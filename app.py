from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# Tworzenie folderów jeśli nie istnieją
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('static', exist_ok=True)

HANDLOWCY_FILE = 'handlowcy.xlsx'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'Brak pliku'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nie wybrano pliku'}), 400

    if file and file.filename.endswith('.csv'):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        return jsonify({'success': True, 'filename': filename})

    return jsonify({'error': 'Nieprawidłowy format pliku'}), 400

@app.route('/process', methods=['POST'])
def process_data():
    data = request.json
    filename = data.get('filename')
    date_from = data.get('date_from')
    date_to = data.get('date_to')
    time_from = data.get('time_from', '00:00')
    time_to = data.get('time_to', '23:59')

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    try:
        # Wczytanie pliku CSV
        df = pd.read_csv(filepath, delimiter=';', encoding='utf-8')

        # Konwersja daty płatności
        df['data płatności'] = pd.to_datetime(df['data płatności'], format='%d.%m.%Y %H:%M')

        # Konwersja kwoty (zamiana przecinka na kropkę)
        df['kwota'] = df['kwota'].astype(str).str.replace(',', '.').astype(float)

        # Tworzenie zakresów dat
        datetime_from = datetime.strptime(f"{date_from} {time_from}", "%Y-%m-%d %H:%M")
        datetime_to = datetime.strptime(f"{date_to} {time_to}", "%Y-%m-%d %H:%M")

        # Filtrowanie danych
        df_filtered = df[(df['data płatności'] >= datetime_from) & 
                        (df['data płatności'] <= datetime_to)].copy()

        # Wczytanie danych handlowców
        df_handlowcy = pd.read_excel(HANDLOWCY_FILE)

        # Tworzenie pliku Excel
        output_excel = f"raport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        output_path = os.path.join('static', output_excel)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Zakładka 1: Cały plik CSV
            df.to_excel(writer, sheet_name='Cały plik', index=False)

            # Zakładka 2: Dane do wyliczeń
            df_filtered.to_excel(writer, sheet_name='DANE DO WYLICZEŃ', index=False)

            # Zakładka 3: Podsumowanie grup
            df_grouped = df_filtered.groupby('Grupa zajęciowa')['kwota'].sum().reset_index()
            df_grouped.columns = ['Grupa zajęciowa', 'Suma kwot']
            df_grouped['10% z sumy'] = df_grouped['Suma kwot'] * 0.1

            # Przypisanie handlowców
            df_grouped = df_grouped.merge(df_handlowcy, on='Grupa zajęciowa', how='left')
            df_grouped['Handlowiec'] = df_grouped['Handlowiec'].fillna('BRAK PRZYPISANIA')

            df_grouped[['Grupa zajęciowa', 'Suma kwot', '10% z sumy', 'Handlowiec']].to_excel(
                writer, sheet_name='PODSUMOWANIE GRUP', index=False
            )

            # Zakładka 4: Podsumowanie handlowców
            df_handlowcy_sum = df_grouped.groupby('Handlowiec')['10% z sumy'].sum().reset_index()
            df_handlowcy_sum.columns = ['Handlowiec', 'Suma prowizji (10%)']
            df_handlowcy_sum.to_excel(writer, sheet_name='PODSUMOWANIE HANDLOWCÓW', index=False)

        # Formatowanie Excela
        wb = openpyxl.load_workbook(output_path)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
        wb.save(output_path)

        # Tworzenie PDF
        output_pdf = f"raport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        pdf_path = os.path.join('static', output_pdf)

        create_pdf(pdf_path, df_grouped, df_handlowcy_sum, date_from, date_to, time_from, time_to)

        return jsonify({
            'success': True,
            'excel_file': output_excel,
            'pdf_file': output_pdf,
            'date_range': f"{date_from} {time_from} - {date_to} {time_to}"
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500

def create_pdf(output_path, df_grouped, df_handlowcy_sum, date_from, date_to, time_from, time_to):
    """Tworzenie raportu PDF"""
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    # Tytuł
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=1
    )

    elements.append(Paragraph('RAPORT WPŁAT', title_style))
    elements.append(Spacer(1, 12))

    # Zakres dat
    date_range = f"Okres: {date_from} {time_from} - {date_to} {time_to}"
    elements.append(Paragraph(date_range, styles['Normal']))
    elements.append(Spacer(1, 20))

    # Tabela podsumowania grup
    elements.append(Paragraph('PODSUMOWANIE GRUP', styles['Heading2']))
    elements.append(Spacer(1, 12))

    data_groups = [['Grupa zajęciowa', 'Suma kwot', '10% z sumy', 'Handlowiec']]
    for _, row in df_grouped.iterrows():
        data_groups.append([
            row['Grupa zajęciowa'],
            f"{row['Suma kwot']:.2f} PLN",
            f"{row['10% z sumy']:.2f} PLN",
            row['Handlowiec']
        ])

    table_groups = Table(data_groups, colWidths=[120*mm/4, 40*mm/4, 40*mm/4, 60*mm/4])
    table_groups.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table_groups)
    elements.append(Spacer(1, 30))

    # Tabela podsumowania handlowców
    elements.append(Paragraph('PODSUMOWANIE HANDLOWCÓW', styles['Heading2']))
    elements.append(Spacer(1, 12))

    data_handlowcy = [['Handlowiec', 'Suma prowizji (10%)']]
    for _, row in df_handlowcy_sum.iterrows():
        data_handlowcy.append([
            row['Handlowiec'],
            f"{row['Suma prowizji (10%)']:.2f} PLN"
        ])

    table_handlowcy = Table(data_handlowcy, colWidths=[100*mm/2, 60*mm/2])
    table_handlowcy.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table_handlowcy)

    doc.build(elements)

@app.route('/download/<filename>')
def download_file(filename):
    filepath = os.path.join('static', filename)
    return send_file(filepath, as_attachment=True)

@app.route('/manage_handlowcy')
def manage_handlowcy():
    return render_template('handlowcy.html')

@app.route('/get_handlowcy')
def get_handlowcy():
    try:
        df = pd.read_excel(HANDLOWCY_FILE)
        return jsonify(df.to_dict('records'))
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/update_handlowcy', methods=['POST'])
def update_handlowcy():
    try:
        data = request.json
        df = pd.DataFrame(data)
        df.to_excel(HANDLOWCY_FILE, index=False)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5100, debug=True)
