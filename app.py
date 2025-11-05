from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
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

# Rejestracja czcionki obsługującej polskie znaki
try:
    pdfmetrics.registerFont(TTFont('DejaVuSans', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
    pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
    DEFAULT_FONT = 'DejaVuSans'
    DEFAULT_FONT_BOLD = 'DejaVuSans-Bold'
except:
    # Fallback - próba innych lokalizacji czcionek
    try:
        pdfmetrics.registerFont(TTFont('DejaVuSans', 'C:\\Windows\\Fonts\\DejaVuSans.ttf'))
        pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', 'C:\\Windows\\Fonts\\DejaVuSans-Bold.ttf'))
        DEFAULT_FONT = 'DejaVuSans'
        DEFAULT_FONT_BOLD = 'DejaVuSans-Bold'
    except:
        # Jeśli nie ma DejaVu, używamy Helvetica (bez pełnej obsługi polskich znaków)
        DEFAULT_FONT = 'Helvetica'
        DEFAULT_FONT_BOLD = 'Helvetica-Bold'

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
        
        # Formatowanie Excela z automatycznym dopasowaniem szerokości kolumn i sumami
        format_excel(output_path, df_grouped, df_handlowcy_sum)
        
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

def format_excel(filepath, df_grouped, df_handlowcy_sum):
    """Formatowanie pliku Excel z automatycznym dopasowaniem szerokości kolumn i sumami"""
    wb = openpyxl.load_workbook(filepath)
    
    # Definiowanie stylów
    header_font = Font(bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    sum_font = Font(bold=True, size=11, color="1F4788")
    sum_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Formatowanie nagłówków
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
        
        # Formatowanie komórek danych
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.border = border
                
                # Formatowanie kwot
                if cell.column in [2, 3]:  # Kolumny z kwotami
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00 "PLN"'
        
        # Automatyczne dopasowanie szerokości kolumn
        for column_cells in ws.columns:
            length = 0
            column = column_cells[0].column_letter
            
            for cell in column_cells:
                try:
                    if cell.value:
                        cell_value = str(cell.value)
                        if len(cell_value) > length:
                            length = len(cell_value)
                except:
                    pass
            
            # Ustawienie szerokości (z małym marginesem)
            adjusted_width = min(length + 4, 80)  # Maksymalna szerokość 80
            ws.column_dimensions[column].width = adjusted_width
        
        # Zamrożenie pierwszego wiersza
        ws.freeze_panes = 'A2'
    
    # Dodanie sum w zakładce PODSUMOWANIE GRUP
    ws_groups = wb['PODSUMOWANIE GRUP']
    last_row = ws_groups.max_row + 1
    
    # Obliczenie sum
    suma_kwot = df_grouped['Suma kwot'].sum()
    suma_prowizji_groups = df_grouped['10% z sumy'].sum()
    
    # Dodanie wiersza z sumą
    ws_groups[f'A{last_row}'] = 'SUMA CAŁKOWITA:'
    ws_groups[f'B{last_row}'] = suma_kwot
    ws_groups[f'C{last_row}'] = suma_prowizji_groups
    
    # Formatowanie wiersza sumy
    for col in ['A', 'B', 'C', 'D']:
        cell = ws_groups[f'{col}{last_row}']
        cell.font = sum_font
        cell.fill = sum_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='left' if col == 'A' else 'left', vertical='center')
        
        if col in ['B', 'C']:
            cell.number_format = '#,##0.00 "PLN"'
    
    # Dodanie sum w zakładce PODSUMOWANIE HANDLOWCÓW
    ws_handlowcy = wb['PODSUMOWANIE HANDLOWCÓW']
    last_row_h = ws_handlowcy.max_row + 1
    
    # Obliczenie sumy prowizji handlowców
    suma_prowizji_handlowcy = df_handlowcy_sum['Suma prowizji (10%)'].sum()
    
    # Dodanie wiersza z sumą
    ws_handlowcy[f'A{last_row_h}'] = 'SUMA CAŁKOWITA:'
    ws_handlowcy[f'B{last_row_h}'] = suma_prowizji_handlowcy
    
    # Formatowanie wiersza sumy
    for col in ['A', 'B']:
        cell = ws_handlowcy[f'{col}{last_row_h}']
        cell.font = sum_font
        cell.fill = sum_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='left' if col == 'A' else 'left', vertical='center')
        
        if col == 'B':
            cell.number_format = '#,##0.00 "PLN"'
    
    wb.save(filepath)

def create_pdf(output_path, df_grouped, df_handlowcy_sum, date_from, date_to, time_from, time_to):
    """Tworzenie raportu PDF z obsługą polskich znaków i logo"""
    # Użycie orientacji poziomej dla lepszego wykorzystania przestrzeni
    doc = SimpleDocTemplate(
        output_path, 
        pagesize=landscape(A4),
        topMargin=15*mm,
        bottomMargin=20*mm,
        leftMargin=15*mm,
        rightMargin=15*mm
    )
    
    elements = []
    styles = getSampleStyleSheet()
    
    # Dodanie logo jeśli istnieje
    logo_path = os.path.join('static', 'logo.png')
    if os.path.exists(logo_path):
        try:
            logo = RLImage(logo_path, width=40*mm, height=40*mm)
            logo.hAlign = 'CENTER'
            elements.append(logo)
            elements.append(Spacer(1, 5))
        except:
            pass  # Jeśli logo nie może być załadowane, kontynuuj bez niego
    
    # Tytuł
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=DEFAULT_FONT_BOLD,
        fontSize=18,
        alignment=TA_CENTER,
        spaceAfter=12,
        textColor=colors.HexColor('#1f4788')
    )
    elements.append(Paragraph('RAPORT WPŁAT', title_style))
    elements.append(Spacer(1, 10))
    
    # Zakres dat
    date_style = ParagraphStyle(
        'DateStyle',
        parent=styles['Normal'],
        fontName=DEFAULT_FONT,
        fontSize=11,
        alignment=TA_CENTER,
        spaceAfter=20
    )
    date_range = f"Okres: {date_from} {time_from} - {date_to} {time_to}"
    elements.append(Paragraph(date_range, date_style))
    elements.append(Spacer(1, 15))
    
    # Nagłówek sekcji
    section_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading2'],
        fontName=DEFAULT_FONT_BOLD,
        fontSize=14,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=10
    )
    
    # Style dla paragrafów w tabelach
    cell_style = ParagraphStyle(
        'CellStyle',
        parent=styles['Normal'],
        fontName=DEFAULT_FONT,
        fontSize=9,
        leading=11
    )
    
    cell_style_bold = ParagraphStyle(
        'CellStyleBold',
        parent=styles['Normal'],
        fontName=DEFAULT_FONT_BOLD,
        fontSize=10,
        textColor=colors.HexColor('#1f4788'),
        leading=12
    )
    
    # TABELA PODSUMOWANIA GRUP
    elements.append(Paragraph('PODSUMOWANIE GRUP', section_style))
    elements.append(Spacer(1, 8))
    
    # Przygotowanie danych
    data_groups = [['Grupa zajęciowa', 'Suma kwot', '10% z sumy', 'Handlowiec']]
    
    total_kwot = 0
    total_prowizji = 0
    
    for _, row in df_grouped.iterrows():
        # Zawijanie długiego tekstu w Paragraph
        grupa_para = Paragraph(str(row['Grupa zajęciowa']), cell_style)
        handlowiec_para = Paragraph(str(row['Handlowiec']), cell_style)
        
        data_groups.append([
            grupa_para,
            f"{row['Suma kwot']:.2f} PLN",
            f"{row['10% z sumy']:.2f} PLN",
            handlowiec_para
        ])
        
        total_kwot += row['Suma kwot']
        total_prowizji += row['10% z sumy']
    
    # Dodanie wiersza podsumowującego
    data_groups.append([
        Paragraph('<b>SUMA CAŁKOWITA:</b>', cell_style_bold),
        f"{total_kwot:.2f} PLN",
        f"{total_prowizji:.2f} PLN",
        ''
    ])
    
    # Szerokości kolumn (lepiej dopasowane)
    col_widths = [110*mm, 45*mm, 45*mm, 65*mm]
    
    table_groups = Table(data_groups, colWidths=col_widths, repeatRows=1)
    table_groups.setStyle(TableStyle([
        # Nagłówek
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (2, 0), 'RIGHT'),
        ('ALIGN', (3, 0), (3, 0), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), DEFAULT_FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('TOPPADDING', (0, 0), (-1, 0), 10),
        
        # Dane
        ('BACKGROUND', (0, 1), (-1, -2), colors.beige),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 1), (2, -1), 'RIGHT'),
        ('ALIGN', (3, 1), (3, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (-1, -2), DEFAULT_FONT),
        ('FONTSIZE', (0, 1), (-1, -2), 9),
        ('VALIGN', (0, 1), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
        
        # Wiersz sumy
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#d9e1f2')),
        ('FONTNAME', (0, -1), (-1, -1), DEFAULT_FONT_BOLD),
        ('FONTSIZE', (0, -1), (-1, -1), 10),
        ('LINEABOVE', (0, -1), (-1, -1), 2, colors.HexColor('#1f4788')),
        
        # Siatka
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BOX', (0, 0), (-1, -1), 1.5, colors.HexColor('#1f4788'))
    ]))
    
    elements.append(table_groups)
    elements.append(Spacer(1, 25))
    
    # TABELA PODSUMOWANIA HANDLOWCÓW
    elements.append(Paragraph('PODSUMOWANIE HANDLOWCÓW', section_style))
    elements.append(Spacer(1, 8))
    
    data_handlowcy = [['Handlowiec', 'Suma prowizji (10%)']]
    
    total_handlowcy = 0
    
    for _, row in df_handlowcy_sum.iterrows():
        handlowiec_para = Paragraph(str(row['Handlowiec']), cell_style)
        data_handlowcy.append([
            handlowiec_para,
            f"{row['Suma prowizji (10%)']:.2f} PLN"
        ])
        total_handlowcy += row['Suma prowizji (10%)']
    
    # Dodanie wiersza podsumowującego
    data_handlowcy.append([
        Paragraph('<b>SUMA CAŁKOWITA:</b>', cell_style_bold),
        f"{total_handlowcy:.2f} PLN"
    ])
    
    table_handlowcy = Table(data_handlowcy, colWidths=[150*mm, 80*mm], repeatRows=1)
    table_handlowcy.setStyle(TableStyle([
        # Nagłówek
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), DEFAULT_FONT_BOLD),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('TOPPADDING', (0, 0), (-1, 0), 10),
        
        # Dane
        ('BACKGROUND', (0, 1), (-1, -2), colors.beige),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 1), (-1, -2), DEFAULT_FONT),
        ('FONTSIZE', (0, 1), (-1, -2), 10),
        ('VALIGN', (0, 1), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
        
        # Wiersz sumy
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#d9e1f2')),
        ('FONTNAME', (0, -1), (-1, -1), DEFAULT_FONT_BOLD),
        ('FONTSIZE', (0, -1), (-1, -1), 10),
        ('LINEABOVE', (0, -1), (-1, -1), 2, colors.HexColor('#1f4788')),
        
        # Siatka
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BOX', (0, 0), (-1, -1), 1.5, colors.HexColor('#1f4788'))
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
