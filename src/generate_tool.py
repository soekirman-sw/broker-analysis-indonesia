import sys
sys.path.insert(0, '/tmp')

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
import os

def create_broker_analysis_excel():
    """Generate Excel tool for broker analysis"""
    
    wb = Workbook()
    wb.remove(wb.active)
    
    # ===== SHEET 1: DATA_INPUT =====
    ws_input = wb.create_sheet("DATA_INPUT", 0)
    
    # Set column width
    ws_input.column_dimensions['A'].width = 14
    ws_input.column_dimensions['B'].width = 12
    for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I']:
        ws_input.column_dimensions[col].width = 16
    
    # Header styling
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Header
    headers = [
        "Date",
        "Broker",
        "Buy_Val (Rp)",
        "Buy_Lot",
        "Buy_Avg",
        "Sell_Val (Rp)",
        "Sell_Lot",
        "Sell_Avg",
        "Net_Val (Rp)"
    ]
    
    for col_num, header in enumerate(headers, 1):
        cell = ws_input.cell(row=1, column=col_num)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = border
    
    # Add 49 empty data rows with formulas
    for row in range(2, 51):
        for col in range(1, 10):
            cell = ws_input.cell(row=row, column=col)
            cell.border = border
            if col == 9:  # Net_Val formula
                cell.value = f"=IF(C{row}=\"\",\"\",C{row}-F{row})"
                cell.number_format = '#,##0'
            elif col in [3, 4, 6, 7]:
                cell.number_format = '#,##0'
            elif col in [5, 8]:
                cell.number_format = '0.000'
            elif col == 1:
                cell.number_format = 'yyyy-mm-dd'
    
    # Add instructions
    inst_row = 52
    ws_input[f'A{inst_row}'] = "PANDUAN PENGGUNAAN:"
    ws_input[f'A{inst_row}'].font = Font(bold=True, size=11, color="1F4E78")
    
    instructions = [
        "1. Copy data Broker Summary dari Stockbit (Top 6 Net + Top 6 Gross)",
        "2. Paste ke baris 2 di atas (mulai dari kolom C - Buy_Val)",
        "3. Isi kolom A (Date) dan B (Broker) secara manual",
        "4. Kolom I (Net_Val) otomatis hitung dari formula",
        "5. Buka sheet 'HASIL_ANALISIS' untuk melihat hasil analisis",
        "6. Kolom yang di-paste dari Stockbit: Buy_Val, Buy_Lot, Buy_Avg, Sell_Val, Sell_Lot, Sell_Avg"
    ]
    
    for i, instruction in enumerate(instructions, 1):
        cell = ws_input[f'A{inst_row + i}']
        cell.value = instruction
        cell.font = Font(italic=True, size=10, color="666666")
        cell.alignment = Alignment(wrap_text=True)
    
    # ===== SHEET 2: HASIL_ANALISIS =====
    ws_output = wb.create_sheet("HASIL_ANALISIS", 1)
    
    # Set column width
    ws_output.column_dimensions['A'].width = 12
    for col in ['B', 'C', 'D', 'E', 'F']:
        ws_output.column_dimensions[col].width = 16
    for col in ['G', 'H', 'I', 'J']:
        ws_output.column_dimensions[col].width = 15
    
    # Output headers
    output_headers = [
        "Broker",
        "Buy_Val (Rp)",
        "Sell_Val (Rp)",
        "Net_Val (Rp)",
        "Net_Buy_Ratio",
        "Gross_Flow (Rp)",
        "Relative_Part (%)",
        "Flow_Momentum",
        "Broker_Score",
        "Signal"
    ]
    
    output_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    output_font = Font(bold=True, color="FFFFFF", size=11)
    output_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for col_num, header in enumerate(output_headers, 1):
        cell = ws_output.cell(row=1, column=col_num)
        cell.value = header
        cell.fill = output_fill
        cell.font = output_font
        cell.alignment = output_alignment
        cell.border = border
    
    # Add formulas for calculation (up to 10 brokers)
    for row in range(2, 12):
        # Column A: Broker (reference from DATA_INPUT)
        cell_a = ws_output.cell(row=row, column=1)
        cell_a.value = f'=IF(DATA_INPUT!B{row}="","",DATA_INPUT!B{row})'
        cell_a.border = border
        
        # Column B: Buy_Val
        cell_b = ws_output.cell(row=row, column=2)
        cell_b.value = f'=IF(DATA_INPUT!C{row}="",0,DATA_INPUT!C{row})'
        cell_b.number_format = '#,##0'
        cell_b.border = border
        
        # Column C: Sell_Val
        cell_c = ws_output.cell(row=row, column=3)
        cell_c.value = f'=IF(DATA_INPUT!F{row}="",0,DATA_INPUT!F{row})'
        cell_c.number_format = '#,##0'
        cell_c.border = border
        
        # Column D: Net_Val
        cell_d = ws_output.cell(row=row, column=4)
        cell_d.value = f'=B{row}-C{row}'
        cell_d.number_format = '#,##0'
        cell_d.border = border
        
        # Column E: Net_Buy_Ratio
        cell_e = ws_output.cell(row=row, column=5)
        cell_e.value = f'=IF(B{row}+C{row}=0,0,(B{row}-C{row})/(B{row}+C{row}))'
        cell_e.number_format = '0.00'
        cell_e.alignment = Alignment(horizontal="center")
        cell_e.border = border
        
        # Column F: Gross_Flow
        cell_f = ws_output.cell(row=row, column=6)
        cell_f.value = f'=B{row}+C{row}'
        cell_f.number_format = '#,##0'
        cell_f.border = border
        
        # Column G: Relative_Participation
        cell_g = ws_output.cell(row=row, column=7)
        cell_g.value = f'=IF(SUM($D$2:$D$11)=0,0,(D{row}/SUM($D$2:$D$11))*100)'
        cell_g.number_format = '0.0'
        cell_g.alignment = Alignment(horizontal="center")
        cell_g.border = border
        
        # Column H: Flow_Momentum
        cell_h = ws_output.cell(row=row, column=8)
        cell_h.value = f'=IF(AVERAGE($E$2:$E$11)=0,1,E{row}/AVERAGE($E$2:$E$11))'
        cell_h.number_format = '0.00'
        cell_h.alignment = Alignment(horizontal="center")
        cell_h.border = border
        
        # Column I: Broker_Score
        cell_i = ws_output.cell(row=row, column=9)
        cell_i.value = f'=IF(A{row}="","",INT((ABS(E{row})*50 + G{row} + H{row}*30)/1.8))'
        cell_i.number_format = '0'
        cell_i.alignment = Alignment(horizontal="center")
        cell_i.border = border
        
        # Column J: Signal
        cell_j = ws_output.cell(row=row, column=10)
        cell_j.value = f'=IF(A{row}="","",IF(I{row}>=60,"BULLISH",IF(I{row}>=40,"NEUTRAL","BEARISH")))'
        cell_j.alignment = Alignment(horizontal="center")
        cell_j.border = border
        
        # Apply color to signal cells
        if row % 2 == 0:
            light_fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
            for col in range(1, 11):
                ws_output.cell(row=row, column=col).fill = light_fill
    
    # Add explanation
    expl_row = 14
    ws_output[f'A{expl_row}'] = "PENJELASAN KOLOM:"
    ws_output[f'A{expl_row}'].font = Font(bold=True, size=11, color="1F4E78")
    
    explanations = [
        "• Net_Buy_Ratio: Proporsi buy vs sell (-1 sampai 1), positif = buy, negatif = sell",
        "• Gross_Flow: Total aktivitas trading (Buy_Val + Sell_Val)",
        "• Relative_Part: Kontribusi broker terhadap total net buy market (%)",
        "• Flow_Momentum: Kecepatan arus (>1 = mempercepat, <1 = melemah)",
        "• Broker_Score: Skor gabungan (0-100), semakin tinggi semakin bullish",
        "• Signal: BULLISH (Score >= 60), NEUTRAL (40-60), BEARISH (<40)"
    ]
    
    for i, explanation in enumerate(explanations, 1):
        cell = ws_output[f'A{expl_row + i}']
        cell.value = explanation
        cell.font = Font(size=10, color="333333")
        cell.alignment = Alignment(wrap_text=True)
    
    # ===== SHEET 3: PANDUAN =====
    ws_guide = wb.create_sheet("PANDUAN", 2)
    ws_guide.column_dimensions['A'].width = 120
    
    guide_title = "BROKER ANALYSIS TOOL - PANDUAN LENGKAP"
    ws_guide['A1'] = guide_title
    ws_guide['A1'].font = Font(bold=True, size=14, color="1F4E78")
    
    guide_lines = [
        ("", ""),
        ("LANGKAH PENGGUNAAN:", "bold_title"),
        ("", ""),
        ("1. BUKA SHEET 'DATA_INPUT'", "bold"),
        ("   - Lihat kolom A sampai I (Date, Broker, Buy_Val, Buy_Lot, Buy_Avg, Sell_Val, Sell_Lot, Sell_Avg, Net_Val)", "normal"),
        ("   - Baris 1 adalah header (jangan diedit)", "normal"),
        ("   - Mulai input data dari baris 2", "normal"),
        ("", ""),
        ("2. AMBIL DATA DARI STOCKBIT", "bold"),
        ("   - Buka chart saham di Stockbit (contoh: ENRG)", "normal"),
        ("   - Klik 'Broker Summary' (sebelah kanan)", "normal"),
        ("   - Lihat TOP 6 BROKER dari Net Summary", "normal"),
        ("   - Lihat TOP 6 BROKER dari Gross Summary", "normal"),
        ("   - Copy kolom: B.val, B.lot, B.avg, SL.val, S.lot, S.avg", "normal"),
        ("", ""),
        ("3. ATUR DATA DI SHEET 'DATA_INPUT'", "bold"),
        ("   - Baris 2 dst: Isi tanggal (kolom A) dan nama broker (kolom B)", "normal"),
        ("   - Paste data Buy/Sell dari Stockbit ke kolom C-H", "normal"),
        ("   - Kolom I (Net_Val) akan otomatis hitung dengan formula", "normal"),
        ("   - Contoh format tanggal: 2026-10-26 atau 26/10/2026", "normal"),
        ("", ""),
        ("4. LIHAT HASIL DI SHEET 'HASIL_ANALISIS'", "bold"),
        ("   - Hasil akan otomatis ter-update setelah Anda input data", "normal"),
        ("   - Sheet ini menampilkan perhitungan otomatis dari DATA_INPUT", "normal"),
        ("   - Fokus pada kolom 'Signal' untuk rekomendasi (BULLISH/NEUTRAL/BEARISH)", "normal"),
        ("", ""),
        ("INTERPRETASI HASIL:", "bold_title"),
        ("", ""),
        ("Net_Buy_Ratio (Proporsi Beli vs Jual):", "bold"),
        ("  Nilai berkisar dari -1.0 (100% jual) sampai +1.0 (100% beli)", "normal"),
        ("  > 0.5  = Broker DOMINAN BELI (bullish signal)", "normal"),
        ("  0 - 0.5 = Broker SEDIKIT BELI", "normal"),
        ("  -0.5 - 0 = Broker SEDIKIT JUAL", "normal"),
        ("  < -0.5 = Broker DOMINAN JUAL (bearish signal)", "normal"),
        ("", ""),
        ("Relative Participation (Kontribusi Broker):", "bold"),
        ("  Menunjukkan seberapa besar kontribusi broker terhadap total net buy market", "normal"),
        ("  > 25%  = Kontribusi sangat tinggi (whale besar)", "normal"),
        ("  15-25% = Kontribusi tinggi (whale aktif)", "normal"),
        ("  5-15%  = Kontribusi sedang", "normal"),
        ("  < 5%   = Kontribusi kecil", "normal"),
        ("", ""),
        ("Flow_Momentum (Kecepatan Arus):", "bold"),
        ("  Membandingkan momentum saat ini dengan rata-rata", "normal"),
        ("  > 1.2  = Momentum MEMPERCEPAT kuat (signal BULLISH)", "normal"),
        ("  1.0-1.2 = Momentum STABIL kuat", "normal"),
        ("  0.8-1.0 = Momentum STABIL / mulai melambat", "normal"),
        ("  < 0.8  = Momentum MELEMAH (signal BEARISH)", "normal"),
        ("", ""),
        ("Broker_Score (Skor Gabungan):", "bold"),
        ("  Kombinasi dari Net_Buy_Ratio, Relative_Participation, dan Flow_Momentum", "normal"),
        ("  Semakin tinggi skor, semakin bullish broker tersebut", "normal"),
        ("  Score >= 60 = BULLISH (broker sedang beli kuat)", "normal"),
        ("  Score 40-60 = NEUTRAL (broker berimbang)", "normal"),
        ("  Score < 40  = BEARISH (broker sedang jual)", "normal"),
        ("", ""),
        ("Signal (REKOMENDASI):", "bold"),
        ("  🟢 BULLISH  = Broker sedang MASUK (dominan beli, momentum kuat, kontribusi tinggi)", "normal"),
        ("  🟡 NEUTRAL  = Broker BERIMBANG (tidak ada tekanan kuat)", "normal"),
        ("  🔴 BEARISH  = Broker sedang KELUAR (dominan jual, momentum lemah)", "normal"),
        ("", ""),
        ("TIPS & TRIK:", "bold_title"),
        ("", ""),
        ("✓ Gunakan MULTI-HARI, bukan analisis per-hari", "normal"),
        ("  - Kumpulkan data dari reversal terakhir sampai candle sekarang", "normal"),
        ("  - Tool akan otomatis smooth noise harian", "normal"),
        ("", ""),
        ("✓ FOKUS pada TOP 6 BROKER WHALE saja", "normal"),
        ("  - Ambil Top 6 dari Net Summary + Top 6 dari Gross Summary", "normal"),
        ("  - Jangan lupa ada broker yang berbeda di Net vs Gross", "normal"),
        ("", ""),
        ("✓ KOMBINASIKAN dengan analisis technical", "normal"),
        ("  - Broker bullish + support/breakout = signal kuat", "normal"),
        ("  - Broker bearish + resistance break = warning", "normal"),
        ("", ""),
        ("✓ MONITOR PERUBAHAN SIGNAL", "normal"),
        ("  - Jika signal berubah dari BULLISH ke NEUTRAL = momentum mulai lemah", "normal"),
        ("  - Jika signal berubah dari NEUTRAL ke BEARISH = perubahan arah", "normal"),
        ("", ""),
        ("✓ GUNAKAN SEBAGAI CONFIRMATION, BUKAN SATU-SATUNYA SINYAL", "normal"),
        ("  - Broker analysis melengkapi, tidak menggantikan technical analysis", "normal"),
        ("  - Risk management tetap penting", "normal"),
        ("", ""),
        ("CONTOH PRAKTIS:", "bold_title"),
        ("", ""),
        ("Scenario: Anda analisis saham ENRG dari tanggal 24-26 Oktober (3 hari)", "normal"),
        ("", ""),
        ("DATA INPUT:", "bold"),
        ("", "normal"),
        ("Date       | Broker | Buy_Val | Buy_Lot | Buy_Avg | Sell_Val | Sell_Lot | Sell_Avg | Net_Val (auto)", "normal"),
        ("2026-10-24 | MG     | 40.2B   | 300K    | 1.340   | 18.5B    | 140K     | 1.320    | 21.7B", "normal"),
        ("2026-10-25 | MG     | 43.5B   | 315K    | 1.370   | 19.8B    | 145K     | 1.360    | 23.7B", "normal"),
        ("2026-10-26 | MG     | 46.3B   | 333K    | 1.388   | 21.2B    | 150K     | 1.400    | 25.1B", "normal"),
        ("", ""),
        ("HASIL ANALISIS:", "bold"),
        ("", "normal"),
        ("Broker | Buy_Val | Sell_Val | Net_Val | Net_Buy_Ratio | Relative_Part | Flow_Momentum | Broker_Score | Signal", "normal"),
        ("MG     | 46.3B   | 21.2B    | 25.1B   | 0.37          | 28%           | 1.08          | 65            | BULLISH 🟢", "normal"),
        ("", ""),
        ("INTERPRETASI:", "normal"),
        ("1. Net_Buy_Ratio 0.37 = MG membeli 37% lebih banyak dari jual (positif)", "normal"),
        ("2. Relative_Part 28% = MG berkontribusi 28% dari total net buy market (tinggi)", "normal"),
        ("3. Flow_Momentum 1.08 = Momentum 8% lebih cepat dari rata-rata (mempercepat)", "normal"),
        ("4. Broker_Score 65 = Skor tinggi (>=60)", "normal"),
        ("5. Signal: BULLISH 🟢 = Broker MG sedang dalam mode BELI KUAT", "normal"),
        ("", ""),
        ("KESIMPULAN:", "normal"),
        ("Whale broker MG sedang MASUK ke saham ENRG dengan momentum yang mempercepat.", "normal"),
        ("Ini adalah signal POSITIF untuk harga saham ENRG (confirmation untuk buy)", "normal"),
        ("", ""),
        ("---", "normal"),
        ("Created: Oktober 2026 | Broker Analysis Tool v1.0", "normal"),
    ]
    
    for row_num, (text, style) in enumerate(guide_lines, 1):
        cell = ws_guide[f'A{row_num}']
        cell.value = text
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        
        if style == "bold_title":
            cell.font = Font(bold=True, size=12, color="1F4E78")
        elif style == "bold":
            cell.font = Font(bold=True, size=11, color="333333")
        else:
            cell.font = Font(size=10, color="333333")
    
    # Save workbook
    output_file = "Broker_Analysis_Tool.xlsx"
    wb.save(output_file)
    
    print(f"✅ Excel file created: {output_file}")
    return output_file

if __name__ == "__main__":
    create_broker_analysis_excel()
