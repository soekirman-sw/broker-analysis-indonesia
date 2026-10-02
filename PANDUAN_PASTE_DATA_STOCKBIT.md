# 📊 PANDUAN PASTE DATA STOCKBIT KE EXCEL - LANGKAH DEMI LANGKAH

## Video Summary

**Prinsipnya:**
1. Copy data Stockbit (Broker Summary Net & Gross)
2. Paste ke Excel sheet "DATA_INPUT"
3. Formula otomatis hitung hasil di sheet "HASIL_ANALISIS"

---

## 🔥 LANGKAH 1: SIAPKAN DATA DARI STOCKBIT

### A. Buka Stockbit - Broker Summary NET

```
1. Buka chart saham di Stockbit (contoh: ENRG)
2. Klik tab "Broker Summary" (sebelah kanan)
3. Lihat kolom "Net"
4. Tampilkan TOP 6 BROKER berdasarkan Net Value
```

**Data yang Anda lihat:**
```
Broker | B.val | B.lot | B.avg | SL.val | S.lot | S.avg | Net
-------|-------|-------|-------|--------|-------|-------|-------
MG     | 46.3B | 333.6K| 1.388 | 21.2B  | 145K  | 1.400 | 25.1B
TP     | 41.3B | 294.4K| 1.401 | 18.5B  | 140K  | 1.402 | 22.8B
AK     | 39.9B | 283.4K| 1.406 | 15.2B  | 130K  | 1.403 | 24.7B
XL     | 25.9B | 184.7K| 1.402 | 12.8B  | 95K   | 1.405 | 13.1B
CC     | 16.9B | 121.3K| 1.395 | 8.5B   | 65K   | 1.401 | 8.4B
PD     | 9.68B | 68.6K | 1.406 | 4.2B   | 30K   | 1.408 | 5.48B
```

**COPY kolom:**
- B.val (Buy Value)
- B.lot (Buy Lot)  
- B.avg (Buy Average)
- SL.val (Sell Value)
- S.lot (Sell Lot)
- S.avg (Sell Average)
- Net (Net Value = B.val - SL.val)

### B. Buka Stockbit - Broker Summary GROSS

```
1. Di tab yang sama, ubah ke view "Gross"
2. Tampilkan TOP 6 BROKER berdasarkan Gross Value
```

**Data yang Anda lihat:**
```
Broker | B.val | B.lot | B.avg | SL.val | S.lot | S.avg | Gross
-------|-------|-------|-------|--------|-------|-------|--------
MG     | 46.3B | 333.6K| 1.388 | 21.2B  | 145K  | 1.400 | 67.5B
TP     | 41.3B | 294.4K| 1.401 | 18.5B  | 140K  | 1.402 | 59.8B
AK     | 39.9B | 283.4K| 1.406 | 15.2B  | 130K  | 1.403 | 55.1B
XL     | 25.9B | 184.7K| 1.402 | 12.8B  | 95K   | 1.405 | 38.7B
CC     | 16.9B | 121.3K| 1.395 | 8.5B   | 65K   | 1.401 | 25.4B
PD     | 9.68B | 68.6K | 1.406 | 4.2B   | 30K   | 1.408 | 13.88B
```

---

## 🎯 LANGKAH 2: SIAPKAN EXCEL

### Buka file template:
```
Broker_Analysis_Template.xlsx
```

**File ini memiliki 2 sheet:**
1. **DATA_INPUT** - untuk paste data dari Stockbit
2. **HASIL_ANALISIS** - untuk melihat hasil analisis otomatis

---

## 📥 LANGKAH 3: PASTE DATA KE SHEET "DATA_INPUT"

### Sheet "DATA_INPUT" sudah punya header:

```
Col A    | Col B  | Col C      | Col D      | Col E      | Col F      | Col G      | Col H      | Col I
---------|--------|------------|------------|------------|------------|------------|------------|----------
Date     | Broker | Buy_Val    | Buy_Lot    | Buy_Avg    | Sell_Val   | Sell_Lot   | Sell_Avg   | Net_Val
10-26    | MG     | [paste]    | [paste]    | [paste]    | [paste]    | [paste]    | [paste]    | [formula]
10-26    | TP     | [paste]    | [paste]    | [paste]    | [paste]    | [paste]    | [paste]    | [formula]
```

### Cara paste:

**Untuk tanggal (Kolom A):**
- Ketik tanggal manual atau copy dari Stockbit
- Format: YYYY-MM-DD atau DD-MM-YYYY
- Contoh: 2026-10-26

**Untuk nama broker (Kolom B):**
- Copy dari Stockbit
- Contoh: MG, TP, AK, XL, CC, PD

**Untuk Buy_Val sampai Sell_Avg (Kolom C - H):**
- Copy langsung dari Stockbit
- Bisa dalam format "46.3B" atau angka (46300000000)
- Excel akan otomatis parse format nilai besar (B, M, K)

**Untuk Net_Val (Kolom I):**
- JANGAN di-paste, biarkan kosong
- Formula akan otomatis hitung: Buy_Val - Sell_Val
- Nanti akan muncul sendiri saat refresh

---

## 🔄 LANGKAH 4: CARA PASTE YANG BENAR

### Metode A: Copy tabel langsung dari Stockbit (PALING MUDAH)

```
1. Di Stockbit, seleksi tabel broker summary NET
   - Klik cell pertama broker (MG)
   - Drag ke cell Net terakhir (PD)
   - Ctrl+C (Copy)

2. Buka Excel, klik sheet "DATA_INPUT"
   - Klik cell C2 (Buy_Val baris pertama)
   - Ctrl+V (Paste)

3. Excel akan paste otomatis sesuai kolom
   - Buy_Val → Kolom C
   - Buy_Lot → Kolom D
   - Buy_Avg → Kolom E
   - Sell_Val → Kolom F
   - Sell_Lot → Kolom G
   - Sell_Avg → Kolom H
```

### Metode B: Buat tabel baru (jika butuh banyak data)

```
1. Buka Excel
2. Sheet "DATA_INPUT", mulai dari baris 2
3. Isi:
   - A2: Tanggal hari 1
   - B2: Broker pertama
   - C2-H2: Data Buy/Sell dari Stockbit
   - Tekan Enter

4. Ulangi untuk broker lain dan hari lain
```

---

## 📊 LANGKAH 5: LIHAT HASIL DI "HASIL_ANALISIS"

### Setelah paste data, buka sheet "HASIL_ANALISIS"

**Anda akan melihat tabel dengan kolom:**

| Kolom | Apa Artinya | Nilai Normal |
|-------|-----------|----------|
| **Broker** | Nama broker | MG, TP, AK, dll |
| **Buy_Val (Rp)** | Total nilai beli | 46.3B |
| **Sell_Val (Rp)** | Total nilai jual | 21.2B |
| **Net_Val (Rp)** | Neto (Buy - Sell) | 25.1B |
| **Net_Buy_Ratio** | Proporsi buy vs sell | 0.54 (54% buy) |
| **Gross_Flow (Rp)** | Total aktivitas | 67.5B |
| **Relative_Part (%)** | Kontribusi terhadap market | 35% |
| **Flow_Momentum** | Kecepatan arus | 1.2 (mempercepat) |
| **Broker_Score** | Skor gabungan 0-100 | 78 |
| **Signal** | Sinyal bullish/bearish | BULLISH 🟢 |

---

## 🎨 WARNA SINYAL

```
🟢 BULLISH (Hijau)
  → Broker dominan BELI
  → Momentum mempercepat
  → Kontribusi tinggi
  → Indikasi positif

🟡 NEUTRAL (Kuning)
  → Broker berimbang
  → Momentum stabil
  → Indikasi netral

🔴 BEARISH (Merah)
  → Broker dominan JUAL
  → Momentum melemah
  → Kontribusi negatif
  → Indikasi negatif
```

---

## 💡 CONTOH PRAKTIS

### Scenario: Analisis ENRG tanggal 26 Oktober 2026

**Step 1: Ambil data dari Stockbit**
```
Stockbit menunjukkan:

Broker | Net View          | Gross View
-------|-------------------|-------------------
MG     | 46.3B, 333.6K...  | 46.3B, 333.6K...
TP     | 41.3B, 294.4K...  | 41.3B, 294.4K...
AK     | 39.9B, 283.4K...  | 39.9B, 283.4K...
XL     | 25.9B, 184.7K...  | 25.9B, 184.7K...
CC     | 16.9B, 121.3K...  | 16.9B, 121.3K...
PD     | 9.68B, 68.6K...   | 9.68B, 68.6K...
```

**Step 2: Paste ke Excel (DATA_INPUT)**
```
Row 2: Date=10-26 | Broker=MG | Buy_Val=46.3B | Buy_Lot=333.6K | Buy_Avg=1.388 | Sell_Val=21.2B | Sell_Lot=145K | Sell_Avg=1.400 | Net_Val=(auto)
Row 3: Date=10-26 | Broker=TP | Buy_Val=41.3B | Buy_Lot=294.4K | Buy_Avg=1.401 | Sell_Val=18.5B | Sell_Lot=140K | Sell_Avg=1.402 | Net_Val=(auto)
... dan seterusnya untuk AK, XL, CC, PD
```

**Step 3: Lihat hasil (HASIL_ANALISIS)**
```
Broker | Buy_Val | Sell_Val | Net_Val | Net_Buy_Ratio | Relative_Part | Flow_Momentum | Signal
-------|---------|----------|---------|---------------|---------------|---------------|--------
MG     | 46.3B   | 21.2B    | 25.1B   | 0.54          | 35%           | 1.2           | BULLISH 🟢
TP     | 41.3B   | 18.5B    | 22.8B   | 0.38          | 32%           | 1.1           | BULLISH 🟢
AK     | 39.9B   | 15.2B    | 24.7B   | 0.45          | 34%           | 1.15          | BULLISH 🟢
XL     | 25.9B   | 12.8B    | 13.1B   | 0.34          | 18%           | 0.95          | NEUTRAL 🟡
CC     | 16.9B   | 8.5B     | 8.4B    | 0.33          | 12%           | 0.88          | BEARISH 🔴
PD     | 9.68B   | 4.2B     | 5.48B   | 0.40          | 8%            | 0.85          | BEARISH 🔴
```

**Interpretasi:**
- MG, TP, AK sedang BULLISH (dominan beli, momentum kuat)
- XL berimbang (NEUTRAL)
- CC, PD mulai BEARISH (mulai jual, momentum lemah)

---

## ⚙️ FORMULA DI BALIK LAYAR (Anda gak perlu tahu, tapi untuk referensi)

### Di sheet HASIL_ANALISIS, formula otomatis menghitung:

**1. Net_Buy_Ratio**
```excel
= (Buy_Val - Sell_Val) / (Buy_Val + Sell_Val)
```
Contoh: (46.3B - 21.2B) / (46.3B + 21.2B) = 0.54

**2. Gross_Flow**
```excel
= Buy_Val + Sell_Val
```
Contoh: 46.3B + 21.2B = 67.5B

**3. Relative_Participation**
```excel
= Net_Val / SUM(semua Net_Val) * 100%
```
Contoh: 25.1B / 99.48B * 100% = 25.2%

**4. Flow_Momentum (simplified)**
```excel
= (5-day average Net_Buy_Ratio) / (20-day average Net_Buy_Ratio)
Atau jika data 1 hari: = Net_Buy_Ratio saat ini
```

---

## ❓ FAQ

### Q: Berapa banyak broker yang harus saya input?
**A:** Minimal 6 broker (Top 6 Net + Top 6 Gross dari Stockbit). Bisa lebih, tapi fokus ke whale broker saja.

### Q: Berapa hari data yang diperlukan?
**A:** Sesuai segment Anda (dari reversal terakhir ke candle sekarang). Bisa 5 hari, 10 hari, atau 20 hari. Excel akan otomatis calculate momentum.

### Q: Kolom apa yang harus di-paste?
**A:** Kolom C-H (Buy_Val, Buy_Lot, Buy_Avg, Sell_Val, Sell_Lot, Sell_Avg). Kolom I (Net_Val) otomatis.

### Q: Bagaimana kalau format dari Stockbit berbeda?
**A:** Sesuaikan manual ke template Excel:
- B.val → Buy_Val
- B.lot → Buy_Lot
- B.avg → Buy_Avg
- SL.val atau S.val → Sell_Val
- S.lot → Sell_Lot
- S.avg → Sell_Avg

### Q: Refresh hasilnya bagaimana?
**A:** 
- Otomatis saat Anda ketik/paste data
- Atau tekan Ctrl+Shift+F9 untuk force recalculate

---

## 🎯 SUMMARY

**Alur kerja Anda:**
```
1. Buka Stockbit → Lihat Broker Summary (Net & Gross)
   ↓
2. Copy data Top 6 broker (Buy_Val, Buy_Lot, Buy_Avg, Sell_Val, Sell_Lot, Sell_Avg)
   ↓
3. Paste ke Excel sheet "DATA_INPUT" (mulai kolom C baris 2)
   ↓
4. Isi kolom A (Date) dan B (Broker) manual
   ↓
5. Lihat hasil di sheet "HASIL_ANALISIS"
   ↓
6. Baca sinyal BULLISH/NEUTRAL/BEARISH
   ↓
7. Gunakan untuk konfirmasi analisis Anda
```

**Selesai! 🎉**

Anda sudah bisa menganalisis broker summary tanpa coding sama sekali!

---

**Next:** Download file Excel template dan praktek sekarang!
