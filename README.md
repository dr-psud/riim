# Grafik Laying Rate & Bobot Telur (Flok G2)

Visualisasi tren modern dari data Excel **Mas Pito_Grafik LR dan Egg Weight**.

## Isi

| File | Keterangan |
|------|------------|
| `charts/dashboard.html` | Dashboard interaktif (Chart.js) — buka di browser |
| `charts/laying_rate_g2.png` | Grafik Laying Rate G2 (%) minggu 19–57 |
| `charts/egg_weight_g2.png` | Grafik rata-rata bobot telur (g) minggu 20–30 |
| `charts/dashboard_lr_egg_weight_g2.png` | Ringkasan gabungan kedua tren |
| `charts/Mas_Pito_Grafik_LR_dan_Egg_Weight.xlsx` | Sumber data |
| `scripts/generate_charts.py` | Script regenerasi PNG |

## Cara pakai

1. Buka `charts/dashboard.html` di browser untuk grafik interaktif.
2. Atau gunakan file PNG di folder `charts/` untuk presentasi / laporan.
3. Regenerasi PNG:

```bash
python3 scripts/generate_charts.py
```

## Ringkasan temuan

- **Laying Rate** puncak **91,53%** di minggu **34**, turun ke ~43% di W51, rebound ke **56,35%** di W57.
- **Bobot telur** naik dari **66,18 g** (W20) ke **72,44 g** (W30); SD mengecil (lebih seragam).
