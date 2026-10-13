# Analisis Kurva S Pengguna dan Merchant QRIS Indonesia 2021–2026

<!-- DOI: diisi setelah rilis pertama diarsipkan Zenodo -->

Kode dan dataset untuk mengukur pola pertumbuhan pengguna dan *merchant* QRIS (*Quick Response Code Indonesian Standard*) dengan model difusi kurva S: logistik, Gompertz, Bass, dan bi-logistik.

*Code and dataset for modelling the growth of QRIS users and merchants in Indonesia (2021–2026) with logistic, Gompertz, Bass and bi-logistic diffusion models.*

Penyusun: Ade Rumintarsih, Boash University, Bogor.

## Isi repositori

| Lokasi | Isi |
| --- | --- |
| `QRIS_analisis_kurva_S.ipynb` | Notebook utama: seluruh kode analisis dalam satu berkas |
| `data/Dataset_QRIS_Indonesia_2021-2026.xlsx` | Dataset lengkap: pengguna dan *merchant*, transaksi tahunan, pertumbuhan transaksi, QRIS antarnegara |
| `data/qris_*.csv` | Dataset yang sama dalam format CSV |
| `data/bangun_dataset.py` | Daftar seluruh titik data beserta sumbernya; membangun berkas Excel dan CSV |
| `data/LICENSE.md` | Lisensi dataset (CC BY 4.0) |
| `analisis/` | Hasil notebook: estimasi model (`hasil_*.json`), proyeksi 2026–2030, gambar, dan pemeriksaan tambahan |
| `hasil_verifikasi_ulang.csv` | Pencocokan 28 angka kunci hasil notebook dengan angka yang dilaporkan |

## Data

- **Periode:** pengguna Desember 2021–Agustus 2026 (30 titik), *merchant* September 2022–Agustus 2026 (29 titik), transaksi tahunan 2020–semester I 2026.
- **Sumber:** siaran pers resmi Bank Indonesia (bi.go.id), pernyataan pejabat Bank Indonesia yang dimuat media, dan data Asosiasi Sistem Pembayaran Indonesia (ASPI).
- **Status sumber** setiap titik dicatat pada kolom `status`:
  - `RESMI BI`: siaran pers resmi Bank Indonesia (nomor dan tautan dicantumkan)
  - `PEJABAT BI`: pernyataan pejabat Bank Indonesia di media
  - `ASPI`: data Asosiasi Sistem Pembayaran Indonesia
  - `TURUNAN`: dihitung dari data lain
- Bank Indonesia tidak menerbitkan seri bulanan QRIS dalam statistik resminya, sehingga jarak antartitik tidak seragam.
- Untuk menambah atau mengoreksi data, ubah daftar di `data/bangun_dataset.py`, lalu jalankan ulang notebook.

## Cara menjalankan

Python 3.10 atau lebih baru.

```
pip install -r requirements.txt
jupyter notebook QRIS_analisis_kurva_S.ipynb
```

Jalankan *Run All* dari folder utama repositori. Notebook membaca `data/`, lalu menulis ulang seluruh isi folder `analisis/` dan `hasil_verifikasi_ulang.csv`. Bagian 9 mencocokkan otomatis 28 angka kunci; hasil yang benar adalah **28 dari 28 angka cocok**.

## Metode singkat

- Waktu *t* dihitung dalam bulan sejak Januari 2020.
- Estimasi kuadrat terkecil nonlinier (SciPy `curve_fit`).
- Perbandingan model: R², RMSE, MAPE, AICc, uji F model bersarang, dan uji *holdout*.
- Selang kepercayaan batas jenuh: *residual bootstrap* dan *moving block bootstrap* (1.000 ulangan, *seed* 2026).
- Diagnostik residual: Shapiro-Wilk, Durbin-Watson, Ljung-Box, Breusch-Pagan, dan *runs test*.
- Uji ketahanan: estimasi ulang hanya dengan titik berstatus `RESMI BI`.
- Pemeriksaan tambahan: laju pertumbuhan antarobservasi, sensitivitas K terhadap periode data, dan validasi *expanding-window*.
- Dekomposisi pertumbuhan volume transaksi menjadi margin ekstensif (pengguna) dan intensif (transaksi per pengguna).

## Catatan

Estimasi batas jenuh (K) adalah hasil model dan sensitif terhadap model serta periode data yang digunakan. Nilai K bukan target atau proyeksi resmi Bank Indonesia.

## Lisensi

- **Kode** (`*.ipynb`, `data/bangun_dataset.py`): [MIT License](LICENSE).
- **Dataset dan hasil** (`data/*.xlsx`, `data/*.csv`, `analisis/`): [CC BY 4.0](data/LICENSE.md). Cantumkan atribusi kepada penyusun apabila menggunakan dataset ini.

## Cara mengutip

Rumintarsih, A. (2026). *Analisis kurva S pengguna dan merchant QRIS Indonesia 2021–2026* [Kode dan dataset]. Zenodo. https://doi.org/(diisi setelah DOI terbit)
