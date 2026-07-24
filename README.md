# Dashboard Kinerja Dosen v2

## File data wajib
- `kinerja_perkuliahan.xlsx`
- `kinerja_ujian.xlsx`

Untuk update data, hapus/ganti workbook lama di GitHub dengan workbook baru **menggunakan nama file yang sama** dan struktur kolom yang sama. Aplikasi membaca Excel langsung saat cache diperbarui/redeploy.

## Jalankan
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Fitur
- Dashboard terpisah Perkuliahan dan Ujian
- Filter semester/kampus/fakultas/prodi
- Top/Bottom 10 dan ranking dosen
- Kategori Hijau >=85, Kuning 70-84.99, Merah <70
- Tren dan delta terhadap semester sebelumnya
- Search dosen + profil individual
- Export hasil pencarian ke Excel
