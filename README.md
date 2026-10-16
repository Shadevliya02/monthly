# Dashboard Paid Monthly Outbound & Inbound

Dashboard web (satu file `index.html`, tanpa build) untuk memantau invoice yang sudah **Paid** per bulan, Outbound dan Inbound, Januari–Desember 2026.

Buka `index.html` langsung di browser.

## Isi
- Pilihan bulan Jan–Des 2026 (titik merah = bulan yang punya data paid).
- KPI: Total Paid, Paid Outbound, Paid Inbound, jumlah invoice yang sempat di-return, rata-rata hari Submit → Paid.
- Grafik rata-rata jarak per tahap (Review awal, Return & revisi, Menuju paid) dan nilai paid per partner.
- Tabel: No, Type of Work, Scope of Work, Description, Partner, Invoice Date, BAST Value, Invoice Number, Submit Invoice Date, Invoice Status, Return Date, Resubmit Date, Paid Date, plus kolom timeline Submit → Paid.
- Klik baris untuk detail timeline dengan jarak hari antar tahap.

## Menambah data
Edit array `DATA` di bagian `<script>` dalam `index.html`. Bulan tampil ditentukan dari `paid`. Format tanggal `YYYY-MM-DD`; `returns` boleh berisi lebih dari satu tanggal. Set `dummy: true` bila tanggal Return/Resubmit masih contoh.

Sumber data September 2026: Report Status Invoice Paid September 2026 (Outbound 9 invoice, Inbound 2 invoice) dan sheet mitra di Database Project Outbound (Return Date, Resubmit Date). Return/Resubmit Inbound masih dummy.
