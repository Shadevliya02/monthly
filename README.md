# Dashboard Paid Monthly Outbound & Inbound

Dashboard web (satu file `index.html`, tanpa build) untuk memantau invoice yang sudah **Paid** per bulan, Outbound dan Inbound, Januari–Desember 2026.

Buka `index.html` langsung di browser.

## Isi
- Pilihan bulan Jan–Des 2026 (titik merah = bulan yang punya data paid).
- KPI: Total Paid, Paid Outbound, Paid Inbound, total return, rata-rata hari Submit → Paid.
- Ringkasan analisis otomatis (konsentrasi nilai, return, tahap terlama, kecepatan bayar, invoice terlama, data yang perlu dicek).
- Grafik waktu per tahap (Review dokumen, Revisi partner, Menuju paid) dan nilai paid per partner.
- Tabel: No, Type of Work, Scope of Work, Description, Partner, BAST Value, Invoice Number, Submit Invoice Date, Invoice Status, Return Date, Resubmit Date, Paid Date, plus kolom timeline Submit → Paid. Return ke-n dan Resubmit ke-n tampil sejajar per siklus.
- Tema terang sebagai default, tombol Mode gelap di kanan atas.
- Klik baris untuk detail timeline dengan jarak hari antar tahap.

## Menambah data
Edit array `DATA` di bagian `<script>` dalam `index.html`. Bulan tampil ditentukan dari `paid`. Format tanggal `YYYY-MM-DD`. `cycles` berisi pasangan `[return ke-n, resubmit ke-n]` berurutan; isi `null` bila resubmit belum tercatat. Set `dummy: true` bila tanggal Return/Resubmit masih contoh.

Sumber data September 2026: Report Status Invoice Paid September 2026 (Outbound 9 invoice, Inbound 2 invoice) dan sheet mitra di Database Project Outbound rev. 5 (Return Date, Resubmit Date). Return/Resubmit Inbound masih dummy.
