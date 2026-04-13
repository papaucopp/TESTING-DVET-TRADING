# PO Complete Set System

Sistem ini ialah prototaip web ringan untuk:

- upload `Purchase Order.pdf`
- baca kandungan PO secara automatik dalam pelayar
- jana `Invoice` dan `Delivery Order`
- muat turun kedua-dua dokumen sebagai PDF

## Cara jalankan

1. Buka Terminal.
2. Pergi ke folder projek:

```bash
cd /Users/wanhakimyusof/Desktop/po-complete-set-system
```

3. Jalankan web server ringkas:

```bash
python3 -m http.server 8080
```

4. Buka pelayar ke:

`http://localhost:8080`

## Nota

- Parser PO menggunakan `pdf.js` dalam browser.
- Eksport PDF menggunakan `html2pdf.js`.
- Jika ada medan yang parser tidak baca dengan tepat, user masih boleh edit manual sebelum download PDF.
- Nombor invoice dan DO dijana secara automatik berdasarkan nombor PO, tetapi masih boleh diubah dalam borang.
