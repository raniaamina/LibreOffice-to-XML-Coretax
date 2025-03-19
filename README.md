# LibreOffice to XML Coretax
Macro script untuk membantu ekspor dokumen excel di LibreOffice Calc ke template XML Coretax


## Dependensi
- APSO (Alternative Python Script Organizer), Unduh via [Apso Gitlab](https://gitlab.com/jmzambon/apso/)

## Penggunaan
- Pastikan ekstensi APSO sudah terpasang
- Salin berkas export-to-xml-cortax.py ke dalam folder config LibreOffice, lokasinya biasanya ada di;
  - Linux: `$HOME/.config/libreoffice/4/user/Scripts/python`
  - Windows: `%APPDATA%\LibreOffice\4\user\Scripts\python`
  - MacOS: `$HOME/Library/'Application Support'/LibreOffice/4/user/Scripts/python`
  
folder `Scripts/python` secara default tidak langsung tersedia, silakan buat folder tersebut dengan nama dan kapitalisasi huruf yang sesuai.

Jika sudah, buka template xslx Anda, lalu jalankan skrip lewat menu; Tools >> Macros >> Run Macro .... Jika pemasangan benar, pada dialog yang muncul akan terlihat opsi export-to-xml-coretax di sisi kiri, dan daftar template yang didukung di sisi kanan.

Pilih salah template yang sesuai, lalu klik Run. Berkas XML akan otomatis muncul di folder tempat berkas xslx Anda di simpan.


## Dukungan Template
Iya, dukungan templatenya masih sangat sedikit, karena skrip ini memang dibuat untuk sekadar membantu dan tidak terafiliasi dengan pihak perpajakan. Untuk Sementara, template yang dapat diekspor menggunakan skrip ini antara lain; 
- SPT PPN DRKB (Daftar Rekap Kendaraan Bermotor)
- SPT PPN Retail (Digunggung 1.A5, 1.A9, 1B)


## Sangkalan
- Do with your own risk, saya tidak memberikan jaminan apa pun dengan skrip ini, ('-_-').
- Saya menguji makro ini di Debian Linux (sid) dengan LibreOffice terbaru, saya tidak memberi jaminan makro ini dapat berjalan di mana saja. Silakan berkabar via tiket jika Anda dapat menjalankan makro ini di sistem operasi Anda.


_If you feel that this tool is helpful, feel free to send a cup of coffee via [Support Dev](https://saweria.co/raniaamina) :")_