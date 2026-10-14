EMAIL OTOMATIS PO KE SUPPLIER (setelah approver terakhir menyetujui)
=====================================================================

Alur
----
1. Approver terakhir menyetujui PO di
     - PurchaseOrderAuthorization.Asp (daftar, bisa banyak PO), atau
     - PurchaseOrderDetailsApproval.Asp (detail satu PO).
2. Halaman approval memasukkan PO ke tabel PO_EMAIL_OUTBOX (po_mail_queue.inc)
   lalu langsung memicu email/send_po_outbox.asp.
3. send_po_outbox.asp membuat PDF PO (dengan QR code berisi nomor PO),
   menyusun email formal, dan mengirim lewat Gmail (pch.fbg@gmail.com).
4. Bila gagal (SMTP / internet), dicoba ulang otomatis (maks. 5x, jeda bertambah)
   oleh Task Scheduler yang menjalankan email/run_po_outbox.vbs.
5. PO yang di-approve ulang (setelah UnRelease/RePlan) akan dikirim ulang.

Penerima
--------
- Mode uji coba (PoMailCfg_TestMode = True, BAWAAN):
    semua email ke nrw.fbg@gmail.com, tanpa CC. Email diberi tanda [UJI COBA]
    dan mencantumkan penerima asli (email supplier) di bagian atas.
- Mode produksi (PoMailCfg_TestMode = False):
    To = email supplier (kolom SUPPLIER.MAIL), CC = rps.fbg@gmail.com; sfk.arjuna@gmail.com.
    Bila email supplier kosong / tidak valid, status FAILED dengan pesan jelas.

Langkah pemasangan
------------------
1. Database: jalankan email/sql/po_email_outbox.sql (tabel + sequence) di DEV, lalu PROD.
2. Gmail pch.fbg@gmail.com: aktifkan Verifikasi 2 Langkah, lalu buat App Password
   (Akun Google > Keamanan > Sandi aplikasi). Isi 16 karakternya ke
   PoMailCfg_SmtpPassword di email/po_mail_config.inc (file ini tidak masuk git).
3. Upload ke server: folder email/ (termasuk po_mail_config.inc), serta
   PurchaseOrderAuthorization.Asp dan PurchaseOrderDetailsApproval.Asp.
4. Izin folder: akun IIS (IUSR / Application Pool) butuh izin tulis ke email/po_pdf/
   (folder dibuat otomatis saat pengiriman pertama).
5. Server harus bisa keluar ke smtp.gmail.com port 465 (cek firewall).
6. PoMailCfg_BaseUrl harus bisa dibuka dari server itu sendiri (bawaan http://localhost).
   PoMailCfg_LinkBaseUrl = alamat aplikasi untuk tautan internal di email.
7. Task Scheduler (cadangan & percobaan ulang), contoh tiap 5 menit:
     schtasks /Create /TN "FB-Plants PO Email" /SC MINUTE /MO 5 /RU SYSTEM ^
       /TR "cscript.exe //nologo <path-website>\email\run_po_outbox.vbs"
8. Uji: approve final satu PO di DEV, lalu cek kotak masuk nrw.fbg@gmail.com dan
     SELECT * FROM PO_EMAIL_OUTBOX ORDER BY ID DESC;
   Hasil proses juga bisa dilihat dengan membuka
     <BaseUrl>/email/send_po_outbox.asp?key=<PoMailCfg_ProcessKey>
9. Setelah lolos uji: ubah PoMailCfg_TestMode = False.

File
----
email/po_mail_config.inc         konfigurasi (TIDAK di git; berisi password)
email/po_mail_config.sample.inc  template konfigurasi
email/po_mail_queue.inc          antrian + pemicu (di-include halaman approval)
email/po_mail_sender.inc         data PO, PDF, isi email, kirim SMTP, proses antrian
email/qrcode.inc                 pembuat QR code (VBScript murni)
email/pdfwriter.inc              pembuat PDF (VBScript murni)
email/send_po_outbox.asp         halaman pemroses antrian (butuh ?key=)
email/run_po_outbox.vbs          dipanggil Task Scheduler
email/sql/po_email_outbox.sql    DDL tabel antrian

Catatan
-------
- Gmail gratis dibatasi +- 500 email/hari.
- Tautan printingpo.asp di email hanya bisa dibuka dari jaringan kantor
  (supplier memakai lampiran PDF).
- Kegagalan email TIDAK menggagalkan approval; detail error tercatat di
  PO_EMAIL_OUTBOX.LAST_ERROR dan /debug/error.log.
