<%
'=====================================================================
' po_mail_config.sample.inc - CONTOH konfigurasi email PO
' Salin menjadi email/po_mail_config.inc lalu isi nilainya.
' File po_mail_config.inc TIDAK masuk git (berisi password).
'=====================================================================
Dim PoMailCfg_Enabled, PoMailCfg_TestMode, PoMailCfg_TestTo, PoMailCfg_TestUseCc, PoMailCfg_Cc
Dim PoMailCfg_SmtpServer, PoMailCfg_SmtpPort, PoMailCfg_SmtpSsl, PoMailCfg_SmtpUser, PoMailCfg_SmtpPassword
Dim PoMailCfg_FromAddress, PoMailCfg_FromName, PoMailCfg_ReplyTo
Dim PoMailCfg_BaseUrl, PoMailCfg_LinkBaseUrl, PoMailCfg_ProcessKey, PoMailCfg_PdfFolder, PoMailCfg_MaxAttempts

' Aktif / nonaktif seluruh fitur (False = tidak ada email yang diantrikan)
PoMailCfg_Enabled      = True

' --- Mode uji coba ---
' True  : semua email dialihkan ke PoMailCfg_TestTo (supplier tidak menerima apa pun)
' False : email dikirim ke alamat email supplier (kolom SUPPLIER.MAIL)
PoMailCfg_TestMode     = True
PoMailCfg_TestTo       = "nrw.fbg@gmail.com"
PoMailCfg_TestUseCc    = False                  ' CC dinonaktifkan selama uji coba
PoMailCfg_Cc           = "rps.fbg@gmail.com;sfk.arjuna@gmail.com"

' --- Server SMTP (Gmail) ---
' Gmail wajib memakai "App Password" 16 karakter (akun harus memakai verifikasi 2 langkah),
' bukan password login biasa. Buat di: Akun Google > Keamanan > Sandi aplikasi.
PoMailCfg_SmtpServer   = "smtp.gmail.com"
PoMailCfg_SmtpPort     = 465                    ' SSL langsung (CDO tidak mendukung STARTTLS 587)
PoMailCfg_SmtpSsl      = True
PoMailCfg_SmtpUser     = "pch.fbg@gmail.com"
PoMailCfg_SmtpPassword = ""                     ' <-- ISI App Password Gmail di sini

PoMailCfg_FromAddress  = "pch.fbg@gmail.com"
PoMailCfg_FromName     = "Purchasing FBG"
PoMailCfg_ReplyTo      = "pch.fbg@gmail.com"

' --- Lain-lain ---
' Alamat aplikasi dilihat dari server IIS sendiri (untuk memicu pengiriman segera setelah approve)
PoMailCfg_BaseUrl      = "http://localhost"
' Alamat aplikasi yang dibuka user di jaringan kantor (untuk tautan internal printingpo.asp di email)
PoMailCfg_LinkBaseUrl  = "http://172.30.61.221"
' Kunci rahasia untuk halaman pemroses antrian (email/send_po_outbox.asp?key=...)
PoMailCfg_ProcessKey   = "GANTI-DENGAN-KUNCI-ACAK"
' Folder penyimpanan PDF lampiran (path virtual; IUSR/AppPool butuh izin tulis)
PoMailCfg_PdfFolder    = "/email/po_pdf"
' Batas percobaan kirim sebelum status FAILED
PoMailCfg_MaxAttempts  = 5
%>
