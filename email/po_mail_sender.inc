<!--#include file="po_mail_config.inc"-->
<!--#include file="qrcode.inc"-->
<!--#include file="pdfwriter.inc"-->
<!--#include file="po_link_token.inc"-->
<%
'=====================================================================
' po_mail_sender.inc - membuat PDF PO (dengan QR nomor PO), menyusun email
' formal, mengirim lewat SMTP (Gmail) dan memproses antrian PO_EMAIL_OUTBOX.
' Dipakai oleh email/send_po_outbox.asp.
'
' Fungsi pembuat PDF & isi email hanya menerima data (Dictionary), sehingga
' bisa diuji tanpa database. Akses database ada di PoMail_LoadPo & PoMail_ProcessOutbox.
'=====================================================================

' ---------------- format ----------------
Function PoMail_Num(v, dec)
	Dim neg, ip, fp, s, out, i, cnt, fr, p
	If IsNull(v) Or Trim("" & v) = "" Then v = 0
	If Not IsNumeric(v) Then PoMail_Num = "" & v : Exit Function
	v = CDbl(v)
	neg = (v < 0) : v = Abs(v)
	p = 10 ^ dec
	v = Int(v * p + 0.5) / p
	ip = Int(v)
	fp = Int((v - ip) * p + 0.5)
	If fp >= p Then ip = ip + 1 : fp = 0
	s = FormatNumber(ip, 0, -1, 0, 0)          ' tanpa pemisah ribuan
	s = Replace(Replace(s, ",", ""), ".", "")
	out = "" : cnt = 0
	For i = Len(s) To 1 Step -1
		out = Mid(s, i, 1) & out
		cnt = cnt + 1
		If cnt Mod 3 = 0 And i > 1 Then out = "," & out
	Next
	If dec > 0 Then out = out & "." & Right(String(dec, "0") & fp, dec)
	If neg Then out = "-" & out
	PoMail_Num = out
End Function

' Qty: sampai 2 desimal, nol di belakang dibuang
Function PoMail_Qty(v)
	Dim s
	s = PoMail_Num(v, 2)
	If Right(s, 3) = ".00" Then
		s = Left(s, Len(s) - 3)
	ElseIf Right(s, 1) = "0" And InStr(s, ".") > 0 Then
		s = Left(s, Len(s) - 1)
	End If
	PoMail_Qty = s
End Function

Function PoMail_Money(v)
	Dim n : n = 0
	If IsNumeric(v) Then n = CDbl(v)
	PoMail_Money = PoMail_Num(n, 2)   ' selalu ribuan + 2 desimal (sama dengan printingpo.asp)
End Function

Function PoMail_DateId(d)
	Dim bln
	bln = Array("", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember")
	If IsDate(d) Then
		PoMail_DateId = Day(CDate(d)) & " " & bln(Month(CDate(d))) & " " & Year(CDate(d))
	Else
		PoMail_DateId = "" & d
	End If
End Function

Function PoMail_H(s)
	s = "" & s
	s = Replace(s, "&", "&amp;") : s = Replace(s, "<", "&lt;") : s = Replace(s, ">", "&gt;") : s = Replace(s, """", "&quot;")
	PoMail_H = s
End Function

Function PoMail_D(d, key)
	If d.Exists(key) Then PoMail_D = "" & d(key) Else PoMail_D = ""
End Function

' ---------------- PDF ----------------
' Tata letak mengikuti cetakan PO (printingpo.asp): kop + logo perusahaan, blok
' Kepada / Tanggal, tabel item, ringkasan biaya, rekening, catatan, tanda tangan.
' Tambahan: QR code nomor PO di pojok kanan atas.
' hdr: Dictionary; items: array of Dictionary (itemcode, itemname, desc, qty, uom, unitprice, disc, amount,
'      otcstock1, otcstock3, othercharge, otherchargestock)
Sub PoMail_BuildPdf(hdr, items, path)
	Dim pdf, L, R, W, x0, xv, y, yTop, qrSize, qrMods, qrBox, img, lw, lh, i, k, it, lines, rowH, pg
	Dim cols, tblX, tblW, ySec, ppnTxt, sums, s, first, sigW, sigLbl, sigName, xx, mataUang, dl, j
	Set pdf = New PoPdf
	pdf.Title = "Purchase Order " & PoMail_D(hdr, "pocode")
	L = 28 : R = pdf.PageW - 28 : W = R - L
	x0 = L + W * 0.62 : xv = x0 + 74               ' kolom kanan (label / nilai)
	pdf.NewPage
	pdf.SetFill 0, 0, 0

	' --- logo perusahaan (sama dengan cetakan PO) ---
	yTop = 26
	img = ""
	If PoMail_D(hdr, "logopath") <> "" Then img = pdf.AddJpeg(PoMail_D(hdr, "logopath"))
	If img <> "" Then
		lh = 56 : lw = lh * pdf.ImageWidth(img) / pdf.ImageHeight(img)
		If lw > 230 Then lw = 230 : lh = lw * pdf.ImageHeight(img) / pdf.ImageWidth(img)
		pdf.Image img, L, yTop, lw, lh
	End If
	' --- QR code nomor PO (pojok kanan atas) ---
	qrMods = PoQr_Encode(PoMail_D(hdr, "pocode"), qrSize)
	If Not IsEmpty(qrMods) Then
		qrBox = 58
		pdf.QrCode qrMods, qrSize, R - qrBox, yTop - 4, qrBox / qrSize
		pdf.TextCenter R - qrBox / 2, yTop + qrBox + 3, 5.5, False, "Scan: No. PO"
	End If
	y = yTop + 70

	' --- baris judul ---
	s = PoMail_D(hdr, "companyname")
	If UCase(Left(s, 3)) <> "PT." And UCase(Left(s, 3)) <> "PT " Then s = "PT. " & s
	pdf.Text L, y, 9, True, pdf.Fit(s, x0 - L - 10, 9, True)
	pdf.Text x0, y, 9, True, "PURCHASE ORDER"
	y = y + 13
	pdf.Text x0, y, 8, False, "No :"
	pdf.Text xv, y, 8, False, pdf.Fit(PoMail_D(hdr, "pocode"), R - xv, 8, False)
	y = y + 11
	pdf.Text x0, y, 8, False, "Classification :"
	pdf.Text xv, y, 8, False, pdf.Fit(PoMail_D(hdr, "classification"), R - xv, 8, False)
	' alamat pemesan (Purchase By), kolom kiri
	ySec = y + 13
	lines = pdf.Wrap(PoMail_D(hdr, "pbaddress"), x0 - L - 14, 8, False)
	For k = 0 To UBound(lines)
		If k < 3 And lines(k) <> "" Then pdf.Text L, ySec, 8, False, lines(k) : ySec = ySec + 10
	Next
	ySec = ySec + 3
	pdf.Text L, ySec, 8, False, "Phone. " & PoMail_D(hdr, "pbphone") : ySec = ySec + 10
	pdf.Text L, ySec, 8, False, "Fax. " & PoMail_D(hdr, "pbfax")
	y = ySec + 18

	' --- blok Kepada (kiri) & Tanggal (kanan) ---
	mataUang = PoMail_D(hdr, "currency")
	If mataUang = "" Or UCase(mataUang) = "IDR" Then mataUang = "Rupiah"
	lines = pdf.Wrap(PoMail_D(hdr, "supplieraddress"), x0 - L - 60, 8, False)
	Call PoMail_PdfPair(pdf, L, y, "Kepada", True, PoMail_D(hdr, "suppliername"), x0 - L - 60)
	Call PoMail_PdfPair(pdf, x0, y, "Tanggal", False, PoMail_DateId(PoMail_D(hdr, "podate")), R - xv)
	y = y + 11
	Call PoMail_PdfPair(pdf, L, y, "Alamat", False, lines(0), x0 - L - 60)
	Call PoMail_PdfPair(pdf, x0, y, "Berlaku S/D", False, "", R - xv)
	For k = 1 To UBound(lines)
		If k < 3 Then y = y + 10 : pdf.Text L + 52, y, 8, False, lines(k)
	Next
	y = y + 11
	Call PoMail_PdfPair(pdf, L, y, "Phone", False, PoMail_D(hdr, "supplierphone"), x0 - L - 60)
	Call PoMail_PdfPair(pdf, x0, y, "Mata Uang :", False, mataUang, R - xv)
	y = y + 11
	Call PoMail_PdfPair(pdf, L, y, "Fax", False, PoMail_D(hdr, "supplierfax"), x0 - L - 60)
	Call PoMail_PdfPair(pdf, x0, y, "Gudang :", False, IIfStr(PoMail_D(hdr, "destname") <> "", PoMail_D(hdr, "destname"), "Kebun"), R - xv)
	y = y + 11
	Call PoMail_PdfPair(pdf, L, y, "U/P", False, "", x0 - L - 60)
	Call PoMail_PdfPair(pdf, x0, y, "No. PR :", False, PoMail_D(hdr, "prcode"), R - xv)
	y = y + 14

	' --- tabel item ---
	cols = Array(Array(24, "No"), Array(255, "Description"), Array(55, "Qty"), Array(45, "UOM"), Array(75, "Price"), Array(85, "Total Amount"))
	tblX = L : tblW = W
	pdf.SetStroke 0, 0, 0 : pdf.SetLineWidth 0.7
	y = PoMail_PdfTableHead(pdf, cols, tblX, y)
	yTop = y                                       ' awal badan tabel (untuk garis kolom)
	For i = 0 To UBound(items)
		Set it = items(i)
		lines = pdf.Wrap(PoMail_D(it, "itemname"), cols(1)(0) - 8, 8, False)
		If PoMail_D(it, "desc") <> "" Then
			dl = pdf.Wrap(PoMail_D(it, "desc"), cols(1)(0) - 8, 8, False)
			For j = 0 To UBound(dl)
				ReDim Preserve lines(UBound(lines) + 1) : lines(UBound(lines)) = dl(j)
			Next
		End If
		rowH = (UBound(lines) + 1) * 10 + 4
		If y + rowH > pdf.PageH - 60 Then
			Call PoMail_PdfTableCols(pdf, cols, tblX, yTop, y)
			pdf.Line tblX, y, tblX + tblW, y
			pdf.NewPage
			pdf.SetFill 0, 0, 0 : pdf.SetStroke 0, 0, 0 : pdf.SetLineWidth 0.7
			pdf.Text L, 40, 8, True, "PURCHASE ORDER No : " & PoMail_D(hdr, "pocode") & " (lanjutan)"
			y = PoMail_PdfTableHead(pdf, cols, tblX, 50)
			yTop = y
		End If
		xx = tblX
		pdf.TextCenter xx + cols(0)(0) / 2, y + 10, 8, False, CStr(i + 1) : xx = xx + cols(0)(0)
		For k = 0 To UBound(lines)
			pdf.Text xx + 4, y + 10 + k * 10, 8, False, lines(k)
		Next
		xx = xx + cols(1)(0)
		pdf.TextCenter xx + cols(2)(0) / 2, y + 10, 8, False, PoMail_Num(PoMail_D(it, "qty"), 2) : xx = xx + cols(2)(0)
		pdf.TextCenter xx + cols(3)(0) / 2, y + 10, 8, False, pdf.Fit(PoMail_D(it, "uom"), cols(3)(0) - 4, 8, False) : xx = xx + cols(3)(0)
		pdf.TextCenter xx + cols(4)(0) / 2, y + 10, 8, False, PoMail_Num(PoMail_D(it, "unitprice"), 2) : xx = xx + cols(4)(0)
		pdf.TextRight xx + cols(5)(0) - 4, y + 10, 8, False, PoMail_Num(PoMail_D(it, "amount"), 2)
		y = y + rowH
	Next
	y = y + 4
	Call PoMail_PdfTableCols(pdf, cols, tblX, yTop, y)

	' --- ringkasan biaya (sama urutan dengan cetakan PO) ---
	ppnTxt = PoMail_D(hdr, "ppn")
	sums = Array( _
		Array("Sub Total", PoMail_D(hdr, "totalamount"), True), _
		Array("Ppn " & ppnTxt & " %", PoMail_D(hdr, "totalppn"), True), _
		Array("Total Amount + PPn " & ppnTxt & " %", PoMail_D(hdr, "totalamountppn"), True))
	For k = 0 To UBound(sums)
		y = PoMail_PdfSumRow(pdf, cols, tblX, tblW, y, sums(k)(0), sums(k)(1), k = 0)
	Next
	If UBound(items) >= 0 Then
		Set first = items(0)
		If PoMail_NumOnly(PoMail_D(first, "otcstock1")) > 0 Then y = PoMail_PdfSumRow(pdf, cols, tblX, tblW, y, "PBBKP (7.5%)", PoMail_D(first, "otcstock1"), False)
		If PoMail_NumOnly(PoMail_D(first, "otcstock3")) > 0 Then y = PoMail_PdfSumRow(pdf, cols, tblX, tblW, y, "OAT", PoMail_D(first, "otcstock3"), False)
		If PoMail_NumOnly(PoMail_D(first, "othercharge")) > 0 Then y = PoMail_PdfSumRow(pdf, cols, tblX, tblW, y, "Other Charge", PoMail_D(first, "othercharge"), False)
		If PoMail_NumOnly(PoMail_D(first, "otherchargestock")) > 0 Then y = PoMail_PdfSumRow(pdf, cols, tblX, tblW, y, "Other Charge", PoMail_D(first, "otherchargestock"), False)
	End If
	y = PoMail_PdfSumRow(pdf, cols, tblX, tblW, y, "Grand Total", PoMail_D(hdr, "grandtotal"), False)

	' --- rekening, catatan, tanda tangan ---
	lines = PoMail_SplitLines(pdf, PoMail_D(hdr, "remarks"), W, 8)
	If y + 60 + (UBound(lines) + 1) * 10 + 95 > pdf.PageH - 30 Then
		pdf.NewPage : pdf.SetFill 0, 0, 0 : pdf.SetStroke 0, 0, 0
		pdf.Text L, 40, 8, True, "PURCHASE ORDER No : " & PoMail_D(hdr, "pocode") & " (lanjutan)"
		y = 50
	End If
	y = y + 18
	pdf.Text L, y, 8, True, "Rekening Pembayaran :"
	y = y + 11
	pdf.Text L + 4, y, 8, False, pdf.Fit("No.Rek. " & PoMail_D(hdr, "bankaccount") & ", Bank : " & PoMail_D(hdr, "bankname") & ",A/N " & PoMail_D(hdr, "bankownership"), W - 8, 8, False)
	y = y + 14
	pdf.Text L, y, 8, True, "Catatan :"
	For k = 0 To UBound(lines)
		y = y + 10
		pdf.Text L + 4, y, 8, False, lines(k)
	Next
	y = y + 22
	sigW = W / 4
	sigLbl = Array("Dibuat Oleh,", "Disetujui Oleh,", "Disetujui Oleh,", "Disetujui Supplier,")
	sigName = Array("PURCHASING", "PURCHASING DEPT", "HEAD COMM & PURCH DEPT", PoMail_D(hdr, "suppliername"))
	For k = 0 To 3
		pdf.TextCenter L + sigW * k + sigW / 2, y, 8, 2, sigLbl(k)
		pdf.TextCenter L + sigW * k + sigW / 2, y + 62, 8, True, pdf.Fit(sigName(k), sigW - 8, 8, True)
	Next

	' --- catatan kaki semua halaman ---
	For pg = 0 To pdf.PageCount - 1
		pdf.SelectPage pg
		pdf.SetFill 110, 110, 110
		pdf.Text L, pdf.PageH - 18, 6.5, False, "Dokumen ini dikirim otomatis oleh sistem FB-Plants setelah PO disetujui. Nomor PO dapat diverifikasi melalui QR code."
		pdf.TextRight R, pdf.PageH - 18, 6.5, False, "Halaman " & (pg + 1) & " / " & pdf.PageCount
	Next
	pdf.SaveTo path
	Set pdf = Nothing
End Sub

' Label + ": nilai" (gaya baris info cetakan PO)
Sub PoMail_PdfPair(pdf, x, y, label, boldValue, value, maxW)
	Dim lx
	pdf.Text x, y, 8, False, label
	If x > 200 Then                                ' kolom kanan: nilai sejajar di satu kolom
		lx = x + 74
		pdf.Text lx, y, 8, boldValue, pdf.Fit(value, maxW, 8, boldValue)
	Else
		lx = x + 48
		pdf.Text lx, y, 8, False, ":"
		pdf.Text lx + 5, y, 8, boldValue, pdf.Fit(value, maxW, 8, boldValue)
	End If
End Sub

' Pecah catatan per baris (enter dipertahankan seperti white-space:pre di cetakan PO)
Function PoMail_SplitLines(pdf, s, maxW, size)
	Dim parts, i, w, j, outArr(), n
	n = 0 : ReDim outArr(-1)
	parts = Split(Replace(Replace("" & s, vbCrLf, vbLf), vbCr, vbLf), vbLf)
	For i = 0 To UBound(parts)
		w = pdf.Wrap(parts(i), maxW - 8, size, False)
		For j = 0 To UBound(w)
			ReDim Preserve outArr(n) : outArr(n) = w(j) : n = n + 1
		Next
	Next
	If n > 12 Then ReDim Preserve outArr(11)
	PoMail_SplitLines = outArr
End Function

Function PoMail_NumOnly(v)
	If IsNumeric(v) Then PoMail_NumOnly = CDbl(v) Else PoMail_NumOnly = 0
End Function

Function IIfStr(c, a, b)
	If c Then IIfStr = a Else IIfStr = b
End Function

' Kepala tabel bergaris (No | Description | Qty | UOM | Price | Total Amount)
Function PoMail_PdfTableHead(pdf, cols, x, y)
	Dim k, xx, w
	w = 0
	For k = 0 To UBound(cols) : w = w + cols(k)(0) : Next
	pdf.Rect x, y, w, 16, "S"
	xx = x
	For k = 0 To UBound(cols)
		If k >= 2 Then pdf.Line xx, y, xx, y + 16
		pdf.TextCenter xx + cols(k)(0) / 2, y + 11, 8, True, cols(k)(1)
		xx = xx + cols(k)(0)
	Next
	PoMail_PdfTableHead = y + 16
End Function

' Garis tepi & pemisah kolom badan tabel (gaya "sideonly" cetakan PO)
Sub PoMail_PdfTableCols(pdf, cols, x, y1, y2)
	Dim k, xx
	xx = x
	pdf.Line x, y1, x, y2
	For k = 0 To UBound(cols)
		xx = xx + cols(k)(0)
		If k >= 1 Or k = UBound(cols) Then pdf.Line xx, y1, xx, y2
	Next
	pdf.Line x, y2, xx, y2
End Sub

' Baris ringkasan: label rata kanan di kiri kolom terakhir, nilai dalam kotak kolom terakhir
Function PoMail_PdfSumRow(pdf, cols, x, w, y, label, value, fullBorder)
	Dim lastW
	lastW = cols(UBound(cols))(0)
	If fullBorder Then pdf.Rect x, y, w, 14, "S" Else pdf.Rect x + w - lastW, y, lastW, 14, "S"
	pdf.TextRight x + w - lastW - 4, y + 10, 8, False, label
	pdf.TextRight x + w - 4, y + 10, 8, False, PoMail_Num(value, 2)
	PoMail_PdfSumRow = y + 14
End Function

' ---------------- isi email ----------------
Function PoMail_Subject(hdr, isTest)
	PoMail_Subject = IIfStr(isTest, "[UJI COBA] ", "") & "Purchase Order No. " & PoMail_D(hdr, "pocode") & " - " & PoMail_D(hdr, "companyname")
End Function

Function PoMail_BuildHtml(hdr, items, isTest, realTo, realCc, linkUrl)
	Dim h, i, it, td, tdr, th
	td = "padding:7px 10px;border-bottom:1px solid #e1e7e5;font-size:13px;color:#17262a;"
	tdr = td & "text-align:right;white-space:nowrap;"
	th = "padding:8px 10px;background:#1B7863;color:#ffffff;font-size:11px;font-weight:bold;text-transform:uppercase;letter-spacing:.4px;text-align:left;"
	h = "<!doctype html><html><head><meta charset=""utf-8""></head>"
	h = h & "<body style=""margin:0;padding:0;background:#f3f6f5;font-family:Segoe UI,Arial,Helvetica,sans-serif;color:#17262a;"">"
	h = h & "<table role=""presentation"" width=""100%"" cellpadding=""0"" cellspacing=""0"" style=""background:#f3f6f5;padding:24px 0;""><tr><td align=""center"">"
	h = h & "<table role=""presentation"" width=""640"" cellpadding=""0"" cellspacing=""0"" style=""max-width:640px;width:100%;background:#ffffff;border:1px solid #e1e7e5;border-radius:10px;overflow:hidden;"">"
	If isTest Then
		h = h & "<tr><td style=""padding:10px 24px;background:#fff4d6;color:#8a5a00;font-size:12px;border-bottom:1px solid #f3dda6;"">"
		h = h & "<b>UJI COBA</b> &mdash; email ini dialihkan dari penerima sebenarnya. Penerima asli: <b>" & IIfStr(realTo <> "", PoMail_H(realTo), "(email supplier belum diisi)") & "</b>"
		If realCc <> "" Then h = h & "; CC: " & PoMail_H(realCc)
		h = h & "</td></tr>"
	End If
	h = h & "<tr><td style=""padding:22px 24px 6px;border-top:5px solid #1B7863;"">"
	h = h & "<div style=""font-size:12px;font-weight:bold;letter-spacing:.6px;color:#1B7863;text-transform:uppercase;"">Purchase Order</div>"
	h = h & "<div style=""font-size:20px;font-weight:bold;margin-top:2px;"">" & PoMail_H(PoMail_D(hdr, "pocode")) & "</div>"
	h = h & "<div style=""font-size:12px;color:#5d6d69;margin-top:2px;"">" & PoMail_H(PoMail_D(hdr, "companyname")) & "</div>"
	h = h & "</td></tr>"
	h = h & "<tr><td style=""padding:14px 24px;font-size:14px;line-height:1.6;"">"
	h = h & "<p style=""margin:0 0 12px;"">Kepada Yth.<br><b>" & PoMail_H(PoMail_D(hdr, "suppliername")) & "</b><br>di tempat</p>"
	h = h & "<p style=""margin:0 0 12px;"">Dengan hormat,</p>"
	h = h & "<p style=""margin:0 0 12px;"">Bersama surat elektronik ini kami sampaikan bahwa <i>Purchase Order</i> (PO) nomor <b>" & PoMail_H(PoMail_D(hdr, "pocode")) & "</b> tertanggal <b>" & PoMail_H(PoMail_DateId(PoMail_D(hdr, "podate"))) & "</b> telah memperoleh persetujuan akhir dan dengan ini resmi kami terbitkan kepada Bapak/Ibu.</p>"
	h = h & "<p style=""margin:0 0 12px;"">Mohon kiranya pesanan tersebut dapat diproses sesuai dengan spesifikasi, jumlah, harga, dan ketentuan yang tercantum pada dokumen PO terlampir. Ringkasan pesanan kami sampaikan sebagai berikut:</p>"
	h = h & "</td></tr>"
	' ringkasan
	h = h & "<tr><td style=""padding:0 24px;""><table role=""presentation"" width=""100%"" cellpadding=""0"" cellspacing=""0"" style=""border:1px solid #e1e7e5;border-radius:8px;"">"
	h = h & PoMail_HtmlRow("No. PO", "<b>" & PoMail_H(PoMail_D(hdr, "pocode")) & "</b>")
	h = h & PoMail_HtmlRow("Tanggal PO", PoMail_H(PoMail_DateId(PoMail_D(hdr, "podate"))))
	h = h & PoMail_HtmlRow("Supplier", PoMail_H(PoMail_D(hdr, "suppliername")) & IIfStr(PoMail_D(hdr, "suppliercode") <> "", " <span style=""color:#5d6d69;"">(" & PoMail_H(PoMail_D(hdr, "suppliercode")) & ")</span>", ""))
	If PoMail_D(hdr, "destname") <> "" Then h = h & PoMail_HtmlRow("Dikirim ke", PoMail_H(PoMail_D(hdr, "destname")) & IIfStr(PoMail_D(hdr, "destaddress") <> "", "<br><span style=""color:#5d6d69;font-size:12px;"">" & PoMail_H(PoMail_D(hdr, "destaddress")) & "</span>", ""))
	If PoMail_D(hdr, "termofpayment") <> "" Then h = h & PoMail_HtmlRow("Pembayaran", PoMail_H(PoMail_D(hdr, "termofpayment")))
	h = h & PoMail_HtmlRow("Grand Total", "<b style=""font-size:15px;color:#12584a;"">" & IIfStr(PoMail_D(hdr, "currency") <> "", PoMail_H(PoMail_D(hdr, "currency")) & " ", "") & PoMail_Money(PoMail_D(hdr, "grandtotal")) & "</b>")
	h = h & "</table></td></tr>"
	' daftar item
	h = h & "<tr><td style=""padding:16px 24px 0;""><table role=""presentation"" width=""100%"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse:collapse;border:1px solid #e1e7e5;"">"
	h = h & "<tr><td style=""" & th & "width:28px;"">No</td><td style=""" & th & """>Nama Barang</td><td style=""" & th & "text-align:right;"">Qty</td><td style=""" & th & """>Sat</td><td style=""" & th & "text-align:right;"">Harga Satuan</td><td style=""" & th & "text-align:right;"">Jumlah</td></tr>"
	For i = 0 To UBound(items)
		Set it = items(i)
		h = h & "<tr><td style=""" & td & "color:#5d6d69;"">" & (i + 1) & "</td>"
		h = h & "<td style=""" & td & """>" & PoMail_H(PoMail_D(it, "itemname")) & IIfStr(PoMail_D(it, "desc") <> "", "<br><span style=""color:#5d6d69;font-size:12px;"">" & PoMail_H(PoMail_D(it, "desc")) & "</span>", "") & IIfStr(PoMail_D(it, "itemcode") <> "", "<br><span style=""color:#8a9a96;font-size:11px;"">" & PoMail_H(PoMail_D(it, "itemcode")) & "</span>", "") & "</td>"
		h = h & "<td style=""" & tdr & """>" & PoMail_Num(PoMail_D(it, "qty"), 2) & "</td>"
		h = h & "<td style=""" & td & """>" & PoMail_H(PoMail_D(it, "uom")) & "</td>"
		h = h & "<td style=""" & tdr & """>" & PoMail_Money(PoMail_D(it, "unitprice")) & "</td>"
		h = h & "<td style=""" & tdr & "font-weight:bold;"">" & PoMail_Money(PoMail_D(it, "amount")) & "</td></tr>"
	Next
	h = h & "</table></td></tr>"
	' penutup
	h = h & "<tr><td style=""padding:16px 24px 6px;font-size:14px;line-height:1.6;"">"
	h = h & "<p style=""margin:0 0 12px;"">Dokumen PO lengkap kami lampirkan dalam format PDF. Keaslian nomor PO dapat diverifikasi melalui <i>QR code</i> yang tercantum pada dokumen tersebut.</p>"
	h = h & "<p style=""margin:0 0 12px;"">Kami mohon konfirmasi penerimaan PO ini dengan membalas surat elektronik ini, sekaligus menginformasikan perkiraan jadwal pengiriman barang. Apabila terdapat hal yang perlu diklarifikasi, Bapak/Ibu dapat menghubungi kami melalui alamat email ini.</p>"
	h = h & "<p style=""margin:0 0 12px;"">Atas perhatian dan kerja sama yang baik, kami ucapkan terima kasih.</p>"
	h = h & "<p style=""margin:0;"">Hormat kami,<br><b>Departemen Purchasing</b><br>" & PoMail_H(PoMail_D(hdr, "companyname")) & "<br><span style=""color:#5d6d69;font-size:13px;"">" & PoMail_H(PoMailCfg_FromAddress) & "</span></p>"
	h = h & "</td></tr>"
	If linkUrl <> "" Then
		h = h & "<tr><td style=""padding:10px 24px 0;font-size:12px;color:#5d6d69;"">Tautan internal (khusus pengguna FB-Plants di jaringan kantor): <a href=""" & PoMail_H(linkUrl) & """ style=""color:#1B7863;"">Lihat / cetak PO</a></td></tr>"
	End If
	h = h & "<tr><td style=""padding:16px 24px 20px;font-size:11px;color:#8a9a96;"">Surat elektronik ini dikirim secara otomatis oleh sistem FB-Plants setelah PO memperoleh persetujuan akhir. Mohon tidak mengubah lampiran dokumen PO.</td></tr>"
	h = h & "</table></td></tr></table></body></html>"
	PoMail_BuildHtml = h
End Function

Function PoMail_HtmlRow(label, valueHtml)
	PoMail_HtmlRow = "<tr><td style=""padding:7px 12px;width:130px;font-size:12px;color:#5d6d69;border-bottom:1px solid #eef2f1;"">" & label & "</td><td style=""padding:7px 12px;font-size:13px;border-bottom:1px solid #eef2f1;"">" & valueHtml & "</td></tr>"
End Function

' ---------------- kirim (CDO, SMTP Gmail) ----------------
' Mengembalikan "" bila berhasil, atau pesan error.
Function PoMail_Send(toList, ccList, subject, html, attachPath)
	Dim m, sch
	sch = "http://schemas.microsoft.com/cdo/configuration/"
	On Error Resume Next
	Set m = CreateObject("CDO.Message")
	With m.Configuration.Fields
		.Item(sch & "sendusing") = 2
		.Item(sch & "smtpserver") = PoMailCfg_SmtpServer
		.Item(sch & "smtpserverport") = PoMailCfg_SmtpPort
		.Item(sch & "smtpusessl") = PoMailCfg_SmtpSsl
		.Item(sch & "smtpauthenticate") = 1
		.Item(sch & "sendusername") = PoMailCfg_SmtpUser
		.Item(sch & "sendpassword") = PoMailCfg_SmtpPassword
		.Item(sch & "smtpconnectiontimeout") = 30
		.Update
	End With
	m.From = """" & PoMailCfg_FromName & """ <" & PoMailCfg_FromAddress & ">"
	If PoMailCfg_ReplyTo <> "" Then m.ReplyTo = PoMailCfg_ReplyTo
	m.To = toList
	If ccList <> "" Then m.Cc = ccList
	m.Subject = subject
	m.HTMLBody = html
	m.HTMLBodyPart.Charset = "utf-8"
	If attachPath <> "" Then m.AddAttachment attachPath
	If Err.Number <> 0 Then
		PoMail_Send = "Gagal menyiapkan email: " & Err.Description
		Err.Clear : Set m = Nothing : Exit Function
	End If
	m.Send
	If Err.Number <> 0 Then
		PoMail_Send = "Gagal mengirim (SMTP): " & Err.Description
		Err.Clear
	Else
		PoMail_Send = ""
	End If
	Set m = Nothing
End Function

' ---------------- data PO dari database ----------------
Function PoMail_Field(rs, name)
	Dim v
	On Error Resume Next
	v = rs(name)
	If Err.Number <> 0 Then
		Err.Clear : PoMail_Field = ""
	ElseIf IsNull(v) Then
		PoMail_Field = ""
	Else
		PoMail_Field = Trim("" & v)
	End If
End Function

' Isi hdr (Dictionary) & items (array Dictionary). Mengembalikan False bila PO tidak ditemukan.
Function PoMail_LoadPo(fcba, pocode, ByRef hdr, ByRef items)
	Dim rs, q, it, n, a()
	PoMail_LoadPo = False
	Set hdr = CreateObject("Scripting.Dictionary")
	q = " SELECT po.*, sp.fcname spname, sp.mail spmail, sp.address spaddress, sp.city spcity, sp.phone spphone, sp.fax spfax," & _
	    "        SUPPLIER_API.GET_BANKACCOUNT(po.suppliercode, po.fcba) bankaccount," & _
	    "        SUPPLIER_API.GET_BANKNAME(po.suppliercode, po.fcba) bankname," & _
	    "        SUPPLIER_API.GET_BANKOWNERSHIP(po.suppliercode, po.fcba) bankownership," & _
	    "        (SELECT fcname FROM document_classification WHERE document_type = 'PO' AND fccode = po.document_classification) classname," & _
	    "        (SELECT MIN(d.prcode) FROM purchaseorderdetails d WHERE d.pocode = po.pocode AND d.fcba = po.fcba) firstprcode," & _
	    "        bu.fccompanyname cmpname, bu.fccompanycode cmpcode, bu.fcname baname," & _
	    "        bu2.fcname pbname, bu2.address pbaddress, bu2.fctelephone pbphone, bu2.fcfax pbfax," & _
	    "        bu3.fcname destname, bu3.address destaddress" & _
	    "   FROM purchaseorder po" & _
	    "   LEFT JOIN supplier sp ON (po.suppliercode = sp.fccode AND po.fcba = sp.fcba)" & _
	    "   LEFT JOIN businessunit bu ON (bu.fccode = po.fcba)" & _
	    "   LEFT JOIN businessunit bu2 ON (bu2.fccode = po.purchaseby)" & _
	    "   LEFT JOIN businessunit bu3 ON (bu3.fccode = po.deliverydestination)" & _
	    "  WHERE po.fcba = '" & PoMail_Sql2(fcba) & "' AND po.pocode = '" & PoMail_Sql2(pocode) & "'"
	Set rs = CreateObject("ADODB.Recordset")
	rs.Open q, connStr
	If rs.EOF Then rs.Close : Set rs = Nothing : Exit Function
	hdr("fcba") = fcba
	hdr("pocode") = PoMail_Field(rs, "pocode")
	hdr("podate") = PoMail_Field(rs, "podate")
	hdr("suppliercode") = PoMail_Field(rs, "suppliercode")
	hdr("suppliername") = PoMail_Field(rs, "spname")
	hdr("supplieremail") = PoMail_Field(rs, "spmail")
	hdr("supplieraddress") = PoMail_Field(rs, "supplieraddress")
	If hdr("supplieraddress") = "" Then hdr("supplieraddress") = PoMail_Field(rs, "spaddress")
	hdr("suppliercity") = PoMail_Field(rs, "spcity")
	hdr("supplierphone") = PoMail_Field(rs, "spphone")
	hdr("supplierfax") = PoMail_Field(rs, "spfax")
	hdr("bankaccount") = PoMail_Field(rs, "bankaccount")
	hdr("bankname") = PoMail_Field(rs, "bankname")
	hdr("bankownership") = PoMail_Field(rs, "bankownership")
	hdr("classification") = PoMail_Field(rs, "classname")
	hdr("prcode") = PoMail_Field(rs, "firstprcode")
	hdr("companycode") = PoMail_Field(rs, "cmpcode")
	hdr("pbaddress") = PoMail_Field(rs, "pbaddress")
	hdr("pbphone") = PoMail_Field(rs, "pbphone")
	hdr("pbfax") = PoMail_Field(rs, "pbfax")
	hdr("totalamountppn") = PoMail_Field(rs, "totalamountppn")
	hdr("companyname") = PoMail_Field(rs, "cmpname")
	hdr("baname") = PoMail_Field(rs, "baname")
	hdr("purchasebyname") = PoMail_Field(rs, "pbname")
	hdr("destname") = PoMail_Field(rs, "destname")
	hdr("destaddress") = PoMail_Field(rs, "destaddress")
	hdr("remarks") = PoMail_Field(rs, "remarks")
	hdr("currency") = PoMail_Field(rs, "currencycode")
	hdr("termofpayment") = PoMail_Field(rs, "termofpayment")
	hdr("totalamount") = PoMail_Field(rs, "totalamount")
	hdr("ppn") = PoMail_Field(rs, "ppn")
	hdr("totalppn") = PoMail_Field(rs, "totalppn")
	hdr("totalcharge") = PoMail_Field(rs, "totalcharge")
	hdr("grandtotal") = PoMail_Field(rs, "grandtotal")
	hdr("dp") = PoMail_Field(rs, "dp")
	rs.Close
	If hdr("suppliername") = "" Then hdr("suppliername") = hdr("suppliercode")
	If hdr("companyname") = "" Then hdr("companyname") = hdr("baname")
	hdr("logopath") = PoMail_LogoPath(hdr("companycode"))

	' item: query sama dengan printingpo.asp
	q = " SELECT pod.itemcode, stock_api.get_stockName(pod.itemcode, pod.fcba) stockname, pod.quantity, st.unitofmeasurement uom," & _
	    "        pod.unitprice, pod.disc," & _
	    "        (pod.quantity * pod.unitprice) - ((pod.quantity * pod.unitprice) * (NVL(pod.disc, 0)) / 100) amount," & _
	    "        TRIM(REPLACE(REPLACE(REPLACE(pod.otheritemdesc, CHR(10), ''), CHR(13), ''), CHR(9), '')) otheritemdesc," & _
	    "        pod.otcstock1, pod.otcstock3, pod.othercharge, pod.otherchargestock" & _
	    "   FROM purchaseorderdetails pod" & _
	    "  INNER JOIN bankstock st ON (pod.itemcode = st.fccode)" & _
	    "  WHERE pod.fcba = '" & PoMail_Sql2(fcba) & "' AND pod.pocode = '" & PoMail_Sql2(pocode) & "'" & _
	    "  ORDER BY pod.linenokey"
	rs.Open q, connStr
	n = 0
	ReDim a(-1)
	Do While Not rs.EOF
		Set it = CreateObject("Scripting.Dictionary")
		it("itemcode") = PoMail_Field(rs, "itemcode")
		it("itemname") = PoMail_Field(rs, "stockname")
		it("desc") = PoMail_Field(rs, "otheritemdesc")
		it("qty") = PoMail_Field(rs, "quantity")
		it("uom") = PoMail_Field(rs, "uom")
		it("unitprice") = PoMail_Field(rs, "unitprice")
		it("disc") = PoMail_Field(rs, "disc")
		it("amount") = PoMail_Field(rs, "amount")
		it("otcstock1") = PoMail_Field(rs, "otcstock1")
		it("otcstock3") = PoMail_Field(rs, "otcstock3")
		it("othercharge") = PoMail_Field(rs, "othercharge")
		it("otherchargestock") = PoMail_Field(rs, "otherchargestock")
		ReDim Preserve a(n)
		Set a(n) = it
		n = n + 1
		rs.MoveNext
	Loop
	rs.Close : Set rs = Nothing
	items = a
	PoMail_LoadPo = True
End Function

' Logo perusahaan sama dengan printingpo.asp (css/login/images); PNG tidak didukung -> tanpa logo
Function PoMail_LogoPath(companyCode)
	Dim f, p
	Select Case UCase(Trim("" & companyCode))
		Case "PT.ARJ" : f = "logoarj.jpeg"
		Case "PT.BRS" : f = "logobrs.jpg"
		Case "PT.BIA" : f = "logobia.jpeg"
		Case "PT.KAA", "POM" : f = "logokaa.jpg"
		Case "PT.GAS" : f = "logogas.jpeg"
		Case "PT.KSR" : f = "logoksr.jpg"
		Case "PT.BRP" : f = "logobrp.jpg"
		Case "PT.CAE" : f = "logocae.jpg"
		Case Else : f = ""
	End Select
	PoMail_LogoPath = ""
	If f = "" Then Exit Function
	On Error Resume Next
	p = Server.MapPath("/css/login/images/" & f)
	If Err.Number = 0 Then
		If CreateObject("Scripting.FileSystemObject").FileExists(p) Then PoMail_LogoPath = p
	End If
	Err.Clear
End Function

Function PoMail_Sql2(s)
	PoMail_Sql2 = Replace("" & s, "'", "''")
End Function

Function PoMail_ValidEmails(s)
	' rapikan daftar email (pisah ; atau ,), buang yang tidak valid
	Dim parts, i, p, out
	out = ""
	parts = Split(Replace("" & s, ",", ";"), ";")
	For i = 0 To UBound(parts)
		p = Trim(parts(i))
		If InStr(p, "@") > 1 And InStr(p, " ") = 0 Then
			If InStr(InStr(p, "@"), p, ".") > 0 Then
				If out <> "" Then out = out & ";"
				out = out & p
			End If
		End If
	Next
	PoMail_ValidEmails = out
End Function

' ---------------- pemroses antrian ----------------
Function PoMail_ProcessOutbox(maxRows)
	Dim rs, q, ids(), fcbas(), pos(), atts(), n, i, affected, hdr, items, realTo, realCc, toList, ccList
	Dim folder, fso, pdfPath, pdfName, errMsg, isTest, linkUrl, report, sent, failed
	report = "" : sent = 0 : failed = 0
	If PoMailCfg_SmtpPassword = "" Then
		PoMail_ProcessOutbox = "Dilewati: App Password Gmail (PoMailCfg_SmtpPassword) belum diisi di email/po_mail_config.inc."
		Exit Function
	End If
	isTest = PoMailCfg_TestMode

	' baris SENDING yang macet (> 15 menit) dikembalikan ke PENDING
	connStr.Execute "UPDATE PO_EMAIL_OUTBOX SET STATUS = 'PENDING', UPDATED_AT = SYSDATE WHERE STATUS = 'SENDING' AND UPDATED_AT < SYSDATE - 15/1440"

	q = "SELECT * FROM (SELECT ID, FCBA, POCODE, ATTEMPTS FROM PO_EMAIL_OUTBOX" & _
	    " WHERE STATUS = 'PENDING' AND ATTEMPTS < " & CLng(PoMailCfg_MaxAttempts) & _
	    " AND (NEXT_TRY IS NULL OR NEXT_TRY <= SYSDATE) ORDER BY ID) WHERE ROWNUM <= " & CLng(maxRows)
	Set rs = CreateObject("ADODB.Recordset")
	rs.Open q, connStr
	n = 0
	ReDim ids(-1), fcbas(-1), pos(-1), atts(-1)
	Do While Not rs.EOF
		ReDim Preserve ids(n), fcbas(n), pos(n), atts(n)
		ids(n) = CLng(rs("ID")) : fcbas(n) = "" & rs("FCBA") : pos(n) = "" & rs("POCODE") : atts(n) = CLng(rs("ATTEMPTS"))
		n = n + 1
		rs.MoveNext
	Loop
	rs.Close : Set rs = Nothing
	If n = 0 Then PoMail_ProcessOutbox = "Tidak ada email PO yang menunggu." : Exit Function

	Set fso = CreateObject("Scripting.FileSystemObject")
	folder = Server.MapPath(PoMailCfg_PdfFolder)
	If Not fso.FolderExists(folder) Then
		On Error Resume Next
		fso.CreateFolder folder
		On Error GoTo 0
	End If

	For i = 0 To n - 1
		' klaim baris (cegah dobel kirim bila pemroses berjalan bersamaan)
		affected = 0
		connStr.Execute "UPDATE PO_EMAIL_OUTBOX SET STATUS = 'SENDING', UPDATED_AT = SYSDATE WHERE ID = " & ids(i) & " AND STATUS = 'PENDING'", affected
		If affected = 1 Then
			errMsg = ""
			On Error Resume Next
			If Not PoMail_LoadPo(fcbas(i), pos(i), hdr, items) Then
				If Err.Number <> 0 Then errMsg = "Gagal membaca PO: " & Err.Description Else errMsg = "PO tidak ditemukan"
			End If
			Err.Clear
			On Error GoTo 0

			If errMsg = "" Then
				realTo = PoMail_ValidEmails(hdr("supplieremail"))
				realCc = PoMail_ValidEmails(PoMailCfg_Cc)
				If isTest Then
					toList = PoMailCfg_TestTo
					If PoMailCfg_TestUseCc Then ccList = realCc Else ccList = ""
				Else
					toList = realTo
					ccList = realCc
					If toList = "" Then errMsg = "Email supplier (SUPPLIER.MAIL) kosong / tidak valid untuk supplier " & hdr("suppliercode")
				End If
			End If

			If errMsg = "" Then
				pdfName = "PO_" & PoMail_SafeName(pos(i)) & "_" & ids(i) & ".pdf"
				pdfPath = folder & "\" & pdfName
				On Error Resume Next
				Call PoMail_BuildPdf(hdr, items, pdfPath)
				If Err.Number <> 0 Then errMsg = "Gagal membuat PDF: " & Err.Description
				Err.Clear
				On Error GoTo 0
			End If

			If errMsg = "" Then
				linkUrl = ""
				If PoMailCfg_LinkBaseUrl <> "" Then linkUrl = PoMailCfg_LinkBaseUrl & "/printingpo.asp?pocode=" & Server.URLEncode(pos(i)) & "&fcba=" & Server.URLEncode(fcbas(i)) & "&purchaseby=&deliverydestination="
				' link bertanda tangan: supplier bisa membuka PO ini tanpa login (lihat po_link_token.inc)
				If linkUrl <> "" Then linkUrl = linkUrl & "&t=" & PoLink_Token(pos(i), fcbas(i), PoMailCfg_ProcessKey)
				errMsg = PoMail_Send(toList, ccList, PoMail_Subject(hdr, isTest), PoMail_BuildHtml(hdr, items, isTest, realTo, IIfStr(isTest And Not PoMailCfg_TestUseCc, realCc, ""), linkUrl), pdfPath)
			End If

			If errMsg = "" Then
				connStr.Execute "UPDATE PO_EMAIL_OUTBOX SET STATUS = 'SENT', SENT_AT = SYSDATE, UPDATED_AT = SYSDATE, LAST_ERROR = NULL," & _
				                " MAIL_TO = '" & PoMail_Sql2(toList) & "', MAIL_CC = '" & PoMail_Sql2(ccList) & "', PDF_FILE = '" & PoMail_Sql2(pdfName) & "'," & _
				                " IS_TEST = '" & IIfStr(isTest, "Y", "N") & "' WHERE ID = " & ids(i)
				sent = sent + 1
				report = report & "TERKIRIM  #" & ids(i) & " " & pos(i) & " -> " & toList & vbCrLf
			Else
				atts(i) = atts(i) + 1
				connStr.Execute "UPDATE PO_EMAIL_OUTBOX SET ATTEMPTS = " & atts(i) & ", UPDATED_AT = SYSDATE," & _
				                " STATUS = '" & IIfStr(atts(i) >= CLng(PoMailCfg_MaxAttempts), "FAILED", "PENDING") & "'," & _
				                " NEXT_TRY = SYSDATE + " & (atts(i) * 5) & "/1440," & _
				                " LAST_ERROR = '" & PoMail_Sql2(Left(errMsg, 1900)) & "' WHERE ID = " & ids(i)
				failed = failed + 1
				report = report & "GAGAL     #" & ids(i) & " " & pos(i) & " (percobaan " & atts(i) & "): " & errMsg & vbCrLf
			End If
		End If
	Next
	PoMail_ProcessOutbox = "Terkirim: " & sent & ", gagal: " & failed & vbCrLf & report
End Function

Function PoMail_SafeName(s)
	Dim re
	Set re = New RegExp
	re.Global = True
	re.Pattern = "[^A-Za-z0-9_-]+"
	PoMail_SafeName = Left(re.Replace("" & s, "-"), 80)
End Function
%>
