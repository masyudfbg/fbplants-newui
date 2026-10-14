<%
'=====================================================================
' pdfwriter.inc - pembuat PDF sederhana (VBScript murni, tanpa komponen luar)
' Halaman A4, font standar Helvetica / Helvetica-Bold / Helvetica-BoldOblique (WinAnsi), teks,
' kotak, garis, gambar JPEG (logo) dan QR code (kotak-kotak vektor, tidak buram saat dicetak).
' Koordinat memakai titik (1/72 inci) dihitung dari KIRI-ATAS halaman.
'
'   Dim pdf : Set pdf = New PoPdf
'   pdf.NewPage
'   pdf.Text 40, 60, 16, True, "PURCHASE ORDER"
'   pdf.SaveTo Server.MapPath("/email/po_pdf/x.pdf")
'=====================================================================

Class PoPdf
	Public PageW, PageH
	Private pages(), nPages, cur, curIdx, docTitle, imgData(), imgW(), imgH(), imgCs(), nImg

	Private Sub Class_Initialize()
		PageW = 595.28 : PageH = 841.89
		nPages = 0 : cur = "" : curIdx = -1 : docTitle = ""
		ReDim pages(0)
		nImg = 0 : ReDim imgData(0), imgW(0), imgH(0), imgCs(0)
	End Sub

	Public Property Let Title(v) : docTitle = v : End Property

	Public Sub NewPage()
		If curIdx >= 0 Then pages(curIdx) = cur
		nPages = nPages + 1
		ReDim Preserve pages(nPages - 1)
		curIdx = nPages - 1
		cur = ""
	End Sub

	' Pindah ke halaman tertentu (0 = pertama), mis. untuk menulis "Halaman x dari n"
	Public Sub SelectPage(i)
		If curIdx >= 0 Then pages(curIdx) = cur
		curIdx = i
		cur = pages(i)
	End Sub

	Public Property Get PageCount() : PageCount = nPages : End Property

	' ---------- angka & teks ----------
	Private Function N(v)
		Dim s
		s = CStr(Round(CDbl(v), 2))
		N = Replace(s, ",", ".")
	End Function

	Private Function Esc(s)
		Dim i, c, o, out
		out = ""
		s = "" & s
		For i = 1 To Len(s)
			c = Mid(s, i, 1)
			o = AscW(c)
			If o < 0 Then o = o + 65536
			If c = "\" Or c = "(" Or c = ")" Then
				out = out & "\" & c
			ElseIf o = 9 Then
				out = out & " "
			ElseIf o < 32 Or (o > 126 And o < 160) Or o > 255 Then
				out = out & "?"
			Else
				out = out & c
			End If
		Next
		Esc = out
	End Function

	Private Function CharW(c, bold)
		Dim o, reg, bld
		o = AscW(c)
		If o < 32 Or o > 126 Then CharW = 556 : Exit Function
		' lebar AFM Helvetica / Helvetica-Bold karakter 32..126
		reg = Array(278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584)
		bld = Array(278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,389,280,389,584)
		If bold = 2 Or bold = True Then CharW = bld(o - 32) Else CharW = reg(o - 32)
	End Function

	Public Function TextWidth(s, size, bold)
		Dim i, w
		w = 0
		s = "" & s
		For i = 1 To Len(s) : w = w + CharW(Mid(s, i, 1), bold) : Next
		TextWidth = w * size / 1000
	End Function

	' Potong teks agar muat di lebar tertentu (diakhiri "...")
	Public Function Fit(s, maxW, size, bold)
		s = "" & s
		If TextWidth(s, size, bold) <= maxW Then Fit = s : Exit Function
		Do While Len(s) > 0 And TextWidth(s & "...", size, bold) > maxW
			s = Left(s, Len(s) - 1)
		Loop
		Fit = s & "..."
	End Function

	' Pecah teks menjadi beberapa baris sesuai lebar (dipisah per kata)
	Public Function Wrap(s, maxW, size, bold)
		Dim words, i, line, test, outArr(), n
		n = 0 : ReDim outArr(0)
		words = Split(Replace(Replace("" & s, vbCr, " "), vbLf, " "), " ")
		line = ""
		For i = 0 To UBound(words)
			If words(i) <> "" Then
				If line = "" Then test = words(i) Else test = line & " " & words(i)
				If TextWidth(test, size, bold) <= maxW Or line = "" Then
					line = test
				Else
					ReDim Preserve outArr(n) : outArr(n) = line : n = n + 1
					line = words(i)
				End If
			End If
		Next
		If line <> "" Or n = 0 Then ReDim Preserve outArr(n) : outArr(n) = line : n = n + 1
		Wrap = outArr
	End Function

	' ---------- gambar ----------
	Public Sub SetFill(r, g, b)
		cur = cur & N(r / 255) & " " & N(g / 255) & " " & N(b / 255) & " rg" & vbLf
	End Sub

	Public Sub SetStroke(r, g, b)
		cur = cur & N(r / 255) & " " & N(g / 255) & " " & N(b / 255) & " RG" & vbLf
	End Sub

	Public Sub SetLineWidth(w)
		cur = cur & N(w) & " w" & vbLf
	End Sub

	' mode: "F" isi, "S" garis tepi, "B" keduanya
	Public Sub Rect(x, yTop, w, h, mode)
		cur = cur & N(x) & " " & N(PageH - yTop - h) & " " & N(w) & " " & N(h) & " re " & mode & vbLf
	End Sub

	Public Sub Line(x1, y1, x2, y2)
		cur = cur & N(x1) & " " & N(PageH - y1) & " m " & N(x2) & " " & N(PageH - y2) & " l S" & vbLf
	End Sub

	' y = garis dasar teks, diukur dari atas halaman; bold: False, True, atau 2 (tebal miring)
	Public Sub Text(x, y, size, bold, s)
		Dim f
		If bold = 2 Then
			f = "/F3"
		ElseIf bold Then
			f = "/F2"
		Else
			f = "/F1"
		End If
		cur = cur & "BT " & f & " " & N(size) & " Tf " & N(x) & " " & N(PageH - y) & " Td (" & Esc(s) & ") Tj ET" & vbLf
	End Sub

	Public Sub TextRight(xRight, y, size, bold, s)
		Text xRight - TextWidth(s, size, bold), y, size, bold, s
	End Sub

	Public Sub TextCenter(xCenter, y, size, bold, s)
		Text xCenter - TextWidth(s, size, bold) / 2, y, size, bold, s
	End Sub

	' Muat gambar JPEG (baseline/progressive, RGB/abu/CMYK). Mengembalikan nama ("/Im1") atau "" bila gagal.
	Public Function AddJpeg(path)
		Dim st, d, i, m, segLen, w, h, comps
		AddJpeg = ""
		On Error Resume Next
		Set st = CreateObject("ADODB.Stream")
		st.Type = 2 : st.Charset = "windows-1252" : st.Open
		st.LoadFromFile path
		d = st.ReadText
		st.Close : Set st = Nothing
		If Err.Number <> 0 Then Err.Clear : Exit Function
		On Error GoTo 0
		If Len(d) < 4 Then Exit Function
		If Asc(Mid(d, 1, 1)) <> 255 Or Asc(Mid(d, 2, 1)) <> 216 Then Exit Function
		i = 3 : w = 0
		Do While i < Len(d) - 8
			If Asc(Mid(d, i, 1)) <> 255 Then
				i = i + 1
			Else
				m = Asc(Mid(d, i + 1, 1))
				If m = 192 Or m = 193 Or m = 194 Then
					h = Asc(Mid(d, i + 5, 1)) * 256 + Asc(Mid(d, i + 6, 1))
					w = Asc(Mid(d, i + 7, 1)) * 256 + Asc(Mid(d, i + 8, 1))
					comps = Asc(Mid(d, i + 9, 1))
					Exit Do
				ElseIf m = 216 Or m = 1 Or (m >= 208 And m <= 215) Or m = 255 Then
					i = i + 1
				Else
					segLen = Asc(Mid(d, i + 2, 1)) * 256 + Asc(Mid(d, i + 3, 1))
					i = i + 2 + segLen
				End If
			End If
		Loop
		If w = 0 Or h = 0 Then Exit Function
		nImg = nImg + 1
		ReDim Preserve imgData(nImg - 1), imgW(nImg - 1), imgH(nImg - 1), imgCs(nImg - 1)
		imgData(nImg - 1) = d : imgW(nImg - 1) = w : imgH(nImg - 1) = h
		Select Case comps
			Case 1 : imgCs(nImg - 1) = "/DeviceGray"
			Case 4 : imgCs(nImg - 1) = "/DeviceCMYK /Decode [1 0 1 0 1 0 1 0]"
			Case Else : imgCs(nImg - 1) = "/DeviceRGB"
		End Select
		AddJpeg = "/Im" & nImg
	End Function

	Public Function ImageWidth(name) : ImageWidth = imgW(CLng(Mid(name, 4)) - 1) : End Function
	Public Function ImageHeight(name) : ImageHeight = imgH(CLng(Mid(name, 4)) - 1) : End Function

	' Gambar yang sudah dimuat: sudut kiri-atas (x, yTop), ukuran w x h titik
	Public Sub Image(name, x, yTop, w, h)
		cur = cur & "q " & N(w) & " 0 0 " & N(h) & " " & N(x) & " " & N(PageH - yTop - h) & " cm " & name & " Do Q" & vbLf
	End Sub

	' QR code: mods(y*size+x)=True gelap; sudut kiri-atas (x, yTop), sisi satu modul = cell
	Public Sub QrCode(mods, size, x, yTop, cell)
		Dim yy, xx, runStart
		For yy = 0 To size - 1
			xx = 0
			Do While xx < size
				If mods(yy * size + xx) Then
					runStart = xx
					Do While xx < size
						If Not mods(yy * size + xx) Then Exit Do
						xx = xx + 1
					Loop
					cur = cur & N(x + runStart * cell) & " " & N(PageH - yTop - (yy + 1) * cell) & " " & N((xx - runStart) * cell) & " " & N(cell) & " re" & vbLf
				Else
					xx = xx + 1
				End If
			Loop
		Next
		cur = cur & "f" & vbLf
	End Sub

	' ---------- keluaran ----------
	Public Function Output()
		Dim objs(), nObj, i, s, offs(), body, xref, pageObj, contObj, kids, d, xobj
		If nPages = 0 Then NewPage
		pages(curIdx) = cur
		nObj = 6 + nPages * 2 + nImg
		ReDim objs(nObj), offs(nObj)
		kids = ""
		For i = 0 To nPages - 1
			pageObj = 7 + i * 2
			kids = kids & pageObj & " 0 R "
		Next
		d = Now()
		objs(1) = "<< /Type /Catalog /Pages 2 0 R >>"
		objs(2) = "<< /Type /Pages /Kids [ " & kids & "] /Count " & nPages & " >>"
		objs(3) = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
		objs(4) = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"
		objs(6) = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-BoldOblique /Encoding /WinAnsiEncoding >>"
		objs(5) = "<< /Producer (FB-Plants) /Title (" & Esc(docTitle) & ") /CreationDate (D:" & Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2) & ") >>"
		xobj = ""
		For i = 0 To nImg - 1
			objs(7 + nPages * 2 + i) = "<< /Type /XObject /Subtype /Image /Width " & imgW(i) & " /Height " & imgH(i) & " /ColorSpace " & imgCs(i) & " /BitsPerComponent 8 /Filter /DCTDecode /Length " & Len(imgData(i)) & " >>" & vbLf & "stream" & vbLf & imgData(i) & vbLf & "endstream"
			xobj = xobj & "/Im" & (i + 1) & " " & (7 + nPages * 2 + i) & " 0 R "
		Next
		If xobj <> "" Then xobj = " /XObject << " & xobj & ">>"
		For i = 0 To nPages - 1
			pageObj = 7 + i * 2 : contObj = pageObj + 1
			objs(pageObj) = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & N(PageW) & " " & N(PageH) & "] /Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 6 0 R >>" & xobj & " >> /Contents " & contObj & " 0 R >>"
			objs(contObj) = "<< /Length " & Len(pages(i)) & " >>" & vbLf & "stream" & vbLf & pages(i) & "endstream"
		Next
		body = "%PDF-1.4" & vbLf & "%" & Chr(226) & Chr(227) & Chr(207) & Chr(211) & vbLf
		For i = 1 To nObj
			offs(i) = Len(body)
			body = body & i & " 0 obj" & vbLf & objs(i) & vbLf & "endobj" & vbLf
		Next
		xref = Len(body)
		body = body & "xref" & vbLf & "0 " & (nObj + 1) & vbLf & "0000000000 65535 f" & vbCrLf
		For i = 1 To nObj
			body = body & Right("0000000000" & offs(i), 10) & " 00000 n" & vbCrLf
		Next
		body = body & "trailer" & vbLf & "<< /Size " & (nObj + 1) & " /Root 1 0 R /Info 5 0 R >>" & vbLf & "startxref" & vbLf & xref & vbLf & "%%EOF" & vbLf
		Output = body
	End Function

	' Simpan ke file (1 karakter = 1 byte). windows-1252 menyalin byte gambar JPEG persis
	' (iso-8859-1 di ADODB.Stream tidak); nilai byte dibaca dengan Asc(), bukan AscW().
	Public Sub SaveTo(path)
		Dim st
		Set st = CreateObject("ADODB.Stream")
		st.Type = 2
		st.Charset = "windows-1252"
		st.Open
		st.WriteText Output()
		st.SaveToFile path, 2
		st.Close
		Set st = Nothing
	End Sub
End Class
%>
