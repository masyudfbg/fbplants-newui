<%
' =====================================================================
' Token tanda tangan untuk link PO di email supplier (printingpo.asp).
' Link berisi &t=<HMAC-SHA256(kunci, POCODE|FCBA)>, sehingga supplier bisa membuka
' PO-nya tanpa login, tetapi tidak bisa menebak/membuka PO lain dengan mengganti pocode.
' Kunci = PoMailCfg_ProcessKey di email/po_mail_config.inc (tidak masuk git).
' =====================================================================
Function PoLink_Hex(bytes)
	Dim i, s
	s = ""
	For i = 1 To LenB(bytes)
		s = s & Right("0" & Hex(AscB(MidB(bytes, i, 1))), 2)
	Next
	PoLink_Hex = LCase(s)
End Function

Function PoLink_Token(pocode, fcba, key)
	Dim enc, h
	PoLink_Token = ""
	If Trim(key & "") = "" Then Exit Function
	Set enc = Server.CreateObject("System.Text.UTF8Encoding")
	Set h = Server.CreateObject("System.Security.Cryptography.HMACSHA256")
	h.Key = enc.GetBytes_4(CStr(key))
	PoLink_Token = PoLink_Hex(h.ComputeHash_2((enc.GetBytes_4(UCase(Trim(pocode & "")) & "|" & UCase(Trim(fcba & ""))))))
End Function

' Kunci dibaca dari email/po_mail_config.inc saat dibutuhkan (bukan lewat #include,
' agar halaman tidak error di server yang belum punya konfigurasi email).
Function PoLink_KeyFromConfig()
	Dim fso, path, ts, txt, re, m
	PoLink_KeyFromConfig = ""
	On Error Resume Next
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	path = Server.MapPath("/email/po_mail_config.inc")
	If Not fso.FileExists(path) Then Exit Function
	Set ts = fso.OpenTextFile(path, 1)
	txt = ts.ReadAll
	ts.Close
	Set re = New RegExp
	re.Pattern = "^\s*PoMailCfg_ProcessKey\s*=\s*""([^""]*)"""
	re.Multiline = True
	re.IgnoreCase = True
	Set m = re.Execute(txt)
	If m.Count > 0 Then PoLink_KeyFromConfig = m(0).SubMatches(0)
	On Error GoTo 0
End Function

' True bila token pada link cocok dengan PO & BA yang diminta
Function PoLink_Valid(pocode, fcba, token)
	Dim key, expect
	PoLink_Valid = False
	If Trim(token & "") = "" Or Trim(pocode & "") = "" Then Exit Function
	key = PoLink_KeyFromConfig()
	If key = "" Then Exit Function
	expect = PoLink_Token(pocode, fcba, key)
	PoLink_Valid = (expect <> "" And LCase(Trim(token & "")) = expect)
End Function
%>
