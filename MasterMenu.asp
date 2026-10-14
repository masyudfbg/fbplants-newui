<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="include/db_param.asp"-->
<%
' =====================================================================
' Master Menu: aktif/nonaktifkan menu (MENU_TAB.ISACTIVE) - SYS-SYSTEM.
' Hanya user yang punya akses ke menu ini (MENU_ROLES_TAB lewat USER_ROLES_API) yang
' bisa membuka & mengubah. Menu ini sendiri dan induknya tidak bisa dimatikan agar admin
' tidak terkunci. Perubahan berlaku di sidebar saat halaman dimuat ulang.
' =====================================================================
Const MM_SELF = "1450000"

Function MmJson(v)
	Dim s, i, c, code, o
	s = "" & v
	o = ""
	For i = 1 To Len(s)
		c = Mid(s, i, 1)
		code = AscW(c)
		If code < 0 Then code = code + 65536
		Select Case code
			Case 34: o = o & "\"""
			Case 92: o = o & "\\"
			Case 10: o = o & "\n"
			Case 13: o = o & "\r"
			Case 9: o = o & "\t"
			Case Else
				If code < 32 Or code > 126 Or code = 60 Or code = 62 Or code = 38 Then
					o = o & "\u" & Right("000" & Hex(code), 4)
				Else
					o = o & c
				End If
		End Select
	Next
	MmJson = """" & o & """"
End Function

Function MmCanAccess()
	Dim r
	MmCanAccess = False
	If Session("CodeName") & "" = "" Then Exit Function
	Set r = DbQuery("select USER_ROLES_API.GetAccessUsers(?, ?) a from dual", Array(Session("CodeName"), MM_SELF))
	If Not r.EOF Then MmCanAccess = (UCase(r("a") & "") = "TRUE")
	r.Close
End Function

' kode menu ini dan semua induknya (tidak boleh dimatikan)
Function MmLockedCodes()
	Dim d, code, r, guard
	Set d = Server.CreateObject("Scripting.Dictionary")
	code = MM_SELF
	guard = 0
	Do While code <> "" And guard < 10
		d(code) = True
		Set r = DbQuery("select menuparentcode from menu_tab where fccode = ?", Array(code))
		If r.EOF Then
			code = ""
		Else
			code = Trim(r("menuparentcode") & "")
			If d.Exists(code) Then code = ""
		End If
		r.Close
		guard = guard + 1
	Loop
	Set MmLockedCodes = d
End Function

Dim mmAllowed
mmAllowed = MmCanAccess()

If Session("mm_tok") & "" = "" Then
	Randomize
	Session("mm_tok") = Hex(Int(Rnd * 2147483647)) & Hex(Int(Rnd * 2147483647)) & Hex(Int(Timer * 1000))
End If

' ---------- AJAX: ubah status ----------
If Request.Form("action") = "toggle" Then
	Response.ContentType = "application/json"
	Response.CacheControl = "no-cache"
	Dim mmCode, mmActive, mmLocked, mmChk
	mmCode = Trim(Request.Form("code") & "")
	mmActive = UCase(Trim(Request.Form("active") & ""))
	If Not mmAllowed Then
		Response.Status = "403 Forbidden"
		Response.Write "{""ok"":false,""error"":""Anda tidak punya akses ke Master Menu.""}"
		Response.End
	End If
	If Request.Form("tok") <> Session("mm_tok") Then
		Response.Status = "403 Forbidden"
		Response.Write "{""ok"":false,""error"":""Token tidak valid, muat ulang halaman.""}"
		Response.End
	End If
	If mmCode = "" Or (mmActive <> "TRUE" And mmActive <> "FALSE") Then
		Response.Status = "400 Bad Request"
		Response.Write "{""ok"":false,""error"":""Permintaan tidak valid.""}"
		Response.End
	End If
	Set mmLocked = MmLockedCodes()
	If mmActive = "FALSE" And mmLocked.Exists(mmCode) Then
		Response.Status = "400 Bad Request"
		Response.Write "{""ok"":false,""error"":""Menu ini tidak bisa dinonaktifkan agar Master Menu tetap bisa dibuka.""}"
		Response.End
	End If
	Set mmChk = DbQuery("select count(*) n from menu_tab where fccode = ?", Array(mmCode))
	If CLng(mmChk("n")) = 0 Then
		Response.Status = "404 Not Found"
		Response.Write "{""ok"":false,""error"":""Menu tidak ditemukan.""}"
		Response.End
	End If
	mmChk.Close
	DbExec "update menu_tab set isactive = ?, fcedit = ?, fcip = ?, lastupdate = sysdate, lasttime = to_char(sysdate, 'HH24:MI') where fccode = ?", _
		Array(mmActive, Left(Session("CodeName") & "", 30), Left(Request.ServerVariables("REMOTE_ADDR") & "", 30), mmCode)
	connStr.Execute "commit"
	Response.Write "{""ok"":true,""code"":" & MmJson(mmCode) & ",""active"":" & LCase(CStr(mmActive = "TRUE")) & ",""by"":" & MmJson(Session("CodeName")) & "}"
	Response.End
End If

' ---------- AJAX: tambah / ubah menu ----------
Sub MmFail(status, msg)
	Response.Status = status
	Response.Write "{""ok"":false,""error"":" & MmJson(msg) & "}"
	Response.End
End Sub

Function MmCount(sql, params)
	Dim r
	Set r = DbQuery(sql, params)
	MmCount = CLng(r(0))
	r.Close
End Function

' satu baris menu dalam format array JSON yang sama dengan data tabel
Function MmRowJson(r)
	MmRowJson = "[" & MmJson(r("fccode")) & "," & MmJson(r("fcname")) & "," & MmJson(r("menuparentcode")) & "," & _
		MmJson(r("formlink")) & "," & MmJson(r("modname")) & "," & LCase(CStr(UCase(Trim(r("isactive") & "")) = "TRUE")) & "," & _
		MmJson(r("sequence")) & "," & MmJson(r("menutypecode")) & "," & MmJson(r("fcedit")) & "," & MmJson(r("lastupd")) & "," & _
		MmJson(r("helpdescription")) & "," & MmJson(r("remarks")) & "," & MmJson(r("module")) & "]"
End Function

Const MM_SELECT = "select m.fccode, m.fcname, m.menuparentcode, m.formlink, nvl(b.name, '') modname, m.isactive, m.sequence, m.menutypecode, m.fcedit, to_char(m.lastupdate, 'DD/MM/YYYY') lastupd, m.helpdescription, m.remarks, m.module from menu_tab m left join module b on m.module = b.code"

If Request.Form("action") = "save" Then
	Response.ContentType = "application/json"
	Response.CacheControl = "no-cache"
	Dim fMode, fCode, fName, fParent, fLink, fModule, fType, fSeq, fActive, fHelp, fRemarks, fRe, fP, fGuard, fRs, fUser, fIp
	If Not mmAllowed Then MmFail "403 Forbidden", "Anda tidak punya akses ke Master Menu."
	If Request.Form("tok") <> Session("mm_tok") Then MmFail "403 Forbidden", "Token tidak valid, muat ulang halaman."
	fMode = Request.Form("mode")
	fCode = Trim(Request.Form("code") & "")
	fName = Trim(Request.Form("name") & "")
	fParent = Trim(Request.Form("parent") & "")
	fLink = Trim(Request.Form("link") & "")
	fModule = Trim(Request.Form("module") & "")
	fType = Trim(Request.Form("type") & "")
	fSeq = Trim(Request.Form("seq") & "")
	fActive = UCase(Trim(Request.Form("active") & ""))
	fHelp = Trim(Request.Form("help") & "")
	fRemarks = Trim(Request.Form("remarks") & "")

	If fMode <> "insert" And fMode <> "update" Then MmFail "400 Bad Request", "Permintaan tidak valid."
	Set fRe = New RegExp
	fRe.Pattern = "^[A-Za-z0-9_.\-]{1,30}$"
	If Not fRe.Test(fCode) Then MmFail "400 Bad Request", "Kode menu wajib diisi (maks. 30 karakter: huruf, angka, titik, garis bawah, atau tanda minus)."
	If fName = "" Or Len(fName) > 100 Then MmFail "400 Bad Request", "Nama menu wajib diisi (maks. 100 karakter)."
	If Len(fLink) > 100 Then MmFail "400 Bad Request", "Link maksimal 100 karakter."
	If Len(fHelp) > 100 Then MmFail "400 Bad Request", "Keterangan bantuan maksimal 100 karakter."
	If Len(fRemarks) > 1000 Then MmFail "400 Bad Request", "Catatan maksimal 1000 karakter."
	If fActive <> "TRUE" And fActive <> "FALSE" Then MmFail "400 Bad Request", "Status tidak valid."
	If fSeq <> "" Then
		If Not IsNumeric(fSeq) Then MmFail "400 Bad Request", "Urutan harus berupa angka."
	End If
	If Not IsNumeric(fModule) Then MmFail "400 Bad Request", "Pilih modul."
	If MmCount("select count(*) from module where code = to_number(?)", Array(fModule)) = 0 Then MmFail "400 Bad Request", "Modul tidak ditemukan."
	If MmCount("select count(*) from menu_type_tab where fccode = ?", Array(fType)) = 0 Then MmFail "400 Bad Request", "Pilih tipe menu."

	' induk: "0" = menu utama; selain itu harus ada, bukan dirinya sendiri, dan bukan turunannya (cegah putaran)
	If fParent = "" Then MmFail "400 Bad Request", "Pilih menu induk (atau 0 untuk menu utama)."
	If fParent <> "0" Then
		If fParent = fCode Then MmFail "400 Bad Request", "Menu induk tidak boleh menu itu sendiri."
		If MmCount("select count(*) from menu_tab where fccode = ?", Array(fParent)) = 0 Then MmFail "400 Bad Request", "Menu induk " & fParent & " tidak ditemukan."
		fP = fParent : fGuard = 0
		Do While fP <> "" And fP <> "0" And fGuard < 20
			Set fRs = DbQuery("select menuparentcode from menu_tab where fccode = ?", Array(fP))
			If fRs.EOF Then fP = "" Else fP = Trim(fRs("menuparentcode") & "")
			fRs.Close
			If fP = fCode Then MmFail "400 Bad Request", "Menu induk tidak boleh salah satu submenu dari menu ini."
			fGuard = fGuard + 1
		Loop
	End If

	fUser = Left(Session("CodeName") & "", 30)
	fIp = Left(Request.ServerVariables("REMOTE_ADDR") & "", 30)
	If fMode = "insert" Then
		If MmCount("select count(*) from menu_tab where fccode = ?", Array(fCode)) > 0 Then MmFail "409 Conflict", "Kode menu " & fCode & " sudah dipakai."
		DbExec "insert into menu_tab (fccode, fcname, menuparentcode, formlink, module, isactive, helpdescription, sequence, remarks, " & _
		       "fcentry, fcedit, fcip, lastupdate, lasttime, menutypecode) values (?, ?, ?, ?, to_number(?), ?, ?, to_number(?), ?, ?, ?, ?, sysdate, to_char(sysdate, 'HH24:MI'), ?)", _
		       Array(fCode, fName, fParent, fLink, fModule, fActive, fHelp, fSeq, fRemarks, fUser, fUser, fIp, fType)
	Else
		If MmCount("select count(*) from menu_tab where fccode = ?", Array(fCode)) = 0 Then MmFail "404 Not Found", "Menu tidak ditemukan."
		If fActive = "FALSE" And MmLockedCodes().Exists(fCode) Then MmFail "400 Bad Request", "Menu ini tidak bisa dinonaktifkan agar Master Menu tetap bisa dibuka."
		DbExec "update menu_tab set fcname = ?, menuparentcode = ?, formlink = ?, module = to_number(?), isactive = ?, helpdescription = ?, " & _
		       "sequence = to_number(?), remarks = ?, menutypecode = ?, fcedit = ?, fcip = ?, lastupdate = sysdate, lasttime = to_char(sysdate, 'HH24:MI') where fccode = ?", _
		       Array(fName, fParent, fLink, fModule, fActive, fHelp, fSeq, fRemarks, fType, fUser, fIp, fCode)
	End If
	connStr.Execute "commit"
	Set fRs = DbQuery(MM_SELECT & " where m.fccode = ?", Array(fCode))
	Response.Write "{""ok"":true,""row"":" & MmRowJson(fRs) & "}"
	fRs.Close
	Response.End
End If

' ---------- data untuk halaman ----------
Dim mmRows, mmRs, mmFirst, mmLockList, mmK, mmModules, mmTypes
mmRows = "" : mmModules = "" : mmTypes = ""
If mmAllowed Then
	Set mmRs = DbQuery(MM_SELECT & " order by m.menutypecode, m.menuparentcode, m.sequence, m.fccode", Array())
	mmFirst = True
	Do While Not mmRs.EOF
		If Not mmFirst Then mmRows = mmRows & ","
		mmFirst = False
		mmRows = mmRows & MmRowJson(mmRs)
		mmRs.MoveNext
	Loop
	mmRs.Close
	Set mmLockList = MmLockedCodes()
	' pilihan modul & tipe untuk form tambah/ubah
	Set mmRs = DbQuery("select code, name from module order by name", Array())
	Do While Not mmRs.EOF
		If mmModules <> "" Then mmModules = mmModules & ","
		mmModules = mmModules & "[" & MmJson(mmRs("code")) & "," & MmJson(mmRs("name")) & "]"
		mmRs.MoveNext
	Loop
	mmRs.Close
	Set mmRs = DbQuery("select fccode, fcname from menu_type_tab order by fccode", Array())
	Do While Not mmRs.EOF
		If mmTypes <> "" Then mmTypes = mmTypes & ","
		mmTypes = mmTypes & "[" & MmJson(mmRs("fccode")) & "," & MmJson(mmRs("fcname")) & "]"
		mmRs.MoveNext
	Loop
	mmRs.Close
End If
%>
<html>
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>ePlaS - Master Menu</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<script type="text/javascript" src="dtree.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
<style>
	.mm { --brand: #1B7863; --brand-dark: #12584a; --tint: #eef6f3; --line: #e1e7e5; --muted: #5d6d69; --ink: #17262a; color: var(--ink); }
	.mm-head { display: flex; flex-wrap: wrap; align-items: flex-end; justify-content: space-between; gap: 12px 20px; margin: 0 0 14px; }
	.mm-head h2 { margin: 0; font-size: 20px; font-weight: 700; }
	.mm-head p { margin: 4px 0 0; color: var(--muted); font-size: 13px; }
	.mm-kpis { display: flex; gap: 10px; flex-wrap: wrap; }
	.mm-kpi { min-width: 110px; padding: 10px 14px; background: #fff; border: 1px solid var(--line); border-radius: 10px; }
	.mm-kpi span { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .3px; color: var(--muted); }
	.mm-kpi b { display: block; margin-top: 2px; font-size: 18px; font-variant-numeric: tabular-nums; }
	.mm-kpi.on b { color: #136a4b; } .mm-kpi.off b { color: #a1342a; }
	.mm-card { background: #fff; border: 1px solid var(--line); border-radius: 12px; overflow: hidden; }
	.mm-tools { display: flex; flex-wrap: wrap; align-items: center; gap: 10px 12px; padding: 12px 14px; border-bottom: 1px solid var(--line); }
	.mm-search { position: relative; flex: 1 1 260px; }
	.mm-search input { width: 100%; height: 38px; padding: 0 12px 0 36px; font: inherit; font-size: 13px; border: 1px solid #d5dfdc; border-radius: 9px; background: #f6f9f8; outline: none; box-sizing: border-box; }
	.mm-search input:focus { background: #fff; border-color: var(--brand); box-shadow: 0 0 0 3px rgba(27,120,99,.18); }
	.mm-search .fa { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: #8a9a96; }
	.mm-seg { display: inline-flex; padding: 3px; background: #eef2f1; border-radius: 10px; }
	.mm-seg button { height: 32px; padding: 0 12px; font: 600 12.5px inherit; color: var(--muted); background: transparent; border: 0; border-radius: 8px; cursor: pointer; }
	.mm-seg button[aria-pressed="true"] { color: var(--brand-dark); background: #fff; box-shadow: 0 1px 3px rgba(23,38,42,.12); }
	.mm-tools select { height: 38px; padding: 0 10px; font: inherit; font-size: 13px; border: 1px solid #d5dfdc; border-radius: 9px; background: #fff; }
	.mm-scroll { overflow: auto; max-height: max(360px, calc(100vh - 360px)); }
	.mm table { width: 100%; border-collapse: separate; border-spacing: 0; }
	.mm thead th { position: sticky; top: 0; z-index: 1; padding: 10px 12px; font-size: 11px; font-weight: 700; letter-spacing: .3px; text-transform: uppercase; text-align: left; white-space: nowrap; color: var(--muted); background: #f6f9f8; border-bottom: 1px solid var(--line); cursor: pointer; user-select: none; }
	.mm thead th.ns { cursor: default; }
	.mm thead th .ar { margin-left: 4px; opacity: .45; }
	.mm thead th[aria-sort] .ar { opacity: 1; color: var(--brand); }
	.mm tbody td { padding: 9px 12px; font-size: 13px; border-bottom: 1px solid #edf1f0; vertical-align: middle; }
	.mm tbody tr:hover td { background: #f6faf8; }
	.mm tbody tr.off td { color: #7d8b88; }
	.mm .code { font: 600 12px ui-monospace, Consolas, monospace; color: var(--brand-dark); background: var(--tint); padding: 2px 7px; border-radius: 6px; white-space: nowrap; }
	.mm .path { display: block; margin-top: 2px; font-size: 11.5px; color: var(--muted); }
	.mm .type { display: inline-block; padding: 2px 8px; font-size: 11px; font-weight: 600; border-radius: 999px; background: #eef1f0; color: #4a5a56; white-space: nowrap; }
	.mm .type.FORM { background: #e8f1fb; color: #1f5a96; }
	.mm .type.REPORT { background: #f3ecfb; color: #6a3d9a; }
	.mm .link { font: 12px ui-monospace, Consolas, monospace; color: #3d4d49; }
	.mm .muted { color: #9aa8a4; }
	.mm mark { background: #ffe58a; color: inherit; border-radius: 3px; padding: 0 1px; }
	/* switch */
	.sw { position: relative; display: inline-flex; align-items: center; gap: 8px; cursor: pointer; }
	.sw input { position: absolute; opacity: 0; width: 1px; height: 1px; }
	.sw .track { position: relative; width: 40px; height: 22px; border-radius: 999px; background: #c9d3d0; transition: background .15s; flex: none; }
	.sw .track::after { content: ""; position: absolute; top: 3px; left: 3px; width: 16px; height: 16px; border-radius: 50%; background: #fff; box-shadow: 0 1px 2px rgba(0,0,0,.25); transition: transform .15s; }
	.sw input:checked + .track { background: var(--brand); }
	.sw input:checked + .track::after { transform: translateX(18px); }
	.sw input:focus-visible + .track { outline: 2px solid var(--brand); outline-offset: 2px; }
	.sw input:disabled + .track { opacity: .55; cursor: not-allowed; }
	.sw .lbl { font-size: 12px; font-weight: 600; min-width: 48px; }
	.sw .lbl.on { color: #136a4b; } .sw .lbl.off { color: #a1342a; }
	.sw.busy .track { opacity: .6; }
	.mm-foot { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 10px; padding: 12px 14px; border-top: 1px solid var(--line); color: var(--muted); font-size: 13px; }
	.mm-foot b { color: var(--ink); }
	.mm-pager { display: flex; flex-wrap: wrap; gap: 4px; }
	.mm-pager button { min-width: 34px; height: 34px; padding: 0 10px; font: 600 12.5px inherit; background: #fff; border: 1px solid var(--line); border-radius: 8px; cursor: pointer; color: var(--ink); }
	.mm-pager button[aria-current="page"] { color: #fff; background: var(--brand); border-color: var(--brand); }
	.mm-pager button:disabled { opacity: .45; cursor: default; }
	.mm-pager span { align-self: center; padding: 0 4px; }
	.mm-size select { height: 34px; }
	.mm-empty { padding: 40px 12px; text-align: center; color: var(--muted); }
	.mm-deny { padding: 26px; background: #fff; border: 1px solid var(--line); border-radius: 12px; color: var(--muted); }
	.mm-deny b { display: block; margin-bottom: 4px; font-size: 15px; color: var(--ink); }
	/* CSS form lama (fbp-forms.css) menata semua input/tombol di .box_both: dikalahkan di sini */
	body .mm .mm-search input { padding-left: 36px !important; height: 38px !important; min-height: 38px !important; border-radius: 9px !important; }
	body .mm .mm-seg button { height: 32px !important; min-height: 32px !important; padding: 0 12px !important; color: var(--muted) !important; background: transparent !important; border: 0 !important; border-radius: 8px !important; box-shadow: none !important; font-weight: 600 !important; }
	body .mm .mm-seg button[aria-pressed="true"] { color: var(--brand-dark) !important; background: #fff !important; box-shadow: 0 1px 3px rgba(23,38,42,.12) !important; }
	body .mm .mm-pager button { min-height: 34px !important; height: 34px !important; padding: 0 10px !important; color: var(--ink) !important; background: #fff !important; border: 1px solid var(--line) !important; border-radius: 8px !important; box-shadow: none !important; }
	body .mm .mm-pager button[aria-current="page"] { color: #fff !important; background: var(--brand) !important; border-color: var(--brand) !important; }
	body .mm .mm-tools select, body .mm .mm-size select { width: auto !important; min-width: 150px; }
	body .mm .mm-size select { min-width: 70px; }
	.mm-toast { position: fixed; right: 20px; bottom: 20px; z-index: 50; max-width: 380px; padding: 11px 16px; color: #fff; background: #17262a; border-radius: 10px; font-size: 13px; box-shadow: 0 10px 30px rgba(0,0,0,.2); display: none; }
	.mm-toast.err { background: #a1342a; }
	/* nomor urut & aksi */
	.mm thead th.no, .mm tbody td.no { width: 1%; text-align: right; color: var(--muted); font-variant-numeric: tabular-nums; white-space: nowrap; }
	.mm thead th.act, .mm tbody td.act { width: 1%; text-align: center; white-space: nowrap; }
	.mm-add { display: inline-flex; align-items: center; gap: 6px; margin-left: auto; height: 38px; padding: 0 14px; font: 600 13px inherit; color: #fff; background: var(--brand); border: 0; border-radius: 9px; cursor: pointer; }
	.mm-add:hover { background: var(--brand-dark); }
	.mm-edit { display: inline-flex; align-items: center; gap: 5px; height: 30px; padding: 0 10px; font: 600 12px inherit; color: var(--brand-dark); background: var(--tint); border: 1px solid #cfe3dc; border-radius: 8px; cursor: pointer; }
	.mm-edit:hover { background: #dff0ea; }
	body .mm .mm-add { height: 38px !important; min-height: 38px !important; padding: 0 14px !important; color: #fff !important; background: var(--brand) !important; border: 0 !important; border-radius: 9px !important; box-shadow: none !important; width: auto !important; }
	body .mm .mm-edit { height: 30px !important; min-height: 30px !important; padding: 0 10px !important; color: var(--brand-dark) !important; background: var(--tint) !important; border: 1px solid #cfe3dc !important; border-radius: 8px !important; box-shadow: none !important; width: auto !important; }
	/* form tambah/ubah */
	.mm-modal { position: fixed; inset: 0; z-index: 60; display: flex; align-items: flex-start; justify-content: center; padding: 6vh 16px 16px; background: rgba(23,38,42,.45); overflow: auto; font-family: "Plus Jakarta Sans", system-ui, sans-serif; color: #17262a; }
	.mm-modal[hidden] { display: none; }
	.mm-dlg { width: 100%; max-width: 640px; background: #fff; border-radius: 14px; box-shadow: 0 20px 50px rgba(0,0,0,.25); }
	.mm-dlg-head { display: flex; align-items: center; justify-content: space-between; padding: 16px 20px; border-bottom: 1px solid #e1e7e5; }
	.mm-dlg-head h3 { margin: 0; font-size: 17px; font-weight: 700; }
	.mm-x { width: 34px; height: 34px; font-size: 22px; line-height: 1; color: #5d6d69; background: transparent; border: 0; border-radius: 8px; cursor: pointer; }
	.mm-x:hover { background: #eef2f1; }
	.mm-dlg-body { padding: 16px 20px 4px; }
	.mm-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 16px; }
	.mm-grid .mm-wide { grid-column: 1 / -1; }
	.mm-grid .f { display: flex; flex-direction: column; gap: 5px; font-size: 12.5px; font-weight: 600; color: #3d4d49; }
	.mm-grid .f em { display: inline; color: #a1342a; font-style: normal; }
	.mm-grid input[type=text], .mm-grid input[type=number], .mm-grid select, .mm-grid textarea { width: 100%; box-sizing: border-box; height: 38px; padding: 0 11px; font: 400 13px "Plus Jakarta Sans", system-ui, sans-serif; color: #17262a; background: #fff; border: 1px solid #d5dfdc; border-radius: 9px; outline: none; }
	.mm-grid textarea { height: auto; min-height: 60px; padding: 8px 11px; resize: vertical; }
	.mm-grid input:focus, .mm-grid select:focus, .mm-grid textarea:focus { border-color: #1B7863; box-shadow: 0 0 0 3px rgba(27,120,99,.18); }
	.mm-grid input[readonly] { color: #5d6d69; background: #f2f5f4; }
	.mm-grid .bad { border-color: #c2483b !important; }
	.mm-grid small { font-size: 11.5px; font-weight: 400; color: #5d6d69; }
	.mm-grid small.warn { color: #a1342a; }
	.mm-err { margin-bottom: 12px; padding: 10px 12px; font-size: 13px; color: #8a2a20; background: #fdecea; border: 1px solid #f3c2bc; border-radius: 9px; }
	.mm-note { margin: 14px 0 12px; font-size: 12px; color: #5d6d69; }
	.mm-dlg-foot { display: flex; justify-content: flex-end; gap: 8px; padding: 14px 20px; border-top: 1px solid #e1e7e5; }
	.mm-btn { height: 38px; padding: 0 18px; font: 600 13px "Plus Jakarta Sans", system-ui, sans-serif; color: #fff; background: #1B7863; border: 1px solid #1B7863; border-radius: 9px; cursor: pointer; }
	.mm-btn:hover { background: #12584a; }
	.mm-btn.ghost { color: #3d4d49; background: #fff; border-color: #d5dfdc; }
	.mm-btn.ghost:hover { background: #f2f5f4; }
	.mm-btn:disabled { opacity: .6; cursor: default; }
	@media (max-width: 560px) { .mm-grid { grid-template-columns: 1fr; } .mm-add { margin-left: 0; } }
</style>
</head>
<body class="dash page-show page-legacy">
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
	<% if Session("fcBa") <> "" then %>
	<!--#include file="menu.asp"-->
	<% end if %>
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
	<div class="judul">Module : User Access Control / Master Menu</div>
	<div class="box_both">
	<div class="mm" data-fbt-auto="off">
<% If Not mmAllowed Then %>
		<div class="mm-deny"><b>Akses ditolak</b>Anda tidak punya akses ke Master Menu. Minta administrator menambahkan menu ini ke role Anda di <i>Combination Roles And Menu</i>.</div>
<% Else %>
		<div class="mm-head">
			<div>
				<h2>Master Menu</h2>
				<p>Aktifkan atau nonaktifkan menu aplikasi (MENU_TAB). Perubahan tersimpan langsung dan tampil di sidebar setelah halaman dimuat ulang.</p>
			</div>
			<div class="mm-kpis">
				<div class="mm-kpi"><span>Total menu</span><b id="kAll">0</b></div>
				<div class="mm-kpi on"><span>Aktif</span><b id="kOn">0</b></div>
				<div class="mm-kpi off"><span>Nonaktif</span><b id="kOff">0</b></div>
			</div>
		</div>
		<div class="mm-card">
			<div class="mm-tools">
				<div class="mm-search"><i class="fa fa-search" aria-hidden="true"></i><input type="search" id="mmQ" placeholder="Cari kode, nama, induk, atau link..." aria-label="Cari menu"></div>
				<div class="mm-seg" role="group" aria-label="Filter status">
					<button type="button" data-st="all" aria-pressed="true">Semua</button>
					<button type="button" data-st="on" aria-pressed="false">Aktif</button>
					<button type="button" data-st="off" aria-pressed="false">Nonaktif</button>
				</div>
				<select id="mmType" aria-label="Filter tipe menu"><option value="">Semua tipe</option></select>
				<select id="mmMod" aria-label="Filter modul"><option value="">Semua modul</option></select>
				<button type="button" class="mm-add" id="mmAdd"><i class="fa fa-plus" aria-hidden="true"></i> Tambah Menu</button>
			</div>
			<div class="mm-scroll">
				<table>
					<thead><tr>
						<th class="ns no">No</th>
						<th data-k="0">Kode<span class="ar">&#8597;</span></th>
						<th data-k="1">Nama Menu<span class="ar">&#8597;</span></th>
						<th data-k="4">Modul<span class="ar">&#8597;</span></th>
						<th data-k="7">Tipe<span class="ar">&#8597;</span></th>
						<th data-k="3">Link<span class="ar">&#8597;</span></th>
						<th data-k="6">Urutan<span class="ar">&#8597;</span></th>
						<th data-k="9">Diubah<span class="ar">&#8597;</span></th>
						<th data-k="5">Status<span class="ar">&#8597;</span></th>
						<th class="ns act">Aksi</th>
					</tr></thead>
					<tbody id="mmBody"></tbody>
				</table>
			</div>
			<div class="mm-foot">
				<div id="mmInfo"></div>
				<label class="mm-size">Tampilkan <select id="mmLen"><option>10</option><option selected>25</option><option>50</option><option>100</option></select></label>
				<nav class="mm-pager" id="mmPager" aria-label="Halaman"></nav>
			</div>
		</div>
<% End If %>
	</div>
	</div>
</div>
<div class="mm-toast" id="mmToast" role="status"></div>
<% If mmAllowed Then %>
<div class="mm-modal" id="mmModal" hidden>
	<div class="mm-dlg" role="dialog" aria-modal="true" aria-labelledby="mmFTitle">
		<form id="mmForm" novalidate autocomplete="off">
			<div class="mm-dlg-head">
				<h3 id="mmFTitle">Tambah Menu</h3>
				<button type="button" class="mm-x" data-close aria-label="Tutup">&times;</button>
			</div>
			<div class="mm-dlg-body">
				<div class="mm-err" id="mmFErr" role="alert" hidden></div>
				<div class="mm-grid">
					<label class="f"><span class="t">Kode Menu <em>*</em></span>
						<input type="text" name="code" maxlength="30" required>
						<small id="mmCodeHint">Huruf, angka, titik, garis bawah atau minus. Tidak bisa diubah setelah disimpan.</small>
					</label>
					<label class="f"><span class="t">Nama Menu <em>*</em></span>
						<input type="text" name="name" maxlength="100" required>
					</label>
					<label class="f mm-wide"><span class="t">Menu Induk <em>*</em></span>
						<input type="text" name="parent" list="mmParents" maxlength="30" required placeholder="Ketik kode atau nama induk, 0 untuk menu utama">
						<small id="mmParentName">&nbsp;</small>
					</label>
					<label class="f"><span class="t">Link (formlink)</span>
						<input type="text" name="link" maxlength="100" placeholder="mis. MasterMenu atau p_laporan.asp">
						<small>Kosongkan untuk menu grup.</small>
					</label>
					<label class="f"><span class="t">Urutan</span>
						<input type="number" name="seq" step="any">
					</label>
					<label class="f"><span class="t">Modul <em>*</em></span>
						<select name="module" required><option value="">Pilih modul</option></select>
					</label>
					<label class="f"><span class="t">Tipe <em>*</em></span>
						<select name="type" required><option value="">Pilih tipe</option></select>
					</label>
					<label class="f mm-wide"><span class="t">Keterangan Bantuan</span>
						<input type="text" name="help" maxlength="100">
					</label>
					<label class="f mm-wide"><span class="t">Catatan</span>
						<textarea name="remarks" maxlength="1000" rows="2"></textarea>
					</label>
					<div class="f mm-wide">
						<label class="sw"><input type="checkbox" name="active" checked><span class="track"></span><span class="lbl on">Aktif</span></label>
						<small id="mmLockHint" hidden>Menu ini dibutuhkan untuk membuka Master Menu, jadi tidak bisa dinonaktifkan.</small>
					</div>
				</div>
				<p class="mm-note" id="mmRoleNote">Menu baru belum muncul di sidebar sebelum ditambahkan ke role pengguna di <i>Combination Roles And Menu</i>.</p>
			</div>
			<div class="mm-dlg-foot">
				<button type="button" class="mm-btn ghost" data-close>Batal</button>
				<button type="submit" class="mm-btn" id="mmSave">Simpan</button>
			</div>
		</form>
	</div>
</div>
<datalist id="mmParents"></datalist>
<script>
(function () {
	'use strict';
	// [kode, nama, induk, link, modul, aktif, urutan, tipe, diubah_oleh, tgl_ubah, keterangan, catatan, kode_modul]
	var ROWS = [<%=mmRows%>];
	var LOCKED = {<% For Each mmK In mmLockList.Keys %><%=MmJson(mmK)%>:1,<% Next %>"":0};
	var TOK = <%=MmJson(Session("mm_tok"))%>;
	var MODULES = [<%=mmModules%>], TYPES = [<%=mmTypes%>];
	var byCode = {}, kids = {};
	ROWS.forEach(function (r) { byCode[r[0]] = r; (kids[r[2]] = kids[r[2]] || []).push(r); });
	var st = { q: '', status: 'all', type: '', mod: '', sort: 0, dir: 1, page: 1, len: 25 };
	var $ = function (id) { return document.getElementById(id); };

	function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
	function fmt(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
	function path(r) {
		var out = [], p = byCode[r[2]], guard = 0;
		while (p && guard++ < 8) { out.unshift(p[1]); p = byCode[p[2]]; }
		return out.join(' \u203a ');
	}
	function reindex() {
		byCode = {}; kids = {};
		ROWS.forEach(function (r) { byCode[r[0]] = r; (kids[r[2]] = kids[r[2]] || []).push(r); });
		ROWS.forEach(function (r) { r.path = path(r); r.hay = (r[0] + ' ' + r[1] + ' ' + r.path + ' ' + r[3] + ' ' + r[4]).toLowerCase(); });
	}
	reindex();
	function mark(s) {
		var t = esc(s), words = st.q.trim().split(/\s+/).filter(Boolean);
		if (!words.length) return t;
		var re = new RegExp('(' + words.map(function (w) { return esc(w).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }).join('|') + ')', 'gi');
		return t.replace(re, '<mark>$1</mark>');
	}
	function toast(m, err) { var t = $('mmToast'); t.textContent = m; t.className = 'mm-toast' + (err ? ' err' : ''); t.style.display = 'block'; clearTimeout(toast.t); toast.t = setTimeout(function () { t.style.display = 'none'; }, 3500); }

	// isi pilihan filter
	function fillSelect(sel, idx) {
		var seen = {};
		ROWS.forEach(function (r) { if (r[idx]) seen[r[idx]] = 1; });
		Object.keys(seen).sort().forEach(function (v) { var o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); });
	}
	fillSelect($('mmType'), 7);
	fillSelect($('mmMod'), 4);

	function kpis() {
		var on = ROWS.filter(function (r) { return r[5]; }).length;
		$('kAll').textContent = fmt(ROWS.length); $('kOn').textContent = fmt(on); $('kOff').textContent = fmt(ROWS.length - on);
	}

	function filtered() {
		var words = st.q.toLowerCase().split(/\s+/).filter(Boolean);
		var list = ROWS.filter(function (r) {
			if (st.status === 'on' && !r[5]) return false;
			if (st.status === 'off' && r[5]) return false;
			if (st.type && r[7] !== st.type) return false;
			if (st.mod && r[4] !== st.mod) return false;
			return words.every(function (w) { return r.hay.indexOf(w) !== -1; });
		});
		var k = st.sort, d = st.dir;
		list.sort(function (a, b) {
			var x = a[k], y = b[k];
			if (k === 6) { x = parseFloat(x) || 0; y = parseFloat(y) || 0; }
			else if (k === 9) { x = String(x).split('/').reverse().join(''); y = String(y).split('/').reverse().join(''); }
			else if (k === 5) { x = x ? 1 : 0; y = y ? 1 : 0; }
			else { x = String(x).toLowerCase(); y = String(y).toLowerCase(); }
			return x < y ? -d : x > y ? d : (a[0] < b[0] ? -1 : 1);
		});
		return list;
	}

	function render() {
		var list = filtered(), pages = Math.max(1, Math.ceil(list.length / st.len));
		st.page = Math.min(st.page, pages);
		var start = (st.page - 1) * st.len, part = list.slice(start, start + st.len);
		$('mmBody').innerHTML = part.length ? part.map(function (r, i) {
			var locked = !!LOCKED[r[0]], child = (kids[r[0]] || []).filter(function (c) { return c[5]; }).length;
			return '<tr class="' + (r[5] ? '' : 'off') + '">' +
				'<td class="no">' + fmt(start + i + 1) + '</td>' +
				'<td><span class="code">' + mark(r[0]) + '</span></td>' +
				'<td>' + mark(r[1]) + (r.path ? '<span class="path">' + mark(r.path) + '</span>' : '') + '</td>' +
				'<td>' + (r[4] ? mark(r[4]) : '<span class="muted">-</span>') + '</td>' +
				'<td><span class="type ' + esc(r[7]) + '">' + esc(r[7] || '-') + '</span></td>' +
				'<td>' + (r[3] ? '<span class="link">' + mark(r[3]) + '</span>' : '<span class="muted">grup</span>') + '</td>' +
				'<td>' + esc(r[6]) + '</td>' +
				'<td>' + (r[9] ? esc(r[9]) + (r[8] ? '<span class="path">' + esc(r[8]) + '</span>' : '') : '<span class="muted">-</span>') + '</td>' +
				'<td><label class="sw" title="' + (locked ? 'Menu ini tidak bisa dinonaktifkan (dibutuhkan untuk membuka Master Menu)' : (child ? child + ' submenu aktif di bawahnya' : 'Aktif / nonaktifkan menu')) + '">' +
				'<input type="checkbox" data-code="' + esc(r[0]) + '"' + (r[5] ? ' checked' : '') + (locked && r[5] ? ' disabled' : '') + ' aria-label="Status menu ' + esc(r[1]) + '">' +
				'<span class="track"></span><span class="lbl ' + (r[5] ? 'on' : 'off') + '">' + (r[5] ? 'Aktif' : 'Nonaktif') + '</span></label></td>' +
				'<td class="act"><button type="button" class="mm-edit" data-edit="' + esc(r[0]) + '" aria-label="Edit menu ' + esc(r[1]) + '"><i class="fa fa-pencil" aria-hidden="true"></i> Edit</button></td>' +
				'</tr>';
		}).join('') : '<tr><td colspan="10" class="mm-empty">Tidak ada menu yang cocok dengan filter.</td></tr>';
		$('mmInfo').innerHTML = list.length
			? 'Menampilkan <b>' + fmt(start + 1) + '&ndash;' + fmt(start + part.length) + '</b> dari <b>' + fmt(list.length) + '</b> menu' + (list.length !== ROWS.length ? ' (difilter dari ' + fmt(ROWS.length) + ')' : '')
			: '0 menu';
		pager(pages);
		Array.prototype.forEach.call(document.querySelectorAll('.mm thead th[data-k]'), function (th) {
			if (+th.getAttribute('data-k') === st.sort) { th.setAttribute('aria-sort', st.dir > 0 ? 'ascending' : 'descending'); th.querySelector('.ar').textContent = st.dir > 0 ? '\u2191' : '\u2193'; }
			else { th.removeAttribute('aria-sort'); th.querySelector('.ar').textContent = '\u2195'; }
		});
	}

	function pager(pages) {
		var cur = st.page, b = function (l, p, o) { o = o || {}; return '<button type="button" data-p="' + p + '"' + (o.dis ? ' disabled' : '') + (o.cur ? ' aria-current="page"' : '') + '>' + l + '</button>'; };
		var h = b('&lsaquo;', cur - 1, { dis: cur <= 1 }), list = [];
		if (pages <= 7) { for (var i = 1; i <= pages; i++) list.push(i); }
		else {
			list.push(1);
			var a = Math.max(2, cur - 1), z = Math.min(pages - 1, cur + 1);
			if (cur <= 3) { a = 2; z = 4; } else if (cur >= pages - 2) { a = pages - 3; z = pages - 1; }
			if (a > 2) list.push(0);
			for (var k = a; k <= z; k++) list.push(k);
			if (z < pages - 1) list.push(0);
			list.push(pages);
		}
		list.forEach(function (n) { h += n === 0 ? '<span>&hellip;</span>' : b(n, n, { cur: n === cur }); });
		$('mmPager').innerHTML = h + b('&rsaquo;', cur + 1, { dis: cur >= pages });
	}

	function save(input) {
		var code = input.getAttribute('data-code'), r = byCode[code], want = input.checked;
		if (!r) return;
		if (!want) {
			var child = (kids[code] || []).filter(function (c) { return c[5]; }).length;
			var msg = 'Nonaktifkan menu "' + r[1] + '" (' + code + ')?' + (child ? '\n\nMenu ini punya ' + child + ' submenu aktif; submenu itu ikut tidak tampil di sidebar selama menu induknya nonaktif.' : '');
			if (!window.confirm(msg)) { input.checked = true; return; }
		}
		var sw = input.closest('.sw');
		sw.classList.add('busy'); input.disabled = true;
		var x = new XMLHttpRequest();
		x.open('POST', 'MasterMenu.asp', true);
		x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
		x.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
		x.onload = function () {
			var d = null; try { d = JSON.parse(x.responseText); } catch (e) { /* bukan JSON */ }
			if (d && d.ok) {
				r[5] = d.active; r[8] = d.by;
				var now = new Date(); r[9] = ('0' + now.getDate()).slice(-2) + '/' + ('0' + (now.getMonth() + 1)).slice(-2) + '/' + now.getFullYear();
				toast('Menu "' + r[1] + '" ' + (d.active ? 'diaktifkan' : 'dinonaktifkan') + '. Sidebar berubah setelah halaman dimuat ulang.');
				kpis(); render();
			} else {
				input.checked = !want; sw.classList.remove('busy'); input.disabled = false;
				toast(d && d.error ? d.error : 'Gagal menyimpan perubahan.', true);
			}
		};
		x.onerror = function () { input.checked = !want; sw.classList.remove('busy'); input.disabled = false; toast('Koneksi gagal, perubahan tidak tersimpan.', true); };
		x.send('action=toggle&code=' + encodeURIComponent(code) + '&active=' + (want ? 'TRUE' : 'FALSE') + '&tok=' + encodeURIComponent(TOK));
	}

	// ---- form tambah / ubah ----
	var modal = $('mmModal'), form = $('mmForm'), F = form.elements, editing = null, lastFocus = null;
	MODULES.forEach(function (m) { var o = document.createElement('option'); o.value = m[0]; o.textContent = m[1] + ' (' + m[0] + ')'; F.module.appendChild(o); });
	TYPES.forEach(function (t) { var o = document.createElement('option'); o.value = t[0]; o.textContent = t[0]; F.type.appendChild(o); });
	function fillParents() {
		var dl = $('mmParents'), h = '<option value="0">0 - Menu utama</option>';
		ROWS.slice().sort(function (a, b) { return a[0] < b[0] ? -1 : 1; }).forEach(function (r) { if (!r[3]) h += '<option value="' + esc(r[0]) + '">' + esc(r[0] + ' - ' + r[1]) + '</option>'; });
		ROWS.forEach(function (r) { if (r[3]) h += '<option value="' + esc(r[0]) + '">' + esc(r[0] + ' - ' + r[1]) + '</option>'; });
		dl.innerHTML = h;
	}
	function parentInfo() {
		var v = F.parent.value.trim(), el = $('mmParentName'), p = byCode[v];
		el.className = '';
		if (!v) el.innerHTML = '&nbsp;';
		else if (v === '0') el.textContent = 'Menu utama (paling atas)';
		else if (p) el.textContent = (p.path ? p.path + ' \u203a ' : '') + p[1];
		else { el.textContent = 'Kode induk tidak ditemukan'; el.className = 'warn'; }
	}
	function setActiveLbl() { var l = form.querySelector('.sw .lbl'); l.textContent = F.active.checked ? 'Aktif' : 'Nonaktif'; l.className = 'lbl ' + (F.active.checked ? 'on' : 'off'); }
	function showErr(m) { var e = $('mmFErr'); e.textContent = m || ''; e.hidden = !m; }
	function openForm(r) {
		editing = r || null;
		form.reset(); showErr('');
		Array.prototype.forEach.call(form.querySelectorAll('.bad'), function (x) { x.classList.remove('bad'); });
		fillParents();
		$('mmFTitle').textContent = r ? 'Edit Menu ' + r[0] : 'Tambah Menu';
		F.code.readOnly = !!r;
		$('mmCodeHint').hidden = !!r;
		$('mmRoleNote').hidden = !!r;
		if (r) {
			F.code.value = r[0]; F.name.value = r[1]; F.parent.value = r[2]; F.link.value = r[3];
			F.seq.value = r[6]; F.module.value = r[12]; F.type.value = r[7];
			F.help.value = r[10] || ''; F.remarks.value = r[11] || ''; F.active.checked = !!r[5];
		} else {
			F.active.checked = true;
			if (st.type) F.type.value = st.type;
		}
		var locked = !!(r && LOCKED[r[0]]);
		F.active.disabled = locked && F.active.checked;
		$('mmLockHint').hidden = !locked;
		setActiveLbl(); parentInfo();
		lastFocus = document.activeElement;
		modal.hidden = false;
		(r ? F.name : F.code).focus();
	}
	function closeForm() { modal.hidden = true; editing = null; if (lastFocus && lastFocus.focus) lastFocus.focus(); }
	// saran urutan: urutan terbesar di bawah induk yang sama + 1
	F.parent.addEventListener('change', function () {
		parentInfo();
		if (!editing && !F.seq.value) {
			var mx = 0; (kids[F.parent.value.trim()] || []).forEach(function (c) { mx = Math.max(mx, parseFloat(c[6]) || 0); });
			if (mx) F.seq.value = mx + 1;
		}
	});
	F.parent.addEventListener('input', parentInfo);
	F.active.addEventListener('change', setActiveLbl);
	modal.addEventListener('click', function (e) { if (e.target === modal || e.target.closest('[data-close]')) closeForm(); });
	modal.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeForm(); });
	form.addEventListener('submit', function (e) {
		e.preventDefault();
		var bad = [], v = function (n) { return F[n].value.trim(); };
		Array.prototype.forEach.call(form.querySelectorAll('.bad'), function (x) { x.classList.remove('bad'); });
		if (!/^[A-Za-z0-9_.\-]{1,30}$/.test(v('code'))) bad.push(['code', 'Kode menu wajib diisi (huruf, angka, titik, garis bawah atau minus).']);
		else if (!editing && byCode[v('code')]) bad.push(['code', 'Kode menu ' + v('code') + ' sudah dipakai.']);
		if (!v('name')) bad.push(['name', 'Nama menu wajib diisi.']);
		if (!v('parent')) bad.push(['parent', 'Pilih menu induk (0 untuk menu utama).']);
		else if (v('parent') !== '0' && !byCode[v('parent')]) bad.push(['parent', 'Menu induk ' + v('parent') + ' tidak ditemukan.']);
		else if (v('parent') === v('code')) bad.push(['parent', 'Menu induk tidak boleh menu itu sendiri.']);
		if (!v('module')) bad.push(['module', 'Pilih modul.']);
		if (!v('type')) bad.push(['type', 'Pilih tipe menu.']);
		if (bad.length) {
			bad.forEach(function (b) { F[b[0]].classList.add('bad'); });
			showErr(bad.map(function (b) { return b[1]; }).join(' '));
			F[bad[0][0]].focus();
			return;
		}
		showErr('');
		var btn = $('mmSave'); btn.disabled = true; btn.textContent = 'Menyimpan...';
		var body = ['action=save', 'mode=' + (editing ? 'update' : 'insert'), 'tok=' + encodeURIComponent(TOK), 'active=' + (F.active.checked ? 'TRUE' : 'FALSE')];
		['code', 'name', 'parent', 'link', 'seq', 'module', 'type', 'help', 'remarks'].forEach(function (n) { body.push(n + '=' + encodeURIComponent(v(n))); });
		var x = new XMLHttpRequest();
		x.open('POST', 'MasterMenu.asp', true);
		x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
		x.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
		x.onload = function () {
			var d = null; try { d = JSON.parse(x.responseText); } catch (er) { /* bukan JSON */ }
			btn.disabled = false; btn.textContent = 'Simpan';
			if (d && d.ok) {
				var row = d.row, idx = -1;
				ROWS.forEach(function (r, i) { if (r[0] === row[0]) idx = i; });
				if (idx >= 0) ROWS[idx] = row; else ROWS.push(row);
				var isNew = idx < 0;
				reindex(); kpis(); render(); closeForm();
				toast('Menu "' + row[1] + '" ' + (isNew ? 'ditambahkan. Tambahkan ke role di Combination Roles And Menu agar tampil di sidebar.' : 'disimpan. Sidebar berubah setelah halaman dimuat ulang.'));
			} else {
				showErr(d && d.error ? d.error : 'Gagal menyimpan menu (HTTP ' + x.status + ').');
			}
		};
		x.onerror = function () { btn.disabled = false; btn.textContent = 'Simpan'; showErr('Koneksi gagal, data tidak tersimpan.'); };
		x.send(body.join('&'));
	});
	$('mmAdd').addEventListener('click', function () { openForm(null); });
	$('mmBody').addEventListener('click', function (e) { var b = e.target.closest('button[data-edit]'); if (b) openForm(byCode[b.getAttribute('data-edit')]); });

	// ---- event ----
	var deb;
	$('mmQ').addEventListener('input', function () { var v = this.value; clearTimeout(deb); deb = setTimeout(function () { st.q = v; st.page = 1; render(); }, 200); });
	Array.prototype.forEach.call(document.querySelectorAll('.mm-seg button'), function (bt) {
		bt.addEventListener('click', function () {
			Array.prototype.forEach.call(document.querySelectorAll('.mm-seg button'), function (x) { x.setAttribute('aria-pressed', x === bt ? 'true' : 'false'); });
			st.status = bt.getAttribute('data-st'); st.page = 1; render();
		});
	});
	$('mmType').addEventListener('change', function () { st.type = this.value; st.page = 1; render(); });
	$('mmMod').addEventListener('change', function () { st.mod = this.value; st.page = 1; render(); });
	$('mmLen').addEventListener('change', function () { st.len = parseInt(this.value, 10) || 25; st.page = 1; render(); });
	$('mmPager').addEventListener('click', function (e) { var bt = e.target.closest('button[data-p]'); if (bt && !bt.disabled) { st.page = +bt.getAttribute('data-p'); render(); } });
	document.querySelector('.mm thead').addEventListener('click', function (e) {
		var th = e.target.closest('th[data-k]'); if (!th) return;
		var k = +th.getAttribute('data-k');
		if (st.sort === k) st.dir = -st.dir; else { st.sort = k; st.dir = 1; }
		render();
	});
	$('mmBody').addEventListener('change', function (e) { if (e.target.matches('input[data-code]')) save(e.target); });

	kpis(); render();
})();
</script>
<% End If %>
<!--#include file="bottom.jsp"-->
</body>
</html>
