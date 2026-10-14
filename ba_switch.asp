<%@Language=VBScript%>
<% Option Explicit %>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="include/ba_session.asp"-->
<%
' =====================================================================
' Ganti Business Unit (BA) tanpa logout - dipakai dropdown BA di dashboard (login.asp).
'   GET  ba_switch.asp?action=list          -> daftar BA yang di-assign ke user (user_fcba_tab)
'   POST ba_switch.asp  action=switch&fcba=XXX&tok=...
'        -> BA default user ikut diganti (sama seperti UsersChangedBusinessUnit.asp),
'           session diisi ulang (include/ba_session.asp), dicatat di HISTORICAL_USER_TAB.
' Hanya BA milik user yang bisa dipilih; token sesi (Session("ba_tok")) wajib untuk POST.
' =====================================================================
Response.ContentType = "application/json"
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

Function JsonStr(v)
	Dim s, i, c, o
	s = "" & v
	o = ""
	For i = 1 To Len(s)
		c = Mid(s, i, 1)
		Select Case AscW(c)
			Case 34: o = o & "\"""
			Case 92: o = o & "\\"
			Case 8: o = o & "\b"
			Case 9: o = o & "\t"
			Case 10: o = o & "\n"
			Case 12: o = o & "\f"
			Case 13: o = o & "\r"
			Case Else
				If AscW(c) < 32 Then
					o = o & "\u" & Right("000" & Hex(AscW(c)), 4)
				Else
					o = o & c
				End If
		End Select
	Next
	JsonStr = """" & o & """"
End Function

Sub Fail(code, msg)
	Response.Status = code
	Response.Write "{""ok"":false,""error"":" & JsonStr(msg) & "}"
	Response.End
End Sub

Dim userId, userSql, action
userId = Trim(Session("CodeName") & "")
If userId = "" Or Trim(Session("USERNAME") & "") = "" Then Fail "401 Unauthorized", "Sesi berakhir, silakan login ulang."
userSql = Replace(userId, "'", "''")
action = LCase(Trim(Request("action") & ""))

' BA yang boleh dipilih: di-assign di user_fcba_tab DAN PT-nya di-assign di user_company_tab.
' User tanpa PT sama sekali di user_company_tab tetap melihat semua BA dari user_fcba_tab.
Dim sqlAllowed
sqlAllowed = " from user_fcba_tab a, businessunit b " & _
             " where b.isactive = 'TRUE' and a.fcba = b.fccode and a.fcusercode = '" & userSql & "' " & _
             " and ( b.fccompanycode in (select c.fccompanycode from user_company_tab c where c.fcusercode = a.fcusercode) " & _
             "       or not exists (select 1 from user_company_tab c2 where c2.fcusercode = a.fcusercode) ) "

Dim sqlList
sqlList = " select b.fccode, b.fcname, b.fccompanycode, b.fccompanyname, a.isdefault " & sqlAllowed & _
          " order by b.fccompanycode, b.fcname "

If action = "list" Then
	Dim rsL, out, first
	Set rsL = connStr.Execute(sqlList)
	out = "{""ok"":true,""current"":" & JsonStr(Session("fcBa")) & ",""items"":["
	first = True
	Do While Not rsL.EOF
		If Not first Then out = out & ","
		first = False
		out = out & "{""code"":" & JsonStr(rsL("fccode")) & _
		            ",""name"":" & JsonStr(rsL("fcname")) & _
		            ",""company"":" & JsonStr(rsL("fccompanycode")) & _
		            ",""companyName"":" & JsonStr(rsL("fccompanyname")) & _
		            ",""isDefault"":" & LCase(CStr(UCase(Trim(rsL("isdefault") & "")) = "TRUE")) & "}"
		rsL.MoveNext
	Loop
	rsL.Close
	Set rsL = Nothing
	Response.Write out & "]}"
	Response.End
End If

If action <> "switch" Then Fail "400 Bad Request", "Aksi tidak dikenal."
If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then Fail "405 Method Not Allowed", "Gunakan POST."
If Trim(Session("ba_tok") & "") = "" Or Request.Form("tok") <> Session("ba_tok") Then Fail "403 Forbidden", "Token tidak valid, muat ulang halaman."

Dim fcba, re
fcba = UCase(Trim(Request.Form("fcba") & ""))
Set re = New RegExp
re.Pattern = "^[A-Z0-9_.\-]{1,20}$"
If Not re.Test(fcba) Then Fail "400 Bad Request", "Kode BA tidak valid."

' BA harus milik user dan aktif
Dim rsChk, allowed
Set rsChk = connStr.Execute(" select count(*) n " & sqlAllowed & " and b.fccode = '" & fcba & "' ")
allowed = CLng(rsChk("n")) > 0
rsChk.Close
Set rsChk = Nothing
If Not allowed Then Fail "403 Forbidden", "BA " & fcba & " tidak di-assign ke user ini."

Dim rsBU
Set rsBU = connStr.Execute("Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where isactive='TRUE' and fcCode='" & fcba & "'")
If rsBU.EOF Then Fail "404 Not Found", "BA tidak ditemukan."

' Session diisi ulang dulu; BA default di database diganti setelah berhasil (sama seperti UsersChangedBusinessUnit.asp)
FbpSetSessionBA rsBU
rsBU.Close
Set rsBU = Nothing

connStr.Execute "update user_fcba_tab set isdefault = 'FALSE' where fcusercode = '" & userSql & "'"
connStr.Execute "update user_fcba_tab set isdefault = 'TRUE' where fcusercode = '" & userSql & "' and fcba = '" & fcba & "'"
connStr.Execute "update user_company_tab set isdefault = 'FALSE' where fcusercode = '" & userSql & "'"
connStr.Execute "update user_company_tab set isdefault = 'TRUE' where fcusercode = '" & userSql & "' and fccompanycode = '" & Replace(Session("COMPANYCODE") & "", "'", "''") & "'"
connStr.Execute "Insert into HISTORICAL_USER_TAB ( USERNAME, IP_ADDRESS, BU, DATE_ACT, TIME_ACT, LASTUPDATE, LASTTIME ) " & _
                "values ('" & userSql & "','" & Replace(Request.ServerVariables("REMOTE_ADDR"), "'", "") & "','" & fcba & "','" & FormatDateTime(Date, 1) & "','" & FormatDateTime(Date, 3) & "', sysdate, to_char(sysdate, 'HH24:MI'))"
connStr.Execute "commit"

Response.Write "{""ok"":true,""fcba"":" & JsonStr(Session("fcBa")) & ",""name"":" & JsonStr(Session("BA")) & ",""company"":" & JsonStr(Session("COMPANYCODE")) & "}"
%>
