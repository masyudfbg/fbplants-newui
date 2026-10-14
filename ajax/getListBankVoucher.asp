<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
bankcode = Request("bankcode")
businessunit = Request("businessunit")
banktype = Request("banktype")
result = ""

html = html & "<select id='bankcode' name='bankcode' style='width: 100%' class='chosen-select'>"
html = html &"<option value=''></option>"

 q = 	" select count(module_01) module_01 from users_tab" & _
		" where fccode ='"& session("CodeName") &"'"

	 ' response.write q
	 set rsAmount = server.CreateObject("Adodb.recordset")
	 rsAmount.Open q, connstr
	 if Not rsAmount.Eof then
		while not rsAmount.eof
			module_01 = rsAmount("module_01")
			rsAmount.movenext
		wend
	End if
if cdbl(to_nol(module_01)) = 0 then
	
	q = " SELECT a.fccode, a.fcname" & _
		"  FROM bank a" & _
		" WHERE a.fcba = '"& businessunit &"'" & _ 
	    "       AND UPPER(a.banktype) <> ('JOURNALRECLASS')" 

	' if UCase(cstr(""&banktype)) = "JOURNAL" then
	' 	q = q & "       AND UPPER(banktype) = UPPER('JOURNAL')"
	' else
	' 	q = q & "       AND UPPER(banktype) IN ('CASH', 'BANK')"
	' end if 
Else
	q = " SELECT a.fccode, a.fcname, b.lu_name, c.module_01 " & _
		"  FROM module_tab b" & _
		"       RIGHT JOIN bank a" & _
		"          ON (a.fccode = b.code)" & _
		"          LEFT JOIN users_tab c ON (c.module_01 = b.fccode)" & _
		" WHERE a.fcba = '"& businessunit &"'" & _ 
	    "       AND UPPER(a.banktype) <> ('JOURNALRECLASS')" & _
		" AND c.fccode = '"& session("CodeName") &"'" 

	' if UCase(cstr(""&banktype)) = "JOURNAL" then
	' 	q = q & "       AND UPPER(banktype) = UPPER('JOURNAL')"
	' else
	' 	q = q & "       AND UPPER(banktype) IN ('CASH', 'BANK')"
	' end if
End If
' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & bankcode <> "" then
			if cstr(""&bankcode) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fccode") & " : " & rsCallCombo("fcname")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>