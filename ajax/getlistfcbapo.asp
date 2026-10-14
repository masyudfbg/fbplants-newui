<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccompanycode = Request("fccompanycode")
fcba = Request("fcba")
result = ""

html = html & "<select id='fcba' name='businessunit' onchange='javascript:getselectedfcba();reloadPoCode();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT fcba, fcname FROM user_fcba_tab uft INNER JOIN businessunit bu ON bu.fccode = uft.fcba WHERE fccompanycode = '" & fccompanycode & "' AND uft.fcusercode = '" & Session("CodeName") & "'"   

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fcba") & "'"
		if "" & fcba <> "" then
			if cstr(""&fcba) = cstr(""&rsCallCombo("fcba")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fcba") & " : " & rsCallCombo("fcname")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop
End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>