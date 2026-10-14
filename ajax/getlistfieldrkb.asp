<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
fcba = Request("fcba")
fieldcode = Request("fieldcode")
unitkerja = Request("unitkerja")
result = ""

html = html & "<select id='fieldcode' name='fieldcode' onchange='javascript:getselectedfieldcode(); javascript:getNoRKB();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

	q = "SELECT DISTINCT fccode, fcname FROM field WHERE fcba = '" & fcba & "' and division = '"& unitkerja &"' ORDER BY fccode"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & fieldcode <> "" then
			if fieldcode = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
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