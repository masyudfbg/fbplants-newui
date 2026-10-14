<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transportercontractno = Request("transportercontractno")
fcba = Request("fcba")
result = ""

html = html & "<select id='transportercontractno' name='transportercontractno' onchange='javascript:getselectedtransportercontract(); javascript:getlistsalesshipment();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT DISTINCT transportercontractno " & _
	" FROM transportercontract " & _
	" WHERE fcba = '"& fcba &"'"   

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("transportercontractno") & "'"
		if "" & transportercontractno <> "" then
			if cstr(""&transportercontractno) = cstr(""&rsCallCombo("transportercontractno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("transportercontractno") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>