<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transportercontractno = Request("transportercontractno")
fcba = Request("fcba")
lineno = Request("lineno")
result = ""

html = html & "<select id='lineno' name='transporterline' onchange='javascript:getTransporter();javascript:getselectedtransporterline();' style='width: 40%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT lineno, shippingaddress " & _
	" FROM TRANSPORTERCONTRACT " & _
	" WHERE transportercontractno= '"& transportercontractno &"' and fcba = '" & fcba & "'"  

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("lineno") & "'"
		if "" & lineno <> "" then
			if cstr(""&lineno) = cstr(""&rsCallCombo("lineno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("lineno") & " : " & rsCallCombo("shippingaddress")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>