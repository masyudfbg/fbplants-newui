<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transporter = Request("transporter")
result = ""

q = " SELECT *" & _
	" FROM transportercontract" & _
	" WHERE transporter = '"& transporter &"' and fcba = '" & Session("fcba") & "'" 

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("transportercontractno") & "'" 
		if Request("transportercontractno") = rsCallCombo("transportercontractno") then html = html &  " selected"
		html = html &  " >"
		html = html & rsCallCombo("transportercontractno")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

else
	html = "<option value=''></option>"
End if
rsCallCombo.Close
set rsCallCombo = Nothing
					  
	response.write(html)
%>