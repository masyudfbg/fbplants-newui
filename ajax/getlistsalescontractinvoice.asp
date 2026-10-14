<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salescontractno = Request("salescontractno")
fcba = Request("fcba")
result = ""

html = html & "<select id='salescontractno' name='salescontractno' onchange='javascript:getSalesInvoiceDp();' style='width: 30%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT salescontractno FROM salesagreementhdr WHERE fcba = '"& fcba &"'"   

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("salescontractno") & "'"
		if "" & salescontractno <> "" then
			if cstr(""&salescontractno) = cstr(""&rsCallCombo("salescontractno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("salescontractno") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>