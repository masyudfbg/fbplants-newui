<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
productcode = Request("productcode")
fcba = Request("fcba")
salescontractno = Request("salescontractno")
result = ""

html = html & "<select id='salescontractno' name='salescontractno' onchange='javascript:getselectedsalescontractno();' style='width: 55%;' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT  sah.customercode, sah.salescontractno, sad.vat " & _
	" FROM SALESAGREEMENTHDR sah" & _
	" INNER JOIN SALESAGREEMENTDET sad ON (sah.salescontractno = sad.salescontractno AND sah.fcba = sad.fcba)" & _
	" WHERE sad.productcode= '"& productcode &"' and sah.fcba = '" & fcba & "' order by sah.salescontractno " 

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