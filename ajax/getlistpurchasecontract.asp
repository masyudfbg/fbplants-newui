<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
customercode = Request("customercode")
fcba = Request("fcba")
salescontractno = Request("salescontractno")
result = ""

html = html & "<select id='salescontractno' name='salescontractno' onchange='javascript:getSalesContract();javascript:getselectedsalescontractno();' style='width: 55%;' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT  sah.suppliercode, sah.purchasecontractno, sad.ppn " & _
	" FROM PURCHASEAGREEMENTHDR sah" & _
	" INNER JOIN PURCHASEAGREEMENTDET sad ON (sah.purchasecontractno = sad.purchasecontractno AND sah.fcba = sad.fcba)" & _
	" WHERE sah.suppliercode= '"& customercode &"' and sah.fcba = '" & fcba & "' order by sah.purchasecontractno " 

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("purchasecontractno") & "'"
		if "" & salescontractno <> "" then
			if cstr(""&salescontractno) = cstr(""&rsCallCombo("purchasecontractno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("purchasecontractno")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>