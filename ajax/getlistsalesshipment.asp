<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transportercontractno = Request("transportercontractno")
fcba = Request("fcba")
salesshipmentno = Request("salesshipmentno")
result = ""

html = html & "<select id='salesshipmentno' name='salesshipmentno'  onchange='javascript:getselectedshipment();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT DISTINCT ssh.salesshipmentno" & _
	"  FROM salesshipmenthdr ssh" & _
	"       INNER JOIN salesshipmentdtl ssd" & _
	"          ON (    ssh.salesshipmentno = ssd.salesshipmentno" & _
	"              AND ssh.fcba = ssd.fcba)" & _
	"       LEFT JOIN deliveryorder do" & _
	"          ON (do.deliveryorderno = ssh.deliveryorderno AND do.fcba = ssh.fcba)" & _
	" WHERE     ssh.fcba = '"& fcba &"'" & _
	"       AND do.transportercontractno = '"& transportercontractno &"'" & _
	"       AND transporter_invoice IS NULL"  

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("salesshipmentno") & "'"
		if "" & salesshipmentno <> "" then
			if cstr(""&salesshipmentno) = cstr(""&rsCallCombo("salesshipmentno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("salesshipmentno") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>