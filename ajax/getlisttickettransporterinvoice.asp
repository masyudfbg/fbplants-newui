<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salesshipmentno = Request("salesshipmentno")
fcba = Request("fcba")
wbt_chitno = Request("wbt_chitno")
result = ""

html = html & "<select id='wbt_chitno' name='wbt_chitno' onchange='javascript:getSalesShipmentInfo();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT wbt_chitno" & _
	"  FROM salesshipmenthdr ssh" & _
	"       INNER JOIN salesshipmentdtl ssd" & _
	"          ON (    ssd.salesshipmentno = ssh.salesshipmentno" & _
	"              AND ssd.fcba = ssh.fcba)" & _
	"       LEFT JOIN deliveryorder do" & _
	"          ON (do.deliveryorderno = ssh.deliveryorderno AND do.fcba = ssh.fcba)" & _
	" WHERE ssh.fcba = '"& fcba &"' AND ssh.salesshipmentno = '"& salesshipmentno &"' AND transporter_invoice is null" 


' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("wbt_chitno") & "'"
		if "" & wbt_chitno <> "" then
			if cstr(""&wbt_chitno) = cstr(""&rsCallCombo("wbt_chitno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("wbt_chitno") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>