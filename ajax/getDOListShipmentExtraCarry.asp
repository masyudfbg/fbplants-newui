<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
deliveryorderno = Request("deliveryorderno")
fcba = Request("fcba")
result = ""

html = html & "<select id='deliveryorderno' name='deliveryorderno' onchange='javascript:getselecteddo();javascript:getMessage();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT DISTINCT shd.deliveryorderno" & _
	"   FROM salesshipmentdtl sdt" & _
	"        INNER JOIN salesshipmenthdr shd" & _
	"           ON (    sdt.fcba = shd.fcba" & _
	"               AND sdt.salesshipmentno = shd.salesshipmentno)" & _
	"  WHERE     sdt.fcba = '"& fcba &"'" & _
	"        AND sdt.validationextracarry = 1" & _
	"        AND shd.deliveryorderno NOT IN (SELECT deliveryorderno" & _
	"                                          FROM salesshipmentdtl_extracarry" & _
	"                                         WHERE fcba = '"& fcba &"')"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("deliveryorderno") & "'"
		if "" & deliveryorderno <> "" then
			if cstr(""&deliveryorderno) = cstr(""&rsCallCombo("deliveryorderno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("deliveryorderno") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>