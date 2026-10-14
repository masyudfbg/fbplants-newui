<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
stocktransno2 = Request("stocktransno2")
dono2 = Request("dono2")
result = ""

html = html & "<select id='dono2' name='dono2' onchange='javascript:getselecteddo();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = "  SELECT DISTINCT stono, do_no" & _
	"    FROM (SELECT a.stono, a.do_no" & _
	"            FROM deliveryorderpurchasedtl a" & _
	"                 INNER JOIN deliveryorderpurchasehdr b" & _
	"                    ON (a.do_no = b.do_no AND a.fcba = b.fcba)" & _
	"           WHERE     (a.delivery_quantity - NVL(a.receivedquantity,0)) > 0" & _
	"                 AND a.stono = '"& stocktransno2 &"'" & _
	"                 AND a.fcba = '"& session("fcba") &"')" & _
	" ORDER BY do_no " 

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("do_no") & "'"
		if "" & dono2 <> "" then
			if cstr(""&dono2) = cstr(""&rsCallCombo("do_no")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("do_no")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>