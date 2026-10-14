<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
deliveryorderno = Request("deliveryorderno")
fcba = Request("fcba")
result = ""

html = html & "<select id='deliveryorderno' name='deliveryorderno' onchange='javascript:getselecteddo();javascript:getDOLine();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = "   SELECT DISTINCT sdtt.fcba," & _
	"          sdtt.salescontractno," & _
	"          sdtt.deliveryorderno," & _
	"          sdtt.acceptquantity," & _
	"          sadd.quantityaftertolerance,sadd.productcode" & _
	"     FROM (  SELECT sc.fcba," & _
	"                    sc.salescontractno," & _
	"                    do.deliveryorderno," & _
	"                    do.lineno," & _
	"                    NVL (SUM (sdt.acceptquantity), 0) AS acceptquantity" & _
	"               FROM salesshipmentdtl sdt" & _
	"                    INNER JOIN salesshipmenthdr ssh" & _
	"                       ON (    ssh.fcba = sdt.fcba" & _
	"                           AND ssh.salesshipmentno = sdt.salesshipmentno)" & _
	"                    RIGHT JOIN deliveryorder do" & _
	"                       ON (    do.fcba = ssh.fcba" & _
	"                           AND do.salescontractno = ssh.salescontractno)" & _
	"                    INNER JOIN salesagreementhdr sc" & _
	"                       ON (    sc.fcba = do.fcba" & _
	"                           AND sc.salescontractno = do.salescontractno)" & _
	"           GROUP BY sc.fcba, sc.salescontractno, do.deliveryorderno, do.lineno) sdtt" & _
	"          INNER JOIN" & _
	"          (  SELECT sad.fcba," & _
	"                    sad.salescontractno," & _
	"                    SUM (sad.quantityaftertolerance) quantityaftertolerance,sad.productcode " & _
	"               FROM salesagreementdet sad" & _
	"           GROUP BY sad.fcba, sad.salescontractno,sad.productcode) sadd" & _
	"             ON (    sadd.fcba = sdtt.fcba" & _
	"                 AND sadd.salescontractno = sdtt.salescontractno)" & _
	"    WHERE /*sdtt.acceptquantity < sadd.quantityaftertolerance" & _
	" 	 AND */ sdtt.fcba = '"& fcba &"'"& _
	" ORDER BY sdtt.fcba, sdtt.salescontractno, sdtt.deliveryorderno" 
  

	
	' "  SELECT deliveryorderno" & _
	' "    FROM (SELECT deliveryorderno" & _
	' "            FROM deliveryorder" & _
	' "           WHERE fcba = '"& fcba &"'" & _
	' "          MINUS" & _
	' "          SELECT deliveryorderno" & _
	' "            FROM salesshipmenthdr" & _
	' "           WHERE fcba = '"& fcba &"')" & _
	' " ORDER BY deliveryorderno" 

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