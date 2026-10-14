<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
orderdate = Request("orderdate")
documentno = Request("documentno")
lineby = Request("lineby")
itemcode = Request("itemcode")
storecode = Request("storecode")
result = ""

q = " SELECT COUNT (1) total" & _
	"   FROM (SELECT rnd.purchaselinenokey AS lineby" & _
	"           FROM receivenotedetail rnd" & _
	"                INNER JOIN receivenote rn" & _
	"                   ON (    rn.receivenotecode = rnd.receivenotecode" & _
	"                       AND rn.fcba = rnd.fcba)" & _
	"          WHERE     rn.fcba = '"& session("fcba") &"'" & _
	"                AND rn.receivenotecode = '"& documentno &"'" & _
	"        		 AND rnd.locationcode = '"& storecode &"'" & _
	"        		 AND rnd.purchaseitemcode = '"& itemcode &"'" & _ 
	"                AND rn.rndate <= TO_DATE ('"& orderdate &"', 'mm/dd/yyyy'))" 


' response.write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	total= cdbl(to_nol(rsrecurtrans("total")))
end if

If total > 0 then

	q = " SELECT rnd.purchaselinenokey as lineby" & _
		"   FROM receivenotedetail rnd" & _
		"        INNER JOIN receivenote rn" & _
		"           ON (rn.receivenotecode = rnd.receivenotecode AND rn.fcba = rnd.fcba)" & _
		"  WHERE     rn.fcba = '"& session("fcba") &"'" & _
		"        AND rn.receivenotecode = '"& documentno &"'" & _
		"        AND rnd.locationcode = '"& storecode &"'" & _
		"        AND rnd.purchaseitemcode = '"& itemcode &"'" & _ 
		"        AND rndate <= TO_DATE ('"& orderdate &"', 'mm/dd/yyyy')" & _
		"	ORDER BY rnd.purchaselinenokey"

	' response.Write q
	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.Open q, connstr
Else
	q = " SELECT rch.orderdate," & _
		"        rct.orderno," & _
		"        rct.stockcode," & _
		"        rct.pricereceived," & _
		"        rct.linenosource as lineby" & _
		"   FROM receivedtaskdtl_tab rct" & _
		"        INNER JOIN receivedtask_tab rch" & _
		"           ON (rct.orderno = rch.orderno AND rct.fcba = rch.fcba)" & _
		"  WHERE     rct.fcba = '"& session("fcba") &"'" & _
		"        AND rch.orderdate <= TO_DATE ('"& orderdate &"', 'mm/dd/yyyy') " & _
		"        AND rct.orderno = '"& documentno &"'"& _
		"        AND rct.locationcode = '"& storecode &"'" & _
		"        AND rct.stockcode = '"& itemcode &"'" & _  
		"		ORDER BY rct.orderno, rct.linenosource"  

	' response.Write q
	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.Open q, connstr
End If

	if not rsCallCombo.eof then
		response.write "<option value=''></option>"
		Do While Not rsCallCombo.eof
		 html = html & "<option value='" & rsCallCombo("lineby") & "'" 
		 if cdbl(to_nol(Request("lineby"))) = cdbl(to_nol(rsCallCombo("lineby"))) then html = html &  " selected"
		 html = html &  " >"
		 html = html & rsCallCombo("lineby")

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