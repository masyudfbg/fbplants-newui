<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
orderdate = Request("orderdate")
documentno = Request("documentno")
lineby = Request("lineby")
itemcode = Request("itemcode")
storecode = Request("storecode")

arr_json_eof = true
q = " SELECT COUNT(1) total FROM (SELECT rn.receivenotecode, rnd.purchaseitemcode, rn.rndate, rnd.purchaselinenokey, (rnd.amountreceived / rnd.quantity) price" & _
	"   FROM receivenotedetail rnd" & _
	"        INNER JOIN receivenote rn" & _
	"           ON (rn.receivenotecode = rnd.receivenotecode AND rn.fcba = rnd.fcba)" & _
	"  WHERE     rn.fcba = '"& session("fcba") &"'" & _
	"        AND rn.receivenotecode = '"& documentno &"'" & _ 
	"        AND rnd.locationcode = '"& storecode &"'" & _
	"        AND rnd.purchaseitemcode = '"& itemcode &"'" & _ 
	"        AND rnd.purchaselinenokey = '"& lineby &"')" 


' response.write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr
if Not rsrecurtrans.eof then
	total= cdbl(to_nol(rsrecurtrans("total")))
end if

If total > 0 then
	q = " SELECT rn.receivenotecode, rnd.purchaseitemcode, rn.rndate, rnd.purchaselinenokey, (rnd.amountreceived / rnd.quantity) price" & _
		"   FROM receivenotedetail rnd" & _
		"        INNER JOIN receivenote rn" & _
		"           ON (rn.receivenotecode = rnd.receivenotecode AND rn.fcba = rnd.fcba)" & _
		"  WHERE     rn.fcba = '"& session("fcba") &"'" & _
		"        AND rn.receivenotecode = '"& documentno &"'" & _ 
		"        AND rnd.purchaseitemcode = '"& itemcode &"'" & _
		"        AND rnd.locationcode = '"& storecode &"'" & _ 
		"        AND rnd.purchaselinenokey = '"& lineby &"'" 


	' response.write q
	set rsrecurtrans = server.CreateObject("Adodb.recordset")
	rsrecurtrans.Open q, connstr
	if Not rsrecurtrans.eof then
		price = cdbl(to_nol(rsrecurtrans("price")))
	end if
	rsrecurtrans.Close : set rsrecurtrans = Nothing

Else

	q = " SELECT rch.orderdate," & _
		"        rct.orderno," & _
		"        rct.stockcode," & _
		"        rct.pricereceived as price," & _
		"        rct.linenosource as lineby" & _
		"   FROM receivedtaskdtl_tab rct" & _
		"        INNER JOIN receivedtask_tab rch" & _
		"           ON (rct.orderno = rch.orderno AND rct.fcba = rch.fcba)" & _
		"  WHERE     rct.fcba = '"& session("fcba") &"'" & _ 
		"        AND rct.orderno = '"& documentno &"'" & _
		"        AND rct.stockcode = '"& itemcode &"'" & _
		"        AND rct.locationcode = '"& storecode &"'" & _ 
		"        AND rct.linenosource = '"& lineby &"'"   


	' response.write q
	set rsrecurtrans = server.CreateObject("Adodb.recordset")
	rsrecurtrans.Open q, connstr

	if Not rsrecurtrans.eof then
		price = cdbl(to_nol(rsrecurtrans("price")))
	end if
	rsrecurtrans.Close : set rsrecurtrans = Nothing

End If


response.write "{"
response.write """price"" : """ & price & """"
response.write "}"

%>