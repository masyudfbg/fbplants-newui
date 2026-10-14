<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
orderdate = Request("orderdate")
documentno = Request("documentno")
lineby = Request("lineby")
itemcode = Request("itemcode")

arr_json_eof = true
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
	"        AND rct.linenosource = '"& lineby &"'"   


' response.write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	arr_json = rsrecurtrans.getrows()
	arr_json_eof = false
end if
rsrecurtrans.Close : set rsrecurtrans = Nothing

if not arr_json_eof then
	response.write "{"
	for xyyy=0 to ubound(arr_json,2)
		' response.write """podate"" : """ & arr_json(0,xyyy) & ""","
		' response.write """suppliercode"" : """ & arr_json(1,xyyy) & ""","
		' response.write """dp"" : """ & arr_json(2,xyyy) & ""","
		' response.write """ppn"" : """ & arr_json(3,xyyy) & ""","
		response.write """price"" : """ & arr_json(3,xyyy) & """"
	next
	response.write "}"
else
	response.write "{"
	' response.write """podate"" : """","
	' response.write """suppliercode"" : """","
	' response.write """dp"" : """","
	' response.write """ppn"" : ""0"","
	response.write """price"" : """""
	response.write "}"
end if

%>