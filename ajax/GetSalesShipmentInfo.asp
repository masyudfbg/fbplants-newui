<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salesshipmentno = Request("salesshipmentno")
wbt_chitno = Request("wbt_chitno")
fcba = Request("fcba")

arr_json_eof = true
q = " SELECT dtl.salesshipmentno" & _
	"        , dtl.wbt_chitno" & _
	"        , dtl.postingdate" & _
	"        , dtl.transporter" & _
	"        , dtl.vehicleno" & _
	"        , dtl.quantity" & _
	"        , dtl.acceptquantity" & _
	"        , dtl.customer_chitno" & _
	"        , dtl.acceptquantitybruto" & _
	"        , hdr.creationdate" & _
	"        , hdr.salescontractno" & _
	"        , hdr.deliverydate" & _
	"        , hdr.deliveryorderno" & _
	"        , hdr.deliveryorderdate" & _
	"        , hdr.quantity_do" & _
	"        , hdr.shipmenttype" & _
	"        , hdr.destinationcountry" & _
	"        , hdr.address" & _
	"        , hdr.remarks" & _
	"        , dtl.fcba" & _
	"		 , do.transportationcost" & _
	"		 , do.losstolerance" & _
	" FROM salesshipmentdtl dtl" & _
	" INNER JOIN salesshipmenthdr hdr ON (dtl.fcba = hdr.fcba AND dtl.salesshipmentno = hdr.salesshipmentno)" & _
	" LEFT JOIN deliveryorder do ON (do.deliveryorderno = hdr.deliveryorderno AND do.fcba = hdr.fcba)"& _
	" WHERE 1 = 1"
if len(trim(fcba)) > 0 then q = q & "       AND dtl.fcba = '"&fcba&"'"
if len(trim(salesshipmentno)) > 0 then q = q & "       AND dtl.salesshipmentno = '"&salesshipmentno&"'"
if len(trim(wbt_chitno)) > 0 then q = q & "       AND dtl.wbt_chitno = '"&wbt_chitno&"'"

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
		response.write """salesshipmentno"" : """ & arr_json(0,xyyy) & ""","
		response.write """wbt_chitno"" : """ & arr_json(1,xyyy) & ""","
		response.write """postingdate"" : """ & arr_json(2,xyyy) & ""","
		response.write """transporter"" : """ & arr_json(3,xyyy) & ""","
		response.write """vehicleno"" : """ & arr_json(4,xyyy) & ""","
		response.write """quantity"" : """ & arr_json(5,xyyy) & ""","
		response.write """acceptquantity"" : """ & arr_json(6,xyyy) & ""","
		response.write """customer_chitno"" : """ & arr_json(7,xyyy) & ""","
		response.write """acceptquantitybruto"" : """ & arr_json(8,xyyy) & ""","
		response.write """creationdate"" : """ & arr_json(9,xyyy) & ""","
		response.write """salescontractno"" : """ & arr_json(10,xyyy) & ""","
		response.write """deliverydate"" : """ & arr_json(11,xyyy) & ""","
		response.write """deliveryorderno"" : """ & arr_json(12,xyyy) & ""","
		response.write """deliveryorderdate"" : """ & arr_json(13,xyyy) & ""","
		response.write """quantity_do"" : """ & arr_json(14,xyyy) & ""","
		response.write """shipmenttype"" : """ & arr_json(15,xyyy) & ""","
		response.write """destinationcountry"" : """ & arr_json(16,xyyy) & ""","
		response.write """address"" : """ & arr_json(17,xyyy) & ""","
		response.write """remarks"" : """ & arr_json(18,xyyy) & ""","
		response.write """transportationcost"" : """ & arr_json(20,xyyy) & ""","
		response.write """losstolerance"" : """ & arr_json(21,xyyy) & """"
	next
	response.write "}"
else
	response.write "{"
	response.write """salesshipmentno"" : """","
	response.write """customer_chitno"" : """","
	response.write """postingdate"" : """","
	response.write """transporter"" : """","
	response.write """vehicleno"" : """","
	response.write """quantity"" : """","
	response.write """acceptquantity"" : """","
	response.write """customer_chitno"" : """","
	response.write """acceptquantitybruto"" : """","
	response.write """creationdate"" : """","
	response.write """salescontractno"" : """","
	response.write """deliverydate"" : """","
	response.write """deliveryorderno"" : """","
	response.write """deliveryorderdate"" : """","
	response.write """quantity_do"" : """","
	response.write """shipmenttype"" : """","
	response.write """destinationcountry"" : """","
	response.write """address"" : """","
	response.write """remarks"" : """","
	response.write """transportationcost"" : """","
	response.write """losstolerance"" : """""
	response.write "}"
end if

%>