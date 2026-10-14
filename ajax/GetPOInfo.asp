<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
pocode = Request("pocode")
fcba = Request("fcba")

arr_json_eof = true
' q = " SELECT hdr.podate" & _
	' "        , hdr.suppliercode" & _
	' "        , hdr.dp" & _ 
	' "        , NVL(hdr.ppn,0) ppn" & _
	' "        , NVL(hdr.totalppn,0) totalppn" & _
	' "        , supp.fcname suppliername" & _
	' "        , hdr.totalcharge" & _
	' " FROM purchaseorder hdr" & _
	' " LEFT JOIN supplier supp ON (hdr.suppliercode = supp.fccode AND hdr.fcba = supp.fcba)" & _
	' " WHERE 1 = 1"
' if len(trim(fcba)) > 0 then q = q & "       AND hdr.fcba = '"&fcba&"'"
' if len(trim(pocode)) > 0 then q = q & "       AND hdr.pocode = '"&pocode&"'"
' response.write q

' " (select sum(totalamount) dp from invoicedphdr where pocode= '"& pocode &"') " & _
' 	" - (select amount from paymentvoucherdetail where invoicecode = (select invoicecode from invoicedphdr where pocode= '"& pocode &"' order by invoicedate desc FETCH FIRST 1 ROW ONLY)) dp, " & _

q = " SELECT po.podate," & _
	"        po.suppliercode," & _
	" (select sum(totalamount) dp from invoicedphdr where pocode= '"& pocode &"') dp, " & _
	"        NVL (po.ppn, 0) ppn," & _
	"        NVL (po.totalppn, 0) totalppn," & _
	"        supp.fcname suppliername," & _
	"        sisa_otch_nonstock AS othercharge,otch_stock_po" & _
	"   FROM iplasprod.purchaseorder po" & _
	"        LEFT JOIN" & _
	"        (SELECT po.fcba," & _
	"                po.pocode," & _
	"                (NVL (po.otch_nonstock_po, 0) - NVL (inv.otch_nonstock_inv, 0))" & _
	"                   AS sisa_otch_nonstock, NVL (po.otch_stock_po, 0) AS otch_stock_po " & _
	"           FROM (  SELECT pocode, SUM (othercharge) AS otch_nonstock_inv" & _
	"                     FROM invoice" & _
	"                 GROUP BY pocode) inv" & _
	"                RIGHT JOIN" & _
	"                (  SELECT fcba, pocode, SUM (othercharge) AS otch_nonstock_po, SUM(otherchargestock) AS otch_stock_po" & _
	"                     FROM purchaseorderdetails" & _
	"                 GROUP BY fcba, pocode) po" & _
	"                   ON (inv.pocode = po.pocode) " & _
	"        ) inv" & _
	"           ON (po.fcba = inv.fcba AND po.pocode = inv.pocode)" & _
	"        LEFT JOIN supplier supp" & _
	"           ON (po.suppliercode = supp.fccode AND po.fcba = supp.fcba)" & _
	"  WHERE po.fcba = '"& fcba &"' AND po.pocode = '"& pocode &"'"
 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	podate= rsrecurtrans("podate")
	suppliercode= rsrecurtrans("suppliercode")
	suppliername= rsrecurtrans("suppliername")
	dp= rsrecurtrans("dp")
	ppn= rsrecurtrans("ppn")
	totalppn= rsrecurtrans("totalppn")
	othercharge= rsrecurtrans("othercharge") 
	otherchargestock= rsrecurtrans("otch_stock_po") 
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """podate"" : """ & podate & ""","
response.write """dp"" : """ & dp & ""","
response.write """suppliercode"" : """ & suppliercode & ""","
response.write """suppliername"" : """ & suppliername & ""","
response.write """ppn"" : """ & ppn & ""","
response.write """totalppn"" : """ & totalppn & ""","
response.write """othercharge"" : """ & othercharge & ""","
response.write """otherchargestock"" : """ & otherchargestock & """"
response.write "}"

%>