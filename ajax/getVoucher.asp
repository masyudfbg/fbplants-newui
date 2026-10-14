<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
xCode = Request("xCode")
datecreated = Request("datecreated")
result = ""

q = "  SELECT a.vouchercode," & _
"         a.bankcode," & _
"         a.totalamount," & _
"         a.datecreated," & _
"         b.sourcetype," & _
"         b.locationtype," & _
"         b.locationcode," & _
"         MAX (b.lineno)" & _
"    FROM paymentvoucher a" & _
"         INNER JOIN paymentvoucherdetail b" & _
"            ON (a.vouchercode = b.vouchercode AND a.fcba = b.fcba)" & _
"   WHERE     a.bankcode = '"& xCode &"'" & _
"         AND a.fcba = '"& session("fcba") &"'" & _
"         and to_char(datecreated,'MM/DD/YYYY') = '"& datecreated & "'" & _
"         GROUP BY a.vouchercode," & _
"         a.bankcode," & _
"         a.totalamount," & _
"         a.datecreated," & _
"         b.sourcetype," & _
"         b.locationtype," & _
"         b.locationcode"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	' suppliername = rsrecurtrans("fcname")
	vouchercode = rsrecurtrans("vouchercode")
else
	vouchercode = ""
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
' response.write """suppliername"" : """ & suppliername & ""","
response.write """vouchercode"" : """ & vouchercode & """"
response.write "}"

%>