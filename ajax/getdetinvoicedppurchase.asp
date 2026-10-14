<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
linenoinvoicedp = Request("linenoinvoicedp")
invoicedpno = Request("invoicedpno")
fcba = Request("fcba")
result = ""

q = " SELECT lineno, dp_aftervat  FROM purchaseinvoicedp WHERE fcba = '"& fcba &"' AND invoicedpno = '"& invoicedpno &"' AND lineno = '"& linenoinvoicedp &"'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	dp_aftervat = rsrecurtrans("dp_aftervat")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """dp_aftervat"" : """ & dp_aftervat & """"
response.write "}"

%>