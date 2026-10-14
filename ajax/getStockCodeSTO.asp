<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
linenokey = Request("linenokey")
fccode = Request("fccode")
result = ""

' q = " select sp.fccode as suppliercode, sp.fcname as suppliername, bu.fccode , bu.fcname, bu.fccompanycode " & _
' 	" from supplier sp Inner join businessunit bu ON (sp.fcba = bu.fccode)" & _
' 	" where bu.fccode = '"& fccode &"' and sp.fccode='"& suppliercode &"'"

q = " select pod.linenokey, pod.quantity, st.fccode, st.fcname || ' - ' || pod.otheritemdesc fcname, st.fcba" & _
	" from purchaseorderdetails pod " & _
	" inner join stock st on (st.fccode = pod.itemcode and st.fcba = pod.fcba)" & _
	" where pod.fcba = '"& fcba &"' and pod.linenokey = '"& linenokey &"' and pod.pocode = '"& fccode &"'"

'response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	itemcode = rsrecurtrans("fccode")
	stockname = rsrecurtrans("fcname")
	quantity = rsrecurtrans("quantity")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """quantity"" : """ & quantity & ""","
response.write """itemcode"" : """ & itemcode & ""","
response.write """stockname"" : """ & stockname & """"
response.write "}"

%>