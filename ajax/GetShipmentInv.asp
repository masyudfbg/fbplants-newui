<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salesshipmentno = Request("salesshipmentno")
fcba = Request("fcba")
result = ""

q =	  " select sd.productcode , sd.salesprices as unitprice, sd.salescontractno, ssd.acceptquantity as quantity, sd.vat as vat, sd.productdescription as product" & _
	  " from ( select sad.productcode, sad.salesprices, sad.vat, sah.salescontractno, sah.fcba, ssh.salesshipmentno, pd.productdescription from salesagreementdet sad" & _
	  "    left join salesagreementhdr sah ON (sad.salescontractno = sah.salescontractno and sad.fcba = sah.fcba)" & _
	  "    left join salesshipmenthdr ssh ON(sad.salescontractno = ssh.salescontractno and sad.fcba = ssh.fcba )" & _
	  "    left join product pd ON (sad.productcode = pd.productcode AND sad.fcba = pd.fcba)" & _
	  "    )sd" & _
	  " left join salesshipmentdtl ssd ON (ssd.salesshipmentno = sd.salesshipmentno and ssd.fcba = sd.fcba)" & _
	  " WHERE sd.fcba = '" & fcba & "'" & _
	  " AND sd.salesshipmentno = '" & salesshipmentno & "'"  


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	quantity = rsrecurtrans("quantity")
	unitprice = rsrecurtrans("unitprice")
	product = rsrecurtrans("product")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing


response.write "{"
response.write """quantity"" : """ & quantity & ""","
response.write """unitprice"" : """ & unitprice & ""","
response.write """product"" : """ & product & """"
response.write "}"
%>