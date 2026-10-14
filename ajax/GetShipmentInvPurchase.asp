<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
salesshipmentno = Request("salesshipmentno")
fcba = Request("fcba")
result = ""

q =	  " select sd.productcode , sd.prices as unitprice, sd.purchasecontractno, ssd.acceptquantity as quantity, sd.ppn as vat, sd.productdescription as product" & _
	  " from ( select sad.productcode, sad.prices, sad.ppn, sah.purchasecontractno, sah.fcba, ssh.purchaseshipmentno, pd.productdescription from purchaseagreementdet sad" & _
	  "    left join purchaseagreementhdr sah ON (sad.purchasecontractno = sah.purchasecontractno and sad.fcba = sah.fcba)" & _
	  "    left join purchaseshipmenthdr ssh ON(sad.salescontractno = ssh.purchasecontractno and sad.fcba = ssh.fcba )" & _
	  "    left join product pd ON (sad.productcode = pd.productcode AND sad.fcba = pd.fcba)" & _
	  "    )sd" & _
	  " left join purchaseshipmentdtl ssd ON (ssd.purchaseshipmentno = sd.purchaseshipmentno and ssd.fcba = sd.fcba)" & _
	  " WHERE sd.fcba = '" & fcba & "'" & _
	  " AND sd.purchaseshipmentno = '" & salesshipmentno & "'"  


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