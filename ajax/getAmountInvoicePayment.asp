<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
documentno = Request("documentno")
locationcode = Request("locationcode")
jobcode = Request("jobcode")
result = ""
amountinvoice = 0

q =	" SELECT pr.parcode," & _
	"       pr.parname," & _
	"       prd.fccode," & _
	"       prd.parheadcode," & _
	"       prd.fcname," & _
	"       ih.totalamount" & _
	"  FROM parameter pr" & _
	"       INNER JOIN parameterdetail prd" & _
	"          ON (pr.parcode = prd.parcode AND pr.fcba = prd.fcba AND pr.fccode = prd.parheadcode)" & _
	"       LEFT JOIN invoice ih ON (pr.fcba = ih.fcba)" & _
	" WHERE     prd.fccode = '"& jobcode &"'" & _
	"       AND ih.invoice_by = '"& session ("fcba") &"'" & _
	"       AND ih.suppliercode = '"& locationcode &"'" & _
	"       AND ih.invoicecode = '"& documentno &"'"


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	JobPurchaseInvoice = rsrecurtrans("fccode")
	totalinvoice = Cdbl(to_nol(rsrecurtrans("totalamount"))) 
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

q =	" SELECT pr.parcode," & _
	"       pr.parname," & _
	"       prd.fccode," & _
	"       prd.parheadcode," & _
	"       prd.fcname," & _
	"       ih.totalamount" & _
	"  FROM parameter pr" & _
	"       INNER JOIN parameterdetail prd" & _
	"          ON (pr.parcode = prd.parcode AND pr.fcba = prd.fcba AND pr.fccode = prd.parheadcode)" & _
	"       LEFT JOIN invoicesto ih ON (pr.fcba = ih.fcba)" & _
	" WHERE     prd.fccode = '"& jobcode &"'" & _
	"       AND ih.invoice_by = '"& session ("fcba") &"'" & _
	"       AND ih.transportercode = '"& locationcode &"'" & _
	"       AND ih.invoicecode = '"& documentno &"'"


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	JobSTOInvoice = rsrecurtrans("fccode")
	totalstoinvoice = Cdbl(to_nol(rsrecurtrans("totalamount"))) 
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

' q =	" SELECT pr.parcode," & _
	' "       pr.parname," & _
	' "       prd.fccode," & _
	' "       prd.parheadcode," & _
	' "       prd.fcname," & _
	' "       ish.grandtotal" & _
	' "  FROM parameter pr" & _
	' "       INNER JOIN parameterdetail prd" & _
	' "          ON (pr.parcode = prd.parcode AND pr.fcba = prd.fcba AND pr.fccode = prd.parheadcode)" & _
	' "       LEFT JOIN invoiceservicehdr ish ON (pr.fcba = ish.fcba)" & _
	' "       INNER JOIN invoiceservicedtl isd" & _
	' "          ON (ish.invoiceno = isd.invoiceno AND ish.fcba = isd.fcba)" & _
	' " WHERE     prd.fccode = '"& jobcode &"'" & _
	' "       AND ish.invoice_by = '"& session("fcba") &"'" & _
	' "       AND ish.suppliername = '"& locationcode &"'" & _
	' "       AND ish.invoiceno = '"& documentno &"'"  

' response.Write q
' set rsrecurtrans = server.CreateObject("Adodb.recordset")
' rsrecurtrans.Open q, connstr

' if Not rsrecurtrans.eof then
	' JobPurchaseInvoice = rsrecurtrans("fccode")
	' total = Cdbl(to_nol(rsrecurtrans("grandtotal"))) 
' end if
' rsrecurtrans.Close
' set rsrecurtrans = Nothing

q =	" SELECT pr.parcode," & _
	"       pr.parname," & _
	"       prd.fccode," & _
	"       prd.parheadcode," & _
	"       prd.fcname," & _
	"       tc.transporter," & _
	"       tid.amountafterloss," & _
	"       ti.vat" & _
	"  FROM parameter pr" & _
	"       INNER JOIN parameterdetail prd" & _
	"          ON (pr.parcode = prd.parcode AND pr.fcba = prd.fcba AND pr.fccode = prd.parheadcode)" & _
	"       LEFT JOIN transportercontract tc ON (pr.fcba = tc.fcba)" & _
	"       INNER JOIN transporterinvoice ti" & _
	"          ON (ti.transportercontractno = tc.transportercontractno AND ti.fcba = tc.fcba)" & _
	"       INNER JOIN transporterinvoicedtl tid ON (tid.invoicecode = ti.invoicecode AND tid.fcba = ti.fcba AND tid.lineno = tc.lineno)" & _
	" WHERE     prd.fccode = '"& jobcode &"'" & _
	"       AND ti.invoice_by = '"& session("fcba") &"'" & _
	"       AND tc.transporter = '"& locationcode &"'" & _
	"       AND ti.invoicecode = '"& documentno &"'"


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	JobTransporterInvoice = rsrecurtrans("fccode")
	vat = Cdbl(to_nol(rsrecurtrans("vat")))  
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

amount = 0
strSuppyingDeclaration	=	" SELECT NVL(amountafterloss,0) amount FROM TRANSPORTERINVOICEDTL" & _
							" WHERE invoicecode = '" & documentno & "'" 
' response.Write strSuppyingDeclaration
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	while Not rsSupplyingDeclaration.Eof
		amount = amount + cdbl(to_nol(rsSupplyingDeclaration("amount")))
		rsSupplyingDeclaration.movenext
	wend
End if
rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing

amountvat = amount * cdbl(to_nol(vat))/100
totaltransporterinvoice = amount + amountvat

q =	"SELECT pr.parcode," & _
	"       pr.parname," & _
	"       prd.fccode," & _
	"       prd.parheadcode," & _
	"       prd.fcname," & _
	"       si.grandtotal" & _
	"  FROM parameter pr" & _
	"       INNER JOIN parameterdetail prd" & _
	"          ON (    pr.parcode = prd.parcode" & _
	"              AND pr.fcba = prd.fcba" & _
	"              AND pr.fccode = prd.parheadcode)" & _
	"       LEFT JOIN salesinvoicedtl sidt ON (pr.fcba = sidt.fcba)" & _
    " 		LEFT JOIN salesinvoicehdr si ON (si.fcba = sidt.fcba and si.invoicecode = sidt.invoicecode)" & _
	" WHERE     prd.fccode = '"& jobcode &"'" & _
	"       AND si.invoice_by = '"& session("fcba") &"'" & _
	"       AND si.customercode = '"& locationcode &"'" & _
	"       AND si.invoicecode = '"& documentno &"'" 

' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	JobSalesInvoice = rsrecurtrans("fccode")
	totalsalesinvoice = Cdbl(to_nol(rsrecurtrans("grandtotal")))  
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

q =	" SELECT pr.parcode," & _
	"       pr.parname," & _
	"       prd.fccode," & _
	"       prd.parheadcode," & _
	"       prd.fcname," & _
	"       cih.grandtotal_aftertax" & _
	"  FROM parameter pr" & _
	"       INNER JOIN parameterdetail prd" & _
	"          ON (    pr.parcode = prd.parcode" & _
	"              AND pr.fcba = prd.fcba" & _
	"              AND pr.fccode = prd.parheadcode)" & _
	"       LEFT JOIN contractorinvoicehdr cih ON (pr.fcba = cih.fcba)" & _
	" WHERE     prd.fccode = '"& jobcode &"'" & _
	" 		AND cih.invoice_by = '"& session("fcba") &"'"& _
	"       AND cih.contractorcode = '"& locationcode &"'" & _
	"       AND cih.invoicecode = '"& documentno &"'" 


' response.Write q
set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	JobContractorInvoice = rsrecurtrans("fccode")
	totalcontractorinvoice = Cdbl(to_nol(rsrecurtrans("grandtotal_aftertax")))  
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

	if JobPurchaseInvoice = "2110101" then
		amountinvoice = totalinvoice 
	Elseif JobSTOInvoice = "2110201" then
		amountinvoice = totalstoinvoice
	Elseif JobTransporterInvoice = "2110201" then
		amountinvoice = totaltransporterinvoice
	Elseif JobContractorInvoice = "2110201" then
		amountinvoice = totalcontractorinvoice
	Else
		amountinvoice = totalsalesinvoice
	End If

		response.write "{"
		response.write """amountinvoice"" : """ & amountinvoice & """"
		' response.write """unitprice"" : """ & unitprice & ""","
		' response.write """productcode"" : """ & productcode & " : " & productdescription & """"
		response.write "}"		

%>