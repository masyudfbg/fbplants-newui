<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
jnsinvoice = Request("jnsinvoice")
fcba = Session("fcBA")
selectnoinv =Request("selectnoinv")

Dim nominal

if jnsinvoice = "1" then
	q = " select h.totalamount - h.payment as grandtotal " & _
	" from invoice h " & _
	" where h.invoicecode='" & selectnoinv & "' and invoice_by='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 
	if Not rsnominal.eof then
		nominal= rsnominal("grandtotal")
		
	end if

	rsnominal.Close
	set rsnominal = Nothing

	response.write "{"
	response.write """nominal"" : """ & nominal & """"
	response.write "}"
end if

if jnsinvoice = "2" then

	q = " select (h.grandtotal+h.totalpph)-h.payment as grandtotal" & _
	" from invoiceservicehdr h " & _
	" where h.invoiceno='" & selectnoinv & "' and invoice_by='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 'response.Write q


if Not rsnominal.eof then
	nominal= rsnominal("grandtotal")
	
end if

rsnominal.Close
set rsnominal = Nothing

response.write "{"
response.write """nominal"" : """ & nominal & """"
response.write "}"

end if

if jnsinvoice = "3" then

	q = "select h.grandtotal_aftertax-h.payment as grandtotal " & _
 " from contractorinvoicehdr h  " & _
 " where h.invoicecode='" & selectnoinv & "' and invoice_by='" & fcba & "' " & _
 " union all " & _
 " select h.dp-h.payment as grandtotal  " & _
 " from CONTRACTORINVOICEDP h " & _
 " where h.invoicecode='" & selectnoinv & "' and invoice_by='" & fcba & "' " 
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 'response.Write q


if Not rsnominal.eof then
	nominal= rsnominal("grandtotal")
	
end if

rsnominal.Close
set rsnominal = Nothing

response.write "{"
response.write """nominal"" : """ & nominal & """"
response.write "}"

end if
if jnsinvoice = "4" then

	q = " select h.totalamount-h.payment as grandtotal" & _
	" from invoicedphdr h " & _
	" where h.invoicecode='" & selectnoinv & "' and invoice_by='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 'response.Write q


if Not rsnominal.eof then
	nominal= rsnominal("grandtotal")
	
end if

rsnominal.Close
set rsnominal = Nothing

response.write "{"
response.write """nominal"" : """ & nominal & """"
response.write "}"

end if

if jnsinvoice = "5" then

	q = " select h.totalamount-h.payment as grandtotal" & _
	" from invoicesto h " & _
	" where h.invoicecode='" & selectnoinv & "' and invoice_by='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 'response.Write q


if Not rsnominal.eof then
	nominal= rsnominal("grandtotal")
	
end if

rsnominal.Close
set rsnominal = Nothing

response.write "{"
response.write """nominal"" : """ & nominal & """"
response.write "}"

end if

if jnsinvoice = "6" then

	q = " select h.grandtotal-h.payment grandtotal" & _
	" from invoiceservicehdr h " & _
	" where REPLACE(h.invoiceno,chr(32),'')='" & selectnoinv & "' and invoice_by='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
'response.Write q


if Not rsnominal.eof then
	nominal= rsnominal("grandtotal")
	
end if

rsnominal.Close
set rsnominal = Nothing

response.write "{"
response.write """nominal"" : """ & nominal & """"
response.write "}"

end if

if jnsinvoice = "7" then

	q = " select h.AMOUNTAFTERPPH-h.payment grandtotal" & _
	" from TRANSPORTERINVOICE h " & _
	" where h.invoicecode='" & selectnoinv & "' and invoice_by='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 'response.Write q


if Not rsnominal.eof then
	nominal= rsnominal("grandtotal")
	
end if

rsnominal.Close
set rsnominal = Nothing

response.write "{"
response.write """nominal"" : """ & nominal & """"
response.write "}"

end if

if jnsinvoice = "8" then

	q = " select grandtotal-RECEIVEAMT grandtotal" & _
	" from salesinvoicehdr" & _
	" where invoicecode='" & selectnoinv & "' and invoice_by='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 'response.Write q


if Not rsnominal.eof then
	nominal= rsnominal("grandtotal")
	
end if

rsnominal.Close
set rsnominal = Nothing

response.write "{"
response.write """nominal"" : """ & nominal & """"
response.write "}"

end if

if jnsinvoice = "9" then
	q = " select h.grandtotal - h.payment as grandtotal " & _
	" from purchaseinvoicehdr h " & _
	" where h.invoicecode='" & selectnoinv & "' and invoice_by='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 
	if Not rsnominal.eof then
		nominal= rsnominal("grandtotal")
		
	end if

	rsnominal.Close
	set rsnominal = Nothing

	response.write "{"
	response.write """nominal"" : """ & nominal & """"
	response.write "}"
end if

if jnsinvoice = "10" then
	q = " select h.amount - h.payment as grandtotal " & _
	" from OPR_INV_PDO h " & _
	" where h.NOINVOICE='" & selectnoinv & "' and invoiceby='" & fcba & "'"
	set rsnominal = server.CreateObject("Adodb.recordset")
	rsnominal.Open q, connstr
 
	if Not rsnominal.eof then
		nominal= rsnominal("grandtotal")
		
	end if

	rsnominal.Close
	set rsnominal = Nothing

	response.write "{"
	response.write """nominal"" : """ & nominal & """"
	response.write "}"
end if

%>