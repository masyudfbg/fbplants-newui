<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - SALES INVOICE PDF</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
 window.print();
</script>
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

dim tglawal,tglakhir,pilihbkm,section,posting

tglawal= trim(request("tglawal"))
tglakhir= trim(request("tglakhir"))
posting = Request("posting")



businessunit = Request("businessunitar")

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
	set rsview = connStr.Execute(q)
	if not rsview.eof then
	   fcbaname = rsview("fcName")
	   CompanyName=rsview("FCCOMPANYNAME")
	   MgrName = rsview("FCMANAGER")
	   KTUName = rsview("FCKTU")
	   CropName=rsview("FCCROP")
	   FcbaType=rsview("fcType")
	end if

Dim sqlview,rsview 
sqlview = "SELECT sih.fcba,sih.invoice_by,sih.invoicecode,sih.transdate, " & _ 
    " sih.customercode,customer_api.get_description (sih.customercode, sih.fcba) customername, " & _
    " sah.productcode,sum(sid.quantity) quantity,sih.vat,AVG(sih.grandtotal)grandtotal, " & _
    " sih.remarks,sih.salescontractno,AVG(sih.amount_beforevat) amount_beforevat, " & _
    " AVG(sih.amountvat)amountvat,sih.VALID_FROM,sih.valid_until, " & _
    " sih.no_faktur,sih.expire_date,sih.fakturdate,sih.invoicedpno,j.postingflg  " & _
    " FROM salesinvoicedtl sid " & _
    " INNER JOIN salesinvoicehdr sih ON (sih.invoicecode = sid.invoicecode AND sih.fcba = sid.fcba) " & _
    " INNER JOIN salesagreementdet sah " & _
    " ON ( sah.fcba = sih.fcba AND sah.salescontractno = sih.salescontractno)  " & _
	" LEFT JOIN BKMJOURNALHDR j on (j.DOCNO=sih.INVOICECODE and j.fcba=sih.fcba) " & _
    " where sih.fcba='" & businessunit & "' AND sih.TRANSDATE BETWEEN to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & tglakhir & "','mm/dd/yyyy') " & _
    " GROUP BY sih.fcba,sih.invoice_by,sih.invoicecode, " & _
    " sih.transdate, sih.customercode, customer_api.get_description (sih.customercode, sih.fcba) , " & _
    " sah.productcode, sih.vat, sih.remarks,sih.salescontractno, sih.VALID_FROM, " & _
    " sih.valid_until, sih.no_faktur,sih.expire_date,sih.fakturdate,sih.invoicedpno,j.postingflg"
set rsview = connStr.Execute(sqlview)


' response.Write(sqlview)
'response.Flush()
%>
	<body>
		<div id="wrapper">
        
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>LISTING SALES INVOICE </h2>
				</div>
			</div>
				</th>
			</tr>
				<!--<div class="printhide">-->
						<tr>
							<th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
						
						
						</tr> 
						<tr>
							
							<th>Invoice Code</th>
							<th>Date</th>
							<th>Sales Contract</th>
							<th>Product</th>
							<th>Customer</th>
							<th>Qty</th>
							<th>Amount Before VAT</th>	  
							<th>Amount VAT</th>
							<th>Grand Total</th>
							<th>Posting</th>
							
						</tr>
                         </thead>

						 <%
						 	dim ttlqty,ttlamtbvat,ttlamtvat,ttlgrand,ttlamount
							ttlqty = 0
							ttlamtbvat = 0
							ttlamtvat = 0
							ttlgrand = 0
							
							While Not rsview.EOF

									 if rsview("POSTINGFLG")= "0" then 
										postingflg= "TIDAK"
									elseif  rsview("POSTINGFLG")= "" then
										postingflg= "TIDAK"
									else 
										postingflg= "YA"
									end if

						 %>

                         <tbody>
						<tr>
														 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("invoicecode")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("transdate")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("salescontractno")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("productcode")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("customername")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsview("quantity"))%></td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsview("amount_beforevat"))%></td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsview("amountvat"))%> </td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsview("grandtotal"))%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=postingflg%> </td>
							
						</tr>
						
						</tbody>
						<%
							ttlqty = ttlqty +cdbl(rsview("quantity"))
							ttlamtbvat = ttlamtbvat+cdbl(rsview("amount_beforevat"))
							ttlamtvat = ttlamtvat+cdbl(rsview("amountvat"))
							ttlgrand = ttlgrand+cdbl(rsview("grandtotal"))
							
						rsview.MoveNext
						Wend
						rsview.close()
						%>
						<tr >
							<td colspan=6><strong>Total</strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlqty)),2)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlamtbvat)),2)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlamtvat)),2)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlgrand)),2)%></strong></td>
							
						</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	