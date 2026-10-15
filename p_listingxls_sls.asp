<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>Reporting - Listing AP</title>
</head>
<%
Dim tglawalsls,tglakhirsls,v_company,v_fcba,v_fctype,filter
tglawalsls      = trim(REQUEST("tglawalsls"))
tglakhirsls      = trim(REQUEST("tglakhirsls"))
v_fcba       = trim(REQUEST("pfcba"))

%>

<script src="/distexcel/tableToExcel.js"></script>
<script type="text/javascript">

	function getExport()
	{
		var tglawalsls = '<%= tglawalsls %>';
		var tglakhirsls = '<%=tglakhirsls %>';
		var businessunitar 	= '<%= v_fcba%>';
		var nm = 'SalesInvoice_' + businessunitar + '_'+tglawalsls+tglakhirsls+".xlsx";  
	
		TableToExcel.convert(document.getElementById("product-table"), {
            name: nm,
            sheet: {
            name: "SalesInvoice"
            }
          });
	}
</script>	
<script language="JavaScript" type="text/javascript" >
	document.getElementById('tableview').style.display = "none";
</script>
<%

	q =   " SELECT sih.fcba,sih.invoice_by,sih.invoicecode,sih.transdate, " & _ 
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
    " where sih.fcba='" & v_fcba & "' AND sih.TRANSDATE BETWEEN to_date('" & tglawalsls & "','mm/dd/yyyy') and to_date('" & tglakhirsls & "','mm/dd/yyyy') " & _
    " GROUP BY sih.fcba,sih.invoice_by,sih.invoicecode, " & _
    " sih.transdate, sih.customercode, customer_api.get_description (sih.customercode, sih.fcba) , " & _
    " sah.productcode, sih.vat, sih.remarks,sih.salescontractno, sih.VALID_FROM, " & _
    " sih.valid_until, sih.no_faktur,sih.expire_date,sih.fakturdate,sih.invoicedpno,j.postingflg"  
		
		' response.write q
Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open q, connstr
%>

<body>
    <span id="loadingjudul1">Proses export data...</span>
    <div id="tableview" style="display:none;">
   	<table border="0" width="100%" id="product-table" >
			<thead>
				<tr>
					<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">SALES INVOICE LISTING  Unit : <%=v_fcba%> Periode : <%=tglawalsls%> / <%=tglakhirsls%></th>		
				</tr>
				
		 <th style="color:#fff;">Invoice Code</th>
		 <th style="color:#fff;">Date</th> 
		 <th style="color:#fff;">Sales Contract</th> 
		 <th style="color:#fff;">Product</th> 
		 <th style="color:#fff;">Customer</th> 
		 <th style="color:#fff;">Qty</th> 
		 <th style="color:#fff;">Amount Before VAT</th> 
		 <th style="color:#fff;">Amount VAT</th> 
		 <th style="color:#fff;">Grand Total</th> 
		 <th style="color:#fff;">Posting</th> 
		
	</tr>
		</thead>
<tbody>
	<%
		While Not rsCat.EOF
		postingflg	= rscat("postingflg")
		if postingflg="1" then
				postingflg = "YA"
		else
				postingflg = "TIDAK"
		end if
	%>
		<tr>			
			<td align="Left" ><%=rscat("invoicecode")%></td>			
			<td align="Left" ><%=rscat("transdate")%></td>			
			<td align="Left" ><%=rscat("salescontractno")%></td>				
			<td align="Left" ><%=rscat("productcode")%></td>				
			<td align="Right" ><%=FbpNum2(rscat("customername"))%></td>		
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("quantity"))%></td>		
			<td data-t="n" align="Right" ><%=rscat("amount_beforevat")%></td>
			<td data-t="n" align="Right" ><%=rscat("amountvat")%></td>				
			<td data-t="n" align="Right" ><%=FbpNum2(rscat("grandtotal"))%></td>				
			<td align="Right" ><%=postingflg%></td>				
			
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null
	Response.Write("<script language='javascript'> getExport();</script>")
%>
</tbody>
</table>
</div>

</BODY>
</HTML>
