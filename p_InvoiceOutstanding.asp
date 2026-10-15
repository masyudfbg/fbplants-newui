<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Invoice Outstanding</title>
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

' txtSQL = "SELECT DATA.invoicecode, DATA.invoicedate, DATA.suppliercode, supplier.fcname,"
' txtSQL = txtSQL + " DATA.PAID, DATA.totalinvoice, DATA.netbalance"
' txtSQL = txtSQL + " FROM (SELECT a.invoicecode, a.invoicedate, a.suppliercode, a.PAID,"
' txtSQL = txtSQL + " b.totalinvoice, (b.totalinvoice - a.PAID) netbalance"
' txtSQL = txtSQL + " FROM (SELECT invoicecode, invoicedate, suppliercode, PAID"
' txtSQL = txtSQL + " FROM invoice) a,"
' txtSQL = txtSQL + " (SELECT invoicecode, SUM (amount) totalinvoice"
' txtSQL = txtSQL + " FROM invoicedetail"
' txtSQL = txtSQL + " GROUP BY invoicecode) b"
' txtSQL = txtSQL + " WHERE a.invoicecode = b.invoicecode AND (b.totalinvoice - a.PAID) > 0) DATA,"
' txtSQL = txtSQL + "  supplier WHERE DATA.suppliercode = supplier.fccode"

txtSQL	= 	" SELECT" & _
			"     FCBA" & _
			"     , INVOICECODE" & _
			"     , TANGGAL" & _
			"     , SUPPLIERCODE" & _
			"     , SUPPLIER_API.GET_SUPP_NAME(SUPPLIERCODE, FCBA) SUPPLIERNAME" & _
			"     , SUM(PAID) PAID" & _
			"     , SUM(TOTALINVOICE) TOTALINVOICE" & _
			"     , SUM(TOTALINVOICE) - SUM(PAID) NETBALANCE" & _
			" FROM (" & _
			"     SELECT" & _
			"         FCBA FCBA" & _
			"         , INVOICECODE INVOICECODE" & _
			"         , INVOICEDATE TANGGAL" & _
			"         , SUPPLIERCODE SUPPLIERCODE" & _
			"         , PAID PAID" & _
			"         , 0 TOTALINVOICE" & _
			"     FROM INVOICE" & _
			"     UNION ALL" & _
			"     SELECT" & _
			"         INVD.FCBA FCBA" & _
			"         , INVD.INVOICECODE INVOICECODE" & _
			"         , INV.INVOICEDATE TANGGAL" & _
			"         , INV.SUPPLIERCODE SUPPLIERCODE" & _
			"         , 0 PAID" & _
			"         , INVD.AMOUNT TOTALINVOICE" & _
			"     FROM INVOICEDETAIL INVD" & _
			"     INNER JOIN INVOICE INV ON (INV.INVOICECODE = INVD.INVOICECODE AND INV.FCBA = INVD.FCBA)" & _
			"     ) WHERE FCBA = '" & Session("FCBA") & "'" & _
			" GROUP BY" & _
			"     FCBA" & _
			"     , INVOICECODE" & _
			"     , TANGGAL" & _
			"     , SUPPLIERCODE" & _
			" HAVING (SUM(TOTALINVOICE) - SUM(PAID) <> 0)" & _
			" ORDER BY TANGGAL, INVOICECODE, SUPPLIERCODE"


'response.Write txtSQL
%>
	<body>
		<div id="wrapper">
			<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					<th colspan="7" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
						<div id="head_judul" style="float:left; text-align:left;">
							<h4>PT. <%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h1 class="judul">INVOICE OUTSTANDING</h1>
						</div>
						<div style="clear:both;"></div>
					</div>
					</th>
				</tr>
				<tr>
					<th>INVOICE CODE</th>
					<th>INVOICE DATE</th>
					<th>SUPPLIER CODE</th>
					<th>SUPPLIER NAME</th>
					<th>PAID</th>
					<th>TOTAL INVOICE</th>
					<th>NET BALANCE</th>
				</tr>
			</thead>
			<tbody>
				
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
				   nCountofField=rsCat.fields.count-1
				end if
				NO = 1
				%>
				<%If (rsCat.BOF and rsCat.EOF) Then%>
					<tr>
						<td Colspan = "7" align = "Left">NO DATA AVAILABLE</td>
					</tr>
				<%Else%>
					<%	While Not rsCat.EOF %>
					<tr>
						<td align="Left"><%=rscat("INVOICECODE")%></td>
						<td align="Right"><%=FbpNum2(rscat("TANGGAL"))%></td>
						<td align="Left"><%=rscat("SUPPLIERCODE")%></td>
						<td align="Left"><%=rscat("SUPPLIERNAME")%></td>
						<td align="Right"><%=FormatNumber(rscat("PAID"),2)%></td>
						<td align="Right"><%=FormatNumber(rscat("totalinvoice"),2)%></td>
						<td align="Right"><%=FormatNumber(rscat("netbalance"),2)%></td>
					</tr>
					<%
						nRptotalinvoice  = nRptotalinvoice+cdbl(rsCat("totalinvoice"))
						nRpNetbalance  = nRpNetbalance+cdbl(rsCat("netbalance"))
						
						rsCat.MoveNext
						'response.Flush()
						'no=no+1
					Wend
					rsCat.close()
					rsCat=null
					%>
					<tr bgcolor="#F4F4A0">
						<td align="Center" colspan="5"><strong><i>T o t a l</i></strong></td>
						<td align="Right"><strong><i><%=FormatNumber(nRptotalinvoice,2)%></i></strong></td>
						<td align="Right"><strong><i><%=FormatNumber(nRpNetbalance,2)%></i></strong></td>
					</tr>
				<%End If%>
			</tbody>
			</table>
		</div>
	</body>
</html>