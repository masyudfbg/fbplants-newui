<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Invoice Ledger</title>
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
cmdstyle	=	"border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop	=	"border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft=	"border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

Dim v_txtsuppliercode
v_txtyear			=	trim(request("txtyear"))
v_txtmonth			=	trim(request("txtmonth"))
v_txtsuppliercode	=	trim(request("txtsuppliercode"))
v_txtinvoicecode	=	trim(request("txtinvoicecode"))

txtSQL	=	" SELECT" & _
			" SUPPLIERCODE," & _
			" SUPPLIER.FCNAME SUPPLIERNAME," & _
			" INVOICECODE," & _
			" INVOICE_DATE," & _
			" YEAR," & _
			" MONTH," & _
			" NVL(SUM(TOTAL_AMOUNT), 0) TOTALAMOUNT," & _
			" NVL(SUM(TOTAL_BALANCE_PAID), 0) TOTALBALANCEPAID," & _
			" NVL(SUM(MONTHLY_PAID), 0) MONTHLYPAID," & _
			" NVL(SUM(OUTSTANDING_PAID), 0) OUTSTANDINGPAID," & _
			" NVL(SUM(POSITION_0), 0) POSITION0," & _
			" NVL(SUM(POSITION_1), 0) POSITION1," & _
			" NVL(SUM(POSITION_2), 0) POSITION2," & _
			" NVL(SUM(POSITION_3), 0) POSITION3" & _
			" FROM INVOICELEDGER INNER JOIN SUPPLIER" &_
			" ON INVOICELEDGER.SUPPLIERCODE = SUPPLIER.FCCODE WHERE INVOICELEDGER.FCBA = '" & Session("fcBA") & "'"

if  len(trim(v_txtyear)) > 0 then
    txtSQL = txtSQL & " AND YEAR =" & v_txtyear
end if
if  len(trim(v_txtmonth)) > 0 then
    txtSQL = txtSQL &  " AND MONTH =" & v_txtmonth
end if
if  len(trim(v_txtsuppliercode)) > 0 then
    txtSQL = txtSQL &  " AND SUPPLIERCODE ='" & v_txtsuppliercode & "'"
end if 
if  len(trim(v_txtinvoicecode)) > 0 then
    txtSQL = txtSQL &  " AND INVOICECODE ='" & v_txtinvoicecode & "'"
end if 

txtsql = (txtsql) & " GROUP BY SUPPLIERCODE, SUPPLIER.FCNAME, INVOICECODE, INVOICE_DATE, YEAR, MONTH"
'response.Write(txtsql)
if v_txtsuppliercode = "" then
   v_txtsuppliercode = "ALL"
end if
response.Flush()
%>
<body>
		<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="50" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>INVOICE LEDGER </h2>
				</div>
			</div>
			  </th>
			</tr>
                        <!--<div class="printhide">-->
						<tr>
							<th colspan="50" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField = rsCat.fields.count - 1
	end if
	NO = 1%>
	</tr>
	<tr>
	    <th> Supplier Code</th>
		<th>Supplier Name</th>
		<th>Invoice Code</th>
		<th>Invoice Date</th>
		<th>Total Invoice</th>
		<th>Total Balance Paid</th>
		<th>Monthly Paid</th>
		<th>Outstanding Paid</th>
		<th>0 - 30 Days</th>
		<th>31 - 90 Days</th>
		<th>91 - 180 Days</th>
		<th>>180 Days</th>
	</tr>
 </thead>
  <tbody>	
	<%While Not rsCat.EOF%>
		<tr>
		
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("SUPPLIERCODE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("SUPPLIERNAME")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("INVOICECODE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("INVOICE_DATE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("TOTALAMOUNT"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("TOTALBALANCEPAID"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("MONTHLYPAID"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("OUTSTANDINGPAID"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("POSITION0"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("POSITION1"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("POSITION2"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("POSITION3"),2)%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	  response.Flush()
	  no = no + 1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>
<!--#include file="../Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>
	