<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
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
 <table width="700">
  <tr> 
    <td rowspan="6" width="100"><img src="../IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">INVOICE LEDGER</FONT></td></tr>
	<tr><td width="100">Period</td><td  width="300"><b><FONT color="#000099"><%=textMonth(v_txtmonth)%> , <%=v_txtyear%></FONT></td></tr>
 </table>
<BR><BR>
<table width="200" cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField = rsCat.fields.count - 1
	end if
	NO = 1%>
	</tr>
	<tr>
	    <td style="<%=cmdstyleLeft%>"><font face="Verdana" size="1">No</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Supplier Code</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Supplier Name</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Invoice Code</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Invoice Date</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Total Invoice</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Total Balance Paid</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Monthly Paid</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Outstanding Paid</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">0 - 30 Days</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">31 - 90 Days</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">91 - 180 Days</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">>180 Days</td>
	</tr>
	<%While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
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
</table>
<!--#include file="../Bottom_Report.asp"-->