<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtSupplierCode=trim(request("txtSupplierCode"))
v_txtInvoiceCode=trim(request("txtInvoiceCode"))

	txtSQL = "SELECT a.suppliercode, a.invoicecode, b.totalAmount, a.invoicedate, a.netbalance,"
	txtSQL = txtSQL + " decode (GetAgingPosition (a.invoicedate), 0, a.netbalance,0) position_0,"
	txtSQL = txtSQL + " decode (GetAgingPosition (a.invoicedate), 1, a.netbalance,0) position_1,"
	txtSQL = txtSQL + " decode (GetAgingPosition (a.invoicedate), 2, a.netbalance,0) position_2,"
	txtSQL = txtSQL + " decode (GetAgingPosition (a.invoicedate), 3, a.netbalance,0) position_3"
	txtSQL = txtSQL + " FROM invoice a, (  SELECT invoicecode, SUM (amount) totalAmount"
	txtSQL = txtSQL + " FROM invoicedetail where fcba = '" & session("fcBA") & "'"
	txtSQL = txtSQL + " GROUP BY invoicecode) b WHERE a.invoicecode = b.invoicecode and a.netbalance <> 0 and"
	txtSQL = txtSQL + " a.fcba = '" & session("fcBA") & "'"	 

if  len(trim(v_txtSupplierCode)) > 0 then
    txtSQL = txtSQL & " AND SUPPLIERCODE ='" & v_txtsuppliercode & "'"
end if
if  len(trim(v_txtinvoicecode)) > 0 then
    txtSQL =txtSQL &  " AND INVOICECODE ='" & v_txtinvoicecode & "'"
end if
' if  len(trim(v_txtPOCODE)) > 0 then
    ' txtSQL = txtSQL & " and pur.pocode='" & v_txtPOCODE & "'"
	' txtSelect = txtSelect & " PO Code " & v_txtPOCODE
' end if
' if  len(trim(v_txtITEMCODE)) > 0 then
    ' txtSQL =txtSQL &  " and ord.itemcode='" & v_txtITEMCODE & "'"
	' txtSelect = txtSelect & " Item Code " & v_txtITEMCODE
' end if 
' if  len(trim(v_txtSUPPLIERCODE)) > 0 then
    ' txtSQL =txtSQL &  " and pur.Suppliercode='" & v_txtSUPPLIERCODE & "'"
	' txtSelect = txtSelect & " Supplier Code " & v_txtSUPPLIERCODE
' end if 

' txtsql=(txtsql) & " group by pur.pocode, pur.podate, sup.fcName, ord.itemcode,s.FCNAME,pur.remarks"

' if len(trim(txtSelect)) > 0 then
   ' txtSelect = "Selected  by : " & txtSelect
' else
   ' txtSelect = "Selected by All Data"
' end if

'response.Write(txtsql)
response.Flush()
%>
 <table width="700">
  <tr> 
    <td rowspan="4" width="100"><img src="../IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">AGING REPORTS</FONT></td></tr>

	
 </table>
<BR><BR>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	<tr>
    <td colspan="11" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" ><font face="Verdana" size="1"><%=txtSelect%></font></td>
	</tr>
	<tr>
    <td colspan="11">&nbsp;</td>
	</tr>
	<tr>
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
	<tr>
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">No</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Supplier Code</font></td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Invoice Code</font></td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Total Amount</font></td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Invoice Date</font></td>
		<td align="center" style="<%=cmdstyleTOP%>" width="200"> <font face="Verdana" size="1">Net Balance</font></td>
		<td align="center" style="<%=cmdstyleTOP%>" width="100"> <font face="Verdana" size="1">0 - 30 Days</font></td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">31 - 90 Days</font></td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">91 - 180 Days</font></td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">> 180 Days</font></td>
	</tr>
	<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
	v_Quantity=0
	v_Amount = 0
	While Not rsCat.EOF%>


		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("suppliercode")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoicecode")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("totalAmount")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoicedate")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("netbalance")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("position_0")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("position_1")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("position_2")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("position_3")%>&nbsp;</td>


		</tr>
	  <%
	  ' v_Quantity=v_Quantity+trim(cdbl(rsCat("Quantity")))
	  ' v_Amount=v_Amount+trim(cdbl(rsCat("Amount")))
	  ' response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="8"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%>&nbsp;</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Amount,2)%>&nbsp;</td>
	</tr>
</table>
<!--#include file="../Bottom_Report.asp"-->