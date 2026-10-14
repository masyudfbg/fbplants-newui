<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtINVOICEDATE_From= trim(request("INVOICEDATE_From"))
v_txtINVOICEDATE_To= trim(request("INVOICEDATE_To"))
v_txtSUPPLIERCODE= trim(request("SUPPLIERCODE"))
v_txtINVOICECODE= trim(request("INVOICECODE"))
v_txtITEMCODE= trim(request("ITEMCODE"))

txtSQL =  "SELECT Inv.INVOICECODE, Inv.INVOICEDATE, Inv.SupplierCode, InvDet.ITEMCODE,Stock.FcName," & _
		 " nvl(sum(InvDet.Quantity),0) as quantity, nvl(sum(InvDet.Amount),0) as Amount" & _
		 " FROM Invoice Inv, InvoiceDetail InvDet, stock stock" & _
		 " WHERE ( ((inv.INVOICECODE = invDet.INVOICECODE) and (inv.fcba = invDet.fcba))" & _
		 " AND ((stock.fccode = InvDet.ITEMCODE) and (stock.fcba = InvDet.fcba)) )" & _
		 " AND InvDet.fcba='" & session("fcBA") & "'"

if  len(trim(v_txtINVOICEDATE_From)) > 0 then
    txtSQL = txtSQL & " and Inv.INVOICEDATE >=to_Date('" & v_txtINVOICEDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtINVOICEDATE_From,1)
end if
if  len(trim(v_txtINVOICEDATE_To)) > 0 then
    txtSQL =txtSQL &  " and Inv.INVOICEDATE <=to_Date('" & v_txtINVOICEDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " to " & FormatDateTime(v_txtINVOICEDATE_To,1)
end if
if  len(trim(v_txtSUPPLIERCODE)) > 0 then
    txtSQL = txtSQL & " and Inv.SupplierCode='" & v_txtSUPPLIERCODE & "'"
	txtSelect = txtSelect & " Supplier Code " & v_txtSUPPLIERCODE
end if
if  len(trim(v_txtINVOICECODE)) > 0 then
    txtSQL =txtSQL &  " and Inv.INVOICECODE='" & v_txtINVOICECODE & "'"
	txtSelect = txtSelect & " Inv. Code " & v_txtINVOICECODE
end if 
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and InvDet.ITEMCODE='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Item. Code " & v_txtITEMCODE
end if 

txtsql=(txtsql) & " group by Inv.INVOICECODE, Inv.INVOICEDATE, Inv.SupplierCode, InvDet.ITEMCODE,Stock.FcName"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()
%>
 <table width="541">
  <tr> 
    <td rowspan="5" width="100"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">INVOICE</FONT></td></tr>
 </table>
<BR><BR>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	<tr>
    <td colspan="10" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" ><font face="Verdana" size="1"><%=txtSelect%></td>
	</tr>
	<tr>
    <td colspan="10">&nbsp;</td>
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
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Date</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Invoice<br>Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Supplier<br>Code</td>	
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Item<br>Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Stock<br>Description</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Quantity</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Amount</td>
	</tr>
	<%
	Dim v_Quantity,v_price,v_amount
	v_Quantity=0
	v_amount=0
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICEDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICECODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SupplierCode")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ITEMCODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FcName")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("quantity"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),2)%></td>
		</tr>
	  <%
	  v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
	  v_amount =v_amount+Trim(cdbl(rsCat("amount")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
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
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="6"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amount,2)%></td>
	</tr>
</table>
<!--#include file="Bottom_Report.asp"-->