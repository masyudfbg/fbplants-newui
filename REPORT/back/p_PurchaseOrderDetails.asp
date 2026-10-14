<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Purchase Order</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtPODATE_From=trim(request("PODATE_From"))
v_txtPODATE_To=trim(request("PODATE_To"))
v_txtPOCODE=trim(request("POCODE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
v_txtSUPPLIERCODE=trim(request("SUPPLIERCODE"))

if request("txtKind") = "POOutstanding" Then
   txtSQL = "SELECT pur.pocode, pur.podate, sup.fcName as Supplier, ord.itemcode, s.FCNAME, nvl(sum(ord.quantity),0) as Quantity,nvl(sum(ord.Received),0) as Received," & _
		 "  nvl(sum(ord.amount),0) as Amount,pur.remarks   FROM purchaseorder pur, purchaseorderdetails ord,stock s,Supplier sup" & _
		 " WHERE( ((pur.pocode = ord.pocode) and (pur.fcba = ord.fcba)) " & _
		 " and ((ord.itemcode = s.fcCode) and (ord.fcba = s.fcba))" & _
		 " and ((pur.SUPPLIERCODE=Sup.FcCode) and (pur.fcba=Sup.fcba)) )" & _
		 " and pur.fcba='" & session("fcBA") & "'" & _
		 " and ord.received < ord.QUANTITY"
else
   txtSQL = "SELECT pur.pocode, pur.podate, sup.fcName as Supplier, ord.itemcode, s.FCNAME, nvl(sum(ord.quantity),0) as Quantity,nvl(sum(ord.Received),0) as Received," & _
		 "  nvl(sum(ord.amount),0) as Amount,pur.remarks   FROM purchaseorder pur, purchaseorderdetails ord,stock s,Supplier sup" & _
		 " WHERE( ((pur.pocode = ord.pocode) and (pur.fcba = ord.fcba)) " & _
		 " and ((ord.itemcode = s.fcCode) and (ord.fcba = s.fcba))" & _
		 " and ((pur.SUPPLIERCODE=Sup.FcCode) and (pur.fcba=Sup.fcba)) )" & _
		 " and pur.fcba='" & session("fcBA") & "'"
end if		 

if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and pur.podate >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtPODATE_From,1)
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and pur.podate <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtPODATE_To,1)
end if
if  len(trim(v_txtPOCODE)) > 0 then
    txtSQL = txtSQL & " and pur.pocode='" & v_txtPOCODE & "'"
	txtSelect = txtSelect & " PO Code " & v_txtPOCODE
end if
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and ord.itemcode='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Item Code " & v_txtITEMCODE
end if 
if  len(trim(v_txtSUPPLIERCODE)) > 0 then
    txtSQL =txtSQL &  " and pur.Suppliercode='" & v_txtSUPPLIERCODE & "'"
	txtSelect = txtSelect & " Supplier Code " & v_txtSUPPLIERCODE
end if 

txtsql=(txtsql) & " group by pur.pocode, pur.podate, sup.fcName, ord.itemcode,s.FCNAME,pur.remarks"

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()
%>
 
	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>PURCHASE ORDER</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="10"><%=txtSelect%></td>
						</tr>
						</div>
						<tr>
							<%	
							Set rsCat = connStr.Execute(txtSQL)
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
						<tr>
							 
							<th>Date</th>
							<th>PO.Code</th>
							<th>Supplier</th>
							<th>Stock Code</th>
							<th>Stock Description</th>
							<th>Remarks</th>
							<th>Order (Qty)</th>
							<th>Received (Qty)</th>
							<th>Amount</th>
						</tr>
							<%
							Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
							v_Quantity=0
							v_Amount = 0
							While Not rsCat.EOF%>
						<tr>
							 
							<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("podate")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Supplier")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("itemcode")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCNAME")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("Quantity")),2)%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("Received")),2)%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("Amount")),2)%> </td>
						</tr>
					  <%
					  v_Quantity=v_Quantity+trim(cdbl(rsCat("Quantity")))
					  v_Amount=v_Amount+trim(cdbl(rsCat("Amount")))
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
					<tr bgcolor="#F4F4A0">
						<td align="Center" style="<%=cmdstyleLeft%>" colspan="7"> <font face="Verdana" size="1">T o t a l</td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%>&nbsp;</td>
						<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Amount,2)%>&nbsp;</td>
					</tr>
				</table>
			<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>
	