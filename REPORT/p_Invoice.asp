<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Invoice</title>
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
<body>
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

if request("txtKind") = "SupplierBalance" Then

	txtSQL =  "SELECT Inv.INVOICECODE, Inv.INVOICEDATE, Inv.SupplierCode, InvDet.ITEMCODE,Stock.FcName," & _
		 " nvl(sum(InvDet.Quantity),0) as quantity, nvl(sum(InvDet.Amount),0) as Amount" & _
		 " FROM Invoice Inv, InvoiceDetail InvDet, stock stock" & _
		 " WHERE ( ((inv.INVOICECODE = invDet.INVOICECODE) and (inv.fcba = invDet.fcba))" & _
		 " AND ((stock.fccode = InvDet.ITEMCODE) and (stock.fcba = InvDet.fcba)) )" & _
		 " AND InvDet.fcba='" & session("fcBA") & "'"
else
	txtSQL =  "SELECT Inv.INVOICECODE, Inv.INVOICEDATE, Inv.SupplierCode, InvDet.ITEMCODE,Stock.FcName," & _
		 " nvl(sum(InvDet.Quantity),0) as quantity, nvl(sum(InvDet.Amount),0) as Amount" & _
		 " FROM Invoice Inv, InvoiceDetail InvDet, stock stock" & _
		 " WHERE ( ((inv.INVOICECODE = invDet.INVOICECODE) and (inv.fcba = invDet.fcba))" & _
		 " AND ((stock.fccode = InvDet.ITEMCODE) and (stock.fcba = InvDet.fcba)) )" & _
		 " AND InvDet.fcba='" & session("fcBA") & "'"
end if


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
	<body>
		<div id="wrapper">
			<table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>INVOICE</h2>
				</div>
			</div>
			  </th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
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
							<th>Invoice Code</th>
							<th>Supplier Code</th>	
							<th>Item Code</th>
							<th>Stock Description</th>
							<th>Quantity</th>
							<th>Amount</th>
						</tr>
						</thead>
						<tbody>
							<%
							Dim v_Quantity,v_price,v_amount
							v_Quantity=0
							v_amount=0
							While Not rsCat.EOF%>
						<tr>
							 
							<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICEDATE")%></td>
							<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICECODE")%> </td>
							<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SupplierCode")%> </td>
							<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ITEMCODE")%> </td>
							<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FcName")%> </td>
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
							<tr bgcolor="#F4F4A0">
							<td align="Center" style="<%=cmdstyleLeft%>" colspan="5"> <font face="Verdana" size="1">T o t a l</td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amount,2)%></td>
							</tr>
							</tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	