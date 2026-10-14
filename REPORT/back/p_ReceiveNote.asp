<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Receive Note</title>
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


v_txtRNDATE_From=trim(request("RNDATE_From"))
v_txtRNDATE_To=trim(request("RNDATE_To"))
v_txtRECEIVENOTECODE=trim(request("RECEIVENOTECODE"))
v_txtPURCHASEITEMCODE=trim(request("PURCHASEITEMCODE"))
v_txtSUPPLIERCODE=trim(request("SUPPLIERCODE"))

txtSQL = "SELECT rec.receivenotecode, rec.pocode, rec.rndate, rd.locationtype," & _
		 "rd.locationcode, rd.purchaseitemcode, stock.fcname,  nvl(sum(rd.quantity),0) as quantity ,  nvl(sum(rd.qtyreturn),0) as  qtyreturn," & _
		 " nvl(sum(rd.amountreceived),0) as amountreceived,  nvl(sum(rd.amountreturn),0) as amountreturn FROM receivenote rec, receivenotedetail rd, stock stock" & _
		 " WHERE ( ((rec.receivenotecode = rd.receivenotecode) and (rec.fcba = rd.fcba))" & _
		 " AND ((stock.fccode = rd.purchaseitemcode) and (stock.fcba = rd.fcba))  )" & _
		 " AND rec.fcba='" & session("fcBA") & "'"

if  len(trim(v_txtRNDATE_From)) > 0 then
    txtSQL = txtSQL & " and rec.rndate >=to_Date('" & v_txtRNDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtRNDATE_From,1)
end if
if  len(trim(v_txtRNDATE_To)) > 0 then
    txtSQL =txtSQL &  " and rec.rndate <=to_Date('" & v_txtRNDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect &" To " & FormatDateTime(v_txtRNDATE_To,1)
end if
if  len(trim(v_txtRECEIVENOTECODE)) > 0 then
    txtSQL = txtSQL & " and rec.receivenotecode='" & v_txtRECEIVENOTECODE & "'"
	txtSelect = txtSelect &" RN Code " & v_txtRECEIVENOTECODE
end if
if  len(trim(v_txtPURCHASEITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and rd.PURCHASEITEMCODE='" & v_txtPURCHASEITEMCODE & "'"
	txtSelect = txtSelect &" PO Code " & v_txtPURCHASEITEMCODE
end if 
if  len(trim(v_txtSUPPLIERCODE)) > 0 then
    txtSQL =txtSQL &  " and rec.SUPPLIERCODE='" & v_txtSUPPLIERCODE & "'"
	txtSelect = txtSelect & " Supplier Code " & v_txtSUPPLIERCODE
end if 

txtsql=(txtsql) & " group by rec.receivenotecode, rec.pocode, rec.rndate, rd.locationtype,rd.locationcode, rd.purchaseitemcode, stock.fcname"

'response.Write(txtsql)
response.Flush()

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

%>
 <body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>RECEIVE NOTE</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="10"><%=txtSelect%></td>
						</tr>
						</div>
						<TR>
							<%	
							Set rsCat = connStr.Execute(txtSQL)
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
						<tr>
							 
							<th>Code</th>
							<th>Date</th>
							<th>PO CODE</th>
							<th>Target Type</th>
							<th>Target Code</th>
							<th>Stock Code.</th>
							<th>Stock Desc.</th>
							<th>Quantity</th>
							<th>Amount</th>
						</tr>
							<%
							Dim v_Quantity,v_qtyretrun,v_amountreceived,v_amountreturn
							v_Quantity=0
							v_qtyretrun=0
							v_amountreceived=0
							v_amountreturn=0
							While Not rsCat.EOF%>
						<tr>
								 
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("receivenotecode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rndate")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationtype")%></td>
								<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationcode")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("purchaseitemcode")%></td>
								<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("quantity"),2)%></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amountreceived"),2)%></td>
						</tr>
						  <%
						  v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
						  'v_qtyretrun  =v_qtyretrun+Trim(cdbl(rsCat("qtyreturn")))
						  v_amountreceived =v_amountreceived+Trim(cdbl(rsCat("amountreceived")))
						  'v_amountreturn  =v_amountreturn+trim(cdbl(rsCat("amountreturn")))
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
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amountreceived,2)%></td>
							</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	