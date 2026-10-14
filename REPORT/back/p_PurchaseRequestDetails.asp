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

v_txtPRDATE_From=trim(request("PRDATE_From"))
v_txtPRDATE_To=trim(request("PRDATE_To"))
v_txtPRCODE=trim(request("PRCODE"))
v_txtITEMCODE=trim(request("ITEMCODE"))
v_txtRequestorID=trim(request("REQUESTORID"))
v_txtRowStateId=trim(request("ROWSTATEID"))


txtSQL = " SELECT a.PRDATE, " & _
				   " a.PRCODE, " & _
				   " a.REQUESTORID, " & _
				   " a.ROWSTATEID, " & _
				   " a.remarks, " & _
				   " B.ITEMCODE, " & _
				   " (SELECT fcname " & _
					   " FROM stock " & _
					   " WHERE B.ITEMCODE = fccode AND a.fcba = fcba) " & _
					   " item_description, " & _
				   " nvl(b.requestedqty,0) requestedqty, " & _
				   " nvl(b.approvedqty,0) approvedqty, " & _
				   " nvl(b.orderedqty,0) orderedqty " & _
			  " FROM purchaserequest a, purchaserequestdetails b " & _
			 " WHERE a.prcode = b.prcode AND a.fcba = b.fcba " 

if  len(trim(v_txtPRDATE_From)) > 0 then
    txtSQL = txtSQL & " and a.prdate >=to_Date('" & v_txtPRDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtPRDATE_From,1)	
end if
if  len(trim(v_txtPRDATE_To)) > 0 then
    txtSQL =txtSQL &  " and a.prdate <=to_Date('" & v_txtPRDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtPRDATE_To,1)	
end if
if  len(trim(v_txtPRCODE)) > 0 then
    txtSQL = txtSQL & " and a.prcode='" & v_txtPRCODE & "'"
	txtSelect = txtSelect & " PR Code " & v_txtPRCODE	
end if
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and b.itemcode='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Item Code " & v_txtITEMCODE	
end if 
if  len(trim(v_txtRequestorID)) > 0 then
    txtSQL =txtSQL &  " and a.RequestorId='" & v_txtRequestorID & "'"
	txtSelect = txtSelect & " Requestor Code " & v_txtRequestorID
end if 

if  len(trim(v_txtRowStateId)) > 0 then
    txtSQL =txtSQL &  " and a.ROWSTATEID='" & v_txtRowStateId & "'"
	txtSelect = txtSelect & " Status  " & v_txtRowStateId	
end if 


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
					<h2>PURCHASE REQUEST</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="11"><%=txtSelect%></td>
						</tr>
						</div>
							<tr>
									<%	
									
									Set rsCat = Server.CreateObject("ADODB.Recordset")
									rsCat.CursorType = 3
									rsCat.CursorLocation = 3
									rsCat.Open txtSQL, connstr

									
									if Not rsCat.eof then
									   nCountofField=rsCat.fields.count-1
									end if
									NO=1%>
							</tr>
							<tr>
								 
								<th>Date</th>
								<th>PR.Code</th>
								<th>Requestor</th>
								<th>Status</th>
								<th>Remarks</th>
								<th>Stock Code</th>
								<th>Stock Description</th>
								<th>Requested (Qty)</th>
								<th>Approved (Qty)</th>
								<th>Ordered (Qty)</th>
							</tr>
								<%
								Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
								v_Quantity=0
								v_Amount = 0
								While Not rsCat.EOF%>


									<tr>
									 
									<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("prdate")%> </td>
									<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("prcode")%> </td>
									<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RequestorId")%> </td>
									<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RowstateId")%> </td>
									<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Remarks")%> </td>
									<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ItemCode")%></td>
									<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("item_description")%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("requestedqty")),2)%> </td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("approvedqty")),2)%></td>
									<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("orderedqty")),2)%> </td>

									</tr>
								  <%

								  rsCat.MoveNext
								  'response.Flush()
								  no=no+1
								Wend
								rsCat.close()
								rsCat=null%>
								<input  type="hidden" value="<%=txtTable%>" name="txtTable">
								<input  type="hidden" value="<%=txtModule%>" name="txtModule">
								</form>
								
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</body>	
</HTML>