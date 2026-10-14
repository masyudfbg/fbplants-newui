<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Delivery Contract</title>
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

v_txt_CONTRACTCODE= trim(request("CONTRACTCODE"))
v_txtCUST_CUSTOMERCODE= trim(request("CUST_CUSTOMERCODE"))
v_txtPRODUCT_PRODUCTCODE= trim(request("PRODUCT_PRODUCTCODE"))
v_txtSTARTDATE_From= trim(request("STARTDATE_From"))
v_txtSTARTDATE_To= trim(request("STARTDATE_To"))


txtSQL = "SELECT CONTRACTCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, nvl(sum(TOTALWEIGHT),0) AS TOTALWEIGHT, STATUS, REFERENCE, " & _
		"CUST_CUSTOMERCODE, PRODUCT_PRODUCTCODE FROM DELIVERYCONTRACT" & _
		 " where fcba='" & session("fcBA") & "'"

if  len(trim(v_txtSTARTDATE_From)) > 0 then
    txtSQL = txtSQL & " and STARTDATE >=to_Date('" & v_txtSTARTDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtSTARTDATE_From,1)
end if
if  len(trim(v_txtSTARTDATE_To)) > 0 then
    txtSQL =txtSQL &  " and STARTDATE <=to_Date('" & v_txtSTARTDATE_To & "','MM/DD/YYYY')"
	txtSelect = " To " & FormatDateTime(v_txtSTARTDATE_To,1)
end if
if  len(trim(v_txt_CONTRACTCODE)) > 0 then
    txtSQL = txtSQL & " and CONTRACTCODE='" & v_txt_CONTRACTCODE & "'"
	txtSelect = txtSelect & " CONTRACTCODE " & v_txt_CONTRACTCODE
end if

if  len(trim(v_txtCUST_CUSTOMERCODE)) > 0 then
    txtSQL = txtSQL & " and CUST_CUSTOMERCODE='" & v_txtCUST_CUSTOMERCODE & "'"
	txtSelect = txtSelect & " CUSTOMERCODE " & v_txtCUST_CUSTOMERCODE
end if
if  len(trim(v_txtPRODUCT_PRODUCTCODE)) > 0 then
    txtSQL = txtSQL & " and PRODUCT_PRODUCTCODE='" & v_txtPRODUCT_PRODUCTCODE & "'"
	txtSelect = txtSelect & " PRODUCTCODE " & v_txtPRODUCT_PRODUCTCODE
end if

txtsql=(txtsql) & " group by CONTRACTCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, STATUS, REFERENCE, CUST_CUSTOMERCODE, PRODUCT_PRODUCTCODE"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'response.Write(txtsql)
'response.Flush()
%>
 <body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">




					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>DELIVERY CONTRACT</h2>
				</div>
			</div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<div class="printhide">

				<tr>
					<td colspan="11"><%=txtSelect%></td>
				</tr>
				</div>
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Start Date</th>
					<th>End Date</th>
					<th>Actual Start</th>
					<th>Actual End</th>
					<th>Contract Code</th>
					<th>Customer Code</th>
					<th>Product Code</th>
					<th>Status</th>
					<th>Reference</th>
					<th>Total Weight</th>
				</tr>
	<%
	Dim v_TOTALWEIGHT
	v_TOTALWEIGHT=0
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STARTDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ENDDATE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ACTUALSTART")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ACTUALEND")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CONTRACTCODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CUST_CUSTOMERCODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCT_PRODUCTCODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STATUS")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REFERENCE")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("TOTALWEIGHT"),2)%></td>
		</tr>
	  <%
	  v_TOTALWEIGHT=v_TOTALWEIGHT+trim(cdbl(rsCat("TOTALWEIGHT")))
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
	<td align="Center" colspan="9"> <strong><i> <font face="Verdana" size="2">T o t a l</i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_TOTALWEIGHT,2)%></i></strong></td>
	</tr>
	
</table>
<!--#include file="Bottom_Report.asp"-->
</div>
	</body>
</html>