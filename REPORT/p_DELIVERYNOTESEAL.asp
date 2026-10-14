<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Delivery Note Seal</title>
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

v_txtDELIVERYNOTECODE= trim(request("DELIVERYNOTECODE"))
v_txtsealno= trim(request("SEALNO"))

txtsql = " select DELIVERYNOTECODE, SEALNO FROM DELIVERYNOTESEAL " &_
		 " where fcba='" & session("fcBA") & "'"
		 
if  len(trim(v_txtDELIVERYNOTECODE)) > 0 then
    txtSQL = txtSQL & " and DELIVERYNOTECODE='" & v_txtDELIVERYNOTECODE & "'"
	txtSelect = txtSelect & "DELIVERYNOTECODE" & v_txtDELIVERYNOTECODE
end if
if  len(trim(v_txtsealno)) > 0 then
    txtSQL = txtSQL & " and SEALNO='" & v_txtsealno & "'"
	txtSelect = txtSelect & "SEALNO" & v_txtsealno
end if

txtsql=(txtsql) & "  GROUP BY DELIVERYNOTECODE, SEALNO "

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
					<h2>DELIVERY NOTE SEAL</h2>
				</div>
			</div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<div class="printhide">

				<tr>
					<td colspan="8"><%=txtSelect%></td>
				</tr>
				</div>
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Delivery Note Code</th>
					<th>Seal No</th>
					
	</tr>
	<% While Not rsCat.EOF %>
	<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DELIVERYNOTECODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SEALNO")%>&nbsp;</td>
		
		</tr>
	  <%
	   rsCat.MoveNext
	  response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null
	%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
</table>
<!--#include file="Bottom_Report.asp"-->