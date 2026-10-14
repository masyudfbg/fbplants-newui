<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Customer</title>
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

v_txtCUSTOMERCODE= trim(request("CUSTOMERCODE"))


txtSQL = "SELECT CUSTOMERCODE, DESCRIPTION, SUPPLIER, ICNO, BANK, BANKACCOUNTNO, ADDRESS, COMPANYNO FROM CUSTOMER" & _
		 " where fcba='" & session("fcBA") & "'"
if  len(trim(v_txtCUSTOMERCODE)) > 0 then
    txtSQL = txtSQL & " and CUSTOMERCODE='" & v_txtCUSTOMERCODE& "'"
	txtSelect = txtSelect & " CUSTOMERCODE " & v_txtCUSTOMERCODE
end if

txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

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
					<h2>CUSTOMER</h2>
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
					<th>Customer Code</th>
					<th>Description</th>
					<th>Supplier</th>
					<th>IC No</th>
					<th>Bank</th>
					<th>Bank Account No</th>
					<th>Address</th>
					<th>Company No</th>
				</tr>
	<%
	Dim v_TOTRNINCHES,v_TOTRNDAYS,v_AVGRAINDAY,v_AVGTEMP,v_AVGWIND
		'v_TOTRNINCHES=0
		'v_TOTRNDAYS=0
		'v_AVGRAINDAY=0
		'v_AVGTEMP=0
		'v_AVGWIND=0				
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CUSTOMERCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DESCRIPTION")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPPLIER")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ICNO")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BANK")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BANKACCOUNTNO")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ADDRESS")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("COMPANYNO")%>&nbsp;</td>
		</tr>
	  <%
		'v_TOTRNINCHES=v_TOTRNINCHES+trim(cdbl(rsCat("TOTRNINCHES")))
		'v_TOTRNDAYS=v_TOTRNDAYS+trim(cdbl(rsCat("TOTRNDAYS")))
		'v_AVGRAINDAY=v_AVGRAINDAY+trim(cdbl(rsCat("AVGRAINDAY")))
		'v_AVGTEMP=v_AVGTEMP+trim(cdbl(rsCat("AVGTEMP")))
		'v_AVGWIND=v_AVGWIND+trim(cdbl(rsCat("AVGWIND")))		

	  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
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
	
</table>
<!--#include file="Bottom_Report.asp"-->