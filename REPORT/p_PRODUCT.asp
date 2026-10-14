<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Product</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtPRODUCTCODE= trim(request("PRODUCTCODE"))
v_txtCROPCODE= trim(request("CROPCODE"))

txtSQL = "SELECT PRODUCTCODE, PRODUCTDESCRIPTION, CROPCODE, UNITOFMEASURE FROM Product" & _
		 " where fcba='" & session("fcBA") & "'"
if  len(trim(v_txtPRODUCTCODE)) > 0 then
    txtSQL = txtSQL & " and PRODUCTCODE='" & v_txtPRODUCTCODE & "'"
	txtSelect = txtSelect & " PRODUCTCODE " & v_txtPRODUCTCODE
end if
if  len(trim(v_txtCROPCODE)) > 0 then
    txtSQL = txtSQL & " and CROPCODE='" & v_txtCROPCODE & "'"
	txtSelect = txtSelect & " CROPCODE " & v_txtCROPCODE
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
					<h2>PRODUCT</h2>
				</div>
			</div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<div class="printhide">







				<tr>
					<td colspan="5"><%=txtSelect%></td>
				</tr>
				</div>
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Product Code</th>
					<th>Product Description</th>
					<th>Crop Code</th>
					<th>Unit Of Measure</th>
			
					
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
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCTCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCTDESCRIPTION")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CROPCODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("UNITOFMEASURE")%>&nbsp;</td>
		
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
	<tr bgcolor="#F4F4A0"> </tr>
	
</table>
<!--#include file="Bottom_Report.asp"-->
</div>
	</body>
</html>