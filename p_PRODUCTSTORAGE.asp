<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
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
<title>ePlaS Reporting - Product Storage</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtSTORAGECODE= trim(request("STORAGECODE"))
v_txtPRODUCT_PRODUCTCODE= trim(request("PRODUCT_PRODUCTCODE"))
txtSQL = "SELECT STORAGECODE,  nvl(SUM(MAXCAPACITY),0) AS MAXCAPACITY, DESCRIPTION,  PRODUCT_PRODUCTCODE,  nvl(SUM(DEADSTOCK_CAPACITY),0) AS DEADSTOCK_CAPACITY,REMARKS  FROM PRODUCTSTORAGE" & _
		 " where fcba='" & session("fcBA") & "'"
if  len(trim(v_txtSTORAGECODE)) > 0 then
    txtSQL = txtSQL & " and STORAGECODE='" & v_txtSTORAGECODE & "'"
	txtSelect = txtSelect & " STORAGECODE " & v_txtSTORAGECODE
end if
if  len(trim(v_txtPRODUCT_PRODUCTCODE)) > 0 then
    txtSQL = txtSQL & " and PRODUCT_PRODUCTCODE='" & v_txtPRODUCT_PRODUCTCODE & "'"
	txtSelect = txtSelect & " PRODUCTCODE " & v_txtPRODUCT_PRODUCTCODE
end if

txtsql=(txtsql) & " group by STORAGECODE,  DESCRIPTION,  PRODUCT_PRODUCTCODE,REMARKS"
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
			 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="6" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PRODUCT STORAGE</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="6" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<!--</div>-->
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Code</th>
					<th>Description</th>
					<th>Product Code</th>
					<th>Remarks</th>
					<th>Max Capacity</th>
					<th>Deadstock Capacity</th>
				</tr>
                </thead>
      <tbody>
	<%
	Dim v_MAXCAPACITY,v_DEADSTOCK_CAPACITY
		v_MAXCAPACITY=0
		v_DEADSTOCK_CAPACITY=0
		
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STORAGECODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DESCRIPTION")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCT_PRODUCTCODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARKS")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("MAXCAPACITY"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("DEADSTOCK_CAPACITY"),2)%></td>
		</tr>
	  <%
		v_MAXCAPACITY=v_MAXCAPACITY+trim(cdbl(rsCat("MAXCAPACITY")))
		v_DEADSTOCK_CAPACITY=v_DEADSTOCK_CAPACITY+trim(cdbl(rsCat("DEADSTOCK_CAPACITY")))
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
	<tr bgcolor="#F4F4A0">
	<td align="Center" colspan="4"> <strong><i> <font face="Verdana" size="2">T o t a l</i></strong></td>
	<td align="Right"> <strong><i>  <font face="Verdana" size="1"><%=FormatNumber(v_MAXCAPACITY,2)%></i></strong></td>
	<td align="Right"> <strong><i>  <font face="Verdana" size="1"><%=FormatNumber(v_DEADSTOCK_CAPACITY,2)%></i></strong></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</div>
	</body>
</html>