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
<title>ePlaS Reporting - Delivery Note</title>
</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtCONTRACTCODE= trim(request("CONTRACTCODE"))
v_txtDNOTE_DATE_From= trim(request("DNOTE_DATE_From"))
v_txtDNOTE_DATE_To= trim(request("DNOTE_DATE_To"))
v_txtDELIVERYNOTECODE= trim(request("DELIVERYNOTECODE"))
v_txtPOSTINGDATE= trim(request("POSTINGDATE"))

txtSQL = "SELECT POSTINGDATE, DNOTE_DATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, " & _
         "NVL(SUM(QUANTITY),0) AS QUANTITY , NVL(SUM(FFA),0) AS FFA, NVL(SUM(DIRT),0) AS DIRT, " & _
		 "NVL(SUM(MOISTURE),0) AS MOISTURE, TRANSPORTCO, DRIVERNAME, ICNO FROM DELIVERYNOTE" & _
		 " where fcba='" & session("fcBA") & "'"

 if  len(trim(v_txtPOSTINGDATE)) > 0 then
    txtSQL = txtSQL & " and POSTINGDATE = to_Date('" & v_txtPOSTINGDATE & "','MM/DD/YYYY')"
	txtSelect = " Posting Date " & FormatDateTime(v_txtPOSTINGDATE,1)
	else
		if  len(trim(v_txtDNOTE_DATE_From)) > 0 then
			txtSQL = txtSQL & " and DNOTE_DATE >=to_Date('" & v_txtDNOTE_DATE_From & "','MM/DD/YYYY')"
			txtSelect = " From " & FormatDateTime(v_txtDNOTE_DATE_From,1)
		end if
		if  len(trim(v_txtDNOTE_DATE_To)) > 0 then
			txtSQL =txtSQL &  " and DNOTE_DATE <=to_Date('" & v_txtDNOTE_DATE_To & "','MM/DD/YYYY')"
			txtSelect = " To " & FormatDateTime(v_txtDNOTE_DATE_To,1)
		end if
end if

if  len(trim(v_txtDELIVERYNOTECODE)) > 2 then
    txtSQL = txtSQL & " and DELIVERYNOTECODE='" & v_txtDELIVERYNOTECODE & "'"
	txtSelect = txtSelect & " DELIVERYNOTECODE " & v_txtDELIVERYNOTECODE
end if
if  len(trim(v_txtCONTRACTCODE)) > 0 then
    txtSQL = txtSQL & " and CONTRACTCODE='" & v_txtCONTRACTCODE & "'"
	txtSelect = txtSelect & " CONTRACTCODE " & v_txtCONTRACTCODE
end if

txtsql=(txtsql) & " group by DNOTE_DATE, POSTINGDATE, DELIVERYNOTECODE, CONTRACTCODE, GATEPASS, WBT_CHITNO, STORAGECODE, TRANSPORTCO, DRIVERNAME, ICNO ORDER BY POSTINGDATE, DELIVERYNOTECODE, CONTRACTCODE"
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
        <!--#include file="print_excel.asp"-->
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
					<h2>DELIVERY NOTE</h2>
				</div>
			</div>
			</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
				<!--</div>-->
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					     <th>Date</th>
						<th>Posting Date</th>
					
						
				
						<th>Delivery Note Code</th>
						<th>Contract Code</th>
						<th>Gate Pass</th>
						<th>Chit No</th>
						<th>Storage</th>
						<th>TransportCo</th>
						<th>Driver</th>
						<th>IC No</th>
						<th>Quantity</th>
						<th>FFA</th>	
						<th>DIRT</th>
						<th>Moisture</th>
				</tr>
    </thead>
    <tbody>
	<%
	Dim v_QUANTITY,v_FFA,v_DIRT,v_MOISTURE
	v_QUANTITY = 0
	v_FFA = 0
	v_DIRT = 0
	v_MOISTURE = 0
	v_TOTALWEIGHT=0
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DNOTE_DATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("POSTINGDATE")%></td>
		<% if  len(trim(v_txtPOSTINGDATE)) > 0 then %>
			
		<% else %>
			
		<% end if %>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DELIVERYNOTECODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CONTRACTCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("GATEPASS")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WBT_CHITNO")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STORAGECODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TRANSPORTCO")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DRIVERNAME")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ICNO")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("QUANTITY"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("FFA"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("DIRT"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("MOISTURE"),2)%></td>
		</tr>
	  <%
	  v_QUANTITY=v_QUANTITY+trim(cdbl(rsCat("QUANTITY")))
	  v_FFA=v_FFA+trim(cdbl(rsCat("FFA")))
	  v_DIRT=v_DIRT+trim(cdbl(rsCat("DIRT")))
	  v_MOISTURE=v_MOISTURE+trim(cdbl(rsCat("MOISTURE")))
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
	<td align="Center" colspan="10"> <strong><i> <font face="Verdana" size="2">T o t a l</td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_QUANTITY,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_FFA,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_DIRT,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_MOISTURE,2)%></i></strong></td>
	</tr>
    </tbody>
	
</table>
<!--#include file="Bottom_Report.asp"-->
</div>
	</body>
</html>