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
<title>ePlaS Reporting - Delivery Return Note</title>
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

v_txtDNOTE_DELIVERYNOTECODE= trim(request("DNOTE_DELIVERYNOTECODE"))
v_txtRETURNDATE_From= trim(request("RETURNDATE_From"))
v_txtRETURNDATE_To= trim(request("RETURNDATE_To"))
v_txtPSTORAGE_STORAGECODE= trim(request("PSTORAGE_STORAGECODE"))
v_txtWBT_CHITNO=trim(request("WBT_CHITNO"))

txtSQL = "SELECT RETURNDATE, REFERENCE, REMARKS, DNOTE_DELIVERYNOTECODE, WBT_CHITNO, RETURN_CODE, nvl(sum(QUANTITY),0) as QUANTITY, STORAGE_CODE FROM DELIVERYRETURNNOTE" & _
		 " where fcba='" & session("fcBA") & "'"

if  len(trim(v_txtRETURNDATE_From)) > 0 then
    txtSQL = txtSQL & " and RETURNDATE >=to_Date('" & v_txtRETURNDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtRETURNDATE_From,1)
end if
if  len(trim(v_txtRETURNDATE_To)) > 0 then
    txtSQL =txtSQL &  " and RETURNDATE <=to_Date('" & v_txtRETURNDATE_To & "','MM/DD/YYYY')"
	txtSelect = " To " & FormatDateTime(v_txtRETURNDATE_To,1)
end if
if  len(trim(v_txtDNOTE_DELIVERYNOTECODE)) > 0 then
    txtSQL = txtSQL & " and DNOTE_DELIVERYNOTECODE='" & v_txtDNOTE_DELIVERYNOTECODE & "'"
	txtSelect = txtSelect & " DELIVERYNOTECODE " & v_txtDNOTE_DELIVERYNOTECODE
end if
if  len(trim(v_txtPSTORAGE_STORAGECODE)) > 0 then
    txtSQL = txtSQL & " and STORAGE_CODE='" & v_txtPSTORAGE_STORAGECODE & "'"
	txtSelect = txtSelect & " STORAGECODE " & v_txtPSTORAGE_STORAGECODE
end if
if  len(trim(v_txtWBT_CHITNO)) > 0 then
    txtSQL = txtSQL & " and WBT_CHITNO='" & v_txtWBT_CHITNO & "'"
	txtSelect = txtSelect & " CHITNO " & v_txtWBT_CHITNO
end if


txtsql=(txtsql) & " group by RETURNDATE, REFERENCE, REMARKS, DNOTE_DELIVERYNOTECODE, WBT_CHITNO, RETURN_CODE,STORAGE_CODE"
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
         <th colspan="8" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>DELIVERY RETURN NOTE</h2>
				</div>
			</div>
			</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="8" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
				</tr>
				<!--</div>-->
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Return Date</th>
					<th>Delivery Note Code</th>
					<th>Chit No</th>
					<th>Return code</th>
					<th>Storage Code</th>
					<th>Reference</th>
					<th>Remarks</th>
					<th>Quantity</th>
	</tr>
    </thead>
    <tbody>
	<%
	Dim v_Quantity
	v_Quantity=0
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RETURNDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DNOTE_DELIVERYNOTECODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WBT_CHITNO")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RETURN_CODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STORAGE_CODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REFERENCE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARKS")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Quantity"),2)%></td>
		</tr>
	  <%
	  v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
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
	<td align="Center" colspan="7"> <strong><i> <font face="Verdana" size="2">T o t a l</i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%></i></strong></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</div>
	</body>
</html>