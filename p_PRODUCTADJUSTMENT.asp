<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - Product Adjusment</title>
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

v_txtADJUSTMENTNO= trim(request("ADJUSTMENTNO"))
v_txtPSTORAGE_STORAGECODE= trim(request("PSTORAGE_STORAGECODE"))
v_txtADJUSTMENTDATE_From= trim(request("ADJUSTMENTDATE_From"))
v_txtADJUSTMENTDATE_To= trim(request("ADJUSTMENTDATE_To"))

txtSQL = "SELECT ADJUSTMENTNO, ADJUSTMENTDATE, PSTORAGE_STORAGECODE,nvl(SUM(QUANTITY),0) AS QUANTITY FROM PRODUCTADJUSTMENT" & _
		 " where fcba='" & session("fcBA") & "'"

if  len(trim(v_txtADJUSTMENTDATE_From)) > 0 then
    txtSQL = txtSQL & " and ADJUSTMENTDATE >=to_Date('" & v_txtADJUSTMENTDATE_From & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtADJUSTMENTDATE_To)) > 0 then
    txtSQL =txtSQL &  " and ADJUSTMENTDATE <=to_Date('" & v_txtADJUSTMENTDATE_To & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtADJUSTMENTNO)) > 0 then
    txtSQL = txtSQL & " and ADJUSTMENTNO='" & v_txtADJUSTMENTNO & "'"
end if
if  len(trim(v_txtPSTORAGE_STORAGECODE)) > 0 then
    txtSQL = txtSQL & " and PSTORAGE_STORAGECODE='" & v_txtPSTORAGE_STORAGECODE & "'"
end if

txtsql=(txtsql) & " group by ADJUSTMENTNO, ADJUSTMENTDATE, PSTORAGE_STORAGECODE"

'response.Write(txtsql)
'response.Flush()
%>
<div id="wrapper"> 
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="5" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%>  / <%=Session("FCBA")%></h5>
                    <h5><%=Session("FCMANAGER")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PRODUCT ADJUSMENT</h2>
				</div>
			</div>
				</th>
			</tr>
            <!--<div class="printhide">-->
						<tr>
							<td colspan="5" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
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
	    <th>NO</th>
		<th>DATE</th>
		<th>ADJUSTMENTNO</th>
		<th>PRODUCT STORAGE CODE</th>	
		<th>QUANTITY</th>	
	</tr>
    </thead>
    <tbody>
	<%
	Dim v_Quantity
	v_Quantity=0
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ADJUSTMENTDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ADJUSTMENTNO")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PSTORAGE_STORAGECODE")%>&nbsp;</td>
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
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="4"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%></td>
	</tr>
	</tbody>
</table>

<!--#include file="Bottom_Report.asp"-->
</div>