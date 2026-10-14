<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Nursery Ledger</title>
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

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
'cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
'cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
'cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"


Dim totalopeningqty : totalopeningqty = 0
Dim totalopeningamount : totalopeningamount = 0
Dim totalqtyreceived : totalqtyreceived = 0
Dim totalqtyculled : totalqtyculled = 0
Dim totalqtyissued : totalqtyissued = 0
Dim totalqtysaldo : totalqtysaldo = 0
Dim totalcostperseedling : totalcostperseedling = 0
Dim totalamountreceived : totalamountreceived = 0
Dim totalamountissued : totalamountissued = 0
Dim totalamountsaldo : totalamountsaldo = 0


v_txtyear=trim(request("txtyear"))
v_txtmonth=trim(request("txtmonth"))
v_txtnurserycode=trim(request("txtnurserycode"))

txtSQL =	" select nurserycode" & _
			"        , year" & _
			"        , month" & _
			"        , nvl(sum(openingqty), 0) as openingqty" & _
			"        , nvl(sum(qtyreceived), 0) as qtyreceived" & _
			"        , nvl(sum(qtyadjustment), 0) as qtyadjustment" & _
			"        , nvl(sum(qtyculled), 0) as qtyculled" & _
			"        , nvl(sum(qtyissued), 0) as qtyissued" & _
			"        , nvl((nvl(sum(openingqty), 0) + nvl(sum(qtyreceived), 0) + nvl(sum(qtyadjustment), 0) - nvl(sum(qtyculled), 0) - nvl(sum(qtyissued), 0)), 0) as qtysaldo" & _
			"        , nvl(sum(openingamount), 0) as openingamount" & _
			"        , nvl(sum(amountreceived), 0) as amountreceived" & _
			"        , nvl(sum(amountissued), 0) as amountissued" & _
			"        , nvl((nvl(sum(openingamount), 0) + nvl(sum(amountreceived), 0) - nvl(sum(amountissued), 0)), 0) as rpsaldo" & _
			"        , round(nvl(sum(costperseedling), 0), 2) as costperseedling" & _
			" from NurseryLedger" & _
			" where fcBA = '" & Session("fcBA") & "'"

if  len(trim(v_txtyear)) > 0 then
    txtSQL = txtSQL & " and year = " & v_txtyear
	txtSelect = txtSelect & " Year " & v_txtyear
end if
if  len(trim(v_txtmonth)) > 0 then
    txtSQL =txtSQL &  " and Month = " & v_txtmonth
	txtSelect = txtSelect & " Month " & v_txtmonth
end if
if  len(trim(v_txtnurserycode)) > 0 then
    txtSQL =txtSQL &  " and NURSERYCODE = '" & v_txtnurserycode & "'"
	txtSelect = txtSelect & " Batch " & v_txtnurserycode
end if 
txtsql = txtsql & " group by nurserycode, year, month"
txtsql = txtsql & " order by substr(nurserycode, -2), substr(nurserycode, 1, 2)"
'response.Write(txtsql)
if v_txtnurserycode="" then
   v_txtnurserycode="ALL"
   txtSelect = txtSelect & " Batch ALL"
end if
response.Flush()
%>
<body>
		<div id="wrapper">
	<!--#include file="print_excel.asp"-->
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="12" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>NURSERY LEDGER </h2>
				</div>
			</div>
			  </th>
			</tr>
	<!--<div class="printhide">-->
						<tr>
							<th colspan="12" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
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
	<TR>
		<th rowspan="2">Nursery Code</th>
		<Th colspan="6">Quantity</Th>
		<Th colspan="4">Amount (Rp)</Th>
		<th rowspan="2">Cost Per Seedling</th>
	</TR>	
	<tr>
		<th>Opening</th>
		<th>Received</th>
		<th>Adjustment</th>
		<th>Culling</th>
		<th>Issued</th>
		<th>Saldo</th>
		<th>Opening</th>
		<th>Received</th>
		<th>Issued</th>
		<th>Saldo</th>
	</tr>
	</thead>
  <tbody>
	<%While Not rsCat.EOF%>
		<tr>
		
		<td align="left"><font face="Verdana" size="1"><%=rsCat("NURSERYCODE")%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("OPENINGQTY"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYRECEIVED"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyadjustment"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYCULLED"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYISSUED"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QtySaldo"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("OPENINGAMOUNT"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMOUNTRECEIVED"), 2)%></font></td>		
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMOUNTISSUED"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("RpSaldo"), 2)%></font></td>
		<td align="right"><font face="Verdana" size="1"><%=FormatNumber(rsCat("COSTPERSEEDLING"), 2)%></font></td>
		</tr>
	  <%
		totalopeningqty = totalopeningqty + Cdbl(rsCat("OPENINGQTY"))
		totalqtyreceived = totalqtyreceived + Cdbl(rsCat("QTYRECEIVED"))
		totalqtyadjustment = totalqtyadjustment + Cdbl(rsCat("qtyadjustment"))
		totalqtyculled = totalqtyculled + Cdbl(rsCat("QTYCULLED"))
		totalqtyissued = totalqtyissued + Cdbl(rsCat("QTYISSUED"))
		totalqtysaldo = totalqtysaldo + Cdbl(rsCat("QtySaldo"))
		totalopeningamount = totalopeningamount + Cdbl(rsCat("OPENINGAMOUNT"))
		totalamountreceived = totalamountreceived + Cdbl(rsCat("AMOUNTRECEIVED"))
		totalamountissued = totalamountissued + Cdbl(rsCat("AMOUNTISSUED"))
		totalamountsaldo = totalamountsaldo + Cdbl(rsCat("RpSaldo"))
		totalcostperseedling = totalcostperseedling + Cdbl(rsCat("COSTPERSEEDLING"))
	  rsCat.MoveNext
	  response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>

		<tr style="background-color:#F4F4A0;">		
			<td align="Center"><font face="Verdana" size="1"><strong>TOTAL</strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalopeningqty,2)%></strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalqtyreceived,2)%></strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalqtyadjustment,2)%></strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalqtyculled,2)%></strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalqtyissued,2)%></strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalqtysaldo,2)%></strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalopeningamount,2)%></strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamountreceived,2)%></strong></font></td>		
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamountissued,2)%></strong></font></td>
			<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(totalamountsaldo,2)%></strong></font></td>
			<% if totalqtysaldo <> 0 then %>
				<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(Round(totalamountsaldo/totalqtysaldo, 2),2)%></strong></font></td>
		    <% else %>
				<td align="right"><font face="Verdana" size="1"><strong><%=FormatNumber(Round(0, 2),2)%></strong></font></td>
			<% end if%>
		</tr>
	
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>