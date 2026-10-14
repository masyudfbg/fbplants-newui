<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockTransHist, strStockTransHist
	Dim v_txtDateFrom, v_txtDateTo, v_txtStockCode
	Dim NO, cClass
	Dim totalqty : totalqty = 0
	
	v_txtDateFrom		=	trim(request("txtDateFrom"))
	v_txtDateTo			=	trim(request("txtDateTo"))
	v_txtStockCode		=	trim(request("txtStockCode"))
	
	if len(trim(v_txtDateFrom)) = 0 then v_txtDateFrom = textdate(date)
	if len(trim(v_txtDateTo)) = 0 then v_txtDateFrom = textdate(date)
	
	dim rsstockbibit, sqlstockbibit, liststockbibit
	
	sqlstockbibit = " select report_posting_api.getPostingDataDirect('" & Session("Fcba") & "','STOCKTRANSHIST','p_StockTransHistBibit.asp','STOCK_BIBIT','STOCK')  liststockbibit" & _
                    " from dual"
					
    'response.write sqlstockbibit
	Set rsstockbibit = Server.CreateObject("ADODB.Recordset")
	rsstockbibit.Open sqlstockbibit, connstr, adOpenStatic
	
	if Not rsstockbibit.eof then
	   liststockbibit = rsstockbibit("liststockbibit")
	end if
	rsstockbibit.Close
	Set rsstockbibit = Nothing
	
	
	strStockTransHist = " select fcba," & _
                        "        stockcode," & _
                        "        stock_api.get_stockname(stockcode, fcba) stockname," & _
                        "        decode(upper(transtype), 'RCVDIRECT', targetcode, sourcecode) batchcode," & _
                        "        nursery_api.get_description(decode(upper(transtype), 'RCVDIRECT', targetcode, sourcecode), fcba) batchname," & _
						"        transtype," & _
                        "        to_char(transdate, 'mm/dd/yyyy') transdate," & _
                        "        quantity," & _
                        "        lineno" & _
                        " from stocktranshist" & _
                        " where fcba = '" & Session("Fcba") & "'" & _
						"       and instr('" & liststockbibit & "', stockcode) <> 0"

if  len(trim(v_txtDateFrom)) > 0 then
    strStockTransHist	=	strStockTransHist & " and transdate >= to_date('" & v_txtDateFrom & "','mm/dd/yyyy')"
end if
if  len(trim(v_txtDateTo)) > 0 then
    strStockTransHist	=	strStockTransHist & " and transdate <= to_date('" & v_txtDateTo & "','mm/dd/yyyy')"
end if
if  len(trim(v_txtStockCode)) > 0 then
    strStockTransHist	=	strStockTransHist & " and decode(upper(transtype), 'RCVDIRECT', targetcode, sourcecode) = '" & v_txtStockCode & "'"
end if

	strStockTransHist	=	strStockTransHist & " order by decode(upper(transtype), 'RCVDIRECT', targetcode, sourcecode) asc, lineno desc"

	'response.write strStockTransHist
	Set rsStockTransHist = Server.CreateObject("ADODB.Recordset")
	rsStockTransHist.Open strStockTransHist, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Stock Transaction History Bibit</title>
<link rel="stylesheet" href="/css/jquery-ui.css" />
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
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
     <script>
	$(function() {
	$( ".datepicker" ).datepicker();
	});
	</script>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
    <table id="form-table" class="centered-margin">
	<tr>
		<th>Dari Tanggal</th>
		<th>Sampai Tanggal</th>
		<th>Kode Batch</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post">
		<td align="center">
		<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			
		</td>
		<td align="center">
		<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			
		</td>
		<td align="center">
			<input type="text" maxlength="11" name="txtStockCode" size="11" value="<%=v_txtStockCode%>" title="Input Stock code">
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>
</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>LAPORAN STOCK HISTORY BIBIT</h3>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>
	</div>
	</div>

 </th>
  </tr>
  
	<tr>
		<th>NO</th>
		<th>BATCH CODE</th>
		<th>BATCH NAME</th>
		<th>TRANSACTION TYPE</th>
		<th>TRANSACTION DATE</th>
		<th>QUANTITY</th>
		<th>LINE NO</th>
	</tr>
</thead>
<tbody>
<%if rsStockTransHist.bof and rsStockTransHist.eof then%>
	<tr>
		<td colspan="7" align="left">NO DATA AVAILABLE</td>
	</tr>
<%else
		NO = 1
		totalqty = 0
		while not rsStockTransHist.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="center"><%=rsStockTransHist("BATCHCODE")%></td>
				<td align="center"><%=rsStockTransHist("BATCHNAME")%></td>
				<td align="center"><%=rsStockTransHist("TRANSTYPE")%></td>
				<td align="center"><%=rsStockTransHist("TRANSDATE")%></td>
				<td align="right"><%=rsStockTransHist("QUANTITY")%></td>
				<td align="right"><%=rsStockTransHist("LINENO")%></td>
			</tr>
<%
			NO = NO + 1
			totalqty = totalqty + cdbl( rsStockTransHist("QUANTITY") )
			rsStockTransHist.MoveNext
		Wend
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="5"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=totalqty%></strong></td>
				<td align="right">&nbsp;</td>
			</tr>
<%
	end if
	rsStockTransHist.Close
	set rsStockTransHist = Nothing
%>
</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
