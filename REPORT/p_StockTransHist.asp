<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockTransHist, strStockTransHist
	Dim v_txtDateFrom, v_txtDateTo, v_txtStockCode
	Dim NO, cClass
	Dim totalqty : totalqty = 0
	
	v_txtDateFrom = trim(request("txtDateFrom"))
	v_txtDateTo = trim(request("txtDateTo"))
	v_txtStockCode = trim(request("txtStockCode"))
	
	if "" & v_txtDateFrom = "" then v_txtDateFrom = textdate(date)
	if "" & v_txtDateTo = "" then v_txtDateTo = textdate(date)
	
	strStockTransHist = " SELECT sth.fcba," & _
                        "        sth.stockcode," & _
						"        sth.documentno," & _
						"        sth.storecode," & _
						"        sth.sourcecode," & _
						"        sth.sourcetype," & _
						"        sth.targettype," & _
						"        sth.targetcode," & _
                        "        st.fcname  stockname," & _
                        "        sth.transtype," & _
                        "        sth.transdate," & _
                        "        sth.quantity," & _
                        "        sth.lineno" & _
                        " FROM stocktranshist sth" & _
                        " LEFT JOIN stock st ON (sth.fcba = st.fcba AND sth.stockcode = st.fccode)" & _
                        " WHERE sth.fcba = '" & Session("Fcba") & "'" & _
						"       AND DECODE(UPPER(sth.transtype), 'RECEIVED', sth.targettype, sth.sourcetype) IN ('ST', 'CB', 'DI')"

if "" & v_txtDateFrom <> "" then
    strStockTransHist	=	strStockTransHist & " AND sth.transdate >= to_date('" & v_txtDateFrom & "','mm/dd/yyyy')"
end if
if "" & v_txtDateTo <> "" then
    strStockTransHist	=	strStockTransHist & " AND sth.transdate <= to_date('" & v_txtDateTo & "','mm/dd/yyyy')"
end if
if "" & v_txtStockCode <> "" then
    strStockTransHist	=	strStockTransHist & " AND sth.stockcode = '" & v_txtStockCode & "'"
end if

	strStockTransHist	=	strStockTransHist & " ORDER BY sth.stockcode ASC, sth.lineno DESC"

	' response.write strStockTransHist
	Set rsStockTransHist = Server.CreateObject("ADODB.Recordset")
	rsStockTransHist.Open strStockTransHist, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Stock Transaction History</title>
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
		<th>Kode Stock</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post">
		<td align="center">
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
		</td>
		<td align="center">
			<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >			
		</td>
		<td align="center">
			<input type="text" maxlength="20" name="txtStockCode" size="20" value="<%=v_txtStockCode%>" title="Input Stock code">
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
		<h2>LAPORAN STOCK HISTORY</h3>
		<% if v_txtdatefrom = v_txtdateto then%>
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
		<th>DOCUMENT NO</th>
		<th>STOCK CODE</th>
		<th>STOCK NAME</th>
		<th>TRANSACTION TYPE</th>
		<th>TRANSACTION DATE</th>
		<th>STORE CODE</th>
		<th>SOURCE TYPE</th>
		<th>SOURCE CODE</th>
		<th>TARGET TYPE</th>
		<th>TARGET CODE</th>
		<th>QUANTITY</th>
		<th>LINE NO</th>
	</tr>
</thead>
<tbody>
<%	if rsStockTransHist.bof and rsStockTransHist.eof then%>
	<tr>
		<td colspan="8" align="left">NO DATA AVAILABLE</td>
	</tr>
<%	else
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
				<td align="center"><%=rsStockTransHist("DOCUMENTNO")%></td>
				<td align="center"><%=rsStockTransHist("STOCKCODE")%></td>
				<td align="left"><%=rsStockTransHist("STOCKNAME")%></td>
				<td align="center"><%=rsStockTransHist("TRANSTYPE")%></td>
				<td align="center"><%=rsStockTransHist("TRANSDATE")%></td>
				<td align="center"><%=rsStockTransHist("storecode")%></td>
				<td align="center"><%=rsStockTransHist("sourcetype")%></td>
				<td align="center"><%=rsStockTransHist("sourcecode")%></td>
				<td align="center"><%=rsStockTransHist("targettype")%></td>
				<td align="center"><%=rsStockTransHist("targetcode")%></td>
				<td align="right"><%=FbpNum2(rsStockTransHist("QUANTITY"))%></td>
				<td align="center"><%=rsStockTransHist("LINENO")%></td>
			</tr>
<%
			NO = NO + 1
			totalqty = totalqty + cdbl( rsStockTransHist("QUANTITY") )
			rsStockTransHist.MoveNext
		Wend
%>
	<%if len(trim(v_txtStockCode)) <> 0 then%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="11"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FbpNum2(totalqty)%></strong></td>
				<td align="right">&nbsp;</td>
			</tr>
	<%end if%>
<%
	end if
	rsStockTransHist.Close
	set rsStockTransHist = Nothing
%>
</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
