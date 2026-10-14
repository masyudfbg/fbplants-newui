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
	
	v_txtDateFrom		=	trim(request("txtDateFrom"))
	v_txtDateTo			=	trim(request("txtDateTo"))
	v_txtStockCode		=	trim(request("txtStockCode"))
	
	
	strStockTransHist 	=	" SELECT * FROM" & _
							" (SELECT" & _
							" STH.FCBA FCBA," & _
							" STH.STOCKCODE STOCKCODE," & _
							" STO.FCNAME STOCKNAME," & _
							" STH.TRANSTYPE TRANSTYPE," & _
							" STH.TRANSDATE TRANSDATE," & _
							" STH.QUANTITY QUANTITY," & _
							" STH.LINENO LINENO" & _
							" FROM STOCKTRANSHIST STH, STOCK STO" & _
							" WHERE" & _
							" (STH.STOCKCODE = STO.FCCODE AND STH.FCBA = STO.FCBA)" & _
							" ORDER BY STH.FCBA, STH.STOCKCODE ASC, LINENO DESC)" & _
							" WHERE FCBA = '" & Session("Fcba") & "'"

if  len(trim(v_txtDateFrom)) > 0 then
    strStockTransHist	=	strStockTransHist & " AND TRANSDATE >= TO_DATE('" & v_txtDateFrom & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
    strStockTransHist	=	strStockTransHist & " AND TRANSDATE <= TO_DATE('" & v_txtDateTo & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtStockCode)) > 0 then
    strStockTransHist	=	strStockTransHist & "AND STOCKCODE = '"& v_txtStockCode &"'"
end if

	
	Set rsStockTransHist = Server.CreateObject("ADODB.Recordset")
	rsStockTransHist.Open strStockTransHist, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Stock Transaction History</title>
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
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN STOCK HISTORY</h3>
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>
	</div>
	</div>
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
		<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
			<input type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDateFrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center">
		<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
			<input type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDateTo');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center">
			<input type="text" maxlength="11" name="txtStockCode" size="11" value="<%=v_txtStockCode%>" title="Input Stock code">
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th>NO</th>
		<th>STOCK CODE</th>
		<th>STOCK NAME</th>
		<th>TRANSACTION TYPE</th>
		<th>TRANSACTION DATE</th>
		<th>QUANTITY</th>
		<th>LINE NO</th>
	</tr>

<%if rsStockTransHist.eof and rsStockTransHist.eof then%>
	<tr>
		<td colspan="7" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1	
		while not rsStockTransHist.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="left"><%=rsStockTransHist("STOCKCODE")%></td>
				<td align="left"><%=rsStockTransHist("STOCKNAME")%></td>
				<td align="left"><%=rsStockTransHist("TRANSTYPE")%></td>
				<td align="left"><%=rsStockTransHist("TRANSDATE")%></td>
				<td align="right"><%=rsStockTransHist("QUANTITY")%></td>
				<td align="right"><%=rsStockTransHist("LINENO")%></td>
			</tr>
<%
			NO = NO + 1	
			rsStockTransHist.MoveNext
		Wend
	rsStockTransHist.Close
	set rsStockTransHist = Nothing
%>
</table>

<!--#include file="../Bottom_Report.asp"-->
