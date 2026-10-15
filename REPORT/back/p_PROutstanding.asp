<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo, v_txtPRCode, v_txtItemCode

v_txtDateFrom	=	trim(Request("txtDateFrom"))
v_txtDateTo		=	trim(Request("txtDateTo"))
v_txtPRCode		=	trim(Request("txtPRCode"))
v_txtItemCode	=	trim(Request("txtItemCode"))
ScriptName		=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - PR Outstanding</title>
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
		<table id="form-table" class="centered-margin">
			<tr>
				<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<td><font face="verdana" size="1">Dari</font></td>
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
				<td><font face="verdana" size="1">Sampai</font></td>
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
				<td><font face="verdana" size="1">NO PR</font></td><td><input type="text" name="txtPRCode" value="<%=v_txtPRCode%>" /></td>
				<td><font face="verdana" size="1">KODE BARANG</font></td><td><input type="text" name="txtItemCode" value="<%=v_txtItemCode%>" /></td>
				<td><input type="submit" value="View" /></td>
				</form>
			</tr>
		</table>
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered">
			<h2>LAPORAN PR OUTSTANDING</h2>		
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>			
		</div>
	</div>

<table border="0" width="100%" id="product-table">
     <tr>
		<th rowspan="2" style="color:#fff;">NO PR</th>		
		<th rowspan="2" style="color:#fff;">LINE ID</th>
		<th rowspan="2" style="color:#fff;">TANGGAL</th>
		<th rowspan="2" style="color:#fff;">KODE BARANG</th>
		<th rowspan="2" style="color:#fff;">NAMA BARANG</th>
		<th rowspan="2" style="color:#fff;">TANGGAL PENGIRIMAN</th>
		<th colspan="4" style="color:#fff;">QUANTITY</th>
	</tr>
	<tr>		
		<th style="color:#fff;">DI MINTA</th>
		<th style="color:#fff;">DI SETUJUI</th>
		<th style="color:#fff;">DI ORDER</th>	
		<th style="color:#fff;">SISA</th>
	</tr>
	<%
	Dim rsPROut, strPROut , noline
	
	strPROut	=	" SELECT data.fcba," & _
					"        data.prcode," & _
					"        data.lineid," & _
					"        data.prdate," & _
					"        data.itemcode," & _
					"        stock.fcname," & _
					"        data.wanteddeliverydate," & _
					"        data.requestedqty," & _
					"        data.approvedqty," & _
					"        data.orderedqty," & _
					"        data.outstandingOrderedQty" & _
					"   FROM (SELECT a.fcba," & _
					"                a.prcode," & _
					"                a.lineid," & _
					"                b.prdate," & _
					"                a.itemcode," & _
					"                a.wanteddeliverydate," & _
					"                a.requestedqty," & _
					"                a.approvedqty," & _
					"                a.orderedqty," & _
					"                (a.approvedqty - a.orderedqty) outstandingOrderedQty" & _
					"           FROM purchaserequestdetails a, purchaserequest b" & _
					"          WHERE a.prcode = b.prcode AND a.fcba = b.fcba) data," & _
					"        stock" & _
					"  WHERE     data.outstandingOrderedQty > 0" & _
					"        AND data.itemcode = stock.fccode" & _
					"        AND data.fcba = stock.fcba" & _
					" AND data.FCBA = '" & Session("FCBA") & "'"
if  len(trim(v_txtDateFrom)) > 0 then
	strPROut = strPROut & " AND data.prdate >=	to_Date('" & v_txtDateFrom & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
	strPROut = strPROut &  " AND data.prdate <= to_Date('" & v_txtDateTo & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtPRCode)) > 0 then
	strPROut = strPROut &  " AND data.prcode = '" & v_txtPRCode & "'"
end if	
if  len(trim(v_txtItemCode)) > 0 then
	strPROut = strPROut &  " AND data.itemcode = '" & v_txtItemCode & "'"
end if					
					
	Set rsPROut = Server.CreateObject("ADODB.Recordset")
	rsPROut.Open strPROut, connstr, adOpenStatic
    
	if rsPROut.bof and rsPROut.eof then
	%>
	<tr>
			<td align="left" colspan="10"><strong>DATA TIDAK DITEMUKAN</strong></td>
	</tr>	
	<%
	elseif Not rsPROut.eof then
		Dim nrequestedqty, napprovedqty, norderedqty, noutstandingOrderedQty
		
		nrequestedqty			=	0
		napprovedqty 			=	0
		norderedqty 			=	0
		noutstandingOrderedQty	=	0
		while Not rsPROut.eof  %>
		           
		  <tr>
			<td align="left"><%=rsPROut("prcode")%></td>
			<td align="right"><%=rsPROut("lineid")%></td>
			<td align="right"><%=FbpNum2(rsPROut("prdate"))%></td>
			<td align="left"><%=rsPROut("itemcode")%></td>
			<td align="left"><%=rsPROut("fcname")%></td>
			<td align="right"><%=FbpNum2(rsPROut("wanteddeliverydate"))%></td>
			<td align="right"><%=formatnumber( cdbl(rsPROut("requestedqty")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPROut("approvedqty")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPROut("orderedqty")),2)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPROut("outstandingOrderedQty")),2)%></td>
		  </tr>	
		      
<%
			nrequestedqty = nrequestedqty + cdbl(rsPROut("requestedqty"))
			napprovedqty = napprovedqty + cdbl(rsPROut("approvedqty"))
			norderedqty = norderedqty + cdbl(rsPROut("orderedqty"))
			noutstandingOrderedQty = noutstandingOrderedQty + cdbl(rsPROut("outstandingOrderedQty"))
			
			rsPROut.MoveNext
		wend
%>
		  <tr bgcolor="#F4F4A0">
			<td align="center" colspan="6"><strong>TOTAL</td>
			<td align="right"><strong><%=FormatNumber(nrequestedqty,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(napprovedqty,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(norderedqty,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(noutstandingOrderedQty,2)%></strong></td>
		  </tr>	
<%
	end if
rsPROut.Close
set rsPROut = Nothing
%>

</table>
<!--#include file="bottom.Asp"-->