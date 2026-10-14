<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtDateFrom, v_txtDateTo, v_txtSupplierCode, v_txtPOCode, v_txtItemCode

v_txtDateFrom		=	trim(Request("txtDateFrom"))
v_txtDateTo			=	trim(Request("txtDateTo"))
v_txtSupplierCode	=	trim(Request("txtSupplierCode"))
v_txtPOCode			=	trim(Request("txtPOCode"))
v_txtItemCode		=	trim(Request("txtItemCode"))
ScriptName			=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - PO Outstanding</title>
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
				<td><font face="verdana" size="1">NO PO</font></td><td><input type="text" name="txtPOCode" value="<%=v_txtPOCode%>" /></td>
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
			<h2>LAPORAN PO OUTSTANDING</h2>		
		<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
			<h5>TANGGAL: <%=v_txtdateto%></h5>
		<% else %>
			<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
		<% end if %>			
		</div>
	</div>

<table border="0" width="100%" id="product-table">
     <tr>
		<th rowspan="2" style="color:#fff;">NO PO</th>		
		<th rowspan="2" style="color:#fff;">LINENOKEY</th>
		<th rowspan="2" style="color:#fff;">TANGGAL</th>
		<th rowspan="2" style="color:#fff;">KODE BARANG</th>
		<th rowspan="2" style="color:#fff;">NAMA BARANG</th>
		<th rowspan="2" style="color:#fff;">JUMLAH (Rp)</th>
		<th colspan="4" style="color:#fff;">QUANTITY</th>
		<th rowspan="2" style="color:#fff;">KETERANGAN</th>
		<th rowspan="2" style="color:#fff;">NO PR</th>
		<th rowspan="2" style="color:#fff;">LINE NO</th>
		<th rowspan="2" style="color:#fff;">NOTES</th>
	</tr>
	<tr>		
		<th style="color:#fff;">DI ORDER</th>
		<th style="color:#fff;">DI TERIMA</th>
		<th style="color:#fff;">DI MINTA</th>	
		<th style="color:#fff;">SISA</th>
	</tr>
	<%
	Dim rsPOOut, strPOOut , noline
	
	strPOOut	=	" SELECT fcba," & _
					"        pocode," & _
					"        linenokey," & _
					"        podate," & _
					"        itemcode," & _
					"        fcname," & _
					"        amount," & _
					"        orderedqty," & _
					"        receivedqty," & _
					"        requestedqty," & _
					"        outstandingreceivedqty," & _
					"        otheritemdesc," & _
					"        prcode," & _
					"        prlineid," & _
					"        notes" & _
					"   FROM (SELECT data.fcba," & _
					"                data.pocode," & _
					"                data.linenokey," & _
					"                data.podate," & _
					"                data.itemcode," & _
					"                stock.fcname," & _
					"                data.amount," & _
					"                data.orderedQty," & _
					"                data.receivedQty," & _
					"                data.requestedqty," & _
					"                data.approvedqty," & _
					"                data.OutstandingReceivedQty," & _
					"                data.otheritemdesc," & _
					"                data.prcode," & _
					"                data.prlineid," & _
					"                'PO without PR Information' NOTES" & _
					"           FROM (SELECT a.fcba," & _
					"                        a.pocode," & _
					"                        a.linenokey," & _
					"                        b.podate," & _
					"                        a.itemcode," & _
					"                        a.quantity orderedQty," & _
					"                        a.amount," & _
					"                        a.received receivedQty," & _
					"                        0 requestedqty," & _
					"                        0 approvedqty," & _
					"                        (a.quantity - a.received) OutstandingReceivedQty," & _
					"                        a.otheritemdesc," & _
					"                        a.prcode," & _
					"                        a.prlineid" & _
					"                   FROM purchaseorderdetails a, purchaseorder b" & _
					"                  WHERE     a.pocode = b.pocode" & _
					"                        AND a.fcba = b.fcba" & _
					"                        AND b.rowstateid = 'Ordered'" & _
					"                        AND a.prcode IS NULL) data," & _
					"                stock" & _
					"          WHERE data.itemcode = stock.fccode AND data.fcba = stock.fcba" & _
					"         UNION" & _
					"         SELECT data.fcba," & _
					"                data.pocode," & _
					"                data.linenokey," & _
					"                data.podate," & _
					"                data.itemcode," & _
					"                stock.fcname," & _
					"                data.amount," & _
					"                data.orderedQty," & _
					"                data.receivedQty," & _
					"                data.requestedqty," & _
					"                data.approvedqty," & _
					"                data.OutstandingReceivedQty," & _
					"                data.otheritemdesc," & _
					"                data.prcode," & _
					"                data.prlineid," & _
					"                'PO with PR Information' NOTES" & _
					"           FROM (SELECT a.fcba," & _
					"                        a.pocode," & _
					"                        a.linenokey," & _
					"                        b.podate," & _
					"                        a.itemcode," & _
					"                        a.quantity orderedQty," & _
					"                        a.amount," & _
					"                        a.received receivedQty," & _
					"                        c.requestedqty," & _
					"                        c.approvedqty," & _
					"                        (a.quantity - a.received) OutstandingReceivedQty," & _
					"                        a.otheritemdesc," & _
					"                        a.prcode," & _
					"                        a.prlineid" & _
					"                   FROM purchaseorderdetails a," & _
					"                        purchaseorder b," & _
					"                        purchaserequestdetails c" & _
					"                  WHERE     a.pocode = b.pocode" & _
					"                        AND a.fcba = b.fcba" & _
					"                        AND A.PRCODE = c.prcode" & _
					"                        AND A.PRLINEID = C.LINEID" & _
					"                        AND b.rowstateid = 'Ordered'" & _
					"                        AND a.prcode IS NOT NULL) data," & _
					"                stock" & _
					"          WHERE data.itemcode = stock.fccode AND data.fcba = stock.fcba)" & _
					"          WHERE FCBA = '" & Session("FCBA") & "'"
if  len(trim(v_txtDateFrom)) > 0 then
	strPOOut = strPOOut & " AND podate >=	to_Date('" & v_txtDateFrom & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDateTo)) > 0 then
	strPOOut = strPOOut &  " AND podate <= to_Date('" & v_txtDateTo & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtPOCode)) > 0 then
	strPOOut = strPOOut &  " AND pocode = '" & v_txtPOCode & "'"
end if	
if  len(trim(v_txtItemCode)) > 0 then
	strPOOut = strPOOut &  " AND itemcode = '" & v_txtItemCode & "'"
end if	
					
	Set rsPOOut = Server.CreateObject("ADODB.Recordset")
	rsPOOut.Open strPOOut, connstr, adOpenStatic
    
	if (rsPOOut.bof or rsPOOut.eof) then
	%>
	<tr>
			<td align="left" colspan="14"><strong>DATA TIDAK DITEMUKAN</strong></td>
	</tr>	
	<%
	elseif Not rsPOOut.eof then		
		while Not rsPOOut.eof  %>
		           
		  <tr>
			<td align="left"><%=rsPOOut("pocode")%></td>
			<td align="right"><%=rsPOOut("linenokey")%></td>
			<td align="right"><%=rsPOOut("podate")%></td>
			<td align="left"><%=rsPOOut("itemcode")%></td>
			<td align="left"><%=rsPOOut("fcname")%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("amount")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("orderedqty")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("receivedqty")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("requestedqty")),0)%></td>
			<td align="right"><%=formatnumber( cdbl(rsPOOut("outstandingReceivedQty")),0)%></td>
			<td align="left"><%=rsPOOut("otheritemdesc")%></td>
			<td align="left"><%=rsPOOut("prcode")%></td>
			<td align="right"><%=rsPOOut("prlineid")%></td>
			<td align="left"><%=rsPOOut("notes")%></td>
		  </tr>	
		      
<%
			
			rsPOOut.MoveNext
		wend
	end if
%>

</table>
<!--#include file="bottom.Asp"-->