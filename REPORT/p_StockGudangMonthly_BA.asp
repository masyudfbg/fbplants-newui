<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockGudang, strStockGudang, strGroup, strOrder, strQuery
	Dim v_txttodaydate, v_txtStock, v_txtStockGroup,strsql, NO, cClass
	Dim rsAccountingPeriod, v_AccountingYear, v_AccountingPeriod, v_txtfcba

	
	v_txttodaydate		=	trim(request.Form("txttodaydate"))
	v_txtStockGroup		=	trim(request.Form("txtStockGroup"))
	v_txtStock			=	trim(request.Form("txtStock"))
	v_txtfcba     		=   trim(request.Form("txtfcba"))

	
	
	strStockGudang  =  " SELECT a.fcba, c.fcname stockgroupname, a.stockcode, b.fcname, " & _
					   " b.unitofmeasurement, SUM (a.opening_balance_day) opening_balance_day, " & _
					   " SUM (a.received) received, SUM (issued) issued, " & _
					   " SUM (a.adjustment) adjustment, " & _
					   " SUM (a.opening_balance_day) " & _
					   " + SUM (a.received) " & _
					   " + SUM (a.issued) " & _
					   " + SUM (a.adjustment) closing_balance_day " & _
				  " FROM (SELECT   fcba, stockcode, SUM (quantity) opening_balance_day, " & _
								 " 0 received, 0 issued, 0 adjustment " & _
							" FROM stocktranshist " & _
						   " WHERE transdate < to_date(to_char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'),'MM') ||'/01/'|| to_char( TO_DATE ('" & _ 
						   v_txttodaydate & "',  " & _
						   " 'MM/DD/YYYY'),'YYYY'), 'MM/DD/YYYY') " & _
						" GROUP BY fcba, stockcode " & _
						" UNION " & _
						" SELECT   fcba, stockcode, 0 opening_balance_day, " & _
								 " SUM (quantity) received, 0 issued, 0 adjustment " & _
							" FROM stocktranshist " & _
						   " WHERE transdate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
							 " AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
							 " AND transtype IN ('RCVDIRECT', 'RECEIVED') " & _
						" GROUP BY fcba, stockcode " & _
						" UNION " & _
						" SELECT   fcba, stockcode, 0 opening_balance_day, 0 received, " & _
								 " SUM (quantity) issued, 0 adjustment " & _
							" FROM stocktranshist " & _
						   " WHERE transdate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
							 " AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
							 " AND transtype IN ('ST_ISSUED', 'MN_ISSUE') " & _
						" GROUP BY fcba, stockcode " & _
						" UNION " & _
						" SELECT   fcba, stockcode, 0 opening_balance_day, 0 received, 0 issued, " & _
								 " SUM (quantity) adjustment " & _
							" FROM stocktranshist " & _
						   " WHERE transdate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
							 " AND to_char(transdate, 'YYYYMM') = to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYYMM') " & _
							 " AND transtype IN ('ST_WRITEOFF') " & _
						" GROUP BY fcba, stockcode) a, " & _
					   " stock b, " & _
					   " stockgroup c " & _
				 " WHERE a.stockcode = b.fccode " & _
				   " AND b.stockgroupcode = c.fccode " & _
				   " AND a.fcba = '" & v_txtfcba & "' " & _
				   " AND a.fcba = b.fcba " & _
				   " AND a.fcba = c.fcba " 
	if len(trim(v_txtStockGroup)) > 0 then
		strStockGudang = strStockGudang & " AND c.fccode = '" & v_txtStockGroup & "'"
	end if
	if len(trim(v_txtStock)) > 0 then
		strStockGudang = strStockGudang & " AND a.stockcode = '" & v_txtStock & "'"
	end if			
			
						
	strGroup	=	" GROUP BY a.fcba, a.stockcode, b.fcname, c.fcname, b.unitofmeasurement"
	strOrder	=	" ORDER BY a.fcba, c.fcname, a.stockcode ASC"
	
	strQuery	=	strStockGudang + strGroup + strOrder

	Set rsStockGudang = Server.CreateObject("ADODB.Recordset")
	rsStockGudang.Open strQuery, connstr, adOpenStatic 
	
	
	Set rsAccountingPeriod = Server.CreateObject("ADODB.Recordset")

	strsql =  " SELECT to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYY') as ACCOUNTINGYEAR, " & _
                      "to_Char(TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'MM') as ACCOUNTINGPERIOD	FROM DUAL"

	rsAccountingPeriod.Open strsql, connstr, adOpenStatic 

	if not rsAccountingPeriod.eof then
	    v_AccountingYear = rsAccountingPeriod("ACCOUNTINGYEAR")
		v_AccountingPeriod = rsAccountingPeriod("ACCOUNTINGPERIOD")
	End if
	rsAccountingPeriod.Close
	set rsAccountingPeriod = nothing
	
	
	
	
	
	
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Monthly Stock Report</title>
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
	<div id="header" class="fixed-1 clearfix left">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader"class="fixed-1 clearfix centered">
		<h2>LAPORAN STOCK GUDANG PER PERIOD</h2>
		<% if v_txttodaydate <> "" then  %>
		    <h3> <%= v_AccountingPeriod  %> - <%= v_AccountingYear %> </h3>
		<% end if %>	
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>
		<th>KEBUN</th>
		<th>STOCK GROUP</th>
		<th>STOCK CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_StockGudangMonthly_BA.asp">
		<td align="center"><input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
			<a href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center">
		<select name="txtfcba">	
		<option value="">--NONE--</option>		
		<%
			Dim rsFcba
			Set rsFcba = Server.CreateObject("ADODB.Recordset")
			rsFcba.CursorLocation = adUseClient
			strsql =  " SELECT fccode FROM businessunit "
			rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsFcba.eof then
				DO while not rsFcba.eof
				if v_txtFcba = rsFcba("fccode") then
		%>			<option value="<%=rsFcba("fccode")%>" Selected><%=rsFcba("fccode")%></option>
				<%else%>
					<option value="<%=rsFcba("fccode")%>"><%=rsFcba("fccode")%></option>
				<%end if%>
		<%			rsFcba.movenext
				loop
			End if
			rsFcba.Close
			set rsFcba = nothing
		%>
		</select>		
		</td>
		<td>
				<select name="txtStockGroup">	
				<option value="">--ALL--</option>
		<%
			Dim rsStockGroup
			Set rsStockGroup = Server.CreateObject("ADODB.Recordset")
			rsStockGroup.CursorLocation = adUseClient
			strsql =  " SELECT FCCODE, FCNAME FROM StockGroup WHERE FCBA = '" & Session("fcBA") & "' ORDER BY FCNAME ASC"
			rsStockGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsStockGroup.eof then
				DO while not rsStockGroup.eof
				if v_txtStockGroup = rsStockGroup("FCCODE") then
		%>			<option value="<%=rsStockGroup("FCCODE")%>" Selected><%=rsStockGroup("FCNAME")%></option>
				<%else%>
					<option value="<%=rsStockGroup("FCCODE")%>"><%=rsStockGroup("FCNAME")%></option>
				<%end if%>
		<%			rsStockGroup.movenext
				loop
			End if
			rsStockGroup.Close
			set rsStockGroup = nothing
		%>
		</select>		
		</td>
		<td>
				<select name="txtStock">	
				<option value="">--ALL--</option>
		<%
			Dim rsStock
			Set rsStock = Server.CreateObject("ADODB.Recordset")
			rsStock.CursorLocation = adUseClient
			strsql =  " SELECT FCCODE, FCNAME FROM Stock WHERE FCBA = '" & Session("fcBA") & "' ORDER BY FCNAME ASC"
			rsStock.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsStock.eof then
				DO while not rsStock.eof
				if v_txtStock = rsStock("FCCODE") then
		%>			<option value="<%=rsStock("FCCODE")%>" Selected><%=rsStock("FCNAME")%></option>
				<%else%>
					<option value="<%=rsStock("FCCODE")%>"><%=rsStock("FCNAME")%></option>
				<%end if%>
		<%			rsStock.movenext
				loop
			End if
			rsStock.Close
			set rsStock = nothing
		%>
		</select>		
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th>NO</th>
		<th>KEBUN</th>
		<th>STOCK GROUP</th>
		<th>STOCK CODE</th>
		<th>STOCK DESCRIPTION</th>
		<th>SATUAN</th>
		<th>STOCK AWAL PERIOD</th>
		<th>PENERIMAAN (SHI per PERIOD)</th>
		<th>PENGELUARAN (SHI per PERIOD)</th>
		<th>ADJUSTMENT (SHI per PERIOD)</th>
		<th>STOCK AKHIR PERIOD</th>
	</tr>
<%
Dim nsubOpening, nsubReceive, nsubIssue, nsubAdj, nsubTotal, ngrandTotal

	if not rsStockGudang.EOF then
		NO = 1
		nsubOpening	= 0
		nsubReceive = 0
		nsubIssue 	= 0
		nsubAdj 	= 0
		nsubTotal 	= 0
		ngrandTotal = 0
		while not rsStockGudang.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="left"><%=rsStockGudang(0)%></td>
				<td align="left"><%=rsStockGudang(1)%></td>
				<td align="left"><%=rsStockGudang(2)%></td>
				<td align="left"><%=rsStockGudang(3)%></td>
				<td align="center"><%=rsStockGudang(4)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang(5),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang(6),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang(7),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang(8),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang(9),2)%></td>
			</tr>
<%
			NO = NO + 1
			nsubOpening = nsubOpening + Cdbl(rsStockGudang(5))
			nsubReceive = nsubReceive + Cdbl(rsStockGudang(6))
			nsubIssue = nsubIssue + Cdbl(rsStockGudang(7))
			nsubAdj = nsubAdj + Cdbl(rsStockGudang(8))
			nsubTotal = nsubTotal + Cdbl(rsStockGudang(9))

			rsStockGudang.MoveNext
		Wend
	end if
	rsStockGudang.Close
	set rsStockGudang = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="6">TOTAL</td>
				<td align="right"><%=FormatNumber(nsubOpening,2)%></td>
				<td align="right"><%=FormatNumber(nsubReceive,2)%></td>
				<td align="right"><%=FormatNumber(nsubIssue,2)%></td>
				<td align="right"><%=FormatNumber(nsubAdj,2)%></td>
				<td align="right"><%=FormatNumber(nsubTotal,2)%></td>
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->
