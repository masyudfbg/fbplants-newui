<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsStockGudang, strStockGudang, strGroup, strOrder, strQuery, strmain
	Dim v_txttodaydate, v_txtStock, v_txtStockGroup,strsql, NO, cClass
	Dim rsAccountingPeriod, v_AccountingYear, v_AccountingPeriod, v_txtfcba
	Dim nsptotal, TempStockGroup, TempStock
	Dim txttodaydate, txtStockGroup, txtStock, txtfcba
	
	txttodaydate		=	request("txttodaydate")
	txtStockGroup		=	request("txtStockGroup")
	txtStock			=	request("txtStock")
	txtfcba     		=  request("txtfcba")
	
	nsptotal= 4
	
	v_txttodaydate		=	txttodaydate
	v_txtStockGroup		=	txtStockGroup
	v_txtStock			=	txtStock
	v_txtfcba     		=   txtfcba

	
	
	strmain  =  " SELECT a.fcba, c.fcname stockgroupname, a.stockcode, b.fcname, " & _
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
						   " WHERE transdate < to_date('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
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
				   " AND a.fcba = b.fcba " & _
				   " AND a.fcba = c.fcba " & _
				   " GROUP BY a.fcba, a.stockcode, b.fcname, c.fcname, b.unitofmeasurement"
	
		
	strStockGudang = " select fcba"
	
		if len(trim(v_txtStockGroup)) > 0 then 
			if v_txtStockGroup <> "--SUMMARY--" then
			strStockGudang = strStockGudang & " , StockGroupname "
			end if
		else
			strStockGudang = strStockGudang & " , StockGroupname "
		end if

		if len(trim(v_txtStock)) > 0 then 
			if v_txtStock <> "--SUMMARY--" then
				strStockGudang = strStockGudang & " , stockcode, fcname, unitofmeasurement "
			end if
		else
			strStockGudang = strStockGudang & " , stockcode , fcname, unitofmeasurement"	
		end if
		
		strStockGudang = strStockGudang & ", SUM (opening_balance_day) opening_balance_day," & _
						" SUM (received) received," & _
						" SUM (issued) issued," & _
						" SUM (adjustment) adjustment," & _
						" SUM (closing_balance_day) closing_balance_day"
						
		strStockGudang = strStockGudang &  " FROM ( " & strMain & " )  " 
		
		
			strStockGudang = strStockGudang & " WHERE fcba = '" & Session("fcba") & "'" 
		
		' where condition
		
		if len(trim(v_txtStockGroup)) > 0 then 
			if v_txtStockGroup <> "--SUMMARY--" then
				strmain = strmain & " AND StockGroup = '" & v_txtStockGroup & "'"
			end if
		end if
		
		if len(trim(v_txtStock)) > 0 then 
			if  v_txtStock <> "--SUMMARY--" then
				strmain = strmain & " AND Stock = '" & v_txtStock & "'"
			end if
		end if
			
		' Group By
		
		strStockGudang = strStockGudang & " GROUP BY FCBA "
		if len(trim(v_txtStockGroup)) > 0 then 
			if v_txtStockGroup <> "--SUMMARY--" then
				strStockGudang = strStockGudang & " , StockGroupname "
			end if
		else
				strStockGudang = strStockGudang & " , StockGroupname "
		end if

		if len(trim(v_txtStock)) > 0 then 
			if v_txtStock <> "--SUMMARY--" then
				strStockGudang = strStockGudang & " , stockcode, fcname, unitofmeasurement "
			end if
		else
				strStockGudang = strStockGudang & " , stockcode, fcname, unitofmeasurement "	
		end if
		
		' Order by
		
		' strStockGudang = strStockGudang & " order BY FCBA "
		' if len(trim(v_txtStockGroup)) > 0 then 
			' if v_txtStockGroup <> "--SUMMARY--" then
				' strmain = strmain & " , StockGroupname "
			' end if
		' else
			' strmain = strmain & " , StockGroupname "
		' end if

		' if len(trim(v_txtStock)) > 0 then 
			' if v_txtStock <> "--SUMMARY--" then
				' strmain = strmain & " , Stock , fcname, unitofmeasurement"
			' end if
		' else
			' strmain = strmain & " , Stock , fcname, unitofmeasurement"	
		' end if
		
' response.write strStockGudang
			
	Set rsStockGudang = Server.CreateObject("ADODB.Recordset")
	rsStockGudang.Open strStockGudang, connstr, adOpenStatic 
	
	
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
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/DashboardManager.asp">Back to Report</a></td>
		</tr>
	</table>
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
		<th>STOCK GROUP</th>
		<th>STOCK CODE</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_StockGudangMonthly_BA_NAF_M.asp">
<td align="center">
		<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
				<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
		<% else %>
				<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if %>
				<a href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
			<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		<td align="center">
		<select name="txtStockGroup">	
		<option value="">--ALL--</option>
		<%
			Dim rsStockGroup
			Set rsStockGroup = Server.CreateObject("ADODB.Recordset")
			rsStockGroup.CursorLocation = adUseClient
			
			
			 if len(trim(v_txtFcba)) > 0 then 
			 
				strsql =  " SELECT '--SUMMARY--' as Fcname from DUAL UNION SELECT DISTINCT Fcname as STOCKGROUP FROM StockGroup WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--SUMMARY--' as Fcname from DUAL UNION SELECT DISTINCT Fcname as STOCKGROUP FROM StockGroup WHERE FCBA  = '" & Session("fcBA") & "'"
			end if
			
			rsStockGroup.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly
			if not rsStockGroup.eof then
				DO while not rsStockGroup.eof
				if v_txtStockGroup = rsStockGroup("Fcname") then
		%>			<option value="<%=rsStockGroup("Fcname")%>" Selected><%=rsStockGroup("Fcname")%></option>
				<%else%>
					<option value="<%=rsStockGroup("Fcname")%>"><%=rsStockGroup("Fcname")%></option>
				<%end if%>
		<%			rsStockGroup.movenext
				loop
			End if
			rsStockGroup.Close
			set rsStockGroup = nothing
		%>
		</select>		
		</td>
		
		<td align="center">
		<select name="txtStock">	
		<option value="">--ALL--</option>
		<%
			Dim rsStock
			Set rsStock = Server.CreateObject("ADODB.Recordset")
			rsStock.CursorLocation = adUseClient
			
			if len(trim(v_txtFcba)) > 0 then 
				strsql =  " SELECT '--SUMMARY--' as STOCK from DUAL UNION SELECT DISTINCT FCNAME as StockName FROM Stock WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--SUMMARY--' as STOCK from DUAL UNION SELECT DISTINCT FCNAME as StockName FROM Stock WHERE FCBA  = '" & Session("fcBA") & "'"
			end if
			
			rsStock.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsStock.eof then
				DO while not rsStock.eof
				if v_txtStock = rsStock("STOCK") then
		%>			<option value="<%=rsStock("STOCK")%>" Selected><%=rsStock("STOCK")%></option>
				<%else%>
					<option value="<%=rsStock("STOCK")%>"><%=rsStock("STOCK")%></option>
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
		
		<% if trim(request("txtStockGroup")) <> "--SUMMARY--" then %>
			<th>STOCK GROUP</th>
		<% else
			nsptotal = nsptotal - 1 			
	    end if %>
		
		
		<% if trim(request("txtStock")) <> "--SUMMARY--" then %>
			 <th>STOCK CODE</th>
			 <th>STOCK DESCRIPTION</th>
			<th>SATUAN</th>
		<% else
			nsptotal = nsptotal - 3 			
	    end if %>
		
		<th>STOCK AWAL PERIOD</th>
		<th>PENERIMAAN (SHI per PERIOD)</th>
		<th>PENGELUARAN (SHI per PERIOD)</th>
		<th>ADJUSTMENT (SHI per PERIOD)</th>
		<th>STOCK AKHIR PERIOD</th>
	</tr>
<%
Dim nsubOpening, nsubReceive, nsubIssue, nsubAdj, nsubTotal, ngrandTotal
Dim sgfcba,sgStockGroup,sgStock

	if not rsStockGudang.EOF then
		NO = 1
		nsubOpening	= 0
		nsubReceive = 0
		nsubIssue 	= 0
		nsubAdj 	= 0
		nsubTotal 	= 0
		ngrandTotal = 0
		
		sgfcba=""
		sgStockGroup =""
		sgStock=""
		
		
		while not rsStockGudang.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>

				<% if trim(request("txtStockGroup")) <> "--SUMMARY--" then %>
				    <% TempStockGroup = rsStockGudang("stockgroupname") %>
					<% TempStock = "" %>
					
					
				    <% if sgStockGroup <> rsStockGudang("stockgroupname") then %>
				        <td align="left">
							<a href="p_StockGudangMonthly_BA_NAF_M.asp?txtfcba=<%=rsStockGudang("fcba")%>&txttodaydate=<%=v_txttodaydate%>&txtStockGroup=<%=TempStockGroup%>&txtStock=<%=TempStock%>">
							<%=rsStockGudang("stockgroupname")%>
							</a>
						</td>
						<% sgStockGroup = rsStockGudang("stockgroupname") %>
					<% else	%>
					    <td align="left"></td>					
					<% end if %>							 
				<% end if %>
				<% if trim(request("txtStock")) <> "--SUMMARY--" then %>
				    <% TempStockGroup = rsStockGudang("stockcode") %>
					<% TempStock =rsStockGudang ("stockcode") %>
				
					<% if sgStock  <> rsStockGudang ("stockcode") then %>
				        <td align="left">
							<a href="p_StockGudangMonthly_BA_NAF_M.asp?txtfcba=<%=rsStockGudang("fcba")%>&txttodaydate=<%=v_txttodaydate%>&txtStockGroup=<%=TempStockGroup%>&txtStock=<%=TempStock%>">
							<%=rsStockGudang("stockcode")%>
							</a>
						</td>
						 <td align="left">
							
							<%=rsStockGudang("fcname")%>
							
						</td>
						<td align="center">
							
							<%=rsStockGudang("unitofmeasurement")%>
							
						</td>
						<% sgStock= rsStockGudang("stockcode") %>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				<td align="right"><%=FormatNumber(rsStockGudang("opening_balance_day"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("received"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("issued"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("adjustment"),2)%></td>
				<td align="right"><%=FormatNumber(rsStockGudang("closing_balance_day"),2)%></td>
			</tr>
<%
		NO = NO + 1
			nsubOpening = nsubOpening + Cdbl(rsStockGudang("opening_balance_day"))
			nsubReceive = nsubReceive + Cdbl(rsStockGudang("received"))
			nsubIssue = nsubIssue + Cdbl(rsStockGudang("issued"))
			nsubAdj = nsubAdj + Cdbl(rsStockGudang("adjustment"))
			nsubTotal = nsubTotal + Cdbl(rsStockGudang("closing_balance_day"))

			rsStockGudang.MoveNext
		Wend
	end if
	rsStockGudang.Close
	set rsStockGudang = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="<%= 1 + nsptotal %>">TOTAL</td>	
				<td align="right"><%=FormatNumber(nsubOpening,2)%></td>
				<td align="right"><%=FormatNumber(nsubReceive,2)%></td>
				<td align="right"><%=FormatNumber(nsubIssue,2)%></td>
				<td align="right"><%=FormatNumber(nsubAdj,2)%></td>
				<td align="right"><%=FormatNumber(nsubTotal,2)%></td>
			</tr>
</table>

<!--#include file="../Bottom_Report.asp"-->
