<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Analisa Bibit</title>
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
	<table id="form-table" class="centered-margin">
		<tr>
			<th>Tahun</th>
			<th>Bulan</th>
			<th>Business Unit</th>
		</tr>
		<tr>
				<form name="frmBTL" method="post" action="<%=ScriptName%>">
				<td>
<% 
    MM=Request.Form("vMonth")
    yyyy=Request.Form("vYear")

    mon = Month(Session("v_Costbook"))
    If len(MM)=0 then MM=mon

    Yr = Year(Session("v_Costbook"))
    MinYr = Yr - 2
    MaxYr = Yr  + 2
    If len(yyyy)= 0 then yyyy = Yr
%>
				<select name="vYear">
<%
    For i = MinYr to MaxYr
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
<%
    Next
%>
				</select>
			</td>
			
			<td>
					<select name="vMonth">
<%
    For i = 1 to 12
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
<%
    Next
%>
					</select>
			</td>
			<td>
				<select name="vBa">
				<option value="ALL" <%if request("vBa")="ALL" then%> selected <%End if%>>::ALL::</option>
				<%
				Set rsBa = Server.CreateObject("ADODB.Recordset")
				rsBa.CursorLocation = adUseClient
				strsql =  " SELECT FCCODE, FCNAME FROM BUSINESSUNIT WHERE FCTYPE = 'E' ORDER BY FCCODE ASC"
				rsBa.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
				if not rsBa.eof then
				DO while not rsBa.eof
				%>			
				<option value="<%=rsBa("FCCODE")%>" <%if request("vBa")= rsBa("FCCODE") then%> selected <%End If%>><%=rsBa("FCNAME")%></option>
				<%rsBa.movenext
				loop
				End if
				rsBa.Close
				set rsBa = Nothing
				%>
				</select>
			</td>
			<td>
				<select name="vBatch">
				<option value="ALL" <%if request("vBatch")="ALL" then%> selected <%End if%>>::ALL::</option>
				<%
				Dim rsBatch
				Set rsBatch = Server.CreateObject("ADODB.Recordset")
				rsBatch.CursorLocation = adUseClient
				strsql =  " SELECT NURSERYCODE FCCODE, DESCRIPTION FCNAME FROM NURSERY ORDER BY NURSERYCODE ASC"
				rsBatch.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
				if not rsBatch.eof then
				DO while not rsBatch.eof
				%>			
				<option value="<%=rsBatch("FCCODE")%>" <%if request("vBatch")= rsBatch("FCCODE") then%> selected <%End If%>><%=rsBatch("FCCODE")%></option>
				<%rsBatch.movenext
				loop
				End if
				rsBatch.Close
				set rsBatch = Nothing
				%>
				</select>
			</td>
			<td><input type="submit" value="View" /></td>
				</form>
		</tr>
	</table>
<%
	Dim rsDescBa, strDescBa, vCompanyName, vFcbaName, vManagerName, vKTUName
	strDescBa	=	" SELECT * FROM BUSINESSUNIT WHERE FCCODE = '" & Request("vBa") & "'"
	Set rsDescBa = Server.CreateObject("ADODB.Recordset")
	rsDescBa.Open strDescBa, connstr, adOpenStatic
	if Not rsDescBa.eof then
		vCompanyName	=	"PT. " & rsDescBa("FCCOMPANYNAME")
		vFcbaName		=	rsDescBa("FCNAME")
		vManagerName	=	rsDescBa("FCMANAGER")
		vKTUName		=	rsDescBa("FCKTU")
	end if
	rsDescBa.Close
	set rsDescBa = Nothing
	
	Dim rsActBudget, strActBudget, vActualBibit, vBudgetBibit
	strActBudget	=	" SELECT NVL(((SUM(OPENINGQTY)+SUM(QTYRECEIVED))-(SUM(QTYCULLED)+SUM(QTYISSUED))),0) ACTUALBIBIT,"
	strActBudget	=	strActBudget	& " 0 BUDGETBIBIT FROM NURSERYLEDGER WHERE YEAR||MONTH = '" & request("vYear") & request("vMonth") & "'"
	if  request("vBa") <> "ALL" then
	  strActBudget = strActBudget & " AND FCBA = '" & request("vBa")  & "'"
	else
	  strActBudget = strActBudget & " "
	end if
	if  request("vBatch") <> "ALL" then
	  strActBudget = strActBudget & " AND NURSERYCODE = '" & request("vBatch")  & "'"
	else
	  strActBudget = strActBudget & " "
	end if
	Set rsActBudget = Server.CreateObject("ADODB.Recordset")
	rsActBudget.Open strActBudget, connstr, adOpenStatic
	if Not rsActBudget.eof then
		vActualBibit	=	rsActBudget("ACTUALBIBIT")
		vBudgetBibit	=	rsActBudget("BUDGETBIBIT")
	end if
	rsActBudget.Close
	set rsActBudget = Nothing
%>
	<!--#include file="menutop.asp"-->
		<div id="subheader" class="centered">
			<h2>LAPORAN PEMBIAYAAN BIBITAN</h2><br/>
			<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(mm)%>&nbsp;<%=yyyy%></h4><br/>
		</div>
		<div id="header">
		<h4><%=vCompanyName%></h4>
		<h5><%=vFcbaName%></h5>
		
		
		</div>
		
		</div>
<table>
<tr><td><h4>Budget Bibit</h4></td><td>&nbsp;</td><td><h4><%=vBudgetBibit%></h4></td></tr>
<tr><td><h4>Actual Bibit</h4></td><td>&nbsp;</td><td><h4><%=vActualBibit%></h4></td></tr>
</table>
<table border="0" width="100%" id="product-table">
	<tr>
		<th colspan="3">&nbsp;</th>
		<th colspan="5">BULAN INI</th>
		<th colspan="5">S/D BULAN INI</th>
		<th colspan="2" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th colspan="3">&nbsp;</th>
		<th colspan="2">REALISASI</th>
		<th colspan="2">ANGGARAN</th>
		<th rowspan="2">%</th>
		<th colspan="2">REALISASI</th>
		<th colspan="2">ANGGARAN</th>
		<th rowspan="2">%</th>
	</tr>
	<tr>
		<th>NO</th>
		<th>CODE</th>
		<th>JENIS PEKERJAAN</th>
		<th>Total</th>
		<th>Rp/Pkk</th>
		<th>Total</th>
		<th>Rp/Pkk</th>
		<th>Total</th>
		<th>Rp/Pkk</th>
		<th>Total</th>
		<th>Rp/Pkk</th>
		<th>Total</th>
		<th>Rp/Pkk</th>
	</tr>
	
<%
Dim strPBA, strPN, strMN, rsPBA, rsPN, rsMN, rsBiayaPBA, strBiayaPBA, rsBiayaPN, strBiayaPN, rsBiayaMN, strBiayaMN
strPBA	=	" SELECT FCCODE, FCNAME FROM JOB WHERE SUBSTR(FCCODE,1,5) = '13101' GROUP BY FCCODE, FCNAME ORDER BY FCCODE ASC"
strPN	=	" SELECT FCCODE, FCNAME FROM JOB WHERE SUBSTR(FCCODE,1,5) = '13102' GROUP BY FCCODE, FCNAME ORDER BY FCCODE ASC"
strMN	=	" SELECT FCCODE, FCNAME FROM JOB WHERE SUBSTR(FCCODE,1,5) = '13103' GROUP BY FCCODE, FCNAME ORDER BY FCCODE ASC"

Set rsPBA = Server.CreateObject("ADODB.Recordset")
Set rsPN = Server.CreateObject("ADODB.Recordset")
Set rsMN = Server.CreateObject("ADODB.Recordset")
	
rsPBA.Open strPBA, connstr, adOpenStatic
rsPN.Open strPN, connstr, adOpenStatic
rsMN.Open strMN, connstr, adOpenStatic
if Not rsPBA.eof then
	Dim SubPBA, SubTyPBA, SubRppPBA, SubTyRppPBA
	SubPBA		=	0
	SubTyPBA		=	0
	SubRppPBA	=	0
	SubTyRppPBA	=	0
	While Not rsPBA.eof
%>
		<tr>
		<td align="center">&nbsp;</td>
		<td align="center"><%=rsPBA("FCCODE")%></td>
		<td align="left"><%=rsPBA("FCNAME")%></td>
<%
		strBiayaPBA	=	" SELECT NVL (SUM (AMOUNT), 0) AMOUNT, NVL (SUM (AMOUNTTY), 0) AMOUNTTY" & _
						" FROM (SELECT FCBA," & _
						" TDATE," & _
						" JOBCODE," & _
						" TARGETCODE," & _
						" AMOUNT," & _
						" 0 AMOUNTTY" & _
						" FROM COST_ACTIVITY" & _
						" WHERE TO_CHAR (TDATE, 'YYYYMM') = '" & request("vYear") & request("vMonth") & "'" & _
						" UNION" & _
						" SELECT FCBA," & _
						" TDATE," & _
						" JOBCODE," & _
						" TARGETCODE," & _
						" 0 AMOUNT," & _
						" AMOUNT AMOUNTTY" & _
						" FROM COST_ACTIVITY" & _
						" WHERE TO_CHAR (TDATE, 'YYYY') =  '" & request("vYear") & "')"
		strBiayaPBA = strBiayaPBA & " WHERE JOBCODE = '" & rsPBA("FCCODE") & "'"
		if  request("vBa") <> "ALL" then
		  strBiayaPBA = strBiayaPBA & " AND FCBA = '" & request("vBa") & "'"
		else
		  strBiayaPBA = strBiayaPBA & " "
		end if
		if  request("vBatch") <> "ALL" then
		  strBiayaPBA = strBiayaPBA & " AND TARGETCODE = '" & request("vBatch")  & "'"
		else
		  strBiayaPBA = strBiayaPBA & " "
		end if
		Set rsBiayaPBA = Server.CreateObject("ADODB.Recordset")
		rsBiayaPBA.Open strBiayaPBA, connstr, adOpenStatic
		if Not rsBiayaPBA.eof then
			While Not rsBiayaPBA.eof
%>
				<td align="right"><%=FormatNumber(rsBiayaPBA("AMOUNT"),0)%></td>
				<td align="right">
				<%if Cdbl(vActualBibit) <> 0 then%>
					<%=FormatNumber(Cdbl(rsBiayaPBA("AMOUNT"))/Cdbl(vActualBibit),0)%>
				<%else%>
					<%=FormatNumber(Cdbl(rsBiayaPBA("AMOUNT"))/1,0)%>
				<%end if%>
				</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right"><%=FormatNumber(rsBiayaPBA("AMOUNTTY"),0)%></td>
				<td align="right">
				<%if Cdbl(vActualBibit) <> 0 then%>
					<%=FormatNumber(Cdbl(rsBiayaPBA("AMOUNTTY"))/Cdbl(vActualBibit),0)%>
				<%else%>
					<%=FormatNumber(Cdbl(rsBiayaPBA("AMOUNTTY"))/1,0)%>
				<%end if%>
				</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
<%
				SubPBA	=	SubPBA	+	Cdbl(rsBiayaPBA("AMOUNT"))
				SubTyPBA	=	SubTyPBA	+	Cdbl(rsBiayaPBA("AMOUNTTY"))
				if Cdbl(vActualBibit) <> 0 then
					SubRppPBA	=	SubRppPBA	+	Cdbl(rsBiayaPBA("AMOUNT"))/Cdbl(vActualBibit)
					SubTyRppPBA	=	SubTyRppPBA	+	Cdbl(rsBiayaPBA("AMOUNTTY"))/Cdbl(vActualBibit)
				else
					SubRppPBA	=	SubRppPBA	+	Cdbl(rsBiayaPBA("AMOUNT"))/1
					SubTyRppPBA	=	SubTyRppPBA	+	Cdbl(rsBiayaPBA("AMOUNTTY"))/1
				end if
				rsBiayaPBA.MoveNext
			Wend
		end if
%>
		</tr>
<%
		SubPBA		=	SubPBA		
		SubTyPBA		=	SubTyPBA		
		SubRppPBA	=	SubRppPBA	
		SubTyRppPBA	=	SubTyRppPBA	
		rsPBA.MoveNext
	Wend
end if
%>
		<tr>
		<td colspan="3" align="center"><strong>Subtotal Persiapan Areal Bibitan</strong></td>
		<td align="right"><strong><%=FormatNumber(SubPBA,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubRppPBA,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubTyPBA,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubTyRppPBA,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		</tr>
<%
if Not rsPN.eof then
	Dim SubPN, SubTyPN, SubRppPN, SubTyRppPN
	SubPN		=	0
	SubTyPN		=	0
	SubRppPN	=	0
	SubTyRppPN	=	0
	While Not rsPN.eof
%>
		<tr>
		<td align="center">&nbsp;</td>
		<td align="center"><%=rsPN("FCCODE")%></td>
		<td align="left"><%=rsPN("FCNAME")%></td>
<%
		strBiayaPN	=	" SELECT NVL (SUM (AMOUNT), 0) AMOUNT, NVL (SUM (AMOUNTTY), 0) AMOUNTTY" & _
						" FROM (SELECT FCBA," & _
						" TDATE," & _
						" JOBCODE," & _
						" TARGETCODE," & _
						" AMOUNT," & _
						" 0 AMOUNTTY" & _
						" FROM COST_ACTIVITY" & _
						" WHERE TO_CHAR (TDATE, 'YYYYMM') = '" & request("vYear") & request("vMonth") & "'" & _
						" UNION" & _
						" SELECT FCBA," & _
						" TDATE," & _
						" JOBCODE," & _
						" TARGETCODE," & _
						" 0 AMOUNT," & _
						" AMOUNT AMOUNTTY" & _
						" FROM COST_ACTIVITY" & _
						" WHERE TO_CHAR (TDATE, 'YYYY') =  '" & request("vYear") & "')"
		strBiayaPN = strBiayaPN & " WHERE JOBCODE = '" & rsPN("FCCODE") & "'"
		if  request("vBa") <> "ALL" then
		  strBiayaPN = strBiayaPN & " AND FCBA = '" & request("vBa") & "'"
		else
		  strBiayaPN = strBiayaPN & " "
		end if
		if  request("vBatch") <> "ALL" then
		  strBiayaPN = strBiayaPN & " AND TARGETCODE = '" & request("vBatch")  & "'"
		else
		  strBiayaPN = strBiayaPN & " "
		end if
		Set rsBiayaPN = Server.CreateObject("ADODB.Recordset")
		rsBiayaPN.Open strBiayaPN, connstr, adOpenStatic
		if Not rsBiayaPN.eof then
			While Not rsBiayaPN.eof
%>
				<td align="right"><%=FormatNumber(rsBiayaPN("AMOUNT"),0)%></td>
				<td align="right">
				<%if Cdbl(vActualBibit) <> 0 then%>
					<%=FormatNumber(Cdbl(rsBiayaPN("AMOUNT"))/Cdbl(vActualBibit),0)%>
				<%else%>
					<%=FormatNumber(Cdbl(rsBiayaPN("AMOUNT"))/1,0)%>
				<%end if%>
				</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right"><%=FormatNumber(rsBiayaPN("AMOUNTTY"),0)%></td>
				<td align="right">
				<%if Cdbl(vActualBibit) <> 0 then%>
					<%=FormatNumber(Cdbl(rsBiayaPN("AMOUNTTY"))/Cdbl(vActualBibit),0)%>
				<%else%>
					<%=FormatNumber(Cdbl(rsBiayaPN("AMOUNTTY"))/1,0)%>
				<%end if%>
				</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
<%
				SubPN	=	SubPN	+	Cdbl(rsBiayaPN("AMOUNT"))
				SubTyPN	=	SubTyPN	+	Cdbl(rsBiayaPN("AMOUNTTY"))
				if Cdbl(vActualBibit) <> 0 then
					SubRppPN	=	SubRppPN	+	Cdbl(rsBiayaPN("AMOUNT"))/Cdbl(vActualBibit)
					SubTyRppPN	=	SubTyRppPN	+	Cdbl(rsBiayaPN("AMOUNTTY"))/Cdbl(vActualBibit)
				else
					SubRppPN	=	SubRppPN	+	Cdbl(rsBiayaPN("AMOUNT"))/1
					SubTyRppPN	=	SubTyRppPN	+	Cdbl(rsBiayaPN("AMOUNTTY"))/1
				end if
				rsBiayaPN.MoveNext
			Wend
		end if
%>
		</tr>
<%
		SubPN		=	SubPN		
		SubTyPN		=	SubTyPN		
		SubRppPN	=	SubRppPN	
		SubTyRppPN	=	SubTyRppPN	
		rsPN.MoveNext
	Wend
end if
%>
		<tr>
		<td colspan="3" align="center"><strong>Subtotal Pre-Nursery</strong></td>
		<td align="right"><strong><%=FormatNumber(SubPN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubRppPN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubTyPN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubTyRppPN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		</tr>

<%
if Not rsMN.eof then
	Dim SubMN, SubTyMN, SubRppMN, SubTyRppMN
	SubMN		=	0
	SubTyMN		=	0
	SubRppMN	=	0
	SubTyRppMN	=	0
	While Not rsMN.eof
%>
		<tr>
		<td align="center">&nbsp;</td>
		<td align="center"><%=rsMN("FCCODE")%></td>
		<td align="left"><%=rsMN("FCNAME")%></td>
<%
		strBiayaMN	=	" SELECT NVL (SUM (AMOUNT), 0) AMOUNT, NVL (SUM (AMOUNTTY), 0) AMOUNTTY" & _
						" FROM (SELECT FCBA," & _
						" TDATE," & _
						" JOBCODE," & _
						" TARGETCODE," & _
						" AMOUNT," & _
						" 0 AMOUNTTY" & _
						" FROM COST_ACTIVITY" & _
						" WHERE TO_CHAR (TDATE, 'YYYYMM') = '" & request("vYear") & request("vMonth") & "'" & _
						" UNION" & _
						" SELECT FCBA," & _
						" TDATE," & _
						" JOBCODE," & _
						" TARGETCODE," & _
						" 0 AMOUNT," & _
						" AMOUNT AMOUNTTY" & _
						" FROM COST_ACTIVITY" & _
						" WHERE TO_CHAR (TDATE, 'YYYY') =  '" & request("vYear") & "')"
		strBiayaMN = strBiayaMN & " WHERE JOBCODE = '" & rsMN("FCCODE") & "'"
		if  request("vBa") <> "ALL" then
		  strBiayaMN = strBiayaMN & " AND FCBA = '" & request("vBa") & "'"
		else
		  strBiayaMN = strBiayaMN & " "
		end if
		if  request("vBatch") <> "ALL" then
		  strBiayaMN = strBiayaMN & " AND TARGETCODE = '" & request("vBatch")  & "'"
		else
		  strBiayaMN = strBiayaMN & " "
		end if
		Set rsBiayaMN = Server.CreateObject("ADODB.Recordset")
		rsBiayaMN.Open strBiayaMN, connstr, adOpenStatic
		if Not rsBiayaMN.eof then
			While Not rsBiayaMN.eof
%>
				<td align="right"><%=FormatNumber(rsBiayaMN("AMOUNT"),0)%></td>
				<td align="right">
				<%if Cdbl(vActualBibit) <> 0 then%>
					<%=FormatNumber(Cdbl(rsBiayaMN("AMOUNT"))/Cdbl(vActualBibit),0)%>
				<%else%>
					<%=FormatNumber(Cdbl(rsBiayaMN("AMOUNT"))/1,0)%>
				<%end if%>
				</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right"><%=FormatNumber(rsBiayaMN("AMOUNTTY"),0)%></td>
				<td align="right">
				<%if Cdbl(vActualBibit) <> 0 then%>
					<%=FormatNumber(Cdbl(rsBiayaMN("AMOUNTTY"))/Cdbl(vActualBibit),0)%>
				<%else%>
					<%=FormatNumber(Cdbl(rsBiayaMN("AMOUNTTY"))/1,0)%>
				<%end if%>
				</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
				<td align="right">0</td>
<%
				SubMN	=	SubMN	+	Cdbl(rsBiayaMN("AMOUNT"))
				SubTyMN	=	SubTyMN	+	Cdbl(rsBiayaMN("AMOUNTTY"))
				if Cdbl(vActualBibit) <> 0 then
					SubRppMN	=	SubRppMN	+	Cdbl(rsBiayaMN("AMOUNT"))/Cdbl(vActualBibit)
					SubTyRppMN	=	SubTyRppMN	+	Cdbl(rsBiayaMN("AMOUNTTY"))/Cdbl(vActualBibit)
				else
					SubRppMN	=	SubRppMN	+	Cdbl(rsBiayaMN("AMOUNT"))/1
					SubTyRppMN	=	SubTyRppMN	+	Cdbl(rsBiayaMN("AMOUNTTY"))/1
				end if
				rsBiayaMN.MoveNext
			Wend
		end if
%>
		</tr>
<%
		SubMN		=	SubMN		
		SubTyMN		=	SubTyMN		
		SubRppMN	=	SubRppMN	
		SubTyRppMN	=	SubTyRppMN	
		rsMN.MoveNext
	Wend
end if
%>
		<tr>
		<td colspan="3" align="center"><strong>Subtotal Main-Nursery</strong></td>
		<td align="right"><strong><%=FormatNumber(SubMN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubRppMN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubTyMN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubTyRppMN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		</tr>

		<tr>
		<td colspan="3" align="center"><strong>Total Biaya Pembibitan</strong></td>
		<td align="right"><strong><%=FormatNumber(SubPBA+SubPN+SubMN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubRppPBA+SubRppPN+SubRppMN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubTyPBA+SubTyPN+SubTyMN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(SubTyRppPBA+SubTyRppPN+SubTyRppMN,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(0,0)%></strong></td>
		</tr>

</table>
<%
rsPBA.Close
rsBiayaPBA.Close
rsPN.Close
rsBiayaPN.Close
rsMN.Close
rsBiayaMN.Close

set rsPBA = Nothing
set rsBiayaPBA = Nothing
set rsPN = Nothing
set rsBiayaPN = Nothing
set rsMN = Nothing
set rsBiayaMN = Nothing
%>
</div>
</body>
</html>