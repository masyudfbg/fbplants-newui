<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
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
<title>iPla'S Reporting - Analisa Biaya Tidak Langsung</title>
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
			<th>&nbsp;</th>
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
        <Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=Bulan(i)%></option>
<%
    Next
%>
					</select>
			</td>
			<td>
				<select name="vBa">
				<option value="ALL" <%if request("vBa")="ALL" then%> selected <%End if%>>::ALL::</option>
				<%
				'reDim rsBa
				Set rsBa = Server.CreateObject("ADODB.Recordset")
				rsBa.CursorLocation = adUseClient
				strsql =  " SELECT FCCODE, FCNAME FROM BUSINESSUNIT ORDER BY FCCODE ASC"
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
			<td><input type="submit" value="View" /></td>
				</form>
		</tr>
	</table>
	<!--#include file="menutop.asp"-->
		<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		</div>
<%
Dim strha, rsha, LuasHa
strha = " SELECT nvl(SUM (hectarageplanted),0) ha FROM field"
		
if  request("vBa") <> "ALL" then
  strha = strha & " WHERE FCBA = '" & request("vBa")  & "'"
else
  strha = strha & " "
end if
		
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strha, connstr, adOpenStatic
if Not rsha.eof then
	if Cdbl(rsha("ha")) > 0 then
		LuasHa	=	rsha("ha")
	elseif Cdbl(rsha("ha")) = 0 then
		LuasHa	=	1
	end if
end if
rsha.Close
Set rsha = Nothing
%>
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
		<th rowspan="3">%</th>
		<th colspan="2">REALISASI</th>
		<th colspan="2">ANGGARAN</th>
		<th rowspan="3">%</th>
	</tr>
	<tr>
		<th colspan="3" align="right">Luas Ha</th>
		<th colspan="2"><%=LuasHa%></th>
		<th colspan="2"><%=LuasHa%></th>
		<th colspan="2"><%=LuasHa%></th>
		<th colspan="2"><%=LuasHa%></th>
		<th colspan="2"><%=LuasHa%></th>
	</tr>
	<tr>
		<th>NO</th>
		<th>CODE</th>
		<th>JENIS PEKERJAAN</th>
		<th>Total</th>
		<th>Rp/ha</th>
		<th>Total</th>
		<th>Rp/ha</th>
		<th>Total</th>
		<th>Rp/ha</th>
		<th>Total</th>
		<th>Rp/ha</th>
		<th>Total</th>
		<th>Rp/ha</th>
	</tr>
<%
Dim rsBTL, strBTL, strGroupBy, rowno, rsalokasiBTL, stralokasiBTL
rowno		=	0
strBTL		=	" SELECT " & _
				" JOBCODE," & _
				" JOBNAME," & _
				" SUM (AMOUNT) AMOUNT," & _
				" SUM (AMOUNTTY) AMOUNTTY" & _
				" FROM (SELECT FCBA," & _
				" JOBCODE," & _
				" JOBNAME," & _
				" AMOUNT," & _
				" 0 AMOUNTTY" & _
				" FROM ( SELECT J.FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" JOBCODE," & _
				" J.FCNAME JOBNAME," & _
				" SUM (AMOUNT) AMOUNT" & _
				" FROM ( SELECT FCBA," & _
				" TO_NUMBER (TO_CHAR (TDATE, 'YYYY')) NYEAR," & _
				" TO_NUMBER (TO_CHAR (TDATE, 'MM')) NMONTH," & _
				" JOBCODE," & _
				" SUM (AMOUNT) AMOUNT" & _
				" FROM ( SELECT FCBA," & _
				" TDATE," & _
				" JOBCODE," & _
				" SUM (AMOUNT) AMOUNT" & _
				" FROM COST_ACTIVITY" & _
				" GROUP BY FCBA, TDATE, JOBCODE" & _
				" ORDER BY FCBA, TDATE, JOBCODE)" & _
				" WHERE SUBSTR (JOBCODE, 1, 2) = '64'" & _
				" GROUP BY FCBA," & _
				" TO_NUMBER (TO_CHAR (TDATE, 'YYYY'))," & _
				" TO_NUMBER (TO_CHAR (TDATE, 'MM'))," & _
				" JOBCODE) CA" & _
				" INNER JOIN" & _
				" JOB J" & _
				" ON (CA.FCBA = J.FCBA AND CA.JOBCODE = J.FCCODE)" & _
				" GROUP BY J.FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" JOBCODE," & _
				" J.FCNAME)" & _
				" WHERE NYEAR || NMONTH = '" & request("vYear") & request("vMonth") & "'" & _
				" UNION" & _
				" SELECT FCBA," & _
				" JOBCODE," & _
				" JOBNAME," & _
				" 0 AMOUNT," & _
				" AMOUNT" & _
				" FROM ( SELECT J.FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" JOBCODE," & _
				" J.FCNAME JOBNAME," & _
				" SUM (AMOUNT) AMOUNT" & _
				" FROM ( SELECT FCBA," & _
				" TO_NUMBER (TO_CHAR (TDATE, 'YYYY')) NYEAR," & _
				" TO_NUMBER (TO_CHAR (TDATE, 'MM')) NMONTH," & _
				" JOBCODE," & _
				" SUM (AMOUNT) AMOUNT" & _
				" FROM ( SELECT FCBA," & _
				" TDATE," & _
				" JOBCODE," & _
				" SUM (AMOUNT) AMOUNT" & _
				" FROM COST_ACTIVITY" & _
				" GROUP BY FCBA, TDATE, JOBCODE" & _
				" ORDER BY FCBA, TDATE, JOBCODE)" & _
				" WHERE SUBSTR (JOBCODE, 1, 2) = '64'" & _
				" GROUP BY FCBA," & _
				" TO_NUMBER (TO_CHAR (TDATE, 'YYYY'))," & _
				" TO_NUMBER (TO_CHAR (TDATE, 'MM'))," & _
				" JOBCODE) CA" & _
				" INNER JOIN" & _
				" JOB J" & _
				" ON (CA.FCBA = J.FCBA AND CA.JOBCODE = J.FCCODE)" & _
				" GROUP BY J.FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" JOBCODE," & _
				" J.FCNAME)" & _
				" WHERE NYEAR || NMONTH <= '" & request("vYear") & request("vMonth") & "')" & _
				" WHERE JOBCODE NOT IN ('6449800')"

				
stralokasiBTL	=	" SELECT " & _
					" JOBCODE," & _
					" JOBNAME," & _
					" SUM (AMOUNT) AMOUNT," & _
					" SUM (AMOUNTTY) AMOUNTTY" & _
					" FROM (SELECT FCBA," & _
					" JOBCODE," & _
					" JOBNAME," & _
					" AMOUNT," & _
					" 0 AMOUNTTY" & _
					" FROM ( SELECT J.FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" JOBCODE," & _
					" J.FCNAME JOBNAME," & _
					" SUM (AMOUNT) AMOUNT" & _
					" FROM ( SELECT FCBA," & _
					" TO_NUMBER (TO_CHAR (TDATE, 'YYYY')) NYEAR," & _
					" TO_NUMBER (TO_CHAR (TDATE, 'MM')) NMONTH," & _
					" JOBCODE," & _
					" SUM (AMOUNT) AMOUNT" & _
					" FROM ( SELECT FCBA," & _
					" TDATE," & _
					" JOBCODE," & _
					" SUM (AMOUNT) AMOUNT" & _
					" FROM COST_ACTIVITY" & _
					" GROUP BY FCBA, TDATE, JOBCODE" & _
					" ORDER BY FCBA, TDATE, JOBCODE)" & _
					" WHERE SUBSTR (JOBCODE, 1, 2) = '64'" & _
					" GROUP BY FCBA," & _
					" TO_NUMBER (TO_CHAR (TDATE, 'YYYY'))," & _
					" TO_NUMBER (TO_CHAR (TDATE, 'MM'))," & _
					" JOBCODE) CA" & _
					" INNER JOIN" & _
					" JOB J" & _
					" ON (CA.FCBA = J.FCBA AND CA.JOBCODE = J.FCCODE)" & _
					" GROUP BY J.FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" JOBCODE," & _
					" J.FCNAME)" & _
					" WHERE NYEAR || NMONTH = '" & request("vYear") & request("vMonth") & "'" & _
					" UNION" & _
					" SELECT FCBA," & _
					" JOBCODE," & _
					" JOBNAME," & _
					" 0 AMOUNT," & _
					" AMOUNT" & _
					" FROM ( SELECT J.FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" JOBCODE," & _
					" J.FCNAME JOBNAME," & _
					" SUM (AMOUNT) AMOUNT" & _
					" FROM ( SELECT FCBA," & _
					" TO_NUMBER (TO_CHAR (TDATE, 'YYYY')) NYEAR," & _
					" TO_NUMBER (TO_CHAR (TDATE, 'MM')) NMONTH," & _
					" JOBCODE," & _
					" SUM (AMOUNT) AMOUNT" & _
					" FROM ( SELECT FCBA," & _
					" TDATE," & _
					" JOBCODE," & _
					" SUM (AMOUNT) AMOUNT" & _
					" FROM COST_ACTIVITY" & _
					" GROUP BY FCBA, TDATE, JOBCODE" & _
					" ORDER BY FCBA, TDATE, JOBCODE)" & _
					" WHERE SUBSTR (JOBCODE, 1, 2) = '64'" & _
					" GROUP BY FCBA," & _
					" TO_NUMBER (TO_CHAR (TDATE, 'YYYY'))," & _
					" TO_NUMBER (TO_CHAR (TDATE, 'MM'))," & _
					" JOBCODE) CA" & _
					" INNER JOIN" & _
					" JOB J" & _
					" ON (CA.FCBA = J.FCBA AND CA.JOBCODE = J.FCCODE)" & _
					" GROUP BY J.FCBA," & _
					" NYEAR," & _
					" NMONTH," & _
					" JOBCODE," & _
					" J.FCNAME)" & _
					" WHERE NYEAR || NMONTH <= '" & request("vYear") & request("vMonth") & "')" & _
					" WHERE JOBCODE IN ('6449800')"


if  request("vBa") <> "ALL" then
  strBTL = strBTL & " AND FCBA = '" & request("vBa")  & "'"
  stralokasiBTL = stralokasiBTL & " AND FCBA = '" & request("vBa")  & "'"
else
  strBTL = strBTL & " "
  stralokasiBTL = stralokasiBTL & " "
end if

strGroupBy	=	" GROUP BY JOBCODE, JOBNAME"

strBTL		=	strBTL	&	strGroupBy
stralokasiBTL		=	stralokasiBTL	&	strGroupBy

Set rsBTL = Server.CreateObject("ADODB.Recordset")
rsBTL.Open strBTL, connstr, adOpenStatic
if Not rsBTL.eof then
Dim TotalBTL, TotalTyBTL
TotalBTL = 0
TotalTyBTL = 0
	While Not rsBTL.eof
			rowno	=	rowno	+	1
%>
	<tr>
		<td align="center"><%=rowno%></td>
		<td><%=rsBTL("JOBCODE")%></td>
		<td><%=rsBTL("JOBNAME")%></td>
		<td align="right"><%=FormatNumber(rsBTL("AMOUNT"),2)%></td>
		<td align="right"><%=FormatNumber((Cdbl(rsBTL("AMOUNT"))/Cdbl(LuasHa)),2)%></td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right"><%=FormatNumber(rsBTL("AMOUNTTY"),2)%></td>
		<td align="right"><%=FormatNumber((Cdbl(rsBTL("AMOUNTTY"))/Cdbl(LuasHa)),2)%></td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
	</tr>
<%
		TotalBTL = TotalBTL + Cdbl(rsBTL("AMOUNT"))
		TotalTyBTL	=	TotalTyBTL	+ Cdbl(rsBTL("AMOUNTTY"))
		rsBTL.MoveNext
	Wend
end if
%>
	<tr>
		<td colspan="3" align="center"><strong>Total Biaya Tak Langsung (A)</strong></td>
		<td align="right"><strong><%=FormatNumber(TotalBTL,2)%></strong></td>		
		<td align="right"><strong><%=FormatNumber((Cdbl(TotalBTL)/Cdbl(LuasHa)),2)%></strong></td>	
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong><%=FormatNumber(TotalTyBTL,2)%></strong></td>		
		<td align="right"><strong><%=FormatNumber((Cdbl(TotalTyBTL)/Cdbl(LuasHa)),2)%></strong></td>	
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
	</tr>
<%
Set rsalokasiBTL = Server.CreateObject("ADODB.Recordset")
rsalokasiBTL.Open stralokasiBTL, connstr, adOpenStatic
if Not rsalokasiBTL.eof then
Dim TotalalokasiBTL, TotalTyalokasiBTL
TotalalokasiBTL = 0
TotalTyalokasiBTL = 0
	While Not rsalokasiBTL.eof
		rowno	=	rowno	+	1
%>
		<tr>
		<td align="center"><%=rowno%></td>
		<td><%=rsalokasiBTL("JOBCODE")%></td>
		<td><%=rsalokasiBTL("JOBNAME")%></td>
		<td align="right"><%=FormatNumber(rsalokasiBTL("AMOUNT"),2)%></td>
		<td align="right"><%=FormatNumber((Cdbl(rsalokasiBTL("AMOUNT"))/Cdbl(LuasHa)),2)%></td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right"><%=FormatNumber(rsalokasiBTL("AMOUNTTY"),2)%></td>
		<td align="right"><%=FormatNumber((Cdbl(rsalokasiBTL("AMOUNTTY"))/Cdbl(LuasHa)),2)%></td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
		<td align="right">0</td>
	</tr>
<%
		TotalalokasiBTL = TotalalokasiBTL + Cdbl(rsalokasiBTL("AMOUNT"))
		TotalTyalokasiBTL = TotalTyalokasiBTL + Cdbl(rsalokasiBTL("AMOUNTTY"))
		rsalokasiBTL.MoveNext
	Wend
end if
%>
	<tr>
		<td colspan="3" align="center"><strong>Total Alokasi Biaya Tak Langsung (B)</strong></td>
		<td align="right"><strong><%=FormatNumber(TotalalokasiBTL,2)%></strong></td>		
		<td align="right"><strong><%=FormatNumber((Cdbl(TotalalokasiBTL)/Cdbl(LuasHa)),2)%></strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong><%=FormatNumber(TotalTyalokasiBTL,2)%></strong></td>		
		<td align="right"><strong><%=FormatNumber((Cdbl(TotalTyalokasiBTL)/Cdbl(LuasHa)),2)%></strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
	</tr>
	<tr>
		<td colspan="3" align="center"><strong>Total Keseluruhan (A) + (B)</td>
		<td align="right"><strong><%=FormatNumber(Cdbl(TotalBTL)+Cdbl(TotalalokasiBTL),2)%></strong></td>		
		<td align="right"><strong><%=FormatNumber(((Cdbl(TotalBTL)+Cdbl(TotalalokasiBTL))/Cdbl(LuasHa)),2)%></strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong><%=FormatNumber(Cdbl(TotalTyBTL)+Cdbl(TotalTyalokasiBTL),2)%></strong></td>		
		<td align="right"><strong><%=FormatNumber(((Cdbl(TotalTyBTL)+Cdbl(TotalTyalokasiBTL))/Cdbl(LuasHa)),2)%></strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
	</tr>
</table>
<!--#include file="bottom.Asp"-->