<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth
Dim ScriptName, NO, cClass
Dim nSALDOAWAL, nDEBET, nKREDIT, nSALDOAKHIR

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Costbook"))
end if

Dim rsNeracaTemp, strNeracaTemp

' strNeracaTemp	=	" SELECT FCBA," & _
					' "        NYEAR," & _
					' "        NMONTH," & _
					' "        JOBCODE," & _
					' "        JOBNAME," & _
					' "        SALDOAWAL," & _
					' "        DEBET," & _
					' "        KREDIT," & _
					' "        SALDOAWAL + DEBET + KREDIT SALDOAKHIR" & _
					' "   FROM (  SELECT TB.FCBA FCBA," & _
					' "                  TB.PERIOD_YEAR NYEAR," & _
					' "                  TB.PERIOD_MONTH NMONTH," & _
					' "                  TB.FCCODE JOBCODE," & _
					' "                  TB.FCNAME JOBNAME," & _
					' "                  NVL (OPENING_BALANCE, 0) SALDOAWAL," & _
					' "                  NVL (SUM (CA.DEBET), 0) DEBET," & _
					' "                  NVL (SUM (CA.KREDIT), 0) KREDIT" & _
					' "             FROM    (SELECT FCBA," & _
					' "                             TO_NUMBER (TO_CHAR (TDATE, 'YYYY')) NYEAR," & _
					' "                             TO_NUMBER (TO_CHAR (TDATE, 'MM')) NMONTH," & _
					' "                             JOBCODE," & _
					' "                             AMOUNT DEBET," & _
					' "                             0 KREDIT" & _
					' "                        FROM COST_ACTIVITY" & _
					' "                       WHERE AMOUNT > 0" & _
					' "                      UNION ALL" & _
					' "                      SELECT FCBA," & _
					' "                             TO_NUMBER (TO_CHAR (TDATE, 'YYYY')) NYEAR," & _
					' "                             TO_NUMBER (TO_CHAR (TDATE, 'MM')) NMONTH," & _
					' "                             JOBCODE," & _
					' "                             0 DEBET," & _
					' "                             AMOUNT KREDIT" & _
					' "                        FROM COST_ACTIVITY" & _
					' "                       WHERE AMOUNT < 0) CA" & _
					' "                  FULL JOIN" & _
					' "                     TRIALBALANCE TB" & _
					' "                  ON (    CA.FCBA = TB.FCBA" & _
					' "                      AND CA.NYEAR = TB.PERIOD_YEAR" & _
					' "                      AND CA.NMONTH = TB.PERIOD_MONTH" & _
					' "                      AND CA.JOBCODE = TB.FCCODE)" & _
					' "         GROUP BY TB.FCBA," & _
					' "                  TB.PERIOD_YEAR," & _
					' "                  TB.PERIOD_MONTH," & _
					' "                  TB.FCCODE," & _
					' "                  TB.FCNAME," & _
					' "                  OPENING_BALANCE)" & _
					' " WHERE FCBA = '" & Session("FCBA") & "' AND NYEAR = " & v_txtYear & " AND NMONTH = " & v_txtmonth

' strNeracaTemp	=	strNeracaTemp	&	" ORDER BY JOBCODE"

'--New Query--
strNeracaTemp	=	" SELECT" & _
					"     TEMPDATA.FCBA" & _
					"     , NYEAR" & _
					"     , NMONTH" & _
					"     , JOBCODE" & _
					"     , FCNAME JOBNAME" & _
					"     , SUM(SALDOAWAL)SALDOAWAL" & _
					"     , SUM(DEBET) DEBET" & _
					"     , SUM(KREDIT) KREDIT" & _
					"     , SUM(SALDOAWAL)+SUM(DEBET)+SUM(KREDIT) SALDOAKHIR " & _
					" FROM (" & _
					"     SELECT" & _
					"         FCBA" & _
					"         , " & v_txtYear & " NYEAR" & _
					"         , " & v_txtmonth & " NMONTH" & _
					"         , FCCODE JOBCODE" & _
					"         , 0 SALDOAWAL" & _
					"         , 0 DEBET" & _
					"         , 0 KREDIT" & _
					"     FROM JOB" & _
					"     WHERE FCBA = '" & Session("FCBA") & "'" & _
					"     UNION ALL" & _
					"     SELECT" & _
					"         FCBA" & _
					"         , PERIOD_YEAR NYEAR" & _
					"         , PERIOD_MONTH NMONTH" & _
					"         , FCCODE JOBCODE" & _
					"         , OPENING_BALANCE SALDOAWAL" & _
					"         , 0 DEBET" & _
					"         , 0 KREDIT" & _
					"     FROM TRIALBALANCE" & _
					"     WHERE FCBA = '" & Session("FCBA") & "' AND PERIOD_YEAR = " & v_txtYear & " AND PERIOD_MONTH = " & v_txtmonth & _
					"     UNION ALL" & _
					"     SELECT" & _
					"         FCBA" & _
					"         , TO_NUMBER(TO_CHAR(TDATE,'YYYY')) NYEAR" & _
					"         , TO_NUMBER(TO_CHAR(TDATE,'MM')) NMONTH" & _
					"         , JOBCODE" & _
					"         , 0 SALDOAWAL" & _
					"         , AMOUNT DEBET" & _
					"         , 0 KREDIT" & _
					"     FROM COST_ACTIVITY" & _
					"     WHERE FCBA = '" & Session("FCBA") & "' AND TO_NUMBER(TO_CHAR(TDATE,'YYYY')) = " & v_txtYear & " AND TO_NUMBER(TO_CHAR(TDATE,'MM')) = " & v_txtmonth & " AND AMOUNT > 0" & _
					"     UNION ALL" & _
					"     SELECT" & _
					"         FCBA" & _
					"         , TO_NUMBER(TO_CHAR(TDATE,'YYYY')) NYEAR" & _
					"         , TO_NUMBER(TO_CHAR(TDATE,'MM')) NMONTH" & _
					"         , JOBCODE" & _
					"         , 0 SALDOAWAL" & _
					"         , 0 DEBET" & _
					"         , AMOUNT KREDIT" & _
					"     FROM COST_ACTIVITY" & _
					"     WHERE FCBA = '" & Session("FCBA") & "' AND TO_NUMBER(TO_CHAR(TDATE,'YYYY')) = " & v_txtYear & " AND TO_NUMBER(TO_CHAR(TDATE,'MM')) = " & v_txtmonth & " AND AMOUNT < 0" & _
					" ) TEMPDATA, JOB" & _
					" WHERE (TEMPDATA.FCBA = JOB.FCBA AND TEMPDATA.JOBCODE = JOB.FCCODE)"

strNeracaTemp	=	strNeracaTemp & " GROUP BY TEMPDATA.FCBA, NYEAR, NMONTH, JOBCODE, FCNAME"
strNeracaTemp	=	strNeracaTemp & " ORDER BY FCBA, NYEAR, NMONTH, JOBCODE"

'Response.Write strNeracaTemp

Set rsNeracaTemp = Server.CreateObject("ADODB.Recordset")
rsNeracaTemp.Open strNeracaTemp, connstr, adOpenStatic

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Neraca Sementara</title>
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

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<table id="form-table" class="centered-margin">	
					<tr>
						<td><font face="verdana" size="1">YEAR</font></td>
				<% 
					Dim mon, Yr, MinYr, MaxYr, i, MM
					mon = Month(Session("v_Costbook"))
					If len(v_txtMonth)=0 then v_txtMonth=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = Yr - 2
					MaxYr = Yr  + 2
					If len(v_txtYear)= 0 then v_txtYear = Yr
				%>
						<td width="4">
							<select name="txtYear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
							</select>
						</td>
						<td><font face="verdana" size="1">MONTH</font></td>
						<td width="2">
							<select name="txtMonth">
				<%
					For i = 1 to 12
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
				<%
					Next
				%>
							</select>
						</td>
						<td><input type="submit" value="View" /></td>
					</tr>	
				</table>
			</div>
			</form>
			<div id="header">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>NERACA SEMENTARA</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
			</div>
	<table border="0" width="100%" id="product-table">
		<tr>
			<th style="background-color:#00A300;color:#fff;">NO</th>
			<th style="background-color:#00A300;color:#fff;">KODE</th>
			<th style="background-color:#00A300;color:#fff;">PERKIRAAN</th>
			<th style="background-color:#00A300;color:#fff;">SALDO AWAL</th>
			<th style="background-color:#00A300;color:#fff;">DEBET</th>
			<th style="background-color:#00A300;color:#fff;">KREDIT</th>
			<th style="background-color:#00A300;color:#fff;">SALDO AKHIR</th>
		</tr>
	<%
	if rsNeracaTemp.bof and rsNeracaTemp.eof then
	%>
		<tr><td colspan="7">NO DATA AVAILABLE</td></tr>
	<%end if
			NO = 1
			while not rsNeracaTemp.eof
				if NO  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
	%>
		<tr class="<%=cClass%>">
			<td align="center"><%=NO%></td>
			<td align="center"><%=rsNeracaTemp("JOBCODE")%></td>
			<td align="left"><%=rsNeracaTemp("JOBNAME")%></td>
			<td align="right"><%=FormatNumber(rsNeracaTemp("SALDOAWAL"),0)%></td>
			<td align="right"><%=FormatNumber(rsNeracaTemp("DEBET"),0)%></td>
			<td align="right"><%=FormatNumber(rsNeracaTemp("KREDIT"),0)%></td>
			<td align="right"><%=FormatNumber(rsNeracaTemp("SALDOAKHIR"),0)%></td>
		</tr>
	<%
				NO = NO + 1
				nSALDOAWAL =  nSALDOAWAL + Cdbl(rsNeracaTemp("SALDOAWAL"))
				nDEBET =  nDEBET + Cdbl(rsNeracaTemp("DEBET"))
				nKREDIT =  nKREDIT + Cdbl(rsNeracaTemp("KREDIT"))
				nSALDOAKHIR =  nSALDOAKHIR+ Cdbl(rsNeracaTemp("SALDOAKHIR"))
				rsNeracaTemp.MoveNext
			Wend
	rsNeracaTemp.Close
	set rsNeracaTemp = Nothing
	%>
		<tr bgcolor="#F4F4A0">
			<td align="center" colspan="3"><strong>TOTAL</strong></td>
			<td align="right"><strong><%=FormatNumber(nSALDOAWAL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(nDEBET,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(nKREDIT,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(nSALDOAKHIR,0)%></strong></td>
		</tr>
	</table>

	<table width="90%" class="data">
		<tr> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
			<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr> 
			<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
			<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
		</tr>
		<tr> 
			<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
			<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
		</tr>
	</table>
</div>
</body>
</html>