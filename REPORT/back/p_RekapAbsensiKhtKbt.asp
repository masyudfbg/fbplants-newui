<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade
Dim ScriptName
Dim rsEmployee, strEmployee, rsAtt, strAtt, rsKJ, strKJ
Dim NO, cClass, totalDay
Dim nKJ
Dim nP1
Dim nP2
Dim nP3
Dim nS1
Dim nS2
Dim nH1
Dim nH2
Dim nCT
Dim nMK
Dim nML
Dim nWH
Dim nWS
Dim nUN
Dim nKK
Dim nTOTAL

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtGangCode	=	Request("txtGangCode")
v_txtGrade		=	Request("txtGrade")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Employee"))
end if

strEmployee	=	" SELECT * FROM (" & _
				" SELECT FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" GANGCODE," & _
				" EMPLOYEECODE," & _
				" EMPLOYEENAME," & _
				" PRESENTGRADE," & _
				" KJ," & _
				" P1," & _
				" P2," & _
				" P3," & _
				" S1," & _
				" S2," & _
				" H1," & _
				" H2," & _
				" CT," & _
				" MK," & _
				" ML," & _
				" WH," & _
				" WS," & _
				" UN," & _
				" KK," & _
				" SUM (KJ" & _
				" + P1" & _
				" + P2" & _
				" + P3" & _
				" + S1" & _
				" + S2" & _
				" + H1" & _
				" + H2" & _
				" + CT" & _
				" + MK" & _
				" + ML" & _
				" + WH" & _
				" + WS" & _
				" + UN" & _
				" + KK)" & _
				" TOTAL" & _
				" FROM (" & _
				" SELECT AG.FCBA," & _
				" TO_NUMBER(TO_CHAR(FDDATE,'YYYY')) NYEAR," & _
				" TO_NUMBER(TO_CHAR(FDDATE,'MM')) NMONTH," & _
				" EM.GANGCODE," & _
				" EMPLOYEECODE," & _
				" EM.FCNAME EMPLOYEENAME," & _
				" EM.PRESENTGRADE," & _
				" SUM(KJ) KJ," & _
				" SUM(P1) P1," & _
				" SUM(P2) P2," & _
				" SUM(P3) P3," & _
				" SUM(S1) S1," & _
				" SUM(S2) S2," & _
				" SUM(H1) H1," & _
				" SUM(H2) H2," & _
				" SUM(CT) CT," & _
				" SUM(MK) MK," & _
				" SUM(ML) ML," & _
				" SUM(WH) WH," & _
				" SUM(WS) WS," & _
				" SUM(UN) UN," & _
				" SUM(KK) KK" & _
				" FROM (" & _
				" SELECT FCBA," & _
				" FDDATE," & _
				" EMPLOYEECODE," & _
				" SUM (DECODE (ATTENDANCE, 'KJ', MANDAYS, 0)) KJ," & _
				" SUM (DECODE (ATTENDANCE, 'P1', 1, 0)) P1," & _
				" SUM (DECODE (ATTENDANCE, 'P2', 1, 0)) P2," & _
				" SUM (DECODE (ATTENDANCE, 'P3', 1, 0)) P3," & _
				" SUM (DECODE (ATTENDANCE, 'S1', 1, 0)) S1," & _
				" SUM (DECODE (ATTENDANCE, 'S2', 1, 0)) S2," & _
				" SUM (DECODE (ATTENDANCE, 'H1', 1, 0)) H1," & _
				" SUM (DECODE (ATTENDANCE, 'H2', 1, 0)) H2," & _
				" SUM (DECODE (ATTENDANCE, 'CT', 1, 0)) CT," & _
				" SUM (DECODE (ATTENDANCE, 'MK', 1, 0)) MK," & _
				" SUM (DECODE (ATTENDANCE, 'ML', 1, 0)) ML," & _
				" SUM (DECODE (ATTENDANCE, 'WH', MANDAYS, 0)) WH," & _
				" SUM (DECODE (ATTENDANCE, 'WS', MANDAYS, 0)) WS," & _
				" SUM (DECODE (ATTENDANCE, 'UN', 1, 0)) UN," & _
				" SUM (DECODE (ATTENDANCE, 'KK', MANDAYS, 0)) KK" & _
				" FROM ATTENDANCE_GAD GROUP BY FCBA, FDDATE, EMPLOYEECODE" & _
				" ) AG LEFT JOIN EMPLOYEE EM ON (AG.FCBA = EM.FCBA AND AG.EMPLOYEECODE = EM.FCCODE)" & _
				" GROUP BY AG.FCBA," & _
				" TO_NUMBER(TO_CHAR(FDDATE,'YYYY'))," & _
				" TO_NUMBER(TO_CHAR(FDDATE,'MM'))," & _
				" EM.GANGCODE," & _
				" EMPLOYEECODE," & _
				" EM.FCNAME," & _
				" EM.PRESENTGRADE" & _
				" )" & _
				" GROUP BY FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" GANGCODE," & _
				" EMPLOYEECODE," & _
				" EMPLOYEENAME," & _
				" PRESENTGRADE," & _
				" KJ," & _
				" P1," & _
				" P2," & _
				" P3," & _
				" S1," & _
				" S2," & _
				" H1," & _
				" H2," & _
				" CT," & _
				" MK," & _
				" ML," & _
				" WH," & _
				" WS," & _
				" UN," & _
				" KK" & _
				" )" & _
				" WHERE PRESENTGRADE IN ('KHT','KBT') AND FCBA = '" & Session ("FCBA") & "'"

if  LEN(TRIM(v_txtYear)) > 0 then
	strEmployee	=	strEmployee & " AND NYEAR = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " AND NMONTH = " & v_txtMonth
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strEmployee	=	strEmployee & " AND GANGCODE = '" & v_txtGangCode & "'"
end if
if  LEN(TRIM(v_txtGrade)) > 0 then
	strEmployee	=	strEmployee & " AND PRESENTGRADE = '" & v_txtGrade & "'"
end if

strEmployee	=	strEmployee & " ORDER BY EMPLOYEECODE ASC"
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Rekap Absensi KHT KBT</title>
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
		<h2>LAPORAN HARI KERJA KHT dan KBT</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
		<td><font face="verdana" size="1">GANGCODE</font></td><td><select name="txtGangCode" ><option value=""></option>
		<%
		Dim rsGangCode, strGangCode
		strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
		
		Set rsGangCode = Server.CreateObject("ADODB.Recordset")
		rsGangCode.Open strGangCode, connstr, adOpenStatic 
		
		if not rsGangCode.eof then
		Do While Not rsGangCode.eof
		if rsGangCode("FCCODE") = v_txtGangCode then%>
		        <option value="<%=rsGangCode("FCCODE")%>" Selected><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option> 
		     <%Else%>
		        <option value="<%=rsGangCode("FCCODE")%>"><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option>  
		    <%End if
			rsGangCode.movenext
		loop
		end if
		rsGangCode.Close
		set rsGangCode = Nothing
		%>
		</select></td>
		<td><font face="verdana" size="1">STATUS</font></td><td><select name="txtGrade" ><option value=""></option>
		<%
		Dim rsGrade, strGrade
		strGrade	=	" SELECT 'KHT' FCCODE, 'Karyawan Harian Tetap' FCNAME FROM DUAL" & _
						" UNION" & _
						" SELECT 'KBT' FCCODE, 'Karyawan Bulanan Tetap' FCNAME FROM DUAL"
		
		Set rsGrade = Server.CreateObject("ADODB.Recordset")
		rsGrade.Open strGrade, connstr, adOpenStatic 
		
		if not rsGrade.eof then
		Do While Not rsGrade.eof
		if rsGrade("FCCODE") = v_txtGrade then%>
		        <option value="<%=rsGrade("FCCODE")%>" Selected><%=rsGrade("FCCODE")%>-<%=rsGrade("FCNAME")%></option> 
		     <%Else%>
		        <option value="<%=rsGrade("FCCODE")%>"><%=rsGrade("FCCODE")%>-<%=rsGrade("FCNAME")%></option>  
		    <%End if
			rsGrade.movenext
		loop
		end if
		rsGrade.Close
		set rsGrade = Nothing
		%>
		</select></td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th>NO</th>
		<th>NO PEKERJA</th>
		<th>NAMA PEKERJA</th>
		<th>STATUS</th>
		<th>KJ</th>
		<th>P1</th>
		<th>P2</th>
		<th>P3</th>
		<th>S1</th>
		<th>S2</th>
		<th>H1</th>
		<th>H2</th>
		<th>CT</th>
		<th>MK</th>
		<th>ML</th>
		<th>WH</th>
		<th>WS</th>
		<th>UN</th>
		<th>KK</th>
		<th>TOTAL</th>
	</tr>

<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic
if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td>NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		while not rsEmployee.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="center"><%=rsEmployee("EMPLOYEECODE")%></td>
				<td align="left"><%=rsEmployee("EMPLOYEENAME")%></td>
				<td align="center"><%=rsEmployee("PRESENTGRADE")%></td>
				<td align="right"><%=rsEmployee("KJ")%></td>
				<td align="right"><%=rsEmployee("P1")%></td>
				<td align="right"><%=rsEmployee("P2")%></td>
				<td align="right"><%=rsEmployee("P3")%></td>
				<td align="right"><%=rsEmployee("S1")%></td>
				<td align="right"><%=rsEmployee("S2")%></td>
				<td align="right"><%=rsEmployee("H1")%></td>
				<td align="right"><%=rsEmployee("H2")%></td>
				<td align="right"><%=rsEmployee("CT")%></td>
				<td align="right"><%=rsEmployee("MK")%></td>
				<td align="right"><%=rsEmployee("ML")%></td>
				<td align="right"><%=rsEmployee("WH")%></td>
				<td align="right"><%=rsEmployee("WS")%></td>
				<td align="right"><%=rsEmployee("UN")%></td>
				<td align="right"><%=rsEmployee("KK")%></td>
				<td align="right"><%=rsEmployee("TOTAL")%></td>
<%
			NO = NO + 1
			nKJ = nKJ + Cdbl(rsEmployee("KJ"))
			nP1 = nP1 + Cdbl(rsEmployee("P1"))
			nP2 = nP2 + Cdbl(rsEmployee("P2"))
			nP3 = nP3 + Cdbl(rsEmployee("P3"))
			nS1 = nS1 + Cdbl(rsEmployee("S1"))
			nS2 = nS2 + Cdbl(rsEmployee("S2"))
			nH1 = nH1 + Cdbl(rsEmployee("H1"))
			nH2 = nH2 + Cdbl(rsEmployee("H2"))
			nCT = nCT + Cdbl(rsEmployee("CT"))
			nMK = nMK + Cdbl(rsEmployee("MK"))
			nML = nML + Cdbl(rsEmployee("ML"))
			nWH = nWH + Cdbl(rsEmployee("WH"))
			nWS = nWS + Cdbl(rsEmployee("WS"))
			nUN = nUN + Cdbl(rsEmployee("UN"))
			nKK = nKK + Cdbl(rsEmployee("KK"))
			nTOTAL = nTOTAL + Cdbl(rsEmployee("TOTAL"))
			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="4">TOTAL</td>
				<td align="right"><%=nKJ%></td>
				<td align="right"><%=nP1%></td>
				<td align="right"><%=nP2%></td>
				<td align="right"><%=nP3%></td>
				<td align="right"><%=nS1%></td>
				<td align="right"><%=nS2%></td>
				<td align="right"><%=nH1%></td>
				<td align="right"><%=nH2%></td>
				<td align="right"><%=nCT%></td>
				<td align="right"><%=nMK%></td>
				<td align="right"><%=nML%></td>
				<td align="right"><%=nWH%></td>
				<td align="right"><%=nWS%></td>
				<td align="right"><%=nUN%></td>
				<td align="right"><%=nKK%></td>
				<td align="right"><%=nTOTAL%></td>
			</tr>
</table>

   <table width="90%" class="data">
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
   </body>
   </html>
