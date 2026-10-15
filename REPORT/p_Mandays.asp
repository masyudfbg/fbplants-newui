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
Dim nKB
Dim nTOTAL


v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtGangCode	=	Request("txtGangCode")
v_txtGrade		=	Request("txtGrade")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Now)
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Now)
end if

strEmployee	=	" SELECT * FROM (" & _
				" SELECT" & _
				" AG.FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" EM.GANGCODE," & _
				" EMPLOYEECODE," & _
				" EM.FCNAME EMPLOYEENAME," & _
				" EM.PRESENTGRADE," & _
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
				" KB," & _
				" SUM(KJ+P1+P2+P3+S1+S2+H1+H2+CT+MK+ML+WH+WS+UN+KK+KB) TOTAL" & _
				" FROM (" & _
				" SELECT" & _
				" FCBA," & _
				" TO_NUMBER(TO_CHAR(FDDATE,'YYYY')) NYEAR," & _
				" TO_NUMBER(TO_CHAR(FDDATE,'MM')) NMONTH," & _
				" EMPLOYEECODE," & _
				" SUM(DECODE(ATTENDANCE,'KJ',MANDAYS,0)) KJ," & _
				" SUM(DECODE(ATTENDANCE,'P1',MANDAYS,0)) P1," & _
				" SUM(DECODE(ATTENDANCE,'P2',MANDAYS,0)) P2," & _
				" SUM(DECODE(ATTENDANCE,'P3',MANDAYS,0)) P3," & _
				" SUM(DECODE(ATTENDANCE,'S1',MANDAYS,0)) S1," & _
				" SUM(DECODE(ATTENDANCE,'S2',MANDAYS,0)) S2," & _
				" SUM(DECODE(ATTENDANCE,'H1',MANDAYS,0)) H1," & _
				" SUM(DECODE(ATTENDANCE,'H2',MANDAYS,0)) H2," & _
				" SUM(DECODE(ATTENDANCE,'CT',MANDAYS,0)) CT," & _
				" SUM(DECODE(ATTENDANCE,'MK',MANDAYS,0)) MK," & _
				" SUM(DECODE(ATTENDANCE,'ML',MANDAYS,0)) ML," & _
				" SUM(DECODE(ATTENDANCE,'WH',MANDAYS,0)) WH," & _
				" SUM(DECODE(ATTENDANCE,'WS',MANDAYS,0)) WS," & _
				" SUM(DECODE(ATTENDANCE,'UN',MANDAYS,0)) UN," & _
				" SUM(DECODE(ATTENDANCE,'KK',MANDAYS,0)) KK," & _
				" SUM(DECODE(ATTENDANCE,'KB',MANDAYS,0)) KB" & _
				" FROM ATTENDANCE_GAD" & _
				" GROUP BY FCBA,TO_NUMBER(TO_CHAR(FDDATE,'YYYY')), TO_NUMBER(TO_CHAR(FDDATE,'MM')), EMPLOYEECODE" & _
				" ) AG LEFT JOIN EMPLOYEE EM ON ( AG.FCBA = EM.FCBA AND AG.EMPLOYEECODE = EM.FCCODE)" & _
				" GROUP BY" & _
				" AG.FCBA," & _
				" NYEAR," & _
				" NMONTH," & _
				" EM.GANGCODE," & _
				" EMPLOYEECODE," & _
				" EM.FCNAME," & _
				" EM.PRESENTGRADE," & _
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
				" KB" & _
				" )" & _
				" WHERE PRESENTGRADE IN ('KHL') AND FCBA = '" & Session ("FCBA") & "'"

if  LEN(TRIM(v_txtYear)) > 0 then
	strEmployee	=	strEmployee & " AND NYEAR = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " AND NMONTH = " & v_txtMonth
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strEmployee	=	strEmployee & " AND GANGCODE = '" & v_txtGangCode & "'"
end if

strEmployee	=	strEmployee & " ORDER BY EMPLOYEECODE ASC"
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Rekap Absen KHL</title>
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
		<h2>LAPORAN HARI KERJA KHL</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
		<td><select name="txtGangCode" ><option value=""></option>
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
		<th>KB</th>
		<th>TOTAL</th>
	</tr>

<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic
if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td colspan="21">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		nKJ = 0
		nP1 = 0
		nP2 = 0
		nP3 = 0
		nS1 = 0
		nS2 = 0
		nH1 = 0
		nH2 = 0
		nCT = 0
		nMK = 0
		nML = 0
		nWH = 0
		nWS = 0
		nUN = 0
		nKK = 0
		nKB = 0
		nTOTAL = 0

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
				<td align="right"><%=FbpNum2(rsEmployee("KJ"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("P1"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("P2"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("P3"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("S1"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("S2"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("H1"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("H2"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("CT"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("MK"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("ML"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("WH"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("WS"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("UN"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("KK"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("KB"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("TOTAL"))%></td>
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
			nKB = nKB + Cdbl(rsEmployee("KB"))
			nTOTAL = nTOTAL + Cdbl(rsEmployee("TOTAL"))

			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="4">TOTAL</td>
				<td align="right"><%=FbpNum2(nKJ)%></td>
				<td align="right"><%=FbpNum2(nP1)%></td>
				<td align="right"><%=FbpNum2(nP2)%></td>
				<td align="right"><%=FbpNum2(nP3)%></td>
				<td align="right"><%=FbpNum2(nS1)%></td>
				<td align="right"><%=FbpNum2(nS2)%></td>
				<td align="right"><%=FbpNum2(nH1)%></td>
				<td align="right"><%=FbpNum2(nH2)%></td>
				<td align="right"><%=FbpNum2(nCT)%></td>
				<td align="right"><%=FbpNum2(nMK)%></td>
				<td align="right"><%=FbpNum2(nML)%></td>
				<td align="right"><%=FbpNum2(nWH)%></td>
				<td align="right"><%=FbpNum2(nWS)%></td>
				<td align="right"><%=FbpNum2(nUN)%></td>
				<td align="right"><%=FbpNum2(nKK)%></td>
				<td align="right"><%=FbpNum2(nKB)%></td>
				<td align="right"><%=FbpNum2(nTOTAL)%></td>
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
