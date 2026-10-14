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
Dim NO, cClass, nday

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

strEmployee	=	" SELECT AG.FCBA FCBA," & _
				" AG.NYEAR NYEAR," & _
				" AG.NMONTH NMONTH," & _
				" AG.EMPLOYEECODE," & _
				" EM.FCNAME EMPLOYEENAME," & _
				" EM.PRESENTGRADE STATUS," & _
				" EM.GANGCODE GANGCODE," & _
				" DAY_1," & _
				" DAY_2," & _
				" DAY_3," & _
				" DAY_4," & _
				" DAY_5," & _
				" DAY_6," & _
				" DAY_7," & _
				" DAY_8," & _
				" DAY_9," & _
				" DAY_10," & _
				" DAY_11," & _
				" DAY_12," & _
				" DAY_13," & _
				" DAY_14," & _
				" DAY_15," & _
				" DAY_16," & _
				" DAY_17," & _
				" DAY_18," & _
				" DAY_19," & _
				" DAY_20," & _
				" DAY_21," & _
				" DAY_22," & _
				" DAY_23," & _
				" DAY_24," & _
				" DAY_25," & _
				" DAY_26," & _
				" DAY_27," & _
				" DAY_28," & _
				" DAY_29," & _
				" DAY_30," & _
				" DAY_31" & _
				" FROM ( SELECT FCBA," & _
				" TO_NUMBER (TO_CHAR (FDDATE, 'YYYY')) NYEAR," & _
				" TO_NUMBER (TO_CHAR (FDDATE, 'MM')) NMONTH," & _
				" EMPLOYEECODE," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 1, ATTENDANCE," & _
				" NULL))" & _
				" DAY_1," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 2, ATTENDANCE," & _
				" NULL))" & _
				" DAY_2," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 3, ATTENDANCE," & _
				" NULL))" & _
				" DAY_3," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 4, ATTENDANCE," & _
				" NULL))" & _
				" DAY_4," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 5, ATTENDANCE," & _
				" NULL))" & _
				" DAY_5," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 6, ATTENDANCE," & _
				" NULL))" & _
				" DAY_6," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 7, ATTENDANCE," & _
				" NULL))" & _
				" DAY_7," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 8, ATTENDANCE," & _
				" NULL))" & _
				" DAY_8," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 9, ATTENDANCE," & _
				" NULL))" & _
				" DAY_9," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 10, ATTENDANCE," & _
				" NULL))" & _
				" DAY_10," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 11, ATTENDANCE," & _
				" NULL))" & _
				" DAY_11," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 12, ATTENDANCE," & _
				" NULL))" & _
				" DAY_12," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 13, ATTENDANCE," & _
				" NULL))" & _
				" DAY_13," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 14, ATTENDANCE," & _
				" NULL))" & _
				" DAY_14," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 15, ATTENDANCE," & _
				" NULL))" & _
				" DAY_15," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 16, ATTENDANCE," & _
				" NULL))" & _
				" DAY_16," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 17, ATTENDANCE," & _
				" NULL))" & _
				" DAY_17," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 18, ATTENDANCE," & _
				" NULL))" & _
				" DAY_18," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 19, ATTENDANCE," & _
				" NULL))" & _
				" DAY_19," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 20, ATTENDANCE," & _
				" NULL))" & _
				" DAY_20," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 21, ATTENDANCE," & _
				" NULL))" & _
				" DAY_21," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 22, ATTENDANCE," & _
				" NULL))" & _
				" DAY_22," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 23, ATTENDANCE," & _
				" NULL))" & _
				" DAY_23," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 24, ATTENDANCE," & _
				" NULL))" & _
				" DAY_24," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 25, ATTENDANCE," & _
				" NULL))" & _
				" DAY_25," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 26, ATTENDANCE," & _
				" NULL))" & _
				" DAY_26," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 27, ATTENDANCE," & _
				" NULL))" & _
				" DAY_27," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 28, ATTENDANCE," & _
				" NULL))" & _
				" DAY_28," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 29, ATTENDANCE," & _
				" NULL))" & _
				" DAY_29," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 30, ATTENDANCE," & _
				" NULL))" & _
				" DAY_30," & _
				" MAX (" & _
				" DECODE (TO_NUMBER (TO_CHAR (FDDATE, 'DD'))," & _
				" 31, ATTENDANCE," & _
				" NULL))" & _
				" DAY_31" & _
				" FROM ( SELECT FCBA," & _
				" FDDATE," & _
				" EMPLOYEECODE," & _
				" ATTENDANCE" & _
				" FROM ATTENDANCE_GAD" & _
				" WHERE ATTENDANCE NOT IN ('UN', 'KB')" & _
				" GROUP BY FCBA," & _
				" FDDATE," & _
				" EMPLOYEECODE," & _
				" ATTENDANCE)" & _
				" GROUP BY FCBA," & _
				" TO_NUMBER (TO_CHAR (FDDATE, 'YYYY'))," & _
				" TO_NUMBER (TO_CHAR (FDDATE, 'MM'))," & _
				" EMPLOYEECODE) AG" & _
				" INNER JOIN" & _
				" EMPLOYEE EM" & _
				" ON (AG.FCBA = EM.FCBA AND AG.EMPLOYEECODE = EM.FCCODE)" & _
				" WHERE EM.PRESENTGRADE IN ('KHT', 'KBT') AND AG.FCBA = '" & Session ("FCBA") & "'"

if  LEN(TRIM(v_txtYear)) > 0 then
	strEmployee	=	strEmployee & " AND AG.NYEAR = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " AND AG.NMONTH = " & v_txtMonth
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strEmployee	=	strEmployee & " AND EM.GANGCODE = '" & v_txtGangCode & "'"
end if
if  LEN(TRIM(v_txtGrade)) > 0 then
	strEmployee	=	strEmployee & " AND EM.PRESENTGRADE = '" & v_txtGrade & "'"
end if
strEmployee	=	strEmployee & " ORDER BY EMPLOYEECODE ASC"
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Absensi KHT dan KBT</title>
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
		<h2>ABSENSI KHT dan KBT</h3>
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
		<th>GANG</th>
<%for nday = 1 to 31%>
<th><%=nday%></th>
<%next%>
	</tr>

<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic
if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td colspan="37">NO DATA AVAILABLE</td>
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
				<td align="center"><%=rsEmployee("STATUS")%></td>
				<td align="center"><%=rsEmployee("GANGCODE")%></td>
				<td align="center"><%=rsEmployee("day_1")%></td>
				<td align="center"><%=rsEmployee("day_2")%></td>
				<td align="center"><%=rsEmployee("day_3")%></td>
				<td align="center"><%=rsEmployee("day_4")%></td>
				<td align="center"><%=rsEmployee("day_5")%></td>
				<td align="center"><%=rsEmployee("day_6")%></td>
				<td align="center"><%=rsEmployee("day_7")%></td>
				<td align="center"><%=rsEmployee("day_8")%></td>
				<td align="center"><%=rsEmployee("day_9")%></td>
				<td align="center"><%=rsEmployee("day_10")%></td>
				<td align="center"><%=rsEmployee("day_11")%></td>
				<td align="center"><%=rsEmployee("day_12")%></td>
				<td align="center"><%=rsEmployee("day_13")%></td>
				<td align="center"><%=rsEmployee("day_14")%></td>
				<td align="center"><%=rsEmployee("day_15")%></td>
				<td align="center"><%=rsEmployee("day_16")%></td>
				<td align="center"><%=rsEmployee("day_17")%></td>
				<td align="center"><%=rsEmployee("day_18")%></td>
				<td align="center"><%=rsEmployee("day_19")%></td>
				<td align="center"><%=rsEmployee("day_20")%></td>
				<td align="center"><%=rsEmployee("day_21")%></td>
				<td align="center"><%=rsEmployee("day_22")%></td>
				<td align="center"><%=rsEmployee("day_23")%></td>
				<td align="center"><%=rsEmployee("day_24")%></td>
				<td align="center"><%=rsEmployee("day_25")%></td>
				<td align="center"><%=rsEmployee("day_26")%></td>
				<td align="center"><%=rsEmployee("day_27")%></td>
				<td align="center"><%=rsEmployee("day_28")%></td>
				<td align="center"><%=rsEmployee("day_29")%></td>
				<td align="center"><%=rsEmployee("day_30")%></td>
				<td align="center"><%=rsEmployee("day_31")%></td>
<%
			NO = NO + 1			
			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>
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
