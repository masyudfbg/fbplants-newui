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

strEmployee	=	" SELECT ag.fcba fcba," & _
				" ag.gangcode gangcode," & _
				" nyear," & _
				" nmonth," & _
				" employeecode," & _
				" em.fcname employeename," & _
				" em.presentgrade status," & _
				" day_1," & _
				" day_2," & _
				" day_3," & _
				" day_4," & _
				" day_5," & _
				" day_6," & _
				" day_7," & _
				" day_8," & _
				" day_9," & _
				" day_10," & _
				" day_11," & _
				" day_12," & _
				" day_13," & _
				" day_14," & _
				" day_15," & _
				" day_16," & _
				" day_17," & _
				" day_18," & _
				" day_19," & _
				" day_20," & _
				" day_21," & _
				" day_22," & _
				" day_23," & _
				" day_24," & _
				" day_25," & _
				" day_26," & _
				" day_27," & _
				" day_28," & _
				" day_29," & _
				" day_30," & _
				" day_31," & _
				" total" & _
				" FROM ( SELECT fcba," & _
				" nyear," & _
				" nmonth," & _
				" gangcode," & _
				" employeecode," & _
				" day_1," & _
				" day_2," & _
				" day_3," & _
				" day_4," & _
				" day_5," & _
				" day_6," & _
				" day_7," & _
				" day_8," & _
				" day_9," & _
				" day_10," & _
				" day_11," & _
				" day_12," & _
				" day_13," & _
				" day_14," & _
				" day_15," & _
				" day_16," & _
				" day_17," & _
				" day_18," & _
				" day_19," & _
				" day_20," & _
				" day_21," & _
				" day_22," & _
				" day_23," & _
				" day_24," & _
				" day_25," & _
				" day_26," & _
				" day_27," & _
				" day_28," & _
				" day_29," & _
				" day_30," & _
				" day_31," & _
				" SUM (" & _
				" day_1" & _
				" + day_2" & _
				" + day_3" & _
				" + day_4" & _
				" + day_5" & _
				" + day_6" & _
				" + day_7" & _
				" + day_8" & _
				" + day_9" & _
				" + day_10" & _
				" + day_11" & _
				" + day_12" & _
				" + day_13" & _
				" + day_14" & _
				" + day_15" & _
				" + day_16" & _
				" + day_17" & _
				" + day_18" & _
				" + day_19" & _
				" + day_20" & _
				" + day_21" & _
				" + day_22" & _
				" + day_23" & _
				" + day_24" & _
				" + day_25" & _
				" + day_26" & _
				" + day_27" & _
				" + day_28" & _
				" + day_29" & _
				" + day_30" & _
				" + day_31)" & _
				" total" & _
				" FROM ( SELECT fcba," & _
				" TO_NUMBER (TO_CHAR (fddate, 'yyyy')) nyear," & _
				" TO_NUMBER (TO_CHAR (fddate, 'mm')) nmonth," & _
				" gangcode," & _
				" employeecode," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 1, mandays," & _
				" 0))" & _
				" day_1," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 2, mandays," & _
				" 0))" & _
				" day_2," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 3, mandays," & _
				" 0))" & _
				" day_3," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 4, mandays," & _
				" 0))" & _
				" day_4," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 5, mandays," & _
				" 0))" & _
				" day_5," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 6, mandays," & _
				" 0))" & _
				" day_6," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 7, mandays," & _
				" 0))" & _
				" day_7," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 8, mandays," & _
				" 0))" & _
				" day_8," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 9, mandays," & _
				" 0))" & _
				" day_9," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 10, mandays," & _
				" 0))" & _
				" day_10," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 11, mandays," & _
				" 0))" & _
				" day_11," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 12, mandays," & _
				" 0))" & _
				" day_12," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 13, mandays," & _
				" 0))" & _
				" day_13," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 14, mandays," & _
				" 0))" & _
				" day_14," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 15, mandays," & _
				" 0))" & _
				" day_15," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 16, mandays," & _
				" 0))" & _
				" day_16," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 17, mandays," & _
				" 0))" & _
				" day_17," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 18, mandays," & _
				" 0))" & _
				" day_18," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 19, mandays," & _
				" 0))" & _
				" day_19," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 20, mandays," & _
				" 0))" & _
				" day_20," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 21, mandays," & _
				" 0))" & _
				" day_21," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 22, mandays," & _
				" 0))" & _
				" day_22," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 23, mandays," & _
				" 0))" & _
				" day_23," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 24, mandays," & _
				" 0))" & _
				" day_24," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 25, mandays," & _
				" 0))" & _
				" day_25," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 26, mandays," & _
				" 0))" & _
				" day_26," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 27, mandays," & _
				" 0))" & _
				" day_27," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 28, mandays," & _
				" 0))" & _
				" day_28," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 29, mandays," & _
				" 0))" & _
				" day_29," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 30, mandays," & _
				" 0))" & _
				" day_30," & _
				" SUM (" & _
				" DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _
				" 31, mandays," & _
				" 0))" & _
				" day_31" & _
				" FROM ( SELECT fcba," & _
				" fddate," & _
				" gangcode," & _
				" employeecode," & _
				" SUM (DECODE (attendance, 'KJ', mandays, 0))" & _
				" mandays" & _
				" FROM attendance_gad" & _
				" GROUP BY fcba," & _
				" fddate," & _
				" gangcode," & _
				" employeecode)" & _
				" GROUP BY fcba," & _
				" gangcode," & _
				" employeecode," & _
				" TO_NUMBER (TO_CHAR (fddate, 'yyyy'))," & _
				" TO_NUMBER (TO_CHAR (fddate, 'mm')))" & _
				" GROUP BY fcba," & _
				" nyear," & _
				" nmonth," & _
				" gangcode," & _
				" employeecode," & _
				" day_1," & _
				" day_2," & _
				" day_3," & _
				" day_4," & _
				" day_5," & _
				" day_6," & _
				" day_7," & _
				" day_8," & _
				" day_9," & _
				" day_10," & _
				" day_11," & _
				" day_12," & _
				" day_13," & _
				" day_14," & _
				" day_15," & _
				" day_16," & _
				" day_17," & _
				" day_18," & _
				" day_19," & _
				" day_20," & _
				" day_21," & _
				" day_22," & _
				" day_23," & _
				" day_24," & _
				" day_25," & _
				" day_26," & _
				" day_27," & _
				" day_28," & _
				" day_29," & _
				" day_30," & _
				" day_31) ag" & _
				" INNER JOIN" & _
				" employee em" & _
				" ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode)" & _
				" WHERE AG.FCBA = '" & Session ("FCBA") & "' AND EM.PRESENTGRADE = 'KHL'"

if  LEN(TRIM(v_txtYear)) > 0 then
	strEmployee	=	strEmployee & " AND NYEAR = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " AND NMONTH = " & v_txtMonth
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strEmployee	=	strEmployee & " AND AG.GANGCODE = '" & v_txtGangCode & "'"
end if
strEmployee	=	strEmployee & " ORDER BY EMPLOYEECODE ASC"
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - HK Karyawan Harian Lepas</title>
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
		<th>TOTAL</th>
	</tr>

<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic

Dim nday_1
Dim nday_2
Dim nday_3
Dim nday_4
Dim nday_5
Dim nday_6
Dim nday_7
Dim nday_8
Dim nday_9
Dim nday_10
Dim nday_11
Dim nday_12
Dim nday_13
Dim nday_14
Dim nday_15
Dim nday_16
Dim nday_17
Dim nday_18
Dim nday_19
Dim nday_20
Dim nday_21
Dim nday_22
Dim nday_23
Dim nday_24
Dim nday_25
Dim nday_26
Dim nday_27
Dim nday_28
Dim nday_29
Dim nday_30
Dim nday_31
Dim nTOTAL

if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td colspan="37">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		
		nday_1 = 0
		nday_2 = 0
		nday_3 = 0
		nday_4 = 0
		nday_5 = 0
		nday_6 = 0
		nday_7 = 0
		nday_8 = 0
		nday_9 = 0
		nday_10 = 0
		nday_11 = 0
		nday_12 = 0
		nday_13 = 0
		nday_14 = 0
		nday_15 = 0
		nday_16 = 0
		nday_17 = 0
		nday_18 = 0
		nday_19 = 0
		nday_20 = 0
		nday_21 = 0
		nday_22 = 0
		nday_23 = 0
		nday_24 = 0
		nday_25 = 0
		nday_26 = 0
		nday_27 = 0
		nday_28 = 0
		nday_29 = 0
		nday_30 = 0
		nday_31 = 0
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
				<td align="center"><%=rsEmployee("STATUS")%></td>
				<td align="center"><%=rsEmployee("GANGCODE")%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_1"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_2"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_3"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_4"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_5"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_6"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_7"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_8"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_9"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_10"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_11"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_12"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_13"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_14"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_15"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_16"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_17"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_18"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_19"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_20"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_21"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_22"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_23"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_24"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_25"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_26"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_27"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_28"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_29"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_30"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("day_31"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("TOTAL"))%></td>
<%
			NO = NO + 1
			
			nday_1 =  nday_1 + Cdbl(rsEmployee("day_1"))
			nday_2 =  nday_2 + Cdbl(rsEmployee("day_2"))
			nday_3 =  nday_3 + Cdbl(rsEmployee("day_3"))
			nday_4 =  nday_4 + Cdbl(rsEmployee("day_4"))
			nday_5 =  nday_5 + Cdbl(rsEmployee("day_5"))
			nday_6 =  nday_6 + Cdbl(rsEmployee("day_6"))
			nday_7 =  nday_7 + Cdbl(rsEmployee("day_7"))
			nday_8 =  nday_8 + Cdbl(rsEmployee("day_8"))
			nday_9 =  nday_9 + Cdbl(rsEmployee("day_9"))
			nday_10 =  nday_10 + Cdbl(rsEmployee("day_10"))
			nday_11 =  nday_11 + Cdbl(rsEmployee("day_11"))
			nday_12 =  nday_12 + Cdbl(rsEmployee("day_12"))
			nday_13 =  nday_13 + Cdbl(rsEmployee("day_13"))
			nday_14 =  nday_14 + Cdbl(rsEmployee("day_14"))
			nday_15 =  nday_15 + Cdbl(rsEmployee("day_15"))
			nday_16 =  nday_16 + Cdbl(rsEmployee("day_16"))
			nday_17 =  nday_17 + Cdbl(rsEmployee("day_17"))
			nday_18 =  nday_18 + Cdbl(rsEmployee("day_18"))
			nday_19 =  nday_19 + Cdbl(rsEmployee("day_19"))
			nday_20 =  nday_20 + Cdbl(rsEmployee("day_20"))
			nday_21 =  nday_21 + Cdbl(rsEmployee("day_21"))
			nday_22 =  nday_22 + Cdbl(rsEmployee("day_22"))
			nday_23 =  nday_23 + Cdbl(rsEmployee("day_23"))
			nday_24 =  nday_24 + Cdbl(rsEmployee("day_24"))
			nday_25 =  nday_25 + Cdbl(rsEmployee("day_25"))
			nday_26 =  nday_26 + Cdbl(rsEmployee("day_26"))
			nday_27 =  nday_27 + Cdbl(rsEmployee("day_27"))
			nday_28 =  nday_28 + Cdbl(rsEmployee("day_28"))
			nday_29 =  nday_29 + Cdbl(rsEmployee("day_29"))
			nday_30 =  nday_30 + Cdbl(rsEmployee("day_30"))
			nday_31 =  nday_31 + Cdbl(rsEmployee("day_31"))
			nTOTAL =  nTOTAL + Cdbl(rsEmployee("TOTAL"))
			
			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>

	<tr>
		<td colspan="5" align="center"><strong>TOTAL</strong></td>
		<td align="right"><%=FbpNum2(nday_1)%></td>
		<td align="right"><%=FbpNum2(nday_2)%></td>
		<td align="right"><%=FbpNum2(nday_3)%></td>
		<td align="right"><%=FbpNum2(nday_4)%></td>
		<td align="right"><%=FbpNum2(nday_5)%></td>
		<td align="right"><%=FbpNum2(nday_6)%></td>
		<td align="right"><%=FbpNum2(nday_7)%></td>
		<td align="right"><%=FbpNum2(nday_8)%></td>
		<td align="right"><%=FbpNum2(nday_9)%></td>
		<td align="right"><%=FbpNum2(nday_10)%></td>
		<td align="right"><%=FbpNum2(nday_11)%></td>
		<td align="right"><%=FbpNum2(nday_12)%></td>
		<td align="right"><%=FbpNum2(nday_13)%></td>
		<td align="right"><%=FbpNum2(nday_14)%></td>
		<td align="right"><%=FbpNum2(nday_15)%></td>
		<td align="right"><%=FbpNum2(nday_16)%></td>
		<td align="right"><%=FbpNum2(nday_17)%></td>
		<td align="right"><%=FbpNum2(nday_18)%></td>
		<td align="right"><%=FbpNum2(nday_19)%></td>
		<td align="right"><%=FbpNum2(nday_20)%></td>
		<td align="right"><%=FbpNum2(nday_21)%></td>
		<td align="right"><%=FbpNum2(nday_22)%></td>
		<td align="right"><%=FbpNum2(nday_23)%></td>
		<td align="right"><%=FbpNum2(nday_24)%></td>
		<td align="right"><%=FbpNum2(nday_25)%></td>
		<td align="right"><%=FbpNum2(nday_26)%></td>
		<td align="right"><%=FbpNum2(nday_27)%></td>
		<td align="right"><%=FbpNum2(nday_28)%></td>
		<td align="right"><%=FbpNum2(nday_29)%></td>
		<td align="right"><%=FbpNum2(nday_30)%></td>
		<td align="right"><%=FbpNum2(nday_31)%></td>
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
