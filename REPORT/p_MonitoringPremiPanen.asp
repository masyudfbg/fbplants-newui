<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
 <%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade
Dim ScriptName
Dim rsEmployee, strEmployee, rsMandor, strMandor, NOMOR
Dim NO, cClass, nday
v_txtYear      =    Request("txtYear")
v_txtMonth     =    Request("txtMonth")
v_txtGangCode  =    Request("txtGangCode")
v_txtGrade     =    Request("txtGrade")
ScriptName     =    Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
	v_txtYear      =     Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
	v_txtMonth     =     Month(Session("v_Employee"))
end if

strMandor       = 	" SELECT EP.FCBA," & _
						" TO_CHAR (EP.ADHOCDATE, 'YYYY') AS YEAR," & _
						" TO_CHAR (EP.ADHOCDATE, 'MM') AS MONTH," & _
						" EM.SECTIONNAME," & _
						" EM.GANGCODE," & _
						" EP.EMPCODE," & _
						" EM.FCNAME," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 1, AMOUNT, 0))" & _
						" AS DAY_1," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 2, AMOUNT, 0))" & _
						" AS DAY_2," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 3, AMOUNT, 0))" & _
						" AS DAY_3," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 4, AMOUNT, 0))" & _
						" AS DAY_4," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 5, AMOUNT, 0))" & _
						" AS DAY_5," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 6, AMOUNT, 0))" & _
						" AS DAY_6," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 7, AMOUNT, 0))" & _
						" AS DAY_7," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 8, AMOUNT, 0))" & _
						" AS DAY_8," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 9, AMOUNT, 0))" & _
						" AS DAY_9," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 10, AMOUNT, 0))" & _
						" AS DAY_10," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 11, AMOUNT, 0))" & _
						" AS DAY_11," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 12, AMOUNT, 0))" & _
						" AS DAY_12," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 13, AMOUNT, 0))" & _
						" AS DAY_13," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 14, AMOUNT, 0))" & _
						" AS DAY_14," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 15, AMOUNT, 0))" & _
						" AS DAY_15," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 16, AMOUNT, 0))" & _
						" AS DAY_16," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 17, AMOUNT, 0))" & _
						" AS DAY_17," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 18, AMOUNT, 0))" & _
						" AS DAY_18," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 19, AMOUNT, 0))" & _
						" AS DAY_19," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 20, AMOUNT, 0))" & _
						" AS DAY_20," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 21, AMOUNT, 0))" & _
						" AS DAY_21," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 22, AMOUNT, 0))" & _
						" AS DAY_22," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 23, AMOUNT, 0))" & _
						" AS DAY_23," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 24, AMOUNT, 0))" & _
						" AS DAY_24," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 25, AMOUNT, 0))" & _
						" AS DAY_25," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 26, AMOUNT, 0))" & _
						" AS DAY_26," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 27, AMOUNT, 0))" & _
						" AS DAY_27," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 28, AMOUNT, 0))" & _
						" AS DAY_28," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 29, AMOUNT, 0))" & _
						" AS DAY_29," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 30, AMOUNT, 0))" & _
						" AS DAY_30," & _
						" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 31, AMOUNT, 0))" & _
						" AS DAY_31," & _
						" SUM (EP.AMOUNT) AS TOTAL" & _
						" FROM EMPALLDED EP" & _
						" INNER JOIN ALLOWDED AL" & _
						" ON (EP.FCBA = AL.FCBA AND EP.ALLOWDEDCODE = AL.FCCODE)" & _
						" INNER JOIN EMPLOYEE EM" & _
						" ON (EP.FCBA = EM.FCBA AND EP.EMPCODE = EM.FCCODE)" & _
						" WHERE EP.ALLOWDEDCODE IN ('A24', 'A25', 'A27', 'A28','A37', 'A38', 'A39', 'A40')" & _
						" AND EM.GANGCODE LIKE 'MD%'" & _
						" AND EP.FCBA = '" & Session ("FCBA") & "'"

if  LEN(TRIM(v_txtYear)) > 0 then
    strMandor	=	strMandor & " AND TO_CHAR (EP.ADHOCDATE, 'YYYY') = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strMandor	=	strMandor & " AND TO_CHAR (EP.ADHOCDATE, 'MM') = " & v_txtMonth
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
    strMandor	=	strMandor & " AND EM.SECTIONNAME = '" & v_txtGangCode & "'"
end if
	strMandor	=	strMandor &   " GROUP BY EP.FCBA," & _
									" TO_CHAR (EP.ADHOCDATE, 'YYYY')," & _
									" TO_CHAR (EP.ADHOCDATE, 'MM')," & _
									" EM.SECTIONNAME," & _
									" EM.GANGCODE," & _
									" EP.EMPCODE," & _
									" EM.FCNAME"
	strMandor	=	strMandor & " ORDER BY EP.EMPCODE ASC"
	
strEmployee		= 	" SELECT EP.FCBA," & _
					" TO_CHAR (EP.ADHOCDATE, 'YYYY') AS YEAR," & _
					" TO_CHAR (EP.ADHOCDATE, 'MM') AS MONTH," & _
					" EM.SECTIONNAME," & _
					" EM.GANGCODE," & _
					" EP.EMPCODE," & _
					" EM.FCNAME," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 1, AMOUNT, 0))" & _
					" AS DAY_1," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 2, AMOUNT, 0))" & _
					" AS DAY_2," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 3, AMOUNT, 0))" & _
					" AS DAY_3," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 4, AMOUNT, 0))" & _
					" AS DAY_4," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 5, AMOUNT, 0))" & _
					" AS DAY_5," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 6, AMOUNT, 0))" & _
					" AS DAY_6," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 7, AMOUNT, 0))" & _
					" AS DAY_7," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 8, AMOUNT, 0))" & _
					" AS DAY_8," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 9, AMOUNT, 0))" & _
					" AS DAY_9," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 10, AMOUNT, 0))" & _
					" AS DAY_10," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 11, AMOUNT, 0))" & _
					" AS DAY_11," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 12, AMOUNT, 0))" & _
					" AS DAY_12," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 13, AMOUNT, 0))" & _
					" AS DAY_13," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 14, AMOUNT, 0))" & _
					" AS DAY_14," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 15, AMOUNT, 0))" & _
					" AS DAY_15," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 16, AMOUNT, 0))" & _
					" AS DAY_16," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 17, AMOUNT, 0))" & _
					" AS DAY_17," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 18, AMOUNT, 0))" & _
					" AS DAY_18," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 19, AMOUNT, 0))" & _
					" AS DAY_19," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 20, AMOUNT, 0))" & _
					" AS DAY_20," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 21, AMOUNT, 0))" & _
					" AS DAY_21," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 22, AMOUNT, 0))" & _
					" AS DAY_22," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 23, AMOUNT, 0))" & _
					" AS DAY_23," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 24, AMOUNT, 0))" & _
					" AS DAY_24," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 25, AMOUNT, 0))" & _
					" AS DAY_25," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 26, AMOUNT, 0))" & _
					" AS DAY_26," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 27, AMOUNT, 0))" & _
					" AS DAY_27," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 28, AMOUNT, 0))" & _
					" AS DAY_28," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 29, AMOUNT, 0))" & _
					" AS DAY_29," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 30, AMOUNT, 0))" & _
					" AS DAY_30," & _
					" SUM (DECODE (TO_NUMBER (TO_CHAR (EP.ADHOCDATE, 'DD')), 31, AMOUNT, 0))" & _
					" AS DAY_31," & _
					" SUM (EP.AMOUNT) AS TOTAL" & _
					" FROM EMPALLDED EP" & _
					" INNER JOIN ALLOWDED AL" & _
					" ON (EP.FCBA = AL.FCBA AND EP.ALLOWDEDCODE = AL.FCCODE)" & _
					" INNER JOIN EMPLOYEE EM" & _
					" ON (EP.FCBA = EM.FCBA AND EP.EMPCODE = EM.FCCODE)" & _
					" WHERE EP.ALLOWDEDCODE IN ('A24', 'A25', 'A27', 'A28','A37', 'A38', 'A39', 'A40')" & _
					" AND EM.GANGCODE LIKE 'PN%'" & _
					" AND EP.FCBA = '" & Session ("FCBA") & "'"

if  LEN(TRIM(v_txtYear)) > 0 then
    strEmployee	=	strEmployee & " AND TO_CHAR (EP.ADHOCDATE, 'YYYY') = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " AND TO_CHAR (EP.ADHOCDATE, 'MM') = " & v_txtMonth
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
    strEmployee	=	strEmployee & " AND EM.SECTIONNAME = '" & v_txtGangCode & "'"
end if
	strEmployee	=	strEmployee &   " GROUP BY EP.FCBA," & _
									" TO_CHAR (EP.ADHOCDATE, 'YYYY')," & _
									" TO_CHAR (EP.ADHOCDATE, 'MM')," & _
									" EM.SECTIONNAME," & _
									" EM.GANGCODE," & _
									" EP.EMPCODE," & _
									" EM.FCNAME"
	strEmployee	=	strEmployee & " ORDER BY EP.EMPCODE ASC"
%>
 <html>
  
 <head>
 <link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
 <title>ePlaS Reporting - Premi Panen </title>
 <script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
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
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<!--#include file="menutop.asp"-->
            <table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
		<tr>
			<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
			<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
			<td><font face="verdana" size="1">DIVISI</font></td><td><select name="txtGangCode" >
				<option value=""></option>
					<%
					Dim rsGangCode, strGangCode
					strGangCode = " SELECT * FROM PARAMETER WHERE PARCODE = 'E0002' AND FCCODE LIKE 'AFD%' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"

					Set rsGangCode = Server.CreateObject("ADODB.Recordset")
					rsGangCode.Open strGangCode, connstr, adOpenStatic 

					if not rsGangCode.eof then
					Do While Not rsGangCode.eof
					if rsGangCode("FCCODE") = v_txtGangCode then%>
					<option value="<%=rsGangCode("FCCODE")%>" Selected><%=rsGangCode("FCCODE")%></option>
					<%Else%>
					<option value="<%=rsGangCode("FCCODE")%>"><%=rsGangCode("FCCODE")%></option>  
					<%End if
					rsGangCode.movenext
					loop
					end if
					rsGangCode.Close
					set rsGangCode = Nothing
					%>
					</select>
			</td>
			<td><input type="submit" value="View" /></td>
        </tr>
	</form>
 </table>
 </div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
			<div id="subheader" class="centered" style="width:50%; margin:auto;">
				<h2>MONITORING PREMI PANEN</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
		</div>
     </th>
  </tr>
  
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
    </thead>
    <tbody>

<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic

Set rsMandor = Server.CreateObject("ADODB.Recordset")
rsMandor.Open strMandor, connstr, adOpenStatic

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

if rsMandor.eof and rsMandor.eof then
%>
	<tr bgcolor="#F4F4A0"><td colspan="37"><strong>MANDOR</strong></td></tr>
	<tr><td colspan="37">NO DATA AVAILABLE</td></tr>
<%
else
NOMOR = 1
%>
<tr bgcolor="#F4F4A0"><td colspan="37"><strong>MANDOR</strong></td></tr>
<%
while not rsMandor.eof
	if no  mod 2 = 1 then
		cClass = "odd"
	else
		cClass ="default"
	end if
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=NOMOR%></td>
		<td align="center"><%=rsMandor("EMPCODE")%></td>
		<td align="left"><%=rsMandor("FCNAME")%></td>
		<td align="center"><%=rsMandor("sectionname")%></td>
		<td align="center"><%=rsMandor("GANGCODE")%></td>
		<td align="right"><%=rsMandor("day_1")%></td>
		<td align="right"><%=rsMandor("day_2")%></td>
		<td align="right"><%=rsMandor("day_3")%></td>
		<td align="right"><%=rsMandor("day_4")%></td>
		<td align="right"><%=rsMandor("day_5")%></td>
		<td align="right"><%=rsMandor("day_6")%></td>
		<td align="right"><%=rsMandor("day_7")%></td>
		<td align="right"><%=rsMandor("day_8")%></td>
		<td align="right"><%=rsMandor("day_9")%></td>
		<td align="right"><%=rsMandor("day_10")%></td>
		<td align="right"><%=rsMandor("day_11")%></td>
		<td align="right"><%=rsMandor("day_12")%></td>
		<td align="right"><%=rsMandor("day_13")%></td>
		<td align="right"><%=rsMandor("day_14")%></td>
		<td align="right"><%=rsMandor("day_15")%></td>
		<td align="right"><%=rsMandor("day_16")%></td>
		<td align="right"><%=rsMandor("day_17")%></td>
		<td align="right"><%=rsMandor("day_18")%></td>
		<td align="right"><%=rsMandor("day_19")%></td>
		<td align="right"><%=rsMandor("day_20")%></td>
		<td align="right"><%=rsMandor("day_21")%></td>
		<td align="right"><%=rsMandor("day_22")%></td>
		<td align="right"><%=rsMandor("day_23")%></td>
		<td align="right"><%=rsMandor("day_24")%></td>
		<td align="right"><%=rsMandor("day_25")%></td>
		<td align="right"><%=rsMandor("day_26")%></td>
		<td align="right"><%=rsMandor("day_27")%></td>
		<td align="right"><%=rsMandor("day_28")%></td>
		<td align="right"><%=rsMandor("day_29")%></td>
		<td align="right"><%=rsMandor("day_30")%></td>
		<td align="right"><%=rsMandor("day_31")%></td>
		<td align="right"><%=rsMandor("TOTAL")%></td>
<%
	NOMOR	=	NOMOR	+	1
	rsMandor.MoveNext
Wend
end if
rsMandor.Close
set rsMandor = Nothing

if rsEmployee.eof and rsEmployee.eof then
%>
	<tr bgcolor="#F4F4A0"><td colspan="37"><strong>ANGGOTA</strong></td></tr>
	<tr><td colspan="37">NO DATA AVAILABLE</td></tr>
<%
else
%>
<tr bgcolor="#F4F4A0"><td colspan="37"><strong>ANGGOTA</strong></td></tr>
<%
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
		cClass ="default"
	end if
%>
	<tr class="<%=cClass%>">
		<td align="center"><%=NO%></td>
		<td align="center"><%=rsEmployee("EMPCODE")%></td>
		<td align="left"><%=rsEmployee("FCNAME")%></td>
		<td align="center"><%=rsEmployee("sectionname")%></td>
		<td align="center"><%=rsEmployee("GANGCODE")%></td>
		<td align="right"><%=rsEmployee("day_1")%></td>
		<td align="right"><%=rsEmployee("day_2")%></td>
		<td align="right"><%=rsEmployee("day_3")%></td>
		<td align="right"><%=rsEmployee("day_4")%></td>
		<td align="right"><%=rsEmployee("day_5")%></td>
		<td align="right"><%=rsEmployee("day_6")%></td>
		<td align="right"><%=rsEmployee("day_7")%></td>
		<td align="right"><%=rsEmployee("day_8")%></td>
		<td align="right"><%=rsEmployee("day_9")%></td>
		<td align="right"><%=rsEmployee("day_10")%></td>
		<td align="right"><%=rsEmployee("day_11")%></td>
		<td align="right"><%=rsEmployee("day_12")%></td>
		<td align="right"><%=rsEmployee("day_13")%></td>
		<td align="right"><%=rsEmployee("day_14")%></td>
		<td align="right"><%=rsEmployee("day_15")%></td>
		<td align="right"><%=rsEmployee("day_16")%></td>
		<td align="right"><%=rsEmployee("day_17")%></td>
		<td align="right"><%=rsEmployee("day_18")%></td>
		<td align="right"><%=rsEmployee("day_19")%></td>
		<td align="right"><%=rsEmployee("day_20")%></td>
		<td align="right"><%=rsEmployee("day_21")%></td>
		<td align="right"><%=rsEmployee("day_22")%></td>
		<td align="right"><%=rsEmployee("day_23")%></td>
		<td align="right"><%=rsEmployee("day_24")%></td>
		<td align="right"><%=rsEmployee("day_25")%></td>
		<td align="right"><%=rsEmployee("day_26")%></td>
		<td align="right"><%=rsEmployee("day_27")%></td>
		<td align="right"><%=rsEmployee("day_28")%></td>
		<td align="right"><%=rsEmployee("day_29")%></td>
		<td align="right"><%=rsEmployee("day_30")%></td>
		<td align="right"><%=rsEmployee("day_31")%></td>
		<td align="right"><%=rsEmployee("TOTAL")%></td>
<%
	NO = NO + 1

	nday_1 =  nday_1+Cdbl(rsEmployee("day_1"))
	nday_2 =  nday_2+Cdbl(rsEmployee("day_2"))
	nday_3 =  nday_3+Cdbl(rsEmployee("day_3"))
	nday_4 =  nday_4+Cdbl(rsEmployee("day_4"))
	nday_5 =  nday_5+Cdbl(rsEmployee("day_5"))
	nday_6 =  nday_6+Cdbl(rsEmployee("day_6"))
	nday_7 =  nday_7+Cdbl(rsEmployee("day_7"))
	nday_8 =  nday_8+Cdbl(rsEmployee("day_8"))
	nday_9 =  nday_9+Cdbl(rsEmployee("day_9"))
	nday_10 =  nday_10+Cdbl(rsEmployee("day_10"))
	nday_11 =  nday_11+Cdbl(rsEmployee("day_11"))
	nday_12 =  nday_12+Cdbl(rsEmployee("day_12"))
	nday_13 =  nday_13+Cdbl(rsEmployee("day_13"))
	nday_14 =  nday_14+Cdbl(rsEmployee("day_14"))
	nday_15 =  nday_15+Cdbl(rsEmployee("day_15"))
	nday_16 =  nday_16+Cdbl(rsEmployee("day_16"))
	nday_17 =  nday_17+Cdbl(rsEmployee("day_17"))
	nday_18 =  nday_18+Cdbl(rsEmployee("day_18"))
	nday_19 =  nday_19+Cdbl(rsEmployee("day_19"))
	nday_20 =  nday_20+Cdbl(rsEmployee("day_20"))
	nday_21 =  nday_21+Cdbl(rsEmployee("day_21"))
	nday_22 =  nday_22+Cdbl(rsEmployee("day_22"))
	nday_23 =  nday_23+Cdbl(rsEmployee("day_23"))
	nday_24 =  nday_24+Cdbl(rsEmployee("day_24"))
	nday_25 =  nday_25+Cdbl(rsEmployee("day_25"))
	nday_26 =  nday_26+Cdbl(rsEmployee("day_26"))
	nday_27 =  nday_27+Cdbl(rsEmployee("day_27"))
	nday_28 =  nday_28+Cdbl(rsEmployee("day_28"))
	nday_29 =  nday_29+Cdbl(rsEmployee("day_29"))
	nday_30 =  nday_30+Cdbl(rsEmployee("day_30"))
	nday_31 =  nday_31+Cdbl(rsEmployee("day_31"))
	nTOTAL =  nTOTAL+Cdbl(rsEmployee("TOTAL"))

	rsEmployee.MoveNext
Wend
end if
rsEmployee.Close
set rsEmployee = Nothing
%>

	<tr>
		<td colspan="5" align="center"><strong>TOTAL</strong></td>
		<td align="right"><%=nday_1%></td>
		<td align="right"><%=nday_2%></td>
		<td align="right"><%=nday_3%></td>
		<td align="right"><%=nday_4%></td>
		<td align="right"><%=nday_5%></td>
		<td align="right"><%=nday_6%></td>
		<td align="right"><%=nday_7%></td>
		<td align="right"><%=nday_8%></td>
		<td align="right"><%=nday_9%></td>
		<td align="right"><%=nday_10%></td>
		<td align="right"><%=nday_11%></td>
		<td align="right"><%=nday_12%></td>
		<td align="right"><%=nday_13%></td>
		<td align="right"><%=nday_14%></td>
		<td align="right"><%=nday_15%></td>
		<td align="right"><%=nday_16%></td>
		<td align="right"><%=nday_17%></td>
		<td align="right"><%=nday_18%></td>
		<td align="right"><%=nday_19%></td>
		<td align="right"><%=nday_20%></td>
		<td align="right"><%=nday_21%></td>
		<td align="right"><%=nday_22%></td>
		<td align="right"><%=nday_23%></td>
		<td align="right"><%=nday_24%></td>
		<td align="right"><%=nday_25%></td>
		<td align="right"><%=nday_26%></td>
		<td align="right"><%=nday_27%></td>
		<td align="right"><%=nday_28%></td>
		<td align="right"><%=nday_29%></td>
		<td align="right"><%=nday_30%></td>
		<td align="right"><%=nday_31%></td>
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
    </tbody>
</table>
    </body>
    </html>
