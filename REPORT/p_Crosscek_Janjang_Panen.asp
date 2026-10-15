<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade
Dim ScriptName
Dim rsEmployee, strEmployee, rsAtt, strAtt, rsUNIT, strUNIT
Dim NO, cClass, nday

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtGangCode	=	Request("txtGangCode")
v_txtGrade		=	Request("txtGrade")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Date)
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Date)
end if

' strEmployee	=	 " SELECT ag.fcba fcba," & _ 
                 ' " ag.gangcode gangcode, " & _
                 ' " nyear, " & _
                 ' " nmonth, " & _
                 ' " employeecode," & _ 
                 ' " em.fcname employeename," & _ 
                 ' " em.presentgrade status, " & _
                 ' " day_1, " & _
                 ' " day_2, " & _
                 ' " day_3, " & _
                 ' " day_4, " & _
                 ' " day_5, " & _
                 ' " day_6, " & _
                 ' " day_7, " & _
                 ' " day_8, " & _
                 ' " day_9, " & _
                 ' " day_10, " & _
                 ' " day_11, " & _
                 ' " day_12, " & _
                 ' " day_13, " & _
                 ' " day_14, " & _
                 ' " day_15, " & _
                 ' " day_16, " & _
                 ' " day_17, " & _
                 ' " day_18, " & _
                 ' " day_19, " & _
                 ' " day_20, " & _
                 ' " day_21, " & _
                 ' " day_22, " & _
                 ' " day_23, " & _
                 ' " day_24, " & _
                 ' " day_25, " & _
                 ' " day_26, " & _
                 ' " day_27, " & _
                 ' " day_28, " & _
                 ' " day_29, " & _
                 ' " day_30, " & _
                 ' " day_31, " & _
                 ' " total " & _
                 ' " FROM ( SELECT fcba," & _ 
                 ' " nyear, " & _
                 ' " nmonth, " & _
                 ' " gangcode," & _
                 ' " employeecode, " & _
                 ' " day_1, " & _
                 ' " day_2, " & _
                 ' " day_3, " & _
                 ' " day_4, " & _
                 ' " day_5, " & _
                 ' " day_6, " & _
                 ' " day_7, " & _
                 ' " day_8, " & _
                 ' " day_9, " & _
                 ' " day_10, " & _
                 ' " day_11, " & _
                 ' " day_12, " & _
                 ' " day_13, " & _
                 ' " day_14, " & _
                 ' " day_15, " & _
                 ' " day_16, " & _
                 ' " day_17, " & _
                 ' " day_18, " & _
                 ' " day_19, " & _
                 ' " day_20, " & _
                 ' " day_21, " & _
                 ' " day_22, " & _
                 ' " day_23, " & _
                 ' " day_24, " & _
                 ' " day_25, " & _
                 ' " day_26, " & _
                 ' " day_27, " & _
                 ' " day_28, " & _
                 ' " day_29, " & _
                 ' " day_30, " & _
                 ' " day_31, " & _
                 ' " SUM ( " & _
                 ' " day_1 " & _
                 ' " + day_2 " & _
                 ' " + day_3 " & _
                 ' " + day_4 " & _
                 ' " + day_5 " & _
                 ' " + day_6 " & _
                 ' " + day_7 " & _
                 ' " + day_8 " & _
                 ' " + day_9 " & _
                 ' " + day_10 " & _
                 ' " + day_11 " & _
                 ' " + day_12 " & _
                 ' " + day_13 " & _
                 ' " + day_14 " & _
                 ' " + day_15 " & _
                 ' " + day_16 " & _
                 ' " + day_17 " & _
                 ' " + day_18 " & _
                 ' " + day_19 " & _
                 ' " + day_20 " & _
                 ' " + day_21 " & _
                 ' " + day_22 " & _
                 ' " + day_23 " & _
                 ' " + day_24 " & _
                 ' " + day_25 " & _
                 ' " + day_26 " & _
                 ' " + day_27 " & _
                 ' " + day_28 " & _
                 ' " + day_29 " & _
                 ' " + day_30 " & _
                 ' " + day_31) " & _
                 ' " total " & _
                 ' " FROM ( SELECT fcba," & _ 
                 ' " TO_NUMBER (TO_CHAR (fddate, 'yyyy')) nyear," & _ 
                 ' " TO_NUMBER (TO_CHAR (fddate, 'mm')) nmonth, " & _
                 ' " gangcode," & _
                 ' " employeecode, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 1, unit, " & _
                 ' " 0)) " & _
                 ' " day_1, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 2, unit, " & _
                 ' " 0)) " & _
                 ' " day_2, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 3, unit, " & _
                 ' " 0)) " & _
                 ' " day_3, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 4, unit, " & _
                 ' " 0)) " & _
                 ' " day_4, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 5, unit, " & _
                 ' " 0)) " & _
                 ' " day_5, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 6, unit, " & _
                 ' " 0)) " & _
                 ' " day_6, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 7, unit, " & _
                 ' " 0)) " & _
                 ' " day_7, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 8, unit, " & _
                 ' " 0)) " & _
                 ' " day_8, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 9, unit, " & _
                 ' " 0)) " & _
                 ' " day_9, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 10, unit, " & _
                 ' " 0)) " & _
                 ' " day_10, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 11, unit, " & _
                 ' " 0)) " & _
                 ' " day_11, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 12, unit, " & _
                 ' " 0)) " & _
                 ' " day_12, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 13, unit, " & _
                 ' " 0)) " & _
                 ' " day_13, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 14, unit, " & _
                 ' " 0)) " & _
                 ' " day_14, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 15, unit, " & _
                 ' " 0)) " & _
                 ' " day_15, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 16, unit, " & _
                 ' " 0)) " & _
                 ' " day_16, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 17, unit, " & _
                 ' " 0)) " & _
                 ' " day_17, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 18, unit, " & _
                 ' " 0)) " & _
                 ' " day_18, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 19, unit, " & _
                 ' " 0)) " & _
                 ' " day_19, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 20, unit, " & _
                 ' " 0)) " & _
                 ' " day_20, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 21, unit, " & _
                 ' " 0)) " & _
                 ' " day_21, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 22, unit, " & _
                 ' " 0)) " & _
                 ' " day_22, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 23, unit, " & _
                 ' " 0)) " & _
                 ' " day_23, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 24, unit, " & _
                 ' " 0)) " & _
                 ' " day_24, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 25, unit, " & _
                 ' " 0)) " & _
                 ' " day_25, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 26, unit, " & _
                 ' " 0)) " & _
                 ' " day_26, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 27, unit, " & _
                 ' " 0)) " & _
                 ' " day_27, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 28, unit, " & _
                 ' " 0)) " & _
                 ' " day_28, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd')), " & _
                 ' " 29, unit, " & _
                 ' " 0)) " & _
                 ' " day_29, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 30, unit, " & _
                 ' " 0)) " & _
                 ' " day_30, " & _
                 ' " SUM ( " & _
                 ' " DECODE (TO_NUMBER (TO_CHAR (fddate, 'dd'))," & _ 
                 ' " 31, unit, " & _
                 ' " 0)) " & _
                 ' " day_31 " & _
                 ' " FROM ( SELECT fcba," & _ 
                 ' " fddate, " & _
                 ' " gangcode," & _
                 ' " employeecode, " & _
                 ' " jobcode," & _
                 ' " SUM (DECODE (jobcode, '6502101', unit, 0)) " & _
                 ' " unit" & _
                 ' " FROM attendance_gad " & _
                 ' " GROUP BY fcba, " & _
                 ' " fddate, " & _
                 ' " gangcode," & _
                 ' " employeecode," & _
                 ' " jobcode " & _
				 ' " UNION ALL " & _
				 ' " SELECT b.fcba," & _ 
                 ' " b.bkmdate fddate, " & _
                 ' " b.gangcode," & _
                 ' " a.employeecode, " & _
                 ' " a.actcode jobcode," & _
                 ' " SUM (DECODE (a.actcode, '6502101', unit, 0)) unit " & _
                 ' " FROM bkmupahdtl a inner join bkmhdr b on (a.bkmno=b.bkmno and a.fcba=b.fcba) " & _
                 ' " GROUP BY b.fcba, " & _
                 ' " b.bkmdate, " & _
                 ' " b.gangcode," & _
                 ' " a.employeecode," & _
                 ' " a.actcode " & _
				 ' " ) ag " & _
                 ' " GROUP BY fcba," & _ 
                 ' " gangcode," & _
                 ' " employeecode, " & _
                 ' " TO_NUMBER (TO_CHAR (fddate, 'yyyy')), " & _
                 ' " TO_NUMBER (TO_CHAR (fddate, 'mm'))) " & _
                 ' " GROUP BY fcba, " & _
                 ' " nyear, " & _
                 ' " nmonth, " & _
                 ' " gangcode," & _
                 ' " employeecode, " & _
                 ' " day_1, " & _
                 ' " day_2, " & _
                 ' " day_3, " & _
                 ' " day_4, " & _
                 ' " day_5, " & _
                 ' " day_6, " & _
                 ' " day_7, " & _
                 ' " day_8, " & _
                 ' " day_9, " & _
                 ' " day_10, " & _
                 ' " day_11, " & _
                 ' " day_12, " & _
                 ' " day_13, " & _
                 ' " day_14, " & _
                 ' " day_15, " & _
                 ' " day_16, " & _
                 ' " day_17, " & _
                 ' " day_18, " & _
                 ' " day_19, " & _
                 ' " day_20, " & _
                 ' " day_21, " & _
                 ' " day_22, " & _
                 ' " day_23, " & _
                 ' " day_24, " & _
                 ' " day_25, " & _
                 ' " day_26, " & _
                 ' " day_27, " & _
                 ' " day_28, " & _
                 ' " day_29, " & _
                 ' " day_30, " & _
                 ' " day_31) ag " & _
                 ' " INNER JOIN " & _
                 ' " employee em " & _
                 ' " ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode)" & _ 
                 ' " WHERE AG.FCBA = '" & Session ("FCBA") & "'"

' if  LEN(TRIM(v_txtYear)) > 0 then
	' strEmployee	=	strEmployee & " AND NYEAR = " & v_txtYear
' end if
' if  LEN(TRIM(v_txtMonth)) > 0 then
	' strEmployee	=	strEmployee & " AND NMONTH = " & v_txtMonth
' end if
' if  LEN(TRIM(v_txtGangCode)) > 0 then
	' strEmployee	=	strEmployee & " AND AG.GANGCODE = '" & v_txtGangCode & "'"
' end if
' strEmployee	=	strEmployee & " ORDER BY EMPLOYEECODE ASC"

strEmployee	= " SELECT  fcba, " & _
				" gangcode, " & _
				" nyear, " & _
				" nmonth, " & _
				" employeecode, " & _
				" employeename, " & _
				" status," & _
				"   day_1, " & _
                "   day_2, " & _
                "   day_3, " & _
                "   day_4, " & _
                "   day_5, " & _
                "   day_6, " & _
                "   day_7, " & _
                "   day_8, " & _
                "   day_9, " & _
                "   day_10, " & _
                "   day_11, " & _
                "   day_12, " & _
                "   day_13, " & _
                "   day_14, " & _
                "   day_15, " & _
                "   day_16, " & _
                "   day_17, " & _
                "   day_18, " & _
                "   day_19, " & _
                "   day_20, " & _
                "   day_21, " & _
                "   day_22, " & _
                "   day_23, " & _
                "   day_24, " & _
                "   day_25, " & _
                "   day_26, " & _
                "   day_27, " & _
                "   day_28, " & _
                "   day_29, " & _
                "   day_30, " & _
                "   day_31, " & _
                "   (" & _
                "   nvl(day_1,0)+ " & _
                "   nvl(day_2,0)+ " & _
                "   nvl(day_3,0)+ " & _
                "   nvl(day_4,0)+ " & _
                "   nvl(day_5,0)+ " & _
                "   nvl(day_6,0)+ " & _
                "   nvl(day_7,0)+ " & _
                "   nvl(day_8,0)+ " & _
                "   nvl(day_9,0)+ " & _
                "   nvl(day_10,0)+ " & _
                "   nvl(day_11,0)+ " & _
                "   nvl(day_12,0)+ " & _
                "   nvl(day_13,0)+ " & _
                "   nvl(day_14,0)+ " & _
                "   nvl(day_15,0)+ " & _
                "   nvl(day_16,0)+ " & _
                "   nvl(day_17,0)+ " & _
                "   nvl(day_18,0)+ " & _
                "   nvl(day_19,0)+ " & _
                "   nvl(day_20,0)+ " & _
                "   nvl(day_21,0)+ " & _
                "   nvl(day_22,0)+ " & _
                "   nvl(day_23,0)+ " & _
                "   nvl(day_24,0)+ " & _
                "   nvl(day_25,0)+ " & _
                "   nvl(day_26,0)+ " & _
                "   nvl(day_27,0)+ " & _
                "   nvl(day_28,0)+ " & _
                "   nvl(day_29,0)+ " & _
                "   nvl(day_30,0)+ " & _
                "   nvl(day_31,0)) total" & _ 
				"   FROM " & _
				"   	(" & _
				"   	SELECT b.fcba, " & _
				"   		   b.gangcode," & _
				"   		   To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
				"   		   extract(YEAR FROM bkmdate) nyear," & _
				"   		   extract(MONTH FROM bkmdate) nmonth," & _
				"   		   a.employeecode, " & _
				"   		   c.fcname employeename," & _
				"   		   c.presentgrade status," & _
				"   		   SUM (Nvl(unit, 0)) unit " & _
				"   	FROM bkmupahdtl a " & _
				"   	INNER JOIN bkmhdr b on (a.bkmno=b.bkmno and a.fcba=b.fcba)" & _
				"   	INNER JOIN employee c ON (a.employeecode=c.fccode AND a.fcba=c.fcba) " & _
				"   	WHERE a.actcode IN ('6502101','1321410') " & _
				"   	AND attendance IN('KJ','KK','WH','WS','TG','KH')" & _
				"   	AND b.postingflg='1' AND extract(year from bkmdate) = " & v_txtYear & " AND extract(month from bkmdate) = " & v_txtMonth & " AND a.chargeunit = '" & Session ("FCBA") & "' " & _
				"       AND b.gangcode like '%" & v_txtGangCode & "%'" & _ 
				"   	GROUP BY b.fcba, b.bkmdate, b.gangcode, a.employeecode, c.fcname,c.presentgrade " & _
				"   	)bkm" & _
				"   	PIVOT ( " & _
				"   	Sum(Nvl(Unit,0)) FOR tgl IN ( 1 AS day_1, 2 AS day_2, 3 AS day_3, 4 AS day_4, 5 AS day_5, 6 AS day_6, 7 AS day_7, 8 AS day_8, 9 AS day_9, 10 AS day_10, 11 AS day_11, 12 AS day_12, 13 AS day_13, 14 AS day_14, 15 AS day_15, 16 AS day_16, 17 AS day_17, 18 AS day_18, 19 AS day_19, 20 AS day_20, 21 AS day_21, 22 AS day_22, 23 AS day_23, 24 AS day_24, 25 AS day_25, 26 AS day_26, 27 AS day_27, 28 AS day_28, 29 AS day_29, 30 AS day_30, 31 AS day_31 ) " & _
				"   		   )" & _
				"   	ORDER BY employeename,gangcode "

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - JANJANG PANEN</title>
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
		<td><font face="verdana" size="1">GANGCODE</font></td><td><select name="txtGangCode" ><option value=""></option>
		<%
		Dim rsGangCode, strGangCode
		strGangCode	=	" SELECT  DISTINCT GANGCODE FROM ATTENDANCE_GAD WHERE JOBCODE ='6502101' AND FCBA = '" & Session("FCBA") & "' ORDER BY GANGCODE ASC"
		
		Set rsGangCode = Server.CreateObject("ADODB.Recordset")
		rsGangCode.Open strGangCode, connstr, adOpenStatic 
		
		if not rsGangCode.eof then
		Do While Not rsGangCode.eof
		if rsGangCode("gangcode") = v_txtGangCode then%>
		        <option value="<%=rsGangCode("gangcode")%>" Selected><%=rsGangCode("gangcode")%></option> 
		     <%Else%>
		        <option value="<%=rsGangCode("gangcode")%>"><%=rsGangCode("gangcode")%></option>  
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
</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;"> 
		<h2>MONITORING JANJANG PANEN</h2>
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
			
			nday_1 =  nday_1 + Cdbl(CheckValue(rsEmployee("day_1")))
			nday_2 =  nday_2 + Cdbl(CheckValue(rsEmployee("day_2")))
			nday_3 =  nday_3 + Cdbl(CheckValue(rsEmployee("day_3")))
			nday_4 =  nday_4 + Cdbl(CheckValue(rsEmployee("day_4")))
			nday_5 =  nday_5 + Cdbl(CheckValue(rsEmployee("day_5")))
			nday_6 =  nday_6 + Cdbl(CheckValue(rsEmployee("day_6")))
			nday_7 =  nday_7 + Cdbl(CheckValue(rsEmployee("day_7")))
			nday_8 =  nday_8 + Cdbl(CheckValue(rsEmployee("day_8")))
			nday_9 =  nday_9 + Cdbl(CheckValue(rsEmployee("day_9")))
			nday_10 =  nday_10 + Cdbl(CheckValue(rsEmployee("day_10")))
			nday_11 =  nday_11 + Cdbl(CheckValue(rsEmployee("day_11")))
			nday_12 =  nday_12 + Cdbl(CheckValue(rsEmployee("day_12")))
			nday_13 =  nday_13 + Cdbl(CheckValue(rsEmployee("day_13")))
			nday_14 =  nday_14 + Cdbl(CheckValue(rsEmployee("day_14")))
			nday_15 =  nday_15 + Cdbl(CheckValue(rsEmployee("day_15")))
			nday_16 =  nday_16 + Cdbl(CheckValue(rsEmployee("day_16")))
			nday_17 =  nday_17 + Cdbl(CheckValue(rsEmployee("day_17")))
			nday_18 =  nday_18 + Cdbl(CheckValue(rsEmployee("day_18")))
			nday_19 =  nday_19 + Cdbl(CheckValue(rsEmployee("day_19")))
			nday_20 =  nday_20 + Cdbl(CheckValue(rsEmployee("day_20")))
			nday_21 =  nday_21 + Cdbl(CheckValue(rsEmployee("day_21")))
			nday_22 =  nday_22 + Cdbl(CheckValue(rsEmployee("day_22")))
			nday_23 =  nday_23 + Cdbl(CheckValue(rsEmployee("day_23")))
			nday_24 =  nday_24 + Cdbl(CheckValue(rsEmployee("day_24")))
			nday_25 =  nday_25 + Cdbl(CheckValue(rsEmployee("day_25")))
			nday_26 =  nday_26 + Cdbl(CheckValue(rsEmployee("day_26")))
			nday_27 =  nday_27 + Cdbl(CheckValue(rsEmployee("day_27")))
			nday_28 =  nday_28 + Cdbl(CheckValue(rsEmployee("day_28")))
			nday_29 =  nday_29 + Cdbl(CheckValue(rsEmployee("day_29")))
			nday_30 =  nday_30 + Cdbl(CheckValue(rsEmployee("day_30")))
			nday_31 =  nday_31 + Cdbl(CheckValue(rsEmployee("day_31")))
			nTOTAL =  nTOTAL + Cdbl(CheckValue(rsEmployee("TOTAL")))
			
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
    </tbody>
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
