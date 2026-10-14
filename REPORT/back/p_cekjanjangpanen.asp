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
   v_txtYear	=	Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Employee"))
end if

strEmployee	=	 " Select ag.fcba fcba,  " &_
                   "   to_number(to_char(AG.fddate,'yyyy')) nYear,  " &_
                   "   to_number(to_char(AG.fddate,'mm')) nMonth,  " &_
                   "  ag.GANGCODE gangcode, " &_
                   "  ag.EMPLOYEECODE, " &_
                    " em.FCNAME employeename, " &_
                    " em.PRESENTGRADE status, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),1, unit,0)) day_1, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),2, unit,0)) day_2, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),3, unit,0))  day_3, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),4, unit,0))  day_4, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),5, unit,0))  day_5, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),6, unit,0))  day_6, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),7, unit,0))  day_7, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),8, unit,0))  day_8, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),9, unit,0))  day_9, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),10, unit,0))  day_10, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),11, unit,0))  day_11, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),12, unit,0))  day_12, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),13, unit,0))  day_13, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),14, unit,0))  day_14, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),15, unit,0))  day_15, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),16, unit,0))  day_16, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),17, unit,0))  day_17, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),18, unit,0))  day_18, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),19, unit,0))  day_19, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),20, unit,0))  day_20, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),21, unit,0))  day_21, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),22, unit,0))  day_22, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),23, unit,0))  day_23, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),24, unit,0))  day_24, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),25, unit,0))  day_25, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),26, unit,0))  day_26, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),27, unit,0))  day_27, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),28, unit,0))  day_28, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),29, unit,0))  day_29, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),30, unit,0))  day_30, " &_
                    " Sum ( decode (to_number (to_char (fddate, 'dd')),31, unit,0))  day_31, " &_
                    " Sum ( ag.unit) total  " &_
                    " From ATTENDANCE_GAD ag , EMPLOYEE Em " &_
                    " Where ag.FCBA = em.FCBA and ag.EMPLOYEECODE = em.FCCODE " &_
                    " and AG.fcba = '" & Session ("FCBA") & "'"
				  

				  
if  LEN(TRIM(v_txtYear)) > 0 then
	strEmployee	=	strEmployee & " AND to_number(to_char(AG.fddate,'yyyy')) = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " AND to_number(to_char(AG.fddate,'mm')) = " & v_txtMonth
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strEmployee	=	strEmployee & " AND AG.GANGCODE = '" & v_txtGangCode & "'"
end if
strEmployee	=	strEmployee &	" Group by AG.FCBA," &_
								"   to_number(to_char(AG.fddate,'yyyy')),  " &_
								"   to_number(to_char(AG.fddate,'mm')),  " &_
								" ag.GANGCODE, " &_
								" ag.EMPLOYEECODE, " &_
								" em.FCNAME, " &_
								" em.PRESENTGRADE " 
strEmployee	=	strEmployee & " ORDER BY AG.EMPLOYEECODE ASC"

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - HK Karyawan Harian Lepas</title>
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
		<h2>LAPORAN JANJANG PANEN</h3>
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
response.write strEmployee
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
   </table>
   </body>
   </html>
