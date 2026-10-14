<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtLocationCode
Dim ScriptName
Dim rsEmployee, strEmployee, rsAtt, strAtt, rsKJ, strKJ
Dim NO, cClass, nday

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtGangCode	=	Request("txtGangCode")
v_txtLocationCode		=	Request("txtLocationCode")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Employee"))
end if

strEmployee	=	" Select * " & _
				" from (" & _
				" Select h.FCBA, to_char(h.receptiondate,'dd-mm-yyyy') as Tanggal, h.FIELDCODE as BLOK,'2. KIRIM' as Janjang," & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),1, quantity,NULL))as day_1, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),2, quantity,NULL))as day_2, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),3, quantity,NULL))as day_3, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),4, quantity,NULL))as day_4, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),5, quantity,NULL))as day_5, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),6, quantity,NULL))as day_6, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),7, quantity,NULL))as day_7, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),8, quantity,NULL))as day_8, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),9, quantity,NULL))as day_9, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),10,quantity,NULL))as day_10, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),11,quantity,NULL))as day_11, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),12,quantity,NULL))as day_12, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),13,quantity,NULL))as day_13, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),14,quantity,NULL))as day_14, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),15,quantity,NULL))as day_15, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),16,quantity,NULL))as day_16, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),17,quantity,NULL))as day_17, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),18,quantity,NULL))as day_18, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),19,quantity,NULL))as day_19, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),20,quantity,NULL))as day_20, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),21,quantity,NULL))as day_21, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),22,quantity,NULL))as day_22, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),23,quantity,NULL))as day_23, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),24,quantity,NULL))as day_24, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),25,quantity,NULL))as day_25, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),26,quantity,NULL))as day_26, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),27,quantity,NULL))as day_27, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),28,quantity,NULL))as day_28, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),29,quantity,NULL))as day_29, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),30,quantity,NULL))as day_30, " & _
				" Sum ( decode (to_number (to_char (h.receptiondate,'dd')),31,quantity,NULL))as day_31," & _
				" Sum ( h.QUANTITY) as Total" & _
				" From harvesting h" & _
				" Group by h.FCBA, to_char(h.receptiondate,'dd-mm-yyyy'), h.FIELDCODE" & _
				" union all" & _
				" Select ag.FCBA, to_char(ag.fddate,'dd-mm-yyyy') as Tanggal, ag.LOCATIONCODE AS BLOK,'1. PANEN' as Janjang," & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),1, unit,NULL))as day_1, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),2, unit,NULL))as day_2, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),3, unit,NULL))as day_3, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),4, unit,NULL))as day_4, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),5, unit,NULL))as day_5, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),6, unit,NULL))as day_6, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),7, unit,NULL))as day_7, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),8, unit,NULL))as day_8, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),9, unit,NULL))as day_9, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),10,unit,NULL))as day_10, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),11,unit,NULL))as day_11, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),12,unit,NULL))as day_12, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),13,unit,NULL))as day_13, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),14,unit,NULL))as day_14, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),15,unit,NULL))as day_15, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),16,unit,NULL))as day_16, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),17,unit,NULL))as day_17, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),18,unit,NULL))as day_18, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),19,unit,NULL))as day_19, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),20,unit,NULL))as day_20, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),21,unit,NULL))as day_21, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),22,unit,NULL))as day_22, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),23,unit,NULL))as day_23, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),24,unit,NULL))as day_24, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),25,unit,NULL))as day_25, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),26,unit,NULL))as day_26, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),27,unit,NULL))as day_27, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),28,unit,NULL))as day_28, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),29,unit,NULL))as day_29, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),30,unit,NULL))as day_30, " & _
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),31,unit,NULL))as day_31," & _
				" Sum ( ag.UNIT ) as Total" & _
				" from attendance_gad ag" & _
				" Where jobcode = '6502101' " & _
				" Group by ag.FCBA, to_char(ag.fddate,'dd-mm-yyyy') , ag.LOCATIONCODE)" & _
				" WHERE FCBA = '" & Session ("FCBA") & "'"


if  LEN(TRIM(v_txtYear)) > 0 then
	strEmployee	=	strEmployee & " and to_number(to_char(to_date(tanggal,'dd-mm-yyyy'),'yyyy')) = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " and to_number(to_char(to_date(tanggal,'dd-mm-yyyy'),'mm')) =" & v_txtMonth
end if
if  LEN(TRIM(v_txtLocationCode)) > 0 then
	strEmployee	=	strEmployee & " AND BLOK = '" & v_txtLocationCode & "'"
end if
	strEmployee	=	strEmployee & " ORDER BY BLOK, TANGGAL, JANJANG ASC"
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Cross Check Janjang Panen vs Kirim</title>
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
		<h2>Cross Check Janjang Panen vs Kirim</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
				<td><font face="verdana" size="1">BLOK</font></td><td><select name="txtLocationCode" ><option value=""></option>
		<%
		Dim rsLocationCode, strLocationCode
		strLocationCode	=	" SELECT FCCODE, FCNAME FROM FIELD WHERE  FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
		
		Set rsLocationCode = Server.CreateObject("ADODB.Recordset")
		rsLocationCode.Open strLocationCode, connstr, adOpenStatic 
		
		if not rsLocationCode.eof then
		Do While Not rsLocationCode.eof
		if rsLocationCode("FCCODE") = v_txtLocationCode then%>
		        <option value="<%=rsLocationCode("FCCODE")%>" Selected><%=rsLocationCode("FCCODE")%></option> 
		     <%Else%>
		        <option value="<%=rsLocationCode("FCCODE")%>"><%=rsLocationCode("FCCODE")%></option>  
		    <%End if
			rsLocationCode.movenext
		loop
		end if
		rsLocationCode.Close
		set rsLocationCode = Nothing
		%>
		</select></td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th>NO</th>
		<th>BLOK</th>
		<th>JANJANG</th>
<%for nday = 1 to 31%>
<th><%=nday%></th>
<%next%>
<th>TOTAL</th>
	</tr>

<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic

if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td colspan="36">NO DATA AVAILABLE</td>
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
				<td align="left"><%=rsEmployee("BLOK")%></td>
				<td align="left"><%=rsEmployee("JANJANG")%></td>
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
