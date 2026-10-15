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

strEmployee	=	" select *"&_
				" from ("&_
				" Select ag.fcba,to_char(ag.fddate,'dd-mm-yyyy') as Tanggal ,ag.jobcode,j.FCNAME,ag.locationtype,ag.locationcode,'HK'as Uraian ,"&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),1, mandays,NULL))as day_1, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),2, mandays,NULL))as day_2, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),3, mandays,NULL))as day_3, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),4, mandays,NULL))as day_4, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),5, mandays,NULL))as day_5, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),6, mandays,NULL))as day_6, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),7, mandays,NULL))as day_7, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),8, mandays,NULL))as day_8, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),9, mandays,NULL))as day_9, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),10,mandays,NULL))as day_10, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),11,mandays,NULL))as day_11, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),12,mandays,NULL))as day_12, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),13,mandays,NULL))as day_13, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),14,mandays,NULL))as day_14, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),15,mandays,NULL))as day_15, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),16,mandays,NULL))as day_16, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),17,mandays,NULL))as day_17, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),18,mandays,NULL))as day_18, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),19,mandays,NULL))as day_19, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),20,mandays,NULL))as day_20, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),21,mandays,NULL))as day_21, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),22,mandays,NULL))as day_22, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),23,mandays,NULL))as day_23, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),24,mandays,NULL))as day_24, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),25,mandays,NULL))as day_25, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),26,mandays,NULL))as day_26, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),27,mandays,NULL))as day_27, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),28,mandays,NULL))as day_28, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),29,mandays,NULL))as day_29, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),30,mandays,NULL))as day_30, "&_
				" Sum ( decode (to_number (to_char (ag.fddate,'dd')),31,mandays,NULL))as day_31"&_
				" From attendance_gad ag, job j"&_
				" Where ag.FCBA = j.FCBA and ag.JOBCODE = j.FCCODE and locationtype ='FF'"&_
				" Group by ag.fcba,to_char(ag.fddate,'dd-mm-yyyy'),ag.jobcode,j.FCNAME,ag.locationtype,ag.locationcode"&_
				" union all"&_
				" Select i.fcba, to_char(i.sivdate,'dd-mm-yyyy') as Tanggal, id.JOBCODE, j.FCNAME,id.LOCATIONTYPE, id.LOCATIONCODE, st.FCNAME as Uraian,"&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),1, id.quantity,NULL))as day_1, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),2, id.quantity,NULL))as day_2, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),3, id.quantity,NULL))as day_3, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),4, id.quantity,NULL))as day_4, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),5, id.quantity,NULL))as day_5, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),6, id.quantity,NULL))as day_6, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),7, id.quantity,NULL))as day_7, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),8, id.quantity,NULL))as day_8, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),9, id.quantity,NULL))as day_9, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),10,id.quantity,NULL))as day_10, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),11,id.quantity,NULL))as day_11, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),12,id.quantity,NULL))as day_12, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),13,id.quantity,NULL))as day_13, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),14,id.quantity,NULL))as day_14, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),15,id.quantity,NULL))as day_15, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),16,id.quantity,NULL))as day_16, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),17,id.quantity,NULL))as day_17, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),18,id.quantity,NULL))as day_18, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),19,id.quantity,NULL))as day_19, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),20,id.quantity,NULL))as day_20, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),21,id.quantity,NULL))as day_21, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),22,id.quantity,NULL))as day_22, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),23,id.quantity,NULL))as day_23, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),24,id.quantity,NULL))as day_24, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),25,id.quantity,NULL))as day_25, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),26,id.quantity,NULL))as day_26, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),27,id.quantity,NULL))as day_27, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),28,id.quantity,NULL))as day_28, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),29,id.quantity,NULL))as day_29, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),30,id.quantity,NULL))as day_30, "&_
				" Sum ( decode (to_number (to_char (i.sivdate,'dd')),31,id.quantity,NULL))as day_31"&_
				" From ISSUE i, issuedetails id, stock st, job j "&_
				" Where i.FCBA = id.FCBA and  i.SIVCODE = id.SIVCODE and st.FCBA = id.fcba and st.FCCODE = id.STOCKCODE and id.FCBA = j.FCBA and id.JOBCODE = j.FCCODE "&_
				" and locationtype ='FF' "&_
				" Group by i.fcba, to_char(i.sivdate,'dd-mm-yyyy'), id.JOBCODE, j.FCNAME,id.LOCATIONTYPE, id.LOCATIONCODE, st.FCNAME"&_
				" )" &_
				" WHERE FCBA = '" & Session ("FCBA") & "' "

if  LEN(TRIM(v_txtYear)) > 0 then
	strEmployee	=	strEmployee & " and to_number(to_char(to_date(tanggal,'dd-mm-yyyy'),'yyyy')) = " & v_txtYear
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strEmployee	=	strEmployee & " and to_number(to_char(to_date(tanggal,'dd-mm-yyyy'),'mm')) =" & v_txtMonth
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strEmployee	=	strEmployee & " AND jobcode = '" & v_txtGangCode & "'"
end if
if  LEN(TRIM(v_txtLocationCode)) > 0 then
	strEmployee	=	strEmployee & " AND LocationCode = '" & v_txtLocationCode & "'"
end if
	strEmployee	=	strEmployee & " ORDER BY JOBCODE, LOCATIONCODE, URAIAN"
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Cross Check HK vs Bahan</title>
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
		<h2>Cross Check HK vs Bahan</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
		<td><font face="verdana" size="1">JOBCODE</font></td><td><select name="txtGangCode" ><option value=""></option>
		<%
		Dim rsGangCode, strGangCode
		strGangCode	=	" SELECT FCCODE, FCNAME FROM JOB WHERE fccode in "&_
		" ('1310209','1310210','1310211','1310212','1310307','1310311','1310312','1310313','1310314', "&_
		"  '1320107','1320108','1320401','1320402','1320404','1320504','1320505','1320506','1322002', "&_
		"  '1322004','1322101','1322102','1322204','1322206','1322207','1322208','1322210','1322211', "&_
		"  '1322306','1322307','1322308','1322309','1322310','1322311','1322313','1322314','1322601', "&_
		"  '6500102','6500104','6500201','6500202','6500203','6500304','6500306','6500307','6500308', "&_
		"  '6500310','6500311','6500405','6500406','6500407','6500408','6500409','6500410','6500411', "&_
		"  '6500701','6500702','6500807') "&_
		"  and FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC" 
		
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
				<td><font face="verdana" size="1">LOCATIONCODE</font></td><td><select name="txtLocationCode" ><option value=""></option>
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
		<th>JOB CODE</th>
		<th>JOB DESCRIPTION</th>
		<th>TARGET</th>
		<th>URAIAN</th>
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
				<td align="center"><%=rsEmployee("JOBCODE")%></td>
				<td align="left"><%=rsEmployee("FCNAME")%></td>
				<td align="center"><%=rsEmployee("LOCATIONCODE")%></td>
				<td align="center"><%=rsEmployee("URAIAN")%></td>
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
