<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode
Dim ScriptName
Dim rsCatuBeras, strCatuBeras
Dim NO, cClass, nTotalPekerja, nTotalPasangan, nTotalTanggungan, nSubTotal

v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")
v_txtGangCode=	Request("txtGangCode")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Employee"))
end if
	
strCatuBeras	=	" SELECT FCBA," & _
					" EMPLOYEECODE," & _
					" EMPLOYEENAME," & _
					" GANGCODE," & _
					" POSITION," & _
					" PRESENTGRADE," & _
					" MARITALSTATUS," & _
					" PEKERJA," & _
					" PASANGAN," & _
					" TANGGUNGAN," & _
					" TOTAL" & _
					" FROM" & _
					" (" & _
					" SELECT EM.FCBA FCBA," & _
					" EM.FCCODE EMPLOYEECODE," & _
					" EM.FCNAME EMPLOYEENAME," & _
					" EM.GANGCODE GANGCODE," & _
					" EP.DESCRIPTION POSITION," & _
					" EM.PRESENTGRADE PRESENTGRADE," & _
					" EM.MARITALSTATUS MARITALSTATUS," & _
					" PRM.RATE_2 PEKERJA," & _
					" PRM.RATE_3 PASANGAN," & _
					" PRM.RATE_4 TANGGUNGAN," & _
					" PRM.RATE_5 TOTAL" & _
					" FROM EMPLOYEE EM, PARAMETER PRM, EMPPOSITION EP" & _
					" WHERE" & _
					" (EM.FCBA = PRM.FCBA AND EM.MARITALSTATUS = PRM.FCCODE)" & _
					" AND (EM.FCBA = EP.FCBA AND EM.POSITION = EP.POSITION)" & _
					" AND EM.DATETERMINATE IS NULL AND EM.NATURETYPE NOT IN ('NN')" & _
					" )" & _
					" WHERE FCBA = '" & Session("FCBA") & "'"
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strCatuBeras	=	strCatuBeras & " AND GANGCODE = '" & v_txtGangCode & "'"
end if				

Set rsCatuBeras = Server.CreateObject("ADODB.Recordset")
rsCatuBeras.Open strCatuBeras, connstr, adOpenStatic 
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Catu Beras</title>
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
		<h2>RINCIAN CATU BERAS KARYAWAN</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" readonly="true" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" readonly="true" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
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
		<th rowspan="2">NO</th>
		<th rowspan="2">NO PEKERJA</th>
		<th rowspan="2">NAMA PEKERJA</th>
		<th rowspan="2">JABATAN</th>
		<th rowspan="2" width="4%">STATUS</th>
		<th rowspan="2" width="4%">MARITAL STATUS</th>
		<th colspan="4">TUNJANGAN BERAS</th>
		<th rowspan="2" colspan="2">TANDA TANGAN</th>
	</tr>
	<tr>
		<th>PEKERJA</th>
		<th>PASANGAN</th>
		<th>TANGGUNGAN</th>
		<th>TOTAL</th>
	</tr>

<%if rsCatuBeras.eof and rsCatuBeras.eof then%>
	<tr>
		<td colspan="7" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		nTotalPekerja = 0
		nTotalPasangan = 0
		nTotalTanggungan = 0
		nSubTotal = 0
		while not rsCatuBeras.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center" height="30" width="3%"><%=NO%></td>
				<td align="center" width="4%"><%=rsCatuBeras("EMPLOYEECODE")%></td>
				<td align="left"><%=rsCatuBeras("EMPLOYEENAME")%></td>
				<td align="left"><%=rsCatuBeras("POSITION")%></td>
				<td align="center"><%=rsCatuBeras("PRESENTGRADE")%></td>
				<td align="center"><%=rsCatuBeras("MARITALSTATUS")%></td>
				<td align="right" width="4%"><%=rsCatuBeras("PEKERJA")%></td>
				<td align="right" width="4%"><%=rsCatuBeras("PASANGAN")%></td>
				<td align="right" width="4%"><%=rsCatuBeras("TANGGUNGAN")%></td>
				<td align="right" width="4%"><%=rsCatuBeras("TOTAL")%></td>
			<%if no  mod 2 = 1 then%>
				   <td align="left" width="100" bgcolor="#ccc"><font face="Verdana" size="1">&nbsp;</font></td>			
				   <td align="left" width="100" bgcolor="#fff"><font face="Verdana" size="1">&nbsp;</font></td>			
			   <%Else%>
				   <td align="left" width="100" bgcolor="#fff"><font face="Verdana" size="1">&nbsp;</font></td>			
				   <td align="left" width="100" bgcolor="#ccc"><font face="Verdana" size="1">&nbsp;</font></td>			
			   <%End if%>
			</tr>
<%
			NO = NO + 1
			
			nTotalPekerja = nTotalPekerja + Cdbl(rsCatuBeras("PEKERJA"))
			nTotalPasangan = nTotalPasangan + Cdbl(rsCatuBeras("PASANGAN"))
			nTotalTanggungan = nTotalTanggungan + Cdbl(rsCatuBeras("TANGGUNGAN"))
			nSubTotal = nSubTotal + Cdbl(rsCatuBeras("TOTAL"))
			rsCatuBeras.MoveNext
		Wend
	rsCatuBeras.Close
	set rsCatuBeras = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="6">TOTAL</td>
				<td align="right"><%=nTotalPekerja%></td>
				<td align="right"><%=nTotalPasangan%></td>
				<td align="right"><%=nTotalTanggungan%></td>
				<td align="right"><%=nSubTotal%></td>
				<td colspan="2">&nbsp;</td>
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
