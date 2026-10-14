<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtFcCode, v_txtFcName, v_txtGangCode, v_txtDivisi

v_txtFcCode		=	trim(Request("txtFcCode"))
v_txtFcName		=	trim(Request("txtFcName"))
v_txtGangCode	=	trim(Request("txtGangCode"))
v_txtDivisi		=	trim(Request("txtDivisi"))
ScriptName		=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Data Master Employee</title>
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
<style>
@media print
{
body {
font-family: 'Arial Narrow', Verdana, Helvetica, sans-serif;
font-size:12px;
background-color:#fff;
color:#000;
}
h1{font-size:16px;}
h2{font-size:14px;}
h3{font-size:12px;}
h4{font-size:11px;}
h5{font-size:10px;}
h6{font-size:9px;}
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<form name="frmAR" method="post" action="<%=ScriptName%>">
						<td><font face="verdana" size="1">KODE</font></td>
						<td><input type="text" maxlength="10" name="txtFcCode" size="10" value="<%=v_txtFcCode%>" title="Input Code"></td>		
						<td><font face="verdana" size="1">NAME</font></td>
						<td><input type="text" maxlength="10" name="txtFcName" size="30" value="<%=v_txtFcName%>" title="Input Name"></td>
						
						<td><font face="verdana" size="1">DIVISI</font></td>
						<td>
							<select name="txtDivisi" >
								<option value=""></option>
						<%
						Dim rsDivisi, strDivisi
						strDivisi	=	" SELECT GANGGROUP FROM (" & _
										" SELECT FCBA, FCCODE, FCNAME,DECODE(GANGGROUP," & _
										" 'KM','KEAMANAN'," & _
										" 'MB','BIBITAN'," & _
										" 'MT','MAINTENANCE'," & _
										" 'OS','OPERATOR/ SUPIR'," & _
										" 'TM','TUKANG MUAT'," & _
										" 'BB','BIBITAN'," & _
										" 'BR','BORONGAN'," & _
										" 'KT','KANTOR'," & _
										" 'SP','SEMPROT'," & _
										" 'TB','SIPIL ENGINEERING'," & _
										" 'MD','MANDOR'," & _
										" 'PR','PROSES'," & _
										" 'LB','LABORTORIUM'," & _
										" 'UM','UMUM'," & _
										" GANGGROUP) GANGGROUP FROM (" & _
										" SELECT FCBA, FCCODE, FCNAME, DECODE(SUBSTR(FCCODE,3,2)," & _
										" '01','AFD-1'," & _
										" '02','AFD-2'," & _
										" '03','AFD-3'," & _
										" '04','AFD-4'," & _
										" '05','AFD-5'," & _
										" '06','AFD-6'," & _
										" '07','AFD-7'," & _
										" '08','AFD-8'," & _
										" '00',SUBSTR(FCCODE,1,2)," & _
										" FCCODE) GANGGROUP FROM GANG))" & _
										" WHERE FCBA = '" & Session("FCBA") & "'" & _
										" GROUP BY GANGGROUP"

						
						Set rsDivisi = Server.CreateObject("ADODB.Recordset")
						rsDivisi.Open strDivisi, connstr, adOpenStatic 
						
						if not rsDivisi.eof then
						Do While Not rsDivisi.eof
						if rsDivisi("GANGGROUP") = v_txtDivisi then%>
								<option value="<%=rsDivisi("GANGGROUP")%>" Selected><%=rsDivisi("GANGGROUP")%></option> 
							<%Else%>
								<option value="<%=rsDivisi("GANGGROUP")%>"><%=rsDivisi("GANGGROUP")%></option>  
							<%End if
							rsDivisi.movenext
						loop
						end if
						rsDivisi.Close
						set rsDivisi = Nothing
						%>
							</select>
						</td>						
						
						<td><font face="verdana" size="1">GANGCODE</font></td>
						<td>
							<select name="txtGangCode" >
								<option value=""></option>
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
							</select>
						</td>

						<td><input type="submit" value="View" /></td>
						</form>
					</tr>
				</table>
			</div>
<%
	Dim rsqGang, qGang, txtqGang
	qGang	=	" SELECT FCNAME FROM GANG WHERE FCBA = '" & Session("FCBA") & "' AND FCCODE = '"& v_txtGangCode & "'"
	
	Set rsqGang = Server.CreateObject("ADODB.Recordset")
	rsqGang.Open qGang, connstr, adOpenStatic
	
	if not rsqGang.eof then
		txtqGang	=	rsqGang("FCNAME")
	end if
		rsqGang.Close
		set rsqGang = Nothing
%>
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered">
			<h2>DATA KARYAWAN <%if v_txtDivisi <> "" then%> <%=v_txtDivisi%> <%end if%></h2>
			<% IF v_txtGangCode <> "" then %>
			<h4>GANG : <%=v_txtGangCode%> - <%=txtqGang%></h4>
			<%END IF%>
			 
		</div>
	</div>
	
<table border="0" width="100%" id="product-table" class="headertable">
     <tr>
		<th style="color:#fff;">NO</th>
		<th style="color:#fff;">NIK</th>		
		<th style="color:#fff;">NAMA KARYAWAN</th>
		<th style="color:#fff;">GANG</th>
		<th style="color:#fff;">DIVISI</th>
		<th style="color:#fff;">TANGGAL LAHIR</th>
		<th style="color:#fff;">AGAMA</th>
		<th style="color:#fff;">SUKU</th>
		<th style="color:#fff;">NO. KTP</th>
		<th style="color:#fff;">PENDIDIKAN</th>		
		<th style="color:#fff;">TANGGAL BERGABUNG</th>		
		<th style="color:#fff;">STATUS</th>
		<th style="color:#fff;">STATUS PEKAWINAN</th>
		<th style="color:#fff;">NAMA PASANGAN</th>
		<th style="color:#fff;">JABATAN</th>
		<th style="color:#fff;">TANGGAL BERHENTI</th>
	</tr>
	<%
	Dim rsEmployee, strEmployee, noline 

			strEmployee	=	" SELECT" & _
							" EM.FCCODE," & _
							" EM.FCNAME," & _
							" GANGCODE," & _
							" GANGGROUP," & _
							" DATEOFBIRTH," & _
							" RELIGION," & _
							" RACE," & _
							" IDENTITYCARD," & _
							" HIGHESTEDUCATION," & _
							" DATEOFJOIN," & _
							" PRESENTGRADE," & _
							" PRESENTGRADE_DETAIL," & _
							" MARITALSTATUS," & _
							" SPOUSE," & _
							" EP.DESCRIPTION," & _
							" DATETERMINATE" & _
							" FROM EMPLOYEE EM" & _
							" INNER JOIN EMPPOSITION EP ON (EM.FCBA = EP.FCBA AND EM.POSITION = EP.POSITION)" & _
							" INNER JOIN (SELECT FCBA, FCCODE, FCNAME, GANGGROUP FROM (" & _
							" SELECT FCBA, FCCODE, FCNAME,DECODE(GANGGROUP," & _
							" 'KM','KEAMANAN'," & _
							" 'MB','BIBITAN'," & _
							" 'MT','MAINTENANCE'," & _
							" 'OS','OPERATOR/ SUPIR'," & _
							" 'TM','TUKANG MUAT'," & _
							" 'BB','BIBITAN'," & _
							" 'BR','BORONGAN'," & _
							" 'KT','KANTOR'," & _
							" 'SP','SEMPROT'," & _
							" 'TB','SIPIL ENGINEERING'," & _
							" 'MD','MANDOR'," & _
							" 'PR','PROSES'," & _
							" 'LB','LABORTORIUM'," & _
							" 'UM','UMUM'," & _
							" GANGGROUP) GANGGROUP FROM (" & _
							" SELECT FCBA, FCCODE, FCNAME, DECODE(SUBSTR(FCCODE,3,2)," & _
							" '01','AFD-1'," & _
							" '02','AFD-2'," & _
							" '03','AFD-3'," & _
							" '04','AFD-4'," & _
							" '05','AFD-5'," & _
							" '06','AFD-6'," & _
							" '07','AFD-7'," & _
							" '08','AFD-8'," & _
							" '00',SUBSTR(FCCODE,1,2)," & _
							" FCCODE) GANGGROUP FROM GANG))) GG ON (EM.FCBA = GG.FCBA AND EM.GANGCODE = GG.FCCODE)" & _
							" WHERE FCBA = '" & SESSION("FCBA") & "'"
		if  len(trim(v_txtFcCode)) > 0 then
			strEmployee	=	strEmployee &  " AND EM.FCCODE = '" & v_txtFcCode & "'"
		end if
		if  len(trim(v_txtFcName)) > 0 then
			strEmployee	=	strEmployee &  " AND EM.FCNAME LIKE '%" & v_txtFcName & "%'"
		end if
		if  len(trim(v_txtDivisi)) > 0 then
			strEmployee	=	strEmployee &  " AND GG.GANGGROUP = '" & v_txtDivisi & "'"
		end if
		if  len(trim(v_txtGangCode)) > 0 then
			strEmployee	=	strEmployee &  " AND EM.GANGCODE = '" & v_txtGangCode & "'"
		end if
		strEmployee	=	strEmployee &  " ORDER BY EM.FCCODE ASC"
	
	Set rsEmployee = Server.CreateObject("ADODB.Recordset")
	rsEmployee.Open strEmployee, connstr, adOpenStatic
    
	if rsEmployee.bof and rsEmployee.eof then
%>
	<tr>
		<td colspan="16">NO DATA AVAILABLE</td>
	</tr>
<%
	elseif Not rsEmployee.eof then
		noline = 1
		
		while Not rsEmployee.eof  %>
		           
		  <tr>
			<td align="center"><%=noline%></td>
			<td align="left"><%=rsEmployee("FCCODE")%></td>
			<td align="left"><%=rsEmployee("FCNAME")%></td>
			<td align="left"><%=rsEmployee("GANGCODE")%></td>
			<td align="left"><%=rsEmployee("GANGGROUP")%></td>
			<td align="left"><%=rsEmployee("DATEOFBIRTH")%></td>
			<td align="left"><%=rsEmployee("RELIGION")%></td>
			<td align="left"><%=rsEmployee("RACE")%></td>
			<td align="left"><%=rsEmployee("IDENTITYCARD")%></td>
			<td align="left"><%=rsEmployee("HIGHESTEDUCATION")%></td>
			<td align="left"><%=rsEmployee("DATEOFJOIN")%></td>
			<td align="left"><%=rsEmployee("PRESENTGRADE")%></td>
			<td align="left"><%=rsEmployee("MARITALSTATUS")%></td>
			<td align="left"><%=rsEmployee("SPOUSE")%></td>
			<td align="left"><%=rsEmployee("DESCRIPTION")%></td>
			<td align="left"><%=rsEmployee("DATETERMINATE")%></td>
		  </tr>	
		      
    <%
			noline	=	noline + 1
			
			rsEmployee.MoveNext
		wend
	end if
	rsEmployee.Close
	Set rsEmployee = Nothing
	%> 
</table>

   <table width="90%" class="data">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
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