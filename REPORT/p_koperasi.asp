<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim ScriptName, v_txtFcCode, v_txtFcName, v_txtGangCode, v_txtDivisi

v_txtFcCode = Request("txtFcCode")

Dim v_txtisactive : v_txtisactive = Request("txtisactive")
ScriptName		=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Koperasi</title>
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
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<form name="frmAR" method="post" action="<%=ScriptName%>">
						<td><font face="verdana" size="1">Cari Kode atau Deskripsi</font></td>
						<td><input type="text" maxlength="50" name="txtFcCode" size="50" value="<%=v_txtFcCode%>" title="Input Code"></td>		

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
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="17" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader" class="centered" style="width:50%; margin:auto;">
			<h2>DATA KOPERASI</h2>
			 
		</div>
	</div>
	
</th>
</tr>
     <tr>		
		<th style="color:#fff;">Kode</th>
		<th style="color:#fff;">Deskripsi</th>
		<th style="color:#fff;">Tanggal Berdiri</th>
		<th style="color:#fff;">No Badan Hukum</th>
		<th style="color:#fff;">Tanggal Badan Hukum</th>
		<th style="color:#fff;">Alamat</th>
		<th style="color:#fff;">Kota</th>
		<th style="color:#fff;">Negara</th>
		<th style="color:#fff;">Telp</th>
		<th style="color:#fff;">Fax</th>
		<th style="color:#fff;">Bentuk Koperasi</th>
		<th style="color:#fff;">Jenis Koperasi</th>
		<th style="color:#fff;">Kelompok Koperasi</th>
		<th style="color:#fff;">Sektor Usaha</th>
		<th style="color:#fff;">NIK</th>
		<th style="color:#fff;">Sertifikasi</th>
		<th style="color:#fff;">Status</th>
	</tr>
 </thead>
 <tbody>   
    
	<%
	Dim rsEmployee, strEmployee, noline 

			strEmployee	= " SELECT fccode" & _
                          "        , fcname" & _
                          "        , datecreated" & _
                          "        , registrationno" & _
                          "        , registrationdate" & _
                          "        , address" & _
                          "        , city" & _
                          "        , state" & _
                          "        , phone" & _
                          "        , fax" & _
                          "        , formco" & _
                          "        , typeco" & _
                          "        , groupco" & _
                          "        , sectorco" & _
                          "        , idno" & _
                          "        , iscertified" & _
                          "        , activation" & _
                          " FROM koperasi" & _
                          " WHERE fcba = '" & session("fcBA") & "'"
		if  len(trim(v_txtFcCode)) > 0 then
			strEmployee	=	strEmployee & " AND (UPPER(fccode) LIKE UPPER('%" & v_txtFcCode & "%') OR UPPER(fcname) LIKE UPPER('%" & v_txtFcCode & "%') )"
		end if
		
		strEmployee	=	strEmployee &  " ORDER BY FCCODE ASC"
	
	Set rsEmployee = Server.CreateObject("ADODB.Recordset")
	rsEmployee.Open strEmployee, connstr, adOpenStatic
	
	arr_data_eof = true
	if not rsEmployee.eof then
		arr_data = rsEmployee.getrows()
		arr_data_eof = false
	end if
    
	if not arr_data_eof then
		for xi=0 to ubound(arr_data,2)
%>
	 	<tr>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(0,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(1,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(2,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(3,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(4,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(5,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(6,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(7,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(8,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(9,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(10,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(11,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(12,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(13,xi)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(14,xi)%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(15,xi)%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=arr_data(16,xi)%></td>
		</tr>
<%
		next
	else
		response.write "<tr><td colspan='17'>NO DATA AVAILABLE</td></tr>"
	end if
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
       </tbody>
   </table>
   </body>
   </html>