<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtdatefrom, v_txtdateuntil
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultDate, varDefaultMonth
Dim ScriptName, v_txtsectionname, v_txtemployeecode, v_txtmethods
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName, no

v_txtFcba = Request("txtFcba")
v_txtdatefrom = Request("txtdatefrom")
v_txtdateuntil = Request("txtdateuntil")
v_txtsectionname = Request("txtsectionname")
v_txtemployeecode = Request("txtemployeecode")
v_txtmethods = Request("txtmethods")
txtpresentgrade = Request("txtpresentgrade")
ScriptName = Request.ServerVariables("SCRIPT_NAME")

strDefaultBa = " select" & _
               "     fccode as fcba" & _
               " from businessunit bu" & _
               " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
               " where isdefault = 'TRUE'" & _
               "       and fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("fcba")
Else
	varDefaultBa = Session("fcba")
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

strcurrentdate = " select currentdate" & _
                 " from periodcontrol" & _                 
                 " where fcba = '" & v_txtFcba & "'"

Set rscurrentdate = Server.CreateObject("ADODB.Recordset")
rscurrentdate.Open strcurrentdate, connstr, adOpenStatic

If Not rscurrentdate.EOF Then
	varDefaultDate	=	rscurrentdate("currentdate")
Else
	varDefaultDate	=	Session("v_Employee")
End If
rscurrentdate.Close
set rscurrentdate = Nothing

If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(varDefaultDate)
If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(varDefaultDate)


strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			"     , FCMANAGER" & _
			"     , FCKTU" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varManagerName = rsBaInfo("FCMANAGER")
	varKTUName = rsBaInfo("FCKTU")
Else
	varFcbaName = ""
	varCompanyName = ""
	varManagerName = ""
	varKTUName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing
%>
<html>
<head>
	<title>ePlaS Reporting - Laporan Sidik Jari</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
	<link rel="stylesheet" href="/css/font/font-awesome.css"/>
	<link rel="stylesheet" href="/css/jquery-ui.css" />
	<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', window);
	  }, false
	);
	</script>
	<script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
	<script>
		$(function() {
		$( ".datepicker" ).datepicker();
		});
	</script>
	<script language="JavaScript">
	var newwin;
	function launchwin(winurl,winname,winfeatures)
	{
		//This launches a new window and then
		//focuses it if window.focus() is supported.
		newwin = window.open(winurl,winname,winfeatures);

	}
	</script>
	
<style type="text/css">
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:24%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:24%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
</head>

<body>
	<div id="wrapper">	
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
			<form name="frmAR" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">Dari</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td>
						<td>
							<font face="verdana" size="1">Sampai</font>
						</td>
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td>
						<td>
							<font face="verdana" size="1">Nama seksi</font>
						</td>
						<td>
							<select name="txtsectionname" >
								<option value=""></option>
						<%
						Dim rssectionname, strsectionname
						strsectionname	=	" select fccode, fcname from sectionname order by fccode"
						
						Set rssectionname = Server.CreateObject("ADODB.Recordset")
						rssectionname.Open strsectionname, connstr, adOpenStatic 
						
						if not rssectionname.eof then
						Do While Not rssectionname.eof
						if rssectionname("fccode") = v_txtsectionname then%>
								<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> - <%=rssectionname("fcname")%></option> 
							 <%Else%>
								<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> - <%=rssectionname("fcname")%></option>  
							<%End if
							rssectionname.movenext
						loop
						end if
						rssectionname.Close
						set rssectionname = Nothing
						%>
						</select></td>
						
						<td><font face="verdana" size="1">Grade</font></td>
						<td>
							<select name="txtpresentgrade" >
								<option value=""></option>
						<%
						strsectionname	= " SELECT fccode, fcname from parameter WHERE parcode = 'E0001' AND fcba = 'JJP2' ORDER BY rate_1"
						
						Set rssectionname = Server.CreateObject("ADODB.Recordset")
						rssectionname.Open strsectionname, connstr, adOpenStatic 
						
						if not rssectionname.eof then
						Do While Not rssectionname.eof
						if rssectionname("fccode") = txtpresentgrade then%>
								<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fcname")%></option> 
							 <%Else%>
								<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fcname")%></option>  
							<%End if
							rssectionname.movenext
						loop
						end if
						rssectionname.Close
						set rssectionname = Nothing
						%>
						</select></td>
						<td>
							<font face="verdana" size="1">Kode Karyawan</font>
						</td>
						<td>
							<input id="fccode" type="text" size="15" name="txtemployeecode" value="<%=v_txtemployeecode%>">
						</td>
						
						<td>
							<font face="verdana" size="1">Tipe</font>
						</td>
						<td>
							<select name="txtmethods" >
								<option value="" <%if len(trim(v_txtmethods)) = 0 then%> Selected <%end if%>>Semua</option>
								<option value="AF" <%if v_txtmethods = "AF" then%> Selected <%end if%>>Automatic Finger</option>
								<option value="MF" <%if v_txtmethods = "MF" then%> Selected <%end if%>>Manual Finger</option>
							</select>
						</td>
						
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>
		
	<table border="0" width="100%" id="product-table" class="xthf-site">
		<thead>
			<tr>
				<th colspan="15" style="background-color:#FFFFFF; border:none;">
					<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<h4>PT. <%=varCompanyName%></h4>
						<h5><%=varFcbaName%></h5>
					</div>
					<div id="subheader" class="centered">
						<h2>LAPORAN DATA SIDIK JARI ( ABSENSI FINGER PRINT )</h2>
						<br/>
						<h4>
							Dari tanggal <%=v_txtdatefrom%> sampai <%=v_txtdateuntil%>
						</h4>
					</div>
					</div>
				</th>
			</tr>
		<tr>
			<th rowspan="2" style="color:#fff;">Kode Karyawan</th>
			<th rowspan="2" style="color:#fff;">Nama Karyawan</th>
			<th rowspan="2" style="color:#fff;">Divisi</th>
			<th rowspan="2" style="color:#fff;">Gang</th>
			<th rowspan="2" style="color:#fff;">Tanggal</th>			
			<th colspan="2" style="color:#fff;">Sidik Jari</th>
			<th colspan="3" style="color:#fff;">Jam Kerja</th>
			<th rowspan="2" style="color:#fff;">Terlambat</th>
			<th rowspan="2" style="color:#fff;">Pulang Cepat</th>
			<th rowspan="2" style="color:#fff;">Tipe</th>
			<th rowspan="2" style="color:#fff;">No. Berita Acara</th>
			<th rowspan="2" style="color:#fff;">Alasan</th>
		</tr>
		<tr>
			<th style="color:#fff;">Check In</th>
			<th style="color:#fff;">Check Out</th>
			<th style="color:#fff;">Model</th>
			<th style="color:#fff;">Mulai</th>
			<th style="color:#fff;">Selesai</th>
		</tr>
	</thead>
	<tbody>
<%
	Dim rsfp, strfp, noline 
	Dim nGet_Hk, nGet_Hk_kj, nget_mkp1, nBase_Salary, nrapelgaji, nOvertime, nPremi, nTunjangan, nBorongan, nP_Jamsostek, nP_bpjskes, nP_Mk_P1, nP_Panen, nP_Pph21, nP_koperasi, nP_Lain, nTotal
		strfp = " SELECT DISTINCT ev.empno user_id" & _
                "        , em.fccode employeecode" & _
                "        , em.fcname employeename" & _
                "        , ev.fcba" & _
                "        , em.sectionname sectioncode" & _
				"        , em.gangcode gangcode" & _
				"        , em.presentgrade presentgrade" & _
				"        , ev.typeattend absensi" & _
                "        , to_char(ev.tdate, 'mm/dd/yyyy') tanggal" & _
                "        , to_char(ev.timein, 'hh24:mi') checkin" & _
                "        , to_char(ev.timeout, 'hh24:mi') checkout" & _
                "        , ev.normaid" & _
                "        , ev.normaname" & _
                "        , to_char(ev.starttime, 'hh24:mi') jammulai" & _
                "        , to_char(ev.endtime, 'hh24:mi') jamselesai" & _
                "        , round(ev.late, 2) terlambat" & _
                "        , round(ev.comehomeearly, 2) pulangcepat" & _
                "        , round(ev.ot, 2) jamlembur" & _
				"        , ev.datatype||' - '||decode(ev.datatype, 'AF', 'Automatic Finger', 'MF', 'Manual Finger', null) datatype" & _
				"        , ev.no_ba" & _
				"        , ev.alasan" & _
                " FROM userfinger.attendanceview ev" & _
                " LEFT JOIN employee em ON (ev.fcba = em.fcba AND ev.empno = em.fccode)" & _
                " WHERE ev.fcba = '" & v_txtFcba & "'"

if  len(trim(v_txtdatefrom)) > 0 then
	strfp = strfp & " AND ev.tdate >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
end if
if  len(trim(v_txtdateuntil)) > 0 then
	strfp = strfp & " AND ev.tdate <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
end if
if  len(trim(v_txtsectionname)) > 0 then
	strfp = strfp & " AND em.sectionname = '" & v_txtsectionname & "'"
end if
if  len(trim(v_txtemployeecode)) > 0 then
	strfp = strfp & " AND ev.empno = '" & v_txtemployeecode & "'"
end if
if  len(trim(v_txtmethods)) > 0 then
	strfp = strfp & " AND ev.datatype = '" & v_txtmethods & "'"
end if

if "" & txtpresentgrade <> "" then strfp = strfp & " AND em.presentgrade = '" & txtpresentgrade & "'"

			
' strfp = strfp & " ORDER BY em.sectionname,em.gangcode,em.presentgrade,em.fccode,ev.tdate" 
' response.write strfp
Set rsfp = Server.CreateObject("ADODB.Recordset")
rsfp.Open strfp, connstr, adOpenStatic
    

	if Not rsfp.eof then
		while Not rsfp.eof
%>
		  <tr>
			<td align="center"><%=rsfp("employeecode")%></td>
			<td align="left"><%=rsfp("employeename")%></td>
			<td align="center"><%=rsfp("sectioncode")%></td>
			<td align="center"><%=rsfp("gangcode")%></td>
			<td align="center"><%= FormatDateTime(rsfp("tanggal"), 1) %></td>
			
			<td align="center"><%=rsfp("checkin")%></td>
			<td align="center"><%=rsfp("checkout")%></td>
			<td align="center"><%=rsfp("normaname")%></td>
			<td align="center"><%=rsfp("jammulai")%></td>
			<td align="center"><%=rsfp("jamselesai")%></td>
			<td align="right"><%=FbpNum2(cdbl(rsfp("terlambat")))%></td>
			<td align="right"><%=FbpNum2(cdbl(rsfp("pulangcepat")))%></td>
			<td align="left"><%=rsfp("datatype")%></td>
			<td align="left"><%=rsfp("no_ba")%></td>
			<td align="left"><%=rsfp("Alasan")%></td>
		  </tr>
    <%
			rsfp.MoveNext
		wend
	Else
	%>
	<tr>
		<td Colspan = "15">NO DATA AVAILABLE</td>
	</tr>
	<%End If%>
	<%
	rsfp.Close
	set rsfp = Nothing
	%>

	</tbody>
</table>
<div id="footer_note">	
	<div class="block">Prepared by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
    <div class="block">Checked by,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <input class="isian" type="text" value="">
    </div>
    <div class="block">Approved by,
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="<%=varManagerName%>">
        <hr/>
        <input class="isian" type="text" value="MANAGER">
    </div>
	<div class="block">Acknowledged by,
        <br/><br/><br/><br/>
        <input class="isian" type="text" value="" >
        <hr/>
        <input class="isian" type="text" value="">
    </div>
 </div>
 <div class="clear"></div>
	<!--
<%
 if len(trim(txtAssistqGang)) > 0 then%>
	

   <table width="90%" class="data" style="margin:auto;">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="30%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="30%" align="center"><font face="Verdana" size="1">Directed by,</font></td>
		<td width="30%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("FCMANAGER")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("KTU")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center" ><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=txtAssistqGang%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i><%=txtPostqGang%></i></strong></td>
	   </tr>
   </table>
   <%else%>
   <table width="90%" class="data" style="margin:auto;">
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
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("FCMANAGER")%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=Session("KTU")%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>

<%end if%>  -->
   </body>
   </html>