<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
 <%
Dim v_txtDateFrom, v_txtDateTo, v_txtGangCode, v_txtGrade, v_txtDivisi, v_txtMandorGangCode
Dim ScriptName
Dim rsEmployee, strEmployee, rsMandor, strMandor, NOMOR
Dim NO, cClass, nday
Dim gt_d27, gt_d28, gt_d29, gt_d30, gt_d31, gt_d32

v_txtDateFrom = Request("txtDateFrom")
v_txtDateTo = Request("txtDateTo")
v_txtDivisi = Request("txtDivisi")
v_txtMandorGangCode = Request("txtMandorGangCode")
v_txtGangCode = Request("txtGangCode")
v_txtGrade = Request("txtGrade")
ScriptName = Request.ServerVariables("SCRIPT_NAME")

If Len(Trim(v_txtDateFrom)) = 0 Then
	v_txtDateFrom = Month(Session("v_Employee")) & "/01/" & Year(Session("v_Employee"))
End If

If Len(Trim(v_txtDateTo)) = 0 Then
	v_txtDateTo = Month(Session("v_Employee")) & "/" & GetLastDay(Month(Session("v_Employee")), Year(Session("v_Employee"))) & "/" & Year(Session("v_Employee"))
End If

Dim grs, gsql, strmain, strmain_close, strmain_open, has_close_payroll, wherecond

gsql = " SELECT NVL(payroll_executed, 'FALSE') result" & _
       " FROM empalldedstatus" & _
	   " WHERE fcba = '" & Session ("FCBA") & "'" & _
	   "       AND year = '" & Year(v_txtDateFrom) & "'" & _
	   "       AND month = '" & Month(v_txtDateFrom) & "'"

'response.write gsql
Set grs = Server.CreateObject("ADODB.Recordset")
grs.Open gsql, connstr, adOpenStatic

if not grs.eof then
	has_close_payroll = grs("result")
else
	has_close_payroll = "FALSE"
end if

	
if  LEN(TRIM(v_txtDateFrom)) > 0 then
    wherecond = wherecond & " AND ADHOCDATE >= TO_DATE('" & v_txtDateFrom & "', 'MM/DD/YYYY')"
end if
if  LEN(TRIM(v_txtDateTo)) > 0 then
	wherecond = wherecond & " AND ADHOCDATE <= TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY')"
end if
if  LEN(TRIM(v_txtDivisi)) > 0 then
    wherecond = wherecond & " AND SECTIONNAME = '" & v_txtDivisi & "'"
end if

if  (LEN(TRIM(v_txtMandorGangCode)) > 0 and LEN(TRIM(v_txtGangCode)) > 0) then
    wherecond = wherecond & " AND (GANGCODE = '" & v_txtMandorGangCode & "'"
    wherecond = wherecond & " OR GANGCODE = '" & v_txtGangCode & "')"
elseif (LEN(TRIM(v_txtMandorGangCode)) > 0 and LEN(TRIM(v_txtGangCode)) = 0) then
    wherecond = wherecond & " AND GANGCODE = '" & v_txtMandorGangCode & "'"
elseif (LEN(TRIM(v_txtMandorGangCode)) = 0 and LEN(TRIM(v_txtGangCode)) > 0) then
    wherecond = wherecond & " AND GANGCODE = '" & v_txtGangCode & "'"
else
	wherecond = wherecond & " "
end if
	
strmain_close = " SELECT ep.fcba," & _
                "        ep.adhocdate," & _
                "        ep.empcode," & _
                "        em.fcname empname," & _
                "        em.sectionname," & _
                "        em.gangcode," & _
                "        ep.allowdedcode," & _
                "        ep.amount" & _
                " FROM empallded ep" & _
                " INNER JOIN employee_hist_tab em" & _
                " ON (ep.fcba = em.fcba AND ep.empcode = em.fccode AND em.year = TO_NUMBER(TO_CHAR(ep.adhocdate, 'yyyy')) AND em.month = TO_NUMBER(TO_CHAR(ep.adhocdate, 'mm')))" & _
                " WHERE ep.fcba = '" & Session ("FCBA") & "'" & _
                "       AND ep.allowdedcode IN ('A24', 'A25', 'A27', 'A37', 'A38', 'A39', 'A40', 'D09', 'D10', 'D12', 'D13', 'D14', 'D15', 'D16', 'D17', 'D19', 'D27', 'D28', 'D29', 'D30', 'D31', 'D32')" & _
                " " & wherecond & " "
				
strmain_open = " SELECT ep.fcba," & _
               "        ep.adhocdate," & _
               "        ep.empcode," & _
               "        em.fcname empname," & _
               "        em.sectionname," & _
               "        em.gangcode," & _
               "        ep.allowdedcode," & _
               "        ep.amount" & _
               " FROM empallded ep" & _
               " INNER JOIN employee em" & _
               " ON (ep.fcba = em.fcba AND ep.empcode = em.fccode)" & _
               " WHERE ep.fcba = '" & Session ("FCBA") & "'" & _
			   "       AND ep.allowdedcode IN ('A24', 'A25', 'A27', 'A37', 'A38', 'A39', 'A40', 'D09', 'D10', 'D12', 'D13', 'D14', 'D15', 'D16', 'D17', 'D19', 'D27', 'D28', 'D29', 'D30', 'D31', 'D32')" & _
               " " & wherecond & " "
			   


if has_close_payroll = "TRUE" then
	strmain = strmain_close
else
	strmain = strmain_open
end if
	
strMandor = "   SELECT FCBA," & _
            "          EMPCODE," & _
            "          EMPNAME," & _
            "          SECTIONNAME," & _
            "          GANGCODE," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'A24', AMOUNT, 0)) PREMILEBIHBORONG," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'A25', AMOUNT, 0)) PREMITANDANHARILIBUR," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'A27', AMOUNT, 0)) PREMISIAPBORONG," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'A37', AMOUNT, 0)) PREMIMANDORPANEN," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'A38', AMOUNT, 0)) PREMIKERANIPANEN," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'A39', AMOUNT, 0)) PREMIMANDOR1," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'A40', AMOUNT, 0)) PREMIKERANIDIVISI," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D09', AMOUNT, 0)) POTONGANBUAHMENTAH," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D10', AMOUNT, 0)) POTONGANTANGKAIPANJANG," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D12', AMOUNT, 0)) POTONGANBUAHTIDAKDIPANEN," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D13', AMOUNT, 0)) POTONGANPELEPAHTIDAKDIPOTONG," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D14', AMOUNT, 0)) POTONGANPELEPAHTIDAKDISUSUN," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D15', AMOUNT, 0)) POTONGANBRONDOLANDIPIRINGAN," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D16', AMOUNT, 0)) POTONGANBUAHTINGGALDIPIRINGAN," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D17', AMOUNT, 0)) POTONGANBUAHMATAHARI," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D19', AMOUNT, 0)) POTONGANTIDAKDAPATBASIS," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D27', AMOUNT, 0)) D27," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D28', AMOUNT, 0)) D28," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D29', AMOUNT, 0)) D29," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D30', AMOUNT, 0)) D30," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D31', AMOUNT, 0)) D31," & _
            "          SUM (DECODE (ALLOWDEDCODE, 'D32', AMOUNT, 0)) D32," & _
            "          SUM (amount) TOTAL" & _
            "     FROM ( " & strmain & " )" & _
            " WHERE gangcode LIKE 'MD%'"
			
strMandor = strMandor & " GROUP BY FCBA," & _
                        "          EMPCODE," & _
                        "          EMPNAME," & _
                        "          SECTIONNAME," & _
                        "          GANGCODE"
strMandor = strMandor & " ORDER BY EMPCODE ASC"
			  
strEmployee = "   SELECT FCBA," & _
              "          EMPCODE," & _
              "          EMPNAME," & _
              "          SECTIONNAME," & _
              "          GANGCODE," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'A24', AMOUNT, 0)) PREMILEBIHBORONG," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'A25', AMOUNT, 0)) PREMITANDANHARILIBUR," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'A27', AMOUNT, 0)) PREMISIAPBORONG," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'A37', AMOUNT, 0)) PREMIMANDORPANEN," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'A38', AMOUNT, 0)) PREMIKERANIPANEN," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'A39', AMOUNT, 0)) PREMIMANDOR1," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'A40', AMOUNT, 0)) PREMIKERANIDIVISI," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D09', AMOUNT, 0)) POTONGANBUAHMENTAH," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D10', AMOUNT, 0)) POTONGANTANGKAIPANJANG," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D12', AMOUNT, 0)) POTONGANBUAHTIDAKDIPANEN," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D13', AMOUNT, 0)) POTONGANPELEPAHTIDAKDIPOTONG," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D14', AMOUNT, 0)) POTONGANPELEPAHTIDAKDISUSUN," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D15', AMOUNT, 0)) POTONGANBRONDOLANDIPIRINGAN," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D16', AMOUNT, 0)) POTONGANBUAHTINGGALDIPIRINGAN," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D17', AMOUNT, 0)) POTONGANBUAHMATAHARI," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D19', AMOUNT, 0)) POTONGANTIDAKDAPATBASIS," & _
			  "          SUM (DECODE (ALLOWDEDCODE, 'D27', AMOUNT, 0)) D27," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D28', AMOUNT, 0)) D28," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D29', AMOUNT, 0)) D29," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D30', AMOUNT, 0)) D30," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D31', AMOUNT, 0)) D31," & _
              "          SUM (DECODE (ALLOWDEDCODE, 'D32', AMOUNT, 0)) D32," & _
              "          SUM (amount) TOTAL" & _
              "     FROM ( " & strmain & " )" & _
              " WHERE gangcode NOT LIKE 'MD%'"

strEmployee = strEmployee & " GROUP BY FCBA," & _
                            "          EMPCODE," & _
                            "          EMPNAME," & _
                            "          SECTIONNAME," & _
                            "          GANGCODE"
strEmployee = strEmployee & " ORDER BY EMPCODE ASC"
%>
 <html>
  
<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Premi Panen Detail</title>
<link rel="stylesheet" href="/css/jquery-ui.css" />
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
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
     <script>
	$(function() {
	$( ".datepicker" ).datepicker();
	});
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

<style type="text/css">
@media print
{
.printhide{display:none;}
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
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
			<table id="form-table" class="centered-margin">
				<tr>
					<td><font face="verdana" size="1">Dari</font></td>
					<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
					<!--<a href="javascript:show_calendar('frmEntry.txtDateFrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
						<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
					</a>-->
				</td>
				<td><font face="verdana" size="1">Sampai</font></td>
				<td align="center">
					<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
					<!--<a href="javascript:show_calendar('frmEntry.txtDateTo');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
						<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
					</a>-->
				</td>
					<td><font face="verdana" size="1">AFDELING</font></td>
					<td><select name="txtDivisi" >
						<option value = "">--ALL--</option>
							<%
							Dim rsDivisi, strDivisi
							strDivisi = " SELECT * FROM PARAMETER WHERE PARCODE = 'E0002' AND FCCODE LIKE 'AFD%' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"

							Set rsDivisi = Server.CreateObject("ADODB.Recordset")
							rsDivisi.Open strDivisi, connstr, adOpenStatic 

							if not rsDivisi.eof then
							Do While Not rsDivisi.eof
							if rsDivisi("FCCODE") = v_txtDivisi then%>
							<option value="<%=rsDivisi("FCCODE")%>" Selected><%=rsDivisi("FCCODE")%></option>
							<%Else%>
							<option value="<%=rsDivisi("FCCODE")%>"><%=rsDivisi("FCCODE")%></option>  
							<%End if
							rsDivisi.movenext
							loop
							end if
							rsDivisi.Close
							set rsDivisi = Nothing
							%>
							</select>
					</td>
					<td><font face="verdana" size="1">MANDOR</font></td>
					<td><select name="txtMandorGangCode" >
						<option value = "">--ALL--</option>
							<%
							Dim rsMandorGangCode, strMandorGangCode
							strMandorGangCode = " SELECT * FROM GANG WHERE FCBA = '" & Session("Fcba") & "' AND FCCODE LIKE 'MD%'" & _
											" AND ACTIVATION = 'Y' ORDER BY FCCODE"

							Set rsMandorGangCode = Server.CreateObject("ADODB.Recordset")
							rsMandorGangCode.Open strMandorGangCode, connstr, adOpenStatic 

							if not rsMandorGangCode.eof then
							Do While Not rsMandorGangCode.eof
							if rsMandorGangCode("FCCODE") = v_txtMandorGangCode then%>
							<option value="<%=rsMandorGangCode("FCCODE")%>" Selected><%=rsMandorGangCode("FCCODE")%></option>
							<%Else%>
							<option value="<%=rsMandorGangCode("FCCODE")%>"><%=rsMandorGangCode("FCCODE")%></option>  
							<%End if
							rsMandorGangCode.movenext
							loop
							end if
							rsMandorGangCode.Close
							set rsMandorGangCode = Nothing
							%>
							</select>
					</td>
					<td><font face="verdana" size="1">GANG</font></td>
					<td><select name="txtGangCode" >
						<option value = "">--ALL--</option>
							<%
							Dim rsGangCode, strGangCode
							strGangCode = " SELECT * FROM GANG WHERE FCBA = '" & Session("Fcba") & "' AND FCCODE LIKE 'PN%'" & _
											" AND ACTIVATION = 'Y' ORDER BY FCCODE"

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
			</table>
			</div>
			</form>
		</div>
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					<th colspan="27" style="background-color:#FFFFFF; border:none;">  
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=Session("COMPANY")%></h4>
								<h5><%=Session("BA")%></h5>
							</div>
							<div id="subheader" class="centered" style="width:50%; margin:auto;">
								<h2>RINCIAN PREMI PANEN</h3>
								<br/>
								<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
									<h5>TANGGAL: <%=v_txtdateto%></h5>
								<% else %>
									<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
								<% end if %>
							</div>
						</div>
					</th>
				</tr>			
				<tr>
					<!--<th rowspan="2" style="color:#fff;">NO</th>-->
					<th rowspan="2" style="color:#fff;">No Pekerja</th>
					<th rowspan="2" style="color:#fff;">Nama Pekerja</th>
					<th rowspan="2" style="color:#fff;">Seksi</th>
					<th rowspan="2" style="color:#fff;">Gang</th>
					<th colspan="7" style="color:#fff;">Premi</th>
					<th colspan="15" style="color:#fff;">Potongan</th>
					<th rowspan="2" style="color:#fff;">Total</th>
				</tr>
				<tr>
					<th style="color:#fff;">Premi Mandor I</th>
					<th style="color:#fff;">Premi Mandor Panen</th>
					<th style="color:#fff;">Premi Kerani Panen</th>
					<th style="color:#fff;">Premi Kerani Divisi</th>
					<th style="color:#fff;">Premi Siap Borong</th>
					<th style="color:#fff;">Premi Lebih Borong</th>
					<th style="color:#fff;">Premi Tandan Hari Libur</th>
					<th style="color:#fff;">Potongan Buah Mentah</th>
					<th style="color:#fff;">Potongan Tangkai Panjang</th>
					<th style="color:#fff;">Potongan Buah Tidak Dipanen</th>
					<th style="color:#fff;">Potongan Pelepah Tdk Dipotong</th>
					<th style="color:#fff;">Potongan Pelepah Tdk Disusun Rapi</th>
					<th style="color:#fff;">Potongan Brondolan Dipiringan</th>
					<th style="color:#fff;">Potongan Buah Tinggal Dipiringan</th>
					<th style="color:#fff;">Potongan Buah Matahari</th>
					<th style="color:#fff;">Potongan Tidak Dapat Basis</th>
					<th style="color:#fff;">Potongan Buah Mentah Diperam</th>
					<th style="color:#fff;">Potongan Over Pruning</th>
					<th style="color:#fff;">Potongan Brondolan Tidak Di Alas</th>
					<th style="color:#fff;">Potongan Brondolan Kotor/Sampah</th>
					<th style="color:#fff;">Potongan Buah Dibelah</th>
					<th style="color:#fff;">Potongan Buah Tidak Disusun Rapi Di Tph</th>
				</tr>
			</thead>
			<tbody>
<%
'response.write has_close_payroll & strEmployee
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic

Set rsMandor = Server.CreateObject("ADODB.Recordset")
rsMandor.Open strMandor, connstr, adOpenStatic

Dim nPREMILEBIHBORONG
Dim nPREMITANDANHARILIBUR
Dim nPREMISIAPBORONG
Dim nPREMIMANDORPANEN
Dim nPREMIKERANIPANEN
Dim nPREMIMANDOR1
Dim nPREMIKERANIDIVISI
Dim nPOTONGANBUAHMENTAH
Dim nPOTONGANTANGKAIPANJANG
Dim nPOTONGANBUAHTIDAKDIPANEN
Dim nPOTONGANPELEPAHTIDAKDIPOTONG
Dim nPOTONGANPELEPAHTIDAKDISUSUN
Dim nPOTONGANBRONDOLANDIPIRINGAN
Dim nPOTONGANBUAHTINGGALDIPIRINGAN
Dim nPOTONGANBUAHMATAHARI
Dim nPOTONGANTIDAKDAPATBASIS
Dim nTOTAL

if rsMandor.BOF and rsMandor.EOF then
%>
	<tr bgcolor="#F4F4A0"><td colspan="27"><strong>MANDOR</strong></td></tr>
	<tr><td colspan="27">NO DATA AVAILABLE</td></tr>
<%
else
NOMOR = 1
%>
<tr bgcolor="#F4F4A0"><td colspan="27"><strong>MANDOR</strong></td></tr>
<%
while not rsMandor.eof
	if no  mod 2 = 1 then
		cClass = "odd"
	else
		cClass ="default"
	end if
%>
	<tr class="<%=cClass%>">
		<!--<td align="center"><%=NOMOR%></td>-->
		<td align="center"><%=rsMandor("EMPCODE")%></td>
		<td align="left"><%=rsMandor("EMPNAME")%></td>
		<td align="center"><%=rsMandor("sectionname")%></td>
		<td align="center"><%=rsMandor("GANGCODE")%></td>
		<td align="right"><%=FormatNumber(rsMandor("PREMIMANDOR1"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("PREMIMANDORPANEN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("PREMIKERANIPANEN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("PREMIKERANIDIVISI"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("PREMISIAPBORONG"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("PREMILEBIHBORONG"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("PREMITANDANHARILIBUR"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANBUAHMENTAH"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANTANGKAIPANJANG"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANBUAHTIDAKDIPANEN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANPELEPAHTIDAKDIPOTONG"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANPELEPAHTIDAKDISUSUN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANBRONDOLANDIPIRINGAN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANBUAHTINGGALDIPIRINGAN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANBUAHMATAHARI"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("POTONGANTIDAKDAPATBASIS"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("D27"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("D28"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("D29"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("D30"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("D31"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("D32"), 0)%></td>
		<td align="right"><%=FormatNumber(rsMandor("TOTAL"), 0)%></td>
<%
	NOMOR	=	NOMOR	+	1
	rsMandor.MoveNext
Wend
end if
rsMandor.Close
set rsMandor = Nothing

if rsEmployee.BOF and rsEmployee.EOF then
%>
	<tr bgcolor="#F4F4A0"><td colspan="27"><strong>ANGGOTA</strong></td></tr>
	<tr><td colspan="27">NO DATA AVAILABLE</td></tr>
<%
else
%>
<tr bgcolor="#F4F4A0"><td colspan="27"><strong>ANGGOTA</strong></td></tr>
<%
NO = 1

nPREMILEBIHBORONG = 0
nPREMITANDANHARILIBUR = 0
nPREMISIAPBORONG = 0
nPREMIMANDORPANEN = 0
nPREMIKERANIPANEN = 0
nPREMIMANDOR1 = 0
nPREMIKERANIDIVISI = 0
nPOTONGANBUAHMENTAH = 0
nPOTONGANTANGKAIPANJANG = 0
nPOTONGANBUAHTIDAKDIPANEN = 0
nPOTONGANPELEPAHTIDAKDIPOTONG = 0
nPOTONGANPELEPAHTIDAKDISUSUN = 0
nPOTONGANBRONDOLANDIPIRINGAN = 0
nPOTONGANBUAHTINGGALDIPIRINGAN = 0
nPOTONGANBUAHMATAHARI = 0
nPOTONGANTIDAKDAPATBASIS = 0
gt_d27 = 0
gt_d28 = 0
gt_d29 = 0
gt_d30 = 0
gt_d31 = 0
gt_d32 = 0

nTOTAL = 0


while not rsEmployee.eof
	if no  mod 2 = 1 then
		cClass = "odd"
	else
		cClass ="default"
	end if
%>
	<tr class="<%=cClass%>">
		<!--<td align="center"><%=NO%></td>-->
		<td align="center"><%=rsEmployee("EMPCODE")%></td>
		<td align="left"><%=rsEmployee("EMPNAME")%></td>
		<td align="center"><%=rsEmployee("sectionname")%></td>
		<td align="center"><%=rsEmployee("GANGCODE")%></td>
		<td align="right"><%=FormatNumber(rsEmployee("PREMIMANDOR1"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("PREMIMANDORPANEN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("PREMIKERANIPANEN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("PREMIKERANIDIVISI"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("PREMISIAPBORONG"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("PREMILEBIHBORONG"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("PREMITANDANHARILIBUR"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANBUAHMENTAH"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANTANGKAIPANJANG"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANBUAHTIDAKDIPANEN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANPELEPAHTIDAKDIPOTONG"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANPELEPAHTIDAKDISUSUN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANBRONDOLANDIPIRINGAN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANBUAHTINGGALDIPIRINGAN"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANBUAHMATAHARI"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("POTONGANTIDAKDAPATBASIS"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("D27"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("D28"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("D29"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("D30"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("D31"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("D32"), 0)%></td>
		<td align="right"><%=FormatNumber(rsEmployee("TOTAL"), 0)%></td>
<%
	NO = NO + 1

	nPREMILEBIHBORONG = nPREMILEBIHBORONG + Cdbl(rsEmployee("PREMILEBIHBORONG"))
	nPREMITANDANHARILIBUR = nPREMITANDANHARILIBUR + Cdbl(rsEmployee("PREMITANDANHARILIBUR"))
	nPREMISIAPBORONG = nPREMISIAPBORONG + Cdbl(rsEmployee("PREMISIAPBORONG"))
	nPREMIMANDORPANEN = nPREMIMANDORPANEN + Cdbl(rsEmployee("PREMIMANDORPANEN"))
	nPREMIKERANIPANEN = nPREMIKERANIPANEN + Cdbl(rsEmployee("PREMIKERANIPANEN"))
	nPREMIMANDOR1 = nPREMIMANDOR1 + Cdbl(rsEmployee("PREMIMANDOR1"))
	nPREMIKERANIDIVISI = nPREMIKERANIDIVISI + Cdbl(rsEmployee("PREMIKERANIDIVISI"))
	nPOTONGANBUAHMENTAH = nPOTONGANBUAHMENTAH + Cdbl(rsEmployee("POTONGANBUAHMENTAH"))
	nPOTONGANTANGKAIPANJANG = nPOTONGANTANGKAIPANJANG + Cdbl(rsEmployee("POTONGANTANGKAIPANJANG"))
	nPOTONGANBUAHTIDAKDIPANEN = nPOTONGANBUAHTIDAKDIPANEN + Cdbl(rsEmployee("POTONGANBUAHTIDAKDIPANEN"))
	nPOTONGANPELEPAHTIDAKDIPOTONG = nPOTONGANPELEPAHTIDAKDIPOTONG + Cdbl(rsEmployee("POTONGANPELEPAHTIDAKDIPOTONG"))
	nPOTONGANPELEPAHTIDAKDISUSUN = nPOTONGANPELEPAHTIDAKDISUSUN + Cdbl(rsEmployee("POTONGANPELEPAHTIDAKDISUSUN"))
	nPOTONGANBRONDOLANDIPIRINGAN = nPOTONGANBRONDOLANDIPIRINGAN + Cdbl(rsEmployee("POTONGANBRONDOLANDIPIRINGAN"))
	nPOTONGANBUAHTINGGALDIPIRINGAN = nPOTONGANBUAHTINGGALDIPIRINGAN + Cdbl(rsEmployee("POTONGANBUAHTINGGALDIPIRINGAN"))
	nPOTONGANBUAHMATAHARI = nPOTONGANBUAHMATAHARI + Cdbl(rsEmployee("POTONGANBUAHMATAHARI"))
	nPOTONGANTIDAKDAPATBASIS = nPOTONGANTIDAKDAPATBASIS + Cdbl(rsEmployee("POTONGANTIDAKDAPATBASIS"))
	gt_d27 = gt_d27 + Cdbl(rsEmployee("d27"))
	gt_d28 = gt_d28 + Cdbl(rsEmployee("d28"))
	gt_d29 = gt_d29 + Cdbl(rsEmployee("d29"))
	gt_d30 = gt_d30 + Cdbl(rsEmployee("d30"))
	gt_d31 = gt_d31 + Cdbl(rsEmployee("d31"))
	gt_d32 = gt_d32 + Cdbl(rsEmployee("d32"))
	nTOTAL =  nTOTAL + Cdbl(rsEmployee("TOTAL"))

	rsEmployee.MoveNext
Wend
end if
rsEmployee.Close
set rsEmployee = Nothing
%>

	<tr bgcolor="#F4F4A0">
		<td colspan="4" align="center"><strong>TOTAL</strong></td>
		<td align="right"><strong><%=FormatNumber(nPREMIMANDOR1, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPREMIMANDORPANEN, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPREMIKERANIPANEN, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPREMIKERANIDIVISI, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPREMISIAPBORONG, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPREMILEBIHBORONG, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPREMITANDANHARILIBUR, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANBUAHMENTAH, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANTANGKAIPANJANG, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANBUAHTIDAKDIPANEN, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANPELEPAHTIDAKDIPOTONG, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANPELEPAHTIDAKDISUSUN, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANBRONDOLANDIPIRINGAN, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANBUAHTINGGALDIPIRINGAN, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANBUAHMATAHARI, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nPOTONGANTIDAKDAPATBASIS, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(gt_d27, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(gt_d28, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(gt_d29, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(gt_d30, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(gt_d31, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(gt_d32, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(nTOTAL, 0)%></strong></td>
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
