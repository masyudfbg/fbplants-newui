<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rscurrentdate, strcurrentdate, varDefaultYear, varDefaultMonth
Dim ScriptName
Dim rsrincian_koperasi, strrincian_koperasi
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName
Dim NO, cClass, nTotalPekerja, nTotalPasangan, nTotalTanggungan, is_payroll_close

v_txtFcba = Request("txtFcba")
v_txtYear =	Request("txtYear")
v_txtMonth = Request("txtMonth")
v_txtGangCode = Request("txtGangCode")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

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
	varDefaultYear	=	Year(rscurrentdate("currentdate"))
	varDefaultMonth	=	Month(rscurrentdate("currentdate"))
Else
	varDefaultYear	=	Year(Session("v_Employee"))
	varDefaultMonth	=	Month(Session("v_Employee"))
End If
rscurrentdate.Close
set rscurrentdate = Nothing

If (v_txtYear = "" Or Len(Trim(v_txtYear)) = 0) Then v_txtYear = varDefaultYear
If (v_txtMonth = "" Or Len(Trim(v_txtMonth)) = 0) Then v_txtMonth = varDefaultMonth

is_payroll_close = "FALSE"
strcurrentdate = " SELECT NVL(payroll_executed, 'FALSE') payroll_executed" & _
                 " FROM empalldedstatus" & _                 
                 " WHERE year = '" & v_txtYear & "' AND month = '" & v_txtMonth & "' AND fcba = '" & v_txtFcba & "'"

Set rscurrentdate = Server.CreateObject("ADODB.Recordset")
rscurrentdate.Open strcurrentdate, connstr, adOpenStatic

If Not rscurrentdate.EOF Then
	is_payroll_close = rscurrentdate("payroll_executed")
Else
	is_payroll_close = "FALSE"
End If
rscurrentdate.Close
set rscurrentdate = Nothing
	
strrincian_koperasi = " SELECT fcba"
if is_payroll_close = "TRUE" then
	strrincian_koperasi = strrincian_koperasi & "        , employee_hist_api.get_sectionname('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode) sectionname" & _
											    "        , empcode employeecode" & _
											    "        , employee_hist_api.get_fcname('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode) employeename" & _
											    "        , employee_hist_api.get_presentgrade('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode) presentgrade"
else
	strrincian_koperasi = strrincian_koperasi & "        , employee_api.get_sectionname(fcba, empcode) sectionname" & _
											    "        , empcode employeecode" & _
											    "        , employee_api.get_employeename(fcba, empcode) employeename" & _
											    "        , employee_api.get_presentgrade(fcba, empcode) presentgrade"
end if
	strrincian_koperasi = strrincian_koperasi & "        , sum(decode(allowdedcode, 'D25', amount, 0)) * -1 iuran_koperasi" & _
											    "        , sum(decode(allowdedcode, 'D26', amount, 0)) * -1 bon_koperasi" & _
											    "        , sum(decode(allowdedcode, 'D25', amount, 'D26', amount, 0)) * -1 total_koperasi" & _
											    " from empallded" & _
											    " where fcba = '" & v_txtFcba & "'" & _
											    "        and allowdedcode in ('D25', 'D26')"
					
if  LEN(TRIM(v_txtYear)) > 0 then
	strrincian_koperasi	=	strrincian_koperasi & "        and to_number(to_char(adhocdate, 'yyyy')) = '" & v_txtYear & "'"
end if
if  LEN(TRIM(v_txtMonth)) > 0 then
	strrincian_koperasi	=	strrincian_koperasi & "        and to_number(to_char(adhocdate, 'mm')) = '" & v_txtMonth & "'"
end if
if  LEN(TRIM(v_txtGangCode)) > 0 then
	if is_payroll_close = "TRUE" then
		strrincian_koperasi = strrincian_koperasi & "       AND employee_hist_api.get_gangcode('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode) = '" & v_txtGangCode & "'"
	else
		strrincian_koperasi = strrincian_koperasi & "       AND employee_api.get_gangcode(fcba, empcode) = '" & v_txtGangCode & "'"
	end if
end if

if is_payroll_close = "TRUE" then
	strrincian_koperasi = strrincian_koperasi & " group by fcba" & _
	                                          "        , employee_hist_api.get_sectionname('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode)" & _
											  "        , empcode" & _
											  "        , employee_hist_api.get_fcname('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode)" & _
											  "        , employee_hist_api.get_presentgrade('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode)"
	strrincian_koperasi = strrincian_koperasi & " ORDER BY employee_hist_api.get_sectionname('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode)" & _
	                                          "          , employee_hist_api.get_presentgrade('" & v_txtYear & "','" & v_txtMonth & "', fcba, empcode)" & _
	                                          "          , empcode"
else
	strrincian_koperasi = strrincian_koperasi & " group by fcba" & _
			    		                      "          , employee_api.get_sectionname(fcba, empcode)" & _
			    		                      "          , empcode " & _
			    		                      "          , employee_api.get_employeename(fcba, empcode)" & _
			    		                      "          , employee_api.get_presentgrade(fcba, empcode)"
	strrincian_koperasi = strrincian_koperasi & " ORDER BY employee_api.get_sectionname(fcba, empcode), employee_api.get_presentgrade(fcba, empcode), empcode"
end if

	'response.write strrincian_koperasi
	
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

	
	

Set rsrincian_koperasi = Server.CreateObject("ADODB.Recordset")
rsrincian_koperasi.Open strrincian_koperasi, connstr, adOpenStatic 
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Rincian Potongan Koperasi</title>
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
		<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
		<table id="form-table" class="centered-margin">
			<tr>
				<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM
						mon = varDefaultMonth
						If len(v_txtMonth)=0 then v_txtMonth=mon

						Yr = varDefaultYear
						MinYr = 2015
						MaxYr = Yr  + 2
						If len(v_txtYear)= 0 then v_txtYear = Yr
					%>
				<td width="4">
					<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
					</select>
				</td>
				<td><font face="verdana" size="1">BULAN</font></td>
				<td width="2">
					<select name="txtMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
					</select>
				</td>
				<td><font face="verdana" size="1">GANGCODE</font></td>
				<td><select name="txtGangCode" ><option value=""></option>
				<%
				Dim rsGangCode, strGangCode
				strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE ASC"
				
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
		</table>
		</div>
	</form>

</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="7" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
		<h4>PT. <%=varCompanyName%></h4>
		<h5><%=varFcbaName%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>RINCIAN POTONGAN KOPERASI</h2>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
 </th>
  </tr>
	<tr>
		<th>Seksi</th>
		<th>No Pekerja</th>
		<th>Nama Pekerja</th>		
		<th width="4%">Status</th>
		<th>Iuran Anggota Koperasi</th>
		<th>Bon Koperasi</th>
		<th>Total Potongan Koperasi</th>
	</tr>
</thead>
<tbody>
<%if rsrincian_koperasi.BOF and rsrincian_koperasi.EOF then%>
	<tr>
		<td colspan="7" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
        Dim grandtotal_iurankoperasi : grandtotal_iurankoperasi = 0
		Dim grandtotal_totalkoperasi : grandtotal_totalkoperasi = 0
		Dim grandtotal_bonkoperasi : grandtotal_bonkoperasi = 0
		
		while not rsrincian_koperasi.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<!--<td align="center" height="30" width="3%"><%=NO%></td>-->
				<td align="left"><%=rsrincian_koperasi("SECTIONNAME")%></td>
				<td align="center" width="4%"><%=rsrincian_koperasi("EMPLOYEECODE")%></td>
				<td align="left"><%=rsrincian_koperasi("EMPLOYEENAME")%></td>
				<td align="center"><%=rsrincian_koperasi("PRESENTGRADE")%></td>
				<td align="right" width="4%"><%=FormatNumber(rsrincian_koperasi("iuran_koperasi"), 0)%></td>
				<td align="right" width="4%"><%=FormatNumber(rsrincian_koperasi("bon_koperasi"), 0)%></td>
				<td align="right" width="4%"><%=FormatNumber(rsrincian_koperasi("total_koperasi"), 0)%></td>
			</tr>
<%
			NO = NO + 1
			
			grandtotal_iurankoperasi = grandtotal_iurankoperasi + Cdbl(rsrincian_koperasi("iuran_koperasi"))
			grandtotal_bonkoperasi = grandtotal_bonkoperasi + Cdbl(rsrincian_koperasi("bon_koperasi"))
			grandtotal_totalkoperasi = grandtotal_totalkoperasi + Cdbl(rsrincian_koperasi("total_koperasi"))
			
			rsrincian_koperasi.MoveNext
		Wend
	rsrincian_koperasi.Close
	set rsrincian_koperasi = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="4"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FormatNumber(grandtotal_iurankoperasi, 0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(grandtotal_bonkoperasi, 0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(grandtotal_totalkoperasi, 0)%></strong></td>
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=varManagerName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=varKTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>