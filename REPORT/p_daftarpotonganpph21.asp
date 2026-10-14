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
Dim rspph21, strpph21
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName
Dim NO, cClass, nTotalPekerja, nTotalPasangan, nTotalTanggungan, nSubTotal, v_txtgrade

v_txtFcba = Request("txtFcba")
v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
v_txtGangCode = Request("txtGangCode")
v_txtgrade = Request("txtgrade")

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
	
' strpph21	=	" select fcba" & _
				' "        , month masa_pajak" & _
				' "        , year tahun_pajak" & _
				' "        , 0 pembetulan" & _
				' "        , employee_api.get_npwp(fcba, empcode) npwp" & _
				' "        , empcode kode" & _
				' "        , employee_api.get_employeename(fcba, empcode) nama" & _
				' "        , '21-100-01' kode_pajak" & _
				' "        , round(grossincome, 0) jumlah_bruto" & _
				' "        , round(tax_value, 0)* -1 jumlah_pph" & _
				' "        , employee_api.get_nationality(fcba, empcode) kode_negara" & _
				' " from TAXALLDEDLEDGER" & _
				' " where fcba = '" & v_txtFcba & "'" & _
				' "       and year = '" & v_txtYear & "'" & _
				' "       and month = '" & v_txtMonth & "'"
				
strpph21	=	" select fcba" & _
				"        , to_number(to_char(adhocdate, 'mm')) masa_pajak" & _
				"        , to_number(to_char(adhocdate, 'yyyy')) tahun_pajak" & _
				"        , 0 pembetulan" & _
				"        , employee_api.get_npwp(fcba, empcode) npwp" & _
				"        , empcode kode" & _
				"        , employee_api.get_employeename(fcba, empcode) nama" & _
				"        , employee_api.get_taxstatus(fcba, empcode) tanggungan" & _
				"        , '21-100-01' kode_pajak" & _
				"        , round(TAXALLDEDLEDGER_API.get_grossincome(to_number(to_char(adhocdate, 'yyyy')), to_number(to_char(adhocdate, 'mm')),empcode, fcba), 0) jumlah_bruto" & _
				"        , round(amount, 0)* -1 jumlah_pph" & _
				"        , employee_api.get_nationality(fcba, empcode) kode_negara" & _
				" from empallded" & _
				" where fcba = '" & v_txtFcba & "'" & _
				"       and allowdedcode = 'D03'" & _
				"       and to_number(to_char(adhocdate, 'yyyy')) = '" & v_txtYear & "'" & _
				"       and to_number(to_char(adhocdate, 'mm')) = '" & v_txtMonth & "'" & _
				"       and empcode in (select fccode from view_payrolldetail where base_salary > 0 and fcba = '" & v_txtFcba & "'" & _ 
				"       and TO_NUMBER (SUBSTR (yearmonth, 1, 4)) = " & v_txtYear & "" & _
				"       and TO_NUMBER (SUBSTR (yearmonth, 5, 2)) = " & v_txtMonth & "" & _
				"       )" & _
				" UNION " & _
				" SELECT a.fcba " & _
				"        , month masa_pajak " & _
				"        , year tahun_pajak " & _
				"        , 0 pembetulan " & _
				"        , employee_api.get_npwp(a.fcba, a.empcode) npwp " & _
				"        , a.empcode kode " & _ 
				"        , employee_api.get_employeename(a.fcba,a.empcode) nama " & _
				"        , employee_api.get_taxstatus(a.fcba, a.empcode) tanggungan " & _
				"        , '21-100-01' kode_pajak " & _ 
				"        , a.amount jumlah_bruto " & _
				"        , round(finalamount, 0)* -1 jumlah_pph " & _
				"        , 'INA' kode_negara " & _
				" FROM taxledger a " & _
                " INNER JOIN (SELECT y.fcba,To_Number(To_Char(bkmdate,'YYYY')) ayear,To_Number(To_Char(bkmdate,'MM')) amonth,employeecode,finalamount " & _
                "             FROM BKMPROSES x INNER JOIN bkmhdr y ON (x.fcba=y.fcba AND x.bkmno=y.bkmno)  " & _
                "             WHERE y.postingflg='1' AND allowdedcode='D03' AND document_classification='599' " & _
                "             AND x.fcba='" & v_txtFcba & "'" & _
                "             AND To_Number(To_Char(bkmdate,'YYYY'))=" & v_txtYear &  "" & _
                "             AND To_Number(To_Char(bkmdate,'MM'))=" & v_txtMonth &  "" & _
                "            )b ON (a.year=b.ayear AND a.month=b.amonth AND a.fcba=b.fcba AND a.empcode=b.employeecode) " & _
				" WHERE a.processid='gajibruto_tm' AND a.fcba = '" & v_txtFcba & "' AND year = " & v_txtYear & " AND month = " & v_txtMonth &  ""

				
if  LEN(TRIM(v_txtGangCode)) > 0 then
	strpph21 = strpph21 & " and employee_api.get_gangcode(fcba, empcode) = '" & v_txtGangCode & "'"
end if

if len("" & trim(v_txtgrade)) <> 0 then
	if v_txtgrade = "TETAP" then
		strpph21 = strpph21 & "       and employee_hist_api.get_presentgrade(" & v_txtYear & ", " & v_txtMonth & ", fcba, empcode) in ('KHT', 'KBT')"
	else
		strpph21 = strpph21 & "       and employee_hist_api.get_presentgrade(" & v_txtYear & ", " & v_txtMonth & ", fcba, empcode)  = '" & v_txtgrade & "'"
	end if
end if

	strpph21 = strpph21 & " order by kode"

'response.write strpph21	
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

Set rspph21 = Server.CreateObject("ADODB.Recordset")
rspph21.Open strpph21, connstr, adOpenStatic 
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - PPh 21</title>
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
						MaxYr = Yr  + 3
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
				<td><font face="verdana" size="1">STATUS</font></td>
				<td>
					<select name="txtgrade">
						<Option Value="" <%If v_txtgrade = "" then%> Selected <%End If%>>--ALL--</option>
						<Option Value="TETAP" <%If v_txtgrade = "TETAP" then%> Selected <%End If%>>KHT + KBT: Karyawan Harian/Bulanan Tetap</option>
						<Option Value="KBT" <%If v_txtgrade = "KBT" then%> Selected <%End If%>>KBT : Karyawan Bulanan Tetap</option>
						<Option Value="KHT" <%If v_txtgrade = "KHT" then%> Selected <%End If%>>KHT : Karyawan Harian Tetap</option>
						<Option Value="KHL" <%If v_txtgrade = "KHL" then%> Selected <%End If%>>KHL : Karyawan Harian Lepas</option>
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
    <th colspan="12" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
		<h4>PT. <%=varCompanyName%></h4>
		<h5><%=varFcbaName%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>DAFTAR POTONGAN PAJAK PENGHASILAN 21 KARYAWAN</h2>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
 </th>
  </tr>
	<tr>
	    <th>No</th>
		<th>Masa Pajak</th>
		<th>Tahun Pajak</th>
		<th>Pembetulan</th>		
		<th>NPWP</th>
		<th>NIK</th>
		<th>Nama</th>
		<th>Status</th>
		<th>Kode Pajak</th>
		<th>Jumlah Bruto</th>
		<th>Jumlah PPh</th>
		<th>Kode Negara</th>
	</tr>
</thead>
<tbody>
<%if rspph21.BOF and rspph21.EOF then%>
	<tr>
		<td colspan="12" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		
		Dim subtotal_jumlah_bruto : subtotal_jumlah_bruto = 0
		Dim subtotal_jumlah_pph: subtotal_jumlah_pph = 0

		while not rspph21.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center" height="30" width="3%"><%=NO%></td>
				<td align="center"><%=rspph21("masa_pajak")%></td>
				<td align="center"><%=rspph21("tahun_pajak")%></td>
				<td align="right"><%=rspph21("pembetulan")%></td>
				<td align="left"><%=rspph21("npwp")%></td>
				<td align="left"><%=rspph21("kode")%></td>
				<td align="left"><%=rspph21("nama")%></td>
				<td align="left"><%=rspph21("tanggungan")%></td>
				<td align="center"><%=rspph21("kode_pajak")%></td>
				<td align="right"><%=FormatNumber(rspph21("jumlah_bruto"), 0)%></td>
				<td align="right"><%=FormatNumber(rspph21("jumlah_pph"), 0)%></td>
				<td align="center"><%=rspph21("kode_negara")%></td>
			</tr>
<%
			NO = NO + 1
			
			subtotal_jumlah_bruto = subtotal_jumlah_bruto + Cdbl(rspph21("jumlah_bruto"))
			subtotal_jumlah_pph = subtotal_jumlah_pph + Cdbl(rspph21("jumlah_pph"))
			
			rspph21.MoveNext
		Wend
	rspph21.Close
	set rspph21 = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="9"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FormatNumber(subtotal_jumlah_bruto, 0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(subtotal_jumlah_pph, 0)%></strong></td>
				<td align="center"><strong>&nbsp;</strong></td>
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