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
Dim rsJamsostek, strJamsostek
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName, varManagerName, varKTUName
Dim NO, cClass, nTotalPekerja, nTotalPasangan, nTotalTanggungan, filter

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

	varDefaultYear	=	Year(Date)
	varDefaultMonth	=	Month(Date)


rscurrentdate.Close
set rscurrentdate = Nothing

If (v_txtYear = "" Or Len(Trim(v_txtYear)) = 0) Then v_txtYear = varDefaultYear
If (v_txtMonth = "" Or Len(Trim(v_txtMonth)) = 0) Then v_txtMonth = varDefaultMonth

if trim(v_txtGangCode) <> "" then 
		filter = " AND e.gangcode='" & v_txtGangCode & "' "
end if

strJamsostek	=	" SELECT " & _
				" e.sectionname AS sectionname, " & _
				" e.fccode AS employeecode, " & _
				" e.fcname AS employeename, " & _
				" e.presentgrade, " & _
				" e.jamsostekno, " & _
				" e.jamsostekcategory, " & _
				" e.presentrate, " & _
				" ROUND( " & _
				"	CASE  " & _
				"		WHEN e.jamsostekcategory IN ('J1','J2','J3','J4','J5') THEN  " & _
				"			(CASE " & _
				"				WHEN e.presentgrade = 'KHL' THEN NVL(e.presentrate,0) * 25 " & _
				"				ELSE NVL(e.presentrate,0) " & _
				"			END) " & _
				"			* (NVL(p.rate_1,0) + NVL(p.rate_2,0) + NVL(p.rate_3,0) + NVL(p.rate_4,0) + NVL(p.rate_5,0)) / 100 " & _
				"		ELSE 0 " & _
				"	END " & _
				" ) AS tunjangan_jamsostek, " & _
				" ROUND( " & _
				"	CASE  " & _
				"		WHEN e.jamsostekcategory IN ('J1','J2') THEN  " & _
				"			(CASE " & _
				"				WHEN e.presentgrade = 'KHL' THEN e.presentrate * 25 " & _
				"				ELSE e.presentrate " & _
				"			END) " & _
				"			* (3/100) " & _
				"		WHEN e.jamsostekcategory IN ('J3') THEN " & _
				"			(CASE " & _
				"				WHEN e.presentgrade = 'KHL' THEN e.presentrate * 25 " & _
				"				ELSE e.presentrate " & _
				"			END) " & _
				"			* (2/100) " & _
				"		WHEN e.jamsostekcategory IN ('J4') THEN  " & _
				"			(CASE " & _
				"				WHEN e.presentgrade = 'KHL' THEN e.presentrate * 25 " & _
				"				ELSE e.presentrate " & _
				"			END) " & _
				"			* (1/100) " & _
				"		ELSE 0 " & _
				"	END " & _
				" ) AS beban_karyawan " & _
			" FROM " & _
			"	employee_hist_tab e " & _
			" INNER JOIN " & _
			"	parameter p ON e.jamsostekcategory = p.fccode AND e.fcba = p.fcba " & _
			" WHERE  " & _
			"	e.year= '" & v_txtYear & "' " & _
			"	AND e.month='" & v_txtMonth & "' " & _
			"	AND e.fcba = '" & v_txtFcba & "' " & _ 
			"   AND e.jamsostekcategory<>'NJ' " & _ 
			"   AND e.dateterminate is null" & _
			filter
											  
					

	
	'response.write strJamsostek
	
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

	
	

Set rsJamsostek = Server.CreateObject("ADODB.Recordset")
rsJamsostek.Open strJamsostek, connstr, adOpenStatic 
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - BPJS Ketenagakerjaan</title>
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
    <th colspan="8" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
		<h4>PT. <%=varCompanyName%></h4>
		<h5><%=varFcbaName%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>RINCIAN BPJS KETENAGAKERJAAN</h2>
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
		<th>No BPJS Ketenagakerjaan</th>
		<th>Tunjangan BPJS Ketenagakerjaan</th>
		<th>Potongan Karyawan</th>
		<th>Total BPJS Ketenagakerjaan yang harus dibayar</th>
	</tr>
</thead>
<tbody>
<%if rsJamsostek.BOF and rsJamsostek.EOF then%>
	<tr>
		<td colspan="8" align="left">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		Dim totalbpjsdibayar
        Dim grandtotaltunjangan : grandtotaltunjangan = 0
		Dim grandtotalkaryawan : grandtotalkaryawan = 0
		Dim grandtotalpotongan : grandtotalpotongan = 0
		
		while not rsJamsostek.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if

			totalbpjsdibayar = cdbl(rsJamsostek("tunjangan_jamsostek")) + cdbl(rsJamsostek("beban_karyawan"))
%>
			<tr class="<%=cClass%>">
				<!--<td align="center" height="30" width="3%"><%=NO%></td>-->
				<td align="left"><%=rsJamsostek("SECTIONNAME")%></td>
				<td align="center" width="4%"><%=rsJamsostek("EMPLOYEECODE")%></td>
				<td align="left"><%=rsJamsostek("EMPLOYEENAME")%></td>
				<td align="center"><%=rsJamsostek("PRESENTGRADE")%></td>
				<td align="left"><%=rsJamsostek("jamsostekno")%></td>
				<td align="right" width="4%"><%=FormatNumber(rsJamsostek("tunjangan_jamsostek"), 0)%></td>
				<td align="right" width="4%"><%=FormatNumber(rsJamsostek("beban_karyawan"), 0)%></td>
				<td align="right" width="4%"><%=FormatNumber(totalbpjsdibayar, 0)%></td>
				
			</tr>
<%
			NO = NO + 1
			
			grandtotaltunjangan = grandtotaltunjangan + Cdbl(rsJamsostek("tunjangan_jamsostek"))
			grandtotalkaryawan = grandtotalkaryawan + Cdbl(rsJamsostek("beban_karyawan"))
			grandtotalpotongan = grandtotalpotongan + Cdbl(totalbpjsdibayar)
			
			rsJamsostek.MoveNext
		Wend
	rsJamsostek.Close
	set rsJamsostek = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="5"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FormatNumber(grandtotaltunjangan, 0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(grandtotalkaryawan, 0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(grandtotalpotongan, 0)%></strong></td>
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