<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay,v_txtDate, v_txtAfdeling, v_txtBlok, v_txtColor
Dim ScriptName
Dim rsPusingan, strPusingan, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName,varCompanyCode, varCompanyName
Dim oRS, oSQL
d_span = 15
maxCol = 0

txtYear = Request("txtYear")
txtMonth = Request("txtMonth")
txtstatus = Request("txtstatus")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
If "" & txtstatus = "" Then txtstatus = "TM"
	

v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")
v_txtFcba		=	Request("txtFcba")

Dim v_txtPersen : v_txtPersen = Request("txtPersen")
If Len(trim(v_txtPersen)) = 0 Then v_txtPersen = 0


strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")


arr_data_eof = true
strPusingan	=	" SELECT *" & _
				" FROM trans_bisa_agrotech_dtl" & _
				" WHERE fcba = '" & v_txtFcba & "'" & _
				"       AND accountingyear = '" & txtYear & "'" & _
				"       AND accountingperiod = '" & txtMonth & "'"
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_data = qORS.getrows()
	arr_data_eof = false
End If
qORS.Close
set qORS = Nothing

Function f_value(data_array__, fieldcode__, fccode__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if "" & fieldcode__ <> "" and "" & fccode__ <> "" then			
			if cstr(arr_1(2, indx)) = cstr(fieldcode__) and cstr(arr_1(3, indx)) = cstr(fccode__) then
				result_ = cdbl(to_nol(arr_1(4, indx)))
			end if
		end if
	next
	f_value = result_
End Function

fddate = txtMonth&"/1/"&txtYear

if "" & v_txtBlok <> "" then
	wblok = v_txtBlok
end if

arr_field_eof = true
strPusingan	= " SELECT hdr.division" & _
              "        , hdr.fieldcode" & _
              "        , hdr.plantingyear" & _
              "        , hdr.hectarageplanted" & _
              "        , MAX(DECODE(hdr.accountingperiod, 1, hdr.nilai_akhir, '')) jan" & _
              "        , MAX(DECODE(hdr.accountingperiod, 2, hdr.nilai_akhir, '')) feb" & _
              "        , MAX(DECODE(hdr.accountingperiod, 3, hdr.nilai_akhir, '')) mar" & _
              "        , MAX(DECODE(hdr.accountingperiod, 4, hdr.nilai_akhir, '')) apr" & _
              "        , MAX(DECODE(hdr.accountingperiod, 5, hdr.nilai_akhir, '')) mei" & _
              "        , MAX(DECODE(hdr.accountingperiod, 6, hdr.nilai_akhir, '')) jun" & _
              "        , MAX(DECODE(hdr.accountingperiod, 7, hdr.nilai_akhir, '')) jul" & _
              "        , MAX(DECODE(hdr.accountingperiod, 8, hdr.nilai_akhir, '')) agu" & _
              "        , MAX(DECODE(hdr.accountingperiod, 9, hdr.nilai_akhir, '')) sep" & _
              "        , MAX(DECODE(hdr.accountingperiod, 10, hdr.nilai_akhir, '')) okt" & _
              "        , MAX(DECODE(hdr.accountingperiod, 11, hdr.nilai_akhir, '')) nov" & _
              "        , MAX(DECODE(hdr.accountingperiod, 12, hdr.nilai_akhir, '')) des" & _
              " FROM trans_bisa_agrotech_hdr hdr" & _
			  " LEFT JOIN (" & _
			  "     SELECT DISTINCT dtl.fcba, dtl.accountingyear, dtl.accountingperiod, dtl.fieldcode" & _
			  "     FROM trans_bisa_agrotech_dtl dtl" & _
			  "     INNER JOIN setup_bisa_agrotech_tab stp ON dtl.fccode = stp.fccode" & _
			  "     WHERE dtl.fcba = '" & v_txtFcba & "'" & _
			  "           AND dtl.accountingyear = '" & txtYear & "'" & _
			  "           AND dtl.accountingperiod <= '" & txtMonth & "'"
if "" & txtstatus <> "" then
	strPusingan = strPusingan & "           AND stp.status = '" & txtstatus & "'"
end if

if "" & v_txtBlok <> "" then
	strPusingan = strPusingan & "           AND dtl.fieldcode = '" & v_txtBlok & "'"
end if
strPusingan = strPusingan & " ) dtl ON ( dtl.fcba = hdr.fcba" & _
                            "            AND dtl.accountingyear = hdr.accountingyear" & _
                            "            AND dtl.accountingperiod = hdr.accountingperiod" & _
                            "            AND dtl.fieldcode = hdr.fieldcode)" & _
                            " WHERE hdr.fcba = '" & v_txtFcba & "'" & _
                            "       AND hdr.accountingyear = '" & txtYear & "'" & _
                            "       AND hdr.accountingperiod <= '" & txtMonth & "'" & _
                            "       AND hdr.status = '" & txtstatus & "'"

if "" & v_txtAfdeling <> "" then
	strPusingan = strPusingan & "       AND hdr.division = '" & v_txtAfdeling & "'"
end if


strPusingan = strPusingan & " GROUP BY hdr.fieldcode, hdr.plantingyear, hdr.hectarageplanted, hdr.division"
strPusingan = strPusingan & " ORDER BY hdr.division, hdr.fieldcode"
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_field = qORS.getrows()
	arr_field_eof = false
End If
qORS.Close
set qORS = Nothing

strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

qSQL = " SELECT fccode, fcname, parent_fccode, priority, status" & _
       " FROM setup_bisa_agrotech_tab" & _
       " ORDER BY SUBSTR(fccode, 1, 2), fccode"
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open qSQL, connstr, adOpenStatic

arr_setup_eof = true
If Not qORS.EOF Then
	arr_setup = qORS.getrows()
	arr_setup_eof = false
end if
qORS.Close
set qORS = Nothing

Function f_colspan(data_array__, code__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if len("" & code__) = 2 and len("" & arr_1(2, indx)) = 2 then			
			if cstr(arr_1(2, indx)) = cstr(code__) then
				result_ = result_ +1
			end if
		end if
	next
	f_colspan = result_
End Function

Function f_maxrow_division(data_array__, code__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if cstr("" & arr_1(0, indx)) = cstr("" & code__) then
			result_ = result_ +1
		end if
	next
	f_maxrow_division = result_
End Function


%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>Rekap BISA Agrotechnical - iPlaS Reporting</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
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
-->

<style type="text/css">
@media print
{
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
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
					<table id="form-table" class="centered-margin">	
						<tr>
						<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM

						Yr = txtYear
						MinYr = 2015
						MaxYr = Yr  + 2
					%>
							<td width="4">
								<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
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
							<Option Value="<%=i%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
												" AND FCTYPE = 'E'" & _
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									
									If rsFcba.Bof Or rsFcba.eof Then
									%>
									<option value=""></option>
									<%
									End If
									
									if not rsFcba.eof then
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if v_txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
									<%
										rsFcba.movenext
									loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">STATUS</font>
							</td>
							<td>
								<select name="txtstatus" >
									<option value="TM" <%if txtstatus = "TM" then%> selected <%end if%>>TM</option>
									<option value="TBM" <%if txtstatus = "TBM" then%> selected <%end if%>>TBM</option>
								</select>
							<td>
								<font face="verdana" size="1">AFDELING</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="">--ALL--</option>
									<%
									Dim rsAfdeling, strAfdeling
									strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = v_txtAfdeling then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE")%></option>  
									<%End if
										rsAfdeling.movenext
									loop
									end if
									rsAfdeling.Close
									set rsAfdeling = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">BLOK</font>
							</td>
							<td>
								<select name="txtBlok" >
									<option value="">--ALL--</option>
									<%
									Dim rsBlok, strBlok
									strBlok	=	" SELECT fccode, fcname FROM field WHERE fcba = '" & v_txtFcba & "'" & _
												" AND field_api.getvalid_isactive_harvested(fccode, fcba, TO_DATE('"&v_txtDate&"', 'mm/dd/yyyy')) = 'TRUE' ORDER BY fccode"
									
									Set rsBlok = Server.CreateObject("ADODB.Recordset")
									rsBlok.Open strBlok, connstr, adOpenStatic 
									
									if not rsBlok.eof then
									Do While Not rsBlok.eof
									if rsBlok("FCCODE") = v_txtBlok then%>
									<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option> 
									<%Else%>
									<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option>  
									<%End if
										rsBlok.movenext
									loop
									end if
									rsBlok.Close
									set rsBlok = Nothing
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
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>BEST IMPROVEMENT SELF ASSESSMENT (BISA) <%=txtstatus%></h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th rowspan="2">Blok</th>
		<th rowspan="2">TT</th>
		<th rowspan="2">Luas (Ha)</th>
		<th colspan="12">Kualitas Total Kondisi Blok (Blok Standart)</th>
	</tr>
	<tr>
		<th>Jan</th>
		<th>Feb</th>
		<th>Mar</th>
		<th>Apr</th>
		<th>Mei</th>
		<th>Jun</th>
		<th>Jul</th>
		<th>Agu</th>
		<th>Sep</th>
		<th>Okt</th>
		<th>Nov</th>
		<th>Des</th>
	</tr>

    </thead>
	<tbody>
<%
	if not arr_field_eof then
		defST = 0
		st_ha = 0
		for k=4 to 15
			execute("st_excellent"&k&" = 0")
			execute("st_optimum"&k&" = 0")
			execute("st_minimum"&k&" = 0")
			execute("st_red"&k&" = 0")
			execute("st_totalblok"&k&" = 0")
			
			execute("gt_excellent"&k&" = 0")
			execute("gt_optimum"&k&" = 0")
			execute("gt_minimum"&k&" = 0")
			execute("gt_red"&k&" = 0")
			execute("gt_totalblok"&k&" = 0")
		next
		for y=0 to ubound(arr_field,2)
			defST = defST +1
			cDivision = arr_field(0,y)
			if cDivision <> pDivision then
				maxST = f_maxrow_division(arr_field, cDivision)
			end if
			
			st_ha = st_ha + cdbl(to_nol(arr_field(3,y)))
			
			response.write "<tr>"
			response.write "	<td align='left'>"&arr_field(1,y)&"</td>"
			response.write "	<td align='center'>"&arr_field(2,y)&"</td>"
			response.write "	<td align='right'>"&arr_field(3,y)&"</td>"
			for k=4 to 15
				execute("st_totalblok"&k&" = st_totalblok"&k&" +1")
				execute("gt_totalblok"&k&" = gt_totalblok"&k&" +1")
				if cstr("" & arr_field(k,y)) = cstr("A") then
					n_style = "background-color:#0074D9;color:#0074D9;"
					
					execute("st_excellent"&k&" = st_excellent"&k&" +1")					
					execute("gt_excellent"&k&" = gt_excellent"&k&" +1")
					
				elseif cstr("" & arr_field(k,y)) = cstr("B") then
					n_style = "background-color:#2ECC40;color:#2ECC40;"
					
					execute("st_optimum"&k&" = st_optimum"&k&" +1")					
					execute("gt_optimum"&k&" = gt_optimum"&k&" +1")

				elseif cstr("" & arr_field(k,y)) = cstr("C") then
					n_style = "background-color:#FFDC00;color:#FFDC00;"
					
					execute("st_minimum"&k&" = st_minimum"&k&" +1")
					execute("gt_minimum"&k&" = gt_minimum"&k&" +1")

				elseif cstr("" & arr_field(k,y)) = cstr("D") then
					n_style = "background-color:#FFDC00;color:#FFDC00;"
					
					execute("st_minimum"&k&" = st_minimum"&k&" +1")
					execute("gt_minimum"&k&" = gt_minimum"&k&" +1")

				elseif cstr("" & arr_field(k,y)) = cstr("E") then
					n_style = "background-color:#FF4136;color:#FF4136;"
					
					execute("st_red"&k&" = st_red"&k&" +1")
					execute("gt_red"&k&" = gt_red"&k&" +1")
				else
					n_style = "background-color:#FFFFFF;color:#111111;"
				end if
			response.write "	<td align='center' style='"&n_style&"'>"&arr_field(k,y)&"</td>"
			next
			response.write "</tr>"
			
			if defST = maxST then				
				response.write "<tr>"
				response.write "	<td align='center' rowspan='6' colspan='2'> Total "&cDivision&"</td>"
				response.write "    <td align='right'><b>"&st_ha&"</b></td>"
				response.write "    <td align='right' colspan='12'>&nbsp;</td>"
				response.write "</tr>"
				response.write "<tr>"
				response.write "	<td align='center' style='background-color:#0074D9;color:#FFFFFF;'>Excellent</td>"
				for k=4 to 15
				response.write "	<td align='right' > "&eval("st_excellent"&k)&"</td>"
				next
				response.write "</tr>"
				
				response.write "<tr>"
				response.write "	<td align='center' style='background-color:#2ECC40;'>Optimum</td>"
				for k=4 to 15
				response.write "	<td align='right' > "&eval("st_optimum"&k)&"</td>"
				next
				response.write "</tr>"
				
				response.write "<tr>"
				response.write "	<td align='center' style='background-color:#FFDC00;'>Minimum</td>"
				for k=4 to 15
				response.write "	<td align='right' > "&eval("st_minimum"&k)&"</td>"
				next
				response.write "</tr>"
				
				response.write "<tr>"
				response.write "	<td align='center' style='background-color:#FF4136;color:#FFFFFF;'>Red</td>"
				for k=4 to 15
				response.write "	<td align='right' > "&eval("st_red"&k)&"</td>"
				next
				response.write "</tr>"
				
				response.write "<tr>"
				response.write "	<td align='center' style='background-color:#DDDDDD;'>&nbsp;</td>"
				for k=4 to 15
				response.write "	<td align='right' style='background-color:#DDDDDD;font-weight:bold;'> "&eval("st_totalblok"&k)&"</td>"
				next
				response.write "</tr>"
				
				if y <> ubound(arr_field,2) then
					response.write "<tr><td colspan='15'></td></tr>"
				end if
				
				defST = 0
				st_ha = 0
				
				for k=4 to 15
					execute("st_excellent"&k&" = 0")
					execute("st_optimum"&k&" = 0")
					execute("st_minimum"&k&" = 0")
					execute("st_red"&k&" = 0")
					execute("st_totalblok"&k&" = 0")
				next
			end if
			
			pDivision = cDivision
		next
		
		response.write "<tr>"
		response.write "	<td align='center' colspan='15'></td></tr>"
		response.write "	<td align='center' colspan='15' style='background-color:#DDDDDD;font-weight:bold;'>REKAPITULASI KEBUN</td></tr>"
		
		response.write "<tr>"
		response.write "	<td align='center' rowspan='5' colspan='2'>Jumlah Blok</td>"
		response.write "	<td align='center' style='background-color:#0074D9;color:#FFFFFF;'>Excellent</td>"
		for k=4 to 15
		response.write "	<td align='right' > "&eval("gt_excellent"&k)&"</td>"
		next
		response.write "</tr>"
		
		response.write "<tr>"
		response.write "	<td align='center' style='background-color:#2ECC40;'>Optimum</td>"
		for k=4 to 15
		response.write "	<td align='right' > "&eval("gt_optimum"&k)&"</td>"
		next
		response.write "</tr>"
		
		response.write "<tr>"
		response.write "	<td align='center' style='background-color:#FFDC00;'>Minimum</td>"
		for k=4 to 15
		response.write "	<td align='right' > "&eval("gt_minimum"&k)&"</td>"
		next
		response.write "</tr>"
		
		response.write "<tr>"
		response.write "	<td align='center' style='background-color:#FF4136;color:#FFFFFF;'>Red</td>"
		for k=4 to 15
		response.write "	<td align='right' > "&eval("gt_red"&k)&"</td>"
		next
		response.write "</tr>"
		
		response.write "<tr>"
		response.write "	<td align='center' style='background-color:#DDDDDD;'>Total</td>"
		for k=4 to 15
		response.write "	<td align='right' style='background-color:#DDDDDD;font-weight:bold;'> "&eval("gt_totalblok"&k)&"</td>"
		next
		response.write "</tr>"
else
	response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
end if
%>
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
	   <%
		DIM rsttd,ttd
		
		ttd = "SELECT FCMANAGER,FCKTU FROM BusinessUnit WHERE FCCODE = '" & v_txtFcba & "'"
		
	    Set rsttd = Server.CreateObject("ADODB.Recordset")
		rsttd.Open ttd, connstr, adOpenStatic

		If Not rsttd.EOF Then
	   %>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCKTU")%></u></strong></td>
	   </tr>
	<%  End If
%> 
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
