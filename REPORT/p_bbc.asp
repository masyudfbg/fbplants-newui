<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%

maxCol = 0

txtYear = Request("txtYear")
txtFcba = Request("txtFcba")
censuscode = Request("censuscode")
txtafdeling = Request("txtafdeling")
censusname = Request("censusname")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))

if ""&censuscode = "" then censuscode = "BBC1"
if ""&censusname = "" then censusname = "BBC 1"

d_span = 40

arr_fcba_eof = true
strDefaultBa = " SELECT fccode AS fcba" & _
				" FROM businessunit BU" & _
				" INNER JOIN user_fcba_tab uft ON (bu.fccode = uft.fcba) " & _
				" WHERE fctype = 'E'" & _
				"       AND fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	arr_fcba = rsDefaultBa.getrows()
	arr_fcba_eof = false
End If
rsDefaultBa.Close : set rsDefaultBa = Nothing

if not arr_fcba_eof then
	for xy=0 to ubound(arr_fcba,2)
		if xy=0 then varDefaultBa = arr_fcba(0,xy)
	next
end if

If "" & txtFcba = "" Then txtFcba = varDefaultBa

arr_afdeling_eof = true
strPusingan	= " SELECT * FROM ( SELECT DISTINCT field_api.getvalid_afdeling(fccode, fcba, TO_DATE('1/1/'||"&txtyear&", 'mm/dd/yyyy')) afdeling" & _
              " FROM field" & _
              " WHERE fcba = '" & txtFcba & "' ) "

strPusingan = strPusingan & " ORDER BY afdeling"
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_afdeling = qORS.getrows()
	arr_afdeling_eof = false
End If
qORS.Close
set qORS = Nothing

if not arr_afdeling_eof then
	for ii=0 to ubound(arr_afdeling,2)
		if ii=0  then defaultafdeling = arr_afdeling(0,ii)
	next
end if

if ""&txtafdeling = "" then txtafdeling = defaultafdeling

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_data_eof = true
strPusingan	= " SELECT * FROM ( SELECT bbc.periodyear" & _
              "        , bbc.fieldcode" & _
              "        , TO_CHAR(field_api.getvalid_plantingdate(bbc.fieldcode, bbc.fcba, TO_DATE('1/1/'||"&txtyear&", 'mm/dd/yyyy')), 'yyyy') tahuntanam" & _
              "        , "&txtyear&" - TO_NUMBER(TO_CHAR(field_api.getvalid_plantingdate(bbc.fieldcode, bbc.fcba, TO_DATE('1/1/'||"&txtyear&", 'mm/dd/yyyy')), 'yyyy')) rut" & _
              "        , field_api.getvalid_ownership(bbc.fieldcode, bbc.fcba, TO_DATE('1/1/'||"&txtyear&", 'mm/dd/yyyy')) ownership" & _
              "        , field_api.getvalid_hectarageplanted(bbc.fieldcode, bbc.fcba, TO_DATE('1/1/'||"&txtyear&", 'mm/dd/yyyy')) planted" & _
              "        , bbc.hectarageharvested panen" & _
              "        , field_api.getvalid_pokok(bbc.fieldcode, bbc.fcba, TO_DATE('1/1/'||"&txtyear&", 'mm/dd/yyyy')) total" & _
              "        , bbc.productivestand produktif" & _
              "        , bbc.samplestand actual" & _
              "        , bbc.samplebunch janjang" & _
              "        , bbc.harvestingbased_abw bjrsensus" & _
              "        , field_api.getvalid_afdeling(bbc.fieldcode, bbc.fcba, TO_DATE('1/1/'||"&txtyear&", 'mm/dd/yyyy')) division" & _
              " FROM bbctransaction bbc" & _
              " WHERE bbc.fcba = '" & txtFcba & "'" & _
              "       AND bbc.censuscode = '"&censuscode&"'" & _
			  " ) WHERE division = '"&txtafdeling&"'"

strPusingan = strPusingan & " "
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_data = qORS.getrows()
	arr_data_eof = false
End If
qORS.Close
set qORS = Nothing

arr_dist_eof = true
strPusingan	= " SELECT fcba" & _
              "        , censuscode" & _
              "        , distribution1" & _
              "        , distribution2" & _
              "        , distribution3" & _
              "        , distribution4" & _
              " FROM bbcdistribution" & _
              " WHERE periodyear = '"&txtyear&"'" & _
              "       AND censuscode = '" & censuscode & "'" & _
              "       AND fcba = '" & txtFcba & "'"
				

strPusingan = strPusingan & " "
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_dist = qORS.getrows()
	arr_dist_eof = false
End If
qORS.Close
set qORS = Nothing

function data_dist(arr__, col__, where__)
	result__ = 0
	arr1 = arr__
	if not arr_dist_eof then
		for xyy=0 to ubound(arr1,2)
			if cstr(""&where__) = cstr(""&arr1(1,xyy)) then
				result__ = result__ +cdbl(to_nol(arr1(col__,xyy)))
			end if
		next
	end if
	
	data_dist = result__
end function

val_dist1 = data_dist(arr_dist, 2, censuscode)
val_dist2 = data_dist(arr_dist, 3, censuscode)
val_dist3 = data_dist(arr_dist, 4, censuscode)
val_dist4 = data_dist(arr_dist, 5, censuscode)

if censuscode = "BBC1" then
	desc_dist1 = "Jan"
	desc_dist2 = "Feb"
	desc_dist3 = "Mar"
	desc_dist4 = "Apr"

elseif censuscode = "BBC2" then
	desc_dist1 = "Mei"
	desc_dist2 = "Jun"
	desc_dist3 = "Jul"
	desc_dist4 = "Agu"

elseif censuscode = "BBC3" then
	desc_dist1 = "Sep"
	desc_dist2 = "Okt"
	desc_dist3 = "Nov"
	desc_dist4 = "Des"
else
	desc_dist1 = "Jan"
	desc_dist2 = "Feb"
	desc_dist3 = "Mar"
	desc_dist4 = "Apr"
end if

arr_adjust_e_eof = true
strPusingan	= " SELECT fcba" & _
              "        , division" & _
              "        , plantingdate" & _
              "        , percent_bunch" & _
              "        , remarks_bunch" & _
			  "        , percent_abw" & _
              "        , remarks_abw" & _
              " FROM bbcadjustment" & _
              " WHERE periodyear = '"&txtyear&"'" & _
              "       AND censuscode = '" & censuscode & "'" & _
              "       AND ruleinfo = 'E'" & _
              "       AND fcba = '" & txtFcba & "'"

strPusingan = strPusingan & " "
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_adjust_e = qORS.getrows()
	arr_adjust_e_eof = false
End If
qORS.Close
set qORS = Nothing

function data_adjust_e_n(arr__, col__, afd__, tt__)
	result__ = 0
	arr1 = arr__
	if not arr_adjust_e_eof then
		for xyy=0 to ubound(arr1,2)
			if cstr(""&arr1(1,xyy)) <> "" and cstr(""&arr1(2,xyy)) <> "" then
			if cstr(""&afd__) = cstr(""&arr1(1,xyy)) and cstr(""&tt__) = cstr(""&arr1(2,xyy)) then
				result__ = result__ + cdbl(to_nol(arr1(col__,xyy)))
			end if
			end if
		next
	end if
	
	data_adjust_e_n = result__
end function

function data_adjust_e_v(arr__, col__, afd__, tt__)
	result__ = ""
	arr1 = arr__
	if not arr_adjust_e_eof then
		for xyy=0 to ubound(arr1,2)
			if cstr(""&arr1(1,xyy)) <> "" and cstr(""&arr1(2,xyy)) <> "" then
			if cstr(""&afd__) = cstr(""&arr1(1,xyy)) and cstr(""&tt__) = cstr(""&arr1(2,xyy)) then
				result__ = ""&arr1(col__,xyy)
			end if
			end if
		next
	end if
	
	data_adjust_e_v = result__
end function

arr_adjust_ro_eof = true
strPusingan	= " SELECT fcba" & _
              "        , plantingdate" & _
              "        , percent_bunch" & _
              "        , remarks_bunch" & _
			  "        , percent_abw" & _
              "        , remarks_abw" & _
              " FROM bbcadjustment" & _
              " WHERE periodyear = '"&txtyear&"'" & _
              "       AND censuscode = '" & censuscode & "'" & _
              "       AND ruleinfo = 'RO'" & _
              "       AND fcba = '" & txtFcba & "'"

strPusingan = strPusingan & " "
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_adjust_ro = qORS.getrows()
	arr_adjust_ro_eof = false
End If
qORS.Close
set qORS = Nothing

function data_adjust_ro_n(arr__, col__, tt__)
	result__ = 0
	arr1 = arr__
	if not arr_adjust_ro_eof then
		for xyy=0 to ubound(arr1,2)
			if cstr(""&tt__) = cstr(""&arr1(1,xyy)) then
				result__ = result__ + cdbl(to_nol(arr1(col__,xyy)))
			end if
		next
	end if
	
	data_adjust_ro_n = result__
end function

function data_adjust_ro_v(arr__, col__, tt__)
	result__ = ""
	arr1 = arr__
	if not arr_adjust_ro_eof then
		for xyy=0 to ubound(arr1,2)
			if cstr(""&tt__) = cstr(""&arr1(1,xyy)) then
				result__ = ""&arr1(col__,xyy)
			end if
		next
	end if
	
	data_adjust_ro_v = result__
end function


strBaInfo = " SELECT fcname AS fcbaname, fccompanycode, fccompanyname AS companyname" & _
			" FROM businessunit bu" & _
			" WHERE fccode = '" & txtFcba & "'"

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
rsBaInfo.Close : set rsBaInfo = Nothing

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
<title>ePlaS Reporting - BBC</title>
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
									if not rsFcba.eof then
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
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
								<font face="verdana" size="1">BBC</font>
							</td>
							<td>
								<select name="censuscode" >
									<%
									strFcba	=	" SELECT fccode, fcname FROM parameter" & _
												" WHERE parcode = 'BBC' AND fcba = '" & txtFcba & "'"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									if not rsFcba.eof then
									Do While Not rsFcba.eof
										if censuscode = rsFcba("FCCODE") then censusname = rsFcba("fcname")
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if censuscode = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
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
								<font face="verdana" size="1">DIVISION</font>
							</td>
							<td>
								<select name="txtafdeling" >
								<%
								if not arr_afdeling_eof then
									for ii=0 to ubound(arr_afdeling,2)
									%>
									<option value="<%=arr_afdeling(0,ii)%>" <%if cstr(""&txtafdeling) = cstr(""&arr_afdeling(0,ii)) then%> Selected <%End If%>>
										<%=arr_afdeling(0,ii)%>
									</option>
									<%
									next
								end if
								%>
								</select>
							</td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>LAPORAN BBC</h2>
				<h3>Periode <%=txtYear%></h3>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th colspan="10">Database</th>
		<th colspan="8">Hasil Sensus Murni</th>
		<th colspan="9">Adjust/ Penyesuaian (+/-) Estate</th>
		<th colspan="9">Adjust/ Penyesuaian (+/-) RO</th>
		<th colspan="4"><%=censusname%> Final</th>
	</tr>
	<tr>
		<th rowspan="3">Kode Blok</th>
		<th rowspan="3">Tahun Tanam</th>
		<th rowspan="3">RUT</th>
		<th rowspan="3">Inti /  Plasma</th>
		<th colspan="2" rowspan="2">Luas (Ha)</th>
		<th colspan="2">Jumlah Pokok</th>
		<th rowspan="3">SPH</th>
		<th rowspan="3">% Pkk Prod</th>
		<th colspan="3">Pokok Periksa</th>
		<th colspan="2" rowspan="2">Jumlah Janjang Diperiksa</th>
		<th rowspan="3">BJR Sensus (Kg)</th>
		<th colspan="2" rowspan="2">Total</th>
		<th colspan="2" rowspan="2">Jjg/Pkk</th>
		<th rowspan="3">Penjelasan Adjustment</th>
		<th colspan="2" rowspan="2">BJR</th>
		<th rowspan="3">Penjelasan Adjustment</th>
		<th rowspan="3">Jjg</th>
		<th rowspan="3">Kg</th>
		<th rowspan="3">Ton/Ha</th>
		<th colspan="2" rowspan="2">Jjg/Pkk</th>
		<th rowspan="3">Penjelasan Adjustment</th>
		<th colspan="2" rowspan="2">BJR</th>
		<th rowspan="3">Penjelasan Adjustment</th>
		<th rowspan="3">Jjg</th>
		<th rowspan="3">Kg</th>
		<th rowspan="3">Ton/Ha</th>
		<th colspan="4">Sebaran Per Bulan</th>
	</tr>
	<tr>
		<th rowspan="2">Total</th>
		<th rowspan="2">Produktif</th>
		<th rowspan="2">Stdt</th>
		<th rowspan="2">Actual</th>
		<th rowspan="2">Var</th>
		<th><%=desc_dist1%></th>
		<th><%=desc_dist2%></th>
		<th><%=desc_dist3%></th>
		<th><%=desc_dist4%></th>
	</tr>
	<tr>
		<th>Planted</th>
		<th>Panen</th>
		<th>Janjang</th>
		<th>Jjg/Pkk</th>
		<th>Jjg</th>
		<th>Kg</th>
		<th>%</th>
		<th>Jjg/Pkk</th>
		<th>%</th>
		<th>Kg/Jjg</th>
		<th>%</th>
		<th>Jjg/Pkk</th>
		<th>%</th>
		<th>Kg/Jjg</th>
		<th><%=val_dist1%></th>
		<th><%=val_dist2%></th>
		<th><%=val_dist3%></th>
		<th><%=val_dist4%></th>
	</tr>
    </thead>
	<tbody>
<%
	if not arr_data_eof then
		r_sph = 0
		r_percent_pkk_produktif = 0
		r_std = 0
		r_act = 0
		r_var = 0
		r_janjang = 0
		r_jjgpkk = 0
		r_bjr = 0
		r_total_jjg = 0
		r_total_kg = 0
		r_tonha = 0
		r_adj_percent_jjg_e = 0 : r_adj_jjgpkk_e = 0 : r_adj_percent_bjr_e = 0 : r_adj_kgjjg_e = 0 : r_adj_jjg_e = 0 : r_adj_kg_e = 0
		
		t_planted = 0 : t_panen = 0 : t_total = 0 : t_produktif = 0 : t_sph = 0 : t_percent_pkk_produktif = 0
		t_stdt = 0 : t_act = 0 : t_var = 0 : t_janjang = 0 : t_jjgpkk = 0 : t_total_jjg = 0 : t_total_kg = 0 : t_bjr = 0
		
		t_adj_percent_jjg_e = 0 : t_adj_jjgpkk_e = 0 : t_adj_percent_bjr_e = 0 : t_adj_kgjjg_e = 0 : t_adj_jjg_e = 0 : t_adj_kg_e = 0 : t_adj_tonha_e = 0
		t_dist1 = 0 : t_dist2 = 0 : t_dist3 = 0 : t_dist4 = 0
		
		for y=0 to ubound(arr_data,2)
			
			if cdbl(to_nol(arr_data(7,y))) = 0 then r_percent_pkk_produktif = 0 else r_percent_pkk_produktif = round(cdbl(to_nol(arr_data(8,y)))/cdbl(to_nol(arr_data(7,y)))*100,0)
			r_std = round(cdbl(to_nol(arr_data(8,y)))*0.19,0)
			r_act = cdbl(to_nol(arr_data(9,y)))
			if r_std = 0 then r_var = 0 else r_var = round( (r_act-r_std)/r_std*100, 0)
			r_janjang = cdbl(to_nol(arr_data(10,y)))
			if r_act = 0 then r_jjgpkk = 0 else r_jjgpkk = r_janjang/r_act
			r_bjr = cdbl(to_nol(arr_data(11,y)))
			r_total_jjg = round( cdbl(to_nol(arr_data(8,y)))*r_jjgpkk, 0)
			r_total_kg = round( r_total_jjg*r_bjr, 0)
			
			if cdbl(to_nol(arr_data(5,y))) = 0 then
				r_sph = 0
				r_tonha = 0
			else
				r_sph = round(cdbl(to_nol(arr_data(7,y)))/cdbl(to_nol(arr_data(5,y))),0)
				r_tonha = r_total_kg/cdbl(to_nol(arr_data(5,y)))/1000
			end if
			
			r_adj_percent_jjg_e = data_adjust_e_n(arr_adjust_e, 3, arr_data(12,y), arr_data(2,y))
			r_adj_jjgpkk_e = (100 +r_adj_percent_jjg_e)/100*r_jjgpkk
			r_adj_r_jjg_e = data_adjust_e_v(arr_adjust_e, 4, arr_data(12,y), arr_data(2,y))
			
			r_adj_percent_bjr_e = data_adjust_e_n(arr_adjust_e, 5, arr_data(12,y), arr_data(2,y))
			r_adj_kgjjg_e = (100 +r_adj_percent_bjr_e)/100*r_bjr
			r_adj_r_bjr_e = data_adjust_e_v(arr_adjust_e, 6, arr_data(12,y), arr_data(2,y))
			
			r_adj_jjg_e = round( cdbl(to_nol(arr_data(8,y)))*r_adj_jjgpkk_e, 0)
			r_adj_kg_e = round( r_adj_jjg_e*r_adj_kgjjg_e, 0)
			if cdbl(to_nol(arr_data(5,y))) = 0 then r_adj_tonha_e = 0 else r_adj_tonha_e = round( r_adj_kg_e/cdbl(to_nol(arr_data(5,y)))/1000, 1)
			
			
			r_adj_percent_jjg_ro = data_adjust_ro_n(arr_adjust_ro, 2, arr_data(2,y))
			r_adj_jjgpkk_ro = (100 +r_adj_percent_jjg_ro)/100*r_jjgpkk
			r_adj_r_jjg_ro = data_adjust_ro_v(arr_adjust_ro, 3, arr_data(2,y))
			
			r_adj_percent_bjr_ro = data_adjust_ro_n(arr_adjust_ro, 4, arr_data(2,y))
			r_adj_kgjjg_ro = (100 +r_adj_percent_bjr_ro)/100*r_bjr
			r_adj_r_bjr_ro = data_adjust_ro_v(arr_adjust_ro, 5, arr_data(2,y))
			
			r_adj_jjg_ro = round( cdbl(to_nol(arr_data(8,y)))*r_adj_jjgpkk_ro, 0)
			r_adj_kg_ro = round( r_adj_jjg_ro*r_adj_kgjjg_ro, 0)
			if cdbl(to_nol(arr_data(5,y))) = 0 then r_adj_tonha_ro = 0 else r_adj_tonha_ro = round( r_adj_kg_ro/cdbl(to_nol(arr_data(5,y)))/1000, 1)
			
			response.write "<tr>"
			response.write "	<td align='center'>"&arr_data(1,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(2,y)&"</td>"
			response.write "	<td align='right'>"&arr_data(3,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(4,y)&"</td>"
			response.write "	<td align='right'>"&arr_data(5,y)&"</td>"
			response.write "	<td align='right'>"&arr_data(6,y)&"</td>"
			response.write "	<td align='right'>"&arr_data(7,y)&"</td>"
			response.write "	<td align='right'>"&arr_data(8,y)&"</td>"
			response.write "	<td align='right'>"&r_sph&"</td>"
			response.write "	<td align='right'>"&r_percent_pkk_produktif&"</td>"
			response.write "	<td align='right'>"&r_std&"</td>"
			response.write "	<td align='right'>"&r_act&"</td>"
			response.write "	<td align='right'>"&r_var&"</td>"
			response.write "	<td align='right'>"&r_janjang&"</td>"
			response.write "	<td align='right'>"&round(r_jjgpkk,2)&"</td>"
			response.write "	<td align='right'>"&r_bjr&"</td>"
			response.write "	<td align='right'>"&r_total_jjg&"</td>"
			response.write "	<td align='right'>"&r_total_kg&"</td>"
			response.write "	<td align='right'>"&round(r_adj_percent_jjg_e,2)&"</td>"
			response.write "	<td align='right'>"&round(r_adj_jjgpkk_e,2)&"</td>"
			response.write "	<td align='left'>"&r_adj_r_jjg_e&"</td>"
			response.write "	<td align='right'>"&round(r_adj_percent_bjr_e,2)&"</td>"
			response.write "	<td align='right'>"&round(r_adj_kgjjg_e,2)&"</td>"
			response.write "	<td align='left'>"&r_adj_r_bjr_e&"</td>"
			response.write "	<td align='right'>"&r_adj_jjg_e&"</td>"
			response.write "	<td align='right'>"&r_adj_kg_e&"</td>"
			response.write "	<td align='right'>"&round(r_adj_tonha_e,1)&"</td>"
			
			response.write "	<td align='right'>"&round(r_adj_percent_jjg_ro,2)&"</td>"
			response.write "	<td align='right'>"&round(r_adj_jjgpkk_ro,2)&"</td>"
			response.write "	<td align='left'>"&r_adj_r_jjg_ro&"</td>"
			response.write "	<td align='right'>"&round(r_adj_percent_bjr_ro,2)&"</td>"
			response.write "	<td align='right'>"&round(r_adj_kgjjg_ro,2)&"</td>"
			response.write "	<td align='left'>"&r_adj_r_bjr_ro&"</td>"
			response.write "	<td align='right'>"&r_adj_jjg_ro&"</td>"
			response.write "	<td align='right'>"&r_adj_kg_ro&"</td>"
			response.write "	<td align='right'>"&round(r_adj_tonha_ro,1)&"</td>"
			
			response.write "	<td align='right'>"&r_adj_kg_e*val_dist1/100&"</td>"
			response.write "	<td align='right'>"&r_adj_kg_e*val_dist2/100&"</td>"
			response.write "	<td align='right'>"&r_adj_kg_e*val_dist3/100&"</td>"
			response.write "	<td align='right'>"&r_adj_kg_e*val_dist4/100&"</td>"
			response.write "</tr>"
			
			
			t_planted = t_planted +cdbl(to_nol(arr_data(5,y)))
			t_panen = t_panen +cdbl(to_nol(arr_data(6,y)))
			t_total = t_total +cdbl(to_nol(arr_data(7,y)))
			t_produktif = t_produktif +cdbl(to_nol(arr_data(8,y)))
			if t_planted = 0 then t_sph = 0 else t_sph = t_total/t_planted
			if t_total = 0 then t_percent_pkk_produktif = 0 else t_percent_pkk_produktif = round(t_produktif/t_total*100,0)
			t_stdt = t_stdt +r_std
			t_act = t_act +r_act
			if t_stdt = 0 then t_var = 0 else t_var = round( (t_act-t_stdt)/t_stdt*100, 0)
			t_janjang = t_janjang +r_janjang
			if t_act = 0 then t_jjgpkk = 0 else t_jjgpkk = t_janjang/t_act
			
			t_total_jjg = t_total_jjg +r_total_jjg
			t_total_kg = t_total_kg +r_total_kg
			if t_total_jjg = 0 then t_bjr = 0 else t_bjr = t_total_kg/t_total_jjg
			
			t_adj_jjg_e = t_adj_jjg_e + r_adj_jjg_e
			t_adj_kg_e = t_adj_kg_e + r_adj_kg_e
			
			
			if t_adj_jjg_e = 0 then t_adj_kgjjg_e = 0 else t_adj_kgjjg_e = round(t_adj_kg_e/t_adj_jjg_e,0)
			if t_bjr = 0 then t_adj_percent_bjr_e = 0 else t_adj_percent_bjr_e = (t_adj_kgjjg_e-t_bjr)/t_bjr*100
			
			if t_produktif = 0 then t_adj_jjgpkk_e = 0 else t_adj_jjgpkk_e = t_adj_jjg_e/t_produktif
			if t_jjgpkk = 0 then t_adj_percent_jjg_e = 0 else t_adj_percent_jjg_e = (t_adj_jjgpkk_e-t_jjgpkk)/t_jjgpkk*100
			if t_planted = 0 then t_adj_tonha_e = 0 else t_adj_tonha_e = t_adj_kg_e/t_planted/1000
			
			t_adj_jjg_ro = t_adj_jjg_ro + r_adj_jjg_ro
			t_adj_kg_ro = t_adj_kg_ro + r_adj_kg_ro
			
			
			if t_adj_jjg_ro = 0 then t_adj_kgjjg_ro = 0 else t_adj_kgjjg_ro = round(t_adj_kg_ro/t_adj_jjg_ro,0)
			if t_bjr = 0 then t_adj_percent_bjr_ro = 0 else t_adj_percent_bjr_ro = (t_adj_kgjjg_ro-t_bjr)/t_bjr*100
			
			if t_produktif = 0 then t_adj_jjgpkk_ro = 0 else t_adj_jjgpkk_ro = t_adj_jjg_ro/t_produktif
			if t_jjgpkk = 0 then t_adj_percent_jjg_ro = 0 else t_adj_percent_jjg_ro = (t_adj_jjgpkk_ro-t_jjgpkk)/t_jjgpkk*100
			if t_planted = 0 then t_adj_tonha_ro = 0 else t_adj_tonha_ro = t_adj_kg_ro/t_planted/1000
			
			t_dist1 = t_dist1 +r_adj_kg_e*val_dist1/100
			t_dist2 = t_dist2 +r_adj_kg_e*val_dist2/100
			t_dist3 = t_dist3 +r_adj_kg_e*val_dist3/100
			t_dist4 = t_dist4 +r_adj_kg_e*val_dist4/100
		next
		
		response.write "<tr>"
		response.write "	<td align='center' colspan='"&d_span&"'></td>"
		response.write "</tr>"
		response.write "<tr style='background-color:#F4F4A0;'>"
		response.write "	<td align='center' colspan='4'>Total "&txtafdeling&"</td>"
		response.write "	<td align='right'>"&t_planted&"</td>"
		response.write "	<td align='right'>"&t_panen&"</td>"
		response.write "	<td align='right'>"&t_total&"</td>"
		response.write "	<td align='right'>"&t_produktif&"</td>"
		response.write "	<td align='right'>"&round(t_sph,0)&"</td>"
		response.write "	<td align='right'>"&t_percent_pkk_produktif&"</td>"
		response.write "	<td align='right'>"&t_stdt&"</td>"
		response.write "	<td align='right'>"&t_act&"</td>"
		response.write "	<td align='right'>"&t_var&"</td>"
		response.write "	<td align='right'>"&t_janjang&"</td>"
		response.write "	<td align='right'>"&round(t_jjgpkk,2)&"</td>"
		response.write "	<td align='right'>"&round(t_bjr,2)&"</td>"
		response.write "	<td align='right'>"&t_total_jjg&"</td>"
		response.write "	<td align='right'>"&t_total_kg&"</td>"
		response.write "	<td align='right'>"&round(t_adj_percent_jjg_e,2)&"</td>"
		response.write "	<td align='right'>"&round(t_adj_jjgpkk_e,2)&"</td>"
		response.write "	<td align='right'></td>"
		response.write "	<td align='right'>"&round(t_adj_percent_bjr_e,4)&"</td>"
		response.write "	<td align='right'>"&t_adj_kgjjg_e&"</td>"
		response.write "	<td align='right'></td>"
		response.write "	<td align='right'>"&t_adj_jjg_e&"</td>"
		response.write "	<td align='right'>"&t_adj_kg_e&"</td>"
		response.write "	<td align='right'>"&round(t_adj_tonha_e,1)&"</td>"
		
		response.write "	<td align='right'>"&round(t_adj_percent_jjg_ro,2)&"</td>"
		response.write "	<td align='right'>"&round(t_adj_jjgpkk_ro,2)&"</td>"
		response.write "	<td align='right'></td>"
		response.write "	<td align='right'>"&round(t_adj_percent_bjr_ro,4)&"</td>"
		response.write "	<td align='right'>"&t_adj_kgjjg_ro&"</td>"
		response.write "	<td align='right'></td>"
		response.write "	<td align='right'>"&t_adj_jjg_ro&"</td>"
		response.write "	<td align='right'>"&t_adj_kg_ro&"</td>"
		response.write "	<td align='right'>"&round(t_adj_tonha_ro,1)&"</td>"
		
		response.write "	<td align='right'>"&t_dist1&"</td>"
		response.write "	<td align='right'>"&t_dist2&"</td>"
		response.write "	<td align='right'>"&t_dist3&"</td>"
		response.write "	<td align='right'>"&t_dist4&"</td>"
		response.write "</tr>"
	else
		response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
	end if
%>
</tbody>
</table>
	</body>
</html>