<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
v_txtFcba = Request("txtFcba")
txtvehiclegroup = Request("txtvehiclegroup")
txtvehiclecode = Request("txtvehiclecode")
txtcategorytransport = Request("txtcategorytransport")
txtDate = Request("txtDate")
If "" & txtDate = "" then txtDate = TextDate(date)

strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
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

If "" & v_txtFcba = "" then v_txtFcba = varDefaultBa

ScriptName = Request.ServerVariables("SCRIPT_NAME")

strrpt = 	"	SELECT fcba,vehiclecode,categorytransport," & _
			"	Sum(unit_hi) unit_hi,Sum(trip_hi) trip_hi,Sum(output_hi) output_hi," & _
			"	Sum(unit_shi) unit_shi,Sum(trip_shi) trip_shi,Sum(output_shi) output_shi," & _
			"	Sum(unit_sbi) unit_sbi,Sum(trip_sbi) trip_sbi,Sum(output_sbi) output_sbi" & _
			"	FROM " & _
			"	( " & _
			"	SELECT fcba,vehiclecode,vehicledate,categorytransport," & _
			"	Sum(unit) unit_hi,Sum(trip) trip_hi,Sum(Nvl(output,0)) output_hi," & _
			"	0 unit_shi,0 trip_shi,0 output_shi," & _
			"	0 unit_sbi,0 trip_sbi,0 output_sbi" & _
			"	FROM vehicleactivity WHERE fcba='" & v_txtFcba & "' AND " & _
			"	vehicledate=TO_DATE('" & txtDate & "', 'mm/dd/yyyy')" & _
			"	GROUP BY fcba,vehiclecode,vehicledate,categorytransport" & _
			"	UNION ALL " & _
			"	SELECT fcba,vehiclecode,vehicledate,categorytransport," & _
			"	0 unit_hi,0 trip_hi,0 output_hi," & _
			"	Sum(unit) unit_shi,Sum(trip) trip_shi,Sum(Nvl(output,0)) output_shi," & _
			"	0 unit_sbi,0 trip_sbi,0 output_sbi " & _
			"	FROM vehicleactivity WHERE fcba='" & v_txtFcba & "' AND " & _
			"	To_Char(vehicledate,'yyyymm') = To_Char(TO_DATE('" & txtDate & "', 'mm/dd/yyyy'),'yyyymm') AND To_Char(vehicledate,'dd') <= To_Char(TO_DATE('" & txtDate & "', 'mm/dd/yyyy'),'dd')" & _
			"	GROUP BY fcba,vehiclecode,vehicledate,categorytransport" & _
			"	UNION ALL " & _
			"	SELECT fcba,vehiclecode,vehicledate,categorytransport," & _
			"	0 unit_hi,0 trip_hi,0 output_hi," & _
			"	0 unit_shi,0 trip_shi,0 output_shi," & _
			"	Sum(unit) unit_sbi,Sum(trip) trip_sbi,Sum(Nvl(output,0)) output_sbi" & _
			"	FROM vehicleactivity WHERE fcba='" & v_txtFcba & "' AND " & _
			"	To_Char(vehicledate,'yyyymm') <= To_Char(TO_DATE('" & txtDate & "', 'mm/dd/yyyy'),'yyyymm')" & _
			"	GROUP BY fcba,vehiclecode,vehicledate,categorytransport" & _
			"	)" & _
			"	GROUP BY fcba,vehiclecode,categorytransport" 
			

strPusingan = " SELECT rpt.fcba" & _
              "        , vh.vehiclegroupcode" & _
              "        , rpt.vehiclecode" & _
              "        , vh.fcname vehiclename" & _
              "        , vac.fcname categotytransport" & _
              "        , rpt.unit_hi" & _
              "        , rpt.trip_hi" & _
              "        , rpt.output_hi" & _
              "        , CASE WHEN rpt.trip_hi = 0 THEN 0 ELSE ROUND(rpt.output_hi/rpt.trip_hi,0) END outputtrip_hi" & _
              "        , rpt.unit_shi" & _
              "        , rpt.trip_shi" & _
              "        , rpt.output_shi" & _
              "        , CASE WHEN rpt.trip_shi = 0 THEN 0 ELSE ROUND(rpt.output_shi/rpt.trip_shi,0) END outputtrip_shi" & _
              "        , rpt.unit_sbi" & _
              "        , rpt.trip_sbi" & _
              "        , rpt.output_sbi" & _
              "        , CASE WHEN rpt.trip_sbi = 0 THEN 0 ELSE ROUND(rpt.output_sbi/rpt.trip_sbi,0) END outputtrip_sbi" & _
              " FROM (" & strrpt & ") rpt" & _
              " LEFT JOIN vehicle vh ON (rpt.fcba = vh.fcba AND rpt.vehiclecode = vh.fccode)" & _
              " LEFT JOIN vehiclegroup vhg ON (vh.fcba = vhg.fcba AND vh.vehiclegroupcode = vhg.vehiclegroupcode)" & _
              " LEFT JOIN vehicleactivitycategory vac ON rpt.categorytransport = vac.fccode" & _
              " WHERE rpt.fcba = '" & v_txtFcba & "'" & _
              " /*AND rpt.vehicledate = TO_DATE('" & txtDate & "', 'mm/dd/yyyy') */"
			  
if "" & txtvehiclegroup <> "" then strPusingan	=	strPusingan	&	" AND vh.vehiclegroupcode = '" & txtvehiclegroup & "'"
if "" & txtvehiclecode <> "" then strPusingan	=	strPusingan	&	" AND rpt.vehiclecode = '" & txtvehiclecode & "'"
if "" & txtcategorytransport <> "" then strPusingan	=	strPusingan	&	" AND rpt.categorytransport = '" & txtcategorytransport & "'"

	strPusingan	=	strPusingan	&	" ORDER BY vh.vehiclegroupcode, rpt.vehiclecode, LPAD(vac.fccode,2,0)"

'response.write strPusingan
Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan, connstr, adOpenStatic
arr_data_eof = true
if not rsPusingan.eof then
	arr_data = rsPusingan.getrows()
	arr_data_eof = false
end if
rsPusingan.close : set rsPusingan = nothing
									
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

function f_row_vehiclecode(arr__, where__)
	arr1 = arr__
	result__ = 0
	for zz=-0 to ubound(arr1,2)
		if cstr(""&where__) = cstr(""&arr1(2,zz)) then
			result__ = result__ +1
		end if
	next
	
	f_row_vehiclecode = result__
end function

function f_row_vehiclegroupcode(arr__, where__)
	arr1 = arr__
	result__ = 0
	for zz=-0 to ubound(arr1,2)
		if cstr(""&where__) = cstr(""&arr1(1,zz)) then
			result__ = result__ +1
		end if
	next
	
	f_row_vehiclegroupcode = result__
end function

function f_addrow_vehiclegroupcode(arr__, where__)
	arr1 = arr__
	result__ = 0
	for zz=-0 to ubound(arr1,2)
		if cstr(""&where__) = cstr(""&arr1(1,zz)) then
			curr_vh = cstr(""&arr1(2,zz))
			if curr_vh <> prev_vh then
				result__ = result__ +1
			end if
			prev_vh = curr_vh
		end if
	next
	
	f_addrow_vehiclegroupcode = result__
end function
			
%>
<html>

<head>
<title>ePlaS Reporting - Harian Kerja Kendaraan</title>
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
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
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
								<font face="verdana" size="1">Group Kendaraan</font>
							</td>
							<td>
								<select name="txtvehiclegroup" >
									<option value="">--ALL--</option>
									<%
									strAfdeling	=	" SELECT vehiclegroupcode fccode, description fcname FROM vehiclegroup WHERE fcba = '" & v_txtFcba & "' ORDER BY vehiclegroupcode"
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = txtvehiclegroup then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE") & " : " & rsAfdeling("fcname")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE") & " : " & rsAfdeling("fcname")%></option>  
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
								<font face="verdana" size="1">Kode Kendaraan</font>
							</td>
							<td>
								<select name="txtvehiclecode" >
									<option value="">--ALL--</option>
									<%
									strAfdeling	=	" SELECT fccode, fcname FROM vehicle WHERE fcba = '" & v_txtFcba & "' ORDER BY fccode"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = txtvehiclecode then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE") & " : " & rsAfdeling("fcname")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE") & " : " & rsAfdeling("fcname")%></option>  
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
								<font face="verdana" size="1">Kategori Transport</font>
							</td>
							<td>
								<select name="txtcategorytransport" >
									<option value="">--ALL--</option>
									<%
									strAfdeling	=	" SELECT fccode, fcname FROM vehicleactivitycategory ORDER BY LPAD(fccode, 2, 0)"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = txtcategorytransport then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE") & " : " & rsAfdeling("fcname")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE") & " : " & rsAfdeling("fcname")%></option>  
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
								<font face="verdana" size="1">Tanggal</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDate" value="<%=txtDate%>" />
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
         <th colspan="16%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If "" & varCompanyName <> "" Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>Harian Kerja Kendaraan</h2>
				<br/>
				<h4>Tanggal : <%=Day(txtDate) & " " & Bulan(Month(txtDate)) & " " & Year(txtDate)%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Group Kendaraan</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Kode Kendaraan</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Nama Kendaraan</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Kategori Transport</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">Hari Ini</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">Sampai Hari Ini</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">Sampai Bulan Ini</th>
	</tr>
	<tr>
		<th>Jumlah HM/KM/Jam</th>
		<th>Jumlah Trip</th>
		<th>Output</th>
		<th>Output per Trip</th>
		
		<th>Jumlah HM/KM/Jam</th>
		<th>Jumlah Trip</th>
		<th>Output</th>
		<th>Output per Trip</th>
		
		<th>Jumlah HM/KM/Jam</th>
		<th>Jumlah Trip</th>
		<th>Output</th>
		<th>Output per Trip</th>
	</tr>
    </thead>
	<tbody>
<%
if not arr_data_eof then
	t_unit_hi = 0
	t_trip_hi = 0
	t_output_hi = 0
	
	t_unit_shi = 0
	t_trip_shi = 0
	t_output_shi = 0
	
	t_unit_sbi = 0
	t_trip_sbi = 0
	t_output_sbi = 0
	p_sp_vhg = 0
	p_sp_vh = 0
	for yy=0 to ubound(arr_data,2)
		p_sp_vhg = p_sp_vhg +1
		p_sp_vh = p_sp_vh +1
		sp_vhg = f_row_vehiclegroupcode(arr_data, arr_data(1,yy))
		addsp_vhg = f_addrow_vehiclegroupcode(arr_data, arr_data(1,yy))

		sp_vh = f_row_vehiclecode(arr_data, arr_data(2,yy))
%>
		<tr>
			<%if p_sp_vhg = 1 then%>
			<td align="center" valign="top" rowspan="<%=sp_vhg +addsp_vhg%>"><%=arr_data(1,yy)%></td>
			<%end if%>
			<%if p_sp_vh = 1 then%>
			<td align="left" valign="top" rowspan="<%=sp_vh +1%>"><%=arr_data(2,yy)%></td>
			<td align="left" valign="top" rowspan="<%=sp_vh +1%>"><%=arr_data(3,yy)%></td>
			<%end if%>
			<td align="left"><%=arr_data(4,yy)%></td>
			<td align="right"><%=FormatNumber(arr_data(5,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(6,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(7,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(8,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(9,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(10,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(11,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(12,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(13,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(14,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(15,yy),2)%></td>
			<td align="right"><%=FormatNumber(arr_data(16,yy),2)%></td>
		
<%
		t_unit_hi = t_unit_hi + cdbl(to_nol(arr_data(5,yy)))
		t_trip_hi = t_trip_hi + cdbl(to_nol(arr_data(6,yy)))
		t_output_hi = t_output_hi + cdbl(to_nol(arr_data(7,yy)))
		
		t_unit_shi = t_unit_shi + cdbl(to_nol(arr_data(9,yy)))
		t_trip_shi = t_trip_shi + cdbl(to_nol(arr_data(10,yy)))
		t_output_shi = t_output_shi + cdbl(to_nol(arr_data(11,yy)))
		
		t_unit_sbi = t_unit_sbi + cdbl(to_nol(arr_data(13,yy)))
		t_trip_sbi = t_trip_sbi + cdbl(to_nol(arr_data(14,yy)))
		t_output_sbi = t_output_sbi + cdbl(to_nol(arr_data(15,yy)))
		
		if p_sp_vh = sp_vh then
			response.write "</tr>"
			response.write "<tr style='background-color:#F4F4A0;'>"
			response.write "	<td><b>Total</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(t_unit_hi,2)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(t_trip_hi,2)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(t_output_hi,2)&"</b></td>"
			if t_trip_hi = 0 then
				response.write "	<td align='right'><b>0</b></td>"
			else
				response.write "	<td align='right'><b>"&FormatNumber(cdbl(t_output_hi)/cdbl(t_trip_hi),2)&"</b></td>"
			end if
			response.write "	<td align='right'><b>"&FormatNumber(t_unit_shi,2)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(t_trip_shi,2)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(t_output_shi,2)&"</b></td>"
			if t_trip_shi = 0 then
				response.write "	<td align='right'><b>0</b></td>"
			else
				response.write "	<td align='right'><b>"&FormatNumber(cdbl(t_output_shi)/cdbl(t_trip_shi),2)&"</b></td>"
			end if
			response.write "	<td align='right'><b>"&FormatNumber(t_unit_sbi,2)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(t_trip_sbi,2)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(t_output_sbi,2)&"</b></td>"
			if t_trip_sbi = 0 then
				response.write "	<td align='right'><b>0</b></td>"
			else
				response.write "	<td align='right'><b>"&FormatNumber(cdbl(t_output_sbi)/cdbl(t_trip_sbi),2)&"</b></td>"
			end if
			response.write "</tr>"
			p_sp_vh = 0
			
			t_unit_hi = 0
			t_trip_hi = 0
			t_output_hi = 0
			
			t_unit_shi = 0
			t_trip_shi = 0
			t_output_shi = 0
			
			t_unit_sbi = 0
			t_trip_sbi = 0
			t_output_sbi = 0
		else
			response.write "</tr>"
		end if
		
		if p_sp_vhg = sp_vhg then
			p_sp_vhg = 0
		end if
		
	next
else
%>
	<tr>
		<td colspan="16">No Data Available</td>
	</tr>
<%
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
