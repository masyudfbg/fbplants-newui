<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtTanggal, v_txtVehicleGroup
Dim ScriptName
Dim rsOKPT, strOKPT, strOKPT2, strMain
Dim NO, cClass, nday, nSpanTotal

nSpanTotal = 6

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, OrderbyCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" where company is not null "


If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--ALL--" then
		whereCons = wherecons & " "
	Elseif v_txtGroupPlantation = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
	end if
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--ALL--" then
		selectCons = selectCons & " , Plantation"
		whereCons = wherecons & " "
	Elseif v_txtPlantation = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , Plantation"
		whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
	end if
End If

If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--ALL--" then
		selectCons = selectCons & " , groupregion"
		whereCons = wherecons & " "
	Elseif v_txtGroupRegion = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , groupregion"
		whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
	end if
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--ALL--" then
		selectCons = selectCons & " , region"
		whereCons = wherecons & " "
	Elseif v_txtRegion = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , region"
		whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
	end if
End If

If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--ALL--" then
		selectCons = selectCons & " , company"
		whereCons = wherecons & " "
	Elseif v_txtCompany = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , company"
		whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
		q = " Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit" & _
			" where FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		   CompanyName=rsCat("FCCOMPANYNAME")			   
		end if
	end if
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--ALL--" then
		selectCons = selectCons & " , NoUrut, FCBA"
		whereCons = wherecons & " "
	Elseif v_txtfcba = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , NoUrut, FCBA, VEHICLECODE"
		whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
		dim q, FcbaName, CompanyName, MgrName, KTUName, CropName, FcbaType
		q = "Select fccode, fccompanycode, FCNAME, FCCOMPANYNAME, FCMANAGER, FCKTU, FCCROP, FCTYPE from BusinessUnit where fcCode = '" & v_txtfcba & "'"
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then	
		   fcbaname		= rsCat("fcName")
		   CompanyName	= rsCat("FCCOMPANYNAME")
		   MgrName		= rsCat("FCMANAGER")
		   KTUName		= rsCat("FCKTU")
		   CropName		= rsCat("FCCROP")
		   FcbaType		= rsCat("fcType")
		end if
	end if
End If

v_txtTanggal	=	Request("txtTanggal")
v_txtVehicleGroup		=	Request("txtVehicleGroup")


If (v_txtTanggal = "" Or Len(Trim(v_txtTanggal)) = 0) Then
	v_txtTanggal = textDate(Date)
End If

If v_txtVehicleGroup = "" Then
	v_txtVehicleGroup	=	"--ALL--"
End If

If Len(Trim(v_txtVehicleGroup)) > 0 Then
	if v_txtVehicleGroup = "--ALL--" then
		whereCons = wherecons & " "
	Elseif v_txtVehicleGroup = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		whereCons = whereCons & " And  VEHICLEGROUPCODE = '" & v_txtVehicleGroup & "' "
	end if
End If

Dim v_Tahun : v_Tahun = Year(v_txtTanggal)
Dim v_Bulan : v_Bulan = Month(v_txtTanggal)

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT " & strFieldCons & _
			"     , VEHICLEGROUPCODE" & _
			"     , VEHICLEGROUPNAME" & _
			"     , URUTVEHICLEGROUP" & _
			"     , VEHICLECODE" & _
			"     , 1 TOTALVEHICLE" & _
			"     , SUM(KMHM_HI) KMHM_HI" & _
			"     , SUM(TONASE_HI) TONASE_HI" & _
			"     , SUM(KMHM_SHI) KMHM_SHI" & _
			"     , SUM(TONASE_SHI) TONASE_SHI" & _
			"     , SUM(KMHM_SBI) KMHM_SBI" & _
			"     , SUM(TONASE_SBI) TONASE_SBI" & _
			" FROM (    " & _
			"     SELECT" & _
			"         FCBA" & _
			"         , TANGGAL" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"         , VEHICLECODE" & _
			"         , SUM(KMHM) KMHM_HI" & _
			"         , SUM(TONASE) TONASE_HI" & _
			"         , 0 KMHM_SHI" & _
			"         , 0 TONASE_SHI" & _
			"         , 0 KMHM_SBI" & _
			"         , 0 TONASE_SBI" & _
			"     FROM (" & _
			"         SELECT" & _
			"             FCBA" & _
			"             , VEHICLEDATE TANGGAL" & _
			"             , VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA) VEHICLEGROUPCODE" & _
			"             , VEHICLEGROUP_API.GET_VEHICLEGROUP_NAME(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) VEHICLEGROUPNAME" & _
			"             , VEHICLEGROUP_API.GET_ORDERREPORT(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) URUTVEHICLEGROUP" & _
			"             , VEHICLECODE" & _
			"             , NVL(SUM(NVL(UNIT, 0)),0) KMHM" & _
			"             , NVL(SUM(NVL(OUTPUT, 0)),0) TONASE" & _
			"         FROM" & _
			"             VEHICLEACTIVITY" & _
			"         WHERE" & _
			"             JOBCODE IN ('6502201','6502203')" & _
			"         GROUP BY" & _
			"             FCBA" & _
			"             , VEHICLEDATE" & _
			"             , VEHICLECODE" & _
			"         )" & _
			"     WHERE VEHICLEGROUPCODE IN ('DT', 'TK', 'TR') AND TANGGAL = TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY')" & _
			"     GROUP BY" & _
			"         FCBA" & _
			"         , TANGGAL" & _
			"         , VEHICLECODE" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"     UNION ALL" & _
			"     SELECT" & _
			"         FCBA" & _
			"         , TANGGAL" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"         , VEHICLECODE" & _
			"         , 0 KMHM_HI" & _
			"         , 0 TONASE_HI" & _
			"         , SUM(KMHM) KMHM_SHI" & _
			"         , SUM(TONASE) TONASE_SHI" & _
			"         , 0 KMHM_SBI" & _
			"         , 0 TONASE_SBI" & _
			"     FROM (" & _
			"         SELECT" & _
			"             FCBA" & _
			"             , VEHICLEDATE TANGGAL" & _
			"             , VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA) VEHICLEGROUPCODE" & _
			"             , VEHICLEGROUP_API.GET_VEHICLEGROUP_NAME(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) VEHICLEGROUPNAME" & _
			"             , VEHICLEGROUP_API.GET_ORDERREPORT(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) URUTVEHICLEGROUP" & _
			"             , VEHICLECODE" & _
			"             , NVL(SUM(NVL(UNIT, 0)),0) KMHM" & _
			"             , NVL(SUM(NVL(OUTPUT, 0)),0) TONASE" & _
			"         FROM" & _
			"             VEHICLEACTIVITY" & _
			"         WHERE" & _
			"             JOBCODE IN ('6502201','6502203')" & _
			"         GROUP BY" & _
			"             FCBA" & _
			"             , VEHICLEDATE" & _
			"             , VEHICLECODE" & _
			"         )" & _
			"     WHERE" & _
			"         VEHICLEGROUPCODE IN ('DT', 'TK', 'TR')" & _
			"         AND TO_NUMBER(TO_CHAR(TANGGAL,'YYYY')) = " & v_Tahun & "" & _
			"         AND TO_NUMBER(TO_CHAR(TANGGAL,'MM')) = " & v_Bulan & "" & _
			"         AND TANGGAL <= TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY')" & _
			"     GROUP BY" & _
			"         FCBA" & _
			"         , TANGGAL" & _
			"         , VEHICLECODE" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"     UNION ALL" & _
			"     SELECT" & _
			"         FCBA" & _
			"         , TANGGAL" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"         , VEHICLECODE" & _
			"         , 0 KMHM_HI" & _
			"         , 0 TONASE_HI" & _
			"         , 0 KMHM_SHI" & _
			"         , 0 TONASE_SHI" & _
			"         , SUM(KMHM) KMHM_SBI" & _
			"         , SUM(TONASE) TONASE_SBI" & _
			"     FROM (" & _
			"         SELECT" & _
			"             FCBA" & _
			"             , VEHICLEDATE TANGGAL" & _
			"             , VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA) VEHICLEGROUPCODE" & _
			"             , VEHICLEGROUP_API.GET_VEHICLEGROUP_NAME(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) VEHICLEGROUPNAME" & _
			"             , VEHICLEGROUP_API.GET_ORDERREPORT(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) URUTVEHICLEGROUP" & _
			"             , VEHICLECODE" & _
			"             , NVL(SUM(NVL(UNIT, 0)),0) KMHM" & _
			"             , NVL(SUM(NVL(OUTPUT, 0)),0) TONASE" & _
			"         FROM" & _
			"             VEHICLEACTIVITY" & _
			"         WHERE" & _
			"             JOBCODE IN ('6502201','6502203')" & _
			"         GROUP BY" & _
			"             FCBA" & _
			"             , VEHICLEDATE" & _
			"             , VEHICLECODE" & _
			"         )" & _
			"     WHERE" & _
			"         VEHICLEGROUPCODE IN ('DT', 'TK', 'TR')" & _
			"         AND TO_NUMBER(TO_CHAR(TANGGAL,'YYYY')) = " & v_Tahun & "" & _
			"         AND TO_NUMBER(TO_CHAR(TANGGAL,'MM')) <= " & v_Bulan & "" & _
			"     GROUP BY" & _
			"         FCBA" & _
			"         , TANGGAL" & _
			"         , VEHICLECODE" & _
			"         , VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUPNAME" & _
			"         , URUTVEHICLEGROUP" & _
			"     )" & _
			" GROUP BY" & _
			"     FCBA" & _
			"     , VEHICLEGROUPCODE" & _
			"     , VEHICLEGROUPNAME" & _
			"     , URUTVEHICLEGROUP" & _
			"     , VEHICLECODE"

			
strOKPT2	=	"        , VEHICLEGROUPCODE" & _
				"        , VEHICLEGROUPNAME" & _
				"        , URUTVEHICLEGROUP" & _
				"        , SUM (TOTALVEHICLE) TOTALVEHICLE" & _
				"     , SUM(KMHM_HI) KMHM_HI" & _
				"     , SUM(TONASE_HI) TONASE_HI" & _
				"     , SUM(KMHM_SHI) KMHM_SHI" & _
				"     , SUM(TONASE_SHI) TONASE_SHI" & _
				"     , SUM(KMHM_SBI) KMHM_SBI" & _
				"     , SUM(TONASE_SBI) TONASE_SBI"
			
strOKPT	=	" select GroupPlantation " & selectCons & strOKPT2 & "  From (" & strMain & " )" & whereCons & _
			" group by GroupPlantation" & selectCons & " , VEHICLEGROUPCODE, VEHICLEGROUPNAME, URUTVEHICLEGROUP" & _
			" Order by GroupPlantation" & selectCons &" , URUTVEHICLEGROUP"
						
'response.write strOKPT
%>
<html>

<head>
<title>ePlaS Reporting - Output Kendaraan Pengangkutan TBS</title>
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
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

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
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSum.html"-->
				<table id="form-table" class="centered-margin">	
					<tr>
						<td><font face="verdana" size="1">TANGGAL</font></td>
						<td align="center">
						<% if v_txtTanggal <> "" and v_txtTanggal = v_txtTanggal then %>
							<input class="datepicker" type="text" maxlength="10" name="txtTanggal" size="10" value="<%=v_txtTanggal%>" title="Type Date mm/dd/yyyy" >
						<% else %>
							<input class="datepicker" type="text" maxlength="10" name="txtTanggal" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% end if%>
							<!--<a href="javascript:show_calendar('frmEntry.txtTanggal');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
								<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
							</a>-->
						</td>
						<td><font face="verdana" size="1">GROUP KENDARAAN</font></td>
						<td>
							<select name="txtVehicleGroup" >
								<option value="--ALL--"<%If v_txtVehicleGroup = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsVehicleGroup, strVehicleGroup
								strVehicleGroup	=	" SELECT DISTINCT VEHICLEGROUPCODE, DESCRIPTION FROM VEHICLEGROUP" & _
													" WHERE VEHICLEGROUPCODE IN ('DT','TK','TR') ORDER BY VEHICLEGROUPCODE"
								
								Set rsVehicleGroup = Server.CreateObject("ADODB.Recordset")
								rsVehicleGroup.Open strVehicleGroup, connstr, adOpenStatic 
								
								if not rsVehicleGroup.eof then
								Do While Not rsVehicleGroup.eof
								%>

								<option value="<%=rsVehicleGroup("VEHICLEGROUPCODE")%>" <% If v_txtVehicleGroup = rsVehicleGroup("VEHICLEGROUPCODE") Then %> Selected <%End if%>><%=rsVehicleGroup("VEHICLEGROUPCODE")%> : <%=rsVehicleGroup("DESCRIPTION")%></option> 
								<%
									rsVehicleGroup.movenext
								loop
								end if
								rsVehicleGroup.Close
								set rsVehicleGroup = Nothing
								%>
							</select>
						</td>
						<td><input type="submit" value="View" /></td>
					</tr>				
				</table>
			</div>
			</form>
   <table border="0" width="100%" id="product-table" class="xthf-site">
        <thead>
        <tr>
        <th colspan="38" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">    
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			</div>
			<div id="subheader" class="centered">
				<h2>OUTPUT KENDARAAN PENGANGKUTAN TBS</h2>
				<h5>TANGGAL: <%=v_txtTanggal%></h5>
			</div>
		</div>
        </th>
    </tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP KENDARAAN</th>
		<% if v_txtGroupPlantation <> "--SUMMARY--" and v_txtGroupPlantation <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
		%>
	    <%end if %>
		
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">COMPANY</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">SITE</th>
			<% if v_txtfcba <>  "--ALL--" then %>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">KODE KENDARAAN</th>
			<% end if %>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<th rowspan="2" style="background-color:#00A300;color:#fff;">JUMLAH KENDARAAN</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">HARI INI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">SAMPAI HARI INI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">KM-HM</th>
		<th style="background-color:#00A300;color:#fff;">TONASE</th>
		<th style="background-color:#00A300;color:#fff;">OUTPUT</th>
		<th style="background-color:#00A300;color:#fff;">KM-HM</th>
		<th style="background-color:#00A300;color:#fff;">TONASE</th>
		<th style="background-color:#00A300;color:#fff;">OUTPUT</th>
		<th style="background-color:#00A300;color:#fff;">KM-HM</th>
		<th style="background-color:#00A300;color:#fff;">TONASE</th>
		<th style="background-color:#00A300;color:#fff;">OUTPUT</th>
	</tr>
        </thead>
 <tbody>
<%
Set rsOKPT = Server.CreateObject("ADODB.Recordset")
rsOKPT.Open strOKPT, connstr, adOpenStatic

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempVehicleGroupCode : tempVehicleGroupCode = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim SprevUraian : SprevUraian = ""

if (rsOKPT.bof and rsOKPT.eof) then
%>
	<tr>
		<td colspan="<%=nSpanTotal+11%>">NO DATA AVAILABLE</td>
	</tr>
<%else
		while not rsOKPT.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<% if v_txtVehicleGroup <>  "--SUMMARY--" and  v_txtVehicleGroup <>  "" then %>
						<% tempVehicleGroupCode = rsOKPT("VEHICLEGROUPCODE")
						   tempGroupPlantation = "--ALL--"
						   tempPlantation = "--SUMMARY--"
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempVehicleGroupCode <> SprevUraian then %>
							<td align="left">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
								<%=rsOKPT("VEHICLEGROUPNAME")%>
								</a>
							</td>
							<% SprevUraian = tempVehicleGroupCode %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
				<% end if %>

				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   tempGroupPlantation = rsOKPT("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsOKPT("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsOKPT("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsOKPT("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsOKPT("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsOKPT("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsOKPT("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsOKPT("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsOKPT("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsOKPT("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsOKPT("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsOKPT("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsOKPT("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsOKPT("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsOKPT("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsOKPT("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsOKPT("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsOKPT("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsOKPT("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsOKPT("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsOKPT("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsOKPT("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsOKPT("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
						    <%=rsOKPT("fcba")%>
							</a>
						</td>
						<% if v_txtFcba <>  "--ALL--" Then%>
							<td align="center"><%=rsOKPT("VEHICLECODE")%></td>
						<% End if%>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>
						<% if v_txtFcba <>  "--ALL--" Then%>
							<td align="center"><%=rsOKPT("VEHICLECODE")%></td>
						<% End if%>
					<% end if %>
				<% end if %>
				
				<td align="right"><%=rsOKPT("TOTALVEHICLE")%></td>
				
				<td align="right"><%=FormatNumber(rsOKPT("KMHM_HI"), 0)%></td>
				<td align="right"><%=FormatNumber(rsOKPT("TONASE_HI"), 0)%></td>
				<%If Cdbl(rsOKPT("TOTALVEHICLE")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 0)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsOKPT("TONASE_HI"))/Cdbl(rsOKPT("TOTALVEHICLE")), 2), 0)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsOKPT("KMHM_SHI"), 0)%></td>
				<td align="right"><%=FormatNumber(rsOKPT("TONASE_SHI"), 0)%></td>	
				<%If Cdbl(rsOKPT("TOTALVEHICLE")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 0)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsOKPT("TONASE_SHI"))/Cdbl(rsOKPT("TOTALVEHICLE")), 2), 0)%></td>
				<%End If%>

				<td align="right"><%=FormatNumber(rsOKPT("KMHM_SBI"), 0)%></td>
				<td align="right"><%=FormatNumber(rsOKPT("TONASE_SBI"), 0)%></td>	
				<%If Cdbl(rsOKPT("TOTALVEHICLE")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 0)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round(Cdbl(rsOKPT("TONASE_SBI"))/Cdbl(rsOKPT("TOTALVEHICLE")), 2), 0)%></td>
				<%End If%>
			</tr>
<%
			rsOKPT.MoveNext
		Wend
	rsOKPT.Close
	set rsOKPT = Nothing
%>

<%End if%>
      <tbody>
</table>
		</div>
	</body>
</html>