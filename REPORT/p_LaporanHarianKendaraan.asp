<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtTanggal, v_txtVehicleGroup, v_TxtTypeFCBA, v_txtVehicleCode
Dim ScriptName
Dim rsKendaraan, strKendaraan, strMain, strKendaraan2
Dim cClass, nday, xNum

Dim var_colspan : var_colspan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_TxtTypeFCBA			= Request("TxtTypeFCBA")
v_txtVehicleCode		= Request("txtVehicleCode")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" where company is not null "

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			CompanyName = rsCat("FCCOMPANYNAME")	
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

v_txtTanggal	=	Request("txtTanggal")

If Len(Trim(v_txtTanggal)) = 0 Then
	v_txtTanggal	=	textDate(Date)
End If

v_txtVehicleGroup	=	Request("txtVehicleGroup")

If Len(Trim(v_txtVehicleGroup)) = 0 Then
	v_txtVehicleGroup	=	"--ALL--"
End If

If Len(Trim(v_txtVehicleGroup)) > 0 Then
	if v_txtVehicleGroup = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtVehicleGroup = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And  VEHICLEGROUPCODE = '" & v_txtVehicleGroup & "'"
		orderCons	=	orderCons	& " "
	End If
End If

If Len(Trim(v_txtVehicleCode)) > 0 Then
	if v_txtVehicleCode = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtVehicleCode = "--ALL--" then
		selectCons	=	selectCons	& " , VEHICLECODE"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , VEHICLECODE"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , VEHICLECODE"
		whereCons	=	wherecons	& " And  VEHICLECODE = '" & v_txtVehicleCode & "'"
		orderCons	=	orderCons	& " , VEHICLECODE"
		var_colspan	=	var_colspan	+ 1
	End If
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" SELECT" & strFieldCons & _
				"     , VEHICLEGROUPCODE" & _
				"     , DESCRIPTION" & _
				"     , VEHICLECODE" & _
				"     , VEHICLEGROUP_API.get_OrderReport(VEHICLEGROUPCODE, FCBA) URUTVEHICLEGROUP" & _
				"     , SUM (TOTALVEHICLE) TOTALVEHICLE" & _
				"     , SUM (TOTALOPERATION) TOTALOPERATION" & _
				"     , ROUND(SUM (TOTALACTIVITYUNIT), 2) TOTALACTIVITYUNIT" & _
				"     FROM (" & _
				"         SELECT" & _
				"             A.FCBA" & _
				"             , A.VEHICLEGROUPCODE" & _
				"             , A.DESCRIPTION" & _
				"             , B.VEHICLECODE" & _
				"             , B.TOTALVEHICLE" & _
				"             , 0 TOTALOPERATION" & _
				"             , 0 TOTALACTIVITYUNIT" & _
				"         FROM VEHICLEGROUP A," & _
				"             (" & _
				"             SELECT " & _
				"                 FCBA" & _
				"                 , VEHICLEGROUPCODE" & _
				"                 , FCCODE VEHICLECODE" & _
				"                 , COUNT (*) TOTALVEHICLE" & _
				"             FROM VEHICLE" & _
				"             WHERE" & _
				"                 ACTIVATION = 'Y' " & _
				"                 OR (ACTIVATION = 'N' AND  TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') < DATETERMINATE)" & _
				"             GROUP BY" & _
				"                 FCBA" & _
				"                 , VEHICLEGROUPCODE" & _
				"                 , FCCODE" & _
				"             ) B" & _
				"         WHERE" & _
				"             A.FCBA = B.FCBA" & _
				"             AND A.VEHICLEGROUPCODE = B.VEHICLEGROUPCODE" & _
				"         UNION" & _
				"         SELECT" & _
				"             A.FCBA" & _
				"             , B.VEHICLEGROUPCODE" & _
				"             , C.DESCRIPTION" & _
				"             , A.VEHICLECODE" & _
				"             , 0 TOTALVEHICLE" & _
				"             , COUNT (*) TOTALOPERATION" & _
				"             , 0 TOTALACTIVITYUNIT" & _
				"         FROM (" & _
				"             SELECT " & _
				"                 DISTINCT FCBA" & _
				"                 , VEHICLECODE" & _
				"             FROM (" & _
				"                 SELECT " & _
				"                     A.FCBA" & _
				"                     , A.VEHICLECODE" & _
				"                 FROM VEHICLEACTIVITY A" & _
				"                 WHERE A.VEHICLEDATE = TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"                 )" & _
				"             ) A, VEHICLE B, VEHICLEGROUP C" & _
				"             WHERE " & _
				"                 A.VEHICLECODE = B.FCCODE" & _
				"                 AND B.VEHICLEGROUPCODE = C.VEHICLEGROUPCODE" & _
				"                 AND A.FCBA = B.FCBA" & _
				"                 AND B.FCBA = C.FCBA" & _
				"             GROUP BY " & _
				"             A.FCBA" & _
				"             , B.VEHICLEGROUPCODE" & _
				"             , C.DESCRIPTION" & _
				"             , A.VEHICLECODE" & _
				"         UNION" & _
				"         SELECT " & _
				"             A.FCBA" & _
				"             , B.VEHICLEGROUPCODE" & _
				"             , C.DESCRIPTION" & _
				"             , B.FCCODE VEHICLECODE" & _
				"             , 0 TOTALVEHICLE" & _
				"             , 0 TOTALOPERATION" & _
				"             , SUM (A.UNIT) TOTALACTIVITYUNIT" & _
				"         FROM VEHICLEACTIVITY A, VEHICLE B, VEHICLEGROUP C" & _
				"         WHERE" & _
				"             A.VEHICLEDATE = TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"             AND A.VEHICLECODE = B.FCCODE" & _
				"             AND B.VEHICLEGROUPCODE = C.VEHICLEGROUPCODE" & _
				"             AND A.FCBA = B.FCBA" & _
				"             AND B.FCBA = C.FCBA" & _
				"         GROUP BY " & _
				"             A.FCBA" & _
				"             , B.VEHICLEGROUPCODE" & _
				"             , C.DESCRIPTION" & _
				"             , B.FCCODE" & _
				"         ) where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = '" & v_TxtTypeFCBA &"' " & _
				"     GROUP BY" & _
				"         FCBA" & _
				"         , VEHICLEGROUPCODE" & _
				"         , DESCRIPTION" & _
				"         , VEHICLECODE" & _
				"     ORDER BY BUSINESSUNIT_API.get_OrderData(fcba)"

								
strKendaraan2	=	"     , VEHICLEGROUPCODE, DESCRIPTION, SUM (TOTALVEHICLE) TOTALVEHICLE" & _
					"     , SUM (TOTALOPERATION) TOTALOPERATION" & _
					"     , SUM (TOTALACTIVITYUNIT) TOTALACTIVITYUNIT"

strKendaraan	=	" select GroupPlantation " & selectCons & strKendaraan2 & "  From (" & strMain & " )" & _
					whereCons & _
					" group by GroupPlantation " & selectCons & ", VEHICLEGROUPCODE, DESCRIPTION, URUTVEHICLEGROUP" & _
					" Order by GroupPlantation " & orderCons & ", URUTVEHICLEGROUP"

'response.write strKendaraan
%>
<html>

<head>
<title>ePlaS Reporting - Laporan Harian Kendaraan</title>
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
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<input type="hidden" name="TxtTypeFCBA" value="<%=v_TxtTypeFCBA%>" >
				<!--#include virtual="/REPORT/FormSelectNewSumME.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TANGGAL</font>
						</td>
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
						<td>
							<font face="verdana" size="1">GRUP KENDARAAN</font>
						</td>
						<td>
							<select name="txtVehicleGroup" >
								<option value="--ALL--"<%If v_txtVehicleGroup = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsVehicleGroup, strVehicleGroup
								strVehicleGroup	=	" SELECT DISTINCT VEHICLEGROUPCODE, DESCRIPTION FROM VEHICLEGROUP" & _
													" ORDER BY VEHICLEGROUPCODE"
								
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
						<td><font face="verdana" size="1">KODE KENDARAAN</font></td>
						<td>
							<select name="txtVehicleCode" >
								<option value="--SUMMARY--"<%If v_txtVehicleCode = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtVehicleCode = "--ALL--" then%> Selected <%End If%>>--ALL--</option>								
								<%
								Dim rsVehicle, strVehicle
								strVehicle	=	" SELECT DISTINCT FCCODE, FCNAME FROM VEHICLE" & _
												" WHERE ACTIVATION = 'Y'" & _
												" ORDER BY FCNAME"
								
								Set rsVehicle = Server.CreateObject("ADODB.Recordset")
								rsVehicle.Open strVehicle, connstr, adOpenStatic 
								
								if not rsVehicle.eof then
								Do While Not rsVehicle.eof
								%>

								<option value="<%=rsVehicle("FCCODE")%>" <% If v_txtVehicleCode = rsVehicle("FCCODE") Then %> Selected <%End if%>><%=rsVehicle("FCCODE")%> : <%=rsVehicle("FCNAME")%></option> 
								<%
									rsVehicle.movenext
								loop
								end if
								rsVehicle.Close
								set rsVehicle = Nothing
								%>
							</select>
						</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
            
   <table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
        <tr>
         <th colspan="<%=var_colspan+8%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
         
         <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>LAPORAN HARIAN KENDARAAN</h3>
                <br/>
                <h4>TANGGAL&nbsp;:&nbsp;<%=v_txtTanggal%></h4>
            </div>
		</div>


          </th>
        </tr>
    
	<tr>
		<%If Len(Trim(v_txtVehicleGroup)) > 0 And v_txtVehicleGroup <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;">GROUP KENDARAAN</th>
		<%End If%>
		
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
		
		<%If Len(Trim(v_txtVehicleCode)) > 0 And v_txtVehicleCode <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;">KENDARAAN</th>
		<%End If%>
		
		<th rowspan="3" style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">JUMLAH UNIT</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">PEMAKAIAN</th>
	</tr>
	<tr>
		<th colspan="2" style="background-color:#00A300;color:#fff;">
        <a class="quick_link" <%IF v_TxtTypeFCBA = "M" then %> href="#" <%ELSE%>href="p_OutputKendTBS.asp?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=v_txtGroupPlantation%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=v_txtVehicleGroup%>" target="_blank"<%END IF%> >OPERASI</a>
        </th>
		<th colspan="2" style="background-color:#00A300;color:#fff;"><a class="quick_link" href="p_LaporanKendaraanTidakOperasi.asp?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=v_txtGroupPlantation%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=v_txtVehicleGroup%>" target="_blank">TIDAK OPERASI</a></th>
		<th colspan="2" style="background-color:#00A300;color:#fff;"><a class="quick_link" href="p_DailyVehicleSolar.asp?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=v_txtGroupPlantation%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtdateto=<%=v_txtTanggal%>&txtVehicleGroup=<%=v_txtVehicleGroup%>" target="_blank">KM-HM-JAM</a></th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th style="background-color:#00A300;color:#fff;">RATA-RATA / KENDARAAN</th>
	</tr>
    	</thead>
<tbody>
<%
Set rsKendaraan = Server.CreateObject("ADODB.Recordset")
rsKendaraan.Open strKendaraan, connstr, adOpenStatic

Dim TotalVehicle : TotalVehicle = 0
Dim TotalOperation : TotalOperation = 0
Dim TotalActivityUnit : TotalActivityUnit = 0
Dim TotalNotOperation : TotalNotOperation = 0
Dim EfisiensiUnit : EfisiensiUnit = 0
Dim PersenOperation : PersenOperation = 0
Dim PersenNotOperation : PersenNotOperation = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempVehicleGroup : tempVehicleGroup = ""
Dim tempVehicleCode : tempVehicleCode = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim sprevVehicleGroup : sprevVehicleGroup = ""
Dim sprevVehicleCode : sprevVehicleCode = ""

if rsKendaraan.bof and rsKendaraan.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+8%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsKendaraan.eof
%>
				<tr>
				
				<% if v_txtVehicleGroup <>  "--SUMMARY--" and  v_txtVehicleGroup <>  "" then %>
					<%	tempVehicleGroup = rsKendaraan("VehicleGroupCode")
						If Cint(visibleP) > 1 Then
							tempGroupPlantation = "--ALL--"
						Else
							If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then
								tempGroupPlantation = rsKendaraan("GroupPlantation")
							Else
								tempGroupPlantation = ""
							End If
						End If
						
						If Cint(visibleGR) > 1 And Cint(visibleP) > 1 Then
							tempPlantation = "--SUMMARY--"
						Else
							If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then
								tempPlantation = rsKendaraan("Plantation")
							Else
								tempPlantation = ""
							End If
						End If
						
						If Cint(visibleR) > 1 And Cint(visibleGR) > 1 Then
							tempGroupRegion = "--SUMMARY--"
						Else
							If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then
								tempGroupRegion = rsKendaraan("GroupRegion")
							Else
								tempGroupRegion = ""
							End If
						End If
						
						If Cint(visibleC) > 1 And Cint(visibleR) > 1 Then
							tempRegion = "--SUMMARY--"
						Else
							If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then
								tempRegion = rsKendaraan("Region")
							Else
								tempRegion = ""
							End If
						End If
						
						If Cint(visibleS) > 1 And Cint(visibleC) > 1 Then
							tempCompany = "--SUMMARY--"
						Else
							If Len(Trim(v_txtcompany)) > 0 And v_txtcompany <>  "--SUMMARY--" Then
								tempCompany = rsKendaraan("Company")
							Else
								tempCompany = ""
							End If
						End If
						
						If Cint(VisibleS) > 1 Then
							tempfcba = "--SUMMARY--"
						Else
							If Len(Trim(v_txtFcba)) > 0 And v_txtFcba <>  "--SUMMARY--" Then
								tempfcba = rsKendaraan("Fcba")
							Else
								tempfcba = ""
							End If
						End If
						tempVehicleCode = "--SUMMARY--"
					%>
						<% if tempVehicleGroup <> SprevVehicleGroup then %>
							<td align="center">
								<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
								<%=rsKendaraan("Description")%>
								</a>
							</td>
							<% SprevVehicleGroup = tempVehicleGroup %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
				<% end if %>
				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroup = rsKendaraan("VehicleGroupCode")
					   else
					      tempVehicleGroup = ""
					   end if
					   tempGroupPlantation = rsKendaraan("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempVehicleCode = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
							<%=rsKendaraan("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroup = rsKendaraan("VehicleGroupCode")
					   else
					      tempVehicleGroup = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsKendaraan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsKendaraan("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempVehicleCode = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
							<%=rsKendaraan("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroup = rsKendaraan("VehicleGroupCode")
					   else
					      tempVehicleGroup = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsKendaraan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsKendaraan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsKendaraan("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempVehicleCode = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
							<%=rsKendaraan("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroup = rsKendaraan("VehicleGroupCode")
					   else
					      tempVehicleGroup = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsKendaraan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsKendaraan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsKendaraan("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsKendaraan("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempVehicleCode = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
							<%=rsKendaraan("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroup = rsKendaraan("VehicleGroupCode")
					   else
					      tempVehicleGroup = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsKendaraan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsKendaraan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsKendaraan("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsKendaraan("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsKendaraan("Company")
					   tempfcba = "--ALL--"
					   tempVehicleCode = "--SUMMARY--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
							<%=rsKendaraan("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroup = rsKendaraan("VehicleGroupCode")
					   else
					      tempVehicleGroup = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsKendaraan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsKendaraan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsKendaraan("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsKendaraan("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsKendaraan("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsKendaraan("fcba")
					   tempVehicleCode = "--ALL--"
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>">
						    <%=rsKendaraan("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtVehicleCode <>  "--SUMMARY--" and v_txtVehicleCode <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroup = rsKendaraan("VehicleGroupCode")
					   else
					      tempVehicleGroup = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsKendaraan("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsKendaraan("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsKendaraan("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsKendaraan("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsKendaraan("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba =rsKendaraan("fcba")
					   else
					       tempfcba =""
					   end if
					   
					   tempVehicleCode = rsKendaraan("VEHICLECODE")	   
					 %>
					
					<% if tempVehicleCode <> sprevVehicleCode then %>
						<td align="center">
						    <%=rsKendaraan("VEHICLECODE")%>
						</td>
						<% sprevVehicleCode = tempVehicleCode %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>

				
				<%
				TotalVehicle	=	Round(Cdbl(rsKendaraan("TOTALVEHICLE")), 0)
				TotalOperation	=	Round(Cdbl(rsKendaraan("TOTALOPERATION")), 0)
				TotalActivityUnit	=	Round(Cdbl(rsKendaraan("TOTALACTIVITYUNIT")), 0)
				TotalNotOperation	=	Round(TotalVehicle - TotalOperation, 0)
				
				If TotalOperation > 0 Then
					EfisiensiUnit	=	Round(TotalActivityUnit / TotalOperation, 0)
				Else
					EfisiensiUnit	=	Round(0, 0)
				End if
				
				If TotalVehicle > 0 Then
					PersenOperation		=	Round(TotalOperation / TotalVehicle * 100, 0)
					PersenNotOperation	=	Round(TotalNotOperation / TotalVehicle * 100, 0)
				Else
					PersenOperation	=	Round(0, 0)
					PersenNotOperation	=	Round(0, 0)
				End if
				
				%>

					<td align="right"><%=FbpNum2(TotalVehicle)%></td>
					<td align="right"><%=FbpNum2(TotalOperation)%></td>
					<td align="right"><%=FbpNum2(PersenOperation)%></td>
					<td align="right"><%=FbpNum2(TotalNotOperation)%></td>
					<td align="right"><%=FbpNum2(PersenNotOperation)%></td>
					<td align="right"><%=FbpNum2(TotalActivityUnit)%></td>
					<td align="right"><%=FbpNum2(EfisiensiUnit)%></td>
					
				</tr>
<%
				TotalVehicle		=	0
				TotalOperation		=	0
				TotalActivityUnit	=	0
				TotalNotOperation	=	0
				EfisiensiUnit		=	0
			
			rsKendaraan.MoveNext
			
			Wend	
rsKendaraan.Close
set rsKendaraan = Nothing
%>
<%End if%>
   </tbody> 
</table>
	
   </body>
   </html>