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
Dim rsKendaraan, strKendaraan, strMain, strKendaraan2
Dim NO, cClass, nday, xNum, nSpanTotal

nSpanTotal = 7

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
		OrderbyCons = OrderbyCons & " , Plantation"
		whereCons = wherecons & " "
	Elseif v_txtPlantation = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , Plantation"
		OrderbyCons = OrderbyCons & " , Plantation"
		whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
	end if
End If

If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--ALL--" then
		selectCons = selectCons & " , groupregion"
		OrderbyCons = OrderbyCons & " , groupregion"
		whereCons = wherecons & " "
	Elseif v_txtGroupRegion = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , groupregion"
		OrderbyCons = OrderbyCons & " , groupregion"
		whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
	end if
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--ALL--" then
		selectCons = selectCons & " , region"
		OrderbyCons = OrderbyCons & " , region"
		whereCons = wherecons & " "
	Elseif v_txtRegion = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , region"
		OrderbyCons = OrderbyCons & " , region"
		whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
	end if
End If

If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--ALL--" then
		selectCons = selectCons & " , company"
		OrderbyCons = OrderbyCons & " , company"
		whereCons = wherecons & " "
	Elseif v_txtCompany = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , company"
		OrderbyCons = OrderbyCons & " , company"
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
		OrderbyCons = OrderbyCons & " , NoUrut"
		whereCons = wherecons & " "
	Elseif v_txtfcba = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , NoUrut, FCBA, VEHICLECODE"
		OrderbyCons = OrderbyCons & " , NoUrut"
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

If v_txtTanggal = "" Then
	v_txtTanggal	=	textDate(Date)
End If

v_txtVehicleGroup	=	Request("txtVehicleGroup")

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

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" SELECT" & strFieldCons & _
				"     , VEHICLEGROUPCODE" & _
				"     , VEHICLEGROUPNAME" & _
				"     , URUTVEHICLEGROUP" & _
				"     , SUM(TOTALVEHICLE) TOTALVEHICLE" & _
				"     , VEHICLECODE" & _
				"     , VEHICLENAME" & _
				"     , SUM(DOWNTIME) DOWNTIME" & _
				"     , DOWNTIME_TYPE" & _
				" FROM (" & _
				"     SELECT" & _
				"         FCBA" & _
				"         , VEHICLEGROUPCODE" & _
				"         , VEHICLEGROUP_API.GET_VEHICLEGROUP_NAME(VEHICLEGROUPCODE, FCBA) VEHICLEGROUPNAME" & _
				"         , VEHICLEGROUP_API.GET_ORDERREPORT(VEHICLEGROUPCODE, FCBA) URUTVEHICLEGROUP" & _
				"         , SUM(TOTALVEHICLE) TOTALVEHICLE" & _
				"         , VEHICLECODE" & _
				"         , VEHICLENAME" & _
				"         , SUM(DOWNTIME) DOWNTIME" & _
				"         , DOWNTIME_TYPE" & _
				"     FROM (" & _
				"         SELECT" & _
				"             FCBA" & _
				"             , VEHICLE_API.GET_VEHICLEGROUPCODE (FCCODE, FCBA) VEHICLEGROUPCODE" & _
				"             , FCCODE VEHICLECODE" & _
				"             , VEHICLE_API.GET_VEHICLECODE_NAME (FCCODE, FCBA) VEHICLENAME" & _
				"             , 1 TOTALVEHICLE" & _
				"             , VEHICLEACTIVITY_API.GETDOWNTIME (FCCODE, TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY'), FCBA) DOWNTIME" & _
				"             , DOWNTIME_TYPE" & _
				"         FROM (" & _
				"             SELECT" & _
				"                 FCBA" & _
				"                 , FCCODE" & _
				"                 , 'STAND BY' DOWNTIME_TYPE" & _
				"             FROM (" & _
				"                 SELECT" & _
				"                     FCBA" & _
				"                     , FCCODE" & _
				"                 FROM VEHICLE" & _
				"                 WHERE" & _
				"                     (ACTIVATION = 'Y'OR DATETERMINATE >    TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY'))" & _
				"             MINUS" & _
				"                 (" & _
				"                 SELECT DISTINCT " & _
				"                     FCBA , VEHICLECODE" & _
				"                 FROM VEHICLEACTIVITY" & _
				"                 WHERE" & _
				"                 VEHICLEDATE = TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"                 UNION" & _
				"                 SELECT DISTINCT " & _
				"                     FCBA" & _
				"                     , LOCATIONCODE" & _
				"                 FROM WORKSHOPACTIVITY" & _
				"                 WHERE" & _
				"                     LOCATIONTYPE = 'VH'" & _
				"                     AND WORKSHOPDATE = TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"                 )" & _
				"                 )" & _
				"             UNION" & _
				"                 SELECT DISTINCT " & _
				"                     FCBA" & _
				"                     , LOCATIONCODE" & _
				"                     , 'BREAK DOWN' DOWNTIME_TYPE" & _
				"                 FROM WORKSHOPACTIVITY" & _
				"                 WHERE" & _
				"                     LOCATIONTYPE = 'VH'" & _
				"                     AND WORKSHOPDATE = TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"                     AND FCBA || LOCATIONCODE NOT IN (" & _
				"                     SELECT DISTINCT FCBA || VEHICLECODE    FROM VEHICLEACTIVITY " & _
				"                     WHERE VEHICLEDATE = TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"                     )" & _
				"             )" & _
				"         ) " & _
				"     GROUP BY" & _
				"         FCBA" & _
				"         , VEHICLEGROUPCODE" & _
				"         , VEHICLEGROUP_API.GET_VEHICLEGROUP_NAME(VEHICLEGROUPCODE, FCBA)" & _
				"         , VEHICLEGROUP_API.GET_ORDERREPORT(VEHICLEGROUPCODE, FCBA)" & _
				"         , VEHICLECODE" & _
				"         , VEHICLENAME" & _
				"         , DOWNTIME_TYPE" & _
				"     )" & _
				" GROUP BY" & _
				"     FCBA" & _
				"     , VEHICLEGROUPCODE" & _
				"     , VEHICLEGROUPNAME" & _
				"     , URUTVEHICLEGROUP" & _
				"     , VEHICLECODE" & _
				"     , VEHICLENAME" & _
				"     , DOWNTIME_TYPE"
				
Dim Selection_2

Selection_2	=	"     , VEHICLEGROUPCODE" & _
				"     , VEHICLEGROUPNAME" & _
				"     , URUTVEHICLEGROUP" & _
				"     , DOWNTIME_TYPE"

strKendaraan2	=	"     , SUM(TOTALVEHICLE) TOTALVEHICLE" & _
					"     , SUM(DOWNTIME) DOWNTIME"

strKendaraan	=	" select GroupPlantation " & selectCons & Selection_2 & strKendaraan2 & "  From (" & strMain & " )" & whereCons & _
					" group by GroupPlantation " & selectCons & Selection_2 & " Order by GroupPlantation " & OrderbyCons & ", URUTVEHICLEGROUP"
					
'response.write strKendaraan
%>
<html>

<head>
<title>ePlaS Reporting - Laporan Kendaraan</title>
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
            <th colspan="10" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
            <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>LAPORAN KENDARAAN</h3>
                <br/>
                <h4>TANGGAL&nbsp;:&nbsp;<%=v_txtTanggal%></h4>
            </div>
		</div>

</th>
</tr>
	<tr>
		<% if v_txtVehicleGroup <>  "--SUMMARY--" and v_txtVehicleGroup <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">GROUP KENDARAAN</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtGroupPlantation <> "--SUMMARY--" and v_txtGroupPlantation <> "" then %>
		      <th style="background-color:#00A300;color:#fff;">GROUP PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
		%>
	    <%end if %>
		
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th style="background-color:#00A300;color:#fff;">PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;">GROUP REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;">REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">COMPANY</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">SITE</th>
			<% if v_txtfcba <>  "--ALL--" then %>
			<th style="background-color:#00A300;color:#fff;">KODE KENDARAAN</th>
			<% end if %>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<th style="background-color:#00A300;color:#fff;">TOTAL KENDARAAN</th>
		<th style="background-color:#00A300;color:#fff;">LAMA RUSAK (HARI)</th>
		<th style="background-color:#00A300;color:#fff;">KETERANGAN</th>
	</tr>
    </thead>
<tbody>
<%
Set rsKendaraan = Server.CreateObject("ADODB.Recordset")
rsKendaraan.Open strKendaraan, connstr, adOpenStatic

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempVehicleGroup : tempVehicleGroup = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim sprevVehicleGroup : sprevVehicleGroup = ""

if rsKendaraan.bof and rsKendaraan.eof then
%>
	<tr>
		<td colspan="12">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsKendaraan.eof
%>
				<tr>

				<% if v_txtVehicleGroup <>  "--SUMMARY--" and  v_txtVehicleGroup <>  "" then %>
						<% tempVehicleGroup = rsKendaraan("VehicleGroupCode")
						   tempGroupPlantation = "--ALL--"
						   tempPlantation = "--SUMMARY--"
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempVehicleGroup <> SprevVehicleGroup then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>">
								<%=rsKendaraan("VEHICLEGROUPNAME")%>
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
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>">
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
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>">
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
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>">
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
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>">
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
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>">
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
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroup%>">
						    <%=rsKendaraan("fcba")%>
							</a>
						</td>
						<% if v_txtFcba <>  "--ALL--" Then%>
							<td align="center"><%=rsKendaraan("VEHICLECODE")%></td>
						<% End if%>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>
						<% if v_txtFcba <>  "--ALL--" Then%>
							<td align="center"><%=rsKendaraan("VEHICLECODE")%></td>
						<% End if%>
					<% end if %>
				<% end if %>
				
					<td align="right"><%=FbpNum2(rsKendaraan("TOTALVEHICLE"))%></td>
					<td align="right"><%=FbpNum2(rsKendaraan("DOWNTIME"))%></td>
					<td align="left"><%=rsKendaraan("DOWNTIME_TYPE")%></td>
				</tr>
<%
			
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
