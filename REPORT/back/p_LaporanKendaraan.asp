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
		selectCons = selectCons & " , NoUrut, FCBA"
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
				"     , DESCRIPTION" & _
				"     , VEHICLEGROUP_API.get_OrderReport(VEHICLEGROUPCODE, FCBA) URUTVEHICLEGROUP" & _
				"     , SUM (TOTALVEHICLE) TOTALVEHICLE" & _
				"     , SUM (TOTALOPERATION) TOTALOPERATION" & _
				"     , ROUND(SUM (TOTALACTIVITYUNIT), 2) TOTALACTIVITYUNIT" & _
				"     FROM (" & _
				"         SELECT" & _
				"             A.FCBA" & _
				"             , A.VEHICLEGROUPCODE" & _
				"             , A.DESCRIPTION" & _
				"             , B.TOTALVEHICLE" & _
				"             , 0 TOTALOPERATION" & _
				"             , 0 TOTALACTIVITYUNIT" & _
				"         FROM VEHICLEGROUP A," & _
				"             (" & _
				"             SELECT " & _
				"                 FCBA" & _
				"                 , VEHICLEGROUPCODE" & _
				"                 , COUNT (*) TOTALVEHICLE" & _
				"             FROM VEHICLE" & _
				"             WHERE" & _
				"                 ACTIVATION = 'Y' " & _
				"                 OR (ACTIVATION = 'N' AND  TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') < DATETERMINATE)" & _
				"             GROUP BY" & _
				"                 FCBA" & _
				"                 , VEHICLEGROUPCODE" & _
				"             ) B" & _
				"         WHERE" & _
				"             A.FCBA = B.FCBA" & _
				"             AND A.VEHICLEGROUPCODE = B.VEHICLEGROUPCODE" & _
				"         UNION" & _
				"         SELECT" & _
				"             A.FCBA" & _
				"             , B.VEHICLEGROUPCODE" & _
				"             , C.DESCRIPTION" & _
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
				"         UNION" & _
				"         SELECT " & _
				"             A.FCBA" & _
				"             , B.VEHICLEGROUPCODE" & _
				"             , C.DESCRIPTION" & _
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
				"         )" & _
				"     GROUP BY" & _
				"         FCBA" & _
				"         , VEHICLEGROUPCODE" & _
				"         , DESCRIPTION" & _
				"     ORDER BY BUSINESSUNIT_API.get_OrderData(fcba)"

								
strKendaraan2	=	"     , VEHICLEGROUPCODE, DESCRIPTION, SUM (TOTALVEHICLE) TOTALVEHICLE" & _
					"     , SUM (TOTALOPERATION) TOTALOPERATION" & _
					"     , SUM (TOTALACTIVITYUNIT) TOTALACTIVITYUNIT"

strKendaraan	=	" select GroupPlantation " & selectCons & strKendaraan2 & "  From (" & strMain & " )" & whereCons & _
					" group by GroupPlantation " & selectCons & ", VEHICLEGROUPCODE, DESCRIPTION, URUTVEHICLEGROUP Order by" & _
					" GroupPlantation " & selectCons & ", URUTVEHICLEGROUP"

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
         <th colspan="13" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
         
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
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP KENDARAAN</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
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
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<th rowspan="2" style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">JUMLAH UNIT</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">KM-HM-JAM</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">EFISIENSI UNIT</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">% TIDAK OPERASI</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">OPERASI</th>
		<th style="background-color:#00A300;color:#fff;">TIDAK OPERASI</th>
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
Dim PersenNotOperation : PersenNotOperation = 0

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
		<td colspan="<%=nSpanTotal+7%>">NO DATA AVAILABLE</td>
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
						<% sprevFcba = tempfcba %>	
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
					PersenNotOperation	=	Round(TotalNotOperation / TotalVehicle * 100, 0)
				Else
					PersenNotOperation	=	Round(0, 0)
				End if
				
				%>

					<td align="right"><%=TotalVehicle%></td>
					<td align="right"><%=TotalOperation%></td>
					<td align="right"><%=TotalNotOperation%></td>
					<td align="right"><%=TotalActivityUnit%></td>
					<td align="right"><%=EfisiensiUnit%></td>
					<td align="right"><%=PersenNotOperation%></td>
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