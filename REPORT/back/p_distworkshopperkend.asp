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
Dim rsDWPK, strDWPK, strDWPK2, strMain
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
			"     , URUTVEHICLEGROUP    " & _
			"     , VEHICLECODE" & _
			"     , SUM(TOTALVEHICLE) TOTALVEHICLE" & _
			"     , SUM(HI) HI" & _
			"     , SUM(SHI) SHI" & _
			"     , SUM(SBI) SBI" & _
			" FROM (" & _
			"     SELECT" & _
			"         FCBA" & _
			"         , VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA) VEHICLEGROUPCODE" & _
			"         , VEHICLEGROUP_API.GET_VEHICLEGROUP_NAME(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) VEHICLEGROUPNAME" & _
			"         , VEHICLEGROUP_API.GET_ORDERREPORT(VEHICLE_API.GET_VEHICLEGROUPCODE(VEHICLECODE, FCBA), FCBA) URUTVEHICLEGROUP    " & _
			"         , VEHICLECODE" & _
			"     	  , SUM(TOTALVEHICLE) TOTALVEHICLE" & _
			"         , SUM(HI) HI" & _
			"         , SUM(SHI) SHI" & _
			"         , SUM(SBI) SBI" & _
			"     FROM (" & _
			"         SELECT" & _
			"             FCBA" & _
			"             , TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY') TANGGAL" & _
			"             , FCCODE VEHICLECODE" & _
			"             , 1 TOTALVEHICLE" & _
			"             , 0 HI" & _
			"             , 0 SHI" & _
			"             , 0 SBI" & _
			"         FROM" & _
			"             VEHICLE" & _
			"         WHERE ACTIVATION = 'Y'" & _
			"         UNION ALL" & _
			"         SELECT" & _
			"             FCBA" & _
			"             , WORKSHOPDATE TANGGAL" & _
			"             , LOCATIONCODE VEHICLECODE" & _
			"             , 0 TOTALVEHICLE" & _
			"             , NVL(UNIT,0) HI" & _
			"             , 0 SHI" & _
			"             , 0 SBI" & _
			"         FROM" & _
			"             WORKSHOPACTIVITY" & _
			"         WHERE" & _
			"             LOCATIONTYPE = 'VH'" & _
			"             AND WORKSHOPDATE = TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY')" & _
			"         UNION ALL" & _
			"         SELECT" & _
			"             FCBA" & _
			"             , WORKSHOPDATE TANGGAL" & _
			"             , LOCATIONCODE VEHICLECODE" & _
			"             , 0 TOTALVEHICLE" & _
			"             , 0 HI" & _
			"             , NVL(UNIT,0) SHI" & _
			"             , 0 SBI" & _
			"         FROM" & _
			"             WORKSHOPACTIVITY" & _
			"         WHERE" & _
			"             LOCATIONTYPE = 'VH'" & _
			"             AND TO_NUMBER(TO_CHAR(WORKSHOPDATE,'YYYY')) = " & v_Tahun & "" & _
			"             AND TO_NUMBER(TO_CHAR(WORKSHOPDATE,'MM')) = " & v_Bulan & "" & _
			"             AND WORKSHOPDATE <= TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY')" & _
			"         UNION ALL    " & _
			"         SELECT" & _
			"             FCBA" & _
			"             , WORKSHOPDATE TANGGAL" & _
			"             , LOCATIONCODE VEHICLECODE" & _
			"             , 0 TOTALVEHICLE" & _
			"             , 0 HI" & _
			"             , 0 SHI" & _
			"             , NVL(UNIT,0) SBI" & _
			"         FROM" & _
			"             WORKSHOPACTIVITY" & _
			"         WHERE" & _
			"             LOCATIONTYPE = 'VH'" & _
			"             AND TO_NUMBER(TO_CHAR(WORKSHOPDATE,'YYYY')) = " & v_Tahun & "" & _
			"             AND TO_NUMBER(TO_CHAR(WORKSHOPDATE,'MM')) <= " & v_Bulan & "" & _
			"         )" & _
			"     GROUP BY" & _
			"         FCBA" & _
			"         , VEHICLECODE" & _
			"     )" & _
			" GROUP BY" & _
			"     FCBA" & _
			"     , VEHICLEGROUPCODE" & _
			"     , VEHICLEGROUPNAME" & _
			"     , URUTVEHICLEGROUP    " & _
			"     , VEHICLECODE"

			
strDWPK2	=	"        , VEHICLEGROUPCODE" & _
				"        , VEHICLEGROUPNAME" & _
				"        , URUTVEHICLEGROUP" & _
				"        , SUM (TOTALVEHICLE) TOTALVEHICLE" & _
				"        , SUM (HI) HI" & _
				"        , SUM (SHI) SHI" & _
				"        , SUM (SBI) SBI"
			
			
strDWPK	=	" select GroupPlantation " & selectCons & strDWPK2 & "  From (" & strMain & " )" & whereCons & _
			" group by GroupPlantation" & selectCons & " , VEHICLEGROUPCODE, VEHICLEGROUPNAME, URUTVEHICLEGROUP" & _
			" Order by GroupPlantation" & selectCons &" , URUTVEHICLEGROUP"
						
'response.write strDWPK
%>
<html>

<head>
<title>ePlaS Reporting - Distribusi Kerja Workshop Per Kendaraan</title>
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
						<td><input type="submit" value="View" /></td>
					</tr>				
				</table>
			</div>
			</form>
            
<table border="0" id="product-table" class="xthf-site" width="100%">
	<thead>
        <tr>
          <th colspan="14" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
			<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			</div>
			<div id="subheader" class="centered">
				<h2>DISTRIBUSI KERJA WORKSHOP PER KENDARAAN</h2>
				<h5>TANGGAL: <%=v_txtTanggal%></h5>
			</div>
		</div>
           </th>
     </tr>
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">GROUP KENDARAAN</th>
		<% if v_txtGroupPlantation <> "--SUMMARY--" and v_txtGroupPlantation <> "" then %>
		    <th rowspan="3" style="background-color:#00A300;color:#fff;">GROUP PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
		%>
	    <%end if %>
		
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;">PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;">GROUP REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="3" style="background-color:#00A300;color:#fff;">REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="3" style="background-color:#00A300;color:#fff;">COMPANY</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
				<th rowspan="3" style="background-color:#00A300;color:#fff;">SITE</th>
			<% if v_txtfcba <>  "--ALL--" then %>
				<th rowspan="3" style="background-color:#00A300;color:#fff;">KODE KENDARAAN</th>
			<% end if %>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<th rowspan="3" style="background-color:#00A300;color:#fff;">JUMLAH KENDARAAN</th>
		<th colspan="6" style="background-color:#00A300;color:#fff;">JAM KERJA</th>
	</tr>
	<tr>
		<th colspan="2" style="background-color:#00A300;color:#fff;">HARI INI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">SAMPAI HARI INI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">JAM</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">JAM</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">JAM</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
	</tr>
</thead>
<tbody>
<%
Set rsDWPK = Server.CreateObject("ADODB.Recordset")
rsDWPK.Open strDWPK, connstr, adOpenStatic

Dim rsGetTotal, strGetTotal
strGetTotal = " Select Sum(HI) HI, Sum(SHI) SHI, Sum(SBI) SBI FROM (" & strDWPK & ")"
Set rsGetTotal = Server.CreateObject("ADODB.Recordset")
rsGetTotal.Open strGetTotal, connstr, adOpenStatic

'response.Write "<strong>" & strGetTotal

Dim Temp_Total_HI : Temp_Total_HI = 0
Dim Temp_Total_SHI : Temp_Total_SHI = 0
Dim Temp_Total_SBI : Temp_Total_SBI = 0

Dim	GrandTotal_HI : GrandTotal_HI = 0
Dim GrandTotal_SHI : GrandTotal_SHI = 0
Dim GrandTotal_SBI : GrandTotal_SBI = 0

Dim Total_Persen_HI : Total_Persen_HI = 0
Dim Total_Persen_SHI : Total_Persen_SHI = 0
Dim Total_Persen_SBI : Total_Persen_SBI = 0

Dim GrandTotal_Persen_HI : GrandTotal_Persen_HI = 0
Dim GrandTotal_Persen_SHI : GrandTotal_Persen_SHI = 0
Dim GrandTotal_Persen_SBI : GrandTotal_Persen_SBI = 0

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

If Not rsGetTotal.EOF Then
	Temp_Total_HI = Temp_Total_HI + Cdbl(rsGetTotal("HI"))
	Temp_Total_SHI = Temp_Total_SHI + Cdbl(rsGetTotal("SHI"))
	Temp_Total_SBI = Temp_Total_SBI + Cdbl(rsGetTotal("SBI"))
Else
	Temp_Total_HI = 0
	Temp_Total_SHI = 0
	Temp_Total_SBI = 0
End If
rsGetTotal.Close
set rsGetTotal = Nothing

if (rsDWPK.bof and rsDWPK.eof) then
%>
	<tr>
		<td colspan="<%=1+nSpanTotal+8%>">NO DATA AVAILABLE</td>
	</tr>
<%else
		while not rsDWPK.eof
			Total_Persen_HI = 0
			Total_Persen_SHI = 0
			Total_Persen_SBI = 0
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<% if v_txtVehicleGroup <>  "--SUMMARY--" and  v_txtVehicleGroup <>  "" then %>
						<% tempVehicleGroupCode = rsDWPK("VEHICLEGROUPCODE")
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
								<%=rsDWPK("VEHICLEGROUPNAME")%>
								</a>
							</td>
							<% SprevUraian = tempVehicleGroupCode %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
				<% end if %>

				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsDWPK("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   tempGroupPlantation = rsDWPK("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsDWPK("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsDWPK("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPK("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsDWPK("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsDWPK("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsDWPK("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPK("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDWPK("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsDWPK("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsDWPK("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsDWPK("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPK("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDWPK("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDWPK("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsDWPK("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsDWPK("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsDWPK("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPK("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDWPK("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDWPK("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsDWPK("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsDWPK("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
							<%=rsDWPK("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtVehicleGroup <>  "--SUMMARY--"  and  v_txtVehicleGroup <> "" then
					      tempVehicleGroupCode = rsDWPK("VEHICLEGROUPCODE")
					   else
					      tempVehicleGroupCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPK("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDWPK("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDWPK("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsDWPK("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsDWPK("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsDWPK("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtVehicleGroup=<%=tempVehicleGroupCode%>">
						    <%=rsDWPK("fcba")%>
							</a>
						</td>
						<% if v_txtFcba <>  "--ALL--" Then%>
							<td align="center"><%=rsDWPK("VEHICLECODE")%></td>
						<% End if%>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>
						<% if v_txtFcba <>  "--ALL--" Then%>
							<td align="center"><%=rsDWPK("VEHICLECODE")%></td>
						<% End if%>
					<% end if %>
				<% end if %>
				
<%
			If Temp_Total_HI = 0 Then
				Total_Persen_HI	= 0
				GrandTotal_Persen_HI = 0
			Else
				Total_Persen_HI	= Round((Cdbl(rsDWPK("HI"))/Temp_Total_HI)*100, 2)
				GrandTotal_Persen_HI	=	GrandTotal_Persen_HI + (Cdbl(rsDWPK("HI"))/Temp_Total_HI)*100
			End If
			If Temp_Total_SHI = 0 Then
				Total_Persen_SHI = 0
				GrandTotal_Persen_SHI = 0
			Else
				Total_Persen_SHI = Round((Cdbl(rsDWPK("SHI"))/Temp_Total_SHI)*100, 2)
				GrandTotal_Persen_SHI	=	GrandTotal_Persen_SHI + (Cdbl(rsDWPK("SHI"))/Temp_Total_SHI)*100
			End If
			If Temp_Total_SBI = 0 Then
				Total_Persen_SBI = 0
				GrandTotal_Persen_SBI = 0
			Else
				Total_Persen_SBI = Round((Cdbl(rsDWPK("SBI"))/Temp_Total_SBI)*100, 2)
				GrandTotal_Persen_SBI	=	GrandTotal_Persen_SBI + (Cdbl(rsDWPK("SBI"))/Temp_Total_SBI)*100
			End If
%>
				
				<td align="right"><%=FbpNum2(rsDWPK("TOTALVEHICLE"))%></td>
				
				<td align="right"><%=FormatNumber(rsDWPK("HI"), 2)%></td>				
				<td align="right"><%=FormatNumber(Total_Persen_HI, 2)%></td>
				
				<td align="right"><%=FormatNumber(rsDWPK("SHI"), 2)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_SHI, 2)%></td>

				<td align="right"><%=FormatNumber(rsDWPK("SBI"), 2)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_SBI, 2)%></td>
			</tr>
<%
			GrandTotal_HI	=	GrandTotal_HI + Cdbl(rsDWPK("HI"))
			GrandTotal_SHI	=	GrandTotal_SHI + Cdbl(rsDWPK("SHI"))
			GrandTotal_SBI	=	GrandTotal_HI + Cdbl(rsDWPK("SBI"))
			
			rsDWPK.MoveNext
		Wend
	rsDWPK.Close
	set rsDWPK = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<% if (v_txtfcba = "--ALL--" OR v_txtfcba = "--SUMMARY--") then %>
			<td align="center" colspan="<%=2+nSpanTotal%>"><strong>TOTAL</strong></td>
		<%Else%>
			<td align="center" colspan="<%=2+nSpanTotal+1%>"><strong>TOTAL</strong></td>
		<%End if%>
		<td align="right"><strong><%=FormatNumber(Temp_Total_HI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_HI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Temp_Total_SHI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_SHI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Temp_Total_SBI, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_SBI, 2)%></strong></td>
	</tr>
<%End if%>
</tbody>
</table>
		</div>
	</body>
</html>