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

v_txtDate = Request("txtDate")

If "" & v_txtDate = "" Then v_txtDate = TextDate(date)

v_txtAfdeling			=	Request("txtAfdeling")	'' Division / Afdeling
v_txtStatus				=	Request("txtStatus")	'' Status	
v_txtCrop				=	Request("txtCrop")		'' Crop
v_txtSoil				=	Request("txtSoil")		'' Soil
v_txtTerrain			=	Request("txtTerrain")	'' Terrain
v_txtOwnership			=	Request("txtOwnership")	'' Ownership
v_txtHarvestingSection	=	Request("txtHarvestingSection")	'' HarvestingSection
v_txtActivation			=	Request("txtActivation")	'' Activation

v_txtBlok			=	Request("txtBlok")		'' Blok
v_txtFcba			=	Request("txtFcba") 		'' Business Unit


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
<!--Additional -->
strPusingan	=   " SELECT fccode, division, round(hectarageplanted,2) hectarageplanted, " & _
			    " 	areal_lc, cadangan, okupasi, bibitan, pabrik, emplasmen, parit, jalan, prasarana_lainnya," & _
				" 	quarry, tanggul, enclave, tambang, konservasi, bukit_sungai, no_imagery, round(luas_gross,2) luas_gross," & _
				" 	ROUND(luas_gross - (areal_lc+cadangan+okupasi+bibitan+pabrik+emplasmen+parit+jalan+prasarana_lainnya+" & _
				" 	quarry+tanggul+enclave+tambang+konservasi+bukit_sungai+no_imagery),2) as Ha_Rounding" & _
				
				" 	, Status, Planting_Date, OWNERSHIP, koperasi " & _
				" 	, Planting_Material, TerrainType, SoilType, CropType " & _
				" 	, Total_Stand " & _
				" 	, decode(hectarageplanted,0,0,round(Total_Stand/hectarageplanted,0)) as Stand_PKK " & _
				" 	, Main_Road, Collection_Road, Main_Drain, Collection_Drain, Subsidiary_Drain " & _
				" 	, Collection_Point, Harvesting_Bridge, Permanent_Bridge, Wood_Bridge " & _
				" 	, Previous_Semester_ABW, Present_Semester_ABW, Harvesting_Based_ABW " & _
				" 	, IsActiveHarvested, HarvestingSection,round(hectarageharvested,2) hectarageharvested, productivestand, Activation, Valid_From,year,month " & _
				
				" 	FROM ( " & _
				" SELECT year,month,fccode" & _
				"        , division" & _
				"        , hectarageplanted" & _
				"        , nvl(areal_lc,0) areal_lc" & _
				"        , nvl(cadangan,0) cadangan" & _
				"        , nvl(okupasi,0) okupasi" & _
				"        , nvl(bibitan,0) bibitan" & _
				"        , nvl(pabrik,0) pabrik" & _
				"        , nvl(emplasmen,0) emplasmen" & _
				"        , nvl(parit,0) parit" & _
				"        , nvl(jalan,0) jalan" & _
				"        , nvl(prasarana_lainnya,0) prasarana_lainnya" & _
				"        , nvl(quarry,0) quarry" & _
				"        , nvl(tanggul,0) tanggul" & _
				"        , nvl(enclave,0) enclave" & _
				"        , nvl(tambang,0) tambang" & _
				"        , nvl(konservasi,0) konservasi" & _
				"        , nvl(bukit_sungai,0) bukit_sungai" & _
				"        , nvl(no_imagery,0) no_imagery" & _
				"        , luas_gross" & _
				"        ,  DX_BLOCK_STATUS as Status		" & _		
				"        ,  PlantingDate				Planting_Date" & _
				"        ,  Ownership" & _
				"        ,  HarvestingSection" & _
			    "		 ,  field_info_01 KOPERASI" & _
				"        ,  PlantingMaterial   Planting_Material" & _
				"        ,  TerrainType" & _
				"        ,  SoilType" & _
				"        ,  CropType" & _
				"        ,  TotalStand         Total_Stand" & _
				"        ,  Main_Road" & _
				"        ,  CollectionRoad     Collection_Road" & _
				"        ,  MainDrain          Main_Drain" & _
				"        ,  CollectionDrain    Collection_Drain" & _
				"        ,  SubsidiaryDrain    Subsidiary_Drain" & _
				"		 ,  CollectionPoint    Collection_Point" & _
				"        ,  HarvestingBridge   Harvesting_Bridge" & _
				"        ,  PermanentBridge    Permanent_Bridge" & _
				"        ,  WoodBridge         Wood_Bridge" & _
				"        ,  PrevSemester_ABW   Previous_Semester_ABW" & _
				"        ,  PresSemester_ABW   Present_Semester_ABW" & _
				"        ,  HarvestingBased_ABW   Harvesting_Based_ABW" & _
				"        ,  Activation " & _
				"        , IsActive_Harvested IsActiveHarvested" & _
				"        , Valid_From " & _
				"        , hectarageharvested " & _
				"        , productivestand " & _
				" 		FROM field_hist_tab" & _
				" 		WHERE fcba = '" & v_txtFcba & "'" & _
				" ) x " & _
				" WHERE 1=1 " 
	
	 ''Filter Include_Date / Valid_From 
	 If Len(Trim(v_txtDate)) > 0 Then
		 strPusingan	= strPusingan	& "  AND Year = EXTRACT(year from TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')) AND Month=EXTRACT(Month from TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy'))"
	 Else
		 strPusingan	= strPusingan	& ""
	 End If	
	 
	''Filter Division / Afdeling
	If v_txtAfdeling <> "ALL" And Len(Trim(v_txtAfdeling)) > 0 Then
		strPusingan	= strPusingan	& "  AND division = '" & v_txtAfdeling & "'"
	Else
		strPusingan	= strPusingan	& ""
	End If
	
	''Filter Status
	If v_txtStatus <> "ALL" And Len(Trim(v_txtStatus)) > 0 Then
		strPusingan	= strPusingan	& "  AND Status = '" & v_txtStatus & "'"
	Else
		strPusingan	= strPusingan	& ""
	End If
	
	''Filter Crop
	If v_txtCrop <> "ALL" And Len(Trim(v_txtCrop)) > 0 Then
		strPusingan	= strPusingan	& "  AND CropType = '" & v_txtCrop & "'"
	Else
		strPusingan	= strPusingan	& ""
	End If
	
	''Filter Soil
	If v_txtSoil <> "ALL" And Len(Trim(v_txtSoil)) > 0 Then
		strPusingan	= strPusingan	& "  AND SoilType = '" & v_txtSoil & "'"
	Else
		strPusingan	= strPusingan	& ""
	End If
	
	''Filter Terrain
	If v_txtTerrain <> "ALL" And Len(Trim(v_txtTerrain)) > 0 Then
		strPusingan	= strPusingan	& "  AND TerrainType = '" & v_txtTerrain & "'"
	Else
		strPusingan	= strPusingan	& ""
	End If
	
	''Filter OWNERSHIP
	If v_txtOwnership <> "ALL" And Len(Trim(v_txtOwnership)) > 0 Then
		strPusingan	= strPusingan	& " AND Ownership = '" & v_txtOwnership & "'"
	Else
		strPusingan	= strPusingan	& " "
	End If		
			
	' ''Filter HarvestingSection
	If v_txtHarvestingSection <> "ALL" And Len(Trim(v_txtHarvestingSection)) > 0 Then
		strPusingan	= strPusingan	& "  AND HarvestingSection = '" & v_txtHarvestingSection & "'"
	Else
		strPusingan	= strPusingan	& ""
	End If	
			
	''Filter Activation
	If Len(Trim(v_txtActivation)) > 0 Then
		strPusingan	= strPusingan	& " AND Activation = '" & v_txtActivation & "'"
	Else
		strPusingan	= strPusingan	& " "
	End If
	
	strPusingan	= strPusingan	& "ORDER BY fccode "
	
	''response.write "Hasil:" + strPusingan 

	
	' If Len(Trim(v_txtYear)) > 0 Then
		' 'strPusingan	=	strPusingan	&	"      AND TO_NUMBER (TO_CHAR (hi.ondate, 'YYYY')) = " & v_txtYear
	' End If
	' If Len(Trim(v_txtMonth)) > 0 Then
		' 'strPusingan	=	strPusingan	&	"      AND TO_NUMBER (TO_CHAR (hi.ondate, 'MM')) = " & v_txtMonth
	' End If		
	' If v_txtAfdeling <> "ALL" And Len(Trim(v_txtAfdeling)) > 0 Then
		' 'strPusingan	=	strPusingan	&	"      AND field_api.getvalid_afdeling(hi.fieldcode, hi.fcba, hi.ondate) = '" & v_txtAfdeling & "'"
	' Else
		' 'strPusingan	=	strPusingan	&	" "
	' End If

	' If v_txtBlok <> "ALL" And Len(Trim(v_txtBlok)) > 0 Then
		' 'strPusingan	=	strPusingan	&	"      AND hi.fieldcode = '" & v_txtBlok & "'"
	' Else
		' strPusingan	=	strPusingan	&	" "
	' End If



									
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

'response.write strPusingan
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
<title>ePlaS Reporting - Areal Statement</title>
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

<body style="padding-top:10em">
	
		<center>
			<p>Manajemen ARESTA dialihkan ke Janjang Emas. Klik link di bawah ini:</p>
			<a href="https://janjangemas.id/" target="_blank" title="Menuju ke Janjang Emas">Janjang Emas &raquo;</a>
	</center>
	<%Response.End%>
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
						<!--<td><font face="verdana" size="1">TAHUN</font></td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon

							Yr = Year(Session("v_Costbook"))
							MinYr = Yr - 2
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
							</td>-->
							
							<!-- Tanggal -->
							<td>
								<font face="verdana" size="1">PERIOD</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDate" value="<%=v_txtDate%>" />
							</td>
							
							<!-- DIVISION / AFDELING -->
							<td>
								<font face="verdana" size="1">DIVISION</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="ALL">--ALL--</option>
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
							
							<!-- STATUS -->
							<td>
								<font face="verdana" size="1">STATUS</font>
							</td>
							<td width="2">
								<select name="txtStatus" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsStatus, strStatus
									strStatus	=	" SELECT DISTINCT P.FCCODE,P.FCNAME FROM PARAMETER P WHERE P.PARCODE= 'F0006' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
									Set rsStatus = Server.CreateObject("ADODB.Recordset")
									rsStatus.Open strStatus, connstr, adOpenStatic 
									
									if not rsStatus.eof then
										Do While Not rsStatus.eof
										if rsStatus("FCCODE") = v_txtStatus then%>
										<option value="<%=rsStatus("FCCODE")%>" Selected><%=rsStatus("FCCODE")%></option> 
										<%Else%>
										<option value="<%=rsStatus("FCCODE")%>"><%=rsStatus("FCCODE")%></option>  
										<%End if
											rsStatus.movenext
										loop
									end if
									rsStatus.Close
									set rsStatus = Nothing
									%>
								</select>
							</td>
							
							<!-- CROP -->
							<td>
								<font face="verdana" size="1">CROP</font>
							</td>
							<td width="2">
								<select name="txtCrop" >
									<!-- <option value="ALL">--ALL--</option> -->
									 <option value="">--ALL--</option> 
									<%
									Dim rsCrop, strCrop
									strCrop	=	" SELECT DISTINCT P.FCCODE,P.FCNAME FROM PARAMETER P WHERE P.PARCODE= 'F0005' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
									Set rsCrop = Server.CreateObject("ADODB.Recordset")
									rsCrop.Open strCrop, connstr, adOpenStatic 
									
									if not rsCrop.eof then
										Do While Not rsCrop.eof
										if rsCrop("FCCODE") = v_txtCrop then%>
										<option value="<%=rsCrop("FCCODE")%>" Selected><%=rsCrop("FCNAME")%></option> 
										<%Else%>
										<option value="<%=rsCrop("FCCODE")%>"><%=rsCrop("FCNAME")%></option>  
										<%End if
											rsCrop.movenext
										loop
									end if
									rsCrop.Close
									set rsCrop = Nothing
									%>
								</select>
							</td>
							
							<!-- SOIL -->
							<td>
								<font face="verdana" size="1">SOIL</font>
							</td>
							<td width="2">
								<select name="txtSoil" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsSoil, strSoil
									strSoil	=	" SELECT DISTINCT P.FCCODE,P.FCNAME FROM PARAMETER P WHERE P.PARCODE= 'F0003' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
									Set rsSoil = Server.CreateObject("ADODB.Recordset")
									rsSoil.Open strSoil, connstr, adOpenStatic 
									
									if not rsSoil.eof then
										Do While Not rsSoil.eof
										if rsSoil("FCCODE") = v_txtSoil then%>
										<option value="<%=rsSoil("FCCODE")%>" Selected><%=rsSoil("FCNAME")%></option> 
										<%Else%>
										<option value="<%=rsSoil("FCCODE")%>"><%=rsSoil("FCNAME")%></option>  
										<%End if
											rsSoil.movenext
										loop
									end if
									rsSoil.Close
									set rsSoil = Nothing
									%>
								</select>
							</td>
							
							<!-- TERRAIN -->
							<td>
								<font face="verdana" size="1">TERRAIN</font>
							</td>
							<td width="2">
								<select name="txtTerrain" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsTerrain, strTerrain
									strTerrain	=	" SELECT DISTINCT P.FCCODE,P.FCNAME FROM PARAMETER P WHERE P.PARCODE= 'F0002' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
									Set rsTerrain = Server.CreateObject("ADODB.Recordset")
									rsTerrain.Open strTerrain, connstr, adOpenStatic 
									
									if not rsTerrain.eof then
										Do While Not rsTerrain.eof
										if rsTerrain("FCCODE") = v_txtTerrain then%>
										<option value="<%=rsTerrain("FCCODE")%>" Selected><%=rsTerrain("FCNAME")%></option> 
										<%Else%>
										<option value="<%=rsTerrain("FCCODE")%>"><%=rsTerrain("FCNAME")%></option>  
										<%End if
											rsTerrain.movenext
										loop
									end if
									rsTerrain.Close
									set rsTerrain = Nothing
									%>
								</select>
							</td>
							
							<!-- OWNERSHIP -->
							<td>
								<font face="verdana" size="1">OWNERSHIP</font>
							</td>
							<td width="2">
								<select name="txtOwnership" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsOwnership, strOwnership
									strOwnership	=	" SELECT DISTINCT P.FCCODE,P.FCNAME FROM PARAMETER P WHERE P.PARCODE= 'F0004' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
									Set rsOwnership = Server.CreateObject("ADODB.Recordset")
									rsOwnership.Open strOwnership, connstr, adOpenStatic 
									
									if not rsOwnership.eof then
										Do While Not rsOwnership.eof
										if rsOwnership("FCCODE") = v_txtOwnership then%>
										<option value="<%=rsOwnership("FCCODE")%>" Selected><%=rsOwnership("FCCODE")%></option> 
										<%Else%>
										<option value="<%=rsOwnership("FCCODE")%>"><%=rsOwnership("FCCODE")%></option>  
										<%End if
											rsOwnership.movenext
										loop
									end if
									rsOwnership.Close
									set rsOwnership = Nothing
									%>
								</select>
							</td>
							
							<!-- HARVESTING SECTION -->
							<td>
								<font face="verdana" size="1">HARV.SECTION</font>
							</td>
							<td width="2">
								<select name="txtHarvestingSection" >
									<option value = "" 
										<%If Len(Trim(txtHARVESTINGSECTION)) = 0 then%> Selected <%End If%>>OTHERS
									</option>
									<%
									Dim rsHarvestingSection, strHarvestingSection
									strHarvestingSection =	" SELECT '' FCCODE, 'ALL' FCNAME FROM DUAL " & _
															" UNION " & _
															" SELECT FCCODE, FCNAME FROM HARVESTINGGROUP" & _
															" WHERE FCBA = '" & Session("FCBA") & "'" & _
															" ORDER BY FCCODE"
									
									Set rsHarvestingSection = Server.CreateObject("ADODB.Recordset")
									rsHarvestingSection.Open strHarvestingSection, connstr, adOpenStatic 
									
									if not rsHarvestingSection.eof then
										Do While Not rsHarvestingSection.eof
										if rsHarvestingSection("FCCODE") = v_txtHarvestingSection then%>
										<option value="<%=rsHarvestingSection("FCCODE")%>" Selected><%=rsHarvestingSection("FCCODE")%></option> 
										<%Else%>
										<option value="<%=rsHarvestingSection("FCCODE")%>"><%=rsHarvestingSection("FCCODE")%></option>  
										<%End if
											rsHarvestingSection.movenext
										loop
									end if
									rsHarvestingSection.Close
									set rsHarvestingSection = Nothing
									%>
									
								</select>
							</td>
							
							<!-- ACTIVATION -->
							<td>
								<font face="verdana" size="1">ACTIVATION</font>
							</td>
							<td width="2">
								<select style="min-width:100px" name="txtActivation" title="ACTIVATION Character a s/d z or 1 s/d 9">
									<%If Len(Trim(Request("ACTIVATION"))) = 0 then%> 
										<Option value="" selected></Option>
										<Option value="Y" >Yes</Option>
										<Option value="N">No</Option>
									<%elseIf Request("ACTIVATION") = "Y" then%> 
										<Option value=""></Option>
										<Option value="Y" selected>Yes</Option>
										<Option value="N">No</Option>
									<%elseIf Request("ACTIVATION") = "N" then%>
										<Option value=""></Option>
										<Option value="Y">Yes</Option>
										<Option value="N" selected>No</Option>
									<%end if%>
								</select>
							</td>
							<!-- View -->
							<td>
								<input type="submit" value="View" />
							</td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
		
				
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=48%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>AREAL STATEMENT</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
<% 
dim selec_day
	if len(trim(DAY(v_txtDate))) = 0 then
		selec_day = GetLastDay(v_txtMonth, v_txtYear)
	else
		selec_day = DAY(v_txtDate)
	end if
%>
	<tr>
		<th rowspan="2"  style="background-color:#00A300;color:#fff;">Blok</th>
		<th rowspan="2"  style="background-color:#00A300;color:#fff;">Division</th>
		<th rowspan="2"  style="background-color:#00A300;color:#fff;">Ha Planted</th>
		<th colspan="17" style="background-color:#00A300;color:#fff;">Land Use</th>
		<th colspan="11" style="background-color:#00A300;color:#fff;">Information</th>
		<th colspan="9"  style="background-color:#00A300;color:#fff;">Infrastructure ( Meter )</th>
		
		<th colspan="3"  style="background-color:#00A300;color:#fff;">Average Bunch Weight ( ABW )</th>
		<th colspan="4"  style="background-color:#00A300;color:#fff;">Harvesting</th>
		
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Active Or No</th>
		
	</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">Areal Sudah di LC</th>
			<th style="background-color:#00A300;color:#fff;">Cadangan</th>
			<th style="background-color:#00A300;color:#fff;">Okupasi</th>
			<th style="background-color:#00A300;color:#fff;">Bibitan</th>
			<th style="background-color:#00A300;color:#fff;">Pabrik</th>
			<th style="background-color:#00A300;color:#fff;">Emplasmen</th>
			<th style="background-color:#00A300;color:#fff;">Parit </th>
			<th style="background-color:#00A300;color:#fff;">Jalan</th>
			<th style="background-color:#00A300;color:#fff;">Prasarana Lainnya</th>
			<th style="background-color:#00A300;color:#fff;">Quarry</th>
			<th style="background-color:#00A300;color:#fff;">Tanggul</th>
			<th style="background-color:#00A300;color:#fff;">Enclave</th>
			<th style="background-color:#00A300;color:#fff;">Tambang</th>
			<th style="background-color:#00A300;color:#fff;">Konservasi</th>
			<th style="background-color:#00A300;color:#fff;">Bukit Sungai</th>
			<th style="background-color:#00A300;color:#fff;">No Imagery</th>
			<th style="background-color:#00A300;color:#fff;">Luas Gross</th>
			
			<th style="background-color:#00A300;color:#fff;">Ha Rounding</th>
			<th style="background-color:#00A300;color:#fff;">Status</th>
			<th style="background-color:#00A300;color:#fff;">Planting Date</th>
			<th style="background-color:#00A300;color:#fff;">Ownership</th>
			<th style="background-color:#00A300;color:#fff;">Koperasi</th>
			<th style="background-color:#00A300;color:#fff;">Planting Material</th>
			<th style="background-color:#00A300;color:#fff;">Terrain Type</th>
			<th style="background-color:#00A300;color:#fff;">Soil Type</th>
			<th style="background-color:#00A300;color:#fff;">Crop Type</th>
			<th style="background-color:#00A300;color:#fff;">Total Stand</th>
			<th style="background-color:#00A300;color:#fff;">Stand/Ha</th>

			<th style="background-color:#00A300;color:#fff;">Main Road</th>
			<th style="background-color:#00A300;color:#fff;">Collection Road</th>
			<th style="background-color:#00A300;color:#fff;">Main Drain</th>
			<th style="background-color:#00A300;color:#fff;">Collection Drain</th>
			<th style="background-color:#00A300;color:#fff;">Subsidiary Drain</th>
			
			<th style="background-color:#00A300;color:#fff;">Collection Point</th>
			<th style="background-color:#00A300;color:#fff;">Harvesting Bridge</th>
			<th style="background-color:#00A300;color:#fff;">Permanent Bridge</th>
			<th style="background-color:#00A300;color:#fff;">Wood Bridge</th>
			
			<th style="background-color:#00A300;color:#fff;">Previous Semester ABW</th>
			<th style="background-color:#00A300;color:#fff;">Present Semester ABW</th>
			<th style="background-color:#00A300;color:#fff;">Harvesting Based ABW</th>
			
			<th style="background-color:#00A300;color:#fff;">Active</th>
			<th style="background-color:#00A300;color:#fff;">Ha.Harvest</th>
			<th style="background-color:#00A300;color:#fff;">Stand.Harvest</th>
			<th style="background-color:#00A300;color:#fff;">Section</th>
	</tr>
    </thead>
	<tbody>
<%
''response.write strPusingan

Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan, connstr, adOpenStatic

Dim xNum

Dim Total_Ha : Total_Ha = 0
Dim Total_LastInterval : Total_LastInterval = 0
Dim DataIsPersen : DataIsPersen = 0
Dim CountData : CountData = 0


if rsPusingan.bof and rsPusingan.eof then

%>
	<tr>
		<td colspan="50">NO DATA AVAILABLE</td>
	</tr>
<%end if


		while not rsPusingan.eof
			CountData = CountData + 1
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>

			<tr class="<%=cClass%>">
				<td align="center"><%=rsPusingan("fccode")%></td>
				<td align="center"><%=rsPusingan("division")%></td>
				<td align="center"><%=rsPusingan("hectarageplanted")%></td>
				
				<td align="center"><%=rsPusingan("areal_lc")%></td>
				<td align="center"><%=rsPusingan("cadangan")%></td>
				<td align="center"><%=rsPusingan("okupasi")%></td>
				<td align="center"><%=rsPusingan("bibitan")%></td>
				<td align="center"><%=rsPusingan("pabrik")%></td>
				<td align="center"><%=rsPusingan("emplasmen")%></td>
				<td align="center"><%=rsPusingan("parit")%></td>
				<td align="center"><%=rsPusingan("jalan")%></td>
				<td align="center"><%=rsPusingan("prasarana_lainnya")%></td>
				<td align="center"><%=rsPusingan("quarry")%></td>
				<td align="center"><%=rsPusingan("tanggul")%></td>
				<td align="center"><%=rsPusingan("enclave")%></td>
				
				<td align="center"><%=rsPusingan("tambang")%></td>
				<td align="center"><%=rsPusingan("konservasi")%></td>
				<td align="center"><%=rsPusingan("bukit_sungai")%></td>
				<td align="center"><%=rsPusingan("no_imagery")%></td>
				<td align="center"><%=rsPusingan("luas_gross")%></td>
				
				<td align="center"><%=rsPusingan("Ha_Rounding")%></td>
				<td align="center"><%=rsPusingan("Status")%></td>
				<td align="center"><%=rsPusingan("Planting_Date")%></td>
				<td align="center"><%=rsPusingan("Ownership")%></td>
				<td align="center"><%=rsPusingan("koperasi")%></td>
				<td align="center"><%=rsPusingan("Planting_Material")%></td>
				
				<td align="center"><%=rsPusingan("TerrainType")%></td>
				<td align="center"><%=rsPusingan("SoilType")%></td>
				<td align="center"><%=rsPusingan("CropType")%></td>
				<td align="center"><%=rsPusingan("Total_Stand")%></td>
				<td align="center"><%=rsPusingan("Stand_PKK")%></td>
				
				<td align="center"><%=rsPusingan("Main_Road")%></td>
				<td align="center"><%=rsPusingan("Collection_Road")%></td>
				<td align="center"><%=rsPusingan("Main_Drain")%></td>
				<td align="center"><%=rsPusingan("Collection_Drain")%></td>
				<td align="center"><%=rsPusingan("Subsidiary_Drain")%></td>
				
				<td align="center"><%=rsPusingan("Collection_Point")%></td>
				<td align="center"><%=rsPusingan("Harvesting_Bridge")%></td>
				<td align="center"><%=rsPusingan("Permanent_Bridge")%></td>
				<td align="center"><%=rsPusingan("Wood_Bridge")%></td>
				
				<td align="center"><%=rsPusingan("Previous_Semester_ABW")%></td>
				<td align="center"><%=rsPusingan("Present_Semester_ABW")%></td>
				<td align="center"><%=rsPusingan("Harvesting_Based_ABW")%></td>
				
				<td align="center"><%=rsPusingan("IsActiveHarvested")%></td>
				<td align="center"><%=rsPusingan("hectarageharvested")%></td>
				<td align="center"><%=rsPusingan("productivestand")%></td>
				<td align="center"><%=rsPusingan("HarvestingSection")%></td>
			
			    <td align="center"><%=rsPusingan("Activation")%></td>
			</tr>

<%
			rsPusingan.MoveNext
		Wend
rsPusingan.Close
set rsPusingan = Nothing
'response.write strPusingan
'response.write  Total_LastInterval
'response.write  Total_Ha 
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
