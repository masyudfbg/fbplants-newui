<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtTanggal, v_txtAfdeling, v_txtBlok, v_TxtTypeFCBA
Dim ScriptName
Dim rsProdWsHarian, strProdWsHarian, strMain, strProdWsHarian2
Dim cClass, nday, xNum

Dim var_ColSpan : var_ColSpan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_TxtTypeFCBA			= Request("TxtTypeFCBA")

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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_ColSpan	=	var_ColSpan	+ 1
		
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_ColSpan	=	var_ColSpan	+ 1
		
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

If (v_txtTanggal = "" Or Len(Trim(v_txtTanggal)) = 0) Then
	v_txtTanggal = textDate(Date)
End If

Dim v_Tahun : v_Tahun = Year(v_txtTanggal)
Dim v_Bulan : v_Bulan = Month(v_txtTanggal)

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" SELECT" & strFieldCons & _
				"     , SUM(HI_HK) HI_HK" & _
				"     , SUM(HI_JAM) HI_JAM" & _
				"     , SUM(SHI_HK) SHI_HK" & _
				"     , SUM(SHI_JAM) SHI_JAM" & _
				"     , SUM(SBI_HK) SBI_HK" & _
				"     , SUM(SBI_JAM) SBI_JAM" & _
				" FROM (    " & _
				"     SELECT" & _
				"         FCBA" & _
				"         , MANDAYS HI_HK" & _
				"         , 0 HI_JAM" & _
				"         , 0 SHI_HK" & _
				"         , 0 SHI_JAM" & _
				"         , 0 SBI_HK" & _
				"         , 0 SBI_JAM" & _
				"     FROM ATTENDANCE_GAD WHERE" & _
				"         LOCATIONTYPE = 'WS'" & _
				"         AND ATTENDANCE IN ('KJ', 'WH', 'WS')" & _
				"         AND FDDATE = TO_DATE('"& v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"     UNION ALL" & _
				"     SELECT" & _
				"         FCBA" & _
				"         , 0 HI_HK" & _
				"         , UNIT HI_JAM" & _
				"         , 0 SHI_HK" & _
				"         , 0 SHI_JAM" & _
				"         , 0 SBI_HK" & _
				"         , 0 SBI_JAM" & _
				"     FROM WORKSHOPACTIVITY WHERE" & _
				"         WORKSHOPDATE = TO_DATE('"& v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"     UNION ALL    " & _
				"     SELECT" & _
				"         FCBA" & _
				"         , 0 HI_HK" & _
				"         , 0 HI_JAM" & _
				"         , MANDAYS SHI_HK" & _
				"         , 0 SHI_JAM" & _
				"         , 0 SBI_HK" & _
				"         , 0 SBI_JAM" & _
				"     FROM ATTENDANCE_GAD WHERE" & _
				"         LOCATIONTYPE = 'WS'" & _
				"         AND ATTENDANCE IN ('KJ', 'WH', 'WS')" & _
				"         AND TO_NUMBER(TO_CHAR(FDDATE,'YYYY')) = " & v_Tahun & "" & _
				"         AND TO_NUMBER(TO_CHAR(FDDATE,'MM')) = " & v_Bulan & "" & _
				"         AND FDDATE <= TO_DATE('"& v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"     UNION ALL" & _
				"     SELECT" & _
				"         FCBA" & _
				"         , 0 HI_HK" & _
				"         , 0 HI_JAM" & _
				"         , 0 SHI_HK" & _
				"         , UNIT SHI_JAM" & _
				"         , 0 SBI_HK" & _
				"         , 0 SBI_JAM" & _
				"     FROM WORKSHOPACTIVITY WHERE" & _
				"         TO_NUMBER(TO_CHAR(WORKSHOPDATE,'YYYY')) = " & v_Tahun & "" & _
				"         AND TO_NUMBER(TO_CHAR(WORKSHOPDATE,'MM')) = " & v_Bulan & "" & _
				"         AND WORKSHOPDATE <= TO_DATE('"& v_txtTanggal & "', 'MM/DD/YYYY')" & _
				"     UNION ALL    " & _
				"     SELECT" & _
				"         FCBA" & _
				"         , 0 HI_HK" & _
				"         , 0 HI_JAM" & _
				"         , 0 SHI_HK" & _
				"         , 0 SHI_JAM" & _
				"         , MANDAYS SBI_HK" & _
				"         , 0 SBI_JAM" & _
				"     FROM ATTENDANCE_GAD WHERE" & _
				"         LOCATIONTYPE = 'WS'" & _
				"         AND ATTENDANCE IN ('KJ', 'WH', 'WS')" & _
				"         AND TO_NUMBER(TO_CHAR(FDDATE,'YYYY')) = " & v_Tahun & "" & _
				"         AND TO_NUMBER(TO_CHAR(FDDATE,'MM')) <= " & v_Bulan & "" & _
				"     UNION ALL" & _
				"     SELECT" & _
				"         FCBA" & _
				"         , 0 HI_HK" & _
				"         , 0 HI_JAM" & _
				"         , 0 SHI_HK" & _
				"         , 0 SHI_JAM" & _
				"         , 0 SBI_HK" & _
				"         , UNIT SBI_JAM" & _
				"     FROM WORKSHOPACTIVITY WHERE" & _
				"         TO_NUMBER(TO_CHAR(WORKSHOPDATE,'YYYY')) = " & v_Tahun & "" & _
				"         AND TO_NUMBER(TO_CHAR(WORKSHOPDATE,'MM')) <= " & v_Bulan & "" & _
				"     ) where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = '" & v_TxtTypeFCBA &"'" & _
				" GROUP BY" & _
				"     FCBA"

								
strProdWsHarian2	=	"     , SUM(HI_HK) HI_HK" & _
						"     , SUM(HI_JAM) HI_JAM" & _
						"     , SUM(SHI_HK) SHI_HK" & _
						"     , SUM(SHI_JAM) SHI_JAM" & _
						"     , SUM(SBI_HK) SBI_HK" & _
						"     , SUM(SBI_JAM) SBI_JAM"

strProdWsHarian	=	" select GroupPlantation " & selectCons & strProdWsHarian2 &"  From (" & strMain & " )" & _
					whereCons & _
					" group by GroupPlantation" & selectCons & _
					" Order by GroupPlantation" & orderCons

'response.write strProdWsHarian
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Laporan Harian Produktivitas Workshop</title>
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
			<th colspan="<%=9+var_ColSpan%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>LAPORAN HARIAN PRODUKTIVITAS WORKSHOP</h3>
					<br/>
					<h4>TANGGAL&nbsp;:&nbsp;<%=v_txtTanggal%></h4>
				</div>
			</th>
		</tr>
	<tr>
		
		<% if v_txtGroupPlantation <> "--SUMMARY--" and v_txtGroupPlantation <> "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<% end if %>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% end if %>
		
		<% if v_txtAfdeling <>  "--SUMMARY--" and v_txtAfdeling <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% end if %>
		
		<% if v_txtBlok <>  "--SUMMARY--" and v_txtBlok <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<% end if %>
		
		<%If var_ColSpan = 0 Then%><th rowspan="2" style="background-color:#00A300;color:#fff;">&nbsp;</th><%End If%>
		
		<th colspan="3" style="background-color:#00A300;color:#fff;">HARI INI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">SAMPAI HARI INI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">HK</th>
		<th style="background-color:#00A300;color:#fff;"><a class="quick_link" href="p_DistHarianWorkshop.asp?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtTanggal=<%=v_txtTanggal%>&txtGroupPlantation=<%=v_txtGroupPlantation%>" target="_blank">JAM</a></th>
		<th style="background-color:#00A300;color:#fff;">JAM/ HK</th>
		<th style="background-color:#00A300;color:#fff;">HK</th>
		<th style="background-color:#00A300;color:#fff;"><a class="quick_link" href="p_DistHarianWorkshop.asp?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtTanggal=<%=v_txtTanggal%>&txtGroupPlantation=<%=v_txtGroupPlantation%>" target="_blank">JAM</a></th>
		<th style="background-color:#00A300;color:#fff;">JAM/ HK</th>
		<th style="background-color:#00A300;color:#fff;">HK</th>
		<th style="background-color:#00A300;color:#fff;"><a class="quick_link" href="p_DistHarianWorkshop.asp?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtTanggal=<%=v_txtTanggal%>&txtGroupPlantation=<%=v_txtGroupPlantation%>" target="_blank">JAM</a></th>
		<th style="background-color:#00A300;color:#fff;">JAM/ HK</th>
	</tr>
    </thead>
	<tbody>
<%
Set rsProdWsHarian = Server.CreateObject("ADODB.Recordset")
rsProdWsHarian.Open strProdWsHarian, connstr, adOpenStatic

Dim val_Hi_HK : val_Hi_HK = 0
Dim val_Hi_Jam : val_Hi_Jam = 0
Dim val_Hi_Jam_per_HK : val_Hi_Jam_per_HK = 0

Dim val_Shi_HK : val_Shi_HK	= 0
Dim val_Shi_Jam : val_Shi_Jam	= 0
Dim val_Shi_Jam_per_HK : val_Shi_Jam_per_HK = 0

Dim val_Sbi_HK : val_Sbi_HK	= 0
Dim val_Sbi_Jam : val_Sbi_Jam	= 0
Dim val_Sbi_Jam_per_HK : val_Sbi_Jam_per_HK = 0


Dim Grand_Hi_HK : Grand_Hi_HK = 0
Dim Grand_Hi_Jam : Grand_Hi_Jam = 0
Dim Grand_Shi_HK : Grand_Shi_HK = 0
Dim Grand_Shi_Jam : Grand_Shi_Jam = 0
Dim Grand_Sbi_HK : Grand_Sbi_HK = 0
Dim Grand_Sbi_Jam : Grand_Sbi_Jam = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""

if rsProdWsHarian.bof and rsProdWsHarian.eof then
%>
	<tr>
		<td colspan="<%=1+var_ColSpan+9%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsProdWsHarian.eof
%>
				<tr>
				
				<%If var_ColSpan = 0 Then%><td align="center">&nbsp;</td><%End If%>
						
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsProdWsHarian("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>">
							<%=rsProdWsHarian("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsProdWsHarian("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsProdWsHarian("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>">
							<%=rsProdWsHarian("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsProdWsHarian("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsProdWsHarian("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsProdWsHarian("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>">
							<%=rsProdWsHarian("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsProdWsHarian("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsProdWsHarian("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsProdWsHarian("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsProdWsHarian("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>">
							<%=rsProdWsHarian("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsProdWsHarian("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsProdWsHarian("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsProdWsHarian("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsProdWsHarian("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsProdWsHarian("Company")
					   tempfcba = "--ALL--"				   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>">
							<%=rsProdWsHarian("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsProdWsHarian("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsProdWsHarian("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsProdWsHarian("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsProdWsHarian("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsProdWsHarian("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsProdWsHarian("fcba")				   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>">
						    <%=rsProdWsHarian("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>

<%
	val_Hi_HK	=	0
	val_Hi_Jam	=	0
	val_Hi_Jam_per_HK = 0
	
	val_Shi_HK	=	0
	val_Shi_Jam	=	0
	val_Shi_Jam_per_HK = 0
	
	val_Sbi_HK	=	0
	val_Sbi_Jam	=	0
	val_Sbi_Jam_per_HK = 0
	
	val_Hi_HK	=	val_Hi_HK + Cdbl(rsProdWsHarian("HI_HK"))
	val_Hi_Jam	=	val_Hi_Jam + Cdbl(rsProdWsHarian("HI_JAM"))
	
	If val_Hi_HK = 0 Then
		val_Hi_Jam_per_HK	=	0
	Else
		val_Hi_Jam_per_HK	=	Round(val_Hi_Jam/ val_Hi_HK, 2)
	End If
	
	val_Shi_HK	=	val_Shi_HK + Cdbl(rsProdWsHarian("SHI_HK"))
	val_Shi_Jam	=	val_Shi_Jam + Cdbl(rsProdWsHarian("SHI_JAM"))
	
	If val_Shi_HK = 0 Then
		val_Shi_Jam_per_HK	=	0
	Else
		val_Shi_Jam_per_HK	=	Round(val_Shi_Jam/ val_Shi_HK, 2)
	End If
	
	val_Sbi_HK	=	val_Sbi_HK + Cdbl(rsProdWsHarian("SBI_HK"))	
	val_Sbi_Jam	=	val_Sbi_Jam + Cdbl(rsProdWsHarian("SBI_JAM"))
	
	If val_Sbi_HK = 0 Then
		val_Sbi_Jam_per_HK	=	0
	Else
		val_Sbi_Jam_per_HK	=	Round(val_Sbi_Jam/ val_Sbi_HK, 2)
	End If
%>
					<td align="right"><%=FormatNumber(val_Hi_HK, 0)%></td>
					<td align="right"><%=FormatNumber(val_Hi_Jam, 0)%></td>
					<td align="right"><%=FormatNumber(val_Hi_Jam_per_HK, 0)%></td>
					
					<td align="right"><%=FormatNumber(val_Shi_HK, 0)%></td>
					<td align="right"><%=FormatNumber(val_Shi_Jam, 0)%></td>
					<td align="right"><%=FormatNumber(val_Shi_Jam_per_HK, 0)%></td>
					
					<td align="right"><%=FormatNumber(val_Sbi_HK, 0)%></td>
					<td align="right"><%=FormatNumber(val_Sbi_Jam, 0)%></td>
					<td align="right"><%=FormatNumber(val_Sbi_Jam_per_HK, 0)%></td>
				</tr>
<%
			Grand_Hi_HK =  Grand_Hi_HK + val_Hi_HK
			Grand_Hi_Jam =  Grand_Hi_Jam + val_Hi_Jam
			
			Grand_Shi_HK =  Grand_Shi_HK + val_Shi_HK
			Grand_Shi_Jam =  Grand_Shi_Jam + val_Shi_Jam
			
			Grand_Sbi_HK =  Grand_Sbi_HK + val_Sbi_HK
			Grand_Sbi_Jam =  Grand_Sbi_Jam + val_Sbi_Jam
			
			rsProdWsHarian.MoveNext			
			Wend	
rsProdWsHarian.Close
set rsProdWsHarian = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td colspan="<%=var_ColSpan%>" align="center"><strong>TOTAL</strong></td>
		<td align="right"><strong><%=FormatNumber(Grand_Hi_HK, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Grand_Hi_Jam, 0)%></strong></td>
		<%If Grand_Hi_HK = 0 Then%>
		<td align="right"><strong><%=FormatNumber(0, 0)%></strong></td>
		<%Else%>
		<td align="right"><strong><%=FormatNumber(Round(Grand_Hi_Jam/Grand_Hi_HK, 2), 0)%></strong></td>
		<%End If%>
		
		<td align="right"><strong><%=FormatNumber(Grand_Shi_HK, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Grand_Shi_Jam, 0)%></strong></td>
		<%If Grand_Shi_HK = 0 Then%>
		<td align="right"><strong><%=FormatNumber(0, 0)%></strong></td>
		<%Else%>
		<td align="right"><strong><%=FormatNumber(Round(Grand_Shi_Jam/Grand_Shi_HK, 2), 0)%></strong></td>
		<%End If%>
		
		<td align="right"><strong><%=FormatNumber(Grand_Sbi_HK, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Grand_Sbi_Jam, 0)%></strong></td>
		<%If Grand_Sbi_HK = 0 Then%>
		<td align="right"><strong><%=FormatNumber(0, 0)%></strong></td>
		<%Else%>
		<td align="right"><strong><%=FormatNumber(Round(Grand_Sbi_Jam/Grand_Sbi_HK, 2), 0)%></strong></td>
		<%End If%>
	</tr>
<%End if%>
	</tbody>
</table>

   </body>
   </html>
