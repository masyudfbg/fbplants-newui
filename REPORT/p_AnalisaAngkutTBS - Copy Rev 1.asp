<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtDateFrom, v_txtDateTo
Dim ScriptName
Dim rsAnalisaAngkutTBS, strAnalisaAngkutTBS, strAnalisaAngkutTBS2, strMain
Dim cClass, nday

Dim var_colspan : var_colspan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")

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

v_txtDateFrom	=	Request("txtDateFrom")
v_txtDateTo		=	Request("txtDateTo")


If Len(Trim(v_txtDateFrom)) = 0 Then
	v_txtDateFrom = textDate(Date)
End If

If Len(Trim(v_txtDateTo)) = 0 Then
	v_txtDateTo = textDate(Date)
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT " & strFieldCons & "" & _
			"        , NO_URUT" & _
			"        , DELIVERY" & _
			"        , SUM (DIRECT_WEIGHT_KG) DIRECT_WEIGHT_KG" & _
			"        , SUM (INDIRECT_WEIGHT_KG) INDIRECT_WEIGHT_KG" & _
			"        , SUM (DIRECT_WEIGHT_KG+INDIRECT_WEIGHT_KG) TOTAL" & _
			" FROM (      " & _
			"       SELECT FCBA," & _
			"              NO_URUT," & _
			"              DELIVERY," & _
			"              SUM (DIRECT_WEIGHT_KG) DIRECT_WEIGHT_KG," & _
			"              SUM (INDIRECT_WEIGHT_KG) INDIRECT_WEIGHT_KG" & _
			"         FROM (SELECT FCCODE FCBA," & _
			"                      1 NO_URUT," & _
			"                      'SENDIRI' DELIVERY," & _
			"                      0 DIRECT_WEIGHT_KG," & _
			"                      0 INDIRECT_WEIGHT_KG" & _
			"                 FROM BUSINESSUNIT" & _
			"               UNION" & _
			"               SELECT FCCODE FCBA," & _
			"                      2 NO_URUT," & _
			"                      'KONTRAK' DELIVERY," & _
			"                      0 DIRECT_WEIGHT_KG," & _
			"                      0 INDIRECT_WEIGHT_KG" & _
			"                 FROM BUSINESSUNIT" & _
			"               UNION" & _
			"                 SELECT FCBA," & _
			"                        1 NO_URUT," & _
			"                        'SENDIRI' DELIVERY," & _
			"                        NVL (SUM (NVL (MILL_WEIGHT, 0)), 0) DIRECT_WEIGHT_KG," & _
			"                        0 INDIRECT_WEIGHT_KG" & _
			"                   FROM HARVESTING" & _
			"                  WHERE       RECEPTIONDATE >= TO_DATE('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
			"                        AND RECEPTIONDATE <= TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
			"                        AND PRODUCTCODE = 'FFB'" & _
			"                        AND OWN = '01'" & _
			"               GROUP BY FCBA" & _
			"               UNION" & _
			"                 SELECT FCBA," & _
			"                        2 NO_URUT," & _
			"                        'KONTRAK' DELIVERY," & _
			"                        NVL (SUM (NVL (MILL_WEIGHT, 0)), 0) DIRECT_WEIGHT_KG," & _
			"                        0 INDIRECT_WEIGHT_KG" & _
			"                   FROM HARVESTING" & _
			"                  WHERE       RECEPTIONDATE >= TO_DATE('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
			"                        AND RECEPTIONDATE <= TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
			"                        AND PRODUCTCODE = 'FFB'" & _
			"                        AND OWN = '02'" & _
			"               GROUP BY FCBA" & _
			"               UNION" & _
			"                 SELECT FCBA," & _
			"                        1 NO_URUT," & _
			"                        'SENDIRI' DELIVERY," & _
			"                        0 DIRECT_WEIGHT_KG," & _
			"                        NVL (SUM (NVL (OUTPUT, 0)), 0) INDIRECT_WEIGHT_KG" & _
			"                   FROM VEHICLEACTIVITY" & _
			"                  WHERE       VEHICLEDATE >= TO_DATE('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
			"                        AND VEHICLEDATE <= TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
			"                        AND JOBCODE = '6502203'" & _
			"                        AND LOCATIONTYPE = 'FF'" & _
			"               GROUP BY FCBA" & _
			"               UNION" & _
			"                 SELECT FCBA," & _
			"                        2 NO_URUT," & _
			"                        'KONTRAK' DELIVERY," & _
			"                        0 DIRECT_WEIGHT_KG," & _
			"                        NVL (SUM (NVL (OUTPUT, 0)), 0) INDIRECT_WEIGHT_KG" & _
			"                   FROM VEHICLEACTIVITY" & _
			"                  WHERE       VEHICLEDATE >= TO_DATE('" & v_txtDateFrom & "', 'MM/DD/YYYY')" & _
			"                        AND VEHICLEDATE <= TO_DATE('" & v_txtDateTo & "', 'MM/DD/YYYY')" & _
			"                        AND JOBCODE = '6502204'" & _
			"                        AND LOCATIONTYPE = 'FF'" & _
			"               GROUP BY FCBA)" & _
			"     GROUP BY FCBA, NO_URUT, DELIVERY" & _
			"     ORDER BY FCBA, NO_URUT" & _
			" )" & _
			" where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E'" & _
			" GROUP BY FCBA, NO_URUT, DELIVERY ORDER BY FCBA, NO_URUT"
			
strAnalisaAngkutTBS2	=	"        , NO_URUT" & _
						"        , DELIVERY" & _
						"        , SUM (DIRECT_WEIGHT_KG) DIRECT_WEIGHT_KG" & _
						"        , SUM (INDIRECT_WEIGHT_KG) INDIRECT_WEIGHT_KG" & _
						"        , SUM (DIRECT_WEIGHT_KG+INDIRECT_WEIGHT_KG) TOTAL"
			
			
strAnalisaAngkutTBS	=	" select GroupPlantation " & selectCons & strAnalisaAngkutTBS2 & "  From (" & strMain & " )" & whereCons & _
						" group by GroupPlantation" & selectCons & " , NO_URUT, DELIVERY Order by GroupPlantation" & selectCons
						
response.write strAnalisaAngkutTBS
%>
<html>

<head>
<title>ePlaS Reporting - Analisa Angkut TBS</title>
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
				<!--#include virtual="/REPORT/FormSelectNewSumE.html"-->
				<table id="form-table" class="centered-margin">	
					<tr>
						<td><font face="verdana" size="1">DARI TANGGAL</font></td>
						<td align="center">
						<% if v_txtDateFrom <> "" and v_txtDateFrom = v_txtDateFrom then %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
						<% else %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% end if%>
							<!--<a href="javascript:show_calendar('frmEntry.txtDateFrom');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
								<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
							</a>-->
						</td>
						<td><font face="verdana" size="1">SAMPAI</font></td>
						<td align="center">
						<% if v_txtDateTo <> "" and v_txtDateTo = v_txtDateTo then %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=v_txtDateTo%>" title="Type Date mm/dd/yyyy" >
						<% else %>
							<input class="datepicker" type="text" maxlength="10" name="txtDateTo" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% end if%>
							<!--<a href="javascript:show_calendar('frmEntry.txtDateTo');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
								<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
							</a>-->
						</td>
						<td><input type="submit" value="View" /></td>
					</tr>				
				</table>
			</div>
			</form>
		</div>

<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
        <tr>
			<th colspan="<%=var_colspan+8%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">    
				<div id="header">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>ANALISA PENGANGKUTAN TBS</h3>
					<% if (v_txtdatefrom <> "" or v_txtdateto <> "") and v_txtdatefrom = v_txtdateto then%>
						<h5>TANGGAL: <%=v_txtdateto%></h5>
					<% else %>
						<h5>TANGGAL: <%=v_txtdatefrom%> - <%=v_txtdateto%></h5>
					<% end if %>
				</div>
			</th>
		</tr>
	<tr>
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
	
		<th rowspan="2" style="background-color:#00A300;color:#fff;">PENGANGKUTAN</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">LANGSUNG</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">TIDAK LANGSUNG</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">% DOUBLE HANDLING</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">KG</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">KG</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">KG</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
	</tr>
</thead>
<tbody>
<%
Set rsAnalisaAngkutTBS = Server.CreateObject("ADODB.Recordset")
rsAnalisaAngkutTBS.Open strAnalisaAngkutTBS, connstr, adOpenStatic

Dim rsGetTotal, strGetTotal
strGetTotal = strAnalisaAngkutTBS
Set rsGetTotal = Server.CreateObject("ADODB.Recordset")
rsGetTotal.Open strGetTotal, connstr, adOpenStatic

Dim Temp_TotalLangsung : Temp_TotalLangsung = 0
Dim Temp_TotalTidakLangsung : Temp_TotalTidakLangsung = 0
Dim Temp_TotalTotal : Temp_TotalTotal = 0

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

If Not rsGetTotal.EOF Then
	Do While Not rsGetTotal.EOF
		Temp_TotalLangsung = Temp_TotalLangsung + Cdbl(rsGetTotal("DIRECT_WEIGHT_KG"))
		Temp_TotalTidakLangsung = Temp_TotalTidakLangsung + Cdbl(rsGetTotal("INDIRECT_WEIGHT_KG"))
		Temp_TotalTotal = Temp_TotalTotal + Cdbl(rsGetTotal("TOTAL"))
		rsGetTotal.MoveNext
	Loop
End If
rsGetTotal.Close
set rsGetTotal = Nothing

if (rsAnalisaAngkutTBS.bof and rsAnalisaAngkutTBS.eof) then
%>
	<tr>
		<td colspan="<%=var_colspan+8%>">NO DATA AVAILABLE</td>
	</tr>
<%else

Dim TotalLangsung : TotalLangsung = 0
Dim TotalTidakLangsung : TotalTidakLangsung = 0
Dim TotalTotal : TotalTotal = 0
		while not rsAnalisaAngkutTBS.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
			
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsAnalisaAngkutTBS("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsAnalisaAngkutTBS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsAnalisaAngkutTBS("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("region")%>
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
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsAnalisaAngkutTBS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsAnalisaAngkutTBS("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsAnalisaAngkutTBS("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
							<%=rsAnalisaAngkutTBS("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsAnalisaAngkutTBS("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsAnalisaAngkutTBS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsAnalisaAngkutTBS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsAnalisaAngkutTBS("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsAnalisaAngkutTBS("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsAnalisaAngkutTBS("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtDateFrom=<%=v_txtDateFrom%>&txtDateTo=<%=v_txtDateTo%>">
						    <%=rsAnalisaAngkutTBS("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<td align="left"><%=rsAnalisaAngkutTBS("DELIVERY")%></td>
				
				<td align="right"><%=FormatNumber(rsAnalisaAngkutTBS("DIRECT_WEIGHT_KG"), 0)%></td>
				<%If Temp_TotalLangsung = 0 Then%>
				<td align="right"><%=FormatNumber(0, 0)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("DIRECT_WEIGHT_KG"))/Temp_TotalLangsung)*100, 0), 0)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsAnalisaAngkutTBS("INDIRECT_WEIGHT_KG"), 0)%></td>
				<%If Temp_TotalTidakLangsung = 0 Then%>
				<td align="right"><%=FormatNumber(0, 0)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("INDIRECT_WEIGHT_KG"))/Temp_TotalTidakLangsung)*100, 0), 0)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsAnalisaAngkutTBS("TOTAL"), 0)%></td>
				<%If Temp_TotalTotal = 0 Then%>
				<td align="right"><%=FormatNumber(0, 0)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Round((Cdbl(rsAnalisaAngkutTBS("TOTAL"))/Temp_TotalTotal)*100, 0), 0)%></td>
				<%End If%>
				
				<%If Cdbl(rsAnalisaAngkutTBS("DIRECT_WEIGHT_KG")) = 0 Then%>
				<td align="right"><%=FormatNumber(0, 0)%></td>
				<%Else%>
				<td align="right"><%=FormatNumber(Cdbl(rsAnalisaAngkutTBS("INDIRECT_WEIGHT_KG"))/Cdbl(rsAnalisaAngkutTBS("DIRECT_WEIGHT_KG"))*100, 0)%></td>
				<%End If%>
			</tr>
<%
			TotalLangsung = TotalLangsung + Cdbl(rsAnalisaAngkutTBS("DIRECT_WEIGHT_KG"))
			TotalTidakLangsung = TotalTidakLangsung + Cdbl(rsAnalisaAngkutTBS("INDIRECT_WEIGHT_KG"))
			TotalTotal = TotalTotal + Cdbl(rsAnalisaAngkutTBS("TOTAL"))
			rsAnalisaAngkutTBS.MoveNext
		Wend
	rsAnalisaAngkutTBS.Close
	set rsAnalisaAngkutTBS = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%=1+var_colspan%>"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=FormatNumber(TotalLangsung, 0)%></strong></td>
				<%If TotalLangsung = 0 Then%>
				<td align="right"><strong><%=FormatNumber(0, 0)%></strong></td>
				<%Else%>
				<td align="right"><strong><%=FormatNumber(TotalLangsung/TotalLangsung*100, 0)%></strong></td>
				<%End if%>
				<td align="right"><strong><%=FormatNumber(TotalTidakLangsung, 0)%></strong></td>
				<%If TotalTidakLangsung = 0 Then%>
				<td align="right"><strong><%=FormatNumber(0, 0)%></strong></td>
				<%Else%>
				<td align="right"><strong><%=FormatNumber(TotalTidakLangsung/TotalTidakLangsung*100, 0)%></strong></td>
				<%End if%>				
				<td align="right"><strong><%=FormatNumber(TotalTotal, 0)%></strong></td>
				<%If TotalTotal = 0 Then%>
				<td align="right"><strong><%=FormatNumber(0, 0)%></strong></td>
				<%Else%>
				<td align="right"><strong><%=FormatNumber(TotalTotal/TotalTotal*100, 0)%></strong></td>
				<%End if%>
				<%If TotalLangsung = 0 Then%>
				<td align="right"><strong><%=FormatNumber(0, 0)%></strong></td>
				<%Else%>
				<td align="right"><strong><%=FormatNumber(TotalTidakLangsung/TotalLangsung*100, 0)%></strong></td>
				<%End if%>				
			</tr>
<%End if%>
</tbody>
</table>
		</div>
	</body>
</html>