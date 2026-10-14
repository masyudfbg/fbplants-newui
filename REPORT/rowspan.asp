<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtDateFrom, v_txtDateTo
Dim ScriptName
Dim qdata, strAnalisaAngkutTBS, strAnalisaAngkutTBS2, strMain
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

v_txtDateFrom	=	Request("txtDateFrom")
v_txtDateTo		=	Request("txtDateTo")


If (v_txtDateFrom = "" Or Len(Trim(v_txtDateFrom)) = 0) Then
	v_txtDateFrom = textDate(Date)
End If

If (v_txtDateTo = "" Or Len(Trim(v_txtDateTo)) = 0) Then
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
			" ) GROUP BY FCBA, NO_URUT, DELIVERY ORDER BY FCBA, NO_URUT"
			
strAnalisaAngkutTBS2	=	"        , NO_URUT" & _
						"        , DELIVERY" & _
						"        , SUM (DIRECT_WEIGHT_KG) DIRECT_WEIGHT_KG" & _
						"        , SUM (INDIRECT_WEIGHT_KG) INDIRECT_WEIGHT_KG" & _
						"        , SUM (DIRECT_WEIGHT_KG+INDIRECT_WEIGHT_KG) TOTAL"
			
			
strAnalisaAngkutTBS	=	""
strAnalisaAngkutTBS	=	" SELECT" & _
						" DIVISION" & _
						" , TO_NUMBER(TO_CHAR(PLANTINGDATE,'YYYY')) TAHUN_TANAM" & _
						" , FCCODE" & _
						" FROM FIELD WHERE FCBA = 'KSB'" & _
						" ORDER BY DIVISION, TO_NUMBER(TO_CHAR(PLANTINGDATE,'YYYY')), FCCODE"

						
'response.write strAnalisaAngkutTBS
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Analisa Angkut TBS</title>
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

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">

<table border="0" width="100%" id="product-table">
	<tr>
		<th style="background-color:#00A300;color:#fff;">DIVISION</th>
		<th style="background-color:#00A300;color:#fff;">TAHUN TANAM</th>
		<th style="background-color:#00A300;color:#fff;">KODE</th>
	</tr>

<%
Set qdata = Server.CreateObject("ADODB.Recordset")
qdata.Open strAnalisaAngkutTBS, connstr, adOpenStatic

if (qdata.bof and qdata.eof) then
%>
	<tr>
		<td colspan="3">NO DATA AVAILABLE</td>
	</tr>
<%else
		while not qdata.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="left"><%=qdata("DIVISION")%></td>
				<td align="left"><%=qdata("TAHUN_TANAM")%></td>
				<td align="left"><%=qdata("FCCODE")%></td>
			</tr>
<%
			qdata.MoveNext
		Wend
	qdata.Close
	set qdata = Nothing
%>
<%End if%>
</table>
		</div>
	</body>
</html>