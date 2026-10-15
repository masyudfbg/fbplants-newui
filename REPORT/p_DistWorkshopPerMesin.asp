<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtTanggal, v_txtStation, v_txtMesin
Dim ScriptName
Dim rsDWPM, strDWPM, strDWPM2, strMain
Dim cClass, nday

Dim var_ColSpan : var_ColSpan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtMesin				= Request("txtMesin")

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
v_txtStation		=	Request("txtStation")


If (v_txtTanggal = "" Or Len(Trim(v_txtTanggal)) = 0) Then
	v_txtTanggal = textDate(Date)
End If

If v_txtStation = "" Then
	v_txtStation	=	"--ALL--"
End If


If Len(Trim(v_txtStation)) > 0 Then
	if v_txtStation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtStation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	whereCons	& " And STATIONCODE = '" & v_txtStation & "'"
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
	End If
End If

If Len(Trim(v_txtMesin)) > 0 Then
	if v_txtMesin = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtMesin = "--ALL--" then
		selectCons	=	selectCons	& " , MACHINECODE, MACHINENAME"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , MACHINECODE, MACHINENAME"
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , MACHINECODE, MACHINENAME"
		whereCons	=	whereCons	& " And MACHINECODE = '" & v_txtMesin & "'"
		orderCons	=	orderCons	& " , MACHINECODE, MACHINENAME"
		var_ColSpan	=	var_ColSpan	+ 1
	End If
End If

Dim v_Tahun : v_Tahun = Year(v_txtTanggal)
Dim v_Bulan : v_Bulan = Month(v_txtTanggal)

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT " & strFieldCons & ", " & _
			"          MECHINE_API.GET_STATIONCODE (MACHINECODE, FCBA) STATIONCODE," & _
			"          MECHINE_API.GET_STATIONNAME (MACHINECODE, FCBA) STATIONNAME," & _
			"          MACHINECODE," & _
			"          MECHINE_API.GET_MECHINENAME (MACHINECODE, FCBA) MACHINENAME," & _
			"          SUM (TOTALMACHINE) TOTALMACHINE," & _
			"          SUM (HI) HI," & _
			"          SUM (SHI) SHI," & _
			"          SUM (SBI) SBI" & _
			"     FROM (  SELECT FCBA," & _
			"                    MACHINECODE," & _
			"                    SUM (TOTALMACHINE) TOTALMACHINE," & _
			"                    SUM (HI) HI," & _
			"                    SUM (SHI) SHI," & _
			"                    SUM (SBI) SBI" & _
			"               FROM (SELECT FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            FCCODE MACHINECODE," & _
			"                            1 TOTALMACHINE," & _
			"                            0 HI," & _
			"                            0 SHI," & _
			"                            0 SBI" & _
			"                       FROM MACHINE" & _
			"                      WHERE ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            LOCATIONCODE MACHINECODE," & _
			"                            0 TOTALMACHINE," & _
			"                            NVL (UNIT, 0) HI," & _
			"                            0 SHI," & _
			"                            0 SBI" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MA'" & _
			"                            AND WORKSHOPDATE =" & _
			"                                   TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
			"                     UNION ALL" & _
			"                     SELECT FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            LOCATIONCODE MACHINECODE," & _
			"                            0 TOTALMACHINE," & _
			"                            0 HI," & _
			"                            NVL (UNIT, 0) SHI," & _
			"                            0 SBI" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE     LOCATIONTYPE = 'MA'" & _
			"                            AND TO_NUMBER (TO_CHAR (WORKSHOPDATE, 'YYYY')) = " & v_Tahun & "" & _
			"                            AND TO_NUMBER (TO_CHAR (WORKSHOPDATE, 'MM')) = " & v_Bulan & "" & _
			"                            AND WORKSHOPDATE <=" & _
			"                                   TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
			"                     UNION ALL" & _
			"                     SELECT FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            LOCATIONCODE MACHINECODE," & _
			"                            0 TOTALMACHINE," & _
			"                            0 HI," & _
			"                            0 SHI," & _
			"                            NVL (UNIT, 0) SBI" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE     LOCATIONTYPE = 'MA'" & _
			"                            AND TO_NUMBER (TO_CHAR (WORKSHOPDATE, 'YYYY')) = " & v_Tahun & "" & _
			"                            AND TO_NUMBER (TO_CHAR (WORKSHOPDATE, 'MM')) <= " & v_Bulan & ")" & _
			"           GROUP BY FCBA, MACHINECODE) where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' " & _
			" GROUP BY FCBA, MACHINECODE"

			
strDWPM2	=	"        , STATIONCODE" & _
				"        , STATIONNAME" & _
				"        , SUM (TOTALMACHINE) TOTALMACHINE" & _
				"        , SUM (HI) HI" & _
				"        , SUM (SHI) SHI" & _
				"        , SUM (SBI) SBI"
			
			
strDWPM	=	" select GroupPlantation " & selectCons & strDWPM2 & "  From (" & strMain & " )" & _
			whereCons & _
			" group by GroupPlantation" & selectCons & " , STATIONCODE, STATIONNAME" & _
			" Order by GroupPlantation" & orderCons &" , STATIONCODE"
						
'response.write strDWPM
%>
<html>

<head>
<title>ePlaS Reporting - Distribusi Kerja Workshop Per Mesin</title>
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
				<!--#include virtual="/REPORT/FormSelectNewSumM.html"-->
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
						<td><font face="verdana" size="1">STASIUN</font></td>
						<td>
							<select name="txtStation" >
								<option value="--ALL--"<%If v_txtStation = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsStation, strStation
								strStation	=	" SELECT DISTINCT FCCODE, FCNAME FROM STATION" & _
													" ORDER BY FCCODE"
								
								Set rsStation = Server.CreateObject("ADODB.Recordset")
								rsStation.Open strStation, connstr, adOpenStatic 
								
								if not rsStation.eof then
								Do While Not rsStation.eof
								%>

								<option value="<%=rsStation("FCCODE")%>" <% If v_txtStation = rsStation("FCCODE") Then %> Selected <%End if%>><%=rsStation("FCCODE")%> : <%=rsStation("FCNAME")%></option> 
								<%
									rsStation.movenext
								loop
								end if
								rsStation.Close
								set rsStation = Nothing
								%>
							</select>
						</td>
						<td><font face="verdana" size="1">MESIN</font></td>
						<td>
							<select name="txtMesin" >
								<option value="--SUMMARY--"<%If v_txtMesin = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtMesin = "--ALL--" then%> Selected <%End If%>>--ALL--</option>								
								<%
								Dim rsMesin, strMesin
								strMesin	=	" SELECT DISTINCT FCCODE, DESCRIPTION FROM MACHINE" & _
													" ORDER BY DESCRIPTION"
								
								Set rsMesin = Server.CreateObject("ADODB.Recordset")
								rsMesin.Open strMesin, connstr, adOpenStatic 
								
								if not rsMesin.eof then
								Do While Not rsMesin.eof
								%>

								<option value="<%=rsMesin("FCCODE")%>" <% If v_txtMesin = rsMesin("FCCODE") Then %> Selected <%End if%>><%=rsMesin("FCCODE")%> : <%=rsMesin("DESCRIPTION")%></option> 
								<%
									rsMesin.movenext
								loop
								end if
								rsMesin.Close
								set rsMesin = Nothing
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
          <th colspan="<%=var_ColSpan+8%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
			<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			</div>
			<div id="subheader" class="centered">
				<h2>DISTRIBUSI KERJA WORKSHOP PER MESIN</h2>
				<h5>TANGGAL: <%=v_txtTanggal%></h5>
			</div>
		</div>
           </th>
     </tr>
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">STASIUN</th>

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
		
		<%If Len(Trim(v_txtMesin)) > 0 And v_txtMesin <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;">MESIN</th>
	    <%End if%>
		
		<th rowspan="3" style="background-color:#00A300;color:#fff;">JUMLAH MESIN</th>
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
Set rsDWPM = Server.CreateObject("ADODB.Recordset")
rsDWPM.Open strDWPM, connstr, adOpenStatic

Dim rsGetTotal, strGetTotal
strGetTotal = " Select Nvl(Sum(HI), 0) HI, Nvl(Sum(SHI), 0) SHI, Nvl(Sum(SBI), 0) SBI FROM (" & strDWPM & ")"
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

Dim GrandTotal_Mesin : GrandTotal_Mesin = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempStationCode : tempStationCode = ""
Dim tempMachineCode : tempMachineCode = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim SprevUraian : SprevUraian = ""
Dim sprevMachineCode : sprevMachineCode = ""

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

if (rsDWPM.bof and rsDWPM.eof) then
%>
	<tr>
		<td colspan="<%=var_ColSpan+8%>">NO DATA AVAILABLE</td>
	</tr>
<%else
		while not rsDWPM.eof
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
				<% if v_txtStation <>  "--SUMMARY--" and  v_txtStation <>  "" then %>
					<% 						
					tempStationCode = rsDWPM("STATIONCODE")
					
					If Cint(visibleP) > 1 Then
						tempGroupPlantation = "--ALL--"
					Else
						If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then
							tempGroupPlantation = rsDWPM("GroupPlantation")
						Else
							tempGroupPlantation = ""
						End If
					End If
					
					If Cint(visibleGR) > 1 And Cint(visibleP) > 1 Then
						tempPlantation = "--SUMMARY--"
					Else
						If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then
							tempPlantation = rsDWPM("Plantation")
						Else
							tempPlantation = ""
						End If
					End If
					
					If Cint(visibleR) > 1 And Cint(visibleGR) > 1 Then
						tempGroupRegion = "--SUMMARY--"
					Else
						If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then
							tempGroupRegion = rsDWPM("GroupRegion")
						Else
							tempGroupRegion = ""
						End If
					End If
					
					If Cint(visibleC) > 1 And Cint(visibleR) > 1 Then
						tempRegion = "--SUMMARY--"
					Else
						If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then
							tempRegion = rsDWPM("Region")
						Else
							tempRegion = ""
						End If
					End If
					
					If Cint(visibleS) > 1 And Cint(visibleC) > 1 Then
						tempCompany = "--SUMMARY--"
					Else
						If Len(Trim(v_txtcompany)) > 0 And v_txtcompany <>  "--SUMMARY--" Then
							tempCompany = rsDWPM("Company")
						Else
							tempCompany = ""
						End If
					End If
					
					If Cint(VisibleS) > 1 Then
						tempfcba = "--SUMMARY--"
					Else
						If Len(Trim(v_txtFcba)) > 0 And v_txtFcba <>  "--SUMMARY--" Then
							tempfcba = rsDWPM("Fcba")
						Else
							tempfcba = ""
						End If
					End If
					
						tempMachineCode = "--SUMMARY--"
					%>
						<% if tempStationCode <> SprevUraian then %>
							<td align="left">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtStation=<%=tempStationCode%>&txtMesin=<%=tempMachineCode%>">
								<%=rsDWPM("STATIONNAME")%>
								</a>
							</td>
							<% SprevUraian = tempStationCode %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
				<% end if %>

				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtStation <>  "--SUMMARY--"  and  v_txtStation <> "" then
					      tempStationCode = rsDWPM("STATIONCODE")
					   else
					      tempStationCode = ""
					   end if
					   tempGroupPlantation = rsDWPM("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempMachineCode = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtStation=<%=tempStationCode%>&txtMesin=<%=tempMachineCode%>">
							<%=rsDWPM("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtStation <>  "--SUMMARY--"  and  v_txtStation <> "" then
					      tempStationCode = rsDWPM("STATIONCODE")
					   else
					      tempStationCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPM("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsDWPM("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempMachineCode = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtStation=<%=tempStationCode%>&txtMesin=<%=tempMachineCode%>">
							<%=rsDWPM("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtStation <>  "--SUMMARY--"  and  v_txtStation <> "" then
					      tempStationCode = rsDWPM("STATIONCODE")
					   else
					      tempStationCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPM("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDWPM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsDWPM("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempMachineCode = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtStation=<%=tempStationCode%>&txtMesin=<%=tempMachineCode%>">
							<%=rsDWPM("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtStation <>  "--SUMMARY--"  and  v_txtStation <> "" then
					      tempStationCode = rsDWPM("STATIONCODE")
					   else
					      tempStationCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPM("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDWPM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDWPM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsDWPM("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempMachineCode = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtStation=<%=tempStationCode%>&txtMesin=<%=tempMachineCode%>">
							<%=rsDWPM("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtStation <>  "--SUMMARY--"  and  v_txtStation <> "" then
					      tempStationCode = rsDWPM("STATIONCODE")
					   else
					      tempStationCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPM("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDWPM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDWPM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsDWPM("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsDWPM("Company")
					   tempfcba = "--ALL--"
					   tempMachineCode = "--SUMMARY--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtStation=<%=tempStationCode%>&txtMesin=<%=tempMachineCode%>">
							<%=rsDWPM("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtStation <>  "--SUMMARY--"  and  v_txtStation <> "" then
					      tempStationCode = rsDWPM("STATIONCODE")
					   else
					      tempStationCode = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDWPM("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDWPM("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDWPM("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsDWPM("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsDWPM("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsDWPM("fcba")
					   tempMachineCode = "--ALL--"					   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtStation=<%=tempStationCode%>&txtMesin=<%=tempMachineCode%>">
						    <%=rsDWPM("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>
					<% end if %>
				<% end if %>
				
				
				<% if v_txtMesin <> "--SUMMARY--" and v_txtMesin <>  ""  then %>
				    <% if sprevMachineCode <> rsDWPM("MACHINECODE") then %>				
					    <td align="left"><%=rsDWPM("MACHINENAME")%></td>				
						<% sprevMachineCode = rsDWPM("MACHINECODE")%>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
				
<%
			If Temp_Total_HI = 0 Then
				Total_Persen_HI	= 0
				GrandTotal_Persen_HI = 0
			Else
				Total_Persen_HI	= Round((Cdbl(rsDWPM("HI"))/Temp_Total_HI)*100, 2)
				GrandTotal_Persen_HI	=	GrandTotal_Persen_HI + (Cdbl(rsDWPM("HI"))/Temp_Total_HI)*100
			End If
			If Temp_Total_SHI = 0 Then
				Total_Persen_SHI = 0
				GrandTotal_Persen_SHI = 0
			Else
				Total_Persen_SHI = Round((Cdbl(rsDWPM("SHI"))/Temp_Total_SHI)*100, 2)
				GrandTotal_Persen_SHI	=	GrandTotal_Persen_SHI + (Cdbl(rsDWPM("SHI"))/Temp_Total_SHI)*100
			End If
			If Temp_Total_SBI = 0 Then
				Total_Persen_SBI = 0
				GrandTotal_Persen_SBI = 0
			Else
				Total_Persen_SBI = Round((Cdbl(rsDWPM("SBI"))/Temp_Total_SBI)*100, 2)
				GrandTotal_Persen_SBI	=	GrandTotal_Persen_SBI + (Cdbl(rsDWPM("SBI"))/Temp_Total_SBI)*100
			End If
%>
				
				<td align="right"><%=FbpNum2(rsDWPM("TOTALMACHINE"))%></td>
				
				<td align="right"><%=FormatNumber(rsDWPM("HI"), 2)%></td>				
				<td align="right"><%=FormatNumber(Total_Persen_HI, 2)%></td>
				
				<td align="right"><%=FormatNumber(rsDWPM("SHI"), 2)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_SHI, 2)%></td>

				<td align="right"><%=FormatNumber(rsDWPM("SBI"), 2)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_SBI, 2)%></td>
			</tr>
<%
			GrandTotal_HI	=	GrandTotal_HI + Cdbl(rsDWPM("HI"))
			GrandTotal_SHI	=	GrandTotal_SHI + Cdbl(rsDWPM("SHI"))
			GrandTotal_SBI	=	GrandTotal_HI + Cdbl(rsDWPM("SBI"))
			
			GrandTotal_Mesin = 	GrandTotal_Mesin + Cdbl(rsDWPM("TOTALMACHINE"))
			
			rsDWPM.MoveNext
		Wend
	rsDWPM.Close
	set rsDWPM = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td align="center" colspan="<%=var_ColSpan%>"><strong>TOTAL</strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Mesin, 2)%></strong></td>
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