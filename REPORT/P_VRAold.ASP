<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Laporan Vehicle Running Account (VRA)</title>
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
</head>

<body>
	<div id="wrapper">
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<!--#include file="FromSelectHDRVRA.html"-->
	
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
		<div id="subheader" class="centered">
			<% Dim ia
			   ia = request.Form("vMonth")
			%>   
			<h3>LAPORAN VEHICLE RUNNING ACCOUNT (VRA)</h3>
			<% IF request.Form("vMonth") <> "" then %>
			<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=request.Form("vYear")%></h5>
			<% END IF %>
			
		</div>
	</div>


<table border="0" width="100%" id="product-table">
	<tr>
		<th  style="color:#FFFFFF" rowspan="3">KODE</th>
		<th  style="color:#FFFFFF" rowspan="3">NAMA KENDARAAN</th>
		<th  style="color:#FFFFFF" rowspan="3">TAHUN PEROLEHAN</th>
		<th  style="color:#FFFFFF" colspan="4">AKTUAL PEMAKAIAN FISIK</th>
		<th  style="color:#FFFFFF" rowspan="3">RASIO S/D B.I (KM/LTR)</th>
		<th  style="color:#FFFFFF" rowspan="3">BUDGET SETAHUN (KM)</th>
		<th  style="color:#FFFFFF" colspan="5">AKTUAL BIAYA OPERASI BULAN INI</th>
		<th  style="color:#FFFFFF" colspan="5">AKTUAL BIAYA OPERASI S/D BULAN INI</th>
		<th  style="color:#FFFFFF" rowspan="3">BUDGET SETAHUN (Rp)</th>
		<th  style="color:#FFFFFF" colspan="3">RP/UNIT(KM/HM)  Rp)</th>
	</tr>
	<tr>
		<th  style="color:#FFFFFF" colspan="2">BI</th>
		<th  style="color:#FFFFFF" colspan="2">s/d BI</th>
		
		<th  style="color:#FFFFFF" rowspan="2">GAJI SOPIR</th>
		<th  style="color:#FFFFFF" rowspan="2">BBM & PELUMAS</th>
		<th  style="color:#FFFFFF" rowspan="2">SUKU CADANG</th>
		<th  style="color:#FFFFFF" rowspan="2">REPARASI & LAIN-LAIN</th>
		<th  style="color:#FFFFFF" rowspan="2">TOTAL</th>
		
		<th  style="color:#FFFFFF" rowspan="2">GAJI SOPIR</th>
		<th  style="color:#FFFFFF" rowspan="2">BBM & PELUMAS</th>
		<th  style="color:#FFFFFF" rowspan="2">SUKU CADANG</th>
		<th  style="color:#FFFFFF" rowspan="2">REPARASI & LAIN-LAIN</th>
		<th  style="color:#FFFFFF" rowspan="2">TOTAL</th>
		
		<th  style="color:#FFFFFF" colspan="2">AKTUAL</th>
		<th  style="color:#FFFFFF" rowspan="2">BUDGET SETAHUN</th>
	</tr>

	<tr>
		
		<th style="color:#FFFFFF" >(KM/HM)</th>
		<th style="color:#FFFFFF" >LTR</th>
		
		<th style="color:#FFFFFF" >(KM/HM)</th>
		<th style="color:#FFFFFF" >LTR</th>
		
		<th style="color:#FFFFFF" >BI</th>
		<th style="color:#FFFFFF" >S/D BI</th>
	</tr>
	
	<%
	Dim rsha, strdtl ,actualBI,actualSBI,totactualBI,totactualSBI,rasioBI, strgrp ,cdgrp,cdgrpdesc  ,noline , subKMHMBI,subKMHMSBI,subltrBI,subltrSBI,subSALBI,subSALSBI,subbbmBI,subbbmSBI,subspartBI,subspartSBI,subrepBI,subrepSBI 
	Dim subrasioBI,subactualBI ,subactualSBI ,subtotactualBI ,subtotactualSBI ,subKMHMbudTHN,subAMTbudTHN
	Dim TotsubKMHMBI,TotsubKMHMSBI,TotsubltrBI,TotsubltrSBI,TotsubSALBI,TotsubSALSBI,TotsubbbmBI,TotsubbbmSBI,TotsubspartBI,TotsubspartSBI,TotsubrepBI,TotsubrepSBI 
	Dim TotsubrasioBI,TotsubactualBI ,TotsubactualSBI ,TotsubtotactualBI ,TotsubtotactualSBI ,TotsubKMHMbudTHN,TotsubAMTbudTHN
	
	
			strdtl = " select vehiclecode,vehicle_name,VEHICLE_API.get_vehiclegroupcode(vehiclecode,fcba) grpvehicle,VEHICLEGROUP_API.get_vehiclegroup_name(VEHICLE_API.get_vehiclegroupcode(vehiclecode,fcba),fcba) grpvehiclename,vehicle_api.get_vehicle_year(vehiclecode,fcba) vehicleYear,sum(KMHMBI) KMHMBI,sum(KMHMSBI) KMHMSBI,sum(ltrBI) ltrBI,sum(ltrSBI) ltrSBI,sum(SALBI) SALBI, sum(SALSBI) SALSBI,sum(BBMBI) bbmBI,sum(BBMSBI) bbmSBI,sum(spartBI) spartBI,sum(spartSBI) spartSBI,sum(repBI) repBI,sum(repSBI) repSBI,sum(KMHMbudTHN) KMHMbudTHN,sum(AmtbudTHN) AmtbudTHN  from ( " & _
					 "	select va.fcba,va.vehiclecode,vehicle_api.get_vehiclecode_name(va.vehiclecode,va.fcba) vehicle_name,sum(va.unit) KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  from VEHICLEACTIVITY va  " & _
					 "	where  to_char(va.vehicledate,'YYYY')= '" & request.Form("vYear") & "'  and to_char(va.vehicledate,'MM') =  lpad('" & request.Form("vMonth") & "',2,0)  and va.fcba='" & session("fcba") & "' " & _
					 "	group by va.fcba,va.vehiclecode " & _
					 "	union " & _
					 "	select va.fcba,va.vehiclecode,vehicle_api.get_vehiclecode_name(va.vehiclecode,va.fcba) vehicle_name,0 KMHMBI,sum(va.unit) KMHMSBI,0 ltrBI,0 ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  from VEHICLEACTIVITY va  " & _
					 "	where  to_char(va.vehicledate,'YYYY')= '" & request.Form("vYear") & "' and to_char(va.vehicledate,'MM') <=  lpad('" & request.Form("vMonth") & "',2,0)  and va.fcba='" & session("fcba") & "' " & _
					 "	group by va.fcba,va.vehiclecode " & _
					 "	union " & _
					 "	select isd.fcba,isd.locationcode vehiclecode,vehicle_api.get_vehiclecode_name(isd.locationcode,isd.fcba) vehicle_name,0 KMHMBI,0 KMHMSBI,sum(isd.quantity) ltrBI , 0 ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN   from issuedetails isd  " & _
					 "	where isd.locationtype='VH' and to_char(get_issuedate(isd.sivcode,isd.fcba),'YYYY')= '" & request.Form("vYear") & "' and to_char(get_issuedate(isd.sivcode,isd.fcba),'MM') =  lpad('" & request.Form("vMonth") & "',2,0)  and isd.fcba='" & session("fcba") & "' " & _
					 "	group by isd.fcba,isd.locationcode " & _
					 "	union " & _
					 "	select isd.fcba,isd.locationcode vehiclecode,vehicle_api.get_vehiclecode_name(isd.locationcode,isd.fcba) vehicle_name,0 KMHMBI,0 KMHMSBI, 0 ltrBI,sum(isd.quantity) ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN    from issuedetails isd  " & _
					 "	where isd.locationtype='VH'  and to_char(get_issuedate(isd.sivcode,isd.fcba),'YYYY')= '" & request.Form("vYear") & "' and to_char(get_issuedate(isd.sivcode,isd.fcba),'MM') <=  lpad('" & request.Form("vMonth") & "',2,0) and isd.fcba='" & session("fcba") & "' " & _
					 "	group by isd.fcba,isd.locationcode " & _
					 "	union " & _
					 "	select c.fcba,c.targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,sum(c.amount) SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  from cost_activity c " & _ 
					 "	where c.targettype='VH'   and to_char(c.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(c.tdate,'MM')= lpad('" & request.Form("vMonth") & "',2,0) AND  c.FCBA='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6310100','6310200','6310300','6310400','6310500','6310600','6310700','6310800','6310900','6311000') " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	union " & _
					 "	select c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,sum(c.amount) SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  from cost_activity c " & _
					 "	where c.targettype='VH'   and to_char(c.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(c.tdate,'MM')<= lpad('" & request.Form("vMonth") & "',2,0) AND  c.FCBA='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6310100','6310200','6310300','6310400','6310500','6310600','6310700','6310800','6310900','6311000') " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	union " & _
					 "	select c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,sum(c.amount) bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  from cost_activity c " & _
					 "	where c.targettype='VH'   and to_char(c.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(c.tdate,'MM')= lpad('" & request.Form("vMonth") & "',2,0) AND  c.FCBA='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6311100') " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	union " & _
					 "	select c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,sum(c.amount) bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN   from cost_activity c " & _
					 "	where c.targettype='VH' and to_char(c.tdate,'YYYY')= '" & request.Form("vYear") & "'  and to_char(c.tdate,'MM')<= lpad('" & request.Form("vMonth") & "',2,0) AND  c.FCBA='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6311100') " & _
					 "	group by c.fcba,c.targetcode  " & _
					 "	union " & _
					 "	select  c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,sum(c.amount) spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  from cost_activity c " & _
					 "	where c.targettype='VH'  and to_char(c.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(c.tdate,'MM')= lpad('" & request.Form("vMonth") & "',2,0) AND  c.FCBA='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6311200') " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	union " & _
					 "	select  c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,sum(c.amount) spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  from cost_activity c " & _
					 "	where c.targettype='VH' and to_char(c.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(c.tdate,'MM')<= lpad('" & request.Form("vMonth") & "',2,0)  AND  c.FCBA='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6311200') " & _
					 "	group by c.fcba,c.targetcode  " & _
					 "	union " & _
					 "	select  c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,sum(c.amount) repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  from cost_activity c " & _
					 "	where c.targettype='VH' and to_char(c.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(c.tdate,'MM')= lpad('" & request.Form("vMonth") & "',2,0)  AND  c.FCBA='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6311300','6311400') " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	union " & _
					 "	select  c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,sum(c.amount) repSBI,0 KMHMbudTHN,0 AmtbudTHN   from cost_activity c " & _
					 "	where c.targettype='VH' and to_char(c.tdate,'YYYY')= '" & request.Form("vYear") & "' and to_char(c.tdate,'MM')<= lpad('" & request.Form("vMonth") & "',2,0)  AND  c.FCBA='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6311300','6311400') " & _
					 "	group by c.fcba,c.targetcode  " & _
					 "	union " & _
					 "	select  c.fcba,c.vehiclecode,vehicle_api.get_vehiclecode_name(c.vehiclecode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,sum(c.qty_01+c.qty_02+c.qty_03+c.qty_04+c.qty_05+c.qty_06+c.qty_07+c.qty_08+c.qty_09+c.qty_10+c.qty_11+c.qty_12) KMHMbudTHN,0 AmtbudTHN   from BUDGET_TRANSACTION_VHHM c " & _
					 "	where c.budget_year ='" & request.Form("vYear") & "'  and c.fcba='" & session("fcba") & "' " & _
					 "	group by c.fcba,c.vehiclecode " & _
					 "	union " & _
					 "	select  c.fcba,c.vehiclecode,vehicle_api.get_vehiclecode_name(c.vehiclecode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,sum(c.budget_01+c.budget_02+c.budget_03+c.budget_04+c.budget_05+c.budget_06+c.budget_07+c.budget_08+c.budget_09+c.budget_10+c.budget_11+c.budget_12) AmtbudTHN  from BUDGET_TRANSACTION_VEHICLE c " & _
					 "	where c.budget_year='" & request.Form("vYear") & "' and c.fcba='" & session("fcba") & "' " & _
					 "	and c.jobcode in ('6310100','6310200','6310300','6310400','6310500','6310600','6310700','6310800','6310900','6311000','6311100','6311200','6311300','6311400') " & _
					 " 	group by c.fcba,c.vehiclecode " & _
					 "	) group by fcba,vehiclecode,vehicle_name  order by VEHICLE_API.get_vehiclegroupcode(vehiclecode,fcba),vehiclecode "	
			
	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
   ' response.write strdtl
   ' response.write request.Form("vMonth") &  "  " & request.Form("vYear") & " " & session("fcba") 
	
	if Not rsha.eof then
		  while Not rsha.eof
		    IF  cdgrp = rsha("grpvehicle") THEN
			    %>
				<tr>  
					<td align="CENTER"><%=rsha("vehiclecode")%></td>
					<td align="LEFT"><%=rsha("vehicle_name")%></td>
					<td align="center"><%=rsha("vehicleYear")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("KMHMBI")),0)%></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("LTRBI")),0)%></td></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("KMHMSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("LTRSBI")),0)%></td></td>
					<% if cdbl(rsha("LTRBI")) > 0 and cdbl(rsha("KMHMBI")) > 0 then
					        rasioBI = cdbl(rsha("LTRBI"))/ cdbl(rsha("KMHMBI")) 
					   else 
					        rasioBI ="0"
				       end if			%>
					<td align="right"><%=formatnumber( cdbl(rasioBI),2)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("KMHMbudTHN")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("SALBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("BBMBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("spartBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("REPBI")),0)%></td>
					<% actualBI = cdbl(rsha("SALBI")) + cdbl(rsha("BBMBI")) + cdbl(rsha("spartBI")) + cdbl(rsha("REPBI"))  %>
					<td align="right"><%=formatnumber( cdbl(actualBI),0)%></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("SALSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("BBMSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("spartSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("REPSBI")),0)%></td>
					<% actualSBI = cdbl(rsha("SALSBI")) + cdbl(rsha("BBMSBI")) + cdbl(rsha("spartSBI")) + cdbl(rsha("REPSBI"))  %>
					<td align="right"><%=formatnumber( cdbl(actualSBI),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("AMTbudTHN")),0)%></td>
					<% if cdbl(rsha("LTRBI")) > 0 then
					        totactualBI = cdbl(actualBI)/ cdbl(rsha("LTRBI")) 
					   else 
					        totactualBI  ="0"
				       end if	
					   if cdbl(rsha("LTRSBI")) > 0 then
					        totactualSBI = cdbl(actualSBI)/ cdbl(rsha("LTRSBI")) 
					   else 
					        totactualSBI  ="0"
				       end if
					   %>
					<td align="right"><%=formatnumber( cdbl(totactualBI),0)%></td>
					<td align="right"><%=formatnumber( cdbl(totactualSBI),0)%></td>
					<% IF cdbl(rsha("KMHMbudTHN")) > 0 THEN %>
					  <td align="right"><%=formatnumber( cdbl(rsha("AMTbudTHN"))/cdbl(rsha("KMHMbudTHN")),0)%></td>
					<% Else %>
					   <td align="right"><%=formatnumber( cdbl(rsha("AMTbudTHN")),0)%></td>
					<% End If %>
				<tr>	
				  <%
					  subKMHMBI= cdbl(subKMHMBI) + cdbl(rsha("KMHMBI"))
					  subKMHMSBI= cdbl(subKMHMSBI) + cdbl(rsha("KMHMSBI"))
					  subltrBI= cdbl(subltrBI) + cdbl(rsha("LTRBI"))
					  subltrSBI= cdbl(subltrSBI) + cdbl(rsha("LTRSBI"))
					  subSALBI= cdbl(subSALBI) + cdbl(rsha("SALBI"))
					  subSALSBI= cdbl(subSALSBI) + cdbl(rsha("SALSBI"))
					  subbbmBI= cdbl(subbbmBI) +  cdbl(rsha("bbmBI"))
					  subbbmSBI= cdbl(subbbmSBI) + cdbl(rsha("bbmSBI"))
					  subspartBI= cdbl(subspartBI) + cdbl(rsha("spartBI"))
					  subspartSBI= cdbl(subspartSBI) + cdbl(rsha("spartSBI"))
					  subrepBI= cdbl(subrepBI) + cdbl(rsha("repBI"))
					  subrepSBI= cdbl(subrepSBI) + cdbl(rsha("repSBI"))
					  subKMHMbudTHN = cdbl(subKMHMbudTHN) + cdbl(rsha("KMHMbudTHN"))
					  subAMTbudTHN = cdbl(subAMTbudTHN) + cdbl(rsha("AMTbudTHN"))
					  subrasioBI =  cdbl(subrasioBI) + cdbl(rasioBI)
					  subactualBI =  cdbl(subactualBI) +  cdbl(actualBI)
					  subactualSBI =  cdbl(subactualSBI) + cdbl(actualSBI)
					  subtotactualBI = cdbl(subtotactualBI) + cdbl(totactualBI)
					  subtotactualSBI = cdbl(subtotactualSBI) + cdbl(totactualSBI)
			ELSE	  
			        if  cdgrp <> "" then  %>
					    <tr>
						   <td style="background-color:#F4F4A0" align="center" colspan="3">SUB TOTAL&nbsp;&nbsp;<%=cdgrpdesc%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subKMHMBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subltrBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subKMHMSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subltrSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subrasioBI),2)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subKMHMbudTHN),2)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subSALBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subbbmBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subspartBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subrepBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subactualBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subSALSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subbbmSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subspartSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subrepSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subactualSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subAMTbudTHN),2)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subtotactualBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subtotactualSBI),0)%></td>
						   <% if cdbl(subKMHMbudTHN) > 0 then %>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subAMTbudTHN)/cdbl(subKMHMbudTHN),0)%></td>  
						   <% Else %>
						   <td style="background-color:#F4F4A0"  align="right"><%=formatnumber( cdbl(subAMTbudTHN),0)%></td> 
						   <% End if %> 
					   </tr> 
				<%	end if  %>  
				
				<tr>  
					<td align="CENTER"><%=rsha("vehiclecode")%></td>
					<td align="LEFT"><%=rsha("vehicle_name")%></td>
					<td align="center"><%=rsha("vehicleYear")%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("KMHMBI")),0)%></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("LTRBI")),0)%></td></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("KMHMSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("LTRSBI")),0)%></td></td>
					<% if cdbl(rsha("LTRBI")) > 0 and cdbl(rsha("KMHMBI")) > 0 then
					        rasioBI = cdbl(rsha("LTRBI"))/ cdbl(rsha("KMHMBI")) 
					   else 
					        rasioBI ="0"
				       end if			
					%>
					<td align="right"><%=formatnumber( cdbl(rasioBI),2)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("KMHMbudTHN")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("SALBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("BBMBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("spartBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("REPBI")),0)%></td>
					<% actualBI = cdbl(rsha("SALBI")) + cdbl(rsha("BBMBI")) + cdbl(rsha("spartBI")) + cdbl(rsha("REPBI"))  %>
					<td align="right"><%=formatnumber( cdbl(actualBI),0)%></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("SALSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("BBMSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("spartSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("REPSBI")),0)%></td>
					<% actualSBI = cdbl(rsha("SALSBI")) + cdbl(rsha("BBMSBI")) + cdbl(rsha("spartSBI")) + cdbl(rsha("REPSBI"))  %>
					<td align="right"><%=formatnumber( cdbl(actualSBI),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("AMTbudTHN")),0)%></td>
					<% if cdbl(rsha("LTRBI")) > 0 then
					        totactualBI = cdbl(actualBI)/ cdbl(rsha("LTRBI")) 
					   else 
					        totactualBI  ="0"
				       end if	
					   if cdbl(rsha("LTRSBI")) > 0 then
					        totactualSBI = cdbl(actualSBI)/ cdbl(rsha("LTRSBI")) 
					   else 
					        totactualSBI  ="0"
				       end if
					%>
					<td align="right"><%=formatnumber( cdbl(totactualBI),0)%></td>
					<td align="right"><%=formatnumber( cdbl(totactualSBI),0)%></td>
					<% IF cdbl(rsha("KMHMbudTHN")) > 0 THEN %>
					  <td align="right"><%=formatnumber( cdbl(rsha("AMTbudTHN"))/cdbl(rsha("KMHMbudTHN")),0)%></td>
					<% Else %>
					   <td align="right"><%=formatnumber( cdbl(rsha("AMTbudTHN")),0)%></td>
					<% End If %>
					
				<tr>	
				  <%
					  subKMHMBI=  cdbl(rsha("KMHMBI"))
					  subKMHMSBI=  cdbl(rsha("KMHMSBI"))
					  subltrBI=  cdbl(rsha("LTRBI"))
					  subltrSBI= cdbl(rsha("LTRSBI"))
					  subSALBI=  cdbl(rsha("SALBI"))
					  subSALSBI=  cdbl(rsha("SALSBI"))
					  subbbmBI=  cdbl(rsha("bbmBI"))
					  subbbmSBI=  cdbl(rsha("bbmSBI"))
					  subspartBI=  cdbl(rsha("spartBI"))
					  subspartSBI=  cdbl(rsha("spartSBI"))
					  subrepBI=  cdbl(rsha("repBI"))
					  subrepSBI=  cdbl(rsha("repSBI"))
					  subKMHMbudTHN = cdbl(rsha("KMHMbudTHN"))
					  subAMTbudTHN = cdbl(rsha("AMTbudTHN"))
					  subrasioBI =   cdbl(rasioBI)
					  subactualBI =    cdbl(actualBI)
					  subactualSBI =   cdbl(actualSBI)
					  subtotactualBI =  cdbl(totactualBI)
					  subtotactualSBI =  cdbl(totactualSBI)
			END IF	  %>
				
				
    <%				  	  
			  TotSubKMHMBI= cdbl(TotSubKMHMBI) + cdbl(rsha("KMHMBI"))
			  TotSubKMHMSBI= cdbl(TotSubKMHMSBI) + cdbl(rsha("KMHMSBI"))
			  TotSubltrBI= cdbl(TotSubltrBI) + cdbl(rsha("LTRBI"))
			  TotSubltrSBI= cdbl(TotSubltrSBI) + cdbl(rsha("LTRSBI"))
			  TotSubSALBI= cdbl(TotSubSALBI) + cdbl(rsha("SALBI"))
			  TotSubSALSBI= cdbl(TotSubSALSBI) + cdbl(rsha("SALSBI"))
			  TotSubbbmBI= cdbl(TotSubbbmBI) +  cdbl(rsha("bbmBI"))
			  TotSubbbmSBI= cdbl(TotSubbbmSBI) + cdbl(rsha("bbmSBI"))
			  TotSubspartBI= cdbl(TotSubspartBI) + cdbl(rsha("spartBI"))
			  TotSubspartSBI= cdbl(TotSubspartSBI) + cdbl(rsha("spartSBI"))
			  TotSubrepBI= cdbl(TotSubrepBI) + cdbl(rsha("repBI"))
			  TotSubrepSBI= cdbl(TotSubrepSBI) + cdbl(rsha("repSBI"))
			  TotsubKMHMbudTHN = cdbl(TotsubKMHMbudTHN) + cdbl(rsha("KMHMbudTHN"))
			  TotsubAMTbudTHN = cdbl(TotsubAMTbudTHN) + cdbl(rsha("AMTbudTHN")) 		  
			  TotSubrasioBI =  cdbl(TotSubrasioBI) + cdbl(rasioBI)
			  TotSubactualBI =  cdbl(TotSubactualBI) +  cdbl(actualBI)
			  TotSubactualSBI =  cdbl(TotSubactualSBI) + cdbl(actualSBI)
			  TotSubtotactualBI = cdbl(TotSubtotactualBI) + cdbl(totactualBI)
			  TotSubtotactualSBI = cdbl(TotSubtotactualSBI) + cdbl(totactualSBI)
			  
			  cdgrp = rsha("grpvehicle")
			  cdgrpdesc = rsha("grpvehiclename")
			  rsha.MoveNext
		  wend
	end if
	%>	
					    <tr>
						   <td style="background-color:#F4F4A0" align="center" colspan="3">SUB TOTAL&nbsp;&nbsp;<%=cdgrpdesc%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subKMHMBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subltrBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subKMHMSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subltrSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrasioBI),2)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotsubKMHMbudTHN),2)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subSALBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subbbmBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subspartBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrepBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subactualBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subSALSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subbbmSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subspartSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrepSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subactualSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subAMTbudTHN),2)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotactualBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotactualSBI),0)%></td>
						   <% if cdbl(subKMHMbudTHN) > 0 then %>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subAMTbudTHN)/cdbl(subKMHMbudTHN),0)%></td>  
						   <% Else %>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subAMTbudTHN),0)%></td> 
						   <% End if %>   
					   </tr> 
					   <tr>
						   <td  colspan="23"></td>
					   </tr> 
	
   <tr>
       <td style="background-color:#F4F4A0" align="center" colspan="3"><b>GRAND TOTAL</td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubKMHMBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubltrBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubKMHMSBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubltrSBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubrasioBI),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotsubKMHMbudTHN),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubSALBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubbbmBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubspartBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubrepBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubactualBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubSALSBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubbbmSBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubspartSBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubrepSBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubactualSBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotsubAMTbudTHN),2)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubtotactualBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubtotactualSBI),0)%></td>
	   <% if cdbl(TotsubKMHMbudTHN) > 0 then %>
			<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotsubAMTbudTHN)/cdbl(TotsubKMHMbudTHN),0)%></td>  
	  <% Else %>
		   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotsubAMTbudTHN),0)%></td> 
	  <% End if %> 
   </tr>    
</table>
<!--#include file="bottom.Asp"-->