<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8, val_vehiclegrouptype
Dim ScriptName,nSpanTotal
Dim vyear,vmonth,v_vJenisReport,v_txtVH,v_TxtTypeFCBA 
Dim SprevCode,Sprevjobname
Dim sprevFcba, SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,SprevVH
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempVH
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,rsJob
            v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	  = Request("txtPlantation")
			v_txtGroupRegion  	  = Request("txtGroupRegion")
			v_txtRegion  		  = Request("txtRegion")
			v_txtcompany  		  = Request("txtCompany")
			v_txtfcba 			  = Request("txtfcba")
			vyear             	  = Request("vYear")
			vmonth            	  = Request("vMonth")
			v_TxtTypeFCBA         = Request("TxtTypeFCBA")
			v_txtVH				  = Request("txtVH")
            nSpanTotal = 5 
			
			 Dim selectFields,strFieldCons,selectCons,whereCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, BUSINESSUNIT_API.get_OrderData(fcba) norut, fcba, " 

						   
			wherecons = " where company is not null "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' " 
			end if
			
			
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' " 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE FROM BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   CompanyName=rsCat("FCCOMPANYNAME")			   
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
	        if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",norut,fcba"
			    orderCons  = orderCons & ",norut "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE FROM BusinessUnit where fcCode='" & v_txtfcba & "' "
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   fcbaname = rsCat("fcName")
				   CompanyName=rsCat("FCCOMPANYNAME")
				   MgrName = rsCat("FCMANAGER")
				   KTUName = rsCat("FCKTU")
				   CropName=rsCat("FCCROP")
				   FcbaType=rsCat("fcType")
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
	        if v_txtVH	<> "--SUMMARY--" and v_txtVH <> "" then
			   if v_txtVH	<> "--ALL--"  then
			      whereCons = wherecons & " And vehiclecode = '" & v_txtVH & "' "
			   end if
		
			  orderCons  = orderCons & ",grpvehicle,vehiclecode  "
			else
			    nSpanTotal = nSpanTotal - 2 
			end if 
%>
<html>

<head>
<title>ePlaS Reporting - Laporan Vehicle Running Account (VRA)</title>
	<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
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
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="../print_excel.asp"-->
				<!--#include file="menutop.asp"-->
			</div>
	<form name="frmVRA" method="post" action="<%=ScriptName%>">
	<input type="hidden" name="TxtTypeFCBA" value="<%=v_TxtTypeFCBA%>" >
	<div class="printhide">
	<!--#include file="FormSelectNewSumME.html"-->
	<!--#include file="FromSelectHDRVRA.html"-->
	</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
			<th colspan="50" style="background-color:#FFFFFF; border:none;">
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<% If  CompanyName <>  "" then %>
							<h4>PT.<%=CompanyName%></h4> 
						<% End if%>
						<% If  FcbaName <> ""  then %>
							   <h5><%=FcbaName%></h5>   
						<% End if%>
					</div>

					<div id="subheader" class="centered">
						<% Dim ia
						   ia = request("vMonth")
						%>   
						<h3>LAPORAN VEHICLE RUNNING ACCOUNT (VRA)</h3>
						<% IF request("vMonth") <> "" then %>
						<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=request("vYear")%></h5>
						<% END IF %>			
					</div>
				</div>
			</th>
		</tr>



	<tr>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="3">KODE</th>
		<% if v_txtVH <> "" then %>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="3">NAMA KENDARAAN</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="3">TAHUN PEROLEHAN</th>
		<% end if %>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="3">PLANTATION</th>			
	    <% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="3">GROUP REGION</th>			
	    <% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="3">REGION</th>		
	    <% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;" rowspan="3">COMPANY</th>		
	    <% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;" rowspan="3">SITE</th>			
	    <% end if %>
		<th  style="background-color:#00A300;color:#fff;"  colspan="4">AKTUAL PEMAKAIAN FISIK</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="3">RASIO SBI</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="3">BUDGET SETAHUN (KM)</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="5">AKTUAL BIAYA OPERASI BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="5">AKTUAL BIAYA OPERASI S/D BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="3">BUDGET SETAHUN (Rp)</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">RP/UNIT(KM/HM)  Rp)</th>
	</tr>
	<tr>
		<th  style="background-color:#00A300;color:#fff;"  colspan="2">BI</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="2">SBI</th>
		
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">GAJI SOPIR</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">BBM & PELUMAS</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">SUKU CADANG</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">REPARASI & LAIN-LAIN</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">TOTAL</th>
		
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">GAJI SOPIR</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">BBM & PELUMAS</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">SUKU CADANG</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">REPARASI & LAIN-LAIN</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">TOTAL</th>
		
		<th  style="background-color:#00A300;color:#fff;"  colspan="2">AKTUAL</th>
		<th  style="background-color:#00A300;color:#fff;"  rowspan="2">BUDGET SETAHUN</th>
	</tr>

	<tr>
		
		<th style="background-color:#00A300;color:#fff;"  >(KM/HM)</th>
		<th style="background-color:#00A300;color:#fff;"  >LTR</th>
		
		<th style="background-color:#00A300;color:#fff;"  >(KM/HM)</th>
		<th style="background-color:#00A300;color:#fff;"  >LTR</th>
		
		<th style="background-color:#00A300;color:#fff;"  >BI</th>
		<th style="background-color:#00A300;color:#fff;"  >SBI</th>
	</tr>
	</thead>
	<tbody>
	<%
	Dim rsha, strdtl ,actualBI,actualSBI,totactualBI,totactualSBI,rasioBI, strgrp ,cdgrp,cdgrpdesc  ,noline , subKMHMBI,subKMHMSBI,subltrBI,subltrSBI,subSALBI,subSALSBI,subbbmBI,subbbmSBI,subspartBI,subspartSBI,subrepBI,subrepSBI 
	Dim subrasioBI,subactualBI ,subactualSBI ,subtotactualBI ,subtotactualSBI ,subKMHMbudTHN,subAMTbudTHN
	Dim TotsubKMHMBI,TotsubKMHMSBI,TotsubltrBI,TotsubltrSBI,TotsubSALBI,TotsubSALSBI,TotsubbbmBI,TotsubbbmSBI,TotsubspartBI,TotsubspartSBI,TotsubrepBI,TotsubrepSBI 
	Dim TotsubrasioBI,TotsubactualBI ,TotsubactualSBI ,TotsubtotactualBI ,TotsubtotactualSBI ,TotsubKMHMbudTHN,TotsubAMTbudTHN
	Dim strjob,driverjob,gasolinejob,spartjob,repairjob,alljob, job_fisik, stock_fisik

	strjob= " SELECT report_posting_api.getPostingDataDirectNEW('PJJP','VRA','p_vra.asp','DRIVER_SALARY','JOB') drv_slr, " & _
	        " report_posting_api.getPostingDataDirectNEW('PJJP','VRA','p_vra.asp','GASOLINE_LUBRICANT','JOB')  gasoline, " & _  
			" report_posting_api.getPostingDataDirectNEW('PJJP','VRA','p_vra.asp','SPARE_PARTS','JOB')  sparts, " & _ 
			" report_posting_api.getPostingDataDirect('PJJP','VRA','p_vra.asp','FISIK_LTR','JOB')  job_fisik, " & _
			" report_posting_api.getPostingDataDirect('PJJP','VRA','p_vra.asp','FISIK_LTR','STOCK')  stock_fisik, " & _
			" report_posting_api.getPostingDataDirectNEW('PJJP','VRA','p_vra.asp','REPARATION_OTHERS','JOB')  repairs FROM DUAL"
	Set rsjob = Server.CreateObject("ADODB.Recordset")
	'response.write strjob
	rsjob.Open strjob, connstr, adOpenStatic
	if Not rsjob.eof then
	   driverjob = rsjob("drv_slr")
	   gasolinejob = rsjob("gasoline")
	   spartjob =  rsjob("sparts")
	   repairjob =  rsjob("repairs")
	   alljob = driverjob & "," & gasolinejob  & "," &  spartjob  & "," &  repairjob
	   job_fisik = rsjob("job_fisik")
	   stock_fisik = rsjob("stock_fisik")
	end if
	
			strdtl = " SELECT" & strFieldCons & " vehiclecode,vehicle_name,vehiclegroup_api.get_OrderReport(VEHICLE_API.get_vehiclegroupcode(vehiclecode,fcba),fcba)  VHrut,vehiclegroup_api.get_vehicletypecode(VEHICLE_API.get_vehiclegroupcode(vehiclecode,fcba),fcba) vehiclegrouptype,VEHICLE_API.get_vehiclegroupcode(vehiclecode,fcba) grpvehicle,VEHICLEGROUP_API.get_vehiclegroup_name(VEHICLE_API.get_vehiclegroupcode(vehiclecode,fcba),fcba) grpvehiclename,vehicle_api.get_vehicle_year(vehiclecode,fcba) vehicleYear,sum(KMHMBI) KMHMBI,sum(KMHMSBI) KMHMSBI,sum(ltrBI) ltrBI,sum(ltrSBI) ltrSBI,sum(SALBI) SALBI, sum(SALSBI) SALSBI,sum(BBMBI) bbmBI,sum(BBMSBI) bbmSBI,sum(spartBI) spartBI,sum(spartSBI) spartSBI,sum(repBI) repBI,sum(repSBI) repSBI,sum(KMHMbudTHN) KMHMbudTHN,sum(AmtbudTHN) AmtbudTHN  FROM ( " & _
					 "	SELECT va.fcba,va.vehiclecode,vehicle_api.get_vehiclecode_name(va.vehiclecode,va.fcba) vehicle_name,sum(va.unit) KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  FROM VEHICLEACTIVITY va  " & _
					 "	where  to_char(va.vehicledate,'YYYY')= '" & request("vYear") & "'  and to_char(va.vehicledate,'MM') =  lpad('" & request("vMonth") & "',2,0)   " & _
					 "	group by va.fcba,va.vehiclecode " & _
					 "	UNION ALL " & _
					 "	SELECT va.fcba,va.vehiclecode,vehicle_api.get_vehiclecode_name(va.vehiclecode,va.fcba) vehicle_name,0 KMHMBI,sum(va.unit) KMHMSBI,0 ltrBI,0 ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  FROM VEHICLEACTIVITY va  " & _
					 "	where  to_char(va.vehicledate,'YYYY')= '" & request("vYear") & "' and to_char(va.vehicledate,'MM') <=  lpad('" & request("vMonth") & "',2,0)   " & _
					 "	group by va.fcba,va.vehiclecode " & _
					 "	UNION ALL " & _
					 "	SELECT isd.fcba,isd.locationcode vehiclecode,vehicle_api.get_vehiclecode_name(isd.locationcode,isd.fcba) vehicle_name,0 KMHMBI,0 KMHMSBI,sum(isd.quantity) ltrBI , 0 ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN   FROM issuedetails isd  " & _
					 "	where isd.locationtype='VH' and to_char(get_issuedate(isd.sivcode,isd.fcba),'YYYY')= '" & request("vYear") & "' and to_char(get_issuedate(isd.sivcode,isd.fcba),'MM') =  lpad('" & request("vMonth") & "',2,0)   " & _
					 "        and instr('" & job_fisik & "', jobcode) <> 0 and instr('" & stock_fisik & "', stockcode) <> 0" & _
					 "	group by isd.fcba,isd.locationcode " & _
					 "	UNION ALL " & _
					 "	SELECT isd.fcba,isd.locationcode vehiclecode,vehicle_api.get_vehiclecode_name(isd.locationcode,isd.fcba) vehicle_name,0 KMHMBI,0 KMHMSBI, 0 ltrBI,sum(isd.quantity) ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN    FROM issuedetails isd  " & _
					 "	where isd.locationtype='VH'  and to_char(get_issuedate(isd.sivcode,isd.fcba),'YYYY')= '" & request("vYear") & "' and to_char(get_issuedate(isd.sivcode,isd.fcba),'MM') <=  lpad('" & request("vMonth") & "',2,0) " & _
					 "        and instr('" & job_fisik & "', jobcode) <> 0 and instr('" & stock_fisik & "', stockcode) <> 0" & _
					 "	group by isd.fcba,isd.locationcode " & _
					 "	UNION ALL " & _
					 "	SELECT rnd.fcba,rnd.locationcode vehiclecode,vehicle_api.get_vehiclecode_name(rnd.locationcode,rnd.fcba) vehicle_name,0 KMHMBI,0 KMHMSBI,sum(rnd.quantity) ltrBI , 0 ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN FROM receivenotedetail rnd, receivenote rn  " & _
					 "	where rnd.fcba = rn.fcba AND rnd.receivenotecode = rn.receivenotecode AND rnd.locationtype='VH' AND TO_CHAR(rn.rndate,'YYYY')= '" & request("vYear") & "' AND TO_CHAR(rn.rndate,'MM') =  LPAD('" & request("vMonth") & "',2,0)   " & _
					 "        AND INSTR('" & job_fisik & "', rnd.jobcode) <> 0 AND INSTR('" & stock_fisik & "', rnd.purchaseitemcode) <> 0" & _
					 "	group by rnd.fcba,rnd.locationcode " & _
					 "	UNION ALL " & _
					 "	SELECT rnd.fcba,rnd.locationcode vehiclecode,vehicle_api.get_vehiclecode_name(rnd.locationcode,rnd.fcba) vehicle_name,0 KMHMBI,0 KMHMSBI, 0 ltrBI , sum(rnd.quantity) ltrSBI,0 SALBI, 0 SALSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN FROM receivenotedetail rnd, receivenote rn  " & _
					 "	where rnd.fcba = rn.fcba AND rnd.receivenotecode = rn.receivenotecode AND rnd.locationtype='VH' AND TO_CHAR(rn.rndate,'YYYY')= '" & request("vYear") & "' AND TO_CHAR(rn.rndate,'MM') <=  LPAD('" & request("vMonth") & "',2,0)   " & _
					 "        AND INSTR('" & job_fisik & "', rnd.jobcode) <> 0 AND INSTR('" & stock_fisik & "', rnd.purchaseitemcode) <> 0" & _
					 "	group by rnd.fcba,rnd.locationcode " & _
					 "	UNION ALL " & _
					 "	SELECT c.fcba,c.targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,sum(c.amount) SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  FROM cost_activity c " & _ 
					 "	where c.jobcode in (" & driverjob & ")   and to_char(c.tdate,'YYYY')= '" & request("vYear") & "' and to_char(c.tdate,'MM')= lpad('" & request("vMonth") & "',2,0) " & _
					 "	and c.targettype='VH'    " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	UNION ALL " & _
					 "	SELECT c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,sum(c.amount) SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  FROM cost_activity c " & _
					 "	where c.jobcode in (" & driverjob & ")   and to_char(c.tdate,'YYYY')= '" & request("vYear") & "' and to_char(c.tdate,'MM')<= lpad('" & request("vMonth") & "',2,0)  " & _
					 "	and  c.targettype='VH'   " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	UNION ALL " & _
					 "	SELECT c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,sum(c.amount) bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  FROM cost_activity c " & _
					 "	where c.jobcode in (" & gasolinejob & ")   and to_char(c.tdate,'YYYY')= '" & request("vYear") & "' and to_char(c.tdate,'MM')= lpad('" & request("vMonth") & "',2,0)  " & _
					 "	and  c.targettype='VH'  " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	UNION ALL " & _
					 "	SELECT c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,sum(c.amount) bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN   FROM cost_activity c " & _
					 "	where  c.jobcode in (" & gasolinejob & ")   and to_char(c.tdate,'YYYY')= '" & request("vYear") & "'  and to_char(c.tdate,'MM')<= lpad('" & request("vMonth") & "',2,0) " & _
					 "	and c.targettype='VH'  " & _
					 "	group by c.fcba,c.targetcode  " & _
					 "	UNION ALL " & _
					 "	SELECT  c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,sum(c.amount) spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  FROM cost_activity c " & _
					 "	where  c.jobcode in (" & spartjob & ")    and to_char(c.tdate,'YYYY')= '" & request("vYear") & "' and to_char(c.tdate,'MM')= lpad('" & request("vMonth") & "',2,0)  " & _
					 "	and c.targettype='VH'  " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	UNION ALL " & _
					 "	SELECT  c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,sum(c.amount) spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  FROM cost_activity c " & _
					 "	where c.jobcode in (" & spartjob & ")  and to_char(c.tdate,'YYYY')= '" & request("vYear") & "' and to_char(c.tdate,'MM')<= lpad('" & request("vMonth") & "',2,0) " & _
					 "	and c.targettype='VH'  " & _
					 "	group by c.fcba,c.targetcode  " & _
					 "	UNION ALL " & _
					 "	SELECT  c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,sum(c.amount) repBI,0 repSBI,0 KMHMbudTHN,0 AmtbudTHN  FROM cost_activity c " & _
					 "	where c.jobcode in (" & repairjob & ")   and to_char(c.tdate,'YYYY')= '" & request("vYear") & "' and to_char(c.tdate,'MM')= lpad('" & request("vMonth") & "',2,0)   " & _
					 "	and c.targettype='VH'  " & _
					 "	group by c.fcba,c.targetcode " & _
					 "	UNION ALL " & _
					 "	SELECT  c.fcba,targetcode vehiclecode,vehicle_api.get_vehiclecode_name(c.targetcode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,sum(c.amount) repSBI,0 KMHMbudTHN,0 AmtbudTHN   FROM cost_activity c " & _
					 "	where c.jobcode in (" & repairjob & ")  and to_char(c.tdate,'YYYY')= '" & request("vYear") & "' and to_char(c.tdate,'MM')<= lpad('" & request("vMonth") & "',2,0)  " & _
					 "	and  c.targettype='VH'  " & _
					 "	group by c.fcba,c.targetcode  " & _
					 "	UNION ALL " & _
					 "	SELECT  c.fcba,c.vehiclecode,vehicle_api.get_vehiclecode_name(c.vehiclecode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,sum(nvl(c.qty_01,0)+nvl(c.qty_02,0)+nvl(c.qty_03,0)+nvl(c.qty_04,0)+nvl(c.qty_05,0)+nvl(c.qty_06,0)+nvl(c.qty_07,0)+nvl(c.qty_08,0)+nvl(c.qty_09,0)+nvl(c.qty_10,0)+nvl(c.qty_11,0)+nvl(c.qty_12,0)) KMHMbudTHN,0 AmtbudTHN   FROM BUDGET_TRANSACTION_VHHM c " & _
					 "	where c.budget_year ='" & request("vYear") & "'  " & _
					 "	group by c.fcba,c.vehiclecode " & _
					 "	UNION ALL " & _
					 "	SELECT  c.fcba,c.vehiclecode,vehicle_api.get_vehiclecode_name(c.vehiclecode,c.fcba) vehicle_name, 0 KMHMBI,0 KMHMSBI,0 ltrBI,0 ltrSBI ,0 SalBI,0 SalSBI,0 bbmBI,0 bbmSBI,0 spartBI,0 spartSBI,0 repBI,0 repSBI,0 KMHMbudTHN,sum(nvl(c.budget_01,0)+nvl(c.budget_02,0)+nvl(c.budget_03,0)+nvl(c.budget_04,0)+nvl(c.budget_05,0)+nvl(c.budget_06,0)+nvl(c.budget_07,0)+nvl(c.budget_08,0)+nvl(c.budget_09,0)+nvl(c.budget_10,0)+nvl(c.budget_11,0)+nvl(c.budget_12,0)) AmtbudTHN  FROM BUDGET_TRANSACTION_VEHICLE c " & _
					 "	where c.budget_year='" & request("vYear") & "' " & _
					 "	and c.jobcode in (" & alljob & ") " & _
					 " 	group by c.fcba,c.vehiclecode " & _
					 "	) where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = '" & Request("TxtTypeFCBA") &"' group by fcba,vehiclecode,vehicle_name   "	
					 
			strdtl = "SELECT vehiclecode,vehicle_name,VHrut,vehiclegrouptype, grpvehicle,grpvehiclename,vehicleYear,GroupPlantation" & selectCons & " ,sum(KMHMBI) KMHMBI,sum(KMHMSBI) KMHMSBI,sum(ltrBI) ltrBI,sum(ltrSBI) ltrSBI,sum(SALBI) SALBI, sum(SALSBI) SALSBI,sum(BBMBI) bbmBI,sum(BBMSBI) bbmSBI,sum(spartBI) spartBI,sum(spartSBI) spartSBI,sum(repBI) repBI,sum(repSBI) repSBI,sum(KMHMbudTHN) KMHMbudTHN,sum(AmtbudTHN) AmtbudTHN  FROM ( " & strdtl & " )" & whereCons & "  group by  vehiclecode,vehicle_name,VHrut,vehiclegrouptype, grpvehicle,grpvehiclename,vehicleYear,GroupPlantation " & selectCons & "  Order by VHrut,vehiclecode,vehicle_name,grpvehiclename,vehicleYear,GroupPlantation " & orderCons  

	'response.write strdtl
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
   
   ' response.write request("vMonth") &  "  " & request("vYear") & " " & session("fcba") 
	
	if Not rsha.eof then
		  while Not rsha.eof
		    IF  cdgrp = rsha("grpvehicle") THEN
			    %>
				<tr>  
					<td align="CENTER"><%=rsha("vehiclecode")%></td>
					<td align="LEFT"><%=rsha("vehicle_name")%></td>
					<td align="center"><%=rsha("vehicleYear")%></td>
					<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     
						   tempPlantation = rsha("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempVH ="--SUMMARY--"
					%>
					<% if SprevPlantation <> rsha("Plantation")   then %>
						<td align="center">
							<a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsha("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  
					    
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsha("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsha("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempVH ="--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsha("GroupRegion") then %>
						<td align="center">
							<a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsha("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsha("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempVH = "--SUMMARY--"
					   
					%>
					
					<% if SprevRegion <> rsha("Region") then %>
						<td align="center">
							<a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("region")%>
							</a>
						</td>
						<% SprevRegion = rsha("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if	
					   tempCompany =rsha("Company")
					   tempfcba = "--ALL--"	
					   tempVH ="--SUMMARY--"
					 %>
					
					<% if SprevCompany <> rsha("Company") then %>
						<td align="center">
							<a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("company")%>
							</a>
						</td>
						<% SprevCompany = rsha("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
					 <% 
					   
					   					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtcompany <>  "--SUMMARY--" And v_txtcompany <>  "" then
					      tempCompany = rsha("Company")
					   else
					      tempCompany = ""
                       end if				   
					   tempfcba = rsha("fcba") 
					   tempVH ="--ALL--"					   
					 %>
					<% if sprevFcba <> rsha("fcba") then %>
						<td align="center">
						    <a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("fcba")%>
						</td>
						<% sprevFcba = rsha("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				
					<td align="right"><%=formatnumber( cdbl(rsha("KMHMBI")),0)%></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("LTRBI")),0)%></td></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("KMHMSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("LTRSBI")),0)%></td></td>
					<%
					If Cstr(rsha("vehiclegrouptype")) = Cstr("KR") Then
						If Cdbl(rsha("LTRSBI")) > 0 Then
							rasioBI = cdbl(rsha("KMHMSBI"))/ cdbl(rsha("LTRSBI"))
						Else
							rasioBI = 0
						End If
					Else
						If Cdbl(rsha("KMHMSBI")) > 0 Then
							rasioBI = cdbl(rsha("LTRSBI"))/ cdbl(rsha("KMHMSBI"))
						Else
							rasioBI = 0
						End If
					End If
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
					<% if cdbl(rsha("KMHMBI")) > 0 then
					        totactualBI = cdbl(actualBI)/ cdbl(rsha("KMHMBI")) 
					   else 
					        totactualBI  ="0"
				       end if	
					   if cdbl(rsha("KMHMSBI")) > 0 then
					        totactualSBI = cdbl(actualSBI)/ cdbl(rsha("KMHMSBI")) 
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
						   <td style="background-color:#F4F4A0" align="center" colspan="<%= 3 + nSpanTotal %>"><b>SUB TOTAL&nbsp;&nbsp;<%=cdgrpdesc%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subKMHMBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subltrBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subKMHMSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subltrSBI),0)%></td>						   
						   <!--<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrasioBI),2)%></td>-->
						   <%If Cstr(val_vehiclegrouptype) = CStr("KR") Then%>
							   <%If subltrSBI > 0 Then%>
								<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber(cdbl(subKMHMSBI)/cdbl(subltrSBI), 2)%></td>
							   <%Else%>
								<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber(0, 2)%></td>
							   <%End If%>
						   <%Else%>
						       <%If subKMHMSBI > 0 Then%>
								<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber(cdbl(subltrSBI)/cdbl(subKMHMSBI), 2)%></td>
							   <%Else%>
								<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber(0, 2)%></td>
							   <%End If%>
						   <%End if%>
						   
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subKMHMbudTHN),2)%></td>
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
						     <% if cdbl(subKMHMBI) > 0 then
									subtotactualBI = cdbl(subactualBI)/ cdbl(subKMHMBI)
							   else 
									subtotactualBI  ="0"
							   end if	
							   if cdbl(subKMHMSBI) > 0 then
									subtotactualSBI = cdbl(subactualSBI)/ cdbl(subKMHMSBI)
							   else 
									subtotactualSBI  ="0"
							   end if
							%>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotactualBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subtotactualSBI),0)%></td>
						   <% if cdbl(subKMHMbudTHN) > 0 then %>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subAMTbudTHN)/cdbl(subKMHMbudTHN),0)%></td>  
						   <% Else %>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subAMTbudTHN),0)%></td> 
						   <% End if %> 
					   </tr> 
				<%	end if  %>  
				
				<tr>  
					<td align="CENTER"><%=rsha("vehiclecode")%></td>
					<td align="LEFT"><%=rsha("vehicle_name")%></td>
					<td align="center"><%=rsha("vehicleYear")%></td>
					<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     
						   tempPlantation = rsha("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						   tempVH ="--SUMMARY--"
					%>
					<% if SprevPlantation <> rsha("Plantation")   then %>
						<td align="center">
							<a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsha("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  
					    
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsha("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsha("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempVH ="--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsha("GroupRegion") then %>
						<td align="center">
							<a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsha("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsha("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempVH = "--SUMMARY--"
					   
					%>
					
					<% if SprevRegion <> rsha("Region") then %>
						<td align="center">
							<a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("region")%>
							</a>
						</td>
						<% SprevRegion = rsha("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if	
					   tempCompany =rsha("Company")
					   tempfcba = "--ALL--"	
					   tempVH ="--SUMMARY--"
					 %>
					
					<% if SprevCompany <> rsha("Company") then %>
						<td align="center">
							<a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("company")%>
							</a>
						</td>
						<% SprevCompany = rsha("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
					 <% 
					   
					   					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsha("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsha("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsha("Region") 
					   else
					       tempRegion = ""
					   end if
					   if v_txtcompany <>  "--SUMMARY--" And v_txtcompany <>  "" then
					      tempCompany = rsha("Company")
					   else
					      tempCompany = ""
                       end if				   
					   tempfcba = rsha("fcba") 
					   tempVH ="--ALL--"					   
					 %>
					<% if sprevFcba <> rsha("fcba") then %>
						<td align="center">
						    <a href="p_vra.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtVH=<%=tempVH%>&vYear=<%=request("vYear")%>&vMonth=<%=request("vMonth")%>">
							<%=rsha("fcba")%>
						</td>
						<% sprevFcba = rsha("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("KMHMBI")),0)%></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("LTRBI")),0)%></td></td>
				    <td align="right"><%=formatnumber( cdbl(rsha("KMHMSBI")),0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("LTRSBI")),0)%></td></td>
					<%
					If rsha("vehiclegrouptype") = "KR" Then
						If Cdbl(rsha("LTRSBI")) > 0 Then
							rasioBI = cdbl(rsha("KMHMSBI"))/ cdbl(rsha("LTRSBI"))
						Else
							rasioBI = 0
						End If
					Else
						If Cdbl(rsha("KMHMSBI")) > 0 Then
							rasioBI = cdbl(rsha("LTRSBI"))/ cdbl(rsha("KMHMSBI"))
						Else
							rasioBI = 0
						End If
					End If
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
					<% if cdbl(rsha("KMHMBI")) > 0 then
					        totactualBI = cdbl(actualBI)/ cdbl(rsha("KMHMBI")) 
					   else 
					        totactualBI  ="0"
				       end if	
					   if cdbl(rsha("KMHMSBI")) > 0 then
					        totactualSBI = cdbl(actualSBI)/ cdbl(rsha("KMHMSBI")) 
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
			  val_vehiclegrouptype = rsha("vehiclegrouptype")
			  rsha.MoveNext
		  wend
	end if
	%>	
					    <tr>
						   <td style="background-color:#F4F4A0" align="center" colspan="<%= 3 + nSpanTotal %>"><b>SUB TOTAL&nbsp;&nbsp;<%=cdgrpdesc%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subKMHMBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subltrBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subKMHMSBI),0)%></td>
						   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subltrSBI),0)%></td>
						   <!--<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(subrasioBI),2)%></td>-->
						   <%If val_vehiclegrouptype = "KR" Then%>
							   <%If subltrSBI > 0 Then%>
								<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber(cdbl(subKMHMSBI)/cdbl(subltrSBI), 2)%></td>
							   <%Else%>
								<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber(0, 2)%></td>
							   <%End If%>
						   <%Else%>
						       <%If subKMHMSBI > 0 Then%>
								<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber(cdbl(subltrSBI)/cdbl(subKMHMSBI), 2)%></td>
							   <%Else%>
								<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber(0, 2)%></td>
							   <%End If%>
						   <%End if%>
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
						   <% if cdbl(subKMHMBI) > 0 then
									subtotactualBI = cdbl(subactualBI)/ cdbl(subKMHMBI)
							   else 
									subtotactualBI  ="0"
							   end if	
							   if cdbl(subKMHMSBI) > 0 then
									subtotactualSBI = cdbl(subactualSBI)/ cdbl(subKMHMSBI)
							   else 
									subtotactualSBI  ="0"
							   end if
							%>
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
       <td style="background-color:#F4F4A0" align="center" colspan="<%= 3 + nSpanTotal %>"><b>GRAND TOTAL</td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubKMHMBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubltrBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubKMHMSBI),0)%></td>
	   <td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubltrSBI),0)%></td>
	   <!--<td style="background-color:#F4F4A0"  align="right"><b><%=formatnumber( cdbl(TotSubrasioBI),2)%></td>-->

	   <td style="background-color:#F4F4A0" align="right"><b>&nbsp;</td>
	   
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
	   <td style="background-color:#F4F4A0"  align="right"><b>&nbsp;</td>
	   <td style="background-color:#F4F4A0"  align="right"><b>&nbsp;</td>
	   <% if cdbl(TotsubKMHMbudTHN) > 0 then %>
			<td style="background-color:#F4F4A0"  align="right"><b>&nbsp;</td>  
	  <% Else %>
		   <td style="background-color:#F4F4A0"  align="right"><b>&nbsp;</td> 
	  <% End if %> 
   </tr>
  </tbody>
</table>
</form>
<!--#include file="bottom.Asp"-->