<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
' Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
' Dim ScriptName,nSpanTotal,no
' Dim v_txtyear,v_txtmonth,v_vJenisReport,v_txtJob,vJenisFactory,v_JenisFactory,v_txtstation,v_txtmachine
' Dim SprevCode,Sprevjobname
' Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
' Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
' Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempjob
' Dim rsSQLPreParameter,strSQLPreParameter,txtdate,txtmonth, 	txtyear
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
' DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,rsJob
	v_txtGroupPlantation  = Request("txtGroupPlantation")
	v_txtPlantation  	  = Request("txtPlantation")
	v_txtGroupRegion  	  = Request("txtGroupRegion")
	v_txtRegion  		  = Request("txtRegion")
	v_txtcompany  		  = Request("txtCompany")
	v_txtfcba 			  = Request("txtfcba")
	v_txtyear             = Request("txtYear")
	v_txtmonth            = Request("txtMonth")
	v_vJenisReport        = request("vJenisReport")
	v_txtstation		  = Request("txtstation")
	v_txtmachine		  = Request("txtmachine")
	v_txtjob    		  = Request("txtjob")
	v_JenisFactory        = request("vJenisFactory")
	nSpanTotal = 11

	if  v_txtstation = "" then
		v_txtstation = "--ALL--"
	end if
		
' Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
	strFieldCons = 	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation," & _
					" USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
					" USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ," & _
					" USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region," & _ 
					" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company," & _
					" fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut,"

	strFieldCons2 = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation," & _
					" USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
					" USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ," & _
					" USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region," & _ 
					" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company," & _
					" fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut,"
						   
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
		selectCons2 = selectCons2 & ",Plantation"
		orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
	else
	    nSpanTotal = nSpanTotal - 1 
	end if
	if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
	    if v_txtGroupRegion <> "--ALL--" then
		   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
		end if
		selectCons = selectCons & " ,GroupRegion"
		selectCons2 = selectCons2 & " ,GroupRegion"
		orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
	else
	    nSpanTotal = nSpanTotal - 1 
	end if
	if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
	    if v_txtRegion <> "--ALL--" then
		   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
		end if
		selectCons = selectCons & " ,Region"
		selectCons2 = selectCons2 & " ,Region"
		orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
	else
	    nSpanTotal = nSpanTotal - 1 
	end if
	if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
	    if v_txtCompany  <> "--ALL--" then
		   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
		end if
		selectCons = selectCons & " ,company"
		selectCons2 = selectCons2 & " ,company"
	    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
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
	    selectCons2 = selectCons2 & ",norut,fcba"
	    orderCons  = orderCons & ",norut "
		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
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


	if  v_txtstation = "--SUMMARY--" then
	    nSpanTotal = nSpanTotal - 2
	end if	
	if  v_txtmachine = "--SUMMARY--" then
	    nSpanTotal = nSpanTotal - 2 
	end if	
	if  v_txtjob = "--SUMMARY--" then
	    nSpanTotal = nSpanTotal - 2 
	end if

	if len(trim(v_txtyear)) = 0 then
	  ' txtdate	= date + 300
	  nSpanTotal = 2
	end if
			
'if len(trim(v_txtyear)) = 0 then  
'  strSQLPreParameter = "select to_char(to_date('" & txtdate & "', 'MM/DD/YYYY'),'YYYY') as CurrentYear, " & _
'		                "to_char(to_date('" & txtdate & "', 'MM/DD/YYYY'),'MM') as CurrentPeriod from dual "
'   Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
'	rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 		
'		if not rsSQLPreParameter.EOF then
'			v_txtyear = rsSQLPreParameter("CurrentYear")
'			v_txtmonth = rsSQLPreParameter("CurrentPeriod")
'		end if
'		rsSQLPreParameter.Close
'		set rsSQLPreParameter = nothing	
'		If v_txtmonth = 1 then
'		   v_txtmonth = "12"
'          v_txtyear =  cstr(cint(v_txtyear) - 1)
'		else
'		   v_txtmonth = cstr(cint(v_txtmonth) - 1) 
'        end if
'        txtmonth=v_txtmonth
'		txtyear=v_txtyear
'		nSpanTotal = 2
'   end if			
%>
<html>
	<head>
		<title>iPla'S Reporting - Analisa Biaya Langsung dan Biaya Tidak Langsung Pabrik</title>
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
				$( "#datepicker" ).datepicker();
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
	</head>
<body>
	<div id="wrapper">
		<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<!--#include file="menutop.asp"-->
			<form name="frmMill" method="post" action="<%=ScriptName%>">
			<!--#include file="FormSelectNewSumM.html"-->
			<!--#include file="FromSelectHDRBiayaPabrik.html"-->			

				<div id="header">
					<% If  CompanyName <>  "" then %>
						<h4>PT.<%=CompanyName%></h4> 
					<% End if%>
					<% If  FcbaName <> ""  then %>
						   <h5><%=FcbaName%></h5>   
					<% End if%>
				</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
			<th colspan="26" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">

			<div id="subheader" class="centered">
			    <% Dim ia,rekap
				   ia = v_txtmonth
				   if   v_vJenisReport  = "Group"  then
				        rekap = "LAPORAN BIAYA PER STASIUN"
				   elseif v_vJenisReport  = "Job"  then	
						rekap = "LAPORAN BIAYA PER STASIUN PER PEKERJAAN"
				   elseif v_vJenisReport  = "Mechine"  then	
						rekap = "LAPORAN BIAYA PER STASIUN PER MESIN"
				   elseif v_vJenisReport  = "Combine"  then	
						rekap = "LAPORAN BIAYA PER STASIUN PER MESIN PER PEKERJAAN"
				   end if
						
				%>   
				<h3>LAPORAN BIAYA LANGSUNG DAN BIAYA TIDAK LANGSUNG PABRIK</h3>
					<% IF v_txtmonth <> "" then %>
				<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=v_txtyear %></h5>
					<% END IF %>
				<h5><%=rekap%></h5> 
			</div>
		</div>
<%
Dim cur_month,strbud_BI,strbud_SBI,strbudHDR_BI,strbudHDR_SBI,x1
strbud_BI = "btm.budget_01" 
strbud_SBI = "btm.budget_01"
strbudHDR_BI = "btm.qty_01"
strbudHDR_SBI = "btm.qty_01"


cur_month = v_txtmonth 

IF Len(cur_month) > 0 then

	strbud_BI = "" 
    strbud_SBI = ""
	strbudHDR_BI = ""
    strbudHDR_SBI = ""
	x1=1
	
	Do While x1<= cint(cur_month) 'Output the values from 0 to 10
	  'response.write(i & "<br \>")
	  'response.write  " Budget_" & Lpad(x1,"0",2) & " +"
	  strbud_SBI = strbud_SBI & " btm.budget_" & Lpad(x1,"0",2) & " +"
	  strbudHDR_SBI = strbudHDR_SBI & "btm.qty_" & Lpad(x1,"0",2) & " +"
	  x1 = x1 + 1 'increment the value of i for next time loop executes
	Loop

	strbud_BI = "btm.budget_" & Lpad(cur_month,"0",2)
	strbudHDR_BI  = "btm.qty_" & Lpad(cur_month,"0",2)
	
	
	strbud_SBI = Mid(strbud_SBI,1,Len(strbud_SBI)-1)
	strbudHDR_SBI = Mid(strbudHDR_SBI,1,Len(strbudHDR_SBI)-1)
	
	
	
    ' response.write  strbudTBS_BI & "<br>"
	' response.write strbudCPO_BI & "<br>"
	' response.write strbudKER_BI & "<br>"
	
	
	' response.write strbudTBS_SBI  & "<br>"
	' response.write strbudCPO_SBI  & "<br>"
	' response.write strbudKER_SBI  & "<br>"
END IF	

Dim strtbs, rsha,TBSprocess_BI,TBSprocess_SBI,ProdCPO_BI,ProdCPO_SBI,ProdKER_BI,ProdKER_SBI,Palm_BI,Palm_SBI, strprod
TBSprocess_BI=0
TBSprocess_SBI=0

strtbs = "  select " & strFieldCons & " sum(TBS_PROCESS_BI) TBS_PROCESS_BI, sum(TBS_PROCESS_SBI) TBS_PROCESS_SBI  from ( " & _
		 "  select t.fcba, sum(t.cageweight) TBS_PROCESS_BI,0 TBS_PROCESS_SBI from AR_SHIFTREPORT t " & _  
	     "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.productiondate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  " & _
		 "  union " & _
		 "  select t.fcba, 0 TBS_PROCESS_BI, sum(t.cageweight) TBS_PROCESS_SBI from AR_SHIFTREPORT t  " & _
		 "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.productiondate,'MM') <= lpad('" & v_txtmonth & "',2,0) " & _
         "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  ) group by fcba " 

 strtbs = "select GroupPlantation , sum(TBS_PROCESS_BI) TBS_PROCESS_BI, sum(TBS_PROCESS_SBI) TBS_PROCESS_SBI  from ( " & strtbs & " )" & whereCons & " group by GroupPlantation"
' response.write strtbs		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    TBSprocess_BI = rsha("TBS_PROCESS_BI")
	TBSprocess_SBI = rsha("TBS_PROCESS_SBI")
end if
rsha.close()

ProdCPO_BI =0
ProdCPO_SBI =0

' strtbs = " select fcba, sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 ' " select t.fcba,  sum(t.quantity) Produksi_CPO_BI,0 Produksi_CPO_SBI from PRODUCTSTORAGEREADING t  " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='CPO')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba, to_char(t.readingdate,'YYYYMM') " & _
		 ' " union " & _
		 ' " select t.fcba, 0 Produksi_CPO_BI,sum(t.quantity) Produksi_CPO_SBI from PRODUCTSTORAGEREADING t " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='CPO')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba, to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select " & strFieldCons & " sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_CPO_BI, 0 Produksi_CPO_SBI from AR_MILL_PRODUCTION t  " & _
		 " where t.product_code='CPO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union " & _
		 " select t.fcba,0 Produksi_CPO_BI, sum(t.production_today) Produksi_CPO_SBI from AR_MILL_PRODUCTION t " & _
		 " where t.product_code='CPO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba " 
		 
 strtbs = "select GroupPlantation , sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI   from ( " & strtbs & " )" & whereCons & " group by GroupPlantation"
 ' response.write strtbs			  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    ProdCPO_BI = rsha("Produksi_CPO_BI")
	ProdCPO_SBI = rsha("Produksi_CPO_SBI")
end if
rsha.close()

ProdKER_BI =0
ProdKER_SBI =0

' strtbs = " select fcba, sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI from ( " & _
	     ' " select t.fcba, sum(t.quantity) Produksi_KER_BI, 0 Produksi_KER_SBI from PRODUCTSTORAGEREADING t  " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='Kernel')   and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba,to_char(t.readingdate,'YYYYMM')  " & _
		 ' " union " & _
		 ' " select t.fcba, 0 Produksi_KER_BI, sum(t.quantity) Produksi_KER_SBI from PRODUCTSTORAGEREADING t " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='Kernel')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba,to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select " & strFieldCons & "  sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_KER_BI, 0 Produksi_KER_SBI from AR_MILL_PRODUCTION t  " & _
		 " where t.product_code='KER'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union " & _
		 " select t.fcba,0 Produksi_KER_BI, sum(t.production_today) Produksi_KER_SBI from AR_MILL_PRODUCTION t " & _
		 " where t.product_code='KER'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)  " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba " 		
strtbs = "select GroupPlantation ,  sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI   from ( " & strtbs & " )" & whereCons & " group by GroupPlantation"
 'response.write strtbs	  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic
'response.write strtbs
if Not rsha.eof then
    ProdKER_BI = rsha("Produksi_KER_BI")
	ProdKER_SBI = rsha("Produksi_KER_SBI")
end if
rsha.close()
Palm_BI = cdbl(ProdKER_BI) +  cdbl(ProdCPO_BI)
Palm_SBI = cdbl(ProdKER_SBI) +  cdbl(ProdCPO_SBI)

' ambil data budget CPO, TBS dan ker

Dim TBSbud_BI,TBSbud_SBI,TBSbud_THN,CPObud_BI,CPObud_SBI,CPObud_THN,KERbud_BI,KERbud_SBI,KERbud_THN,PALMbud_BI,PALMbud_SBI,PALMbud_THN
strtbs = "  select" & strFieldCons & "  sum(TBSbudBI) TBSbudBI,sum(tbsbudSBI) tbsbudSBI,sum(tbsbudthn) tbsbudthn,sum(cpobudBI) cpobudBI,sum(cpobudSBI) cpobudSBI,sum(cpobudthn) cpobudthn,sum(kerbudBI) kerbudBI,sum(kerbudSBI) kerbudSBI, sum(kerbudthn) kerbudthn  from ( " & _
		 "	 select btm.fcba , sum(" & strbudHDR_BI & ") TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='TBS' " & _
		 "	 group by btm.fcba " & _
		 "	 union " & _
		 "	 select btm.fcba , 0 TBSbudBI,sum(" & strbudHDR_SBI & ") tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='TBS' " & _
		 "	 group by btm.fcba " & _
		 "	 union " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='TBS' " & _
		 "	 group by btm.fcba " & _
		 "	 union " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,sum(" & strbudHDR_BI & ") cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='CPO' " & _
		 "	 group by btm.fcba " & _
		 "	 union " & _
		 "	 select btm.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,sum(" & strbudHDR_SBI & ") cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='CPO' " & _
		 "	 group by btm.fcba " & _
		 "	 union " & _
		 "	 select btm.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) cpobudthn,0 kerbudBI, 0 kerbudSBI , 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='CPO' " & _
		 "	 group by btm.fcba " & _
		 "	  union " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudthn,0 cpobudSBI,sum(" & strbudHDR_BI & ") kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='KER' " & _
		 "	 group by btm.fcba " & _
		 "	  union " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, sum(" & strbudHDR_SBI & ") kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='KER' " & _
		 "	 group by btm.fcba " & _
		 "	 union " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='KER' " & _
		 "	 group by btm.fcba ) group by fcba " 

		 strtbs = "select GroupPlantation , sum(TBSbudBI) TBSbudBI,sum(tbsbudSBI) tbsbudSBI,sum(tbsbudthn) tbsbudthn,sum(cpobudBI) cpobudBI,sum(cpobudSBI) cpobudSBI,sum(cpobudthn) cpobudthn,sum(kerbudBI) kerbudBI,sum(kerbudSBI) kerbudSBI, sum(kerbudthn) kerbudthn   from ( " & strtbs & " )" & whereCons & " group by GroupPlantation"
' response.write strtbs		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    TBSbud_BI = rsha("TBSbudBI")
	TBSbud_SBI = rsha("TBSbudSBI")
	TBSbud_THN = rsha("TBSbudTHN")
	CPObud_BI = rsha("cpobudBI")
	CPObud_SBI = rsha("cpobudSBI")
	CPObud_THN = rsha("cpobudTHN")
	KERbud_BI = rsha("kerbudBI")
	KERbud_SBI = rsha("kerbudSBI")
	KERbud_THN = rsha("kerbudTHN")
	PALMbud_BI =   cdbl(KERbud_BI) +  cdbl(CPObud_BI)
	PALMbud_SBI =  cdbl(KERbud_SBI) +  cdbl(CPObud_SBI)
	PALMbud_THN =  cdbl(KERbud_THN) +  cdbl(CPObud_THN)
end if
rsha.close()
if cdbl(TBSbud_BI)=0 then
     TBSbud_BI=1
end if
if cdbl(TBSbud_SBI)=0 then
     TBSbud_SBI=1
end if
if cdbl(TBSbud_THN)=0 then
     TBSbud_THN=1
end if
if cdbl(CPObud_BI) = 0 then
   CPObud_BI=1
end if   
if cdbl(CPObud_SBI)=0 then
     CPObud_SBI=1
end if
if cdbl(CPObud_THN)=0 then
     CPObud_THN=1
end if
if cdbl(KERbud_BI)=0 then
     KERbud_BI=1
end if
if cdbl(KERbud_SBI)=0 then
     KERbud_SBI=1
end if
if cdbl(KERbud_THN)=0 then
     KERbud_THN=1
end if
if cdbl(PALMbud_BI)=0 then
     PALMbud_BI=1
end if
if cdbl(PALMbud_SBI)=0 then
     PALMbud_SBI=1
end if
if cdbl(PALMbud_THN)=0 then
     PALMbud_THN=1
end if
%>	
</th>
</tr>
	<tr>
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= nSpanTotal %>"></th>
		
		<th  style="background-color:#00A300;color:#fff;"  colspan="7">BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="7">S/D BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= nSpanTotal %>"></th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">REALISASI</th>
		
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;"  >%</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">REALISASI</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;"  >%</th>
	</tr>
	<tr>
	    <% 'TBSbud_BI,TBSbud_SBI,TBSbud_THN,CPObud_BI,CPObud_SBI,CPObud_THN,KERbud_BI,KERbud_SBI,KERbud_THN,PALMbud_BI,PALMbud_SBI,PALMbud_THN %>
		<th  style="background-color:#00A300;color:#fff;"   colspan="<%= nSpanTotal %>" align="right">TBS PROSES (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSprocess_BI,2)%></th>
		<% if cdbl(TBSbud_BI) <> 1 then %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSbud_BI,2)%></th>
			   <th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(TBSprocess_BI)-cdbl(TBSbud_BI))/(cdbl(TBSbud_BI))*100,2)%></th>
		<% Else %>
		       <th style="background-color:#00A300;color:#fff;"  align="right" colspan="3" >0</th>
			   <th style="background-color:#00A300;color:#fff;"  align="right" >0</th>
		<% End if %>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSprocess_SBI,2)%></th>
		<% if cdbl(TBSbud_SBI) <> 1 then %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSbud_SBI,2)%></th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(TBSprocess_SBI)-cdbl(TBSbud_SBI))/(cdbl(TBSbud_SBI))*100,2)%></th>
	   <% Else %>
		       <th style="background-color:#00A300;color:#fff;"  align="right" colspan="3" >0</th>
			   <th style="background-color:#00A300;color:#fff;"  align="right" >0</th>
		<% End if %>	
		<% if cdbl(TBSbud_THN) <> 1 then %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSbud_THN,2)%></th>
		 <% Else %>
		       <th style="background-color:#00A300;color:#fff;"  align="right" colspan="3" >0</th>
		<% End if %>	   
		
		
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= nSpanTotal  %>" align="right">PRODUKSI CPO (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdCPO_BI,2)%></th>
		
		<% if cdbl(CPObud_BI) <> 1 then %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(CPObud_BI,2)%></th>
		       <th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(ProdCPO_BI)-cdbl(CPObud_BI))/(cdbl(CPObud_BI))*100,2)%></th>
		<% Else %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		       <th style="background-color:#00A300;color:#fff;"  align="right" >0</th>
		<% End if %>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdCPO_SBI,2)%></th>
		
		<% if cdbl(CPObud_SBI) <> 1 then %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(CPObud_SBI,2)%></th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(ProdCPO_SBI)-cdbl(CPObud_SBI))/(cdbl(CPObud_SBI))*100,2)%></th>
		<% Else %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		       <th style="background-color:#00A300;color:#fff;"  align="right" >0</th>
		<% End if %>
		<% if cdbl(CPObud_THN) <> 1 then %>
		       <th style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(CPObud_THN,2)%></th>
		<% Else %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		<% End if %>	   
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= nSpanTotal %>" align="right">PRODUKSI KER (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdKER_BI,2)%></th>
		<% if cdbl(KERbud_BI) <> 1 then %>
		      <th style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(KERbud_BI,2)%></th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(ProdKER_BI)-cdbl(KERbud_BI))/(cdbl(KERbud_BI))*100,2)%></th>
		<% Else %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" >0</th>
		<% End if %>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdKER_SBI,2)%></th>
		<% if cdbl(KERbud_SBI) <> 1 then %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(KERbud_SBI,2)%></th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(ProdKER_SBI)-cdbl(KERbud_SBI))/(cdbl(KERbud_SBI))*100,2)%></th>
		<% Else %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" >0</th>
		<% End if %>
		<% if cdbl(KERbud_SBI) <> 1 then %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(KERbud_THN,2)%></th>
		<% Else %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		<% End if %>	  
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= nSpanTotal %>" align="right">PRODUK (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(Palm_BI,2)%></th>
		<% if cdbl(Palmbud_BI) <> 1 then %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(PALMbud_BI,2)%></th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(Palm_BI)-cdbl(Palmbud_BI))/(cdbl(Palmbud_BI))*100,2)%></th>
		<% Else %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" >0</th>
		<% End if %>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(Palm_SBI,2)%></th>
		<% if cdbl(Palmbud_SBI) <> 1 then %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(PALMbud_SBI,2)%></th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(Palm_SBI)-cdbl(Palmbud_SBI))/(cdbl(Palmbud_SBI))*100,2)%></th>
		<% Else %>
		      <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		      <th style="background-color:#00A300;color:#fff;"  align="right" >0</th>
		<% End if %>
		<% if cdbl(PALMbud_THN) <> 1 then %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(PALMbud_THN,2)%></th>
		<% Else %>
		       <th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3">0</th>
		<% End if %>
	</tr>
	
	<tr>
	    <%  'response.write  nSpanTotal 
		    IF nSpanTotal = -1 then %>
		       <th style="background-color:#00A300;color:#fff;" rowspan="2" colspan="<%= nSpanTotal -2 %>"></th>	   
	    <% END IF %>
	    	  
		<%  if v_txtstation <> "--SUMMARY--" and v_txtstation <> ""  then %>
			  <th style="background-color:#00A300;color:#fff;" rowspan="2" >Biaya</th>	
		      <th style="background-color:#00A300;color:#fff;" rowspan="2" >Stasiun</th> 
		<% end if %>
		<%  if v_txtmachine <> "--SUMMARY--" and v_txtmachine <> ""  then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="2" >Kode Mesin</th>
		      <th style="background-color:#00A300;color:#fff;" rowspan="2" >Nama Mesin</th>
		<% end if %>
        <%  if v_txtjob <> "--SUMMARY--" and v_txtjob <> ""  then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="2" >Kode Job</th>
		      <th style="background-color:#00A300;color:#fff;" rowspan="2" >Nama Job</th>
		<% end if %>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="2">PLANTATION</th>			
	    <% end if %>
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="2">GROUP REGION</th>			
	    <% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="2">REGION</th>		
	    <% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;" rowspan="2">COMPANY</th>		
	    <% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;" rowspan="2">SITE</th>			
	    <% end if %>

		
		<th style="background-color:#00A300;color:#fff;"  rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">Rp./Kg.</th>
		
		<th style="background-color:#00A300;color:#fff;"  rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">Rp./Kg.</th>
		<th style="background-color:#00A300;color:#fff;"  rowspan="2" align="right" ></th>
		<th style="background-color:#00A300;color:#fff;"  rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">Rp./Kg.</th>
		<th style="background-color:#00A300;color:#fff;"  rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">Rp./Kg.</th>
		<th style="background-color:#00A300;color:#fff;"  rowspan="2" align="right" ></th>
		<th style="background-color:#00A300;color:#fff;"  rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;"  colspan="2">Rp./Kg.</th>
	</tr>
	<tr>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >TBS</th>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >TBS</th>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >TBS</th>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >TBS</th>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >TBS</th>
	</tr>
</thead>
<tbody>	
	<%
		    if  v_txtstation <> "--SUMMARY--"  and v_txtstation <> "" then
			    if v_txtstation  <> "--ALL--" then
			       whereCons = whereCons & " And  stationcode = '" & v_txtstation & "' "
				end if
				strFieldCons = strFieldCons & "stationcode,STATIONCODE_DESCRIPTION,"
				strFieldCons2 = strFieldCons2 & "'' stationcode, ''STATIONCODE_DESCRIPTION,"
			    selectCons = selectCons & ",stationcode,STATIONCODE_DESCRIPTION"
			    selectCons2 = selectCons2 & ",stationcode,STATIONCODE_DESCRIPTION"
			    orderCons  = orderCons & ",station_api.get_orderdata(stationcode,'PSR') " 
			end if	
			if  v_txtmachine <> "--SUMMARY--"  and v_txtmachine <> "" then
			    if v_txtmachine  <> "--ALL--" then
			       whereCons = whereCons & " And  MACHINECODE = '" & v_txtmachine & "' "
				end if
				strFieldCons = strFieldCons & "MACHINECODE,MACHINECODE_DESCRIPTION,"
				strFieldCons2 = strFieldCons2 & "'' MACHINECODE, ''MACHINECODE_DESCRIPTION,"
			    selectCons = selectCons & ",MACHINECODE,MACHINECODE_DESCRIPTION"
			    selectCons2 = selectCons2 & ",MACHINECODE,MACHINECODE_DESCRIPTION"
			    orderCons  = orderCons & ",MACHINE_API.get_orderdata(MACHINECODE,'PSR') "
			end if	
            if  v_txtjob <> "--SUMMARY--"  and v_txtjob <> "" then
			    if v_txtjob  <> "--ALL--" then
			       whereCons = whereCons & " And  JobCode = '" & v_txtjob & "' "
				end if
				strFieldCons = strFieldCons & "JobCode, DESCRIPTION,"
				' strFieldCons2 = strFieldCons2 & "'' jobcode, '' DESCRIPTION,"
			    selectCons = selectCons & ",jobcode ,DESCRIPTION"
			    ' selectCons2 = selectCons2 & ",jobcode ,DESCRIPTION"
			    orderCons  = orderCons & ",jobcode ,DESCRIPTION" 
			end if
			
			   whereCons = wherecons & " And  businessunit_api.get_FCTYPE(fcba) = 'M' "
 if len(trim(v_txtyear)) > 0 then			
	Dim strdtl , strgrp ,cdgrp,cdmesin,cdgrpdesc  ,noline , subAmtgrpBI , GRAmtgrpBI, subAmtgrpSBI , GRAmtgrpSBI , grpDesc,subbudBI,subbudSBI,Subbudthn,GRbudBI,GRbudSBI,GRbudthn

			strdtl = " SELECT " & strFieldCons & "  SUM(BI) BI,SUM(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  FROM ( " & _
                     " SELECT *  FROM (" & _
                     "     SELECT data.ba fcba" & _
                     "            , ma.station_stationcode stationcode" & _
                     "            , st.fcname stationcode_description" & _
                     "            , data.machinecode" & _
                     "            , ma.description machinecode_description" & _
                     "            , data.jobcode" & _
                     "            , jj.fcname description" & _
                     "            , data.bi" & _
                     "            , data.sbi" & _
                     "            , data.budbi" & _
                     "            , data.budsbi" & _
                     "            , data.budthn" & _
                     "     FROM (" & _
                     "         SELECT YEAR period_year" & _
                     "                , MONTH period_month" & _
                     "                , fcba ba" & _
                     "                , NULL croptype" & _
                     "                , NULL sourcetype" & _
                     "                , NULL ownership" & _
                     "                , jobcode" & _
                     "                , targetcode machinecode" & _
                     "                , act_bi bi" & _
                     "                , bud_bi budbi" & _
                     "                , act_sbi sbi" & _
                     "                , bud_sbi budsbi" & _
                     "                , bud_annual budthn" & _
                     "         FROM ar_mill" & _
                     "         WHERE year = '" & v_txtyear & "'" & _
                     "               AND month = '" & v_txtmonth & "'" & _
                     "				 AND sourcetype not in ('GC','BM')" & _
                     "     ) data" & _
                     "     LEFT join job jj ON (jj.fcba = data.ba AND jj.fccode = data.jobcode)" & _
                     "     LEFT join machine ma ON (ma.fcba = data.ba AND ma.fccode = data.machinecode)" & _
                     "     LEFT join station st ON (st.fcba = data.ba AND st.fccode = ma.station_stationcode)" & _
					 " )" & _
					 " ) " & _
					 " GROUP BY  fcba,stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION ORDER BY stationcode "


             strdtl = "select GroupPlantation " & selectCons & " ,'BL' biaya, '' targettype," & _
					 " '' targetcode, '' detail, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & "group by GroupPlantation" & selectCons & " " & _
             		 " UNION ALL  select GroupPlantation " & selectCons & "" & _
             		 "		   , 'BTL' biaya" & _
			         "         , targettype" & _
			         "         , targetcode" & _
			         "         , detail" & _
			         "         , sum(bi) BI" & _
			         "         , sum(sbi) SBI" & _
			         "         , sum(budBI) budBI" & _
			         "         , sum(budSBI) budSBI" & _
			         "         , sum(budthn) budthn" & _
			         "  from ( " & _
                     "  select " & strFieldCons2 & "" & _
			         "         jobcode" & _
			         "         , description" & _
				     "         , targettype" & _
				     "         , targetcode" & _
				     "         , decode(targettype, 'GC', targetcode||' - '||GCMASTER_API.get_fcname(fcba, targettype, targetcode)" & _
				     "                            , 'BM', targetcode||' - '||BUILDING_API.get_FCNAME(targetcode, fcba)" & _
				     "                            , null) detail" & _
			         "         , sum(bi) BI" & _
			         "         , sum(sbi) SBI" & _
			         "         , sum(budBI) budBI" & _
			         "         , sum(budSBI) budSBI" & _
			         "         , sum(budthn) budthn" & _
			         "  from ( " & _
					 "         SELECT data.*, jj.fcname description" & _
					 "         FROM (" & _
					 "             SELECT fcba, status grp, jobcode, targettype, targetcode, SUM(act_bi) bi, SUM(bud_bi) budbi, SUM(act_sbi) sbi, SUM(bud_sbi) budsbi, SUM(bud_annual) budthn" & _
					 "             FROM ar_btl" & _
					 "             WHERE year = '" & v_txtyear & "'" & _
					 "                   AND month = '" & v_txtmonth & "'" & _
					 "             GROUP BY fcba, status, jobcode, targettype, targetcode" & _
					 "         ) data" & _
					 "         LEFT JOIN job jj ON (data.fcba = jj.fcba AND data.jobcode = jj.fccode)" & _
					 "	 ) " & _
					 "	group by fcba" & _
					 "           , jobcode" & _
					 "           , description" & _
					 "           , targettype" & _
					 "           , targetcode" & _
					 "           , decode(targettype, 'GC', targetcode||' - '||GCMASTER_API.get_fcname(fcba, targettype, targetcode)" & _
					 "                              , 'BM', targetcode||' - '||BUILDING_API.get_FCNAME(targetcode, fcba)" & _
					 "                              , null)" & _
					 "   )   " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' group by GroupPlantation, jobcode,description, targettype, targetcode, detail" & selectCons & " "
			
             ' Order by GroupPlantation " & orderCons  	

   			 
	
		' response.write strdtl
		Set rsha = Server.CreateObject("ADODB.Recordset")
		rsha.Open strdtl, connstr, adOpenStatic
		
		if cdbl(Palm_BI) = 0 then
		   Palm_BI = "1"
		end if
		if cdbl(Palm_SBI) = 0 then
		   Palm_SBI = "1"
		end if
		if cdbl(TBSprocess_BI) = 0 then
		   TBSprocess_BI = "1"
		end if
		if cdbl(TBSprocess_SBI) = 0 then
			TBSprocess_SBI = "1"
		end if
		subAmtgrpBI ="0"
		subAmtgrpSBI ="0"
		cdmesin =""
		
		' response.write "budget BI </br>" & cdbl(rsha("budBI"))
		' 	response.write "budget TBS BUd </br>" & cdbl(TBSbud_BI)
			
		' Dim rp_ha_BI,rp_ha_budBI,rp_kg_BI,rp_kg_budBI, rp_ha_SBI,rp_ha_budSBI ,rp_kg_SBI,rp_kg_budSBI 
		' Dim sub_rp_ha_BI,sub_rp_ha_budBI,sub_rp_kg_BI,sub_rp_kg_budBI, sub_rp_ha_SBI,sub_rp_ha_budSBI ,sub_rp_kg_SBI,sub_rp_kg_budSBI  	 
		' Dim SprevStn,SprevMachine,sprevjob,tempMachine,tempstation
		
		IF rsha.BOF  and rsha.EOF  then %>
		<tr>
			<td colspan="13" align="left">NO DATA AVAILABLE</td>
		</tr>
		   <% end if
				NO = 1
				rows = 1
				
		   		While Not rsha.eof
			   	if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				' GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
				' GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
				' GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))
				' GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
				' GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))	
						 
			%> 
				<tr class="<%=cClass%>">
					<% 	if rows = 1 then
						sqlrow = "  select count(biaya) rowsd  from ( " & strdtl & " ) where biaya = '" & rsha("biaya") & "'"
						' response.write sqlrow
						Set rsrow = Server.CreateObject("ADODB.Recordset")
						rsrow.Open sqlrow, connstr, adOpenStatic
						if not rsrow.EOF then
							temprows = CDbl(rsrow("rowsd"))
						else
							temprows = 0
						end if
					end if 
					%>
					<%if rows = 1 then%>
						<td align="center" Rowspan="<%=0+temprows%>"><%=rsha("biaya")%></td> 
					<%end if%>

					 <%
						IF v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
						   IF v_txtfcba <> "--ALL--" then					      
						      tempPlantation = "--SUMMARY--"
						      tempGroupRegion = "--SUMMARY--"
						      tempRegion = "--SUMMARY--"
						      tempCompany ="--SUMMARY--"
						      tempfcba = rsha("fcba")
						    ELSE
                              tempPlantation = "--SUMMARY--"
						      tempGroupRegion = "--SUMMARY--"
						      tempRegion = "--ALL--"
						      tempCompany ="--SUMMARY--"
						      tempfcba = rsha("fcba") 							
							END IF
                         END IF							
						
		                IF v_txtstation <> "--SUMMARY--" and v_txtstation<>"" then
						   tempstation = rsha("stationCode") 
						   If SprevStn <> rsha("stationCode") then %>
							   <td align="left" >
								   <%=rsha("StationCode")%> : <%=rsha("STATIONCODE_DESCRIPTION")%></td>   
						       <% SprevStn =rsha("StationCode")
						    Else %>
							   <td align="center" style="background-color: #7F8C8D;"></td> 
						 <% End If 
                        ELSE
						    tempstation =  "--SUMMARY--"
                        END IF 
					    IF v_txtmachine <> "--SUMMARY--" and v_txtmachine<>"" then
						   tempMachine = rsha("MACHINECODE")
						   If Sprevmachine <> rsha("MACHINECODE") then %>
						      <td align="center">
							      <a href="p_AccountReportMILL.asp?txtstation=<%=tempstation%>&txtmachine=<%=rsha("MACHINECODE")%>&txtjob=--ALL--&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 
								  <%=rsha("MACHINECODE")%>
							     </a>
							  </td>
							  <td align="left"><%=rsha("MACHINECODE_DESCRIPTION")%></td>
						      <% Sprevmachine=rsha("MACHINECODE")
						   Else %>
								<td align="center"></td>
								<td align="center"></td>
						 <% End If  
                        ELSE
                           tempMachine = "--SUMMARY--" 						
				        END IF	
                        IF v_txtjob <> "--SUMMARY--" and v_txtjob <> "" then
						      if tempfcba = "" then
							     tempfcba = "--ALL--"
						      end if
						   If Sprevjob <> rsha("jobcode") then 				       
						          %>						      
						      <td align="right">
							      <a href="p_AccountReportMILL.asp?txtstation=<%=tempstation%>&txtmachine=<%=tempMachine%>&txtjob=<%=rsha("jobcode")%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 
								  <%=rsha("jobcode")%>
							     </a>
							  </td>
							  <td align="left"><%=rsha("DESCRIPTION")%></td>
						      <% Sprevjob=rsha("jobcode")
						   Else %>
								<td align="center"></td>
								<td align="center"></td>
						 <% End If   
				        END IF						
						IF v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
						   if SprevPlantation <> rsha("Plantation") then %>
							   <td align="center">																																																																																									
								  <a href="p_AccountReportMILL.asp?txtstation=<%=SprevStn%>&txtmachine=<%=Sprevmachine%>&txtjob=<%=Sprevjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("Plantation")%>
								  </a>
							   </td>
							  <% SprevPlantation = rsha("Plantation") 	
							else	%>
							   <td align="center"></td>					
							  <% 
							end if 									   
						END IF 
						IF v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
						   if SprevGroupRegion <> rsha("GroupRegion") then %>
							  <td align="center">
								  <a href="p_AccountReportMILL.asp?txtstation=<%=SprevStn%>&txtmachine=<%=Sprevmachine%>&txtjob=<%=Sprevjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("GroupRegion")%>
								  </a>
							  </td>
							  <% SprevGroupRegion = rsha("GroupRegion")
							else	%>
							  <td align="center"></td>					
							  <% 
							end if									   
						END IF
						IF v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
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
							   IF SprevRegion <> rsha("Region") then %>
								  <td align="center">
									  <a href="p_AccountReportMILL.asp?txtstation=<%=SprevStn%>&txtmachine=<%=Sprevmachine%>&txtjob=<%=Sprevjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			<%=rsha("Region")%>
									  </a>
								  </td>
								  <% SprevRegion = rsha("Region")	
							   ELSE	%>
								  <td align="center"></td>					
								   <% 
							   END IF									   
						END IF
						IF v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
						 
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
						   tempCompany =rsha("Company") 
						   tempfcba = "--ALL--"
						   if SprevCompany <> rsha("Company") then %>
							 <td align="center">
								 <a href="p_AccountReportMILL.asp?txtstation=<%=SprevStn%>&txtmachine=<%=Sprevmachine%>&txtjob=<%=Sprevjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			  <%=rsha("Company")%>
								 </a>
							 </td>
							 <% SprevCompany = rsha("Company") 	
						   else	%>
							 <td align="center"></td>					
							 <% 
						   end if									   
						END IF
						IF v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
							  tempPlantation = rsha("Plantation")
						   else
							  tempPlantation = ""
						   end if
			               if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
							 tempGroupRegion = rsha("GroupRegion")
						   else
						     tempGroupRegion = ""
						   end if
						   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
						      tempRegion = rsha("Region") 
						   else
							  tempRegion = ""
						   end if
					       if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
							   tempCompany =rsha("Company")
						   else
							   tempCompany =""
						   end if							   
						   tempfcba = rsha("fcba")
						   if Sprevfcba <> rsha("fcba") then %>
							  <td align="center">
								  <a href="p_AccountReportMILL.asp?txtstation=<%=SprevStn%>&txtmachine=<%=Sprevmachine%>&txtjob=<%=Sprevjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>">
								     <%=rsha("fcba")%>
								  </a>
							  </td>
							  <% Sprevfcba = rsha("fcba") 	
						   else	%>
								<td align="center"></td>
                                <%								
						   end if 									   
					    END IF %>				     								   
						<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(palm_BI)),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)),2)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("budBI")),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)),2)%></td>
						<% if (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)) = 0 then %>
							<td align="right"><%=formatnumber((cdbl(rsha("budBI"))/cdbl(TBSbud_BI)) -(cdbl(rsha("BI"))/cdbl(TBSprocess_BI)) ,2)%></td>
						<% else %>
							<td align="right"><%=formatnumber(((cdbl(rsha("budBI"))/cdbl(TBSbud_BI)) -(cdbl(rsha("BI"))/cdbl(TBSprocess_BI)) )/(cdbl(rsha("budBI"))/cdbl(TBSbud_BI))*100,2)%></td>
						<% end if %>

						<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(palm_SBI)),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)),2)%></td>
						<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)),2)%></td>
						<% if (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) = 0 then %>
							<td align="right"><%=formatnumber((cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) -(cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)) ,2)%></td>
						<% else %>
							<td align="right"><%=formatnumber(((cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) -(cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)) )/(cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI))*100,2)%></td>
						<% end if %>
						<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)),2)%></td>
						<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)),2)%></td>				
				</tr>	 
	<%	if  rows = temprows then  
				GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
				GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))  
				Prod_BI_R = cdbl(GRAmtgrpBI)/cdbl(palm_BI)
				If cdbl(TBSprocess_BI) = 0 Then TBS_BI_R = 0 Else TBS_BI_R = cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)
				Prod_BI_A = cdbl(GRbudBI)/cdbl(PALMbud_BI)
				If cdbl(TBSbud_BI) = 0 Then TBS_BI_A = 0 Else TBS_BI_A = cdbl(GRbudBI)/cdbl(TBSbud_BI)
				PERC_BI_A = (cdbl(GRbudBI)/cdbl(TBSbud_BI)) - (cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI))
				PERC_BI_A_2 = cdbl(TBS_BI_A) - cdbl(TBS_BI_R) / cdbl(TBS_BI_A) *100

				GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
				GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
				Prod_SBI_R = cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)
				If cdbl(TBSprocess_SBI) = 0 Then TBS_SBI_R = 0 Else TBS_SBI_R = cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)
				Prod_SBI_A = cdbl(GRbudSBI)/cdbl(PALMbud_SBI)
				If cdbl(TBSbud_SBI) = 0 Then TBS_SBI_A = 0 Else TBS_SBI_A = cdbl(GRbudSBI)/cdbl(TBSbud_SBI)
				PERC_SBI_A =  (cdbl(GRbudSBI)/cdbl(TBSbud_SBI)) - ( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI))
				PERC_SBI_A_2 =cdbl(TBS_SBI_A) - cdbl(TBS_SBI_R) / cdbl(TBS_SBI_A)*100

				GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))	
				AT_Prod = cdbl(rsha("budThn"))/cdbl(PALMbud_THN)
				AT_TBS =  cdbl(rsha("budThn"))/cdbl(TBSbud_THN)
			rows = 1 	
	%>
	<tr>  		
		<!-- Bulan ini  -->
		<td style="background-color:#F4F4A0;" align="center" colspan="<%= nSpanTotal %>"><b>Sub Total <%=rsHa("biaya")%></td>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
		<% IF cdbl(palm_BI) > 0 then %>
		      <td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(Prod_BI_R),2)%></td>
		<% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %>
		<% IF cdbl(TBSprocess_BI) > 0 then %>
		      <td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(TBS_BI_R),2)%></td>	
		<% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudBI,2)%></td>
		<% IF cdbl(PALMbud_BI) > 0 then %> 
		      <td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(Prod_BI_A),2)%></td>
	    <% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %> 	
        <% IF cdbl(TBSbud_BI) > 0 then %> 		
		      <td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(TBS_BI_A),2)%></td>
		<% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %> 	
        <% If cdbl(TBSbud_BI) > 0 and  cdbl(TBSprocess_BI) > 0 then %>		
				<% If ( cdbl(TBS_BI_A)) = 0 then %>
					  <td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(PERC_BI_A) ,2)%></td>
				<% Else %>
					  <td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber( cdbl(PERC_BI_A_2),2)%></td>
				<% End If %>
		<% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %>		
		<!-- ENd Bulan Ini -->

		<!-- Sampai Dengan Bulan Ini -->
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
		<% IF cdbl(palm_SBI) > 0 then %>
		      <td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(Prod_SBI_R),2)%></td>
	    <% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %> 
        <% IF cdbl(TBSprocess_SBI) > 0 then %>		
		      <td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(TBS_SBI_R),2)%></td>
		<% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %> 	  
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudSBI,2)%></td>
		<% IF cdbl(PALMbud_SBI) > 0 then %>	
		      <td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(Prod_SBI_A),2)%></td>
	    <% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %> 	
        <% IF cdbl(TBSbud_SBI) > 0 then %>			
		      <td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(TBS_SBI_A),2)%></td>
		<% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %> 
        <% If cdbl(TBSbud_SBI) > 0 and  cdbl(TBSprocess_SBI) > 0 then %>			
				<% If ( cdbl(GRbudBI)/cdbl(TBSbud_SBI)) = 0 then %>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(PERC_SBI_A) ,2)%></td>
				<% Else %>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(cdbl(PERC_SBI_A_2),2)%></td>
				<% End If %>
		<% ELSE %>
		      <td  style="background-color:#F4F4A0" align="right"><b>0</td>
		<% END IF %> 		
		<!-- End Sampai Dengan Bulan Ini -->
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudthn,2)%></td>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(PALMbud_THN)),2)%></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(TBSbud_THN)),2)%></td>														
	</tr> 
	
		<%
			GRAmtgrpBI = cdbl("0") 
			GRbudBI  = cdbl("0") 	
			Prod_BI_R = cdbl("0")
			TBS_BI_R = cdbl("0")
			Prod_BI_A = cdbl("0")
			TBS_BI_A = cdbl("0")
			PERC_BI_A = cdbl("0")
			PERC_BI_A_2 = cdbl("0")

			GRAmtgrpSBI = cdbl("0")
			GRbudSBI = cdbl("0")
			Prod_SBI_R = cdbl("0")
			TBS_SBI_R = cdbl("0")
			Prod_SBI_A = cdbl("0")
			TBS_SBI_A = cdbl("0")

			GRbudthn  = cdbl("0")
			AT_Prod = cdbl("0")
			AT_TBS =  cdbl("0")
		else
		' cari sub total
			GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
			GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))  
			Prod_BI_R = cdbl(GRAmtgrpBI)/cdbl(palm_BI)
			If cdbl(TBSprocess_BI) = 0 Then TBS_BI_R = 0 Else TBS_BI_R = cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)
			Prod_BI_A = cdbl(GRbudBI)/cdbl(PALMbud_BI)
			If cdbl(TBSbud_BI) = 0 Then TBS_BI_A = 0 Else TBS_BI_A = cdbl(GRbudBI)/cdbl(TBSbud_BI)
			PERC_BI_A = (cdbl(GRbudBI)/cdbl(TBSbud_BI)) - (cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI))
			PERC_BI_A_2 = cdbl(TBS_BI_A) - cdbl(TBS_BI_R) / cdbl(TBS_BI_A) *100

			GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
			GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
			Prod_SBI_R = cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)
			If cdbl(TBSprocess_SBI) = 0 Then TBS_SBI_R = 0 Else TBS_SBI_R = cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)
			Prod_SBI_A = cdbl(GRbudSBI)/cdbl(PALMbud_SBI)
			If cdbl(TBSbud_SBI) = 0 Then TBS_SBI_A = 0 Else TBS_SBI_A = cdbl(GRbudSBI)/cdbl(TBSbud_SBI)
			PERC_SBI_A =  (cdbl(GRbudSBI)/cdbl(TBSbud_SBI)) - ( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI))
			PERC_SBI_A_2 =cdbl(TBS_SBI_A) - cdbl(TBS_SBI_R) / cdbl(TBS_SBI_A)*100
 			
 			GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))	
			AT_Prod = cdbl(rsha("budThn"))/cdbl(PALMbud_THN)
			AT_TBS =  cdbl(rsha("budThn"))/cdbl(TBSbud_THN)
			rows = rows + 1
		End if
							  
			' cari grand total 
				gt_GRAmtgrpBI = cdbl(gt_GRAmtgrpBI) + cdbl(rsha("BI")) 
				gt_Prod_BI_R = cdbl(gt_Prod_BI_R) + cdbl(rsha("BI")) /cdbl(palm_BI)
				If cdbl(TBSprocess_BI) = 0 Then gt_TBS_BI_R = 0 Else gt_TBS_BI_R  = cdbl(gt_TBS_BI_R ) + cdbl(rsha("BI")) / cdbl(TBSprocess_BI)
				gt_GRbudBI  = cdbl(gt_GRbudBI ) + cdbl(rsha("budBI"))
				gt_Prod_BI_A  = cdbl(gt_Prod_BI_A ) + cdbl(rsha("budBI"))/cdbl(PALMbud_BI)
				if cdbl(TBSbud_BI) = 0 Then gt_TBS_BI_A = 0 Else gt_TBS_BI_A = cdbl(gt_TBS_BI_A ) + cdbl(rsha("budBI"))/cdbl(TBSbud_BI)

				if (cdbl(gt_TBS_BI_R)) = 0 then  
					gt_Persent = (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)) - (cdbl(rsha("BI"))/cdbl(TBSprocess_BI))  
				else 
					gt_Persent = cdbl(gt_TBS_BI_A) - cdbl(gt_TBS_BI_R) / cdbl(gt_TBS_BI_A)*100
			 	End If

			 	gt_GRAmtgrpSBI = cdbl(gt_GRAmtgrpSBI) + cdbl(rsha("SBI")) 
				gt_Prod_SBI_R = cdbl(gt_Prod_SBI_R) + cdbl(rsha("SBI")) /cdbl(palm_SBI)
				If cdbl(TBSprocess_SBI) = 0 Then gt_TBS_SBI_R = 0 Else gt_TBS_SBI_R  = cdbl(gt_TBS_SBI_R ) + cdbl(rsha("SBI")) / cdbl(TBSprocess_SBI)
				gt_GRbudSBI  = cdbl(gt_GRbudSBI ) + cdbl(rsha("budSBI"))
				gt_Prod_SBI_A  = cdbl(gt_Prod_SBI_A ) + cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)
				if cdbl(TBSbud_SBI) = 0 Then gt_TBS_SBI_A = 0 Else gt_TBS_SBI_A = cdbl(gt_TBS_SBI_A ) + cdbl(rsha("budBI"))/cdbl(TBSbud_SBI)

				if (cdbl(gt_TBS_SBI_R)) = 0 then  
					gt_Persent2 = (cdbl(rsha("budBI"))/cdbl(TBSbud_SBI)) - (cdbl(rsha("BI"))/cdbl(TBSprocess_SBI))  
				else 
					gt_Persent2 = cdbl(gt_TBS_SBI_A) - cdbl(gt_TBS_SBI_R) / cdbl(gt_TBS_SBI_A)*100
			 	End If
				
				gt_GRbudthn = gt_GRbudthn + cdbl(rsha("budThn")) 
				gt_AT_Prod = gt_AT_Prod + cdbl(rsha("budThn"))/cdbl(PALMbud_THN)
				gt_AT_TBS = gt_AT_TBS + cdbl(rsha("budThn"))/cdbl(TBSbud_THN)
 

			rsha.MoveNext			
			Wend
		rsha.Close
		set rsha = Nothing
	%> 
<% End If %>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="<%= nSpanTotal %>"><font style="font-weight: bold;">GRAND TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_GRAmtgrpBI)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_Prod_BI_R)),2)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_TBS_BI_R)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_GRbudBI)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_Prod_BI_A)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_TBS_BI_A)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_Persent)),2)%></font></td>  

			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_GRAmtgrpSBI)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_Prod_SBI_R)),2)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_TBS_SBI_R)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_GRbudSBI)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_Prod_SBI_A)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_TBS_SBI_A)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_Persent2)),2)%></font></td>   

			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_GRbudthn)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_AT_Prod)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(gt_AT_TBS)),2)%></font></td>   
		</tr>

</tbody>
</table>

</form>
<!--#include file="bottom.Asp"-->