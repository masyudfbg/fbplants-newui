<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal,no
Dim v_txtyear,v_txtmonth,v_vJenisReport,v_txtJob,vJenisFactory,v_JenisFactory
Dim SprevCode,Sprevjobname
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempjob
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,rsJob

Dim varKernelProcess_BI : varKernelProcess_BI = 0
Dim varKernelProcess_SBI : varKernelProcess_SBI = 0
Dim varProdCPKO_BI : varProdCPKO_BI = 0
Dim varProdCPKO_SBI : varProdCPKO_SBI = 0
Dim varProdPKE_BI : varProdPKE_BI = 0
Dim varProdPKE_SBI : varProdPKE_SBI = 0
Dim varProduk_BI : varProduk_BI = 0
Dim varProduk_SBI : varProduk_SBI = 0

Dim RowRpPerProdAct_BI : RowRpPerProdAct_BI = 0
Dim RowRpPerKerAct_BI : RowRpPerKerAct_BI = 0
Dim RowRpPerProdBud_BI : RowRpPerProdBud_BI = 0
Dim RowRpPerKerBud_BI : RowRpPerKerBud_BI = 0
Dim RowVariance_BI : RowVariance_BI = 0

Dim RowRpPerProdAct_SBI : RowRpPerProdAct_SBI = 0
Dim RowRpPerKerAct_SBI : RowRpPerKerAct_SBI = 0
Dim RowRpPerProdBud_SBI : RowRpPerProdBud_SBI = 0
Dim RowRpPerKerBud_SBI : RowRpPerKerBud_SBI = 0
Dim RowVariance_SBI : RowVariance_SBI = 0
Dim RowRpPerProdBud_THN : RowRpPerProdBud_THN = 0
Dim RowRpPerKerBud_THN : RowRpPerKerBud_THN = 0

Dim TotalRpPerProdAct_BI : TotalRpPerProdAct_BI = 0
Dim TotalRpPerKerAct_BI : TotalRpPerKerAct_BI = 0
Dim TotalRpPerProdBud_BI : TotalRpPerProdBud_BI = 0
Dim TotalRpPerKerBud_BI : TotalRpPerKerBud_BI = 0
Dim TotalVariance_BI : TotalVariance_BI = 0

Dim TotalRpPerProdAct_SBI : TotalRpPerProdAct_SBI = 0
Dim TotalRpPerKerAct_SBI : TotalRpPerKerAct_SBI = 0
Dim TotalRpPerProdBud_SBI : TotalRpPerProdBud_SBI = 0
Dim TotalRpPerKerBud_SBI : TotalRpPerKerBud_SBI = 0
Dim TotalVariance_SBI : TotalVariance_SBI = 0

Dim TotalRpPerProdBud_THN : TotalRpPerProdBud_THN = 0
Dim TotalRpPerKerBud_THN : TotalRpPerKerBud_THN = 0

Dim GrandTotalRpPerProdAct_BI : GrandTotalRpPerProdAct_BI = 0
Dim GrandTotalRpPerKerAct_BI : GrandTotalRpPerKerAct_BI = 0
Dim GrandTotalRpPerProdBud_BI : GrandTotalRpPerProdBud_BI = 0
Dim GrandTotalRpPerKerBud_BI : GrandTotalRpPerKerBud_BI = 0
Dim GrandTotalVariance_BI : GrandTotalVariance_BI = 0

Dim GrandTotalRpPerProdAct_SBI : GrandTotalRpPerProdAct_SBI = 0
Dim GrandTotalRpPerKerAct_SBI : GrandTotalRpPerKerAct_SBI = 0
Dim GrandTotalRpPerProdBud_SBI : GrandTotalRpPerProdBud_SBI = 0
Dim GrandTotalRpPerKerBud_SBI : GrandTotalRpPerKerBud_SBI = 0
Dim GrandTotalVariance_SBI : GrandTotalVariance_SBI = 0

Dim GrandTotalRpPerProdBud_THN : GrandTotalRpPerProdBud_THN = 0
Dim GrandTotalRpPerKerBud_THN : GrandTotalRpPerKerBud_THN = 0

Dim TotalSubPerProdAct_BI : TotalSubPerProdAct_BI = 0
Dim TotalSubPerKerAct_BI : TotalSubPerKerAct_BI = 0
Dim TotalSubPerProdBud_BI : TotalSubPerProdBud_BI = 0
Dim TotalSubPerKerBud_BI : TotalSubPerKerBud_BI = 0
Dim TotalSubVariance_BI : TotalSubVariance_BI = 0

Dim TotalSubPerProdAct_SBI : TotalSubPerProdAct_SBI = 0
Dim TotalSubPerKerAct_SBI : TotalSubPerKerAct_SBI = 0
Dim TotalSubPerProdBud_SBI : TotalSubPerProdBud_SBI = 0
Dim TotalSubPerKerBud_SBI : TotalSubPerKerBud_SBI = 0
Dim TotalSubVariance_SBI : TotalSubVariance_SBI = 0

Dim TotalSubPerProdBud_THN : TotalSubPerProdBud_THN = 0
Dim TotalSubPerKerBud_THN : TotalSubPerKerBud_THN = 0


            v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	  = Request("txtPlantation")
			v_txtGroupRegion  	  = Request("txtGroupRegion")
			v_txtRegion  		  = Request("txtRegion")
			v_txtcompany  		  = Request("txtCompany")
			v_txtfcba 			  = Request("txtfcba")
			v_txtyear             = Request("txtYear")
			v_txtmonth            = Request("txtMonth")
			v_vJenisReport        = request("vJenisReport")
			v_txtjob			  = Request("txtjob")
			v_JenisFactory         = request("vJenisFactory")
            nSpanTotal = 5 
      Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut, " 

						   
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
%>
<html>

<head>
<title>ePlaS Reporting - Analisa Biaya Pabrik</title>
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
	<!--#include file="FormSelectNewSumP.html"-->
	<!--#include file="FormSelectHDRPlant.html"-->
	
	
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
<th colspan="25" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
    
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
		<h3>LAPORAN BIAYA LANGSUNG PABRIK</h3>
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
		 "  select t.fcba, sum(t.cageweight) TBS_PROCESS_BI, 0 TBS_PROCESS_SBI from SHIFTREPORT t " & _  
	     "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.productiondate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  " & _
		 "  union all" & _
		 "  select t.fcba, 0 TBS_PROCESS_BI, sum(t.cageweight) TBS_PROCESS_SBI from SHIFTREPORT t  " & _
		 "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.productiondate,'MM') <= lpad('" & v_txtmonth & "',2,0) " & _
         "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  )  group by fcba " 

strtbs = " select GroupPlantation" & _
         "        , sum(TBS_PROCESS_BI) TBS_PROCESS_BI" & _
         "        , sum(TBS_PROCESS_SBI) TBS_PROCESS_SBI" & _
         " from ( " & strtbs & " )" & _
         " " & whereCons & "" & _
         "     and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "', fcba) = 'P'" & _
         " group by GroupPlantation"
'response.write strtbs
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    TBSprocess_BI = rsha("TBS_PROCESS_BI")
	TBSprocess_SBI = rsha("TBS_PROCESS_SBI")
end if
rsha.close()

ProdCPO_BI =0
ProdCPO_SBI =0

		 
strtbs = " select " & strFieldCons & " sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_CPO_BI, 0 Produksi_CPO_SBI from MILL_PRODUCTION t  " & _
		 " where t.product_code='CPKO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union all" & _
		 " select t.fcba,0 Produksi_CPO_BI, sum(t.production_today) Produksi_CPO_SBI from MILL_PRODUCTION t " & _
		 " where t.product_code='CPKO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select GroupPlantation" & _
         "        , sum(Produksi_CPO_BI) Produksi_CPO_BI" & _
         "        , sum(Produksi_CPO_SBI) Produksi_CPO_SBI" & _
         " from ( " & strtbs & ")" & _
         " " & whereCons & "" & _
		 "     and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "', fcba) = 'P'" & _
         " group by GroupPlantation"
 'response.write strtbs			  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic

if Not rsha.eof then
    ProdCPO_BI = rsha("Produksi_CPO_BI")
	ProdCPO_SBI = rsha("Produksi_CPO_SBI")
end if
rsha.close()

ProdKER_BI =0
ProdKER_SBI =0
		 
strtbs = " select " & strFieldCons & "  sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_KER_BI, 0 Produksi_KER_SBI from MILL_PRODUCTION t  " & _
		 " where t.product_code='PKE'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union all" & _
		 " select t.fcba,0 Produksi_KER_BI, sum(t.production_today) Produksi_KER_SBI from MILL_PRODUCTION t " & _
		 " where t.product_code='PKE'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)  " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba " 		
strtbs = " select GroupPlantation" & _
         "        , sum(Produksi_KER_BI) Produksi_KER_BI" & _
         "        , sum(Produksi_KER_SBI) Produksi_KER_SBI" & _
         " from ( " & strtbs & " )" & _
         " " & whereCons & "" & _
		 "     and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "', fcba) = 'P'" & _
         " group by GroupPlantation"
' response.write strtbs	  
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
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='KER' " & _
		 "	 group by btm.fcba " & _
		 "	 union all" & _
		 "	 select btm.fcba , 0 TBSbudBI,sum(" & strbudHDR_SBI & ") tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='KER' " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='KER' " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,sum(" & strbudHDR_BI & ") cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='CPKO' " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,sum(" & strbudHDR_SBI & ") cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='CPKO' " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) cpobudthn,0 kerbudBI, 0 kerbudSBI , 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='CPKO' " & _
		 "	 group by btm.fcba " & _
		 "	  union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudthn,0 cpobudSBI,sum(" & strbudHDR_BI & ") kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='PKE' " & _
		 "	 group by btm.fcba " & _
		 "	  union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, sum(" & strbudHDR_SBI & ") kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='PKE' " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='PKE' " & _
		 "	 group by btm.fcba ) group by fcba " 

strtbs = " select GroupPlantation" & _
         "        , sum(TBSbudBI) TBSbudBI" & _
         "        , sum(tbsbudSBI) tbsbudSBI" & _
         "        , sum(tbsbudthn) tbsbudthn" & _
         "        , sum(cpobudBI) cpobudBI" & _
         "        , sum(cpobudSBI) cpobudSBI" & _
         "        , sum(cpobudthn) cpobudthn" & _
         "        , sum(kerbudBI) kerbudBI" & _
         "        , sum(kerbudSBI) kerbudSBI" & _
         "        , sum(kerbudthn) kerbudthn" & _
         " from ( " & strtbs & " )" & _
         " " & whereCons & "" & _
		 "     and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "', fcba) = 'P'" & _
         " group by GroupPlantation"
'response.write strtbs		  
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

%>	
</th>
</tr>
	<tr>
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>"></th>
		
		<th  style="background-color:#00A300;color:#fff;"  colspan="7">BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="7">S/D BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>"></th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">REALISASI</th>
		
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;"  >%</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">REALISASI</th>
		<th  style="background-color:#00A300;color:#fff;"  colspan="3">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;"  >%</th>
	</tr>
	<tr>
	    <% 
		'TBSbud_BI,TBSbud_SBI,TBSbud_THN,CPObud_BI,CPObud_SBI,CPObud_THN,KERbud_BI,KERbud_SBI,KERbud_THN,PALMbud_BI,PALMbud_SBI,PALMbud_THN 
		%>
		<th  style="background-color:#00A300;color:#fff;"   colspan="<%= 2 + nSpanTotal %>" align="right">KERNEL PROSES (KG)</th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSprocess_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSbud_BI,2)%></th>
		<%If Cdbl(TBSbud_BI) = 0 Then%>
			<%If cdbl(TBSprocess_BI) = 0 Then%>
				<%varKernelProcess_BI = 0%>
			<%Else%>
				<%varKernelProcess_BI = 100%>
			<%End If%>
		<%Else%>
				<%varKernelProcess_BI = Round((cdbl(TBSprocess_BI)-cdbl(TBSbud_BI))/(cdbl(TBSbud_BI))*100, 2)%>
		<%End If%>

		<th style="background-color:#00A300;color:#fff;" align="right" >
			<%=formatnumber(varKernelProcess_BI ,2)%>
		</th>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSprocess_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSbud_SBI,2)%></th>
		
		<%If Cdbl(TBSbud_SBI) = 0 Then%>
			<%If cdbl(TBSprocess_SBI) = 0 Then%>
				<%varKernelProcess_SBI = 0%>
			<%Else%>
				<%varKernelProcess_SBI = 100%>
			<%End If%>
		<%Else%>
				<%varKernelProcess_SBI = Round((cdbl(TBSprocess_SBI)-cdbl(TBSbud_SBI))/(cdbl(TBSbud_SBI))*100, 2)%>
		<%End If%>
		<th style="background-color:#00A300;color:#fff;" align="right" >
			<%=formatnumber(varKernelProcess_SBI ,2)%>
		</th>

		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSbud_THN,2)%></th>
		
		
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI CPKO (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdCPO_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(CPObud_BI,2)%></th>
		
		<%If Cdbl(CPObud_BI) = 0 Then%>
			<%If cdbl(ProdCPO_BI) = 0 Then%>
				<%varProdCPKO_BI = 0%>
			<%Else%>
				<%varProdCPKO_BI = 100%>
			<%End If%>
		<%Else%>
				<%varProdCPKO_BI = Round((cdbl(ProdCPO_BI)-cdbl(CPObud_BI))/(cdbl(CPObud_BI))*100, 2)%>
		<%End If%>
		<th style="background-color:#00A300;color:#fff;" align="right" >
			<%=formatnumber(varProdCPKO_BI ,2)%>
		</th>

		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(ProdCPO_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(CPObud_SBI,2)%></th>
		
		<%If Cdbl(CPObud_SBI) = 0 Then%>
			<%If cdbl(ProdCPO_SBI) = 0 Then%>
				<%varProdCPKO_SBI = 0%>
			<%Else%>
				<%varProdCPKO_SBI = 100%>
			<%End If%>
		<%Else%>
				<%varProdCPKO_SBI = Round((cdbl(ProdCPO_SBI)-cdbl(CPObud_SBI))/(cdbl(CPObud_SBI))*100, 2)%>
		<%End If%>
		<th style="background-color:#00A300;color:#fff;" align="right" >
			<%=formatnumber(varProdCPKO_SBI ,2)%>
		</th>
		
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(CPObud_THN,2)%></th>
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI PKE (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdKER_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(KERbud_BI,2)%></th>
		
		<%If Cdbl(KERbud_BI) = 0 Then%>
			<%If cdbl(ProdKER_BI) = 0 Then%>
				<%varProdPKE_BI = 0%>
			<%Else%>
				<%varProdPKE_BI = 100%>
			<%End If%>
		<%Else%>
				<%varProdPKE_BI = Round((cdbl(ProdKER_BI)-cdbl(KERbud_BI))/(cdbl(KERbud_BI))*100, 2)%>
		<%End If%>
		<th style="background-color:#00A300;color:#fff;" align="right" >
			<%=formatnumber(varProdPKE_BI ,2)%>
		</th>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(ProdKER_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(KERbud_SBI,2)%></th>
		
		<%If Cdbl(KERbud_SBI) = 0 Then%>
			<%If cdbl(ProdKER_SBI) = 0 Then%>
				<%varProdPKE_SBI = 0%>
			<%Else%>
				<%varProdPKE_SBI = 100%>
			<%End If%>
		<%Else%>
				<%varProdPKE_SBI = Round((cdbl(ProdKER_SBI)-cdbl(KERbud_SBI))/(cdbl(KERbud_SBI))*100, 2)%>
		<%End If%>
		<th style="background-color:#00A300;color:#fff;" align="right" >
			<%=formatnumber(varProdPKE_SBI ,2)%>
		</th>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(KERbud_THN,2)%></th>
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">PRODUK (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(Palm_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(PALMbud_BI,2)%></th>
		
		<%If Cdbl(Palmbud_BI) = 0 Then%>
			<%If cdbl(Palm_BI) = 0 Then%>
				<%varProduk_BI = 0%>
			<%Else%>
				<%varProduk_BI = 100%>
			<%End If%>
		<%Else%>
				<%varProduk_BI = Round((cdbl(Palm_BI)-cdbl(Palmbud_BI))/(cdbl(Palmbud_BI))*100, 2)%>
		<%End If%>
		<th style="background-color:#00A300;color:#fff;" align="right" >
			<%=formatnumber(varProduk_BI ,2)%>
		</th>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(Palm_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(PALMbud_SBI,2)%></th>
		
		<%If Cdbl(PALMbud_SBI) = 0 Then%>
			<%If cdbl(Palm_SBI) = 0 Then%>
				<%varProduk_SBI = 0%>
			<%Else%>
				<%varProduk_SBI = 100%>
			<%End If%>
		<%Else%>
				<%varProduk_SBI = Round((cdbl(Palm_SBI)-cdbl(PALMbud_SBI))/(cdbl(PALMbud_SBI))*100, 2)%>
		<%End If%>
		<th style="background-color:#00A300;color:#fff;" align="right" >
			<%=formatnumber(varProduk_SBI ,2)%>
		</th>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(PALMbud_THN,2)%></th>
	</tr>
	
	<tr>
	  
	    <th style="background-color:#00A300;color:#fff;" rowspan="2" >Kode</th>
		<%  if len(v_vJenisReport )<1 then %>
		      <th style="background-color:#00A300;color:#fff;" rowspan="2" >Nama Stasiun</th>
		<% end if %>
		<% if   v_vJenisReport  = "Job"  then %>
			<th style="background-color:#00A300;color:#fff;" rowspan="2" >Nama Pekerjaan</th>
		<% elseif v_vJenisReport  = "Group"  then %>
			<th style="background-color:#00A300;color:#fff;" rowspan="2" >Nama Stasiun</th>
		<% elseif v_vJenisReport  = "Mechine"  then %>
			<th style="background-color:#00A300;color:#fff;" rowspan="2" >Nama Mesin</th>
		<% elseif v_vJenisReport  = "Combine"  then %>	
			<th style="background-color:#00A300;color:#fff;" rowspan="2" >Nama Pekerjaan</th>
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
		<th style="background-color:#00A300;color:#fff;"  >Kernel</th>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >Kernel</th>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >Kernel</th>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >Kernel</th>
		
		<th style="background-color:#00A300;color:#fff;"  >Produk</th>
		<th style="background-color:#00A300;color:#fff;"  >Kernel</th>
	</tr>
</thead>
<tbody>	
<%

if v_txtjob <> "--SUMMARY--"  and v_txtjob <> "--ALL--" and v_txtjob <> "" then
	whereCons = wherecons & " And jobcode = '" & v_txtjob & "' "
end if

whereCons = wherecons & " And businessunit_api.get_FCTYPE(fcba) = 'P' "
			
	Dim strdtl , strgrp ,cdgrp,cdmesin,cdgrpdesc  ,noline , subAmtgrpBI , GRAmtgrpBI, subAmtgrpSBI , GRAmtgrpSBI , grpDesc,subbudBI,subbudSBI,Subbudthn,GRbudBI,GRbudSBI,GRbudthn

	if     v_vJenisReport  = "Group" or len(v_vJenisReport ) < 1  then
			strdtl = " select " & strFieldCons & " jobcode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from  ( " & _
					 " 	select t.ba fcba,t.stationcode jobcode, t.STATIONCODE_DESCRIPTION DESCRIPTION, sum(t.amount) BI, 0 sbi, " & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.ba," & v_txtyear  & ",1,1,'','','','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.ba," & v_txtyear  & ",1,1,'','','','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.ba," & v_txtyear  & ",1,1,'','','','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY'))  budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by  t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION  " & _
					 "	union all" & _
					 "	select t.ba fcba,t.stationcode jobcode, t.STATIONCODE_DESCRIPTION DESCRIPTION, 0 BI, sum(t.amount) SBI ,0 budBI,0 budSBI,0 budthn  " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by  t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION  " & _
					 "	 ) group by fcba,jobcode,DESCRIPTION  " 	
            
              strdtl = "select GroupPlantation " & selectCons & ", jobcode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by jobcode,DESCRIPTION,GroupPlantation" & selectCons & "  Order by jobcode,DESCRIPTION,GroupPlantation" & orderCons  	
			
	elseif  v_vJenisReport  = "Job"  then  	
           	strdtl = " SELECT " & strFieldCons & " stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn FROM ( " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION DESCRIPTION , sum(t.amount) BI, 0 SBI ," & _
					 "  0 budBI, " & _
                     "  0 budSBI, " & _
                     "  0 budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)   " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "	UNION ALL " & _
					 "	select distinct t.fcba,t.stationcode, station_api.get_StationName(t.stationcode,t.fcba) STATIONCODE_DESCRIPTION , t.JobCode,job_api.get_jobcode_name(t.jobcode,t.fcba) DESCRIPTION , 0 BI, 0 SBI ," & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,'',t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||'" & v_txtyear  & "','DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,'',t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||" & v_txtyear  & ",'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,'',t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||" & v_txtyear  & ",'DD/MM/YYYY'))  budthn " & _
					 "	from budget_transaction_mill t where t.budget_year = '" & v_txtyear  & "'    " & _
					 "	and t.targettype ='MA' and instr('11-12-13-14-15-16-17-18-19', t.stationcode) <> 0" & _
					 "	UNION ALL " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION DESCRIPTION , 0 BI, sum(t.amount) SBI ,0 budBI,0 budSBI,0 budthn  " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "  ) " & _
					 "	GROUP BY  fcba,stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION  "
			strdtl = "select GroupPlantation " & selectCons & ", stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION,GroupPlantation" & selectCons & "  Order by stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION,GroupPlantation" & orderCons  	
		 
	elseif  v_vJenisReport  = "Mechine"  then
			strdtl = " SELECT " & strFieldCons & " stationcode,STATIONCODE_DESCRIPTION,MACHINECODE jobcode,DESCRIPTION, SUM(BI) BI,SUM(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  FROM ( " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.MACHINECODE,t.MACHINECODE_DESCRIPTION DESCRIPTION,t.JobCode,t.ACCOUNTCODE_DESCRIPTION  , sum(t.amount) BI, 0 SBI ," & _
					 "  0 budBI, " & _
                     "  0 budSBI, " & _
                     "  0 budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION, t.MACHINECODE,t.MACHINECODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "	UNION ALL " & _
					 "	select distinct t.fcba,t.stationcode, station_api.get_StationName(t.stationcode,t.fcba) STATIONCODE_DESCRIPTION ,t.machinecode,mechine_api.get_MechineName(t.machinecode,t.fcba) MACHINECODE_DESCRIPTION, t.JobCode,job_api.get_jobcode_name(t.jobcode,t.fcba) DESCRIPTION , 0 BI, 0 SBI ," & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,t.MACHINECODE,t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||'" & v_txtyear  & "','DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,t.MACHINECODE,t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||" & v_txtyear  & ",'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,t.MACHINECODE,t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||" & v_txtyear  & ",'DD/MM/YYYY'))  budthn " & _
					 "	from budget_transaction_mill t where t.budget_year = '" & v_txtyear  & "'    " & _
					 "	and t.targettype ='MA' and instr('11-12-13-14-15-16-17-18-19', t.stationcode) <> 0 " & _
					 "	UNION ALL " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.MACHINECODE,t.MACHINECODE_DESCRIPTION DESCRIPTION, t.JobCode,t.ACCOUNTCODE_DESCRIPTION , 0 BI, sum(t.amount) SBI,0 budBI,0 budSBI,0 budthn  " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)   " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION, t.MACHINECODE,t.MACHINECODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "	 ) " & _
					 "	GROUP BY  fcba,stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,DESCRIPTION " 	
			
             strdtl = "select GroupPlantation " & selectCons & ", stationcode,STATIONCODE_DESCRIPTION,jobcode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION,GroupPlantation" & selectCons & "  Order by stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION,GroupPlantation" & orderCons  	
			 'response.write v_vJenisReport & " - " & strdtl
			 
	elseif  v_vJenisReport  = "Combine"  then
			strdtl = " SELECT " & strFieldCons & " stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION, SUM(BI) BI,SUM(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  FROM ( " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.MACHINECODE,t.MACHINECODE_DESCRIPTION ,t.JobCode,t.ACCOUNTCODE_DESCRIPTION  DESCRIPTION, sum(t.amount) BI, 0 SBI ," & _
					 "  0 budBI, " & _
                     "  0 budSBI, " & _
                     "  0 budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION, t.MACHINECODE,t.MACHINECODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "	UNION ALL " & _
					  "	select distinct t.fcba,t.stationcode, station_api.get_StationName(t.stationcode,t.fcba) STATIONCODE_DESCRIPTION ,t.machinecode,mechine_api.get_MechineName(t.machinecode,t.fcba) MACHINECODE_DESCRIPTION, t.JobCode,job_api.get_jobcode_name(t.jobcode,t.fcba) DESCRIPTION , 0 BI, 0 SBI ," & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,t.MACHINECODE,t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||'" & v_txtyear  & "','DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,t.MACHINECODE,t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||" & v_txtyear  & ",'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.fcba," & v_txtyear  & ",1,1,'',t.JobCode,'','',t.targettype,t.MACHINECODE,t.stationcode,to_date('01/'||lpad('" & v_txtmonth & "',2,0)||'/'||" & v_txtyear  & ",'DD/MM/YYYY'))  budthn " & _
					 "	from budget_transaction_mill t where t.budget_year = '" & v_txtyear  & "'    " & _
					 "	and t.targettype ='MA' and instr('11-12-13-14-15-16-17-18-19', t.stationcode) <> 0 " & _
					 "	UNION ALL " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.MACHINECODE,t.MACHINECODE_DESCRIPTION ,t.JobCode,t.ACCOUNTCODE_DESCRIPTION  DESCRIPTION, 0 BI, sum(t.amount) SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION, t.MACHINECODE,t.MACHINECODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION " & _ 
					 "	 ) " & _
					 "	GROUP BY  fcba,stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION ORDER BY stationcode " 	
             strdtl = "select GroupPlantation " & selectCons & ", stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION , sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION ,GroupPlantation" & selectCons & "  Order by stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION ,GroupPlantation" & orderCons  	
   					 
	end if		 
			 
				
	if len(v_vJenisReport ) >  1 then
	
	
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    
	subAmtgrpBI ="0"
	subAmtgrpSBI ="0"
	cdmesin =""
	
	 Dim rp_ha_BI,rp_ha_budBI,rp_kg_BI,rp_kg_budBI, rp_ha_SBI,rp_ha_budSBI ,rp_kg_SBI,rp_kg_budSBI 
     Dim sub_rp_ha_BI,sub_rp_ha_budBI,sub_rp_kg_BI,sub_rp_kg_budBI, sub_rp_ha_SBI,sub_rp_ha_budSBI ,sub_rp_kg_SBI,sub_rp_kg_budSBI  	 

	if Not rsha.eof then
		  while Not rsha.eof
		  
	
         IF  v_vJenisReport  = "Group" then 
		     GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
			 GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
			 GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))
			 GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
			 GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))
			 	 
			 
			 %>
			 
	      <tr>  
		      <%
				        tempjob = rsha("JobCode")
						tempPlantation = "--ALL--"
						tempGroupRegion = "--SUMMARY--"
						tempRegion = "--SUMMARY--"
						tempCompany ="--SUMMARY--"
						tempfcba = "--SUMMARY--"
						If SprevCode <> rsha("JobCode") then %>
				           <td align="center">
						          <a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 
						          <%=rsha("JobCode")%>
								  </a>
						   </td>
					    <% Else %>
								 <td align="center"></td>
						<% End If %>	
						<% If Sprevjobname <> rsha("DESCRIPTION") then %>
				                 <td><%=rsha("DESCRIPTION")%></td>
						<% Else %>
								 <td align="center"></td>
						<% End If %>	   
				<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
				                     tempjob = rsha("JobCode")
									 tempPlantation = rsha("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 IF SprevPlantation <> rsha("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 		<%=rsha("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsha("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
							         tempjob = rsha("JobCode")
									 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										  tempPlantation = rsha("Plantation")
									 else
										  tempPlantation = ""
									 end if
									 tempGroupRegion = rsha("GroupRegion")
									 tempRegion = "--ALL--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"

									 IF SprevGroupRegion <> rsha("GroupRegion") then %>
										   <td align="center">
												<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 		<%=rsha("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsha("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
									   tempjob = rsha("JobCode")
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
													<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 			<%=rsha("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsha("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
									   tempjob = rsha("JobCode")
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

										 IF SprevCompany <> rsha("Company") then %>
											   <td align="center">
												    <a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 			  <%=rsha("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsha("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
									   tempjob = rsha("JobCode")
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
									   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
										   tempCompany =rsha("Company")
									   else
										   tempCompany =""
									   end if							   
									   tempfcba = rsha("fcba")

										 IF Sprevfcba <> rsha("fcba") then %>
											   <td align="center">
													<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
					<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
					<%If Cdbl(palm_BI) = 0 Then%>
						<%RowRpPerProdAct_BI = 0%>
					<%Else%>
						<%RowRpPerProdAct_BI = Round( (cdbl(rsha("BI"))/cdbl(palm_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdAct_BI, 2)%></td>
					
					<%If Cdbl(TBSprocess_BI) = 0 Then%>
						<%RowRpPerKerAct_BI = 0%>
					<%Else%>
						<%RowRpPerKerAct_BI = Round( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerAct_BI, 2)%></td>

					<td align="right"><%=formatnumber( cdbl(rsha("budBI")),2)%></td>
					
					<%If Cdbl(PALMbud_BI) = 0 Then%>
						<%RowRpPerProdBud_BI = 0%>
					<%Else%>
						<%RowRpPerProdBud_BI = Round( (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_BI, 2)%></td>
					
					<%If Cdbl(TBSbud_BI) = 0 Then%>
						<%RowRpPerKerBud_BI = 0%>
					<%Else%>
						<%RowRpPerKerBud_BI = Round( (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_BI, 2)%></td>
					
					<%If RowRpPerKerBud_BI = 0 then%>
						<%If RowRpPerKerAct_BI = 0 Then%>
							<%RowVariance_BI = 0%>
						<%Else%>
							<%RowVariance_BI = -100%>
						<%End If%>
					<%Else%>
						<%RowVariance_BI = Round( ((RowRpPerKerBud_BI - RowRpPerKerAct_BI) / RowRpPerKerBud_BI) * 100, 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowVariance_BI, 2)%></td>
					
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<%If Cdbl(palm_SBI) = 0 Then%>
						<%RowRpPerProdAct_SBI = 0%>
					<%Else%>
						<%RowRpPerProdAct_SBI = Round( (cdbl(rsha("SBI"))/cdbl(palm_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdAct_SBI, 2)%></td>
					
					<%If Cdbl(TBSprocess_SBI) = 0 Then%>
						<%RowRpPerKerAct_SBI = 0%>
					<%Else%>
						<%RowRpPerKerAct_SBI = Round( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerAct_SBI, 2)%></td>
					
					<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
					
					<%If Cdbl(PALMbud_SBI) = 0 Then%>
						<%RowRpPerProdBud_SBI = 0%>
					<%Else%>
						<%RowRpPerProdBud_SBI = Round( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_SBI, 2)%></td>
					
					<%If Cdbl(TBSbud_SBI) = 0 Then%>
						<%RowRpPerKerBud_SBI = 0%>
					<%Else%>
						<%RowRpPerKerBud_SBI = Round( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_SBI, 2)%></td>
					
					<%If RowRpPerKerBud_SBI = 0 then%>
						<%If RowRpPerKerAct_SBI = 0 Then%>
							<%RowVariance_SBI = 0%>
						<%Else%>
							<%RowVariance_SBI = -100%>
						<%End If%>
					<%Else%>
						<%RowVariance_SBI = Round( ((RowRpPerKerBud_SBI - RowRpPerKerAct_SBI) / RowRpPerKerBud_SBI) * 100, 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowVariance_SBI, 2)%></td>

					<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
					
					<%If Cdbl(PALMbud_THN) = 0 Then%>
						<%RowRpPerProdBud_THN = 0%>
					<%Else%>
						<%RowRpPerProdBud_THN = Round( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_THN, 2)%></td>
					
					<%If Cdbl(TBSbud_THN) = 0 Then%>
						<%RowRpPerKerBud_THN = 0%>
					<%Else%>
						<%RowRpPerKerBud_THN = Round( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_THN, 2)%></td>
				
		    </tr>		
	   <% ELSE 
	            
				   
				   GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
			       GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI"))
				   GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))
				   GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
			       GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))
				   
				   IF  cdgrp = rsha("stationcode") then 
				         subAmtgrpSBI = cdbl(subAmtgrpSBI) + cdbl(rsha("SBI"))
						 subAmtgrpBI = cdbl(subAmtgrpBI) + cdbl(rsha("BI"))						 
						 subbudBI  = cdbl(subbudBI) + cdbl(rsha("budBI"))
						 subbudSBI = cdbl(subbudSBI) + cdbl(rsha("budSBI"))
						 subbudthn  = cdbl(subbudthn) + cdbl(rsha("budthn"))                                                 %>	

                         <% IF  v_vJenisReport  = "Combine"  then 
					          if cdmesin <> rsha("MACHINECODE") then %>
								  <tr>  								
										<td colspan="<%= 19 + nSpanTotal %>" >&nbsp;&nbsp;&nbsp;<b><%=rsha("MACHINECODE")%>. <%=rsha("MACHINECODE_DESCRIPTION")%></td>
								  </tr> 
							<% end if 
								  
					      END IF %>



																		  
						<tr>  
							<%
				        tempjob = rsha("JobCode")
						tempPlantation = "--ALL--"
						tempGroupRegion = "--SUMMARY--"
						tempRegion = "--SUMMARY--"
						tempCompany ="--SUMMARY--"
						tempfcba = "--SUMMARY--"
						If SprevCode <> rsha("JobCode") then %>
				           <td align="center">
						          <a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 
						          <%=rsha("JobCode")%>
								  </a>
						   </td>
					    <% Else %>
								 <td align="center"></td>
						<% End If %>	
						<% If Sprevjobname <> rsha("DESCRIPTION") then %>
				                 <td><%=rsha("DESCRIPTION")%></td>
						<% Else %>
								 <td align="center"></td>
						<% End If %>	   
				<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
				                     tempjob = rsha("JobCode")
									 tempPlantation = rsha("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 IF SprevPlantation <> rsha("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 		<%=rsha("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsha("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
							         tempjob = rsha("JobCode")
									 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										  tempPlantation = rsha("Plantation")
									 else
										  tempPlantation = ""
									 end if
									 tempGroupRegion = rsha("GroupRegion")
									 tempRegion = "--ALL--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"

									 IF SprevGroupRegion <> rsha("GroupRegion") then %>
										   <td align="center">
												<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 		<%=rsha("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsha("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
									   tempjob = rsha("JobCode")
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
													<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 			<%=rsha("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsha("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
									   tempjob = rsha("JobCode")
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

										 IF SprevCompany <> rsha("Company") then %>
											   <td align="center">
												    <a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 			  <%=rsha("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsha("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
									   tempjob = rsha("JobCode")
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
									   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
										   tempCompany =rsha("Company")
									   else
										   tempCompany =""
									   end if							   
									   tempfcba = rsha("fcba")

										 IF Sprevfcba <> rsha("fcba") then %>
											   <td align="center">
													<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
					<%If Cdbl(palm_BI) = 0 Then%>
						<%RowRpPerProdAct_BI = 0%>
					<%Else%>
						<%RowRpPerProdAct_BI = Round( (cdbl(rsha("BI"))/cdbl(palm_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdAct_BI, 2)%></td>
					
					<%If Cdbl(TBSprocess_BI) = 0 Then%>
						<%RowRpPerKerAct_BI = 0%>
					<%Else%>
						<%RowRpPerKerAct_BI = Round( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerAct_BI, 2)%></td>

					<td align="right"><%=formatnumber( cdbl(rsha("budBI")),2)%></td>
					
					<%If Cdbl(PALMbud_BI) = 0 Then%>
						<%RowRpPerProdBud_BI = 0%>
					<%Else%>
						<%RowRpPerProdBud_BI = Round( (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_BI, 2)%></td>
					
					<%If Cdbl(TBSbud_BI) = 0 Then%>
						<%RowRpPerKerBud_BI = 0%>
					<%Else%>
						<%RowRpPerKerBud_BI = Round( (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_BI, 2)%></td>
					
					<%If RowRpPerKerBud_BI = 0 then%>
						<%If RowRpPerKerAct_BI = 0 Then%>
							<%RowVariance_BI = 0%>
						<%Else%>
							<%RowVariance_BI = -100%>
						<%End If%>
					<%Else%>
						<%RowVariance_BI = Round( ((RowRpPerKerBud_BI - RowRpPerKerAct_BI) / RowRpPerKerBud_BI) * 100, 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowVariance_BI, 2)%></td>
					
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<%If Cdbl(palm_SBI) = 0 Then%>
						<%RowRpPerProdAct_SBI = 0%>
					<%Else%>
						<%RowRpPerProdAct_SBI = Round( (cdbl(rsha("SBI"))/cdbl(palm_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdAct_SBI, 2)%></td>
					
					<%If Cdbl(TBSprocess_SBI) = 0 Then%>
						<%RowRpPerKerAct_SBI = 0%>
					<%Else%>
						<%RowRpPerKerAct_SBI = Round( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerAct_SBI, 2)%></td>
					
					<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
					
					<%If Cdbl(PALMbud_SBI) = 0 Then%>
						<%RowRpPerProdBud_SBI = 0%>
					<%Else%>
						<%RowRpPerProdBud_SBI = Round( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_SBI, 2)%></td>
					
					<%If Cdbl(TBSbud_SBI) = 0 Then%>
						<%RowRpPerKerBud_SBI = 0%>
					<%Else%>
						<%RowRpPerKerBud_SBI = Round( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_SBI, 2)%></td>
					
					<%If RowRpPerKerBud_SBI = 0 then%>
						<%If RowRpPerKerAct_SBI = 0 Then%>
							<%RowVariance_SBI = 0%>
						<%Else%>
							<%RowVariance_SBI = -100%>
						<%End If%>
					<%Else%>
						<%RowVariance_SBI = Round( ((RowRpPerKerBud_SBI - RowRpPerKerAct_SBI) / RowRpPerKerBud_SBI) * 100, 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowVariance_SBI, 2)%></td>

					<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
					
					<%If Cdbl(PALMbud_THN) = 0 Then%>
						<%RowRpPerProdBud_THN = 0%>
					<%Else%>
						<%RowRpPerProdBud_THN = Round( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_THN, 2)%></td>
					
					<%If Cdbl(TBSbud_THN) = 0 Then%>
						<%RowRpPerKerBud_THN = 0%>
					<%Else%>
						<%RowRpPerKerBud_THN = Round( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_THN, 2)%></td>
						</tr>		
				<% ELSE  
				         
				          If  cdgrp <> "" then  %>
								<tr>  						
									<td style="background-color:#F4F4A0" colspan="<%= 2 + nSpanTotal %>"><b>Total &nbsp;<%=cdgrpdesc%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpBI,2)%></td>
								<%If Cdbl(palm_BI) = 0 Then%>
									<%TotalSubPerProdAct_BI = 0%>
								<%Else%>
									<%TotalSubPerProdAct_BI = Round( (cdbl(subAmtgrpBI)/cdbl(palm_BI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdAct_BI, 2)%></b></td>
								
								<%If Cdbl(palm_BI) = 0 Then%>
									<%TotalSubPerKerAct_BI = 0%>
								<%Else%>
									<%TotalSubPerKerAct_BI = Round( (cdbl(subAmtgrpBI)/cdbl(TBSprocess_BI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerAct_BI, 2)%></b></td>
								
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudBI,2)%></td>
								<%If Cdbl(PALMbud_BI) = 0 Then%>
									<%TotalSubPerProdBud_BI = 0%>
								<%Else%>
									<%TotalSubPerProdBud_BI = Round( (cdbl(subbudBI)/cdbl(PALMbud_BI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdBud_BI, 2)%></b></td>
								
								<%If Cdbl(TBSbud_BI) = 0 Then%>
									<%TotalSubPerKerBud_BI = 0%>
								<%Else%>
									<%TotalSubPerKerBud_BI = Round( (cdbl(subbudBI)/cdbl(TBSbud_BI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerBud_BI, 2)%></b></td>
								
								<%If TotalSubPerKerAct_BI = 0 then%>
									<%If subbudBI = 0 Then%>
										<%If subAmtgrpBI = 0 Then%>
											<%TotalSubVariance_BI = 0%>
										<%Else%>
											<%TotalSubVariance_BI = -100%>
										<%End If%>
									<%Else%>
										<%TotalSubVariance_BI = Round( ((subbudBI - subAmtgrpBI) / subbudBI) * 100, 2)%>
									<%End If%>
								<%Else%>
									<% If TotalSubPerKerBud_BI = 0 Then%>
										<%TotalSubVariance_BI = -100%>
									<%Else%>
										<%TotalSubVariance_BI = Round( ((TotalSubPerKerBud_BI - TotalSubPerKerAct_BI) / TotalSubPerKerBud_BI) * 100, 2)%>
									<%End If%>									
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubVariance_BI, 2)%></b></td>
								
								
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpSBI,2)%></td>
								<%If Cdbl(palm_SBI) = 0 Then%>
									<%TotalSubPerProdAct_SBI = 0%>
								<%Else%>
									<%TotalSubPerProdAct_SBI = Round( (cdbl(subAmtgrpSBI)/cdbl(palm_SBI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdAct_SBI, 2)%></b></td>
								
								<%If Cdbl(TBSprocess_SBI) = 0 Then%>
									<%TotalSubPerKerAct_SBI = 0%>
								<%Else%>
									<%TotalSubPerKerAct_SBI = Round( (cdbl(subAmtgrpSBI)/cdbl(TBSprocess_SBI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerAct_SBI, 2)%></b></td>
								
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudSBI,2)%></td>
								<%If Cdbl(PALMbud_SBI) = 0 Then%>
									<%TotalSubPerProdBud_SBI = 0%>
								<%Else%>
									<%TotalSubPerProdBud_SBI = Round( (cdbl(subbudSBI)/cdbl(PALMbud_SBI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdBud_SBI, 2)%></b></td>
								
								<%If Cdbl(TBSbud_SBI) = 0 Then%>
									<%TotalSubPerKerBud_SBI = 0%>
								<%Else%>
									<%TotalSubPerKerBud_SBI = Round( (cdbl(subbudSBI)/cdbl(TBSbud_SBI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerBud_SBI, 2)%></b></td>
								
								<%If TotalSubPerKerAct_SBI = 0 then%>
									<%If subbudSBI = 0 Then%>
										<%If subAmtgrpSBI = 0 Then%>
											<%TotalSubVariance_BI = 0%>
										<%Else%>
											<%TotalSubVariance_BI = -100%>
										<%End If%>
									<%Else%>
										<%TotalSubVariance_BI = Round( ((subbudSBI - subAmtgrpSBI) / subbudSBI) * 100, 2)%>
									<%End If%>
								<%Else%>
									<% If TotalSubPerKerBud_SBI = 0 Then%>
										<%TotalSubVariance_BI = -100%>
									<%Else%>
										<%TotalSubVariance_BI = Round( ((TotalSubPerKerBud_SBI - TotalSubPerKerAct_SBI) / TotalSubPerKerBud_SBI) * 100, 2)%>
									<%End If%>									
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubVariance_BI, 2)%></b></td>
								
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudthn,2)%></td>
								<%If Cdbl(PALMbud_THN) = 0 Then%>
									<%TotalSubPerProdBud_THN = 0%>
								<%Else%>
									<%TotalSubPerProdBud_THN = Round( (cdbl(subbudthn)/cdbl(PALMbud_THN)), 2)%>
								<%End If%>
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdBud_THN, 2)%></b></td>
								
								<%If Cdbl(TBSbud_THN) = 0 Then%>
									<%TotalSubPerKerBud_THN = 0%>
								<%Else%>
									<%TotalSubPerKerBud_THN = Round( (cdbl(subbudthn)/cdbl(TBSbud_THN)), 2)%>
								<%End If%>
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerBud_THN, 2)%></b></td>	
								</tr>
                       <% end if %>
                          <tr>  
								
								<td colspan="<%= 19 + nSpanTotal %>" ><b><%=rsha("stationcode")%>. <%=rsha("STATIONCODE_DESCRIPTION")%></td>
								
						  </tr> 
					      
						  <% IF  v_vJenisReport  = "Combine"  then %>
								  <tr>  								
										<td colspan="<%= 19 + nSpanTotal %>" >&nbsp;&nbsp;&nbsp;<b><%=rsha("MACHINECODE")%>. <%=rsha("MACHINECODE_DESCRIPTION")%></td>
								  </tr> 
						  <% END IF %>
					   
					   
						  <tr>  
							<%
				        tempjob = rsha("JobCode")
						tempPlantation = "--ALL--"
						tempGroupRegion = "--SUMMARY--"
						tempRegion = "--SUMMARY--"
						tempCompany ="--SUMMARY--"
						tempfcba = "--SUMMARY--"
						If SprevCode <> rsha("JobCode") then %>
				           <td align="center">
						          <a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 
						          <%=rsha("JobCode")%>
								  </a>
						   </td>
					    <% Else %>
								 <td align="center"></td>
						<% End If %>	
						<% If Sprevjobname <> rsha("DESCRIPTION") then %>
				                 <td><%=rsha("DESCRIPTION")%></td>
						<% Else %>
								 <td align="center"></td>
						<% End If %>	   
				<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
				                     tempjob = rsha("JobCode")
									 tempPlantation = rsha("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 IF SprevPlantation <> rsha("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 		<%=rsha("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsha("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
							         tempjob = rsha("JobCode")
									 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										  tempPlantation = rsha("Plantation")
									 else
										  tempPlantation = ""
									 end if
									 tempGroupRegion = rsha("GroupRegion")
									 tempRegion = "--ALL--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"

									 IF SprevGroupRegion <> rsha("GroupRegion") then %>
										   <td align="center">
												<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 		<%=rsha("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsha("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
									   tempjob = rsha("JobCode")
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
													<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 			<%=rsha("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsha("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
									   tempjob = rsha("JobCode")
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

										 IF SprevCompany <> rsha("Company") then %>
											   <td align="center">
												    <a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 			  <%=rsha("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsha("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
									   tempjob = rsha("JobCode")
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
									   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
										   tempCompany =rsha("Company")
									   else
										   tempCompany =""
									   end if							   
									   tempfcba = rsha("fcba")

										 IF Sprevfcba <> rsha("fcba") then %>
											   <td align="center">
													<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
					<%If Cdbl(palm_BI) = 0 Then%>
						<%RowRpPerProdAct_BI = 0%>
					<%Else%>
						<%RowRpPerProdAct_BI = Round( (cdbl(rsha("BI"))/cdbl(palm_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdAct_BI, 2)%></td>
					
					<%If Cdbl(TBSprocess_BI) = 0 Then%>
						<%RowRpPerKerAct_BI = 0%>
					<%Else%>
						<%RowRpPerKerAct_BI = Round( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerAct_BI, 2)%></td>

					<td align="right"><%=formatnumber( cdbl(rsha("budBI")),2)%></td>
					
					<%If Cdbl(PALMbud_BI) = 0 Then%>
						<%RowRpPerProdBud_BI = 0%>
					<%Else%>
						<%RowRpPerProdBud_BI = Round( (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_BI, 2)%></td>
					
					<%If Cdbl(TBSbud_BI) = 0 Then%>
						<%RowRpPerKerBud_BI = 0%>
					<%Else%>
						<%RowRpPerKerBud_BI = Round( (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_BI, 2)%></td>
					
					<%If RowRpPerKerBud_BI = 0 then%>
						<%If RowRpPerKerAct_BI = 0 Then%>
							<%RowVariance_BI = 0%>
						<%Else%>
							<%RowVariance_BI = -100%>
						<%End If%>
					<%Else%>
						<%RowVariance_BI = Round( ((RowRpPerKerBud_BI - RowRpPerKerAct_BI) / RowRpPerKerBud_BI) * 100, 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowVariance_BI, 2)%></td>
					
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<%If Cdbl(palm_SBI) = 0 Then%>
						<%RowRpPerProdAct_SBI = 0%>
					<%Else%>
						<%RowRpPerProdAct_SBI = Round( (cdbl(rsha("SBI"))/cdbl(palm_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdAct_SBI, 2)%></td>
					
					<%If Cdbl(TBSprocess_SBI) = 0 Then%>
						<%RowRpPerKerAct_SBI = 0%>
					<%Else%>
						<%RowRpPerKerAct_SBI = Round( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerAct_SBI, 2)%></td>
					
					<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
					
					<%If Cdbl(PALMbud_SBI) = 0 Then%>
						<%RowRpPerProdBud_SBI = 0%>
					<%Else%>
						<%RowRpPerProdBud_SBI = Round( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_SBI, 2)%></td>
					
					<%If Cdbl(TBSbud_SBI) = 0 Then%>
						<%RowRpPerKerBud_SBI = 0%>
					<%Else%>
						<%RowRpPerKerBud_SBI = Round( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_SBI, 2)%></td>
					
					<%If RowRpPerKerBud_SBI = 0 then%>
						<%If RowRpPerKerAct_SBI = 0 Then%>
							<%RowVariance_SBI = 0%>
						<%Else%>
							<%RowVariance_SBI = -100%>
						<%End If%>
					<%Else%>
						<%RowVariance_SBI = Round( ((RowRpPerKerBud_SBI - RowRpPerKerAct_SBI) / RowRpPerKerBud_SBI) * 100, 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowVariance_SBI, 2)%></td>

					<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
					
					<%If Cdbl(PALMbud_THN) = 0 Then%>
						<%RowRpPerProdBud_THN = 0%>
					<%Else%>
						<%RowRpPerProdBud_THN = Round( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_THN, 2)%></td>
					
					<%If Cdbl(TBSbud_THN) = 0 Then%>
						<%RowRpPerKerBud_THN = 0%>
					<%Else%>
						<%RowRpPerKerBud_THN = Round( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_THN, 2)%></td>
						</tr>
						  
                          <% 
						     subAmtgrpSBI = cdbl(rsha("SBI"))	
							 subAmtgrpBI = cdbl(rsha("BI"))
							 subbudBI  =  cdbl(rsha("budBI"))
						     subbudSBI =  cdbl(rsha("budSBI"))
						     subbudthn  = cdbl(rsha("budthn")) 
				    End If  
					  if v_vJenisReport  = "Combine" THEN
                         cdmesin = rsha("MACHINECODE")	
                      end if						 
				     cdgrp = rsha("stationcode")
				     cdgrpdesc = rsha("STATIONCODE_DESCRIPTION")  
		 End If  		  
	     
		
		rsha.MoveNext
		wend
	  
		IF  v_vJenisReport  <> "Group" then %>
			<tr>  							
				<td style="background-color:#F4F4A0" colspan="<%= 2 + nSpanTotal %>"><b>Total &nbsp;<%=cdgrpdesc%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpBI,2)%></td>
				<%If Cdbl(palm_BI) = 0 Then%>
					<%TotalSubPerProdAct_BI = 0%>
				<%Else%>
					<%TotalSubPerProdAct_BI = Round( (cdbl(subAmtgrpBI)/cdbl(palm_BI)), 2)%>
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdAct_BI, 2)%></b></td>
				
				<%If Cdbl(palm_BI) = 0 Then%>
					<%TotalSubPerKerAct_BI = 0%>
				<%Else%>
					<%TotalSubPerKerAct_BI = Round( (cdbl(subAmtgrpBI)/cdbl(TBSprocess_BI)), 2)%>
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerAct_BI, 2)%></b></td>
				
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudBI,2)%></td>
				<%If Cdbl(PALMbud_BI) = 0 Then%>
					<%TotalSubPerProdBud_BI = 0%>
				<%Else%>
					<%TotalSubPerProdBud_BI = Round( (cdbl(subbudBI)/cdbl(PALMbud_BI)), 2)%>
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdBud_BI, 2)%></b></td>
				
				<%If Cdbl(TBSbud_BI) = 0 Then%>
					<%TotalSubPerKerBud_BI = 0%>
				<%Else%>
					<%TotalSubPerKerBud_BI = Round( (cdbl(subbudBI)/cdbl(TBSbud_BI)), 2)%>
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerBud_BI, 2)%></b></td>
				
				<%If TotalSubPerKerAct_BI = 0 then%>
					<%If subbudBI = 0 Then%>
						<%If subAmtgrpBI = 0 Then%>
							<%TotalSubVariance_BI = 0%>
						<%Else%>
							<%TotalSubVariance_BI = -100%>
						<%End If%>
					<%Else%>
						<%TotalSubVariance_BI = Round( ((subbudBI - subAmtgrpBI) / subbudBI) * 100, 2)%>
					<%End If%>
				<%Else%>
					<% If TotalSubPerKerBud_BI = 0 Then%>
						<%TotalSubVariance_BI = -100%>
					<%Else%>
						<%TotalSubVariance_BI = Round( ((TotalSubPerKerBud_BI - TotalSubPerKerAct_BI) / TotalSubPerKerBud_BI) * 100, 2)%>
					<%End If%>									
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubVariance_BI, 2)%></b></td>
				
				
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpSBI,2)%></td>
				<%If Cdbl(palm_SBI) = 0 Then%>
					<%TotalSubPerProdAct_SBI = 0%>
				<%Else%>
					<%TotalSubPerProdAct_SBI = Round( (cdbl(subAmtgrpSBI)/cdbl(palm_SBI)), 2)%>
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdAct_SBI, 2)%></b></td>
				
				<%If Cdbl(TBSprocess_SBI) = 0 Then%>
					<%TotalSubPerKerAct_SBI = 0%>
				<%Else%>
					<%TotalSubPerKerAct_SBI = Round( (cdbl(subAmtgrpSBI)/cdbl(TBSprocess_SBI)), 2)%>
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerAct_SBI, 2)%></b></td>
				
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudSBI,2)%></td>
				<%If Cdbl(PALMbud_SBI) = 0 Then%>
					<%TotalSubPerProdBud_SBI = 0%>
				<%Else%>
					<%TotalSubPerProdBud_SBI = Round( (cdbl(subbudSBI)/cdbl(PALMbud_SBI)), 2)%>
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdBud_SBI, 2)%></b></td>
				
				<%If Cdbl(TBSbud_SBI) = 0 Then%>
					<%TotalSubPerKerBud_SBI = 0%>
				<%Else%>
					<%TotalSubPerKerBud_SBI = Round( (cdbl(subbudSBI)/cdbl(TBSbud_SBI)), 2)%>
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerBud_SBI, 2)%></b></td>
				
				<%If TotalSubPerKerAct_SBI = 0 then%>
					<%If subbudSBI = 0 Then%>
						<%If subAmtgrpSBI = 0 Then%>
							<%TotalSubVariance_BI = 0%>
						<%Else%>
							<%TotalSubVariance_BI = -100%>
						<%End If%>
					<%Else%>
						<%TotalSubVariance_BI = Round( ((subbudSBI - subAmtgrpSBI) / subbudSBI) * 100, 2)%>
					<%End If%>
				<%Else%>
					<% If TotalSubPerKerBud_SBI = 0 Then%>
						<%TotalSubVariance_BI = -100%>
					<%Else%>
						<%TotalSubVariance_BI = Round( ((TotalSubPerKerBud_SBI - TotalSubPerKerAct_SBI) / TotalSubPerKerBud_SBI) * 100, 2)%>
					<%End If%>									
				<%End If%>
				<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubVariance_BI, 2)%></b></td>
				
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudthn,2)%></td>
				<%If Cdbl(PALMbud_THN) = 0 Then%>
					<%TotalSubPerProdBud_THN = 0%>
				<%Else%>
					<%TotalSubPerProdBud_THN = Round( (cdbl(subbudthn)/cdbl(PALMbud_THN)), 2)%>
				<%End If%>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerProdBud_THN, 2)%></b></td>
				
				<%If Cdbl(TBSbud_THN) = 0 Then%>
					<%TotalSubPerKerBud_THN = 0%>
				<%Else%>
					<%TotalSubPerKerBud_THN = Round( (cdbl(subbudthn)/cdbl(TBSbud_THN)), 2)%>
				<%End If%>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalSubPerKerBud_THN, 2)%></b></td>	
			</tr>
	   <%END IF 
	end if  
	
	
	if   v_vJenisReport  = "Job"  then %>
	         <tr>  
				
				<td colspan="19"><b>Rekapitulasi</td>
				
			</tr>
     <%			
			strdtl = " SELECT " & strFieldCons & "  JobCode,DESCRIPTION, sum(BI) BI, sum(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn from (  " & _
					 "	select  t.ba fcba,t.JobCode,t.ACCOUNTCODE_DESCRIPTION DESCRIPTION, sum(t.amount) BI, 0 SBI ," & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.ba," & v_txtyear  & ",1,1,'',t.JobCode,'','','','','',to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.ba," & v_txtyear  & ",1,1,'',t.JobCode,'','','','','',to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.ba," & v_txtyear  & ",1,1,'',t.JobCode,'','','','','',to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY'))  budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by  t.ba,t.period_year,t.period_month,t.JobCode,t.ACCOUNTCODE_DESCRIPTION " & _
					 "	UNION ALL" & _
					 "	select t.ba fcba,t.JobCode,t.ACCOUNTCODE_DESCRIPTION DESCRIPTION, 0 BI ,sum(t.amount) SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by  t.ba,t.period_year,t.period_month,t.JobCode,t.ACCOUNTCODE_DESCRIPTION " & _
					 "	 ) " & _
					 "	group by fcba,JobCode,DESCRIPTION " 
					 
			strdtl = "select GroupPlantation " & selectCons & ", JobCode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by JobCode,DESCRIPTION,GroupPlantation" & selectCons & "  Order by JobCode,DESCRIPTION,GroupPlantation" & orderCons  	
	     
			
			Set rsha = Server.CreateObject("ADODB.Recordset")
			rsha.Open strdtl, connstr, adOpenStatic 
			if Not rsha.eof then
				  while Not rsha.eof 
			 %>
					 <tr>  
						<%
				        tempjob = rsha("JobCode")
						tempPlantation = "--ALL--"
						tempGroupRegion = "--SUMMARY--"
						tempRegion = "--SUMMARY--"
						tempCompany ="--SUMMARY--"
						tempfcba = "--SUMMARY--"
						If SprevCode <> rsha("JobCode") then %>
				           <td align="center">
						          <a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 
						          <%=rsha("JobCode")%>
								  </a>
						   </td>
					    <% Else %>
								 <td align="center"></td>
						<% End If %>	
						<% If Sprevjobname <> rsha("DESCRIPTION") then %>
				                 <td><%=rsha("DESCRIPTION")%></td>
						<% Else %>
								 <td align="center"></td>
						<% End If %>	   
				<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
				                     tempjob = rsha("JobCode")
									 tempPlantation = rsha("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 IF SprevPlantation <> rsha("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 		<%=rsha("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsha("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
							         tempjob = rsha("JobCode")
									 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										  tempPlantation = rsha("Plantation")
									 else
										  tempPlantation = ""
									 end if
									 tempGroupRegion = rsha("GroupRegion")
									 tempRegion = "--ALL--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"

									 IF SprevGroupRegion <> rsha("GroupRegion") then %>
										   <td align="center">
												<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 		<%=rsha("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsha("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
									   tempjob = rsha("JobCode")
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
													<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 			<%=rsha("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsha("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
									   tempjob = rsha("JobCode")
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

										 IF SprevCompany <> rsha("Company") then %>
											   <td align="center">
												    <a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>"> 			  <%=rsha("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsha("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
									   tempjob = rsha("JobCode")
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
									   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
										   tempCompany =rsha("Company")
									   else
										   tempCompany =""
									   end if							   
									   tempfcba = rsha("fcba")

										 IF Sprevfcba <> rsha("fcba") then %>
											   <td align="center">
													<a href="p_accountreportplant.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("BI")),2)%></td>
					<%If Cdbl(palm_BI) = 0 Then%>
						<%RowRpPerProdAct_BI = 0%>
					<%Else%>
						<%RowRpPerProdAct_BI = Round( (cdbl(rsha("BI"))/cdbl(palm_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdAct_BI, 2)%></td>
					
					<%If Cdbl(TBSprocess_BI) = 0 Then%>
						<%RowRpPerKerAct_BI = 0%>
					<%Else%>
						<%RowRpPerKerAct_BI = Round( (cdbl(rsha("BI"))/cdbl(TBSprocess_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerAct_BI, 2)%></td>

					<td align="right"><%=formatnumber( cdbl(rsha("budBI")),2)%></td>
					
					<%If Cdbl(PALMbud_BI) = 0 Then%>
						<%RowRpPerProdBud_BI = 0%>
					<%Else%>
						<%RowRpPerProdBud_BI = Round( (cdbl(rsha("budBI"))/cdbl(PALMbud_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_BI, 2)%></td>
					
					<%If Cdbl(TBSbud_BI) = 0 Then%>
						<%RowRpPerKerBud_BI = 0%>
					<%Else%>
						<%RowRpPerKerBud_BI = Round( (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_BI, 2)%></td>
					
					<%If RowRpPerKerBud_BI = 0 then%>
						<%If RowRpPerKerAct_BI = 0 Then%>
							<%RowVariance_BI = 0%>
						<%Else%>
							<%RowVariance_BI = -100%>
						<%End If%>
					<%Else%>
						<%RowVariance_BI = Round( ((RowRpPerKerBud_BI - RowRpPerKerAct_BI) / RowRpPerKerBud_BI) * 100, 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowVariance_BI, 2)%></td>
					
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<%If Cdbl(palm_SBI) = 0 Then%>
						<%RowRpPerProdAct_SBI = 0%>
					<%Else%>
						<%RowRpPerProdAct_SBI = Round( (cdbl(rsha("SBI"))/cdbl(palm_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdAct_SBI, 2)%></td>
					
					<%If Cdbl(TBSprocess_SBI) = 0 Then%>
						<%RowRpPerKerAct_SBI = 0%>
					<%Else%>
						<%RowRpPerKerAct_SBI = Round( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerAct_SBI, 2)%></td>
					
					<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
					
					<%If Cdbl(PALMbud_SBI) = 0 Then%>
						<%RowRpPerProdBud_SBI = 0%>
					<%Else%>
						<%RowRpPerProdBud_SBI = Round( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_SBI, 2)%></td>
					
					<%If Cdbl(TBSbud_SBI) = 0 Then%>
						<%RowRpPerKerBud_SBI = 0%>
					<%Else%>
						<%RowRpPerKerBud_SBI = Round( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_SBI, 2)%></td>
					
					<%If RowRpPerKerBud_SBI = 0 then%>
						<%If RowRpPerKerAct_SBI = 0 Then%>
							<%RowVariance_SBI = 0%>
						<%Else%>
							<%RowVariance_SBI = -100%>
						<%End If%>
					<%Else%>
						<%RowVariance_SBI = Round( ((RowRpPerKerBud_SBI - RowRpPerKerAct_SBI) / RowRpPerKerBud_SBI) * 100, 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowVariance_SBI, 2)%></td>

					<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
					
					<%If Cdbl(PALMbud_THN) = 0 Then%>
						<%RowRpPerProdBud_THN = 0%>
					<%Else%>
						<%RowRpPerProdBud_THN = Round( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerProdBud_THN, 2)%></td>
					
					<%If Cdbl(TBSbud_THN) = 0 Then%>
						<%RowRpPerKerBud_THN = 0%>
					<%Else%>
						<%RowRpPerKerBud_THN = Round( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)), 2)%>
					<%End If%>
					<td align="right"><%=formatnumber(RowRpPerKerBud_THN, 2)%></td>
								
					</tr>	
			 <%
					 rsha.MoveNext
				  wend
			end if
	end if		
	%>
	
	
	<tr>  							
		<td style="background-color:#F4F4A0" colspan="<%= 2 + nSpanTotal %>"><b> Grand Total</td>
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
								<%If Cdbl(palm_BI) = 0 Then%>
									<%TotalRpPerProdAct_BI = 0%>
								<%Else%>
									<%TotalRpPerProdAct_BI = Round( (cdbl(GRAmtgrpBI)/cdbl(palm_BI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerProdAct_BI, 2)%></b></td>
								
								<%If Cdbl(TBSprocess_BI) = 0 Then%>
									<%TotalRpPerKerAct_BI = 0%>
								<%Else%>
									<%TotalRpPerKerAct_BI = Round( (cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerKerAct_BI, 2)%></b></td>

								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudBI,2)%></td>
								
								<%If Cdbl(PALMbud_BI) = 0 Then%>
									<%TotalRpPerProdBud_BI = 0%>
								<%Else%>
									<%TotalRpPerProdBud_BI = Round( (cdbl(GRbudBI)/cdbl(PALMbud_BI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerProdBud_BI, 2)%></b></td>
								
								<%If Cdbl(TBSbud_BI) = 0 Then%>
									<%TotalRpPerKerBud_BI = 0%>
								<%Else%>
									<%TotalRpPerKerBud_BI = Round( (cdbl(GRbudBI)/cdbl(TBSbud_BI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerKerBud_BI, 2)%></b></td>
								
								<%If TotalRpPerKerAct_BI = 0 then%>
									<%If GRbudBI = 0 Then%>
										<%If GRAmtgrpBI = 0 Then%>
											<%TotalVariance_BI = 0%>
										<%Else%>
											<%TotalVariance_BI = -100%>
										<%End If%>
									<%Else%>
										<%TotalVariance_BI = Round( ((GRbudBI - GRAmtgrpBI) / GRbudBI) * 100, 2)%>
									<%End If%>
								<%Else%>
									<% If TotalRpPerKerBud_BI = 0 Then%>
										<%TotalVariance_BI = -100%>
									<%Else%>
										<%TotalVariance_BI = Round( ((TotalRpPerKerBud_BI - TotalRpPerKerAct_BI) / TotalRpPerKerBud_BI) * 100, 2)%>
									<%End If%>									
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalVariance_BI, 2)%></b></td>
								
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
								
								<%If Cdbl(palm_SBI) = 0 Then%>
									<%TotalRpPerProdAct_SBI = 0%>
								<%Else%>
									<%TotalRpPerProdAct_SBI = Round( (cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerProdAct_SBI, 2)%></b></td>
								
								<%If Cdbl(TBSprocess_SBI) = 0 Then%>
									<%TotalRpPerKerAct_SBI = 0%>
								<%Else%>
									<%TotalRpPerKerAct_SBI = Round( (cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerKerAct_SBI, 2)%></b></td>

								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudSBI,2)%></td>
								
								<%If Cdbl(PALMbud_SBI) = 0 Then%>
									<%TotalRpPerProdBud_SBI = 0%>
								<%Else%>
									<%TotalRpPerProdBud_SBI = Round( (cdbl(GRbudSBI)/cdbl(PALMbud_SBI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerProdBud_SBI, 2)%></b></td>
								
								<%If Cdbl(TBSbud_SBI) = 0 Then%>
									<%TotalRpPerKerBud_SBI = 0%>
								<%Else%>
									<%TotalRpPerKerBud_SBI = Round( (cdbl(GRbudSBI)/cdbl(TBSbud_SBI)), 2)%>
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerKerBud_SBI, 2)%></b></td>
								
								
								<%If TotalRpPerKerAct_SBI = 0 then%>
									<%If GRbudSBI = 0 Then%>
										<%If GRAmtgrpSBI = 0 Then%>
											<%TotalVariance_BI = 0%>
										<%Else%>
											<%TotalVariance_BI = -100%>
										<%End If%>
									<%Else%>
										<%TotalVariance_BI = Round( ((GRbudSBI - GRAmtgrpSBI) / GRbudSBI) * 100, 2)%>
									<%End If%>
								<%Else%>
									<% If TotalRpPerKerBud_SBI = 0 Then%>
										<%TotalVariance_BI = -100%>
									<%Else%>
										<%TotalVariance_BI = Round( ((TotalRpPerKerBud_SBI - TotalRpPerKerAct_SBI) / TotalRpPerKerBud_SBI) * 100, 2)%>
									<%End If%>									
								<%End If%>
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalVariance_BI, 2)%></b></td>

								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudthn,2)%></td>
								
								<%If Cdbl(PALMbud_THN) = 0 Then%>
									<%TotalRpPerProdBud_THN = 0%>
								<%Else%>
									<%TotalRpPerProdBud_THN = Round( (cdbl(GRbudthn)/cdbl(PALMbud_THN)), 2)%>
								<%End If%>
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerProdBud_THN, 2)%></b></td>
								
								<%If Cdbl(TBSbud_THN) = 0 Then%>
									<%TotalRpPerKerBud_THN = 0%>
								<%Else%>
									<%TotalRpPerKerBud_THN = Round( (cdbl(GRbudthn)/cdbl(TBSbud_THN)), 2)%>
								<%End If%>
								<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(TotalRpPerKerBud_THN, 2)%></b></td>
	</tr>
 <% END IF %>	
    </tbody>
</table>
</form>
<!--#include file="bottom.Asp"-->