<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim v_txtDivision, v_txtPlantingDate, v_txtBlok
Dim v_txtyear,v_txtmonth,no
Dim txtDivision, txtBlok, txtPlantingDate
Dim v_vstatus,v_vCROPTYPE,v_vJenisReport, v_vOwnership
Dim tempyear,tempmonth
Dim tempDivision, tempBlok, tempPlantingDate 
Dim SprevCode,Sprevjobname
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

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

Dim v_txttype : v_txttype = Request("txttype")
if len(trim(v_txttype)) = 0 then v_txttype = "rekap"
            
			 v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	  = Request("txtPlantation")
			v_txtGroupRegion  	  = Request("txtGroupRegion")
			v_txtRegion  		  = Request("txtRegion")
			v_txtcompany  		  = Request("txtCompany")
			v_txtfcba 			  = Request("txtfcba")
			v_txtDivision		  = Request("txtDivision")
			v_txtPlantingDate	  = Request("txtPlantingDate")
			v_txtBlok			  = Request("txtBlok")
			v_txtyear             = Request("txtYear")
			v_txtmonth            = Request("txtMonth")
			' v_vstatus             = Request("vStatus")
			' v_vCROPTYPE           = Request("vCROPTYPE")
			' v_vJenisReport        = request("vJenisReport")
			' v_vOwnership	      = Request("vOwnership")
			nSpanTotal = 5
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " TT" & v_txtPlantingDate
			'response.write  "YEAR " & v_txtyear & "  month " & v_txtmonth& "  status " & v_vstatus & " crop " & v_vCROPTYPE  & " owner " & v_vOwnership & " jenis " & v_vJenisReport  

					
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
<title>ePlaS Reporting - Analisa Biaya Tak Langsung</title>
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
	<form name="frmBTL" method="post" action="<%=ScriptName%>">
	<!--#include file="FormSelectNewSumP.html"-->
	<!--#include file="FromSelectHDRBTLMILL.html"-->
	</form>
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
		   if   request.Form("vJenisReport") = "Group"  then
		        rekap = "REKAP BIAYA PER STASIUN"
		   else		
				rekap = "REKAP BIAYA PER STASIUN PER PEKERJAAN"
		   end if
				
		%>   
		<h2>LAPORAN BIAYA TAK LANGSUNG</h3>
		<% IF v_txtmonth <> "" then %>
		<h5>PERIODE     : <%=UCASE(MonthName(cint(ia)))%> &nbsp; <%=v_txtyear%></h5>
		<% END IF %>
		
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

strtbs = " select GroupPlantation ,sum(TBS_PROCESS_BI) TBS_PROCESS_BI, sum(TBS_PROCESS_SBI) TBS_PROCESS_SBI From ( " & _
         "  select " & strFieldCons & " sum(TBS_PROCESS_BI) TBS_PROCESS_BI, sum(TBS_PROCESS_SBI) TBS_PROCESS_SBI  from ( " & _
		 "  select t.fcba, sum(t.cageweight) TBS_PROCESS_BI,0 TBS_PROCESS_SBI from SHIFTREPORT t " & _  
	     "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.productiondate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  " & _
		 "  union all" & _
		 "  select t.fcba, 0 TBS_PROCESS_BI, sum(t.cageweight) TBS_PROCESS_SBI from SHIFTREPORT t  " & _
		 "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.productiondate,'MM') <= lpad('" & v_txtmonth & "',2,0) " & _
         "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  ) group by fcba )  " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'P' group by GroupPlantation"
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic
'response.write strtbs
if Not rsha.eof then
    TBSprocess_BI = rsha("TBS_PROCESS_BI")
	TBSprocess_SBI = rsha("TBS_PROCESS_SBI")
end if
rsha.close()

ProdCPO_BI =0
ProdCPO_SBI =0
		 
strtbs = " select GroupPlantation ,sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  From ( " & _
         " select  " & strFieldCons & " sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_CPO_BI, 0 Produksi_CPO_SBI from MILL_PRODUCTION t  " & _
		 " where t.product_code='CPKO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union all" & _
		 " select t.fcba,0 Produksi_CPO_BI, sum(t.production_today) Produksi_CPO_SBI from MILL_PRODUCTION t " & _
		 " where t.product_code='CPKO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba )  " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'P' group by GroupPlantation"	 
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic
'response.write strtbs

if Not rsha.eof then
    ProdCPO_BI = rsha("Produksi_CPO_BI")
	ProdCPO_SBI = rsha("Produksi_CPO_SBI")
end if
rsha.close()

ProdKER_BI =0
ProdKER_SBI =0
		 
strtbs = " select GroupPlantation ,sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI From ( " & _
         " select " & strFieldCons & " sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_KER_BI, 0 Produksi_KER_SBI from MILL_PRODUCTION t  " & _
		 " where t.product_code='PKE'  and to_char(t.tdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union all" & _
		 " select t.fcba,0 Produksi_KER_BI, sum(t.production_today) Produksi_KER_SBI from MILL_PRODUCTION t " & _
		 " where t.product_code='PKE'  and to_char(t.tdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba  )  " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'P' group by GroupPlantation"	  		 
		 
	  
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
strtbs = " select GroupPlantation ,sum(TBSbudBI) TBSbudBI,sum(tbsbudSBI) tbsbudSBI,sum(tbsbudthn) tbsbudthn,sum(cpobudBI) cpobudBI,sum(cpobudSBI) cpobudSBI,sum(cpobudthn) cpobudthn,sum(kerbudBI) kerbudBI,sum(kerbudSBI) kerbudSBI, sum(kerbudthn) kerbudthn  from ( " & _
         "  select " & strFieldCons & " sum(TBSbudBI) TBSbudBI,sum(tbsbudSBI) tbsbudSBI,sum(tbsbudthn) tbsbudthn,sum(cpobudBI) cpobudBI,sum(cpobudSBI) cpobudSBI,sum(cpobudthn) cpobudthn,sum(kerbudBI) kerbudBI,sum(kerbudSBI) kerbudSBI, sum(kerbudthn) kerbudthn  from ( " & _
		 "	 select btm.fcba , sum(" & strbudHDR_BI & ") TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='FFB'   " & _
		 "	 group by btm.fcba " & _
		 "	 union all" & _
		 "	 select btm.fcba , 0 TBSbudBI,sum(" & strbudHDR_SBI & ") tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='FFB'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='FFB'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,sum(" & strbudHDR_BI & ") cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='CPKO'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,sum(" & strbudHDR_SBI & ") cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='CPKO'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) cpobudthn,0 kerbudBI, 0 kerbudSBI , 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='CPKO'  " & _
		 "	 group by btm.fcba " & _
		 "	  union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudthn,0 cpobudSBI,sum(" & strbudHDR_BI & ") kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='PKE'  " & _
		 "	 group by btm.fcba " & _
		 "	  union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, sum(" & strbudHDR_SBI & ") kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='PKE'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='PKE'  " & _
		 "	 group by btm.fcba ) group by fcba  )  " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'P' group by GroupPlantation" 
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic
'response.write strtbs
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
	<tr>
		<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>"th>
		
		<th  style="background-color:#00A300;color:#fff;" colspan="7">BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="7">S/D BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="3" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>"th>
		<th  style="background-color:#00A300;color:#fff;" colspan="3">REALISASI</th>
		
		<th  style="background-color:#00A300;color:#fff;" colspan="3">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;" >%</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="3">REALISASI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="3">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;" >%</th>
	</tr>
	<tr>
	    <%
		'TBSbud_BI,TBSbud_SBI,TBSbud_THN,CPObud_BI,CPObud_SBI,CPObud_THN,KERbud_BI,KERbud_SBI,KERbud_THN,PALMbud_BI,PALMbud_SBI,PALMbud_THN
		%>
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">KERNEL PROSES (KG)</th>
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
	    
		<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI CPKO (KG)</th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(ProdCPO_BI,2)%></th>
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
	    
		<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI PKE (KG)</th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(ProdKER_BI,2)%></th>
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
	    
		<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">PRODUK (KG)</th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(Palm_BI,2)%></th>
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
	<tr   >
	  
	    <th style="background-color:#00A300;color:#fff;" rowspan="2" >Kode</th>
		<th style="background-color:#00A300;color:#fff;"rowspan="2" >Nama Pekerjaan</th>
		
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

		
		
		
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Rp./Kg.</th>
		
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Rp./Kg.</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2" align="right" ></th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Rp./Kg.</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Rp./Kg.</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2" align="right" ></th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Total</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Rp./Kg.</th>
	</tr>
	<tr>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >Kernel</th>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >Kernel</th>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >Kernel</th>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >Kernel</th>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >Kernel</th>
	</tr>
</thead>
<tbody>	
	<%
	Dim strdtl , strgrp ,cdgrp,cdgrpdesc  ,noline ,  subAmtgrpBI , GRAmtgrpBI, subAmtgrpSBI , GRAmtgrpSBI , grpDesc,GRbudBI,GRbudSBI,GRbudthn
	Dim rsjob,strjob,debetjob,kreditjob,notAlokasi

	strjob= "select report_posting_api.getPostingDataDirect('PSD','ANALISABTL','p_AccountReportBTL.asp','DATADEBET_MILL','JOB') debet, report_posting_api.getPostingDataDirect('PSD','ANALISABTL','p_AccountReportBTL.asp','DATAKREDIT_MILL','JOB')   kredit  from dual"
	Set rsjob = Server.CreateObject("ADODB.Recordset")
	'response.write strjob
	rsjob.Open strjob, connstr, adOpenStatic
	if Not rsjob.eof then
	   debetjob = rsjob("debet")
	   kreditjob = rsjob("kredit")
	end if
	
	
			strdtl = "  select GroupPlantation " & selectCons & "  ,jobcode,description, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from ( " & _
                     "  select " & strFieldCons & " jobcode,description, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from ( " & _
					 "	select t.fcba,t.jobcode,job_api.get_jobcode_name(t.jobcode,t.fcba) Description, sum(t.amount) BI, 0 SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from cost_activity t where instr('" & debetjob & "',t.jobcode) <> 0 and to_char(t.tdate,'YYYY')=  '" & v_txtyear & "' and to_char(t.tdate,'MM')=  lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.fcba, t.jobcode " & _
					 "	union " & _
					 "	select t.fcba, t.jobcode,job_api.get_jobcode_name(t.jobcode,t.fcba) Description,0 BI, sum(t.amount) SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from cost_activity t where instr('" & debetjob & "',t.jobcode) <> 0 and to_char(t.tdate,'YYYY')=  '" & v_txtyear & "' and to_char(t.tdate,'MM') <=   lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.fcba, t.jobcode  " & _
					 "	union " & _
					 "	select t.fcba, t.jobcode,job_api.get_jobcode_name(t.jobcode,t.fcba) Description, sum(t.amount)* -1 BI, 0 SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from cost_activity t where instr('" & kreditjob & "',t.jobcode)  <> 0   and to_char(t.tdate,'YYYY')=  '" & v_txtyear & "' and to_char(t.tdate,'MM')=  lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.fcba, t.jobcode " & _
					 "	union " & _
					 "	select  t.fcba,t.jobcode,job_api.get_jobcode_name(t.jobcode,t.fcba) Description,0 BI, sum(t.amount)* -1  SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from cost_activity t where instr('" & kreditjob & "',t.jobcode)  <> 0   and to_char(t.tdate,'YYYY')=  '" & v_txtyear & "' and to_char(t.tdate,'MM') <=   lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.fcba, t.jobcode " & _
					 "	union " & _
					 "	select btm.fcba,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,sum(" & strbud_BI & ")* -1 budBI,0 budSBI,0 budthn  from BUDGET_TRANSACTION_MILL btm " & _
					 "	where btm.budget_year ='" & v_txtyear & "'  and instr('" & kreditjob & "',btm.jobcode) <> 0  " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.fcba,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,0 budBI,sum(" & strbud_SBI & ")* -1 budSBI,0 budthn  from BUDGET_TRANSACTION_MILL btm " & _
					 "	where btm.budget_year ='" & v_txtyear & "'  and instr('" & kreditjob & "',btm.jobcode) <> 0  " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.fcba,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,0 budBI,0 budSBI,sum(btm.budget_01+btm.budget_02+btm.budget_03+btm.budget_04+btm.budget_05+btm.budget_06+btm.budget_07+btm.budget_08+btm.budget_09+btm.budget_10+btm.budget_11+btm.budget_12)* -1 budthn  from BUDGET_TRANSACTION_MILL btm " & _
					 "	where btm.budget_year ='" & v_txtyear & "'  and instr('" & kreditjob & "',btm.jobcode) <> 0  " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.fcba,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,sum(" & strbud_BI & ") budBI,0 budSBI,0 budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _
					 "	where btm.budget_year ='" & v_txtyear & "' and  instr('" & debetjob & "',btm.jobcode) <> 0 " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.fcba,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,0 budBI,sum(" & strbud_SBI & ") budSBI,0 budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _
					 "	where btm.budget_year ='" & v_txtyear & "'  and instr('" & debetjob & "',btm.jobcode) <> 0 " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	union " & _
					 "	select btm.fcba,btm.jobcode ,job_api.get_jobcode_name(btm.jobcode,btm.fcba) DESCRIPTION,0 BI, 0 sbi ,0 budBI,0 budSBI,sum(btm.budget_01+btm.budget_02+btm.budget_03+btm.budget_04+btm.budget_05+btm.budget_06+btm.budget_07+btm.budget_08+btm.budget_09+btm.budget_10+btm.budget_11+btm.budget_12) budthn  from BUDGET_TRANSACTION_INDIRCOST btm " & _
					 "	where btm.budget_year ='" & v_txtyear & "'  and  instr('" & debetjob & "',btm.jobcode) <> 0 " & _
					 "	group by btm.fcba,btm.jobcode " & _
					 "	 ) " & _
					 "	group by fcba, jobcode,description )   " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'P' group by GroupPlantation, jobcode,description " & selectCons & " order by jobcode " & selectCons 
	
		
		
			 strdtl =  strdtl  + strsql1  + strgrp 
	'response.write strdtl
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic

	GRAmtgrpBI ="0"
	GRAmtgrpSBI ="0"
	if Not rsha.eof then
		  while Not rsha.eof
		    
			 IF rsha("JobCode") <> "6610600"	 THEN  %>
				<tr>  
					       <% If SprevCode <> rsha("JOBCODE") then %>
								 <td align="center"><%=rsha("JOBCODE")%></td>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>				   
							<% If Sprevjobname <> rsha("description") then %>
								 <td><%=rsha("description")%></td>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>	
						   <%
						   SprevCode   = rsha("JOBCODE")
						   Sprevjobname = rsha("description")
						   notAlokasi = "T"
						   %>
				           <% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
									 tempPlantation = rsha("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevPlantation <> rsha("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
												<%=rsha("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsha("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
									 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										  tempPlantation = rsha("Plantation")
									 else
										  tempPlantation = ""
									 end if
									 tempGroupRegion = rsha("GroupRegion")
									 tempRegion = "--ALL--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevGroupRegion <> rsha("GroupRegion") then %>
										   <td align="center">
												<a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
												<%=rsha("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsha("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
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
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevRegion <> rsha("Region") then %>
											   <td align="center">
													<a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
													<%=rsha("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsha("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
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
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevCompany <> rsha("Company") then %>
											   <td align="center">
												    <a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
													  <%=rsha("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsha("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
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
									   tempDivision = "--ALL--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF Sprevfcba <> rsha("fcba") then %>
											   <td align="center">
													<a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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

				<tr>	
            <% ELSE  
			  
				   if notAlokasi = "T"   then   %>	
						  <tr>
								<td  style="background-color:#F4F4A0"  colspan="<%= 2 + nSpanTotal %>"><b>Total Biaya Tak Langsung</td>
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
							<tr>   		  
								 
		  <%            notAlokasi = "F"
		         end if 
                
					 ' GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
					 ' GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
					 ' GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))
					 ' GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
					 ' GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))
 				 %>
				<tr>  
					<% If SprevCode <> rsha("JOBCODE") then %>
								 <td align="center"><%=rsha("JOBCODE")%></td>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>				   
							<% If Sprevjobname <> rsha("description") then %>
								 <td><%=rsha("description")%></td>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>	
						   <%
						   SprevCode   = rsha("JOBCODE")
						   Sprevjobname = rsha("description")
						   %>
				           <% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
									 tempPlantation = rsha("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevPlantation <> rsha("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
												<%=rsha("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsha("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
									 if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
										  tempPlantation = rsha("Plantation")
									 else
										  tempPlantation = ""
									 end if
									 tempGroupRegion = rsha("GroupRegion")
									 tempRegion = "--ALL--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevGroupRegion <> rsha("GroupRegion") then %>
										   <td align="center">
												<a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
												<%=rsha("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsha("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
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
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevRegion <> rsha("Region") then %>
											   <td align="center">
													<a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
													<%=rsha("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsha("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
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
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevCompany <> rsha("Company") then %>
											   <td align="center">
												    <a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
													  <%=rsha("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsha("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
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
									   tempDivision = "--ALL--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF Sprevfcba <> rsha("fcba") then %>
											   <td align="center">
													<a href="p_accountreportbtlplant.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
              					
               
          <% END IF %>				
    <%        
	                 GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
					 GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
					 GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))
					 GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
					 GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))
   			 
		      rsha.MoveNext
		  wend
	end if    
	%>		
               <tr>  							
					<td  style="background-color:#F4F4A0" colspan="<%= 2 + nSpanTotal %>"><b>Total Keseluruhan</td>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>

					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>
					
					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>

					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudSBI,2)%></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>
					<td  style="background-color:#F4F4A0" align="right"><b></b></td>

					<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudthn,2)%></td>
					<td style="background-color:#F4F4A0" align="right"><b></b></td>
					<td style="background-color:#F4F4A0" align="right"><b></b></td>														
				<tr>  
         </tbody>	
</table>
<!--#include file="bottom.Asp"-->