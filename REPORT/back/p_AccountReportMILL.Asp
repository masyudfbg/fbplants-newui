<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa, strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
Dim v_txtyear,v_txtmonth,v_vJenisReport,v_txtJob
Dim SprevCode,Sprevjobname
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempjob
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,rsJob
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
<title>iPla'S Reporting - Analisa Biaya Pabrik</title>
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
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<form name="frmMill" method="post" action="<%=ScriptName%>">
	<!--#include file="FormSelectNewSum.html"-->
	<!--#include file="FromSelectHDRMILL.html"-->
	
<div id="wrapper">
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
<tr>
<th colspan="24" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">

	<div id="header">
		   <% If  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
	</div>
 </th>
 </tr>
 <tr>
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
		 "  select t.fcba, sum(t.cageweight) TBS_PROCESS_BI,0 TBS_PROCESS_SBI from SHIFTREPORT t " & _  
	     "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.productiondate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  " & _
		 "  union " & _
		 "  select t.fcba, 0 TBS_PROCESS_BI, sum(t.cageweight) TBS_PROCESS_SBI from SHIFTREPORT t  " & _
		 "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.productiondate,'MM') <= lpad('" & v_txtmonth & "',2,0) " & _
         "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  ) group by fcba " 

 strtbs = "select GroupPlantation , sum(TBS_PROCESS_BI) TBS_PROCESS_BI, sum(TBS_PROCESS_SBI) TBS_PROCESS_SBI  from ( " & strtbs & " )" & whereCons & " group by GroupPlantation"
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

' strtbs = " select fcba, sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 ' " select t.fcba,  sum(t.quantity) Produksi_CPO_BI,0 Produksi_CPO_SBI from PRODUCTSTORAGEREADING t  " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='CPO')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba, to_char(t.readingdate,'YYYYMM') " & _
		 ' " union " & _
		 ' " select t.fcba, 0 Produksi_CPO_BI,sum(t.quantity) Produksi_CPO_SBI from PRODUCTSTORAGEREADING t " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='CPO')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba, to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select " & strFieldCons & " sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_CPO_BI, 0 Produksi_CPO_SBI from MILL_PRODUCTION t  " & _
		 " where t.product_code='CPO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union " & _
		 " select t.fcba,0 Produksi_CPO_BI, sum(t.production_today) Produksi_CPO_SBI from MILL_PRODUCTION t " & _
		 " where t.product_code='CPO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba " 
		 
 strtbs = "select GroupPlantation , sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI   from ( " & strtbs & " )" & whereCons & " group by GroupPlantation"
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

' strtbs = " select fcba, sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI from ( " & _
	     ' " select t.fcba, sum(t.quantity) Produksi_KER_BI, 0 Produksi_KER_SBI from PRODUCTSTORAGEREADING t  " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='Kernel')   and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') = lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba,to_char(t.readingdate,'YYYYMM')  " & _
		 ' " union " & _
		 ' " select t.fcba, 0 Produksi_KER_BI, sum(t.quantity) Produksi_KER_SBI from PRODUCTSTORAGEREADING t " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='Kernel')  and to_char(t.readingdate,'YYYY') = '" & request.Form("vYear") & "'  and  to_char(t.readingdate,'MM') <= lpad('" & request.Form("vMonth") & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba,to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select " & strFieldCons & "  sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_KER_BI, 0 Produksi_KER_SBI from MILL_PRODUCTION t  " & _
		 " where t.product_code='KER'  and to_char(t.tdate,'YYYY') = '" & v_txtyear  & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union " & _
		 " select t.fcba,0 Produksi_KER_BI, sum(t.production_today) Produksi_KER_SBI from MILL_PRODUCTION t " & _
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
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='FFB' " & _
		 "	 group by btm.fcba " & _
		 "	 union " & _
		 "	 select btm.fcba , 0 TBSbudBI,sum(" & strbudHDR_SBI & ") tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='FFB' " & _
		 "	 group by btm.fcba " & _
		 "	 union " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear  & "' and btm.productcode='FFB' " & _
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
	    <% 'TBSbud_BI,TBSbud_SBI,TBSbud_THN,CPObud_BI,CPObud_SBI,CPObud_THN,KERbud_BI,KERbud_SBI,KERbud_THN,PALMbud_BI,PALMbud_SBI,PALMbud_THN %>
		<th  style="background-color:#00A300;color:#fff;"   colspan="<%= 2 + nSpanTotal %>" align="right">TBS PROSES (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSprocess_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSbud_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(TBSprocess_BI)-cdbl(TBSbud_BI))/(cdbl(TBSbud_BI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSprocess_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSbud_SBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(TBSprocess_SBI)-cdbl(TBSbud_SBI))/(cdbl(TBSbud_SBI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(TBSbud_THN,2)%></th>
		
		
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI CPO (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdCPO_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(CPObud_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(ProdCPO_BI)-cdbl(CPObud_BI))/(cdbl(CPObud_BI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdCPO_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(CPObud_SBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(ProdCPO_SBI)-cdbl(CPObud_SBI))/(cdbl(CPObud_SBI))*100,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(CPObud_THN,2)%></th>
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI KER (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdKER_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(KERbud_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(ProdKER_BI)-cdbl(KERbud_BI))/(cdbl(KERbud_BI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(ProdKER_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(KERbud_SBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(ProdKER_SBI)-cdbl(KERbud_SBI))/(cdbl(KERbud_SBI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(KERbud_THN,2)%></th>
	</tr>
	<tr>
	    
		<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">PRODUK (KG)</th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(Palm_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(PALMbud_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(Palm_BI)-cdbl(Palmbud_BI))/(cdbl(Palmbud_BI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(Palm_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(PALMbud_SBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;"  align="right" ><%=formatnumber((cdbl(Palm_SBI)-cdbl(Palmbud_SBI))/(cdbl(Palmbud_SBI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;"  align="right" colspan="3"><%=formatnumber(PALMbud_THN,2)%></th>
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
	       if v_txtjob <> "--SUMMARY--"  and v_txtjob <> "--ALL--" and v_txtjob <> "" then
				whereCons = wherecons & " And jobcode = '" & v_txtjob & "' "
			end if
	Dim strdtl , strgrp ,cdgrp,cdmesin,cdgrpdesc  ,noline , subAmtgrpBI , GRAmtgrpBI, subAmtgrpSBI , GRAmtgrpSBI , grpDesc,subbudBI,subbudSBI,Subbudthn,GRbudBI,GRbudSBI,GRbudthn
	if     v_vJenisReport  = "Group" or len(v_vJenisReport ) < 1  then
			strdtl = " select " & strFieldCons & " jobcode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from  ( " & _
					 " 	select t.ba fcba,t.stationcode jobcode, t.STATIONCODE_DESCRIPTION DESCRIPTION, sum(t.amount) BI, 0 sbi, " & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.ba,2013,1,1,'','','','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.ba,2013,1,1,'','','','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.ba,2013,1,1,'','','','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY'))  budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by  t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION  " & _
					 "	union " & _
					 "	select t.ba fcba,t.stationcode jobcode, t.STATIONCODE_DESCRIPTION DESCRIPTION, 0 BI, sum(t.amount) SBI ,0 budBI,0 budSBI,0 budthn  " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by  t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION  " & _
					 "	 ) group by fcba,jobcode,DESCRIPTION  " 	
            
              strdtl = "select GroupPlantation " & selectCons & ", jobcode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by jobcode,DESCRIPTION,GroupPlantation" & selectCons & "  Order by jobcode,DESCRIPTION,GroupPlantation" & orderCons  	
			
	elseif  v_vJenisReport  = "Job"  then  	
           	strdtl = " SELECT " & strFieldCons & " stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn FROM ( " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION DESCRIPTION , sum(t.amount) BI, 0 SBI ," & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.ba,2013,1,1,'',t.JobCode,'','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.ba,2013,1,1,'',t.JobCode,'','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.ba,2013,1,1,'',t.JobCode,'','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY'))  budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)   " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "	UNION " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION DESCRIPTION , 0 BI, sum(t.amount) SBI ,0 budBI,0 budSBI,0 budthn  " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "  ) " & _
					 "	GROUP BY  fcba,stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION  "
			strdtl = "select GroupPlantation " & selectCons & ", stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION,GroupPlantation" & selectCons & "  Order by stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION,GroupPlantation" & orderCons  	
		 
	elseif  v_vJenisReport  = "Mechine"  then
			strdtl = " SELECT " & strFieldCons & " stationcode,STATIONCODE_DESCRIPTION,MACHINECODE jobcode,DESCRIPTION, SUM(BI) BI,SUM(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  FROM ( " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.MACHINECODE,t.MACHINECODE_DESCRIPTION DESCRIPTION,t.JobCode,t.ACCOUNTCODE_DESCRIPTION  , sum(t.amount) BI, 0 SBI ," & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.ba,2013,1,1,'',t.JobCode,'','','',t.MACHINECODE,t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.ba,2013,1,1,'',t.JobCode,'','','',t.MACHINECODE,t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.ba,2013,1,1,'',t.JobCode,'','','',t.MACHINECODE,t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY'))  budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION, t.MACHINECODE,t.MACHINECODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "	UNION " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.MACHINECODE,t.MACHINECODE_DESCRIPTION DESCRIPTION, t.JobCode,t.ACCOUNTCODE_DESCRIPTION , 0 BI, sum(t.amount) SBI,0 budBI,0 budSBI,0 budthn  " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)   " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION, t.MACHINECODE,t.MACHINECODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "	 ) " & _
					 "	GROUP BY  fcba,stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,DESCRIPTION " 	
			
             strdtl = "select GroupPlantation " & selectCons & ", stationcode,STATIONCODE_DESCRIPTION,jobcode,DESCRIPTION, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION,GroupPlantation" & selectCons & "  Order by stationcode,STATIONCODE_DESCRIPTION,JobCode,DESCRIPTION,GroupPlantation" & orderCons  	
			
	elseif  v_vJenisReport  = "Combine"  then
			strdtl = " SELECT " & strFieldCons & " stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION, SUM(BI) BI,SUM(SBI) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  FROM ( " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.MACHINECODE,t.MACHINECODE_DESCRIPTION ,t.JobCode,t.ACCOUNTCODE_DESCRIPTION  DESCRIPTION, sum(t.amount) BI, 0 SBI ," & _
					 "  budget_transaction_mill_api.GetBudgetMonth(t.ba,2013,1,1,'',t.JobCode,'','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.ba,2013,1,1,'',t.JobCode,'','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.ba,2013,1,1,'',t.JobCode,'','','','',t.stationcode,to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY'))  budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month = lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION, t.MACHINECODE,t.MACHINECODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION  " & _
					 "	UNION " & _
					 "	select t.ba fcba,t.stationcode, t.STATIONCODE_DESCRIPTION , t.MACHINECODE,t.MACHINECODE_DESCRIPTION ,t.JobCode,t.ACCOUNTCODE_DESCRIPTION  DESCRIPTION, 0 BI, sum(t.amount) SBI,0 budBI,0 budSBI,0 budthn " & _
					 "	from VIEW_ACCOUNTREPORT_MILL t where t.period_year = '" & v_txtyear  & "' and t.period_month <= lpad('" & v_txtmonth & "',2,0)  " & _
					 "	group by t.ba,t.period_year,t.period_month,t.stationcode, t.STATIONCODE_DESCRIPTION, t.MACHINECODE,t.MACHINECODE_DESCRIPTION , t.JobCode,t.ACCOUNTCODE_DESCRIPTION " & _ 
					 "	 ) " & _
					 "	GROUP BY  fcba,stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION ORDER BY stationcode " 	
             strdtl = "select GroupPlantation " & selectCons & ", stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION , sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn   From (" & strdtl & " )" & whereCons & " group by stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION ,GroupPlantation" & selectCons & "  Order by stationcode,STATIONCODE_DESCRIPTION,MACHINECODE,MACHINECODE_DESCRIPTION,JOBCODE, DESCRIPTION ,GroupPlantation" & orderCons  	
   					 
	end if		 
			 
				
	'response.write strdtl
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
						          <a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 
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
												<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("Plantation")%>
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
												<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("GroupRegion")%>
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
													<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			<%=rsha("Region")%>
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
												    <a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			  <%=rsha("Company")%>
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
													<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
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
						          <a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 
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
												<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("Plantation")%>
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
												<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("GroupRegion")%>
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
													<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			<%=rsha("Region")%>
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
												    <a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			  <%=rsha("Company")%>
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
													<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
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
				<% ELSE  
				         
				          If  cdgrp <> "" then  %>
								<tr>  						
									<td style="background-color:#F4F4A0" colspan="<%= 2 + nSpanTotal %>"><b>Total &nbsp;<%=cdgrpdesc%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpBI,2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subAmtgrpBI)/cdbl(palm_BI)),2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(subAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></td>	
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudBI,2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudBI)/cdbl(PALMbud_BI)),2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudBI)/cdbl(TBSbud_BI)),2)%></td>
									<% If (cdbl(subbudBI)/cdbl(TBSbud_BI)) = 0 then %>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudBI)/cdbl(TBSbud_BI)) - ( cdbl(subAmtgrpBI)/cdbl(TBSprocess_BI)) ,2)%></td>
									<% Else %>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((( cdbl(subAmtgrpBI)/cdbl(TBSprocess_BI)) - (cdbl(subbudBI)/cdbl(TBSbud_BI))) / (cdbl(subbudBI)/cdbl(TBSbud_BI)),2)%></td>
									<% End If %>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpSBI,2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subAmtgrpSBI)/cdbl(palm_SBI)),2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(subAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></td>	
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudSBI,2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudSBI)/cdbl(TBSbud_SBI)),2)%></td>
									<% If (cdbl(subbudSBI)/cdbl(TBSbud_SBI)) = 0 then %>
										<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudSBI)/cdbl(TBSbud_SBI)) - ( cdbl(subAmtgrpSBI)/cdbl(TBSprocess_SBI)) ,2)%></td>
									<% Else %>
										<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(subbudSBI)/cdbl(TBSbud_SBI)) - ( cdbl(subAmtgrpSBI)/cdbl(TBSprocess_SBI))) / (cdbl(subbudSBI)/cdbl(TBSbud_SBI)),2)%></td>
									<% End If %>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudthn,2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudTHN)/cdbl(PALMbud_THN)),2)%></td>
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudTHN)/cdbl(TBSbud_THN)),2)%></td>		
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
						          <a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 
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
												<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("Plantation")%>
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
												<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("GroupRegion")%>
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
													<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			<%=rsha("Region")%>
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
												    <a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			  <%=rsha("Company")%>
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
													<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
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
							<% if (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI))= 0 then %>
								<td align="right"><%=formatnumber((cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) -(cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)) ,2)%></td>
							<% else %>
								<td align="right"><%=formatnumber(((cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) -(cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)) )/(cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI))*100,2)%></td>
							<% end if %>
							<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
							<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)),2)%></td>
							<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)),2)%></td>
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
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subAmtgrpBI)/cdbl(palm_BI)),2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(subAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></td>	
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudBI,2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudBI)/cdbl(PALMbud_BI)),2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudBI)/cdbl(TBSbud_BI)),2)%></td>
				<% If (cdbl(subbudBI)/cdbl(TBSbud_BI)) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudBI)/cdbl(TBSbud_BI)) - ( cdbl(subAmtgrpBI)/cdbl(TBSprocess_BI)) ,2)%></td>
				<% Else %>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((( cdbl(subAmtgrpBI)/cdbl(TBSprocess_BI)) - (cdbl(subbudBI)/cdbl(TBSbud_BI))) / (cdbl(subbudBI)/cdbl(TBSbud_BI)),2)%></td>
				<% End If %>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpSBI,2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subAmtgrpSBI)/cdbl(palm_SBI)),2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(subAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></td>	
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudSBI,2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudSBI)/cdbl(TBSbud_SBI)),2)%></td>
				<% If (cdbl(subbudSBI)/cdbl(TBSbud_SBI)) = 0 then %>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudSBI)/cdbl(TBSbud_SBI)) - ( cdbl(subAmtgrpSBI)/cdbl(TBSprocess_SBI)) ,2)%></td>
				<% Else %>
					<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(subbudSBI)/cdbl(TBSbud_SBI)) - ( cdbl(subAmtgrpSBI)/cdbl(TBSprocess_SBI))) / (cdbl(subbudSBI)/cdbl(TBSbud_SBI)),2)%></td>
				<% End If %>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudthn,2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudTHN)/cdbl(PALMbud_THN)),2)%></td>
				<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(subbudTHN)/cdbl(TBSbud_THN)),2)%></td>												
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
					 "  budget_transaction_mill_api.GetBudgetMonth(t.ba,2013,1,1,'',t.JobCode,'','','','','',to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budBI, " & _
                     "  budget_transaction_mill_api.GetBudgetUntilMonth(t.ba,2013,1,1,'',t.JobCode,'','','','','',to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY')) budSBI, " & _
                     "  budget_transaction_mill_api.GetBudget_perYear(t.ba,2013,1,1,'',t.JobCode,'','','','','',to_date('01/'||t.period_month||'/'||t.period_year,'DD/MM/YYYY'))  budthn " & _
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
						          <a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 
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
												<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("Plantation")%>
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
												<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 		<%=rsha("GroupRegion")%>
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
													<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			<%=rsha("Region")%>
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
												    <a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>"> 			  <%=rsha("Company")%>
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
													<a href="p_AccountReportMILL.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&vJenisReport=<%=v_vJenisReport%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
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
			 <%
					 rsha.MoveNext
				  wend
			end if
	end if		
	%>
	
	
	<tr>  							
		<td style="background-color:#F4F4A0" colspan="<%= 2 + nSpanTotal %>"><b> Grand Total</b></td>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></b></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpBI)/cdbl(palm_BI)),2)%></b></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></b></td>	
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudBI,2)%></td>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></b></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(TBSbud_BI)),2)%></b></td>
		<% If ( cdbl(GRbudBI)/cdbl(TBSbud_BI)) = 0 then %>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRbudBI)/cdbl(TBSbud_BI)) - (cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI))) ,2)%></b></td>
		<% Else %>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(((cdbl(GRbudBI)/cdbl(TBSbud_BI)) - (cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI))) / ( cdbl(GRbudBI)/cdbl(TBSbud_BI))*100,2)%></b></td>
		<% End If %>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></b></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)),2)%></b></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></b></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudSBI,2)%></td>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></b></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(TBSbud_SBI)),2)%></b></td>
		<% If ( cdbl(GRbudBI)/cdbl(TBSbud_SBI)) = 0 then %>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(TBSbud_SBI)) - ( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)) ,2)%></b></td>
		<% Else %>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber((( cdbl(GRbudSBI)/cdbl(TBSbud_SBI)) - ( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI))) / ( cdbl(GRbudSBI)/cdbl(TBSbud_SBI))*100,2)%></b></td>
		<% End If %>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudthn,2)%></b></td>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(PALMbud_THN)),2)%></b></td>
		<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(TBSbud_THN)),2)%></b></td>												
	</tr>
 </tbody>
</table>
</form>
<!--#include file="bottom.Asp"-->