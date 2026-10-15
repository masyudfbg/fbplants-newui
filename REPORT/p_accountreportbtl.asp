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
Dim SprevCode,Sprevjobname, sprevdetail
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

Dim v_txttype : v_txttype = Request("txttype")
if len(trim(v_txttype)) = 0 then v_txttype = "rekap"

Dim currjob : currjob = ""
Dim prevjob : prevjob = ""
Dim currjobdesc : currjobdesc = ""
Dim prevjobdesc : prevjobdesc = ""
Dim sub_act_bi : sub_act_bi = 0
Dim sub_act_sbi : sub_act_sbi = 0
Dim sub_bud_bi : sub_bud_bi = 0
Dim sub_bud_sbi : sub_bud_sbi = 0
Dim sub_bud_ann : sub_bud_ann = 0

Dim sub_rpprod_act_bi : sub_rpprod_act_bi = 0
Dim sub_rpprod_act_sbi : sub_rpprod_act_sbi = 0
Dim sub_rpprod_bud_bi : sub_rpprod_bud_bi = 0
Dim sub_rpprod_bud_sbi : sub_rpprod_bud_sbi = 0
Dim sub_rpprod_bud_ann : sub_rpprod_bud_ann = 0

Dim sub_rptbs_act_bi : sub_rptbs_act_bi = 0
Dim sub_rptbs_act_sbi : sub_rptbs_act_sbi = 0
Dim sub_rptbs_bud_bi : sub_rptbs_bud_bi = 0
Dim sub_rptbs_bud_sbi : sub_rptbs_bud_sbi = 0
Dim sub_rptbs_bud_ann : sub_rptbs_bud_ann = 0

Dim sub_var_act_bi : sub_var_act_bi = 0
Dim sub_var_act_sbi : sub_var_act_sbi = 0

Dim addcolspan : addcolspan = 0
            
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
				addcolspan = addcolspan + 1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
				addcolspan = addcolspan + 1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
				addcolspan = addcolspan + 1
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
				addcolspan = addcolspan + 1
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
				addcolspan = addcolspan + 1
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
			
			' if len(trim(v_txtDivision)) > 0 then 
			   ' if v_txtDivision <> "--SUMMARY--" then
			      ' if v_txtDivision <>  "--ALL--"  then
					' whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      ' end if
                  ' selectCons = selectCons &  " , DIVISION "
			   ' else
			      ' nSpanTotal = nSpanTotal - 1 	  
			   ' end if
			' end if
            
			' if len(trim(v_txtPlantingDate)) > 0 then 
				' if v_txtPlantingDate <> "--SUMMARY--"   then
				   ' if v_txtPlantingDate <>  "--ALL--" then
					  ' whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   ' end if
				   ' selectCons = selectCons &  " , PLANTINGYEAR "
				' else
			      ' nSpanTotal = nSpanTotal - 1 	  
			   ' end if
				
			' end if
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
		<form name="frmBTL" method="post" action="<%=ScriptName%>">
			<div class="printhide">
				<!--#include file="FormSelectNewSumM.html"-->
				<!--#include file="FromSelectHDRBTLMILL.html"-->
			</div>
		</form>
	</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
			<th colspan="22" style="background-color:#FFFFFF; border:none;">
				<div id="header-wrapper">
				<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
						<% If  CompanyName <>  "" then %>
							<h4>PT.<%=CompanyName%></h4> 
						<% End if%>
						<% If  FcbaName <> ""  then %>
							   <h5><%=FcbaName%></h5>   
						<% End if%>
				</div>
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
			</th>
		</tr>
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
		 "  select t.fcba, sum(t.cageweight) TBS_PROCESS_BI,0 TBS_PROCESS_SBI from AR_SHIFTREPORT t " & _  
	     "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.productiondate,'MM') = lpad('" & v_txtmonth & "',2,0) " & _
		 "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  " & _
		 "  union all " & _
		 "  select t.fcba, 0 TBS_PROCESS_BI, sum(t.cageweight) TBS_PROCESS_SBI from AR_SHIFTREPORT t  " & _
		 "	where to_char(t.productiondate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.productiondate,'MM') <= lpad('" & v_txtmonth & "',2,0) " & _
         "  group by t.fcba,to_char(t.productiondate,'YYYYMM')  ) group by fcba )  " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' group by GroupPlantation"
		  
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

' strtbs = " select fcba, sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 ' " select t.fcba,  sum(t.quantity) Produksi_CPO_BI,0 Produksi_CPO_SBI from PRODUCTSTORAGEREADING t  " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='CPO')  and to_char(t.readingdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.readingdate,'MM') = lpad('" & v_txtmonth & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba, to_char(t.readingdate,'YYYYMM') " & _
		 ' " union all " & _
		 ' " select t.fcba, 0 Produksi_CPO_BI,sum(t.quantity) Produksi_CPO_SBI from PRODUCTSTORAGEREADING t " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='CPO')  and to_char(t.readingdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.readingdate,'MM') <= lpad('" & v_txtmonth & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba, to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select GroupPlantation ,sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  From ( " & _
         " select  " & strFieldCons & " sum(Produksi_CPO_BI) Produksi_CPO_BI,sum(Produksi_CPO_SBI) Produksi_CPO_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_CPO_BI, 0 Produksi_CPO_SBI from AR_MILL_PRODUCTION t  " & _
		 " where t.product_code='CPO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union all " & _
		 " select t.fcba,0 Produksi_CPO_BI, sum(t.production_today) Produksi_CPO_SBI from AR_MILL_PRODUCTION t " & _
		 " where t.product_code='CPO'  and to_char(t.tdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba )  " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' group by GroupPlantation"	 
		  
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

' strtbs = " select fcba, sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI from ( " & _
	     ' " select t.fcba, sum(t.quantity) Produksi_KER_BI, 0 Produksi_KER_SBI from PRODUCTSTORAGEREADING t  " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='Kernel')   and to_char(t.readingdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.readingdate,'MM') = lpad('" & v_txtmonth & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba,to_char(t.readingdate,'YYYYMM')  " & _
		 ' " union all " & _
		 ' " select t.fcba, 0 Produksi_KER_BI, sum(t.quantity) Produksi_KER_SBI from PRODUCTSTORAGEREADING t " & _
		 ' " where t.pstorage_storagecode in ( select STORAGECODE  from PRODUCTSTORAGE where IsQtyProd='Y' and PRODUCT_PRODUCTCODE='Kernel')  and to_char(t.readingdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.readingdate,'MM') <= lpad('" & v_txtmonth & "',2,0)  and t.fcba ='" & session("fcba")  & "' " & _
		 ' " group by t.fcba,to_char(t.readingdate,'YYYYMM') ) group by fcba " 
		 
strtbs = " select GroupPlantation ,sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI From ( " & _
         " select " & strFieldCons & " sum(Produksi_KER_BI) Produksi_KER_BI,sum(Produksi_KER_SBI) Produksi_KER_SBI  from ( " & _
		 " select t.fcba, sum(t.production_today) Produksi_KER_BI, 0 Produksi_KER_SBI from AR_MILL_PRODUCTION t  " & _
		 " where t.product_code='KER'  and to_char(t.tdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.tdate,'MM') = lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') " & _
		 " union all " & _
		 " select t.fcba,0 Produksi_KER_BI, sum(t.production_today) Produksi_KER_SBI from AR_MILL_PRODUCTION t " & _
		 " where t.product_code='KER'  and to_char(t.tdate,'YYYY') = '" & v_txtyear & "'  and  to_char(t.tdate,'MM') <= lpad('" & v_txtmonth & "',2,0)   " & _
		 " group by t.fcba, to_char(t.tdate,'YYYYMM') ) group by fcba  )  " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' group by GroupPlantation"	  		 
		 
	  
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
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='TBS'   " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,sum(" & strbudHDR_SBI & ") tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='TBS'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='TBS'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,sum(" & strbudHDR_BI & ") cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='CPO'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,sum(" & strbudHDR_SBI & ") cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='CPO'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba ,0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) cpobudthn,0 kerbudBI, 0 kerbudSBI , 0 kerbudthn  from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='CPO'  " & _
		 "	 group by btm.fcba " & _
		 "	  union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudthn,0 cpobudSBI,sum(" & strbudHDR_BI & ") kerbudBI, 0 kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='KER'  " & _
		 "	 group by btm.fcba " & _
		 "	  union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, sum(" & strbudHDR_SBI & ") kerbudSBI, 0 kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='KER'  " & _
		 "	 group by btm.fcba " & _
		 "	 union all " & _
		 "	 select btm.fcba , 0 TBSbudBI,0 tbsbudSBI,0 tbsbudthn,0 cpobudBI,0 cpobudSBI,0 cpobudthn,0 kerbudBI, 0 kerbudSBI, sum(btm.qty_01+btm.qty_02+btm.qty_03+btm.qty_04+btm.qty_05+btm.qty_06+btm.qty_07+btm.qty_08+btm.qty_09+btm.qty_10+btm.qty_11+btm.qty_12) kerbudthn   from BUDGET_TRANSACTION_MILLQTY btm " & _
		 "	 where btm.budget_year='" & v_txtyear & "' and btm.productcode='KER'  " & _
		 "	 group by btm.fcba ) group by fcba  )  " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' group by GroupPlantation" 
		  
Set rsha = Server.CreateObject("ADODB.Recordset")
rsha.Open strtbs, connstr, adOpenStatic
' response.write strtbs
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
		<%if v_txttype = "detail" then%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 1 + 2 + nSpanTotal %>">&nbsp;</th>
		<%else%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
		<%end if%>
		
		
		<th  style="background-color:#00A300;color:#fff;" colspan="7">BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="7">S/D BULAN INI</th>
		<th  style="background-color:#00A300;color:#fff;" colspan="3" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<%if v_txttype = "detail" then%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 1 + 2 + nSpanTotal %>">&nbsp;</th>
		<%else%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
		<%end if%>
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
		<%if v_txttype = "detail" then%>
			<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 1 + 2 + nSpanTotal %>" align="right">TBS PROSES (KG)</th>
		<%else%>
			<th  style="background-color:#00A300;color:#fff;"  colspan="<%= 2 + nSpanTotal %>" align="right">TBS PROSES (KG)</th>
		<%end if%>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSprocess_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSbud_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(TBSprocess_BI)-cdbl(TBSbud_BI))/(cdbl(TBSbud_BI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSprocess_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSbud_SBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(TBSprocess_SBI)-cdbl(TBSbud_SBI))/(cdbl(TBSbud_SBI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(TBSbud_THN,2)%></th>
		
		
	</tr>
	<tr>
		<%if v_txttype = "detail" then%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 1 + 2 + nSpanTotal %>" align="right">PRODUKSI CPO (KG)</th>
		<%else%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI CPO (KG)</th>
		<%end if%>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(ProdCPO_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(CPObud_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(ProdCPO_BI)-cdbl(CPObud_BI))/(cdbl(CPObud_BI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(ProdCPO_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(CPObud_SBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(ProdCPO_SBI)-cdbl(CPObud_SBI))/(cdbl(CPObud_SBI))*100,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(CPObud_THN,2)%></th>
	</tr>
	<tr>
		<%if v_txttype = "detail" then%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 1 + 2 + nSpanTotal %>" align="right">PRODUKSI KER (KG)</th>
		<%else%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI KER (KG)</th>
		<%end if%>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(ProdKER_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(KERbud_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(ProdKER_BI)-cdbl(KERbud_BI))/(cdbl(KERbud_BI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(ProdKER_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(KERbud_SBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(ProdKER_SBI)-cdbl(KERbud_SBI))/(cdbl(KERbud_SBI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(KERbud_THN,2)%></th>
	</tr>
	<tr>
		<%if v_txttype = "detail" then%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 1 + 2 + nSpanTotal %>" align="right">PRODUK (KG)</th>
		<%else%>
			<th  style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">PRODUK (KG)</th>
		<%end if%>
		
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(Palm_BI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(PALMbud_BI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(Palm_BI)-cdbl(Palmbud_BI))/(cdbl(Palmbud_BI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(Palm_SBI,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(PALMbud_SBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(Palm_SBI)-cdbl(Palmbud_SBI))/(cdbl(Palmbud_SBI))*100,2)%></th>
		<th  style="background-color:#00A300;color:#fff;" align="right" colspan="3"><%=formatnumber(PALMbud_THN,2)%></th>
	</tr>
	<tr   >
	  
	    <th style="background-color:#00A300;color:#fff;" rowspan="2" >Kode</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2" >Jenis Pekerjaan</th>
		<%if v_txttype = "detail" then%>
			<th style="background-color:#00A300;color:#fff;" rowspan="2" >&nbsp;</th>
		<%end if%>
		
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
		<th style="background-color:#00A300;color:#fff;" >TBS</th>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >TBS</th>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >TBS</th>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >TBS</th>
		
		<th style="background-color:#00A300;color:#fff;" >Produk</th>
		<th style="background-color:#00A300;color:#fff;" >TBS</th>
	</tr>
</thead>
<tbody>	
	<%
	Dim strdtl , strgrp ,cdgrp,cdgrpdesc  ,noline ,  subAmtgrpBI , GRAmtgrpBI, subAmtgrpSBI , GRAmtgrpSBI , grpDesc,GRbudBI,GRbudSBI,GRbudthn
	Dim rsjob,strjob,debetjob,kreditjob,notAlokasi

	strjob= "select report_posting_api.getPostingDataDirect('PJJP','ANALISABTL','p_AccountReportBTL.asp','DATADEBET_MILL','JOB') debet, report_posting_api.getPostingDataDirect('PJJP','ANALISABTL','p_AccountReportBTL.asp','DATAKREDIT_MILL','JOB')   kredit  from dual"
	Set rsjob = Server.CreateObject("ADODB.Recordset")
	'response.write strjob
	rsjob.Open strjob, connstr, adOpenStatic
	if Not rsjob.eof then
	   debetjob = rsjob("debet")
	   kreditjob = rsjob("kredit")
	end if
	
		if v_txttype = "detail" then
			strdtl = "  select GroupPlantation " & selectCons & "" & _
			         "         , jobcode" & _
			         "         , description" & _
			         "         , targettype" & _
			         "         , targetcode" & _
			         "         , detail" & _
			         "         , sum(bi) BI" & _
			         "         , sum(sbi) SBI" & _
			         "         , sum(budBI) budBI" & _
			         "         , sum(budSBI) budSBI" & _
			         "         , sum(budthn) budthn" & _
			         "  from ( " & _
                     "  select " & strFieldCons & "" & _
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
					 "   )   " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' group by GroupPlantation, jobcode,description, targettype, targetcode, detail" & selectCons & " order by jobcode " & selectCons & ", detail"
	
		
		
			 strdtl =  strdtl  + strsql1  + strgrp
		else
			strdtl = "  select GroupPlantation " & selectCons & "  ,jobcode,description, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from ( " & _
                     "  select " & strFieldCons & " jobcode,description, sum(bi) BI,sum(sbi) SBI,sum(budBI) budBI,sum(budSBI) budSBI,sum(budthn) budthn  from ( " & _
					 "         SELECT data.*, jj.fcname description" & _
					 "         FROM (" & _
					 "             SELECT fcba, status grp, jobcode, SUM(act_bi) bi, SUM(bud_bi) budbi, SUM(act_sbi) sbi, SUM(bud_sbi) budsbi, SUM(bud_annual) budthn" & _
					 "             FROM ar_btl" & _
					 "             WHERE year = '" & v_txtyear & "'" & _
					 "                   AND month = '" & v_txtmonth & "'" & _
					 "             GROUP BY fcba, status, jobcode" & _
					 "         ) data" & _
					 "         LEFT JOIN job jj ON (data.fcba = jj.fcba AND data.jobcode = jj.fccode)" & _
					 "	 ) " & _
					 "	group by fcba, jobcode,description )   " & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M' group by GroupPlantation, jobcode,description " & selectCons & " order by jobcode " & selectCons 
	
		
		
			 strdtl =  strdtl  + strsql1  + strgrp
		end if
	' response.write strdtl
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
    '
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
	GRAmtgrpBI ="0"
	GRAmtgrpSBI ="0"
	if Not rsha.eof then
		  while Not rsha.eof
		    
			 IF rsha("JobCode") <> "6610600"	 THEN  %>
			 
			<%
				if v_txttype = "detail" then
					currjob = rsha("JOBCODE")
					currjobdesc = rsha("description")
					if currjob <> prevjob then
						if len(trim(prevjob)) = 0 then
							sub_act_bi = 0
							sub_act_sbi = 0
							sub_bud_bi = 0
							sub_bud_sbi = 0
							sub_bud_ann = 0
							sub_rpprod_act_bi = 0
							sub_rpprod_act_sbi = 0
							sub_rpprod_bud_bi = 0
							sub_rpprod_bud_sbi = 0
							sub_rpprod_bud_ann = 0
							sub_rptbs_act_bi = 0
							sub_rptbs_act_sbi = 0
							sub_rptbs_bud_bi = 0
							sub_rptbs_bud_sbi = 0
							sub_rptbs_bud_ann = 0
							sub_var_act_bi = 0
							sub_var_act_sbi = 0
						else
			%>
			
			<%
							sub_rpprod_act_bi = 0
							sub_rpprod_act_sbi = 0
							sub_rpprod_bud_bi = 0
							sub_rpprod_bud_sbi = 0
							sub_rpprod_bud_ann = 0
							sub_rptbs_act_bi = 0
							sub_rptbs_act_sbi = 0
							sub_rptbs_bud_bi = 0
							sub_rptbs_bud_sbi = 0
							sub_rptbs_bud_ann = 0
							sub_var_act_bi = 0
							sub_var_act_sbi = 0
							if Palm_BI <> 0 then
								sub_rpprod_act_bi = sub_act_bi/Palm_BI
							else
								sub_rpprod_act_bi = 0
							end if
							if Palm_SBI <> 0 then
								sub_rpprod_act_sbi = sub_act_sbi/Palm_SBI
							else
								sub_rpprod_act_sbi = 0
							end if
							
							if cdbl(TBSprocess_BI) <> 0 then
								sub_rptbs_act_bi = sub_act_bi/cdbl(TBSprocess_BI)
							else
								sub_rptbs_act_bi = 0
							end if
							if cdbl(TBSprocess_SBI) <> 0 then
								sub_rptbs_act_sbi = sub_act_sbi/cdbl(TBSprocess_SBI)
							else
								sub_rptbs_act_sbi = 0
							end if
							
							if PALMbud_BI <> 0 then
								sub_rpprod_bud_bi = sub_bud_bi/PALMbud_BI
							else
								sub_rpprod_bud_bi = 0
							end if
							
							if PALMbud_SBI <> 0 then
								sub_rpprod_bud_sbi = sub_bud_sbi/PALMbud_SBI
							else
								sub_rpprod_bud_sbi = 0
							end if
							
							if cdbl(TBSbud_BI) <> 0 then
								sub_rptbs_bud_bi = sub_bud_bi/cdbl(TBSbud_BI)
							else
								sub_rptbs_bud_bi = 0
							end if
							if cdbl(TBSbud_SBI) <> 0 then
								sub_rptbs_bud_sbi = sub_bud_sbi/cdbl(TBSbud_SBI)
							else
								sub_rptbs_bud_sbi = 0
							end if
							
							if cdbl(PALMbud_THN) <> 0 then
								sub_rpprod_bud_ann = sub_bud_ann/cdbl(PALMbud_THN)
							else
								sub_rpprod_bud_ann = 0
							end if
							
							if cdbl(TBSbud_THN) <> 0 then
								sub_rptbs_bud_ann = sub_bud_ann/cdbl(TBSbud_THN)
							else
								sub_rptbs_bud_ann = 0
							end if
							
							if sub_rptbs_bud_bi = 0 and sub_rptbs_act_bi <> 0 then
								sub_var_act_bi = -100
							elseif sub_rptbs_bud_bi <> 0 and sub_rptbs_act_bi = 0 then
								sub_var_act_bi = 100
							elseif sub_rptbs_bud_bi = 0 and sub_rptbs_act_bi = 0 then
								sub_var_act_bi = 0
							elseif sub_rptbs_bud_bi <> 0 and sub_rptbs_act_bi <> 0 then
								sub_var_act_bi = Round(((sub_rptbs_bud_bi - sub_rptbs_act_bi)/ sub_rptbs_bud_bi)*100, 2)
							else
								sub_var_act_bi = 0
							end if
							
							if sub_rptbs_bud_sbi = 0 and sub_rptbs_act_sbi <> 0 then
								sub_var_act_sbi = -100
							elseif sub_rptbs_bud_sbi <> 0 and sub_rptbs_act_sbi = 0 then
								sub_var_act_sbi = 100
							elseif sub_rptbs_bud_sbi = 0 and sub_rptbs_act_sbi = 0 then
								sub_var_act_sbi = 0
							elseif sub_rptbs_bud_sbi <> 0 and sub_rptbs_act_sbi <> 0 then
								sub_var_act_sbi = Round(((sub_rptbs_bud_sbi - sub_rptbs_act_sbi)/ sub_rptbs_bud_sbi)*100, 2)
							else
								sub_var_act_sbi = 0
							end if
			%>
							<tr bgcolor="#F4F4A0">
								<td colspan="<%=3 + addcolspan%>"><strong>Subtotal <%=prevjobdesc%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_bud_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_bud_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_var_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_bud_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_bud_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_var_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_ann, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_bud_ann, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_bud_ann, 2)%></strong></td>
							</tr>
			<%
							sub_act_bi = 0
							sub_act_sbi = 0
							sub_bud_bi = 0
							sub_bud_sbi = 0
							sub_bud_ann = 0
							sub_rpprod_act_bi = 0
							sub_rpprod_act_sbi = 0
							sub_rpprod_bud_bi = 0
							sub_rpprod_bud_sbi = 0
							sub_rpprod_bud_ann = 0
							sub_rptbs_act_bi = 0
							sub_rptbs_act_sbi = 0
							sub_rptbs_bud_bi = 0
							sub_rptbs_bud_sbi = 0
							sub_rptbs_bud_ann = 0
							sub_var_act_bi = 0
							sub_var_act_sbi = 0
						end if
						sub_act_bi = sub_act_bi + cdbl(rsha("bi"))
						sub_act_sbi = sub_act_sbi + cdbl(rsha("sbi"))
						sub_bud_bi = sub_bud_bi + cdbl(rsha("budBI"))
						sub_bud_sbi = sub_bud_sbi + cdbl(rsha("budSBI"))
						sub_bud_ann = sub_bud_ann + cdbl(rsha("budthn"))
					else
						sub_act_bi = sub_act_bi + cdbl(rsha("bi"))
						sub_act_sbi = sub_act_sbi + cdbl(rsha("sbi"))
						sub_bud_bi = sub_bud_bi + cdbl(rsha("budBI"))
						sub_bud_sbi = sub_bud_sbi + cdbl(rsha("budSBI"))
						sub_bud_ann = sub_bud_ann + cdbl(rsha("budthn"))
					end if
				end if
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
						   <%if v_txttype = "detail" then%>
								<td><%=rsha("detail")%></td>
						   <% End If %>
						   <%
						   SprevCode   = rsha("JOBCODE")
						   Sprevjobname = rsha("description")
						   if v_txttype = "detail" then sprevdetail = rsha("detail")
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
												<a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
												<a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
													<a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
												    <a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
													<a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
						<td align="right"><%=formatnumber(((cdbl(rsha("BI"))/cdbl(TBSprocess_BI))- (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)) ),2)%></td>
					<% else %>
						<td align="right"><%=formatnumber(((cdbl(rsha("BI"))/cdbl(TBSprocess_BI)) - (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)) )/(cdbl(rsha("budBI"))/cdbl(TBSbud_BI))*100,2)%></td>
					<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(palm_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)),2)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)),2)%></td>
					<% if (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) = 0 then %>
						<td align="right"><%=formatnumber(((cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI))- (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) ),2)%></td>
					<% else %>
						<td align="right"><%=formatnumber(((cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)) - (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) )/(cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI))*100,2)%></td>
					<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)),2)%></td>
				</tr>	
            <% ELSE  
			  
				   if notAlokasi = "T"   then   %>
						<%
						if v_txttype="detail" then
							sub_rpprod_act_bi = 0
							sub_rpprod_act_sbi = 0
							sub_rpprod_bud_bi = 0
							sub_rpprod_bud_sbi = 0
							sub_rpprod_bud_ann = 0
							sub_rptbs_act_bi = 0
							sub_rptbs_act_sbi = 0
							sub_rptbs_bud_bi = 0
							sub_rptbs_bud_sbi = 0
							sub_rptbs_bud_ann = 0
							sub_var_act_bi = 0
							sub_var_act_sbi = 0
							if Palm_BI <> 0 then
								sub_rpprod_act_bi = sub_act_bi/Palm_BI
							else
								sub_rpprod_act_bi = 0
							end if
							if Palm_SBI <> 0 then
								sub_rpprod_act_sbi = sub_act_sbi/Palm_SBI
							else
								sub_rpprod_act_sbi = 0
							end if
							
							if cdbl(TBSprocess_BI) <> 0 then
								sub_rptbs_act_bi = sub_act_bi/cdbl(TBSprocess_BI)
							else
								sub_rptbs_act_bi = 0
							end if
							if cdbl(TBSprocess_SBI) <> 0 then
								sub_rptbs_act_sbi = sub_act_sbi/cdbl(TBSprocess_SBI)
							else
								sub_rptbs_act_sbi = 0
							end if
							
							if PALMbud_BI <> 0 then
								sub_rpprod_bud_bi = sub_bud_bi/PALMbud_BI
							else
								sub_rpprod_bud_bi = 0
							end if
							
							if PALMbud_SBI <> 0 then
								sub_rpprod_bud_sbi = sub_bud_sbi/PALMbud_SBI
							else
								sub_rpprod_bud_sbi = 0
							end if
							
							if cdbl(TBSbud_BI) <> 0 then
								sub_rptbs_bud_bi = sub_bud_bi/cdbl(TBSbud_BI)
							else
								sub_rptbs_bud_bi = 0
							end if
							if cdbl(TBSbud_SBI) <> 0 then
								sub_rptbs_bud_sbi = sub_bud_sbi/cdbl(TBSbud_SBI)
							else
								sub_rptbs_bud_sbi = 0
							end if
							
							if cdbl(PALMbud_THN) <> 0 then
								sub_rpprod_bud_ann = sub_bud_ann/cdbl(PALMbud_THN)
							else
								sub_rpprod_bud_ann = 0
							end if
							
							if cdbl(TBSbud_THN) <> 0 then
								sub_rptbs_bud_ann = sub_bud_ann/cdbl(TBSbud_THN)
							else
								sub_rptbs_bud_ann = 0
							end if
							
							if sub_rptbs_bud_bi = 0 and sub_rptbs_act_bi <> 0 then
								sub_var_act_bi = -100
							elseif sub_rptbs_bud_bi <> 0 and sub_rptbs_act_bi = 0 then
								sub_var_act_bi = 100
							elseif sub_rptbs_bud_bi = 0 and sub_rptbs_act_bi = 0 then
								sub_var_act_bi = 0
							elseif sub_rptbs_bud_bi <> 0 and sub_rptbs_act_bi <> 0 then
								sub_var_act_bi = Round(((sub_rptbs_bud_bi - sub_rptbs_act_bi)/ sub_rptbs_bud_bi)*100, 2)
							else
								sub_var_act_bi = 0
							end if
							
							if sub_rptbs_bud_sbi = 0 and sub_rptbs_act_sbi <> 0 then
								sub_var_act_sbi = -100
							elseif sub_rptbs_bud_sbi <> 0 and sub_rptbs_act_sbi = 0 then
								sub_var_act_sbi = 100
							elseif sub_rptbs_bud_sbi = 0 and sub_rptbs_act_sbi = 0 then
								sub_var_act_sbi = 0
							elseif sub_rptbs_bud_sbi <> 0 and sub_rptbs_act_sbi <> 0 then
								sub_var_act_sbi = Round(((sub_rptbs_bud_sbi - sub_rptbs_act_sbi)/ sub_rptbs_bud_sbi)*100, 2)
							else
								sub_var_act_sbi = 0
							end if
			%>
							<tr bgcolor="#F4F4A0">
								<td colspan="<%=3 + addcolspan%>"><strong>Subtotal <%=prevjobdesc%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_bud_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_bud_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_var_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_bud_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_bud_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_var_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_bud_ann, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rpprod_bud_ann, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(sub_rptbs_bud_ann, 2)%></strong></td>
							</tr>
						<%end if%>
				   
						  <tr bgcolor="#ffff01">
								<%if v_txttype = "detail" then%>
									<td   colspan="<%= 1 + 2 + nSpanTotal %>"><b>Total Biaya Tak Langsung</td>
								<%else%>
									<td   colspan="<%= 2 + nSpanTotal %>"><b>Total Biaya Tak Langsung</td>
								<%end if%>
								
								<td align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
								<td  align="right"><b><%=formatnumber((cdbl(GRAmtgrpBI)/cdbl(palm_BI)),2)%></td>
								<td  align="right"><b><%=formatnumber(( cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></td>	
								<td align="right"><b><%=formatnumber(GRbudBI,2)%></td>
								<td align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></td>
								<td  align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(TBSbud_BI)),2)%></td>
								<% If ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) = 0 then %>
								<td  align="right"><b><%=formatnumber(((cdbl(GRAmtgrpBI)/cdbl(palm_BI)) - ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) ) ,2)%></td>
								<% Else %>
								<td  align="right"><b><%=formatnumber(((cdbl(GRAmtgrpBI)/cdbl(palm_BI)) - ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) ) / ( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></td>
								<% End If %>
								<td  align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
								<td  align="right"><b><%=formatnumber((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)),2)%></td>
								<td  align="right"><b><%=formatnumber(( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></td>
								<td  align="right"><b><%=formatnumber(GRbudSBI,2)%></td>
								<td align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
								<td  align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(TBSbud_SBI)),2)%></td>
								<% IF ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) = 0 then %>
									<td  align="right"><b><%=formatnumber(((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)) - ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) ) ,2)%></td>
								<% Else %>
									<td  align="right"><b><%=formatnumber(((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)) - ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) ) / ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
								<% End If %>
								<td  align="right"><b><%=formatnumber(GRbudthn,2)%></td>
								<td align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(PALMbud_THN)),2)%></td>
								<td  align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(TBSbud_THN)),2)%></td>														
							</tr>   		  
								 
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
						   <%if v_txttype = "detail" then%>
								<td><%=rsha("detail")%></td>
						   <% End If %>
						   <%
						   SprevCode   = rsha("JOBCODE")
						   Sprevjobname = rsha("description")
						   if v_txttype = "detail" then sprevdetail = rsha("detail")
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
												<a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
												<a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
													<a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
												    <a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
													<a href="p_AccountReportBTL.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
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
						<td align="right"><%=formatnumber(((cdbl(rsha("BI"))/cdbl(TBSprocess_BI))- (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)) ),2)%></td>
					<% else %>
						<td align="right"><%=formatnumber(((cdbl(rsha("BI"))/cdbl(TBSprocess_BI)) - (cdbl(rsha("budBI"))/cdbl(TBSbud_BI)) )/(cdbl(rsha("budBI"))/cdbl(TBSbud_BI))*100,2)%></td>
					<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("SBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(palm_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)),2)%></td>
					<td align="right"><%=formatnumber( cdbl(rsha("budSBI")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(PALMbud_SBI)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)),2)%></td>
					<% if (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) = 0 then %>
						<td align="right"><%=formatnumber(((cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI))- (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) ),2)%></td>
					<% else %>
						<td align="right"><%=formatnumber(((cdbl(rsha("SBI"))/cdbl(TBSprocess_SBI)) - (cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI)) )/(cdbl(rsha("budSBI"))/cdbl(TBSbud_SBI))*100,2)%></td>
					<% end if %>
					<td align="right"><%=formatnumber( cdbl(rsha("budThn")),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(PALMbud_THN)),2)%></td>
					<td align="right"><%=formatnumber( (cdbl(rsha("budThn"))/cdbl(TBSbud_THN)),2)%></td>
				</tr>
              					
               
          <% END IF %>				
    <%        
	                 GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(rsha("SBI")) 
					 GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(rsha("BI")) 
					 GRbudBI  = cdbl(GRbudBI) + cdbl(rsha("budBI"))
					 GRbudSBI = cdbl(GRbudSBI) + cdbl(rsha("budSBI"))
					 GRbudthn  = cdbl(GRbudthn) + cdbl(rsha("budthn"))
					 
					if v_txttype = "detail" then
						prevjob = currjob
						prevjobdesc = currjobdesc
					end if
					
		      rsha.MoveNext
		  wend
	end if    
	%>		
               <tr bgcolor="#ffff01">
					<%if v_txttype = "detail" then%>
						<td  colspan="<%= 1 + 2 + nSpanTotal %>"><b>Total Keseluruhan</td>
					<%else%>
						<td  colspan="<%= 2 + nSpanTotal %>"><b>Total Keseluruhan</td>
					<%end if%>
					
					<td align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
					<td  align="right"><b><%=formatnumber((cdbl(GRAmtgrpBI)/cdbl(palm_BI)),2)%></td>
					<td  align="right"><b><%=formatnumber(( cdbl(GRAmtgrpBI)/cdbl(TBSprocess_BI)),2)%></td>	
					<td align="right"><b><%=formatnumber(GRbudBI,2)%></td>
					<td align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></td>
					<td  align="right"><b><%=formatnumber(( cdbl(GRbudBI)/cdbl(TBSbud_BI)),2)%></td>
					<% If ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) = 0 then %>
					<td  align="right"><b><%=formatnumber(((cdbl(GRAmtgrpBI)/cdbl(palm_BI)) - ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) ) ,2)%></td>
					<% Else %>
					<td  align="right"><b><%=formatnumber(((cdbl(GRAmtgrpBI)/cdbl(palm_BI)) - ( cdbl(GRbudBI)/cdbl(PALMbud_BI)) ) / ( cdbl(GRbudBI)/cdbl(PALMbud_BI)),2)%></td>
					<% End If %>
					<td  align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
					<td  align="right"><b><%=formatnumber((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)),2)%></td>
					<td  align="right"><b><%=formatnumber(( cdbl(GRAmtgrpSBI)/cdbl(TBSprocess_SBI)),2)%></td>
					<td  align="right"><b><%=formatnumber(GRbudSBI,2)%></td>
					<td align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
					<td  align="right"><b><%=formatnumber(( cdbl(GRbudSBI)/cdbl(TBSbud_SBI)),2)%></td>
					<% IF ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) = 0 then %>
						<td  align="right"><b><%=formatnumber(((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)) - ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) ) ,2)%></td>
					<% Else %>
						<td  align="right"><b><%=formatnumber(((cdbl(GRAmtgrpSBI)/cdbl(palm_SBI)) - ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)) ) / ( cdbl(GRbudSBI)/cdbl(PALMbud_SBI)),2)%></td>
					<% End If %>
					<td  align="right"><b><%=formatnumber(GRbudthn,2)%></td>
					<td align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(PALMbud_THN)),2)%></td>
					<td  align="right"><b><%=formatnumber(( cdbl(GRbudthn)/cdbl(TBSbud_THN)),2)%></td>														
				</tr>  
         </tbody>	
</table>
<!--#include file="bottom.Asp"-->