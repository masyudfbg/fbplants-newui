<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%

Dim v_txtDivision, v_txtPlantingDate, v_txtBlok
Dim v_txtyear,v_txtmonth
Dim txtDivision, txtBlok, txtPlantingDate
Dim v_vstatus,v_vCROPTYPE,v_vJenisReport, v_vOwnership,v_txtCostElm,v_txtjob,v_txtASPCode
Dim tempyear,tempmonth,no
Dim tempDivision, tempBlok, tempPlantingDate 
Dim SprevCode,Sprevjobname
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation,tempjob
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, rsBa,strsql1,strsql2,strsql3,strsql4,strsql5,strsql6,strsql7,strsql8
Dim ScriptName,nSpanTotal
DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType,rsJob
Dim rsSQLPreParameter,strSQLPreParameter,txtdate,txtmonth, 	txtyear
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
           
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
			v_vstatus             = Request("vStatus")
			v_vCROPTYPE           = Request("vCROPTYPE")
			v_vJenisReport        = request("vJenisReport")
			v_vOwnership	      = Request("vOwnership")
			v_txtCostElm		  = Request("txtCostElm")	
			v_txtjob			  = Request("txtjob")
			v_txtASPCode          = Request("txtASPCode")
			
			nSpanTotal = 9
			 
					
			Dim selectFields,strFieldCons,selectCons,whereCons,orderCons
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut,businessunit_api.Isactive(fcba) fcbaActive, " 

						   
			wherecons = " where company is not null and fcbaActive='TRUE'  "
			
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
	        if v_txtCostElm ="Tampil" then
			    selectCons = selectCons & ", sourcetype,sourcename"
				orderCons  = orderCons & ",sourcetype "
			else
			    nSpanTotal = nSpanTotal - 2 
			end if   	
			
			if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
				  orderCons  = orderCons & " , DIVISION "
			else
			      nSpanTotal = nSpanTotal - 1 	  
			end if
			
			
            
			if  v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""  then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
				   orderCons  = orderCons & " , PLANTINGYEAR "
			else
			      nSpanTotal = nSpanTotal - 1 	  
			end if
				
			
			 ' if len(trim(v_txtBlok)) > 0 then 
				 ' if v_txtBlok <> "--SUMMARY--"  then
				    ' if v_txtBlok <>  "--ALL--" then
					   ' whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
				    ' end if
				    ' selectCons = selectCons & " , FIELDCODE "
                 ' else
			       ' nSpanTotal = nSpanTotal - 1 	  
			    ' end if				   	
             ' end if
			 if len(trim(v_vJenisReport))=0 then
			   v_vJenisReport = "Group" 
			end if  
			' if len(trim(v_vstatus)) = 0 then
			'    v_vstatus ="TM" 
           '  end if	
             if len(trim(v_vCROPTYPE)) = 0 then	
			    v_vCROPTYPE ="OP" 
             end if			 
			 if len(trim(v_vOwnership)) = 0 then	
			    v_vOwnership ="INTI" 
             end if	 
			 if len(trim(v_txtyear)) = 0 then
			   txtdate	= date
			end if
			  if len(trim(v_txtyear)) = 0 then  
			   strSQLPreParameter = "select to_char(to_date('" & txtdate & "', 'MM/DD/YYYY'),'YYYY') as CurrentYear, " & _
					                "to_char(to_date('" & txtdate & "', 'MM/DD/YYYY'),'MM') as CurrentPeriod from dual "
			    Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
				rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 		
				if not rsSQLPreParameter.EOF then
					v_txtyear = rsSQLPreParameter("CurrentYear")
					v_txtmonth = rsSQLPreParameter("CurrentPeriod")
				end if
				rsSQLPreParameter.Close
				set rsSQLPreParameter = nothing	
				If v_txtmonth = 1 then
				   v_txtmonth = "12"
                   v_txtyear =  cstr(cint(v_txtyear) - 1)
				else
				   v_txtmonth = cstr(cint(v_txtmonth) - 1) 
                end if
                txtmonth=v_txtmonth
				txtyear=v_txtyear
            end if			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " TT" & v_txtPlantingDate
			'response.write  "YEAR " & v_txtyear & "  month " & v_txtmonth& "  status " & v_vstatus & " crop " & v_vCROPTYPE  & " owner " & v_vOwnership & " jenis " & len(trim(v_vJenisReport))  & " jenis2 " & v_vJenisReport

%>
<html>

<head>
<title>iPla'S Reporting - Account Report Biaya TM/TBM</title>
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
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
</style>
</head>

<body onload="javascript:filtering()">
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
		<div class="printhide">
		
		<form action="<%=ScriptName%>" method="post">
		<!--#include file="menutop.asp"-->
		<!--#include file="FormSelectNewsum.html"-->
		<!--#include file="FormSelectTMTBM.html"-->
		</form>
		</div>
     </div>   
<table border="0" id="product-table" class="xthf-site" width="100%">
	<thead>
        <tr>
          <th colspan="19" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">	    
		<div id="header">
		   <% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			<% If  v_txtASPCode <> "NONE"  then %>
				   <h5><%=v_txtASPCode%></h5>   
			<% End if%>
		</div>
		<div id="subheader" class="centered">
			<% Dim ia,thntnm
			   ia = v_txtmonth
			   if ia ="" then
				  ia="0" 
			   end if
			   if   v_txtPlantingDate = "0"  then
					thntnm = "ALL"
			   else		
					thntnm = v_txtPlantingDate
			   end if
					
			%>
			<%  IF v_vstatus  = "TM" THEN %>		
			   <h3>LAPORAN BIAYA TANAMAN MENGHASILKAN (TM) </h3>
			<%  ELSE %>		
			   <h3>LAPORAN BIAYA TANAMAN BELUM MENGHASILKAN (TBM) </h3>   
			<%  END IF %>	   
			<% IF v_txtmonth <> "" then %>
			<h5>PERIODE     : <%=ucase(MonthName(cint(ia)))%> &nbsp; <%=v_txtyear%></h5>
			<% END IF %>
			<h5>CROP        : <%=request.Form("vCROPTYPE")%> &nbsp; KEPEMILIKAN : <%=request.Form("vOwnership")%></h5>
			<h5>DIVISI	    : <%=request.Form("vDivision")%> &nbsp; TAHUN TANAM : <%=thntnm%></h5>
			<h5>SOURCE      : <%=ucase(request.Form("vJenisReport"))%></h5>
		</div>
        </th>
    </tr>
	
	
<%


' Budget HA dan luas HA

Dim strha, rsha,luasHA,budHA,ProdTon ,prodkg, strprod,cur_month,strbud_BI,strbud_SBI,x1,strbudprod_BI,strbudprod_SBI
	    strbud_BI = "a.Budget_01" 
	    strbud_SBI = "a.Budget_01"
		strbudprod_BI = "f.qty_01"
		strbudprod_SBI = "f.qty_01"
		
		cur_month = v_txtMonth 
		
		IF Len(cur_month) > 0 then

			strbud_BI = "" 
	        strbud_SBI = ""
			strbudprod_BI = ""
		    strbudprod_SBI = ""
			x1=1
			
			Do While x1<= cint(cur_month) 'Output the values from 0 to 10
			  'response.write(i & "<br \>")
			  'response.write  " Budget_" & Lpad(x1,"0",2) & " +"
			  strbud_SBI = strbud_SBI & "nvl(a.Budget_" & Lpad(x1,"0",2) & ",0) +"
		      strbudprod_SBI = strbudprod_SBI &  "nvl(f.qty_" & Lpad(x1,"0",2) &",0) +"
			  x1 = x1 + 1 'increment the value of i for next time loop executes
			Loop

			strbud_BI = "a.Budget_" & Lpad(cur_month,"0",2)
			strbud_SBI = Mid(strbud_SBI,1,Len(strbud_SBI)-1)
			strbudprod_BI = "f.qty_" & Lpad(cur_month,"0",2)
			strbudprod_SBI = Mid(strbudprod_SBI,1,Len(strbudprod_SBI)-1)
            'response.write  strbud_SBI	& "<br \>"		
	        'response.write  strbud_BI	& "<br \>"	
		END IF	
		
luasHA=0
budHA=0


		
' strha  = " SELECT " & strFieldCons & "division,PLANTINGYEAR, DECODE(SUM (ha), NULL,1,0,1,SUM (ha)) HA,DECODE(SUM (bud_ha), NULL,1,0,1,SUM (bud_ha)) bud_ha FROM " & _
		 ' "	( SELECT fcba, croptype, ownership, division, status, TO_CHAR(plantingdate,'YYYY') PLANTINGYEAR, SUM (hectarageplanted) ha,0 bud_ha FROM field  " & _
		 ' "	  GROUP BY fcba, croptype, ownership, division, status, TO_CHAR(plantingdate, 'YYYY') " & _
		 ' "	  UNION ALL" & _
		 ' "	  select fcba,croptype,ownership,division,status,PLANTINGYEAR,0 ha, sum(hectarageplanted)  bud_ha  from  " & _
		 ' "		( select  distinct t.fcba,t.croptype,t.ownership,t.division,t.status,to_char(t.plantingdate,'YYYY') PLANTINGYEAR, hectarageplanted " & _
		 ' "		  from BUDGET_TRANSACTION_ESTATE t ) " & _
		 ' "	  GROUP BY fcba, croptype, ownership, division, status, PLANTINGYEAR " & _
		 ' "	) WHERE STATUS = '" & v_vStatus  & "' AND CROPTYPE = '"& v_vCROPTYPE &"' "         
if v_txtyear = "" then
   v_txtyear = 2010
end if
		
strha  = " SELECT " & strFieldCons & "division,PLANTINGYEAR, SUM (nvl(ha,0)) ha, SUM (nvl(bud_ha,0)) bud_ha FROM " & _
		 "	( SELECT fcba,field_api.GetValid_CropType(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )  croptype,field_api.GetValid_ownership(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )  ownership, field_api.getValid_Afdeling(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) division, field_api.GetValid_status(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) status,field_api.GetValid_Field_Info_01(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) field_info_01, field_api.getValid_Plantingyear(fccode,fcba,last_day(to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )) PLANTINGYEAR, SUM (nvl(field_api.GetValid_hectarageplanted(fccode,fcba,last_day(to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )),0)) ha,0 bud_ha FROM field  " & _
		 "    where  field_api.GetValid_Activation(fccode,fcba,last_day(to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy')) ) = 'Y' " & _
		 "	  GROUP BY fcba, field_api.GetValid_CropType(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ), field_api.GetValid_ownership(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ), field_api.getValid_Afdeling(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ),  field_api.GetValid_status(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ),field_api.GetValid_Field_Info_01(fccode,fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ),  field_api.getValid_Plantingyear(fccode,fcba,last_day(to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )) " & _
		 "	  UNION ALL" & _
		 "	  select fcba,croptype,ownership,division,status,FIELD_INFO_01,PLANTINGYEAR,0 ha, sum(hectarageplanted)  bud_ha  from  " & _
		 "		( select  t.fcba,t.croptype,t.ownership,t.division,t.status,t.FIELD_INFO_01,to_char(t.plantingdate,'YYYY') PLANTINGYEAR, hectarageplanted " & _
		 "		  from Budget_Field t where t.budget_year = " & v_txtyear & ") " & _
		 "	      GROUP BY fcba, croptype, ownership, division, status,FIELD_INFO_01, PLANTINGYEAR " & _
		 "	) WHERE STATUS = '" & v_vStatus  & "' AND CROPTYPE = '"& v_vCROPTYPE &"' "  
	   
		  if  v_vOwnership <> "ALL" then
			  strsql2 = "  and ownership ='" & v_vOwnership & "'" 
		  else
			  strsql2= " "
		  end if
		  if v_txtASPCode <> "NONE"  and v_txtASPCode <> "" then
			  strsql2 =  strsql2 + " AND Field_Info_01 = '" &  v_txtASPCode & "'" 
		 else
               strsql2 =  strsql2 + ""		 
		 end if
		    ' if   v_txtPlantingDate <> "0"  then
			   ' ' strsql3 = " and PLANTINGYEAR ='" & v_txtPlantingDate & "'"
		    ' else
			   ' ' strsql3 = " "
		    ' end if
		    ' if  v_txtDivision <>"ALL" then
			    ' strsql4 = " and Division='" &  v_txtDivision & "'" 
		    ' else
			    ' strsql4 = " "
		    ' end if
		  
		  strha = strha + strsql2 +  strsql3 + strsql4	 & " group by FCBA, division,PLANTINGYEAR"
		  
	selectFields = " select GroupPlantation,sum(HA) HA, sum(bud_ha) bud_ha from ( " & strha  & " ) " & whereCons & " group by GroupPlantation "
		  strha  = selectFields 
	'response.write selectFields	  
Set rsha = Server.CreateObject("ADODB.Recordset")
'response.write strha
rsha.Open strha, connstr, adOpenStatic


if Not rsha.eof then
    luasHA = rsha("ha")  ' SEMENTARA AMBIL field budget  recomend Pak Dion  20 Nov 2013 rsha("ha")
	budHA = rsha("bud_ha")  
	' ubah luas bud ha menjadi rsha 18 nov 2013 ( recomended Pak dion , Pak Sahat ) rsha("bud_ha")
end if
rsha.close()

Dim prodtonBI,prodkgBI,budProdtonBI,budProdkgBI,prodtonSBI,prodkgSBI,budProdtonSBI,budProdkgSBI,budProdtonthn,budProdkgthn,TbudProdtonthn,TbudProdkgthn, TprodtonBI,TprodtonSBI,TbudProdtonBI,TbudProdtonSBI 

' prodtonBI=1
' prodkgBI=1
' budProdtonBI=1
' budProdkgBI=1
' prodtonSBI=1
' prodkgSBI=1
' budProdtonSBI=1
' budProdkgSBI=1		

if v_vStatus = "TM" then
				
				  if  v_vOwnership <>"ALL" then
					  strsql2 = "  and field_api.GetValid_ownership(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )  ='" & v_vOwnership & "'" 
				  else
					  strsql2= " "
				  end if
				  
				   ' if  v_txtPlantingDate <> "0"  then
					   ' strsql3 = " and to_char(f.PlantingDate,'YYYY') ='" & request.Form("vPlantingDate") & "'"
				   ' else
					   ' strsql3 = " ",
				   ' end if
				   ' if  v_txtDivision <> "ALL" then
					   ' strsql4 = " and f.Division='" & request.Form("vDivision") & "'" 
				   ' else
					   ' strsql4 = " "
				   ' end if
				 
		
        strprod = " select " & strFieldCons & " division,plantingyear, sum(prod_tonBI) prod_tonBI,sum(budProdBI) budProdBI,sum(prod_tonSBI) prod_tonSBI,sum(budProdSBI) budProdSBI,sum(budthn) budthn  from " & _
				  "	( select h.fcba,field_api.getValid_Afdeling(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) division,field_api.getValid_Plantingyear(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) plantingyear,field_api.GetValid_Field_Info_01(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )  field_info_01, round(nvl(sum(h.mill_weight),0),0) prod_tonBI,0 budProdBI,0  prod_tonSBI,0 budProdSBI,0 budthn from harvesting h , field f " & _
				  "	  where h.fieldcode = f.fccode and h.fcba = f.fcba and field_api.GetValid_status(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )  = '" & v_vStatus  & "' AND field_api.GetValid_CropType(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) =  '"& v_vCROPTYPE &"' " & strsql2 & strsql3 &  strsql4 & "  " & _
				  "	  and to_char(h.receptiondate,'YYYY')='" & v_txtyear & "' and to_char(h.receptiondate,'MM')=lpad('" & v_txtmonth & "',2,0) group by h.fcba,field_api.getValid_Afdeling(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ),field_api.getValid_Plantingyear(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ),field_api.GetValid_Field_Info_01(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) " & _
				  "	  UNION ALL " & _
				  "	  select h.fcba,field_api.getValid_Afdeling(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) division, field_api.getValid_Plantingyear(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) plantingyear,field_api.GetValid_Field_Info_01(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) field_info_01,0 prod_tonBI,0 budProdBI,round(nvl(sum(h.mill_weight),0),0)  prod_tonSBI,0 budProdSBI,0 budthn from harvesting h , field f " & _
				  "	  where h.fieldcode = f.fccode and h.fcba = f.fcba and field_api.GetValid_status(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) = '" & v_vStatus  & "' AND field_api.GetValid_CropType(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ) = '"& v_vCROPTYPE &"' " & strsql2 & strsql3 &  strsql4 & "  " & _
				  "	  and to_char(h.receptiondate,'YYYY')='" & v_txtyear & "' and to_char(h.receptiondate,'MM')<=lpad('" & v_txtmonth & "',2,0) group by h.fcba,field_api.getValid_Afdeling(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ),field_api.getValid_Plantingyear(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') ),field_api.GetValid_Field_Info_01(f.fccode,h.fcba,to_date( '01/" & v_txtmonth & "/" & v_txtyear & "','dd/mm/yyyy') )  " & _
				  "	  UNION ALL " & _
				  "	  select f.fcba,f.division,to_char(f.plantingdate,'YYYY') plantingyear,field_api.Get_Field_Info_01(f.FIELDCODE,f.fcba) Field_Info_01,0 prod_tonBI,sum( " & strbudprod_BI & ") budProdBI,0  prod_tonSBI,0 budProdSBI,0 budthn from BUDGET_TRANSACTION_PRODQTY f " & _
				  "	  where f.budget_year ='" & v_txtyear & "' "  & strsql3 &  strsql4 & "  " & _
				  "	  group by f.fcba,f.FIELDCODE,f.division,to_char(f.plantingdate,'YYYY') " & _
				  "	  UNION ALL " & _
				  "	  select f.fcba,f.division,to_char(f.plantingdate,'YYYY') plantingyear,field_api.Get_Field_Info_01(f.FIELDCODE,f.fcba) Field_Info_01,0 prod_tonBI,0 budProdBI,0  prod_tonSBI,sum(" & strbudprod_SBI & ") budProdSBI,0 budthn from BUDGET_TRANSACTION_PRODQTY f " & _
				  "	  where f.budget_year ='" & v_txtyear & "'  "  & strsql3 &  strsql4 & "  " & _
				  "	  group by f.fcba,f.FIELDCODE,f.division,to_char(f.plantingdate,'YYYY') " & _
				  "    UNION ALL " & _
				  "	  select f.fcba,f.division,to_char(f.plantingdate,'YYYY') plantingyear,field_api.Get_Field_Info_01(f.FIELDCODE,f.fcba) Field_Info_01,0 prod_tonBI,0 budProdBI,0  prod_tonSBI,0 budProdSBI,sum(f.qty_01 + f.qty_02 + f.qty_03 + f.qty_04 + f.qty_05 + f.qty_06 + f.qty_07 + f.qty_08 + f.qty_09 + f.qty_10 + f.qty_11 + f.qty_12 ) budthn from BUDGET_TRANSACTION_PRODQTY f " & _
				  "	  where f.budget_year ='" & v_txtyear & "' "  & strsql3 &  strsql4 & "  " & _
				  "	  group by f.fcba,f.FIELDCODE,f.division,to_char(f.plantingdate,'YYYY') " & _
				  "	) "
				 
				if v_txtASPCode <> "" and v_txtASPCode <> "NONE"  then
				     strprod = strprod +  " where Field_Info_01 = '" & v_txtASPCode & "' group by fcba,division,plantingyear	"
			    else
                     strprod = strprod +  " group by fcba,division,plantingyear	"
                end if				
				 
				strprod = " select GroupPlantation,sum(prod_tonBI) prod_tonBI,sum(budProdBI) budProdBI,sum(prod_tonSBI) prod_tonSBI,sum(budProdSBI) budProdSBI,sum(budthn) budthn  from ( " & strprod & " ) " & whereCons & " group by GroupPlantation "
		'response.write strprod
		Set rsha = Server.CreateObject("ADODB.Recordset")
		rsha.Open strprod, connstr, adOpenStatic
        
		if Not rsha.eof then
			prodtonBI=cdbl(rsha("prod_tonBI"))/1000
			prodkgBI=cdbl(rsha("prod_tonBI"))
			budProdtonBI=cdbl(rsha("budProdBI"))/1000
			budProdkgBI=cdbl(rsha("budProdBI"))
			prodtonSBI=cdbl(rsha("prod_tonSBI"))/1000
			prodkgSBI=cdbl(rsha("prod_tonSBI"))
			budProdtonSBI=cdbl(rsha("budProdSBI"))/1000
			budProdkgSBI=cdbl(rsha("budProdSBI"))
			budProdtonthn=cdbl(rsha("budthn"))/1000
			budProdkgthn=cdbl(rsha("budthn"))
		end if
		rsha.close()
end if

TprodtonBI=prodtonBI
TprodtonSBI=prodtonSBI 
TbudProdtonBI=budProdtonBI
TbudProdtonSBI=budProdtonSBI
TbudProdtonthn=budProdtonthn

if cdbl(prodtonBI)=0 then
   prodtonBI=1
end if
if cdbl(prodkgBI)=0 then
   prodkgBI=1
end if
if cdbl(budProdtonBI)=0 then
   budProdtonBI=1
end if
if cdbl(budProdkgBI)=0 then 
   budProdkgBI=1
end if
if cdbl(prodtonSBI)=0 then
	prodtonSBI=1
end if
if cdbl(prodkgSBI)=0 then
	prodkgSBI=1
end if
if budProdtonSBI=0 then
	budProdtonSBI=1
end if
if cdbl(budProdkgSBI)=0 then
	budProdkgSBI=1	
end if
if cdbl(budProdkgthn)=0 then
	budProdkgthn=1	
end if
%>	
 
	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
		<th style="background-color:#00A300;color:#fff;" colspan="5">BULAN INI</th>
		<th style="background-color:#00A300;color:#fff;" colspan="5">S/D BULAN INI</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2" rowspan="2">ANGGARAN TAHUNAN</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>">&nbsp;</th>
		
		<th style="background-color:#00A300;color:#fff;" colspan="2">REALISASI</th>
		
		<th style="background-color:#00A300;color:#fff;" colspan="2">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;" >%</th>
		
		<th style="background-color:#00A300;color:#fff;" colspan="2">REALISASI</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">ANGGARAN</th>
		<th style="background-color:#00A300;color:#fff;" >%</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">LUAS Ha</th>

		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(luasHA,2) %></th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(budHA,2) %></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ></th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(luasHA,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(budHA,2) %></th>
		<th style="background-color:#00A300;color:#fff;" align="right" ></th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(budHA,2) %></th>

	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">PRODUKSI - Ton</th>

		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(TprodtonBI,2) %> </th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(TbudProdtonBI,2) %></th>
		
		<% 
		' Prosentase 
		%>
		<% if cdbl(TbudProdtonBI) = 0 and cdbl(TprodtonBI) > 0 then %>
		    <th style="background-color:#00A300;color:#fff;" align="right" >100</th>
		<% elseif cdbl(TbudProdtonBI) > 0 and cdbl(TprodtonBI) = 0 then%>
		    <th style="background-color:#00A300;color:#fff;" align="right" >-100</th>
		<% elseif cdbl(TbudProdtonBI) = 0 and cdbl(TprodtonBI) = 0 then%>
		    <th style="background-color:#00A300;color:#fff;" align="right" >0</th>
		<% else %>	
		    <th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(TprodtonBI) - cdbl(TbudProdtonBI)) / cdbl(TbudProdtonBI)*100,2) %></th>
		<% End If %>
		
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(TprodtonSBI,2)%></th>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(TbudProdtonSBI,2) %></th>
		
		
		<% 
		' Prosentase SBI 
		%>
		<% if cdbl(TbudProdtonSBI) = 0 and cdbl(TprodtonSBI) > 0 then %>
		    <th style="background-color:#00A300;color:#fff;" align="right" >100</th>
		<% elseif cdbl(TbudProdtonSBI) > 0 and cdbl(TprodtonSBI) = 0 then%>
		    <th style="background-color:#00A300;color:#fff;" align="right" >-100</th>
		<% elseif cdbl(TbudProdtonSBI) = 0 and cdbl(TprodtonSBI) = 0 then%>
		    <th style="background-color:#00A300;color:#fff;" align="right" >0</th>
		<% else %>	
		    <th style="background-color:#00A300;color:#fff;" align="right" ><%=formatnumber((cdbl(TprodtonSBI) - cdbl(TbudProdtonSBI)) / cdbl(TbudProdtonSBI)*100,2) %></th>
		<% End If %>
		
		
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(TbudProdtonthn,2) %></th>
		

	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" colspan="<%= 2 + nSpanTotal %>" align="right">Yield - Ton/Ha </th>
        <% IF  cdbl(luasHA) > 0 then %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(cdbl(TprodtonBI)/cdbl(luasHA),2) %></th>
		<% else %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2">0</th>
		<% END IF %>
		<% IF  cdbl(budHA) > 0 then %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(cdbl(TbudProdtonBI)/cdbl(budHA),2) %></th>
		<% else %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2">0</th>
		<% END IF %>
		<th style="background-color:#00A300;color:#fff;" align="right" ></th>
		<% IF  cdbl(luasHA) > 0 then %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(cdbl(TprodtonSBI)/cdbl(luasHA),2) %></th>
		<% else %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2">0</th>
		<% END IF %>
		<% IF  cdbl(budHA) > 0 then %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(cdbl(TbudProdtonSBI)/cdbl(budHA),2) %></th>
		<% else %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2">0</th>
		<% END IF %>
		<th style="background-color:#00A300;color:#fff;" align="right" ></th>
		<% IF  cdbl(budHA) > 0 then %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2"><%=formatnumber(cdbl(TbudProdtonthn)/cdbl(budHA),2) %></th>
		<% else %>
		<th style="background-color:#00A300;color:#fff;" align="right" colspan="2">0</th>
		<% END IF %>

	</tr>
	<tr>
		<% 'Response.write nSpanTotal %>
		<th  style="background-color:#00A300;color:#fff;" >KODE</th>
		<% if   request.Form("vJenisReport") = "Group"  then %>
		    <th  style="background-color:#00A300;color:#fff;" >GROUP PEKERJAAN</th>
	    <% else %>
		    <th  style="background-color:#00A300;color:#fff;" >JENIS PEKERJAAN</th>
		<% End If %>
		
		<% if  v_txtCostElm ="Tampil" then %>
		      <th style="background-color:#00A300;color:#fff;">SOURCE TYPE</th>	
              <th style="background-color:#00A300;color:#fff;">SOURCE NAME</th>				  
	    <% end if %>
		
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th style="background-color:#00A300;color:#fff;">PLANTATION</th>			
	    <% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;">GROUP REGION</th>			
	    <% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th style="background-color:#00A300;color:#fff;">REGION</th>		
	    <% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">COMPANY</th>		
	    <% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th style="background-color:#00A300;color:#fff;">SITE</th>			
	    <% end if %>

		
		<% if v_txtDivision <> "--SUMMARY--"  and v_txtDivision <> "" then %>
		    <th style="background-color:#00A300;color:#fff;">AFDELING</th>	
	   <% end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> "" then %>
		    <th style="background-color:#00A300;color:#fff;">TT</th>			
	   <% end if %>
       

		
		
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Ha</th>
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Ha</th>
		<th style="background-color:#00A300;color:#fff;" >&nbsp;</th>
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Ha</th>
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Ha</th>
		<th style="background-color:#00A300;color:#fff;" >&nbsp;</th>
		<th style="background-color:#00A300;color:#fff;" >Total</th>
		<th style="background-color:#00A300;color:#fff;" >Rp/Ha</th>
	</tr>
</thead>	
<tbody>
	<%
	 Dim strdtl ,strsql2b, strgrp ,cdgrp ,noline , subAmtgrpBI,subAmtgrpSBI , GRAmtgrpBI,GRAmtgrpSBI, subbudgrpBI,subbudgrpSBI,subbudgrpthn , GRbudgrpBI,GRbudgrpSBI ,GRbudgrpthn  , grpDesc
     dim wherejob		
			' if   request.Form("vPlantingDate") <> "0"  then
				  ' strsql2 = " and substr(a.PlantingDate,-4) ='" & request.Form("vPlantingDate") & "'"
			 ' else
				  ' strsql2 = " "
			 ' end if
			 ' if   request.Form("vPlantingDate") <> "0"  then
				  ' strsql2b = " and to_char(a.PlantingDate,'YYYY') ='" & request.Form("vPlantingDate") & "'"
			 ' else
				  ' strsql2b = " "
			 ' end if
			 ' if  request.Form("vDivision") <>"ALL" then
				  ' strsql3 = " and a.division='" & request.Form("vDivision") & "'" 
			 ' else
				  ' strsql3 = " "
			 ' end if
			 if  v_txtyear <>"ALL" then
				  strsql4 = " and a.period_year='" & v_txtyear & "'" 
			 else
				  strsql4 = " "
			 end if
			 if  v_txtmonth <>"ALL" then
				  strsql5 = " and a.period_month='" & v_txtmonth & "'" 
			 else 
				  strsql5 = " "
			 end if
			if  v_vOwnership <>"ALL" then
				strsql6 = "  and a.ownership  ='" & v_vOwnership & "'" 
			else
				strsql6= " "
			end if
	        
			
			if v_txtASPCode <> "" and v_txtASPCode <> "NONE"  then
				     strsql7 = " and  Field_Info_01 = '" & v_txtASPCode & "' "
             end if	
			 
			 if v_txtjob <> "--SUMMARY--" and v_txtjob <> "" then
			      if v_txtjob <>  "--ALL--"  then
					wherejob = wherejob & " where code like '" & v_txtjob & "'"
			      end if 	  
			end if
			 
	 if   v_vJenisReport = "Group"  then
					  
		   ' strdtl =  " select trim(code) code,jobheading,trim(job_name) job_name,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn  from " & _
					 ' " ( select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name ,sum(a.tm_amt_actual) totalBI,0 BudBI,0 totalSBI,0 BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
					 ' "	  where a.period_year='" & v_txtyear & "' and a.period_month='" & v_txtmonth & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode " & _
					 ' "	  UNION " & _
					 ' "	  select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name ,0 totalBI,0 BudBI,sum(a.tm_amt_actual) totalSBI,0 BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
					 ' "	  where a.period_year='" & v_txtyear & "' and a.period_month <='" & v_txtmonth & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode " & _
					 ' "	  UNION " & _
					 ' "	  select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name ,0 totalBI,sum( " & strbud_BI & ")  BudBI  ,0 totalSBI,0 BudSBI,0 Budthn from view_budget_ar_tm_tbm a " & _
					 ' "	  where a.BUDGET_YEAR ='" & v_txtyear & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode " & _
					 ' "	  UNION " & _
					 ' "	  select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name ,0 totalBI,0  BudBI  ,0 totalSBI,sum(" & strbud_SBI & ") BudSBI,0 Budthn  from view_budget_ar_tm_tbm a " & _
					 ' "	  where a.BUDGET_YEAR ='" & v_txtyear & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode " & _
					 ' "	  UNION " & _
					 ' "	  select a.accountgroupcode code,a.activity_heading jobheading ,get_job_name(a.accountgroupcode) job_name ,0 totalBI,0  BudBI  ,0 totalSBI,0 BudSBI,sum(a.BUDGET_01+a.BUDGET_02+a.BUDGET_03+a.BUDGET_04+a.BUDGET_05+a.BUDGET_06+a.BUDGET_07+a.BUDGET_08+a.BUDGET_09+a.BUDGET_10+a.BUDGET_11+a.BUDGET_12) Budthn from view_budget_ar_tm_tbm a " & _
					 ' "	  where a.BUDGET_YEAR ='" & v_txtyear & "'  and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					 ' "	  GROUP BY a.activity_heading,a.accountgroupcode " & _
				     ' " ) group by code,jobheading,job_name order by code,jobheading,job_name " 
		     
			strdtl =  " select " & strFieldCons & " division,PLANTINGYEAR,trim(code) code,jobheading,jobgroup_api.get_jobgroup_name(code,fcba) job_name,sourcetype,sourcename,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn  from " & _
					  " ( select a.ba fcba,a.accountgroupcode code,a.activity_heading jobheading,a.division,to_char(to_date(PLANTINGDATE,'MM/DD/YYYY'),'YYYY') PLANTINGYEAR,a.sourcetype,a.sourcename, a.tm_amt_actual totalBI, " & _  
					  "   a.tm_amt_budget  BudBI,  " & _
                      "   0 totalSBI,  " & _
                      "   0 BudSBI,  " & _
                      "   a.ty_amt_budget Budthn  " & _
					  "   from sac_accountreport_tm_tbm a " & _
					  "   where a.period_year='" & v_txtyear & "' and a.period_month=lpad('" & v_txtmonth & "',2,0)and a.status='" & v_vStatus  &"' AND a.croptype = '"& v_vCROPTYPE &"'  " &   strsql6 & "  " &   strsql7 & "  " & _
					  "	  UNION All" & _
					  "	  select a.ba fcba,a.accountgroupcode code,a.activity_heading jobheading,a.division,to_char(to_date(PLANTINGDATE,'MM/DD/YYYY'),'YYYY') PLANTINGYEAR ,a.sourcetype,a.sourcename ,0 totalBI,0 BudBI,a.tm_amt_actual totalSBI,a.tm_amt_budget BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
					  "	  where a.period_year='" & v_txtyear & "' and a.period_month <=lpad('" & v_txtmonth & "',2,0)and a.status='" & v_vStatus  &"' AND a.croptype = '"& v_vCROPTYPE &"' " &  strsql6 & "  " &   strsql7 & "  " & _
					  "	)  " & wherejob & "  group by BUSINESSUNIT_API.get_OrderData(fcba),fcba,division,PLANTINGYEAR,code,jobheading,sourcetype,sourcename "
		
		  'response.write 
		  strdtl = "select GroupPlantation " & selectCons & ", code,jobheading,job_name,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn  From (" & strdtl & " )" & whereCons & " group by GroupPlantation,code,jobheading,job_name" & selectCons & "  Order by code,jobheading,job_name,GroupPlantation" & orderCons  	


          'strdtl = " select GroupPlantation,code,jobheading,job_name,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn  from ( " & strdtl & " ) " & whereCons & " group by GroupPlantation,code,jobheading,job_name order by code,jobheading"


					 
	 else  	
					
		   ' strdtl = " select trim(code) code,jobheading,trim(job_name) job_name,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn  from " & _
					' " ( select a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,sum(a.tm_amt_actual) totalBI,0 BudBI,0 totalSBI,0 BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
					' "  where a.period_year='" & v_txtyear & "' and a.period_month='" & v_txtmonth & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					' "  GROUP BY a.activity_heading,a.accountcode,a.Description " & _
					' "  UNION " & _
					' "  select a.accountcode code,a.activity_heading jobheading ,a.Description job_name ,0 totalBI,0 BudBI,sum(a.tm_amt_actual) totalSBI,0 BudSBI,0 Budthn from sac_accountreport_tm_tbm a " & _
					' "  where a.period_year='" & v_txtyear & "' and a.period_month <='" & v_txtmonth & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					' "  GROUP BY a.activity_heading,a.accountcode,a.Description " & _
					' "  UNION " & _
					' "  SELECT  a.accountcode code,a.activity_heading jobheading ,a.Description job_name,0 totalBI,sum(" & strbud_BI & ")  BudBI  ,0 totalSBI,0 BudSBI,0 Budthn from view_budget_ar_tm_tbm a " & _
					' "  where  a.BUDGET_YEAR = '" & v_txtyear & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					' "  GROUP BY a.activity_heading,a.accountcode,a.Description " & _
					' "  UNION " & _
					' "  SELECT  a.accountcode code,a.activity_heading jobheading ,a.Description job_name,0 totalBI,0  BudBI  ,0 totalSBI,sum(" & strbud_SBI & ") BudSBI,0 Budthn from view_budget_ar_tm_tbm a " & _
					' "  where  a.BUDGET_YEAR = '" & v_txtyear & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					' "  GROUP BY a.activity_heading,a.accountcode,a.Description " & _
		            ' "  UNION " & _
					' "  SELECT  a.accountcode code,a.activity_heading jobheading ,a.Description job_name,0 totalBI,0  BudBI  ,0 totalSBI,0 BudSBI,sum(a.BUDGET_01+a.BUDGET_02+a.BUDGET_03+a.BUDGET_04+a.BUDGET_05+a.BUDGET_06+a.BUDGET_07+a.BUDGET_08+a.BUDGET_09+a.BUDGET_10+a.BUDGET_11+a.BUDGET_12) Budthn  from view_budget_ar_tm_tbm a " & _
					' "  where  a.BUDGET_YEAR = '" & v_txtyear & "' and a.status='" & request.Form("vStatus") &"' AND a.croptype = '"& v_vCROPTYPE &"' and a.ba='" & session("fcba") & "' " & strsql2 & strsql3 &  strsql6 & "  " & _
					' "  GROUP BY a.activity_heading,a.accountcode,a.Description " & _
					' " ) group by code,jobheading,job_name order by code,jobheading,job_name "
					
		   strdtl = " select " & strFieldCons & " division,PLANTINGYEAR,trim(code) code,jobheading,job_api.get_jobcode_name(code,fcba) job_name,sourcetype,sourcename,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn   from  " & _
					"	( select a.ba fcba,a.accountcode code,a.activity_heading jobheading,a.division,to_char(to_date(PLANTINGDATE,'MM/DD/YYYY'),'YYYY') PLANTINGYEAR ,a.sourcetype,a.sourcename,a.tm_amt_actual totalBI, " & _
					"     a.tm_amt_budget  BudBI,  " & _
                    "     0 totalSBI,  " & _
                    "     0 BudSBI,  " & _
                    "     a.ty_amt_budget Budthn  " & _				
					"     from sac_accountreport_tm_tbm a " & _
					"	  where a.period_year='" & v_txtyear & "' and a.period_month=lpad('" & v_txtmonth & "',2,0)and a.status='" & v_vStatus &"' AND a.croptype = '"& v_vCROPTYPE &"'   " &  strsql6 & "  " &   strsql7 & "  " & _
					"	  UNION ALL " & _
					"	  select a.ba fcba,a.accountcode code,a.activity_heading jobheading,a.division,to_char(to_date(PLANTINGDATE,'MM/DD/YYYY'),'YYYY') PLANTINGYEAR,a.sourcetype,a.sourcename,0 totalBI,0 BudBI,a.tm_amt_actual totalSBI,a.tm_amt_budget BudSBI,0 Budthn  from sac_accountreport_tm_tbm a  " & _
					"	  where a.period_year='" & v_txtyear & "' and a.period_month <=lpad('" & v_txtmonth & "',2,0)and a.status='" & v_vStatus &"' AND a.croptype = '"& v_vCROPTYPE &"' "  &  strsql6 & "  " &   strsql7 & "  " & _
					"   )  " & wherejob & " group by BUSINESSUNIT_API.get_OrderData(fcba),fcba,division,PLANTINGYEAR,code,jobheading,sourcetype,sourcename"
					
				 strdtl =  "select GroupPlantation " & selectCons & " ,code,jobheading,job_name,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn  From (" & strdtl & " )" & whereCons & " group by GroupPlantation,code,jobheading,job_name" & selectCons & "   Order by code,jobheading,job_name,GroupPlantation" & orderCons  	
	
				' strdtl = " select GroupPlantation,code,jobheading,job_name,sum(totalBI) totalBI,sum(BudBI) BudBI,sum(totalSBI) totalSBI,sum(BudSBI) BudSBI,sum(Budthn) Budthn  from ( " & strdtl & " ) " & whereCons & " group by GroupPlantation,code,jobheading,job_name order by code,jobheading,job_name"

		   
	 end if					 
			 
	 		
			 
     Dim rp_ha_BI,rp_ha_budBI,rp_kg_BI,rp_kg_budBI, rp_ha_SBI,rp_ha_budSBI ,rp_kg_SBI,rp_kg_budSBI 
     Dim sub_rp_ha_BI,sub_rp_ha_budBI,sub_rp_kg_BI,sub_rp_kg_budBI, sub_rp_ha_SBI,sub_rp_ha_budSBI ,sub_rp_kg_SBI,sub_rp_kg_budSBI  
    	 
   ' response.write strdtl
	Set rsha = Server.CreateObject("ADODB.Recordset")
	rsha.Open strdtl, connstr, adOpenStatic
  ' response.write budHA
	   
	%>
	
	<%
	  if Not rsha.eof then
		  while Not rsha.eof
	             
				  if cdbl(luasHA) = 0  then 
				      rp_ha_BI = 0
					  rp_ha_SBI = 0
				   else				   
				      rp_ha_BI = cdbl(rsha("totalBI")) / cdbl(luasHA)
					  rp_ha_SBI = cdbl(rsha("totalSBI")) / cdbl(luasHA)
				   end if	
				    ' response.write budHA
                   if cdbl(budHA) = 0 then
                      rp_ha_budBI = 0
					  rp_ha_budSBI = 0
                   else 					  
				      rp_ha_budBI = cdbl(rsha("budBI")) / cdbl(budHA)
					  rp_ha_budSBI = cdbl(rsha("budSBI")) / cdbl(budHA)
				   end if
				   
				   if cdbl(prodkgBI) = 0 then
				      rp_kg_BI = 0 
				   else
				      rp_kg_BI = cdbl(rsha("totalBI")) / cdbl(prodkgBI)
				   end if
				   
				   if cdbl(budProdkgBI) = 0 then
				      rp_kg_budBI = 0
				   else
				      rp_kg_budBI = cdbl(rsha("budBI")) / cdbl(budProdkgBI)
				   end if
				   
				   if cdbl(prodkgSBI) = 0 then
				      rp_kg_SBI = 0 
				   else
				      rp_kg_SBI = cdbl(rsha("totalSBI")) / cdbl(prodkgSBI)
				   end if
				   
				   if cdbl(budProdkgSBI) = 0 then
				      rp_kg_budSBI = 0
				   else
				      rp_kg_budSBI = cdbl(rsha("budSBI")) / cdbl(budProdkgSBI)
				   end if	  
		    
	%> 

	   <tr>  
		
		
		    <% IF  cdgrp = rsha("jobheading")  then
			       noline = cint(noline )+ 1
				   subAmtgrpBI = cdbl(subAmtgrpBI) + cdbl(rsha("totalBI"))
				   subAmtgrpSBI = cdbl(subAmtgrpSBI) + cdbl(rsha("totalSBI"))
				   subbudgrpBI = cdbl(subbudgrpBI) + cdbl(rsha("budBI"))
				   subbudgrpSBI = cdbl(subbudgrpSBI) + cdbl(rsha("budSBI"))
				   subbudgrpthn  = cdbl(subbudgrpthn)  + cdbl(rsha("budthn"))
				  
			%>	   
				   
						   <%   tempjob = rsha("code") 
						        tempPlantation = "--SUMMARY--"
								tempGroupRegion = "--SUMMARY--"
								tempRegion =  "--ALL--"
								tempCompany ="--SUMMARY--"
								tempfcba = "--SUMMARY--"
								tempDivision = "--SUMMARY--"
								tempPlantingDate = "--SUMMARY--"
								 
						      If SprevCode <> rsha("code") then 
							       if v_vJenisReport = "Group"  then
                                      v_vJenisReport = "Job"
                                   end if 									  
								   %>
								 <td align="center">
								 <a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
								    <%=rsha("code")%>
								 </a>
								 </td>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>				   
							<% If Sprevjobname <> rsha("job_name") then %>
								 <td><%=rsha("job_name")%></td>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>	
						   <%
						   SprevCode   = rsha("code")
						   Sprevjobname = rsha("job_name")
						   %>
						    <% if  v_txtCostElm ="Tampil" then %>
								  <td><%=rsha("SOURCETYPE")%></td>
								  <td><%=rsha("SOURCENAME")%></td>			  
							<% end if %>
				           <% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
						             if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
									    tempjob = rsha("code")
									 else
									    tempjob = ""
                                     end if									 
									 tempPlantation = rsha("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevPlantation <> rsha("Plantation") then %>
										   <td align="center">																																																																																									
												<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
												<%=rsha("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsha("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
									 if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
									    tempjob = rsha("code")
									 else
									    tempjob = ""
                                     end if	
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
												<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
												<%=rsha("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsha("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
									   tempCompany = "--SUMMARY--"
									   tempfcba = "--ALL--"
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevRegion <> rsha("Region") then %>
											   <td align="center">
													<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsha("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsha("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
												    <a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													  <%=rsha("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsha("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
													<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
				            <% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> ""  then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
									   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
										   tempfcba = rsha("fcba")
									   else
										   tempfcba = ""
									   end if
									   tempDivision = rsha("Division")
									   tempPlantingDate = "--ALL--"
									   tempBlok = "--SUMMARY--"
										 IF SprevDivision <> rsha("Division") then %>
											   <td align="center">
													<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsha("Division")%>
													</a>
											   </td>
											   <% SprevDivision = rsha("Division") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
												
							<% if v_txtPlantingDate <> "--SUMMARY--" AND v_txtPlantingDate <> "" then %>
								<td><%=rsha("PlantingYear")%></td>			
						   <% end if %>


							
				   
				   <td align="right"><%=formatnumber(rsha("totalBI"),2)%></td>	
				 
				   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
							<td align="right"><%= formatnumber( cdbl(rp_kg_BI),2) %></td>
					<% else  %> 
							<td align="right"><%= formatnumber( cdbl(rp_ha_BI),2) %></td>
					<% end if %>	
				   <td align="right"><%=formatnumber(rsha("budBI"),2)%></td>
				   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
							<td align="right"><%= formatnumber( cdbl(rp_kg_budBI),2) %></td>
					<% else  %> 
							<td align="right"><%= formatnumber( cdbl(rp_ha_budBI),2) %></td>
					<% end if %>
					
					<% ' prosentase BI %>
					 <% IF ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then %>
							<% IF cdbl(rp_kg_budBI) = 0  and cdbl(rp_kg_BI) > 0 then %>
							   <td align="right">-100</td>
							<% Elseif cdbl(rp_kg_budBI) > 0  and cdbl(rp_kg_BI) = 0 then %>
							   <td align="right">100</td>
							<% Elseif cdbl(rp_kg_budBI) = 0  and cdbl(rp_kg_BI) = 0 then %>
							   <td align="right">0</td>  
							<% elseif cdbl(rp_kg_budBI) <> 0   and  cdbl(rp_kg_BI) <> 0   then %>   
					           <td align="right"><%= formatnumber((cdbl(rp_kg_budBI) - cdbl(rp_kg_BI))/cdbl(rp_kg_budBI)*100,,,-1) %></td>
							<% end if %>   
					<%  ELSE %>	
					        <% IF cdbl(rp_ha_budBI) = 0  and cdbl(rp_ha_BI) > 0 then %>
							   <td align="right">-100</td>
							<% Elseif cdbl(rp_ha_budBI) > 0  and cdbl(rp_ha_BI) = 0 then %>
							   <td align="right">100</td>
							<% Elseif cdbl(rp_ha_budBI) = 0  and cdbl(rp_ha_BI) = 0 then %>
							   <td align="right">0</td>  
							<% elseif cdbl(rp_ha_budBI) <> 0   and  cdbl(rp_ha_BI) <> 0   then %>   
						       <td align="right"><%= formatnumber((cdbl(rp_ha_budBI) - cdbl(rp_ha_BI))/cdbl(rp_ha_budBI)*100,,,-1) %></td>
					        <% end if %>
				    <% End If %>
					
					
					
				   <td align="right"><%=formatnumber(rsha("totalSBI"),2)%></td>
                    <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
							<td align="right"><%= formatnumber( cdbl(rp_kg_SBI),2) %></td>
					<% else  %> 
							<td align="right"><%= formatnumber( cdbl(rp_ha_SBI),2) %></td>
					<% end if %>	
					
				   <td align="right"><%=formatnumber(rsha("budSBI"),2)%></td>
				   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
							<td align="right"><%= formatnumber( cdbl(rp_kg_budSBI),2) %></td>
					<% else  %> 
							<td align="right"><%= formatnumber( cdbl(rp_ha_budSBI),2) %></td>
					<% end if %>
					
					<% ' prosentase SBI %>
					 <% IF ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then %>
							<% IF cdbl(rp_kg_budSBI) = 0  and cdbl(rp_kg_SBI) > 0 then %>
							   <td align="right">-100</td>
							<% Elseif cdbl(rp_kg_budSBI) > 0  and cdbl(rp_kg_SBI) = 0 then %>
							   <td align="right">100</td>
							<% Elseif cdbl(rp_kg_budSBI) = 0  and cdbl(rp_kg_SBI) = 0 then %>
							   <td align="right">0</td>  
							<% elseif cdbl(rp_kg_budSBI) <> 0   and  cdbl(rp_kg_SBI) <> 0   then %>   
					           <td align="right"><%= formatnumber((cdbl(rp_kg_budSBI) - cdbl(rp_kg_SBI))/cdbl(rp_kg_budSBI)*100,,,-1) %></td>
							<% end if %>   
					<%  ELSE %>	
					        <% IF cdbl(rp_ha_budSBI) = 0  and cdbl(rp_ha_SBI) > 0 then %>
							   <td align="right">-100</td>
							<% Elseif cdbl(rp_ha_budSBI) > 0  and cdbl(rp_ha_SBI) = 0 then %>
							   <td align="right">100</td>
							<% Elseif cdbl(rp_ha_budSBI) = 0  and cdbl(rp_ha_SBI) = 0 then %>
							   <td align="right">0</td>  
							<% elseif cdbl(rp_ha_budSBI) <> 0   and  cdbl(rp_ha_SBI) <> 0   then %>   
						       <td align="right"><%= formatnumber((cdbl(rp_ha_budSBI) - cdbl(rp_ha_SBI))/cdbl(rp_ha_budSBI)*100,,,-1) %></td>
					        <% end if %>
				    <% End If %>
					
					<td align="right"><%=formatnumber(rsha("Budthn"),2)%></td>
					<% IF cdbl(budHA) <> 0 THEN  %>
						<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
								  <td align="right"><%= formatnumber( cdbl(rsha("Budthn")) / cdbl(budProdkgthn),2) %></td>
						<% else  %> 
								  <td align="right"><%= formatnumber( cdbl(rsha("Budthn")) / cdbl(budHA),2) %></td>
						<% end if %>
					<% ELSE %>
						<td align="right">0</td>
                    <% END IF %>					
				<%   
			   ELSE %>
			      
					   <%  if  cdgrp <> "" then  %>
							 <tr>
								<% GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(subAmtgrpBI) 
								   GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(subAmtgrpSBI)
								   GRbudgrpBI = cdbl(GRbudgrpBI) + cdbl(subbudgrpBI) 
								   GRbudgrpSBI = cdbl(GRbudgrpSBI) + cdbl(subbudgrpSBI)
								   GRbudgrpthn = cdbl(GRbudgrpthn) + cdbl(subbudgrpthn)
								   
								%>
								
								<%  
								   If cdgrp = "PRM" Then
								      grpDesc = "Permulaan "
								   elseif 	 cdgrp = "PML" Then 
								      grpDesc = "Pemeliharaan "
								   elseif 	 cdgrp = "GC" Then 
								      grpDesc = "General Charge "	
  								   elseif	cdgrp = "PNN" Then 								   
								      grpDesc = "Panen "	
                                   end if									  
								%>
								
								<%
								   if cdbl(luasHA) = 0 then
								      sub_rp_ha_BI = 0
									  sub_rp_ha_SBI = 0
								   else
									   sub_rp_ha_BI = cdbl(subAmtgrpBI) / cdbl(luasHA)
									   sub_rp_ha_SBI = cdbl(subAmtgrpSBI) / cdbl(luasHA)
								   end if
								  
								   
								   if cdbl(budHA) = 0 then 
								       sub_rp_ha_budBI = 0
									   sub_rp_ha_budSBI = 0
								   else
									   sub_rp_ha_budBI = cdbl(subbudgrpBI) / cdbl(budHA)
									   sub_rp_ha_budSBI = cdbl(subbudgrpSBI) / cdbl(budHA)
								   end if
								   
								   If cdbl(prodkgBI) = 0 then
								      sub_rp_kg_BI = 0
                                   else									  
								      sub_rp_kg_BI = cdbl(subAmtgrpBI) / cdbl(prodkgBI)
								   end if	  
								   if cdbl(budprodkgBI) = 0 then
								      sub_rp_kg_budBI = 0
								   else
								      sub_rp_kg_budBI = cdbl(subbudgrpBI) / cdbl(budprodkgBI)
								   end if
								   
								   if cdbl(prodkgSBI ) = 0 then
								      sub_rp_kg_SBI  = 0
								   else
								      sub_rp_kg_SBI  = cdbl(subAmtgrpSBI) / cdbl(prodkgSBI )
								   end if
								   
								   if cdbl(budProdkgSBI)=0 then
								      sub_rp_kg_budSBI =  0
								   else
								      sub_rp_kg_budSBI =  cdbl(subbudgrpSBI) / cdbl(budProdkgSBI)
								   end if
								   'response.write budProdkgSBI  &  "-" & subbudgrpSBI&  "-" & sub_rp_kg_budSBI & " - " & sub_rp_ha_budSBI & " - " & rsha("jobheading") & " - " & cdgrp
								   
								%>
								
								<td style="background-color:#F4F4A0" colspan="<%= 2 + nSpanTotal %>"><b> Total&nbsp;<%=grpDesc%> </td>	
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpBI,2)%></td>
								<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
									  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(sub_rp_kg_BI),2) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(sub_rp_ha_BI),2) %></td>
								<% end if %>
								
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudgrpBI,2)%></td>
								<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
									  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(sub_rp_kg_budBI),2) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(sub_rp_ha_budBI),2) %></td>
								<% end if %>
								
								<% ' prosentase BI %>
								<%  IF  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" THEN %>
										<% IF cdbl(sub_rp_kg_budBI) = 0  and cdbl(sub_rp_kg_BI) > 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">-100</td>
										<% Elseif cdbl(sub_rp_kg_budBI) > 0  and cdbl(sub_rp_kg_BI) = 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">100</td>
										<% Elseif cdbl(sub_rp_kg_budBI) = 0  and cdbl(sub_rp_kg_BI) = 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">0</td>  
										<% elseif cdbl(sub_rp_kg_budBI) <> 0  and cdbl(sub_rp_kg_BI) <> 0 then %>   							
									       <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( (cdbl(sub_rp_kg_budBI) - cdbl(sub_rp_kg_BI))/cdbl(sub_rp_kg_budBI)*100,,,-1) %></td>
									    <% end if %>
								<% ELSE  %> 
								        <% IF cdbl(sub_rp_ha_budBI) = 0  and cdbl(sub_rp_ha_BI) > 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">-100</td>
										<% Elseif cdbl(sub_rp_ha_budBI) > 0  and cdbl(sub_rp_ha_BI) = 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">100</td>
										<% Elseif cdbl(sub_rp_ha_budBI) = 0  and cdbl(sub_rp_ha_BI) = 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">0</td>  
										<% elseif cdbl(sub_rp_ha_budBI) > 0  and cdbl(sub_rp_ha_BI) <> 0 then %>   																    
									       <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( (cdbl(sub_rp_ha_budBI) - cdbl(sub_rp_ha_BI))/cdbl(sub_rp_ha_budBI)*100,,,-1) %></td>
									    <% end if %>
								<% END IF %>
								
								
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subAmtgrpSBI,2)%></td>
								<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
									  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(sub_rp_kg_SBI),2) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(sub_rp_ha_SBI),2) %></td>
								<% end if %>
								
								<td  style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudgrpSBI,2)%></td>
								<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
									  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(sub_rp_kg_budSBI),2) %></td>
								<% else  %> 
									  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(sub_rp_ha_budSBI),2) %></td>
								<% end if %>
								
								<% ' prosentase SBI %>
								
								<%  IF  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" THEN %>
										<% IF cdbl(sub_rp_kg_budSBI) = 0  and cdbl(sub_rp_kg_BI) > 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">-100</td>
										<% Elseif cdbl(sub_rp_kg_budSBI) > 0  and cdbl(sub_rp_kg_BI) = 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">100</td>
										<% Elseif cdbl(sub_rp_kg_budSBI) = 0  and cdbl(sub_rp_kg_BI) = 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">0</td>  
										<% elseif cdbl(sub_rp_kg_budSBI) <> 0  and cdbl(sub_rp_kg_BI) <> 0 then %>   							
									       <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( (cdbl(sub_rp_kg_budSBI) - cdbl(sub_rp_kg_SBI))/cdbl(sub_rp_kg_budSBI)*100,,,-1) %></td>
									    <% end if %>
								<% ELSE  %> 
								        <% IF cdbl(sub_rp_ha_budSBI) = 0  and cdbl(sub_rp_ha_SBI) > 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">-100</td>
										<% Elseif cdbl(sub_rp_ha_budSBI) > 0  and cdbl(sub_rp_ha_SBI) = 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">100</td>
										<% Elseif cdbl(sub_rp_ha_budSBI) = 0  and cdbl(sub_rp_ha_SBI) = 0 then %>
										   <td  style="background-color:#F4F4A0" align="right">0</td>  
										<% elseif cdbl(sub_rp_ha_budSBI) > 0  and cdbl(sub_rp_ha_SBI) <> 0 then %>   																    
									       <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( (cdbl(sub_rp_ha_budSBI) - cdbl(sub_rp_ha_SBI))/cdbl(sub_rp_ha_budSBI)*100,,,-1) %></td>
									    <% end if %>
								<% END IF %>
								
								
									<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(subbudgrpthn,2)%></td>
								<% IF 	cdbl(budHA) <> 0 THEN %>
									 <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
										  <td style="background-color:#F4F4A0"  align="right"><b><%= formatnumber( cdbl(subbudgrpthn) / cdbl(budProdkgthn),2) %></td>
									<% else  %> 
										  <td  style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(subbudgrpthn) / cdbl(budHA),2) %></td>
									<% end if %>
								<% ELSE %>
									<td align="right">0</td>
								<% END IF %>	
							 </tr>
														
						<%  end if %>
				
						<%   noline = 1
							 subAmtgrpBI = rsha("totalBI")
							 subAmtgrpSBI = rsha("totalSBI")
							 subbudgrpBI = rsha("budBI")
				             subbudgrpSBI = rsha("budSBI")
							 subbudgrpthn  = cdbl(rsha("budthn"))
						%>
				        
						<%
								   if cdbl(luasHA) = 0 then
								      sub_rp_ha_BI = 0
									  sub_rp_ha_SBI = 0
								   else
									   sub_rp_ha_BI = cdbl(subAmtgrpBI) / cdbl(luasHA)
									   sub_rp_ha_SBI = cdbl(subAmtgrpSBI) / cdbl(luasHA)
								   end if
								  
								   
								   if cdbl(budHA) = 0 then 
								       sub_rp_ha_budBI = 0
									   sub_rp_ha_budSBI = 0
								   else
									   sub_rp_ha_budBI = cdbl(subbudgrpBI) / cdbl(budHA)
									   sub_rp_ha_budSBI = cdbl(subbudgrpSBI) / cdbl(budHA)
								   end if
								   
								   If cdbl(prodkgBI) = 0 then
								      sub_rp_kg_BI = 0
                                   else									  
								      sub_rp_kg_BI = cdbl(subAmtgrpBI) / cdbl(prodkgBI)
								   end if	  
								   if cdbl(budprodkgBI) = 0 then
								      sub_rp_kg_budBI = 0
								   else
								      sub_rp_kg_budBI = cdbl(subbudgrpBI) / cdbl(budprodkgBI)
								   end if
								   
								   if cdbl(prodkgSBI ) = 0 then
								      sub_rp_kg_SBI  = 0
								   else
								      sub_rp_kg_SBI  = cdbl(subAmtgrpSBI) / cdbl(prodkgSBI )
								   end if
								   
								   if cdbl(budProdkgSBI)=0 then
								      sub_rp_kg_budSBI =  0
								   else
								      sub_rp_kg_budSBI =  cdbl(subbudgrpSBI) / cdbl(budProdkgSBI)
								   end if
								   'response.write budProdkgSBI  &  "-" & subbudgrpSBI&  "-" & sub_rp_kg_budSBI & " - " & sub_rp_ha_budSBI & " - " & rsha("jobheading") & " - " & cdgrp
								   
								%>
						
					   <%  if  cdgrp = "PML" and v_vStatus = "TM" then  %>
							<tr>
								<td colspan="<%= 2 + nSpanTotal %>">&nbsp;</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td align="center">Rp/Kg</td>
							</tr>		
					   <%  end if %>
                       					   
					   <tr>
						     <%   tempjob = rsha("code") 
						        tempPlantation = "--SUMMARY--"
								tempGroupRegion = "--SUMMARY--"
								tempRegion = "--ALL--"
								tempCompany ="--SUMMARY--"
								tempfcba = "--SUMMARY--"
								tempDivision = "--SUMMARY--"
								tempPlantingDate = "--SUMMARY--"
						      If SprevCode <> rsha("code") then %>
								 <td align="center">
								 <a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
								    <%=rsha("code")%>
								 </a>
								 </td>
						   <% Else %>
								 <td align="center"></td>
						   <% End If %>				   
							<% If Sprevjobname <> rsha("job_name") then %>
								 <td><%=rsha("job_name")%></td>
						   <% Else %>
								 <td align="center"></td>
						   <% End If 
						   
						   SprevCode   = rsha("code")
						   Sprevjobname = rsha("job_name")
						   
						   %>
						    <% if  v_txtCostElm ="Tampil" then %>
								  <td><%=rsha("SOURCETYPE")%></td>
								  <td><%=rsha("SOURCENAME")%></td>			  
							<% end if %>
				           <% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
									 if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
									 tempPlantation = rsha("Plantation")
									 tempGroupRegion = "--ALL--"
									 tempRegion = "--SUMMARY--"
									 tempCompany ="--SUMMARY--"
									 tempfcba = "--SUMMARY--"
									 tempDivision = "--SUMMARY--"
									 tempPlantingDate = "--SUMMARY--"
									 IF SprevPlantation <> rsha("Plantation") then %>
										   <td align="center">
												<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
												<%=rsha("Plantation")%>
												</a>
										   </td>
										   <% SprevPlantation = rsha("Plantation") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then
									 if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
												<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
												<%=rsha("GroupRegion")%>
												</a>
										   </td>
										   <% SprevGroupRegion = rsha("GroupRegion") %>	
								  <% else	%>
										   <td align="center"></td>					
								   <% end if %>									   
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
									   tempCompany ="--SUMMARY--"
									   tempfcba = "--ALL--"
									   tempDivision = "--SUMMARY--"
									   tempPlantingDate = "--SUMMARY--"
									   tempBlok = "--SUMMARY--"
										 IF SprevRegion <> rsha("Region") then %>
											   <td align="center">
													<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsha("Region")%>
													</a>
											   </td>
											   <% SprevRegion = rsha("Region") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  ""  then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
												   <a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsha("Company")%>
													</a>
											   </td>
											   <% SprevCompany = rsha("Company") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  ""   then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
													<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsha("fcba")%>
													</a>
											   </td>
											   <% Sprevfcba = rsha("fcba") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
				            <% if v_txtDivision <> "--SUMMARY--" and v_txtDivision <> ""  then
									   if v_txtjob <> "--SUMMARY--" and  v_txtjob <> ""  then
											tempjob = rsha("code")
										 else
											tempjob = ""
										 end if	
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
									   if   v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then
										   tempfcba = rsha("fcba")
									   else
										   tempfcba = ""
									   end if
									   tempDivision = rsha("Division")
									   tempPlantingDate = "--ALL--"
									   tempBlok = "--SUMMARY--"
										 IF SprevDivision <> rsha("Division") then %>
											   <td align="center">
													<a href="p_AccountReportTMTBM.asp?txtjob=<%=tempjob%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&vstatus=<%=v_vstatus%>&vCROPTYPE=<%=v_vCROPTYPE%>&vJenisReport=<%=v_vJenisReport%>&vOwnership=<%=v_vOwnership%>">
													<%=rsha("Division")%>
													</a>
											   </td>
											   <% SprevDivision = rsha("Division") %>	
									  <% else	%>
											   <td align="center"></td>					
									   <% end if %>									   
							<% end if %>
							
							
												
							<% if v_txtPlantingDate <> "--SUMMARY--" AND v_txtPlantingDate <> "" then %>
								<td><%=rsha("Plantingyear")%></td>			
						   <% end if %>
							
							<td align="right"><%=formatnumber(rsha("totalBI"),2)%></td>	
                            <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
									<td align="right"><%= formatnumber( cdbl(rp_kg_BI),2) %></td>
							<% else  %> 
									<td align="right"><%= formatnumber( cdbl(rp_ha_BI),2) %></td>
							<% end if %>	
						   <td align="right"><%=formatnumber(rsha("budBI"),2)%></td>
						   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
									<td align="right"><%= formatnumber( cdbl(rp_kg_budBI),2) %></td>
							<% else  %> 
									<td align="right"><%= formatnumber( cdbl(rp_ha_budBI),2) %></td>
							<% end if %>
							
							<% ' prosentase BI %>
							 <% IF ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then %>
									<% IF cdbl(rp_kg_budBI) = 0  and cdbl(rp_kg_BI) > 0 then %>
									   <td align="right">-100</td>
									<% Elseif cdbl(rp_kg_budBI) > 0  and cdbl(rp_kg_BI) = 0 then %>
									   <td align="right">100</td>
									<% Elseif cdbl(rp_kg_budBI) = 0  and cdbl(rp_kg_BI) = 0 then %>
									   <td align="right">0</td>  
									<% elseif cdbl(rp_kg_budBI) <> 0   and  cdbl(rp_kg_BI) <> 0   then %>   
									   <td align="right"><%= formatnumber((cdbl(rp_kg_budBI) - cdbl(rp_kg_BI))/cdbl(rp_kg_budBI)*100,,,-1) %></td>
									<% end if %>   
							<%  ELSE %>	
									<% IF cdbl(rp_ha_budBI) = 0  and cdbl(rp_ha_BI) > 0 then %>
									   <td align="right">-100</td>
									<% Elseif cdbl(rp_ha_budBI) > 0  and cdbl(rp_ha_BI) = 0 then %>
									   <td align="right">100</td>
									<% Elseif cdbl(rp_ha_budBI) = 0  and cdbl(rp_ha_BI) = 0 then %>
									   <td align="right">0</td>  
									<% elseif cdbl(rp_ha_budBI) <> 0   and  cdbl(rp_ha_BI) <> 0   then %>   
									   <td align="right"><%= formatnumber((cdbl(rp_ha_budBI) - cdbl(rp_ha_BI))/cdbl(rp_ha_budBI)*100,,,-1) %></td>
									<% end if %>   
							<% End If %>

						   
							
							
						   <td align="right"><%=formatnumber(rsha("totalSBI"),2)%></td>
							<%  if  ( cdgrp = "PNN" or cdgrp = "GC" or rsha("jobheading")= "PNN" or rsha("jobheading")= "GC") and v_vStatus = "TM" then  %>
									<td align="right"><%= formatnumber( cdbl(rp_kg_SBI),2) %></td>
							<% else  %> 
									<td align="right"><%= formatnumber( cdbl(rp_ha_SBI),2) %></td>
							<% end if %>	
							
						   <td align="right"><%=formatnumber(rsha("budSBI"),2)%></td>
						   <%  if  ( cdgrp = "PNN" or cdgrp = "GC" or rsha("jobheading")= "PNN" or rsha("jobheading")= "GC") and v_vStatus = "TM" then  %>
									<td align="right"><%= formatnumber( cdbl(rp_kg_budSBI),2) %></td>
							<% else  %> 
									<td align="right"><%= formatnumber( cdbl(rp_ha_budSBI),2) %></td>
							<% end if %>
							
							<% ' prosentase SBI %>
							
								<%  IF  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" THEN %>
										<% IF cdbl(sub_rp_kg_budSBI) = 0  and cdbl(sub_rp_kg_SBI) > 0 then %>
										   <td  align="right">-100</td>
										<% Elseif cdbl(sub_rp_kg_budSBI) > 0  and cdbl(sub_rp_kg_SBI) = 0 then %>
										   <td   align="right">100</td>
										<% Elseif cdbl(sub_rp_kg_budSBI) = 0  and cdbl(sub_rp_kg_SBI) = 0 then %>
										   <td   align="right">0</td>  
										<% elseif cdbl(sub_rp_kg_budSBI) <> 0  and cdbl(sub_rp_kg_SBI) <> 0 then %>   							
									       <td   align="right"><%= formatnumber( (cdbl(sub_rp_kg_budSBI) - cdbl(sub_rp_kg_SBI))/cdbl(sub_rp_kg_budSBI)*100,,,-1) %></td>
									    <% end if %>
								<% ELSE  %> 
								        <% IF cdbl(sub_rp_ha_budSBI) = 0  and cdbl(sub_rp_ha_SBI) > 0 then %>
										   <td   align="right">-100</td>
										<% Elseif cdbl(sub_rp_ha_budSBI) > 0  and cdbl(sub_rp_ha_SBI) = 0 then %>
										   <td   align="right">100</td>
										<% Elseif cdbl(sub_rp_ha_budSBI) = 0  and cdbl(sub_rp_ha_SBI) = 0 then %>
										   <td   align="right">0</td>  
										<% elseif cdbl(sub_rp_ha_budSBI) > 0  and cdbl(sub_rp_ha_SBI) <> 0 then %>   																    
									       <td   align="right"><%= formatnumber( (cdbl(sub_rp_ha_budSBI) - cdbl(sub_rp_ha_SBI))/cdbl(sub_rp_ha_budSBI)*100,,,-1) %></td>
									    <% end if %>
								<% END IF %>
							
							
							
							
							
							<td align="right"><%=formatnumber(rsha("Budthn"),2)%></td>
							<% IF cdbl(budHA) <> 0 THEN %>
								<%  if  ( cdgrp = "PNN" or cdgrp = "GC" or rsha("jobheading")= "PNN" or rsha("jobheading")= "GC") and v_vStatus = "TM" then  %>
										  <td align="right"><%= formatnumber( cdbl(rsha("Budthn")) / cdbl(budProdkgthn),2) %></td>
								<% else  %> 
										  <td align="right"><%= formatnumber( cdbl(rsha("Budthn")) / cdbl(budHA),2) %></td>
								<% end if %>
						   <% ELSE %>	
						          <td align="right">0</td>
                           <% END IF%>						
					   </tr>
					   
					   
				   <%
			   END IF
			%>   
		   
		</tr>
	<%
	      
		  cdgrp = rsha("jobheading")
		rsha.MoveNext
		  wend
	  end if
	%>
	
	<tr>
	     <% GRAmtgrpBI = cdbl(GRAmtgrpBI) + cdbl(subAmtgrpBI) 
		    GRAmtgrpSBI = cdbl(GRAmtgrpSBI) + cdbl(subAmtgrpSBI) 
			GRbudgrpBI = cdbl(GRbudgrpBI) + cdbl(subbudgrpBI) 
			GRbudgrpSBI = cdbl(GRbudgrpSBI) + cdbl(subbudgrpSBI)
			GRbudgrpthn = cdbl(GRbudgrpthn) + cdbl(subbudgrpthn) 
			
			Dim grandtot_rp_ha_BI,grandtot_rp_ha_budBI,grandtot_rp_kg_BI,grandtot_rp_kg_budBI, grandtot_rp_ha_SBI,grandtot_rp_ha_budSBI ,grandtot_rp_kg_SBI,grandtot_rp_kg_budSBI 
			                       
								   if cdbl(luasHA) = 0 then
								      grandtot_rp_ha_BI = 0
									  grandtot_rp_ha_SBI = 0
								   else
									   grandtot_rp_ha_BI = cdbl(GRAmtgrpBI) / cdbl(luasHA)
									   grandtot_rp_ha_SBI = cdbl(GRAmtgrpSBI) / cdbl(luasHA)
								   end if
								  
								   
								   if cdbl(budHA) = 0 then 
								       grandtot_rp_ha_budBI = 0
									   grandtot_rp_ha_budSBI = 0
								   else
									   grandtot_rp_ha_budBI = cdbl(GRbudgrpBI) / cdbl(budHA)
									   grandtot_rp_ha_budSBI = cdbl(GRbudgrpSBI) / cdbl(budHA)
								   end if
								   
								   If cdbl(prodkgBI) = 0 then
								      grandtot_rp_kg_BI = 0
                                   else									  
								      grandtot_rp_kg_BI = cdbl(GRAmtgrpBI) / cdbl(prodkgBI)
								   end if	  
								   if cdbl(budprodkgBI) = 0 then
								      grandtot_rp_kg_budBI = 0
								   else
								      grandtot_rp_kg_budBI = cdbl(GRbudgrpBI) / cdbl(budprodkgBI)
								   end if
								   
								   if cdbl(prodkgSBI ) = 0 then
								      grandtot_rp_kg_SBI  = 0
								   else
								      grandtot_rp_kg_SBI  = cdbl(GRAmtgrpSBI) / cdbl(prodkgSBI )
								   end if
								   
								   if cdbl(budProdkgSBI)=0 then
								      grandtot_rp_kg_budSBI =  0
								   else
								      grandtot_rp_kg_budSBI =  cdbl(GRbudgrpSBI) / cdbl(budProdkgSBI)
								   end if
		 %>
		<td style="background-color:#F4F4A0" colspan="<%= 2 + nSpanTotal %>"><b> Grand Total&nbsp;</td>	
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpBI,2)%></td>
		<%  if   v_vStatus = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_kg_BI),2) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_ha_BI),2) %></b></td>
		<% end if %>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudgrpBI,2)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_kg_budBI),2) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_ha_budBI),2) %></b></td>
		<% end if %>
		
		
		<% ' prosentase BI %>
		<% IF cdbl(grandtot_rp_kg_budBI) = 0  and cdbl(grandtot_rp_kg_BI) > 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>-100</td>
		<% Elseif cdbl(grandtot_rp_kg_budBI) > 0  and cdbl(grandtot_rp_kg_BI) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>100</td>
		<% Elseif cdbl(grandtot_rp_kg_budBI) = 0  and cdbl(grandtot_rp_kg_BI) < 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>100</td>
		<% Elseif cdbl(grandtot_rp_kg_budBI) = 0  and cdbl(grandtot_rp_kg_BI) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>0</td>  
		<% else %>   
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( (cdbl(grandtot_rp_kg_budBI) - cdbl(grandtot_rp_kg_BI) )/cdbl(grandtot_rp_kg_budBI)*100,,,-1) %></td>
		<% End If %>
		
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRAmtgrpSBI,2)%></td>
		<%  if   v_vStatus = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_kg_SBI),2) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_ha_SBI),2) %></b></td>
		<% end if %>
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudgrpSBI,2)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_kg_budSBI),2) %></b></td>
		<% else  %> 
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(grandtot_rp_ha_budSBI),2) %></b></td>
		<% end if %>
		
		
		<% ' prosentase SBI %>
		<% IF cdbl(grandtot_rp_kg_budSBI) = 0  and cdbl(grandtot_rp_kg_SBI) > 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>-100</td>
		<% Elseif cdbl(grandtot_rp_kg_budSBI) > 0  and cdbl(grandtot_rp_kg_SBI) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>100</td>
		<% Elseif cdbl(grandtot_rp_kg_budSBI) = 0  and cdbl(grandtot_rp_kg_SBI) = 0 then %>
				<td style="background-color:#F4F4A0" align="right"><b>0</td>  
		<% else %>   
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( (cdbl(grandtot_rp_kg_budSBI) - cdbl(grandtot_rp_kg_SBI) )/cdbl(grandtot_rp_kg_budSBI)*100,,,-1) %></td>
		<% End If %>
		
		<td style="background-color:#F4F4A0" align="right"><b><%=formatnumber(GRbudgrpthn,2)%></td>
		<%  if  ( cdgrp = "PNN" or cdgrp = "GC" ) and v_vStatus = "TM" then  %>
				<td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(GRbudgrpthn) / cdbl(budProdkgthn),2) %></b></td>
		<% else  %> 
		       <% IF  cdbl(luasHA) <> 0 then %>
			   <td style="background-color:#F4F4A0" align="right"><b><%= formatnumber( cdbl(GRbudgrpthn) / cdbl(luasHA),2) %></b></td>
			   <% Else %>
			   <td style="background-color:#F4F4A0" align="right"><b>0</b></td>
			   <% End If %>
		<% end if %>			
	<tr>
 </tbody>	
</table>
<!--#include file="bottom.Asp"-->
