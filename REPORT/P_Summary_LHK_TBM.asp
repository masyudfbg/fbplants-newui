<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%	
	Dim rsEPS, strEPS,rscomp , strcomp,	strEPS2
	Dim v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass
	Dim strMain, nSpanTotal
	Dim txtyear, txtmonth, txttodaydate, txtDivision, txtBlok, txtPlantingDate,txtComp,v_txtComp
    Dim tempDivision, tempBlok, tempPlantingDate 
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	Dim nSubMandays_Amount_hi, nSubMandays_Amount_shi
	Dim nOutput_Mandays_Amount_hi, nOutput_Mandays_Amount_shi
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	
	        v_txtGroupPlantation  = Request("txtGroupPlantation")
			v_txtPlantation  	= Request("txtPlantation")
			v_txtGroupRegion  	= Request("txtGroupRegion")
			v_txtRegion  		= Request("txtRegion")
			v_txtcompany  		= Request("txtCompany")
			v_txtfcba 			= Request("txtfcba")
			v_txtDivision		= Request("txtDivision")
			v_txtBlok			= Request("txtBlok")
			v_txtPlantingDate	= Request("txtPlantingDate")
			v_txttodaydate		= Request("txttodaydate")
			
			txtyear 		= right(Request("txttodaydate"),4)
			txtmonth 		= left(Request("txttodaydate"),2)
			txttodaydate 	= Request("txttodaydate")
			
			txtComp         = Request("txtCompany")
			nSpanTotal = 8
			v_txtComp           =   txtComp
			v_txtyear			=	txtyear
			v_txtmonth			=	txtmonth
			
			
			
			'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba  &  " D " & v_txtDivision & " B "  & v_txtBlok & " TT" & v_txtPlantingDate
			
			Dim selectFields,strFieldCons,selectCons,whereCons,GroupCons,orderCons 
			strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company, fcba, BUSINESSUNIT_API.get_OrderData(fcba) norut, " 

						   
			wherecons = " where company is not null "
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
			    selectCons = selectCons & ",fcba"
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
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
			
			
			
			   if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
			      if v_txtDivision <>  "--ALL--"  then
					whereCons = whereCons & " AND DIVISION = '" & v_txtDivision & "'"
			      end if
                  selectCons = selectCons &  " , DIVISION "
				  orderCons  = orderCons & ",DIVISION "
			   else
			    nSpanTotal = nSpanTotal - 1 	  
			   end if   

            
			
				if v_txtPlantingDate <> "--SUMMARY--" and v_txtPlantingDate <> ""   then
				   if v_txtPlantingDate <>  "--ALL--" then
					  whereCons = whereCons & " AND PLANTINGYEAR = '" & v_txtplantingdate & "'"
				   end if
				   selectCons = selectCons &  " , PLANTINGYEAR "
				   orderCons  = orderCons & ",PLANTINGYEAR "
				else
			    nSpanTotal = nSpanTotal - 1    
				end if
		
			
				if v_txtBlok <> "--SUMMARY--" and v_txtBlok <> "" then
				   if v_txtBlok <>  "--ALL--" then
					  whereCons = whereCons & " AND FIELDCODE = '" & v_txtBlok & "'"
				   end if
				   selectCons = selectCons & " , FIELDCODE "
				   orderCons  = orderCons & ",FIELDCODE "
				else
			    nSpanTotal = nSpanTotal - 1    
				end if
            
			
	strMain = " SELECT 	groupplantationcode groupplantation,plantationcode plantation,groupregioncode groupregion,regioncode region,companycode company,fcba,  " & _
              "        	todate  Act_Stack_hi,  " & _
              "        	tomonth Act_Stack_shi,  " & _
              " 		budgettodate budget_Stack_hi,  " & _
              " 		budgettomonth budget_Stack_shi,  " & _
              " 		toyear Act_Stack_sbi,  " & _
              " 		budgettoyear budget_Stack_sbi,  " & _
              " 		0 Act_Tnm_hi,  " & _
              " 		0 Act_Tnm_shi,  " & _
              " 		0 budget_Tnm_hi,  " & _
              " 		0 budget_Tnm_shi,  " & _
              " 		0 Act_Tnm_sbi,  " & _
              " 		0 budget_Tnm_sbi  " & _
              " from production_summary_tab  " & _
			  " where to_char(ondate, 'YYYYMMDD') = to_char(TO_DATE('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYYMMDD')  " & _
              " and stepnumber = 'STEP_01' and  jobgencode = 'STACKING' " & _
			  " Union " & _
			  " SELECT groupplantationcode groupplantation,plantationcode plantation,groupregioncode groupregion,regioncode region,companycode company,fcba,  " & _
              " 		0  Act_Stack_hi,  " & _
              " 		0 Act_Stack_shi,  " & _
              " 		0 budget_Stack_hi,  " & _
              " 		0 budget_Stack_shi,  " & _
			   " 		0 Act_Stack_sbi,  " & _
              " 		0 budget_Stack_sbi,  " & _
              " 		todate  Act_Tnm_hi,  " & _
              " 		tomonth Act_Tnm_shi,  " & _
              " 		budgettodate budget_Tnm_hi,  " & _
              " 		budgettomonth budget_Tnm_shi,  " & _
              " 		toyear Act_Tnm_sbi,  " & _
              " 		budgettoyear budget_Tnm_sbi  " & _
              " from production_summary_tab  " & _
              " where to_char(ondate, 'YYYYMMDD') = to_char(TO_DATE('" & v_txttodaydate & "', 'MM/DD/YYYY'), 'YYYYMMDD')  " & _
              " and stepnumber = 'STEP_01' and  jobgencode = 'TANAMSAWIT' "
             		
	 strEPS2  =  " ,sum(Act_Stack_hi)  Act_Stack_hi,   " & _
				 " sum(Act_Stack_shi) Act_Stack_shi,   " & _
                 " sum(budget_Stack_hi) budget_Stack_hi,   " & _
				 " sum(budget_Stack_shi) budget_Stack_shi,   " & _
				 " sum(Act_Stack_shi) - sum(budget_Stack_shi) var_stack_shi,   " & _
                 " decode(sum(budget_Stack_shi),0,0,(sum(Act_Stack_shi) - sum(budget_Stack_shi)) / sum(budget_Stack_shi) * 100) varPer_stack_shi,   " & _
                 " sum(Act_Stack_sbi) Act_Stack_sbi,   " & _
                 " sum(budget_Stack_sbi) budget_Stack_sbi,   " & _
				 " sum(Act_Stack_sbi) - sum(budget_Stack_sbi) var_stack_sbi,   " & _
                 " decode(sum(budget_Stack_sbi),0,0,sum(Act_Stack_sbi) - sum(budget_Stack_sbi)/ sum(budget_Stack_sbi) * 100) varPer_stack_sbi,  " & _
				 " sum(Act_Tnm_hi)  Act_Tnm_hi,   " & _
				 " sum(Act_Tnm_shi) Act_Tnm_shi,   " & _
                 " sum(budget_Tnm_hi) budget_Tnm_hi,   " & _
                 " sum(budget_Tnm_shi) budget_Tnm_shi,   " & _
                 " sum(Act_Tnm_shi) - sum(budget_Tnm_shi) var_Tnm_shi,   " & _
                 " decode(sum(budget_Tnm_shi),0,0,sum(Act_Tnm_shi) - sum(budget_Tnm_shi)/ sum(budget_Tnm_shi)*100) varPer_Tnm_shi,   " & _
                 " sum(Act_Tnm_sbi) Act_Tnm_sbi,   " & _
                 " sum(budget_Tnm_sbi) budget_Tnm_sbi,   " & _
                 " sum(Act_Tnm_sbi) - sum(budget_Tnm_sbi)   var_tnm_sbi,   " & _
                 " decode(sum(budget_Tnm_sbi),0,0,sum(Act_Tnm_sbi) - sum(budget_Tnm_sbi)/sum(budget_Tnm_sbi) * 100) varPer_tnm_sbi   "  
								 
								 

	strEPS =  "select GroupPlantation " & selectCons & strEPS2 &"  From (" & strMain & " )" & whereCons & " and USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'E' group by  GroupPlantation" & selectCons & "  Order by GroupPlantation" & orderCons  	

	

				
Set rsEPS = Server.CreateObject("ADODB.Recordset")
'response.write strEPS

rsEPS.Open strEPS , connstr, adOpenStatic 


%>

<html>

<head>
<title>ePlaS Reporting - Laporan Harian Produksi</title>
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
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
<form name="frmEntry" method="post" action="P_Summary_LHK_TBM.asp">


	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	

<!--#include file="FormSelectNewSum.html"-->
<table border="0" id="form-table" class="centered-margin">
    
	<tr>
		<th>TAHUN</th>
		<th>BULAN</th>
		<th>TANGGAL</th>
		
		<th>&nbsp;</th>
	</tr>
	<input type="hidden" name="txtComp" value="<%=v_txtComp%>">
	<tr>
	    <!--%response.write v_txtfcba %-->
	    <!--%response.write v_txtyear %-->
	    <!--%response.write v_txtmonth %-->
		<!--%response.write len(trim(request("txtYear"))) %-->
		<td align="center"><input type="text" name="txtyear" value="<%=v_txtyear%>"></td>
		<td align="center"><input type="text" name="txtmonth" value="<%=v_txtmonth%>"></td>
		<td align="center">
		<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
			<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
		<% else%>
			<input class="datepicker" type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<!--<a class="calendar" href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>-->
		</td>
		
		
		
		
		<td align="center"><input type="submit" value="View" /></td>

	</tr>
</table>



<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
<tr>
<th colspan="38" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
		<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
		</div>
		<div id="subheader"class="centered">
			<h2>LAPORAN HARIAN KEBUN (TBM) </h2>
			<% if v_txttodaydate <> "" and v_txtFcba <> "" then %>
				<h5>PERIODE : <%=v_txttodaydate%>, KEBUN : <%=v_txtFcba%></h5>
			<% end if %>
			<h3></br></h3>
			<h3></br></h3>
		</div>
	</div>

	<%
		if len(trim(request("txtYear"))) = 0 then
			v_txtyear = year(Now)	 
		end if

		if len(trim(request("txtmonth"))) = 0 then
			v_txtmonth = Month(Now)	 
		end if
	%>

</th>
</tr>
	<tr>
		 <th rowspan="4" style="  background-color:#00A300;color:#fff; "></th>
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% 			
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% 			
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% 		
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="4" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% 		
	    end if %>

		
		<% if v_txtDivision <> "--SUMMARY--"   and v_txtDivision <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% 		
	    end if %>
		
		
		<% if v_txtPlantingDate <> "--SUMMARY--" And v_txtPlantingDate <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">TT</th>
		<% 		
	    end if %>


		<% if v_txtBlok <> "--SUMMARY--" And v_txtBlok <> "" then %>
		    <th rowspan="4" style="background-color:#00A300;color:#fff;">BLOK</th>
		<% 		
	    end if %>
		
		<th colspan="10">STACKING ( HA )</th>
		<th colspan="10">TANAM (HA)</th>

	</tr>	
	<tr>
		<th colspan="6">BULAN INI (BI)</th>
		<th colspan="4">SAMPAI BULAN INI (SBI)</th>
		<th colspan="6">BULAN INI (BI)</th>
		<th colspan="4">SAMPAI DGN BULAN INI (S/D BI)</th>
	</tr>	
	<tr>
		<th colspan="2">AKTUAL</th>
		<th colspan="2">BUDGET</th>
		<th colspan="2">VAR SHI</th>
		<th rowspan="2">AKTUAL</th>
		<th rowspan="2">BUDGET</th>
		<th colspan="2">VAR SBI</th>
		<th colspan="2">AKTUAL</th>
		<th colspan="2">BUDGET</th>
		<th colspan="2">VAR SHI</th>
		<th rowspan="2">AKTUAL</th>
		<th rowspan="2">BUDGET</th>
		<th colspan="2">VAR SBI</th>
	</tr>	
	<tr>
		<th>HI</th>
		<th>SHI</th>
		<th>HI</th>
		<th>SHI</th>
		<th>Ha</th>
		<th>%</th>
		<th>Ha</th>
		<th>%</th>
		<th>HI</th>
		<th>SHI</th>
		<th>HI</th>
		<th>SHI</th>
		<th>Ha</th>
		<th>%</th>
		<th>Ha</th>
		<th>%</th>	
	</tr>
        </thead>
 <tbody>
<%

Dim Gtot_Act_Tnm_hi,Gtot_Act_Tnm_Shi ,Gtot_budget_Tnm_hi,Gtot_budget_Tnm_Shi,Gtot_Act_Tnm_sbi ,Gtot_budget_Tnm_sbi
Dim Gtot_Act_Stack_hi ,Gtot_Act_Stack_Shi ,Gtot_budget_Stack_hi ,Gtot_budget_Stack_Shi ,Gtot_Act_Stack_sbi,Gtot_budget_Stack_sbi
Dim TotCmp_Act_Tnm_hi ,TotCmp_Act_Tnm_Shi ,TotCmp_budget_Tnm_hi ,TotCmp_budget_Tnm_Shi ,TotCmp_Act_Tnm_sbi ,TotCmp_budget_Tnm_sbi 				
Dim	TotCmp_Act_Stack_hi ,TotCmp_Act_Stack_Shi ,TotCmp_budget_Stack_hi ,TotCmp_budget_Stack_Shi ,TotCmp_Act_Stack_sbi,TotCmp_budget_Stack_sbi

Dim gtotvar_tnm_shi, gtotvar_tnm_sbi,gtotvarper_tnm_shi, gtotvarper_tnm_sbi
Dim gtotvar_Stack_shi,gtotvar_Stack_sbi,gtotvarper_Stack_shi,gtotvarper_Stack_sbi
Dim TotCmpvar_tnm_shi, TotCmpvar_tnm_sbi,TotCmpvarper_tnm_shi, TotCmpvarper_tnm_sbi
Dim TotCmpvar_Stack_shi,TotCmpvar_Stack_sbi,TotCmpvarper_Stack_shi,TotCmpvarper_Stack_sbi

Dim nsubhi, nsubshi, nsubbi, nsubsbi, nsubtonase_hi
Dim nsubtonase_shi, nsubtonase_bi, nsubtonase_sbi, nsubpom_hi, nsubpom_shi
Dim nsubpom_bi, nsubpom_sbi, nsubpom_tonase_hi, nsubpom_tonase_shi, nsubpom_tonase_bi, nsubpom_tonase_sbi, nsubrestan
Dim nsubbudjjg_shi,nsubbudjjg_sbi,nsubbudqty_shi,nsubbudqty_sbi,nsubbudmandays_shi,nsubbudmandays_sbi,nsubbudOutput_shi,nsubbudOutput_sbi
Dim nsubmandays_hi, nsubmandays_shi, nsubmandays_bi, nsubmandays_sbi
Dim SprevFcba, sprevDivision, sprevPlantingYear, sprevFieldCode
Dim SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation,GantiCompany
Dim nsubOutput_hi, nsubOutput_shi, nsubOutput_bi, nsubOutput_sbi
Dim nOutput_hi, nOutput_shi, nOutput_bi, nOutput_sbi ,nbudOutput_shi,nbudOutput_sbi
Dim nsubwo_plus_hi, nsubwo_plus_shi, nsubwo_plus_sbi,nsubwo_min_hi, nsubwo_min_shi, nsubwo_min_sbi


	if not rsEPS.EOF then
		NO = 1
		Gtot_Act_Tnm_hi = 0
		Gtot_Act_Tnm_Shi = 0
		GantiCompany = "FALSE"
		SprevPlantation= ""
		SprevGroupRegion= ""
		SprevRegion= ""
		SprevCompany= ""
		sprevFcba = ""
		sprevDivision = ""
		sprevPlantingYear = ""
		sprevFieldCode = ""
		while not rsEPS.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
	
					    if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then 
						    if SprevCompany <> rsEPS("Company") then
							  if SprevCompany <> "" then
							       TotCmpvar_tnm_shi = TotCmp_Act_Tnm_Shi - TotCmp_budget_Tnm_Shi
								   if TotCmp_budget_Tnm_Shi = 0 then
									  TotCmpvarper_tnm_shi = 0
								   else 
									  TotCmpvarper_tnm_shi = (TotCmp_Act_Tnm_Shi - TotCmp_budget_Tnm_Shi) / TotCmp_budget_Tnm_Shi * 100
								   end if	  
								   
								   TotCmpvar_tnm_sbi = TotCmp_Act_Tnm_Sbi - TotCmp_budget_Tnm_Sbi
								   if TotCmp_budget_Tnm_Sbi = 0 then
									  TotCmpvarper_tnm_sbi = 0
								   else 
									  TotCmpvarper_tnm_sbi = (TotCmp_Act_Tnm_Sbi - TotCmp_budget_Tnm_Sbi) / TotCmp_budget_Tnm_Sbi * 100
								   end if
								   
								   TotCmpvar_Stack_shi = TotCmp_Act_Stack_Shi - TotCmp_budget_Stack_Shi
									if TotCmp_budget_Stack_Shi = 0 then
									  TotCmpvarper_Stack_shi = 0
								   else 
									  TotCmpvarper_Stack_shi = (TotCmp_Act_Stack_Shi - TotCmp_budget_Stack_Shi) / TotCmp_budget_Stack_Shi * 100
								   end if	  
								   
								   TotCmpvar_Stack_sbi = TotCmp_Act_Stack_Sbi - TotCmp_budget_Stack_Sbi
								   if TotCmp_budget_Stack_Sbi = 0 then
									  TotCmpvarper_Stack_sbi = 0
								   else 
									  TotCmpvarper_Stack_sbi = (TotCmp_Act_Stack_Sbi - TotCmp_budget_Stack_Sbi) / TotCmp_budget_Stack_Sbi * 100
			                       end if 
					%>
								<tr bgcolor="#F4F4A0">
									<td align="center" colspan="<%= 1+ nSpanTotal %>">TOTAL&nbsp;&nbsp<%=tempCompany%></td>				
									<td align="right"><%=FormatNumber(TotCmp_Act_Stack_hi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_Act_Stack_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Stack_hi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Stack_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvar_stack_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvarper_Stack_shi,0)%></td>
									
									<td align="right"><%=FormatNumber(TotCmp_act_Stack_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Stack_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvar_stack_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvarper_Stack_sbi,0)%></td>
											
									<td align="right"><%=FormatNumber(TotCmp_Act_Tnm_hi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_Act_Tnm_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Tnm_hi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Tnm_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvar_tnm_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvarper_tnm_shi,0)%></td>
									
									<td align="right"><%=FormatNumber(TotCmp_act_Tnm_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Tnm_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvar_Tnm_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvarper_Tnm_sbi,0)%></td>
								</tr>
                                								
<%						     
                                TotCmp_Act_Tnm_hi = 0
								TotCmp_Act_Tnm_Shi = 0
								TotCmp_budget_Tnm_hi = 0
								TotCmp_budget_Tnm_Shi = 0
								TotCmp_Act_Tnm_sbi = 0
								TotCmp_budget_Tnm_sbi = 0
								
								TotCmp_Act_Stack_hi = 0
								TotCmp_Act_Stack_Shi = 0
								TotCmp_budget_Stack_hi = 0
								TotCmp_budget_Stack_Shi = 0
								TotCmp_Act_Stack_sbi = 0
								TotCmp_budget_Stack_sbi = 0
							end if 
						end if
				            
				        end if  
			
			
%>
			<tr class="<%=cClass%>">
			    
			
				<td align="center"></td>
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% tempPlantation = rsEPS("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if SprevPlantation <> rsEPS("Plantation") then %>
						<td align="center">
							<a href="P_Summary_LHK_TBM.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsEPS("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsEPS("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsEPS("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevGroupRegion <> rsEPS("GroupRegion") then %>
						<td align="center">
							<a href="P_Summary_LHK_TBM.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsEPS("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsEPS("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if SprevRegion <> rsEPS("Region") then %>
						<td align="center">
							<a href="P_Summary_LHK_TBM.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("region")%>
							</a>
						</td>
						<% SprevRegion = rsEPS("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsEPS("Company")
					   tempfcba = "--ALL--"
					   tempDivision = "--SUMMARY--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if SprevCompany <> rsEPS("Company") then %>
						<td align="center">
							<a href="P_Summary_LHK_TBM.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
							<%=rsEPS("company")%>
							</a>
						</td>
						<% SprevCompany = rsEPS("Company")
                           						%>	
					<% else	%>
						<td align="center"></td>	
						<% GantiCompany = "True"    %>
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsEPS("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsEPS("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsEPS("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsEPS("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsEPS("fcba")
					   tempDivision = "--ALL--"
					   tempPlantingDate = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"					   
					 %>
					
					<% if sprevFcba <> rsEPS("fcba") then %>
						<td align="center">
							<a href="P_Summary_LHK_TBM.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">
						    <%=rsEPS("fcba")%>
							</a>
						</td>
						<% sprevFcba = rsEPS("fcba") 
						   %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				

				
				<td align="right"><%=FormatNumber(rsEPS("Act_Stack_hi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Act_Stack_Shi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("budget_Stack_hi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("budget_Stack_Shi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("var_stack_shi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("varPer_stack_shi"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Act_Stack_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("budget_Stack_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("var_stack_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("varPer_stack_sbi"),0)%></td>
				
                <td align="right"><%=FormatNumber(rsEPS("Act_Tnm_hi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("Act_Tnm_Shi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("budget_Tnm_hi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("budget_Tnm_Shi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("var_Tnm_shi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("varPer_Tnm_shi"),0)%></td>
				
				<td align="right"><%=FormatNumber(rsEPS("Act_Tnm_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("budget_Tnm_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("var_Tnm_sbi"),0)%></td>
				<td align="right"><%=FormatNumber(rsEPS("varPer_Tnm_sbi"),0)%></td>
				
				
				<%
                
                %>
			</tr>
<%
			NO = NO + 1
			    ' Tot Per Company
				TotCmp_Act_Tnm_hi = TotCmp_Act_Tnm_hi + cdbl(rsEPS("Act_Tnm_hi"))
				TotCmp_Act_Tnm_Shi = TotCmp_Act_Tnm_Shi + cdbl(rsEPS("Act_Tnm_shi"))
				TotCmp_budget_Tnm_hi = TotCmp_budget_Tnm_hi + cdbl(rsEPS("budget_tnm_hi"))
				TotCmp_budget_Tnm_Shi = TotCmp_budget_Tnm_shi  + cdbl(rsEPS("budget_tnm_shi"))
				TotCmp_Act_Tnm_sbi = TotCmp_Act_Tnm_sbi + cdbl(rsEPS("Act_tnm_sbi"))
				TotCmp_budget_Tnm_sbi = TotCmp_budget_Tnm_sbi   + cdbl(rsEPS("budget_tnm_sbi")) 
				
				TotCmp_Act_Stack_hi = TotCmp_Act_Stack_hi + cdbl(rsEPS("Act_Stack_hi"))
				TotCmp_Act_Stack_Shi = TotCmp_Act_Stack_Shi + cdbl(rsEPS("Act_Stack_shi"))
				TotCmp_budget_Stack_hi = TotCmp_budget_Stack_hi + cdbl(rsEPS("budget_Stack_hi"))
				TotCmp_budget_Stack_Shi = TotCmp_budget_Stack_shi  + cdbl(rsEPS("budget_Stack_shi"))
				TotCmp_Act_Stack_sbi = TotCmp_Act_Stack_sbi + cdbl(rsEPS("Act_Stack_sbi"))
				TotCmp_budget_Stack_sbi = TotCmp_budget_Stack_sbi   + cdbl(rsEPS("budget_Stack_sbi")) 
				
				
				
				
				
				' Grand Total
			    Gtot_Act_Tnm_hi = Gtot_Act_Tnm_hi + cdbl(rsEPS("Act_Tnm_hi"))
				Gtot_Act_Tnm_Shi = Gtot_Act_Tnm_Shi + cdbl(rsEPS("Act_Tnm_shi"))
				Gtot_budget_Tnm_hi = Gtot_budget_Tnm_hi + cdbl(rsEPS("budget_tnm_hi"))
				Gtot_budget_Tnm_Shi = Gtot_budget_Tnm_shi  + cdbl(rsEPS("budget_tnm_shi"))
				Gtot_Act_Tnm_sbi = Gtot_Act_Tnm_sbi + cdbl(rsEPS("Act_tnm_sbi"))
				Gtot_budget_Tnm_sbi = Gtot_budget_Tnm_sbi   + cdbl(rsEPS("budget_tnm_sbi")) 
				
				Gtot_Act_Stack_hi = Gtot_Act_Stack_hi + cdbl(rsEPS("Act_Stack_hi"))
				Gtot_Act_Stack_Shi = Gtot_Act_Stack_Shi + cdbl(rsEPS("Act_Stack_shi"))
				Gtot_budget_Stack_hi = Gtot_budget_Stack_hi + cdbl(rsEPS("budget_Stack_hi"))
				Gtot_budget_Stack_Shi = Gtot_budget_Stack_shi  + cdbl(rsEPS("budget_Stack_shi"))
				Gtot_Act_Stack_sbi = Gtot_Act_Stack_sbi + cdbl(rsEPS("Act_Stack_sbi"))
				Gtot_budget_Stack_sbi = Gtot_budget_Stack_sbi   + cdbl(rsEPS("budget_Stack_sbi")) 
               
			rsEPS.MoveNext			
		Wend
	end if
	rsEPS.Close
	set rsEPS = Nothing
	
	
       gtotvar_tnm_shi = Gtot_Act_Tnm_Shi - Gtot_budget_Tnm_Shi
	   if Gtot_budget_Tnm_Shi = 0 then
	      gtotvarper_tnm_shi = 0
	   else 
	      gtotvarper_tnm_shi = (Gtot_Act_Tnm_Shi - Gtot_budget_Tnm_Shi) / Gtot_budget_Tnm_Shi * 100
       end if	  
	   
	   gtotvar_tnm_sbi = Gtot_Act_Tnm_Sbi - Gtot_budget_Tnm_Sbi
	   if Gtot_budget_Tnm_Sbi = 0 then
	      gtotvarper_tnm_sbi = 0
	   else 
	      gtotvarper_tnm_sbi = (Gtot_Act_Tnm_Sbi - Gtot_budget_Tnm_Sbi) / Gtot_budget_Tnm_Sbi * 100
       end if
	   
	   gtotvar_Stack_shi = Gtot_Act_Stack_Shi - Gtot_budget_Stack_Shi
	    if Gtot_budget_Stack_Shi = 0 then
	      gtotvarper_Stack_shi = 0
	   else 
	      gtotvarper_Stack_shi = (Gtot_Act_Stack_Shi - Gtot_budget_Stack_Shi) / Gtot_budget_Stack_Shi * 100
       end if	  
	   
	   gtotvar_Stack_sbi = Gtot_Act_Stack_Sbi - Gtot_budget_Stack_Sbi
	   if Gtot_budget_Stack_Sbi = 0 then
	      gtotvarper_Stack_sbi = 0
	   else 
	      gtotvarper_Stack_sbi = (Gtot_Act_Stack_Sbi - Gtot_budget_Stack_Sbi) / Gtot_budget_Stack_Sbi * 100
       end if
%>
		             <%	
			                       TotCmpvar_tnm_shi = TotCmp_Act_Tnm_Shi - TotCmp_budget_Tnm_Shi
								   if TotCmp_budget_Tnm_Shi = 0 then
									  TotCmpvarper_tnm_shi = 0
								   else 
									  TotCmpvarper_tnm_shi = (TotCmp_Act_Tnm_Shi - TotCmp_budget_Tnm_Shi) / TotCmp_budget_Tnm_Shi * 100
								   end if	  
								   
								   TotCmpvar_tnm_sbi = TotCmp_Act_Tnm_Sbi - TotCmp_budget_Tnm_Sbi
								   if TotCmp_budget_Tnm_Sbi = 0 then
									  TotCmpvarper_tnm_sbi = 0
								   else 
									  TotCmpvarper_tnm_sbi = (TotCmp_Act_Tnm_Sbi - TotCmp_budget_Tnm_Sbi) / TotCmp_budget_Tnm_Sbi * 100
								   end if
								   
								   TotCmpvar_Stack_shi = TotCmp_Act_Stack_Shi - TotCmp_budget_Stack_Shi
									if TotCmp_budget_Stack_Shi = 0 then
									  TotCmpvarper_Stack_shi = 0
								   else 
									  TotCmpvarper_Stack_shi = (TotCmp_Act_Stack_Shi - TotCmp_budget_Stack_Shi) / TotCmp_budget_Stack_Shi * 100
								   end if	  
								   
								   TotCmpvar_Stack_sbi = TotCmp_Act_Stack_Sbi - TotCmp_budget_Stack_Sbi
								   if TotCmp_budget_Stack_Sbi = 0 then
									  TotCmpvarper_Stack_sbi = 0
								   else 
									  TotCmpvarper_Stack_sbi = (TotCmp_Act_Stack_Sbi - TotCmp_budget_Stack_Sbi) / TotCmp_budget_Stack_Sbi * 100
			                       end if 
					%>
					
								<tr bgcolor="#F4F4A0">
									<td align="center" colspan="<%= 1+ nSpanTotal %>">TOTAL&nbsp;&nbsp<%=tempCompany%></td>				
									<td align="right"><%=FormatNumber(TotCmp_Act_Stack_hi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_Act_Stack_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Stack_hi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Stack_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvar_stack_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvarper_Stack_shi,0)%></td>
									
									<td align="right"><%=FormatNumber(TotCmp_act_Stack_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Stack_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvar_stack_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvarper_Stack_sbi,0)%></td>
											
									<td align="right"><%=FormatNumber(TotCmp_Act_Tnm_hi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_Act_Tnm_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Tnm_hi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Tnm_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvar_tnm_shi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvarper_tnm_shi,0)%></td>
									
									<td align="right"><%=FormatNumber(TotCmp_act_Tnm_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmp_budget_Tnm_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvar_Tnm_sbi,0)%></td>
									<td align="right"><%=FormatNumber(TotCmpvarper_Tnm_sbi,0)%></td>
								</tr>
			
			<tr bgcolor="#F4F6A2">
			   
				<td align="center" colspan="<%= 1+ nSpanTotal %>">GRAND TOTAL</td>				
				<td align="right"><%=FormatNumber(Gtot_Act_Stack_hi,0)%></td>
				<td align="right"><%=FormatNumber(Gtot_Act_Stack_shi,0)%></td>
				<td align="right"><%=FormatNumber(Gtot_budget_Stack_hi,0)%></td>
				<td align="right"><%=FormatNumber(Gtot_budget_Stack_shi,0)%></td>
				<td align="right"><%=FormatNumber(gtotvar_stack_shi,0)%></td>
				<td align="right"><%=FormatNumber(gtotvarper_Stack_shi,0)%></td>
				
				<td align="right"><%=FormatNumber(Gtot_act_Stack_sbi,0)%></td>
				<td align="right"><%=FormatNumber(Gtot_budget_Stack_sbi,0)%></td>
				<td align="right"><%=FormatNumber(gtotvar_stack_sbi,0)%></td>
				<td align="right"><%=FormatNumber(gtotvarper_Stack_sbi,0)%></td>
						
				<td align="right"><%=FormatNumber(Gtot_Act_Tnm_hi,0)%></td>
				<td align="right"><%=FormatNumber(Gtot_Act_Tnm_shi,0)%></td>
				<td align="right"><%=FormatNumber(Gtot_budget_Tnm_hi,0)%></td>
				<td align="right"><%=FormatNumber(Gtot_budget_Tnm_shi,0)%></td>
				<td align="right"><%=FormatNumber(gtotvar_tnm_shi,0)%></td>
				<td align="right"><%=FormatNumber(gtotvarper_tnm_shi,0)%></td>
				
				<td align="right"><%=FormatNumber(Gtot_act_Tnm_sbi,0)%></td>
				<td align="right"><%=FormatNumber(Gtot_budget_Tnm_sbi,0)%></td>
				<td align="right"><%=FormatNumber(gtotvar_Tnm_sbi,0)%></td>
				<td align="right"><%=FormatNumber(gtotvarper_Tnm_sbi,0)%></td>
			</tr>
</tbody>
</table>
</form>
</div>
</body>
</html>

