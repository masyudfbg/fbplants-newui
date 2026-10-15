<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom
	Dim no
	
	Dim ScriptName
	Dim strMain, sqlLHPP, sqlLHPP2, rsLHPP
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType

    v_txtGroupPlantation = Request("txtGroupPlantation")
	v_txtPlantation = Request("txtPlantation")
	v_txtGroupRegion = Request("txtGroupRegion")
	v_txtRegion = Request("txtRegion")
	v_txtcompany = Request("txtCompany")
	v_txtfcba = Request("txtfcba")
	
	v_txtBudgetYear = Request("txtBudgetYear")
	v_txtBudgetVersion = Request("txtBudgetVersion")
	v_txtModuleVersion = Request("txtModuleVersion")
	
	v_txtdatefrom		= Request("txtdatefrom")
	if len(trim(v_txtdatefrom)) = 0 then
	   v_txtdatefrom	= date -2
	   v_txtfcba = "--SUMMARY--"
	   v_txtRegion = "--ALL--"
	end if
	
	If Len(Trim(v_txtBudgetYear)) = 0 Then
		v_txtBudgetYear = Year(textDate(Date))
	Else
		v_txtBudgetYear = Year(v_txtdatefrom)
	End If
	If Len(Trim(v_txtBudgetVersion)) = 0 Then v_txtBudgetVersion = 1
	If Len(Trim(v_txtModuleVersion)) = 0 Then v_txtModuleVersion = 1
	
	If Len(Trim(v_txtdatefrom)) = 0 Then v_txtdatefrom = textDate(Date)
	
	Dim var_colspan : var_colspan = 1
	
Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" and businessunit_api.get_fctype(fcba) = 'M' "

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
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
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
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
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
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
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
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
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				"        , FCCOMPANYCODE" & _
				"        , FCNAME" & _
				"        , FCCOMPANYNAME" & _
				"        , FCMANAGER" & _
				"        , FCKTU" & _
				"        , FCCROP" & _
				"        , FCTYPE" & _
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
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				"        , FCCOMPANYCODE" & _
				"        , FCNAME" & _
				"        , FCCOMPANYNAME" & _
				"        , FCMANAGER" & _
				"        , FCKTU" & _
				"        , FCCROP" & _
				"        , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName = rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName = rsCat("FCCROP")
			FcbaType = rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

ScriptName = Request.ServerVariables("SCRIPT_NAME")

strMain = " "
		  
sqlLHPP2 = " "

sqlLHPP = "Select  CASE " & _ 
		  "		 WHEN REGION IS NOT NULL AND FCBA IS NULL AND MILL_CAP IS NULL THEN " & _
		  "		  'Total ' || Region " & _
		  "      WHEN REGION IS NULL AND FCBA IS NULL AND MILL_CAP IS NULL THEN" & _
          "       'Grand Total '" & _
		  "		 ELSE" & _
		  "		   Region" & _
		  "	   END Region," & _
		  "    case" & _
          "       when region is null and fcba is null and mill_cap is null then" & _
          "        10000 + SEQUENCE_REGION" & _
          "       else" & _
          "        SEQUENCE_REGION" & _
          "    end nextSequence," & _
		  "	   SEQUENCE_REGION," & _
		  "		   FCBA," & _
		  "		   MILL_CAP," & _
		  "		   SUM(FFB_HI) FFB_HI," & _
		  "		   SUM(FFB_SHI) FFB_SHI," & _
		  "        SUM(ffb_Process_HI) ffb_Process_HI," & _
		  "        SUM(ffb_Process_SHI) ffb_Process_SHI," & _
		  "		   SUM(FFB_BUDGET_SHI) FFB_BUDGET_SHI," & _
		  "		   SUM(FFB_VAR_SHI) FFB_VAR_SHI," & _
		  "		   SUM(FFB_VAR_PERCENT_SHI) FFB_VAR_PERCENT_SHI," & _
		  "		   SUM(THROUGHPUT_HI) THROUGHPUT_HI," & _
		  "		   SUM(THROUGHPUT_SHI) THROUGHPUT_SHI," & _
		  "		   SUM(CPO_ACT_HI)  CPO_ACT_HI," & _
		  "		   SUM(CPO_ACT_SHI) CPO_ACT_SHI," & _		  
		  "		   SUM(OER_ACT_HI) OER_ACT_HI," & _
		  "		   SUM(OER_ACT_SHI) OER_ACT_SHI," & _
		  "		   SUM(OER_BUD_SHI) OER_BUD_SHI," & _
		  "		   SUM(OER_VAR_PERCENT) OER_VAR_PERCENT," & _
		  "		   SUM(KER_ACT_HI) KER_ACT_HI," & _
		  "		   SUM(KER_ACT_SHI) KER_ACT_SHI," & _
		  "		   SUM(KER_BUD_SHI) KER_BUD_SHI," & _
		  "		   SUM(KER_VAR_PERCENT) KER_VAR_PERCENT," & _
		  "		   SUM(FFA_ACT_HI) FFA_ACT_HI," & _
		  "		   SUM(FFA_ACT_SHI) FFA_ACT_SHI," & _
		  " 	   SUM(FFA_BUD_SHI) FFA_BUD_SHI," & _
		  "		   SUM(FFA_VAR_PERCENT) FFA_VAR_PERCENT" & _
		  "	  from ( " & _
          "Select  SEQUENCE_REGION," & _
		  "		   REGION," & _
		  "		   FCBA," & _
		  "		   MILL_CAP," & _
		  "		   SUM(FFB_HI) FFB_HI," & _
		  "		   SUM(FFB_SHI) FFB_SHI," & _
		  "        SUM(ffb_Process_HI) ffb_Process_HI," & _
		  "        SUM(ffb_Process_SHI) ffb_Process_SHI," & _
		  "		   SUM(FFB_BUDGET_SHI) FFB_BUDGET_SHI," & _
		  "		   SUM(FFB_VAR_SHI) FFB_VAR_SHI," & _
		  "		   SUM(FFB_VAR_PERCENT_SHI) FFB_VAR_PERCENT_SHI," & _
		  "		   SUM(THROUGHPUT_HI) THROUGHPUT_HI," & _
		  "		   SUM(THROUGHPUT_SHI) THROUGHPUT_SHI," & _
		  "		   SUM(CPO_ACT_HI)  CPO_ACT_HI," & _
		  "		   SUM(CPO_ACT_SHI) CPO_ACT_SHI," & _
		  "		   SUM(OER_ACT_HI) OER_ACT_HI," & _
		  "		   SUM(OER_ACT_SHI) OER_ACT_SHI," & _
		  "		   SUM(OER_BUD_SHI) OER_BUD_SHI," & _
		  "		   SUM(OER_VAR_PERCENT) OER_VAR_PERCENT," & _
		  "		   SUM(KER_ACT_HI) KER_ACT_HI," & _
		  "		   SUM(KER_ACT_SHI) KER_ACT_SHI," & _
		  "		   SUM(KER_BUD_SHI) KER_BUD_SHI," & _
		  "		   SUM(KER_VAR_PERCENT) KER_VAR_PERCENT," & _
		  "		   SUM(FFA_ACT_HI) FFA_ACT_HI," & _
		  "		   SUM(FFA_ACT_SHI) FFA_ACT_SHI," & _
		  " 	   SUM(FFA_BUD_SHI) FFA_BUD_SHI," & _
		  "		   SUM(FFA_VAR_PERCENT) FFA_VAR_PERCENT" & _
		  "	  from (select case" & _
		  "					 when a.company = 'PT.THIP' then" & _
		  "					  1" & _
		  "					 when a.company = 'PT.CRS' then" & _
		  "					  3" & _
		  "					 else" & _
		  "					  2" & _
		  "				   end sequence_region," & _
		  "				   case" & _
		  "					 when a.company = 'PT.THIP' then" & _
		  "					  'THIP'" & _
		  "					 when a.company = 'PT.CRS' then" & _
		  "					  'CRS'" & _
		  "					 else" & _
		  "					  'AMS Ganda'" & _
		  "				   end region," & _
		  "				   a.fcba," & _
		  "				   a.mill_cap," & _
		  "				   case" & _
		  "					 when a.reference = 'HI' then" & _
		  "					  a.ffb_total" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ffb_HI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ffb_total" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ffb_SHI," & _	  
		  "				   case" & _
		  "					 when a.reference = 'HI' then" & _
		  "					  a.ffb_process" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ffb_Process_HI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ffb_process" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ffb_process_SHI," & _	  
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ffb_budget" & _
          "					 else" & _
		  "					  0" & _
		  "				   end ffb_budget_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ffb_var_mt" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ffb_var_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ffb_var_percent" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ffb_var_percent_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'HI' then" & _
		  "					  a.throughput" & _
		  "	 				 else" & _
		  "					  0" & _
		  "				   end throughput_HI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.throughput" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end throughput_SHI," & _	  
		  "				   case" & _
		  "					 when a.reference = 'HI' then" & _
		  " 					  a.cpo_act" & _
	      "					 else" & _
		  "					  0" & _
		  "				   end cpo_act_HI," & _
		  " 				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.cpo_act" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end cpo_act_SHI," & _	  
		  "				   case" & _
		  "					 when a.reference = 'HI' then" & _
		  " 					  a.oer_act" & _
	      "					 else" & _
		  "					  0" & _
		  "				   end oer_act_HI," & _
		  " 				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.oer_act" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end oer_act_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.oer_bud" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end oer_bud_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.oer_var_percent" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end oer_var_percent," & _
		  "				   case" & _
		  "					 when a.reference = 'HI' then" & _
		  "					  a.ker_act" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ker_act_HI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ker_act" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ker_act_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ker_bud" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ker_bud_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ker_var_percent" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ker_var_percent," & _
		  "				   case" & _
		  "					 when a.reference = 'HI' then" & _
		  "					  a.ffa_act" & _
		  " 					 else" & _
		  "					  0" & _
		  "				   end ffa_act_HI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ffa_act" & _
		  "					 else" & _
		  "					  0" & _
		  "				   end ffa_act_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ffa_bud" & _
		  "					 else" & _
		  "					  0" & _
		  "  				   end ffa_bud_SHI," & _
		  "				   case" & _
		  "					 when a.reference = 'SHI' then" & _
		  "					  a.ffa_var_percent" & _
		  "					 else" & _
		  "					  0" & _
		  " 				   end ffa_var_percent" & _
		  "			  from mgt_daily_rpt_mill_tab a" & _
		  "			 where trunc(ondate) = trunc(to_date('" & v_txtdatefrom & "', 'mm/dd/yyyy')))" & _
		  "	 group by SEQUENCE_REGION, REGION, FCBA, MILL_CAP ) " & _
		  "  GROUP BY GROUPING SETS(( REGION,SEQUENCE_REGION, FCBA,MILL_CAP),( REGION,SEQUENCE_REGION),())" & _
		  "	 order by SEQUENCE_REGION, businessunit_api.get_OrderData(fcba),nextsequence"
 
'response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Summary Harian Produksi Pabrik</title>
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
				@media print{.printhide{display:none;}}
			</style>
       </head>


	<body>
		<div id="wrapper">
		<div class="printhide">
		<!--#include file="../print_excel.asp"-->
		</div>
			<form name="frmEntry" method="post" action="<%=ScriptName%>">
				<div class="printhide">
				<!--#include file="menutop.asp"-->
				<!--#include file="FormSelectNewSumMNew.html"-->
				
					<table id="form-table" class="centered-margin">
						<tr>
							<td>	<!--<font face="verdana" size="1">BUDGET YEAR</font>--></td>
							<td>	<!--<font face="verdana" size="1">BUDGET VERSION</font>--></td>
							<td>	<!--<font face="verdana" size="1">MODULE VERSION</font>--></td>
							<td><font face="verdana" size="1">DATE</font></td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td align="center">
								<input type="hidden" name="txtBudgetYear" value="<%=v_txtBudgetYear%>">
							</td>
							<td align="center">
								<input type="hidden" name="txtBudgetVersion" value="<%=v_txtBudgetVersion%>">
							</td>
							<td align="center">
								<input type="hidden" name="txtModuleVersion" value="<%=v_txtModuleVersion%>">
							</td>
							<td align="center">
								<input id="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>" title="Type Date mm/dd/yyyy" >
							</td>	
							<td align="center">
								<input type="submit" value="View" />
							</td>							
						</tr>
					</table>
				</div>
			</form>

			<table border="1" width="100%" id="product-table" class='xthf-site'>
				<thead>
					<tr>
						<th colspan="33" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<h1>AMS PLANTATION</h1>
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN HARIAN PRODUKSI PABRIK</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Tanggal: <%=Day(v_txtdatefrom)%>&nbsp;<%=Bulan(Month(v_txtdatefrom))%>&nbsp;<%=Year(v_txtdatefrom)%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th rowspan="3">MILL</th>
						<th rowspan="3">Cap.</th>
						<th colspan="4">FFB PROCESSED (KG</th>
						<th rowspan="2" colspan="2">THROUGHPUT</th>
						<th colspan="3">OER (%)</th>
						<th colspan="3">FFA (%)</th>
						<th colspan="3">KER (%)</th>
					</tr>
					<tr>
						
						<th rowspan="2">HI</th>
						<th colspan="3">SHI</th>
						<th rowspan="2">HI</th>
						<th colspan="2">SHI</th>
						<th rowspan="2">HI</th>
						<th colspan="2">SHI</th>
						<th rowspan="2">HI</th>
						<th colspan="2">SHI</th>
					</tr>
					<tr>
						
						<th>ACT</th>
						<th>BGT</th>
						<th>VAR</th>
						<th>HI</th>
						<th>SHI</th>
						<th>ACT</th>
						<th>BGT</th>
						<th>ACT</th>
						<th>BGT</th>
						<th>ACT</th>
						<th>BGT</th>
					</tr>
				</thead>
				
				<tbody>
<%
Set rsLHPP = Server.CreateObject("ADODB.Recordset")
rsLHPP.Open sqlLHPP, connstr, adOpenStatic

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""

Dim currfcba : currfcba = ""
Dim prevfcba : prevfcba = ""

Dim currnextregion : currnextregion = ""
Dim prevnextregion : prevnextregion = ""

Dim TEMPSEQUENCE_REGION : TEMPSEQUENCE_REGION = ""
Dim prevnextregion1 : prevnextregion1 = ""

if rsLHPP.BOF and rsLHPP.EOF then
%>
					<tr>
						<td colspan="<%=33%>">NO DATA AVAILABLE</td>
					</tr>
<%
else
		while not rsLHPP.eof

            currfcba = rsLHPP("FCBA")
			currnextregion =  rsLHPP("REGION")
%>

					
					<%If (rsLHPP("REGION") = "Grand Total" Or Left(Cstr(rsLHPP("REGION")), 5) = "Total") Then%>
					<tr >
						<%If currnextregion <> prevnextregion or rsLHPP("REGION") = "Grand Total" then%>
							<td align="Center" colspan="2" style=" border-bottom:none;">
								<strong><%=rsLHPP("REGION")%></strong>
							</td>
							<%prevnextregion = currnextregion%>
						<%Else%>
							<td align="center" colspan="2" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
						<%End If%>
										
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_HI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_SHI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_BUDGET_SHI"), 2)%></strong></td>
						<% IF cdbl(rsLHPP("FFB_BUDGET_SHI")) > 0 then %>
						<td align="right"><strong><%=FormatNumber((cdbl(rsLHPP("FFB_SHI")) - cdbl(rsLHPP("FFB_BUDGET_SHI")) )/cdbl(rsLHPP("FFB_BUDGET_SHI"))*100, 2)%></strong></td>
						<% else %>
						<td align="right"><strong>0</strong></td>
						<% end if %>
						<td align="right"><strong><%=FormatNumber(rsLHPP("THROUGHPUT_HI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("THROUGHPUT_SHI"), 2)%></strong></td>
						<% IF cdbl(rsLHPP("ffb_process_HI")) > 0 then %>
						<td align="right"><strong><%=FormatNumber(cdbl(rsLHPP("CPO_ACT_HI"))/cdbl(rsLHPP("ffb_process_HI"))*100, 2)%></strong></td>
						<% else %>
						<td align="right"><strong>0</strong></td>
						<% end if %>
						<% IF cdbl(rsLHPP("ffb_process_SHI")) > 0 then %>
						<td align="right"><%=FormatNumber(cdbl(rsLHPP("CPO_ACT_SHI"))/cdbl(rsLHPP("ffb_process_SHI"))*100, 2)%></td>
						<% else %>
						<td align="right">0</strong></td>
						<% end if %>	
						<td align="right"><strong><%=FormatNumber(rsLHPP("OER_BUD_SHI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KER_ACT_HI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KER_ACT_SHI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KER_BUD_SHI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFA_ACT_HI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFA_ACT_SHI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFA_BUD_SHI"), 2)%></strong></td>
					</tr>
					<%Else%>
					<tr>
					    <%If currfcba <> prevfcba then%>
							<td align="Center" style=" border-bottom:none;"><%=rsLHPP("FCBA")%></td>
							<%prevfcba = currfcba%>
							<td align="Center" style=" border-bottom:none;"><%=FormatNumber(rsLHPP("MILL_CAP"), 2)%></td>
						<%Else%>
							<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
							<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
						<%End If%>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_HI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_SHI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_BUDGET_SHI"), 2)%></td>
						<% IF cdbl(rsLHPP("FFB_BUDGET_SHI")) > 0 then %>
						<td align="right"><strong><%=FormatNumber((cdbl(rsLHPP("FFB_SHI")) - cdbl(rsLHPP("FFB_BUDGET_SHI")) )/cdbl(rsLHPP("FFB_BUDGET_SHI"))*100, 2)%></strong></td>
						<% else %>
						<td align="right"><strong>0</strong></td>
						<% end if %>
						<td align="right"><%=FormatNumber(rsLHPP("THROUGHPUT_HI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("THROUGHPUT_SHI"), 2)%></td>
						<% IF cdbl(rsLHPP("ffb_process_HI")) > 0 then %>
						<td align="right"><%=FormatNumber(cdbl(rsLHPP("CPO_ACT_HI"))/cdbl(rsLHPP("ffb_process_HI"))*100, 2)%></td>
						<% else %>
						<td align="right">0</strong></td>
						<% end if %>
						<% IF cdbl(rsLHPP("ffb_process_SHI")) > 0 then %>
						<td align="right"><%=FormatNumber(cdbl(rsLHPP("CPO_ACT_SHI"))/cdbl(rsLHPP("ffb_process_SHI"))*100, 2)%></td>
						<% else %>
						<td align="right">0</strong></td>
						<% end if %>					
						<td align="right"><%=FormatNumber(rsLHPP("OER_BUD_SHI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_ACT_HI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_ACT_SHI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_BUD_SHI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_ACT_HI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_ACT_SHI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_BUD_SHI"), 2)%></td>
					</tr>
					<%End If%>


<%				    
				TEMPSEQUENCE_REGION =  rsLHPP("SEQUENCE_REGION")
				rsLHPP.MoveNext
			Wend
%>

<%
End if
rsLHPP.Close
set rsLHPP = Nothing
%>

				</tbody>
			</table>

		</div>
	</body>
</html>
