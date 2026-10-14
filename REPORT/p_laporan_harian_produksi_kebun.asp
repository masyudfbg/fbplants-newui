<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom
	
	Dim ScriptName
	Dim strMain, strLHPK, strLHPK2, rsLHPK
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType
	Dim rsCountRow, strCountRow, SortNumber

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
					
wherecons	=	" and businessunit_api.get_fctype(fcba) = 'E' "

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
		  
strLHPK2 = " "


strLHPK = "   SELECT CASE" & _
          "             WHEN GroupRegion IS NOT NULL AND Region IS NULL AND fcba IS NULL" & _
          "             THEN" & _
          "                'Total ' || GroupRegion" & _
          "             ELSE" & _
          "                GroupRegion" & _
          "          END" & _
          "             GroupRegion," & _
          "          orderdatagroupregion," & _
          "          region," & _
          "          fcba," & _
          "          SUM (HA_TM) HA_TM," & _
          "          SUM (PROD_EST_HI) PROD_EST_HI," & _
          "          SUM (PROD_EST_SHI) PROD_EST_SHI," & _
          "          SUM (PROD_MILL_HI) PROD_MILL_HI," & _
          "          SUM (PROD_MILL_SHI) PROD_MILL_SHI," & _
          "          SUM (BUD_PROD_HI) BUD_PROD_HI," & _
          "          SUM (BUD_PROD_SHI) BUD_PROD_SHI," & _
          "          SUM (VAR_SHI_KG) VAR_SHI_KG," & _
          "          SUM (VAR_SHI_PERCENT) VAR_SHI_PERCENT," & _
          "          SUM (RESTAN_HI) RESTAN_HI," & _
          "          SUM (PROD_MILL_SBI) PROD_MILL_SBI," & _
          "          SUM (BUD_PROD_SBI) BUD_PROD_SBI," & _
          "          SUM (VAR_SBI_KG) VAR_SBI_KG," & _
          "          SUM (VAR_SBI_PERCENT) VAR_SBI_PERCENT," & _
          "          SUM (PUSINGAN_AVG) PUSINGAN_AVG," & _
          "          SUM (PUSINGAN_UPTOTEN) PUSINGAN_UPTOTEN," & _
          "          SUM (MANDAYS_HI) MANDAYS_HI," & _
          "          SUM (MANDAYS_SHI) MANDAYS_SHI," & _
          "          SUM (MANDAYS_SB_HI) MANDAYS_SB_HI," & _
          "          SUM (MANDAYS_SB_SHI) MANDAYS_SB_SHI," & _
          "          SUM (PERCENT_SB_HI) PERCENT_SB_HI," & _
          "          SUM (PERCENT_SB_SHI) PERCENT_SB_SHI," & _
          "          SUM (WI_VOL_HI) WI_VOL_HI," & _
          "          SUM (WI_VOL_SHI) WI_VOL_SHI," & _
          "          SUM (WI_HH_SHI) WI_HH_SHI," & _
          "          SUM (WI_VOL_SBI) WI_VOL_SBI," & _
          "          SUM (WI_HH_SBI) WI_HH_SBI" & _
          "     FROM (  SELECT CASE WHEN company = 'PT.THIP' THEN 'THIP' ELSE 'AMS Ganda' END" & _
          "                       groupregion," & _
          "                    CASE WHEN company = 'PT.THIP' THEN 1 ELSE 2 END" & _
          "                       orderdatagroupregion," & _
          "                    region," & _
          "                    fcba," & _
          "                    SUM (HA_TM) HA_TM," & _
          "                    SUM (PROD_EST_HI) PROD_EST_HI," & _
          "                    SUM (PROD_EST_SHI) PROD_EST_SHI," & _
          "                    SUM (PROD_MILL_HI) PROD_MILL_HI," & _
          "                    SUM (PROD_MILL_SHI) PROD_MILL_SHI," & _
          "                    SUM (BUD_PROD_HI) BUD_PROD_HI," & _
          "                    SUM (BUD_PROD_SHI) BUD_PROD_SHI," & _
          "                    SUM (VAR_SHI_KG) VAR_SHI_KG," & _
          "                    SUM (VAR_SHI_PERCENT) VAR_SHI_PERCENT," & _
          "                    SUM (RESTAN_HI) RESTAN_HI," & _
          "                    SUM (PROD_MILL_SBI) PROD_MILL_SBI," & _
          "                    SUM (BUD_PROD_SBI) BUD_PROD_SBI," & _
          "                    SUM (VAR_SBI_KG) VAR_SBI_KG," & _
          "                    SUM (VAR_SBI_PERCENT) VAR_SBI_PERCENT," & _
          "                    SUM (PUSINGAN_AVG) PUSINGAN_AVG," & _
          "                    SUM (PUSINGAN_UPTOTEN) PUSINGAN_UPTOTEN," & _
          "                    SUM (MANDAYS_HI) MANDAYS_HI," & _
          "                    SUM (MANDAYS_SHI) MANDAYS_SHI," & _
          "                    SUM (MANDAYS_SB_HI) MANDAYS_SB_HI," & _
          "                    SUM (MANDAYS_SB_SHI) MANDAYS_SB_SHI," & _
          "                    SUM (PERCENT_SB_HI) PERCENT_SB_HI," & _
          "                    SUM (PERCENT_SB_SHI) PERCENT_SB_SHI," & _
          "                    SUM (WI_VOL_HI) WI_VOL_HI," & _
          "                    SUM (WI_VOL_SHI) WI_VOL_SHI," & _
          "                    SUM (WI_HH_SHI) WI_HH_SHI," & _
          "                    SUM (WI_VOL_SBI) WI_VOL_SBI," & _
          "                    SUM (WI_HH_SBI) WI_HH_SBI" & _
          "               FROM mgt_daily_rpt_est_tab" & _
          "               WHERE to_char(ondate, 'yyyymmdd') = to_char(to_date('" & v_txtdatefrom & "', 'mm/dd/yyyy'), 'yyyymmdd')" & _
          "                     " & whereCons & "" & _
          "           GROUP BY company, region, fcba)" & _
          " GROUP BY GROUPING SETS ( (GroupRegion," & _
          "                           orderdatagroupregion," & _
          "                           region," & _
          "                           fcba)," & _
          "                         (GroupRegion, orderdatagroupregion))" & _
          " ORDER BY orderdatagroupregion, REGION_API.get_OrderData (region), BUSINESSUNIT_API.get_OrderData (fcba)"

'response.write strLHPK
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Harian Produksi Kebun</title>
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
				<!--#include file="FormSelectNewSumE.html"-->
				
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

			<table border="0" width="100%" id="product-table" class='xthf-site'>
				<thead>
					<tr>
						<th colspan="30" style="background-color:#FFFFFF; border:none;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<h1>AMS PLANTATION</h1>
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN HARIAN PRODUKSI KEBUN</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Tanggal: <%=Day(v_txtdatefrom)%>&nbsp;<%=Bulan(Month(v_txtdatefrom))%>&nbsp;<%=Year(v_txtdatefrom)%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th rowspan="4">WILAYAH</th>
						<th rowspan="4">KEBUN</th>
						<th rowspan="4">Luas (Ha) TM</th>
						<th colspan="9">BULAN INI (BI)</th>
						<th colspan="4">SAMPAI DENGAN BULAN INI (SBI)</th>
						<th colspan="2">PUSINGAN</th>
						<th colspan="6">TK POTONG BUAH (HK)</th>
						<th colspan="5">CURAH HUJAN</th>
					</tr>
					<tr>
						<th colspan="4">AKTUAL PRODUKSI (Kg)</th>
						<th rowspan="2" colspan="2">BUDGET (Kg)</th>
						<th rowspan="2" colspan="2">VAR SHI</th>
						<th rowspan="3">RESTAN</th>
						<th rowspan="3">PMKS (Kg)</th>
						<th rowspan="3">BUDGET (Kg)</th>
						<th rowspan="2" colspan="2">VAR SBI</th>
						<th rowspan="2">RATA-RATA</th>
						<th rowspan="2">>10 HARI (%)</th>
						<th rowspan="2" colspan="2">PEMANEN</th>
						<th rowspan="2" colspan="2">SIAP BORONG</th>
						<th rowspan="2" colspan="2">% SIAP BORONG</th>
						<th colspan="3">BI</th>
						<th colspan="2">SBI</th>
					</tr>
					<tr>
						<th colspan="2">KEBUN</th>
						<th colspan="2">PMKS</th>
						<th colspan="2">VOL (mm)</th>
						<th>HH</th>
						<th rowspan="2">VOL (mm)</th>
						<th rowspan="2">HH</th>
					</tr>
					<tr>
						<th>HI</th>
						<th>SHI</th>
						<th>HI</th>
						<th>SHI</th>
						<th>HI</th>
						<th>SHI</th>
						<th>Kg</th>
						<th>%</th>
						<th>Kg</th>
						<th>%</th>
						<th>HI</th>
						<th>HI</th>
						<th>HI</th>
						<th>SHI</th>
						<th>HI</th>
						<th>SHI</th>
						<th>HI</th>
						<th>SHI</th>
						<th>HI</th>
						<th>SHI</th>
						<th>SHI</th>
					</tr>
					<tr>
						<td align="Center" colspan="30" style="padding-left:10px;"  class="separator">
						</td>
					</tr>
				</thead>
				
				<tbody>
<%
Set rsLHPK = Server.CreateObject("ADODB.Recordset")
rsLHPK.Open strLHPK, connstr, adOpenStatic

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

Dim CurrRegion : CurrRegion = ""
Dim PrevRegion : PrevRegion = ""

Dim CurrGroupRegion : CurrGroupRegion = ""
Dim PrevGroupRegion : PrevGroupRegion = ""

Dim st_gr_ha_tm : st_gr_ha_tm = 0
Dim st_gr_prod_est_hi : st_gr_prod_est_hi = 0
Dim st_gr_prod_est_shi : st_gr_prod_est_shi = 0
Dim st_gr_prod_mill_hi : st_gr_prod_mill_hi = 0
Dim st_gr_prod_mill_shi : st_gr_prod_mill_shi = 0
Dim st_gr_bud_prod_hi : st_gr_bud_prod_hi = 0
Dim st_gr_bud_prod_shi : st_gr_bud_prod_shi = 0
Dim st_gr_var_shi_kg : st_gr_var_shi_kg = 0
Dim st_gr_var_shi_percent : st_gr_var_shi_percent = 0
Dim st_gr_restan_hi : st_gr_restan_hi = 0
Dim st_gr_prod_mill_sbi : st_gr_prod_mill_sbi = 0
Dim st_gr_bud_prod_sbi : st_gr_bud_prod_sbi = 0
Dim st_gr_var_sbi_kg : st_gr_var_sbi_kg = 0
Dim st_gr_var_sbi_percent : st_gr_var_sbi_percent = 0
Dim st_gr_pusingan_avg : st_gr_pusingan_avg = 0
Dim st_gr_pusingan_uptoten : st_gr_pusingan_uptoten = 0
Dim st_gr_mandays_hi : st_gr_mandays_hi = 0
Dim st_gr_mandays_shi : st_gr_mandays_shi = 0
Dim st_gr_mandays_sb_hi : st_gr_mandays_sb_hi = 0
Dim st_gr_mandays_sb_shi : st_gr_mandays_sb_shi = 0
Dim st_gr_percent_sb_hi : st_gr_percent_sb_hi = 0
Dim st_gr_percent_sb_shi : st_gr_percent_sb_shi = 0
Dim st_gr_wi_vol_hi : st_gr_wi_vol_hi = 0
Dim st_gr_wi_vol_shi : st_gr_wi_vol_shi = 0
Dim st_gr_wi_hh_shi : st_gr_wi_hh_shi = 0
Dim st_gr_wi_vol_sbi : st_gr_wi_vol_sbi = 0
Dim st_gr_wi_hh_sbi : st_gr_wi_hh_sbi = 0

Dim Row_var_shi_kg : Row_var_shi_kg = 0
Dim Row_var_shi_percent : Row_var_shi_percent = 0

Dim Row_var_sbi_kg : Row_var_sbi_kg = 0
Dim Row_var_sbi_percent : Row_var_sbi_percent = 0

Dim Row_percent_sb_hi : Row_percent_sb_hi = 0
Dim Row_percent_sb_shi : Row_percent_sb_shi = 0

Dim st_ha_tm : st_ha_tm = 0
Dim st_prod_est_hi : st_prod_est_hi = 0
Dim st_prod_est_shi : st_prod_est_shi = 0
Dim st_prod_mill_hi : st_prod_mill_hi = 0
Dim st_prod_mill_shi : st_prod_mill_shi = 0
Dim st_bud_prod_hi : st_bud_prod_hi = 0
Dim st_bud_prod_shi : st_bud_prod_shi = 0
Dim st_var_shi_kg : st_var_shi_kg = 0
Dim st_var_shi_percent : st_var_shi_percent = 0
Dim st_restan_hi : st_restan_hi = 0
Dim st_prod_mill_sbi : st_prod_mill_sbi = 0
Dim st_bud_prod_sbi : st_bud_prod_sbi = 0
Dim st_var_sbi_kg : st_var_sbi_kg = 0
Dim st_var_sbi_percent : st_var_sbi_percent = 0
Dim st_pusingan_avg : st_pusingan_avg = 0
Dim st_pusingan_uptoten : st_pusingan_uptoten = 0
Dim st_mandays_hi : st_mandays_hi = 0
Dim st_mandays_shi : st_mandays_shi = 0
Dim st_mandays_sb_hi : st_mandays_sb_hi = 0
Dim st_mandays_sb_shi : st_mandays_sb_shi = 0
Dim st_percent_sb_hi : st_percent_sb_hi = 0
Dim st_percent_sb_shi : st_percent_sb_shi = 0
Dim st_wi_vol_hi : st_wi_vol_hi = 0
Dim st_wi_vol_shi : st_wi_vol_shi = 0
Dim st_wi_hh_shi : st_wi_hh_shi = 0
Dim st_wi_vol_sbi : st_wi_vol_sbi = 0
Dim st_wi_hh_sbi : st_wi_hh_sbi = 0

Dim gt_ha_tm : gt_ha_tm = 0
Dim gt_prod_est_hi : gt_prod_est_hi = 0
Dim gt_prod_est_shi : gt_prod_est_shi = 0
Dim gt_prod_mill_hi : gt_prod_mill_hi = 0
Dim gt_prod_mill_shi : gt_prod_mill_shi = 0
Dim gt_bud_prod_hi : gt_bud_prod_hi = 0
Dim gt_bud_prod_shi : gt_bud_prod_shi = 0
Dim gt_var_shi_kg : gt_var_shi_kg = 0
Dim gt_var_shi_percent : gt_var_shi_percent = 0
Dim gt_restan_hi : gt_restan_hi = 0
Dim gt_prod_mill_sbi : gt_prod_mill_sbi = 0
Dim gt_bud_prod_sbi : gt_bud_prod_sbi = 0
Dim gt_var_sbi_kg : gt_var_sbi_kg = 0
Dim gt_var_sbi_percent : gt_var_sbi_percent = 0
Dim gt_pusingan_avg : gt_pusingan_avg = 0
Dim gt_pusingan_uptoten : gt_pusingan_uptoten = 0
Dim gt_mandays_hi : gt_mandays_hi = 0
Dim gt_mandays_shi : gt_mandays_shi = 0
Dim gt_mandays_sb_hi : gt_mandays_sb_hi = 0
Dim gt_mandays_sb_shi : gt_mandays_sb_shi = 0
Dim gt_percent_sb_hi : gt_percent_sb_hi = 0
Dim gt_percent_sb_shi : gt_percent_sb_shi = 0
Dim gt_wi_vol_hi : gt_wi_vol_hi = 0
Dim gt_wi_vol_shi : gt_wi_vol_shi = 0
Dim gt_wi_hh_shi : gt_wi_hh_shi = 0
Dim gt_wi_vol_sbi : gt_wi_vol_sbi = 0
Dim gt_wi_hh_sbi : gt_wi_hh_sbi = 0


if rsLHPK.BOF and rsLHPK.EOF then
%>
					<tr>
						<td colspan="<%=29%>">NO DATA AVAILABLE</td>
					</tr>
<%
else
		SortNumber = 1
		while not rsLHPK.EOF
%>
			<%If Ucase(Left(rsLHPK("GroupRegion"), 5)) = Ucase(Cstr("Total")) Then%>
				<tr bgcolor="#ffff01">
					<td align="Center" colspan="2">
						<font face="Verdana" size="1"><strong><%=rsLHPK("GroupRegion")%></strong></font>
					</td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("HA_TM"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("PROD_EST_HI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("PROD_EST_SHI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("PROD_MILL_HI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("PROD_MILL_SHI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("BUD_PROD_HI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("BUD_PROD_SHI"), 0)%></strong></td>
					<%Row_var_shi_kg = Cdbl(rsLHPK("PROD_MILL_SHI")) - Cdbl(rsLHPK("BUD_PROD_SHI"))%>
					<td align="right"><strong><%=FormatNumber(Row_var_shi_kg, 0)%></strong></td>
					<%
					If Row_var_shi_kg = 0 then
						Row_var_shi_percent = 0
					Else
						If Cdbl(rsLHPK("BUD_PROD_SHI")) = 0 Then
							Row_var_shi_percent = 100
						Else
							Row_var_shi_percent = Round((Row_var_shi_kg/ Cdbl(rsLHPK("BUD_PROD_SHI")))*100, 2)
						End If
					End If
					%>
					<td align="right"><strong><%=FormatNumber(Row_var_shi_percent, 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("RESTAN_HI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("PROD_MILL_SBI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("BUD_PROD_SBI"), 0)%></strong></td>
					<%Row_var_sbi_kg = Cdbl(rsLHPK("PROD_MILL_SBI")) - Cdbl(rsLHPK("BUD_PROD_SBI"))%>
					<td align="right"><strong><%=FormatNumber(Row_var_sbi_kg, 0)%></strong></td>
					<%
					If Row_var_sbi_kg = 0 then
						Row_var_sbi_percent = 0
					Else
						If Cdbl(rsLHPK("BUD_PROD_SBI")) = 0 Then
							Row_var_sbi_percent = 100
						Else
							Row_var_sbi_percent = Round((Row_var_sbi_kg/ Cdbl(rsLHPK("BUD_PROD_SBI")))*100, 2)
						End If
					End If
					%>
					<td align="right"><strong><%=FormatNumber(Row_var_sbi_percent, 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("PUSINGAN_AVG"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("PUSINGAN_UPTOTEN"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("MANDAYS_HI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("MANDAYS_SHI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("MANDAYS_SB_HI"), 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(rsLHPK("MANDAYS_SB_SHI"), 0)%></strong></td>
					<%
					'HI
					If Cdbl(rsLHPK("MANDAYS_SB_HI")) = 0 Then
						Row_percent_sb_hi = 0
					Else
						If Cdbl(rsLHPK("MANDAYS_HI")) = 0 Then
							Row_percent_sb_hi = 100
						Else
							Row_percent_sb_hi = Round(( Cdbl(rsLHPK("MANDAYS_SB_HI"))/Cdbl(rsLHPK("MANDAYS_HI")) )*100, 2)
						End If
					End If
					'SHI
					If Cdbl(rsLHPK("MANDAYS_SB_SHI")) = 0 Then
						Row_percent_sb_shi = 0
					Else
						If Cdbl(rsLHPK("MANDAYS_SHI")) = 0 Then
							Row_percent_sb_shi = 100
						Else
							Row_percent_sb_shi = Round(( Cdbl(rsLHPK("MANDAYS_SB_SHI"))/Cdbl(rsLHPK("MANDAYS_SHI")) )*100, 2)
						End If
					End If
					%>
					<td align="right"><strong><%=FormatNumber(Row_percent_sb_hi, 0)%></strong></td>
					<td align="right"><strong><%=FormatNumber(Row_percent_sb_shi, 0)%></strong></td>
					<td align="right" class="grey" colspan="5"><strong>&nbsp;</strong></td>
				</tr>
				<tr>
					<td align="Center" colspan="30" style="padding-left:10px;"  class="separator">
					</td>
				</tr>
			<%Else%>
				<tr>
				<%
					strCountRow = " select count(*) CountRow" & _
					              " from mgt_daily_rpt_est_tab" & _
								  " where to_char(ondate, 'yyyymmdd') = to_char(to_date('" & v_txtdatefrom & "', 'mm/dd/yyyy'), 'yyyymmdd')" & _
		                          "       " & whereCons & "" & _
								  "       and region = '" & rsLHPK("Region") & "'"
					Set rsCountRow = Server.CreateObject("ADODB.Recordset")
					rsCountRow.Open strCountRow ,connstr, adOpenStatic
					'response.write CDbl(rsCountRow("CountRow")) & ". " & strCountRow & " ; "
				%>
					<%If SortNumber = 1 then%>
						<td align="center" Rowspan="<%=CDbl(rsCountRow("CountRow"))+1 %>"><%=rsLHPK("Region")%></td>
					<%End If%>
					
					<td align="left"><%=rsLHPK("fcba")%></td>
					<td align="right"><%=FormatNumber(rsLHPK("HA_TM"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("PROD_EST_HI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("PROD_EST_SHI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("PROD_MILL_HI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("PROD_MILL_SHI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("BUD_PROD_HI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("BUD_PROD_SHI"), 0)%></td>
					<%Row_var_shi_kg = Cdbl(rsLHPK("PROD_MILL_SHI")) - Cdbl(rsLHPK("BUD_PROD_SHI"))%>
					<td align="right"><%=FormatNumber(Row_var_shi_kg, 0)%></td>
					<%
					If Row_var_shi_kg = 0 then
						Row_var_shi_percent = 0
					Else
						If Cdbl(rsLHPK("BUD_PROD_SHI")) = 0 Then
							Row_var_shi_percent = 100
						Else
							Row_var_shi_percent = Round((Row_var_shi_kg/ Cdbl(rsLHPK("BUD_PROD_SHI")))*100, 2)
						End If
					End If
					%>
					<td align="right"><%=FormatNumber(Row_var_shi_percent, 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("RESTAN_HI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("PROD_MILL_SBI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("BUD_PROD_SBI"), 0)%></td>
					<%Row_var_sbi_kg = Cdbl(rsLHPK("PROD_MILL_SBI")) - Cdbl(rsLHPK("BUD_PROD_SBI"))%>
					<td align="right"><%=FormatNumber(Row_var_sbi_kg, 0)%></td>
					<%
					If Row_var_sbi_kg = 0 then
						Row_var_sbi_percent = 0
					Else
						If Cdbl(rsLHPK("BUD_PROD_SBI")) = 0 Then
							Row_var_sbi_percent = 100
						Else
							Row_var_sbi_percent = Round((Row_var_sbi_kg/ Cdbl(rsLHPK("BUD_PROD_SBI")))*100, 2)
						End If
					End If
					%>
					<td align="right"><%=FormatNumber(Row_var_sbi_percent, 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("PUSINGAN_AVG"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("PUSINGAN_UPTOTEN"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("MANDAYS_HI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("MANDAYS_SHI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("MANDAYS_SB_HI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("MANDAYS_SB_SHI"), 0)%></td>
					<%
					'HI
					If Cdbl(rsLHPK("MANDAYS_SB_HI")) = 0 Then
						Row_percent_sb_hi = 0
					Else
						If Cdbl(rsLHPK("MANDAYS_HI")) = 0 Then
							Row_percent_sb_hi = 100
						Else
							Row_percent_sb_hi = Round(( Cdbl(rsLHPK("MANDAYS_SB_HI"))/Cdbl(rsLHPK("MANDAYS_HI")) )*100, 2)
						End If
					End If
					'SHI
					If Cdbl(rsLHPK("MANDAYS_SB_SHI")) = 0 Then
						Row_percent_sb_shi = 0
					Else
						If Cdbl(rsLHPK("MANDAYS_SHI")) = 0 Then
							Row_percent_sb_shi = 100
						Else
							Row_percent_sb_shi = Round(( Cdbl(rsLHPK("MANDAYS_SB_SHI"))/Cdbl(rsLHPK("MANDAYS_SHI")) )*100, 2)
						End If
					End If
					%>
					<td align="right"><%=FormatNumber(Row_percent_sb_hi, 0)%></td>
					<td align="right"><%=FormatNumber(Row_percent_sb_shi, 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("WI_VOL_HI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("WI_VOL_SHI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("WI_HH_SHI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("WI_VOL_SBI"), 0)%></td>
					<td align="right"><%=FormatNumber(rsLHPK("WI_HH_SBI"), 0)%></td>
				</tr>
				<%
				st_ha_tm = st_ha_tm + Cdbl(rsLHPK("HA_TM"))
				st_prod_est_hi = st_prod_est_hi + Cdbl(rsLHPK("PROD_EST_HI"))
				st_prod_est_shi = st_prod_est_shi + Cdbl(rsLHPK("PROD_EST_SHI"))
				st_prod_mill_hi = st_prod_mill_hi + Cdbl(rsLHPK("PROD_MILL_HI"))
				st_prod_mill_shi = st_prod_mill_shi + Cdbl(rsLHPK("PROD_MILL_SHI"))
				st_bud_prod_hi = st_bud_prod_hi + Cdbl(rsLHPK("BUD_PROD_HI"))
				st_bud_prod_shi = st_bud_prod_shi + Cdbl(rsLHPK("BUD_PROD_SHI"))
				st_var_shi_kg = st_var_shi_kg + Cdbl(rsLHPK("VAR_SHI_KG"))
				st_var_shi_percent = st_var_shi_percent + Cdbl(rsLHPK("VAR_SHI_PERCENT"))
				st_restan_hi = st_restan_hi + Cdbl(rsLHPK("RESTAN_HI"))
				st_prod_mill_sbi = st_prod_mill_sbi + Cdbl(rsLHPK("PROD_MILL_SBI"))
				st_bud_prod_sbi = st_bud_prod_sbi + Cdbl(rsLHPK("BUD_PROD_SBI"))
				st_var_sbi_kg = st_var_sbi_kg + Cdbl(rsLHPK("VAR_SBI_KG"))
				st_var_sbi_percent = st_var_sbi_percent + Cdbl(rsLHPK("VAR_SBI_PERCENT"))
				st_pusingan_avg = st_pusingan_avg + Cdbl(rsLHPK("PUSINGAN_AVG"))
				st_pusingan_uptoten = st_pusingan_uptoten + Cdbl(rsLHPK("PUSINGAN_UPTOTEN"))
				st_mandays_hi = st_mandays_hi + Cdbl(rsLHPK("MANDAYS_HI"))
				st_mandays_shi = st_mandays_shi + Cdbl(rsLHPK("MANDAYS_SHI"))
				st_mandays_sb_hi = st_mandays_sb_hi + Cdbl(rsLHPK("MANDAYS_SB_HI"))
				st_mandays_sb_shi = st_mandays_sb_shi + Cdbl(rsLHPK("MANDAYS_SB_SHI"))
				st_percent_sb_hi = st_percent_sb_hi + Cdbl(rsLHPK("PERCENT_SB_HI"))
				st_percent_sb_shi = st_percent_sb_shi + Cdbl(rsLHPK("PERCENT_SB_SHI"))
				st_wi_vol_hi = st_wi_vol_hi + Cdbl(rsLHPK("WI_VOL_HI"))
				st_wi_vol_shi = st_wi_vol_shi + Cdbl(rsLHPK("WI_VOL_SHI"))
				st_wi_hh_shi = st_wi_hh_shi + Cdbl(rsLHPK("WI_HH_SHI"))
				st_wi_vol_sbi = st_wi_vol_sbi + Cdbl(rsLHPK("WI_VOL_SBI"))
				st_wi_hh_sbi = st_wi_hh_sbi + Cdbl(rsLHPK("WI_HH_SBI"))
				
				if SortNumber = CDbl(rsCountRow("CountRow"))then
					SortNumber = 1
				%>
					<tr bgcolor="#F4F4A0">
						<td align="Left" style="padding-left:10px;">
							<font face="Verdana" size="1"><strong>Sub Total</strong></font>
						</td>
						<td align="right"><strong><%=FormatNumber(st_ha_tm, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_prod_est_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_prod_est_shi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_prod_mill_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_prod_mill_shi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_bud_prod_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_bud_prod_shi, 0)%></strong></td>
						<%' Var SHI Kg
						st_var_shi_kg = st_prod_mill_shi - st_bud_prod_shi
						  ' Var SHI %
						If st_var_shi_kg = 0 Then
							st_var_shi_percent = 0
						Else
							If st_bud_prod_shi = 0 Then
								st_var_shi_percent = 100
							Else
								st_var_shi_percent = Round((st_var_shi_kg/st_bud_prod_shi)*100, 2)
							End If
						End If
						%>
						<td align="right"><strong><%=FormatNumber(st_var_shi_kg, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_var_shi_percent, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_restan_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_prod_mill_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_bud_prod_sbi, 0)%></strong></td>
						<%' Var SBI Kg
						st_var_sbi_kg = st_prod_mill_sbi - st_bud_prod_sbi
						  ' Var SBI %
						If st_var_sbi_kg = 0 Then
							st_var_sbi_percent = 0
						Else
							If st_bud_prod_sbi = 0 Then
								st_var_sbi_percent = 100
							Else
								st_var_sbi_percent = Round((st_var_sbi_kg/st_bud_prod_sbi)*100, 2)
							End If
						End If
						%>
						<td align="right"><strong><%=FormatNumber(st_var_sbi_kg, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_var_sbi_percent, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_pusingan_avg, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_pusingan_uptoten, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_mandays_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_mandays_shi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_mandays_sb_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_mandays_sb_shi, 0)%></strong></td>
						<%
						'HI
						If st_mandays_sb_hi = 0 Then
							st_percent_sb_hi = 0
						Else
							If st_mandays_hi = 0 Then
								st_percent_sb_hi = 100
							Else
								st_percent_sb_hi = Round((st_mandays_sb_hi/st_mandays_hi)*100, 2)
							End If
						End If
						'SHI
						If st_mandays_sb_shi = 0 Then
							st_percent_sb_shi = 0
						Else
							If st_mandays_shi = 0 Then
								st_percent_sb_shi = 100
							Else
								st_percent_sb_shi = Round((st_mandays_sb_shi/st_mandays_shi)*100, 2)
							End If
						End If
						%>
						<td align="right"><strong><%=FormatNumber(st_percent_sb_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(st_percent_sb_shi, 0)%></strong></td>
						<td align="right" class="grey" colspan="5"><strong>&nbsp;</strong></td>
					</tr>
				<%
					st_ha_tm = 0
					st_prod_est_hi = 0
					st_prod_est_shi = 0
					st_prod_mill_hi = 0
					st_prod_mill_shi = 0
					st_bud_prod_hi = 0
					st_bud_prod_shi = 0
					st_var_shi_kg = 0
					st_var_shi_percent = 0
					st_restan_hi = 0
					st_prod_mill_sbi = 0
					st_bud_prod_sbi = 0
					st_var_sbi_kg = 0
					st_var_sbi_percent = 0
					st_pusingan_avg = 0
					st_pusingan_uptoten = 0
					st_mandays_hi = 0
					st_mandays_shi = 0
					st_mandays_sb_hi = 0
					st_mandays_sb_shi = 0
					st_percent_sb_hi = 0
					st_percent_sb_shi = 0
					st_wi_vol_hi = 0
					st_wi_vol_shi = 0
					st_wi_hh_shi = 0
					st_wi_vol_sbi = 0
					st_wi_hh_sbi = 0
					
					
					
					
				else 
					SortNumber = SortNumber + 1
				end if
				    gt_ha_tm = gt_ha_tm + Cdbl(rsLHPK("HA_TM"))
				    gt_prod_est_hi = gt_prod_est_hi + Cdbl(rsLHPK("PROD_EST_HI"))
					gt_prod_est_shi = gt_prod_est_shi + Cdbl(rsLHPK("PROD_EST_SHI"))
					gt_prod_mill_hi = gt_prod_mill_hi + Cdbl(rsLHPK("PROD_MILL_HI"))
					gt_prod_mill_shi = gt_prod_mill_shi + Cdbl(rsLHPK("PROD_MILL_SHI"))
					gt_bud_prod_hi = gt_bud_prod_hi + Cdbl(rsLHPK("BUD_PROD_HI"))
					gt_bud_prod_shi = gt_bud_prod_shi + Cdbl(rsLHPK("BUD_PROD_SHI"))
					gt_var_shi_kg = gt_var_shi_kg + Cdbl(rsLHPK("VAR_SHI_KG"))
					gt_var_shi_percent = gt_var_shi_percent + Cdbl(rsLHPK("VAR_SHI_PERCENT"))
					gt_restan_hi = gt_restan_hi + Cdbl(rsLHPK("RESTAN_HI"))
					gt_prod_mill_sbi = gt_prod_mill_sbi + Cdbl(rsLHPK("PROD_MILL_SBI"))
					gt_bud_prod_sbi = gt_bud_prod_sbi + Cdbl(rsLHPK("BUD_PROD_SBI"))
					gt_var_sbi_kg = gt_var_sbi_kg + Cdbl(rsLHPK("VAR_SBI_KG"))
					gt_var_sbi_percent = gt_var_sbi_percent + Cdbl(rsLHPK("VAR_SBI_PERCENT"))
					gt_pusingan_avg = gt_pusingan_avg + Cdbl(rsLHPK("PUSINGAN_AVG"))
					gt_pusingan_uptoten = gt_pusingan_uptoten + Cdbl(rsLHPK("PUSINGAN_UPTOTEN"))
					gt_mandays_hi = gt_mandays_hi + Cdbl(rsLHPK("MANDAYS_HI"))
					gt_mandays_shi = gt_mandays_shi + Cdbl(rsLHPK("MANDAYS_SHI"))
					gt_mandays_sb_hi = gt_mandays_sb_hi + Cdbl(rsLHPK("MANDAYS_SB_HI"))
					gt_mandays_sb_shi = gt_mandays_sb_shi + Cdbl(rsLHPK("MANDAYS_SB_SHI"))
					gt_percent_sb_hi = gt_percent_sb_hi + Cdbl(rsLHPK("PERCENT_SB_HI"))
					gt_percent_sb_shi = gt_percent_sb_shi + Cdbl(rsLHPK("PERCENT_SB_SHI"))
					gt_wi_vol_hi = gt_wi_vol_hi + Cdbl(rsLHPK("WI_VOL_HI"))
					gt_wi_vol_shi = gt_wi_vol_shi + Cdbl(rsLHPK("WI_VOL_SHI"))
					gt_wi_hh_shi = gt_wi_hh_shi + Cdbl(rsLHPK("WI_HH_SHI"))
					gt_wi_vol_sbi = gt_wi_vol_sbi + Cdbl(rsLHPK("WI_VOL_SBI"))
					gt_wi_hh_sbi = gt_wi_hh_sbi + Cdbl(rsLHPK("WI_HH_SBI"))
				%>

			<%End If%>
<%
				rsLHPK.MoveNext
				
				PrevRegion = CurrRegion
				PrevGroupRegion = CurrGroupRegion
			Wend	
rsLHPK.Close
set rsLHPK = Nothing
%>

					<tr bgcolor="#ffff01">
						<td align="Center" colspan="2">
							<font face="Verdana" size="1"><strong>Grand Total</strong></font>
						</td>
						<td align="right"><strong><%=FormatNumber(gt_ha_tm, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_prod_est_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_prod_est_shi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_prod_mill_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_prod_mill_shi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_bud_prod_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_bud_prod_shi, 0)%></strong></td>
						<%' Var SHI Kg
						gt_var_shi_kg = gt_prod_mill_shi - gt_bud_prod_shi
						  ' Var SHI %
						If gt_var_shi_kg = 0 Then
							gt_var_shi_percent = 0
						Else
							If gt_bud_prod_shi = 0 Then
								gt_var_shi_percent = 100
							Else
								gt_var_shi_percent = Round((gt_var_shi_kg/gt_bud_prod_shi)*100, 2)
							End If
						End If
						%>
						<td align="right"><strong><%=FormatNumber(gt_var_shi_kg, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_var_shi_percent, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_restan_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_prod_mill_sbi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_bud_prod_sbi, 0)%></strong></td>
						<%' Var SBI Kg
						gt_var_sbi_kg = gt_prod_mill_sbi - gt_bud_prod_sbi
						  ' Var SBI %
						If gt_var_sbi_kg = 0 Then
							gt_var_sbi_percent = 0
						Else
							If gt_bud_prod_sbi = 0 Then
								gt_var_sbi_percent = 100
							Else
								gt_var_sbi_percent = Round((gt_var_sbi_kg/gt_bud_prod_sbi)*100, 2)
							End If
						End If
						%>
						<td align="right"><strong><%=FormatNumber(gt_var_sbi_kg, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_var_sbi_percent, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_pusingan_avg, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_pusingan_uptoten, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_mandays_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_mandays_shi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_mandays_sb_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_mandays_sb_shi, 0)%></strong></td>
						<%
						'HI
						If gt_mandays_sb_hi = 0 Then
							gt_percent_sb_hi = 0
						Else
							If gt_mandays_hi = 0 Then
								gt_percent_sb_hi = 100
							Else
								gt_percent_sb_hi = Round((gt_mandays_sb_hi/gt_mandays_hi)*100, 2)
							End If
						End If
						'SHI
						If gt_mandays_sb_shi = 0 Then
							gt_percent_sb_shi = 0
						Else
							If gt_mandays_shi = 0 Then
								gt_percent_sb_shi = 100
							Else
								gt_percent_sb_shi = Round((gt_mandays_sb_shi/gt_mandays_shi)*100, 2)
							End If
						End If
						%>
						<td align="right"><strong><%=FormatNumber(gt_percent_sb_hi, 0)%></strong></td>
						<td align="right"><strong><%=FormatNumber(gt_percent_sb_shi, 0)%></strong></td>
						<td align="right" class="grey" colspan="5"><strong>&nbsp;</strong></td>
					</tr>
<%End if%>

				</tbody>
			</table>

		</div>
	</body>
</html>