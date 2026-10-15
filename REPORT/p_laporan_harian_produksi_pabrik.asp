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

sqlLHPP = " select * from (" & _
          " select case when region is not null and fcba is null and mill_cap is null then 'Total '||region" & _
          "             when region is null and fcba is null and mill_cap is null then 'Grand Total'" & _
          "             else region" & _
          "        end nextRegion," & _
          "        case when region is not null and fcba is null and mill_cap is null then 1000+Sequence" & _
          "             when region is null and fcba is null and mill_cap is null then 10000+Sequence" & _
          "             else Sequence" & _
          "        end nextSequence" & _
          "          , aa.*        " & _
          " from (" & _
          "   SELECT region," & _
          "          orderdataregion," & _
          "          fcba," & _
          "          mill_cap," & _
          "          sequence," & _
          "          reference," & _
          "          SUM (FFB_INTI) FFB_INTI," & _
          "          SUM (FFB_PLASMA) FFB_PLASMA," & _
          "          SUM (FFB_T_GILING) FFB_T_GILING," & _
          "          SUM (FFB_OUTSIDER) FFB_OUTSIDER," & _
          "          SUM (FFB_TOTAL) FFB_TOTAL," & _
          "          SUM (FFB_BUDGET) FFB_BUDGET," & _
          "          SUM (FFB_VAR_MT) FFB_VAR_MT," & _
          "          SUM (FFB_VAR_PERCENT) FFB_VAR_PERCENT," & _
          "          SUM (FFB_PROCESS) FFB_PROCESS," & _
          "          SUM (FFB_UNPROCESS) FFB_UNPROCESS," & _
          "          SUM (THROUGHPUT) THROUGHPUT," & _
          "          SUM (CPO_ACT) CPO_ACT," & _
          "          SUM (CPO_BUD) CPO_BUD," & _
          "          SUM (CPO_VAR_MT) CPO_VAR_MT," & _
          "          SUM (CPO_VAR_PERCENT) CPO_VAR_PERCENT," & _
          "          SUM (OER_ACT) OER_ACT," & _
          "          SUM (OER_BUD) OER_BUD," & _
          "          SUM (OER_VAR_PERCENT) OER_VAR_PERCENT," & _
          "          SUM (FFA_ACT) FFA_ACT," & _
          "          SUM (FFA_BUD) FFA_BUD," & _
          "          SUM (FFA_VAR_PERCENT) FFA_VAR_PERCENT," & _
          "          SUM (KERNEL_ACT) KERNEL_ACT," & _
          "          SUM (KERNEL_BUD) KERNEL_BUD," & _
          "          SUM (KERNEL_VAR_MT) KERNEL_VAR_MT," & _
          "          SUM (KERNEL_VAR_PERCENT) KERNEL_VAR_PERCENT," & _
          "          SUM (KER_ACT) KER_ACT," & _
          "          SUM (KER_BUD) KER_BUD," & _
          "          SUM (KER_VAR_PERCENT) KER_VAR_PERCENT," & _
          "          SUM (QUALITY_MOIST) QUALITY_MOIST," & _
          "          SUM (QUALITY_DIRT) QUALITY_DIRT" & _
          "     FROM (  SELECT case when BUSINESSUNIT_API.get_CompanyCode(fcba) = 'PT.THIP' then 'THIP'" & _
          "                         when BUSINESSUNIT_API.get_CompanyCode(fcba) = 'PT.CRS' then 'CRS'" & _
          "                         else 'AMS Ganda'" & _
          "                    end region," & _
          "                    case when BUSINESSUNIT_API.get_CompanyCode(fcba) = 'PT.THIP' then 1" & _
          "                         when BUSINESSUNIT_API.get_CompanyCode(fcba) = 'PT.CRS' then 3" & _
          "                         else 2" & _
          "                    end OrderDataRegion," & _
          "                    fcba," & _
          "                    mill_cap," & _
          "                    sequence," & _
          "                    reference," & _
          "                    SUM (FFB_INTI) FFB_INTI," & _
          "                    SUM (FFB_PLASMA) FFB_PLASMA," & _
          "                    SUM (FFB_T_GILING) FFB_T_GILING," & _
          "                    SUM (FFB_OUTSIDER) FFB_OUTSIDER," & _
          "                    SUM (FFB_TOTAL) FFB_TOTAL," & _
          "                    SUM (FFB_BUDGET) FFB_BUDGET," & _
          "                    SUM (FFB_VAR_MT) FFB_VAR_MT," & _
          "                    SUM (FFB_VAR_PERCENT) FFB_VAR_PERCENT," & _
          "                    SUM (FFB_PROCESS) FFB_PROCESS," & _
          "                    SUM (FFB_UNPROCESS) FFB_UNPROCESS," & _
          "                    SUM (THROUGHPUT) THROUGHPUT," & _
          "                    SUM (CPO_ACT) CPO_ACT," & _
          "                    SUM (CPO_BUD) CPO_BUD," & _
          "                    SUM (CPO_VAR_MT) CPO_VAR_MT," & _
          "                    SUM (CPO_VAR_PERCENT) CPO_VAR_PERCENT," & _
          "                    SUM (OER_ACT) OER_ACT," & _
          "                    SUM (OER_BUD) OER_BUD," & _
          "                    SUM (OER_VAR_PERCENT) OER_VAR_PERCENT," & _
          "                    SUM (FFA_ACT) FFA_ACT," & _
          "                    SUM (FFA_BUD) FFA_BUD," & _
          "                    SUM (FFA_VAR_PERCENT) FFA_VAR_PERCENT," & _
          "                    SUM (KERNEL_ACT) KERNEL_ACT," & _
          "                    SUM (KERNEL_BUD) KERNEL_BUD," & _
          "                    SUM (KERNEL_VAR_MT) KERNEL_VAR_MT," & _
          "                    SUM (KERNEL_VAR_PERCENT) KERNEL_VAR_PERCENT," & _
          "                    SUM (KER_ACT) KER_ACT," & _
          "                    SUM (KER_BUD) KER_BUD," & _
          "                    SUM (KER_VAR_PERCENT) KER_VAR_PERCENT," & _
          "                    SUM (QUALITY_MOIST) QUALITY_MOIST," & _
          "                    SUM (QUALITY_DIRT) QUALITY_DIRT" & _
          "             FROM mgt_daily_rpt_mill_tab" & _
          "             WHERE to_char(ondate, 'yyyymmdd') = to_char(to_date('" & v_txtdatefrom & "', 'mm/dd/yyyy'), 'yyyymmdd')" & _
          "                   " & whereCons & "" & _
          "           GROUP BY fcba," & _
          "                    mill_cap," & _
          "                    sequence," & _
          "                    reference)" & _
          " GROUP BY GROUPING SETS ( (region," & _
          "                           orderdataregion," & _
          "                           fcba," & _
          "                           mill_cap," & _
          "                           sequence," & _
          "                           reference)," & _
          "                           (region," & _
          "                           orderdataregion," & _
          "                           sequence," & _
          "                           reference)," & _
          "                           (sequence," & _
          "                           reference) )" & _
          " ) aa" & _
          " ) order by orderdataregion, BUSINESSUNIT_API.get_OrderData (fcba), nextsequence"

'response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Harian Produksi Pabrik</title>
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
				<!--#include file="FormSelectNewSumM.html"-->
				
					<table id="form-table" class="centered-margin">
						<tr>
							<td><font face="verdana" size="1">BUDGET YEAR</font></td>
							<td><font face="verdana" size="1">BUDGET VERSION</font></td>
							<td><font face="verdana" size="1">MODULE VERSION</font></td>
							<td><font face="verdana" size="1">DATE</font></td>
							<td>&nbsp;</td>
						</tr>
						<tr>
							<td align="center">
								<input type="text" name="txtBudgetYear" value="<%=v_txtBudgetYear%>">
							</td>
							<td align="center">
								<input type="text" name="txtBudgetVersion" value="<%=v_txtBudgetVersion%>">
							</td>
							<td align="center">
								<input type="text" name="txtModuleVersion" value="<%=v_txtModuleVersion%>">
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
						<th colspan="11">FFB PROCESSED (MT)</th>
						<th rowspan="3">THROUGHPUT</th>
						<th colspan="4">CPO PRODUCTION (MT)</th>
						<th colspan="3">OER (%)</th>
						<th colspan="3">FFA (%)</th>
						<th colspan="4">KERNEL PRODUCTION (MT)</th>
						<th colspan="3">KER (%)</th>
						<th colspan="2">QUALITY (%)</th>
					</tr>
					<tr>
						<th colspan="5">RECEIVED</th>
						<th rowspan="2">TOTAL RECEIVED</th>
						<th rowspan="2">BGT</th>
						<th colspan="2">VARIANCE</th>
						<th rowspan="2">TOTAL PROCESS</th>
						<th rowspan="2">UN PROCESS</th>
						<th rowspan="2">ACT</th>
						<th rowspan="2">BGT</th>
						<th colspan="2">VARIANCE</th>
						<th rowspan="2">ACT</th>
						<th rowspan="2">BGT</th>
						<th rowspan="2">VAR (%)</th>
						<th rowspan="2">ACT</th>
						<th rowspan="2">BGT</th>
						<th rowspan="2">VAR (%)</th>
						<th rowspan="2">ACT</th>
						<th rowspan="2">BGT</th>
						<th colspan="2">VARIANCE</th>
						<th rowspan="2">ACT</th>
						<th rowspan="2">BGT</th>
						<th rowspan="2">VAR (%)</th>
						<th rowspan="2">MOIST CONTENT (7%)</th>
						<th rowspan="2">DIRT CONTENT (7%)</th>
					</tr>
					<tr>
						<th></th>
						<th>INTI</th>
						<th>PLASMA</th>
						<th>TUMPANG GILING</th>
						<th>OUTSIDE</th>
						<th>MT</th>
						<th>%</th>
						<th>MT</th>
						<th>%</th>
						<th>MT</th>
						<th>%</th>
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


if rsLHPP.BOF and rsLHPP.EOF then
%>
					<tr>
						<td colspan="<%=33%>">NO DATA AVAILABLE</td>
					</tr>
<%
else
		while not rsLHPP.eof
			currfcba = rsLHPP("FCBA")
			currnextregion =  rsLHPP("NEXTREGION")
%>

					
					<%If (rsLHPP("NEXTREGION") = "Grand Total" Or Left(Cstr(rsLHPP("NEXTREGION")), 5) = "Total") Then%>
					<tr bgcolor="#F4F4A0">
						<%If currnextregion <> prevnextregion then%>
							<td align="Center" colspan="2" style=" border-bottom:none;">
								<strong><%=rsLHPP("NEXTREGION")%></strong>
							</td>
							<%prevnextregion = currnextregion%>
						<%Else%>
							<td align="center" colspan="2" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
						<%End If%>
						<td align="Center"><strong><%=rsLHPP("REFERENCE")%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_INTI"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_PLASMA"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_T_GILING"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_OUTSIDER"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_TOTAL"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_BUDGET"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_VAR_MT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_VAR_PERCENT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_PROCESS"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFB_UNPROCESS"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("THROUGHPUT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("CPO_ACT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("CPO_BUD"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("CPO_VAR_MT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("CPO_VAR_PERCENT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("OER_ACT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("OER_BUD"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("OER_VAR_PERCENT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFA_ACT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFA_BUD"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("FFA_VAR_PERCENT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KERNEL_ACT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KERNEL_BUD"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KERNEL_VAR_MT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KERNEL_VAR_PERCENT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KER_ACT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KER_BUD"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("KER_VAR_PERCENT"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("QUALITY_MOIST"), 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(rsLHPP("QUALITY_DIRT"), 2)%></strong></td>
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
						<td align="Center"><%=rsLHPP("REFERENCE")%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_INTI"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_PLASMA"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_T_GILING"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_OUTSIDER"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_TOTAL"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_BUDGET"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_VAR_MT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_VAR_PERCENT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_PROCESS"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFB_UNPROCESS"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("THROUGHPUT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("CPO_ACT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("CPO_BUD"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("CPO_VAR_MT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("CPO_VAR_PERCENT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("OER_ACT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("OER_BUD"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("OER_VAR_PERCENT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_ACT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_BUD"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_VAR_PERCENT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KERNEL_ACT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KERNEL_BUD"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KERNEL_VAR_MT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KERNEL_VAR_PERCENT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_ACT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_BUD"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_VAR_PERCENT"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("QUALITY_MOIST"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("QUALITY_DIRT"), 2)%></td>
					</tr>
					<%End If%>


<%				
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
