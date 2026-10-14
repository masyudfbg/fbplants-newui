<%@Language=VBScript%>
<%Option Explicit%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtMonth, v_txtYear
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
	
	v_txtYear		= Request("txtYear")
	v_txtMonth		= Request("txtMonth")
	
	If Len(Trim(v_txtYear)) = 0 Then v_txtYear = Year(Session("v_Costbook"))
	If Len(Trim(v_txtMonth)) = 0 Then v_txtMonth = Month(Session("v_Costbook"))
	
	If Len(Trim(v_txtBudgetYear)) = 0 Then
		v_txtBudgetYear = Year(textDate(Date))
	Else
		v_txtBudgetYear = v_txtYear
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
					
wherecons	=	" and businessunit_api.get_fctype(fcba) = 'M' AND USER_COMPANY_API.IsAuthorizedFcba('" & Session("CodeName") & "', fcba) =  'TRUE'"

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
		whereCons	=	whereCons	& " And Plantationcode = '" & v_txtPlantation & "'"
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
		whereCons	=	whereCons	& " And GroupRegioncode = '" & v_txtGroupRegion & "'"
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
		whereCons	=	whereCons	& " And Regioncode = '" & v_txtRegion & "'"
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
		whereCons	=	whereCons	& " And companycode = '" & v_txtCompany & "'"
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

' sqlLHPP = " select case when region is not null and fcba is null then 'Total '||region" & _
          ' "             when region is null and fcba is null then 'Grand Total'" & _
          ' "             else region" & _
          ' "        end nextRegion," & _
          ' "        aa.*" & _
          ' " from (" & _
          ' "     select OrderDataRegion," & _
          ' "            region," & _
          ' "            fcba," & _
		  ' "            SUM(FFBPROCESS_THISMONTH) FFBPROCESS_THISMONTH," & _
		  ' "            SUM(FFBPROCESS_TODATE) FFBPROCESS_TODATE," & _
		  ' "            SUM(CPOPROD_THISMONTH) CPOPROD_THISMONTH," & _
		  ' "            SUM(CPOPROD_TODATE) CPOPROD_TODATE," & _
		  ' "            SUM(KERPROD_THISMONTH) KERPROD_THISMONTH," & _
		  ' "            SUM(KERPROD_TODATE) KERPROD_TODATE," & _
		  ' "            SUM(SHLPROD_THISMONTH) SHLPROD_THISMONTH," & _
		  ' "            SUM(SHLPROD_TODATE) SHLPROD_TODATE," & _
		  ' "            SUM(FFAPROD_THISMONTH) FFAPROD_THISMONTH," & _
		  ' "            SUM(FFAPROD_TODATE) FFAPROD_TODATE," & _
		  ' "            SUM(FFBPROCESS_BUD_TODATE) FFBPROCESS_BUD_TODATE," & _
		  ' "            SUM(FFAPROD_BUD_TODATE) FFAPROD_BUD_TODATE," & _
		  ' "            SUM(CPOPROD_BUD_TODATE) CPOPROD_BUD_TODATE," & _
		  ' "            SUM(KERPROD_BUD_TODATE) KERPROD_BUD_TODATE," & _
		  ' "            SUM(SHLPROD_BUD_TODATE) SHLPROD_BUD_TODATE," & _
          ' "            AVG(OER_ACT_THISMONTH) OER_ACT_THISMONTH," & _
          ' "            AVG(OER_ACT_TODATE) OER_ACT_TODATE," & _
          ' "            AVG(OER_BUD_TODATE) OER_BUD_TODATE," & _
          ' "            AVG(KER_ACT_THISMONTH) KER_ACT_THISMONTH," & _
          ' "            AVG(KER_ACT_TODATE) KER_ACT_TODATE," & _
          ' "            AVG(KER_BUD_TODATE) KER_BUD_TODATE," & _
          ' "            AVG(CER_ACT_THISMONTH) CER_ACT_THISMONTH," & _
		  ' "            AVG(CER_ACT_TODATE) CER_ACT_TODATE," & _
		  ' "            AVG(CER_BUD_TODATE) CER_BUD_TODATE," & _
		  ' "            AVG(FFA_ACT_THISMONTH) FFA_ACT_THISMONTH," & _
		  ' "            AVG(FFA_ACT_TODATE) FFA_ACT_TODATE," & _
		  ' "            AVG(FFA_BUD_TODATE) FFA_BUD_TODATE" & _
          ' "     from ( SELECT case when companycode = 'PT.THIP' then 'A'" & _
          ' "                                           when companycode = 'PT.CRS' then 'B'" & _
          ' "                                           else 'C'" & _
          ' "                                      end OrderDataRegion," & _
          ' "                                      case when companycode = 'PT.THIP' then 'THIP'" & _
          ' "                                           when companycode = 'PT.CRS' then 'CRS'" & _
          ' "                                           else 'AMS Ganda'" & _
          ' "                                      end region," & _
          ' "                                      fcba," & _
          ' "                                      NVL(SUM(FFBPROCESS_THISMONTH), 0) FFBPROCESS_THISMONTH," & _
		  ' "                                      NVL(SUM(FFBPROCESS_TODATE), 0) FFBPROCESS_TODATE," & _
		  ' "                                      NVL(SUM(CPOPROD_THISMONTH), 0) CPOPROD_THISMONTH," & _
		  ' "                                      NVL(SUM(CPOPROD_TODATE), 0) CPOPROD_TODATE," & _
		  ' "                                      NVL(SUM(KERPROD_THISMONTH), 0) KERPROD_THISMONTH," & _
		  ' "                                      NVL(SUM(KERPROD_TODATE), 0) KERPROD_TODATE," & _
		  ' "                                      NVL(SUM(SHLPROD_THISMONTH), 0) SHLPROD_THISMONTH," & _
		  ' "                                      NVL(SUM(SHLPROD_TODATE), 0) SHLPROD_TODATE," & _
		  ' "                                      NVL(SUM(FFAPROD_THISMONTH), 0) FFAPROD_THISMONTH," & _
		  ' "                                      NVL(SUM(FFAPROD_TODATE), 0) FFAPROD_TODATE," & _
		  ' "                                      NVL(SUM(FFBPROCESS_BUD_ANNUAL), 0) FFBPROCESS_BUD_TODATE," & _
		  ' "                                      NVL(SUM(FFAPROD_BUD_TODATE), 0) FFAPROD_BUD_TODATE," & _
		  ' "                                      NVL(SUM(CPOPROD_BUD_TODATE), 0) CPOPROD_BUD_TODATE," & _
		  ' "                                      NVL(SUM(KERPROD_BUD_TODATE), 0) KERPROD_BUD_TODATE," & _
		  ' "                                      NVL(SUM(SHLPROD_BUD_TODATE), 0) SHLPROD_BUD_TODATE," & _
		  ' "                                      AVG(OER_ACT_THISMONTH) OER_ACT_THISMONTH," & _
          ' "                                      AVG(OER_ACT_TODATE) OER_ACT_TODATE," & _
          ' "                                      AVG(OER_BUD_TODATE) OER_BUD_TODATE," & _
          ' "                                      AVG(KER_ACT_THISMONTH) KER_ACT_THISMONTH," & _
          ' "                                      AVG(KER_ACT_TODATE) KER_ACT_TODATE," & _
          ' "                                      AVG(KER_BUD_TODATE) KER_BUD_TODATE," & _
          ' "                                      AVG(CER_ACT_THISMONTH) CER_ACT_THISMONTH," & _
		  ' "                                      AVG(CER_ACT_TODATE) CER_ACT_TODATE," & _
		  ' "                                      AVG(CER_BUD_TODATE) CER_BUD_TODATE," & _
		  ' "                                      AVG(FFA_ACT_THISMONTH) FFA_ACT_THISMONTH," & _
		  ' "                                      AVG(FFA_ACT_TODATE) FFA_ACT_TODATE," & _
		  ' "                                      AVG(FFA_BUD_TODATE) FFA_BUD_TODATE" & _
          ' "                               FROM mgt2_millratio_det_tab" & _
          ' "                               WHERE year = " & v_txtYear & "" & _
          ' "                                     and month = " & v_txtMonth & "" & _
          ' "                                     " & whereCons & "" & _
          ' "                               GROUP BY companycode," & _
          ' "                                        fcba" & _
          ' "     )" & _
          ' "     group by grouping sets( (OrderDataRegion," & _
          ' "                              region," & _
          ' "                              fcba)," & _
          ' "                             (OrderDataRegion," & _
          ' "                              region)," & _
          ' "                             () )" & _
          ' " ) aa" & _
		  ' " order by orderdataregion, BUSINESSUNIT_API.get_OrderData (fcba)"
		  
sqlLHPP = " select case when region is not null and fcba is null then 'Total '||region" & _
          "             when region is null and fcba is null then 'Grand Total'" & _
          "             else region" & _
          "        end nextRegion," & _
          "        aa.*" & _
          " from (" & _
          "     select OrderDataRegion," & _
          "            region," & _
          "            fcba," & _
		  "            SUM(FFBPROCESS_THISMONTH) FFBPROCESS_THISMONTH," & _
		  "            SUM(FFBPROCESS_TODATE) FFBPROCESS_TODATE," & _
		  "            SUM(CPOPROD_THISMONTH) CPOPROD_THISMONTH," & _
		  "            SUM(CPOPROD_TODATE) CPOPROD_TODATE," & _
		  "            SUM(KERPROD_THISMONTH) KERPROD_THISMONTH," & _
		  "            SUM(KERPROD_TODATE) KERPROD_TODATE," & _
		  "            SUM(SHLPROD_THISMONTH) SHLPROD_THISMONTH," & _
		  "            SUM(SHLPROD_TODATE) SHLPROD_TODATE," & _
		  "            SUM(FFAPROD_THISMONTH) FFAPROD_THISMONTH," & _
		  "            SUM(FFAPROD_TODATE) FFAPROD_TODATE," & _
		  "            SUM(FFBPROCESS_BUD_ANNUAL) FFBPROCESS_BUD_ANNUAL," & _
		  "            SUM(FFAPROD_BUD_ANNUAL) FFAPROD_BUD_ANNUAL," & _
		  "            SUM(CPOPROD_BUD_ANNUAL) CPOPROD_BUD_ANNUAL," & _
		  "            SUM(KERPROD_BUD_ANNUAL) KERPROD_BUD_ANNUAL," & _
		  "            SUM(SHLPROD_BUD_ANNUAL) SHLPROD_BUD_ANNUAL," & _
          "            AVG(OER_ACT_THISMONTH) OER_ACT_THISMONTH," & _
          "            AVG(OER_ACT_TODATE) OER_ACT_TODATE," & _
          "            AVG(OER_BUD_ANNUAL) OER_BUD_ANNUAL," & _
          "            AVG(KER_ACT_THISMONTH) KER_ACT_THISMONTH," & _
          "            AVG(KER_ACT_TODATE) KER_ACT_TODATE," & _
          "            AVG(KER_BUD_ANNUAL) KER_BUD_ANNUAL," & _
          "            AVG(CER_ACT_THISMONTH) CER_ACT_THISMONTH," & _
		  "            AVG(CER_ACT_TODATE) CER_ACT_TODATE," & _
		  "            AVG(CER_BUD_ANNUAL) CER_BUD_ANNUAL," & _
		  "            AVG(FFA_ACT_THISMONTH) FFA_ACT_THISMONTH," & _
		  "            AVG(FFA_ACT_TODATE) FFA_ACT_TODATE," & _
		  "            AVG(FFA_BUD_ANNUAL) FFA_BUD_ANNUAL" & _
          "     from ( SELECT case when companycode = 'PT.THIP' then 'A'" & _
          "                                           when companycode = 'PT.CRS' then 'B'" & _
          "                                           else 'C'" & _
          "                                      end OrderDataRegion," & _
          "                                      case when companycode = 'PT.THIP' then 'THIP'" & _
          "                                           when companycode = 'PT.CRS' then 'CRS'" & _
          "                                           else 'AMS Ganda'" & _
          "                                      end region," & _
          "                                      fcba," & _
          "                                      NVL(SUM(FFBPROCESS_THISMONTH), 0) FFBPROCESS_THISMONTH," & _
		  "                                      NVL(SUM(FFBPROCESS_TODATE), 0) FFBPROCESS_TODATE," & _
		  "                                      NVL(SUM(CPOPROD_THISMONTH), 0) CPOPROD_THISMONTH," & _
		  "                                      NVL(SUM(CPOPROD_TODATE), 0) CPOPROD_TODATE," & _
		  "                                      NVL(SUM(KERPROD_THISMONTH), 0) KERPROD_THISMONTH," & _
		  "                                      NVL(SUM(KERPROD_TODATE), 0) KERPROD_TODATE," & _
		  "                                      NVL(SUM(SHLPROD_THISMONTH), 0) SHLPROD_THISMONTH," & _
		  "                                      NVL(SUM(SHLPROD_TODATE), 0) SHLPROD_TODATE," & _
		  "                                      NVL(SUM(FFAPROD_THISMONTH), 0) FFAPROD_THISMONTH," & _
		  "                                      NVL(SUM(FFAPROD_TODATE), 0) FFAPROD_TODATE," & _
		  "                                      NVL(SUM(FFBPROCESS_BUD_ANNUAL), 0) FFBPROCESS_BUD_ANNUAL," & _
		  "                                      NVL(SUM(FFAPROD_BUD_ANNUAL), 0) FFAPROD_BUD_ANNUAL," & _
		  "                                      NVL(SUM(CPOPROD_BUD_ANNUAL), 0) CPOPROD_BUD_ANNUAL," & _
		  "                                      NVL(SUM(KERPROD_BUD_ANNUAL), 0) KERPROD_BUD_ANNUAL," & _
		  "                                      NVL(SUM(SHLPROD_BUD_ANNUAL), 0) SHLPROD_BUD_ANNUAL," & _
		  "                                      NVL(AVG(OER_ACT_THISMONTH), 0) OER_ACT_THISMONTH," & _
          "                                      NVL(AVG(OER_ACT_TODATE), 0) OER_ACT_TODATE," & _
          "                                      NVL(AVG(OER_BUD_ANNUAL), 0) OER_BUD_ANNUAL," & _
          "                                      NVL(AVG(KER_ACT_THISMONTH), 0) KER_ACT_THISMONTH," & _
          "                                      NVL(AVG(KER_ACT_TODATE), 0) KER_ACT_TODATE," & _
          "                                      NVL(AVG(KER_BUD_ANNUAL), 0) KER_BUD_ANNUAL," & _
          "                                      NVL(AVG(CER_ACT_THISMONTH), 0) CER_ACT_THISMONTH," & _
		  "                                      NVL(AVG(CER_ACT_TODATE), 0) CER_ACT_TODATE," & _
		  "                                      NVL(AVG(CER_BUD_ANNUAL), 0) CER_BUD_ANNUAL," & _
		  "                                      NVL(AVG(FFA_ACT_THISMONTH), 0) FFA_ACT_THISMONTH," & _
		  "                                      NVL(AVG(FFA_ACT_TODATE), 0) FFA_ACT_TODATE," & _
		  "                                      NVL(AVG(FFA_BUD_ANNUAL), 0) FFA_BUD_ANNUAL" & _
          "                               FROM mgt2_millratio_det_tab" & _
          "                               WHERE year = " & v_txtYear & "" & _
          "                                     and month = " & v_txtMonth & "" & _
          "                                     " & whereCons & "" & _
          "                               GROUP BY companycode," & _
          "                                        fcba" & _
          "     )" & _
          "     group by grouping sets( (OrderDataRegion," & _
          "                              region," & _
          "                              fcba)," & _
          "                             (OrderDataRegion," & _
          "                              region)," & _
          "                             () )" & _
          " ) aa" & _
		  " order by orderdataregion, BUSINESSUNIT_API.get_OrderData (fcba)"
		  
'response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Bulanan Rendemen</title>
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
							<td><font face="verdana" size="1">YEAR</font></td>
							<td><font face="verdana" size="1">MONTH</font></td>
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
							<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon

							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(v_txtYear)= 0 then v_txtYear = Yr
							%>
							<td  align="center" width="4">
								<select name="txtYear">
									<%For i = MinYr to MaxYr%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
									<%Next%>
								</select>
							</td>
							<td width="2">
								<select name="txtMonth">
									<%For i = 1 to 12%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
									<%Next%>
								</select>
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
						<th colspan="14" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<h1>AMS PLANTATION</h1>
									<div style="margin-botton:10px;"/>
									<h3>LAPORAN BULANAN RENDEMEN</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th rowspan="3">NO</th>
						<th rowspan="3">MILL</th>
						<th colspan="3">OER</th>
						<th colspan="3">KER</th>
						<th colspan="3">CER</th>
						<th colspan="3">FFA</th>
					</tr>
					<tr>
						<th colspan="2">ACTUAL</th>
						<th rowspan="2">BGT</th>
						<th colspan="2">ACTUAL</th>
						<th rowspan="2">BGT</th>
						<th colspan="2">ACTUAL</th>
						<th rowspan="2">BGT</th>
						<th colspan="2">ACTUAL</th>
						<th rowspan="2">BGT</th>
					</tr>
					<tr>
						<th>T Month</th>
						<th>To date</th>
						<th>T Month</th>
						<th>To date</th>
						<th>T Month</th>
						<th>To date</th>
						<th>T Month</th>
						<th>To date</th>
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

Dim CurrOrderDataRegion : CurrOrderDataRegion = ""
Dim PrevOrderDataRegion : PrevOrderDataRegion = ""

Dim currnextregion : currnextregion = ""
Dim prevnextregion : prevnextregion = ""

Dim Row_percent_thismonth : Row_percent_thismonth = 0
Dim Row_percent_todate : Row_percent_todate = 0


Dim Sub_ffbrocess_act_bi : Sub_ffbrocess_act_bi = 0
Dim Sub_cpoprod_act_bi : Sub_cpoprod_act_bi = 0
Dim Sub_kerprod_act_bi : Sub_kerprod_act_bi = 0
Dim Sub_shlprod_act_bi : Sub_shlprod_act_bi = 0
Dim Sub_oer_act_bi : Sub_oer_act_bi = 0
Dim Sub_ker_act_bi : Sub_ker_act_bi = 0
Dim Sub_cer_act_bi : Sub_cer_act_bi = 0

Dim Sub_ffbrocess_act_sbi : Sub_ffbrocess_act_sbi = 0
Dim Sub_cpoprod_act_sbi : Sub_cpoprod_act_sbi = 0
Dim Sub_kerprod_act_sbi : Sub_kerprod_act_sbi = 0
Dim Sub_shlprod_act_sbi : Sub_shlprod_act_sbi = 0
Dim Sub_oer_act_sbi : Sub_oer_act_sbi = 0
Dim Sub_ker_act_sbi : Sub_ker_act_sbi = 0
Dim Sub_cer_act_sbi : Sub_cer_act_sbi = 0

Dim Sub_ffa_act_bi : Sub_ffa_act_bi = 0
Dim Sub_ffa_act_sbi : Sub_ffa_act_sbi = 0
Dim Sub_ffa_bud_annual : Sub_ffa_bud_annual = 0

Dim Sub_qty_acid_act_bi : Sub_qty_acid_act_bi = 0
Dim Sub_qty_acid_act_sbi : Sub_qty_acid_act_sbi = 0

Dim Sub_ffbrocess_bud_annual : Sub_ffbrocess_bud_annual = 0
Dim Sub_qty_acid_bud_annual : Sub_qty_acid_bud_annual = 0
Dim Sub_cpoprod_bud_annual : Sub_cpoprod_bud_annual = 0
Dim Sub_kernelprod_bud_annual : Sub_kernelprod_bud_annual = 0
Dim Sub_shellprod_bud_annual : Sub_shellprod_bud_annual = 0

Dim Sub_oer_bud_annual : Sub_oer_bud_annual = 0
Dim Sub_ker_bud_annual : Sub_ker_bud_annual = 0
Dim Sub_cer_bud_annual : Sub_cer_bud_annual = 0


if rsLHPP.BOF and rsLHPP.EOF then
%>
					<tr>
						<td colspan="<%=14%>">NO DATA AVAILABLE</td>
					</tr>
<%
else
		while not rsLHPP.eof
			CurrOrderDataRegion = rsLHPP("orderdataregion")
%>
				<%if rsLHPP("NEXTREGION") = "Grand Total" then%>
					<tr bgcolor="#ffff01" >
						<td colspan="2" align="left"><strong><%=rsLHPP("nextregion")%></strong></td>
						<%
						Sub_ffbrocess_act_bi = cdbl(rsLHPP("FFBPROCESS_THISMONTH"))
						Sub_cpoprod_act_bi = cdbl(rsLHPP("CPOPROD_THISMONTH"))
						Sub_kerprod_act_bi = cdbl(rsLHPP("KERPROD_THISMONTH"))
						Sub_shlprod_act_bi = cdbl(rsLHPP("SHLPROD_THISMONTH"))
						Sub_qty_acid_act_bi = cdbl(rsLHPP("FFAPROD_THISMONTH"))
						
						Sub_ffbrocess_act_sbi = cdbl(rsLHPP("FFBPROCESS_TODATE"))
						Sub_cpoprod_act_sbi = cdbl(rsLHPP("CPOPROD_TODATE"))
						Sub_kerprod_act_sbi = cdbl(rsLHPP("KERPROD_TODATE"))
						Sub_shlprod_act_sbi = cdbl(rsLHPP("SHLPROD_TODATE"))
						Sub_qty_acid_act_sbi = cdbl(rsLHPP("FFAPROD_TODATE"))
						
						Sub_ffbrocess_bud_annual = cdbl(rsLHPP("FFBPROCESS_BUD_ANNUAL"))
						Sub_qty_acid_bud_annual = cdbl(rsLHPP("FFAPROD_BUD_ANNUAL"))
						Sub_cpoprod_bud_annual = cdbl(rsLHPP("CPOPROD_BUD_ANNUAL"))
						Sub_kernelprod_bud_annual = cdbl(rsLHPP("KERPROD_BUD_ANNUAL"))
						Sub_shellprod_bud_annual = cdbl(rsLHPP("SHLPROD_BUD_ANNUAL"))
						
						if Sub_ffbrocess_act_bi = 0 then
							Sub_oer_act_bi = 0
							Sub_ker_act_bi = 0
							Sub_cer_act_bi = 0
						else
							Sub_oer_act_bi = Round( (Sub_cpoprod_act_bi/Sub_ffbrocess_act_bi) * 100 , 2)
							Sub_ker_act_bi = Round( (Sub_kerprod_act_bi/Sub_ffbrocess_act_bi) * 100 , 2)
							Sub_cer_act_bi = Round( (Sub_shlprod_act_bi/Sub_ffbrocess_act_bi) * 100 , 2)
						end if
						
						if Sub_ffbrocess_act_sbi = 0 then
							Sub_oer_act_sbi = 0
							Sub_ker_act_sbi = 0
							Sub_cer_act_sbi = 0
						else
							Sub_oer_act_sbi = Round( (Sub_cpoprod_act_sbi/Sub_ffbrocess_act_sbi) * 100 , 2)
							Sub_ker_act_sbi = Round( (Sub_kerprod_act_sbi/Sub_ffbrocess_act_sbi) * 100 , 2)
							Sub_cer_act_sbi = Round( (Sub_shlprod_act_sbi/Sub_ffbrocess_act_sbi) * 100 , 2)
						end if
						
						if Sub_cpoprod_act_bi = 0 then
							Sub_ffa_act_bi = 0
						else
							Sub_ffa_act_bi = Round( (Sub_qty_acid_act_bi/Sub_cpoprod_act_bi) * 100, 2)
						end if
						
						if Sub_cpoprod_act_sbi = 0 then
							Sub_ffa_act_sbi = 0
						else
							Sub_ffa_act_sbi = Round( (Sub_qty_acid_act_sbi/Sub_cpoprod_act_sbi) * 100, 2)
						end if
						
						if Sub_ffbrocess_bud_annual = 0 then
							Sub_oer_bud_annual = 0
							Sub_ker_bud_annual = 0
							Sub_cer_bud_annual = 0
						else
							Sub_oer_bud_annual = Round( (Sub_cpoprod_bud_annual/Sub_ffbrocess_bud_annual) * 100, 2)
							Sub_ker_bud_annual = Round( (Sub_kernelprod_bud_annual/Sub_ffbrocess_bud_annual) * 100, 2)
							Sub_cer_bud_annual = Round( (Sub_shellprod_bud_annual/Sub_ffbrocess_bud_annual) * 100, 2)
						end if
						
						if Sub_cpoprod_bud_annual = 0 then
							Sub_ffa_bud_annual = 0
						else
							Sub_ffa_bud_annual = Round( (Sub_qty_acid_bud_annual/Sub_cpoprod_bud_annual) * 100, 2)
						end if
						%>
						<td align="right"><strong><%=FormatNumber(Sub_oer_act_bi, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_oer_act_sbi, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_oer_bud_annual, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_ker_act_bi, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_ker_act_sbi, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_ker_bud_annual, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_cer_act_bi, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_cer_act_sbi, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_cer_bud_annual, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_ffa_act_bi, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_ffa_act_sbi, 2)%></strong></td>
						<td align="right"><strong><%=FormatNumber(Sub_ffa_bud_annual, 2)%></strong></td>
					</tr>
				<%else%>
					<%if CurrOrderDataRegion <> PrevOrderDataRegion then%>
					<tr>
						<td align="Center" style=" border-bottom:none;"><strong><%=rsLHPP("orderdataregion")%></strong></td>
						<td align="left"><strong><%=rsLHPP("region")%></strong></td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
						<td align="right">&nbsp;</td>
					</tr>
					<tr>
						<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
						<td align="left">&nbsp;&nbsp;&nbsp;&nbsp;<%=rsLHPP("FCBA")%></td>
						<td align="right"><%=FormatNumber(rsLHPP("OER_ACT_THISMONTH"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("OER_ACT_TODATE"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("OER_BUD_ANNUAL"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_ACT_THISMONTH"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_ACT_TODATE"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("KER_BUD_ANNUAL"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("CER_ACT_THISMONTH"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("CER_ACT_TODATE"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("CER_BUD_ANNUAL"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_ACT_THISMONTH"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_ACT_TODATE"), 2)%></td>
						<td align="right"><%=FormatNumber(rsLHPP("FFA_BUD_ANNUAL"), 2)%></td>
					</tr>
						<%PrevOrderDataRegion = CurrOrderDataRegion%>
					<%else%>
						<%if Left(Cstr(rsLHPP("NEXTREGION")), 5) = "Total" Then%>
							<tr bgcolor="#F4F4A0" >
								<td bgcolor="#ffffff" align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
								<td align="left"><strong><%=rsLHPP("nextregion")%></strong></td>
								<%
								Sub_ffbrocess_act_bi = cdbl(rsLHPP("FFBPROCESS_THISMONTH"))
								Sub_cpoprod_act_bi = cdbl(rsLHPP("CPOPROD_THISMONTH"))
								Sub_kerprod_act_bi = cdbl(rsLHPP("KERPROD_THISMONTH"))
								Sub_shlprod_act_bi = cdbl(rsLHPP("SHLPROD_THISMONTH"))
								Sub_qty_acid_act_bi = cdbl(rsLHPP("FFAPROD_THISMONTH"))
								
								Sub_ffbrocess_act_sbi = cdbl(rsLHPP("FFBPROCESS_TODATE"))
								Sub_cpoprod_act_sbi = cdbl(rsLHPP("CPOPROD_TODATE"))
								Sub_kerprod_act_sbi = cdbl(rsLHPP("KERPROD_TODATE"))
								Sub_shlprod_act_sbi = cdbl(rsLHPP("SHLPROD_TODATE"))
								Sub_qty_acid_act_sbi = cdbl(rsLHPP("FFAPROD_TODATE"))
								
								Sub_ffbrocess_bud_annual = cdbl(rsLHPP("FFBPROCESS_BUD_ANNUAL"))
								Sub_qty_acid_bud_annual = cdbl(rsLHPP("FFAPROD_BUD_ANNUAL"))
								Sub_cpoprod_bud_annual = cdbl(rsLHPP("CPOPROD_BUD_ANNUAL"))
								Sub_kernelprod_bud_annual = cdbl(rsLHPP("KERPROD_BUD_ANNUAL"))
								Sub_shellprod_bud_annual = cdbl(rsLHPP("SHLPROD_BUD_ANNUAL"))
								
								if Sub_ffbrocess_act_bi = 0 then
									Sub_oer_act_bi = 0
									Sub_ker_act_bi = 0
									Sub_cer_act_bi = 0
								else
									Sub_oer_act_bi = Round( (Sub_cpoprod_act_bi/Sub_ffbrocess_act_bi) * 100 , 2)
									Sub_ker_act_bi = Round( (Sub_kerprod_act_bi/Sub_ffbrocess_act_bi) * 100 , 2)
									Sub_cer_act_bi = Round( (Sub_shlprod_act_bi/Sub_ffbrocess_act_bi) * 100 , 2)
								end if
								
								if Sub_ffbrocess_act_sbi = 0 then
									Sub_oer_act_sbi = 0
									Sub_ker_act_sbi = 0
									Sub_cer_act_sbi = 0
								else
									Sub_oer_act_sbi = Round( (Sub_cpoprod_act_sbi/Sub_ffbrocess_act_sbi) * 100 , 2)
									Sub_ker_act_sbi = Round( (Sub_kerprod_act_sbi/Sub_ffbrocess_act_sbi) * 100 , 2)
									Sub_cer_act_sbi = Round( (Sub_shlprod_act_sbi/Sub_ffbrocess_act_sbi) * 100 , 2)
								end if
								
								if Sub_cpoprod_act_bi = 0 then
									Sub_ffa_act_bi = 0
								else
									Sub_ffa_act_bi = Round( (Sub_qty_acid_act_bi/Sub_cpoprod_act_bi) * 100, 2)
								end if
								
								if Sub_cpoprod_act_sbi = 0 then
									Sub_ffa_act_sbi = 0
								else
									Sub_ffa_act_sbi = Round( (Sub_qty_acid_act_sbi/Sub_cpoprod_act_sbi) * 100, 2)
								end if
								
								if Sub_ffbrocess_bud_annual = 0 then
									Sub_oer_bud_annual = 0
									Sub_ker_bud_annual = 0
									Sub_cer_bud_annual = 0
								else
									Sub_oer_bud_annual = Round( (Sub_cpoprod_bud_annual/Sub_ffbrocess_bud_annual) * 100, 2)
									Sub_ker_bud_annual = Round( (Sub_kernelprod_bud_annual/Sub_ffbrocess_bud_annual) * 100, 2)
									Sub_cer_bud_annual = Round( (Sub_shellprod_bud_annual/Sub_ffbrocess_bud_annual) * 100, 2)
								end if
								
								if Sub_cpoprod_bud_annual = 0 then
									Sub_ffa_bud_annual = 0
								else
									Sub_ffa_bud_annual = Round( (Sub_qty_acid_bud_annual/Sub_cpoprod_bud_annual) * 100, 2)
								end if
								%>
								<td align="right"><strong><%=FormatNumber(Sub_oer_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_oer_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_oer_bud_annual, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_ker_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_ker_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_ker_bud_annual, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_cer_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_cer_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_cer_bud_annual, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_ffa_act_bi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_ffa_act_sbi, 2)%></strong></td>
								<td align="right"><strong><%=FormatNumber(Sub_ffa_bud_annual, 2)%></strong></td>
							</tr>
						<%else%>
							<tr>
								<td align="center" style="border:none;border-right:1px #3A5F0B solid; border-left:1px #3A5F0B solid;"></td>
								<td align="left">&nbsp;&nbsp;&nbsp;&nbsp;<%=rsLHPP("FCBA")%></td>
								<td align="right"><%=FormatNumber(rsLHPP("OER_ACT_THISMONTH"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("OER_ACT_TODATE"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("OER_BUD_ANNUAL"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("KER_ACT_THISMONTH"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("KER_ACT_TODATE"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("KER_BUD_ANNUAL"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("CER_ACT_THISMONTH"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("CER_ACT_TODATE"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("CER_BUD_ANNUAL"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("FFA_ACT_THISMONTH"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("FFA_ACT_TODATE"), 2)%></td>
								<td align="right"><%=FormatNumber(rsLHPP("FFA_BUD_ANNUAL"), 2)%></td>
							</tr>
						<%end if%>
					<%end if%>
				<%end if%>

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
