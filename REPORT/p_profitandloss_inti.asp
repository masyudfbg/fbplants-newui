<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtjobgroup, v_txtjobcode, v_txtnonull
Dim ScriptName
Dim rsRevenue, strRevenueTotal, strMain, strRevenue
Dim NO, cClass, nday, xNum, nSpanTotal, nCountofField

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

nSpanTotal = 9

v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
v_txtjobgroup = Request("txtjobgroup")
v_txtjobcode = Request("txtjobcode")
v_txtnonull = Request("no-null")

if Request("no-null") = "" Or Len(Trim(Request("no-null"))) = 0 then v_txtnonull = "full"

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Costbook"))
end if

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtOwnership			= Request("txtOwnership")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons
strDefaultBa = 	" SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     FCTYPE = 'HO'" & _
				"     AND ISACTIVE= 'TRUE'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"
' response.write strDefaultBa
Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If	
strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba,"

strFieldCons2	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',mp.fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',mp.fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',mp.fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',mp.fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',mp.fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(mp.fcba) NoUrut, mp.fcba,"
					
wherecons	=	" where company is not null "

wherecons2	=	" where company is not null "

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		selectCons2 = 	selectCons2 & " "
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		selectCons2 = 	selectCons2 & " "
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		selectCons2 = 	selectCons2 & " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		whereCons2	=	wherecons2	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		GroupCons   =   GroupCons   & " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		selectCons2 = 	selectCons2 & " "
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		selectCons2 = 	selectCons2 & " , Plantation"
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " , Plantation"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		selectCons2 = 	selectCons2 & " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		whereCons2	=	whereCons2	& " And Plantation = '" & v_txtPlantation & "'"
		GroupCons   =   GroupCons   & " , Plantation"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		selectCons2 = 	selectCons2 & " "
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		selectCons2 = 	selectCons2 & " , GroupRegion"
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " , GroupRegion"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		selectCons2 = 	selectCons2 & " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		whereCons2	=	whereCons2	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		GroupCons   =   GroupCons   & " , GroupRegion"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		selectCons2 = 	selectCons2 & " "
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		selectCons2 = 	selectCons2 & " , Region"
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " , Region"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		selectCons2 = 	selectCons2 & " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		whereCons2	=	whereCons2	& " And Region = '" & v_txtRegion & "'"
		GroupCons   =   GroupCons   & " , Region"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		selectCons2 = 	selectCons2 & " "
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		selectCons2 = 	selectCons2 & " , company"
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " , company"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
	Else
		selectCons	=	selectCons	& " , company"
		selectCons2 = 	selectCons2 & " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		whereCons2	=	whereCons2	& " And company = '" & v_txtCompany & "'"
		GroupCons   =   GroupCons   & " , company"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsRevenue = connStr.Execute(q)		
		if not rsRevenue.eof then
			CompanyName = rsRevenue("FCCOMPANYNAME")	
		end if
		q	=	""
		rsRevenue.Close
		set rsRevenue = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) = 0 Then v_txtfcba = "--SUMMARY--"

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		selectCons2 = 	selectCons2 & " "
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		selectCons2 = 	selectCons2 & " , nourut, 'HOA' fcba"
		whereCons	=	wherecons	& " "
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " , nourut, fcba"
		orderCons	=	orderCons	& " , nourut, fcba"
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		selectCons2 = 	selectCons2 & " , nourut, 'HOA' fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		whereCons2	=	whereCons2	& " "
		GroupCons   =   GroupCons   & " , nourut, fcba"
		orderCons	=	orderCons	& " , nourut, fcba"
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsRevenue = connStr.Execute(q)		
		if not rsRevenue.eof then
			fcbaname = rsRevenue("FCNAME")
			CompanyName=rsRevenue("FCCOMPANYNAME")
			MgrName = rsRevenue("FCMANAGER")
			KTUName = rsRevenue("FCKTU")
			CropName=rsRevenue("FCCROP")
			FcbaType=rsRevenue("FCTYPE")
		end if
		q	=	""
		rsRevenue.Close
		set rsRevenue = Nothing
	End If
End If



ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain = 	"    		SELECT "& strFieldCons &"" & _
			"                    DECODE (param_1, 0, FCCODE, param_1) param_1," & _
			"                    param_2," & _
			"                    param_3," & _
			"                    param_group," & _
			"                    fccode," & _
			"                    jobprofitandlosscode," & _
			"                    SUM (amount_act) amount_act," & _
			"                    SUM (amount_sbi) amount_sbi," & _
			"                    SUM (amount_prev_act) amount_prev_act," & _
			"                    SUM (amount_prev_sbi) amount_prev_sbi" & _
			"               FROM (  SELECT STAGE1.fcba," & _
			"                              STAGE1.param_1," & _
			"                              STAGE1.param_2," & _
			"                              prd.param_3," & _
			"                              prd.fcname param_group," & _
			"                              STAGE1.fccode," & _
			"                              STAGE1.jobprofitandlosscode," & _
			"                              SUM (NVL (STAGE2.amount_act, 0)) amount_act," & _
			"                              SUM (NVL (STAGE2.amount_sbi, 0)) amount_sbi," & _
			"                              SUM (NVL (STAGE2.amount_prev_act, 0)) amount_prev_act," & _
			"                              SUM (NVL (STAGE2.amount_prev_sbi, 0)) amount_prev_sbi" & _
			"                         FROM (    SELECT fcba," & _
			"                                          fccode," & _
			"                                          param_1," & _
			"                                          param_2," & _
			"                                          param_3," & _
			"                                             LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                          || TO_CHAR (fcname)" & _
			"                                             jobprofitandlosscode" & _
			"                                     FROM parameterdetail" & _
			"                                    WHERE     parcode = 'JOBCODE_PROFITANDLOSS'" & _
			"                                          AND param_1 <> -1" & _
			"                                          AND parheadcode = 'JOBCODE_COST'" & _
			"                               START WITH fccode = '0'" & _
			"                               CONNECT BY PRIOR fccode = param_1" & _
			"                                 GROUP BY fcba," & _
			"                                          fccode," & _
			"                                          param_1," & _
			"                                          param_2," & _
			"                                          param_3," & _
			"                                             LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                          || TO_CHAR (fcname)) STAGE1" & _
			"                              INNER JOIN" & _
			"                              (SELECT fcba," & _
			"                                      fccode param_3," & _
			"                                      param_1," & _
			"                                      param_2," & _
			"                                      fcname" & _
			"                                 FROM parameterdetail" & _
			"                                WHERE     parcode = 'JOBCODE_PROFITANDLOSS_GROUP'" & _
			"                                      AND parameterdetail.param_1 <> -1) prd" & _
			"                                 ON (    stage1.fcba = prd.fcba" & _
			"                                     AND stage1.param_3 = prd.param_3)" & _
			"                              LEFT JOIN" & _
			"                              (  SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        ( (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtYear &"' AND period_month = '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT ca.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        (-1 * SUM (amount)) amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM cost_activity ca" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    ca.fcba = jb.fcba" & _
			"                                               AND ca.jobcode = jb.fccode)" & _
			"                                  WHERE     accountingyear = '"& v_txtYear &"'" & _
			"                                        AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                        AND sourcetype = 'SL'" & _
			"                                        AND jobcode IN ('5110101', '5110102')" & _
			"                               GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT hpp.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        SUM (hpp.amount) amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM (SELECT 'CPO' AS product_code," & _
			"                                                '5210101' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"'" & _
			"                                                AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140205'" & _
			"                                         UNION ALL" & _
			"                                         SELECT 'KER' AS product_code," & _
			"                                                '5210102' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"'" & _
			"                                                AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140206') hpp" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    hpp.jobcode_HPP = jb.fccode" & _
			"                                               AND hpp.fcba = jb.fcba)" & _
			"                               GROUP BY hpp.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        ( (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtYear &"' AND period_month >= '7' AND period_month <= '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT ca.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        (-1 * SUM (amount)) amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM cost_activity ca" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    ca.fcba = jb.fcba" & _
			"                                               AND ca.jobcode = jb.fccode)" & _
			"                                  WHERE     accountingyear = '"& v_txtYear &"'" & _
			"                                        AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                        AND sourcetype = 'SL'" & _
			"                                        AND jobcode IN ('5110101', '5110102')" & _
			"                               GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT hpp.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        SUM (hpp.amount) amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM (SELECT 'CPO' AS product_code," & _
			"                                                '5210101' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"'" & _
			"                                                AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140205'" & _
			"                                         UNION ALL" & _
			"                                         SELECT 'KER' AS product_code," & _
			"                                                '5210102' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"'" & _
			"                                                AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140206') hpp" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    hpp.jobcode_HPP = jb.fccode" & _
			"                                               AND hpp.fcba = jb.fcba)" & _
			"                               GROUP BY hpp.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        ( (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE     period_year = '"& v_txtYear &"' - 1" & _
			"                                        AND period_month = '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT ca.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        (-1 * SUM (amount)) amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM cost_activity ca" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    ca.fcba = jb.fcba" & _
			"                                               AND ca.jobcode = jb.fccode)" & _
			"                                  WHERE     accountingyear = '"& v_txtYear &"' - 1" & _
			"                                        AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                        AND sourcetype = 'SL'" & _
			"                                        AND jobcode IN ('5110101', '5110102')" & _
			"                               GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT hpp.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        SUM (hpp.amount) amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM (SELECT 'CPO' AS product_code," & _
			"                                                '5210101' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"' - 1" & _
			"                                                AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140205'" & _
			"                                         UNION ALL" & _
			"                                         SELECT 'KER' AS product_code," & _
			"                                                '5210102' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"' - 1" & _
			"                                                AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140206') hpp" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    hpp.jobcode_HPP = jb.fccode" & _
			"                                               AND hpp.fcba = jb.fcba)" & _
			"                               GROUP BY hpp.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        ( (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_prev_sbi" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE     period_year = '"& v_txtYear &"' - 1" & _
			"                                        AND period_month >= '7' AND period_month <= '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT ca.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        (-1 * SUM (amount)) amount_prev_sbi" & _
			"                                   FROM cost_activity ca" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    ca.fcba = jb.fcba" & _
			"                                               AND ca.jobcode = jb.fccode)" & _
			"                                  WHERE     accountingyear = '"& v_txtYear &"' - 1" & _
			"                                        AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                        AND sourcetype = 'SL'" & _
			"                                        AND jobcode IN ('5110101', '5110102')" & _
			"                               GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT hpp.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        SUM (hpp.amount) amount_prev_sbi" & _
			"                                   FROM (SELECT 'CPO' AS product_code," & _
			"                                                '5210101' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"' - 1" & _
			"                                                AND accountingperiod >= '7' AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140205'" & _
			"                                         UNION ALL" & _
			"                                         SELECT 'KER' AS product_code," & _
			"                                                '5210102' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"' - 1" & _
			"                                                AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140206') hpp" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    hpp.jobcode_HPP = jb.fccode" & _
			"                                               AND hpp.fcba = jb.fcba)" & _
			"                               GROUP BY hpp.fcba, jb.job_forprofitandlosscode)" & _
			"                              STAGE2" & _
			"                                 ON (    STAGE1.FCBA = STAGE2.FCBA" & _
			"                                     AND STAGE1.FCCODE =" & _
			"                                            STAGE2.JOB_FORPROFITANDLOSSCODE)" & _
			"                     GROUP BY STAGE1.fcba," & _
			"                              STAGE1.param_1," & _
			"                              STAGE1.param_2," & _
			"                              prd.param_3," & _
			"                              prd.fcname," & _
			"                              STAGE1.fccode," & _
			"                              STAGE1.jobprofitandlosscode" & _
			"                     ORDER BY fccode)" & _
			"           GROUP BY fcba," & _
			"                    DECODE (param_1, 0, FCCODE, param_1)," & _
			"                    param_2," & _
			"                    param_3," & _
			"                    param_group," & _
			"                    fccode," & _
			"                    jobprofitandlosscode" & _
			"           UNION ALL" & _
			"             SELECT "& strFieldCons &"" & _
			"                    DECODE (param_1, 0, FCCODE, param_1) param_1," & _
			"                    param_2," & _
			"                    param_3," & _
			"                    param_group," & _
			"                    fccode," & _
			"                    jobprofitandlosscode," & _
			"                    SUM (amount_act) amount_act," & _
			"                    SUM (amount_sbi) amount_sbi," & _
			"                    SUM (amount_prev_act) amount_prev_act," & _
			"                    SUM (amount_prev_sbi) amount_prev_sbi" & _
			"               FROM (  SELECT STAGE1.fcba," & _
			"                              STAGE1.param_1," & _
			"                              STAGE1.param_2," & _
			"                              prd.param_3," & _
			"                              prd.fcname param_group," & _
			"                              STAGE1.fccode," & _
			"                              STAGE1.jobprofitandlosscode," & _
			"                              SUM (NVL (STAGE2.amount_act, 0)) amount_act," & _
			"                              SUM (NVL (STAGE2.amount_sbi, 0)) amount_sbi," & _
			"                              SUM (NVL (STAGE2.amount_prev_act, 0)) amount_prev_act," & _
			"                              SUM (NVL (STAGE2.amount_prev_sbi, 0)) amount_prev_sbi" & _
			"                         FROM (    SELECT fcba," & _
			"                                          fccode," & _
			"                                          param_1," & _
			"                                          param_2," & _
			"                                          param_3," & _
			"                                             LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                          || TO_CHAR (fcname)" & _
			"                                             jobprofitandlosscode" & _
			"                                     FROM parameterdetail" & _
			"                                    WHERE     parcode = 'JOBCODE_PROFITANDLOSS'" & _
			"                                          AND param_1 <> -1" & _
			"                                          AND parheadcode = 'JOBCODE_REVENUE'" & _
			"                               START WITH fccode = '0'" & _
			"                               CONNECT BY PRIOR fccode = param_1" & _
			"                                 GROUP BY fcba," & _
			"                                          fccode," & _
			"                                          param_1," & _
			"                                          param_2," & _
			"                                          param_3," & _
			"                                             LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                          || TO_CHAR (fcname)) STAGE1" & _
			"                              INNER JOIN" & _
			"                              (SELECT fcba," & _
			"                                      fccode param_3," & _
			"                                      param_1," & _
			"                                      param_2," & _
			"                                      fcname" & _
			"                                 FROM parameterdetail" & _
			"                                WHERE     parcode = 'JOBCODE_PROFITANDLOSS_GROUP'" & _
			"                                      AND parameterdetail.param_1 <> -1) prd" & _
			"                                 ON (    stage1.fcba = prd.fcba" & _
			"                                     AND stage1.param_3 = prd.param_3)" & _
			"                              LEFT JOIN" & _
			"                              (  SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        ( (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtYear &"' AND period_month = '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT ca.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        (-1 * SUM (amount)) amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM cost_activity ca" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    ca.fcba = jb.fcba" & _
			"                                               AND ca.jobcode = jb.fccode)" & _
			"                                  WHERE     accountingyear = '"& v_txtYear &"'" & _
			"                                        AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                        AND sourcetype = 'SL'" & _
			"                                        AND jobcode IN ('5110101', '5110102')" & _
			"                               GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT hpp.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        SUM (hpp.amount) amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM (SELECT 'CPO' AS product_code," & _
			"                                                '5210101' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"'" & _
			"                                                AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140205'" & _
			"                                         UNION ALL" & _
			"                                         SELECT 'KER' AS product_code," & _
			"                                                '5210102' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"'" & _
			"                                                AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140206') hpp" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    hpp.jobcode_HPP = jb.fccode" & _
			"                                               AND hpp.fcba = jb.fcba)" & _
			"                               GROUP BY hpp.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        ( (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtYear &"' AND period_month >= '7' AND period_month <= '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT ca.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        (-1 * SUM (amount)) amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM cost_activity ca" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    ca.fcba = jb.fcba" & _
			"                                               AND ca.jobcode = jb.fccode)" & _
			"                                  WHERE     accountingyear = '"& v_txtYear &"'" & _
			"                                        AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                        AND sourcetype = 'SL'" & _
			"                                        AND jobcode IN ('5110101', '5110102')" & _
			"                               GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT hpp.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        SUM (hpp.amount) amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM (SELECT 'CPO' AS product_code," & _
			"                                                '5210101' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"'" & _
			"                                                AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140205'" & _
			"                                         UNION ALL" & _
			"                                         SELECT 'KER' AS product_code," & _
			"                                                '5210102' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"'" & _
			"                                                AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140206') hpp" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    hpp.jobcode_HPP = jb.fccode" & _
			"                                               AND hpp.fcba = jb.fcba)" & _
			"                               GROUP BY hpp.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        ( (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE     period_year = '"& v_txtYear &"' - 1" & _
			"                                        AND period_month = '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT ca.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        (-1 * SUM (amount)) amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM cost_activity ca" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    ca.fcba = jb.fcba" & _
			"                                               AND ca.jobcode = jb.fccode)" & _
			"                                  WHERE     accountingyear = '"& v_txtYear &"' - 1" & _
			"                                        AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                        AND sourcetype = 'SL'" & _
			"                                        AND jobcode IN ('5110101', '5110102')" & _
			"                               GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT hpp.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        SUM (hpp.amount) amount_prev_act," & _
			"                                        0 amount_prev_sbi" & _
			"                                   FROM (SELECT 'CPO' AS product_code," & _
			"                                                '5210101' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"' - 1" & _
			"                                                AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140205'" & _
			"                                         UNION ALL" & _
			"                                         SELECT 'KER' AS product_code," & _
			"                                                '5210102' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"' - 1" & _
			"                                                AND accountingperiod = '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140206') hpp" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    hpp.jobcode_HPP = jb.fccode" & _
			"                                               AND hpp.fcba = jb.fcba)" & _
			"                               GROUP BY hpp.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        ( (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_prev_sbi" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE     period_year = '"& v_txtYear &"' - 1" & _
			"                                        AND period_month >= '7' AND period_year <= '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT ca.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        (-1 * SUM (amount)) amount_prev_sbi" & _
			"                                   FROM cost_activity ca" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    ca.fcba = jb.fcba" & _
			"                                               AND ca.jobcode = jb.fccode)" & _
			"                                  WHERE     accountingyear = '"& v_txtYear &"' - 1" & _
			"                                        AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                        AND sourcetype = 'SL'" & _
			"                                        AND jobcode IN ('5110101', '5110102')" & _
			"                               GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT hpp.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev_act," & _
			"                                        SUM (hpp.amount) amount_prev_sbi" & _
			"                                   FROM (SELECT 'CPO' AS product_code," & _
			"                                                '5210101' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"' - 1" & _
			"                                                AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140205'" & _
			"                                         UNION ALL" & _
			"                                         SELECT 'KER' AS product_code," & _
			"                                                '5210102' AS jobcode_HPP," & _
			"                                                amount * -1 AS amount," & _
			"                                                'HOA' fcba" & _
			"                                           FROM cost_activity ca" & _
			"                                                INNER JOIN businessunit bu" & _
			"                                                   ON (bu.fccode = ca.fcba)" & _
			"                                          WHERE     fctype = 'M'" & _
			"                                                AND accountingyear = '"& v_txtYear &"' - 1" & _
			"                                                AND accountingperiod >= '7' AND accountingperiod <= '"& v_txtMonth &"'" & _
			"                                                AND referenceno =" & _
			"                                                       'Cr-HARGAPOKOKPPENJUALAN'" & _
			"                                                AND sourcetype IN ('SL')" & _
			"                                                AND jobcode = '1140206') hpp" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    hpp.jobcode_HPP = jb.fccode" & _
			"                                               AND hpp.fcba = jb.fcba)" & _
			"                               GROUP BY hpp.fcba, jb.job_forprofitandlosscode)" & _
			"                              STAGE2" & _
			"                                 ON (    STAGE1.FCBA = STAGE2.FCBA" & _
			"                                     AND STAGE1.FCCODE =" & _
			"                                            STAGE2.JOB_FORPROFITANDLOSSCODE)" & _
			"                     GROUP BY STAGE1.fcba," & _
			"                              STAGE1.param_1," & _
			"                              STAGE1.param_2," & _
			"                              prd.param_3," & _
			"                              prd.fcname," & _
			"                              STAGE1.fccode," & _
			"                              STAGE1.jobprofitandlosscode" & _
			"                     ORDER BY fccode)" & _
			"           GROUP BY fcba," & _
			"                    DECODE (param_1, 0, FCCODE, param_1)," & _
			"                    param_2," & _
			"                    param_3," & _
			"                    param_group," & _
			"                    fccode," & _
			"                    jobprofitandlosscode" 

if  len(trim(v_txtjobgroup)) > 0 then
    strMain = strMain & " and job_api.get_jobgroupCode(fccode, fcba) = " & v_txtjobgroup & ""
end if
if  len(trim(v_txtjobcode)) > 0 then
    strMain = strMain & " and fccode = '" & v_txtjobcode & "'"
end if 
	' strMain = strMain & " group by fcba" & _
 '                        "          , fccode" 
							
	strRevenue = 	"          ,DECODE (param_1, 0, FCCODE, param_1) param_1," & _
					"          param_2," & _
					"          param_3," & _
					"          param_group," & _
					"          fccode," & _
					"          jobprofitandlosscode," & _
					"          NVL (SUM (amount_act), 0) AMOUNT," & _
					"          NVL (SUM (amount_sbi), 0) AMOUNT2," & _
					"          NVL (SUM (amount_prev_act), 0) AMOUNT3," & _
					"          NVL (SUM (amount_prev_sbi), 0) AMOUNT4" 

strRevenueTotal	=	" select GroupPlantation " & selectCons & strRevenue &"  From (" & strMain & " )" & _ 
                    "" & whereCons & _
                    " group by GroupPlantation" & selectCons & ", DECODE (param_1, 0, FCCODE, param_1), param_2, param_3, param_group, fccode, jobprofitandlosscode"

strRevenueTotal = strRevenueTotal & " Order by fccode, GroupPlantation" & orderCons & ", param_1, param_2, param_3, param_group, jobprofitandlosscode"

strMain3 = 	"    SELECT "& strFieldCons2 &"" & _
			"                    mp.product_code," & _
			"                    MAX (TO_NUMBER (TO_CHAR (mp.tdate, 'mm'))) month_prod_bi," & _
			"                    MAX (TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy'))) year_prod_bi," & _
			"                    SUM (mp.production_today) qty_prod_bi," & _
			"                    0 qty_prod_sbi" & _
			"               FROM mill_production mp" & _
			"               INNER JOIN businessunit bu ON (bu.fccode = mp.fcba)" & _
			"              WHERE     TO_NUMBER (TO_CHAR (mp.tdate, 'mm')) = '"& v_txtMonth &"'" & _
			"                    AND TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy')) = '"& v_txtYear &"'" & _
			"                    AND mp.product_code IN ('CPO', 'KER')" & _
			"					 AND bu.fctype = 'M'"& _
			"					 AND bu.isactive = 'TRUE'"& _
			"           GROUP BY mp.fcba, mp.product_code" & _
			"           UNION ALL" & _
			"             SELECT "& strFieldCons2 &"" & _
			"                    mp.product_code," & _
			"                    MAX (TO_NUMBER (TO_CHAR (mp.tdate, 'mm'))) month_prod_bi," & _
			"                    MAX (TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy'))) year_prod_bi," & _
			"                    0 qty_prod_bi," & _
			"                    SUM (mp.production_today) qty_prod_sbi" & _
			"               FROM mill_production mp" & _
			"               INNER JOIN businessunit bu ON (bu.fccode = mp.fcba)" & _
			"              WHERE     TO_NUMBER (TO_CHAR (mp.tdate, 'mm')) <= '"& v_txtMonth &"'" & _
			"                    AND TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy')) = '"& v_txtYear &"'" & _
			"                    AND mp.product_code IN ('CPO', 'KER')" & _
			"					 AND bu.fctype = 'M'"& _
			"					 AND bu.isactive = 'TRUE'"& _
			"           GROUP BY mp.fcba, mp.product_code"  

 
if  len(trim(v_txtjobcode)) > 0 then
    strMain3 = strMain3 & " and fccode = '" & v_txtjobcode & "'"
end if 
	' strMain3 = strMain3 & " group by fcba" & _
 '                        "          , fccode" 
								
strProd =   "          , product_code," & _
			"          month_prod_bi," & _
			"          year_prod_bi," & _
			"          SUM (qty_prod_bi) qty_prod_bi," & _
			"          SUM (qty_prod_sbi) qty_prod_sbi" 

strProdTotal	=	" select GroupPlantation " & selectCons2 & strProd &"  From (" & strMain3 & " )" & _ 
                    "" & whereCons2 & _
                    " group by GroupPlantation, " & _
					"          product_code," & _
					"          month_prod_bi," & _
					"          year_prod_bi" & GroupCons & ""

strProdTotal = strProdTotal & " Order by GroupPlantation" & orderCons & ""

' response.write strProdTotal

'Response.write strRevenueTotal
%>
<head>
<title>ePlaS Reporting - Statement Of Profit Or Loss And Other Comprehensive Income</title>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link rel="stylesheet" href="/css/Neraca.css" type="text/css" />
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<style type="text/css">
@media print
{
body{font-family:'Arial Narrow';font-size:12px;}
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
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSumPNL.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td><font face="verdana" size="1">YEAR</font></td>
				<% 
					Dim mon, Yr, MinYr, MaxYr, i, MM
					mon = Month(Session("v_Costbook"))
					If len(v_txtMonth)=0 then v_txtMonth=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = 2015
					MaxYr = Yr  + 2
					If len(v_txtYear)= 0 then v_txtYear = Yr
				%>
						<td width="4">
							<select name="txtYear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
							</select>
						</td>
						<td><font face="verdana" size="1">MONTH</font></td>
						<td width="2">
							<select name="txtMonth">
				<%
						For i = 1 to 15
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=BulanAP(i)%></option>
				<%
					Next
				%>
							</select>
						</td>
				<!-- <td>
					<font face="verdana" size="1">OWNERSHIP</font>
				</td>
				<td width="2">
					<select name="txtOwnership" >
						<option value="ALL">--ALL--</option>
						<%
						Dim rsOwnership, strOwnership
						strOwnership	=	" SELECT DISTINCT P.FCCODE,P.FCNAME FROM PARAMETER P WHERE P.PARCODE= 'F0004' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
						
						Set rsOwnership = Server.CreateObject("ADODB.Recordset")
						rsOwnership.Open strOwnership, connstr, adOpenStatic 
						
						if not rsOwnership.eof then
							Do While Not rsOwnership.eof
							if rsOwnership("FCCODE") = v_txtOwnership then%>
							<option value="<%=rsOwnership("FCCODE")%>" Selected><%=rsOwnership("FCCODE")%></option> 
							<%Else%>
							<option value="<%=rsOwnership("FCCODE")%>"><%=rsOwnership("FCCODE")%></option>  
							<%End if
								rsOwnership.movenext
							loop
						end if
						rsOwnership.Close
						set rsOwnership = Nothing
						%>
					</select>
				</td> -->
						<td><input type="submit" value="View" /></td>
						<td><input type="hidden" name="no-null" value="<%=v_txtnonull%>"></td>
					</tr>
				</table>
			</div>
			</form>
		</div>

<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
		<thead >
		<tr> 
			<div id="header-wrapper"> 
				<div id="subheader" class="centered">
					<h2>PT. <%=CompanyName%></h2>
					<br>
					<h2>STATEMENT OF PROFIT OR LOSS</h2>
					<br>
					<h3>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h3> 
					<br>
				</div> 
			</div> 
		</tr> 
	</thead> 
</table> 
<table border="0" width="100%" id="product-table" class="xthf-site">
  <thead>
    <tr> 
		<th align="center" style="color: #fff; width: 400px" rowspan="2">SUBJECT</th>
		<th align="center" style="color: #fff; width: 300px" colspan="2">BULAN INI</th> 
		<th align="center" style="color: #fff; width: 300px" colspan="2">S/D BULAN INI</th> 
	</tr>
	<tr>
		<th align="center" style="color: #fff; width: 150px"><%=v_txtyear%></th> 
		<th align="center" style="color: #fff; width: 150px"><%=v_txtyear-1%></th> 

		<th align="center" style="color: #fff; width: 150px"><%=v_txtyear%></th> 
		<th align="center" style="color: #fff; width: 150px"><%=v_txtyear-1%></th> 
	</tr>
  </thead>
<tbody>
<%	Set rsRevenue = connStr.Execute(strRevenueTotal)
	if not rsRevenue.BOF and rsRevenue.EOF then%>
  <tr>
    <td colspan="<%= 8 + var_colspan%>" align="left">NO DATA AVAILABLE</td>
  </tr>
<%end if
NO = 1 
NO = 1    
    
while not rsRevenue.EOF
 
AMOUNT = rsRevenue("AMOUNT")
AMOUNT2 = rsRevenue("AMOUNT2")
' AMOUNT3 = rsRevenue("AMOUNT3")
Currweekly = rsRevenue("param_group")

If Currweekly <> Prevweekly Then
	tempAmount = tempAmount + cdbl(to_nol(rsRevenue("AMOUNT")))
	tempAmount2 = tempAmount2 + cdbl(to_nol(rsRevenue("AMOUNT2")))
	tempAmount3 = tempAmount3 + cdbl(to_nol(rsRevenue("AMOUNT3")))
	tempAmount4 = tempAmount4 + cdbl(to_nol(rsRevenue("AMOUNT4")))

	v_tempAmount = v_tempAmount + cdbl(to_nol(rsRevenue("AMOUNT")))
	v_tempAmount2 = v_tempAmount2 + cdbl(to_nol(rsRevenue("AMOUNT2"))) 
	v_tempAmount3 = v_tempAmount3 + cdbl(to_nol(rsRevenue("AMOUNT3")))
	v_tempAmount4 = v_tempAmount4 + cdbl(to_nol(rsRevenue("AMOUNT4"))) 

	v_total_hpp = v_total_hpp + cdbl(to_nol(rsRevenue("AMOUNT")))
	v_total_hpp2 = v_total_hpp2 + cdbl(to_nol(rsRevenue("AMOUNT2")))
	v_total_hpp3 = v_total_hpp3 + cdbl(to_nol(rsRevenue("AMOUNT3")))
	v_total_hpp4 = v_total_hpp4 + cdbl(to_nol(rsRevenue("AMOUNT4")))

	v_total_pb = v_total_pb + cdbl(to_nol(rsRevenue("AMOUNT")))
	v_total_pb2 = v_total_pb2 + cdbl(to_nol(rsRevenue("AMOUNT2")))
	v_total_pb3 = v_total_pb3 + cdbl(to_nol(rsRevenue("AMOUNT3")))
	v_total_pb4 = v_total_pb4 + cdbl(to_nol(rsRevenue("AMOUNT4")))

	v_total_pl = v_total_pl + cdbl(to_nol(rsRevenue("AMOUNT")))
	v_total_pl2 = v_total_pl2 + cdbl(to_nol(rsRevenue("AMOUNT2")))
	v_total_pl3 = v_total_pl3 + cdbl(to_nol(rsRevenue("AMOUNT3")))
	v_total_pl4 = v_total_pl4 + cdbl(to_nol(rsRevenue("AMOUNT4")))
 
	If 	rsRevenue("param_3") = 401000000 then  
		total_pb =  tempAmount
		total_pb2 = tempAmount2 
		total_pb3 =  tempAmount3
		total_pb4 = tempAmount4 
		Else If	rsRevenue("param_3") = 402010000 then  
				tempAmount =  tempAmount
				tempAmount2 = tempAmount2
				tempAmount3 =  tempAmount3
				tempAmount4 = tempAmount4

				total_hpp =  tempAmount
				total_hpp2 =  tempAmount2 
				total_hpp3 =  tempAmount3
				total_hpp4 =  tempAmount4 
				Else If	rsRevenue("param_3") = 402010999 then 
					tempAmount =  total_pb - total_hpp
					tempAmount2 =  total_pb2 - total_hpp2 
					tempAmount3 =  total_pb3 - total_hpp3
					tempAmount4 =  total_pb4 - total_hpp4 

					total_lk =  total_pb - total_hpp
					total_lk2 =  total_pb2 - total_hpp2  
					total_lk3 =  total_pb3 - total_hpp3
					total_lk4 =  total_pb4 - total_hpp4  
					Else If	rsRevenue("param_3") = 402020000 then 
						total_bup =  tempAmount
						total_bup2 =  tempAmount2
						total_bup3 =  tempAmount3
						total_bup4 =  tempAmount4
						Else If	rsRevenue("param_3") = 602000000 then 
							total_bua =  tempAmount
							total_bua2 =  tempAmount2
							total_bua3 =  tempAmount3
							total_bua4 =  tempAmount4
							Else If	rsRevenue("param_3") = 602199999 then 
								tempAmount =  total_bup + total_bua
								tempAmount2 =  total_bup2 + total_bua2  
								tempAmount3 =  total_bup3 + total_bua3
								tempAmount4 =  total_bup4 + total_bua4  

								total_bo =  total_bup + total_bua
								total_bo2 =  total_bup2 + total_bua2  
								total_bo3 =  total_bup3 + total_bua3
								total_bo4 =  total_bup4 + total_bua4  
								Else If	rsRevenue("param_3") = 602299999 then 
									tempAmount =  total_lk - total_bo
									tempAmount2 =  total_lk2 - total_bo2
									tempAmount3 =  total_lk3 - total_bo3
									tempAmount4 =  total_lk4 - total_bo4

									total_lo =  total_lk - total_bo
									total_lo2 =  total_lk2 - total_bo2
									total_lo3 =  total_lk3 - total_bo3
									total_lo4 =  total_lk4 - total_bo4
									Else If	rsRevenue("param_3") = 700000000 then  
										total_pl =  tempAmount
										total_pl2 =  tempAmount2 
										total_pl3 =  tempAmount3
										total_pl4 =  tempAmount4 
										Else If	rsRevenue("param_3") = 799999999 then 
											tempAmount =  total_lo - total_pl
											tempAmount2 =  total_lo2 - total_pl2  
											tempAmount3 =  total_lo3 - total_pl3
											tempAmount4 =  total_lo4 - total_pl4  
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If
%>
<%If Len(Trim(Prevweekly)) = 0 Then%>
		<tr>
			<td align="left" style="padding-left:10px;" bgcolor="#F4F4A0"><font face="Verdana" size="1"><strong><%=Currweekly%></strong></font></td> 
			<!-- This Month Formula -->
			<td align="Right" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 401000000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_pb)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_hpp)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010999  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602199999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602299999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 799999999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 2)%></strong></font>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
			</td>  
			<td align="Right" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 401000000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_pb3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_hpp3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010999  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602199999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602299999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 799999999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 2)%></strong></font>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
			</td>  
			<!-- End This Month Formula -->
			<!-- To Date Formula -->
			<td align="Right" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 401000000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_pb2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_hpp2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010999  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602199999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602299999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 799999999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 2)%></strong></font>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
			</td>  
			<td align="Right" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 401000000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_pb4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_hpp4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010999  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602199999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602299999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 799999999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount4)), 2)%></strong></font>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
			</td>  
			<!-- End To Date Formula -->
		</tr> 
<%
	tempAmount = 0
	tempAmount2 = 0
	tempAmount3 = 0 
	tempAmount4 = 0 

	v_total_hpp = 0
	v_total_hpp2 = 0
	v_total_hpp3 = 0
	v_total_hpp4 = 0
	v_total_bua = 0
	v_total_bua2 = 0
	v_total_bua3 = 0
	v_total_bua4 = 0
	v_total_bup = 0
	v_total_bup2 = 0
	v_total_bup3 = 0
	v_total_bup4 = 0
	v_total_bo = 0
	v_total_bo2 = 0
	v_total_bo3 = 0
	v_total_bo4 = 0
	v_total_pl = 0
	v_total_pl2 = 0
	v_total_pl3 = 0
	v_total_pl4 = 0

	v_tempAmount = 0
	v_tempAmount2 = 0
	v_tempAmount3 = 0 
%>
<%Else%>
		<tr>
			<td align="left" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 402020000 Then %>
					<font face="Verdana" size="1"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=Currweekly%></strong></font>
				<%Else If rsRevenue("param_3") = 602000000 Then %>
					<font face="Verdana" size="1"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=Currweekly%></strong></font>
				<%Else%>
					<font face="Verdana" size="1"><strong><%=Currweekly%></strong></font>
				<%End If%>
				<%End If%>
			</td>   
			<!-- This Month Formula -->
			<td align="Right" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 402010000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_hpp)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010999  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602199999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602299999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 700000000  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_pl)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 799999999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 2)%></strong></font>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
			</td> 
			<td align="Right" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 402010000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_hpp3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010999  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602199999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602299999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 700000000  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_pl3)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 799999999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 2)%></strong></font>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
			</td> 
			<!-- End This Month Formula -->
			<!-- To Date Formula -->
			<td align="Right" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 402010000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_hpp2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010999  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602199999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602299999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 700000000  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_pl2)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 799999999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 2)%></strong></font>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
			</td> 
			<td align="Right" style="padding-left:10px;" bgcolor="#F4F4A0">
				<%If rsRevenue("param_3") = 402010000  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_hpp4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 402010999  Then %> 
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602199999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 602299999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 700000000  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(v_total_pl4)), 2)%></strong></font>
				<%Else If rsRevenue("param_3") = 799999999  Then%>
					<font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount4)), 2)%></strong></font>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
				<%End If%>
			</td> 
			<!-- End To Date Formula -->
		</tr>
<%	 
	tempAmount = 0
	tempAmount2 = 0
	tempAmount3 = 0 
	tempAmount4 = 0 

	v_total_hpp = 0
	v_total_hpp2 = 0
	v_total_hpp3 = 0
	v_total_hpp4 = 0
	v_total_bua = 0
	v_total_bua2 = 0
	v_total_bua3 = 0
	v_total_bua4 = 0
	v_total_bup = 0
	v_total_bup2 = 0
	v_total_bup3 = 0
	v_total_bup4 = 0
	v_total_bo = 0
	v_total_bo2 = 0
	v_total_bo3 = 0
	v_total_bo4 = 0
	v_total_pl = 0
	v_total_pl2 = 0
	v_total_pl3 = 0
	v_total_pl4 = 0

	v_tempAmount = 0
	v_tempAmount2 = 0
	v_tempAmount3 = 0 
%>
<%End If%>
<tr>
	<%If rsRevenue("fccode") <> 402019990 and rsRevenue("fccode") <> 402019991 and rsRevenue("fccode") <>  602199991 and rsRevenue("fccode") <> 602299991  and rsRevenue("fccode") <>799999991 Then %> 
		<td align="left" style="padding-left:10px;">
			<font face="Verdana" size="1">
				<strong><%=Replace(rsRevenue("jobprofitandlosscode"),"_","&nbsp;")%></strong>
			</font>
		</td>
    <td align="left"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT"))),2)%></font></td> 
    <td align="left"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT3"))),2)%></font></td> 
    <td align="left"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT2"))),2)%></font></td> 
    <td align="left"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT4"))),2)%></font></td> 
	<%End If%>
</tr>
<%
	tempAmount = tempAmount + cdbl(to_nol(rsRevenue("AMOUNT")))
	tempAmount2 = tempAmount2 + cdbl(to_nol(rsRevenue("AMOUNT2")))  
	tempAmount3 = tempAmount3 + cdbl(to_nol(rsRevenue("AMOUNT3")))
	tempAmount4 = tempAmount4 + cdbl(to_nol(rsRevenue("AMOUNT4")))  

	If 	rsRevenue("param_3") = 401000000 then  
		total_pb =  tempAmount
		total_pb2 = tempAmount2 
		total_pb3 =  tempAmount3
		total_pb4 = tempAmount4 
		Else If	rsRevenue("param_3") = 402010000 then  
				tempAmount =  tempAmount
				tempAmount2 = tempAmount2
				tempAmount3 =  tempAmount3
				tempAmount4 = tempAmount4

				total_hpp =  tempAmount
				total_hpp2 =  tempAmount2 
				total_hpp3 =  tempAmount3
				total_hpp4 =  tempAmount4 
				Else If	rsRevenue("param_3") = 402010999 then 
					tempAmount =  total_pb - total_hpp
					tempAmount2 =  total_pb2 - total_hpp2 
					tempAmount3 =  total_pb3 - total_hpp3
					tempAmount4 =  total_pb4 - total_hpp4 

					total_lk =  total_pb - total_hpp
					total_lk2 =  total_pb2 - total_hpp2  
					total_lk3 =  total_pb3 - total_hpp3
					total_lk4 =  total_pb4 - total_hpp4  
					Else If	rsRevenue("param_3") = 402020000 then 
						total_bup =  tempAmount
						total_bup2 =  tempAmount2
						total_bup3 =  tempAmount3
						total_bup4 =  tempAmount4
						Else If	rsRevenue("param_3") = 602000000 then 
							total_bua =  tempAmount
							total_bua2 =  tempAmount2
							total_bua3 =  tempAmount3
							total_bua4 =  tempAmount4
							Else If	rsRevenue("param_3") = 602199999 then 
								tempAmount =  total_bup + total_bua
								tempAmount2 =  total_bup2 + total_bua2  
								tempAmount3 =  total_bup3 + total_bua3
								tempAmount4 =  total_bup4 + total_bua4  

								total_bo =  total_bup + total_bua
								total_bo2 =  total_bup2 + total_bua2  
								total_bo3 =  total_bup3 + total_bua3
								total_bo4 =  total_bup4 + total_bua4  
								Else If	rsRevenue("param_3") = 602299999 then 
									tempAmount =  total_lk - total_bo
									tempAmount2 =  total_lk2 - total_bo2
									tempAmount3 =  total_lk3 - total_bo3
									tempAmount4 =  total_lk4 - total_bo4

									total_lo =  total_lk - total_bo
									total_lo2 =  total_lk2 - total_bo2
									total_lo3 =  total_lk3 - total_bo3
									total_lo4 =  total_lk4 - total_bo4
									Else If	rsRevenue("param_3") = 700000000 then  
										total_pl =  tempAmount
										total_pl2 =  tempAmount2 
										total_pl3 =  tempAmount3
										total_pl4 =  tempAmount4 
										Else If	rsRevenue("param_3") = 799999999 then 
											tempAmount =  total_lo - total_pl
											tempAmount2 =  total_lo2 - total_pl2  
											tempAmount3 =  total_lo3 - total_pl3
											tempAmount4 =  total_lo4 - total_pl4  
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If
%>
<%Else%>
<tr>
	<% If rsRevenue("fccode") <> 402019990 and rsRevenue("fccode") <> 402019991 and rsRevenue("fccode") <>  602199991 and rsRevenue("fccode") <> 602299991  and rsRevenue("fccode") <>799999991 Then%> 
		<td align="left" style="padding-left:10px;">
			<font face="Verdana" size="1">
				<strong><%=Replace(rsRevenue("jobprofitandlosscode"),"_","&nbsp;")%></strong>
			</font>
		</td>
    <td align="left"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT"))),2)%></font></td> 
    <td align="left"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT3"))),2)%></font></td> 
    <td align="left"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT2"))),2)%></font></td> 
    <td align="left"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT4"))),2)%></font></td> 
	<%End If%>
</tr>
		
<%
	tempAmount = tempAmount + cdbl(to_nol(rsRevenue("AMOUNT")))
	tempAmount2 = tempAmount2 + cdbl(to_nol(rsRevenue("AMOUNT2")))  
	tempAmount3 = tempAmount3 + cdbl(to_nol(rsRevenue("AMOUNT3")))
	tempAmount4 = tempAmount4 + cdbl(to_nol(rsRevenue("AMOUNT4")))  

	If 	rsRevenue("param_3") = 401000000 then  
		total_pb =  tempAmount
		total_pb2 = tempAmount2 
		total_pb3 =  tempAmount3
		total_pb4 = tempAmount4 
		Else If	rsRevenue("param_3") = 402010000 then  
				tempAmount =  tempAmount
				tempAmount2 = tempAmount2
				tempAmount3 =  tempAmount3
				tempAmount4 = tempAmount4

				total_hpp =  tempAmount
				total_hpp2 =  tempAmount2 
				total_hpp3 =  tempAmount3
				total_hpp4 =  tempAmount4 
				Else If	rsRevenue("param_3") = 402010999 then 
					tempAmount =  total_pb - total_hpp
					tempAmount2 =  total_pb2 - total_hpp2 
					tempAmount3 =  total_pb3 - total_hpp3
					tempAmount4 =  total_pb4 - total_hpp4 

					total_lk =  total_pb - total_hpp
					total_lk2 =  total_pb2 - total_hpp2  
					total_lk3 =  total_pb3 - total_hpp3
					total_lk4 =  total_pb4 - total_hpp4  
					Else If	rsRevenue("param_3") = 402020000 then 
						total_bup =  tempAmount
						total_bup2 =  tempAmount2
						total_bup3 =  tempAmount3
						total_bup4 =  tempAmount4
						Else If	rsRevenue("param_3") = 602000000 then 
							total_bua =  tempAmount
							total_bua2 =  tempAmount2
							total_bua3 =  tempAmount3
							total_bua4 =  tempAmount4
							Else If	rsRevenue("param_3") = 602199999 then 
								tempAmount =  total_bup + total_bua
								tempAmount2 =  total_bup2 + total_bua2  
								tempAmount3 =  total_bup3 + total_bua3
								tempAmount4 =  total_bup4 + total_bua4  

								total_bo =  total_bup + total_bua
								total_bo2 =  total_bup2 + total_bua2  
								total_bo3 =  total_bup3 + total_bua3
								total_bo4 =  total_bup4 + total_bua4  
								Else If	rsRevenue("param_3") = 602299999 then 
									tempAmount =  total_lk - total_bo
									tempAmount2 =  total_lk2 - total_bo2
									tempAmount3 =  total_lk3 - total_bo3
									tempAmount4 =  total_lk4 - total_bo4

									total_lo =  total_lk - total_bo
									total_lo2 =  total_lk2 - total_bo2
									total_lo3 =  total_lk3 - total_bo3
									total_lo4 =  total_lk4 - total_bo4
									Else If	rsRevenue("param_3") = 700000000 then  
										total_pl =  tempAmount
										total_pl2 =  tempAmount2 
										total_pl3 =  tempAmount3
										total_pl4 =  tempAmount4 
										Else If	rsRevenue("param_3") = 799999999 then 
											tempAmount =  total_lo - total_pl
											tempAmount2 =  total_lo2 - total_pl2  
											tempAmount3 =  total_lo3 - total_pl3
											tempAmount4 =  total_lo4 - total_pl4  
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If
%>

<%End If%>
<% 

	rsRevenue.MoveNext
	response.Flush()
	no = no + 1
	Prevweekly= Currweekly
	Wend
	rsRevenue.close()
	rsRevenue=null%>  
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>