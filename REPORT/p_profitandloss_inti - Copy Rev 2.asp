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
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
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

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
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
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut, fcba"
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		whereCons2	=	whereCons2	& " And mp.fcba = '" & v_txtfcba & "'"
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

strMain = 	"   SELECT "& strFieldCons &"" & _
			"                    DECODE (param_1, 0, FCCODE, param_1) param_1," & _
			"                    param_2," & _
			"                    param_3," & _
			"                    param_group," & _
			"                    fccode," & _
			"                    jobprofitandlosscode," & _
			"                    SUM (amount_act) amount_act," & _
			"                    SUM (amount_sbi) amount_sbi," & _
			"                    SUM (amount_prev) amount_prev" & _
			"               FROM (  SELECT STAGE1.fcba," & _
			"                              STAGE1.param_1," & _
			"                              STAGE1.param_2," & _
			"                              prd.param_3," & _
			"                              prd.fcname param_group," & _
			"                              STAGE1.fccode," & _
			"                              STAGE1.jobprofitandlosscode," & _
			"                              SUM (NVL (STAGE2.amount_act, 0)) amount_act," & _
			"                              SUM (NVL (STAGE2.amount_sbi, 0)) amount_sbi," & _
			"                              SUM (NVL (STAGE2.amount_prev, 0)) amount_prev" & _
			"                         FROM (    SELECT fcba," & _
			"                                          fccode," & _
			"                                          param_1," & _
			"                                          param_2," & _
			"                                          param_3," & _
			"                                             LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                          || TO_CHAR (fcname)" & _
			"                                             jobprofitandlosscode" & _
			"                                     FROM iplasprod.parameterdetail" & _
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
			"                                      AND parameterdetail.param_1 <> -1 ) prd" & _
			"                                 ON (    stage1.fcba = prd.fcba" & _
			"                                     AND stage1.param_3 = prd.param_3)" & _
			"                              LEFT JOIN" & _
			"                              (  SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        (-1 * (SUM (tb.debet) + SUM (tb.credit))) amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtyear &"' AND period_month = '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        (-1 * SUM (tb.ending_balance)) amount_sbi," & _
			"                                        0 amount_prev" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtyear &"' AND period_month = '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        (-1 * SUM (tb.ending_balance)) amount_prev" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtyear &"' - 1 AND period_month = 12" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode) STAGE2" & _
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
			"                    SUM (amount_prev) amount_prev" & _
			"               FROM (  SELECT STAGE1.fcba," & _
			"                              STAGE1.param_1," & _
			"                              STAGE1.param_2," & _
			"                              prd.param_3," & _
			"                              prd.fcname param_group," & _
			"                              STAGE1.fccode," & _
			"                              STAGE1.jobprofitandlosscode," & _
			"                              SUM (NVL (STAGE2.amount_act, 0)) amount_act," & _
			"                              SUM (NVL (STAGE2.amount_sbi, 0)) amount_sbi," & _
			"                              SUM (NVL (STAGE2.amount_prev, 0)) amount_prev" & _
			"                         FROM (    SELECT fcba," & _
			"                                          fccode," & _
			"                                          param_1," & _
			"                                          param_2," & _
			"                                          param_3," & _
			"                                             LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                          || TO_CHAR (fcname)" & _
			"                                             jobprofitandlosscode" & _
			"                                     FROM iplasprod.parameterdetail" & _
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
			"                                      AND parameterdetail.param_1 <> -1 ) prd" & _
			"                                 ON (    stage1.fcba = prd.fcba" & _
			"                                     AND stage1.param_3 = prd.param_3)" & _
			"                              LEFT JOIN" & _
			"                              (  SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        ( -1 * (SUM (tb.debet) + SUM (tb.credit)))" & _
			"                                           amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        0 amount_prev" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtyear &"' AND period_month = '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        (-1 * SUM (tb.ending_balance)) amount_sbi," & _
			"                                        0 amount_prev" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtyear &"' AND period_month = '"& v_txtMonth &"'" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode" & _
			"                               UNION ALL" & _
			"                                 SELECT tb.fcba," & _
			"                                        jb.job_forprofitandlosscode," & _
			"                                        0 amount_act," & _
			"                                        0 amount_sbi," & _
			"                                        (-1 * SUM (tb.ending_balance)) amount_prev" & _
			"                                   FROM trialbalance tb" & _
			"                                        INNER JOIN job jb" & _
			"                                           ON (    tb.fcba = jb.fcba" & _
			"                                               AND tb.fccode = jb.fccode)" & _
			"                                  WHERE period_year = '"& v_txtyear &"' - 1 AND period_month = 12" & _
			"                               GROUP BY tb.fcba, jb.job_forprofitandlosscode) STAGE2" & _
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
					"          SUM (amount_act) AMOUNT," & _
					"          SUM (amount_sbi) AMOUNT2," & _
					"          SUM (amount_prev) AMOUNT3" 

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
			"              WHERE     TO_NUMBER (TO_CHAR (mp.tdate, 'mm')) = '"& v_txtMonth &"'" & _
			"                    AND TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy')) = '"& v_txtYear &"'" & _
			"                    AND mp.product_code IN ('CPO', 'KER')" & _
			"					 "& whereCons2 &""& _
			"           GROUP BY mp.fcba, mp.product_code" & _
			"           UNION ALL" & _
			"             SELECT "& strFieldCons2 &"" & _
			"                    mp.product_code," & _
			"                    MAX (TO_NUMBER (TO_CHAR (mp.tdate, 'mm'))) month_prod_bi," & _
			"                    MAX (TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy'))) year_prod_bi," & _
			"                    0 qty_prod_bi," & _
			"                    SUM (mp.production_today) qty_prod_sbi" & _
			"               FROM mill_production mp" & _
			"              WHERE     TO_NUMBER (TO_CHAR (mp.tdate, 'mm')) <= '"& v_txtMonth &"'" & _
			"                    AND TO_NUMBER (TO_CHAR (mp.tdate, 'yyyy')) = '"& v_txtYear &"'" & _
			"                    AND mp.product_code IN ('CPO', 'KER')" & _
			"					 "& whereCons2 &""& _
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

strProdTotal	=	" select GroupPlantation " & selectCons & strProd &"  From (" & strMain3 & " )" & _ 
                    "" & whereCons & _
                    " group by GroupPlantation, " & _
					"          product_code," & _
					"          month_prod_bi," & _
					"          year_prod_bi" & selectCons & ""

strProdTotal = strProdTotal & " Order by GroupPlantation" & orderCons & ""

' response.write strProdTotal

' Response.write strRevenueTotal
%>
<head>
<title>iPla'S Reporting - Statement Of Profit Or Loss And Other Comprehensive Income</title>
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
			<th colspan="<%=var_colspan+10%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header-wrapper">
					<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
					<% if  CompanyName <>  "" then %>
						<h4>PT.<%=CompanyName%></h4> 
					<% End if%>
					<% If  FcbaName <> ""  then %>
						   <h5><%=FcbaName%></h5>   
					<% End if%>
					</div>
				<div id="subheader" class="centered">
					<h3>STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME</h3>
					<br/>
					<h4>For The Month Ended&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
			  </div>
			</th>
		</tr>
		<tr>
		<%

		Set rsProduction = connStr.Execute(strProdTotal)
		if Not rsProduction.eof then
		   nCountofField = rsProduction.fields.count - 1
		end if
		NO=1%>
		</tr>
		<tr>    
			<th align="center" style="width: 400px; color: #fff;">PRODUCTION</th>
			<th align="center" style="width: 135px; color: #fff;">THIS MONTH</th> 
			<th align="center" style="width: 135px; color: #fff;">TO DATE</th> 
		</tr>
	</thead>
	<tbody>
		<%
		v_qty_prod_bi = 0
		v_qty_prod_sbi = 0 
		While Not rsProduction.EOF
		' QUANTITY = cdbl(to_nol(rsProduction("QUANTITY"))) / 1000
		' QUANTITY2 = cdbl(to_nol(rsProduction("QUANTITY2"))) / 1000
		' QUANTITY3 = cdbl(to_nol(rsProduction("QUANTITY3"))) / 1000

		' BUDGETQTY = cdbl(to_nol(rsProduction("BUDGETQTY"))) / 1000
		' BUDGETQTY2 = cdbl(to_nol(rsProduction("BUDGETQTY2"))) / 1000
		' BUDGETQTY3 = cdbl(to_nol(rsProduction("BUDGETQTY3"))) / 1000
		

		' VARIANCE_QUANTITY_THISMONTH = cdbl(to_nol(QUANTITY)) - cdbl(to_nol(BUDGETQTY))
		' VARIANCE_QUANTITY_TODATE = cdbl(to_nol(QUANTITY2)) - cdbl(to_nol(BUDGETQTY2))
	%>
			<tr class="<%=cClass%>"> 
				<td style="width: 425px"><strong><i><%=rsProduction("product_code")%></i></strong></td>
				<td align="right" style="width: 135px;"><strong><i><%=FormatNumber(cdbl(to_nol(rsProduction("qty_prod_bi"))),0)%></i></strong></td>
				<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(to_nol(rsProduction("qty_prod_sbi"))),0)%></i></strong></td> 
			</tr>
		  <%
		 ' 	v_QUANTITY = v_QUANTITY + cdbl(to_nol(rsProduction("QUANTITY"))) / 1000
			' v_QUANTITY2 = v_QUANTITY2 + cdbl(to_nol(rsProduction("QUANTITY2")))	/ 1000
			' v_QUANTITY3 = v_QUANTITY3 + cdbl(to_nol(rsProduction("QUANTITY3")))	/ 1000

			' v_BUDGETQTY = v_BUDGETQTY + cdbl(to_nol(rsProduction("budgetqty"))) / 1000
			' v_BUDGETQTY2 = v_BUDGETQTY2 + cdbl(to_nol(rsProduction("budgetqty2")))	/ 1000
			' v_BUDGETQTY3 = v_BUDGETQTY3 + cdbl(to_nol(rsProduction("budgetqty3")))	/ 1000

			' v_variance_QUANTITY_thismonth = v_variance_QUANTITY_thismonth + cdbl(to_nol(VARIANCE_QUANTITY_THISMONTH))
			' v_variance_QUANTITY_todate = v_variance_QUANTITY_todate + cdbl(to_nol(VARIANCE_QUANTITY_TODATE))	
		  'response.write cdbl(trim(rsProduction("MTHOPENBALANCE")))+1
		  ' v_qty_prod_bi = v_qty_prod_bi + cdbl(to_nol(rsProduction("qty_prod_bi")))
		  ' v_qty_prod_sbi = v_qty_prod_sbi + cdbl(to_nol(rsProduction("qty_prod_sbi")))
		  rsProduction.MoveNext
		  'response.Flush()
		  no=no+1
		Wend
		rsProduction.close()
		rsProduction=null%>
		</form> 
	<tbody>
</table>

<table border="0" width="100%" id="product-table" class="xthf-site">
  <thead>
    <tr> 
		<th align="center" style="color: #fff; width: 400px">SUBJECT</th>
		<th align="center" style="color: #fff; width: 135px">THIS MONTH</th>
		<th align="center" style="color: #fff; width: 135px">EST. THIS MONTH</th>
		<th align="center" style="color: #fff; width: 135px">VARIANCE</th>
		<th align="center" style="color: #fff; width: 135px">TO DATE</th>
		<th align="center" style="color: #fff; width: 135px">EST. TO DATE</th>
		<th align="center" style="color: #fff; width: 135px">VARIANCE</th>
		<th align="center" style="color: #fff; width: 135px">LAST YEAR TO DATE</th>
		<th align="center" style="color: #fff; width: 135px" >EST. FOR THE YEAR</th> 
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
    
    while not rsRevenue.EOF

    AMOUNT = rsRevenue("AMOUNT")
	AMOUNT2 = rsRevenue("AMOUNT2")
	AMOUNT3 = rsRevenue("AMOUNT3")
    Currweekly = rsRevenue("param_group")
	' judul = rsRevenue("param_group")

  	If Currweekly <> Prevweekly Then
%>
<%If Len(Trim(Prevweekly)) = 0 Then

    tempAmount = 0
	tempAmount2 = 0
	tempAmount3 = 0

	v_tempAmount = 0
	v_tempAmount2 = 0
	v_tempAmount3 = 0 
%>
<%Else%>
		<tr style="<%=style%>">
			<td align="left" style="padding-left:10px;"><font face="Verdana" size="1"><strong><%=Replace(UCase(Prevweekly),"_","&nbsp;")%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 0)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 0)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 0)%></strong></font></td>
			<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
		</tr>
<%
	tempAmount = 0
	tempAmount2 = 0
	tempAmount3 = 0

	v_tempAmount = 0
	v_tempAmount2 = 0
	v_tempAmount3 = 0 
%>
<%End If%>
    <tr>  
        <%If rsRevenue("fccode") <> 1510001 AND rsRevenue("fccode") <> 1610001 AND rsRevenue("fccode") <> 2110001 Then %> 
	  		<td align="left"><%=Replace(rsRevenue("jobprofitandlosscode"),"_","&nbsp;")%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT"))),0)%></td> 
			<td align="right"></td> 
			<td align="right"></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT2"))),0)%></td> 
			<td align="right"></td> 
			<td align="right"></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT3"))),0)%></td> 
			<td align="right"></td>       
		<%End If%>
    </tr>
	<%
		tempAmount = tempAmount + cdbl(to_nol(rsRevenue("AMOUNT")))
		tempAmount2 = tempAmount2 + cdbl(to_nol(rsRevenue("AMOUNT2")))
		tempAmount3 = tempAmount3 + cdbl(to_nol(rsRevenue("AMOUNT3")))

		v_tempAmount = v_tempAmount + cdbl(to_nol(rsRevenue("AMOUNT")))
		v_tempAmount2 = v_tempAmount2 + cdbl(to_nol(rsRevenue("AMOUNT2")))
		v_tempAmount3 = v_tempAmount3 + cdbl(to_nol(rsRevenue("AMOUNT3"))) 

		If 	rsRevenue("param_3") = "1100000" then
			style = "background-color: #F4F4A0;"
			total_or =  tempAmount
			total_or2 = tempAmount2
			total_or3 = tempAmount3 

			tempAmount = tempAmount
			tempAmount2 = tempAmount2
			tempAmount3 = tempAmount3 
			Else If	rsRevenue("param_3") = "1210000" then
					style = "background-color: #F4F4A0;"
					total_um =  tempAmount
					total_um2 =  tempAmount2
					total_um3 =  tempAmount3
				Else if rsRevenue("param_3") = "1310000" then  
						style = "background-color: #F4F4A0;"
						total_hc  =  tempAmount
						total_hc2 =  tempAmount2
						total_hc3 =  tempAmount3 
					else if rsRevenue("param_3") = "1410000" then   
							style = "background-color: #F4F4A0;"
							total_est_gc  = tempAmount
							total_est_gc2 = tempAmount2
							total_est_gc3 = tempAmount3
						else if rsRevenue("param_3") = "1510000" then   
								style = "background-color: #f3f33e;"
								tempAmount  = total_um + total_hc + total_est_gc
								tempAmount2 = total_um2 + total_hc2 + total_est_gc2
								tempAmount3 = total_um3 + total_hc3 + total_est_gc3
							else if rsRevenue("param_3") = "1610000" then   
									style = "background-color: #9CB770;"
									tempAmount  = total_or + (total_um + total_hc + total_est_gc)
									tempAmount2  = total_or2 + (total_um2 + total_hc2 + total_est_gc2)
									tempAmount3  = total_or3 + (total_um3 + total_hc3 + total_est_gc3)

									total_gp  = total_or + (total_um + total_hc + total_est_gc)
									total_gp2  = total_or2 + (total_um2 + total_hc2 + total_est_gc2)
									total_gp3  = total_or3 + (total_um3 + total_hc3 + total_est_gc3)
								else if rsRevenue("param_3") = "1710000" then
										style = "background-color: #F4F4A0;"   
										total_oth_ga  = tempAmount
										total_oth_ga2 = tempAmount2
										total_oth_ga3 = tempAmount3
									else if rsRevenue("param_3") = "1810000" then 
											style = "background-color: #F4F4A0;"  
											total_int  = tempAmount
											total_int2 = tempAmount2
											total_int3 = tempAmount3
										else if rsRevenue("param_3") = "1910000" then   
												style = "background-color: #F4F4A0;"
												total_non_or  = tempAmount
												total_non_or2 = tempAmount2
												total_non_or3 = tempAmount3
											else if rsRevenue("param_3") = "2010000" then  
													style = "background-color: #F4F4A0;" 
													total_non_oc  = tempAmount
													total_non_oc2 = tempAmount2
													total_non_oc3 = tempAmount3
												else if rsRevenue("param_3") = "2110000" then   
														style = "background-color: #9CB770;"
														tempAmount  = (total_or + (total_um + total_hc + total_est_gc)) + total_oth_ga + total_int + total_non_or + total_non_oc
														tempAmount2  = (total_or2 + (total_um2 + total_hc2 + total_est_gc2)) + total_oth_ga2 + total_int2 + total_non_or2 + total_non_oc2
														tempAmount3  = (total_or3 + (total_um3 + total_hc3 + total_est_gc3)) + total_oth_ga3 + total_int3 + total_non_or3 + total_non_oc3
													else if rsRevenue("param_3") = "2210000" then   
															style = "background-color: #9CB770;"
															tempAmount  = ((total_or + (total_um + total_hc + total_est_gc)) + total_oth_ga + total_int + total_non_or + total_non_oc) +  v_tempAmount
															tempAmount2  = ((total_or2 + (total_um2 + total_hc2 + total_est_gc2)) + total_oth_ga2 + total_int2 + total_non_or2 + total_non_oc2) + v_tempAmount2
															tempAmount3  = ((total_or3 + (total_um3 + total_hc3 + total_est_gc3)) + total_oth_ga3 + total_int3 + total_non_or3 + total_non_oc3) + v_tempAmount3 

															x_tempAmount  = ((total_or + (total_um + total_hc + total_est_gc)) + total_oth_ga - total_int + total_non_or + total_non_oc) + v_tempAmount
															x_tempAmount2  = ((total_or2 + (total_um2 + total_hc2 + total_est_gc2)) + total_oth_ga2 + total_int2 + total_non_or2 +total_non_oc2) + v_tempAmount2
															x_tempAmount3  = ((total_or3 + (total_um3 + total_hc3 + total_est_gc3)) + total_oth_ga3 + total_int3 + total_non_or3 + total_non_oc3) + v_tempAmount3
														else if rsRevenue("param_3") = "2310000" then 
																style = "background-color: #9CB770;"  
															 	tempAmount  = x_tempAmount + v_tempAmount
																tempAmount2  = x_tempAmount2 + v_tempAmount2
																tempAmount3  = x_tempAmount3 + v_tempAmount3 

																ebf_tempAmount  = x_tempAmount + v_tempAmount
																ebf_tempAmount2  = x_tempAmount2 + v_tempAmount2
																ebf_tempAmount3  = x_tempAmount3 + v_tempAmount3
															else if rsRevenue("param_3") = "2410000" then   
																	style = "background-color: #9CB770;"
																	tempAmount  = (ebf_tempAmount + v_tempAmount)
																	tempAmount2  = (ebf_tempAmount2 + v_tempAmount2)
																	tempAmount3  =  (ebf_tempAmount3 + v_tempAmount3)

																	ebe_tempAmount  = (ebf_tempAmount + v_tempAmount)
																	ebe_tempAmount2  = (ebf_tempAmount2 + v_tempAmount2)
																	ebe_tempAmount3  = (ebf_tempAmount3 + v_tempAmount3)
																else if rsRevenue("param_3") = "2510000" then   
																		style = "background-color: #9CB770;"
																		tempAmount  = (ebe_tempAmount + v_tempAmount)
																		tempAmount2  = (ebe_tempAmount2 + v_tempAmount2)
																		tempAmount3  = (ebe_tempAmount3 + v_tempAmount3)

																		pbtb_tempAmount  = (ebe_tempAmount + v_tempAmount)
																		pbtb_tempAmount2  = (ebe_tempAmount2 + v_tempAmount2)
																		pbtb_tempAmount3  = (ebe_tempAmount3 + v_tempAmount3)
																	else if rsRevenue("param_3") = "2610000" then 
																			style = "background-color: #9CB770;"  
																			tempAmount  = (pbtb_tempAmount + v_tempAmount)
																			tempAmount2  = (pbtb_tempAmount2 + v_tempAmount2)
																			tempAmount3  = (pbtb_tempAmount3 + v_tempAmount3)

																			pbta_tempAmount  = (pbtb_tempAmount + v_tempAmount)
																			pbta_tempAmount2  = (pbtb_tempAmount2 + v_tempAmount2)
																			pbta_tempAmount3  = (pbtb_tempAmount3 + v_tempAmount3)
																		else if rsRevenue("param_3") = "2710000" then   
																				style = "background-color: #9CB770;"
																				tempAmount  = pbta_tempAmount + v_tempAmount
																				tempAmount2  = pbta_tempAmount2 + v_tempAmount2
																				tempAmount3  = pbta_tempAmount3 + v_tempAmount3 
																		End If
																	End If
																End If
															End If
														End If
													End If
												End If
											End If
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
		<%If rsRevenue("fccode") <> 1510001 AND rsRevenue("fccode") <> 1610001 AND rsRevenue("fccode") <> 2110001 Then %> 
	  		<td align="left"><%=Replace(rsRevenue("jobprofitandlosscode"),"_","&nbsp;")%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT"))),0)%></td> 
			<td align="right"></td> 
			<td align="right"></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT2"))),0)%></td> 
			<td align="right"></td> 
			<td align="right"></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsRevenue("AMOUNT3"))),0)%></td> 
			<td align="right"></td>       
		<%End If%>    
	</tr>
    <%
    	tempAmount = tempAmount + cdbl(to_nol(rsRevenue("AMOUNT")))
		tempAmount2 = tempAmount2 + cdbl(to_nol(rsRevenue("AMOUNT2")))
		tempAmount3 = tempAmount3 + cdbl(to_nol(rsRevenue("AMOUNT3")))

		v_tempAmount = v_tempAmount + cdbl(to_nol(rsRevenue("AMOUNT")))
		v_tempAmount2 = v_tempAmount2 + cdbl(to_nol(rsRevenue("AMOUNT2")))
		v_tempAmount3 = v_tempAmount3 + cdbl(to_nol(rsRevenue("AMOUNT3"))) 

		If 	rsRevenue("param_3") = "1100000" then
			style = "background-color: #F4F4A0;"
			total_or =  tempAmount
			total_or2 = tempAmount2
			total_or3 = tempAmount3 

			tempAmount = tempAmount
			tempAmount2 = tempAmount2
			tempAmount3 = tempAmount3 
			Else If	rsRevenue("param_3") = "1210000" then
					style = "background-color: #F4F4A0;"
					total_um =  tempAmount
					total_um2 =  tempAmount2
					total_um3 =  tempAmount3
				Else if rsRevenue("param_3") = "1310000" then  
						style = "background-color: #F4F4A0;"
						total_hc  =  tempAmount
						total_hc2 =  tempAmount2
						total_hc3 =  tempAmount3 
					else if rsRevenue("param_3") = "1410000" then   
							style = "background-color: #F4F4A0;"
							total_est_gc  = tempAmount
							total_est_gc2 = tempAmount2
							total_est_gc3 = tempAmount3
						else if rsRevenue("param_3") = "1510000" then   
								style = "background-color: #f3f33e;"
								tempAmount  = total_um + total_hc + total_est_gc
								tempAmount2 = total_um2 + total_hc2 + total_est_gc2
								tempAmount3 = total_um3 + total_hc3 + total_est_gc3
							else if rsRevenue("param_3") = "1610000" then   
									style = "background-color: #9CB770;"
									tempAmount  = total_or + (total_um + total_hc + total_est_gc)
									tempAmount2  = total_or2 + (total_um2 + total_hc2 + total_est_gc2)
									tempAmount3  = total_or3 + (total_um3 + total_hc3 + total_est_gc3)

									total_gp  = total_or + (total_um + total_hc + total_est_gc)
									total_gp2  = total_or2 + (total_um2 + total_hc2 + total_est_gc2)
									total_gp3  = total_or3 + (total_um3 + total_hc3 + total_est_gc3)
								else if rsRevenue("param_3") = "1710000" then
										style = "background-color: #F4F4A0;"   
										total_oth_ga  = tempAmount
										total_oth_ga2 = tempAmount2
										total_oth_ga3 = tempAmount3
									else if rsRevenue("param_3") = "1810000" then 
											style = "background-color: #F4F4A0;"  
											total_int  = tempAmount
											total_int2 = tempAmount2
											total_int3 = tempAmount3
										else if rsRevenue("param_3") = "1910000" then   
												style = "background-color: #F4F4A0;"
												total_non_or  = tempAmount
												total_non_or2 = tempAmount2
												total_non_or3 = tempAmount3
											else if rsRevenue("param_3") = "2010000" then  
													style = "background-color: #F4F4A0;" 
													total_non_oc  = tempAmount
													total_non_oc2 = tempAmount2
													total_non_oc3 = tempAmount3
												else if rsRevenue("param_3") = "2110000" then   
														style = "background-color: #9CB770;"
														tempAmount  = (total_or + (total_um + total_hc + total_est_gc)) + total_oth_ga + total_int + total_non_or + total_non_oc
														tempAmount2  = (total_or2 + (total_um2 + total_hc2 + total_est_gc2)) + total_oth_ga2 + total_int2 + total_non_or2 + total_non_oc2
														tempAmount3  = (total_or3 + (total_um3 + total_hc3 + total_est_gc3)) + total_oth_ga3 + total_int3 + total_non_or3 + total_non_oc3
													else if rsRevenue("param_3") = "2210000" then   
															style = "background-color: #9CB770;"
															tempAmount  = ((total_or + (total_um + total_hc + total_est_gc)) + total_oth_ga + total_int + total_non_or + total_non_oc) +  v_tempAmount
															tempAmount2  = ((total_or2 + (total_um2 + total_hc2 + total_est_gc2)) + total_oth_ga2 + total_int2 + total_non_or2 + total_non_oc2) + v_tempAmount2
															tempAmount3  = ((total_or3 + (total_um3 + total_hc3 + total_est_gc3)) + total_oth_ga3 + total_int3 + total_non_or3 + total_non_oc3) + v_tempAmount3 

															x_tempAmount  = ((total_or + (total_um + total_hc + total_est_gc)) + total_oth_ga - total_int + total_non_or + total_non_oc) + v_tempAmount
															x_tempAmount2  = ((total_or2 + (total_um2 + total_hc2 + total_est_gc2)) + total_oth_ga2 + total_int2 + total_non_or2 +total_non_oc2) + v_tempAmount2
															x_tempAmount3  = ((total_or3 + (total_um3 + total_hc3 + total_est_gc3)) + total_oth_ga3 + total_int3 + total_non_or3 + total_non_oc3) + v_tempAmount3
														else if rsRevenue("param_3") = "2310000" then 
																style = "background-color: #9CB770;"  
															 	tempAmount  = x_tempAmount + v_tempAmount
																tempAmount2  = x_tempAmount2 + v_tempAmount2
																tempAmount3  = x_tempAmount3 + v_tempAmount3 

																ebf_tempAmount  = x_tempAmount + v_tempAmount
																ebf_tempAmount2  = x_tempAmount2 + v_tempAmount2
																ebf_tempAmount3  = x_tempAmount3 + v_tempAmount3
															else if rsRevenue("param_3") = "2410000" then   
																	style = "background-color: #9CB770;"
																	tempAmount  = (ebf_tempAmount + v_tempAmount)
																	tempAmount2  = (ebf_tempAmount2 + v_tempAmount2)
																	tempAmount3  =  (ebf_tempAmount3 + v_tempAmount3)

																	ebe_tempAmount  = (ebf_tempAmount + v_tempAmount)
																	ebe_tempAmount2  = (ebf_tempAmount2 + v_tempAmount2)
																	ebe_tempAmount3  = (ebf_tempAmount3 + v_tempAmount3)
																else if rsRevenue("param_3") = "2510000" then   
																		style = "background-color: #9CB770;"
																		tempAmount  = (ebe_tempAmount + v_tempAmount)
																		tempAmount2  = (ebe_tempAmount2 + v_tempAmount2)
																		tempAmount3  = (ebe_tempAmount3 + v_tempAmount3)

																		pbtb_tempAmount  = (ebe_tempAmount + v_tempAmount)
																		pbtb_tempAmount2  = (ebe_tempAmount2 + v_tempAmount2)
																		pbtb_tempAmount3  = (ebe_tempAmount3 + v_tempAmount3)
																	else if rsRevenue("param_3") = "2610000" then 
																			style = "background-color: #9CB770;"  
																			tempAmount  = (pbtb_tempAmount + v_tempAmount)
																			tempAmount2  = (pbtb_tempAmount2 + v_tempAmount2)
																			tempAmount3  = (pbtb_tempAmount3 + v_tempAmount3)

																			pbta_tempAmount  = (pbtb_tempAmount + v_tempAmount)
																			pbta_tempAmount2  = (pbtb_tempAmount2 + v_tempAmount2)
																			pbta_tempAmount3  = (pbtb_tempAmount3 + v_tempAmount3)
																		else if rsRevenue("param_3") = "2710000" then   
																				style = "background-color: #9CB770;"
																				tempAmount  = pbta_tempAmount + v_tempAmount
																				tempAmount2  = pbta_tempAmount2 + v_tempAmount2
																				tempAmount3  = pbta_tempAmount3 + v_tempAmount3 
																		End If
																	End If
																End If
															End If
														End If
													End If
												End If
											End If
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
     ' gt_gross = gt_gross + cdbl(to_nol(rsRevenue("gross")))
     ' gt_grading = totalCHEMICAL + cdbl(to_nol(rsRevenue("grading")))
     ' gt_netto  = gt_netto + cdbl(to_nol(rsRevenue("netto")))
     ' gtp_supplying = gtp_supplying + cdbl(to_nol(supplying))
     ' gtp_grading = gtp_grading + cdbl(to_nol(grading))
     rsRevenue.MoveNext
    Prevweekly= Currweekly
  wend 
  %>
	<tr style="<%=style%>">
		<td align="left"  style="padding-left:10px;"><font face="Verdana" size="1"><strong><%=Replace(UCase(Prevweekly),"_","&nbsp;")%></strong></font></td>
		<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount)), 0)%></strong></font></td>
		<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
		<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
		<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount2)), 0)%></strong></font></td>
		<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
		<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
		<td align="Right"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(tempAmount3)), 0)%></strong></font></td>
		<td align="Right"><font face="Verdana" size="1"><strong></strong></font></td>
	</tr>
  </tbody>
</table>
<%If Len(Trim(v_txtFcba)) > 0 And (v_txtFcba <>  Cstr("--SUMMARY--") And v_txtFcba <>  Cstr("--ALL--"))then%>
   <table width="90%" class="data" style="margin:auto;">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=MgrName%>&nbsp;&nbsp;</u></strong></td>
		<td align="center"><font face="Verdana" size="2"><strong><u>&nbsp;&nbsp;<%=KTUName%>&nbsp;&nbsp;</u></strong></td>
	   </tr>
	   	  <tr> 
		<td align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
<%End If%>
   </table>