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
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
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
		selectCons	=	selectCons	& " , nourut"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
	Else
		selectCons	=	selectCons	& " , nourut"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		
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

strMain = 	"   SELECT" & strFieldCons & _
			"                  fcba," & _
			"                  param_1," & _
			"                  param_2," & _
			"                  fccode," & _
			"                  jobprofitandlosscode," & _
			"                  JOBCODE_PROFITANDLOSS," & _
			"                  0 JOBCODE_PROFITANDLOSS2," & _
			"                  0 JOBCODE_PROFITANDLOSS3" & _
			"             FROM (  SELECT STAGE1.fcba," & _
			"                            STAGE1.param_1," & _
			"                            STAGE1.param_2," & _
			"                            STAGE1.fccode," & _
			"                            STAGE1.jobprofitandlosscode," & _
			"                            NVL (STAGE2.amount, 0) JOBCODE_PROFITANDLOSS" & _
			"                       FROM (    SELECT fcba," & _
			"                                        fccode," & _
			"                                        param_1," & _
			"                                        param_2," & _
			"                                           LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                        || TO_CHAR (fcname)" & _
			"                                           jobprofitandlosscode" & _
			"                                   FROM iplasprod.parameterdetail" & _
			"                                  WHERE     parcode = 'JOBCODE_PROFITANDLOSS'" & _
			"                                        AND param_1 <> -1" & _
			"                             START WITH fccode = '0'" & _
			"                             CONNECT BY PRIOR fccode = param_1" & _
			"                               GROUP BY fcba," & _
			"                                        fccode," & _
			"                                        param_1," & _
			"                                        param_2," & _
			"                                           LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                        || TO_CHAR (fcname)) STAGE1" & _
			"                            LEFT JOIN" & _
			"                            (  SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                      LEFT JOIN field ff" & _
			"                                         ON (    ff.fcba = ca.fcba" & _
			"                                             AND ff.fccode = ca.targetcode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &"" & _
			"                                      AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) = "& v_txtMonth &"" & _
			"                                      AND ff.ownership <> 'INTI'" & _
			"                                      AND ca.targettype = 'FF'" & _
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                             UNION ALL" & _
			"                               SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                      LEFT JOIN field ff" & _
			"                                         ON (    ff.fcba = ca.fcba" & _
			"                                             AND ff.fccode = ca.targetcode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &"" & _
			"                                      AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) = "& v_txtMonth &"" & _
			"                                      AND jb.fccode IN" & _
			"                                             (SELECT DISTINCT field_info_03" & _
			"                                                FROM field" & _
			"                                               WHERE ownership <> 'INTI')" & _ 
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                             UNION ALL" & _
			"                               SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &"" & _
			"                                      AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) = "& v_txtMonth &"" & _
			"                                      AND ca.targettype = 'GC'" & _
			"                                      AND ca.targetcode = '04'" & _
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode) STAGE2" & _
			"                               ON (    STAGE1.FCBA = STAGE2.FCBA" & _
			"                                   AND STAGE1.FCCODE =" & _
			"                                          STAGE2.JOB_FORPROFITANDLOSSCODE)" & _
			"                   ORDER BY fccode)" & _
			"           UNION ALL" & _
			"   		SELECT" & strFieldCons & _
			"                  fcba," & _
			"                  param_1," & _
			"                  param_2," & _
			"                  fccode," & _
			"                  jobprofitandlosscode," & _
			"                  0 JOBCODE_PROFITANDLOSS," & _
			"                  JOBCODE_PROFITANDLOSS2," & _
			"                  0 JOBCODE_PROFITANDLOSS3" & _
			"             FROM (  SELECT STAGE1.fcba," & _
			"                            STAGE1.param_1," & _
			"                            STAGE1.param_2," & _
			"                            STAGE1.fccode," & _
			"                            STAGE1.jobprofitandlosscode," & _
			"                            NVL (STAGE2.amount, 0) JOBCODE_PROFITANDLOSS2" & _
			"                       FROM (    SELECT fcba," & _
			"                                        fccode," & _
			"                                        param_1," & _
			"                                        param_2," & _
			"                                           LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                        || TO_CHAR (fcname)" & _
			"                                           jobprofitandlosscode" & _
			"                                   FROM iplasprod.parameterdetail" & _
			"                                  WHERE     parcode = 'JOBCODE_PROFITANDLOSS'" & _
			"                                        AND param_1 <> -1" & _
			"                             START WITH fccode = '0'" & _
			"                             CONNECT BY PRIOR fccode = param_1" & _
			"                               GROUP BY fcba," & _
			"                                        fccode," & _
			"                                        param_1," & _
			"                                        param_2," & _
			"                                           LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                        || TO_CHAR (fcname)) STAGE1" & _
			"                            LEFT JOIN" & _
			"                            (  SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                      LEFT JOIN field ff" & _
			"                                         ON (    ff.fcba = ca.fcba" & _
			"                                             AND ff.fccode = ca.targetcode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &"" & _
			"                                      AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) <=" & _
			"                                             "& v_txtMonth &"" & _
			"                                      AND ff.ownership <> 'INTI'" & _
			"                                      AND ca.targettype = 'FF'" & _
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                             UNION ALL" & _
			"                               SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                      LEFT JOIN field ff" & _
			"                                         ON (    ff.fcba = ca.fcba" & _
			"                                             AND ff.fccode = ca.targetcode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &"" & _
			"                                      AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) <=" & _
			"                                             "& v_txtMonth &"" & _
			"                                      AND jb.fccode IN" & _
			"                                             (SELECT DISTINCT field_info_03" & _
			"                                                FROM field" & _
			"                                               WHERE ownership <> 'INTI')" & _ 
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                             UNION ALL" & _
			"                               SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &"" & _
			"                                      AND TO_NUMBER (TO_CHAR (ca.tdate, 'mm')) <=" & _
			"                                             "& v_txtMonth &"" & _
			"                                      AND ca.targettype = 'GC'" & _
			"                                      AND ca.targetcode = '04'" & _
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode) STAGE2" & _
			"                               ON (    STAGE1.FCBA = STAGE2.FCBA" & _
			"                                   AND STAGE1.FCCODE =" & _
			"                                          STAGE2.JOB_FORPROFITANDLOSSCODE)" & _
			"                   ORDER BY fccode)" & _
			"           UNION ALL" & _
			"   		SELECT" & strFieldCons & _
			"                  fcba," & _
			"                  param_1," & _
			"                  param_2," & _
			"                  fccode," & _
			"                  jobprofitandlosscode," & _
			"                  0 JOBCODE_PROFITANDLOSS," & _
			"                  0 JOBCODE_PROFITANDLOSS2," & _
			"                  JOBCODE_PROFITANDLOSS3" & _
			"             FROM (  SELECT STAGE1.fcba," & _
			"                            STAGE1.param_1," & _
			"                            STAGE1.param_2," & _
			"                            STAGE1.fccode," & _
			"                            STAGE1.jobprofitandlosscode," & _
			"                            NVL (STAGE2.amount, 0) JOBCODE_PROFITANDLOSS3" & _
			"                       FROM (    SELECT fcba," & _
			"                                        fccode," & _
			"                                        param_1," & _
			"                                        param_2," & _
			"                                           LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                        || TO_CHAR (fcname)" & _
			"                                           jobprofitandlosscode" & _
			"                                   FROM iplasprod.parameterdetail" & _
			"                                  WHERE     parcode = 'JOBCODE_PROFITANDLOSS'" & _
			"                                        AND param_1 <> -1" & _
			"                             START WITH fccode = '0'" & _
			"                             CONNECT BY PRIOR fccode = param_1" & _
			"                               GROUP BY fcba," & _
			"                                        fccode," & _
			"                                        param_1," & _
			"                                        param_2," & _
			"                                           LPAD ('_', 5 * (LEVEL - 1), '_')" & _
			"                                        || TO_CHAR (fcname)) STAGE1" & _
			"                            LEFT JOIN" & _
			"                            (  SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                      LEFT JOIN field ff" & _
			"                                         ON (    ff.fcba = ca.fcba" & _
			"                                             AND ff.fccode = ca.targetcode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &" - 1" & _
			"                                      AND ff.ownership <> 'INTI'" & _
			"                                      AND ca.targettype = 'FF'" & _
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                             UNION ALL" & _
			"                               SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                      LEFT JOIN field ff" & _
			"                                         ON (    ff.fcba = ca.fcba" & _
			"                                             AND ff.fccode = ca.targetcode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &" - 1" & _
			"                                      AND jb.fccode IN" & _
			"                                             (SELECT DISTINCT field_info_03" & _
			"                                                FROM field" & _
			"                                               WHERE ownership <> 'INTI')" & _ 
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode" & _
			"                             UNION ALL" & _
			"                               SELECT ca.fcba," & _
			"                                      jb.job_forprofitandlosscode," & _
			"                                      SUM (ca.amount) amount" & _
			"                                 FROM iplasprod.cost_activity ca" & _
			"                                      INNER JOIN iplasprod.job jb" & _
			"                                         ON (    ca.fcba = jb.fcba" & _
			"                                             AND ca.jobcode = jb.fccode)" & _
			"                                WHERE     TO_NUMBER (TO_CHAR (ca.tdate, 'yyyy')) =" & _
			"                                             "& v_txtYear &" - 1" & _
			"                                      AND ca.targettype = 'GC'" & _
			"                                      AND ca.targetcode = '04'" & _
			"                             GROUP BY ca.fcba, jb.job_forprofitandlosscode) STAGE2" & _
			"                               ON (    STAGE1.FCBA = STAGE2.FCBA" & _
			"                                   AND STAGE1.FCCODE =" & _
			"                                          STAGE2.JOB_FORPROFITANDLOSSCODE)" & _
			"                   ORDER BY fccode)" 

if  len(trim(v_txtjobgroup)) > 0 then
    strMain = strMain & " and job_api.get_jobgroupCode(fccode, fcba) = " & v_txtjobgroup & ""
end if
if  len(trim(v_txtjobcode)) > 0 then
    strMain = strMain & " and fccode = '" & v_txtjobcode & "'"
end if 
	' strMain = strMain & " group by fcba" & _
 '                        "          , fccode" 
							
	strRevenue = 	"         ,decode(param_1,0,FCCODE,param_1)param_1," & _
					"         param_2," & _
					"         fccode," & _
					"         jobprofitandlosscode," & _
					"         SUM (JOBCODE_PROFITANDLOSS) AMOUNT," & _
					"         SUM (JOBCODE_PROFITANDLOSS2) AMOUNT2," & _
					"         SUM (JOBCODE_PROFITANDLOSS3) AMOUNT3" 

strRevenueTotal	=	" select GroupPlantation " & selectCons & strRevenue &"  From (" & strMain & " )" & _ 
                    "" & whereCons & _
                    " group by GroupPlantation" & selectCons & " , param_1, param_2, fccode, jobprofitandlosscode"

strRevenueTotal = strRevenueTotal & " Order by fccode, GroupPlantation" & orderCons & ", param_1, param_2, jobprofitandlosscode"

strMain3 = 	"  SELECT " & strFieldCons &"," & _
			"         SUM (mill_weight) quantity," & _
			"         0 quantity2," & _
			"         0 quantity3" & _
			"    FROM (  SELECT salesshipmentdtl.fcba," & _
			"                   salesagreementdet.salescontractno," & _
			"                   salesagreementhdr.customercode," & _
			"                   salesshipmentdtl.salesshipmentno," & _
			"                   salesshipmenthdr.creationdate," & _
			"                   harvesting.chitno," & _
			"                   harvesting.fieldcode," & _
			"                   field.ownership," & _
			"                   field.field_info_01 AS ownership_info," & _
			"                   salesshipmentdtl.postingdate," & _
			"                   SUM (harvesting.mill_weight) AS mill_weight" & _
			"              FROM salesshipmentdtl" & _
			"                   INNER JOIN salesshipmenthdr" & _
			"                      ON (    salesshipmentdtl.fcba = salesshipmenthdr.fcba" & _
			"                          AND salesshipmentdtl.salesshipmentno =" & _
			"                                 salesshipmenthdr.salesshipmentno)" & _
			"                   LEFT JOIN deliveryorder" & _
			"                      ON (    salesshipmenthdr.fcba = deliveryorder.fcba" & _
			"                          AND UPPER (TRIM (salesshipmenthdr.deliveryorderno)) =" & _
			"                                 UPPER (TRIM (deliveryorder.deliveryorderno)))" & _
			"                   INNER JOIN salesagreementdet" & _
			"                      ON (    deliveryorder.fcba = salesagreementdet.fcba" & _
			"                          AND deliveryorder.salescontractno =" & _
			"                                 salesagreementdet.salescontractno)" & _
			"                   INNER JOIN salesagreementhdr" & _
			"                      ON (    salesagreementdet.fcba = salesagreementhdr.fcba" & _
			"                          AND salesagreementdet.salescontractno =" & _
			"                                 salesagreementhdr.salescontractno)" & _
			"                   FULL JOIN harvesting" & _
			"                      ON (    salesshipmentdtl.fcba = harvesting.fcba" & _
			"                          AND salesshipmentdtl.customer_chitno =" & _
			"                                 harvesting.chitno)" & _
			"                   INNER JOIN field" & _
			"                      ON (    harvesting.fcba = field.fcba" & _
			"                          AND harvesting.fieldcode = field.fccode)" & _
			"             WHERE TO_NUMBER (TO_CHAR (salesshipmentdtl.postingdate, 'YYYY')) = '"& v_txtYear &"'" & _
			"                   AND TO_NUMBER (TO_CHAR (salesshipmentdtl.postingdate, 'MM')) ='"& v_txtMonth &"'" & _
			"					AND field.ownership = 'INTI'" & _
			"          GROUP BY salesshipmentdtl.fcba," & _
			"                   salesagreementdet.salescontractno," & _
			"                   salesagreementhdr.customercode," & _
			"                   salesshipmentdtl.salesshipmentno," & _
			"                   salesshipmenthdr.creationdate," & _
			"                   harvesting.chitno," & _
			"                   harvesting.fieldcode," & _
			"                   field.ownership," & _
			"                   field.field_info_01," & _
			"                   salesshipmentdtl.postingdate)" & _
			" GROUP BY fcba" & _
			" UNION ALL" & _
			" SELECT " & strFieldCons &"," & _
			"         0 quantity," & _
			"         SUM (mill_weight) quantity2," & _
			"         0 quantity3" & _
			"    FROM (  SELECT salesshipmentdtl.fcba," & _
			"                   salesagreementdet.salescontractno," & _
			"                   salesagreementhdr.customercode," & _
			"                   salesshipmentdtl.salesshipmentno," & _
			"                   salesshipmenthdr.creationdate," & _
			"                   harvesting.chitno," & _
			"                   harvesting.fieldcode," & _
			"                   field.ownership," & _
			"                   field.field_info_01 AS ownership_info," & _
			"                   salesshipmentdtl.postingdate," & _
			"                   SUM (harvesting.mill_weight) AS mill_weight" & _
			"              FROM salesshipmentdtl" & _
			"                   INNER JOIN salesshipmenthdr" & _
			"                      ON (    salesshipmentdtl.fcba = salesshipmenthdr.fcba" & _
			"                          AND salesshipmentdtl.salesshipmentno =" & _
			"                                 salesshipmenthdr.salesshipmentno)" & _
			"                   LEFT JOIN deliveryorder" & _
			"                      ON (    salesshipmenthdr.fcba = deliveryorder.fcba" & _
			"                          AND UPPER (TRIM (salesshipmenthdr.deliveryorderno)) =" & _
			"                                 UPPER (TRIM (deliveryorder.deliveryorderno)))" & _
			"                   INNER JOIN salesagreementdet" & _
			"                      ON (    deliveryorder.fcba = salesagreementdet.fcba" & _
			"                          AND deliveryorder.salescontractno =" & _
			"                                 salesagreementdet.salescontractno)" & _
			"                   INNER JOIN salesagreementhdr" & _
			"                      ON (    salesagreementdet.fcba = salesagreementhdr.fcba" & _
			"                          AND salesagreementdet.salescontractno =" & _
			"                                 salesagreementhdr.salescontractno)" & _
			"                   FULL JOIN harvesting" & _
			"                      ON (    salesshipmentdtl.fcba = harvesting.fcba" & _
			"                          AND salesshipmentdtl.customer_chitno =" & _
			"                                 harvesting.chitno)" & _
			"                   INNER JOIN field" & _
			"                      ON (    harvesting.fcba = field.fcba" & _
			"                          AND harvesting.fieldcode = field.fccode)" & _
			"             WHERE  TO_NUMBER (TO_CHAR (salesshipmentdtl.postingdate, 'YYYY')) = '"& v_txtYear &"'" & _
			"                   AND TO_NUMBER (TO_CHAR (salesshipmentdtl.postingdate, 'MM')) <='"& v_txtMonth &"'" & _
			"					AND field.ownership = 'INTI'" & _
			"          GROUP BY salesshipmentdtl.fcba," & _
			"                   salesagreementdet.salescontractno," & _
			"                   salesagreementhdr.customercode," & _
			"                   salesshipmentdtl.salesshipmentno," & _
			"                   salesshipmenthdr.creationdate," & _
			"                   harvesting.chitno," & _
			"                   harvesting.fieldcode," & _
			"                   field.ownership," & _
			"                   field.field_info_01," & _
			"                   salesshipmentdtl.postingdate)" & _
			" GROUP BY fcba" & _
			" UNION ALL" & _
			" SELECT " & strFieldCons &"," & _
			"         0 quantity," & _
			"         0 quantity2," & _
			"         SUM (mill_weight) quantity3" & _
			"    FROM (  SELECT salesshipmentdtl.fcba," & _
			"                   salesagreementdet.salescontractno," & _
			"                   salesagreementhdr.customercode," & _
			"                   salesshipmentdtl.salesshipmentno," & _
			"                   salesshipmenthdr.creationdate," & _
			"                   harvesting.chitno," & _
			"                   harvesting.fieldcode," & _
			"                   field.ownership," & _
			"                   field.field_info_01 AS ownership_info," & _
			"                   salesshipmentdtl.postingdate," & _
			"                   SUM (harvesting.mill_weight) AS mill_weight" & _
			"              FROM salesshipmentdtl" & _
			"                   INNER JOIN salesshipmenthdr" & _
			"                      ON (    salesshipmentdtl.fcba = salesshipmenthdr.fcba" & _
			"                          AND salesshipmentdtl.salesshipmentno =" & _
			"                                 salesshipmenthdr.salesshipmentno)" & _
			"                   LEFT JOIN deliveryorder" & _
			"                      ON (    salesshipmenthdr.fcba = deliveryorder.fcba" & _
			"                          AND UPPER (TRIM (salesshipmenthdr.deliveryorderno)) =" & _
			"                                 UPPER (TRIM (deliveryorder.deliveryorderno)))" & _
			"                   INNER JOIN salesagreementdet" & _
			"                      ON (    deliveryorder.fcba = salesagreementdet.fcba" & _
			"                          AND deliveryorder.salescontractno =" & _
			"                                 salesagreementdet.salescontractno)" & _
			"                   INNER JOIN salesagreementhdr" & _
			"                      ON (    salesagreementdet.fcba = salesagreementhdr.fcba" & _
			"                          AND salesagreementdet.salescontractno =" & _
			"                                 salesagreementhdr.salescontractno)" & _
			"                   FULL JOIN harvesting" & _
			"                      ON (    salesshipmentdtl.fcba = harvesting.fcba" & _
			"                          AND salesshipmentdtl.customer_chitno =" & _
			"                                 harvesting.chitno)" & _
			"                   INNER JOIN field" & _
			"                      ON (    harvesting.fcba = field.fcba" & _
			"                          AND harvesting.fieldcode = field.fccode)" & _
			"             WHERE TO_NUMBER (TO_CHAR (salesshipmentdtl.postingdate, 'YYYY')) = '"& v_txtYear &"' - 1" & _
			"					AND field.ownership = 'INTI'" & _
			"          GROUP BY salesshipmentdtl.fcba," & _
			"                   salesagreementdet.salescontractno," & _
			"                   salesagreementhdr.customercode," & _
			"                   salesshipmentdtl.salesshipmentno," & _
			"                   salesshipmenthdr.creationdate," & _
			"                   harvesting.chitno," & _
			"                   harvesting.fieldcode," & _
			"                   field.ownership," & _
			"                   field.field_info_01," & _
			"                   salesshipmentdtl.postingdate)" & _
			" GROUP BY fcba" 

if  len(trim(v_txtjobgroup)) > 0 then
    strMain3 = strMain3 & " and job_api.get_jobgroupCode(fccode, fcba) = " & v_txtjobgroup & ""
end if
if  len(trim(v_txtjobcode)) > 0 then
    strMain3 = strMain3 & " and fccode = '" & v_txtjobcode & "'"
end if 
	' strMain3 = strMain3 & " group by fcba" & _
 '                        "          , fccode" 

								
strProd = 	"  , SUM(quantity) quantity, SUM(quantity2) quantity2, SUM(quantity3) quantity3" 

strProdTotal	=	" select GroupPlantation " & selectCons & strProd &"  From (" & strMain3 & " )" & _ 
                    "" & whereCons & _
                    " group by GroupPlantation" & selectCons & ""

strProdTotal = strProdTotal & " Order by GroupPlantation" & orderCons & ""

' response.write strProdTotal

' Response.write strRevenueTotal
%>
<head>
<title>ePlaS Reporting - Profit And Loss</title>
<script language="JavaScript" src="/js/date-picker.js"></script> 
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
				<!--#include virtual="/REPORT/FormSelectNewSum.html"-->
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
					<h2>PROFIT AND LOSS</h2>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
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
			<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
				<th style="border-right: double; width: 135px"><%=titleGroupPlantation%></th>
			<%End If%>

			<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
				<th style="border-right: double; width: 135px"><%=titlePlantation%></th>
			<%End If%>

			<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
				<th style="border-right: double; width: 135px"><%=titleGroupRegion%></th>
			<%End If%>

			<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
				<th style="border-right: double; width: 135px"><%=titleRegion%></th>
			<%End If%>

				<th align="center" colspan="2" style="border-right: double;  width: 400px">PRODUCTION</th>
				<th align="center" style="border-right: double; width: 135px">THIS MONTH</th>
				<th align="center" style="border-right: double; width: 135px">EST. THIS MONTH</th>
				<th align="center" style="border-right: double; width: 135px">VARIANCE</th>
				<th align="center" style="border-right: double; width: 135px">TO DATE</th>
				<th align="center" style="border-right: double; width: 135px">EST. TO DATE</th>
				<th align="center" style="border-right: double; width: 135px">VARIANCE</th>
				<th align="center" style="border-right: double; width: 135px">LAST YEAR TO DATE</th>
				<th align="center" style=" width: 135px" >EST. FOR THE YEAR</th>
				<tr><td><strong><i></td></i></strong></tr>	
		</tr>
	</thead>
	<tbody>
		<%
		v_QUANTITY = 0
		v_variance_QUANTITY_thismonth = 0
		v_variance_QUANTITY_todate = 0
		v_QUANTITY2 = 0
		v_QUANTITY3 = 0
		While Not rsProduction.EOF
		QUANTITY = cdbl(to_nol(rsProduction("QUANTITY"))) / 1000
		VARIANCE_QUANTITY_THISMONTH = cdbl(to_nol(QUANTITY)) - cdbl(to_nol(QUANTITY))
		QUANTITY2 = cdbl(to_nol(rsProduction("QUANTITY2"))) / 1000
		VARIANCE_QUANTITY_TODATE = cdbl(to_nol(QUANTITY)) - cdbl(to_nol(QUANTITY))
		QUANTITY3 = cdbl(to_nol(rsProduction("QUANTITY3"))) / 1000
	%>
			<tr class="<%=cClass%>">
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
						<% tempGroupPlantation = rsProduction("GroupPlantation")
						   tempPlantation = "--ALL--"
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempGroupPlantation <> SprevGroupPlantation then %>
							<td align="center" style="border-right: double; width: 135px">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
								<%=rsProduction("GroupPlantation")%>
								</a>
							</td>
							<% SprevGroupPlantation = tempGroupPlantation %>	
						<% else	%>
							<td align="center" style="border-right: double; width: 135px"></td>					
						<% end if %>
					<% end if %>
					
					
					<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
						      tempGroupPlantation = rsProduction("GroupPlantation")
						   else
						      tempGroupPlantation = ""
						   end if
						   
						   tempPlantation = rsProduction("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempPlantation <> SprevPlantation then %>
							<td align="center" style=" border-bottom:none;">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
								<%=rsProduction("Plantation")%>
								</a>
							</td>
							<% SprevPlantation = tempPlantation %>	
						<% else	%>
							<td align="center" style="border-right: double; width: 135px"></td>					
						<% end if %>	
					<% end if %>
					
					<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
						      tempGroupPlantation = rsProduction("GroupPlantation")
						   else
						      tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
						      tempPlantation = rsProduction("Plantation")
						   else
						      tempPlantation = ""
						   end if
						   
						   tempGroupRegion = rsProduction("GroupRegion")
						   tempRegion = "--ALL--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempGroupRegion <> SprevGroupRegion then %>
							<td align="center" style="border-right: double; width: 135px">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
								<%=rsProduction("GroupRegion")%>
								</a>
							</td>
							<% SprevGroupRegion= tempGroupRegion %>	
						<% else	%>
							<td align="center" style="border-right: double; width: 135px"></td>					
						<% end if %>	
					<% end if %>
					
					<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
					    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
						      tempGroupPlantation = rsProduction("GroupPlantation")
						   else
						      tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
						      tempPlantation = rsProduction("Plantation")
						   else
						      tempPlantation = ""
						   end if
						   
						   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
						       tempGroupRegion = rsProduction("GroupRegion")
						   else
						       tempGroupRegion = ""
						   end if
						   
						   tempRegion = rsProduction("Region") 
						   tempCompany ="--ALL--"
						   tempfcba = "--SUMMARY--"
						%>
						
						<% if tempRegion <> SprevRegion then %>
							<td align="center" style="border-right: double; width: 135px">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
								<%=rsProduction("region")%>
								</a>
							</td>
							<% SprevRegion = tempRegion %>	
						<% else	%>
							<td align="center" style="border-right: double; width: 135px"></td>					
						<% end if %>	
					<% end if %>
					
						<td></td>
						<td style="border-right: double; width: 425px"><strong><i>FFB PRODUCTION (MT)</i></strong></td>
						<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(quantity),2)%></i></strong></td>
						<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(quantity),2)%></i></strong></td>
						<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(to_nol(VARIANCE_QUANTITY_THISMONTH)),2)%></i></strong></td>
						<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(quantity2),2)%></i></strong></td>
						<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(quantity2),2)%></i></strong></td>
						<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(to_nol(VARIANCE_QUANTITY_TODATE)),2)%></i></strong></td>
						<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(quantity3),2)%></i></strong></td>
						<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(quantity3),2)%></i></strong></td>
						<td align="right" ><strong><i></i></strong></td>
				</tr>
		  <%
		 	v_QUANTITY = v_QUANTITY + cdbl(to_nol(rsProduction("QUANTITY"))) / 1000
			v_QUANTITY2 = v_QUANTITY2 + cdbl(to_nol(rsProduction("QUANTITY2")))	/ 1000
			v_QUANTITY3 = v_QUANTITY3 + cdbl(to_nol(rsProduction("QUANTITY3")))	/ 1000
			v_variance_QUANTITY_thismonth = v_variance_QUANTITY_thismonth + cdbl(to_nol(VARIANCE_QUANTITY_THISMONTH))
			v_variance_QUANTITY_todate = v_variance_QUANTITY_todate + cdbl(to_nol(VARIANCE_QUANTITY_TODATE))	
		  'response.write cdbl(trim(rsProduction("MTHOPENBALANCE")))+1
		  rsProduction.MoveNext
		  'response.Flush()
		  no=no+1
		Wend
		rsProduction.close()
		rsProduction=null%>
		</form>
		<tr>
		<th align="Center" colspan="<%=var_colspan+1%>"> <font face="Verdana" size="1">T o t a l</th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_QUANTITY,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_QUANTITY,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_variance_QUANTITY_thismonth,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_QUANTITY2,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_QUANTITY2,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_variance_QUANTITY_todate,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_QUANTITY3,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_QUANTITY3,2)%></th>
		</tr>
	<tbody>
</table>

<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
	<thead >
	<tr>
	<%
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

	Set rsRevenue = connStr.Execute(strRevenueTotal)
	if Not rsRevenue.eof then
	   nCountofField = rsRevenue.fields.count - 1
	end if
	NO=1%>
	</tr>
	<tr>
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th style="border-right: double; width: 135px"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th style="border-right: double; width: 135px"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th style="border-right: double; width: 135px"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th style="border-right: double; width: 135px"><%=titleRegion%></th>
		<%End If%>
			<th align="center" colspan="2" style="border-right: double;  width: 400px">SUBJECT</th>
			<th align="center" style="border-right: double; width: 135px">THIS MONTH</th>
			<th align="center" style="border-right: double; width: 135px">EST. THIS MONTH</th>
			<th align="center" style="border-right: double; width: 135px">VARIANCE</th>
			<th align="center" style="border-right: double; width: 135px">TO DATE</th>
			<th align="center" style="border-right: double; width: 135px">EST. TO DATE</th>
			<th align="center" style="border-right: double; width: 135px">VARIANCE</th>
			<th align="center" style="border-right: double; width: 135px">LAST YEAR TO DATE</th>
			<th align="center" style=" width: 135px" >EST. FOR THE YEAR</th>
			<tr><td><strong><i></td></i></strong></tr>	
	</tr>
</thead>
<tbody>
	<%
	v_amountrevenue = 0
	v_variance_revenue_thismonth = 0
	v_variance_revenue_todate = 0
	v_amountrevenue2 = 0
	v_amountrevenue3 = 0
	tempAmount = 0 
	tempAmount2 = 0
	tempAmount3 = 0
	tempAmount_THISMONTH = 0
	tempAmount2_TODATE = 0
	tempVARIANCE_THISMONTH = 0
	tempVARIANCE_TODATE = 0 
	param_1 = "X"

	While Not rsRevenue.EOF
		AMOUNT = rsRevenue("AMOUNT")
		VARIANCE_THISMONTH = cdbl(to_nol(rsRevenue("AMOUNT"))) - cdbl(to_nol(rsRevenue("AMOUNT")))
		AMOUNT2 = rsRevenue("AMOUNT2")
		VARIANCE_TODATE = cdbl(to_nol(rsRevenue("AMOUNT2"))) - cdbl(to_nol(rsRevenue("AMOUNT2")))
		AMOUNT3 = rsRevenue("AMOUNT3") 

	if (param_1 = "X") then
		param_1 = rsRevenue("PARAM_1")
	end if
%>
<% 
			tempAmount = tempAmount + cdbl(rsRevenue("AMOUNT"))
			tempAmount2 = tempAmount2 + cdbl(rsRevenue("AMOUNT2"))
			tempAmount3 = tempAmount3 + cdbl(rsRevenue("AMOUNT3"))
			' tempVARIANCE_THISMONTH = cdbl(to_nol(tempAmount)) - cdbl(to_nol(tempAmount))
			' tempVARIANCE_TODATE = cdbl(to_nol(tempAmount2)) - cdbl(to_nol(tempAmount2))

			' tempAmount_THISMONTH = tempAmount_THISMONTH + cdbl(to_nol(tempVARIANCE_THISMONTH))
			' tempAmount2_TODATE = tempAmount2_TODATE + cdbl(to_nol(tempVARIANCE_TODATE))

			total_estate_thismonth = cdbl(rsRevenue("AMOUNT")) - tempAmount
			total_estate_todate =  cdbl(rsRevenue("AMOUNT2")) - tempAmount2
			total_estate_lastyear =  cdbl(rsRevenue("AMOUNT3")) - tempAmount3
			
			if param_1 <> rsRevenue("PARAM_1") then
				' Judul Total 
					if param_1 = "4200000" then  result = "TOTAL OTHER REVENUE / EXPENSES"
					if param_1 = "5100000" then  result = "TOTAL ESTATE EXPENDITURE"
				' End Judul Total
				' response.write total_estate
		%>
		<tr>
			<% 	if param_1 = "4100000" then 
					result = ""
			%>
				<td colspan="2" style="border-right: double; width: 600px;"><strong><i> </i></strong></td>
				<td align="right" ><strong><i> </i></strong></td>
				<td align="right" ><strong><i> </i></strong></td>
				<td align="right" ><strong><i> </i></strong></td>
			<%else%>
				<td colspan="2" style="border-right: double; width: 600px; padding-left: 30px; border-bottom: double; border-top: double;"><strong><i><%=result%></i></strong></td>
				<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount,2) %></i></strong></td>
				<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount,2) %></i></strong></td>
				<td align="right" style="border-right: double; width: 250px; border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount_THISMONTH,2) %></i></strong></td>

				<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount2,2) %></i></strong></td>
				<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount2,2) %></i></strong></td>
				<td align="right" style="border-right: double; width: 250px; border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount2_TODATE,2) %></i></strong></td>

				<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount3,2) %></i></strong></td>
				<td align="right" style=" width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount3,2) %></i></strong></td>
			<%end if %>

		</tr>

		<tr>
			<% 	if param_1 = "5100000" then
					result  = "ESTATE PROFIT / (LOSS)" 
			%>
			<td colspan="2" style="border-right: double; width: 600px; padding-left: 30px; border-bottom: double; border-top: double;"><strong><i><%=result%></i></strong></td>
			<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(total_estate_thismonth,2) %></i></strong></td>
			<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(total_estate_thismonth,2) %></i></strong></td>
			<td align="right" style="border-right: double; width: 250px; border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount_THISMONTH,2) %></i></strong></td>

			<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(total_estate_todate,2) %></i></strong></td>
			<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(total_estate_todate,2) %></i></strong></td>
			<td align="right" style="border-right: double; width: 250px; border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount2_TODATE,2) %></i></strong></td>

			<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(total_estate_lastyear,2) %></i></strong></td>
			<td align="right" style=" width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(total_estate_lastyear,2) %></i></strong></td>
			<%end if %> 
		</tr>
			<% 	 
				param_1 = rsRevenue("PARAM_1")
				tempAmount = 0
				tempAmount2 = 0
				tempAmount3 = 0
				total_estate  = 0
			end if 
		
			%>
		<tr class="<%=cClass%>">
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsRevenue("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center" style="border-right: double; width: 135px">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rsRevenue("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center" style="border-right: double; width: 135px"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRevenue("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsRevenue("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center" style=" border-bottom:none;">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rsRevenue("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center" style="border-right: double; width: 135px"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRevenue("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsRevenue("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsRevenue("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center" style="border-right: double; width: 135px">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rsRevenue("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center" style="border-right: double; width: 135px"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsRevenue("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsRevenue("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsRevenue("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsRevenue("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center" style="border-right: double; width: 135px">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtjobgroup=<%=v_txtjobgroup%>&txtjobcode=<%=v_txtjobcode%>&no-null=<%=v_txtnonull%>">
							<%=rsRevenue("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center" style="border-right: double; width: 135px"></td>					
					<% end if %>	
				<% end if %>
				
					<td></td>
					<td style="border-right: double; width: 800px"><strong><i><%=Replace(rsRevenue("jobprofitandlosscode"),"_","&nbsp;")%></i></strong></td>
				<% if cdbl(to_nol(AMOUNTREVENUE)) = 0 AND cdbl(to_nol(AMOUNTREVENUE2)) = 0 AND cdbl(to_nol(param_1)) = 0  then %>
					<td align="right" style="width: 135px"><strong><i></i></strong></td>
					<td align="right" style="width: 135px"><strong><i></i></strong></td>
					<td align="right" style="width: 135px"><strong><i></i></strong></td>
					<td align="right" style="width: 135px"><strong><i></i></strong></td>
					<td align="right" style="width: 135px"><strong><i></i></strong></td>
					<td align="right" style="width: 135px"><strong><i></i></strong></td>
					<td align="right" style="width: 135px"><strong><i></i></strong></td>
					<td align="right" style="width: 135px"><strong><i></i></strong></td>
					<td align="right"><strong><i></i></strong></td>
				<%else%>
					<td align="right" style="width: 135px"><strong><i><%=FormatNumber(rsRevenue("AMOUNT"),2)%></i></strong></td>
					<td align="right" style="width: 135px"><strong><i><%=FormatNumber(rsRevenue("AMOUNT"),2)%></i></strong></td>
					<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(to_nol(VARIANCE_THISMONTH)),2)%></i></strong></td>
					<td align="right" style="width: 135px"><strong><i><%=FormatNumber(rsRevenue("AMOUNT2"),2)%></i></strong></td>
					<td align="right" style="width: 135px"><strong><i><%=FormatNumber(rsRevenue("AMOUNT2"),2)%></i></strong></td>
					<td align="right" style="width: 135px"><strong><i><%=FormatNumber(cdbl(to_nol(VARIANCE_TODATE)),2)%></i></strong></td>
					<td align="right" style="width: 135px"><strong><i><%=FormatNumber(rsRevenue("AMOUNT3"),2)%></i></strong></td>
					<td align="right" style="width: 135px"><strong><i><%=FormatNumber(rsRevenue("AMOUNT3"),2)%></i></strong></td>
					<td align="right" ><strong><i></i></strong></td>
				<%End if%>
		</tr>
	  <%
	 	v_amountrevenue = v_amountrevenue + cdbl(to_nol(rsRevenue("AMOUNT")))
		v_amountrevenue2 = v_amountrevenue2 + cdbl(to_nol(rsRevenue("AMOUNT2")))	
		v_amountrevenue3 = v_amountrevenue3 + cdbl(to_nol(rsRevenue("AMOUNT3")))	
		v_variance_revenue_thismonth = v_variance_revenue_thismonth + cdbl(to_nol(VARIANCE_THISMONTH))
		v_variance_revenue_todate = v_variance_revenue_todate + cdbl(to_nol(VARIANCE_TODATE))	
	  'response.write cdbl(trim(rsRevenue("MTHOPENBALANCE")))+1
	  rsRevenue.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsRevenue.close()
	rsRevenue=null%>
	</form>
	<tr>
		<td colspan="2" style="border-right: double; width: 600px; padding-left: 30px; border-bottom: double; border-top: double;"><strong><i>TOTAL</i></strong></td>
		<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount,2) %></i></strong></td>
		<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount,2) %></i></strong></td>
		<td align="right" style="border-right: double; width: 250px; border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount_THISMONTH,2) %></i></strong></td>

		<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount2,2) %></i></strong></td>
		<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount2,2) %></i></strong></td>
		<td align="right" style="border-right: double; width: 250px; border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount2_TODATE,2) %></i></strong></td>

		<td align="right" style="border-right: double; width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount3,2) %></i></strong></td>
		<td align="right" style=" width: 250px;border-bottom: double; border-top: double;"><strong><i><%= FormatNumber(tempAmount3,2) %></i></strong></td>
	</tr>

	<tr>
		<th align="Center" colspan="<%=var_colspan+1%>"> <font face="Verdana" size="1">GRAND TOTAL</th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_amountrevenue,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_amountrevenue,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_variance_revenue_thismonth,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_amountrevenue2,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_amountrevenue2,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_variance_revenue_todate,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_amountrevenue3,2)%></th>
		<th align="Right"> <font face="Verdana" size="1"><%=FormatNumber(v_amountrevenue3,2)%></th>
	</tr>
<tbody>
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