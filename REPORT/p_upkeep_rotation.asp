<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay,v_txtDate, v_txtAfdeling, v_txtBlok, v_txtColor
Dim ScriptName
Dim rsPusingan, strPusingan, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName,varCompanyCode, varCompanyName
Dim oRS, oSQL
maxCol = 0

txtYear = Request("txtYear")
txtFcba = Request("txtFcba")
txtAfdeling = Request("txtAfdeling")
txtactivity = Request("txtactivity")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtactivity = "" Then txtactivity = "1"

d_span = 31

strDefaultBa = " SELECT fccode AS fcba" & _
				" FROM businessunit BU" & _
				" INNER JOIN user_fcba_tab uft ON (bu.fccode = uft.fcba) " & _
				" WHERE isdefault = 'TRUE'" & _
				"       AND fctype = 'E'" & _
				"       AND fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If "" & txtFcba = "" Then txtFcba = varDefaultBa

qSQL = " SELECT param_1 FROM parameter WHERE parcode = 'RPT_UR001' AND fccode = '"&txtactivity&"' AND fcba = 'KDP'"
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open qSQL, connstr, adOpenStatic

If Not qORS.EOF Then
	xparam_1 = qORS("param_1")
end if
qORS.Close
set qORS = Nothing

qSQL = " SELECT rpt_upkeeprotation_api.customquery('"&xparam_1&"', ',') resultx FROM DUAL"
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open qSQL, connstr, adOpenStatic

If Not qORS.EOF Then
	wherejob = qORS("resultx")
end if
qORS.Close
set qORS = Nothing

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_data_eof = true
strPusingan	= 	" SELECT tr.fcba" & _
				"        , bf.division" & _
				"        , TO_CHAR(bf.plantingdate, 'yyyy') tt" & _
				"        , tr.fieldcode" & _
				"        , bf.hectarageplanted" & _
				"        , tr.prog_total_01" & _
				"        , tr.prog_total_02" & _
				"        , tr.prog_total_03" & _
				"        , tr.prog_total_04" & _
				"        , tr.prog_total_05" & _
				"        , tr.prog_total_06" & _
				"        , tr.prog_total_07" & _
				"        , tr.prog_total_08" & _
				"        , tr.prog_total_09" & _
				"        , tr.prog_total_10" & _
				"        , tr.prog_total_11" & _
				"        , tr.prog_total_12" & _
				"        , tr.prog_total_bo_01" & _
				"        , tr.prog_total_bo_02" & _
				"        , tr.prog_total_bo_03" & _
				"        , tr.prog_total_bo_04" & _
				"        , tr.prog_total_bo_05" & _
				"        , tr.prog_total_bo_06" & _
				"        , tr.prog_total_bo_07" & _
				"        , tr.prog_total_bo_08" & _
				"        , tr.prog_total_bo_09" & _
				"        , tr.prog_total_bo_10" & _
				"        , tr.prog_total_bo_11" & _
				"        , tr.prog_total_bo_12" & _
				"        , tr.real_total_01" & _
				"        , tr.real_total_02" & _
				"        , tr.real_total_03" & _
				"        , tr.real_total_04" & _
				"        , tr.real_total_05" & _
				"        , tr.real_total_06" & _
				"        , tr.real_total_07" & _
				"        , tr.real_total_08" & _
				"        , tr.real_total_09" & _
				"        , tr.real_total_10" & _
				"        , tr.real_total_11" & _
				"        , tr.real_total_12" & _
				"        , tr.real_total_bo_01" & _
				"        , tr.real_total_bo_02" & _
				"        , tr.real_total_bo_03" & _
				"        , tr.real_total_bo_04" & _
				"        , tr.real_total_bo_05" & _
				"        , tr.real_total_bo_06" & _
				"        , tr.real_total_bo_07" & _
				"        , tr.real_total_bo_08" & _
				"        , tr.real_total_bo_09" & _
				"        , tr.real_total_bo_10" & _
				"        , tr.real_total_bo_11" & _
				"        , tr.real_total_bo_12" & _
				"        , tr.prog_total" & _
				"        , tr.prog_total_bo" & _
				"        , tr.real_total" & _
				"        , tr.real_total_bo" & _
				" FROM (" & _
				" SELECT fcba" & _
				"        , fieldcode" & _
				"        , ROUND(SUM(prog_total_01),2) prog_total_01" & _
				"        , ROUND(SUM(prog_total_02),2) prog_total_02" & _
				"        , ROUND(SUM(prog_total_03),2) prog_total_03" & _
				"        , ROUND(SUM(prog_total_04),2) prog_total_04" & _
				"        , ROUND(SUM(prog_total_05),2) prog_total_05" & _
				"        , ROUND(SUM(prog_total_06),2) prog_total_06" & _
				"        , ROUND(SUM(prog_total_07),2) prog_total_07" & _
				"        , ROUND(SUM(prog_total_08),2) prog_total_08" & _
				"        , ROUND(SUM(prog_total_09),2) prog_total_09" & _
				"        , ROUND(SUM(prog_total_10),2) prog_total_10" & _
				"        , ROUND(SUM(prog_total_11),2) prog_total_11" & _
				"        , ROUND(SUM(prog_total_12),2) prog_total_12" & _
				"        , ROUND(SUM(prog_total_bo_01),2) prog_total_bo_01" & _
				"        , ROUND(SUM(prog_total_bo_02),2) prog_total_bo_02" & _
				"        , ROUND(SUM(prog_total_bo_03),2) prog_total_bo_03" & _
				"        , ROUND(SUM(prog_total_bo_04),2) prog_total_bo_04" & _
				"        , ROUND(SUM(prog_total_bo_05),2) prog_total_bo_05" & _
				"        , ROUND(SUM(prog_total_bo_06),2) prog_total_bo_06" & _
				"        , ROUND(SUM(prog_total_bo_07),2) prog_total_bo_07" & _
				"        , ROUND(SUM(prog_total_bo_08),2) prog_total_bo_08" & _
				"        , ROUND(SUM(prog_total_bo_09),2) prog_total_bo_09" & _
				"        , ROUND(SUM(prog_total_bo_10),2) prog_total_bo_10" & _
				"        , ROUND(SUM(prog_total_bo_11),2) prog_total_bo_11" & _
				"        , ROUND(SUM(prog_total_bo_12),2) prog_total_bo_12" & _
				"        , ROUND(SUM(real_total_01),2) real_total_01" & _
				"        , ROUND(SUM(real_total_02),2) real_total_02" & _
				"        , ROUND(SUM(real_total_03),2) real_total_03" & _
				"        , ROUND(SUM(real_total_04),2) real_total_04" & _
				"        , ROUND(SUM(real_total_05),2) real_total_05" & _
				"        , ROUND(SUM(real_total_06),2) real_total_06" & _
				"        , ROUND(SUM(real_total_07),2) real_total_07" & _
				"        , ROUND(SUM(real_total_08),2) real_total_08" & _
				"        , ROUND(SUM(real_total_09),2) real_total_09" & _
				"        , ROUND(SUM(real_total_10),2) real_total_10" & _
				"        , ROUND(SUM(real_total_11),2) real_total_11" & _
				"        , ROUND(SUM(real_total_12),2) real_total_12" & _
				"        , ROUND(SUM(real_total_bo_01),2) real_total_bo_01" & _
				"        , ROUND(SUM(real_total_bo_02),2) real_total_bo_02" & _
				"        , ROUND(SUM(real_total_bo_03),2) real_total_bo_03" & _
				"        , ROUND(SUM(real_total_bo_04),2) real_total_bo_04" & _
				"        , ROUND(SUM(real_total_bo_05),2) real_total_bo_05" & _
				"        , ROUND(SUM(real_total_bo_06),2) real_total_bo_06" & _
				"        , ROUND(SUM(real_total_bo_07),2) real_total_bo_07" & _
				"        , ROUND(SUM(real_total_bo_08),2) real_total_bo_08" & _
				"        , ROUND(SUM(real_total_bo_09),2) real_total_bo_09" & _
				"        , ROUND(SUM(real_total_bo_10),2) real_total_bo_10" & _
				"        , ROUND(SUM(real_total_bo_11),2) real_total_bo_11" & _
				"        , ROUND(SUM(real_total_bo_12),2) real_total_bo_12" & _
				"        , ROUND(SUM(prog_total_01)+SUM(prog_total_02)+SUM(prog_total_03)+SUM(prog_total_04)+SUM(prog_total_05)+SUM(prog_total_06)+" & _
				"          SUM(prog_total_07)+SUM(prog_total_08)+SUM(prog_total_09)+SUM(prog_total_10)+SUM(prog_total_11)+SUM(prog_total_12),2) prog_total" & _
				"        , ROUND(SUM(prog_total_bo_01)+SUM(prog_total_bo_02)+SUM(prog_total_bo_03)+SUM(prog_total_bo_04)+SUM(prog_total_bo_05)+SUM(prog_total_bo_06)+" & _
				"          SUM(prog_total_bo_07)+SUM(prog_total_bo_08)+SUM(prog_total_bo_09)+SUM(prog_total_bo_10)+SUM(prog_total_bo_11)+SUM(prog_total_bo_12),2) prog_total_bo" & _
				"        , ROUND(SUM(real_total),2) real_total" & _
				"        ,ROUND( SUM(real_total_bo),2) real_total_bo" & _
				" FROM (" & _
				" SELECT ur.fcba" & _
				"        , ur.fieldcode" & _
				"        , CASE WHEN ur.period_month = 1 THEN ur.mandays ELSE 0 END prog_total_01" & _
				"        , CASE WHEN ur.period_month = 2 THEN ur.mandays ELSE 0 END prog_total_02" & _
				"        , CASE WHEN ur.period_month = 3 THEN ur.mandays ELSE 0 END prog_total_03" & _
				"        , CASE WHEN ur.period_month = 4 THEN ur.mandays ELSE 0 END prog_total_04" & _
				"        , CASE WHEN ur.period_month = 5 THEN ur.mandays ELSE 0 END prog_total_05" & _
				"        , CASE WHEN ur.period_month = 6 THEN ur.mandays ELSE 0 END prog_total_06" & _
				"        , CASE WHEN ur.period_month = 7 THEN ur.mandays ELSE 0 END prog_total_07" & _
				"        , CASE WHEN ur.period_month = 8 THEN ur.mandays ELSE 0 END prog_total_08" & _
				"        , CASE WHEN ur.period_month = 9 THEN ur.mandays ELSE 0 END prog_total_09" & _
				"        , CASE WHEN ur.period_month = 10 THEN ur.mandays ELSE 0 END prog_total_10" & _
				"        , CASE WHEN ur.period_month = 11 THEN ur.mandays ELSE 0 END prog_total_11" & _
				"        , CASE WHEN ur.period_month = 12 THEN ur.mandays ELSE 0 END prog_total_12" & _
				"        , 0 prog_total" & _
				"        , CASE WHEN ur.period_month = 1 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_01" & _
				"        , CASE WHEN ur.period_month = 2 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_02" & _
				"        , CASE WHEN ur.period_month = 3 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_03" & _
				"        , CASE WHEN ur.period_month = 4 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_04" & _
				"        , CASE WHEN ur.period_month = 5 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_05" & _
				"        , CASE WHEN ur.period_month = 6 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_06" & _
				"        , CASE WHEN ur.period_month = 7 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_07" & _
				"        , CASE WHEN ur.period_month = 8 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_08" & _
				"        , CASE WHEN ur.period_month = 9 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_09" & _
				"        , CASE WHEN ur.period_month = 10 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_10" & _
				"        , CASE WHEN ur.period_month = 11 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_11" & _
				"        , CASE WHEN ur.period_month = 12 THEN bf.hectarageplanted ELSE 0 END prog_total_bo_12" & _
				"        , 0 prog_total_bo" & _
				"        , 0 real_total_01" & _
				"        , 0 real_total_02" & _
				"        , 0 real_total_03" & _
				"        , 0 real_total_04" & _
				"        , 0 real_total_05" & _
				"        , 0 real_total_06" & _
				"        , 0 real_total_07" & _
				"        , 0 real_total_08" & _
				"        , 0 real_total_09" & _
				"        , 0 real_total_10" & _
				"        , 0 real_total_11" & _
				"        , 0 real_total_12" & _
				"        , 0 real_total" & _
				"        , 0 real_total_bo_01" & _
				"        , 0 real_total_bo_02" & _
				"        , 0 real_total_bo_03" & _
				"        , 0 real_total_bo_04" & _
				"        , 0 real_total_bo_05" & _
				"        , 0 real_total_bo_06" & _
				"        , 0 real_total_bo_07" & _
				"        , 0 real_total_bo_08" & _
				"        , 0 real_total_bo_09" & _
				"        , 0 real_total_bo_10" & _
				"        , 0 real_total_bo_11" & _
				"        , 0 real_total_bo_12" & _
				"        , 0 real_total_bo" & _
				" FROM upkeeprotation ur" & _
				" LEFT JOIN budget_field bf ON (ur.fcba = bf.fcba AND ur.fieldcode = bf.fccode AND bf.budget_year = '"&txtYear&"' AND bf.budget_field_version = '1')" & _
				" WHERE ur.period_year = '"&txtYear&"'" & _
				"       AND ur.fcba = '" & txtFcba & "'" & _
				"       AND ur.jobcode "&wherejob&"" & _
				" UNION ALL" & _
				" SELECT fcba" & _
				"        , fieldcode" & _
				"        , 0 prog_total_01" & _
				"        , 0 prog_total_02" & _
				"        , 0 prog_total_03" & _
				"        , 0 prog_total_04" & _
				"        , 0 prog_total_05" & _
				"        , 0 prog_total_06" & _
				"        , 0 prog_total_07" & _
				"        , 0 prog_total_08" & _
				"        , 0 prog_total_09" & _
				"        , 0 prog_total_10" & _
				"        , 0 prog_total_11" & _
				"        , 0 prog_total_12" & _
				"        , 0 prog_total" & _
				"        , 0 prog_total_bo_01" & _
				"        , 0 prog_total_bo_02" & _
				"        , 0 prog_total_bo_03" & _
				"        , 0 prog_total_bo_04" & _
				"        , 0 prog_total_bo_05" & _
				"        , 0 prog_total_bo_06" & _
				"        , 0 prog_total_bo_07" & _
				"        , 0 prog_total_bo_08" & _
				"        , 0 prog_total_bo_09" & _
				"        , 0 prog_total_bo_10" & _
				"        , 0 prog_total_bo_11" & _
				"        , 0 prog_total_bo_12" & _
				"        , 0 prog_total_bo" & _
				"        , CASE WHEN month = LPAD('1',2,0) THEN total_01 ELSE 0 END real_total_01" & _
				"        , CASE WHEN month = LPAD('2',2,0) THEN total_02 ELSE 0 END real_total_02" & _
				"        , CASE WHEN month = LPAD('3',2,0) THEN total_03 ELSE 0 END real_total_03" & _
				"        , CASE WHEN month = LPAD('4',2,0) THEN total_04 ELSE 0 END real_total_04" & _
				"        , CASE WHEN month = LPAD('5',2,0) THEN total_05 ELSE 0 END real_total_05" & _
				"        , CASE WHEN month = LPAD('6',2,0) THEN total_06 ELSE 0 END real_total_06" & _
				"        , CASE WHEN month = LPAD('7',2,0) THEN total_07 ELSE 0 END real_total_07" & _
				"        , CASE WHEN month = LPAD('8',2,0) THEN total_08 ELSE 0 END real_total_08" & _
				"        , CASE WHEN month = LPAD('9',2,0) THEN total_09 ELSE 0 END real_total_09" & _
				"        , CASE WHEN month = LPAD('10',2,0) THEN total_10 ELSE 0 END real_total_10" & _
				"        , CASE WHEN month = LPAD('11',2,0) THEN total_11 ELSE 0 END real_total_11" & _
				"        , CASE WHEN month = LPAD('12',2,0) THEN total_12 ELSE 0 END real_total_12" & _
				"        , CASE WHEN month = LPAD('1',2,0) THEN total_01" & _
				"               WHEN month = LPAD('2',2,0) THEN total_02" & _
				"               WHEN month = LPAD('3',2,0) THEN total_03" & _
				"               WHEN month = LPAD('4',2,0) THEN total_04" & _
				"               WHEN month = LPAD('5',2,0) THEN total_05" & _
				"               WHEN month = LPAD('6',2,0) THEN total_06" & _
				"               WHEN month = LPAD('7',2,0) THEN total_07" & _
				"               WHEN month = LPAD('8',2,0) THEN total_08" & _
				"               WHEN month = LPAD('9',2,0) THEN total_09" & _
				"               WHEN month = LPAD('10',2,0) THEN total_10" & _
				"               WHEN month = LPAD('11',2,0) THEN total_11" & _
				"               WHEN month = LPAD('12',2,0) THEN total_12 ELSE 0 END real_total" & _
				"        , CASE WHEN month = LPAD('1',2,0) THEN total_bo_01 ELSE 0 END real_total_bo_01" & _
				"        , CASE WHEN month = LPAD('2',2,0) THEN total_bo_02 ELSE 0 END real_total_bo_02" & _
				"        , CASE WHEN month = LPAD('3',2,0) THEN total_bo_03 ELSE 0 END real_total_bo_03" & _
				"        , CASE WHEN month = LPAD('4',2,0) THEN total_bo_04 ELSE 0 END real_total_bo_04" & _
				"        , CASE WHEN month = LPAD('5',2,0) THEN total_bo_05 ELSE 0 END real_total_bo_05" & _
				"        , CASE WHEN month = LPAD('6',2,0) THEN total_bo_06 ELSE 0 END real_total_bo_06" & _
				"        , CASE WHEN month = LPAD('7',2,0) THEN total_bo_07 ELSE 0 END real_total_bo_07" & _
				"        , CASE WHEN month = LPAD('8',2,0) THEN total_bo_08 ELSE 0 END real_total_bo_08" & _
				"        , CASE WHEN month = LPAD('9',2,0) THEN total_bo_09 ELSE 0 END real_total_bo_09" & _
				"        , CASE WHEN month = LPAD('10',2,0) THEN total_bo_10 ELSE 0 END real_total_bo_10" & _
				"        , CASE WHEN month = LPAD('11',2,0) THEN total_bo_11 ELSE 0 END real_total_bo_11" & _
				"        , CASE WHEN month = LPAD('12',2,0) THEN total_bo_12 ELSE 0 END real_total_bo_12" & _
				"        , CASE WHEN month = LPAD('1',2,0) THEN total_bo_01" & _
				"               WHEN month = LPAD('2',2,0) THEN total_bo_02" & _
				"               WHEN month = LPAD('3',2,0) THEN total_bo_03" & _
				"               WHEN month = LPAD('4',2,0) THEN total_bo_04" & _
				"               WHEN month = LPAD('5',2,0) THEN total_bo_05" & _
				"               WHEN month = LPAD('6',2,0) THEN total_bo_06" & _
				"               WHEN month = LPAD('7',2,0) THEN total_bo_07" & _
				"               WHEN month = LPAD('8',2,0) THEN total_bo_08" & _
				"               WHEN month = LPAD('9',2,0) THEN total_bo_09" & _
				"               WHEN month = LPAD('10',2,0) THEN total_bo_10" & _
				"               WHEN month = LPAD('11',2,0) THEN total_bo_11" & _
				"               WHEN month = LPAD('12',2,0) THEN total_bo_12 ELSE 0 END real_total_bo" & _
				" FROM merumput_monthly_report_tab" & _
				" WHERE year = '"&txtYear&"'" & _
				"       AND fcba = '" & txtFcba & "'" & _
				"       AND jobcode "&wherejob&"" & _
				" )" & _
				" GROUP BY fcba" & _
				"          , fieldcode" & _
				" ) tr" & _
				" LEFT JOIN budget_field bf ON (tr.fcba = bf.fcba AND tr.fieldcode = bf.fccode AND bf.budget_year = '"&txtYear&"' AND bf.budget_field_version = '1')" & _
				" LEFT JOIN ( SELECT ur.fieldcode, rpt_upkeeprotation_api.fix_orderdata(ur.orderdata) orderdata" & _
				"             FROM upkeeprotation ur" & _
				"             INNER JOIN ( SELECT fieldcode, MIN(period_month) period_month" & _
				"                          FROM upkeeprotation" & _
				"                          WHERE period_year = '"&txtYear&"' AND fcba = '" & txtFcba & "' AND jobcode "&wherejob&"" & _
				"                          GROUP BY fieldcode) urr" & _
				"             ON (ur.fieldcode = urr.fieldcode AND ur.period_month = urr.period_month)" & _
				"             WHERE ur.period_year = '"&txtYear&"' AND ur.fcba = '" & txtFcba & "' AND jobcode "&wherejob&") orderdata ON (tr.fieldcode = orderdata.fieldcode)" & _
				" WHERE tr.fcba = '" & txtFcba & "'"
				
if "" & txtAfdeling <> "" then strPusingan = strPusingan & " AND bf.division = '" & txtAfdeling & "'"
strPusingan = strPusingan & " ORDER BY tr.fcba, bf.division, orderdata.orderdata"
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_data = qORS.getrows()
	arr_data_eof = false
End If
qORS.Close
set qORS = Nothing

Function f_value(data_array__, fieldcode__, fccode__)
	result_ = 0
	arr_1 = data_array__
	if not arr_data_eof then
	for indx = 0 to UBound(arr_1, 2)
		if "" & fieldcode__ <> "" and "" & fccode__ <> "" then			
			if cstr(arr_1(2, indx)) = cstr(fieldcode__) and cstr(arr_1(3, indx)) = cstr(fccode__) then
				result_ = cdbl(to_nol(arr_1(4, indx)))
			end if
		end if
	next
	end if
	f_value = result_
End Function

fddate = txtMonth&"/1/"&txtYear

if "" & v_txtBlok <> "" then
	wblok = v_txtBlok
end if


strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

Function f_coloringprog(a_)
	result_ = ""
	
	if cdbl(to_nol(a_)) > 0 then
		result_ = "style='background-color:#01FF70;'"
	else
		result_ = "style='background-color:#FFFFFF;'"
	end if
	f_coloringprog = result_
End Function

Function f_coloringreal(a_, b_)
	result_ = ""
	
	if cdbl(to_nol(a_)) = 0 then
		if cdbl(to_nol(b_)) > 0 then
			result_ = "style='background-color:#FF4136;'"
		else
			result_ = "style='background-color:#FFFFFF;'"
		end if
	else
		if cdbl(to_nol(b_)) = 0 then
			result_ = "style='background-color:#FFFFFF;'"
		else
			if cdbl(to_nol(b_)) < cdbl(to_nol(a_)) then
				result_ = "style='background-color:#FF4136;'"
			elseif cdbl(to_nol(b_)) >= cdbl(to_nol(a_)) then
				result_ = "style='background-color:#0074D9;'"
			end if
		end if
	end if
	f_coloringreal = result_
End Function

Function f_maxrow_division(data_array__, code__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if cstr("" & arr_1(1, indx)) = cstr("" & code__) then
			result_ = result_ +1
		end if
	next
	f_maxrow_division = result_
End Function

%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>Upkeep Rotation - iPlaS Reporting</title>
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}

th.rotate-270 {
  height: 200px; /* 1 */
  white-space: nowrap;
  min-width:40px;
  max-width:40px;
  overflow:hidden;
}
th.rotate-270 > div {
	transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-webkit-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-ms-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
    width:15px; /* 4 */
	font-size:10px;
	font-weight:normal;
}

/*
1. Mengatur tinggi table head.
2. mengatur posisi X div
3. mengatur posisi Y div.
4. mengatur lebarnya DIV yang akan mempengaruhi Lebarnya Table Head.
*/
</style>
</head>

<body>
	<div id="wrapper">
	 <!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
				<div class="printhide">
					<table id="form-table" class="centered-margin">	
						<tr>
						<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM

						Yr = txtYear
						MinYr = 2015
						MaxYr = Yr  + 2
					%>
							<td width="4">
								<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
												" AND FCTYPE = 'E'" & _
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									if not rsFcba.eof then
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
									<%
										rsFcba.movenext
									loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>

							<td>
								<font face="verdana" size="1">AFDELING</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="">--All--</option>
									<%
									Dim rsAfdeling, strAfdeling
									strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' AND FCBA = '" & txtFcba & "' ORDER BY FCCODE"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = txtAfdeling then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE")%></option>  
									<%End if
										rsAfdeling.movenext
									loop
									end if
									rsAfdeling.Close
									set rsAfdeling = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">Activity</font>
							</td>
							<td>
								<select name="txtactivity" >
									<%
									strBlok	= " SELECT rate_1 orderdata, fccode, fcname, param_1, param_2, param_3 FROM parameter" & _
									          " WHERE parcode = 'RPT_UR001' AND fcba = 'KDP' ORDER BY rate_1"
									
									Set rsBlok = Server.CreateObject("ADODB.Recordset")
									rsBlok.Open strBlok, connstr, adOpenStatic 
									
									if not rsBlok.eof then
									Do While Not rsBlok.eof
									if rsBlok("FCCODE") = txtactivity then
										activityname = rsBlok("fcname")%>
									<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("orderdata")%> - <%=rsBlok("FCNAME")%></option> 
									<%Else%>
									<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("orderdata")%> - <%=rsBlok("FCNAME")%></option>  
									<%End if
										rsBlok.movenext
									loop
									end if
									rsBlok.Close
									set rsBlok = Nothing
									%>
								</select>
							</td>

							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>PROGRAM & REALISASI TAHUN <%=txtYear%></h2>
				<h3><%=activityname%></h3>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th rowspan="2">Kebun</th>
		<th rowspan="2">Afdeling</th>
		<th rowspan="2">TT</th>
		<th rowspan="2">Blok</th>
		<th rowspan="2">Luas</th>
		
		<th colspan="2">Jan</th>
		<th colspan="2">Feb</th>
		<th colspan="2">Mar</th>
		<th colspan="2">Apr</th>
		<th colspan="2">Mei</th>
		<th colspan="2">Jun</th>
		<th colspan="2">Jul</th>
		<th colspan="2">Agu</th>
		<th colspan="2">Sep</th>
		<th colspan="2">Okt</th>
		<th colspan="2">Nov</th>
		<th colspan="2">Des</th>
		<th colspan="2">Total</th>
	</tr>
	<tr>
		<%for ci=1 to 13%>
		<th>Prog</th>
		<th>Real</th>
		<%next%>
	</tr>
    </thead>
	<tbody>
<%
	if not arr_data_eof then
		defST = 0
		
		stha_prog_01 = 0
		stha_real_01 = 0
		stha_prog_02 = 0
		stha_real_02 = 0
		stha_prog_03 = 0
		stha_real_03 = 0
		stha_prog_04 = 0
		stha_real_04 = 0
		stha_prog_05 = 0
		stha_real_05 = 0
		stha_prog_06 = 0
		stha_real_06 = 0
		stha_prog_07 = 0
		stha_real_07 = 0
		stha_prog_08 = 0
		stha_real_08 = 0
		stha_prog_09 = 0
		stha_real_09 = 0
		stha_prog_10 = 0
		stha_real_10 = 0
		stha_prog_11 = 0
		stha_real_11 = 0
		stha_prog_12 = 0
		stha_real_12 = 0
		stha_prog_total = 0
		stha_real_total = 0
		
		sthk_prog_01 = 0
		sthk_real_01 = 0
		sthk_prog_02 = 0
		sthk_real_02 = 0
		sthk_prog_03 = 0
		sthk_real_03 = 0
		sthk_prog_04 = 0
		sthk_real_04 = 0
		sthk_prog_05 = 0
		sthk_real_05 = 0
		sthk_prog_06 = 0
		sthk_real_06 = 0
		sthk_prog_07 = 0
		sthk_real_07 = 0
		sthk_prog_08 = 0
		sthk_real_08 = 0
		sthk_prog_09 = 0
		sthk_real_09 = 0
		sthk_prog_10 = 0
		sthk_real_10 = 0
		sthk_prog_11 = 0
		sthk_real_11 = 0
		sthk_prog_12 = 0
		sthk_real_12 = 0
		sthk_prog_total = 0
		sthk_real_total = 0

		varha_01 = 0
		varha_02 = 0
		varha_03 = 0
		varha_04 = 0
		varha_05 = 0
		varha_06 = 0
		varha_07 = 0
		varha_08 = 0
		varha_09 = 0
		varha_10 = 0
		varha_11 = 0
		varha_12 = 0
		varha_total = 0


		for y=0 to ubound(arr_data,2)
			defST = defST +1
			curr_division = arr_data(1,y)
			if curr_division <> prev_division then
				maxST = f_maxrow_division(arr_data, curr_division)
			end if
			
			stha_prog_01 = stha_prog_01 + cdbl(to_nol(arr_data(17,y)))
			stha_real_01 = stha_real_01 + cdbl(to_nol(arr_data(41,y)))
			stha_prog_02 = stha_prog_02 + cdbl(to_nol(arr_data(18,y)))
			stha_real_02 = stha_real_02 + cdbl(to_nol(arr_data(42,y)))
			stha_prog_03 = stha_prog_03 + cdbl(to_nol(arr_data(19,y)))
			stha_real_03 = stha_real_03 + cdbl(to_nol(arr_data(43,y)))
			stha_prog_04 = stha_prog_04 + cdbl(to_nol(arr_data(20,y)))
			stha_real_04 = stha_real_04 + cdbl(to_nol(arr_data(44,y)))
			stha_prog_05 = stha_prog_05 + cdbl(to_nol(arr_data(21,y)))
			stha_real_05 = stha_real_05 + cdbl(to_nol(arr_data(45,y)))
			stha_prog_06 = stha_prog_06 + cdbl(to_nol(arr_data(22,y)))
			stha_real_06 = stha_real_06 + cdbl(to_nol(arr_data(46,y)))
			stha_prog_07 = stha_prog_07 + cdbl(to_nol(arr_data(23,y)))
			stha_real_07 = stha_real_07 + cdbl(to_nol(arr_data(47,y)))
			stha_prog_08 = stha_prog_08 + cdbl(to_nol(arr_data(24,y)))
			stha_real_08 = stha_real_08 + cdbl(to_nol(arr_data(48,y)))
			stha_prog_09 = stha_prog_09 + cdbl(to_nol(arr_data(25,y)))
			stha_real_09 = stha_real_09 + cdbl(to_nol(arr_data(49,y)))
			stha_prog_10 = stha_prog_10 + cdbl(to_nol(arr_data(26,y)))
			stha_real_10 = stha_real_10 + cdbl(to_nol(arr_data(50,y)))
			stha_prog_11 = stha_prog_11 + cdbl(to_nol(arr_data(27,y)))
			stha_real_11 = stha_real_11 + cdbl(to_nol(arr_data(51,y)))
			stha_prog_12 = stha_prog_12 + cdbl(to_nol(arr_data(28,y)))
			stha_real_12 = stha_real_12 + cdbl(to_nol(arr_data(52,y)))
			stha_prog_total = stha_prog_total + cdbl(to_nol(arr_data(54,y)))
			stha_real_total = stha_real_total + cdbl(to_nol(arr_data(56,y)))
			
			if stha_prog_01 = 0 then sthkperha_prog_01 = 0 else sthkperha_prog_01 = round(sthk_prog_01/stha_prog_01,2)
			if stha_real_01 = 0 then sthkperha_real_01 = 0 else sthkperha_real_01 = round(sthk_real_01/stha_real_01,2)
			if stha_prog_02 = 0 then sthkperha_prog_02 = 0 else sthkperha_prog_02 = round(sthk_prog_02/stha_prog_02,2)
			if stha_real_02 = 0 then sthkperha_real_02 = 0 else sthkperha_real_02 = round(sthk_real_02/stha_real_02,2)
			if stha_prog_03 = 0 then sthkperha_prog_03 = 0 else sthkperha_prog_03 = round(sthk_prog_03/stha_prog_03,2)
			if stha_real_03 = 0 then sthkperha_real_03 = 0 else sthkperha_real_03 = round(sthk_real_03/stha_real_03,2)
			if stha_prog_04 = 0 then sthkperha_prog_04 = 0 else sthkperha_prog_04 = round(sthk_prog_04/stha_prog_04,2)
			if stha_real_04 = 0 then sthkperha_real_04 = 0 else sthkperha_real_04 = round(sthk_real_04/stha_real_04,2)
			if stha_prog_05 = 0 then sthkperha_prog_05 = 0 else sthkperha_prog_05 = round(sthk_prog_05/stha_prog_05,2)
			if stha_real_05 = 0 then sthkperha_real_05 = 0 else sthkperha_real_05 = round(sthk_real_05/stha_real_05,2)
			if stha_prog_06 = 0 then sthkperha_prog_06 = 0 else sthkperha_prog_06 = round(sthk_prog_06/stha_prog_06,2)
			if stha_real_06 = 0 then sthkperha_real_06 = 0 else sthkperha_real_06 = round(sthk_real_06/stha_real_06,2)
			if stha_prog_07 = 0 then sthkperha_prog_07 = 0 else sthkperha_prog_07 = round(sthk_prog_07/stha_prog_07,2)
			if stha_real_07 = 0 then sthkperha_real_07 = 0 else sthkperha_real_07 = round(sthk_real_07/stha_real_07,2)
			if stha_prog_08 = 0 then sthkperha_prog_08 = 0 else sthkperha_prog_08 = round(sthk_prog_08/stha_prog_08,2)
			if stha_real_08 = 0 then sthkperha_real_08 = 0 else sthkperha_real_08 = round(sthk_real_08/stha_real_08,2)
			if stha_prog_09 = 0 then sthkperha_prog_09 = 0 else sthkperha_prog_09 = round(sthk_prog_09/stha_prog_09,2)
			if stha_real_09 = 0 then sthkperha_real_09 = 0 else sthkperha_real_09 = round(sthk_real_09/stha_real_09,2)
			if stha_prog_10 = 0 then sthkperha_prog_10 = 0 else sthkperha_prog_10 = round(sthk_prog_10/stha_prog_10,2)
			if stha_real_10 = 0 then sthkperha_real_10 = 0 else sthkperha_real_10 = round(sthk_real_10/stha_real_10,2)
			if stha_prog_11 = 0 then sthkperha_prog_11 = 0 else sthkperha_prog_11 = round(sthk_prog_11/stha_prog_11,2)
			if stha_real_11 = 0 then sthkperha_real_11 = 0 else sthkperha_real_11 = round(sthk_real_11/stha_real_11,2)
			if stha_prog_12 = 0 then sthkperha_prog_12 = 0 else sthkperha_prog_12 = round(sthk_prog_12/stha_prog_12,2)
			if stha_real_12 = 0 then sthkperha_real_12 = 0 else sthkperha_real_12 = round(sthk_real_12/stha_real_12,2)
			if stha_prog_total = 0 then sthkperha_prog_total = 0 else sthkperha_prog_total = round(sthk_prog_total/stha_prog_total,2)
			if stha_real_total = 0 then sthkperha_real_total = 0 else sthkperha_real_total = round(sthk_real_total/stha_real_total,2)

			
			varha_01 = stha_prog_01 - stha_real_01
			varha_02 = stha_prog_02 - stha_real_02
			varha_03 = stha_prog_03 - stha_real_03
			varha_04 = stha_prog_04 - stha_real_04
			varha_05 = stha_prog_05 - stha_real_05
			varha_06 = stha_prog_06 - stha_real_06
			varha_07 = stha_prog_07 - stha_real_07
			varha_08 = stha_prog_08 - stha_real_08
			varha_09 = stha_prog_09 - stha_real_09
			varha_10 = stha_prog_10 - stha_real_10
			varha_11 = stha_prog_11 - stha_real_11
			varha_12 = stha_prog_12 - stha_real_12
			varha_total = stha_prog_total - stha_real_total
			
			sthk_prog_01 = sthk_prog_01 + cdbl(to_nol(arr_data(5,y)))
			sthk_real_01 = sthk_real_01 + cdbl(to_nol(arr_data(29,y)))
			sthk_prog_02 = sthk_prog_02 + cdbl(to_nol(arr_data(6,y)))
			sthk_real_02 = sthk_real_02 + cdbl(to_nol(arr_data(30,y)))
			sthk_prog_03 = sthk_prog_03 + cdbl(to_nol(arr_data(7,y)))
			sthk_real_03 = sthk_real_03 + cdbl(to_nol(arr_data(31,y)))
			sthk_prog_04 = sthk_prog_04 + cdbl(to_nol(arr_data(8,y)))
			sthk_real_04 = sthk_real_04 + cdbl(to_nol(arr_data(32,y)))
			sthk_prog_05 = sthk_prog_05 + cdbl(to_nol(arr_data(9,y)))
			sthk_real_05 = sthk_real_05 + cdbl(to_nol(arr_data(33,y)))
			sthk_prog_06 = sthk_prog_06 + cdbl(to_nol(arr_data(10,y)))
			sthk_real_06 = sthk_real_06 + cdbl(to_nol(arr_data(34,y)))
			sthk_prog_07 = sthk_prog_07 + cdbl(to_nol(arr_data(11,y)))
			sthk_real_07 = sthk_real_07 + cdbl(to_nol(arr_data(35,y)))
			sthk_prog_08 = sthk_prog_08 + cdbl(to_nol(arr_data(12,y)))
			sthk_real_08 = sthk_real_08 + cdbl(to_nol(arr_data(36,y)))
			sthk_prog_09 = sthk_prog_09 + cdbl(to_nol(arr_data(13,y)))
			sthk_real_09 = sthk_real_09 + cdbl(to_nol(arr_data(37,y)))
			sthk_prog_10 = sthk_prog_10 + cdbl(to_nol(arr_data(14,y)))
			sthk_real_10 = sthk_real_10 + cdbl(to_nol(arr_data(38,y)))
			sthk_prog_11 = sthk_prog_11 + cdbl(to_nol(arr_data(15,y)))
			sthk_real_11 = sthk_real_11 + cdbl(to_nol(arr_data(39,y)))
			sthk_prog_12 = sthk_prog_12 + cdbl(to_nol(arr_data(16,y)))
			sthk_real_12 = sthk_real_12 + cdbl(to_nol(arr_data(40,y)))
			sthk_prog_total = sthk_prog_total + cdbl(to_nol(arr_data(53,y)))
			sthk_real_total = sthk_real_total + cdbl(to_nol(arr_data(55,y)))


			
			response.write "<tr>"
			response.write "	<td align='center'>"&arr_data(0,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(1,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(2,y)&"</td>"
			response.write "	<td align='center'>"&arr_data(3,y)&"</td>"
			response.write "	<td align='right'>"&arr_data(4,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(17,y))&">"&arr_data(17,y)&"</td>" 'prog_total_bo_01
			response.write "	<td align='right' "&f_coloringreal(arr_data(17,y),arr_data(41,y))&">"&arr_data(41,y)&"</td>" 'real_total_bo_01
			response.write "	<td align='right' "&f_coloringprog(arr_data(18,y))&">"&arr_data(18,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(18,y),arr_data(42,y))&">"&arr_data(42,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(19,y))&">"&arr_data(19,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(19,y),arr_data(43,y))&">"&arr_data(43,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(20,y))&">"&arr_data(20,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(20,y),arr_data(44,y))&">"&arr_data(44,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(21,y))&">"&arr_data(21,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(21,y),arr_data(45,y))&">"&arr_data(45,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(22,y))&">"&arr_data(22,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(22,y),arr_data(46,y))&">"&arr_data(46,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(23,y))&">"&arr_data(23,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(23,y),arr_data(47,y))&">"&arr_data(47,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(24,y))&">"&arr_data(24,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(24,y),arr_data(48,y))&">"&arr_data(48,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(25,y))&">"&arr_data(25,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(25,y),arr_data(49,y))&">"&arr_data(49,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(26,y))&">"&arr_data(26,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(26,y),arr_data(50,y))&">"&arr_data(50,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(27,y))&">"&arr_data(27,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(27,y),arr_data(51,y))&">"&arr_data(51,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(28,y))&">"&arr_data(28,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(28,y),arr_data(52,y))&">"&arr_data(52,y)&"</td>"
			response.write "	<td align='right' "&f_coloringprog(arr_data(54,y))&">"&arr_data(54,y)&"</td>"
			response.write "	<td align='right' "&f_coloringreal(arr_data(54,y),arr_data(56,y))&">"&arr_data(56,y)&"</td>"
			response.write "</tr>"
			
			if defST = maxST then
				response.write "<tr><td colspan='"&d_span&"'></td></tr>"
				response.write "<tr>"
				response.write "	<td align='center' rowspan='3' colspan='4'><b>Rencana</b></td>"
				response.write "	<td align='center'><b>Ha</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_01&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_01&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_02&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_02&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_03&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_03&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_04&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_04&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_05&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_05&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_06&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_06&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_07&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_07&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_08&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_08&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_09&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_09&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_10&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_10&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_11&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_11&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_12&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_12&"</b></td>"
				response.write "	<td align='right'><b>"&stha_prog_total&"</b></td>"
				response.write "	<td align='right'><b>"&stha_real_total&"</b></td>"

				response.write "</tr>"
				response.write "<tr>"
				response.write "	<td align='center'><b>HK</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_01&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_01&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_02&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_02&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_03&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_03&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_04&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_04&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_05&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_05&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_06&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_06&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_07&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_07&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_08&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_08&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_09&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_09&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_10&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_10&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_11&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_11&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_12&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_12&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_prog_total&"</b></td>"
				response.write "	<td align='right'><b>"&sthk_real_total&"</b></td>"
				response.write "</tr>"
				response.write "<tr>"
				response.write "	<td align='center'><b>HK/Ha</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_01&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_01&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_02&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_02&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_03&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_03&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_04&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_04&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_05&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_05&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_06&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_06&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_07&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_07&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_08&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_08&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_09&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_09&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_10&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_10&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_11&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_11&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_12&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_12&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_prog_total&"</b></td>"
				response.write "	<td align='right'><b>"&sthkperha_real_total&"</b></td>"
				response.write "</tr>"
				response.write "<tr>"
				response.write "	<td align='center' colspan='4'><b>Varian</b></td>"
				response.write "	<td align='center'><b>+/-</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_01&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_02&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_03&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_04&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_05&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_06&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_07&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_08&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_09&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_10&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_11&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_12&"</b></td>"
				response.write "	<td align='center' colspan='2'><b>"&varha_total&"</b></td>"
				response.write "</tr>"
				
				if y <> cdbl(to_nol(ubound(arr_data,2))) then
					response.write "<tr><td colspan='"&d_span&"'>&nbsp;</td></tr>"
				end if
				
				
				defST = 0
				st_luas = 0
				stha_prog_01 = 0
				stha_real_01 = 0
				stha_prog_02 = 0
				stha_real_02 = 0
				stha_prog_03 = 0
				stha_real_03 = 0
				stha_prog_04 = 0
				stha_real_04 = 0
				stha_prog_05 = 0
				stha_real_05 = 0
				stha_prog_06 = 0
				stha_real_06 = 0
				stha_prog_07 = 0
				stha_real_07 = 0
				stha_prog_08 = 0
				stha_real_08 = 0
				stha_prog_09 = 0
				stha_real_09 = 0
				stha_prog_10 = 0
				stha_real_10 = 0
				stha_prog_11 = 0
				stha_real_11 = 0
				stha_prog_12 = 0
				stha_real_12 = 0
				stha_prog_total = 0
				stha_real_total = 0
				
				sthk_prog_01 = 0
				sthk_real_01 = 0
				sthk_prog_02 = 0
				sthk_real_02 = 0
				sthk_prog_03 = 0
				sthk_real_03 = 0
				sthk_prog_04 = 0
				sthk_real_04 = 0
				sthk_prog_05 = 0
				sthk_real_05 = 0
				sthk_prog_06 = 0
				sthk_real_06 = 0
				sthk_prog_07 = 0
				sthk_real_07 = 0
				sthk_prog_08 = 0
				sthk_real_08 = 0
				sthk_prog_09 = 0
				sthk_real_09 = 0
				sthk_prog_10 = 0
				sthk_real_10 = 0
				sthk_prog_11 = 0
				sthk_real_11 = 0
				sthk_prog_12 = 0
				sthk_real_12 = 0
				sthk_prog_total = 0
				sthk_real_total = 0
				
				varha_01 = 0
				varha_02 = 0
				varha_03 = 0
				varha_04 = 0
				varha_05 = 0
				varha_06 = 0
				varha_07 = 0
				varha_08 = 0
				varha_09 = 0
				varha_10 = 0
				varha_11 = 0
				varha_12 = 0
				varha_total = 0



			end if
			
			prev_division = curr_division
		next
else
	response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
end if
%>
</tbody>
</table>

<style>
.stick { display:none;}
.xfTBarF { height:15px; line-height: 15px; font-size:10px; padding-left:20px; background-color:#008001;}
.xfContainer { border:1px #008001 solid; box-shadow:1px 1px 1px rgba(0,0,0,0.5); background:#FFFFFF;}
.fenster-content { padding:5px; padding-top:10px;}
#indicator_color { border:1px #999999 solid ; width:650px; padding:2px; font-family:Arial, Helvetica, sans-serif; }
#indicator_color table { border-collapse:collapse; border-spacing:0px;}
#indicator_color table, #indicator_color tr, #indicator_color td { border: none; line-height:30px; }
#indicator_color th { font-size:12px; background-color:#222222; border-left:1px #666666 dashed; color:#FFFFFF;}
#indicator_color .2nd { background-color:#f1f1f1;}
#indicator_color td { font-size:12px; }
.merah { width:15px; height:15px; background-color:#FF0000; float:left; margin-right:5px; margin-top:7px;
   	      border-radius:3px;
		  -moz-border-radius:3px;
		  -webkit-border-radius:3px;
		  -o-border-radius:3px;
		  box-shadow:1px 1px 0px #ccc;
		  -moz-box-shadow:1px 1px 0px #ccc;
		  -webkit-box-shadow:1px 1px 0px #ccc;
		  -o-box-shadow:1px 1px 0px #ccc;}
.kuning { width:15px; height:15px; background-color:#FFFF00; float:left; margin-right:5px; margin-top:7px;
		   border-radius:3px;
		  -moz-border-radius:3px;
		  -webkit-border-radius:3px;
		  -o-border-radius:3px;
		  box-shadow:1px 1px 0px #ccc;
		  -moz-box-shadow:1px 1px 0px #ccc;
		  -webkit-box-shadow:1px 1px 0px #ccc;
		  -o-box-shadow:1px 1px 0px #ccc;}
.hijau { width:15px; height:15px; background-color:#008001; float:left; margin-right:5px; margin-top:7px;
          border-radius:3px;
		  -moz-border-radius:3px;
		  -webkit-border-radius:3px;
		  -o-border-radius:3px;
		  box-shadow:1px 1px 0px #ccc;
		  -moz-box-shadow:1px 1px 0px #ccc;
		  -webkit-box-shadow:1px 1px 0px #ccc;
		  -o-box-shadow:1px 1px 0px #ccc;}
#indicator_color .span_color { float:left;}
#indicator_color .point_table {}
#indicator_color .MsoNormal { margin-bottom:3px; border-bottom:1px #ccc dashed; padding:0px 5px;}
.symbol_1 { font-size:18px; color:#008001; margin-right:5px; }
</style>

<div id="indicator_color">
	<div>
	<table cellpadding="0" border="1" style="border:none">
		<thead>
			<tr>
				<th valign="center" style="min-width:100px;max-width:100px;"><p>Warna</p></th>
				<th valign="center" style="min-width:400px;max-width:400px;"><p>Deskripsi</p></th>
			</tr>
		</thead>
		<tbody>
			<tr bgcolor="#f1f1f1">
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#01FF70;padding-left:10px;">Hectare (Ha)</p>
				</td>
				
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;width:35%;">
					<p class="MsoNormal">Program kerja</p>
				</td>
			</tr>
			
			<tr>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#FF4136;padding-left:10px;">Hectare (Ha)</p>
				</td>
				
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Tidak selesai dikerjakan atau dikerjakan tanpa program</p>
				</td>
			</tr>
			
			<tr bgcolor="#f1f1f1">
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#0074D9;padding-left:10px;">Hectare (Ha)</p>
				</td>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Sesuai program kerja</p>
				</td>
			</tr>
			
			<tr>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#FFFFFF;padding-left:10px;">Hectare (Ha)</p>
				</td>
				
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Tidak dikerjakan atau tidak ada program kerja</p>
				</td>
			</tr>
		</tbody>
	</table>
	</div>
</div>
	</body>
</html>