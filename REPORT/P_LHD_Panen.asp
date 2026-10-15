<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtdatefrom, v_txtMonth, v_txtYear, v_txtgrade, v_txtDivisi, v_txtemployeetype
	Dim no
	Dim rsDefaultBa, strDefaultBa, varDefaultBa
	Dim ScriptName
	Dim strMain, sqlLHPP, sqlLHPP2, rsLHPP
	Dim q, CompanyName, fcbaname, MgrName, KTUName, CropName, FcbaType
	Dim OrderSection : OrderSection = 0
	Dim gt_hk_bi : gt_hk_bi = 0
	Dim gt_gaji_bi : gt_gaji_bi = 0
	Dim gt_premilembur_bi : gt_premilembur_bi = 0
	Dim gt_beras_bi : gt_beras_bi = 0
	Dim gt_thr_bi : gt_thr_bi = 0
	Dim gt_bonus_bi : gt_bonus_bi = 0
	Dim gt_bpjs_bi : gt_bpjs_bi = 0
	Dim gt_pengobatan_bi : gt_pengobatan_bi = 0
	Dim gt_perumahan_bi : gt_perumahan_bi = 0
	Dim gt_seragam_bi : gt_seragam_bi = 0
	Dim gt_extrafooding_bi : gt_extrafooding_bi = 0
	Dim gt_hk_sbi : gt_hk_sbi = 0
	Dim gt_gaji_sbi : gt_gaji_sbi = 0
	Dim gt_premilembur_sbi : gt_premilembur_sbi = 0
	Dim gt_beras_sbi : gt_beras_sbi = 0
	Dim gt_thr_sbi : gt_thr_sbi = 0
	Dim gt_bonus_sbi : gt_bonus_sbi = 0
	Dim gt_bpjs_sbi : gt_bpjs_sbi = 0
	Dim gt_pengobatan_sbi : gt_pengobatan_sbi = 0
	Dim gt_perumahan_sbi : gt_perumahan_sbi = 0
	Dim gt_seragam_sbi : gt_seragam_sbi = 0
	Dim gt_extrafooding_sbi : gt_extrafooding_sbi = 0
	
	Dim row_total_bi : row_total_bi = 0
	Dim row_total_sbi : row_total_sbi = 0
	
	Dim gt_total_bi : gt_total_bi = 0
	Dim gt_total_sbi : gt_total_sbi = 0
	
	' **** Sub Total (Kebun)
	Dim rsST, strST, arrayST, arrayST_eof, iii, st_indx
	
	Dim row_st_total_bi : row_st_total_bi = 0
	Dim row_st_total_sbi : row_st_total_sbi = 0
	
	arrayST_eof = true
	' ****End Sub Total (Kebun)
	
	Dim rowspan_division : rowspan_division = 0
	Dim rsRowspan, strRowspan
	Dim posting_rowspan : posting_rowspan = 0


    v_txtGroupPlantation = Request("txtGroupPlantation")
	v_txtPlantation = Request("txtPlantation")
	v_txtGroupRegion = Request("txtGroupRegion")
	v_txtRegion = Request("txtRegion")
	v_txtcompany = Request("txtCompany")
	v_txtfcba = Request("txtfcba")
	
	v_txtDivisi = Request("txtDivisi")
	v_txtemployeetype = Request("txtemployeetype")
	
	v_txtBudgetYear = Request("txtBudgetYear")
	v_txtBudgetVersion = Request("txtBudgetVersion")
	v_txtModuleVersion = Request("txtModuleVersion")
	
	v_txtdatefrom		= Request("txtdatefrom")
	
	v_txtYear = Request("txtYear")
	v_txtMonth = Request("txtMonth")
	v_txtDate = Request("txtDate")
	v_txtJobCode = Request("txtJobCode")
	v_txtjobgroup = Request("txtjobgroup")
	
	
	If Len(Trim(v_txtBudgetYear)) = 0 Then
		v_txtBudgetYear = Year(textDate(Date))
	Else
		v_txtBudgetYear = v_txtYear
	End If
	If Len(Trim(v_txtBudgetVersion)) = 0 Then v_txtBudgetVersion = 1
	If Len(Trim(v_txtModuleVersion)) = 0 Then v_txtModuleVersion = 1
	
	If Len(Trim(v_txtdatefrom)) = 0 Then v_txtdatefrom = textDate(Date)
	
	if "" & v_txtDate = "" then v_txtDate = textDate(Date)
	v_txtYear = Year(v_txtDate)
	v_txtMonth = Month(v_txtDate)
	
	
	strDefaultBa = " select" & _
				   "     fccode as fcba" & _
				   " from businessunit bu" & _
				   " inner join user_fcba_tab uft on (bu.fccode = uft.fcba) " & _
				   " where isdefault = 'TRUE'" & _
				   "       and fcusercode = '" & Session("CodeName") & "'"

	Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
	rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

	If Not rsDefaultBa.EOF Then
		varDefaultBa = rsDefaultBa("fcba")
	Else
		varDefaultBa = Session("fcba")
	End If
	rsDefaultBa.Close
	set rsDefaultBa = Nothing

	if not isnull(v_txtFcba) then
		If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
			v_txtFcba = varDefaultBa
		End If
	End If
	
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

Dim array_data, array_eof, xx

array_eof = true

strMain = " SELECT fcba" & _
          "        , jobcode" & _
          "        , jobname" & _
          "        , division" & _
          "        , SUM(sku) sku" & _
          "        , SUM(khl) khl" & _
          "        , SUM(unit_daily) unit_daily" & _
          "        , SUM(unit_contract) unit_contract" & _
          "        , SUM(sku_1) sku_1" & _
          "        , SUM(khl_1) khl_1" & _
          "        , SUM(unit_daily_1) unit_daily_1" & _
          "        , SUM(unit_contract_1) unit_contract_1" & _
          "        , SUM(sku_2) sku_2" & _
          "        , SUM(khl_2) khl_2" & _
          "        , SUM(unit_daily_2) unit_daily_2" & _
          "        , SUM(unit_contract_2) unit_contract_2" & _
          "        , SUM(bud_hk) bud_hk" & _
          "        , SUM(bud_ha) bud_ha" & _
          " FROM (" & _
          "     SELECT agem.fcba" & _
          "            , agem.jobcode" & _
          "            , jj.fcname jobname" & _
		  "            , jj.jobgroupcode" & _
          "            , agem.division" & _
          "            , SUM(DECODE(agem.presentgrade, 'KBT', agem.mandays, 'KHT', agem.mandays, 0)) sku" & _
          "            , SUM(DECODE(agem.presentgrade, 'KHL', agem.mandays, 0)) khl" & _
          "            , 0 unit_daily" & _
          "            , 0 unit_contract" & _
          "            , 0 sku_1" & _
          "            , 0 khl_1" & _
          "            , 0 unit_daily_1" & _
          "            , 0 unit_contract_1" & _
          "            , 0 sku_2" & _
          "            , 0 khl_2" & _
          "            , 0 unit_daily_2" & _
          "            , 0 unit_contract_2" & _
          "            , 0 bud_hk" & _
          "            , 0 bud_ha" & _
          "     FROM (" & _
          "         SELECT ag.fcba, ag.jobcode, em.presentgrade, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL) division, SUM(ag.mandays) mandays" & _
          "         FROM (" & _
          "             SELECT fcba, fddate, jobcode, locationtype, locationcode, employeecode, SUM(mandays) mandays" & _
          "             FROM attendance_gad" & _
          "             WHERE fcba = '" & v_txtFcba & "'" & _
          "                   AND fddate = TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
          "                   AND attendance IN ('KJ', 'DL')" & _
          "                   AND locationtype IN ('FF', 'MN')" & _
          "             GROUP BY fcba, fddate, jobcode, locationtype, locationcode, employeecode" & _
          "         ) ag" & _
          "         LEFT JOIN employee_hist_tab em ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode AND em.year = TO_NUMBER(TO_CHAR(fddate, 'yyyy')) AND month = TO_NUMBER(TO_CHAR(fddate, 'mm')))" & _
          "         GROUP BY ag.fcba, ag.jobcode, em.presentgrade, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL)" & _
          "     ) agem" & _
          "     LEFT JOIN job jj ON (agem.fcba = jj.fcba AND agem.jobcode = jj.fccode)" & _
          "     GROUP BY agem.fcba" & _
          "              , agem.jobcode" & _
          "              , jj.fcname" & _
          "              , jj.jobgroupcode" & _
          "              , agem.division" & _
          "     UNION ALL" & _
          "     SELECT agem.fcba" & _
          "            , agem.jobcode" & _
          "            , jj.fcname jobname" & _
		  "            , jj.jobgroupcode" & _
          "            , agem.division" & _
          "            , 0 sku" & _
          "            , 0 khl" & _
          "            , 0 unit_daily" & _
          "            , 0 unit_contract" & _
          "            , SUM(DECODE(agem.presentgrade, 'KBT', agem.mandays, 'KHT', agem.mandays, 0)) sku_1" & _
          "            , SUM(DECODE(agem.presentgrade, 'KHL', agem.mandays, 0)) khl_1" & _
          "            , 0 unit_daily_1" & _
          "            , 0 unit_contract_1" & _
          "            , 0 sku_2" & _
          "            , 0 khl_2" & _
          "            , 0 unit_daily_2" & _
          "            , 0 unit_contract_2" & _
          "            , 0 bud_hk" & _
          "            , 0 bud_ha" & _
          "     FROM (" & _
          "         SELECT ag.fcba, ag.jobcode, em.presentgrade, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL) division, SUM(ag.mandays) mandays" & _
          "         FROM (" & _
          "             SELECT fcba, fddate, jobcode, locationtype, locationcode, employeecode, SUM(mandays) mandays" & _
          "             FROM attendance_gad" & _
          "             WHERE fcba = '" & v_txtFcba & "'" & _
          "                   AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "'||LPAD('" & v_txtMonth & "',2,0)" & _
          "                   AND fddate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
          "                   AND attendance IN ('KJ', 'DL')" & _
          "                   AND locationtype IN ('FF', 'MN')" & _
          "             GROUP BY fcba, fddate, jobcode, locationtype, locationcode, employeecode" & _
          "         ) ag" & _
          "         LEFT JOIN employee_hist_tab em ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode AND em.year = TO_NUMBER(TO_CHAR(fddate, 'yyyy')) AND month = TO_NUMBER(TO_CHAR(fddate, 'mm')))" & _
          "         GROUP BY ag.fcba, ag.jobcode, em.presentgrade, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL)" & _
          "     ) agem" & _
          "     LEFT JOIN job jj ON (agem.fcba = jj.fcba AND agem.jobcode = jj.fccode)" & _
          "     GROUP BY agem.fcba" & _
          "              , agem.jobcode" & _
          "              , jj.fcname" & _
		  "              , jj.jobgroupcode" & _
          "              , agem.division" & _
          "     UNION ALL" & _
          "     SELECT agem.fcba" & _
          "            , agem.jobcode" & _
          "            , jj.fcname jobname" & _
		  "            , jj.jobgroupcode" & _
          "            , agem.division" & _
          "            , 0 sku" & _
          "            , 0 khl" & _
          "            , 0 unit_daily" & _
          "            , 0 unit_contract" & _
          "            , 0 sku_1" & _
          "            , 0 khl_1" & _
          "            , 0 unit_daily_1" & _
          "            , 0 unit_contract_1" & _
          "            , SUM(DECODE(agem.presentgrade, 'KBT', agem.mandays, 'KHT', agem.mandays, 0)) sku_2" & _
          "            , SUM(DECODE(agem.presentgrade, 'KHL', agem.mandays, 0)) khl_2" & _
          "            , 0 unit_daily_2" & _
          "            , 0 unit_contract_2" & _
          "            , 0 bud_hk" & _
          "            , 0 bud_ha" & _
          "     FROM (" & _
          "         SELECT ag.fcba, ag.jobcode, em.presentgrade, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL) division, SUM(ag.mandays) mandays" & _
          "         FROM (" & _
          "             SELECT fcba, fddate, jobcode, locationtype, locationcode, employeecode, SUM(mandays) mandays" & _
          "             FROM attendance_gad" & _
          "             WHERE fcba = '" & v_txtFcba & "'" & _
          "                   AND TO_NUMBER(TO_CHAR(fddate, 'yyyy')) = '" & v_txtYear & "'" & _
          "                   AND TO_NUMBER(TO_CHAR(fddate, 'mm')) <= '" & v_txtMonth & "'" & _
          "                   AND fddate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
          "                   AND attendance IN ('KJ', 'DL')" & _
          "                   AND locationtype IN ('FF', 'MN')" & _
          "             GROUP BY fcba, fddate, jobcode, locationtype, locationcode, employeecode" & _
          "         ) ag" & _
          "         LEFT JOIN employee_hist_tab em ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode AND em.year = TO_NUMBER(TO_CHAR(fddate, 'yyyy')) AND month = TO_NUMBER(TO_CHAR(fddate, 'mm')))" & _
          "         GROUP BY ag.fcba, ag.jobcode, em.presentgrade, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL)" & _
          "     ) agem" & _
          "     LEFT JOIN job jj ON (agem.fcba = jj.fcba AND agem.jobcode = jj.fccode)" & _
          "     GROUP BY agem.fcba" & _
          "              , agem.jobcode" & _
          "              , jj.fcname" & _
		  "              , jj.jobgroupcode" & _
          "              , agem.division" & _
          "     UNION ALL" & _
          "     SELECT agem.fcba" & _
          "            , agem.jobcode" & _
          "            , jj.fcname jobname" & _
		  "            , jj.jobgroupcode" & _
          "            , agem.division" & _
          "            , 0 sku" & _
          "            , 0 khl" & _
          "            , SUM(DECODE(agem.attendance, 'KJ', agem.unit, 'DL', agem.unit, 0)) unit_daily" & _
          "            , SUM(DECODE(agem.attendance, 'KB', agem.unit, 0)) unit_contract" & _
          "            , 0 sku_1" & _
          "            , 0 khl_1" & _
          "            , 0 unit_daily_1" & _
          "            , 0 unit_contract_1" & _
          "            , 0 sku_2" & _
          "            , 0 khl_2" & _
          "            , 0 unit_daily_2" & _
          "            , 0 unit_contract_2" & _
          "            , 0 bud_hk" & _
          "            , 0 bud_ha" & _
          "     FROM (" & _
          "         SELECT ag.fcba, ag.jobcode, ag.attendance, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL) division, SUM(ag.unit) unit" & _
          "         FROM (" & _
          "             SELECT fcba, fddate, jobcode, locationtype, locationcode, attendance, SUM(unit) unit" & _
          "             FROM attendance_gad" & _
          "             WHERE fcba = '" & v_txtFcba & "'" & _
          "                   AND fddate = TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
          "                   AND attendance IN ('KJ', 'DL', 'KB')" & _
          "                   AND locationtype IN ('FF', 'MN')" & _
          "             GROUP BY fcba, fddate, jobcode, locationtype, locationcode, attendance" & _
          "         ) ag" & _
          "         GROUP BY ag.fcba, ag.jobcode, ag.attendance, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL)" & _
          "     ) agem" & _
          "     LEFT JOIN job jj ON (agem.fcba = jj.fcba AND agem.jobcode = jj.fccode)" & _
          "     GROUP BY agem.fcba" & _
          "              , agem.jobcode" & _
          "              , jj.fcname" & _
		  "              , jj.jobgroupcode" & _
          "              , agem.division" & _
          "     UNION ALL" & _
          "     SELECT agem.fcba" & _
          "            , agem.jobcode" & _
          "            , jj.fcname jobname" & _
		  "            , jj.jobgroupcode" & _
          "            , agem.division" & _
          "            , 0 sku" & _
          "            , 0 khl" & _
          "            , 0 unit_daily" & _
          "            , 0 unit_contract" & _
          "            , 0 sku_1" & _
          "            , 0 khl_1" & _
          "            , SUM(DECODE(agem.attendance, 'KJ', agem.unit, 'DL', agem.unit, 0)) unit_daily_1" & _
          "            , SUM(DECODE(agem.attendance, 'KB', agem.unit, 0)) unit_contract_1" & _
          "            , 0 sku_2" & _
          "            , 0 khl_2" & _
          "            , 0 unit_daily_2" & _
          "            , 0 unit_contract_2" & _
          "            , 0 bud_hk" & _
          "            , 0 bud_ha" & _
          "     FROM (" & _
          "         SELECT ag.fcba, ag.jobcode, ag.attendance, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL) division, SUM(ag.unit) unit" & _
          "         FROM (" & _
          "             SELECT fcba, fddate, jobcode, locationtype, locationcode, attendance, SUM(unit) unit" & _
          "             FROM attendance_gad" & _
          "             WHERE fcba = '" & v_txtFcba & "'" & _
          "                   AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "'||LPAD('" & v_txtMonth & "',2,0)" & _
          "                   AND fddate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
          "                   AND attendance IN ('KJ', 'DL', 'KB')" & _
          "                   AND locationtype IN ('FF', 'MN')" & _
          "             GROUP BY fcba, fddate, jobcode, locationtype, locationcode, attendance" & _
          "         ) ag" & _
          "         GROUP BY ag.fcba, ag.jobcode, ag.attendance, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL)" & _
          "     ) agem" & _
          "     LEFT JOIN job jj ON (agem.fcba = jj.fcba AND agem.jobcode = jj.fccode)" & _
          "     GROUP BY agem.fcba" & _
          "              , agem.jobcode" & _
          "              , jj.fcname" & _
		  "              , jj.jobgroupcode" & _
          "              , agem.division" & _
          "     UNION ALL" & _
          "     SELECT agem.fcba" & _
          "            , agem.jobcode" & _
          "            , jj.fcname jobname" & _
		  "            , jj.jobgroupcode" & _
          "            , agem.division" & _
          "            , 0 sku" & _
          "            , 0 khl" & _
          "            , 0 unit_daily" & _
          "            , 0 unit_contract" & _
          "            , 0 sku_1" & _
          "            , 0 khl_1" & _
          "            , 0 unit_daily_1" & _
          "            , 0 unit_contract_1" & _
          "            , 0 sku_2" & _
          "            , 0 khl_2" & _
          "            , SUM(DECODE(agem.attendance, 'KJ', agem.unit, 'DL', agem.unit, 0)) unit_daily_2" & _
          "            , SUM(DECODE(agem.attendance, 'KB', agem.unit, 0)) unit_contract_2" & _
          "            , 0 bud_hk" & _
          "            , 0 bud_ha" & _
          "     FROM (" & _
          "         SELECT ag.fcba, ag.jobcode, ag.attendance, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL) division, SUM(ag.unit) unit" & _
          "         FROM (" & _
          "             SELECT fcba, fddate, jobcode, locationtype, locationcode, attendance, SUM(unit) unit" & _
          "             FROM attendance_gad" & _
          "             WHERE fcba = '" & v_txtFcba & "'" & _
          "                   AND TO_NUMBER(TO_CHAR(fddate, 'yyyy')) = '" & v_txtYear & "'" & _
          "                   AND TO_NUMBER(TO_CHAR(fddate, 'mm')) <= '" & v_txtMonth & "'" & _
          "                   AND fddate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
          "                   AND attendance IN ('KJ', 'DL', 'KB')" & _
          "                   AND locationtype IN ('FF', 'MN')" & _
          "             GROUP BY fcba, fddate, jobcode, locationtype, locationcode, attendance" & _
          "         ) ag" & _
          "         GROUP BY ag.fcba, ag.jobcode, ag.attendance, DECODE(ag.locationtype, 'MN', 'BBT', 'FF', field_api.getvalid_afdeling(ag.locationcode, ag.fcba, ag.fddate), NULL)" & _
          "     ) agem" & _
          "     LEFT JOIN job jj ON (agem.fcba = jj.fcba AND agem.jobcode = jj.fccode)" & _
          "     GROUP BY agem.fcba" & _
          "              , agem.jobcode" & _
          "              , jj.fcname" & _
		  "              , jj.jobgroupcode" & _
          "              , agem.division" & _
          "     UNION ALL" & _
		  "     SELECT agem.fcba" & _
		  "            , agem.jobcode" & _
		  "            , jj.fcname jobname" & _
		  "            , jj.jobgroupcode" & _
		  "            , agem.division" & _
		  "            , 0 sku" & _
		  "            , 0 khl" & _
		  "            , 0 unit_daily" & _
		  "            , 0 unit_contract" & _
		  "            , 0 sku_1" & _
		  "            , 0 khl_1" & _
		  "            , 0 unit_daily_1" & _
		  "            , 0 unit_contract_1" & _
		  "            , 0 sku_2" & _
		  "            , 0 khl_2" & _
		  "            , 0 unit_daily_2" & _
		  "            , 0 unit_contract_2" & _
		  "            , SUM(agem.bud_hk) bud_hk" & _
		  "            , SUM(agem.bud_ha) bud_ha" & _
		  "     FROM (" & _
		  "         SELECT fcba, jobcode, division, SUM( CASE" & _
		  "                                                  WHEN " & v_txtMonth & " = 1 THEN qty_01" & _
		  "                                                  WHEN " & v_txtMonth & " = 2 THEN qty_02" & _
		  "                                                  WHEN " & v_txtMonth & " = 3 THEN qty_03" & _
		  "                                                  WHEN " & v_txtMonth & " = 4 THEN qty_04" & _
		  "                                                  WHEN " & v_txtMonth & " = 5 THEN qty_05" & _
		  "                                                  WHEN " & v_txtMonth & " = 6 THEN qty_06" & _
		  "                                                  WHEN " & v_txtMonth & " = 7 THEN qty_07" & _
		  "                                                  WHEN " & v_txtMonth & " = 8 THEN qty_08" & _
		  "                                                  WHEN " & v_txtMonth & " = 9 THEN qty_09" & _
		  "                                                  WHEN " & v_txtMonth & " = 10 THEN qty_10" & _
		  "                                                  WHEN " & v_txtMonth & " = 11 THEN qty_11" & _
		  "                                                  WHEN " & v_txtMonth & " = 12 THEN qty_12" & _
		  "                                              ELSE 0 END ) bud_hk, 0 bud_ha" & _
		  "         FROM budget_transaction_estate" & _
		  "         WHERE fcba = '" & v_txtFcba & "'" & _
		  "               AND budget_year = '" & v_txtYear & "'" & _
		  "               AND sourcetype = 'LOC'" & _
		  "               AND sourcecode <> 'FISIK'" & _
		  "         GROUP BY fcba, jobcode, division" & _
		  "         UNION ALL" & _
		  "         SELECT fcba, jobcode, division, 0 bud_hk, SUM( CASE" & _
		  "                                                            WHEN " & v_txtMonth & " = 1 THEN qty_01" & _
		  "                                                            WHEN " & v_txtMonth & " = 2 THEN qty_02" & _
		  "                                                            WHEN " & v_txtMonth & " = 3 THEN qty_03" & _
		  "                                                            WHEN " & v_txtMonth & " = 4 THEN qty_04" & _
		  "                                                            WHEN " & v_txtMonth & " = 5 THEN qty_05" & _
		  "                                                            WHEN " & v_txtMonth & " = 6 THEN qty_06" & _
		  "                                                            WHEN " & v_txtMonth & " = 7 THEN qty_07" & _
		  "                                                            WHEN " & v_txtMonth & " = 8 THEN qty_08" & _
		  "                                                            WHEN " & v_txtMonth & " = 9 THEN qty_09" & _
		  "                                                            WHEN " & v_txtMonth & " = 10 THEN qty_10" & _
		  "                                                            WHEN " & v_txtMonth & " = 11 THEN qty_11" & _
		  "                                                            WHEN " & v_txtMonth & " = 12 THEN qty_12" & _
		  "                                                        ELSE 0 END ) bud_ha" & _
		  "         FROM budget_transaction_estate" & _
		  "         WHERE fcba = '" & v_txtFcba & "'" & _
		  "               AND budget_year = '" & v_txtYear & "'" & _
		  "               AND sourcetype = 'LOC'" & _
		  "               AND sourcecode = 'FISIK'" & _
		  "         GROUP BY fcba, jobcode, division" & _
		  "     ) agem" & _
		  "     LEFT JOIN job jj ON (agem.fcba = jj.fcba AND agem.jobcode = jj.fccode)" & _
		  "     GROUP BY agem.fcba" & _
		  "              , agem.jobcode" & _
		  "              , jj.fcname" & _
		  "              , jj.jobgroupcode" & _
		  "              , agem.division" & _
          " ) WHERE fcba = '" & v_txtFcba & "'"
		  
if "" & v_txtDivisi <> "" then
		strMain = strMain & " AND division = '" & v_txtDivisi & "'"
end if

if "" & v_txtjobgroup <> "" then
		strMain = strMain & " AND jobgroupcode = '" & v_txtjobgroup & "'"
end if

if "" & v_txtJobCode <> "" then
		strMain = strMain & " AND jobcode = '" & v_txtJobCode & "'"
end if

strMain = strMain & " GROUP BY fcba" & _
                    "          , jobcode" & _
                    "          , jobname" & _
                    "          , division" & _
                    " ORDER BY jobcode, division"
		  
sqlLHPP2 = " "

sqlLHPP = " SELECT * FROM ("&strMain&")"
		  
' response.write sqlLHPP
%>


<html>

	<head>
		<title>ePlaS Reporting - Laporan Produktivitas Tenaga Kerja</title>
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
				<!--#include virtual="/REPORT/FormSelectFcbaOnly.html"-->
				
					<table id="form-table" class="centered-margin">
						<tr>
							<td><font face="verdana" size="1">TAHUN</font></td>
							<td><font face="verdana" size="1">BULAN</font></td>
							<td><font face="verdana" size="1">TANGGAL</font></td>
							<td><font face="verdana" size="1">DIVISI</font></td>
							<td><font face="verdana" size="1">JOBGROUP</font></td>
							<td><font face="verdana" size="1">JOBCODE</font></td>
							<td>&nbsp;</td>
						</tr>
						<tr>
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
								<input class="datepicker" type="text" maxlength="10" name="txtDate" size="10" value="<%=v_txtDate%>" title="Type Date mm/dd/yyyy" >
							</td>
							<td>
								<select name="txtDivisi" >
								<Option Value="" <%If v_txtDivisi = "" then%> Selected <%End If%>>--ALL--</option>
							<%
							Dim rsDivisi, strDivisi
							strDivisi   =    " SELECT DISTINCT emp.sectionname as fccode, sc.fcname" & _
									         "    FROM employee emp" & _
									         "         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
									         "   WHERE fcba = '"& v_txtFcba &"' ORDER BY emp.sectionname "
							<!--strDivisi	=	" select distinct fccode, fcname, orderdata" & _                                                     -->
							<!--				" from (" & _                                                                                        -->                        
							<!--				"     select pd.fccode, pd.fcname, pd.rate_1 orderdata" & _                                          -->
							<!--				"     from parameter pr inner join parameterdetail pd" & _                                           -->  
							<!--				"     on (pr.fcba = pd.fcba and pr.parcode = pd.parcode and pr.fccode = pd.parheadcode)" & _         -->       
							<!--				"     inner join ( select fcba, fccode" & _                                                          -->
							<!--				"                         , rpt_upr_api.mapping_loccode_to_detail (fccode, 'SECTION') division" & _  --> 
							<!--				"                  from gcmaster" & _                                                                -->   
							<!--				"                  where fcba = '" & v_txtFcba & "'" & _                                             --> 
							<!--				"                        and item = 'LOC'" & _                                                       -->        
							<!--				"                        and activation = 'Y' ) loc" & _                                             -->
							<!--				"     on (pd.fcba = loc.fcba and pd.fccode = loc.division)" & _                                      --> 
							<!--				"     where pd.fcba = '" & v_txtFcba & "'" & _                                                       -->   
							<!--				"           and pd.parcode = 'RPT_UPR'" & _                                                          -->
							<!--				"           and pd.parheadcode = 'DIVISION'" & _                                                     -->  
							<!--				" )" & _
							<!--				" order by orderdata"
						   

							Set rsDivisi = Server.CreateObject("ADODB.Recordset")
							rsDivisi.Open strDivisi, connstr, adOpenStatic 

							if not rsDivisi.eof then
								Do While Not rsDivisi.eof
									if rsDivisi("fccode") = v_txtDivisi then%>
										<option value="<%=rsDivisi("fccode")%>" Selected><%=rsDivisi("fccode")%> : <%=rsDivisi("fcname")%></option> 
									<%Else%>
										<option value="<%=rsDivisi("fccode")%>"><%=rsDivisi("fccode")%> : <%=rsDivisi("fcname")%></option>  
									<%End if
									rsDivisi.movenext
								loop
							end if
							rsDivisi.Close
							set rsDivisi = Nothing
							%>
							</select>
							</td>
							<td>
								<select name="txtjobgroup" >
									<option value="">--ALL--</option>
								<%
								strJobCode = " SELECT DISTINCT jg.fccode, jg.fcname" & _
								             " FROM jobgroup jg" & _
								             " INNER JOIN job jj ON (jg.fcba = jj.fcba AND jj.jobgroupcode = jg.fccode)" & _
											 " WHERE jg.fcba = '" & v_txtFcba & "'" & _
											 "       AND jj.job_category IN ('FF', 'MN')" & _
											 " ORDER BY jg.fccode ASC"
								
								Set rsJobCode = Server.CreateObject("ADODB.Recordset")
								rsJobCode.Open strJobCode, connstr, adOpenStatic 
								
								if not rsJobCode.eof then
								Do While Not rsJobCode.eof
								if rsJobCode("FCCODE") = v_txtjobgroup then%>
										<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option> 
									 <%Else%>
										<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>  
									<%End if
									rsJobCode.movenext
								loop
								end if
								rsJobCode.Close
								set rsJobCode = Nothing
								%>
								</select>
							</td>
							<td>
								<select name="txtJobCode" >
									<option value="">--ALL--</option>
								<%
								Dim rsJobCode, strJobCode
								strJobCode	=	" SELECT fccode, fcname FROM job WHERE fcba = '" & v_txtFcba & "' AND job_category IN ('FF', 'MN')ORDER BY fccode ASC"
								
								Set rsJobCode = Server.CreateObject("ADODB.Recordset")
								rsJobCode.Open strJobCode, connstr, adOpenStatic 
								
								if not rsJobCode.eof then
								Do While Not rsJobCode.eof
								if rsJobCode("FCCODE") = v_txtJobCode then%>
										<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option> 
									 <%Else%>
										<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>  
									<%End if
									rsJobCode.movenext
								loop
								end if
								rsJobCode.Close
								set rsJobCode = Nothing
								%>
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
						<th colspan="<%=25%>" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
							<div id="header-wrapper">
								<div id="subheader" class="centered">
									<div style="margin-botton:10px;"/>
									<h3>PRODUKTIVITAS TENAGA KERJA</h3>
									<br/>
								</div>
								<div style="text-align:left;">
									<h4>Periode : <%=Day(v_txtDate) & " " & Bulan(v_txtMonth) & " " & v_txtYear%></h4>
									<h4 style="margin-top:5px;"><%=PCase(fcbaname)%></h4>
								</div>
							</div>
						</th>
					</tr>
					<tr>
						<th style="background-color:#00a300;color:#fff;" rowspan="3">Kode Pekerjaan</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="3">Jenis Pekerjaan</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="3">Divisi</th>
						<th style="background-color:#00a300;color:#fff;" colspan="7">HI</th>
						<th style="background-color:#00a300;color:#fff;" colspan="7">SHI</th>
						<th style="background-color:#00a300;color:#fff;" colspan="7">SBI</th>
						<th style="background-color:#00a300;color:#fff;" rowspan="2">BUDGET</th>
					</tr>
					<tr>
						<th style="background-color:#00a300;color:#fff;" colspan="3">Tenaga Kerja (HK)</th>
						<th style="background-color:#00a300;color:#fff;" colspan="3">Hasil Kerja (Unit)</th>
						<th style="background-color:#00a300;color:#fff;">Output</th>
						
						<th style="background-color:#00a300;color:#fff;" colspan="3">Tenaga Kerja (HK)</th>
						<th style="background-color:#00a300;color:#fff;" colspan="3">Hasil Kerja (Unit)</th>
						<th style="background-color:#00a300;color:#fff;">Output</th>
						
						<th style="background-color:#00a300;color:#fff;" colspan="3">Tenaga Kerja (HK)</th>
						<th style="background-color:#00a300;color:#fff;" colspan="3">Hasil Kerja (Unit)</th>
						<th style="background-color:#00a300;color:#fff;">Output</th>
					</tr>
					<tr>
						<th style="background-color:#00a300;color:#fff;">SKU</th>
						<th style="background-color:#00a300;color:#fff;">KHL</th>
						<th style="background-color:#00a300;color:#fff;">Total</th>
						<th style="background-color:#00a300;color:#fff;">Harian</th>
						<th style="background-color:#00a300;color:#fff;">Borongan</th>
						<th style="background-color:#00a300;color:#fff;">Total</th>
						<th style="background-color:#00a300;color:#fff;">HK/Unit</th>
						<th style="background-color:#00a300;color:#fff;">SKU</th>
						<th style="background-color:#00a300;color:#fff;">KHL</th>
						<th style="background-color:#00a300;color:#fff;">Total</th>
						<th style="background-color:#00a300;color:#fff;">Harian</th>
						<th style="background-color:#00a300;color:#fff;">Borongan</th>
						<th style="background-color:#00a300;color:#fff;">Total</th>
						<th style="background-color:#00a300;color:#fff;">HK/Unit</th>
						<th style="background-color:#00a300;color:#fff;">SKU</th>
						<th style="background-color:#00a300;color:#fff;">KHL</th>
						<th style="background-color:#00a300;color:#fff;">Total</th>
						<th style="background-color:#00a300;color:#fff;">Harian</th>
						<th style="background-color:#00a300;color:#fff;">Borongan</th>
						<th style="background-color:#00a300;color:#fff;">Total</th>
						<th style="background-color:#00a300;color:#fff;">HK/Unit</th>
						<th style="background-color:#00a300;color:#fff;">HK/Unit</th>
					</tr>
				</thead>
				
				<tbody>
<%

Function search_span(data_array__, col__, where__)
	result_ = 0
	arr_1 = data_array__
	for indx = LBound(arr_1, 2) to UBound(arr_1, 2)
		if arr_1(col__, indx) = where__ then
			result_ = result_ + 1
		end if
	next
	search_span = result_
End Function

Set rsLHPP = Server.CreateObject("ADODB.Recordset")
rsLHPP.Open sqlLHPP, connstr, adOpenStatic

if not rsLHPP.EOF then
	array_data = rsLHPP.getrows()
	array_eof = false
end if
rsLHPP.Close
set rsLHPP = Nothing
%>
<%if not array_eof then
	gt_sku = 0
	gt_khl = 0
	gt_unit_daily = 0
	gt_unit_contract = 0
	gt_sku_1 = 0
	gt_khl_1 = 0
	gt_unit_daily_1 = 0
	gt_unit_contract_1 = 0
	gt_sku_2 = 0
	gt_khl_2 = 0
	gt_unit_daily_2 = 0
	gt_unit_contract_2 = 0
	gt_bud_hk = 0
	gt_bud_ha = 0

	For xx = LBound(array_data, 2) To UBound(array_data, 2)
	
		curr_job = array_data(1, xx)
	
			r_hk_hi = 0
			r_hk_shi = 0
			r_hk_sbi = 0
			
			r_ha_hi = 0
			r_ha_shi = 0
			r_ha_sbi = 0
			
			r_hk_hi = Cdbl(array_data(4, xx)) + Cdbl(array_data(5, xx))
			r_hk_shi = Cdbl(array_data(8, xx)) + Cdbl(array_data(9, xx))
			r_hk_sbi = Cdbl(array_data(12, xx)) + Cdbl(array_data(13, xx))
			
			r_ha_hi = Cdbl(array_data(6, xx)) + Cdbl(array_data(7, xx))
			r_ha_shi = Cdbl(array_data(10, xx)) + Cdbl(array_data(11, xx))
			r_ha_sbi = Cdbl(array_data(14, xx)) + Cdbl(array_data(15, xx))
	
	
	%>
					<tr>
						<%if curr_job <> prev_job then%>
						<%span_job = search_span(array_data, 1, array_data(1, xx))%>
						<td align="center" rowspan="<%=span_job%>" valign="top"><%=array_data(1, xx)%></td>
						<td align="left" rowspan="<%=span_job%>" valign="top"><%=array_data(2, xx)%></td>
						<%end if%>
						<td align="center"><%=array_data(3, xx)%></td>
						<td align="right"><%=FormatNumber(array_data(4, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(5, xx), 2)%></td>
						<td align="right"><%=FormatNumber(r_hk_hi, 2)%></td>
						<td align="right"><%=FormatNumber(array_data(6, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(7, xx), 2)%></td>
						<td align="right"><%=FormatNumber(r_ha_hi, 2)%></td>
						
						<%if r_ha_hi = 0 then%>
							<td align="right"><%=FormatNumber(0, 2)%></td>
						<%else%>
							<td align="right"><%=FormatNumber(r_hk_hi/r_ha_hi, 2)%></td>
						<%end if%>
						
						<td align="right"><%=FormatNumber(array_data(8, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(9, xx), 2)%></td>
						<td align="right"><%=FormatNumber(r_hk_shi, 2)%></td>
						<td align="right"><%=FormatNumber(array_data(10, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(11, xx), 2)%></td>
						<td align="right"><%=FormatNumber(r_ha_shi, 2)%></td>
						
						<%if r_ha_shi = 0 then%>
							<td align="right"><%=FormatNumber(0, 2)%></td>
						<%else%>
							<td align="right"><%=FormatNumber(r_hk_shi/r_ha_shi, 2)%></td>
						<%end if%>
						
						<td align="right"><%=FormatNumber(array_data(12, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(13, xx), 2)%></td>
						<td align="right"><%=FormatNumber(r_hk_sbi, 2)%></td>
						<td align="right"><%=FormatNumber(array_data(14, xx), 2)%></td>
						<td align="right"><%=FormatNumber(array_data(15, xx), 2)%></td>
						<td align="right"><%=FormatNumber(r_ha_sbi, 2)%></td>
						
						<%if r_ha_sbi = 0 then%>
							<td align="right"><%=FormatNumber(0, 2)%></td>
						<%else%>
							<td align="right"><%=FormatNumber(r_hk_sbi/r_ha_sbi, 2)%></td>
						<%end if%>
						
						<%if cdbl(array_data(17, xx)) = 0 then%>
							<td align="right"><%=FormatNumber(0, 2)%></td>
						<%else%>
							<td align="right"><%=FormatNumber(cdbl(array_data(16, xx))/cdbl(array_data(17, xx)), 2)%></td>
						<%end if%>
					
					</tr>
					
<%
		gt_sku = gt_sku + Cdbl(array_data(4, xx))
		gt_khl = gt_khl + Cdbl(array_data(5, xx))
		gt_unit_daily = gt_unit_daily + Cdbl(array_data(6, xx))
		gt_unit_contract = gt_unit_contract + Cdbl(array_data(7, xx))
		gt_sku_1 = gt_sku_1 + Cdbl(array_data(8, xx))
		gt_khl_1 = gt_khl_1 + Cdbl(array_data(9, xx))
		gt_unit_daily_1 = gt_unit_daily_1 + Cdbl(array_data(10, xx))
		gt_unit_contract_1 = gt_unit_contract_1 + Cdbl(array_data(11, xx))
		gt_sku_2 = gt_sku_2 + Cdbl(array_data(12, xx))
		gt_khl_2 = gt_khl_2 + Cdbl(array_data(13, xx))
		gt_unit_daily_2 = gt_unit_daily_2 + Cdbl(array_data(14, xx))
		gt_unit_contract_2 = gt_unit_contract_2 + Cdbl(array_data(15, xx))
		gt_bud_hk = gt_bud_hk + Cdbl(array_data(16, xx))
		gt_bud_ha = gt_bud_ha + Cdbl(array_data(17, xx))
		
		prev_job = curr_job

	next
	
	gt_hk_hi = 0
	gt_hk_shi = 0
	gt_hk_sbi = 0

	gt_ha_hi = 0
	gt_ha_shi = 0
	gt_ha_sbi = 0

	gt_hk_hi = gt_sku + gt_khl
	gt_hk_shi = gt_sku_1 + gt_khl_1
	gt_hk_sbi = gt_sku_2 + gt_khl_2

	gt_ha_hi = gt_unit_daily + gt_unit_contract
	gt_ha_shi = gt_unit_daily_1 + gt_unit_contract_1
	gt_ha_sbi = gt_unit_daily_2 + gt_unit_contract_2
%>
					<tr><td colspan="25"></td></tr>
					<tr style="background-color:#f4f4a0;">					
						<td align="center" colspan="3"><b>Total</b></td>
						<td align="right"><b><%=FormatNumber(gt_sku, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_khl, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_hk_hi, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_daily, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_contract, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_ha_hi, 2)%></b></td>

						<%if gt_ha_hi = 0 then%>
							<td align="right"><b><%=FormatNumber(0, 2)%></b></td>
						<%else%>
							<td align="right"><b><%=FormatNumber(gt_hk_hi/gt_ha_hi, 2)%></b></td>
						<%end if%>
						
						<td align="right"><b><%=FormatNumber(gt_sku_1, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_khl_1, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_hk_shi, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_daily_1, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_contract_1, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_ha_shi, 2)%></b></td>
						
						<%if gt_ha_shi = 0 then%>
							<td align="right"><b><%=FormatNumber(0, 2)%></b></td>
						<%else%>
							<td align="right"><b><%=FormatNumber(gt_hk_shi/gt_ha_shi, 2)%></b></td>
						<%end if%>
						
						<td align="right"><b><%=FormatNumber(gt_sku_2, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_khl_2, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_hk_sbi, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_daily_2, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_unit_contract_2, 2)%></b></td>
						<td align="right"><b><%=FormatNumber(gt_ha_sbi, 2)%></b></td>
						
						<%if gt_ha_sbi = 0 then%>
							<td align="right"><b><%=FormatNumber(0, 2)%></b></td>
						<%else%>
							<td align="right"><b><%=FormatNumber(gt_hk_sbi/gt_ha_sbi, 2)%></b></td>
						<%end if%>
						
						<%if gt_bud_ha = 0 then%>
							<td align="right"><b><%=FormatNumber(0, 2)%></b></td>
						<%else%>
							<td align="right"><b><%=FormatNumber(gt_bud_hk/gt_bud_ha, 2)%></b></td>
						<%end if%>
					
					</tr>
<%else%>
					<tr>
						<td colspan="<%=25%>">NO DATA AVAILABLE</td>
					</tr>
<%end if%>
				</tbody>
			</table>

		</div>
	</body>
</html>
