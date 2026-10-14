<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtTanggal, v_txtUraian, v_TxtTypeFCBA
Dim ScriptName
Dim rsDKWH, strDKWH, strDKWH2, strMain
Dim cClass, nday

Dim var_ColSpan : var_ColSpan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_TxtTypeFCBA			= Request("TxtTypeFCBA")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

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
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_ColSpan	=	var_ColSpan	+ 1
		
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_ColSpan	=	var_ColSpan	+ 1
		
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
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

v_txtTanggal	=	Request("txtTanggal")
v_txtUraian		=	Request("txtUraian")


If (v_txtTanggal = "" Or Len(Trim(v_txtTanggal)) = 0) Then
	v_txtTanggal = textDate(Date)
End If

If v_txtUraian = "" Then
	v_txtUraian	=	"--ALL--"
End If

If Len(Trim(v_txtUraian)) > 0 Then
	if v_txtUraian = "--ALL--" then
		whereCons = wherecons & " "
	Elseif v_txtUraian = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		whereCons = whereCons & " And  URAIAN = '" & v_txtUraian & "' "
	end if
End If

Dim v_Tahun : v_Tahun = Year(v_txtTanggal)
Dim v_Bulan : v_Bulan = Month(v_txtTanggal)

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT " & strFieldCons & "," & _
			"          SEQ," & _
			"          URAIAN," & _
			"          SUM (HI) HI," & _
			"          SUM (SHI) SHI," & _
			"          SUM (SBI) SBI" & _
			"     FROM (  SELECT FCBA," & _
			"                    SEQ," & _
			"                    URAIAN," & _
			"                    SUM (JAM) HI," & _
			"                    0 SHI," & _
			"                    0 SBI" & _
			"               FROM (SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            'TBM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            'TM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'KENDARAAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'POM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 6 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MN'" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TBM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'KENDARAAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'VH'" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'POM' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MA'" & _
			"                     UNION ALL" & _
			"                     SELECT 6 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE NOT IN ('MN', 'FF', 'VH', 'MA'))" & _
			"              WHERE TANGGAL = TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
			"           GROUP BY FCBA, SEQ, URAIAN" & _
			"           UNION ALL" & _
			"             SELECT FCBA," & _
			"                    SEQ," & _
			"                    URAIAN," & _
			"                    0 HI," & _
			"                    SUM (JAM) SHI," & _
			"                    0 SBI" & _
			"               FROM (SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            'TBM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            'TM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'KENDARAAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'POM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 6 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MN'" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TBM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'KENDARAAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'VH'" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'POM' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MA'" & _
			"                     UNION ALL" & _
			"                     SELECT 6 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE NOT IN ('MN', 'FF', 'VH', 'MA'))" & _
			"              WHERE     TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY')) = " & v_Tahun & "" & _
			"                    AND TO_NUMBER (TO_CHAR (TANGGAL, 'MM')) = " & v_Bulan & "" & _
			"                    AND TANGGAL <= TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
			"           GROUP BY FCBA, SEQ, URAIAN" & _
			"           UNION ALL" & _
			"             SELECT FCBA," & _
			"                    SEQ," & _
			"                    URAIAN," & _
			"                    0 HI," & _
			"                    0 SHI," & _
			"                    SUM (JAM) SBI" & _
			"               FROM (SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            'TBM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            'TM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'KENDARAAN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'POM' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 6 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCCODE FCBA," & _
			"                            TO_DATE ('" & v_txtTanggal & "', 'MM/DD/YYYY') TANGGAL," & _
			"                            0 JAM" & _
			"                       FROM BUSINESSUNIT" & _
			"                     UNION ALL" & _
			"                     SELECT 1 SEQ," & _
			"                            'BIBITAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MN'" & _
			"                     UNION ALL" & _
			"                     SELECT 2 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TBM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 3 SEQ," & _
			"                            FF.STATUS," & _
			"                            WA.FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM    WORKSHOPACTIVITY WA" & _
			"                            INNER JOIN" & _
			"                               FIELD FF" & _
			"                            ON (WA.FCBA = FF.FCBA AND WA.LOCATIONCODE = FF.FCCODE)" & _
			"                      WHERE     LOCATIONTYPE = 'FF'" & _
			"                            AND FF.STATUS = 'TM'" & _
			"                            AND ACTIVATION = 'Y'" & _
			"                     UNION ALL" & _
			"                     SELECT 4 SEQ," & _
			"                            'KENDARAAN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'VH'" & _
			"                     UNION ALL" & _
			"                     SELECT 5 SEQ," & _
			"                            'POM' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE = 'MA'" & _
			"                     UNION ALL" & _
			"                     SELECT 6 SEQ," & _
			"                            'LAIN-LAIN' URAIAN," & _
			"                            FCBA," & _
			"                            WORKSHOPDATE TANGGAL," & _
			"                            NVL (UNIT, 0) JAM" & _
			"                       FROM WORKSHOPACTIVITY" & _
			"                      WHERE LOCATIONTYPE NOT IN ('MN', 'FF', 'VH', 'MA'))" & _
			"              WHERE TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY')) = " & v_Tahun & "" & _
			"                    AND TO_NUMBER (TO_CHAR (TANGGAL, 'MM')) <= " & v_Bulan & "" & _
			"           GROUP BY FCBA, SEQ, URAIAN) where USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = '" & v_TxtTypeFCBA &"' " & _
			" GROUP BY FCBA, SEQ, URAIAN" & _
			" ORDER BY FCBA, SEQ, URAIAN"

			
strDKWH2	=	"        , SEQ" & _
				"        , URAIAN" & _
				"        , SUM (HI) HI" & _
				"        , SUM (SHI) SHI" & _
				"        , SUM (SBI) SBI"
			
			
strDKWH	=	" select GroupPlantation " & selectCons & strDKWH2 & "  From (" & strMain & " )" & _
			whereCons & _
			" group by GroupPlantation" & selectCons & " , SEQ, URAIAN" & _
			" Order by GroupPlantation" & selectCons & " , SEQ"
						
'response.write strDKWH
%>
<html>

<head>
<title>ePlaS Reporting - Distribusi Kerja Workshop Harian</title>
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
@media print
{
.printhide{display:none;}
}
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
				<input type="hidden" name="TxtTypeFCBA" value="<%=v_TxtTypeFCBA%>" >
				<!--#include virtual="/REPORT/FormSelectNewSumME.html"-->
				<table id="form-table" class="centered-margin">	
					<tr>
						<td><font face="verdana" size="1">TANGGAL</font></td>
						<td align="center">
						<% if v_txtTanggal <> "" and v_txtTanggal = v_txtTanggal then %>
							<input class="datepicker" type="text" maxlength="10" name="txtTanggal" size="10" value="<%=v_txtTanggal%>" title="Type Date mm/dd/yyyy" >
						<% else %>
							<input class="datepicker" type="text" maxlength="10" name="txtTanggal" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
						<% end if%>
							<!--<a href="javascript:show_calendar('frmEntry.txtTanggal');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
								<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
							</a>-->
						</td>
						<td><font face="verdana" size="1">URAIAN</font></td>
						<td>
							<select name="txtUraian" >
								<option value="--ALL--"<%If v_txtUraian = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<option value="BIBITAN"<%If v_txtUraian = "BIBITAN" then%> Selected <%End If%>>BIBITAN</option>
								<option value="TBM"<%If v_txtUraian = "TBM" then%> Selected <%End If%>>TBM</option>
								<option value="TM"<%If v_txtUraian = "TM" then%> Selected <%End If%>>TM</option>
								<option value="KENDARAAN"<%If v_txtUraian = "KENDARAAN" then%> Selected <%End If%>>KENDARAAN</option>
								<option value="POM"<%If v_txtUraian = "POM" then%> Selected <%End If%>>POM</option>
								<option value="LAIN-LAIN"<%If v_txtUraian = "LAIN-LAIN" then%> Selected <%End If%>>LAIN-LAIN</option>
							</select>
						</td>
						<td><input type="submit" value="View" /></td>
					</tr>				
				</table>
			</div>
			</form>
 <table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
        <tr>
            <th colspan="13" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
			<div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
			</div>
			<div id="subheader" class="centered">
				<h2>DISTRIBUSI KERJA WORKSHOP HARIAN</h2>
				<h5>TANGGAL: <%=v_txtTanggal%></h5>
			</div>
		</div>
        </th>
    </tr>
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">URAIAN</th>

		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
		
		<th colspan="6" style="background-color:#00A300;color:#fff;">JAM KERJA</th>
	</tr>
	<tr>
		<th colspan="2" style="background-color:#00A300;color:#fff;">HARI INI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">SAMPAI HARI INI</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">SAMPAI BULAN INI</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;"><a class="quick_link" href="<%IF v_TxtTypeFCBA = "M" then %>p_DistWorkshopPerMesin.asp<%ELSE%>p_DistWorkshopPerKend.asp<%END IF%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtTanggal=<%=v_txtTanggal%>&txtGroupPlantation=<%=v_txtGroupPlantation%>&txtUraian=<%=v_txtUraian%>" target="_blank">JAM</a></th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;"><a class="quick_link" href="<%IF v_TxtTypeFCBA = "M" then %>p_DistWorkshopPerMesin.asp<%ELSE%>p_DistWorkshopPerKend.asp<%END IF%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtTanggal=<%=v_txtTanggal%>&txtGroupPlantation=<%=v_txtGroupPlantation%>&txtUraian=<%=v_txtUraian%>" target="_blank">JAM</a></th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;"><a class="quick_link" href="<%IF v_TxtTypeFCBA = "M" then %>p_DistWorkshopPerMesin.asp<%ELSE%>p_DistWorkshopPerKend.asp<%END IF%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtplantation=<%=v_txtplantation%>&txtgroupregion=<%=v_txtgroupregion%>&txtregion=<%=v_txtregion%>&txtcompany=<%=v_txtcompany%>&txtfcba=<%=v_txtfcba%>&txtTanggal=<%=v_txtTanggal%>&txtGroupPlantation=<%=v_txtGroupPlantation%>&txtUraian=<%=v_txtUraian%>" target="_blank">JAM</a></th>
		<th style="background-color:#00A300;color:#fff;">%</th>
	</tr>
    </thead>
<tbody>

<%
Set rsDKWH = Server.CreateObject("ADODB.Recordset")
rsDKWH.Open strDKWH, connstr, adOpenStatic

Dim rsGetTotal, strGetTotal
strGetTotal = " Select Sum(HI) HI, Sum(SHI) SHI, Sum(SBI) SBI FROM (" & strDKWH & ")"
Set rsGetTotal = Server.CreateObject("ADODB.Recordset")
rsGetTotal.Open strGetTotal, connstr, adOpenStatic

Dim Temp_Total_HI : Temp_Total_HI = 0
Dim Temp_Total_SHI : Temp_Total_SHI = 0
Dim Temp_Total_SBI : Temp_Total_SBI = 0

Dim	GrandTotal_HI : GrandTotal_HI = 0
Dim GrandTotal_SHI : GrandTotal_SHI = 0
Dim GrandTotal_SBI : GrandTotal_SBI = 0

Dim Total_Persen_HI : Total_Persen_HI = 0
Dim Total_Persen_SHI : Total_Persen_SHI = 0
Dim Total_Persen_SBI : Total_Persen_SBI = 0

Dim GrandTotal_Persen_HI : GrandTotal_Persen_HI = 0
Dim GrandTotal_Persen_SHI : GrandTotal_Persen_SHI = 0
Dim GrandTotal_Persen_SBI : GrandTotal_Persen_SBI = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempUraian : tempUraian = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim SprevUraian : SprevUraian = ""

If Not rsGetTotal.EOF Then
	Temp_Total_HI = Temp_Total_HI + Cdbl(rsGetTotal("HI"))
	Temp_Total_SHI = Temp_Total_SHI + Cdbl(rsGetTotal("SHI"))
	Temp_Total_SBI = Temp_Total_SBI + Cdbl(rsGetTotal("SBI"))
Else
	Temp_Total_HI = 0
	Temp_Total_SHI = 0
	Temp_Total_SBI = 0
End If
rsGetTotal.Close
set rsGetTotal = Nothing

if (rsDKWH.bof and rsDKWH.eof) then
%>
	<tr>
		<td colspan="<%=1+var_ColSpan+8%>">NO DATA AVAILABLE</td>
	</tr>
<%else
		while not rsDKWH.eof
			Total_Persen_HI = 0
			Total_Persen_SHI = 0
			Total_Persen_SBI = 0
			
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<% if v_txtUraian <>  "--SUMMARY--" and  v_txtUraian <>  "" then %>
					<%	tempUraian = rsDKWH("URAIAN")
					If Cint(visibleP) > 1 Then
						tempGroupPlantation = "--ALL--"
					Else
						If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then
							tempGroupPlantation = rsDKWH("GroupPlantation")
						Else
							tempGroupPlantation = ""
						End If
					End If
					
					If Cint(visibleGR) > 1 And Cint(visibleP) > 1 Then
						tempPlantation = "--SUMMARY--"
					Else
						If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then
							tempPlantation = rsDKWH("Plantation")
						Else
							tempPlantation = ""
						End If
					End If
					
					If Cint(visibleR) > 1 And Cint(visibleGR) > 1 Then
						tempGroupRegion = "--SUMMARY--"
					Else
						If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then
							tempGroupRegion = rsDKWH("GroupRegion")
						Else
							tempGroupRegion = ""
						End If
					End If
					
					If Cint(visibleC) > 1 And Cint(visibleR) > 1 Then
						tempRegion = "--SUMMARY--"
					Else
						If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then
							tempRegion = rsDKWH("Region")
						Else
							tempRegion = ""
						End If
					End If
					
					If Cint(visibleS) > 1 And Cint(visibleC) > 1 Then
						tempCompany = "--SUMMARY--"
					Else
						If Len(Trim(v_txtcompany)) > 0 And v_txtcompany <>  "--SUMMARY--" Then
							tempCompany = rsDKWH("Company")
						Else
							tempCompany = ""
						End If
					End If
					
					If Cint(VisibleS) > 1 Then
						tempfcba = "--SUMMARY--"
					Else
						If Len(Trim(v_txtFcba)) > 0 And v_txtFcba <>  "--SUMMARY--" Then
							tempfcba = rsDKWH("Fcba")
						Else
							tempfcba = ""
						End If
					End If
					%>
						<% if tempUraian <> SprevUraian then %>
							<td align="left">
								<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
								<%=rsDKWH("URAIAN")%>
								</a>
							</td>
							<% SprevUraian = tempUraian %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
				<% end if %>

				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   tempGroupPlantation = rsDKWH("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsDKWH("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDKWH("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsDKWH("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDKWH("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDKWH("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsDKWH("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDKWH("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDKWH("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsDKWH("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsDKWH("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
							<%=rsDKWH("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtUraian <>  "--SUMMARY--"  and  v_txtUraian <> "" then
					      tempUraian = rsDKWH("URAIAN")
					   else
					      tempUraian = ""
					   end if
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsDKWH("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsDKWH("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsDKWH("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsDKWH("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsDKWH("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsDKWH("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?TxtTypeFCBA=<%=v_TxtTypeFCBA%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtTanggal=<%=v_txtTanggal%>&txtUraian=<%=tempUraian%>">
						    <%=rsDKWH("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
<%
			If Temp_Total_HI = 0 Then
				Total_Persen_HI	= 0
				GrandTotal_Persen_HI = 0
			Else
				Total_Persen_HI	= Round((Cdbl(rsDKWH("HI"))/Temp_Total_HI)*100, 2)
				GrandTotal_Persen_HI	=	GrandTotal_Persen_HI + (Cdbl(rsDKWH("HI"))/Temp_Total_HI)*100
			End If
			If Temp_Total_SHI = 0 Then
				Total_Persen_SHI = 0
				GrandTotal_Persen_SHI = 0
			Else
				Total_Persen_SHI = Round((Cdbl(rsDKWH("SHI"))/Temp_Total_SHI)*100, 2)
				GrandTotal_Persen_SHI	=	GrandTotal_Persen_SHI + (Cdbl(rsDKWH("SHI"))/Temp_Total_SHI)*100
			End If
			If Temp_Total_SBI = 0 Then
				Total_Persen_SBI = 0
				GrandTotal_Persen_SBI = 0
			Else
				Total_Persen_SBI = Round((Cdbl(rsDKWH("SBI"))/Temp_Total_SBI)*100, 2)
				GrandTotal_Persen_SBI	=	GrandTotal_Persen_SBI + (Cdbl(rsDKWH("SBI"))/Temp_Total_SBI)*100
			End If
%>
				
				<td align="right"><%=FormatNumber(rsDKWH("HI"), 0)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_HI, 0)%></td>

				<td align="right"><%=FormatNumber(rsDKWH("SHI"), 0)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_SHI, 0)%></td>

				<td align="right"><%=FormatNumber(rsDKWH("SBI"), 0)%></td>
				<td align="right"><%=FormatNumber(Total_Persen_SBI, 0)%></td>
			</tr>
<%
			GrandTotal_HI	=	GrandTotal_HI + Cdbl(rsDKWH("HI"))
			GrandTotal_SHI	=	GrandTotal_SHI + Cdbl(rsDKWH("SHI"))
			GrandTotal_SBI	=	GrandTotal_HI + Cdbl(rsDKWH("SBI"))
			rsDKWH.MoveNext
		Wend
	rsDKWH.Close
	set rsDKWH = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td align="center" colspan="<%=1+var_ColSpan%>"><strong>TOTAL</strong></td>
		<td align="right"><strong><%=FormatNumber(Temp_Total_HI, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_HI, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Temp_Total_SHI, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_SHI, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(Temp_Total_SBI, 0)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandTotal_Persen_SBI, 0)%></strong></td>
	</tr>
<%End if%>
</tbody>
</table>
		</div>
	</body>
</html>