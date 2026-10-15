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

v_txtDate = Request("txtDate")

If "" & v_txtDate = "" Then v_txtDate = TextDate(date)
If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))

v_txtMonth = month(v_txtDate)
v_txtDay = DAY(v_txtDate)
v_txtYear =	year(v_txtDate)
	
'v_txtYear		=	Request("txtYear")
'v_txtMonth		=	Request("txtMonth")
v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")
v_txtFcba		=	Request("txtFcba")
v_txtColor		=	Request("txtColor")

Dim v_txtPersen : v_txtPersen = Request("txtPersen")
If Len(trim(v_txtPersen)) = 0 Then v_txtPersen = 0


strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

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



ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strPusingan	=	" SELECT" & _
				"      hi.fcba" & _
				"      , TO_NUMBER (TO_CHAR (hi.ondate, 'yyyy')) tahun" & _
				"      , TO_NUMBER (TO_CHAR (hi.ondate, 'mm')) bulan" & _
				"      , field_api.getvalid_afdeling(hi.fieldcode, hi.fcba, hi.ondate) division" & _
				"      , field_api.getvalid_harvestingsection(hi.fieldcode, hi.fcba, hi.ondate) harvestingsection" & _
				"      , field_api.getvalid_harvestingsectionnumb(hi.fieldcode, hi.fcba, hi.ondate) harvestingsectionnumb" & _
				"      , hi.fieldcode" & _
				"      , TO_NUMBER (TO_CHAR (field_api.getvalid_plantingdate(hi.fieldcode, hi.fcba, hi.ondate), 'yyyy')) tahuntanam" & _
				"      , field_api.getvalid_hectarageplanted(hi.fieldcode, hi.fcba, hi.ondate) hectarageplanted" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 1, ORIGINALNUMBER, 0)), 0) DAY_1" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 1, COLORIDENTIFICATION, 0)) COLOR_1" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 2, ORIGINALNUMBER, 0)), 0) DAY_2" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 2, COLORIDENTIFICATION, 0)) COLOR_2" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 3, ORIGINALNUMBER, 0)), 0) DAY_3" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 3, COLORIDENTIFICATION, 0)) COLOR_3" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 4, ORIGINALNUMBER, 0)), 0) DAY_4" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 4, COLORIDENTIFICATION, 0)) COLOR_4" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 5, ORIGINALNUMBER, 0)), 0) DAY_5" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 5, COLORIDENTIFICATION, 0)) COLOR_5" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 6, ORIGINALNUMBER, 0)), 0) DAY_6" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 6, COLORIDENTIFICATION, 0)) COLOR_6" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 7, ORIGINALNUMBER, 0)), 0) DAY_7" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 7, COLORIDENTIFICATION, 0)) COLOR_7" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 8, ORIGINALNUMBER, 0)), 0) DAY_8" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 8, COLORIDENTIFICATION, 0)) COLOR_8" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 9, ORIGINALNUMBER, 0)), 0) DAY_9" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 9, COLORIDENTIFICATION, 0)) COLOR_9" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 10, ORIGINALNUMBER, 0)), 0) DAY_10" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 10, COLORIDENTIFICATION, 0)) COLOR_10" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 11, ORIGINALNUMBER, 0)), 0) DAY_11" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 11, COLORIDENTIFICATION, 0)) COLOR_11" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 12, ORIGINALNUMBER, 0)), 0) DAY_12" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 12, COLORIDENTIFICATION, 0)) COLOR_12" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 13, ORIGINALNUMBER, 0)), 0) DAY_13" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 13, COLORIDENTIFICATION, 0)) COLOR_13" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 14, ORIGINALNUMBER, 0)), 0) DAY_14" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 14, COLORIDENTIFICATION, 0)) COLOR_14" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 15, ORIGINALNUMBER, 0)), 0) DAY_15" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 15, COLORIDENTIFICATION, 0)) COLOR_15" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 16, ORIGINALNUMBER, 0)), 0) DAY_16" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 16, COLORIDENTIFICATION, 0)) COLOR_16" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 17, ORIGINALNUMBER, 0)), 0) DAY_17" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 17, COLORIDENTIFICATION, 0)) COLOR_17" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 18, ORIGINALNUMBER, 0)), 0) DAY_18" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 18, COLORIDENTIFICATION, 0)) COLOR_18" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 19, ORIGINALNUMBER, 0)), 0) DAY_19" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 19, COLORIDENTIFICATION, 0)) COLOR_19" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 20, ORIGINALNUMBER, 0)), 0) DAY_20" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 20, COLORIDENTIFICATION, 0)) COLOR_20" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 21, ORIGINALNUMBER, 0)), 0) DAY_21" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 21, COLORIDENTIFICATION, 0)) COLOR_21" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 22, ORIGINALNUMBER, 0)), 0) DAY_22" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 22, COLORIDENTIFICATION, 0)) COLOR_22" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 23, ORIGINALNUMBER, 0)), 0) DAY_23" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 23, COLORIDENTIFICATION, 0)) COLOR_23" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 24, ORIGINALNUMBER, 0)), 0) DAY_24" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 24, COLORIDENTIFICATION, 0)) COLOR_24" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 25, ORIGINALNUMBER, 0)), 0) DAY_25" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 25, COLORIDENTIFICATION, 0)) COLOR_25" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 26, ORIGINALNUMBER, 0)), 0) DAY_26" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 26, COLORIDENTIFICATION, 0)) COLOR_26" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 27, ORIGINALNUMBER, 0)), 0) DAY_27" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 27, COLORIDENTIFICATION, 0)) COLOR_27" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 28, ORIGINALNUMBER, 0)), 0) DAY_28" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 28, COLORIDENTIFICATION, 0)) COLOR_28" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 29, ORIGINALNUMBER, 0)), 0) DAY_29" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 29, COLORIDENTIFICATION, 0)) COLOR_29" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 30, ORIGINALNUMBER, 0)), 0) DAY_30" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 30, COLORIDENTIFICATION, 0)) COLOR_30" & _
				"      , ROUND( SUM(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 31, ORIGINALNUMBER, 0)), 0) DAY_31" & _
				"      , MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 31, COLORIDENTIFICATION, 0)) COLOR_31" & _
				"      , new_pusingan_api.GetLastInterval(hi.fcba, hi.fieldcode, to_date('" & v_txtDate & "', 'mm/dd/yyyy')) lastinterval" & _
				" FROM " & _
				"      new_pusingan_tab hi" & _
				" WHERE " & _
				"      field_api.getvalid_activation(hi.fieldcode, hi.fcba, hi.ondate) = 'Y'" & _
				"      AND field_api.getvalid_isactive_harvested(hi.fieldcode, hi.fcba, hi.ondate) = 'TRUE'" & _
				"      AND field_api.getvalid_harvestingsection(hi.fieldcode, hi.fcba, hi.ondate) IS NOT NULL" & _
				"      AND hi.fcba = '" & v_txtFcba & "'"
				
If Len(Trim(v_txtYear)) > 0 Then
	strPusingan	=	strPusingan	&	"      AND TO_NUMBER (TO_CHAR (hi.ondate, 'YYYY')) = " & v_txtYear
End If
If Len(Trim(v_txtMonth)) > 0 Then
	strPusingan	=	strPusingan	&	"      AND TO_NUMBER (TO_CHAR (hi.ondate, 'MM')) = " & v_txtMonth
End If		
If v_txtAfdeling <> "ALL" And Len(Trim(v_txtAfdeling)) > 0 Then
	strPusingan	=	strPusingan	&	"      AND field_api.getvalid_afdeling(hi.fieldcode, hi.fcba, hi.ondate) = '" & v_txtAfdeling & "'"
Else
	strPusingan	=	strPusingan	&	" "
End If
If v_txtBlok <> "ALL" And Len(Trim(v_txtBlok)) > 0 Then
	strPusingan	=	strPusingan	&	"      AND hi.fieldcode = '" & v_txtBlok & "'"
Else
	strPusingan	=	strPusingan	&	" "
End If

If "" & v_txtColor <> "" Then
	qSQL = " SELECT DISTINCT fieldcode" & _
		   " FROM new_pusingan_tab" & _
		   " WHERE fcba = '" & v_txtFcba & "'" & _
		   "       AND TO_CHAR(ondate, 'yyyymm') = '" & v_txtYear & "'||LPAD('" & v_txtMonth & "', 2, 0)" & _
		   "       AND ondate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" &_
		   "       AND coloridentification = '" & v_txtColor & "'"
	Set qORS = Server.CreateObject("ADODB.Recordset")
	qORS.Open qSQL, connstr, adOpenStatic

	arr_yellow_eof = true
	If Not qORS.EOF Then
		arr_yellow = qORS.getrows()
		arr_yellow_eof = false
	end if
	qORS.Close
	set qORS = Nothing
	
	if not arr_yellow_eof then
		filterbycolor = ""
		for ii=0 to ubound(arr_yellow,2)
			f_color = ""
			if ii = ubound(arr_yellow,2) then
				f_color = "'" & arr_yellow(0,ii) & "'"
			else
				f_color = "'" & arr_yellow(0,ii) & "',"
			end if
			filterbycolor = filterbycolor & f_color
		next
	end if
	strPusingan = strPusingan & "       AND fieldcode IN (" & filterbycolor & ")"
End If

	strPusingan	=	strPusingan	&	" GROUP BY " & _
									"      hi.fcba" & _
									"      , to_number (to_char (hi.ondate, 'yyyy'))" & _
									"      , to_number (to_char (hi.ondate, 'mm'))" & _
									"      , field_api.getvalid_afdeling(hi.fieldcode, hi.fcba, hi.ondate)" & _
									"      , field_api.getvalid_harvestingsection(hi.fieldcode, hi.fcba, hi.ondate)" & _
									"      , field_api.getvalid_harvestingsectionnumb(hi.fieldcode, hi.fcba, hi.ondate)" & _
									"      , hi.fieldcode" & _
									"      , field_api.getvalid_hectarageplanted(hi.fieldcode, hi.fcba, hi.ondate)" & _
									"      , to_number (to_char (field_api.getvalid_plantingdate(hi.fieldcode, hi.fcba, hi.ondate), 'yyyy'))"
	
If Len(Trim(v_txtColor)) > 0 Then
	strPusingan	=	strPusingan	& " HAVING MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 1, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 2, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 3, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 4, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 5, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 6, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 7, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 8, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 9, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 10, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 11, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 12, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 13, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 14, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 15, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 16, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 17, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 18, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 19, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 20, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 21, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 22, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 23, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 24, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 25, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 26, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 27, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 28, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 29, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 30, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'" & _
					" OR MAX(DECODE (TO_NUMBER (TO_CHAR (HI.ONDATE, 'DD')), 31, COLORIDENTIFICATION, 0)) = '" & v_txtColor & "'"
End If

	strPusingan	=	strPusingan	&	" ORDER BY " & _
									"      hi.fcba" & _
									"      , to_number (to_char (hi.ondate, 'yyyy'))" & _
									"      , to_number (to_char (hi.ondate, 'mm'))" & _
									"      , field_api.getvalid_afdeling(hi.fieldcode, hi.fcba, hi.ondate)" & _
									"      , field_api.getvalid_harvestingsection(hi.fieldcode, hi.fcba, hi.ondate)" & _
									"      , field_api.getvalid_harvestingsectionnumb(hi.fieldcode, hi.fcba, hi.ondate)"
									
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

' response.write strPusingan


qSQL = " SELECT fieldcode, NVL(SUM(NVL(hectare_harvested, 0)), 0) hectare_harvested" & _
       " FROM new_pusingan_tab" & _
       " WHERE fcba = '" & v_txtFcba & "'" & _
       "       AND TO_CHAR(ondate, 'yyyymm') = '" & v_txtYear & "'||LPAD('" & v_txtMonth & "', 2, 0)" & _
       "       AND ondate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')" & _
       " GROUP BY fieldcode"
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open qSQL, connstr, adOpenStatic

arr_rotasi_eof = true
If Not qORS.EOF Then
	arr_rotasi = qORS.getrows()
	arr_rotasi_eof = false
end if
qORS.Close
set qORS = Nothing

Function f_rotasi(data_array__, fieldcode__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if Cstr(arr_1(0, indx)) = Cstr(fieldcode__) then
			result_ = Cdbl(To_Nol(arr_1(1, indx)))
		end if
	next
	f_rotasi = result_
End Function

qSQL = " SELECT TO_NUMBER(TO_CHAR(ondate, 'dd')) ondate, fieldcode, CASE WHEN hectare_harvested IS NULL THEN originalnumber WHEN hectare_harvested = 0 THEN originalnumber ELSE ROUND(hectare_harvested, 2) END hectare_harvested, field_api.getvalid_hectarageplanted (fieldcode, fcba, ondate) ha_planted" & _
       " FROM new_pusingan_tab" & _
       " WHERE fcba = '" & v_txtFcba & "'" & _
       "       AND TO_CHAR(ondate, 'yyyymm') = '" & v_txtYear & "'||LPAD('" & v_txtMonth & "', 2, 0)" & _
       "       AND ondate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')"
' response.write qSQL
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open qSQL, connstr, adOpenStatic

arr_hectare_eof = true
If Not qORS.EOF Then
	arr_hectare = qORS.getrows()
	arr_hectare_eof = false
end if
qORS.Close
set qORS = Nothing

Function f_hectare(data_array__, day__, fieldcode__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if Cstr(arr_1(0, indx)) = Cstr(day__) and Cstr(arr_1(1, indx)) = Cstr(fieldcode__) then
			result_ = Cdbl(To_Nol(arr_1(2, indx)))
		end if
	next
	f_hectare = result_
End Function

' Addded Function By SAG 22 Agu 2019
Function f_hectare2(data_array__, day__, fieldcode__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if Cstr(arr_1(0, indx)) = Cstr(day__) and Cstr(arr_1(1, indx)) = Cstr(fieldcode__) then
			result_ = Cdbl(To_Nol(arr_1(3, indx)))
		end if
	next
	f_hectare2 = result_
End Function
' End Added Function By SAG

qSQL = " SELECT TO_NUMBER(TO_CHAR(ondate, 'dd')) ondate, fieldcode, hectare_harvested" & _
       " FROM new_pusingan_tab" & _
       " WHERE fcba = '" & v_txtFcba & "'" & _
       "       AND TO_CHAR(ondate, 'yyyymm') = '" & v_txtYear & "'||LPAD('" & v_txtMonth & "', 2, 0)" & _
       "       AND ondate <= TO_DATE('" & v_txtDate & "', 'mm/dd/yyyy')"

' response.write qSQL
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open qSQL, connstr, adOpenStatic

arr_hectare2_eof = true
If Not qORS.EOF Then
	arr_hectare2 = qORS.getrows()
	arr_hectare2_eof = false
end if
qORS.Close
set qORS = Nothing

Function f_colorbyhectare(data_array__, day__, fieldcode__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if Cstr(arr_1(0, indx)) = Cstr(day__) and Cstr(arr_1(1, indx)) = Cstr(fieldcode__) then
			result_ = Cdbl(To_Nol(arr_1(2, indx)))
		end if
	next
	f_colorbyhectare = result_
End Function

Function f_colorsection(cond__)
	result_ = ""
	if cstr(cond__) = "A" then
		result_ = "background-color:#e26b0a;"
	elseif cstr(cond__) = "B" then
		result_ = "background-color:#92cddc;"
	elseif cstr(cond__) = "C" then
		result_ = "background-color:#76933c;"
	elseif cstr(cond__) = "D" then
		result_ = "background-color:#ffc000;"
	elseif cstr(cond__) = "E" then
		result_ = "background-color:#ff00ff;"
	elseif cstr(cond__) = "F" then
		result_ = "background-color:#66ff33;"
	else
		result_ = "background-color:#ffffff;"
	end if
	f_colorsection = result_
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
<title>ePlaS Reporting - Rotasi Panen</title>
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
						<!--<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM
						mon = Month(Session("v_Costbook"))
						If len(v_txtMonth)=0 then v_txtMonth=mon

						Yr = Year(Session("v_Costbook"))
						MinYr = Yr - 2
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
							<td><font face="verdana" size="1">BULAN</font></td>
							<td width="2">
								<select name="txtMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>-->
							<td>
								<font face="verdana" size="1">DATE</font>
							</td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDate" value="<%=v_txtDate%>" />
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
									
									If rsFcba.Bof Or rsFcba.eof Then
									%>
									<option value=""></option>
									<%
									End If
									
									if not rsFcba.eof then
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if v_txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
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
									<option value="ALL">--ALL--</option>
									<%
									Dim rsAfdeling, strAfdeling
									strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = v_txtAfdeling then%>
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
								<font face="verdana" size="1">BLOK</font>
							</td>
							<td>
								<select name="txtBlok" >
									<option value="ALL">--ALL--</option>
									<%
									Dim rsBlok, strBlok
									strBlok	=	" SELECT fccode, fcname FROM field WHERE fcba = '" & v_txtFcba & "'" & _
												" AND field_api.getvalid_isactive_harvested(fccode, fcba, TO_DATE('"&v_txtDate&"', 'mm/dd/yyyy')) = 'TRUE' ORDER BY fccode"
									
									Set rsBlok = Server.CreateObject("ADODB.Recordset")
									rsBlok.Open strBlok, connstr, adOpenStatic 
									
									if not rsBlok.eof then
									Do While Not rsBlok.eof
									if rsBlok("FCCODE") = v_txtBlok then%>
									<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option> 
									<%Else%>
									<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option>  
									<%End if
										rsBlok.movenext
									loop
									end if
									rsBlok.Close
									set rsBlok = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">COLOR</font>
							</td>
							<td>
								<select name="txtColor" >
									<option value="" <%If Len(Trim(v_txtColor)) = 0 Then%> Selected <%End If%>>--ALL--</option>
									<option value="RED" <%If v_txtColor = "RED" Then%> Selected <%End If%>>RED</option>
									<option value="YELLOW" <%If v_txtColor = "YELLOW" Then%> Selected <%End If%>>YELLOW</option>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">PERSENTASE DIATAS (HARI)</font>
							</td>
							<td width="2">
								<input type="text" name="txtPersen" value="<%=v_txtPersen%>" />
							</td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=6+GetLastDay(v_txtMonth, v_txtYear)%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>ROTASI PANEN</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				<h4>s/d Tgl&nbsp;:&nbsp;<%=v_txtDay%>&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;&nbsp;<%=v_txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
<% 
dim selec_day
	if len(trim(DAY(v_txtDate))) = 0 then
		selec_day = GetLastDay(v_txtMonth, v_txtYear)
	else
		selec_day = DAY(v_txtDate)
	end if
%>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">SEKSI</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">TT</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">LUAS</th>
		<th colspan="<%=selec_day%>" style="background-color:#00A300;color:#fff;">TANGGAL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">Jml Rotasi</th>
	</tr>
	<tr>
	
	
<%
Dim rsGetColorTanggal, strGetColorTanggal
Dim val_FirstDay, val_LastDay
Dim val_Background

val_FirstDay = v_txtMonth & "/" & 1 & "/" & v_txtYear
val_LastDay = v_txtMonth & "/" & v_txtDay & "/" & v_txtYear

strGetColorTanggal =	" SELECT" & _
						"     FCBA" & _
						"     , TO_NUMBER(TO_CHAR(TANGGAL, 'DD')) LIST_DAY" & _
						"     , DECODE(" & _
						"         BUDGET_TRANSACTION_WORKDAY_API.GETQTYWORKINGDAY_HI(FCBA, TO_NUMBER(TO_CHAR(TANGGAL, 'YYYY')), 1, 1,TANGGAL)" & _
						"         , 0,'LIBUR', 'NORMAL') KETERANGAN" & _
						" FROM (" & _
						"     SELECT" & _
						"         FCCODE FCBA" & _
						"         , TANGGAL" & _
						"     FROM (" & _
						"         SELECT TO_DATE('" & val_FirstDay & "' , 'MM/DD/YYYY') + ROWNUM - 1 TANGGAL" & _
						"         FROM ALL_OBJECTS" & _
						"         WHERE ROWNUM <= TO_DATE('" & val_LastDay & "','MM/DD/YYYY') - TO_DATE('" & val_FirstDay & "' , 'MM/DD/YYYY') + 1" & _
						"         )" & _
						"     , BUSINESSUNIT" & _
						"     ) WHERE FCBA = '" & v_txtFcba & "' ORDER BY TO_NUMBER(TO_CHAR(TANGGAL, 'DD'))"
						
'response.write(strGetColorTanggal)
Set rsGetColorTanggal = Server.CreateObject("ADODB.Recordset")
rsGetColorTanggal.Open strGetColorTanggal, connstr, adOpenStatic

If Not rsGetColorTanggal.EOF Then

	While Not rsGetColorTanggal.EOF
		If rsGetColorTanggal("KETERANGAN") = "LIBUR" Then
			val_Background = "background-color:#FF851B;color:#000;"
		Else
			val_Background = "background-color:#00A300;color:#fff;"
		End If
		
		v_txtMonth		=	month(v_txtDate)
v_txtDay		=	DAY(v_txtDate)
v_txtYear		=	year(v_txtDate)
%>
	<th style="<%=val_Background%>">
	<a target="_blank" href="http://172.30.1.122/ppb/Content/PPB_Map.aspx?typeModule=OLM&Company=<%=varCompanyCode%>&Estate=<%=v_txtFcba%>&Date=<%=Lpad(v_txtMonth,"0",2)%>/<%=Lpad(rsGetColorTanggal("LIST_DAY"),"0",2)%>/<%=v_txtYear%>">
	<%=rsGetColorTanggal("LIST_DAY")%>
	</a>
	</th>
<%
	
		rsGetColorTanggal.MoveNext
	Wend 
		
End If
%>
	</tr>
    </thead>
	<tbody>
<%
Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan, connstr, adOpenStatic

Dim xNum

Dim Total_Ha : Total_Ha = 0
Dim Total_LastInterval : Total_LastInterval = 0
Dim DataIsPersen : DataIsPersen = 0
Dim CountData : CountData = 0


if rsPusingan.bof and rsPusingan.eof then

%>
	<tr>
		<td colspan="<%=6+v_txtDay%>">NO DATA AVAILABLE</td>
	</tr>
<%end if

		
		while not rsPusingan.eof
			CountData = CountData + 1
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>

			<tr class="<%=cClass%>">
				<td align="center"><%=rsPusingan("DIVISION")%></td>
				<td align="center" style="<%=f_colorsection(rsPusingan("HARVESTINGSECTION"))%>"><%=rsPusingan("HARVESTINGSECTION")%></td>
				<td align="center" style="<%=f_colorsection(rsPusingan("HARVESTINGSECTION"))%>"><%=rsPusingan("FIELDCODE")%></td>
				<td align="center" style="<%=f_colorsection(rsPusingan("HARVESTINGSECTION"))%>"><%=rsPusingan("TAHUNTANAM")%></td>
				<td align="center" style="<%=f_colorsection(rsPusingan("HARVESTINGSECTION"))%>"><%=rsPusingan("HECTARAGEPLANTED")%></td>
				<%
				DIM pusingan_,Npusingan_,fontID
				Npusingan_ =0 			
				row_pusingan = 0
				row_temp=0	
				For xNum = 1 to selec_day %>
				<%
					'response.write row_temp & " atas <br>"
					If rsPusingan("color_"&xNum ) = "BLUE" Then
						ColorID = "#2ecc71"
						fontID = "#000"
						row_pusingan = f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE"))
					ElseIf rsPusingan("color_"&xNum ) = "YELLOW" Then
						yellowid = f_colorbyhectare(arr_hectare2, xNum, rsPusingan("FIELDCODE"))
						row_pusingan = f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE"))

						row_pusingan1 = f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE"))
						row_Ha = f_hectare2(arr_hectare, xNum, rsPusingan("FIELDCODE"))

						if "" & yellowid = "" or yellowid = 0 then		
							row_temp = cdbl(row_temp) + cdbl(f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE")))						
							ColorID = "#fff59d"
						else 
							row_temp = cdbl(row_temp) + cdbl(f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE")))	
							ColorID = "#ffdc00" 
						end if
						fontID = "#000000"
						row_temp =0
					ElseIf rsPusingan("color_"&xNum ) = "RED" Then 
						row_pusingan1 = f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE"))
						row_Ha = f_hectare2(arr_hectare, xNum, rsPusingan("FIELDCODE"))
						
						' response.write "<br> ha row_temp > " &row_temp
						' response.write "<br> ha row_Ha " & row_Ha

						If (cdbl(to_nol(row_temp)) + cdbl(f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE")))) > cdbl(to_nol(row_Ha)) Then
							ColorID = "#E74c3c"
							fontID = "#000000"
							row_pusingan = f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE"))
							' response.write "<br> Merah" 
						Else
							ColorID = "#09f"
							fontID = "#000000"
							row_pusingan = f_hectare(arr_hectare, xNum, rsPusingan("FIELDCODE"))							
							' response.write row_temp & "<br>"
						End If
						row_temp = 0
					Else
						ColorID = "#ffffff"
						fontID = "#000000"
						row_pusingan = rsPusingan("day_"&xNum)
					End If
				%>
				<td align="center" bgcolor="<%=ColorID%>" style="color:<%=fontID%>;"><%=row_pusingan%></td>
				<%
				Next
				
				if not arr_rotasi_eof then
					if Cdbl(To_nol(rsPusingan("HECTARAGEPLANTED"))) = 0 then
						Npusingan_ = 0
					else
						Npusingan_ = Round(f_rotasi(arr_rotasi, rsPusingan("FIELDCODE"))/ Cdbl(To_nol(rsPusingan("HECTARAGEPLANTED"))), 1)
					end if
				else
					Npusingan_ = 0
				end if
				%>
				<td width="50" align="center">
				<%=Npusingan_%>
				</td>
			</tr>

<%
			Total_Ha = Total_Ha + Cdbl(rsPusingan("HECTARAGEPLANTED"))
			Total_LastInterval = Total_LastInterval + (Cdbl(rsPusingan("lastinterval"))*Cdbl(rsPusingan("HECTARAGEPLANTED")))
			rsPusingan.MoveNext
		Wend
rsPusingan.Close
set rsPusingan = Nothing
'response.write strPusingan
'response.write  Total_LastInterval
'response.write  Total_Ha 
%>
			<tr>
				<td colspan="<%=5+v_txtDay%>" align="left"><strong>Rata-rata Pusingan pada tanggal <%=v_txtDay%>&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></strong></td>
				<%If Total_Ha > 0 Then%>
				<td align="Right"><strong><%=FbpNum2(Round(Total_LastInterval/Total_Ha, 2))%></strong></td>
				<%Else%>
				<td align="Right"><strong><%=FbpNum2(Round(0, 2))%></strong></td>
				<%End If%>
			</tr>
<%
If Len(Trim(v_txtPersen)) > 0 Then
	oSQL = " SELECT count(1) total" & _
	       " FROM harvesting_interval_tab" & _
		   " WHERE fcba = '" & v_txtFcba & "'" & _
		   "       AND ondate = to_date('" & v_txtDate & "', 'mm/dd/yyyy')" & _
		   "       AND field_api.GetValid_Activation(fieldcode,fcba,ondate)='Y' " & _
		   "       AND FIELD_API.GetValid_isactive_harvested(fieldcode,fcba,ondate) = 'TRUE'" & _
		   "       AND FIELD_API.GetValid_Harvestingsection(fieldcode,fcba,ondate) IS NOT NULL" & _
		   "       AND originalnumber > '" & v_txtPersen & "'"
		   
	Set oRS = Server.CreateObject("ADODB.Recordset")
	oRS.Open oSQL, connstr, adOpenStatic

	If Not oRS.EOF Then
		DataIsPersen = Cdbl(oRS("total"))
	Else
		DataIsPersen = 0
	End If
	
	oRS.Close
	set oRS = Nothing
Else
	DataIsPersen = 0
End If
%>
			<tr>
				<td colspan="<%=5+v_txtDay%>" align="left"><strong>Persentase Pusingan > <%=v_txtPersen%> hari</strong></td>
				<%If CountData > 0 Then%>
				<td align="Right"><strong><%=Round((DataIsPersen/CountData)*100, 0)%> %</strong></td>
				<%Else%>
				<td align="Right"><strong><%=Round(0, 0)%> %</strong></td>
				<%End If%>
			</tr>
</tbody>
</table>

   <table width="90%" class="data">
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
	   <%
		DIM rsttd,ttd
		
		ttd = "SELECT FCMANAGER,FCKTU FROM BusinessUnit WHERE FCCODE = '" & v_txtFcba & "'"
		
	    Set rsttd = Server.CreateObject("ADODB.Recordset")
		rsttd.Open ttd, connstr, adOpenStatic

		If Not rsttd.EOF Then
	   %>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCKTU")%></u></strong></td>
	   </tr>
	<%  End If
%> 
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
