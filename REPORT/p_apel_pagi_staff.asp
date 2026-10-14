<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtJobcode, v_txtAfdeling, v_txtTahunTanam, v_txtBlok, v_Isactive
Dim ScriptName
Dim rsmonitoringclosing, strmonitoringclosing, strMain, strmonitoringclosing2
Dim NO, cClass, nday, xNum, nSpanTotal

Dim var_colspan : var_colspan = 1
d_span = 36

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

function f_color (datefrom__, dateto__)
	result__ = "background-color:#f53333;"
	if "" & datefrom__ <> "" and "" & dateto__ <> "" then
		if DateDiff("d",Cdate(datefrom__),CDate(dateto__)) >= 0 and DateDiff("d",Cdate(datefrom__),CDate(dateto__)) <= 2 then
			result__ = "background-color:#92d050;"
		elseif DateDiff("d",Cdate(datefrom__),CDate(dateto__)) > 2 then
			result__ = "background-color:#ffc000;"
		else
			result__ = "background-color:#f53333;"
		end if
	end if
	f_color = result__
end function

function f_fcbaspan (arr__, where__)
	result__ = 0
	arr1 = arr__
	
	for xyy=0 to ubound(arr1,2)
		if cstr("" & arr1(1, xyy)) = cstr("" & where__) then
				result__ = result__ +1
		end if
	next
	
	f_fcbaspan = result__
end function

nSpanTotal = 9

v_Isactive = ""
v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_Isactive              = Request("txtIsactive")

If v_Isactive = "" then
   v_Isactive="TRUE"
end if




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
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
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
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
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
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
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
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
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

If Len(Trim(v_txtfcba)) = 0 Then v_txtfcba = "--SUMMARY--"

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " , fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
	Else
		selectCons	=	selectCons	& " , fcba"
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

If v_Isactive <> "ALL" then
   whereCons	=	whereCons	& " and businessunit_api.Isactive(fcba) = '" & v_Isactive & "'"	
end if

txtYear = Request("txtYear")
txtMonth = Request("txtMonth")
txttdate = Request("txttdate")
txtjabatan = Request("txtjabatan")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
If "" & txttdate = "" Then txttdate = textDate(date)

txttdate = txtMonth & "/01/" & txtYear


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain = " SELECT" & strFieldCons & _
          "        , fccode" & _
          "        , company_nik" & _
          "        , fcname" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 1, location, NULL)) day_1" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 2, location, NULL)) day_2" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 3, location, NULL)) day_3" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 4, location, NULL)) day_4" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 5, location, NULL)) day_5" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 6, location, NULL)) day_6" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 7, location, NULL)) day_7" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 8, location, NULL)) day_8" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 9, location, NULL)) day_9" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 10, location, NULL)) day_10" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 11, location, NULL)) day_11" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 12, location, NULL)) day_12" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 13, location, NULL)) day_13" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 14, location, NULL)) day_14" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 15, location, NULL)) day_15" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 16, location, NULL)) day_16" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 17, location, NULL)) day_17" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 18, location, NULL)) day_18" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 19, location, NULL)) day_19" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 20, location, NULL)) day_20" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 21, location, NULL)) day_21" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 22, location, NULL)) day_22" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 23, location, NULL)) day_23" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 24, location, NULL)) day_24" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 25, location, NULL)) day_25" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 26, location, NULL)) day_26" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 27, location, NULL)) day_27" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 28, location, NULL)) day_28" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 29, location, NULL)) day_29" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 30, location, NULL)) day_30" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 31, location, NULL)) day_31" & _
		  "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 1, fcba_loc, NULL)) xxday_1" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 2, fcba_loc, NULL)) xxday_2" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 3, fcba_loc, NULL)) xxday_3" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 4, fcba_loc, NULL)) xxday_4" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 5, fcba_loc, NULL)) xxday_5" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 6, fcba_loc, NULL)) xxday_6" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 7, fcba_loc, NULL)) xxday_7" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 8, fcba_loc, NULL)) xxday_8" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 9, fcba_loc, NULL)) xxday_9" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 10, fcba_loc, NULL)) xxday_10" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 11, fcba_loc, NULL)) xxday_11" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 12, fcba_loc, NULL)) xxday_12" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 13, fcba_loc, NULL)) xxday_13" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 14, fcba_loc, NULL)) xxday_14" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 15, fcba_loc, NULL)) xxday_15" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 16, fcba_loc, NULL)) xxday_16" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 17, fcba_loc, NULL)) xxday_17" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 18, fcba_loc, NULL)) xxday_18" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 19, fcba_loc, NULL)) xxday_19" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 20, fcba_loc, NULL)) xxday_20" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 21, fcba_loc, NULL)) xxday_21" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 22, fcba_loc, NULL)) xxday_22" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 23, fcba_loc, NULL)) xxday_23" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 24, fcba_loc, NULL)) xxday_24" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 25, fcba_loc, NULL)) xxday_25" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 26, fcba_loc, NULL)) xxday_26" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 27, fcba_loc, NULL)) xxday_27" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 28, fcba_loc, NULL)) xxday_28" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 29, fcba_loc, NULL)) xxday_29" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 30, fcba_loc, NULL)) xxday_30" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 31, fcba_loc, NULL)) xxday_31" & _
		  "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 1, attendance, NULL)) xday_1" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 2, attendance, NULL)) xday_2" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 3, attendance, NULL)) xday_3" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 4, attendance, NULL)) xday_4" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 5, attendance, NULL)) xday_5" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 6, attendance, NULL)) xday_6" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 7, attendance, NULL)) xday_7" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 8, attendance, NULL)) xday_8" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 9, attendance, NULL)) xday_9" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 10, attendance, NULL)) xday_10" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 11, attendance, NULL)) xday_11" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 12, attendance, NULL)) xday_12" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 13, attendance, NULL)) xday_13" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 14, attendance, NULL)) xday_14" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 15, attendance, NULL)) xday_15" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 16, attendance, NULL)) xday_16" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 17, attendance, NULL)) xday_17" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 18, attendance, NULL)) xday_18" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 19, attendance, NULL)) xday_19" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 20, attendance, NULL)) xday_20" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 21, attendance, NULL)) xday_21" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 22, attendance, NULL)) xday_22" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 23, attendance, NULL)) xday_23" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 24, attendance, NULL)) xday_24" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 25, attendance, NULL)) xday_25" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 26, attendance, NULL)) xday_26" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 27, attendance, NULL)) xday_27" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 28, attendance, NULL)) xday_28" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 29, attendance, NULL)) xday_29" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 30, attendance, NULL)) xday_30" & _
          "        , MAX(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 31, attendance, NULL)) xday_31" & _
          "        , positiondescription" & _
          "        , position" & _
          " FROM (" & _
          " SELECT em.fcba" & _
          "        , em.fccode" & _
          "        , em.company_nik" & _
          "        , em.fcname" & _
          "        , em.tdate" & _
          "        , pos.description positiondescription" & _
          "        , em.position" & _
          "        , CASE WHEN uf.timein IS NULL THEN ''" & _
          "               WHEN uf.timein BETWEEN TO_DATE(TO_CHAR(em.tdate, 'mm/dd/yyyy')||'00:00:01', 'mm/dd/yyyy hh24:mi:ss') AND TO_DATE(TO_CHAR(em.tdate, 'mm/dd/yyyy')||'06:00:00', 'mm/dd/yyyy hh24:mi:ss') THEN 'KJ'" & _
          "               ELSE '' END attendance " & _
		  "        , CASE WHEN uf.timein IS NULL THEN ''" & _
          "               WHEN uf.timein BETWEEN TO_DATE(TO_CHAR(em.tdate, 'mm/dd/yyyy')||'00:00:01', 'mm/dd/yyyy hh24:mi:ss') AND TO_DATE(TO_CHAR(em.tdate, 'mm/dd/yyyy')||'06:00:00', 'mm/dd/yyyy hh24:mi:ss') THEN mf.fcba" & _
          "               ELSE '' END fcba_loc " & _
		  "        , CASE WHEN uf.timein IS NULL THEN ''" & _
          "               WHEN uf.timein BETWEEN TO_DATE(TO_CHAR(em.tdate, 'mm/dd/yyyy')||'00:00:01', 'mm/dd/yyyy hh24:mi:ss') AND TO_DATE(TO_CHAR(em.tdate, 'mm/dd/yyyy')||'06:00:00', 'mm/dd/yyyy hh24:mi:ss') THEN mf.location" & _
          "               ELSE '' END location " & _
          " FROM (" & _
          "     SELECT em.fcba, em.fccode, em.company_nik, em.fcname, dd.tdate, em.position" & _
          "     FROM iplasprod.employee em," & _
          "          ( SELECT TO_DATE ('"&txttdate&"', 'MM/DD/YYYY') +ROWNUM -1 tdate" & _
          "            FROM ALL_OBJECTS" & _
          "            WHERE ROWNUM <= LAST_DAY(TO_DATE ('"&txttdate&"', 'MM/DD/YYYY')) - TO_DATE ('"&txttdate&"', 'MM/DD/YYYY') + 1) dd" & _
          "     WHERE em.presentgrade = 'STF'" & _
          "           AND em.company_nik IS NOT NULL" & _
          "           AND em.dateterminate IS NULL"
if ""&txtjabatan <> "" then strMain = strMain & " AND em.position = '"&txtjabatan&"'"
strMain = strMain & " ) em" & _
                    " LEFT JOIN userfinger.attendance_tab uf ON (em.fcba = uf.fcbatobe AND em.fccode = uf.empnotobe and em.tdate = uf.tdate)" & _
                    " LEFT JOIN empposition pos ON (pos.fcba = em.fcba AND pos.position = em.position)" & _
                    " LEFT JOIN businessunit bu ON em.fcba = bu.fccode" & _
                    " LEFT JOIN userfinger.checkinout ci ON ( ci.fingerprint_id = uf.empno AND TO_CHAR(ci.checktime, 'mm/dd/yyyy hh24:mi') =  TO_CHAR(uf.timein, 'mm/dd/yyyy hh24:mi') )" & _
                    " LEFT JOIN master_machine_fingerprint mf ON mf.serial_number = REGEXP_REPLACE(ci.machine_sn, '[^0-9]', '')" & _
                    " ) GROUP BY fccode, company_nik, fcname, positiondescription, position, fcba"
								
strmonitoringclosing2 = " , company_nik, fcname, position, positiondescription" & _
                        " , day_1, day_2, day_3, day_4, day_5, day_6, day_7, day_8, day_9, day_10" & _
                        " , day_11, day_12, day_13, day_14, day_15, day_16, day_17, day_18, day_19, day_20" & _
						" , day_21, day_22, day_23, day_24, day_25, day_26, day_27, day_28, day_29, day_30, day_31" & _
						" , xxday_1, xxday_2, xxday_3, xxday_4, xxday_5, xxday_6, xxday_7, xxday_8, xxday_9, xxday_10" & _
                        " , xxday_11, xxday_12, xxday_13, xxday_14, xxday_15, xxday_16, xxday_17, xxday_18, xxday_19, xxday_20" & _
						" , xxday_21, xxday_22, xxday_23, xxday_24, xxday_25, xxday_26, xxday_27, xxday_28, xxday_29, xxday_30, xxday_31" & _
						" , xday_1, xday_2, xday_3, xday_4, xday_5, xday_6, xday_7, xday_8, xday_9, xday_10" & _
                        " , xday_11, xday_12, xday_13, xday_14, xday_15, xday_16, xday_17, xday_18, xday_19, xday_20" & _
						" , xday_21, xday_22, xday_23, xday_24, xday_25, xday_26, xday_27, xday_28, xday_29, xday_30, xday_31"

strmonitoringclosing = " SELECT GroupPlantation" & selectCons & strmonitoringclosing2 &"  FROM (" & strMain & " )" & whereCons & _
                       " ORDER BY nourut"

'Response.write strmonitoringclosing

arr_field_eof = true
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strmonitoringclosing, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_field = qORS.getrows()
	arr_field_eof = false
End If
qORS.Close
set qORS = Nothing
%>
<html>

<head>
<title>ePlaS Reporting - Monitoring Closing</title>
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
				<!--#include virtual="/REPORT/FormSelectNewSum.html"-->
						<table id="form-table" class="centered-margin">
						<tr>
							<td>
								<font face="verdana" size="1">Jabatan</font>
							</td>
							<td>
								<select name="txtjabatan" >
									<option value="">--All--</option>
									<%
									strFcba	=	" SELECT position fccode, description fcname FROM empposition" & _
												" WHERE fcba = 'JJP2' AND fces_em = 'SF'"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic

									if not rsFcba.eof then
									Do While Not rsFcba.eof
										if txtjabatan = rsFcba("FCCODE") then
											judul = rsFcba("fcname")
										end if
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if txtjabatan = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("fcname")%></option>
									<%
										rsFcba.movenext
									loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>
							
						<td><font face="verdana" size="1">Tahun</font></td>
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
							<td><font face="verdana" size="1">Bulan</font></td>
							<td width="2">
								<select name="txtMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>

<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span +add_colspan%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>APEL PAGI <%=UCase(judul)%></h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th rowspan="2"><%=titleSite%></th>
		<th rowspan="2">NIK</th>
		<th rowspan="2">Nama</th>
		<th rowspan="2">Jabatan</th>
		<th colspan="31">Tanggal</th>
		<th style="min-width:40px;max-width:40px;" rowspan="2">Total</th>
	</tr>
	<tr>
		<%for xi=1 to 31%>
		<th style="min-width:20px;max-width:20px;"><%=xi%></th>
		<%next%>
	</tr>

    </thead>
	<tbody>
<%
if not arr_field_eof then
	currfcba = ""
	prevfcba = ""
	for y=0 to ubound(arr_field,2)
		ntotal = 0
		
		response.write "<tr>"
		currfcba = arr_field(1,y)
		if currfcba <> prevfcba then			
			response.write "	<td rowspan='"&f_fcbaspan(arr_field, arr_field(1,y))&"'align='center' valign='top'>"&arr_field(1,y)&"</td>"
		end if
		response.write "	<td align='left'>"&arr_field(2,y)&"</td>"
		response.write "	<td align='left'>"&arr_field(3,y)&"</td>"
		response.write "	<td align='left'>"&arr_field(5,y)&"</td>"
		for xi=6 to 36
			response.write "	<td align='center'>"
			if ""&arr_field(4,y) = "GEM" and ""&arr_field(xi +31,y) <> "" then response.write arr_field(xi +31,y) & "."
			response.write 			arr_field(xi,y)
			if cstr(""&arr_field(xi,y)) <> "" then ntotal = ntotal +1
			response.write "	</td>"
		next
		response.write "	<td align='right'>"&ntotal&"</td>"
		response.write "</tr>"
		
		prevfcba = currfcba
	next
else
	response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
end if
%>
	</tbody>
</table>

   </body>
   </html>