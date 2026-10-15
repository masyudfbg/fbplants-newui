<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtCustomer,v_txtOwnstate,v_txtDivision
Dim ScriptName
Dim rstbsluar, strtbsluar, strMain, strtbsluar2,no
Dim cClass, nday, xNum

Dim var_ColSpan : var_ColSpan = 0

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType,v_txtCUSTFFBTYPE,conCUSTFFBTYPE,conCUSTFFBTYPE2

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtCUSTFFBTYPE		= Request("CUSTFFBTYPE")
v_txtOwnstate			= Request("txtOwnstate")
v_txtDivision			= Request("txtDivision")


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

v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

v_txtCustomer	=	Request("txtCustomer")

If Len(Trim(v_txtCustomer)) = 0 Then
	v_txtCustomer	=	"--ALL--"
End If

If Len(Trim(v_txtCustomer)) > 0 Then
	If v_txtCustomer = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtCustomer = "--ALL--" then
		selectCons	=	selectCons	& " , CUSTOMERCODE , CUSTOMERNAME"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , CUSTOMERCODE"
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , CUSTOMERCODE , CUSTOMERNAME" 
		whereCons	=	wherecons	& " And CUSTOMERCODE = '" & v_txtCustomer & "'"
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
	End if
End If

v_txtDivision	=	Request("txtDivision")

If Len(Trim(v_txtDivision)) = 0 Then
	v_txtDivision	=	"--SUMMARY--"
End If



	If v_txtDivision = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtDivision = "--ALL--" then
		selectCons	=	selectCons	& " , division"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , division"
		var_ColSpan	=	var_ColSpan	+ 1
	
	End if

If Len(Trim(v_txtCUSTFFBTYPE)) > 0 Then
	
	if v_txtCUSTFFBTYPE = "TUMPANGOLAH" then
		selectCons	=	selectCons	& " , CUSTFFBTYPE, Own_State" 
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
		conCUSTFFBTYPE = " AND customer_api.GetCUSTFFBTYPE(cust_customercode, fcba) = '" & v_txtCUSTFFBTYPE & "'"
		conCUSTFFBTYPE2 = " AND customer_api.GetCUSTFFBTYPE(B.cust_customercode, A.fcba) = '" & v_txtCUSTFFBTYPE & "'"
		
		if v_txtOwnstate = "Y" then 
		conCUSTFFBTYPE = conCUSTFFBTYPE & " AND customer_api.GetInternalCustomer(cust_customercode, fcba) = '" & v_txtOwnstate & "'" 
		conCUSTFFBTYPE2 = conCUSTFFBTYPE2 & " AND customer_api.GetInternalCustomer(B.cust_customercode, A.fcba) = '" & v_txtOwnstate & "'" 
		else 
		conCUSTFFBTYPE = conCUSTFFBTYPE & " AND customer_api.GetInternalCustomer(cust_customercode, fcba) <> 'Y'"
		conCUSTFFBTYPE2 = conCUSTFFBTYPE2 & " AND customer_api.GetInternalCustomer(B.cust_customercode, A.fcba) <> 'Y'"
		end if

		
		
	Else 
		selectCons	=	selectCons	& " , CUSTFFBTYPE"
		whereCons	=	whereCons	& " "
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
		conCUSTFFBTYPE = " AND customer_api.GetCUSTFFBTYPE(cust_customercode, fcba) = '" & v_txtCUSTFFBTYPE & "'"
		conCUSTFFBTYPE2 = " AND customer_api.GetCUSTFFBTYPE(B.cust_customercode, A.fcba) = '" & v_txtCUSTFFBTYPE & "'"
	End If
else
	if v_txtCUSTFFBTYPE = "" then
		selectCons	=	selectCons	& " , CUSTFFBTYPE"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
		conCUSTFFBTYPE = ""
		conCUSTFFBTYPE2 = ""
	End If
End If
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" SELECT" & strFieldCons & _  
				"     , CUSTOMERCODE" & _
				"     , division" & _
				"     , CUSTFFBTYPE" & _
				"     , Own_State" & _
				"     , CUSTOMER_API.GET_DESCRIPTION(CUSTOMERCODE, FCBA) CUSTOMERNAME" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 1, KG, 0)) KG_1" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 2, KG, 0)) KG_2" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 3, KG, 0)) KG_3" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 4, KG, 0)) KG_4" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 5, KG, 0)) KG_5" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 6, KG, 0)) KG_6" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 7, KG, 0)) KG_7" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 8, KG, 0)) KG_8" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 9, KG, 0)) KG_9" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 10, KG, 0)) KG_10" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 11, KG, 0)) KG_11" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 12, KG, 0)) KG_12" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 13, KG, 0)) KG_13" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 14, KG, 0)) KG_14" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 15, KG, 0)) KG_15" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 16, KG, 0)) KG_16" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 17, KG, 0)) KG_17" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 18, KG, 0)) KG_18" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 19, KG, 0)) KG_19" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 20, KG, 0)) KG_20" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 21, KG, 0)) KG_21" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 22, KG, 0)) KG_22" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 23, KG, 0)) KG_23" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 24, KG, 0)) KG_24" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 25, KG, 0)) KG_25" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 26, KG, 0)) KG_26" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 27, KG, 0)) KG_27" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 28, KG, 0)) KG_28" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 29, KG, 0)) KG_29" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 30, KG, 0)) KG_30" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 31, KG, 0)) KG_31" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'YYYY')), " & v_txtYear & ", KG, TO_NUMBER (TO_CHAR (ONDATE, 'MM')), " & v_txtMonth & ", KG, 0)) KG_TOTAL" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 1, KG_SORTAGE, 0)) KG_SORTAGE_1" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 2, KG_SORTAGE, 0)) KG_SORTAGE_2" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 3, KG_SORTAGE, 0)) KG_SORTAGE_3" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 4, KG_SORTAGE, 0)) KG_SORTAGE_4" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 5, KG_SORTAGE, 0)) KG_SORTAGE_5" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 6, KG_SORTAGE, 0)) KG_SORTAGE_6" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 7, KG_SORTAGE, 0)) KG_SORTAGE_7" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 8, KG_SORTAGE, 0)) KG_SORTAGE_8" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 9, KG_SORTAGE, 0)) KG_SORTAGE_9" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 10, KG_SORTAGE, 0)) KG_SORTAGE_10" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 11, KG_SORTAGE, 0)) KG_SORTAGE_11" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 12, KG_SORTAGE, 0)) KG_SORTAGE_12" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 13, KG_SORTAGE, 0)) KG_SORTAGE_13" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 14, KG_SORTAGE, 0)) KG_SORTAGE_14" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 15, KG_SORTAGE, 0)) KG_SORTAGE_15" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 16, KG_SORTAGE, 0)) KG_SORTAGE_16" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 17, KG_SORTAGE, 0)) KG_SORTAGE_17" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 18, KG_SORTAGE, 0)) KG_SORTAGE_18" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 19, KG_SORTAGE, 0)) KG_SORTAGE_19" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 20, KG_SORTAGE, 0)) KG_SORTAGE_20" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 21, KG_SORTAGE, 0)) KG_SORTAGE_21" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 22, KG_SORTAGE, 0)) KG_SORTAGE_22" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 23, KG_SORTAGE, 0)) KG_SORTAGE_23" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 24, KG_SORTAGE, 0)) KG_SORTAGE_24" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 25, KG_SORTAGE, 0)) KG_SORTAGE_25" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 26, KG_SORTAGE, 0)) KG_SORTAGE_26" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 27, KG_SORTAGE, 0)) KG_SORTAGE_27" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 28, KG_SORTAGE, 0)) KG_SORTAGE_28" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 29, KG_SORTAGE, 0)) KG_SORTAGE_29" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 30, KG_SORTAGE, 0)) KG_SORTAGE_30" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'DD')), 31, KG_SORTAGE, 0)) KG_SORTAGE_31" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (ONDATE, 'YYYY')), " & v_txtYear & ", KG_SORTAGE, TO_NUMBER (TO_CHAR (ONDATE, 'MM')), " & v_txtMonth & ", KG_SORTAGE, 0)) KG_SORTAGE_TOTAL" & _
				" FROM (" & _
				"     SELECT" & _
				"         FCBA" & _
				"         , CUST_CUSTOMERCODE CUSTOMERCODE" & _
				"         , division" & _ 
				"		  , CUSTFFBTYPE" & _
				"         , Own_State" & _
				"         , ONDATE" & _
				"         , SUM(KG) KG" & _
				"         , SUM(KG_SORTAGE) KG_SORTAGE" & _
				"     FROM (" & _
				"         SELECT" & _
				"             FCBA" & _
				"             , CUST_CUSTOMERCODE" & _ 
				"             , division" & _ 
				"             , customer_api.GetCUSTFFBTYPE(cust_customercode, fcba) CUSTFFBTYPE" & _
				"             , customer_api.GetInternalCustomer(cust_customercode, fcba) Own_State" & _
				"             , POSTINGDATE ONDATE" & _
				"             , SUM (NVL (WEIGHTIN, 0) - NVL (WEIGHTOUT, 0)) KG" & _
				"             , 0 KG_SORTAGE" & _
				"         FROM" & _
				"             WEIGHBRIDGETICKET" & _
				"         WHERE" & _
				"             TRANSACTIONTYPE = '1'" & _
				"             AND PRODUCT_PRODUCTCODE = 'TBS'" & conCUSTFFBTYPE & _
				"             AND FCBA||CUST_CUSTOMERCODE IN (SELECT T.FCBA||T.CUSTOMERCODE FROM CUSTOMER T WHERE T.SUPPLIER = 'Y')" & _
				"             AND TO_NUMBER (TO_CHAR (POSTINGDATE, 'YYYY')) = " & v_txtYear & "" & _
				"             AND TO_NUMBER (TO_CHAR (POSTINGDATE, 'MM')) = " & v_txtMonth & "" & _
				"         GROUP BY" & _
				"             FCBA" & _
				"             , CUST_CUSTOMERCODE,division" & _
				"             , POSTINGDATE" & _
				"         UNION" & _
				"         SELECT" & _
				"             A.FCBA" & _
				"             , B.CUST_CUSTOMERCODE" & _
				"             , B.division" & _ 
				"             , customer_api.GetCUSTFFBTYPE(B.cust_customercode, A.fcba) CUSTFFBTYPE" & _
				"             , customer_api.GetInternalCustomer(B.cust_customercode, A.fcba) Own_State" & _
				"             , A.POSTINGDATE ONDATE" & _
				"             , 0 KG" & _
				"             , SUM (WEIGHT) KG_SORTAGE" & _
				"         FROM" & _
				"             GRADINGRECORD A" & _
				"             , WEIGHBRIDGETICKET B" & _
				"         WHERE" & _
				"             B.CHITNO = A.WBT_CHITNO" & _
				"             AND B.FCBA = A.FCBA" & conCUSTFFBTYPE2 & _
				"             AND TO_NUMBER (TO_CHAR (A.POSTINGDATE, 'YYYY')) = " & v_txtYear & "" & _
				"             AND TO_NUMBER (TO_CHAR (A.POSTINGDATE, 'MM')) = " & v_txtMonth & "" & _
				"         GROUP BY" & _
				"         A.FCBA" & _
				"         , B.CUST_CUSTOMERCODE,B.division" & _
				"         , A.POSTINGDATE" & _
				"         )" & _
				"     GROUP BY" & _
				"         FCBA" & _
				"         , CUST_CUSTOMERCODE,division" & _
				"         , ONDATE" & _
				"         , CUSTFFBTYPE" & _
				"         , Own_State" & _
				"     ) WHERE USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) = 'M'" & _
				" GROUP BY" & _
				"     FCBA" & _
				"     , CUSTOMERCODE,division" & _ 
				"     , CUSTFFBTYPE" & _
				"     , Own_State" & _
				" ORDER BY" & _
				"     FCBA" & _
				"     , CUSTOMERCODE"								
								
strtbsluar2	=	"     , SUM(KG_1) KG_1" & _
				"     , SUM(KG_2) KG_2" & _
				"     , SUM(KG_3) KG_3" & _
				"     , SUM(KG_4) KG_4" & _
				"     , SUM(KG_5) KG_5" & _
				"     , SUM(KG_6) KG_6" & _
				"     , SUM(KG_7) KG_7" & _
				"     , SUM(KG_8) KG_8" & _
				"     , SUM(KG_9) KG_9" & _
				"     , SUM(KG_10) KG_10" & _
				"     , SUM(KG_11) KG_11" & _
				"     , SUM(KG_12) KG_12" & _
				"     , SUM(KG_13) KG_13" & _
				"     , SUM(KG_14) KG_14" & _
				"     , SUM(KG_15) KG_15" & _
				"     , SUM(KG_16) KG_16" & _
				"     , SUM(KG_17) KG_17" & _
				"     , SUM(KG_18) KG_18" & _
				"     , SUM(KG_19) KG_19" & _
				"     , SUM(KG_20) KG_20" & _
				"     , SUM(KG_21) KG_21" & _
				"     , SUM(KG_22) KG_22" & _
				"     , SUM(KG_23) KG_23" & _
				"     , SUM(KG_24) KG_24" & _
				"     , SUM(KG_25) KG_25" & _
				"     , SUM(KG_26) KG_26" & _
				"     , SUM(KG_27) KG_27" & _
				"     , SUM(KG_28) KG_28" & _
				"     , SUM(KG_29) KG_29" & _
				"     , SUM(KG_30) KG_30" & _
				"     , SUM(KG_31) KG_31" & _
				"     , SUM(KG_TOTAL) KG_TOTAL" & _
				"     , SUM(KG_SORTAGE_1) KG_SORTAGE_1" & _
				"     , SUM(KG_SORTAGE_2) KG_SORTAGE_2" & _
				"     , SUM(KG_SORTAGE_3) KG_SORTAGE_3" & _
				"     , SUM(KG_SORTAGE_4) KG_SORTAGE_4" & _
				"     , SUM(KG_SORTAGE_5) KG_SORTAGE_5" & _
				"     , SUM(KG_SORTAGE_6) KG_SORTAGE_6" & _
				"     , SUM(KG_SORTAGE_7) KG_SORTAGE_7" & _
				"     , SUM(KG_SORTAGE_8) KG_SORTAGE_8" & _
				"     , SUM(KG_SORTAGE_9) KG_SORTAGE_9" & _
				"     , SUM(KG_SORTAGE_10) KG_SORTAGE_10" & _
				"     , SUM(KG_SORTAGE_11) KG_SORTAGE_11" & _
				"     , SUM(KG_SORTAGE_12) KG_SORTAGE_12" & _
				"     , SUM(KG_SORTAGE_13) KG_SORTAGE_13" & _
				"     , SUM(KG_SORTAGE_14) KG_SORTAGE_14" & _
				"     , SUM(KG_SORTAGE_15) KG_SORTAGE_15" & _
				"     , SUM(KG_SORTAGE_16) KG_SORTAGE_16" & _
				"     , SUM(KG_SORTAGE_17) KG_SORTAGE_17" & _
				"     , SUM(KG_SORTAGE_18) KG_SORTAGE_18" & _
				"     , SUM(KG_SORTAGE_19) KG_SORTAGE_19" & _
				"     , SUM(KG_SORTAGE_20) KG_SORTAGE_20" & _
				"     , SUM(KG_SORTAGE_21) KG_SORTAGE_21" & _
				"     , SUM(KG_SORTAGE_22) KG_SORTAGE_22" & _
				"     , SUM(KG_SORTAGE_23) KG_SORTAGE_23" & _
				"     , SUM(KG_SORTAGE_24) KG_SORTAGE_24" & _
				"     , SUM(KG_SORTAGE_25) KG_SORTAGE_25" & _
				"     , SUM(KG_SORTAGE_26) KG_SORTAGE_26" & _
				"     , SUM(KG_SORTAGE_27) KG_SORTAGE_27" & _
				"     , SUM(KG_SORTAGE_28) KG_SORTAGE_28" & _
				"     , SUM(KG_SORTAGE_29) KG_SORTAGE_29" & _
				"     , SUM(KG_SORTAGE_30) KG_SORTAGE_30" & _
				"     , SUM(KG_SORTAGE_31) KG_SORTAGE_31" & _
				"     , SUM(KG_SORTAGE_TOTAL) KG_SORTAGE_TOTAL"

strtbsluar	=	" select GroupPlantation " & selectCons & strtbsluar2 & " From (" & strMain & ")" & _
				whereCons & _
				" group by GroupPlantation" & selectCons & " " & _
				" Order by GroupPlantation" & orderCons 

' response.write strtbsluar

%>
<html>

<head>
<title>ePlaS Reporting - Penerimaan TBS Luar</title>
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
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSumM.html"-->
				
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TAHUN</font>
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
						<td width="4">
							<select name="txtYear">
							<%For i = MinYr to MaxYr%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">BULAN</font>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">CUSTOMER</font>
						</td>
						<td>
							<select name="txtCustomer" >
								<option value="--ALL--"<%If v_txtCustomer = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsCustomer, strCustomer
								strCustomer	=	" SELECT DISTINCT CUSTOMERCODE, DESCRIPTION FROM CUSTOMER" & _
												" WHERE SUPPLIER = 'Y' AND FCBA = '" & v_txtfcba & "'" & _
											" ORDER BY CUSTOMERCODE"
								
								Set rsCustomer = Server.CreateObject("ADODB.Recordset")
								rsCustomer.Open strCustomer, connstr, adOpenStatic 
								
								if not rsCustomer.eof then
								Do While Not rsCustomer.eof
								if rsCustomer("CUSTOMERCODE") = v_txtCustomer then%>
								<option value="<%=rsCustomer("CUSTOMERCODE")%>" Selected><%=rsCustomer("CUSTOMERCODE")%> : <%=rsCustomer("DESCRIPTION")%></option> 
								<%Else%>
								<option value="<%=rsCustomer("CUSTOMERCODE")%>"><%=rsCustomer("CUSTOMERCODE")%> : <%=rsCustomer("DESCRIPTION")%></option>  
								<%End if
									rsCustomer.movenext
								loop
								end if
								rsCustomer.Close
								set rsCustomer = Nothing
								%>
							</select>
						</td>
					
						<td>
							<font face="verdana" size="1">DIVISION</font>
						</td>
						<td>
						
							<select name="txtDivision">
								<option VALUE="--SUMMARY--" <% if v_txtDivision = "--SUMMARY--" then %>SELECTED<% end if %>>--SUMMARY--</option>
								<option VALUE="--ALL--" <% if v_txtDivision = "--ALL--" then %>SELECTED<% end if %>>--ALL--</option>
							</select>
					

							
							
						</td>	
					
						<td><font face="verdana" size="1">FFB Type</font></td>
						<td>
						
						<script>
						function tree_olah(){
							  var TumpangOlah = document.getElementById("CUSTFFBTYPE");
							  var SUN_tumpan_olah = document.getElementById("SUN_tumpan_olah");
							  var v_txtTumpangOlah = TumpangOlah.options[TumpangOlah.selectedIndex].value
								if (v_txtTumpangOlah == "TUMPANGOLAH") {
								
									SUN_tumpan_olah.innerHTML="<select name='txtOwnstate'><option VALUE='Y'>TBS Afiliasi</option><option VALUE='N'>TBS Luar</option></select>";
									
								}else{
									
									SUN_tumpan_olah.innerHTML="";
								
								}
							}
							
							
						</script>
						<select id="CUSTFFBTYPE" name="CUSTFFBTYPE" onchange="tree_olah()">
							<option VALUE="" ></option>
							<% 
							DIM cffbT,rscffbT,SELECTED
							cffbT="select FCCODE, FCNAME from CUSTOMERFFBTYPE"
							Set rscffbT = Server.CreateObject("ADODB.Recordset")
							'response.Write(cffbT)
							'response.Flush()
							rscffbT.CursorType = 3
							rscffbT.CursorLocation = 3
							rscffbT.Open cffbT, connstr 
							if Not rscffbT.eof then
							  While Not rscffbT.EOF
							  IF v_txtCUSTFFBTYPE = rscffbT("FCCODE") THEN
							  SELECTED = "SELECTED"
							  ELSE
							  SELECTED = ""
							  END IF
							%>
								<option VALUE="<%=rscffbT("FCCODE")%>" <%=SELECTED%>><%=rscffbT("FCNAME")%></option>
						<%
							rscffbT.MoveNext
							  Wend
							end if 
							rscffbT.Close
							set rscffbT = Nothing
						
						%>
							</select>
						</td>
						<td id="SUN_tumpan_olah">
						<% if v_txtCUSTFFBTYPE = "TUMPANGOLAH" then%>
							<select name="txtOwnstate">
								<option VALUE="Y" <% if v_txtOwnstate = "Y" then %>SELECTED<% end if %>>TBS Afiliasi</option>
								<option VALUE="N" <% if v_txtOwnstate = "N" then %>SELECTED<% end if %>>TBS Luar</option>
							</select>
						<% end if %>

							
							
						</td>	
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>


<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
		<tr>
			<th colspan="<%=var_ColSpan+97%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					<h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>LAPORAN HARIAN PENERIMAAN TBS</h3>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
			</th>
		</tr>
	<tr>

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
		
		<%If Len(Trim(v_txtCustomer)) > 0 And v_txtCustomer <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;">SUPPLIER/KEBUN</th>
		<%End If%>
		
		<% IF v_txtDivision <> "--SUMMARY--"  then %>
			<th rowspan="3" style="background-color:#00A300;color:#fff;">DIVISION</th>
		<% END IF %>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">FFB Type</th>

		<th colspan="93" style="background-color:#00A300;color:#fff;">TANGGAL</th>
		<th rowspan="2" colspan="3" style="background-color:#00A300;color:#fff;">TOTAL</th>
	</tr>
	<tr>
		<%for nday = 1 to 31%>
		<th colspan="3" style="background-color:#00A300;color:#fff;"><%=nday%></th>
		<%next%>
	</tr>
	<tr>
		<%for nday = 1 to 32%>
		<th style="background-color:#00A300;color:#fff;">KG</th>
		<th style="background-color:#00A300;color:#fff;">SORTASI</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<%next%>
	</tr>
    </thead>
	<tbody>
<%
'response.write strtbsluar
Set rstbsluar = Server.CreateObject("ADODB.Recordset")
rstbsluar.Open strtbsluar, connstr, adOpenStatic

Dim GrandKg_1 : GrandKg_1 = 0
Dim GrandKg_2 : GrandKg_2 = 0
Dim GrandKg_3 : GrandKg_3 = 0
Dim GrandKg_4 : GrandKg_4 = 0
Dim GrandKg_5 : GrandKg_5 = 0
Dim GrandKg_6 : GrandKg_6 = 0
Dim GrandKg_7 : GrandKg_7 = 0
Dim GrandKg_8 : GrandKg_8 = 0
Dim GrandKg_9 : GrandKg_9 = 0
Dim GrandKg_10 : GrandKg_10 = 0
Dim GrandKg_11 : GrandKg_11 = 0
Dim GrandKg_12 : GrandKg_12 = 0
Dim GrandKg_13 : GrandKg_13 = 0
Dim GrandKg_14 : GrandKg_14 = 0
Dim GrandKg_15 : GrandKg_15 = 0
Dim GrandKg_16 : GrandKg_16 = 0
Dim GrandKg_17 : GrandKg_17 = 0
Dim GrandKg_18 : GrandKg_18 = 0
Dim GrandKg_19 : GrandKg_19 = 0
Dim GrandKg_20 : GrandKg_20 = 0
Dim GrandKg_21 : GrandKg_21 = 0
Dim GrandKg_22 : GrandKg_22 = 0
Dim GrandKg_23 : GrandKg_23 = 0
Dim GrandKg_24 : GrandKg_24 = 0
Dim GrandKg_25 : GrandKg_25 = 0
Dim GrandKg_26 : GrandKg_26 = 0
Dim GrandKg_27 : GrandKg_27 = 0
Dim GrandKg_28 : GrandKg_28 = 0
Dim GrandKg_29 : GrandKg_29 = 0
Dim GrandKg_30 : GrandKg_30 = 0
Dim GrandKg_31 : GrandKg_31 = 0
Dim GrandKg_Total : GrandKg_Total = 0

Dim GrandKgSortasi_1 : GrandKgSortasi_1 = 0
Dim GrandKgSortasi_2 : GrandKgSortasi_2 = 0
Dim GrandKgSortasi_3 : GrandKgSortasi_3 = 0
Dim GrandKgSortasi_4 : GrandKgSortasi_4 = 0
Dim GrandKgSortasi_5 : GrandKgSortasi_5 = 0
Dim GrandKgSortasi_6 : GrandKgSortasi_6 = 0
Dim GrandKgSortasi_7 : GrandKgSortasi_7 = 0
Dim GrandKgSortasi_8 : GrandKgSortasi_8 = 0
Dim GrandKgSortasi_9 : GrandKgSortasi_9 = 0
Dim GrandKgSortasi_10 : GrandKgSortasi_10 = 0
Dim GrandKgSortasi_11 : GrandKgSortasi_11 = 0
Dim GrandKgSortasi_12 : GrandKgSortasi_12 = 0
Dim GrandKgSortasi_13 : GrandKgSortasi_13 = 0
Dim GrandKgSortasi_14 : GrandKgSortasi_14 = 0
Dim GrandKgSortasi_15 : GrandKgSortasi_15 = 0
Dim GrandKgSortasi_16 : GrandKgSortasi_16 = 0
Dim GrandKgSortasi_17 : GrandKgSortasi_17 = 0
Dim GrandKgSortasi_18 : GrandKgSortasi_18 = 0
Dim GrandKgSortasi_19 : GrandKgSortasi_19 = 0
Dim GrandKgSortasi_20 : GrandKgSortasi_20 = 0
Dim GrandKgSortasi_21 : GrandKgSortasi_21 = 0
Dim GrandKgSortasi_22 : GrandKgSortasi_22 = 0
Dim GrandKgSortasi_23 : GrandKgSortasi_23 = 0
Dim GrandKgSortasi_24 : GrandKgSortasi_24 = 0
Dim GrandKgSortasi_25 : GrandKgSortasi_25 = 0
Dim GrandKgSortasi_26 : GrandKgSortasi_26 = 0
Dim GrandKgSortasi_27 : GrandKgSortasi_27 = 0
Dim GrandKgSortasi_28 : GrandKgSortasi_28 = 0
Dim GrandKgSortasi_29 : GrandKgSortasi_29 = 0
Dim GrandKgSortasi_30 : GrandKgSortasi_30 = 0
Dim GrandKgSortasi_31 : GrandKgSortasi_31 = 0
Dim GrandKgSortasi_Total : GrandKgSortasi_Total = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempCustomer : tempCustomer = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim sPrevCustomer : sPrevCustomer = ""

if rstbsluar.bof and rstbsluar.eof then
%>
	<tr>
		<td colspan="<%=var_ColSpan+97%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rstbsluar.eof
%>
				<tr>

				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rstbsluar("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rstbsluar("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstbsluar("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   tempPlantation = rstbsluar("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rstbsluar("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstbsluar("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rstbsluar("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rstbsluar("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rstbsluar("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstbsluar("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rstbsluar("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rstbsluar("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rstbsluar("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rstbsluar("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstbsluar("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rstbsluar("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rstbsluar("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rstbsluar("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rstbsluar("Company")
					   tempfcba = "--ALL--"			   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rstbsluar("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rstbsluar("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rstbsluar("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rstbsluar("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rstbsluar("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rstbsluar("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rstbsluar("fcba")
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
						    <%=rstbsluar("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtCustomer <>  "--SUMMARY--" and  v_txtCustomer <>  "" then %>
					<%	tempCustomer = rstbsluar("CUSTOMERCODE")
						If Cint(visibleP) > 1 Then
							tempGroupPlantation = "--ALL--"
						Else
							If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then
								tempGroupPlantation = rstbsluar("GroupPlantation")
							Else
								tempGroupPlantation = ""
							End If
						End If
						
						If Cint(visibleGR) > 1 And Cint(visibleP) > 1 Then
							tempPlantation = "--SUMMARY--"
						Else
							If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then
								tempPlantation = rstbsluar("Plantation")
							Else
								tempPlantation = ""
							End If
						End If
						
						If Cint(visibleR) > 1 And Cint(visibleGR) > 1 Then
							tempGroupRegion = "--SUMMARY--"
						Else
							If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then
								tempGroupRegion = rstbsluar("GroupRegion")
							Else
								tempGroupRegion = ""
							End If
						End If
						
						If Cint(visibleC) > 1 And Cint(visibleR) > 1 Then
							tempRegion = "--SUMMARY--"
						Else
							If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then
								tempRegion = rstbsluar("Region")
							Else
								tempRegion = ""
							End If
						End If
						
						If Cint(visibleS) > 1 And Cint(visibleC) > 1 Then
							tempCompany = "--SUMMARY--"
						Else
							If Len(Trim(v_txtcompany)) > 0 And v_txtcompany <>  "--SUMMARY--" Then
								tempCompany = rstbsluar("Company")
							Else
								tempCompany = ""
							End If
						End If
						
						If Cint(VisibleS) > 1 Then
							tempfcba = "--SUMMARY--"
						Else
							If Len(Trim(v_txtFcba)) > 0 And v_txtFcba <>  "--SUMMARY--" Then
								tempfcba = rstbsluar("Fcba")
							Else
								tempfcba = ""
							End If
						End If
					%>
						<% if tempCustomer <> sPrevCustomer then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
								<%=rstbsluar("CUSTOMERNAME")%>
								</a>
							</td>
							<% sPrevCustomer = tempCustomer %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
				<% end if %>
				<% IF v_txtDivision <> "--SUMMARY--"  then %>
				   <td align="center"><%=rstbsluar("DIVISION")%></td>
				<% END IF %>
				<td align="center"><%=rstbsluar("CUSTFFBTYPE")%></td>
		
				<%For nday = 1 to 31%>
					<td align="right"><a href="../p_WEIGHBRIDGETICKET.asp?CUST_CUSTOMERCODE=<%=tempCustomer%>&Posting_Date=<%=v_txtMonth%>/<%=nday%>/<%=v_txtYear%>&PRODUCT_PRODUCTCODE=TBS" target="_blank"><%=FormatNumber(rstbsluar("KG_" & nday & ""), 2)%></a></td>
					<td align="right"><%=FormatNumber(rstbsluar("KG_SORTAGE_" & nday & ""), 2)%></td>
					<% If Cdbl(rstbsluar("KG_" & nday & "")) = 0 Then%>
						<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
						<td align="right"><%=FormatNumber(Cdbl(rstbsluar("KG_SORTAGE_" & nday & ""))/Cdbl(rstbsluar("KG_" & nday & ""))*100, 2)%></td>
					<%End If%>
				<%Next%>
				
					<td align="right"><%=FormatNumber(rstbsluar("KG_TOTAL"), 2)%></td>
					<td align="right"><%=FormatNumber(rstbsluar("KG_SORTAGE_TOTAL"), 2)%></td>
					<% If Cdbl(rstbsluar("KG_TOTAL")) = 0 Then%>
						<td align="right"><%=FormatNumber(0, 2)%></td>
					<%Else%>
						<td align="right"><%=FormatNumber(Cdbl(rstbsluar("KG_SORTAGE_TOTAL"))/Cdbl(rstbsluar("KG_TOTAL"))*100, 2)%></td>
					<%End If%>
				</tr>
<%
			GrandKg_1 =  GrandKg_1 + Cdbl(rstbsluar("KG_1"))
			GrandKg_2 =  GrandKg_2 + Cdbl(rstbsluar("KG_2"))
			GrandKg_3 =  GrandKg_3 + Cdbl(rstbsluar("KG_3"))
			GrandKg_4 =  GrandKg_4 + Cdbl(rstbsluar("KG_4"))
			GrandKg_5 =  GrandKg_5 + Cdbl(rstbsluar("KG_5"))
			GrandKg_6 =  GrandKg_6 + Cdbl(rstbsluar("KG_6"))
			GrandKg_7 =  GrandKg_7 + Cdbl(rstbsluar("KG_7"))
			GrandKg_8 =  GrandKg_8 + Cdbl(rstbsluar("KG_8"))
			GrandKg_9 =  GrandKg_9 + Cdbl(rstbsluar("KG_9"))
			GrandKg_10 =  GrandKg_10 + Cdbl(rstbsluar("KG_10"))
			GrandKg_11 =  GrandKg_11 + Cdbl(rstbsluar("KG_11"))
			GrandKg_12 =  GrandKg_12 + Cdbl(rstbsluar("KG_12"))
			GrandKg_13 =  GrandKg_13 + Cdbl(rstbsluar("KG_13"))
			GrandKg_14 =  GrandKg_14 + Cdbl(rstbsluar("KG_14"))
			GrandKg_15 =  GrandKg_15 + Cdbl(rstbsluar("KG_15"))
			GrandKg_16 =  GrandKg_16 + Cdbl(rstbsluar("KG_16"))
			GrandKg_17 =  GrandKg_17 + Cdbl(rstbsluar("KG_17"))
			GrandKg_18 =  GrandKg_18 + Cdbl(rstbsluar("KG_18"))
			GrandKg_19 =  GrandKg_19 + Cdbl(rstbsluar("KG_19"))
			GrandKg_20 =  GrandKg_20 + Cdbl(rstbsluar("KG_20"))
			GrandKg_21 =  GrandKg_21 + Cdbl(rstbsluar("KG_21"))
			GrandKg_22 =  GrandKg_22 + Cdbl(rstbsluar("KG_22"))
			GrandKg_23 =  GrandKg_23 + Cdbl(rstbsluar("KG_23"))
			GrandKg_24 =  GrandKg_24 + Cdbl(rstbsluar("KG_24"))
			GrandKg_25 =  GrandKg_25 + Cdbl(rstbsluar("KG_25"))
			GrandKg_26 =  GrandKg_26 + Cdbl(rstbsluar("KG_26"))
			GrandKg_27 =  GrandKg_27 + Cdbl(rstbsluar("KG_27"))
			GrandKg_28 =  GrandKg_28 + Cdbl(rstbsluar("KG_28"))
			GrandKg_29 =  GrandKg_29 + Cdbl(rstbsluar("KG_29"))
			GrandKg_30 =  GrandKg_30 + Cdbl(rstbsluar("KG_30"))
			GrandKg_31 =  GrandKg_31 + Cdbl(rstbsluar("KG_31"))
			GrandKg_Total =  GrandKg_Total + Cdbl(rstbsluar("KG_TOTAL"))
			
			GrandKgSortasi_1 =  GrandKgSortasi_1 + Cdbl(rstbsluar("KG_SORTAGE_1"))
			GrandKgSortasi_2 =  GrandKgSortasi_2 + Cdbl(rstbsluar("KG_SORTAGE_2"))
			GrandKgSortasi_3 =  GrandKgSortasi_3 + Cdbl(rstbsluar("KG_SORTAGE_3"))
			GrandKgSortasi_4 =  GrandKgSortasi_4 + Cdbl(rstbsluar("KG_SORTAGE_4"))
			GrandKgSortasi_5 =  GrandKgSortasi_5 + Cdbl(rstbsluar("KG_SORTAGE_5"))
			GrandKgSortasi_6 =  GrandKgSortasi_6 + Cdbl(rstbsluar("KG_SORTAGE_6"))
			GrandKgSortasi_7 =  GrandKgSortasi_7 + Cdbl(rstbsluar("KG_SORTAGE_7"))
			GrandKgSortasi_8 =  GrandKgSortasi_8 + Cdbl(rstbsluar("KG_SORTAGE_8"))
			GrandKgSortasi_9 =  GrandKgSortasi_9 + Cdbl(rstbsluar("KG_SORTAGE_9"))
			GrandKgSortasi_10 =  GrandKgSortasi_10 + Cdbl(rstbsluar("KG_SORTAGE_10"))
			GrandKgSortasi_11 =  GrandKgSortasi_11 + Cdbl(rstbsluar("KG_SORTAGE_11"))
			GrandKgSortasi_12 =  GrandKgSortasi_12 + Cdbl(rstbsluar("KG_SORTAGE_12"))
			GrandKgSortasi_13 =  GrandKgSortasi_13 + Cdbl(rstbsluar("KG_SORTAGE_13"))
			GrandKgSortasi_14 =  GrandKgSortasi_14 + Cdbl(rstbsluar("KG_SORTAGE_14"))
			GrandKgSortasi_15 =  GrandKgSortasi_15 + Cdbl(rstbsluar("KG_SORTAGE_15"))
			GrandKgSortasi_16 =  GrandKgSortasi_16 + Cdbl(rstbsluar("KG_SORTAGE_16"))
			GrandKgSortasi_17 =  GrandKgSortasi_17 + Cdbl(rstbsluar("KG_SORTAGE_17"))
			GrandKgSortasi_18 =  GrandKgSortasi_18 + Cdbl(rstbsluar("KG_SORTAGE_18"))
			GrandKgSortasi_19 =  GrandKgSortasi_19 + Cdbl(rstbsluar("KG_SORTAGE_19"))
			GrandKgSortasi_20 =  GrandKgSortasi_20 + Cdbl(rstbsluar("KG_SORTAGE_20"))
			GrandKgSortasi_21 =  GrandKgSortasi_21 + Cdbl(rstbsluar("KG_SORTAGE_21"))
			GrandKgSortasi_22 =  GrandKgSortasi_22 + Cdbl(rstbsluar("KG_SORTAGE_22"))
			GrandKgSortasi_23 =  GrandKgSortasi_23 + Cdbl(rstbsluar("KG_SORTAGE_23"))
			GrandKgSortasi_24 =  GrandKgSortasi_24 + Cdbl(rstbsluar("KG_SORTAGE_24"))
			GrandKgSortasi_25 =  GrandKgSortasi_25 + Cdbl(rstbsluar("KG_SORTAGE_25"))
			GrandKgSortasi_26 =  GrandKgSortasi_26 + Cdbl(rstbsluar("KG_SORTAGE_26"))
			GrandKgSortasi_27 =  GrandKgSortasi_27 + Cdbl(rstbsluar("KG_SORTAGE_27"))
			GrandKgSortasi_28 =  GrandKgSortasi_28 + Cdbl(rstbsluar("KG_SORTAGE_28"))
			GrandKgSortasi_29 =  GrandKgSortasi_29 + Cdbl(rstbsluar("KG_SORTAGE_29"))
			GrandKgSortasi_30 =  GrandKgSortasi_30 + Cdbl(rstbsluar("KG_SORTAGE_30"))
			GrandKgSortasi_31 =  GrandKgSortasi_31 + Cdbl(rstbsluar("KG_SORTAGE_31"))
			GrandKgSortasi_Total =  GrandKgSortasi_Total + Cdbl(rstbsluar("KG_SORTAGE_TOTAL"))
			
			rstbsluar.MoveNext
			
			Wend	
rstbsluar.Close
set rstbsluar = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td colspan="<%=var_colspan%>" align="center"><strong>GRAND TOTAL</strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKg_1, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_1, 2)%></strong></td>
		<%If GrandKg_1 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_1/GrandKg_1*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_2, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_2, 2)%></strong></td>
		<%If GrandKg_2 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_2/GrandKg_2*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_3, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_3, 2)%></strong></td>
		<%If GrandKg_3 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_3/GrandKg_3*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_4, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_4, 2)%></strong></td>
		<%If GrandKg_4 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_4/GrandKg_4*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_5, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_5, 2)%></strong></td>
		<%If GrandKg_5 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_5/GrandKg_5*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_6, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_6, 2)%></strong></td>
		<%If GrandKg_6 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_6/GrandKg_6*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_7, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_7, 2)%></strong></td>
		<%If GrandKg_7 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_7/GrandKg_7*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_8, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_8, 2)%></strong></td>
		<%If GrandKg_8 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_8/GrandKg_8*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_9, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_9, 2)%></strong></td>
		<%If GrandKg_9 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_9/GrandKg_9*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_10, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_10, 2)%></strong></td>
		<%If GrandKg_10 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_10/GrandKg_10*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_11, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_11, 2)%></strong></td>
		<%If GrandKg_11 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_11/GrandKg_11*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_12, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_12, 2)%></strong></td>
		<%If GrandKg_12 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_12/GrandKg_12*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_13, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_13, 2)%></strong></td>
		<%If GrandKg_13 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_13/GrandKg_13*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_14, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_14, 2)%></strong></td>
		<%If GrandKg_14 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_14/GrandKg_14*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_15, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_15, 2)%></strong></td>
		<%If GrandKg_15 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_15/GrandKg_15*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_16, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_16, 2)%></strong></td>
		<%If GrandKg_16 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_16/GrandKg_16*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_17, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_17, 2)%></strong></td>
		<%If GrandKg_17 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_17/GrandKg_17*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_18, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_18, 2)%></strong></td>
		<%If GrandKg_18 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_18/GrandKg_18*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_19, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_19, 2)%></strong></td>
		<%If GrandKg_19 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_19/GrandKg_19*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_20, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_20, 2)%></strong></td>
		<%If GrandKg_20 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_20/GrandKg_20*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_21, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_21, 2)%></strong></td>
		<%If GrandKg_21 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_21/GrandKg_21*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_22, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_22, 2)%></strong></td>
		<%If GrandKg_22 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_22/GrandKg_22*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_23, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_23, 2)%></strong></td>
		<%If GrandKg_23 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_23/GrandKg_23*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_24, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_24, 2)%></strong></td>
		<%If GrandKg_24 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_24/GrandKg_24*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_25, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_25, 2)%></strong></td>
		<%If GrandKg_25 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_25/GrandKg_25*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_26, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_26, 2)%></strong></td>
		<%If GrandKg_26 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_26/GrandKg_26*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_27, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_27, 2)%></strong></td>
		<%If GrandKg_27 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_27/GrandKg_27*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_28, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_28, 2)%></strong></td>
		<%If GrandKg_28 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_28/GrandKg_28*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_29, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_29, 2)%></strong></td>
		<%If GrandKg_29 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_29/GrandKg_29*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_30, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_30, 2)%></strong></td>
		<%If GrandKg_30 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_30/GrandKg_30*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_31, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_31, 2)%></strong></td>
		<%If GrandKg_31 = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_31/GrandKg_31*100, 2)%></strong></td>
		<%End If%>
		<td align="right"><strong><%=FormatNumber(GrandKg_Total, 2)%></strong></td>
		<td align="right"><strong><%=FormatNumber(GrandKgSortasi_Total, 2)%></strong></td>
		<%If GrandKg_Total = 0 Then%>
			<td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		<%Else%>
			<td align="right"><strong><%=FormatNumber(GrandKgSortasi_Total/GrandKg_Total*100, 2)%></strong></td>
		<%End If%>
	</tr>
<%End if%>
	</tbody>
</table>

   </body>
   </html>
