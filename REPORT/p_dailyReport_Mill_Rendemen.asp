<%@Language=VBScript%>

<%
' Option Explicit 
%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
	' DateTo and DateFrom will be EQUAL to remove miss-understanding
	
	Dim txtToDayDate, txtyear, txtmonth, txtday
	Dim strSQLPreParameter, rsSQLPreParameter, v_txtToDayDate, v_txtyear, v_txtmonth
	Dim strProduction, rsProduction
	Dim tempRefCode, strHRef, rsHRef, tempDescription, tempUOM
	Dim tempGrading_HI, tempGrading_SHI, tempGrading_SBI
	Dim tempDecrease_hi, tempDecrease_shi, tempDecrease_sbi
	Dim tempOpeningCPOStock_HI, tempDespatchCPOStock_HI
	Dim tempOpeningKernelStock_HI, tempDespatchKernelStock_HI
	Dim tempRounding
    Dim rscomp , strcomp,txcompanyname,str1,txtComp,v_txtComp,nColSpan
	Dim rsHourCapacityFinal
	Dim rsCPO_FFAQuality, rsCPO_MoistureQuality, rsCPO_DirtQuality
	Dim finalRowSpan
	Dim tempDataDivBudget_SHI, tempDataDivBudget_SBI
	Dim v_txtReportTitle	
	
	
    txtGroupPlantation  = Request("txtGroupPlantation")
	txtPlantation  		= Request("txtPlantation")
	txtGroupRegion  	= Request("txtGroupRegion")
	txtRegion  			= Request("txtRegion")
	txtcompany  		= Request("txtCompany")
	txtfcba 			= Request("txtfcba")
	txtToDayDate		= Request("txtdatefrom")
	txtMonth 		    = Request("txtMonth")
	
	' Now we make the HardCode...
	txtBudgetVersion = 1
	txtModuleVersion = 1
	
	v_txtReportTitle = "LAPORAN HARIAN RENDEMEN" 	
	
	
	if txtToDayDate = "" then
	    Dim strSQLDate, rsDate
        ' Get date from Server Date	by NAF
	    strSQLDate = "select to_char(sysdate,'MM/DD/YYYY') txtToDayDate from dual"
		Set rsDate = Server.CreateObject("ADODB.Recordset")
		rsDate.Open  strSQLDate,connstr, adOpenStatic
		if not rsDate.eof then
			txtToDayDate   = rsDate("txtToDayDate")
		End if
		rsDate.Close
		set rsDate = nothing
	end if
		
				
	
	strcomp= " select a.fccode, a.fccompanycode,a.fccompanyname from BUSINESSUNIT a  where a.fccompanycode = '" & txtcompany & "'  and a.fctype='M'"
	Set rscomp = Server.CreateObject("ADODB.Recordset")
	rscomp.Open  strcomp,connstr, adOpenStatic
	'response.write strcomp
	if not rscomp.eof then
		txcompanyname   = rscomp("fccompanyname")
        txtfcba 		= rscomp("fccode")
	End if
	rscomp.Close
	set rscomp = nothing
	

	strSQLPreParameter = "select to_char(to_date('" & txtToDayDate & "', 'MM/DD/YYYY'),'YYYY') as CurrentYear, " & _
	       "to_char(to_date('" & txtToDayDate & "', 'MM/DD/YYYY'),'MM') as CurrentPeriod, " & _
	       "to_char(to_date('" & txtToDayDate & "', 'MM/DD/YYYY'),'DD') as CurrentDay " & _
  		   " from dual "
	
	Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
	rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 
	
    if not rsSQLPreParameter.EOF then
		txtyear = rsSQLPreParameter("CurrentYear")
		txtmonth = rsSQLPreParameter("CurrentPeriod")
		txtday = cdbl(rsSQLPreParameter("CurrentDay"))
	end if
	rsSQLPreParameter.Close
	set rsSQLPreParameter = nothing	

	'response.write v_txtfcba
	
	v_txtGroupPlantation 	=   txtGroupPlantation
	v_txtPlantation 		=   txtPlantation
	v_txtGroupRegion 		=   txtGroupRegion
	v_txtRegion 			=   txtRegion
	v_txtCompany			=   txtCompany
	v_txtfcba     			=   txtfcba
	v_txtToDayDate			=	txtToDayDate
	v_txtdatefrom		    = 	txtToDayDate
	v_txtyear 				=   txtyear
	v_txtmonth				=   txtmonth
	v_txtBudgetYear 		= 	txtyear
	v_txtBudgetVersion 		= 	txtBudgetVersion
	v_txtModuleVersion 		= 	txtModuleVersion	
	
	
	' First
	
	
	strSELECTStmt = " CODEDESC, HEADGROUP, GROUPNAME, UNITOFMEASUREMENT "
	strSELECTStmt_Rendemen = ""

	if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , groupplantationcode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , groupplantationcode "
		end if
	else
	    strSELECTStmt = strSELECTStmt & " , groupplantationcode "
	    strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , groupplantationcode "
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , plantationcode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , plantationcode "
		end if
	else
	    strSELECTStmt = strSELECTStmt & " , plantationcode "
	    strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , groupregioncode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , groupregioncode "
		end if
	else
	    strSELECTStmt = strSELECTStmt & " , groupregioncode "
	    strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , regioncode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , regioncode "
		end if
	else
	    strSELECTStmt = strSELECTStmt & " , regioncode "
	    strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , companycode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , companycode "
		end if
	else
	    strSELECTStmt = strSELECTStmt & " , companycode "
	    strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , fcba "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , fcba "
		end if
	else
	    strSELECTStmt = strSELECTStmt & " , fcba "
	    strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , fcba "
	end if	

	strMAINStmt =   " , sum(nvl(seqnumber,0)) seqnumber, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),1, todate, 0 )) h_01, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 1, 1, 0)) div_01, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),2, todate, 0 )) h_02, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 2, 1, 0)) div_02, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),3, todate, 0 )) h_03, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 3, 1, 0)) div_03, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),4, todate, 0 )) h_04, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 4, 1, 0)) div_04, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),5, todate, 0 )) h_05, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 5, 1, 0)) div_05, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),6, todate, 0 )) h_06, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 6, 1, 0)) div_06, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),7, todate, 0 )) h_07, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 7, 1, 0)) div_07, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),8, todate, 0 )) h_08, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 8, 1, 0)) div_08, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),9, todate, 0 )) h_09, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 9, 1, 0)) div_09, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),10, todate, 0 )) h_10, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 10, 1, 0)) div_10, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),11, todate, 0 )) h_11, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 11, 1, 0)) div_11, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),12, todate, 0 )) h_12, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 12, 1, 0)) div_12, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),13, todate, 0 )) h_13, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 13, 1, 0)) div_13, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),14, todate, 0 )) h_14, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 14, 1, 0)) div_14, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),15, todate, 0 )) h_15, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 15, 1, 0)) div_15, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),16, todate, 0 )) h_16, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 16, 1, 0)) div_16, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),17, todate, 0 )) h_17, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 17, 1, 0)) div_17, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),18, todate, 0 )) h_18, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 18, 1, 0)) div_18, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),19, todate, 0 )) h_19, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 19, 1, 0)) div_19, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),20, todate, 0 )) h_20, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 20, 1, 0)) div_20, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),21, todate, 0 )) h_21, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 21, 1, 0)) div_21, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),22, todate, 0 )) h_22, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 22, 1, 0)) div_22, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),23, todate, 0 )) h_23, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 23, 1, 0)) div_23, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),24, todate, 0 )) h_24, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 24, 1, 0)) div_24, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),25, todate, 0 )) h_25, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 25, 1, 0)) div_25, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),26, todate, 0 )) h_26, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 26, 1, 0)) div_26, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),27, todate, 0 )) h_27, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 27, 1, 0)) div_27, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),28, todate, 0 )) h_28, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 28, 1, 0)) div_28, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),29, todate, 0 )) h_29, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 29, 1, 0)) div_29, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),30, todate, 0 )) h_30, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 30, 1, 0)) div_30, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),31, todate, 0 )) h_31, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), 31, 1, 0)) div_31, " & _					
					" sum(decode(to_number(to_char(ondate,'DD'))," & txtDay & ", tomonth, 0 )) shi, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), " & txtDay & ", 1, 0)) div_shi, " & _					
					" sum(decode(to_number(to_char(ondate,'DD'))," & txtDay & ", toyear, 0 )) sbi, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), " & txtDay & ", 1, 0)) div_sbi, " & _					
					" sum(decode(to_number(to_char(ondate,'DD'))," & txtDay & ", budgettomonth, 0 )) budget_shi, " & _
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), " & txtDay & ", 1, 0)) div_budget_shi, " & _					
					" sum(decode(to_number(to_char(ondate,'DD'))," & txtDay & ", budgettoyear, 0 )) budget_sbi, " & _ 
					" sum(DECODE (TO_NUMBER (TO_CHAR (ondate, 'DD')), " & txtDay & ", 1, 0)) div_budget_sbi " 					

	
				
	strMAINAVGStmt = " , avg(nvl(seqnumber,0)) seqnumber, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),1, todate, 0 )) h_01, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),2, todate, 0 )) h_02, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),3, todate, 0 )) h_03, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),4, todate, 0 )) h_04, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),5, todate, 0 )) h_05, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),6, todate, 0 )) h_06, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),7, todate, 0 )) h_07, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),8, todate, 0 )) h_08, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),9, todate, 0 )) h_09, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),10, todate, 0 )) h_10, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),11, todate, 0 )) h_11, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),12, todate, 0 )) h_12, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),13, todate, 0 )) h_13, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),14, todate, 0 )) h_14, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),15, todate, 0 )) h_15, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),16, todate, 0 )) h_16, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),17, todate, 0 )) h_17, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),18, todate, 0 )) h_18, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),19, todate, 0 )) h_19, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),20, todate, 0 )) h_20, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),21, todate, 0 )) h_21, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),22, todate, 0 )) h_22, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),23, todate, 0 )) h_23, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),24, todate, 0 )) h_24, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),25, todate, 0 )) h_25, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),26, todate, 0 )) h_26, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),27, todate, 0 )) h_27, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),28, todate, 0 )) h_28, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),29, todate, 0 )) h_29, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),30, todate, 0 )) h_30, " & _
					" avg(decode(to_number(to_char(ondate,'DD')),31, todate, 0 )) h_31, " & _
					" avg(budgettomonth) budget_shi,  " & _
					" avg(budgettoyear) budget_sbi   " 
				
	
	
    strMAIN2Stmt =   " , seqnumber, " & _
					" h_01, " & _
					" h_02, " & _
					" h_03, " & _
					" h_04, " & _
					" h_05, " & _
					" h_06, " & _
					" h_07, " & _
					" h_08, " & _
					" h_09, " & _
					" h_10, " & _
					" h_11, " & _
					" h_12, " & _
					" h_13, " & _
					" h_14, " & _
					" h_15, " & _
					" h_16, " & _
					" h_17, " & _
					" h_18, " & _
					" h_19, " & _
					" h_20, " & _
					" h_21, " & _
					" h_22, " & _
					" h_23, " & _
					" h_24, " & _
					" h_25, " & _
					" h_26, " & _
					" h_27, " & _
					" h_28, " & _
					" h_29, " & _
					" h_30, " & _
					" h_31, " & _
					" div_01, " & _
					" div_02, " & _
					" div_03, " & _
					" div_04, " & _
					" div_05, " & _
					" div_06, " & _
					" div_07, " & _
					" div_08, " & _
					" div_09, " & _
					" div_10, " & _
					" div_11, " & _
					" div_12, " & _
					" div_13, " & _
					" div_14, " & _
					" div_15, " & _
					" div_16, " & _
					" div_17, " & _
					" div_18, " & _
					" div_19, " & _
					" div_20, " & _
					" div_21, " & _
					" div_22, " & _
					" div_23, " & _
					" div_24, " & _
					" div_25, " & _
					" div_26, " & _
					" div_27, " & _
					" div_28, " & _
					" div_29, " & _
					" div_30, " & _
					" div_31, " & _
					" shi,  " & _
					" div_shi,  " & _
					" sbi,  " & _
					" div_sbi,  " & _
					" budget_shi,  " & _
					" div_budget_shi,  " & _
					" budget_sbi,   " & _					
					" div_budget_sbi   " 					

    strMAIN2AVGStmt =   " , avg(seqnumber) seqnumber, " & _
					" avg(h_01) h_01, " & _
					" avg(h_02) h_02, " & _
					" avg(h_03) h_03, " & _
					" avg(h_04) h_04, " & _
					" avg(h_05) h_05, " & _
					" avg(h_06) h_06, " & _
					" avg(h_07) h_07, " & _
					" avg(h_08) h_08, " & _
					" avg(h_09) h_09, " & _
					" avg(h_10) h_10, " & _
					" avg(h_11) h_11, " & _
					" avg(h_12) h_12, " & _
					" avg(h_13) h_13, " & _
					" avg(h_14) h_14, " & _
					" avg(h_15) h_15, " & _
					" avg(h_16) h_16, " & _
					" avg(h_17) h_17, " & _
					" avg(h_18) h_18, " & _
					" avg(h_19) h_19, " & _
					" avg(h_20) h_20, " & _
					" avg(h_21) h_21, " & _
					" avg(h_22) h_22, " & _
					" avg(h_23) h_23, " & _
					" avg(h_24) h_24, " & _
					" avg(h_25) h_25, " & _
					" avg(h_26) h_26, " & _
					" avg(h_27) h_27, " & _
					" avg(h_28) h_28, " & _
					" avg(h_29) h_29, " & _
					" avg(h_30) h_30, " & _
					" avg(h_31) h_31, " & _
					" avg(budget_shi) budget_shi,  " & _
					" avg(budget_sbi) budget_sbi   " 					
					
					

				
					
					
	strFROMStmt = " from mill_summary_tab  " & _
					" where  " & _
					" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) is not null   and "  & _
					" to_number(to_char(ondate,'YYYY')) = '" & v_txtYear & "' and " & _
					" to_number(to_char(ondate,'MM')) = '" & v_txtMonth & "' and " 
					
					
    ' Where Condion
	strWHEREStmt = ""
	
	if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND GroupPlantationCode like '%' " 
		else
		    if v_txtGroupPlantation <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND GroupPlantationCode = '" & v_txtGroupPlantation& "' " 
			end if
		end if
	else
	    strWHEREStmt = strWHEREStmt & " AND GroupPlantationCode like '%' " 
	end if
	
	
	if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND PlantationCode like '%' " 
		else
		    if v_txtPlantation <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND PlantationCode = '" & v_txtPlantation& "' " 
			end if
		end if
	else
	    strWHEREStmt = strWHEREStmt & " AND PlantationCode like '%' " 
	end if
	
	
	if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND GroupRegionCode like '%' " 
		else
		    if v_txtGroupRegion <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND GroupRegionCode = '" & v_txtGroupRegion& "' " 
			end if
		end if
	else
	    strWHEREStmt = strWHEREStmt & " AND GroupRegionCode like '%' " 
	end if
	

	if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND RegionCode like '%' " 
		else
		    if v_txtRegion <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND RegionCode = '" & v_txtRegion& "' " 
			end if
		end if
	else
	    strWHEREStmt = strWHEREStmt & " AND RegionCode like '%' " 
	end if
	
	if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND CompanyCode like '%' " 
		else
		    if v_txtcompany <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND CompanyCode = '" & v_txtcompany& "' " 
			end if
		end if
	else
	    strWHEREStmt = strWHEREStmt & " AND CompanyCode like '%' " 
	end if	


	if len(trim(v_txtfcba)) > 0 then 
		if v_txtfcba = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND fcba like '%' " 
		else
		    if v_txtfcba <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND fcba = '" & v_txtfcba& "' " 
			end if
		end if
	else
	    strWHEREStmt = strWHEREStmt & " AND fcba like '%' " 
	end if	

					
	' Group By
	
	strGROUPStmt = ""
	
	strGROUPStmt = strGROUPStmt & " GROUP BY CODEDESC, HEADGROUP, GROUPNAME, UNITOFMEASUREMENT "
	
    if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , groupplantationcode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , groupplantationcode "
		end if
	else
	    strGROUPStmt = strGROUPStmt & " , groupplantationcode "
	    strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , groupplantationcode "
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , plantationcode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , plantationcode "
		end if
	else
	    strGROUPStmt = strGROUPStmt & " , plantationcode "
	    strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , groupregioncode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , groupregioncode "
		end if
	else
	    strGROUPStmt = strGROUPStmt & " , groupregioncode "
	    strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , regioncode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , regioncode "
		end if
	else
	    strGROUPStmt = strGROUPStmt & " , regioncode "
	    strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , companycode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , companycode "
		end if
	else
	    strGROUPStmt = strGROUPStmt & " , companycode "
	    strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , fcba "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , fcba "
		end if
	else
	    strGROUPStmt = strGROUPStmt & " , fcba "
	    strGROUPRENDEMENStmt = strGROUPRENDEMENStmt & " , fcba "
	end if	
	
	strORDERStmt = " ORDER BY 'x' " 	
	
    if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , groupplantationcode "
		end if
	else
	    strORDERStmt = strORDERStmt & " , groupplantationcode "
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , plantationcode "
		end if
	else
	    strORDERStmt = strORDERStmt & " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , groupregioncode "
		end if
	else
	    strORDERStmt = strORDERStmt & " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , regioncode "
		end if
	else
	    strORDERStmt = strORDERStmt & " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , companycode "
		end if
	else
	    strORDERStmt = strORDERStmt & " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , fcba "
		end if
	else
	    strORDERStmt = strORDERStmt & " , fcba "
	end if		
	
    strORDERStmt = strORDERStmt & " , seqnumber "



	
	
' *************************SPECIAL FOR RENDEMEN************************
	

	' Step RENDEMEN UTAMA
	strSTEPStmt = " stepnumber = 'STEP_03' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'FFB0701' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) " 

	
	strRENDEMENstmt =   ", sum(h_01) tbs_process_h_01, " & _
				" sum(h_02) tbs_process_h_02, " & _
				" sum(h_03) tbs_process_h_03, " & _
				" sum(h_04) tbs_process_h_04, " & _
				" sum(h_05) tbs_process_h_05, " & _
				" sum(h_06) tbs_process_h_06, " & _
				" sum(h_07) tbs_process_h_07, " & _
				" sum(h_08) tbs_process_h_08, " & _
				" sum(h_09) tbs_process_h_09, " & _
				" sum(h_10) tbs_process_h_10, " & _
				" sum(h_11) tbs_process_h_11, " & _
				" sum(h_12) tbs_process_h_12, " & _
				" sum(h_13) tbs_process_h_13, " & _
				" sum(h_14) tbs_process_h_14, " & _
				" sum(h_15) tbs_process_h_15, " & _
				" sum(h_16) tbs_process_h_16, " & _
				" sum(h_17) tbs_process_h_17, " & _
				" sum(h_18) tbs_process_h_18, " & _
				" sum(h_19) tbs_process_h_19, " & _
				" sum(h_20) tbs_process_h_20, " & _
				" sum(h_21) tbs_process_h_21, " & _
				" sum(h_22) tbs_process_h_22, " & _
				" sum(h_23) tbs_process_h_23, " & _
				" sum(h_24) tbs_process_h_24, " & _
				" sum(h_25) tbs_process_h_25, " & _
				" sum(h_26) tbs_process_h_26, " & _
				" sum(h_27) tbs_process_h_27, " & _
				" sum(h_28) tbs_process_h_28, " & _
				" sum(h_29) tbs_process_h_29, " & _
				" sum(h_30) tbs_process_h_30, " & _
				" sum(h_31) tbs_process_h_31, " & _
				" sum(shi) tbs_process_shi, " & _
				" sum(sbi) tbs_process_sbi, " & _
				" sum(budget_shi) tbs_process_budget_shi,  " & _
				" sum(budget_sbi) tbs_process_budget_sbi   " 	
	strRENDEMENstmt = strRENDEMENstmt &  ", 0 cpo_production_h_01, " & _
				" 0 cpo_production_h_02, " & _
				" 0 cpo_production_h_03, " & _
				" 0 cpo_production_h_04, " & _
				" 0 cpo_production_h_05, " & _
				" 0 cpo_production_h_06, " & _
				" 0 cpo_production_h_07, " & _
				" 0 cpo_production_h_08, " & _
				" 0 cpo_production_h_09, " & _
				" 0 cpo_production_h_10, " & _
				" 0 cpo_production_h_11, " & _
				" 0 cpo_production_h_12, " & _
				" 0 cpo_production_h_13, " & _
				" 0 cpo_production_h_14, " & _
				" 0 cpo_production_h_15, " & _
				" 0 cpo_production_h_16, " & _
				" 0 cpo_production_h_17, " & _
				" 0 cpo_production_h_18, " & _
				" 0 cpo_production_h_19, " & _
				" 0 cpo_production_h_20, " & _
				" 0 cpo_production_h_21, " & _
				" 0 cpo_production_h_22, " & _
				" 0 cpo_production_h_23, " & _
				" 0 cpo_production_h_24, " & _
				" 0 cpo_production_h_25, " & _
				" 0 cpo_production_h_26, " & _
				" 0 cpo_production_h_27, " & _
				" 0 cpo_production_h_28, " & _
				" 0 cpo_production_h_29, " & _
				" 0 cpo_production_h_30, " & _
				" 0 cpo_production_h_31, " & _	
				" 0 cpo_production_shi, " & _
				" 0 cpo_production_sbi, " & _
				" 0 cpo_production_budget_shi,  " & _
				" 0 cpo_production_budget_sbi   " 	
	strRENDEMENstmt =  strRENDEMENstmt & ", 0 kernel_production_h_01, " & _
				" 0 kernel_production_h_02, " & _
				" 0 kernel_production_h_03, " & _
				" 0 kernel_production_h_04, " & _
				" 0 kernel_production_h_05, " & _
				" 0 kernel_production_h_06, " & _
				" 0 kernel_production_h_07, " & _
				" 0 kernel_production_h_08, " & _
				" 0 kernel_production_h_09, " & _
				" 0 kernel_production_h_10, " & _
				" 0 kernel_production_h_11, " & _
				" 0 kernel_production_h_12, " & _
				" 0 kernel_production_h_13, " & _
				" 0 kernel_production_h_14, " & _
				" 0 kernel_production_h_15, " & _
				" 0 kernel_production_h_16, " & _
				" 0 kernel_production_h_17, " & _
				" 0 kernel_production_h_18, " & _
				" 0 kernel_production_h_19, " & _
				" 0 kernel_production_h_20, " & _
				" 0 kernel_production_h_21, " & _
				" 0 kernel_production_h_22, " & _
				" 0 kernel_production_h_23, " & _
				" 0 kernel_production_h_24, " & _
				" 0 kernel_production_h_25, " & _
				" 0 kernel_production_h_26, " & _
				" 0 kernel_production_h_27, " & _
				" 0 kernel_production_h_28, " & _
				" 0 kernel_production_h_29, " & _
				" 0 kernel_production_h_30, " & _
				" 0 kernel_production_h_31, " & _	
				" 0 kernel_production_shi, " & _
				" 0 kernel_production_sbi, " & _
				" 0 kernel_production_budget_shi,  " & _
				" 0 kernel_production_budget_sbi   " 	


	strRendemen = "SELECT 'RENDEMEN' title " & strSELECTStmt_Rendemen & strRENDEMENstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN_DETAIL' " & strGROUPRENDEMENStmt
	
	'======================CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_06' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'PROD01' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strRENDEMENstmt =   ", 0 tbs_process_h_01, " & _
				" 0 tbs_process_h_02, " & _
				" 0 tbs_process_h_03, " & _
				" 0 tbs_process_h_04, " & _
				" 0 tbs_process_h_05, " & _
				" 0 tbs_process_h_06, " & _
				" 0 tbs_process_h_07, " & _
				" 0 tbs_process_h_08, " & _
				" 0 tbs_process_h_09, " & _
				" 0 tbs_process_h_10, " & _
				" 0 tbs_process_h_11, " & _
				" 0 tbs_process_h_12, " & _
				" 0 tbs_process_h_13, " & _
				" 0 tbs_process_h_14, " & _
				" 0 tbs_process_h_15, " & _
				" 0 tbs_process_h_16, " & _
				" 0 tbs_process_h_17, " & _
				" 0 tbs_process_h_18, " & _
				" 0 tbs_process_h_19, " & _
				" 0 tbs_process_h_20, " & _
				" 0 tbs_process_h_21, " & _
				" 0 tbs_process_h_22, " & _
				" 0 tbs_process_h_23, " & _
				" 0 tbs_process_h_24, " & _
				" 0 tbs_process_h_25, " & _
				" 0 tbs_process_h_26, " & _
				" 0 tbs_process_h_27, " & _
				" 0 tbs_process_h_28, " & _
				" 0 tbs_process_h_29, " & _
				" 0 tbs_process_h_30, " & _
				" 0 tbs_process_h_31, " & _		
				" 0 tbs_process_shi, " & _
				" 0 tbs_process_sbi, " & _
				" 0 tbs_process_budget_shi,  " & _
				" 0 tbs_process_budget_sbi   " 	
	strRENDEMENstmt = strRENDEMENstmt &  ", sum(h_01) cpo_production_h_01, " & _
				" sum(h_02) cpo_production_h_02, " & _
				" sum(h_03) cpo_production_h_03, " & _
				" sum(h_04) cpo_production_h_04, " & _
				" sum(h_05) cpo_production_h_05, " & _
				" sum(h_06) cpo_production_h_06, " & _
				" sum(h_07) cpo_production_h_07, " & _
				" sum(h_08) cpo_production_h_08, " & _
				" sum(h_09) cpo_production_h_09, " & _
				" sum(h_10) cpo_production_h_10, " & _
				" sum(h_11) cpo_production_h_11, " & _
				" sum(h_12) cpo_production_h_12, " & _
				" sum(h_13) cpo_production_h_13, " & _
				" sum(h_14) cpo_production_h_14, " & _
				" sum(h_15) cpo_production_h_15, " & _
				" sum(h_16) cpo_production_h_16, " & _
				" sum(h_17) cpo_production_h_17, " & _
				" sum(h_18) cpo_production_h_18, " & _
				" sum(h_19) cpo_production_h_19, " & _
				" sum(h_20) cpo_production_h_20, " & _
				" sum(h_21) cpo_production_h_21, " & _
				" sum(h_22) cpo_production_h_22, " & _
				" sum(h_23) cpo_production_h_23, " & _
				" sum(h_24) cpo_production_h_24, " & _
				" sum(h_25) cpo_production_h_25, " & _
				" sum(h_26) cpo_production_h_26, " & _
				" sum(h_27) cpo_production_h_27, " & _
				" sum(h_28) cpo_production_h_28, " & _
				" sum(h_29) cpo_production_h_29, " & _
				" sum(h_30) cpo_production_h_30, " & _
				" sum(h_31) cpo_production_h_31, " & _	
				" sum(shi) cpo_production_shi, " & _
				" sum(sbi) cpo_production_sbi, " & _
				" sum(budget_shi) cpo_production_budget_shi,  " & _
				" sum(budget_sbi) cpo_production_budget_sbi   " 	
	strRENDEMENstmt =  strRENDEMENstmt & ", 0 kernel_production_h_01, " & _
				" 0 kernel_production_h_02, " & _
				" 0 kernel_production_h_03, " & _
				" 0 kernel_production_h_04, " & _
				" 0 kernel_production_h_05, " & _
				" 0 kernel_production_h_06, " & _
				" 0 kernel_production_h_07, " & _
				" 0 kernel_production_h_08, " & _
				" 0 kernel_production_h_09, " & _
				" 0 kernel_production_h_10, " & _
				" 0 kernel_production_h_11, " & _
				" 0 kernel_production_h_12, " & _
				" 0 kernel_production_h_13, " & _
				" 0 kernel_production_h_14, " & _
				" 0 kernel_production_h_15, " & _
				" 0 kernel_production_h_16, " & _
				" 0 kernel_production_h_17, " & _
				" 0 kernel_production_h_18, " & _
				" 0 kernel_production_h_19, " & _
				" 0 kernel_production_h_20, " & _
				" 0 kernel_production_h_21, " & _
				" 0 kernel_production_h_22, " & _
				" 0 kernel_production_h_23, " & _
				" 0 kernel_production_h_24, " & _
				" 0 kernel_production_h_25, " & _
				" 0 kernel_production_h_26, " & _
				" 0 kernel_production_h_27, " & _
				" 0 kernel_production_h_28, " & _
				" 0 kernel_production_h_29, " & _
				" 0 kernel_production_h_30, " & _
				" 0 kernel_production_h_31, " & _	
				" 0 kernel_production_shi, " & _
				" 0 kernel_production_sbi, " & _
				" 0 kernel_production_budget_shi,  " & _
				" 0 kernel_production_budget_sbi   " 		

	strRendemen = strRendemen & " UNION SELECT 'RENDEMEN' title " & strSELECTStmt_Rendemen & strRENDEMENstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN' " & strGROUPRENDEMENStmt
	
	'======================CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_06' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'PROD02' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strRENDEMENstmt =   ", 0 tbs_process_h_01, " & _
				" 0 tbs_process_h_02, " & _
				" 0 tbs_process_h_03, " & _
				" 0 tbs_process_h_04, " & _
				" 0 tbs_process_h_05, " & _
				" 0 tbs_process_h_06, " & _
				" 0 tbs_process_h_07, " & _
				" 0 tbs_process_h_08, " & _
				" 0 tbs_process_h_09, " & _
				" 0 tbs_process_h_10, " & _
				" 0 tbs_process_h_11, " & _
				" 0 tbs_process_h_12, " & _
				" 0 tbs_process_h_13, " & _
				" 0 tbs_process_h_14, " & _
				" 0 tbs_process_h_15, " & _
				" 0 tbs_process_h_16, " & _
				" 0 tbs_process_h_17, " & _
				" 0 tbs_process_h_18, " & _
				" 0 tbs_process_h_19, " & _
				" 0 tbs_process_h_20, " & _
				" 0 tbs_process_h_21, " & _
				" 0 tbs_process_h_22, " & _
				" 0 tbs_process_h_23, " & _
				" 0 tbs_process_h_24, " & _
				" 0 tbs_process_h_25, " & _
				" 0 tbs_process_h_26, " & _
				" 0 tbs_process_h_27, " & _
				" 0 tbs_process_h_28, " & _
				" 0 tbs_process_h_29, " & _
				" 0 tbs_process_h_30, " & _
				" 0 tbs_process_h_31, " & _		
				" 0 tbs_process_shi, " & _
				" 0 tbs_process_sbi, " & _
				" 0 tbs_process_budget_shi,  " & _
				" 0 tbs_process_budget_sbi   " 		
	strRENDEMENstmt = strRENDEMENstmt &  ", 0 cpo_production_h_01, " & _
				" 0 cpo_production_h_02, " & _
				" 0 cpo_production_h_03, " & _
				" 0 cpo_production_h_04, " & _
				" 0 cpo_production_h_05, " & _
				" 0 cpo_production_h_06, " & _
				" 0 cpo_production_h_07, " & _
				" 0 cpo_production_h_08, " & _
				" 0 cpo_production_h_09, " & _
				" 0 cpo_production_h_10, " & _
				" 0 cpo_production_h_11, " & _
				" 0 cpo_production_h_12, " & _
				" 0 cpo_production_h_13, " & _
				" 0 cpo_production_h_14, " & _
				" 0 cpo_production_h_15, " & _
				" 0 cpo_production_h_16, " & _
				" 0 cpo_production_h_17, " & _
				" 0 cpo_production_h_18, " & _
				" 0 cpo_production_h_19, " & _
				" 0 cpo_production_h_20, " & _
				" 0 cpo_production_h_21, " & _
				" 0 cpo_production_h_22, " & _
				" 0 cpo_production_h_23, " & _
				" 0 cpo_production_h_24, " & _
				" 0 cpo_production_h_25, " & _
				" 0 cpo_production_h_26, " & _
				" 0 cpo_production_h_27, " & _
				" 0 cpo_production_h_28, " & _
				" 0 cpo_production_h_29, " & _
				" 0 cpo_production_h_30, " & _
				" 0 cpo_production_h_31, " & _	
				" 0 cpo_production_shi, " & _
				" 0 cpo_production_sbi, " & _
				" 0 cpo_production_budget_shi,  " & _
				" 0 cpo_production_budget_sbi   " 	
	strRENDEMENstmt = strRENDEMENstmt &  ", sum(h_01) kernel_production_h_01, " & _
				" sum(h_02) kernel_production_h_02, " & _
				" sum(h_03) kernel_production_h_03, " & _
				" sum(h_04) kernel_production_h_04, " & _
				" sum(h_05) kernel_production_h_05, " & _
				" sum(h_06) kernel_production_h_06, " & _
				" sum(h_07) kernel_production_h_07, " & _
				" sum(h_08) kernel_production_h_08, " & _
				" sum(h_09) kernel_production_h_09, " & _
				" sum(h_10) kernel_production_h_10, " & _
				" sum(h_11) kernel_production_h_11, " & _
				" sum(h_12) kernel_production_h_12, " & _
				" sum(h_13) kernel_production_h_13, " & _
				" sum(h_14) kernel_production_h_14, " & _
				" sum(h_15) kernel_production_h_15, " & _
				" sum(h_16) kernel_production_h_16, " & _
				" sum(h_17) kernel_production_h_17, " & _
				" sum(h_18) kernel_production_h_18, " & _
				" sum(h_19) kernel_production_h_19, " & _
				" sum(h_20) kernel_production_h_20, " & _
				" sum(h_21) kernel_production_h_21, " & _
				" sum(h_22) kernel_production_h_22, " & _
				" sum(h_23) kernel_production_h_23, " & _
				" sum(h_24) kernel_production_h_24, " & _
				" sum(h_25) kernel_production_h_25, " & _
				" sum(h_26) kernel_production_h_26, " & _
				" sum(h_27) kernel_production_h_27, " & _
				" sum(h_28) kernel_production_h_28, " & _
				" sum(h_29) kernel_production_h_29, " & _
				" sum(h_30) kernel_production_h_30, " & _
				" sum(h_31) kernel_production_h_31, " & _	
				" sum(shi) kernel_production_shi, " & _
				" sum(sbi) kernel_production_sbi, " & _
				" sum(budget_shi) kernel_production_budget_shi,  " & _
				" sum(budget_sbi) kernel_production_budget_sbi   " 					

	strRendemen = strRendemen & " UNION SELECT 'RENDEMEN' title " & strSELECTStmt_Rendemen & strRENDEMENstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN' " & strGROUPRENDEMENStmt
	
	
	
	strRENDEMENFinalStmt = " , sum(tbs_process_h_01) tbs_process_h_01 , " & _
							" sum(tbs_process_h_02) tbs_process_h_02, " & _
							" sum(tbs_process_h_03) tbs_process_h_03, " & _
							" sum(tbs_process_h_04) tbs_process_h_04, " & _
							" sum(tbs_process_h_05) tbs_process_h_05, " & _
							" sum(tbs_process_h_06) tbs_process_h_06, " & _
							" sum(tbs_process_h_07) tbs_process_h_07, " & _
							" sum(tbs_process_h_08) tbs_process_h_08, " & _
							" sum(tbs_process_h_09) tbs_process_h_09, " & _
							" sum(tbs_process_h_10) tbs_process_h_10, " & _
							" sum(tbs_process_h_11) tbs_process_h_11, " & _
							" sum(tbs_process_h_12) tbs_process_h_12, " & _
							" sum(tbs_process_h_13) tbs_process_h_13, " & _
							" sum(tbs_process_h_14) tbs_process_h_14, " & _
							" sum(tbs_process_h_15) tbs_process_h_15, " & _
							" sum(tbs_process_h_16) tbs_process_h_16, " & _
							" sum(tbs_process_h_17) tbs_process_h_17, " & _
							" sum(tbs_process_h_18) tbs_process_h_18, " & _
							" sum(tbs_process_h_19) tbs_process_h_19, " & _
							" sum(tbs_process_h_20) tbs_process_h_20, " & _
							" sum(tbs_process_h_21) tbs_process_h_21, " & _
							" sum(tbs_process_h_22) tbs_process_h_22, " & _
							" sum(tbs_process_h_23) tbs_process_h_23, " & _
							" sum(tbs_process_h_24) tbs_process_h_24, " & _
							" sum(tbs_process_h_25) tbs_process_h_25, " & _
							" sum(tbs_process_h_26) tbs_process_h_26, " & _
							" sum(tbs_process_h_27) tbs_process_h_27, " & _
							" sum(tbs_process_h_28) tbs_process_h_28, " & _
							" sum(tbs_process_h_29) tbs_process_h_29, " & _
							" sum(tbs_process_h_30) tbs_process_h_30, " & _
							" sum(tbs_process_h_31) tbs_process_h_31, " & _	
						    " sum(tbs_process_shi) tbs_process_shi , " & _
						    " sum(tbs_process_sbi) tbs_process_sbi , " & _
						    " sum(tbs_process_budget_shi) tbs_process_budget_shi , " & _
						    " sum(tbs_process_budget_sbi) tbs_process_budget_sbi , " & _
							" sum(cpo_production_h_01) cpo_production_h_01 , " & _
							" sum(cpo_production_h_02) cpo_production_h_02, " & _
							" sum(cpo_production_h_03) cpo_production_h_03, " & _
							" sum(cpo_production_h_04) cpo_production_h_04, " & _
							" sum(cpo_production_h_05) cpo_production_h_05, " & _
							" sum(cpo_production_h_06) cpo_production_h_06, " & _
							" sum(cpo_production_h_07) cpo_production_h_07, " & _
							" sum(cpo_production_h_08) cpo_production_h_08, " & _
							" sum(cpo_production_h_09) cpo_production_h_09, " & _
							" sum(cpo_production_h_10) cpo_production_h_10, " & _
							" sum(cpo_production_h_11) cpo_production_h_11, " & _
							" sum(cpo_production_h_12) cpo_production_h_12, " & _
							" sum(cpo_production_h_13) cpo_production_h_13, " & _
							" sum(cpo_production_h_14) cpo_production_h_14, " & _
							" sum(cpo_production_h_15) cpo_production_h_15, " & _
							" sum(cpo_production_h_16) cpo_production_h_16, " & _
							" sum(cpo_production_h_17) cpo_production_h_17, " & _
							" sum(cpo_production_h_18) cpo_production_h_18, " & _
							" sum(cpo_production_h_19) cpo_production_h_19, " & _
							" sum(cpo_production_h_20) cpo_production_h_20, " & _
							" sum(cpo_production_h_21) cpo_production_h_21, " & _
							" sum(cpo_production_h_22) cpo_production_h_22, " & _
							" sum(cpo_production_h_23) cpo_production_h_23, " & _
							" sum(cpo_production_h_24) cpo_production_h_24, " & _
							" sum(cpo_production_h_25) cpo_production_h_25, " & _
							" sum(cpo_production_h_26) cpo_production_h_26, " & _
							" sum(cpo_production_h_27) cpo_production_h_27, " & _
							" sum(cpo_production_h_28) cpo_production_h_28, " & _
							" sum(cpo_production_h_29) cpo_production_h_29, " & _
							" sum(cpo_production_h_30) cpo_production_h_30, " & _
							" sum(cpo_production_h_31) cpo_production_h_31, " & _				
						    " sum(cpo_production_shi) cpo_production_shi , " & _
						    " sum(cpo_production_sbi) cpo_production_sbi , " & _
						    " sum(cpo_production_budget_shi) cpo_production_budget_shi , " & _
						    " sum(cpo_production_budget_sbi) cpo_production_budget_sbi , " & _						   
							" sum(kernel_production_h_01) kernel_production_h_01 , " & _
							" sum(kernel_production_h_02) kernel_production_h_02, " & _
							" sum(kernel_production_h_03) kernel_production_h_03, " & _
							" sum(kernel_production_h_04) kernel_production_h_04, " & _
							" sum(kernel_production_h_05) kernel_production_h_05, " & _
							" sum(kernel_production_h_06) kernel_production_h_06, " & _
							" sum(kernel_production_h_07) kernel_production_h_07, " & _
							" sum(kernel_production_h_08) kernel_production_h_08, " & _
							" sum(kernel_production_h_09) kernel_production_h_09, " & _
							" sum(kernel_production_h_10) kernel_production_h_10, " & _
							" sum(kernel_production_h_11) kernel_production_h_11, " & _
							" sum(kernel_production_h_12) kernel_production_h_12, " & _
							" sum(kernel_production_h_13) kernel_production_h_13, " & _
							" sum(kernel_production_h_14) kernel_production_h_14, " & _
							" sum(kernel_production_h_15) kernel_production_h_15, " & _
							" sum(kernel_production_h_16) kernel_production_h_16, " & _
							" sum(kernel_production_h_17) kernel_production_h_17, " & _
							" sum(kernel_production_h_18) kernel_production_h_18, " & _
							" sum(kernel_production_h_19) kernel_production_h_19, " & _
							" sum(kernel_production_h_20) kernel_production_h_20, " & _
							" sum(kernel_production_h_21) kernel_production_h_21, " & _
							" sum(kernel_production_h_22) kernel_production_h_22, " & _
							" sum(kernel_production_h_23) kernel_production_h_23, " & _
							" sum(kernel_production_h_24) kernel_production_h_24, " & _
							" sum(kernel_production_h_25) kernel_production_h_25, " & _
							" sum(kernel_production_h_26) kernel_production_h_26, " & _
							" sum(kernel_production_h_27) kernel_production_h_27, " & _
							" sum(kernel_production_h_28) kernel_production_h_28, " & _
							" sum(kernel_production_h_29) kernel_production_h_29, " & _
							" sum(kernel_production_h_30) kernel_production_h_30, " & _
							" sum(kernel_production_h_31) kernel_production_h_31, " & _							   
						   " sum(kernel_production_shi) kernel_production_shi , " & _
						   " sum(kernel_production_sbi) kernel_production_sbi , " & _
						   " sum(kernel_production_budget_shi) kernel_production_budget_shi , " & _
						   " sum(kernel_production_budget_sbi) kernel_production_budget_sbi  " 
						   
	strRendemenFinal = "SELECT title HEADGROUP, '%' UnitOfMeasurement, '' CodeDesc " & strSELECTStmt_Rendemen & strRENDEMENFinalStmt & " FROM ( " & strRendemen & " ) " & _
	                   " GROUP BY title " & strGROUPRENDEMENStmt
	
	

	Set rsRendemenFinal = Server.CreateObject("ADODB.Recordset")
	'response.write strRendemenFinal
	rsRendemenFinal.Open strRendemenFinal, connstr, 3,3	

	
	' *************************SPECIAL FOR RENDEMEN************************
	
	

	
    %>


<html>

	<head>
		<title>iPla'S Reporting - Daily Production CPO Quality Report</title>
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
       </head>


	<body>
		<div id="wrapper">
			<form name="frmEntry" method="post" action="p_DailyReport_Mill_Rendemen.asp">
				<!--#include file="menutop.asp"-->	
				<!--#include file="FormSelectNewSum.html"-->
				<!--#include file="FormSelectDailyReportMill_CPOQuality.html"-->					
	
				

					<table border="0" width="100%" id="product-table" class='xthf-site'>
					<!--#include file="FormSelectDailyReportHEADERMill.html"-->
					
					
						<tbody>
						
						
							<!-- Start COUNT ROWSPAN Rendemen -->
							<!--#include file="FormCountRowRendemen.asp"-->																										
							<!-- Start Rendemen -->		
							<!--#include file="FormDisplayDailyReportRendemen.html"-->						
						
							

							
							
							<tr>
								<th class="separator" colspan="<%= 41 + nColSpan %>" align="left"></th>				
							</tr>

														
							
						</tbody>
					</table>
				

			</form>
			<!--#include file="indicator_color.asp"-->

		</div>
	</body>
</html>
