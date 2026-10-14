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
	
	v_txtReportTitle = "LAPORAN HARIAN TBS PROSES" 	
	
	
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
					" sum(decode(to_number(to_char(ondate,'DD')),2, todate, 0 )) h_02, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),3, todate, 0 )) h_03, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),4, todate, 0 )) h_04, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),5, todate, 0 )) h_05, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),6, todate, 0 )) h_06, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),7, todate, 0 )) h_07, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),8, todate, 0 )) h_08, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),9, todate, 0 )) h_09, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),10, todate, 0 )) h_10, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),11, todate, 0 )) h_11, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),12, todate, 0 )) h_12, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),13, todate, 0 )) h_13, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),14, todate, 0 )) h_14, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),15, todate, 0 )) h_15, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),16, todate, 0 )) h_16, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),17, todate, 0 )) h_17, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),18, todate, 0 )) h_18, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),19, todate, 0 )) h_19, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),20, todate, 0 )) h_20, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),21, todate, 0 )) h_21, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),22, todate, 0 )) h_22, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),23, todate, 0 )) h_23, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),24, todate, 0 )) h_24, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),25, todate, 0 )) h_25, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),26, todate, 0 )) h_26, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),27, todate, 0 )) h_27, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),28, todate, 0 )) h_28, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),29, todate, 0 )) h_29, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),30, todate, 0 )) h_30, " & _
					" sum(decode(to_number(to_char(ondate,'DD')),31, todate, 0 )) h_31, " & _
					" sum(decode(to_number(to_char(ondate,'DD'))," & txtDay & ", tomonth, 0 )) shi, " & _
					" sum(decode(to_number(to_char(ondate,'DD'))," & txtDay & ", toyear, 0 )) sbi, " & _
					" sum(decode(to_number(to_char(ondate,'DD'))," & txtDay & ", budgettomonth, 0 )) budget_shi, " & _
					" sum(decode(to_number(to_char(ondate,'DD'))," & txtDay & ", budgettoyear, 0 )) budget_sbi   "					

	
				
	strMAINAVGStmt = "" 
				
	
	
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
					" shi,  " & _
					" sbi,  " & _
					" budget_shi,  " & _
					" budget_sbi   "

    strMAIN2AVGStmt =   "" 					
					
					

				
					
					
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



	' Step 03
	strSTEPStmt = " stepnumber = 'STEP_03' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsProses = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsProses.Open strSQL, connstr, 3,3	
	
	'*************************SPECIAL FOR TBS BRUTO = TBS NETTO + SORTASI ************************
	
	'Step SORTASI UTAMA
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'FFB07' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) "

	strNETTOstmt = " , sum(h_01) h_01" & _
							" , sum(h_02) h_02" & _
							" , sum(h_03) h_03" & _
							" , sum(h_04) h_04" & _
							" , sum(h_05) h_05" & _
							" , sum(h_06) h_06" & _
							" , sum(h_07) h_07" & _
							" , sum(h_08) h_08" & _
							" , sum(h_09) h_09" & _
							" , sum(h_10) h_10" & _
							" , sum(h_11) h_11" & _
							" , sum(h_12) h_12" & _
							" , sum(h_13) h_13" & _
							" , sum(h_14) h_14" & _
							" , sum(h_15) h_15" & _
							" , sum(h_16) h_16" & _
							" , sum(h_17) h_17" & _
							" , sum(h_18) h_18" & _
							" , sum(h_19) h_19" & _
							" , sum(h_20) h_20" & _
							" , sum(h_21) h_21" & _
							" , sum(h_22) h_22" & _
							" , sum(h_23) h_23" & _
							" , sum(h_24) h_24" & _
							" , sum(h_25) h_25" & _
							" , sum(h_26) h_26" & _
							" , sum(h_27) h_27" & _
							" , sum(h_28) h_28" & _
							" , sum(h_29) h_29" & _
							" , sum(h_30) h_30" & _
							" , sum(h_31) h_31" & _
							" , sum(shi) shi " & _
							" , sum(sbi) sbi " & _
							" , sum(budget_shi) budget_shi  " & _
							" , sum(budget_sbi) budget_sbi   "
	
	
	
	
	strSortasi = "SELECT 'TBS_BRUTO' title " & strSELECTStmt_Rendemen & strNETTOstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'TBS_BRUTO' " & strGROUPRENDEMENStmt
					

	'=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_03' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
	
	strNETTOstmt =   " , sum(h_01) h_01" & _
							" , sum(h_02) h_02" & _
							" , sum(h_03) h_03" & _
							" , sum(h_04) h_04" & _
							" , sum(h_05) h_05" & _
							" , sum(h_06) h_06" & _
							" , sum(h_07) h_07" & _
							" , sum(h_08) h_08" & _
							" , sum(h_09) h_09" & _
							" , sum(h_10) h_10" & _
							" , sum(h_11) h_11" & _
							" , sum(h_12) h_12" & _
							" , sum(h_13) h_13" & _
							" , sum(h_14) h_14" & _
							" , sum(h_15) h_15" & _
							" , sum(h_16) h_16" & _
							" , sum(h_17) h_17" & _
							" , sum(h_18) h_18" & _
							" , sum(h_19) h_19" & _
							" , sum(h_20) h_20" & _
							" , sum(h_21) h_21" & _
							" , sum(h_22) h_22" & _
							" , sum(h_23) h_23" & _
							" , sum(h_24) h_24" & _
							" , sum(h_25) h_25" & _
							" , sum(h_26) h_26" & _
							" , sum(h_27) h_27" & _
							" , sum(h_28) h_28" & _
							" , sum(h_29) h_29" & _
							" , sum(h_30) h_30" & _
							" , sum(h_31) h_31" & _
							" , sum(shi) shi " & _
							" , sum(sbi) sbi " & _
							" , sum(budget_shi) budget_shi  " & _
							" , sum(budget_sbi) budget_sbi   "
	
	strSortasi = strSortasi & " UNION SELECT 'TBS_BRUTO' title " & strSELECTStmt_Rendemen & strNETTOstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'TBS_BRUTO' " & strGROUPRENDEMENStmt
	
	strNETTOFinalStmt = " , sum(h_01) h_01" & _
							" , sum(h_02) h_02" & _
							" , sum(h_03) h_03" & _
							" , sum(h_04) h_04" & _
							" , sum(h_05) h_05" & _
							" , sum(h_06) h_06" & _
							" , sum(h_07) h_07" & _
							" , sum(h_08) h_08" & _
							" , sum(h_09) h_09" & _
							" , sum(h_10) h_10" & _
							" , sum(h_11) h_11" & _
							" , sum(h_12) h_12" & _
							" , sum(h_13) h_13" & _
							" , sum(h_14) h_14" & _
							" , sum(h_15) h_15" & _
							" , sum(h_16) h_16" & _
							" , sum(h_17) h_17" & _
							" , sum(h_18) h_18" & _
							" , sum(h_19) h_19" & _
							" , sum(h_20) h_20" & _
							" , sum(h_21) h_21" & _
							" , sum(h_22) h_22" & _
							" , sum(h_23) h_23" & _
							" , sum(h_24) h_24" & _
							" , sum(h_25) h_25" & _
							" , sum(h_26) h_26" & _
							" , sum(h_27) h_27" & _
							" , sum(h_28) h_28" & _
							" , sum(h_29) h_29" & _
							" , sum(h_30) h_30" & _
							" , sum(h_31) h_31" & _
							" , sum(shi) shi " & _
							" , sum(sbi) sbi " & _
							" , sum(budget_shi) budget_shi  " & _
							" , sum(budget_sbi) budget_sbi   "

						   
	strNettoFinal = "SELECT title HEADGROUP, 'Kg' UnitOfMeasurement, '' CodeDesc " & strSELECTStmt_Rendemen & strNETTOFinalStmt & " FROM ( " & strSortasi & " ) " & _
	                   " GROUP BY title " & strGROUPRENDEMENStmt
	
	

	Set rsNettoFinal = Server.CreateObject("ADODB.Recordset")
	'response.write strNettoFinal
	rsNettoFinal.Open strNettoFinal, connstr, 3,3			

	
    %>


<html>

	<head>
		<title>iPla'S Reporting - Daily Production TBS Process Report</title>
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
			<form name="frmEntry" method="post" action="p_DailyReport_Mill_FFBNettoBruto.asp">
				<!--#include file="menutop.asp"-->	
				<!--#include file="FormSelectNewSum.html"-->
				<!--#include file="FormSelectDailyReportMill_FFBNettoBruto.html"-->					
	
				

					<table border="0" width="100%" id="product-table" class='xthf-site'>
					<!--#include file="FormSelectDailyReportHEADERMill.html"-->
					
					
						<tbody>
						
							
							<!-- Start COUNT ROWSPAN FFBProductionProcess -->
							<!--#include file="FormCountRowDailyReportFFBProductionProcess.asp"-->
							
							<!-- Start COUNT ROWSPAN FFBBrutoProductionProcess -->
							<!--#include file="FormCountRowDailyReportFFBBrutoProductionProcess.asp"-->
							
							<!-- Start PROSES -->
							<!--#include file="FormDisplayDailyReportFFBProductionProcess.html"-->
							
							<!-- Start PROSES -->
							<!--#include file="FormDisplayDailyReportFFBBrutoProductionProcess.html"-->
							
							
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
