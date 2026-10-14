<%@Language=VBScript%>

<%
' Option Explicit 
%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	' DateTo and DateFrom will be EQUAL to remove miss-understanding
	
	Dim txtToDayDate, txtyear, txtmonth
	Dim strSQLPreParameter, rsSQLPreParameter, v_txtToDayDate, v_txtyear, v_txtmonth
	Dim strProduction, rsProduction
	Dim tempRefCode, strHRef, rsHRef, tempDescription, tempUOM
	Dim tempGrading_HI, tempGrading_SHI, tempGrading_SBI
	Dim tempDecrease_hi, tempDecrease_shi, tempDecrease_sbi
	Dim tempOpeningCPOStock_HI, tempDespatchCPOStock_HI
	Dim tempOpeningKernelStock_HI, tempDespatchKernelStock_HI
	Dim tempRounding
    Dim rscomp , strcomp,txcompanyname,str1,txtComp,v_txtComp,nColSpan
	
	
	
    txtGroupPlantation  = Request("txtGroupPlantation")
	txtPlantation  		= Request("txtPlantation")
	txtGroupRegion  	= Request("txtGroupRegion")
	txtRegion  			= Request("txtRegion")
	txtcompany  		= Request("txtCompany")
	txtfcba 			= Request("txtfcba")
	txtToDayDate		= Request("txtdatefrom")
	
	' Now we make the HardCode...
	txtBudgetVersion = 1
	txtModuleVersion = 1
	
	
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
	       "to_char(to_date('" & txtToDayDate & "', 'MM/DD/YYYY'),'MM') as CurrentPeriod from dual "
	
	Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
	rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 
	
    if not rsSQLPreParameter.EOF then
		txtyear = rsSQLPreParameter("CurrentYear")
		txtmonth = rsSQLPreParameter("CurrentPeriod")
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
					" sum(todate) hi, " & _
					" sum(tomonth) shi, " & _
					" sum(toyear) sbi, " & _
					" sum(budgettodate) budget_hi,  " & _
					" sum(budgettomonth) budget_shi,  " & _
					" sum(budgettoyear) budget_sbi   " 
	
    strMAIN2Stmt =   " , seqnumber, " & _
					" hi, " & _
					" shi, " & _
					" sbi, " & _
					" budget_hi,  " & _
					" budget_shi,  " & _
					" budget_sbi   " 					


				
					
					
	strFROMStmt = " from mill_summary_tab  " & _
					" where  " & _
					" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) is not null   and "  & _
					" to_char(ondate,'YYYYMMDD') = to_char(TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY'),'YYYYMMDD')  and " 
					
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
	
	strORDERStmt = " ORDER BY groupplantationcode " 
	
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


	' Step 01 
	strSTEPStmt = " stepnumber = 'STEP_01' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsTerimaTBS = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsTerimaTBS.Open strSQL, connstr, 3,3	


	' Step 02 
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsSortasi = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsSortasi.Open strSQL, connstr, 3,3	
	

	' Step 03
	strSTEPStmt = " stepnumber = 'STEP_03' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsProses = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsProses.Open strSQL, connstr, 3,3			
	
	' Step 04 
	strSTEPStmt = " stepnumber = 'STEP_04' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsKapasitasOlah = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsKapasitasOlah.Open strSQL, connstr, 3,3	

	' Step 05 
	strSTEPStmt = " stepnumber = 'STEP_05' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsSisaTBS = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsSisaTBS.Open strSQL, connstr, 3,3	
	
	
    ' Step 051 -- Closing Stock 
	strSTEPStmt = " stepnumber = 'STEP_051' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsClosingTBS = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsClosingTBS.Open strSQL, connstr, 3,3	

	' Step 06
	strSTEPStmt = " stepnumber = 'STEP_06' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsOilKernelProduction = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsOilKernelProduction.Open strSQL, connstr, 3,3	
	
	
	
	' Step 07
	strSTEPStmt = " stepnumber = 'STEP_07' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsCPOStock = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsCPOStock.Open strSQL, connstr, 3,3		
	
	' Step 08
	strSTEPStmt = " stepnumber = 'STEP_08' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsKernelStock = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsKernelStock.Open strSQL, connstr, 3,3	
	
	
	' Step 09
	strSTEPStmt = " stepnumber = 'STEP_09' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsProductionGrading = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsProductionGrading.Open strSQL, connstr, 3,3	

	
	' Step 10
	strSTEPStmt = " stepnumber = 'STEP_10' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strSQL = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 
 	Set rsOilProduction = Server.CreateObject("ADODB.Recordset")
	'response.write strSQL
	rsOilProduction.Open strSQL, connstr, 3,3		
	
	
	' *************************SPECIAL FOR SORTASI PERCENTAGE************************
	
	' Step SORTASI UTAMA
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'FFB07' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) " 

	
	strSORTASIstmt =   ", sum(hi) sortasi_hi, " & _
				" sum(shi) sortasi_shi, " & _
				" sum(sbi) sortasi_sbi, " & _
				" sum(budget_hi) sortasi_budget_hi,  " & _
				" sum(budget_shi) sortasi_budget_shi,  " & _
				" sum(budget_sbi) sortasi_budget_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 tbs_received_hi, " & _
				" 0 tbs_received_shi, " & _
				" 0 tbs_received_sbi, " & _
				" 0 tbs_received_budget_hi,  " & _
				" 0 tbs_received_budget_shi,  " & _
				" 0 tbs_received_budget_sbi   " 	



	strSortasi = "SELECT 'SORTASI_PERCENTAGE' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'SORTASI_PERCENTAGE' " & strGROUPRENDEMENStmt
	
	'=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_01' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'FFB02' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strSORTASIstmt =   ", 0 sortasi_hi, " & _
				" 0 sortasi_shi, " & _
				" 0 sortasi_sbi, " & _
				" 0 sortasi_budget_hi,  " & _
				" 0 sortasi_budget_shi,  " & _
				" 0 sortasi_budget_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", sum(hi) tbs_received_hi, " & _
				" sum(shi) tbs_received_shi, " & _
				" sum(sbi) tbs_received_sbi, " & _
				" sum(budget_hi) tbs_received_budget_hi,  " & _
				" sum(budget_shi) tbs_received_budget_shi,  " & _
				" sum(budget_sbi) tbs_received_budget_sbi   " 	
	

	strSortasi = strSortasi & " UNION SELECT 'SORTASI_PERCENTAGE' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'SORTASI_PERCENTAGE' " & strGROUPRENDEMENStmt
	
	
	
	strSORTASIFinalStmt = " , sum(sortasi_hi) sortasi_hi , " & _
						   " sum(sortasi_shi) sortasi_shi , " & _
						   " sum(sortasi_sbi) sortasi_sbi , " & _
						   " sum(sortasi_budget_hi) sortasi_budget_hi , " & _
						   " sum(sortasi_budget_shi) sortasi_budget_shi , " & _
						   " sum(sortasi_budget_sbi) sortasi_budget_sbi , " & _
						   " sum(tbs_received_hi) tbs_received_hi , " & _
						   " sum(tbs_received_shi) tbs_received_shi , " & _
						   " sum(tbs_received_sbi) tbs_received_sbi , " & _
						   " sum(tbs_received_budget_hi) tbs_received_budget_hi , " & _
						   " sum(tbs_received_budget_shi) tbs_received_budget_shi , " & _
						   " sum(tbs_received_budget_sbi) tbs_received_budget_sbi  " 
						   
	strSortasiFinal = "SELECT title HEADGROUP, '%' UnitOfMeasurement, '' CodeDesc " & strSELECTStmt_Rendemen & strSORTASIFinalStmt & " FROM ( " & strSortasi & " ) " & _
	                   " GROUP BY title " & strGROUPRENDEMENStmt
	
	

	Set rsSortasiFinal = Server.CreateObject("ADODB.Recordset")
	'response.write strSortasiFinal
	rsSortasiFinal.Open strSortasiFinal, connstr, 3,3		
	
	
	
	' *************************SPECIAL FOR RENDEMEN************************
	

	' Step RENDEMEN UTAMA
	strSTEPStmt = " stepnumber = 'STEP_03' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'FFB0701' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) " 

	
	strRENDEMENstmt =   ", sum(hi) tbs_process_hi, " & _
				" sum(shi) tbs_process_shi, " & _
				" sum(sbi) tbs_process_sbi, " & _
				" sum(budget_hi) tbs_process_budget_hi,  " & _
				" sum(budget_shi) tbs_process_budget_shi,  " & _
				" sum(budget_sbi) tbs_process_budget_sbi   " 	
	strRENDEMENstmt = strRENDEMENstmt &  ", 0 cpo_production_hi, " & _
				" 0 cpo_production_shi, " & _
				" 0 cpo_production_sbi, " & _
				" 0 cpo_production_budget_hi,  " & _
				" 0 cpo_production_budget_shi,  " & _
				" 0 cpo_production_budget_sbi   " 	
	strRENDEMENstmt =  strRENDEMENstmt & ", 0 kernel_production_hi, " & _
				" 0 kernel_production_shi, " & _
				" 0 kernel_production_sbi, " & _
				" 0 kernel_production_budget_hi,  " & _
				" 0 kernel_production_budget_shi,  " & _
				" 0 kernel_production_budget_sbi   " 	


	strRendemen = "SELECT 'RENDEMEN' title " & strSELECTStmt_Rendemen & strRENDEMENstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN_DETAIL' " & strGROUPRENDEMENStmt
	
	'======================CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_06' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'PROD01' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strRENDEMENstmt =   ", 0 tbs_process_hi, " & _
				" 0 tbs_process_shi, " & _
				" 0 tbs_process_sbi, " & _
				" 0 tbs_process_budget_hi,  " & _
				" 0 tbs_process_budget_shi,  " & _
				" 0 tbs_process_budget_sbi   " 	
	strRENDEMENstmt = strRENDEMENstmt &  ", sum(hi) cpo_production_hi, " & _
				" sum(shi) cpo_production_shi, " & _
				" sum(sbi) cpo_production_sbi, " & _
				" sum(budget_hi) cpo_production_budget_hi,  " & _
				" sum(budget_shi) cpo_production_budget_shi,  " & _
				" sum(budget_sbi) cpo_production_budget_sbi   " 	
	strRENDEMENstmt =  strRENDEMENstmt & ", 0 kernel_production_hi, " & _
				" 0 kernel_production_shi, " & _
				" 0 kernel_production_sbi, " & _
				" 0 kernel_production_budget_hi,  " & _
				" 0 kernel_production_budget_shi,  " & _
				" 0 kernel_production_budget_sbi   " 	

	strRendemen = strRendemen & " UNION SELECT 'RENDEMEN' title " & strSELECTStmt_Rendemen & strRENDEMENstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN' " & strGROUPRENDEMENStmt
	
	'======================CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_06' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'PROD02' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strRENDEMENstmt =   ", 0 tbs_process_hi, " & _
				" 0 tbs_process_shi, " & _
				" 0 tbs_process_sbi, " & _
				" 0 tbs_process_budget_hi,  " & _
				" 0 tbs_process_budget_shi,  " & _
				" 0 tbs_process_budget_sbi   " 	
	strRENDEMENstmt = strRENDEMENstmt &  ", 0 cpo_production_hi, " & _
				" 0 cpo_production_shi, " & _
				" 0 cpo_production_sbi, " & _
				" 0 cpo_production_budget_hi,  " & _
				" 0 cpo_production_budget_shi,  " & _
				" 0 cpo_production_budget_sbi   " 	
	strRENDEMENstmt =  strRENDEMENstmt & ", sum(hi) kernel_production_hi, " & _
				" sum(shi) kernel_production_shi, " & _
				" sum(sbi) kernel_production_sbi, " & _
				" sum(budget_hi) kernel_production_budget_hi,  " & _
				" sum(budget_shi) kernel_production_budget_shi,  " & _
				" sum(budget_sbi) kernel_production_budget_sbi   " 	

	strRendemen = strRendemen & " UNION SELECT 'RENDEMEN' title " & strSELECTStmt_Rendemen & strRENDEMENstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN' " & strGROUPRENDEMENStmt
	
	
	
	strRENDEMENFinalStmt = " , sum(tbs_process_hi) tbs_process_hi , " & _
						   " sum(tbs_process_shi) tbs_process_shi , " & _
						   " sum(tbs_process_sbi) tbs_process_sbi , " & _
						   " sum(tbs_process_budget_hi) tbs_process_budget_hi , " & _
						   " sum(tbs_process_budget_shi) tbs_process_budget_shi , " & _
						   " sum(tbs_process_budget_sbi) tbs_process_budget_sbi , " & _
						   " sum(cpo_production_hi) cpo_production_hi , " & _
						   " sum(cpo_production_shi) cpo_production_shi , " & _
						   " sum(cpo_production_sbi) cpo_production_sbi , " & _
						   " sum(cpo_production_budget_hi) cpo_production_budget_hi , " & _
						   " sum(cpo_production_budget_shi) cpo_production_budget_shi , " & _
						   " sum(cpo_production_budget_sbi) cpo_production_budget_sbi , " & _
						   " sum(kernel_production_hi) kernel_production_hi , " & _
						   " sum(kernel_production_shi) kernel_production_shi , " & _
						   " sum(kernel_production_sbi) kernel_production_sbi , " & _
						   " sum(kernel_production_budget_hi) kernel_production_budget_hi , " & _
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
		<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
		<title>iPla'S Reporting - Daily Production Statement Report</title>

		<script type='text/javascript' src='/js/x.js'></script>
		<script type='text/javascript' src='/js/xtableheaderfixed.js'></script>
		<script type='text/javascript'>
			xAddEventListener(window, 'load',
			  function() {
				new xTableHeaderFixed('xthf-site','tbl-container',0);
				new xTableHeaderFixed('xthf-green', window);
				new xTableHeaderFixed('xthf-blue', window);
			  }, false
			);
		</script>

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
	</head>


	<body>
		<div id="wrapper">
			<form name="frmEntry" method="post" action="p_Summary_Production_NAF_2.asp">
				<!--#include file="menutop.asp"-->	
				<!--#include file="FormSelectNewSum.html"-->
				<!--#include file="FormSelectSUMMARYMill.html"-->					
	
				<div id="tbl-container">

					<table border="0" width="100%" id="product-table" class='xthf-site'>
					<!--#include file="FormSelectHEADERMill.html"-->
					
					
						<tbody>
							<!-- Start Terima TBS -->
							<!--#include file="FormDisplayFFBReceived.html"-->
							
							<tr>
								<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
							</tr>	
							
							<!-- Start SORTASI -->	
							<!--#include file="FormDisplayFFBGradingSortage.html"-->
							

							
							<!-- Start PERCENTAGE SORTASI -->	
							<!--#include file="FormDisplayFFBPercentageSortage.html"-->
							

							<tr>
								<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
							</tr>	
							
							
							
							
							
							<!-- Start PROSES -->	
							<!--#include file="FormDisplayFFBProductionProcess.html"-->


							<tr>
								<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
							</tr>	
							
							

							<!-- Start KAPASITAS OLAH -->		
							<!--#include file="FormDisplayMillCapacity.html"-->
							
							<tr>
								<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
							</tr>	

							<!-- Start Sisa TBS -->		
							<!--#include file="FormDisplayFFBRemainMill.html"-->
							
							
							<!-- Start Opening Stock -->
							<!--#include file="FormDisplayFFBClosingStock.html"-->
							
							<tr>
								<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
							</tr>	
							
							

							<!-- Start OIL Kernel Production -->		
							<!--#include file="FormDisplayOilKernelProduction.html"-->
							
							<tr>
								<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
							</tr>	
							
							<!-- Start Rendemen -->		
							<!--#include file="FormDisplayRendemen.html"-->

							<tr>
								<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
							</tr>	

							<!-- Start CPO Stock -->		
							<!--#include file="FormDisplayCPOStock.html"-->							

							<tr>
								<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
							</tr>	

							<!-- Start KERNEL Stock -->		
							<!--#include file="FormDisplayKernelStock.html"-->							
							
							
						</tbody>
					</table>
				</div>

			</form>
		</div>
	</body>
</html>
