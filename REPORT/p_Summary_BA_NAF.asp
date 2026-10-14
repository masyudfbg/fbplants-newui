<%@Language=VBScript%>
<% Option Explicit %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	
	Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion
	Dim txtBudgetYear, txtBudgetVersion, txtModuleVersion
    Dim strSELECTStmt, strMAINStmt, strWHEREStmt, strGROUPStmt, strORDERStmt,  strFROMStmt, strMAIN2Stmt, strSTEPStmt, strMAINMAXStmt
	Dim strTempLandClearing, strTempProduction, nColSpan
	Dim strGROUPRENDEMENStmt, strTemp, strSORTASIstmt, strSELECTStmt_Rendemen, strSortasi, strSORTASIFinalStmt, strSortasiFinal
	Dim rsBiayaMengumpulFinal, rsTBSPengangkutanFinal, strTBSPengangkutan
	Dim strOutputPanen, rsOutputPanenFinal
	Dim rsPusingan, rsCurahHujan
	
	Dim txtDetailInfo, v_txtDetailInfo, tempUnitOfMeasurement
	Dim strRatio, rsRatioFinal
	Dim strRatioAmount, rsRatioAmountFinal
	Dim finalRowSpan, strCodeDescHead, rsCodeDescHead, tempCodeHeadDescription, tempSelectedDate
	Dim tempHI, tempSHI, tempSBI
	Dim temp_HI, temp_SHI, temp_SBI
	
	
	Dim strCompare
	Dim rsCompare
	Dim tempStrCompare

	dim temporiginalscriptname, tempfilternextreport
	
	
	Dim tempGroupPlantation, sPrevGroupPlantation
	Dim tempPlantation, sPrevPlantation
	Dim tempGroupRegion, sPrevGroupRegion
	Dim tempRegion, sPrevRegion
	Dim tempCompany, sPrevCompany
	Dim tempFcba, sPrevFcba, sPrevSubDesc
	Dim tempCodeSubDescription
	
	Dim tempNumber

	
	Dim v_txtdatefrom, v_txtdateto,  v_txtyear, v_txtmonth,v_txtComp, v_txtGroupItem
	Dim txtdatefrom, txtdateto,  txtyear, txtmonth , txtComp
	Dim strProduction, strHandling, strTempHandling, strHandlingAmount, strTempHandlingAmount, strFertilizer, strTempFertilizer
	Dim strEndingBalance, strTempEndingBalance,  strMandays, strTempMandays
	Dim strTempVehicleGroup, strVehicleGroup
	Dim rsProduction, rsHandling, rsHandlingAmount, rsFertilizer, rsEndingBalance, rsMandays, rsVehicleGroup
	Dim no, cClass, tempData
	Dim tempRefCode, strHRef, rsHRef, tempReference, tempExtendedReference, tempRowSpan, tempExtendedReferenceAll, tempExtendedReferenceDetail
	Dim strSQLPreParameter, rsSQLPreParameter
	Dim strCodeDesc, rsCodeDesc, tempCodeDescription
	Dim sPrevDesc
	Dim tempRounding
    Dim tempBudget_HI, tempBudget_SHI, tempBudget_SBI, tempBudgetType
	Dim strBudget, rsBudget 
					
	
	Dim tempUOM, strOutput, strTempOutput, rsOutput
	
	Dim tempDataBudget_HI, tempDataBudget_SHI, tempDataBudget_SBI
	Dim tempCompare_HI, tempCompare_SHI, tempCompare_SBI
	Dim cClassAssigned_HI, cClassAssigned_SHI, cClassAssigned_SBI
	Dim strLandClearing, rsLandClearing,rsCPO_FFAQuality

	Dim tempTBSProduksi, tempTBSProduksi_shi, tempTBSProduksi_sbi
	Dim tempBiayaAngkut, tempBiayaAngkut_shi, tempBiayaAngkut_sbi
	Dim tempData_shi, tempData_sbi 	
	Dim rscomp , strcomp, txcompanyname,str1
	Dim strGetColor, rsGetColor
	Dim txtGroupItem
	
    
			




	' DateTo and DateFrom will be EQUAL to remove miss-understanding
	
	
	txtGroupPlantation  = Request("txtGroupPlantation")
	txtPlantation  		= Request("txtPlantation")
	txtGroupRegion  	= Request("txtGroupRegion")
	txtRegion  			= Request("txtRegion")
	txtcompany  		= Request("txtCompany")
	txtfcba 			= Request("txtfcba")
	
	
	txtdatefrom			= Request("txtdatefrom")
	txtdateto			= Request("txtdatefrom")
    txtyear 			= Request("txtyear")
	txtmonth 			= Request("txtmonth")
	
	txtGroupItem  		= Request("txtGroupItem")
	txtBudgetYear 		= Request("txtBudgetYear")
	txtBudgetVersion	= Request("txtBudgetVersion")
	txtModuleVersion	= Request("txtModuleVersion")
	
	txtDetailInfo		= Request("txtDetailInfo")
	
	if len(trim(txtdatefrom)) = 0 then
	   txtdatefrom	= date -1
	   txtdateto	= date -1
	end if
	
    if len(trim(txtdatefrom)) > 0 then  
		strSQLPreParameter = "select to_char(to_date('" & txtdatefrom & "', 'MM/DD/YYYY'),'YYYY') as CurrentYear, " & _
			   "to_char(to_date('" & txtdatefrom & "', 'MM/DD/YYYY'),'MM') as CurrentPeriod from dual "
		
		Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
		rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 
		
		if not rsSQLPreParameter.EOF then
			txtyear = rsSQLPreParameter("CurrentYear")
			txtmonth = rsSQLPreParameter("CurrentPeriod")
		end if
		rsSQLPreParameter.Close
		set rsSQLPreParameter = nothing
	else
		strSQLPreParameter = "select to_char(sysdate,'YYYY') as CurrentYear, " & _
			   "to_char(sysdate,'MM') as CurrentPeriod from dual "
		
		Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
		rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 
		
		if not rsSQLPreParameter.EOF then
			txtyear = rsSQLPreParameter("CurrentYear")
			txtmonth = rsSQLPreParameter("CurrentPeriod")
			txtBudgetYear = rsSQLPreParameter("CurrentYear")
			txtBudgetVersion = 1
			txtModuleVersion = 1
		end if
		rsSQLPreParameter.Close
		set rsSQLPreParameter = nothing
	
	end if
	
     
	'response.write  txtcompany
	
	

	
	
	v_txtGroupPlantation 	= txtGroupPlantation
	v_txtPlantation 		= txtPlantation
	v_txtGroupRegion 		= txtGroupRegion
	v_txtRegion 			= txtRegion
	v_txtCompany			=   txtCompany
	v_txtfcba     			=   txtfcba
	v_txtGroupItem     	    =   txtGroupItem
	
	v_txtdatefrom			=	txtdatefrom
	v_txtdateto				=	txtdateto
	v_txtyear 				=   txtyear
	v_txtmonth				=   txtmonth
	v_txtBudgetYear 		= 	txtBudgetYear
	v_txtBudgetVersion 		= 	txtBudgetVersion
	v_txtModuleVersion 		= 	txtModuleVersion
	v_txtDetailInfo			= 	txtDetailInfo
	
	
	' Zero 
	' Land Clearing
	
    
	
	strSELECTStmt = " codedesc, jobgencode, description, unitofmeasurement "
	strSELECTStmt_Rendemen = ""
	
	
    if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , groupplantationcode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , groupplantationcode "
		end if
	'else
	'    strSELECTStmt = strSELECTStmt & " , groupplantationcode "
	'	strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , groupplantationcode "
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , plantationcode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , plantationcode "
		end if
	'else
	'    strSELECTStmt = strSELECTStmt & " , plantationcode "
	'	strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , groupregioncode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , groupregioncode "
		end if
	'else
	'    strSELECTStmt = strSELECTStmt & " , groupregioncode "
	'	strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , regioncode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , regioncode "
		end if
	'else
	'    strSELECTStmt = strSELECTStmt & " , regioncode "
	'	strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , companycode "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , companycode "
		end if
	'else
	 '   strSELECTStmt = strSELECTStmt & " , companycode "
	'	strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strSELECTStmt = strSELECTStmt & " , fcba "
			strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , fcba "
		end if
	'else
	'    strSELECTStmt = strSELECTStmt & " , fcba "
	'	strSELECTStmt_Rendemen = strSELECTStmt_Rendemen & " , fcba "
	end if	
	
	
	
	strMAINStmt =   " , sum(nvl(seqnumber,0)) seqnumber, " & _
					" sum(todate) hi, " & _
					" sum(tomonth) shi, " & _
					" sum(toyear) sbi, " & _
					" sum(budgettodate) budget_hi,  " & _
					" sum(budgettomonth) budget_shi,  " & _
					" sum(budgettoyear) budget_sbi   " 
					
	strMAINMaxStmt =   " , sum(nvl(seqnumber,0)) seqnumber, " & _
					" max(todate) hi, " & _
					" max(tomonth) shi, " & _
					" max(toyear) sbi, " & _
					" max(budgettodate) budget_hi,  " & _
					" max(budgettomonth) budget_shi,  " & _
					" max(budgettoyear) budget_sbi   " 					

    strMAIN2Stmt =   " , seqnumber, " & _
					" hi, " & _
					" shi, " & _
					" sbi, " & _
					" budget_hi,  " & _
					" budget_shi,  " & _
					" budget_sbi   " 					
					
	strFROMStmt = " from production_summary_tab  " & _
					" where  " & _
					" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) is not null   and "  & _
					" to_char(ondate,'YYYYMMDD') = to_char(TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'),'YYYYMMDD')  and " 
					
					
										
										


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
	'else
	'    strWHEREStmt = strWHEREStmt & " AND GroupPlantationCode like '%' " 
	end if
	
	
	if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND PlantationCode like '%' " 
		else
		    if v_txtPlantation <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND PlantationCode = '" & v_txtPlantation& "' " 
			end if
		end if
	'else
	'    strWHEREStmt = strWHEREStmt & " AND PlantationCode like '%' " 
	end if
	
	
	if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND GroupRegionCode like '%' " 
		else
		    if v_txtGroupRegion <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND GroupRegionCode = '" & v_txtGroupRegion& "' " 
			end if
		end if
	'else
	'    strWHEREStmt = strWHEREStmt & " AND GroupRegionCode like '%' " 
	end if
	

	if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND RegionCode like '%' " 
		else
		    if v_txtRegion <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND RegionCode = '" & v_txtRegion& "' " 
			end if
		end if
	'else
	'    strWHEREStmt = strWHEREStmt & " AND RegionCode like '%' " 
	end if
	
	if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND CompanyCode like '%' " 
		else
		    if v_txtcompany <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND CompanyCode = '" & v_txtcompany& "' " 
			end if
		end if
	'else
	'    strWHEREStmt = strWHEREStmt & " AND CompanyCode like '%' " 
	end if	


	if len(trim(v_txtfcba)) > 0 then 
		if v_txtfcba = "--ALL--" then
	        strWHEREStmt = strWHEREStmt & " AND fcba like '%' " 
		else
		    if v_txtfcba <> "--SUMMARY--" then 
	            strWHEREStmt = strWHEREStmt & " AND fcba = '" & v_txtfcba& "' " 
			end if
		end if
	'else
	'    strWHEREStmt = strWHEREStmt & " AND fcba like '%' " 
	end if	

	
	

	' Group By
	
	strGROUPStmt = ""
	
	strGROUPStmt = strGROUPStmt & " GROUP BY codedesc, jobgencode, description, unitofmeasurement "
	
    if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , groupplantationcode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , groupplantationcode "			
		end if
	'else
	'    strGROUPStmt = strGROUPStmt & " , groupplantationcode "
	'	strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , groupplantationcode "
		
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , plantationcode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , plantationcode "
		end if
	'else
	'    strGROUPStmt = strGROUPStmt & " , plantationcode "
	'	strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , groupregioncode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , groupregioncode "
		end if
	'else
	'    strGROUPStmt = strGROUPStmt & " , groupregioncode "
	'	strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , regioncode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , regioncode "
		end if
	'else
	'    strGROUPStmt = strGROUPStmt & " , regioncode "
	'	strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , companycode "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , companycode "
		end if
	'else
	'    strGROUPStmt = strGROUPStmt & " , companycode "
	'	strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strGROUPStmt = strGROUPStmt & " , fcba "
			strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , fcba "
		end if
	'else
	'    strGROUPStmt = strGROUPStmt & " , fcba "
	'	strGROUPRENDEMENStmt = strGROUPRENDEMENStmt &  " , fcba "
	end if	

	
	
	
	' Order By
	
	'strORDERStmt = " ORDER BY 'x' " 
	strORDERStmt = " ORDER BY 'x', seqnumber " 
	
	
    if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , groupplantationcode "
		end if
	'else
	'    strORDERStmt = strORDERStmt & " , groupplantationcode "
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , plantationcode "
		end if
	'else
	'    strORDERStmt = strORDERStmt & " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , groupregioncode "
		end if
	'else
	'    strORDERStmt = strORDERStmt & " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , regioncode "
		end if
	'else
	'    strORDERStmt = strORDERStmt & " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , companycode "
		end if
	'else
	'    strORDERStmt = strORDERStmt & " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strORDERStmt = strORDERStmt & " , fcba "
		end if
	'else
	'    strORDERStmt = strORDERStmt & " , fcba "
	end if		
	
    ' strORDERStmt = strORDERStmt & " , seqnumber "
    strORDERStmt = strORDERStmt & " "
	
	
	' Step 01 -- Land Clearing
	strSTEPStmt = " stepnumber = 'STEP_01' " 	
	strTempLandClearing = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strLandClearing = "SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempLandClearing & strORDERStmt & " ) " 
 	Set rsLandClearing = Server.CreateObject("ADODB.Recordset")
	'response.write strLandClearing
	
	rsLandClearing.Open strLandClearing, connstr, 3,3


	' Step 02 -- Production
	' SR003_BA - Biaya Pengangkutan TBS
	' SR003_05_AMT Amount biaya Mengumpul
	' SR003_05_TON Kg Mengumpul
	' SR003_OP ==> HK atau Mandays dari Potong Buah, Output Panen
	' SR003_04 Curah Hujan
	' 'SR003_BA_CURR_QTY','SR003_BA_PREV_AMT', 'SR003_BA_PREV_QTY' ==> for Biaya Pengangkutan TBS
	
	' TBS Production and Restan
	strSTEPStmt = " stepnumber = 'STEP_02' and codedesc in ('SR003_01','SR003_02') " 	
	strTempProduction = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strProduction = "SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempProduction & strORDERStmt & " ) " 
 	Set rsProduction = Server.CreateObject("ADODB.Recordset")
	' response.write strProduction
	
	rsProduction.Open strProduction, connstr, 3,3
	
	
	Dim strSTEPPusinganStmt, strMAIN2PusinganStmt
	' Pusingan ==> Mengambil yang maximum
	strSTEPPusinganStmt = " SELECT aa.codedesc," & _
                          "        aa.jobgencode," & _
                          "        aa.description," & _
                          "        aa.unitofmeasurement," & _
                          "        aa.groupplantationcode," & _
                          "        aa.plantationcode," & _
                          "        aa.groupregioncode," & _
                          "        aa.regioncode," & _
                          "        aa.companycode," & _
                          "        aa.fcba," & _
                          "        aa.seqnumber seqnumber," & _
                          "        aa.todate hi," & _
                          "        aa.tomonth shi," & _
                          "        aa.toyear sbi," & _
                          "        aa.budgettodate budget_hi," & _
                          "        aa.budgettomonth budget_shi," & _
                          "        aa.budgettoyear budget_sbi," & _
                          "        bb.ha ha_hi," & _
                          "        bb.interval_ha interval_ha_hi" & _
                          " FROM production_summary_tab aa" & _
                          " LEFT JOIN harvesting_interval_avg_tab bb" & _
                          " ON (aa.fcba = bb.fcba AND aa.ondate = bb.ondate)" & _
                          " WHERE USER_COMPANY_API.GETCOMPANYCODE ('" & Session("CodeName") & "', aa.fcba) IS NOT NULL" & _
                          "       AND TO_CHAR (aa.ondate,'YYYYMMDD') = TO_CHAR (TO_DATE ('" & txtdatefrom & "', 'MM/DD/YYYY'), 'YYYYMMDD')" & _
                          "       AND stepnumber = 'STEP_02'" & _
                          "       AND codedesc = 'SR003_03'"
				  
	
	strMAIN2PusinganStmt =  " , CASE SUM(ha_hi) WHEN 0 THEN 0 ELSE ROUND(SUM(interval_ha_hi)/SUM(ha_hi), 2) END hi," & _
							"   SUM(shi) shi," & _
							"   SUM(sbi) sbi," & _
							"   SUM(budget_hi) budget_hi," & _
							"   SUM(budget_shi) budget_shi," & _
							"   SUM(budget_sbi) budget_sbi"

			
			
	strTempProduction = strSTEPPusinganStmt 
	strProduction = " SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2PusinganStmt & " from ( " & strTempProduction & " )" & _
	                " WHERE codedesc = 'SR003_03' " & strWHEREStmt & " GROUP BY seqnumber, " & strSELECTStmt & strORDERStmt
	Set rsPusingan = Server.CreateObject("ADODB.Recordset")
	' response.write "Select Statement Pusingan : " & strProduction
		
	rsPusingan.Open strProduction, connstr, 3,3
	
	
	' Curah Hujan ==> Total Summary, namun diletakan paling akhir dari Group Produksi
	strSTEPStmt = " stepnumber = 'STEP_02' and codedesc in ('SR003_04') " 	
	strTempProduction = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strProduction = "SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempProduction & strORDERStmt & " ) " 
 	Set rsCurahHujan = Server.CreateObject("ADODB.Recordset")
	' response.write strProduction
		
	rsCurahHujan.Open strProduction, connstr, 3,3
	
	

	' Step 03 -- Output
	strSTEPStmt = " stepnumber = 'STEP_03' " 	
	strTempOutput = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strOutput = "SELECT JobGenCode theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempOutput & " ) "  & _
	            " ORDER BY 'x' , JobGenCode, description,  UnitOfMeasurement " & strGROUPRENDEMENStmt
	

 	Set rsOutput = Server.CreateObject("ADODB.Recordset")
	' response.write strOutput
	
	rsOutput.Open strOutput, connstr, 3,3
	
	' This Step already changed with strRatio by NAF
	' Step 04 -- Handling
	'strSTEPStmt = " stepnumber = 'STEP_04' "  	
	'strTempHandling = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	'strHandling = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempHandling & strORDERStmt & " ) " 
 	'Set rsHandling = Server.CreateObject("ADODB.Recordset")
	'response.write strHandling
	
	'rsHandling.Open strHandling, connstr, 3,3
	
	
	
	' This Step already changed with strRatioAmount by NAF
	' Step 05 -- HandlingAmount
	'strSTEPStmt = " stepnumber = 'STEP_05' " 	
	'strTempHandlingAmount = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	'strHandlingAmount = "SELECT rownum theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempHandlingAmount & strORDERStmt & " ) " 
 	'Set rsHandlingAmount = Server.CreateObject("ADODB.Recordset")
	'response.write strHandlingAmount
	
	'rsHandlingAmount.Open strHandlingAmount, connstr, 3,3

	
	' Step 06 -- Fertilizer
	strSTEPStmt = " stepnumber = 'STEP_06' " 	
	strTempFertilizer = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strFertilizer = "SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempFertilizer & strORDERStmt & " ) " 
 	Set rsFertilizer = Server.CreateObject("ADODB.Recordset")
	' response.write strFertilizer
	
	rsFertilizer.Open strFertilizer, connstr, 3,3
	

	' Step 07 -- Mandays
	strSTEPStmt = " stepnumber = 'STEP_07' " 	
	strTempMandays = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strMandays = "SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempMandays & strORDERStmt & " ) " 
 	Set rsMandays = Server.CreateObject("ADODB.Recordset")
	' response.write strMandays
	
	rsMandays.Open strMandays, connstr, 3,3	
	
	
	' Step 08 -- VehicleGroup
	strSTEPStmt = " stepnumber = 'STEP_08' " 	
	strTempVehicleGroup = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strVehicleGroup = "SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempVehicleGroup & strORDERStmt & " ) " 
 	Set rsVehicleGroup = Server.CreateObject("ADODB.Recordset")
	' response.write strVehicleGroup
	
	rsVehicleGroup.Open strVehicleGroup, connstr, 3,3	
	
	' Step 09 -- Ending Balance
	strSTEPStmt = " stepnumber = 'STEP_09' " 	
	strTempEndingBalance = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & strGROUPStmt 
	strEndingBalance = "SELECT seqnumber theNumber, " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTempEndingBalance & strORDERStmt & " ) " 
 	Set rsEndingBalance = Server.CreateObject("ADODB.Recordset")
	' response.write strEndingBalance
	
	rsEndingBalance.Open strEndingBalance, connstr, 3,3		

	
	' *************************SPECIAL FOR RATA-RATA BIAYA MENGUMPUL************************
	' Rata-Rata Biaya Mengumpul = SR003_05_AMT / SR003_05_TON
	' Step SORTASI UTAMA
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_05_AMT' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) " 

	
	strSORTASIstmt =   ", sum(hi) amount_hi, " & _
				" sum(shi) amount_shi, " & _
				" sum(sbi) amount_sbi, " & _
				" sum(budget_hi) amount_budget_hi,  " & _
				" sum(budget_shi) amount_budget_shi,  " & _
				" sum(budget_sbi) amount_budget_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 tonase_hi, " & _
				" 0 tonase_shi, " & _
				" 0 tonase_sbi, " & _
				" 0 tonase_budget_hi,  " & _
				" 0 tonase_budget_shi,  " & _
				" 0 tonase_budget_sbi   " 	



	strSortasi = "SELECT 'RATA_RATA_MENGUMPUL' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RATA_RATA_MENGUMPUL' " & strGROUPRENDEMENStmt
	
	'=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_05_TON' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strSORTASIstmt =   ", 0 amount_hi, " & _
				" 0 amount_shi, " & _
				" 0 amount_sbi, " & _
				" 0 amount_budget_hi,  " & _
				" 0 amount_budget_shi,  " & _
				" 0 amount_budget_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", sum(hi) tonase_hi, " & _
				" sum(shi) tonase_shi, " & _
				" sum(sbi) tonase_sbi, " & _
				" sum(budget_hi) tonase_budget_hi,  " & _
				" sum(budget_shi) tonase_budget_shi,  " & _
				" sum(budget_sbi) tonase_budget_sbi   " 	
	

	strSortasi = strSortasi & " UNION SELECT 'RATA_RATA_MENGUMPUL' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RATA_RATA_MENGUMPUL' " & strGROUPRENDEMENStmt
	
	
	
	strSORTASIFinalStmt = " , sum(amount_hi) amount_hi , " & _
						   " sum(amount_shi) amount_shi , " & _
						   " sum(amount_sbi) amount_sbi , " & _
						   " sum(amount_budget_hi) amount_budget_hi , " & _
						   " sum(amount_budget_shi) amount_budget_shi , " & _
						   " sum(amount_budget_sbi) amount_budget_sbi , " & _
						   " sum(tonase_hi) tonase_hi , " & _
						   " sum(tonase_shi) tonase_shi , " & _
						   " sum(tonase_sbi) tonase_sbi , " & _
						   " sum(tonase_budget_hi) tonase_budget_hi , " & _
						   " sum(tonase_budget_shi) tonase_budget_shi , " & _
						   " sum(tonase_budget_sbi) tonase_budget_sbi  " 
					
	
	strSortasiFinal = "SELECT title HEADGROUP,'' JobGenCode, 'Biaya Panen' description, 'Rp/Kg' UnitOfMeasurement, 'SR003_05' CodeDesc " & strSELECTStmt_Rendemen & strSORTASIFinalStmt & " FROM ( " & strSortasi & " ) " & _
	                   " GROUP BY title " & strGROUPRENDEMENStmt
	
	

	Set rsBiayaMengumpulFinal = Server.CreateObject("ADODB.Recordset")
	' response.write strSortasiFinal
	rsBiayaMengumpulFinal.Open strSortasiFinal, connstr, 3,3		
	
	' *************************END OF SPECIAL************************
	


	' *************************SPECIAL FOR TBS_PENGANGKUTAN************************
	' TBS Pengangkutan = sum(SR003_BA_CURR_QTY) *  ( sum(SR003_BA_PREV_AMT) / sum(SR003_BA_PREV_QTY)  ) / sum(SR003_01)
	' TBS Pengangkutan = Current QTY * Factor Harga Bulan Sebelumnya  / Jumlah TBS 
	' Step SORTASI UTAMA
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_BA_CURR_QTY' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) " 

	
	strSORTASIstmt =   ", sum(hi) curr_qty_hi, " & _
				" sum(shi) curr_qty_shi, " & _
				" sum(sbi) curr_qty_sbi, " & _
				" sum(budget_hi) curr_budget_qty_hi,  " & _
				" sum(budget_shi) curr_budget_qty_shi,  " & _
				" sum(budget_sbi) curr_budget_qty_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 prev_amt_hi, " & _
				" 0 prev_amt_shi, " & _
				" 0 prev_amt_sbi, " & _
				" 0 prev_budget_amt_hi,  " & _
				" 0 prev_budget_amt_shi,  " & _
				" 0 prev_budget_amt_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 prev_qty_hi, " & _
				" 0 prev_qty_shi, " & _
				" 0 prev_qty_sbi, " & _
				" 0 prev_budget_qty_hi,  " & _
				" 0 prev_budget_qty_shi,  " & _
				" 0 prev_budget_qty_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 tbs_qty_hi, " & _
				" 0 tbs_qty_shi, " & _
				" 0 tbs_qty_sbi, " & _
				" 0 tbs_budget_qty_hi,  " & _
				" 0 tbs_budget_qty_shi,  " & _
				" 0 tbs_budget_qty_sbi   " 	


	strSortasi = "SELECT 'TBS_PENGANGKUTAN' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'TBS_PENGANGKUTAN' " & strGROUPRENDEMENStmt
	
	'=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_BA_PREV_AMT' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strSORTASIstmt =   ", 0 curr_qty_hi, " & _
				" 0 curr_qty_shi, " & _
				" 0 curr_qty_sbi, " & _
				" 0 curr_budget_qty_hi,  " & _
				" 0 curr_budget_qty_shi,  " & _
				" 0 curr_budget_qty_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", sum(hi) prev_amt_hi, " & _
				" sum(shi) prev_amt_shi, " & _
				" sum(sbi) prev_amt_sbi, " & _
				" sum(budget_hi) prev_budget_amt_hi,  " & _
				" sum(budget_shi) prev_budget_amt_shi,  " & _
				" sum(budget_sbi) prev_budget_amt_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 prev_qty_hi, " & _
				" 0 prev_qty_shi, " & _
				" 0 prev_qty_sbi, " & _
				" 0 prev_budget_qty_hi,  " & _
				" 0 prev_budget_qty_shi,  " & _
				" 0 prev_budget_qty_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 tbs_qty_hi, " & _
				" 0 tbs_qty_shi, " & _
				" 0 tbs_qty_sbi, " & _
				" 0 tbs_budget_qty_hi,  " & _
				" 0 tbs_budget_qty_shi,  " & _
				" 0 tbs_budget_qty_sbi   " 					
	

	strSortasi = strSortasi & " UNION SELECT 'TBS_PENGANGKUTAN' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'TBS_PENGANGKUTAN' " & strGROUPRENDEMENStmt
	
	'=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_BA_PREV_QTY' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strSORTASIstmt =   ", 0 curr_qty_hi, " & _
				" 0 curr_qty_shi, " & _
				" 0 curr_qty_sbi, " & _
				" 0 curr_budget_qty_hi,  " & _
				" 0 curr_budget_qty_shi,  " & _
				" 0 curr_budget_qty_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 prev_amt_hi, " & _
				" 0 prev_amt_shi, " & _
				" 0 prev_amt_sbi, " & _
				" 0 prev_budget_amt_hi,  " & _
				" 0 prev_budget_amt_shi,  " & _
				" 0 prev_budget_amt_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", sum(hi) prev_qty_hi, " & _
				" sum(shi) prev_qty_shi, " & _
				" sum(sbi) prev_qty_sbi, " & _
				" sum(budget_hi) prev_budget_qty_hi,  " & _
				" sum(budget_shi) prev_budget_qty_shi,  " & _
				" sum(budget_sbi) prev_budget_qty_sbi   " 		
	strSORTASIstmt = strSORTASIstmt &  ", 0 tbs_qty_hi, " & _
				" 0 tbs_qty_shi, " & _
				" 0 tbs_qty_sbi, " & _
				" 0 tbs_budget_qty_hi,  " & _
				" 0 tbs_budget_qty_shi,  " & _
				" 0 tbs_budget_qty_sbi   " 					
	

	strSortasi = strSortasi & " UNION SELECT 'TBS_PENGANGKUTAN' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'TBS_PENGANGKUTAN' " & strGROUPRENDEMENStmt

    '=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_01' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strSORTASIstmt =   ", 0 curr_qty_hi, " & _
				" 0 curr_qty_shi, " & _
				" 0 curr_qty_sbi, " & _
				" 0 curr_budget_qty_hi,  " & _
				" 0 curr_budget_qty_shi,  " & _
				" 0 curr_budget_qty_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 prev_amt_hi, " & _
				" 0 prev_amt_shi, " & _
				" 0 prev_amt_sbi, " & _
				" 0 prev_budget_amt_hi,  " & _
				" 0 prev_budget_amt_shi,  " & _
				" 0 prev_budget_amt_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 prev_qty_hi, " & _
				" 0 prev_qty_shi, " & _
				" 0 prev_qty_sbi, " & _
				" 0 prev_budget_qty_hi,  " & _
				" 0 prev_budget_qty_shi,  " & _
				" 0 prev_budget_qty_sbi   " 		
	strSORTASIstmt = strSORTASIstmt &  ", sum(hi) tbs_qty_hi, " & _
				" sum(shi) tbs_qty_shi, " & _
				" sum(sbi) tbs_qty_sbi, " & _
				" sum(budget_hi) tbs_budget_qty_hi,  " & _
				" sum(budget_shi) tbs_budget_qty_shi,  " & _
				" sum(budget_sbi) tbs_budget_qty_sbi   " 					
	

	strSortasi = strSortasi & " UNION SELECT 'TBS_PENGANGKUTAN' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'TBS_PENGANGKUTAN' " & strGROUPRENDEMENStmt
	
	
	
	
	
	strSORTASIFinalStmt = " , sum(curr_qty_hi) curr_qty_hi , " & _
						   " sum(curr_qty_shi) curr_qty_shi , " & _
						   " sum(curr_qty_sbi) curr_qty_sbi , " & _
						   " sum(curr_budget_qty_hi) curr_budget_qty_hi , " & _
						   " sum(curr_budget_qty_shi) curr_budget_qty_shi , " & _
						   " sum(curr_budget_qty_sbi) curr_budget_qty_sbi , " & _
						   " sum(prev_amt_hi) prev_amt_hi , " & _
						   " sum(prev_amt_shi) prev_amt_shi , " & _
						   " sum(prev_amt_sbi) prev_amt_sbi , " & _
						   " sum(prev_budget_amt_hi) prev_budget_amt_hi , " & _
						   " sum(prev_budget_amt_shi) prev_budget_amt_shi , " & _
						   " sum(prev_budget_amt_sbi) prev_budget_amt_sbi,  " & _
						   " sum(prev_qty_hi) prev_qty_hi , " & _
						   " sum(prev_qty_shi) prev_qty_shi , " & _
						   " sum(prev_qty_sbi) prev_qty_sbi , " & _
						   " sum(prev_budget_qty_hi) prev_budget_qty_hi , " & _
						   " sum(prev_budget_qty_shi) prev_budget_qty_shi , " & _
						   " sum(prev_budget_qty_sbi) prev_budget_qty_sbi , " & _	
						   " sum(tbs_qty_hi) tbs_qty_hi , " & _
						   " sum(tbs_qty_shi) tbs_qty_shi , " & _
						   " sum(tbs_qty_sbi) tbs_qty_sbi , " & _
						   " sum(tbs_budget_qty_hi) tbs_budget_qty_hi , " & _
						   " sum(tbs_budget_qty_shi) tbs_budget_qty_shi , " & _
						   " sum(tbs_budget_qty_sbi) tbs_budget_qty_sbi  " 						   
	
	strTBSPengangkutan = "SELECT title HEADGROUP,'' JobGenCode, 'TBS Pengangkutan' description, 'Rp/Kg' UnitOfMeasurement, 'SR003_BA' CodeDesc " & strSELECTStmt_Rendemen & strSORTASIFinalStmt & " FROM ( " & strSortasi & " ) " & " GROUP BY title " & strGROUPRENDEMENStmt
	
	
	

	Set rsTBSPengangkutanFinal = Server.CreateObject("ADODB.Recordset")
	'response.write "NAFIS2 : " & strTBSPengangkutan
	rsTBSPengangkutanFinal.Open strTBSPengangkutan, connstr, 3,3		
	
	' *************************END OF SPECIAL************************

	' *************************SPECIAL FOR OUTPUT PANEN************************
	' Output Panen = sum(SR003_01) / sum(SR003_OP) ==> Kg/HK
	' Output Panen = Jumlah TBS / Jumlah HK 
	' Step SORTASI UTAMA
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_01' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) " 

	
	strSORTASIstmt =   ", sum(hi) tbs_qty_hi, " & _
				" sum(shi) tbs_qty_shi, " & _
				" sum(sbi) tbs_qty_sbi, " & _
				" sum(budget_hi) tbs_budget_qty_hi,  " & _
				" sum(budget_shi) tbs_budget_qty_shi,  " & _
				" sum(budget_sbi) tbs_budget_qty_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 mandays_qty_hi, " & _
				" 0 mandays_qty_shi, " & _
				" 0 mandays_qty_sbi, " & _
				" 0 mandays_qty_budget_hi,  " & _
				" 0 mandays_qty_budget_shi,  " & _
				" 0 mandays_qty_budget_sbi   " 	


	strSortasi = "SELECT 'OUTPUT_PANEN' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'OUTPUT_PANEN' " & strGROUPRENDEMENStmt
	
	'=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_02' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_OP' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strSORTASIstmt =   ", 0 tbs_qty_hi, " & _
				" 0 tbs_qty_shi, " & _
				" 0 tbs_qty_sbi, " & _
				" 0 tbs_budget_qty_hi,  " & _
				" 0 tbs_budget_qty_shi,  " & _
				" 0 tbs_budget_qty_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", sum(hi) mandays_qty_hi, " & _
				" sum(shi) mandays_qty_shi, " & _
				" sum(sbi) mandays_qty_sbi, " & _
				" sum(budget_hi) mandays_qty_budget_hi,  " & _
				" sum(budget_shi) mandays_qty_budget_shi,  " & _
				" sum(budget_sbi) mandays_qty_budget_sbi   " 					
	

	strSortasi = strSortasi & " UNION SELECT 'OUTPUT_PANEN' title " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'OUTPUT_PANEN' " & strGROUPRENDEMENStmt
	
	
	
	strSORTASIFinalStmt = " , sum(tbs_qty_hi) tbs_qty_hi , " & _
						   " sum(tbs_qty_shi) tbs_qty_shi , " & _
						   " sum(tbs_qty_sbi) tbs_qty_sbi , " & _
						   " sum(tbs_budget_qty_hi) tbs_budget_qty_hi , " & _
						   " sum(tbs_budget_qty_shi) tbs_budget_qty_shi , " & _
						   " sum(tbs_budget_qty_sbi) tbs_budget_qty_sbi , " & _
						   " sum(mandays_qty_hi) mandays_qty_hi , " & _
						   " sum(mandays_qty_shi) mandays_qty_shi , " & _
						   " sum(mandays_qty_sbi) mandays_qty_sbi , " & _
						   " sum(mandays_qty_budget_hi) mandays_qty_budget_hi , " & _
						   " sum(mandays_qty_budget_shi) mandays_qty_budget_shi , " & _
						   " sum(mandays_qty_budget_sbi) mandays_qty_budget_sbi  " 
					   
	
	strOutputPanen = "SELECT title HEADGROUP,'' JobGenCode, 'Output Panen' description, 'Kg/HK' UnitOfMeasurement, 'SR003_OP' CodeDesc " & strSELECTStmt_Rendemen & strSORTASIFinalStmt & " FROM ( " & strSortasi & " ) " & " GROUP BY title " & strGROUPRENDEMENStmt
	
	

	Set rsOutputPanenFinal = Server.CreateObject("ADODB.Recordset")
	'response.write "NAFIS2 : " & strOutputPanen
	rsOutputPanenFinal.Open strOutputPanen, connstr, 3,3		
	
	' *************************END OF SPECIAL************************


	
	' *************************SPECIAL FOR RATIO HK/HA************************
	' codedesc in ('SR003_06_MANDAYS','SR003_06_HA') "
	' Rata-Rata Biaya Mandays = SR003_06_MANDAYS / SR003_06_HA
	' Step SORTASI UTAMA
	strSTEPStmt = " stepnumber = 'STEP_04'  " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_06_MANDAYS' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) " 

	
	strSORTASIstmt =   ", sum(hi) mandays_hi, " & _
				" sum(shi) mandays_shi, " & _
				" sum(sbi) mandays_sbi, " & _
				" sum(budget_hi) mandays_budget_hi,  " & _
				" sum(budget_shi) mandays_budget_shi,  " & _
				" sum(budget_sbi) mandays_budget_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 hectare_hi, " & _
				" 0 hectare_shi, " & _
				" 0 hectare_sbi, " & _
				" 0 hectare_budget_hi,  " & _
				" 0 hectare_budget_shi,  " & _
				" 0 hectare_budget_sbi   " 	



	strSortasi = "SELECT 'RATIO' title , JobGenCode, description,  UnitOfMeasurement " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RATIO' , JobGenCode, description,  UnitOfMeasurement " & strGROUPRENDEMENStmt
	
	'=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_04' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_06_HA' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strSORTASIstmt =   ", 0 mandays_hi, " & _
				" 0 mandays_shi, " & _
				" 0 mandays_sbi, " & _
				" 0 mandays_budget_hi,  " & _
				" 0 mandays_budget_shi,  " & _
				" 0 mandays_budget_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", sum(hi) hectare_hi, " & _
				" sum(shi) hectare_shi, " & _
				" sum(sbi) hectare_sbi, " & _
				" sum(budget_hi) hectare_budget_hi,  " & _
				" sum(budget_shi) hectare_budget_shi,  " & _
				" sum(budget_sbi) hectare_budget_sbi   " 	
	

	strSortasi = strSortasi & " UNION SELECT 'RATIO' title, JobGenCode, description,  UnitOfMeasurement  " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RATIO', JobGenCode, description,  UnitOfMeasurement  " & strGROUPRENDEMENStmt
	
	
	
	strSORTASIFinalStmt = " , sum(mandays_hi) mandays_hi , " & _
						   " sum(mandays_shi) mandays_shi , " & _
						   " sum(mandays_sbi) mandays_sbi , " & _
						   " sum(mandays_budget_hi) mandays_budget_hi , " & _
						   " sum(mandays_budget_shi) mandays_budget_shi , " & _
						   " sum(mandays_budget_sbi) mandays_budget_sbi , " & _
						   " sum(hectare_hi) hectare_hi , " & _
						   " sum(hectare_shi) hectare_shi , " & _
						   " sum(hectare_sbi) hectare_sbi , " & _
						   " sum(hectare_budget_hi) hectare_budget_hi , " & _
						   " sum(hectare_budget_shi) hectare_budget_shi , " & _
						   " sum(hectare_budget_sbi) hectare_budget_sbi  " 
					
	
	strRatio = "SELECT title HEADGROUP, JobGenCode, description,  UnitOfMeasurement, 'SR003_06' CodeDesc " & strSELECTStmt_Rendemen & strSORTASIFinalStmt & " FROM ( " & strSortasi & " ) " & _
	                   " GROUP BY title , JobGenCode, description,  UnitOfMeasurement " & strGROUPRENDEMENStmt & _
					   " ORDER BY title , JobGenCode, description,  UnitOfMeasurement " & strGROUPRENDEMENStmt
	
	

	Set rsRatioFinal = Server.CreateObject("ADODB.Recordset")
	'response.write strRatio
	rsRatioFinal.Open strRatio, connstr, 3,3		
	
	' *************************END OF SPECIAL************************	
	
	
	' *************************SPECIAL FOR RP/HA************************
	' codedesc in ('SR003_07_AMOUNT','SR003_07_HECTARE') "
	' Rata-Rata Biaya Mandays = SR003_06_MANDAYS / SR003_06_HA
	' Step SORTASI UTAMA
	strSTEPStmt = " stepnumber = 'STEP_05'  " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_07_AMOUNT' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & " ) " 

	
	strSORTASIstmt =   ", sum(hi) amount_hi, " & _
				" sum(shi) amount_shi, " & _
				" sum(sbi) amount_sbi, " & _
				" sum(budget_hi) amount_budget_hi,  " & _
				" sum(budget_shi) amount_budget_shi,  " & _
				" sum(budget_sbi) amount_budget_sbi   " 	
	strSORTASIstmt = strSORTASIstmt &  ", 0 hectare_hi, " & _
				" 0 hectare_shi, " & _
				" 0 hectare_sbi, " & _
				" 0 hectare_budget_hi,  " & _
				" 0 hectare_budget_shi,  " & _
				" 0 hectare_budget_sbi   " 	



	strSortasi = "SELECT 'RATIO' title , JobGenCode, description,  UnitOfMeasurement " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RATIO' , JobGenCode, description,  UnitOfMeasurement " & strGROUPRENDEMENStmt
	
	'=================== CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_05' " 	
	strTemp = "SELECT " & strSELECTStmt & strMAINStmt & strFROMStmt & strSTEPStmt & strWHEREStmt & " AND CodeDesc = 'SR003_07_HECTARE' " & strGROUPStmt 
	strSQL = "SELECT " & strSELECTStmt & strMAIN2Stmt & " from ( " & strTemp & strORDERStmt & " ) " 

	strSORTASIstmt =   ", 0 amount_hi, " & _
				" 0 amount_shi, " & _
				" 0 amount_sbi, " & _
				" 0 amount_budget_hi,  " & _
				" 0 amount_budget_shi,  " & _
				" 0 amount_budget_sbi   " 
	strSORTASIstmt = strSORTASIstmt &  ", sum(hi) hectare_hi, " & _
				" sum(shi) hectare_shi, " & _
				" sum(sbi) hectare_sbi, " & _
				" sum(budget_hi) hectare_budget_hi,  " & _
				" sum(budget_shi) hectare_budget_shi,  " & _
				" sum(budget_sbi) hectare_budget_sbi   " 	
	

	strSortasi = strSortasi & " UNION SELECT 'RATIO' title, JobGenCode, description,  UnitOfMeasurement  " & strSELECTStmt_Rendemen & strSORTASIstmt & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RATIO', JobGenCode, description,  UnitOfMeasurement  " & strGROUPRENDEMENStmt
	
	
	
	strSORTASIFinalStmt = " , sum(amount_hi) amount_hi , " & _
						   " sum(amount_shi) amount_shi , " & _
						   " sum(amount_sbi) amount_sbi , " & _
						   " sum(amount_budget_hi) amount_budget_hi , " & _
						   " sum(amount_budget_shi) amount_budget_shi , " & _
						   " sum(amount_budget_sbi) amount_budget_sbi , " & _
						   " sum(hectare_hi) hectare_hi , " & _
						   " sum(hectare_shi) hectare_shi , " & _
						   " sum(hectare_sbi) hectare_sbi , " & _
						   " sum(hectare_budget_hi) hectare_budget_hi , " & _
						   " sum(hectare_budget_shi) hectare_budget_shi , " & _
						   " sum(hectare_budget_sbi) hectare_budget_sbi  " 
					
	
	strRatioAmount = "SELECT title HEADGROUP, JobGenCode, description,  UnitOfMeasurement, 'SR003_07' CodeDesc " & strSELECTStmt_Rendemen & strSORTASIFinalStmt & " FROM ( " & strSortasi & " ) " & _
	                   " GROUP BY title , JobGenCode, description,  UnitOfMeasurement " & strGROUPRENDEMENStmt & _
					   " ORDER BY title , JobGenCode, description,  UnitOfMeasurement "  & strGROUPRENDEMENStmt 
	
	

	Set rsRatioAmountFinal = Server.CreateObject("ADODB.Recordset")
	'response.write strRatioAmount
	rsRatioAmountFinal.Open strRatioAmount, connstr, 3,3		
	
	' *************************END OF SPECIAL************************	
	
	
	
	
%>

<!--#include file="FormAdditionalQueryRendemen.asp"-->	


<html>

<head>
<title>ePlaS Reporting - Maintenance Report</title>
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

<script type="text/javascript">
	function ajaxurlpost(url, param) {
	
		$.ajax({
			type: "POST",
			url: url,
			data: { RF_username: '<%=Session("CodeName")%>',
					RF_ipaddress: '<%=Request.ServerVariables("REMOTE_ADDR")%>',
					RF_fcba: '<%=Session("fcBA")%>',
					RF_date_act: '<%=FormatDateTime(Date, 1)%>',
					RF_object_name: param
			},
			success: function(){
				var view = document.getElementById("frmEntry");
				view.submit();
				//alert("test");
			}
		});
		
	}
</script>
</head>
<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<form id="frmEntry" name="frmEntry" method="post" action="p_Summary_BA_NAF.asp">


			<!--#include file="menutop.asp"-->	
			<!--#include file="FormSelectNewSum.html"-->
			<!--#include file="FormSelectSUMMARYProduction.html"-->


			<table border="0" width="100%" id="product-table" class='xthf-site'>
			<!--#include file="FormSelectHEADERProduction.html"-->
				<tbody>
					<tr>
						<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
					</tr>	
					
					<%
						sPrevGroupPlantation = ""			
						sPrevPlantation = ""
						sPrevGroupRegion = ""
						sPrevRegion = ""
						sPrevCompany = ""
						sPrevFcba = ""
						sPrevDesc = ""
						sPrevSubDesc = ""
					%>

					<!-- Start COUNT ROWSPAN LandClearing -->
					<!--#include file="FormCountRowLandClearing.asp"-->					
					
					<!-- Start LandClearing -->	
					<!--#include file="FormDisplayLandClearing.html"-->
					

					<tr>
						<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
					</tr>
					
					<!-- Start COUNT ROWSPAN Production -->
					<!--#include file="FormCountRowProduction.asp"-->
					
					

					<!-- Start Production -->	
					<!--#include file="FormDisplayProduction.html"-->
					<!-- Start Production -->	
					<!--#include file="FormDisplayPusingan.html"-->
					<!-- Start Output Panen -->	
					<!--#include file="FormDisplayOutputPanen.html"-->							
					<!-- Start Biaya Mengumpul -->	
					<!--#include file="FormDisplayBiayaMengumpul.html"-->				
					<!-- Start TBS Pengangkutan -->	
					<!--#include file="FormDisplayTBSPengangkutan.html"-->							
					<!-- Start Curah Hujan -->	
					<!--#include file="FormDisplayCurahHujan.html"-->	
					
					
					<tr>
						<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
					</tr>
					
					<!-- Start COUNT ROWSPAN Rendemen - O E R -->
					<!--#include file="FormCountRowRendemenAdd.asp"-->																										
					<!-- Start COUNT ROWSPAN Rendemen - K E R -->
					<!--#include file="FormCountRowRendemenAddKer.asp"-->																										

					<!-- Start Rendemen -->		
					<!--#include file="FormDisplayRendemenAdd.html"-->
					<!-- Start Rendemen -->		
					<!--#include file="FormDisplayRendemenAddKer.html"-->

					
					<tr>
						<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"></th>				
					</tr>	
					
					<!-- Start COUNT ROWSPAN Quality -->
					<!--#include file="FormCountRowQualityAdd.asp"-->			
					
					<!-- CPO Quality -->		
					<!-- Start FFA Quality for CPO Production -->		
					<!--#include file="FormDisplayCPO_FFAQualityAdd.html"-->	
			
					
					<tr>
					<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"  ></th>				
					</tr>
					<!-- Start COUNT ROWSPAN Output -->
					<!--#include file="FormCountRowOutput.asp"-->
					<!-- Start Output -->	
					<!--#include file="FormDisplayOutput.html"-->
					<!------------------------------------------------------------->	
					

					<!-- Start COUNT ROWSPAN Handling -->
					<!--#include file="FormCountRowHandling.asp"-->
					<!-- Start Handling -->	
					<!--#include file="FormDisplayHandling.html"-->
					<!------------------------------------------------------------->	

					<!-- Start COUNT ROWSPAN HandlingAmount -->
					<!--#include file="FormCountRowHandlingAmount.asp"-->
					<!-- Start HandlingAmount -->	
					<!--#include file="FormDisplayHandlingAmount.html"-->
					<!------------------------------------------------------------->	
					
					<tr>
					<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"  ></th>				
					</tr>

					<!-- Start COUNT ROWSPAN Fertilizer -->
					<!--#include file="FormCountRowFertilizer.asp"-->
					<!-- Start Fertilizer -->	
					<!--#include file="FormDisplayFertilizer.html"-->
			
					<tr>
					<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"  ></th>				
					</tr>

					<!-- Start COUNT ROWSPAN Mandays -->
					<!--#include file="FormCountRowMandays.asp"-->
					<!-- Start Tenaga Kerja -->	
					<!--#include file="FormDisplayMandays.html"-->
					
					<tr>
					<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"  ></th>				
					</tr>

					<!-- Start COUNT ROWSPAN VehicleGroup -->
					<!--#include file="FormCountRowVehicleGroup.asp"-->
					<!-- Start Vehicle Group -->	
					<!--#include file="FormDisplayVehicleGroup.html"-->
					
					<tr>
					<th class="separator" colspan="<%= 13 + nColSpan %>" align="left"  ></th>				
					</tr>

					<!-- Start COUNT ROWSPAN EndingBalance -->
					<!--#include file="FormCountRowEndingBalance.asp"-->
					<!-- Start Ending Balance -->	
					<!--#include file="FormDisplayEndingBalance.html"-->

				</tbody>

			</table>

		</form>
     <!--#include file="indicator_color.asp"-->
	</div>
    
</body>
</html>




