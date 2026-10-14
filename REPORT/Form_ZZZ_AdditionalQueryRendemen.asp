
<%
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	' *************************SPECIAL FOR RENDEMEN************************
	
	Dim strRendemenAddFinal, rsRendemenFinalAdd, strSELECTStmtAdd_Rendemen, strRENDEMENFinalStmt, strRendemenAdd, strGROUPRENDEMENStmtAdd
	Dim strSELECTStmtAdd, strMAINStmtAdd, strMAINAVGStmtAdd, strMAIN2StmtAdd, strFROMStmtAdd, strWHEREStmtAdd, strGROUPStmtAdd, strORDERStmtAdd
	Dim strRENDEMENstmtAdd
	
	
	strSELECTStmtAdd = " CODEDESC, HEADGROUP, GROUPNAME, UNITOFMEASUREMENT "
	strSELECTStmtAdd_Rendemen = ""

	if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strSELECTStmtAdd = strSELECTStmtAdd & " , groupplantationcode "
			strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , groupplantationcode "
		end if
	else
	    strSELECTStmtAdd = strSELECTStmtAdd & " , groupplantationcode "
	    strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , groupplantationcode "
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strSELECTStmtAdd = strSELECTStmtAdd & " , plantationcode "
			strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , plantationcode "
		end if
	else
	    strSELECTStmtAdd = strSELECTStmtAdd & " , plantationcode "
	    strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strSELECTStmtAdd = strSELECTStmtAdd & " , groupregioncode "
			strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , groupregioncode "
		end if
	else
	    strSELECTStmtAdd = strSELECTStmtAdd & " , groupregioncode "
	    strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strSELECTStmtAdd = strSELECTStmtAdd & " , regioncode "
			strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , regioncode "
		end if
	else
	    strSELECTStmtAdd = strSELECTStmtAdd & " , regioncode "
	    strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strSELECTStmtAdd = strSELECTStmtAdd & " , companycode "
			strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , companycode "
		end if
	else
	    strSELECTStmtAdd = strSELECTStmtAdd & " , companycode "
	    strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strSELECTStmtAdd = strSELECTStmtAdd & " , CUSTOMER_FCBA "
			strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , CUSTOMER_FCBA "
		end if
	else
	    strSELECTStmtAdd = strSELECTStmtAdd & " , CUSTOMER_FCBA "
	    strSELECTStmtAdd_Rendemen = strSELECTStmtAdd_Rendemen & " , CUSTOMER_FCBA "
	end if	

	strMAINStmtAdd =   " , sum(nvl(seqnumber,0)) seqnumber, " & _
					" sum(todate) hi, " & _
					" sum(tomonth) shi, " & _
					" sum(toyear) sbi, " & _
					" sum(budgettodate) budget_hi,  " & _
					" sum(budgettomonth) budget_shi,  " & _
					" sum(budgettoyear) budget_sbi   " 
	
	strMAINAVGStmtAdd =   " , avg(nvl(seqnumber,0)) seqnumber, " & _
					" avg(todate) hi, " & _
					" avg(tomonth) shi, " & _
					" avg(toyear) sbi, " & _
					" avg(budgettodate) budget_hi,  " & _
					" avg(budgettomonth) budget_shi,  " & _
					" avg(budgettoyear) budget_sbi   " 	
	
	
    strMAIN2StmtAdd =   " , seqnumber, " & _
					" hi, " & _
					" shi, " & _
					" sbi, " & _
					" budget_hi,  " & _
					" budget_shi,  " & _
					" budget_sbi   " 					


				
					
					
	strFROMStmtAdd = " from mill_summary_tab, customer  " & _
					" where  " & _
					" mill_summary_tab.fcba = customer.fcba and  " & _ 
					" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', customer.fcba) is not null   and "  & _
					" to_char(ondate,'YYYYMMDD') = to_char(TO_DATE ('" & v_txtdatefrom & "', 'MM/DD/YYYY'),'YYYYMMDD')  and " 
					
    ' Where Condion
	strWHEREStmtAdd = ""
	
	if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation = "--ALL--" then
	        strWHEREStmtAdd = strWHEREStmtAdd & " AND GroupPlantationCode like '%' " 
		else
		    if v_txtGroupPlantation <> "--SUMMARY--" then 
	            strWHEREStmtAdd = strWHEREStmtAdd & " AND GroupPlantationCode = '" & v_txtGroupPlantation& "' " 
			end if
		end if
	else
	    strWHEREStmtAdd = strWHEREStmtAdd & " AND GroupPlantationCode like '%' " 
	end if
	
	
	if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation = "--ALL--" then
	        strWHEREStmtAdd = strWHEREStmtAdd & " AND PlantationCode like '%' " 
		else
		    if v_txtPlantation <> "--SUMMARY--" then 
	            strWHEREStmtAdd = strWHEREStmtAdd & " AND PlantationCode = '" & v_txtPlantation& "' " 
			end if
		end if
	else
	    strWHEREStmtAdd = strWHEREStmtAdd & " AND PlantationCode like '%' " 
	end if
	
	
	if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion = "--ALL--" then
	        strWHEREStmtAdd = strWHEREStmtAdd & " AND GroupRegionCode like '%' " 
		else
		    if v_txtGroupRegion <> "--SUMMARY--" then 
	            strWHEREStmtAdd = strWHEREStmtAdd & " AND GroupRegionCode = '" & v_txtGroupRegion& "' " 
			end if
		end if
	else
	    strWHEREStmtAdd = strWHEREStmtAdd & " AND GroupRegionCode like '%' " 
	end if
	

	if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion = "--ALL--" then
	        strWHEREStmtAdd = strWHEREStmtAdd & " AND RegionCode like '%' " 
		else
		    if v_txtRegion <> "--SUMMARY--" then 
	            strWHEREStmtAdd = strWHEREStmtAdd & " AND RegionCode = '" & v_txtRegion& "' " 
			end if
		end if
	else
	    strWHEREStmtAdd = strWHEREStmtAdd & " AND RegionCode like '%' " 
	end if
	
	if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany = "--ALL--" then
	        strWHEREStmtAdd = strWHEREStmtAdd & " AND CompanyCode like '%' " 
		else
		    if v_txtcompany <> "--SUMMARY--" then 
	            strWHEREStmtAdd = strWHEREStmtAdd & " AND CompanyCode = '" & v_txtcompany& "' " 
			end if
		end if
	else
	    strWHEREStmtAdd = strWHEREStmtAdd & " AND CompanyCode like '%' " 
	end if	


	if len(trim(v_txtfcba)) > 0 then 
		if v_txtfcba = "--ALL--" then
	        strWHEREStmtAdd = strWHEREStmtAdd & " AND CUSTOMER_FCBA like '%' " 
		else
		    if v_txtfcba <> "--SUMMARY--" then 
	            strWHEREStmtAdd = strWHEREStmtAdd & " AND CUSTOMER_FCBA = '" & v_txtfcba& "' " 
			end if
		end if
	else
	    strWHEREStmtAdd = strWHEREStmtAdd & " AND CUSTOMER_FCBA like '%' " 
	end if	

					
	' Group By
	
	strGROUPStmtAdd = ""
	
	strGROUPStmtAdd = strGROUPStmtAdd & " GROUP BY CODEDESC, HEADGROUP, GROUPNAME, UNITOFMEASUREMENT "
	
    if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strGROUPStmtAdd = strGROUPStmtAdd & " , groupplantationcode "
			strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd &  " , groupplantationcode "
		end if
	else
	    strGROUPStmtAdd = strGROUPStmtAdd & " , groupplantationcode "
	    strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , groupplantationcode "
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strGROUPStmtAdd = strGROUPStmtAdd & " , plantationcode "
			strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , plantationcode "
		end if
	else
	    strGROUPStmtAdd = strGROUPStmtAdd & " , plantationcode "
	    strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strGROUPStmtAdd = strGROUPStmtAdd & " , groupregioncode "
			strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , groupregioncode "
		end if
	else
	    strGROUPStmtAdd = strGROUPStmtAdd & " , groupregioncode "
	    strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strGROUPStmtAdd = strGROUPStmtAdd & " , regioncode "
			strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , regioncode "
		end if
	else
	    strGROUPStmtAdd = strGROUPStmtAdd & " , regioncode "
	    strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strGROUPStmtAdd = strGROUPStmtAdd & " , companycode "
			strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , companycode "
		end if
	else
	    strGROUPStmtAdd = strGROUPStmtAdd & " , companycode "
	    strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strGROUPStmtAdd = strGROUPStmtAdd & " , CUSTOMER_FCBA "
			strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , CUSTOMER_FCBA "
		end if
	else
	    strGROUPStmtAdd = strGROUPStmtAdd & " , CUSTOMER_FCBA "
	    strGROUPRENDEMENStmtAdd = strGROUPRENDEMENStmtAdd & " , CUSTOMER_FCBA "
	end if	
	
	strORDERStmtAdd = " ORDER BY 'x' " 	
	
    if len(trim(v_txtGroupPlantation)) > 0 then 
		if v_txtGroupPlantation <> "--SUMMARY--" then
			strORDERStmtAdd = strORDERStmtAdd & " , groupplantationcode "
		end if
	else
	    strORDERStmtAdd = strORDERStmtAdd & " , groupplantationcode "
	end if		
	
    if len(trim(v_txtPlantation)) > 0 then 
		if v_txtPlantation <> "--SUMMARY--" then
			strORDERStmtAdd = strORDERStmtAdd & " , plantationcode "
		end if
	else
	    strORDERStmtAdd = strORDERStmtAdd & " , plantationcode "
	end if	
	
	
    if len(trim(v_txtGroupRegion)) > 0 then 
		if v_txtGroupRegion <> "--SUMMARY--" then
			strORDERStmtAdd = strORDERStmtAdd & " , groupregioncode "
		end if
	else
	    strORDERStmtAdd = strORDERStmtAdd & " , groupregioncode "
	end if	


    if len(trim(v_txtRegion)) > 0 then 
		if v_txtRegion <> "--SUMMARY--" then
			strORDERStmtAdd = strORDERStmtAdd & " , regioncode "
		end if
	else
	    strORDERStmtAdd = strORDERStmtAdd & " , regioncode "
	end if	


    if len(trim(v_txtcompany)) > 0 then 
		if v_txtcompany <> "--SUMMARY--" then
			strORDERStmtAdd = strORDERStmtAdd & " , companycode "
		end if
	else
	    strORDERStmtAdd = strORDERStmtAdd & " , companycode "
	end if		

	if len(trim(v_txtFcba)) > 0 then 
		if v_txtFcba <> "--SUMMARY--" then
			strORDERStmtAdd = strORDERStmtAdd & " , CUSTOMER_FCBA "
		end if
	else
	    strORDERStmtAdd = strORDERStmtAdd & " , CUSTOMER_FCBA "
	end if		
	
    strORDERStmtAdd = strORDERStmtAdd & " , seqnumber "
	
	

	' Step RENDEMEN UTAMA
	strSTEPStmt = " stepnumber = 'STEP_03' " 	
	strTemp = "SELECT " & strSELECTStmtAdd & strMAINStmtAdd & strFROMStmtAdd & strSTEPStmt & strWHEREStmtAdd & " AND CodeDesc = 'FFB0701' " & strGROUPStmtAdd 
	strSQL = "SELECT " & strSELECTStmtAdd & strMAIN2StmtAdd & " from ( " & strTemp & " ) " 

	
	strRENDEMENstmtAdd =   ", sum(hi) tbs_process_hi, " & _
				" sum(shi) tbs_process_shi, " & _
				" sum(sbi) tbs_process_sbi, " & _
				" sum(budget_hi) tbs_process_budget_hi,  " & _
				" sum(budget_shi) tbs_process_budget_shi,  " & _
				" sum(budget_sbi) tbs_process_budget_sbi   " 	
	strRENDEMENstmtAdd = strRENDEMENstmtAdd &  ", 0 cpo_production_hi, " & _
				" 0 cpo_production_shi, " & _
				" 0 cpo_production_sbi, " & _
				" 0 cpo_production_budget_hi,  " & _
				" 0 cpo_production_budget_shi,  " & _
				" 0 cpo_production_budget_sbi   " 	
	strRENDEMENstmtAdd =  strRENDEMENstmtAdd & ", 0 kernel_production_hi, " & _
				" 0 kernel_production_shi, " & _
				" 0 kernel_production_sbi, " & _
				" 0 kernel_production_budget_hi,  " & _
				" 0 kernel_production_budget_shi,  " & _
				" 0 kernel_production_budget_sbi   " 	


	strRendemenAdd = "SELECT 'RENDEMEN' title " & strSELECTStmtAdd_Rendemen & strRENDEMENstmtAdd & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN_DETAIL' " & strGROUPRENDEMENStmtAdd
	
	'======================CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_06' " 	
	strTemp = "SELECT " & strSELECTStmtAdd & strMAINStmtAdd & strFROMStmtAdd & strSTEPStmt & strWHEREStmtAdd & " AND CodeDesc = 'PROD01' " & strGROUPStmtAdd 
	strSQL = "SELECT " & strSELECTStmtAdd & strMAIN2StmtAdd & " from ( " & strTemp & strORDERStmtAdd & " ) " 

	strRENDEMENstmtAdd =   ", 0 tbs_process_hi, " & _
				" 0 tbs_process_shi, " & _
				" 0 tbs_process_sbi, " & _
				" 0 tbs_process_budget_hi,  " & _
				" 0 tbs_process_budget_shi,  " & _
				" 0 tbs_process_budget_sbi   " 	
	strRENDEMENstmtAdd = strRENDEMENstmtAdd &  ", sum(hi) cpo_production_hi, " & _
				" sum(shi) cpo_production_shi, " & _
				" sum(sbi) cpo_production_sbi, " & _
				" sum(budget_hi) cpo_production_budget_hi,  " & _
				" sum(budget_shi) cpo_production_budget_shi,  " & _
				" sum(budget_sbi) cpo_production_budget_sbi   " 	
	strRENDEMENstmtAdd =  strRENDEMENstmtAdd & ", 0 kernel_production_hi, " & _
				" 0 kernel_production_shi, " & _
				" 0 kernel_production_sbi, " & _
				" 0 kernel_production_budget_hi,  " & _
				" 0 kernel_production_budget_shi,  " & _
				" 0 kernel_production_budget_sbi   " 	

	strRendemenAdd = strRendemenAdd & " UNION SELECT 'RENDEMEN' title " & strSELECTStmtAdd_Rendemen & strRENDEMENstmtAdd & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN' " & strGROUPRENDEMENStmtAdd
	
	'======================CONTINUE
	
	strSTEPStmt = " stepnumber = 'STEP_06' " 	
	strTemp = "SELECT " & strSELECTStmtAdd & strMAINStmtAdd & strFROMStmtAdd & strSTEPStmt & strWHEREStmtAdd & " AND CodeDesc = 'PROD02' " & strGROUPStmtAdd 
	strSQL = "SELECT " & strSELECTStmtAdd & strMAIN2StmtAdd & " from ( " & strTemp & strORDERStmtAdd & " ) " 

	strRENDEMENstmtAdd =   ", 0 tbs_process_hi, " & _
				" 0 tbs_process_shi, " & _
				" 0 tbs_process_sbi, " & _
				" 0 tbs_process_budget_hi,  " & _
				" 0 tbs_process_budget_shi,  " & _
				" 0 tbs_process_budget_sbi   " 	
	strRENDEMENstmtAdd = strRENDEMENstmtAdd &  ", 0 cpo_production_hi, " & _
				" 0 cpo_production_shi, " & _
				" 0 cpo_production_sbi, " & _
				" 0 cpo_production_budget_hi,  " & _
				" 0 cpo_production_budget_shi,  " & _
				" 0 cpo_production_budget_sbi   " 	
	strRENDEMENstmtAdd =  strRENDEMENstmtAdd & ", sum(hi) kernel_production_hi, " & _
				" sum(shi) kernel_production_shi, " & _
				" sum(sbi) kernel_production_sbi, " & _
				" sum(budget_hi) kernel_production_budget_hi,  " & _
				" sum(budget_shi) kernel_production_budget_shi,  " & _
				" sum(budget_sbi) kernel_production_budget_sbi   " 	

	strRendemenAdd = strRendemenAdd & " UNION SELECT 'RENDEMEN' title " & strSELECTStmtAdd_Rendemen & strRENDEMENstmtAdd & " FROM ( " & strSQL & " ) " & _
					" GROUP BY 'RENDEMEN' " & strGROUPRENDEMENStmtAdd
	
	
	
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
						   
	strRendemenAddFinal = "SELECT title HEADGROUP, '%' UnitOfMeasurement, '' CodeDesc " & strSELECTStmtAdd_Rendemen & strRENDEMENFinalStmt & " FROM ( " & strRendemenAdd & " ) " & _
	                   " GROUP BY title " & strGROUPRENDEMENStmtAdd
	
	
	

	Set rsRendemenFinalAdd = Server.CreateObject("ADODB.Recordset")
	response.write strRendemenAddFinal
	rsRendemenFinalAdd.Open strRendemenAddFinal, connstr, 3,3	

	
	' *************************SPECIAL FOR RENDEMEN************************	
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
%>