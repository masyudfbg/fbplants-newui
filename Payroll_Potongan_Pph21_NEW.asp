<%
	'Amandemen
	'1. Allowded Code Untuk Potongan PPh 21 = D03 
	' By NAF
	
	
	

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Potongan_Pph21.asp"
	tempRefCodeReceive = "ATTENDANCE_GAD"

	qReference = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
				  " from process_tableref_tab " & _
				  " where " & _
				  " processcode = '" & tempProcessCode & "' and " & _
				  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
				  " fccode =  '" & tempRefCodeReceive & "'"

	Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
	rsReferenceReceive.CursorType = 3
	rsReferenceReceive.CursorLocation = 3
	rsReferenceReceive.Open qReference, connstr

	doneReceive = 1
	if not rsReferenceReceive.eof then
		txtRefNumberKeys_Receive = rsReferenceReceive("REF_NUMBERKEYS")
		txtRefStringSeparator_Receive = rsReferenceReceive("REF_STRINGSEPARATOR")
		txtRefStringKeys_Receive = rsReferenceReceive("REF_STRINGKEYS")
		doneReceive = 1
	else
		doneReceive = 0
	end if
	rsReferenceReceive.close
	set rsReferenceReceive = Nothing
	' End of Picking Referential Integrity
	
	
	' System must track it as an AUTOMATIC Process 
	tempIncludedPayrollProcessing = "FALSE"
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select includedpayrollprocessing, man_auto " & _
		              " From ALLOWDED where FCCODE='D03'  " & _
					  " and fcBA='" & Session("fcBA") & "' " & _
					  " and man_auto = 'A'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		 tempIncludedPayrollProcessing= rsParameter("includedpayrollprocessing")
	else
		 tempIncludedPayrollProcessing = "FALSE"
	End if
	rsParameter.Close
	set rsParameter = Nothing
	


	' System must track it as an AUTOMATIC Process 
	' Jabatan
	tempIncludedPayrollProcessingA = "FALSE"
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select includedpayrollprocessing, man_auto, rate_1 " & _
		              " From ALLOWDED where FCCODE='D03A'  " & _
					  " and fcBA='" & Session("fcBA") & "' " & _
					  " and man_auto = 'A'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		 tempIncludedPayrollProcessingA= rsParameter("includedpayrollprocessing")
		 tempPercentageJabatanDeduction= cdbl(rsParameter("rate_1"))
	else
		 tempIncludedPayrollProcessingA = "FALSE"
		 tempPercentageJabatanDeduction= 0
	End if
	rsParameter.Close
	set rsParameter = Nothing

	' System must track it as an AUTOMATIC Process 
	' Astec
	tempIncludedPayrollProcessingB = "FALSE"
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select includedpayrollprocessing, man_auto, rate_1 " & _
		              " From ALLOWDED where FCCODE='D03B'  " & _
					  " and fcBA='" & Session("fcBA") & "' " & _
					  " and man_auto = 'A'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		 tempIncludedPayrollProcessingB= rsParameter("includedpayrollprocessing")
		 tempPercentageAstekDeduction= cdbl(rsParameter("rate_1"))
	else
		 tempIncludedPayrollProcessingB = "FALSE"
		 tempPercentageAstekDeduction= 0
	End if
	rsParameter.Close
	set rsParameter = Nothing


	' MaxTaxLineNo from TaxAllded
    maxTaxLineNo = 0
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
		              " From TAXALLDED where  fcBA='" & Session("fcBA") & "' " 
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		 maxTaxLineNo = cdbl(rsParameter("maxLineNo"))
	else
		 maxTaxLineNo = 0
	End if
	rsParameter.Close
	set rsParameter = Nothing


	
	
	if tempIncludedPayrollProcessing = "TRUE" then 
	

	
		qQuery = " select fcba, " & _
				 " last_day(to_date('" & Month(Session("v_Employee")) & "/01/" & Year(Session("v_Employee")) & "','mm/dd/yyyy')) fddate," & _
				 " empcode employeecode, maritalstatus, IncomeBeforeTax, " & _
				 " untaxableincome_api.getTotalUti(maritalstatus, fcba) totalUti " & _
						" from ( " & _
						"  SELECT a.fcba,  a.empcode, b.taxstatus maritalStatus, SUM ( Allowded_APi.GetAlldedValue (a.allowdedcode, amount, a.fcba)) IncomeBeforeTax " & _
						" 	FROM empallded a, employee b " & _
						"    WHERE	 a.empcode = b.fccode " & _
						" 		 AND TO_NUMBER (TO_CHAR (a.adhocdate, 'YYYY')) = " & Year(Session("v_Employee")) & _
						" 		 AND TO_NUMBER (TO_CHAR (a.adhocdate, 'MM')) = " & Month(Session("v_Employee")) &  _
						" 		 AND allowded_api.GetIncludedTaxCalculation (A.ALLOWDEDCODE, a.fcba) = 'TRUE' " & _
						" 		 AND employee_api.GetIncludedTaxCalculation (a.empcode, a.fcba) = 'TRUE' " & _
						"        AND b.dateterminate is null " & _
						" 		 AND a.fcba = '" & Session("fcBA") & "' " & _
						" 		 AND a.fcba = b.fcba " & _
						" GROUP BY a.fcba, a.empcode, b.taxstatus " & _
						" ) " 

		
		'response.Write "<br><br>" & " Recordset " & qReceive & "<br><br>"
		'response.Flush()			    
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Tax for PPh 21 ..."& "<br>" 
		response.Flush()
		rsEmpAllDed.Open qQuery, connstr
		if not rsEmpAllDed.eof then
		   Do While Not rsEmpAllDed.eof
		   
				tempIncomeBeforeTax = cdbl(rsEmpAllDed ("IncomeBeforeTax"))
				
				
				if tempIncludedPayrollProcessingA = "TRUE" then 
				    tempJabatanValue = -1 * tempIncomeBeforeTax * tempPercentageJabatanDeduction / 100
					
					maxTaxLineNo = maxTaxLineNo + 1
					qAllowance="Insert Into  TaxAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
							"(" & maxTaxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D03A" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & tempJabatanValue & _
							",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
							
					connstr.execute(qAllowance)
					
				else
				
					tempJabatanValue = 0
				end if
				

				
				strQueryData = " select sum(amount) totalAmount from empallded " & _
								" where " & _
								" 		 AND TO_NUMBER (TO_CHAR (adhocdate, 'YYYY')) = " & Year(Session("v_Employee")) & _
								" 		 AND TO_NUMBER (TO_CHAR (adhocdate, 'MM')) = " & Month(Session("v_Employee")) &  _
								"        AND allowdedcode in ('A17','D02') " & _
								" 		 AND a.fcba = '" & Session("fcBA") & "' " 
				
				totalAstekDeduction = 0
				
				Set rsQueryData = Server.CreateObject("ADODB.Recordset")
				rsQueryData.CursorType = 3
				rsQueryData.CursorLocation = 3
				rsQueryData.Open strQueryData, connstr
				if not rsQueryData.eof then
					totalAstekDeduction = cdbl(rsQueryData("totalAmount"))
				else 
					totalAstekDeduction = 0
				end if
				
				
				
				
				
				if tempIncludedPayrollProcessingB = "TRUE" then 
				    tempAstekValue = totalAstekDeduction
					
					maxTaxLineNo = maxTaxLineNo + 1
					qAllowance="Insert Into  TaxAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
							"(" & maxTaxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D03B" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & tempAstekValue & _
							",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
							
					connstr.execute(qAllowance)
					
				else
					tempAstekValue = 0
				
				end if
				
				
				
				tempTotalUti = cdbl(rsEmpAllDed ("totalUti"))
				
				tempDeduction = tempJabatanValue + tempAstekValue 
				
				tempNetIncome = tempIncomeBeforeTax + tempDeduction
				
				'tempIncomeIncludedTax = tempNetIncome - tempTotalUti 
				tempIncomeIncludedTax = tempNetIncome 
				
				if tempIncomeIncludedTax >= 0 then 
					tempYearlyIncomeIncludedTax = tempIncomeIncludedTax * 12
					tempYearlyTotalUti = tempTotalUti * 12
					
					qGetPercentage = " select RANGETAXGROUP_API.GetValueProgressive(" & tempYearlyIncomeIncludedTax & _
									 "," & tempYearlyTotalUti & ",'" & Session("fcBA") & "') totalPercentage from dual"
									 
					Set rsGetPercentage = Server.CreateObject("ADODB.Recordset")
					rsGetPercentage.CursorType = 3
					rsGetPercentage.CursorLocation = 3
					rsGetPercentage.Open qGetPercentage, connstr
					if not rsGetPercentage.eof then
						if rsGetPercentage ("totalPercentage") <> "" then 
							tempTotalPercentage = cdbl(rsGetPercentage ("totalPercentage"))				
						else
							tempTotalPercentage = 0
						end if
					else
						tempTotalPercentage = 0
					end if
					
					tempFinalTax = -1 * tempTotalPercentage 
				else
				    tempFinalTax = 0
				end if
				
				' response.write " tempIncomeBeforeTax : " & tempIncomeBeforeTax 
				' response.write " tempTotalUti : " & tempTotalUti
				' response.write " tempYearlyIncomeIncludedTax : " & tempYearlyIncomeIncludedTax
				' response.write " tempTotalPercentage : " & tempTotalPercentage 
				' response.write " tempIncomeIncludedTax : " & tempIncomeIncludedTax 
				' response.write " tempFinalTax : " & tempFinalTax 
				
				
				if tempFinalTax <> 0 then 



				
					maxLineNo = maxLineNo + 1
					qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
							"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D03" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & tempFinalTax & _
							",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
							
					connstr.execute(qAllowance)
					
					
					
					qLedger = " Insert into taxalldedledger " & _
								" ( " & _
								" EMPCODE, YEAR, MONTH, GROSSINCOME, TAXDEDUCTION, NETINCOME, UTI_VALUE, ITX_VALUE, TAXPERCENTAGE_VALUE, TAX_VALUE, " & _
								" FCENTRY, FCEDIT, FCIP, FCBA, REMARKS, LASTUPDATE, LASTTIME " & _
								" ) " & _
								" values ( " & _
								"'" & rsEmpAllDed("EMPLOYEECODE") & "'," & Year(Session("v_Employee")) & "," &  Month(Session("v_Employee")) & "," & _
								tempIncomeBeforeTax & "," &  tempDeduction & "," &  tempNetIncome & "," &  tempTotalUti & "," &  tempIncomeIncludedTax & "," &  tempTotalPercentage & "," & tempFinalTax & _
								",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', '', sysdate, to_char(sysdate,'HH24:MI')" & _					
								" ) "
					connstr.execute(qLedger)
				
				end if
				 
				' Sejarahnya di REMARK karena untuk mempercepat Process...
				' by NAF
				' qReceiveDetail ="  SELECT a.fcba,  a.lineno, Allowded_APi.GetAlldedValue (a.allowdedcode, amount, a.fcba) amount " & _
								' " 	FROM empallded a, employee b " & _
								' "    WHERE	 a.empcode = b.fccode " & _
								' " 		 AND TO_NUMBER (TO_CHAR (a.adhocdate, 'YYYY')) = " & Year(Session("v_Employee")) & _
								' " 		 AND TO_NUMBER (TO_CHAR (a.adhocdate, 'MM')) = " & Month(Session("v_Employee")) &  _
								' " 		 AND allowded_api.GetIncludedTaxCalculation (A.ALLOWDEDCODE, a.fcba) = 'TRUE' " & _
								' " 		 AND employee_api.GetIncludedTaxCalculation (a.empcode, a.fcba) = 'TRUE' " & _
								' "        AND b.dateterminate is null " & _								
								' " 		 AND a.fcba = '" & Session("fcBA") & "' " & _
								' " 		 AND a.fcba = b.fcba " & _
								' " 	     AND a.allowdedcode <> 'D03' "

				
								
				' Set rsReceive = Server.CreateObject("ADODB.Recordset")
				' rsReceive.CursorType = 3
				' rsReceive.CursorLocation = 3

				' rsReceive.Open qReceiveDetail, connstr
				' if not rsReceive.eof then
					' detailLineNo = 0
					' Do While Not rsReceive.eof	

						' '	Insert the History of EmpallDed Data 
						' ' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
						
						' txtRefStringValues_Receive = rsReceive("LINENO") & txtRefStringSeparator_Receive & _
													 ' Session("fcBA")
						
						' detailLineNo = detailLineNo + 1
						' qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
															' " LINEHEADERNO," & _
															' " LINENO," & _
															' " FCBA," & _
															' " REF_PROCESS," & _
															' " REF_PROCESSNAME," & _
															' " REF_TABLESOURCE," & _
															' " REF_NUMBERKEYS," & _
															' " REF_STRINGSEPARATOR," & _
															' " REF_STRINGKEYS," & _
															' " REF_STRINGVALUES," & _
															' " OTHRS, " & _
															' " RATE, " & _
															' " AMOUNT, " & _
															' " FCENTRY," & _
															' " FCEDIT," & _
															' " FCIP," & _
															' " LASTUPDATE," & _
															' " LASTTIME)" & _
													' " VALUES ( " & _
															' " '" & maxLineNo & "'," & _
															' " '" & detailLineNo & "'," & _
															' " '" & Session("fcBA") & "'," & _
															' " '" & tempProcessCode & "'," & _
															' " '" & tempProcessNameCode & "'," & _
															' " '" & tempRefCodeReceive & "'," & _
															' " '" & txtRefNumberKeys_Receive & "'," & _
															' " '" & txtRefStringSeparator_Receive & "'," & _
															' " '" & txtRefStringKeys_Receive & "'," & _
															' " '" & txtRefStringValues_Receive & "'," & _
															' " 0," & _
															' " 0," & _
															' cdbl(rsReceive("amount")) & "," & _ 
															' " '" & Session("CodeName") & "'," & _
															' " '" & Session("CodeName") & "'," & _
															' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
															' " sysdate, " & _
															' " to_char(sysdate, 'HH24:MI') " & _
															' " ) "					
						' connstr.execute(qAllowanceHistory)
						' rsReceive.MoveNext
					' loop
				' end if
				' rsReceive.Close
				' set rsReceive = Nothing		
				
				rsEmpAllDed.MoveNext
		  loop
		End if
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
	end if
%>