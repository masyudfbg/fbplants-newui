<%
	'Amandemen
	'1. Allowded Code Untuk Premi Basis Tunas = A57
	'1. Allowded Code Untuk Premi Basis Lebih Tunas = A57A	
	'By NAF
	
	

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Tunas.asp"
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
	
	
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select nvl(RATE_1,0) as Rate_1, includedpayrollprocessing " & _
				" From ALLOWDED where FCCODE='A57'  " & _
				" and man_auto = 'A' " & _
				" and fcBA='" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		' nRatePremiTunas will be inputted by 100 (percent), if 100%
		nRatePremiBasis= cdbl(rsParameter("Rate_1")) 
		tempIncludePayrollProcessingBasis = rsParameter("includedpayrollprocessing")
	else
		nRatePremiBasis = 0 
		tempIncludePayrollProcessingBasis = "FALSE"		 
	End if
	rsParameter.Close
	set rsParameter = Nothing
	
    Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select nvl(RATE_1,0) as Rate_1, includedpayrollprocessing " & _
				" From ALLOWDED where FCCODE='A57A'  " & _
				" and man_auto = 'A' " & _
				" and fcBA='" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		' nRatePremiTunas will be inputted by 100 (percent), if 100%
		nRatePremiLebihBasis= cdbl(rsParameter("Rate_1")) 
		tempIncludePayrollProcessingLebihBasis = rsParameter("includedpayrollprocessing")
	else
		nRatePremiLebihBasis = 0 
		tempIncludePayrollProcessingLebihBasis = "FALSE"		 
	End if
	rsParameter.Close
	set rsParameter = Nothing
	
	
	  
	if tempIncludePayrollProcessingBasis = "TRUE" then 
	 
		qReceive = " 	SELECT " & _
				   "			   fcba," & _
				   " 			   fddate," & _
				   "			   employeecode," & _
				   "			   firstreference reference," & _
				   " 			   SUM (totalUnitPerDay) totalUnitPerDay" & _
				   "			FROM" & _
				   " 			   (	SELECT " & _
				   " 					   fcba, " & _
				   "					   fddate, " & _
				   "					   employeecode, " & _
				   " 					   reference, " & _
				   " 					   oppremirate_api.getFirstReference (fcba, fddate, employeecode, 'A57') firstReference, " & _
				   "					   oppremirate_api.convertQty (fcba, " & _
			 	   "												   fddate, " & _
				   " 												   employeecode, " & _
				   "												   reference, " & _
				   "		                                            'A57','TUNAS'," & _
				   "												   TotalUnitPerDay) " & _
				   "						   AS TotalUnitPerDay " & _
				   "					FROM " & _
				   "					   (	SELECT " & _
				   "							   gad.fcba, " & _
				   "							   gad.FDDATE, " & _
				   "							   gad.EMPLOYEECODE, " & _
				   "							   gad.reference reference, " & _
				   "							   NVL (SUM (gad.unit), 0) TotalUnitPerDay " & _
				   "							FROM " & _
				   "							   Employee emp, ATTENDANCE_GAD gad " & _
				   "							WHERE " & _
				   "							   (emp.fccode = gad.EMPLOYEECODE AND " & _
				   "								emp.fcba = gad.fcba) AND " & _
				   "							   gad.attendance IN ('KJ', 'WH', 'WS') AND " & _
				   "							   DATETERMINATE IS NULL AND " & _
				   "							   process_posting_api.GetPostingValue ('PAYROLL', " & _
				   "																	'A57', " & _
				   "																	'JOB', " & _
				   "																	gad.JOBCODE, " & _
				   "																	'" & Session("fcBA") & "') = 'TRUE' AND " & _
				   "							   gad.unit > 0 AND " & _
				   "							   TO_NUMBER (TO_CHAR (gad.FDDATE, 'mm')) = " & Month(Session("v_Employee")) & " AND " & _
				   "							   TO_NUMBER (TO_CHAR (gad.FDDATE, 'yyyy')) = " & Year(Session("v_Employee")) & " AND " & _
				   "							   gad.fcBA = '" & Session("fcBA") & "' " & _
				   "							GROUP BY " & _
				   "							   gad.fcba, " & _
				   "							   gad.FDDATE, " & _
				   "							   gad.EMPLOYEECODE, " & _
				   "							   gad.reference) " & _
				   "					GROUP BY " & _
				   "					   fcba, " & _
				   "					   fddate, " & _
				   "					   employeecode, " & _
				   "					   reference, " & _
				   "					   oppremirate_api.getFirstReference (fcba, fddate, employeecode,'A57'), " & _
				   "									  oppremirate_api.convertQty (fcba, " & _
				   "												   fddate, " & _
				   "												   employeecode, " & _
				   "												   reference, " & _
				   "											       'A57','TUNAS'," & _
				   "												   TotalUnitPerDay)) " & _
				   "			GROUP BY " & _
				   "			   fcba, " & _
				   "			   fddate, " & _
				   "			   employeecode, " & _
				   "			   firstreference		 "  
				 

		    
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Premi Tunas ..."& "<br>" 
		response.Flush()
		'response.write qReceive
		rsEmpAllDed.Open qReceive, connstr
		if not rsEmpAllDed.eof then
		    Do While Not rsEmpAllDed.eof
		   
				Set rsOpPremiRate = Server.CreateObject("ADODB.Recordset")
				qParameter = " select basis, basisrate, extrabasisrate from oppremirate " & _
							  " where " & _
							  " premitypecode ='TUNAS' and " & _
							  " fccode = '" & rsEmpAllDed("REFERENCE") & "' and " & _
							  " fcba = '" & Session("fcBA") & "' " 
							
				'response.Write " qParameter = " & qParameter				  
				rsOpPremiRate.CursorType = 3
				rsOpPremiRate.CursorLocation = 3
				rsOpPremiRate.Open qParameter, connstr
				if not rsOpPremiRate.eof then
				    ntempBasis = cdbl(rsOpPremiRate("basis"))
					ntempBasisRate = cdbl(rsOpPremiRate("basisrate"))
					ntempExtraBasisRate = cdbl(rsOpPremiRate("extrabasisrate"))
				else
				    ntempBasis = 0
					ntempBasisRate = 0
					ntempExtraBasisRate = 0
				End if
				rsOpPremiRate.Close
				set rsOpPremiRate = Nothing		   
				
				' Get Total TREE/ ambil jumlah Pokok-nya/Pohon yang di tunas oleh Employee
				tempTotalUnitPerDay = cdbl(rsEmpAllDed ("TotalUnitPerDay"))
				
				if tempTotalUnitPerDay >= ntempBasis then 
				    ' the Employee will have the BASIS.
		   
					maxLineNo = maxLineNo + 1
					nPremiBasis = ntempBasisRate * nRatePremiBasis / 100
					qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
							"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A57" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nPremiBasis & _
							",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
							
					connstr.execute(qAllowance)
					
					
					if tempIncludePayrollProcessingLebihBasis = "TRUE" then 
					    'the Employee will have extra basis 
						tempTotalLebihBasis = tempTotalUnitPerDay - ntempBasis
						nPremiLebihBasis = tempTotalLebihBasis * ntempExtraBasisRate * nRatePremiBasis  / 100
						
						if nPremiLebihBasis <> 0 then 
							
							maxLineNo = maxLineNo + 1
							qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
									"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A57A" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nPremiLebihBasis & _
									",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
									
							connstr.execute(qAllowance)
							maxLineNoLebih = maxLineNo 
							
						end if
					
					end if					
					
					 
					qReceiveDetail =" select gad.FDDATE, " & _
											 " gad.EMPLOYEECODE, " & _
											 " gad.JOBCODE, " & _
											 " gad.LOCATIONTYPE, " & _
											 " gad.LOCATIONCODE, " & _
											 " nvl(gad.unit,0) TotalUnitPerDay " & _
									" From Employee emp,ATTENDANCE_GAD gad Where " & _
									" (emp.fccode = gad.EMPLOYEECODE and emp.fcba = gad.fcba) " & _
									" and gad.attendance in ('KJ','WH','WS') " & _
									" and DATETERMINATE is null " & _
									" and process_posting_api.GetPostingValue('PAYROLL','A57','JOB',gad.JOBCODE,'" & Session("fcBA") & "')= 'TRUE' " & _
									" and gad.unit > 0 " & _
									" and gad.fcBA='" & Session("fcBA") & "'" & _
									" and gad.reference = '" & rsEmpAllDed("reference") & "'" & _
									" and to_number(to_Char(gad.FDDATE,'YYYYMMDD'))= to_number(to_char(to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy'),'YYYYMMDD')) " & _
									" and gad.EMPLOYEECODE = '" &  rsEmpAllDed("EMPLOYEECODE") & "' " 
									
					Set rsReceive = Server.CreateObject("ADODB.Recordset")
					rsReceive.CursorType = 3
					rsReceive.CursorLocation = 3

					rsReceive.Open qReceiveDetail, connstr
					if not rsReceive.eof then
						detailLineNo = 0
						Do While Not rsReceive.eof	

							' Insert the History of EmpallDed Data 
							' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
							
							txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
														 rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
														 rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
														 rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
														 rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
														 Session("fcBA")
														 
														 
							nUnitDetail = cdbl(rsReceive("TotalUnitPerDay"))
							nRateDetail = nRatePremiBasis / 100
							nAmountDetail = nUnitDetail * nPremiBasis / tempTotalUnitPerDay 
							
							' Insert history for BASIS
							detailLineNo = detailLineNo + 1
							qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
																" LINEHEADERNO," & _
																" LINENO," & _
																" FCBA," & _
																" REF_PROCESS," & _
																" REF_PROCESSNAME," & _
																" REF_TABLESOURCE," & _
																" REF_NUMBERKEYS," & _
																" REF_STRINGSEPARATOR," & _
																" REF_STRINGKEYS," & _
																" REF_STRINGVALUES," & _
																" UNIT, " & _
																" RATE, " & _
																" AMOUNT, " & _
																" FCENTRY," & _
																" FCEDIT," & _
																" FCIP," & _
																" LASTUPDATE," & _
																" LASTTIME)" & _
														" VALUES ( " & _
																" '" & maxLineNo & "'," & _
																" '" & detailLineNo & "'," & _
																" '" & Session("fcBA") & "'," & _
																" '" & tempProcessCode & "'," & _
																" '" & tempProcessNameCode & "'," & _
																" '" & tempRefCodeReceive & "'," & _
																" '" & txtRefNumberKeys_Receive & "'," & _
																" '" & txtRefStringSeparator_Receive & "'," & _
																" '" & txtRefStringKeys_Receive & "'," & _
																" '" & txtRefStringValues_Receive & "'," & _
																nUnitDetail & "," & _
																nRateDetail & "," & _
																nAmountDetail & "," & _ 
																" '" & Session("CodeName") & "'," & _
																" '" & Session("CodeName") & "'," & _
																" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
																" sysdate, " & _
																" to_char(sysdate, 'HH24:MI') " & _
																" ) "					
							connstr.execute(qAllowanceHistory)
							
							' insert history for BASIS LEBIH
							
							if tempIncludePayrollProcessingLebihBasis = "TRUE" then 

								if nPremiLebihBasis <> 0 then 
									detailLineNo = detailLineNo + 1
									qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
																		" LINEHEADERNO," & _
																		" LINENO," & _
																		" FCBA," & _
																		" REF_PROCESS," & _
																		" REF_PROCESSNAME," & _
																		" REF_TABLESOURCE," & _
																		" REF_NUMBERKEYS," & _
																		" REF_STRINGSEPARATOR," & _
																		" REF_STRINGKEYS," & _
																		" REF_STRINGVALUES," & _
																		" UNIT, " & _
																		" RATE, " & _
																		" AMOUNT, " & _
																		" FCENTRY," & _
																		" FCEDIT," & _
																		" FCIP," & _
																		" LASTUPDATE," & _
																		" LASTTIME)" & _
																" VALUES ( " & _
																		" '" & maxLineNoLebih & "'," & _
																		" '" & detailLineNo & "'," & _
																		" '" & Session("fcBA") & "'," & _
																		" '" & tempProcessCode & "'," & _
																		" '" & tempProcessNameCode & "'," & _
																		" '" & tempRefCodeReceive & "'," & _
																		" '" & txtRefNumberKeys_Receive & "'," & _
																		" '" & txtRefStringSeparator_Receive & "'," & _
																		" '" & txtRefStringKeys_Receive & "'," & _
																		" '" & txtRefStringValues_Receive & "'," & _
																		nUnitDetail & "," & _
																		nRateDetail & "," & _
																		nAmountDetail & "," & _ 
																		" '" & Session("CodeName") & "'," & _
																		" '" & Session("CodeName") & "'," & _
																		" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
																		" sysdate, " & _
																		" to_char(sysdate, 'HH24:MI') " & _
																		" ) "					
									connstr.execute(qAllowanceHistory)

								end if
							end if			
							
							rsReceive.MoveNext
						loop
					end if
					rsReceive.Close
					set rsReceive = Nothing		

				end if
				rsEmpAllDed.MoveNext
			loop
		End if
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
	end if
%>