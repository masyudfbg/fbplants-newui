<%
	
	' Alowded Premi Semprot Mandor = A47 
	' ketentuan Supervisi_2 di GAD dengan minimial jumlah anggota yg bekerja hari itu 10 orang
	' permi yg di dapat 10 % dari total premi yg di dapat dari anggotanya
    ' by Akbar, 24 Okt 2013
	
	Dim TempPremiMandor
	
	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Semprot_Mandor.asp"
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
	
	
	
	qBA="select nvl(Rate_1,0) as Rate_1, nvl(Rate_2,0) as Rate_2, includedpayrollprocessing " & _
        " from ALLOWDED where FCCODE='A47' and fcBA='" & Session("fcBA") & "'"
		
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   Min_Anggota=cdbl(rsBA("Rate_1")) 
	   Rate_Mandore=cdbl(rsBA("Rate_2"))
	   tempIncludePayrollProcessing = rsBA("includedpayrollprocessing")
	   
	else
	   Rate_Mandore=0
	   Min_Anggota=0
	   tempIncludePayrollProcessing = "FALSE"
	   
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY12001 Rate Mandore 1 A47 is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
    
    qBA="select nvl(Rate_1,0) as Rate_1 from ALLOWDED where FCCODE='A46' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   Rate_Anggota=cdbl(rsBA("Rate_1"))
	else
	   Rate_Anggota=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY12001 Rate Mandore 1 A46 is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
	
	
	if tempIncludePayrollProcessing = "TRUE" then 
	
	
		qSupervisi = 	" SELECT fddate, gangcode, supervision_2, to_number(trim(countEmp)) countEmp, totalHK FROM " & _
						   " (	SELECT " & _
							" 	   gad.FDDATE, " & _
							" 	   gad.GangCode, " & _
							" 	   gad.supervision_2, " & _
							" 	   COUNT (gad.supervision_2) countEmp, " & _
							" 	   SUM (nvl(gad.Mandays,0)) totalHK " & _
							" 	FROM " & _
							" 	   Employee emp, ATTENDANCE_GAD gad " & _
							" 	WHERE " & _
							" 	   (emp.fccode = gad.EMPLOYEECODE AND " & _
							" 		emp.fcba = gad.fcba) AND " & _
							" 	   gad.attendance IN ('KJ', 'WH', 'WS') AND " & _
							" 	   DATETERMINATE IS NULL AND " & _
							" 	   process_posting_api.GetPostingValue ('PAYROLL', " & _
							" 											'A47', " & _
							" 											'JOB', " & _
							" 											gad.JOBCODE, " & _
							" 											'" & Session("fcBA") & "') = 'TRUE' AND " & _
							" 	   gad.mandays > 0 AND " & _
							"	   trim(gad.supervision_2) IS NOT NULL AND " & _
							"	   TO_NUMBER (TO_CHAR (gad.FDDATE, 'mm')) = " & Month(Session("v_Employee")) &  _
							" 		  and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
							" 		  and gad.fcBA='" & Session("fcBA") & "' " & _
							" 		  group by gad.FDDATE,gad.GangCode, gad.supervision_2 " & _
							" 		  )  " & _
							" where countEmp >= " & Min_Anggota & " " 

		'response.write 	qSupervisi	
		Set rsSupervisi = Server.CreateObject("ADODB.Recordset")
		rsSupervisi.CursorType = 3
		rsSupervisi.CursorLocation = 3
		response.Write "Calculate Premi Mandore Semprot ..."  & "<br>" 
		response.Flush()
		rsSupervisi.Open qSupervisi, connstr
		if not rsSupervisi.eof then
		   ' Loop Supervisi
		   Do While Not rsSupervisi.eof
			    ' By NAF
			    ' nPremiMandore = 10% * ( jumlahAnggota * Rate_A46 ) 
				
				'response.write Rate_Mandore
				'response.write rsSupervisi("countEmp") 
				'response.write Rate_Anggota


			    tempPremiMandor = cdbl(Rate_Mandore) * (  cdbl(Rate_Anggota) * cdbl( rsSupervisi("totalHK")))
			
				
				
				maxLineNo = maxLineNo + 1
				qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _ 
						" values  " & _
						"(" & maxLineNo & ",'" & _ 
						rsSupervisi("SUPERVISION_2") & _ 
						"','A47'," & _
						" to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') " & _
						"," & TempPremiMandor & _
						",'" & Session("CodeName") & _
						"','" & Session("CodeName") & _
						"','" & _
						Request.ServerVariables("REMOTE_ADDR") & _
						"','" & Session("fcBA") & _
						"', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"

				connstr.execute(qAllowance)
				
				' Remarks History By Iwa : 08/01/2019
			    'XXXXXXXXXXXXX
				' Detail History
				
			    'qSupervisiDetail = " SELECT " & _
				'					" 	   gad.FDDATE, " & _
				'					" 	   gad.EMPLOYEECODE, " & _
				'					" 	   gad.JOBCODE, " & _
				'					" 	   gad.LOCATIONTYPE, " & _
				'					" 	   gad.LOCATIONCODE, " & _
				'					" 	   gad.FDDATE, " & _
				'					" 	   nvl(gad.Mandays,0) totalHK " & _
				'					" 	FROM " & _
				'					" 	   Employee emp, ATTENDANCE_GAD gad " & _
				'					" 	WHERE " & _
				'					" 	   (emp.fccode = gad.EMPLOYEECODE AND " & _
				'					" 		emp.fcba = gad.fcba) AND " & _
				'					" 	   gad.attendance IN ('KJ', 'WH', 'WS') AND " & _
				'					" 	   DATETERMINATE IS NULL AND " & _
				'					" 	   process_posting_api.GetPostingValue ('PAYROLL', " & _
				'					" 											'A47', " & _
				'					" 											'JOB', " & _
				'					" 											gad.JOBCODE, " & _
				'					" 											'" & Session("fcBA") & "') = 'TRUE' AND " & _
				'					" 	   gad.mandays > 0 AND " & _
				'					"	   trim(gad.supervision_2) IS NOT NULL  " & _
				'					" 	   and to_number(to_Char(gad.FDDATE,'YYYYMMDD'))= to_number(to_char(to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy'),'YYYYMMDD')) " & _
				'					" 	   and gad.fcBA='" & Session("fcBA") & "' " & _
				'					" 	   and gad.Supervision_2 = '" & rsSupervisi("SUPERVISION_2") & "'" & _
				'					" 	   and gad.GANGCODE='" & rsSupervisi("GANGCODE") & "'" 
				'								
					
				'Set rsReceive = Server.CreateObject("ADODB.Recordset")
				'rsReceive.CursorType = 3
				'rsReceive.CursorLocation = 3

				'response.write qSupervisiDetail
				'rsReceive.Open qSupervisiDetail, connstr
						
				'if not rsReceive.eof then
				'	' Loop 1
				'	Do While Not rsReceive.eof	
				'		
				'		' Insert the History of EmpallDed Data 
				'		' We add NEW reference string
				'		
				'		txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
				'									 rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
				'									 rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
				'									 rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
				'									 rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
				'									 Session("fcBA")
				'									 
				'		nRateDetail = cdbl(Rate_Anggota) * cdbl(Rate_Mandore)
				'		nUnitDetail = cdbl(rsReceive("totalHK"))
				'		nAmountDetail = nRateDetail * nUnitDetail
				'									 
				'		
				'		detailLineNo = detailLineNo + 1
				'		qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
				'											" LINEHEADERNO," & _
				'											" LINENO," & _
				'											" FCBA," & _
				'											" REF_PROCESS," & _
				'											" REF_PROCESSNAME," & _
				'											" REF_TABLESOURCE," & _
				'											" REF_NUMBERKEYS," & _
				'											" REF_STRINGSEPARATOR," & _
				'											" REF_STRINGKEYS," & _
				'											" REF_STRINGVALUES," & _
				'											" UNIT, " & _
				'											" RATE, " & _
				'											" AMOUNT, " & _
				'											" FCENTRY," & _
				'											" FCEDIT," & _
				'											" FCIP," & _
				'											" LASTUPDATE," & _
				'											" LASTTIME)" & _
				'									" VALUES ( " & _
				'											" '" & maxLineNo & "'," & _
				'											" '" & detailLineNo & "'," & _
				'											" '" & Session("fcBA") & "'," & _
				'											" '" & tempProcessCode & "'," & _
				'											" '" & tempProcessNameCode & "'," & _
				'											" '" & tempRefCodeReceive & "'," & _
				'											" '" & txtRefNumberKeys_Receive & "'," & _
				'											" '" & txtRefStringSeparator_Receive & "'," & _
				'											" '" & txtRefStringKeys_Receive & "'," & _
				'											" '" & txtRefStringValues_Receive & "'," & _
				'											nUnitDetail & "," & _
				'											nRateDetail & "," & _
				'											nAmountDetail & "," & _ 
				'											" '" & Session("CodeName") & "'," & _
				'											" '" & Session("CodeName") & "'," & _
				'											" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
				'											" sysdate, " & _
				'											" to_char(sysdate, 'HH24:MI') " & _
				'											" ) "					
				'		connstr.execute(qAllowanceHistory)
				'		rsReceive.MoveNext
				'		tempCounter = tempCounter + 1
				'	loop
					' End loop 1
				'end if
				'rsReceive.Close
				'set rsReceive = Nothing	
			    'XXXXXXXXXXXXX
				 rsSupervisi.MoveNext
			loop ' End Loop 2
		End if
		rsSupervisi.Close
		set rsSupervisi = Nothing
	end if
	
%>