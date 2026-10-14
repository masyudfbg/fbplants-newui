<%
	'Amandemen
	'1. Allowded Code Untuk Premi Pupuk = A48 
	

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Pupuk.asp"
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
	
	Dim nRatePremiPupuk, nPremiPupuk
	  Set rsParameter = Server.CreateObject("ADODB.Recordset")
		  qParameter =" Select nvl(RATE_1,0) as Rate_1, includedpayrollprocessing " & _
		              " From ALLOWDED where FCCODE='A48'  " & _
					  " and fcBA='" & Session("fcBA") & "'"
	  'response.Write " qParameter = " & qParameter				  
	  rsParameter.CursorType = 3
	  rsParameter.CursorLocation = 3
	  rsParameter.Open qParameter, connstr
	  if not rsParameter.eof then
		 nRatePremiPupuk= cdbl(rsParameter("Rate_1")) 
		 tempIncludePayrollProcessing = rsParameter("includedpayrollprocessing")
	  else
		 nRatePremiPupuk = 0 
		 tempIncludePayrollProcessing = "FALSE"		 
	  End if
	  rsParameter.Close
	  set rsParameter = Nothing
	  
	  
	if tempIncludePayrollProcessing = "TRUE" then 
	 
		qReceive=" select gad.FDDATE, " & _
						  " gad.EMPLOYEECODE, " & _
						  " nvl(sum(gad.mandays),0) as Mandays  " & _
				 " From Employee emp,ATTENDANCE_GAD gad Where  " & _
				 " (emp.fccode = gad.EMPLOYEECODE and emp.fcba = gad.fcba)  " & _
				 " and gad.attendance in ('KJ','WH','WS') " & _
				 " and DATETERMINATE is null " & _
				 " and process_posting_api.GetPostingValue('PAYROLL','A48','JOB',gad.JOBCODE,'" & Session("fcBA") & "')= 'TRUE' " & _
				 " and gad.mandays > 0 " & _
				 " and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
				 " and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
				 " and gad.fcBA='" & Session("fcBA") & "' " & _
				 " group by gad.FDDATE,gad.EMPLOYEECODE "

		    
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Premi Pupuk ..."& "<br>" 
		response.Flush()
		rsEmpAllDed.Open qReceive, connstr
		if not rsEmpAllDed.eof then
		   Do While Not rsEmpAllDed.eof
					maxLineNo = maxLineNo + 1
					nPremiPupuk = nRatePremiPupuk * cdbl(rsEmpAllDed("Mandays"))
					qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
							"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A48" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nPremiPupuk & _
							",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
							
					connstr.execute(qAllowance)
				 
					' Remarks History By Iwa : 08/01/2019					
					'qReceiveDetail =" select gad.FDDATE, " & _
					'						 " gad.EMPLOYEECODE, " & _
					'						 " gad.JOBCODE, " & _
					'						 " gad.LOCATIONTYPE, " & _
					'						 " gad.LOCATIONCODE, " & _
					'						 " nvl(gad.Mandays,0) mandays " & _
					'				" From Employee emp,ATTENDANCE_GAD gad Where " & _
					'				" (emp.fccode = gad.EMPLOYEECODE and emp.fcba = gad.fcba) " & _
					'				" and gad.attendance in ('KJ','WH','WS') " & _
					'				" and DATETERMINATE is null " & _
					'				" and process_posting_api.GetPostingValue('PAYROLL','A48','JOB',gad.JOBCODE,'" & Session("fcBA") & "')= 'TRUE' " & _
					'				" and gad.mandays > 0 " & _
					'				" and gad.fcBA='" & Session("fcBA") & "'" & _
					'				" and to_number(to_Char(gad.FDDATE,'YYYYMMDD'))= to_number(to_char(to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy'),'YYYYMMDD')) " & _
					'				" and gad.EMPLOYEECODE = '" &  rsEmpAllDed("EMPLOYEECODE") & "' " 
									
					'Set rsReceive = Server.CreateObject("ADODB.Recordset")
					'rsReceive.CursorType = 3
					'rsReceive.CursorLocation = 3

					'rsReceive.Open qReceiveDetail, connstr
					'if not rsReceive.eof then
					'	detailLineNo = 0
					'	Do While Not rsReceive.eof	

							'	Insert the History of EmpallDed Data 
							' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
							
					'		txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
					'									 Session("fcBA")
					'									 
					'									 
					'		nUnitDetail = cdbl(rsReceive("mandays"))
					'		nRateDetail = nRatePremiPupuk
					'		nAmountDetail = nUnitDetail * nRateDetail
							
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
					'	loop
					'end if
					'rsReceive.Close
					'set rsReceive = Nothing									
			  rsEmpAllDed.MoveNext
		  loop
		End if
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
	end if
%>