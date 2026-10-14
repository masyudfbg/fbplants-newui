<%
	'Amandemen
	'1. Allowded Code Untuk Lembur = A18 Rate 1 = Harga Beras
	'2. Parameter Master Status Pekerja di Parameter = E0006
	'   Rate 2 = Kg Beras Pekerja

	
	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Lembur.asp"
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
	
	
	
	
	
	'Dim nBerasPekerja,nBerasPasangan,nBerasTanggungan,nHargaBeras
	  Set rsParameter = Server.CreateObject("ADODB.Recordset")
		  qParameter =" Select nvl(sum(RATE_1),0) as Rate_1, " & _
							   " nvl(sum(RATE_2),0) as Rate_2, " & _
							   " nvl(sum(RATE_3),0) as Rate_3, " & _
							   " nvl(sum(RATE_4),0) as Rate_4  " & _
                      " From ALLOWDED where FCCODE='A18'  " & _
					  " and fcBA='" & Session("fcBA") & "'"
	  'response.Write " qParameter = " & qParameter				  
	  rsParameter.CursorType = 3
	  rsParameter.CursorLocation = 3
	  rsParameter.Open qParameter, connstr
	  if not rsParameter.eof then
		 nHargaBeras= cdbl(rsParameter("Rate_1")) 
		 nKonstanta_KBT= cdbl(rsParameter("Rate_2")) 
		 nKonstanta_KHT= cdbl(rsParameter("Rate_3")) 
		 nKonstanta_KHL= cdbl(rsParameter("Rate_4")) 
	  else
		 nHargaBeras= 0
		 nKonstanta_KBT = 0
		 nKonstanta_KHT = 0
		 nKonstanta_KHL = 0 
	  End if
	  rsParameter.Close
	  set rsParameter = Nothing
	  

	' We don't use UN anymore as per UN are not acceptable for OTHRS (Overtime hours)
	' By NAF, 3 Dec 2013
	' qReceive=" select gad.FDDATE,gad.EMPLOYEECODE,emp.MARITALSTATUS,emp.PRESENTGRADE,nvl(sum(emp.PRESENTRATE),0) as PRESENTRATE ,nvl(sum(gad.OTHRS),0) as OTHRS " & _
	         ' " From Employee emp,ATTENDANCE_GAD gad Where " & _
			 ' " (emp.fccode = gad.EMPLOYEECODE and emp.fcba = gad.fcba) " & _
			 ' '" and gad.OTHRS >0  and gad.attendance in ('KJ','UN','WH','WS') and DATETERMINATE is null and " & _
			 ' " to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			 ' " and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
			 ' " group by gad.FDDATE,gad.EMPLOYEECODE,emp.MARITALSTATUS,emp.PRESENTGRADE"


	qReceive=" select gad.FDDATE,gad.EMPLOYEECODE, nvl(emp.naturetype, 'NN') naturetype, emp.MARITALSTATUS,emp.PRESENTGRADE, " & _			 
			 " nvl(emp.PRESENTRATE, 0) as PRESENTRATE ," & _
			 " TO_NUMBER(TO_CHAR(LAST_DAY(FDDATE),'dd')) LASTDAY, " & _
			 " nvl(sum(gad.OTHRS),0) as OTHRS " & _
	         " From Employee emp,ATTENDANCE_GAD gad Where " & _
			 " (emp.fccode = gad.EMPLOYEECODE and emp.fcba = gad.fcba) " & _
			 " and gad.OTHRS > 0  and gad.attendance in ('KJ', 'WH', 'WS', 'KL','KK') and DATETERMINATE is null and " & _
			 " to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			 " and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
			 " group by gad.FDDATE,gad.EMPLOYEECODE, nvl(emp.naturetype, 'NN'), emp.MARITALSTATUS,emp.PRESENTGRADE, nvl(emp.PRESENTRATE, 0)"

    'response.Write "<br><br>" & " Recordset " & qReceive & "<br><br>"
	'response.Flush()			    
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate Over Time / Lembur ..."& "<br>" 
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
	   Do While Not rsEmpAllDed.eof

			If rsEmpAllDed("naturetype") = "NN" Then
				nBerasPekerja   = 0
			Else
				'Dim rsParameter,qParameter
				Set rsParameter = Server.CreateObject("ADODB.Recordset")
				qParameter =" Select nvl(sum(RATE_1),0) as Rate_1,nvl(sum(RATE_2),0) as Rate_2 From PARAMETER where PARCODE='E0006' and " & _
							  " FCCODE='" & rsEmpAllDed("MARITALSTATUS") & "'" & _
							  " and fcBA='" & Session("fcBA") & "'"
				rsParameter.CursorType = 3
				rsParameter.CursorLocation = 3
				rsParameter.Open qParameter, connstr
				if not rsParameter.eof then
					nBerasPekerja   = cdbl(rsParameter("Rate_2")) * nHargaBeras
				else
					nBerasPekerja   = 0
				End if
				rsParameter.Close
				set rsParameter = Nothing
			End If
		  
		  qAllowance=""
		  Dim nLembur
		  if rsEmpAllded("PRESENTGRADE")="KHL" then
				'nLembur = cdbl(rsEmpAllded("PRESENTRATE")) / 7 * cdbl(rsEmpAllded("OTHRS"))
				'nLembur = (cdbl(rsEmpAllded("PRESENTRATE"))*25 / 173) * cdbl(rsEmpAllded("OTHRS"))
				nLembur = (cdbl(rsEmpAllded("PRESENTRATE"))*nKonstanta_KHL) * cdbl(rsEmpAllded("OTHRS"))
				

				
				maxLineNo = maxLineNo + 1
				qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
						"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A18" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nLembur & _
						",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
						Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
						
				connstr.execute(qAllowance)
			 
			 '* Remarks History By Iwa : 08/01/2019
				 
			'	qReceiveDetail =" select gad.FDDATE, " & _
			'							 " gad.EMPLOYEECODE, " & _
			'							 " gad.JOBCODE, " & _
			'							 " gad.LOCATIONTYPE, " & _
			'							 " gad.LOCATIONCODE, " & _
			'							 " emp.MARITALSTATUS, " & _
			'							 " emp.PRESENTGRADE, " & _
			'							 " nvl(emp.PRESENTRATE,0) as PRESENTRATE, " & _
			'							 " nvl(gad.OTHRS,0) as OTHRS " & _
			'					" From Employee emp,ATTENDANCE_GAD gad Where " & _
			'					" (emp.fccode = gad.EMPLOYEECODE and emp.fcba = gad.fcba) " & _
			'					" and gad.OTHRS >0  and gad.attendance in ('KJ', 'WH', 'WS', 'KL') and DATETERMINATE is null and " & _
			'					" to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			'					" and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
			'					" and gad.FDDATE = to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & _
			'					" and gad.EMPLOYEECODE = '" &  rsEmpAllDed("EMPLOYEECODE") & "' " & _
			'					" and emp.MARITALSTATUS = '" &  rsEmpAllDed("MARITALSTATUS") & "' " & _
			'					" and emp.PRESENTGRADE = '" &  rsEmpAllDed("PRESENTGRADE") & "' " 
								
			'	Set rsReceive = Server.CreateObject("ADODB.Recordset")
			'	rsReceive.CursorType = 3
			'	rsReceive.CursorLocation = 3
				'response.write qReceiveDetail & "</br>"
				
			'	rsReceive.Open qReceiveDetail, connstr
			'	if not rsReceive.eof then
			'		detailLineNo = 0
			'		Do While Not rsReceive.eof	

						'	Insert the History of EmpallDed Data 
						' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
						
			'			txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
			'										 rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
			'										 rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
			'										 rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
			'										 rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
			'										 Session("fcBA")
													 
			'			nLemburDetail = (cdbl(rsReceive("PRESENTRATE"))*nKonstanta_KHL) * cdbl(rsReceive("OTHRS"))
			'			nRateDetail = cdbl(rsReceive("PRESENTRATE"))*nKonstanta_KHL
			'			nOthrsDetail = cdbl(rsReceive("OTHRS"))
			'										 
			'			
			'			detailLineNo = detailLineNo + 1
			'			qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
			'												" LINEHEADERNO," & _
			'												" LINENO," & _
			'												" FCBA," & _
			'												" REF_PROCESS," & _
			'												" REF_PROCESSNAME," & _
			'												" REF_TABLESOURCE," & _
			'												" REF_NUMBERKEYS," & _
			'												" REF_STRINGSEPARATOR," & _
			'												" REF_STRINGKEYS," & _
			'												" REF_STRINGVALUES," & _
			'												" OTHRS, " & _
			'												" RATE, " & _
			'												" AMOUNT, " & _
			'												" FCENTRY," & _
			'												" FCEDIT," & _
			'												" FCIP," & _
			'												" LASTUPDATE," & _
			'												" LASTTIME)" & _
			'										" VALUES ( " & _
			'												" '" & maxLineNo & "'," & _
			'												" '" & detailLineNo & "'," & _
			'												" '" & Session("fcBA") & "'," & _
			'												" '" & tempProcessCode & "'," & _
			'												" '" & tempProcessNameCode & "'," & _
			'												" '" & tempRefCodeReceive & "'," & _
			'												" '" & txtRefNumberKeys_Receive & "'," & _
			'												" '" & txtRefStringSeparator_Receive & "'," & _
			'												" '" & txtRefStringKeys_Receive & "'," & _
			'												" '" & txtRefStringValues_Receive & "'," & _
			'												nOthrsDetail & "," & _
			'												nRateDetail & "," & _
			'												nLemburDetail & "," & _ 
			'												" '" & Session("CodeName") & "'," & _
			'												" '" & Session("CodeName") & "'," & _
			'												" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
			'												" sysdate, " & _
			'												" to_char(sysdate, 'HH24:MI') " & _
			'												" ) "					
			'			connstr.execute(qAllowanceHistory)
			'			rsReceive.MoveNext
			'		loop
			'	end if
			'	rsReceive.Close
			'	set rsReceive = Nothing	
								
								
			 
		  elseif rsEmpAllded("PRESENTGRADE")="KBT" OR rsEmpAllded("PRESENTGRADE")="KHT" then
				' Remarked by SAHAT and NAFIS
				'nLembur = 1/173* ( cdbl(rsEmpAllded("PRESENTRATE"))+ (nBerasPekerja/30))  * cdbl(rsEmpAllded("OTHRS"))
				'nLembur = 1/173* ( cdbl(rsEmpAllded("PRESENTRATE"))+ (nBerasPekerja))  * cdbl(rsEmpAllded("OTHRS"))
				
				if rsEmpAllded("PRESENTGRADE")="KBT" then 
					nLembur = nKonstanta_KBT * ( cdbl(rsEmpAllded("PRESENTRATE"))+ (nBerasPekerja))  * cdbl(rsEmpAllded("OTHRS"))
				 
				end if
				
				if rsEmpAllded("PRESENTGRADE")="KHT" then 
					'nLembur = nKonstanta_KHT * ((cdbl(rsEmpAllded("PRESENTRATE")) / 30 * cdbl(rsEmpAllded("LASTDAY")))+ nBerasPekerja)  * cdbl(rsEmpAllded("OTHRS"))
					nLembur = nKonstanta_KHT * ( cdbl(rsEmpAllded("PRESENTRATE"))+ (nBerasPekerja))  * cdbl(rsEmpAllded("OTHRS"))
		        end if
				
				if rsEmpAllDed("EMPLOYEECODE") = "101200095" then
					if rsEmpAllDed("FDDATE") = "07/05/2019" then
						response.write "nKonstanta_KHT = " & nKonstanta_KHT & "</br>"
						response.write "PRESENTRATE = " & cdbl(rsEmpAllded("PRESENTRATE")) & "</br>"
						response.write "LASTDAY = " & cdbl(rsEmpAllded("LASTDAY")) & "</br>"
						response.write "OTHRS = " & cdbl(rsEmpAllded("OTHRS")) & "</br>"
					end if
				end if
				
			    maxLineNo = maxLineNo + 1
				qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
						"(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A18" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nLembur & _
						",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
						Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
				 
				connstr.execute(qAllowance)
				
				' Remarks History By Iwa : 08/01/2019
				'qReceiveDetail = " select gad.FDDATE, " & _
				'				 " gad.EMPLOYEECODE, " & _
				'				 " gad.JOBCODE, " & _
				'				 " gad.LOCATIONTYPE, " & _
				'				 " gad.LOCATIONCODE, " & _
				'				 " emp.MARITALSTATUS, " & _
				'				 " emp.PRESENTGRADE, " & _
				'				 " nvl(emp.PRESENTRATE,0) as PRESENTRATE, " & _
				'				 " nvl(gad.OTHRS,0) as OTHRS " & _
				'				 " From Employee emp,ATTENDANCE_GAD gad Where " & _
				'				 " (emp.fccode = gad.EMPLOYEECODE and emp.fcba = gad.fcba) " & _
				'				 " and gad.OTHRS >0  and gad.attendance in ('KJ', 'WH', 'WS', 'KL') and DATETERMINATE is null and " & _
				'				 " to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
				'				 " and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
				'				 " and gad.FDDATE = to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & _
				'				 " and gad.EMPLOYEECODE = '" &  rsEmpAllDed("EMPLOYEECODE") & "' " & _
				'				 " and emp.MARITALSTATUS = '" &  rsEmpAllDed("MARITALSTATUS") & "' " & _
				'				 " and emp.PRESENTGRADE = '" &  rsEmpAllDed("PRESENTGRADE") & "' " 
								
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
													 
				'		nLemburDetail = 1/173* ( cdbl(rsReceive("PRESENTRATE"))+ (nBerasPekerja))  * cdbl(rsReceive("OTHRS"))
				'		nRateDetail = 1/173* ( cdbl(rsReceive("PRESENTRATE"))+ (nBerasPekerja))
				'		nOthrsDetail = cdbl(rsReceive("OTHRS"))
													 
						
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
				'											" OTHRS, " & _
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
				'											nOthrsDetail & "," & _
				'											nRateDetail & "," & _
				'											nLemburDetail & "," & _ 
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

		  End if	
		 ' response.Write "sql :  p" & qAllowance & "p<br>"
		  response.Flush()
		  rsEmpAllDed.MoveNext
	  loop
	End if
	rsEmpAllDed.Close
	set rsEmpAllDed = Nothing
%>