<%
	'Amandemen
	'1. Parameter Master Status Pekerja di Parameter = E0006
	'   Rate 2 = Kg Beras Pekerja, Rate 3 = Kg Beras Pasangan, Rate 4 = Kg Beras Tanggungan
	
	'2. Allowded Code Untuk Tunjangan Beras Pekerja = A12,Rate 1 = Harga Beras
	'   Allowded Code Untuk Tunjangan Beras Pasangan  =A13
	'   Allowded Code Untuk Tunjangan Beras Tanggungan =A14
	Dim nBerasPekerja,nBerasPasangan,nBerasTanggungan,nHargaBeras, temp_rsParameter_rate_2
	
	
	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_TunjanganKeluarga.asp"
	tempRefCodeReceive = "EMPLOYEE"

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
	
	' Cek Apakah allowded akan di proses payroll?
	Set objectRS = Server.CreateObject("ADODB.Recordset")
	objectSQL = " SELECT INCLUDEDPAYROLLPROCESSING" & _
	                       " FROM ALLOWDED" & _
	                       " WHERE FCBA = '" & Session("fcBA") & "'" & _
	                       "       AND FCCODE = 'A12'"

	'response.write objectSQL
	objectRS.Open objectSQL, connstr

	tempIncludedPayrollProcessing = "FALSE"
	If Not objectRS.EOF then
		tempIncludedPayrollProcessing = objectRS("INCLUDEDPAYROLLPROCESSING")
	Else
		tempIncludedPayrollProcessing = "FALSE"
	End if
	objectRS.Close
	set objectRS = Nothing
	
	
	IF tempIncludedPayrollProcessing = "TRUE" THEN
		
		
		

		qReceive="select FCCODE,MARITALSTATUS, NATURETYPE From Employee " & _
		  " WHERE DATETERMINATE is null and MARITALSTATUS<>'NT' " & _
		  " and fcBA='" & Session("fcBA") & "' GROUP BY FCCODE,MARITALSTATUS, NATURETYPE" 
		  
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Tunjangan Keluarga..."& "<br>"
		response.Flush()
		
		Set rsParameter = Server.CreateObject("ADODB.Recordset")
			qParameter =" Select nvl(sum(RATE_1),0) as Rate_1 From ALLOWDED where FCCODE='A12' and " & _
						" fcBA='" & Session("fcBA") & "'"
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		rsParameter.Open qParameter, connstr
		if not rsParameter.eof then
		 nHargaBeras= cdbl(rsParameter("Rate_1")) 
		else
		 nHargaBeras= 0
		End if
		rsParameter.Close
		set rsParameter = Nothing

		rsEmpAllDed.Open qReceive, connstr
		if not rsEmpAllDed.eof then
		   Do While Not rsEmpAllDed.eof
			  Dim rsParameter,qParameter
			  Set rsParameter = Server.CreateObject("ADODB.Recordset")
				  qParameter =" Select nvl(sum(RATE_2),0) as Rate_2,nvl(sum(RATE_3),0) as Rate_3," & _
							  " nvl(sum(RATE_4),0) as Rate_4 From PARAMETER where PARCODE='E0006' and " & _
							  " FCCODE='" & rsEmpAllDed("MARITALSTATUS") & "'" & _
							  " and fcBA='" & Session("fcBA") & "'"
			  rsParameter.CursorType = 3
			  rsParameter.CursorLocation = 3
			  'response.write qParameter
			  rsParameter.Open qParameter, connstr
			  if not rsParameter.eof then
				 nBerasPekerja   = cdbl(rsParameter("Rate_2")) * nHargaBeras
				 nBerasPasangan  = cdbl(rsParameter("Rate_3")) * nHargaBeras
				 nBerasTanggungan= cdbl(rsParameter("Rate_4")) * nHargaBeras
			  else
				 nBerasPekerja   = 0 
				 nBerasPasangan  = 0
				 nBerasTanggungan= 0
			  End if
			  temp_rsParameter_rate_2 = cdbl(rsParameter("Rate_2"))
			  temp_rsParameter_rate_3 = cdbl(rsParameter("Rate_3"))
			  temp_rsParameter_rate_4 = cdbl(rsParameter("Rate_4"))
			  
			  rsParameter.Close
			  set rsParameter = Nothing
			  
			  if rsEmpAllDed("NATURETYPE") = "TA" then 
				  if nBerasPekerja > 0 then
						maxLineNo = maxLineNo + 1
						qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
									"(" & maxLineNo & ",'" & rsEmpAllDed("FCCODE") & "','" & "A12" & "'," &  " to_date('" & Session("v_Employee") & "','mm/dd/yyyy') " & "," & nBerasPekerja & _
									",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',sysdate, to_char(SYSDATE, 'HH24:MI'))"
						connstr.execute(qAllowance)
						
						'**** Remarks by UKIE, to disable history, 21 April 2016
					 
						' ' Insert the History of Beras Pekerja
						' txtRefStringValues_Receive = rsEmpAllDed("FCCODE") & txtRefStringSeparator_Receive & _
													 ' Session("fcBA")
						
						' detailLineNo = 1
						
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
															' " UNIT, " & _
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
															' temp_rsParameter_rate_2 & "," & _
															' nHargaBeras & "," & _
															' nBerasPekerja & "," & _ 
															' " '" & Session("CodeName") & "'," & _
															' " '" & Session("CodeName") & "'," & _
															' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
															' " sysdate, " & _
															' " to_char(sysdate, 'HH24:MI') " & _
															' " ) "				

															
						' connstr.execute(qAllowanceHistory)
					 
				  End if		
				  if nBerasPasangan > 0 then
						maxLineNo = maxLineNo + 1
						qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
								"(" & maxLineNo & ",'" & rsEmpAllDed("FCCODE") & "','" & "A13" & "'," &  " to_date('" & Session("v_Employee") & "','mm/dd/yyyy') " & "," & nBerasPasangan & _
								",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',sysdate, to_char(SYSDATE, 'HH24:MI'))"
						connstr.execute(qAllowance)
						
						'**** Remarks by UKIE, to disable history, 21 April 2016
					 
						' ' Insert the History of Beras Pekerja
						' txtRefStringValues_Receive = rsEmpAllDed("FCCODE") & txtRefStringSeparator_Receive & _
													 ' Session("fcBA")
						
						' detailLineNo = 1
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
															' " UNIT, " & _
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
															' temp_rsParameter_rate_3 & "," & _
															' nHargaBeras & "," & _
															' nBerasPasangan & "," & _ 
															' " '" & Session("CodeName") & "'," & _
															' " '" & Session("CodeName") & "'," & _
															' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
															' " sysdate, " & _
															' " to_char(sysdate, 'HH24:MI') " & _
															' " ) "					
						' connstr.execute(qAllowanceHistory)
					 
					 
				  End if		
				  if nBerasTanggungan > 0 then
						maxLineNo = maxLineNo + 1
						qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
								"(" & maxLineNo & ",'" & rsEmpAllDed("FCCODE") & "','" & "A14" & "'," &  " to_date('" & Session("v_Employee") & "','mm/dd/yyyy') " & "," & nBerasTanggungan & _
								",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',sysdate, to_char(SYSDATE, 'HH24:MI'))"

						connstr.execute(qAllowance)
						
						'**** Remarks by UKIE, to disable history, 21 April 2016
					 
						' ' Insert the History of Beras Pekerja
						' txtRefStringValues_Receive = rsEmpAllDed("FCCODE") & txtRefStringSeparator_Receive & _
													 ' Session("fcBA")
						
						' detailLineNo = 1
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
															' " UNIT, " & _
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
															' temp_rsParameter_rate_4 & "," & _
															' nHargaBeras & "," & _
															' nBerasTanggungan & "," & _ 
															' " '" & Session("CodeName") & "'," & _
															' " '" & Session("CodeName") & "'," & _
															' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
															' " sysdate, " & _
															' " to_char(sysdate, 'HH24:MI') " & _
															' " ) "					
						' connstr.execute(qAllowanceHistory)				 
					 
				  End if
			  end if
			  response.Flush()
			  rsEmpAllDed.MoveNext
		  loop
		  set nBerasPekerja = Nothing
		  set nBerasPasangan = Nothing
		  set nBerasTanggungan = Nothing
		  set nHargaBeras = Nothing
		End if
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
		
	END IF
%>