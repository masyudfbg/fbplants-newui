<%
	'Amandemen
	'1. Allowded Code Untuk Premi Tetap = A51 
	'1. Allowded Code Untuk Tunjangan Jabatan = A52
	'1. Allowded Code Untuk Premi Kehadiran = A22
	' NAF : In this methodology, system will give availability to user for defining another PREMI Addition for particular Employee
	
	
	Dim varKodePremiTetap 
	Dim ValueIncludeProses, ValueManAuto, oRS, oSQL

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Additional.asp"
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
	
	If doneReceive = 1 Then
		Response.Write "<br> Start Calculating Payroll Additional.."

		' General Calculation and Selection for Additional Payroll Premi MONTHLY		
		qSelection = " SELECT " & _
					 "	   a.employeecode, " & _
					 "     a.allowdedcode, " & _
					 "     NVL (a.rate_1, 0) + NVL (a.rate_2, 0) + NVL (a.rate_3, 0) + NVL (a.rate_4, 0) + NVL (a.rate_4, 0) totalAllowance " & _
					 "	FROM " & _
					 "	   employeeallowdeddetail a, employee b, allowded c " & _
					 "	WHERE " & _
					 "	   A.EMPLOYEECODE = B.FCCODE AND " & _
					 "	   A.ALLOWDEDCODE = C.FCcode AND " & _
					 "	   a.fcba = '" & Session("FCBA") & "' AND " & _
					 "	   a.fcba = b.fcba AND " & _
					 "	   a.fcba = c.fcba AND " & _
					 "	   B.DATETERMINATE IS NULL AND " & _
					 "	   C.INCLUDEDPAYROLLPROCESSING = 'TRUE' AND " & _
					 "	   C.MAN_AUTO = 'A' AND " & _
					 "     C.TYPE_1 = 'M'"
		
		Set rsGeneralSelection = Server.CreateObject("ADODB.Recordset")
		rsGeneralSelection.CursorType = 3
		rsGeneralSelection.CursorLocation = 3
		rsGeneralSelection.Open qSelection, connstr		
		
		If Not rsGeneralSelection.EOF then
			Do While Not rsGeneralSelection.EOF
				If cdbl(rsGeneralSelection("totalAllowance")) <> 0 Then
					varKodePremiTetap = rsGeneralSelection("allowdedcode")
				
					maxLineNo	=	maxLineNo + 1				
					qAllowance	=	" Insert Into EmpAllded (" & _
									"     EMPCODE" & _
									"     , ALLOWDEDCODE" & _
									"     , ADHOCDATE" & _
									"     , AMOUNT" & _
									"     , FIXEDALLDED" & _
									"     , FCENTRY" & _
									"     , FCEDIT" & _
									"     , FCIP" & _
									"     , FCBA" & _
									"     , REMARKS" & _
									"     , LASTUPDATE" & _
									"     , LASTTIME" & _
									"     , LINENO" & _
									" ) Values (" & _
									"     '" & rsGeneralSelection("EMPLOYEECODE") & "'" & _
									"     , '" & varKodePremiTetap & "'" & _
									"     , to_date('" & Month(Session("v_Employee")) & "/01/" & Year(Session("v_Employee")) & "','mm/dd/yyyy')" & _
									"     , " & cdbl(rsGeneralSelection("totalAllowance")) & "" & _
									"     , ''" & _
									"     , '" & Session("CodeName") & "'" & _
									"     , '" & Session("CodeName") & "'" & _
									"     , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"     , '" & Session("fcBA") & "'" & _
									"     , ''" & _
									"     , sysdate" & _
									"     , to_char(sysdate,'HH24:MI')" & _
									"     , " & maxLineNo & "" & _
									" )"
					connstr.execute(qAllowance)
					
					'Remarks By Iwa : 08/01/2019
					' Insert the History of EmpallDed Data 
					' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
					'txtRefStringValues_Receive	=	rsGeneralSelection("EMPLOYEECODE") & _
					'								txtRefStringSeparator_Receive & _
					'								Session("fcBA")
			
					'detailLineNo = detailLineNo + 1
					'qAllowanceHistory = " INSERT INTO empallded_hist_tab ( " & _
					'							" LINEHEADERNO," & _
					'							" LINENO," & _
					'							" FCBA," & _
					'							" REF_PROCESS," & _
					'							" REF_PROCESSNAME," & _
					'							" REF_TABLESOURCE," & _
					'							" REF_NUMBERKEYS," & _
					'							" REF_STRINGSEPARATOR," & _
					'							" REF_STRINGKEYS," & _
					'							" REF_STRINGVALUES," & _
					'							" UNIT, " & _
					'							" RATE, " & _
					'							" AMOUNT, " & _
					'							" FCENTRY," & _
					'							" FCEDIT," & _
					'							" FCIP," & _
					'							" LASTUPDATE," & _
					'							" LASTTIME" & _
					'					" ) VALUES ( " & _
					'							" '" & maxLineNo & "'" & _
					'							" , '" & detailLineNo & "'" & _
					'							" , '" & Session("fcBA") & "'" & _
					'							" , '" & tempProcessCode & "'" & _
					'							" , '" & tempProcessNameCode & "'" & _
					'							" , '" & tempRefCodeReceive & "'" & _
					'							" , '" & txtRefNumberKeys_Receive & "'" & _
					'							" , '" & txtRefStringSeparator_Receive & "'" & _
					'							" , '" & txtRefStringKeys_Receive & "'" & _
					'							" , '" & txtRefStringValues_Receive & "'" & _
					'							" , 1" & _
					'							" , " & cdbl(rsGeneralSelection("totalAllowance")) & "" & _
					'							" , " & cdbl(rsGeneralSelection("totalAllowance")) & "" & _
					'							" , '" & Session("CodeName") & "'" & _
					'							" , '" & Session("CodeName") & "'" & _
					'							" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					'							" , sysdate" & _
					'							" , to_char(sysdate, 'HH24:MI')" & _
					'							" ) "
		  
					'connstr.execute(qAllowanceHistory)
					
					'Response.Write "."
					'End Remarks By Iwa : 08/01/2019
				End If
				
				rsGeneralSelection.MoveNext
			Loop
		End If
		rsGeneralSelection.close
		Set rsGeneralSelection = Nothing
		
		' General Calculation and Selection for Additional Payroll Premi DAILY		
		qSelection = " SELECT " & _
					 " 	   d.fddate, " & _
					 "	   a.employeecode, " & _
					 "     a.allowdedcode, " & _
					 "     NVL (a.rate_1, 0) + NVL (a.rate_2, 0) + NVL (a.rate_3, 0) + NVL (a.rate_4, 0) + NVL (a.rate_4, 0) totalAllowance,  " & _
					 "     d.totalHK " & _
					 "	FROM " & _
					 "	   employeeallowdeddetail a, employee b, allowded c, " & _
					       " ( select " & _
							"	fcba," & _
							"	employeecode," & _
							"	fddate," & _
							"	sum(nvl(mandays,0)) totalHK" & _
						    " from attendance_gad " & _
							" where" & _
						    " attendance in ('KJ', 'WH', 'WS', 'DL') " & _
							" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
							" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
							" group by  fcba,employeecode, fddate" & _
						    " )d " & _
					 "	WHERE " & _
					 "	   A.EMPLOYEECODE = B.FCCODE AND " & _
					 "     A.EmployeeCode = d.employeecode and " & _
					 "	   A.ALLOWDEDCODE = C.FCcode AND " & _
					 "	   a.fcba = '" & Session("FCBA") & "' AND " & _
					 "	   a.fcba = b.fcba AND " & _
					 "	   a.fcba = c.fcba AND " & _
					 "	   a.fcba = d.fcba AND " & _
					 "	   B.DATETERMINATE IS NULL AND " & _
					 "	   C.INCLUDEDPAYROLLPROCESSING = 'TRUE' AND " & _
					 "	   C.MAN_AUTO = 'A' AND " & _
					 "     C.TYPE_1 = 'D'"
		
		Set rsGeneralSelection = Server.CreateObject("ADODB.Recordset")
		rsGeneralSelection.CursorType = 3
		rsGeneralSelection.CursorLocation = 3
		'response.write qSelection
		rsGeneralSelection.Open qSelection, connstr		
		
		If Not rsGeneralSelection.EOF then
			Do While Not rsGeneralSelection.EOF
				If cdbl(rsGeneralSelection("totalAllowance")) <> 0 Then
				
					varKodePremiTetap = rsGeneralSelection("allowdedcode")
					
					maxLineNo	=	maxLineNo + 1				
					qAllowance	=	" Insert Into EmpAllded (" & _
									"     EMPCODE" & _
									"     , ALLOWDEDCODE" & _
									"     , ADHOCDATE" & _
									"     , AMOUNT" & _
									"     , FIXEDALLDED" & _
									"     , FCENTRY" & _
									"     , FCEDIT" & _
									"     , FCIP" & _
									"     , FCBA" & _
									"     , REMARKS" & _
									"     , LASTUPDATE" & _
									"     , LASTTIME" & _
									"     , LINENO" & _
									" ) Values (" & _
									"     '" & rsGeneralSelection("EMPLOYEECODE") & "'" & _
									"     , '" & varKodePremiTetap & "'" & _
									"     , to_date('" & rsGeneralSelection("fddate") & "','mm/dd/yyyy')" & _
									"     , " & cdbl(rsGeneralSelection("totalAllowance")) * cdbl(rsGeneralSelection("totalHK"))  & "" & _
									"     , ''" & _
									"     , '" & Session("CodeName") & "'" & _
									"     , '" & Session("CodeName") & "'" & _
									"     , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
									"     , '" & Session("fcBA") & "'" & _
									"     , ''" & _
									"     , sysdate" & _
									"     , to_char(sysdate,'HH24:MI')" & _
									"     , " & maxLineNo & "" & _
									" )"
					connstr.execute(qAllowance)
					
					'Remarks By Iwa : 08/01/2019
					' Insert the History of EmpallDed Data 
					' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
					'txtRefStringValues_Receive	=	rsGeneralSelection("EMPLOYEECODE") & _
					'								txtRefStringSeparator_Receive & _
					'								Session("fcBA")
			
					'detailLineNo = detailLineNo + 1
					'qAllowanceHistory = " INSERT INTO empallded_hist_tab ( " & _
					'							" LINEHEADERNO," & _
					'							" LINENO," & _
					'							" FCBA," & _
					'							" REF_PROCESS," & _
					'							" REF_PROCESSNAME," & _
					'							" REF_TABLESOURCE," & _
					'							" REF_NUMBERKEYS," & _
					'							" REF_STRINGSEPARATOR," & _
					'							" REF_STRINGKEYS," & _
					'							" REF_STRINGVALUES," & _
					'							" UNIT, " & _
					'						    " RATE, " & _
					'							" AMOUNT, " & _
					'							" FCENTRY," & _
					'							" FCEDIT," & _
					'							" FCIP," & _
					'							" LASTUPDATE," & _
					'							" LASTTIME" & _
					'					" ) VALUES ( " & _
					'							" '" & maxLineNo & "'" & _
					'							" , '" & detailLineNo & "'" & _
					'							" , '" & Session("fcBA") & "'" & _
					'							" , '" & tempProcessCode & "'" & _
					'							" , '" & tempProcessNameCode & "'" & _
					'							" , '" & tempRefCodeReceive & "'" & _
					'							" , '" & txtRefNumberKeys_Receive & "'" & _
					'							" , '" & txtRefStringSeparator_Receive & "'" & _
					'							" , '" & txtRefStringKeys_Receive & "'" & _
					'							" , '" & txtRefStringValues_Receive & "'" & _
					'							" , " & cdbl(rsGeneralSelection("totalHK"))  & "" & _
					'							" , " & cdbl(rsGeneralSelection("totalAllowance"))  & "" & _
					'							" , " & cdbl(rsGeneralSelection("totalAllowance")) * cdbl(rsGeneralSelection("totalHK"))  & "" & _
					'							" , '" & Session("CodeName") & "'" & _
					'							" , '" & Session("CodeName") & "'" & _
					'							" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					'							" , sysdate" & _
					'							" , to_char(sysdate, 'HH24:MI')" & _
					'							" ) "
		  
					'connstr.execute(qAllowanceHistory)
					
					'Response.Write "."
					'End Remarks By Iwa : 08/01/2019
				End If
				
				rsGeneralSelection.MoveNext
			Loop
		End If
		rsGeneralSelection.close
		Set rsGeneralSelection = Nothing
		
		
		Response.Write "End Calculating Payroll Additional.. Conratulations....<br>"
	Else
		Response.Write "Start Calculating Payroll Additional.."
		Response.Write "End Calculating Payroll Additional.."
	End If
%>