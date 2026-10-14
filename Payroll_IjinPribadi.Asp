<%
	'Amandemen Potongan Ijin Pribadi
	'Code Deduction untuk Ijin Pribadi = D18
	'Potongan Ijin Pribadi diproses Setelah Tunjangan Keluarga
	'Sehingga Tunjangan Beras Pekerja A12, Pasangan A13, Tanggungan A14 Tinggal 
	'Ambil Dari Payslip 
	
	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.
	
	Dim var_P1 : var_P1 = "D18"
	Dim ConstDividerP1_KHT : ConstDividerP1_KHT = 0
	Dim ConstDividerP1_KBT : ConstDividerP1_KBT = 0

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_IjinPribadi.asp"
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
	
	' Cek Apakah allowded akan di proses payroll?
	Set objectRS = Server.CreateObject("ADODB.Recordset")
	objectSQL = " SELECT INCLUDEDPAYROLLPROCESSING" & _
	                       " FROM ALLOWDED" & _
	                       " WHERE FCBA = '" & Session("fcBA") & "'" & _
	                       "       AND FCCODE = '" & var_P1 & "'"

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
	
	
	Set objectRS = Server.CreateObject("ADODB.Recordset")
	objectSQL = " SELECT NVL(param_1, 'A12,A13,A14') param" & _
                " FROM parameter" & _
                " WHERE fcba = '" & Session("fcBA") & "'" & _
                "       AND parcode = 'PAYROLL_IJINPRIBADI'" & _
                "       AND fccode = 'ALLOWDEDBERAS'"

	'response.write objectSQL
	objectRS.Open objectSQL, connstr

	p_allowdedberas = "A12,A13,A14"
	If Not objectRS.EOF then
		p_allowdedberas = objectRS("param")
	Else
		p_allowdedberas = "A12,A13,A14"
	End if
	objectRS.Close
	set objectRS = Nothing
	
	
	IF tempIncludedPayrollProcessing = "TRUE" THEN

	
		Set objectRS = Server.CreateObject("ADODB.Recordset")
		objectSQL = " SELECT SUM(DECODE(RATE_2, 0, 30, RATE_2)) DIVIDER_KBT" & _
					"        , SUM(DECODE(RATE_3, 0, 30, 100, TO_NUMBER(TO_CHAR(LAST_DAY(TO_DATE('" & Lpad(Month(Session("v_Employee")), "0", "2") & "/01/" & Year(Session("v_Employee")) & "', 'MM/DD/YYYY')), 'DD')), RATE_3)) DIVIDER_KHT" & _
					" FROM ALLOWDED" & _
					" WHERE FCBA = '" & Session("fcBA") & "'" & _
					"       AND FCCODE = '" & var_P1 & "'"

		objectRS.CursorType = 3
		objectRS.CursorLocation = 3
		'response.write objectSQL
		objectRS.Open objectSQL, connstr
		
		
		' Val_Kategory = Kategory cara perhitungan Gaji : Bisa Bulanan/Harian
		If Not objectRS.EOF then
			ConstDividerP1_KBT = cdbl(objectRS("DIVIDER_KBT"))
			ConstDividerP1_KHT = cdbl(objectRS("DIVIDER_KHT"))
		Else
			ConstDividerP1_KBT = 30
			ConstDividerP1_KHT = 30
		End if
		
		
		objectRS.Close
		set objectRS = Nothing
		
		
		
		
		qReceive =	" SELECT ATT_GAD.EMPLOYEECODE AS FCCODE" & _
					"        , EMP.PRESENTGRADE AS PRESENTGRADE" & _
					"        , NVL(EMP.PRESENTRATE, 0) AS PRESENTRATE" & _
					"        , ATT_GAD.ATTENDANCE AS ATTENDANCE" & _								
					"        , NVL(COUNT(*),0) AS NJLHHARI" & _
					" FROM ATTENDANCE_GAD ATT_GAD" & _
					"      ,EMPLOYEE EMP" & _
					" WHERE (ATT_GAD.EMPLOYEECODE = EMP.FCCODE AND ATT_GAD.FCBA = EMP.FCBA)" & _
					"       AND DATETERMINATE IS NULL" & _
					"       AND ATT_GAD.ATTENDANCE IN ('P1')" & _
					"       AND EMP.PRESENTGRADE IN ('KHT','KBT')" & _
					"       AND TO_NUMBER(TO_CHAR(FDDATE,'MM')) = " & Month(Session("v_Employee")) & "" & _
					"       AND TO_NUMBER(TO_CHAR(FDDATE,'YYYY')) = " & Year(Session("v_Employee")) & "" & _
					"       and ATT_GAD.FCBA = '" & Session("fcBA") & "'" & _
					" GROUP BY ATT_GAD.EMPLOYEECODE" & _
					"          , EMP.PRESENTGRADE" & _
					"          , NVL(EMP.PRESENTRATE, 0)" & _
					"          , ATT_GAD.ATTENDANCE"
		  
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Ijin Pribadi ..." & "<br>"
		response.Flush()
		rsEmpAllDed.Open qReceive, connstr
		if not rsEmpAllDed.eof then
		   Do While Not rsEmpAllDed.eof
				nPresentRate=cdbl(rsEmpAllDed("PRESENTRATE"))
				nDays=cdbl(rsEmpAllDed("nJlhHari"))
				Dim nTunjanganKeluarga1
				qReceive=" Select nvl(sum(AMOUNT),0) as Amount From EMPALLDED where INSTR('"&p_allowdedberas&"', ALLOWDEDCODE) <> 0" & _
					   " and EMPCODE='" & rsEmpAllDed("FCCODE") & "'" & _
					   " and to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					   " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and fcBA='" & Session("fcBA") & "'" 
				Set rsPayslip = Server.CreateObject("ADODB.Recordset")
				rsPayslip.Open qReceive, connstr,3,3
				if not rsPayslip.eof then		  
				 nTunjanganKeluarga1=cdbl(rsPayslip("Amount"))
				else
				 nTunjanganKeluarga1=0	 
				end if
				
				If rsEmpAllDed("PRESENTGRADE") = "KBT" Then
					nAmount	=	nDays*(nPresentRate/ConstDividerP1_KBT+(nTunjanganKeluarga1/ConstDividerP1_KBT))
				Elseif rsEmpAllDed("PRESENTGRADE") = "KHT" Then
					nAmount	=	nDays*(nPresentRate/ConstDividerP1_KHT+(nTunjanganKeluarga1/ConstDividerP1_KHT))
				Else
					nAmount	=	0
				End If
				
				maxLineNo = maxLineNo + 1
				qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
						  "(" & maxLineNo & ",'" & rsEmpAllDed("FCCODE") & "','" & "D18" & "'," &  " to_date('" & Session("v_Employee") & "','mm/dd/yyyy') " & "," & nAmount & _
						  ",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
						  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
				'response.Write(qAllowance) & "<br>"		  
				connstr.execute(qAllowance)
				
				'**** Remarks by UKIE, to disable history, 21 April 2016
			  
			  
				' ' Start to insert history detail
				' qReceiveDetail = "select att_gad.fddate, " & _
										 ' " att_gad.employeecode, " & _
										 ' " emp.presentgrade, " & _
										 ' " att_gad.jobcode," & _
										 ' " att_gad.locationtype," & _
										 ' " att_gad.locationcode" & _
								' " from attendance_gad Att_Gad,Employee Emp" & _
								' " where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
								' " and DATETERMINATE is null and Att_Gad.Attendance in ('P1') " & _
								' " and EMP.PRESENTGRADE in ('KHT','KBT') " & _
								' " and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
								' " and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
								' " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
								' " and Att_Gad.employeecode='" & rsEmpAllDed("FCCODE") & "'" 
								
				' Set rsReceive = Server.CreateObject("ADODB.Recordset")
				' rsReceive.CursorType = 3
				' rsReceive.CursorLocation = 3

				' 'response.write qReceiveDetail
				' rsReceive.Open qReceiveDetail, connstr
				' if not rsReceive.eof then
					' detailLineNo = 0
					' Do While Not rsReceive.eof	

						' '	Insert the History of EmpallDed Data 
						' ' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
						
						' txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
													 ' rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
													 ' rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
													 ' rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
													 ' rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
													 ' Session("fcBA")
													 
													 
						' If rsReceive("PRESENTGRADE") = "KBT" Then
							' nRateDetail = (nPresentRate/ConstDividerP1_KBT+(nTunjanganKeluarga/ConstDividerP1_KBT))
						' Elseif rsReceive("PRESENTGRADE") = "KHT" Then
							' nRateDetail = (nPresentRate/ConstDividerP1_KHT+(nTunjanganKeluarga/ConstDividerP1_KHT))
						' Else
							' nRateDetail	=	0
						' End If
													 
						' nUnitDetail = -1
						' nMangkirDetail = nRateDetail * nUnitDetail
													 
						
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
															' nUnitDetail & "," & _
															' nRateDetail & "," & _
															' nMangkirDetail & "," & _ 
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
		
	END IF
%>