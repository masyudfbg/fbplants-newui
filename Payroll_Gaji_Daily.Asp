<%

	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.
	
	Dim objectRS, objectSQL, qfingger
	Dim var_Gaji : var_Gaji = "A01"
	Dim val_Divider, param_paidmandays, param_paidcount
	Dim nTotalDayThisMonth : nTotalDayThisMonth = 0
	Dim nilaiGaji : nilaiGaji = 0
	tglnow = Session("tglnow")

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Gaji_Daily.asp"
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
	
    tempProcessCode1 = "PAYROLL"
	tempProcessNameCode1 = "Payroll_Gaji_Daily.asp"
	tempRefCodeReceive1 = "ATTENDANCE_GAD"

	qReference1 = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
				  " from process_tableref_tab " & _
				  " where " & _
				  " processcode = '" & tempProcessCode1 & "' and " & _
				  " upper(processnamecode) =  upper('" & tempProcessNameCode1 & "') and " & _
				  " fccode =  '" & tempRefCodeReceive1 & "'"

	Set rsReferenceReceive1 = Server.CreateObject("ADODB.Recordset")
	rsReferenceReceive1.CursorType = 3
	rsReferenceReceive1.CursorLocation = 3
	rsReferenceReceive1.Open qReference1, connstr

	doneReceive = 1
	if not rsReferenceReceive1.eof then
		txtRefNumberKeys_Receive1 = rsReferenceReceive1("REF_NUMBERKEYS")
		txtRefStringSeparator_Receive1 = rsReferenceReceive1("REF_STRINGSEPARATOR")
		txtRefStringKeys_Receive1 = rsReferenceReceive1("REF_STRINGKEYS")
		doneReceive = 1
	else
		doneReceive = 0
	end if
	rsReferenceReceive1.close
	set rsReferenceReceive1 = Nothing
	' End of Picking Referential Integrity
	
	
	
	' Cek Apakah allowded akan di proses payroll?
	Set objectRS = Server.CreateObject("ADODB.Recordset")
	objectSQL = " SELECT INCLUDEDPAYROLLPROCESSING" & _
	                       " FROM ALLOWDED" & _
	                       " WHERE FCBA = '" & Session("fcBA") & "'" & _
	                       "       AND FCCODE = '" & var_Gaji & "'"

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
	
	'**** Cek FingerValidation ( Attendance Pakai FInger/Tidak)
	Set objectRS = Server.CreateObject("ADODB.Recordset")
	objectSQL = " SELECT f.isvalidation FROM Fingervalidation f WHERE f.isvalidation ='TRUE' " & _ 
	            " AND TO_NUMBER(TO_CHAR (fddate, 'yyyymm')) <= TO_NUMBER('" & Year(Session("v_Employee")) & "'||LPAD('" & Month(Session("v_Employee")) & "', 2, 0)) " & _  
			    " AND f.fcba ='" & Session("fcBA") & "'"
	
	objectRS.Open objectSQL, connstr

	If Not objectRS.EOF then
		Fingervalidation = objectRS("isvalidation")
	Else
		Fingervalidation = "FALSE"
	End if
	objectRS.Close
	set objectRS = Nothing
    
    IF 	Fingervalidation <> "FALSE" THEN
	
		'IF (Year(Session("v_Employee")) >= 2021) AND (Month(Session("v_Employee"))>=7) THEN
		Response.Write "<font face=Verdana size=1 >" & "Calculate Mandays From Fingerprint...." & "<br><br></font>"
		Response.Flush
		
		'**** Update Absensi dari Finger
		'1. restore dari attendance_gad_temp
		qfingger = 	"MERGE INTO attendance_gad x " & _
						"USING ( " & _ 
						"SELECT gangcode,fddate,fcba,employeecode,jobcode,locationtype,locationcode,linenokey, attendance,mandays FROM attendance_gad_temp  " & _ 
						"WHERE " & _ 
						"fcba= '" & Session("fcBA") & "'" & _
						"AND fddate = to_date('" & tglnow & "','mm/dd/yyyy')" & _
						") y " & _
						"ON (x.employeecode=y.employeecode and x.fddate  = y.fddate and x.FCBA=y.FCBA and x.gangcode=y.gangcode AND x.locationtype=y.locationtype AND x.locationcode=y.locationcode AND x.linenokey=y.linenokey ) " & _
						"WHEN MATCHED THEN " & _
						"	UPDATE SET x.attendance = y.attendance, " & _
						"			   x.mandays = y.mandays " 
        
		connstr.execute(qfingger)
		
		'2. Backup Attendance dan MANDAYS
		qfingger = 	"UPDATE attendance_gad x " & _
						"SET " & _ 
						"	Attendance_fin = attendance,mandays_fin=mandays " & _ 
						"WHERE fcba = '" & Session("fcBA") & "'" & _
						"AND fddate = to_date('" & tglnow & "','mm/dd/yyyy')" 

		connstr.execute(qfingger)
		
		
		'2. Jika tidak ada finger jad mangkir Kecuali Panen utk absen KJ,WS,WH      
		qfingger = 	"MERGE INTO attendance_gad x " & _
						"USING ( " & _ 
						"SELECT employeecode,a.gangcode,fddate,a.fcba,b.empno,b.tdate,a.linenokey FROM attendance_gad a " & _ 
						"LEFT JOIN userfinger.attendanceview b ON a.employeecode= b.empno AND a.fddate=b.tdate AND a.fcba=b.fcba " & _ 
						"WHERE " & _ 
						"a.fcba= '" & Session("fcBA") & "'" & _
						"AND fddate = to_date('" & tglnow & "','mm/dd/yyyy') " & _
						"AND b.empno IS NULL " & _ 
						"AND a.JobCode <> '6502101' " & _
						"AND a.attendance in ('KJ','WS','WH') " & _
						"AND TO_CHAR(tdate,'D') <> '1' " & _ 
						"AND TO_CHAR(a.fddate,'yyyymmdd') NOT IN (SELECT DISTINCT TO_CHAR(holidaydate,'yyyymmdd')  FROM budget_transaction_holiday  WHERE To_Char(holidaydate,'yyyymm')='" & Year(Session("v_Employee")) & "'||LPAD('" & Month(Session("v_Employee")) & "', 2, 0)) " & _
						") y " & _
						"ON (x.employeecode=y.employeecode and x.fddate  = y.fddate and x.FCBA=y.FCBA and x.gangcode=y.gangcode  AND x.linenokey=y.linenokey) " & _
						"WHEN MATCHED THEN " & _
						"	UPDATE SET x.attendance = 'MK', " & _
						"			   x.mandays = 0 " 
		connstr.execute(qfingger)
	
	    '2. JIka ada finger maka HK Proposi ,sabtu /5 ,h kerja /7 , Kecuali panen dan PKS 
	    qfingger = 	"MERGE INTO attendance_gad x " & _
						"USING ( " & _ 
						"SELECT DISTINCT employeecode,a.gangcode,fddate,a.fcba,a.linenokey,Round(((timeout-timein)*1440)/60,0) QtyHrs, " & _ 
						"Case When To_Char(tdate,'D') in ('1','6','7') Then " & _ 
						"	Case When Round(((timeout-timein)*1440)/60,0) >= 5 then 1 else  round(Round(((timeout-timein)*1440)/60,0) / 5,2) End "& _ 
						"Else " & _
						"	Case When Round(((timeout-timein)*1440)/60,0) >= 7 then 1 else  round(Round(((timeout-timein)*1440)/60,0) / 7,2) End "& _ 
						"End NewHK " & _ 
						"FROM attendance_gad a " &_ 
						"LEFT JOIN userfinger.attendanceview b ON a.employeecode= b.empno AND a.fddate=b.tdate AND a.fcba=b.fcba " & _ 
						"WHERE " & _ 
						"a.fcba= '" & Session("fcBA") & "'" & _
						"AND fddate = to_date('" & tglnow & "','mm/dd/yyyy') " & _
						"AND b.empno IS NOT NULL " & _ 
						"AND a.JobCode <> '6502101' " & _
						"/*AND a.attendance in ('KJ','WS','WH')*/ " & _
						"AND a.attendance in ('KJ') " & _
						"AND businessunit_api.get_FCTYPE('" & Session("fcBA") & "') <> 'M' " & _ 
						"AND TO_CHAR(tdate,'D') <> '1' " & _ 
						"AND TO_CHAR(a.fddate,'yyyymmdd') NOT IN (SELECT DISTINCT TO_CHAR(holidaydate,'yyyymmdd')  FROM budget_transaction_holiday  WHERE To_Char(holidaydate,'yyyymm')='" & Year(Session("v_Employee")) & "'||LPAD('" & Month(Session("v_Employee")) & "', 2, 0)) " & _
						") y " & _
						"ON (x.employeecode=y.employeecode and x.fddate  = y.fddate and x.FCBA=y.FCBA and x.gangcode=y.gangcode AND x.linenokey=y.linenokey ) " & _
						"WHEN MATCHED THEN " & _
						"	UPDATE SET x.mandays = round(y.NewHK * x.mandays_fin,2) " 
		
		connstr.execute(qfingger)
		
		Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Mandays From Fingerprint success...." & "<br><br></font>"
		Response.Flush
      
      'END IF	  
	ELSE
		Response.Write "<font face=Verdana size=1 color=Red>" & "Skip Calculate Mandays From Fingerprint success...." & "<br><br></font>"
		Response.Flush
	END IF
		'**** Simpan Semua Parameter ke dalam variable Untuk Perhitungan Golongan KHT ****'
		Dim val_kategori_kbt : val_kategori_kbt = "M"
		Dim val_kategori_kht : val_kategori_kht = "M"		
		
		Set objectRS = Server.CreateObject("ADODB.Recordset")
		objectSQL = " SELECT NVL(parameterdetail_api.get_param_1('" & Session("fcBA") & "', 'PAYROLL_GAJI', 'KBT', 'CALC_CATEGORY'), 'M') kategori_kbt" & _
                    "        , NVL(parameterdetail_api.get_param_1('" & Session("fcBA") & "', 'PAYROLL_GAJI', 'KHT', 'CALC_CATEGORY'), 'M') kategori_kht" & _
					" FROM DUAL"
							   
		'response.write objectSQL
		objectRS.Open objectSQL, connstr
		
		If Not objectRS.EOF then
			val_kategori_kbt = objectRS("kategori_kbt")
			val_kategori_kht = objectRS("kategori_kht")
		Else
			val_kategori_kbt = "M"
			val_kategori_kht = "M"
		End if						
		objectRS.Close
		set objectRS = Nothing
		
		emp_presentrate = 0	
		
		val_pembagi_kbt = 1 'Default KBT
		val_divider = 25 'Default KHT
		
		val_pembagi_kbt = GetLastDay(Month(Session("v_Employee")), Year(Session("v_Employee")))
		'val_pembagi_kbt = 25
		
						
		Set objectRS = Server.CreateObject("ADODB.Recordset")
		objectSQL = " SELECT NVL(parameterdetail_api.get_rate_1('" & Session("fcBA") & "', 'PAYROLL_GAJI', 'KHT', 'DIVIDER'), 0) DIVIDER_KHT" & _
					" 		 , NVL(parameterdetail_api.get_rate_2('" & Session("fcBA") & "', 'PAYROLL_GAJI', 'KHT', 'DIVIDER'), 0) DIVIDER_KHT_VAL" & _
                    "        , NVL(parameterdetail_api.get_param_1('" & Session("fcBA") & "', 'PAYROLL_GAJI', 'KHT', 'PAIDMANDAYS'), 'N/A') PAIDMANDAYS" & _
                    "        , NVL(parameterdetail_api.get_param_1('" & Session("fcBA") & "', 'PAYROLL_GAJI', 'KHT', 'PAIDCOUNT'), 'N/A') PAIDCOUNT" & _
                    " FROM DUAL"
							   
		' response.write objectSQL
		objectRS.Open objectSQL, connstr
		
		If Not objectRS.EOF then
			val_Divider = cdbl(to_nol(objectRS("DIVIDER_KHT")))
			DIVIDER_KHT_VAL = cdbl(to_nol(objectRS("DIVIDER_KHT_VAL")))
			param_paidmandays = objectRS("PAIDMANDAYS")
			param_paidcount = objectRS("PAIDCOUNT")
		Else
			val_Divider = 25
			param_paidmandays = "N/A"
			param_paidcount = "N/A"
		End if						
		objectRS.Close
		set objectRS = Nothing
		
		IF "" & val_Divider = "" OR val_Divider = 0 THEN
			val_Divider = 25
		ELSE
			IF val_Divider = 100 THEN
				'val_Divider = GetLastDay(Month(Session("v_Employee")), Year(Session("v_Employee")))
				' pembagi ambil dari rate_2
				val_Divider = DIVIDER_KHT_VAL
			ELSE
				val_Divider = val_Divider
			END IF
		END IF

	''	qReceive = " SELECT fccode, presentgrade, NVL(presentrate,0) AS presentrate FROM employee " & _
      ''             " WHERE DATETERMINATE IS NULL  " & _
     ''              "       AND fcBA = '" & Session("fcBA") & "'" 
	 qReceive ="SELECT EMPLOYEECODE fccode, EMPLOYEE_API.get_presentgrade(fcba ,EMPLOYEECODE) presentgrade ,  " & _
        " NVL(EMPLOYEE_API.get_presentrate(fcba ,EMPLOYEECODE),0) AS presentrate " & _
		" FROM ATTENDANCE_GAD a " & _
		" WHERE a.fcBA = '" & Session("fcBA") & "' and a.fddate=to_date('" & tglnow & "','mm/dd/yyyy') " & _
        " and EMPLOYEE_API.GET_DATETERMINATE(fcba ,EMPLOYEECODE) is null      GROUP BY EMPLOYEECODE, fcba"
		  
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Basic Salary..." & "<br>"
		response.Flush()
		rsEmpAllDed.Open qReceive, connstr
		if not rsEmpAllDed.eof then
		   Do While Not rsEmpAllDed.eof
			  in_empallded = 0
			  emp_presentrate = 0
			  count_data = 0
			  curr_data = 0
			  
			  emp_presentrate = Cdbl(to_nol(rsEmpAllDed("PRESENTRATE")))
			  
			  if rsEmpAllDed("PRESENTGRADE")="KBT" then
			  
				IF val_kategori_kbt = "D" THEN
						arr_data_kbt_eof = true
						objectSQL = " SELECT fcba" & _
									"        , fddate" & _
									"        , employeecode" & _
									"        , attendance" & _
									"        , mandays" & _
									"        , payroll_gaji_api.get_actual_mandays_kbt (attendance, mandays) act_mandays" & _
									"        , TO_CHAR(fddate, 'dd') xlastday" & _
									" FROM ( SELECT fcba" & _
									"               , fddate" & _
									"               , employeecode" & _
									"               , attendance" & _
									"               , SUM (mandays) mandays" & _
									"        FROM attendance_gad" & _
									"        WHERE employeecode = '" & rsEmpAllDed("FCCODE") & "'" & _
									"              AND fcba = '" & Session("fcBA") & "'" & _
									"              AND fddate = to_date('" & tglnow & "','mm/dd/yyyy') " & _
									"              AND attendance NOT IN ('KB', 'KL')" & _
									"        GROUP BY fcba" & _
									"                 , fddate" & _
									"                 , employeecode" & _
									"                 , attendance" & _
									" ) ORDER BY fddate"
						'response.write objectSQL
						Set objectRS = Server.CreateObject("ADODB.Recordset")
						objectRS.Open objectSQL, connstr
						If Not objectRS.EOF then
							arr_data_kbt = objectRS.getrows()
							arr_data_kbt_eof = false
						End if
						objectRS.Close
						set objectRS = Nothing
						
						if not arr_data_kbt_eof then
							for yyi=0 to ubound(arr_data_kbt,2)
								count_data = count_data +1
								nilaiGaji = Round( (emp_presentrate/val_pembagi_kbt)*cdbl(to_nol(arr_data_kbt(5,yyi))), 0)
								
								
			  
								'if cdbl(to_nol(val_pembagi_kbt)) = count_data then
								if ubound(arr_data_kbt,2) +1 = count_data then
									nilaiGaji = emp_presentrate - in_empallded
								else
									in_empallded = in_empallded +nilaiGaji
								end if
								
								if nilaiGaji > 0 then
									maxLineNo = maxLineNo + 1
									qAllowance = " INSERT INTO empallded (lineno, empcode, allowdedcode, adhocdate, amount, fixedallded, fcentry, fcedit, fcip, fcba, lastupdate, lasttime)" & _
												 " VALUES ( '" & maxLineNo & "'" & _
									             "          , '" & rsEmpAllDed("FCCODE") & "'" & _
									             "          , 'A01'" & _
									             "          , TO_DATE('" & arr_data_kbt(1,yyi) & "', 'mm/dd/yyyy')" & _
									             "          , '" & nilaiGaji & "'" & _
												 "          , 0" & _
									             "          , '" & Session("CodeName") & "'" & _
									             "          , '" & Session("CodeName") & "'" & _
									             "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									             "          , '" & Session("fcBA") & "'" & _
									             "          , SYSDATE" & _
									             "          , TO_CHAR(SYSDATE, 'HH24:MI'))"
												  'response.write "</br>" & qAllowance
									connstr.execute(qAllowance)
								end if
							next
						end if
				ELSE
					'Karyawan Bulanan Tetap Dari Present Rate
					if cdbl(to_nol(rsEmpAllDed("PRESENTRATE"))) > 0 then
						maxLineNo = maxLineNo + 1
						qAllowance = " INSERT INTO empallded (lineno, empcode, allowdedcode, adhocdate, amount,fixedallded, fcentry, fcedit, fcip, fcba, lastupdate, lasttime)" & _
							         " VALUES ( '" & maxLineNo & "'" & _
									 "          , '" & rsEmpAllDed("FCCODE") & "'" & _
									 "          , 'A01'" & _
									 "          , TO_DATE('" & Session("v_Employee") & "','mm/dd/yyyy')" & _
									 "          , '" & rsEmpAllDed("PRESENTRATE") & "'" & _
									 "          , 0" & _
									 "          , '" & Session("CodeName") & "'" & _
									 "          , '" & Session("CodeName") & "'" & _
									 "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									 "          , '" & Session("fcBA") & "'" & _
									 "          , SYSDATE" & _
									 "          , TO_CHAR(SYSDATE, 'HH24:MI'))"
						connstr.execute(qAllowance)
					end if
				  

				  
					'	Insert the History of EmpallDed Data 
					' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
					
					'**** Remarks by UKIE, to disable history, 19 April 2016
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
														' rsEmpAllDed("PRESENTRATE") & "," & _
														' " '" & Session("CodeName") & "'," & _
														' " '" & Session("CodeName") & "'," & _
														' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
														' " sysdate, " & _
														' " to_char(sysdate, 'HH24:MI') " & _
														' " ) "
				  
					' connstr.execute(qAllowanceHistory)
				END IF
				
			  elseif rsEmpAllDed("PRESENTGRADE")="KHT" then 
					'Karyawan Harian Tetap 2 Ketegori:
					'    1. Dibayar Bulanan langsung dari Presentrate
					'    2. Dibayar Harian dikali tanggal hitam kalender
					' Note: Kategori dibuat dikolom Type_3 pada allowded A01, DEFAULT untuk Kategori 1, H untuk Kategori 2
					
					
					'***** Update Perhitungan Upah KHT di bagi dalam 3 kategori
					' 1. Dibayar Bulanan.
					'      - Kode Allowded adalah A01
					'      - Kolom Type_3 diisi M
					'      - Maka Perhitungan gaji langsung mengambil kolom presentrate pada table employee.
					' 2. Dibayar Harian versi asiatic.
					'      - Kode Allowded adalah A01
					'      - Kolom Type_3 diisi D1
					'      - Kolom Rate_3 diisi 25(sebagai pembagi)
					'      - Maka Perhitungan gaji mengambil kolom presentrate pada table employee dibagi rate_3 dikali dengan jumlah hari pada table BUDGET_TRANSACTION_WORKINGDAY
					'        BUDGET_TRANSACTION_HOLIDAY
					' 3. Dibayar Harian versi THIP.
					'      - Kode Allowded adalah A01
					'      - Kolom Type_3 diisi D2
					'      - Attendensi KJ, WH, WS, KK, P2, P3, S1, S2, H1, H2, ML, CT
					'      - Maka Perhitungan gaji mengambil kolom presentrate pada table employee dibagi Jumlah hari pada bulan yang di kalkulasi payroll di kali jumlah mandays
					
					
					'***** Update Versi 2 Perhitungan Upah KHT di bagi dalam 2 kategori
					' 1. Dibayar Bulanan.
					'      - Kode Allowded adalah A01
					'      - Kolom Type_3 diisi M
					'      - Maka Perhitungan gaji langsung mengambil kolom presentrate pada table employee.
					' 2. Dibayar Harian.
					'      - Kode Allowded adalah A01
					'      - Kolom Type_3 diisi D
					'      - Kolom Rate_3 => sebagai pembagi, diisi 0 (default) maka = 25, 100 maka = Jumlah Hari dalam sebulan, atau sesuai pembagi diisikan angka pembaginya
					'      - Attendensi yang dibayar bisa dilihat pada tabel Parameter dengan parcode E0102, Kolom    diisikan PAID = Attendensi yang dibayar dan NOPAID = Attendensi tidak dibayar
					'      - Maka Perhitungan gaji mengambil kolom presentrate pada table employee dibagi Jumlah hari pada Kolom Rate_3 dikalikan jumlah Attendensi.
					
					
					If val_kategori_kht = "D" Then
						arr_data_kht_eof = true
						objectSQL = " SELECT * FROM (" & _
						            " SELECT fcba" & _
									"        , fddate" & _
									"        , employeecode" & _
									"        , attendance" & _
									"        , mandays" & _
									"        , payroll_gaji_api.get_actual_mandays (fcba" & _
									"                                                , fddate" & _
									"                                                , attendance" & _
									"                                                , mandays) act_mandays" & _
									" FROM ( SELECT fcba" & _
									"               , fddate" & _
									"               , employeecode" & _
									"               , attendance" & _
									"               , SUM (mandays) mandays" & _
									"        FROM attendance_gad" & _
									"        WHERE employeecode = '" & rsEmpAllDed("FCCODE") & "'" & _
									"              AND fcba = '" & Session("fcBA") & "'" & _
									"              AND fddate = to_date('" & tglnow & "','mm/dd/yyyy') " & _
									"              AND INSTR('" & param_paidmandays & "', attendance) <> 0" & _
									"        GROUP BY fcba" & _
									"                 , fddate" & _
									"                 , employeecode" & _
									"                 , attendance" & _
									" )" & _
									" UNION ALL" & _
									" SELECT fcba, fddate, employeecode, attendance, 0 mandays, count(1) act_mandays" & _
									" FROM ( SELECT fcba, employeecode, fddate, attendance" & _
									"        FROM attendance_gad" & _
									"        WHERE employeecode = '" & rsEmpAllDed("FCCODE") & "'" & _
									"              AND fcba = '" & Session("fcBA") & "'" & _
									"              AND fddate = to_date('" & tglnow & "','mm/dd/yyyy') " & _
									"              AND INSTR('" & param_paidcount & "', attendance) <> 0" & _
									"        GROUP BY fcba, employeecode, fddate, attendance" & _
									" )" & _
									" GROUP BY fcba, employeecode, fddate, attendance" & _
									" ) ORDER BY fddate"

						'response.write objectSQL
						Set objectRS = Server.CreateObject("ADODB.Recordset")
						objectRS.Open objectSQL, connstr
						If Not objectRS.EOF then
							arr_data_kht = objectRS.getrows()
							arr_data_kht_eof = false
						End if
						objectRS.Close
						set objectRS = Nothing
						
						if not arr_data_kht_eof then
							for yyi=0 to ubound(arr_data_kht,2)
								count_data = count_data +cdbl(to_nol(arr_data_kht(5,yyi)))
							next
						end if
						
						if not arr_data_kht_eof then
							for yyi=0 to ubound(arr_data_kht,2)
								nilaiGaji = Round((emp_presentrate/val_divider)*cdbl(to_nol(arr_data_kht(5,yyi))), 0)
								
								curr_data = curr_data +cdbl(to_nol(arr_data_kht(5,yyi)))
								if count_data <= cdbl(to_nol(val_divider)) then
									if cdbl(to_nol(val_divider)) = curr_data then
										nilaiGaji = emp_presentrate -in_empallded
									else
										
										in_empallded = in_empallded +nilaiGaji
									end if
								end if
								
								maxLineNo = maxLineNo + 1
								qAllowance = " INSERT INTO empallded ( lineno, empcode, allowdedcode, adhocdate, amount, fixedallded, fcentry, fcedit, fcip, fcba, lastupdate, lasttime)" & _
                                             " VALUES ( '" & maxLineNo & "'" & _
											 "          , '" & rsEmpAllDed("FCCODE") & "'" & _
											 "          , 'A01'" & _
											 "          , TO_DATE('" & arr_data_kht(1,yyi) & "','mm/dd/yyyy')" & _
											 "          , '" & nilaiGaji & "'" & _
											 "          , 0" & _
											 "          , '" & Session("CodeName") & "'" & _
											 "          , '" & Session("CodeName") & "'" & _
											 "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
											 "          , '" & Session("fcBA") & "'" & _
											 "          , SYSDATE" & _
											 "          , TO_CHAR(SYSDATE, 'HH24:MI'))" 
								'response.write qAllowance
								connstr.execute(qAllowance)
							next
						end if
					  

					  
						'	Insert the History of EmpallDed Data 
						' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
						'**** Remarks by UKIE, to disable history, 19 April 2016
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
															' nilaiGaji & "," & _
															' " '" & Session("CodeName") & "'," & _
															' " '" & Session("CodeName") & "'," & _
															' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
															' " sysdate, " & _
															' " to_char(sysdate, 'HH24:MI') " & _
															' " ) "
					  
						' connstr.execute(qAllowanceHistory)				
						
					Else
						if cdbl(to_nol(rsEmpAllDed("PRESENTRATE"))) > 0 then
							maxLineNo = maxLineNo + 1
							qAllowance = " INSERT INTO empallded (lineno, empcode, allowdedcode, adhocdate, amount, fixedallded, fcentry, fcedit, fcip, fcba, lastupdate, lasttime)" & _
                                         " VALUES ( '" & maxLineNo & "'" & _
										 "          , '" & rsEmpAllDed("FCCODE") & "'" & _
										 "          , 'A01'" & _
										 "          , TO_DATE('" & Session("v_Employee") & "', 'mm/dd/yyyy')" & _
										 "          , '" & rsEmpAllDed("PRESENTRATE") & "'" & _
										 "          , 0" & _
										 "          , '" & Session("CodeName") & "'" & _
										 "          , '" & Session("CodeName") & "'" & _
										 "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
										 "          , '" & Session("fcBA") & "'" & _
										 "          , SYSDATE" & _
										 "          , TO_CHAR(SYSDATE, 'HH24:MI'))"
							connstr.execute(qAllowance)
						end if
					  

					  
						'	Insert the History of EmpallDed Data 
						' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
						'**** Remarks by UKIE, to disable history, 19 April 2016
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
															' rsEmpAllDed("PRESENTRATE") & "," & _
															' " '" & Session("CodeName") & "'," & _
															' " '" & Session("CodeName") & "'," & _
															' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
															' " sysdate, " & _
															' " to_char(sysdate, 'HH24:MI') " & _
															' " ) "
					  
						' connstr.execute(qAllowanceHistory)
						
					End If
				  
			  Elseif rsEmpAllDed("PRESENTGRADE")="KHL"  Then 
				'Karyawan Harian Lepas
				' Modify by NAF,  
				arr_data_khl_eof = true
				qGAD = " SELECT fcba" & _
					   "        , employeecode" & _
					   "        , fddate" & _
					   "        , attendance" & _
					   "        , NVL(SUM(mandays), 0) mandays" & _
					   " FROM attendance_gad" & _
					   " WHERE employeecode = '" & rsEmpAllDed("FCCODE") & "'" & _
					   "       AND fcba = '" & Session("fcBA") & "'" & _
                       "       AND fddate = to_date('" & tglnow & "','mm/dd/yyyy') " & _
					   "       AND attendance IN ('KJ', 'DL','KK')" & _
					   " GROUP BY fcba, employeecode, fddate, attendance" & _
					   " ORDER BY fddate"

				 'response.write qGAD
				 Set rsGAD = Server.CreateObject("ADODB.Recordset")
				 rsGAD.Open qGAD, connstr
				 if not rsGAD.eof then
					arr_data_khl = rsGAD.getrows()
					arr_data_khl_eof = false
				 end if
				 rsGAD.Close
				 set rsGAD = Nothing
				 
				 
				 if not arr_data_khl_eof then
					for yyi=0 to ubound(arr_data_khl,2)
						nAmount = emp_presentrate * cdbl(to_nol(arr_data_khl(4,yyi)))
						
						if nAmount > 0 then
							maxLineNo = maxLineNo + 1
							qAllowance = " INSERT INTO empallded (lineno, empcode, allowdedcode, adhocdate, amount, fcentry, fcedit, fcip, fcba, lastupdate, lasttime)" & _
							             " VALUES ( '" & maxLineNo & "'" & _
										 "          , '" & rsEmpAllDed("FCCODE") & "'" & _
										 "          , 'A01'" & _
										 "          , TO_DATE('" & arr_data_khl(2,yyi) & "', 'mm/dd/yyyy')" & _
										 "          , '" & nAmount & "'" & _
										 "          , '" & Session("CodeName") & "'" & _
										 "          , '" & Session("CodeName") & "'" & _
										 "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
										 "          , '" & Session("fcBA") & "'" & _
										 "          , SYSDATE" & _
										 "          , TO_CHAR(SYSDATE, 'HH24:MI'))"
							connstr.execute(qAllowance)
						end if
					next
				 end if
				 

				 
				 
					
					
					'**** Remarks by UKIE, to disable history, 19 April 2016
					' qGADDetail =	" SELECT " & _
										' " FDDATE, EMPLOYEECODE, JOBCODE, LOCATIONTYPE, LOCATIONCODE, FCBA, " & _
										' " NVL(DECODE(GreaterThan(MANDAYS,0),'TRUE', mandays ), 0) MANDAYS FROM(" & _
									' " SELECT FCBA, EMPLOYEECODE, FDDATE, ATTENDANCE, JOBCODE, LOCATIONTYPE, LOCATIONCODE,   NVL(MANDAYS,0) MANDAYS" & _
									' " from attendance_GAD WHERE ATTENDANCE='KJ' " & _
									' " )" & _
									' " where EMPLOYEECODE = '" & rsEmpAllDed("FCCODE") & "' and " & _
									' " to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
									' " and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and fcBA='" & Session("fcBA") & "'"
									
					' Set rsReceive = Server.CreateObject("ADODB.Recordset")
					' rsReceive.CursorType = 3
					' rsReceive.CursorLocation = 3
					' ' response.write qGADDetail
					' rsReceive.Open qGADDetail, connstr
					' if not rsReceive.eof then
						' detailLineNo = 0
						' Do While Not rsReceive.eof	

							'	Insert the History of EmpallDed Data 
							' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
							'**** Remarks by UKIE, to disable history, 19 April 2016
							' txtRefStringValues_Receive1 = rsReceive("FDDATE") & txtRefStringSeparator_Receive1 & _
														 ' rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive1 & _
														 ' rsReceive("JOBCODE") & txtRefStringSeparator_Receive1 & _
														 ' rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive1 & _
														 ' rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive1 & _
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
																' " MANDAYS, " & _
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
																' " '" & tempProcessCode1 & "'," & _
																' " '" & tempProcessNameCode1 & "'," & _
																' " '" & tempRefCodeReceive1 & "'," & _
																' " '" & txtRefNumberKeys_Receive1 & "'," & _
																' " '" & txtRefStringSeparator_Receive1 & "'," & _
																' " '" & txtRefStringKeys_Receive1 & "'," & _
																' " '" & txtRefStringValues_Receive1 & "'," & _
																' cdbl(rsReceive("MANDAYS")) & "," & _
																' cdbl(rsEmpAllDed("PRESENTRATE")) & "," & _
																' cdbl(rsReceive("MANDAYS")) * cdbl(rsEmpAllDed("PRESENTRATE")) & "," & _ 
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
						
			  else 'ConTract Dan Lain - Lain adalah Dari Present Rate
					maxLineNo = maxLineNo + 1
					qAllowance = " INSERT INTO empallded (lineno, empcode, allowdedcode, adhocdate, amount, fixedallded, fcentry, fcedit, fcip, fcba, lastupdate, lasttime)" & _
					             " VALUES ( '" & maxLineNo & "'" & _
								 "          , '" & rsEmpAllDed("FCCODE") & "'" & _
								 "          , 'A01'" & _
								 "          , TO_DATE('" & Session("v_Employee") & "', 'mm/dd/yyyy')" & _
								 "          , '" & rsEmpAllDed("PRESENTRATE") & "'" & _
								 "          , 0" & _
								 "          , '" & Session("CodeName") & "'" & _
								 "          , '" & Session("CodeName") & "'" & _
								 "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
								 "          , '" & Session("fcBA") & "'" & _
								 "          , SYSDATE" & _
								 "          , TO_CHAR(SYSDATE, 'HH24:MI'))"
					connstr.execute(qAllowance)
				  
				  '	Insert the History of EmpallDed Data 
					' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
					'**** Remarks by UKIE, to disable history, 19 April 2016
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
														' rsEmpAllDed("PRESENTRATE") & "," & _
														' " '" & Session("CodeName") & "'," & _
														' " '" & Session("CodeName") & "'," & _
														' " '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
														' " sysdate, " & _
														' " to_char(sysdate, 'HH24:MI') " & _
														' " ) "
				  
					' connstr.execute(qAllowanceHistory)
				  
			  End if
			  response.Flush()
			  rsEmpAllDed.MoveNext
		  loop
		End if
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
		
END IF
%>