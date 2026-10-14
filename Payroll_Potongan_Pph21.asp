<%
	'Amandemen
	'1. Allowded Code Untuk Potongan PPh 21 = D03 
	' By NAF

    '26 May 2014, Permintaan dari THIP Project :
    ' Yaitu : 
    ' 2. Tunjangan PPh 21 oleh Perusahaan ==> A58
    ' By NAF ==> ini belum ada tolong dibuatkan


	'New Definition
	Dim temp_salaries : temp_salaries = 0
	Dim temp_allowances : temp_allowances = 0
	Dim temp_irregular_allowances : temp_irregular_allowances = 0
	Dim temp_ded_employment : temp_ded_employment = 0
	Dim temp_ded_retirement : temp_ded_retirement = 0
	
	Dim tempNettoMonth : tempNettoMonth = 0
	Dim tempNettoYear : tempNettoYear = 0
	
	Dim array_empallded, array_empallded_eof, array_empallded_record
	Dim insertdatatotable, GlobalRS, GlobalSQL, temp_tunjpph_ty

	array_empallded_eof = true
	
	
	Dim temp_gajibruto_tm : temp_gajibruto_tm = 0
	Dim temp_potjabatan_tm : temp_potjabatan_tm = 0
	Dim temp_potastek_tm : temp_potastek_tm = 0
	Dim temp_gajinetto_tm : temp_gajinetto_tm = 0
	Dim temp_ptkp_tm : temp_ptkp_tm = 0
	Dim temp_pkp_tm : temp_pkp_tm = 0
	Dim temp_pph21_tm : temp_pph21_tm = 0
	Dim temp_tunjpph21_tm : temp_tunjpph21_tm = 0

	Dim temp_gajibruto_ty : temp_gajibruto_ty = 0
	Dim temp_potjabatan_ty : temp_potjabatan_ty = 0
	Dim temp_potastek_ty : temp_potastek_ty = 0
	Dim temp_gajinetto_ty : temp_gajinetto_ty = 0
	Dim temp_ptkp_ty : temp_ptkp_ty = 0
	Dim temp_pkp_ty : temp_pkp_ty = 0
	Dim temp_pph21_ty : temp_pph21_ty = 0
	Dim temp_tunjpph21_ty : temp_tunjpph21_ty = 0
	
	Dim temp_npwp : temp_npwp = "FALSE"
	
	

	

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
	
	' Agar Tunjangan Pph 21 keluar dan dihitung sistem, adalah sbb:
	' System must track it as an AUTOMATIC Process dan includedpayrollprocessing = TRUE
	' Kode Allowded: A58
	' Parameter yang menentukan ada pembiayaan yang di tanggung oleh perusahaan, dalam hal ini TYPE_2, jika ada isi TRUE jika tidak isi FALSE
	
	Dim tempIncludedPayrollProcessingTunjanganPph21 : tempIncludedPayrollProcessingTunjanganPph21 = "FALSE"
	Dim ParamIsPaidByCompany : ParamIsPaidByCompany = "FALSE"
	Dim tempFinalTaxTunjanganPph21 : tempFinalTaxTunjanganPph21 = 0
	Dim rsIsEmployeePaidByCompany, strIsEmployeePaidByCompany
	Dim IsEmployeePaidByCompany : IsEmployeePaidByCompany = "FALSE"
	tempIncludedPayrollProcessingTunjanganPph21 = "FALSE"
	
	Dim sqljht, rsjht, temp_amount_jht
	
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter = " select includedpayrollprocessing" & _
	             "        , man_auto" & _
				 "        , type_2 ParamIsPaidByCompany" & _
		         " from allowded" & _
				 " where fccode = 'A58'" & _
				 "       and fcba = '" & Session("fcBA") & "'" & _
				 "       and man_auto = 'A'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		 tempIncludedPayrollProcessingTunjanganPph21= rsParameter("includedpayrollprocessing")
		 ParamIsPaidByCompany = rsParameter("ParamIsPaidByCompany")
	else
		 tempIncludedPayrollProcessingTunjanganPph21 = "FALSE"
		 ParamIsPaidByCompany = "FALSE"
	End if
	rsParameter.Close
	set rsParameter = Nothing
	


	' System must track it as an AUTOMATIC Process 
	' Jabatan
	
	Dim temp_max_Jabatan : temp_max_Jabatan = 0
	tempIncludedPayrollProcessingA = "FALSE"
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter =" Select includedpayrollprocessing, man_auto, rate_1, nvl(max_amt, 0) max_amt" & _
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
		 temp_max_Jabatan = cdbl(rsParameter("max_amt"))
	else
		 tempIncludedPayrollProcessingA = "FALSE"
		 tempPercentageJabatanDeduction= 0
		 temp_max_Jabatan = 0
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
	
	
	'**** Normalisasi Data Data Taxledger ****'
'	GlobalSQL = " select distinct empcode" & _
'				" from taxledger" & _
'				" where fcba = '" & Session("fcBA") & "'" & _
'				"       and year = " & Year(Session("v_Employee")) & "" & _
'				"       and month = " & Month(Session("v_Employee")) & ""

		'response.Write GlobalSQL
		'response.Flush()			    
'		Set GlobalRS = Server.CreateObject("ADODB.Recordset")
'		GlobalRS.CursorType = 3
'		GlobalRS.CursorLocation = 3
'		GlobalRS.Open GlobalSQL, connstr
'		if not GlobalRS.eof then		
'			array_empallded = GlobalRS.getrows()
'			array_empallded_eof = false
'		end if
'		GlobalRS.Close
'		set GlobalRS = Nothing		
	
'		if not array_empallded_eof then
'			For array_empallded_record = LBound(array_empallded, 2) To UBound(array_empallded, 2)
'				maxLineNo = maxLineNo + 1
				connstr.BeginTrans
				insertdatatotable = " delete from taxledger" & _
									" where fcba = '" & Session("fcBA") & "'" & _
									"       and year = " & Year(Session("v_Employee")) & "" & _
									"       and month = " & Month(Session("v_Employee")) & "" 
'									& _
'									"       and empcode = '" & array_empallded(0, array_empallded_record) & "'"
						
				connstr.execute(insertdatatotable)
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
'			Next
'		End if
		
'	array_empallded_eof = true	
'	GlobalSQL = " SELECT DISTINCT empcode" & _
'				" FROM taxledger_thr" & _
'				" WHERE fcba = '" & Session("fcBA") & "'" & _
'				"       AND year = " & Year(Session("v_Employee")) & "" & _
'				"       AND month = " & Month(Session("v_Employee")) & ""

		'response.Write GlobalSQL
		'response.Flush()			    
'		Set GlobalRS = Server.CreateObject("ADODB.Recordset")
'		GlobalRS.CursorType = 3
'		GlobalRS.CursorLocation = 3
'		GlobalRS.Open GlobalSQL, connstr
'		if not GlobalRS.eof then		
'			array_empallded = GlobalRS.getrows()
'			array_empallded_eof = false
'		end if
'		GlobalRS.Close
'		set GlobalRS = Nothing		
	
'		if not array_empallded_eof then
'			For array_empallded_record = LBound(array_empallded, 2) To UBound(array_empallded, 2)
'				maxLineNo = maxLineNo + 1
				connstr.BeginTrans
				insertdatatotable = " DELETE FROM taxledger_thr" & _
									" WHERE fcba = '" & Session("fcBA") & "'" & _
									"       AND year = " & Year(Session("v_Employee")) & "" & _
									"       AND month = " & Month(Session("v_Employee")) & "" 
'									& _
'									"       AND empcode = '" & array_empallded(0, array_empallded_record) & "'"
						
				connstr.execute(insertdatatotable)
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
'			Next
'		End if

'	array_empallded_eof = true	
'	GlobalSQL = " SELECT DISTINCT empcode" & _
'				" FROM taxledger_bonus" & _
'				" WHERE fcba = '" & Session("fcBA") & "'" & _
'				"       AND year = " & Year(Session("v_Employee")) & "" & _
'				"       AND month = " & Month(Session("v_Employee")) & ""

		'response.Write GlobalSQL
		'response.Flush()			    
'		Set GlobalRS = Server.CreateObject("ADODB.Recordset")
'		GlobalRS.CursorType = 3
'		GlobalRS.CursorLocation = 3
'		GlobalRS.Open GlobalSQL, connstr
'		if not GlobalRS.eof then		
'			array_empallded = GlobalRS.getrows()
'			array_empallded_eof = false
'		end if
'		GlobalRS.Close
'		set GlobalRS = Nothing		
	
'		if not array_empallded_eof then
'			For array_empallded_record = LBound(array_empallded, 2) To UBound(array_empallded, 2)
'				maxLineNo = maxLineNo + 1
				connstr.BeginTrans
				insertdatatotable = " DELETE FROM taxledger_bonus" & _
									" WHERE fcba = '" & Session("fcBA") & "'" & _
									"       AND year = " & Year(Session("v_Employee")) & "" & _
									"       AND month = " & Month(Session("v_Employee")) & "" 
'									& _
'									"       AND empcode = '" & array_empallded(0, array_empallded_record) & "'"
						
				connstr.execute(insertdatatotable)
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
'			Next
'		End if
	

				connstr.BeginTrans
				insertdatatotable = " DELETE FROM empallded" & _
									" WHERE fcba = '" & Session("fcBA") & "'" & _
									"       AND TO_NUMBER (TO_CHAR (adhocdate, 'YYYY')) = " & Year(Session("v_Employee")) & "" & _
									"       AND TO_NUMBER (TO_CHAR (adhocdate, 'MM')) = " & Month(Session("v_Employee")) & "" & _
                                    "       AND allowdedcode in ('D03','A58') "						
				connstr.execute(insertdatatotable)
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
'
	response.Write "Normalisasi ..."& "<br>" 
	response.Flush()
	
	if tempIncludedPayrollProcessing = "TRUE" then
	
		qQuery = " select fcba" & _
				 "        , last_day(to_date('" & Month(Session("v_Employee")) & "/01/" & Year(Session("v_Employee")) & "', 'mm/dd/yyyy')) fddate" & _
				 "        , empcode employeecode" & _
				 "        , maritalstatus" & _
				 "        , IncomeBeforeTax" & _
				 "        , untaxableincome_api.getTotalUti(maritalstatus, fcba) totalUti" & _
                 " from(SELECT a.fcba" & _
				 "             , a.empcode" & _
				 "             , taxledger_api.get_valid_taxstatus(" & Year(Session("v_Employee")) & ", " & Month(Session("v_Employee")) & ", a.fcba, a.empcode, b.taxstatus) maritalStatus" & _
				 "             , SUM ( Allowded_APi.GetAlldedValue (a.allowdedcode, amount, a.fcba)) IncomeBeforeTax " & _
                 "      FROM empallded a" & _
				 "           , employee b" & _
                 "      WHERE a.empcode = b.fccode" & _
                 "            AND TO_NUMBER (TO_CHAR (a.adhocdate, 'YYYY')) = " & Year(Session("v_Employee")) & "" & _
                 "            AND TO_NUMBER (TO_CHAR (a.adhocdate, 'MM')) = " & Month(Session("v_Employee")) & "" & _
                 "            AND allowded_api.GetIncludedTaxCalculation (A.ALLOWDEDCODE, a.fcba) = 'TRUE'" & _
				 "            AND employee_api.GetIncludedTaxCalculation (a.empcode, a.fcba) = 'TRUE'" & _
                 "            AND b.dateterminate is null" & _
                 "            AND a.fcba = '" & Session("fcBA") & "'" & _
                 "            AND a.fcba = b.fcba" & _
                 "      GROUP BY a.fcba" & _
                 "              , a.empcode" & _
                 "              , b.taxstatus " & _
                 " ) " 
		
		'response.Write "<br><br>" & " Recordset " & qQuery & "<br><br>"
		'response.Flush()			    
		Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
		rsEmpAllDed.CursorType = 3
		rsEmpAllDed.CursorLocation = 3
		response.Write "Calculate Tax for PPh 21 ..."& "<br>" 
		response.Flush()
		rsEmpAllDed.Open qQuery, connstr
		if not rsEmpAllDed.eof then		
			array_empallded = rsEmpAllDed.getrows()
			array_empallded_eof = false
		end if
		rsEmpAllDed.Close
		set rsEmpAllDed = Nothing
		
		
		
		if not array_empallded_eof then
		
			For array_empallded_record = LBound(array_empallded, 2) To UBound(array_empallded, 2)
			
				temp_gajibruto_tm = 0
				temp_potjabatan_tm = 0
				temp_potastek_tm = 0
				temp_gajinetto_tm = 0
				temp_ptkp_tm = 0
				temp_pkp_tm = 0
				temp_pph21_tm = 0
				temp_tunjpph21_tm = 0

				temp_gajibruto_ty = 0
				temp_potjabatan_ty = 0
				temp_potastek_ty = 0
				temp_gajinetto_ty = 0
				temp_ptkp_ty = 0
				temp_pkp_ty = 0
				temp_pph21_ty = 0
				temp_tunjpph21_ty = 0

			
				' **** MENCARI GAJI BRUTO **** '

				temp_gajibruto_tm = Cdbl(array_empallded(4, array_empallded_record))
				
				
				temp_amount_jht = 0
				sqljht = " select payroll_jamsostek_dtl.get_jht(" & Year(Session("v_Employee")) & "" & _
				         "                                      , " & Month(Session("v_Employee")) & "" & _
						 "                                      , '" & Session("fcBA") & "'" & _
						 "                                      , '" & array_empallded(2, array_empallded_record) & "'" & _
						 "                                      , 'A17') jht" & _
						 " from dual"
				Set rsjht = Server.CreateObject("ADODB.Recordset")
				rsjht.Open sqljht, connstr
				
				if not rsjht.EOF then
					temp_amount_jht = cdbl( rsjht("jht") )
				else 
					temp_amount_jht = 0
				end if
				rsjht.Close
				Set rsjht = Nothing
				
				temp_amount_jp = 0
				sqljp = " select payroll_jamsostek_dtl.get_jp(" & Year(Session("v_Employee")) & "" & _
                        "                                      , " & Month(Session("v_Employee")) & "" & _
                        "                                      , '" & Session("fcBA") & "'" & _
                        "                                      , '" & array_empallded(2, array_empallded_record) & "'" & _
                        "                                      , 'A17') jp" & _
                        " from dual"
				Set rsjp = Server.CreateObject("ADODB.Recordset")
				rsjp.Open sqljp, connstr
				
				if not rsjp.EOF then
					temp_amount_jp = cdbl( rsjp("jp") )
				else 
					temp_amount_jp = 0
				end if
				rsjp.Close
				Set rsjp = Nothing
				
				temp_gajibruto_tm = temp_gajibruto_tm - (temp_amount_jht + temp_amount_jp)
				temp_gajibruto_ty = temp_gajibruto_tm * 12
				
				' **** MENCARI POTONGAN JABATAN **** '
				' *** Validasi Jika KHL tidak dapat potongan Jabatan
				
				qgetEmpPresentGrade = "SELECT EMPLOYEE_API.get_presentgrade('" & Session("fcBA") & "', '" & array_empallded(2, array_empallded_record) & "')PresentGrade FROM DUAL"
						
					Set rsqgetEmpPresentGrade = Server.CreateObject("ADODB.Recordset")
					rsqgetEmpPresentGrade.CursorType = 3
					rsqgetEmpPresentGrade.CursorLocation = 3
					rsqgetEmpPresentGrade.Open qgetEmpPresentGrade, connstr

					if not rsqgetEmpPresentGrade.eof then
						EmpPresentGrade = rsqgetEmpPresentGrade("PresentGrade")
					else
						EmpPresentGrade = "KHT"
					end if
					rsqgetEmpPresentGrade.close
					set rsqgetEmpPresentGrade = Nothing		
					
					if EmpPresentGrade <> "KHL" then
						temp_potjabatan_tm = temp_gajibruto_tm * tempPercentageJabatanDeduction / 100
						
						if temp_potjabatan_tm > temp_max_Jabatan then
							temp_potjabatan_tm = temp_max_Jabatan
						else
							temp_potjabatan_tm = temp_potjabatan_tm
						end if
					else
							temp_potjabatan_tm = 0
					end if 
					
				'*** Validasi Jika KHL tidak dapat potongan Jabatan
				
				temp_potjabatan_ty = temp_potjabatan_tm * 12
				
				' **** MENCARI POTONGAN ASTEK **** '
				
				strQueryData = " SELECT " & _
									   " NVL (SUM (NVL (DECODE (B.ALLDED, 'D', -1 * ABS (amount), amount), 0)), 0) totalAmount " & _
									" FROM " & _
									   " empallded a, allowded b " & _
									" WHERE " & _
									   " a.allowdedcode = b.fccode AND " & _
									   " TO_NUMBER (TO_CHAR (a.adhocdate, 'YYYYMM')) = to_number(to_char(" & Year(Session("v_Employee")) & ")||lpad ( to_char( " & Month(Session("v_Employee")) & "),2,'0')) AND " & _
									   " a.allowdedcode IN ('A17', 'D02') AND" & _
									   " a.empcode =  '" &  array_empallded(2, array_empallded_record) & "' AND " & _
									   " a.fcba = b.fcba AND " & _
									   " a.fcba = '" & Session("fcBA") & "' " 		
				
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
				
				rsQueryData.Close
				set rsQueryData = Nothing
				
				temp_potastek_tm = totalAstekDeduction * -1
				temp_potastek_ty = temp_potastek_tm * 12				
				
				' **** HITUNG GAJI NETTO **** '
				
				temp_gajinetto_tm = temp_gajibruto_tm - (temp_potjabatan_tm + temp_potastek_tm)
				temp_gajinetto_ty = temp_gajinetto_tm * 12
				
				' **** MENCARI PTKP **** '
				
				temp_ptkp_tm = Cdbl(to_nol(array_empallded(5, array_empallded_record)))
				temp_ptkp_ty = temp_ptkp_tm * 12
				
				' **** HITUNG PKP **** '
				temp_pkp_tm = temp_gajinetto_tm - temp_ptkp_tm
				temp_pkp_ty = temp_pkp_tm * 12
				
				
				
				' **** INSERT DATA TO LEDGER MONTHLY **** '
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'gajibruto_tm'" & _
				                    "          , " & temp_gajibruto_tm & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'Gaji Bruto Sebulan'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'potjabatan_tm'" & _
				                    "          , " & temp_potjabatan_tm & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'Potongan Jabatan Sebulan'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'potastek_tm'" & _
				                    "          , " & temp_potastek_tm & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'Potongan ASTEK Sebulan'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'gajinetto_tm'" & _
				                    "          , " & temp_gajinetto_tm & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'Gaji Netto Sebulan'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'ptkp_tm'" & _
				                    "          , " & temp_ptkp_tm & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'PTKP Sebulan'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'pkp_tm'" & _
				                    "          , " & temp_pkp_tm & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'PKP Sebulan'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				' **** INSERT DATA TO LEDGER ANNUALLY **** '
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'gajibruto_ty'" & _
				                    "          , " & temp_gajibruto_ty & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'Gaji Bruto Setahun'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'potjabatan_ty'" & _
				                    "          , " & temp_potjabatan_ty & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'Potongan Jabatan Setahun'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'potastek_ty'" & _
				                    "          , " & temp_potastek_ty & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'Potongan ASTEK Setahun'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'gajinetto_ty'" & _
				                    "          , " & temp_gajinetto_ty & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'Gaji Netto Setahun'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'ptkp_ty'" & _
				                    "          , " & temp_ptkp_ty & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'PTKP Setahun'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				insertdatatotable = " insert into taxledger ( empcode" & _
				                    "                         , year" & _
				                    "                         , month" & _
				                    "                         , processid" & _
				                    "                         , amount" & _
				                    "                         , fcentry" & _
				                    "                         , fcedit" & _
				                    "                         , fcip" & _
				                    "                         , fcba" & _
				                    "                         , remarks" & _
				                    "                         , lastupdate" & _
				                    "                         , lasttime)" & _
				                    " values ( '" & array_empallded(2, array_empallded_record) & "'" & _
				                    "          , " & Year(Session("v_Employee")) & "" & _
				                    "          , " & Month(Session("v_Employee")) & "" & _
				                    "          , 'pkp_ty'" & _
				                    "          , " & temp_pkp_ty & "" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Session("CodeName") & "'" & _
				                    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				                    "          , '" & Session("fcBA") & "'" & _
									"          , 'PKP Setahun'" & _
				                    "          , sysdate," & _
				                    "           to_char(sysdate,'HH24:MI'))"
							
				connstr.execute(insertdatatotable)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				
				if temp_pkp_ty > 0 then
					' **** HITUNG PPH 21 BULANAN **** '
					GlobalSQL = " select RANGETAXGROUP_API.GetValueProgressive(" & temp_pkp_ty & "" & _
							    "                                              , " & temp_ptkp_ty & "" & _
							    "                                              , '" & Session("fcBA") & "') result" & _
							    " from dual"
									 
					Set GlobalRS = Server.CreateObject("ADODB.Recordset")
					GlobalRS.CursorType = 3
					GlobalRS.CursorLocation = 3
					GlobalRS.Open GlobalSQL, connstr
					if not GlobalRS.eof then
						if GlobalRS ("result") <> "" then 
							temp_pph21_tm = cdbl(GlobalRS ("result"))				
						else
							temp_pph21_tm = 0
						end if
					else
						temp_pph21_tm = 0
					end if
					GlobalRS.Close
					set GlobalRS = Nothing
					
					' **** HITUNG TUNJANGAN PPH 21 BULANAN **** '
					GlobalSQL = " select taxledger_api.get_tax_subsidy(" & temp_pkp_ty & ") result" & _
							    " from dual"
									 
					Set GlobalRS = Server.CreateObject("ADODB.Recordset")
					GlobalRS.CursorType = 3
					GlobalRS.CursorLocation = 3
					GlobalRS.Open GlobalSQL, connstr
					if not GlobalRS.eof then
						if GlobalRS ("result") <> "" then 
							temp_tunjpph21_ty = cdbl(GlobalRS ("result"))				
						else
							temp_tunjpph21_ty = 0
						end if
					else
						temp_tunjpph21_ty = 0
					end if
					GlobalRS.Close
					set GlobalRS = Nothing
					
					temp_tunjpph21_tm = temp_tunjpph21_ty / 12					
					
				else
					temp_pph21_tm = 0
				end if
				
				temp_pph21_ty = temp_pph21_tm * 12
				
				' **** CARI PUNYA NPWP atau TIDAK **** '
				GlobalSQL = " select nvl(npwp, '000000000311000') result" & _
								 " from employee" & _
								 " where fccode = '" & array_empallded(2, array_empallded_record) & "'" & _
								 "       and fcba = '" & Session("fcBA") & "'"
								 
				Set GlobalRS = Server.CreateObject("ADODB.Recordset")
				GlobalRS.CursorType = 3
				GlobalRS.CursorLocation = 3
				GlobalRS.Open GlobalSQL, connstr
				if not GlobalRS.eof then
					if (IsNull(GlobalRS ("result")) Or GlobalRS ("result") = "") then 
						temp_npwp = "000000000311000"
					else
						temp_npwp = GlobalRS ("result")
					end if
				else
					temp_npwp = "000000000311000"
				end if
				GlobalRS.Close
				set GlobalRS = Nothing
				
				if temp_npwp = "000000000311000" then
					temp_tunjpph21_tm = temp_tunjpph21_tm * 120 / 100
					temp_tunjpph21_ty = temp_tunjpph21_ty * 120 / 100
					temp_pph21_tm = temp_pph21_tm * 120 / 100
					temp_pph21_ty = temp_pph21_ty * 120 / 100					
				end if
				

				
				If tempIncludedPayrollProcessingTunjanganPph21 = "TRUE" Then
					If ParamIsPaidByCompany = "TRUE" Then
						Set rsIsEmployeePaidByCompany = Server.CreateObject("ADODB.Recordset")
						strIsEmployeePaidByCompany = " select nvl(taxpaidbycompany, 'FALSE') taxpaidbycompany" & _
													 " from employee" & _
													 " where fccode = '" & array_empallded(2, array_empallded_record) & "'" & _
													 "       and fcba = '" & Session("fcBA") & "'"
													 
						'response.Write " strIsEmployeePaidByCompany = " & strIsEmployeePaidByCompany				  
						rsIsEmployeePaidByCompany.CursorType = 3
						rsIsEmployeePaidByCompany.CursorLocation = 3
						rsIsEmployeePaidByCompany.Open strIsEmployeePaidByCompany, connstr
						if not rsIsEmployeePaidByCompany.eof then
							 IsEmployeePaidByCompany= rsIsEmployeePaidByCompany("taxpaidbycompany")
						else
							 IsEmployeePaidByCompany = "FALSE"
						End if
						rsIsEmployeePaidByCompany.Close
						set rsIsEmployeePaidByCompany = Nothing

						If IsEmployeePaidByCompany = "TRUE" Then
							connstr.BeginTrans
							insertdatatotable = " insert into taxledger ( empcode" & _
												"                         , year" & _
												"                         , month" & _
												"                         , processid" & _
												"                         , amount" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , remarks" & _
												"                         , lastupdate" & _
												"                         , lasttime)" & _
												" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
												"          , " & Year(Session("v_Employee")) & "" & _
												"          , " & Month(Session("v_Employee")) & "" & _
												"          , 'pph21_tm'" & _
												"          , " & temp_tunjpph21_tm & "" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
												"          , '" & Session("fcBA") & "'" & _
												"          , 'Pph21 Sebulan'" & _
												"          , sysdate," & _
												"           to_char(sysdate,'HH24:MI'))"
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							connstr.BeginTrans
							insertdatatotable = " insert into taxledger ( empcode" & _
												"                         , year" & _
												"                         , month" & _
												"                         , processid" & _
												"                         , amount" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , remarks" & _
												"                         , lastupdate" & _
												"                         , lasttime)" & _
												" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
												"          , " & Year(Session("v_Employee")) & "" & _
												"          , " & Month(Session("v_Employee")) & "" & _
												"          , 'pph21_ty'" & _
												"          , " & temp_tunjpph21_ty & "" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
												"          , '" & Session("fcBA") & "'" & _
												"          , 'Pph21 Setahun'" & _
												"          , sysdate," & _
												"           to_char(sysdate,'HH24:MI'))"
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							connstr.BeginTrans
							insertdatatotable = " insert into taxledger ( empcode" & _
												"                         , year" & _
												"                         , month" & _
												"                         , processid" & _
												"                         , amount" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , remarks" & _
												"                         , lastupdate" & _
												"                         , lasttime)" & _
												" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
												"          , " & Year(Session("v_Employee")) & "" & _
												"          , " & Month(Session("v_Employee")) & "" & _
												"          , 'tunjpph21_tm'" & _
												"          , " & temp_tunjpph21_tm & "" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
												"          , '" & Session("fcBA") & "'" & _
												"          , 'Tunjangan Pph21 Sebulan'" & _
												"          , sysdate," & _
												"           to_char(sysdate,'HH24:MI'))"
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							connstr.BeginTrans
							insertdatatotable = " insert into taxledger ( empcode" & _
												"                         , year" & _
												"                         , month" & _
												"                         , processid" & _
												"                         , amount" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , remarks" & _
												"                         , lastupdate" & _
												"                         , lasttime)" & _
												" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
												"          , " & Year(Session("v_Employee")) & "" & _
												"          , " & Month(Session("v_Employee")) & "" & _
												"          , 'tunjpph21_ty'" & _
												"          , " & temp_tunjpph21_ty & "" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
												"          , '" & Session("fcBA") & "'" & _
												"          , 'Tunjangan Pph21 Setahun'" & _
												"          , sysdate," & _
												"           to_char(sysdate,'HH24:MI'))"
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							connstr.BeginTrans
							insertdatatotable = " update taxledger" & _
							                    " set amount = " & (temp_gajibruto_tm + temp_tunjpph21_tm) - (temp_potjabatan_tm + temp_potastek_tm) & "" & _
												" where fcba = '" & Session("fcBA") & "'" & _
												"       and year = " & Year(Session("v_Employee")) & "" & _
												"       and month = " & Month(Session("v_Employee")) & "" & _
												"       and empcode = '" & array_empallded(2, array_empallded_record) & "'" & _
												"       and processid = 'gajinetto_tm'"
							'response.write "<br/>" & array_empallded(2, array_empallded_record) & " tm : " & insertdatatotable
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							connstr.BeginTrans
							insertdatatotable = " update taxledger" & _
							                    " set amount = " & (temp_gajibruto_ty + temp_tunjpph21_ty) - (temp_potjabatan_ty + temp_potastek_ty) & "" & _
												" where fcba = '" & Session("fcBA") & "'" & _
												"       and year = " & Year(Session("v_Employee")) & "" & _
												"       and month = " & Month(Session("v_Employee")) & "" & _
												"       and empcode = '" & array_empallded(2, array_empallded_record) & "'" & _
												"       and processid = 'gajinetto_ty'"
							'response.write "<br/>" & array_empallded(2, array_empallded_record) & " ty : " & insertdatatotable
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							connstr.BeginTrans
							insertdatatotable = " update taxledger" & _
							                    " set amount = " & ( (temp_gajibruto_tm + temp_tunjpph21_tm) - (temp_potjabatan_tm + temp_potastek_tm) ) - temp_ptkp_tm & "" & _
												" where fcba = '" & Session("fcBA") & "'" & _
												"       and year = " & Year(Session("v_Employee")) & "" & _
												"       and month = " & Month(Session("v_Employee")) & "" & _
												"       and empcode = '" & array_empallded(2, array_empallded_record) & "'" & _
												"       and processid = 'pkp_tm'"
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							connstr.BeginTrans
							insertdatatotable = " update taxledger" & _
							                    " set amount = " & ( (temp_gajibruto_ty + temp_tunjpph21_ty) - (temp_potjabatan_ty + temp_potastek_ty) ) - temp_ptkp_ty & "" & _
												" where fcba = '" & Session("fcBA") & "'" & _
												"       and year = " & Year(Session("v_Employee")) & "" & _
												"       and month = " & Month(Session("v_Employee")) & "" & _
												"       and empcode = '" & array_empallded(2, array_empallded_record) & "'" & _
												"       and processid = 'pkp_ty'"
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							
						else
							connstr.BeginTrans
							insertdatatotable = " insert into taxledger ( empcode" & _
												"                         , year" & _
												"                         , month" & _
												"                         , processid" & _
												"                         , amount" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , remarks" & _
												"                         , lastupdate" & _
												"                         , lasttime)" & _
												" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
												"          , " & Year(Session("v_Employee")) & "" & _
												"          , " & Month(Session("v_Employee")) & "" & _
												"          , 'pph21_tm'" & _
												"          , " & temp_pph21_tm & "" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
												"          , '" & Session("fcBA") & "'" & _
												"          , 'Pph21 Sebulan'" & _
												"          , sysdate," & _
												"           to_char(sysdate,'HH24:MI'))"
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
							
							connstr.BeginTrans
							insertdatatotable = " insert into taxledger ( empcode" & _
												"                         , year" & _
												"                         , month" & _
												"                         , processid" & _
												"                         , amount" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , remarks" & _
												"                         , lastupdate" & _
												"                         , lasttime)" & _
												" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
												"          , " & Year(Session("v_Employee")) & "" & _
												"          , " & Month(Session("v_Employee")) & "" & _
												"          , 'pph21_ty'" & _
												"          , " & temp_pph21_ty & "" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Session("CodeName") & "'" & _
												"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
												"          , '" & Session("fcBA") & "'" & _
												"          , 'Pph21 Setahun'" & _
												"          , sysdate," & _
												"           to_char(sysdate,'HH24:MI'))"
										
							connstr.execute(insertdatatotable)
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
								Response.Flush
							End If
						end if
					else
						connstr.BeginTrans
						insertdatatotable = " insert into taxledger ( empcode" & _
											"                         , year" & _
											"                         , month" & _
											"                         , processid" & _
											"                         , amount" & _
											"                         , fcentry" & _
											"                         , fcedit" & _
											"                         , fcip" & _
											"                         , fcba" & _
											"                         , remarks" & _
											"                         , lastupdate" & _
											"                         , lasttime)" & _
											" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
											"          , " & Year(Session("v_Employee")) & "" & _
											"          , " & Month(Session("v_Employee")) & "" & _
											"          , 'pph21_tm'" & _
											"          , " & temp_pph21_tm & "" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
											"          , '" & Session("fcBA") & "'" & _
											"          , 'Pph21 Sebulan'" & _
											"          , sysdate," & _
											"           to_char(sysdate,'HH24:MI'))"
									
						connstr.execute(insertdatatotable)
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
							Response.Flush
						End If
						
						connstr.BeginTrans
						insertdatatotable = " insert into taxledger ( empcode" & _
											"                         , year" & _
											"                         , month" & _
											"                         , processid" & _
											"                         , amount" & _
											"                         , fcentry" & _
											"                         , fcedit" & _
											"                         , fcip" & _
											"                         , fcba" & _
											"                         , remarks" & _
											"                         , lastupdate" & _
											"                         , lasttime)" & _
											" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
											"          , " & Year(Session("v_Employee")) & "" & _
											"          , " & Month(Session("v_Employee")) & "" & _
											"          , 'pph21_ty'" & _
											"          , " & temp_pph21_ty & "" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Session("CodeName") & "'" & _
											"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
											"          , '" & Session("fcBA") & "'" & _
											"          , 'Pph21 Setahun'" & _
											"          , sysdate," & _
											"           to_char(sysdate,'HH24:MI'))"
									
						connstr.execute(insertdatatotable)
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
							Response.Flush
						End If
					end if
				else
					connstr.BeginTrans
					insertdatatotable = " insert into taxledger ( empcode" & _
										"                         , year" & _
										"                         , month" & _
										"                         , processid" & _
										"                         , amount" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , remarks" & _
										"                         , lastupdate" & _
										"                         , lasttime)" & _
										" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
										"          , " & Year(Session("v_Employee")) & "" & _
										"          , " & Month(Session("v_Employee")) & "" & _
										"          , 'pph21_tm'" & _
										"          , " & temp_pph21_tm & "" & _
										"          , '" & Session("CodeName") & "'" & _
										"          , '" & Session("CodeName") & "'" & _
										"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
										"          , '" & Session("fcBA") & "'" & _
										"          , 'Pph21 Sebulan'" & _
										"          , sysdate," & _
										"           to_char(sysdate,'HH24:MI'))"
								
					connstr.execute(insertdatatotable)
					
					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
						Response.Flush
					End If
					
					connstr.BeginTrans
					insertdatatotable = " insert into taxledger ( empcode" & _
										"                         , year" & _
										"                         , month" & _
										"                         , processid" & _
										"                         , amount" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , remarks" & _
										"                         , lastupdate" & _
										"                         , lasttime)" & _
										" values ( '" & array_empallded(2, array_empallded_record) & "'" & _
										"          , " & Year(Session("v_Employee")) & "" & _
										"          , " & Month(Session("v_Employee")) & "" & _
										"          , 'pph21_ty'" & _
										"          , " & temp_pph21_ty & "" & _
										"          , '" & Session("CodeName") & "'" & _
										"          , '" & Session("CodeName") & "'" & _
										"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
										"          , '" & Session("fcBA") & "'" & _
										"          , 'Pph21 Setahun'" & _
										"          , sysdate," & _
										"           to_char(sysdate,'HH24:MI'))"
								
					connstr.execute(insertdatatotable)
					
					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
						Response.Flush
					End If
					
				end if
				
				
			Next
		End if
		
		
		DIM arr2_has_data, arr2_eof, oQuery
		DIM amt_gajibruto_ty, amt_potastek_ty, amt_pph21_ty, rrr
		
		IF NOT array_empallded_eof THEN
			FOR array_empallded_record = LBound(array_empallded, 2) TO UBound(array_empallded, 2)
				arr2_eof = true
				GlobalSQL = " SELECT fcba, adhocdate, empcode, amount" & _
							" FROM empallded" & _
							" WHERE fcba = '" & Session("fcBA") & "'" & _
							"       AND TO_CHAR(ADHOCDATE, 'YYYYMM') = " & Year(Session("v_Employee")) & "||LPAD(" & Month(Session("v_Employee")) & ", 2, 0)" & _
							"       AND empcode = '" & array_empallded(2, array_empallded_record) & "'" & _
							"       AND allowdedcode = 'A61'"
							
				Set GlobalRS = Server.CreateObject("ADODB.Recordset")
				GlobalRS.CursorType = 3
				GlobalRS.CursorLocation = 3
				GlobalRS.Open GlobalSQL, connstr
				if not GlobalRS.eof then		
					arr2_has_data = GlobalRS.getrows()
					arr2_eof = false
				end if
				GlobalRS.Close
				set GlobalRS = Nothing
				
				IF NOT arr2_eof THEN
					FOR rrr = LBound(arr2_has_data, 2) TO UBound(arr2_has_data, 2)
						GlobalSQL = " SELECT NVL( taxledger_api.get_amount('" & Session("fcBA") & "', " & Year(Session("v_Employee")) & ", " & Month(Session("v_Employee")) & ", '" & array_empallded(2, array_empallded_record) & "', 'gajibruto_ty'), 0) amt_gajibruto_ty" & _
									" , NVL( taxledger_api.get_amount('" & Session("fcBA") & "', " & Year(Session("v_Employee")) & ", " & Month(Session("v_Employee")) & ", '" & array_empallded(2, array_empallded_record) & "', 'potastek_ty'), 0) amt_potastek_ty" & _
									" , NVL( taxledger_api.get_amount('" & Session("fcBA") & "', " & Year(Session("v_Employee")) & ", " & Month(Session("v_Employee")) & ", '" & array_empallded(2, array_empallded_record) & "', 'pph21_ty'), 0) amt_pph21_ty" & _
									" FROM DUAL"
									
						Set GlobalRS = Server.CreateObject("ADODB.Recordset")
						GlobalRS.CursorType = 3
						GlobalRS.CursorLocation = 3
						GlobalRS.Open GlobalSQL, connstr
						if not GlobalRS.eof then		
							amt_gajibruto_ty = Cdbl(GlobalRS("amt_gajibruto_ty"))
							amt_potastek_ty = Cdbl(GlobalRS("amt_potastek_ty"))
							amt_pph21_ty = Cdbl(GlobalRS("amt_pph21_ty"))
						Else
							amt_gajibruto_ty = 0
							amt_potastek_ty = 0
							amt_pph21_ty = 0
						end if
						GlobalRS.Close
						set GlobalRS = Nothing
						
						oQuery	= 	" BEGIN TAXLEDGER_API.PPH_THRBONUS( '" & Session("FCBA") & "'" & _
									"                                   , " & Year(Session("v_Employee")) & "" & _
									"                                   , " & Month(Session("v_Employee")) & "" & _
									"                                   , '" & array_empallded(2, array_empallded_record) & "'" & _
									"                                   , '" & array_empallded(3, array_empallded_record) & "'" & _
									"                                   , " & amt_gajibruto_ty & "" & _
									"                                   , " & amt_potastek_ty & "" & _
									"                                   , " & amt_pph21_ty & "" & _
									"                                   , 'A61'" & _
									"                                   , " & arr2_has_data(3, rrr) & "" & _
									" ); END; "
									
						'Response.Write oQuery & "<br/>"

						connStr.execute(oQuery)
					NEXT
				END IF
				
				arr2_eof = true
				GlobalSQL = " SELECT fcba, adhocdate, empcode, amount" & _
							" FROM empallded" & _
							" WHERE fcba = '" & Session("fcBA") & "'" & _
							"       AND TO_CHAR(ADHOCDATE, 'YYYYMM') = " & Year(Session("v_Employee")) & "||LPAD(" & Month(Session("v_Employee")) & ", 2, 0)" & _
							"       AND empcode = '" & array_empallded(2, array_empallded_record) & "'" & _
							"       AND allowdedcode = 'A59'"
							
				Set GlobalRS = Server.CreateObject("ADODB.Recordset")
				GlobalRS.CursorType = 3
				GlobalRS.CursorLocation = 3
				GlobalRS.Open GlobalSQL, connstr
				if not GlobalRS.eof then		
					arr2_has_data = GlobalRS.getrows()
					arr2_eof = false
				end if
				GlobalRS.Close
				set GlobalRS = Nothing
				
				IF NOT arr2_eof THEN
					FOR rrr = LBound(arr2_has_data, 2) TO UBound(arr2_has_data, 2)
						GlobalSQL = " SELECT NVL( taxledger_api.get_amount('" & Session("fcBA") & "', " & Year(Session("v_Employee")) & ", " & Month(Session("v_Employee")) & ", '" & array_empallded(2, array_empallded_record) & "', 'gajibruto_ty'), 0) amt_gajibruto_ty" & _
									" , NVL( taxledger_api.get_amount('" & Session("fcBA") & "', " & Year(Session("v_Employee")) & ", " & Month(Session("v_Employee")) & ", '" & array_empallded(2, array_empallded_record) & "', 'potastek_ty'), 0) amt_potastek_ty" & _
									" , NVL( taxledger_api.get_amount('" & Session("fcBA") & "', " & Year(Session("v_Employee")) & ", " & Month(Session("v_Employee")) & ", '" & array_empallded(2, array_empallded_record) & "', 'pph21_ty'), 0) amt_pph21_ty" & _
									" FROM DUAL"
									
						Set GlobalRS = Server.CreateObject("ADODB.Recordset")
						GlobalRS.CursorType = 3
						GlobalRS.CursorLocation = 3
						GlobalRS.Open GlobalSQL, connstr
						if not GlobalRS.eof then		
							amt_gajibruto_ty = Cdbl(GlobalRS("amt_gajibruto_ty"))
							amt_potastek_ty = Cdbl(GlobalRS("amt_potastek_ty"))
							amt_pph21_ty = Cdbl(GlobalRS("amt_pph21_ty"))
						Else
							amt_gajibruto_ty = 0
							amt_potastek_ty = 0
							amt_pph21_ty = 0
						end if
						GlobalRS.Close
						set GlobalRS = Nothing
						
						oQuery	= 	" BEGIN TAXLEDGER_API.PPH_THRBONUS( '" & Session("FCBA") & "'" & _
									"                                   , " & Year(Session("v_Employee")) & "" & _
									"                                   , " & Month(Session("v_Employee")) & "" & _
									"                                   , '" & array_empallded(2, array_empallded_record) & "'" & _
									"                                   , '" & array_empallded(3, array_empallded_record) & "'" & _
									"                                   , " & amt_gajibruto_ty & "" & _
									"                                   , " & amt_potastek_ty & "" & _
									"                                   , " & amt_pph21_ty & "" & _
									"                                   , 'A59'" & _
									"                                   , " & arr2_has_data(3, rrr) & "" & _
									" ); END; "
									
						'Response.Write oQuery & "<br/>"

						connStr.execute(oQuery)
					NEXT
				END IF
			
			NEXT
		END IF
		
		' To Pick max Line No in Cost_Activity
		qMaxLineNo = " select nvl(max(nvl(lineno, 0)), 0) maxLineNo " & _
						" from empallded " & _
						" where " & _
						" fcba = '" & Session("fcBA") & "' "
						
		Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
		rsqMaxLineNo.CursorType = 3
		rsqMaxLineNo.CursorLocation = 3
		rsqMaxLineNo.Open qMaxLineNo, connstr

		if not rsqMaxLineNo.eof then
			maxLineNo = cdbl(rsqMaxLineNo("maxLineNo"))
		else
			maxLineNo = 0
		end if
		rsqMaxLineNo.close
		set rsqMaxLineNo = Nothing	
		
		GlobalSQL = " SELECT empcode" & _
		            "        , allowdedcode" & _
					"        , adhocdate" & _
					"        , NVL(SUM(amount), 0) amount" & _
					" FROM (" & _					
		            "     SELECT empcode" & _
					"            , 'D03' allowdedcode" & _
					"            , last_day(to_date(month||'/01/'||year, 'mm/dd/yyyy')) adhocdate" & _
					"            , round(NVL(amount, 0), 0) * -1 amount" & _
					"     FROM taxledger" & _
					"     WHERE fcba = '" & Session("fcBA") & "'" & _
					"          AND year = " & Year(Session("v_Employee")) & "" & _
					"           AND month = " & Month(Session("v_Employee")) & "" & _
					"           AND processid = 'pph21_tm'" & _
					"     UNION ALL" & _
					"     SELECT empcode" & _
					"            , 'D03' allowdedcode" & _
					"            , last_day(to_date(month||'/01/'||year, 'mm/dd/yyyy')) adhocdate" & _
					"            , round(NVL(amount, 0), 0) * -1 amount" & _
					"     FROM taxledger_bonus" & _
					"     WHERE fcba = '" & Session("fcBA") & "'" & _
					"           AND year = " & Year(Session("v_Employee")) & "" & _
					"           AND month = " & Month(Session("v_Employee")) & "" & _
					"           AND processid = 'pph21_ty'" & _
					"     UNION ALL" & _
					"     SELECT empcode" & _
					"            , 'D03' allowdedcode" & _
					"            , last_day(to_date(month||'/01/'||year, 'mm/dd/yyyy')) adhocdate" & _
					"            , round(NVL(amount, 0), 0) * -1 amount" & _
					"     FROM taxledger_thr" & _
					"     WHERE fcba = '" & Session("fcBA") & "'" & _
					"           AND year = " & Year(Session("v_Employee")) & "" & _
					"           AND month = " & Month(Session("v_Employee")) & "" & _
					"           AND processid = 'pph21_ty'" & _
					" ) GROUP BY empcode" & _
		            "          , allowdedcode" & _
					"          , adhocdate" & _
					" UNION ALL" & _
					" SELECT empcode" & _
		            "        , allowdedcode" & _
					"        , adhocdate" & _
					"        , NVL(SUM(amount), 0) amount" & _
					" FROM (" & _					
		            "     SELECT empcode" & _
					"            , 'A58' allowdedcode" & _
					"            , last_day(to_date(month||'/01/'||year, 'mm/dd/yyyy')) adhocdate" & _
					"            , round(NVL(amount, 0), 0) amount" & _
					"     FROM taxledger" & _
					"     WHERE fcba = '" & Session("fcBA") & "'" & _
					"          AND year = " & Year(Session("v_Employee")) & "" & _
					"           AND month = " & Month(Session("v_Employee")) & "" & _
					"           AND processid = 'tunjpph21_tm'" & _
					"     UNION ALL" & _
					"     SELECT empcode" & _
					"            , 'A58' allowdedcode" & _
					"            , last_day(to_date(month||'/01/'||year, 'mm/dd/yyyy')) adhocdate" & _
					"            , round(NVL(amount, 0), 0) amount" & _
					"     FROM taxledger_bonus" & _
					"     WHERE fcba = '" & Session("fcBA") & "'" & _
					"           AND year = " & Year(Session("v_Employee")) & "" & _
					"           AND month = " & Month(Session("v_Employee")) & "" & _
					"           AND processid = 'finaltunjpph21_ty'" & _
					"     UNION ALL" & _
					"     SELECT empcode" & _
					"            , 'A58' allowdedcode" & _
					"            , last_day(to_date(month||'/01/'||year, 'mm/dd/yyyy')) adhocdate" & _
					"            , round(NVL(amount, 0), 0) amount" & _
					"     FROM taxledger_thr" & _
					"     WHERE fcba = '" & Session("fcBA") & "'" & _
					"           AND year = " & Year(Session("v_Employee")) & "" & _
					"           AND month = " & Month(Session("v_Employee")) & "" & _
					"           AND processid = 'finaltunjpph21_ty'" & _
					" ) GROUP BY empcode" & _
		            "          , allowdedcode" & _
					"          , adhocdate"

			'response.Write GlobalSQL
			'response.Flush()			    
			Set GlobalRS = Server.CreateObject("ADODB.Recordset")
			GlobalRS.CursorType = 3
			GlobalRS.CursorLocation = 3
			GlobalRS.Open GlobalSQL, connstr
			if not GlobalRS.eof then		
				array_empallded = GlobalRS.getrows()
				array_empallded_eof = false
			end if
			GlobalRS.Close
			set GlobalRS = Nothing		
		
		response.Write "Insert empallded ..."& "<br>" 
		response.Flush()
		
			if not array_empallded_eof then
				For array_empallded_record = LBound(array_empallded, 2) To UBound(array_empallded, 2)
					maxLineNo = maxLineNo + 1
					connstr.BeginTrans
					insertdatatotable = " insert into empallded ( lineno" & _
                                        "                         , empcode" & _
                                        "                         , allowdedcode" & _
                                        "                         , adhocdate" & _
                                        "                         , amount" & _
                                        "                         , fcentry" & _
                                        "                         , fcedit" & _
                                        "                         , fcip" & _
                                        "                         , fcba" & _
                                        "                         , lastupdate" & _
                                        "                         , lasttime)" & _
                                        " values ( " & maxLineNo & "" & _
                                        "          , '" & array_empallded(0, array_empallded_record) & "'" & _
                                        "          , '" & array_empallded(1, array_empallded_record) & "'" & _
                                        "          , to_date('" & array_empallded(2, array_empallded_record) & "', 'mm/dd/yyyy')" & _
                                        "          , '" & Cdbl( array_empallded(3, array_empallded_record) ) & "'" & _
                                        "          , '" & Session("CodeName") & "'" & _
                                        "          , '" & Session("CodeName") & "'" & _
                                        "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
                                        "          , '" & Session("fcBA") & "'" & _
                                        "          , sysdate" & _
                                        "          , to_char(sysdate,'HH24:MI'))"
'response.write(insertdatatotable)							
					connstr.execute(insertdatatotable)
					If Err.Number = 0 Then
						connstr.CommitTrans
						Response.Flush
					Else
						connstr.RollbackTrans
						Response.write("Error Save ["&insertdatatotable&"] ("&Err.Description&")")
						Response.Flush
					End If
				Next
			End if
		
	end if
%>