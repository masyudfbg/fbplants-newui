<%
	' Amandemen Premi Tandan Hari libur Mandor1
	' Kode Allowded = 'A56'
	
    ' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.
	
	KAPTHLM   = Cstr("A56")    'KAPTHLM  = Kode Allowded Premi Tandan Hari Libur Mandor1
	KAPTHL    = Cstr("A54")    'KAPTHL   = Kode Allowded Premi Tandan Hari Libur Anak Buah Dalam hal ini mandor panen
	KAPPTHLM  = Cstr("D23")    'KAPPTHLM = Kode Allowded Potongan Premi Tandan Hari Libur Mandor1
	
	nRateBagi = 0
	nFaktorKali = 0
    nPremiTandanHariLibur = 0
	
	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_TandanHariLiburMandor1.asp"
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
	
	
	' Kode Allowded Premi Tandan Hari Libur Mandor1, By NAF
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	' qParameter = " select nvl(rate_1, 0) as rate_1" & _
	             ' "        , nvl(rate_2, 0) as rate_2,type_2" & _
				 ' "        , includedpayrollprocessing" & _
				 ' " from allowded" & _
				 ' " where fccode = '" & KAPTHLM & "'" & _
				 ' "       and man_auto = 'A'" & _
				 ' "       and fcba = '" & Session("fcBA") & "'"
				 
	qParameter = " select nvl(rate_1, 0) as rate_1" & _
	             "        , nvl(rate_2, 0) as rate_2" & _
				 "        , nvl(rate_3, 0) as rate_3" & _
	             "        , type_2" & _
				 "        , nvl(includedpayrollprocessing, 'FALSE') includedpayrollprocessing" & _
				 " from allowded" & _
				 " where fccode = '" & KAPTHLM & "'" & _
				 "       and man_auto = 'A'" & _
				 "       and fcba = '" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		' nRate 
		' nRateBagi = cdbl(rsParameter("Rate_1"))
		' nFaktorKali = Cdbl(rsParameter("rate_2"))
		' tempIncludePayrollProcessingBasis = rsParameter("includedpayrollprocessing")
		' nPaymentType = rsParameter("type_2")
		
		nFaktorKali_1 = cdbl(rsParameter("Rate_1"))
		nFaktorKali_2 = Cdbl(rsParameter("rate_2"))
		nFaktorKali_3 = Cdbl(rsParameter("rate_3"))
		tempIncludePayrollProcessingBasis = rsParameter("includedpayrollprocessing")
		nPaymentType = rsParameter("type_2")
		
	else
		' nPaymentType = "NK"
		' nRateBagi  = 0
		' nFaktorKali = 0
		' tempIncludePayrollProcessingBasis = "FALSE"
		
		nFaktorKali_1 = 0
		nFaktorKali_2 = 0
		nFaktorKali_3 = 0
		tempIncludePayrollProcessingBasis = "FALSE"
		nPaymentType = "NK"
	End if
	rsParameter.Close
	set rsParameter = Nothing
	
	
	' Kode Allowded Potongan Premi Tandan Hari Libur Mandor, By NAF
	Set rsParameter = Server.CreateObject("ADODB.Recordset")
	qParameter = " select includedpayrollprocessing" & _
				 " from allowded" & _
				 " where fccode = '" & KAPPTHLM & "'" & _
				 "       and fcba = '" & Session("fcBA") & "'"
	'response.Write " qParameter = " & qParameter				  
	rsParameter.CursorType = 3
	rsParameter.CursorLocation = 3
	rsParameter.Open qParameter, connstr
	if not rsParameter.eof then
		tempIncludePayrollProcessingBasisKontanan = rsParameter("includedpayrollprocessing")
	else
		tempIncludePayrollProcessingBasisKontanan = "FALSE"		 
	End if
	rsParameter.Close
	set rsParameter = Nothing
	
	
	
	
	If tempIncludePayrollProcessingBasis = "TRUE" Then

		' qSupervisi =	" select gad.fddate" & _
						' "        , gad.gangcode" & _
						' "        , gad.supervision_1" & _
						' "        , nvl(sum(gad.mandays),0) as hk" & _
						' " from attendance_gad gad" & _
						' " where process_posting_api.GetPostingValue ('PAYROLL', '" & KAPTHLM & "', 'JOB', gad.jobcode, '" & Session("fcBA") & "') = 'TRUE'" & _
						' "       and supervision_1 is not null" & _
						' "       and to_number(to_char(gad.fddate,'mm')) = " & Month(Session("v_Employee")) & "" & _
						' "       and to_number(to_char(gad.fddate,'yyyy'))= " & Year(Session("v_Employee")) & "" & _
						' "       and gad.fcba = '" & Session("fcBA") & "'" & _
						' " group by gad.fddate" & _
						' "          , gad.gangcode" & _
						' "          , gad.supervision_1"
						
		qSupervisi =	" select fddate" & _
						"        , supervision_1" & _
						"        , count(*) hk" & _
						" from (" & _
						"     select gad.fddate" & _
						"            , gad.supervision_1" & _
						"            , supervision_2" & _
						"     from attendance_gad gad" & _
						"     where process_posting_api.GetPostingValue ('PAYROLL', '" & KAPTHLM & "', 'JOB', gad.jobcode, '" & Session("fcBA") & "') = 'TRUE'" & _
						"           and supervision_1 is not null" & _
						"           and to_number(to_char(gad.fddate,'mm')) = " & Month(Session("v_Employee")) & "" & _
						"           and to_number(to_char(gad.fddate,'yyyy'))= " & Year(Session("v_Employee")) & "" & _
						"           and gad.fcba = '" & Session("fcBA") & "'" & _
						"           and gad.attendance = 'KK'" & _
						"     group by gad.fddate" & _
						"              , gad.supervision_1" & _
						"              , supervision_2" & _
						"     )" & _
						" group by fddate" & _
						"          , supervision_1"

	   
		Set rsSupervisi = Server.CreateObject("ADODB.Recordset")
		rsSupervisi.CursorType = 3
		rsSupervisi.CursorLocation = 3
		response.Write "Calculate Premi Mandor1Tandan Hari Libur ..."  & "<br>" 
		response.Flush()
		rsSupervisi.Open qSupervisi, connstr
		if not rsSupervisi.eof then
		   Do While Not rsSupervisi.eof
			  response.Write "."
			  response.Flush()
			 			  
			  ' qAnggota	=	" select gad.fddate" & _
			                ' "        , gad.gangcode" & _
			                ' "        , gad.employeecode" & _
			                ' "        , gad.supervision_1" & _
			                ' "        , nvl(sum(gad.mandays),0) as hk" & _
							' " from attendance_gad gad" & _
			                ' " where process_posting_api.GetPostingValue ('PAYROLL', '" & KAPTHLM & "', 'JOB', gad.jobcode, '" & Session("fcBA") & "') = 'TRUE'" & _
							' "       and to_char(gad.fddate,'YYYYMMDD') = TO_CHAR(to_date('" & rsSupervisi("fddate") & "','mm/dd/yyyy'),'YYYYMMDD') "  &  _
							' "       and gad.gangcode = '" & rsSupervisi("gangcode") & "'" & _
							' "       and gad.fcba = '" & Session("fcBA") & "'" & _
							' "       and gad.supervision_1 = '" &  rsSupervisi("supervision_1")   & "'" & _
							' " group by gad.fddate" & _
			                ' "          , gad.gangcode" & _
			                ' "          , gad.employeecode" & _
			                ' "          , gad.supervision_1"
							
			  qAnggota	=	" select gad.fddate" & _
			                "        , gad.supervision_1" & _
			                "        , gad.supervision_2" & _
			                "        , nvl(sum(gad.mandays),0) as hk" & _
							" from attendance_gad gad" & _
			                " where process_posting_api.GetPostingValue ('PAYROLL', '" & KAPTHLM & "', 'JOB', gad.jobcode, '" & Session("fcBA") & "') = 'TRUE'" & _
							"       and TO_CHAR(gad.fddate,'YYYYMMDD') = TO_CHAR(to_date('" & rsSupervisi("fddate") & "','mm/dd/yyyy'),'YYYYMMDD') "  &  _
							"       and gad.fcba = '" & Session("fcBA") & "'" & _
							"       and gad.attendance = 'KK'" & _
							"       and gad.supervision_1 = '" &  rsSupervisi("supervision_1")   & "'" & _
							" group by gad.fddate" & _
			                "          , gad.supervision_1" & _
			                "          , gad.supervision_2"
							
							
			  Set rsAnggota = Server.CreateObject("ADODB.Recordset")
			  rsAnggota.CursorType = 3
			  rsAnggota.CursorLocation = 3
			  'response.Write "Supervisi ........" & rsSupervisi("supervision_1")  & "<br>"
			  rsAnggota.Open qAnggota, connstr
			  if not rsAnggota.eof then
				 Do While Not rsAnggota.eof
					qEmpAllded=" Select nvl(sum(nvl(amount,0)),0) as PremiTandanHariLibur " & _
							   " from empallded" & _
							   " where empcode = '" & rsAnggota("supervision_2") & "'" & _
							   "       and TO_CHAR(adhocdate,'YYYYMMDD') = TO_CHAR(to_date('" & rsAnggota("fddate") & "','mm/dd/yyyy'),'YYYYMMDD') "  &  _  
							   "       and allowdedcode = '" & KAPTHL & "' " & _
							   "       and fcba = '" & Session("fcBA") & "'"
					Set rsEmpAllded = Server.CreateObject("ADODB.Recordset")
					response.Flush()
					rsEmpAllded.CursorType = 3
					rsEmpAllded.CursorLocation = 3
					rsEmpAllded.Open qEmpAllded, connstr
					if not rsEmpAllded.eof then
					   nPremiTandanHariLibur = nPremiTandanHariLibur + cdbl(rsEmpAllded("PremiTandanHariLibur"))
					   response.Flush()
					end if
					rsEmpAllded.Close
					set rsEmpAllded = Nothing
					
					rsAnggota.MoveNext
				 Loop
			  end if
			  rsAnggota.Close
			  set rsAnggota = Nothing
			  
			   ' if cdbl(rsSupervisi("HK")) > nRateBagi then
			      ' nRateBagi = cdbl(rsSupervisi("HK"))
			   ' end if	
			  
			  
			  ' nPremiMandore= ((nPremiTandanHariLibur) / nRateBagi ) * nFaktorKali
			  
			  if cdbl(rsSupervisi("HK")) = 0 then
				nPremiMandore = 0
			  elseif cdbl(rsSupervisi("HK")) = 1 then
				nPremiMandore= ((nPremiTandanHariLibur) / cdbl(rsSupervisi("HK")) ) * nFaktorKali_1
			  elseif cdbl(rsSupervisi("HK")) = 2 then
				nPremiMandore= ((nPremiTandanHariLibur) / cdbl(rsSupervisi("HK")) ) * nFaktorKali_2
			  elseif cdbl(rsSupervisi("HK")) >= 3 then
				nPremiMandore= ((nPremiTandanHariLibur) / cdbl(rsSupervisi("HK")) ) * nFaktorKali_3
			  else
				nPremiMandore = 0
			  end if
			  
			  nPremiMandore = round(nPremiMandore, 0)
			  if nPremiMandore > 0 then
					maxLineNo = maxLineNo + 1
					qAllowance = " insert into empallded ( empcode" & _
					             "                            , allowdedcode" & _
					             "                            , adhocdate" & _
					             "                            , amount" & _
					             "                            , fixedallded" & _
					             "                            , fcentry" & _
					             "                            , fcedit" & _
					             "                            , fcip" & _
					             "                            , fcba" & _
					             "                            , remarks" & _
					             "                            , lastupdate" & _
					             "                            , lasttime" & _
					             "                            , lineno" & _
					             "                        ) values (" & _
					             " '" & rsSupervisi("supervision_1") & "'" & _
					             " , '" & KAPTHLM & "'" & _
					             " , to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy')" & _
					             " , " & nPremiMandore & "" & _
								 " , 0 "& _
								 " , '" & Session("CodeName") & "'" & _
					             " , '" & Session("CodeName") & "'" & _
					             " , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					             " , '" & Session("fcBA") & "'" & _
								 " , null" & _
					             " , sysdate" & _
					             " , to_char(sysdate,'hh24:mi')" & _
								 " , " & maxLineNo & "" & _
								 " )"

					connstr.execute(qAllowance)
					response.Flush()
					
					' Remarks History By Iwa : 08/01/2019
					'qSupervisiCount = " select count(*) total " & _
					'				" from ATTENDANCE_GAD gad where process_posting_api.GetPostingValue ('PAYROLL', '" & KAPTHLM & "', 'JOB', gad.jobcode, '" & Session("fcBA") & "') = 'TRUE'" & _
					'									" and supervision_1 is not null" & _
					'									" and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					'									" and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
					'									" and gad.FDDATE = to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') " & _
					'									" and gad.attendance = 'KK'" & _
					'									" and gad.supervision_1 = '" & rsSupervisi("supervision_1") & "' "
														
					'Set rsReceive = Server.CreateObject("ADODB.Recordset")
					'rsReceive.CursorType = 3
					'rsReceive.CursorLocation = 3	

					'rsReceive.Open qSupervisiCount, connstr
					'if not rsReceive.eof then
					'	tempTotal = cdbl(rsReceive("total"))
					'else
					'	tempTotal = 1
					'end if
					'rsReceive.Close
					'set rsReceive = Nothing
					
					'qSupervisiDetail = " select " & _
					'							" gad.FDDATE, " & _
					'							" gad.LocationType, " & _
					'							" gad.LocationCode, " & _
					'							" gad.EmployeeCode, " & _
					'							" gad.JobCode, " & _
					'							" gad.GANGCODE, " & _
					'							" gad.supervision_1, " & _
					'							" nvl(gad.mandays,0) as HK" & _
					'					" from ATTENDANCE_GAD gad where process_posting_api.GetPostingValue ('PAYROLL', '" & KAPTHLM & "', 'JOB', gad.jobcode, '" & Session("fcBA") & "') = 'TRUE'" & _
					'					" and supervision_1 is not null" & _
					'					" and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					'					" and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
					'					" and TO_CHAR(gad.FDDATE,'YYYYMMDD') = TO_CHAR(to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy'),'YYYYMMDD') " & _
					'					" and gad.attendance = 'KK'" & _
					'					" and gad.supervision_1 = '" & rsSupervisi("supervision_1") & "' "
										
					' Insert history Data.
						
					'nAmountDetail = round(nPremiMandore / tempTotal, 2)
					'
					'Set rsReceive = Server.CreateObject("ADODB.Recordset")
					'rsReceive.CursorType = 3
					'rsReceive.CursorLocation = 3

					'response.write qSupervisiDetail
					'rsReceive.Open qSupervisiDetail, connstr
					'if not rsReceive.eof then
					'	detailLineNo = 0
					'	tempCounter = 0
					'	totalAmountDetail = 0
					'	Do While Not rsReceive.eof	
					'		if tempCounter = tempTotal - 1 then 
					'			nAmountDetail = nPremiMandore - totalAmountDetail 
					'		else
					'			totalAmountDetail = totalAmountDetail + nAmountDetail
					'		end if
					'		
					'		'	Insert the History of EmpallDed Data 
					'		' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
					'		
					'		txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
					'									 Session("fcBA")
					'									 
					'		nRateDetail = 0
					'		nUnitDetail = 0
					
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
					'end if
					'rsReceive.Close
					'set rsReceive = Nothing	
				    ' End Remarks History By Iwa : 08/01/2019
					if nPaymentType = "KK" then 
					    if tempIncludePayrollProcessingBasisKontanan = "TRUE" Then
							maxLineNo = maxLineNo + 1
				            qAllowance = " insert into empallded (    empcode" & _
					             "                            , allowdedcode" & _
					             "                            , adhocdate" & _
					             "                            , amount" & _
					             "                            , fixedallded" & _
					             "                            , fcentry" & _
					             "                            , fcedit" & _
					             "                            , fcip" & _
					             "                            , fcba" & _
					             "                            , remarks" & _
					             "                            , lastupdate" & _
					             "                            , lasttime" & _
					             "                            , lineno" & _
					             "                        ) values (" & _
					             " '" & rsSupervisi("supervision_1") & "'" & _
					             " , '" & KAPPTHLM & "'" & _
					             " , to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy')" & _
					             " , " & nPremiMandore * - 1 & "" & _
								 " , 0 "& _
								 " , '" & Session("CodeName") & "'" & _
					             " , '" & Session("CodeName") & "'" & _
					             " , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					             " , '" & Session("fcBA") & "'" & _
								 " , null" & _
					             " , sysdate" & _
					             " , to_char(sysdate,'hh24:mi')" & _
								 " , " & maxLineNo & "" & _
								 " )"

					        connstr.execute(qAllowance)
							
							' Remarks History By Iwa : 08/01/2019
							
							'qSupervisiDetail = " select " & _
							'					" gad.FDDATE, " & _
							'					" gad.LocationType, " & _
							'					" gad.LocationCode, " & _
							'					" gad.EmployeeCode, " & _
							'					" gad.JobCode, " & _
							'					" gad.GANGCODE, " & _
							'					" gad.supervision_1, " & _
							'					" nvl(gad.mandays,0) as HK" & _
							'			" from ATTENDANCE_GAD gad where process_posting_api.GetPostingValue ('PAYROLL', '" & KAPPTHLM & "', 'JOB', gad.jobcode, '" & Session("fcBA") & "') = 'TRUE'" & _
							'			" and supervision_1 is not null" & _
							'			" and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
							'			" and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
							'			" and TO_CHAR(gad.FDDATE,'YYYYMMDD') = TO_CHAR(to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy'),'YYYYMMDD') " & _
							'			" and gad.attendance = 'KK'" & _
							'			" and gad.supervision_1 = '" & rsSupervisi("supervision_1") & "' "
							'	
							'	nAmountDetail = round(nPremiMandore / tempTotal, 2)
							'
							'	Set rsReceive = Server.CreateObject("ADODB.Recordset")
							'	rsReceive.CursorType = 3
							'	rsReceive.CursorLocation = 3

							'	'response.write qSupervisiDetail
							'	rsReceive.Open qSupervisiDetail, connstr
							'	if not rsReceive.eof then
							'		detailLineNo = 0
							'		tempCounter = 0
							'		totalAmountDetail = 0
							'		Do While Not rsReceive.eof	
							'			if tempCounter = tempTotal - 1 then 
							'				nAmountDetail = nPremiMandore - totalAmountDetail 
							'			else
							'				totalAmountDetail = totalAmountDetail + nAmountDetail
							'			end if
							'			
							'			'	Insert the History of EmpallDed Data 
							'			' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
							'			
							'			txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
							'										 rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
							'										 rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
							'										 rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
							'										 rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
							'										 Session("fcBA")
																	 
							'			nRateDetail = 0
							'			nUnitDetail = 0
										
																	 
										
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
							'												" UNIT, " & _
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
							'												nUnitDetail & "," & _
							'												nRateDetail & "," & _
							'												nAmountDetail * -1 & "," & _ 
							'												" '" & Session("CodeName") & "'," & _
							'												" '" & Session("CodeName") & "'," & _
							'												" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
							'												" sysdate, " & _
							'												" to_char(sysdate, 'HH24:MI') " & _
							'												" ) "					
							'			connstr.execute(qAllowanceHistory)
							'			rsReceive.MoveNext
							'			tempCounter = tempCounter + 1
							'		loop
							'	end if
							'	rsReceive.Close
							'	set rsReceive = Nothing								
				        end if
				    end if 
					
			    end if
			  nPremiTandanHariLibur=0
			  rsSupervisi.MoveNext
		   loop
		End if
		rsSupervisi.Close
		set rsSupervisi = Nothing		
	End If
%>