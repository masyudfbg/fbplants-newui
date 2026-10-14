<%
	'Amandemen Premi Supervisi Panen Mandore, Mandore, Kerani Panen, Kerani Afd
	'                  Mandore   = A38 Rate 1 nya = Rate Kerani Afdeling	
    ' Cari Rate Kerani Buah
	
	
    ' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Kerani_Buah.asp"
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
	
	
	
	qBA = " SELECT includedpayrollprocessing, man_auto, NVL(rate_1,0) rate_1 FROM allowded WHERE fccode = 'A38' AND fcba = '" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
		p_includedpayrollprocessing = rsBA("includedpayrollprocessing")
	    p_man_auto = rsBA("man_auto")
	    Rate_Kerani_op = cdbl(to_nol(rsBA("rate_1")))
	else
		p_includedpayrollprocessing = "FALSE"
	    p_man_auto = "X"
	    Rate_Kerani_op = 0
	end if
	
	if p_includedpayrollprocessing = "TRUE" and p_man_auto = "A" then
		'qSupervisi=	 " select gad.FDDATE,gad.GANGCODE,gad.SUPERVISION_3,nvl(sum(gad.mandays),0) as HK" & _
		'			 " from ATTENDANCE_GAD gad where  gad.jobcode IN ('1323001', '6502101')" & _
		'			 " and supervision_3 is not null" & _
		'			 " and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
		'			 " and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
		'			 " group by gad.FDDATE,gad.GANGCODE,gad.SUPERVISION_3"
		
	    qSupervisi=" select " & _  				 
				 " hdb.FDDATE,hdb.GANGCODE,hdb.SUPERVISION_3, " & _ 
				 " case " & _
				 " when MAX(pm.param_1) = 'D' then " & _
				 "     (nvl(sum(hdb.QUANTITY_BYDAILY),0) * MAX(allowded.RATE_1)/100)" & _
				 " else " & _
				 "     (nvl(sum(hdb.QUANTITY_BYMONTHLY),0) * MAX(allowded.RATE_1)/100)" & _
				 " end UNDER_RIPE_AMOUNT" & _
				 " from HARVESTING_DAILY_BJR hdb " & _
				 " INNER JOIN allowded ON allowded.fccode='A38' and allowded.fcba=hdb.fcBA" & _
				 " inner join iplasprod.parameter pm on (hdb.fcba = pm.fcba) " & _
				 " where " & _
				 " hdb.SUPERVISION_3 is not null " & _
				 " and to_number(to_Char(TO_DATE(hdb.FDDATE,'MM/DD/YYYY') ,'mm'))=  " & Month(Session("v_Employee")) &  _
				 " and to_number(to_Char(TO_DATE(hdb.FDDATE,'MM/DD/YYYY') ,'yyyy'))=  " & Year(Session("v_Employee")) & _
				 " and hdb.fcBA='" & Session("fcBA") & "'" & _
				 " and pm.parcode = 'HARDCODE'" & _
				 " and pm.fccode = 'JOBCODE_PANEN' " & _
				 " group by hdb.FDDATE, hdb.GANGCODE, hdb.SUPERVISION_3"
				 
		Set rsSupervisi = Server.CreateObject("ADODB.Recordset")
		rsSupervisi.CursorType = 3
		rsSupervisi.CursorLocation = 3
		response.Write "Calculate Premi Kerani Buah ..."  & "<br>" 
		response.Flush()
		rsSupervisi.Open qSupervisi, connstr
		if not rsSupervisi.eof then
		   Do While Not rsSupervisi.eof
			' nPremiSortasi=0
			' qAnggota=" select gad.FDDATE,gad.GANGCODE,gad.EMPLOYEECODE,gad.SUPERVISION_3,nvl(sum(gad.mandays),0) as HK" & _
			'		   " from ATTENDANCE_GAD gad where  gad.jobcode IN ('1323001', '6502101')" & _
			'		   " and gad.FDDATE=to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') "  &  _
			'		   " and gad.GANGCODE='" & rsSupervisi("GANGCODE") & "'" & _
			'		   " and gad.fcBA='" & Session("fcBA") & "'" & _
			'		   " group by gad.FDDATE,gad.GANGCODE,gad.EMPLOYEECODE,gad.SUPERVISION_3"
			'		   
			'  
			'  Set rsAnggota = Server.CreateObject("ADODB.Recordset")
			'  rsAnggota.CursorType = 3
			'  rsAnggota.CursorLocation = 3
			'  'response.Write "Supervisi ........" & rsSupervisi("SUPERVISION_2")  & "<br>"
			'  response.Flush()
			'  rsAnggota.Open qAnggota, connstr
			'  response.Flush()
			' if not rsAnggota.eof then
			'	 Do While Not rsAnggota.eof
			'		qEmpAllded=" Select nvl(sum(decode(substr(allowdedcode,1,1),'A',amount,amount*-1)),0) as PremiSortasi " & _		                   
			'				   " from EmpAllded where empcode='" & rsAnggota("EMPLOYEECODE") & "'" & _
			'				   " and adHocDate=to_date('" & rsAnggota("FDDATE") & "','mm/dd/yyyy') "  &  _  
			'				   " and AllowdedCode in ('A24','A27','D09','D10') " & _
			'				   " and fcBA='" & Session("fcBA") & "'"
			'		Set rsEmpAllded = Server.CreateObject("ADODB.Recordset")
			'		'response.Write "rsAnggota " & rsAnggota("EMPLOYEECODE")  &  "<br>"
			'		response.Flush()
			'		rsEmpAllded.CursorType = 3
			'		rsEmpAllded.CursorLocation = 3
			'		rsEmpAllded.Open qEmpAllded, connstr
			'		if not rsEmpAllded.eof then
			'			nPremiSortasi=nPremiSortasi+cdbl(rsEmpAllded("PremiSortasi"))
			'			'response.Write  "Payslip : " & rsEmpAllded("PremiSortasi")  & "<br>"
			'			response.Flush()
			'		end if
			'		rsEmpAllded.Close
			'		set rsEmpAllded = Nothing
			'		
			'		rsAnggota.MoveNext
			'	 Loop
			'  end if
			'  rsAnggota.Close
			'  set rsAnggota = Nothing
			'  if cdbl(rsSupervisi("HK")) = 0 then
			'	 nPremiKeraniBuah_op = 0
			'  else
			'	  if cdbl(rsSupervisi("HK")) < 15 then
			'		 nPremiKeraniBuah_op= ((nPremiSortasi) / 15 ) * (Rate_Kerani_op/100)
			'	  else
			'		 nPremiKeraniBuah_op= ((nPremiSortasi) / cdbl(rsSupervisi("HK")) ) * (Rate_Kerani_op/100)
			'	  end if
			'  'end if
			  
			  nPremiKeraniBuah_op = round(cdbl(rsSupervisi("UNDER_RIPE_AMOUNT")), 2)
			  if nPremiKeraniBuah_op > 0 then
					maxLineNo = maxLineNo + 1
					qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
							"(" & maxLineNo & ",'" & rsSupervisi("SUPERVISION_3") & "','" & "A38" & "'," &  " to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') " & "," & nPremiKeraniBuah_op & _
							",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
					'response.Write " Total Untuk Supervisi......" & qAllowance & "<br>"
					'response.Flush()
					connstr.execute(qAllowance)
					response.Flush()
					
					' Remarks History By Iwa : 08/01/2019
					
					'qSupervisiCount = " select count(*) total " & _
					'				" from ATTENDANCE_GAD gad where  gad.jobcode IN ('112030101', '505030101')" & _
					'									" and supervision_3 is not null" & _
					'									" and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					'									" and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
					'									" and gad.FDDATE = to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') " & _
					'									" and gad.gangcode = '" & rsSupervisi("GANGCODE") & "' " & _
					'									" and gad.supervision_3 = '" & rsSupervisi("supervision_3") & "' "
														
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
					'							" gad.SUPERVISION_3, " & _
					'							" nvl(gad.mandays,0) as HK" & _
					'					" from ATTENDANCE_GAD gad where  gad.jobcode IN ('112030101', '505030101')" & _
					'					" and supervision_3 is not null" & _
					'					" and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					'					" and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
					'					" and gad.FDDATE = to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') " & _
					'					" and gad.gangcode = '" & rsSupervisi("GANGCODE") & "' " & _
					'					" and gad.SUPERVISION_3 = '" & rsSupervisi("SUPERVISION_3") & "' " 
										
					'if tempTotal = 0 then nAmountDetail = 0 else nAmountDetail = round(nPremiKeraniBuah_op / tempTotal, 2)
					
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
					'			nAmountDetail = nPremiKeraniBuah_op - totalAmountDetail 
					'		else
					'			totalAmountDetail = totalAmountDetail + nAmountDetail
					'		end if
					'		
					'		' Insert the History of EmpallDed Data 
					'		' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
							
					'		txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
					'									 rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
					'									 Session("fcBA")
					'									 
					'		nRateDetail = 0
					'		nUnitDetail = 0
					'		
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
					'end if
					'rsReceive.Close
					'set rsReceive = Nothing										
				 
			  end if
			  nPremiSortasi=0
			  rsSupervisi.MoveNext
		   loop
		End if
		rsSupervisi.Close
		set rsSupervisi = Nothing
	end if
%>