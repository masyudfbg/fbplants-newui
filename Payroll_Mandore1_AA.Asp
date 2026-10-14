<%
	'Amandemen Premi Supervisi Panen Mandore1, Mandore, Kerani Panen, Kerani Afd
	'1. Code Allowance Mandore 1 = A39 Rate 1 nya = Rate Mandore 1
	
	
    ' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Mandore1_AA.asp"
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
	
	
	
	
	
    ' Cari Rate Mandore 1
	qBA = " SELECT includedpayrollprocessing, man_auto, NVL(rate_1,0) rate_1, NVL(rate_2,0) rate_2, NVL(rate_3,0) rate_3 FROM allowded WHERE fccode = 'A39' AND fcba = '" & Session("fcBA") & "'"

	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   p_includedpayrollprocessing = rsBA("includedpayrollprocessing")
	   p_man_auto = rsBA("man_auto")
	   Rate_Mandore_1=cdbl(to_nol(rsBA("Rate_1")))
	   Rate_Mandore_2=cdbl(to_nol(rsBA("Rate_2")))
	   Rate_Mandore_3=cdbl(to_nol(rsBA("Rate_3")))
	else
	   p_includedpayrollprocessing = "FALSE"
	   p_man_auto = "X"
	   Rate_Mandore_1 = 0
	   Rate_Mandore_2 = 0
	   Rate_Mandore_3 = 0
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
	
	if p_includedpayrollprocessing = "TRUE" and p_man_auto = "A" then

		qSupervisi=" select gad.FDDATE,gad.SUPERVISION_1,nvl(sum(gad.mandays),0) as HK" & _
				 " from ATTENDANCE_GAD gad where  gad.jobcode IN ('112030101', '505030101')" & _
				 " and supervision_1 is not null" & _
				 " and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
				 " and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
				 " and nvl(gad.BUCKET,0) > 0  " & _
				 " and gad.ATTENDANCE IN ('KJ','WH','WS','KK')" & _
				 " group by gad.FDDATE,gad.SUPERVISION_1"

		Set rsSupervisi = Server.CreateObject("ADODB.Recordset")
		rsSupervisi.CursorType = 3
		rsSupervisi.CursorLocation = 3
		response.Write "Calculate Premi Mandore 1..."  & "<br>" 
		response.Flush()
		rsSupervisi.Open qSupervisi, connstr
		if not rsSupervisi.eof then
		   Do While Not rsSupervisi.eof
		   
				' Add by NAF
				qTotalAnggota = " select count(*) as TotalAnggota from " & _ 
					   " ( select distinct supervision_2 " & _
					   " from ATTENDANCE_GAD gad where  gad.jobcode IN ('112030101', '505030101')" & _
					   " and gad.FDDATE=to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') "  &  _
					   " and gad.SUPERVISION_1 = '" & rsSupervisi("SUPERVISION_1") & "'" & _
					   " and gad.SUPERVISION_2 IS NOT NULL " & _
					   " and gad.ATTENDANCE IN ('KJ','WH','WS','KK')" & _
					   " and gad.fcBA='" & Session("fcBA") & "' ) " 
					   
				
				Set rsTotalAnggota = Server.CreateObject("ADODB.Recordset")
				rsTotalAnggota.CursorType = 3
				rsTotalAnggota.CursorLocation = 3
				'response.Write "Supervisi ........" & rsSupervisi("SUPERVISION_1")  & "<br>"
				response.Flush()
				rsTotalAnggota.Open qTotalAnggota, connstr
				response.Flush()

				if not rsTotalAnggota.eof then
				  nTotalAnggota = cdbl(rsTotalAnggota ("TotalAnggota"))
				  rsTotalAnggota.Close
				else
				  nTotalAnggota = 0
				  rsTotalAnggota.Close			  
				end if
				set rsTotalAnggota = Nothing
			  
				
			  
				qSumTotalAmount = " select nvl(sum(nvl(amount,0)),0) as nSumTotalAmount from empallded	" & _
								" where " & _
								" fcba = '" & Session("fcBA")  & "' and " & _ 
								" allowdedcode ='A37' and " & _
								" adhocdate = to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') and " & _
								" empcode in " & _
							" ( select distinct supervision_2 " & _
								   " from ATTENDANCE_GAD gad where  gad.jobcode IN ('112030101', '505030101')" & _
								   " and gad.FDDATE=to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') "  &  _
								   " and gad.SUPERVISION_1 = '" & rsSupervisi("SUPERVISION_1") & "'" & _
								   " and gad.SUPERVISION_2 IS NOT NULL " & _
								   " and gad.fcBA='" & Session("fcBA") & "' ) " 
			  
				Set rsSumTotalAmount = Server.CreateObject("ADODB.Recordset")
					'response.Write "rsAnggota " & rsAnggota("EMPLOYEECODE")  &  "<br>"
				response.Flush()
				rsSumTotalAmount.CursorType = 3
				rsSumTotalAmount.CursorLocation = 3
				rsSumTotalAmount.Open qSumTotalAmount, connstr
				if not rsSumTotalAmount.eof then
					nSumTotalAmount=cdbl(rsSumTotalAmount("nSumTotalAmount"))
					response.Flush()
				else
					nSumTotalAmount = 0
				end if
				rsSumTotalAmount.Close
				set rsSumTotalAmount = Nothing
					
			  
			  
				if nTotalAnggota >= 3 then 
					nPremiMandore1 = (nSumTotalAmount/nTotalAnggota) * Rate_Mandore_1 / 100 
				else
					if nTotalAnggota = 2 then 
						nPremiMandore1 = (nSumTotalAmount/nTotalAnggota) * Rate_Mandore_1 / 100 
					else
						if nTotalAnggota = 1 then 
							nPremiMandore1 = (nSumTotalAmount/nTotalAnggota) * Rate_Mandore_1 / 100 			  
						else
							nPremiMandore1 = 0
						end if
					end if
				end if
			  
				if nPremiMandore1 > 0 then
					 'response.Write " Total Untuk Supervisi......" & nPremiMandore1 & "<br>"
					 maxLineNo = maxLineNo + 1
					 qAllowance="Insert Into  EmpAllded ( " & _
									  " LINENO, " & _
									  " EMPCODE, " & _
									  " ALLOWDEDCODE, " & _
									  " ADHOCDATE, " & _
									  " AMOUNT, " & _
									  " FCENTRY, " & _
									  " FCEDIT, " & _
									  " FCIP, " & _
									  " FCBA, " & _
									  " LASTUPDATE," & _ 
									  " LASTTIME) " & _
									  " values  " & _
								"("  & maxLineNo &  ",'" & _
									  rsSupervisi("SUPERVISION_1") &  "','" & _
									  "A39" & "'," & _
									  " to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') " & "," & _
									  nPremiMandore1 & ",'" & _
									  Session("CodeName") & "','" & _
									  Session("CodeName") & "','" & _
									  Request.ServerVariables("REMOTE_ADDR") &"','" & _
									  Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
									  
					 connstr.execute(qAllowance)
					 response.Flush()
					 
				end if
				nSumTotalAmount=0
				rsSupervisi.MoveNext
		   loop
		End if
		rsSupervisi.Close
		set rsSupervisi = Nothing
	end if
%>