<%
	'Amandemen Jamsostek
	'1. Parameter Master Status Pekerja di Parameter = E0006
	'   Rate 1 = Harga Beras, Rate 2 = Kg Beras Pekerja, Rate 3 = Kg Beras Pasangan, Rate 4 = Kg Beras Tanggungan
	'2. Allowded Code Untuk Jamsostek = A17 Tunjangan D02 Potongan 
	'   A17 : Rate 1 Harga Beras Jamsostek
	'       : Rate2 z Kebun
	'       : Rate3 z Pabrik
	'   D02 : Rate2 y Kebun
	'       : Rate3 y Pabrik
	
	'3. Bedakan Pabrik dan Kebun dari Business Unit Kolom fcType E or M
	
	'''''''' NOTE : Yang di atas tidak di pakai lagi !!! ''''''''
	' 1. Harga Beras Jamsostek di ambil dari tabel ALLOWDED kolom RATE_1 dengan kondisi FCCODE = A17 (Tunjangan Jamsostek)
	' 2. Kg Beras Pekerja di ambil dari tabel PARAMETER kolom RATE_1 dengan kondisi parcode = E0006 dan FCCODE di sesuaikan dengan MARITAL STATUS di tabel MASTER EMPLOYEE
	' 3. Penunjang dan Pemotong (ex. 4.54/6.54, 4.89/6.89, 0.84/0.84) di ambil dari tabel PARAMETER kolom RATE_1, RATE_2, RATE_3, RATE_4 s/d RATE_10
	'                                                                               dengan kondisi PARCODE = E0007
	'                                                                               dan FCCODE disesuaikan dengan kategori jamsostek di tabel MASTER EMPLOYEE
	
	Function GetPercentAllowded(in_jamsostekcategory, in_fcba)
		Dim temp_totalpercent, sqltotalpercent, rstotalpercent
		sqltotalpercent = " select nvl( sum(rate_1 +" & _
		                  "                 rate_2 +" & _
						  "                 rate_3 +" & _
						  "                 rate_4 +" & _
						  " 				rate_5), 0) total" & _
                          " from parameter" & _
                          " where fccode = '"& in_jamsostekcategory &"'" & _
                          "       and parcode = 'E0007'" & _
                          "       and fcba = '" & in_fcba & "'"
						  
		'response.write "<br/> All Query: " & sqltotalpercent
		Set rstotalpercent = Server.CreateObject("ADODB.Recordset")
		rstotalpercent.Open sqltotalpercent, connstr, 3
		
		If Not rstotalpercent.EOF Then
			temp_totalpercent = Cdbl(rstotalpercent("total"))
		Else
			temp_totalpercent = 0
		End If
		
		sqltotalpercent =	""
		rstotalpercent.Close
		set rstotalpercent = Nothing

		GetPercentAllowded = temp_totalpercent
	End Function
	
	Function GetPercentDeduction(in_jamsostekcategory, in_fcba)
		Dim temp_totalpercent, sqltotalpercent, rstotalpercent
		sqltotalpercent = " select nvl( sum(rate_1 +" & _
		                  "                 rate_2 +" & _
						  "                 rate_3 +" & _
						  "                 rate_4 +" & _
						  "                 rate_5 +" & _
						  "                 rate_6 +" & _
						  "                 rate_7 +" & _
						  "                 rate_8 +" & _
						  "                 rate_9 +" & _
						  " 				rate_10), 0) total" & _
                          " from parameter" & _
                          " where fccode = '"& in_jamsostekcategory &"'" & _
                          "       and parcode = 'E0007'" & _
                          "       and fcba = '" & in_fcba & "'"
						  
		'response.write "<br/> Ded Query: " & sqltotalpercent
		Set rstotalpercent = Server.CreateObject("ADODB.Recordset")
		rstotalpercent.Open sqltotalpercent, connstr, 3
		
		If Not rstotalpercent.EOF Then
			temp_totalpercent = Cdbl(rstotalpercent("total"))
		Else
			temp_totalpercent = 0
		End If
		
		sqltotalpercent =	""
		rstotalpercent.Close
		set rstotalpercent = Nothing

		GetPercentDeduction = temp_totalpercent
	End Function
	
	Dim PercentAllowded, PercentDeduction

	Dim nKgBerasPekerja, nHargaBerasJamsostek, zKebun, zPabrik, zOthers, yKebun, yPabrik, yOthers, vAll_JamSostek, vDed_JamSostek
    ' Cari Kebun / Pabrik   Type E = Estate M = Mill
    Dim rsBA,qBA,vEstateorMill
	
	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Jamsostek.asp"
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

	
	
	
	
	qBA="select FCTYPE from BUSINESSUNIT where fcCode='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.CursorType = 3
	rsBA.CursorLocation = 3
	rsBA.Open qBA, connstr
	if not rsBA.eof then
	   vEstateorMill=rsBA("FCTYPE")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY9001 Business Area ( Type) is Empty"
	   err.number=9002
	   vEstateorMill=""
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

    ' Cari yKebun,yPabrik (Potongan)
	qBA = " select nvl(sum(rate_2), 0) as rate_2" & _
		  "        , nvl(sum(rate_3), 0) as rate_3" & _
		  "        , nvl(sum(rate_4), 0) as rate_4" & _
		  " from allowded" & _
		  " where fccode = 'D02'" & _
		  "       and fcba = '" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.CursorType = 3
	rsBA.CursorLocation = 3
	rsBA.Open qBA, connstr
	if not rsBA.eof then
	   yKebun = cdbl(rsBA("Rate_2"))
	   yPabrik = cdbl(rsBA("Rate_3"))
	   yOthers = Cdbl(rsBA("rate_4"))
	else
	   yKebun = 0
	   yPabrik = 0
	   yOthers = 0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY9002 Harga Beras AllowDed Code D02 is Empty"
	   err.number = 9002
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

    ' Cari Harga Beras Jamsostek, zKebun, zPabrik
	qBA = " select nvl(sum(rate_1), 0) as rate_1" & _
		  "        , nvl(sum(rate_2), 0) as rate_2" & _
		  "        , nvl(sum(rate_3), 0) as rate_3" & _
		  "        , nvl(sum(rate_4), 0) as rate_4" & _
		  " from allowded" & _
		  " where fccode = 'A17'" & _
		  "       and fcba = '" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.CursorType = 3
	rsBA.CursorLocation = 3
	rsBA.Open qBA, connstr
	if not rsBA.eof then
	   nHargaBerasJamsostek = cdbl(rsBA("Rate_1"))
	   zKebun = cdbl(rsBA("Rate_2"))
	   zPabrik = cdbl(rsBA("Rate_3"))
	   zOthers = Cdbl(rsBA("rate_4"))
	else
	   nHargaBerasJamsostek=0
	   zKebun = 0
	   zPabrik = 0
	   zOthers = 0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY9002 Harga Beras AllowDed Code A17 is Empty"
	   err.number=9002
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

	qReceive = " SELECT em.fccode, NVL(em.naturetype, 'NN') naturetype, em.maritalstatus, em.presentgrade, NVL(SUM(em.presentrate),0) AS presentrate, NVL(SUM(emp.amount),0) AS gaji, em.emp_type" & _
	           " FROM employee em" & _
	           " LEFT JOIN (SELECT fcba, empcode, SUM(amount) amount FROM empallded WHERE fcba = '" & Session("fcBA") & "' AND TO_CHAR(adhocdate, 'yyyymm') = '" & Year(Session("v_Employee")) & "'||LPAD('" & Month(Session("v_Employee")) & "', 2, 0) AND allowdedcode = 'A01' GROUP BY fcba, empcode ) emp ON (em.fcba = emp.fcba AND em.fccode = emp.empcode)" & _
               " WHERE em.DATETERMINATE IS NULL" & _
               "       AND em.presentgrade IN ('KBT','KHT','KHL')" & _
			   "       AND em.jamsostekcategory <> 'NJ' " & _
               "       AND em.fcba = '" & Session("fcBA") & "'" & _
			   " GROUP BY em.fccode, NVL(em.naturetype, 'NN'), em.maritalstatus, em.presentgrade, em.emp_type" 
	  
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate Jamsostek..."& "<br>"
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
	   Do While Not rsEmpAllDed.eof
	   
			If rsEmpAllDed("naturetype") = "NN" Then
				nKgBerasPekerja   = 0
			Else
				'Dim rsParameter,qParameter
				Set rsParameter = Server.CreateObject("ADODB.Recordset")
				qParameter =" Select nvl(sum(RATE_2),0) as Rate_2 From PARAMETER where PARCODE='E0006' and " & _
			              " FCCODE='" & rsEmpAllDed("MARITALSTATUS") & "'" & _
				          " and fcBA='" & Session("fcBA") & "'"
				rsParameter.CursorType = 3
				rsParameter.CursorLocation = 3
				rsParameter.Open qParameter, connstr
				if not rsParameter.eof then
					nKgBerasPekerja   = cdbl(rsParameter("Rate_2"))
				else
					nKgBerasPekerja   = 0
				End if
				rsParameter.Close
				set rsParameter = Nothing
			End If
		  
		  
			Dim rsJamsostekCategory, qJamsostekCategory, vJamsostekCategory
			Set rsJamsostekCategory = Server.CreateObject("ADODB.Recordset")

			qJamsostekCategory	=	" SELECT TRIM(JAMSOSTEKCATEGORY) JAMSOSTEKCATEGORY"& _
									" FROM EMPLOYEE" & _
									" WHERE FCBA = '" & Session("FCBA") & "'" & _
									"       AND FCCODE = '" & rsEmpAllDed("FCCODE") & "'"
			rsJamsostekCategory.CursorType = 3
			rsJamsostekCategory.CursorLocation = 3
			rsJamsostekCategory.Open qJamsostekCategory, connstr

			if not rsJamsostekCategory.eof then
				vJamsostekCategory   = rsJamsostekCategory("JAMSOSTEKCATEGORY")
			else
				vJamsostekCategory   = "NJ"
			End if

			rsJamsostekCategory.Close
			set rsJamsostekCategory = Nothing

			vAll_JamSostek=0
			vDed_JamSostek=0
			
			PercentAllowded = GetPercentAllowded(vJamsostekCategory, Session("FCBA"))
			PercentDeduction = GetPercentDeduction(vJamsostekCategory, Session("FCBA"))
			
			
			If rsEmpAllDed("PRESENTGRADE") = "KHL" Then
				If rsEmpAllDed("emp_type") = "PEMANEN" Then
					vAll_JamSostek = ( ( ( ( cdbl(rsEmpAllDed("PRESENTRATE")) - 250)* zOthers) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * PercentAllowded ) / 100
					vDed_JamSostek = ( ( ( ( cdbl(rsEmpAllDed("PRESENTRATE")) - 250)* zOthers) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * PercentDeduction ) / 100
				Else 
					vAll_JamSostek = ( ( (cdbl(rsEmpAllDed("PRESENTRATE")) * zOthers) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * PercentAllowded ) / 100
					vDed_JamSostek = ( ( (cdbl(rsEmpAllDed("PRESENTRATE")) * zOthers) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * PercentDeduction ) / 100
				End If					
			ElseIf rsEmpAllDed("PRESENTGRADE") = "KBT" OR rsEmpAllDed("PRESENTGRADE") = "KHT" Then
				vAll_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * PercentAllowded ) / 100
				vDed_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * PercentDeduction ) / 100
			Else
				vAll_JamSostek = 0
				vDed_JamSostek = 0
			End If
			
			'Response.Write " <br/> Emp: " & rsEmpAllDed("FCCODE") & ", All: " & PercentAllowded & ", value: " & vAll_JamSostek
			'Response.Write " <br/> Emp: " & rsEmpAllDed("FCCODE") & ", Ded: " & PercentDeduction & ", value:" & vDed_JamSostek
			
			' If vJamsostekCategory = "J1" Then
				' If rsEmpAllDed("PRESENTGRADE") = "KHL" Then
					' vAll_JamSostek = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zKebun ) / 100
					' vDed_JamSostek = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yKebun ) / 100
				' ElseIf rsEmpAllDed("PRESENTGRADE") = "KBT" OR rsEmpAllDed("PRESENTGRADE") = "KHT" Then
					' vAll_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zKebun ) / 100
					' vDed_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yKebun ) / 100
				' Else
					' vAll_JamSostek = 0
					' vDed_JamSostek = 0
				' End If
			' ElseIf vJamsostekCategory = "J2" Then
				' If rsEmpAllDed("PRESENTGRADE") = "KHL" Then
					' vAll_JamSostek = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zPabrik ) / 100
					' vDed_JamSostek = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yPabrik ) / 100
				' ElseIf rsEmpAllDed("PRESENTGRADE") = "KBT" OR rsEmpAllDed("PRESENTGRADE") = "KHT" Then
					' vAll_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zPabrik ) / 100
					' vDed_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yPabrik ) / 100
				' Else
					' vAll_JamSostek = 0
					' vDed_JamSostek = 0
				' End If
			' ElseIf vJamsostekCategory = "J3" Then
				' If rsEmpAllDed("PRESENTGRADE") = "KHL" Then
					' vAll_JamSostek = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zOthers ) / 100
					' vDed_JamSostek = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yOthers ) / 100
				' ElseIf rsEmpAllDed("PRESENTGRADE") = "KBT" OR rsEmpAllDed("PRESENTGRADE") = "KHT" Then
					' vAll_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zOthers ) / 100
					' vDed_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yOthers ) / 100
				' Else
					' vAll_JamSostek = 0
					' vDed_JamSostek = 0
				' End If
			' Else
				' vAll_JamSostek = 0
				' vDed_JamSostek = 0
			' End If
			
		  ' if vEstateorMill="M" and rsEmpAllDed("PRESENTGRADE")="KHL" then
		     ' vAll_JamSostek=( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zPabrik ) / 100
		     ' vDed_JamSostek=( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yPabrik ) / 100
		  ' elseif vEstateorMill="E" and (rsEmpAllDed("PRESENTGRADE")="KBT" OR rsEmpAllDed("PRESENTGRADE")="KHT" ) then
		     ' vAll_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zKebun ) / 100
		     ' vDed_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yKebun ) / 100
		  ' elseif vEstateorMill="E" and (rsEmpAllDed("PRESENTGRADE")="KHL" ) then
		     ' vAll_JamSostek=( ( ( cdbl(rsEmpAllDed("PRESENTRATE")) * 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zKebun ) / 100
		     ' vDed_JamSostek=( ( ( cdbl(rsEmpAllDed("PRESENTRATE")) * 25)+ (nKgBerasPekerja*nHargaBerasJamsostek) ) * yKebun ) / 100
		  ' elseif vEstateorMill="M" and (rsEmpAllDed("PRESENTGRADE")="KHT" or rsEmpAllDed("PRESENTGRADE")="KBT") then
		     ' vAll_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zKebun ) / 100
		     ' vDed_JamSostek=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yKebun ) / 100
		  ' end if
		  
		  
		  if vAll_JamSostek <> 0 then
				maxLineNo = maxLineNo + 1
				qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
					    "(" & maxLineNo & ",'" & rsEmpAllDed("FCCODE") & "','" & "A17" & "'," &  " to_date('" & Session("v_Employee") & "','mm/dd/yyyy') " & "," & vAll_JamSostek & _
					    ",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					    Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
				connstr.execute(qAllowance)
			    
				' *** Remarked History By Iwa 08/01/2019 ***
				
				'	Insert the History of EmpallDed Data 
			    ' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
				' txtRefStringValues_Receive = rsEmpAllDed("FCCODE") & txtRefStringSeparator_Receive & _
									         Session("fcBA")
				
				'detailLineNo = 1
				'qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
				'									" LINEHEADERNO," & _
				'									" LINENO," & _
				'									" FCBA," & _
				'									" REF_PROCESS," & _
				'									" REF_PROCESSNAME," & _
				'									" REF_TABLESOURCE," & _
				'									" REF_NUMBERKEYS," & _
				'									" REF_STRINGSEPARATOR," & _
				'									" REF_STRINGKEYS," & _
				'									" REF_STRINGVALUES," & _
				'									" AMOUNT, " & _
				'									" FCENTRY," & _
				'									" FCEDIT," & _
				'									" FCIP," & _
				'									" LASTUPDATE," & _
				'									" LASTTIME)" & _
				'							" VALUES ( " & _
				'									" '" & maxLineNo & "'," & _
				'									" '" & detailLineNo & "'," & _
				'									" '" & Session("fcBA") & "'," & _
				'									" '" & tempProcessCode & "'," & _
				'									" '" & tempProcessNameCode & "'," & _
				'									" '" & tempRefCodeReceive & "'," & _
				'									" '" & txtRefNumberKeys_Receive & "'," & _
				'									" '" & txtRefStringSeparator_Receive & "'," & _
				'									" '" & txtRefStringKeys_Receive & "'," & _
				'									" '" & txtRefStringValues_Receive & "'," & _
				'									vAll_JamSostek & "," & _
				'									" '" & Session("CodeName") & "'," & _
				'									" '" & Session("CodeName") & "'," & _
				'									" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
				'									" sysdate, " & _
				'									" to_char(sysdate, 'HH24:MI') " & _
				'									" ) "
				
			    'connstr.execute(qAllowanceHistory)			 
		  End if		

		  if vDed_JamSostek <> 0 then
		        maxLineNo = maxLineNo + 1
				qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
					    "(" & maxLineNo & ",'" & rsEmpAllDed("FCCODE") & "','" & "D02" & "'," &  " to_date('" & Session("v_Employee") & "','mm/dd/yyyy') " & "," & (vDed_JamSostek * -1) & _
					    ",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					    Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
				connstr.execute(qAllowance)
			 
				' *** Remarked History by Iwa : 08/01/2019 
				'	Insert the History of EmpallDed Data 
			    ' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
				' txtRefStringValues_Receive = rsEmpAllDed("FCCODE") & txtRefStringSeparator_Receive & _
									         Session("fcBA")
				
				' detailLineNo = 1
				' qAllowanceHistory = "INSERT INTO empallded_hist_tab ( " & _
				'									" LINEHEADERNO," & _
				'									" LINENO," & _
				'									" FCBA," & _
				'									" REF_PROCESS," & _
				'									" REF_PROCESSNAME," & _
				'									" REF_TABLESOURCE," & _
				'									" REF_NUMBERKEYS," & _
				'									" REF_STRINGSEPARATOR," & _
				'									" REF_STRINGKEYS," & _
				'									" REF_STRINGVALUES," & _
				'									" AMOUNT, " & _
				'									" FCENTRY," & _
				'									" FCEDIT," & _
				'									" FCIP," & _
				'									" LASTUPDATE," & _
				'									" LASTTIME)" & _
				'							" VALUES ( " & _
				'									" '" & maxLineNo & "'," & _
				'									" '" & detailLineNo & "'," & _
				'									" '" & Session("fcBA") & "'," & _
				'									" '" & tempProcessCode & "'," & _
				'									" '" & tempProcessNameCode & "'," & _
				'									" '" & tempRefCodeReceive & "'," & _
				'									" '" & txtRefNumberKeys_Receive & "'," & _
				'									" '" & txtRefStringSeparator_Receive & "'," & _
				'									" '" & txtRefStringKeys_Receive & "'," & _
				'									" '" & txtRefStringValues_Receive & "'," & _
				'									-1 * vDed_JamSostek & "," & _
				'									" '" & Session("CodeName") & "'," & _
				'									" '" & Session("CodeName") & "'," & _
				'									" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
				'									" sysdate, " & _
				'									" to_char(sysdate, 'HH24:MI') " & _
				'									" ) "
			  
			    'connstr.execute(qAllowanceHistory)			 
		  End if		

		  response.Flush()
		  rsEmpAllDed.MoveNext
	  loop
	  set nKgBerasPekerja = Nothing
	  set nHargaBerasJamsostek = Nothing
	End if
	rsEmpAllDed.Close
	set rsEmpAllDed = Nothing
%>