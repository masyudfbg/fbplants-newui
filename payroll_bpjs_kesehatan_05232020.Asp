<%
	'Amandemen BPJS Kesehatan
	'1. Parameter Master Status Pekerja di Parameter = E0006
	'   Rate 1 = Harga Beras, Rate 2 = Kg Beras Pekerja, Rate 3 = Kg Beras Pasangan, Rate 4 = Kg Beras Tanggungan
	'2. Allowded Code Untuk BPJS Kesehatan = A60 Tunjangan D24 Potongan 
	'   A60 : Rate 1 Harga Beras BPJS Kesehatan
	'       : Rate2 z Kebun
	'       : Rate3 z Pabrik
	'       : Rate4 z others
	
	'   D24 : Rate2 y Kebun
	'       : Rate3 y Pabrik
	'       : Rate4 y others
	
	'3. Bedakan Pabrik dan Kebun dari Business Unit Kolom fcType E or M

	Dim vall_bpjskesehatan, vded_bpjskesehatan
    ' Cari Kebun / Pabrik   Type E = Estate M = Mill
	
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
		  " where fccode = 'D24'" & _
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
		  " where fccode = 'A60'" & _
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
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY9002 Harga Beras AllowDed Code A60 is Empty"
		err.number = 9002
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

	qReceive = " select fccode" & _
	           "        , nvl(naturetype, 'NN') naturetype" & _
	           "        , maritalstatus" & _
	           "        , presentgrade" & _
	           "        , nvl(sum(presentrate),0) as presentrate" & _
			   "        , emp_type" & _
	           " from employee " & _
	           " where dateterminate is null" & _
	           "       and presentgrade in ('KBT','KHT','KHL')" & _
	           "       and bpjskes_category <> 'NB'" & _
	           "       and fcba='" & Session("fcBA") & "'" & _
	           " group by fccode" & _
	           "          , nvl(naturetype, 'NN')" & _
	           "          , maritalstatus" & _
	           "          , presentgrade" & _
	           "          , emp_type "	   
	  
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate BPJS Kesehatan..."& "<br>"
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
		Do While Not rsEmpAllDed.eof
		
			If rsEmpAllDed("naturetype") = "NN" Then
				nKgBerasPekerja   = 0
			Else
				'Dim rsParameter,qParameter
				Set rsParameter = Server.CreateObject("ADODB.Recordset")
				qParameter = " select nvl(sum(rate_2),0) as rate_2" & _
	                         " from parameter" & _
	                         " where parcode = 'E0006'" & _
	                         "       and fccode = '" & rsEmpAllDed("MARITALSTATUS") & "'" & _
				             "       and fcba = '" & Session("fcBA") & "'"
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
		  
		  
			Dim rsbpjskesehatan, sqlbpjskesehatan, vbpjscategory
			Set rsbpjskesehatan = Server.CreateObject("ADODB.Recordset")

			sqlbpjskesehatan	=	" select trim(bpjskes_category) bpjskes_category"& _
									" from employee" & _
									" where fcba = '" & Session("FCBA") & "'" & _
									"       and fccode = '" & rsEmpAllDed("FCCODE") & "'"
			rsbpjskesehatan.CursorType = 3
			rsbpjskesehatan.CursorLocation = 3
			rsbpjskesehatan.Open sqlbpjskesehatan, connstr

			if not rsbpjskesehatan.eof then
				vbpjscategory   = rsbpjskesehatan("bpjskes_category")
			else
				vbpjscategory   = "NB"
			End if

			rsbpjskesehatan.Close
			set rsbpjskesehatan = Nothing

			vall_bpjskesehatan=0
			vded_bpjskesehatan=0
			
			If vbpjscategory = "BK01" Then
				If rsEmpAllDed("PRESENTGRADE") = "KHL" Then
					If rsEmpAllDed("emp_type") = "PEMANEN" Then
						vall_bpjskesehatan = ( ( ( ( cdbl(rsEmpAllDed("PRESENTRATE")) - 250)* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zKebun ) / 100
						vded_bpjskesehatan = ( ( ( ( cdbl(rsEmpAllDed("PRESENTRATE")) - 250)* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yKebun ) / 100
					Else 
						vall_bpjskesehatan = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zKebun ) / 100
						vded_bpjskesehatan = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yKebun ) / 100
					End If					
				ElseIf rsEmpAllDed("PRESENTGRADE") = "KBT" OR rsEmpAllDed("PRESENTGRADE") = "KHT" Then
					vall_bpjskesehatan=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zKebun ) / 100
					vded_bpjskesehatan=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yKebun ) / 100
				Else
					vall_bpjskesehatan = 0
					vded_bpjskesehatan = 0
				End If
			ElseIf vbpjscategory = "BK02" Then
					If rsEmpAllDed("PRESENTGRADE") = "KHL" Then
						
					'Edit Tanggal 02/04/2020 atas permintaan pak gibran
					
						if cdbl(rsEmpAllDed("PRESENTRATE")) >= yOthers Then
							vTempPresentRateKHL = cdbl(rsEmpAllDed("PRESENTRATE")) 
						Else
							vTempPresentRateKHL = yOthers 
						End If

						If rsEmpAllDed("emp_type") = "PEMANEN" Then
							vall_bpjskesehatan = ( ( ( ( vTempPresentRateKHL - 0)* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zPabrik ) / 100
							vded_bpjskesehatan = ( ( ( ( vTempPresentRateKHL - 0)* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yPabrik ) / 100
						Else
							vall_bpjskesehatan = ( ( ( vTempPresentRateKHL* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zPabrik ) / 100
							vded_bpjskesehatan = ( ( ( vTempPresentRateKHL* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yPabrik ) / 100
						End If					
						
					ElseIf rsEmpAllDed("PRESENTGRADE") = "KBT" OR rsEmpAllDed("PRESENTGRADE") = "KHT" Then
						
						if cdbl(rsEmpAllDed("PRESENTRATE")) >= zOthers Then
							vTempPresentRate = cdbl(rsEmpAllDed("PRESENTRATE")) 
						Else
							vTempPresentRate = zOthers 
						End If
						
						vall_bpjskesehatan=( ( vTempPresentRate + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zPabrik ) / 100
						vded_bpjskesehatan=( ( vTempPresentRate + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yPabrik ) / 100

					Else
						vall_bpjskesehatan = 0
						vded_bpjskesehatan = 0
					End If
					' End Edit

			ElseIf vbpjscategory = "BK03" Then
				If rsEmpAllDed("PRESENTGRADE") = "KHL" Then				
					If rsEmpAllDed("emp_type") = "PEMANEN" Then
						vall_bpjskesehatan = ( ( ( ( cdbl(rsEmpAllDed("PRESENTRATE")) - 250)* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zOthers ) / 100
						vded_bpjskesehatan = ( ( ( ( cdbl(rsEmpAllDed("PRESENTRATE")) - 250)* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yOthers ) / 100
					Else
						vall_bpjskesehatan = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zOthers ) / 100
						vded_bpjskesehatan = ( ( (cdbl(rsEmpAllDed("PRESENTRATE"))* 25) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yOthers ) / 100
					End If					
				ElseIf rsEmpAllDed("PRESENTGRADE") = "KBT" OR rsEmpAllDed("PRESENTGRADE") = "KHT" Then
					vall_bpjskesehatan=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * zOthers ) / 100
					vded_bpjskesehatan=( ( cdbl(rsEmpAllDed("PRESENTRATE")) + (nKgBerasPekerja*nHargaBerasJamsostek) ) * yOthers ) / 100
				Else
					vall_bpjskesehatan = 0
					vded_bpjskesehatan = 0
				End If
			Else
				vall_bpjskesehatan = 0
				vded_bpjskesehatan = 0
			End If		  
		  
			if vall_bpjskesehatan <> 0 then
			
				maxLineNo = maxLineNo + 1
				qAllowance = " insert into empallded (lineno" & _
				             "                        , empcode" & _
				             "                        , allowdedcode" & _
				             "                        , adhocdate" & _
				             "                        , amount" & _
				             "                        , fcentry" & _
				             "                        , fcedit" & _
				             "                        , fcip" & _
				             "                        , fcba" & _
				             "                        , lastupdate" & _
				             "                        , lasttime)" & _
				             " values(" & maxLineNo & "" & _
				             "        , '" & rsEmpAllDed("FCCODE") & "'" & _
				             "        , 'A60'" & _
				             "        , last_day(to_date('" & Session("v_Employee") & "', 'mm/dd/yyyy'))" & _
				             "        , " & vall_bpjskesehatan & "" & _
				             "        , '" & Session("CodeName") & "'" & _
				             "        , '" & Session("CodeName") & "'" & _
				             "        , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				             "        , '" & Session("fcBA") & "'" & _
				             "        , sysdate" & _
				             "        , to_char(sysdate,'HH24:MI'))"
				connstr.execute(qAllowance)
				
			End if

			if vded_bpjskesehatan <> 0 then
		        maxLineNo = maxLineNo + 1
				qAllowance = " insert into empallded (lineno" & _
				             "                        , empcode" & _
				             "                        , allowdedcode" & _
				             "                        , adhocdate" & _
				             "                        , amount" & _
				             "                        , fcentry" & _
				             "                        , fcedit" & _
				             "                        , fcip" & _
				             "                        , fcba" & _
				             "                        , lastupdate" & _
				             "                        , lasttime)" & _
				             " values(" & maxLineNo & "" & _
				             "        , '" & rsEmpAllDed("FCCODE") & "'" & _
				             "        , 'D24'" & _
				             "        , last_day(to_date('" & Session("v_Employee") & "', 'mm/dd/yyyy'))" & _
				             "        , " & vded_bpjskesehatan & "" & _
				             "        , '" & Session("CodeName") & "'" & _
				             "        , '" & Session("CodeName") & "'" & _
				             "        , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				             "        , '" & Session("fcBA") & "'" & _
				             "        , sysdate" & _
				             "        , to_char(sysdate,'HH24:MI'))"

				connstr.execute(qAllowance)
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