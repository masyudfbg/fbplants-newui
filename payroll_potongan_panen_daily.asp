<%
	Dim rate_buah_mentah
		' Cari RateBuah Mentah
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D09' and fcBA='" & Session("fcBA") & "'"
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		'response.Write(qba)
		response.Flush()
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_buah_mentah=cdbl(rsBA("Rate_1"))
		else
		   rate_buah_mentah=0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Mentah is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		Dim rate_tangkai_panjang
			' Cari RateBuah Masak
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D10' and fcBA='" & Session("fcBA") & "'"
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		'response.Write(qba)
		response.Flush()
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_tangkai_panjang=cdbl(rsBA("Rate_1"))
		else
		   rate_tangkai_panjang=0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing

		Dim rate_buah_tidak_dipanen
			' Cari rate_buah_tidak_dipanen
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D12' and fcBA='" & Session("fcBA") & "'"
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		'response.Write(qba)
		response.Flush()
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_buah_tidak_dipanen=cdbl(rsBA("Rate_1"))
		else
		   rate_buah_tidak_dipanen=0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing

		Dim rate_pelepah_tidak_disusun
			' Cari rate_pelepah_tidak_disusun
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D14' and fcBA='" & Session("fcBA") & "'"
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		'response.Write(qba)
		response.Flush()
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_pelepah_tidak_disusun=cdbl(rsBA("Rate_1"))
		else
		   rate_pelepah_tidak_disusun=0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		Dim rate_brondolan_dipiringan
			' Cari rate_brondolan_dipiringan
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D15' and fcBA='" & Session("fcBA") & "'"
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		'response.Write(qba)
		response.Flush()
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_brondolan_dipiringan=cdbl(rsBA("Rate_1"))
		else
		   rate_brondolan_dipiringan=0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing


		Dim rate_buah_tinggal_dipiringan
			' Cari rate_buah_tinggal_dipiringan
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D16' and fcBA='" & Session("fcBA") & "'"
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		'response.Write(qba)
		response.Flush()
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_buah_tinggal_dipiringan=cdbl(rsBA("Rate_1"))
		else
		   rate_buah_tinggal_dipiringan=0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing


		Dim rate_buah_matahari
			' Cari rate_buah_matahari
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D17' and fcBA='" & Session("fcBA") & "'"
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		'response.Write(qba)
		response.Flush()
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_buah_matahari=cdbl(rsBA("Rate_1"))
		else
		   rate_buah_matahari=0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		
		' **** POTONGAN BUAH MENTAH DIPERAM **** '
		Dim rate_buah_mentah_diperam
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D27' and fcBA='" & Session("fcBA") & "'"
		' response.Write(qba)
		' response.Flush()
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_buah_mentah_diperam = cdbl(rsBA("Rate_1"))
		else
		   rate_buah_mentah_diperam = 0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Mentah Diperam is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		
		' **** POTONGAN OVER PRUNING **** '
		Dim rate_over_pruning
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D28' and fcBA='" & Session("fcBA") & "'"
		' response.Write(qba)
		' response.Flush()
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_over_pruning = cdbl(rsBA("Rate_1"))
		else
		   rate_over_pruning = 0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Over Pruning is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		
		' **** POTONGAN BRONDOLAN TIDAK DI ALAS **** '
		Dim rate_brondolan_tidak_di_alas
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D29' and fcBA='" & Session("fcBA") & "'"
		' response.Write(qba)
		' response.Flush()
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_brondolan_tidak_di_alas = cdbl(rsBA("Rate_1"))
		else
		   rate_brondolan_tidak_di_alas = 0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Brondolan Tidak di Alas is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		
		' **** POTONGAN BRONDOLAN KOTOR/SAMPAH **** '
		Dim rate_brondolan_kotor_sampah
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D30' and fcBA='" & Session("fcBA") & "'"
		' response.Write(qba)
		' response.Flush()
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_brondolan_kotor_sampah = cdbl(rsBA("Rate_1"))
		else
		   rate_brondolan_kotor_sampah = 0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Brondolan Kotor/ Sampah is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		
		' **** POTONGAN BUAH DIBELAH **** '
		Dim rate_buah_dibelah
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D31' and fcBA='" & Session("fcBA") & "'"
		' response.Write(qba)
		' response.Flush()
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_buah_dibelah = cdbl(rsBA("Rate_1"))
		else
		   rate_buah_dibelah = 0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah di Belah is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		
		' **** POTONGAN PELEPAH SENGKLEH **** '
		Dim rate_pelepah_sengkleh
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D13' and fcBA='" & Session("fcBA") & "'"
		' response.Write(qba)
		' response.Flush()
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_pelepah_sengkleh = cdbl(rsBA("Rate_1"))
		else
		   rate_pelepah_sengkleh = 0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Pelepah Sengkleh is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		
		
		' **** POTONGAN BUAH TIDAK DISUSUN RAPI DI TPH **** '
		Dim rate_buah_tidak_disusun_rapi
		qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D32' and fcBA='" & Session("fcBA") & "'"
		' response.Write(qba)
		' response.Flush()
		Set rsBA = Server.CreateObject("ADODB.Recordset")
		rsBA.Open qBA, connstr,3,3
		if not rsBA.eof then
		   rate_buah_tidak_disusun_rapi = cdbl(rsBA("Rate_1"))
		else
		   rate_buah_tidak_disusun_rapi = 0
		   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Tidak Disusun Rapi di TPH is Null "
		   err.number=10001
		end if
		rsBA.close
		set rsBA = Nothing
		set qBA = Nothing
		response.Write "Start Potongan Panen..." & "<br>" 
		response.Flush()
		
		qty_buah_mentah = 0
		qty_tangkai_panjang = 0
		qty_buah_tidak_dipanen = 0
		qty_pelepah_tidak_disusun = 0
		qty_brondolan_dipiringan = 0
		qty_buah_tinggal_dipiringan = 0
		qty_buah_matahari = 0
		qty_buah_mentah_diperam = 0
		qty_over_pruning = 0
		qty_brondolan_tidak_di_alas = 0
		qty_brondolan_kotor_sampah = 0
		qty_buah_dibelah = 0
		qty_pelepah_sengkleh = 0
		qty_buah_tidak_disusun_rapi = 0
								
						  'Cari Potongan Kalau ada
			  qBA = " SELECT EMPCODE as EMPLOYEECODE, FDDATE, NVL(SUM(UNDER_RIPE),0) as UNDER_RIPE" & _
					"        , NVL(SUM(LONG_STALK),0) as LONG_STALK" & _
					"        , NVL(SUM(UNHARVEST_FFB),0) as UNHARVEST_FFB" & _
					"        , NVL(SUM(QE_1),0) as QE_1" & _
					"        , NVL(SUM(UNCOLLECT_LF_CIRCLE),0) as UNCOLLECT_LF_CIRCLE" & _
					"        , NVL(SUM(UNCOLLECT_LF_PIECE),0) as UNCOLLECT_LF_PIECE" & _
					"        , NVL(SUM(QE_2),0) as QE_2" & _
					"        , NVL(SUM(QE_4),0) as QE_4" & _
					"        , NVL(SUM(QE_5),0) as QE_5" & _
					"        , NVL(SUM(QE_6),0) as QE_6" & _
					"        , NVL(SUM(QE_7),0) as QE_7" & _
					"        , NVL(SUM(QE_8),0) as QE_8" & _
					"        , NVL(SUM(UNPRUNE_FROND),0) as UNPRUNE_FROND" & _
					"        , NVL(SUM(UNARRANGE_FFB),0) as UNARRANGE_FFB" & _
					" FROM HARVESTINGQUALITY" & _
					" WHERE " & _
					" FDDATE= to_Date('" & tglnow & "','mm/dd/yyyy')" &  _
					" AND fcBA = '" & Session("fcBA") & "'" & _
					" group by EMPCODE, FDDATE"
						  
			 Set rsBA = Server.CreateObject("ADODB.Recordset")
			 'response.Write(qba)
			 'response.Flush()
			 rsBA.Open qBA, connstr
			 if not rsBA.eof then
				Do While Not rsBA.eof
					qty_buah_mentah=cdbl(rsBA("UNDER_RIPE"))
					qty_tangkai_panjang=cdbl(rsBA("LONG_STALK"))
					qty_buah_tidak_dipanen=cdbl(rsBA("UNHARVEST_FFB"))
					qty_pelepah_tidak_disusun=cdbl(rsBA("QE_1"))
					qty_brondolan_dipiringan=cdbl(rsBA("UNCOLLECT_LF_CIRCLE"))
					qty_buah_tinggal_dipiringan=cdbl(rsBA("UNCOLLECT_LF_PIECE"))
					qty_buah_matahari=cdbl(rsBA("QE_2"))
					qty_buah_mentah_diperam=cdbl(rsBA("QE_4"))
					qty_over_pruning=cdbl(rsBA("QE_5"))
					qty_brondolan_tidak_di_alas=cdbl(rsBA("QE_6"))
					qty_brondolan_kotor_sampah=cdbl(rsBA("QE_7"))
					qty_buah_dibelah=cdbl(rsBA("QE_8"))
					qty_pelepah_sengkleh=cdbl(rsBA("UNPRUNE_FROND"))
					qty_buah_tidak_disusun_rapi=cdbl(rsBA("UNARRANGE_FFB"))
								  	
					'response.Write rsBA("EMPLOYEECODE") & " - " & rsBA("FDDATE") &"<br>" 
					response.Flush()

								 if qty_buah_mentah*rate_buah_mentah>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
												"(" & maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D09" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_buah_mentah*rate_buah_mentah & _
												",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
									
								  end if
								  
							    'if qty_buah_mentah*rate_buah_mentah>0 then
								'	 maxLineNo = maxLineNo + 1
								'	 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
								'				"(" & maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D09" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_buah_mentah*rate_buah_mentah & _
								'				",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								'				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
								'	 connstr.execute(qAllowance)	
								'end if

								  if qty_tangkai_panjang*rate_tangkai_panjang>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
												"(" & maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D10" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_tangkai_panjang*rate_tangkai_panjang & _
												",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
									 'response.write 	qAllowance		
									 connstr.execute(qAllowance)
									 
								  end if


								  if qty_buah_tidak_dipanen*rate_buah_tidak_dipanen>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
												"(" & maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D12" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_buah_tidak_dipanen*rate_buah_tidak_dipanen & _
												",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
									 
								  end if


								  if qty_pelepah_tidak_disusun*rate_pelepah_tidak_disusun>0 then
									maxLineNo = maxLineNo + 1
									qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
												"(" & maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D14" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_pelepah_tidak_disusun*rate_pelepah_tidak_disusun & _
												",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
									connstr.execute(qAllowance)
									 
								  end if

								  if qty_brondolan_dipiringan*rate_brondolan_dipiringan>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
												"(" & maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D15" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_brondolan_dipiringan*rate_brondolan_dipiringan & _
												",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
									 
									
								  end if

								  if qty_buah_tinggal_dipiringan*rate_buah_tinggal_dipiringan>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
												"(" & maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D16" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_buah_tinggal_dipiringan*rate_buah_tinggal_dipiringan & _
												",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
									 
								  end if


								  if qty_buah_matahari*rate_buah_matahari>0 then
										maxLineNo = maxLineNo + 1
										qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
												"("& maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D17" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_buah_matahari*rate_buah_matahari & _
												",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
										connstr.execute(qAllowance)
									 
								  end if
								  
								  ' **** POTONGAN BUAH MENTAH DIPERAM **** '
								  if qty_buah_mentah_diperam*rate_buah_mentah_diperam>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance = " INSERT INTO empallded ( lineno" & _
												   "                        , empcode" & _
												   "                        , allowdedcode" & _
												   "                        , adhocdate" & _
												   "                        , amount" & _
												   "                        , fixedallded" & _
												   "                        , fcentry" & _
												   "                        , fcedit" & _
												   "                        , fcip" & _
												   "                        , fcba" & _
												   "                        , lastupdate" & _
												   "                        , lasttime)" & _
												   " VALUES ( " & maxLineNo & "" & _
												   "           , '" & rsBA("EMPLOYEECODE") & "'" & _
												   "           , 'D27'" & _
												   "           , TO_DATE('" & rsBA("FDDATE") & "', 'mm/dd/yyyy') " & _
												   "           , '" & qty_buah_mentah_diperam*rate_buah_mentah_diperam & "'" & _
												   "           , '0'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												   "           , '" & Session("fcBA") & "'" & _
												   "           , SYSDATE" & _
												   "           , TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
								  end if
								  
								  ' **** POTONGAN OVER PRUNING **** '
								  if qty_over_pruning*rate_over_pruning>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance = " INSERT INTO empallded ( lineno" & _
												   "                        , empcode" & _
												   "                        , allowdedcode" & _
												   "                        , adhocdate" & _
												   "                        , amount" & _
												   "                        , fixedallded" & _
												   "                        , fcentry" & _
												   "                        , fcedit" & _
												   "                        , fcip" & _
												   "                        , fcba" & _
												   "                        , lastupdate" & _
												   "                        , lasttime)" & _
												   " VALUES ( " & maxLineNo & "" & _
												   "           , '" & rsBA("EMPLOYEECODE") & "'" & _
												   "           , 'D28'" & _
												   "           , TO_DATE('" & rsBA("FDDATE") & "', 'mm/dd/yyyy') " & _
												   "           , '" & qty_over_pruning*rate_over_pruning & "'" & _
												   "           , '0'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												   "           , '" & Session("fcBA") & "'" & _
												   "           , SYSDATE" & _
												   "           , TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
								  end if
								  
								  ' **** POTONGAN BRONDOLAN TIDAK DI ALAS **** '
								  if qty_brondolan_tidak_di_alas*rate_brondolan_tidak_di_alas>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance = " INSERT INTO empallded ( lineno" & _
												   "                        , empcode" & _
												   "                        , allowdedcode" & _
												   "                        , adhocdate" & _
												   "                        , amount" & _
												   "                        , fixedallded" & _
												   "                        , fcentry" & _
												   "                        , fcedit" & _
												   "                        , fcip" & _
												   "                        , fcba" & _
												   "                        , lastupdate" & _
												   "                        , lasttime)" & _
												   " VALUES ( " & maxLineNo & "" & _
												   "           , '" & rsBA("EMPLOYEECODE") & "'" & _
												   "           , 'D29'" & _
												   "           , TO_DATE('" & rsBA("FDDATE") & "', 'mm/dd/yyyy') " & _
												   "           , '" & qty_brondolan_tidak_di_alas*rate_brondolan_tidak_di_alas & "'" & _
												   "           , '0'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												   "           , '" & Session("fcBA") & "'" & _
												   "           , SYSDATE" & _
												   "           , TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
								  end if
								  
								  ' **** POTONGAN BRONDOLAN KOTOR/SAMPAH **** '
								  if qty_brondolan_kotor_sampah*rate_brondolan_kotor_sampah>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance = " INSERT INTO empallded ( lineno" & _
												   "                        , empcode" & _
												   "                        , allowdedcode" & _
												   "                        , adhocdate" & _
												   "                        , amount" & _
												   "                        , fixedallded" & _
												   "                        , fcentry" & _
												   "                        , fcedit" & _
												   "                        , fcip" & _
												   "                        , fcba" & _
												   "                        , lastupdate" & _
												   "                        , lasttime)" & _
												   " VALUES ( " & maxLineNo & "" & _
												   "           , '" & rsBA("EMPLOYEECODE") & "'" & _
												   "           , 'D30'" & _
												   "           , TO_DATE('" & rsBA("FDDATE") & "', 'mm/dd/yyyy') " & _
												   "           , '" & qty_brondolan_kotor_sampah*rate_brondolan_kotor_sampah & "'" & _
												   "           , '0'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												   "           , '" & Session("fcBA") & "'" & _
												   "           , SYSDATE" & _
												   "           , TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
								  end if
								  
								  ' **** POTONGAN BUAH DIBELAH **** '
								  if qty_buah_dibelah*rate_buah_dibelah>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance = " INSERT INTO empallded ( lineno" & _
												   "                        , empcode" & _
												   "                        , allowdedcode" & _
												   "                        , adhocdate" & _
												   "                        , amount" & _
												   "                        , fixedallded" & _
												   "                        , fcentry" & _
												   "                        , fcedit" & _
												   "                        , fcip" & _
												   "                        , fcba" & _
												   "                        , lastupdate" & _
												   "                        , lasttime)" & _
												   " VALUES ( " & maxLineNo & "" & _
												   "           , '" & rsBA("EMPLOYEECODE") & "'" & _
												   "           , 'D31'" & _
												   "           , TO_DATE('" & rsBA("FDDATE") & "', 'mm/dd/yyyy') " & _
												   "           , '" & qty_buah_dibelah*rate_buah_dibelah & "'" & _
												   "           , '0'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												   "           , '" & Session("fcBA") & "'" & _
												   "           , SYSDATE" & _
												   "           , TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
								  end if
								  
								  if qty_pelepah_sengkleh*rate_pelepah_sengkleh>0 then
										  maxLineNo = maxLineNo + 1
										  qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
													"(" & maxLineNo & ",'" & rsBA("EMPLOYEECODE") & "','" & "D13" & "'," &  " to_date('" & rsBA("FDDATE") & "','mm/dd/yyyy') " & "," & qty_pelepah_sengkleh*rate_pelepah_sengkleh & _
													",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
													Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI'))"
										 connstr.execute(qAllowance)	 
								   end if
								  
								  ' **** POTONGAN BUAH TIDAK DISUSUN RAPI DI TPH **** '
								  if qty_buah_tidak_disusun_rapi*rate_buah_tidak_disusun_rapi>0 then
									 maxLineNo = maxLineNo + 1
									 qAllowance = " INSERT INTO empallded ( lineno" & _
												   "                        , empcode" & _
												   "                        , allowdedcode" & _
												   "                        , adhocdate" & _
												   "                        , amount" & _
												   "                        , fixedallded" & _
												   "                        , fcentry" & _
												   "                        , fcedit" & _
												   "                        , fcip" & _
												   "                        , fcba" & _
												   "                        , lastupdate" & _
												   "                        , lasttime)" & _
												   " VALUES ( " & maxLineNo & "" & _
												   "           , '" & rsBA("EMPLOYEECODE") & "'" & _
												   "           , 'D32'" & _
												   "           , TO_DATE('" & rsBA("FDDATE") & "', 'mm/dd/yyyy') " & _
												   "           , '" & qty_buah_tidak_disusun_rapi*rate_buah_tidak_disusun_rapi & "'" & _
												   "           , '0'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Session("CodeName") & "'" & _
												   "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												   "           , '" & Session("fcBA") & "'" & _
												   "           , SYSDATE" & _
												   "           , TO_CHAR(SYSDATE,'HH24:MI'))"
									 connstr.execute(qAllowance)
								  end if			
					rsBA.MoveNext
				loop
			end if
			rsBA.Close
			set rsBA = Nothing
		
		response.Write "Finish Calculate Potongan Panen..." & "<br>" 
		response.Flush()
%>