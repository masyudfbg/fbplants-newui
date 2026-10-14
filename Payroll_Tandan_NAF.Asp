<%
	'Amandemen Premi Pandan
	'1. Code Allowance Panen = A24
	'   Qty Basis = Rate1 
	'   Premi Pertandan = Rate 2
	'2. Potongan Buah Mentah Code : D09 Rate1 = Rate
	'2. Potongan Buah Terlalu Masak Code : D10 Rate1 = Rate
	Dim nQtyBasis,nQtyPremiTandan,nRpPremiTandan,nRatePremiTandan
	Dim nQtyBuahMentah,nQtyTerlaluMasak,nTangkaiPanjang,nBuahTinggal,nTdkDisusun,nPelepah,nPremiPanen
	
  

	Dim nRateBuahMentah
    ' Cari RateBuah Mentah
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D09' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nRateBuahMentah=cdbl(rsBA("Rate_1"))
	else
	   nRateBuahMentah=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Mentah is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

	Dim nRateBuahMasak
        ' Cari RateBuah Masak
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D10' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nRateBuahMasak=cdbl(rsBA("Rate_1"))
	else
	   nRateBuahMasak=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

	Dim nBuahtdkdipanen
        ' Cari nBuahtdkdipanen
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D12' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nBuahtdkdipanen=cdbl(rsBA("Rate_1"))
	else
	   nBuahtdkdipanen=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

	Dim nPelepahtdkdipotong
        ' Cari nPelepahtdkdipotong
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D13' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nPelepahtdkdipotong=cdbl(rsBA("Rate_1"))
	else
	   nPelepahtdkdipotong=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

	Dim nPelepahtdkdisusun
        ' Cari nPelepahtdkdisusun
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D14' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nPelepahtdkdisusun=cdbl(rsBA("Rate_1"))
	else
	   nPelepahtdkdisusun=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
	
	Dim nBrondolandipiringan
        ' Cari nBrondolandipiringan
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D15' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nBrondolandipiringan=cdbl(rsBA("Rate_1"))
	else
	   nBrondolandipiringan=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing


	Dim nbuahTdkdiTPH
        ' Cari nbuahTdkdiTPH
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D16' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nbuahTdkdiTPH=cdbl(rsBA("Rate_1"))
	else
	   nbuahTdkdiTPH=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing


	Dim nPulangtdkMelapor
        ' Cari nPulangtdkMelapor
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='D17' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nPulangtdkMelapor=cdbl(rsBA("Rate_1"))
	else
	   nPulangtdkMelapor=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Rate Buah Terlalu Masak is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
	
	
	
	qDecision = " SELECT   fddate, employeecode, COUNT (*) as totalBlock " & _
				" FROM (SELECT DISTINCT fddate, employeecode, locationcode " & _
								 "FROM (SELECT   att_gad.fddate, att_gad.locationcode, " & _
												"att_gad.employeecode, " & _
												"NVL (SUM (att_gad.unit), 0) AS unit, " & _
												"emp.presentrate AS presentrate, " & _
												"emp.presentgrade AS presentgrade " & _
										   "FROM attendance_gad att_gad, employee emp " & _
										  "WHERE (    att_gad.employeecode = emp.fccode " & _
												 "AND att_gad.fcba = emp.fcba " & _
												") " & _
											"AND emp.dateterminate IS NULL " & _
											"AND att_gad.attendance = 'KJ' " & _
											"AND att_gad.jobcode = '6502101' " & _
											"AND att_gad.locationtype = 'FF' " & _
											"AND att_gad.mandays > 0 " & _
											"AND att_gad.unit > 0 " & _
											"AND att_gad.rate = 0 " & _
											"AND TO_NUMBER (TO_CHAR (att_gad.fddate, 'mm')) = " & Month(Session("v_Employee")) &  _
											"AND TO_NUMBER (TO_CHAR (att_gad.fddate,'yyyy')) = " & Year(Session("v_Employee")) & _
											"AND att_gad.fcba = " & Session("fcBA") & "' " & _
									   "GROUP BY att_gad.fddate, " & _
												"att_gad.locationcode, " & _
												"att_gad.employeecode, " & _
												"emp.presentrate, " & _
												"emp.presentgrade)) " & _
				" GROUP BY fddate, employeecode"
	
    Set rsEmpAllDedDecision = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDedDecision.CursorType = 3
	rsEmpAllDedDecision.CursorLocation = 3
	response.Write "Start Calculate Premi/Potongan Panen..." & "<br>" 
	response.Flush()
	rsEmpAllDedDecision.Open qDecision, connstr
	if not rsEmpAllDedDecision.eof then
	    Do While Not rsEmpAllDedDecision.eof			
			nDecisionTotal = rsEmpAllDedDecision("totalBlock")
			sEmployeeCode = rsEmpAllDedDecision("EmployeeCode")
			dFdDate = rsEmpAllDedDecision("fddate")
			
			if nDecisionTotal = 1 then
			
				qReceive="select att_gad.FDDATE,att_gad.LOCATIONCODE,att_gad.EMPLOYEECODE,nvl(sum(att_gad.unit),0) as Unit, Emp.PresentRate as PresentRate, " & _
						" Emp.PresentGrade as PresentGrade " & _
						" from attendance_gad Att_Gad,Employee Emp" & _
						" where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
						" and emp.DATETERMINATE is null and Att_Gad.Attendance='KJ'" & _
						" and att_gad.Jobcode='6502101'" & _
						" and att_gad.EmployeeCode = '" & sEmployeeCode  & "'" & _
						" and att_gad.LOCATIONTYPE='FF'" & _
						" and Att_Gad.MANDAYS>0" & _
						" and Att_Gad.UNIT>0" & _
						" and Att_Gad.rate=0" & _
						" and Att_Gad.FDDATE = to_date('" & dFdDate & "','MM/DD/YYYY') " & _  
						" and to_number(to_Char(Att_Gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
						" and to_number(to_Char(Att_Gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
						" group by att_gad.FDDATE,att_gad.LOCATIONCODE,att_gad.EMPLOYEECODE, Emp.PresentRate, Emp.PresentGrade"
				  
				Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
				rsEmpAllDed.CursorType = 3
				rsEmpAllDed.CursorLocation = 3
				response.Write "Calculate Premi/Potongan Panen..." & "<br>" 
				response.Flush()
				rsEmpAllDed.Open qReceive, connstr
				if not rsEmpAllDed.eof then
				   Do While Not rsEmpAllDed.eof
					  'Cari Potongan Kalau ada
					  
							  
					  'Cari sensus Blok
					  qBA=" SELECT nvl(sum(HARVESTINGBASED_ABW),0)  as n_BjrBlok" & _
						  " FROM FIELD WHERE fccode = '" & rsEmpAllDed("LocationCode") & "'"  & _
						  " and fcBA='" & Session("fcBA") & "'"
					  Set rsBA = Server.CreateObject("ADODB.Recordset")
					  
					  response.Flush()
					  rsBA.Open qBA, connstr,3,3
					  if not rsBA.eof then
						 n_BjrBlok =cdbl(rsBA("n_BjrBlok"))
					  else
						 n_BjrBlok =0
						 Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Qty Basis Panen A24 is Null "
						 err.number=10001
					  end if
					  rsBA.close
					  set rsBA = Nothing
					  set qBA = Nothing



					  ' Cari Basis
							  qBA=" SELECT nvl(sum(BASIS_QTY),0) as nBasisQty,nvl(sum(BASIS_Rate),0) as nRpDapatBasis,nvl(sum(EXTRA_BASIS_RATE),0) as nRpExtraBasis " & _
								  " FROM OPHARVRATE WHERE " & n_BjrBlok & " >= ABW_MIN  and " & n_BjrBlok & " <= ABW_MAX "  & _
						  " and fcBA='" & Session("fcBA") & "'"
					  Set rsBA = Server.CreateObject("ADODB.Recordset")
					  'response.Write(qba)
					  response.Flush()
					  rsBA.Open qBA, connstr,3,3
					  if not rsBA.eof then
								 nQtyBasis=cdbl(rsBA("nBasisQty"))
						 nRpExtraBasis=cdbl(rsBA("nRpExtraBasis"))
								 nRpDapatBasis=cdbl(rsBA("nRpDapatBasis"))
					  else
								 nQtyBasis=0
						 nRpExtraBasis=0
						 nRpDapatBasis=0

						 Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Qty Basis Panen A24 is Null "
						 err.number=10001
					  end if
					  rsBA.close
					  set rsBA = Nothing
					  set qBA = Nothing


							  


					  qBA="select nvl(sum(UNDER_RIPE),0) as UNDER_RIPE,nvl(sum(OVER_RIPE),0) as OVER_RIPE " & _
						  ",nvl(sum(UNHARVEST_FFB),0) as UNHARVEST_FFB,nvl(sum(UNCOLLECT_LF_CIRCLE),0) as UNCOLLECT_LF_CIRCLE " & _
						  ",nvl(sum(UNARRANGE_FFB),0) as UNARRANGE_FFB,nvl(sum(UNPRUNE_FROND),0) as UNPRUNE_FROND " & _
						  ",nvl(sum(QE_1),0) as QE_1,nvl(sum(QE_2),0) as QE_2 " & _
						  " from HARVESTINGQUALITY where EMPCODE='" & rsEmpAllDed("EMPLOYEECODE") & "'" & _
						  " and FDDATE =to_Date('" & rsEmpAllDed("FDDATE") & "','MM/DD/YYYY')" & _
						  " and fcBA='" & Session("fcBA") & "'"
					  Set rsBA = Server.CreateObject("ADODB.Recordset")
					  'response.Write(qba)
					  response.Flush()
					  rsBA.Open qBA, connstr,3,3
					  if not rsBA.eof then
						 nQtyBuahMentah=cdbl(rsBA("UNDER_RIPE"))
						 nQtyTerlaluMasak=cdbl(rsBA("OVER_RIPE"))
								 nQtyBuahtdkdipanen=cdbl(rsBA("UNHARVEST_FFB"))
							 nQtyPelepahtdkdipotong=cdbl(rsBA("UNPRUNE_FROND"))
							 nQtyPelepahtdkdisusun=cdbl(rsBA("QE_1"))
							 nQtyBrondolandipiringan=cdbl(rsBA("UNCOLLECT_LF_CIRCLE"))
							 nQtybuahTdkdiTPH=cdbl(rsBA("UNARRANGE_FFB"))
							 nQtyPulangtdkMelapor=cdbl(rsBA("QE_2"))		    
					  else
						 nQtyBuahMentah=0
						 nQtyTerlaluMasak=0     
								 nQtyBuahtdkdipanen=0
							 nQtyPelepahtdkdipotong=0
							 nQtyPelepahtdkdisusun=0
							 nQtyBrondolandipiringan=0
							 nQtybuahTdkdiTPH=0
							 nQtyPulangtdkMelapor=0

						 Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Qty Basis Panen A24 is Null "
						 err.number=10001
					  end if
					  rsBA.close
					  set rsBA = Nothing
					  set qBA = Nothing
					  
					  nQtyPremiTandan=cdbl(rsEmpAllDed("Unit"))-(nQtyBuahMentah+nQtyTerlaluMasak)

					  if nQtyPremiTandan >= nQtyBasis then
								 'siap borong
								 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A27" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nRpDapatBasis & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)

						 nRpPremiTandan = (nQtyPremiTandan-nQtyBasis) * nRpExtraBasis
					  else
						 ' Insert Deduction untuk D19 POTONGAN TIDAK DAPAT BASIS
						   
							 nRpDapatBasis = (nQtyBasis - nQtyPremiTandan)/nQtyBasis * cdbl(rsEmpAllDed("PresentRate"))
							 
							 if rsEmpAllDed("PresentGrade") = "KHT" or rsEmpAllDed("PresentGrade") = "KBT" then 
								 nRpDapatBasis = nRpDapatBasis / 30
							 end if
							 
							 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
										"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D19" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nRpDapatBasis & _
										",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
										Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
							 connstr.execute(qAllowance)
							 nRpPremiTandan=0
					  end if 

							  'RESPONSE.WRITE "....." & "pREMI = " & nRpPremiTandan 
					  
					  if nRpPremiTandan>0 then
							 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A24" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremiTandan & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if

					  if nQtyBuahMentah*nRateBuahMentah>0 then
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D09" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nQtyBuahMentah*nRateBuahMentah & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if

					  if nQtyTerlaluMasak*nRateBuahMasak>0 then
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D10" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nQtyTerlaluMasak*nRateBuahMasak & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if


					  if nQtyBuahtdkdipanen*nBuahtdkdipanen>0 then
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D12" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nQtyBuahtdkdipanen*nBuahtdkdipanen & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if

							 if nQtyPelepahtdkdipotong*nPelepahtdkdipotong>0 then
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D13" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nQtyPelepahtdkdipotong*nPelepahtdkdipotong & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if


					  if nQtyPelepahtdkdisusun*nPelepahtdkdisusun>0 then
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D14" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nQtyPelepahtdkdisusun*nPelepahtdkdisusun & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if

							 if nQtyBrondolandipiringan*nBrondolandipiringan>0 then
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D15" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nQtyBrondolandipiringan*nBrondolandipiringan & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if

							 if nQtybuahTdkdiTPH*nbuahTdkdiTPH>0 then
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D16" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nQtybuahTdkdiTPH*nbuahTdkdiTPH & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if


							 if nQtyPulangtdkMelapor*nPulangtdkMelapor>0 then
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
									"('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "D17" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nQtyPulangtdkMelapor*nPulangtdkMelapor & _
									",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
						 connstr.execute(qAllowance)
					  end if
							

					  rsEmpAllDed.MoveNext
				   loop
				End if
				rsEmpAllDed.Close
				set rsEmpAllDed = Nothing
			else
			    ' Start for Double/Multi Block
				
				' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
				
				qTotalUnit = 
				" select  employeecode, sum(unit) as totalUnit " & _
				" from ( " & _ 
					"select att_gad.FDDATE,att_gad.LOCATIONCODE,att_gad.EMPLOYEECODE,nvl(sum(att_gad.unit),0) as Unit, Emp.PresentRate as PresentRate, " & _
						" Emp.PresentGrade as PresentGrade " & _
						" from attendance_gad Att_Gad,Employee Emp" & _
						" where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
						" and emp.DATETERMINATE is null and Att_Gad.Attendance='KJ'" & _
						" and att_gad.Jobcode='6502101'" & _
						" and att_gad.EmployeeCode = '" & sEmployeeCode & "'" & _
						" and att_gad.LOCATIONTYPE='FF'" & _
						" and Att_Gad.MANDAYS>0" & _
						" and Att_Gad.UNIT>0" & _
						" and Att_Gad.rate=0" & _
						" and Att_Gad.FDDATE = to_date('" & dFdDate & "','MM/DD/YYYY') " & _  
						" and to_number(to_Char(Att_Gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
						" and to_number(to_Char(Att_Gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
						" group by att_gad.FDDATE,att_gad.LOCATIONCODE,att_gad.EMPLOYEECODE, Emp.PresentRate, Emp.PresentGrade " & _
				" ) group by employeecode "
				
				Set rsTotalUnit = Server.CreateObject("ADODB.Recordset")
				rsTotalUnit.CursorType = 3
				rsTotalUnit.CursorLocation = 3
				response.Write "Total Unit for All Block ..." & "<br>" 
				response.Flush()
				rsTotalUnit.Open qTotalUnit, connstr
				
				if not rsTotalUnit.eof then
					nTotalUnit = rsTotalUnit("totalUnit")
				end if

				
				
				
				' Decide what is the first Row
				
				qFirstRowBlock = 
				" select fddate, locationcode, employeecode, unit, presentrate, presentgrade " & _
				" from ( " & _ 
					"select att_gad.FDDATE,att_gad.LOCATIONCODE,att_gad.EMPLOYEECODE,nvl(sum(att_gad.unit),0) as Unit, Emp.PresentRate as PresentRate, " & _
						" Emp.PresentGrade as PresentGrade " & _
						" from attendance_gad Att_Gad,Employee Emp" & _
						" where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
						" and emp.DATETERMINATE is null and Att_Gad.Attendance='KJ'" & _
						" and att_gad.Jobcode='6502101'" & _
						" and att_gad.EmployeeCode = '" & sEmployeeCode & "'" & _
						" and att_gad.LOCATIONTYPE='FF'" & _
						" and Att_Gad.MANDAYS>0" & _
						" and Att_Gad.UNIT>0" & _
						" and Att_Gad.rate=0" & _
						" and Att_Gad.FDDATE = to_date('" & dFdDate & "','MM/DD/YYYY') " & _  
						" and to_number(to_Char(Att_Gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
						" and to_number(to_Char(Att_Gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
						" group by att_gad.FDDATE,att_gad.LOCATIONCODE,att_gad.EMPLOYEECODE, Emp.PresentRate, Emp.PresentGrade " & _
				" ) where rownum = 1 " 	
				
				Set rsFirstRowBlock = Server.CreateObject("ADODB.Recordset")
				rsFirstRowBlock.CursorType = 3
				rsFirstRowBlock.CursorLocation = 3
				response.Write "Decide First Block ..." & "<br>" 
				response.Flush()
				rsFirstRowBlock.Open qFirstRowBlock, connstr
				
				if not rsFirstRowBlock.eof then
					sFirstLocationCode = rsFirstRowBlock("locationcode")
					dFirstFdDate = cdate(rsFirstRowBlock("fddate"))
					nFirstUnit =  cdbl(rsFirstRowBlock("unit"))
					nPresentRate = cdbl(rsFirstRowBlock("presentrate"))
				end if

				
				'Cari sensus Blok
				qBA=" SELECT nvl(sum(HARVESTINGBASED_ABW),0)  as n_BjrBlok" & _
					" FROM FIELD WHERE fccode = '" & sFirstLocationCode & "'"  & _
					" and fcBA='" & Session("fcBA") & "'"
				Set rsBA = Server.CreateObject("ADODB.Recordset")
					  
				response.Flush()
				rsBA.Open qBA, connstr,3,3
				if not rsBA.eof then
					n_BjrBlok =cdbl(rsBA("n_BjrBlok"))
				else
					n_BjrBlok =0
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Qty Basis Panen A24 is Null "
					err.number=10001
				end if
				rsBA.close
				set rsBA = Nothing
				set qBA = Nothing
				
				
				' Cari Basis
				qBA=" SELECT nvl(sum(BASIS_QTY),0) as nBasisQty,nvl(sum(BASIS_Rate),0) as nRpDapatBasis,nvl(sum(EXTRA_BASIS_RATE),0) as nRpExtraBasis " & _
					" FROM OPHARVRATE WHERE " & n_BjrBlok & " >= ABW_MIN  and " & n_BjrBlok & " <= ABW_MAX "  & _
					" and fcBA='" & Session("fcBA") & "'"

				Set rsBA = Server.CreateObject("ADODB.Recordset")
				'response.Write(qba)
				response.Flush()
				rsBA.Open qBA, connstr,3,3
				if not rsBA.eof then
					nFirstQtyBasis=cdbl(rsBA("nBasisQty"))
					nFirstRpExtraBasis=cdbl(rsBA("nRpExtraBasis"))
					nFirstRpDapatBasis=cdbl(rsBA("nRpDapatBasis"))
				else
					nFirstQtyBasis=0
					nFirstRpExtraBasis=0
					nFirstRpDapatBasis=0

					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Qty Basis Panen A24 is Null "
					err.number=10001
				end if
				rsBA.close
				set rsBA = Nothing
				set qBA = Nothing


				
				qReceive="select att_gad.FDDATE,att_gad.LOCATIONCODE,att_gad.EMPLOYEECODE,nvl(sum(att_gad.unit),0) as Unit, Emp.PresentRate as PresentRate, " & _
						" Emp.PresentGrade as PresentGrade " & _
						" from attendance_gad Att_Gad,Employee Emp" & _
						" where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
						" and emp.DATETERMINATE is null and Att_Gad.Attendance='KJ'" & _
						" and att_gad.Jobcode='6502101'" & _
						" and att_gad.EmployeeCode = '" & sEmployeeCode & "'" & _
						" and att_gad.LOCATIONTYPE='FF'" & _
						" and Att_Gad.MANDAYS>0" & _
						" and Att_Gad.UNIT>0" & _
						" and Att_Gad.rate=0" & _
						" and Att_Gad.FDDATE = to_date('" & dFdDate & "','MM/DD/YYYY') " & _  
						" and to_number(to_Char(Att_Gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
						" and to_number(to_Char(Att_Gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
						" group by att_gad.FDDATE,att_gad.LOCATIONCODE,att_gad.EMPLOYEECODE, Emp.PresentRate, Emp.PresentGrade"
				  
				Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
				rsEmpAllDed.CursorType = 3
				rsEmpAllDed.CursorLocation = 3
				response.Write "Calculate Premi/Potongan Panen..." & "<br>" 
				response.Flush()
				rsEmpAllDed.Open qReceive, connstr
				
				' Get the TotalConversion for All Block
				nTotalUnitConversion = 0
				if not rsEmpAllDed.eof then
				    Do While Not rsEmpAllDed.eof
						
						'Cari sensus Blok
						qBA=" SELECT nvl(sum(HARVESTINGBASED_ABW),0)  as n_BjrBlok" & _
							" FROM FIELD WHERE fccode = '" & rsEmpAllDed("LocationCode") & "'"  & _
							" and fcBA='" & Session("fcBA") & "'"
						Set rsBA = Server.CreateObject("ADODB.Recordset")
							  
						response.Flush()
						rsBA.Open qBA, connstr,3,3
						if not rsBA.eof then
							n_BjrBlok =cdbl(rsBA("n_BjrBlok"))
						else
							n_BjrBlok =0
							Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Qty Basis Panen A24 is Null "
							err.number=10001
						end if
						rsBA.close
						set rsBA = Nothing
						set qBA = Nothing
						
						
						' Cari Basis
						qBA=" SELECT nvl(sum(BASIS_QTY),0) as nBasisQty,nvl(sum(BASIS_Rate),0) as nRpDapatBasis,nvl(sum(EXTRA_BASIS_RATE),0) as nRpExtraBasis " & _
							" FROM OPHARVRATE WHERE " & n_BjrBlok & " >= ABW_MIN  and " & n_BjrBlok & " <= ABW_MAX "  & _
							" and fcBA='" & Session("fcBA") & "'"

						Set rsBA = Server.CreateObject("ADODB.Recordset")
						'response.Write(qba)
						response.Flush()
						rsBA.Open qBA, connstr,3,3
						if not rsBA.eof then
							nQtyBasis=cdbl(rsBA("nBasisQty"))
							nRpExtraBasis=cdbl(rsBA("nRpExtraBasis"))
							nRpDapatBasis=cdbl(rsBA("nRpDapatBasis"))
						else
							nQtyBasis=0
							nRpExtraBasis=0
							nRpDapatBasis=0

							Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Qty Basis Panen A24 is Null "
							err.number=10001
						end if
						rsBA.close
						set rsBA = Nothing
						set qBA = Nothing			

						nTotalUnitConversion = nTotalUnitConversion + ( cdbl(rsEmpAllDed("Unit")) / nQtyBasis * nFirstQtyBasis )

					    rsEmpAllDed.MoveNext
				    loop
				End if


				qBA="select nvl(sum(UNDER_RIPE),0) as UNDER_RIPE,nvl(sum(OVER_RIPE),0) as OVER_RIPE " & _
					",nvl(sum(UNHARVEST_FFB),0) as UNHARVEST_FFB,nvl(sum(UNCOLLECT_LF_CIRCLE),0) as UNCOLLECT_LF_CIRCLE " & _
					",nvl(sum(UNARRANGE_FFB),0) as UNARRANGE_FFB,nvl(sum(UNPRUNE_FROND),0) as UNPRUNE_FROND " & _
					",nvl(sum(QE_1),0) as QE_1,nvl(sum(QE_2),0) as QE_2 " & _
					" from HARVESTINGQUALITY where EMPCODE='" & sEmployeeCode & "'" & _
					" and FDDATE = to_date('" & dFdDate & "','MM/DD/YYYY') " & _  
					" and fcBA='" & Session("fcBA") & "'"
				Set rsBA = Server.CreateObject("ADODB.Recordset")
				'response.Write(qba)
				response.Flush()
				rsBA.Open qBA, connstr,3,3
				if not rsBA.eof then
					nQtyBuahMentah=cdbl(rsBA("UNDER_RIPE"))
					nQtyTerlaluMasak=cdbl(rsBA("OVER_RIPE"))
					nQtyBuahtdkdipanen=cdbl(rsBA("UNHARVEST_FFB"))
					nQtyPelepahtdkdipotong=cdbl(rsBA("UNPRUNE_FROND"))
					nQtyPelepahtdkdisusun=cdbl(rsBA("QE_1"))
					nQtyBrondolandipiringan=cdbl(rsBA("UNCOLLECT_LF_CIRCLE"))
					nQtybuahTdkdiTPH=cdbl(rsBA("UNARRANGE_FFB"))
					nQtyPulangtdkMelapor=cdbl(rsBA("QE_2"))		    
				else
					nQtyBuahMentah=0
					nQtyTerlaluMasak=0     
					nQtyBuahtdkdipanen=0
					nQtyPelepahtdkdipotong=0
					nQtyPelepahtdkdisusun=0
					nQtyBrondolandipiringan=0
					nQtybuahTdkdiTPH=0
					nQtyPulangtdkMelapor=0

					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY10001 Qty Basis Panen A24 is Null "
					err.number=10001
				end if
				rsBA.close
				set rsBA = Nothing
				set qBA = Nothing
					  
				nQtyPremiTandan=nTotalUnitConversion-(nQtyBuahMentah+nQtyTerlaluMasak)

				if nQtyPremiTandan >= nQtyBasis then
						 'siap borong
					
						 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
							"('" & sEmployeeCode & "','" & "A27" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nFirstRpDapatBasis & _
							",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				            connstr.execute(qAllowance)

				     nRpPremiTandan = (nQtyPremiTandan-nFirstQtyBasis) * nFirstRpExtraBasis
				else
				 ' Insert Deduction untuk D19 POTONGAN TIDAK DAPAT BASIS
				   
					 nFirstRpDapatBasis = (nFirstQtyBasis - nQtyPremiTandan)/nFirstQtyBasis * nPresentRate
					 
					 if rsEmpAllDed("PresentGrade") = "KHT" or rsEmpAllDed("PresentGrade") = "KBT" then 
						 nFirstRpDapatBasis = nFirstRpDapatBasis / 30
					 end if
					 
					 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
								"('" & sEmployeeCode & "','" & "D19" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nFirstRpDapatBasis & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					 connstr.execute(qAllowance)
					 nRpPremiTandan=0
				end if 

							  'RESPONSE.WRITE "....." & "pREMI = " & nRpPremiTandan 
					  
				if nRpPremiTandan>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
							"('" & sEmployeeCode & "','" & "A24" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nRpPremiTandan & _
							",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				    connstr.execute(qAllowance)
				end if

				if nQtyBuahMentah*nRateBuahMentah>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
								"('" & sEmployeeCode & "','" & "D09" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nQtyBuahMentah*nRateBuahMentah & _
								",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					 connstr.execute(qAllowance)
				end if

				if nQtyTerlaluMasak*nRateBuahMasak>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					"('" & sEmployeeCode & "','" & "D10" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nQtyTerlaluMasak*nRateBuahMasak & _
					",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					connstr.execute(qAllowance)
				end if


				if nQtyBuahtdkdipanen*nBuahtdkdipanen>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					"('" & sEmployeeCode & "','" & "D12" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nQtyBuahtdkdipanen*nBuahtdkdipanen & _
					",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					connstr.execute(qAllowance)
				end if

				if nQtyPelepahtdkdipotong*nPelepahtdkdipotong>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					"('" & sEmployeeCode & "','" & "D13" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nQtyPelepahtdkdipotong*nPelepahtdkdipotong & _
					",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					connstr.execute(qAllowance)
				end if


				if nQtyPelepahtdkdisusun*nPelepahtdkdisusun>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					"('" & sEmployeeCode & "','" & "D14" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nQtyPelepahtdkdisusun*nPelepahtdkdisusun & _
					",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					connstr.execute(qAllowance)
				end if

				if nQtyBrondolandipiringan*nBrondolandipiringan>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					"('" & sEmployeeCode & "','" & "D15" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nQtyBrondolandipiringan*nBrondolandipiringan & _
					",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					connstr.execute(qAllowance)
				end if

				if nQtybuahTdkdiTPH*nbuahTdkdiTPH>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					"('" & sEmployeeCode & "','" & "D16" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nQtybuahTdkdiTPH*nbuahTdkdiTPH & _
					",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					connstr.execute(qAllowance)
				end if


				if nQtyPulangtdkMelapor*nPulangtdkMelapor>0 then
					qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					"('" & sEmployeeCode & "','" & "D17" & "'," &  " to_date('" & dFdDate & "','mm/dd/yyyy') " & "," & nQtyPulangtdkMelapor*nPulangtdkMelapor & _
					",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
					connstr.execute(qAllowance)
				end if
							
				' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX				
			end if
			rsEmpAllDedDecision.MoveNext
		loop
	end if	
	rsEmpAllDedDecision.Close
	set rsEmpAllDedDecision = Nothing		
%>