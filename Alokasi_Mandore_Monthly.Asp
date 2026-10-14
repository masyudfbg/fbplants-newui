<%
'--------------------- Find Basis by Unit
Set rsBasis = Server.CreateObject("ADODB.Recordset")
rsBasis.CursorType = 3
rsBasis.CursorLocation = 3
nBasis=0


qBasis =  " Select nvl(sum(UNIT),0) as Basis From ATTENDANCE_GAD" & _
		" Where nvl(UNIT,0)<>0   " & _
		" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
		" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
		" and decode( '" & tempActivityAllocSupervisor & "','ALL',jobcode, '" & tempActivityAllocSupervisor & "') = jobcode " & _ 		
        " and  INSTR (   SUPERVISION_1||'-'||SUPERVISION_2||'-'||SUPERVISION_3||'-'||SUPERVISION_4, '" & rsEmployee("FcCode") & "') <> 0 " & _
		" and fcBA='" & Session("fcBA") & "' " & _
		" and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' "

	
rsBasis.Open qBasis, connstr
if Not rsBasis.eof then
   if cdbl(rsBasis("Basis"))<> 0 then
      nBasis=cdbl(rsBasis("Basis"))
      tBasisType="UNIT"
   else
      nBasis=0
   end if	  
   rsBasis.Close
   set rsBasis=Nothing
else
   nBasis=0   
end if   
If nBasis=0 Then
   'rsBasis.Close
   'set rsBasis=Nothing
   '------------------------ Find Basis by Mandays
   Set rsBasis = Server.CreateObject("ADODB.Recordset")
   rsBasis.CursorType = 3
   rsBasis.CursorLocation = 3
   
   qBasis =  " Select nvl(sum(Mandays),0) as Basis From ATTENDANCE_GAD" & _
			" Where nvl(Mandays,0)<>0  " & _
			" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
		    " and decode( '" & tempActivityAllocSupervisor & "','ALL',jobcode, '" & tempActivityAllocSupervisor & "') = jobcode " & _ 
            " and  INSTR (   SUPERVISION_1||'-'||SUPERVISION_2||'-'||SUPERVISION_3||'-'||SUPERVISION_4, '" & rsEmployee("FcCode") & "') <> 0 " & _
			" and fcBA='" & Session("fcBA") & "' " & _
		    " and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' "
   
			
   rsBasis.Open qBasis, connstr
   if Not rsBasis.eof then
	  nBasis=cdbl(rsBasis("Basis"))
	  tBasisType="MANDAYS"
   End if
   rsBasis.Close
   set rsBasis=Nothing
End if

if nBasis > 0 Then
	'--------------------- Start GAD
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.CursorType = 3
	rsGAD.CursorLocation = 3
	'--- Save Credit							
	if cdbl(rsEmpAllDed("Amount")) <>0 then
	   maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1						 	
	   qInsertCR=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
				   " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
				   " values ('" & rsEmployee("FcCode") & "','ES','" & "CB" & "','" & _
				   "BS" & "','" & txtControlJob & "'," & 0 & ","& _
				   nxGaji*cdbl(rsEmpAllDed("Amount"))  & _
				   "," & " last_day(to_date('" & Year(Session("v_Employee")) & "'||" & _
												"lpad('" & Month(Session("v_Employee")) & "',2,'0')||'01'" & _
											   ",'YYYYMMDD')),'" & _ 
				   rsAllowded("FCCODE") & "','" & _
				   Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
	  'response.Write(qInsertCR)
	  ' XXX response.Flush()			  
	  connStr.execute(qInsertCR)
	End if
	
	
	qGAD =  " Select LOCATIONTYPE,LOCATIONCODE,FDDATE, JOBCODE, nvl(sum(UNIT),0) as UNIT,sum(Mandays) as Mandays From ATTENDANCE_GAD" & _
			" Where " & _
			" to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
		    " and decode( '" & tempActivityAllocSupervisor & "','ALL',jobcode, '" & tempActivityAllocSupervisor & "') = jobcode " & _ 
            " and  INSTR (   SUPERVISION_1||'-'||SUPERVISION_2||'-'||SUPERVISION_3||'-'||SUPERVISION_4, '" & rsEmployee("FcCode") & "') <> 0 " & _			
			" and fcBA='" & Session("fcBA") &  "' " & _
		    " and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' " & _
			" Group by LOCATIONTYPE,LOCATIONCODE,FDDATE, JOBCODE"
	
	
	rsGAD.Open qGAD, connstr
	nAlokasi=0
	nTotalAlokasi=0
	nSelisih=0
	if Not rsGAD.eof then
	   Do While Not rsGAD.eof
		  if nBasis > 0 Then
		     if tBasisType="UNIT" Then
			    nAlokasi = round( ( cdbl(rsGAD("UNIT"))/cdbl(nBasis)*cdbl(rsEmpAllDed("Amount")) * nxActivity ),0)
			 else
			    nAlokasi = round( ( cdbl(rsGAD("MANDAYS"))/cdbl(nBasis)*cdbl(rsEmpAllDed("Amount")) * nxActivity ),0)
			 end if
			 if nAlokasi <> 0 then
				' Get for Location Type, Location Code and Job Code based on Parameter Table
				Set rsInfoParameter = Server.CreateObject("ADODB.Recordset")
				rsInfoParameter.CursorType = 3
				rsInfoParameter.CursorLocation = 3
				
				qInfoParameter = " select GetInfoParameter ( '" & rsAllowded("FCCODE") &"', '" & _ 
								   rsGAD("LOCATIONTYPE") & "', '" & rsGAD("JOBCODE") &"', 1, '" & Session("fcBA") & "') " & _
								   " as JobCode , " & _
								   " GetInfoParameter ( '" & rsAllowded("FCCODE") &"', '" & _ 
								   rsGAD("LOCATIONTYPE") & "', '" & rsGAD("LOCATIONTYPE") &"', 2, '" & Session("fcBA") & "') " & _
								   " as locationtype ,  " & _
								   " GetInfoParameter ( '" & rsAllowded("FCCODE") &"', '" & _ 
								   rsGAD("LOCATIONTYPE") & "', '" & rsGAD("LOCATIONCODE") &"', 3, '" & Session("fcBA") & "') " & _
								   " as locationcode from dual " 
				
												
				rsInfoParameter.Open qInfoParameter, connstr	
				if Not rsInfoParameter.eof then
					temp_Info_JobCode = rsInfoParameter("jobcode")
					temp_Info_LocationType = rsInfoParameter("locationtype")
					temp_Info_LocationCode = rsInfoParameter("locationcode")
				else
					temp_Info_JobCode = rsGAD("JOBCODE")
					temp_Info_LocationType = rsGAD("LOCATIONTYPE")
					temp_Info_LocationCode = rsGAD("LOCATIONCODE")
				end if
				rsInfoParameter.Close
				Set rsInfoParameter = Nothing

			    maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1						 
			 	qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, Mandays, AMOUNT,TRANSDATE, REFERENCENO,FCENTRY, " & _
						  " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
						  " values ('" & rsEmployee("FcCode") & "','ES','" & temp_Info_LocationType & "','" & _
						  temp_Info_LocationCode & "','" & temp_Info_JobCode & "'," & rsGAD("Unit") & "," &  rsGAD("Mandays") & "," & _
						  nAlokasi  & "," & " to_date('" & rsGAD("FDDATE") & "','mm/dd/yyyy'),'" & rsAllowded("FCCODE") & "','" & _
						  Session("CodeName") & "','" & Session("CodeName") & "','" & _
						  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI')," & maxLineNoEmployeeCB & ")"	
			 
				txtDate=rsGAD("FDDATE")			  
				connStr.execute(qInsertDB)
			  End if	  
		  End if
		  LT    = temp_Info_LocationType
		  LC    = temp_Info_LocationCode
		  If rsEmployee("SECTIONNAME")="BTL" then
			 JOBCODE=rsAllowded("CONTROL_JOB")
		  else
			 JOBCODE=rsGAD("JOBCODE")
		  end if	 
		  FDDATE= rsGAD("FDDATE")
		  
		  rsGAD.MoveNext
		  nTotalAlokasi=nTotalAlokasi+nAlokasi								  
	   Loop
	   'Selisih Pembulatan
	   if nTotalAlokasi<>cdbl(rsEmpAllDed("Amount")) then
		  'response.Write "Dialokasikan " & nTotalAlokasi & "<br>"
		  'response.Write "Payslip " & cdbl(rsEmpAllDed("Amount")) & "<br>"
		  if rsAllowded("ALLDED")="P" or rsAllowded("ALLDED")="A" then
			 nSelisih = cdbl(rsEmpAllDed("Amount"))-nTotalAlokasi
		  else
			 nSelisih = (cdbl(rsEmpAllDed("Amount"))+nTotalAlokasi)*-1
		  end if
		  maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1						 		  
		  qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
						  " FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME,  LINENO)" & _
						  " values ('" & rsEmployee("FcCode") & "','ES','" & LT & "','" & _
						  LC & "','" & temp_Info_JobCode & "'," & 0 & ","& _
						  nSelisih  & "," & " to_date('" & FDDATE & "','mm/dd/yyyy'),'" & rsAllowded("FCCODE") & "','" & _
						  Session("CodeName") & "','" & Session("CodeName") & "','" & _
						  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
		  connStr.execute(qInsertDB)
		  nSelisih=0
	   End if
	   
	   
	End if
	rsGAD.Close
	Set rsGAD = Nothing
	'--------------------- End of GAD
Else '----- Jika Basis Tidak Ada Debet Dari Control Job
	'--- Save Credit							
	'response.Write "OOO UNIT " & "<BR>"
    maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1						 	
	qInsertCR=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
			    " FCEDIT, FCIP,FCBA, LASTUPDATE, LASTTIME,  LINENO )" & _
			    " values ('" & rsEmployee("FcCode") & "','EM','" & "CB" & "','" & _
			    "BS" & "','" & txtControlJob & "'," & 0 & ","& _
			    nxGaji*cdbl(rsEmpAllDed("Amount"))  & _
				"," & " last_day(to_date('" & Year(Session("v_Employee")) & "'||" & _
												"lpad('" & Month(Session("v_Employee")) & "',2,'0')||'01'" & _
											    ",'YYYYMMDD')),'"  & _
				rsAllowded("FCCODE") & "','" & _
			    Session("CodeName") & "','" & Session("CodeName") & "','" & _
			    Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',  sysdate, to_char(sysdate, 'HH24:MI')," & maxLineNoEmployeeCB & ")"
	'response.Write(qInsertCR)
	' XXX response.Flush()			  
	connStr.execute(qInsertCR)
   'Jika Bukan Pekerja dan Bukan Mandore 
	'Maka akan di Charge sebagai Accrued Others
	
    maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1						 	
	qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
				" FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
				" values ('" & rsEmployee("FcCode") & "','EM','" & "CB" & "','" & _
				"BS" & "','" & "1140800" & "'," & 0 & ","& _
				nxActivity * cdbl(rsEmpAllDed("Amount"))  & _
				"," & " last_day(to_date('" & Year(Session("v_Employee")) & "'||" & _
												"lpad('" & Month(Session("v_Employee")) & "',2,'0')||'01'" & _
											    ",'YYYYMMDD')),'" & _
				rsAllowded("FCCODE")+" But No Activity" & "','" & _
				Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
	'response.Write(qInsertDB)
	' XXX response.Flush()			  
	connStr.execute(qInsertDB)
End if '-------------- End if Basis
%>