<%
	'Premi Borongan
	'1. Code Allowance Panen = A45
	'   unit x Rate

	
	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' First, take the Original Table of RECEIVENOTE.

	tempProcessCode = "PAYROLL"
	tempProcessNameCode = "Payroll_Borongan.asp"
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
	
	
	qAllowdedBrondolan = "select allowded_api.GetIncludedPayrollProcessing('A26', '" & Session("fcBA") & "') tempAllowdedBrondolan from dual "
	
	Set rsqAllowdedBrondolan = Server.CreateObject("ADODB.Recordset")
	rsqAllowdedBrondolan.CursorType = 3
	rsqAllowdedBrondolan.CursorLocation = 3
	rsqAllowdedBrondolan.Open qAllowdedBrondolan, connstr
	
	if not rsqAllowdedBrondolan.eof then
		tempData = rsqAllowdedBrondolan("tempAllowdedBrondolan")
	else
		tempData = "FALSE"
	end if
	rsqAllowdedBrondolan.close
	set rsqAllowdedBrondolan = Nothing
	
	
	
    qReceive="select att_gad.fcba, att_gad.jobcode, att_gad.FDDATE,att_gad.EMPLOYEECODE,att_gad.locationCODE,((sum(nvl(att_gad.unit,0) * nvl(att_gad.Rate,0)))+(sum(nvl(att_gad.kg_brondolan,0) * nvl(att_gad.Rate,0)))) as nRpBorongan" & _
			 " from attendance_gad Att_Gad,Employee Emp" & _
			 " where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
			 " and emp.DATETERMINATE is null and Att_Gad.Attendance in ('KB')" & _
			 " and Att_Gad.MANDAYS=0" & _
             " and to_number(to_Char(Att_Gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			 " and to_number(to_Char(Att_Gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
			 " group by att_gad.fcba, att_gad.jobcode, att_gad.FDDATE,att_gad.EMPLOYEECODE,att_gad.locationCODE"
			
	if tempData = "TRUE" then 
		qReceive = " select FDDATE, EMPLOYEECODE, locationCODE, " & _
		           " decode ( process_posting_api.GetPostingValue('PAYROLL', 'A26', 'JOB', jobcode, fcba),  'TRUE', 'A26','A45') allowdedcode, " & _
		           " sum(nRpBorongan) nRpBorongan from (" & qReceive & ") " & _
				   " group by FDDATE, EMPLOYEECODE, locationCODE, decode ( process_posting_api.GetPostingValue('PAYROLL', 'A26', 'JOB', jobcode, fcba),  'TRUE', 'A26','A45') "	
    else
		qReceive = " select FDDATE, EMPLOYEECODE, locationCODE, " & _
		           " 'A45' allowdedcode, " & _
		           " sum(nRpBorongan) nRpBorongan from (" & qReceive & ") " & _
				   " group by FDDATE, EMPLOYEECODE, locationCODE "	
	end if
			
			
	  
	'response.Write " sql " & qReceive
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate Premi Borongan ..." & "<br>" 
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
	   Do While Not rsEmpAllDed.eof
         
			maxLineNo = maxLineNo + 1
	        qAllowance="Insert Into  EmpAllded (LINENO, EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) values  " & _
			            "(" & maxLineNo & ",'" & rsEmpAllDed("EMPLOYEECODE") & "','" & rsEmpAllDed("allowdedcode") & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & rsEmpAllDed("nRpBorongan") & _
			            ",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			            Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',sysdate, to_char(SYSDATE, 'HH24:MI'))"
		    'response.Write "<br/>qAllowance : " & qAllowance
			connstr.execute(qAllowance)
			 
			' Put the Historical Detail Data in EmpallDed_Hist_Tab 
			' ** Remarks History By Iwa 08/01/2019 **
						
			'qReceiveDetail =" select " & _
			'						 " att_gad.FDDATE,  " & _
			'						 " att_gad.EMPLOYEECODE,  " & _
			'						 " att_gad.JOBCODE,  " & _
			'						 " att_gad.LOCATIONTYPE,  " & _
			'						 " att_gad.LOCATIONCODE,  " & _
			'						 " nvl(att_gad.unit,0) unit,  " & _
			'						 " nvl(att_gad.Rate,0) rate,  " & _
			'						 " (nvl(att_gad.unit,0) * nvl(att_gad.Rate,0)) + (nvl(att_gad.KG_BRONDOLAN,0)*nvl(att_gad.Rate,0)) AMOUNT" & _
			'				" from attendance_gad Att_Gad,Employee Emp" & _
			'				" where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
			'				" and emp.DATETERMINATE is null and Att_Gad.Attendance in ('KB')" & _
			'				" and Att_Gad.MANDAYS=0" & _
			'				" and to_number(to_Char(Att_Gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			'				" and to_number(to_Char(Att_Gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
			'			    " and att_gad.FDDATE =  to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & _
			'				" and att_gad.EMPLOYEECODE = '" & rsEmpAllDed("EMPLOYEECODE") & "' " & _
			'				" and att_gad.locationCODE = '" & rsEmpAllDed ("LOCATIONCODE") & "' "
			
			 
			'Set rsReceive = Server.CreateObject("ADODB.Recordset")
			'rsReceive.CursorType = 3
			'rsReceive.CursorLocation = 3

			'rsReceive.Open qReceiveDetail, connstr
			'if not rsReceive.eof then
			'	detailLineNo = 0
			'	Do While Not rsReceive.eof	

					'	Insert the History of EmpallDed Data 
					' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
					
			'		txtRefStringValues_Receive = rsReceive("FDDATE") & txtRefStringSeparator_Receive & _
			'									 rsReceive("EMPLOYEECODE") & txtRefStringSeparator_Receive & _
			'									 rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
			'									 rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
			'									 rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
			'									 Session("fcBA")
					
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
			'											cdbl(rsReceive("UNIT")) & "," & _
			'											cdbl(rsReceive("RATE")) & "," & _
			'											cdbl(rsReceive("AMOUNT")) & "," & _ 
			'											" '" & Session("CodeName") & "'," & _
			'											" '" & Session("CodeName") & "'," & _
			'											" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
			'											" sysdate, " & _
			'											" to_char(sysdate, 'HH24:MI') " & _
			'											" ) "					
			'		connstr.execute(qAllowanceHistory)
			'		rsReceive.MoveNext
			'	loop
			'end if
			'rsReceive.Close
			'set rsReceive = Nothing	
				 

	        rsEmpAllDed.MoveNext
	   loop
	End if
	rsEmpAllDed.Close
	set rsEmpAllDed = Nothing
%>