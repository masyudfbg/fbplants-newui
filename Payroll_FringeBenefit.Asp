
<%
	qReceive="select Att_Gad.employeecode as FCCODE,nvl(emp.PRESENTRATE,0) as PRESENTRATE ,Att_Gad.attendance as Attendance,nvl(count(*),0) as nJlhHari" & _
			" from attendance_gad Att_Gad,Employee Emp" & _
			" where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
			" and DATETERMINATE is null and Att_Gad.Attendance in ('CT','S1','S2','H1','H2','P1','P2','P3','HL') " & _
			" and EMP.PRESENTGRADE ='KHT' " & _
            " and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
			" group by Att_Gad.employeecode,emp.PRESENTRATE,Att_Gad.attendance"
	  
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate Fringe Benefit..." & "<br>"
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
	   Do While Not rsEmpAllDed.eof
          nPresentRate=cdbl(rsEmpAllDed("PRESENTRATE")/30)
		  nDays=cdbl(rsEmpAllDed("nJlhHari"))
		  nAmount=nPresentRate*nDays
		  if rsEmpAllDed("Attendance")="CT" Then
		     v_AllowanceCode="A03"
		  elseif rsEmpAllDed("Attendance")="S1" then
		     v_AllowanceCode="A04"
		  elseif rsEmpAllDed("Attendance")="S2" then
		     v_AllowanceCode="A05"
		  elseif rsEmpAllDed("Attendance")="H1" then
		     v_AllowanceCode="A06"
		  elseif rsEmpAllDed("Attendance")="H2" then
		     v_AllowanceCode="A07"
		  elseif rsEmpAllDed("Attendance")="P1" then
		     v_AllowanceCode="A08"
		  elseif rsEmpAllDed("Attendance")="P2" then
		     v_AllowanceCode="A09"
		  elseif rsEmpAllDed("Attendance")="P3" then
		     v_AllowanceCode="A10"
		  elseif rsEmpAllDed("Attendance")="HL" then
		     v_AllowanceCode="A11"
		  end if
	      qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
			          "('" & rsEmpAllDed("FCCODE") & "','" & v_AllowanceCode & "'," &  " to_date('" & Session("v_Employee") & "','mm/dd/yyyy') " & "," & nAmount & _
			          ",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			          Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		  'response.Write(qAllowance) & "<br>"		  
	      connstr.execute(qAllowance)
	      rsEmpAllDed.MoveNext
	   loop
	End if
	rsEmpAllDed.Close
	set rsEmpAllDed = Nothing
%>