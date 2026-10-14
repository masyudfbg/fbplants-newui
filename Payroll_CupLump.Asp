
<%
	'Amandemen Premi Cup Lump
	'1. Code Allowance Panen = A28
	'   Rate1 = Rupiah Per Kg
    'Jobcode Cup Lump = 312004
	
	'Dim nRatePerKg,nRatePengali
    ' Cari Basis
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='A28' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nRatePerKg=cdbl(rsBA("Rate_1"))
	else
	   nRatePerKg=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY13001 Rate Premi Cup Lump is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
	
	qReceive="select att_gad.FDDATE,att_gad.EMPLOYEECODE,nvl(sum(att_gad.unit),0) as Unit" & _
			" from attendance_gad Att_Gad,Employee Emp" & _
			" where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
			" and emp.DATETERMINATE is null and Att_Gad.Attendance in ('KJ')" & _
			" and att_gad.Jobcode='312004'" & _
			" and Att_Gad.MANDAYS>0" & _
			" and Att_Gad.UNIT>0" & _
            " and to_number(to_Char(Att_Gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			" and to_number(to_Char(Att_Gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
			" group by att_gad.FDDATE,att_gad.EMPLOYEECODE"
	  
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate Premi Cup Lump ..." & "<br>" 
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
	   Do While Not rsEmpAllDed.eof
         
		  nRpPremiCupLump=cdbl(rsEmpAllDed("Unit"))*nRatePerKg
		  if nRpPremiCupLump > 0 then
	         qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
			            "('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A28" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremiCupLump & _
			            ",0,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			            Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			 connstr.execute(qAllowance)
		  end if

	      rsEmpAllDed.MoveNext
	   loop
	End if
	rsEmpAllDed.Close
	set rsEmpAllDed = Nothing
%>