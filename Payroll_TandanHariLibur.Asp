
<%
	'Amandemen Premi Tandan Hari Libur
	'1. Code Allowance Panen = A25
	'   Rate1 = Rupiah Per Tandan
	'   Rate2 = Rupiah Pengali

	'Dim nRatePerTandan,nRatePengali
    ' Cari Basis
	qBA="select nvl(sum(Rate_1),0) as Rate_1,nvl(sum(Rate_2),0) as Rate_2 from ALLOWDED where FCCODE='A25' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	'response.Write(qba)
	response.Flush()
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   nRatePerTandan=cdbl(rsBA("Rate_1"))
	   nRatePengali=cdbl(rsBA("Rate_2"))
	else
	   nRatePerTandan=0
	   nRatePengali=0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY11001 Rate Premi Tandan Work Holiday is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing
	
	qReceive="select att_gad.FDDATE,att_gad.EMPLOYEECODE,nvl(sum(att_gad.unit),0) as Unit" & _
			" from attendance_gad Att_Gad,Employee Emp" & _
			" where (att_gad.EMPLOYEECODE = emp.FCCODE and att_gad.FCBA = emp.FCBA)" & _
			" and emp.DATETERMINATE is null and Att_Gad.Attendance in ('WH','UN','WS')" & _
			" and att_gad.Jobcode='312001'" & _
			" and Att_Gad.MANDAYS>0" & _
			" and Att_Gad.UNIT>0" & _
            " and to_number(to_Char(Att_Gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			" and to_number(to_Char(Att_Gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and Att_Gad.fcBA='" & Session("fcBA") & "'" & _
			" group by att_gad.FDDATE,att_gad.EMPLOYEECODE"
	  
    Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
	rsEmpAllDed.CursorType = 3
	rsEmpAllDed.CursorLocation = 3
	response.Write "Calculate Premi Tandan Hari Minggu ..." & "<br>" 
	response.Flush()
	rsEmpAllDed.Open qReceive, connstr
	if not rsEmpAllDed.eof then
	   Do While Not rsEmpAllDed.eof
         
		  nRpPremiTandan=cdbl(rsEmpAllDed("Unit"))*nRatePerTandan*nRatePengali
		  if nRpPremiTandan > 0 then
	         qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT,FIXEDALLDED, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
			            "('" & rsEmpAllDed("EMPLOYEECODE") & "','" & "A25" & "'," &  " to_date('" & rsEmpAllDed("FDDATE") & "','mm/dd/yyyy') " & "," & nRpPremiTandan & _
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