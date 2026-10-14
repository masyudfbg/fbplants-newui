<%
	'Amandemen Premi Supervisi Panen Mandore1, Mandore, Kerani Panen, Kerani Afd
	'1. Code Allowance Mandore 1 = A39 Rate 1 nya = Rate Mandore 1
	
    ' Cari Rate Mandore 1
	qBA="select nvl(sum(Rate_1),0) as Rate_1 from ALLOWDED where FCCODE='A39' and fcBA='" & Session("fcBA") & "'"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.Open qBA, connstr,3,3
	if not rsBA.eof then
	   Rate_Mandore_1=cdbl(rsBA("Rate_1"))
	else
	   Rate_Mandore_1 = 0
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SACPAY12001 Rate Mandore 1 A39 is Null "
	   err.number=10001
	end if
	rsBA.close
	set rsBA = Nothing
	set qBA = Nothing

	qSupervisi=" select gad.FDDATE,gad.GANGCODE,gad.SUPERVISION_1,nvl(sum(gad.mandays),0) as HK" & _
	         " from ATTENDANCE_GAD gad where  gad.jobcode='6502101'" & _
			 " and supervision_1 is not null" & _
             " and to_number(to_Char(gad.FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			 " and to_number(to_Char(gad.FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and gad.fcBA='" & Session("fcBA") & "'" & _
			 " group by gad.FDDATE,gad.GANGCODE,gad.SUPERVISION_1"
   
    Set rsSupervisi = Server.CreateObject("ADODB.Recordset")
	rsSupervisi.CursorType = 3
	rsSupervisi.CursorLocation = 3
	response.Write "Calculate Premi Mandore 1..."  & "<br>" 
	response.Flush()
	rsSupervisi.Open qSupervisi, connstr
	if not rsSupervisi.eof then
	   Do While Not rsSupervisi.eof
	      nPremiSortasi=0
	      qAnggota=" select gad.FDDATE,gad.GANGCODE,gad.EMPLOYEECODE,gad.SUPERVISION_1,nvl(sum(gad.mandays),0) as HK" & _
				   " from ATTENDANCE_GAD gad where  gad.jobcode='6502101'" & _
				   " and gad.FDDATE=to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') "  &  _
				   " and gad.GANGCODE='" & rsSupervisi("GANGCODE") & "'" & _
				   " and gad.fcBA='" & Session("fcBA") & "'" & _
				   " group by gad.FDDATE,gad.GANGCODE,gad.EMPLOYEECODE,gad.SUPERVISION_1"
		  Set rsAnggota = Server.CreateObject("ADODB.Recordset")
		  rsAnggota.CursorType = 3
		  rsAnggota.CursorLocation = 3
		  'response.Write "Supervisi ........" & rsSupervisi("SUPERVISION_1")  & "<br>"
		  response.Flush()
		  rsAnggota.Open qAnggota, connstr
		  response.Flush()
		  if not rsAnggota.eof then
		     Do While Not rsAnggota.eof
			    qEmpAllded=" Select nvl(sum(decode(substr(allowdedcode,1,1),'A',amount,amount*-1)),0) as PremiSortasi " & _
		                   " from EmpAllded where empcode='" & rsAnggota("EMPLOYEECODE") & "'" & _
                           " and adHocDate=to_date('" & rsAnggota("FDDATE") & "','mm/dd/yyyy') "  &  _  
						   " and AllowdedCode in ('A24','A27','A28','D09','D10') " & _
			               " and fcBA='" & Session("fcBA") & "'"
				Set rsEmpAllded = Server.CreateObject("ADODB.Recordset")
				'response.Write "rsAnggota " & rsAnggota("EMPLOYEECODE")  &  "<br>"
				response.Flush()
				rsEmpAllded.CursorType = 3
				rsEmpAllded.CursorLocation = 3
				rsEmpAllded.Open qEmpAllded, connstr
				if not rsEmpAllded.eof then
                   nPremiSortasi=nPremiSortasi+cdbl(rsEmpAllded("PremiSortasi"))
				   'response.Write  "Payslip : " & rsEmpAllded("PremiSortasi")  & "<br>"
				   response.Flush()
                end if
				rsEmpAllded.Close
				set rsEmpAllded = Nothing
				
			    rsAnggota.MoveNext
			 Loop
          end if
          rsAnggota.Close
		  set rsAnggota = Nothing
		  if cdbl(rsSupervisi("HK")) = 0 then
		     nPremiMandore1 = 0 
		  end if
		  if cdbl(rsSupervisi("HK")) < 15 then
		     nPremiMandore1 = ((nPremiSortasi) / 15 ) * (Rate_Mandore_1/100)
		  else
             nPremiMandore1 = ((nPremiSortasi) / cdbl(rsSupervisi("HK")) ) * (Rate_Mandore_1/100)
		  end if
		  if nPremiMandore1 > 0 then
		     'response.Write " Total Untuk Supervisi......" & nPremiMandore1 & "<br>"
			 qAllowance="Insert Into  EmpAllded (EMPCODE, ALLOWDEDCODE, ADHOCDATE, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA) values  " & _
					    "('" & rsSupervisi("SUPERVISION_1") & "','" & "A39" & "'," &  " to_date('" & rsSupervisi("FDDATE") & "','mm/dd/yyyy') " & "," & nPremiMandore1 & _
					    ",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					    Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			 connstr.execute(qAllowance)
			 response.Flush()
		  end if
		  nPremiSortasi=0
	      rsSupervisi.MoveNext
	   loop
	End if
	rsSupervisi.Close
	set rsSupervisi = Nothing
%>