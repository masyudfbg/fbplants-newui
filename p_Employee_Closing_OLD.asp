<%@Language=VBScript%>
<%Response.Buffer =  True
 Response.Expires = 0
 Response.ExpiresAbsolute = now() -1
 Response.CacheControl = "no-cache"%>

<%
'Option Explicit
%>
<%
'On Error Resume Next
%> 
<%
'response.buffer=True
%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Employee Closing</title>
</head>
<body>



<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
Dim TimeAwal,TimeAkhir,Hitung,Lama

' Add variable by NAF
Dim strSQLHutangGaji, rsHutangGaji, tempYear, tempMonth, tempDate, strDate, rsDate
Dim strJobClearing, rsstrJobClearing, tempJobClearing
Dim strJobHutang, rsstrJobHutang, tempJobHutang
Dim tempJobCode
Dim qInfoParameter, rsInfoParameter, temp_Info_JobCode, temp_Info_LocationType, temp_Info_LocationCode
Dim tempActivityAllocSupervisor
Dim qCostActivityMax, rsCostActivityMax, maxLineNo
Dim qMaxLineNo, rsqMaxLineNo, maxLineNoEmployeeCB
Dim rsAllowded, qAllowded, nxActivity, nxGaji, rsEmployee, qEmployee
Dim tBasisType, qEmpAllDed, n_Credit, rsBasis, nBasis, qBasis
Dim qInsertCR, qInsertDR, rsGAD, qGAD, nAlokasi, nTotalAlokasi, nSelisih
Dim LT, LC, FDDATE, qInsertDB, txtDate
Dim JOBCODE, rsAllowLembur, nDay, rsBorongan, qBorongan, qABC, q
Dim qLineMax, rsqLineMax, tempMaxLine, strClosingLog, strLenghtime, rsLenghtime, upClosingLog
Dim txtLastDay, txtNextMonth
											
' End of variable Definition

TimeAwal=Time()

'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))
'--End Added--	
%>


<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Employee Last Period Close <%=FormatDateTime(Session("cls_Employee"),1)%> <br> 
Employee Current Period is <%=FormatDateTime(Session("v_Employee"),1)%> <br><br>
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br><br>" 

dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl

qPeriodControl=" select  to_char(add_months(LASTPERIODPOSTED,1), 'MM/DD/YYYY')  LastPeriodPosted from PERIODCONTROL  " & _
  " where SYSTEM='Employee' and fcba='" & Session("fcBA") & "'"

Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
rsPeriodControl.CursorType = 3
rsPeriodControl.CursorLocation = 3
rsPeriodControl.Open qPeriodControl, connstr

if not rsPeriodControl.eof then
    temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Employee " & "<br>"
end if
' End Added
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX







'---- Control Job
qControlJob="select ControlJob from Module where NAME='Employee'"
Set rsControlJob = Server.CreateObject("ADODB.Recordset")
rsControlJob.CursorType = 3
rsControlJob.CursorLocation = 3
rsControlJob.Open qControlJob, connstr
if not rsControlJob.eof then
   txtControlJob=rsControlJob("ControlJob")
else
   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For Employee is Empty"
   txtControlJob=""
end if
' XXX response.Flush()

if txtControlJob<>"" then
	'---- Normalisasi
	connstr.BeginTrans
	Response.Write " Normalisasi Employee Base Costing" & "<br>"
	' XXX Response.Flush()

   	
    qNormalisasi=" Delete From Employeecb where  to_number(to_Char(transdate,'mm'))=" & Month(Session("v_Employee")) &  _
				 " and to_number(to_Char(transdate,'yyyy'))=" & Year(Session("v_Employee")) & " and fcBA='" & Session("fcBA") & "'"
	connstr.execute(qNormalisasi)

	
	' To Pick max Line No in Cost_Activity
	qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
					" from employeecb " & _
					" where " & _
					" fcba = '" & Session("fcBA") & "' "
					
	Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
	rsqMaxLineNo.CursorType = 3
	rsqMaxLineNo.CursorLocation = 3
	rsqMaxLineNo.Open qMaxLineNo, connstr

	if not rsqMaxLineNo.eof then
		maxLineNoEmployeeCB = cdbl(rsqMaxLineNo("maxLineNo"))
	else
		maxLineNoEmployeeCB = 0
	end if
	rsqMaxLineNo.close
	set rsqMaxLineNo = Nothing	
	' End of Pick Max Line No.
	
	
	
	qNormalisasi=" Delete From COST_ACTIVITY where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Employee")) &  _
				 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Employee")) & " and fcBA='" & Session("fcBA") & "'" & _
				 " and SOURCETYPE in ('EM','ES')"
	connstr.execute(qNormalisasi)

	Response.Write " Negative Non Allowance..." & "<br>"
	' XXX Response.Flush()
    'Positifkan dulu potongan untuk alokasi
    qNormalisasi="Update EMPALLDED set Amount=amount*-1 where  to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
				 " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
				 " and ALLOWDEDCODE in (select FCCODE from ALLOWDED where ALLDED<>'A' and fcBA='" & Session("fcBA") & "') " & _
				 " and fcBA='" & Session("fcBA") & "'"
	connstr.execute(qNormalisasi)

    Set rsAllowded = Server.CreateObject("ADODB.Recordset")
	rsAllowded.CursorType = 3
	rsAllowded.CursorLocation = 3
	qAllowded = " Select FCCODE, FCBA, FCNAME, ALLDED_GROUP, ALLDED, CROP, CONTROL_JOB, " & _
	            " MAN_AUTO, TYPE_1, TYPE_2, TYPE_3, TYPE_4, " & _
	            " MIN_AMT, MAX_AMT, RATE_1, RATE_2, RATE_3, RATE_4, job_alloc_supervisor from ALLOWDED Where FCCODE in " & _
	            " (Select ALLOWDEDCODE From EmpAllDed Where to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
				                                      " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
													  " and fcBA='" & Session("fcBA") & "')" & _
				" and FCCODE not in ('A18','A45') " & _									  
			    " and fcBA='" & Session("fcBA") & "'" '& _

	
	rsAllowded.Open qAllowded, connstr
	if not rsAllowded.eof then
	   Do While Not rsAllowded.eof
   		  Response.Write "to Allocate the  " & rsAllowded("FCNAME") & "<br>"
		  ' XXX Response.Flush()

	      if rsAllowded("ALLDED")="P" or rsAllowded("ALLDED")="A" Then '--- Premi or Allowance
		     nxActivity =  1
		     nxGaji     = -1
	      else
		     nxActivity = -1
		     nxGaji     =  1
	      end if
		  
		  tempActivityAllocSupervisor = rsAllowded("job_alloc_supervisor")

		  Set rsEmployee = Server.CreateObject("ADODB.Recordset")
		  rsEmployee.CursorType = 3
		  rsEmployee.CursorLocation = 3
		  qEmployee = " Select FCCODE, PRESENTGRADE,SECTIONNAME from EMPLOYEE Where FCCODE in " & _
					  " (Select EMPCODE From EmpAllDed Where ALLOWDEDCODE='" & rsAllowded("FCCODE") & "' and AMOUNT<>0 " & _
					                                          " and to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
															  " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
															  " and fcBA='" & Session("fcBA") & "')" & _
					  " and fcBA='" & Session("fcBA") & "'" '& _

			
		  rsEmployee.Open qEmployee, connstr
		  if not rsEmployee.eof then
			 Do While Not rsEmployee.eof
                'Cari Nilai PayslipNya
				if rsAllowded("TYPE_1")="D" Then '-------------------Daily
				   '------------------------ Daily
				   for nDay = 1 to 31
				      
					   Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
					   rsEmpAllDed.CursorType = 3
					   rsEmpAllDed.CursorLocation = 3
				       qEmpAllDed = " Select nvl(sum(amount),0) as Amount,ADHOCDATE From EmpAllDed Where ALLOWDEDCODE='" & rsAllowded("FCCODE") & "' and AMOUNT<>0 " & _
				                " and to_number(to_Char(ADHOCDATE,'DD'))=" & nDay &  _
								" and to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
								" and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
								" and EMPCODE='" & rsEmployee("FcCode") & "'" & _
								" and fcBA='" & Session("fcBA") & "' group by ADHOCDATE"

						rsEmpAllDed.Open qEmpAllDed, connstr
						n_Credit=0
						
						if Not rsEmpAllDed.eof then
						     tBasisType=""
						    '--------------------- Find Basis by Unit
						    Set rsBasis = Server.CreateObject("ADODB.Recordset")
						    rsBasis.CursorType = 3
						    rsBasis.CursorLocation = 3
							nBasis=0
	
							
						    qBasis =  " Select nvl(sum(UNIT),0) as Basis From ATTENDANCE_GAD" & _
								" Where nvl(UNIT,0) <> 0 " & _
								" and to_number(to_Char(FDDATE,'DD'))=" & nDay &  _
								" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
								" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
								" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
								" and fcBA='" & Session("fcBA") & "' " & _
		                        " and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' "
						
							
									
							rsBasis.Open qBasis, connstr
							if Not rsBasis.eof then
							   if cdbl(rsBasis("Basis")) <> 0 then
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
							if nBasis=0 Then
							   'rsBasis.Close
							   'set rsBasis=Nothing
							   '------------------------ Find Basis by Mandays
							   Set rsBasis = Server.CreateObject("ADODB.Recordset")
							   rsBasis.CursorType = 3
							   rsBasis.CursorLocation = 3
							   
							   qBasis =  " Select nvl(sum(Mandays),0) as Basis From ATTENDANCE_GAD" & _
									" Where nvl(Mandays,0) <> 0 " & _
									" and to_number(to_Char(FDDATE,'DD'))=" & nDay &  _
									" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
									" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
									" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
									" and fcBA='" & Session("fcBA") & "' " & _
		                            " and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' "

							   rsBasis.Open qBasis, connstr
							   if Not rsBasis.eof then
							      if cdbl(rsBasis("Basis"))<> 0 then
								     nBasis=cdbl(rsBasis("Basis"))
									 tBasisType="MANDAYS"
								  else
								     nBasis=0
								  end if	 
							   else
							      nBasis=0	  
							   End if
							   rsBasis.Close
							   set rsBasis=Nothing
                            End if
							
						    if nBasis > 0 Then
								'--- Save Credit							
								if cdbl(rsEmpAllDed("Amount")) <>0 then
								   ' Insert the Credit for CLEARING ACCOUNT
								   maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
								   qInsertCR=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
											   " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
											   " values ('" & rsEmployee("FcCode") & "','EM','" & "CB" & "','" & _
											   "BS" & "','" & txtControlJob & "'," & 0 & ","& _
											   nxGaji*cdbl(rsEmpAllDed("Amount"))  & "," & " to_date('" & rsEmpAllDed("ADHOCDATE") & "','mm/dd/yyyy'),'" & rsAllowded("FCCODE") &  "-NAD1','" & _
											   Session("CodeName") & "','" & Session("CodeName") & "','" & _
											   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNoEmployeeCB & ")"
								  'response.Write(qInsertCR)
								  ' XXX response.Flush()			  
								  connStr.execute(qInsertCR)
								End if
								'--------------------- Start GAD
								Set rsGAD = Server.CreateObject("ADODB.Recordset")
								rsGAD.CursorType = 3
								rsGAD.CursorLocation = 3
								
								qGAD =  " Select JOBCODE, LOCATIONTYPE,LOCATIONCODE,FDDATE,nvl(sum(UNIT),0) as UNIT,sum(Mandays) as Mandays From ATTENDANCE_GAD" & _
										" Where " & _
										" to_number(to_Char(FDDATE,'DD'))=" & nDay &  _
										" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
										" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
										" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
										" and fcBA='" & Session("fcBA") & "' " & _
										" and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' " & _
										" Group by JOBCODE, LOCATIONTYPE,LOCATIONCODE,FDDATE"
								
								
								rsGAD.Open qGAD, connstr
								nAlokasi=0
								nTotalAlokasi=0
								nSelisih=0
								if Not rsGAD.eof then
								   Do While Not rsGAD.eof
									  if nBasis > 0 Then
									     if tBasisType="UNIT" Then
										    nAlokasi = Round ( ( cdbl(rsGAD("UNIT"))/cdbl(nBasis)*cdbl(rsEmpAllDed("Amount")) * nxActivity ) ,0 )
										 else
										    nAlokasi = Round ( ( cdbl(rsGAD("MANDAYS"))/cdbl(nBasis)*cdbl(rsEmpAllDed("Amount")) * nxActivity ) ,0 )
										 End if
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
													  " FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME, LINENO)" & _
													  " values ('" & rsEmployee("FcCode") & "','EM','" & temp_Info_LocationType & "','" & _
													  temp_Info_LocationCode & "','" & temp_Info_JobCode & "'," & rsGAD("Unit") & "," &  rsGAD("Mandays") & "," & _
													  nAlokasi  & "," & " to_date('" & rsGAD("FDDATE") & "','mm/dd/yyyy'),'" & rsAllowded("FCCODE") & "-NAD3','" & _
													  Session("CodeName") & "','" & Session("CodeName") & "','" & _
													  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"											
																

											txtDate=rsGAD("FDDATE")			  
											connStr.execute(qInsertDB)
										  End if	  
									  End if
									  LT    = temp_Info_LocationType
									  LC    = temp_Info_LocationCode
									   
									  FDDATE= rsGAD("FDDATE")
									  
									  rsGAD.MoveNext
									  nTotalAlokasi=nTotalAlokasi+nAlokasi
								   Loop
								   'Selisih Pembulatan
								   if nTotalAlokasi<>cdbl(rsEmpAllDed("Amount")) then
									  if rsAllowded("ALLDED")="P" or rsAllowded("ALLDED")="A" then
									     nSelisih = cdbl(rsEmpAllDed("Amount"))-nTotalAlokasi
									  else
									     nSelisih = (cdbl(rsEmpAllDed("Amount"))+nTotalAlokasi)*-1
									  end if
									  maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1											
							          qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
													  " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME,  LINENO)" & _
													  " values ('" & rsEmployee("FcCode") & "','EM','" & LT & "','" & _
													  LC & "','" & temp_Info_JobCode & "'," & 0 & ","& _
													  nSelisih  & "," & " to_date('" & FDDATE & "','mm/dd/yyyy'),'" & rsAllowded("FCCODE") & "-NAD4','" & _
													  Session("CodeName") & "','" & Session("CodeName") & "','" & _
													  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
									  connStr.execute(qInsertDB)
									  nSelisih=0
								   End if
								End if
								rsGAD.Close
								Set rsGAD = Nothing
								'--------------------- End of GAD
						    Else 
							    'Jika Unit Atau Mandays 0 Cek Apakah Mandore Atau Tidak
								%>
								<!--#include virtual="Alokasi_Mandore_Daily.Asp"-->
							    <%
							End if '-------------- End if Basis
						end if '----- End if Empallded
						rsEmpAllDed.Close
						set rsEmpAllDed= Nothing						
				   Next			
				   '------------------------ Enf Daily
				else 
				   '------------------------ Monthly
					   tBasisType=""
					   Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
					   rsEmpAllDed.CursorType = 3
					   rsEmpAllDed.CursorLocation = 3
				       qEmpAllDed = " Select nvl(sum(amount),0) as Amount From EmpAllDed Where ALLOWDEDCODE='" & rsAllowded("FCCODE") & "' and AMOUNT<>0 " & _
								" and to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
								" and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
								" and EMPCODE='" & rsEmployee("FcCode") & "'" & _
								" and fcBA='" & Session("fcBA") & "' group by to_char(ADHOCDATE,'YYYYMM') "
						'response.Write qEmpAllDed & "<br>"
						' XXX response.Flush()
						rsEmpAllDed.Open qEmpAllDed, connstr
						n_Credit=0
						tBasisType=""
						if Not rsEmpAllDed.eof then
						
						    '--------------------- Find Basis by Unit
						    Set rsBasis = Server.CreateObject("ADODB.Recordset")
						    rsBasis.CursorType = 3
						    rsBasis.CursorLocation = 3
							nBasis=0
							
							
							' Add the Attendance <> 'KB' to handle for ALL NON KB - Kerja Borongan data
							' By NAF 31 October 2013
							qBasis =  " Select nvl(sum(UNIT),0) as Basis From ATTENDANCE_GAD" & _
							        " Where nvl(UNIT,0) <>0   " & _
									" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
									" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
									" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
									" and fcBA='" & Session("fcBA") & "' " & _
									" and attendance <> 'KB' " & _
									" and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' "
							

							' XXX response.Flush()		
							rsBasis.Open qBasis, connstr
							if Not rsBasis.eof then
							   if cdbl(rsBasis("Basis")) <> 0 then
                                  nBasis=cdbl(rsBasis("Basis"))
							      tBasisType="UNIT"
							   else
                                  nBasis=0
							   end if	  
							   rsBasis.Close
							   set rsBasis=Nothing
							ELSE
							   nBasis=0   
							END IF
							if nBasis=0 Then

							   '------------------------ Find Basis by Mandays
							   Set rsBasis = Server.CreateObject("ADODB.Recordset")
							   rsBasis.CursorType = 3
							   rsBasis.CursorLocation = 3
							   
							   ' Add the Attendance <> 'KB' to handle for ALL NON KB - Kerja Borongan data
							   ' By NAF 31 October 2013
							
							   qBasis =  " Select nvl(sum(Mandays),0) as Basis From ATTENDANCE_GAD" & _
										" Where  nvl(Mandays,0) <>0   " & _
										" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
										" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
										" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
										" and fcBA='" & Session("fcBA") & "' " & _
										" and attendance <> 'KB' " & _
		                                " and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' "
							   
			
							   rsBasis.Open qBasis, connstr
							   if Not rsBasis.eof then
							      if cdbl(rsBasis("Basis"))<> 0 then
								     nBasis=cdbl(rsBasis("Basis"))
								     tBasisType="MANDAYS"
								  else
								     nBasis=0
								  end if	 
							   else
							      nBasis=0	  
							   End if
							   rsBasis.Close
							   set rsBasis=Nothing
                            End if
							
						    if nBasis > 0 Then
								'--- Save Credit							
								'response.Write "OOO UNIT " & "<BR>"
								if cdbl(rsEmpAllDed("Amount")) <>0 then
								   maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1									   
								   qInsertCR=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
											   " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
											   " values ('" & rsEmployee("FcCode") & "','EM','" & "CB" & "','" & _
											   "BS" & "','" & txtControlJob & "'," & 0 & ","& _
											   nxGaji*cdbl(rsEmpAllDed("Amount"))  & "," & " last_day(to_date('" & Year(Session("v_Employee")) & "'||" & _
												"lpad('" & Month(Session("v_Employee")) & "',2,'0')||'01'" & _
											   ",'YYYYMMDD')),'" & rsAllowded("FCCODE") & "-NAD5','" & _
											   Session("CodeName") & "','" & Session("CodeName") & "','" & _
											   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
								  'response.Write(qInsertCR)
								  ' XXX response.Flush()			  
								  connStr.execute(qInsertCR)
								End if
								'--------------------- Start GAD
								Set rsGAD = Server.CreateObject("ADODB.Recordset")
								rsGAD.CursorType = 3
								rsGAD.CursorLocation = 3
								
								qGAD =  " Select JOBCODE, LOCATIONTYPE,LOCATIONCODE,FDDATE,nvl(sum(UNIT),0) as UNIT,sum(Mandays) as Mandays From ATTENDANCE_GAD" & _
										" Where " & _
										" to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
										" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
										" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
										" and fcBA='" & Session("fcBA") & "' " & _
										" and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' " & _
										" Group by JOBCODE, LOCATIONTYPE,LOCATIONCODE,FDDATE"
								
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
										    nAlokasi = round( ( cdbl(rsGAD("Mandays"))/cdbl(nBasis)*cdbl(rsEmpAllDed("Amount")) * nxActivity ),0)
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
													  " values ('" & rsEmployee("FcCode") & "','EM','" & temp_Info_LocationType & "','" & _
													  temp_Info_LocationCode & "','" & temp_Info_JobCode & "'," & rsGAD("Unit") & "," &  rsGAD("Mandays") & "," & _
													  nAlokasi  & "," & " to_date('" & rsGAD("FDDATE") & "','mm/dd/yyyy'),'" & rsAllowded("FCCODE") & "-NAD3','" & _
													  Session("CodeName") & "','" & Session("CodeName") & "','" & _
													  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',  sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"	
													  

											txtDate=rsGAD("FDDATE")			  
											connStr.execute(qInsertDB)
										  End if	  
									  End if
									  
									  LT    = temp_Info_LocationType
									  LC    = temp_Info_LocationCode
									  
									  FDDATE= rsGAD("FDDATE")
									  
									  rsGAD.MoveNext
									  nTotalAlokasi=nTotalAlokasi+nAlokasi								  
								   Loop
								   'Selisih Pembulatan
								   if nTotalAlokasi<>cdbl(rsEmpAllDed("Amount")) then
									  if rsAllowded("ALLDED")="P" or rsAllowded("ALLDED")="A" then
									     nSelisih = cdbl(rsEmpAllDed("Amount"))-nTotalAlokasi
									  else
									     nSelisih = (cdbl(rsEmpAllDed("Amount"))+nTotalAlokasi)*-1
									  end if
									  maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1						 								   
							          qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
													  " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
													  " values ('" & rsEmployee("FcCode") & "','EM','" & LT & "','" & _
													  LC & "','" & temp_Info_JobCode & "'," & 0 & ","& _
													  nSelisih  & "," & " to_date('" & FDDATE & "','mm/dd/yyyy'),'" & rsAllowded("FCCODE") & "-NAD8" & "','" & _
													  Session("CodeName") & "','" & Session("CodeName") & "','" & _
													  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',  sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
									  connStr.execute(qInsertDB)
									  nSelisih=0
								   End if
								   
								   
								End if
								rsGAD.Close
								Set rsGAD = Nothing
								'--------------------- End of GAD
						    Else '----- Jika Basis Tidak Ada Debet Dari Control Job
							    %>
							    <!--#include virtual="Alokasi_Mandore_Monthly.Asp"-->
								<%
							End if '-------------- End if Basis
						end if '----- End if Empallded
						rsEmpAllDed.Close
						set rsEmpAllDed= Nothing				   '------------------------ Akhir Monthly
				end if
								
		        rsEmployee.MoveNext
		     Loop '----------- Loop For Employee  
			 
			 rsEmployee.Close
			 set rsEmployee = Nothing
		  End if    
		  rsAllowded.MoveNext
       loop '--------- Loop For Allowance
	   rsAllowded.close
	   set rsAllowded = Nothing
    end if
	
	'---------------Lembur  
    Set rsAllowLembur = Server.CreateObject("ADODB.Recordset")
	rsAllowLembur.CursorType = 3
	rsAllowLembur.CursorLocation = 3
	qAllowded = " Select FCCODE, FCBA, FCNAME, ALLDED_GROUP, ALLDED, CROP, CONTROL_JOB, " & _
	            " MAN_AUTO, TYPE_1, TYPE_2, TYPE_3, TYPE_4, " & _
	            " MIN_AMT, MAX_AMT, RATE_1, RATE_2, RATE_3, RATE_4 from ALLOWDED Where FCCODE in " & _
	            " (Select ALLOWDEDCODE From EmpAllDed Where to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
				                                      " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
													  " and fcBA='" & Session("fcBA") & "')" & _
				" and FCCODE in ('A18')" & _									  
			    " and fcBA='" & Session("fcBA") & "'" '& _
				'" and TYPE_1='D'"
	
	rsAllowLembur.Open qAllowded, connstr
	if not rsAllowLembur.eof then
	   Do While Not rsAllowLembur.eof
   		  Response.Write "to Allocate the  " & rsAllowLembur("FCNAME") & "<br>"
		  ' XXX Response.Flush()

	      if rsAllowLembur("ALLDED")="P" or rsAllowLembur("ALLDED")="A" Then '--- Premi or Allowance
		     nxActivity =  1
		     nxGaji     = -1
	      else
		     nxActivity = -1
		     nxGaji     =  1
	      end if

		  Set rsEmployee = Server.CreateObject("ADODB.Recordset")
		  rsEmployee.CursorType = 3
		  rsEmployee.CursorLocation = 3
		  qEmployee = " Select FCCODE, PRESENTGRADE,SECTIONNAME from EMPLOYEE Where FCCODE in " & _
					  " (Select EMPCODE From EmpAllDed Where ALLOWDEDCODE='" & rsAllowLembur("FCCODE") & "' and AMOUNT<>0 " & _
					                                          " and to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
															  " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
															  " and fcBA='" & Session("fcBA") & "')" & _
					  " and fcBA='" & Session("fcBA") & "'" '& _
			
		  rsEmployee.Open qEmployee, connstr
		  if not rsEmployee.eof then
			 Do While Not rsEmployee.eof
                'Cari Nilai PayslipNya
				if rsAllowLembur("TYPE_1")="D" Then '-------------------Daily
				   '------------------------ Daily
				   for nDay = 1 to 31
					   Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
					   rsEmpAllDed.CursorType = 3
					   rsEmpAllDed.CursorLocation = 3
				       qEmpAllDed = " Select nvl(sum(amount),0) as Amount,ADHOCDATE From EmpAllDed Where ALLOWDEDCODE='" & rsAllowLembur("FCCODE") & "' and AMOUNT<>0 " & _
				                " and to_number(to_Char(ADHOCDATE,'DD'))=" & nDay &  _
								" and to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
								" and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
								" and EMPCODE='" & rsEmployee("FcCode") & "'" & _
								" and fcBA='" & Session("fcBA") & "' group by ADHOCDATE"

						rsEmpAllDed.Open qEmpAllDed, connstr
						n_Credit=0
						
						if Not rsEmpAllDed.eof then
							'--- Save Credit							
							if cdbl(rsEmpAllDed("Amount")) <>0 then
							   maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							   qInsertCR=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
							               " FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME,  LINENO)" & _
										   " values ('" & rsEmployee("FcCode") & "','EM','" & "CB" & "','" & _
										   "BS" & "','" & txtControlJob & "'," & 0 & ","& _
										   nxGaji*cdbl(rsEmpAllDed("Amount"))  & "," & " to_date('" & rsEmpAllDed("ADHOCDATE") & "','mm/dd/yyyy'),'" & rsAllowLembur("FCCODE") &"-NAF4" & "','" & _
										   Session("CodeName") & "','" & Session("CodeName") & "','" & _
										   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',  sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
							  'response.Write(qInsertCR)
							  ' XXX response.Flush()			  
							  connStr.execute(qInsertCR)
                            End if
							' -- cdbl(rsEmpAllDed("Amount")) 
						
						    '--------------------- Find Basis by Unit
						    Set rsBasis = Server.CreateObject("ADODB.Recordset")
						    rsBasis.CursorType = 3
						    rsBasis.CursorLocation = 3
							nBasis=0
							
						    qBasis =  " Select nvl(sum(nvl(OTHRS,0)),0) as Basis From ATTENDANCE_GAD" & _
							        " Where OTHRS > 0  " & _
									" and to_number(to_Char(FDDATE,'DD'))=" & nDay &  _
									" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
									" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
									" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
									" and fcBA='" & Session("fcBA") & "' " & _
									" and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' "						
							
	
							rsBasis.Open qBasis, connstr
							if Not rsBasis.eof then
                               nBasis=cdbl(rsBasis("Basis"))
							   rsBasis.Close
							   set rsBasis=Nothing
							elseif nBasis=0 Then
							   rsBasis.Close
							   set rsBasis=Nothing
							   '------------------------ Find Basis by Mandays
							   Set rsBasis = Server.CreateObject("ADODB.Recordset")
							   rsBasis.CursorType = 3
							   rsBasis.CursorLocation = 3
							   
							   qBasis =  " Select nvl(sum(Mandays),0) as Basis From ATTENDANCE_GAD" & _
										" Where OTHRS > 0  " & _
										" and to_number(to_Char(FDDATE,'DD'))=" & nDay &  _
										" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
										" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
										" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
										" and fcBA='" & Session("fcBA") & "' " & _
										" and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' "
							   
							   
							   rsBasis.Open qBasis, connstr
							   if Not rsBasis.eof then
								  nBasis=cdbl(rsBasis("Basis"))
							   End if
							   rsBasis.Close
							   set rsBasis=Nothing
                            End if
							
						    if nBasis > 0 Then
								'--------------------- Start GAD
								Set rsGAD = Server.CreateObject("ADODB.Recordset")
								rsGAD.CursorType = 3
								rsGAD.CursorLocation = 3
								
								qGAD =  " Select LOCATIONTYPE,LOCATIONCODE,FDDATE,jobcode,nvl(sum(OTHRS),0) as OTHRS From ATTENDANCE_GAD" & _
										" Where OTHRS > 0  " & _
										" and to_number(to_Char(FDDATE,'DD'))=" & nDay &  _
										" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
										" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
										" and EMPLOYEECODE='" & rsEmployee("FcCode") & "'" & _
										" and fcBA='" & Session("fcBA") & "' " & _
										" and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' " & _
										" Group by LOCATIONTYPE,LOCATIONCODE,FDDATE,jobcode"			
										
	
								rsGAD.Open qGAD, connstr
								nAlokasi=0
								nTotalAlokasi=0
								nSelisih=0
								if Not rsGAD.eof then
								    Do While Not rsGAD.eof
										if nBasis > 0 Then
											 nAlokasi = Round ( ( cdbl(rsGAD("OTHRS"))/cdbl(nBasis)*cdbl(rsEmpAllDed("Amount")) * nxActivity ) ,0 )
											 if nAlokasi <> 0 then
											 
													' Get for Location Type, Location Code and Job Code based on Parameter Table
													
													
													Set rsInfoParameter = Server.CreateObject("ADODB.Recordset")
													rsInfoParameter.CursorType = 3
													rsInfoParameter.CursorLocation = 3
													
													
													qInfoParameter = " select GetInfoParameter ( '" & rsAllowLembur("FCCODE") &"', '" & _ 
               											       rsGAD("LOCATIONTYPE") & "', '" & rsGAD("JOBCODE") &"', 1, '" & Session("fcBA") & "') " & _
															   " as JobCode , " & _
															   " GetInfoParameter ( '" & rsAllowLembur("FCCODE") &"', '" & _ 
               											       rsGAD("LOCATIONTYPE") & "', '" & rsGAD("LOCATIONTYPE") &"', 2, '" & Session("fcBA") & "') " & _
															   " as locationtype ,  " & _
															   " GetInfoParameter ( '" & rsAllowLembur("FCCODE") &"', '" & _ 
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
													qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
															  " FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME,  LINENO)" & _
															  " values ('" & rsEmployee("FcCode") & "','EM','" & temp_Info_LocationType & "','" & _
															  temp_Info_LocationCode & "','" & temp_Info_JobCode & "'," & 0 & ","& _
															  nAlokasi  & "," & " to_date('" & rsGAD("FDDATE") & "','mm/dd/yyyy'),'" & rsAllowLembur("FCCODE") &"-NAF5" & "','" & _
															  Session("CodeName") & "','" & Session("CodeName") & "','" & _
															  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',  sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
			
													txtDate=rsGAD("FDDATE")			  
													connStr.execute(qInsertDB)
											  End if	  
										End if
										LT    = temp_Info_LocationType
										LC    = temp_Info_LocationCode
									   
										FDDATE= rsGAD("FDDATE")
									  

									  
										rsGAD.MoveNext
										nTotalAlokasi=nTotalAlokasi+nAlokasi
								    Loop
								    'Selisih Pembulatan
								    if nTotalAlokasi<>cdbl(rsEmpAllDed("Amount")) then
								  
										  if rsAllowLembur("ALLDED")="P" or rsAllowLembur("ALLDED")="A" then
												 if cdbl(rsEmpAllDed("Amount")) > nTotalAlokasi then 
													nSelisih = nTotalAlokasi-cdbl(rsEmpAllDed("Amount"))
												 else
													nSelisih = cdbl(rsEmpAllDed("Amount"))-nTotalAlokasi
												 end if
										  else
												 if cdbl(rsEmpAllDed("Amount")) > nTotalAlokasi then 
													nSelisih = nTotalAlokasi-cdbl(rsEmpAllDed("Amount"))
												 else
													nSelisih = cdbl(rsEmpAllDed("Amount"))-nTotalAlokasi
												 end if
										  end if								  
								  
									      maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
										  qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
														  " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
														  " values ('" & rsEmployee("FcCode") & "','EM','" & LT & "','" & _
														  LC & "','" & temp_Info_JobCode & "'," & 0 & ","& _
														  nSelisih  & "," & " to_date('" & FDDATE & "','mm/dd/yyyy'),'" & rsAllowLembur("FCCODE") & "-NAF1" & "','" & _
														  Session("CodeName") & "','" & Session("CodeName") & "','" & _
														  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',  sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
										  connStr.execute(qInsertDB)
										  nSelisih=0
									End if
									' End Selisih Pembulatan
								End if
								rsGAD.Close
								Set rsGAD = Nothing
								'--------------------- End of GAD
						    End if
							'End if NBASIS
						end if '----- End if Empallded
						rsEmpAllDed.Close
						set rsEmpAllDed= Nothing	
				   Next			
				End if				
		        rsEmployee.MoveNext
				 tBasisType=""
		     Loop '----------- Loop For Employee  
			 
			 rsEmployee.Close
			 set rsEmployee = Nothing
			  tBasisType=""
		  End if    
		  rsAllowLembur.MoveNext
		  tBasisType=""
       loop '--------- Loop For Allowance
	   rsAllowLembur.close
	   set rsAllowLembur = Nothing
	   tBasisType=""
    end if
	
	'---------------Akhir Lembur
end if
Response.Write "to Allocate the  GAJI BORONGAN"
' XXX Response.Flush()
Set rsBorongan = Server.CreateObject("ADODB.Recordset")
rsBorongan.CursorType = 3
rsBorongan.CursorLocation = 3

'sum(nvl(att_gad.unit,0) * nvl(att_gad.Rate,0))

' nvl(sum(UNIT),0) as UNIT,nvl(sum(Rate),0) as Rate 
		
qBorongan =  " Select employeecode,LOCATIONTYPE,LOCATIONCODE,Jobcode,FDDATE, " & _
	    " sum(nvl(unit,0) * nvl(Rate,0)) totalAmount " & _
        " From ATTENDANCE_GAD" & _
		" Where unit <> 0 and rate <> 0 and (mandays=0 or mandays is null) " & _
		" and Attendance in ('KB') " & _
		" and to_number(to_Char(FDDATE,'mm'))=" & Month(Session("v_Employee")) &  _
		" and to_number(to_Char(FDDATE,'yyyy'))=" & Year(Session("v_Employee")) & _ 
		" and fcBA='" & Session("fcBA") & "' " & _
		" and ATTENDANCE_GAD_API.get_GangTypeCode (fcba, fddate, employeecode, jobcode, locationtype, locationcode) <> 'SUPERVISION' " & _
		" Group by employeecode,LOCATIONTYPE,LOCATIONCODE,Jobcode,FDDATE"
		
rsBorongan.Open qBorongan, connstr
if Not rsBorongan.eof then
   Do While Not rsBorongan.eof
      maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
	  qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
				  " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
				  " values ('" & rsBorongan("employeecode") & "','EM','" & rsBorongan("LOCATIONTYPE") & "','" & _
				  rsBorongan("LOCATIONCODE") & "','" & rsBorongan("Jobcode") & "',"  & _
				  cdbl(rsBorongan("totalAmount")) & "," & " to_date('" & rsBorongan("FDDATE") & "','mm/dd/yyyy'),'" & "A45" & "','" & _
				  Session("CodeName") & "','" & Session("CodeName") & "','" & _
				  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',  sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
	  ' response.Write(qInsertDB)			  
      connStr.execute(qInsertDB)
	  maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
	  qInsertCR=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
				   " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
				   " values ('" & rsBorongan("employeecode") & "','EM','" & "CB" & "','" & _
				   "BS" & "','" & txtControlJob & "'," & 0 & ","& _
				   -1*cdbl(rsBorongan("totalAmount"))  & "," & " to_date('" & rsBorongan("FDDATE") & "','mm/dd/yyyy'),'" & "A45" & "','" & _
				   Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
	  connStr.execute(qInsertCR)
	  
	  rsBorongan.MoveNext
   Loop
End if
rsBorongan.close
set rsBorongan = Nothing

Response.Write " Negative Non Allowance..." & "<br>"
' XXX Response.Flush()
'Positifkan dulu potongan untuk alokasi
qNormalisasi="Update EMPALLDED set Amount=amount*-1 where  to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
			 " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
			 " and ALLOWDEDCODE in (select FCCODE from ALLOWDED where ALLDED<>'A' and fcBA='" & Session("fcBA") & "') " & _
			 " and fcBA='" & Session("fcBA") & "'"
connstr.execute(qNormalisasi)



' Get the Max LineNo of Cost Activity


Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
rsCostActivityMax.CursorType = 3
rsCostActivityMax.CursorLocation = 3
								
qCostActivityMax =  " select max(lineno) maxLineNo from cost_activity " & _
					" where fcba = '" &  Session("fcBA") & "' " 
rsCostActivityMax.Open qCostActivityMax, connstr
maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))

rsCostActivityMax.close
set rsCostActivityMax = Nothing
' End of Get the Max LineNo of Cost Activity



'----- Tranfer to Activity Costing
response.Write "Transfer to Activity Based Costing..." & "<br>"

qABC = " INSERT INTO COST_ACTIVITY (" & _
       " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
	   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, LASTUPDATE, LASTTIME, LINENO ) " & _
	      " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
		  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, sysdate LASTUPDATE, to_char(sysdate,'HH24:MI') LASTTIME, rownum + " & maxLineNo & " LINENO " & _
	      " FROM (SELECT cb.transdate tdate, " & _
				         " cb.jobcode jobcode,  " & _
						 " '' source,  " & _
						 " CB.SOURCETYPE sourcetype,  " & _
						 " em.presentgrade sourcecode,  " & _
						 " '' sourcecode_status,  " & _
						 " cb.targettype targettype,   " & _
						 " cb.targetcode targetcode,  " & _
						 " '' targetcode_status,  " & _
						 " cb.referenceno referenceno,  " & _
						 " 0 quantity, " & _
						 " 0 output, " & _
						 " sum(cb.amount) Amount,  " & _
						 " cb.fcba fcba ,  " & _
						 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date  " & _
          " FROM employee em, employeecb cb WHERE ((cb.empcode = em.fccode) AND (cb.fcba = em.fcba))" & _
		  " and to_number(to_Char(cb.transdate,'mm'))=" & Month(Session("v_Employee")) &  _
		  " and to_number(to_Char(cb.transdate,'yyyy'))=" & Year(Session("v_Employee")) & _ 
		  " and CB.fcBA='" & Session("fcBA") & "'" & _
          " group by " & _
          " cb.transdate,cb.jobcode,'',CB.SOURCETYPE,em.presentgrade,'',cb.targettype, cb.targetcode,'',cb.referenceno,cb.fcba ) "
	   
connstr.Execute(qABC)


' Start Modification to Add Setting Data for Allocation HUTANG GAJI

Response.Write "Calculate Reverse for HUTANG GAJI..." & "<br>"
' XXX Response.Flush()


strSQLHutangGaji = " select a.source, a.SOURCEtype, a.sourcecode, a.sourcecode_status, a.targettype, a.targetcode, " & _ 
					" a.targetcode_status, sum(amount) totalAmount from " & _
					" cost_activity a, " & _
					" (select fcba, param_1 from parameter " & _
					" where " & _
					" parcode = 'ALLOC_SALARY' and " & _
					" fccode = 'CLEARINGACCPAY') b " & _
					" where " & _
					" a.jobcode = b.param_1 " & _
					" and to_number(to_Char(a.tdate,'mm'))=" & Month(Session("v_Employee"))  &  _
					" and to_number(to_Char(a.tdate,'yyyy'))=" & Year(Session("v_Employee")) & _ 
					" and a.fcba = '" & Session("fcBA") & "' " & _
					" and a.fcba = b.fcba  " & _
					" group by " & _
					" a.source, a.SOURCEtype, a.sourcecode, a.sourcecode_status, a.targettype, a.targetcode, a.targetcode_status "
					
Set rsHutangGaji = Server.CreateObject("ADODB.Recordset")					
'response.write strSQLHutangGaji
rsHutangGaji.Open strSQLHutangGaji, connstr

if Not rsHutangGaji.eof then
   ' Get the tempDate : the Last date in particular month
   tempYear = Year(Session("v_Employee"))
   if Month(Session("v_Employee")) + 1 = 13 then 
       tempYear = Year(Session("v_Employee")) + 1
	   strDate = " select to_date('01/01/" & tempYear & "','MM/DD/YYYY') - 1 as tempDate from dual "
   else
       tempMonth = Month(Session("v_Employee")) + 1
	   strDate = " select to_date('"& tempMonth & "/01/" & tempYear & "','MM/DD/YYYY') - 1 as tempDate from dual "
   end if
   
   Set rsDate = Server.CreateObject("ADODB.Recordset")					
   rsDate.Open strDate, connstr
   if Not rsDate.eof then 
       tempDate = rsDate("tempDate")
   end if   



   ' Get The Job Code for Hutang Gaji
   strJobClearing = " select param_1 as JobClearing from parameter " & _
						" where " & _
						" parcode = 'ALLOC_SALARY' and " & _
						" fcba = '" & Session("fcBA") & "'  and " & _
						" fccode = 'CLEARINGACCPAY' " 
						

   Set rsstrJobClearing = Server.CreateObject("ADODB.Recordset")					
						
   rsstrJobClearing.Open strJobClearing, connstr
   if Not rsstrJobClearing.eof then 
       tempJobClearing = rsstrJobClearing("JobClearing")
   end if   
   rsstrJobClearing.close
   set rsstrJobClearing = Nothing

   
   ' Get The Job Code for Hutang Gaji
   strJobHutang = " select param_1 as JobHutang from parameter " & _
						" where " & _
						" parcode = 'ALLOC_SALARY' and " & _
						" fcba = '" & Session("fcBA") & "'  and " & _
						" fccode = 'SALARYACCPAY' " 
						
						
   Set rsstrJobHutang = Server.CreateObject("ADODB.Recordset")					

   rsstrJobHutang.Open strJobHutang, connstr
   if Not rsstrJobHutang.eof then 
       tempJobHutang = rsstrJobHutang("JobHutang")
   end if   
   rsstrJobHutang.close
   set rsstrJobHutang = Nothing


   
   


	' Get the Max LineNo of Cost Activity


	Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
	rsCostActivityMax.CursorType = 3
	rsCostActivityMax.CursorLocation = 3
									
	qCostActivityMax =  " select max(lineno) maxLineNo from cost_activity " & _
						" where fcba = '" &  Session("fcBA") & "' " 
	rsCostActivityMax.Open qCostActivityMax, connstr
	maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))

	rsCostActivityMax.close
	set rsCostActivityMax = Nothing
	' End of Get the Max LineNo of Cost Activity   

	Do While Not rsHutangGaji.eof
	  maxLineNo = maxLineNo + 1
      qInsertDB = " insert into cost_activity " & _	  
				    " (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
					    " SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
                        " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, closing_date, process_date, lastupdate, lasttime, lineno ) " & _
						" values ( " & _
						" to_date('" & tempDate & "','MM/DD/YYYY'),'" & _
						tempJobHutang & "','" & _
						rsHutangGaji("source") & _
						"','" & rsHutangGaji("SOURCETYPE") & _
						"','" & rsHutangGaji("SOURCECODE") & _
						"','" & rsHutangGaji("SOURCECODE_STATUS") & _ 
						"','" & rsHutangGaji("TARGETTYPE") & _
						"','" & rsHutangGaji("TARGETCODE") & _
						"','" & rsHutangGaji("TARGETCODE_STATUS") & _
						"','Hutang Gaji Created',0,0," &  cdbl(rsHutangGaji("totalAmount"))& ", '" & _
						Session("fcBA") & "' , to_date('" & temp_datePeriodControl & "','MM/DD/YYYY') , sysdate, sysdate, to_char(sysdate, 'HH24:MI') ," & maxLineNo & "  )" 
						
		  
      connStr.execute(qInsertDB)
	   
	  maxLineNo = maxLineNo + 1 
      qInsertCR = " insert into cost_activity " & _	  
				    " (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
					    " SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
                        " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA , closing_date, process_date, lastupdate, lasttime, lineno ) " & _
						" values ( " & _
						" to_date('" & tempDate & "','MM/DD/YYYY'),'" & _
						tempJobClearing & "','" & _
						rsHutangGaji("source") & _
						"','" & rsHutangGaji("SOURCETYPE") & _
						"','" & rsHutangGaji("SOURCECODE") & _
						"','" & rsHutangGaji("SOURCECODE_STATUS") & _ 
						"','" & rsHutangGaji("TARGETTYPE") & _
						"','" & rsHutangGaji("TARGETCODE") & _
						"','" & rsHutangGaji("TARGETCODE_STATUS") & _
						"','Job Clearing Created',0,0," & -1 * cdbl(rsHutangGaji("totalAmount"))& ", '" & _
						Session("fcBA") & "' , to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, sysdate, to_char(sysdate, 'HH24:MI') ," & maxLineNo & "  )" 
						

	  connStr.execute(qInsertCR)
	  
	  rsHutangGaji.MoveNext
   Loop
End if
rsHutangGaji.close
set rsHutangGaji = Nothing




' End of Modification HUTANG GAJI




%>

<%if err.number <> 0 then%>
    <br><img src="images/error.gif"><%
	connStr.RollbackTrans
	'response.Clear()
	'response.Write qInsertDB
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
	cmdView	= "True"%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>	
<%Else
    connStr.CommitTrans
	'response.Clear()

    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)),CURRENTDATE=add_months(CURRENTDATE,1) " & _
	  " where SYSTEM='Employee' and fcba='" & Session("fcBA") & "'"
	connStr.execute(q)
	

    
	TimeAkhir=Time()
	Lama =TimeAkhir-TimeAwal
	Response.Write "<br>" & "Finish ... " & time() & "<br><br>"
	Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	

	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Employee'" 
							
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
	else
		tempMaxLine = 0
	end if 

	tempMaxLine = tempMaxLine + 1

	strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME) VALUES (" & _
					" 'Employee', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" " & tempMaxLine & "" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"
	'response.write strClosingLog
	connStr.execute(strClosingLog)

	strLenghtime = 	" SELECT" & _
					" FCBA," & _
					" LINENO," & _
					" SYSTEM," & _
					" TO_CHAR(STARTINGTIME, 'DD.MM.YYYY:HH24:MI:SS') STARTINGTIME," & _
					" TO_CHAR(ENDINGTIME, 'DD.MM.YYYY:HH24:MI:SS') ENDINGTIME," & _
					" TRUNC((((86400*(ENDINGTIME-STARTINGTIME))/60)/60)/24) DAYS," & _
					" TRUNC(((86400*(ENDINGTIME-STARTINGTIME))/60)/60)-24*(TRUNC((((86400*(ENDINGTIME-STARTINGTIME))/60)/60)/24)) HRS," & _
					" TRUNC((86400*(ENDINGTIME-STARTINGTIME))/60)-60*(TRUNC(((86400*(ENDINGTIME-STARTINGTIME))/60)/60)) MIN," & _
					" TRUNC(86400*(ENDINGTIME-STARTINGTIME))-60*(TRUNC((86400*(ENDINGTIME-STARTINGTIME))/60)) SEC" & _
					" FROM CLOSINGLOG" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Employee'"
			
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.CursorType = 3
	rsLenghtime.CursorLocation = 3
	rsLenghtime.Open strLenghtime, connstr

	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Employee'"		
	'connStr.execute(upClosingLog)
	'--End Added--	
		
	%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Employee Closing is succesfully...
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		
	<%
	txtLastDay  =cmdFindPeriod("Employee")
    txtNextMonth=cmdPeriodClose("Employee")	
	
	Session("v_Employee")=txtLastDay
	Session("cls_Employee")= txtNextMonth	
	%>
		
		
<%End if%>

</body>
</html>