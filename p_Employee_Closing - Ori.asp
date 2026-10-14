<%@Language=VBScript%>
<%Response.Buffer=true%>
<%Response.Clear%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<title>iPlaS - Employee Closing</title>
	<style>
		*{font-family:Verdana;font-size:10px;}
	</style>
	</head>
	<body>
		<%
			'Declare All Variabel
			Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow
			Dim qCheckStatus, rsCheckStatus, txtClosingStatus
			Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl
			Dim qNormalisasi
			Dim Startingtime, Endingtime
			Dim qMaxLineNo, rsqMaxLineNo, maxLineNoEmployeeCB
			Dim qAllowdedBrondolan, rsqAllowdedBrondolan, tempData
			Dim qBorongan, rsBorongan
			Dim qNoActivity, rsNoActivity
			Dim TimeAwal, TimeAkhir, Lama
			Dim nAlokasi
			Dim qLembur, rsLembur
			Dim qDaily, rsDaily
			Dim qMonthly, rsMonthly
			Dim sqlunittype, rsunittype, tempunittype

			TimeAwal=Time()
			Startingtime = FormatDateTime(CDate(Now))
		%>
		<div id="status_running">
			<img src="ICON/working.gif" />
		</div>
		<br/>
		Employee Last Period Close <%=FormatDateTime(Session("cls_Employee"),1)%> <br> 
		Employee Current Period is <%=FormatDateTime(Session("v_Employee"),1)%> <br><br>
		Please Wait...
		<br/>
		<br/>
		<%
		Response.Write "Start ..." & Time() & "<br><br>"
		Response.Flush
		
		%>
		<%
		
		Function GetRSArray(strSQL)
			set objRSArray = connStr.Execute(strSQL)
			If Not objRSArray.EOF Then
				arrRS = objRSArray.GetRows()
				GetRSArray = arrRS
			Else
				Response.Write("There are currently no records returned.<br/>")
                Exit Function
			End If
			objRSArray.Close()
			Set objRSArray = Nothing
		End Function
		' Water Flow Check, By NAF
		' 2 December 2013
		qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Employee','" & Session("fcBA")  & "') waterFlowCheck from dual "

		Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
		rsWaterFlowCheck.CursorType = 3
		rsWaterFlowCheck.CursorLocation = 3
		rsWaterFlowCheck.Open qWaterFlowCheck, connstr

		if not rsWaterFlowCheck.eof then
			CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
		else
			Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Employee " & "<br>"
		end if
		rsWaterFlowCheck.close
		set rsWaterFlowCheck = Nothing
		
		'-- Check for Empallded PreProcess
		qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreEmployeeClosing (" & Year(Session("v_Employee")) & "," & _
																				Month(Session("v_Employee")) & ",'" & _
																				Session("fcBA") & "') closingStatus from dual "

		Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
		rsCheckStatus.CursorType = 3
		rsCheckStatus.CursorLocation = 3
		rsCheckStatus.Open qCheckStatus, connstr
		if not rsCheckStatus.eof then
		   txtClosingStatus = rsCheckStatus("closingStatus")
		else
		   Response.Write  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
		   txtClosingStatus="FALSE"
		end if
		rsCheckStatus.close
		set rsCheckStatus = Nothing
		
		if CheckWaterFlow = "TRUE" and txtClosingStatus = "TRUE" then
		
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
			rsPeriodControl.close
			set rsPeriodControl = Nothing
			
			 dim QSession
			     QSession = " insert into ActiveSession(Sid,Audsid,Fcba,Modul,User_Iplas,Datestart)  select sid,Audsid,'" & Session("fcBA")  & "','EMPLOYEE' Modul,'" & Session("CodeName") & "',sysdate  from V$SESSION " & _
                             " where audsid in (select sys_context('userenv','sessionid') Session_ID from dual)" 
	             ' Response.write QSession
				  connstr.execute(QSession)
				  connStr.execute("Commit")
			
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
			rsControlJob.close
			set rsControlJob = Nothing
			
			if txtControlJob<>"" then
			    Response.Write "Deleting Data Process" & "<br/>"
				Response.Flush
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from process_empallded_daily" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from process_empallded_monthly" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from process_empallded_overtime" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_es_monthly" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_es_daily" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_em_monthly" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_em_daily" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_overtime" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
                'Add by 16 Oktober 2015 -- To Normalisasi history HK KHL Monthly'
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from monthly_mandays_tab" & _
                             " where fcba = '" & Session("fcBA") & "'" & _
                             "       and year = " & Year(Session("v_Employee")) & "" & _
							 "       and month = " & Month(Session("v_Employee")) & ""


				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				'Add by 16 Oktober 2015 -- To save history HK KHL Monthly'
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_closing" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				Response.write("<font face=Verdana size=1 color=Green>Deleting Data Process, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush


				Response.Write "Normalisasi EmployeeCB" & "<br/>"
				Response.Flush
				connstr.BeginTrans
				qNormalisasi = " delete from employeecb" & _
				               " where to_number(to_Char(transdate,'mm')) = " & Month(Session("v_Employee")) & "" & _
							   "       and to_number(to_Char(transdate,'yyyy')) = " & Year(Session("v_Employee")) & "" & _
							   "       and fcBA = '" & Session("fcBA") & "'"
				connstr.execute(qNormalisasi)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.write("<font face=Verdana size=1 color=Green>Normalisasi EmployeeCB, Success! " & " Time: " & time() & " ..<br/><br/></font>")
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qNormalisasi&"] ("&Err.Description&")")
					Response.Flush
				End If
				
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

				Response.Write "Normalisasi Cost_Activity" & "<br/>"
				Response.Flush
				connstr.BeginTrans
				
				' Modify to make and Update State, by NAF, 3 Nov 2014
				qNormalisasi = " update COST_ACTIVITY set deleteflag = 'REMOVE' " & _
				               " where to_number(to_Char(TDATE,'mm')) = " & Month(Session("v_Employee")) & "" & _
					           "       and to_number(to_Char(TDATE,'yyyy')) = " & Year(Session("v_Employee")) & "" & _
							   "       and fcBA = '" & Session("fcBA") & "'" & _
					           "       and SOURCETYPE in ('EM','ES')"
				connstr.execute(qNormalisasi)

				qNormalisasi = " delete COST_ACTIVITY " & _
				               " where deleteflag = 'REMOVE' " & _
					           "       and fcBA = '" & Session("fcBA") & "'" 
				connstr.execute(qNormalisasi)

				' End of Modify to make and Update State, by NAF, 3 Nov 2014
				
				'qNormalisasi = " Delete From COST_ACTIVITY" & _
				'               " where to_number(to_Char(TDATE,'mm')) = " & Month(Session("v_Employee")) & "" & _
				'	           "       and to_number(to_Char(TDATE,'yyyy')) = " & Year(Session("v_Employee")) & "" & _
				'			   "       and fcBA = '" & Session("fcBA") & "'" & _
				'	           "       and SOURCETYPE in ('EM','ES')"
				'connstr.execute(qNormalisasi)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.write("<font face=Verdana size=1 color=Green>Normalisasi Cost_Activity, Success! " & " Time: " & time() & " ..<br/><br/></font>")
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qNormalisasi&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				Response.Write "Generating Data..." & "<br/>"
				Response.Flush
				connstr.BeginTrans
				ExecuteSQL = " insert into process_empallded_daily" & _
                             " select fcba" & _
                             "        , adhocdate" & _
                             "        , allowdedcode" & _
                             "        , empcode" & _
                             "        , sum(amount) amount" & _
                             " from empallded" & _
                             " where fcba = '" & Session("fcBA") & "'" & _
                             "       and to_number(to_char(adhocdate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
                             "       and to_number(to_char(adhocdate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
                             "       and allowdedcode not in ('A18', 'A45', 'A59', 'A61')" & _
                             "       and parameter_api.get_param_2(allowdedcode, 'ALLOCATION', fcba) = 'D'" & _
                             " group by fcba, adhocdate, allowdedcode, empcode" & _
                             " having sum(amount) <> 0"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " insert into process_empallded_monthly" & _
                             " select fcba" & _
                             "        , last_day(adhocdate)adhocdate" & _
                             "        , allowdedcode" & _
                             "        , empcode" & _
                             "        , sum(amount) amount" & _
                             " from empallded" & _
                             " where fcba = '" & Session("fcBA") & "'" & _
                             "       and to_number(to_char(adhocdate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
                             "       and to_number(to_char(adhocdate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
                             "       and allowdedcode not in ('A18', 'A45', 'A59', 'A61')" & _
                             "       and parameter_api.get_param_2(allowdedcode, 'ALLOCATION', fcba) = 'M'" & _
                             " group by fcba, last_day(adhocdate), allowdedcode, empcode" & _
                             " having sum(amount) <> 0"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " insert into process_empallded_overtime" & _
                             " select fcba" & _
                             "        , adhocdate" & _
                             "        , allowdedcode" & _
                             "        , empcode" & _
                             "        , sum(amount) amount" & _
                             " from empallded" & _
                             " where fcba = '" & Session("fcBA") & "'" & _
                             "       and to_number(to_char(adhocdate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
                             "       and to_number(to_char(adhocdate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
                             "       and allowdedcode = 'A18'" & _
                             " group by fcba, adhocdate, allowdedcode, empcode" & _
                             " having sum(amount) <> 0"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				
				'**** Edit By Iwa, Remove not In KB,, cause KB get A60 and D24 in Agri Andalas **
				'ExecuteSQL = " insert into attendance_gad_closing" & _
                '             " select *" & _
                '             " from attendance_gad" & _
                '             " where fcba = '" & Session("fcBA") & "'" & _
                '             "       and to_number(to_char(fddate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
                '             "       and to_number(to_char(fddate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
                '             "       and attendance not in ('UN', 'KB', 'KL')" & _
				'			  "       and (unit <> 0 Or mandays <> 0)"
				
				'**** End Edit By Iwa, Remove not In KB,, cause KB get A60 and D24 in Agri Andalas **				
				
				ExecuteSQL = " insert into attendance_gad_closing" & _
                             " SELECT GANGCODE," & _
							"        FDDATE," & _
							"        SUPERVISION_1," & _
							"        SUPERVISION_2," & _
							"        SUPERVISION_3," & _
							"        SUPERVISION_4," & _
							"        EMPLOYEECODE," & _
							"        ATTENDANCE," & _
							"        JOBCODE," & _
							"        LOCATIONTYPE," & _
							"        LOCATIONCODE," & _
							"        MANDAYS," & _
							"        OTHRS," & _
							"        RATE," & _
							"        UNIT," & _
							"        OUTPUT," & _
							"        REFERENCE," & _
							"        REMARKS," & _
							"        FCENTRY," & _
							"        FCEDIT," & _
							"        FCIP," & _
							"        FCBA," & _
							"        LASTUPDATE," & _
							"        LASTTIME," & _
							"        LINENOKEY," & _
							"        OVERTIME_HOURS," & _
							"        TYPE_OVERTIME," & _
							"        CHARGEJOB," & _
							"        CHARGETYPE," & _
							"        CHARGECODE," & _
							"        BUCKET," & _
							"        SPBNO," & _
							"        KG_JANJANG," & _
							"        KG_BRONDOLAN" & _
							"   FROM attendance_gad" & _
                             " where fcba = '" & Session("fcBA") & "'" & _
                             "       and to_number(to_char(fddate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
                             "       and to_number(to_char(fddate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
                             "       and attendance not in ('UN', 'KL')" & _
							 "       and (unit <> 0 Or mandays <> 0)"


				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				'Add by 16 Oktober 2015 -- To save history HK KHL Monthly'
				
				connstr.BeginTrans
				ExecuteSQL = " insert into monthly_mandays_tab" & _
                             " select fcba" & _
                             "        , year" & _
                             "        , month" & _
                             "        , employeecode" & _
                             "        , employee_hist_api.get_presentgrade(year, month, fcba, employeecode) grade" & _
                             "        , total_mandays" & _
                             "        , decode(greaterthan(total_mandays, 20), 'TRUE', 1, 0) isvalidation" & _
							 "        , '" & Session("CodeName") & "' fcentry" & _
							 "        , '" & Session("CodeName") & "' fcedit" & _
							 "        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
							 "        , sysdate lastupdate" & _
							 "        , to_char(sysdate, 'HH24:MI') lasttime" & _
                             " from ( select fcba" & _
                             "               , to_number(to_char(fddate, 'yyyy')) year" & _
                             "               , to_number(to_char(fddate, 'mm')) month" & _
                             "               , employeecode" & _
                             "               , sum(mandays) total_mandays" & _
                             "        from attendance_gad_closing" & _
                             "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and attendance IN ('KJ', 'DL')" & _
                             "        group by fcba" & _
                             "                 , to_number(to_char(fddate, 'yyyy'))" & _
                             "                 , to_number(to_char(fddate, 'mm'))" & _
                             "                 , employeecode" & _
                             " ) where employee_hist_api.get_presentgrade(year, month, fcba, employeecode) = 'KHL'"


				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				'Add by 16 Oktober 2015 -- To save history HK KHL Monthly'
				
				connstr.BeginTrans
				ExecuteSQL = " insert into attendance_gad_em_daily" & _
                             " select fcba" & _
                             "        , fddate" & _
                             "        , employeecode" & _
                             "        , jobcode" & _
                             "        , locationtype" & _
                             "        , locationcode" & _
                             "        , sum(mandays) mandays" & _
                             "        , sum(unit) unit" & _
                             "        , sum(othrs) othrs" & _
                             " from attendance_gad_closing" & _
							 " where fcba = '" & Session("fcBA") & "'" & _
                             " group by fcba" & _
                             "          , fddate" & _
                             "          , employeecode" & _
                             "          , jobcode" & _
                             "          , locationtype" & _
                             "          , locationcode"


				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " insert into attendance_gad_em_monthly" & _
                             " select fcba" & _
                             "        , last_day(fddate) fddate" & _
                             "        , employeecode" & _
                             "        , jobcode" & _
                             "        , locationtype" & _
                             "        , locationcode" & _
                             "        , sum(mandays) mandays" & _
                             "        , sum(unit) unit" & _
                             "        , sum(othrs) othrs" & _
                             " from attendance_gad_closing" & _
							 " where fcba = '" & Session("fcBA") & "'" & _
                             " group by fcba" & _
                             "          , last_day(fddate)" & _
                             "          , employeecode" & _
                             "          , jobcode" & _
                             "          , locationtype" & _
                             "          , locationcode"


				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " insert into attendance_gad_es_daily" & _
                             " select fcba" & _
                             "        , fddate" & _
                             "        , employeecode" & _
                             "        , jobcode" & _
                             "        , locationtype" & _
                             "        , locationcode" & _
                             "        , sum (mandays) mandays" & _
                             "        , sum (unit) unit" & _
                             "        , sum (othrs) othrs" & _
                             " from ( select fcba" & _
                             "               , fddate" & _
                             "               , supervision_1 employeecode" & _
                             "               , jobcode" & _
                             "               , locationtype" & _
                             "               , locationcode" & _
                             "               , mandays" & _
                             "               , unit" & _
                             "               , othrs" & _
                             "        from attendance_gad_closing" & _
							 "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and supervision_1 is not null" & _
                             "        union all" & _
                             "        select fcba" & _
                             "               , fddate" & _
                             "               , supervision_2 employeecode" & _
                             "               , jobcode" & _
                             "               , locationtype" & _
                             "               , locationcode" & _
                             "               , mandays" & _
                             "               , unit" & _
                             "               , othrs" & _
                             "        from attendance_gad_closing" & _
                             "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and supervision_2 is not null" & _
                             "        union all" & _
                             "        select fcba" & _
                             "               , fddate" & _
                             "               , supervision_3 employeecode" & _
                             "               , jobcode" & _
                             "               , locationtype" & _
                             "               , locationcode" & _
                             "               , mandays" & _
                             "               , unit" & _
                             "               , othrs" & _
                             "        from attendance_gad_closing" & _
                             "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and supervision_3 is not null" & _
                             "        union all" & _
                             "        select fcba" & _
                             "               , fddate" & _
                             "               , supervision_4 employeecode" & _
                             "               , jobcode" & _
                             "               , locationtype" & _
                             "               , locationcode" & _
                             "               , mandays" & _
                             "               , unit" & _
                             "               , othrs" & _
                             "        from attendance_gad_closing" & _
                             "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and supervision_4 is not null" & _
                             "        )" & _
                             " group by fcba" & _
                             "          , fddate" & _
                             "          , employeecode" & _
                             "          , jobcode" & _
                             "          , locationtype" & _
                             "          , locationcode"


				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " insert into attendance_gad_es_monthly" & _
                             " select fcba" & _
                             "        , last_day(fddate) fddate" & _
                             "        , employeecode" & _
                             "        , jobcode" & _
                             "        , locationtype" & _
                             "        , locationcode" & _
                             "        , sum (mandays) mandays" & _
                             "        , sum (unit) unit" & _
                             "        , sum (othrs) othrs" & _
                             " from ( select fcba" & _
                             "               , fddate" & _
                             "               , supervision_1 employeecode" & _
                             "               , jobcode" & _
                             "               , locationtype" & _
                             "               , locationcode" & _
                             "               , mandays" & _
                             "               , unit" & _
                             "               , othrs" & _
                             "        from attendance_gad_closing" & _
                             "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and supervision_1 is not null" & _
                             "        union all" & _
                             "        select fcba" & _
                             "               , fddate" & _
                             "               , supervision_2 employeecode" & _
                             "               , jobcode" & _
                             "               , locationtype" & _
                             "               , locationcode" & _
                             "               , mandays" & _
                             "               , unit" & _
                             "               , othrs" & _
                             "        from attendance_gad_closing" & _
                             "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and supervision_2 is not null" & _
                             "        union all" & _
                             "        select fcba" & _
                             "               , fddate" & _
                             "               , supervision_3 employeecode" & _
                             "               , jobcode" & _
                             "               , locationtype" & _
                             "               , locationcode" & _
                             "               , mandays" & _
                             "               , unit" & _
                             "               , othrs" & _
                             "        from attendance_gad_closing" & _
                             "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and supervision_3 is not null" & _
                             "        union all" & _
                             "        select fcba" & _
                             "               , fddate" & _
                             "               , supervision_4 employeecode" & _
                             "               , jobcode" & _
                             "               , locationtype" & _
                             "               , locationcode" & _
                             "               , mandays" & _
                             "               , unit" & _
                             "               , othrs" & _
                             "        from attendance_gad_closing" & _
                             "        where fcba = '" & Session("fcBA") & "'" & _
                             "              and supervision_4 is not null" & _
                             "        )" & _
                             " group by fcba" & _
                             "          , last_day(fddate)" & _
                             "          , employeecode" & _
                             "          , jobcode" & _
                             "          , locationtype" & _
                             "          , locationcode"


				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " insert into attendance_gad_overtime" & _
                             " select fcba" & _
                             "        , fddate" & _
                             "        , employeecode" & _
                             "        , jobcode" & _
                             "        , locationtype" & _
                             "        , locationcode" & _
                             "        , sum(mandays) mandays" & _
                             "        , sum(unit) unit" & _
                             "        , sum(othrs) othrs" & _
                             " from attendance_gad" & _
							 " where fcba = '" & Session("fcBA") & "'" & _
							 "       and to_number(to_char(fddate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
                             "       and to_number(to_char(fddate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
							 "       and attendance in ('KJ', 'WH', 'WS', 'KL')" & _
							 "       and othrs <> 0" & _							 
                             " group by fcba" & _
                             "          , fddate" & _
                             "          , employeecode" & _
                             "          , jobcode" & _
                             "          , locationtype" & _
                             "          , locationcode"


				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				Response.write("<font face=Verdana size=1 color=Green>Generating Data, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
sqlunittype = " select fctype from businessunit where fccode = '" & Session("fcBA") & "'"

Set rsunittype = Server.CreateObject("ADODB.Recordset")
rsunittype.CursorType = 3
rsunittype.CursorLocation = 3
rsunittype.Open sqlunittype, connstr
if not rsunittype.eof then
   tempunittype = rsunittype("fctype")
else
   Response.Write  "<font face=Verdana size=1 color=red>" & " Tipe dari business unit " & Session("fcBA") & " belum di definisikan!"
   tempunittype = ""
end if
rsunittype.close
set rsunittype = Nothing

if tempunittype = "E" then
				
%>
	<!--#include virtual="/p_employee_closing_estate.asp"-->
<%
Else
				
				Response.Write "Allocate Gaji Borongan" & "<br/>"
				Response.Flush
				
				qBorongan = " Select fcba" & _
                            "        , employeecode" & _
                            "        , LOCATIONTYPE" & _
                            "        , LOCATIONCODE" & _
                            "        , Jobcode" & _
                            "        , FDDATE" & _
                            "        , decode(InclPayProcess,'TRUE',decode(PostingStatus,'TRUE','A45','A45'),'A45') allowdedcode" & _
                            "        , sum(totalAmount) totalAmount" & _
                            " from ( Select fcba" & _
                            "               , employeecode" & _
                            "               , LOCATIONTYPE" & _
                            "               , LOCATIONCODE" & _
                            "               , Jobcode" & _
                            "               , FDDATE" & _
                            "               , nvl(unit,0) * nvl(Rate,0) totalAmount" & _
                            "               , process_posting_api.GetPostingValue('PAYROLL', 'A45', 'JOB', jobcode, fcba) PostingStatus" & _
                            "               , allowded_api.GetIncludedPayrollProcessing('A45', fcba) InclPayProcess" & _
                            "        From ATTENDANCE_GAD" & _
                            "        Where unit <> 0" & _
                            "              and rate <> 0" & _
                            "              and (mandays = 0 or mandays is null)" & _
                            "              and Attendance = 'KB'" & _
                            "              and to_number(to_Char(FDDATE,'mm')) =  " & Month(Session("v_Employee")) & "" & _
                            "              and to_number(to_Char(FDDATE,'yyyy')) =  " & Year(Session("v_Employee")) & "" & _
                            "              and fcBA = '" & Session("fcBA") & "'" & _
							"              and employee_api.get_dateterminate(fcba, employeecode) is null" & _
                            "      )     " & _
                            " Group by fcba" & _
                            "          , employeecode" & _
                            "          , LOCATIONTYPE" & _
                            "          , LOCATIONCODE" & _
                            "          , Jobcode" & _
                            "          , FDDATE" & _
                            "          , decode(InclPayProcess,'TRUE',decode(PostingStatus,'TRUE','A45','A45'),'A45')"

				arrRS = GetRSArray(qBorongan)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
						connstr.BeginTrans
						maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
						qInsertDB=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
								  " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
								  " values ('" & arrRS(1, i) & "','EM','" & arrRS(2, i)  & "','" & _
								  arrRS(3, i)  & "','" & arrRS(4, i)  & "',"  & _
								  cdbl(arrRS(7, i) ) & "," & " to_date('" & arrRS(5, i)  & "','mm/dd/yyyy'),'" & arrRS(6, i)  & "','" & _
								  Session("CodeName") & "','" & Session("CodeName") & "','" & _
								  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',  sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
						response.Write(qInsertDB)			  
						connStr.execute(qInsertDB)

						maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1					  
						qInsertCR=  " Insert into Employeecb (EMPCODE, SOURCETYPE,TARGETTYPE, TARGETCODE, JOBCODE, UNITS, AMOUNT, TRANSDATE, REFERENCENO,FCENTRY, " & _
								   " FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
								   " values ('" & arrRS(1, i)  & "','EM','" & "CB" & "','" & _
								   "BS" & "','" & txtControlJob & "'," & 0 & ","& _
								   -1*cdbl(arrRS(7, i) )  & "," & " to_date('" & arrRS(5, i)  & "','mm/dd/yyyy'),'" & arrRS(6, i)  & "','" & _
								   Session("CodeName") & "','" & Session("CodeName") & "','" & _
								   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoEmployeeCB & ")"
						connStr.execute(qInsertCR)
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
							Response.Flush
						End If
					Next
					Erase arrRS
				End If
				
				Response.write("<font face=Verdana size=1 color=Green> Allocate Gaji Borongan, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
				
				Response.Write "Allocate No Activity" & "<br/>"
				Response.Flush
				
				qNoActivity = " select fcba" & _
							  "        , last_day(adhocdate) adhocdate" & _
							  "        , empcode" & _
							  "        , allowdedcode" & _
							  "        , sum(amount) amount" & _
							  "        , employee_closing_api.get_InfoParameterNoActivity(allowdedcode, 1, fcba) jobcode" & _
							  "        , employee_closing_api.get_InfoParameterNoActivity(allowdedcode, 2, fcba) targettype" & _
							  "        , employee_closing_api.get_InfoParameterNoActivity(allowdedcode, 3, fcba) targetcode" & _
							  "        , allowded_api.getallded(allowdedcode, fcba) allded" & _
							  " from empallded" & _
							  " where fcba = '" & Session("fcBA") & "'" & _
							  "               and to_number(to_char(adhocdate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
							  "               and to_number(to_char(adhocdate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
							  "               and employee_closing_api.get_IsNoActivity(last_day(adhocdate), empcode, fcba) = 'TRUE'" & _
							  "               and allowdedcode not in ('A45','A18', 'A59', 'A61') "  & _
							  " group by fcba" & _
							  "          , last_day(adhocdate)" & _
							  "          , empcode" & _
							  "          , allowdedcode"
							  
				arrRS = GetRSArray(qNoActivity)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
						connstr.BeginTrans
						If Cstr(arrRS(8, i)) = "A" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertDB = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , '" & arrRS(6, i) & "'" & _
										"         , '" & arrRS(7, i) & "'" & _
										"         , '" & arrRS(5, i) & "'" & _
										"         , " & cdbl(arrRS(4, i)) & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-DB-NOACT" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertDB)			  
							connStr.execute(qInsertDB)
							
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertCR = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-CR-NOACT" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertCR)			  
							connStr.execute(qInsertCR)
						ElseIf Cstr(arrRS(8, i)) = "D" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertDB = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-DB-NOACT" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertDB)			  
							connStr.execute(qInsertDB)
							
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertCR = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , '" & arrRS(6, i) & "'" & _
										"         , '" & arrRS(7, i) & "'" & _
										"         , '" & arrRS(5, i) & "'" & _
										"         , " & cdbl(arrRS(4, i)) & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-CR-NOACT" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertCR)			  
							connStr.execute(qInsertCR)
						Else
						End If

						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
							Response.Flush
						End If
					Next
					Erase arrRS
				End If
				
				Response.write("<font face=Verdana size=1 color=Green> Allocate No Activity, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
				Response.Write "Allocate Lembur" & "<br/>"
				Response.Flush
				
				qLembur = " select ed.fcba" & _
						  "        , adhocdate" & _
						  "        , empcode" & _
						  "        , allowdedcode" & _
						  "        , amount" & _
						  "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
						  " from process_empallded_overtime ed" & _
						  " inner join(select distinct fcba, fddate, employeecode" & _
						  "           from attendance_gad_overtime" & _
						  "           where fcba = '" & Session("fcBA") & "'" & _
						  "                 ) em" & _
						  " on (ed.fcba = em.fcba and ed.adhocdate = em.fddate and ed.empcode = em.employeecode)" & _
						  " where ed.fcba = '" & Session("fcBA") & "'"
						  
				arrRS = GetRSArray(qLembur)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						connstr.BeginTrans
						If Cstr(arrRS(5, i)) = "A" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertCR = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-CR" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertCR)			  
							connStr.execute(qInsertCR)
						ElseIf Cstr(arrRS(5, i)) = "D" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertDB = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-DB" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertDB)			  
							connStr.execute(qInsertDB)
						Else
						End If
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
							Response.Flush
						End If
						
					Next
					Erase arrRS
				End If
				
				qLembur = " select ed.fcba" & _
						 "        , adhocdate" & _
						 "        , empcode" & _
						 "        , allowdedcode" & _
						 "        , amount" & _
						 "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
						 "        , employee_closing_api.get_basis(adhocdate, allowdedcode, empcode, ed.fcba) basis" & _
						 "        , ag.jobcode" & _
						 "        , ag.locationtype" & _
						 "        , ag.locationcode" & _
						 "        , employee_closing_api.get_basisperactivity(adhocdate, allowdedcode, empcode, ag.jobcode, ag.locationtype, ag.locationcode, ed.fcba) basisperactivity" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.jobcode, 1, ed.fcba) jobcode_alokasi" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationtype, 2, ed.fcba) tt_alokasi" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationcode, 3, ed.fcba) tc_alokasi" & _
						 " from (select fcba" & _
						 "              , adhocdate" & _
						 "              , allowdedcode" & _
						 "              , empcode" & _
						 "              , amount" & _
						 "       from process_empallded_overtime" & _
						 "       where fcba = '" & Session("fcBA") & "'" & _
						 "             ) ed" & _
						 " inner join (select fcba" & _
						 "                    , fddate" & _
						 "                    , employeecode" & _
						 "                    , jobcode" & _
						 "                    , locationtype" & _
						 "                    , locationcode" & _
						 "                    , othrs" & _
						 "             from attendance_gad_overtime" & _
						 "             where fcba = '" & Session("fcBA") & "'" & _
						 "             ) ag" & _
						 " on (ed.fcba = ag.fcba and ed.adhocdate = ag.fddate and ed.empcode = ag.employeecode)"
						 
				arrRS = GetRSArray(qLembur)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)

						If Cdbl(arrRS(6, i)) <> 0 Then
							connstr.BeginTrans
							nAlokasi = 0
							nAlokasi = Round(Cdbl(arrRS(10, i))/Cdbl(arrRS(6, i))*Cdbl(arrRS(4, i)), 0)
							If nAlokasi <> 0 Then
								If Cstr(arrRS(5, i)) = "A" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertDB = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'EM'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-DB" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertDB)			  
									connStr.execute(qInsertDB)
								ElseIf Cstr(arrRS(5, i)) = "D" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertCR = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'EM'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-CR" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertCR)			  
									connStr.execute(qInsertCR)
								Else
								End If
							End If
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
								Response.Flush
							End If
						End If
						
					Next
					Erase arrRS
				End If
				
				Response.write("<font face=Verdana size=1 color=Green> Allocate Lembur, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
				Response.Write "Allocate Daily" & "<br/>"
				Response.Flush
				
				qDaily = " select ed.fcba" & _
						 "        , adhocdate" & _
						 "        , empcode" & _
						 "        , allowdedcode" & _
						 "        , amount" & _
						 "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
						 " from process_empallded_daily ed" & _
						 " inner join(select distinct fcba, fddate, employeecode" & _
						 "           from attendance_gad_em_daily" & _
						 "           where fcba = '" & Session("fcBA") & "'" & _
						 "           ) em" & _
						 " on (ed.fcba = em.fcba and ed.adhocdate = em.fddate and ed.empcode = em.employeecode)" & _
						 " where ed.fcba = '" & Session("fcBA") & "'" & _
						 "       and employee_closing_api.get_dailySourcetype('D', adhocdate, empcode, allowdedcode, ed.fcba) = 'EM'"
						 
				arrRS = GetRSArray(qDaily)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						connstr.BeginTrans
						If Cstr(arrRS(5, i)) = "A" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertCR = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-CR" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertCR)			  
							connStr.execute(qInsertCR)
						ElseIf Cstr(arrRS(5, i)) = "D" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertDB = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-DB" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertDB)			  
							connStr.execute(qInsertDB)
						Else
						End If
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
							Response.Flush
						End If
						
					Next
					Erase arrRS
				End If
				
				qDaily = " select ed.fcba" & _
						 "        , adhocdate" & _
						 "        , empcode" & _
						 "        , allowdedcode" & _
						 "        , amount" & _
						 "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
						 "        , employee_closing_api.get_basis(adhocdate, allowdedcode, empcode, ed.fcba) basis" & _
						 "        , ag.jobcode" & _
						 "        , ag.locationtype" & _
						 "        , ag.locationcode" & _
						 "        , employee_closing_api.get_basisperactivity(adhocdate, allowdedcode, empcode, ag.jobcode, ag.locationtype, ag.locationcode, ed.fcba) basisperactivity" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.jobcode, 1, ed.fcba) jobcode_alokasi" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationtype, 2, ed.fcba) tt_alokasi" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationcode, 3, ed.fcba) tc_alokasi" & _
						 " from process_empallded_daily ed" & _
						 " inner join attendance_gad_em_daily ag on (ed.fcba = ag.fcba and ed.adhocdate = ag.fddate and ed.empcode = ag.employeecode)" & _
						 " where ed.fcba = '" & Session("fcBA") & "'" & _
						 "       and employee_closing_api.get_dailySourcetype('D', adhocdate, empcode, allowdedcode, ed.fcba) = 'EM'"
						 
				arrRS = GetRSArray(qDaily)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						If Cdbl(arrRS(6, i)) <> 0 Then
							connstr.BeginTrans
							nAlokasi = 0
							nAlokasi = Round(Cdbl(arrRS(10, i))/Cdbl(arrRS(6, i))*Cdbl(arrRS(4, i)), 0)
							If nAlokasi <> 0 Then
								If Cstr(arrRS(5, i)) = "A" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertDB = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'EM'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-DB" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertDB)			  
									connStr.execute(qInsertDB)
								ElseIf Cstr(arrRS(5, i)) = "D" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertCR = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'EM'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-CR" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertCR)			  
									connStr.execute(qInsertCR)
								Else
								End If
							End If
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
								Response.Flush
							End If
						End If
						
					Next
					Erase arrRS
				End If
				
				
				'Supervisi Daily
				qDaily = " select ed.fcba" & _
						 "        , adhocdate" & _
						 "        , empcode" & _
						 "        , allowdedcode" & _
						 "        , amount" & _
						 "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
						 " from process_empallded_daily ed" & _
						 " inner join(select distinct fcba, fddate, employeecode" & _
						 "           from attendance_gad_es_daily" & _
						 "           where fcba = '" & Session("fcBA") & "'" & _
						 "           ) em" & _
						 " on (ed.fcba = em.fcba and ed.adhocdate = em.fddate and ed.empcode = em.employeecode)" & _
						 " where ed.fcba = '" & Session("fcBA") & "'" & _
						 "       and employee_closing_api.get_dailySourcetype('D', adhocdate, empcode, allowdedcode, ed.fcba) = 'ES'"
						 
				arrRS = GetRSArray(qDaily)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						connstr.BeginTrans
						If Cstr(arrRS(5, i)) = "A" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertCR = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'ES'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-CR" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertCR)			  
							connStr.execute(qInsertCR)
						ElseIf Cstr(arrRS(5, i)) = "D" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertDB = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'ES'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-DB" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertDB)			  
							connStr.execute(qInsertDB)
						Else
						End If
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
							Response.Flush
						End If
						
					Next
					Erase arrRS
				End If
				
				qDaily = " select ed.fcba" & _
						 "        , adhocdate" & _
						 "        , empcode" & _
						 "        , allowdedcode" & _
						 "        , amount" & _
						 "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
						 "        , employee_closing_api.get_basis_supervisi(adhocdate, allowdedcode, empcode, ed.fcba) basis" & _
						 "        , ag.jobcode" & _
						 "        , ag.locationtype" & _
						 "        , ag.locationcode" & _
						 "        , employee_closing_api.get_basisperactivity_supervisi(adhocdate, allowdedcode, empcode, ag.jobcode, ag.locationtype, ag.locationcode, ed.fcba) basisperactivity" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.jobcode, 1, ed.fcba) jobcode_alokasi" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationtype, 2, ed.fcba) tt_alokasi" & _
						 "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationcode, 3, ed.fcba) tc_alokasi" & _
						 " from process_empallded_daily ed" & _
						 " inner join attendance_gad_es_daily ag on (ed.fcba = ag.fcba and ed.adhocdate = ag.fddate and ed.empcode = ag.employeecode)" & _
						 " where ed.fcba = '" & Session("fcBA") & "'" & _
						 "       and employee_closing_api.get_dailySourcetype('D', adhocdate, empcode, allowdedcode, ed.fcba) = 'ES'"
						 
				arrRS = GetRSArray(qDaily)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						If Cdbl(arrRS(6, i)) <> 0 Then
							connstr.BeginTrans
							nAlokasi = 0
							nAlokasi = Round(Cdbl(arrRS(10, i))/Cdbl(arrRS(6, i))*Cdbl(arrRS(4, i)), 0)
							If nAlokasi <> 0 Then
								If Cstr(arrRS(5, i)) = "A" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertDB = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'ES'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-DB" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertDB)			  
									connStr.execute(qInsertDB)
								ElseIf Cstr(arrRS(5, i)) = "D" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertCR = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'ES'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-CR" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertCR)			  
									connStr.execute(qInsertCR)
								Else
								End If
							End If
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
								Response.Flush
							End If
						End If
						
					Next
					Erase arrRS
				End If
				
				Response.write("<font face=Verdana size=1 color=Green> Allocate Daily, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
				Response.Write "Allocate Monthly" & "<br/>"
				Response.Flush
				
				qMonthly = " select ed.fcba" & _
						   "        , adhocdate" & _
						   "        , empcode" & _
						   "        , allowdedcode" & _
						   "        , amount" & _
						   "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
						   " from process_empallded_monthly ed" & _
						   " inner join(select distinct fcba, fddate, employeecode" & _
						   "            from attendance_gad_em_monthly" & _
						   "            where fcba = '" & Session("fcBA") & "'" & _
						   "           ) em" & _
						   " on (ed.fcba = em.fcba and ed.adhocdate = em.fddate and ed.empcode = em.employeecode)" & _
						   " where ed.fcba = '" & Session("fcBA") & "'" & _
						   "       and employee_closing_api.get_dailySourcetype('M', adhocdate, empcode, allowdedcode, ed.fcba) = 'EM'"
						   
				arrRS = GetRSArray(qMonthly)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						connstr.BeginTrans
						If Cstr(arrRS(5, i)) = "A" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertCR = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-CR" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertCR)			  
							connStr.execute(qInsertCR)
						ElseIf Cstr(arrRS(5, i)) = "D" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertDB = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'EM'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-DB" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertDB)			  
							connStr.execute(qInsertDB)
						Else
						End If
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
							Response.Flush
						End If
						   
					Next
					Erase arrRS
				End If
				
				qMonthly = " select ed.fcba" & _
                           "        , adhocdate" & _
                           "        , empcode" & _
                           "        , allowdedcode" & _
                           "        , amount" & _
                           "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
                           "        , employee_closing_api.get_basis(adhocdate, allowdedcode, empcode, ed.fcba) basis" & _
                           "        , ag.jobcode" & _
                           "        , ag.locationtype" & _
                           "        , ag.locationcode" & _
                           "        , employee_closing_api.get_basisperactivity(adhocdate, allowdedcode, empcode, ag.jobcode, ag.locationtype, ag.locationcode, ed.fcba) basisperactivity" & _
                           "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.jobcode, 1, ed.fcba) jobcode_alokasi" & _
                           "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationtype, 2, ed.fcba) tt_alokasi" & _
                           "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationcode, 3, ed.fcba) tc_alokasi" & _
                           " from (select fcba" & _
                           "              , adhocdate" & _
                           "              , empcode" & _
                           "              , allowdedcode" & _
                           "              , amount" & _
                           "       from process_empallded_monthly" & _
						   "       where fcba = '" & Session("fcBA") & "'" & _
                           "             and employee_closing_api.get_dailySourcetype('M', adhocdate, empcode, allowdedcode, fcba) = 'EM'" & _
                           "       ) ed" & _
                           " inner join (select fcba" & _
                           "                          , fddate" & _
                           "                          , employeecode" & _
                           "                          , jobcode" & _
                           "                          , locationtype" & _
                           "                          , locationcode" & _
                           "                   from attendance_gad_em_monthly" & _
						   "                   where fcba = '" & Session("fcBA") & "'" & _
                           "             ) ag" & _
                           " on (ed.fcba = ag.fcba and ed.adhocdate = ag.fddate and ed.empcode = ag.employeecode)"
						   
				arrRS = GetRSArray(qMonthly)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						If Cdbl(arrRS(6, i)) <> 0 Then
							connstr.BeginTrans
							nAlokasi = 0
							nAlokasi = Round(Cdbl(arrRS(10, i))/Cdbl(arrRS(6, i))*Cdbl(arrRS(4, i)), 0)
							If nAlokasi <> 0 Then
								If Cstr(arrRS(5, i)) = "A" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertDB = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'EM'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-DB" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertDB)			  
									connStr.execute(qInsertDB)
								ElseIf Cstr(arrRS(5, i)) = "D" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertCR = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'EM'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-CR" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertCR)			  
									connStr.execute(qInsertCR)
								Else
								End If
							End If
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
								Response.Flush
							End If
						End If
						   
					Next
					Erase arrRS
				End If
				
				'Supervisi Monthly
				qMonthly = " select ed.fcba" & _
						   "        , adhocdate" & _
						   "        , empcode" & _
						   "        , allowdedcode" & _
						   "        , amount" & _
						   "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
						   " from process_empallded_monthly ed" & _
						   " inner join(select distinct fcba, fddate, employeecode" & _
						   "            from attendance_gad_es_monthly" & _
						   "            where fcba = '" & Session("fcBA") & "'" & _
						   "           ) em" & _
						   " on (ed.fcba = em.fcba and ed.adhocdate = em.fddate and ed.empcode = em.employeecode)" & _
						   " where ed.fcba = '" & Session("fcBA") & "'" & _
						   "       and employee_closing_api.get_dailySourcetype('M', adhocdate, empcode, allowdedcode, ed.fcba) = 'ES'"
						   
				arrRS = GetRSArray(qMonthly)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						connstr.BeginTrans
						If Cstr(arrRS(5, i)) = "A" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertCR = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'ES'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-CR" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertCR)			  
							connStr.execute(qInsertCR)
						ElseIf Cstr(arrRS(5, i)) = "D" Then
							maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
							qInsertDB = " insert into employeecb (empcode" & _
										"                         , sourcetype" & _
										"                         , targettype" & _
										"                         , targetcode" & _
										"                         , jobcode" & _
										"                         , amount" & _
										"                         , transdate" & _
										"                         , referenceno" & _
										"                         , fcentry" & _
										"                         , fcedit" & _
										"                         , fcip" & _
										"                         , fcba" & _
										"                         , lastupdate" & _
										"                         , lasttime" & _
										"                         , lineno)" & _
										" values ('" & arrRS(2, i) & "'" & _
										"         , 'ES'" & _
										"         , 'CB'" & _
										"         , 'BS'" & _
										"         , '" & txtControlJob & "'" & _
										"         , " & cdbl(arrRS(4, i)) * -1 & "" & _
										"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
										"         , '" & arrRS(3, i) & "-DB" & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Session("CodeName") & "'" & _
										"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										"         , '" & Session("fcBA") & "'" & _
										"         , sysdate" & _
										"         , to_char(sysdate, 'HH24:MI')" & _
										"         , " & maxLineNoEmployeeCB & ")"
							' response.Write(qInsertDB)			  
							connStr.execute(qInsertDB)
						Else
						End If
						
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
							Response.Flush
						End If
						   
					Next
					Erase arrRS
				End If
				
				qMonthly = " select ed.fcba" & _
                           "        , adhocdate" & _
                           "        , empcode" & _
                           "        , allowdedcode" & _
                           "        , amount" & _
                           "        , allowded_api.getallded(allowdedcode, ed.fcba) allded" & _
                           "        , employee_closing_api.get_basis_supervisi(adhocdate, allowdedcode, empcode, ed.fcba) basis" & _
                           "        , ag.jobcode" & _
                           "        , ag.locationtype" & _
                           "        , ag.locationcode" & _
                           "        , employee_closing_api.get_basisperactivity_supervisi(adhocdate, allowdedcode, empcode, ag.jobcode, ag.locationtype, ag.locationcode, ed.fcba) basisperactivity" & _
                           "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.jobcode, 1, ed.fcba) jobcode_alokasi" & _
                           "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationtype, 2, ed.fcba) tt_alokasi" & _
                           "        , GetInfoParameter(allowdedcode, ag.locationtype, ag.locationcode, 3, ed.fcba) tc_alokasi" & _
                           " from (select fcba" & _
                           "              , adhocdate" & _
                           "              , empcode" & _
                           "              , allowdedcode" & _
                           "              , amount" & _
                           "       from process_empallded_monthly" & _
						   "       where fcba = '" & Session("fcBA") & "'" & _
                           "             and employee_closing_api.get_dailySourcetype('M', adhocdate, empcode, allowdedcode, fcba) = 'ES'" & _
                           " ) ed" & _
                           " inner join (select fcba" & _
                           "                    , fddate" & _
                           "                    , employeecode" & _
                           "                    , jobcode" & _
                           "                    , locationtype" & _
                           "                    , locationcode" & _
                           "             from attendance_gad_es_monthly" & _
						   "             where fcba = '" & Session("fcBA") & "'" & _
                           "             ) ag" & _
                           " on (ed.fcba = ag.fcba and ed.adhocdate = ag.fddate and ed.empcode = ag.employeecode)"
						   
				arrRS = GetRSArray(qMonthly)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
					
						If Cdbl(arrRS(6, i)) <> 0 Then
							connstr.BeginTrans
							nAlokasi = 0
							nAlokasi = Round(Cdbl(arrRS(10, i))/Cdbl(arrRS(6, i))*Cdbl(arrRS(4, i)), 0)
							If nAlokasi <> 0 Then
								If Cstr(arrRS(5, i)) = "A" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertDB = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'ES'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-DB" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertDB)			  
									connStr.execute(qInsertDB)
								ElseIf Cstr(arrRS(5, i)) = "D" Then
									maxLineNoEmployeeCB = maxLineNoEmployeeCB + 1
									qInsertCR = " insert into employeecb (empcode" & _
												"                         , sourcetype" & _
												"                         , targettype" & _
												"                         , targetcode" & _
												"                         , jobcode" & _
												"                         , amount" & _
												"                         , transdate" & _
												"                         , referenceno" & _
												"                         , fcentry" & _
												"                         , fcedit" & _
												"                         , fcip" & _
												"                         , fcba" & _
												"                         , lastupdate" & _
												"                         , lasttime" & _
												"                         , lineno)" & _
												" values ('" & arrRS(2, i) & "'" & _
												"         , 'ES'" & _
												"         , '" & arrRS(12, i) & "'" & _
												"         , '" & arrRS(13, i) & "'" & _
												"         , '" & arrRS(11, i) & "'" & _
												"         , " & nAlokasi & "" & _
												"         , " & " to_date('" & arrRS(1, i) & "','mm/dd/yyyy')" & _
												"         , '" & arrRS(3, i) & "-CR" & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Session("CodeName") & "'" & _
												"         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
												"         , '" & Session("fcBA") & "'" & _
												"         , sysdate" & _
												"         , to_char(sysdate, 'HH24:MI')" & _
												"         , " & maxLineNoEmployeeCB & ")"
									' response.Write(qInsertCR)			  
									connStr.execute(qInsertCR)
								Else
								End If
							End If
							
							If Err.Number = 0 Then
								connstr.CommitTrans
								Response.Flush
							Else
								connstr.RollbackTrans
								Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
								Response.Flush
							End If
						End If
						   
					Next
					Erase arrRS
				End If
				
				Response.write("<font face=Verdana size=1 color=Green> Allocate Monthly, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
				Response.Write "Normalisasi Rounding" & "<br/>"
				Response.Flush
				
				qSelisih = " select fcba" & _
                           "        , transdate" & _
                           "        , empcode" & _
                           "        , allowdedcode" & _
                           "        , allded" & _
                           "        , amount" & _
                           " from (" & _
                           "     select fcba" & _
                           "            , transdate" & _
                           "            , empcode" & _
                           "            , substr (referenceno, 1, 3) allowdedcode" & _
                           "            , allowded_api.getallded(substr (referenceno, 1, 3), fcba) allded" & _
                           "            , sum (amount) amount" & _
                           "     from employeecb" & _
                           "     where fcba = '" & Session("fcBA") & "'" & _
                           "           and to_number (to_char (transdate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
                           "           and to_number (to_char (transdate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
                           "     group by fcba" & _
                           "              , transdate" & _
                           "              , empcode" & _
                           "              , substr(referenceno, 1, 3)" & _
                           "              , allowded_api.getallded(substr(referenceno, 1, 3), fcba)" & _
                           "     )" & _
                           " where amount <> 0"
						   
				arrRS = GetRSArray(qSelisih)
				
				If IsArray(arrRS) Then
					For i = LBound(arrRS, 2) To UBound(arrRS, 2)
						If Cstr(arrRS(4, i)) = "A" Then
							Set objectRS = Server.CreateObject("ADODB.Recordset")
							objectSQL = " select min (lineno) minlineno" & _
										" from employeecb" & _
										" where fcba = '" & Session("fcBA") & "'" & _
										"       and transdate = to_date('" & arrRS(1, i) & "', 'mm/dd/yyyy')" & _
										"       and empcode = '" & arrRS(2, i) & "'" & _
										"       and substr(referenceno, 1, 3) = '" & arrRS(3, i) & "'" & _
										"       and jobcode <> '" & txtControlJob & "'"
							objectRS.Open objectSQL, connstr
							If Not objectRS.EOF Then
								FirstRecord = objectRS("minlineno")
							End If
							objectRS.close
							set objectRS = Nothing
							
							UpdateAmoutSelisih = " update employeecb" & _
												 " set amount = amount - " & cdbl(arrRS(5, i)) & "" & _
												 " where lineno = '" & FirstRecord & "'" & _
												 "       and fcba = '" & Session("fcBA") & "'"
							connStr.execute(UpdateAmoutSelisih)
						ElseIf Cstr(arrRS(4, i)) = "D" Then
							Set objectRS = Server.CreateObject("ADODB.Recordset")
							objectSQL = " select min (lineno) minlineno" & _
										" from employeecb" & _
										" where fcba = '" & Session("fcBA") & "'" & _
										"       and transdate = to_date('" & arrRS(1, i) & "', 'mm/dd/yyyy')" & _
										"       and empcode = '" & arrRS(2, i) & "'" & _
										"       and substr(referenceno, 1, 3) = '" & arrRS(3, i) & "'" & _
										"       and jobcode <> '" & txtControlJob & "'"
							objectRS.Open objectSQL, connstr
							If Not objectRS.EOF Then
								FirstRecord = objectRS("minlineno")
							End If
							objectRS.close
							set objectRS = Nothing
							
							UpdateAmoutSelisih = " update employeecb" & _
												 " set amount = amount - " & cdbl(arrRS(5, i)) & "" & _
												 " where lineno = '" & FirstRecord & "'" & _
												 "       and fcba = '" & Session("fcBA") & "'"
							connStr.execute(UpdateAmoutSelisih)
						Else
						End If
					
					Next
					Erase arrRS
				End If
				
				Response.write("<font face=Verdana size=1 color=Green> Normalisasi Rounding, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
End If
				
				Response.Write "Transfer to Cost_Activity" & "<br/>"
				Response.Flush
				
				' Get the Max LineNo of Cost Activity


				Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
				rsCostActivityMax.CursorType = 3
				rsCostActivityMax.CursorLocation = 3
												
				qCostActivityMax =  " select nvl(max(nvl(lineno,0)),0) maxLineNo from cost_activity " & _
									" where fcba = '" &  Session("fcBA") & "' " 
				rsCostActivityMax.Open qCostActivityMax, connstr
				maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))

				rsCostActivityMax.close
				set rsCostActivityMax = Nothing
				' End of Get the Max LineNo of Cost Activity
				
				connStr.BeginTrans
				qABC = " insert into cost_activity (tdate" & _
					   "                            , jobcode" & _
					   "                            , source" & _
					   "                            , sourcetype" & _
					   "                            , sourcecode" & _
					   "                            , sourcecode_status" & _
					   "                            , targettype" & _
					   "                            , targetcode" & _
					   "                            , targetcode_status" & _
					   "                            , referenceno" & _
					   "                            , quantity" & _
					   "                            , output" & _
					   "                            , amount" & _
					   "                            , fcba" & _
					   "                            , closing_date" & _
					   "                            , process_date" & _
					   "                            , title_reference_01" & _
					   "                            , auto_reference_01" & _
					   "                            , title_reference_02" & _
					   "                            , auto_reference_02" & _
					   "                            , title_reference_03" & _
					   "                            , auto_reference_03" & _
					   "                            , title_reference_04" & _
					   "                            , auto_reference_04" & _
					   "                            , title_reference_05" & _
					   "                            , auto_reference_05" & _
					   "                            , debetcredit" & _
					   "                            , lastupdate" & _
					   "                            , lasttime" & _
					   "                            , lineno" & _
					   "                            , accountingperiod" & _
                       "                            , accountingyear)" & _
					   " select tdate" & _
					   "        , jobcode" & _
					   "        , source" & _
					   "        , sourcetype" & _
					   "        , sourcecode" & _
					   "        , sourcecode_status" & _
					   "        , targettype" & _
					   "        , targetcode" & _
					   "        , targetcode_status" & _
					   "        , referenceno" & _
					   "        , quantity" & _
					   "        , output" & _
					   "        , amount" & _
					   "        , fcba" & _
					   "        , closing_date" & _
					   "        , sysdate process_date" & _
					   "        , cost_activity_info_api.get_title_reference_01 (sourcetype,targettype,'Dr')" & _
					   "        , targetcode" & _
					   "        , cost_activity_info_api.get_title_reference_02 (sourcetype,targettype,'Dr')" & _
					   "        , referenceno" & _
					   "        , ''" & _
					   "        , ''" & _
					   "        , ''" & _
					   "        , ''" & _
					   "        , ''" & _
					   "        , ''" & _
					   "        , decode ( greaterorequalthan ( amount, 0), 'TRUE','Dr','Cr')" & _
					   "        , sysdate lastupdate" & _
					   "        , to_char(sysdate,'hh24:mi') lasttime" & _
					   "        , rownum + " & maxLineNo & " lineno " & _
					   "        , " & Month(Session("v_Employee")) & " accountingperiod " & _
					   "        , " & Year(Session("v_Employee")) & " accountingyear " & _
					   " from (select transdate tdate" & _
					   "              , jobcode jobcode" & _
					   "              ,  '' source" & _
					   "              , sourcetype sourcetype" & _
					   "              , employee_api.get_presentgrade(fcba, empcode) sourcecode" & _
					   "              , '' sourcecode_status" & _
					   "              , targettype targettype" & _
					   "              , targetcode targetcode" & _
					   "              , '' targetcode_status" & _
					   "              , referenceno referenceno" & _
					   "              , 0 quantity" & _
					   "              , 0 output" & _
					   "              , sum(amount) amount" & _
					   "              , fcba" & _
					   "              , to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date  " & _
					   "       from employeecb" & _
					   "       where fcba = '" & Session("fcBA") & "'" & _
					   "             and to_number(to_char(transdate,'yyyy')) = " & Year(Session("v_Employee")) & "" & _
					   "             and to_number(to_char(transdate,'mm')) = " & Month(Session("v_Employee")) & "" & _
					   "       group by transdate" & _
					   "                , jobcode" & _
					   "                , ''" & _
					   "                , sourcetype" & _
					   "                , employee_api.get_presentgrade(fcba, empcode)" & _
					   "                , ''" & _
					   "                , targettype" & _
					   "                , targetcode" & _
					   "                , ''" & _
					   "                , referenceno" & _
					   "                , fcba" & _
					   "       )"
					   
				connstr.execute(qABC)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.write("<font face=Verdana size=1 color=Green>Transfer to Cost_Activity, Success! " & " Time: " & time() & " ..<br/><br/></font>")
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qABC&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				
				Response.Write "Calculate Journal for Hutang Gaji" & "<br/>"
				Response.Flush
				
				Set rsHutangGaji = Server.CreateObject("ADODB.Recordset")
				rsHutangGaji.CursorType = 3
				rsHutangGaji.CursorLocation = 3
				
				strSQLHutangGaji = "  SELECT a.fcba," & _
								   "          LAST_DAY (a.tdate) tdate," & _
								   "          a.source," & _
								   "          a.SOURCEtype," & _
								   "          a.sourcecode," & _
								   "          a.sourcecode_status," & _
								   "          a.targettype," & _
								   "          a.targetcode," & _
								   "          a.targetcode_status," & _
								   "          SUM (amount) totalAmount" & _
								   "     FROM cost_activity a," & _
								   "          (SELECT fcba, param_1" & _
								   "             FROM parameter" & _
								   "            WHERE parcode = 'ALLOC_SALARY' AND fccode = 'CLEARINGACCPAY') b" & _
								   "    WHERE     a.jobcode = b.param_1" & _
								   "          AND a.fcba = '" & Session("fcBA") & "'" & _
								   "          AND TO_NUMBER (TO_CHAR (a.tdate, 'yyyy')) = " & Year(Session("v_Employee")) & "" & _
								   "          AND TO_NUMBER (TO_CHAR (a.tdate, 'mm')) = " & Month(Session("v_Employee")) & "" & _
								   "          AND a.fcba = b.fcba" & _
								   " GROUP BY a.fcba," & _
								   "          LAST_DAY (a.tdate)," & _
								   "          a.source," & _
								   "          a.SOURCEtype," & _
								   "          a.sourcecode," & _
								   "          a.sourcecode_status," & _
								   "          a.targettype," & _
								   "          a.targetcode," & _
								   "          a.targetcode_status"
								   
				Set rsHutangGaji = Server.CreateObject("ADODB.Recordset")					
				'response.write strSQLHutangGaji
				rsHutangGaji.Open strSQLHutangGaji, connstr

				if Not rsHutangGaji.eof then
					Do While Not rsHutangGaji.eof
						connStr.BeginTrans
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
						' End Get The Job Code for Hutang Gaji
						
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
						
						maxLineNo = maxLineNo + 1
						qInsertDB = " insert into cost_activity " & _	  
									" (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
										" SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
										" REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, closing_date, process_date, " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
										" lastupdate, lasttime, lineno, accountingperiod, accountingyear) " & _
										" values ( " & _
										" to_date('" & rsHutangGaji("tdate") & "','MM/DD/YYYY'),'" & _
										tempJobHutang & "','" & _
										rsHutangGaji("source") & _
										"','" & rsHutangGaji("SOURCETYPE") & _
										"','" & rsHutangGaji("SOURCECODE") & _
										"','" & rsHutangGaji("SOURCECODE_STATUS") & _ 
										"','" & rsHutangGaji("TARGETTYPE") & _
										"','" & rsHutangGaji("TARGETCODE") & _
										"','" & rsHutangGaji("TARGETCODE_STATUS") & _
										"','Hutang Gaji Created',0,0," &  cdbl(rsHutangGaji("totalAmount"))& ", '" & _
										Session("fcBA") & "' , to_date('" & temp_datePeriodControl & "','MM/DD/YYYY') , sysdate," & _
										" Cost_Activity_Info_Api.Get_Title_Reference_01 ('" & rsHutangGaji("SOURCETYPE") & "','" & rsHutangGaji("TARGETTYPE") & "','Dr'), '" & rsHutangGaji("TARGETCODE") & "', " & _
										" Cost_Activity_Info_Api.Get_Title_Reference_01 ('" & rsHutangGaji("SOURCETYPE") & "','" & rsHutangGaji("TARGETTYPE") & "','Dr'), 'Hutang Gaji Created', " & _
										" '','', " & _
										" '','', " & _
										" '','','Dr', " & _
										" sysdate, to_char(sysdate, 'HH24:MI') ," & maxLineNo & ", " & Month(Session("v_Employee")) & ", " & Year(Session("v_Employee")) & "  )" 
										
						  
						connStr.execute(qInsertDB)

						maxLineNo = maxLineNo + 1 
						qInsertCR = " insert into cost_activity " & _	  
									" (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
										" SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
										" REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA , closing_date, process_date, " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
										" lastupdate, lasttime, lineno, accountingperiod, accountingyear) " & _
										" values ( " & _
										" to_date('" & rsHutangGaji("tdate") & "','MM/DD/YYYY'),'" & _
										tempJobClearing & "','" & _
										rsHutangGaji("source") & _
										"','" & rsHutangGaji("SOURCETYPE") & _
										"','" & rsHutangGaji("SOURCECODE") & _
										"','" & rsHutangGaji("SOURCECODE_STATUS") & _ 
										"','" & rsHutangGaji("TARGETTYPE") & _
										"','" & rsHutangGaji("TARGETCODE") & _
										"','" & rsHutangGaji("TARGETCODE_STATUS") & _
										"','Job Clearing Created',0,0," & -1 * cdbl(rsHutangGaji("totalAmount"))& ", '" & _
										Session("fcBA") & "' , to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
										" Cost_Activity_Info_Api.Get_Title_Reference_01 ('" & rsHutangGaji("SOURCETYPE") & "','" & rsHutangGaji("TARGETTYPE") & "','Cr'), '" & rsHutangGaji("TARGETCODE") & "', " & _
										" Cost_Activity_Info_Api.Get_Title_Reference_01 ('" & rsHutangGaji("SOURCETYPE") & "','" & rsHutangGaji("TARGETTYPE") & "','Cr'), 'Job Clearing Created', " & _
										" '','', " & _
										" '','', " & _
										" '','','Cr', " & _
										" sysdate, to_char(sysdate, 'HH24:MI') ," & maxLineNo & ", " & Month(Session("v_Employee")) & ", " & Year(Session("v_Employee")) & "  )" 
										

						connStr.execute(qInsertCR)
						  
						If Err.Number = 0 Then
							connstr.CommitTrans
							Response.Flush
						Else
							connstr.RollbackTrans
							Response.write("Error Save ["&qInsertDB&"] - ["&qInsertCR&"] ("&Err.Description&")")
							Response.Flush
						End If
		  
						rsHutangGaji.MoveNext
					Loop
				End If
				rsHutangGaji.close
				set rsHutangGaji = Nothing
								   
				Response.write("<font face=Verdana size=1 color=Green> Calculate Journal for Hutang Gaji, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
				Response.Write "Update EmpalldedStatus" & "<br/>"
				Response.Flush
				connstr.BeginTrans
				'Add by NAF to handle Payroll Empallded Status
				qProcessPayroll = "BEGIN EMPALLDEDSTATUS_API.EmployeeClosing (" & Year(Session("v_Employee")) & "," & _
																				Month(Session("v_Employee")) & ",'" & _
																				Session("fcBA") & "','" & _
																				Session("CodeName") & "','" & _
																				Session("CodeName") & "','" & _
																				Request.ServerVariables("REMOTE_ADDR")  & "'); END;"
				
				connstr.execute(qProcessPayroll)
				' End of modification.
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.write("<font face=Verdana size=1 color=Green>Update EmpalldedStatus, Success! " & " Time: " & time() & " ..<br/><br/></font>")
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qProcessPayroll&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				Response.Write "Deleting Data Process" & "<br/>"
				Response.Flush
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from process_empallded_daily" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from process_empallded_monthly" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from process_empallded_overtime" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_es_monthly" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_es_daily" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_em_monthly" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_em_daily" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_overtime" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				connstr.BeginTrans
				ExecuteSQL = " delete" & _
                             " from attendance_gad_closing" & _
                             " where fcba = '" & Session("fcBA") & "'"

				connstr.execute(ExecuteSQL)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				Response.write("<font face=Verdana size=1 color=Green>Deleting Data Process, Success! " & " Time: " & time() & " ..<br/><br/></font>")
				Response.Flush
				
				
				Response.Write "Update PeriodControl" & "<br/>"
				Response.Flush
				connstr.BeginTrans
				' Add By Akbar 18 Desemeber 2015
		
			Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear  " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='EMPLOYEE' and fcba='" & Session("fcBA") & "'"
			Set rsAccountPeriod = Server.CreateObject("ADODB.Recordset")
				rsAccountPeriod.CursorType = 3
				rsAccountPeriod.CursorLocation = 3
				rsAccountPeriod.Open QAccountPeriod, connstr
				if not rsAccountPeriod.eof then
					temp_NextCloseAccountingPeriod =  Cint(rsAccountPeriod("NextCloseAccountingPeriod"))
					temp_NextCloseAccountingYear   =  Cint(rsAccountPeriod("NextCloseAccountingYear"))
					temp_NextCurrentAccountingPeriod = Cint(rsAccountPeriod("NextCurrentAccountingPeriod"))
					temp_NextCurrentAccountingYear = Cint(rsAccountPeriod("NextCurrentAccountingYear"))
				else
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for EMPLOYEE " & "<br>"
				end if
			rsAccountPeriod.close
			set rsAccountPeriod = Nothing		
			
			IF temp_NextCurrentAccountingPeriod < 12   THEN
			
				q= " Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)),CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				   " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				   " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				   " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _    
				   " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				   " where UPPER(system) ='EMPLOYEE' and fcba='" & Session("fcBA") & "'"
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where UPPER(system) ='EMPLOYEE' and fcba='" & Session("fcBA") & "'"
				ELSE
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
					   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where  UPPER(system) ='EMPLOYEE' and fcba='" & Session("fcBA") & "'"
			    END IF
			END IF
			
			IF temp_NextCloseAccountingPeriod = 13 AND temp_NextCurrentAccountingPeriod = 1  THEN
			
				q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where UPPER(system) ='EMPLOYEE' and and fcba='" & Session("fcBA") & "'"
			END IF
		
        connStr.execute(q)
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.write("<font face=Verdana size=1 color=Green>Update PeriodControl, Success! " & " Time: " & time() & " ..<br/><br/></font>")
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&qUpdatePeriod&"] ("&Err.Description&")")
					Response.Flush
				End If
				
				txtLastDay  =cmdFindPeriod("Employee")
				txtNextMonth=cmdPeriodClose("Employee")		
			
				Session("v_Employee")=txtLastDay
				Session("cls_Employee")= txtNextMonth	
				
				Response.Write "Update Closing Log" & "<br/>"
				Response.Flush
				connstr.BeginTrans
				
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
				connStr.execute(upClosingLog)
				'--End Added--
				
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.write("<font face=Verdana size=1 color=Green>Update Closing Log, Success! " & " Time: " & time() & " ..<br/><br/></font>")
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&strClosingLog&"] - ["&upClosingLog&"] ("&Err.Description&")")
					Response.Flush
				End If

				
			End If
		
		TimeAkhir = Time()
		Lama = TimeAkhir - TimeAwal
		Response.Write "<br/>" & "Finish ... " & time() & "<br/><br/>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	
%>
			<script language="JavaScript" type="text/javascript" >
				document.getElementById('status_running').style.visibility = 'hidden';
			</script>
			<br/><br/>
			<img src="images/ok.gif" /><b>Employee Closing is Completed and Success</b>
<%
		else 
%>
			<script language="JavaScript" type="text/javascript" >
				document.getElementById('status_running').style.visibility = 'hidden';
			</script>
			<br/><br/>
			<img src="images/error.gif" /><b>Employee Closing is Not Completed</b>
			<br/>
			Cause:
			<%
			tempErrorWaterFlow = "Please Check Previous Closing Process or Re-Process your Payroll, its still not DONE yet !!!"
			response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempErrorWaterFlow & "</font>"
	
		End If
		
		Dim oQuery
		oQuery	= 	" BEGIN activesession_api.DeleteSession('EMPLOYEE','" & Session("fcBA")  & "'); END; "
		connStr.execute(oQuery)	
		%>
	</body>
</html>