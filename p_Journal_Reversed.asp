<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Journal Reversed</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
Dim TimeAwal,TimeAkhir,Hitung,Lama
Dim temp_DebetCredit


'**** Normalisasi Session ****'
Session("v_Reversed") = cmdFindPeriod("Reversed")
Session("cls_Reversed") = cmdPeriodClose("Reversed")
		
Session("CurP_Reversed") = cmdFindAccountPeriod("Reversed")
Session("ClsP_Reversed") = cmdAccountPeriodClose("Reversed")

Dim YearProcess
Dim MonthProcess

YearProcess = Cdbl(Left(Session("CurP_Reversed"), 4))
MonthProcess = Cdbl(Right(Session("CurP_Reversed"), 2))



TimeAwal=Time()

'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))
'--End Added--	
%>
<div id="status_running">
	<img src="ICON/working.gif">
</div>
<font face="Verdana" size="1">
<%
' Start Modification to Add Reversed Process

Response.Write "Calculate Reverse Account for TBM..." & "<br/>"
Response.Write "<strong>&nbsp;&nbsp;&nbsp;&nbsp;Unit : " & Session("fcBA")  & "</strong>"
Response.Write "<strong>&nbsp;&nbsp;|&nbsp;&nbsp;Year : " & YearProcess & "</strong>"
Response.Write "<strong>&nbsp;&nbsp;|&nbsp;&nbsp;Month : " & MonthProcess & "</strong>"
Response.Flush()

Response.Write "<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;Start ..." & Time() & "<br/><br/>"
Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please wait ..." & Time() & "<br/><br/>"
Response.Flush()

Dim strSQLReversed, rsstrSQLReversed,  rsDate
Dim tempYear, tempMonth, strDate, tempDate
Dim qInsertDB, qInsertCR

Dim GlobalRS, GlobalSQL, array_data, array_i, array_eof
Dim strJobClearing, rsstrJobClearing, tempJobDb, tempJobCr, tempJobDbPLASMA, tempJobCrPLASMA, rsCostActivityMax
Dim qCostActivityMax, maxLineNo, QAccountPeriod, rsAccountPeriod, q
Dim qLineMax, rsqLineMax, tempMaxLine, strClosingLog, rsLenghtime, strLenghtime, upClosingLog
Dim temp_jobDb, temp_jobCr

array_eof = true

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl


' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Reversed','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Reversed " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 

	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='Reversed' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Reversed " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	
	
	   ' Get The Job Code for Hutang Gaji
	   strJobClearing = " select param_1 as JobDB, param_2 as JobCR, param_3 as JobDBPLASMA, param_4 as JobCRPLASMA, param_5 as JobCRDEPINTI  from parameter " & _
							" where " & _
							" parcode = 'PROCREVERSED_001' and " & _
							" fcba = '" & Session("fcBA") & "'  and " & _
							" fccode = 'PROCESS_001' " 
							
	   Set rsstrJobClearing = Server.CreateObject("ADODB.Recordset")					
							
	   rsstrJobClearing.Open strJobClearing, connstr
	   if Not rsstrJobClearing.eof then 
		   tempJobDb = rsstrJobClearing("JobDB")
		   tempJobCr = rsstrJobClearing("JobCR")	   
		   tempJobDbPLASMA = rsstrJobClearing("JobDBPLASMA")
		   tempJobCrPLASMA = rsstrJobClearing("JobCRPLASMA")
		   tempJobCrDepInti = rsstrJobClearing("JobCRDEPINTI")
	   end if   
	   rsstrJobClearing.close
	   set rsstrJobClearing = Nothing
	
	
	   ' **** Normalisasi Job Investasi dan Alokasi TBM ****'
	   GlobalSQL = " select fcba, lineno from cost_activity" & _
                           " where fcba = '" & Session("fcBA") & "'" & _
                           "          and accountingyear = " & YearProcess & "" & _
                           "          and accountingperiod = " & MonthProcess & "" & _
                           "          and jobcode in ('" & tempJobDb & "', '" & tempJobCr & "','" & tempJobDbPLASMA & "', '" & tempJobCrPLASMA & "')" & _
						   " 		  and sourcetype not in('FA','BA')" & _
						   " 		  and targettype = 'CB'"
 
		Set GlobalRS = Server.CreateObject("ADODB.Recordset")
		GlobalRS.CursorType = 3
		GlobalRS.CursorLocation = 3
		GlobalRS.Open GlobalSQL, connstr
		if not GlobalRS.eof then		
			array_data = GlobalRS.getrows()
			array_eof = false
		end if
		GlobalRS.Close
		set GlobalRS = Nothing
		
		if not array_eof then
			For array_i = LBound(array_data, 2) To UBound(array_data, 2)
				connstr.BeginTrans
				GlobalSQL = " delete from cost_activity" & _
									" where lineno = '" & array_data(1, array_i)  & "'" & _
									"           and fcba = '" & Session("fcBA") & "'"
						
				connstr.execute(GlobalSQL)
				If Err.Number = 0 Then
					connstr.CommitTrans
					Response.Flush
				Else
					connstr.RollbackTrans
					Response.write("Error Save ["&GlobalSQL&"] ("&Err.Description&")")
					Response.Flush
				End If
				
			Next
		end if

' Get the Max LineNo of Cost Activity
	Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
	rsCostActivityMax.CursorType = 3
	rsCostActivityMax.CursorLocation = 3
												
	qCostActivityMax =  " select nvl(max(lineno),0) maxLineNo from cost_activity " & _
									" where fcba = '" &  Session("fcBA") & "' " 
	rsCostActivityMax.Open qCostActivityMax, connstr
	
	if not rsCostActivityMax.eof then
		maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))
	else
		maxLineNo = 0
	end if
	
	rsCostActivityMax.close
	set rsCostActivityMax = Nothing
' End of Get the Max LineNo of Cost Activity

' Begin Get the tempDate : the Last date in particular month
	if MonthProcess > 12 then 
		strDate = " select last_day(to_date('12/01/" & YearProcess & "','MM/DD/YYYY') ) as tempDate from dual "
	else
		strDate = " select last_day(to_date('"& MonthProcess & "/01/" & YearProcess & "','MM/DD/YYYY')) as tempDate from dual "
	end if
			   
	Set rsDate = Server.CreateObject("ADODB.Recordset")	
	'response.write  strDate  
	rsDate.Open strDate, connstr
	if Not rsDate.eof then 
		tempDate = rsDate("tempDate")
	end if  
' End Get the tempDate : the Last date in particular month		   

	' Begin INTI
		' Begin Debet
			strSQLDebetInti =   "select " & _
								 " a.source,   " & _
								 " a.Sourcetype,   " & _
								 " a.sourcecode,   " & _
								 " a.sourcecode_status,   " & _
								 " a.targettype,   " & _
								 " a.targetcode,    " & _ 
								 " a.targetcode_status,   " & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)fieldstatus, " & _
								 " sum(amount) totalAmount   " & _
								 " from" & _
								 " cost_activity a,    " & _
								 " (select fcba, param_2, param_3, param_4, param_6 from parameterdetail  " & _
								 " where  " & _
								 " parcode = 'PROCREVERSED_001' and  " & _
								 " parheadcode = 'PROCESS_001'  " & _
								 " ) b    " & _
								 " where    " & _
								 " (a.jobcode >= b.param_2 and a.jobcode <= b.param_3 or a.jobcode = b.param_4 or a.jobcode = b.param_6)  " & _
								 " and to_number(to_Char(a.tdate,'mm'))=" & MonthProcess  &  _
								 " and to_number(to_Char(a.tdate,'yyyy'))=" & YearProcess & _ 
								 " and a.fcba = '" & Session("fcBA") & "'   " & _
								 " and a.fcba = b.fcba" & _
								 " and FIELD_API.Get_owenership(a.targetcode, a.fcba) = 'INTI'" & _
								 " group by" & _
								 " a.source, a.SOURCEtype, a.sourcecode, " & _
								 " a.sourcecode_status, a.targettype, " & _
								 " a.targetcode, a.targetcode_status," & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)"
						  
			Set rsstrSQLDebetInti = Server.CreateObject("ADODB.Recordset")					
			'response.write strSQLDebetInti
			rsstrSQLDebetInti.Open strSQLDebetInti, connstr

			if Not rsstrSQLDebetInti.eof then				  
				Do While Not rsstrSQLDebetInti.eof
					maxLineNo = maxLineNo + 1
					
					temp_DebetCredit = "Dr"
					
					if rsstrSQLDebetInti("fieldstatus") = "INTI" then 
						temp_jobDb = tempJobDb
					end if

					qInsertDB = " insert into cost_activity " & _	  
								" (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
									" SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
									" REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, closing_date, process_date,  " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
										" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _								
										" LASTUPDATE, LASTTIME, LINENO ) " & _
									" values ( " & _
									" to_date('" & tempDate & "','MM/DD/YYYY'),'" & _
									temp_jobDb & "','" & _
									rsstrSQLDebetInti("source") & _
									"','" & rsstrSQLDebetInti("SOURCETYPE") & _
									"','" & rsstrSQLDebetInti("SOURCECODE") & _
									"','" & rsstrSQLDebetInti("SOURCECODE_STATUS") & _ 
									"','" & rsstrSQLDebetInti("TARGETTYPE") & _
									"','" & rsstrSQLDebetInti("TARGETCODE") & _
									"','" & rsstrSQLDebetInti("TARGETCODE_STATUS") & _
									"','TBM Reversed DB Created',0,0," &  cdbl(rsstrSQLDebetInti("totalAmount"))& ", '" & _
									Session("fcBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'Ref :','TBM Reversed DB Created', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '" & temp_DebetCredit & "', " & _
									 MonthProcess & ", " & _
									 YearProcess & ", " & _
									 " sysdate, to_char(sysdate, 'HH:MI'), " & maxLineNo & " ) "
					  
					connStr.execute(qInsertDB)
				    
					rsstrSQLDebetInti.MoveNext
				Loop
			End if
			rsstrSQLDebetInti.close
			set rsstrSQLDebetInti = Nothing
		' End Debet

		' Begin Credit Not In 99
			strSQLCreditIntiGC =  "select " & _
								 " a.source,  " & _
								 " a.Sourcetype,  " & _
								 " a.sourcecode,  " & _ 
								 " a.sourcecode_status," & _   
								 " a.targettype,   " & _
								 " a.targetcode,     " & _
								 " a.targetcode_status,   " & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)fieldstatus, " & _
								 " sum(amount) totalAmount   " & _
								 " from" & _
								 " cost_activity a,    " & _
								 " (select fcba, param_2, param_3, param_4, param_6 from parameterdetail  " & _
								 " where  " & _
								 " parcode = 'PROCREVERSED_001' and  " & _
								 " parheadcode = 'PROCESS_001'  " & _
								 " ) b    " & _
								 " where    " & _
								 " (a.jobcode >= b.param_2 and a.jobcode <= b.param_3 or a.jobcode = b.param_4 )  " & _
								 " and to_number(to_Char(a.tdate,'mm'))=" & MonthProcess  &  _
								 " and to_number(to_Char(a.tdate,'yyyy'))=" & YearProcess & _ 
								 " and a.fcba = '" & Session("fcBA") & "'   " & _
								 " and a.fcba = b.fcba" & _
								 " and FIELD_API.Get_owenership(a.targetcode, a.fcba) = 'INTI'	" & _
								 " group by" & _
								 " a.source, a.SOURCEtype, a.sourcecode, " & _
								 " a.sourcecode_status, a.targettype, " & _
								 " a.targetcode, a.targetcode_status," & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)"
						  
			Set rsstrSQLCreditIntiGC = Server.CreateObject("ADODB.Recordset")					
			'response.write strSQLCreditIntiGC
			rsstrSQLCreditIntiGC.Open strSQLCreditIntiGC, connstr

			if Not rsstrSQLCreditIntiGC.eof then				  
				Do While Not rsstrSQLCreditIntiGC.eof
					
					temp_DebetCredit = "Cr"
					
					if rsstrSQLCreditIntiGC("fieldstatus") = "INTI" then 
						temp_jobCr = tempJobCr
					end if
													
					maxLineNo = maxLineNo + 1	   
					qInsertCR = " insert into cost_activity " & _	  
								" (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
									" SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
									" REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _								
										" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
										" LASTUPDATE, LASTTIME, LINENO ) " & _
									" values ( " & _
									" to_date('" & tempDate & "','MM/DD/YYYY'),'" & _
									temp_jobCr & "','" & _
									rsstrSQLCreditIntiGC("source") & _
									"','" & rsstrSQLCreditIntiGC("SOURCETYPE") & _
									"','" & rsstrSQLCreditIntiGC("SOURCECODE") & _
									"','" & rsstrSQLCreditIntiGC("SOURCECODE_STATUS") & _ 
									"','" & rsstrSQLCreditIntiGC("TARGETTYPE") & _
									"','" & rsstrSQLCreditIntiGC("TARGETCODE") & _
									"','" & rsstrSQLCreditIntiGC("TARGETCODE_STATUS") & _
									"','TBM Reversed CR Created',0,0," & -1 * cdbl(rsstrSQLCreditIntiGC("totalAmount"))& ", '" & _
									Session("fcBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'Ref :','TBM Reversed CR Created', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '" & temp_DebetCredit & "', " & _	
									 MonthProcess & ", " & _
									 YearProcess & ", " & _							 
									 " sysdate, to_char(sysdate, 'HH:MI'), " & maxLineNo & " ) "

					connStr.execute(qInsertCR)
				    
					rsstrSQLCreditIntiGC.MoveNext
				Loop
			End if
			rsstrSQLCreditIntiGC.close
			set rsstrSQLCreditIntiGC = Nothing
		' End Credit Not In 99
		
		' Begin Credit 99 / Depresiasi 
		  strSQLCreditIntiDep =  " select " & _
								 " a.source, " & _
								 " a.Sourcetype, " & _ 
								 " a.sourcecode, " & _
								 " a.sourcecode_status, " & _  
								 " a.targettype, " & _
								 " a.targetcode, " & _
								 " a.targetcode_status, " & _  
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)fieldstatus, " & _
								 " sum(amount) totalAmount  " & _
								 " from" & _
								 " cost_activity a, " & _
								 " (select fcba, param_2, param_3, param_4, param_6 from parameterdetail " & _
								 " where " & _
								 " parcode = 'PROCREVERSED_001' and " & _
								 " parheadcode = 'PROCESS_001' " & _
								 " ) b " & _
								 " where " & _
								 " (a.jobcode >= b.param_2 and a.jobcode <= b.param_3 or a.jobcode = b.param_4 or a.jobcode = b.param_6) " & _
								 " and to_number(to_Char(a.tdate,'mm'))=" & MonthProcess  &  _
								 " and to_number(to_Char(a.tdate,'yyyy'))=" & YearProcess & _ 
								 " and a.fcba = '" & Session("fcBA") & "'   " & _
								 " and a.fcba = b.fcba " & _
								 " and a.sourcetype || a.sourcecode = 'GC00999' " & _
								 " and FIELD_API.Get_owenership(a.targetcode, a.fcba) = 'INTI' " & _
								 " group by" & _
								 " a.source, a.SOURCEtype, a.sourcecode, " & _
								 " a.sourcecode_status, a.targettype, " & _
								 " a.targetcode, a.targetcode_status," & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)"
						  
			Set rsstrSQLCreditIntiDep = Server.CreateObject("ADODB.Recordset")					
			'response.write strSQLCreditIntiDep
			rsstrSQLCreditIntiDep.Open strSQLCreditIntiDep, connstr

			if Not rsstrSQLCreditIntiDep.eof then				  
				Do While Not rsstrSQLCreditIntiDep.eof
					
					temp_DebetCredit = "Cr"
					
					if rsstrSQLCreditIntiDep("fieldstatus") = "INTI" then 
						temp_jobCr = tempJobCrDepInti
					end if
													
					maxLineNo = maxLineNo + 1	   
					qInsertCR = " insert into cost_activity " & _	  
								" (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
									" SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
									" REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _								
										" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
										" LASTUPDATE, LASTTIME, LINENO ) " & _
									" values ( " & _
									" to_date('" & tempDate & "','MM/DD/YYYY'),'" & _
									temp_jobCr & "','" & _
									rsstrSQLCreditIntiDep("source") & _
									"','" & rsstrSQLCreditIntiDep("SOURCETYPE") & _
									"','" & rsstrSQLCreditIntiDep("SOURCECODE") & _
									"','" & rsstrSQLCreditIntiDep("SOURCECODE_STATUS") & _ 
									"','" & rsstrSQLCreditIntiDep("TARGETTYPE") & _
									"','" & rsstrSQLCreditIntiDep("TARGETCODE") & _
									"','" & rsstrSQLCreditIntiDep("TARGETCODE_STATUS") & _
									"','TBM Reversed CR Created',0,0," & -1 * cdbl(rsstrSQLCreditIntiDep("totalAmount"))& ", '" & _
									Session("fcBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'Ref :','TBM Reversed CR Created', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '" & temp_DebetCredit & "', " & _	
									 MonthProcess & ", " & _
									 YearProcess & ", " & _							 
									 " sysdate, to_char(sysdate, 'HH:MI'), " & maxLineNo & " ) "

					connStr.execute(qInsertCR)
				    
					rsstrSQLCreditIntiDep.MoveNext
				Loop
			End if
			rsstrSQLCreditIntiDep.close
			set rsstrSQLCreditIntiDep = Nothing
		' End Credit 99 / Depresiasi
	' End INTI
	

	' Begin PLASMA
	  ' Begin Debet
			strSQLDebetPlasma =   "select " & _
								 " a.source, " & _
								 " a.Sourcetype, " & _
								 " a.sourcecode, " & _
								 " a.sourcecode_status, " & _
								 " a.targettype, " & _
								 " a.targetcode,  " & _
								 " a.targetcode_status, " & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)fieldstatus, " & _
								 " sum(amount) totalAmount " & _
								 " from" & _
								 " cost_activity a, " & _
								 " (select fcba, param_2, param_3, param_5 from parameterdetail  " & _
								 " where  " & _
								 " parcode = 'PROCREVERSED_001' and  " & _
								 " parheadcode = 'PROCESS_001'  " & _
								 " ) b  " & _
								 " where    " & _
								 " (a.jobcode >= b.param_2 and a.jobcode <= b.param_3 or a.jobcode = b.param_5)  " & _
						" and to_number(to_Char(a.tdate,'mm'))=" & MonthProcess  &  _
						" and to_number(to_Char(a.tdate,'yyyy'))=" & YearProcess & _ 
						" and a.fcba = '" & Session("fcBA") & "'   " & _
								 " and a.fcba = b.fcba" & _
								 " and FIELD_API.Get_owenership(a.targetcode, a.fcba) != 'INTI'" & _
								 " group by" & _
								 " a.source, a.SOURCEtype, a.sourcecode, " & _
								 " a.sourcecode_status, a.targettype, " & _
								 " a.targetcode, a.targetcode_status," & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)"
	
			Set rsstrSQLDebetPlasma = Server.CreateObject("ADODB.Recordset")					
			'response.write strSQLDebetPlasma
			rsstrSQLDebetPlasma.Open strSQLDebetPlasma, connstr

			if Not rsstrSQLDebetPlasma.eof then				  
				Do While Not rsstrSQLDebetPlasma.eof
					maxLineNo = maxLineNo + 1
					
					temp_DebetCredit = "Dr"
					
					if rsstrSQLDebetPlasma("fieldstatus") <> "INTI" then 
						temp_jobDb = tempJobDbPLASMA
					end if

					qInsertDB = " insert into cost_activity " & _	  
								" (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
									" SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
									" REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, closing_date, process_date,  " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
										" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _								
										" LASTUPDATE, LASTTIME, LINENO ) " & _
									" values ( " & _
									" to_date('" & tempDate & "','MM/DD/YYYY'),'" & _
									temp_jobDb & "','" & _
									rsstrSQLDebetPlasma("source") & _
									"','" & rsstrSQLDebetPlasma("SOURCETYPE") & _
									"','" & rsstrSQLDebetPlasma("SOURCECODE") & _
									"','" & rsstrSQLDebetPlasma("SOURCECODE_STATUS") & _ 
									"','" & rsstrSQLDebetPlasma("TARGETTYPE") & _
									"','" & rsstrSQLDebetPlasma("TARGETCODE") & _
									"','" & rsstrSQLDebetPlasma("TARGETCODE_STATUS") & _
									"','TBM Reversed DB Created',0,0," &  cdbl(rsstrSQLDebetPlasma("totalAmount"))& ", '" & _
									Session("fcBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'Ref :','TBM Reversed DB Created', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '" & temp_DebetCredit & "', " & _
									 MonthProcess & ", " & _
									 YearProcess & ", " & _
									 " sysdate, to_char(sysdate, 'HH:MI'), " & maxLineNo & " ) "
					  
					connStr.execute(qInsertDB)
				    
					rsstrSQLDebetPlasma.MoveNext
				Loop
			End if
			rsstrSQLDebetPlasma.close
			set rsstrSQLDebetPlasma = Nothing
		' End Debet

		' Begin Credit
			strSQLCreditPlasma =  "select " & _
								 " a.source," & _
								 " a.Sourcetype," & _
								 " a.sourcecode," & _
								 " a.sourcecode_status," & _
								 " a.targettype,   " & _
								 " a.targetcode,     " & _
								 " a.targetcode_status,   " & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)fieldstatus, " & _
								 " sum(amount) totalAmount   " & _
								 " from" & _
								 " cost_activity a,    " & _
								 " (select fcba, param_2, param_3, param_5 from parameterdetail  " & _
								 " where  " & _
								 " parcode = 'PROCREVERSED_001' and  " & _
								 " parheadcode = 'PROCESS_001'  " & _
								 " ) b    " & _
								 " where    " & _
								 " (a.jobcode >= b.param_2 and a.jobcode <= b.param_3 or a.jobcode = b.param_5)  " & _
						" and to_number(to_Char(a.tdate,'mm'))=" & MonthProcess  &  _
						" and to_number(to_Char(a.tdate,'yyyy'))=" & YearProcess & _ 
						" and a.fcba = '" & Session("fcBA") & "'   " & _
								 " and a.fcba = b.fcba" & _
								 " and FIELD_API.Get_owenership(a.targetcode, a.fcba) != 'INTI'" & _
								 " group by" & _
								 " a.source, a.SOURCEtype, a.sourcecode, " & _
								 " a.sourcecode_status, a.targettype, " & _
								 " a.targetcode, a.targetcode_status," & _
								 " FIELD_API.Get_owenership(a.targetcode, a.fcba)"
				
			Set rsstrSQLCreditPlasma = Server.CreateObject("ADODB.Recordset")					
			'response.write strSQLCreditPlasma
			rsstrSQLCreditPlasma.Open strSQLCreditPlasma, connstr

			if Not rsstrSQLCreditPlasma.eof then				  
				Do While Not rsstrSQLCreditPlasma.eof
					
					temp_DebetCredit = "Cr"
					if rsstrSQLCreditPlasma("fieldstatus") <> "INTI" then 
						temp_jobCr = tempJobCrPLASMA
					end if
													
					maxLineNo = maxLineNo + 1	   
					qInsertCR = " insert into cost_activity " & _	  
								" (TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " & _
									" SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _ 
									" REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
										" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
										" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
										" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
										" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
										" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _								
										" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
										" LASTUPDATE, LASTTIME, LINENO ) " & _
									" values ( " & _
									" to_date('" & tempDate & "','MM/DD/YYYY'),'" & _
									temp_jobCr & "','" & _
									rsstrSQLCreditPlasma("source") & _
									"','" & rsstrSQLCreditPlasma("SOURCETYPE") & _
									"','" & rsstrSQLCreditPlasma("SOURCECODE") & _
									"','" & rsstrSQLCreditPlasma("SOURCECODE_STATUS") & _ 
									"','" & rsstrSQLCreditPlasma("TARGETTYPE") & _
									"','" & rsstrSQLCreditPlasma("TARGETCODE") & _
									"','" & rsstrSQLCreditPlasma("TARGETCODE_STATUS") & _
									"','TBM Reversed CR Created',0,0," & -1 * cdbl(rsstrSQLCreditPlasma("totalAmount"))& ", '" & _
									Session("fcBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'Ref :','TBM Reversed CR Created', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '','', " & _
									 " '" & temp_DebetCredit & "', " & _	
									 MonthProcess & ", " & _
									 YearProcess & ", " & _							 
									 " sysdate, to_char(sysdate, 'HH:MI'), " & maxLineNo & " ) "

					connStr.execute(qInsertCR)
				    
					rsstrSQLCreditPlasma.MoveNext
				Loop
			End if
			rsstrSQLCreditPlasma.close
			set rsstrSQLCreditPlasma = Nothing
		' End Credit
	' End PLASMA
	
	' End of Modification Reversed Account
	%>

	<%if err.number <> 0 then%>
		<br><img src="images/error.gif"><%
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		cmdView	= "True"%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>	
	<%Else        
		' Add By Akbar 18 Desemeber 2015
		
			Dim temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear  " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='REVERSED' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for REVERSED " & "<br>"
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
				   " where UPPER(system) ='REVERSED' and fcba='" & Session("fcBA") & "'"
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " "  & _ 
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where UPPER(system) ='REVERSED' and fcba='" & Session("fcBA") & "'"
				ELSE
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
					   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where  UPPER(system) ='REVERSED' and fcba='" & Session("fcBA") & "'"
			    END IF
			END IF
			
			IF temp_NextCloseAccountingPeriod = 13 AND temp_NextCurrentAccountingPeriod = 1  THEN
			
				q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where UPPER(system) ='REVERSED' and and fcba='" & Session("fcBA") &  "'"
			END IF
			
		connStr.execute(q)
		
		
		Session("v_Reversed") = cmdFindPeriod("Reversed")
		Session("cls_Reversed") = cmdPeriodClose("Reversed")
		
		Session("CurP_Reversed") = cmdFindAccountPeriod("Reversed")
		Session("ClsP_Reversed") = cmdAccountPeriodClose("Reversed")
		
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "<br/>&nbsp;&nbsp;&nbsp;&nbsp;Finish ... " & time() & "<br><br>"
		Response.Write "Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	

		'--Added by Ukie at 25-Oct-2012--	
		Endingtime = FormatDateTime(CDate(Now))

		qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Reversed'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, lastupdate, lasttime) VALUES (" & _
					" 'Reversed', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" " & tempMaxLine & "" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Reversed'"
				
		Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
		rsLenghtime.CursorType = 3
		rsLenghtime.CursorLocation = 3
		rsLenghtime.Open strLenghtime, connstr

		upClosingLog = " UPDATE CLOSINGLOG SET" & _
						" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Reversed'"		
		connStr.execute(upClosingLog)
		'--End Added--	
		
	%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Journal Reversed is succesfully...
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
	<%End if%>

<%
else 
%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/error.gif">Journal Reversed is NOT completed...
	<%
    tempErrorWaterFlow = "Please Check Previous Closing Process, its still not DONE yet"
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempErrorWaterFlow & "</font>" %>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>	
	<img src="images/error.gif" />
<%
end if
%>	
	
</body>
</html>