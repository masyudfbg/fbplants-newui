<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Workshop Closing</title>
</head>
<body>
<%
Session("v_Workshop") = cmdFindPeriod("Workshop")
Session("cls_Workshop") = cmdPeriodClose("Workshop")

Session("CurP_Workshop") = cmdFindAccountPeriod("Workshop")
Session("ClsP_Workshop") = cmdAccountPeriodClose("Workshop")

Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()

'--Added by Ukie at 25-Oct-2012--	
Dim nSelisih
Dim Startingtime, Endingtime
Dim qCostActivityMax, rsCostActivityMax, maxLineNo
Dim tempProcessCode, tempProcessNameCode, tempRefCodePVDetail, qReference
Dim rsReferencePVDetail, donePVDetail, maxLineNoWorkshopCB, qNormalisasiDetail, rsWorkshop
Dim qWorkshop, nxActivity, nxGaji, rsCostWorkshop, qCostWorkshop
Dim nCostTotal, rsBasis, nBasis, qBasis, qInsertCR, rsWSActivity
Dim qWSActivity, nWSAlokasi, qInsertDB, LT, LC, FDDATE, FCJOB, nTotalAlokasi
Dim qVoucher, rsVoucher,  txtLastDay, q
Dim txtNextMonth, qLineMax, rsqLineMax, tempMaxLine
Dim strClosingLog, strLenghtime, rsLenghtime, upClosingLog
Dim qCost_Activity, qVoucherDetail, rsVoucherDetail
Dim maxLineNoDetail
Dim txtRefStringSeparator_PVDetail, txtRefStringValues_PVDetail, nQuantity, nOutput
Dim txtRefNumberKeys_PVDetail, txtRefStringKeys_PVDetail, qCost_ActivityHistory, qCost_ActivityDBHistory
Dim temp_DebetCredit
Startingtime = FormatDateTime(CDate(Now))
'--End Added--	
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Workshop Last Period Close <%=FormatDateTime(Session("cls_Workshop"),1)%> <br> 
Workshop Current Period is <%=FormatDateTime(Session("v_Workshop"),1)%> <br><br>
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br><br>" 

dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl

' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Workshop','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Workshop " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 

	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='Workshop' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Workshop " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	' Remarked by NAF, as per New Concept, 7 May 2013
	' For Workshop Closing, for 2nd Book in Accounting

	tempProcessCode = "WORKSHOP_CLOSING"
	tempProcessNameCode = "p_Workshop_Closing.asp"
	tempRefCodePVDetail = "WORKSHOPCB"

	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' Take the Original Table of PAYMENTVOUCHERDETAIL.
	qReference = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
				  " from process_tableref_tab " & _
				  " where " & _
				  " processcode = '" & tempProcessCode & "' and " & _
				  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
				  " fccode =  '" & tempRefCodePVDetail & "'"

	Set rsReferencePVDetail = Server.CreateObject("ADODB.Recordset")
	rsReferencePVDetail.CursorType = 3
	rsReferencePVDetail.CursorLocation = 3
	rsReferencePVDetail.Open qReference, connstr

	donePVDetail = 1
	if not rsReferencePVDetail.eof then
		txtRefNumberKeys_PVDetail = rsReferencePVDetail("REF_NUMBERKEYS")
		txtRefStringSeparator_PVDetail = rsReferencePVDetail("REF_STRINGSEPARATOR")
		txtRefStringKeys_PVDetail = rsReferencePVDetail("REF_STRINGKEYS")
		donePVDetail = 1
	else
		donePVDetail = 0
	end if
	rsReferencePVDetail.close
	set rsReferencePVDetail = Nothing


	' Get the Max LineNo of Cost Activity
	Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
	rsCostActivityMax.CursorType = 3
	rsCostActivityMax.CursorLocation = 3
									
	qCostActivityMax =  " select nvl(max(nvl(lineno, 0)), 0) maxLineNo from workshopcb " & _
						" where fcba = '" &  Session("fcBA") & "' " 
	rsCostActivityMax.Open qCostActivityMax, connstr
	maxLineNoWorkshopCB = cdbl(rsCostActivityMax("maxLineNo"))

	rsCostActivityMax.close
	set rsCostActivityMax = Nothing
	' End of Get the Max LineNo of Cost Activity



	'---- Control Job
	qControlJob=" SELECT parameterdetail_api.get_param_1 ('" & Session("FCBA") & "', 'HARDCODE', 'CONTROLJOB_MODULE', 'WORKSHOP') ControlJob FROM DUAL"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For Workshop is Empty"
	   txtControlJob=""
	end if
	rsControlJob.close
	set rsControlJob = Nothing
	
	'---- No Activity
	qControlJob=" SELECT fccode FROM parameterdetail WHERE parcode = 'HARDCODE' AND fcba = '" & Session("FCBA") & "' AND parheadcode = 'JOBCODE_NOACTIVITY'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   job_noactivity = rsControlJob("fccode")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "No Job For No Activity of Workshop !"
	   job_noactivity = "2140403"
	end if
	rsControlJob.close
	set rsControlJob = Nothing

	' XXX response.Flush()

	if txtControlJob<>"" then
		'---- Normalisasi
		connstr.BeginTrans
		Response.Write "Normalisasi Workshop Base Costing" & "<br>"
		' XXX Response.Flush()
		qNormalisasi="Delete From Workshopcb where  to_number(to_Char(transdate,'mm'))=" & Month(Session("v_Workshop")) &  _
					 " and to_number(to_Char(transdate,'yyyy'))=" & Year(Session("v_Workshop")) & " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)
		connStr.execute("Commit")
		
		'---- Normalisasi
		qNormalisasi="Delete From cost_activity where Sourcetype = 'WS' and ACCOUNTINGPERIOD=" & Month(Session("v_Workshop")) &  _
					 " and ACCOUNTINGYEAR=" & Year(Session("v_Workshop")) & " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)
		connStr.execute("Commit")
		
		
		' arr_normalisasi_eof = true
		' qPeriodControl = " select lineno from cost_activity " & _
						 ' " where ACCOUNTINGPERIOD = " & Mid(Session("CurP_Workshop"),5,2) & "" & _
						 ' "       and ACCOUNTINGYEAR = " & Mid(Session("CurP_Workshop"),1,4) & "" & _
						 ' "       and fcBA = '" & Session("fcBA") & "'" & _
						 ' "       and Sourcetype = 'WS'"

		' Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
		' rsPeriodControl.Open qPeriodControl, connstr

		' if not rsPeriodControl.eof then
		  ' arr_normalisasi = rsPeriodControl.getrows()
		  ' arr_normalisasi_eof = false
		' end if
		
		' if not arr_normalisasi_eof then
			' for xyz=0 to ubound(arr_normalisasi,2)
				' qNormalisasiDetail = " DELETE FROM cost_activity_hist_tab WHERE lineheaderno = '"&arr_normalisasi(0,xyz)&"' AND fcba =  '" & Session("fcBA") & "'"
				' connstr.execute(qNormalisasiDetail)			
			' next
		' end if
		
		Set rsWorkshop = Server.CreateObject("ADODB.Recordset")
		rsWorkshop.CursorType = 3
		rsWorkshop.CursorLocation = 3
		qWorkshop = " Select * from Workshop where ACTIVATION='Y'" & _									  
					" and fcBA='" & Session("fcBA") & "'" '& _
		'Response.Write " qAllowded  " & qWorkshop & "<br>"
		'Response.Flush()
			  
		rsWorkshop.Open qWorkshop, connstr
		if not rsWorkshop.eof then
		   Do While Not rsWorkshop.eof
		      curr_vouchercode = rsWorkshop("fccode")
			  if curr_vouchercode <> prev_vouchercode then response.Write "Allocate : " & rsWorkshop("FCNAME") & "<br>" : response.Flush()

			  nxActivity =  1
			  nxGaji     = -1
		  
			  '--------------------- Total Cost of Workshop
			  qCostWorkshop = " Select nvl(sum(Amount),0) as CostTotal from COST_ACTIVITY Where to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Workshop")) &  _
						  " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Workshop")) & _ 
						  " and TargetType='WS'" & _
						  " and TargetCode='" & rsWorkshop("FCCODE") & "'" & _
						  " and fcBA='" & Session("fcBA") & "'" '& _
						  '" and FCCODE='00024'"  '--- Temporary Testing
			  Set rsCostWorkshop = Server.CreateObject("ADODB.Recordset")
			  rsCostWorkshop.Open qCostWorkshop, connstr
			  if not rsCostWorkshop.eof then
				 nCostTotal=cdbl(rsCostWorkshop("CostTotal"))
			  else
				 nCostTotal=0
			  End if   
			  rsCostWorkshop.Close
			  set rsCostWorkshop=Nothing
			  'Response.Write " Cost Total  " & nCostTotal & "<br>"
			  'Response.Flush()
			  
			  '--------------------- Find Basis by Unit
			  nBasis=0
			  qBasis =  " Select nvl(sum(UNIT),0) as Basis From WORKSHOPACTIVITY" & _
					" Where  UNIT<>0 and UNIT is not null  " & _
					" and to_number(to_Char(WORKSHOPDATE,'mm'))=" & Month(Session("v_Workshop")) &  _
					" and to_number(to_Char(WORKSHOPDATE,'yyyy'))=" & Year(Session("v_Workshop")) & _ 
					" and WORKSHOPCODE='" & rsWorkshop("FCCODE") & "'" & _
					" and fcBA='" & Session("fcBA") & "'"
					
			  Set rsBasis = Server.CreateObject("ADODB.Recordset")
			  rsBasis.Open qBasis, connstr
			  if Not rsBasis.eof then
				 nBasis=cdbl(rsBasis("Basis"))
			  Else
				 nBasis=0
			  End if
			  rsBasis.Close
			  set rsBasis=Nothing
			  'Response.Write " Basis   " & nBasis & "<br>"
			  'Response.Flush()
			  
			  if  nCostTotal<>0 and nBasis<>0 then
				  maxLineNoWorkshopCB = maxLineNoWorkshopCB + 1
				  qInsertCR=  " Insert into WORKSHOPCB (WORKSHOPCODE, TARGETTYPE, TARGETCODE, JOBCODE, QUANTITY, AMOUNT, TRANSDATE, REFERENCENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
							" values ('" & rsWorkshop("FCCODE") & "','" & "CB" & "','" & _
							"BS" & "','" & txtControlJob & "'," & 0 & "," &  nxGaji*nCostTotal & "," & _
							" to_date('" & Session("v_Workshop") & "','mm/dd/yyyy'),'" & rsWorkshop("FCCODE")+"-Cr" & "','" & _
							Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoWorkshopCB & ")"
							
				  'response.Write(qInsertCR)
				  ' XXX response.Flush()			  
				  connStr.execute(qInsertCR)		  
			  
				  '--------------------- Find Workshop Activity
				  Set rsWSActivity = Server.CreateObject("ADODB.Recordset")

				  qWSActivity = " SELECT pid, workshopcode, workshopdate,job, locationtype, locationcode, nvl(unit,0) as unit" & _
								" FROM WORKSHOPACTIVITY" & _
								" WHERE UNIT<>0 AND UNIT IS NOT NULL  " & _
								"       AND TO_NUMBER(TO_CHAR(workshopdate,'mm'))=" & Month(Session("v_Workshop")) &  _
								"       AND TO_NUMBER(TO_CHAR(workshopdate,'yyyy'))=" & Year(Session("v_Workshop")) & _ 
								"       AND workshopcode = '" & rsWorkshop("FCCODE") & "'" & _
								"       AND fcBA = '" & Session("fcBA") & "'"
					
				  rsWSActivity.Open qWSActivity, connstr
				  if not rsWSActivity.eof then
				     nTotalAlokasi=0
					 Do While Not rsWSActivity.eof	
						nWSAlokasi = Round(cdbl(rsWSActivity("Unit"))/nBasis*nCostTotal,0)*nxActivity
						
						maxLineNoWorkshopCB = maxLineNoWorkshopCB + 1			  
						qInsertDB = " INSERT INTO workshopcb ( workshopcode, targettype, targetcode, jobcode, quantity, amount, transdate, referenceno, fcentry, fcedit, fcip, fcba, lastupdate, lasttime, lineno, pid)" & _
						" values ('" & rsWorkshop("FCCODE") & "','" & rsWSActivity("LOCATIONTYPE") & "','" & _
						rsWSActivity("LOCATIONCODE") & "','" & rsWSActivity("JOB") & "'," & rsWSActivity("Unit") & "," &  nWSAlokasi & "," & _
						" to_date('" & rsWSActivity("WORKSHOPDATE") & "','mm/dd/yyyy'),'" & rsWorkshop("FCCODE")+"-Db" & "','" & _
						Session("CodeName") & "','" & Session("CodeName") & "','" & _
						Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI')," & maxLineNoWorkshopCB & ", '" & rsWSActivity("pid") & "')"
						
						'Response.Write " qInsertDB   " & qInsertDB & "<br>"
						' XXX Response.Flush()
						connStr.execute(qInsertDB)
						LT    = rsWSActivity("LOCATIONTYPE")
						LC    = rsWSActivity("LOCATIONCODE")
						FDDATE= rsWSActivity("WORKSHOPDATE")
						FCJOB = rsWSActivity("JOB")
						fpid = rsWSActivity("pid")
						
						rsWSActivity.MoveNext
						nTotalAlokasi=nTotalAlokasi+nWSAlokasi
					 Loop '----------- Loop For rsWSActivity  
					 
					 if nTotalAlokasi<>nCostTotal then
						  'response.Write "Dialokasikan " & nTotalAlokasi & "<br>"
						  'response.Write "Payslip " & cdbl(rsEmpAllDed("Amount")) & "<br>"
						  nSelisih = nCostTotal-nTotalAlokasi
						  
						  qInsertDB = " UPDATE workshopactivity SET amount = amount + '" & nSelisih & "'" & _
								      " WHERE fcba = '" & Session("fcBA") & "'" & _
								      "       AND pid = '" & fpid & "'"
						
						  connstr.execute(qInsertDB)
							
						  qInsertDB = " UPDATE workshopcb SET amount = amount + '" & nSelisih & "'" & _
										  " WHERE fcba = '" & Session("fcBA") & "'" & _
										  "       AND pid = '" & fpid & "'"
							
						  connstr.execute(qInsertDB)
						  
						  nSelisih=0
					 End if				 
					 
				  End if   
				  rsWSActivity.Close
				  set rsWSActivity = Nothing
			  Elseif nCostTotal<>0 and nBasis=0 Then
				  'Jika tidak ada activity sama sekali
				  
				  maxLineNoWorkshopCB = maxLineNoWorkshopCB + 1
				  qInsertCR=  " Insert into WORKSHOPCB (WORKSHOPCODE, TARGETTYPE, TARGETCODE, JOBCODE, QUANTITY, AMOUNT, TRANSDATE, REFERENCENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
							" values ('" & rsWorkshop("FCCODE") & "','" & "CB" & "','" & _
							"BS" & "','" & txtControlJob & "'," & 0 & "," &  nxGaji*nCostTotal & "," & _
							" to_date('" & Session("v_Workshop") & "','mm/dd/yyyy'),'" & rsWorkshop("FCCODE")+"-Cr" & "','" & _
							Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoWorkshopCB & ")"
							
				  'response.Write(qInsertCR)
				  ' XXX response.Flush()			  
				  connStr.execute(qInsertCR)
				  
				  'namun Biaya ada dicharge ke Accrued Others
				  maxLineNoWorkshopCB = maxLineNoWorkshopCB + 1
				  qInsertDB=" Insert into WORKSHOPCB (WORKSHOPCODE, TARGETTYPE, TARGETCODE, JOBCODE, QUANTITY, AMOUNT, TRANSDATE, REFERENCENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
						  " values ('" & rsWorkshop("FCCODE") & "','" & "CB" & "','" & _
						  "BS"& "','" & job_noactivity & "'," & 0 & "," &  nxActivity*nCostTotal & "," & _
						  " to_date('" & Session("v_Workshop") & "','mm/dd/yyyy'),'" & rsWorkshop("FCCODE")+"-No Act." & "','" & _
						  Session("CodeName") & "','" & Session("CodeName") & "','" & _
						  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoWorkshopCB & ")"
				  'response.Write(qInsertDB)
				  ' XXX response.Flush()			  
				  connStr.execute(qInsertDB)
			  End if ' End if Basis <> 0 and CostTotal <> 0
			  
			  prev_vouchercode = curr_vouchercode

			  rsWorkshop.MoveNext
		   loop '--------- Loop For rsWorkshop
		end if
		rsWorkshop.close
		set rsWorkshop = Nothing
		
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		%>
	<%end if

	'----- Tranfer to Activity Costing
	response.Write "Transfer to Activity Based Costing..." & "<br>"



	' Get the Max LineNo of Cost Activity
	Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
	rsCostActivityMax.CursorType = 3
	rsCostActivityMax.CursorLocation = 3
									
	qCostActivityMax =  " select nvl(max(lineno),0) maxLineNo from cost_activity " & _
						" where fcba = '" &  Session("fcBA") & "' " 
	rsCostActivityMax.Open qCostActivityMax, connstr
	maxLineNo = cdbl(to_nol(rsCostActivityMax("maxLineNo")))

	rsCostActivityMax.close
	set rsCostActivityMax = Nothing
	' End of Get the Max LineNo of Cost Activity



	' XXX response.flush()


	qVoucher= " SELECT TRANSDATE Tdate, " & _
							" JOBCODE jobcode, " & _
							" 'WS' source, " & _
							" 'WS' sourcetype, " & _
							" WORKSHOPCODE sourcecode, " & _
							" '' sourcecode_status," & _
							" TARGETTYPE, " & _
							" TARGETCODE, " & _
							" '' TARGETCODE_STATUS , " & _ 
							" REFERENCENO, " & _
							" SUM(QUANTITY) QUANTITY, " & _
							" 0 OUTPUT, " & _ 
							" SUM(AMOUNT) AMOUNT, " & _
							" FCBA , " & _
							" to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') CLOSING_DATE " & _
				  " FROM WORKSHOPCB" & _
				  " WHERE to_number(to_Char(TRANSDATE,'mm'))=" & Month(Session("v_Workshop")) &  _
				  " and to_number(to_Char(TRANSDATE,'yyyy'))=" & Year(Session("v_Workshop")) & _ 
				  " and fcBA='" & Session("fcBA") & "'" & _
				  " group by " & _
				  " TRANSDATE, JOBCODE, 'WS', 'WS', WORKSHOPCODE, '', TARGETTYPE, TARGETCODE, '', REFERENCENO,  FCBA "
				  
	Set rsVoucher = Server.CreateObject("ADODB.Recordset")
	rsVoucher.CursorType = 3
	rsVoucher.CursorLocation = 3
	rsVoucher.Open qVoucher, connstr

	if not rsVoucher.eof  then
		Do While Not rsVoucher.eof
	   
			maxLineNo = maxLineNo + 1
			
			if cdbl(rsVoucher("amount")) >= 0 then 
			    temp_DebetCredit = "Dr"
			else
			    temp_DebetCredit = "Cr"
		    end if
			
			qCost_Activity="Insert into Cost_Activity " & _
								" ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								"   TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
								"   QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, " & _
								" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								" LASTUPDATE, LASTTIME, LINENO ,journalprefix) " & _
								" values ( to_date('" & rsVoucher("tdate") & "','mm/dd/yyyy') " & ",'" & rsVoucher("jobcode")  & _ 
								"','" & rsVoucher("source") & "','" & rsVoucher("sourcetype") & "','" & rsVoucher("sourcecode") & "'" & _
								",'" & rsVoucher("sourcecode_status") & "','" & rsVoucher("targettype") & "','" & rsVoucher("targetcode") & "'" & _
								",'" & rsVoucher("targetcode_status") & "','" & rsVoucher("referenceno") & "'," & _
								cdbl(rsVoucher("quantity")) & "," & _
								cdbl(rsVoucher("output")) & "," & _
								cdbl(rsVoucher("amount")) & ",'" & _
								Session("fcBA") & "'," & _
								" to_date('" & rsVoucher("CLOSING_DATE") & "','mm/dd/yyyy') , sysdate, " & _
								" Cost_Activity_Info_Api.Get_Title_Reference_01 ('" & rsVoucher("SOURCETYPE") & "','" & rsVoucher("TARGETTYPE") & "','" & temp_DebetCredit & "'), '" & rsVoucher("TARGETCODE") & "', " & _
								" '','', " & _
								" '','', " & _
								" '','', " & _
								" '','','" & temp_DebetCredit & "', " & _
								Month(Session("v_Workshop")) & ", " & _
								Year(Session("v_Workshop")) & ", " & _
								" sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & " ,'WRS') "
			'response.write 	qCost_Activity				
			connstr.execute(qCost_Activity)

			rsVoucher.MoveNext
		loop
	End if

				qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT fcba,accountingyear,accountingperiod,referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT fcba,accountingyear,accountingperiod,sourcecode referenceno,journalprefix,company_api.get_companyno(businessunit_api.get_companycode(fcba))||'/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Workshop")) & " AND accountingyear=" & Year(Session("v_Workshop")) & " AND journalprefix='WRS' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.fcba=y.fcba and  x.sourcecode =  y.referenceno AND  x.accountingyear=y.accountingyear AND x.accountingperiod=y.accountingperiod AND x.journalprefix=y.journalprefix) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
		connstr.execute(qjournal)
		

	%>

	<%if err.number <> 0 then%>
		<br><img src="images/error.gif"><%
		connStr.RollbackTrans
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		cmdView	= "True"%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>	
	<%Else
		connStr.CommitTrans
		
		' Add By Akbar 18 Desemeber 2015
		
			Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear  " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='WORKSHOP' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for WORKSHOP " & "<br>"
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
				   " where UPPER(system) ='WORKSHOP' and fcba='" & Session("fcBA") & "'"
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _ 
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where UPPER(system) ='WORKSHOP' and fcba='" & Session("fcBA") & "'"
				ELSE
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
					   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where  UPPER(system) ='WORKSHOP' and fcba='" & Session("fcBA") & "'"
			    END IF
			END IF
			
			IF temp_NextCloseAccountingPeriod = 13 AND temp_NextCurrentAccountingPeriod = 1  THEN
			
				q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where UPPER(system) ='WORKSHOP' and and fcba='" & Session("fcBA") & "'"
			END IF
			
		connStr.execute(q)
		
		
		txtLastDay  =cmdFindPeriod("Workshop")
		txtNextMonth=cmdPeriodClose("Workshop")		
		
		Session("v_Workshop")=txtLastDay
		Session("cls_Workshop")= txtNextMonth	
		
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "<br>" & "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	
		
		'--Added by Ukie at 25-Oct-2012--	
		Endingtime = FormatDateTime(CDate(Now))

		qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Workshop'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME, fcentry) VALUES (" & _
					" 'Workshop', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" " & tempMaxLine & "," & _
					" sysdate , " & _
					" to_char(sysdate, 'HH:MI') " & _
					" ,'" & Session("CodeName") & "'" & _ 
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
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Workshop'"
				
		Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
		rsLenghtime.CursorType = 3
		rsLenghtime.CursorLocation = 3
		rsLenghtime.Open strLenghtime, connstr

		upClosingLog = " UPDATE CLOSINGLOG SET" & _
						" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Workshop'"		
		connStr.execute(upClosingLog)
		'--End Added--
		
		%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Workshop Closing is succesfully...
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
	<br><br><img src="images/error.gif">Workshop Closing is NOT completed...
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