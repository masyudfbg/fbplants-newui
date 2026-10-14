<%
   Option Explicit
%>
<%
   Response.Buffer=true
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Contract Closing</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()

Dim rsVehicle, qVehicle, nxActivity, nxGaji, nCostTotal
Dim nTotalAlokasi, rsCostVehicle, qCostVehicle, nBasis, rsBasis, qBasis, qInsertCR
Dim rsVHActivity, qVHActivity, nWSAlokasi, qInsertDB, LT, LC, FDDATE, FCJOB, nSelisih
Dim qABC, q, qLineMax, rsqLineMax, tempMaxLine
Dim strClosingLog, strLenghtime, rsLenghtime, upClosingLog, txtLastDay, txtNextMonth
Dim maxLineNoVehicleCB, tempProcessCode, tempProcessNameCode, tempRefCodePVDetail, qReference
Dim rsReferencePVDetail, donePVDetail, qNormalisasiDetail, qVoucher, rsVoucher, qCost_Activity,qContractInvDB1, qContractInvDB2, qContractInvCR1, qContractInvCR2
Dim qVoucherDetail, rsVoucherDetail, maxLineNoDetail, rsContractInvDB1, rsContractInvDB2, rsContractInvCR1, rsContractInvCR2
Dim txtRefStringSeparator_PVDetail, txtRefStringValues_PVDetail, nQuantity, nOutput, qCost_ActivityTRDB1, qCost_ActivityTRDB2, qCost_ActivityTRCR1, qCost_ActivityTRCR2 
Dim txtRefNumberKeys_PVDetail, txtRefStringKeys_PVDetail, qCost_ActivityHistory, maxLineNoContractorCB,maxLineNoINT 
Dim rsContractDB, qContractDB, rsContractCR, qContractCR, rsAggrement, qCostContract, nRate,qjournal
Dim temp_DebetCredit


'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))
'--End Added--	
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Contract Last Period Close <%=FormatDateTime(Session("cls_Contract"),1)%> <br> 
Contract Current Period is <%=FormatDateTime(Session("v_Contract"),1)%> <br><br>
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br><br>" 

dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode
Dim NormalisasiContractorLedger, SelectContractorLedger, rsContractorLedger, InsertContractorLedger ,vchar
Dim qCostActivityMax, rsCostActivityMax, maxLineNo

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted



' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Contract','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Contract " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 


	Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl

	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='Contract' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Contract " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	' Remarked by NAF, as per New Concept, 7 May 2013
	' For Workshop Closing, for 2nd Book in Accounting

	tempProcessCode = "CONTRACT_CLOSING"
	tempProcessNameCode = "p_Contract_Closing.asp"
	tempRefCodePVDetail = "CONTRACTORCB"

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
									
	qCostActivityMax =  " select  nvl(max(lineno),0) maxLineNo from contractorcb " & _
						" where fcba = '" &  Session("fcBA") & "' " 
	rsCostActivityMax.Open qCostActivityMax, connstr
	maxLineNoContractorCB = cdbl(rsCostActivityMax("maxLineNo"))

	rsCostActivityMax.close
	set rsCostActivityMax = Nothing
	' End of Get the Max LineNo of Cost Activity



	'---- Control Job
	qControlJob="select ControlJob from Module where NAME='Contract'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For Contract is Empty"
	   txtControlJob=""
	end if
	rsControlJob.close
	set rsControlJob = Nothing

	response.Flush()

	' if txtControlJob<>"" then ' Remarked By UKIE 21 feb 2014 purpose: Controljob from Contractor Master Not from Module
		'---- Normalisasi
		connstr.BeginTrans
		Response.Write " Normalisasi Contract Base Costing" & "<br>"
		Response.Flush()
		qNormalisasi="Delete From CONTRACTORCB where  to_number(to_Char(transdate,'mm'))=" & Month(Session("v_Contract")) &  _
					 " and to_number(to_Char(transdate,'yyyy'))=" & Year(Session("v_Contract")) & " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)
		connStr.execute("Commit")
		
		'---- Normalisasi

		' qNormalisasiDetail = " delete from cost_activity_hist_tab " & _
							 ' " where " & _
							 ' " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
									' " where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Contract")) &  _
									' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Contract")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='WC')"
							 
		' connstr.execute(qNormalisasiDetail)	
		
		' connStr.execute("Commit")
		
		' qNormalisasi=" Delete From COST_ACTIVITY where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Contract")) &  _
					 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Contract")) & " and fcBA='" & Session("fcBA") & "'" & _
					 ' " and SOURCETYPE = 'WC'"
					 
		' connstr.execute(qNormalisasi)	
		
		' connStr.execute("Commit")
		
		
		qNormalisasiDetail = " update cost_activity_hist_tab set deleteflag = 'REMOVE'" & _
							 " where " & _
							 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
									" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Contract")) &  _
									" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Contract")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='WC')"
							 
		connstr.execute(qNormalisasiDetail)	
		
		connStr.execute("Commit")
		
		qNormalisasiDetail = " delete from cost_activity_hist_tab where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

		connstr.execute(qNormalisasiDetail)
		connStr.execute("Commit")
		
		
		
		qNormalisasiDetail=	 " DELETE FROM  Cost_Activity WHERE  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Contract")) & _
							 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Contract")) & _
							 " and (fcBA='" & Session("fcBA") & "')" & _
							 " and SOURCETYPE = 'WC' AND journalprefix='WIP'"
							 
		connstr.execute(qNormalisasiDetail)	
		
		connStr.execute("Commit")
		
		qNormalisasiDetail=	 " DELETE FROM  Cost_Activity WHERE  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Contract")) & _
							 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Contract")) & _
							 " and (fcBA='" & Session("fcBA") & "' or chargeunit='" & Session("fcBA") & "')" & _
							 " and journalprefix='WIV'"
							 
		connstr.execute(qNormalisasiDetail)	
		
		connStr.execute("Commit")
		
		'qNormalisasiDetail = " delete cost_activity where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

		'connstr.execute(qNormalisasiDetail)
		'connStr.execute("Commit")
		
		
		'if Session("Estate_Mill") <> "HO" then
			' Get the Max LineNo of Cost Activity
			'-- AMH 05-22 Modify to 4 contract Unit 
			' JIka FCBA = CHARGEUNIT
		 	'-- Journal Db : Biaya Unit
			'-- Journal Cr : Clearing Accounting
			' JIka FCBA <> CHARGEUNIT
			'-- FCBA = FCBA 
		 	'-- Journal Db : INTRACO 
			'-- Journal Cr : Clearing Accounting
			'-- Charge Unit: CHARGEUNIT
			
			'-- CHARGEUNIT = FCBA 
		 	'-- Journal Db : BIAYA Unit 
			'-- Journal Cr : INTRACO
			'-- Charge Unit: FCBA
			
			
			Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
			rsCostActivityMax.CursorType = 3
			rsCostActivityMax.CursorLocation = 3
			qCostActivityMax =  " select max(lineno) maxLineNo from cost_activity " & _
								" where fcba = '" &  Session("fcBA") & "' " 
			rsCostActivityMax.Open qCostActivityMax, connstr
			maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))
			rsCostActivityMax.close
			set rsCostActivityMax = Nothing
			
			response.Write "Process WIP Contract BEGIN.... <br>"
		    response.Flush
			
			' JIka FCBA = CHARGEUNIT
			Set rsContractDB = Server.CreateObject("ADODB.Recordset")
			rsContractDB.CursorType = 3
			rsContractDB.CursorLocation = 3
			qContractDB = " SELECT "& _
			             " TO_CHAR(LVL1.WIPDATE,'fmMM') ACCOUNTINGPERIOD, "& _
						 " TO_CHAR(LVL1.WIPDATE,'yyyy') ACCOUNTINGYEAR, "& _
						 " LVL1.AGREEMENTCODE AGREEMENTCODE, LVL1.WIPDATE WIPDATE, CA.CONTRACTORCODE CONTRACTORCODE, "& _
						 " Contractor_api.get_controljob(CA.CONTRACTORCODE, CA.FCBA) controljob,"& _
						 " PAR.param_7 controljobtemp,  "& _
						 " CA.DESCRIPTION CONTRACTORDESC, LVL1.JOBCODE JOB, LVL1.LOCATIONTYPE LOCATIONTYPE, LVL1.LOCATIONCODE LOCATIONCODE, "& _
						 " SUM(LVL1.QTY * CAD.RATE) AMOUNT, "& _
						 " SUM(LVL1.QTY) QTY, "& _
						 " job_api.get_jobcode_name(LVL1.JOBCODE , CA.FCBA) JOBNAME,"& _
						 " contractor_api.get_contractorname(CA.CONTRACTORCODE, CA.FCBA) CONTRACTORNAME,LVL1.FCBA "& _
						 " FROM( "& _
						 " Select "& _
						 " HDR.FCBA, HDR.WIPDATE, HDR.AGREEMENTCODE,  DTL.JOBCODE, HDR.AGREEMENTLINENO,"& _
						 " DTL.LOCATIONTYPE, DTL.LOCATIONCODE, SUM(HDR.QTY) QTY"& _
						 " FROM "& _
						 " WORKINPROGRESSPLAN DTL "& _
						 " INNER JOIN WORKINPROGRESS HDR ON (DTL.AGREEMENTCODE = HDR.AGREEMENTCODE AND "& _
						 " DTL.FCBA = HDR.FCBA AND HDR.AGREEMENTFINALNO = DTL.AGREEMENTFINALNO AND HDR.AGREEMENTLINENO = DTL.AGREEMENTLINENO "& _
						 " AND DTL.FCBA=DTL.CHARGEUNIT)  "& _
						 " WHERE "& _
						 " HDR.FCBA='" & Session("fcBA") & "'" & _
						 " and to_number(to_Char(HDR.WIPDATE,'mm'))=" & Month(Session("v_Contract")) &  _
						 " and to_number(to_Char(HDR.WIPDATE,'yyyy'))=" & Year(Session("v_Contract")) &  _						
 						 " GROUP BY "& _
						 " HDR.FCBA, HDR.WIPDATE, HDR.AGREEMENTCODE,DTL.JOBCODE, HDR.AGREEMENTLINENO,DTL.LOCATIONTYPE, DTL.LOCATIONCODE "& _
						 " )LVL1 "& _
						 " INNER JOIN CONTRACTAGREEMENTDETAIL CAD ON LVL1.AGREEMENTCODE = CAD.AGREEMENTCODE AND LVL1.FCBA = CAD.FCBA AND  LVL1.AGREEMENTLINENO = CAD.LINENO "& _
						 " INNER JOIN CONTRACTAGREEMENT CA ON CAD.AGREEMENTCODE = CA.AGREEMENTCODE AND CAD.FCBA = CA.FCBA "& _
						 " LEFT JOIN PARAMETER PAR ON CAD.FCBA = PAR.FCBA AND parcode = 'JOBCODE_CONTRACTOR_INVOICE'  "& _
						 " GROUP BY CA.FCBA, LVL1.AGREEMENTCODE, LVL1.WIPDATE, CA.CONTRACTORCODE, Contractor_api.get_controljob(CA.CONTRACTORCODE, CA.FCBA), "& _
						 " CA.DESCRIPTION,  LVL1.JOBCODE, LVL1.LOCATIONTYPE, LVL1.LOCATIONCODE,PAR.param_7,LVL1.FCBA"& _
						 " ORDER BY agreementcode,job,locationcode "
		
			rsContractDB.Open qContractDB, connstr
			if not rsContractDB.eof then
			   Do While Not rsContractDB.eof
				  Response.Write "FCBA = CHARGEUNIT to Allocate the  " & rsContractDB("AGREEMENTCODE") & "<br>"
				  Response.Flush()		

				  '-- UnitWIP 
					 maxLineNo = maxLineNo + 1
					 qCost_ActivityTRDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy') " & ",'" & rsContractDB("JOB") & "','WC','WC','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '','" & rsContractDB("LOCATIONTYPE") & "','" & rsContractDB("LOCATIONCODE") & "'," & _
									 " '" &  cdbl(rsContractDB("AMOUNT")) & "','" & rsContractDB("AGREEMENTCODE") & "'," & rsContractDB("QTY") &",0,'" & rsContractDB("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '',''," & _
									 " '',''," & _
									 " '',''," & _
									 " '','No.Kontrak/Kontraktor: " & rsContractDB("AGREEMENTCODE") & "/" & rsContractDB("CONTRACTORNAME") &" Pekerjaan: " & rsContractDB("JOBNAME") & "/" & rsContractDB("LOCATIONCODE")  & "', " & _
									 " 'JOB-BLOK', '"& rsContractDB("JOB") &"-" & rsContractDB("LOCATIONCODE") &"','Db', " & _
									 " " & rsContractDB("ACCOUNTINGPERIOD") & "," & rsContractDB("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIP') "
					connstr.execute(qCost_ActivityTRDB1)
					 
					maxLineNoContractorCB = maxLineNoContractorCB + 1
					qInsertDB=  " Insert into CONTRACTORCB (CONTRACTORCODE, TARGETTYPE, TARGETCODE, JOBCODE, " & _
								" QUANTITY, AMOUNT, TRANSDATE, REFERENCENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
							    " values ('" & rsContractDB("CONTRACTORCODE") & "','" & rsContractDB("LOCATIONTYPE") & "','" & _
							    rsContractDB("LOCATIONCODE") & "','" & rsContractDB("JOB") & "'," & rsContractDB("QTY") & "," &  cdbl(rsContractDB("AMOUNT")) & "," & _
							    " to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy'),'" & rsContractDB("AGREEMENTCODE")+"-Db" & "','" & _
							    Session("CodeName") & "','" & Session("CodeName") & "','" & _
							    Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoContractorCB & " )"
					
					response.Flush()			  
					connStr.execute(qInsertDB)	
					  
					maxLineNo = maxLineNo + 1
					qCost_ActivityTRCR1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy') " & ",'" & rsContractDB("controljobtemp") & "','WC','WC','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '','CA','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '" &  cdbl(rsContractDB("AMOUNT")) *-1 & "','" & rsContractDB("AGREEMENTCODE") & "'," & rsContractDB("QTY") &",0,'" & rsContractDB("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '',''," & _
									 " '',''," & _
									 " '',''," & _
									 " '','No.Kontrak/Kontraktor: " & rsContractDB("AGREEMENTCODE") & "/" & rsContractDB("CONTRACTORNAME") &" Pekerjaan: " & rsContractDB("JOBNAME") & "/" & rsContractDB("LOCATIONCODE") & "', " & _
									 " 'JOB-BLOK', '"& rsContractDB("JOB") &"-" & rsContractDB("LOCATIONCODE") &"','Cr', " & _
									 " " & rsContractDB("ACCOUNTINGPERIOD") & "," & rsContractDB("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIP') "
					connstr.execute(qCost_ActivityTRCR1)
									  
				 	  maxLineNoContractorCB = maxLineNoContractorCB + 1
					  qInsertCR=  " Insert into CONTRACTORCB (CONTRACTORCODE, TARGETTYPE, TARGETCODE, JOBCODE, " & _
								  " QUANTITY, AMOUNT, TRANSDATE, REFERENCENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
								" values ('" & rsContractDB("CONTRACTORCODE") & "','CA','" & rsContractDB("CONTRACTORCODE") & "','" & rsContractDB("controljobtemp") & "'," & 0 & "," & cdbl(rsContractDB("AMOUNT"))*-1 & "," & _
								" to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy'),'"  & rsContractDB("AGREEMENTCODE")+"-Cr" & "','" & _
								Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") &  "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoContractorCB & " )"
					  response.Flush()			  
					  connStr.execute(qInsertCR)  

            	  
				  rsContractDB.MoveNext
			   loop '--------- Loop For rsContractDB
			end if
			rsContractDB.close
			set rsContractDB = Nothing
			
			
			' JIka FCBA <> CHARGEUNIT [FCBA=FCBA]
			Set rsContractDB = Server.CreateObject("ADODB.Recordset")
			rsContractDB.CursorType = 3
			rsContractDB.CursorLocation = 3
			qContractDB = " SELECT "& _
			             " TO_CHAR(LVL1.WIPDATE,'fmMM') ACCOUNTINGPERIOD, "& _
						 " TO_CHAR(LVL1.WIPDATE,'yyyy') ACCOUNTINGYEAR, "& _
						 " LVL1.AGREEMENTCODE AGREEMENTCODE, LVL1.WIPDATE WIPDATE, CA.CONTRACTORCODE CONTRACTORCODE, "& _
						 " businessunit_api.get_jobintraco( CA.FCBA) controljob,"& _
						 " PAR.param_7 controljobtemp,  "& _
						 " CA.DESCRIPTION CONTRACTORDESC, LVL1.JOBCODE JOB, LVL1.LOCATIONTYPE LOCATIONTYPE, LVL1.LOCATIONCODE LOCATIONCODE, "& _
						 " SUM(LVL1.QTY * CAD.RATE) AMOUNT, "& _
						 " SUM(LVL1.QTY) QTY, "& _
						 " job_api.get_jobcode_name(LVL1.JOBCODE , CA.FCBA) JOBNAME,"& _
						 " contractor_api.get_contractorname(CA.CONTRACTORCODE, CA.FCBA) CONTRACTORNAME,LVL1.FCBA,LVL1.CHARGEUNIT"& _
						 " FROM( "& _
						 " Select "& _
						 " HDR.FCBA, HDR.WIPDATE, HDR.AGREEMENTCODE,  DTL.JOBCODE, HDR.AGREEMENTLINENO,"& _
						 " DTL.LOCATIONTYPE, DTL.LOCATIONCODE, SUM(HDR.QTY) QTY, DTL.CHARGEUNIT"& _
						 " FROM "& _
						 " WORKINPROGRESSPLAN DTL "& _
						 " INNER JOIN WORKINPROGRESS HDR ON DTL.AGREEMENTCODE = HDR.AGREEMENTCODE AND "& _
						 " DTL.FCBA = HDR.FCBA AND HDR.AGREEMENTFINALNO = DTL.AGREEMENTFINALNO AND HDR.AGREEMENTLINENO = DTL.AGREEMENTLINENO "& _
						 " AND DTL.FCBA <> DTL.CHARGEUNIT  "& _
						 " WHERE "& _
						 " HDR.FCBA='" & Session("fcBA") & "'" & _
						 " and to_number(to_Char(HDR.WIPDATE,'mm'))=" & Month(Session("v_Contract")) &  _
						 " and to_number(to_Char(HDR.WIPDATE,'yyyy'))=" & Year(Session("v_Contract")) &  _
						 " GROUP BY "& _
						 " HDR.FCBA, HDR.WIPDATE, HDR.AGREEMENTCODE,DTL.JOBCODE, HDR.AGREEMENTLINENO,DTL.LOCATIONTYPE, DTL.LOCATIONCODE,DTL.CHARGEUNIT "& _
						 " )LVL1 "& _
						 " INNER JOIN CONTRACTAGREEMENTDETAIL CAD ON LVL1.AGREEMENTCODE = CAD.AGREEMENTCODE AND LVL1.FCBA = CAD.FCBA AND  LVL1.AGREEMENTLINENO = CAD.LINENO "& _
						 " INNER JOIN CONTRACTAGREEMENT CA ON CAD.AGREEMENTCODE = CA.AGREEMENTCODE AND CAD.FCBA = CA.FCBA "& _
						 " LEFT JOIN PARAMETER PAR ON CAD.FCBA = PAR.FCBA AND parcode = 'JOBCODE_CONTRACTOR_INVOICE'  "& _
						 " GROUP BY CA.FCBA, LVL1.AGREEMENTCODE, LVL1.WIPDATE, CA.CONTRACTORCODE, Contractor_api.get_controljob(CA.CONTRACTORCODE, CA.FCBA), "& _
						 " CA.DESCRIPTION,  LVL1.JOBCODE, LVL1.LOCATIONTYPE, LVL1.LOCATIONCODE,PAR.param_7,LVL1.FCBA,LVL1.CHARGEUNIT "& _
						 " ORDER BY agreementcode,chargeunit,job,locationcode "
		
			rsContractDB.Open qContractDB, connstr
			if not rsContractDB.eof then
			   Do While Not rsContractDB.eof
				  Response.Write "FCBA <> CHARGEUNIT [FCBA=FCBA] to Allocate the  " & rsContractDB("AGREEMENTCODE") & "- " &  rsContractDB("CHARGEUNIT") & "<br>"
				  Response.Flush()		

				  '-- UnitWIP 
					 maxLineNo = maxLineNo + 1
					 qCost_ActivityTRDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix,CHARGEUNIT ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy') " & ",'" & rsContractDB("controljob") & "','WC','WC','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '','CB','BS'," & _
									 " '" &  cdbl(rsContractDB("AMOUNT")) & "','" & rsContractDB("AGREEMENTCODE") & "'," & rsContractDB("QTY") &",0,'" & rsContractDB("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '',''," & _
									 " '',''," & _
									 " 'INTRACO TO ','" & rsContractDB("CHARGEUNIT")  & "'," & _
									 " '','No.Kontrak/Kontraktor: " & rsContractDB("AGREEMENTCODE") & "/" & rsContractDB("CONTRACTORNAME") &" Pekerjaan: " & rsContractDB("JOBNAME") & "/" & rsContractDB("LOCATIONCODE") & "', " & _
									 " 'JOB-BLOK', '"& rsContractDB("JOB") &"-" & rsContractDB("LOCATIONCODE") &"','Db', " & _
									 " " & rsContractDB("ACCOUNTINGPERIOD") & "," & rsContractDB("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIP','" & rsContractDB("CHARGEUNIT") & "') "
					connstr.execute(qCost_ActivityTRDB1)
					 
					maxLineNoContractorCB = maxLineNoContractorCB + 1
					qInsertDB=  " Insert into CONTRACTORCB (CONTRACTORCODE, TARGETTYPE, TARGETCODE, JOBCODE, " & _
								" QUANTITY, AMOUNT, TRANSDATE, REFERENCENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
							    " values ('" & rsContractDB("CONTRACTORCODE") & "','CB','BS','" & rsContractDB("controljob") & "'," & rsContractDB("QTY") & "," &  cdbl(rsContractDB("AMOUNT")) & "," & _
							    " to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy'),'" & rsContractDB("AGREEMENTCODE")+"-Db" & "','" & _
							    Session("CodeName") & "','" & Session("CodeName") & "','" & _
							    Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoContractorCB & " )"
					
					response.Flush()			  
					connStr.execute(qInsertDB)	
					  
					maxLineNo = maxLineNo + 1
					qCost_ActivityTRCR1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy') " & ",'" & rsContractDB("controljobtemp") & "','WC','WC','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '','CA','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '" &  cdbl(rsContractDB("AMOUNT")) *-1 & "','" & rsContractDB("AGREEMENTCODE") & "'," & rsContractDB("QTY") &",0,'" & rsContractDB("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '',''," & _
									 " '',''," & _
									 " 'INTRACO TO ','" & rsContractDB("CHARGEUNIT")  & "'," & _
									 " '','No.Kontrak/Kontraktor: " & rsContractDB("AGREEMENTCODE") & "/" & rsContractDB("CONTRACTORNAME") &" Pekerjaan: " & rsContractDB("JOBNAME") & "/" & rsContractDB("LOCATIONCODE") & "', " & _
									 " 'JOB-BLOK', '"& rsContractDB("JOB") &"-" & rsContractDB("LOCATIONCODE") &"','Cr', " & _
									 " " & rsContractDB("ACCOUNTINGPERIOD") & "," & rsContractDB("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIP') "
					  connstr.execute(qCost_ActivityTRCR1)
									  
				 	  maxLineNoContractorCB = maxLineNoContractorCB + 1
					  qInsertCR=  " Insert into CONTRACTORCB (CONTRACTORCODE, TARGETTYPE, TARGETCODE, JOBCODE, " & _
								  " QUANTITY, AMOUNT, TRANSDATE, REFERENCENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, LINENO)" & _
								" values ('" & rsContractDB("CONTRACTORCODE") & "','CA','" & rsContractDB("CONTRACTORCODE") & "','" & rsContractDB("controljobtemp") & "'," & 0 & "," &  -1*cdbl(rsContractDB("AMOUNT")) & "," & _
								" to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy'),'"  & rsContractDB("AGREEMENTCODE")+"-Cr" & "','" & _
								Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") &  "', sysdate, to_char(sysdate, 'HH24:MI'), " & maxLineNoContractorCB & " )"
					  response.Flush()			  
					  connStr.execute(qInsertCR)  

                      connStr.execute("Commit")					  
				  
				  rsContractDB.MoveNext
			   loop '--------- Loop For rsContractDB
			end if
			rsContractDB.close
			set rsContractDB = Nothing
			
		    
			' JIka FCBA <> CHARGEUNIT [CHARGEUNIT=FCBA]
			Set rsContractDB = Server.CreateObject("ADODB.Recordset")
			rsContractDB.CursorType = 3
			rsContractDB.CursorLocation = 3
			qContractDB = " SELECT "& _
			             " TO_CHAR(LVL1.WIPDATE,'fmMM') ACCOUNTINGPERIOD, "& _
						 " TO_CHAR(LVL1.WIPDATE,'yyyy') ACCOUNTINGYEAR, "& _
						 " LVL1.AGREEMENTCODE AGREEMENTCODE, LVL1.WIPDATE WIPDATE, CA.CONTRACTORCODE CONTRACTORCODE, "& _
						 " businessunit_api.get_jobintraco( CA.FCBA) controljob,"& _
						 " PAR.param_7 controljobtemp,  "& _
						 " CA.DESCRIPTION CONTRACTORDESC, LVL1.JOBCODE JOB, LVL1.LOCATIONTYPE LOCATIONTYPE, LVL1.LOCATIONCODE LOCATIONCODE, "& _
						 " SUM(LVL1.QTY * CAD.RATE) AMOUNT, "& _
						 " SUM(LVL1.QTY) QTY, "& _
						 " job_api.get_jobcode_name(LVL1.JOBCODE , CA.FCBA) JOBNAME,"& _
						 " contractor_api.get_contractorname(CA.CONTRACTORCODE, CA.FCBA) CONTRACTORNAME,LVL1.FCBA,LVL1.CHARGEUNIT"& _
						 " FROM( "& _
						 " Select "& _
						 " HDR.FCBA, DTL.CHARGEUNIT, HDR.WIPDATE, HDR.AGREEMENTCODE,  DTL.JOBCODE, HDR.AGREEMENTLINENO,"& _
						 " DTL.LOCATIONTYPE, DTL.LOCATIONCODE, SUM(HDR.QTY) QTY"& _
						 " FROM "& _
						 " WORKINPROGRESSPLAN DTL "& _
						 " INNER JOIN WORKINPROGRESS HDR ON DTL.AGREEMENTCODE = HDR.AGREEMENTCODE AND "& _
						 " DTL.FCBA = HDR.FCBA AND HDR.AGREEMENTFINALNO = DTL.AGREEMENTFINALNO AND HDR.AGREEMENTLINENO = DTL.AGREEMENTLINENO "& _
						 " AND DTL.FCBA <> DTL.CHARGEUNIT  "& _
						 " GROUP BY "& _
						 " HDR.FCBA,  DTL.CHARGEUNIT, HDR.WIPDATE, HDR.AGREEMENTCODE,DTL.JOBCODE, HDR.AGREEMENTLINENO,DTL.LOCATIONTYPE, DTL.LOCATIONCODE "& _
						 " )LVL1 "& _
						 " INNER JOIN CONTRACTAGREEMENTDETAIL CAD ON LVL1.AGREEMENTCODE = CAD.AGREEMENTCODE AND LVL1.FCBA = CAD.FCBA AND  LVL1.AGREEMENTLINENO = CAD.LINENO "& _
						 " INNER JOIN CONTRACTAGREEMENT CA ON CAD.AGREEMENTCODE = CA.AGREEMENTCODE AND CAD.FCBA = CA.FCBA "& _
						 " LEFT JOIN PARAMETER PAR ON CAD.FCBA = PAR.FCBA AND parcode = 'JOBCODE_CONTRACTOR_INVOICE'  "& _
						 " Where "& _
						 " LVL1.CHARGEUNIT='" & Session("fcBA") & "'" & _
						 " and to_number(to_Char(LVL1.WIPDATE,'mm'))=" & Month(Session("v_Contract")) &  _
						 " and to_number(to_Char(LVL1.WIPDATE,'yyyy'))=" & Year(Session("v_Contract")) &  _
						 " GROUP BY CA.FCBA, LVL1.AGREEMENTCODE, LVL1.WIPDATE, CA.CONTRACTORCODE, Contractor_api.get_controljob(CA.CONTRACTORCODE, CA.FCBA), "& _
						 " CA.DESCRIPTION,  LVL1.JOBCODE, LVL1.LOCATIONTYPE, LVL1.LOCATIONCODE,PAR.param_7,LVL1.FCBA,LVL1.CHARGEUNIT"& _
						 " ORDER BY CHARGEUNIT,agreementcode,job,locationcode "
		
			rsContractDB.Open qContractDB, connstr
			if not rsContractDB.eof then
			   Do While Not rsContractDB.eof
				  Response.Write "FCBA <> CHARGEUNIT [CHARGEUNIT=FCBA] to Allocate the  " & rsContractDB("AGREEMENTCODE") & "<br>"
				  Response.Flush()		
				  
					Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
					rsCostActivityMax.CursorType = 3
					rsCostActivityMax.CursorLocation = 3
					qCostActivityMax =  " select nvl(max(lineno),0) maxLineNo from cost_activity " & _
										" where fcba = '" &  rsContractDB("CHARGEUNIT") & "' " 
					rsCostActivityMax.Open qCostActivityMax, connstr
					maxLineNoINT = cdbl(rsCostActivityMax("maxLineNo"))
					rsCostActivityMax.close
					set rsCostActivityMax = Nothing


				  '-- UnitWIP 
					 maxLineNoINT = maxLineNoINT + 1
					 qCost_ActivityTRDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix,CHARGEUNIT ) " & _
									 " values (" & maxLineNoINT & ",to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy') " & ",'" & rsContractDB("job") & "','WC','WC','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '','" & rsContractDB("LOCATIONTYPE") & "','" & rsContractDB("LOCATIONCODE") & "'," & _
									 " '" &  cdbl(rsContractDB("AMOUNT")) & "','" & rsContractDB("AGREEMENTCODE") &"-"&rsContractDB("CHARGEUNIT") & "'," & rsContractDB("QTY") &",0,'" & rsContractDB("CHARGEUNIT") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '',''," & _
									 " '',''," & _
									 " 'INTRACO FROM ','" & rsContractDB("FCBA")  & "'," & _
									 " '','No.Kontrak/Kontraktor: " & rsContractDB("AGREEMENTCODE") & "/" & rsContractDB("CONTRACTORNAME") &" Pekerjaan: " & rsContractDB("JOBNAME") & "/" & rsContractDB("LOCATIONCODE") & "', " & _
									 " 'JOB-BLOK', '"& rsContractDB("JOB") &"-" & rsContractDB("LOCATIONCODE") &"','Db', " & _
									 " " & rsContractDB("ACCOUNTINGPERIOD") & "," & rsContractDB("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIP','" & rsContractDB("FCBA") & "') "
					connstr.execute(qCost_ActivityTRDB1)
					 
					maxLineNoINT = maxLineNoINT + 1
					qCost_ActivityTRCR1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix,CHARGEUNIT ) " & _
									 " values (" & maxLineNoINT & ",to_date('" & rsContractDB("WIPDATE") & "','mm/dd/yyyy') " & ",'" & rsContractDB("controljob") & "','WC','WC','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '','CA','" & rsContractDB("CONTRACTORCODE") & "'," & _
									 " '" &  cdbl(rsContractDB("AMOUNT")) * -1 & "','" & rsContractDB("AGREEMENTCODE") &"-"&rsContractDB("CHARGEUNIT") & "'," & rsContractDB("QTY") &",0,'" & rsContractDB("CHARGEUNIT") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '',''," & _
									 " '',''," & _
									 " 'INTRACO FROM ','" & rsContractDB("FCBA")  & "'," & _
									 " '','No.Kontrak/Kontraktor: " & rsContractDB("AGREEMENTCODE") & "/" & rsContractDB("CONTRACTORNAME") &" Pekerjaan: " & rsContractDB("JOBNAME") & "/" & rsContractDB("LOCATIONCODE") & "', " & _
									 " 'JOB-BLOK', '"& rsContractDB("JOB") &"-" & rsContractDB("LOCATIONCODE") &"','Cr', " & _
									 " " & rsContractDB("ACCOUNTINGPERIOD") & "," & rsContractDB("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIP','" & rsContractDB("FCBA") & "') "
					  connstr.execute(qCost_ActivityTRCR1)
									  
				 	  connStr.execute("Commit")					  
				  
				  rsContractDB.MoveNext
			   loop '--------- Loop For rsContractDB
			end if
			rsContractDB.close
			set rsContractDB = Nothing
			
			
			qjournal=	" MERGE INTO cost_activity x " & _
						" USING " & _
						" ( " & _
						"  SELECT fcba,accountingyear,accountingperiod,referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
						"  ( " & _
						"  SELECT DISTINCT fcba,accountingyear,accountingperiod,referenceno,journalprefix,company_api.get_companyno(businessunit_api.get_companycode(fcba))||'/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
						"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Contract")) & " AND accountingyear=" & Year(Session("v_Contract")) & " AND journalprefix='WIP' " & _
						"  Order BY referenceno " & _
						"  ) " & _
						" )y ON (x.fcba=y.fcba and  x.referenceno =  y.referenceno AND  x.accountingyear=y.accountingyear AND x.accountingperiod=y.accountingperiod AND x.journalprefix=y.journalprefix) " & _
						" WHEN MATCHED THEN " & _
						" UPDATE SET x.journalno = y.seq " 
			connstr.execute(qjournal)
			
			response.Write "Process WIP Contract Complete.... <br>"
		    response.Flush
		
			
        ' -- INVOICE		
			'-- AMH 05-22 Modify to 4 contract HO 
			'-- Jika FCBA = INVOICE BY
			'-  Fcba = Invoice By
			'-- Journal Db : Clearing Account 
			'-- Journal Db : PPN MAsukan sudah di terima
			'-- Journal Cr : Hutang PPH 23
			'-- Journal Cr : HUtang Contractor
			
			
			'-- Jika FCBA <> INVOICE BY
			'-  Fcba = Invoice By
			'-- Journal Db : Clearing Account 
			'-- Journal Db : PPN MAsukan sudah di terima
			'-- Journal Cr : Hutang PPH 23
			'-- Journal Cr : Intraco
			
			'-- fcba = fcba
			'-- Journal Db : Hutang COntractor
			'-- Journal Cr : Intraco 	
			
			
			

			response.Write "Process Invoice Contract begin.... <br>"
		    response.Flush
			
			Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
			rsCostActivityMax.CursorType = 3
			rsCostActivityMax.CursorLocation = 3
											
			qCostActivityMax =  " select max(lineno) maxLineNo from cost_activity " & _
								" where fcba = '" &  Session("fcBA") & "' " 
			rsCostActivityMax.Open qCostActivityMax, connstr

			maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))

			rsCostActivityMax.close
			set rsCostActivityMax = Nothing
		
		'-- FCBA = INVOCIE BY 
			qContractInvCR2 =" SELECT " &_
							" FCBA, ROWNUM LINENO,  " &_
							" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_		 
							" TDATE, JOBCODE_DB1,JOBCODE_DB2,JOBCODE_CR1,JOBCODE_CR2, SOURCE, SOURCETYPE ,SOURCECODE, TARGETTYPE, TARGETCODE, " &_ 
							" DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT_DB1,AMOUNT_DB2,AMOUNT_CR1,AMOUNT_CR2, " &_ 
							" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " &_ 
							" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " &_ 
							" TITLE_REFERENCE_03, AUTO_REFERENCE_03  " &_
							" FROM  " &_
							" (  " &_
							" SELECT  " &_ 
							" MAX(HDR.INVOICE_BY)FCBA, " &_ 
							" MAX(TO_CHAR(HDR.TRANSDATE,'fmMM')) ACCOUNTINGPERIOD,  " &_
							" MAX(TO_CHAR(HDR.TRANSDATE,'yyyy')) ACCOUNTINGYEAR, " &_ 
							" MAX(PERIODCONTROL.CURRENTDATE) TDATE, " &_ 
							" MAX(PARAMETER.PARAM_7) JOBCODE_DB1, " &_ 
							" MAX(PARAMETER.PARAM_3) JOBCODE_DB2, " &_ 
							" MAX(PARAMETER.PARAM_8) JOBCODE_CR1, " &_ 
							" MAX(CONTRACTOR.CONTROLJOB) JOBCODE_CR2 , " &_ 
							" 'WC' SOURCE,  " &_
							" 'WC' SOURCETYPE,  " &_
							" MAX(HDR.INVOICECODE) SOURCECODE,  " &_
							" MAX(JOB.JOB_CATEGORY) TARGETTYPE, " &_ 
							" 'CA' TARGETTYPE_CR2, " &_ 
							" MAX(HDR.CONTRACTORCODE) TARGETCODE,  " &_
							" MAX(HDR.AGREEMENTCODE) DOCUMENTNO, " &_ 
							" 0 QUANTITY,  " &_
							" 0 OUTPUT, " &_ 
							" MAX(NVL(HDR.TOTALAMOUNT_AFTERRETENTION,0)) AMOUNT_DB1," &_ 
							" MAX(NVL(HDR.AMOUNTPPN,0)) AMOUNT_DB2," &_ 
							" -1*MAX(NVL(HDR.AMOUNTPPH,0))AMOUNT_CR1," &_ 
							" -1*MAX(NVL(HDR.grandtotal_aftertax,0)) AMOUNT_CR2," &_ 
							" 'TC - Contractor Invoice Code :' TITLE_REFERENCE_01,  " &_
							" MAX(HDR.INVOICECODE) AUTO_REFERENCE_01, " &_ 
							" 'Auto - Aggrement Code :' TITLE_REFERENCE_02, " &_ 
							" MAX(HDR.AGREEMENTCODE) AUTO_REFERENCE_02, " &_ 
							" 'Contractor Code ' TITLE_REFERENCE_03,  " &_
							" MAX(HDR.CONTRACTORCODE)||'-'||Max(contractor_api.get_contractorname(HDR.CONTRACTORCODE,HDR.FCBA)) AUTO_REFERENCE_03 " &_
							" FROM  " &_
							" CONTRACTORINVOICEHDR HDR " &_
							" LEFT JOIN CONTRACTOR ON HDR.FCBA = CONTRACTOR.FCBA AND HDR.CONTRACTORCODE = CONTRACTOR.FCCODE " &_ 
							" INNER JOIN PERIODCONTROL ON HDR.INVOICE_BY = PERIODCONTROL.FCBA AND UPPER(PERIODCONTROL.SYSTEM)='CONTRACT'  " &_ 
							" INNER JOIN PARAMETER ON HDR.INVOICE_BY = PARAMETER.FCBA AND PARAMETER.PARCODE = 'JOBCODE_CONTRACTOR_INVOICE' " &_  
							" INNER JOIN JOB ON HDR.INVOICE_BY = JOB.FCBA AND PARAMETER.PARAM_8 = JOB.FCCODE " &_  
							" INNER JOIN GCMASTER ON HDR.INVOICE_BY = GCMASTER.FCBA AND GCMASTER.ITEM = JOB.JOB_CATEGORY   " &_
							" WHERE " &_  
							" TO_CHAR(HDR.TRANSDATE,'fmMM') = " & Month(Session("v_Contract")) &  _
							" AND TO_CHAR(HDR.TRANSDATE,'yyyy') =  " & Year(Session("v_Contract")) &  _
							" AND HDR.INVOICE_BY = HDR.FCBA " & _
							" AND HDR.INVOICE_BY = '" & Session("fcBA") & "'" & _
							" GROUP BY HDR.INVOICE_BY,HDR.FCBA,HDR.INVOICECODE)"
							
		    'response.write qContractInvCR2 & "<br>" 
			
			Set rsContractInvCR2 = Server.CreateObject("ADODB.Recordset")
			rsContractInvCR2.Open qContractInvCR2, connstr
			if not rsContractInvCR2.eof then
				Do While Not rsContractInvCR2.eof
					 '--Insert Debet 1
					 maxLineNo = maxLineNo + 1
					 
					 qCost_ActivityTRDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_DB1") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 " '" &  Round(rsContractInvCR2("AMOUNT_DB1")) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Db', " & _
									 " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIV') "
			         
					 connstr.execute(qCost_ActivityTRDB1)
					 '--Insert Debet 2
					  maxLineNo = maxLineNo + 1
					 qCost_ActivityTRDB2="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ,journalprefix) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_DB2") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 " '" &  round(cdbl(to_nol(rsContractInvCR2("AMOUNT_DB2")))) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Db', " & _
									 " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIV') "
                     
					 connstr.execute(qCost_ActivityTRDB2)
					 '--Insert Credit 1
					 maxLineNo = maxLineNo + 1
					 qCost_ActivityTRCR1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ,journalprefix) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_CR1") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 " '" &  round(cdbl(to_nol(rsContractInvCR2("AMOUNT_CR1")))) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Cr', " & _
									 " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIV') "
                     
					 connstr.execute(qCost_ActivityTRCR1)
					 '--Insert Credit 2
					 maxLineNo = maxLineNo + 1
					 qCost_ActivityTRCR2="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_CR2") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 " '" &  round(cdbl(to_nol(rsContractInvCR2("AMOUNT_CR2")))) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Cr', " & _
									 " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIV') "
									 
					 connstr.execute(qCost_ActivityTRCR2)
					 connStr.execute("Commit")
					 
					rsContractInvCR2.MoveNext
				loop
		end if								
		'-- End Insert Credit 2
		
		
		
		
		'-- FCBA <> INVOCIE BY 
			qContractInvCR2 =" SELECT " &_
							" FCBA, CHARGEUNIT, ROWNUM LINENO,  " &_
							" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_		 
							" TDATE, JOBCODE_INT, JOBCODE_DB1,JOBCODE_DB2,JOBCODE_CR1,JOBCODE_CR2, SOURCE, SOURCETYPE ,SOURCECODE, TARGETTYPE, TARGETCODE, " &_ 
							" DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT_DB1,AMOUNT_DB2,AMOUNT_CR1,AMOUNT_CR2, " &_ 
							" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " &_ 
							" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " &_ 
							" TITLE_REFERENCE_03, AUTO_REFERENCE_03  " &_
							" FROM  " &_
							" (  " &_
							" SELECT  " &_ 
							" MAX(HDR.INVOICE_BY)FCBA, " &_ 
							" MAX(HDR.FCBA)CHARGEUNIT, " &_ 
							" MAX(TO_CHAR(HDR.TRANSDATE,'fmMM')) ACCOUNTINGPERIOD,  " &_
							" MAX(TO_CHAR(HDR.TRANSDATE,'yyyy')) ACCOUNTINGYEAR, " &_ 
							" MAX(PERIODCONTROL.CURRENTDATE) TDATE, " &_ 
							" businessunit_api.get_jobintraco( HDR.INVOICE_BY) JOBCODE_INT, " &_ 
							" MAX(PARAMETER.PARAM_7) JOBCODE_DB1, " &_ 
							" MAX(PARAMETER.PARAM_3) JOBCODE_DB2, " &_ 
							" MAX(PARAMETER.PARAM_8) JOBCODE_CR1, " &_ 
							" MAX(CONTRACTOR.CONTROLJOB) JOBCODE_CR2 , " &_ 
							" 'WC' SOURCE,  " &_
							" 'WC' SOURCETYPE,  " &_
							" MAX(HDR.INVOICECODE) SOURCECODE,  " &_
							" MAX(JOB.JOB_CATEGORY) TARGETTYPE, " &_ 
							" 'CA' TARGETTYPE_CR2, " &_ 
							" MAX(HDR.CONTRACTORCODE) TARGETCODE,  " &_
							" MAX(HDR.AGREEMENTCODE) DOCUMENTNO, " &_ 
							" 0 QUANTITY,  " &_
							" 0 OUTPUT, " &_ 
							" MAX(NVL(HDR.TOTALAMOUNT_AFTERRETENTION,0)) AMOUNT_DB1," &_ 
							" MAX(NVL(HDR.AMOUNTPPN,0)) AMOUNT_DB2," &_ 
							" -1*MAX(NVL(HDR.AMOUNTPPH,0))AMOUNT_CR1," &_ 
							" -1*MAX(NVL(HDR.grandtotal_aftertax,0)) AMOUNT_CR2," &_ 
							" 'TC - Contractor Invoice Code :' TITLE_REFERENCE_01,  " &_
							" MAX(HDR.INVOICECODE) AUTO_REFERENCE_01, " &_ 
							" 'Auto - Aggrement Code :' TITLE_REFERENCE_02, " &_ 
							" MAX(HDR.AGREEMENTCODE) AUTO_REFERENCE_02, " &_ 
							" 'Contractor Code ' TITLE_REFERENCE_03,  " &_
							" MAX(HDR.CONTRACTORCODE)||'-'||Max(contractor_api.get_contractorname(HDR.CONTRACTORCODE,HDR.FCBA)) AUTO_REFERENCE_03 " &_
							" FROM  " &_
							" CONTRACTORINVOICEHDR HDR " &_
							" LEFT JOIN CONTRACTOR ON HDR.FCBA = CONTRACTOR.FCBA AND HDR.CONTRACTORCODE = CONTRACTOR.FCCODE " &_ 
							" INNER JOIN PERIODCONTROL ON HDR.INVOICE_BY = PERIODCONTROL.FCBA AND UPPER(PERIODCONTROL.SYSTEM)='CONTRACT'  " &_ 
							" INNER JOIN PARAMETER ON HDR.INVOICE_BY = PARAMETER.FCBA AND PARAMETER.PARCODE = 'JOBCODE_CONTRACTOR_INVOICE' " &_  
							" INNER JOIN JOB ON HDR.INVOICE_BY = JOB.FCBA AND PARAMETER.PARAM_8 = JOB.FCCODE " &_  
							" INNER JOIN GCMASTER ON HDR.INVOICE_BY = GCMASTER.FCBA AND GCMASTER.ITEM = JOB.JOB_CATEGORY   " &_
							" WHERE " &_  
							" TO_CHAR(HDR.TRANSDATE,'fmMM') = " & Month(Session("v_Contract")) &  _
							" AND TO_CHAR(HDR.TRANSDATE,'yyyy') =  " & Year(Session("v_Contract")) &  _
							" AND HDR.INVOICE_BY <> HDR.FCBA " & _
							" AND HDR.INVOICE_BY = '" & Session("fcBA") & "'" & _
							" GROUP BY HDR.INVOICE_BY,HDR.FCBA,HDR.INVOICECODE)"
							
		    'response.write qContractInvCR2 & "<br>" 
			
			Set rsContractInvCR2 = Server.CreateObject("ADODB.Recordset")
			rsContractInvCR2.Open qContractInvCR2, connstr
			if not rsContractInvCR2.eof then
				Do While Not rsContractInvCR2.eof
					 '--Insert Debet 1
					 maxLineNo = maxLineNo + 1
					 
					 qCost_ActivityTRDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix,CHARGEUNIT ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_DB1") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 " '" &  Round(rsContractInvCR2("AMOUNT_DB1")) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Db', " & _
									 " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIV','" & rsContractInvCR2("CHARGEUNIT") & "') "
			         
					 connstr.execute(qCost_ActivityTRDB1)
					 '--Insert Debet 2
					  maxLineNo = maxLineNo + 1
					 qCost_ActivityTRDB2="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ,journalprefix,CHARGEUNIT) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_DB2") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 " '" &  round(cdbl(to_nol(rsContractInvCR2("AMOUNT_DB2")))) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Db', " & _
									 " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIV','" & rsContractInvCR2("CHARGEUNIT") & "') "
                     
					 connstr.execute(qCost_ActivityTRDB2)
					 '--Insert Credit 1
					 maxLineNo = maxLineNo + 1
					 qCost_ActivityTRCR1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ,journalprefix,CHARGEUNIT) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_CR1") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 " '" &  round(cdbl(to_nol(rsContractInvCR2("AMOUNT_CR1")))) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Cr', " & _
									 " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIV','" & rsContractInvCR2("CHARGEUNIT") & "') "
                     
					 connstr.execute(qCost_ActivityTRCR1)
					 '--Insert Credit 2
					 maxLineNo = maxLineNo + 1
					 qCost_ActivityTRCR2="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME,journalprefix ,CHARGEUNIT) " & _
									 " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_INT") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 " '" &  round(cdbl(to_nol(rsContractInvCR2("AMOUNT_CR2")))) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Cr', " & _
									 " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI'),'WIV','" & rsContractInvCR2("CHARGEUNIT") & "') "
									 
					 connstr.execute(qCost_ActivityTRCR2)
					 
					 
					 ' '--- INTRACO
					  ' maxLineNo = maxLineNo + 1
					 
					 ' qCost_ActivityTRDB1="Insert into Cost_Activity " & _
									 ' "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 ' "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 ' " LASTUPDATE, LASTTIME,journalprefix,CHARGEUNIT ) " & _
									 ' " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_DB1") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 ' " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 ' " '" &  Round(rsContractInvCR2("AMOUNT_DB1")) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("CHARGEUNIT") & "',"  & _
									 ' " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 ' " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 ' " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 ' " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 ' " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 ' " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Db', " & _
									 ' " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 ' " sysdate, to_char(sysdate,'HH24:MI'),'WIV','" & rsContractInvCR2("FCBA") & "') "
			         
					 ' connstr.execute(qCost_ActivityTRDB1)
					 
					 ' maxLineNo = maxLineNo + 1
					 ' qCost_ActivityTRCR2="Insert into Cost_Activity " & _
									 ' "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 ' "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 ' " LASTUPDATE, LASTTIME,journalprefix,CHARGEUNIT ) " & _
									 ' " values (" & maxLineNo & ",to_date('" & rsContractInvCR2("TDATE") & "','mm/dd/yyyy') " & ",'" & rsContractInvCR2("JOBCODE_INT") & "','" & rsContractInvCR2("SOURCE") & "','" & rsContractInvCR2("SOURCETYPE") & "','" & rsContractInvCR2("SOURCECODE") & "'," & _
									 ' " '','" & rsContractInvCR2("TARGETTYPE") & "','" & rsContractInvCR2("TARGETCODE") & "'," & _
									 ' " '" &  round(cdbl(to_nol(rsContractInvCR2("AMOUNT_CR2")))) & "','" & rsContractInvCR2("DOCUMENTNO") & "'," & "0,0,'" & rsContractInvCR2("CHARGEUNIT") & "',"  & _
									 ' " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 ' " '" & rsContractInvCR2("TITLE_REFERENCE_01") & "','"  & rsContractInvCR2("AUTO_REFERENCE_01") & "'," & _
									 ' " '" & rsContractInvCR2("TITLE_REFERENCE_02") & "','"  & rsContractInvCR2("AUTO_REFERENCE_02") & "'," & _
									 ' " '" & rsContractInvCR2("TITLE_REFERENCE_03") & "','"  & rsContractInvCR2("AUTO_REFERENCE_03") & "'," & _
									 ' " '','INVOICE WIP: " & rsContractInvCR2("AUTO_REFERENCE_01") & " KONTRAKTOR:" & rsContractInvCR2("AUTO_REFERENCE_03") & "', " & _
									 ' " 'Line No:', '" & rsContractInvCR2("LINENO") & "','Cr', " & _
									 ' " " & rsContractInvCR2("ACCOUNTINGPERIOD") & "," & rsContractInvCR2("ACCOUNTINGYEAR") & "," & _
									 ' " sysdate, to_char(sysdate,'HH24:MI'),'WIV','" & rsContractInvCR2("FCBA") & "') "
									 
					 ' connstr.execute(qCost_ActivityTRCR2)
					 
					 ' connStr.execute("Commit")
					 
					 
					 
					 
					rsContractInvCR2.MoveNext
				loop
		end if				
	
		'end if
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		%>
	<%' end if ' End Remarked By UKIE 21 feb 2014 purpose: Controljob from Contractor Master Not from Module
	
	'----- Tranfer to Activity Costing
	response.Write "Process Invoice Contract Complete.... !!<br>"
	response.Flush
	
	' Get the Max LineNo of Cost Activity
	' Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
	' rsCostActivityMax.CursorType = 3
	' rsCostActivityMax.CursorLocation = 3
									
	' qCostActivityMax =  " select nvl(max(lineno),0) maxLineNo from cost_activity " & _
						' " where fcba = '" &  Session("fcBA") & "' " 
	' rsCostActivityMax.Open qCostActivityMax, connstr

	' maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))

	' rsCostActivityMax.close
	' set rsCostActivityMax = Nothing
	' ' End of Get the Max LineNo of Cost Activity
	' ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

	' qVoucher = " SELECT TRANSDATE tdate, " & _
					   ' " JOBCODE jobcode," & _ 
					   ' " 'WC' source, " & _
					   ' " 'WC' sourcetype, " & _
					   ' " CONTRACTORCODE sourcecode, " & _
					   ' " '' sourcecode_status, " & _
					   ' " TARGETTYPE, " & _
					   ' " TARGETCODE , " & _
					   ' " '' targetcode_status, " & _
					   ' " REFERENCENO, " & _
					   ' " SUM(QUANTITY) quantity, " & _
					   ' " 0 output, " & _
					   ' " SUM(AMOUNT) amount, " & _
					   ' " FCBA," & _
					   ' " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
			  ' " FROM ContractorCB" & _
			  ' " WHERE to_number(to_Char(TRANSDATE,'mm'))=" & Month(Session("v_Contract")) &  _
			  ' " and to_number(to_Char(TRANSDATE,'yyyy'))=" & Year(Session("v_Contract")) & _ 
			  ' " and fcBA='" & Session("fcBA") & "'" & _
			  ' " group by " & _
			  ' " TRANSDATE, JOBCODE, 'WC', 'WC', CONTRACTORCODE, '', TARGETTYPE, TARGETCODE, '', REFERENCENO,  FCBA"

			  
	' Set rsVoucher = Server.CreateObject("ADODB.Recordset")
	' rsVoucher.CursorType = 3
	' rsVoucher.CursorLocation = 3
	' rsVoucher.Open qVoucher, connstr



	' if not rsVoucher.eof  then
		' Do While Not rsVoucher.eof
	   
			' maxLineNo = maxLineNo + 1

			' if cdbl(rsVoucher("amount")) >= 0 then 
			    ' temp_DebetCredit = "Dr"
			' else
			    ' temp_DebetCredit = "Cr"
		    ' end if	
			
			' qCost_Activity="Insert into Cost_Activity " & _
								' " ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								' "   TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
								' "   QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, " & _
								' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
								' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								' " LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
								' " values ( to_date('" & rsVoucher("tdate") & "','mm/dd/yyyy') " & ",'" & rsVoucher("jobcode")  & _ 
								' "','" & rsVoucher("source") & "','" & rsVoucher("sourcetype") & "','" & rsVoucher("sourcecode") & "'" & _
								' ",'" & rsVoucher("sourcecode_status") & "','" & rsVoucher("targettype") & "','" & rsVoucher("targetcode") & "'" & _
								' ",'" & rsVoucher("targetcode_status") & "','" & rsVoucher("referenceno") & "'," & _
								' cdbl(rsVoucher("quantity")) & "," & _
								' cdbl(rsVoucher("output")) & "," & _
								' cdbl(rsVoucher("amount")) & ",'" & _
								' Session("fcBA") & "'," & _
								' " to_date('" & rsVoucher("CLOSING_DATE") & "','mm/dd/yyyy') , sysdate, " & _
								' " Cost_Activity_Info_Api.Get_Title_Reference_01 ('" & rsVoucher("SOURCETYPE") & "','" & rsVoucher("TARGETTYPE") & "','" & temp_DebetCredit & "'), '" & rsVoucher("TARGETCODE") & "', " & _
								' " '','', " & _
								' " '','', " & _
								' " '','', " & _
								' " '','','" & temp_DebetCredit & "', " & _
								' Month(Session("v_Contract")) & ", " & _
								' Year(Session("v_Contract")) & ", " & _
								' " sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & ",'WIP' ) "
			' 'response.write 	qCost_Activity				
			' connstr.execute(qCost_Activity)		

			' ' Insert History Transaction,
			
			
			

			' qVoucherDetail = " SELECT LINENO, " & _
							' " QUANTITY, " & _
							' " AMOUNT " & _
				  ' " FROM CONTRACTORCB" & _
				  ' " WHERE to_number(to_Char(TRANSDATE,'mm'))=" & Month(Session("v_Contract")) &  _
				  ' " and to_number(to_Char(TRANSDATE,'yyyy'))=" & Year(Session("v_Contract")) & _ 
				  ' " and fcBA='" & Session("fcBA") & "'" & _
				  ' " and transdate = to_date('" & rsVoucher("tdate") & "','mm/dd/yyyy') " & _
				  ' "	and jobcode = '" & rsVoucher("jobcode") & "' " & _
				  ' "	and contractorcode ='" & rsVoucher("sourcecode") & "' " & _
				  ' "	and targettype ='" & rsVoucher("targettype") & "' " & _
				  ' "	and targetcode = '" & rsVoucher("targetcode") & "' " & _
				  ' "	and referenceno = '" & rsVoucher("referenceno") & "' " 
						
							
			
			' Set rsVoucherDetail = Server.CreateObject("ADODB.Recordset")
			' rsVoucherDetail.CursorType = 3
			' rsVoucherDetail.CursorLocation = 3
			' rsVoucherDetail.Open qVoucherDetail, connstr
			

			
			' if not rsVoucherDetail.eof  then	
				' maxLineNoDetail = 0
				' Do While Not rsVoucherDetail.eof


					' maxLineNoDetail = maxLineNoDetail + 1
					
					
					' txtRefStringValues_PVDetail = rsVoucherDetail("LINENO")  & txtRefStringSeparator_PVDetail & _
												  ' Session("fcBA") 	

					' nQuantity = cdbl(rsVoucherDetail("quantity"))
					' nOutput = 0
					' nAmount = cdbl(rsVoucherDetail("amount"))

					' qCost_ActivityHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
									' "( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
									' "  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
									' " QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
									' maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodePVDetail & "','" & txtRefNumberKeys_PVDetail & "','" & _
									' txtRefStringSeparator_PVDetail & "','" & txtRefStringKeys_PVDetail & "','" &  txtRefStringValues_PVDetail & "' , " & _
									' nQuantity & "," & nOutput & "," & nAmount &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
									' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
								 
					' connstr.execute(qCost_ActivityHistory)				
					
					' rsVoucherDetail.MoveNext
				' loop
			' end if
			' rsVoucherDetail.Close
			' set rsVoucherDetail = Nothing	

			' rsVoucher.MoveNext
		' loop
	' End if		
	
	qjournal=		" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT fcba,accountingyear,accountingperiod,referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT fcba,accountingyear,accountingperiod,sourcecode referenceno,journalprefix,company_api.get_companyno(businessunit_api.get_companycode(fcba))||'/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Contract")) & " AND accountingyear=" & Year(Session("v_Contract")) & " AND journalprefix='WIV' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.fcba=y.fcba and  x.sourcecode =  y.referenceno AND  x.accountingyear=y.accountingyear AND x.accountingperiod=y.accountingperiod AND x.journalprefix=y.journalprefix) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
	connstr.execute(qjournal)
		
	' Normalisasi Contractor Ledger
	NormalisasiContractorLedger =	" Delete From CONTRACTORLEDGER WHERE" & _
								" MONTH =" & Month(Session("v_Contract")) & " AND" & _
								" YEAR =" & Year(Session("v_Contract")) & " AND" & _ 
								" FCBA='" & Session("fcBA") & "'"
	connstr.execute(NormalisasiContractorLedger)

	vchar = Year(Session("v_Contract")) 
	vchar = vchar + Month(Session("v_Contract"))
		
	'Update To Invoice Ledger
	SelectContractorLedger =  " select fcba, agreementcode, contractorcode, agreementlineno, year, month,progresspaymentdate, " & _
						   " sum(totalAmount) totalAmount, " & _
						   " sum(totalBeginningBalancePaid) totalBeginningBalancePaid, " & _
						   " sum(monthlyPaid) monthlyPaid, " & _
						   " sum(outstandingPaid) outstandingPaid, " & _
						   " sum(position_0) position_0, " & _
						   " sum(position_1) position_1, " & _
						   " sum(position_2) position_2, " & _
						   " sum(position_3) position_3 " & _
						   " FROM ( " & _
							" SELECT" & _
							" fcba," & _
							" agreementcode," & _
							" contractorcode," & _
							" agreementlineno," & _
							" year," & _
							" month," & _
							" to_char(progresspaymentdate, 'mm/dd/yyyy') progresspaymentdate," & _
							" totalAmount," & _
							" totalBeginningBalancePaid," & _
							" monthlyPaid," & _
							" (totalAmount - (totalBeginningBalancePaid + monthlyPaid)) outstandingPaid," & _
							" decode (GetAgingPosition (progresspaymentdate), 0, (totalAmount - (totalBeginningBalancePaid + monthlyPaid)),0) position_0," & _
							" decode (GetAgingPosition (progresspaymentdate), 1, (totalAmount - (totalBeginningBalancePaid + monthlyPaid)),0) position_1," & _
							" decode (GetAgingPosition (progresspaymentdate), 2, (totalAmount - (totalBeginningBalancePaid + monthlyPaid)),0) position_2," & _
							" decode (GetAgingPosition (progresspaymentdate), 3, (totalAmount - (totalBeginningBalancePaid + monthlyPaid)),0) position_3" & _
							" from (" & _
							" SELECT a.fcba," & _
							" a.agreementcode," & _
							" b.contractorcode," & _
							" a.agreementlineno," & _
							" TO_CHAR (a.progresspaymentdate, 'YYYY') year," & _
							" TO_CHAR (a.progresspaymentdate, 'MM') month," & _
							" a.progresspaymentdate," & _
							" b.totalAmount," & _
							" NVL ( (SELECT totalBeginningBalancePaid" & _
							" FROM ( (  SELECT fcba," & _
							" agreementcode," & _
							" agreementlineno," & _
							" SUM (NVL (paid, 0)) totalBeginningBalancePaid" & _
							" FROM workinprogresspaid" & _
							" WHERE TO_NUMBER (" & _
							" TO_CHAR (wippaiddate, 'YYYYMM')) <" & vchar & _
							" GROUP BY fcba, agreementcode,agreementlineno ))" & _
							" WHERE fcba = a.fcba AND agreementcode = a.agreementcode AND agreementlineno = a.agreementlineno)," & _
							" 0)" & _
							" totalBeginningBalancePaid," & _
							" NVL ( (SELECT monthlyPaid" & _
							" FROM ( (  SELECT fcba," & _
							" agreementcode," & _
							" agreementlineno," & _
							" SUM (NVL (paid, 0)) monthlyPaid" & _
							" FROM workinprogresspaid" & _
							" WHERE TO_NUMBER (" & _
							" TO_CHAR (wippaiddate, 'YYYYMM')) =" & vchar & _
							" GROUP BY fcba, agreementcode,agreementlineno ))" & _
							" WHERE fcba = a.fcba AND agreementcode = a.agreementcode AND agreementlineno = a.agreementlineno)," & _
							" 0)" & _
							" monthlyPaid" & _
							" FROM workinprogress a," & _
							" (  SELECT fcba, agreementcode, contractorcode, lineno,  SUM (round(qty * rate,0)) totalAmount" & _
							" FROM agreementdetail" & _
							" GROUP BY fcba, agreementcode, contractorcode, lineno) b" & _
							" WHERE a.fcba = b.fcba AND a.agreementcode = b.agreementcode AND a.agreementlineno = b.lineno)" & _
							" WHERE MONTH =" & Month(Session("v_Contract")) & " AND" & _
							" YEAR =" & Year(Session("v_Contract")) & " AND" & _ 
							" FCBA='" & Session("fcBA") & "' ) " & _
							" GROUP BY fcba, agreementcode, contractorcode, agreementlineno, year, month,progresspaymentdate "
						   
	  
	Set rsContractorLedger = Server.CreateObject("ADODB.Recordset")
	rsContractorLedger.CursorType = 3
	rsContractorLedger.CursorLocation = 3
	rsContractorLedger.Open SelectContractorLedger, connstr
	if not rsContractorLedger.eof then
	   Do While Not rsContractorLedger.eof	
		  InsertContractorLedger =	"INSERT INTO CONTRACTORLEDGER (" & _
								" AGREEMENTCODE, CONTRACTORCODE, AGREEMENTLINENO," & _
								" YEAR, MONTH, PROGRESSPAYMENTDATE, TOTALAMOUNT," & _
								" TOTALBEGINNINGBALANCEPAID, MONTHLYPAID, OUTSTANDINGPAID," & _
								" POSITION_0, POSITION_1, POSITION_2, POSITION_3," & _
								" FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
								" VALUES (" & _
								" '" & rsContractorLedger("agreementcode") & "'," & _
								" '" & rsContractorLedger("contractorcode") & "'," & _
								" " & rsContractorLedger("agreementlineno") & "," & _
								" " & rsContractorLedger("year") & "," & _
								" " & rsContractorLedger("month") & "," & _
								" to_date('" & rsContractorLedger("progresspaymentdate") & "', 'mm/dd/yyyy') ," & _
								" " & rsContractorLedger("totalAmount") & "," & _
								" " & rsContractorLedger("totalBeginningBalancePaid") & "," & _
								" " & rsContractorLedger("monthlyPaid") & "," & _
								" " & rsContractorLedger("outstandingPaid") & "," & _
								" " & rsContractorLedger("position_0") & "," & _
								" " & rsContractorLedger("position_1") & "," & _
								" " & rsContractorLedger("position_2") & "," & _
								" " & rsContractorLedger("position_3") & "," & _
								" '" & Session("CodeName") & "'," & _
								" '" & Session("CodeName") & "'," & _
								" '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
								" '" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI') " & _
								" )"
		  ' response.write(InsertContractorLedger)
		  connstr.execute(InsertContractorLedger)
		  rsContractorLedger.MoveNext
	  loop
	End if
	Response.Write "<font face=Verdana size=1 color=blue>" & "Update To Contractor Ledger is Successfully." & "</font>"
	response.Flush()
	%>
		<br/><br/>
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
		response.Clear()
		
		' Add By Akbar 18 Desemeber 2015
		
			Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear  " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='CONTRACT' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for CONTRACT " & "<br>"
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
				   " where UPPER(system) ='CONTRACT' and fcba='" & Session("fcBA") & "'"
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _ 
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where UPPER(system) ='CONTRACT' and fcba='" & Session("fcBA") & "'"
				ELSE
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
					   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where  UPPER(system) ='CONTRACT' and fcba='" & Session("fcBA") & "'"
			    END IF
			END IF
			
			IF temp_NextCloseAccountingPeriod = 13 AND temp_NextCurrentAccountingPeriod = 1  THEN
			
				q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where UPPER(system) ='CONTRACT' and and fcba='" & Session("fcBA") & "'"
			END IF
			
		connStr.execute(q)
		

		
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "<br>" & "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	
		
		'--Added by Ukie at 25-Oct-2012--	
		Endingtime = FormatDateTime(CDate(Now))

		qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Contract'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME,fcentry) VALUES (" & _
						" 'Contract', " & _
						" '" & Session("fcBA") & "'," & _
						" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" " & tempMaxLine & "" & _
						", SYSDATE, TO_CHAR(SYSDATE,'HH24:MI') " & _
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
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Contract'"
					
		Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
		rsLenghtime.CursorType = 3
		rsLenghtime.CursorLocation = 3
		rsLenghtime.Open strLenghtime, connstr

		upClosingLog = " UPDATE CLOSINGLOG SET" & _
						" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Contract'"		
		connStr.execute(upClosingLog)
		'--End Added--

		%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Contract Closing is succesfully...
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		<%
			txtLastDay  =cmdFindPeriod("Contract")
			txtNextMonth=cmdPeriodClose("Contract")			

			Session("v_Contract")=txtLastDay
			Session("cls_Contract")= txtNextMonth	
		%>
	<%End if%>

<%
else 
%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/error.gif">Contract Closing is NOT completed...
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