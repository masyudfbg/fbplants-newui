<%Option Explicit%>
<%response.buffer=True%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()

		
txtLastDay  = cmdFindPeriod("Costbook")
txtNextMonth = cmdPeriodClose("Costbook")	

Session("v_Nursery") = txtLastDay
Session("cls_Nursery") = txtNextMonth

Response.Write "<br/>" & Session("v_Nursery")
Response.Write "<br/>" & Session("cls_Nursery")


Dim rsVehicle, qVehicle, nxActivity, nxGaji, nCostTotal
Dim nTotalAlokasi, rsCostVehicle, qCostVehicle, nBasis, rsBasis, qBasis, qInsertCR
Dim rsVHActivity, qVHActivity, nWSAlokasi, qInsertDB, LT, LC, FDDATE, FCJOB, nSelisih
Dim qABC, q, qLineMax, rsqLineMax, tempMaxLine
Dim strClosingLog, strLenghtime, rsLenghtime, upClosingLog, txtLastDay, txtNextMonth
Dim maxLineNoVehicleCB, tempProcessCode, tempProcessNameCode, tempRefCodePVDetail, qReference
Dim rsReferencePVDetail, donePVDetail, qNormalisasiDetail, qVoucher, rsVoucher, qCost_Activity
Dim qVoucherDetail, rsVoucherDetail, maxLineNoDetail
Dim txtRefStringSeparator_PVDetail, txtRefStringValues_PVDetail, nQuantity, nOutput
Dim txtRefNumberKeys_PVDetail, txtRefStringKeys_PVDetail, qCost_ActivityHistory, maxLineNoContractorCB
Dim rsContract, qContract, rsAggrement, qCostContract, nRate, MSQL, rsReceive, rsCARI, rsIssue, rsCARII, nAmountIssue
Dim rsCostActivityMax, qCostActivityMax, maxLineNo

'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))
'--End Added--



'***** Added By Ukie at 25 feb 2014 *****'
Dim oQuery : oQuery = 0
Dim oRS : oRS = 0
Dim SaldoOnNurseryLedger : SaldoOnNurseryLedger = 0
Dim objectQuery : objectQuery = 0
Dim objectRS : objectRS = 0
Dim tempLineNoKey : tempLineNoKey = 0
Dim SaldoOnIssue : SaldoOnIssue = 0
'***** End Added By Ukie at 25 feb 2014 *****'
%>
<font face="Verdana" size="1">
Nursery Last Period Close <%=FormatDateTime(Session("cls_Nursery"),1)%> <br> 
Nursery Current Period is <%=FormatDateTime(Session("v_Nursery"),1)%> <br><br>
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br>" 

dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl


' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('General_Charges','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for General_Charges " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 

	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='Costbook' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Costbook " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	' Remarked by NAF, as per New Concept, 7 May 2013
	' For Workshop Closing, for 2nd Book in Accounting

	tempProcessCode = "NURSERY_CLOSING"
	tempProcessNameCode = "p_Nursery_Closing.asp"
	tempRefCodePVDetail = "NURSERYLEDGER"

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


	'---- Control Job
	qControlJob="select ControlJob from Module where NAME='Nursery'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For Nursery is Empty"
	   txtControlJob=""
	end if
	response.Flush()

	if txtControlJob<>"" then
		'---- Normalisasi
		Response.Write " Normalisasi Nursery Base Costing" & "<br>"
		Response.Flush()
		
		' qNormalisasi=" Delete From COST_ACTIVITY where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Nursery")) &  _
				 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Nursery")) & " and fcBA='" & Session("fcBA") & "'" & _
				 ' " and SOURCETYPE = 'MN'"
		' connstr.execute(qNormalisasi)
		' connStr.execute("Commit")
		
		qNormalisasiDetail="   update Cost_Activity set deleteflag = 'REMOVE' where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Nursery")) &  _
				 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Nursery")) & " and fcBA='" & Session("fcBA") & "'" & _
				 " and SOURCETYPE = 'MN'"
		connstr.execute(qNormalisasiDetail)
		connStr.execute("Commit")
		
		qNormalisasiDetail = " delete cost_activity where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

		connstr.execute(qNormalisasiDetail)
		connStr.execute("Commit")
		
		
		
		Response.Write " Normalisasi Nursery Ledger" & "<br/>"
		Response.Flush()
		
		qNormalisasi = " update nurseryledger" & _
		               " set qtyreceived = 0" & _
					   "     , qtyculled = 0" & _
					   "     , qtyadjustment = 0" & _
					   "     , qtyissued = 0" & _
					   "     , costperseedling = 0" & _
					   "     , amountreceived = 0" & _
					   "     , amountissued = 0" & _
					   "     , fcedit = '" & Session("CodeName") & "'" & _
					   "     , lastupdate = sysdate" & _
					   "     , lasttime = to_char(sysdate, 'hh24:mi')" & _					   
		               " where month = " & Month(Session("v_Nursery")) & "" & _
					   "       and year = " & Year(Session("v_Nursery")) & "" & _
					   "       and fcBA = '" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)
		
		'Qty Order
		MSQL=" SELECT RD.LOCATIONCODE AS NURSERYCODE,RD.PURCHASEITEMCODE AS ITEMCODE,TO_CHAR(RN.RNDATE,'YYYY') AS NYEAR," & _
			" TO_CHAR(RN.RNDATE,'MM') AS NMONTH,nvl(SUM(RD.QUANTITY),0) AS QUANTITY " & _
			" FROM RECEIVENOTE RN,RECEIVENOTEDETAIL RD WHERE " & _
			" (RN.RECEIVENOTECODE=RD.RECEIVENOTECODE AND RN.FCBA=RD.FCBA) AND " & _
			" to_number(to_Char(RN.RNDATE,'mm'))=" & Month(Session("v_Nursery")) &  _
			" and to_number(to_Char(RN.RNDATE,'yyyy'))=" & Year(Session("v_Nursery")) & _
			" and rd.LOCATIONTYPE='" & "MN" & "'" & _
			" and RN.fcBA='" & Session("fcBA") & "'" & _
			" GROUP BY RD.LOCATIONCODE,RD.PURCHASEITEMCODE,TO_CHAR(RN.RNDATE,'YYYY'),TO_CHAR(RN.RNDATE,'MM')"	  
		'RESPONSE.Write(MSQL)
		'RESPONSE.Flush()
		Set rsReceive = Server.CreateObject("ADODB.Recordset")
		rsReceive.CursorType = 3
		rsReceive.CursorLocation = 3
		rsReceive.Open MSQL, connstr
		'JIKA DATA BELUM ADA		
		if not rsReceive.eof then
			Do While Not rsReceive.eof
					'RESPONSE.Write("A")
				q="SELECT NURSERYCODE FROM NURSERYLEDGER WHERE NURSERYCODE ='" & rsReceive("NURSERYCODE") & _
				  "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
				Set rsCARI = Server.CreateObject("ADODB.Recordset")
				rsCARI.CursorType = 3
				rsCARI.CursorLocation = 3
				rsCARI.Open q, connstr	 
				if rsCARI.eof Then
					q= " Insert into NURSERYLEDGER " & _
					   " (NURSERYCODE, YEAR, MONTH, QTYRECEIVED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					   " values ('"  &  rsReceive("NURSERYCODE") & "',"  & rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
					   rsReceive("QUANTITY") & ",'"  & _
					   Session("CodeName") & "','" & Session("CodeName") & "','" & _
					   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
					  ' Response.Write(q)
					  ' response.Flush()
					   connStr.execute(q)
				Else
					q= "Update NURSERYLEDGER set QTYRECEIVED =" & rsReceive("QUANTITY") & _
					   " where  NURSERYCODE ='"  &  rsReceive("NURSERYCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
					   " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
					connStr.execute(q)
					'RESPONSE.Write(q)
					'RESPONSE.Flush()
					connStr.execute("commit")
				End if
				rsCari.Close
				set rsCari=nothing
				rsReceive.MOVENEXT
			loop
			rsReceive.close
			set rsReceive=Nothing
		End if
		Response.Write "<br><font color=Blue>" & "Closing Order Nursery is completed..." & "</font>"
		
		'Qty Issue 
		
		

		MSQL= " select nid.nurserycode" & _
			  "        , to_number(to_char(nih.issuedate, 'mm')) as nmonth" & _
			  "        , to_number(to_char(nih.issuedate, 'yyyy')) as nyear" & _
			  "        , nvl(sum(nvl(nid.quantity, 0)), 0) quantity" & _
			  " from nurseryissuedtl nid" & _
			  "      , nurseryissuehdr nih" & _
			  " where nid.fcba = nih.fcba" & _
			  "       and nid.documentno = nih.documentno" & _
			  "       and nid.fcba = '" & Session("fcBA") & "'" & _
			  "       and to_number(to_char(nih.issuedate, 'mm')) = " & Month(Session("v_Nursery")) & "" & _
			  "       and to_number(to_char(nih.issuedate, 'yyyy')) = " & Year(Session("v_Nursery")) & "" & _
			  " group by nid.nurserycode" & _
			  "          , to_number(to_char(nih.issuedate, 'mm'))" & _
			  "          , to_number(to_char(nih.issuedate, 'yyyy'))"
		'RESPONSE.Write(MSQL)
		'RESPONSE.Flush()
		Set rsReceive = Server.CreateObject("ADODB.Recordset")
		rsReceive.CursorType = 3
		rsReceive.CursorLocation = 3
		rsReceive.Open MSQL, connstr
		'JIKA DATA BELUM ADA		
		if not rsReceive.eof then
			Do While Not rsReceive.eof
					'RESPONSE.Write("A")
				q="SELECT NURSERYCODE FROM NURSERYLEDGER WHERE NURSERYCODE ='" & rsReceive("NURSERYCODE") & _
				  "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
				Set rsCARI = Server.CreateObject("ADODB.Recordset")
				rsCARI.CursorType = 3
				rsCARI.CursorLocation = 3
				rsCARI.Open q, connstr	 
				if rsCARI.eof Then
					q= " Insert into NURSERYLEDGER " & _
					   " (NURSERYCODE, YEAR, MONTH, QTYISSUED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					   " values ('"  &  rsReceive("NURSERYCODE") & "',"  & rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
					   rsReceive("QUANTITY") & ",'"  & _
					   Session("CodeName") & "','" & Session("CodeName") & "','" & _
					   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
					   'Response.Write(q)
					   'response.Flush()
					   connStr.execute(q)
				Else
					q= "Update NURSERYLEDGER set QTYISSUED =" & rsReceive("QUANTITY") & _
					   " where  NURSERYCODE ='"  &  rsReceive("NURSERYCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
					   " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
					connStr.execute(q)
					'RESPONSE.Write(q)
					'RESPONSE.Flush()
					connStr.execute("commit")
				End if
				rsCari.Close
				set rsCari=nothing
				rsReceive.MOVENEXT
			loop
			rsReceive.close
			set rsReceive=Nothing
		End if
		Response.Write "<br><font color=Blue>" & "Closing Issue Nursery is completed..." & "</font>"
		
		'Qty Culling
		MSQL= " select ncd.nurserycode" & _
			  "        , to_number(to_char(nch.cullingdate, 'mm')) as nmonth" & _
			  "        , to_number(to_char(nch.cullingdate, 'yyyy')) as nyear" & _
			  "        , nvl(sum(nvl(ncd.quantity, 0)), 0) quantity" & _
			  " from nurseryculldtl ncd" & _
			  "      , nurserycullhdr nch" & _
			  " where ncd.fcba = nch.fcba" & _
			  "       and ncd.documentno = nch.documentno" & _
			  "       and ncd.fcba = '" & Session("fcBA") & "'" & _
			  "       and to_number(to_char(nch.cullingdate, 'mm')) = " & Month(Session("v_Nursery")) & "" & _
			  "       and to_number(to_char(nch.cullingdate, 'yyyy')) = " & Year(Session("v_Nursery")) & "" & _
			  " group by ncd.nurserycode" & _
			  "          , to_number(to_char(nch.cullingdate, 'mm'))" & _
			  "          , to_number(to_char(nch.cullingdate, 'yyyy'))"
		'RESPONSE.Write(MSQL)
		'RESPONSE.Flush()
		Set rsReceive = Server.CreateObject("ADODB.Recordset")
		rsReceive.CursorType = 3
		rsReceive.CursorLocation = 3
		rsReceive.Open MSQL, connstr
		'JIKA DATA BELUM ADA		
		if not rsReceive.eof then
			Do While Not rsReceive.eof
					'RESPONSE.Write("A")
				q="SELECT NURSERYCODE FROM NURSERYLEDGER WHERE NURSERYCODE ='" & rsReceive("NURSERYCODE") & _
				  "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
				Set rsCARI = Server.CreateObject("ADODB.Recordset")
				rsCARI.CursorType = 3
				rsCARI.CursorLocation = 3
				rsCARI.Open q, connstr	 
				if rsCARI.eof Then
					q= " Insert into NURSERYLEDGER " & _
					   " (NURSERYCODE, YEAR, MONTH, QTYCULLED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					   " values ('"  &  rsReceive("NURSERYCODE") & "',"  & rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
					   rsReceive("QUANTITY") & ",'"  & _
					   Session("CodeName") & "','" & Session("CodeName") & "','" & _
					   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
					   'Response.Write(q)
					   'response.Flush()
					   connStr.execute(q)
				Else
					q= "Update NURSERYLEDGER set QTYCULLED =" & rsReceive("QUANTITY") & _
					   " where  NURSERYCODE ='"  &  rsReceive("NURSERYCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
					   " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
					connStr.execute(q)
					'RESPONSE.Write(q)
					'RESPONSE.Flush()
					connStr.execute("commit")
				End if
				rsCari.Close
				set rsCari=nothing
				rsReceive.MOVENEXT
			loop
			rsReceive.close
			set rsReceive=Nothing
		End if
		Response.Write "<br><font color=Blue>" & "Closing Culling Nursery is completed..." & "</font>"
		
		'Qty Adjustment
		MSQL = " select nad.nurserycode" & _
			   "        , to_number(to_char(nah.adjustmentdate, 'mm')) as nmonth" & _
			   "        , to_number(to_char(nah.adjustmentdate, 'yyyy')) as nyear" & _
			   "        , nvl(sum(nvl(nad.quantity, 0)), 0) quantity" & _
			   " from nurseryadjdtl nad" & _
			   "      , nurseryadjhdr nah" & _
			   " where nad.fcba = nah.fcba" & _
			   "       and nad.documentno = nah.documentno" & _
			   "       and nad.fcba = '" & Session("fcBA") & "'" & _
			   "       and to_number(to_char(nah.adjustmentdate, 'mm')) = " & Month(Session("v_Nursery")) & "" & _
			   "       and to_number(to_char(nah.adjustmentdate, 'yyyy')) = " & Year(Session("v_Nursery")) & "" & _
			   " group by nad.nurserycode" & _
			   "          , to_number(to_char(nah.adjustmentdate, 'mm'))" & _
			   "          , to_number(to_char(nah.adjustmentdate, 'yyyy'))"
		'RESPONSE.Write(MSQL)
		'RESPONSE.Flush()
		Set rsReceive = Server.CreateObject("ADODB.Recordset")
		rsReceive.CursorType = 3
		rsReceive.CursorLocation = 3
		rsReceive.Open MSQL, connstr
		'JIKA DATA BELUM ADA		
		if not rsReceive.eof then
			Do While Not rsReceive.eof
					'RESPONSE.Write("A")
				q = " select nurserycode" & _
				    " from nurseryledger" & _
				    " where nurserycode = '" & rsReceive("NURSERYCODE") & "'" & _
					"       and year = " & rsReceive("NYEAR") & "" & _
					"       and month = "  &  rsReceive("NMONTH") & "" & _
					"       and fcba = '" & Session("fcBA") & "'"
				Set rsCARI = Server.CreateObject("ADODB.Recordset")
				rsCARI.CursorType = 3
				rsCARI.CursorLocation = 3
				rsCARI.Open q, connstr	 
				if rsCARI.eof Then
					q = " insert into nurseryledger (" & _
					    " nurserycode, year, month, qtyadjustment, fcentry, fcedit, fcip, fcba, lastupdate, lasttime" & _
					    " ) values (" & _
						" '" & rsReceive("NURSERYCODE") & "'" & _
						" , " & rsReceive("NYEAR") & "" & _
						" , " & rsReceive("NMONTH") & "" & _
						" , " & rsReceive("QUANTITY") & "" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						" , '" & Session("fcBA") & "'" & _
						" , sysdate" & _
						" , to_char(sysdate, 'HH24:MI')" & _
						" )"
					   'Response.Write(q)
					   'response.Flush()
					   connStr.execute(q)
					   connStr.execute("commit")
				Else
					q = " update nurseryledger" & _
					    " set qtyadjustment = " & rsReceive("QUANTITY") & "" & _
						" where nurserycode = '" & rsReceive("NURSERYCODE") & "'" & _
						"       and year = " & rsReceive("NYEAR") & "" & _
						"       and month = " & rsReceive("NMONTH") & "" & _
						"       and fcba = '" & Session("fcBA") & "'"
					connStr.execute(q)
					'RESPONSE.Write(q)
					'RESPONSE.Flush()
					connStr.execute("commit")
				End if
				rsCari.Close
				set rsCari = nothing
				rsReceive.MOVENEXT
			loop
			rsReceive.close
			set rsReceive=Nothing
		End if
		Response.Write "<br><font color=Blue>" & "Closing Adjustment Nursery is completed..." & "</font>"

		'-------- Capture Cost To Nursery
		MSQL=" SELECT TARGETCODE as NURSERYCODE,TO_CHAR(TDATE,'YYYY') AS NYEAR," & _
			" TO_CHAR(TDATE,'MM') AS NMONTH,nvl(SUM(AMOUNT),0) AS AMOUNT" & _
			" FROM COST_ACTIVITY WHERE " & _
			" to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Nursery")) &  _
			" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Nursery")) & _
			" and TARGETTYPE='" & "MN" & "'" & _
			" and fcBA='" & Session("fcBA") & "'" & _
			" GROUP BY TARGETCODE,TO_CHAR(TDATE,'YYYY'),TO_CHAR(TDATE,'MM')"	  
		'RESPONSE.Write(MSQL)
		'RESPONSE.Flush()
		Set rsReceive = Server.CreateObject("ADODB.Recordset")
		rsReceive.CursorType = 3
		rsReceive.CursorLocation = 3
		rsReceive.Open MSQL, connstr
		'JIKA DATA BELUM ADA		
		if not rsReceive.eof then
			Do While Not rsReceive.eof
					'RESPONSE.Write("A")
				q="SELECT NURSERYCODE FROM NURSERYLEDGER WHERE NURSERYCODE ='" & rsReceive("NURSERYCODE") & _
				  "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
				Set rsCARI = Server.CreateObject("ADODB.Recordset")
				rsCARI.CursorType = 3
				rsCARI.CursorLocation = 3
				rsCARI.Open q, connstr	 
				if rsCARI.eof Then
					q= " Insert into NURSERYLEDGER " & _
					   " (NURSERYCODE, YEAR, MONTH, AMOUNTRECEIVED,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					   " values ('"  &  rsReceive("NURSERYCODE") & "',"  & rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
					   rsReceive("AMOUNT") & ",'" & _
					   Session("CodeName") & "','" & Session("CodeName") & "','" & _
					   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
					  ' Response.Write(q)
					  ' response.Flush()
					   connStr.execute(q)
				Else
					q= "Update NURSERYLEDGER set AMOUNTRECEIVED=" & rsReceive("AMOUNT") & _
					   " where  NURSERYCODE ='"  &  rsReceive("NURSERYCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
					   " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
					connStr.execute(q)
					'RESPONSE.Write(q)
					'RESPONSE.Flush()
					connStr.execute("commit")
				End if
				rsCari.Close
				set rsCari=nothing
				rsReceive.MOVENEXT
			loop
			rsReceive.close
			set rsReceive=Nothing
		End if
		Response.Write "<br><font color=Blue>" & "Capture Cost to Nursery is completed..." & "</font>"
		
		'Update Amt Issue di Tabel Nursery Ledger
		Dim rsAmtissue,nHargapokok,nAmtissue,nRp,nQty
		nRp=0
		nQty=0
		nHargapokok=0
		nAmtissue=0
		q= " SELECT NURSERYCODE, YEAR, MONTH, OPENINGQTY, OPENINGAMOUNT, QTYRECEIVED, QTYCULLED, qtyadjustment, QTYISSUED, COSTPERSEEDLING, AMOUNTRECEIVED, AMOUNTISSUED " & _ 
		   " FROM NURSERYLEDGER WHERE  FCBA= '" & Session("fcBA") & "'" & _
		   " AND YEAR="  &  Year(Session("v_Nursery")) & " AND MONTH="  &  Month(Session("v_Nursery"))
			Set rsAmtissue = Server.CreateObject("ADODB.Recordset")
			rsAmtissue.CursorType = 3
			rsAmtissue.CursorLocation = 3
			rsAmtissue.Open q, connstr
			Do While Not rsAmtissue.eof 
				nRp=(cdbl(rsAmtissue("OPENINGAMOUNT")) + cdbl(rsAmtissue("AMOUNTRECEIVED")))
				nQty=((cdbl(rsAmtissue("OPENINGQTY")) + cdbl(rsAmtissue("QTYRECEIVED")))+cdbl(rsAmtissue("qtyadjustment"))-Cdbl(rsAmtissue("QTYCULLED")))
				' Add modification by NAF, to handle divided by Zero value, 2012-11-29
				if nQty = 0 then
					nQty = 1
				end if
				' End of Modification 
				
				nHargapokok=nRp/nQty
				nAmtissue=Round(nRp/nQty * cdbl(rsAmtissue("QTYISSUED")), 0)
			   q="Update NURSERYLEDGER set AMOUNTISSUED =" & nAmtissue & ",COSTPERSEEDLING=" & nHargaPokok & "" & _
				" where  NURSERYCODE='"  &  rsAmtissue("NURSERYCODE") &  "' AND YEAR="  &  _
				rsAmtissue("YEAR") & " AND MONTH="  &  rsAmtissue("MONTH") & " AND FCBA= '" & Session("fcBA") & "'"
				'RESPONSE.Write(q)
				RESPONSE.Flush()
				connStr.execute(q)
				
				connStr.execute("commit")			
				response.Flush()
			rsAmtissue.MOVENEXT
		loop	
		rsAmtissue.Close
		set rsAmtissue=nothing
		Response.Write "<br><font color=Blue>" & "Update Cost Per Seedling is completed..." & "</font>"
		
		'Update Nilai Amount di IssueDetails
		MSQL =  " select nid.documentno" & _
		        "        , nid.linenokey" & _
		        "        , nid.nurserycode" & _
		        "        , to_char(nih.issuedate,'yyyy') as nyear" & _
		        "        , to_char(nih.issuedate,'mm') as nmonth" & _
		        "        , sum(quantity) as quantity" & _
				" from nurseryissuedtl nid" & _
				"      , nurseryissuehdr nih" & _
		        " where nid.fcba = nih.fcba" & _
				"       and nid.documentno = nih.documentno" & _
				"       and to_number(to_Char(issuedate,'mm')) = " & Month(Session("v_Nursery")) & "" & _
				"       and to_number(to_Char(issuedate,'yyyy')) = " & Year(Session("v_Nursery")) & "" & _
				"       and nid.fcba = '" & Session("fcBA") & "'" & _
				" group by nid.documentno" & _
		        "          , nid.linenokey" & _
		        "          , nid.nurserycode" & _
		        "          , to_char(nih.issuedate,'yyyy')" & _
		        "          , to_char(nih.issuedate,'mm')"
		'RESPONSE.Write(MSQL)
		'RESPONSE.Flush()
		Set rsIssue = Server.CreateObject("ADODB.Recordset")
		rsIssue.CursorType = 3
		rsIssue.CursorLocation = 3
		rsIssue.Open MSQL, connstr
		'JIKA DATA BELUM ADA		
		if not rsIssue.eof then
			Do While Not rsIssue.eof
				'RESPONSE.Write("A")
				q="SELECT nvl(sum(COSTPERSEEDLING),0) as COSTPERSEEDLING  FROM NURSERYLEDGER WHERE NURSERYCODE='"  &  rsIssue("NURSERYCODE") & "' AND YEAR="  &  _
				  rsIssue("NYEAR") & " AND MONTH="  &  rsIssue("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
				'response.Write(q)		
				'response.Flush()
				Set rsCARII = Server.CreateObject("ADODB.Recordset")
				rsCARII.CursorType = 3
				rsCARII.CursorLocation = 3
				rsCARII.Open q, connstr	 
				nAmountIssue=0
				if Not rsCARII.eof Then
				   nAmountIssue = Round(cdbl(rsCARII("COSTPERSEEDLING"))*cdbl(rsIssue("QUANTITY")), 0)
				   q = " update nurseryissuedtl" & _
                       " set amount = " & nAmountIssue & "" & _
					   " where documentno = '" & rsIssue("documentno") & "'" & _
					   "       and linenokey = '" & rsIssue("linenokey") & "'" & _
					   "       and nurserycode = '" & rsIssue("nurserycode") & "'" & _
					   "       and fcba = '" & Session("fcBA") & "'"
					  'response.Write(q)
					  'response.Flush()
					connStr.execute(q)
					connStr.execute("commit")
				End if
				rsCARII.Close
				set rsCARII=nothing
				rsIssue.MOVENEXT
			loop
			rsIssue.close
			set rsIssue=Nothing
		End if
		
		
		
		'***** Added By Ukie at 25 feb 2014 *****'
		oQuery = 	" select nurserycode" & _
					"        , year" & _
					"        , month" & _
					"        , nvl(sum(openingqty), 0) openingqty" & _
					"        , nvl(sum(qtyreceived), 0) qtyreceived" & _
					"        , nvl(sum(qtyadjustment), 0) qtyadjustment" & _
					"        , nvl(sum(qtyculled)*-1, 0) qtyculled" & _
					"        , nvl(sum(qtyissued)*-1, 0) qtyissued" & _
					"        , nvl(sum(openingqty), 0) + nvl(sum(qtyreceived), 0) + nvl(sum(qtyadjustment), 0) + nvl(sum(qtyculled)*-1, 0) + nvl(sum(qtyissued)*-1, 0) qtysaldo" & _
					"        , nvl(sum(amountissued), 0) amtissue" & _
					"        , nvl(sum(openingamount), 0) + nvl(sum(amountreceived), 0) + nvl(sum(amountissued)*-1, 0) amtsaldo" & _
					" from nurseryledger" & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and year = " & Year(Session("v_Nursery")) & "" & _
					"       and month = " & Month(Session("v_Nursery")) & "" & _
					" group by nurserycode" & _
					"          , year" & _
					"          , month" & _
					" "
					
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open oQuery, connstr
		
		If Not oRS.EOF Then
			Do While Not oRS.EOF
				SaldoOnNurseryLedger = Cdbl(oRS("amtissue"))
				objectQuery =	" select nvl(max(nid.linenokey), 0.000000001) linenokey, nvl(sum(nid.amount), 0) amtissue" & _
								" from nurseryissuedtl nid" & _
								"      , nurseryissuehdr nih" & _
								" where nid.fcba = nih.fcba" & _
								"       and nid.documentno = nih.documentno" & _
								"       and nid.fcba = '" & Session("fcBA") & "'" & _
								"       and to_number(to_char(nih.issuedate, 'yyyy')) = " & Cdbl(oRS("year")) & "" & _
								"       and to_number(to_char(nih.issuedate, 'mm')) = " & Cdbl(oRS("month")) & "" & _
								"       and nid.nurserycode = '" & oRS("nurserycode") & "'"
								
				Set objectRS = Server.CreateObject("ADODB.Recordset")
				objectRS.Open objectQuery, connstr
				
				If Not objectRS.EOF Then
					tempLineNoKey = Cdbl(objectRS("linenokey"))
					SaldoOnIssue = Cdbl(objectRS("amtissue"))
				Else
					tempLineNoKey = 0.000000001
					SaldoOnIssue = 0
				End If
				objectRS.Close
				set objectRS = nothing
				
				If (SaldoOnNurseryLedger - SaldoOnIssue) <> 0 Then
					objectQuery =	" select nid.documentno" & _
					                "        , nid.linenokey" & _
					                "        , nid.nurserycode" & _
					                "        , sum(quantity) as quantity" & _
									" from nurseryissuedtl nid" & _
									"      , nurseryissuehdr nih" & _
					                " where nid.fcba = nih.fcba" & _
									"       and nid.documentno = nih.documentno" & _
									"       and to_number(to_char(nih.issuedate,'mm')) = " & Month(Session("v_Nursery")) & _
									"       and to_number(to_char(nih.issuedate,'yyyy')) = " & Year(Session("v_Nursery")) & "" & _
					                "       and nid.fcba = '" & Session("fcBA") & "'" & _
									"       and nid.linenokey = " & tempLineNoKey & "" & _
									" group by nid.documentno" & _
					                "          , nid.linenokey" & _
					                "          , nid.nurserycode"
									
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectQuery, connstr
					
					If Not objectRS.EOF Then
						Do While Not objectRS.EOF						
							q	=	" update nurseryissuedtl" & _
									" set amount = amount + " & SaldoOnNurseryLedger - SaldoOnIssue & _
									"     , lastupdate = sysdate" & _
									"     , lasttime = to_char(sysdate,'hh24:mi')" & _
									" where documentno = '" & objectRS("documentno") & "'" & _
									"       and linenokey = '"& objectRS("linenokey") & "'" & _
									"       and nurserycode = '"  &  objectRS("nurserycode") & "'" & _
									"       and fcba = '" & session("fcBA") & "'"

							connStr.execute(q)
						objectRS.MoveNext
						Loop
					End If
					objectRS.Close
					set objectRS = nothing
				End If		
				oRS.MoveNext
			Loop
		End If
		oRS.Close
		set oRS = nothing
		'***** End Added By Ukie at 25 feb 2014 *****'
		
		
		Response.Write "<br><font color=Blue>" & "Update Amount at Nursery Issue is completed..." & "</font>"
		
		'Insert Saldo Awal Bulan Depan
		Dim rsSaldoawal,QtySaldoawal,Bulan,Tahun,RpSaldoawal,rsCARISaldo
		q= " SELECT NURSERYCODE, YEAR, MONTH, OPENINGQTY, OPENINGAMOUNT, QTYRECEIVED, qtyadjustment, QTYCULLED, QTYISSUED, COSTPERSEEDLING, AMOUNTRECEIVED, AMOUNTISSUED " & _ 
		   " FROM NURSERYLEDGER Where MONTH=" & Month(Session("v_Nursery")) &  _
		   " and YEAR=" & Year(Session("v_Nursery")) & " and fcBA='" & Session("fcBA") & "'"
			Set rsSaldoawal = Server.CreateObject("ADODB.Recordset")
			rsSaldoawal.CursorType = 3
			rsSaldoawal.CursorLocation = 3
			rsSaldoawal.Open q, connstr

			Do While Not rsSaldoawal.eof 
				Bulan=Month(Session("v_Nursery"))
				Tahun =Year(Session("v_Nursery"))
				'Response.Write(Bulan)
				'Response.Write(Tahun)
				if bulan = 12 then
					tahun=tahun + 1
					bulan= 1
				else
					tahun=tahun
					bulan= bulan + 1
				end if
			RpSaldoawal=cdbl(rsSaldoawal("OPENINGAMOUNT"))+cdbl(rsSaldoawal("AMOUNTRECEIVED"))-cdbl(rsSaldoawal("AMOUNTISSUED"))
			QtySaldoawal=cdbl(rsSaldoawal("OPENINGQTY"))+cdbl(rsSaldoawal("QTYRECEIVED"))+cdbl(rsSaldoawal("qtyadjustment"))-cdbl(rsSaldoawal("QTYISSUED"))-cdbl(rsSaldoawal("QTYCULLED"))

			q="SELECT NURSERYCODE FROM NURSERYLEDGER WHERE NURSERYCODE='"  &  rsSaldoawal("NURSERYCODE") &  _
			  "' AND YEAR="  &  tahun & " AND MONTH="  &  bulan & " AND FCBA= '" & Session("fcBA") & "'"
			Set rsCARISaldo = Server.CreateObject("ADODB.Recordset")
			rsCARISaldo.CursorType = 3
			rsCARISaldo.CursorLocation = 3
			rsCARISaldo.Open q, connstr	 
			if rsCARISaldo.eof Then			
				q="Insert into NURSERYLEDGER (NURSERYCODE, YEAR, MONTH, OPENINGQTY, OPENINGAMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
				"values ('"  &  rsSaldoawal("NURSERYCODE") & "',"  & tahun & _
				","  &  bulan & ","  &  QtySaldoawal & ","  &  RpSaldoawal & _
				",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
				connStr.execute(q)
			else
			   q="Update NURSERYLEDGER set OPENINGQTY =" & QtySaldoawal & ",OPENINGAMOUNT=" & RpSaldoawal & "" & _
				" where  NURSERYCODE='"  &  rsSaldoawal("NURSERYCODE") & "' AND YEAR="  &  _
				tahun & " AND MONTH="  &  bulan & " AND FCBA= '" & Session("fcBA") & "'"
				connStr.execute(q)
				connStr.execute("commit")
			End if
				rsSaldoawal.MOVENEXT
		loop
		'rsCARISaldo.close
		'set rsCARISaldo	=nothing
		rsSaldoawal.Close
		set rsSaldoawal=nothing
		Response.Write "<br><font color=Blue>" & "Update Nursery Opening Balance  is completed..." & "</font>"

		
		' Added by NAF to pick the Control Job from Module table
		Dim strControlJob, rsTempControlJob, sControlJob
		strControlJob = " select controljob from module " & _
						" where " & _
						" code = '10' " 
						
		Set rsTempControlJob = Server.CreateObject("ADODB.Recordset")
		rsTempControlJob.CursorType = 3
		rsTempControlJob.CursorLocation = 3
		rsTempControlJob.Open strControlJob, connstr	 
		if rsTempControlJob.eof Then			
			sControlJob = rsTempControlJob("controljob")
		else 
			sControlJob = "1310401"
		End if				
		
		Dim strControlJobInvestasi, rsTempControlJobInvestasi, sControlJobInvestasiDr, sControlJobInvestasiCr
		
		strControlJobInvestasi = " select param_1 from parameter " & _
								 " where " & _
								 " parcode = 'MN001' and " & _
								 " fccode = 'MNINVESTDR' and " & _
								 " fcba = '" & Session("fcBA") & "'" 
		
		Set rsTempControlJobInvestasi = Server.CreateObject("ADODB.Recordset")
		rsTempControlJobInvestasi.CursorType = 3
		rsTempControlJobInvestasi.CursorLocation = 3
		rsTempControlJobInvestasi.Open strControlJobInvestasi, connstr	 
		if rsTempControlJobInvestasi.eof Then			
			sControlJobInvestasiDr = rsTempControlJobInvestasi("param_1")
		else 
			sControlJobInvestasiDr = "1310501"
		End if				


		
		strControlJobInvestasi = " select param_1 from parameter " & _
								 " where " & _
								 " parcode = 'MN001' and " & _
								 " fccode = 'MNINVESTCR' and " & _							 
								 " fcba = '" & Session("fcBA") & "'" 
		
		Set rsTempControlJobInvestasi = Server.CreateObject("ADODB.Recordset")
		rsTempControlJobInvestasi.CursorType = 3
		rsTempControlJobInvestasi.CursorLocation = 3
		rsTempControlJobInvestasi.Open strControlJobInvestasi, connstr	 
		if rsTempControlJobInvestasi.eof Then			
			sControlJobInvestasiCr = rsTempControlJobInvestasi("param_1")
		else 
			sControlJobInvestasiCr = "1310402"
		End if		
		
						
		
		'----- Issue Bibitan
		'----- DB. Activity Sesuai User Pilih Ke Job Mana
		'----------- Cr. Control Job 1310401 TRANSFER BIBIT KE LAPANGAN 
		
		'----- Tranfer to Activity Costing Debet ke Lap, Nota dll tergantung Pemakaian
		Response.Flush()
		
		

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
		
		
		qABC = " INSERT INTO COST_ACTIVITY (" & _
			   " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
			   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, " & _
								" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								" LASTUPDATE, LASTTIME, LINENO ) " & _
			   " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
				  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, " & _
				  " Cost_Activity_Info_Api.Get_Title_Reference_01 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),TARGETCODE, " & _
				  " Cost_Activity_Info_Api.Get_Title_Reference_02 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),referenceno, " & _
				  " '','', " & _
				  " '','', " & _
				  " '','', " & _
			      " decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr'), " & _
				  Month(Session("v_Nursery")) & ",  " & _
				  Year(Session("v_Nursery")) & " ,  " & _
				  " sysdate LASTUPDATE, to_char(sysdate,'HH:MI') LASTTIME, rownum + " & maxLineNo & " LINENO " & _
				  " FROM ( " & _	
					  " SELECT ISSUEDATE tdate, " & _
							   " JOBCODE jobcode, " & _
							   " 'MN' source, " & _
							   " 'MN' sourcetype, " & _
							   " NURSERYCODE sourcecode, " & _ 
							   " '' sourcecode_status, " & _
							   " LOCATIONTYPE targettype, " & _
							   " LOCATIONCODE targetcode, " & _
							   " '' TARGETCODE_STATUS, " & _
							   " nid.documentno referenceno, " & _ 
							   " SUM(QUANTITY) quantity, " & _
							   " 0 output, " & _
							   " SUM(AMOUNT)*1 amount, " & _
							   " nid.FCBA," & _
							   " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
							   " FROM nurseryissuedtl nid" & _
							   "      , nurseryissuehdr nih" & _
							   " WHERE nid.fcba = nih.fcba" & _
							   "       and nid.documentno = nih.documentno" & _
							   "       and to_number(to_Char(ISSUEDATE,'mm')) = " & Month(Session("v_Nursery")) & "" & _
							   "       and to_number(to_Char(ISSUEDATE,'yyyy')) = " & Year(Session("v_Nursery")) & "" & _
							   "       and nid.fcBA = '" & Session("fcBA") & "'" & _
							   " GROUP BY ISSUEDATE, JOBCODE, 'MN', 'MN', NURSERYCODE, '', LOCATIONTYPE, LOCATIONCODE, '', nid.documentno,nid.fcBA)"
							  
		'response.Write " qABC "& qABC
		response.Write "<br><br>Transfer to Activity Based Costing..." & "<br>"
		response.Flush()		  
		connstr.Execute(qABC)
		

		Response.Flush()
		
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
		
		
		
		qABC = " INSERT INTO COST_ACTIVITY (" & _
			   " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
			   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, " & _
								" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								" LASTUPDATE, LASTTIME, LINENO ) " & _
			   " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
				  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, " & _
				  " Cost_Activity_Info_Api.Get_Title_Reference_01 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),TARGETCODE, " & _
				  " Cost_Activity_Info_Api.Get_Title_Reference_02 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),referenceno, " & _
				  " '','', " & _
				  " '','', " & _
				  " '','', " & _
			      " decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr'), " & _
				  Month(Session("v_Nursery")) & ",  " & _
				  Year(Session("v_Nursery")) & " ,  " & _
				  " sysdate LASTUPDATE, to_char(sysdate,'HH:MI') LASTTIME, rownum + " & maxLineNo & " LINENO " & _
				  " FROM ( " & _	
					  " SELECT ISSUEDATE tdate, " & _
					  " '" & sControlJob & "' jobcode, " & _
					  " 'MN' source, " & _
					  " 'MN' sourcetype, " & _
					  " NURSERYCODE sourcecode," & _ 
					  " '' sourcecode_status, " & _
					  " 'CB' targettype, " & _
					  " 'BS' targetcode, " & _
					  " '' targetcode_status, " & _
					  " nid.documentno referenceno, " & _
					  " SUM(QUANTITY) quantity, " & _
					  " 0 output, " & _
					  " SUM(AMOUNT)*-1 amount, " & _ 
					  " nid.FCBA," & _
					  " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
					  " FROM nurseryissuedtl nid" & _
				      "      , nurseryissuehdr nih" & _
				      " WHERE nid.fcba = nih.fcba" & _
					  "       and nid.documentno = nih.documentno" & _
					  "       and to_number(to_Char(ISSUEDATE,'mm'))=" & Month(Session("v_Nursery")) &  _
					  " and to_number(to_Char(ISSUEDATE,'yyyy'))=" & Year(Session("v_Nursery")) & " and nid.fcBA='" & Session("fcBA") & "'" & _
					  " GROUP BY ISSUEDATE, '" & sControlJob & "', 'MN', 'MN', NURSERYCODE, '', 'CB', 'BS', '', nid.documentno,nid.fcBA)"
		'response.Write " qABC "& qABC
		response.Flush()		  
		connstr.Execute(qABC)

		'Jurnal to Balanced ( Activa )
		'-- db. 1740100 = BIAYA AKUISISI BIBITAN
		'       -- Cr 1310499 = TRANSFER KE BIAYA AKUISISI BIBITAN
		' 1330100
		Response.Flush()
		
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
		
		
		qABC = " INSERT INTO COST_ACTIVITY (" & _
			   " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
			   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
								" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								" LASTUPDATE, LASTTIME, LINENO ) " & _
			   " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
				  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, " & _
				  " Cost_Activity_Info_Api.Get_Title_Reference_01 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),TARGETCODE, " & _
				  " Cost_Activity_Info_Api.Get_Title_Reference_02 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),referenceno, " & _
				  " '','', " & _
				  " '','', " & _
				  " '','', " & _
			      " decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr'), " & _
				  Month(Session("v_Nursery")) & ",  " & _
				  Year(Session("v_Nursery")) & " ,  " & _
				  " sysdate LASTUPDATE, to_char(sysdate,'HH:MI') LASTTIME, rownum + " & maxLineNo & " LINENO " & _
				  " FROM ( " & _
				  " SELECT to_Date('" & Session("v_Nursery") & "','mm/dd/yyyy') tdate, " & _
							"'" & sControlJobInvestasiDr & "' jobcode, " & _
							"'MN' source, " & _
							"'MN' sourcetype, " & _
							"NURSERYCODE sourcecode, " & _
							"'' sourcecode_status, " & _
							"'CB' targettype, " & _
							"'BS' targetcode, " & _
							"'' targetcode_status, " & _
							"'DB MN TO BS' referenceno, " & _
							"0 quantity, " & _
							"0 output, " & _
							"(SUM(AMOUNTRECEIVED)-sum(AMOUNTISSUED))*1 amount, " & _ 
							"FCBA ," & _
							"to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
				  " FROM NURSERYLEDGER " & _
				  " WHERE Month=" & Month(Session("v_Nursery")) &  _
				  " and   Year=" & Year(Session("v_Nursery")) & " and fcBA='" & Session("fcBA") & "'" & _
				  " GROUP BY to_Date('"  & Session("v_Nursery") & "','mm/dd/yyyy'), '" & sControlJobInvestasiDr & "', 'MN', 'MN', NURSERYCODE, '', 'CB', 'BS', '', 'DB MN TO BS',fcBA)"
		'response.Write " qABC "& qABC
		response.Flush()		  
		connstr.Execute(qABC)

		'1310402
		'----- Tranfer to Activity Costing Debet ke Lap, Nota dll tergantung Pemakaian
		Response.Flush()

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


		qABC = " INSERT INTO COST_ACTIVITY (" & _
			   " TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
			   " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,  " & _
								" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								" LASTUPDATE, LASTTIME, LINENO ) " & _
			   " select TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, " & _
				  " REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, sysdate PROCESS_DATE, " & _
				  " Cost_Activity_Info_Api.Get_Title_Reference_01 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),TARGETCODE, " & _
				  " Cost_Activity_Info_Api.Get_Title_Reference_02 (SOURCETYPE,TARGETTYPE,decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr')),referenceno, " & _
				  " '','', " & _
				  " '','', " & _
				  " '','', " & _
			      " decode ( GreaterOrEqualThan ( amount, 0), 'TRUE','Dr','Cr'), " & _
				  Month(Session("v_Nursery")) & ",  " & _
				  Year(Session("v_Nursery")) & " ,  " & _
				  " sysdate LASTUPDATE, to_char(sysdate,'HH:MI') LASTTIME, rownum + " & maxLineNo & " LINENO " & _
				  " FROM ( " & _
					  " SELECT to_Date('" & Session("v_Nursery") & "','mm/dd/yyyy') tdate, " & _
					  "'" & sControlJobInvestasiCr & "' jobcode, " & _
					  "'MN' source, " & _
					  "'MN' sourcetype, " & _
					  "NURSERYCODE sourcecode, " & _
					  "'' sourcecode_status, " & _
					  "'CB' targettype, " & _
					  "'BS' targetcode, " & _
					  "'' targetcode_status, " & _
					  "'CR MN TO BS' referenceno, " & _
					  "0 quantity, " & _
					  "0 output, " & _
					  "(SUM(AMOUNTRECEIVED)-sum(AMOUNTISSUED))*-1 amount, " & _
					  "FCBA, " & _
					  "to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date " & _
					  " FROM NURSERYLEDGER " & _
					  " WHERE Month=" & Month(Session("v_Nursery")) &  _
					  " and   Year=" & Year(Session("v_Nursery")) & " and fcBA='" & Session("fcBA") & "'" & _
					  " GROUP BY to_Date('"  & Session("v_Nursery") & "','mm/dd/yyyy'), '" & sControlJobInvestasiCr & "', 'MN', 'MN', NURSERYCODE, '', 'CB', 'BS', '', 'CR MN TO BS',fcBA)"
		'response.Write " qABC "& qABC
		response.Write "<br><br>Transfer to Nursery Acquisition Cost..." & "<br>"
		response.Flush()		  
		connstr.Execute(qABC)


		' Add By Akbar 18 Desemeber 2015
		
			Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear  " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='NURSERY' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for NURSERY " & "<br>"
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
				   " where UPPER(system) ='NURSERY' and fcba='" & Session("fcBA") & "'"
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where UPPER(system) ='NURSERY' and fcba='" & Session("fcBA") & "'"
				ELSE
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
					   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where  UPPER(system) ='NURSERY' and fcba='" & Session("fcBA") & "'"
			    END IF
			END IF
			
			IF temp_NextCloseAccountingPeriod = 13 AND temp_NextCurrentAccountingPeriod = 1  THEN
			
				q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where UPPER(system) ='NURSERY' and and fcba='" & Session("fcBA") & "'"
			END IF
		
        'connStr.execute(q)

		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "<br>Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
		
		'--Added by Ukie at 25-Oct-2012--	
		Endingtime = FormatDateTime(CDate(Now))

		qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Nursery'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, lastupdate, lasttime, fcentry) VALUES (" & _
						" 'Nursery', " & _
						" '" & Session("fcBA") & "'," & _
						" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" " & tempMaxLine & "" & _
						" , sysdate" & _
						" , to_char(sysdate, 'hh24:mi')" & _
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
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Nursery'"
					
		Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
		rsLenghtime.CursorType = 3
		rsLenghtime.CursorLocation = 3
		rsLenghtime.Open strLenghtime, connstr

		upClosingLog = " UPDATE CLOSINGLOG SET" & _
						" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Nursery'"		
		connStr.execute(upClosingLog)
		'--End Added--
		
		%>
	<%end if%>

	<%if err.number <> 0 then%>
		<br><img src="images/error.gif"><%
		'connStr.RollbackTrans
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		cmdView	= "True"%>

	<%Else%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Nursery Closing is succesfully...
		
		<%
		
			txtLastDay  =cmdFindPeriod("Nursery")
			txtNextMonth=cmdPeriodClose("Nursery")	
			
			Session("v_Nursery")=txtLastDay
			Session("cls_Nursery")= txtNextMonth			
		%>
		
	<%End if%>
<%
else 
%>
	<br><br><img src="images/error.gif">Nursery Closing is NOT completed...
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
<%
Response.Write "<br/>" & Session("v_Nursery")
Response.Write "<br/>" & Session("cls_Nursery")
%>