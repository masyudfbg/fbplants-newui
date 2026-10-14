<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Store Closing</title>
</head>
<body>
<%
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Store Last Period Close <%=FormatDateTime(Session("cls_Stores"),1)%> <br> 
Store Current Period is <%=FormatDateTime(Session("v_Stores"),1)%> <br><br>
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br><br>"
TimeAwal=Time()

'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Dim fine_result_01, fine_result_02, fine_result_03

Startingtime = FormatDateTime(CDate(Now))
'--End Added--	

Dim TDATE, STORAGE_CODE, QUANTITY_IN, QUANTITY_OUT, REFERENCES, REMARKS
DIM MSQL, rsReceive,rsWrite,rsIssue
DIM TGL,q,rsCARIW,rsCARII
Dim strCursorSivCode, rsCursorSivCode
Dim strCursorCredit, rsCursorCredit, strInsertCredit
Dim tempTotalAmountCredit, tempSIVCODE, qInsertCredit, strCursorCount, rsCursorCount, totalRecord, totalRowInserted, tempSelisih
Dim temp_Db_Job, temp_Db_TargetType, temp_Db_TargetCode
Dim temp_Cr_Job, temp_Cr_TargetType, temp_Cr_TargetCode
Dim queryWriteOffDetail, qCost_ActivityCr
Dim rsRealPrice,realprice,realAmount
dim AllQty, AllAmount

Dim txdate

'***** Added By Ukie at 25 feb 2014 *****'
Dim oQuery : oQuery = 0
Dim oRS : oRS = 0
Dim SaldoOnStockLedger : SaldoOnStockLedger = 0
Dim objectQuery : objectQuery = 0
Dim objectRS : objectRS = 0
Dim tempLineNokey : tempLineNokey = 0
Dim SaldoOnIssue : SaldoOnIssue = 0
'***** End Added By Ukie at 25 feb 2014 *****'



' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl


' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Stores','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Stores " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 


	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='Stores' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Stores " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	response.Write "Current Period for Stock Ledger is " & FormatDateTime(Session("v_Stores"),1) & "<br>"


	' qNormalisasiDetail = " delete from cost_activity_hist_tab " & _
						 ' " where " & _
						 ' " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
								' " where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Stores")) &  _
								' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Stores")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='ST')"
						 
	' connstr.execute(qNormalisasiDetail)
    ' connStr.execute("Commit")

	' qNormalisasi=" Delete From COST_ACTIVITY where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Stores")) &  _
				 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Stores")) & " and fcBA='" & Session("fcBA") & "'" & _
				 ' " and SOURCETYPE = 'ST'"
	' connstr.execute(qNormalisasi)
    ' connStr.execute("Commit")
    
	
	qNormalisasiDetail = " update cost_activity_hist_tab set deleteflag = 'REMOVE'" & _
						 " where " & _
						 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
								" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Stores")) &  _
								" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Stores")) & " and fcBA='" & Session("fcBA") & "' and SOURCETYPE='ST')"
						 
	connstr.execute(qNormalisasiDetail)
    connStr.execute("Commit")
	
	qNormalisasiDetail = " delete from cost_activity_hist_tab where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

	connstr.execute(qNormalisasiDetail)
	connStr.execute("Commit")

	qNormalisasiDetail=" update Cost_Activity set deleteflag = 'REMOVE' where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Stores")) &  _
				 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Stores")) & " and fcBA='" & Session("fcBA") & "'" & _
				 " and SOURCETYPE = 'ST'"
	connstr.execute(qNormalisasiDetail)
    connStr.execute("Commit")

	qNormalisasiDetail = " delete cost_activity where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

	connstr.execute(qNormalisasiDetail)
	connStr.execute("Commit")
	
	' To Pick max Line No in Cost_Activity
	qMaxLineNo = " select nvl(max(nvl(lineno, 0)), 0) maxLineNo " & _
					" from cost_activity " & _
					" where " & _
					" fcba = '" & Session("fcBA") & "' "
					
	Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
	rsqMaxLineNo.CursorType = 3
	rsqMaxLineNo.CursorLocation = 3
	rsqMaxLineNo.Open qMaxLineNo, connstr

	if not rsqMaxLineNo.eof then
		maxLineNo = cdbl(rsqMaxLineNo("maxLineNo"))
	else
		maxLineNo = 0
	end if
	rsqMaxLineNo.close
	set rsqMaxLineNo = Nothing

    
		
	MSQL= " SELECT RD.LOCATIONCODE AS STORECODE,RD.PURCHASEITEMCODE AS ITEMCODE,TO_CHAR(RN.RNDATE,'YYYY') AS NYEAR," & _
			" TO_CHAR(RN.RNDATE,'MM') AS NMONTH,nvl(SUM(RD.QUANTITY),0) AS QUANTITY,nvl(SUM(RD.AMOUNTRECEIVED),0) AS AMOUNT" & _
			" FROM RECEIVENOTE RN,RECEIVENOTEDETAIL RD WHERE " & _
			" (RN.RECEIVENOTECODE=RD.RECEIVENOTECODE AND RN.FCBA=RD.FCBA) AND " & _
			" to_number(to_Char(RN.RNDATE,'mm'))=" & Month(Session("v_Stores")) &  _
			" and to_number(to_Char(RN.RNDATE,'yyyy'))=" & Year(Session("v_Stores")) & _
			" and rd.LOCATIONTYPE='" & "ST" & "'" & _
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
			q="SELECT STORECODE FROM STOCKLEDGER WHERE STORECODE ='"  &  rsReceive("STORECODE") & "' AND  ITEMCODE='"  & _
			  rsReceive("ITEMCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"

			' response.write q
			Set rsCARI = Server.CreateObject("ADODB.Recordset")
			rsCARI.CursorType = 3
			rsCARI.CursorLocation = 3
			rsCARI.Open q, connstr	 
			if rsCARI.eof Then
				q="Insert into STOCKLEDGER (STORECODE, ITEMCODE, YEAR, MONTH, QTYRECEIVED, AMTRECEIVED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				   "values ('"  &  rsReceive("STORECODE") & "','"  &  rsReceive("ITEMCODE") & "',"  &  rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
				   rsReceive("QUANTITY") & ","  &  rsReceive("AMOUNT") &  _
				   ",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
				   ' Response.Write(q)
				   response.Flush()
				   connStr.execute(q)
			Else
			   q="Update STOCKLEDGER set QTYRECEIVED =" & rsReceive("QUANTITY") & ",AMTRECEIVED=" & rsReceive("AMOUNT") & _
				 ", LASTUPDATE = sysdate, LASTTIME = to_char(SYSDATE,'HH24:MI') " & _
				 " where  STORECODE ='"  &  rsReceive("STORECODE") & "' AND  ITEMCODE='"  &  rsReceive("ITEMCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
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
	Response.Write "<br><font color=Blue>" & "Closing Recieve Note is completed..." & "</font>"
	
	
		
	MSQL= " SELECT RD.LOCATIONCODE AS STORECODE,RD.STOCKCODE AS ITEMCODE,TO_CHAR(RN.ORDERDATE,'YYYY') AS NYEAR," & _
			" TO_CHAR(RN.ORDERDATE,'MM') AS NMONTH,nvl(SUM(RD.QUANTITYRECEIVED),0) AS QUANTITY,nvl(SUM(RD.AMOUNTRECEIVED),0) AS AMOUNT" & _
			" FROM receivedtask_tab RN,receivedtaskdtl_tab RD WHERE " & _
			" (RN.orderno=RD.orderno AND RN.FCBA=RD.FCBA) AND " & _
			" to_number(to_Char(RN.ORDERDATE,'mm'))=" & Month(Session("v_Stores")) &  _
			" and to_number(to_Char(RN.ORDERDATE,'yyyy'))=" & Year(Session("v_Stores")) & _
			" and rd.LOCATIONTYPE='ST'" & _
			" and RN.fcBA='" & Session("fcBA") & "'" & _
			" GROUP BY RD.LOCATIONCODE,RD.STOCKCODE,TO_CHAR(RN.ORDERDATE,'YYYY'),TO_CHAR(RN.ORDERDATE,'MM')"	  
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
			q="SELECT STORECODE FROM STOCKLEDGER WHERE STORECODE ='"  &  rsReceive("STORECODE") & "' AND  ITEMCODE='"  & _
			  rsReceive("ITEMCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			Set rsCARI = Server.CreateObject("ADODB.Recordset")
			rsCARI.CursorType = 3
			rsCARI.CursorLocation = 3
			rsCARI.Open q, connstr	 
			if rsCARI.eof Then
				q="Insert into STOCKLEDGER (STORECODE, ITEMCODE, YEAR, MONTH, QTYRECEIVEDTASK, AMTRECEIVEDTASK, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				   "values ('"  &  rsReceive("STORECODE") & "','"  &  rsReceive("ITEMCODE") & "',"  &  rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
				   rsReceive("QUANTITY") & ","  &  rsReceive("AMOUNT") &  _
				   ",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
				   'Response.Write(q)
				   response.Flush()
				   connStr.execute(q)
			Else
			   q="Update STOCKLEDGER set QTYRECEIVEDTASK =" & rsReceive("QUANTITY") & ",AMTRECEIVEDTASK=" & rsReceive("AMOUNT") & _
				 ", LASTUPDATE = sysdate, LASTTIME = to_char(SYSDATE,'HH24:MI') " & _
				 " where  STORECODE ='"  &  rsReceive("STORECODE") & "' AND  ITEMCODE='"  &  rsReceive("ITEMCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
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
	Response.Write "<br><font color=Blue>" & "Closing Recieve Task is completed..." & "</font>"
	
	
	' Add by Akbar 17 Sept 2016 to Update Transport Task
    
	MSQL =  "  SELECT TTD.STORECODE," & _
			"          TTD.STOCKCODE AS ITEMCODE," & _
			"          TO_CHAR (TRANSPORTTASK_API.GET_RCVDATE (TTD.FCBA, TTD.ORDERNO),  'YYYY') AS NYEAR," & _
			"          TO_CHAR (TRANSPORTTASK_API.GET_RCVDATE (TTD.FCBA, TTD.ORDERNO), 'MM') AS NMONTH," & _
			"          NVL (SUM (TTD.QUANTITYTRANSFERED), 0) AS QUANTITY," & _
			"          NVL (ROUND(SUM (TTD.AMOUNT),2), 0) AS AMOUNT" & _
			"     FROM TRANSPORTTASKDTL_TAB TTD" & _
			"     INNER JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = TTD.ORDERNO AND TT.FCBA = TTD.FCBA)" & _
			"    WHERE     TO_CHAR (TRANSPORTTASK_API.GET_RCVDATE (TTD.FCBA, TTD.ORDERNO), 'MM') = '" & LPAD(Month(Session("v_Stores")),"0",2) & "'" & _
			"          AND TO_CHAR (TRANSPORTTASK_API.GET_RCVDATE (TTD.FCBA, TTD.ORDERNO), 'YYYY') = '" & Year(Session("v_Stores")) & "'" & _
			"          AND TTD.FCBA = '"& Session("fcBA") &"'" & _
			"          AND TT.ROWSTATEID = 'Approved'" & _
			" GROUP BY TTD.STORECODE," & _
			"          TTD.STOCKCODE," & _
			"          TO_CHAR (TRANSPORTTASK_API.GET_RCVDATE (TTD.FCBA, TTD.ORDERNO), 'YYYY')," & _
			"          TO_CHAR (TRANSPORTTASK_API.GET_RCVDATE (TTD.FCBA, TTD.ORDERNO), 'MM')"  


			' "SELECT TTD.STORECODE,TTD.STOCKCODE AS ITEMCODE, to_char(TRANSPORTTASK_API.GET_RCVDATE(TTD.FCBA,TTD.ORDERNO),'YYYY') AS NYEAR," & _
			' " to_char(TRANSPORTTASK_API.GET_RCVDATE(TTD.FCBA,TTD.ORDERNO),'MM') AS NMONTH,nvl(SUM(TTD.QUANTITYTRANSFERED),0) AS QUANTITY,nvl(SUM(TTD.AMOUNT),0) AS AMOUNT " & _
			' "  FROM TRANSPORTTASKDTL_TAB TTD WHERE " & _
			' " to_char(TRANSPORTTASK_API.GET_RCVDATE(TTD.FCBA,TTD.ORDERNO),'MM')='" & LPAD(Month(Session("v_Stores")),"0",2) & "'" &  _
			' " and to_char(TRANSPORTTASK_API.GET_RCVDATE(TTD.FCBA,TTD.ORDERNO),'YYYY')='" & Year(Session("v_Stores")) & "' and TTD.FCBA='" & Session("fcBA") & "'" & _
			' "  GROUP BY TTD.STORECODE,TTD.STOCKCODE, to_char(TRANSPORTTASK_API.GET_RCVDATE(TTD.FCBA,TTD.ORDERNO),'YYYY'), to_char(TRANSPORTTASK_API.GET_RCVDATE(TTD.FCBA,TTD.ORDERNO),'MM')"
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
			q="SELECT STORECODE FROM STOCKLEDGER WHERE STORECODE='"  &  rsIssue("STORECODE") & "' AND ITEMCODE='"  &  rsIssue("ITEMCODE") & "' AND YEAR="  &  _
			  rsIssue("NYEAR") & " AND MONTH="  &  rsIssue("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			'response.Write(q)		
			'response.Flush()
			Set rsCARII = Server.CreateObject("ADODB.Recordset")
			rsCARII.CursorType = 3
			rsCARII.CursorLocation = 3
			rsCARII.Open q, connstr	 
			if rsCARII.eof Then		
				q= "Insert into STOCKLEDGER (STORECODE, ITEMCODE, YEAR, MONTH, QTYTRANSPORTTASK, AMTTRANSPORTTASK, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				   "values ('"  &  rsIssue("STORECODE") & "','"  &  rsIssue("ITEMCODE") & "',"  &  rsIssue("NYEAR") & ","  &  rsIssue("NMONTH") & _
				   "," &  rsIssue("QUANTITY") & ","  &  rsIssue("AMOUNT") & "," & _
				   "'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI') )"
				   connStr.execute(q)
			Else
				q= "Update STOCKLEDGER set QTYTRANSPORTTASK =" & rsIssue("QUANTITY") & ",AMTTRANSPORTTASK=" & rsIssue("AMOUNT") &"" & _
				   ", LASTUPDATE = sysdate, LASTTIME = to_char(SYSDATE,'HH24:MI') " & _
				   " where  STORECODE='"  &  rsIssue("STORECODE") & "' AND ITEMCODE='"  &  rsIssue("ITEMCODE") & "' AND YEAR="  &  rsIssue("NYEAR") & _
				   " AND MONTH="  &  rsIssue("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
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
	Response.Write "<br><font color=Blue>" & "Closing Transport Task is completed..." & "</font>"
	

	'Issue
	'ISSUE=SIVCODE, STORECODE, SIVDATE, FCENTRY, FCEDIT, FCIP, FCBA
	'ISSUEDETAILS=SIVCODE, STOCKCODE, LOCATIONCODE, LOCATIONTYPE, QUANTITY, PRICE, AMOUNT, JOBCODE, FCENTRY, FCEDIT, FCIP, FCBA
	MSQL=   "SELECT ISU.STORECODE AS STORECODE,ID.STOCKCODE AS ITEMCODE,TO_CHAR(ISU.SIVDATE,'YYYY') AS NYEAR," & _
			" TO_CHAR(ISU.SIVDATE,'MM') AS NMONTH,nvl(SUM(ID.QUANTITY),0) AS QUANTITY,nvl(SUM(ID.AMOUNT),0) AS AMOUNT" & _
			" FROM ISSUE ISU,ISSUEDETAILS ID WHERE " & _
			" (ISU.SIVCODE=ID.SIVCODE AND ISU.FCBA=ID.FCBA) AND " & _
			" to_number(to_Char(ISU.SIVDATE,'mm'))=" & Month(Session("v_Stores")) &  _
			" and to_number(to_Char(ISU.SIVDATE,'yyyy'))=" & Year(Session("v_Stores")) & " and ISU.fcBA='" & Session("fcBA") & "'" & _
			" GROUP BY ISU.STORECODE,ID.STOCKCODE,TO_CHAR(ISU.SIVDATE,'YYYY'),TO_CHAR(ISU.SIVDATE,'MM')"
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
			q="SELECT STORECODE FROM STOCKLEDGER WHERE STORECODE='"  &  rsIssue("STORECODE") & "' AND ITEMCODE='"  &  rsIssue("ITEMCODE") & "' AND YEAR="  &  _
			  rsIssue("NYEAR") & " AND MONTH="  &  rsIssue("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			'response.Write(q)		
			'response.Flush()
			Set rsCARII = Server.CreateObject("ADODB.Recordset")
			rsCARII.CursorType = 3
			rsCARII.CursorLocation = 3
			rsCARII.Open q, connstr	 
			if rsCARII.eof Then		
				q="Insert into STOCKLEDGER (STORECODE, ITEMCODE, YEAR, MONTH, QTYISSUED, AMTISSUED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				  "values ('"  &  rsIssue("STORECODE") & "','"  &  rsIssue("ITEMCODE") & "',"  &  rsIssue("NYEAR") & ","  &  rsIssue("NMONTH") & _
				  "," &  rsIssue("QUANTITY") & ","  &  rsIssue("AMOUNT") & "," & _
				  "'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI') )"
				  connStr.execute(q)
			Else
				q="Update STOCKLEDGER set QTYISSUED =" & rsIssue("QUANTITY") & ",AMTISSUED=" & rsIssue("AMOUNT") &"" & _
					", LASTUPDATE = sysdate, LASTTIME = to_char(SYSDATE,'HH24:MI') " & _
					" where  STORECODE='"  &  rsIssue("STORECODE") & "' AND ITEMCODE='"  &  rsIssue("ITEMCODE") & "' AND YEAR="  &  rsIssue("NYEAR") & _
					" AND MONTH="  &  rsIssue("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
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
	Response.Write "<br><font color=Blue>" & "Closing Issue is completed..." & "</font>"



	'Write Off
	'WRITEOFF=WRITEOFFCODE, STORECODE, WODATE, FCENTRY, FCEDIT, FCIP, FCBA
	'WRITEOFFDETAILS=WRITEOFFCODE, STOCKCODE, LOCATIONCODE, LOCATIONTYPE, QUANTITY, PRICE, AMOUNT, JOBCODE, FCENTRY, FCEDIT, FCIP, FCBA
	MSQL= " SELECT WF.STORECODE AS STORECODE,WD.STOCKCODE AS ITEMCODE,TO_CHAR(WF.WODATE,'YYYY') AS NYEAR," & _
			" TO_CHAR(WF.WODATE,'MM') AS NMONTH,nvl(SUM(WD.QUANTITY),0) AS QUANTITY,nvl(SUM(WD.AMOUNT),0) AS AMOUNT" & _
			" FROM WRITEOFF WF,WRITEOFFDETAILS WD WHERE " & _
			" (WF.WRITEOFFCODE=WD.WRITEOFFCODE AND WF.FCBA=WD.FCBA) AND " & _
			" to_number(to_Char(WF.WODATE,'mm'))=" & Month(Session("v_Stores")) &  _
			" and to_number(to_Char(WF.WODATE,'yyyy'))=" & Year(Session("v_Stores")) & " and WF.fcBA='" & Session("fcBA") & "'" & _
			" GROUP BY WF.STORECODE,WD.STOCKCODE,TO_CHAR(WF.WODATE,'YYYY'),TO_CHAR(WF.WODATE,'MM')"
	'RESPONSE.Write(MSQL)
	'RESPONSE.Flush()
	Set rsWrite = Server.CreateObject("ADODB.Recordset")
	rsWrite.CursorType = 3
	rsWrite.CursorLocation = 3
	rsWrite.Open MSQL, connstr
	'JIKA DATA BELUM ADA		
	if not rsWrite.eof then
		Do While Not rsWrite.eof
			'RESPONSE.Write("A")
			q="SELECT STORECODE FROM STOCKLEDGER WHERE STORECODE='"  &  rsWrite("STORECODE") & "' AND ITEMCODE='"  &  rsWrite("ITEMCODE") & _
			  "' AND YEAR="  &  rsWrite("NYEAR") & " AND MONTH="  &  rsWrite("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			Set rsCARIW = Server.CreateObject("ADODB.Recordset")
			rsCARIW.CursorType = 3
			rsCARIW.CursorLocation = 3
			rsCARIW.Open q, connstr	 
			if rsCARIW.eof Then
				q="Insert into STOCKLEDGER (STORECODE, ITEMCODE, YEAR, MONTH, QTYWRITEOFF, AMTWRITEOFF, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				"values ('"  &  rsWrite("STORECODE") & "','"  &  rsWrite("ITEMCODE") & "',"  &  rsWrite("NYEAR") & _
				","  &  rsWrite("NMONTH") & ","  &  rsWrite("QUANTITY") & ","  &  rsWrite("AMOUNT") & _
				",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI') )"
				connStr.execute(q)
			Else
			   q="Update STOCKLEDGER set QTYWRITEOFF =" & rsWrite("QUANTITY") & ",AMTWRITEOFF=" & rsWrite("AMOUNT") &"" & _
				", LASTUPDATE = sysdate, LASTTIME = to_char(SYSDATE,'HH24:MI') " & _		   
				" where  STORECODE='"  &  rsWrite("STORECODE") & "' AND ITEMCODE='"  &  rsWrite("ITEMCODE") & "' AND YEAR="  &  rsWrite("NYEAR") & _
				" AND MONTH="  &  rsWrite("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
				connStr.execute(q)
				'RESPONSE.Write(q)
				'RESPONSE.Flush()
				connStr.execute("commit")
			End if
			  rsCARIW.Close
			  set rsCARIW=nothing
			  rsWrite.MOVENEXT
		loop
		rsWrite.close
		set rsWrite=Nothing
	End if
	Response.Write "<br><font color=Blue>" & "Closing Write Off is completed..." & "</font>"


	'Update Amt transport task, Issue di Tabel Stock Ledger
		Dim rsAmtissue,nHargapokok,nAmtissue,nRp,nQty
		nRp=0
		nQty=0
		nHargapokok=0
		nAmtissue=0
		' UPDATED BY UKIE AT 24 OKTOBER 2012 BY DROPPING LINE "QTYRECEIVED > 0"
		q="SELECT STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, QTYRECEIVED, AMTRECEIVED, QTYRECEIVEDTASK, AMTRECEIVEDTASK , QTYRETURNED,AMTRETURNED, " & _
		   "QTYWRITEOFF, AMTWRITEOFF, QTYTRANSFERED, AMTTRANSFERED, QTYISSUED, AMTISSUED, QTYTRANSPORTTASK, AMTTRANSPORTTASK " & _ 
		   "FROM STOCKLEDGER WHERE MONTH=" & Month(Session("v_Stores")) &  _
		   " and YEAR=" & Year(Session("v_Stores")) & " and fcBA='" & Session("fcBA") & "'"
			Set rsAmtissue = Server.CreateObject("ADODB.Recordset")
			rsAmtissue.CursorType = 3
			rsAmtissue.CursorLocation = 3
			rsAmtissue.Open q, connstr
			Do While Not rsAmtissue.eof 
				' Remarked by NAF to handle the Amount Write OFF as it is., 
				' nRp=(cdbl(rsAmtissue("OPENINGBALANCEAMT")) + cdbl(rsAmtissue("AMTRECEIVED")))-cdbl(rsAmtissue("AMTWRITEOFF"))
				' nQty=(cdbl(rsAmtissue("OPENINGBALANCEQTY")) + cdbl(rsAmtissue("QTYRECEIVED")))-cdbl(rsAmtissue("QTYWRITEOFF"))
				' if user input negative then it will be negative, and if user input positif then it will be positif
				' by NAF, 23 January 2014
				nRp= cdbl(rsAmtissue("OPENINGBALANCEAMT")) + cdbl(rsAmtissue("AMTRECEIVED")) + cdbl(rsAmtissue("AMTWRITEOFF")) + cdbl(rsAmtissue("AMTRECEIVEDTASK")) -cdbl(rsAmtissue("AMTTRANSPORTTASK"))
				nQty= cdbl(rsAmtissue("OPENINGBALANCEQTY")) + cdbl(rsAmtissue("QTYRECEIVED")) + cdbl(rsAmtissue("QTYWRITEOFF")) + cdbl(rsAmtissue("QTYRECEIVEDTASK")) -cdbl(rsAmtissue("QTYTRANSPORTTASK"))
				
				' Added by NAF, to handle zero value, 24-10-2012
				if nQty = 0 then 
					nQty = 1
				end if
				' End added.
				nHargapokok= nRp/nQty
				nAmtissue=round(nRp/nQty * cdbl(rsAmtissue("QTYISSUED")), 2)
			   q="Update STOCKLEDGER set AMTISSUED =" & nAmtissue & ",AMTPERUNIT=" & nHargaPokok & "" & _
				 ", LASTUPDATE = sysdate, LASTTIME = to_char(SYSDATE,'HH24:MI') " & _		   					   
				" where  STORECODE='"  &  rsAmtissue("STORECODE") & "' AND ITEMCODE='"  &  rsAmtissue("ITEMCODE") & "' AND YEAR="  &  _
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

        

		'Update Nilai Amount di IssueDetails
		MSQL =	" SELECT id.LineNoKey" & _
		        "        , isu.SIVCODE" & _
		        "        , ISU.STORECODE AS STORECODE" & _
		        "        , ID.STOCKCODE AS ITEMCODE" & _
		        "        , ID.JOBCODE" & _
		        "        , LOCATIONCODE" & _
		        "        , LOCATIONTYPE" & _
		        "        , ISU.SIVDATE" & _
		        "        , TO_NUMBER(TO_CHAR(ISU.SIVDATE,'YYYY')) AS NYEAR" & _
		        "        , TO_NUMBER(TO_CHAR(ISU.SIVDATE,'MM')) AS NMONTH" & _
		        "        , nvl(SUM(ID.QUANTITY),0) AS QUANTITY" & _
				" FROM ISSUE ISU" & _
		        "      , ISSUEDETAILS ID" & _
		        " WHERE (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA)" & _
		        "        AND to_number(to_Char(ISU.SIVDATE,'mm')) = " & Month(Session("v_Stores")) & _
				"        and to_number(to_Char(ISU.SIVDATE,'yyyy')) = " & Year(Session("v_Stores")) & _
				"        and ISU.fcBA = '" & Session("fcBA") & "'" & _
				" GROUP BY id.LineNoKey" & _
		        "          , isu.SIVCODE" & _
		        "          , ISU.STORECODE" & _
		        "          , ID.STOCKCODE" & _
		        "          , ID.JOBCODE" & _
		        "          , LOCATIONCODE" & _
		        "          , LOCATIONTYPE" & _
		        "          , ISU.SIVDATE" & _
		        "          , TO_CHAR(ISU.SIVDATE,'YYYY')" & _
		        "          , TO_CHAR(ISU.SIVDATE,'MM')"
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
				q="SELECT nvl(sum(AMTPERUNIT),0) as AMTPERUNIT  FROM STOCKLEDGER WHERE STORECODE='"  &  rsIssue("STORECODE") & "' AND ITEMCODE='"  &  rsIssue("ITEMCODE") & "' AND YEAR="  &  _
				  rsIssue("NYEAR") & " AND MONTH="  &  rsIssue("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
				'response.Write(q)		
				'response.Flush()
				Set rsCARII = Server.CreateObject("ADODB.Recordset")
				rsCARII.CursorType = 3
				rsCARII.CursorLocation = 3
				rsCARII.Open q, connstr	 
				nAmountIssue=0
				if Not rsCARII.eof Then
					nAmountIssue = Round(cdbl(rsCARII("AMTPERUNIT"))*cdbl(rsIssue("QUANTITY")), 2)
					q	=	" Update ISSUEDETAILS" & _
							" set AMOUNT = " & nAmountIssue & _
							"     , LASTUPDATE = sysdate" & _
							"     , LASTTIME = to_char(SYSDATE,'HH24:MI')" & _
							" where LINENOKEY = " & rsIssue("LINENOKEY") & "" & _
							"       and SIVCODE = '" & rsIssue("SIVCODE") & "'" & _
							"       and STOCKCODE = '"  &  rsIssue("ITEMCODE") & "'" & _
							"       AND LOCATIONTYPE = '"  &  rsIssue("LOCATIONTYPE") & "'" & _
							"       AND JOBCODE = '"& rsIssue("JOBCODE") & "'" & _
							"       AND LOCATIONCODE = '"&  rsISsue("LOCATIONCODE") & "'" & _
							"       and fcba = '" & Session("fcBA") & "'" & _
							"       and SIVCODE in (Select SIVCODE FROM ISSUE WHERE SIVDATE = to_date('" & rsIssue("SIVDATE") & "','mm/dd/yyyy')" & _
							"           and SIVCODE = '" & rsIssue("SIVCODE") & "' and STORECODE = '" & rsIssue("STORECODE") & "')"
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
		oQuery = 	" select storecode" & _
					"        , itemcode" & _
					"        , year" & _
					"        , month" & _
					"        , nvl(sum(openingbalanceqty), 0) openingbalanceqty" & _
					"        , nvl(sum(qtyreceived), 0) qtyreceived" & _
					"        , nvl(sum(qtywriteoff), 0) qtywriteoff" & _
					"        , nvl(sum(qtyissued)*-1, 0) qtyissued" & _
					"        , nvl(sum(openingbalanceqty), 0) + nvl(sum(qtyreceivedtask), 0)  + nvl(sum(qtyreceived), 0) + nvl(sum(qtywriteoff), 0) + nvl(sum(qtytransporttask)*-1, 0) + nvl(sum(qtyissued)*-1, 0) qtysaldo" & _
					"        , nvl(sum(amtissued), 0) amtissue" & _
					"        , nvl(sum(openingbalanceamt), 0) + nvl(sum(amtreceivedtask), 0) + nvl(sum(amtreceived), 0) + nvl(sum(amtwriteoff), 0) + nvl(sum(amttransporttask)*-1, 0) + nvl(sum(amtissued)*-1, 0) amtsaldo" & _
					" from stockledger" & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and year = " & Year(Session("v_Stores")) & "" & _
					"       and month = " & Month(Session("v_Stores")) & "" & _
					" group by storecode" & _
					"          , itemcode" & _
					"          , year" & _
					"          , month" & _
					" "
					
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open oQuery, connstr
		
		If Not oRS.EOF Then
			Do While Not oRS.EOF
				SaldoOnStockLedger = Cdbl(oRS("amtissue"))
				objectQuery =	" select nvl(max(linenokey), 0.000000001) linenokey, nvl(sum(amount), 0) amtissue" & _
								" from issue iss" & _
								"      , issuedetails isd" & _
								" where iss.fcba = isd.fcba" & _
								"       and iss.sivcode = isd.sivcode" & _
								"       and isd.fcba = '" & Session("fcBA") & "'" & _
								"       and to_number(to_char(sivdate, 'yyyy')) = " & Cdbl(oRS("year")) & "" & _
								"       and to_number(to_char(sivdate, 'mm')) = " & Cdbl(oRS("month")) & "" & _
								"       and stockcode = '" & oRS("itemcode") & "'" & _
								"       and storecode = '" & oRS("storecode") & "'"
								
				Set objectRS = Server.CreateObject("ADODB.Recordset")
				objectRS.Open objectQuery, connstr
				
				'response.write objectQuery
				
				If Not objectRS.EOF Then
					tempLineNokey = Cdbl(objectRS("linenokey"))
					SaldoOnIssue = Cdbl(objectRS("amtissue"))
				Else
					tempLineNokey = 0.000000001
					SaldoOnIssue = 0
				End If
				objectRS.Close
				set objectRS = nothing
				
				If (SaldoOnStockLedger - SaldoOnIssue) <> 0 Then
					objectQuery =	" select id.linenokey" & _
									"        , isu.sivcode" & _
									"        , isu.storecode as storecode" & _
									"        , id.stockcode as itemcode" & _
									"        , id.jobcode" & _
									"        , locationcode" & _
									"        , locationtype" & _
									"        , isu.sivdate" & _
									"        , to_number(to_char(isu.sivdate,'yyyy')) as nyear" & _
									"        , to_number(to_char(isu.sivdate,'mm')) as nmonth" & _
									"        , nvl(sum(id.quantity),0) as quantity" & _
									" from issue isu" & _
									"      , issuedetails id" & _
									" where (isu.sivcode = id.sivcode and isu.fcba = id.fcba)" & _
									"        and to_number(to_char(isu.sivdate,'mm')) = " & Month(Session("v_Stores")) & _
									"        and to_number(to_char(isu.sivdate,'yyyy')) = " & Year(Session("v_Stores")) & _
									"        and isu.fcba = '" & Session("fcBA") & "'" & _
									"        and id.linenokey = " & tempLineNokey & "" & _
									" group by id.linenokey" & _
									"          , isu.sivcode" & _
									"          , isu.storecode" & _
									"          , id.stockcode" & _
									"          , id.jobcode" & _
									"          , locationcode" & _
									"          , locationtype" & _
									"          , isu.sivdate" & _
									"          , to_char(isu.sivdate,'yyyy')" & _
									"          , to_char(isu.sivdate,'mm')"
									
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectQuery, connstr
					
					If Not objectRS.EOF Then
						Do While Not objectRS.EOF						
							q	=	" update issuedetails" & _
									" set amount = amount + " & SaldoOnStockLedger - SaldoOnIssue & _
									"     , lastupdate = sysdate" & _
									"     , lasttime = to_char(sysdate,'hh24:mi')" & _
									" where linenokey = " & objectrs("linenokey") & "" & _
									"       and sivcode = '" & objectrs("sivcode") & "'" & _
									"       and stockcode = '"  &  objectrs("itemcode") & "'" & _
									"       and locationtype = '"  &  objectrs("locationtype") & "'" & _
									"       and jobcode = '"& objectrs("jobcode") & "'" & _
									"       and locationcode = '"&  objectrs("locationcode") & "'" & _
									"       and fcba = '" & session("fcBA") & "'" & _
									"       and sivcode in (select sivcode from issue where sivdate = to_date('" & objectrs("sivdate") & "','mm/dd/yyyy')" & _
									"           and sivcode = '" & objectrs("sivcode") & "' and storecode = '" & objectrs("storecode") & "')"

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
		
		
		Response.Write "<br><font color=Blue>" & "Update Amount at Issue Details is completed..." & "</font>"
		
		'Insert Saldo Awal Bulan Depan
			Dim rsSaldoawal,QtySaldoawal,Bulan,Tahun,RpSaldoawal,rsCARISaldo
			q= " SELECT STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, QTYRECEIVED, AMTRECEIVED, QTYRETURNED, AMTRETURNED, " & _
			   " QTYWRITEOFF, AMTWRITEOFF, QTYTRANSFERED, AMTTRANSFERED, QTYISSUED, AMTISSUED,QTYRECEIVEDTASK, AMTRECEIVEDTASK, QTYTRANSPORTTASK, AMTTRANSPORTTASK   " & _ 
			   " FROM STOCKLEDGER Where MONTH=" & Month(Session("v_Stores")) &  _
			   " AND YEAR=" & Year(Session("v_Stores")) & " AND FCBA='" & Session("fcBA") & "'"
				Set rsSaldoawal = Server.CreateObject("ADODB.Recordset")
				rsSaldoawal.CursorType = 3
				rsSaldoawal.CursorLocation = 3
				rsSaldoawal.Open q, connstr

			Do While Not rsSaldoawal.eof 
				Bulan=Month(Session("v_Stores"))
				Tahun =Year(Session("v_Stores"))

				if bulan = 12 then
					tahun=tahun + 1
					bulan= 1
				else
					tahun=tahun
					bulan= bulan + 1
				end if

				' Remarked by NAF to handle the Amount Write OFF as it is., 
				' RpSaldoawal=cdbl(rsSaldoawal("OPENINGBALANCEAMT"))+cdbl(rsSaldoawal("AMTRECEIVED"))-cdbl(rsSaldoawal("AMTISSUED"))-cdbl(rsSaldoawal("AMTWRITEOFF"))
				' QtySaldoawal=cdbl(rsSaldoawal("OPENINGBALANCEQTY"))+cdbl(rsSaldoawal("QTYRECEIVED"))-cdbl(rsSaldoawal("QTYISSUED"))-cdbl(rsSaldoawal("QTYWRITEOFF"))
				' if user input negative then it will be negative, and if user input positif then it will be positif
				' by NAF, 23 January 2014
				

				RpSaldoawal= (cdbl(rsSaldoawal("OPENINGBALANCEAMT"))+cdbl(rsSaldoawal("AMTRECEIVED"))+cdbl(rsSaldoawal("AMTRECEIVEDTASK"))+cdbl(rsSaldoawal("AMTWRITEOFF"))) - (cdbl(rsSaldoawal("AMTISSUED"))+cdbl(rsSaldoawal("AMTTRANSPORTTASK")))
				QtySaldoawal=(cdbl(rsSaldoawal("OPENINGBALANCEQTY"))+cdbl(rsSaldoawal("QTYRECEIVED"))+ cdbl(rsSaldoawal("QTYRECEIVEDTASK"))+cdbl(rsSaldoawal("QTYWRITEOFF")))- (cdbl(rsSaldoawal("QTYISSUED"))+cdbl(rsSaldoawal("QTYTRANSPORTTASK")))


				q="SELECT STORECODE FROM STOCKLEDGER WHERE STORECODE='"  &  rsSaldoawal("STORECODE") & "' AND ITEMCODE='"  &  rsSaldoawal("ITEMCODE") & _
				  "' AND YEAR="  &  tahun & " AND MONTH="  &  bulan & " AND FCBA= '" & Session("fcBA") & "'"
				Set rsCARISaldo = Server.CreateObject("ADODB.Recordset")
				rsCARISaldo.CursorType = 3
				rsCARISaldo.CursorLocation = 3
				rsCARISaldo.Open q, connstr	 
				if rsCARISaldo.eof Then			
					q="Insert into STOCKLEDGER (STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					"values ('"  &  rsSaldoawal("STORECODE") & "','"  &  rsSaldoawal("ITEMCODE") & "',"  & tahun & _
					","  &  bulan & ","  &  QtySaldoawal & ","  &  RpSaldoawal & _
					",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))"
					connStr.execute(q)
				else
				   q="Update STOCKLEDGER set OPENINGBALANCEQTY =" & QtySaldoawal & ",OPENINGBALANCEAMT=" & RpSaldoawal & "" & _
					 ", LASTUPDATE = sysdate, LASTTIME = to_char(SYSDATE,'HH24:MI') " & _					   
					" where  STORECODE='"  &  rsSaldoawal("STORECODE") & "' AND ITEMCODE='"  &  rsSaldoawal("ITEMCODE") & "' AND YEAR="  &  _
					tahun & " AND MONTH="  &  bulan & " AND FCBA= '" & Session("fcBA") & "'"
					connStr.execute(q)
					connStr.execute("commit")
				End if
				rsSaldoawal.MOVENEXT
			loop
			' rsCARISaldo.close
			' set rsCARISaldo	=nothing
			rsSaldoawal.Close
			set rsSaldoawal=nothing
		
			'----- Tranfer to Activity Costing kREDIT Control Job 
			'edited by ukie on 24 october 2012 fro loc.type= 'CB' and loccode ='BS'


			'----- Tranfer to Activity Costing
			Response.Flush()
			
			
			tempProcessCode = "STORE_CLOSING"
			tempProcessNameCode = "p_Store_Closing.asp"
			tempRefCodeReceive = "WRITEOFFDETAILS"

			' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
			' by NAF, 12 May 2013
			' First, take the Original Table of RECEIVENOTE.
			qReference = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
						  " from process_tableref_tab " & _
						  " where " & _
						  " processcode = '" & tempProcessCode & "' and " & _
						  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
						  " fccode =  '" & tempRefCodeReceive & "'"

			Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
			rsReferenceReceive.CursorType = 3
			rsReferenceReceive.CursorLocation = 3
			rsReferenceReceive.Open qReference, connstr

			doneReceive = 1
			if not rsReferenceReceive.eof then
				txtRefNumberKeys_Receive = rsReferenceReceive("REF_NUMBERKEYS")
				txtRefStringSeparator_Receive = rsReferenceReceive("REF_STRINGSEPARATOR")
				txtRefStringKeys_Receive = rsReferenceReceive("REF_STRINGKEYS")
				doneReceive = 1
			else
				doneReceive = 0
			end if
			rsReferenceReceive.close
			set rsReferenceReceive = Nothing

		
			' Start to create the Journal based on New Concept for Write OFF
			' The System will create POSITIF as a POSITIF and
			' The System will create NEGATIF as a NEGATIF in Cost Activity
			' Means :
			' POSITIF ==> System will INCREASE Inventory
			' 		Db. Control_Job from Inventory/StockGroup
			'       Cr. Job from Write Off
			' NEGATIF ==> System will DECREASE Inventory
			'       Db. Job from Write Off
			'       Cr. Control Job from Inventory/StockGroup
			' By NAF, 23 January 2014
			' If you all have any question, please let me know, Ph. 081949939865
			
			queryWriteOffDetail = 	" SELECT isu.WODATE wodate, " & _
									" id.JOBCODE JOBCODE, " & _
									" id.stockcode stockcode, " & _
									" stockgroup_api.get_controljob(STOCK_API.GET_STOCKGROUP(id.stockcode, id.fcba),id.fcba) ControlJob, " & _
									" 'ST' source, " & _
									" 'ST' sourcetype, " & _
									" ISU.STORECODE sourcecode, " & _
									" '' sourcecode_status, " & _
									" id.locationtype targettype, " & _
									" id.locationcode targetcode, " & _
									" '' targetcode_status, " & _
									" id.WRITEOFFCODE REFERENCENO, " & _
									" SUM(id.QUANTITY) QUANTITY, " & _
									" 0 output, " & _
									" SUM(id.AMOUNT) AMOUNT, " & _
									" id.linenokey, " & _
									" id.FCBA , " & _
									"  to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date, " & _
									" sysdate process_date, " & _
									" sysdate lastupdate, " & _
									" to_char(sysdate, 'HH24:MI') lasttime " & _
								    " FROM WRITEOFF ISU,WRITEOFFDETAILS ID  " & _
								    " WHERE (ISU.WRITEOFFCODE=ID.WRITEOFFCODE AND ISU.FCBA=ID.FCBA) " & _ 
								    " and to_number(to_Char(ISU.WODATE,'mm'))=" & Month(Session("v_Stores")) &  _
								    " and to_number(to_Char(ISU.WODATE,'yyyy'))= " & Year(Session("v_Stores")) & " and ISU.fcBA='" & Session("fcBA") & "' " & _ 
								    " GROUP BY isu.WODATE, id.JOBCODE, id.stockcode, stockgroup_api.get_controljob(STOCK_API.GET_STOCKGROUP(id.stockcode, id.fcba),id.fcba), " & _
								    " 'ST', 'ST', ISU.STORECODE, '', id.LOCATIONTYPE, id.LOCATIONCODE, '', id.WRITEOFFCODE, id.linenokey, id.fcBA " 
			
			Set rsReceive = Server.CreateObject("ADODB.Recordset")
			rsReceive.CursorType = 3
			rsReceive.CursorLocation = 3
			'response.write queryWriteOffDetail
			rsReceive.Open queryWriteOffDetail, connstr
			
			if not rsReceive.eof then
				Do While Not rsReceive.eof
					' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
					txtRefStringValues_Receive = rsReceive("REFERENCENO") & txtRefStringSeparator_Receive & _
											   Session("fcBA") & txtRefStringSeparator_Receive & _
											   rsReceive("linenokey") 
											   

					if cdbl(rsReceive("AMOUNT")) >= 0 then 
					    ' This Write OFF will ADJUST POSITIF to Inventory / INCREASE the Inventory
						' 		Db. Control_Job from Inventory/StockGroup, CB, BS
						'       Cr. Job from Write Off, LocationType, LocationCode	
						
						temp_Db_Job = rsReceive("CONTROLJOB")
						temp_Db_TargetType = "CB"
						temp_Db_TargetCode = "BS"
						
						
						temp_Cr_Job = rsReceive("JOBCODE")
						temp_Cr_TargetType = rsReceive("targettype")
						temp_Cr_TargetCode = rsReceive("targetcode")
						
						'response.write " Db -> " & temp_Db_Job  & "   , Cr ->  " & temp_Cr_Job
						
					else
					    ' This Write OFF will ADJUST NEGATIF to Inventory / DECREASE the Inventory
						'       Dr. Job from Write Off, LocationType, LocationCode	
						' 		Cr. Control_Job from Inventory/StockGroup, CB, BS

						temp_Db_Job = rsReceive("JOBCODE")
						temp_Db_TargetType = rsReceive("targettype")
						temp_Db_TargetCode = rsReceive("targetcode")

						temp_Cr_Job = rsReceive("CONTROLJOB")
						temp_Cr_TargetType = "CB"
						temp_Cr_TargetCode = "BS"						
					end if
					
					' To Pick the TITLE REFERENCE, by NAF, 5 Feb 2014
					queryReference =   " select Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST','" & temp_Db_TargetType & "','Dr') temp_Db_Title_Reference_01, " & _
									   " Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST','" & temp_Db_TargetType & "','Dr') temp_Db_Title_Reference_02, " & _
									   " Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST','" & temp_Db_TargetType & "','Dr') temp_Db_Title_Reference_03 from dual " 
					
					Set rsqueryReference = Server.CreateObject("ADODB.Recordset")
					rsqueryReference.CursorType = 3
					rsqueryReference.CursorLocation = 3
					rsqueryReference.Open queryReference, connstr

					if not rsqueryReference.eof then
						temp_Db_Title_Reference_01 = rsqueryReference("temp_Db_Title_Reference_01")
						temp_Db_Auto_Reference_01 = temp_Db_TargetCode
						temp_Db_Title_Reference_02 = rsqueryReference("temp_Db_Title_Reference_02")
						temp_Db_Auto_Reference_02 = rsReceive("REFERENCENO")
						temp_Db_Title_Reference_03 = rsqueryReference("temp_Db_Title_Reference_03")
						temp_Db_Auto_Reference_03 = rsReceive("stockcode")
						temp_Db_Title_Reference_04 = ""
						temp_Db_Auto_Reference_04 = ""
						temp_Db_Title_Reference_05 = ""
						temp_Db_Auto_Reference_05 = ""
					else
						temp_Db_Title_Reference_01 = ""
						temp_Db_Auto_Reference_01 = ""
						temp_Db_Title_Reference_02 = ""
						temp_Db_Auto_Reference_02 = ""
						temp_Db_Title_Reference_03 = ""
						temp_Db_Auto_Reference_03 = ""
						temp_Db_Title_Reference_04 = ""
						temp_Db_Auto_Reference_04 = ""
						temp_Db_Title_Reference_05 = ""
						temp_Db_Auto_Reference_05 = ""						
					end if
					rsqueryReference.close
					set rsqueryReference = Nothing
					' End of To Pick the TITLE REFERENCE, by NAF, 5 Feb 2014
					
					' To Pick the TITLE REFERENCE, by NAF, 5 Feb 2014  ==> Cr
					queryReference = "select Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST','" & temp_Cr_TargetType & "','Cr') temp_Cr_Title_Reference_01, " & _
										   " Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST','" & temp_Cr_TargetType & "','Cr') temp_Cr_Title_Reference_02, " & _
										   " Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST','" & temp_Cr_TargetType & "','Cr') temp_Cr_Title_Reference_03 from dual " 
					
					Set rsqueryReference = Server.CreateObject("ADODB.Recordset")
					rsqueryReference.CursorType = 3
					rsqueryReference.CursorLocation = 3
					rsqueryReference.Open queryReference, connstr

					if not rsqueryReference.eof then
						temp_Cr_Title_Reference_01 = rsqueryReference("temp_Cr_Title_Reference_01")
						temp_Cr_Auto_Reference_01 = temp_Cr_TargetCode
						temp_Cr_Title_Reference_02 = rsqueryReference("temp_Cr_Title_Reference_02")
						temp_Cr_Auto_Reference_02 = rsReceive("REFERENCENO")
						temp_Cr_Title_Reference_03 = rsqueryReference("temp_Cr_Title_Reference_03")
						temp_Cr_Auto_Reference_03 = rsReceive("stockcode")
						temp_Cr_Title_Reference_04 = ""
						temp_Cr_Auto_Reference_04 = ""
						temp_Cr_Title_Reference_05 = ""
						temp_Cr_Auto_Reference_05 = ""
					else
						temp_Cr_Title_Reference_01 = ""
						temp_Cr_Auto_Reference_01 = ""
						temp_Cr_Title_Reference_02 = ""
						temp_Cr_Auto_Reference_02 = ""
						temp_Cr_Title_Reference_03 = ""
						temp_Cr_Auto_Reference_03 = ""
						temp_Cr_Title_Reference_04 = ""
						temp_Cr_Auto_Reference_04 = ""
						temp_Cr_Title_Reference_05 = ""
						temp_Cr_Auto_Reference_05 = ""						
					end if
					rsqueryReference.close
					set rsqueryReference = Nothing
					' End of To Pick the TITLE REFERENCE, by NAF, 5 Feb 2014
					
					 
					' Insert the DEBET Record, 
					maxLineNo = maxLineNo + 1		
					qCost_ActivityDB = "Insert into Cost_Activity " & _
								 "(LINENO, " & _
									" TDATE, " & _ 
									" JOBCODE,  " & _
									" SOURCE, " & _
									" SOURCETYPE, " & _
									" SOURCECODE, " & _
									" SOURCECODE_STATUS, " & _
									" TARGETTYPE, " & _
									" TARGETCODE, " & _
									" TARGETCODE_STATUS,  " & _
									" REFERENCENO, " & _
									" AMOUNT, " & _
									" QUANTITY, " & _
									" OUTPUT," & _
									" FCBA, " & _
									" closing_date, " & _
									" process_date, " & _
									" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _		
                                    " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _									
									" LASTUPDATE, " & _
									" LASTTIME ) " & _
								 " values (" & maxLineNo & _
								 ", to_date('" & rsReceive("WODATE") & "','mm/dd/yyyy') ,'" & _
								 temp_Db_Job & "','" & _
								 rsReceive("SOURCE") & "','" & _
								 rsReceive("SOURCETYPE") & "','" & _
								 rsReceive("SOURCECODE") & "','" & _
								 rsReceive("SOURCECODE_STATUS") & "','" & _
								 temp_Db_TargetType & "','" & _
								 temp_Db_TargetCode & "','" & _
								 rsReceive("TARGETCODE_STATUS") & "','" & _
								 rsReceive("REFERENCENO") & "'," & _
								 "abs(" & cdbl(rsReceive("AMOUNT")) & ")," & _
								 cdbl(rsReceive("QUANTITY")) & "," & _
								 cdbl(rsReceive("OUTPUT")) & ",'" & _
								 rsReceive("FCBA") & "'," & _
								 "to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') ," & _
								 " sysdate,'" & _
								 temp_Db_Title_Reference_01 & "','" & _
								 temp_Db_Auto_Reference_01  & "','" & _
								 temp_Db_Title_Reference_02  & "','" & _
								 temp_Db_Auto_Reference_02  & "','" & _
								 temp_Db_Title_Reference_03  & "','" & _
								 temp_Db_Auto_Reference_03  & "','" & _
								 temp_Db_Title_Reference_04  & "','" & _
								 temp_Db_Auto_Reference_04  & "','" & _
								 temp_Db_Title_Reference_05  & "','" & _
								 temp_Db_Auto_Reference_05  & "','Dr'," & _
								 Month(Session("v_Stores")) & ", " & _
								 Year(Session("v_Stores")) & ", " & _
								 " sysdate," & _
								 " to_char(sysdate, 'HH24:MI'))" 

					'response.write qCost_ActivityDB
					response.Flush()
					connstr.execute(qCost_ActivityDB)	
					
					
					
					'* Remarks By Iwa * 
					'qABCDetail = " SELECT isu.WODATE wodate, " & _
					'				" stockgroup.CONTROLJOB controljob, " & _
					'				" 'ST' source, " & _
					'				" 'ST' sourcetype, " & _
					'				" ISU.STORECODE sourcecode, " & _
					'				" '' sourcecode_status, " & _
					'				" id.locationtype targettype, " & _
					'				" id.locationcode targetcode, " & _
					'				" '' targetcode_status, " & _
					'				" id.WRITEOFFCODE REFERENCENO, " & _
					'				" nvl(id.QUANTITY,0) QUANTITY, " & _
					'				" 0 output, " & _
					'				" nvl(id.AMOUNT,0) AMOUNT, " & _
					'				" id.linenokey, " & _
					'				" id.FCBA , " & _
					'				"  to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date, " & _
					'				" sysdate process_date, " & _
					'				" sysdate lastupdate, " & _
					'				" to_char(sysdate, 'HH24:MI') lasttime " & _
					'			" FROM WRITEOFF ISU,WRITEOFFDETAILS ID,Stock,STOCKGROUP  " & _
					'			" WHERE (ISU.WRITEOFFCODE=ID.WRITEOFFCODE AND ISU.FCBA=ID.FCBA) " & _ 
					'			" and (id.STOCKCODE =stock.FCCODE and id.fcba = stock.fcba)  " & _
					'			" and (stock.STOCKGROUPCODE =stockgroup.FCCODE and stock.fcba = stockgroup.fcba)  " & _
					'			" and to_number(to_Char(ISU.WODATE,'mm'))=" & Month(Session("v_Stores")) &  _
					'			" and to_number(to_Char(ISU.WODATE,'yyyy'))= " & Year(Session("v_Stores")) & " and ISU.fcBA='" & Session("fcBA") & "' " & _ 
					'			" and id.WRITEOFFCODE = '" & rsReceive("REFERENCENO") & "'" & _
					'			" and id.linenokey = '" & rsReceive("linenokey") & "'"
					
					'Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					'rsReceiveDetail.CursorType = 3
					'rsReceiveDetail.CursorLocation = 3
					'rsReceiveDetail.Open qABCDetail, connstr
					'if not rsReceiveDetail.eof then
					'	Do While Not rsReceiveDetail.eof
					'	
					'		maxLineNoDetail = maxLineNoDetail + 1
					'		qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
					'						"( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
					'						"  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
					'						" QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
					'						maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
					'						txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
					'						cdbl(rsReceiveDetail("QUANTITY")) & ",0,abs("& cdbl(rsReceiveDetail("AMOUNT")) &"),'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
					'						Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
					'		
					'		response.write qCost_ActivityDBHistory & "</br>"
					'		connstr.execute(qCost_ActivityDBHistory)		
					'	
					'		rsReceiveDetail.MoveNext
					'	loop
					'end if
					'rsReceiveDetail.Close
					'set rsReceiveDetail = Nothing			
					
					'* Remarks By Iwa * 
					
					' Insert the CREDIT Record
					
					maxLineNo = maxLineNo + 1		
					qCost_ActivityCr = "Insert into Cost_Activity " & _
									"(LINENO, " & _
									" TDATE, " & _ 
									" JOBCODE,  " & _
									" SOURCE, " & _
									" SOURCETYPE, " & _
									" SOURCECODE, " & _
									" SOURCECODE_STATUS, " & _
									" TARGETTYPE, " & _
									" TARGETCODE, " & _
									" TARGETCODE_STATUS,  " & _
									" REFERENCENO, " & _
									" AMOUNT, " & _
									" QUANTITY, " & _
									" OUTPUT," & _
									" FCBA, " & _
									" closing_date, " & _
									" process_date, " & _
									" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _		
                                    " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _										
									" LASTUPDATE, " & _
									" LASTTIME ) " & _
								 " values (" & maxLineNo & _
								 ", to_date('" & rsReceive("WODATE") & "','mm/dd/yyyy') ,'" & _
								 temp_Cr_Job & "','" & _
								 rsReceive("SOURCE") & "','" & _
								 rsReceive("SOURCETYPE") & "','" & _
								 rsReceive("SOURCECODE") & "','" & _
								 rsReceive("SOURCECODE_STATUS") & "','" & _
								 temp_Cr_TargetType & "','" & _
								 temp_Cr_TargetCode & "','" & _
								 rsReceive("TARGETCODE_STATUS") & "','" & _
								 rsReceive("REFERENCENO") & "'," & _
								 "-1 * abs(" & cdbl(rsReceive("AMOUNT")) & ")," & _
								 cdbl(rsReceive("QUANTITY")) & "," & _
								 cdbl(rsReceive("OUTPUT")) & ",'" & _
								 rsReceive("FCBA") & "'," & _
								 "to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') ," & _
								 " sysdate, '" & _
								 temp_Cr_Title_Reference_01 & "','" & _
								 temp_Cr_Auto_Reference_01  & "','" & _
								 temp_Cr_Title_Reference_02  & "','" & _
								 temp_Cr_Auto_Reference_02  & "','" & _
								 temp_Cr_Title_Reference_03  & "','" & _
								 temp_Cr_Auto_Reference_03  & "','" & _
								 temp_Cr_Title_Reference_04  & "','" & _
								 temp_Cr_Auto_Reference_04  & "','" & _
								 temp_Cr_Title_Reference_05  & "','" & _
								 temp_Cr_Auto_Reference_05  & "','Cr'," & _
								 Month(Session("v_Stores")) & ",  " & _
								 Year(Session("v_Stores")) & ",  " & _
								 " sysdate," & _
								 " to_char(sysdate, 'HH24:MI'))" 

                    'response.write  qCost_ActivityCr
					response.Flush()
					connstr.execute(qCost_ActivityCr)	
					
					
					'* Remarks By Iwa * 
					
					'qABCDetail = " SELECT isu.WODATE wodate, " & _
					'				" stockgroup.CONTROLJOB controljob, " & _
					'				" 'ST' source, " & _
					'				" 'ST' sourcetype, " & _
					'				" ISU.STORECODE sourcecode, " & _
					'				" '' sourcecode_status, " & _
					'				" id.locationtype targettype, " & _
					'				" id.locationcode targetcode, " & _
					'				" '' targetcode_status, " & _
					'				" id.WRITEOFFCODE REFERENCENO, " & _
					'				" nvl(id.QUANTITY,0) QUANTITY, " & _
					'				" 0 output, " & _
					'				" nvl(id.AMOUNT,0) AMOUNT, " & _
					'				" id.linenokey, " & _
					'				" id.FCBA , " & _
					'				"  to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date, " & _
					'				" sysdate process_date, " & _
					'				" sysdate lastupdate, " & _
					'				" to_char(sysdate, 'HH24:MI') lasttime " & _
					'				" FROM WRITEOFF ISU,WRITEOFFDETAILS ID,Stock,STOCKGROUP  " & _
					'				" WHERE (ISU.WRITEOFFCODE=ID.WRITEOFFCODE AND ISU.FCBA=ID.FCBA) " & _ 
					'				" and (id.STOCKCODE =stock.FCCODE and id.fcba = stock.fcba)  " & _
					'				" and (stock.STOCKGROUPCODE =stockgroup.FCCODE and stock.fcba = stockgroup.fcba)  " & _
					'				" and to_number(to_Char(ISU.WODATE,'mm'))=" & Month(Session("v_Stores")) &  _
					'				" and to_number(to_Char(ISU.WODATE,'yyyy'))= " & Year(Session("v_Stores")) & " and ISU.fcBA='" & Session("fcBA") & "' " & _ 
					'				" and id.WRITEOFFCODE = '" & rsReceive("REFERENCENO") & "'" & _
					'				" and id.linenokey = '" & rsReceive("linenokey") & "'"
					'
					'Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					'rsReceiveDetail.CursorType = 3
					'rsReceiveDetail.CursorLocation = 3
					'rsReceiveDetail.Open qABCDetail, connstr
					'if not rsReceiveDetail.eof then
					'	Do While Not rsReceiveDetail.eof
					'	
					'		maxLineNoDetail = maxLineNoDetail + 1
					'		qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
					'						"( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
					'						"  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
					'						" QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
					'						maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
					'						txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
					'						cdbl(rsReceiveDetail("QUANTITY")) & ",0,-1 * abs("& cdbl(rsReceiveDetail("AMOUNT")) &"),'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
					'						Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
					'		
					'		response.write qCost_ActivityDBHistory & "</br>"
					'		connstr.execute(qCost_ActivityDBHistory)		
					'	
					'		rsReceiveDetail.MoveNext
					'	loop
					'end if
					'rsReceiveDetail.Close
					'set rsReceiveDetail = Nothing		
					
					'* Remarks By Iwa * 
					
					rsReceive.MoveNext
				loop
			end if
			rsReceive.Close
			set rsReceive = Nothing	
			
			'----- Tranfer to Activity Costing Completed 
			Response.Write "<br><font color=Blue>" & "(Write Off) Transfer to Activity Based Costing is completed..." & "</font>"
			Response.Flush()
					
					
		'  Add By Akbar 30 sept 2016  Process Cost Activity from Transport Task
		
		tempRefCodeReceive = "TRANSPORTTASK"

	
		qReference =  "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
					  " from PROCESS_TABLEREF_TAB " & _
					  " where " & _
					  " processcode = '" & tempProcessCode & "' and " & _
					  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
					  " fccode =  '" & tempRefCodeReceive & "'"

		Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
		rsReferenceReceive.CursorType = 3
		rsReferenceReceive.CursorLocation = 3
		rsReferenceReceive.Open qReference, connstr

		doneReceive = 1
		if not rsReferenceReceive.eof then
			txtRefNumberKeys_Receive = rsReferenceReceive("REF_NUMBERKEYS")
			txtRefStringSeparator_Receive = rsReferenceReceive("REF_STRINGSEPARATOR")
			txtRefStringKeys_Receive = rsReferenceReceive("REF_STRINGKEYS")
			doneReceive = 1
		else
			doneReceive = 0
		end if
		rsReferenceReceive.close
		set rsReferenceReceive = Nothing	
		
		
		strCursorSivCode =  " Select ORDERNO " & _
							" from TRANSPORTTASK_TAB " & _
							" where " & _
							" to_number(to_char(orderdate , 'mm')) = " & month (session ("v_stores")) & _
							" and to_number(to_char(orderdate , 'yyyy')) = " & year (session ("v_stores")) & _
							" and fcba = '" & Session("fcBA") & "' and ROWSTATEID ='Approved'" 
							
		Set rsCursorSivCode = Server.CreateObject("ADODB.Recordset")
		rsCursorSivCode.CursorType = 3
		rsCursorSivCode.CursorLocation = 3
		rsCursorSivCode.Open strCursorSivCode, connstr		

		if not rsCursorSivCode.eof then
			Do While Not rsCursorSivCode.eof
			
				' NAF : Jangan dirubah format string-nya menjadi lurus atau via excel, biarkan seperti apa adanya atau 
				'       Jangan menggunakan TRIM untuk menghilangkan SPASI-nya "
			

				
				strCursorCredit =   " SELECT  ID.FCBA," & _
									"		  ID.ORDERNO REFERENCENO," & _
									"		  ID.STOCKCODE, " & _ 
									"		  TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) TDATE , " & _
									"		  STOCKGROUP_API.GET_CONTROLJOB(STOCK_API.GET_STOCKGROUP(ID.STOCKCODE,ID.FCBA),ID.FCBA) JOBCODE," & _
									"		  ID.JOBCODE JOBCODEORIGINAL," & _
									"		  ID.STORECODE SOURCECODE, " & _
									"		  ID.LINENO LINENOKEY, " & _
									"		  ID.LOCATIONTYPE, " & _
									"		  ID.LOCATIONCODE, " & _
									"		  'ST' SOURCE, " & _
									"		  'ST' SOURCETYPE, " & _
									"		  '' SOURCECODE_STATUS, " & _
									"		  'CB' TARGETTYPE, " & _
									"		  'BS' TARGETCODE, " & _
									"		  ''   TARGETCODE_STATUS, " & _
									"		  SUM(ID.QUANTITYTRANSFERED) QUANTITY, " & _
									"		  0 OUTPUT, " & _
									"		  ROUND( (SUM (ID.AMOUNT) * -1) ,2) AMOUNT, " & _
									"		  TO_DATE('" & TEMP_DATEPERIODCONTROL & "','MM/DD/YYYY') CLOSING_DATE, " & _
									"		  SYSDATE PROCESS_DATE " & _
									"  FROM TRANSPORTTASKDTL_TAB ID " & _
									"	INNER JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = ID.ORDERNO AND TT.FCBA = ID.FCBA)" & _
									"  WHERE  " & _
								    "  ID.ORDERNO = '" & rsCursorSivCode("ORDERNO") & "' " & _
									"  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'mm')) = " & Month(Session("v_Stores")) &  _
									"  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'yyyy')) =" & Year(Session("v_Stores"))  & _
									"  AND ID.fcBA = '" & Session("fcBA") & "'" & _
									"  AND TT.ROWSTATEID = 'Approved'" & _
								    "  GROUP BY  ID.FCBA," & _
									"	  ID.ORDERNO," & _
									"	  ID.STOCKCODE, " & _ 
									"	  TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) ," & _ 
									"	  STOCKGROUP_API.GET_CONTROLJOB(STOCK_API.GET_STOCKGROUP(ID.STOCKCODE,ID.FCBA),ID.FCBA)," & _
									"	  ID.JOBCODE," & _
									"	  ID.STORECODE," & _ 
									"	  ID.LINENO, " & _
									"	  ID.LOCATIONTYPE," & _ 
									"	  ID.LOCATIONCODE "
										 
				
				Set rsCursorCredit = Server.CreateObject("ADODB.Recordset")
				rsCursorCredit.CursorType = 3
				rsCursorCredit.CursorLocation = 3
				rsCursorCredit.Open strCursorCredit, connstr	
				
				
				if not rsCursorCredit.eof then
				
					Do While Not rsCursorCredit.eof
				
						tempTotalAmountCredit = round(cdbl(rsCursorCredit("AMOUNT")),2)
						tempSIVCODE = rsCursorCredit("referenceno")
						tempStockCode = rsCursorCredit("STOCKCODE")

						txtRefStringValues_Receive = tempSIVCODE & txtRefStringSeparator_Receive & _
													 rsCursorCredit("linenokey") & txtRefStringSeparator_Receive & _
													 rsCursorCredit("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
													 rsCursorCredit("LOCATIONCODE") & txtRefStringSeparator_Receive & _
													 rsCursorCredit("jobcodeoriginal") & txtRefStringSeparator_Receive & _
													 rsCursorCredit("FCBA") 
													 
						
						' NAF : Jangan dirubah format string-nya menjadi lurus atau via excel, biarkan seperti apa adanya atau 
						'       Jangan menggunakan TRIM untuk menghilangkan SPASI-nya "		

						
						strCursorCount = 	" select count(*) totalRecord from ( " & _
											" SELECT TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO)," & _
											"	 id.JOBCODE, " & _
											"	 id.STORECODE, " & _
											"	 id.LOCATIONTYPE, " & _
											"	 id.LOCATIONCODE, " & _
											"	 id.ORDERNO, " & _
											"    id.LINENO, " & _
											"	 id.FCBA " & _
											" FROM TRANSPORTTASKDTL_TAB ID   " & _
											"	INNER JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = ID.ORDERNO AND TT.FCBA = ID.FCBA)" & _
										    " WHERE	  " & _
											"	 TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'mm')) = " & Month(Session("v_Stores")) &  _
											"	 AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'yyyy'))  =" & Year(Session("v_Stores"))  & _
											"	 AND ID.fcBA = '" & Session("fcBA") & "'" & _
											"	 AND ID.ORDERNO = '" & tempSIVCODE & "'  " & _
											"    AND id.lineno = '" & rsCursorCredit ("linenokey") & "' " & _
											"    AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
											"    AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
											"    AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") &  "' " & _
											"    AND ID.STOCKCODE = '" & tempStockCode & "' " & _
											"    AND TT.ROWSTATEID = 'Approved'" & _
										    " GROUP BY TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO),  " & _
											"	 id.JOBCODE,  " & _
											"	 id.STORECODE,  " & _
											"	 id.LOCATIONTYPE,  " & _
											"	 id.LOCATIONCODE,  " & _
											"	 id.ORDERNO,  " & _
											"    id.LINENO, " & _
											"	 id.fcBA  " & _
										    " ) "  

					
						Set rsCursorCount = Server.CreateObject("ADODB.Recordset")
						rsCursorCount.CursorType = 3
						rsCursorCount.CursorLocation = 3
						

						rsCursorCount.Open strCursorCount, connstr		

						totalRecord = cdbl(rsCursorCount("totalRecord"))
						totalRowInserted = 0
						tempTotalAmountDebit = 0

						
						Response.Flush()
						
						
						maxLineNo = maxLineNo + 1
						strInsertCredit = " INSERT INTO cost_activity ( LINENO , " & _
																  " TDATE, " & _
																  " JOBCODE, " & _
																  " SOURCE, " & _
																  " SOURCETYPE, " & _
																  " SOURCECODE, " & _
																  " SOURCECODE_STATUS, " & _
																  " TARGETTYPE, " & _
																  " TARGETCODE, " & _
																  " TARGETCODE_STATUS, " & _
																  " REFERENCENO, " & _
																  " QUANTITY, " & _
																  " OUTPUT, " & _
																  " AMOUNT, " & _
																  " FCBA, " & _
																  " CLOSING_DATE, " & _ 
																  " PROCESS_DATE , " & _
																  " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																  " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																  " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																  " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																  " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
																  " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _	
																  " LASTUPDATE,  " & _
																  " LASTTIME ) " & _
												" VALUES (" & _
												maxLineNo & "," & _
												" to_date('" & rsCursorCredit("tdate") & "','mm/dd/yyyy'),'" & _
												rsCursorCredit("JOBCODE") & "','" & _
												rsCursorCredit("SOURCE") & "','" & _
												rsCursorCredit("SOURCETYPE") & "','" & _
												rsCursorCredit("SOURCECODE") & "','" & _
												rsCursorCredit("SOURCECODE_STATUS") & "','" & _
												rsCursorCredit("TARGETTYPE") & "','" & _
												rsCursorCredit("TARGETCODE") & "','" & _
												"NAF||" & rsCursorCredit("TARGETCODE_STATUS") & "','" & _
												rsCursorCredit("REFERENCENO") & "'," & _
												rsCursorCredit("QUANTITY") & "," & _
												rsCursorCredit("OUTPUT") & "," & _
												tempTotalAmountCredit & ",'" & _
												rsCursorCredit("FCBA") & "'," & _
												" to_date('" & rsCursorCredit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
												" sysdate, " & _
												" Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & rsCursorCredit("TARGETCODE") & "'," & _
												" Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & tempSIVCODE & "'," & _
												" Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & tempStockCode & "'," & _
												" '',''," & _
												" '','','Cr'," & _
												Month(Session("v_Stores")) & ", " & _
												Year(Session("v_Stores")) & ", " & _
												" sysdate, " & _
												" to_char(sysdate,'HH24:MI')) " 
 						' response.write strInsertCredit
						response.Flush()		  
						connstr.Execute(strInsertCredit)		
						
						'* Remarks By Iwa * 
						'strCursorCreditDetail = " SELECT  ID.FCBA," & _
						'						"		  ID.ORDERNO REFERENCENO," & _
						'						"		  ID.STOCKCODE, " & _ 
						'						"		  TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) TDATE , " & _
						'						"		  STOCKGROUP_API.GET_CONTROLJOB(STOCK_API.GET_STOCKGROUP(ID.STOCKCODE,ID.FCBA),ID.FCBA) JOBCODE," & _
						'						"		  ID.JOBCODE JOBCODEORIGINAL," & _
						'						"		  ID.STORECODE SOURCECODE, " & _
						'						"		  ID.LINENO LINENOKEY, " & _
						'						"		  ID.LOCATIONTYPE, " & _
						'						"		  ID.LOCATIONCODE, " & _
						'						"		  'ST' SOURCE, " & _
						'						"		  'ST' SOURCETYPE, " & _
						'						"		  '' SOURCECODE_STATUS, " & _
						'						"		  'CB' TARGETTYPE, " & _
						'						"		  'BS' TARGETCODE, " & _
						'						"		  ''   TARGETCODE_STATUS, " & _
						'						"		  ID.QUANTITYTRANSFERED QUANTITY, " & _
						'						"		  0 OUTPUT, " & _
						'						"		  ROUND(ID.AMOUNT * -1) AMOUNT, " & _
						'						"		  TO_DATE('" & TEMP_DATEPERIODCONTROL & "','MM/DD/YYYY') CLOSING_DATE, " & _
						'						"		  SYSDATE PROCESS_DATE " & _
						'						"  FROM TRANSPORTTASKDTL_TAB ID " & _
						'						"	INNER JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = ID.ORDERNO AND TT.FCBA = ID.FCBA)" & _
						'						"  WHERE  " & _
						'						"  ID.ORDERNO = '" & rsCursorSivCode("ORDERNO") & "' " & _
						'						"  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'mm')) = " & Month(Session("v_Stores")) &  _
						'						"  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'yyyy')) =" & Year(Session("v_Stores"))  & _
						'						"  AND ID.fcBA = '" & Session("fcBA") & "'" & _
						'						"  AND TT.ROWSTATEID = 'Approved'" & _
						'						"  AND ID.ORDERNO = '" & tempSIVCODE & "'  " & _
						'						"  AND id.lineno = '" & rsCursorCredit ("linenokey") & "' " & _
						'						"  AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
						'						"  AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
						'						"  AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") &  "' " & _
						'						"  AND ID.STOCKCODE = '" & tempStockCode & "' " 
'

						'Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
						'rsReceiveDetail.CursorType = 3
						'rsReceiveDetail.CursorLocation = 3
						'rsReceiveDetail.Open strCursorCreditDetail, connstr
						'if not rsReceiveDetail.eof then
						'	Do While Not rsReceiveDetail.eof
						'	
						'		maxLineNoDetail = maxLineNoDetail + 1
						'		qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
						'						"( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
						'						"  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
						'						" QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
						'						maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
						'						txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
						'						cdbl(rsReceiveDetail("QUANTITY")) & ",0,"& cdbl(rsReceiveDetail("AMOUNT")) &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
						'						Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
						'		
						'		response.write qCost_ActivityDBHistory & "</br>"								
						'		connstr.execute(qCost_ActivityDBHistory)		
						'	
						'		rsReceiveDetail.MoveNext
						'	loop
						'end if
						'rsReceiveDetail.Close
						'set rsReceiveDetail = Nothing			
						
						'* Remarks By Iwa * 
						
						' NAF : Jangan dirubah format string-nya menjadi lurus atau via excel, biarkan seperti apa adanya atau 
						'       Jangan menggunakan TRIM untuk menghilangkan SPASI-nya "			
						     strCursorDebit = 	" SELECT  ID.FCBA," & _
												"		  ID.ORDERNO REFERENCENO," & _
												"		  ID.STOCKCODE, " & _ 
												"		  TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) TDATE , " & _
												"		  ID.JOBCODE," & _
												"		  ID.STORECODE SOURCECODE, " & _
												"		  ID.LINENO LINENOKEY, " & _
												"		  ID.LOCATIONTYPE, " & _
												"		  ID.LOCATIONCODE, " & _
												"		  'ST' SOURCE, " & _
												"		  'ST' SOURCETYPE, " & _
												"		  '' SOURCECODE_STATUS, " & _
												"		  'CB' TARGETTYPE, " & _
												"		  'BS' TARGETCODE, " & _
												"		  ''   TARGETCODE_STATUS, " & _
												"		  SUM(ID.QUANTITYTRANSFERED) QUANTITY, " & _
												"		  0 OUTPUT, " & _
												"		  ROUND( (SUM (ID.AMOUNT) * 1) ,2) AMOUNT, " & _
												"		  TO_DATE('" & TEMP_DATEPERIODCONTROL & "','MM/DD/YYYY') CLOSING_DATE, " & _
												"		  SYSDATE PROCESS_DATE " & _
												"  FROM TRANSPORTTASKDTL_TAB ID " & _
												"	INNER JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = ID.ORDERNO AND TT.FCBA = ID.FCBA)" & _
												"  WHERE  " & _
												"  ID.ORDERNO = '" & rsCursorSivCode("ORDERNO") & "' " & _
												"  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'mm')) = " & Month(Session("v_Stores")) &  _
												"  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'yyyy')) =" & Year(Session("v_Stores"))  & _
												"  AND ID.fcBA = '" & Session("fcBA") & "'" & _
												"  AND ID.ORDERNO = '" & tempSIVCODE & "'  " & _
												"  AND ID.STOCKCODE = '" & tempStockCode & "' " & _
												"  AND ID.LINENO = '" & rsCursorCredit ("linenokey") & "' " & _
												"  AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
												"  AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
												"  AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") & "' " & _	
												"  AND TT.ROWSTATEID = 'Approved'" & _												 
												"  GROUP BY  ID.FCBA," & _
												"	  		ID.ORDERNO," & _
												"	  		ID.STOCKCODE, " & _ 
												"	  		TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) ," & _ 
												"	  		ID.JOBCODE," & _
												"	  		ID.STORECODE," & _ 
												"	  		ID.LINENO, " & _
												"	  		ID.LOCATIONTYPE," & _ 
												"	  		ID.LOCATIONCODE "
												
						Set rsCursorDebit = Server.CreateObject("ADODB.Recordset")
						rsCursorDebit.CursorType = 3
						rsCursorDebit.CursorLocation = 3
						rsCursorDebit.Open strCursorDebit, connstr		
						
						tempRowNum = 1

						if not rsCursorDebit.eof then
							Do While Not rsCursorDebit.eof
								'		   If totalRowInserted = n  then 
								'			   tempTotalAmountDebit = tempTotalAmountDebit + dari Cursor 2 setelah di round
								'			   tempSelisih = tempTotalAmountCredit - tempTotalAmountDebit				   
								'			   Insert Debit ==> Amount setelah di Round + tempSelisih
								'           else
								'		       Insert Debit
								'			   tempTotalAmountDebit = tempTotalAmountDebit + dari Cursor 2 setelah di round
								'			   totalRowInserted = totalRowInserted + 1
								' 		   end if
								
								
								tempDebit = round(cdbl(rsCursorDebit("amount")),2)
								
								if totalRowInserted = totalRecord - 1 then 
									tempTotalAmountDebit = tempTotalAmountDebit + tempDebit
									
									if tempTotalAmountDebit < (-1 * tempTotalAmountCredit) then
										tempSelisih = (-1 * tempTotalAmountCredit) - tempTotalAmountDebit				   
										tempDebit = tempDebit + tempSelisih
									else
										tempSelisih = tempTotalAmountDebit - (-1 * tempTotalAmountCredit) 				   
										tempDebit = tempDebit - tempSelisih								
									end if
									Response.Flush()
									maxLineNo = maxLineNo + 1
									qInsertDebit = " INSERT INTO cost_activity ( LINENO ," & _
																			  " TDATE, " & _
																			  " JOBCODE, " & _
																			  " SOURCE, " & _
																			  " SOURCETYPE, " & _
																			  " SOURCECODE, " & _
																			  " SOURCECODE_STATUS, " & _
																			  " TARGETTYPE, " & _
																			  " TARGETCODE, " & _
																			  " TARGETCODE_STATUS, " & _
																			  " REFERENCENO, " & _
																			  " QUANTITY, " & _
																			  " OUTPUT, " & _
																			  " AMOUNT, " & _
																			  " FCBA, " & _
																			  " CLOSING_DATE, " & _ 
																			  " PROCESS_DATE, " & _		
																			  " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																			  " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																			  " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																			  " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																			  " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _		
																			  " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
																			  " LASTUPDATE,  " & _
																			  " LASTTIME ) " & _
															" VALUES (" & _
															maxLineNo & "," & _
															" to_date('" & rsCursorDebit("tdate") & "','mm/dd/yyyy'),'" & _
															rsCursorDebit("JOBCODE") & "','" & _
															rsCursorDebit("SOURCE") & "','" & _
															rsCursorDebit("SOURCETYPE") & "','" & _
															rsCursorDebit("SOURCECODE") & "','" & _
															rsCursorDebit("SOURCECODE_STATUS") & "','" & _
															rsCursorDebit("TARGETTYPE") & "','" & _
															rsCursorDebit("TARGETCODE") & "','" & _
															rsCursorDebit("TARGETCODE_STATUS") & "','" & _
															rsCursorDebit("REFERENCENO") & "'," & _
															rsCursorDebit("QUANTITY") & "," & _
															rsCursorDebit("OUTPUT") & "," & _
															tempDebit & ",'" & _
															rsCursorDebit("FCBA") & "'," & _
															" to_date('" & rsCursorDebit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
															" sysdate, " & _
															" Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & rsCursorDebit("TARGETCODE") & "'," & _
															" Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempSIVCODE & "'," & _
															" Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempStockCode & "'," & _
															" '',''," & _
															" '','','Dr'," & _
															Month(Session("v_Stores")) & ", " & _
															Year(Session("v_Stores")) & ", " & _
															" sysdate, " & _
															" to_char(sysdate,'HH24:MI')) "
									response.Flush()		  
									'response.write qInsertDebit
									connstr.Execute(qInsertDebit)			
									
									'* Remarks By Iwa * 
									'strCursorDebitDetail =  " SELECT  ID.FCBA," & _
									'						"		  ID.ORDERNO REFERENCENO," & _
									'						"		  ID.STOCKCODE, " & _ 
									'						"		  TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) TDATE , " & _
									'						"		  ID.JOBCODE," & _
									'						"		  ID.STORECODE SOURCECODE, " & _
									'						"		  ID.LINENO LINENOKEY, " & _
									'						"		  ID.LOCATIONTYPE, " & _
									'						"		  ID.LOCATIONCODE, " & _
									'						"		  'ST' SOURCE, " & _
									'						"		  'ST' SOURCETYPE, " & _
									'						"		  '' SOURCECODE_STATUS, " & _
									'						"		  'CB' TARGETTYPE, " & _
									'						"		  'BS' TARGETCODE, " & _
									'						"		  ''   TARGETCODE_STATUS, " & _
									'						"		  ID.QUANTITYTRANSFERED QUANTITY, " & _
									'						"		  0 OUTPUT, " & _
									'						"		  ROUND(ID.AMOUNT * 1) AMOUNT, " & _
									'						"		  TO_DATE('" & TEMP_DATEPERIODCONTROL & "','MM/DD/YYYY') CLOSING_DATE, " & _
									'						"		  SYSDATE PROCESS_DATE " & _
									'						"  FROM TRANSPORTTASKDTL_TAB ID " & _
									'						"	INNER JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = ID.ORDERNO AND TT.FCBA = ID.FCBA)" & _
									'						"  WHERE  " & _
									'						"  ID.ORDERNO = '" & rsCursorSivCode("ORDERNO") & "' " & _
									'						"  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'mm')) = " & Month(Session("v_Stores")) &  _
									'						"  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'yyyy')) =" & Year(Session("v_Stores"))  & _
									'						"  AND ID.fcBA = '" & Session("fcBA") & "'" & _
									'						"  AND TT.ROWSTATEID = 'Approved'" & _	
									'						"  AND ID.ORDERNO = '" & tempSIVCODE & "'  " & _
									'						"  AND ID.STOCKCODE = '" & tempStockCode & "' " & _
									'						"  AND ID.LINENO = '" & rsCursorCredit ("linenokey") & "' " & _
									'						"  AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
									'						"  AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
									'						"  AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") & "' " 
															
									' Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
									' rsReceiveDetail.CursorType = 3
									' rsReceiveDetail.CursorLocation = 3
									' rsReceiveDetail.Open strCursorDebitDetail, connstr
									' if not rsReceiveDetail.eof then
										' Do While Not rsReceiveDetail.eof
										
											' maxLineNoDetail = maxLineNoDetail + 1
											' qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
															' "( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
															' "  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
															' " QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
															' maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
															' txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
															' cdbl(rsReceiveDetail("QUANTITY")) & ",0,"& cdbl(rsReceiveDetail("AMOUNT")) &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
															' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
											
											' response.write qCost_ActivityDBHistory & "</br>"
											' connstr.execute(qCost_ActivityDBHistory)		
										
											' rsReceiveDetail.MoveNext
										' loop
									' end if
									' rsReceiveDetail.Close
									' set rsReceiveDetail = Nothing							
								'* Remarks By Iwa * 
								
								
								else
									Response.Flush()
									maxLineNo = maxLineNo + 1
									qInsertDebit = " INSERT INTO cost_activity ( LINENO ," & _
																			  " TDATE, " & _
																			  " JOBCODE, " & _
																			  " SOURCE, " & _
																			  " SOURCETYPE, " & _
																			  " SOURCECODE, " & _
																			  " SOURCECODE_STATUS, " & _
																			  " TARGETTYPE, " & _
																			  " TARGETCODE, " & _
																			  " TARGETCODE_STATUS, " & _
																			  " REFERENCENO, " & _
																			  " QUANTITY, " & _
																			  " OUTPUT, " & _
																			  " AMOUNT, " & _
																			  " FCBA, " & _
																			  " CLOSING_DATE, " & _ 
																			  " PROCESS_DATE, " & _		
																			  " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																			  " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																			  " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																			  " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																			  " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _																				  
																			  " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _	
																			  " LASTUPDATE,  " & _
																			  " LASTTIME ) " & _
															" VALUES (" & _
															maxLineNo & "," & _
															" to_date('" & rsCursorDebit("tdate") & "','mm/dd/yyyy'),'" & _
															rsCursorDebit("JOBCODE") & "','" & _
															rsCursorDebit("SOURCE") & "','" & _
															rsCursorDebit("SOURCETYPE") & "','" & _
															rsCursorDebit("SOURCECODE") & "','" & _
															rsCursorCredit("SOURCECODE_STATUS") & "','" & _
															rsCursorDebit("TARGETTYPE") & "','" & _
															rsCursorDebit("TARGETCODE") & "','" & _
															rsCursorDebit("TARGETCODE_STATUS") & "','" & _
															rsCursorDebit("REFERENCENO") & "'," & _
															rsCursorDebit("QUANTITY") & "," & _
															rsCursorDebit("OUTPUT") & "," & _
															tempDebit & ",'" & _
															rsCursorDebit("FCBA") & "'," & _
															" to_date('" & rsCursorDebit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
															" sysdate, " & _
															" Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & rsCursorDebit("TARGETCODE") & "'," & _
															" Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempSIVCODE & "'," & _
															" Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempStockCode & "'," & _
															" '',''," & _
															" '','','Dr'," & _
															Month(Session("v_Stores")) & ", " & _
															Year(Session("v_Stores")) & ", " & _
															" sysdate, " & _
															" to_char(sysdate,'HH24:MI')) "
									response.Flush()		  
									connstr.Execute(qInsertDebit)	
									
									'* Remarks By Iwa * 
									' strCursorDebitDetail =  " SELECT  ID.FCBA," & _
															' "		  ID.ORDERNO REFERENCENO," & _
															' "		  ID.STOCKCODE, " & _ 
															' "		  TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO) TDATE , " & _
															' "		  ID.JOBCODE," & _
															' "		  ID.STORECODE SOURCECODE, " & _
															' "		  ID.LINENO LINENOKEY, " & _
															' "		  ID.LOCATIONTYPE, " & _
															' "		  ID.LOCATIONCODE, " & _
															' "		  'ST' SOURCE, " & _
															' "		  'ST' SOURCETYPE, " & _
															' "		  '' SOURCECODE_STATUS, " & _
															' "		  'CB' TARGETTYPE, " & _
															' "		  'BS' TARGETCODE, " & _
															' "		  ''   TARGETCODE_STATUS, " & _
															' "		  ID.QUANTITYTRANSFERED QUANTITY, " & _
															' "		  0 OUTPUT, " & _
															' "		  ROUND(ID.AMOUNT * 1) AMOUNT, " & _
															' "		  TO_DATE('" & TEMP_DATEPERIODCONTROL & "','MM/DD/YYYY') CLOSING_DATE, " & _
															' "		  SYSDATE PROCESS_DATE " & _
															' "  FROM TRANSPORTTASKDTL_TAB ID " & _
															' "	INNER JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = ID.ORDERNO AND TT.FCBA = ID.FCBA)" & _
															' "  WHERE  " & _
															' "  ID.ORDERNO = '" & rsCursorSivCode("ORDERNO") & "' " & _
															' "  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'mm')) = " & Month(Session("v_Stores")) &  _
															' "  AND TO_NUMBER(TO_CHAR(TRANSPORTTASK_API.GET_RCVDATE(ID.FCBA,ID.ORDERNO), 'yyyy')) =" & Year(Session("v_Stores"))  & _
															' "  AND ID.fcBA = '" & Session("fcBA") & "'" & _
															' "  AND TT.ROWSTATEID = 'Approved'" & _	
															' "  AND ID.ORDERNO = '" & tempSIVCODE & "'  " & _
															' "  AND ID.STOCKCODE = '" & tempStockCode & "' " & _
															' "  AND ID.LINENO = '" & rsCursorCredit ("linenokey") & "' " & _
															' "  AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
															' "  AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
															' "  AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") & "' " 											 
							
							
									' Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
									' rsReceiveDetail.CursorType = 3
									' rsReceiveDetail.CursorLocation = 3
									' rsReceiveDetail.Open strCursorDebitDetail, connstr
									' if not rsReceiveDetail.eof then
										' Do While Not rsReceiveDetail.eof
										
											' maxLineNoDetail = maxLineNoDetail + 1
											' qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
															' "( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
															' "  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
															' " QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
															' maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
															' txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
															' cdbl(rsReceiveDetail("QUANTITY")) & ",0,"& cdbl(rsReceiveDetail("AMOUNT")) &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
															' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
											
											' response.write qCost_ActivityDBHistory & "</br>"
											' connstr.execute(qCost_ActivityDBHistory)		
										
											' rsReceiveDetail.MoveNext
										' loop
									' end if
									' rsReceiveDetail.Close
									' set rsReceiveDetail = Nothing							
									'* Remarks By Iwa * 
								
									tempTotalAmountDebit = tempTotalAmountDebit + tempDebit
									totalRowInserted = totalRowInserted + 1
								end if	
								tempRowNum = tempRowNum + 1
								rsCursorDebit.MOVENEXT
							loop
							rsCursorDebit.close
							set rsCursorDebit=Nothing				
						end if
						rsCursorCredit.MOVENEXT
					loop
					rsCursorCredit.close
					set rsCursorCredit=Nothing		
				end if
				rsCursorSivCode.MOVENEXT
			loop
			rsCursorSivCode.close
			set rsCursorSivCode=Nothing
		End if
		
		Response.Write "<br><font color=Blue>" & "Proses Cost Activity from Transport Task is completed..." & "</font>"	
		
		' Add By Akbar 30 Sept 2016 Transport Task for Cost Activity 
		
		' Add By Akbar  7 Oct 2016  Received Task for Cost Activity 
		
		tempRefCodeReceive = "RECEIVEDTASK"

	
		qReference =  "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
					  " from PROCESS_TABLEREF_TAB " & _
					  " where " & _
					  " processcode = '" & tempProcessCode & "' and " & _
					  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
					  " fccode =  '" & tempRefCodeReceive & "'"

		Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
		rsReferenceReceive.CursorType = 3
		rsReferenceReceive.CursorLocation = 3
		rsReferenceReceive.Open qReference, connstr

		doneReceive = 1
		if not rsReferenceReceive.eof then
			txtRefNumberKeys_Receive = rsReferenceReceive("REF_NUMBERKEYS")
			txtRefStringSeparator_Receive = rsReferenceReceive("REF_STRINGSEPARATOR")
			txtRefStringKeys_Receive = rsReferenceReceive("REF_STRINGKEYS")
			doneReceive = 1
		else
			doneReceive = 0
		end if
		rsReferenceReceive.close
		set rsReferenceReceive = Nothing
		
		qReceive=   " select param_1 from iplasprod.parameterdetail" & _
					" where parcode = 'HARDCODE'" & _
					" and parheadcode = 'CONTROLJOB_MODULE'" & _
					" and fccode = 'RECEIVEDTASK_SBU'" & _
					" and fcba = '"& session("fcba") &"'"  
		
		Set rsTemp = Server.CreateObject("ADODB.Recordset")			
		rsTemp.Open qReceive, connstr
		if Not rsTemp.Eof then 
			param_1 = cdbl(to_nol(rsTemp("param_1")))  
		End if
		
		qReceive= " Select RD.FCBA,RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO) TDATE, RD.ORDERNO,RD.STOCKCODE," & _
				  "		   INTRACO_API.GET_JOBCODE(RD.FCBASOURCE) CONTROLJOB," & _
				  "			INTRACO_API2.GET_JOBCODE(RD.FCBASOURCE,'DBU') CONTROLJOB2," & _
				  "		   TD.JOBCODE JOBCODE_TT, RD.JOBCODE,RD.QUANTITYRECEIVED QUANTITY,RD.LOCATIONTYPE,RD.LOCATIONCODE," & _
				  "		   RD.FCBASOURCE ,RD.ORDERSOURCE, RD.LINENOSOURCE, ROUND(NVL(sum(RD.AMOUNTRECEIVED),0),2) as AMOUNTRECEIVED" & _
				  " FROM RECEIVEDTASKDTL_TAB RD " & _
				  " INNER JOIN TRANSPORTTASKDTL_TAB TD ON (TD.ORDERNO = RD.ORDERSOURCE AND TD.FCBADISTINATION = RD.FCBA AND TD.STOCKCODE = RD.STOCKCODE AND TD.LINENO = RD.LINENOSOURCE)" & _
				  " WHERE to_number(to_Char(RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO),'mm'))= " & Month(Session("v_Stores")) &  _
				  " And to_number(to_Char(RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO),'yyyy'))= " & Year(Session("v_Stores")) & _
				  " And RD.fcBA='" & Session("fcBA") & "'" & _
				  " Group by RD.FCBA,RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO), RD.ORDERNO,RD.STOCKCODE,RD.JOBCODE, TD.JOBCODE ,RD.QUANTITYRECEIVED,RD.LOCATIONTYPE,RD.LOCATIONCODE," & _
				  " RD.FCBASOURCE ,RD.ORDERSOURCE, RD.LINENOSOURCE"
		  
		Set rsReceive = Server.CreateObject("ADODB.Recordset")
		rsReceive.CursorType = 3
		rsReceive.CursorLocation = 3
		'response.Write(qReceive)
		rsReceive.Open qReceive, connstr
		if not rsReceive.eof then
		   Do While Not rsReceive.eof
		   
				' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
				txtRefStringValues_Receive = rsReceive("ORDERNO") & txtRefStringSeparator_Receive & _
										   rsReceive("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
										   rsReceive("LOCATIONCODE") & txtRefStringSeparator_Receive & _
										   rsReceive("JOBCODE") & txtRefStringSeparator_Receive & _
										   rsReceive("LINENOSOURCE") & txtRefStringSeparator_Receive & _
										   Session("fcBA")

				maxLineNo = maxLineNo + 1		  
				qCost_ActivityDB="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ", to_date('" & rsReceive("TDATE") & "','mm/dd/yyyy') " & ",'" & rsReceive("JOBCODE") & "','ST','" & "ST" & "','" & rsReceive("ORDERSOURCE") & "'," & _
							 "'" &  rsReceive("STOCKCODE") & "','" & rsReceive("LOCATIONTYPE") & "','" & rsReceive("LOCATIONCODE") & "'," & _
							 "" &  rsReceive("AMOUNTRECEIVED") & ",'" & rsReceive("ORDERNO") & "'," & rsReceive("QUANTITY") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate , " & _
							 " Cost_Activity_Info_Api.Get_Title_Reference_01('ST', '" & rsReceive("LOCATIONTYPE") & "','Dr'),'"  & rsReceive("LOCATIONCODE") & "'," & _
							 " Cost_Activity_Info_Api.Get_Title_Reference_02('ST', '" & rsReceive("LOCATIONTYPE") & "','Dr'),'"  & rsReceive("ORDERNO") & "'," & _
							 " Cost_Activity_Info_Api.Get_Title_Reference_03('ST', '" & rsReceive("LOCATIONTYPE") & "','Dr'),'"  & rsReceive("ORDERSOURCE") & "'," & _
							 " Cost_Activity_Info_Api.Get_Title_Reference_04('ST', '" & rsReceive("LOCATIONTYPE") & "','Dr'),'"  & rsReceive("STOCKCODE") & "'," & _
							 " '','','Dr', " & _
							 Month(Session("v_Stores")) & ", " & _
							 Year(Session("v_Stores")) & ", " & _
							 " sysdate, to_char(sysdate,'HH24:MI'))"
							 
				
				response.Flush()
				connstr.execute(qCost_ActivityDB)
			    connStr.execute("Commit")
				 
				 '* Remarks By Iwa * 
				 
				' qVoucherDetail =  " Select RD.FCBA,RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO) TDATE, RD.ORDERNO,RD.STOCKCODE," & _
								  ' "		   INTRACO_API.GET_JOBCODE(RD.FCBASOURCE) CONTROLJOB," & _
								  ' "		   RD.JOBCODE,RD.QUANTITYRECEIVED QUANTITY,RD.LOCATIONTYPE,RD.LOCATIONCODE," & _
								  ' "		   RD.FCBASOURCE ,RD.ORDERSOURCE, RD.LINENOSOURCE, NVL(RD.AMOUNTRECEIVED,0) as AMOUNTRECEIVED" & _
								  ' " FROM RECEIVEDTASKDTL_TAB RD " & _
								  ' " WHERE to_number(to_Char(RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO),'mm'))= " & Month(Session("v_Stores")) &  _
								  ' " And to_number(to_Char(RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO),'yyyy'))= " & Year(Session("v_Stores")) & _
								  ' " And RD.fcBA='" & Session("fcBA") & "'" & _
								  ' " and RD.ORDERNO = '" & rsReceive("ORDERNO") & "'" & _
								  ' "	and RD.LOCATIONTYPE	= '" &  rsReceive("LOCATIONTYPE") & "'" & _
								  ' "	and RD.LOCATIONCODE	= '" & rsReceive("LOCATIONCODE") & "'" & _
								  ' " and RD.JOBCODE = '" &  rsReceive("JOBCODE")& "'" & _
								  ' " and RD.LINENOSOURCE = '" & rsReceive("LINENOSOURCE")& "'" 		  
								  
							  
				' Set rsVoucherDetail = Server.CreateObject("ADODB.Recordset")
				' rsVoucherDetail.CursorType = 3
				' rsVoucherDetail.CursorLocation = 3
				' rsVoucherDetail.Open qVoucherDetail, connstr
				
				' if not rsVoucherDetail.eof then
					' maxLineNoDetail = 0
					' Do While Not rsVoucherDetail.eof
					

						' maxLineNoDetail = maxLineNoDetail + 1
						' qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
										' "( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
										' "  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
										' " QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME , OTHER_REFERENCE_C01, OTHER_REFERENCE_C02 ) values (" & _
										' maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
										' txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
										' rsVoucherDetail("QUANTITY") & ",0,"& rsVoucherDetail("AMOUNTRECEIVED")&",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
										' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'), '" & rsVoucherDetail("ORDERSOURCE") & _
										' "','" & rsVoucherDetail("LINENOSOURCE") & "')"
						
				' response.write qCost_ActivityDBHistory & "</br>"						
						' connstr.execute(qCost_ActivityDBHistory)		
						' connStr.execute("Commit")
						' rsVoucherDetail.MoveNext
					' loop
				' end if
				' rsVoucherDetail.Close
				' set rsVoucherDetail = Nothing	
				'* Remarks By Iwa *  
				
				If param_1 = rsReceive("JOBCODE_TT")  Then
					maxLineNo = maxLineNo + 1
					qCost_ActivityCR="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _							 
								 " LASTUPDATE, LASTTIME) " & _
								 " values (" & maxLineNo & ", to_date('" & rsReceive("TDATE") & "','mm/dd/yyyy') " & ",'" & rsReceive("CONTROLJOB") & "','ST','" & "ST" & "','" & rsReceive("ORDERSOURCE") & "'," & _
								 "'" &  rsReceive("STOCKCODE") & "','" & "CB" & "','" & "BS" & "'," & _
								 -1*cdbl(rsReceive("AMOUNTRECEIVED")) & ",'" & rsReceive("ORDERNO") & "'," & rsReceive("QUANTITY") & ",0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " Cost_Activity_Info_Api.Get_Title_Reference_01('ST', 'CB','Dr'),'"  & rsReceive("FCBASOURCE") & "'," & _
								 " Cost_Activity_Info_Api.Get_Title_Reference_02('ST', 'CB','Dr'),'"  & rsReceive("ORDERNO") & "'," & _
								 " Cost_Activity_Info_Api.Get_Title_Reference_03('ST', 'CB','Dr'),'"  & rsReceive("ORDERSOURCE") & "'," & _
								 " Cost_Activity_Info_Api.Get_Title_Reference_04('ST', 'CB','Dr'),'"  & rsReceive("STOCKCODE") & "'," & _
								 " '','','Cr', " & _
								 Month(Session("v_Stores")) & ", " & _
								 Year(Session("v_Stores")) & ", " & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "
								 
					connstr.execute(qCost_ActivityCR)
					connStr.execute("Commit")
				ELSE 
					maxLineNo = maxLineNo + 1
					qCost_ActivityCR="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _							 
								 " LASTUPDATE, LASTTIME) " & _
								 " values (" & maxLineNo & ", to_date('" & rsReceive("TDATE") & "','mm/dd/yyyy') " & ",'" & rsReceive("CONTROLJOB2") & "','ST','" & "ST" & "','" & rsReceive("ORDERSOURCE") & "'," & _
								 "'" &  rsReceive("STOCKCODE") & "','" & "CB" & "','" & "BS" & "'," & _
								 -1*cdbl(rsReceive("AMOUNTRECEIVED")) & ",'" & rsReceive("ORDERNO") & "'," & rsReceive("QUANTITY") & ",0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " Cost_Activity_Info_Api.Get_Title_Reference_01('ST', 'CB','Dr'),'"  & rsReceive("FCBASOURCE") & "'," & _
								 " Cost_Activity_Info_Api.Get_Title_Reference_02('ST', 'CB','Dr'),'"  & rsReceive("ORDERNO") & "'," & _
								 " Cost_Activity_Info_Api.Get_Title_Reference_03('ST', 'CB','Dr'),'"  & rsReceive("ORDERSOURCE") & "'," & _
								 " Cost_Activity_Info_Api.Get_Title_Reference_04('ST', 'CB','Dr'),'"  & rsReceive("STOCKCODE") & "'," & _
								 " '','','Cr', " & _
								 Month(Session("v_Stores")) & ", " & _
								 Year(Session("v_Stores")) & ", " & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "
								 
					connstr.execute(qCost_ActivityCR)
					connStr.execute("Commit")
				End If 
				'* Remarks By Iwa * 
				
				' qVoucherDetail =  " Select RD.FCBA,RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO) TDATE, RD.ORDERNO,RD.STOCKCODE," & _
								  ' "		   INTRACO_API.GET_JOBCODE(RD.FCBASOURCE) CONTROLJOB," & _
								  ' "		   RD.JOBCODE,RD.QUANTITYRECEIVED QUANTITY,RD.LOCATIONTYPE,RD.LOCATIONCODE," & _
								  ' "		   RD.FCBASOURCE ,RD.ORDERSOURCE, RD.LINENOSOURCE, NVL(RD.AMOUNTRECEIVED,0) as AMOUNTRECEIVED" & _
								  ' " FROM RECEIVEDTASKDTL_TAB RD " & _
								  ' " WHERE to_number(to_Char(RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO),'mm'))= " & Month(Session("v_Stores")) &  _
								  ' " And to_number(to_Char(RECEIVEDTASK_API.GET_RCVDATE(RD.FCBA,RD.ORDERNO),'yyyy'))= " & Year(Session("v_Stores")) & _
								  ' " And RD.fcBA='" & Session("fcBA") & "'" & _
								  ' " and RD.ORDERNO = '" & rsReceive("ORDERNO") & "'" & _
								  ' "	and RD.LOCATIONTYPE	= '" &  rsReceive("LOCATIONTYPE") & "'" & _
								  ' "	and RD.LOCATIONCODE	= '" & rsReceive("LOCATIONCODE") & "'" & _
								  ' " and RD.JOBCODE = '" &  rsReceive("JOBCODE")& "'" & _
								  ' " and RD.LINENOSOURCE = '" & rsReceive("LINENOSOURCE")& "'" 
							  
				' Set rsVoucherDetail = Server.CreateObject("ADODB.Recordset")
				' rsVoucherDetail.CursorType = 3
				' rsVoucherDetail.CursorLocation = 3
				' rsVoucherDetail.Open qVoucherDetail, connstr
				
				' if not rsVoucherDetail.eof then
					' maxLineNoDetail = 0
					' Do While Not rsVoucherDetail.eof

						' maxLineNoDetail = maxLineNoDetail + 1
				
						
						' qCost_ActivityDBHistory = "INSERT INTO COST_ACTIVITY_HIST_TAB " & _
										' "( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
										' "  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
										' " QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, OTHER_REFERENCE_C01, OTHER_REFERENCE_C02  ) values (" & _
										' maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
										' txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
										' rsVoucherDetail("QUANTITY") & ",0,"& -1 * cdbl(rsVoucherDetail("AMOUNTRECEIVED")) &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
										' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'), '" & rsVoucherDetail("ORDERSOURCE") & _
										' "','" & rsVoucherDetail("LINENOSOURCE") & "')"
						' response.write qCost_ActivityDBHistory & "</br>"
						'response.write qCost_ActivityDBHistory & "</br>"
						' connstr.execute(qCost_ActivityDBHistory)		
						' connStr.execute("Commit")
						' rsVoucherDetail.MoveNext
					' loop
				' end if
				' rsVoucherDetail.Close
				' set rsVoucherDetail = Nothing		
				'* Remarks By Iwa * 				
			  
			  'response.Write "Process GRN " & rsReceive("RECEIVENOTECODE") & "<br>"
			  response.Flush()
			  rsReceive.MoveNext
		  loop
		End if
		
		Response.Write "<br><font color=Blue>" & "Proses Cost Activity from Received Task is completed..." & "</font>"	
		
		
		' NAF : We change the Old Process become the New ONE
		' Cause, we should remove COMMA decimal at the last point of each number
		
		' Logic : 
		' Buat Cursor 1 per SIVCODE
		' Open Cursor 1
		'      Insert Credit Into Cost Activity
		'      tempTotalAmountCredit, didapatkan dari Cursor 1
		'      Buat Cursor 2 per SIVCODE and LocationType, LocationCode
		'      Buat Count(*) total Recordnya ==> n
		'      totalRowInserted = 0
		' 	   tempTotalAmountDebit = 0
		'	   Open Cursor 2
		'		   If totalRowInserted = n  then 
		'			   tempTotalAmountDebit = tempTotalAmountDebit + dari Cursor 2 setelah di round
		'			   tempSelisih = tempTotalAmountCredit - tempTotalAmountDebit				   
		'			   Insert Debit ==> Amount setelah di Round + tempSelisih
		'           else
		'		       Insert Debit
		'			   tempTotalAmountDebit = tempTotalAmountDebit + dari Cursor 2 setelah di round
		'			   n = n + 1
		' 		   end if
		' 	   End Loop
		' End Loop
		
		
		
		

		tempRefCodeReceive = "ISSUEDETAILS"

		' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
		' by NAF, 12 May 2013
		' First, take the Original Table of RECEIVENOTE.
		qReference = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
					  " from process_tableref_tab " & _
					  " where " & _
					  " processcode = '" & tempProcessCode & "' and " & _
					  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
					  " fccode =  '" & tempRefCodeReceive & "'"

		Set rsReferenceReceive = Server.CreateObject("ADODB.Recordset")
		rsReferenceReceive.CursorType = 3
		rsReferenceReceive.CursorLocation = 3
		rsReferenceReceive.Open qReference, connstr

		doneReceive = 1
		if not rsReferenceReceive.eof then
			txtRefNumberKeys_Receive = rsReferenceReceive("REF_NUMBERKEYS")
			txtRefStringSeparator_Receive = rsReferenceReceive("REF_STRINGSEPARATOR")
			txtRefStringKeys_Receive = rsReferenceReceive("REF_STRINGKEYS")
			doneReceive = 1
		else
			doneReceive = 0
		end if
		rsReferenceReceive.close
		set rsReferenceReceive = Nothing	
		
		
		strCursorSivCode = " select sivcode " & _
							" from issue " & _
							" where " & _
							" to_number(to_char(sivdate, 'mm')) = " & month (session ("v_stores")) & _
							" and to_number(to_char(sivdate, 'yyyy')) = " & year (session ("v_stores")) & _
							" and fcba = '" & Session("fcBA") & "'" 
							
		Set rsCursorSivCode = Server.CreateObject("ADODB.Recordset")
		rsCursorSivCode.CursorType = 3
		rsCursorSivCode.CursorLocation = 3
		rsCursorSivCode.Open strCursorSivCode, connstr		

		if not rsCursorSivCode.eof then
			Do While Not rsCursorSivCode.eof
			
				' NAF : Jangan dirubah format string-nya menjadi lurus atau via excel, biarkan seperti apa adanya atau 
				'       Jangan menggunakan TRIM untuk menghilangkan SPASI-nya "
			

				
				strCursorCredit = " SELECT isu.SIVDATE tdate , " & _
										 " stockgroup.CONTROLJOB jobcode, " & _
										 " 'ST' source, " & _
										 " 'ST' sourcetype, " & _
										 " ISU.STORECODE sourcecode, " & _
										 " '' sourcecode_status, " & _
										 " 'CB' targettype, " & _
										 " 'BS' targetcode, " & _
										 " ''   targetcode_status, " & _
										 " id.SIVCODE referenceno, " & _
										 " id.linenokey, " & _
										 " id.locationtype, " & _
										 " id.locationcode, " & _
										 " id.jobcode jobcodeoriginal, " & _
										 " SUM (id.QUANTITY) quantity, " & _
										 " 0 output, " & _
										 " round( (SUM (id.AMOUNT) * -1) ,2) amount, " & _
										 " id.FCBA fcba, " & _
										 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date, " & _
										 " SYSDATE process_date, " & _
										 " ID.STOCKCODE STOCKCODE " & _
									" FROM ISSUE ISU, " & _
										 " ISSUEDETAILS ID, " & _
										 " Stock, " & _
										 " STOCKGROUP " & _
								   " WHERE (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA) " & _
										 " AND ISU.SIVCODE = '" & rsCursorSivCode("sivcode") & "' " & _
										 " AND (id.STOCKCODE = stock.FCCODE AND id.fcba = stock.fcba) " & _
										 " AND (stock.STOCKGROUPCODE = stockgroup.FCCODE " & _
										 " AND stock.fcba = stockgroup.fcba) " & _
										 " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
										 " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy')) =" & Year(Session("v_Stores"))  & _
										 " AND ISU.fcBA = '" & Session("fcBA") & "'" & _
								" GROUP BY isu.SIVDATE, " & _
										 " stockgroup.CONTROLJOB, " & _
										 " 'ST', " & _
										 " 'ST', " & _
										 " ISU.STORECODE, " & _
										 " '', " & _
										 " 'CB', " & _
										 " 'BS', " & _
										 " '', " & _
										 " id.SIVCODE, " & _
										 " id.linenokey, " & _
										 " id.locationtype, " & _
										 " id.locationcode, " & _
										 " id.jobcode, " & _
										 " id.fcBA, " & _
										 " id.STOCKCODE "
										 
				
				Set rsCursorCredit = Server.CreateObject("ADODB.Recordset")
				rsCursorCredit.CursorType = 3
				rsCursorCredit.CursorLocation = 3
				rsCursorCredit.Open strCursorCredit, connstr	
				
				
				if not rsCursorCredit.eof then
				
					Do While Not rsCursorCredit.eof
				
						tempTotalAmountCredit = round(cdbl(rsCursorCredit("AMOUNT")),2)
						tempSIVCODE = rsCursorCredit("referenceno")
						tempStockCode = rsCursorCredit("STOCKCODE")

						txtRefStringValues_Receive = tempSIVCODE & txtRefStringSeparator_Receive & _
													 rsCursorCredit("linenokey") & txtRefStringSeparator_Receive & _
													 rsCursorCredit("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
													 rsCursorCredit("LOCATIONCODE") & txtRefStringSeparator_Receive & _
													 rsCursorCredit("jobcodeoriginal") & txtRefStringSeparator_Receive & _
													 rsCursorCredit("FCBA") 
													 
						
						' NAF : Jangan dirubah format string-nya menjadi lurus atau via excel, biarkan seperti apa adanya atau 
						'       Jangan menggunakan TRIM untuk menghilangkan SPASI-nya "		

						
						strCursorCount = " select count(*) totalRecord from ( " & _
												" SELECT isu.SIVDATE, " & _
												"	 id.JOBCODE, " & _
												"	 ISU.STORECODE, " & _
												"	 id.LOCATIONTYPE, " & _
												"	 id.LOCATIONCODE, " & _
												"	 id.SIVCODE, " & _
												"    id.LINENOKEY, " & _
												"	 id.FCBA " & _
												" FROM ISSUE ISU, ISSUEDETAILS ID  " & _
											   " WHERE	 (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA)  " & _
												"	 AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
												"	 AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy'))  =" & Year(Session("v_Stores"))  & _
												"	 AND ISU.fcBA = '" & Session("fcBA") & "'" & _
												"	 AND ISU.sivcode = '" & tempSIVCODE & "'  " & _
												"    AND id.linenokey = '" & rsCursorCredit ("linenokey") & "' " & _
												"    AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
												"    AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
												"    AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") &  "' " & _
												"    AND ID.STOCKCODE = '" & tempStockCode & "' " & _
											" GROUP BY isu.SIVDATE,  " & _
												"	 id.JOBCODE,  " & _
												"	 ISU.STORECODE,  " & _
												"	 id.LOCATIONTYPE,  " & _
												"	 id.LOCATIONCODE,  " & _
												"	 id.SIVCODE,  " & _
												"    id.LINENOKEY, " & _
												"	 id.fcBA  " & _
											" ) "  

					
						Set rsCursorCount = Server.CreateObject("ADODB.Recordset")
						rsCursorCount.CursorType = 3
						rsCursorCount.CursorLocation = 3
						

						rsCursorCount.Open strCursorCount, connstr		

						totalRecord = cdbl(rsCursorCount("totalRecord"))
						totalRowInserted = 0
						tempTotalAmountDebit = 0

						
						Response.Flush()
						
						
						maxLineNo = maxLineNo + 1
						strInsertCredit = " INSERT INTO cost_activity ( LINENO , " & _
																  " TDATE, " & _
																  " JOBCODE, " & _
																  " SOURCE, " & _
																  " SOURCETYPE, " & _
																  " SOURCECODE, " & _
																  " SOURCECODE_STATUS, " & _
																  " TARGETTYPE, " & _
																  " TARGETCODE, " & _
																  " TARGETCODE_STATUS, " & _
																  " REFERENCENO, " & _
																  " QUANTITY, " & _
																  " OUTPUT, " & _
																  " AMOUNT, " & _
																  " FCBA, " & _
																  " CLOSING_DATE, " & _ 
																  " PROCESS_DATE , " & _
																  " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																  " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																  " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																  " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																  " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
																  " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _	
																  " LASTUPDATE,  " & _
																  " LASTTIME ) " & _
												" VALUES (" & _
												maxLineNo & "," & _
												" to_date('" & rsCursorCredit("tdate") & "','mm/dd/yyyy'),'" & _
												rsCursorCredit("JOBCODE") & "','" & _
												rsCursorCredit("SOURCE") & "','" & _
												rsCursorCredit("SOURCETYPE") & "','" & _
												rsCursorCredit("SOURCECODE") & "','" & _
												rsCursorCredit("SOURCECODE_STATUS") & "','" & _
												rsCursorCredit("TARGETTYPE") & "','" & _
												rsCursorCredit("TARGETCODE") & "','" & _
												"NAF||" & rsCursorCredit("TARGETCODE_STATUS") & "','" & _
												rsCursorCredit("REFERENCENO") & "'," & _
												rsCursorCredit("QUANTITY") & "," & _
												rsCursorCredit("OUTPUT") & "," & _
												tempTotalAmountCredit & ",'" & _
												rsCursorCredit("FCBA") & "'," & _
												" to_date('" & rsCursorCredit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
												" sysdate, " & _
												" Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & rsCursorCredit("TARGETCODE") & "'," & _
												" Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & tempSIVCODE & "'," & _
												" Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & tempStockCode & "'," & _
												" '',''," & _
												" '','','Cr'," & _
												Month(Session("v_Stores")) & ", " & _
												Year(Session("v_Stores")) & ", " & _
												" sysdate, " & _
												" to_char(sysdate,'HH24:MI')) " 
 
						response.Flush()		  
						connstr.Execute(strInsertCredit)		
						
						'* Remarks By Iwa * 
						' strCursorCreditDetail = " SELECT isu.SIVDATE tdate , " & _
													 ' " stockgroup.CONTROLJOB jobcode, " & _
													 ' " 'ST' source, " & _
													 ' " 'ST' sourcetype, " & _
													 ' " ISU.STORECODE sourcecode, " & _
													 ' " '' sourcecode_status, " & _
													 ' " 'CB' targettype, " & _
													 ' " 'BS' targetcode, " & _
													 ' " ''   targetcode_status, " & _
													 ' " id.SIVCODE referenceno, " & _
													 ' " id.linenokey, " & _
													 ' " id.locationtype, " & _
													 ' " id.locationcode, " & _
													 ' " id.jobcode jobcodeoriginal, " & _
													 ' " id.QUANTITY quantity, " & _
													 ' " 0 output, " & _
													 ' " round(id.AMOUNT * -1) amount, " & _
													 ' " id.FCBA fcba, " & _
													 ' " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date, " & _
													 ' " SYSDATE process_date, " & _
													 ' " ID.STOCKCODE STOCKCODE " & _
												' " FROM ISSUE ISU, " & _
													 ' " ISSUEDETAILS ID, " & _
													 ' " Stock, " & _
													 ' " STOCKGROUP " & _
											   ' " WHERE (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA) " & _
													 ' " AND ISU.SIVCODE = '" & rsCursorSivCode("sivcode") & "' " & _
													 ' " AND (id.STOCKCODE = stock.FCCODE AND id.fcba = stock.fcba) " & _
													 ' " AND (stock.STOCKGROUPCODE = stockgroup.FCCODE " & _
													 ' " AND stock.fcba = stockgroup.fcba) " & _
													 ' " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
													 ' " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy')) =" & Year(Session("v_Stores"))  & _
													 ' " AND ISU.fcBA = '" & Session("fcBA") & "'" & _
													 ' " AND ISU.sivcode = '" & tempSIVCODE & "'  " & _
													 ' " AND id.linenokey = '" & rsCursorCredit ("linenokey") & "' " & _
													 ' " AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
													 ' " AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
													 ' " AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") &  "' " & _
													 ' " AND ID.STOCKCODE = '" & tempStockCode & "' " 


						' Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
						' rsReceiveDetail.CursorType = 3
						' rsReceiveDetail.CursorLocation = 3
						' rsReceiveDetail.Open strCursorCreditDetail, connstr
						' if not rsReceiveDetail.eof then
							' Do While Not rsReceiveDetail.eof
							
								' maxLineNoDetail = maxLineNoDetail + 1
								' qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
												' "( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
												' "  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
												' " QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
												' maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
												' txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
												' cdbl(rsReceiveDetail("QUANTITY")) & ",0,"& cdbl(rsReceiveDetail("AMOUNT")) &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
												' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
								
								' response.write qCost_ActivityDBHistory & "</br>"
								' connstr.execute(qCost_ActivityDBHistory)		
								
								' rsReceiveDetail.MoveNext
							' loop
						' end if
						' rsReceiveDetail.Close
						' set rsReceiveDetail = Nothing			
						'* Remarks By Iwa * 
						
						' NAF : Jangan dirubah format string-nya menjadi lurus atau via excel, biarkan seperti apa adanya atau 
						'       Jangan menggunakan TRIM untuk menghilangkan SPASI-nya "			
						strCursorDebit = " SELECT isu.SIVDATE tdate, " & _
													 " id.JOBCODE jobcode,  " & _
													 " 'ST' source,  " & _
													 " 'ST' sourcetype,  " & _
													 " ISU.STORECODE sourcecode,  " & _
													 " '' sourcecode_status,  " & _
													 " id.LOCATIONTYPE targettype,  " & _
													 " id.LOCATIONCODE targetcode,  " & _
													 " '' targetcode_status,  " & _
													 " id.SIVCODE referenceno,  " & _
													 " id.linenokey, " & _
													 " id.locationtype, " & _
													 " id.locationcode, " & _
													 " id.jobcode, " & _												 
													 " SUM (id.QUANTITY) quantity,  " & _
													 " 0 output,  " & _
													 " round( (SUM (id.AMOUNT) * 1) ,2) amount,  " & _
													 " id.FCBA fcba,  " & _
													 " TO_DATE ('" & temp_datePeriodControl & "', 'mm/dd/yyyy') closing_date,  " & _
													 " SYSDATE process_date, " & _
													 " ID.STOCKCODE " & _
												" FROM ISSUE ISU, ISSUEDETAILS ID  " & _
											   " WHERE	 (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA)  " & _
													 " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
													 " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy'))  =" & Year(Session("v_Stores"))  & _
													 " AND ISU.fcBA = '" & Session("fcBA") & "'" & _
													 " AND ISU.sivcode = '" & tempSIVCODE & "'  " & _
													 " AND ID.STOCKCODE = '" & tempStockCode & "' " & _
													 " AND ID.LINENOKEY = '" & rsCursorCredit ("linenokey") & "' " & _
													 " AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
													 " AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
													 " AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") & "' " & _													 
											" GROUP BY isu.SIVDATE,  " & _
													 " id.JOBCODE,  " & _
													 " 'ST',  " & _
													 " 'ST',  " & _
													 " ISU.STORECODE,  " & _
													 " '',  " & _
													 " id.LOCATIONTYPE,  " & _
													 " id.LOCATIONCODE,  " & _
													 " '',  " & _
													 " id.SIVCODE,  " & _
													 " id.linenokey, " & _
													 " id.locationtype, " & _
													 " id.locationcode, " & _
													 " id.jobcode, " & _												 
													 " id.fcBA, " & _
													 " id.STOCKCODE "
						Set rsCursorDebit = Server.CreateObject("ADODB.Recordset")
						rsCursorDebit.CursorType = 3
						rsCursorDebit.CursorLocation = 3
						rsCursorDebit.Open strCursorDebit, connstr		
						
						tempRowNum = 1

						if not rsCursorDebit.eof then
							Do While Not rsCursorDebit.eof
								'		   If totalRowInserted = n  then 
								'			   tempTotalAmountDebit = tempTotalAmountDebit + dari Cursor 2 setelah di round
								'			   tempSelisih = tempTotalAmountCredit - tempTotalAmountDebit				   
								'			   Insert Debit ==> Amount setelah di Round + tempSelisih
								'           else
								'		       Insert Debit
								'			   tempTotalAmountDebit = tempTotalAmountDebit + dari Cursor 2 setelah di round
								'			   totalRowInserted = totalRowInserted + 1
								' 		   end if
								
								
								tempDebit = round(cdbl(rsCursorDebit("amount")),2)
								
								if totalRowInserted = totalRecord - 1 then 
									tempTotalAmountDebit = tempTotalAmountDebit + tempDebit
									
									if tempTotalAmountDebit < (-1 * tempTotalAmountCredit) then
										tempSelisih = (-1 * tempTotalAmountCredit) - tempTotalAmountDebit				   
										tempDebit = tempDebit + tempSelisih
									else
										tempSelisih = tempTotalAmountDebit - (-1 * tempTotalAmountCredit) 				   
										tempDebit = tempDebit - tempSelisih								
									end if
									Response.Flush()
									maxLineNo = maxLineNo + 1
									qInsertDebit = " INSERT INTO cost_activity ( LINENO ," & _
																			  " TDATE, " & _
																			  " JOBCODE, " & _
																			  " SOURCE, " & _
																			  " SOURCETYPE, " & _
																			  " SOURCECODE, " & _
																			  " SOURCECODE_STATUS, " & _
																			  " TARGETTYPE, " & _
																			  " TARGETCODE, " & _
																			  " TARGETCODE_STATUS, " & _
																			  " REFERENCENO, " & _
																			  " QUANTITY, " & _
																			  " OUTPUT, " & _
																			  " AMOUNT, " & _
																			  " FCBA, " & _
																			  " CLOSING_DATE, " & _ 
																			  " PROCESS_DATE, " & _		
																			  " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																			  " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																			  " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																			  " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																			  " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _		
																			  " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
																			  " LASTUPDATE,  " & _
																			  " LASTTIME ) " & _
															" VALUES (" & _
															maxLineNo & "," & _
															" to_date('" & rsCursorDebit("tdate") & "','mm/dd/yyyy'),'" & _
															rsCursorDebit("JOBCODE") & "','" & _
															rsCursorDebit("SOURCE") & "','" & _
															rsCursorDebit("SOURCETYPE") & "','" & _
															rsCursorDebit("SOURCECODE") & "','" & _
															rsCursorDebit("SOURCECODE_STATUS") & "','" & _
															rsCursorDebit("TARGETTYPE") & "','" & _
															rsCursorDebit("TARGETCODE") & "','" & _
															rsCursorDebit("TARGETCODE_STATUS") & "','" & _
															rsCursorDebit("REFERENCENO") & "'," & _
															rsCursorDebit("QUANTITY") & "," & _
															rsCursorDebit("OUTPUT") & "," & _
															tempDebit & ",'" & _
															rsCursorDebit("FCBA") & "'," & _
															" to_date('" & rsCursorDebit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
															" sysdate, " & _
															" Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & rsCursorDebit("TARGETCODE") & "'," & _
															" Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempSIVCODE & "'," & _
															" Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempStockCode & "'," & _
															" '',''," & _
															" '','','Dr'," & _
															Month(Session("v_Stores")) & ", " & _
															Year(Session("v_Stores")) & ", " & _
															" sysdate, " & _
															" to_char(sysdate,'HH24:MI')) "
									response.Flush()		  
									'response.write qInsertDebit
									connstr.Execute(qInsertDebit)			
		'* Remarks By Iwa * 
									' strCursorDebitDetail = " SELECT isu.SIVDATE tdate, " & _
																 ' " id.JOBCODE jobcode,  " & _
																 ' " 'ST' source,  " & _
																 ' " 'ST' sourcetype,  " & _
																 ' " ISU.STORECODE sourcecode,  " & _
																 ' " '' sourcecode_status,  " & _
																 ' " id.LOCATIONTYPE targettype,  " & _
																 ' " id.LOCATIONCODE targetcode,  " & _
																 ' " '' targetcode_status,  " & _
																 ' " id.SIVCODE referenceno,  " & _
																 ' " id.linenokey, " & _
																 ' " id.locationtype, " & _
																 ' " id.locationcode, " & _
																 ' " id.jobcode, " & _												 
																 ' " id.QUANTITY quantity,  " & _
																 ' " 0 output,  " & _
																 ' " round(id.AMOUNT * 1) amount,  " & _
																 ' " id.FCBA fcba,  " & _
																 ' " TO_DATE ('" & temp_datePeriodControl & "', 'mm/dd/yyyy') closing_date,  " & _
																 ' " SYSDATE process_date, " & _
																 ' " ID.STOCKCODE " & _
															' " FROM ISSUE ISU, ISSUEDETAILS ID  " & _
														   ' " WHERE	 (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA)  " & _
																 ' " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
																 ' " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy'))  =" & Year(Session("v_Stores"))  & _
																 ' " AND ISU.fcBA = '" & Session("fcBA") & "'" & _
																 ' " AND ISU.sivcode = '" & tempSIVCODE & "'  " & _
																 ' " AND ID.STOCKCODE = '" & tempStockCode & "' " & _
																 ' " AND ID.LINENOKEY = '" & rsCursorCredit ("linenokey") & "' " & _
																 ' " AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
																 ' " AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
																 ' " AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") & "' " 											 
							
							
									' Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
									' rsReceiveDetail.CursorType = 3
									' rsReceiveDetail.CursorLocation = 3
									' rsReceiveDetail.Open strCursorDebitDetail, connstr
									' if not rsReceiveDetail.eof then
										' Do While Not rsReceiveDetail.eof
										
											' maxLineNoDetail = maxLineNoDetail + 1
											' qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
															' "( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
															' "  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
															' " QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
															' maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
															' txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
															' cdbl(rsReceiveDetail("QUANTITY")) & ",0,"& cdbl(rsReceiveDetail("AMOUNT")) &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
															' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
											' response.write qCost_ActivityDBHistory & "</br>"
											
											' connstr.execute(qCost_ActivityDBHistory)		
										    'response.write qCost_ActivityDBHistory & " && "
											' rsReceiveDetail.MoveNext
										' loop
									' end if
									' rsReceiveDetail.Close
									' set rsReceiveDetail = Nothing							
							
								'* Remarks By Iwa * 
								
								else
									Response.Flush()
									maxLineNo = maxLineNo + 1
									qInsertDebit = " INSERT INTO cost_activity ( LINENO ," & _
																			  " TDATE, " & _
																			  " JOBCODE, " & _
																			  " SOURCE, " & _
																			  " SOURCETYPE, " & _
																			  " SOURCECODE, " & _
																			  " SOURCECODE_STATUS, " & _
																			  " TARGETTYPE, " & _
																			  " TARGETCODE, " & _
																			  " TARGETCODE_STATUS, " & _
																			  " REFERENCENO, " & _
																			  " QUANTITY, " & _
																			  " OUTPUT, " & _
																			  " AMOUNT, " & _
																			  " FCBA, " & _
																			  " CLOSING_DATE, " & _ 
																			  " PROCESS_DATE, " & _		
																			  " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																			  " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																			  " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																			  " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																			  " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _																				  
																			  " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _	
																			  " LASTUPDATE,  " & _
																			  " LASTTIME ) " & _
															" VALUES (" & _
															maxLineNo & "," & _
															" to_date('" & rsCursorDebit("tdate") & "','mm/dd/yyyy'),'" & _
															rsCursorDebit("JOBCODE") & "','" & _
															rsCursorDebit("SOURCE") & "','" & _
															rsCursorDebit("SOURCETYPE") & "','" & _
															rsCursorDebit("SOURCECODE") & "','" & _
															rsCursorCredit("SOURCECODE_STATUS") & "','" & _
															rsCursorDebit("TARGETTYPE") & "','" & _
															rsCursorDebit("TARGETCODE") & "','" & _
															rsCursorDebit("TARGETCODE_STATUS") & "','" & _
															rsCursorDebit("REFERENCENO") & "'," & _
															rsCursorDebit("QUANTITY") & "," & _
															rsCursorDebit("OUTPUT") & "," & _
															tempDebit & ",'" & _
															rsCursorDebit("FCBA") & "'," & _
															" to_date('" & rsCursorDebit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
															" sysdate, " & _
															" Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & rsCursorDebit("TARGETCODE") & "'," & _
															" Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempSIVCODE & "'," & _
															" Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempStockCode & "'," & _
															" '',''," & _
															" '','','Dr'," & _
															Month(Session("v_Stores")) & ", " & _
															Year(Session("v_Stores")) & ", " & _
															" sysdate, " & _
															" to_char(sysdate,'HH24:MI')) "
									response.Flush()		  
									connstr.Execute(qInsertDebit)	
'* Remarks By Iwa * 
									' strCursorDebitDetail = " SELECT isu.SIVDATE tdate, " & _
																 ' " id.JOBCODE jobcode,  " & _
																 ' " 'ST' source,  " & _
																 ' " 'ST' sourcetype,  " & _
																 ' " ISU.STORECODE sourcecode,  " & _
																 ' " '' sourcecode_status,  " & _
																 ' " id.LOCATIONTYPE targettype,  " & _
																 ' " id.LOCATIONCODE targetcode,  " & _
																 ' " '' targetcode_status,  " & _
																 ' " id.SIVCODE referenceno,  " & _
																 ' " id.linenokey, " & _
																 ' " id.locationtype, " & _
																 ' " id.locationcode, " & _
																 ' " id.jobcode, " & _												 
																 ' " id.QUANTITY quantity,  " & _
																 ' " 0 output,  " & _
																 ' " round(id.AMOUNT * 1) amount,  " & _
																 ' " id.FCBA fcba,  " & _
																 ' " TO_DATE ('" & temp_datePeriodControl & "', 'mm/dd/yyyy') closing_date,  " & _
																 ' " SYSDATE process_date, " & _
																 ' " ID.STOCKCODE " & _
															' " FROM ISSUE ISU, ISSUEDETAILS ID  " & _
														   ' " WHERE	 (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA)  " & _
																 ' " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
																 ' " AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy'))  =" & Year(Session("v_Stores"))  & _
																 ' " AND ISU.fcBA = '" & Session("fcBA") & "'" & _
																 ' " AND ISU.sivcode = '" & tempSIVCODE & "'  " & _
																 ' " AND ID.STOCKCODE = '" & tempStockCode & "' " & _
																 ' " AND ID.LINENOKEY = '" & rsCursorCredit ("linenokey") & "' " & _
																 ' " AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
																 ' " AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
																 ' " AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") & "' " 												 
							
							
									' Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
									' rsReceiveDetail.CursorType = 3
									' rsReceiveDetail.CursorLocation = 3
									' rsReceiveDetail.Open strCursorDebitDetail, connstr
									' if not rsReceiveDetail.eof then
										' Do While Not rsReceiveDetail.eof
										
											' maxLineNoDetail = maxLineNoDetail + 1
											' qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
															' "( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
															' "  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
															' " QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
															' maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodeReceive & "','" & txtRefNumberKeys_Receive & "','" & _
															' txtRefStringSeparator_Receive & "','" & txtRefStringKeys_Receive & "','" &  txtRefStringValues_Receive & "' , " & _
															' cdbl(rsReceiveDetail("QUANTITY")) & ",0,"& cdbl(rsReceiveDetail("AMOUNT")) &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
															' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
											' response.write qCost_ActivityDBHistory & "</br>"
											' connstr.execute(qCost_ActivityDBHistory)		
										
											' rsReceiveDetail.MoveNext
										' loop
									' end if
									' rsReceiveDetail.Close
									' set rsReceiveDetail = Nothing							
'* Remarks By Iwa * 
								
									tempTotalAmountDebit = tempTotalAmountDebit + tempDebit
									totalRowInserted = totalRowInserted + 1
								end if	
								tempRowNum = tempRowNum + 1
								rsCursorDebit.MOVENEXT
							loop
							rsCursorDebit.close
							set rsCursorDebit=Nothing				
						end if
						rsCursorCredit.MOVENEXT
					loop
					rsCursorCredit.close
					set rsCursorCredit=Nothing		
				end if
				rsCursorSivCode.MOVENEXT
			loop
			rsCursorSivCode.close
			set rsCursorSivCode=Nothing
		End if
		Response.Write "<br><font color=Blue>" & "Proses Cost Activity for Issue Details is completed..." & "</font>"	
		
		

		
		' Remarked by NAF, as per we change the Old Process with the New ONE
		' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
		' Response.Flush()
		' qABC =    " INSERT INTO COST_ACTIVITY(" & _
				  ' " SELECT isu.SIVDATE, stockgroup.CONTROLJOB, 'ST', 'ST', ISU.STORECODE, '', 'CB', 'BS', '', id.SIVCODE, SUM(id.QUANTITY), 0, SUM(id.AMOUNT)*-1, id.FCBA , to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate " & _
				  ' " FROM ISSUE ISU,ISSUEDETAILS ID,Stock,STOCKGROUP " & _
				  ' " WHERE (ISU.SIVCODE=ID.SIVCODE AND ISU.FCBA=ID.FCBA) " & _
				  ' " and (id.STOCKCODE =stock.FCCODE and id.fcba = stock.fcba) " & _
				  ' " and (stock.STOCKGROUPCODE =stockgroup.FCCODE and stock.fcba = stockgroup.fcba) " & _
				  ' " and to_number(to_Char(ISU.SIVDATE,'mm'))=" & Month(Session("v_Stores")) &  _
				  ' " and to_number(to_Char(ISU.SIVDATE,'yyyy'))=" & Year(Session("v_Stores")) & " and ISU.fcBA='" & Session("fcBA") & "'" & _
				  ' " GROUP BY isu.SIVDATE, stockgroup.CONTROLJOB, 'ST', 'ST', ISU.STORECODE, '', 'CB', 'BS', '', id.SIVCODE,id.fcBA)"
		' response.Write " qABC "& qABC
		' response.Flush()		  
		' connstr.Execute(qABC)
		

		' '----- Tranfer to Activity Costing
		' Response.Flush()
		' qABC =  " INSERT INTO COST_ACTIVITY(" & _
				' " SELECT isu.SIVDATE, id.JOBCODE, 'ST', 'ST', ISU.STORECODE, '', id.LOCATIONTYPE, id.LOCATIONCODE, '', id.SIVCODE, SUM(id.QUANTITY), 0, SUM(id.AMOUNT)*1, id.FCBA , to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate " & _
				' " FROM ISSUE ISU,ISSUEDETAILS ID WHERE " & _
				' " (ISU.SIVCODE=ID.SIVCODE AND ISU.FCBA=ID.FCBA) AND " & _
				' " to_number(to_Char(ISU.SIVDATE,'mm'))=" & Month(Session("v_Stores")) &  _
				' " and to_number(to_Char(ISU.SIVDATE,'yyyy'))=" & Year(Session("v_Stores")) & " and ISU.fcBA='" & Session("fcBA") & "'" & _
				' " GROUP BY isu.SIVDATE, id.JOBCODE, 'ST', 'ST', ISU.STORECODE, '', id.LOCATIONTYPE, id.LOCATIONCODE, '', id.SIVCODE,id.fcBA)"
		' response.Write " qABC "& qABC
		' response.Flush()		  
		' connstr.Execute(qABC)
		' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
		
		
		
		Response.Write "<br><font color=Blue>" & "(Issue Details) Transfer to Activity Based Costing is completed..." & "</font>"
		
		qautonumber = " SELECT TO_CHAR(ADD_MONTHS(TO_DATE('" & Session("v_Stores") & "', 'mm/dd/yyyy'), 1),'yyyymm') xdate FROM DUAL"
		Set rsautonumber = Server.CreateObject("ADODB.Recordset")
		rsautonumber.Open qautonumber, connstr
		if not rsautonumber.eof then
			tempDateYear = MID(rsautonumber("xdate"), 3, 2)
			tempDateMonth = RIGHT(rsautonumber("xdate"), 2)
		end if
		rsautonumber.close : set rsautonumber = nothing
        
		' Format RECEIVEDTASK No adalah RCVFCBAYYMMXXXX
		tempPrefix = "RCV" &  Session("fcBA") & tempDateYear & "" & tempDateMonth 
		tempSuffix = ""
		tempLengthPad = 4
		q = "begin ControlNumber_Api.Generate_Data ('RECEIVEDTASK',add_months(to_date('" & Session("v_Stores") & "','MM/DD/YYYY'),1),'" & _ 
		                                            tempPrefix & "','" & tempSuffix & "'," & _
													tempLengthPad & ",'" &  Session("fcBA") & "','" & Session("CodeName") & "','" & _
													Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "'); end;" 
	    connStr.execute(q)		
		
		' Add By Akbar 18 Desemeber 2015
		
			Dim temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear  " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='STORES' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for STORES " & "<br>"
				end if
			rsAccountPeriod.close
			set rsAccountPeriod = Nothing		
			
			IF temp_NextCurrentAccountingPeriod < 12   THEN
			
				q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)),CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _    
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where UPPER(system) ='STORES' and fcba='" & Session("fcBA") & "'"
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					    " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _    
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where UPPER(system) ='STORES' and fcba='" & Session("fcBA") & "'"
				ELSE
				    IF temp_NextCloseAccountingPeriod = 12 THEN
						q=  " Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
							" ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							" ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							" ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							" ,CurrentAccountingYear= " & temp_NextCurrentAccountingYear & " " & _  
							" ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							" where  UPPER(system) ='STORES' and fcba='" & Session("fcBA") & "'"
					ELSE
					    q=  " Update PeriodControl Set" & _ 
							" CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							" ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							" ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							" ,CurrentAccountingYear= " & temp_NextCurrentAccountingYear & " " & _  
							" ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							" where  UPPER(system) ='STORES' and fcba='" & Session("fcBA") & "'"
					END IF
			    END IF
			END IF
			
			IF temp_NextCloseAccountingPeriod = 13 AND temp_NextCurrentAccountingPeriod = 1  THEN
			
				q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where  UPPER(system) ='STORES' and fcba='" & Session("fcBA") & "'"
				ww = " eMPAT " 
			END IF
			
		connStr.execute(q)
		

	TimeAkhir=Time()
	Lama =TimeAkhir-TimeAwal	
	%>

	<%if err.number <> 0 then%>
		<br><img src="images/error.gif"><%
		'connStr.RollbackTrans
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		cmdView	= "True"%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>	
	<%Else
	   ' connStr.CommitTrans
		response.Clear()
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "<br><br>" & "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."

		'--Added by Ukie at 25-Oct-2012--	
		Endingtime = FormatDateTime(CDate(Now))

		qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Store'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME) VALUES (" & _
						" 'Store', " & _
						" '" & Session("fcBA") & "'," & _
						" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" " & tempMaxLine & _
						" , SYSDATE, to_char(SYSDATE,'HH24:MI') " & _
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
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Store'"
				
		Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
		rsLenghtime.CursorType = 3
		rsLenghtime.CursorLocation = 3
		rsLenghtime.Open strLenghtime, connstr

		upClosingLog = " UPDATE CLOSINGLOG SET" & _
						" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
						", LASTUPDATE = Sysdate, LastTime = To_Char(Sysdate, 'HH24:MI') " & _
						" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Store'"		
		connStr.execute(upClosingLog)
		'--End Added--
		
		%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Store Closing is succesfully...
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		
		<%	
			txtLastDay  =cmdFindPeriod("Stores")
			txtNextMonth=cmdPeriodClose("Stores")
			
			
			Session("v_Stores")=txtLastDay
			Session("cls_Stores")= txtNextMonth	
		%>	
		
		
		
	<%End if%>
<%
else 
%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/error.gif">Store Closing is NOT completed...
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