<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>iPla'S - Bank Closing</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()

'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Dim fine_result_01, fine_result_02, fine_result_03

Startingtime = FormatDateTime(CDate(Now))
'--End Added--	
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Bank Last Period Close <%=FormatDateTime(Session("cls_Bank"),1)%> <br> 
Bank Current Period is <%=FormatDateTime(Session("v_Bank"),1)%> <br><br>
Please Wait...<br>
<%
 Response.Write "Start ..." & Time() & "<br>" 
dim rsVoucher,rsBank,qVoucher,qPO,nPurchaseControl,txtControlJob,qCost_ActivityDB,qCost_ActivityCR,nQtyReceived,qNormalisasi,txtOriginalST

' New Concept Variable, By NAF, 12 May 2013
Dim tempProcessCode, tempProcessNameCode, tempRefCodePVDetail, qReference
Dim rsReferencePVDetail
Dim txtRefNumberKeys_PVDetail, txtRefStringSeparator_PVDetail , txtRefStringKeys_PVDetail
Dim txtRefStringValues_PVDetail
Dim donePVDetail

' End Modification.


' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Bank','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Bank " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 

	'---- Normalisasi

	qNormalisasiDetail = " delete from cost_activity_hist_tab " & _
						 " where " & _
						 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
								" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Bank")) &  _
								" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Bank")) & " and fcBA='" & Session("fcBA") & "' and Source='BA')"
						 
	connstr.execute(qNormalisasiDetail)
    connStr.execute("Commit")

	qNormalisasi="Delete From Cost_Activity where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Bank")) &  _
				 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Bank")) & " and fcBA='" & Session("fcBA") & "' and SourceType='BA'"
	connstr.execute(qNormalisasi)
    connStr.execute("Commit")
	' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	' Add by NAF to select lastperiodposted
	Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl

	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='Bank' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Bank " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	' Remarked by NAF, as per New Concept, 7 May 2013
	' qControlJob="select ControlJob from Module where NAME='Purchase'"
	' Now we use new concept for picking Control Job, by using Process Control Tables
	' as per last concept are NOT able to pick data per Business Unit.

	tempProcessCode = "BANK_CLOSING"
	tempProcessNameCode = "p_Bank_Closing.asp"
	tempRefCodePVDetail = "PAYMENTVOUCHERDETAIL"

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



	qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
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

	' End of Modification


	qVoucher="select p.VOUCHERCODE,p.BANKCODE,p.DATECREATED," & _
			" pv.SOURCETYPE,pv.LOCATIONTYPE,pv.LOCATIONCODE,pv.JOBCODE,pv.REFERENCE,sum(pv.AMOUNT) as Amount" & _
			" from PAYMENTVOUCHER P,PAYMENTVOUCHERDETAIL PV" & _
			" where (p.VOUCHERCODE = pv.VOUCHERCODE and p.fcba = pv.fcba)" & _
			" AND to_number(to_Char(p.DATECREATED,'mm'))=" & Month(Session("v_Bank")) &  _
			" and to_number(to_Char(p.DATECREATED,'yyyy'))=" & Year(Session("v_Bank")) & _
			" and pv.fcBA='" & Session("fcBA") & "'" & _
			" group by p.VOUCHERCODE,p.BANKCODE,p.DATECREATED," & _
			" pv.SOURCETYPE,pv.LOCATIONTYPE,pv.LOCATIONCODE,pv.JOBCODE,pv.REFERENCE"
			
	Set rsVoucher = Server.CreateObject("ADODB.Recordset")
	rsVoucher.CursorType = 3
	rsVoucher.CursorLocation = 3
	rsVoucher.Open qVoucher, connstr
	if not rsVoucher.eof and donePVDetail = 1 then
	   Do While Not rsVoucher.eof
		  '-- Control Job Bank     
		  txtControlJob=""
		  qPO=" select CONTROLJOB as CONTROLJOB FROM Bank" & _
			  " WHERE FCCode ='" & rsVoucher("BANKCODE")  & "'" & _
			  " and fcBA='" & Session("fcBA") & "'"
		  Set rsBank = Server.CreateObject("ADODB.Recordset")
		  rsBank.CursorType = 3
		  rsBank.CursorLocation = 3
		  rsBank.Open qPO, connstr
		  if not rsBank.eof then
			 txtControlJob = rsBank("CONTROLJOB") 

			 '---- Cek Source Type Original
			 if rsVoucher("SOURCETYPE")<>"" then
				txtOriginalST= rsVoucher("SOURCETYPE")
			 else
				txtOriginalST= "ST"
			 end if 
			 
			 ' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
			 txtRefStringValues_PVDetail = Session("fcBA") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("VOUCHERCODE") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("LOCATIONTYPE") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("JOBCODE") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("LOCATIONCODE") 
										   

			 maxLineNo = maxLineNo + 1
			 qCost_ActivityDB="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsVoucher("DATECREATED") & "','mm/dd/yyyy') " & ",'" & rsVoucher("JOBCODE") & "','" & txtOriginalST & "','" & "BA" & "','" & rsVoucher("BANKCODE") & "'," & _
							 "'" &  rsVoucher("BANKCODE") & "','" & rsVoucher("LOCATIONTYPE") & "','" & rsVoucher("LOCATIONCODE") & "'," & _
							 "'" &  rsVoucher("Amount") & "','" & rsVoucher("VOUCHERCODE") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
							 
			 connstr.execute(qCost_ActivityDB)
			

			' Insert History Transaction,
			
			qVoucherDetail ="select p.VOUCHERCODE,p.BANKCODE,p.DATECREATED," & _
							" pv.SOURCETYPE,pv.LOCATIONTYPE,pv.LOCATIONCODE,pv.JOBCODE,pv.REFERENCE,pv.AMOUNT as Amount" & _
							" from PAYMENTVOUCHER P,PAYMENTVOUCHERDETAIL PV" & _
							" where (p.VOUCHERCODE = pv.VOUCHERCODE and p.fcba = pv.fcba)" & _
							" AND to_number(to_Char(p.DATECREATED,'mm'))=" & Month(Session("v_Bank")) &  _
							" and to_number(to_Char(p.DATECREATED,'yyyy'))=" & Year(Session("v_Bank")) & _
							" and pv.fcBA='" & Session("fcBA") & "'" & _
							" and pv.VOUCHERCODE = '" & rsVoucher("VOUCHERCODE") & "' and " & _
							" pv.LOCATIONTYPE = '" & rsVoucher("LOCATIONTYPE") & "' and " & _
							" pv.JOBCODE = '" & rsVoucher("JOBCODE") & "' and " & _
							" pv.LOCATIONCODE = '" & rsVoucher("LOCATIONCODE")  & "'"
			
			Set rsVoucherDetail = Server.CreateObject("ADODB.Recordset")
			rsVoucherDetail.CursorType = 3
			rsVoucherDetail.CursorLocation = 3
			rsVoucherDetail.Open qVoucherDetail, connstr
			
			if not rsVoucherDetail.eof then
				maxLineNoDetail = 0
				Do While Not rsVoucherDetail.eof

					maxLineNoDetail = maxLineNoDetail + 1
					qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
									"( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
									"  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
									" QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
									maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodePVDetail & "','" & txtRefNumberKeys_PVDetail & "','" & _
									txtRefStringSeparator_PVDetail & "','" & txtRefStringKeys_PVDetail & "','" &  txtRefStringValues_PVDetail & "' , " & _
									"0,0,"& rsVoucherDetail("amount")&",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
								 
					connstr.execute(qCost_ActivityDBHistory)		
					 connStr.execute("Commit")
					rsVoucherDetail.MoveNext
				loop
			end if
			rsVoucherDetail.Close
			set rsVoucherDetail = Nothing	
			
			nPurchaseControl=-1*cdbl(rsVoucher("Amount"))
			maxLineNo = maxLineNo + 1

			qCost_ActivityCR="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA , closing_date, process_date, " & _
							 " LASTUPDATE, LASTTIME) " & _
							 " values (" & maxLineNo & ", to_date('" & rsVoucher("DATECREATED") & "','mm/dd/yyyy') " & ",'" & txtControlJob & "','" & txtOriginalST & "','" & "BA" & "','" & rsVoucher("BANKCODE") & "'," & _
							 "'" &  rsVoucher("BANKCODE") & "','" & "CB" & "','" & "BS" & "'," & _
							 "'" & nPurchaseControl  & "','" & rsVoucher("VOUCHERCODE") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
							 
							 
			 connstr.execute(qCost_ActivityCR)
			 connStr.execute("Commit")
			' Insert History Transaction,
			
			qVoucherDetail ="select p.VOUCHERCODE,p.BANKCODE,p.DATECREATED," & _
							" pv.SOURCETYPE,pv.LOCATIONTYPE,pv.LOCATIONCODE,pv.JOBCODE,pv.REFERENCE,pv.AMOUNT as Amount" & _
							" from PAYMENTVOUCHER P,PAYMENTVOUCHERDETAIL PV" & _
							" where (p.VOUCHERCODE = pv.VOUCHERCODE and p.fcba = pv.fcba)" & _
							" AND to_number(to_Char(p.DATECREATED,'mm'))=" & Month(Session("v_Bank")) &  _
							" and to_number(to_Char(p.DATECREATED,'yyyy'))=" & Year(Session("v_Bank")) & _
							" and pv.fcBA='" & Session("fcBA") & "'" & _
							" and pv.VOUCHERCODE = '" & rsVoucher("VOUCHERCODE") & "' and " & _
							" pv.LOCATIONTYPE = '" & rsVoucher("LOCATIONTYPE") & "' and " & _
							" pv.JOBCODE = '" & rsVoucher("JOBCODE") & "' and " & _
							" pv.LOCATIONCODE = '" & rsVoucher("LOCATIONCODE")  & "'"
			
			Set rsVoucherDetail = Server.CreateObject("ADODB.Recordset")
			rsVoucherDetail.CursorType = 3
			rsVoucherDetail.CursorLocation = 3
			rsVoucherDetail.Open qVoucherDetail, connstr
			
			if not rsVoucherDetail.eof and donePVDetail = 1 then
				maxLineNoDetail = 0
				Do While Not rsVoucherDetail.eof

					maxLineNoDetail = maxLineNoDetail + 1
					nPurchaseControl = -1*cdbl(rsVoucherDetail("Amount"))
					qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
									"( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
									"  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
									" QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
									maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodePVDetail & "','" & txtRefNumberKeys_PVDetail & "','" & _
									txtRefStringSeparator_PVDetail & "','" & txtRefStringKeys_PVDetail & "','" &  txtRefStringValues_PVDetail & "' , " & _
									"0,0,"& nPurchaseControl &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
								 
					connstr.execute(qCost_ActivityDBHistory)		
					 connStr.execute("Commit")
					rsVoucherDetail.MoveNext
				loop
			end if
			rsVoucherDetail.Close
			set rsVoucherDetail = Nothing		
			
			response.Write "Process Bank... " & rsVoucher("VOUCHERCODE") & "<br>"
			response.Flush()
		  else
			 Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20011 Control Job For Bank " & rsVoucher("BANKCODE") & " is Empty" & "<br>"
		  End if
		  rsBank.Close
		  set rsBank = Nothing
		  
		  rsVoucher.MoveNext
	  loop
	End if





	TimeAkhir=Time()
	Lama =TimeAkhir-TimeAwal
	Response.Write "Finish ... " & time() & "<br><br>"
	Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."

	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Bank'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

	strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME) VALUES (" & _
					" 'Bank', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy HH:MI:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy HH:MI:ssam')," & _
					" " & tempMaxLine & ", sysdate, to_char(sysdate,'HH24:MI')" & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Bank'"
				
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.CursorType = 3
	rsLenghtime.CursorLocation = 3
	rsLenghtime.Open strLenghtime, connstr

	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Bank'"		
	connStr.execute(upClosingLog)
	'--End Added--

	%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/ok.gif">Bank Closing is completed...
	<%
	 if err.number <> 0 then
		connStr.RollbackTrans
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		cmdView	= "True"%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		<img src="images/error.gif">
		 <% fine_result_01 = 0 %>	
	<% else %>
		 <% fine_result_01 = 1 %>
	<% End if%>
	 
	 <%
	 if donePVDetail <> 1 then 
		connStr.RollbackTrans
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>Payment Voucher Detail String still not defined in PROCESS_TABLEREF_TAB, Call your Administrator </font>"
		cmdView	= "True"%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		<img src="images/error.gif" />
		 <% fine_result_02 = 0 %>	
	<% else %>
		 <% fine_result_02 = 1 %>	
	<% end if %>
	 
	<% 
		if fine_result_01 = 1 and fine_result_02 = 1 then 
		
			q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)),CURRENTDATE=add_months(CURRENTDATE,1) " & _
			  " where SYSTEM='Bank' and fcba='" & Session("fcBA") & "'"

			connStr.execute(q)

			' Query ke Period Control untuk mengambli Session(v_bank) & Session(cls_Bank)	
		
		
			txtLastDay  =cmdFindPeriod("Bank")
			txtNextMonth=cmdPeriodClose("Bank")

			Session("v_Bank")= txtLastDay
			Session("cls_Bank")= txtNextMonth
		end if
	 
	%> 
<%
else 
%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/error.gif">Bank Closing is NOT completed...
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