<%
'Option Explicit
%>
<%
'On Error Resume Next
%>
<%response.buffer=True%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Costbook Closing</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()

'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))
'--End Added--	
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Financial Last Period Close <%=FormatDateTime(Session("cls_Costbook"),1)%> <br> 
Financial Current Period is <%=FormatDateTime(Session("v_Costbook"),1)%> <br><br>
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br>" 

dim rsCek,RsCost_ActivityDb,RsCost_ActivityCr,RsCost_Opening
dim v_AllowanceCode



' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Costbook','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Costbook " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 

	'---- Control Job
	qControlJob="select ControlJob from Module where NAME='Costbook'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC30025 Control Job For Costbook is Empty"
	   err.number=30025
	   txtControlJob=""
	end if
	response.Flush()
	
	' Add By Akbar 15 December 2015  For Accounting Period 13,14 and 15
	
	 Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear, temp_EndingAccountingPeriod
		
	
	 QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
							" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
							" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
							" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear  " & _ 
							" from periodcontrol p  " & _ 
							" where UPPER(system) ='COSTBOOK' and fcba='" & Session("fcBA") & "'"
	Set rsAccountPeriod = Server.CreateObject("ADODB.Recordset")
		rsAccountPeriod.CursorType = 3
		rsAccountPeriod.CursorLocation = 3
		rsAccountPeriod.Open QAccountPeriod, connstr
		if not rsAccountPeriod.eof then
			temp_NextCloseAccountingPeriod =  rsAccountPeriod("NextCloseAccountingPeriod")
			temp_NextCloseAccountingYear   =  rsAccountPeriod("NextCloseAccountingYear")
			temp_NextCurrentAccountingPeriod = rsAccountPeriod("NextCurrentAccountingPeriod")
			temp_NextCurrentAccountingYear = rsAccountPeriod("NextCurrentAccountingYear")
		else
			Response.Write  "<font face=Verdana size=1 color=red>" & "IPLAS Error, No PERIODCONTROL for COSTBOOK " & "<br>"
		end if
		rsAccountPeriod.close
		set rsAccountPeriod = Nothing	

	if txtControlJob<>"" then
		'---- Normalisasi
		qNormalisasi=" Delete From TrialBalance where  Period_Month=" & temp_NextCurrentAccountingPeriod &  _
					 " and Period_Year=" & temp_NextCurrentAccountingYear & " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)	
		'response.write qNormalisasi

		Set rsJob = Server.CreateObject("ADODB.Recordset")
		rsJob.CursorType = 3
		rsJob.CursorLocation = 3
		qContract = " Select FCCODE as AccountCode,FCNAME from Job where fcBA='" & Session("fcBA") & "' order by FCCODE"
			  
		rsJob.Open qContract, connstr
		if not rsJob.eof then
		   Do While Not rsJob.eof
			  nEnding  = 0 
			  nOpening = 0
			  nDebet   = 0
			  nCredit  = 0
			  
			  Response.Write "to Generated Jurnal " & rsJob("AccountCode") & "	" & rsJob("FCNAME") & "<br>"
			  Response.Flush()

			  '--------------------- Debet
			  Set RsCost_ActivityDb = Server.CreateObject("ADODB.Recordset")
			  RsCost_ActivityDb.CursorType = 3
			  RsCost_ActivityDb.CursorLocation = 3
			  qCostContractDb = " Select nvl(Sum(AMOUNT),0) as nDebet" & _
						  " from Cost_Activity where JobCode='" & rsJob("AccountCode") & "' and Amount > 0" & _
						  " and AccountingPeriod = " & temp_NextCloseAccountingPeriod & " And AccountingYear =" & temp_NextCloseAccountingYear &  _
						  " and fcBA='" & Session("fcBA") & "'"
			  response.Flush()		  
			  RsCost_ActivityDb.Open qCostContractDb, connstr
			  if not RsCost_ActivityDb.eof then
				 nDebet=cdbl(RsCost_ActivityDb("nDebet"))
			  else
				 nDebet=0
			  End if   
			  RsCost_ActivityDb.Close
			  set RsCost_ActivityDb = Nothing
			  
			  '--------------------- Credit
			  Set RsCost_ActivityCr = Server.CreateObject("ADODB.Recordset")
			  RsCost_ActivityCr.CursorType = 3
			  RsCost_ActivityCr.CursorLocation = 3
			  qCostContractCr = " Select nvl(Sum(AMOUNT),0) as nCredit" & _
						  " from Cost_Activity where JobCode='" & rsJob("AccountCode") & "' and Amount < 0" & _
						  " And AccountingPeriod = " & temp_NextCloseAccountingPeriod & " And AccountingYear =" & temp_NextCloseAccountingYear &  _
						  " and fcBA='" & Session("fcBA") & "'"
			  response.Flush()		  
			  RsCost_ActivityCr.Open qCostContractCr, connstr
			  if not RsCost_ActivityCr.eof then
				 nCredit=cdbl(RsCost_ActivityCr("nCredit"))
			  else
				 nCredit=0
			  End if   
			  RsCost_ActivityCr.Close
			  set RsCost_ActivityCr = Nothing

			  '--------------------- Trial Balance
			  Set RsCost_Opening = Server.CreateObject("ADODB.Recordset")
			  RsCost_Opening.CursorType = 3
			  RsCost_Opening.CursorLocation = 3
			  qOpening = " Select nvl(Sum(OPENING_BALANCE),0) as nOpening" & _
								" from TRIALBALANCE where FCCODE='" & rsJob("AccountCode") & "'" & _
								" and Period_Month=" & temp_NextCloseAccountingPeriod &  _
								" and Period_Year=" & temp_NextCloseAccountingYear & " and fcBA='" & Session("fcBA") & "'"
		'	  response.write qOpening
			  response.Flush()		  
			  RsCost_Opening.Open qOpening, connstr
			  if not RsCost_Opening.eof then
				 nOpening=cdbl(RsCost_Opening("nOpening"))
			  else
				 nOpening=0
			  End if   
			  RsCost_Opening.Close
			  set RsCost_Opening = Nothing
			  
			  nEnding = nOpening + nDebet + nCredit

			  Set RsCek = Server.CreateObject("ADODB.Recordset")
			  RsCek.CursorType = 3
			  RsCek.CursorLocation = 3
			  qCek = " Select nvl(count(*),0) as JlhRecord" & _
								" from TRIALBALANCE where FCCODE='" & rsJob("AccountCode") & "'" & _
								" and Period_Month=" & temp_NextCloseAccountingPeriod &  _
								" and Period_Year=" & temp_NextCloseAccountingYear & " and fcBA='" & Session("fcBA") & "'"
			  response.Flush()		  
			  RsCek.Open qCek, connstr
			  if not RsCek.eof then
				  if cdbl(RsCek("JlhRecord")) = 0 then
					 qUpdate=" Insert into TrialBalance " & _
						  "(PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
						  " values " & _
						  "(" & temp_NextCloseAccountingPeriod & "," & temp_NextCloseAccountingYear & ",'" & rsJob("AccountCode") & "','"  & _
						  rsJob("FCNAME") & "',0," & nDebet & "," & nCredit & "," & nEnding  & ",'"  & Session("CodeName") & "','" & Session("CodeName") & "','" & _
						  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))" 
						  'response.Write qUpdate
						  response.Flush()
					  connstr.execute(qUpdate)	
				  else
					 qUpdate=" Update TrialBalance set Debet =" & nDebet & ",Credit=" & nCredit & ",ENDING_BALANCE=" & nEnding & _
						  " where  FCCODE ='" & rsJob("AccountCode") & "' and Period_Month=" & temp_NextCloseAccountingPeriod &  _
						  " and Period_Year=" & temp_NextCloseAccountingYear& " and fcBA='" & Session("fcBA") & "'"
					 connstr.execute(qUpdate)	
				  End if	 
			  else
				  qUpdate=" Insert into TrialBalance " & _
						  "(PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA)" & _
						  " values " & _
						  "(" & temp_NextCloseAccountingPeriod & "," & temp_NextCloseAccountingYear & ",'" & rsJob("AccountCode") & "','"  & _
						  rsJob("AccountCode") & "',0," & nDebet & "," & nCredit & "," & nEnding  & ",'"  & Session("CodeName") & "','" & Session("CodeName") & "','" & _
						  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')" 
				  connstr.execute(qUpdate)	
			  End if   
			  RsCek.Close
			  set RsCek = Nothing
			  
			  rsJob.MoveNext
		   loop '--------- Loop For rsJob
		end if
		rsJob.close
		set rsJob = Nothing
		
		'Insert Opening Balance For Next Month Financial
		' ttemp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear

		qUpdate=" Insert into TrialBalance " & _
				" (PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				"  (select " &  temp_NextCurrentAccountingPeriod & " as PERIOD_MONTH,"  & temp_NextCurrentAccountingYear & " as PERIOD_YEAR" & ", FCCODE," & _
				"  FCNAME, OPENING_BALANCE+DEBET+CREDIT as OPENING_BALANCE,0 as DEBET, 0 as CREDIT,  0 as ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, SYSDATE, TO_CHAR(SYSDATE,'HH24:MI') " & _
				"  From TrialBalance Where Period_Month=" & temp_NextCloseAccountingPeriod &  _
				" and Period_Year=" & temp_NextCloseAccountingYear & " and fcBA='" & Session("fcBA") & "')"		 
		'Response.Write qUpdate
		response.Flush()
		connstr.execute(qUpdate)
		
		' Add By Akbar 18 Desemeber 2015
		
    ' AMH Apr 21 Pindahkan kesini 
	'**** Eksekusi ERP Report ****'
	Server.Execute("p_Generate_ERP_Report.asp")
	
	
	'**** Eksekusi AR Report ****'
	Server.Execute("p_Generate_Account_Report.asp")
        
		
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear,  " & _ 
								" periodcontrol_api.GetEndingAccountingPeriodTY(system,fcba)  EndingAccountingPeriod " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='COSTBOOK' and fcba='" & Session("fcBA") & "'"
			Set rsAccountPeriod = Server.CreateObject("ADODB.Recordset")
				rsAccountPeriod.CursorType = 3
				rsAccountPeriod.CursorLocation = 3
				rsAccountPeriod.Open QAccountPeriod, connstr
				if not rsAccountPeriod.eof then
					temp_NextCloseAccountingPeriod =  Cint(rsAccountPeriod("NextCloseAccountingPeriod"))
					temp_NextCloseAccountingYear   =  Cint(rsAccountPeriod("NextCloseAccountingYear"))
					temp_NextCurrentAccountingPeriod = Cint(rsAccountPeriod("NextCurrentAccountingPeriod"))
					temp_NextCurrentAccountingYear = Cint(rsAccountPeriod("NextCurrentAccountingYear"))
					temp_EndingAccountingPeriod = Cint(rsAccountPeriod("EndingAccountingPeriod"))
				else
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for COSTBOOK " & "<br>"
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
				   " where UPPER(system) ='COSTBOOK' and fcba='" & Session("fcBA") & "'"
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				    q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " "  & _  
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where UPPER(system) ='COSTBOOK' and fcba='" & Session("fcBA") & "'"
				ELSE
				     IF temp_NextCloseAccountingPeriod = 12 THEN
						 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
							" ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							" ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							" ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							" ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
							" ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							" where  UPPER(system) ='COSTBOOK' and fcba='" & Session("fcBA") & "'"
					 ELSE
                         q=" Update PeriodControl Set " & _ 
							" CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							" ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							" ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							" ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
							" ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							" where  UPPER(system) ='COSTBOOK' and fcba='" & Session("fcBA") & "'"
                     END IF					 
			    END IF
			END IF
			
			IF temp_NextCloseAccountingPeriod = temp_EndingAccountingPeriod AND temp_NextCurrentAccountingPeriod = 1  THEN
			
				q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
				  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
				  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
				  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
				  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
				  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
				  " where UPPER(system) ='COSTBOOK' and  fcba='" & Session("fcBA") & "'"
			END IF
			
			
           ' response.write q			
			
			connStr.execute(q)
            
			txtLastDay  =cmdFindPeriod("Costbook")
			txtNextMonth=cmdPeriodClose("Costbook")		
		
			Session("v_Costbook")=txtLastDay
			Session("cls_Costbook")= txtNextMonth	

							
			TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
		
	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Costbook'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

	strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME) VALUES (" & _
					" 'Costbook', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" " & tempMaxLine & "" & _
					", SYSDATE, TO_CHAR(SYSDATE,'HH24:MI') " & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Costbook'"
				
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.CursorType = 3
	rsLenghtime.CursorLocation = 3
	rsLenghtime.Open strLenghtime, connstr

	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Costbook'"		
	connStr.execute(upClosingLog)
	'--End Added--
%>
	<!--#include virtual="/insert_nextmonthbjr.asp"-->
<%
	
	'**** Eksekusi ERP Report ****'
	'Server.Execute("p_Generate_ERP_Report.asp")
	
	
	'**** Eksekusi AR Report ****'
	'Server.Execute("p_Generate_Account_Report.asp")
	
	year_process = Cdbl(Mid(cmdAccountPeriodClose("Costbook"), 1, 4))
	month_process = Cdbl(Mid(cmdAccountPeriodClose("Costbook"), 5, 2))
	
	qLineMax = " SELECT NVL(MAX(lineno),0) tempMaxLine FROM process_month_log WHERE fcba = '" & session("fcBA") & "'" 
							
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		new_max_log = cdbl(rsqLineMax("tempMaxLine"))
	else
		new_max_log = 0
	end if

	new_max_log = new_max_log + 1
	
	strClosingLog = " INSERT INTO process_month_log( lineno" & _
	                "                                , fccode" & _
	                "                                , fcname" & _
	                "                                , year_process" & _
	                "                                , month_process" & _
	                "                                , execution_date" & _
	                "                                , fcentry" & _
	                "                                , fcedit" & _
	                "                                , fcip" & _
	                "                                , fcba" & _
	                "                                , lastupdate" & _
	                "                                , lasttime)" & _
	                "  VALUES ( '" & new_max_log & "'" & _
					"           , NULL" & _
					"           , 'Final Closing'" & _
					"           , '" & year_process & "'" & _
					"           , '" & month_process & "'" & _
					"           , TO_DATE(TO_CHAR(SYSDATE, 'mm/dd/yyyy'), 'mm/dd/yyyy')" & _
					"           , '" & Session("CodeName") & "'" & _
					"           , '" & Session("CodeName") & "'" & _
					"           , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"           , '" & Session("fcBA") & "'" & _
					"           , SYSDATE" & _
					"           , TO_CHAR(SYSDATE, 'hh24:mi') )"
	connStr.execute(strClosingLog)
	
	
	%>
	<%end if%>

	<%if err.number <> 0 then%>
		<br><img src="images/error.gif"><%
		'connStr.RollbackTrans
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		cmdView	= "True"%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>	
	<%Else%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Financial Monthly Closing is succesfully...
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
	<br><br><img src="images/error.gif">Financial Monthly is NOT completed...
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