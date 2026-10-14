<%
'Option Explicit
%>
<%On Error Resume Next%>
<%response.buffer=True%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>iPS - Delivery Closing</title>
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
Delivery Last Period Close <%=FormatDateTime(Session("cls_Delivery"),1)%> <br> 
Delivery Current Period is <%=FormatDateTime(Session("v_Delivery"),1)%> <br><br>
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

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Delivery','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Delivery " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 

	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	" where SYSTEM='Delivery' and fcba='" & Session("fcBA") & "'"


	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Delivery " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

	'---- Control Job
	qControlJob="select ControlJob from Module where NAME='Delivery'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For Delivery is Empty"
	   txtControlJob=""
	end if
	response.Flush()
	
		'**** Remark By ukie 1 Februari 2016 ****'
		' txtLastDay  =FormatDateTime(LastDayOfMonth(Session("v_Delivery"),"2"),2)
		' txtNextMonth=FormatDateTime(LastDayOfMonth(Session("v_Delivery"),"2")+1,2)
		' q=" Update PeriodControl Set LASTPERIODPOSTED=" & _
		  ' " to_date('" & txtLastDay & "','mm/dd/yyyy') ,CURRENTDATE=" & _
		  ' " to_date('" & txtNextMonth & "','mm/dd/yyyy')  " & _
		  ' " where SYSTEM='Delivery' and fcba='" & Session("fcBA") & "'"
		' connStr.execute(q)
		
		'**** Remark By ukie 1 Februari 2016 ****'
		' q=" Update PeriodControl Set LASTPERIODPOSTED=add_months(LASTPERIODPOSTED,1),CURRENTDATE=add_months(CURRENTDATE,1) , lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi')" & _
		  ' " where SYSTEM='Delivery' and fcba='" & Session("fcBA") & "'"
		' connStr.execute(q)
		
		'**** Add By ukie 1 Februari 2016 ****'						Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear,temp_EndingAccountingPeriod
							
						QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
											" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
											" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
											" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear,  " & _ 
											" periodcontrol_api.GetEndingAccountingPeriodTY(system,'" & Session("fcBA") & "')  EndingAccountingPeriod " & _ 
											" from periodcontrol p  " & _ 
											" where UPPER(system) = UPPER('Delivery')  and fcba='" & Session("fcBA") & "'"
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
								Response.Write  "<font face=Verdana size=1 color=red>" & "IPLAS Error, No PERIODCONTROL for Delivery " & "<br>"
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
							   " where UPPER(system) = UPPER('Delivery')  and fcba='" & Session("fcBA") & "'"
						ELSE
							IF temp_NextCloseAccountingPeriod < 12 THEN
								q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
								  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
								  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
								  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
								  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
								  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
								  " where UPPER(system) = UPPER('Delivery')  and fcba='" & Session("fcBA") & "'"
							ELSE
								 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
								   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
								  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
								  " where  UPPER(system) = UPPER('Delivery')  and fcba='" & Session("fcBA") & "'"
							END IF
						END IF
						
						IF temp_NextCloseAccountingPeriod = temp_EndingAccountingPeriod AND temp_NextCurrentAccountingPeriod = 1  THEN
						
							q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
							  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
							  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							  " where UPPER(system) = UPPER('Delivery') and and fcba='" & Session("fcBA") & "'"
						END IF			
			
			
		connStr.execute(q)
		
		txtLastDay  = cmdFindPeriod("Delivery")
		txtNextMonth = cmdPeriodClose("Delivery")
		
		Session("v_Delivery") = txtLastDay
		Session("cls_Delivery") = txtNextMonth	

		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
		
	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Delivery'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

	strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, fcentry) VALUES (" & _
					" 'Delivery', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" " & tempMaxLine & "" & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Delivery'"
				
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.CursorType = 3
	rsLenghtime.CursorLocation = 3
	rsLenghtime.Open strLenghtime, connstr

	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Delivery'"		
	connStr.execute(upClosingLog)
	'--End Added--

	if err.number <> 0 then%>
		<br><img src="images/error.gif"><%
		'connStr.RollbackTrans
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		cmdView	= "True"%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>	
	<%Else%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Delivery Closing is succesfully...
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
	<br><br><img src="images/error.gif">Delivery Closing is NOT completed...
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