<%
  'Option Explicit
%>
<%
   'On Error Resume Next
 %>
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
Dim CLOSINGYEAR : CLOSINGYEAR = Request("CLOSINGYEAR")

%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Please Wait...FINANCIAL YEARLY CLOSING Executed...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br>" 
 Response.Write "You are executing Yearly Closing " & CLOSINGYEAR & "<br><br>" 
 
 
dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode
Dim qDataDeleteQuery13MonthYear, qDataDeleteQueryNextYear, QueryCek,CurrentYear,rsendPeriod,Temp_EndingAccountingPeriodTY

Set rsCek1 = Server.CreateObject("ADODB.Recordset")
	rsCek1.CursorType = 3
	rsCek1.CursorLocation = 3
	
	' QueryCek = " select to_number(to_char(lastperiodposted, 'YYYY')) CurrentYear , to_char(lastperiodposted, 'Month-YYYY') CurrentPeriod" & _
    '            " from periodcontrol where fcba= '" & Session("fcBA") & "' and upper(system)='COSTBOOK' " 
			   
	 QueryCek = " select Closeaccountingperiod CurrentPeriod , Closeaccountingyear CurrentYear " & _
			    " from periodcontrol where fcba= '" & Session("fcBA") & "' and upper(system)='COSTBOOK' " 
			   
	Response.Write	QueryCek	   
 rsCek1.Open QueryCek, connstr	
 if not rsCek1.eof then
    CurrentYear = cdbl(rsCek1("CurrentYear"))
    CurrentPeriod = rsCek1("CurrentPeriod")
 else
    CurrentYear = "0"   
 end if

' rsCek1.close
' set rsCek1 = Nothing	

'Response.Write	CLOSINGYEAR & CurrentYear 

IF  CurrentYear <> "0"  and cdbl(CLOSINGYEAR)=cdbl(CurrentYear) then 

 
   
   Set  rsendPeriod = Server.CreateObject("ADODB.Recordset")
		rsendPeriod.CursorType = 3
		rsendPeriod.CursorLocation = 3

   QueryCek = " select periodcontrol_api.GetEndingAccountingPeriodTY('COSTBOOK','" & Session("fcBA") & "') EndingAccountingPeriodTY from dual "

   response.write QueryCek
   rsendPeriod.Open QueryCek, connstr	
	if not rsendPeriod.eof then
		Temp_EndingAccountingPeriodTY = cdbl(rsendPeriod("EndingAccountingPeriodTY"))
	else
		Temp_EndingAccountingPeriodTY = "0"
	end if

	
	Set rsCek = Server.CreateObject("ADODB.Recordset")
		rsCek.CursorType = 3
		rsCek.CursorLocation = 3
		'QueryCek = " select to_number(to_char(lastperiodposted,'MM')) temp_SourceMonth,to_number(to_char(lastperiodposted,'YYYY')) temp_SourceYear, " & _
		'		    " to_number(to_char(Currentdate,'MM')) temp_TargetNextMonth , to_number(to_char(Currentdate,'YYYY')) temp_TargetNextYear " & _
		'		    " from periodcontrol where fcba='" & Session("fcBA") & "'  and upper(system)='COSTBOOK' and  to_char(lastperiodposted,'MMYYYY')='12" & CLOSINGYEAR  & "'"
		
		' Add By Akbar, 21 Desember 2015
		
		QueryCek = " select closeaccountingperiod temp_SourceMonth, closeaccountingyear   temp_SourceYear,  " & _
				   " currentaccountingperiod  temp_TargetNextMonth, currentaccountingyear temp_TargetNextYear, " & _
				   " periodcontrol_api.GetEndingAccountingPeriodTY('COSTBOOK','" & Session("fcBA") & "' ) TargetMonth  " & _
				   " from periodcontrol where fcba='" & Session("fcBA") & "'  and upper(system)='COSTBOOK' and  closeaccountingperiod  = " & Temp_EndingAccountingPeriodTY 

		
	response.write QueryCek	
	rsCek.Open QueryCek, connstr	
	if not rsCek.eof then
			
		temp_SourceYear = cdbl(rsCek("temp_SourceYear"))
		temp_SourceMonth = cdbl(rsCek("temp_SourceMonth")) - 1
		
		temp_TargetYear = cdbl(rsCek("temp_SourceYear"))
		temp_TargetMonth = cdbl(rsCek("TargetMonth"))
		
		temp_TargetNextYear = cdbl(rsCek("temp_TargetNextYear"))
		temp_TargetNextMonth = cdbl(rsCek("temp_TargetNextMonth"))

	else
		temp_SourceYear = "0"
		temp_SourceMonth = "0"
		
		temp_TargetYear = "0"
		temp_TargetMonth = "0"
		
		temp_TargetNextYear = "0"
		temp_TargetNextMonth = "0"	
	  
	end if
	
	    ' Response.write " temp_SourceYear  "  & temp_SourceYear
		' Response.write "temp_SourceMonth  " & temp_SourceMonth
		
		' Response.write "temp_TargetYear " & temp_SourceYear
		' Response.write "temp_TargetMonth " & temp_TargetMonth
		
		' Response.write "temp_TargetNextYear " & temp_TargetNextYear
		' Response.write "temp_TargetNextMonth " & temp_TargetNextMonth
	
	
    IF temp_SourceMonth <>  "0" and temp_SourceMonth <>  "" then 
			rsCek.close
			set rsCek = Nothing	

			'---- Control Job
			' Untuk Sementara pakai hard-code, nantinya kita harus rubah menggunakan parameter.
			' temp_ReturnedEarningCode = "3110300"
			' temp_ReturnedEarningDesc = "Rugi/Laba yang Ditahan"

			' txtControlJob = temp_ReturnedEarningCode


			' if txtControlJob<>"" then
				'---- Normalisasi
				
				' Yang diinput oleh User adalah Closing Tahun "2013", 
				' Informasi dibawah ini harus di generate langsung oleh System 
				' Source Month = 12, TargetMonth = 13
				' Target Next Year = 2014 (2013 + 1)
				' Target Next Month = 1

				' Validasinya adalah :
				' Tahun Closing yang akan diexecute adalah harus lebih besar dari Account Period yang aktif dan sudah close
				' Last Period Postednya harus bulan 12 (201312) tahun Closing
				' Current Datenya harus bulan 1 (201401) tahun Closing + 1

					
				' temp_SourceYear = 2013
				' temp_SourceMonth = 12
				
				' temp_TargetYear = 2013
				' temp_TargetMonth = 13
				
				' temp_TargetNextYear = 2014
				' temp_TargetNextMonth = 1
				
					' Next Year, By NAF
				' First we have to remove the Next Year TrialBalance
				' Second we can insert data from 13'th Period
					
				qDataDeleteQuery13MonthYear = " DELETE " & _
											  " FROM trialbalance " & _
											  " WHERE fcba = '" & Session("fcBA") & "' AND period_year = " & temp_TargetYear & _
											  " AND period_month = " & temp_TargetMonth & " " 
				
				connstr.execute(qDataDeleteQuery13MonthYear)
				
				
				' Next Year, By NAF
				' First we have to remove the Next Year TrialBalance
				' Second we can insert data from 13'th Period
					
				qDataDeleteQueryNextYear = " DELETE " & _
										   " FROM trialbalance " & _
										   " WHERE fcba = '" & Session("fcBA") & "' AND period_year = " & temp_TargetNextYear & _
										   " AND period_month = " & temp_TargetNextMonth & " " 
				
				connStr.execute(qDataDeleteQueryNextYear)
				
				
				

				
				' Process to closing financial Year
				' 1. Get Data from Ending Of Year in Ending_Balance Column in Trial Balance
				' 2. Insert data into 
				
				

				

				
				Set rsTransactionBalance = Server.CreateObject("ADODB.Recordset")
				rsTransactionBalance.CursorType = 3
				rsTransactionBalance.CursorLocation = 3
				
				qTransactionBalance = " SELECT " & _
									  " sum(debet - credit) transaction_balance " & _
									  " FROM ( SELECT period_month, " & _
													   " period_year, " & _
													   " fccode, " & _
													   " fcname, " & _
													   " ending_balance opening_balance, " & _
													   " DECODE (YEAR_FINANCIAL_CLOSING_API.IsRevenues (fcba, fccode), " & _
															   " 'TRUE', ending_balance, " & _
															   " 0) " & _
														  " DEBET, " & _
													   " DECODE (YEAR_FINANCIAL_CLOSING_API.IsExpenses (fcba, fccode), " & _
															   " 'TRUE', ending_balance, " & _
															   " 0) " & _
														 "  CREDIT, " & _
													   " fcba " & _
												  " FROM trialbalance " & _
												 " WHERE fcba = '" & Session("fcBA") & "' AND period_year = " & temp_SourceYear & " AND period_month = " & temp_SourceMonth & " " & _
												 " ) " 
				' response.write qTransactionBalance
				rsTransactionBalance.Open qTransactionBalance, connstr
				
				
				if not rsTransactionBalance.eof then
					tempTransactionBalance = cdbl(rsTransactionBalance("transaction_balance"))
				end if
				
				if tempTransactionBalance > 0 then 
					tempTransaction_Flag = "PLUS"
				else 
					tempTransaction_Flag = "MINUS"
				end if

				rsTransactionBalance.close
				set rsTransactionBalance = Nothing	
				
				
				
				
				
				Set rsDataQuery = Server.CreateObject("ADODB.Recordset")
				rsDataQuery.CursorType = 3
				rsDataQuery.CursorLocation = 3
				
				qDataQuery = 	" SELECT " & _
								" fccode, " & _
								" fcname, " & _
								" sum (opening_balance) opening_balance, " & _
								" sum(DEBET) debet, " & _
								" sum (CREDIT) credit " & _
								" FROM (SELECT fccode, " & _
							   	" fcname, " & _
							    " ending_balance opening_balance, " & _
							    " YEAR_FINANCIAL_CLOSING_API.GetValueDB (fcba, fccode, ending_balance) DEBET, " & _
							    " YEAR_FINANCIAL_CLOSING_API.GetValueCR (fcba, fccode, ending_balance) CREDIT " & _
								" FROM trialbalance " & _
								" WHERE fcba = '" & Session("fcBA") & "' AND period_year = " & temp_SourceYear & " AND period_month = " & temp_SourceMonth & " ) " 	& _
								" group by fccode, fcname "

							' " SELECT " & _
							' " fccode, " & _
							' " fcname, " & _
							' " sum (opening_balance) opening_balance, " & _
							' " sum(DEBET) debet, " & _
							' " sum (CREDIT) credit " & _
							' " FROM (SELECT fccode, " & _
						 '   	" fcname, " & _
						 '    " ending_balance opening_balance, " & _
						 '    " YEAR_FINANCIAL_CLOSING_API.GetValueDB (fcba, fccode, ending_balance) DEBET, " & _
						 '    " YEAR_FINANCIAL_CLOSING_API.GetValueCR (fcba, fccode, ending_balance) CREDIT " & _
							' " FROM trialbalance " & _
							' " WHERE fcba = '" & Session("fcBA") & "' AND period_year = " & temp_SourceYear & " AND period_month = " & temp_SourceMonth & " " & _
							' " UNION " & _
							' " select '" &  temp_ReturnedEarningCode & "'," & _
							' " '" & temp_ReturnedEarningDesc & "'," & _
							' " 0 opening_balance, " & _
							' " decode( '" & tempTransaction_Flag & "','MINUS',abs(" & tempTransactionBalance & "),0) DEBET, " & _  
							' " decode( '" & tempTransaction_Flag & "','PLUS',abs(" & tempTransactionBalance & "),0) CREDIT " & _  
							' " FROM dual " & _
							' " ) " 	& _
							' " group by fccode, fcname "
				'response.write qDataQuery
				rsDataQuery.Open qDataQuery, connstr	
				
				
				
				if not rsDataQuery.eof then
					Do While Not rsDataQuery.eof
					
						nOpening = cdbl(rsDataQuery("opening_balance"))
						nDebet = cdbl(rsDataQuery("debet"))
						nCredit = cdbl(rsDataQuery("credit"))
						nEnding = nOpening + nDebet + nCredit
						
						
						' To Check Counter of Data 
						Set rsDataQueryCheck = Server.CreateObject("ADODB.Recordset")
						rsDataQueryCheck.CursorType = 3
						rsDataQueryCheck.CursorLocation = 3
						
						tempQueryCheck = " select count(*) total " & _
										" from TrialBalance " & _
										" where " & _
										" PERIOD_YEAR = " & temp_TargetYear & " and "  & _
										" PERIOD_MONTH = " & temp_TargetMonth & " and " & _
										" FCCODE = '" & rsDataQuery("fccode") & "' and " & _
										" FCBA = '" & Session("fcBA") & "' " 
										
						rsDataQueryCheck.Open tempQueryCheck, connstr		

						if not rsDataQueryCheck.eof then
							tempCheckCounter = cdbl(rsDataQueryCheck("total"))
						end if
						rsDataQueryCheck.close
						set rsDataQueryCheck = Nothing	
						' End of Check Counter 
						
						if tempCheckCounter = 0 then 
						
							qInsertData =" Insert into TrialBalance " & _
									  "(PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
									  " values " & _
									  "(" & temp_TargetMonth & "," & temp_TargetYear & ",'" & rsDataQuery("fccode") & "','"  & _
									  rsDataQuery("fcname") & "'," & nOpening & "," & nDebet & "," & nCredit & "," & nEnding  & ",'"  & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))" 

							'response.write 	qInsertData & "</br>"	  
							connstr.execute(qInsertData)	
						else
							qUpdateData = " Update TrialBalance " & _
											" set " & _
											" debet = debet + " & nDebet & "," & _
											" credit = credit + " & nCredit & "," & _
											" ENDING_BALANCE = ENDING_BALANCE + " & nEnding & "" & _
											" where " & _
											" PERIOD_YEAR = " & temp_TargetYear & " and "  & _
											" PERIOD_MONTH = " & temp_TargetMonth & " and " & _
											" FCCODE = '" & rsDataQuery("fccode") & "' and " & _
											" FCBA = '" & Session("fcBA") & "' " 

							'response.write 	qUpdateData & "</br>"	  
							connstr.execute(qUpdateData)	
						end if
						

						
						rsDataQuery.MoveNext
					Loop '----------- Loop For rsDataQuery  
				
				end if
				rsDataQuery.close
				set rsDataQuery = Nothing			
				
				
				
				
				
				' Create Next Year
				Set rsDataQueryNextYear = Server.CreateObject("ADODB.Recordset")
				rsDataQueryNextYear.CursorType = 3
				rsDataQueryNextYear.CursorLocation = 3
				
				qDataQueryNextYear = " SELECT " & _
													   " fccode, " & _
													   " fcname, " & _
													   " ending_balance opening_balance, " & _
													   " 0 DEBET, " & _
													   " 0 CREDIT, " & _
													   " 0 ending_balance " & _ 
												  " FROM trialbalance " & _
												 " WHERE fcba = '" & Session("fcBA") & "' AND period_year = " & temp_TargetYear & " AND period_month = " & temp_TargetMonth & " " 
				
				rsDataQueryNextYear.Open qDataQueryNextYear, connstr	
				
				if not rsDataQueryNextYear.eof then
					Do While Not rsDataQueryNextYear.eof
					
						nOpening = cdbl(rsDataQueryNextYear("opening_balance"))
						nDebet = cdbl(rsDataQueryNextYear("debet"))
						nCredit = cdbl(rsDataQueryNextYear("credit"))
						nEnding = cdbl(rsDataQueryNextYear("ending_balance"))
						
						qInsertData =" Insert into TrialBalance " & _
								  "(PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
								  " values " & _
								  "(" & temp_TargetNextMonth & "," & temp_TargetNextYear & ",'" & rsDataQueryNextYear("fccode") & "','"  & _
								  rsDataQueryNextYear("fcname") & "'," & nOpening & "," & nDebet & "," & nCredit & "," & nEnding  & ",'"  & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate, 'HH24:MI'))" 
						connstr.execute(qInsertData)	
						
						rsDataQueryNextYear.MoveNext
					Loop '----------- Loop For rsDataQuery  
				
				end if
				rsDataQueryNextYear.close
				set rsDataQueryNextYear = Nothing			
				
				connstr.execute("commit")	
				TimeAkhir=Time()
				Lama =TimeAkhir-TimeAwal
				Response.Write "Finish ... " & time() & "<br><br>"
				Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."%>
			
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
				<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Congratulations, -Financial Year Closing- is succesfully generated...
				<script language="JavaScript" type="text/javascript" >
				document.getElementById('status_running').style.visibility = 'hidden';
				</script>
			<%End if 
	Else 
     Response.Write "You can't execute Yearly Closing " & CLOSINGYEAR & ", Your Current Period Is " & CurrentPeriod & ", Your End Period This Year Is " & Temp_EndingAccountingPeriodTY &"<br>"
    End if%>		
<% 
Else 
     Response.Write "You can't execute Yearly Closing " & CLOSINGYEAR & ", Your Current Yearly Period Is " & CurrentYear & "<br>"
End if%>
  


