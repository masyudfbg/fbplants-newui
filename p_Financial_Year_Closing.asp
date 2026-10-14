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
Dim fcba
Dim companycode : companycode =  Session("COMPANYCODE") 

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

' set default fcba : HO /companycode


Set rsCek1 = Server.CreateObject("ADODB.Recordset")
	rsCek1.CursorType = 3
	rsCek1.CursorLocation = 3
	
  QueryCek = " select fccode " & _
			 " from businessunit where fccompanycode = '" & Session("COMPANYCODE")  & "' AND fctype='HO' and isactive='TRUE' " 
			   
 'Response.Write	QueryCek	   
 rsCek1.Open QueryCek, connstr	
 if not rsCek1.eof then
    fcba = rsCek1("fccode")
 else
    fcba = session("fcba")  
 end if

 rsCek1.close
 set rsCek1 = Nothing	

 Response.Write "All Unit " & Session("COMPANYCODE") & " - " & fcba & "<br><br>" 
  
 Set rsCek1 = Server.CreateObject("ADODB.Recordset")
	rsCek1.CursorType = 3
	rsCek1.CursorLocation = 3

	
	' QueryCek = " select to_number(to_char(lastperiodposted, 'YYYY')) CurrentYear , to_char(lastperiodposted, 'Month-YYYY') CurrentPeriod" & _
    '            " from periodcontrol where fcba= '" & fcba & "' and upper(system)='COSTBOOK' " 
			   
	 QueryCek = " select Closeaccountingperiod CurrentPeriod , Closeaccountingyear CurrentYear " & _
			    " from periodcontrol where fcba= '" & fcba & "' and upper(system)='COSTBOOK' " 
			   
 'Response.Write	QueryCek	   
 rsCek1.Open QueryCek, connstr	
 if not rsCek1.eof then
    CurrentYear = cdbl(rsCek1("CurrentYear"))
    CurrentPeriod = rsCek1("CurrentPeriod")
 else
    CurrentYear = "0"   
 end if

 rsCek1.close
 set rsCek1 = Nothing	

Response.Write	CLOSINGYEAR & CurrentYear 

IF  CurrentYear <> "0"  and cdbl(CLOSINGYEAR)=cdbl(CurrentYear) then 

 
   
   Set  rsendPeriod = Server.CreateObject("ADODB.Recordset")
		rsendPeriod.CursorType = 3
		rsendPeriod.CursorLocation = 3

   QueryCek = " select periodcontrol_api.GetEndingAccountingPeriodTY('COSTBOOK','" & fcba & "') EndingAccountingPeriodTY from dual "

  ' response.write QueryCek
   rsendPeriod.Open QueryCek, connstr	
	if not rsendPeriod.eof then
		Temp_EndingAccountingPeriodTY = cdbl(rsendPeriod("EndingAccountingPeriodTY"))
	else
		Temp_EndingAccountingPeriodTY = "0"
	end if

	
	Set rsCek = Server.CreateObject("ADODB.Recordset")
		rsCek.CursorType = 3
		rsCek.CursorLocation = 3
		
		QueryCek = " select closeaccountingperiod temp_SourceMonth, closeaccountingyear   temp_SourceYear,  " & _
				   " currentaccountingperiod  temp_TargetNextMonth, currentaccountingyear temp_TargetNextYear, " & _
				   " periodcontrol_api.GetEndingAccountingPeriodTY('COSTBOOK','" & fcba & "' ) TargetMonth  " & _
				   " from periodcontrol where fcba='" & fcba & "'  and upper(system)='COSTBOOK' and  closeaccountingperiod  = " & Temp_EndingAccountingPeriodTY 

		
	'response.write QueryCek	
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
	
	     Response.write "temp_SourceYear  "  & temp_SourceYear & "<br>"
		 Response.write "temp_SourceMonth  " & temp_SourceMonth & "<br>"
		
		 Response.write "temp_TargetYear " & temp_SourceYear & "<br>"
		 Response.write "temp_TargetMonth " & temp_TargetMonth & "<br>"
		
		 Response.write "temp_TargetNextYear " & temp_TargetNextYear & "<br>"
		 Response.write "temp_TargetNextMonth " & temp_TargetNextMonth & "<br>"
	
	
    IF temp_SourceMonth <>  "0" and temp_SourceMonth <>  "" then 
			rsCek.close
			set rsCek = Nothing	
            Response.write "Start running Yearly Closing  " & temp_SourceMonth & "/" & temp_SourceYear & "<br>"
	
			'---- Control Job 
			' Jobcode ambil dari parameter parcode='JOBCODE_YEAR_CLOSE' 
			' dipindahkan ke 3210101 - Rugi/Laba yang Ditahan
			

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

					
				' temp_SourceYear = 2021
				' temp_SourceMonth = 14
				
				' temp_TargetYear = 2021
				' temp_TargetMonth = 15
				
				' temp_TargetNextYear = 2022
				' temp_TargetNextMonth = 1
				
				' Next Year, By NAF
				' First we have to remove the Next Year TrialBalance
				' Second we can insert data from 13'th Period
					
				qDataDeleteQuery13MonthYear = " DELETE " & _
											  " FROM trialbalance " & _
											  " WHERE fcba in (select fccode from businessunit where fccompanycode= '" & companycode & "') AND period_year = " & temp_TargetYear & _
											  " AND period_month = " & temp_TargetMonth & " " 
				Response.write qDataDeleteQuery13MonthYear & "<br>"
				connstr.execute(qDataDeleteQuery13MonthYear)
				
				
				' Next Year, By NAF
				' First we have to remove the Next Year TrialBalance
				' Second we can insert data from 13'th Period
					
				qDataDeleteQueryNextYear = " UPDATE  trialbalance SET opening_balance=0 " & _
										   " WHERE fcba in (select fccode from businessunit where fccompanycode= '" & companycode & "') AND period_year = " & temp_TargetNextYear & _
										   " AND period_month = " & temp_TargetNextMonth & " " 
				Response.write qDataDeleteQueryNextYear & "<br>"
				connStr.execute(qDataDeleteQueryNextYear)
				
				
				
				

				
				' Process to closing financial Year
				' 1. Get Data from Ending Of Year in Ending_Balance Column in Trial Balance
				' 2. Insert data into 
				
				qUpdate=" Insert into TrialBalance " & _
				" (PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				"  SELECT " &  temp_TargetMonth & " as PERIOD_MONTH,"  & temp_TargetYear & " as PERIOD_YEAR" & ", FCCODE," & _
				"  FCNAME, ENDING_BALANCE as OPENING_BALANCE,0 as DEBET, 0 as CREDIT,  0 as ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, SYSDATE, TO_CHAR(SYSDATE,'HH24:MI') " & _
				"  From TrialBalance Where Period_Month=" & temp_SourceMonth &  _
				"  and Period_Year=" & temp_SourceYear & " and fcba in (select fccode from businessunit where fccompanycode= '" & companycode & "') " & _
				"  ORDER BY FCBA, FCCODE "		 
				Response.write qUpdate & "<br>"
				connstr.execute(qUpdate)
				
				'3. Get Sum Coa closing year 
				qUpdate=" MERGE INTO TrialBalance x USING " & _
				        " ( " & _ 
						" 	SELECT fcba," & temp_TargetMonth & " period_month," & temp_TargetYear & " period_year,(select param_1 from parameter where rownum=1 and parcode='JOBCODE_YEAR_CLOSE' and fcba='" & fcba & "') coa, SUM(Ending_Balance) Amount FROM TrialBalance " & _ 
						"   WHERE fcba in (select fccode from businessunit where fccompanycode= '" & companycode & "') AND period_year = " & temp_SourceYear & " AND period_month = " & temp_SourceMonth & _ 
						"   AND fccode in (select fccode from parameter where parcode='JOBCODE_YEAR_CLOSE' and fcba='" & fcba & "')" & _ 
						"   GROUP BY fcba " & _
						" )y ON (x.fcba=y.fcba AND x.period_month=y.period_month AND x.period_year=y.period_year AND x.fccode=y.coa) " & _ 
						" WHEN MATCHED THEN " & _ 
						" UPDATE SET x.DEBET = CASE WHEN y.Amount > 0 then y.amount else 0 End, x.CREDIT= CASE WHEN y.Amount < 0 then y.amount else 0 End"
				Response.write qUpdate & "<br>"
				connstr.execute(qUpdate)
				
				
				'4. Get reclass Coa closing year 
				qUpdate=" MERGE INTO TrialBalance x USING " & _
				        " ( " & _ 
						" 	SELECT fcba," & temp_TargetMonth & " period_month," & temp_TargetYear & " period_year,fccode,ending_balance amount " & _
						"   FROM TrialBalance " & _
						"   WHERE fcba in (select fccode from businessunit where fccompanycode= '" & companycode & "') AND period_year = " & temp_SourceYear & " AND period_month = " & temp_SourceMonth & _ 
						"   AND fccode in (select fccode from parameter where parcode='JOBCODE_YEAR_CLOSE' and fcba='" & fcba & "')" & _ 
						" )y ON (x.fcba=y.fcba AND x.period_month=y.period_month AND x.period_year=y.period_year AND x.fccode=y.fccode) " & _ 
						" WHEN MATCHED THEN " & _ 
						" UPDATE SET x.DEBET = CASE WHEN y.Amount < 0 then y.amount*-1 else 0 End, x.CREDIT= CASE WHEN y.Amount > 0 then y.amount*-1 else 0 End "
				Response.write qUpdate & "<br>"
				connstr.execute(qUpdate)
				
				'5. Update ending_balance
				qUpdate=" UPDATE TrialBalance " & _
				        " SET " & _ 
						" 	ENDING_BALANCE = opening_balance + debet + credit " & _ 
						"   WHERE fcba in (select fccode from businessunit where fccompanycode= '" & companycode & "') AND period_year = " & temp_TargetYear & " AND period_month = " & temp_TargetMonth 
				Response.write qUpdate & "<br>"		
				connstr.execute(qUpdate)
				
				
				
				' Create Next Year
				qUpdate=" MERGE INTO TrialBalance x USING " & _
				        " ( " & _ 
						" 	SELECT fcba," & temp_TargetNextMonth & " period_month," & temp_TargetNextYear & " period_year,fccode,ending_balance " & _
						"   FROM TrialBalance " & _
						"   WHERE fcba in (select fccode from businessunit where fccompanycode= '" & companycode & "') AND period_year = " & temp_TargetYear & " AND period_month = " & temp_TargetMonth & _ 
						" )y ON (x.fcba=y.fcba AND x.period_month=y.period_month AND x.period_year=y.period_year AND x.fccode=y.fccode) " & _ 
						" WHEN MATCHED THEN " & _ 
						" UPDATE SET x.OPENING_BALANCE = y.ending_balance " 
				Response.write qUpdate & "<br>"
				connstr.execute(qUpdate)
				
				' Copy to fac_trialbalance
				if cdbl(temp_TargetYear) = 2022 then
				
				qUpdate="DELETE FROM FAC_TRIALBALANCE WHERE period_year=" & temp_TargetYear & " AND period_month=15 AND fcba IN (SELECT fccode FROM businessunit WHERE fccompanycode='" & companycode & "') "
				Response.write qUpdate & "<br>"
				connstr.execute(qUpdate)
				
				qUpdate="DELETE FROM FAC_TRIALBALANCE WHERE period_year="& temp_TargetNextYear & " AND period_month=1 AND fcba IN (SELECT fccode FROM businessunit WHERE fccompanycode='" & companycode & "')"
                Response.write qUpdate & "<br>"
				connstr.execute(qUpdate)

				qUpdate="INSERT INTO FAC_TRIALBALANCE " & _
				" (PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME) " & _
				" SELECT PERIOD_MONTH, PERIOD_YEAR,fccode,fcname,OPENING_BALANCE ,DEBET, " & _
				" CREDIT,ENDING_BALANCE,'System' FCENTRY,'System' FCEDIT,'::1' FCIP,fcba, SYSDATE LASTUPDATE, TO_CHAR(SYSDATE,'HH24:MI') LASTTIME " & _
				" FROM " & _
				" Trialbalance " & _
				" WHERE fcba IN (SELECT fccode FROM businessunit WHERE fccompanycode='" & companycode & "') " & _
				" AND period_month IN (15) AND period_year=" & temp_TargetYear & "" & _
				" ORDER BY fcba,period_month,fccode "
				Response.write qUpdate & "<br>"
				connstr.execute(qUpdate)

				qUpdate="Insert into FAC_TRIALBALANCE " & _
				" (PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME) " & _
				" SELECT PERIOD_MONTH, PERIOD_YEAR,fccode,fcname,OPENING_BALANCE ,DEBET, " & _
				" CREDIT,ENDING_BALANCE,'System' FCENTRY,'System' FCEDIT,'::1' FCIP,fcba, SYSDATE LASTUPDATE, TO_CHAR(SYSDATE,'HH24:MI') LASTTIME " & _
				" FROM " & _
				" Trialbalance " & _
				" WHERE fcba IN (SELECT fccode FROM businessunit WHERE fccompanycode='" & companycode & "') " & _
				" AND period_month IN (1) AND period_year= " & temp_TargetNextYear & "" & _
				" ORDER BY fcba,period_month,fccode "
				Response.write qUpdate & "<br>"
				connstr.execute(qUpdate)
			
				end if
				
				' Set rsDataQueryNextYear = Server.CreateObject("ADODB.Recordset")
				' rsDataQueryNextYear.CursorType = 3
				' rsDataQueryNextYear.CursorLocation = 3
				
				' qDataQueryNextYear = " SELECT " & _
													   ' " fccode, " & _
													   ' " fcname, " & _
													   ' " ending_balance opening_balance, " & _
													   ' " 0 DEBET, " & _
													   ' " 0 CREDIT, " & _
													   ' " 0 ending_balance " & _ 
												  ' " FROM trialbalance " & _
												 ' " WHERE fcba = '" & fcba & "' AND period_year = " & temp_TargetYear & " AND period_month = " & temp_TargetMonth & " " 
				
				' rsDataQueryNextYear.Open qDataQueryNextYear, connstr	
				
				' if not rsDataQueryNextYear.eof then
					' Do While Not rsDataQueryNextYear.eof
					
						' nOpening = cdbl(rsDataQueryNextYear("opening_balance"))
						' nDebet = cdbl(rsDataQueryNextYear("debet"))
						' nCredit = cdbl(rsDataQueryNextYear("credit"))
						' nEnding = cdbl(rsDataQueryNextYear("ending_balance"))
						
						' qInsertData =" Insert into TrialBalance " & _
								  ' "(PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
								  ' " values " & _
								  ' "(" & temp_TargetNextMonth & "," & temp_TargetNextYear & ",'" & rsDataQueryNextYear("fccode") & "','"  & _
								  ' rsDataQueryNextYear("fcname") & "'," & nOpening & "," & nDebet & "," & nCredit & "," & nEnding  & ",'"  & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								  ' Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, to_char(sysdate, 'HH24:MI'))" 
						' connstr.execute(qInsertData)	
						
						' rsDataQueryNextYear.MoveNext
					' Loop '----------- Loop For rsDataQuery  
				
				' end if
				' rsDataQueryNextYear.close
				' set rsDataQueryNextYear = Nothing			
				
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
  


