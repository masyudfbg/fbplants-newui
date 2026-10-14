<%'Option Explicit%>
<%'On Error Resume Next%>
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
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Project Last Period Close <%=FormatDateTime(Session("cls_Project"),1)%> <br> 
Project Current Period is <%=FormatDateTime(Session("v_Project"),1)%> <br><br>
Please Wait...<br><br>
<%
 Response.Write "Start ..." & Time() & "<br>" 

dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode
'---- Control Job
qControlJob="select ControlJob from Module where NAME='Project'"
Set rsControlJob = Server.CreateObject("ADODB.Recordset")
rsControlJob.CursorType = 3
rsControlJob.CursorLocation = 3
rsControlJob.Open qControlJob, connstr
if not rsControlJob.eof then
   txtControlJob=rsControlJob("ControlJob")
else
   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For Project is Empty"
   txtControlJob=""
end if
response.Flush()

if txtControlJob<>"" then
	'---- Normalisasi
	Response.Write " Normalisasi Project Base Costing" & "<br>"
	Response.Flush()
	
	qNormalisasi=" Delete From COST_ACTIVITY where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Project")) &  _
				 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Project")) & " and fcBA='" & Session("fcBA") & "'" & _
				 " and SOURCETYPE in ('PJ')"
	connstr.execute(qNormalisasi)
	
	'Qty Order
	MSQL=" SELECT RD.LOCATIONCODE AS ProjectCODE,RD.PURCHASEITEMCODE AS ITEMCODE,TO_CHAR(RN.RNDATE,'YYYY') AS NYEAR," & _
		" TO_CHAR(RN.RNDATE,'MM') AS NMONTH,nvl(SUM(RD.QUANTITY),0) AS QUANTITY " & _
		" FROM RECEIVENOTE RN,RECEIVENOTEDETAIL RD WHERE " & _
		" (RN.RECEIVENOTECODE=RD.RECEIVENOTECODE AND RN.FCBA=RD.FCBA) AND " & _
		" to_number(to_Char(RN.RNDATE,'mm'))=" & Month(Session("v_Project")) &  _
		" and to_number(to_Char(RN.RNDATE,'yyyy'))=" & Year(Session("v_Project")) & _
	    " and rd.LOCATIONTYPE='" & "PJ" & "'" & _
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
			q="SELECT ProjectCODE FROM Project_LEDGER WHERE ProjectCODE ='" & rsReceive("ProjectCODE") & _
			  "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			Set rsCARI = Server.CreateObject("ADODB.Recordset")
			rsCARI.CursorType = 3
			rsCARI.CursorLocation = 3
			rsCARI.Open q, connstr
			'RESPONSE.Write(Q)	 
			if rsCARI.eof Then
				q= " Insert into Project_LEDGER " & _
				   " (ProjectCODE, YEAR, MONTH, QTYRECEIVED, FCENTRY, FCEDIT, FCIP, FCBA) " & _
				   " values ('"  &  rsReceive("ProjectCODE") & "',"  & rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
				   rsReceive("QUANTITY") & ",'"  & _
				   Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				   Response.Write(q)
				   response.Flush()
				   connStr.execute(q)
			Else
				q= "Update Project_LEDGER set QTYRECEIVED =" & rsReceive("QUANTITY") & _
				   " where  ProjectCODE ='"  &  rsReceive("ProjectCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
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
	Response.Write "<br><font color=Blue>" & "Closing Order Project is completed..." & "</font>"
	
    'Qty Issue 
	MSQL= " SELECT ProjectCODE, to_number(to_Char(ISSUEDATE,'mm')) AS NMONTH ,to_number(to_Char(ISSUEDATE,'yyyy')) AS NYEAR," & _
	      " SUM(QUANTITY) AS QUANTITY from Project_ISSUE where " & _
		  " to_number(to_Char(ISSUEDATE,'mm'))=" & Month(Session("v_Project")) &  _
		  " and to_number(to_Char(ISSUEDATE,'yyyy'))=" & Year(Session("v_Project")) & _
		  " and fcBA='" & Session("fcBA") & "'" & _
		  " Group by ProjectCODE, to_number(to_Char(ISSUEDATE,'mm')) ,to_number(to_Char(ISSUEDATE,'yyyy'))"
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
			q="SELECT ProjectCODE FROM Project_LEDGER WHERE ProjectCODE ='" & rsReceive("ProjectCODE") & _
			  "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			Set rsCARI = Server.CreateObject("ADODB.Recordset")
			rsCARI.CursorType = 3
			rsCARI.CursorLocation = 3
			rsCARI.Open q, connstr	 
			if rsCARI.eof Then
				q= " Insert into Project_LEDGER " & _
				   " (ProjectCODE, YEAR, MONTH, QTYISSUED, FCENTRY, FCEDIT, FCIP, FCBA) " & _
				   " values ('"  &  rsReceive("ProjectCODE") & "',"  & rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
				   rsReceive("QUANTITY") & ",'"  & _
				   Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				   'Response.Write(q)
				   'response.Flush()
				   connStr.execute(q)
			Else
				q= "Update Project_LEDGER set QTYISSUED =" & rsReceive("QUANTITY") & _
				   " where  ProjectCODE ='"  &  rsReceive("ProjectCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
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
	Response.Write "<br><font color=Blue>" & "Closing Issue Project is completed..." & "</font>"
	
    'Qty Culling
	MSQL= " SELECT PROJECTCODE, to_number(to_Char(ADJUSMENTDATE,'mm')) AS NMONTH ,to_number(to_Char(ADJUSMENTDATE,'yyyy')) AS NYEAR," & _
	      " SUM(QUANTITY) AS QUANTITY,SUM(AMOUNT) AS AMOUNT from PROJECT_ADJUSMENT where " & _
		  " to_number(to_Char(ADJUSMENTDATE,'mm'))=" & Month(Session("v_Project")) &  _
		  " and to_number(to_Char(ADJUSMENTDATE,'yyyy'))=" & Year(Session("v_Project")) & _
		  " and fcBA='" & Session("fcBA") & "'" & _
		  " Group by ProjectCODE, to_number(to_Char(ADJUSMENTDATE,'mm')) ,to_number(to_Char(ADJUSMENTDATE,'yyyy'))"
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
			q="SELECT ProjectCODE FROM Project_LEDGER WHERE ProjectCODE ='" & rsReceive("ProjectCODE") & _
			  "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			Set rsCARI = Server.CreateObject("ADODB.Recordset")
			rsCARI.CursorType = 3
			rsCARI.CursorLocation = 3
			rsCARI.Open q, connstr	 
			if rsCARI.eof Then
				q= " Insert into Project_LEDGER " & _
				   " (ProjectCODE, YEAR, MONTH, QTYCULLED, AMOUNTADJUSTMENT,FCENTRY, FCEDIT, FCIP, FCBA) " & _
				   " values ('"  &  rsReceive("ProjectCODE") & "',"  & rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
				   rsReceive("AMOUNT") & "," & rsReceive("QUANTITY") & ",'"  & _
				   Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				   'Response.Write(q)
				   'response.Flush()
				   connStr.execute(q)
			Else
				q= "Update Project_LEDGER set QTYCULLED =" & rsReceive("QUANTITY") & ",AMOUNTADJUSTMENT =" & rsReceive("AMOUNT") & _
				   " where  ProjectCODE ='"  &  rsReceive("ProjectCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
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
	Response.Write "<br><font color=Blue>" & "Closing Culling Project is completed..." & "</font>"

	'-------- Capture Cost To Project
	MSQL=" SELECT TARGETCODE as ProjectCODE,TO_CHAR(TDATE,'YYYY') AS NYEAR," & _
		" TO_CHAR(TDATE,'MM') AS NMONTH,nvl(SUM(AMOUNT),0) AS AMOUNT" & _
		" FROM COST_ACTIVITY WHERE " & _
		" to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Project")) &  _
		" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Project")) & _
	    " and TARGETTYPE='" & "PJ" & "'" & _
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
			q="SELECT ProjectCODE FROM Project_LEDGER WHERE ProjectCODE ='" & rsReceive("ProjectCODE") & _
			  "' AND YEAR="  &  rsReceive("NYEAR") & " AND MONTH="  &  rsReceive("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			Set rsCARI = Server.CreateObject("ADODB.Recordset")
			rsCARI.CursorType = 3
			rsCARI.CursorLocation = 3
			rsCARI.Open q, connstr	 
			if rsCARI.eof Then
				q= " Insert into Project_LEDGER " & _
				   " (ProjectCODE, YEAR, MONTH, AMOUNTRECEIVED,FCENTRY, FCEDIT, FCIP, FCBA) " & _
				   " values ('"  &  rsReceive("ProjectCODE") & "',"  & rsReceive("NYEAR") & ","  &  rsReceive("NMONTH") & ","  &  _
				   rsReceive("AMOUNT") & ",'" & _
				   Session("CodeName") & "','" & Session("CodeName") & "','" & _
				   Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				  ' Response.Write(q)
				  ' response.Flush()
				   connStr.execute(q)
			Else
				q= "Update Project_LEDGER set AMOUNTRECEIVED=" & rsReceive("AMOUNT") & _
				   " where  ProjectCODE ='"  &  rsReceive("ProjectCODE") & "' AND YEAR="  &  rsReceive("NYEAR") & _
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
	Response.Write "<br><font color=Blue>" & "Capture Cost to Project is completed..." & "</font>"
	
    'Update Amt Issue di Tabel Project Ledger
    Dim rsAmtissue,nHargapokok,nAmtissue,nRp,nQty
	nRp=0
	nQty=0
	nHargapokok=0
	nAmtissue=0
	q= " SELECT ProjectCODE, YEAR, MONTH, OPENINGQTY, OPENINGAMOUNT, QTYRECEIVED, QTYCULLED, QTYISSUED, COSTPERSEEDLING, AMOUNTRECEIVED,AMOUNTADJUSTMENT, AMOUNTISSUED " & _ 
	   " FROM PROJECT_LEDGER WHERE QTYRECEIVED > 0 AND FCBA= '" & Session("fcBA") & "'" & _
	   " AND YEAR="  &  Year(Session("v_Project")) & " AND MONTH="  &  Month(Session("v_Project"))
		Set rsAmtissue = Server.CreateObject("ADODB.Recordset")
		rsAmtissue.CursorType = 3
		rsAmtissue.CursorLocation = 3
		rsAmtissue.Open q, connstr
	    Do While Not rsAmtissue.eof 
			nRp=(cdbl(rsAmtissue("OPENINGAMOUNT")) + cdbl(rsAmtissue("AMOUNTRECEIVED")) - cdbl(rsAmtissue("AMOUNTADJUSTMENT")))
			nQty=(cdbl(rsAmtissue("OPENINGQTY")) + cdbl(rsAmtissue("QTYRECEIVED")))-cdbl(rsAmtissue("QTYCULLED"))
			nHargapokok=int(nRp/nQty)
			nAmtissue=nHargapokok * cdbl(rsAmtissue("QTYISSUED"))
		   q="Update PROJECT_LEDGER set AMOUNTISSUED =" & nAmtissue & ",COSTPERSEEDLING=" & nHargaPokok & "" & _
			" where  ProjectCODE='"  &  rsAmtissue("ProjectCODE") &  "' AND YEAR="  &  _
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
	MSQL=   "SELECT ProjectCODE, ISSUEDATE, TO_CHAR(ISSUEDATE,'YYYY') AS NYEAR," & _
		    " TO_CHAR(ISSUEDATE,'MM') AS NMONTH,LOCATIONTYPE, LOCATIONCODE, JOBCODE, sum(QUANTITY) as QUANTITY" & _
			" FROM PROJECT_ISSUE WHERE " & _
			" to_number(to_Char(ISSUEDATE,'mm'))=" & Month(Session("v_Project")) &  _
			" and to_number(to_Char(ISSUEDATE,'yyyy'))=" & Year(Session("v_Project")) & " and fcBA='" & Session("fcBA") & "'" & _
			" GROUP BY ProjectCODE, ISSUEDATE,TO_CHAR(ISSUEDATE,'YYYY'),TO_CHAR(ISSUEDATE,'MM'), LOCATIONTYPE, LOCATIONCODE, JOBCODE"
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
			q="SELECT nvl(sum(COSTPERSEEDLING),0) as COSTPERSEEDLING  FROM PROJECT_LEDGER WHERE PROJECTCODE='"  &  rsIssue("ProjectCODE") & "' AND YEAR="  &  _
			  rsIssue("NYEAR") & " AND MONTH="  &  rsIssue("NMONTH") & " AND FCBA= '" & Session("fcBA") & "'"
			'response.Write(q)		
		    'response.Flush()
			Set rsCARII = Server.CreateObject("ADODB.Recordset")
			rsCARII.CursorType = 3
			rsCARII.CursorLocation = 3
			rsCARII.Open q, connstr	 
			nAmountIssue=0
			if Not rsCARII.eof Then
			   nAmountIssue = cdbl(rsCARII("COSTPERSEEDLING"))*cdbl(rsIssue("QUANTITY"))	
			   q="Update PROJECT_ISSUE set AMOUNT =" & nAmountIssue & _
				" where  ProjectCODE='" & rsIssue("ProjectCODE") & "' AND LOCATIONTYPE='"  &  rsIssue("LOCATIONTYPE") & "'" & _
				" AND LOCATIONCODE='"&  rsISsue("LOCATIONCODE") & "' and fcba='" & Session("fcBA") & "'" & _
				" and ISSUEDATE =to_date('" & rsIssue("ISSUEDATE") & "','mm/dd/yyyy') "  
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
	Response.Write "<br><font color=Blue>" & "Update Amount at Project Issue is completed..." & "</font>"
	
    'Insert Saldo Awal Bulan Depan
    Dim rsSaldoawal,QtySaldoawal,Bulan,Tahun,RpSaldoawal,rsCARISaldo
	q= " SELECT ProjectCODE, YEAR, MONTH, OPENINGQTY, OPENINGAMOUNT, QTYRECEIVED, QTYCULLED, QTYISSUED, COSTPERSEEDLING, AMOUNTRECEIVED,AMOUNTADJUSTMENT, AMOUNTISSUED " & _ 
	   " FROM PROJECT_LEDGER Where MONTH=" & Month(Session("v_Project")) &  _
       " and YEAR=" & Year(Session("v_Project")) & " and fcBA='" & Session("fcBA") & "'"
		Set rsSaldoawal = Server.CreateObject("ADODB.Recordset")
		rsSaldoawal.CursorType = 3
		rsSaldoawal.CursorLocation = 3
		rsSaldoawal.Open q, connstr

	    Do While Not rsSaldoawal.eof 
			Bulan=Month(Session("v_Project"))
			Tahun =Year(Session("v_Project"))
			'Response.Write(Bulan)
			'Response.Write(Tahun)
			if bulan = 12 then
				tahun=tahun + 1
				bulan= 1
			else
				tahun=tahun
				bulan= bulan + 1
			end if
		RpSaldoawal=cdbl(rsSaldoawal("OPENINGAMOUNT"))+cdbl(rsSaldoawal("AMOUNTRECEIVED"))-cdbl(rsSaldoawal("AMOUNTISSUED"))-cdbl(rsSaldoawal("AMOUNTADJUSTMENT"))
		QtySaldoawal=cdbl(rsSaldoawal("OPENINGQTY"))+cdbl(rsSaldoawal("QTYRECEIVED"))-cdbl(rsSaldoawal("QTYISSUED"))-cdbl(rsSaldoawal("QTYCULLED"))

		q="SELECT ProjectCODE FROM PROJECT_LEDGER WHERE ProjectCODE='"  &  rsSaldoawal("ProjectCODE") &  _
		  "' AND YEAR="  &  tahun & " AND MONTH="  &  bulan & " AND FCBA= '" & Session("fcBA") & "'"
		Set rsCARISaldo = Server.CreateObject("ADODB.Recordset")
		rsCARISaldo.CursorType = 3
		rsCARISaldo.CursorLocation = 3
		rsCARISaldo.Open q, connstr	 
		if rsCARISaldo.eof Then			
			q="Insert into PROJECT_LEDGER (ProjectCODE, YEAR, MONTH, OPENINGQTY, OPENINGAMOUNT, FCENTRY, FCEDIT, FCIP, FCBA) " & _
			"values ('"  &  rsSaldoawal("ProjectCODE") & "',"  & tahun & _
			","  &  bulan & ","  &  QtySaldoawal & ","  &  RpSaldoawal & _
			",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
			connStr.execute(q)
		else
		   q="Update PROJECT_LEDGER set OPENINGQTY =" & RpSaldoawal & ",OPENINGAMOUNT=" & QtySaldoawal & "" & _
			" where  ProjectCODE='"  &  rsSaldoawal("ProjectCODE") & "' AND YEAR="  &  _
			tahun & " AND MONTH="  &  bulan & " AND FCBA= '" & Session("fcBA") & "'"
			connStr.execute(q)
			connStr.execute("commit")
		End if
	    rsCARISaldo.close
	    set rsCARISaldo	=nothing
		rsSaldoawal.MOVENEXT
	loop
	rsSaldoawal.Close
	set rsSaldoawal=nothing
    Response.Write "<br><font color=Blue>" & "Update Project Opening Balance  is completed..." & "</font>"

	'----- Issue Bibitan
	'----- DB. Activity Sesuai User Pilih Ke Job Mana
	'----------- Cr. Control Job 800499 TRANSFER BIAYA PROJECT KE LUAR
	
	'----- Tranfer to Activity Costing Debet ke Lap, Nota dll tergantung Pemakaian
	Response.Flush()
	qABC =    " INSERT INTO COST_ACTIVITY(" & _
			  " SELECT ISSUEDATE, JOBCODE, 'PJ', 'PJ', ProjectCODE, '', LOCATIONTYPE, LOCATIONCODE, '', REFERENCE, SUM(QUANTITY), 0, SUM(AMOUNT)*1, FCBA" & _
			  " FROM Project_ISSUE " & _
			  " WHERE to_number(to_Char(ISSUEDATE,'mm'))=" & Month(Session("v_Project")) &  _
			  " and to_number(to_Char(ISSUEDATE,'yyyy'))=" & Year(Session("v_Project")) & " and fcBA='" & Session("fcBA") & "'" & _
			  " GROUP BY ISSUEDATE, JOBCODE, 'PJ', 'PJ', ProjectCODE, '', LOCATIONTYPE, LOCATIONCODE, '', REFERENCE,fcBA)"
	'response.Write " qABC "& qABC
	response.Write "<br><br>Transfer to Activity Based Costing..." & "<br>"
	response.Flush()		  
	connstr.Execute(qABC)

	Response.Flush()
	qABC =    " INSERT INTO COST_ACTIVITY(" & _
			  " SELECT ISSUEDATE, '820199', 'PJ', 'PJ', ProjectCODE, '', 'CB', 'BS', '', REFERENCE, SUM(QUANTITY), 0, SUM(AMOUNT)*-1, FCBA" & _
			  " FROM Project_ISSUE " & _
			  " WHERE to_number(to_Char(ISSUEDATE,'mm'))=" & Month(Session("v_Project")) &  _
			  " and to_number(to_Char(ISSUEDATE,'yyyy'))=" & Year(Session("v_Project")) & " and fcBA='" & Session("fcBA") & "'" & _
			  " GROUP BY ISSUEDATE, '820199', 'PJ', 'PJ', ProjectCODE, '', 'CB', 'BS', '', REFERENCE,fcBA)"
	'response.Write " qABC "& qABC
	response.Flush()		  
	connstr.Execute(qABC)

    'Jurnal to Balanced ( Activa )
    '-- db. 107005 = BIAYA AKUSISI PROJECT
	'       -- Cr 809902 = TRANSFER KE BIAYA AKUISISI PROJECT

	'Response.Flush()
	'qABC =    " INSERT INTO COST_ACTIVITY(" & _
    '		  " SELECT to_Date('" & Session("v_Project") & "','mm/dd/yyyy'), '107005', 'PJ', 'PJ', ProjectCODE, '', 'CB', 'BS', '', 'DB PJ TO BS', 0, 0, (SUM(AMOUNTRECEIVED)-sum(AMOUNTISSUED))*1, FCBA" & _
	'		  " FROM Project_LEDGER " & _
	'		  " WHERE Month=" & Month(Session("v_Project")) &  _
	'		  " and   Year=" & Year(Session("v_Project")) & " and fcBA='" & Session("fcBA") & "'" & _
	'		  " GROUP BY to_Date('"  & Session("v_Project") & "','mm/dd/yyyy'), '107005', 'PJ', 'PJ', ProjectCODE, '', 'CB', 'BS', '', 'DB PJ TO BS',fcBA)"
	'response.Write " qABC "& qABC
	'response.Flush()		  
	'connstr.Execute(qABC)


	'----- Tranfer to Activity Costing Debet ke Lap, Nota dll tergantung Pemakaian
	'Response.Flush()
	'Response.Flush()
	'qABC =    " INSERT INTO COST_ACTIVITY(" & _
	'		  " SELECT to_Date('" & Session("v_Project") & "','mm/dd/yyyy'), '809902', 'PJ', 'PJ', ProjectCODE, '', 'CB', 'BS', '', 'CR PJ TO BS', 0, 0, (SUM(AMOUNTRECEIVED)-sum(AMOUNTISSUED))*-1, FCBA" & _
	'		  " FROM Project_LEDGER " & _
	'		  " WHERE Month=" & Month(Session("v_Project")) &  _
	'		  " and   Year=" & Year(Session("v_Project")) & " and fcBA='" & Session("fcBA") & "'" & _
	'		  " GROUP BY to_Date('"  & Session("v_Project") & "','mm/dd/yyyy'), '809902', 'PJ', 'PJ', ProjectCODE, '', 'CB', 'BS', '', 'CR PJ TO BS',fcBA)"
	'response.Write " qABC "& qABC
	'response.Write "<br><br>Transfer to Project Acquisition Cost..." & "<br>"
	'response.Flush()		  
	'connstr.Execute(qABC)

	txtLastDay  =FormatDateTime(LastDayOfMonth(Session("v_Project"),"2"),2)
	txtNextMonth=FormatDateTime(LastDayOfMonth(Session("v_Project"),"2")+1,2)
        q=" Update PeriodControl Set LASTPERIODPOSTED=add_months(LASTPERIODPOSTED,1),CURRENTDATE=add_months(CURRENTDATE,1)   , lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
	  " where SYSTEM='Project' and fcba='" & Session("fcBA") & "'"
	connStr.execute(q)
	Session("v_Project")=txtLastDay
	Session("cls_Project")= txtNextMonth	

    TimeAkhir=Time()
	Lama =TimeAkhir-TimeAwal
	Response.Write "<br>Finish ... " & time() & "<br><br>"
	Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."%>
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
	<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Project Closing is succesfully...
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
<%End if%>
  


