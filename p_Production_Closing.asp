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
<title>iPS - Mill_Production Closing</title>
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

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl



' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Mill_Production','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Mill_Production " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 


	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='Mill_Production' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Mill Production " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	'--End Added--	
	%>
	<div id="status_running"><img src="ICON/working.gif"></div><br>
	<font face="Verdana" size="1">
	Mill Production Last Period Close <%=FormatDateTime(Session("cls_Mill_Production"),1)%> <br> 
	Mill Production Current Period is <%=FormatDateTime(Session("v_Mill_Production"),1)%> <br><br>
	Please Wait...<br><br>
	<%
	 Response.Write "Start ..." & Time() & "<br>" 

	dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,qQuery,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
	dim v_AllowanceCode
	'---- Control Job
	qControlJob="select ControlJob from Module where NAME='Mill_Production'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For MillProduction is Empty"
	   txtControlJob=""
	end if
	response.Flush()

	
		'**** Remark By ukie 1 Februari 2016 ****'
	LastMonth = FormatDateTime(LastDayOfMonth(Session("v_Mill_Production"),"2"),2)
	
		' txtLastDay  =FormatDateTime(LastDayOfMonth(Session("v_Mill_Production"),"2"),2)
		' txtNextMonth=FormatDateTime(LastDayOfMonth(Session("v_Mill_Production"),"2")+1,2)
		' q=" Update PeriodControl Set LASTPERIODPOSTED=" & _
		  ' " to_date('" & txtLastDay & "','mm/dd/yyyy') ,CURRENTDATE=" & _
		  ' " to_date('" & txtNextMonth & "','mm/dd/yyyy')  " & _
		  ' " where SYSTEM='Mill_Production' and fcba='" & Session("fcBA") & "'"
		' connStr.execute(q)
		'---- Normalisasi
		
	response.write "Kode :" & Session("Estate_Mill") & "<br>" 
	if Session("Estate_Mill") = "M" then
	
		Dim TBS_QtySisaBulanlalu,TBS_Hargakemarin,TBS_QtyKemarin,TBS_SisaHarini,TBS_RpKemarin,TBS_QtyBulanini,TBS_RPBulanini,TBS_QtyOlah,TBS_RpOlah 
		Dim qCost_Activity,qNormalisasiDetail 
		
		'-- Bahan Baku TBS
		connstr.BeginTrans
		qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Mill_Production")) &  _
					 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Mill_Production")) & " and fcBA='" & Session("fcBA") & "'" & _
					 " and sourcecode='POM'"
		connstr.execute(qNormalisasiDetail)	
		
		If Err.Number = 0 Then
			connstr.CommitTrans
		Else
			connstr.RollbackTrans
			Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		End If
	    Response.Write  "<font face=Verdana size=1 color=Blue>" & "Normalisasi For MillProduction ,success " & "<br/></font>" 
		TBS_QtySisaBulanlalu=0
		TBS_Hargakemarin=0
		TBS_QtyKemarin=0
		TBS_SisaHarini=0
		TBS_RpKemarin=0
		TBS_QtyBulanini=0
		TBS_RPBulanini=0
		TBS_QtyOlah=0
		TBS_RpOlah=0
		
		CPO_QtyBlnini = 0
		PK_QtyBlnini  = 0
		
		CPO_RendemenBlnini = 0
		PK_RendemenBlnini = 0
		
		
		qQuery = "SELECT tbssisa TBS_QtySisaBulanlalu FROM productiontrn WHERE fcba='" & Session("fcBA") & "' AND proddate=trunc(To_Date('" & LastMonth & "','mm/dd/yyyy'),'month')-1 " 
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			TBS_QtySisaBulanlalu = cdbl(rsqLineMax("TBS_QtySisaBulanlalu"))
		else
			TBS_QtySisaBulanlalu = 0
		end if 
		'response.write(qQuery & "<br>")
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		qQuery = "SELECT tbssisa TBS_QtyKemarin FROM productiontrn WHERE fcba='" & Session("fcBA") & "' AND proddate=To_Date('" & Session("v_Mill_Production") & "','mm/dd/yyyy')-1 " 
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			TBS_QtyKemarin = cdbl(rsqLineMax("TBS_QtyKemarin"))
		else
			TBS_QtyKemarin = 0
		end if 
		
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		qQuery = "SELECT tbssisa TBS_SisaHarini FROM productiontrn WHERE fcba='" & Session("fcBA") & "' AND proddate=To_Date('" & LastMonth & "','mm/dd/yyyy') " 
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			TBS_SisaHarini = cdbl(rsqLineMax("TBS_SisaHarini"))
		else
			TBS_SisaHarini = 0
		end if 
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		qQuery = "SELECT  sum(tbsterima) TBS_QtyBulanini FROM productiontrn WHERE  To_Number(To_Char(proddate,'mm'))=" & Month(Session("v_Mill_Production")) & " AND To_Number(To_Char(proddate,'yyyy'))= " & Year(Session("v_Mill_Production")) & " and fcba='" & Session("fcBA") & "'" 
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			TBS_QtyBulanini = cdbl(rsqLineMax("TBS_QtyBulanini"))
		else
			TBS_QtyBulanini = 0
		end if 
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		qQuery = "SELECT round(Sum(amount)/Sum(quantity),0) TBS_Hargakemarin FROM cost_activity WHERE fcba='" & Session("fcBA") & "' AND accountingperiod=" & Month(Session("v_Mill_Production")) & " AND accountingyear= " & Year(Session("v_Mill_Production")) & _  
				 " AND journalprefix='SFB' AND tdate=last_day(tdate)-1 "
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			TBS_Hargakemarin = cdbl(rsqLineMax("TBS_Hargakemarin"))
		else
			TBS_Hargakemarin = 0
		end if 
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		qQuery = "SELECT Sum(amount)*-1 TBS_RPBulanini FROM cost_activity WHERE fcba='" & Session("fcBA") & "' AND accountingperiod=" & Month(Session("v_Mill_Production")) & " AND accountingyear= " & Year(Session("v_Mill_Production")) & _  
				 " AND journalprefix='SFB' AND  jobcode IN ('2110608','2110609') "
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			TBS_RPBulanini = cdbl(rsqLineMax("TBS_RPBulanini"))
		else
			TBS_RPBulanini = 0
		end if 
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		TBS_QtyOlah = TBS_QtySisaBulanlalu + TBS_QtyBulanini - TBS_SisaHarini 
		TBS_RpKemarin = TBS_QtyKemarin * TBS_Hargakemarin 
		
		TBS_RpOlah =  (TBS_RPBulanini / ( TBS_QtyBulanini + TBS_QtyKemarin) ) * TBS_QtyOlah
		
		
		' Get the Max LineNo of Cost Activity
		Set rsCostActivityMax = Server.CreateObject("ADODB.Recordset")
		rsCostActivityMax.CursorType = 3
		rsCostActivityMax.CursorLocation = 3
												
		qCostActivityMax =  " select nvl(max(lineno), 0) maxLineNo from cost_activity  where fcba = '" &  Session("fcBA") & "' " 
		rsCostActivityMax.Open qCostActivityMax, connstr
		If Not rsCostActivityMax.EOF Then
			maxLineNo = cdbl(rsCostActivityMax("maxLineNo"))
		Else
			maxLineNo = 0
		End If
		
		rsCostActivityMax.close
		set rsCostActivityMax = Nothing
		
		
		' response.write(TBS_QtySisaBulanlalu & "<br>")
		' response.write(TBS_QtyKemarin & "<br>")
		' response.write(TBS_SisaHarini & "<br>")
		' response.write(TBS_QtyBulanini & "<br>")
		' response.write(TBS_Hargakemarin & "<br>")
		' response.write(TBS_RPBulanini & "<br>")


		' response.write(TBS_QtyOlah & "<br>")
		' response.write(TBS_RpKemarin & "<br>")
		' response.write(TBS_RpOlah & "<br>")

		' response.flush
		
		connstr.BeginTrans
		maxLineNo = maxLineNo + 1
		qCost_Activity= "Insert into Cost_Activity " & _
						" ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
						" TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
						" QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE," & _
						" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
						" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
						" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
						" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
						" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
						" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
						" LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
						" SELECT To_Date('"& Session("v_Mill_Production") & "','mm/dd/yyyy') datein, '6610199' jobcode,'ST' source,'ST' SOURCETYPE,'POM' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						" 'CB' TARGETTYPE ,'BS' TARGETCODE,'','TBSOlah'||'/'||'" & Month(Session("v_Mill_Production")) & "'||'/'||'" & Year(Session("v_Mill_Production")) &"', " & _ 
						" " & TBS_QtyOlah & " QUANTITY, 0, " & TBS_RpOlah & " amount,'" & Session("fcBA") & "',to_date('" & Session("v_Mill_Production") & "','mm/dd/yyyy'),sysdate, " & _ 
						" '','', " & _ 
						" '','', " & _
						" '','', " & _
						" '','', " & _
						" '','','Dr', " & _
						"" & Month(Session("v_Mill_Production")) & " , " & Year(Session("v_Mill_Production")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'POM' FROM DUAL " 
		
		connstr.execute(qCost_Activity)
							
		If Err.Number = 0 Then
			connstr.CommitTrans
			Response.Flush
		Else
			connstr.RollbackTrans
			Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			Response.Flush
		End If
		
		connstr.BeginTrans
		maxLineNo = maxLineNo + 1
		qCost_Activity= "Insert into Cost_Activity " & _
						" ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
						" TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
						" QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE," & _
						" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
						" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
						" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
						" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
						" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
						" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
						" LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
						" SELECT To_Date('" & Session("v_Mill_Production") & "','mm/dd/yyyy') datein, '1140299' jobcode,'ST' source,'ST' SOURCETYPE,'POM' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						" 'CB' TARGETTYPE ,'BS' TARGETCODE,'','TBSOlah'||'/'||'" & Month(Session("v_Mill_Production")) & "'||'/'||'" & Year(Session("v_Mill_Production")) &"', " & _ 
						" " & TBS_QtyOlah & " QUANTITY, 0, " & TBS_RpOlah*-1 & " amount,'" & Session("fcBA") & "',to_date('" & Session("v_Mill_Production") & "','mm/dd/yyyy'),sysdate, " & _ 
						" '','', " & _ 
						" '','', " & _
						" '','', " & _
						" '','', " & _
						" '','','Cr', " & _
						"" & Month(Session("v_Mill_Production")) & " , " & Year(Session("v_Mill_Production")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'POM'  FROM DUAL" 
				
		connstr.execute(qCost_Activity)
							
		If Err.Number = 0 Then
			connstr.CommitTrans
			Response.Flush
		Else
			connstr.RollbackTrans
			Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			Response.Flush
		End If
		
		
			
		'------------------------------------------------------------------------------------------------------------------------------
		'-- Persediaan CPO PK 
		Dim CPOkirimshi,totcpohi,totcposbl
		Dim kernelkrmshi,totpkhi,totpksbl
		
		Dim PKS_RPBulanini
		      
		Dim CPO_QtyBlnini,PK_QtyBlnini
		Dim CPO_RendemenBlnini,PK_RendemenBlnini
		
		
		qQuery = "select sum(cpokirim) cpokirimshi, sum(kernelkirim) kernelkrmshi from productiontrn " & _ 
                 "where To_Number(To_Char(proddate,'mm'))=" & Month(Session("v_Mill_Production")) & " AND To_Number(To_Char(proddate,'yyyy'))= " & Year(Session("v_Mill_Production")) & " and fcba='" & Session("fcBA") & "'" 

		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			CPOkirimshi = cdbl(rsqLineMax("CPOkirimshi"))
			kernelkrmshi = cdbl(rsqLineMax("kernelkrmshi"))
		else
			CPOkirimshi = 0
			kernelkrmshi = 0
		end if 
		'response.write(qQuery & "<br>")
		
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		qQuery = " select (cpokirim+cpocot+cpocst+cpoot+cpotankisi) totcpohi ,( kernelsilo+pktankisi) totpkhi from productiontrn " & _ 
		         " WHERE fcba='" & Session("fcBA") & "' AND proddate=To_Date('" & LastMonth & "','mm/dd/yyyy') " 

		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			totcpohi = cdbl(rsqLineMax("totcpohi"))
			totpkhi = cdbl(rsqLineMax("totpkhi"))
		else
			totcpohi = 0
			totpkhi = 0
		end if 
		'response.write(qQuery & "<br>")
		
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		qQuery = " select ((cpocot) +(cpocst) +(cpoot) + (cpotankisi)) totcposbl ,((kernelsilo) + (pktankisi)) totpksbl from productiontrn " & _
                 " WHERE fcba='" & Session("fcBA") & "' AND proddate=trunc(To_Date('" & LastMonth & "','mm/dd/yyyy'),'month')-1 " 

		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			totcposbl = cdbl(rsqLineMax("totcposbl"))
			totpksbl = cdbl(rsqLineMax("totpksbl"))
		else
			totcposbl = 0
			totpksbl = 0
		end if 
		'response.write(qQuery & "<br>")
		
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		
		qQuery = "SELECT fcba,Sum(amount) amount " & _ 
				 " 	FROM " & _ 
				 " 	( " & _ 
				 "		SELECT a.fcba,jobgroupcode,Sum(a.amount) amount " & _ 
				 "		FROM cost_activity a " & _ 
				 "		INNER JOIN job b ON a.jobcode=b.fccode  AND a.fcba=b.fcba " & _ 
				 "		INNER JOIN jobgroup c ON  b.jobgroupcode=c.fccode AND b.fcba=c.fcba " & _ 
                 " 		WHERE a.fcba = '" & Session("fcBA")  & "'" & _ 
                 " 		AND b.jobgroupcode IN ('6110000','6210000','6310000','6400000','6610100','6620100') " & _ 
				 " 		AND To_Char(a.tdate,'yyyymm')='" & Year(Session("v_Mill_Production")) & "'||LPAD('" & Month(Session("v_Mill_Production")) & "', 2, 0) " & _
				 "		AND to_number(a.accountingperiod)=" & Month(Session("v_Mill_Production")) & "" &  _
			     "      AND to_number(a.accountingyear)=" & Year(Session("v_Mill_Production")) & "" &  _
			     " 		GROUP BY  a.fcba,jobgroupcode " & _ 
				 "  )x " & _
				 " GROUP BY fcba"
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qQuery, connstr
		if not rsqLineMax.eof then
			PKS_RPBulanini = cdbl(rsqLineMax("amount"))
		else
			PKS_RPBulanini = 0
		end if 
		'response.write(qQuery & "<br>")
		
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		
		 CPO_QtyBlnini = CPOkirimshi +  totcpohi - totcposbl
		 PK_QtyBlnini  = kernelkrmshi + totpkhi - totpksbl 
		 
		 CPO_RendemenBlnini = round((CPO_QtyBlnini/TBS_QtyOlah)*100,2)
		 PK_RendemenBlnini  = round((PK_QtyBlnini/TBS_QtyOlah)*100,2)
		 
		 Dim AvgCostAll,AvgCostCPO,AvgCostPK
		 Dim CostCPO_1,CostPK_1
		 Dim PsnCostCPO,PsnCostPK
		 Dim AmountCPO,AmountPK
		 
		 AvgCostAll  = round(PKS_RPBulanini / (CPO_QtyBlnini+PK_QtyBlnini),2)
		 AvgCostCPO  = round((CPO_RendemenBlnini/ (CPO_RendemenBlnini+PK_RendemenBlnini)) * AvgCostAll,2)
		 AvgCostPK   = round((PK_RendemenBlnini/ (CPO_RendemenBlnini+PK_RendemenBlnini)) * AvgCostAll,2)
		 
		 CostCPO_1 =   round(CPO_RendemenBlnini * AvgCostCPO,2)
		 CostPK_1  =   round(PK_RendemenBlnini  * AvgCostPK,2)
		 
		 PsnCostCPO = (CostCPO_1 / (CostCPO_1+CostPK_1))
		 PsnCostPK  = (CostPK_1 / (CostCPO_1+CostPK_1))
		 
		 AmountCPO = PsnCostCPO * PKS_RPBulanini
		 AmountPK  = PsnCostPK  * PKS_RPBulanini
		 
		 'response.write("TBS_QtyOlah : " &TBS_QtyOlah & "<br>")
		 'response.write("CPO_QtyBlnini : " &CPO_QtyBlnini & "<br>")
		 'response.write("PK_QtyBlnini : "& PK_QtyBlnini & "<br>")
		 
		 'response.write("CPO_RendemenBlnini :" & CPO_RendemenBlnini & "<br>")
		 'response.write("PK_RendemenBlnini : " & PK_RendemenBlnini & "<br>")
		 
		 'response.write("PKS_RPBulanini : " & PKS_RPBulanini & "<br>")
		 
		 'response.write("AvgCostAll : " & AvgCostAll & "<br>")
		 
		 'response.write("AvgCostCPO : " & AvgCostCPO & "<br>")
		 
		 'response.write("AvgCostPK : " & AvgCostPK & "<br>")
		 
		 
		 'response.write("PsnCostCPO : " & PsnCostCPO & "<br>")
		 'response.write("PsnCostPK : " & PsnCostPK & "<br>")
		 
		 'response.write("AmountCPO : " & AmountCPO & "<br>")
		 'response.write("AmountPK : " & AmountPK & "<br>")

		 'response.flush
		 
		'CPO
		connstr.BeginTrans
		maxLineNo = maxLineNo + 1
		qCost_Activity= "Insert into Cost_Activity " & _
						" ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
						" TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
						" QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE," & _
						" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
						" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
						" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
						" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
						" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
						" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
						" LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
						" SELECT To_Date('"& Session("v_Mill_Production") & "','mm/dd/yyyy') datein, '1140205' jobcode,'ST' source,'ST' SOURCETYPE,'POM' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						" 'CB' TARGETTYPE ,'BS' TARGETCODE,'','PKSOLAH'||'/'||'" & Month(Session("v_Mill_Production")) & "'||'/'||'" & Year(Session("v_Mill_Production")) &"', " & _ 
						" " & CPO_QtyBlnini & " QUANTITY, 0, " & AmountCPO & " amount,'" & Session("fcBA") & "',to_date('" & Session("v_Mill_Production") & "','mm/dd/yyyy'),sysdate, " & _ 
						" '','', " & _ 
						" '','', " & _
						" '','', " & _
						" '','', " & _
						" '','','Dr', " & _
						"" & Month(Session("v_Mill_Production")) & " , " & Year(Session("v_Mill_Production")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'POM' FROM DUAL " 
		
		connstr.execute(qCost_Activity)
							
		If Err.Number = 0 Then
			connstr.CommitTrans
			Response.Flush
		Else
			connstr.RollbackTrans
			Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			Response.Flush
		End If
		
		'PK
		
		connstr.BeginTrans
		maxLineNo = maxLineNo + 1
		qCost_Activity= "Insert into Cost_Activity " & _
						" ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
						" TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
						" QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE," & _
						" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
						" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
						" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
						" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
						" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
						" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
						" LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
						" SELECT To_Date('"& Session("v_Mill_Production") & "','mm/dd/yyyy') datein, '1140206' jobcode,'ST' source,'ST' SOURCETYPE,'POM' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						" 'CB' TARGETTYPE ,'BS' TARGETCODE,'','PKSOLAH'||'/'||'" & Month(Session("v_Mill_Production")) & "'||'/'||'" & Year(Session("v_Mill_Production")) &"', " & _ 
						" " & PK_QtyBlnini & " QUANTITY, 0, " & AmountPK & " amount,'" & Session("fcBA") & "',to_date('" & Session("v_Mill_Production") & "','mm/dd/yyyy'),sysdate, " & _ 
						" '','', " & _ 
						" '','', " & _
						" '','', " & _
						" '','', " & _
						" '','','Dr', " & _
						"" & Month(Session("v_Mill_Production")) & " , " & Year(Session("v_Mill_Production")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'POM' FROM DUAL " 
		
		connstr.execute(qCost_Activity)
							
		If Err.Number = 0 Then
			connstr.CommitTrans
			Response.Flush
		Else
			connstr.RollbackTrans
			Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			Response.Flush
		End If
		
	
		connstr.BeginTrans
		maxLineNo = maxLineNo + 1
		qCost_Activity= "Insert into Cost_Activity " & _
						" ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
						" TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
						" QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE," & _
						" TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
						" TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
						" TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
						" TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
						" TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
						" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
						" LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
						" SELECT To_Date('" & Session("v_Mill_Production") & "','mm/dd/yyyy') datein, '6699999' jobcode,'ST' source,'ST' SOURCETYPE,'POM' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						" 'CB' TARGETTYPE ,'BS' TARGETCODE,'','PKSOLAH'||'/'||'" & Month(Session("v_Mill_Production")) & "'||'/'||'" & Year(Session("v_Mill_Production")) &"', " & _ 
						" " & (CPO_QtyBlnini+PK_QtyBlnini) & " QUANTITY, 0, " & (AmountCPO+AmountPK)*-1 & " amount,'" & Session("fcBA") & "',to_date('" & Session("v_Mill_Production") & "','mm/dd/yyyy'),sysdate, " & _ 
						" '','', " & _ 
						" '','', " & _
						" '','', " & _
						" '','', " & _
						" '','','Cr', " & _
						"" & Month(Session("v_Mill_Production")) & " , " & Year(Session("v_Mill_Production")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'POM'  FROM DUAL" 
				
		connstr.execute(qCost_Activity)
							
		If Err.Number = 0 Then
			connstr.CommitTrans
			Response.Flush
		Else
			connstr.RollbackTrans
			Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			Response.Flush
		End If
		
		
		qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT referenceno,journalprefix,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Mill_Production")) & " AND accountingyear=" & Year(Session("v_Mill_Production")) & " AND journalprefix='POM' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.referenceno =  y.referenceno and x.journalprefix=y.journalprefix) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
		
		connstr.execute(qjournal)
	
		 

		Response.write("<font face=Verdana size=1 color=Green>Transfer to Activity Based Costing..., Success!.<br/><br/></font>")
		Response.Flush
		 

	
	End IF	
		
		
						Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear,temp_EndingAccountingPeriod
							
						QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
											" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
											" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
											" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear,  " & _ 
											" periodcontrol_api.GetEndingAccountingPeriodTY(system,'" & Session("fcBA") & "')  EndingAccountingPeriod " & _ 
											" from periodcontrol p  " & _ 
											" where UPPER(system) ='MILL_PRODUCTION' and fcba='" & Session("fcBA") & "'"
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
								Response.Write  "<font face=Verdana size=1 color=red>" & "IPLAS Error, No PERIODCONTROL for MILL_PRODUCTION " & "<br>"
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
							   " where UPPER(system) ='MILL_PRODUCTION' and fcba='" & Session("fcBA") & "'"
						ELSE
							IF temp_NextCloseAccountingPeriod < 12 THEN
								q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
								  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
								  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
								  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
								  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
								  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
								  " where UPPER(system) ='MILL_PRODUCTION' and fcba='" & Session("fcBA") & "'"
							ELSE
								 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
								   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
								  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
								  " where  UPPER(system) ='MILL_PRODUCTION' and fcba='" & Session("fcBA") & "'"
							END IF
						END IF
						
						IF temp_NextCloseAccountingPeriod = temp_EndingAccountingPeriod AND temp_NextCurrentAccountingPeriod = 1  THEN
						
							q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
							  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
							  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							  " where UPPER(system) ='MILL_PRODUCTION' and and fcba='" & Session("fcBA") & "'"
						END IF
			
			
		
        connStr.execute(q)
		
		'**** Remark By ukie 1 Februari 2016 ****'
		'connStr.execute(q)
		
		'**** Add By ukie 1 Februari 2016 ****'
		txtLastDay  = cmdFindPeriod("Mill_Production")
		txtNextMonth = cmdPeriodClose("Mill_Production")
			
		Session("v_Mill_Production") = txtLastDay
		Session("cls_Mill_Production") = txtNextMonth	
		
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
		
	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Mill_Production'" 
								
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.CursorType = 3
	rsqLineMax.CursorLocation = 3
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 
		
		rsqLineMax.close
		set rsqLineMax = Nothing
		
		

		tempMaxLine = tempMaxLine + 1
		
		
	strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME, fcentry) VALUES (" & _
						" 'Mill_Production', " & _
						" '" & Session("fcBA") & "'," & _
						" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
						" " & tempMaxLine & _
						" , SYSDATE, to_char(SYSDATE,'HH24:MI') " & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Mill_Production'"
				
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.CursorType = 3
	rsLenghtime.CursorLocation = 3
	rsLenghtime.Open strLenghtime, connstr

	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Mill_Production'"		
    				
	connStr.execute(upClosingLog)
	'--End Added--
	rsLenghtime.close
	set rsLenghtime = Nothing
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
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Mill Production Closing is succesfully...
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
	<br><br><img src="images/error.gif">Mill Production Closing is NOT completed...
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