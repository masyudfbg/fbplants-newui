<%'Option Explicit
%>
<%On Error Resume Next
%>
<%response.buffer=True
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>iPS - Weighbridge Closing</title>
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
WeighBridge Last Period Close <%=FormatDateTime(Session("cls_WeighBridge"),1)%> <br> 
WeighBridge Current Period is <%=FormatDateTime(Session("v_WeighBridge"),1)%> <br><br>
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

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('WeighBridge','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for WeighBridge " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then 


	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  " where SYSTEM='WeighBridge' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
		temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for WeighBridge " & "<br>"
	end if
	rsPeriodControl.close
    set rsPeriodControl = Nothing

	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX





	'---- Control Job
	qControlJob="select ControlJob from Module where NAME='WeighBridge'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For WeighBridge is Empty"
	   txtControlJob=""
	end if
	rsControlJob.close
    set rsControlJob = Nothing
	response.Flush()

	
'--	List Proses Journal Auto
'--1. Journal Persediaan TBS INTERNAL (SFB-Stock Fruit Bunch) 
'--   PERSEDIAAN TANDAN BUAH SEGAR (TBS) KEBUN INTI	(1140201) (D)
'--   INVOICE TBS INTERNAL YANG BELUM DITERIMA (2110608) (K)
'
'--2. Journal Persediaan TBS EXTERNAL (SFB-Stock Fruit Bunch) 
'--   PERSEDIAAN TANDAN BUAH SEGAR (TBS) EXTERNAL	(1140203) (D)
'--   INVOICE TBS EXTERNAL YANG BELUM DITERIMA (2110609) (K)
'
'--5. Journal TRANSPORTIER PRODUCT/SHIPMENT (CTR-Cost Transporter)
'--   Biaya Angkutan/Ansurasi CPO (7110101) (D)
'--   BIAYA ASURANSI & ANGKUTAN KERNEL (7110105) (D)
'--   BIAYA ASURANSI/ANGKUTAN CANGKANG (7110109) (D)
'--   Invoice blm di tagih (2110607) (K)
'



		'---- Normalisasi
	if Session("Estate_Mill") = "M" then
		'connstr.BeginTrans

'--1. Journal Persediaan TBS INTERNAL (SFB-Stock Fruit Bunch) 
'--2. Journal Persediaan TBS EXTERNAL (SFB-Stock Fruit Bunch) 

		qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_WeighBridge")) &  _
					 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_WeighBridge")) & " and fcBA='" & Session("fcBA") & "'" & _
					 " and journalprefix='SFB'"
		connstr.execute(qNormalisasiDetail)	
		
		'If Err.Number = 0 Then
		'	connstr.CommitTrans
			Response.Write  "<font face=Verdana size=1 color=Blue>" & "Normalisasi For Persediaan TBS ,success " & "<br/></font>" 		
		'Else
		'	connstr.RollbackTrans
		'	Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
			Response.Flush
		'End If
		
		
	    
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
				
		
		'connstr.BeginTrans
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
						" SELECT datein, decode(own_estate,'Y','1140201','1140203') jobcode,'ST' source,'ST' SOURCETYPE,'SFB' SOURCECODE,customercode SOURCECODE_STATUS , " & _ 
						" 'CB' TARGETTYPE ,'BS' TARGETCODE,'',customercode||'/'||to_char(datein,'yyyymmdd'), " & _ 
						" netto2 QUANTITY, 0, netto2*Nvl(price,0) amount,'" & Session("fcBA") & "',to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'),sysdate, " & _ 
						" 'TC - Balance Sheets :','BS', " & _ 
						" '',customercode||'/'||to_char(datein,'yyyymmdd'), " & _
						" 'Customer TBS :',customername, " & _
						" 'own_estate',decode(own_estate,'Y','OWN','EXT'), " & _
						" '','','Dr', " & _
						"" & Month(Session("v_WeighBridge")) & " , " & Year(Session("v_WeighBridge")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'SFB' " & _
						" FROM " & _
						" ( " & _
						"	SELECT own_estate ,c.eplascode customercode,d.description customername,docketdate datein,Sum(a.ffbweight) netto2 " & _
						"	FROM lhpffbwbtrn a " & _
						"	left JOIN lhptikethdr b ON a.docketid=b.docketid " & _
						"	left JOIN lhpareamst c ON  a.areaid=c.areaid " & _
						"	left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
						"	WHERE  To_Char(docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0)" & _
                        "			AND b.fcba='" & Session("fcBA") & "' " & _
						"	GROUP BY own_estate ,c.eplascode ,d.description ,docketdate " & _
						" )x " & _
						" LEFT JOIN Invoice_pricetbs y ON x.customercode=y.suppliercode AND x.datein=y.ondate " & _
						" ORDER BY datein,customername "
		
		connstr.execute(qCost_Activity)
							
		'If Err.Number = 0 Then
		'	connstr.CommitTrans
		'	Response.Flush
		'Else
		'	connstr.RollbackTrans
		'	Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
		'	Response.Flush
		'End If
		
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
		
		'connstr.BeginTrans
		maxLineNo = maxLineNo + 1
		qCost_Activity ="Insert into Cost_Activity " & _
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
						" SELECT to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'), decode(own_estate,'Y','2110608','2110609') jobcode,'AP' Source,'AP' SOURCETYPE,'SFB',decode(own_estate,'Y','OWN','EXT') SOURCECODE_STATUS , " & _ 
						" 'CB' TARGETTYPE ,'BS' TARGETCODE,'','', " & _ 
						" 0 QUANTITY, 0, amount * -1 ,'" & Session("fcBA") & "',to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'),sysdate, " & _ 
						" '','', " & _ 
						" '',decode(own_estate,'Y','OWN','EXT'), " & _
						" '','', " & _
						" 'own_estate',decode(own_estate,'Y','OWN','EXT'), " & _
						" '','','Cr', " & _
						"" & Month(Session("v_WeighBridge")) & " , " & Year(Session("v_WeighBridge")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'SFB' " & _
						" FROM " & _
						" ( SELECT own_estate,SUM(amount) amount FROM " & _
						" 	( " & _
						" 	SELECT own_estate,netto2*Nvl(price,0) amount " & _ 
						" 	FROM " & _
						" 	( " & _
						"		SELECT own_estate ,c.eplascode customercode,d.description customername,docketdate datein,Sum(a.ffbweight) netto2 " & _
						"		FROM lhpffbwbtrn a " & _
						"		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _
						"		left JOIN lhpareamst c ON  a.areaid=c.areaid " & _
						"		left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
						"		WHERE  To_Char(docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0)" & _
                        "			AND b.fcba='" & Session("fcBA") & "' " & _
						"		GROUP BY own_estate ,c.eplascode ,d.description ,docketdate " & _
						" 	)x " & _
						" 	LEFT JOIN Invoice_pricetbs y ON x.customercode=y.suppliercode AND x.datein=y.ondate " & _
						" 	)Z GROUP BY own_estate " & _ 
						" )j "
						
		
		connstr.execute(qCost_Activity)
							
		'If Err.Number = 0 Then
		'	connstr.CommitTrans
		'	Response.Flush
		'Else
		'	connstr.RollbackTrans
		'	Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
		'	Response.Flush
		'End If
		
		qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT To_Char(tdate,'yyyymm')||AUTO_REFERENCE_04 referenceno,journalprefix,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_WeighBridge")) & " AND accountingyear=" & Year(Session("v_WeighBridge")) & " AND journalprefix='SFB' " & _
					"  Order BY To_Char(tdate,'yyyymm')||AUTO_REFERENCE_04 " & _
					"  ) " & _
					" )y ON (To_Char(x.tdate,'yyyymm')||x.AUTO_REFERENCE_04 =  y.referenceno and x.journalprefix=y.journalprefix) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
		connstr.execute(qjournal)
		
		Response.Write  "<font face=Verdana size=1 color=Blue>" & "Prosses Persediaan TBS, success " & "<br/></font>" 
		Response.Flush
		

' '--3. Journal Penjualan TBS INTERNAL (ARF-Account Receiveable FruitBunch) 
		' Dim strCursorSivCode, rsCursorSivCode
        ' Dim qCostAMax,rsCostAMax
		
		' qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_WeighBridge")) &  _
					 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_WeighBridge")) & " and journalprefix='ARF'"
		' connstr.execute(qNormalisasiDetail)	
		
		' qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_WeighBridge")) &  _
					 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_WeighBridge")) & " and journalprefix='HPF'" 
		' connstr.execute(qNormalisasiDetail)	
		
		' Response.Write  "<font face=Verdana size=1 color=Blue>" & "Normalisasi For Penjualan TBS Internal & HPP TBS ,success " & "<br/></font>" 
		
		 ' strCursorSivCode = "SELECT customer_fcba,c.eplascode customercode,Sum(a.ffbweight) netto2 " & _
						    ' " FROM lhpffbwbtrn a " & _
						    ' " left JOIN lhptikethdr b ON a.docketid=b.docketid " & _
						    ' " left JOIN lhpareamst c ON  a.areaid=c.areaid " & _
						    ' " left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
						    ' " WHERE  To_Char(docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0)" & _
                            ' " AND b.fcba='" & Session("fcBA") & "' AND own_estate='Y' " & _
					        ' " GROUP BY c.eplascode ,customer_fcba "

						
		' Set rsCursorSivCode = Server.CreateObject("ADODB.Recordset")
		 ' rsCursorSivCode.CursorType = 3
		 ' rsCursorSivCode.CursorLocation = 3
		 ' rsCursorSivCode.Open strCursorSivCode, connstr		

		 ' if not rsCursorSivCode.eof then
			 ' Do While Not rsCursorSivCode.eof
			 
		
		
		 ' qCostAMax =  "select nvl(max(lineno), 0) maxLineNo from cost_activity  where fcba = '" & rsCursorSivCode("customer_fcba") & "'" 
		 
		 ' Set rsCostAMax = Server.CreateObject("ADODB.Recordset")
		 ' rsCostAMax.Open qCostAMax, connstr
		 ' If Not rsCostAMax.EOF Then
		 ' maxLNO = cdbl(rsCostAMax("maxLineNo"))
		 ' Else
		 ' maxLNO = 0
		 ' End If
		
		 ' rsCostAMax.close
		 ' set rsCostAMax = Nothing
		
		 ' maxLNO = maxLNO + 1
		
		' Response.Write  "<font face=Verdana size=1 color=Green>" & "Prosses Penjualan TBS Internal " & rsCursorSivCode("customer_fcba") & "<br/></font>" 
		' Response.Flush
				
		 ' connstr.BeginTrans
		 ' qCost_Activity ="Insert into Cost_Activity " & _
						 ' " ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
						 ' " TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
						 ' " QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE," & _
						 ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
						 ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
						 ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
						 ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
						 ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
						 ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
						 ' " LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
						 ' " SELECT to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'), decode(own_estate,'Y','2110608','2110609') jobcode,'AP' Source,'AP' SOURCETYPE,'SFB',decode(own_estate,'Y','OWN','EXT') SOURCECODE_STATUS , " & _ 
						 ' " 'CB' TARGETTYPE ,'BS' TARGETCODE,'','', " & _ 
						 ' " 0 QUANTITY, 0, amount ,'" & rsCursorSivCode("customer_fcba") & "',to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'),sysdate, " & _ 
						 ' " '','', " & _ 
						 ' " '',decode(own_estate,'Y','OWN','EXT'), " & _
						 ' " '','', " & _
						 ' " 'own_estate',decode(own_estate,'Y','OWN','EXT')||'" & rsCursorSivCode("customercode") & "', " & _
						 ' " '','','Dr', " & _
						 ' "" & Month(Session("v_WeighBridge")) & " , " & Year(Session("v_WeighBridge")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'ARF' " & _
						 ' " FROM " & _
						 ' " ( SELECT customer_fcba,own_estate,SUM(amount) amount FROM " & _
						 ' " 	( " & _
						 ' " 	SELECT customer_fcba,own_estate,netto2*Nvl(price,0) amount " & _ 
						 ' " 	FROM " & _
						 ' " 	( " & _
						 ' "		SELECT customer_fcba,own_estate ,c.eplascode customercode,d.description customername,docketdate datein,Sum(a.ffbweight) netto2 " & _
						 ' "		FROM lhpffbwbtrn a " & _
						 ' "		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _
						 ' "		left JOIN lhpareamst c ON  a.areaid=c.areaid " & _
						 ' "		left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
						 ' "		WHERE  To_Char(docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0)" & _
                         ' "			AND b.fcba='" & Session("fcBA") & "' AND own_estate='Y' AND customer_fcba = '" & rsCursorSivCode("customer_fcba") & "'" & _
						 ' "		GROUP BY own_estate ,c.eplascode ,d.description ,docketdate ,customer_fcba" & _
						 ' " 	)x " & _
						 ' " 	LEFT JOIN Invoice_pricetbs y ON x.customercode=y.suppliercode AND x.datein=y.ondate " & _
						 ' " 	)Z GROUP BY own_estate,customer_fcba " & _ 
						 ' " )j "
						
		
		 ' connstr.execute(qCost_Activity)
							
		 ' If Err.Number = 0 Then
			 ' connstr.CommitTrans
			 ' Response.Flush
		 ' Else
			 ' connstr.RollbackTrans
			 ' Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			 ' Response.Flush
		 ' End If
		
		 ' maxLNO = maxLNO + 1
		 ' connstr.BeginTrans
		 ' qCost_Activity ="Insert into Cost_Activity " & _
						 ' " ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
						 ' " TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _
						 ' " QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE," & _
						 ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
						 ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
						 ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
						 ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
						 ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
						 ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
						 ' " LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _
						 ' " SELECT to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'), 5110106 jobcode,'AR' Source,'AR' SOURCETYPE,'" & rsCursorSivCode("customercode") & "',decode(own_estate,'Y','OWN','EXT') SOURCECODE_STATUS , " & _ 
						 ' " 'CB' TARGETTYPE ,'BS' TARGETCODE,'','', " & _ 
						 ' " 0 QUANTITY, 0, amount * -1 ,'" & rsCursorSivCode("customer_fcba") & "',to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'),sysdate, " & _ 
						 ' " '','', " & _ 
						 ' " '',decode(own_estate,'Y','OWN','EXT'), " & _
						 ' " '','', " & _
						 ' " 'own_estate',decode(own_estate,'Y','OWN','EXT')||'" & rsCursorSivCode("customercode") & "', " & _
						 ' " '','','Cr', " & _
						 ' "" & Month(Session("v_WeighBridge")) & " , " & Year(Session("v_WeighBridge")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'ARF' " & _
						 ' " FROM " & _
						 ' " ( SELECT customer_fcba,own_estate,SUM(amount) amount FROM " & _
						 ' " 	( " & _
						 ' " 	SELECT customer_fcba,own_estate,netto2*Nvl(price,0) amount " & _ 
						 ' " 	FROM " & _
						 ' " 	( " & _
						 ' "		SELECT customer_fcba,own_estate ,c.eplascode customercode,d.description customername,docketdate datein,Sum(a.ffbweight) netto2 " & _
						 ' "		FROM lhpffbwbtrn a " & _
						 ' "		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _
						 ' "		left JOIN lhpareamst c ON  a.areaid=c.areaid " & _
						 ' "		left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
						 ' "		WHERE  To_Char(docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0)" & _
                         ' "			AND b.fcba='" & Session("fcBA") & "' AND own_estate='Y' AND customer_fcba = '" & rsCursorSivCode("customer_fcba") & "'" & _
						 ' "		GROUP BY own_estate ,c.eplascode ,d.description ,docketdate ,customer_fcba" & _
						 ' " 	)x " & _
						 ' " 	LEFT JOIN Invoice_pricetbs y ON x.customercode=y.suppliercode AND x.datein=y.ondate " & _
						 ' " 	)Z GROUP BY own_estate,customer_fcba " & _ 
						 ' " )j "
		 ' connstr.execute(qCost_Activity)
							
		 ' If Err.Number = 0 Then
			 ' connstr.CommitTrans
			 ' Response.Flush
		 ' Else
			 ' connstr.RollbackTrans
			 ' Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			 ' Response.Flush
		 ' End If
		 
		 
' '--4. Journal HPP TBS INTERNAL (HPF-Harga Pokok  FruitBunch) 
		 
		' Response.Write  "<font face=Verdana size=1 color=Green>" & "Prosses HPP TBS Internal " & rsCursorSivCode("customer_fcba") & "<br/></font>" 
		' Response.Flush
		
		  ' maxLNO = maxLNO + 1
		 ' connstr.BeginTrans
		 ' qCost_Activity =" Insert into Cost_Activity " & _ 
						 ' " ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _ 
						 ' " TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _ 
						 ' " QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, " & _ 
						 ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _ 
						 ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _ 
						 ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _ 
						 ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _ 
						 ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _ 
						 ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _ 
						 ' " LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _ 
						 ' " SELECT to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'), '5210201' jobcode,'PL' Source,'PL' SOURCETYPE,'HPP' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						 ' " 'CB' TARGETTYPE ,'BS' TARGETCODE,'' TARGETCODE_STATUS,'HPP/TBS/'||fcba REFERENCENO, " & _ 
						 ' " 0 QUANTITY, 0, amount ,fcba,to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'),sysdate, " & _ 
						 ' " '','', " & _ 
						 ' " '','', " & _ 
						 ' " '','', " & _ 
						 ' " '','HPP/TBS/'||fcba, " & _ 
						 ' " '','','Dr', " & _ 
						 ' "" & Month(Session("v_WeighBridge")) & " , " & Year(Session("v_WeighBridge")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'HPT' " & _
						 ' " FROM " & _ 
						 ' " (	SELECT fcba,Sum(amount) amount " & _ 
						 ' " 	FROM " & _ 
						 ' " 	( " & _ 
						 ' "		SELECT a.fcba,jobgroupcode,Sum(a.amount) amount " & _ 
						 ' "		FROM cost_activity a " & _ 
						 ' "		INNER JOIN job b ON a.jobcode=b.fccode  AND a.fcba=b.fcba " & _ 
						 ' "		INNER JOIN jobgroup c ON  b.jobgroupcode=c.fccode AND b.fcba=c.fcba " & _ 
                         ' " 		WHERE a.fcba = '" & rsCursorSivCode("customer_fcba") & "'" & _ 
                         ' " 		AND b.jobgroupcode IN ('6110000','6210000','6310000','6500100','6500200','6500300','6500400','6500500','6500600','6500700','6500800','6500900','6501000','6502100','6502200','6400000','6509100','6510100') " & _ 
						 ' " 		AND To_Char(a.tdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0) " & _
						 ' " 		GROUP BY  a.fcba,jobgroupcode " & _ 
						 ' "  )x " & _
						 ' " GROUP BY fcba " & _
						 ' " )j "

                  
          		 ' connstr.execute(qCost_Activity)
							
				 ' If Err.Number = 0 Then
					 ' connstr.CommitTrans
					 ' Response.Flush
				 ' Else
					 ' connstr.RollbackTrans
					 ' Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
					 ' Response.Flush
				 ' End If


 		 ' maxLNO = maxLNO + 1
		 ' connstr.BeginTrans
		 ' qCost_Activity =" Insert into Cost_Activity " & _ 
						 ' " ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _ 
						 ' " TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _ 
						 ' " QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, " & _ 
						 ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _ 
						 ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _ 
						 ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _ 
						 ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _ 
						 ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _ 
						 ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _ 
						 ' " LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _ 
						 ' " SELECT to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'), '6599999' jobcode,'PL' Source,'PL' SOURCETYPE,'HPP' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						 ' " 'CB' TARGETTYPE ,'BS' TARGETCODE,'' TARGETCODE_STATUS,'HPP/TBS/'||fcba REFERENCENO, " & _ 
						 ' " 0 QUANTITY, 0, amount *-1,fcba,to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'),sysdate, " & _ 
						 ' " '','', " & _ 
						 ' " '','', " & _ 
						 ' " '','', " & _ 
						 ' " '','HPP/TBS/'||fcba, " & _ 
						 ' " '','','Cr', " & _ 
						 ' "" & Month(Session("v_WeighBridge")) & " , " & Year(Session("v_WeighBridge")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'HPT' " & _
						 ' " FROM " & _ 
						 ' " (	SELECT fcba,Sum(amount) amount " & _ 
						 ' " 	FROM " & _ 
						 ' " 	( " & _ 
						 ' "		SELECT a.fcba,jobgroupcode,Sum(a.amount) amount " & _ 
						 ' "		FROM cost_activity a " & _ 
						 ' "		INNER JOIN job b ON a.jobcode=b.fccode  AND a.fcba=b.fcba " & _ 
						 ' "		INNER JOIN jobgroup c ON  b.jobgroupcode=c.fccode AND b.fcba=c.fcba " & _ 
                         ' " 		WHERE a.fcba = '" & rsCursorSivCode("customer_fcba") & "'" & _ 
                         ' " 		AND b.jobgroupcode IN ('6110000','6210000','6310000','6500100','6500200','6500300','6500400','6500500','6500600','6500700','6500800','6500900','6501000','6502100','6502200','6400000','6509100','6510100') " & _ 
						 ' " 		AND To_Char(a.tdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0) " & _
						 ' " 		GROUP BY  a.fcba,jobgroupcode " & _ 
						 ' "  )x " & _
						 ' " GROUP BY fcba " & _
						 ' " )j "

          		 ' connstr.execute(qCost_Activity)
							
				 ' If Err.Number = 0 Then
					 ' connstr.CommitTrans
					 ' Response.Flush
				 ' Else
					 ' connstr.RollbackTrans
					 ' Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
					 ' Response.Flush
				 ' End If

				
			 ' rsCursorSivCode.MoveNext
			 ' loop
		 ' END IF
	     ' rsCursorSivCode.close
		 ' set rsCursorSivCode=Nothing
		
		
		         ' qjournal=	" MERGE INTO cost_activity x " & _
					 ' " USING " & _
					 ' " ( " & _
					 ' "  SELECT referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					 ' "  ( " & _
					 ' "  SELECT DISTINCT AUTO_REFERENCE_04 referenceno,journalprefix,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					 ' "  FROM cost_activity WHERE accountingperiod="& Month(Session("v_WeighBridge")) & " AND accountingyear=" & Year(Session("v_WeighBridge")) & " AND journalprefix='ARF' " & _
					 ' "  Order BY referenceno " & _
					 ' "  ) " & _
					 ' " )y ON (x.AUTO_REFERENCE_04 =  y.referenceno and x.journalprefix=y.journalprefix and x.accountingperiod="& Month(Session("v_WeighBridge")) & " AND x.accountingyear=" & Year(Session("v_WeighBridge")) & ") " & _
					 ' " WHEN MATCHED THEN " & _
					 ' " UPDATE SET x.journalno = y.seq " 
		
				 ' connstr.execute(qjournal)
		
		         ' Response.Write  "<font face=Verdana size=1 color=Blue>" & "Prosses Penjualan TBS Internal, success " & "<br/></font>" 
				 ' Response.Flush
		
		         ' qjournal=	" MERGE INTO cost_activity x " & _
					 ' " USING " & _
					 ' " ( " & _
					 ' "  SELECT referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					 ' "  ( " & _
					 ' "  SELECT DISTINCT AUTO_REFERENCE_04 referenceno,journalprefix,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					 ' "  FROM cost_activity WHERE accountingperiod="& Month(Session("v_WeighBridge")) & " AND accountingyear=" & Year(Session("v_WeighBridge")) & " AND journalprefix='HPT' " & _
					 ' "  Order BY referenceno " & _
					 ' "  ) " & _
					 ' " )y ON (x.AUTO_REFERENCE_04 =  y.referenceno and x.journalprefix=y.journalprefix and x.accountingperiod="& Month(Session("v_WeighBridge")) & " AND x.accountingyear=" & Year(Session("v_WeighBridge")) & ") " & _
					 ' " WHEN MATCHED THEN " & _
					 ' " UPDATE SET x.journalno = y.seq " 
		
				 ' connstr.execute(qjournal)
		
		         ' Response.Write  "<font face=Verdana size=1 color=Blue>" & "Prosses HPP TBS Internal, success " & "<br/></font>" 
				 ' Response.Flush
				 
'--5. Journal TRANSPORTIER PRODUCT/SHIPMENT (CTR-Cost Transporter)
'--   Biaya Angkutan/Ansurasi CPO (7110101) (D)
'--   BIAYA ASURANSI & ANGKUTAN KERNEL (7110105) (D)
'--   BIAYA ASURANSI/ANGKUTAN CANGKANG (7110109) (D)
'--   Invoice blm di tagih (2110607) (K)

        'connstr.BeginTrans
		qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_WeighBridge")) &  _
					 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_WeighBridge")) & " and fcBA='" & Session("fcBA") & "'" & _
					 " and journalprefix='CTR'"
		connstr.execute(qNormalisasiDetail)	
		
		'If Err.Number = 0 Then
		'	connstr.CommitTrans
		'Else
		'	connstr.RollbackTrans
		'	Response.write("Error Save ["&qNormalisasiDetail&"] ("&Err.Description&")")
		'	Response.Flush
		'End If
		
		
	    Response.Write  "<font face=Verdana size=1 color=Blue>" & "Normalisasi For Transportir Product ,success " & "<br/></font>" 
		
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

		 
	     maxLineNo = maxLineNo + 1
		' connstr.BeginTrans
		 qCost_Activity =" Insert into Cost_Activity " & _ 
						 " ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _ 
						 " TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _ 
						 " QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, " & _ 
						 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _ 
						 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _ 
						 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _ 
						 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _ 
						 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _ 
						 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _ 
						 " LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _ 
						 " SELECT to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'), case productid when '1' then '7110101' when '3' then '7110105' when '5' then '7110109' end jobcode,'CA' Source,'CA' SOURCETYPE,transporter SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						 " 'CB' TARGETTYPE ,'BS' TARGETCODE,'' TARGETCODE_STATUS,transporter||'-'||transportercontractno REFERENCENO, " & _ 
						 " Qty QUANTITY, 0, amount ,'"&  Session("fcBA") &"' fcba,to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'),sysdate, " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '',trim(transportername)||' DO:'||trim(deliveryorderno)||' Qty:'||Qty||' Price:'||transportationcost, " & _ 
						 " '','','Dr', " & _ 
						 "" & Month(Session("v_WeighBridge")) & " , " & Year(Session("v_WeighBridge")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'CTR' " & _
						 " FROM " & _ 
						" ( " & _ 
						"		SELECT " & _ 
						"		productid,Nvl(salescontractno,'-') salescontractno ,Nvl(deliveryorderno,doorderid) deliveryorderno , " & _
						" 		Nvl(transportercontractno,'-') transportercontractno, " & _
						" 		Nvl(transporter,'-') transporter , " & _
						" 		Nvl(transportername,'-') transportername, " & _
						" 		Nvl(transportationcost,1) transportationcost, " & _
						" 		Sum(netweight) Qty,Sum(netweight*Nvl(transportationcost,1)) Amount " & _ 
						"		FROM " & _ 
						"		( " & _ 
						"		SELECT  c.productid, a.docketid,a.netweight,b.docketdate,a.doid,c.doorderid, " & _ 
						"				d.salescontractno,d.deliveryorderno, " & _ 
						"				d.transportercontractno,d.transporter,transportationcost,losstolerance,Upper(e.contractorname) transportername " & _ 
						"		FROM lhptiketcpo a " & _ 
						"		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _ 
						"		left JOIN lhpdomst c ON a.doid=c.doid AND b.fcba=c.fcba " & _ 
						"		left JOIN deliveryorder d ON Trim(c.doorderid)=Trim(d.deliveryorderno)  AND Trim(c.fcba)=Trim(d.fcba) " & _ 
						"		Left JOIN lhpcontractormst e ON d.transporter=e.eplascode " & _ 
						"		WHERE b.fcba='"&  Session("fcBA") &"' AND To_Char(b.docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0) " & _
						"		UNION ALL " & _ 
						"		SELECT  c.productid, a.docketid,a.netweight,b.docketdate,a.doid,c.doorderid, " & _ 
						"				d.salescontractno,d.deliveryorderno, " & _ 
						"				d.transportercontractno,d.transporter,transportationcost,losstolerance,Upper(e.contractorname) transportername " & _ 
						"		FROM lhptiketkernel a " & _ 
						"		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _ 
						"		left JOIN lhpdomst c ON a.doid=c.doid AND b.fcba=c.fcba " & _ 
						"		left JOIN deliveryorder d ON Trim(c.doorderid)=Trim(d.deliveryorderno)  AND Trim(c.fcba)=Trim(d.fcba) " & _ 
						"		Left JOIN lhpcontractormst e ON d.transporter=e.eplascode " & _ 
						"		WHERE b.fcba='"&  Session("fcBA") &"' AND To_Char(b.docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0) " & _
						"		UNION ALL " & _ 
						"		SELECT  c.productid, a.docketid,a.netweight,b.docketdate,a.doid,c.doorderid, " & _ 
						"				d.salescontractno,d.deliveryorderno, " & _ 
						"				d.transportercontractno,d.transporter,transportationcost,losstolerance,Upper(e.contractorname) transportername " & _ 
						"		FROM lhptiketckg a  " & _ 
						"		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _ 
						"		left JOIN lhpdomst c ON a.doid=c.doid AND b.fcba=c.fcba " & _ 
						"		left JOIN deliveryorder d ON Trim(c.doorderid)=Trim(d.deliveryorderno)  AND Trim(c.fcba)=Trim(d.fcba) " & _ 
						"		Left JOIN lhpcontractormst e ON d.transporter=e.eplascode " & _ 
						"		WHERE b.fcba='"&  Session("fcBA") &"' AND To_Char(b.docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0) " & _
						"		)X " & _ 
						"		GROUP BY productid,salescontractno,deliveryorderno,doorderid,transportercontractno,transporter ,transportername ,transportationcost" & _ 
						" )j "

                  
          		 connstr.execute(qCost_Activity)
							
		'		 If Err.Number = 0 Then
		'			 connstr.CommitTrans
		'			 Response.Flush
		'		 Else
		'			 connstr.RollbackTrans
		'			 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
		'			 Response.Flush
		'		 End If

		
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
		
 		 maxLineNo = maxLineNo + 1
		' connstr.BeginTrans
		 qCost_Activity =" Insert into Cost_Activity " & _ 
						 " ( TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _ 
						 " TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, " & _ 
						 " QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, " & _ 
						 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _ 
						 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _ 
						 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _ 
						 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _ 
						 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _ 
						 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _ 
						 " LASTUPDATE, LASTTIME, LINENO,journalprefix ) " & _ 
						 " SELECT to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'), '2110607' jobcode,'CA' Source,'CA' SOURCETYPE,transporter SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						 " 'AP' TARGETTYPE ,transporter TARGETCODE,'' TARGETCODE_STATUS,transporter||'-'||transportercontractno REFERENCENO, " & _ 
						 " Qty QUANTITY, 0, amount *-1,'"&  Session("fcBA") &"' fcba,to_date('" & Session("v_WeighBridge") & "','mm/dd/yyyy'),sysdate, " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '',trim(transportername)||' DO:'||trim(deliveryorderno)||' Qty:'||Qty||' Price:'||transportationcost, " & _ 
						 " '','','Cr', " & _ 
						 "" & Month(Session("v_WeighBridge")) & " , " & Year(Session("v_WeighBridge")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLineNo & "+rownum ,'CTR' " & _
						 " FROM " & _ 
						" ( " & _ 
						"		SELECT " & _ 
						"		productid,Nvl(salescontractno,'-') salescontractno ,Nvl(deliveryorderno,doorderid) deliveryorderno , " & _
						" 		Nvl(transportercontractno,'-') transportercontractno, " & _
						" 		Nvl(transporter,'-') transporter , " & _
						" 		Nvl(transportername,'-') transportername, " & _
						" 		Nvl(transportationcost,1) transportationcost, " & _
						" 		Sum(netweight) Qty,Sum(netweight*Nvl(transportationcost,1)) Amount " & _ 
						"		FROM " & _ 
						"		( " & _ 
						"		SELECT  c.productid, a.docketid,a.netweight,b.docketdate,a.doid,c.doorderid, " & _ 
						"				d.salescontractno,d.deliveryorderno, " & _ 
						"				d.transportercontractno,d.transporter,transportationcost,losstolerance,Upper(e.contractorname) transportername " & _ 
						"		FROM lhptiketcpo a " & _ 
						"		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _ 
						"		left JOIN lhpdomst c ON a.doid=c.doid AND b.fcba=c.fcba " & _ 
						"		left JOIN deliveryorder d ON Trim(c.doorderid)=Trim(d.deliveryorderno)  AND Trim(c.fcba)=Trim(d.fcba) " & _ 
						"		Left JOIN lhpcontractormst e ON d.transporter=e.eplascode " & _ 
						"		WHERE b.fcba='"&  Session("fcBA") &"' AND To_Char(b.docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0) " & _
						"		UNION ALL " & _ 
						"		SELECT  c.productid, a.docketid,a.netweight,b.docketdate,a.doid,c.doorderid, " & _ 
						"				d.salescontractno,d.deliveryorderno, " & _ 
						"				d.transportercontractno,d.transporter,transportationcost,losstolerance,Upper(e.contractorname) transportername " & _ 
						"		FROM lhptiketkernel a " & _ 
						"		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _ 
						"		left JOIN lhpdomst c ON a.doid=c.doid AND b.fcba=c.fcba " & _ 
						"		left JOIN deliveryorder d ON Trim(c.doorderid)=Trim(d.deliveryorderno)  AND Trim(c.fcba)=Trim(d.fcba) " & _ 
						"		Left JOIN lhpcontractormst e ON d.transporter=e.eplascode " & _ 
						"		WHERE b.fcba='"&  Session("fcBA") &"' AND To_Char(b.docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0) " & _
						"		UNION ALL " & _ 
						"		SELECT  c.productid, a.docketid,a.netweight,b.docketdate,a.doid,c.doorderid, " & _ 
						"				d.salescontractno,d.deliveryorderno, " & _ 
						"				d.transportercontractno,d.transporter,transportationcost,losstolerance,Upper(e.contractorname) transportername " & _ 
						"		FROM lhptiketckg a  " & _ 
						"		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _ 
						"		left JOIN lhpdomst c ON a.doid=c.doid AND b.fcba=c.fcba " & _ 
						"		left JOIN deliveryorder d ON Trim(c.doorderid)=Trim(d.deliveryorderno)  AND Trim(c.fcba)=Trim(d.fcba) " & _ 
						"		Left JOIN lhpcontractormst e ON d.transporter=e.eplascode " & _ 
						"		WHERE b.fcba='"&  Session("fcBA") &"' AND To_Char(b.docketdate,'yyyymm')='" & Year(Session("v_WeighBridge")) & "'||LPAD('" & Month(Session("v_WeighBridge")) & "', 2, 0) " & _
						"		)X " & _ 
						"		GROUP BY productid,salescontractno,deliveryorderno,doorderid,transportercontractno,transporter ,transportername ,transportationcost" & _ 
						" )j "

          		 connstr.execute(qCost_Activity)
							
		'		 If Err.Number = 0 Then
		'			 connstr.CommitTrans
		'			 Response.Flush
		'		 Else
		'			 connstr.RollbackTrans
		'			 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
		'			 Response.Flush
		'		 End If
				
			
                qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT AUTO_REFERENCE_04 referenceno,journalprefix,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_WeighBridge")) & " AND accountingyear=" & Year(Session("v_WeighBridge")) & " AND journalprefix='CTR' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.AUTO_REFERENCE_04 =  y.referenceno and x.journalprefix=y.journalprefix) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
				
				 connstr.execute(qjournal)
		
		         Response.Write  "<font face=Verdana size=1 color=Blue>" & "Prosses Transportir Product, success " & "<br/></font>" 
				 Response.Flush





			Response.write("<font face=Verdana size=1 color=Green>Transfer to Activity Based Costing..., Success!.<br/><br/></font>")
				 Response.Flush
		 

		 
		 
		 
	
	End IF
		

		' q=" Update PeriodControl Set LASTPERIODPOSTED=" & _
		  ' " to_date('" & txtLastDay & "','mm/dd/yyyy') ,CURRENTDATE=" & _
		  ' " to_date('" & txtNextMonth & "','mm/dd/yyyy')  " & _
		  ' " where SYSTEM='WeighBridge' and fcba='" & Session("fcBA") & "'"
		' connStr.execute(q)
		
		
		'**** Remark By ukie 1 Februari 2016 ****'
		' q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)),CURRENTDATE=add_months(CURRENTDATE,1)  , lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
		  ' " where SYSTEM='WeighBridge' and fcba='" & Session("fcBA") & "'"
		' connStr.execute(q)
		
		'**** Add By ukie 1 Februari 2016 ****'
			
			Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear,temp_EndingAccountingPeriod
							
						QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
											" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
											" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
											" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear,  " & _ 
											" periodcontrol_api.GetEndingAccountingPeriodTY(system,'" & Session("fcBA") & "')  EndingAccountingPeriod " & _ 
											" from periodcontrol p  " & _ 
											" where UPPER(system) =UPPER('WeighBridge')  and fcba='" & Session("fcBA") & "'"
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
								Response.Write  "<font face=Verdana size=1 color=red>" & "IPLAS Error, No PERIODCONTROL for WeighBridge " & "<br>"
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
							   " where UPPER(system) = UPPER('WeighBridge')  and fcba='" & Session("fcBA") & "'"
						ELSE
							IF temp_NextCloseAccountingPeriod < 12 THEN
								q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
								  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
								  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
								  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
								  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
								  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
								  " where UPPER(system) = UPPER('WeighBridge')  and fcba='" & Session("fcBA") & "'"
							ELSE
								 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
								   " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _     
								  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
								  " where  UPPER(system) = UPPER('WeighBridge')  and fcba='" & Session("fcBA") & "'"
							END IF
						END IF
						
						IF temp_NextCloseAccountingPeriod = temp_EndingAccountingPeriod AND temp_NextCurrentAccountingPeriod = 1  THEN
						
							q=" Update PeriodControl Set CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
							  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
							  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
							  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
							  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
							  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
							  " where UPPER(system) = UPPER('WeighBridge')  and and fcba='" & Session("fcBA") & "'"
						END IF	
			
		connStr.execute(q)
		
		


		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "<br>Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."

	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Weighbridge'" 
								
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
						" 'Weighbridge', " & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Weighbridge'"
				
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.CursorType = 3
	rsLenghtime.CursorLocation = 3
	rsLenghtime.Open strLenghtime, connstr

	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Weighbridge'"		
    				
	connStr.execute(upClosingLog)
	'--End Added--
	rsLenghtime.close
	set rsLenghtime = Nothing

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
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>WeighBridge Closing is succesfully...
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		
		<%
			txtLastDay  = cmdFindPeriod("WeighBridge")
			txtNextMonth = cmdPeriodClose("WeighBridge")		
			
			Session("v_WeighBridge") = txtLastDay
			Session("cls_WeighBridge") = txtNextMonth			
		%>
	<%End if%>
<%
else 
%>
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/error.gif">WeighBridge Closing is NOT completed...
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