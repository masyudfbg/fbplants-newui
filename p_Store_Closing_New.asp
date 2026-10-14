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
<%
			'Declare All Variabel
			

			
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


Dim fine_result_01, fine_result_02, fine_result_03
Dim Startingtime, Endingtime
Dim TimeAwal, TimeAkhir, Lama
Dim nAlokasi
Dim BulanBfr, TahunBfr
Dim v_tahun : v_tahun = Year(Session("v_Stores"))
Dim v_Bulan : v_Bulan = Month(Session("v_Stores"))
			
TimeAwal=Time()
Startingtime = FormatDateTime(CDate(Now))

				if v_Bulan   = 1 then
					TahunBfr = v_tahun - 1
					BulanBfr = 12
				else
					TahunBfr = v_tahun
					BulanBfr = v_Bulan - 1
				end if
				
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

Response.Write "<br>Calculate Stock..." 
Response.Write "<br>Periode : " & v_Bulan & "/" & v_tahun 

Response.Flush()

' *** NORMALISASI ***
   MSQL =   "DELETE FROM STOCKLEDGER " & _
			"WHERE MONTH=" & v_Bulan &  _
			" 	   AND YEAR=" & v_tahun & _ 
			"      AND fcBA='" & Session("fcBA") & "'" 
			
    'response.write MSQL	& "<br>"
	connStr.execute(MSQL)
	connStr.execute("commit")
	Response.Write "<br><font color=Blue>" & "Normalisasi Stock Ledger is completed..." & "</font>"  
	Response.Flush()
   '0.*** OPENINGBALANCE ***'
   MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " &_ 
			"			 SELECT STORECODE,ITEMCODE,FCBA,(OPENINGBALANCEQTY+QTYRECEIVED+QTYRECEIVEDTASK+QTYWRITEOFF-QTYISSUED-QTYTRANSPORTTASK) QUANTITY, " & _ 
			"            (OPENINGBALANCEAMT+AMTRECEIVED+AMTRECEIVEDTASK+ AMTWRITEOFF - AMTISSUED-AMTTRANSPORTTASK) AMOUNT,AMTPERUNIT,      "& _
			"            '" & v_tahun & "' AS YEAR, '" & v_Bulan & "' AS MONTH " &  _
			"            FROM STOCKLEDGER WHERE MONTH=" & BulanBfr &  _
			" 			 and YEAR=" & TahunBfr & " and fcBA='" & Session("fcBA") & "'" &  _
			") y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.FCBA=y.FCBA and x.Month=y.Month and x.YEAR=y.YEAR ) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.OPENINGBALANCEQTY = y.QUANTITY, " &_
			"			   x.OPENINGBALANCEAMT = y.AMOUNT," &_
			"			   x.AMTPERUNIT = DECODE(y.QUANTITY,0,y.AMTPERUNIT,ABS(y.AMOUNT/y.QUANTITY)) " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, AMTPERUNIT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, y.YEAR,y.MONTH,  y.QUANTITY, y.AMOUNT ,DECODE(y.QUANTITY,0,y.AMTPERUNIT,ABS(y.AMOUNT/y.QUANTITY)), 'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	connStr.execute(MSQL)
	connStr.execute("commit")
	
	Response.Write "<br><font color=Blue>" & "Set Opening Balance is completed..." & "</font>"  
	Response.Flush()

'01.insert stock dari opname

	MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING (SELECT  STORECODE, ITEMCODE, YEAR, MONTH, QTY, AMT, AMTPERUNIT, FCBA FROM STOCKOPNAME WHERE fcBA='" & Session("fcBA") & "' and Month=" & v_Bulan & " and Year=" & v_tahun & " )y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.Year=y.YEAR and x.Month=y.MONTH ) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.OPENINGBALANCEQTY = NVL(y.QTY,0), " &_
			"			   x.OPENINGBALANCEAMT = NVL(y.AMT,0), " &_
			"              x.AMTPERUNIT        = NVL(y.AMTPERUNIT,0) " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, AMTPERUNIT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.YEAR), to_number(y.MONTH),  y.QTY, y.AMT ,y.AMTPERUNIT, 'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	connStr.execute(MSQL)
	connStr.execute("commit")
	
	Response.Write "<br><font color=Blue>" & "Set Stock Opname is completed..." & "</font>"
    Response.Flush()

'02. WRITEOFF/ Adjment ***'
	MSQL =  "MERGE INTO STOCKLEDGER x " & _
			"USING ( " & _ 
			" 	SELECT WF.STORECODE AS STORECODE,WD.STOCKCODE AS ITEMCODE,TO_CHAR(WF.WODATE,'YYYY') AS NYEAR," & _
			" 	TO_CHAR(WF.WODATE,'MM') AS NMONTH,nvl(SUM(WD.QUANTITY),0) AS QUANTITY,nvl(SUM(WD.AMOUNT),0) AS AMOUNT, WF.fcBA " & _
			" 	FROM WRITEOFF WF INNER JOIN WRITEOFFDETAILS WD ON (WF.WRITEOFFCODE=WD.WRITEOFFCODE AND WF.FCBA=WD.FCBA)   " & _
			" 	WHERE " & _
			" 	to_number(to_Char(WF.WODATE,'mm'))=" & v_Bulan & "" & _
			" 	and to_number(to_Char(WF.WODATE,'yyyy'))=" & v_tahun & "" & _ 
			" 	and WF.fcBA='" & Session("fcBA") & "'" & _
			" 	GROUP BY WF.STORECODE,WD.STOCKCODE,TO_CHAR(WF.WODATE,'YYYY'),TO_CHAR(WF.WODATE,'MM'), WF.fcBA" & _
			" ) y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.Year=to_number(y.NYEAR) and x.Month=to_number(y.NMONTH)) " & _
			"WHEN MATCHED THEN " & _
			"	UPDATE SET x.QTYWRITEOFF = y.QUANTITY, " & _
			"			   x.AMTWRITEOFF = y.AMOUNT, " & _
			"			   x.AMTPERUNIT = DECODE((NVL(x.OPENINGBALANCEQTY,0)+NVL(x.QTYRECEIVED,0)+NVL(x.QTYRECEIVEDTASK,0)+NVL(y.QUANTITY,0)-NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)),0,0,ABS((NVL(OPENINGBALANCEAMT,0)+NVL(AMTRECEIVED,0)+NVL(AMTRECEIVEDTASK,0)+ NVL(y.AMOUNT,0) - NVL(AMTISSUED,0)-NVL(AMTTRANSPORTTASK,0))/(NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(y.QUANTITY,0)-NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)))) " & _
			" WHEN NOT MATCHED THEN " & _
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYWRITEOFF, AMTWRITEOFF, AMTPERUNIT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT ,DECODE(y.QUANTITY,0,0,ABS(y.AMOUNT/y.QUANTITY)),'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	'Response.write(MSQL) & "<br>"
	connStr.execute(MSQL)
	connStr.execute("commit")
	
	Response.Write "<br><font color=Blue>" & "Closing Write Off is completed..." & "</font>"
    Response.Flush()
	
'1.*** RECEIVENOTE ***'
    MSQL =  "MERGE INTO RECEIVENOTEDETAIL x " &_
			"USING ( " &_ 
			"			 SELECT RN.FCBA,RN.RECEIVENOTECODE ,RD.LOCATIONCODE,RD.PURCHASEITEMCODE, " & _
			"            SUM(round(NVL((pod.quantity /rd.quantity) * (pod.otherchargestock-pod.otcstock3) ,0),2)) Othercharge " & _
			"			 FROM RECEIVENOTE RN  " & _ 
			"            INNER JOIN RECEIVENOTEDETAIL RD ON ( RN.RECEIVENOTECODE=RD.RECEIVENOTECODE AND RN.FCBA=RD.FCBA ) " & _
			" 			 LEFT JOIN purchaseorder po ON (po.pocode = RN.pocode AND po.fcba = RN.fcba)" & _
		    "            LEFT JOIN purchaseorderdetails pod ON (pod.pocode = po.pocode AND pod.fcba = po.fcba AND pod.linenokey = RD.purchaselinenokey)" & _
			"			 WHERE " & _
			"			 to_number(to_Char(RN.RNDATE,'mm'))= " & v_Bulan & _
			"			 and to_number(to_Char(RN.RNDATE,'yyyy'))= " & v_tahun & _
			"			 and rd.LOCATIONTYPE='" & "ST" & "'" & _
			"			 and RN.fcBA='" & Session("fcBA") & "'" & _
			"            AND pod.otherchargestock <> 0 " & _
			"            GROUP BY RN.FCBA,RN.RECEIVENOTECODE ,RD.LOCATIONCODE,RD.PURCHASEITEMCODE " & _
			"		) y " & _
			"ON (x.RECEIVENOTECODE=y.RECEIVENOTECODE and x.LOCATIONCODE=y.LOCATIONCODE and x.PURCHASEITEMCODE  = y.PURCHASEITEMCODE and x.FCBA=y.FCBA ) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.othercharge = y.Othercharge " 
			
	
	connStr.execute(MSQL)
	connStr.execute("commit")


    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " &_ 
			"        SELECT STORECODE,ITEMCODE,NYEAR,NMONTH,FCBA,SUM(QUANTITY) QUANTITY ,SUM(TotAmountReceived) AMOUNT FROM " & _ 
			"          ( " & _ 
			"			 SELECT RD.LOCATIONCODE AS STORECODE,RD.PURCHASEITEMCODE AS ITEMCODE,TO_CHAR(RN.RNDATE,'YYYY') AS NYEAR, " & _
			"			 TO_CHAR(RN.RNDATE,'MM') AS NMONTH, " & _
			"			 RN.FCBA, " & _
			"            nvl(RD.QUANTITY,0) AS QUANTITY, " & _
			"            nvl(RD.AMOUNTRECEIVED,0) + nvl(RD.othercharge,0) As TotAmountReceived " & _
			"			 FROM RECEIVENOTE RN  " & _ 
			"            INNER JOIN RECEIVENOTEDETAIL RD ON ( RN.RECEIVENOTECODE=RD.RECEIVENOTECODE AND RN.FCBA=RD.FCBA ) " & _
			"			 WHERE " & _
			"			 to_number(to_Char(RN.RNDATE,'mm'))= " & v_Bulan & _
			"			 and to_number(to_Char(RN.RNDATE,'yyyy'))= " & v_tahun & _
			"			 and rd.LOCATIONTYPE='" & "ST" & "'" & _
			"			 and RN.fcBA='" & Session("fcBA") & "'" & _
			"			) " & _
			"			GROUP BY STORECODE,ITEMCODE,NYEAR,NMONTH,FCBA " & _ 
			"		) y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.FCBA=y.FCBA and x.Year=to_number(y.NYEAR) and x.Month=to_number(y.NMONTH)) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.QTYRECEIVED = y.QUANTITY, " &_
			"			   x.AMTRECEIVED = y.AMOUNT, " &_
			"			   x.AMTPERUNIT = DECODE(NVL(x.OPENINGBALANCEQTY,0)+NVL(y.QUANTITY,0),0,NVL(x.AMTPERUNIT,0),((NVL(y.AMOUNT,0)+NVL(x.OPENINGBALANCEAMT,0))/(NVL(x.OPENINGBALANCEQTY,0)+NVL(y.QUANTITY,0)))) " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYRECEIVED, AMTRECEIVED, AMTPERUNIT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT , DECODE(y.QUANTITY,0,0,ABS(y.AMOUNT/y.QUANTITY)), 'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	ConnStr.execute(MSQL)
	connStr.execute("commit")
	
	
	Response.Write "<br><font color=Blue>" & "Closing Recieve Note is completed..." & "</font>"
	Response.Flush()
	
'1.*** ISSUE ***'	
    MSQL =  "MERGE INTO ISSUEDETAILS x " &_
			"USING ( " & _ 
			"	SELECT STORECODE, ITEMCODE, YEAR, MONTH,FCBA, " & _ 
			"	nvl(AMTPERUNIT,0) as AMTPERUNIT " & _ 
			"	FROM STOCKLEDGER WHERE Storecode='ST1' AND MONTH=" & v_Bulan &  _
			" 	and YEAR=" & v_tahun & " and fcBA='" & Session("fcBA") & "'" &  _
			" ) y " & _
			"ON (x.STOCKCODE  = y.ITEMCODE and x.fcba=y.fcba and " & _
			"    x.SIVCODE in (SELECT SIVCODE FROM ISSUE WHERE fcba='" & Session("fcBA") & "'" &  _
			"        AND to_number(to_Char(SIVDATE,'mm')) = " & v_Bulan & _
			"        and to_number(to_Char(SIVDATE,'yyyy')) = " & v_tahun & _
			"    )) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMOUNT = x.QUANTITY * y.AMTPERUNIT,x.Price=y.AMTPERUNIT " 
			
	connStr.execute(MSQL)
	connStr.execute("commit")
	
    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " & _ 
			"	SELECT ISU.STORECODE AS STORECODE,ID.STOCKCODE AS ITEMCODE,TO_CHAR(ISU.SIVDATE,'YYYY') AS NYEAR," & _
			" 	TO_CHAR(ISU.SIVDATE,'MM') AS NMONTH,nvl(SUM(ID.QUANTITY),0) AS QUANTITY,nvl(SUM(ID.AMOUNT),0) AS AMOUNT,ISU.fcBA" & _
			" 	FROM ISSUE ISU INNER JOIN ISSUEDETAILS ID ON (ISU.SIVCODE=ID.SIVCODE AND ISU.FCBA=ID.FCBA)  " & _
			" 	WHERE " & _
			" 	to_number(to_Char(ISU.SIVDATE,'mm'))=" & v_Bulan &  _
			" 	and to_number(to_Char(ISU.SIVDATE,'yyyy'))=" & v_tahun & _ 
			"   and ISU.fcBA='" & Session("fcBA") & "'" & _
			" 	GROUP BY ISU.STORECODE,ID.STOCKCODE,TO_CHAR(ISU.SIVDATE,'YYYY'),TO_CHAR(ISU.SIVDATE,'MM'),ISU.fcBA " &  _
			" ) y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.Year=to_number(y.NYEAR) and x.Month=to_number(y.NMONTH)) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.QTYISSUED = y.QUANTITY, " &_
			"			   x.AMTISSUED = y.AMOUNT " &_
			"	WHERE x.QTYISSUED  <> y.QUANTITY OR " &_
			"		   x.AMTISSUED <> y.AMOUNT " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYISSUED, AMTISSUED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT ,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	connStr.execute(MSQL)
	connStr.execute("commit")
	
	
	Response.Write "<br><font color=Blue>" & "Closing Issue is completed..." & "</font>"
    Response.Flush()
	
'3.*** TRANSPORTTASK_TAB ***'	 
    MSQL =  "MERGE INTO TRANSPORTTASKDTL_TAB x " &_
			"USING ( " &_ 
			"			 SELECT DISTINCT RN.FCBA,RN.ORDERNO ,RD.STOCKCODE,RD.StoreCode,STK.AMTPERUNIT,RD.LineNoHist   " & _
			"			 FROM TRANSPORTTASK_TAB RN  " & _ 
			"            INNER JOIN TRANSPORTTASKDTL_TAB RD ON ( RN.ORDERNO=RD.ORDERNO AND RN.FCBA=RD.FCBA ) " & _
			" 			 INNER JOIN STOCKLEDGER STK ON (STK.fcba = RN.fcba AND STK.MONTH=to_number(to_Char(RN.ORDERDATE,'mm')) AND STK.YEAR=to_number(to_Char(RN.ORDERDATE,'yyyy')) AND RD.StoreCode=Stk.Storecode AND RD.STOCKCODE=STK.ITEMCODE )" & _
		   "			 WHERE " & _
			"			 to_number(to_Char(RN.ORDERDATE,'mm'))= " & v_Bulan & _
			"			 and to_number(to_Char(RN.ORDERDATE,'yyyy'))= " & v_tahun & _
			"			 and RN.fcBA='" & Session("fcBA") & "'" & _
			"		) y " & _
			"ON (x.ORDERNO=y.ORDERNO and x.STOCKCODE  = y.STOCKCODE and x.FCBA=y.FCBA AND x.StoreCode=y.StoreCode and x.linenohist=y.linenohist) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMOUNT = x.QUANTITYRECEIVED * y.AMTPERUNIT, " & _ 
			"	           x.PRICE = y.AMTPERUNIT " 
	'response.write(MSQL)		
	connStr.execute(MSQL)
	connStr.execute("commit")
	
    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " & _ 
			"  SELECT TTD.STORECODE," & _
			"          TTD.STOCKCODE AS ITEMCODE," & _
			"          TO_CHAR (TT.ORDERDATE,  'YYYY') AS NYEAR," & _
			"          TO_CHAR (TT.ORDERDATE, 'MM') AS NMONTH," & _
			"          NVL (SUM (TTD.QUANTITYTRANSFERED), 0) AS QUANTITY," & _
			"          NVL (ROUND(SUM (TTD.AMOUNT),2), 0) AS AMOUNT," & _
			"          TT.FCBA " & _ 
			"     FROM TRANSPORTTASKDTL_TAB TTD" & _
			"     INNER JOIN TRANSPORTTASK_TAB TT ON (TT.ORDERNO = TTD.ORDERNO AND TT.FCBA = TTD.FCBA)" & _
			"     WHERE    " & _
			" 		        to_number(to_Char(TT.ORDERDATE,'mm'))=" & v_Bulan &  _
			" 		   AND	to_number(to_Char(TT.ORDERDATE,'yyyy'))=" & v_tahun & _
			"          AND TTD.FCBA = '"& Session("fcBA") &"'" & _
			"          AND TT.ROWSTATEID = 'Approved'" & _
			" GROUP BY TTD.STORECODE," & _
			"          TTD.STOCKCODE," & _
			"          TO_CHAR (TT.ORDERDATE,  'YYYY')," & _
			"          TO_CHAR (TT.ORDERDATE, 'MM') , TT.FCBA " &  _	
			" ) y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.Year=to_number(y.NYEAR) and x.Month=to_number(y.NMONTH)) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.QTYTRANSPORTTASK = y.QUANTITY, " &_
			"			   x.AMTTRANSPORTTASK = y.AMOUNT " &_
			"	WHERE x.QTYTRANSPORTTASK  <> y.QUANTITY OR " &_
			"		   x.AMTTRANSPORTTASK <> y.AMOUNT " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYTRANSPORTTASK, AMTTRANSPORTTASK, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT ,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	connStr.execute(MSQL)
	connStr.execute("commit")
	
	
	Response.Write "<br><font color=Blue>" & "Closing Transport Task is completed..." & "</font>"
	Response.Flush()

'2.*** RECEIVEDTASK ***'	
	MSQL =  "MERGE INTO RECEIVEDTASKDTL_TAB x " &_
			"USING ( " &_ 
			"			 SELECT DISTINCT RN.FCBA,RN.ORDERNO ,RD.STOCKCODE,RD.LOCATIONCODE,STK.AMTPERUNIT " & _
			"			 FROM RECEIVEDTASK_TAB RN  " & _ 
			"            INNER JOIN RECEIVEDTASKDTL_TAB RD ON ( RN.ORDERNO=RD.ORDERNO AND RN.FCBA=RD.FCBA ) " & _
			" 			 INNER JOIN STOCKLEDGER STK ON (STK.fcba = RN.fcba AND STK.MONTH=to_number(to_Char(RN.ORDERDATE,'mm')) AND STK.YEAR=to_number(to_Char(RN.ORDERDATE,'yyyy')) AND RD.LOCATIONCODE=Stk.Storecode AND RD.STOCKCODE=STK.ITEMCODE)" & _
	  	    "			 WHERE " & _
			"			 to_number(to_Char(RN.ORDERDATE,'mm'))= " & v_Bulan & _
			"			 and to_number(to_Char(RN.ORDERDATE,'yyyy'))= " & v_tahun & _
			"			 and rd.LOCATIONTYPE='" & "ST" & "'" & _
			"			 and RN.fcBA='" & Session("fcBA") & "'" & _
			"		) y " & _
			"ON (x.ORDERNO=y.ORDERNO and x.STOCKCODE  = y.STOCKCODE and x.FCBA=y.FCBA AND x.LOCATIONCODE=y.LOCATIONCODE) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMOUNTRECEIVED = x.QUANTITYRECEIVED * y.AMTPERUNIT, " & _ 
			"	           x.PRICERECEIVED = y.AMTPERUNIT " 
	
	connStr.execute(MSQL)
	connStr.execute("commit")

    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " & _ 
			" SELECT RD.LOCATIONCODE AS STORECODE,RD.STOCKCODE AS ITEMCODE,TO_CHAR(RN.ORDERDATE,'YYYY') AS NYEAR," & _
			" TO_CHAR(RN.ORDERDATE,'MM') AS NMONTH,nvl(SUM(RD.QUANTITYRECEIVED),0) AS QUANTITY,nvl(SUM(RD.AMOUNTRECEIVED),0) AS AMOUNT, RN.FCBA" & _
			" FROM receivedtask_tab RN INNER JOIN receivedtaskdtl_tab RD ON (RN.orderno=RD.orderno AND RN.FCBA=RD.FCBA) " & _
			" WHERE " & _
			" to_number(to_Char(RN.ORDERDATE,'mm'))=" & v_Bulan &  _
			" and to_number(to_Char(RN.ORDERDATE,'yyyy'))=" & v_tahun & _
			" and rd.LOCATIONTYPE='ST'" & _
			" and RN.fcBA='" & Session("fcBA") & "'" & _
			" GROUP BY RD.LOCATIONCODE,RD.STOCKCODE,TO_CHAR(RN.ORDERDATE,'YYYY'),TO_CHAR(RN.ORDERDATE,'MM'),RN.FCBA" &  _	
			" ) y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.Year=to_number(y.NYEAR) and x.Month=to_number(y.NMONTH)) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.QTYRECEIVEDTASK = y.QUANTITY, " &_
			"			   x.AMTRECEIVEDTASK = y.AMOUNT " &_
			"	WHERE x.QTYRECEIVEDTASK  <> y.QUANTITY OR " &_
			"		   x.AMTRECEIVEDTASK <> y.AMOUNT " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYRECEIVEDTASK, AMTRECEIVEDTASK, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT ,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	connStr.execute(MSQL)
	connStr.execute("commit")
	
	
	Response.Write "<br><font color=Blue>" & "Closing Recieve Task is completed..." & "</font>"
	Response.Flush()
	

	' 'Update Amt Issue di Tabel Stock Ledger
	' MSQL =  "MERGE INTO STOCKLEDGER x " &_
			' "USING ( " & _ 
			' "	SELECT STORECODE, ITEMCODE, YEAR, MONTH,FCBA, " & _ 
			' "	(OPENINGBALANCEAMT+AMTRECEIVED+AMTWRITEOFF+AMTRECEIVEDTASK-AMTTRANSPORTTASK) AS nRp, " & _ 
			' "	(OPENINGBALANCEQTY+QTYRECEIVED+QTYWRITEOFF+QTYRECEIVEDTASK-QTYTRANSPORTTASK) AS nQty, QTYISSUED ," & _ 
			' "	DECODE((OPENINGBALANCEQTY+QTYRECEIVED+QTYWRITEOFF+QTYRECEIVEDTASK-QTYTRANSPORTTASK),0,(OPENINGBALANCEAMT+AMTRECEIVED+AMTWRITEOFF+AMTRECEIVEDTASK-AMTTRANSPORTTASK),round((OPENINGBALANCEAMT+AMTRECEIVED+AMTWRITEOFF+AMTRECEIVEDTASK-AMTTRANSPORTTASK) / (OPENINGBALANCEQTY+QTYRECEIVED+QTYWRITEOFF+QTYRECEIVEDTASK-QTYTRANSPORTTASK),2)) AS nPrice " & _ 
			' "	FROM STOCKLEDGER WHERE MONTH=" & v_Bulan &  _
			' " 	and YEAR=" & v_tahun & " and fcBA='" & Session("fcBA") & "'" &  _
			' " ) y " & _
			' "ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.Year=y.year and x.Month=y.month) " & _
			' "WHEN MATCHED THEN " &_
			' "	UPDATE SET x.AMTISSUED = y.QTYISSUED * y.nPrice, " &_
			' "			   x.AMTPERUNIT = y.nPrice " 
			
	' connStr.execute(MSQL)
	' connStr.execute("commit")
	
	'Update Amt Issue di Tabel Stock Issue
	
	
	' Response.Write "<br><font color=Blue>" & "Update Amount at Issue Details is completed..." & "</font>"
	' Response.Flush()	
	' 'Update Amt transport task, Issue di Tabel Stock Ledger
	' MSQL =  "MERGE INTO TRANSPORTTASKDTL_TAB x " &_
			' "USING ( " & _ 
			' "	SELECT STORECODE, ITEMCODE, YEAR, MONTH,FCBA, AMTPERUNIT " & _ 
			' "	FROM STOCKLEDGER WHERE MONTH=" & v_Bulan &  _
			' " 	and YEAR=" & v_tahun & " and fcBA='" & Session("fcBA") & "'"  & _ 
			' " ) y " & _
			' "ON (x.STORECODE=y.STORECODE and x.STOCKCODE  = y.ITEMCODE and x.fcba=y.fcba " & _ 
			' " AND to_number(TO_CHAR (TRANSPORTTASK_API.GET_RCVDATE (x.FCBA, x.ORDERNO), 'MM'))= y.MONTH " &  _
			' " AND to_number(TO_CHAR (TRANSPORTTASK_API.GET_RCVDATE (x.FCBA, x.ORDERNO), 'YYYY'))= y.year ) " & _
			' "WHEN MATCHED THEN " &_
			' "	UPDATE SET x.AMOUNT = x.QUANTITYRECEIVED * y.AMTPERUNIT, " &_
			' "			   x.PRICE  = y.AMTPERUNIT " 
			
	' connStr.execute(MSQL)
	' connStr.execute("commit")
	
	
	' Response.Write "<br><font color=Blue>" & "Update Amount at Transport Task Details is completed..." & "</font>"
	' Response.Flush()	
	' 'Update Amt receivedtaskdtl 
	' MSQL =  "MERGE INTO receivedtaskdtl_tab x " &_
			' "USING ( " & _ 
			' "	SELECT STORECODE, ITEMCODE, YEAR, MONTH,FCBA, AMTPERUNIT " & _ 
			' "	FROM STOCKLEDGER WHERE MONTH=" & v_Bulan &  _
			' " 	and YEAR=" & v_tahun & " and fcBA='" & Session("fcBA") & "'"  & _ 
			' " ) y " & _
			' "ON (x.LOCATIONCODE=y.STORECODE and x.STOCKCODE= y.ITEMCODE and x.fcba=y.fcba " & _ 
			' " AND to_number(TO_CHAR (RECEIVEDTASK_API.GET_RCVDATE (x.FCBA, x.ORDERNO), 'MM'))= y.MONTH " &  _
			' " AND to_number(TO_CHAR (RECEIVEDTASK_API.GET_RCVDATE (x.FCBA, x.ORDERNO), 'YYYY'))= y.year ) " & _
			' "WHEN MATCHED THEN " &_
			' "	UPDATE SET x.AMOUNTRECEIVED = x.QUANTITYRECEIVED * y.AMTPERUNIT, " &_
			' "			   x.PRICERECEIVED  = y.AMTPERUNIT " 
			
	' connStr.execute(MSQL)
	' connStr.execute("commit")
	
	
	' Response.Write "<br><font color=Blue>" & "Update Amount at Receive Task Details is completed..." & "</font>"
	' Response.Flush()
	
	Response.Write "<br>Tranfer to Activity Costing..." 
	Response.Flush()
	
	qNormalisasiDetail = " DELETE FROM cost_activity_hist_tab where fcba =  '" & Session("fcBA") & "'"
	connstr.execute(qNormalisasiDetail)
	connStr.execute("Commit")

	qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Stores")) &  _
				 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Stores")) & " and fcBA='" & Session("fcBA") & "'" & _
				 " and SOURCETYPE = 'ST'"
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

		
			'----- Tranfer to Activity Costing
			'-
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
			
'-- <CA - WRITE OFF>			
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
									" LASTTIME,journalprefix ) " & _
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
								 " to_char(sysdate, 'HH24:MI'),'SWO')" 

					'response.write qCost_ActivityDB
					response.Flush()
					connstr.execute(qCost_ActivityDB)	
					
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
									" LASTTIME ,journalprefix) " & _
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
								 " to_char(sysdate, 'HH24:MI'),'SWO')" 

                    'response.write  qCost_ActivityCr
					response.Flush()
					connstr.execute(qCost_ActivityCr)	
					
					rsReceive.MoveNext
				loop
			end if
			rsReceive.Close
			set rsReceive = Nothing	
			
		qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT referenceno,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT referenceno,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Stores")) & " AND accountingyear=" & Year(Session("v_Stores")) & " AND journalprefix='SWO' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.referenceno =  y.referenceno) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq "
			connstr.execute(qjournal)
			
			'----- Tranfer to Activity Costing Completed 
			Response.Write "<br><font color=Blue>" & "Proses Cost Activity from WRITEOFF is completed..." & "</font>"	
			Response.Flush()
		    
			
'-- </CA - WRITE OFF>						
		
'-- </CA - TRANSPORTTASK>						
		
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
									"		  ROUND((SUM(NVL(ID.AMOUNT,0)) * -1) ,2) AMOUNT, " & _
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
																  " LASTTIME,journalprefix ) " & _
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
												" to_char(sysdate,'HH24:MI'),'STT') " 
 						' response.write strInsertCredit
						response.Flush()		  
						connstr.Execute(strInsertCredit)		
						
						
						
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
												"		  ROUND( (SUM (NVL(ID.AMOUNT,0)) * 1) ,2) AMOUNT, " & _
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
																			  " LASTTIME,journalprefix ) " & _
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
															" to_char(sysdate,'HH24:MI'),'STT') "
									response.Flush()		  
									'response.write qInsertDebit
									connstr.Execute(qInsertDebit)			
									
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
																			  " LASTTIME,journalprefix ) " & _
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
															" to_char(sysdate,'HH24:MI'),'STT') "
									response.Flush()		  
									connstr.Execute(qInsertDebit)	
									
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
		
		qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT referenceno,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT referenceno,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Stores")) & " AND accountingyear=" & Year(Session("v_Stores")) & " AND journalprefix='STT' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.referenceno =  y.referenceno) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq "
			connstr.execute(qjournal)
			
		Response.Write "<br><font color=Blue>" & "Proses Cost Activity from Transport Task is completed..." & "</font>"	
		Response.Flush()
'--</CA-TRANSPORTTASK>		
	
'--</CA-RECEIVEDTASK>		
		
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
		
		qReceive=   " select param_1 from parameterdetail" & _
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
							 " LASTUPDATE, LASTTIME,journalprefix ) " & _
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
							 " sysdate, to_char(sysdate,'HH24:MI'),'SRT')"
							 
				
				response.Flush()
				connstr.execute(qCost_ActivityDB)
			    connStr.execute("Commit")
				 
			
				
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
								 " LASTUPDATE, LASTTIME ,journalprefix) " & _
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
								 " sysdate, to_char(sysdate,'HH24:MI'),'SRT') "
								 
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
								 " LASTUPDATE, LASTTIME ,journalprefix) " & _
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
								 " sysdate, to_char(sysdate,'HH24:MI'),'SRT') "
								 
					connstr.execute(qCost_ActivityCR)
					connStr.execute("Commit")
				End If 
				
			  response.Flush()
			  rsReceive.MoveNext
		  loop
		End if
		
		qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT referenceno,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT referenceno,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Stores")) & " AND accountingyear=" & Year(Session("v_Stores")) & " AND journalprefix='SRT' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.referenceno =  y.referenceno) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq "
			connstr.execute(qjournal)
			
		
		Response.Write "<br><font color=Blue>" & "Proses Cost Activity from Received Task is completed..." & "</font>"	
        response.Flush()
'--</CA-RECEIVEDTASK>

'--</CA-ISSUE>
		
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
																  " LASTTIME ,journalprefix) " & _
			" SELECT  " &_
			"                                 ROWNUM+" &  maxLineNo & " LineNo," & _  
            "                                 tdate	, " &_		
			"								  jobcode, " &_
			"								  source, " &_
			"								  sourcetype, " &_
			"								  sourcecode, " &_
			"								  sourcecode_status, " &_
			"								  targettype, " &_
			"								  targetcode, " &_
			"								  targetcode_status, " &_
			"								  referenceno, " &_
			"								  quantity, " &_
			"								  output, " &_
			"                                 amount, " &_
			"								  fcba, " &_
			"                                 to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') closing_date, " & _
			"								  process_date, " &_
			"                                 '' TITLE_REFERENCE_01,'' AUTO_REFERENCE_01 , " & _
			"                                 '' TITLE_REFERENCE_02,'' AUTO_REFERENCE_02 , " & _
			"                                 '' TITLE_REFERENCE_03,'' AUTO_REFERENCE_03 , " & _
			"                                 '' TITLE_REFERENCE_04,STOCKCODE AUTO_REFERENCE_04 , " & _
			"                                 '' TITLE_REFERENCE_05,'' AUTO_REFERENCE_05, " & _
			"                                 DEBETCREDIT, " & _
			"" & Month(Session("v_Stores")) & ", " & _
			"" & Year(Session("v_Stores")) & ", " & _
			"                                 sysdate, " & _
			"                                 to_char(sysdate,'HH24:MI'),'SIS' " &_
			" FROM " &_
			" ( " &_
			"                    SELECT isu.SIVDATE tdate , " &_
			"										  id.jobcode, " &_
			"										  'ST' source, " &_
			"										  'ST' sourcetype, " &_
			"										  ISU.STORECODE sourcecode, " &_
			"										  '' sourcecode_status, " &_
			"										  id.locationtype targettype, " &_
			"										  id.locationcode targetcode, " &_
			"										  ''   targetcode_status, " &_
			"										  id.SIVCODE referenceno, " &_
			"										  sum(id.QUANTITY) quantity, " &_
			"										  0 output, " &_
			"                                         SUM(id.amount) amount, " &_
			"										  id.FCBA fcba, " &_
			"										  SYSDATE process_date, " &_
			"										  ID.STOCKCODE||'-'||Stock.FCNAME STOCKCODE, " &_     
            "                                         'Dr' DEBETCREDIT " &_			
			"									 FROM ISSUE ISU, " &_
			"										  ISSUEDETAILS ID, " &_
			"										  Stock, " &_
			"										  STOCKGROUP " &_
			"								    WHERE (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA) " &_
			"										  AND (id.STOCKCODE = stock.FCCODE AND id.fcba = stock.fcba) " &_
			"										  AND (stock.STOCKGROUPCODE = stockgroup.FCCODE " &_
			"										  AND stock.fcba = stockgroup.fcba) " &_
			"										  AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
			"										  AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy')) = " & Year(Session("v_Stores"))  & _
			"										  AND ISU.fcBA = '" & Session("fcBA") & "'" & _
			"                                         AND isu.sivcode NOT LIKE 'BKM%' " &_
			"                                   GROUP BY " &_
            "                             				isu.SIVDATE , " &_
			"							  				id.jobcode, " &_
			"							  				ISU.STORECODE , " &_
			"							  				id.locationtype , " &_
			"							  				id.locationcode , " &_
			"							  				id.SIVCODE , " &_
			"							  				id.FCBA , " &_
			"							  				ID.STOCKCODE||'-'||Stock.FCNAME " &_
			"                  UNION " &_
			"                  SELECT isu.SIVDATE tdate , " &_
			"										  stockgroup.CONTROLJOB jobcode, " &_
			"										  'ST' source, " &_
			"										  'ST' sourcetype, " &_
			"										  ISU.STORECODE sourcecode, " &_
			"										  '' sourcecode_status, " &_
			"										  'CB' targettype, " &_
			"										  'BS' targetcode, " &_
			"										  ''   targetcode_status, " &_
			"										  id.SIVCODE referenceno, " &_
			"										  Sum(id.QUANTITY) quantity, " &_
			"										  0 output, " &_
			"                                         Sum(id.amount)*-1, " &_
			"										  id.FCBA fcba, " &_
			"										  SYSDATE process_date, " &_
			"										  stockgroup.FCCODE||'-'||stockgroup.fcname STOCKCODE, " &_
			"                                         'Cr' DEBETCREDIT " &_
			"									 FROM ISSUE ISU, " &_
			"										  ISSUEDETAILS ID, " &_
			"										  Stock, " &_
			"										  STOCKGROUP " &_
			"								    WHERE (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA) " &_
			"										  AND (id.STOCKCODE = stock.FCCODE AND id.fcba = stock.fcba) " &_
			"										  AND (stock.STOCKGROUPCODE = stockgroup.FCCODE " &_
			"										  AND stock.fcba = stockgroup.fcba) " &_
			"										  AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
			"										  AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy')) = " & Year(Session("v_Stores"))  & _
			"										  AND ISU.fcBA = '" & Session("fcBA") & "'" & _
			"                                         AND isu.sivcode NOT LIKE 'BKM%' " &_
			"								 GROUP BY isu.SIVDATE, " &_
			"										  stockgroup.CONTROLJOB, " &_
			"										  ISU.STORECODE, " &_
			"										  id.SIVCODE, " &_
			"										  id.fcBA ," &_
			"                                         stockgroup.FCCODE||'-'||stockgroup.fcname " &_
			" ORDER BY referenceno " & _
			" ) " 
			'response.write strInsertCredit
			response.Flush()				
			connstr.Execute(strInsertCredit)	

		
		
		
		
		
		
		
		' strCursorSivCode = " select sivcode " & _
							' " from issue " & _
							' " where " & _
							' " to_number(to_char(sivdate, 'mm')) = " & month (session ("v_stores")) & _
							' " and to_number(to_char(sivdate, 'yyyy')) = " & year (session ("v_stores")) & _
							' " and fcba = '" & Session("fcBA") & "'" 
							
		' Set rsCursorSivCode = Server.CreateObject("ADODB.Recordset")
		' rsCursorSivCode.CursorType = 3
		' rsCursorSivCode.CursorLocation = 3
		' rsCursorSivCode.Open strCursorSivCode, connstr		

		' if not rsCursorSivCode.eof then
			' Do While Not rsCursorSivCode.eof
			
				' ' NAF : Jangan dirubah format string-nya menjadi lurus atau via excel, biarkan seperti apa adanya atau 
				' '       Jangan menggunakan TRIM untuk menghilangkan SPASI-nya "
			

				
				' strCursorCredit = " SELECT isu.SIVDATE tdate , " & _
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
										 ' " SUM (id.QUANTITY) quantity, " & _
										 ' " 0 output, " & _
										 ' " round( (SUM (nvl(id.AMOUNT,0)) * -1) ,2) amount, " & _
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
								' " GROUP BY isu.SIVDATE, " & _
										 ' " stockgroup.CONTROLJOB, " & _
										 ' " 'ST', " & _
										 ' " 'ST', " & _
										 ' " ISU.STORECODE, " & _
										 ' " '', " & _
										 ' " 'CB', " & _
										 ' " 'BS', " & _
										 ' " '', " & _
										 ' " id.SIVCODE, " & _
										 ' " id.linenokey, " & _
										 ' " id.locationtype, " & _
										 ' " id.locationcode, " & _
										 ' " id.jobcode, " & _
										 ' " id.fcBA, " & _
										 ' " id.STOCKCODE "
										 
				
				' Set rsCursorCredit = Server.CreateObject("ADODB.Recordset")
				' rsCursorCredit.CursorType = 3
				' rsCursorCredit.CursorLocation = 3
				' rsCursorCredit.Open strCursorCredit, connstr	
				
				
				' if not rsCursorCredit.eof then
				
					' Do While Not rsCursorCredit.eof
				
						' tempTotalAmountCredit = round(cdbl(rsCursorCredit("AMOUNT")),2)
						' tempSIVCODE = rsCursorCredit("referenceno")
						' tempStockCode = rsCursorCredit("STOCKCODE")

						' txtRefStringValues_Receive = tempSIVCODE & txtRefStringSeparator_Receive & _
													 ' rsCursorCredit("linenokey") & txtRefStringSeparator_Receive & _
													 ' rsCursorCredit("LOCATIONTYPE") & txtRefStringSeparator_Receive & _
													 ' rsCursorCredit("LOCATIONCODE") & txtRefStringSeparator_Receive & _
													 ' rsCursorCredit("jobcodeoriginal") & txtRefStringSeparator_Receive & _
													 ' rsCursorCredit("FCBA") 
													 
						
						' ' NAF : Jangan dirubah format string-nya menjadi lurus atau via excel, biarkan seperti apa adanya atau 
						' '       Jangan menggunakan TRIM untuk menghilangkan SPASI-nya "		

						
						' strCursorCount = " select count(*) totalRecord from ( " & _
												' " SELECT isu.SIVDATE, " & _
												' "	 id.JOBCODE, " & _
												' "	 ISU.STORECODE, " & _
												' "	 id.LOCATIONTYPE, " & _
												' "	 id.LOCATIONCODE, " & _
												' "	 id.SIVCODE, " & _
												' "    id.LINENOKEY, " & _
												' "	 id.FCBA " & _
												' " FROM ISSUE ISU, ISSUEDETAILS ID  " & _
											   ' " WHERE	 (ISU.SIVCODE = ID.SIVCODE AND ISU.FCBA = ID.FCBA)  " & _
												' "	 AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'mm')) = " & Month(Session("v_Stores")) &  _
												' "	 AND TO_NUMBER (TO_CHAR (ISU.SIVDATE, 'yyyy'))  =" & Year(Session("v_Stores"))  & _
												' "	 AND ISU.fcBA = '" & Session("fcBA") & "'" & _
												' "	 AND ISU.sivcode = '" & tempSIVCODE & "'  " & _
												' "    AND id.linenokey = '" & rsCursorCredit ("linenokey") & "' " & _
												' "    AND id.locationtype = '" & rsCursorCredit ("locationtype") & "' " & _
												' "    AND id.locationcode = '" & rsCursorCredit ("locationcode") & "' " & _
												' "    AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") &  "' " & _
												' "    AND ID.STOCKCODE = '" & tempStockCode & "' " & _
											' " GROUP BY isu.SIVDATE,  " & _
												' "	 id.JOBCODE,  " & _
												' "	 ISU.STORECODE,  " & _
												' "	 id.LOCATIONTYPE,  " & _
												' "	 id.LOCATIONCODE,  " & _
												' "	 id.SIVCODE,  " & _
												' "    id.LINENOKEY, " & _
												' "	 id.fcBA  " & _
											' " ) "  

					
						' Set rsCursorCount = Server.CreateObject("ADODB.Recordset")
						' rsCursorCount.CursorType = 3
						' rsCursorCount.CursorLocation = 3
						

						' rsCursorCount.Open strCursorCount, connstr		

						' totalRecord = cdbl(rsCursorCount("totalRecord"))
						' totalRowInserted = 0
						' tempTotalAmountDebit = 0

						
						' Response.Flush()
						
						
						' maxLineNo = maxLineNo + 1
						' strInsertCredit = " INSERT INTO cost_activity ( LINENO , " & _
																  ' " TDATE, " & _
																  ' " JOBCODE, " & _
																  ' " SOURCE, " & _
																  ' " SOURCETYPE, " & _
																  ' " SOURCECODE, " & _
																  ' " SOURCECODE_STATUS, " & _
																  ' " TARGETTYPE, " & _
																  ' " TARGETCODE, " & _
																  ' " TARGETCODE_STATUS, " & _
																  ' " REFERENCENO, " & _
																  ' " QUANTITY, " & _
																  ' " OUTPUT, " & _
																  ' " AMOUNT, " & _
																  ' " FCBA, " & _
																  ' " CLOSING_DATE, " & _ 
																  ' " PROCESS_DATE , " & _
																  ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																  ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																  ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																  ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																  ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _	
																  ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _	
																  ' " LASTUPDATE,  " & _
																  ' " LASTTIME ,journalprefix) " & _
												' " VALUES (" & _
												' maxLineNo & "," & _
												' " to_date('" & rsCursorCredit("tdate") & "','mm/dd/yyyy'),'" & _
												' rsCursorCredit("JOBCODE") & "','" & _
												' rsCursorCredit("SOURCE") & "','" & _
												' rsCursorCredit("SOURCETYPE") & "','" & _
												' rsCursorCredit("SOURCECODE") & "','" & _
												' rsCursorCredit("SOURCECODE_STATUS") & "','" & _
												' rsCursorCredit("TARGETTYPE") & "','" & _
												' rsCursorCredit("TARGETCODE") & "','" & _
												' "NAF||" & rsCursorCredit("TARGETCODE_STATUS") & "','" & _
												' rsCursorCredit("REFERENCENO") & "'," & _
												' rsCursorCredit("QUANTITY") & "," & _
												' rsCursorCredit("OUTPUT") & "," & _
												' tempTotalAmountCredit & ",'" & _
												' rsCursorCredit("FCBA") & "'," & _
												' " to_date('" & rsCursorCredit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
												' " sysdate, " & _
												' " Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & rsCursorCredit("TARGETCODE") & "'," & _
												' " Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & tempSIVCODE & "'," & _
												' " Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorCredit("TARGETTYPE") & "','Cr'), '" & tempStockCode & "'," & _
												' " '',''," & _
												' " '','','Cr'," & _
												' Month(Session("v_Stores")) & ", " & _
												' Year(Session("v_Stores")) & ", " & _
												' " sysdate, " & _
												' " to_char(sysdate,'HH24:MI'),'SIS') " 
 
						' response.Flush()		  
						' connstr.Execute(strInsertCredit)		
						
						' strCursorDebit = " SELECT isu.SIVDATE tdate, " & _
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
													 ' " SUM (id.QUANTITY) quantity,  " & _
													 ' " 0 output,  " & _
													 ' " round( (SUM (id.AMOUNT) * 1) ,2) amount,  " & _
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
													 ' " AND id.jobcode = '" & rsCursorCredit ("jobcodeoriginal") & "' " & _													 
											' " GROUP BY isu.SIVDATE,  " & _
													 ' " id.JOBCODE,  " & _
													 ' " 'ST',  " & _
													 ' " 'ST',  " & _
													 ' " ISU.STORECODE,  " & _
													 ' " '',  " & _
													 ' " id.LOCATIONTYPE,  " & _
													 ' " id.LOCATIONCODE,  " & _
													 ' " '',  " & _
													 ' " id.SIVCODE,  " & _
													 ' " id.linenokey, " & _
													 ' " id.locationtype, " & _
													 ' " id.locationcode, " & _
													 ' " id.jobcode, " & _												 
													 ' " id.fcBA, " & _
													 ' " id.STOCKCODE "
						' Set rsCursorDebit = Server.CreateObject("ADODB.Recordset")
						' rsCursorDebit.CursorType = 3
						' rsCursorDebit.CursorLocation = 3
						' rsCursorDebit.Open strCursorDebit, connstr		
						
						' tempRowNum = 1

						' if not rsCursorDebit.eof then
							' Do While Not rsCursorDebit.eof
								
								' tempDebit = round(cdbl(rsCursorDebit("amount")),2)
								
								' if totalRowInserted = totalRecord - 1 then 
									' tempTotalAmountDebit = tempTotalAmountDebit + tempDebit
									
									' if tempTotalAmountDebit < (-1 * tempTotalAmountCredit) then
										' tempSelisih = (-1 * tempTotalAmountCredit) - tempTotalAmountDebit				   
										' tempDebit = tempDebit + tempSelisih
									' else
										' tempSelisih = tempTotalAmountDebit - (-1 * tempTotalAmountCredit) 				   
										' tempDebit = tempDebit - tempSelisih								
									' end if
									' Response.Flush()
									' maxLineNo = maxLineNo + 1
									' qInsertDebit = " INSERT INTO cost_activity ( LINENO ," & _
																			  ' " TDATE, " & _
																			  ' " JOBCODE, " & _
																			  ' " SOURCE, " & _
																			  ' " SOURCETYPE, " & _
																			  ' " SOURCECODE, " & _
																			  ' " SOURCECODE_STATUS, " & _
																			  ' " TARGETTYPE, " & _
																			  ' " TARGETCODE, " & _
																			  ' " TARGETCODE_STATUS, " & _
																			  ' " REFERENCENO, " & _
																			  ' " QUANTITY, " & _
																			  ' " OUTPUT, " & _
																			  ' " AMOUNT, " & _
																			  ' " FCBA, " & _
																			  ' " CLOSING_DATE, " & _ 
																			  ' " PROCESS_DATE, " & _		
																			  ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																			  ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																			  ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																			  ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																			  ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _		
																			  ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
																			  ' " LASTUPDATE,  " & _
																			  ' " LASTTIME,journalprefix ) " & _
															' " VALUES (" & _
															' maxLineNo & "," & _
															' " to_date('" & rsCursorDebit("tdate") & "','mm/dd/yyyy'),'" & _
															' rsCursorDebit("JOBCODE") & "','" & _
															' rsCursorDebit("SOURCE") & "','" & _
															' rsCursorDebit("SOURCETYPE") & "','" & _
															' rsCursorDebit("SOURCECODE") & "','" & _
															' rsCursorDebit("SOURCECODE_STATUS") & "','" & _
															' rsCursorDebit("TARGETTYPE") & "','" & _
															' rsCursorDebit("TARGETCODE") & "','" & _
															' rsCursorDebit("TARGETCODE_STATUS") & "','" & _
															' rsCursorDebit("REFERENCENO") & "'," & _
															' rsCursorDebit("QUANTITY") & "," & _
															' rsCursorDebit("OUTPUT") & "," & _
															' tempDebit & ",'" & _
															' rsCursorDebit("FCBA") & "'," & _
															' " to_date('" & rsCursorDebit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
															' " sysdate, " & _
															' " Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & rsCursorDebit("TARGETCODE") & "'," & _
															' " Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempSIVCODE & "'," & _
															' " Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempStockCode & "'," & _
															' " '',''," & _
															' " '','','Dr'," & _
															' Month(Session("v_Stores")) & ", " & _
															' Year(Session("v_Stores")) & ", " & _
															' " sysdate, " & _
															' " to_char(sysdate,'HH24:MI'),'SIS') "
									' response.Flush()		  
									' 'response.write qInsertDebit
									' connstr.Execute(qInsertDebit)			
									
								' else
									' Response.Flush()
									' maxLineNo = maxLineNo + 1
									' qInsertDebit = " INSERT INTO cost_activity ( LINENO ," & _
																			  ' " TDATE, " & _
																			  ' " JOBCODE, " & _
																			  ' " SOURCE, " & _
																			  ' " SOURCETYPE, " & _
																			  ' " SOURCECODE, " & _
																			  ' " SOURCECODE_STATUS, " & _
																			  ' " TARGETTYPE, " & _
																			  ' " TARGETCODE, " & _
																			  ' " TARGETCODE_STATUS, " & _
																			  ' " REFERENCENO, " & _
																			  ' " QUANTITY, " & _
																			  ' " OUTPUT, " & _
																			  ' " AMOUNT, " & _
																			  ' " FCBA, " & _
																			  ' " CLOSING_DATE, " & _ 
																			  ' " PROCESS_DATE, " & _		
																			  ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
																			  ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
																			  ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
																			  ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
																			  ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _																				  
																			  ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _	
																			  ' " LASTUPDATE,  " & _
																			  ' " LASTTIME,journalprefix ) " & _
															' " VALUES (" & _
															' maxLineNo & "," & _
															' " to_date('" & rsCursorDebit("tdate") & "','mm/dd/yyyy'),'" & _
															' rsCursorDebit("JOBCODE") & "','" & _
															' rsCursorDebit("SOURCE") & "','" & _
															' rsCursorDebit("SOURCETYPE") & "','" & _
															' rsCursorDebit("SOURCECODE") & "','" & _
															' rsCursorCredit("SOURCECODE_STATUS") & "','" & _
															' rsCursorDebit("TARGETTYPE") & "','" & _
															' rsCursorDebit("TARGETCODE") & "','" & _
															' rsCursorDebit("TARGETCODE_STATUS") & "','" & _
															' rsCursorDebit("REFERENCENO") & "'," & _
															' rsCursorDebit("QUANTITY") & "," & _
															' rsCursorDebit("OUTPUT") & "," & _
															' tempDebit & ",'" & _
															' rsCursorDebit("FCBA") & "'," & _
															' " to_date('" & rsCursorDebit("CLOSING_DATE") & "','mm/dd/yyyy'), " & _
															' " sysdate, " & _
															' " Cost_Activity_Info_Api.Get_Title_Reference_01 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & rsCursorDebit("TARGETCODE") & "'," & _
															' " Cost_Activity_Info_Api.Get_Title_Reference_02 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempSIVCODE & "'," & _
															' " Cost_Activity_Info_Api.Get_Title_Reference_03 ('ST', '" &  rsCursorDebit("TARGETTYPE") & "','Dr'), '" & tempStockCode & "'," & _
															' " '',''," & _
															' " '','','Dr'," & _
															' Month(Session("v_Stores")) & ", " & _
															' Year(Session("v_Stores")) & ", " & _
															' " sysdate, " & _
															' " to_char(sysdate,'HH24:MI'),'SIS') "
									' response.Flush()		  
									' connstr.Execute(qInsertDebit)	
								
									' tempTotalAmountDebit = tempTotalAmountDebit + tempDebit
									' totalRowInserted = totalRowInserted + 1
								' end if	
								' tempRowNum = tempRowNum + 1
								' rsCursorDebit.MOVENEXT
							' loop
							' rsCursorDebit.close
							' set rsCursorDebit=Nothing				
						' end if
						' rsCursorCredit.MOVENEXT
					' loop
					' rsCursorCredit.close
					' set rsCursorCredit=Nothing		
				' end if
				' rsCursorSivCode.MOVENEXT
			' loop
			' rsCursorSivCode.close
			' set rsCursorSivCode=Nothing
		' End if
		
		qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT referenceno,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT referenceno,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Stores")) & " AND accountingyear=" & Year(Session("v_Stores")) & " AND journalprefix='SIS' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.referenceno =  y.referenceno) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq "
		connstr.execute(qjournal)
		
		Response.Write "<br><font color=Blue>" & "Proses Cost Activity for Issue Details is completed..." & "</font>"	
		response.Flush()
		
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

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME, fcentry) VALUES (" & _
						" 'Store', " & _
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
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>Process Inventory is succesfully...
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		
				
		
		
	<%End if%>


</body>
</html>