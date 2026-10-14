<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Inventory Process</title>
</head>
<body>
<%
			'Declare All Variabel
			

			Dim Startingtime, Endingtime
			Dim TimeAwal, TimeAkhir, Lama
			Dim nAlokasi
			Dim BulanBfr, TahunBfr
			Dim v_tahun : v_tahun = Request("txtYear")
			Dim v_Bulan : v_Bulan = Request("txtMonth")
			
			TimeAwal=Time()
			Startingtime = FormatDateTime(CDate(Now))
			
			
				if v_Bulan   = 1 then
					TahunBfr = v_tahun - 1
					BulanBfr = 12
				else
					TahunBfr = v_tahun
					BulanBfr = v_Bulan - 1
				end if
		%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
<br/>
		Unit Process : <% response.write Session("fcBA")%> <br> 
		Inventory Process Period is <% response.write v_Bulan & "/" & v_tahun%>    <br><br>
		Please Wait...
		<br/>
		<br/>
		<%
		Response.Write "Start ..." & Time() & "<br><br>"
		Response.Flush
		
		%>
<%

'--Added by Ukie at 25-Oct-2012--	

Dim fine_result_01, fine_result_02, fine_result_03

Dim TDATE, STORAGE_CODE, QUANTITY_IN, QUANTITY_OUT, REFERENCES, REMARKS
DIM MSQL,Qry, rsReceive,rsWrite,rsIssue
DIM strLineMax, rsstrLineMax
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
Dim amtprice : amtprice = 0
'***** End Added By Ukie at 25 feb 2014 *****'
' *** NORMALISASI ***
   MSQL =   "DELETE FROM STOCKLEDGER " & _
			"WHERE MONTH=" & v_Bulan &  _
			" 	   AND YEAR=" & v_tahun & _ 
			"      AND fcBA='" & Session("fcBA") & "'" 
			
    'response.write MSQL	& "<br>"
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	Response.Write "<br><font color=Blue>" & "Normalisasi Stock Ledger is completed..." & "</font>"  
	Response.Flush()
	
'0.*** OPENINGBALANCE ***'
if cint(v_tahun)*100+cint(v_Bulan) <> 202401 then
   MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " &_ 
			"			 SELECT STORECODE,ITEMCODE,FCBA,(NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)-NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) QUANTITY, " & _ 
			"            Decode((NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)-NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)),0,0,"& _
            "                   (NVL(OPENINGBALANCEAMT,0)+NVL(AMTRECEIVED,0)+NVL(AMTRECEIVEDTASK,0)+NVL(AMTRECEIVETASKDIV,0)+NVL(AMTWRITEOFF,0)+NVL(AMTRETURNED,0)-NVL(AMTISSUED,0)-NVL(AMTTRANSPORTTASK,0)-NVL(AMTTRANSPORTTASKDIV,0))) AMOUNT,"& _
			"            AMTPERUNIT,'" & v_tahun & "' AS YEAR, '" & v_Bulan & "' AS MONTH " &  _
			"            FROM STOCKLEDGER WHERE STORECODE in ('ST1','MN1') AND MONTH=" & BulanBfr &  _
			" 			 and YEAR=" & TahunBfr & " and fcBA='" & Session("fcBA") & "'" &  _
			") y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.FCBA=y.FCBA and x.Month=y.Month and x.YEAR=y.YEAR ) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.OPENINGBALANCEQTY = y.QUANTITY, " &_
			"			   x.OPENINGBALANCEAMT = round(y.AMOUNT,2)," &_
			"			   x.AMTPERUNIT = round(DECODE(y.QUANTITY,0,y.AMTPERUNIT,ABS(y.AMOUNT/y.QUANTITY)),2) " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, AMTPERUNIT,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, y.YEAR,y.MONTH,  y.QUANTITY, y.AMOUNT ,DECODE(y.QUANTITY,0,round(y.AMTPERUNIT,2),round(ABS(y.AMOUNT/y.QUANTITY),2)),'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
else
  MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " &_ 
			"	SELECT STORECODE,ITEMCODE,FCBA,Sum(QUANTITY)QUANTITY,round(Sum(AMOUNT),2) AMOUNT,round(MAX(AMTPERUNIT),2) AMTPERUNIT," &_ 
			"	YEAR,  MONTH " &_ 
			"	FROM " &_ 
			"	( " &_ 
			"			 SELECT STORECODE,b.NEWCODE ITEMCODE,a.FCBA," & _
			"            NVL(ENDINGBALANCEQTY,0) QUANTITY, " & _ 
			"            NVL(ENDINGBALANCEAMT,0) AMOUNT,"& _
			"            AMTPERUNIT,'" & v_tahun & "' AS YEAR, '1' AS MONTH " & _
			"            FROM STOCKLEDGER a  INNER JOIN stock_2024 b ON (a.itemcode=b.fccode) " & _ 
			"  			 WHERE  isdelete='0' AND STORECODE in ('ST1','MN1') AND MONTH=" & BulanBfr &  _
			" 			 AND YEAR=" & TahunBfr & " and a.fcBA='" & Session("fcBA") & "'" &  _
			"			 AND Endingbalanceqty <> 0 " & _
			"            UNION ALL " & _
			"			 SELECT STORECODE,NEWCODE ITEMCODE,x.FCBA,QUANTITY,AMOUNT,AMTPERUNIT,YEAR,  MONTH " & _
			"			 FROM " & _
			"			  ( " & _
			"				SELECT STORECODE,b.KodePengganti ITEMCODE,a.FCBA," & _
			"				NVL(ENDINGBALANCEQTY,0) QUANTITY, " & _
			"				NVL(ENDINGBALANCEAMT,0) AMOUNT," & _
			"				AMTPERUNIT,'" & v_tahun & "' AS YEAR, '1' AS MONTH " & _ 
			"				FROM STOCKLEDGER a INNER JOIN stock_2024 b ON a.itemcode=b.fccode  " & _ 
			"	  			WHERE  isdelete='1' AND kodepengganti <> '0'  AND STORECODE in ('ST1','MN1') AND MONTH=" & BulanBfr &  _
			" 			 	AND YEAR=" & TahunBfr & " and a.fcBA='" & Session("fcBA") & "'" &  _
			"				AND Endingbalanceqty <> 0 " & _
			"			  )x INNER JOIN stock_2024 y ON X.itemcode=y.fccode " & _
			"  )  GROUP BY  STORECODE,ITEMCODE,STORECODE,ITEMCODE,FCBA,YEAR,MONTH " & _
			") y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.FCBA=y.FCBA and x.Month=y.Month and x.YEAR=y.YEAR ) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.OPENINGBALANCEQTY = y.QUANTITY, " &_
			"			   x.OPENINGBALANCEAMT = y.AMOUNT," &_
			"			   x.AMTPERUNIT = y.AMTPERUNIT " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, AMTPERUNIT,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, y.YEAR,y.MONTH,  y.QUANTITY, y.AMOUNT ,y.AMTPERUNIT,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            

end if   
   'response.write MSQL
   ExecuteAndLog "p_inventory_closing",MSQL 
   
	
		
	
Response.Write "<br><font color=Blue>" & "Set Opening Balance is completed..." & "</font>"  
Response.Flush()

'01.insert stock dari opname

	MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING (SELECT  STORECODE, ITEMCODE, YEAR, MONTH, NVL(QTY,0) QTY, round(NVL(AMT,0),2) AMT, round(NVL(AMTPERUNIT,0),2) AMTPERUNIT, FCBA FROM STOCKOPNAME WHERE fcBA='" & Session("fcBA") & "' and Month=" & v_Bulan & " and Year=" & v_tahun & " )y " & _
			"ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.Year=y.YEAR and x.Month=y.MONTH ) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.OPENINGBALANCEQTY = NVL(y.QTY,0), " &_
			"			   x.OPENINGBALANCEAMT = NVL(y.AMT,0), " &_
			"              x.AMTPERUNIT        = NVL(y.AMTPERUNIT,0) " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, AMTPERUNIT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.YEAR), to_number(y.MONTH),  y.QTY, y.AMT ,y.AMTPERUNIT, 'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	'response.write MSQL & "<br>"
	ExecuteAndLog "p_inventory_closing",MSQL 
	


	Response.Write "<br><font color=Blue>" & "Closing Opname is completed..." & "</font>"
    Response.Flush()

'02. WRITEOFF/ Adjment ***'
	MSQL =  "MERGE INTO STOCKLEDGER x " & _
			"USING ( " & _ 
			" 	SELECT WF.STORECODE AS STORECODE,WD.STOCKCODE AS ITEMCODE,TO_CHAR(WF.WODATE,'YYYY') AS NYEAR," & _
			" 	TO_CHAR(WF.WODATE,'MM') AS NMONTH,nvl(SUM(WD.QUANTITY),0) AS QUANTITY, round(nvl(SUM(WD.AMOUNT),0),2) AS AMOUNT, WF.fcBA " & _
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
			"			   x.AMTPERUNIT = round(DECODE((NVL(x.OPENINGBALANCEQTY,0)+NVL(x.QTYRECEIVED,0)+NVL(x.QTYRECEIVEDTASK,0)+NVL(y.QUANTITY,0)-NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)),0,0,ABS((NVL(OPENINGBALANCEAMT,0)+NVL(AMTRECEIVED,0)+NVL(AMTRECEIVEDTASK,0)+ NVL(y.AMOUNT,0) - NVL(AMTISSUED,0)-NVL(AMTTRANSPORTTASK,0))/(NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(y.QUANTITY,0)-NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)))),2) " & _
			" WHEN NOT MATCHED THEN " & _
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYWRITEOFF, AMTWRITEOFF, AMTPERUNIT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, round(y.AMOUNT,2) , round(DECODE(y.QUANTITY,0,0,ABS(y.AMOUNT/y.QUANTITY)),2),'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	'Response.write(MSQL) & "<br>"
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	

	
	Response.Write "<br><font color=Blue>" & "Closing Write Off is completed..." & "</font>"
    Response.Flush()
		
	
'1.*** RECEIVENOTE ***'
    MSQL =  "MERGE INTO RECEIVENOTEDETAIL x " &_
			"USING ( " &_ 
			"			 SELECT RN.FCBA,RN.RECEIVENOTECODE ,RD.LOCATIONCODE,RD.PURCHASEITEMCODE, " & _
			"            round(SUM(NVL((rd.quantity/pod.quantity) * pod.otherchargestock ,0)),2) Othercharge " & _
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
			
	
	ExecuteAndLog "p_inventory_closing",MSQL 
	


    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " &_ 
			"        SELECT STORECODE,ITEMCODE,NYEAR,NMONTH,FCBA,SUM(QUANTITY) QUANTITY ,SUM(TotAmountReceived) AMOUNT FROM " & _ 
			"          ( " & _ 
			"			 SELECT RD.LOCATIONCODE AS STORECODE,RD.PURCHASEITEMCODE AS ITEMCODE,TO_CHAR(RN.RNDATE,'YYYY') AS NYEAR, " & _
			"			 TO_CHAR(RN.RNDATE,'MM') AS NMONTH, " & _
			"			 RN.FCBA, " & _
			"            nvl(RD.QUANTITY,0) AS QUANTITY, " & _
			"            round(nvl(RD.AMOUNTRECEIVED,0) + nvl(RD.othercharge,0),2) As TotAmountReceived " & _
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
			"			   x.AMTPERUNIT = round(DECODE(NVL(x.OPENINGBALANCEQTY,0)+NVL(y.QUANTITY,0),0,NVL(x.AMTPERUNIT,0),((NVL(y.AMOUNT,0)+NVL(x.OPENINGBALANCEAMT,0))/(NVL(x.OPENINGBALANCEQTY,0)+NVL(y.QUANTITY,0)))),2) " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYRECEIVED, AMTRECEIVED, AMTPERUNIT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, round(y.AMOUNT,2) , round(DECODE(y.QUANTITY,0,0,ABS(y.AMOUNT/y.QUANTITY)),2), 'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	

	
	Response.Write "<br><font color=Blue>" & "Closing Recieve Note is completed..." & "</font>"
	Response.Flush()
	

	

'2.*** ISSUE ***'	
    MSQL =  "MERGE INTO ISSUEDETAILS x " &_
			"USING ( " & _ 
			"	SELECT STORECODE, ITEMCODE, YEAR, MONTH,FCBA, " & _ 
			"	round(nvl(AMTPERUNIT,0),2) as AMTPERUNIT " & _ 
			"	FROM STOCKLEDGER WHERE MONTH=" & v_Bulan &  _
			" 	and YEAR=" & v_tahun & " and fcBA='" & Session("fcBA") & "' and STORECODE='ST1' " &  _
			" ) y " & _
			"ON (x.STOCKCODE  = y.ITEMCODE and x.fcba=y.fcba and " & _
			"    x.SIVCODE in (SELECT H.SIVCODE FROM ISSUE H INNER JOIN ISSUEDETAILS D ON H.SIVCODE=D.SIVCODE AND H.FCBA=D.FCBA " & _ 
			" 	 WHERE NVL(D.REMARKS,'-') NOT LIKE('AUTO GENERATE BKM BAHAN%') AND H.fcba='" & Session("fcBA") & "'" &  _
			"        AND to_number(to_Char(H.SIVDATE,'mm')) = " & v_Bulan & _
			"        and to_number(to_Char(H.SIVDATE,'yyyy')) = " & v_tahun & _
			"    )) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMOUNT = round(x.QUANTITY * y.AMTPERUNIT,2) ,x.Price=y.AMTPERUNIT " 
			
	'response.write MSQL
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " & _ 
			"	SELECT ISU.STORECODE AS STORECODE,ID.STOCKCODE AS ITEMCODE,TO_CHAR(ISU.SIVDATE,'YYYY') AS NYEAR," & _
			" 	TO_CHAR(ISU.SIVDATE,'MM') AS NMONTH,nvl(SUM(ID.QUANTITY),0) AS QUANTITY,round(nvl(SUM(ID.AMOUNT),0),2) AS AMOUNT,ISU.fcBA" & _
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
			"	/*WHERE x.QTYISSUED  <> y.QUANTITY OR */" &_
			"	/*	   x.AMTISSUED <> y.AMOUNT */ " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYISSUED, AMTISSUED, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT ,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	

	
	Response.Write "<br><font color=Blue>" & "Closing Issue is completed..." & "</font>"
    Response.Flush()
	
'3.*** TRANSPORTTASK_TAB ***'	 

    MSQL =  "MERGE INTO TRANSPORTTASKDTL_TAB x " &_
			"USING ( " &_ 
			"			 SELECT DISTINCT RN.FCBA,RN.ORDERNO ,RD.STOCKCODE,RD.StoreCode,round(STK.AMTPERUNIT,2) AMTPERUNIT,RD.LineNoHist  " & _
			"			 FROM TRANSPORTTASK_TAB RN  " & _ 
			"            INNER JOIN TRANSPORTTASKDTL_TAB RD ON ( RN.ORDERNO=RD.ORDERNO AND RN.FCBA=RD.FCBA ) " & _
			" 			 INNER JOIN STOCKLEDGER STK ON (STK.fcba = RN.fcba AND STK.MONTH=to_number(to_Char(RN.ORDERDATE,'mm')) AND STK.YEAR=to_number(to_Char(RN.ORDERDATE,'yyyy')) AND RD.StoreCode=Stk.Storecode AND RD.STOCKCODE=STK.ITEMCODE )" & _
		   "			 WHERE " & _
			"			 RN.ROWSTATEID='Approved' AND Extract(month from RN.ORDERDATE)= " & v_Bulan & _
			"			 and Extract(year from RN.ORDERDATE)= " & v_tahun & _
			"			 and RN.fcBA='" & Session("fcBA") & "'" & _
			"		) y " & _
			"ON (trim(x.ORDERNO)=trim(y.ORDERNO) and trim(x.STOCKCODE)  = trim(y.STOCKCODE) and trim(x.FCBA) = trim(y.FCBA) AND trim(x.StoreCode)=trim(y.StoreCode) and x.linenohist=y.linenohist) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMOUNT = round(x.QUANTITYTRANSFERED * y.AMTPERUNIT,2)  , " & _ 
			"	           x.PRICE = y.AMTPERUNIT " 
	
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " & _ 
			"  SELECT TTD.STORECODE," & _
			"          TTD.STOCKCODE AS ITEMCODE," & _
			"          TO_CHAR (TT.ORDERDATE,  'YYYY') AS NYEAR," & _
			"          TO_CHAR (TT.ORDERDATE, 'MM') AS NMONTH," & _
			"          NVL (SUM (TTD.QUANTITYTRANSFERED), 0) AS QUANTITY," & _
			"          ROUND(NVL (SUM (TTD.AMOUNT), 0),2) AS AMOUNT," & _
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
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	


	
	Response.Write "<br><font color=Blue>" & "Closing Transport Task is completed..." & "</font>"
	Response.Flush()

'4.*** RECEIVEDTASK ***'	
	MSQL =  "MERGE INTO RECEIVEDTASKDTL_TAB x " &_
			"USING ( " &_ 
			"			 SELECT RN.FCBA,RN.ORDERNO ,RD.STOCKCODE,RD.LOCATIONCODE,round(TT.PRICE,2) AMTPERUNIT,RD.linenosource" & _
			"			 FROM RECEIVEDTASK_TAB RN  " &_ 
			"			 INNER JOIN RECEIVEDTASKDTL_TAB RD ON ( RN.ORDERNO=RD.ORDERNO AND RN.FCBA=RD.FCBA ) " &_ 
		    "            INNER JOIN TRANSPORTTASKDTL_TAB TT ON ( RD.ordersource=TT.orderno AND RD.FCBA=TT.fcbadistination AND RD.fcbasource=TT.fcba  AND RD.stockcode=TT.stockcode AND RD.linenosource=TT.LINENO) " &_ 
            "			 WHERE " & _
			"			 to_number(to_Char(RN.ORDERDATE,'mm'))= " & v_Bulan & _
			"			 and to_number(to_Char(RN.ORDERDATE,'yyyy'))= " & v_tahun & _
			"			 and rd.LOCATIONTYPE='" & "ST" & "'" & _
			"			 and RN.fcBA='" & Session("fcBA") & "'" & _
			"		) y " & _
			"ON (x.ORDERNO=y.ORDERNO and x.STOCKCODE  = y.STOCKCODE and x.FCBA=y.FCBA AND x.LOCATIONCODE=y.LOCATIONCODE AND x.linenosource=y.linenosource) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMOUNTRECEIVED = round(x.QUANTITYRECEIVED * y.AMTPERUNIT,2) , " & _ 
			"	           x.PRICERECEIVED = y.AMTPERUNIT " 
	
	ExecuteAndLog "p_inventory_closing",MSQL 
	

    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " & _ 
			" SELECT RD.LOCATIONCODE AS STORECODE,RD.STOCKCODE AS ITEMCODE,TO_CHAR(RN.ORDERDATE,'YYYY') AS NYEAR," & _
			" TO_CHAR(RN.ORDERDATE,'MM') AS NMONTH,nvl(SUM(RD.QUANTITYRECEIVED),0) AS QUANTITY, round(nvl(SUM(RD.AMOUNTRECEIVED),0),2) AS AMOUNT, RN.FCBA" & _
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
			"			   x.AMTRECEIVEDTASK = y.AMOUNT, " &_
			"			   x.AMTPERUNIT = round(DECODE(y.QUANTITY*y.Amount,0,x.AMTPERUNIT,y.AMOUNT/y.QUANTITY),2) " &_
			"WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYRECEIVEDTASK, AMTRECEIVEDTASK, AMTPERUNIT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT ,DECODE(y.QUANTITY,0,0,y.AMOUNT/y.QUANTITY), 'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	ExecuteAndLog "p_inventory_closing",MSQL 
	

	
	Response.Write "<br><font color=Blue>" & "Closing Recieve Task is completed..." & "</font>"
	Response.Flush()


	'4.*** TRANSPORTTASKDIV ***'
	MSQL =  "MERGE INTO TRANSPORTTASKDIVDTL_TAB x " &_
			"USING ( " &_ 
			"			SELECT RD.STOCKCODE, RN.FCBA,RN.ORDERNO ,round(NVL(STK.AMTPERUNIT,0),2) AMTPERUNIT,SUM(NVL(RD.QUANTITYTRANSFERED,0)) QUANTITYTRANSFERED,RD.LINENO " & _
			"			FROM TRANSPORTTASKDIV_TAB  RN   " & _
			"			INNER JOIN TRANSPORTTASKDIVDTL_TAB  RD ON ( RN.ORDERNO=RD.ORDERNO AND RN.FCBA=RD.FCBA ) " & _
			"			LEFT JOIN STOCKLEDGER STK ON (STK.fcba = RN.fcba AND STK.MONTH=to_number(to_Char(RN.ORDERDATE,'mm')) AND STK.YEAR=to_number(to_Char(RN.ORDERDATE,'yyyy')) " & _
			"			AND STK.STORECODE='ST1' AND RD.STOCKCODE=STK.ITEMCODE) " & _
			"			WHERE  " & _
			"			to_number(to_Char(RN.ORDERDATE,'mm'))=" & v_Bulan &  _
			"			and to_number(to_Char(RN.ORDERDATE,'yyyy'))=" & v_tahun & _
			"			and RN.fcBA='" & Session("fcBA") & "' " & _
			"			GROUP BY RD.STOCKCODE, RN.FCBA,RN.ORDERNO ,STK.AMTPERUNIT,RD.LINENO " & _
			"		) y " & _
			"ON (x.ORDERNO=y.ORDERNO and x.STOCKCODE  = y.STOCKCODE and x.FCBA=y.FCBA and x.lineno=y.lineno) " & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMOUNT = round(x.QUANTITYTRANSFERED * y.AMTPERUNIT,2) , " & _ 
			"	           x.PRICE = y.AMTPERUNIT " 
	'Response.Write MSQL
	ExecuteAndLog "p_inventory_closing",MSQL 
	

	MSQL = "MERGE INTO STOCKLEDGER x  " &_
			" USING (  " &_
			"  SELECT  TTD.STORECODEFROM , " &_
			"          TTD.STOCKCODE AS ITEMCODE, " &_
			"          TO_CHAR (TT.ORDERDATE,  'YYYY') AS NYEAR, " &_
			"          TO_CHAR (TT.ORDERDATE, 'MM') AS NMONTH, " &_
			"          NVL (SUM (TTD.QUANTITYTRANSFERED), 0) AS QUANTITY, " &_
			"          ROUND(NVL (SUM (TTD.AMOUNT), 0),2) AS AMOUNT, " &_
			"          TT.FCBA  " &_
			"     FROM TRANSPORTTASKDIVDTL_TAB TTD " &_
			"     INNER JOIN TRANSPORTTASKDIV_TAB TT ON (TT.ORDERNO = TTD.ORDERNO AND TT.FCBA = TTD.FCBA) " &_
			"     WHERE     " &_
			" 		        to_number(to_Char(TT.ORDERDATE,'mm'))= " & v_Bulan &  _ 
			" 		   AND	to_number(to_Char(TT.ORDERDATE,'yyyy'))= " & v_tahun & _
			"          AND TTD.FCBA = '" & Session("fcBA") & "'  " &_
			" GROUP BY TTD.STORECODEFROM ,  " &_
			"          TTD.STOCKCODE, " &_
			"          TO_CHAR (TT.ORDERDATE,  'YYYY'), " &_
			"          TO_CHAR (TT.ORDERDATE, 'MM') , TT.FCBA  " &_		
			" ) y  " &_
			" ON (x.STORECODE=y.STORECODEFROM and x.ITEMCODE = y.ITEMCODE and x.fcba=y.fcba and x.Year=to_number(y.NYEAR) and x.Month=to_number(y.NMONTH))  " &_
			" WHEN MATCHED THEN " &_
			"	UPDATE SET x.QTYTRANSPORTTASKDIV  = y.QUANTITY, " &_
			"			   x.AMTTRANSPORTTASKDIV  = y.AMOUNT " &_
			" WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYTRANSPORTTASKDIV , AMTTRANSPORTTASKDIV , FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODEFROM, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT ,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
	MSQL = "MERGE INTO STOCKLEDGER x  " &_
			" USING (  " &_
			"  SELECT  TTD.STORECODETO , " &_
			"          TTD.STOCKCODE AS ITEMCODE, " &_
			"          TO_CHAR (TT.ORDERDATE,  'YYYY') AS NYEAR, " &_
			"          TO_CHAR (TT.ORDERDATE, 'MM') AS NMONTH, " &_
			"          NVL (SUM (TTD.QUANTITYTRANSFERED), 0) AS QUANTITY, " &_
			"          ROUND(NVL (SUM (TTD.AMOUNT), 0),2) AS AMOUNT, " &_
			"          TT.FCBA  " &_
			"     FROM TRANSPORTTASKDIVDTL_TAB TTD " &_
			"     INNER JOIN TRANSPORTTASKDIV_TAB TT ON (TT.ORDERNO = TTD.ORDERNO AND TT.FCBA = TTD.FCBA) " &_
			"     WHERE     " &_
			" 		        to_number(to_Char(TT.ORDERDATE,'mm'))= " & v_Bulan &  _ 
			" 		   AND	to_number(to_Char(TT.ORDERDATE,'yyyy'))= " & v_tahun & _
			"          AND TTD.FCBA = '" & Session("fcBA") & "'  " &_
			" GROUP BY TTD.STORECODETO ,  " &_
			"          TTD.STOCKCODE, " &_
			"          TO_CHAR (TT.ORDERDATE,  'YYYY'), " &_
			"          TO_CHAR (TT.ORDERDATE, 'MM') , TT.FCBA  " &_		
			" ) y  " &_
			" ON (x.STORECODE=y.STORECODETO and x.ITEMCODE = y.ITEMCODE and x.fcba=y.fcba and x.Year=to_number(y.NYEAR) and x.Month=to_number(y.NMONTH))  " &_
			" WHEN MATCHED THEN " &_
			"	UPDATE SET x.qtyreceivetaskdiv  = y.QUANTITY, " &_
			"			   x.amtreceivetaskdiv  = y.AMOUNT " &_
			" WHEN NOT MATCHED THEN " &_
			"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, qtyreceivetaskdiv , amtreceivetaskdiv , FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			"	VALUES(y.STORECODETO, y.ITEMCODE, to_number(y.NYEAR), to_number(y.NMONTH),  y.QUANTITY, y.AMOUNT ,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
	

	
	Response.Write "<br><font color=Blue>" & "Closing Transport Task Divisi is completed..." & "</font>"
	Response.Flush()
	
	
	'4.*** STOCK RETURN ***'
		MSQL = "MERGE INTO STOCKLEDGER x  " &_
			" USING (  " &_
					"	SELECT " &_
					"	to_number(to_char(returndate,'yyyy')) accountingyear," &_
					"	to_number(to_char(returndate,'mm')) accountingperiod," &_
					"	FCBA," &_
					"	STOCKCODE," &_
					"	STORE," &_
					"	Sum(QUANTITY) Quantity," &_
					"	round(Sum(AMOUNT),2) Amount" &_
					"	FROM" &_
					"	(" &_
					"		SELECT ID.FCBA," &_
					"		TT.returndate," &_
					"		ID.returnno REFERENCENO," &_
					"		TT.NOREF refno," &_
					"		ID.ITEMCODE STOCKCODE," &_
					"		TT.STORE," &_
					"		SUM(ID.qtyreturn) QUANTITY," &_
					"		Sum(ID.qtyreturn * IDSIS.PRICE) AMOUNT" &_
					"		FROM returnissuedtl ID" &_
					"		INNER JOIN returnissuehdr TT ON (TT.returnno = ID.returnno AND TT.FCBA = ID.FCBA)" &_
					"		INNER JOIN issuedetails IDSIS ON (TT.NOREF=IDSIS.SIVCODE AND TT.FCBA=IDSIS.FCBA AND ID.ITEMCODE = IDSIS.STOCKCODE AND ID.LinenoKey=IDSIS.LinenoKey)" &_
					"		WHERE" &_
					"		TO_NUMBER(TO_CHAR(TT.returndate, 'mm')) =  " & v_Bulan & "" &_
					"		AND TO_NUMBER(TO_CHAR(TT.returndate, 'yyyy')) = " & v_tahun & "" &_
					"		AND ID.fcBA = '" & Session("fcBA") & "'" &_
					"		GROUP BY  ID.FCBA,TT.NOREF,TT.returndate,ID.returnno,ID.ITEMCODE,TT.STORE" &_					
					"	)" &_
					"	GROUP BY to_number(to_char(returndate,'yyyy')),to_number(to_char(returndate,'mm')),fcba,stockcode,store " &_
					" ) y  " &_
					" ON (x.STORECODE=y.STORE and x.ITEMCODE = y.STOCKCODE and x.fcba=y.fcba and x.Year=to_number(y.accountingyear) and x.Month=to_number(y.accountingperiod))  " &_
					" WHEN MATCHED THEN " &_
					"	UPDATE SET x.QTYRETURNED  = y.QUANTITY, " &_
					"			   x.AMTRETURNED  = y.AMOUNT " &_
					" WHEN NOT MATCHED THEN " &_
					"	Insert (STORECODE, ITEMCODE, YEAR, MONTH, QTYRETURNED , AMTRETURNED , FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
					"	VALUES(y.STORE, y.STOCKCODE, to_number(y.accountingyear), to_number(y.accountingperiod),  y.QUANTITY, y.AMOUNT ,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
	Response.Write "<br><font color=Blue>" & "Closing Stock Return is completed..." & "</font>"
	Response.Flush()

	
	
	'******************** UPDATE AMTPERUNIT All
    MSQL =  "MERGE INTO STOCKLEDGER x " &_
			"USING ( " & _ 
			"	SELECT ITEMCODE, YEAR, MONTH,FCBA, " & _ 
			"	ROUND(NVL(AMTPERUNIT,0),2) as AMTPERUNIT " & _ 
			"	FROM STOCKLEDGER WHERE MONTH=" & v_Bulan &  _
			" 	and YEAR=" & v_tahun & " and fcBA='" & Session("fcBA") & "' and STORECODE='ST1' and nvl(AMTPERUNIT,0) <> 0 " &  _
			" ) y " & _
			"ON (x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.MONTH=y.month and x.year=y.year )" & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMTPERUNIT = y.AMTPERUNIT , x.AMTISSUED = x.QTYISSUED * y.AMTPERUNIT " 
			
	'response.write MSQL
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
	MSQL =  " UPDATE STOCKLEDGER  " &_
			" SET endingbalanceqty =  (NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)-NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)) ," & _ 
			" endingbalanceamt = Decode((NVL(OPENINGBALANCEQTY,0)+NVL(QTYRECEIVED,0)+NVL(QTYRECEIVEDTASK,0)+NVL(QTYRECEIVETASKDIV,0)+NVL(QTYWRITEOFF,0)+NVL(QTYRETURNED,0)-NVL(QTYISSUED,0)-NVL(QTYTRANSPORTTASK,0)-NVL(QTYTRANSPORTTASKDIV,0)),0,0,"& _
            "                   (NVL(OPENINGBALANCEAMT,0)+NVL(AMTRECEIVED,0)+NVL(AMTRECEIVEDTASK,0)+NVL(AMTRECEIVETASKDIV,0)+NVL(AMTWRITEOFF,0)+NVL(AMTRETURNED,0)-NVL(AMTISSUED,0)-NVL(AMTTRANSPORTTASK,0)-NVL(AMTTRANSPORTTASKDIV,0))) "& _
			" WHERE MONTH=" & v_Bulan &  _
			" and YEAR=" & v_tahun & " AND fcBA='" & Session("fcBA") & "' AND nvl(AMTPERUNIT,0) <> 0 " 
			
	'response.write MSQL
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
	Response.Write "<br><font color=Blue>" & "Update Price to all Store is completed..." & "</font>"
	Response.Flush()

	
	MSQL =  "MERGE INTO ISSUEDETAILS x " &_
			"USING ( " & _ 
			"    SELECT round(a.amtperunit,2) amtperunit,b.sivcode,b.linenokey ,b.fcba,b.STORECODE,b.STOCKCODE,b.YEAR,b.MONTH FROM stockledger a " & _
            "    INNER JOIN " & _
            "    ( " & _
            "      SELECT x.sivcode,y.linenokey,x.FCBA,STORECODE,SIVDATE,STOCKCODE,To_Char(SIVDATE,'YYYY') YEAR,To_Char(SIVDATE,'MM') MONTH " & _
            "      FROM ISSUE x INNER JOIN ISSUEDETAILS y ON (x.SIVCODE=Y.SIVCODE AND x.fcba=y.fcba) " & _
            "      WHERE " & _
            "      x.fcba = '" & Session("fcBA") & "' " & _
            "      AND To_Char(SIVDATE,'YYYY') = " & v_tahun & "  " & _
            "      AND To_Char(SIVDATE,'MM') = " & v_Bulan & " " & _
            "    ) b ON (a.fcba=b.fcba AND a.itemcode=b.stockcode AND a.MONTH=b.MONTH AND a.YEAR=b.YEAR AND a.storecode=b.storecode ) " & _
			" ) y " & _
			"ON (x.STOCKCODE  = y.STOCKCODE and x.fcba=y.fcba and x.sivcode=y.sivcode and x.linenokey=y.linenokey )" & _
			"WHEN MATCHED THEN " &_
			"	UPDATE SET x.AMOUNT = round(NVL(x.QUANTITY,0) * NVL(y.AMTPERUNIT,0),2) ,x.Price=NVL(y.AMTPERUNIT,2) " 
			
	'response.write MSQL
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
	Response.Write "<br><font color=Blue>" & "Update Amount ISSUE is completed..." & "</font>"
	Response.Flush()
	
	MSQL =  " MERGE INTO BKMPROSES_BAHAN x " & _
			" USING (  " & _
			"   SELECT round(a.amount,2) amount ,b.sivcode,a.linenokey,b.FCBA,a.STOCKCODE " & _
            "   FROM ISSUEDETAILS a INNER JOIN ISSUE b ON (a.SIVCODE=b.SIVCODE AND a.fcba=b.fcba)  " & _
            "   WHERE " & _
            "   a.fcba = '" & Session("fcBA") & "' " & _
            "   AND To_Char(SIVDATE,'YYYY') = " & v_tahun & "  " & _
            "   AND To_Char(SIVDATE,'MM') = " & v_Bulan & " " & _
            "   AND NVL(a.REMARKS,'-') LIKE('AUTO GENERATE BKM BAHAN%') " & _
            "  ) y " & _
			" ON (x.bkmno=y.sivcode AND x.fcba=y.fcba AND x.ITEMCODE  = y.STOCKCODE and x.linenokey=y.linenokey ) " & _
			" WHEN MATCHED THEN " & _
			" UPDATE SET " & _
			" x.AMOUNT      = NVL(y.AMOUNT,0)," & _
			" x.FINALAMOUNT = NVL(y.AMOUNT,0) "

			
	'response.write MSQL
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	
	Response.Write "<br><font color=Blue>" & "Update BKM BAHAN is completed..." & "</font>"
	Response.Flush()
	
	MSQL =  "BEGIN bkmjournal_api.Update_bahan_monthly(" & v_tahun & "," & v_Bulan & ",'" & Session("fcBA") & "','System'); END;" 
    'response.write MSQL
	ExecuteAndLog "p_inventory_closing",MSQL 
	
	Response.Write "<br><font color=Blue>" & "Update JOURNAL BKM BAHAN is completed..." & "</font>"
	Response.Flush()
	
	'insert stock dari opname

	' MSQL =  "MERGE INTO STOCKpLEDGER x " &_
			' "USING (SELECT  STORECODE, ITEMCODE, YEAR, MONTH, QTY, AMT, FCBA FROM STOCKOPNAME WHERE fcBA='" & Session("fcBA") & "' and Month=" & v_Bulan & " and Year=" & v_tahun & " )y " & _
			' "ON (x.STORECODE=y.STORECODE and x.ITEMCODE  = y.ITEMCODE and x.fcba=y.fcba and x.Year=to_number(y.YEAR) and x.Month=to_number(y.MONTH) ) " & _
			' "WHEN MATCHED THEN " &_
			' "	UPDATE SET x.OPENINGBALANCEQTY = NVL(y.QTY,0), " &_
			' "			   x.OPENINGBALANCEAMT = NVL(y.AMT,0) " &_
			' "	WHERE NVL(x.OPENINGBALANCEQTY,0)  <> NVL(y.QTY,0) OR " &_
			' "		   NVL(x.OPENINGBALANCEAMT,0) <> NVL(y.AMT,0) " &_
			' "WHEN NOT MATCHED THEN " &_
			' "	Insert (STORECODE, ITEMCODE, YEAR, MONTH, OPENINGBALANCEQTY, OPENINGBALANCEAMT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " &_
			' "	VALUES(y.STORECODE, y.ITEMCODE, to_number(y.YEAR), to_number(y.MONTH),  y.QTY, y.AMT ,'SYSTEM','SYSTEM','1:::',y.fcba,SYSDATE,to_char(sysdate, 'HH24:MI'))"            
	
	' response.write MSQL
	' ExecuteAndLog "p_inventory_closing",MSQL 
	' 
	
	' Response.Write "<br><font color=Blue>" & "Closing Opname is completed..." & "</font>"
    ' Response.Flush()


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

		'qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Store'" 
								
		'set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		'rsqLineMax.Open qLineMax, connstr
		'if not rsqLineMax.eof then
		'	tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		'else
		'	tempMaxLine = 0
		'end if 

		'tempMaxLine = tempMaxLine + 1

		'strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME, fcentry) VALUES (" & _
		'				" 'Store', " & _
		'				" '" & Session("fcBA") & "'," & _
		'				" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
		'				" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
		'				" " & tempMaxLine & _
		'				" , SYSDATE, to_char(SYSDATE,'HH24:MI') " & _
		'				" )"
		'connStr.execute(strClosingLog)

		'strLenghtime = 	" SELECT" & _
		'				" FCBA," & _
		'				" LINENO," & _
		'				" SYSTEM," & _
		'				" TO_CHAR(STARTINGTIME, 'DD.MM.YYYY:HH24:MI:SS') STARTINGTIME," & _
		'				" TO_CHAR(ENDINGTIME, 'DD.MM.YYYY:HH24:MI:SS') ENDINGTIME," & _
		'				" TRUNC((((86400*(ENDINGTIME-STARTINGTIME))/60)/60)/24) DAYS," & _
		'				" TRUNC(((86400*(ENDINGTIME-STARTINGTIME))/60)/60)-24*(TRUNC((((86400*(ENDINGTIME-STARTINGTIME))/60)/60)/24)) HRS," & _
		'				" TRUNC((86400*(ENDINGTIME-STARTINGTIME))/60)-60*(TRUNC(((86400*(ENDINGTIME-STARTINGTIME))/60)/60)) MIN," & _
		'				" TRUNC(86400*(ENDINGTIME-STARTINGTIME))-60*(TRUNC((86400*(ENDINGTIME-STARTINGTIME))/60)) SEC" & _
		'				" FROM CLOSINGLOG" & _
		'				" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Store'"
				
		'Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
		'rsLenghtime.CursorType = 3
		'rsLenghtime.CursorLocation = 3
		'rsLenghtime.Open strLenghtime, connstr

		'upClosingLog = " UPDATE CLOSINGLOG SET" & _
		'				" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
		'				", LASTUPDATE = Sysdate, LastTime = To_Char(Sysdate, 'HH24:MI') " & _
		'				" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Store'"		
		'connStr.execute(upClosingLog)
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