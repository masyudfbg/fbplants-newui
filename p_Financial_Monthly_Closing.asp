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
dim v_AllowanceCode,qCost_Activity,qNormalisasiDetail



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

CheckWaterFlow = "TRUE"
if CheckWaterFlow = "TRUE" then 

'--1. Journal Alokasi Asset 
'--  ALOKASI BTL DEPRESIASI TM INTI (6510102) (D)
'--  ALOKASI DEPRESIASI INTI (6429998) (K)
 qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  fcba='" & Session("fcBA") & "' AND to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Costbook")) &  _
					 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Costbook")) & " and journalprefix='ARF'" & _
					 " AND to_number(accountingperiod)=" & Month(Session("v_Costbook")) & "" &  _
			         " AND to_number(accountingyear)=" & Year(Session("v_Costbook")) 
					 
connstr.execute(qNormalisasiDetail)	

Response.Write  "<font face=Verdana size=1 color=Green>" & "Prosses Alokasi DEPRESIASI " & "<br/></font>" 
Response.Flush
				
		qCostAMax =  "select nvl(max(lineno), 0) maxLineNo from cost_activity  where fcba = '" & Session("fcBA")  & "'" 
		 
		 Set rsCostAMax = Server.CreateObject("ADODB.Recordset")
		 rsCostAMax.Open qCostAMax, connstr
		 If Not rsCostAMax.EOF Then
		 maxLNO = cdbl(rsCostAMax("maxLineNo"))
		 Else
		 maxLNO = 0
		 End If
		
		 rsCostAMax.close
		 set rsCostAMax = Nothing
		 
		 maxLNO = maxLNO + 1
		 
		 connstr.BeginTrans
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
						 " SELECT to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'), '6510102' jobcode,'CB' Source,'CB' SOURCETYPE,'FAX' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						 " 'CB' TARGETTYPE ,'BS' TARGETCODE,'' TARGETCODE_STATUS,'FAX/'||fcba REFERENCENO, " & _ 
						 " 0 QUANTITY, 0, amount ,fcba,to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'),sysdate, " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','FAX/'||fcba, " & _ 
						 " '','','Dr', " & _ 
						 "" & Month(Session("v_Costbook")) & " , " & Year(Session("v_Costbook")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'FAX' " & _
						 " FROM " & _ 
						 " (	SELECT fcba,Sum(amount) amount " & _ 
						 " 	FROM " & _ 
						 " 	( " & _ 
						 "		SELECT a.fcba,jobgroupcode,Sum(a.amount) amount " & _ 
						 "		FROM cost_activity a " & _ 
						 "		INNER JOIN job b ON a.jobcode=b.fccode  AND a.fcba=b.fcba " & _ 
						 "		INNER JOIN jobgroup c ON  b.jobgroupcode=c.fccode AND b.fcba=c.fcba " & _ 
                         " 		WHERE a.fcba = '" & Session("fcBA") & "'" & _ 
                         " 		AND b.jobgroupcode IN ('6420101','6420102','6420103','6420104','6420105','6420106','6420107','6420108') " & _ 
						 " 		AND To_Char(a.tdate,'yyyymm')='" & Year(Session("v_Costbook")) & "'||LPAD('" & Month(Session("v_Costbook")) & "', 2, 0) " & _
						 " 		GROUP BY  a.fcba,jobgroupcode " & _ 
						 "  )x " & _
						 " GROUP BY fcba " & _
						 " )j "
						
		
		 connstr.execute(qCost_Activity)
							
		 If Err.Number = 0 Then
			 connstr.CommitTrans
			 Response.Flush
		 Else
			 connstr.RollbackTrans
			 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			 Response.Flush
		 End If

		maxLNO = maxLNO + 1
		connstr.BeginTrans
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
						 " SELECT to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'), '6429998' jobcode,'CB' Source,'CB' SOURCETYPE,'FAX' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						 " 'CB' TARGETTYPE ,'BS' TARGETCODE,'' TARGETCODE_STATUS,'FAX/'||fcba REFERENCENO, " & _ 
						 " 0 QUANTITY, 0, amount* -1  ,fcba,to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'),sysdate, " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','FAX/'||fcba, " & _ 
						 " '','','Cr', " & _ 
						 "" & Month(Session("v_Costbook")) & " , " & Year(Session("v_Costbook")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'FAX' " & _
						 " FROM " & _ 
						 " (	SELECT fcba,Sum(amount) amount " & _ 
						 " 	FROM " & _ 
						 " 	( " & _ 
						 "		SELECT a.fcba,jobgroupcode,Sum(a.amount) amount " & _ 
						 "		FROM cost_activity a " & _ 
						 "		INNER JOIN job b ON a.jobcode=b.fccode  AND a.fcba=b.fcba " & _ 
						 "		INNER JOIN jobgroup c ON  b.jobgroupcode=c.fccode AND b.fcba=c.fcba " & _ 
                         " 		WHERE a.fcba = '" & Session("fcBA") & "'" & _ 
                         " 		AND b.jobgroupcode IN ('6420101','6420102','6420103','6420104','6420105','6420106','6420107','6420108') " & _ 
						 " 		AND To_Char(a.tdate,'yyyymm')='" & Year(Session("v_Costbook")) & "'||LPAD('" & Month(Session("v_Costbook")) & "', 2, 0) " & _
						 " 		GROUP BY  a.fcba,jobgroupcode " & _ 
						 "  )x " & _
						 " GROUP BY fcba " & _
						 " )j "
						
		
		 connstr.execute(qCost_Activity)
							
		 If Err.Number = 0 Then
			 connstr.CommitTrans
			 Response.Flush
		 Else
			 connstr.RollbackTrans
			 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			 Response.Flush
		 End If

		qjournal=" MERGE INTO cost_activity x " & _
					 " USING " & _
					 " ( " & _
					 "  SELECT fcba,accountingyear,accountingperiod,referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					 "  ( " & _
					 "  SELECT DISTINCT fcba,accountingyear,accountingperiod,AUTO_REFERENCE_04 referenceno,journalprefix,company_api.get_companyno(businessunit_api.get_companycode(fcba))||'/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					 "  FROM cost_activity WHERE accountingperiod="& Month(Session("v_Costbook")) & " AND accountingyear=" & Year(Session("v_Costbook")) & " AND journalprefix='FAX' " & _
					 "  Order BY referenceno " & _
					 "  ) " & _
					 " )y ON (x.fcba=y.fcba and  x.AUTO_REFERENCE_04 =  y.referenceno AND  x.accountingyear=y.accountingyear AND x.accountingperiod=y.accountingperiod AND x.journalprefix=y.journalprefix) " & _
					 " WHEN MATCHED THEN " & _
					 " UPDATE SET x.journalno = y.seq " 
		
		          connstr.execute(qjournal)
				  
		Response.Write  "<font face=Verdana size=1 color=Blue>" & "Prosses HPP TBS Internal, success " & "<br/></font>" 
		Response.Flush

IF Session("Estate_Mill") = "E" THEN
'--1. Journal Penjualan TBS INTERNAL (ARF-Account Receiveable FruitBunch) 
'--   INVOICE TBS INTERNAL YANG BELUM DITERIMA (2110608) (D)
'--   PENJUALAN TBS (5110106) (K)
'
'--2. Journal HPP TBS INTERNAL (HPF-Harga Pokok  FruitBunch)  202304 temp Close 
'--   HARGA POKOK PENJUALAN TBS (5210201) (D)
'--   ALOKASI BTL  (6599999) (K)

'--2. Journal HPP TBS INTERNAL (HPF-Harga Pokok  FruitBunch)  202304 temp Close 
'--   HARGA POKOK PENJUALAN TBS (5210201) (D)
'--   ALOKASI BTL  (6599999) (K)


'--1. Journal Penjualan TBS INTERNAL (ARF-Account Receiveable FruitBunch) 
		Dim strCursorSivCode, rsCursorSivCode
        Dim qCostAMax,rsCostAMax
		
		qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  fcba='" & Session("fcBA") & "' AND to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Costbook")) &  _
					 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Costbook")) & " and journalprefix='ARF'" & _
					 " AND to_number(accountingperiod)=" & Month(Session("v_Costbook")) & "" &  _
			         " AND to_number(accountingyear)=" & Year(Session("v_Costbook")) 
					 
		connstr.execute(qNormalisasiDetail)	
		
		qNormalisasiDetail=" DELETE FROM Cost_Activity WHERE  fcba='" & Session("fcBA") & "' AND to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Costbook")) &  _
					 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Costbook")) & " and journalprefix='HPT'"  & _
					 " AND to_number(accountingperiod)=" & Month(Session("v_Costbook")) & "" &  _
			         " AND to_number(accountingyear)=" & Year(Session("v_Costbook")) 
		connstr.execute(qNormalisasiDetail)	
		
		Response.Write  "<font face=Verdana size=1 color=Blue>" & "Normalisasi For Penjualan TBS Internal & HPP TBS ,success " & "<br/></font>" 
		
		strCursorSivCode = "SELECT customer_fcba,c.eplascode customercode,Sum(a.ffbweight) netto2 " & _
						    " FROM lhpffbwbtrn a " & _
						    " left JOIN lhptikethdr b ON a.docketid=b.docketid " & _
						    " left JOIN lhpareamst c ON  a.areaid=c.areaid " & _
						    " left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
						    " WHERE  To_Char(docketdate,'yyyymm')='" & Year(Session("v_Costbook")) & "'||LPAD('" & Month(Session("v_Costbook")) & "', 2, 0)" & _
                            " AND customer_fcba='" & Session("fcBA") & "' AND own_estate='Y' " & _
					        " GROUP BY c.eplascode ,customer_fcba "

						
		Set rsCursorSivCode = Server.CreateObject("ADODB.Recordset")
		 rsCursorSivCode.CursorType = 3
		 rsCursorSivCode.CursorLocation = 3
		 rsCursorSivCode.Open strCursorSivCode, connstr		

		 if not rsCursorSivCode.eof then
			 Do While Not rsCursorSivCode.eof
			 
		
		
		 qCostAMax =  "select nvl(max(lineno), 0) maxLineNo from cost_activity  where fcba = '" & rsCursorSivCode("customer_fcba") & "'" 
		 
		 Set rsCostAMax = Server.CreateObject("ADODB.Recordset")
		 rsCostAMax.Open qCostAMax, connstr
		 If Not rsCostAMax.EOF Then
		 maxLNO = cdbl(rsCostAMax("maxLineNo"))
		 Else
		 maxLNO = 0
		 End If
		
		 rsCostAMax.close
		 set rsCostAMax = Nothing
		
		 maxLNO = maxLNO + 1
		
		Response.Write  "<font face=Verdana size=1 color=Green>" & "Prosses Penjualan TBS Internal " & rsCursorSivCode("customer_fcba") & "<br/></font>" 
		Response.Flush
				
		 connstr.BeginTrans
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
						 " SELECT to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'), decode(own_estate,'Y','2110608','2110609') jobcode,'AP' Source,'AP' SOURCETYPE,'SFB',decode(own_estate,'Y','OWN','EXT') SOURCECODE_STATUS , " & _ 
						 " 'CB' TARGETTYPE ,'BS' TARGETCODE,'','', " & _ 
						 " 0 QUANTITY, 0, amount ,'" & rsCursorSivCode("customer_fcba") & "',to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'),sysdate, " & _ 
						 " '','', " & _ 
						 " '',decode(own_estate,'Y','OWN','EXT'), " & _
						 " '','', " & _
						 " 'own_estate',decode(own_estate,'Y','OWN','EXT')||'" & rsCursorSivCode("customercode") & "', " & _
						 " '','','Dr', " & _
						 "" & Month(Session("v_Costbook")) & " , " & Year(Session("v_Costbook")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'ARF' " & _
						 " FROM " & _
						 " ( SELECT customer_fcba,own_estate,SUM(amount) amount FROM " & _
						 " 	( " & _
						 " 	SELECT customer_fcba,own_estate,netto2*Nvl(price,0) amount " & _ 
						 " 	FROM " & _
						 " 	( " & _
						 "		SELECT customer_fcba,own_estate ,c.eplascode customercode,d.description customername,docketdate datein,Sum(a.ffbweight) netto2 " & _
						 "		FROM lhpffbwbtrn a " & _
						 "		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _
						 "		left JOIN lhpareamst c ON  a.areaid=c.areaid " & _
						 "		left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
						 "		WHERE  To_Char(docketdate,'yyyymm')='" & Year(Session("v_Costbook")) & "'||LPAD('" & Month(Session("v_Costbook")) & "', 2, 0)" & _
                         "			AND customer_fcba='" & Session("fcBA") & "' AND own_estate='Y' AND customer_fcba = '" & rsCursorSivCode("customer_fcba") & "'" & _
						 "		GROUP BY own_estate ,c.eplascode ,d.description ,docketdate ,customer_fcba" & _
						 " 	)x " & _
						 " 	LEFT JOIN Invoice_pricetbs y ON x.customercode=y.suppliercode AND x.datein=y.ondate " & _
						 " 	)Z GROUP BY own_estate,customer_fcba " & _ 
						 " )j "
						
		
		 connstr.execute(qCost_Activity)
							
		 If Err.Number = 0 Then
			 connstr.CommitTrans
			 Response.Flush
		 Else
			 connstr.RollbackTrans
			 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			 Response.Flush
		 End If
		
		 maxLNO = maxLNO + 1
		 connstr.BeginTrans
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
						 " SELECT to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'), 5110106 jobcode,'AR' Source,'AR' SOURCETYPE,'" & rsCursorSivCode("customercode") & "',decode(own_estate,'Y','OWN','EXT') SOURCECODE_STATUS , " & _ 
						 " 'CB' TARGETTYPE ,'BS' TARGETCODE,'','', " & _ 
						 " 0 QUANTITY, 0, amount * -1 ,'" & rsCursorSivCode("customer_fcba") & "',to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'),sysdate, " & _ 
						 " '','', " & _ 
						 " '',decode(own_estate,'Y','OWN','EXT'), " & _
						 " '','', " & _
						 " 'own_estate',decode(own_estate,'Y','OWN','EXT')||'" & rsCursorSivCode("customercode") & "', " & _
						 " '','','Cr', " & _
						 "" & Month(Session("v_Costbook")) & " , " & Year(Session("v_Costbook")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'ARF' " & _
						 " FROM " & _
						 " ( SELECT customer_fcba,own_estate,SUM(amount) amount FROM " & _
						 " 	( " & _
						 " 	SELECT customer_fcba,own_estate,netto2*Nvl(price,0) amount " & _ 
						 " 	FROM " & _
						 " 	( " & _
						 "		SELECT customer_fcba,own_estate ,c.eplascode customercode,d.description customername,docketdate datein,Sum(a.ffbweight) netto2 " & _
						 "		FROM lhpffbwbtrn a " & _
						 "		left JOIN lhptikethdr b ON a.docketid=b.docketid " & _
						 "		left JOIN lhpareamst c ON  a.areaid=c.areaid " & _
						 "		left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
						 "		WHERE  To_Char(docketdate,'yyyymm')='" & Year(Session("v_Costbook")) & "'||LPAD('" & Month(Session("v_Costbook")) & "', 2, 0)" & _
                         "			AND customer_fcba='" & Session("fcBA") & "' AND own_estate='Y' AND customer_fcba = '" & rsCursorSivCode("customer_fcba") & "'" & _
						 "		GROUP BY own_estate ,c.eplascode ,d.description ,docketdate ,customer_fcba" & _
						 " 	)x " & _
						 " 	LEFT JOIN Invoice_pricetbs y ON x.customercode=y.suppliercode AND x.datein=y.ondate " & _
						 " 	)Z GROUP BY own_estate,customer_fcba " & _ 
						 " )j "
		 connstr.execute(qCost_Activity)
							
		 If Err.Number = 0 Then
			 connstr.CommitTrans
			 Response.Flush
		 Else
			 connstr.RollbackTrans
			 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
			 Response.Flush
		 End If
		 
		rsCursorSivCode.MoveNext
		loop
		END IF
	    rsCursorSivCode.close
		set rsCursorSivCode=Nothing
		
		qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT fcba,accountingyear,accountingperiod,referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT fcba,accountingyear,accountingperiod,AUTO_REFERENCE_04 referenceno,journalprefix,company_api.get_companyno(businessunit_api.get_companycode(fcba))||'/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE accountingperiod="& Month(Session("v_Costbook")) & " AND accountingyear=" & Year(Session("v_Costbook")) & " AND journalprefix='ARF' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.fcba=y.fcba and  x.AUTO_REFERENCE_04 =  y.referenceno AND  x.accountingyear=y.accountingyear AND x.accountingperiod=y.accountingperiod AND x.journalprefix=y.journalprefix) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
					
		   
		connstr.execute(qjournal)
		
		Response.Write  "<font face=Verdana size=1 color=Blue>" & "Prosses Penjualan TBS Internal, success " & "<br/></font>" 
		Response.Flush
		
		
'--4. Journal HPP TBS INTERNAL (HPF-Harga Pokok  FruitBunch) 
'-- Temp Close 2023-04		 
		Response.Write  "<font face=Verdana size=1 color=Green>" & "Prosses HPP TBS Internal <br/></font>" 
		Response.Flush
		
		  maxLNO = maxLNO + 1
		 connstr.BeginTrans
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
						 " SELECT to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'), '5210201' jobcode,'PL' Source,'PL' SOURCETYPE,'HPP' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						 " 'CB' TARGETTYPE ,'BS' TARGETCODE,'' TARGETCODE_STATUS,'HPP/'||fcba REFERENCENO, " & _ 
						 " 0 QUANTITY, 0, amount ,fcba,to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'),sysdate, " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','HPP/'||fcba, " & _ 
						 " '','','Dr', " & _ 
						 "" & Month(Session("v_Costbook")) & " , " & Year(Session("v_Costbook")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'HPT' " & _
						 " FROM " & _ 
						 " (	SELECT fcba,Sum(amount) amount " & _ 
						 " 	FROM " & _ 
						 " 	( " & _ 
						 "		SELECT a.fcba,jobgroupcode,Sum(a.amount) amount " & _ 
						 "		FROM cost_activity a " & _ 
						 "		INNER JOIN job b ON a.jobcode=b.fccode  AND a.fcba=b.fcba " & _ 
						 "		INNER JOIN jobgroup c ON  b.jobgroupcode=c.fccode AND b.fcba=c.fcba " & _ 
                         " 		WHERE a.fcba = '" & session("fcba") & "'" & _ 
						 "      AND a.jobcode between '6110101' AND '6510104' " & _
                         " 		/*AND b.jobgroupcode IN ('6110000','6210000','6310000','6500100','6500200','6500300','6500400','6500500','6500600','6500700','6500800','6500900','6501000','6502100','6502200','6400000','6509100','6510100') */" & _ 
						 " 		AND To_Char(a.tdate,'yyyymm')='" & Year(Session("v_Costbook")) & "'||LPAD('" & Month(Session("v_Costbook")) & "', 2, 0) " & _
						 " 		GROUP BY  a.fcba,jobgroupcode " & _ 
						 "  )x " & _
						 " GROUP BY fcba " & _
						 " )j "

                  
          		 connstr.execute(qCost_Activity)
							
				 If Err.Number = 0 Then
					 connstr.CommitTrans
					 Response.Flush
				 Else
					 connstr.RollbackTrans
					 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
					 Response.Flush
				 End If


 		 maxLNO = maxLNO + 1
		 connstr.BeginTrans
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
						 " SELECT to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'), '6599999' jobcode,'PL' Source,'PL' SOURCETYPE,'HPP' SOURCECODE,'' SOURCECODE_STATUS , " & _ 
						 " 'CB' TARGETTYPE ,'BS' TARGETCODE,'' TARGETCODE_STATUS,'HPP/'||fcba REFERENCENO, " & _ 
						 " 0 QUANTITY, 0, amount *-1,fcba,to_date('" & Session("v_Costbook") & "','mm/dd/yyyy'),sysdate, " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','', " & _ 
						 " '','HPP/'||fcba, " & _ 
						 " '','','Cr', " & _ 
						 "" & Month(Session("v_Costbook")) & " , " & Year(Session("v_Costbook")) & " ,sysdate, to_char(sysdate,'HH24:MI'), " & maxLNO & "+rownum ,'HPT' " & _
						 " FROM " & _ 
						 " (	SELECT fcba,Sum(amount) amount " & _ 
						 " 	FROM " & _ 
						 " 	( " & _ 
						 "		SELECT a.fcba,jobgroupcode,Sum(a.amount) amount " & _ 
						 "		FROM cost_activity a " & _ 
						 "		INNER JOIN job b ON a.jobcode=b.fccode  AND a.fcba=b.fcba " & _ 
						 "		INNER JOIN jobgroup c ON  b.jobgroupcode=c.fccode AND b.fcba=c.fcba " & _ 
                         " 		WHERE a.fcba = '" & session("fcba") & "'" & _ 
						 "      AND a.jobcode between '6110101' AND '6510104' " & _
                         " 		/*AND b.jobgroupcode IN ('6110000','6210000','6310000','6500100','6500200','6500300','6500400','6500500','6500600','6500700','6500800','6500900','6501000','6502100','6502200','6400000','6509100','6510100') */" & _ 
						 " 		AND To_Char(a.tdate,'yyyymm')='" & Year(Session("v_Costbook")) & "'||LPAD('" & Month(Session("v_Costbook")) & "', 2, 0) " & _
						 " 		GROUP BY  a.fcba,jobgroupcode " & _ 
						 "  )x " & _
						 " GROUP BY fcba " & _
						 " )j "

          		 connstr.execute(qCost_Activity)
							
				 If Err.Number = 0 Then
					 connstr.CommitTrans
					 Response.Flush
				 Else
					 connstr.RollbackTrans
					 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_Activity&"] ("&Err.Description&") <br/></font>")
					 Response.Flush
				 End If

				
		    qjournal=" MERGE INTO cost_activity x " & _
					 " USING " & _
					 " ( " & _
					 "  SELECT fcba,accountingyear,accountingperiod,referenceno,journalprefix,idx||LPAD(ROWNUM,4,0) seq from " & _
					 "  ( " & _
					 "  SELECT DISTINCT fcba,accountingyear,accountingperiod,AUTO_REFERENCE_04 referenceno,journalprefix,company_api.get_companyno(businessunit_api.get_companycode(fcba))||'/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					 "  FROM cost_activity WHERE accountingperiod="& Month(Session("v_Costbook")) & " AND accountingyear=" & Year(Session("v_Costbook")) & " AND journalprefix='HPT' " & _
					 "  AND fcba='" & session("fcba") & "'" & _
					 "  Order BY referenceno " & _
					 "  ) " & _
					 " )y ON (x.fcba=y.fcba and  x.AUTO_REFERENCE_04 =  y.referenceno AND  x.accountingyear=y.accountingyear AND x.accountingperiod=y.accountingperiod AND x.journalprefix=y.journalprefix) " & _
					 " WHEN MATCHED THEN " & _
					 " UPDATE SET x.journalno = y.seq " 
		
		          connstr.execute(qjournal)
		
		         Response.Write  "<font face=Verdana size=1 color=Blue>" & "Prosses HPP TBS Internal, success " & "<br/></font>" 
				 Response.Flush

END IF


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
	
	 Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear 
	 Dim temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear, temp_EndingAccountingPeriod
	 Dim temp_BeforeAccountingPeriod, temp_BeforeAccountingYear
	 
	 
		
	
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
			temp_NextCurrentAccountingYear =  rsAccountPeriod("NextCurrentAccountingYear")
		else
			Response.Write  "<font face=Verdana size=1 color=red>" & "IPLAS Error, No PERIODCONTROL for COSTBOOK " & "<br>"
		end if
		rsAccountPeriod.close
		
		'temp_NextCurrentAccountingPeriod = CInt (temp_NextCurrentAccountingPeriod) - 1 
		
		set rsAccountPeriod = Nothing	
		
		if temp_NextCloseAccountingPeriod = "1" Then 
			temp_BeforeAccountingPeriod = "15" 
		else
			temp_BeforeAccountingPeriod = CInt (temp_NextCloseAccountingPeriod) - 1
		end if
		
		if temp_NextCloseAccountingPeriod = "1" Then 
			temp_BeforeAccountingYear = CInt (temp_NextCloseAccountingYear) - 1 
		else
			temp_BeforeAccountingYear = temp_NextCloseAccountingYear
		end if
		

	if txtControlJob<>"" then
	
	response.write "Before:"  & temp_BeforeAccountingYear & "/" & temp_BeforeAccountingPeriod & "</br>"
	response.write "Current:"  & temp_NextCurrentAccountingPeriod & "/" & temp_NextCurrentAccountingYear & "</br>"
	response.write "Next : " & temp_NextCloseAccountingPeriod & "/" & temp_NextCloseAccountingYear & "</br>"
	
	    'UPDATE CHARGEUNIT
		qNormalisasi=" UPDATE cost_activity SET CHARGEUNIT='' " & _ 
		             " WHERE To_Char(tdate,'yyyymm')='" & Year(Session("v_Costbook")) & "'||LPAD('" & Month(Session("v_Costbook")) & "', 2, 0) " & _
					 " and fcBA='" & Session("fcBA") & "'" & _
					 " and jobcode NOT IN (SELECT JobIntraco FROM Businessunit WHERE fccode='" & Session("fcBA") & "')" 
		connstr.execute(qNormalisasi)
		
		'---- Normalisasi
		qNormalisasi=" Delete From TrialBalance where  Period_Month=" & temp_NextCurrentAccountingPeriod &  _
					 " and Period_Year=" & temp_NextCurrentAccountingYear & " and fcBA='" & Session("fcBA") & "'"
		response.write "1-" & qNormalisasi & "<br>"
		connstr.execute(qNormalisasi)	
		
		qNormalisasi=" Delete From TrialBalance where  Period_Month=" & temp_NextCloseAccountingPeriod &  _
					 " and Period_Year=" & temp_NextCloseAccountingYear & " and fcBA='" & Session("fcBA") & "'"
		response.write "2-" & qNormalisasi & "<br>"
		connstr.execute(qNormalisasi)	
		
		Response.Write  "<font face=Verdana size=1 color=Blue>" & "FCBA :  " & Session("fcBA")  & "<br>"
		Response.Write  "<font face=Verdana size=1 color=Blue>" & "Periode :  " & temp_NextCurrentAccountingPeriod & "/" & temp_NextCurrentAccountingYear  & "<br>"
		Response.Write  "<font face=Verdana size=1 color=Blue>" & "Normalisasi  " & "<br>"
		Response.Flush() 
		
		'response.write qNormalisasi
		
		 Response.Write "Generated TrialBalance " & "<br>"
     	 Response.Flush()

		
		qUpdate=" Insert into TrialBalance " & _
				"(PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME)" & _
				" SELECT "& temp_NextCloseAccountingPeriod & "," & temp_NextCloseAccountingYear & ",a.fccode,a.fcname,Nvl(nOpening,0) nOpening ,Nvl(c.nDebet,0) nDebet ,Nvl(d.nCredit,0) nCredit,Nvl(nOpening,0)+Nvl(c.nDebet,0)+Nvl(d.nCredit,0) AS nEnding, " & _
				" '"  & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', SYSDATE, TO_CHAR(SYSDATE,'HH24:MI')  " & _ 
				" FROM JOB a  " & _
				" LEFT JOIN (SELECT fcba,fccode,Nvl(ending_balance,0) nOpening FROM trialbalance WHERE fcba='" & Session("fcBA")  & "' AND period_year=" & temp_BeforeAccountingYear & " AND period_month= " & temp_BeforeAccountingPeriod & ") b ON a.fcba=b.fcba AND a.fccode=b.fccode " & _
				" LEFT JOIN (SELECT fcba,JobCode,nvl(Sum(AMOUNT),0) as nDebet FROM Cost_Activity  WHERE  Amount > 0 AND fcba='" & Session("fcBA")  & "' AND AccountingYear=" & temp_NextCloseAccountingYear & " AND AccountingPeriod=" & temp_NextCloseAccountingPeriod & " GROUP BY fcba,JobCode) c ON a.fcba=c.fcba AND a.fccode=c.jobcode " & _
				" LEFT JOIN (SELECT fcba,JobCode,nvl(Sum(AMOUNT),0) as nCredit FROM Cost_Activity  WHERE  Amount < 0 AND fcba='" & Session("fcBA")  & "' AND AccountingYear=" & temp_NextCloseAccountingYear & " AND AccountingPeriod=" & temp_NextCloseAccountingPeriod & " GROUP BY fcba,JobCode) d ON a.fcba=d.fcba AND a.fccode=d.jobcode " & _
				" WHERE a.fcba='" & Session("fcBA")  &"' " & _
				" ORDER BY a.fccode " 
	    
		response.write "3-" & qUpdate & "<br>"
		connstr.execute(qUpdate)	
				  

		qUpdate=" Insert into TrialBalance " & _
				" (PERIOD_MONTH, PERIOD_YEAR, FCCODE, FCNAME, OPENING_BALANCE, DEBET, CREDIT, ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				"  (select " &  temp_NextCurrentAccountingPeriod & " as PERIOD_MONTH,"  & temp_NextCurrentAccountingYear & " as PERIOD_YEAR" & ", FCCODE," & _
				"  FCNAME, ENDING_BALANCE as OPENING_BALANCE,0 as DEBET, 0 as CREDIT,  0 as ENDING_BALANCE, FCENTRY, FCEDIT, FCIP, FCBA, SYSDATE, TO_CHAR(SYSDATE,'HH24:MI') " & _
				"  From TrialBalance Where Period_Month=" & temp_NextCloseAccountingPeriod &  _
				" and Period_Year=" & temp_NextCloseAccountingYear & " and fcBA='" & Session("fcBA") & "')"		 
		
		response.write "4-" & qUpdate & "<br>"
		connstr.execute(qUpdate)
		
		' Add By Akbar 18 Desemeber 2015
		
    ' AMH Apr 21 Pindahkan kesini 
	'**** Eksekusi ERP Report ****'
	'Server.Execute("p_Generate_ERP_Report.asp")
	
	
	'**** Eksekusi AR Report ****'
	'Server.Execute("p_Generate_Account_Report.asp")
        
		
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

	strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME,fcentry) VALUES (" & _
					" 'Costbook', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" " & tempMaxLine & "" & _
					", SYSDATE, TO_CHAR(SYSDATE,'HH24:MI') " & _
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
	
	if Session("Estate_Mill") = "E"  AND temp_NextCloseAccountingPeriod <= 12 then
%>
	<!--#include virtual="/insert_nextmonthbjr.asp"-->
<%
    end if
	
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