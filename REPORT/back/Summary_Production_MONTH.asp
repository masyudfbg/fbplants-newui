<%@Language=VBScript%>

<%
' Option Explicit 
%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	' DateTo and DateFrom will be EQUAL to remove miss-understanding
	
	Dim txtToDayDate, txtyear, txtmonth, txtfcba
	Dim strSQLPreParameter, rsSQLPreParameter, v_txtToDayDate, v_txtfcba, v_txtyear, v_txtmonth
	Dim strProduction, rsProduction
	Dim tempRefCode, strHRef, rsHRef, tempDescription, tempUOM
	Dim tempGrading_HI, tempGrading_SHI, tempGrading_SBI
	Dim tempDecrease_hi, tempDecrease_shi, tempDecrease_sbi
	Dim tempOpeningCPOStock_HI, tempDespatchCPOStock_HI
	Dim tempOpeningKernelStock_HI, tempDespatchKernelStock_HI
	Dim tempRounding
    Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i,tglinput
	
	' DateTo and DateFrom will be EQUAL to remove miss-understanding
	
	'txtdatefrom			= Request("txtdatefrom")
	'txtdateto			= Request("txtdatefrom")
    'txtfcba 			= Request("txtfcba")
	txtyear 			= request.Form("v_txtyear")
	txtmonth 			= request.Form("v_txtmonth")
    
	txtToDayDate		= Request("txtToDayDate")
    txtfcba 			= Request("txtfcba")	

	'strSQLPreParameter = "select to_char(to_date('" & txtToDayDate & "', 'MM/DD/YYYY'),'YYYY') as CurrentYear, " & _
	 '      "to_char(to_date('" & txtToDayDate & "', 'MM/DD/YYYY'),'MM') as CurrentPeriod from dual "
	
	'Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
	'rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 
	
    'if not rsSQLPreParameter.EOF then
	'	txtyear = rsSQLPreParameter("CurrentYear")
	'	txtmonth = rsSQLPreParameter("CurrentPeriod")
	'end if
	
	
	
	tglinput = "10-" & txtmonth & "-" & txtyear
	strSQLPreParameter = "select TRUNC(to_date('"& tglinput & "','dd-mm-yyyy'), 'month') firstday, to_date('3-10-2012','dd-mm-yyyy') today, last_day(to_date('"& tglinput & "','dd-mm-yyyy')) lastday from dual"
	Set rsSQLPreParameter = Server.CreateObject("ADODB.Recordset")
	rsSQLPreParameter.Open strSQLPreParameter, connstr, adOpenStatic 
	
    if not rsSQLPreParameter.EOF then
	    txtdatefrom	= rsSQLPreParameter("firstday")
		txtdateto     = rsSQLPreParameter("lastday")
	end if
	
	v_txtToDayDate		=	txtdateto
	v_txtdatefrom		=	txtdatefrom
	v_txtfcba     		=   txtfcba
	v_txtyear 			=   txtyear
	v_txtmonth			=   txtmonth
	
	
	
	
	
	
	' First

    ' To Handle the Query of Production Grading
	
	strTerimaTbs = 	" select b.theHeadGroup, b.theGroupName,  a.desccode, b.UoM UnitOfMeasurement, a.hi, a.shi, a.sbi " & _
				    " from ( " & _
	                     " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB02' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB02' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB02' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY DescCode " & _
						" UNION " & _
	                     " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB03' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001', '20002', '30001') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB03' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001', '20002', '30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB03' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001', '20002', '30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY DescCode " & _
						" UNION " & _
						  " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB0301' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0301' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0301' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY  DescCode " & _
						" UNION " & _
						 " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB0302' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20002') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0302' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20002') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0302' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20002') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY DescCode " & _
						" UNION " & _
						 " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB0303' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('30001') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0303' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0303' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY  DescCode " & _
						" UNION " & _
						 " SELECT  DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB04' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode not in ('20001','20002','30001') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB04' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode not in ('20001','20002','30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB04' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode not in ('20001','20002','30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY DescCode " & _
						" ) a, " & _
						" (select parheadcode theHeadGroup, fccode theGroup, fcname theGroupName, rate_1 OrderData, param_1 UoM from  " & _
							" parameterdetail " & _
							" where " & _
							" parcode = 'SR005'  and " & _
							" parheadcode = 'TERIMA_TBS' and " & _
							" param_2 = 'DEFAULT' and " & _
							" fcba = '" & v_txtFcba & "' " & _
						" ) b " & _
						" where " & _
						" a.desccode = b.theGroup  " & _
						" order by b.orderdata "
						

						
 	Set rsTerimaTBS = Server.CreateObject("ADODB.Recordset")
	
	rsTerimaTBS.Open strTerimaTBS, connstr, adOpenStatic 		
						
	strSortasi = 	" select b.theHeadGroup, b.theGroupName,  a.desccode, b.UoM UnitOfMeasurement, a.hi, a.shi, a.sbi " & _
				    " from ( " & _
					" SELECT 'FFB07' desccode, SUM (hi) hi, SUM (shi) shi, SUM (sbi) sbi " & _
					  " FROM (SELECT SUM (weight) hi, 0 shi, 0 sbi " & _
							  " FROM gradingrecord " & _
							 " WHERE datetime = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								   " AND fcba = '" & v_txtfcba & "' " & _
							" UNION " & _
							" SELECT 0 hi, SUM (weight) shi, 0 sbi " & _
							  " FROM gradingrecord " & _
							 " WHERE     datetime <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								   " AND TO_CHAR (datetime, 'YYYY') = '" & v_txtYear & "' " & _
								   " AND TO_CHAR (datetime, 'MM') = '" & v_txtMonth & "' " & _ 
								   " AND fcba = '" & v_txtfcba & "'" & _ 
							" UNION " & _ 
							" SELECT 0 hi, 0 shi, SUM (weight) sbi " & _
							  " FROM gradingrecord " & _
							 " WHERE     datetime <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								   " AND TO_CHAR (datetime, 'YYYY') = '" & v_txtYear & "' " & _
								   " AND fcba = '" & v_txtfcba & "') " & _
							" ) a, " & _
						" (select parheadcode theHeadGroup, fccode theGroup, fcname theGroupName, rate_1 OrderData, param_1 UoM from  " & _
							" parameterdetail " & _
							" where " & _
							" parcode = 'SR005'  and " & _
							" parheadcode = 'SORTASI' and " & _
							" param_2 = 'DEFAULT' and " & _
							" fcba = '" & v_txtFcba & "' " & _
						" ) b " & _
						" where " & _
						" a.desccode = b.theGroup  " & _
						" order by b.orderdata "								   
	
 	Set rsSortasi = Server.CreateObject("ADODB.Recordset")
	rsSortasi.Open strSortasi, connstr, adOpenStatic 							
		
	
    strProses = 	" select b.theHeadGroup, b.theGroupName,  a.desccode, b.UoM UnitOfMeasurement, a.hi, a.shi, a.sbi " & _
				    " from ( " & _
	                        " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB0701' DescCode, " & _
										 " SUM (cageweight) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM shiftreport " & _
								   " WHERE productiondate = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
								  " UNION " & _
								  " SELECT 'FFB0701' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (cageweight) SHI, " & _
										 " 0 SBI " & _
									" FROM shiftreport " & _
								   " WHERE     TO_CHAR (productiondate, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (productiondate, 'MM') = '" & v_txtMonth & "' " & _
										 " AND productiondate <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
								  " UNION " & _
								  " SELECT 'FFB0701' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (cageweight) SBI " & _
									" FROM shiftreport " & _
								   " WHERE     TO_CHAR (productiondate, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND productiondate <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
										 " AND fcba = '" & v_txtFcba & "') " & _
						" GROUP BY DescCode " & _
						" ) a, " & _
						" (select parheadcode theHeadGroup, fccode theGroup, fcname theGroupName, rate_1 OrderData, param_1 UoM from  " & _
							" parameterdetail " & _
							" where " & _
							" parcode = 'SR005'  and " & _
							" parheadcode = 'TBS_PROSES' and " & _
							" param_2 = 'DEFAULT' and " & _
							" fcba = '" & v_txtFcba & "' " & _
						" ) b " & _
						" where " & _
						" a.desccode = b.theGroup  " & _
						" order by b.orderdata "
						
	Set rsProses = Server.CreateObject("ADODB.Recordset")
	rsProses.Open strProses, connstr, adOpenStatic 			


    strKapasitasOlah = " select b.theHeadGroup, b.theGroupName,  a.desccode, b.UoM UnitOfMeasurement, a.hi, a.shi, a.sbi " & _
				    " from ( " & _
                      	" SELECT desccode, " & _ 
							 " SUM (hi) hi, " & _ 
							 " SUM (shi) shi, " & _ 
							 " SUM (sbi) sbi " & _ 
						" FROM (SELECT 'FFB09' DescCode," & _ 
									 " SUM (totalHourShift1 + totalHourShift2) hi, " & _ 
									 " 0 shi, " & _ 
									 " 0 sbi " & _ 
								" FROM view_shifthourrange " & _ 
							   " WHERE fcba = '" & v_txtFcba & "' " & _ 
									 " AND TO_CHAR (productiondate, 'MM/DD/YYYY') = '" & v_txtToDayDate & "' " & _ 
							  " UNION " & _ 
							  " SELECT 'FFB09' DescCode, " & _ 
									 " 0 hi, " & _ 
									 " SUM (totalHourShift1 + totalHourShift2) shi, " & _ 
									 " 0 sbi " & _ 
								" FROM view_shifthourrange " & _ 
							   " WHERE     fcba = '" & v_txtFcba & "' " & _ 
									 " AND TO_CHAR (productiondate, 'MM/DD/YYYY') <= '" & v_txtToDayDate & "' " & _ 
									 " AND TO_CHAR (productiondate, 'MM') = '" & v_txtMonth & "' " & _ 
									 " AND TO_CHAR (productiondate, 'YYYY') = '" & v_txtYear & "' " & _ 
							  " UNION " & _ 
							  " SELECT 'FFB09' DescCode, " & _ 
									 " 0 hi, " & _ 
									 " 0 shi, " & _ 
									 " SUM (totalHourShift1 + totalHourShift2) sbi " & _ 
								" FROM view_shifthourrange " & _ 
							   " WHERE     fcba = '" & v_txtFcba & "' " & _ 
									 " AND TO_CHAR (productiondate, 'MM/DD/YYYY') <= '" & v_txtToDayDate & "' " & _ 
									 " AND TO_CHAR (productiondate, 'YYYY') = '" & v_txtYear & "') " & _ 
					" GROUP BY desccode " & _
						" ) a, " & _
						" (select parheadcode theHeadGroup, fccode theGroup, fcname theGroupName, rate_1 OrderData, param_1 UoM from  " & _
							" parameterdetail " & _
							" where " & _
							" parcode = 'SR005'  and " & _
							" parheadcode = 'KAPASITAS_OLAH' and " & _
							" param_2 = 'DEFAULT' and " & _
							" fcba = '" & v_txtFcba & "' " & _
						" ) b " & _
						" where " & _
						" a.desccode = b.theGroup  " & _
						" order by b.orderdata "					
					
	Set rsKapasitasOlah = Server.CreateObject("ADODB.Recordset")
	rsKapasitasOlah.Open strKapasitasOlah, connstr, adOpenStatic 		



	strSisaTBS = " select b.theHeadGroup, b.theGroupName,  a.desccode, b.UoM UnitOfMeasurement, a.hi, a.shi, a.sbi " & _
				    " from ( " & _
					" SELECT 'FFB01' DescCode, " & _
							   " SUM (rampbalance + cagesbalance + platformbalance) HI," & _
							   " 0 SHI," & _
							   " 0 SBI " & _
						  " FROM shiftreport " & _
						 " WHERE     productiondate = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') - 1 " & _
							   " AND fcba = '" & v_txtFcba & "' " & _
							   " AND shift_shiftcode = 'SHIFT2' " & _
						" ) a, " & _
						" (select parheadcode theHeadGroup, fccode theGroup, fcname theGroupName, rate_1 OrderData, param_1 UoM from  " & _
							" parameterdetail " & _
							" where " & _
							" parcode = 'SR005'  and " & _
							" parheadcode = 'SISA_TBS' and " & _
							" param_2 = 'DEFAULT' and " & _
							" fcba = '" & v_txtFcba & "' " & _
						" ) b " & _
						" where " & _
						" a.desccode = b.theGroup  " & _
						" order by b.orderdata "		

	Set rsSisaTBS = Server.CreateObject("ADODB.Recordset")
	rsSisaTBS.Open strSisaTBS, connstr, adOpenStatic 		
	
	
	
	strOilKernelProduction = " select b.theHeadGroup, b.theGroupName,  a.desccode, b.UoM UnitOfMeasurement, a.hi, a.shi, a.sbi " & _
				    " from ( " & _
	               "  SELECT desccode, " & _
						" SUM (hi) hi, " & _
						" SUM (shi) shi, " & _
						" SUM (sbi) sbi " & _
					" FROM (SELECT 'PROD01' DescCode, " & _
								 " SUM (quantity) hi, " & _
								 " 0 shi, " & _
								 " 0 sbi " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND pstorage_storagecode = 'CIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
						  " UNION " & _
						  " SELECT 'PROD01' DescCode, " & _
								 " 0 hi, " & _
								 " SUM (quantity) shi, " & _
								 " 0 sbi  " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND TO_CHAR (readingdate, 'YYYY') = '" & v_txtYear & "' " & _
								 " AND TO_CHAR (readingdate, 'MM') = '" & v_txtMonth & "' " & _
								 " AND pstorage_storagecode = 'CIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
						  " UNION " & _
						  " SELECT 'PROD01' DescCode, " & _
								 " 0 hi, " & _
								 " 0 shi, " & _
								 " SUM (quantity) sbi " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND TO_CHAR (readingdate, 'YYYY') = '" & v_txtYear & "' " & _
								 " AND pstorage_storagecode = 'CIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
							" UNION " & _	 
							" SELECT 'PROD02' DescCode, " & _
								 " SUM (quantity) hi, " & _
								 " 0 shi, " & _
								 " 0 sbi " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND pstorage_storagecode = 'KIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
						  " UNION " & _
						  " SELECT 'PROD02' DescCode, " & _
								 " 0 hi, " & _
								 " SUM (quantity) shi, " & _
								 " 0 sbi  " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND TO_CHAR (readingdate, 'YYYY') = '" & v_txtYear & "' " & _
								 " AND TO_CHAR (readingdate, 'MM') = '" & v_txtMonth & "' " & _
								 " AND pstorage_storagecode = 'KIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
						  " UNION " & _
						  " SELECT 'PROD02' DescCode, " & _
								 " 0 hi, " & _
								 " 0 shi, " & _
								 " SUM (quantity) sbi " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND TO_CHAR (readingdate, 'YYYY') = '" & v_txtYear & "' " & _
								 " AND pstorage_storagecode = 'KIP1' " & _
								 " AND fcba = '" & v_txtfcba & "'	 ) " & _
				      " GROUP BY desccode " & _
					  " ) a, " & _
						" (select parheadcode theHeadGroup, fccode theGroup, fcname theGroupName, rate_1 OrderData, param_1 UoM from  " & _
							" parameterdetail " & _
							" where " & _
							" parcode = 'SR005'  and " & _
							" parheadcode = 'PRODUCTION' and " & _
							" param_2 = 'DEFAULT' and " & _
							" fcba = '" & v_txtFcba & "' " & _
						" ) b " & _
						" where " & _
						" a.desccode = b.theGroup  " & _
						" order by b.orderdata "		
	
	
 	Set rsOilKernelProduction = Server.CreateObject("ADODB.Recordset")
	
	rsOilKernelProduction.Open strOilKernelProduction, connstr, adOpenStatic 	
	
	

	strCPOStock = " select b.theHeadGroup, b.theGroupName,  a.desccode, b.UoM UnitOfMeasurement, a.hi, a.shi, a.sbi " & _
				    " from ( " & _
	                    " SELECT 'CPO001' DescCode, " & _
							   " SUM (Quantity) hi,  " & _
							   " 0 shi, " & _
							   " 0 sbi " & _
						  " FROM productstoragereading  " & _
						 " WHERE     pstorage_storagecode = 'BST1'  " & _
							   " AND readingdate = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY')  " & _
							   " AND fcba = '" & v_txtfcba & "'  " & _
						" UNION  " & _
						  " SELECT desccode,  " & _
								 " SUM (hi) hi,  " & _
								 " SUM (shi) shi,  " & _
								 " SUM (sbi) sbi  " & _
							" FROM (SELECT 'CPO002' DescCode,  " & _
										 " SUM (weightout - weightin) hi,  " & _
										 " 0 shi,  " & _
										 " 0 sbi  " & _
									" FROM weighbridgeticket  " & _
								   " WHERE     transactiontype = '2'  " & _
										 " AND product_productcode = 'CPO'  " & _
										 " AND dateout = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY')  " & _
										 " AND fcba = '" & v_txtfcba & "'  " & _
								  " UNION  " & _
								  " SELECT 'CPO002' DescCode,  " & _
										 " 0 hi,  " & _
										 " SUM (weightout - weightin) shi,  " & _
										 " 0 sbi  " & _
									" FROM weighbridgeticket  " & _
								   " WHERE     transactiontype = '2'  " & _
										 " AND product_productcode = 'CPO'  " & _
 										 " AND dateout <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY')  " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "'  " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "'  " & _
										 " AND fcba = '" & v_txtfcba & "'  " & _
								  " UNION  " & _
								  " SELECT 'CPO002' DescCode,  " & _
										 " 0 hi,  " & _
										 " 0 shi,  " & _
										 " SUM (weightout - weightin) sbi  " & _
									" FROM weighbridgeticket  " & _
								   " WHERE     transactiontype = '2'  " & _
										 " AND product_productcode = 'CPO'  " & _
										 " AND dateout <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY')  " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "'  " & _
										 " AND fcba = '" & v_txtfcba & "')  " & _
						" GROUP BY desccode  " & _
					  " ) a, " & _
						" (select parheadcode theHeadGroup, fccode theGroup, fcname theGroupName, rate_1 OrderData, param_1 UoM from  " & _
							" parameterdetail " & _
							" where " & _
							" parcode = 'SR005'  and " & _
							" parheadcode = 'CPOSTOCK' and " & _
							" param_2 = 'DEFAULT' and " & _
							" fcba = '" & v_txtFcba & "' " & _
						" ) b " & _
						" where " & _
						" a.desccode = b.theGroup  " & _
						" order by b.orderdata "							
	
	
 	Set rsCPOStock = Server.CreateObject("ADODB.Recordset")
	
	rsCPOStock.Open strCPOStock, connstr, adOpenStatic 		

	
	
	strKernelStock = " select b.theHeadGroup, b.theGroupName,  a.desccode, b.UoM UnitOfMeasurement, a.hi, a.shi, a.sbi " & _
				    " from ( " & _
	                    " SELECT 'PK001' DescCode, " & _
							   " SUM (Quantity) hi,  " & _
							   " 0 shi, " & _
							   " 0 sbi " & _
						  " FROM productstoragereading  " & _
						 " WHERE     pstorage_storagecode = 'BS1'  " & _
							   " AND readingdate = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY')  " & _
							   " AND fcba = '" & v_txtfcba & "'  " & _
						" UNION  " & _
						  " SELECT desccode,  " & _
								 " SUM (hi) hi,  " & _
								 " SUM (shi) shi,  " & _
								 " SUM (sbi) sbi  " & _
							" FROM (SELECT 'PK002' DescCode,  " & _
										 " SUM (weightout - weightin) hi,  " & _
										 " 0 shi,  " & _
										 " 0 sbi  " & _
									" FROM weighbridgeticket  " & _
								   " WHERE     transactiontype = '2'  " & _
										 " AND product_productcode = 'KER'  " & _
										 " AND dateout = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY')  " & _
										 " AND fcba = '" & v_txtfcba & "'  " & _
								  " UNION  " & _
								  " SELECT 'PK002' DescCode,  " & _
										 " 0 hi,  " & _
										 " SUM (weightout - weightin) shi,  " & _
										 " 0 sbi  " & _
									" FROM weighbridgeticket  " & _
								   " WHERE     transactiontype = '2'  " & _
										 " AND product_productcode = 'KER'  " & _
 										 " AND dateout <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY')  " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "'  " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtmonth & "'  " & _
										 " AND fcba = '" & v_txtfcba & "'  " & _
								  " UNION  " & _
								  " SELECT 'PK002' DescCode,  " & _
										 " 0 hi,  " & _
										 " 0 shi,  " & _
										 " SUM (weightout - weightin) sbi  " & _
									" FROM weighbridgeticket  " & _
								   " WHERE     transactiontype = '2'  " & _
										 " AND product_productcode = 'KER'  " & _
										 " AND dateout <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY')  " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "'  " & _
										 " AND fcba = '" & v_txtfcba & "')  " & _
						" GROUP BY desccode  " & _
					  " ) a, " & _
						" (select parheadcode theHeadGroup, fccode theGroup, fcname theGroupName, rate_1 OrderData, param_1 UoM from  " & _
							" parameterdetail " & _
							" where " & _
							" parcode = 'SR005'  and " & _
							" parheadcode = 'PKSTOCK' and " & _
							" param_2 = 'DEFAULT' and " & _
							" fcba = '" & v_txtFcba & "' " & _
						" ) b " & _
						" where " & _
						" a.desccode = b.theGroup  " & _
						" order by b.orderdata "							
	
	
 	Set rsKernelStock = Server.CreateObject("ADODB.Recordset")
	
	rsKernelStock.Open strKernelStock, connstr, adOpenStatic 		
	
	
	
	

    ' To Handle the Query of Production Grading
	strProductionGrading = 	" SELECT 'FFB01' DescCode, " & _
							   " SUM (rampbalance + cagesbalance + platformbalance) HI," & _
							   " 0 SHI," & _
							   " 0 SBI " & _
						  " FROM shiftreport " & _
						 " WHERE     productiondate = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') - 1 " & _
							   " AND fcba = '" & v_txtFcba & "' " & _
							   " AND shift_shiftcode = 'SHIFT2' " & _
						" UNION " & _
						 " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB02' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB02' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB02' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY DescCode " & _
						" UNION " & _
						  " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB03' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001', '20002', '30001') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB03' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001', '20002', '30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB03' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001', '20002', '30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY DescCode " & _
						" UNION " & _
						  " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB0301' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0301' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0301' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY  DescCode " & _
						" UNION " & _
						 " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB0302' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20002') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0302' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20002') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0302' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('20002') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY DescCode " & _
						" UNION " & _
						 " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB0303' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('30001') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0303' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB0303' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode in ('30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY  DescCode " & _
						" UNION " & _
						 " SELECT  DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB04' DescCode, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode not in ('20001','20002','30001') " & _
										 " AND dateout = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB04' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SHI, " & _
										 " 0 SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode not in ('20001','20002','30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (dateout, 'MM') = '" & v_txtMonth & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT 'FFB04' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (WEIGHTIN - WEIGHTOUT) SBI " & _
									" FROM WEIGHBRIDGETICKET " & _
								   " WHERE     transactiontype = '1' " & _
										 " AND product_productcode = 'FFB' " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
										 " AND cust_customercode not in ('20001','20002','30001') " & _
										 " AND TO_CHAR (dateout, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND dateout <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY')) " & _
						" GROUP BY DescCode " & _
						" UNION " & _
						  " SELECT DescCode, " & _
								 " SUM (HI) HI, " & _
								 " SUM (SHI) SHI, " & _
								 " SUM (SBI) SBI " & _
							" FROM (SELECT 'FFB07' DescCode, " & _
										 " SUM (cageweight) HI, " & _
										 " 0 SHI, " & _
										 " 0 SBI " & _
									" FROM shiftreport " & _
								   " WHERE productiondate = TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
								  " UNION " & _
								  " SELECT 'FFB07' DescCode, " & _
										 " 0 HI, " & _
										 " SUM (cageweight) SHI, " & _
										 " 0 SBI " & _
									" FROM shiftreport " & _
								   " WHERE     TO_CHAR (productiondate, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND TO_CHAR (productiondate, 'MM') = '" & v_txtMonth & "' " & _
										 " AND productiondate <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
										 " AND fcba = '" & v_txtFcba & "' " & _
								  " UNION " & _
								  " SELECT 'FFB07' DescCode, " & _
										 " 0 HI, " & _
										 " 0 SHI, " & _
										 " SUM (cageweight) SBI " & _
									" FROM shiftreport " & _
								   " WHERE     TO_CHAR (productiondate, 'YYYY') = '" & v_txtYear & "' " & _
										 " AND productiondate <= TO_DATE ('" & v_txtToDayDate & "', 'MM/DD/YYYY') " & _
										 " AND fcba = '" & v_txtFcba & "') " & _
						" GROUP BY DescCode " 
	
	
	

						
 	Set rsProductionGrading = Server.CreateObject("ADODB.Recordset")
	
	rsProductionGrading.Open strProductionGrading, connstr, adOpenStatic 

	
	
	strOilProduction = "  SELECT desccode, " & _
						" SUM (hi) hi, " & _
						" SUM (shi) shi, " & _
						" SUM (sbi) sbi " & _
					" FROM (SELECT 'PROD01' DescCode, " & _
								 " SUM (quantity) hi, " & _
								 " 0 shi, " & _
								 " 0 sbi " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND pstorage_storagecode = 'CIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
						  " UNION " & _
						  " SELECT 'PROD01' DescCode, " & _
								 " 0 hi, " & _
								 " SUM (quantity) shi, " & _
								 " 0 sbi  " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND TO_CHAR (readingdate, 'YYYY') = '" & v_txtYear & "' " & _
								 " AND TO_CHAR (readingdate, 'MM') = '" & v_txtMonth & "' " & _
								 " AND pstorage_storagecode = 'CIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
						  " UNION " & _
						  " SELECT 'PROD01' DescCode, " & _
								 " 0 hi, " & _
								 " 0 shi, " & _
								 " SUM (quantity) sbi " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND TO_CHAR (readingdate, 'YYYY') = '" & v_txtYear & "' " & _
								 " AND pstorage_storagecode = 'CIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
							" UNION " & _	 
							" SELECT 'PROD02' DescCode, " & _
								 " SUM (quantity) hi, " & _
								 " 0 shi, " & _
								 " 0 sbi " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate = TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND pstorage_storagecode = 'KIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
						  " UNION " & _
						  " SELECT 'PROD02' DescCode, " & _
								 " 0 hi, " & _
								 " SUM (quantity) shi, " & _
								 " 0 sbi  " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND TO_CHAR (readingdate, 'YYYY') = '" & v_txtYear & "' " & _
								 " AND TO_CHAR (readingdate, 'MM') = '" & v_txtMonth & "' " & _
								 " AND pstorage_storagecode = 'KIP1' " & _
								 " AND fcba = '" & v_txtfcba & "' " & _
						  " UNION " & _
						  " SELECT 'PROD02' DescCode, " & _
								 " 0 hi, " & _
								 " 0 shi, " & _
								 " SUM (quantity) sbi " & _
							" FROM productstoragereading " & _
						   " WHERE     readingdate <= TO_DATE ('" & v_txtTodayDate & "', 'MM/DD/YYYY') " & _
								 " AND TO_CHAR (readingdate, 'YYYY') = '" & v_txtYear & "' " & _
								 " AND pstorage_storagecode = 'KIP1' " & _
								 " AND fcba = '" & v_txtfcba & "'	 ) " & _
				      " GROUP BY desccode "
	
	
 	Set rsOilProduction = Server.CreateObject("ADODB.Recordset")
	
	rsOilProduction.Open strOilProduction, connstr, adOpenStatic 	
	
	


	
	


	
    %>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>IPla'S Reporting - Laporan Bulanan Produksi</title>

<script type='text/javascript' src='/js/x.js'></script>
<script type='text/javascript' src='/js/xtableheaderfixed.js'></script>
<script type='text/javascript'>
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site','tbl-container',0);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>

<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>
<form name="frmEntry" method="post" action="Summary_Production_MONTH.asp">

<body>
	<div id="wrapper">
	<div id="header-wrapper">
		<table id="menu-table" class="centered-margin">
			<tr>
				<td align="center"><a href="/login.asp">Back to Home</a></td>
				<td align="center"><a href="/report/DashboardGroupManager.asp">Back to Report</a></td>
			</tr>
		</table>
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
		</div>
		<div id="subheader" class="centered">
			<h2>LAPORAN BULANAN PABRIK</h2>
			<h5>TANGGAL:  <%=v_txtdatefrom%> - <%=v_txtToDayDate%>, PABRIK : <%=v_txtFcba%></h5>
		</div>
	</div>
	
	<table id="form-table" class="centered-margin">
		<tr>
			<th>TAHUN</th>
			<th>BULAN</th>
			<th>PABRIK</th>
			<th>&nbsp;</th>
		</tr>
		<tr>

			<td>
				<font size="2"  color="#0000FF" ></font>
                <% 
					MM=Request.Form("v_txtmonth")
					yyyy=Request.Form("v_txtyear")

					mon = Month(Session("v_Costbook"))
					If len(MM)=0 then MM=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = Yr - 2
					MaxYr = Yr  + 2
					If len(yyyy)= 0 then yyyy = Yr
				%>
				<select name="v_txtyear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
				</select>
			</td>
			
			<td>
				<font size="2" color="#0000FF" ></font>
					<select name="v_txtmonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
					<%
						Next
					%>
					</select>
			</td>	


			
			<td align="center">
				<select name="txtFcba">	
				<option value="">--ALL--</option>		
				<%
					Dim rsFcba
					Set rsFcba = Server.CreateObject("ADODB.Recordset")
					rsFcba.CursorLocation = adUseClient
					strsql =  " SELECT fccode FROM businessunit "
					rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not rsFcba.eof then
						DO while not rsFcba.eof
						if v_txtFcba = rsFcba("fccode") then
				%>			<option value="<%=rsFcba("fccode")%>" Selected><%=rsFcba("fccode")%></option>
						<%else%>
							<option value="<%=rsFcba("fccode")%>"><%=rsFcba("fccode")%></option>
						<%end if%>
				<%			rsFcba.movenext
						loop
					End if
					rsFcba.Close
					set rsFcba = nothing
				%>
				</select>		
			</td>
			<td align="center"><input type="submit" value="View" /></td>
		</tr>
	</table>
	
<div id="tbl-container">
	
	<table border="0" width="100%" id="product-table" class='xthf-site'>
<thead>	
		<tr>
			<th class="information" colspan="3"> INFORMASI </th>
			<th class="uom" rowspan="2" width="20px"> UOM </th>
			<th class="actual" colspan="2"> ACTUAL </th>
			<th class="budget" colspan="2"> TARGET-BUDGET-STANDARD </th>
			<th class="indicater" colspan="2"> INDICATOR </th>
		</tr>
		<tr>
			<th class="information" width="20px"> GROUP </th>
			<th class="information" width="20px" > NO. </th>
			<th class="information" width="200px"> DESKRIPSI </th>
			
			
			<th class="actual" width="20px"> BI  </th>
			<th class="actual" width="20px"> SBI  </th>
			
			
			<th class="budget" width="20px"> BI  </th>
			<th class="budget" width="20px"> SBI  </th>
			
			
			<th class="indicater" width="20px"> BI  </th>
			<th class="indicater" width="20px"> SBI  </th>
		</tr>

		<tr>
			<th class="separator" colspan="13" align="left"></th>				
		</tr>	

</thead>

	<!-- Start Terima TBS -->
<tbody>

	<%

		if not rsTerimaTBS.EOF then
			no = 0
			while not rsTerimaTBS.EOF
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				
				if rsTerimaTBS("DescCode") = "FFB0301" or rsTerimaTBS("DescCode") = "FFB0302" or rsTerimaTBS("DescCode") = "FFB0303" then 
				    no = no
				else
					no = no + 1
				end if
	%>				


				<tr class="<%=cClass%>">
				
					<%
					    strCodeDescHead = " select fcname from parameter " & _
										" where " & _
										" parcode = 'SR005' and " & _
										" fccode = '" & rsTerimaTBS("theHeadGroup") & "' and " & _
										" fcba = '" & v_txtFcba & "'"
										
						Set rsCodeDescHead = Server.CreateObject("ADODB.Recordset")	
						rsCodeDescHead.Open strCodeDescHead, connstr, adOpenStatic 		

						if not rsCodeDescHead.EOF then
							tempCodeHeadDescription = rsCodeDescHead ("fcname")
						else
							tempCodeHeadDescription = ""
						end if		
						rsCodeDescHead.Close
						set rsCodeDescHead = Nothing
					
					
						strCodeDesc = " select fcname codeDescription, rate_2 theRowSpan, rate_3 rounding, param_6 reference  " & _
									  " from parameterdetail " & _
									  "	where " & _
									    " parcode = 'SR005'  and " & _
										" parheadcode = 'TERIMA_TBS' and " & _
										" param_2 = 'DEFAULT' and " & _
										" fccode = '" & rsTerimaTBS("DescCode") & "'" 
					
					
					
						Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
						rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

						if not rsCodeDesc.EOF then
							tempCodeDescription = rsCodeDesc ("CodeDescription")
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = rsCodeDesc ("reference")
							tempRounding = cdbl(rsCodeDesc ("rounding"))
							
						else
							tempCodeDescription = ""					
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = ""
							tempRounding = 0
						end if		
					    tempUOM = rsTerimaTBS("UnitOfMeasurement")			

						rsCodeDesc.Close
						set rsCodeDesc = Nothing
										
										
					%>
					
					<% if no = 1 then %>
						<% if sPrevDesc <> tempCodeHeadDescription then %>				
							<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeHeadDescription%></td>
							<% sPrevDesc = tempCodeHeadDescription %>
						<% else %>
							<td align="left" ></td>
						<% end if %>
					<% end if %>
									
					<%
					if rsTerimaTBS("DescCode") = "FFB0301" or rsTerimaTBS("DescCode") = "FFB0302" or rsTerimaTBS("DescCode") = "FFB0303" then 
					%>
										
					<%
					else
					%>
						<td align="center"  ><%= no %></td>					
					<%
					end if
					%>
				
				
					
					
					<td align="left" ><a href="<%=tempReference%>"><%=tempCodeDescription%></a></td>
					<td align="center" ><%=tempUOM%></td>				
					
					<% ' Start Compare to Budget 

						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	

						if tempDataBudget_HI <> 0 then 
							tempCompare_HI = cdbl(rsTerimaTBS("hi"))/tempDataBudget_HI * 100
						else
							tempCompare_HI = 0
						end if
						
						if tempDataBudget_SHI <> 0 then 
							tempCompare_SHI = cdbl(rsTerimaTBS("shi"))/tempDataBudget_SHI * 100
						else
							tempCompare_SHI = 0
						end if

						if tempDataBudget_SBI <> 0 then 
							tempCompare_SBI = cdbl(rsTerimaTBS("sbi"))/tempDataBudget_SBI * 100
						else
							tempCompare_SBI = 0
						end if
						
					%>					
					
					<%  if rsTerimaTBS("hi") <>  "" then 
							tempData = cdbl(rsTerimaTBS("hi"))
						else
							tempData = 0
						end if
					%>
					
					<%
						if rsTerimaTBS("DescCode") = "FFB02" then 
						   tempReceived = tempData
					    end if
					%>
					
					<%  if rsTerimaTBS("shi") <>  "" then 
							tempData = cdbl(rsTerimaTBS("shi"))
						else
							tempData = 0
						end if
					%>
					<td align="center" ><%=FormatNumber(tempData,tempRounding)%></td>
					<%  if rsTerimaTBS("sbi") <>  "" then 
							tempData = cdbl(rsTerimaTBS("sbi"))
						else
							tempData = 0
						end if
					%>
					<td align="center" ><%=FormatNumber(tempData,tempRounding)%></td>
					
					

					
					
					
					
					
					
					<td align="center" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>

					<%
						if tempCompare_HI > 95 then
							cClassAssigned_HI = "green"
						else
							if tempCompare_HI > 85 then 
								 cClassAssigned_HI = "yellow"
							else
								 cClassAssigned_HI = "red"
							end if
						end if
						
						if tempCompare_SHI > 95 then
							cClassAssigned_SHI = "green"
						else
							if tempCompare_SHI > 85 then 
								 cClassAssigned_SHI = "yellow"
							else
								 cClassAssigned_SHI = "red"
							end if
						end if
						
						if tempCompare_SBI > 95 then
							cClassAssigned_SBI = "green"
						else
							if tempCompare_SBI > 85 then 
								 cClassAssigned_SBI = "yellow"
							else
								 cClassAssigned_SBI = "red"
							end if
						end if
						
					%>
					
					
					<td class="<%=cClassAssigned_SHI%>" align="center" ><%=FormatNumber(tempCompare_SHI,tempRounding)%></td>
					<td class="<%=cClassAssigned_SBI%>" align="center" ><%=FormatNumber(tempCompare_SBI,tempRounding)%></td>
					
				</tr>
	<%

				rsTerimaTBS.MoveNext
			Wend
		end if
		rsTerimaTBS.Close
		set rsTerimaTBS = Nothing
	%>
	
	
<!-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX -->	
	<tr>
		<th class="separator" colspan="10" align="left"></th>				
	</tr>	


	
	
	<!-- Start SORTASI -->	

	<%

		if not rsSortasi.EOF then
			no = 0
			while not rsSortasi.EOF
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				no = no + 1

	%>				


				<tr class="<%=cClass%>">
				
					<%
					    strCodeDescHead = " select fcname from parameter " & _
										" where " & _
										" parcode = 'SR005' and " & _
										" fccode = '" & rsSortasi("theHeadGroup") & "' and " & _
										" fcba = '" & v_txtFcba & "'"
										
						Set rsCodeDescHead = Server.CreateObject("ADODB.Recordset")	
						rsCodeDescHead.Open strCodeDescHead, connstr, adOpenStatic 		

						if not rsCodeDescHead.EOF then
							tempCodeHeadDescription = rsCodeDescHead ("fcname")
						else
							tempCodeHeadDescription = ""
						end if		
						rsCodeDescHead.Close
						set rsCodeDescHead = Nothing
					
					
						strCodeDesc = " select fcname codeDescription, rate_2 theRowSpan, rate_3 rounding, param_6 reference from parameterdetail " & _
									  "	where " & _
									    " parcode = 'SR005'  and " & _
										" parheadcode = 'SORTASI' and " & _
										" param_2 = 'DEFAULT' and " & _
										" fccode = '" & rsSortasi("DescCode") & "'" 
					
					
					
						Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
						rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

						if not rsCodeDesc.EOF then
							tempCodeDescription = rsCodeDesc ("CodeDescription")
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = rsCodeDesc ("reference")	
							tempRounding = rsCodeDesc ("rounding")
						else
							tempCodeDescription = ""					
							tempRowSpan = 1
							tempReference = ""						
							tempRounding = 0
						end if		
					    tempUOM = rsSortasi("UnitOfMeasurement")			
						rsCodeDesc.Close
						set rsCodeDesc = Nothing
										
										
					%>
					
					<% if no = 1 then %>
						<% if sPrevDesc <> tempCodeHeadDescription then %>				
							<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeHeadDescription%></td>
							<% sPrevDesc = tempCodeHeadDescription %>
						<% else %>
							<td align="left" ></td>
						<% end if %>
					<% end if %>
									
					<!--td align="left"  ><%= no %></td-->					
					<td align="left"  ></td>					
					
				
				
					
					
					<td align="left" ><a href="<%=tempReference%>"><%=tempCodeDescription%></a></td>
					<td align="center" ><%=tempUOM%></td>				
					
					<% ' Start Compare to Budget 

						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	

						if tempDataBudget_HI <> 0 then 
							tempCompare_HI = cdbl(rsSortasi("hi"))/tempDataBudget_HI * 100
						else
							tempCompare_HI = 0
						end if
						
						if tempDataBudget_SHI <> 0 then 
							tempCompare_SHI = cdbl(rsSortasi("shi"))/tempDataBudget_SHI * 100
						else
							tempCompare_SHI = 0
						end if

						if tempDataBudget_SBI <> 0 then 
							tempCompare_SBI = cdbl(rsSortasi("sbi"))/tempDataBudget_SBI * 100
						else
							tempCompare_SBI = 0
						end if
						
					%>					
					
					
					<%  
					    if rsSortasi("HI") <> "" then 
					        tempHI = cdbl(rsSortasi("HI"))
							if rsSortasi("DescCode") = "FFB07" then 
							    tempGrading_HI = tempHI
							end if 
					    else
							tempHI = 0
						end if
					
					    if rsSortasi("SHI") <> "" then 
					        tempSHI = cdbl(rsSortasi("SHI"))	
							if rsSortasi("DescCode") = "FFB07" then 
							    tempGrading_SHI = tempSHI
							end if 							
					    else
							tempSHI = 0
						end if

					    if rsSortasi("SBI") <> "" then 
					        tempSBI = cdbl(rsSortasi("SBI"))		
							if rsSortasi("DescCode") = "FFB07" then 
							    tempGrading_SBI = tempSBI
							end if							
					    else
							tempSBI = 0
						end if
					%>
					
								


					
					<td align="center" ><%=FormatNumber(tempSHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempSBI,tempRounding)%></td>
					
				
					<td align="center" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>

					<%
						if tempCompare_HI > 95 then
							cClassAssigned_HI = "green"
						else
							if tempCompare_HI > 85 then 
								 cClassAssigned_HI = "yellow"
							else
								 cClassAssigned_HI = "red"
							end if
						end if
						
						if tempCompare_SHI > 95 then
							cClassAssigned_SHI = "green"
						else
							if tempCompare_SHI > 85 then 
								 cClassAssigned_SHI = "yellow"
							else
								 cClassAssigned_SHI = "red"
							end if
						end if
						
						if tempCompare_SBI > 95 then
							cClassAssigned_SBI = "green"
						else
							if tempCompare_SBI > 85 then 
								 cClassAssigned_SBI = "yellow"
							else
								 cClassAssigned_SBI = "red"
							end if
						end if
						
					%>
					
					
					<td class="<%=cClassAssigned_SHI%>" align="center" ><%=FormatNumber(tempCompare_SHI,tempRounding)%></td>
					<td class="<%=cClassAssigned_SBI%>" align="center" ><%=FormatNumber(tempCompare_SBI,tempRounding)%></td>
					
				</tr>
	<%

				rsSortasi.MoveNext
			Wend
		end if
		rsSortasi.Close
		set rsSortasi = Nothing
	%>
	
	
<!-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX -->	

	<tr>
		<th class="separator" colspan="10" align="left"></th>				
	</tr>	


	
	<!-- Start PROSES -->	

	<%

		if not rsProses.EOF then
			no = 0
			while not rsProses.EOF
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				no = no + 1

	%>				


				<tr class="<%=cClass%>">
				
					<%
					    strCodeDescHead = " select fcname from parameter " & _
										" where " & _
										" parcode = 'SR005' and " & _
										" fccode = '" & rsProses("theHeadGroup") & "' and " & _
										" fcba = '" & v_txtFcba & "'"
										
						Set rsCodeDescHead = Server.CreateObject("ADODB.Recordset")	
						rsCodeDescHead.Open strCodeDescHead, connstr, adOpenStatic 		

						if not rsCodeDescHead.EOF then
							tempCodeHeadDescription = rsCodeDescHead ("fcname")
						else
							tempCodeHeadDescription = ""
						end if		
						rsCodeDescHead.Close
						set rsCodeDescHead = Nothing
					
					
						strCodeDesc = " select fcname codeDescription, rate_2 theRowSpan, rate_3 rounding, param_3 decrease_hi, param_4 decrease_shi, param_5 decrease_sbi , param_6 reference from parameterdetail " & _
									  "	where " & _
									    " parcode = 'SR005'  and " & _
										" parheadcode = 'TBS_PROSES' and " & _
										" param_2 = 'DEFAULT' and " & _
										" fccode = '" & rsProses("DescCode") & "'" 
					
					
					
						Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
						rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

						if not rsCodeDesc.EOF then
							tempCodeDescription = rsCodeDesc ("CodeDescription")
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = rsCodeDesc ("reference")	
							tempDecrease_hi = rsCodeDesc ("decrease_hi")
							tempDecrease_shi = rsCodeDesc ("decrease_shi")
							tempDecrease_sbi = rsCodeDesc ("decrease_sbi")	
							tempRounding = rsCodeDesc ("rounding")	
						else
							tempCodeDescription = ""					
							tempRowSpan = 1
							tempReference = ""	
							tempDecrease_hi = ""
							tempDecrease_shi = ""
							tempDecrease_sbi = ""								
							tempRounding = 0	
						end if		
					    tempUOM = rsProses("UnitOfMeasurement")			
						rsCodeDesc.Close
						set rsCodeDesc = Nothing
						
										
					%>
					
					<% if no = 1 then %>
						<% if sPrevDesc <> tempCodeHeadDescription then %>				
							<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeHeadDescription%></td>
							<% sPrevDesc = tempCodeHeadDescription %>
						<% else %>
							<td align="left" ></td>
						<% end if %>
					<% end if %>
									
					<!--td align="left"  ><%= no %></td-->					
					<td align="left"  ></td>							
				
				
					
					
					<td align="left" ><a href="<%=tempReference%>"><%=tempCodeDescription%></a></td>
					<td align="center" ><%=tempUOM%></td>				
					
					<% ' Start Compare to Budget 

						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	

						if tempDataBudget_HI <> 0 then 
							tempCompare_HI = cdbl(rsProses("hi"))/tempDataBudget_HI * 100
						else
							tempCompare_HI = 0
						end if
						
						if tempDataBudget_SHI <> 0 then 
							tempCompare_SHI = cdbl(rsProses("shi"))/tempDataBudget_SHI * 100
						else
							tempCompare_SHI = 0
						end if

						if tempDataBudget_SBI <> 0 then 
							tempCompare_SBI = cdbl(rsProses("sbi"))/tempDataBudget_SBI * 100
						else
							tempCompare_SBI = 0
						end if
						
					%>					
					
					
					<%  
					    if rsProses("HI") <> "" then 
					        tempHI = cdbl(rsProses("HI"))
							if tempDecrease_hi = "DECREASE_GRADING" then 
								tempHI = tempHI - tempGrading_HI
							end if
					    else
							tempHI = 0
						end if
					
					    if rsProses("SHI") <> "" then 
					        tempSHI = cdbl(rsProses("SHI"))		
							if tempDecrease_Shi = "DECREASE_GRADING" then 
								tempSHI = tempSHI - tempGrading_SHI
							end if							
					    else
							tempSHI = 0
						end if

					    if rsProses("SBI") <> "" then 
					        tempSBI = cdbl(rsProses("SBI"))		
							if tempDecrease_Sbi = "DECREASE_GRADING" then 
								tempSBI = tempSBI - tempGrading_SBI
							end if								
					    else
							tempSBI = 0
						end if
					%>


					
					<td align="center" ><%=FormatNumber(tempSHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempSBI,tempRounding)%></td>
					
					<%
						if rsProses("DescCode") = "FFB0701" then 
						    tempData_Proses_HI = tempHI
							tempData_Proses_SHI = tempSHI
							tempData_Proses_SBI = tempSBI
					    end if
					
					%>
					
					
					
					
					<td align="center" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>

					<%
						if tempCompare_HI > 95 then
							cClassAssigned_HI = "green"
						else
							if tempCompare_HI > 85 then 
								 cClassAssigned_HI = "yellow"
							else
								 cClassAssigned_HI = "red"
							end if
						end if
						
						if tempCompare_SHI > 95 then
							cClassAssigned_SHI = "green"
						else
							if tempCompare_SHI > 85 then 
								 cClassAssigned_SHI = "yellow"
							else
								 cClassAssigned_SHI = "red"
							end if
						end if
						
						if tempCompare_SBI > 95 then
							cClassAssigned_SBI = "green"
						else
							if tempCompare_SBI > 85 then 
								 cClassAssigned_SBI = "yellow"
							else
								 cClassAssigned_SBI = "red"
							end if
						end if
						
					%>
					
					
					<td class="<%=cClassAssigned_SHI%>" align="center" ><%=FormatNumber(tempCompare_SHI,tempRounding)%></td>
					<td class="<%=cClassAssigned_SBI%>" align="center" ><%=FormatNumber(tempCompare_SBI,tempRounding)%></td>
					
				</tr>
	<%

				rsProses.MoveNext
			Wend
		end if
		rsProses.Close
		set rsProses = Nothing
	%>
	
	
<!-- XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX -->	

	<tr>
		<th class="separator" colspan="10" align="left"></th>				
	</tr>	


	
	<!-- Start KAPASITAS OLAH -->		
	
	

	<%

		if not rsKapasitasOlah.EOF then
			no = 0
			while not rsKapasitasOlah.EOF
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				no = no + 1

	%>				


				<tr class="<%=cClass%>">
				
					<%
					    strCodeDescHead = " select fcname from parameter " & _
										" where " & _
										" parcode = 'SR005' and " & _
										" fccode = '" & rsKapasitasOlah("theHeadGroup") & "' and " & _
										" fcba = '" & v_txtFcba & "'"
										
						Set rsCodeDescHead = Server.CreateObject("ADODB.Recordset")	
						rsCodeDescHead.Open strCodeDescHead, connstr, adOpenStatic 		

						if not rsCodeDescHead.EOF then
							tempCodeHeadDescription = rsCodeDescHead ("fcname")
						else
							tempCodeHeadDescription = ""
						end if		
						rsCodeDescHead.Close
						set rsCodeDescHead = Nothing
					
					
						strCodeDesc = " select fcname codeDescription, rate_2 theRowSpan, rate_3 rounding, param_6 reference from parameterdetail " & _
									  "	where " & _
									    " parcode = 'SR005'  and " & _
										" parheadcode = 'KAPASITAS_OLAH' and " & _
										" param_2 = 'DEFAULT' and " & _
										" fccode = '" & rsKapasitasOlah("DescCode") & "'" 
					
					
					
						Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
						rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

						if not rsCodeDesc.EOF then
							tempCodeDescription = rsCodeDesc ("CodeDescription")
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = rsCodeDesc ("reference")	
							tempRounding = rsCodeDesc ("rounding")
						else
							tempCodeDescription = ""					
							tempRowSpan = 1
							tempReference = ""	
							tempRounding = 0
						end if		
					    tempUOM = rsKapasitasOlah("UnitOfMeasurement")			
						rsCodeDesc.Close
						set rsCodeDesc = Nothing
						
										
					%>
					
					<% if no = 1 then %>
						<% if sPrevDesc <> tempCodeHeadDescription then %>				
							<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeHeadDescription%></td>
							<% sPrevDesc = tempCodeHeadDescription %>
						<% else %>
							<td align="left" ></td>
						<% end if %>
					<% end if %>
									
					<!--td align="left"  ><%= no %></td-->					
					<td align="left"  ></td>		
					
				
				
					
					
					<td align="left" ><a href="<%=tempReference%>"><%=tempCodeDescription%></a></td>
					<td align="center" ><%=tempUOM%></td>				
					
					<% ' Start Compare to Budget 

						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	

						if tempDataBudget_HI <> 0 then 
							tempCompare_HI = cdbl(rsKapasitasOlah("hi"))/tempDataBudget_HI * 100
						else
							tempCompare_HI = 0
						end if
						
						if tempDataBudget_SHI <> 0 then 
							tempCompare_SHI = cdbl(rsKapasitasOlah("shi"))/tempDataBudget_SHI * 100
						else
							tempCompare_SHI = 0
						end if

						if tempDataBudget_SBI <> 0 then 
							tempCompare_SBI = cdbl(rsKapasitasOlah("sbi"))/tempDataBudget_SBI * 100
						else
							tempCompare_SBI = 0
						end if
						
					%>					
					
					
					<%  
					    if rsKapasitasOlah("HI") <> ""  then
                            if  cdbl(rsKapasitasOlah("HI")) <> 0 then 						
								tempHI = tempData_Proses_HI / cdbl(rsKapasitasOlah("HI")) / 1000
							else
							    tempHI = 0
							end if
					    else
							tempHI = 0
						end if
					
					    if rsKapasitasOlah("SHI") then 
						    if  cdbl(rsKapasitasOlah("SHI")) <> 0 then 
								tempSHI = tempData_Proses_SHI / cdbl(rsKapasitasOlah("SHI")) / 1000						
							else
							    tempSHI = 0
							end if
					    else
							tempSHI = 0
						end if

					    if rsKapasitasOlah("SBI") <> "" then 
						    if cdbl(rsKapasitasOlah("SBI")) <> 0 then 
					            tempSBI = tempData_Proses_SBI / cdbl(rsKapasitasOlah("SBI")) / 1000									
							else
							    tempSBI = 0
							end if
					    else
							tempSBI = 0
						end if
					%>


					
					<td align="center" ><%=FormatNumber(tempSHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempSBI,tempRounding)%></td>
					
					
					<td align="center" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>

					<%
						if tempCompare_HI > 95 then
							cClassAssigned_HI = "green"
						else
							if tempCompare_HI > 85 then 
								 cClassAssigned_HI = "yellow"
							else
								 cClassAssigned_HI = "red"
							end if
						end if
						
						if tempCompare_SHI > 95 then
							cClassAssigned_SHI = "green"
						else
							if tempCompare_SHI > 85 then 
								 cClassAssigned_SHI = "yellow"
							else
								 cClassAssigned_SHI = "red"
							end if
						end if
						
						if tempCompare_SBI > 95 then
							cClassAssigned_SBI = "green"
						else
							if tempCompare_SBI > 85 then 
								 cClassAssigned_SBI = "yellow"
							else
								 cClassAssigned_SBI = "red"
							end if
						end if
						
					%>
					
					
					<td class="<%=cClassAssigned_SHI%>" align="center" ><%=FormatNumber(tempCompare_SHI,tempRounding)%></td>
					<td class="<%=cClassAssigned_SBI%>" align="center" ><%=FormatNumber(tempCompare_SBI,tempRounding)%></td>
					
				</tr>
	<%

				rsKapasitasOlah.MoveNext
			Wend
		end if
		rsKapasitasOlah.Close
		set rsKapasitasOlah = Nothing
	%>	
	
	
	
	<tr>
		<th class="separator" colspan="13" align="left"></th>				
	</tr>	


	

	<%

		if not rsSisaTBS.EOF then
			no = 0
			while not rsSisaTBS.EOF
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				no = no + 1

	%>				


				<tr class="<%=cClass%>">
				
					<%
					    strCodeDescHead = " select fcname from parameter " & _
										" where " & _
										" parcode = 'SR005' and " & _
										" fccode = '" & rsSisaTBS("theHeadGroup") & "' and " & _
										" fcba = '" & v_txtFcba & "'"
										
						Set rsCodeDescHead = Server.CreateObject("ADODB.Recordset")	
						rsCodeDescHead.Open strCodeDescHead, connstr, adOpenStatic 		

						if not rsCodeDescHead.EOF then
							tempCodeHeadDescription = rsCodeDescHead ("fcname")
						else
							tempCodeHeadDescription = ""
						end if		
						rsCodeDescHead.Close
						set rsCodeDescHead = Nothing
					
					
						strCodeDesc = " select fcname codeDescription, rate_2 theRowSpan, rate_3 rounding, param_6 reference from parameterdetail " & _
									  "	where " & _
									    " parcode = 'SR005'  and " & _
										" parheadcode = 'SISA_TBS' and " & _
										" param_2 = 'DEFAULT' and " & _
										" fccode = '" & rsSisaTBS("DescCode") & "'" 
					
					
					
						Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
						rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

						if not rsCodeDesc.EOF then
							tempCodeDescription = rsCodeDesc ("CodeDescription")
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = rsCodeDesc ("reference")	
							tempRounding = rsCodeDesc ("rounding")
						else
							tempCodeDescription = ""					
							tempRowSpan = 1
							tempReference = ""		
						    tempRounding = 0							
						end if		
					    tempUOM = rsSisaTBS("UnitOfMeasurement")			
						rsCodeDesc.Close
						set rsCodeDesc = Nothing
						
										
					%>
					
					<% if no = 1 then %>
						<% if sPrevDesc <> tempCodeHeadDescription then %>				
							<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeHeadDescription%></td>
							<% sPrevDesc = tempCodeHeadDescription %>
						<% else %>
							<td align="left" ></td>
						<% end if %>
					<% end if %>
									
					<td align="center"  ><%= no %></td>					
					
				
				
					
					
					<td align="left" ><a href="<%=tempReference%>"><%=tempCodeDescription%></a></td>
					<td align="center" ><%=tempUOM%></td>				
					
					<% ' Start Compare to Budget 

						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	

						if tempDataBudget_HI <> 0 then 
							tempCompare_HI = cdbl(rsSisaTBS("hi"))/tempDataBudget_HI * 100
						else
							tempCompare_HI = 0
						end if
						
						if tempDataBudget_SHI <> 0 then 
							tempCompare_SHI = cdbl(rsSisaTBS("shi"))/tempDataBudget_SHI * 100
						else
							tempCompare_SHI = 0
						end if

						if tempDataBudget_SBI <> 0 then 
							tempCompare_SBI = cdbl(rsSisaTBS("sbi"))/tempDataBudget_SBI * 100
						else
							tempCompare_SBI = 0
						end if
						
					%>					
					
					
					<%  
					    if rsSisaTBS("HI") <> "" then 
					        tempHI = cdbl(rsSisaTBS("HI")) 
					    else
							tempHI = 0
						end if
					
					    if rsSisaTBS("SHI") <> "" then 
					        tempSHI = cdbl(rsSisaTBS("SHI")) 						
					    else
							tempSHI = 0
						end if

					    if rsSisaTBS("SBI") <> "" then 
					        tempSBI = cdbl(rsSisaTBS("SBI")) 									
					    else
							tempSBI = 0
						end if
					%>
					
					<%
						if rsSisaTBS("DescCode") = "FFB01" then 
                           tempOpeningStock = tempHI
					    end if
					%>


					<td align="center" ><%=FormatNumber(tempHI,tempRounding)%></td>
					<td class="grey" align="right" ></td>
					<td class="grey" align="right" ></td>					

					<td class="grey" align="right" ></td>
					<td class="grey" align="right" ></td>
					<td class="grey" align="right" ></td>					
					
					<%
						if tempCompare_HI > 95 then
							cClassAssigned_HI = "green"
						else
							if tempCompare_HI > 85 then 
								 cClassAssigned_HI = "yellow"
							else
								 cClassAssigned_HI = "red"
							end if
						end if
						
						if tempCompare_SHI > 95 then
							cClassAssigned_SHI = "green"
						else
							if tempCompare_SHI > 85 then 
								 cClassAssigned_SHI = "yellow"
							else
								 cClassAssigned_SHI = "red"
							end if
						end if
						
						if tempCompare_SBI > 95 then
							cClassAssigned_SBI = "green"
						else
							if tempCompare_SBI > 85 then 
								 cClassAssigned_SBI = "yellow"
							else
								 cClassAssigned_SBI = "red"
							end if
						end if
						
					%>
					
					
					
				</tr>
	<%

				rsSisaTBS.MoveNext
			Wend
		end if
		rsSisaTBS.Close
		set rsSisaTBS = Nothing
	%>	
	
	
	<tr>
		<td align="center" >2</td>
		<td align="left" >Bulan ini</td>
		<td align="center" ><%=tempUOM%></td>

		<% 
		
		   tempProcessed = tempData_Proses_HI
		   if (tempOpeningStock + tempReceived - tempProcessed - tempGrading_HI) > 0 then 
		       theClass = "black"
		   end if %>
		
		<td class="<%=theClass%>" align="center" ><%=FormatNumber((tempOpeningStock + tempReceived - tempProcessed),tempRounding) %></td>
		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>
		
		
		
	</tr>	
	
	
	
	<tr>
		<th class="separator" colspan="10" align="left"></th>				
	</tr>	

	
	

	<%

		if not rsOilKernelProduction.EOF then
			no = 0
			while not rsOilKernelProduction.EOF
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				no = no + 1

	%>				


				<tr class="<%=cClass%>">
				
					<%
					    strCodeDescHead = " select fcname from parameter " & _
										" where " & _
										" parcode = 'SR005' and " & _
										" fccode = '" & rsOilKernelProduction("theHeadGroup") & "' and " & _
										" fcba = '" & v_txtFcba & "'"
										
						Set rsCodeDescHead = Server.CreateObject("ADODB.Recordset")	
						rsCodeDescHead.Open strCodeDescHead, connstr, adOpenStatic 		

						if not rsCodeDescHead.EOF then
							tempCodeHeadDescription = rsCodeDescHead ("fcname")
						else
							tempCodeHeadDescription = ""
						end if		
						rsCodeDescHead.Close
						set rsCodeDescHead = Nothing
					
					
						strCodeDesc = " select fcname codeDescription, rate_2 theRowSpan, rate_3 rounding, param_6 reference from parameterdetail " & _
									  "	where " & _
									    " parcode = 'SR005'  and " & _
										" parheadcode = 'PRODUCTION' and " & _
										" param_2 = 'DEFAULT' and " & _
										" fccode = '" & rsOilKernelProduction("DescCode") & "'" 
					
					
					
						Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
						rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

						if not rsCodeDesc.EOF then
							tempCodeDescription = rsCodeDesc ("CodeDescription")
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = rsCodeDesc ("reference")	
							tempRounding = rsCodeDesc ("rounding")								
						else
							tempCodeDescription = ""					
							tempRowSpan = 1
							tempReference = ""		
							tempRounding = 0
						end if		
					    tempUOM = rsOilKernelProduction("UnitOfMeasurement")			
						rsCodeDesc.Close
						set rsCodeDesc = Nothing
						
										
					%>
					
					<% if no = 1 then %>
						<% if sPrevDesc <> tempCodeHeadDescription then %>				
							<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeHeadDescription%></td>
							<% sPrevDesc = tempCodeHeadDescription %>
						<% else %>
							<td align="left" ></td>
						<% end if %>
					<% end if %>
									
					<td align="center"  ><%= no %></td>					
					
				
				
					
					
					<td align="left" ><a href="<%=tempReference%>"><%=tempCodeDescription%></a></td>
					<td align="center" ><%=tempUOM%></td>				
					
					<% ' Start Compare to Budget 

						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	

						if tempDataBudget_HI <> 0 then 
							tempCompare_HI = cdbl(rsOilKernelProduction("hi"))/tempDataBudget_HI * 100
						else
							tempCompare_HI = 0
						end if
						
						if tempDataBudget_SHI <> 0 then 
							tempCompare_SHI = cdbl(rsOilKernelProduction("shi"))/tempDataBudget_SHI * 100
						else
							tempCompare_SHI = 0
						end if

						if tempDataBudget_SBI <> 0 then 
							tempCompare_SBI = cdbl(rsOilKernelProduction("sbi"))/tempDataBudget_SBI * 100
						else
							tempCompare_SBI = 0
						end if
						
					%>					
					
					
					<%  
					    if rsOilKernelProduction("HI") <> "" then 
					        tempHI = cdbl(rsOilKernelProduction("HI")) 
					    else
							tempHI = 0
						end if
					
					    if rsOilKernelProduction("SHI") <> "" then 
					        tempSHI = cdbl(rsOilKernelProduction("SHI")) 						
					    else
							tempSHI = 0
						end if

					    if rsOilKernelProduction("SBI") <> "" then 
					        tempSBI = cdbl(rsOilKernelProduction("SBI")) 									
					    else
							tempSBI = 0
						end if
					%>
					
					<%
						if rsOilKernelProduction("DescCode") = "FFB01" then 
                           tempOpeningStock = tempHI
					    end if
					%>


					
					<td align="center" ><%=FormatNumber(tempSHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempSBI,tempRounding)%></td>
					
					<%
				    if rsOilKernelProduction("DescCode") = "PROD01" then 
					   tempOilProcessed = tempHI
					   tempOilProcessed_SHI = tempSHI
					   tempOilProcessed_SBI = tempSBI
				    end if		

				    if rsOilKernelProduction("DescCode") = "PROD02" then 
					   tempKernelProcessed = tempHI
					   tempKernelProcessed_SHI = tempSHI
					   tempKernelProcessed_SBI = tempSBI
				    end if		
					%>
					
				
					<td align="center" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
					<td align="center" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>
					
					
					<%
						if tempCompare_HI > 95 then
							cClassAssigned_HI = "green"
						else
							if tempCompare_HI > 85 then 
								 cClassAssigned_HI = "yellow"
							else
								 cClassAssigned_HI = "red"
							end if
						end if
						
						if tempCompare_SHI > 95 then
							cClassAssigned_SHI = "green"
						else
							if tempCompare_SHI > 85 then 
								 cClassAssigned_SHI = "yellow"
							else
								 cClassAssigned_SHI = "red"
							end if
						end if
						
						if tempCompare_SBI > 95 then
							cClassAssigned_SBI = "green"
						else
							if tempCompare_SBI > 85 then 
								 cClassAssigned_SBI = "yellow"
							else
								 cClassAssigned_SBI = "red"
							end if
						end if
						
					%>
					
					
					<td class="<%=cClassAssigned_SHI%>" align="center" ><%=FormatNumber(tempCompare_SHI,tempRounding)%></td>
					<td class="<%=cClassAssigned_SBI%>" align="center" ><%=FormatNumber(tempCompare_SBI,tempRounding)%></td>
					
					
				</tr>
	<%

				rsOilKernelProduction.MoveNext
			Wend
		end if
		rsOilKernelProduction.Close
		set rsOilKernelProduction = Nothing
	%>	
	
	
	
	
	<tr>
		<th class="separator" colspan="10" align="left"></th>				
	</tr>	

		
	
	
	
	<!-- RENDEMEN -->	
	
	
	<%
	    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
		' Process for Oil After Grading	
		no = 0
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass = "default"
		end if
		no = no + 1
		tempDescription = "OER"
		tempUoM = "%"
		
		if tempData_Proses_HI <> 0 then 
			tempHI = tempOilProcessed / tempData_Proses_HI  * 100
			tempSHI = tempOilProcessed_SHI / tempData_Proses_SHI * 100
			tempSBI = tempOilProcessed_SBI / tempData_Proses_SBI * 100
		else
			tempHI = 0
			tempSHI = 0
			tempSBI = 0
		end if
	%>
	<tr class="<%=cClass%>">
	    <td class="boldfont" rowspan="2" align="center" >RENDEMEN</td>

		<td align="center" ><%=no%></td>
		<td align="left" ><%=tempDescription%></td>
		<td align="center" ><%=tempUOM%></td>
		
		
		<td align="center" ><%=FormatNumber(tempSHI,1)%></td>
		<td align="center" ><%=FormatNumber(tempSBI,1)%></td>	
		
		
		<td align="center" ><%=FormatNumber(0,2)%></td>
		<td align="center" ><%=FormatNumber(0,2)%></td>	
		
		
		<td align="center" ><%=FormatNumber(0,2)%></td>
		<td align="center" ><%=FormatNumber(0,2)%></td>		
	</tr>
	
	<%
	    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
		' Process for Kernel After Grading	
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass = "default"
		end if
		no = no + 1
		tempDescription = "KER"
		tempUoM = "%"
		
		if tempData_Proses_HI <> 0 then 
			tempHI = tempKernelProcessed / tempData_Proses_HI  * 100
			tempSHI = tempKernelProcessed_SHI / tempData_Proses_SHI * 100
			tempSBI = tempKernelProcessed_SBI / tempData_Proses_SBI * 100
		else
			tempHI = 0
			tempSHI = 0
			tempSBI = 0
		end if
	%>

	<tr class="<%=cClass%>">

		<td align="center" ><%=no%></td>
		<td align="left" ><%=tempDescription%></td>
		<td align="center" ><%=tempUOM%></td>
		
		
		<td align="center" ><%=FormatNumber(tempSHI,1)%></td>
		<td align="center" ><%=FormatNumber(tempSBI,1)%></td>	
		
		
		<td align="center" ><%=FormatNumber(0,2)%></td>
		<td align="center" ><%=FormatNumber(0,2)%></td>	
		
		
		<td align="center" ><%=FormatNumber(0,2)%></td>
		<td align="center" ><%=FormatNumber(0,2)%></td>			
	</tr>
	
	
	
	<tr>
		<th class="separator" colspan="10" align="left"></th>				
	</tr>	


	
	
	
	
<%

		if not rsCPOStock.EOF then
			no = 0
			while not rsCPOStock.EOF
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				no = no + 1

	%>				


				<tr class="<%=cClass%>">
				
					<%
					    strCodeDescHead = " select fcname from parameter " & _
										" where " & _
										" parcode = 'SR005' and " & _
										" fccode = '" & rsCPOStock("theHeadGroup") & "' and " & _
										" fcba = '" & v_txtFcba & "'"
										
						Set rsCodeDescHead = Server.CreateObject("ADODB.Recordset")	
						rsCodeDescHead.Open strCodeDescHead, connstr, adOpenStatic 		

						if not rsCodeDescHead.EOF then
							tempCodeHeadDescription = rsCodeDescHead ("fcname")
						else
							tempCodeHeadDescription = ""
						end if		
						rsCodeDescHead.Close
						set rsCodeDescHead = Nothing
					
					
						strCodeDesc = " select fcname codeDescription, rate_2 theRowSpan,  rate_3 rounding, param_6 reference from parameterdetail " & _
									  "	where " & _
									    " parcode = 'SR005'  and " & _
										" parheadcode = 'CPOSTOCK' and " & _
										" param_2 = 'DEFAULT' and " & _
										" fccode = '" & rsCPOStock("DescCode") & "'" 
					
					
					
						Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
						rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

						if not rsCodeDesc.EOF then
							tempCodeDescription = rsCodeDesc ("CodeDescription")
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = rsCodeDesc ("reference")		
							tempRounding =  rsCodeDesc ("rounding")
						else
							tempCodeDescription = ""					
							tempRowSpan = 1
							tempReference = ""		
							tempRounding = 0
						end if		
					    tempUOM = rsCPOStock("UnitOfMeasurement")			
						rsCodeDesc.Close
						set rsCodeDesc = Nothing
						
										
					%>
					
					<% if no = 1 then %>
						<% if sPrevDesc <> tempCodeHeadDescription then %>				
							<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeHeadDescription%></td>
							<% sPrevDesc = tempCodeHeadDescription %>
						<% else %>
							<td align="left" ></td>
						<% end if %>
					<% end if %>
									
					<td align="center"  ><%= no %></td>					
					
				
				
					
					
					<td align="left" ><a href="<%=tempReference%>"><%=tempCodeDescription%></a></td>
					<td align="center" ><%=tempUOM%></td>				
					
					<% ' Start Compare to Budget 

						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	

						if tempDataBudget_HI <> 0 then 
							tempCompare_HI = cdbl(rsCPOStock("hi"))/tempDataBudget_HI * 100
						else
							tempCompare_HI = 0
						end if
						
						if tempDataBudget_SHI <> 0 then 
							tempCompare_SHI = cdbl(rsCPOStock("shi"))/tempDataBudget_SHI * 100
						else
							tempCompare_SHI = 0
						end if

						if tempDataBudget_SBI <> 0 then 
							tempCompare_SBI = cdbl(rsCPOStock("sbi"))/tempDataBudget_SBI * 100
						else
							tempCompare_SBI = 0
						end if
						
					%>					
					
					
					<%  
					    if rsCPOStock("HI") <> "" then 
					        tempHI = cdbl(rsCPOStock("HI")) 
					    else
							tempHI = 0
						end if
					
					    if rsCPOStock("SHI") <> "" then 
					        tempSHI = cdbl(rsCPOStock("SHI")) 						
					    else
							tempSHI = 0
						end if

					    if rsCPOStock("SBI") <> "" then 
					        tempSBI = cdbl(rsCPOStock("SBI")) 									
					    else
							tempSBI = 0
						end if
					%>
					


					
					
					<% 					   
					   if rsCPOStock("DescCode") = "CPO001" then 
						   tempOpeningCPOStock_HI = tempHI
					   end if		

					   if rsCPOStock("DescCode") = "CPO002" then 
						   tempDespatchCPOStock_HI = tempHI
					   end if					   
					   
					%>
					
					<% if rsCPOStock("DescCode") = "CPO001" then %> 
						<td class="grey" align="right" ></td>
						<td class="grey" align="right" ></td>
						
						<td class="grey" align="right" ></td>
						<td class="grey" align="right" ></td>
						<td class="grey" align="right" ></td>
						
						<td class="grey" align="right" ></td>
						
						
					<% else %>
						<td align="center" ><%=FormatNumber(tempSHI,tempRounding)%></td>
						<td align="center" ><%=FormatNumber(tempSBI,tempRounding)%></td>
						

						
						
						<td align="center" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
						<td align="center" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>
						
						
						<%
							if tempCompare_HI > 95 then
								cClassAssigned_HI = "green"
							else
								if tempCompare_HI > 85 then 
									 cClassAssigned_HI = "yellow"
								else
									 cClassAssigned_HI = "red"
								end if
							end if
							
							if tempCompare_SHI > 95 then
								cClassAssigned_SHI = "green"
							else
								if tempCompare_SHI > 85 then 
									 cClassAssigned_SHI = "yellow"
								else
									 cClassAssigned_SHI = "red"
								end if
							end if
							
							if tempCompare_SBI > 95 then
								cClassAssigned_SBI = "green"
							else
								if tempCompare_SBI > 85 then 
									 cClassAssigned_SBI = "yellow"
								else
									 cClassAssigned_SBI = "red"
								end if
							end if
							
						%>
						
						
						
						<td class="<%=cClassAssigned_SHI%>" align="center" ><%=FormatNumber(tempCompare_SHI,tempRounding)%></td>
						<td class="<%=cClassAssigned_SBI%>" align="center" ><%=FormatNumber(tempCompare_SBI,tempRounding)%></td>
					<% end if %>
					
				</tr>
	<%

				rsCPOStock.MoveNext
			Wend
		end if
		rsCPOStock.Close
		set rsCPOStock = Nothing
	%>	
	
    <%
	    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
		' Closing CPO Stock
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass = "default"
		end if
		no = no + 1
		tempDescription = "Saldo Akhir"
		tempUoM = "Kg"
		
		tempHI = tempOpeningCPOStock_HI - tempDespatchCPOStock_HI + tempOilProcessed
		tempSHI = 0
		tempSBI = 0	%>
		
	<tr class="<%=cClass%>">

		<td align="center" ><%=no%></td>
		<td align="left" ><%=tempDescription%></td>
		<td align="center" ><%=tempUOM%></td>
		

		<td class="<%=theClass%>" align="right" ><%=FormatNumber(tempHI,tempRounding)%></td>
		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>

		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>
		
		
			
		
	</tr>	
	
	
		

	<tr>
		<th class="separator" colspan="10" align="left"></th>				
	</tr>	


	
	
	
	
    <%

		if not rsKernelStock.EOF then
			no = 0
			while not rsKernelStock.EOF
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
				no = no + 1

	%>				


				<tr class="<%=cClass%>">
				
					<%
					    strCodeDescHead = " select fcname from parameter " & _
										" where " & _
										" parcode = 'SR005' and " & _
										" fccode = '" & rsKernelStock("theHeadGroup") & "' and " & _
										" fcba = '" & v_txtFcba & "'"
										
						Set rsCodeDescHead = Server.CreateObject("ADODB.Recordset")	
						rsCodeDescHead.Open strCodeDescHead, connstr, adOpenStatic 		

						if not rsCodeDescHead.EOF then
							tempCodeHeadDescription = rsCodeDescHead ("fcname")
						else
							tempCodeHeadDescription = ""
						end if		
						rsCodeDescHead.Close
						set rsCodeDescHead = Nothing
					
					
						strCodeDesc = " select fcname codeDescription, rate_2 theRowSpan,  rate_3 rounding, param_6 reference from parameterdetail " & _
									  "	where " & _
									    " parcode = 'SR005'  and " & _
										" parheadcode = 'PKSTOCK' and " & _
										" param_2 = 'DEFAULT' and " & _
										" fccode = '" & rsKernelStock("DescCode") & "'" 
					
					
					
						Set rsCodeDesc = Server.CreateObject("ADODB.Recordset")	
						rsCodeDesc.Open strCodeDesc, connstr, adOpenStatic 		

						if not rsCodeDesc.EOF then
							tempCodeDescription = rsCodeDesc ("CodeDescription")
							tempRowSpan = cdbl(rsCodeDesc ("theRowSpan"))
							tempReference = rsCodeDesc ("reference")
							tempRounding = rsCodeDesc ("rounding")
						else
							tempCodeDescription = ""					
							tempRowSpan = 1
							tempReference = ""								
							tempRounding = 0
						end if		
					    tempUOM = rsKernelStock("UnitOfMeasurement")			
						rsCodeDesc.Close
						set rsCodeDesc = Nothing
						
										
					%>
					
					<% if no = 1 then %>
						<% if sPrevDesc <> tempCodeHeadDescription then %>				
							<td class="boldfont" rowspan="<%= tempRowSpan %>" align="center" ><%=tempCodeHeadDescription%></td>
							<% sPrevDesc = tempCodeHeadDescription %>
						<% else %>
							<td align="left" ></td>
						<% end if %>
					<% end if %>
									
					<td align="center"  ><%= no %></td>					
					
				
				
					
					
					<td align="left" ><a href="<%=tempReference%>"><%=tempCodeDescription%></a></td>
					<td align="center" ><%=tempUOM%></td>				
					
					<% ' Start Compare to Budget 

						tempDataBudget_HI = 0
						tempDataBudget_SHI = 0
						tempDataBudget_SBI = 0	

						if tempDataBudget_HI <> 0 then 
							tempCompare_HI = cdbl(rsKernelStock("hi"))/tempDataBudget_HI * 100
						else
							tempCompare_HI = 0
						end if
						
						if tempDataBudget_SHI <> 0 then 
							tempCompare_SHI = cdbl(rsKernelStock("shi"))/tempDataBudget_SHI * 100
						else
							tempCompare_SHI = 0
						end if

						if tempDataBudget_SBI <> 0 then 
							tempCompare_SBI = cdbl(rsKernelStock("sbi"))/tempDataBudget_SBI * 100
						else
							tempCompare_SBI = 0
						end if
						
					%>					
					
					
					<%  
					    if rsKernelStock("HI") <> "" then 
					        tempHI = cdbl(rsKernelStock("HI")) 
					    else
							tempHI = 0
						end if
					
					    if rsKernelStock("SHI") <> "" then 
					        tempSHI = cdbl(rsKernelStock("SHI")) 						
					    else
							tempSHI = 0
						end if

					    if rsKernelStock("SBI") <> "" then 
					        tempSBI = cdbl(rsKernelStock("SBI")) 									
					    else
							tempSBI = 0
						end if
					%>
					
					<% 

					   
					   if rsKernelStock("DescCode") = "PK001" then 
                           tempOpeningKernelStock_HI = tempHI
					   end if		

					   if rsKernelStock("DescCode") = "PK002" then 
                           tempDespatchKernelStock_HI = tempHI
					   end if						   
					   
					%>	


					
					
					
					<% if rsKernelStock("DescCode") = "PK001" then %> 
						<td class="grey" align="right" ></td>
						<td class="grey" align="right" ></td>
						
						<td class="grey" align="right" ></td>
						<td class="grey" align="right" ></td>
						<td class="grey" align="right" ></td>
						<td class="grey" align="right" ></td>
					
						
					<% else %>

						<td align="center" ><%=FormatNumber(tempSHI,tempRounding)%></td>
						<td align="center" ><%=FormatNumber(tempSBI,tempRounding)%></td>
						

						
						
						<td align="center" ><%=FormatNumber(tempDataBudget_SHI,tempRounding)%></td>
						<td align="center" ><%=FormatNumber(tempDataBudget_SBI,tempRounding)%></td>
						
						
						<%
							if tempCompare_HI > 95 then
								cClassAssigned_HI = "green"
							else
								if tempCompare_HI > 85 then 
									 cClassAssigned_HI = "yellow"
								else
									 cClassAssigned_HI = "red"
								end if
							end if
							
							if tempCompare_SHI > 95 then
								cClassAssigned_SHI = "green"
							else
								if tempCompare_SHI > 85 then 
									 cClassAssigned_SHI = "yellow"
								else
									 cClassAssigned_SHI = "red"
								end if
							end if
							
							if tempCompare_SBI > 95 then
								cClassAssigned_SBI = "green"
							else
								if tempCompare_SBI > 85 then 
									 cClassAssigned_SBI = "yellow"
								else
									 cClassAssigned_SBI = "red"
								end if
							end if
							
						%>
						
						
						<td class="<%=cClassAssigned_SHI%>" align="center" ><%=FormatNumber(tempCompare_SHI,tempRounding)%></td>
						<td class="<%=cClassAssigned_SBI%>" align="center" ><%=FormatNumber(tempCompare_SBI,tempRounding)%></td>
					<% end if %>
					
					
				</tr>
	<%

				rsKernelStock.MoveNext
			Wend
		end if
		rsKernelStock.Close
		set rsKernelStock = Nothing
	%>	
		
    <%
	    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
		' Closing Kernel Stock
		if no  mod 2 = 1 then
			cClass = "odd"
		else
			cClass = "default"
		end if
		no = no + 1
		tempDescription = "Saldo Akhir"
		tempUoM = "Kg"
		
		tempHI = tempOpeningKernelStock_HI - tempDespatchKernelStock_HI + tempKernelProcessed
		tempSHI = 0
		tempSBI = 0	%>
	<tr class="<%=cClass%>">

		<td align="center" ><%=no%></td>
		<td align="left" ><%=tempDescription%></td>
		<td align="center" ><%=tempUOM%></td>
		

		<td class="<%=theClass%>" align="center" ><%=FormatNumber(tempHI,tempRounding)%></td>
		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>

		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>
		<td class="grey" align="right" ></td>
		
		
		
	</tr>		
</tbody>

</table>

</div>

</form>

</div>
</body>
</html>
