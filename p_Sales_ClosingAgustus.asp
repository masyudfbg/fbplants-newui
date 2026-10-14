<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="adovbs.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Sales Closing</title>
</head>
<body>
<%

Session("v_Sales") = cmdFindPeriod("Sales")
Session("cls_Sales") = cmdPeriodClose("Sales")
Session("CurP_Sales") = cmdFindAccountPeriod("Sales")
Session("ClsP_Sales") = cmdAccountPeriodClose("Sales")

Sub NormalisasiData
	arr_normalisasi_eof = true
	qPeriodControl = " select lineno from cost_activity " & _
				     " where ACCOUNTINGPERIOD = " & Mid(Session("CurP_Sales"),5,2) & "" & _
				     "       and ACCOUNTINGYEAR = " & Mid(Session("CurP_Sales"),1,4) & "" & _
				     "       and fcBA = '" & Session("fcBA") & "'" & _
				     "       and Sourcetype = 'SL'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
	  arr_normalisasi = rsPeriodControl.getrows()
	  arr_normalisasi_eof = false
	end if
	
	if not arr_normalisasi_eof then
		for xyz=0 to ubound(arr_normalisasi,2)
			qNormalisasiDetail = " DELETE FROM cost_activity_hist_tab WHERE lineheaderno = '"&arr_normalisasi(0,xyz)&"' AND fcba =  '" & Session("fcBA") & "'"
			connstr.execute(qNormalisasiDetail)
			
			qNormalisasiDetail = " DELETE FROM cost_activity WHERE lineno = '"&arr_normalisasi(0,xyz)&"' AND fcba =  '" & Session("fcBA") & "'"
			connstr.execute(qNormalisasiDetail) 
		next
	end if
	' Begin Normalisasi Insert Table HPP Penjualan Added By SAG 08 Juni 2019
		qNormalisasiHPP = " DELETE FROM hargapokokpenjualan WHERE month_qty_sales_sbi = '"& Mid(Session("CurP_Sales"),5,2) &"'" & _
							 " AND year_qty_sales_sbi = '"& Mid(Session("CurP_Sales"),1,4) &"'	AND fcba =  '" & Session("fcBA") & "'"
		' response.write qNormalisasiHPP
		connstr.execute(qNormalisasiHPP)
	' End Normalisasi Insert Table HPP Penjualan Added By SAG 08 Juni 2019
End Sub

Sub postingData
	' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX 
	' Add by NAF to select lastperiodposted
	Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl
    
	qPeriodControl="select add_months(LASTPERIODPOSTED,1) LastPeriodPosted, LASTPERIODPOSTED LastEndPeriodPosted from PERIODCONTROL " & _
	  " where SYSTEM='Sales' and fcba='" & Session("fcBA") & "'"
	
	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr
	
	if not rsPeriodControl.eof then
	   	LastPeriod = rsPeriodControl("LASTPERIODPOSTED")
	   IF  Cint(MID(Session("CurP_Sales"),5,2)) <= 12 Then
		   temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")     ' add Akbar 28 Des 2015 , for Account Period <= 12
	   ELSE
	       temp_datePeriodControl = rsPeriodControl("LASTENDPERIODPOSTED")  ' for Account Period > 12
       END IF
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Sales " & "<br>"
	end if	
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	' qControlJob="select ControlJob from Module where NAME='Purchase'"
	' Now we use new concept for picking Control Job, by using Process Control Tables
	' as per last concept are NOT able to pick data per Business Unit.
	
	tempProcessCode = "SALES_CLOSING"
	tempProcessNameCode = "p_Sales_Closing.asp"
	tempRefCodePVDetail = "PAYMENTVOUCHERDETAIL"

	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' Take the Original Table of PAYMENTVOUCHERDETAIL.
	qReference =  " select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
				  " from process_tableref_tab " & _
				  " where " & _
				  " processcode = '" & tempProcessCode & "' and " & _
				  " upper(processnamecode) =  upper('" & tempProcessNameCode & "') and " & _
				  " fccode =  '" & tempRefCodePVDetail & "'"

	Set rsReferencePVDetail = Server.CreateObject("ADODB.Recordset")
	rsReferencePVDetail.CursorType = 3
	rsReferencePVDetail.CursorLocation = 3
	rsReferencePVDetail.Open qReference, connstr

	if not rsReferencePVDetail.eof then
		txtRefNumberKeys_PVDetail = rsReferencePVDetail("REF_NUMBERKEYS")
		txtRefStringSeparator_PVDetail = rsReferencePVDetail("REF_STRINGSEPARATOR")
		txtRefStringKeys_PVDetail = rsReferencePVDetail("REF_STRINGKEYS")
	end if
	rsReferencePVDetail.close
	set rsReferencePVDetail = Nothing 
	
	qMaxLineNo = " select nvl(max(lineno),0) maxLineNo " & _
					" from cost_activity " & _
					" where " & _
					" fcba = '" & Session("fcBA") & "' "
					
	Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
	rsqMaxLineNo.CursorType = 3
	rsqMaxLineNo.CursorLocation = 3
	rsqMaxLineNo.Open qMaxLineNo, connstr

	if not rsqMaxLineNo.eof then
		maxLineNo = cdbl(to_nol(rsqMaxLineNo("maxLineNo")))
	else
		maxLineNo = 0
	end if
	rsqMaxLineNo.close
	set rsqMaxLineNo = Nothing
	
	
	' End of Modification

	' Add By IWA  24 Mar 2019
	
	if Session("Estate_Mill") = "E" then
	'-- Begin sales Shipment
	
	'-- Insert Debet
	qSalesDB =" SELECT" &_
				" ROWNUM LINENO,FCBA, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
				" TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " &_
				" TARGETTYPE, TARGETCODE, DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT," &_
				" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB,  " &_
				" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB " &_
				" FROM" &_
				" (" &_
				" SELECT" &_
				" MAX(ROWNUM) LINENO," &_
				" MAX(DTL.FCBA)FCBA," &_
				" MAX(DTL.ACCOUNTINGPERIOD)ACCOUNTINGPERIOD," &_
				" MAX(DTL.ACCOUNTINGYEAR)ACCOUNTINGYEAR," &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE," &_
				" MAX(IPLASPROD.PARAMETER.PARAM_2) JOBCODE," &_
				" 'SL' SOURCE," &_
				" 'SL' SOURCETYPE," &_
				" MAX(DTL.SALESSHIPMENTNO) SOURCECODE," &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE," &_
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE," &_
				" MAX(DTL.SALESCONTRACTNO) DOCUMENTNO," &_
				" 0 QUANTITY," &_
				" 0 OUTPUT," &_
				" SUM(DTL.AMOUNT_ADVANCE_SALES) AMOUNT," &_
				" 'TC - Sales Contract No :' TITLE_REFERENCE_01_DB, " &_
				" MAX(DTL.SALESCONTRACTNO) AUTO_REFERENCE_01_DB, " &_
				" 'Auto - Sales Shipment Code :' TITLE_REFERENCE_02_DB, " &_
				" MAX(DTL.SALESSHIPMENTNO) AUTO_REFERENCE_02_DB " &_
				" FROM " &_
				" IPLASPROD.SALES_VIEW DTL" &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES'" &_
				" INNER JOIN IPLASPROD.PARAMETER ON DTL.FCBA = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES'" &_
				" INNER JOIN IPLASPROD.JOB ON DTL.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_2 = IPLASPROD.JOB.FCCODE" &_
				" INNER JOIN IPLASPROD.GCMASTER ON DTL.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY" &_
				" WHERE" &_
				" DTL.ACCOUNTINGPERIOD = " & Mid(Session("CurP_Sales"),5,2) &_ 
				" AND DTL.ACCOUNTINGYEAR = " & Mid(Session("CurP_Sales"),1,4) & _
				" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
				" GROUP BY DTL.FCBA,DTL.SALESCONTRACTNO,DTL.CUSTOMERCODE, DTL.SALESSHIPMENTNO" &_
				" ) ORDER BY LINENO"
	
	'response.write qSalesDB
	Set rsSalesDB = Server.CreateObject("ADODB.Recordset")
	rsSalesDB.Open qSalesDB, connstr
	
	if not rsSalesDB.eof then
		Do While Not rsSalesDB.eof		  
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
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("TDATE") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("JOBCODE") & "','" & rsSalesDB("SOURCE") & "','" & rsSalesDB("SOURCETYPE") & "','" & rsSalesDB("SOURCECODE") & "'," & _
							 "'','" & rsSalesDB("TARGETTYPE") & "','" & rsSalesDB("TARGETCODE") & "'," & _
							 "'" &  rsSalesDB("AMOUNT") & "','" & rsSalesDB("DOCUMENTNO") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsSalesDB("TITLE_REFERENCE_01_DB") & "' ,'"  & rsSalesDB("AUTO_REFERENCE_01_DB") & "'," & _
							 " '" & rsSalesDB("TITLE_REFERENCE_02_DB") & "' ,'"  & rsSalesDB("AUTO_REFERENCE_02_DB") & "'," & _
							 " '','', " & _
							 " '','', " & _
							 " 'Line No:', '" & rsSalesDB("LINENO") & "','Db', " & _
							 " " & rsSalesDB("ACCOUNTINGPERIOD") & "," & rsSalesDB("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDB)
			 connStr.execute("Commit")

		rsSalesDB.MoveNext
	  loop
	  response.Write "Process Sales Shipment TBS Debet Complete.... !!<br>"
	  response.Flush
	end if
	'-- End Insert Debet
	
	'-- Insert Credit Inti
	qSalesCRInti =" SELECT" &_
				" ROWNUM LINENO,FCBA, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
				" TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " &_
				" TARGETTYPE, TARGETCODE, DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT, " &_
				" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR," &_
				" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR " &_
				" FROM " &_
				" ( " &_
				" SELECT " &_
				" MAX(ROWNUM) LINENO, " &_
				" MAX(DTL.FCBA)FCBA, " &_
				" MAX(DTL.ACCOUNTINGPERIOD)ACCOUNTINGPERIOD, " &_
				" MAX(DTL.ACCOUNTINGYEAR)ACCOUNTINGYEAR, " &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE, " &_
				" MAX(IPLASPROD.FIELD.FIELD_INFO_04) JOBCODE, " &_
				" 'SL' SOURCE, " &_
				" 'SL' SOURCETYPE, " &_
				" MAX(DTL.SALESSHIPMENTNO) SOURCECODE, " &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE, " &_
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE, " &_
				" MAX(DTL.SALESCONTRACTNO) DOCUMENTNO, " &_
				" 0 QUANTITY, " &_
				" 0 OUTPUT, " &_
				" -1*SUM(DTL.AMOUNT_ADVANCE_SALES) AMOUNT, " &_
				" 'TC - Sales Contract No :' TITLE_REFERENCE_01_CR,  " &_
				" MAX(DTL.SALESCONTRACTNO) AUTO_REFERENCE_01_CR,  " &_
				" 'Auto - Sales Shipment Code :' TITLE_REFERENCE_02_CR,  " &_
				" MAX(DTL.SALESSHIPMENTNO) AUTO_REFERENCE_02_CR  " &_
				" FROM  " &_
				" IPLASPROD.SALES_VIEW DTL " &_
				" INNER JOIN IPLASPROD.FIELD ON DTL.FIELDCODE = IPLASPROD.FIELD.FCCODE AND DTL.FCBA = IPLASPROD.FIELD.FCBA " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES' " &_
				" INNER JOIN IPLASPROD.JOB ON DTL.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.FIELD.FIELD_INFO_04 = IPLASPROD.JOB.FCCODE " &_
				" INNER JOIN IPLASPROD.GCMASTER ON DTL.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY " &_
				" WHERE " &_
				" DTL.ACCOUNTINGPERIOD = " & Mid(Session("CurP_Sales"),5,2) &_ 
				" AND DTL.ACCOUNTINGYEAR = " & Mid(Session("CurP_Sales"),1,4) & _
				" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
				" AND DTL.OWNERSHIP = 'INTI' " &_
				" GROUP BY DTL.FCBA,DTL.SALESCONTRACTNO,DTL.CUSTOMERCODE, DTL.SALESSHIPMENTNO " &_
				" ) ORDER BY LINENO"
	
	'response.write qSalesCR
	Set rsSalesCRInti = Server.CreateObject("ADODB.Recordset")
	rsSalesCRInti.Open qSalesCRInti, connstr
	
	if not rsSalesCRInti.eof then
		Do While Not rsSalesCRInti.eof	   
		
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityCr="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesCRInti("TDATE") & "','mm/dd/yyyy') " & ",'" & rsSalesCRInti("JOBCODE") & "','" & rsSalesCRInti("SOURCE") & "','" & rsSalesCRInti("SOURCETYPE") & "','" & rsSalesCRInti("SOURCECODE") & "'," & _
							 "'','" & rsSalesCRInti("TARGETTYPE") & "','" & rsSalesCRInti("TARGETCODE") & "'," & _
							 "'" &  rsSalesCRInti("AMOUNT") & "','" & rsSalesCRInti("DOCUMENTNO") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsSalesCRInti("TITLE_REFERENCE_01_CR") & "' ,'"  & rsSalesCRInti("AUTO_REFERENCE_01_CR") & "'," & _
							 " '" & rsSalesCRInti("TITLE_REFERENCE_02_CR") & "' ,'"  & rsSalesCRInti("AUTO_REFERENCE_02_CR") & "'," & _
							 " '','', " & _
							 " '','', " & _
							 " 'Line No:', '" & rsSalesCRInti("LINENO") & "','Cr', " & _
							 " " & rsSalesCRInti("ACCOUNTINGPERIOD") & "," & rsSalesCRInti("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityCr)
			 connStr.execute("Commit")

		  rsSalesCRInti.MoveNext
	  loop
	  	  response.Write "Process Sales Shipment TBS Credit (Inti) Complete.... !!<br>"
		  response.Flush
	end if
	'-- End Insert Credit Inti
	
	'-- Insert Credit Plasma
	qSalesCRPlasma =" SELECT" &_
				" ROWNUM LINENO,FCBA, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
				" TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, " &_
				" TARGETTYPE, TARGETCODE, DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT, " &_
				" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR," &_
				" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR " &_
				" FROM " &_
				" ( " &_
				" SELECT " &_
				" MAX(ROWNUM) LINENO, " &_
				" MAX(DTL.FCBA)FCBA, " &_
				" MAX(DTL.ACCOUNTINGPERIOD)ACCOUNTINGPERIOD, " &_
				" MAX(DTL.ACCOUNTINGYEAR)ACCOUNTINGYEAR, " &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE, " &_
				" MAX(IPLASPROD.FIELD.FIELD_INFO_04) JOBCODE, " &_
				" 'SL' SOURCE, " &_
				" 'SL' SOURCETYPE, " &_
				" MAX(DTL.SALESSHIPMENTNO) SOURCECODE, " &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE, " &_
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE, " &_
				" MAX(DTL.SALESCONTRACTNO) DOCUMENTNO, " &_
				" 0 QUANTITY, " &_
				" 0 OUTPUT, " &_
				" -1*SUM(DTL.AMOUNT_ADVANCE_SALES) AMOUNT, " &_
				" 'TC - Sales Contract No :' TITLE_REFERENCE_01_CR,  " &_
				" MAX(DTL.SALESCONTRACTNO) AUTO_REFERENCE_01_CR,  " &_
				" 'Auto - Sales Shipment Code :' TITLE_REFERENCE_02_CR,  " &_
				" MAX(DTL.SALESSHIPMENTNO) AUTO_REFERENCE_02_CR  " &_
				" FROM  " &_
				" IPLASPROD.SALES_VIEW DTL " &_
				" INNER JOIN IPLASPROD.FIELD ON DTL.FIELDCODE = IPLASPROD.FIELD.FCCODE AND DTL.FCBA = IPLASPROD.FIELD.FCBA " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES' " &_
				" INNER JOIN IPLASPROD.JOB ON DTL.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.FIELD.FIELD_INFO_04 = IPLASPROD.JOB.FCCODE " &_
				" INNER JOIN IPLASPROD.GCMASTER ON DTL.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY " &_
				" WHERE " &_
				" DTL.ACCOUNTINGPERIOD = " & Mid(Session("CurP_Sales"),5,2) &_ 
				" AND DTL.ACCOUNTINGYEAR = " & Mid(Session("CurP_Sales"),1,4) & _
				" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
				" AND DTL.OWNERSHIP = 'PLASMA' " &_
				" GROUP BY DTL.FCBA,DTL.SALESCONTRACTNO,DTL.CUSTOMERCODE, DTL.SALESSHIPMENTNO " &_
				" ) ORDER BY LINENO"
	
	'response.write qSalesCR
	Set rsSalesCRPlasma = Server.CreateObject("ADODB.Recordset")
	rsSalesCRPlasma.Open qSalesCRPlasma, connstr
	
	if not rsSalesCRPlasma.eof then
		Do While Not rsSalesCRPlasma.eof	   
		
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityCr="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesCRPlasma("TDATE") & "','mm/dd/yyyy') " & ",'" & rsSalesCRPlasma("JOBCODE") & "','" & rsSalesCRPlasma("SOURCE") & "','" & rsSalesCRPlasma("SOURCETYPE") & "','" & rsSalesCRPlasma("SOURCECODE") & "'," & _
							 "'','" & rsSalesCRPlasma("TARGETTYPE") & "','" & rsSalesCRPlasma("TARGETCODE") & "'," & _
							 "'" &  rsSalesCRPlasma("AMOUNT") & "','" & rsSalesCRPlasma("DOCUMENTNO") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsSalesCRPlasma("TITLE_REFERENCE_01_CR") & "' ,'"  & rsSalesCRPlasma("AUTO_REFERENCE_01_CR") & "'," & _
							 " '" & rsSalesCRPlasma("TITLE_REFERENCE_02_CR") & "' ,'"  & rsSalesCRPlasma("AUTO_REFERENCE_02_CR") & "'," & _
							 " '','', " & _
							 " '','', " & _
							 " 'Line No:', '" & rsSalesCRPlasma("LINENO") & "','Cr', " & _
							 " " & rsSalesCRPlasma("ACCOUNTINGPERIOD") & "," & rsSalesCRPlasma("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityCr)
			 connStr.execute("Commit")

		  rsSalesCRPlasma.MoveNext
	  loop
	  	  response.Write "Process Sales Shipment Credit TBS (Plasma) Complete.... !!<br>"
		  response.Flush
	end if
	'-- End Insert Credit Plasma
	
	'-- End sales Shipment
	end if
	
	if Session("Estate_Mill") = "M" then
	'-- Begin sales Shipment Part 2 Added By Sendy 27/05/2019
	
	'-- Insert Debet /Credit
	qSalesDB =	"   SELECT fcba," & _
				"          month_prod_bi," & _
				"          year_prod_bi," & _
				"          DebetCredit," & _
				" 'HARGAPOKOKPRODUKSI' HARGAPOKOKPRODUKSI, " & _
				"          jobcode," & _
				"          SUM (amount_prod_bi) AS amount_prod_bi" & _
				"     FROM (SELECT fcba," & _
				"                  product_code," & _
				"                  month_prod_bi," & _
				"                  year_prod_bi," & _
				"                  'Db' AS DebetCredit," & _
				"                  jobcode," & _
				"                  amount_prod_bi" & _
				"             FROM (SELECT fcba," & _
				"                          product_code," & _
				"                          month_prod_bi," & _
				"                          year_prod_bi," & _
				"                          jobcode," & _
				"                          qty_prod_bi," & _
				"                            NVL (" & _
				"                               IPLASPROD.GetLastMonthamountprodsbimid (" & _
				"                                  product_code," & _
				"                                  fcba," & _
				"                                  LAST_DAY (" & _
				"                                     ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -0)))," & _
				"                               0)" & _
				"                          - NVL (" & _
				"                               IPLASPROD.GetLastMonthamountprodsbimid (" & _
				"                                  product_code," & _
				"                                  fcba," & _
				"                                  LAST_DAY (" & _
				"                                     ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1)))," & _
				"                               0)" & _
				"                             AS amount_prod_bi" & _
				"                     FROM (SELECT STAGE1.fcba," & _
				"                                  STAGE1.product_code," & _
				"                                  STAGE1.month_prod_bi," & _
				"                                  STAGE1.year_prod_bi," & _
				"                                  STAGE2.param_1 AS jobcode," & _
				"                                  STAGE1.qty_cpo_prod_bi AS qty_prod_bi," & _
				"                                  STAGE1.prop_amount_cpo_prod_bi" & _
				"                                     AS amount_prod_bi," & _
				"                                  (  STAGE1.prop_amount_cpo_prod_bi" & _
				"                                   / STAGE1.qty_cpo_prod_bi)" & _
				"                                     AS Price_prod_bi" & _
				"                             FROM (SELECT fcba," & _
				"                                          'CPO' AS product_code," & _
				"                                          month_prod_bi," & _
				"                                          year_prod_bi," & _
				"                                          qty_cpo_prod_bi," & _
				"                                          prop_amount_cpo_prod_bi" & _
				"                                     FROM (SELECT qty_bi.fcba," & _
				"                                                  qty_bi.month_prod_bi," & _
				"                                                  qty_bi.year_prod_bi," & _
				"                                                  qty_bi.qty_cpo_prod_bi," & _
				"                                                  qty_bi.qty_ker_prod_bi," & _
				"                                                  biaya_bi.amount_prod_bi," & _
				"                                                  (  biaya_bi.amount_prod_bi" & _
				"                                                   / (  qty_bi.qty_cpo_prod_bi" & _
				"                                                      + (  qty_bi.qty_ker_prod_bi" & _
				"                                                         / 2)))" & _
				"                                                     AS prop_cpo_prod_bi," & _
				"                                                  (  (  biaya_bi.amount_prod_bi" & _
				"                                                      / (  qty_bi.qty_cpo_prod_bi" & _
				"                                                         + (  qty_bi.qty_ker_prod_bi" & _
				"                                                            / 2)))" & _
				"                                                   / 2)" & _
				"                                                     AS prop_ker_prod_bi," & _
				"                                                  ROUND (" & _
				"                                                       qty_bi.qty_cpo_prod_bi" & _
				"                                                     * ( (  biaya_bi.amount_prod_bi" & _
				"                                                          / (  qty_bi.qty_cpo_prod_bi" & _
				"                                                             + (  qty_bi.qty_ker_prod_bi" & _
				"                                                                / 2))))," & _
				"                                                     2)" & _
				"                                                     AS prop_amount_cpo_prod_bi," & _
				"                                                  ROUND (" & _
				"                                                       qty_bi.qty_ker_prod_bi" & _
				"                                                     * ( (  (  biaya_bi.amount_prod_bi" & _
				"                                                             / (  qty_bi.qty_cpo_prod_bi" & _
				"                                                                + (  qty_bi.qty_ker_prod_bi" & _
				"                                                                   / 2)))" & _
				"                                                          / 2))," & _
				"                                                     2)" & _
				"                                                     AS prop_amount_ker_prod_bi" & _
				"                                             FROM (  SELECT fcba," & _
				"                                                            month_prod_bi," & _
				"                                                            year_prod_bi," & _
				"                                                            SUM (" & _
				"                                                               DECODE (" & _
				"                                                                  product_code," & _
				"                                                                  'CPO', qty_prod_bi," & _
				"                                                                  0))" & _
				"                                                               AS qty_cpo_prod_bi," & _
				"                                                            SUM (" & _
				"                                                               DECODE (" & _
				"                                                                  product_code," & _
				"                                                                  'KER', qty_prod_bi," & _
				"                                                                  0))" & _
				"                                                               AS qty_ker_prod_bi" & _
				"                                                       FROM (  SELECT iplasprod.mill_production.fcba" & _
				"                                                                         AS fcba," & _
				"                                                                      iplasprod.mill_production.product_code," & _
				"                                                                      MAX (" & _
				"                                                                         TO_NUMBER (" & _
				"                                                                            TO_CHAR (" & _
				"                                                                               iplasprod.mill_production.tdate," & _
				"                                                                               'mm')))" & _
				"                                                                         AS month_prod_bi," & _
				"                                                                      MAX (" & _
				"                                                                         TO_NUMBER (" & _
				"                                                                            TO_CHAR (" & _
				"                                                                               iplasprod.mill_production.tdate," & _
				"                                                                               'yyyy')))" & _
				"                                                                         AS year_prod_bi," & _
				"                                                                      SUM (" & _
				"                                                                         iplasprod.mill_production.production_today)" & _
				"                                                                         AS qty_prod_bi" & _
				"                                                                 FROM iplasprod.mill_production" & _
				"                                                                WHERE     iplasprod.mill_production.fcba =" & _
				"                                                                             '"& Session("fcba") &"'" & _
				"                                                                      AND TO_NUMBER (" & _
				"                                                                             TO_CHAR (" & _
				"                                                                                iplasprod.mill_production.tdate," & _
				"                                                                                'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
				"                                                                      AND TO_NUMBER (" & _
				"                                                                             TO_CHAR (" & _
				"                                                                                iplasprod.mill_production.tdate," & _
				"                                                                                'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _
				"                                                                      AND iplasprod.mill_production.product_code IN" & _
				"                                                                             ('CPO'," & _
				"                                                                              'KER')" & _
				"                                                             GROUP BY iplasprod.mill_production.fcba," & _
				"                                                                      iplasprod.mill_production.product_code)" & _
				"                                                   GROUP BY fcba," & _
				"                                                            month_prod_bi," & _
				"                                                            year_prod_bi) qty_bi" & _
				"                                                  INNER JOIN" & _
				"                                                  (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                               AS fcba," & _
				"                                                            MAX (" & _
				"                                                               TO_NUMBER (" & _
				"                                                                  TO_CHAR (" & _
				"                                                                     iplasprod.cost_activity.tdate," & _
				"                                                                     'mm')))" & _
				"                                                               AS month_prod_bi," & _
				"                                                            MAX (" & _
				"                                                               TO_NUMBER (" & _
				"                                                                  TO_CHAR (" & _
				"                                                                     iplasprod.cost_activity.tdate," & _
				"                                                                     'yyyy')))" & _
				"                                                               AS year_prod_bi," & _
				"                                                            SUM (" & _
				"                                                               iplasprod.cost_activity.amount)" & _
				"                                                               AS amount_prod_bi" & _
				"                                                       FROM iplasprod.cost_activity" & _
				"                                                      WHERE     cost_activity.fcba =" & _
				"                                                                   '"& Session("fcba") &"'" & _
				"                                                            AND TO_NUMBER (" & _
				"                                                                   TO_CHAR (" & _
				"                                                                      iplasprod.cost_activity.tdate," & _
				"                                                                      'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
				"                                                            AND TO_NUMBER (" & _
				"                                                                   TO_CHAR (" & _
				"                                                                      iplasprod.cost_activity.tdate," & _
				"                                                                      'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _
				"                                                            AND cost_activity.jobcode IN" & _
				"                                                                   (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                      FROM iplasprod.parameterdetail" & _
				"                                                                     WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                  'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                           AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                  'CPO'" & _
				"                                                                           AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                  '"& Session("fcba") &"')" & _
				"                                                   GROUP BY iplasprod.cost_activity.fcba)" & _
				"                                                  biaya_bi" & _
				"                                                     ON (    qty_bi.fcba =" & _
				"                                                                biaya_bi.fcba" & _
				"                                                         AND qty_bi.month_prod_bi =" & _
				"                                                                biaya_bi.month_prod_bi" & _
				"                                                         AND qty_bi.year_prod_bi =" & _
				"                                                                biaya_bi.year_prod_bi))" & _
				"                                   UNION ALL" & _
				"                                   SELECT fcba," & _
				"                                          'KER' AS product_code," & _
				"                                          month_prod_bi," & _
				"                                          year_prod_bi," & _
				"                                          qty_ker_prod_bi," & _
				"                                          prop_amount_ker_prod_bi" & _
				"                                     FROM (SELECT qty_bi.fcba," & _
				"                                                  qty_bi.month_prod_bi," & _
				"                                                  qty_bi.year_prod_bi," & _
				"                                                  qty_bi.qty_cpo_prod_bi," & _
				"                                                  qty_bi.qty_ker_prod_bi," & _
				"                                                  biaya_bi.amount_prod_bi," & _
				"                                                  (  biaya_bi.amount_prod_bi" & _
				"                                                   / (  qty_bi.qty_cpo_prod_bi" & _
				"                                                      + (  qty_bi.qty_ker_prod_bi" & _
				"                                                         / 2)))" & _
				"                                                     AS prop_cpo_prod_bi," & _
				"                                                  (  (  biaya_bi.amount_prod_bi" & _
				"                                                      / (  qty_bi.qty_cpo_prod_bi" & _
				"                                                         + (  qty_bi.qty_ker_prod_bi" & _
				"                                                            / 2)))" & _
				"                                                   / 2)" & _
				"                                                     AS prop_ker_prod_bi," & _
				"                                                  ROUND (" & _
				"                                                       qty_bi.qty_cpo_prod_bi" & _
				"                                                     * ( (  biaya_bi.amount_prod_bi" & _
				"                                                          / (  qty_bi.qty_cpo_prod_bi" & _
				"                                                             + (  qty_bi.qty_ker_prod_bi" & _
				"                                                                / 2))))," & _
				"                                                     2)" & _
				"                                                     AS prop_amount_cpo_prod_bi," & _
				"                                                  ROUND (" & _
				"                                                       qty_bi.qty_ker_prod_bi" & _
				"                                                     * ( (  (  biaya_bi.amount_prod_bi" & _
				"                                                             / (  qty_bi.qty_cpo_prod_bi" & _
				"                                                                + (  qty_bi.qty_ker_prod_bi" & _
				"                                                                   / 2)))" & _
				"                                                          / 2))," & _
				"                                                     2)" & _
				"                                                     AS prop_amount_ker_prod_bi" & _
				"                                             FROM (  SELECT fcba," & _
				"                                                            month_prod_bi," & _
				"                                                            year_prod_bi," & _
				"                                                            SUM (" & _
				"                                                               DECODE (" & _
				"                                                                  product_code," & _
				"                                                                  'CPO', qty_prod_bi," & _
				"                                                                  0))" & _
				"                                                               AS qty_cpo_prod_bi," & _
				"                                                            SUM (" & _
				"                                                               DECODE (" & _
				"                                                                  product_code," & _
				"                                                                  'KER', qty_prod_bi," & _
				"                                                                  0))" & _
				"                                                               AS qty_ker_prod_bi" & _
				"                                                       FROM (  SELECT iplasprod.mill_production.fcba" & _
				"                                                                         AS fcba," & _
				"                                                                      iplasprod.mill_production.product_code," & _
				"                                                                      MAX (" & _
				"                                                                         TO_NUMBER (" & _
				"                                                                            TO_CHAR (" & _
				"                                                                               iplasprod.mill_production.tdate," & _
				"                                                                               'mm')))" & _
				"                                                                         AS month_prod_bi," & _
				"                                                                      MAX (" & _
				"                                                                         TO_NUMBER (" & _
				"                                                                            TO_CHAR (" & _
				"                                                                               iplasprod.mill_production.tdate," & _
				"                                                                               'yyyy')))" & _
				"                                                                         AS year_prod_bi," & _
				"                                                                      SUM (" & _
				"                                                                         iplasprod.mill_production.production_today)" & _
				"                                                                         AS qty_prod_bi" & _
				"                                                                 FROM iplasprod.mill_production" & _
				"                                                                WHERE     iplasprod.mill_production.fcba =" & _
				"                                                                             '"& Session("fcba") &"'" & _
				"                                                                      AND TO_NUMBER (" & _
				"                                                                             TO_CHAR (" & _
				"                                                                                iplasprod.mill_production.tdate," & _
				"                                                                                'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
				"                                                                      AND TO_NUMBER (" & _
				"                                                                             TO_CHAR (" & _
				"                                                                                iplasprod.mill_production.tdate," & _
				"                                                                                'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _
				"                                                                      AND iplasprod.mill_production.product_code IN" & _
				"                                                                             ('CPO'," & _
				"                                                                              'KER')" & _
				"                                                             GROUP BY iplasprod.mill_production.fcba," & _
				"                                                                      iplasprod.mill_production.product_code)" & _
				"                                                   GROUP BY fcba," & _
				"                                                            month_prod_bi," & _
				"                                                            year_prod_bi) qty_bi" & _
				"                                                  INNER JOIN" & _
				"                                                  (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                               AS fcba," & _
				"                                                            MAX (" & _
				"                                                               TO_NUMBER (" & _
				"                                                                  TO_CHAR (" & _
				"                                                                     iplasprod.cost_activity.tdate," & _
				"                                                                     'mm')))" & _
				"                                                               AS month_prod_bi," & _
				"                                                            MAX (" & _
				"                                                               TO_NUMBER (" & _
				"                                                                  TO_CHAR (" & _
				"                                                                     iplasprod.cost_activity.tdate," & _
				"                                                                     'yyyy')))" & _
				"                                                               AS year_prod_bi," & _
				"                                                            SUM (" & _
				"                                                               iplasprod.cost_activity.amount)" & _
				"                                                               AS amount_prod_bi" & _
				"                                                       FROM iplasprod.cost_activity" & _
				"                                                      WHERE     cost_activity.fcba =" & _
				"                                                                   '"& Session("fcba") &"'" & _
				"                                                            AND TO_NUMBER (" & _
				"                                                                   TO_CHAR (" & _
				"                                                                      iplasprod.cost_activity.tdate," & _
				"                                                                      'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
				"                                                            AND TO_NUMBER (" & _
				"                                                                   TO_CHAR (" & _
				"                                                                      iplasprod.cost_activity.tdate," & _
				"                                                                      'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _
				"                                                            AND cost_activity.jobcode IN" & _
				"                                                                   (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                      FROM iplasprod.parameterdetail" & _
				"                                                                     WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                  'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                           AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                  'KER'" & _
				"                                                                           AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                  '"& Session("fcba") &"')" & _
				"                                                   GROUP BY iplasprod.cost_activity.fcba)" & _
				"                                                  biaya_bi" & _
				"                                                     ON (    qty_bi.fcba =" & _
				"                                                                biaya_bi.fcba" & _
				"                                                         AND qty_bi.month_prod_bi =" & _
				"                                                                biaya_bi.month_prod_bi" & _
				"                                                         AND qty_bi.year_prod_bi =" & _
				"                                                                biaya_bi.year_prod_bi)))" & _
				"                                  STAGE1" & _
				"                                  INNER JOIN iplasprod.parameter STAGE2" & _
				"                                     ON (    STAGE1.fcba = STAGE2.fcba" & _
				"                                         AND STAGE1.product_code = STAGE2.fccode)" & _
				"                            WHERE STAGE2.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'))" & _
				"           UNION ALL" & _
				"           SELECT fcba," & _
				"                  productcode," & _
				"                  month_prod_bi," & _
				"                  year_prod_bi," & _
				"                  DebetCredit," & _
				"                  (SELECT DISTINCT param_11" & _
				"                     FROM iplasprod.parameter" & _
				"                    WHERE     parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                          AND fccode IN ('CPO', 'KER')" & _
				"                          AND fcba = '"& Session("fcba") &"')" & _
				"                     AS jobcode," & _
				"                  amount" & _
				"             FROM (  SELECT fcba," & _
				"                            '' AS productcode," & _
				"                            accountingperiod AS month_prod_bi," & _
				"                            accountingyear AS year_prod_bi," & _
				"                            'Cr' AS DebetCredit," & _
				"                            jobcode," & _
				"                            (SUM (amount) * -1) AS amount" & _
				"                       FROM cost_activity" & _
				"                      WHERE     fcba = '"& Session("fcba") &"'" & _
				"                            AND TO_CHAR (tdate, 'mmyyyy') = '"& Mid(Session("CurP_Sales"),5,2) & Mid(Session("CurP_Sales"),1,4) &"'" & _
				"                            AND jobcode IN" & _
				"                                   (SELECT DISTINCT param_3" & _
				"                                      FROM iplasprod.parameterdetail" & _
				"                                     WHERE     parcode =" & _
				"                                                  'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                           AND parheadcode IN ('CPO', 'KER')" & _
				"                                           AND fcba = '"& Session("fcba") &"'" & _
				"                                           AND param_3 IS NOT NULL)" & _
				"                   GROUP BY fcba," & _
				"                            ''," & _
				"                            accountingperiod," & _
				"                            accountingyear," & _
				"                            'Cr'," & _
				"                            jobcode)" & _
				"           UNION ALL" & _
				"           SELECT fcba," & _
				"                  '' AS productcode," & _
				"                  month_prod_bi," & _
				"                  year_prod_bi," & _
				"                  DebetCredit," & _
				"                  (SELECT DISTINCT param_12" & _
				"                     FROM iplasprod.parameter" & _
				"                    WHERE     parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                          AND fccode IN ('CPO', 'KER')" & _
				"                          AND fcba = '"& Session("fcba") &"')" & _
				"                     AS jobcode," & _
				"                  amount" & _
				"             FROM (  SELECT fcba," & _
				"                            '' AS productcode," & _
				"                            accountingperiod AS month_prod_bi," & _
				"                            accountingyear AS year_prod_bi," & _
				"                            'Cr' AS DebetCredit," & _
				"                            (SUM (amount) * -1) AS amount" & _
				"                       FROM cost_activity" & _
				"                      WHERE     fcba = '"& Session("fcba") &"'" & _
				"                            AND TO_CHAR (tdate, 'mmyyyy') = '"& Mid(Session("CurP_Sales"),5,2) & Mid(Session("CurP_Sales"),1,4) &"'" & _
				"                            AND jobcode IN" & _
				"                                   (SELECT DISTINCT param_2" & _
				"                                      FROM iplasprod.parameterdetail" & _
				"                                     WHERE     parcode =" & _
				"                                                  'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                           AND parheadcode IN ('CPO', 'KER')" & _
				"                                           AND fcba = '"& Session("fcba") &"'" & _
				"                                           AND param_2 IS NOT NULL)" & _
				"                   GROUP BY fcba," & _
				"                            ''," & _
				"                            accountingperiod," & _
				"                            accountingyear," & _
				"                            'Cr'))" & _
				" GROUP BY fcba," & _
				"          month_prod_bi," & _
				"          year_prod_bi," & _
				"          DebetCredit," & _
				"          jobcode" 

	' response.write qSalesDB
	Set rsSalesDB = Server.CreateObject("ADODB.Recordset")
	rsSalesDB.Open qSalesDB, connstr
	
	if not rsSalesDB.eof then
		Do While Not rsSalesDB.eof		  
			 
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityDBHPP="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & temp_datePeriodControl & "','mm/dd/yyyy'),'" & rsSalesDB("JOBCODE") & "','SL','SL','" & rsSalesDB("HARGAPOKOKPRODUKSI") & "'," & _
								 "'','CB','BS'," & _
								 "'" &  rsSalesDB("amount_prod_bi") & "','" & rsSalesDB("HARGAPOKOKPRODUKSI") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '',''," & _
								 " '',''," & _
								 " '','', " & _
								 " '','', " & _
								 " '', '','" & rsSalesDB("DebetCredit") & "', " & _
								 " " & rsSalesDB("month_prod_bi") & "," & rsSalesDB("year_prod_bi") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "
			' response.write qCost_ActivityDBHPP
			 connstr.execute(qCost_ActivityDBHPP)
			 connStr.execute("Commit")

		rsSalesDB.MoveNext
	  loop
	  response.Write "Process Harga Pokok Produksi Sales Debet Credit Complete.... !!<br>"
	  response.Flush
	end if
	'-- End Insert Debet
	
	'-- Insert Credit 
	' qSalesCRInti =	" SELECT ROWNUM lineno," & _
					' "        fcba," & _
					' "        (SELECT DISTINCT fccode" & _
					' "           FROM iplasprod.parameterdetail" & _
					' "          WHERE     parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
					' "                AND fcba = '"& session("fcBA") &"'" & _
					' "                AND param_1 = 'ALOKASI_TBS')" & _
					' "           jobcode," & _
					' "        hargapokokproduksi," & _
					' "        month_prod_bi," & _
					' "        year_prod_bi," & _
					' "        amount_prod_bi" & _
					' "   FROM (  SELECT fcba," & _
					' "                  hargapokokproduksi," & _
					' "                  month_prod_bi," & _
					' "                  year_prod_bi," & _
					' "                  (-1 * SUM (amount_prod_bi)) amount_prod_bi" & _
					' "             FROM (  SELECT iplasprod.cost_activity.fcba," & _
					' "                            iplasprod.cost_activity.jobcode," & _
					' "                            'CB'," & _
					' "                            'BS'," & _
					' "                            MAX (" & _
					' "                                  'HARGAPOKOKPRODUKSI'" & _
					' "                               || '-'" & _
					' "                               || TO_CHAR (iplasprod.cost_activity.tdate, 'mm')" & _
					' "                               || '/'" & _
					' "                               || TO_CHAR (iplasprod.cost_activity.tdate, 'yyyy'))" & _
					' "                               HARGAPOKOKPRODUKSI," & _
					' "                            MAX (" & _
					' "                               TO_NUMBER (" & _
					' "                                  TO_CHAR (iplasprod.cost_activity.tdate, 'mm')))" & _
					' "                               AS month_prod_bi," & _
					' "                            MAX (" & _
					' "                               TO_NUMBER (" & _
					' "                                  TO_CHAR (iplasprod.cost_activity.tdate, 'yyyy')))" & _
					' "                               AS year_prod_bi," & _
					' "                            SUM (iplasprod.cost_activity.amount) AS amount_prod_bi" & _
					' "                       FROM iplasprod.cost_activity" & _
					' "                      WHERE     cost_activity.fcba = '"& session("fcBA") &"'" & _
					' "                            AND TO_NUMBER (" & _
					' "                                   TO_CHAR (iplasprod.cost_activity.tdate, 'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
					' "                            AND TO_NUMBER (" & _
					' "                                   TO_CHAR (iplasprod.cost_activity.tdate, 'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _ 
					' "                            AND cost_activity.jobcode IN" & _
					' "                                   (SELECT fccode" & _
					' "                                      FROM iplasprod.parameterdetail" & _
					' "                                     WHERE     parcode =" & _
					' "                                                  'JOBCODE_HARGAPOKOKPRODUKSI'" & _
					' "                                           AND fcba = '"& session("fcBA") &"')" & _
					' "                   GROUP BY iplasprod.cost_activity.fcba," & _
					' "                            iplasprod.cost_activity.jobcode)" & _
					' "         GROUP BY fcba," & _
					' "                  hargapokokproduksi," & _
					' "                  month_prod_bi," & _
					' "                  year_prod_bi)"  

	
	' response.write qSalesCR
	' Set rsSalesCRInti = Server.CreateObject("ADODB.Recordset")
	' rsSalesCRInti.Open qSalesCRInti, connstr
	
	' if not rsSalesCRInti.eof then
		' Do While Not rsSalesCRInti.eof	   
		
			 ' maxLineNo = maxLineNo + 1
			 ' qCost_ActivityCr="Insert into Cost_Activity " & _
							 ' "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 ' "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 ' " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 ' " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 ' " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 ' " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 ' " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 ' " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 ' " LASTUPDATE, LASTTIME ) " & _
							 ' " values (" & maxLineNo & ",sysdate,'" & rsSalesCRInti("JOBCODE") & "','SL','SL','" & rsSalesCRInti("HARGAPOKOKPRODUKSI") & "'," & _
							 ' "'','CB','BS'," & _
							 ' "'" &  rsSalesCRInti("amount_prod_bi") & "','" & rsSalesCRInti("HARGAPOKOKPRODUKSI") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 ' " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 ' " '',''," & _
							 ' " '',''," & _
							 ' " '','', " & _
							 ' " '','', " & _
							 ' " 'Line No:', '" & rsSalesCRInti("LINENO") & "','Cr', " & _
							 ' " " & rsSalesCRInti("month_prod_bi") & "," & rsSalesCRInti("year_prod_bi") & "," & _
							 ' " sysdate, to_char(sysdate,'HH24:MI')) "

			 ' connstr.execute(qCost_ActivityCr)
			 ' connStr.execute("Commit")

		  ' rsSalesCRInti.MoveNext
	  ' loop
	  	  ' response.Write "Process Harga Pokok Produksi Sales Credit Complete.... !!<br>"
		  ' response.Flush
	' end if
	'-- End Insert Credit 
	

	'-- Insert Table Harga Pokok Penjualan
	qSalesDB =	" select " & _
				" fcba," & _
				" product_code," & _
				" month_saldo_awal," & _
				" year_saldo_awal," & _
				" saldo_qty_awal," & _
				" saldo_amount_awal," & _
				" Price_saldo_awal," & _
				" month_prod_sbi," & _
				" year_prod_sbi," & _
				" qty_prod_sbi," & _
				" last_month_qty_prod_sbi," & _
				" amount_prod_sbi," & _
				" last_month_amount_prod_sbi," & _
				" Price_prod_sbi," & _
				" month_qty_sales_sbi," & _
				" year_qty_sales_sbi," & _
				" qty_sales_sbi," & _
				" last_month_qty_sales_sbi," & _
				" (qty_sales_sbi-0) as qty_sales_bi," & _
				" HPP_Qty_level_1," & _
				" HPP_Amount_level_1," & _
				" HPP_Price_level_1," & _
				" HPP_Qty_level_2," & _
				" HPP_Amount_level_2," & _
				" HPP_Price_level_2," & _
				" amount_sales_sbi," & _
				" last_month_amount_sales_sbi," & _
				" (amount_sales_sbi-0) as amount_sales_bi," & _
				" price_sales_sbi," & _
				" ((amount_sales_sbi-0)/(qty_sales_sbi-0)) as price_sales_bi," & _
				" saldo_Qty_akhir," & _
				" saldo_amount_akhir," & _
				" price_saldo_akhir" & _
				" from (" & _
				"     select " & _
				"     fcba," & _
				"     product_code," & _
				"     month_saldo_awal," & _
				"     year_saldo_awal," & _
				"     saldo_qty_awal," & _
				"     saldo_amount_awal," & _
				"     Price_saldo_awal," & _
				"     month_prod_sbi," & _
				"     year_prod_sbi," & _
				"     qty_prod_sbi," & _
				"     nvl(IPLASPROD.GetLastMonthqtysbi(product_code, fcba,LAST_DAY(ADD_MONTHS(to_date('"& Mid(Session("CurP_Sales"),5,2) &"','fmMM'),-1))),0) last_month_qty_prod_sbi," & _
				"     amount_prod_sbi," & _
				"     nvl(IPLASPROD.GetLastMonthamountprodsbi(product_code, fcba,LAST_DAY(ADD_MONTHS(to_date('"& Mid(Session("CurP_Sales"),5,2) &"','fmMM'),-1))),0) last_month_amount_prod_sbi," & _
				"     Price_prod_sbi," & _
				"     month_qty_sales_sbi," & _
				"     year_qty_sales_sbi," & _
				"     qty_sales_sbi," & _
				"     nvl(IPLASPROD.GetLastMonthqtysalessbi(product_code, fcba,LAST_DAY(ADD_MONTHS(to_date('"& Mid(Session("CurP_Sales"),5,2) &"','fmMM'),-1))),0) last_month_qty_sales_sbi," & _
				"     HPP_Qty_level_1," & _
				"     HPP_Amount_level_1," & _
				"     HPP_Price_level_1," & _
				"     HPP_Qty_level_2," & _
				"     HPP_Amount_level_2," & _
				"     HPP_Price_level_2," & _
				"     amount_sales_sbi," & _
				"     nvl(IPLASPROD.GetLastMonthamountsalessbi(product_code, fcba,LAST_DAY(ADD_MONTHS(to_date('"& Mid(Session("CurP_Sales"),5,2) &"','fmMM'),-1))),0) last_month_amount_sales_sbi," & _
				"     price_sales_sbi," & _
				"     saldo_Qty_akhir," & _
				"     saldo_amount_akhir," & _
				"     price_saldo_akhir" & _
				"     from (" & _
				"         select " & _
				"         fcba," & _
				"         product_code," & _
				"         month_saldo_awal," & _
				"         year_saldo_awal," & _
				"         saldo_qty_awal," & _
				"         saldo_amount_awal," & _
				"         Price_saldo_awal," & _
				"         month_prod_sbi," & _
				"         year_prod_sbi," & _
				"         qty_prod_sbi," & _
				"         amount_prod_sbi," & _
				"         Price_prod_sbi," & _
				"         month_qty_sales_sbi," & _
				"         year_qty_sales_sbi," & _
				"         qty_sales_sbi," & _
				"         HPP_Qty_level_1," & _
				"         HPP_Amount_level_1," & _
				"         HPP_Price_level_1," & _
				"         HPP_Qty_level_2," & _
				"         HPP_Amount_level_2," & _
				"         HPP_Price_level_2," & _
				"         amount_sales_sbi," & _
				"         price_sales_sbi," & _
				"         saldo_Qty_akhir," & _
				"         saldo_amount_akhir," & _
				"         price_saldo_akhir" & _
				"         from (" & _
				"             select" & _
				"             fcba," & _
				"             product_code," & _
				"             month_saldo_awal," & _
				"             year_saldo_awal," & _
				"             saldo_qty_awal," & _
				"             saldo_amount_awal," & _
				"             Price_saldo_awal," & _
				"             month_prod_sbi," & _
				"             year_prod_sbi," & _
				"             qty_prod_sbi," & _
				"             amount_prod_sbi," & _
				"             Price_prod_sbi," & _
				"             month_qty_sales_sbi," & _
				"             year_qty_sales_sbi," & _
				"             qty_sales_sbi," & _
				"             HPP_Qty_level_1," & _
				"             HPP_Amount_level_1," & _
				"             HPP_Price_level_1," & _
				"             HPP_Qty_level_2," & _
				"             HPP_Amount_level_2," & _
				"             HPP_Price_level_2," & _
				"             (HPP_Amount_level_1+HPP_Amount_level_2) as amount_sales_sbi," & _
				"             decode(qty_sales_sbi,0,0,((HPP_Amount_level_1+HPP_Amount_level_2)/qty_sales_sbi)) as price_sales_sbi," & _
				"             (saldo_qty_awal+qty_prod_sbi-qty_sales_sbi) as saldo_Qty_akhir," & _
				"             (saldo_amount_awal+amount_prod_sbi-HPP_Amount_level_1-HPP_Amount_level_2) as saldo_amount_akhir," & _
				"             ((saldo_amount_awal+amount_prod_sbi-HPP_Amount_level_1-HPP_Amount_level_2)/(saldo_qty_awal+qty_prod_sbi-qty_sales_sbi)) as price_saldo_akhir" & _
				"             from (" & _
				"                 select" & _
				"                 fcba," & _
				"                 product_code," & _
				"                 month_saldo_awal," & _
				"                 year_saldo_awal," & _
				"                 saldo_qty_awal," & _
				"                 saldo_amount_awal," & _
				"                 Price_saldo_awal," & _
				"                 month_prod_sbi," & _
				"                 year_prod_sbi," & _
				"                 qty_prod_sbi," & _
				"                 amount_prod_sbi," & _
				"                 Price_prod_sbi," & _
				"                 month_qty_sales_sbi," & _
				"                 year_qty_sales_sbi," & _
				"                 qty_sales_sbi," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then saldo_qty_awal" & _
				"                     else qty_sales_sbi" & _
				"                     end as HPP_Qty_level_1," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then Price_saldo_awal" & _
				"                     else Price_saldo_awal" & _
				"                     end as HPP_Price_level_1," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then (saldo_qty_awal*Price_saldo_awal)" & _
				"                     else (saldo_qty_awal*Price_saldo_awal)" & _
				"                     end as HPP_Amount_level_1," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then (qty_sales_sbi-saldo_qty_awal)" & _
				"                     else 0" & _
				"                     end as HPP_Qty_level_2," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then Price_prod_sbi" & _
				"                     else 0" & _
				"                     end as HPP_Price_level_2," & _
				"                 case when" & _
				"                     qty_sales_sbi >= saldo_qty_awal then ((qty_sales_sbi-saldo_qty_awal)*Price_prod_sbi)" & _
				"                     else 0" & _
				"                     end as HPP_Amount_level_2" & _
				"                 from (" & _
				"                     select" & _
				"                     fcba," & _
				"                     product_code," & _
				"                     max(month_saldo_awal) as month_saldo_awal," & _
				"                     max(year_saldo_awal) as year_saldo_awal," & _
				"                     max(saldo_qty_awal) as saldo_qty_awal," & _
				"                     max(saldo_amount_awal) as saldo_amount_awal," & _
				"                     max(Price_saldo_awal) as Price_saldo_awal," & _
				"                     max(month_prod_sbi) as month_prod_sbi," & _
				"                     max(year_prod_sbi) as year_prod_sbi," & _
				"                     max(qty_prod_sbi) as qty_prod_sbi," & _
				"                     max(amount_prod_sbi) as amount_prod_sbi," & _
				"                     max(Price_prod_sbi) as Price_prod_sbi," & _
				"                     max(month_qty_sales_sbi) as month_qty_sales_sbi," & _
				"                     max(year_qty_sales_sbi) as year_qty_sales_sbi," & _
				"                     max(qty_sales_sbi) as qty_sales_sbi" & _
				"                     from (" & _
				"                         select" & _
				"                         fcba," & _
				"                         product_code," & _
				"                         month_saldo_awal," & _
				"                         year_saldo_awal," & _
				"                         qty_cpo_prod_awal as saldo_qty_awal," & _
				"                         prop_amount_cpo_prod_awal as saldo_amount_awal," & _
				"                         (prop_amount_cpo_prod_awal/qty_cpo_prod_awal) as Price_saldo_awal," & _
				"                         0 as month_prod_sbi," & _
				"                         0 as year_prod_sbi," & _
				"                         0 as qty_prod_sbi," & _
				"                         0 as amount_prod_sbi," & _
				"                         0 as Price_prod_sbi," & _
				"                         0 as month_qty_sales_sbi," & _
				"                         0 as year_qty_sales_sbi," & _
				"                         0 as qty_sales_sbi" & _
				"                         from " & _
				"                         (" & _
				"                             select" & _
				"                             fcba," & _
				"                             'CPO' as product_code," & _
				"                             month_saldo_awal," & _
				"                             year_saldo_awal," & _
				"                             qty_cpo_prod_awal," & _
				"                             prop_amount_cpo_prod_awal" & _
				"                             from (" & _
				"                                 select" & _
				"                                 qty_awal.fcba," & _
				"                                 qty_awal.month_saldo_awal," & _
				"                                 qty_awal.year_saldo_awal," & _
				"                                 qty_awal.qty_cpo_prod_awal," & _
				"                                 qty_awal.qty_ker_prod_awal," & _
				"                                 biaya_awal.saldo_amount_awal," & _
				"                                 (biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2))) as prop_cpo_prod_awal," & _
				"                                 ((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))/2) as prop_ker_prod_awal," & _
				"                                 round(qty_awal.qty_cpo_prod_awal*((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))),0) as prop_amount_cpo_prod_awal," & _
				"                                 round(qty_awal.qty_ker_prod_awal*(((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))/2)),0) as prop_amount_ker_prod_awal" & _
				"                                 from (" & _
				"                                     select" & _
				"                                     fcba,month_saldo_awal,year_saldo_awal," & _
				"                                     sum(decode(product_code,'CPO',saldo_qty_awal,0)) as qty_cpo_prod_awal," & _
				"                                     sum(decode(product_code,'KER',saldo_qty_awal,0)) as qty_ker_prod_awal" & _
				"                                     from (" & _
				"                                         select iplasprod.mill_production.fcba as fcba,iplasprod.mill_production.product_code," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'mm'))) as month_saldo_awal," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'yyyy'))) as year_saldo_awal," & _
				"                                         sum(iplasprod.mill_production.production_today) as saldo_qty_awal " & _
				"                                         from iplasprod.mill_production" & _
				"                                         where iplasprod.mill_production.fcba = '"& session("fcba") &"'" & _
				"                                         and to_char(iplasprod.mill_production.tdate,'mm')= '12'" & _
				"                                         and to_char(iplasprod.mill_production.tdate,'yyyy')= '"& Mid(Session("CurP_Sales"),1,4) &"' - 1" & _
				"                                         and iplasprod.mill_production.product_code in ('CPO','KER')" & _
				"                                         group by iplasprod.mill_production.fcba,iplasprod.mill_production.product_code" & _
				"                                     )" & _
				"                                     group by fcba,month_saldo_awal,year_saldo_awal" & _
				"                                 ) qty_awal" & _
				"                                 inner join (" & _
				"                                     select iplasprod.cost_activity.fcba as fcba," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'mm'))) as month_saldo_awal," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))) as year_saldo_awal," & _
				"                                     sum(iplasprod.cost_activity.amount) as saldo_amount_awal " & _
				"                                     from iplasprod.cost_activity" & _
				"                                     where cost_activity.fcba = '" & session("fcba") & "'" & _
				"                                     and to_char(iplasprod.cost_activity.tdate,'mm')= '12'" & _
				"                                     and to_char(iplasprod.cost_activity.tdate,'yyyy')= '"& Mid(Session("CurP_Sales"),1,4) &"' - 1" & _
				"                                     and cost_activity.jobcode in (select fccode from parameterdetail where parcode = 'JOBCODE_HARGAPOKOKPRODUKSI' and fcba= '"& session("fcba") &"')" & _
				"                                     group by iplasprod.cost_activity.fcba" & _
				"                                 ) biaya_awal" & _
				"                                 on (" & _
				"                                 qty_awal.fcba = biaya_awal.fcba and" & _
				"                                 qty_awal.month_saldo_awal = biaya_awal.month_saldo_awal and" & _
				"                                 qty_awal.year_saldo_awal = biaya_awal.year_saldo_awal)" & _
				"                             )" & _
				"                             union all" & _
				"                             select" & _
				"                             fcba," & _
				"                             'KER' as product_code," & _
				"                             month_saldo_awal," & _
				"                             year_saldo_awal," & _
				"                             qty_ker_prod_awal," & _
				"                             prop_amount_ker_prod_awal" & _
				"                             from (" & _
				"                                 select" & _
				"                                 qty_awal.fcba," & _
				"                                 qty_awal.month_saldo_awal," & _
				"                                 qty_awal.year_saldo_awal," & _
				"                                 qty_awal.qty_cpo_prod_awal," & _
				"                                 qty_awal.qty_ker_prod_awal," & _
				"                                 biaya_awal.saldo_amount_awal," & _
				"                                 (biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2))) as prop_cpo_prod_awal," & _
				"                                 ((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))/2) as prop_ker_prod_awal," & _
				"                                 round(qty_awal.qty_cpo_prod_awal*((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))),0) as prop_amount_cpo_prod_awal," & _
				"                                 round(qty_awal.qty_ker_prod_awal*(((biaya_awal.saldo_amount_awal/(qty_awal.qty_cpo_prod_awal+(qty_awal.qty_ker_prod_awal/2)))/2)),0) as prop_amount_ker_prod_awal" & _
				"                                 from (" & _
				"                                     select" & _
				"                                     fcba,month_saldo_awal,year_saldo_awal," & _
				"                                     sum(decode(product_code,'CPO',saldo_qty_awal,0)) as qty_cpo_prod_awal," & _
				"                                     sum(decode(product_code,'KER',saldo_qty_awal,0)) as qty_ker_prod_awal" & _
				"                                     from (" & _
				"                                         select iplasprod.mill_production.fcba as fcba,iplasprod.mill_production.product_code," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'mm'))) as month_saldo_awal," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'yyyy'))) as year_saldo_awal," & _
				"                                         sum(iplasprod.mill_production.production_today) as saldo_qty_awal " & _
				"                                         from iplasprod.mill_production" & _
				"                                         where iplasprod.mill_production.fcba = '"& session("fcba") &"'" & _
				"                                         and to_char(iplasprod.mill_production.tdate,'mm')= '12'" & _
				"                                         and to_char(iplasprod.mill_production.tdate,'yyyy')= '"& Mid(Session("CurP_Sales"),1,4) &"' - 1" & _
				"                                         and iplasprod.mill_production.product_code in ('CPO','KER')" & _
				"                                         group by iplasprod.mill_production.fcba,iplasprod.mill_production.product_code" & _
				"                                     )" & _
				"                                     group by fcba,month_saldo_awal,year_saldo_awal" & _
				"                                 ) qty_awal" & _
				"                                 inner join (" & _
				"                                     select iplasprod.cost_activity.fcba as fcba," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'mm'))) as month_saldo_awal," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))) as year_saldo_awal," & _
				"                                     sum(iplasprod.cost_activity.amount) as saldo_amount_awal " & _
				"                                     from iplasprod.cost_activity" & _
				"                                     where cost_activity.fcba = '"& session("fcba") &"'" & _
				"                                     and to_char(iplasprod.cost_activity.tdate,'mm')= '12'" & _
				"                                     and to_char(iplasprod.cost_activity.tdate,'yyyy')= '"& Mid(Session("CurP_Sales"),1,4) &"' - 1" & _
				"                                     and cost_activity.jobcode in (select fccode from parameterdetail where parcode = 'JOBCODE_HARGAPOKOKPRODUKSI' and fcba= '"& session("fcba") &"')" & _
				"                                     group by iplasprod.cost_activity.fcba" & _
				"                                 ) biaya_awal" & _
				"                                 on (" & _
				"                                 qty_awal.fcba = biaya_awal.fcba and" & _
				"                                 qty_awal.month_saldo_awal = biaya_awal.month_saldo_awal and" & _
				"                                 qty_awal.year_saldo_awal = biaya_awal.year_saldo_awal)" & _
				"                             )" & _
				"                         )" & _
				"                         where product_code in ('CPO','KER')" & _
				"                         union all" & _
				"                         select            " & _
				"                         fcba," & _
				"                         product_code," & _
				"                         0 as month_saldo_awal," & _
				"                         0 as year_saldo_awal," & _
				"                         0 as saldo_qty_awal," & _
				"                         0 as saldo_amount_awal," & _
				"                         0 as Price_saldo_awal," & _
				"                         month_prod_sbi," & _
				"                         year_prod_sbi," & _
				"                         qty_cpo_prod_sbi as qty_prod_sbi," & _
				"                         prop_amount_cpo_prod_sbi as amount_prod_sbi," & _
				"                         (prop_amount_cpo_prod_sbi/qty_cpo_prod_sbi) as Price_prod_sbi," & _
				"                         0 as month_qty_sales_sbi," & _
				"                         0 as year_qty_sales_sbi," & _
				"                         0 as qty_sales_sbi" & _
				"                         from (" & _
				"                             select" & _
				"                             fcba," & _
				"                             'CPO' as product_code," & _
				"                             month_prod_sbi," & _
				"                             year_prod_sbi," & _
				"                             qty_cpo_prod_sbi," & _
				"                             prop_amount_cpo_prod_sbi" & _
				"                             from (" & _
				"                                 select" & _
				"                                 qty_sbi.fcba," & _
				"                                 qty_sbi.month_prod_sbi," & _
				"                                 qty_sbi.year_prod_sbi," & _
				"                                 qty_sbi.qty_cpo_prod_sbi," & _
				"                                 qty_sbi.qty_ker_prod_sbi," & _
				"                                 biaya_sbi.amount_prod_sbi," & _
				"                                 (biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2))) as prop_cpo_prod_sbi," & _
				"                                 ((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))/2) as prop_ker_prod_sbi," & _
				"                                 round(qty_sbi.qty_cpo_prod_sbi*((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))),0) as prop_amount_cpo_prod_sbi," & _
				"                                 round(qty_sbi.qty_ker_prod_sbi*(((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))/2)),0) as prop_amount_ker_prod_sbi" & _
				"                                 from (" & _
				"                                     select" & _
				"                                     fcba,month_prod_sbi,year_prod_sbi," & _
				"                                     sum(decode(product_code,'CPO',qty_prod_sbi,0)) as qty_cpo_prod_sbi," & _
				"                                     sum(decode(product_code,'KER',qty_prod_sbi,0)) as qty_ker_prod_sbi" & _
				"                                     from (" & _
				"                                         select iplasprod.mill_production.fcba as fcba,iplasprod.mill_production.product_code," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'mm'))) as month_prod_sbi," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'yyyy'))) as year_prod_sbi," & _
				"                                         sum(iplasprod.mill_production.production_today) as qty_prod_sbi " & _
				"                                         from iplasprod.mill_production" & _
				"                                         where iplasprod.mill_production.fcba = '"& session("fcba") &"'" & _
				"                                         and (iplasprod.mill_production.tdate between to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') and sysdate) " & _
				"                                         and iplasprod.mill_production.product_code in ('CPO','KER')" & _
				"                                         group by iplasprod.mill_production.fcba,iplasprod.mill_production.product_code" & _
				"                                     )" & _
				"                                     group by fcba,month_prod_sbi,year_prod_sbi" & _
				"                                 ) qty_sbi" & _
				"                                 inner join (" & _
				"                                     select iplasprod.cost_activity.fcba as fcba," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'mm'))) as month_prod_sbi," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))) as year_prod_sbi," & _
				"                                     sum(iplasprod.cost_activity.amount) as amount_prod_sbi " & _
				"                                     from iplasprod.cost_activity" & _
				"                                     where cost_activity.fcba = '"& session("fcba") &"'" & _
				"                                     and (iplasprod.cost_activity.tdate between to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') and sysdate) " & _
				"                                     and cost_activity.jobcode in (select fccode from parameterdetail where parcode = 'JOBCODE_HARGAPOKOKPRODUKSI' and fcba= '"& session("fcba") &"')" & _
				"                                     group by iplasprod.cost_activity.fcba" & _
				"                                 ) biaya_sbi" & _
				"                                 on (" & _
				"                                 qty_sbi.fcba = biaya_sbi.fcba and" & _
				"                                 qty_sbi.month_prod_sbi = biaya_sbi.month_prod_sbi and" & _
				"                                 qty_sbi.year_prod_sbi = biaya_sbi.year_prod_sbi)" & _
				"                             )" & _
				"                             union all" & _
				"                             select" & _
				"                             fcba," & _
				"                             'KER' as product_code," & _
				"                             month_prod_sbi," & _
				"                             year_prod_sbi," & _
				"                             qty_ker_prod_sbi," & _
				"                             prop_amount_ker_prod_sbi" & _
				"                             from (" & _
				"                                 select" & _
				"                                 qty_sbi.fcba," & _
				"                                 qty_sbi.month_prod_sbi," & _
				"                                 qty_sbi.year_prod_sbi," & _
				"                                 qty_sbi.qty_cpo_prod_sbi," & _
				"                                 qty_sbi.qty_ker_prod_sbi," & _
				"                                 biaya_sbi.amount_prod_sbi," & _
				"                                 (biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2))) as prop_cpo_prod_sbi," & _
				"                                 ((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))/2) as prop_ker_prod_sbi," & _
				"                                 round(qty_sbi.qty_cpo_prod_sbi*((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))),0) as prop_amount_cpo_prod_sbi," & _
				"                                 round(qty_sbi.qty_ker_prod_sbi*(((biaya_sbi.amount_prod_sbi/(qty_sbi.qty_cpo_prod_sbi+(qty_sbi.qty_ker_prod_sbi/2)))/2)),0) as prop_amount_ker_prod_sbi" & _
				"                                 from (" & _
				"                                     select" & _
				"                                     fcba,month_prod_sbi,year_prod_sbi," & _
				"                                     sum(decode(product_code,'CPO',qty_prod_sbi,0)) as qty_cpo_prod_sbi," & _
				"                                     sum(decode(product_code,'KER',qty_prod_sbi,0)) as qty_ker_prod_sbi" & _
				"                                     from (" & _
				"                                         select iplasprod.mill_production.fcba as fcba,iplasprod.mill_production.product_code," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'mm'))) as month_prod_sbi," & _
				"                                         max(to_number(to_char(iplasprod.mill_production.tdate,'yyyy'))) as year_prod_sbi," & _
				"                                         sum(iplasprod.mill_production.production_today) as qty_prod_sbi " & _
				"                                         from iplasprod.mill_production" & _
				"                                         where iplasprod.mill_production.fcba = '"& session("fcba") &"'" & _
				"                                         and (iplasprod.mill_production.tdate between to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') and sysdate) " & _
				"                                         and iplasprod.mill_production.product_code in ('CPO','KER')" & _
				"                                         group by iplasprod.mill_production.fcba,iplasprod.mill_production.product_code" & _
				"                                     )" & _
				"                                     group by fcba,month_prod_sbi,year_prod_sbi" & _
				"                                 ) qty_sbi" & _
				"                                 inner join (" & _
				"                                     select iplasprod.cost_activity.fcba as fcba," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'mm'))) as month_prod_sbi," & _
				"                                     max(to_number(to_char(iplasprod.cost_activity.tdate,'yyyy'))) as year_prod_sbi," & _
				"                                     sum(iplasprod.cost_activity.amount) as amount_prod_sbi " & _
				"                                     from iplasprod.cost_activity" & _
				"                                     where cost_activity.fcba = '"& session("fcba") &"'" & _
				"                                     and (iplasprod.cost_activity.tdate between to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') and sysdate) " & _
				"                                     and cost_activity.jobcode in (select fccode from parameterdetail where parcode = 'JOBCODE_HARGAPOKOKPRODUKSI' and fcba= '"& session("fcba") &"')" & _
				"                                     group by iplasprod.cost_activity.fcba" & _
				"                                 ) biaya_sbi" & _
				"                                 on (" & _
				"                                 qty_sbi.fcba = biaya_sbi.fcba and" & _
				"                                 qty_sbi.month_prod_sbi = biaya_sbi.month_prod_sbi and" & _
				"                                 qty_sbi.year_prod_sbi = biaya_sbi.year_prod_sbi)" & _
				"                             )" & _
				"                         )" & _
				"                         where product_code in ('CPO','KER')" & _
				"                         union all" & _
				"                         select" & _
				"                         fcba," & _
				"                         product_code," & _
				"                         0 as month_saldo_awal," & _
				"                         0 as year_saldo_awal," & _
				"                         0 as saldo_qty_awal," & _
				"                         0 as saldo_amount_awal," & _
				"                         0 as Price_saldo_awal," & _
				"                         0 as month_prod_sbi," & _
				"                         0 as year_prod_sbi," & _
				"                         0 as qty_prod_sbi," & _
				"                         0 as amount_prod_sbi," & _
				"                         0 as Price_prod_sbi," & _
				"                         month_qty_sales_sbi," & _
				"                         year_qty_sales_sbi," & _
				"                         qty_sales_sbi" & _
				"                         from (" & _
				"                             select " & _
				"                             fcba," & _
				"                             product_code," & _
				"                             month_qty_sales_sbi," & _
				"                             year_qty_sales_sbi," & _
				"                             weightin," & _
				"                             weightout," & _
				"                             (weightout-weightin) as qty_sales_sbi" & _
				"                             from (        " & _
				"                                 select iplasprod.weighbridgeticket.fcba,iplasprod.weighbridgeticket.product_productcode as product_code," & _
				"                                 max(to_number(to_char(iplasprod.weighbridgeticket.dateout,'mm'))) as month_qty_sales_sbi," & _
				"                                 max(to_number(to_char(iplasprod.weighbridgeticket.dateout,'yyyy'))) as year_qty_sales_sbi," & _
				"                                 sum(iplasprod.weighbridgeticket.weightin) as weightin," & _
				"                                 sum(iplasprod.weighbridgeticket.weightout) as weightout" & _
				"                                 from iplasprod.weighbridgeticket" & _
				"                                 where iplasprod.weighbridgeticket.fcba = '"& session("fcba") &"'" & _
				"                                 and (iplasprod.weighbridgeticket.dateout between to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') and  sysdate)" & _
				"                                 and iplasprod.weighbridgeticket.product_productcode in ('CPO','KER')" & _
				"                                 group by iplasprod.weighbridgeticket.fcba,iplasprod.weighbridgeticket.product_productcode" & _
				"                             )" & _
				"                         )  " & _
				"                     ) " & _
				"                     group by fcba,product_code" & _
				"                 )" & _
				"             )" & _
				"         )" & _
				"     )" & _
				" )" 

	' response.write qSalesDB & "<br>"
	Set rsSalesDB = Server.CreateObject("ADODB.Recordset")
	rsSalesDB.Open qSalesDB, connstr
	
	if not rsSalesDB.eof then
		Do While Not rsSalesDB.eof		  
			 
			 maxLineNo = maxLineNo + 1
			 q_HPP_Tab		=	 "Insert into hargapokokpenjualan " & _
								 "( product_code, month_saldo_awal, year_saldo_awal, saldo_qty_awal, saldo_amount_awal, Price_saldo_awal, month_prod_sbi, year_prod_sbi, qty_prod_sbi, last_month_qty_prod_sbi," & _
								 "	amount_prod_sbi, last_month_amount_prod_sbi, Price_prod_sbi, month_qty_sales_sbi, year_qty_sales_sbi, qty_sales_sbi, last_month_qty_sales_sbi, qty_sales_bi," & _
								 "	HPP_Qty_level_1, HPP_Amount_level_1, HPP_Price_level_1, HPP_Qty_level_2, HPP_Amount_level_2, HPP_Price_level_2, amount_sales_sbi, last_month_amount_sales_sbi, " & _
								 "	amount_sales_bi, price_sales_sbi, price_sales_bi, saldo_Qty_akhir, saldo_amount_akhir, price_saldo_akhir, fcentry, fcedit, fcip, fcba, lastupdate, lasttime ) " & _
								 " values ('" & rsSalesDB("product_code") & "'" & _
								 " ," & rsSalesDB("month_saldo_awal") &"," & rsSalesDB("year_saldo_awal") & "," & rsSalesDB("saldo_qty_awal") & "" & _
								 " ," & rsSalesDB("saldo_amount_awal") & "," & rsSalesDB("Price_saldo_awal") & "," & rsSalesDB("month_prod_sbi") & "" & _
								 " ," & rsSalesDB("year_prod_sbi") & "," & rsSalesDB("qty_prod_sbi") & "," & rsSalesDB("last_month_qty_prod_sbi") & "" & _ 
								 " ," & rsSalesDB("amount_prod_sbi") & "," & rsSalesDB("last_month_amount_prod_sbi") & "," & rsSalesDB("Price_prod_sbi") & "" & _ 
								 " ," & rsSalesDB("month_qty_sales_sbi") & "," & rsSalesDB("year_qty_sales_sbi") & "," & rsSalesDB("qty_sales_sbi") & "" & _ 
								 " ," & rsSalesDB("last_month_qty_sales_sbi") & "," & rsSalesDB("qty_sales_bi") & "," & rsSalesDB("HPP_Qty_level_1") & "" & _ 
								 " ," & rsSalesDB("HPP_Amount_level_1") & "," & rsSalesDB("HPP_Price_level_1") & "," & rsSalesDB("HPP_Qty_level_2") & "" & _ 
								 " ," & rsSalesDB("HPP_Amount_level_2") & "," & rsSalesDB("HPP_Price_level_2") & "," & rsSalesDB("amount_sales_sbi") & "" & _ 
								 " ," & rsSalesDB("last_month_amount_sales_sbi") & "," & rsSalesDB("amount_sales_bi") & "," & rsSalesDB("price_sales_sbi") & "" & _ 
								 " ," & rsSalesDB("price_sales_bi") & "," & rsSalesDB("saldo_Qty_akhir") & "," & rsSalesDB("saldo_amount_akhir") & "" & _ 
								 " ," & rsSalesDB("price_saldo_akhir") & ",'"& session("codename") &"','"& session("codename") &"'" & _
								 " ,'"& Request.ServerVariables("REMOTE_ADDR") &"','"& session("fcba") &"',sysdate, to_char(sysdate,'HH24:MI')) "
			' response.write q_HPP_Tab
			 connstr.execute(q_HPP_Tab)
			 connStr.execute("Commit")

		rsSalesDB.MoveNext
	  loop
	  response.Write "Insert Harga Pokok Penjualan Sales Complete.... !!<br>"
	  response.Flush
	end if
	'-- End Insert Table Harga Pokok Penjualan 

	'-- Insert Debet Harga Pokok Penjualan
	qSalesDB =	" select" & _
				" fcba," & _
				" month_qty_sales_bi," & _
				" year_qty_sales_bi," & _
				" jobcode_intraco_HO," & _
				" sum(total_amount_sales_bi) as total_amount_sales_bi" & _
				" from (" & _
				" SELECT ROWNUM lineno," & _
				"    HPP_prod_bi.fcba," & _
				"    HPP_prod_bi.product_code," & _
				"    HPP_prod_bi.month_qty_sales_bi," & _
				"    HPP_prod_bi.year_qty_sales_bi," & _
				"    pm.param_1 AS jobcode_persediaan," & _
				"    pm.param_13 AS jobcode_intraco_HO," & _
				"    HPP_prod_bi.qty_sales_bi," & _
				"    hpp_penjualan_mtd.price_sales_bi," & _
				"    ROUND ( (-1 * (HPP_prod_bi.qty_sales_bi * hpp_penjualan_mtd.price_sales_bi)), 2) AS total_amount_sales_bi" & _
				" FROM (  SELECT fcba," & _
				"              product_code," & _
				"              month_qty_sales_bi," & _
				"              year_qty_sales_bi," & _
				"              SUM (qty_sales_bi) AS qty_sales_bi_before_adj," & _
				"              SUM (qty_adjust_bi) AS qty_adjust_bi," & _
				"              (SUM (qty_sales_bi) + SUM (qty_adjust_bi)) AS qty_sales_bi" & _
				"         FROM (  SELECT iplasprod.weighbridgeticket.fcba," & _
				"                        iplasprod.weighbridgeticket.product_productcode AS product_code," & _
				"                        MAX ( TO_NUMBER ( TO_CHAR (iplasprod.weighbridgeticket.dateout, 'mm')))" & _
				"                           AS month_qty_sales_bi," & _
				"                        MAX ( TO_NUMBER ( TO_CHAR (iplasprod.weighbridgeticket.dateout, 'yyyy')))" & _
				"                           AS year_qty_sales_bi," & _
				"                          SUM (iplasprod.weighbridgeticket.weightout)" & _
				"                        - SUM (iplasprod.weighbridgeticket.weightin)" & _
				"                           AS qty_sales_bi," & _
				"                        0 AS qty_adjust_bi" & _
				"                   FROM iplasprod.weighbridgeticket" & _
				"                  WHERE     iplasprod.weighbridgeticket.fcba = '"& Session("fcba") &"'" & _
				"                        AND TO_NUMBER ( TO_CHAR (iplasprod.weighbridgeticket.dateout, 'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
				"                        AND TO_NUMBER ( TO_CHAR (iplasprod.weighbridgeticket.dateout, 'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _
				"                        AND iplasprod.weighbridgeticket.product_productcode IN ('CPO', 'KER')" & _
				"               GROUP BY iplasprod.weighbridgeticket.fcba," & _
				"                        iplasprod.weighbridgeticket.product_productcode" & _
				"               UNION ALL" & _
				"                 SELECT iplasprod.mill_production.fcba AS fcba," & _
				"                        iplasprod.mill_production.product_code," & _
				"                        MAX ( TO_NUMBER ( TO_CHAR (iplasprod.mill_production.tdate, 'mm')))month_prod_bi," & _
				"                        MAX ( TO_NUMBER ( TO_CHAR (iplasprod.mill_production.tdate, 'yyyy')))year_prod_bi," & _
				"                        0 AS qty_sales_bi," & _
				"                        SUM (iplasprod.mill_production.qty_adjust) * -1 AS qty_adjust_bi" & _
				"                   FROM iplasprod.mill_production" & _
				"                  WHERE     iplasprod.mill_production.fcba = '"& Session("fcba") &"'" & _
				"                        AND TO_NUMBER ( TO_CHAR (iplasprod.mill_production.tdate, 'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
				"                        AND TO_NUMBER ( TO_CHAR (iplasprod.mill_production.tdate, 'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _
				"                        AND iplasprod.mill_production.product_code IN ('CPO', 'KER')" & _
				"               GROUP BY iplasprod.mill_production.fcba," & _
				"                        iplasprod.mill_production.product_code)" & _
				"     GROUP BY fcba," & _
				"              product_code," & _
				"              month_qty_sales_bi," & _
				"              year_qty_sales_bi) HPP_prod_bi" & _
				"    INNER JOIN" & _
				"    (SELECT fcba," & _
				"            product_code," & _
				"            month_saldo_awal," & _
				"            year_saldo_awal," & _
				"            saldo_qty_awal," & _
				"            saldo_amount_awal," & _
				"            Price_saldo_awal," & _
				"            month_prod_sbi," & _
				"            year_prod_sbi," & _
				"            qty_prod_sbi," & _
				"            last_month_qty_prod_sbi," & _
				"            amount_prod_sbi," & _
				"            last_month_amount_prod_sbi," & _
				"            Price_prod_sbi," & _
				"            month_qty_sales_sbi," & _
				"            year_qty_sales_sbi," & _
				"            qty_sales_sbi," & _
				"            last_month_qty_sales_sbi," & _
				"            (qty_sales_sbi - last_month_qty_sales_sbi) AS qty_sales_bi," & _
				"            HPP_Qty_level_1," & _
				"            HPP_Amount_level_1," & _
				"            HPP_Price_level_1," & _
				"            HPP_Qty_level_2," & _
				"            HPP_Amount_level_2," & _
				"            HPP_Price_level_2," & _
				"            amount_sales_sbi," & _
				"            last_month_amount_sales_sbi," & _
				"            (amount_sales_sbi - last_month_amount_sales_sbi) AS amount_sales_bi," & _
				"            price_sales_sbi," & _
				"            CASE" & _
				"               WHEN HPP_Qty_level_2 = 0" & _
				"               THEN" & _
				"                  price_saldo_awal" & _
				"               ELSE" & _
				"                  ( (amount_sales_sbi - last_month_amount_sales_sbi) / (qty_sales_sbi - last_month_qty_sales_sbi))" & _
				"            END price_sales_bi," & _
				"            saldo_Qty_akhir," & _
				"            saldo_amount_akhir," & _
				"            price_saldo_akhir" & _
				"       FROM (SELECT fcba," & _
				"                    product_code," & _
				"                    month_saldo_awal," & _
				"                    year_saldo_awal," & _
				"                    saldo_qty_awal," & _
				"                    saldo_amount_awal," & _
				"                    Price_saldo_awal," & _
				"                    month_prod_sbi," & _
				"                    year_prod_sbi," & _
				"                    qty_prod_sbi," & _
				"                    NVL ( IPLASPROD.GetLastMonthqtysbifrommidyear ( product_code, fcba, LAST_DAY ( ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1))), 0) " & _
				"					 last_month_qty_prod_sbi," & _
				"                    amount_prod_sbi," & _
				"                    NVL ( IPLASPROD.GetLastMonthamountprodsbimid ( product_code, fcba, LAST_DAY ( ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1))), 0) " & _
				"					 last_month_amount_prod_sbi," & _
				"                    Price_prod_sbi," & _
				"                    month_qty_sales_sbi," & _
				"                    year_qty_sales_sbi," & _
				"                    qty_sales_sbi," & _
				"                    NVL ( IPLASPROD.GetLastMonthqtysalessbimidyear ( product_code, fcba, LAST_DAY ( ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1))), 0) " & _
				"					 last_month_qty_sales_sbi," & _
				"                    HPP_Qty_level_1," & _
				"                    HPP_Amount_level_1," & _
				"                    HPP_Price_level_1," & _
				"                    HPP_Qty_level_2," & _
				"                    HPP_Amount_level_2," & _
				"                    HPP_Price_level_2," & _
				"                    amount_sales_sbi," & _
				"                    NVL ( IPLASPROD.GetLastMonthamountsalessbimid ( product_code, fcba, LAST_DAY ( ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1))), 0) " & _
				"					 last_month_amount_sales_sbi," & _
				"                    price_sales_sbi," & _
				"                    saldo_Qty_akhir," & _
				"                    saldo_amount_akhir," & _
				"                    price_saldo_akhir" & _
				"               FROM (SELECT fcba," & _
				"                            product_code," & _
				"                            month_saldo_awal," & _
				"                            year_saldo_awal," & _
				"                            saldo_qty_awal," & _
				"                            saldo_amount_awal," & _
				"                            Price_saldo_awal," & _
				"                            month_prod_sbi," & _
				"                            year_prod_sbi," & _
				"                            qty_prod_sbi," & _
				"                            amount_prod_sbi," & _
				"                            Price_prod_sbi," & _
				"                            month_qty_sales_sbi," & _
				"                            year_qty_sales_sbi," & _
				"                            qty_sales_sbi," & _
				"                            HPP_Qty_level_1," & _
				"                            HPP_Amount_level_1," & _
				"                            HPP_Price_level_1," & _
				"                            HPP_Qty_level_2," & _
				"                            HPP_Amount_level_2," & _
				"                            HPP_Price_level_2," & _
				"                            amount_sales_sbi," & _
				"                            price_sales_sbi," & _
				"                            saldo_Qty_akhir," & _
				"                            saldo_amount_akhir," & _
				"                            price_saldo_akhir" & _
				"                       FROM (SELECT fcba," & _
				"                                    product_code," & _
				"                                    month_saldo_awal," & _
				"                                    year_saldo_awal," & _
				"                                    saldo_qty_awal," & _
				"                                    saldo_amount_awal," & _
				"                                    Price_saldo_awal," & _
				"                                    month_prod_sbi," & _
				"                                    year_prod_sbi," & _
				"                                    qty_prod_sbi," & _
				"                                    amount_prod_sbi," & _
				"                                    Price_prod_sbi," & _
				"                                    month_qty_sales_sbi," & _
				"                                    year_qty_sales_sbi," & _
				"                                    qty_sales_sbi," & _
				"                                    HPP_Qty_level_1," & _
				"                                    HPP_Amount_level_1," & _
				"                                    HPP_Price_level_1," & _
				"                                    HPP_Qty_level_2," & _
				"                                    HPP_Amount_level_2," & _
				"                                    HPP_Price_level_2," & _
				"                                    (  HPP_Amount_level_1" & _
				"                                     + HPP_Amount_level_2)" & _
				"                                       AS amount_sales_sbi," & _
				"                                    DECODE (" & _
				"                                       qty_sales_sbi," & _
				"                                       0, 0," & _
				"                                       (  (  HPP_Amount_level_1" & _
				"                                           + HPP_Amount_level_2)" & _
				"                                        / qty_sales_sbi))" & _
				"                                       AS price_sales_sbi," & _
				"                                    (  saldo_qty_awal" & _
				"                                     + qty_prod_sbi" & _
				"                                     - qty_sales_sbi)" & _
				"                                       AS saldo_Qty_akhir," & _
				"                                    (  saldo_amount_awal" & _
				"                                     + amount_prod_sbi" & _
				"                                     - HPP_Amount_level_1" & _
				"                                     - HPP_Amount_level_2)" & _
				"                                       AS saldo_amount_akhir," & _
				"                                    (  (  saldo_amount_awal" & _
				"                                        + amount_prod_sbi" & _
				"                                        - HPP_Amount_level_1" & _
				"                                        - HPP_Amount_level_2)" & _
				"                                     / (  saldo_qty_awal" & _
				"                                        + qty_prod_sbi" & _
				"                                        - qty_sales_sbi))" & _
				"                                       AS price_saldo_akhir" & _
				"                               FROM (SELECT fcba," & _
				"                                            product_code," & _
				"                                            month_saldo_awal," & _
				"                                            year_saldo_awal," & _
				"                                            saldo_qty_awal," & _
				"                                            saldo_amount_awal," & _
				"                                            Price_saldo_awal," & _
				"                                            month_prod_sbi," & _
				"                                            year_prod_sbi," & _
				"                                            qty_prod_sbi," & _
				"                                            amount_prod_sbi," & _
				"                                            Price_prod_sbi," & _
				"                                            month_qty_sales_sbi," & _
				"                                            year_qty_sales_sbi," & _
				"                                            qty_sales_sbi," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  saldo_qty_awal" & _
				"                                               ELSE" & _
				"                                                  qty_sales_sbi" & _
				"                                            END" & _
				"                                               AS HPP_Qty_level_1," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  Price_saldo_awal" & _
				"                                               ELSE" & _
				"                                                  Price_saldo_awal" & _
				"                                            END" & _
				"                                               AS HPP_Price_level_1," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  (  saldo_qty_awal" & _
				"                                                   * Price_saldo_awal)" & _
				"                                               ELSE" & _
				"                                                  (  (  qty_sales_sbi" & _
				"                                                      / saldo_qty_awal)" & _
				"                                                   * saldo_amount_awal)" & _
				"                                            END" & _
				"                                               AS HPP_Amount_level_1," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  (  qty_sales_sbi" & _
				"                                                   - saldo_qty_awal)" & _
				"                                               ELSE 0" & _
				"                                            END" & _
				"                                               AS HPP_Qty_level_2," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  Price_prod_sbi" & _
				"                                               ELSE 0" & _
				"                                            END" & _
				"                                               AS HPP_Price_level_2," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  (  (  qty_sales_sbi" & _
				"                                                      - saldo_qty_awal)" & _
				"                                                   * Price_prod_sbi)" & _
				"                                               ELSE 0" & _
				"                                            END" & _
				"                                               AS HPP_Amount_level_2" & _
				"                                       FROM (  SELECT fcba," & _
				"                                                      product_code," & _
				"                                                      MAX (month_saldo_awal)" & _
				"                                                         AS month_saldo_awal," & _
				"                                                      MAX (year_saldo_awal)" & _
				"                                                         AS year_saldo_awal," & _
				"                                                      MAX (saldo_qty_awal)" & _
				"                                                         AS saldo_qty_awal," & _
				"                                                      MAX (" & _
				"                                                         saldo_amount_awal)" & _
				"                                                         AS saldo_amount_awal," & _
				"                                                      MAX (Price_saldo_awal)" & _
				"                                                         AS Price_saldo_awal," & _
				"                                                      MAX (month_prod_sbi)" & _
				"                                                         AS month_prod_sbi," & _
				"                                                      MAX (year_prod_sbi)" & _
				"                                                         AS year_prod_sbi," & _
				"                                                      MAX (qty_prod_sbi)" & _
				"                                                         AS qty_prod_sbi," & _
				"                                                      MAX (amount_prod_sbi)" & _
				"                                                         AS amount_prod_sbi," & _
				"                                                      MAX (Price_prod_sbi)" & _
				"                                                         AS Price_prod_sbi," & _
				"                                                      MAX (" & _
				"                                                         month_qty_sales_sbi)" & _
				"                                                         AS month_qty_sales_sbi," & _
				"                                                      MAX (" & _
				"                                                         year_qty_sales_sbi)" & _
				"                                                         AS year_qty_sales_sbi," & _
				"                                                      MAX (qty_sales_sbi)" & _
				"                                                         AS qty_sales_sbi" & _
				"                                                 FROM (SELECT fcba," & _
				"                                                              product_code," & _
				"                                                              month_saldo_awal," & _
				"                                                              year_saldo_awal," & _
				"                                                              qty_cpo_prod_awal" & _
				"                                                                 AS saldo_qty_awal," & _
				"                                                              prop_amount_cpo_prod_awal" & _
				"                                                                 AS saldo_amount_awal," & _
				"                                                              (  prop_amount_cpo_prod_awal" & _
				"                                                               / qty_cpo_prod_awal)" & _
				"                                                                 AS Price_saldo_awal," & _
				"                                                              0 AS month_prod_sbi," & _
				"                                                              0 AS year_prod_sbi," & _
				"                                                              0 AS qty_prod_sbi," & _
				"                                                              0 AS amount_prod_sbi," & _
				"                                                              0 AS Price_prod_sbi," & _
				"                                                              0 AS month_qty_sales_sbi," & _
				"                                                              0 AS year_qty_sales_sbi," & _
				"                                                              0 AS qty_sales_sbi" & _
				"                                                         FROM (SELECT fcba," & _
				"                                                                      'CPO' AS product_code," & _
				"                                                                      month_saldo_awal," & _
				"                                                                      year_saldo_awal," & _
				"                                                                      qty_cpo_prod_awal," & _
				"                                                                      15758662530.5 AS prop_amount_cpo_prod_awal" & _
				"                                                                 FROM (SELECT qty_awal.fcba," & _
				"                                                                              qty_awal.month_saldo_awal," & _
				"                                                                              qty_awal.year_saldo_awal," & _
				"                                                                              qty_awal.qty_cpo_prod_awal," & _
				"                                                                              qty_awal.qty_ker_prod_awal," & _
				"                                                                              biaya_awal.saldo_amount_awal," & _
				"                                                                              (  biaya_awal.saldo_amount_awal" & _
				"                                                                               / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                  + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                     / 2)))" & _
				"                                                                                 AS prop_cpo_prod_awal," & _
				"                                                                              (  (  biaya_awal.saldo_amount_awal" & _
				"                                                                                  / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                     + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                        / 2)))" & _
				"                                                                               / 2)" & _
				"                                                                                 AS prop_ker_prod_awal," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_awal.qty_cpo_prod_awal" & _
				"                                                                                 * ( (  biaya_awal.saldo_amount_awal" & _
				"                                                                                      / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                         + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                            / 2))))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_cpo_prod_awal," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_awal.qty_ker_prod_awal" & _
				"                                                                                 * ( (  (  biaya_awal.saldo_amount_awal" & _
				"                                                                                         / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                            + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                               / 2)))" & _
				"                                                                                      / 2))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_ker_prod_awal" & _
				"                                                                         FROM (  SELECT fcba," & _
				"                                                                                        month_saldo_awal," & _
				"                                                                                        year_saldo_awal," & _
				"                                                                                        SUM ( DECODE ( product_code, 'CPO', saldo_qty_awal, 0))" & _
				"                                                                                           AS qty_cpo_prod_awal," & _
				"                                                                                        SUM ( DECODE ( product_code, 'KER', saldo_qty_awal, 0))" & _
				"                                                                                           AS qty_ker_prod_awal" & _
				"                                                                                   FROM (SELECT iplasprod.mill_summary_tab.fcba," & _
				"                                                                                                CASE" & _
				"                                                                                                   WHEN iplasprod.mill_summary_tab.codedesc =" & _
				"                                                                                                           'CPO003'" & _
				"                                                                                                   THEN" & _
				"                                                                                                      'CPO'" & _
				"                                                                                                   ELSE" & _
				"                                                                                                      'KER'" & _
				"                                                                                                END" & _
				"                                                                                                   AS product_code," & _
				"                                                                                                TO_NUMBER ( TO_CHAR ( iplasprod.mill_summary_tab.ondate, 'mm'))" & _
				"                                                                                                   AS month_saldo_awal," & _
				"                                                                                                TO_NUMBER ( TO_CHAR ( iplasprod.mill_summary_tab.ondate, 'yyyy'))" & _
				"                                                                                                   AS year_saldo_awal," & _
				"                                                                                                NVL ( iplasprod.mill_summary_tab.todate, 0)" & _
				"                                                                                                   AS saldo_qty_awal" & _
				"                                                                                           FROM iplasprod.mill_summary_tab" & _
				"                                                                                          WHERE     iplasprod.mill_summary_tab.fcba =" & _
				"                                                                                                       '"& Session("fcba") &"'" & _
				"                                                                                                AND iplasprod.mill_summary_tab.codedesc IN" & _
				"                                                                                                       ('CPO003'," & _
				"                                                                                                        'PK003')" & _
				"                                                                                                AND iplasprod.mill_summary_tab.ondate = " & _
				"																								 TO_DATE ( '07/2019', 'mm/yyyy') - 1)" & _
				"                                                                               GROUP BY fcba," & _
				"                                                                                        month_saldo_awal," & _
				"                                                                                        year_saldo_awal) qty_awal" & _
				"                                                                              INNER JOIN" & _
				"                                                                              (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                                                           AS fcba," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'mm')))" & _
				"                                                                                           AS month_saldo_awal," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'yyyy')))" & _
				"                                                                                           AS year_saldo_awal," & _
				"                                                                                        SUM (" & _
				"                                                                                           iplasprod.cost_activity.amount)" & _
				"                                                                                           AS saldo_amount_awal" & _
				"                                                                                   FROM iplasprod.cost_activity" & _
				"                                                                                  WHERE     cost_activity.fcba =" & _
				"                                                                                               '"& Session("fcba") &"'" & _
				"                                                                                        AND iplasprod.cost_activity.tdate < TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                        AND cost_activity.jobcode IN" & _
				"                                                                                               (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                                                  FROM iplasprod.parameterdetail" & _
				"                                                                                                 WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                                              'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                                                       AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                                              'CPO'" & _
				"                                                                                                       AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                                              '"& Session("fcba") &"')" & _
				"                                                                               GROUP BY iplasprod.cost_activity.fcba) biaya_awal" & _
				"                                                                                 ON (    qty_awal.fcba =" & _
				"                                                                                            biaya_awal.fcba" & _
				"                                                                                     AND qty_awal.month_saldo_awal =" & _
				"                                                                                            biaya_awal.month_saldo_awal" & _
				"                                                                                     AND qty_awal.year_saldo_awal =" & _
				"                                                                                            biaya_awal.year_saldo_awal))" & _
				"                                                               UNION ALL" & _
				"                                                               SELECT fcba," & _
				"                                                                      'KER'" & _
				"                                                                         AS product_code," & _
				"                                                                      month_saldo_awal," & _
				"                                                                      year_saldo_awal," & _
				"                                                                      qty_ker_prod_awal," & _
				" 																	   1051395609.64 AS prop_amount_ker_prod_awal" & _
				"                                                                 FROM (SELECT qty_awal.fcba," & _
				"                                                                              qty_awal.month_saldo_awal," & _
				"                                                                              qty_awal.year_saldo_awal," & _
				"                                                                              qty_awal.qty_cpo_prod_awal," & _
				"                                                                              qty_awal.qty_ker_prod_awal," & _
				"                                                                              biaya_awal.saldo_amount_awal," & _
				"                                                                              (  biaya_awal.saldo_amount_awal" & _
				"                                                                               / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                  + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                     / 2)))" & _
				"                                                                                 AS prop_cpo_prod_awal," & _
				"                                                                              (  (  biaya_awal.saldo_amount_awal" & _
				"                                                                                  / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                     + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                        / 2)))" & _
				"                                                                               / 2)" & _
				"                                                                                 AS prop_ker_prod_awal," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_awal.qty_cpo_prod_awal" & _
				"                                                                                 * ( (  biaya_awal.saldo_amount_awal" & _
				"                                                                                      / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                         + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                            / 2))))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_cpo_prod_awal," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_awal.qty_ker_prod_awal" & _
				"                                                                                 * ( (  (  biaya_awal.saldo_amount_awal" & _
				"                                                                                         / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                            + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                               / 2)))" & _
				"                                                                                      / 2))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_ker_prod_awal" & _
				"                                                                         FROM (  SELECT fcba," & _
				"                                                                                        month_saldo_awal," & _
				"                                                                                        year_saldo_awal," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'CPO', saldo_qty_awal," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_cpo_prod_awal," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'KER', saldo_qty_awal," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_ker_prod_awal" & _
				"                                                                                   FROM (SELECT iplasprod.mill_summary_tab.fcba," & _
				"                                                                                                CASE" & _
				"                                                                                                   WHEN iplasprod.mill_summary_tab.codedesc =" & _
				"                                                                                                           'CPO003'" & _
				"                                                                                                   THEN" & _
				"                                                                                                      'CPO'" & _
				"                                                                                                   ELSE" & _
				"                                                                                                      'KER'" & _
				"                                                                                                END" & _
				"                                                                                                   AS product_code," & _
				"                                                                                                TO_NUMBER ( TO_CHAR ( iplasprod.mill_summary_tab.ondate, 'mm'))" & _
				"                                                                                                   AS month_saldo_awal," & _
				"                                                                                                TO_NUMBER ( TO_CHAR ( iplasprod.mill_summary_tab.ondate, 'yyyy'))" & _
				"                                                                                                   AS year_saldo_awal," & _
				"                                                                                                NVL (" & _
				"                                                                                                   iplasprod.mill_summary_tab.todate," & _
				"                                                                                                   0)" & _
				"                                                                                                   AS saldo_qty_awal" & _
				"                                                                                           FROM iplasprod.mill_summary_tab" & _
				"                                                                                          WHERE     iplasprod.mill_summary_tab.fcba = '"& Session("fcba") &"'" & _
				"                                                                                                AND iplasprod.mill_summary_tab.codedesc IN" & _
				"                                                                                                       ('CPO003'," & _
				"                                                                                                        'PK003')" & _
				"                                                                                                AND iplasprod.mill_summary_tab.ondate =" & _
				" 																								 TO_DATE ( '07/2019', 'mm/yyyy') - 1)" & _
				"                                                                               GROUP BY fcba," & _
				"                                                                                        month_saldo_awal," & _
				"                                                                                        year_saldo_awal) qty_awal" & _
				"                                                                              INNER JOIN" & _
				"                                                                              (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                                                           AS fcba," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'mm')))" & _
				"                                                                                           AS month_saldo_awal," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'yyyy')))" & _
				"                                                                                           AS year_saldo_awal," & _
				"                                                                                        SUM (" & _
				"                                                                                           iplasprod.cost_activity.amount)" & _
				"                                                                                           AS saldo_amount_awal" & _
				"                                                                                   FROM iplasprod.cost_activity" & _
				"                                                                                  WHERE     cost_activity.fcba =" & _
				"                                                                                               '"& Session("fcba") &"'" & _
				"                                                                                        AND iplasprod.cost_activity.tdate <" & _
				"                                                                                               TO_DATE (" & _
				"                                                                                                  '07/1/2019'," & _
				"                                                                                                  'mm/dd/yyyy')" & _
				"                                                                                        AND cost_activity.jobcode IN" & _
				"                                                                                               (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                                                  FROM iplasprod.parameterdetail" & _
				"                                                                                                 WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                                              'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                                                       AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                                              'KER'" & _
				"                                                                                                       AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                                              '"& Session("fcba") &"')" & _
				"                                                                               GROUP BY iplasprod.cost_activity.fcba) biaya_awal" & _
				"                                                                                 ON (    qty_awal.fcba =" & _
				"                                                                                            biaya_awal.fcba" & _
				"                                                                                     AND qty_awal.month_saldo_awal =" & _
				"                                                                                            biaya_awal.month_saldo_awal" & _
				"                                                                                     AND qty_awal.year_saldo_awal =" & _
				"                                                                                            biaya_awal.year_saldo_awal)))" & _
				"                                                        WHERE product_code IN" & _
				"                                                                 ('CPO'," & _
				"                                                                  'KER')" & _
				"                                                       UNION ALL" & _
				"                                                       SELECT fcba," & _
				"                                                              product_code," & _
				"                                                              0 AS month_saldo_awal," & _
				"                                                              0 AS year_saldo_awal," & _
				"                                                              0 AS saldo_qty_awal," & _
				"                                                              0 AS saldo_amount_awal," & _
				"                                                              0 AS Price_saldo_awal," & _
				"                                                              month_prod_sbi," & _
				"                                                              year_prod_sbi," & _
				"                                                              qty_cpo_prod_sbi" & _
				"                                                                 AS qty_prod_sbi," & _
				"                                                              prop_amount_cpo_prod_sbi" & _
				"                                                                 AS amount_prod_sbi," & _
				"                                                              (  prop_amount_cpo_prod_sbi" & _
				"                                                               / qty_cpo_prod_sbi)" & _
				"                                                                 AS Price_prod_sbi," & _
				"                                                              0 AS month_qty_sales_sbi," & _
				"                                                              0 AS year_qty_sales_sbi," & _
				"                                                              0 AS qty_sales_sbi" & _
				"                                                         FROM (SELECT fcba," & _
				"                                                                      'CPO'" & _
				"                                                                         AS product_code," & _
				"                                                                      month_prod_sbi," & _
				"                                                                      year_prod_sbi," & _
				"                                                                      qty_cpo_prod_sbi," & _
				"                                                                      prop_amount_cpo_prod_sbi" & _
				"                                                                 FROM (SELECT qty_sbi.fcba," & _
				"                                                                              qty_sbi.month_prod_sbi," & _
				"                                                                              qty_sbi.year_prod_sbi," & _
				"                                                                              qty_sbi.qty_cpo_prod_sbi," & _
				"                                                                              qty_sbi.qty_ker_prod_sbi," & _
				"                                                                              biaya_sbi.amount_prod_sbi," & _
				"                                                                              (  biaya_sbi.amount_prod_sbi" & _
				"                                                                               / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                  + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                     / 2)))" & _
				"                                                                                 AS prop_cpo_prod_sbi," & _
				"                                                                              (  (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                  / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                     + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                        / 2)))" & _
				"                                                                               / 2)" & _
				"                                                                                 AS prop_ker_prod_sbi," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                 * ( (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                      / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                         + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                            / 2))))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_cpo_prod_sbi," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                 * ( (  (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                         / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                            + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                               / 2)))" & _
				"                                                                                      / 2))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_ker_prod_sbi" & _
				"                                                                         FROM (  SELECT fcba," & _
				"                                                                                        month_prod_sbi," & _
				"                                                                                        year_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'CPO', qty_prod_sbi," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_cpo_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'KER', qty_prod_sbi," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_ker_prod_sbi" & _
				"                                                                                   FROM (  SELECT iplasprod.mill_production.fcba" & _
				"                                                                                                     AS fcba," & _
				"                                                                                                  iplasprod.mill_production.product_code," & _
				"                                                                                                  MAX (" & _
				"                                                                                                     TO_NUMBER (" & _
				"                                                                                                        TO_CHAR (" & _
				"                                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                                           'mm')))" & _
				"                                                                                                     AS month_prod_sbi," & _
				"                                                                                                  MAX (" & _
				"                                                                                                     TO_NUMBER (" & _
				"                                                                                                        TO_CHAR (" & _
				"                                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                                           'yyyy')))" & _
				"                                                                                                     AS year_prod_sbi," & _
				"                                                                                                  SUM (" & _
				"                                                                                                     iplasprod.mill_production.production_today)" & _
				"                                                                                                     AS qty_prod_sbi" & _
				"                                                                                             FROM iplasprod.mill_production" & _
				"                                                                                            WHERE     iplasprod.mill_production.fcba =" & _
				"                                                                                                         '"& Session("fcba") &"'" & _
				"                                                                                                  AND (iplasprod.mill_production.tdate BETWEEN TO_DATE (" & _
				"                                                                                                                                                  '07/1/2019'," & _
				"                                                                                                                                                  'mm/dd/yyyy')" & _
				"                                                                                                                                           AND LAST_DAY (" & _
				"                                                                                                                                                  TO_DATE (" & _
				"                                                                                                                                                     '"& LastPeriod &"'," & _
				"                                                                                                                                                     'mm/dd/yyyy')))" & _
				"                                                                                                  AND iplasprod.mill_production.product_code IN" & _
				"                                                                                                         ('CPO'," & _
				"                                                                                                          'KER')" & _
				"                                                                                         GROUP BY iplasprod.mill_production.fcba," & _
				"                                                                                                  iplasprod.mill_production.product_code)" & _
				"                                                                               GROUP BY fcba," & _
				"                                                                                        month_prod_sbi," & _
				"                                                                                        year_prod_sbi) qty_sbi" & _
				"                                                                              INNER JOIN" & _
				"                                                                              (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                                                           AS fcba," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'mm')))" & _
				"                                                                                           AS month_prod_sbi," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'yyyy')))" & _
				"                                                                                           AS year_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           iplasprod.cost_activity.amount)" & _
				"                                                                                           AS amount_prod_sbi" & _
				"                                                                                   FROM iplasprod.cost_activity" & _
				"                                                                                  WHERE     cost_activity.fcba =" & _
				"                                                                                               '"& Session("fcba") &"'" & _
				"                                                                                        AND (iplasprod.cost_activity.tdate BETWEEN TO_DATE (" & _
				"                                                                                                                                      '07/1/2019'," & _
				"                                                                                                                                      'mm/dd/yyyy')" & _
				"                                                                                                                               AND LAST_DAY (" & _
				"                                                                                                                                      TO_DATE (" & _
				"                                                                                                                                         '"& LastPeriod &"'," & _
				"                                                                                                                                         'mm/dd/yyyy')))" & _
				"                                                                                        AND cost_activity.jobcode IN" & _
				"                                                                                               (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                                                  FROM iplasprod.parameterdetail" & _
				"                                                                                                 WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                                              'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                                                       AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                                              'CPO'" & _
				"                                                                                                       AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                                              '"& Session("fcba") &"')" & _
				"                                                                               GROUP BY iplasprod.cost_activity.fcba) biaya_sbi" & _
				"                                                                                 ON (    qty_sbi.fcba =" & _
				"                                                                                            biaya_sbi.fcba" & _
				"                                                                                     AND qty_sbi.month_prod_sbi =" & _
				"                                                                                            biaya_sbi.month_prod_sbi" & _
				"                                                                                     AND qty_sbi.year_prod_sbi =" & _
				"                                                                                            biaya_sbi.year_prod_sbi))" & _
				"                                                               UNION ALL" & _
				"                                                               SELECT fcba," & _
				"                                                                      'KER'" & _
				"                                                                         AS product_code," & _
				"                                                                      month_prod_sbi," & _
				"                                                                      year_prod_sbi," & _
				"                                                                      qty_ker_prod_sbi," & _
				"                                                                      prop_amount_ker_prod_sbi" & _
				"                                                                 FROM (SELECT qty_sbi.fcba," & _
				"                                                                              qty_sbi.month_prod_sbi," & _
				"                                                                              qty_sbi.year_prod_sbi," & _
				"                                                                              qty_sbi.qty_cpo_prod_sbi," & _
				"                                                                              qty_sbi.qty_ker_prod_sbi," & _
				"                                                                              biaya_sbi.amount_prod_sbi," & _
				"                                                                              (  biaya_sbi.amount_prod_sbi" & _
				"                                                                               / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                  + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                     / 2)))" & _
				"                                                                                 AS prop_cpo_prod_sbi," & _
				"                                                                              (  (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                  / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                     + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                        / 2)))" & _
				"                                                                               / 2)" & _
				"                                                                                 AS prop_ker_prod_sbi," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                 * ( (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                      / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                         + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                            / 2))))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_cpo_prod_sbi," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                 * ( (  (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                         / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                            + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                               / 2)))" & _
				"                                                                                      / 2))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_ker_prod_sbi" & _
				"                                                                         FROM (  SELECT fcba," & _
				"                                                                                        month_prod_sbi," & _
				"                                                                                        year_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'CPO', qty_prod_sbi," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_cpo_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'KER', qty_prod_sbi," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_ker_prod_sbi" & _
				"                                                                                   FROM (  SELECT iplasprod.mill_production.fcba" & _
				"                                                                                                     AS fcba," & _
				"                                                                                                  iplasprod.mill_production.product_code," & _
				"                                                                                                  MAX (" & _
				"                                                                                                     TO_NUMBER (" & _
				"                                                                                                        TO_CHAR (" & _
				"                                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                                           'mm')))" & _
				"                                                                                                     AS month_prod_sbi," & _
				"                                                                                                  MAX (" & _
				"                                                                                                     TO_NUMBER (" & _
				"                                                                                                        TO_CHAR (" & _
				"                                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                                           'yyyy')))" & _
				"                                                                                                     AS year_prod_sbi," & _
				"                                                                                                  SUM (" & _
				"                                                                                                     iplasprod.mill_production.production_today)" & _
				"                                                                                                     AS qty_prod_sbi" & _
				"                                                                                             FROM iplasprod.mill_production" & _
				"                                                                                            WHERE     iplasprod.mill_production.fcba = '"& Session("fcba") &"'" & _
				"                                                                                                  AND (iplasprod.mill_production.tdate BETWEEN" & _
				"																								   TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                                  AND LAST_DAY ( TO_DATE ( '"& LastPeriod &"', 'mm/dd/yyyy')))" & _
				"                                                                                                  AND iplasprod.mill_production.product_code IN" & _
				"                                                                                                         ('CPO'," & _
				"                                                                                                          'KER')" & _
				"                                                                                         GROUP BY iplasprod.mill_production.fcba," & _
				"                                                                                                  iplasprod.mill_production.product_code)" & _
				"                                                                               GROUP BY fcba," & _
				"                                                                                        month_prod_sbi," & _
				"                                                                                        year_prod_sbi) qty_sbi" & _
				"                                                                              INNER JOIN" & _
				"                                                                              (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                                                           AS fcba," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'mm')))" & _
				"                                                                                           AS month_prod_sbi," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'yyyy')))" & _
				"                                                                                           AS year_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           iplasprod.cost_activity.amount)" & _
				"                                                                                           AS amount_prod_sbi" & _
				"                                                                                   FROM iplasprod.cost_activity" & _
				"                                                                                  WHERE     cost_activity.fcba =" & _
				"                                                                                               '"& Session("fcba") &"'" & _
				"                                                                                        AND (iplasprod.cost_activity.tdate BETWEEN TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                        AND LAST_DAY ( TO_DATE ( '"& LastPeriod &"', 'mm/dd/yyyy')))" & _
				"                                                                                        AND cost_activity.jobcode IN" & _
				"                                                                                               (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                                                  FROM iplasprod.parameterdetail" & _
				"                                                                                                 WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                                              'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                                                       AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                                              'KER'" & _
				"                                                                                                       AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                                              '"& Session("fcba") &"')" & _
				"                                                                               GROUP BY iplasprod.cost_activity.fcba) biaya_sbi" & _
				"                                                                                 ON (    qty_sbi.fcba =" & _
				"                                                                                            biaya_sbi.fcba" & _
				"                                                                                     AND qty_sbi.month_prod_sbi =" & _
				"                                                                                            biaya_sbi.month_prod_sbi" & _
				"                                                                                     AND qty_sbi.year_prod_sbi =" & _
				"                                                                                            biaya_sbi.year_prod_sbi)))" & _
				"                                                        WHERE product_code IN" & _
				"                                                                 ('CPO'," & _
				"                                                                  'KER')" & _
				"                                                       UNION ALL" & _
				"                                                       SELECT fcba," & _
				"                                                              product_code," & _
				"                                                              0 AS month_saldo_awal," & _
				"                                                              0 AS year_saldo_awal," & _
				"                                                              0 AS saldo_qty_awal," & _
				"                                                              0 AS saldo_amount_awal," & _
				"                                                              0 AS Price_saldo_awal," & _
				"                                                              0 AS month_prod_sbi," & _
				"                                                              0 AS year_prod_sbi," & _
				"                                                              0 AS qty_prod_sbi," & _
				"                                                              0 AS amount_prod_sbi," & _
				"                                                              0 AS Price_prod_sbi," & _
				"                                                              month_qty_sales_sbi," & _
				"                                                              year_qty_sales_sbi," & _
				"                                                              qty_sales_sbi" & _
				"                                                         FROM (  SELECT fcba," & _
				"                                                                        product_code," & _
				"                                                                        month_qty_sales_sbi," & _
				"                                                                        year_qty_sales_sbi," & _
				"                                                                        SUM (" & _
				"                                                                           qty_sales_sbi)" & _
				"                                                                           AS qty_sales_sbi_before_adj," & _
				"                                                                        SUM (" & _
				"                                                                           qty_adjust_sbi)" & _
				"                                                                           AS qty_adjust_sbi," & _
				"                                                                        (  SUM (" & _
				"                                                                              qty_sales_sbi)" & _
				"                                                                         + SUM (" & _
				"                                                                              qty_adjust_sbi))" & _
				"                                                                           AS qty_sales_sbi" & _
				"                                                                   FROM (  SELECT iplasprod.weighbridgeticket.fcba," & _
				"                                                                                  iplasprod.weighbridgeticket.product_productcode" & _
				"                                                                                     AS product_code," & _
				"                                                                                  MAX (" & _
				"                                                                                     TO_NUMBER (" & _
				"                                                                                        TO_CHAR (" & _
				"                                                                                           iplasprod.weighbridgeticket.dateout," & _
				"                                                                                           'mm')))" & _
				"                                                                                     AS month_qty_sales_sbi," & _
				"                                                                                  MAX (" & _
				"                                                                                     TO_NUMBER (" & _
				"                                                                                        TO_CHAR (" & _
				"                                                                                           iplasprod.weighbridgeticket.dateout," & _
				"                                                                                           'yyyy')))" & _
				"                                                                                     AS year_qty_sales_sbi," & _
				"                                                                                    SUM (" & _
				"                                                                                       iplasprod.weighbridgeticket.weightout)" & _
				"                                                                                  - SUM (" & _
				"                                                                                       iplasprod.weighbridgeticket.weightin)" & _
				"                                                                                     AS qty_sales_sbi, 0" & _
				"                                                                                     AS qty_adjust_sbi" & _
				"                                                                             FROM iplasprod.weighbridgeticket" & _
				"                                                                            WHERE     iplasprod.weighbridgeticket.fcba =" & _
				"                                                                                         '"& Session("fcba") &"'" & _
				"                                                                                  AND (iplasprod.weighbridgeticket.dateout BETWEEN TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                  AND LAST_DAY ( TO_DATE ( '"& LastPeriod &"', 'mm/dd/yyyy')))" & _
				"                                                                                  AND iplasprod.weighbridgeticket.product_productcode IN" & _
				"                                                                                         ('CPO'," & _
				"                                                                                          'KER')" & _
				"                                                                         GROUP BY iplasprod.weighbridgeticket.fcba," & _
				"                                                                                  iplasprod.weighbridgeticket.product_productcode" & _
				"                                                                         UNION ALL" & _
				"                                                                           SELECT iplasprod.mill_production.fcba" & _
				"                                                                                     AS fcba," & _
				"                                                                                  iplasprod.mill_production.product_code," & _
				"                                                                                  MAX (" & _
				"                                                                                     TO_NUMBER (" & _
				"                                                                                        TO_CHAR (" & _
				"                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                           'mm')))" & _
				"                                                                                     AS month_prod_sbi," & _
				"                                                                                  MAX (" & _
				"                                                                                     TO_NUMBER (" & _
				"                                                                                        TO_CHAR (" & _
				"                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                           'yyyy')))" & _
				"                                                                                     AS year_prod_sbi, 0" & _
				"                                                                                     AS qty_sales_sbi," & _
				"                                                                                    SUM (" & _
				"                                                                                       iplasprod.mill_production.qty_adjust)" & _
				"                                                                                  * -1" & _
				"                                                                                     AS qty_adjust_sbi" & _
				"                                                                             FROM iplasprod.mill_production" & _
				"                                                                            WHERE     iplasprod.mill_production.fcba =" & _
				"                                                                                         '"& Session("fcba") &"'" & _
				"                                                                                  AND (iplasprod.mill_production.tdate BETWEEN TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                  AND LAST_DAY ( TO_DATE ( '"& LastPeriod &"', 'mm/dd/yyyy')))" & _
				"                                                                                  AND iplasprod.mill_production.product_code IN" & _
				"                                                                                         ('CPO', 'KER')" & _
				"                                                                         GROUP BY iplasprod.mill_production.fcba," & _
				"                                                                                  iplasprod.mill_production.product_code)" & _
				"                                                               GROUP BY fcba," & _
				"                                                                        product_code," & _
				"                                                                        month_qty_sales_sbi," & _
				"                                                                        year_qty_sales_sbi))" & _
				"                                             GROUP BY fcba, product_code))))))" & _
				"    hpp_penjualan_mtd" & _
				"       ON (    HPP_prod_bi.fcba = hpp_penjualan_mtd.fcba" & _
				"           AND HPP_prod_bi.product_code =" & _
				"                  hpp_penjualan_mtd.product_code" & _
				"           AND HPP_prod_bi.month_qty_sales_bi =" & _
				"                  hpp_penjualan_mtd.month_prod_sbi" & _
				"           AND HPP_prod_bi.year_qty_sales_bi =" & _
				"                  hpp_penjualan_mtd.year_prod_sbi)" & _
				"    INNER JOIN iplasprod.parameter pm" & _
				"       ON (    HPP_prod_bi.fcba = pm.fcba" & _
				"           AND HPP_prod_bi.product_code = pm.fccode)" & _
				" WHERE pm.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				" )" & _
				" group by" & _
				" fcba," & _
				" month_qty_sales_bi," & _
				" year_qty_sales_bi," & _
				" jobcode_intraco_HO "

	' response.write qSalesDB
	Set rsSalesDB = Server.CreateObject("ADODB.Recordset")
	rsSalesDB.Open qSalesDB, connstr
	
	if not rsSalesDB.eof then
		Do While Not rsSalesDB.eof		  
			 
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityDBHPP="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & temp_datePeriodControl & "','mm/dd/yyyy'),'" & rsSalesDB("jobcode_intraco_HO") & "','SL','SL','Db-HARGAPOKOKPPENJUALAN'," & _
								 "'','CB','BS'," & _
								 "'" &  rsSalesDB("total_amount_sales_bi") & "','Db-HARGAPOKOKPPENJUALAN'," & "0,0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '' ,''," & _
								 " '',''," & _
								 " '','', " & _
								 " '','', " & _
								 " '', '','Db', " & _
								 " " & rsSalesDB("month_qty_sales_bi") & "," & rsSalesDB("year_qty_sales_bi") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "
			' response.write "<br>" & qCost_ActivityDBHPP
			 connstr.execute(qCost_ActivityDBHPP)
			 connStr.execute("Commit")

		rsSalesDB.MoveNext
	  loop
	  response.Write "Process Harga Pokok Penjualan Sales Debet Complete.... !!<br>"
	  response.Flush
	end if
	'-- End Insert Debet Harga Pokok Penjualan

	'-- Insert Credit Harga Pokok Penjualan
	qSalesCR =	" SELECT ROWNUM lineno," & _
				"    HPP_prod_bi.fcba," & _
				"    HPP_prod_bi.product_code," & _
				"    HPP_prod_bi.month_qty_sales_bi," & _
				"    HPP_prod_bi.year_qty_sales_bi," & _
				"    pm.param_1 AS jobcode_persediaan," & _
				"    pm.param_2 AS jobcode_hargapokokpenjualan," & _
				"    HPP_prod_bi.qty_sales_bi," & _
				"    hpp_penjualan_mtd.price_sales_bi," & _
				"    ROUND ( (-1 * (HPP_prod_bi.qty_sales_bi * hpp_penjualan_mtd.price_sales_bi)), 2) AS total_amount_sales_bi" & _
				" FROM (  SELECT fcba," & _
				"              product_code," & _
				"              month_qty_sales_bi," & _
				"              year_qty_sales_bi," & _
				"              SUM (qty_sales_bi) AS qty_sales_bi_before_adj," & _
				"              SUM (qty_adjust_bi) AS qty_adjust_bi," & _
				"              (SUM (qty_sales_bi) + SUM (qty_adjust_bi)) AS qty_sales_bi" & _
				"         FROM (  SELECT iplasprod.weighbridgeticket.fcba," & _
				"                        iplasprod.weighbridgeticket.product_productcode AS product_code," & _
				"                        MAX ( TO_NUMBER ( TO_CHAR (iplasprod.weighbridgeticket.dateout, 'mm')))" & _
				"                           AS month_qty_sales_bi," & _
				"                        MAX ( TO_NUMBER ( TO_CHAR (iplasprod.weighbridgeticket.dateout, 'yyyy')))" & _
				"                           AS year_qty_sales_bi," & _
				"                          SUM (iplasprod.weighbridgeticket.weightout)" & _
				"                        - SUM (iplasprod.weighbridgeticket.weightin)" & _
				"                           AS qty_sales_bi," & _
				"                        0 AS qty_adjust_bi" & _
				"                   FROM iplasprod.weighbridgeticket" & _
				"                  WHERE     iplasprod.weighbridgeticket.fcba = '"& Session("fcba") &"'" & _
				"                        AND TO_NUMBER ( TO_CHAR (iplasprod.weighbridgeticket.dateout, 'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
				"                        AND TO_NUMBER ( TO_CHAR (iplasprod.weighbridgeticket.dateout, 'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _
				"                        AND iplasprod.weighbridgeticket.product_productcode IN ('CPO', 'KER')" & _
				"               GROUP BY iplasprod.weighbridgeticket.fcba," & _
				"                        iplasprod.weighbridgeticket.product_productcode" & _
				"               UNION ALL" & _
				"                 SELECT iplasprod.mill_production.fcba AS fcba," & _
				"                        iplasprod.mill_production.product_code," & _
				"                        MAX ( TO_NUMBER ( TO_CHAR (iplasprod.mill_production.tdate, 'mm')))month_prod_bi," & _
				"                        MAX ( TO_NUMBER ( TO_CHAR (iplasprod.mill_production.tdate, 'yyyy')))year_prod_bi," & _
				"                        0 AS qty_sales_bi," & _
				"                        SUM (iplasprod.mill_production.qty_adjust) * -1 AS qty_adjust_bi" & _
				"                   FROM iplasprod.mill_production" & _
				"                  WHERE     iplasprod.mill_production.fcba = '"& Session("fcba") &"'" & _
				"                        AND TO_NUMBER ( TO_CHAR (iplasprod.mill_production.tdate, 'mm')) = "& Mid(Session("CurP_Sales"),5,2) &"" & _
				"                        AND TO_NUMBER ( TO_CHAR (iplasprod.mill_production.tdate, 'yyyy')) = "& Mid(Session("CurP_Sales"),1,4) &"" & _
				"                        AND iplasprod.mill_production.product_code IN ('CPO', 'KER')" & _
				"               GROUP BY iplasprod.mill_production.fcba," & _
				"                        iplasprod.mill_production.product_code)" & _
				"     GROUP BY fcba," & _
				"              product_code," & _
				"              month_qty_sales_bi," & _
				"              year_qty_sales_bi) HPP_prod_bi" & _
				"    INNER JOIN" & _
				"    (SELECT fcba," & _
				"            product_code," & _
				"            month_saldo_awal," & _
				"            year_saldo_awal," & _
				"            saldo_qty_awal," & _
				"            saldo_amount_awal," & _
				"            Price_saldo_awal," & _
				"            month_prod_sbi," & _
				"            year_prod_sbi," & _
				"            qty_prod_sbi," & _
				"            last_month_qty_prod_sbi," & _
				"            amount_prod_sbi," & _
				"            last_month_amount_prod_sbi," & _
				"            Price_prod_sbi," & _
				"            month_qty_sales_sbi," & _
				"            year_qty_sales_sbi," & _
				"            qty_sales_sbi," & _
				"            last_month_qty_sales_sbi," & _
				"            (qty_sales_sbi - last_month_qty_sales_sbi) AS qty_sales_bi," & _
				"            HPP_Qty_level_1," & _
				"            HPP_Amount_level_1," & _
				"            HPP_Price_level_1," & _
				"            HPP_Qty_level_2," & _
				"            HPP_Amount_level_2," & _
				"            HPP_Price_level_2," & _
				"            amount_sales_sbi," & _
				"            last_month_amount_sales_sbi," & _
				"            (amount_sales_sbi - last_month_amount_sales_sbi) AS amount_sales_bi," & _
				"            price_sales_sbi," & _
				"            CASE" & _
				"               WHEN HPP_Qty_level_2 = 0" & _
				"               THEN" & _
				"                  price_saldo_awal" & _
				"               ELSE" & _
				"                  ( (amount_sales_sbi - last_month_amount_sales_sbi) / (qty_sales_sbi - last_month_qty_sales_sbi))" & _
				"            END price_sales_bi," & _
				"            saldo_Qty_akhir," & _
				"            saldo_amount_akhir," & _
				"            price_saldo_akhir" & _
				"       FROM (SELECT fcba," & _
				"                    product_code," & _
				"                    month_saldo_awal," & _
				"                    year_saldo_awal," & _
				"                    saldo_qty_awal," & _
				"                    saldo_amount_awal," & _
				"                    Price_saldo_awal," & _
				"                    month_prod_sbi," & _
				"                    year_prod_sbi," & _
				"                    qty_prod_sbi," & _
				"                    NVL ( IPLASPROD.GetLastMonthqtysbifrommidyear ( product_code, fcba, LAST_DAY ( ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1))), 0) " & _
				"					 last_month_qty_prod_sbi," & _
				"                    amount_prod_sbi," & _
				"                    NVL ( IPLASPROD.GetLastMonthamountprodsbimid ( product_code, fcba, LAST_DAY ( ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1))), 0) " & _
				"					 last_month_amount_prod_sbi," & _
				"                    Price_prod_sbi," & _
				"                    month_qty_sales_sbi," & _
				"                    year_qty_sales_sbi," & _
				"                    qty_sales_sbi," & _
				"                    NVL ( IPLASPROD.GetLastMonthqtysalessbimidyear ( product_code, fcba, LAST_DAY ( ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1))), 0) " & _
				"					 last_month_qty_sales_sbi," & _
				"                    HPP_Qty_level_1," & _
				"                    HPP_Amount_level_1," & _
				"                    HPP_Price_level_1," & _
				"                    HPP_Qty_level_2," & _
				"                    HPP_Amount_level_2," & _
				"                    HPP_Price_level_2," & _
				"                    amount_sales_sbi," & _
				"                    NVL ( IPLASPROD.GetLastMonthamountsalessbimid ( product_code, fcba, LAST_DAY ( ADD_MONTHS (TO_DATE ('"& Mid(Session("CurP_Sales"),5,2) &"', 'fmMM'), -1))), 0) " & _
				"					 last_month_amount_sales_sbi," & _
				"                    price_sales_sbi," & _
				"                    saldo_Qty_akhir," & _
				"                    saldo_amount_akhir," & _
				"                    price_saldo_akhir" & _
				"               FROM (SELECT fcba," & _
				"                            product_code," & _
				"                            month_saldo_awal," & _
				"                            year_saldo_awal," & _
				"                            saldo_qty_awal," & _
				"                            saldo_amount_awal," & _
				"                            Price_saldo_awal," & _
				"                            month_prod_sbi," & _
				"                            year_prod_sbi," & _
				"                            qty_prod_sbi," & _
				"                            amount_prod_sbi," & _
				"                            Price_prod_sbi," & _
				"                            month_qty_sales_sbi," & _
				"                            year_qty_sales_sbi," & _
				"                            qty_sales_sbi," & _
				"                            HPP_Qty_level_1," & _
				"                            HPP_Amount_level_1," & _
				"                            HPP_Price_level_1," & _
				"                            HPP_Qty_level_2," & _
				"                            HPP_Amount_level_2," & _
				"                            HPP_Price_level_2," & _
				"                            amount_sales_sbi," & _
				"                            price_sales_sbi," & _
				"                            saldo_Qty_akhir," & _
				"                            saldo_amount_akhir," & _
				"                            price_saldo_akhir" & _
				"                       FROM (SELECT fcba," & _
				"                                    product_code," & _
				"                                    month_saldo_awal," & _
				"                                    year_saldo_awal," & _
				"                                    saldo_qty_awal," & _
				"                                    saldo_amount_awal," & _
				"                                    Price_saldo_awal," & _
				"                                    month_prod_sbi," & _
				"                                    year_prod_sbi," & _
				"                                    qty_prod_sbi," & _
				"                                    amount_prod_sbi," & _
				"                                    Price_prod_sbi," & _
				"                                    month_qty_sales_sbi," & _
				"                                    year_qty_sales_sbi," & _
				"                                    qty_sales_sbi," & _
				"                                    HPP_Qty_level_1," & _
				"                                    HPP_Amount_level_1," & _
				"                                    HPP_Price_level_1," & _
				"                                    HPP_Qty_level_2," & _
				"                                    HPP_Amount_level_2," & _
				"                                    HPP_Price_level_2," & _
				"                                    (  HPP_Amount_level_1" & _
				"                                     + HPP_Amount_level_2)" & _
				"                                       AS amount_sales_sbi," & _
				"                                    DECODE (" & _
				"                                       qty_sales_sbi," & _
				"                                       0, 0," & _
				"                                       (  (  HPP_Amount_level_1" & _
				"                                           + HPP_Amount_level_2)" & _
				"                                        / qty_sales_sbi))" & _
				"                                       AS price_sales_sbi," & _
				"                                    (  saldo_qty_awal" & _
				"                                     + qty_prod_sbi" & _
				"                                     - qty_sales_sbi)" & _
				"                                       AS saldo_Qty_akhir," & _
				"                                    (  saldo_amount_awal" & _
				"                                     + amount_prod_sbi" & _
				"                                     - HPP_Amount_level_1" & _
				"                                     - HPP_Amount_level_2)" & _
				"                                       AS saldo_amount_akhir," & _
				"                                    (  (  saldo_amount_awal" & _
				"                                        + amount_prod_sbi" & _
				"                                        - HPP_Amount_level_1" & _
				"                                        - HPP_Amount_level_2)" & _
				"                                     / (  saldo_qty_awal" & _
				"                                        + qty_prod_sbi" & _
				"                                        - qty_sales_sbi))" & _
				"                                       AS price_saldo_akhir" & _
				"                               FROM (SELECT fcba," & _
				"                                            product_code," & _
				"                                            month_saldo_awal," & _
				"                                            year_saldo_awal," & _
				"                                            saldo_qty_awal," & _
				"                                            saldo_amount_awal," & _
				"                                            Price_saldo_awal," & _
				"                                            month_prod_sbi," & _
				"                                            year_prod_sbi," & _
				"                                            qty_prod_sbi," & _
				"                                            amount_prod_sbi," & _
				"                                            Price_prod_sbi," & _
				"                                            month_qty_sales_sbi," & _
				"                                            year_qty_sales_sbi," & _
				"                                            qty_sales_sbi," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  saldo_qty_awal" & _
				"                                               ELSE" & _
				"                                                  qty_sales_sbi" & _
				"                                            END" & _
				"                                               AS HPP_Qty_level_1," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  Price_saldo_awal" & _
				"                                               ELSE" & _
				"                                                  Price_saldo_awal" & _
				"                                            END" & _
				"                                               AS HPP_Price_level_1," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  (  saldo_qty_awal" & _
				"                                                   * Price_saldo_awal)" & _
				"                                               ELSE" & _
				"                                                  (  (  qty_sales_sbi" & _
				"                                                      / saldo_qty_awal)" & _
				"                                                   * saldo_amount_awal)" & _
				"                                            END" & _
				"                                               AS HPP_Amount_level_1," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  (  qty_sales_sbi" & _
				"                                                   - saldo_qty_awal)" & _
				"                                               ELSE 0" & _
				"                                            END" & _
				"                                               AS HPP_Qty_level_2," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  Price_prod_sbi" & _
				"                                               ELSE 0" & _
				"                                            END" & _
				"                                               AS HPP_Price_level_2," & _
				"                                            CASE" & _
				"                                               WHEN qty_sales_sbi >=" & _
				"                                                       saldo_qty_awal" & _
				"                                               THEN" & _
				"                                                  (  (  qty_sales_sbi" & _
				"                                                      - saldo_qty_awal)" & _
				"                                                   * Price_prod_sbi)" & _
				"                                               ELSE 0" & _
				"                                            END" & _
				"                                               AS HPP_Amount_level_2" & _
				"                                       FROM (  SELECT fcba," & _
				"                                                      product_code," & _
				"                                                      MAX (month_saldo_awal)" & _
				"                                                         AS month_saldo_awal," & _
				"                                                      MAX (year_saldo_awal)" & _
				"                                                         AS year_saldo_awal," & _
				"                                                      MAX (saldo_qty_awal)" & _
				"                                                         AS saldo_qty_awal," & _
				"                                                      MAX (" & _
				"                                                         saldo_amount_awal)" & _
				"                                                         AS saldo_amount_awal," & _
				"                                                      MAX (Price_saldo_awal)" & _
				"                                                         AS Price_saldo_awal," & _
				"                                                      MAX (month_prod_sbi)" & _
				"                                                         AS month_prod_sbi," & _
				"                                                      MAX (year_prod_sbi)" & _
				"                                                         AS year_prod_sbi," & _
				"                                                      MAX (qty_prod_sbi)" & _
				"                                                         AS qty_prod_sbi," & _
				"                                                      MAX (amount_prod_sbi)" & _
				"                                                         AS amount_prod_sbi," & _
				"                                                      MAX (Price_prod_sbi)" & _
				"                                                         AS Price_prod_sbi," & _
				"                                                      MAX (" & _
				"                                                         month_qty_sales_sbi)" & _
				"                                                         AS month_qty_sales_sbi," & _
				"                                                      MAX (" & _
				"                                                         year_qty_sales_sbi)" & _
				"                                                         AS year_qty_sales_sbi," & _
				"                                                      MAX (qty_sales_sbi)" & _
				"                                                         AS qty_sales_sbi" & _
				"                                                 FROM (SELECT fcba," & _
				"                                                              product_code," & _
				"                                                              month_saldo_awal," & _
				"                                                              year_saldo_awal," & _
				"                                                              qty_cpo_prod_awal" & _
				"                                                                 AS saldo_qty_awal," & _
				"                                                              prop_amount_cpo_prod_awal" & _
				"                                                                 AS saldo_amount_awal," & _
				"                                                              (  prop_amount_cpo_prod_awal" & _
				"                                                               / qty_cpo_prod_awal)" & _
				"                                                                 AS Price_saldo_awal," & _
				"                                                              0 AS month_prod_sbi," & _
				"                                                              0 AS year_prod_sbi," & _
				"                                                              0 AS qty_prod_sbi," & _
				"                                                              0 AS amount_prod_sbi," & _
				"                                                              0 AS Price_prod_sbi," & _
				"                                                              0 AS month_qty_sales_sbi," & _
				"                                                              0 AS year_qty_sales_sbi," & _
				"                                                              0 AS qty_sales_sbi" & _
				"                                                         FROM (SELECT fcba," & _
				"                                                                      'CPO' AS product_code," & _
				"                                                                      month_saldo_awal," & _
				"                                                                      year_saldo_awal," & _
				"                                                                      qty_cpo_prod_awal," & _
				"                                                                      15758662530.5 AS prop_amount_cpo_prod_awal" & _
				"                                                                 FROM (SELECT qty_awal.fcba," & _
				"                                                                              qty_awal.month_saldo_awal," & _
				"                                                                              qty_awal.year_saldo_awal," & _
				"                                                                              qty_awal.qty_cpo_prod_awal," & _
				"                                                                              qty_awal.qty_ker_prod_awal," & _
				"                                                                              biaya_awal.saldo_amount_awal," & _
				"                                                                              (  biaya_awal.saldo_amount_awal" & _
				"                                                                               / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                  + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                     / 2)))" & _
				"                                                                                 AS prop_cpo_prod_awal," & _
				"                                                                              (  (  biaya_awal.saldo_amount_awal" & _
				"                                                                                  / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                     + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                        / 2)))" & _
				"                                                                               / 2)" & _
				"                                                                                 AS prop_ker_prod_awal," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_awal.qty_cpo_prod_awal" & _
				"                                                                                 * ( (  biaya_awal.saldo_amount_awal" & _
				"                                                                                      / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                         + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                            / 2))))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_cpo_prod_awal," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_awal.qty_ker_prod_awal" & _
				"                                                                                 * ( (  (  biaya_awal.saldo_amount_awal" & _
				"                                                                                         / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                            + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                               / 2)))" & _
				"                                                                                      / 2))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_ker_prod_awal" & _
				"                                                                         FROM (  SELECT fcba," & _
				"                                                                                        month_saldo_awal," & _
				"                                                                                        year_saldo_awal," & _
				"                                                                                        SUM ( DECODE ( product_code, 'CPO', saldo_qty_awal, 0))" & _
				"                                                                                           AS qty_cpo_prod_awal," & _
				"                                                                                        SUM ( DECODE ( product_code, 'KER', saldo_qty_awal, 0))" & _
				"                                                                                           AS qty_ker_prod_awal" & _
				"                                                                                   FROM (SELECT iplasprod.mill_summary_tab.fcba," & _
				"                                                                                                CASE" & _
				"                                                                                                   WHEN iplasprod.mill_summary_tab.codedesc =" & _
				"                                                                                                           'CPO003'" & _
				"                                                                                                   THEN" & _
				"                                                                                                      'CPO'" & _
				"                                                                                                   ELSE" & _
				"                                                                                                      'KER'" & _
				"                                                                                                END" & _
				"                                                                                                   AS product_code," & _
				"                                                                                                TO_NUMBER ( TO_CHAR ( iplasprod.mill_summary_tab.ondate, 'mm'))" & _
				"                                                                                                   AS month_saldo_awal," & _
				"                                                                                                TO_NUMBER ( TO_CHAR ( iplasprod.mill_summary_tab.ondate, 'yyyy'))" & _
				"                                                                                                   AS year_saldo_awal," & _
				"                                                                                                NVL ( iplasprod.mill_summary_tab.todate, 0)" & _
				"                                                                                                   AS saldo_qty_awal" & _
				"                                                                                           FROM iplasprod.mill_summary_tab" & _
				"                                                                                          WHERE     iplasprod.mill_summary_tab.fcba =" & _
				"                                                                                                       '"& Session("fcba") &"'" & _
				"                                                                                                AND iplasprod.mill_summary_tab.codedesc IN" & _
				"                                                                                                       ('CPO003'," & _
				"                                                                                                        'PK003')" & _
				"                                                                                                AND iplasprod.mill_summary_tab.ondate = " & _
				"																								 TO_DATE ( '07/2019', 'mm/yyyy') - 1)" & _
				"                                                                               GROUP BY fcba," & _
				"                                                                                        month_saldo_awal," & _
				"                                                                                        year_saldo_awal) qty_awal" & _
				"                                                                              INNER JOIN" & _
				"                                                                              (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                                                           AS fcba," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'mm')))" & _
				"                                                                                           AS month_saldo_awal," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'yyyy')))" & _
				"                                                                                           AS year_saldo_awal," & _
				"                                                                                        SUM (" & _
				"                                                                                           iplasprod.cost_activity.amount)" & _
				"                                                                                           AS saldo_amount_awal" & _
				"                                                                                   FROM iplasprod.cost_activity" & _
				"                                                                                  WHERE     cost_activity.fcba =" & _
				"                                                                                               '"& Session("fcba") &"'" & _
				"                                                                                        AND iplasprod.cost_activity.tdate < TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                        AND cost_activity.jobcode IN" & _
				"                                                                                               (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                                                  FROM iplasprod.parameterdetail" & _
				"                                                                                                 WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                                              'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                                                       AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                                              'CPO'" & _
				"                                                                                                       AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                                              '"& Session("fcba") &"')" & _
				"                                                                               GROUP BY iplasprod.cost_activity.fcba) biaya_awal" & _
				"                                                                                 ON (    qty_awal.fcba =" & _
				"                                                                                            biaya_awal.fcba" & _
				"                                                                                     AND qty_awal.month_saldo_awal =" & _
				"                                                                                            biaya_awal.month_saldo_awal" & _
				"                                                                                     AND qty_awal.year_saldo_awal =" & _
				"                                                                                            biaya_awal.year_saldo_awal))" & _
				"                                                               UNION ALL" & _
				"                                                               SELECT fcba," & _
				"                                                                      'KER'" & _
				"                                                                         AS product_code," & _
				"                                                                      month_saldo_awal," & _
				"                                                                      year_saldo_awal," & _
				"                                                                      qty_ker_prod_awal," & _
				" 																	   1051395609.64 AS prop_amount_ker_prod_awal" & _
				"                                                                 FROM (SELECT qty_awal.fcba," & _
				"                                                                              qty_awal.month_saldo_awal," & _
				"                                                                              qty_awal.year_saldo_awal," & _
				"                                                                              qty_awal.qty_cpo_prod_awal," & _
				"                                                                              qty_awal.qty_ker_prod_awal," & _
				"                                                                              biaya_awal.saldo_amount_awal," & _
				"                                                                              (  biaya_awal.saldo_amount_awal" & _
				"                                                                               / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                  + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                     / 2)))" & _
				"                                                                                 AS prop_cpo_prod_awal," & _
				"                                                                              (  (  biaya_awal.saldo_amount_awal" & _
				"                                                                                  / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                     + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                        / 2)))" & _
				"                                                                               / 2)" & _
				"                                                                                 AS prop_ker_prod_awal," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_awal.qty_cpo_prod_awal" & _
				"                                                                                 * ( (  biaya_awal.saldo_amount_awal" & _
				"                                                                                      / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                         + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                            / 2))))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_cpo_prod_awal," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_awal.qty_ker_prod_awal" & _
				"                                                                                 * ( (  (  biaya_awal.saldo_amount_awal" & _
				"                                                                                         / (  qty_awal.qty_cpo_prod_awal" & _
				"                                                                                            + (  qty_awal.qty_ker_prod_awal" & _
				"                                                                                               / 2)))" & _
				"                                                                                      / 2))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_ker_prod_awal" & _
				"                                                                         FROM (  SELECT fcba," & _
				"                                                                                        month_saldo_awal," & _
				"                                                                                        year_saldo_awal," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'CPO', saldo_qty_awal," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_cpo_prod_awal," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'KER', saldo_qty_awal," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_ker_prod_awal" & _
				"                                                                                   FROM (SELECT iplasprod.mill_summary_tab.fcba," & _
				"                                                                                                CASE" & _
				"                                                                                                   WHEN iplasprod.mill_summary_tab.codedesc =" & _
				"                                                                                                           'CPO003'" & _
				"                                                                                                   THEN" & _
				"                                                                                                      'CPO'" & _
				"                                                                                                   ELSE" & _
				"                                                                                                      'KER'" & _
				"                                                                                                END" & _
				"                                                                                                   AS product_code," & _
				"                                                                                                TO_NUMBER ( TO_CHAR ( iplasprod.mill_summary_tab.ondate, 'mm'))" & _
				"                                                                                                   AS month_saldo_awal," & _
				"                                                                                                TO_NUMBER ( TO_CHAR ( iplasprod.mill_summary_tab.ondate, 'yyyy'))" & _
				"                                                                                                   AS year_saldo_awal," & _
				"                                                                                                NVL (" & _
				"                                                                                                   iplasprod.mill_summary_tab.todate," & _
				"                                                                                                   0)" & _
				"                                                                                                   AS saldo_qty_awal" & _
				"                                                                                           FROM iplasprod.mill_summary_tab" & _
				"                                                                                          WHERE     iplasprod.mill_summary_tab.fcba = '"& Session("fcba") &"'" & _
				"                                                                                                AND iplasprod.mill_summary_tab.codedesc IN" & _
				"                                                                                                       ('CPO003'," & _
				"                                                                                                        'PK003')" & _
				"                                                                                                AND iplasprod.mill_summary_tab.ondate =" & _
				" 																								 TO_DATE ( '07/2019', 'mm/yyyy') - 1)" & _
				"                                                                               GROUP BY fcba," & _
				"                                                                                        month_saldo_awal," & _
				"                                                                                        year_saldo_awal) qty_awal" & _
				"                                                                              INNER JOIN" & _
				"                                                                              (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                                                           AS fcba," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'mm')))" & _
				"                                                                                           AS month_saldo_awal," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'yyyy')))" & _
				"                                                                                           AS year_saldo_awal," & _
				"                                                                                        SUM (" & _
				"                                                                                           iplasprod.cost_activity.amount)" & _
				"                                                                                           AS saldo_amount_awal" & _
				"                                                                                   FROM iplasprod.cost_activity" & _
				"                                                                                  WHERE     cost_activity.fcba =" & _
				"                                                                                               '"& Session("fcba") &"'" & _
				"                                                                                        AND iplasprod.cost_activity.tdate <" & _
				"                                                                                               TO_DATE (" & _
				"                                                                                                  '07/1/2019'," & _
				"                                                                                                  'mm/dd/yyyy')" & _
				"                                                                                        AND cost_activity.jobcode IN" & _
				"                                                                                               (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                                                  FROM iplasprod.parameterdetail" & _
				"                                                                                                 WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                                              'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                                                       AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                                              'KER'" & _
				"                                                                                                       AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                                              '"& Session("fcba") &"')" & _
				"                                                                               GROUP BY iplasprod.cost_activity.fcba) biaya_awal" & _
				"                                                                                 ON (    qty_awal.fcba =" & _
				"                                                                                            biaya_awal.fcba" & _
				"                                                                                     AND qty_awal.month_saldo_awal =" & _
				"                                                                                            biaya_awal.month_saldo_awal" & _
				"                                                                                     AND qty_awal.year_saldo_awal =" & _
				"                                                                                            biaya_awal.year_saldo_awal)))" & _
				"                                                        WHERE product_code IN" & _
				"                                                                 ('CPO'," & _
				"                                                                  'KER')" & _
				"                                                       UNION ALL" & _
				"                                                       SELECT fcba," & _
				"                                                              product_code," & _
				"                                                              0 AS month_saldo_awal," & _
				"                                                              0 AS year_saldo_awal," & _
				"                                                              0 AS saldo_qty_awal," & _
				"                                                              0 AS saldo_amount_awal," & _
				"                                                              0 AS Price_saldo_awal," & _
				"                                                              month_prod_sbi," & _
				"                                                              year_prod_sbi," & _
				"                                                              qty_cpo_prod_sbi" & _
				"                                                                 AS qty_prod_sbi," & _
				"                                                              prop_amount_cpo_prod_sbi" & _
				"                                                                 AS amount_prod_sbi," & _
				"                                                              (  prop_amount_cpo_prod_sbi" & _
				"                                                               / qty_cpo_prod_sbi)" & _
				"                                                                 AS Price_prod_sbi," & _
				"                                                              0 AS month_qty_sales_sbi," & _
				"                                                              0 AS year_qty_sales_sbi," & _
				"                                                              0 AS qty_sales_sbi" & _
				"                                                         FROM (SELECT fcba," & _
				"                                                                      'CPO'" & _
				"                                                                         AS product_code," & _
				"                                                                      month_prod_sbi," & _
				"                                                                      year_prod_sbi," & _
				"                                                                      qty_cpo_prod_sbi," & _
				"                                                                      prop_amount_cpo_prod_sbi" & _
				"                                                                 FROM (SELECT qty_sbi.fcba," & _
				"                                                                              qty_sbi.month_prod_sbi," & _
				"                                                                              qty_sbi.year_prod_sbi," & _
				"                                                                              qty_sbi.qty_cpo_prod_sbi," & _
				"                                                                              qty_sbi.qty_ker_prod_sbi," & _
				"                                                                              biaya_sbi.amount_prod_sbi," & _
				"                                                                              (  biaya_sbi.amount_prod_sbi" & _
				"                                                                               / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                  + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                     / 2)))" & _
				"                                                                                 AS prop_cpo_prod_sbi," & _
				"                                                                              (  (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                  / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                     + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                        / 2)))" & _
				"                                                                               / 2)" & _
				"                                                                                 AS prop_ker_prod_sbi," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                 * ( (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                      / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                         + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                            / 2))))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_cpo_prod_sbi," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                 * ( (  (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                         / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                            + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                               / 2)))" & _
				"                                                                                      / 2))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_ker_prod_sbi" & _
				"                                                                         FROM (  SELECT fcba," & _
				"                                                                                        month_prod_sbi," & _
				"                                                                                        year_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'CPO', qty_prod_sbi," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_cpo_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'KER', qty_prod_sbi," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_ker_prod_sbi" & _
				"                                                                                   FROM (  SELECT iplasprod.mill_production.fcba" & _
				"                                                                                                     AS fcba," & _
				"                                                                                                  iplasprod.mill_production.product_code," & _
				"                                                                                                  MAX (" & _
				"                                                                                                     TO_NUMBER (" & _
				"                                                                                                        TO_CHAR (" & _
				"                                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                                           'mm')))" & _
				"                                                                                                     AS month_prod_sbi," & _
				"                                                                                                  MAX (" & _
				"                                                                                                     TO_NUMBER (" & _
				"                                                                                                        TO_CHAR (" & _
				"                                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                                           'yyyy')))" & _
				"                                                                                                     AS year_prod_sbi," & _
				"                                                                                                  SUM (" & _
				"                                                                                                     iplasprod.mill_production.production_today)" & _
				"                                                                                                     AS qty_prod_sbi" & _
				"                                                                                             FROM iplasprod.mill_production" & _
				"                                                                                            WHERE     iplasprod.mill_production.fcba =" & _
				"                                                                                                         '"& Session("fcba") &"'" & _
				"                                                                                                  AND (iplasprod.mill_production.tdate BETWEEN TO_DATE (" & _
				"                                                                                                                                                  '07/1/2019'," & _
				"                                                                                                                                                  'mm/dd/yyyy')" & _
				"                                                                                                                                           AND LAST_DAY (" & _
				"                                                                                                                                                  TO_DATE (" & _
				"                                                                                                                                                     '"& LastPeriod &"'," & _
				"                                                                                                                                                     'mm/dd/yyyy')))" & _
				"                                                                                                  AND iplasprod.mill_production.product_code IN" & _
				"                                                                                                         ('CPO'," & _
				"                                                                                                          'KER')" & _
				"                                                                                         GROUP BY iplasprod.mill_production.fcba," & _
				"                                                                                                  iplasprod.mill_production.product_code)" & _
				"                                                                               GROUP BY fcba," & _
				"                                                                                        month_prod_sbi," & _
				"                                                                                        year_prod_sbi) qty_sbi" & _
				"                                                                              INNER JOIN" & _
				"                                                                              (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                                                           AS fcba," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'mm')))" & _
				"                                                                                           AS month_prod_sbi," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'yyyy')))" & _
				"                                                                                           AS year_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           iplasprod.cost_activity.amount)" & _
				"                                                                                           AS amount_prod_sbi" & _
				"                                                                                   FROM iplasprod.cost_activity" & _
				"                                                                                  WHERE     cost_activity.fcba =" & _
				"                                                                                               '"& Session("fcba") &"'" & _
				"                                                                                        AND (iplasprod.cost_activity.tdate BETWEEN TO_DATE (" & _
				"                                                                                                                                      '07/1/2019'," & _
				"                                                                                                                                      'mm/dd/yyyy')" & _
				"                                                                                                                               AND LAST_DAY (" & _
				"                                                                                                                                      TO_DATE (" & _
				"                                                                                                                                         '"& LastPeriod &"'," & _
				"                                                                                                                                         'mm/dd/yyyy')))" & _
				"                                                                                        AND cost_activity.jobcode IN" & _
				"                                                                                               (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                                                  FROM iplasprod.parameterdetail" & _
				"                                                                                                 WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                                              'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                                                       AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                                              'CPO'" & _
				"                                                                                                       AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                                              '"& Session("fcba") &"')" & _
				"                                                                               GROUP BY iplasprod.cost_activity.fcba) biaya_sbi" & _
				"                                                                                 ON (    qty_sbi.fcba =" & _
				"                                                                                            biaya_sbi.fcba" & _
				"                                                                                     AND qty_sbi.month_prod_sbi =" & _
				"                                                                                            biaya_sbi.month_prod_sbi" & _
				"                                                                                     AND qty_sbi.year_prod_sbi =" & _
				"                                                                                            biaya_sbi.year_prod_sbi))" & _
				"                                                               UNION ALL" & _
				"                                                               SELECT fcba," & _
				"                                                                      'KER'" & _
				"                                                                         AS product_code," & _
				"                                                                      month_prod_sbi," & _
				"                                                                      year_prod_sbi," & _
				"                                                                      qty_ker_prod_sbi," & _
				"                                                                      prop_amount_ker_prod_sbi" & _
				"                                                                 FROM (SELECT qty_sbi.fcba," & _
				"                                                                              qty_sbi.month_prod_sbi," & _
				"                                                                              qty_sbi.year_prod_sbi," & _
				"                                                                              qty_sbi.qty_cpo_prod_sbi," & _
				"                                                                              qty_sbi.qty_ker_prod_sbi," & _
				"                                                                              biaya_sbi.amount_prod_sbi," & _
				"                                                                              (  biaya_sbi.amount_prod_sbi" & _
				"                                                                               / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                  + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                     / 2)))" & _
				"                                                                                 AS prop_cpo_prod_sbi," & _
				"                                                                              (  (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                  / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                     + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                        / 2)))" & _
				"                                                                               / 2)" & _
				"                                                                                 AS prop_ker_prod_sbi," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                 * ( (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                      / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                         + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                            / 2))))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_cpo_prod_sbi," & _
				"                                                                              ROUND (" & _
				"                                                                                   qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                 * ( (  (  biaya_sbi.amount_prod_sbi" & _
				"                                                                                         / (  qty_sbi.qty_cpo_prod_sbi" & _
				"                                                                                            + (  qty_sbi.qty_ker_prod_sbi" & _
				"                                                                                               / 2)))" & _
				"                                                                                      / 2))," & _
				"                                                                                 2)" & _
				"                                                                                 AS prop_amount_ker_prod_sbi" & _
				"                                                                         FROM (  SELECT fcba," & _
				"                                                                                        month_prod_sbi," & _
				"                                                                                        year_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'CPO', qty_prod_sbi," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_cpo_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           DECODE (" & _
				"                                                                                              product_code," & _
				"                                                                                              'KER', qty_prod_sbi," & _
				"                                                                                              0))" & _
				"                                                                                           AS qty_ker_prod_sbi" & _
				"                                                                                   FROM (  SELECT iplasprod.mill_production.fcba" & _
				"                                                                                                     AS fcba," & _
				"                                                                                                  iplasprod.mill_production.product_code," & _
				"                                                                                                  MAX (" & _
				"                                                                                                     TO_NUMBER (" & _
				"                                                                                                        TO_CHAR (" & _
				"                                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                                           'mm')))" & _
				"                                                                                                     AS month_prod_sbi," & _
				"                                                                                                  MAX (" & _
				"                                                                                                     TO_NUMBER (" & _
				"                                                                                                        TO_CHAR (" & _
				"                                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                                           'yyyy')))" & _
				"                                                                                                     AS year_prod_sbi," & _
				"                                                                                                  SUM (" & _
				"                                                                                                     iplasprod.mill_production.production_today)" & _
				"                                                                                                     AS qty_prod_sbi" & _
				"                                                                                             FROM iplasprod.mill_production" & _
				"                                                                                            WHERE     iplasprod.mill_production.fcba = '"& Session("fcba") &"'" & _
				"                                                                                                  AND (iplasprod.mill_production.tdate BETWEEN" & _
				"																								   TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                                  AND LAST_DAY ( TO_DATE ( '"& LastPeriod &"', 'mm/dd/yyyy')))" & _
				"                                                                                                  AND iplasprod.mill_production.product_code IN" & _
				"                                                                                                         ('CPO'," & _
				"                                                                                                          'KER')" & _
				"                                                                                         GROUP BY iplasprod.mill_production.fcba," & _
				"                                                                                                  iplasprod.mill_production.product_code)" & _
				"                                                                               GROUP BY fcba," & _
				"                                                                                        month_prod_sbi," & _
				"                                                                                        year_prod_sbi) qty_sbi" & _
				"                                                                              INNER JOIN" & _
				"                                                                              (  SELECT iplasprod.cost_activity.fcba" & _
				"                                                                                           AS fcba," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'mm')))" & _
				"                                                                                           AS month_prod_sbi," & _
				"                                                                                        MAX (" & _
				"                                                                                           TO_NUMBER (" & _
				"                                                                                              TO_CHAR (" & _
				"                                                                                                 iplasprod.cost_activity.tdate," & _
				"                                                                                                 'yyyy')))" & _
				"                                                                                           AS year_prod_sbi," & _
				"                                                                                        SUM (" & _
				"                                                                                           iplasprod.cost_activity.amount)" & _
				"                                                                                           AS amount_prod_sbi" & _
				"                                                                                   FROM iplasprod.cost_activity" & _
				"                                                                                  WHERE     cost_activity.fcba =" & _
				"                                                                                               '"& Session("fcba") &"'" & _
				"                                                                                        AND (iplasprod.cost_activity.tdate BETWEEN TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                        AND LAST_DAY ( TO_DATE ( '"& LastPeriod &"', 'mm/dd/yyyy')))" & _
				"                                                                                        AND cost_activity.jobcode IN" & _
				"                                                                                               (SELECT iplasprod.parameterdetail.fccode" & _
				"                                                                                                  FROM iplasprod.parameterdetail" & _
				"                                                                                                 WHERE     iplasprod.parameterdetail.parcode =" & _
				"                                                                                                              'JOBCODE_HARGAPOKOKPRODUKSI'" & _
				"                                                                                                       AND iplasprod.parameterdetail.parheadcode =" & _
				"                                                                                                              'KER'" & _
				"                                                                                                       AND iplasprod.parameterdetail.fcba =" & _
				"                                                                                                              '"& Session("fcba") &"')" & _
				"                                                                               GROUP BY iplasprod.cost_activity.fcba) biaya_sbi" & _
				"                                                                                 ON (    qty_sbi.fcba =" & _
				"                                                                                            biaya_sbi.fcba" & _
				"                                                                                     AND qty_sbi.month_prod_sbi =" & _
				"                                                                                            biaya_sbi.month_prod_sbi" & _
				"                                                                                     AND qty_sbi.year_prod_sbi =" & _
				"                                                                                            biaya_sbi.year_prod_sbi)))" & _
				"                                                        WHERE product_code IN" & _
				"                                                                 ('CPO'," & _
				"                                                                  'KER')" & _
				"                                                       UNION ALL" & _
				"                                                       SELECT fcba," & _
				"                                                              product_code," & _
				"                                                              0 AS month_saldo_awal," & _
				"                                                              0 AS year_saldo_awal," & _
				"                                                              0 AS saldo_qty_awal," & _
				"                                                              0 AS saldo_amount_awal," & _
				"                                                              0 AS Price_saldo_awal," & _
				"                                                              0 AS month_prod_sbi," & _
				"                                                              0 AS year_prod_sbi," & _
				"                                                              0 AS qty_prod_sbi," & _
				"                                                              0 AS amount_prod_sbi," & _
				"                                                              0 AS Price_prod_sbi," & _
				"                                                              month_qty_sales_sbi," & _
				"                                                              year_qty_sales_sbi," & _
				"                                                              qty_sales_sbi" & _
				"                                                         FROM (  SELECT fcba," & _
				"                                                                        product_code," & _
				"                                                                        month_qty_sales_sbi," & _
				"                                                                        year_qty_sales_sbi," & _
				"                                                                        SUM (" & _
				"                                                                           qty_sales_sbi)" & _
				"                                                                           AS qty_sales_sbi_before_adj," & _
				"                                                                        SUM (" & _
				"                                                                           qty_adjust_sbi)" & _
				"                                                                           AS qty_adjust_sbi," & _
				"                                                                        (  SUM (" & _
				"                                                                              qty_sales_sbi)" & _
				"                                                                         + SUM (" & _
				"                                                                              qty_adjust_sbi))" & _
				"                                                                           AS qty_sales_sbi" & _
				"                                                                   FROM (  SELECT iplasprod.weighbridgeticket.fcba," & _
				"                                                                                  iplasprod.weighbridgeticket.product_productcode" & _
				"                                                                                     AS product_code," & _
				"                                                                                  MAX (" & _
				"                                                                                     TO_NUMBER (" & _
				"                                                                                        TO_CHAR (" & _
				"                                                                                           iplasprod.weighbridgeticket.dateout," & _
				"                                                                                           'mm')))" & _
				"                                                                                     AS month_qty_sales_sbi," & _
				"                                                                                  MAX (" & _
				"                                                                                     TO_NUMBER (" & _
				"                                                                                        TO_CHAR (" & _
				"                                                                                           iplasprod.weighbridgeticket.dateout," & _
				"                                                                                           'yyyy')))" & _
				"                                                                                     AS year_qty_sales_sbi," & _
				"                                                                                    SUM (" & _
				"                                                                                       iplasprod.weighbridgeticket.weightout)" & _
				"                                                                                  - SUM (" & _
				"                                                                                       iplasprod.weighbridgeticket.weightin)" & _
				"                                                                                     AS qty_sales_sbi, 0" & _
				"                                                                                     AS qty_adjust_sbi" & _
				"                                                                             FROM iplasprod.weighbridgeticket" & _
				"                                                                            WHERE     iplasprod.weighbridgeticket.fcba =" & _
				"                                                                                         '"& Session("fcba") &"'" & _
				"                                                                                  AND (iplasprod.weighbridgeticket.dateout BETWEEN TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                  AND LAST_DAY ( TO_DATE ( '"& LastPeriod &"', 'mm/dd/yyyy')))" & _
				"                                                                                  AND iplasprod.weighbridgeticket.product_productcode IN" & _
				"                                                                                         ('CPO'," & _
				"                                                                                          'KER')" & _
				"                                                                         GROUP BY iplasprod.weighbridgeticket.fcba," & _
				"                                                                                  iplasprod.weighbridgeticket.product_productcode" & _
				"                                                                         UNION ALL" & _
				"                                                                           SELECT iplasprod.mill_production.fcba" & _
				"                                                                                     AS fcba," & _
				"                                                                                  iplasprod.mill_production.product_code," & _
				"                                                                                  MAX (" & _
				"                                                                                     TO_NUMBER (" & _
				"                                                                                        TO_CHAR (" & _
				"                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                           'mm')))" & _
				"                                                                                     AS month_prod_sbi," & _
				"                                                                                  MAX (" & _
				"                                                                                     TO_NUMBER (" & _
				"                                                                                        TO_CHAR (" & _
				"                                                                                           iplasprod.mill_production.tdate," & _
				"                                                                                           'yyyy')))" & _
				"                                                                                     AS year_prod_sbi, 0" & _
				"                                                                                     AS qty_sales_sbi," & _
				"                                                                                    SUM (" & _
				"                                                                                       iplasprod.mill_production.qty_adjust)" & _
				"                                                                                  * -1" & _
				"                                                                                     AS qty_adjust_sbi" & _
				"                                                                             FROM iplasprod.mill_production" & _
				"                                                                            WHERE     iplasprod.mill_production.fcba =" & _
				"                                                                                         '"& Session("fcba") &"'" & _
				"                                                                                  AND (iplasprod.mill_production.tdate BETWEEN TO_DATE ( '07/1/2019', 'mm/dd/yyyy')" & _
				"                                                                                  AND LAST_DAY ( TO_DATE ( '"& LastPeriod &"', 'mm/dd/yyyy')))" & _
				"                                                                                  AND iplasprod.mill_production.product_code IN" & _
				"                                                                                         ('CPO', 'KER')" & _
				"                                                                         GROUP BY iplasprod.mill_production.fcba," & _
				"                                                                                  iplasprod.mill_production.product_code)" & _
				"                                                               GROUP BY fcba," & _
				"                                                                        product_code," & _
				"                                                                        month_qty_sales_sbi," & _
				"                                                                        year_qty_sales_sbi))" & _
				"                                             GROUP BY fcba, product_code))))))" & _
				"    hpp_penjualan_mtd" & _
				"       ON (    HPP_prod_bi.fcba = hpp_penjualan_mtd.fcba" & _
				"           AND HPP_prod_bi.product_code =" & _
				"                  hpp_penjualan_mtd.product_code" & _
				"           AND HPP_prod_bi.month_qty_sales_bi =" & _
				"                  hpp_penjualan_mtd.month_prod_sbi" & _
				"           AND HPP_prod_bi.year_qty_sales_bi =" & _
				"                  hpp_penjualan_mtd.year_prod_sbi)" & _
				"    INNER JOIN iplasprod.parameter pm" & _
				"       ON (    HPP_prod_bi.fcba = pm.fcba" & _
				"           AND HPP_prod_bi.product_code = pm.fccode)" & _
				" WHERE pm.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" 
				
	'response.write qSalesCR
	Set rsSalesCR = Server.CreateObject("ADODB.Recordset")
	rsSalesCR.Open qSalesCR, connstr
	
	if not rsSalesCR.eof then
		Do While Not rsSalesCR.eof		  
			 
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityCRHPP="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & temp_datePeriodControl & "','mm/dd/yyyy'),'" & rsSalesCR("jobcode_persediaan") & "','SL','SL','Cr-HARGAPOKOKPPENJUALAN'," & _
								 "'','CB','BS'," & _
								 "'" &  rsSalesCR("total_amount_sales_bi") & "','Cr-HARGAPOKOKPPENJUALAN'," & "0,0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " 'Product Code : ' ,'"  & rsSalesCR("product_code") & "'," & _
								 " '',''," & _
								 " '','', " & _
								 " '','', " & _
								 " '', '','Cr', " & _
								 " " & rsSalesCR("month_qty_sales_bi") & "," & rsSalesCR("year_qty_sales_bi") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "
			' response.write qCost_ActivityCRHPP
			 connstr.execute(qCost_ActivityCRHPP)
			 connStr.execute("Commit")

		rsSalesCR.MoveNext
	  loop
	  response.Write "Process Harga Pokok Penjualan Sales Credit Complete.... !!<br>"
	  response.Flush
	end if
	'-- End Insert Credit Harga Pokok Penjualan

	'-- End sales Shipment
	end if

	if Session("Estate_Mill") = "HO" then
	'-- Begin sales Invoice	

	'-- Insert Debet Piutang Penjualan
	qSalesInvDB1 = 	" SELECT salesinv.fcba," & _
					"        salesinv.invoice_by," & _
					"        salesinv.transdate," & _
					"        to_number(to_char(salesinv.transdate,'MM')) ACCOUNTINGPERIOD," & _
					"        to_number(to_char(salesinv.transdate,'YYYY')) ACCOUNTINGYEAR," & _
					"        salesinv.salescontractno," & _
					"        salesinv.customercode," & _
					"        salesinv.dp_vat," & _
					"        salesinv.amount_dpvat," & _
					"        salesinv.dp_aftervat," & _
					"        salesinv.percentage_dp," & _
					"        salesinv.invoicecode," & _
					"        salesinv.lineno," & _
					"        salesinv.productcode," & _
					"        pm.param_5 AS Jobcode_Sales," & _
					"        salesinv.salesinvoice_beforeppn, " & _
					"        salesinv.ppn," & _
					"        salesinv.salesinvoice_ppn," & _
					"        salesinv.salesinvoice_afterppn," & _
					"        pm.param_3 AS Jobcode_UangMukaPenjualan," & _
					"        salesinv.Penyelesaian_DP, " & _
					"        pm.param_4 AS Jobcode_PPnKeluaran," & _
					"        salesinv.Penyelesaian_PPN, " & _
					"        pm.param_6 as Jobcode_PiutangPenjualan," & _
					"        CASE" & _
					"           WHEN salesinv.penyelesaian_dp = 0" & _
					"           THEN" & _
					"              (salesinv.salesinvoice_beforeppn + salesinv.Penyelesaian_PPN)" & _
					"           ELSE" & _
					"              (  (salesinv.salesinvoice_beforeppn + salesinv.Penyelesaian_PPN)" & _
					"               - salesinv.Penyelesaian_DP)" & _
					"        END" & _
					"           AS Piutang_Penjualan " & _
					"   FROM (SELECT fcba," & _
					"                invoice_by," & _
					"                transdate," & _
					"                salescontractno," & _
					"                customercode," & _
					"                dp_vat," & _
					"                amount_dpvat," & _
					"                dp_aftervat," & _
					"                percentage_dp," & _
					"                invoicecode," & _
					"                lineno," & _
					"                productcode," & _
					"                salesinvoice_beforeppn," & _
					"                ppn," & _
					"                salesinvoice_ppn," & _
					"                salesinvoice_afterppn," & _
					"                Penyelesaian_DP," & _
					"                CASE" & _
					"                   WHEN (penyelesaian_dp = 0 AND dp_vat = 0)" & _
					"                   THEN" & _
					"                      salesinvoice_ppn" & _
					"                   WHEN (    penyelesaian_dp = salesinvoice_beforeppn" & _
					"                         AND dp_vat > 0)" & _
					"                   THEN" & _
					" 0" & _
					"                   WHEN (penyelesaian_dp > 0 AND dp_vat = 0)" & _
					"                   THEN" & _
					"                      salesinvoice_ppn" & _
					"                   WHEN (penyelesaian_dp > 0 AND dp_vat > 0)" & _
					"                   THEN" & _
					"                      ( (salesinvoice_ppn * percentage_dp) / 100)" & _
					"                   ELSE" & _
					" 0" & _
					"                END" & _
					"                   AS Penyelesaian_PPN" & _
					"           FROM (SELECT iplasprod.salesinvoicedtl.fcba," & _
					"                        iplasprod.salesinvoicehdr.invoice_by," & _
					"                        iplasprod.salesinvoicehdr.transdate," & _
					"                        iplasprod.salesagreementhdr.salescontractno," & _
					"                        iplasprod.salesagreementhdr.customercode," & _
					"                        iplasprod.salesagreementhdr.dp_vat," & _
					"                        iplasprod.salesagreementhdr.amount_dpvat," & _
					"                        iplasprod.salesagreementhdr.dp_aftervat," & _
					"                        iplasprod.salesagreementhdr.percentage_dp," & _
					"                        iplasprod.salesinvoicedtl.invoicecode," & _
					"                        iplasprod.salesinvoicedtl.lineno," & _
					"                        iplasprod.salesinvoicedtl.productcode," & _
					"                        iplasprod.salesinvoicedtl.amount" & _
					"                           AS salesinvoice_beforeppn," & _
					"                        iplasprod.salesinvoicehdr.vat AS ppn," & _
					"                        ROUND ( (  (  iplasprod.salesinvoicedtl.amount" & _
					"                            * iplasprod.salesinvoicehdr.vat)" & _
					"                         / 100),0)" & _
					"                           AS salesinvoice_ppn," & _
					"                        (  iplasprod.salesinvoicedtl.amount" & _
					"                         + (  (  iplasprod.salesinvoicedtl.amount" & _
					"                               * iplasprod.salesinvoicehdr.vat)" & _
					"                            / 100))" & _
					"                           AS salesinvoice_afterppn," & _
					"                        CASE" & _
					"                           WHEN iplasprod.salesagreementhdr.percentage_dp > 0" & _
					"                           THEN" & _
					"                              (  (  iplasprod.salesinvoicedtl.amount" & _
					"                                  * iplasprod.salesagreementhdr.percentage_dp)" & _
					"                               / 100)" & _
					"                           ELSE" & _
					" 0" & _
					"                        END" & _
					"                           AS Penyelesaian_DP" & _
					"                   FROM iplasprod.salesinvoicedtl" & _
					"                        INNER JOIN iplasprod.salesinvoicehdr" & _
					"                           ON (    iplasprod.salesinvoicedtl.fcba =" & _
					"                                      iplasprod.salesinvoicehdr.fcba" & _
					"                               AND iplasprod.salesinvoicedtl.invoicecode =" & _
					"                                      iplasprod.salesinvoicehdr.invoicecode)" & _
					"                        INNER JOIN iplasprod.salesagreementhdr" & _
					"                           ON (    iplasprod.salesinvoicehdr.fcba =" & _
					"                                      iplasprod.salesagreementhdr.fcba" & _
					"                               AND iplasprod.salesinvoicehdr.salescontractno =" & _
					"                                      iplasprod.salesagreementhdr.salescontractno)" & _
					"                  WHERE     TO_NUMBER (" & _
					"                               TO_CHAR (iplasprod.salesinvoicehdr.transdate," & _
					"                                        'mm')) =" & _
					"                               (SELECT currentaccountingperiod FROM iplasprod.periodcontrol" & _
					"                                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
					"                                       AND UPPER ( iplasprod.periodcontrol.SYSTEM) = 'SALES') " & _
					"                  AND     TO_NUMBER (" & _
					"                               TO_CHAR (iplasprod.salesinvoicehdr.transdate," & _
					"                                        'yyyy')) =" & _
					"                               (SELECT currentaccountingyear FROM iplasprod.periodcontrol" & _
					"                                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
					"                                       AND UPPER ( iplasprod.periodcontrol.SYSTEM) = 'SALES') " & _
					"                        AND iplasprod.salesinvoicehdr.invoice_by = '"& session("fcBA") &"' )) salesinv" & _
					"        INNER JOIN iplasprod.parameter pm" & _
					"           ON (salesinv.fcba = pm.fcba AND salesinv.productcode = pm.fccode)" & _
					"        INNER JOIN iplasprod.customer cs" & _
					"           ON (salesinv.fcba = cs.fcba AND salesinv.customercode = cs.customercode)" & _
					"  WHERE pm.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" 
	
	' response.write qSalesInvDB1
	Set rsSalesInvDB1 = Server.CreateObject("ADODB.Recordset")
	rsSalesInvDB1.Open qSalesInvDB1, connstr
	if not rsSalesInvDB1.eof then
		Do While Not rsSalesInvDB1.eof	   
			' tmpAMountDB1 = Round(rsSalesInvDB1("AMOUNT_DB1"))
			maxLineNo = maxLineNo + 1

			qJob_PiutangPenjualan = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"' and fccode = '"& rsSalesInvDB1("Jobcode_PiutangPenjualan") &"'"
					
			Set rsJob_PiutangPenjualan = Server.CreateObject("ADODB.Recordset")
			rsJob_PiutangPenjualan.CursorType = 3
			rsJob_PiutangPenjualan.CursorLocation = 3
			rsJob_PiutangPenjualan.Open qJob_PiutangPenjualan, connstr

			if not rsJob_PiutangPenjualan.eof then
				job_category = rsJob_PiutangPenjualan("job_category") 
			end if
			rsJob_PiutangPenjualan.close
			set rsJob_PiutangPenjualan = Nothing

			qCost_ActivityDB1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesInvDB1("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesInvDB1("Jobcode_PiutangPenjualan") & "','SL','SL','" & rsSalesInvDB1("invoicecode") & "'," & _
							 "'','" & job_category & "','" & rsSalesInvDB1("customercode") & "'," & _
							 "'" & rsSalesInvDB1("Piutang_Penjualan") & "','" & rsSalesInvDB1("salescontractno") & "'," & "0,0,'" & rsSalesInvDB1("INVOICE_BY") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " 'TC - Sales Invoice Code : ', '"  & rsSalesInvDB1("invoicecode") & "'," & _
							 " 'Auto - Sales Contract No : ', '"  & rsSalesInvDB1("salescontractno") & "'," & _
							 " 'Auto - Customer Code: ', '"  & rsSalesInvDB1("customercode") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsSalesInvDB1("LINENO") & "','Db', " & _
							 " " & rsSalesInvDB1("ACCOUNTINGPERIOD") & "," & rsSalesInvDB1("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
			' response.Write qCost_ActivityDB1 
			 connstr.execute(qCost_ActivityDB1)
			 connStr.execute("Commit")
			 
			rsSalesInvDB1.MoveNext
		loop
		response.Write "Process Sales Invoice Debet Piutang Penjualan Complete.... !!<br>"
		response.Flush
	end if
	'-- End Insert Debet Piutang Penjualan

	'-- Insert Debet Uang Muka Penjualan
	qSalesInvDB1 = 	" SELECT salesinv.fcba," & _
					"        salesinv.invoice_by," & _
					"        salesinv.transdate," & _
					"        to_number(to_char(salesinv.transdate,'MM')) ACCOUNTINGPERIOD," & _
					"        to_number(to_char(salesinv.transdate,'YYYY')) ACCOUNTINGYEAR," & _
					"        salesinv.salescontractno," & _
					"        salesinv.customercode," & _
					"        salesinv.dp_vat," & _
					"        salesinv.amount_dpvat," & _
					"        salesinv.dp_aftervat," & _
					"        salesinv.percentage_dp," & _
					"        salesinv.invoicecode," & _
					"        salesinv.lineno," & _
					"        salesinv.productcode," & _
					"        pm.param_5 AS Jobcode_Sales," & _
					"        salesinv.salesinvoice_beforeppn, " & _
					"        salesinv.ppn," & _
					"        salesinv.salesinvoice_ppn," & _
					"        salesinv.salesinvoice_afterppn," & _
					"        pm.param_3 AS Jobcode_UangMukaPenjualan," & _
					"        salesinv.Penyelesaian_DP, " & _
					"        pm.param_4 AS Jobcode_PPnKeluaran," & _
					"        salesinv.Penyelesaian_PPN, " & _
					"        cs.controljob as jobcode_piutangpenjualan," & _
					"        CASE" & _
					"           WHEN salesinv.penyelesaian_dp = 0" & _
					"           THEN" & _
					"              (salesinv.salesinvoice_beforeppn + salesinv.Penyelesaian_PPN)" & _
					"           ELSE" & _
					"              (  (salesinv.salesinvoice_beforeppn + salesinv.Penyelesaian_PPN)" & _
					"               - salesinv.Penyelesaian_DP)" & _
					"        END" & _
					"           AS Piutang_Penjualan " & _
					"   FROM (SELECT fcba," & _
					"                invoice_by," & _
					"                transdate," & _
					"                salescontractno," & _
					"                customercode," & _
					"                dp_vat," & _
					"                amount_dpvat," & _
					"                dp_aftervat," & _
					"                percentage_dp," & _
					"                invoicecode," & _
					"                lineno," & _
					"                productcode," & _
					"                salesinvoice_beforeppn," & _
					"                ppn," & _
					"                salesinvoice_ppn," & _
					"                salesinvoice_afterppn," & _
					"                Penyelesaian_DP," & _
					"                CASE" & _
					"                   WHEN (penyelesaian_dp = 0 AND dp_vat = 0)" & _
					"                   THEN" & _
					"                      salesinvoice_ppn" & _
					"                   WHEN (    penyelesaian_dp = salesinvoice_beforeppn" & _
					"                         AND dp_vat > 0)" & _
					"                   THEN" & _
					" 0" & _
					"                   WHEN (penyelesaian_dp > 0 AND dp_vat = 0)" & _
					"                   THEN" & _
					"                      salesinvoice_ppn" & _
					"                   WHEN (penyelesaian_dp > 0 AND dp_vat > 0)" & _
					"                   THEN" & _
					"                      ( (salesinvoice_ppn * percentage_dp) / 100)" & _
					"                   ELSE" & _
					" 0" & _
					"                END" & _
					"                   AS Penyelesaian_PPN" & _
					"           FROM (SELECT iplasprod.salesinvoicedtl.fcba," & _
					"                        iplasprod.salesinvoicehdr.invoice_by," & _
					"                        iplasprod.salesinvoicehdr.transdate," & _
					"                        iplasprod.salesagreementhdr.salescontractno," & _
					"                        iplasprod.salesagreementhdr.customercode," & _
					"                        iplasprod.salesagreementhdr.dp_vat," & _
					"                        iplasprod.salesagreementhdr.amount_dpvat," & _
					"                        iplasprod.salesagreementhdr.dp_aftervat," & _
					"                        iplasprod.salesagreementhdr.percentage_dp," & _
					"                        iplasprod.salesinvoicedtl.invoicecode," & _
					"                        iplasprod.salesinvoicedtl.lineno," & _
					"                        iplasprod.salesinvoicedtl.productcode," & _
					"                        iplasprod.salesinvoicedtl.amount" & _
					"                           AS salesinvoice_beforeppn," & _
					"                        iplasprod.salesinvoicehdr.vat AS ppn," & _
					"                        ROUND ( (  (  iplasprod.salesinvoicedtl.amount" & _
					"                            * iplasprod.salesinvoicehdr.vat)" & _
					"                         / 100),0)" & _
					"                           AS salesinvoice_ppn," & _
					"                        (  iplasprod.salesinvoicedtl.amount" & _
					"                         + (  (  iplasprod.salesinvoicedtl.amount" & _
					"                               * iplasprod.salesinvoicehdr.vat)" & _
					"                            / 100))" & _
					"                           AS salesinvoice_afterppn," & _
					"                        CASE" & _
					"                           WHEN iplasprod.salesagreementhdr.percentage_dp > 0" & _
					"                           THEN" & _
					"                              (  (  iplasprod.salesinvoicedtl.amount" & _
					"                                  * iplasprod.salesagreementhdr.percentage_dp)" & _
					"                               / 100)" & _
					"                           ELSE" & _
					" 0" & _
					"                        END" & _
					"                           AS Penyelesaian_DP" & _
					"                   FROM iplasprod.salesinvoicedtl" & _
					"                        INNER JOIN iplasprod.salesinvoicehdr" & _
					"                           ON (    iplasprod.salesinvoicedtl.fcba =" & _
					"                                      iplasprod.salesinvoicehdr.fcba" & _
					"                               AND iplasprod.salesinvoicedtl.invoicecode =" & _
					"                                      iplasprod.salesinvoicehdr.invoicecode)" & _
					"                        INNER JOIN iplasprod.salesagreementhdr" & _
					"                           ON (    iplasprod.salesinvoicehdr.fcba =" & _
					"                                      iplasprod.salesagreementhdr.fcba" & _
					"                               AND iplasprod.salesinvoicehdr.salescontractno =" & _
					"                                      iplasprod.salesagreementhdr.salescontractno)" & _
					"                  WHERE     TO_NUMBER (" & _
					"                               TO_CHAR (iplasprod.salesinvoicehdr.transdate," & _
					"                                        'mm')) =" & _
					"                               (SELECT currentaccountingperiod FROM iplasprod.periodcontrol" & _
					"                                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
					"                                       AND UPPER ( iplasprod.periodcontrol.SYSTEM) = 'SALES') " & _
					"                  AND     TO_NUMBER (" & _
					"                               TO_CHAR (iplasprod.salesinvoicehdr.transdate," & _
					"                                        'yyyy')) =" & _
					"                               (SELECT currentaccountingyear FROM iplasprod.periodcontrol" & _
					"                                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
					"                                       AND UPPER ( iplasprod.periodcontrol.SYSTEM) = 'SALES') " & _
					"                        AND iplasprod.salesinvoicehdr.invoice_by = '"& session("fcBA") &"' )) salesinv" & _
					"        INNER JOIN iplasprod.parameter pm" & _
					"           ON (salesinv.fcba = pm.fcba AND salesinv.productcode = pm.fccode)" & _
					"        INNER JOIN iplasprod.customer cs" & _
					"           ON (salesinv.fcba = cs.fcba AND salesinv.customercode = cs.customercode)" & _
					"  WHERE pm.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" 

	Set rsSalesInvDB1 = Server.CreateObject("ADODB.Recordset")
	rsSalesInvDB1.Open qSalesInvDB1, connstr
	if not rsSalesInvDB1.eof then
		Do While Not rsSalesInvDB1.eof	   
			' tmpAMountDB1 = Round(rsSalesInvDB1("AMOUNT_DB1"))
			maxLineNo = maxLineNo + 1

			' Target Type
			qJob_UangMukaPenjualan = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"' and fccode = '"& rsSalesInvDB1("jobcode_uangmukapenjualan") &"'"
					
			Set rsUangMukaPenjualan = Server.CreateObject("ADODB.Recordset")
			rsUangMukaPenjualan.CursorType = 3
			rsUangMukaPenjualan.CursorLocation = 3
			rsUangMukaPenjualan.Open qJob_UangMukaPenjualan, connstr

			if not rsUangMukaPenjualan.eof then
				tt_job_category = rsUangMukaPenjualan("job_category") 
			end if
			rsUangMukaPenjualan.close
			set rsUangMukaPenjualan = Nothing
			' End Target Type

			' Target Code
			qJob_UangMukaPenjualan = 	" SELECT fccode" & _
										"   FROM iplasprod.gcmaster" & _
										"  WHERE     fcba = '"& session("fcba") &"'" & _
										"        AND item =" & _
										"               (SELECT job_category" & _
										"                  FROM iplasprod.job" & _
										"                 WHERE fcba = '"& session("fcBA") &"' AND fccode = '"& rsSalesInvDB1("jobcode_uangmukapenjualan") &"')"  
					
			Set rsUangMukaPenjualan = Server.CreateObject("ADODB.Recordset")
			rsUangMukaPenjualan.CursorType = 3
			rsUangMukaPenjualan.CursorLocation = 3
			rsUangMukaPenjualan.Open qJob_UangMukaPenjualan, connstr

			if not rsUangMukaPenjualan.eof then
				tc_job_category = rsUangMukaPenjualan("fccode") 
			end if
			rsUangMukaPenjualan.close
			set rsUangMukaPenjualan = Nothing
			' End Target Code

			qCost_ActivityDB1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesInvDB1("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesInvDB1("jobcode_uangmukapenjualan") & "','SL','SL','" & rsSalesInvDB1("invoicecode") & "'," & _
							 "'','" & tt_job_category & "','" & tc_job_category & "'," & _
							 "'" & rsSalesInvDB1("penyelesaian_dp") & "','" & rsSalesInvDB1("salescontractno") & "'," & "0,0,'" & rsSalesInvDB1("INVOICE_BY") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " 'TC - Sales Invoice Code : ', '"  & rsSalesInvDB1("invoicecode") & "'," & _
							 " 'Auto - Sales Contract No : ', '"  & rsSalesInvDB1("salescontractno") & "'," & _
							 " 'Auto - Customer Code: ', '"  & rsSalesInvDB1("customercode") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsSalesInvDB1("LINENO") & "','Db', " & _
							 " " & rsSalesInvDB1("ACCOUNTINGPERIOD") & "," & rsSalesInvDB1("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDB1)
			 connStr.execute("Commit")
			 
			rsSalesInvDB1.MoveNext
		loop
		response.Write "Process Sales Invoice Debet Uang Muka Penjualan Complete.... !!<br>"
		response.Flush
	end if
	'-- End Insert Debet Uang Muka Penjualan

	'-- Insert Credit Sales
	qSalesInvCR1 = 	" SELECT salesinv.fcba," & _
					"        salesinv.invoice_by," & _
					"        salesinv.transdate," & _
					"        to_number(to_char(salesinv.transdate,'MM')) ACCOUNTINGPERIOD," & _
					"        to_number(to_char(salesinv.transdate,'YYYY')) ACCOUNTINGYEAR," & _
					"        salesinv.salescontractno," & _
					"        salesinv.customercode," & _
					"        salesinv.dp_vat," & _
					"        salesinv.amount_dpvat," & _
					"        salesinv.dp_aftervat," & _
					"        salesinv.percentage_dp," & _
					"        salesinv.invoicecode," & _
					"        salesinv.lineno," & _
					"        salesinv.productcode," & _
					"        pm.param_5 AS Jobcode_Sales," & _
					"        ( -1 * salesinv.salesinvoice_beforeppn) salesinvoice_beforeppn, " & _
					"        salesinv.ppn," & _
					"        salesinv.salesinvoice_ppn," & _
					"        salesinv.salesinvoice_afterppn," & _
					"        pm.param_3 AS Jobcode_UangMukaPenjualan," & _
					"        salesinv.Penyelesaian_DP, " & _
					"        pm.param_4 AS Jobcode_PPnKeluaran," & _
					"        salesinv.Penyelesaian_PPN, " & _
					"        cs.controljob as jobcode_piutangpenjualan," & _
					"        CASE" & _
					"           WHEN salesinv.penyelesaian_dp = 0" & _
					"           THEN" & _
					"              (salesinv.salesinvoice_beforeppn + salesinv.Penyelesaian_PPN)" & _
					"           ELSE" & _
					"              (  (salesinv.salesinvoice_beforeppn + salesinv.Penyelesaian_PPN)" & _
					"               - salesinv.Penyelesaian_DP)" & _
					"        END" & _
					"           AS Piutang_Penjualan " & _
					"   FROM (SELECT fcba," & _
					"                invoice_by," & _
					"                transdate," & _
					"                salescontractno," & _
					"                customercode," & _
					"                dp_vat," & _
					"                amount_dpvat," & _
					"                dp_aftervat," & _
					"                percentage_dp," & _
					"                invoicecode," & _
					"                lineno," & _
					"                productcode," & _
					"                salesinvoice_beforeppn," & _
					"                ppn," & _
					"                salesinvoice_ppn," & _
					"                salesinvoice_afterppn," & _
					"                Penyelesaian_DP," & _
					"                CASE" & _
					"                   WHEN (penyelesaian_dp = 0 AND dp_vat = 0)" & _
					"                   THEN" & _
					"                      salesinvoice_ppn" & _
					"                   WHEN (    penyelesaian_dp = salesinvoice_beforeppn" & _
					"                         AND dp_vat > 0)" & _
					"                   THEN" & _
					" 0" & _
					"                   WHEN (penyelesaian_dp > 0 AND dp_vat = 0)" & _
					"                   THEN" & _
					"                      salesinvoice_ppn" & _
					"                   WHEN (penyelesaian_dp > 0 AND dp_vat > 0)" & _
					"                   THEN" & _
					"                      ( (salesinvoice_ppn * percentage_dp) / 100)" & _
					"                   ELSE" & _
					" 0" & _
					"                END" & _
					"                   AS Penyelesaian_PPN" & _
					"           FROM (SELECT iplasprod.salesinvoicedtl.fcba," & _
					"                        iplasprod.salesinvoicehdr.invoice_by," & _
					"                        iplasprod.salesinvoicehdr.transdate," & _
					"                        iplasprod.salesagreementhdr.salescontractno," & _
					"                        iplasprod.salesagreementhdr.customercode," & _
					"                        iplasprod.salesagreementhdr.dp_vat," & _
					"                        iplasprod.salesagreementhdr.amount_dpvat," & _
					"                        iplasprod.salesagreementhdr.dp_aftervat," & _
					"                        iplasprod.salesagreementhdr.percentage_dp," & _
					"                        iplasprod.salesinvoicedtl.invoicecode," & _
					"                        iplasprod.salesinvoicedtl.lineno," & _
					"                        iplasprod.salesinvoicedtl.productcode," & _
					"                        iplasprod.salesinvoicedtl.amount" & _
					"                           AS salesinvoice_beforeppn," & _
					"                        iplasprod.salesinvoicehdr.vat AS ppn," & _
					"                        ROUND ( ( (  iplasprod.salesinvoicedtl.amount" & _
					"                            * iplasprod.salesinvoicehdr.vat)" & _
					"                         / 100),0)" & _
					"                           AS salesinvoice_ppn," & _
					"                        (  iplasprod.salesinvoicedtl.amount" & _
					"                         + (  (  iplasprod.salesinvoicedtl.amount" & _
					"                               * iplasprod.salesinvoicehdr.vat)" & _
					"                            / 100))" & _
					"                           AS salesinvoice_afterppn," & _
					"                        CASE" & _
					"                           WHEN iplasprod.salesagreementhdr.percentage_dp > 0" & _
					"                           THEN" & _
					"                              (  (  iplasprod.salesinvoicedtl.amount" & _
					"                                  * iplasprod.salesagreementhdr.percentage_dp)" & _
					"                               / 100)" & _
					"                           ELSE" & _
					" 0" & _
					"                        END" & _
					"                           AS Penyelesaian_DP" & _
					"                   FROM iplasprod.salesinvoicedtl" & _
					"                        INNER JOIN iplasprod.salesinvoicehdr" & _
					"                           ON (    iplasprod.salesinvoicedtl.fcba =" & _
					"                                      iplasprod.salesinvoicehdr.fcba" & _
					"                               AND iplasprod.salesinvoicedtl.invoicecode =" & _
					"                                      iplasprod.salesinvoicehdr.invoicecode)" & _
					"                        INNER JOIN iplasprod.salesagreementhdr" & _
					"                           ON (    iplasprod.salesinvoicehdr.fcba =" & _
					"                                      iplasprod.salesagreementhdr.fcba" & _
					"                               AND iplasprod.salesinvoicehdr.salescontractno =" & _
					"                                      iplasprod.salesagreementhdr.salescontractno)" & _
					"                  WHERE     TO_NUMBER (" & _
					"                               TO_CHAR (iplasprod.salesinvoicehdr.transdate," & _
					"                                        'mm')) =" & _
					"                               (SELECT currentaccountingperiod FROM iplasprod.periodcontrol" & _
					"                                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
					"                                       AND UPPER ( iplasprod.periodcontrol.SYSTEM) = 'SALES') " & _
					"                  AND     TO_NUMBER (" & _
					"                               TO_CHAR (iplasprod.salesinvoicehdr.transdate," & _
					"                                        'yyyy')) =" & _
					"                               (SELECT currentaccountingyear FROM iplasprod.periodcontrol" & _
					"                                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
					"                                       AND UPPER ( iplasprod.periodcontrol.SYSTEM) = 'SALES') " & _
					"                        AND iplasprod.salesinvoicehdr.invoice_by = '"& session("fcBA") &"' )) salesinv" & _
					"        INNER JOIN iplasprod.parameter pm" & _
					"           ON (salesinv.fcba = pm.fcba AND salesinv.productcode = pm.fccode)" & _
					"        INNER JOIN iplasprod.customer cs" & _
					"           ON (salesinv.fcba = cs.fcba AND salesinv.customercode = cs.customercode)" & _
					"  WHERE pm.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" 

	Set rsSalesInvCR1 = Server.CreateObject("ADODB.Recordset")
	rsSalesInvCR1.Open qSalesInvCR1, connstr
	if not rsSalesInvCR1.eof then
		Do While Not rsSalesInvCR1.eof	   
			' tmpAMountDB1 = Round(rsSalesInvCR1("AMOUNT_DB1"))
			maxLineNo = maxLineNo + 1

			' Target Type
			qJob_Sales = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"' and fccode = '"& rsSalesInvCR1("jobcode_uangmukapenjualan") &"'"
					
			Set rsSales = Server.CreateObject("ADODB.Recordset")
			rsSales.CursorType = 3
			rsSales.CursorLocation = 3
			rsSales.Open qJob_Sales, connstr

			if not rsSales.eof then
				tt_job_category = rsSales("job_category") 
			end if
			rsSales.close
			set rsSales = Nothing
			' End Target Type

			' Target Code
			qJob_Sales = 	" SELECT fccode" & _
							"   FROM iplasprod.gcmaster" & _
							"  WHERE     fcba = '"& session("fcba") &"'" & _
							"        AND item =" & _
							"               (SELECT job_category" & _
							"                  FROM iplasprod.job" & _
							"                 WHERE fcba = '"& session("fcBA") &"' AND fccode = '"& rsSalesInvCR1("Jobcode_Sales") &"')"  
					
			Set rsSales = Server.CreateObject("ADODB.Recordset")
			rsSales.CursorType = 3
			rsSales.CursorLocation = 3
			rsSales.Open qJob_Sales, connstr

			if not rsSales.eof then
				tc_job_category = rsSales("fccode")
			end if
			rsSales.close
			set rsSales = Nothing
			' End Target Code

			qCost_ActivityDB1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesInvCR1("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesInvCR1("Jobcode_Sales") & "','SL','SL','" & rsSalesInvCR1("invoicecode") & "'," & _
							 "'','" & tt_job_category & "','" & tc_job_category & "'," & _
							 "'" & rsSalesInvCR1("salesinvoice_beforeppn") & "','" & rsSalesInvCR1("salescontractno") & "'," & "0,0,'" & rsSalesInvCR1("INVOICE_BY") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " 'TC - Sales Invoice Code : ', '"  & rsSalesInvCR1("invoicecode") & "'," & _
							 " 'Auto - Sales Contract No : ', '"  & rsSalesInvCR1("salescontractno") & "'," & _
							 " 'Auto - Customer Code: ', '"  & rsSalesInvCR1("customercode") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsSalesInvCR1("LINENO") & "','Cr', " & _
							 " " & rsSalesInvCR1("ACCOUNTINGPERIOD") & "," & rsSalesInvCR1("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDB1)
			 connStr.execute("Commit")
			 
			rsSalesInvCR1.MoveNext
		loop
		response.Write "Process Sales Invoice Credit Sales Complete.... !!<br>"
		response.Flush
	end if
	'-- End Insert Credit Sales
	
	'-- Insert Credit PPN Keluaran
	qSalesInvCR1 = 	" SELECT salesinv.fcba," & _
					"        salesinv.invoice_by," & _
					"        salesinv.transdate," & _
					"        to_number(to_char(salesinv.transdate,'MM')) ACCOUNTINGPERIOD," & _
					"        to_number(to_char(salesinv.transdate,'YYYY')) ACCOUNTINGYEAR," & _
					"        salesinv.salescontractno," & _
					"        salesinv.customercode," & _
					"        salesinv.dp_vat," & _
					"        salesinv.amount_dpvat," & _
					"        salesinv.dp_aftervat," & _
					"        salesinv.percentage_dp," & _
					"        salesinv.invoicecode," & _
					"        salesinv.lineno," & _
					"        salesinv.productcode," & _
					"        pm.param_5 AS Jobcode_Sales," & _
					"        salesinvoice_beforeppn, " & _
					"        salesinv.ppn," & _
					"        salesinv.salesinvoice_ppn," & _
					"        salesinv.salesinvoice_afterppn," & _
					"        pm.param_3 AS Jobcode_UangMukaPenjualan," & _
					"        salesinv.Penyelesaian_DP, " & _
					"        pm.param_4 AS Jobcode_PPnKeluaran," & _
					"        ( -1 * salesinv.Penyelesaian_PPN) Penyelesaian_PPN, " & _
					"        cs.controljob as jobcode_piutangpenjualan," & _
					"        CASE" & _
					"           WHEN salesinv.penyelesaian_dp = 0" & _
					"           THEN" & _
					"              (salesinv.salesinvoice_beforeppn + salesinv.Penyelesaian_PPN)" & _
					"           ELSE" & _
					"              (  (salesinv.salesinvoice_beforeppn + salesinv.Penyelesaian_PPN)" & _
					"               - salesinv.Penyelesaian_DP)" & _
					"        END" & _
					"           AS Piutang_Penjualan " & _
					"   FROM (SELECT fcba," & _
					"                invoice_by," & _
					"                transdate," & _
					"                salescontractno," & _
					"                customercode," & _
					"                dp_vat," & _
					"                amount_dpvat," & _
					"                dp_aftervat," & _
					"                percentage_dp," & _
					"                invoicecode," & _
					"                lineno," & _
					"                productcode," & _
					"                salesinvoice_beforeppn," & _
					"                ppn," & _
					"                salesinvoice_ppn," & _
					"                salesinvoice_afterppn," & _
					"                Penyelesaian_DP," & _
					"                CASE" & _
					"                   WHEN (penyelesaian_dp = 0 AND dp_vat = 0)" & _
					"                   THEN" & _
					"                      salesinvoice_ppn" & _
					"                   WHEN (    penyelesaian_dp = salesinvoice_beforeppn" & _
					"                         AND dp_vat > 0)" & _
					"                   THEN" & _
					" 0" & _
					"                   WHEN (penyelesaian_dp > 0 AND dp_vat = 0)" & _
					"                   THEN" & _
					"                      salesinvoice_ppn" & _
					"                   WHEN (penyelesaian_dp > 0 AND dp_vat > 0)" & _
					"                   THEN" & _
					"                      ( (salesinvoice_ppn * percentage_dp) / 100)" & _
					"                   ELSE" & _
					" 0" & _
					"                END" & _
					"                   AS Penyelesaian_PPN" & _
					"           FROM (SELECT iplasprod.salesinvoicedtl.fcba," & _
					"                        iplasprod.salesinvoicehdr.invoice_by," & _
					"                        iplasprod.salesinvoicehdr.transdate," & _
					"                        iplasprod.salesagreementhdr.salescontractno," & _
					"                        iplasprod.salesagreementhdr.customercode," & _
					"                        iplasprod.salesagreementhdr.dp_vat," & _
					"                        iplasprod.salesagreementhdr.amount_dpvat," & _
					"                        iplasprod.salesagreementhdr.dp_aftervat," & _
					"                        iplasprod.salesagreementhdr.percentage_dp," & _
					"                        iplasprod.salesinvoicedtl.invoicecode," & _
					"                        iplasprod.salesinvoicedtl.lineno," & _
					"                        iplasprod.salesinvoicedtl.productcode," & _
					"                        iplasprod.salesinvoicedtl.amount" & _
					"                           AS salesinvoice_beforeppn," & _
					"                        iplasprod.salesinvoicehdr.vat AS ppn," & _
					"                        ROUND ( (  (  iplasprod.salesinvoicedtl.amount" & _
					"                            * iplasprod.salesinvoicehdr.vat)" & _
					"                         / 100),0)" & _
					"                           AS salesinvoice_ppn," & _
					"                        (  iplasprod.salesinvoicedtl.amount" & _
					"                         + (  (  iplasprod.salesinvoicedtl.amount" & _
					"                               * iplasprod.salesinvoicehdr.vat)" & _
					"                            / 100))" & _
					"                           AS salesinvoice_afterppn," & _
					"                        CASE" & _
					"                           WHEN iplasprod.salesagreementhdr.percentage_dp > 0" & _
					"                           THEN" & _
					"                              (  (  iplasprod.salesinvoicedtl.amount" & _
					"                                  * iplasprod.salesagreementhdr.percentage_dp)" & _
					"                               / 100)" & _
					"                           ELSE" & _
					" 0" & _
					"                        END" & _
					"                           AS Penyelesaian_DP" & _
					"                   FROM iplasprod.salesinvoicedtl" & _
					"                        INNER JOIN iplasprod.salesinvoicehdr" & _
					"                           ON (    iplasprod.salesinvoicedtl.fcba =" & _
					"                                      iplasprod.salesinvoicehdr.fcba" & _
					"                               AND iplasprod.salesinvoicedtl.invoicecode =" & _
					"                                      iplasprod.salesinvoicehdr.invoicecode)" & _
					"                        INNER JOIN iplasprod.salesagreementhdr" & _
					"                           ON (    iplasprod.salesinvoicehdr.fcba =" & _
					"                                      iplasprod.salesagreementhdr.fcba" & _
					"                               AND iplasprod.salesinvoicehdr.salescontractno =" & _
					"                                      iplasprod.salesagreementhdr.salescontractno)" & _
					"                  WHERE     TO_NUMBER (" & _
					"                               TO_CHAR (iplasprod.salesinvoicehdr.transdate," & _
					"                                        'mm')) =" & _
					"                               (SELECT currentaccountingperiod FROM iplasprod.periodcontrol" & _
					"                                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
					"                                       AND UPPER ( iplasprod.periodcontrol.SYSTEM) = 'SALES') " & _
					"                  AND     TO_NUMBER (" & _
					"                               TO_CHAR (iplasprod.salesinvoicehdr.transdate," & _
					"                                        'yyyy')) =" & _
					"                               (SELECT currentaccountingyear FROM iplasprod.periodcontrol" & _
					"                                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
					"                                       AND UPPER ( iplasprod.periodcontrol.SYSTEM) = 'SALES') " & _
					"                        AND iplasprod.salesinvoicehdr.invoice_by = '"& session("fcBA") &"' )) salesinv" & _
					"        INNER JOIN iplasprod.parameter pm" & _
					"           ON (salesinv.fcba = pm.fcba AND salesinv.productcode = pm.fccode)" & _
					"        INNER JOIN iplasprod.customer cs" & _
					"           ON (salesinv.fcba = cs.fcba AND salesinv.customercode = cs.customercode)" & _
					"  WHERE pm.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" 

	Set rsSalesInvCR1 = Server.CreateObject("ADODB.Recordset")
	rsSalesInvCR1.Open qSalesInvCR1, connstr
	if not rsSalesInvCR1.eof then
		Do While Not rsSalesInvCR1.eof	   
			' tmpAMountDB1 = Round(rsSalesInvCR1("AMOUNT_DB1"))
			maxLineNo = maxLineNo + 1

			' Target Type
			qJob_PPnKeluaran = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"' and fccode = '"& rsSalesInvCR1("jobcode_ppnkeluaran") &"'"
					
			Set rsPPnKeluaran = Server.CreateObject("ADODB.Recordset")
			rsPPnKeluaran.CursorType = 3
			rsPPnKeluaran.CursorLocation = 3
			rsPPnKeluaran.Open qJob_PPnKeluaran, connstr

			if not rsPPnKeluaran.eof then
				tt_job_category = rsPPnKeluaran("job_category") 
			end if
			rsPPnKeluaran.close
			set rsPPnKeluaran = Nothing
			' End Target Type

			' Target Code
			qJob_PPnKeluaran = 	" SELECT fccode" & _
								"   FROM iplasprod.gcmaster" & _
								"  WHERE     fcba = '"& session("fcba") &"'" & _
								"        AND item =" & _
								"               (SELECT job_category" & _
								"                  FROM iplasprod.job" & _
								"                 WHERE fcba = '"& session("fcBA") &"' AND fccode = '"& rsSalesInvCR1("jobcode_ppnkeluaran") &"')"  
					
			Set rsPPnKeluaran = Server.CreateObject("ADODB.Recordset")
			rsPPnKeluaran.CursorType = 3
			rsPPnKeluaran.CursorLocation = 3
			rsPPnKeluaran.Open qJob_PPnKeluaran, connstr

			if not rsPPnKeluaran.eof then
				tc_job_category = rsPPnKeluaran("fccode") 
			end if
			rsPPnKeluaran.close
			set rsPPnKeluaran = Nothing
			' End Target Code

			qCost_ActivityDB1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesInvCR1("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesInvCR1("jobcode_ppnkeluaran") & "','SL','SL','" & rsSalesInvCR1("invoicecode") & "'," & _
							 "'','" & tt_job_category & "','" & tc_job_category & "'," & _
							 "'" & rsSalesInvCR1("Penyelesaian_PPN") & "','" & rsSalesInvCR1("salescontractno") & "'," & "0,0,'" & rsSalesInvCR1("INVOICE_BY") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " 'TC - Sales Invoice Code : ', '"  & rsSalesInvCR1("invoicecode") & "'," & _
							 " 'Auto - Sales Contract No : ', '"  & rsSalesInvCR1("salescontractno") & "'," & _
							 " 'Auto - Customer Code: ', '"  & rsSalesInvCR1("customercode") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsSalesInvCR1("LINENO") & "','Cr', " & _
							 " " & rsSalesInvCR1("ACCOUNTINGPERIOD") & "," & rsSalesInvCR1("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDB1)
			 connStr.execute("Commit")
			 
			rsSalesInvCR1.MoveNext
		loop
		response.Write "Process Sales Invoice Credit PPn Keluaran Complete.... !!<br>"
		response.Flush
	end if
	'-- End Insert Credit PPN Keluaran
	
	'-- End sales Invoice

	' -- Begin Transporter Invoice
	' -- Insert Debet Biaya Transporter
		 qTransSalesInvDB1 = 	" SELECT iplasprod.transporterinvoicedtl.invoicecode," & _
								"        iplasprod.transporterinvoice.creationdate," & _
								"        to_number(to_char(transporterinvoice.creationdate,'MM')) ACCOUNTINGPERIOD," & _
								"        to_number(to_char(transporterinvoice.creationdate,'YYYY')) ACCOUNTINGYEAR," & _
								"        iplasprod.transporterinvoice.invoice_by," & _
								"        iplasprod.transporterinvoice.transportercontractno," & _
								"        iplasprod.transporter.fccode," & _
								"        iplasprod.transporter.fcname," & _
								"        iplasprod.transporterinvoicedtl.wbt_chitno," & _
								"        iplasprod.transporterinvoicedtl.lineno," & _
								"        iplasprod.transporterinvoice.vat," & _
								"        iplasprod.parameter.param_7 AS jobcode_biayatransporter," & _
								"        iplasprod.parameter.param_8 AS jobcode_accruedtransporter," & _
								"        iplasprod.transporterinvoicedtl.amountafterloss," & _
								"        iplasprod.parameter.param_8 AS jobcode_ppnmasukan," & _
								"        (  (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"            * iplasprod.transporterinvoice.vat)" & _
								"         / 100)" & _
								"           AS ppn_transporterinvoice," & _
								"        iplasprod.transporter.controljob AS jobcode_hutangtransporter," & _
								"        (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"         + (  (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"               * iplasprod.transporterinvoice.vat)" & _
								"            / 100))" & _
								"           AS hutang_transporter" & _
								"   FROM iplasprod.transporterinvoicedtl" & _
								"        INNER JOIN iplasprod.transporterinvoice" & _
								"           ON (    iplasprod.transporterinvoicedtl.fcba =" & _
								"                      iplasprod.transporterinvoice.fcba" & _
								"               AND iplasprod.transporterinvoicedtl.invoicecode =" & _
								"                      iplasprod.transporterinvoice.invoicecode)" & _
								"        INNER JOIN iplasprod.transportercontract" & _
								"           ON (    iplasprod.transporterinvoice.fcba =" & _
								"                      iplasprod.transportercontract.fcba" & _
								"               AND iplasprod.transporterinvoice.transportercontractno =" & _
								"                      iplasprod.transportercontract.transportercontractno)" & _
								" inner join iplasprod.salesshipmentdtl on " & _
								" (iplasprod.transporterinvoice.fcba = iplasprod.salesshipmentdtl.fcba and " & _
								" iplasprod.transporterinvoicedtl.wbt_chitno = iplasprod.salesshipmentdtl.wbt_chitno)" & _
								" inner join iplasprod.salesshipmenthdr on " & _
								" (iplasprod.salesshipmentdtl.fcba = iplasprod.salesshipmenthdr.fcba and " & _
								" iplasprod.salesshipmentdtl.salesshipmentno = iplasprod.salesshipmenthdr.salesshipmentno)" & _
								"        inner join (" & _
								"            select distinct " & _
								"            fcba," & _
								"            deliveryorderno," & _
								"            productcode," & _
								"            transportercontractno," & _
								"            transporterline," & _
								"            transporter " & _
								"            from deliveryorder" & _
								"        ) DO on (" & _
								"        iplasprod.transportercontract.fcba = DO.fcba and" & _
								"        iplasprod.transportercontract.transportercontractno = DO.transportercontractno and" & _
								"		 iplasprod.salesshipmenthdr.deliveryorderno = DO.deliveryorderno)" & _
								"        INNER JOIN iplasprod.transporter" & _
								"           ON (    iplasprod.transportercontract.fcba =" & _
								"                      iplasprod.transporter.fcba" & _
								"               AND iplasprod.transportercontract.transporter =" & _
								"                      iplasprod.transporter.fccode)" & _
								"        INNER JOIN iplasprod.parameter" & _
								"           ON (    DO.fcba = iplasprod.parameter.fcba" & _
								"               AND DO.productcode =" & _
								"                      iplasprod.parameter.fccode)" & _
								"  WHERE     TO_NUMBER (" & _
								"               TO_CHAR (iplasprod.transporterinvoice.creationdate, 'mm')) =" & _
								"               (SELECT currentaccountingperiod" & _
								"                  FROM iplasprod.periodcontrol" & _
								"                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcba") &"'" & _
								"                       AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"        AND TO_NUMBER (" & _
								"               TO_CHAR (iplasprod.transporterinvoice.creationdate, 'yyyy')) =" & _
								"               (SELECT currentaccountingyear" & _
								"                  FROM iplasprod.periodcontrol" & _
								"                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcba") &"'" & _
								"                       AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"        AND iplasprod.transporterinvoice.invoice_by = '"& session("fcba") &"'" & _
								"        AND iplasprod.parameter.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" 

			Set rsTransSalesInvDB1 = Server.CreateObject("ADODB.Recordset")
			rsTransSalesInvDB1.Open qTransSalesInvDB1, connstr
			if not rsTransSalesInvDB1.eof then
				Do While Not rsTransSalesInvDB1.eof	   
					' tmpAMountDB1 = Round(rsTransSalesInvDB1("AMOUNT_DB1"))
					maxLineNo = maxLineNo + 1

					' Target Type
					qJob_BiayaTransporter = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"' and fccode = '"& rsTransSalesInvDB1("jobcode_biayatransporter") &"'"
							
					Set rsBiayaTransporter = Server.CreateObject("ADODB.Recordset")
					rsBiayaTransporter.CursorType = 3
					rsBiayaTransporter.CursorLocation = 3
					rsBiayaTransporter.Open qJob_BiayaTransporter, connstr

					if not rsBiayaTransporter.eof then
						tt_job_category = rsBiayaTransporter("job_category") 
					end if
					rsBiayaTransporter.close
					set rsBiayaTransporter = Nothing
					' End Target Type

					' Target Code
					qJob_BiayaTransporter = 	" SELECT fccode" & _
												"   FROM iplasprod.gcmaster" & _
												"  WHERE     fcba = '"& session("fcba") &"'" & _
												"        AND item =" & _
												"               (SELECT job_category" & _
												"                  FROM iplasprod.job" & _
												"                 WHERE fcba = '"& session("fcBA") &"' AND fccode = '"& rsTransSalesInvDB1("jobcode_biayatransporter") &"')"  
							
					Set rsBiayaTransporter = Server.CreateObject("ADODB.Recordset")
					rsBiayaTransporter.CursorType = 3
					rsBiayaTransporter.CursorLocation = 3
					rsBiayaTransporter.Open qJob_BiayaTransporter, connstr

					if not rsBiayaTransporter.eof then
						tc_job_category = rsBiayaTransporter("fccode") 
					end if
					rsBiayaTransporter.close
					set rsBiayaTransporter = Nothing
					' End Target Code

					qCost_ActivityDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsTransSalesInvDB1("creationdate") & "','mm/dd/yyyy') " & ",'" & rsTransSalesInvDB1("jobcode_biayatransporter") & "','SL','SL','" & rsTransSalesInvDB1("invoicecode") & "'," & _
									 "'','" & tt_job_category & "','" & tc_job_category & "'," & _
									 "'" & rsTransSalesInvDB1("amountafterloss") & "','" & rsTransSalesInvDB1("wbt_chitno") & "'," & "0,0,'" & rsTransSalesInvDB1("INVOICE_BY") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'TC - Transporter Invoice No : ', '"  & rsTransSalesInvDB1("invoicecode") & "'," & _
									 " 'Auto - Transporter Contract No : ', '"  & rsTransSalesInvDB1("transportercontractno") & "'," & _
									 " 'Auto - Chitno : ', '"  & rsTransSalesInvDB1("wbt_chitno") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsTransSalesInvDB1("LINENO") & "','Db', " & _
									 " " & rsTransSalesInvDB1("ACCOUNTINGPERIOD") & "," & rsTransSalesInvDB1("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "

					 connstr.execute(qCost_ActivityDB1)
					 connStr.execute("Commit")
					 
					rsTransSalesInvDB1.MoveNext
				loop
				response.Write "Process Transporter Invoice Debet Biaya Transporter Complete.... !!<br>"
				response.Flush
			End if

	' -- End Debet Biaya Transporter

	' -- Insert Debet PPN Masukan
		 qTransSalesInvDB1 = 	" SELECT iplasprod.transporterinvoicedtl.invoicecode," & _
								"        iplasprod.transporterinvoice.creationdate," & _
								"        to_number(to_char(transporterinvoice.creationdate,'MM')) ACCOUNTINGPERIOD," & _
								"        to_number(to_char(transporterinvoice.creationdate,'YYYY')) ACCOUNTINGYEAR," & _
								"        iplasprod.transporterinvoice.invoice_by," & _
								"        iplasprod.transporterinvoice.transportercontractno," & _
								"        iplasprod.transporter.fccode," & _
								"        iplasprod.transporter.fcname," & _
								"        iplasprod.transporterinvoicedtl.wbt_chitno," & _
								"        iplasprod.transporterinvoicedtl.lineno," & _
								"        iplasprod.transporterinvoice.vat," & _
								"        iplasprod.parameter.param_7 AS jobcode_biayatransporter," & _
								"        iplasprod.parameter.param_8 AS jobcode_accruedtransporter," & _
								"        iplasprod.transporterinvoicedtl.amountafterloss," & _
								"        iplasprod.parameter.param_9 AS jobcode_ppnmasukan," & _
								"        (  (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"            * iplasprod.transporterinvoice.vat)" & _
								"         / 100)" & _
								"           AS ppn_transporterinvoice," & _
								"        iplasprod.transporter.controljob AS jobcode_hutangtransporter," & _
								"        (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"         + (  (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"               * iplasprod.transporterinvoice.vat)" & _
								"            / 100))" & _
								"           AS hutang_transporter" & _
								"   FROM iplasprod.transporterinvoicedtl" & _
								"        INNER JOIN iplasprod.transporterinvoice" & _
								"           ON (    iplasprod.transporterinvoicedtl.fcba =" & _
								"                      iplasprod.transporterinvoice.fcba" & _
								"               AND iplasprod.transporterinvoicedtl.invoicecode =" & _
								"                      iplasprod.transporterinvoice.invoicecode)" & _
								"        INNER JOIN iplasprod.transportercontract" & _
								"           ON (    iplasprod.transporterinvoice.fcba =" & _
								"                      iplasprod.transportercontract.fcba" & _
								"               AND iplasprod.transporterinvoice.transportercontractno =" & _
								"                      iplasprod.transportercontract.transportercontractno)" & _
								" inner join iplasprod.salesshipmentdtl on " & _
								" (iplasprod.transporterinvoice.fcba = iplasprod.salesshipmentdtl.fcba and " & _
								" iplasprod.transporterinvoicedtl.wbt_chitno = iplasprod.salesshipmentdtl.wbt_chitno)" & _
								" inner join iplasprod.salesshipmenthdr on " & _
								" (iplasprod.salesshipmentdtl.fcba = iplasprod.salesshipmenthdr.fcba and " & _
								" iplasprod.salesshipmentdtl.salesshipmentno = iplasprod.salesshipmenthdr.salesshipmentno)" & _
								"        inner join (" & _
								"            select distinct " & _
								"            fcba," & _
								"            deliveryorderno," & _
								"            productcode," & _
								"            transportercontractno," & _
								"            transporterline," & _
								"            transporter " & _
								"            from deliveryorder" & _
								"        ) DO on (" & _
								"        iplasprod.transportercontract.fcba = DO.fcba and" & _
								"        iplasprod.transportercontract.transportercontractno = DO.transportercontractno and" & _
								"		 iplasprod.salesshipmenthdr.deliveryorderno = DO.deliveryorderno)" & _
								"        INNER JOIN iplasprod.transporter" & _
								"           ON (    iplasprod.transportercontract.fcba =" & _
								"                      iplasprod.transporter.fcba" & _
								"               AND iplasprod.transportercontract.transporter =" & _
								"                      iplasprod.transporter.fccode)" & _
								"        INNER JOIN iplasprod.parameter" & _
								"           ON (    DO.fcba = iplasprod.parameter.fcba" & _
								"               AND DO.productcode =" & _
								"                      iplasprod.parameter.fccode)" & _
								"  WHERE     TO_NUMBER (" & _
								"               TO_CHAR (iplasprod.transporterinvoice.creationdate, 'mm')) =" & _
								"               (SELECT currentaccountingperiod" & _
								"                  FROM iplasprod.periodcontrol" & _
								"                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcba") &"'" & _
								"                       AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"        AND TO_NUMBER (" & _
								"               TO_CHAR (iplasprod.transporterinvoice.creationdate, 'yyyy')) =" & _
								"               (SELECT currentaccountingyear" & _
								"                  FROM iplasprod.periodcontrol" & _
								"                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcba") &"'" & _
								"                       AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"        AND iplasprod.transporterinvoice.invoice_by = '"& session("fcba") &"'" & _
								"        AND iplasprod.parameter.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" 

			Set rsTransSalesInvDB1 = Server.CreateObject("ADODB.Recordset")
			rsTransSalesInvDB1.Open qTransSalesInvDB1, connstr
			if not rsTransSalesInvDB1.eof then
				Do While Not rsTransSalesInvDB1.eof	   
					' tmpAMountDB1 = Round(rsTransSalesInvDB1("AMOUNT_DB1"))
					maxLineNo = maxLineNo + 1

					' Target Type
					qJob_BiayaTransporter = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"' and fccode = '"& rsTransSalesInvDB1("jobcode_ppnmasukan") &"'"
							
					Set rsBiayaTransporter = Server.CreateObject("ADODB.Recordset")
					rsBiayaTransporter.CursorType = 3
					rsBiayaTransporter.CursorLocation = 3
					rsBiayaTransporter.Open qJob_BiayaTransporter, connstr

					if not rsBiayaTransporter.eof then
						tt_job_category = rsBiayaTransporter("job_category") 
					end if
					rsBiayaTransporter.close
					set rsBiayaTransporter = Nothing
					' End Target Type

					' Target Code
					qJob_BiayaTransporter = 	" SELECT fccode" & _
												"   FROM iplasprod.gcmaster" & _
												"  WHERE     fcba = '"& session("fcba") &"'" & _
												"        AND item =" & _
												"               (SELECT job_category" & _
												"                  FROM iplasprod.job" & _
												"                 WHERE fcba = '"& session("fcBA") &"' AND fccode = '"& rsTransSalesInvDB1("jobcode_ppnmasukan") &"')"  
							
					Set rsBiayaTransporter = Server.CreateObject("ADODB.Recordset")
					rsBiayaTransporter.CursorType = 3
					rsBiayaTransporter.CursorLocation = 3
					rsBiayaTransporter.Open qJob_BiayaTransporter, connstr

					if not rsBiayaTransporter.eof then
						tc_job_category = rsBiayaTransporter("fccode") 
					end if
					rsBiayaTransporter.close
					set rsBiayaTransporter = Nothing
					' End Target Code

					qCost_ActivityDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsTransSalesInvDB1("creationdate") & "','mm/dd/yyyy') " & ",'" & rsTransSalesInvDB1("jobcode_ppnmasukan") & "','SL','SL','" & rsTransSalesInvDB1("invoicecode") & "'," & _
									 "'','" & tt_job_category & "','" & tc_job_category & "'," & _
									 "'" & rsTransSalesInvDB1("ppn_transporterinvoice") & "','" & rsTransSalesInvDB1("wbt_chitno") & "'," & "0,0,'" & rsTransSalesInvDB1("INVOICE_BY") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'TC - Transporter Invoice No : ', '"  & rsTransSalesInvDB1("invoicecode") & "'," & _
									 " 'Auto - Transporter Contract No : ', '"  & rsTransSalesInvDB1("transportercontractno") & "'," & _
									 " 'Auto - Chitno : ', '"  & rsTransSalesInvDB1("wbt_chitno") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsTransSalesInvDB1("LINENO") & "','Db', " & _
									 " " & rsTransSalesInvDB1("ACCOUNTINGPERIOD") & "," & rsTransSalesInvDB1("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "

					 connstr.execute(qCost_ActivityDB1)
					 connStr.execute("Commit")
					 
					rsTransSalesInvDB1.MoveNext
				loop
				response.Write "Process Transporter Invoice Debet PPn Masukan Complete.... !!<br>"
				response.Flush
			End if
	' -- End Debet PPN Masukan

	' -- Insert Credit Hutang  Transporter
		 qTransSalesInvDB1 = 	" SELECT iplasprod.transporterinvoicedtl.invoicecode," & _
								"        iplasprod.transporterinvoice.creationdate," & _
								"        to_number(to_char(transporterinvoice.creationdate,'MM')) ACCOUNTINGPERIOD," & _
								"        to_number(to_char(transporterinvoice.creationdate,'YYYY')) ACCOUNTINGYEAR," & _
								"        iplasprod.transporterinvoice.invoice_by," & _
								"        iplasprod.transporterinvoice.transportercontractno," & _
								"        iplasprod.transporter.fccode," & _
								"        iplasprod.transporter.fcname," & _
								"        iplasprod.transporterinvoicedtl.wbt_chitno," & _
								"        iplasprod.transporterinvoicedtl.lineno," & _
								"        iplasprod.transporterinvoice.vat," & _
								"        iplasprod.parameter.param_7 AS jobcode_biayatransporter," & _
								"        iplasprod.parameter.param_8 AS jobcode_accruedtransporter," & _
								"        iplasprod.transporterinvoicedtl.amountafterloss," & _
								"        iplasprod.parameter.param_9 AS jobcode_ppnmasukan," & _
								"        (  (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"            * iplasprod.transporterinvoice.vat)" & _
								"         / 100)" & _
								"           AS ppn_transporterinvoice," & _
								"        iplasprod.transporter.controljob AS jobcode_hutangtransporter," & _
								"        (-1 * (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"         + (  (  iplasprod.transporterinvoicedtl.amountafterloss" & _
								"               * iplasprod.transporterinvoice.vat)" & _
								"            / 100)))" & _
								"           AS hutang_transporter" & _
								"   FROM iplasprod.transporterinvoicedtl" & _
								"        INNER JOIN iplasprod.transporterinvoice" & _
								"           ON (    iplasprod.transporterinvoicedtl.fcba =" & _
								"                      iplasprod.transporterinvoice.fcba" & _
								"               AND iplasprod.transporterinvoicedtl.invoicecode =" & _
								"                      iplasprod.transporterinvoice.invoicecode)" & _
								"        INNER JOIN iplasprod.transportercontract" & _
								"           ON (    iplasprod.transporterinvoice.fcba =" & _
								"                      iplasprod.transportercontract.fcba" & _
								"               AND iplasprod.transporterinvoice.transportercontractno =" & _
								"                      iplasprod.transportercontract.transportercontractno)" & _
								" inner join iplasprod.salesshipmentdtl on " & _
								" (iplasprod.transporterinvoice.fcba = iplasprod.salesshipmentdtl.fcba and " & _
								" iplasprod.transporterinvoicedtl.wbt_chitno = iplasprod.salesshipmentdtl.wbt_chitno)" & _
								" inner join iplasprod.salesshipmenthdr on " & _
								" (iplasprod.salesshipmentdtl.fcba = iplasprod.salesshipmenthdr.fcba and " & _
								" iplasprod.salesshipmentdtl.salesshipmentno = iplasprod.salesshipmenthdr.salesshipmentno)" & _
								"        inner join (" & _
								"            select distinct " & _
								"            fcba," & _
								"            deliveryorderno," & _
								"            productcode," & _
								"            transportercontractno," & _
								"            transporterline," & _
								"            transporter " & _
								"            from deliveryorder" & _
								"        ) DO on (" & _
								"        iplasprod.transportercontract.fcba = DO.fcba and" & _
								"        iplasprod.transportercontract.transportercontractno = DO.transportercontractno and" & _
								"		 iplasprod.salesshipmenthdr.deliveryorderno = DO.deliveryorderno)" & _
								"        INNER JOIN iplasprod.transporter" & _
								"           ON (    iplasprod.transportercontract.fcba =" & _
								"                      iplasprod.transporter.fcba" & _
								"               AND iplasprod.transportercontract.transporter =" & _
								"                      iplasprod.transporter.fccode)" & _
								"        INNER JOIN iplasprod.parameter" & _
								"           ON (    DO.fcba = iplasprod.parameter.fcba" & _
								"               AND DO.productcode =" & _
								"                      iplasprod.parameter.fccode)" & _
								"  WHERE     TO_NUMBER (" & _
								"               TO_CHAR (iplasprod.transporterinvoice.creationdate, 'mm')) =" & _
								"               (SELECT currentaccountingperiod" & _
								"                  FROM iplasprod.periodcontrol" & _
								"                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcba") &"'" & _
								"                       AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"        AND TO_NUMBER (" & _
								"               TO_CHAR (iplasprod.transporterinvoice.creationdate, 'yyyy')) =" & _
								"               (SELECT currentaccountingyear" & _
								"                  FROM iplasprod.periodcontrol" & _
								"                 WHERE     iplasprod.periodcontrol.fcba = '"& session("fcba") &"'" & _
								"                       AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"        AND iplasprod.transporterinvoice.invoice_by = '"& session("fcba") &"'" & _
								"        AND iplasprod.parameter.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" 

			Set rsTransSalesInvDB1 = Server.CreateObject("ADODB.Recordset")
			rsTransSalesInvDB1.Open qTransSalesInvDB1, connstr
			if not rsTransSalesInvDB1.eof then
				Do While Not rsTransSalesInvDB1.eof	   
					' tmpAMountDB1 = Round(rsTransSalesInvDB1("AMOUNT_DB1"))
					maxLineNo = maxLineNo + 1

					' Target Type
					qJob_BiayaTransporter = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"' and fccode = '"& rsTransSalesInvDB1("jobcode_hutangtransporter") &"'"
							
					Set rsBiayaTransporter = Server.CreateObject("ADODB.Recordset")
					rsBiayaTransporter.CursorType = 3
					rsBiayaTransporter.CursorLocation = 3
					rsBiayaTransporter.Open qJob_BiayaTransporter, connstr

					if not rsBiayaTransporter.eof then
						tt_job_category = rsBiayaTransporter("job_category")
					end if
					rsBiayaTransporter.close
					set rsBiayaTransporter = Nothing
					' End Target Type

					 
					qCost_ActivityDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsTransSalesInvDB1("creationdate") & "','mm/dd/yyyy') " & ",'" & rsTransSalesInvDB1("jobcode_hutangtransporter") & "','SL','SL','" & rsTransSalesInvDB1("invoicecode") & "'," & _
									 "'','" & tt_job_category & "','" & rsTransSalesInvDB1("fccode") & "'," & _
									 "'" & rsTransSalesInvDB1("hutang_transporter") & "','" & rsTransSalesInvDB1("wbt_chitno") & "'," & "0,0,'" & rsTransSalesInvDB1("INVOICE_BY") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'TC - Transporter Invoice No : ', '"  & rsTransSalesInvDB1("invoicecode") & "'," & _
									 " 'Auto - Transporter Contract No : ', '"  & rsTransSalesInvDB1("transportercontractno") & "'," & _
									 " 'Auto - Chitno : ', '"  & rsTransSalesInvDB1("wbt_chitno") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsTransSalesInvDB1("LINENO") & "','Cr', " & _
									 " " & rsTransSalesInvDB1("ACCOUNTINGPERIOD") & "," & rsTransSalesInvDB1("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "

					 connstr.execute(qCost_ActivityDB1)
					 connStr.execute("Commit")
					 
					rsTransSalesInvDB1.MoveNext
				loop
				response.Write "Process Credit Hutang Transporter Complete.... !!<br>"
				response.Flush
			end if

	' -- End Credit Hutang Transporter

	' -- End Transporter Invoice

	' -- Begin Jurnal Penalty
	' -- Insert Debet Jurnal Penalty 
		 qTransSalesInvDB1 = 	"  SELECT iplasprod.transporterinvoicedtl.invoicecode," & _
								"          iplasprod.transporterinvoice.transportercontractno," & _
								"          iplasprod.transporterinvoice.invoice_by," & _
								"          iplasprod.transporterinvoice.creationdate," & _
								"          to_number(to_char(transporterinvoice.creationdate,'MM')) ACCOUNTINGPERIOD," & _
								"          to_number(to_char(transporterinvoice.creationdate,'YYYY')) ACCOUNTINGYEAR," & _
								"          iplasprod.transporterinvoice.invoice_by," & _
								"          iplasprod.transporter.fccode," & _
								"          iplasprod.transporter.fcname," & _
								"          iplasprod.transporterinvoicedtl.salesshipmentno," & _
								"          iplasprod.transporterinvoicedtl.wbt_chitno," & _	
								"          iplasprod.parameter.param_7 AS jobcode_biayatransporter," & _
								"          iplasprod.transporter.controljob AS jobcode_hutangtransporter," & _
								"          SUM (iplasprod.transporterinvoicedtl.penaltyamount) penaltyamount" & _
								"     FROM iplasprod.transporterinvoicedtl" & _
								"          INNER JOIN iplasprod.transporterinvoice" & _
								"             ON (    iplasprod.transporterinvoicedtl.fcba =" & _
								"                        iplasprod.transporterinvoice.fcba" & _
								"                 AND iplasprod.transporterinvoicedtl.invoicecode =" & _
								"                        iplasprod.transporterinvoice.invoicecode)" & _
								"          INNER JOIN iplasprod.transportercontract" & _
								"             ON (    iplasprod.transporterinvoice.fcba =" & _
								"                        iplasprod.transportercontract.fcba" & _
								"                 AND iplasprod.transporterinvoice.transportercontractno =" & _
								"                        iplasprod.transportercontract.transportercontractno)" & _
								" inner join iplasprod.salesshipmentdtl on " & _
								" (iplasprod.transporterinvoice.fcba = iplasprod.salesshipmentdtl.fcba and " & _
								" iplasprod.transporterinvoicedtl.wbt_chitno = iplasprod.salesshipmentdtl.wbt_chitno)" & _
								" inner join iplasprod.salesshipmenthdr on " & _
								" (iplasprod.salesshipmentdtl.fcba = iplasprod.salesshipmenthdr.fcba and " & _
								" iplasprod.salesshipmentdtl.salesshipmentno = iplasprod.salesshipmenthdr.salesshipmentno)" & _
								"          INNER JOIN iplasprod.deliveryorder" & _
								"             ON (    iplasprod.transportercontract.fcba =" & _
								"                        iplasprod.deliveryorder.fcba" & _
								"                 AND iplasprod.transportercontract.transportercontractno =" & _
								"                        iplasprod.deliveryorder.transportercontractno)" & _
								"          INNER JOIN iplasprod.transporter" & _
								"             ON (    iplasprod.transportercontract.fcba =" & _
								"                        iplasprod.transporter.fcba" & _
								"                 AND iplasprod.transportercontract.transporter =" & _
								"                        iplasprod.transporter.fccode)" & _
								"          INNER JOIN iplasprod.parameter" & _
								"             ON (    iplasprod.deliveryorder.fcba = iplasprod.parameter.fcba" & _
								"                 AND iplasprod.deliveryorder.productcode =" & _
								"                        iplasprod.parameter.fccode)" & _
								"    WHERE     TO_NUMBER (" & _
								"                 TO_CHAR (iplasprod.transporterinvoice.creationdate, 'mm')) =" & _
								"                 (SELECT currentaccountingperiod " & _
								"                    FROM iplasprod.periodcontrol" & _
								"                   WHERE     iplasprod.periodcontrol.fcba = '"& session("fcba") &"'" & _
								"                         AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"          AND TO_NUMBER (" & _
								"                 TO_CHAR (iplasprod.transporterinvoice.creationdate, 'yyyy')) =" & _
								"                 (SELECT currentaccountingyear" & _
								"                    FROM iplasprod.periodcontrol" & _
								"                   WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
								"                         AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"          AND iplasprod.transporterinvoice.invoice_by = '"& session("fcba") &"'" & _
								"          AND iplasprod.parameter.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
								" GROUP BY iplasprod.transporterinvoicedtl.invoicecode," & _
								"          iplasprod.transporterinvoice.transportercontractno," & _
								"          iplasprod.transporter.fccode," & _
								"          iplasprod.transporter.fcname," & _
								"          iplasprod.transporterinvoicedtl.salesshipmentno," & _
								"          iplasprod.transporterinvoicedtl.wbt_chitno," & _	
								"          iplasprod.parameter.param_7," & _
								"          iplasprod.transporter.controljob," & _
								"          iplasprod.transporterinvoice.creationdate," & _
								"          to_number(to_char(transporterinvoice.creationdate,'MM')) ," & _
								"          to_number(to_char(transporterinvoice.creationdate,'YYYY')), " & _
								"          iplasprod.transporterinvoice.invoice_by" 

			Set rsTransSalesInvDB1 = Server.CreateObject("ADODB.Recordset")
			rsTransSalesInvDB1.Open qTransSalesInvDB1, connstr
			if not rsTransSalesInvDB1.eof then
				Do While Not rsTransSalesInvDB1.eof	   
					' tmpAMountDB1 = Round(rsTransSalesInvDB1("AMOUNT_DB1"))
					maxLineNo = maxLineNo + 1

					' Target Type
					qJob_HutangTransporter = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"'" & _
											 " and fccode = '"& rsTransSalesInvDB1("jobcode_hutangtransporter") &"'"
							
					Set rsHutangTransporter = Server.CreateObject("ADODB.Recordset")
					rsHutangTransporter.CursorType = 3
					rsHutangTransporter.CursorLocation = 3
					rsHutangTransporter.Open qJob_HutangTransporter, connstr

					if not rsHutangTransporter.eof then
						tt_job_category = rsHutangTransporter("job_category") 
					end if
					rsHutangTransporter.close
					set rsHutangTransporter = Nothing
					' End Target Type
 
					qCost_ActivityDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsTransSalesInvDB1("creationdate") & "','mm/dd/yyyy') " & ",'" & rsTransSalesInvDB1("jobcode_hutangtransporter") & "','SL','SL','" & rsTransSalesInvDB1("salesshipmentno") & "'," & _
									 "'','" & tt_job_category & "','" & rsTransSalesInvDB1("transportercontractno") & "'," & _
									 "'" & rsTransSalesInvDB1("penaltyamount") & "','" & rsTransSalesInvDB1("wbt_chitno") & "'," & "0,0,'" & rsTransSalesInvDB1("INVOICE_BY") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'TC - Sales Shipment No : ', '"  & rsTransSalesInvDB1("invoicecode") & "'," & _ 
									 " 'Auto - Chitno : ', '"  & rsTransSalesInvDB1("wbt_chitno") & "'," & _
									 " '','', " & _
									 " '','', " & _
									 " '', '','Db', " & _
									 " " & rsTransSalesInvDB1("ACCOUNTINGPERIOD") & "," & rsTransSalesInvDB1("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "

					 connstr.execute(qCost_ActivityDB1)
					 connStr.execute("Commit")
					 
					rsTransSalesInvDB1.MoveNext
				loop
				response.Write "Process Debet Jurnal Penalaty Transporter Complete.... !!<br>"
				response.Flush
			End if
	' -- End Debet Jurnal Penalty

	' -- Insert Credit Jurnal Penalty 
		 qTransSalesInvDB1 = 	"  SELECT iplasprod.transporterinvoicedtl.invoicecode," & _
								"          iplasprod.transporterinvoice.transportercontractno," & _
								"          iplasprod.transporterinvoice.creationdate," & _
								"          to_number(to_char(transporterinvoice.creationdate,'MM')) ACCOUNTINGPERIOD," & _
								"          to_number(to_char(transporterinvoice.creationdate,'YYYY')) ACCOUNTINGYEAR," & _
								"          iplasprod.transporterinvoice.invoice_by," & _
								"          iplasprod.transporter.fccode," & _
								"          iplasprod.transporter.fcname," & _
								"          iplasprod.transporterinvoicedtl.salesshipmentno," & _
								"          iplasprod.transporterinvoicedtl.wbt_chitno," & _	
								"          iplasprod.parameter.param_7 AS jobcode_biayatransporter," & _
								"          iplasprod.transporter.controljob AS jobcode_hutangtransporter," & _
								"          SUM ((-1 *iplasprod.transporterinvoicedtl.penaltyamount)) penaltyamount" & _
								"     FROM iplasprod.transporterinvoicedtl" & _
								"          INNER JOIN iplasprod.transporterinvoice" & _
								"             ON (    iplasprod.transporterinvoicedtl.fcba =" & _
								"                        iplasprod.transporterinvoice.fcba" & _
								"                 AND iplasprod.transporterinvoicedtl.invoicecode =" & _
								"                        iplasprod.transporterinvoice.invoicecode)" & _
								"          INNER JOIN iplasprod.transportercontract" & _
								"             ON (    iplasprod.transporterinvoice.fcba =" & _
								"                        iplasprod.transportercontract.fcba" & _
								"                 AND iplasprod.transporterinvoice.transportercontractno =" & _
								"                        iplasprod.transportercontract.transportercontractno)" & _
								" inner join iplasprod.salesshipmentdtl on " & _
								" (iplasprod.transporterinvoice.fcba = iplasprod.salesshipmentdtl.fcba and " & _
								" iplasprod.transporterinvoicedtl.wbt_chitno = iplasprod.salesshipmentdtl.wbt_chitno)" & _
								" inner join iplasprod.salesshipmenthdr on " & _
								" (iplasprod.salesshipmentdtl.fcba = iplasprod.salesshipmenthdr.fcba and " & _
								" iplasprod.salesshipmentdtl.salesshipmentno = iplasprod.salesshipmenthdr.salesshipmentno)" & _
								"          INNER JOIN iplasprod.deliveryorder" & _
								"             ON (    iplasprod.transportercontract.fcba =" & _
								"                        iplasprod.deliveryorder.fcba" & _
								"                 AND iplasprod.transportercontract.transportercontractno =" & _
								"                        iplasprod.deliveryorder.transportercontractno)" & _
								"          INNER JOIN iplasprod.transporter" & _
								"             ON (    iplasprod.transportercontract.fcba =" & _
								"                        iplasprod.transporter.fcba" & _
								"                 AND iplasprod.transportercontract.transporter =" & _
								"                        iplasprod.transporter.fccode)" & _
								"          INNER JOIN iplasprod.parameter" & _
								"             ON (    iplasprod.deliveryorder.fcba = iplasprod.parameter.fcba" & _
								"                 AND iplasprod.deliveryorder.productcode =" & _
								"                        iplasprod.parameter.fccode)" & _
								"    WHERE     TO_NUMBER (" & _
								"                 TO_CHAR (iplasprod.transporterinvoice.creationdate, 'mm')) =" & _
								"                 (SELECT currentaccountingperiod " & _
								"                    FROM iplasprod.periodcontrol" & _
								"                   WHERE     iplasprod.periodcontrol.fcba = '"& session("fcba") &"'" & _
								"                         AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"          AND TO_NUMBER (" & _
								"                 TO_CHAR (iplasprod.transporterinvoice.creationdate, 'yyyy')) =" & _
								"                 (SELECT currentaccountingyear" & _
								"                    FROM iplasprod.periodcontrol" & _
								"                   WHERE     iplasprod.periodcontrol.fcba = '"& session("fcBA") &"'" & _
								"                         AND UPPER (iplasprod.periodcontrol.SYSTEM) = 'SALES')" & _
								"          AND iplasprod.transporterinvoice.invoice_by = '"& session("fcba") &"'" & _
								"          AND iplasprod.parameter.parcode = 'JOBCODE_HARGAPOKOKPRODUKSI'" & _
								" GROUP BY iplasprod.transporterinvoicedtl.invoicecode," & _
								"          iplasprod.transporterinvoice.transportercontractno," & _
								"          iplasprod.transporter.fccode," & _
								"          iplasprod.transporter.fcname," & _
								"          iplasprod.transporterinvoicedtl.salesshipmentno," & _
								"          iplasprod.transporterinvoicedtl.wbt_chitno," & _	
								"          iplasprod.parameter.param_7," & _
								"          iplasprod.transporter.controljob," & _
								"          iplasprod.transporterinvoice.creationdate," & _
								"          to_number(to_char(transporterinvoice.creationdate,'MM')) ," & _
								"          to_number(to_char(transporterinvoice.creationdate,'YYYY')), " & _
								"          iplasprod.transporterinvoice.invoice_by" 

			Set rsTransSalesInvDB1 = Server.CreateObject("ADODB.Recordset")
			rsTransSalesInvDB1.Open qTransSalesInvDB1, connstr
			if not rsTransSalesInvDB1.eof then
				Do While Not rsTransSalesInvDB1.eof	   
					' tmpAMountDB1 = Round(rsTransSalesInvDB1("AMOUNT_DB1"))
					maxLineNo = maxLineNo + 1

					' Target Type
					qJob_HutangTransporter = " SELECT job_category FROM iplasprod.job where fcba = '"& session("fcBA") &"'" & _
											 " and fccode = '"& rsTransSalesInvDB1("jobcode_biayatransporter") &"'"
							
					Set rsHutangTransporter = Server.CreateObject("ADODB.Recordset")
					rsHutangTransporter.CursorType = 3
					rsHutangTransporter.CursorLocation = 3
					rsHutangTransporter.Open qJob_HutangTransporter, connstr

					if not rsHutangTransporter.eof then
						tt_job_category = rsHutangTransporter("job_category") 
					end if
					rsHutangTransporter.close
					set rsHutangTransporter = Nothing
					' End Target Type

					' Target Code
					qJob_BiayaTransporter = 	" SELECT fccode" & _
												"   FROM iplasprod.gcmaster" & _
												"  WHERE     fcba = '"& session("fcba") &"'" & _
												"        AND item =" & _
												"               (SELECT job_category" & _
												"                  FROM iplasprod.job" & _
												"                 WHERE fcba = '"& session("fcBA") &"' AND fccode = '"& rsTransSalesInvDB1("jobcode_biayatransporter") &"')"  
							
					Set rsBiayaTransporter = Server.CreateObject("ADODB.Recordset")
					rsBiayaTransporter.CursorType = 3
					rsBiayaTransporter.CursorLocation = 3
					rsBiayaTransporter.Open qJob_BiayaTransporter, connstr

					if not rsBiayaTransporter.eof then
						tc_job_category = rsBiayaTransporter("fccode")
					end if
					rsBiayaTransporter.close
					set rsBiayaTransporter = Nothing
					' End Target Code
 
					qCost_ActivityDB1="Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsTransSalesInvDB1("creationdate") & "','mm/dd/yyyy') " & ",'" & rsTransSalesInvDB1("jobcode_biayatransporter") & "','SL','SL','" & rsTransSalesInvDB1("salesshipmentno") & "'," & _
									 "'','" & tt_job_category & "','" & tc_job_category & "'," & _
									 "'" & rsTransSalesInvDB1("penaltyamount") & "','" & rsTransSalesInvDB1("wbt_chitno") & "'," & "0,0,'" & rsTransSalesInvDB1("INVOICE_BY") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " 'TC - Sales Shipment No : ', '"  & rsTransSalesInvDB1("invoicecode") & "'," & _ 
									 " 'Auto - Chitno : ', '"  & rsTransSalesInvDB1("wbt_chitno") & "'," & _
									 " '','', " & _
									 " '','', " & _
									 " '', '','Cr', " & _
									 " " & rsTransSalesInvDB1("ACCOUNTINGPERIOD") & "," & rsTransSalesInvDB1("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "

					 connstr.execute(qCost_ActivityDB1)
					 connStr.execute("Commit")
					 
					rsTransSalesInvDB1.MoveNext
				loop
				response.Write "Process Credit Jurnal Penalaty Transporter Complete.... !!<br>"
				response.Flush
			end if
	' -- End Credit PPN Masukan
	' -- End Jurnal Penalty
	end if
	
	if Session("Estate_Mill") = "E" then
	'-- Begin Biaya Transporter
	'-- Insert Debet
		qTransporterDB =" SELECT" &_
				" FCBA, " &_
				" ROWNUM LINENO," &_
				" ACCOUNTINGPERIOD," &_
				" ACCOUNTINGYEAR, " &_
				" TDATE_DB, " &_
				" JOBCODE_DB, " &_
				" SOURCE_DB, " &_
				" SOURCETYPE_DB, " &_
				" SOURCECODE_DB, " &_
				" TARGETTYPE_DB, " &_
				" TARGETCODE_DB, " &_
				" DOCUMENTNO_DB, " &_
				" QUANTITY_DB, " &_
				" OUTPUT_DB, " &_
				" AMOUNT_DB," &_
				" TITLE_REFERENCE_01_DB, " &_
				" AUTO_REFERENCE_01_DB, " &_
				" TITLE_REFERENCE_02_DB, " &_
				" AUTO_REFERENCE_02_DB, " &_
				" TITLE_REFERENCE_03_DB, " &_
				" AUTO_REFERENCE_03_DB" &_
				" FROM" &_
				" (" &_
				" SELECT " &_
				" MAX(DTL.FCBA)FCBA, " &_
				" MAX(DTL.ACCOUNTINGPERIOD)ACCOUNTINGPERIOD, " &_
				" MAX(DTL.ACCOUNTINGYEAR)ACCOUNTINGYEAR, " &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB, " &_
				" MAX(IPLASPROD.PARAMETER.PARAM_6) JOBCODE_DB, " &_
				" 'SL' SOURCE_DB, " &_
				" 'SL' SOURCETYPE_DB, " &_
				" MAX(DTL.SALESSHIPMENTNO) SOURCECODE_DB, " &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB, " &_
				" MAX(DTL.FIELDCODE) TARGETCODE_DB, " &_
				" MAX(DTL.TRANSPORTERCODE) DOCUMENTNO_DB, " &_
				" 0 QUANTITY_DB, " &_
				" 0 OUTPUT_DB, " &_
				" MAX(DTL.TRANSPORTERCOST) AMOUNT_DB," &_
				" 'TC - Sales Shipment No :' TITLE_REFERENCE_01_DB," &_
				" MAX(DTL.SALESSHIPMENTNO) AUTO_REFERENCE_01_DB," &_
				" 'Auto - Transporter Code :' TITLE_REFERENCE_02_DB," &_
				" MAX(DTL.TRANSPORTERCODE) AUTO_REFERENCE_02_DB," &_
				" 'Auto - Chitno :' TITLE_REFERENCE_03_DB," &_
				" MAX(DTL.CHITNO) AUTO_REFERENCE_03_DB" &_	
				" FROM" &_
				" IPLASPROD.TRANSPORTER_VIEW DTL " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES' " &_
				" INNER JOIN IPLASPROD.PARAMETER ON DTL.FCBA = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES' " &_
				" INNER JOIN IPLASPROD.JOB ON DTL.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_6 = IPLASPROD.JOB.FCCODE " &_
				" WHERE " &_
				" TO_CHAR(DTL.POSTINGDATE,'fmMM') = " & Mid(Session("CurP_Sales"),5,2) &_
				" AND TO_CHAR(DTL.POSTINGDATE,'yyyy') =  " & Mid(Session("CurP_Sales"),1,4) & _
				" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
				" GROUP BY DTL.FCBA,DTL.SALESSHIPMENTNO, DTL.FIELDCODE, DTL.CHITNO)"
	
	Set rsTransporterDB = Server.CreateObject("ADODB.Recordset")
	rsTransporterDB.Open qTransporterDB, connstr
	
	if not rsTransporterDB.eof then
		xError = "0"
		'response.write qTransporter
		Do While Not rsTransporterDB.eof
		
			if rsTransporterDB("TARGETCODE_DB") <> "" and (rsTransporterDB("AMOUNT_DB") <> "" or rsTransporterDB("AMOUNT_DB") <> "0") then
				xtTargetCodeDB =rsTransporterDB("TARGETCODE_DB")
				xtJobcodeDB = rsTransporterDB("JOBCODE_DB")
				
				maxLineNo = maxLineNo + 1
				qCost_ActivityDB11="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsTransporterDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & xtJobcodeDB & "','" & rsTransporterDB("SOURCE_DB") & "','" & rsTransporterDB("SOURCETYPE_DB") & "','" & rsTransporterDB("SOURCECODE_DB") & "'," & _
								 " '','" & rsTransporterDB("TARGETTYPE_DB") & "','" & xtTargetCodeDB & "'," & _
								 " '" &  rsTransporterDB("AMOUNT_DB") & "','" & rsTransporterDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsTransporterDB("FCBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsTransporterDB("TITLE_REFERENCE_01_DB") & "','"  & rsTransporterDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsTransporterDB("TITLE_REFERENCE_02_DB") & "','"  & rsTransporterDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '" & rsTransporterDB("TITLE_REFERENCE_03_DB") & "','"  & rsTransporterDB("AUTO_REFERENCE_03_DB") & "'," & _
								 " '','', " & _
								 " 'Line No:', '" & rsTransporterDB("LINENO") & "','Db', " & _
								 " " & rsTransporterDB("ACCOUNTINGPERIOD") & "," & rsTransporterDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				connstr.execute(qCost_ActivityDB11)
				connStr.execute("Commit")
			else 
				response.Write "No Shipment <font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & rsTransporterDB("SOURCECODE_DB") & "</font> Tidak ada Harvesting !!<br>"
				response.Flush
				xError = "1"
			end if			
			rsTransporterDB.MoveNext
		loop
		
		if xError = "0" then
			response.Write "Process Biaya Transport Debet Complete.... !!<br>"
			response.Flush
		else
			response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif> Closing Biaya Transport Gagal !!!</font><br>"
			response.Flush
		end if
	end if
	'-- End Insert Debet
	
	'-- Insert Credit
		qTransporterCR =" SELECT" &_
				" FCBA, " &_
				" ROWNUM LINENO," &_
				" ACCOUNTINGPERIOD," &_
				" ACCOUNTINGYEAR, " &_
				" TDATE_CR, " &_
				" JOBCODE_CR, " &_
				" SOURCE_CR, " &_
				" SOURCETYPE_CR, " &_
				" SOURCECODE_CR, " &_
				" TARGETTYPE_CR, " &_
				" TARGETCODE_CR, " &_
				" DOCUMENTNO_CR, " &_
				" QUANTITY_CR, " &_
				" OUTPUT_CR, " &_
				" AMOUNT_CR," &_
				" TITLE_REFERENCE_01_CR," &_
				" AUTO_REFERENCE_01_CR," &_
				" TITLE_REFERENCE_02_CR," &_
				" AUTO_REFERENCE_02_CR," &_
				" TITLE_REFERENCE_03_CR," &_
				" AUTO_REFERENCE_03_CR" &_
				" FROM" &_
				" (" &_
				" SELECT " &_
				" MAX(DTL.FCBA)FCBA, " &_
				" MAX(DTL.ACCOUNTINGPERIOD)ACCOUNTINGPERIOD, " &_
				" MAX(DTL.ACCOUNTINGYEAR)ACCOUNTINGYEAR, " &_	
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR, " &_
				" MAX(IPLASPROD.PARAMETER.PARAM_7) JOBCODE_CR, " &_
				" 'SL' SOURCE_CR, " &_
				" 'SL' SOURCETYPE_CR, " &_
				" MAX(DTL.SALESSHIPMENTNO) SOURCECODE_CR, " &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_CR, " &_
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_CR, " &_
				" MAX(DTL.TRANSPORTERCODE) DOCUMENTNO_CR, " &_
				" 0 QUANTITY_CR, " &_
				" 0 OUTPUT_CR, " &_
				" MAX(-1*DTL.TRANSPORTERCOST) AMOUNT_CR, " &_
				" 'TC - Sales Shipment No :' TITLE_REFERENCE_01_CR," &_
				" MAX(DTL.SALESSHIPMENTNO) AUTO_REFERENCE_01_CR," &_
				" 'Auto - Transporter Code :' TITLE_REFERENCE_02_CR," &_
				" MAX(DTL.TRANSPORTERCODE) AUTO_REFERENCE_02_CR," &_
				" 'Auto - Chitno :' TITLE_REFERENCE_03_CR," &_
				" MAX(DTL.CHITNO) AUTO_REFERENCE_03_CR" &_
				" FROM" &_
				" IPLASPROD.TRANSPORTER_VIEW DTL " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES' " &_
				" INNER JOIN IPLASPROD.PARAMETER ON DTL.FCBA = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES' " &_
				" INNER JOIN IPLASPROD.JOB ON DTL.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_7 = IPLASPROD.JOB.FCCODE " &_
				" INNER JOIN IPLASPROD.GCMASTER ON DTL.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY" &_ 
				" WHERE " &_
				" TO_CHAR(DTL.POSTINGDATE,'fmMM') = " & Mid(Session("CurP_Sales"),5,2) &_
				" AND TO_CHAR(DTL.POSTINGDATE,'yyyy') =  " & Mid(Session("CurP_Sales"),1,4) & _
				" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
				" GROUP BY DTL.FCBA,DTL.SALESSHIPMENTNO, DTL.FIELDCODE, DTL.CHITNO)"
	
	Set rsTransporterCR = Server.CreateObject("ADODB.Recordset")
	rsTransporterCR.Open qTransporterCR, connstr
	
	if not rsTransporterCR.eof then
		xError = "0"
		'response.write rsTransporterCR
		Do While Not rsTransporterCR.eof
		
			if rsTransporterCR("TARGETCODE_CR") <> "" and (rsTransporterCR("AMOUNT_CR") <> "" or rsTransporterCR("AMOUNT_CR") <> "0") then
				xtTargetCodeCR =rsTransporterCR("TARGETCODE_CR")
				xtJobcodeCR = rsTransporterCR("JOBCODE_CR")

			 maxLineNo = maxLineNo + 1
			 qCost_ActivityCR11="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsTransporterCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & xtJobcodeCR & "','" & rsTransporterCR("SOURCE_CR") & "','" & rsTransporterCR("SOURCETYPE_CR") & "','" & rsTransporterCR("SOURCECODE_CR") & "'," & _
							 " '','" & rsTransporterCR("TARGETTYPE_CR") & "','" & xtTargetCodeCR & "'," & _
							 " '" &  rsTransporterCR("AMOUNT_CR") & "','" & rsTransporterCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsTransporterCR("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsTransporterCR("TITLE_REFERENCE_01_CR") & "','"  & rsTransporterCR("AUTO_REFERENCE_01_CR") & "'," & _
							 " '" & rsTransporterCR("TITLE_REFERENCE_02_CR") & "','"  & rsTransporterCR("AUTO_REFERENCE_02_CR") & "'," & _
							 " '" & rsTransporterCR("TITLE_REFERENCE_03_CR") & "','"  & rsTransporterCR("AUTO_REFERENCE_03_CR") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsTransporterCR("LINENO") & "','Cr', " & _
							 " " & rsTransporterCR("ACCOUNTINGPERIOD") & "," & rsTransporterCR("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
			 connstr.execute(qCost_ActivityCR11)
			 connStr.execute("Commit")
			else
			 xError = "1"
			end if
			
			rsTransporterCR.MoveNext
		loop
		
		if xError = "0" then
			response.Write "Process Biaya Transport Credit Complete.... !!<br>"
			response.Flush
		end if
		
	end if
	'-- End Insert Credit

	'-- End Biaya Transporter
	end if
	
	if Session("Estate_Mill") = "HO" then
	'-- Begin Transporter Invoice
	
	'-- Insert Debet 1
		qTransporterInvDB1 = " SELECT" &_		
				" FCBA, ROWNUM LINENO, " &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_ 
				" TDATE_DB1, JOBCODE_DB1, SOURCE_DB1, SOURCETYPE_DB1, SOURCECODE_DB1, TARGETTYPE_DB1, TARGETCODE_DB1,  " &_
				" DOCUMENTNO_DB1, QUANTITY_DB1, OUTPUT_DB1, AMOUNT_DB1, " &_
				" TITLE_REFERENCE_01_DB1, AUTO_REFERENCE_01_DB1, " &_
				" TITLE_REFERENCE_02_DB1, AUTO_REFERENCE_02_DB1, " &_
				" TITLE_REFERENCE_03_DB1, AUTO_REFERENCE_03_DB1 " &_
				" FROM " &_
				" ( " &_
				" SELECT  " &_
				" MAX(HDR.INVOICE_BY)FCBA, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR," &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB1, " &_
				" MAX(IPLASPROD.PARAMETER.PARAM_7) JOBCODE_DB1, " &_
				" 'SL' SOURCE_DB1, " &_
				" 'SL' SOURCETYPE_DB1, " &_
				" MAX(DTL.INVOICECODE) SOURCECODE_DB1, " &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB1, " &_
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_DB1, " &_
				" MAX(DTL.WBT_CHITNO) DOCUMENTNO_DB1, " &_
				" 0 QUANTITY_DB1, " &_
				" 0 OUTPUT_DB1, " &_
				" MAX(DTL.AMOUNTAFTERLOSS) AMOUNT_DB1, " &_
				" 'TC - Sales Shipment No :' TITLE_REFERENCE_01_DB1, " &_
				" MAX(DTL.INVOICECODE) AUTO_REFERENCE_01_DB1, " &_
				" 'Auto - Transporter Code :' TITLE_REFERENCE_02_DB1, " &_
				" MAX(HDR.TRANSPORTERCONTRACTNO) AUTO_REFERENCE_02_DB1, " &_
				" 'Auto - Chitno :' TITLE_REFERENCE_03_DB1, " &_
				" MAX(DTL.WBT_CHITNO) AUTO_REFERENCE_03_DB1" &_
				" FROM " &_
				" IPLASPROD.TRANSPORTERINVOICEDTL DTL " &_
				" INNER JOIN IPLASPROD.TRANSPORTERINVOICE HDR ON DTL.FCBA = HDR.FCBA AND DTL.INVOICECODE = HDR.INVOICECODE " &_
				" LEFT JOIN IPLASPROD.TRANSPORTERCONTRACT ON HDR.FCBA = IPLASPROD.TRANSPORTERCONTRACT.FCBA AND HDR.TRANSPORTERCONTRACTNO = IPLASPROD.TRANSPORTERCONTRACT.TRANSPORTERCONTRACTNO " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.INVOICE_BY = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES'  " &_
				" INNER JOIN IPLASPROD.PARAMETER ON HDR.INVOICE_BY = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES'  " &_
				" INNER JOIN IPLASPROD.JOB ON HDR.INVOICE_BY = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_7 = IPLASPROD.JOB.FCCODE  " &_
				" INNER JOIN IPLASPROD.GCMASTER ON HDR.INVOICE_BY = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY " &_ 
				" WHERE  " &_
				" TO_CHAR(HDR.CREATIONDATE,'fmMM') = " & Mid(Session("CurP_Sales"),5,2) &_
				" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') =  " & Mid(Session("CurP_Sales"),1,4) & _
				" AND HDR.INVOICE_BY = '" & Session("fcBA") & "'" &_
				" GROUP BY HDR.INVOICE_BY,HDR.FCBA,DTL.INVOICECODE, DTL.WBT_CHITNO)"

	Set rsTransporterInvDB1 = Server.CreateObject("ADODB.Recordset")
	rsTransporterInvDB1.Open qTransporterInvDB1, connstr
	if not rsTransporterInvDB1.eof then
		Do While Not rsTransporterInvDB1.eof
		
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityTRDB1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsTransporterInvDB1("TDATE_DB1") & "','mm/dd/yyyy') " & ",'" & rsTransporterInvDB1("JOBCODE_DB1") & "','" & rsTransporterInvDB1("SOURCE_DB1") & "','" & rsTransporterInvDB1("SOURCETYPE_DB1") & "','" & rsTransporterInvDB1("SOURCECODE_DB1") & "'," & _
							 " '','" & rsTransporterInvDB1("TARGETTYPE_DB1") & "','" & rsTransporterInvDB1("TARGETCODE_DB1") & "'," & _
							 " '" &  rsTransporterInvDB1("AMOUNT_DB1") & "','" & rsTransporterInvDB1("DOCUMENTNO_DB1") & "'," & "0,0,'" & rsTransporterInvDB1("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsTransporterInvDB1("TITLE_REFERENCE_01_DB1") & "','"  & rsTransporterInvDB1("AUTO_REFERENCE_01_DB1") & "'," & _
							 " '" & rsTransporterInvDB1("TITLE_REFERENCE_02_DB1") & "','"  & rsTransporterInvDB1("AUTO_REFERENCE_02_DB1") & "'," & _
							 " '" & rsTransporterInvDB1("TITLE_REFERENCE_03_DB1") & "','"  & rsTransporterInvDB1("AUTO_REFERENCE_03_DB1") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsTransporterInvDB1("LINENO") & "','Db', " & _
							 " " & rsTransporterInvDB1("ACCOUNTINGPERIOD") & "," & rsTransporterInvDB1("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityTRDB1)
			 connStr.execute("Commit")
			 
			rsTransporterInvDB1.MoveNext
		loop
		response.Write "Process Biaya Transport Debet Complete.... !!<br>"
		response.Flush
	end if	
	'-- End Insert Debet 1
	
	'-- Insert Debet 2
	qTransporterInvDB2 =" SELECT" &_				
				" FCBA, ROWNUM LINENO, " &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_ 
				" TDATE_DB2, JOBCODE_DB2, SOURCE_DB2, SOURCETYPE_DB2, SOURCECODE_DB2, TARGETTYPE_DB2, TARGETCODE_DB2, " &_ 
				" DOCUMENTNO_DB2, QUANTITY_DB2, OUTPUT_DB2, AMOUNT_DB2, " &_
				" TITLE_REFERENCE_01_DB2, AUTO_REFERENCE_01_DB2, " &_
				" TITLE_REFERENCE_02_DB2, AUTO_REFERENCE_02_DB2, " &_
				" TITLE_REFERENCE_03_DB2, AUTO_REFERENCE_03_DB2 " &_
				" FROM " &_
				" ( " &_
				" SELECT  " &_
				" MAX(HDR.INVOICE_BY)FCBA, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR," &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB2, " &_
				" MAX(IPLASPROD.PARAMETER.PARAM_3) JOBCODE_DB2, " &_
				" 'SL' SOURCE_DB2, " &_
				" 'SL' SOURCETYPE_DB2, " &_
				" MAX(DTL.INVOICECODE) SOURCECODE_DB2, " &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB2, " &_
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_DB2, " &_
				" MAX(DTL.WBT_CHITNO) DOCUMENTNO_DB2, " &_
				" 0 QUANTITY_DB2, " &_
				" 0 OUTPUT_DB2, " &_
				" MAX(DTL.AMOUNTAFTERLOSS) * MAX(HDR.VAT)/100 AMOUNT_DB2," &_
				" 'TC - Sales Shipment No :' TITLE_REFERENCE_01_DB2, " &_
				" MAX(DTL.INVOICECODE) AUTO_REFERENCE_01_DB2, " &_
				" 'Auto - Transporter Code :' TITLE_REFERENCE_02_DB2, " &_
				" MAX(HDR.TRANSPORTERCONTRACTNO) AUTO_REFERENCE_02_DB2, " &_
				" 'Auto - Chitno :' TITLE_REFERENCE_03_DB2, " &_
				" MAX(DTL.WBT_CHITNO) AUTO_REFERENCE_03_DB2 " &_
				" FROM " &_
				" IPLASPROD.TRANSPORTERINVOICEDTL DTL " &_
				" INNER JOIN IPLASPROD.TRANSPORTERINVOICE HDR ON DTL.FCBA = HDR.FCBA AND DTL.INVOICECODE = HDR.INVOICECODE " &_
				" LEFT JOIN IPLASPROD.TRANSPORTERCONTRACT ON HDR.FCBA = IPLASPROD.TRANSPORTERCONTRACT.FCBA AND HDR.TRANSPORTERCONTRACTNO = IPLASPROD.TRANSPORTERCONTRACT.TRANSPORTERCONTRACTNO " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.INVOICE_BY = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES'  " &_
				" INNER JOIN IPLASPROD.PARAMETER ON HDR.INVOICE_BY = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES'  " &_
				" INNER JOIN IPLASPROD.JOB ON HDR.INVOICE_BY = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_3 = IPLASPROD.JOB.FCCODE  " &_
				" INNER JOIN IPLASPROD.GCMASTER ON HDR.INVOICE_BY = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY " &_ 
				" WHERE  " &_
				" TO_CHAR(HDR.CREATIONDATE,'fmMM') = " & Mid(Session("CurP_Sales"),5,2) &_
				" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') =  " & Mid(Session("CurP_Sales"),1,4) & _
				" AND HDR.INVOICE_BY = '" & Session("fcBA") & "'" &_
				" GROUP BY HDR.INVOICE_BY,HDR.FCBA,DTL.INVOICECODE, DTL.WBT_CHITNO)"

	Set rsTransporterInvDB2 = Server.CreateObject("ADODB.Recordset")
	rsTransporterInvDB2.Open qTransporterInvDB2, connstr
	if not rsTransporterInvDB2.eof then
		Do While Not rsTransporterInvDB2.eof			
			
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityTRDB2="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsTransporterInvDB2("TDATE_DB2") & "','mm/dd/yyyy') " & ",'" & rsTransporterInvDB2("JOBCODE_DB2") & "','" & rsTransporterInvDB2("SOURCE_DB2") & "','" & rsTransporterInvDB2("SOURCETYPE_DB2") & "','" & rsTransporterInvDB2("SOURCECODE_DB2") & "'," & _
							 " '','" & rsTransporterInvDB2("TARGETTYPE_DB2") & "','" & rsTransporterInvDB2("TARGETCODE_DB2") & "'," & _
							 " '" &  rsTransporterInvDB2("AMOUNT_DB2") & "','" & rsTransporterInvDB2("DOCUMENTNO_DB2") & "'," & "0,0,'" & rsTransporterInvDB2("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsTransporterInvDB2("TITLE_REFERENCE_01_DB2") & "','"  & rsTransporterInvDB2("AUTO_REFERENCE_01_DB2") & "'," & _
							 " '" & rsTransporterInvDB2("TITLE_REFERENCE_02_DB2") & "','"  & rsTransporterInvDB2("AUTO_REFERENCE_02_DB2") & "'," & _
							 " '" & rsTransporterInvDB2("TITLE_REFERENCE_03_DB2") & "','"  & rsTransporterInvDB2("AUTO_REFERENCE_03_DB2") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsTransporterInvDB2("LINENO") & "','Db', " & _
							 " " & rsTransporterInvDB2("ACCOUNTINGPERIOD") & "," & rsTransporterInvDB2("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityTRDB2)
			 connStr.execute("Commit")
			 
			rsTransporterInvDB2.MoveNext
		loop
		response.Write "Process Biaya Transport Debet VAT Complete.... !!<br>"
		response.Flush
	end if	
	'-- End Insert Debet 2
	
	'-- Begin Insert Credit 1
	qTransporterInvCR1 =" SELECT" &_			
				" FCBA, ROWNUM LINENO, " &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
				" TDATE_CR1, JOBCODE_CR1, SOURCE_CR1, SOURCETYPE_CR1, SOURCECODE_CR1, TARGETTYPE_CR1, TARGETCODE_CR1, " &_
				" DOCUMENTNO_CR1, QUANTITY_CR1, OUTPUT_CR1, AMOUNT_CR1, " &_
				" TITLE_REFERENCE_01_CR1, AUTO_REFERENCE_01_CR1, " &_
				" TITLE_REFERENCE_02_CR1, AUTO_REFERENCE_02_CR1, " &_
				" TITLE_REFERENCE_03_CR1, AUTO_REFERENCE_03_CR1 " &_
				" FROM " &_
				" ( " &_
				" SELECT  " &_
				" MAX(HDR.INVOICE_BY)FCBA, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR," &_ 
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR1, " &_
				" MAX(IPLASPROD.PARAMETER.PARAM_8) JOBCODE_CR1, " &_
				" 'SL' SOURCE_CR1, " &_
				" 'SL' SOURCETYPE_CR1, " &_
				" MAX(DTL.INVOICECODE) SOURCECODE_CR1, " &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_CR1, " &_
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_CR1, " &_
				" MAX(DTL.WBT_CHITNO) DOCUMENTNO_CR1, " &_
				" 0 QUANTITY_CR1, " &_
				" 0 OUTPUT_CR1, " &_
				" -1*(MAX(DTL.AMOUNTAFTERLOSS) * MAX(HDR.PPH)/100) AMOUNT_CR1," &_
				" 'TC - Sales Shipment No :' TITLE_REFERENCE_01_CR1, " &_
				" MAX(DTL.INVOICECODE) AUTO_REFERENCE_01_CR1, " &_
				" 'Auto - Transporter Code :' TITLE_REFERENCE_02_CR1, " &_
				" MAX(HDR.TRANSPORTERCONTRACTNO) AUTO_REFERENCE_02_CR1, " &_
				" 'Auto - Chitno :' TITLE_REFERENCE_03_CR1, " &_
				" MAX(DTL.WBT_CHITNO) AUTO_REFERENCE_03_CR1 " &_
				" FROM " &_
				" IPLASPROD.TRANSPORTERINVOICEDTL DTL " &_
				" INNER JOIN IPLASPROD.TRANSPORTERINVOICE HDR ON DTL.FCBA = HDR.FCBA AND DTL.INVOICECODE = HDR.INVOICECODE " &_
				" LEFT JOIN IPLASPROD.TRANSPORTERCONTRACT ON HDR.FCBA = IPLASPROD.TRANSPORTERCONTRACT.FCBA AND HDR.TRANSPORTERCONTRACTNO = IPLASPROD.TRANSPORTERCONTRACT.TRANSPORTERCONTRACTNO " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.INVOICE_BY = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES'  " &_
				" INNER JOIN IPLASPROD.PARAMETER ON HDR.INVOICE_BY = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES'  " &_
				" INNER JOIN IPLASPROD.JOB ON HDR.INVOICE_BY = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_8 = IPLASPROD.JOB.FCCODE  " &_
				" INNER JOIN IPLASPROD.GCMASTER ON HDR.INVOICE_BY = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY " &_ 
				" WHERE  " &_
				" TO_CHAR(HDR.CREATIONDATE,'fmMM') = " & Mid(Session("CurP_Sales"),5,2) &_
				" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') =  " & Mid(Session("CurP_Sales"),1,4) & _
				" AND HDR.INVOICE_BY = '" & Session("fcBA") & "'" &_
				" GROUP BY HDR.INVOICE_BY,HDR.FCBA,DTL.INVOICECODE, DTL.WBT_CHITNO)"

	Set rsTransporterInvCR1 = Server.CreateObject("ADODB.Recordset")
	rsTransporterInvCR1.Open qTransporterInvCR1, connstr
	if not rsTransporterInvCR1.eof then
		Do While Not rsTransporterInvCR1.eof			
			
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityTRCR1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsTransporterInvCR1("TDATE_CR1") & "','mm/dd/yyyy') " & ",'" & rsTransporterInvCR1("JOBCODE_CR1") & "','" & rsTransporterInvCR1("SOURCE_CR1") & "','" & rsTransporterInvCR1("SOURCETYPE_CR1") & "','" & rsTransporterInvCR1("SOURCECODE_CR1") & "'," & _
							 " '','" & rsTransporterInvCR1("TARGETTYPE_CR1") & "','" & rsTransporterInvCR1("TARGETCODE_CR1") & "'," & _
							 " '" &  rsTransporterInvCR1("AMOUNT_CR1") & "','" & rsTransporterInvCR1("DOCUMENTNO_CR1") & "'," & "0,0,'" & rsTransporterInvCR1("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsTransporterInvCR1("TITLE_REFERENCE_01_CR1") & "','"  & rsTransporterInvCR1("AUTO_REFERENCE_01_CR1") & "'," & _
							 " '" & rsTransporterInvCR1("TITLE_REFERENCE_02_CR1") & "','"  & rsTransporterInvCR1("AUTO_REFERENCE_02_CR1") & "'," & _
							 " '" & rsTransporterInvCR1("TITLE_REFERENCE_03_CR1") & "','"  & rsTransporterInvCR1("AUTO_REFERENCE_03_CR1") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsTransporterInvCR1("LINENO") & "','Cr', " & _
							 " " & rsTransporterInvCR1("ACCOUNTINGPERIOD") & "," & rsTransporterInvCR1("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityTRCR1)
			 connStr.execute("Commit")
			 
			rsTransporterInvCR1.MoveNext
		loop
		response.Write "Process Biaya Transport Credit Complete.... !!<br>"
		response.Flush
	end if	
	'-- End Insert Credit 1

	'-- Insert Credit 2
	qTransporterInvCR2 =" SELECT" &_				
				" FCBA, ROWNUM LINENO, " &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_ 			 
				" TDATE_CR2, JOBCODE_CR2, SOURCE_CR2, SOURCETYPE_CR2, SOURCECODE_CR2, TARGETTYPE_CR2, TARGETCODE_CR2, " &_
				" DOCUMENTNO_CR2, QUANTITY_CR2, OUTPUT_CR2, AMOUNT_CR2, " &_
				" TITLE_REFERENCE_01_CR2, AUTO_REFERENCE_01_CR2, " &_
				" TITLE_REFERENCE_02_CR2, AUTO_REFERENCE_02_CR2, " &_
				" TITLE_REFERENCE_03_CR2, AUTO_REFERENCE_03_CR2 " &_
				" FROM " &_
				" ( " &_
				" SELECT  " &_
				" MAX(HDR.INVOICE_BY)FCBA, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR," &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR2, " &_
				" MAX(IPLASPROD.TRANSPORTER.CONTROLJOB) JOBCODE_CR2,  " &_
				" 'SL' SOURCE_CR2, " &_
				" 'SL' SOURCETYPE_CR2, " &_
				" MAX(DTL.INVOICECODE) SOURCECODE_CR2, " &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_CR2, " &_
				" MAX(IPLASPROD.TRANSPORTERCONTRACT.TRANSPORTER) TARGETCODE_CR2, " &_
				" MAX(DTL.WBT_CHITNO) DOCUMENTNO_CR2, " &_
				" 0 QUANTITY_CR2, " &_
				" 0 OUTPUT_CR2, " &_
				" -1*(MAX(DTL.AMOUNTAFTERLOSS)+(MAX(DTL.AMOUNTAFTERLOSS) * MAX(HDR.VAT)/100)-(MAX(DTL.AMOUNTAFTERLOSS) * MAX(HDR.PPH)/100)) AMOUNT_CR2, " &_
				" 'TC - Sales Shipment No :' TITLE_REFERENCE_01_CR2, " &_
				" MAX(DTL.INVOICECODE) AUTO_REFERENCE_01_CR2, " &_
				" 'Auto - Transporter Code :' TITLE_REFERENCE_02_CR2, " &_
				" MAX(HDR.TRANSPORTERCONTRACTNO) AUTO_REFERENCE_02_CR2, " &_
				" 'Auto - Chitno :' TITLE_REFERENCE_03_CR2, " &_
				" MAX(DTL.WBT_CHITNO) AUTO_REFERENCE_03_CR2 " &_
				" FROM " &_
				" IPLASPROD.TRANSPORTERINVOICEDTL DTL " &_
				" INNER JOIN IPLASPROD.TRANSPORTERINVOICE HDR ON DTL.FCBA = HDR.FCBA AND DTL.INVOICECODE = HDR.INVOICECODE " &_
				" LEFT JOIN IPLASPROD.TRANSPORTERCONTRACT ON HDR.FCBA = IPLASPROD.TRANSPORTERCONTRACT.FCBA AND HDR.TRANSPORTERCONTRACTNO = IPLASPROD.TRANSPORTERCONTRACT.TRANSPORTERCONTRACTNO " &_
				" INNER JOIN IPLASPROD.TRANSPORTER ON IPLASPROD.TRANSPORTERCONTRACT.FCBA = IPLASPROD.TRANSPORTER.FCBA AND IPLASPROD.TRANSPORTERCONTRACT.TRANSPORTER = IPLASPROD.TRANSPORTER.FCCODE "&_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.INVOICE_BY = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES'  " &_
				" INNER JOIN IPLASPROD.PARAMETER ON HDR.INVOICE_BY = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES'  " &_
				" INNER JOIN IPLASPROD.JOB ON HDR.INVOICE_BY = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_8 = IPLASPROD.JOB.FCCODE  " &_
				" INNER JOIN IPLASPROD.GCMASTER ON HDR.INVOICE_BY = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY " &_ 
				" WHERE  " &_
				" TO_CHAR(HDR.CREATIONDATE,'fmMM') = " & Mid(Session("CurP_Sales"),5,2) &_
				" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') =  " & Mid(Session("CurP_Sales"),1,4) & _
				" AND HDR.INVOICE_BY = '" & Session("fcBA") & "'" &_
				" GROUP BY HDR.INVOICE_BY,HDR.FCBA,DTL.INVOICECODE, DTL.WBT_CHITNO)"

	Set rsTransporterInvCR2 = Server.CreateObject("ADODB.Recordset")
	rsTransporterInvCR2.Open qTransporterInvCR2, connstr
	if not rsTransporterInvCR2.eof then
		Do While Not rsTransporterInvCR2.eof
			 
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityTRCR2="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsTransporterInvCR2("TDATE_CR2") & "','mm/dd/yyyy') " & ",'" & rsTransporterInvCR2("JOBCODE_CR2") & "','" & rsTransporterInvCR2("SOURCE_CR2") & "','" & rsTransporterInvCR2("SOURCETYPE_CR2") & "','" & rsTransporterInvCR2("SOURCECODE_CR2") & "'," & _
							 " '','" & rsTransporterInvCR2("TARGETTYPE_CR2") & "','" & rsTransporterInvCR2("TARGETCODE_CR2") & "'," & _
							 " '" &  rsTransporterInvCR2("AMOUNT_CR2") & "','" & rsTransporterInvCR2("DOCUMENTNO_CR2") & "'," & "0,0,'" & rsTransporterInvCR2("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsTransporterInvCR2("TITLE_REFERENCE_01_CR2") & "','"  & rsTransporterInvCR2("AUTO_REFERENCE_01_CR2") & "'," & _
							 " '" & rsTransporterInvCR2("TITLE_REFERENCE_02_CR2") & "','"  & rsTransporterInvCR2("AUTO_REFERENCE_02_CR2") & "'," & _
							 " '" & rsTransporterInvCR2("TITLE_REFERENCE_03_CR2") & "','"  & rsTransporterInvCR2("AUTO_REFERENCE_03_CR2") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsTransporterInvCR2("LINENO") & "','Cr', " & _
							 " " & rsTransporterInvCR2("ACCOUNTINGPERIOD") & "," & rsTransporterInvCR2("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
			 connstr.execute(qCost_ActivityTRCR2)
			 connStr.execute("Commit")
			 
			rsTransporterInvCR2.MoveNext
		loop
		response.Write "Process Biaya Transport Credit VAT Complete.... !!<br>"
		response.Flush
	end if	
	'-- End Insert Credit 2
	
	'-- End Transporter Invoice
	end if
	
	if Session("Estate_Mill") = "E" then
	'-- Begin Transporter Deduction
	'-- Insert Debet
		qTransporterDeductionDB =" SELECT" &_
				" FCBA, ROWNUM LINENO," & _
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR," & _ 
				" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB, " & _
				" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB," & _
				" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB, " & _
				" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB," & _
				" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB" & _
				" FROM" & _
				" (" & _
				" SELECT" & _ 
				" MAX(DTL.FCBA)FCBA," & _ 
				" MAX(TO_CHAR(DTL.POSTINGDATE,'fmMM')) ACCOUNTINGPERIOD," & _
				" MAX(TO_CHAR(DTL.POSTINGDATE,'yyyy')) ACCOUNTINGYEAR, " & _
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB, " & _
				" MAX(IPLASPROD.PARAMETER.PARAM_7) JOBCODE_DB, " & _
				" 'SL' SOURCE_DB, " & _
				" 'SL' SOURCETYPE_DB, " & _
				" MAX(DTL.SALESSHIPMENTNO) SOURCECODE_DB, " & _
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB," & _ 
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_DB, " & _
				" MAX(DTL.CHITNO) DOCUMENTNO_DB, " & _
				" 0 QUANTITY_DB, " & _
				" 0 OUTPUT_DB, " & _
				" MAX(DTL.PENALTYDEDUCTION) AMOUNT_DB," & _
				" 'TC - Sales Shipment No :' TITLE_REFERENCE_01_DB," & _
				" MAX(DTL.SALESSHIPMENTNO) AUTO_REFERENCE_01_DB," & _
				" 'Auto - Chitno :' TITLE_REFERENCE_02_DB," & _
				" MAX(DTL.CHITNO) AUTO_REFERENCE_02_DB," & _
				" 'Auto - Field Code :' TITLE_REFERENCE_03_DB," & _
				" MAX(DTL.FIELDCODE) AUTO_REFERENCE_03_DB" & _
				" FROM" & _
				" IPLASPROD.TRANSPORTER_PENALTY_VIEW DTL" & _
				" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES' " & _
				" INNER JOIN IPLASPROD.PARAMETER ON DTL.FCBA = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES' " & _
				" INNER JOIN IPLASPROD.JOB ON DTL.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_7 = IPLASPROD.JOB.FCCODE " & _
				" INNER JOIN IPLASPROD.GCMASTER ON DTL.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY" & _ 
				" WHERE " & _
				" TO_CHAR(DTL.POSTINGDATE,'fmMM') = " & Mid(Session("CurP_Sales"),5,2) &_
				" AND TO_CHAR(DTL.POSTINGDATE,'yyyy') =  " & Mid(Session("CurP_Sales"),1,4) & _
				" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
				" GROUP BY DTL.FCBA,DTL.SALESSHIPMENTNO, DTL.CHITNO, DTL.FIELDCODE)"

	Set rsTransporterInvoiceDedDB = Server.CreateObject("ADODB.Recordset")
	rsTransporterInvoiceDedDB.Open qTransporterDeductionDB, connstr
	if not rsTransporterInvoiceDedDB.eof then
		Do While Not rsTransporterInvoiceDedDB.eof
		
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityDDB1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsTransporterInvoiceDedDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsTransporterInvoiceDedDB("JOBCODE_DB") & "','" & rsTransporterInvoiceDedDB("SOURCE_DB") & "','" & rsTransporterInvoiceDedDB("SOURCETYPE_DB") & "','" & rsTransporterInvoiceDedDB("SOURCECODE_DB") & "'," & _
							 " '','" & rsTransporterInvoiceDedDB("TARGETTYPE_DB") & "','" & rsTransporterInvoiceDedDB("TARGETCODE_DB") & "'," & _
							 " '" &  rsTransporterInvoiceDedDB("AMOUNT_DB") & "','" & rsTransporterInvoiceDedDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsTransporterInvoiceDedDB("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsTransporterInvoiceDedDB("TITLE_REFERENCE_01_DB") & "','"  & rsTransporterInvoiceDedDB("AUTO_REFERENCE_01_DB") & "'," & _
							 " '" & rsTransporterInvoiceDedDB("TITLE_REFERENCE_02_DB") & "','"  & rsTransporterInvoiceDedDB("AUTO_REFERENCE_02_DB") & "'," & _
							 " '" & rsTransporterInvoiceDedDB("TITLE_REFERENCE_03_DB") & "','"  & rsTransporterInvoiceDedDB("AUTO_REFERENCE_03_DB") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsTransporterInvoiceDedDB("LINENO") & "','Db', " & _
							 " " & rsTransporterInvoiceDedDB("ACCOUNTINGPERIOD") & "," & rsTransporterInvoiceDedDB("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDDB1)
			 connStr.execute("Commit")

			rsTransporterInvoiceDedDB.MoveNext
		loop
		response.Write "Process Penalty Debet Complete.... !!<br>"
		response.Flush
	end if
	'-- End Insert Debet
	
	'-- Insert Credit
		qTransporterDeductionCR =" SELECT" &_
				" FCBA, ROWNUM LINENO," & _
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR," & _ 				
				" TDATE_CR, JOBCODE_CR, SOURCE_CR, SOURCETYPE_CR, SOURCECODE_CR, TARGETTYPE_CR, TARGETCODE_CR, " & _
				" DOCUMENTNO_CR, QUANTITY_CR, OUTPUT_CR, AMOUNT_CR," & _
				" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR, " & _
				" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR," & _
				" TITLE_REFERENCE_03_CR, AUTO_REFERENCE_03_CR" & _
				" FROM" & _
				" (" & _
				" SELECT" & _ 
				" MAX(DTL.FCBA)FCBA," & _ 
				" MAX(TO_CHAR(DTL.POSTINGDATE,'fmMM')) ACCOUNTINGPERIOD," & _
				" MAX(TO_CHAR(DTL.POSTINGDATE,'yyyy')) ACCOUNTINGYEAR, " & _				
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR, " & _
				" MAX(IPLASPROD.PARAMETER.PARAM_6) JOBCODE_CR, " & _
				" 'SL' SOURCE_CR, " & _
				" 'SL' SOURCETYPE_CR, " & _
				" MAX(DTL.SALESSHIPMENTNO) SOURCECODE_CR, " & _
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_CR," & _ 
				" MAX(DTL.FIELDCODE) TARGETCODE_CR, " & _
				" MAX(DTL.CHITNO) DOCUMENTNO_CR, " & _
				" 0 QUANTITY_CR, " & _
				" 0 OUTPUT_CR, " & _
				" MAX(-1*DTL.PENALTYDEDUCTION) AMOUNT_CR," & _
				" 'TC - Sales Shipment No :' TITLE_REFERENCE_01_CR," & _
				" MAX(DTL.SALESSHIPMENTNO) AUTO_REFERENCE_01_CR," & _
				" 'Auto - Chitno :' TITLE_REFERENCE_02_CR," & _
				" MAX(DTL.CHITNO) AUTO_REFERENCE_02_CR," & _
				" 'Auto - Field Code :' TITLE_REFERENCE_03_CR," & _
				" MAX(DTL.FIELDCODE) AUTO_REFERENCE_03_CR" & _
				" FROM" & _
				" IPLASPROD.TRANSPORTER_PENALTY_VIEW DTL" & _
				" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='SALES' " & _
				" INNER JOIN IPLASPROD.PARAMETER ON DTL.FCBA = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'JOBCODE_SALES' " & _
				" INNER JOIN IPLASPROD.JOB ON DTL.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_6 = IPLASPROD.JOB.FCCODE " & _
				" WHERE " & _
				" TO_CHAR(DTL.POSTINGDATE,'fmMM') = " & Mid(Session("CurP_Sales"),5,2) &_
				" AND TO_CHAR(DTL.POSTINGDATE,'yyyy') =  " & Mid(Session("CurP_Sales"),1,4) & _
				" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
				" GROUP BY DTL.FCBA,DTL.SALESSHIPMENTNO, DTL.CHITNO, DTL.FIELDCODE)"

	Set rsTransporterInvoiceDedCR = Server.CreateObject("ADODB.Recordset")
	rsTransporterInvoiceDedCR.Open qTransporterDeductionCR, connstr
	if not rsTransporterInvoiceDedCR.eof then
		Do While Not rsTransporterInvoiceDedCR.eof			 			 
			 
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityDCR1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsTransporterInvoiceDedCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsTransporterInvoiceDedCR("JOBCODE_CR") & "','" & rsTransporterInvoiceDedCR("SOURCE_CR") & "','" & rsTransporterInvoiceDedCR("SOURCETYPE_CR") & "','" & rsTransporterInvoiceDedCR("SOURCECODE_CR") & "'," & _
							 " '','" & rsTransporterInvoiceDedCR("TARGETTYPE_CR") & "','" & rsTransporterInvoiceDedCR("TARGETCODE_CR") & "'," & _
							 " '" &  rsTransporterInvoiceDedCR("AMOUNT_CR") & "','" & rsTransporterInvoiceDedCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsTransporterInvoiceDedCR("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsTransporterInvoiceDedCR("TITLE_REFERENCE_01_CR") & "','"  & rsTransporterInvoiceDedCR("AUTO_REFERENCE_01_CR") & "'," & _
							 " '" & rsTransporterInvoiceDedCR("TITLE_REFERENCE_02_CR") & "','"  & rsTransporterInvoiceDedCR("AUTO_REFERENCE_02_CR") & "'," & _
							 " '" & rsTransporterInvoiceDedCR("TITLE_REFERENCE_03_CR") & "','"  & rsTransporterInvoiceDedCR("AUTO_REFERENCE_03_CR") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsTransporterInvoiceDedCR("LINENO") & "','Cr', " & _
							 " " & rsTransporterInvoiceDedCR("ACCOUNTINGPERIOD") & "," & rsTransporterInvoiceDedCR("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
			 connstr.execute(qCost_ActivityDCR1)
			 connStr.execute("Commit")
			 
			rsTransporterInvoiceDedCR.MoveNext
		loop
		response.Write "Process Penalty Credit Complete.... !!<br>"
		response.Flush
	end if
	'-- End Insert Credit
	
	'-- eND Transporter Deduction
	end if
	
	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))
	
	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND UPPER(SYSTEM) = 'SALES'" 

		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME) VALUES (" & _
					" 'Sales', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy HH:MI:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy HH:MI:ssam')," & _
					" " & tempMaxLine & ", sysdate, to_char(sysdate,'HH24:MI')" & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Sales'"
	
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.Open strLenghtime, connstr
	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Sales'"		
	connStr.execute(upClosingLog)
	'--End Added--
	
	' Add By Akbar 18 Desemeber 2015	
			Dim QAccountPeriod,rsAccountPeriod,temp_NextCloseAccountingPeriod, temp_NextCloseAccountingYear, temp_NextCurrentAccountingPeriod, temp_NextCurrentAccountingYear,temp_EndingAccountingPeriod
			Dim ww
				
			QAccountPeriod	=	" select periodcontrol_api.GetNextCloseAccountingPeriod(system,fcba) NextCloseAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCloseAccountingYear(system,fcba) NextCloseAccountingYear ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingPeriod(system,fcba) NextCurrentAccountingPeriod ,  " & _ 
								" periodcontrol_api.GetNextCurrentAccountingYear(system,fcba) NextCurrentAccountingYear,  " & _ 
								" periodcontrol_api.GetEndingAccountingPeriodTY(system,'" & Session("fcBA") & "')  EndingAccountingPeriod " & _ 
								" from periodcontrol p  " & _ 
								" where UPPER(system) ='SALES' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Sales " & "<br>"
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
				  " where SYSTEM='Sales' and fcba='" & Session("fcBA") & "'"
				ww = " Satu " 
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where SYSTEM='Sales' and fcba='" & Session("fcBA") & "'"
					  ww = " Dua " 
				ELSE
				    IF temp_NextCloseAccountingPeriod = 12 THEN
						 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
						  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear= " & temp_NextCurrentAccountingYear & " " & _  
						  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						  " where  SYSTEM='Sales' AND fcba='" & Session("fcBA") & "'"
					ELSE
					    q=" Update PeriodControl Set " & _ 
						  " CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear= " & temp_NextCurrentAccountingYear & " " & _  
						  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						  " where  postingposition <= 16  AND fcba='" & Session("fcBA") & "'"
					 
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
				  " where postingposition < 17 and fcba='" & Session("fcBA") & "'"
				ww = " eMPAT " 
			END IF
			
           'response.write ww 		
		   'response.write q
			
			connStr.execute(q)

			' Query ke Period Control untuk mengambli Session(v_Sales) & Session(cls_Sales)	
		
		
			txtLastDay  =cmdFindPeriod("Sales")
			txtNextMonth=cmdPeriodClose("Sales")

			Session("v_Sales")= txtLastDay
			Session("cls_Sales")= txtNextMonth
End Sub

Sub PostEndTime
	TimeAkhir=Time()
	Lama = TimeAkhir-TimeAwal
	Response.Write " Finish ... " & time() & "<br/><br/>"
	Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."
End Sub
Dim TimeAwal,TimeAkhir,Hitung,Lama
TimeAwal=Time()

'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Dim fine_result_01, fine_result_02, fine_result_03

Startingtime = FormatDateTime(CDate(Now))
'--End Added--

%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Sales Last Period Close <%=FormatDateTime(Session("cls_Sales"),1)%> <br> 
Sales Current Period is <%=FormatDateTime(Session("v_Sales"),1)%> <br><br>
Please Wait...<br>
<%
 Response.Write "Start ..." & Time() & "<br>" 
 
 ' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Sales','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Sales " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing

if CheckWaterFlow = "TRUE" then	
		Call NormalisasiData
		Call postingData
		Call PostEndTime
%>
		<br/><br/>
		<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>
			<img src="images/ok.gif"/>Sales Closing is Completed!<br/>
		</font>
<%
else
		Call PostEndTime
%>
		<br/><br/>
		<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>
			<img src="images/error.gif"/>Sales Closing is Not Completed!!!<br/>
		</font>
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