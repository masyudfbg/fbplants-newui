<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="adovbs.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Plasma Reconciliation</title>
</head>
<body>
<%

	Session("v_PlasmaReconsil") = cmdFindPeriod("Plasma_Reconsil")
	Session("cls_PlasmaReconsil") = cmdPeriodClose("Plasma_Reconsil")
	Session("CurP_PlasmaReconsil") = cmdFindAccountPeriod("Plasma_Reconsil")
	Session("ClsP_PlasmaReconsil") = cmdAccountPeriodClose("Plasma_Reconsil")

Sub NormalisasiData
	arr_normalisasi_eof = true
	qPeriodControl = " select lineno from cost_activity " & _
				     " where ACCOUNTINGPERIOD = " & Mid(Session("CurP_PlasmaReconsil"),5,2) & "" & _
				     "       and ACCOUNTINGYEAR = " & Mid(Session("CurP_PlasmaReconsil"),1,4) & "" & _
				     "       and fcBA = '" & Session("fcBA") & "'" & _
				     "       and Sourcetype = 'PL'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
	  arr_normalisasi = rsPeriodControl.getrows()
	  arr_normalisasi_eof = false
	end if
	
	if not arr_normalisasi_eof then
		for xyz=0 to ubound(arr_normalisasi,2)
			qNormalisasiDetail = " DELETE FROM cost_activity_hist_tab WHERE lineheaderno = '"&arr_normalisasi(0,xyz)&"' AND fcba =  '" & Session("fcBA") & "'"
			' connstr.execute(qNormalisasiDetail)
			
			qNormalisasiDetail = " DELETE FROM cost_activity WHERE lineno = '"&arr_normalisasi(0,xyz)&"' AND fcba =  '" & Session("fcBA") & "'"
			connstr.execute(qNormalisasiDetail)
		next
	end if
End Sub

Sub postingData
	' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX 
	' Begin
	Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl
		qPeriodControl="select add_months(LASTPERIODPOSTED,1) LastPeriodPosted, LASTPERIODPOSTED LastEndPeriodPosted from PERIODCONTROL " & _
	  " where SYSTEM='Plasma_Reconsil' and fcba='" & Session("fcBA") & "'"
	
	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr
	
	if not rsPeriodControl.eof then
	   IF  Cint(MID(Session("CurP_PlasmaReconsil"),5,2)) <= 12 Then
		   temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")     ' add Akbar 28 Des 2015 , for Account Period <= 12
	   ELSE
	       temp_datePeriodControl = rsPeriodControl("LASTENDPERIODPOSTED")  ' for Account Period > 12
       END IF
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Plasma Reconsil " & "<br>"
	end if	
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	
	
	' qControlJob="select ControlJob from Module where NAME='Purchase'"
	' Now we use new concept for picking Control Job, by using Process Control Tables
	' as per last concept are NOT able to pick data per Business Unit.
	
	tempProcessCode = "PLASMA_RECONCILIATION"
	tempProcessNameCode = "p_Plasma_Reconciliation.asp"
	tempRefCodePVDetail = "PLASMARECONCILIATION"

	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' Take the Original Table of PLASMARECONCILIATION.
	qReference = "select REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS " & _
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
	
	if Session("Estate_Mill") = "E" then
	
		'-- Begin Plasma Reconsil Plus (+)
		'-- Insert Debet
		qPlasmaDB =" SELECT" &_				
					" ROWNUM LINENO,FCBA, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
					" TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE," &_  
					" TARGETTYPE, TARGETCODE, DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT, " &_
					" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB, " &_
					" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB  " &_
					" FROM " &_
					" ( " &_		 
					" SELECT " &_
					" MAX(ROWNUM) LINENO, " &_
					" MAX(DTL.FCBA)FCBA, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTACCOUNTINGPERIOD)ACCOUNTINGPERIOD, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTACCOUNTINGYEAR)ACCOUNTINGYEAR, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE, " &_
					" MAX(DTL.JOBCODE_PIUTANG_PLASMA) JOBCODE, " &_
					" 'PL' SOURCE, " &_
					" 'PL' SOURCETYPE, " &_
					" MAX(DTL.NAMA_KOPERASI) SOURCECODE, " &_
					" 'CB' TARGETTYPE, " &_
					" 'BS' TARGETCODE, " &_
					" MAX(DTL.NAMA_KOPERASI) DOCUMENTNO, " &_
					" 0 QUANTITY, " &_
					" 0 OUTPUT, " &_
					" SUM(DTL.AMOUNT_PLASMA_RECONCILIATION) AMOUNT, " &_
					" 'TC - Nama Koperasi No :' TITLE_REFERENCE_01_DB,  " &_
					" MAX(DTL.NAMA_KOPERASI) AUTO_REFERENCE_01_DB,  " &_
					" '' TITLE_REFERENCE_02_DB,  " &_
					" '' AUTO_REFERENCE_02_DB  " &_
					" FROM  " &_
					" IPLASPROD.PLASMA_RECONCILIATION_VIEW DTL " &_
					" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='PLASMA_RECONSIL' " &_
					" WHERE " &_
					" to_number(to_char(DTL.CLOSING_DATE,'mm')) = " & Mid(Session("CurP_PlasmaReconsil"),5,2) &_
					" AND to_number(to_char(DTL.CLOSING_DATE,'yyyy')) = " & Mid(Session("CurP_PlasmaReconsil"),1,4) & _
					" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
					" AND DTL.DEBETCREDIT = 'Db' " &_
					" GROUP BY DTL.FCBA,DTL.NAMA_KOPERASI,DTL.JOBCODE_PIUTANG_PLASMA, DTL.JOBCODE_HUTANG_PLASMA " &_
					" ) ORDER BY LINENO "
		'response.write qPlasmaDB
		Set rsPlasmaDB = Server.CreateObject("ADODB.Recordset")
		rsPlasmaDB.Open qPlasmaDB, connstr
				
		if not rsPlasmaDB.eof then
			Do While Not rsPlasmaDB.eof				
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
								 " values (" & maxLineNo & ",to_date('" & rsPlasmaDB("TDATE") & "','mm/dd/yyyy') " & ",'" & rsPlasmaDB("JOBCODE") & "','" & rsPlasmaDB("SOURCE") & "','" & rsPlasmaDB("SOURCETYPE") & "','" & rsPlasmaDB("SOURCECODE") & "'," & _
								 "'','" & rsPlasmaDB("TARGETTYPE") & "','" & rsPlasmaDB("TARGETCODE") & "'," & _
								 "'" &  rsPlasmaDB("AMOUNT") & "','" & rsPlasmaDB("DOCUMENTNO") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsPlasmaDB("TITLE_REFERENCE_01_DB") & "' ,'"  & rsPlasmaDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsPlasmaDB("TITLE_REFERENCE_02_DB") & "' ,'"  & rsPlasmaDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '','', " & _
								 " '','', " & _
								 " 'Line No:', '" & rsPlasmaDB("LINENO") & "','Db', " & _
								 " " & rsPlasmaDB("ACCOUNTINGPERIOD") & "," & rsPlasmaDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDB)
				 connStr.execute("Commit")
				'-- End Insert Debet
				
			rsPlasmaDB.MoveNext
		  loop
		  response.Write "Process Plasma Reconciliation Plus Debet Complete.... !!<br>"
		  response.Flush
		end if
		'-- End Insert Debet
		
		'-- Insert Credit
		qPlasmaCR =" SELECT" &_				
					" ROWNUM LINENO,FCBA, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
					" TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE," &_  
					" TARGETTYPE, TARGETCODE, DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT, " &_
					" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB, " &_
					" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB  " &_
					" FROM " &_
					" ( " &_		 
					" SELECT " &_
					" MAX(ROWNUM) LINENO, " &_
					" MAX(DTL.FCBA)FCBA, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTACCOUNTINGPERIOD)ACCOUNTINGPERIOD, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTACCOUNTINGYEAR)ACCOUNTINGYEAR, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE, " &_
					" MAX(DTL.JOBCODE_HUTANG_PLASMA) JOBCODE, " &_
					" 'PL' SOURCE, " &_
					" 'PL' SOURCETYPE, " &_
					" MAX(DTL.NAMA_KOPERASI) SOURCECODE, " &_
					" 'CB' TARGETTYPE, " &_
					" 'BS' TARGETCODE, " &_
					" MAX(DTL.NAMA_KOPERASI) DOCUMENTNO, " &_
					" 0 QUANTITY, " &_
					" 0 OUTPUT, " &_
					" SUM(DTL.AMOUNT_PLASMA_RECONCILIATION) * -1 AMOUNT , " &_
					" 'TC - Nama Koperasi No :' TITLE_REFERENCE_01_DB,  " &_
					" MAX(DTL.NAMA_KOPERASI) AUTO_REFERENCE_01_DB,  " &_
					" '' TITLE_REFERENCE_02_DB,  " &_
					" '' AUTO_REFERENCE_02_DB  " &_
					" FROM  " &_
					" IPLASPROD.PLASMA_RECONCILIATION_VIEW DTL " &_
					" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='PLASMA_RECONSIL' " &_
					" WHERE " &_
					" to_number(to_char(DTL.CLOSING_DATE,'mm')) = " & Mid(Session("CurP_PlasmaReconsil"),5,2) &_
					" AND to_number(to_char(DTL.CLOSING_DATE,'yyyy')) = " & Mid(Session("CurP_PlasmaReconsil"),1,4) & _
					" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
					" AND DTL.DEBETCREDIT = 'Db' " &_
					" GROUP BY DTL.FCBA,DTL.NAMA_KOPERASI,DTL.JOBCODE_PIUTANG_PLASMA, DTL.JOBCODE_HUTANG_PLASMA " &_
					" ) ORDER BY LINENO "
		'response.write qPlasmaCR
		Set rsPlasmaCR = Server.CreateObject("ADODB.Recordset")
		rsPlasmaCR.Open qPlasmaCR, connstr
				
		if not rsPlasmaCR.eof then
			Do While Not rsPlasmaCR.eof				
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
								 " values (" & maxLineNo & ",to_date('" & rsPlasmaCR("TDATE") & "','mm/dd/yyyy') " & ",'" & rsPlasmaCR("JOBCODE") & "','" & rsPlasmaCR("SOURCE") & "','" & rsPlasmaCR("SOURCETYPE") & "','" & rsPlasmaCR("SOURCECODE") & "'," & _
								 "'','" & rsPlasmaCR("TARGETTYPE") & "','" & rsPlasmaCR("TARGETCODE") & "'," & _
								 "'" &  rsPlasmaCR("AMOUNT") & "','" & rsPlasmaCR("DOCUMENTNO") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsPlasmaCR("TITLE_REFERENCE_01_DB") & "' ,'"  & rsPlasmaCR("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsPlasmaCR("TITLE_REFERENCE_02_DB") & "' ,'"  & rsPlasmaCR("AUTO_REFERENCE_02_DB") & "'," & _
								 " '','', " & _
								 " '','', " & _
								 " 'Line No:', '" & rsPlasmaCR("LINENO") & "','Cr', " & _
								 " " & rsPlasmaCR("ACCOUNTINGPERIOD") & "," & rsPlasmaCR("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDB)
				 connStr.execute("Commit")
				'-- End Insert Debet
				
			rsPlasmaCR.MoveNext
		  loop
		  response.Write "Process Plasma Reconciliation Plus Credit Complete.... !!<br>"
		  response.Flush
		end if
		'-- End Insert Credit
		
		'-- Begin Plasma Reconsil Plus (+)
		
		'-- Begin Plasma Reconsil Minus (-)
		'-- Insert Debet
		qPlasmaMinDB =" SELECT" &_				
					" ROWNUM LINENO,FCBA, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
					" TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE," &_  
					" TARGETTYPE, TARGETCODE, DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT, " &_
					" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB, " &_
					" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB  " &_
					" FROM " &_
					" ( " &_		 
					" SELECT " &_
					" MAX(ROWNUM) LINENO, " &_
					" MAX(DTL.FCBA)FCBA, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTACCOUNTINGPERIOD)ACCOUNTINGPERIOD, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTACCOUNTINGYEAR)ACCOUNTINGYEAR, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE, " &_
					" MAX(DTL.JOBCODE_HUTANG_PLASMA) JOBCODE, " &_
					" 'PL' SOURCE, " &_
					" 'PL' SOURCETYPE, " &_
					" MAX(DTL.NAMA_KOPERASI) SOURCECODE, " &_
					" 'CB' TARGETTYPE, " &_
					" 'BS' TARGETCODE, " &_
					" MAX(DTL.NAMA_KOPERASI) DOCUMENTNO, " &_
					" 0 QUANTITY, " &_
					" 0 OUTPUT, " &_
					" SUM(DTL.AMOUNT_PLASMA_RECONCILIATION) * -1 AMOUNT, " &_
					" 'TC - Nama Koperasi No :' TITLE_REFERENCE_01_DB,  " &_
					" MAX(DTL.NAMA_KOPERASI) AUTO_REFERENCE_01_DB,  " &_
					" '' TITLE_REFERENCE_02_DB,  " &_
					" '' AUTO_REFERENCE_02_DB  " &_
					" FROM  " &_
					" IPLASPROD.PLASMA_RECONCILIATION_VIEW DTL " &_
					" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='PLASMA_RECONSIL' " &_
					" WHERE " &_
					" to_number(to_char(DTL.CLOSING_DATE,'mm')) = " & Mid(Session("CurP_PlasmaReconsil"),5,2) &_
					" AND to_number(to_char(DTL.CLOSING_DATE,'yyyy')) = " & Mid(Session("CurP_PlasmaReconsil"),1,4) & _
					" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
					" AND DTL.DEBETCREDIT = 'Cr' " &_
					" GROUP BY DTL.FCBA,DTL.NAMA_KOPERASI,DTL.JOBCODE_PIUTANG_PLASMA, DTL.JOBCODE_HUTANG_PLASMA " &_
					" ) ORDER BY LINENO "
		'response.write qPlasmaDB
		Set rsPlasmaMinDB = Server.CreateObject("ADODB.Recordset")
		rsPlasmaMinDB.Open qPlasmaMinDB, connstr
				
		if not rsPlasmaMinDB.eof then
			Do While Not rsPlasmaMinDB.eof				
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
								 " values (" & maxLineNo & ",to_date('" & rsPlasmaMinDB("TDATE") & "','mm/dd/yyyy') " & ",'" & rsPlasmaMinDB("JOBCODE") & "','" & rsPlasmaMinDB("SOURCE") & "','" & rsPlasmaMinDB("SOURCETYPE") & "','" & rsPlasmaMinDB("SOURCECODE") & "'," & _
								 "'','" & rsPlasmaMinDB("TARGETTYPE") & "','" & rsPlasmaMinDB("TARGETCODE") & "'," & _
								 "'" &  rsPlasmaMinDB("AMOUNT") & "','" & rsPlasmaMinDB("DOCUMENTNO") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsPlasmaMinDB("TITLE_REFERENCE_01_DB") & "' ,'"  & rsPlasmaMinDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsPlasmaMinDB("TITLE_REFERENCE_02_DB") & "' ,'"  & rsPlasmaMinDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '','', " & _
								 " '','', " & _
								 " 'Line No:', '" & rsPlasmaMinDB("LINENO") & "','Db', " & _
								 " " & rsPlasmaMinDB("ACCOUNTINGPERIOD") & "," & rsPlasmaMinDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDB)
				 connStr.execute("Commit")
				'-- End Insert Debet
				
			rsPlasmaMinDB.MoveNext
		  loop
		  response.Write "Process Plasma Reconciliation Minus Debet Complete.... !!<br>"
		  response.Flush
		end if
		'-- End Insert Debet
		
		'-- Insert Credit
		qPlasmaMinCR =" SELECT" &_				
					" ROWNUM LINENO,FCBA, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
					" TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE," &_  
					" TARGETTYPE, TARGETCODE, DOCUMENTNO, QUANTITY, OUTPUT, AMOUNT, " &_
					" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB, " &_
					" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB  " &_
					" FROM " &_
					" ( " &_		 
					" SELECT " &_
					" MAX(ROWNUM) LINENO, " &_
					" MAX(DTL.FCBA)FCBA, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTACCOUNTINGPERIOD)ACCOUNTINGPERIOD, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTACCOUNTINGYEAR)ACCOUNTINGYEAR, " &_
					" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE, " &_
					" MAX(DTL.JOBCODE_PIUTANG_PLASMA) JOBCODE, " &_
					" 'PL' SOURCE, " &_
					" 'PL' SOURCETYPE, " &_
					" MAX(DTL.NAMA_KOPERASI) SOURCECODE, " &_
					" 'CB' TARGETTYPE, " &_
					" 'BS' TARGETCODE, " &_
					" MAX(DTL.NAMA_KOPERASI) DOCUMENTNO, " &_
					" 0 QUANTITY, " &_
					" 0 OUTPUT, " &_
					" SUM(DTL.AMOUNT_PLASMA_RECONCILIATION) AMOUNT , " &_
					" 'TC - Nama Koperasi No :' TITLE_REFERENCE_01_DB,  " &_
					" MAX(DTL.NAMA_KOPERASI) AUTO_REFERENCE_01_DB,  " &_
					" '' TITLE_REFERENCE_02_DB,  " &_
					" '' AUTO_REFERENCE_02_DB  " &_
					" FROM  " &_
					" IPLASPROD.PLASMA_RECONCILIATION_VIEW DTL " &_
					" INNER JOIN IPLASPROD.PERIODCONTROL ON DTL.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='PLASMA_RECONSIL' " &_
					" WHERE " &_
					" to_number(to_char(DTL.CLOSING_DATE,'mm')) = " & Mid(Session("CurP_PlasmaReconsil"),5,2) &_
					" AND to_number(to_char(DTL.CLOSING_DATE,'yyyy')) = " & Mid(Session("CurP_PlasmaReconsil"),1,4) & _
					" AND DTL.FCBA = '" & Session("fcBA") & "'" &_
					" AND DTL.DEBETCREDIT = 'Cr' " &_
					" GROUP BY DTL.FCBA,DTL.NAMA_KOPERASI,DTL.JOBCODE_PIUTANG_PLASMA, DTL.JOBCODE_HUTANG_PLASMA " &_
					" ) ORDER BY LINENO "
		'response.write qPlasmaCR
		Set rsPlasmaMinCR = Server.CreateObject("ADODB.Recordset")
		rsPlasmaMinCR.Open qPlasmaMinCR, connstr
				
		if not rsPlasmaMinCR.eof then
			Do While Not rsPlasmaMinCR.eof				
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
								 " values (" & maxLineNo & ",to_date('" & rsPlasmaMinCR("TDATE") & "','mm/dd/yyyy') " & ",'" & rsPlasmaMinCR("JOBCODE") & "','" & rsPlasmaMinCR("SOURCE") & "','" & rsPlasmaMinCR("SOURCETYPE") & "','" & rsPlasmaMinCR("SOURCECODE") & "'," & _
								 "'','" & rsPlasmaMinCR("TARGETTYPE") & "','" & rsPlasmaMinCR("TARGETCODE") & "'," & _
								 "'" &  rsPlasmaMinCR("AMOUNT") & "','" & rsPlasmaMinCR("DOCUMENTNO") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsPlasmaMinCR("TITLE_REFERENCE_01_DB") & "' ,'"  & rsPlasmaMinCR("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsPlasmaMinCR("TITLE_REFERENCE_02_DB") & "' ,'"  & rsPlasmaMinCR("AUTO_REFERENCE_02_DB") & "'," & _
								 " '','', " & _
								 " '','', " & _
								 " 'Line No:', '" & rsPlasmaMinCR("LINENO") & "','Cr', " & _
								 " " & rsPlasmaMinCR("ACCOUNTINGPERIOD") & "," & rsPlasmaMinCR("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDB)
				 connStr.execute("Commit")
				'-- End Insert Debet
				
			rsPlasmaMinCR.MoveNext
		  loop
		  response.Write "Process Plasma Reconciliation Minus Credit Complete.... !!<br>"
		  response.Flush
		end if
		'-- End Insert Credit
		
		'-- Begin Plasma Reconsil Minus (-)
	end if
	
	
	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))
	
	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND UPPER(SYSTEM) = 'PLASMA_RECONSIL'" 

		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME, fcentry) VALUES (" & _
					" 'Plasma_Reconsil', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy HH:MI:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy HH:MI:ssam')," & _
					" " & tempMaxLine & ", sysdate, to_char(sysdate,'HH24:MI')" & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Plasma_Reconsil'"
	
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.Open strLenghtime, connstr
	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Plasma_Reconsil'"		
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
								" where UPPER(system) ='PLASMA_RECONSIL' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for PLASMA RECONCILIATION " & "<br>"
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
				  " where SYSTEM='Plasma_Reconsil' and fcba='" & Session("fcBA") & "'"
				ww = " Satu " 
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where SYSTEM='Plasma_Reconsil' and fcba='" & Session("fcBA") & "'"
					  ww = " Dua " 
				ELSE
				    IF temp_NextCloseAccountingPeriod = 12 THEN
						 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
						  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear= " & temp_NextCurrentAccountingYear & " " & _  
						  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						  " where  SYSTEM='Plasma_Reconsil' AND fcba='" & Session("fcBA") & "'"
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

			' Query ke Period Control untuk mengambli Session(v_PlasmaReconsil) & Session(cls_PlasmaReconsil)	
		
		
			txtLastDay  =cmdFindPeriod("Plasma_Reconsil")
			txtNextMonth=cmdPeriodClose("Plasma_Reconsil")

			Session("v_PlasmaReconsil")= txtLastDay
			Session("cls_PlasmaReconsil")= txtNextMonth

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
Plasma Reconsil Last Period Close <%=FormatDateTime(Session("cls_PlasmaReconsil"),1)%> <br> 
Plasma Reconsil Current Period is <%=FormatDateTime(Session("v_PlasmaReconsil"),1)%> <br><br>
Please Wait...<br>
<%
 Response.Write "Start ..." & Time() & "<br>" 
 
 ' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Plasma_Reconsil','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Plasma Reconsil " & "<br>"
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
			<img src="images/ok.gif"/>Plasma Reconsil is Completed!<br/>
		</font>
<%
else
		Call PostEndTime
%>
		<br/><br/>
		<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>
			<img src="images/error.gif"/>Plasma Reconsil is Not Completed!!!<br/>
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