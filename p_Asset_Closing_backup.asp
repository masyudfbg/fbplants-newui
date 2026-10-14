<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="adovbs.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Asset Closing</title>
</head>
<body>
<%
Session("v_Asset") = cmdFindPeriod("Asset")
Session("cls_Asset") = cmdPeriodClose("Asset")

Session("CurP_Asset") = cmdFindAccountPeriod("Asset")
Session("ClsP_Asset") = cmdAccountPeriodClose("Asset")

Sub NormalisasiData
	
	arr_normalisasi_eof = true
	qPeriodControl = " select lineno from cost_activity " & _
				     " where ACCOUNTINGPERIOD = " & Mid(Session("CurP_Asset"),5,2) & "" & _
				     "       and ACCOUNTINGYEAR = " & Mid(Session("CurP_Asset"),1,4) & "" & _
				     "       and fcBA = '" & Session("fcBA") & "'" & _
				     "       and Sourcetype = 'FA'"

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
	
	'Delete Next Month Ledger
	qNMLedger = " DELETE FROM FIXEDASSET_DEPRECIATION_LEDGER " &_
				" WHERE to_number(to_Char(DATE_TRANSACTION,'yyyy')) = "& Cdbl(Mid(Session("CurP_Asset"),1,4)) &" AND  " &_ 
				" to_number(to_Char(DATE_TRANSACTION,'mm')) = "& Cdbl(Mid(Session("CurP_Asset"),5,2))+1 &" AND " &_
				" fcba =  '" & Session("fcBA") & "' AND " &_
				" BEGINING_BALANCE = 'FALSE'"
	connstr.execute(qNMLedger)
	
	qNMLedgerAdd = " DELETE FROM FIXEDASSET_DEPRECIATION_LEDGER " &_
				   " WHERE FCTYPE = 'ADDITION' AND " &_ 
				   " to_number(to_Char(DATE_TRANSACTION,'yyyy')) = "& Cdbl(Mid(Session("CurP_Asset"),1,4)) &" AND " &_ 
				   " to_number(to_Char(DATE_TRANSACTION,'mm')) = "& Cdbl(Mid(Session("CurP_Asset"),5,2)) &" AND " &_ 
				   " fcba =  '" & Session("fcBA") & "' AND " &_
				   " BEGINING_BALANCE = 'FALSE'"
	connstr.execute(qNMLedgerAdd)
	
	'Normalisasi FIXEDASSET
	qFAHistA =	" SELECT  " &_
				" FCBA, TRANSDATE, FCCODE, FCTYPE, NVL(AMOUNT,0) AMOUNT " &_ 
				" FROM " &_
				" FIXEDASSET_HISTORY " &_
				" WHERE " &_
				" to_number(to_Char(TRANSDATE,'yyyy')) = "& Cdbl(Mid(Session("CurP_Asset"),1,4)) & _
				" AND to_number(to_Char(TRANSDATE,'mm')) = "& Cdbl(Mid(Session("CurP_Asset"),5,2))  & _
				" AND FCBA = '" & Session("fcBA") & "'"
	
	Set rsFAHistA = Server.CreateObject("ADODB.Recordset")
	rsFAHistA.Open qFAHistA, connstr
	
	if not rsFAHistA.eof then
		Do While Not rsFAHistA.eof
			
			if rsFAHistA("FCTYPE") = "FROM STRAIGHT LINE" then 
				qTYPEAC= "UPDATE FIXEDASSET SET ACQUISITIONVALUE = 0 "&_
						 " WHERE FCBA = '" & rsFAHistA("FCBA") & "' AND "&_
						 " FCCODE = '" & rsFAHistA("FCCODE") & "' AND "&_
						 " TO_CHAR(TRANSDATE,'fmMM/dd/yyyy') = '" & rsFAHistA("TRANSDATE") & "'" 
				connstr.execute(qTYPEAC)
				connStr.execute("Commit")
			end if
			
			if rsFAHistA("FCTYPE") = "ADDITION" then
				qTYPEAD= "UPDATE FIXEDASSET SET ADDITIONALVALUE = " &_
						" CASE WHEN " & Cdbl(rsFAHistA("AMOUNT")) & " > 0 THEN " &_
						" ADDITIONALVALUE - decode("& Cdbl(rsFAHistA("AMOUNT")) &",null,0,"& Cdbl(rsFAHistA("AMOUNT")) &")" &_
						" ELSE " &_
						" 0 " &_
						" END " &_
				" WHERE FCBA = '" & rsFAHistA("FCBA") & "' AND FCCODE = '" & rsFAHistA("FCCODE") & "' AND TO_CHAR(TRANSDATE,'fmMM/dd/yyyy') = '" & rsFAHistA("TRANSDATE") & "'"
				'response.write qTYPEAD
				connstr.execute(qTYPEAD)
				connStr.execute("Commit")
			end if
			
			rsFAHistA.MoveNext
		loop
		response.Write "Normalisasi Fixed Asset.... <br>"
		response.Flush
	end if
	
	'Delete Next FIXEDASSET hist
	qFAHIST = " DELETE FROM FIXEDASSET_HISTORY WHERE to_number(to_Char(TRANSDATE,'yyyy')) = "& Cdbl(Mid(Session("CurP_Asset"),1,4)) &" AND to_number(to_Char(TRANSDATE,'mm')) = "& Cdbl(Mid(Session("CurP_Asset"),5,2)) &" AND FCBA =  '" & Session("fcBA") & "'"
	connstr.execute(qFAHIST)
End Sub

Sub postingData
	' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	' Add by NAF to select lastperiodposted
	Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl
    
	qPeriodControl="select add_months(LASTPERIODPOSTED,1) LastPeriodPosted, LASTPERIODPOSTED LastEndPeriodPosted from PERIODCONTROL " & _
	  " where SYSTEM='Asset' and fcba='" & Session("fcBA") & "'"
	
	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr
	
	if not rsPeriodControl.eof then
	   IF  Cint(MID(Session("CurP_Asset"),5,2)) <= 12 Then
		   temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")     ' add Akbar 28 Des 2015 , for Account Period <= 12
	   ELSE
	       temp_datePeriodControl = rsPeriodControl("LASTENDPERIODPOSTED")  ' for Account Period > 12
       END IF
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Asset " & "<br>"
	end if	
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	' Remarked by NAF, as per New Concept, 7 May 2013
	' qControlJob="select ControlJob from Module where NAME='Purchase'"
	' Now we use new concept for picking Control Job, by using Process Control Tables
	' as per last concept are NOT able to pick data per Business Unit.
	
	tempProcessCode = "ASSET_CLOSING"
	tempProcessNameCode = "p_Asset_Closing.asp"
	tempRefCodePVDetail = "PAYMENTVOUCHERDETAIL"

	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
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
	
	'-- Begin Of AUCH
	'-- Insert Debet
	qAssetAUCHDB =	" SELECT " &_
					" ROWNUM LINENO, " &_ 
					" TO_CHAR (creationdate, 'fmMM') ACCOUNTINGPERIOD, " &_
					" TO_CHAR (creationdate, 'yyyy') ACCOUNTINGYEAR, " &_		 
					" IPLASPROD.PERIODCONTROL.CURRENTDATE TDATE_DB, " &_
					" IPLASPROD.ASSETGROUP.CONTROLJOB JOBCODE_DB, " &_
					" 'FA' SOURCE_DB, " &_
					" 'FA' SOURCETYPE_DB, " &_
					" HDR.FIXEDASSETCODE SOURCECODE_DB, " &_
					" 'FA' TARGETTYPE_DB, " &_
					" HDR.FIXEDASSETCODE TARGETCODE_DB, " &_
					" HDR.DOCUMENTNO DOCUMENTNO_DB, " &_
					" 0 QUANTITY_DB, " &_
					" 0 OUTPUT_DB, " &_
					" HDR.AMOUNT AMOUNT_DB, " &_
					" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_
					" HDR.FIXEDASSETCODE AUTO_REFERENCE_01_DB, " &_
					" 'Auto - Fixed Asset AUC No :' TITLE_REFERENCE_02_DB, " &_
					" HDR.DOCUMENTNO AUTO_REFERENCE_02_DB, " &_
					" '' TITLE_REFERENCE_03_DB, " &_
					" '' AUTO_REFERENCE_03_DB " &_
					" FROM " &_
					" ( " &_
					" Select HDR.FCBA, HDR.DOCUMENTNO, HDR.CREATIONDATE, HDR.FIXEDASSETCODE, SUM(DTL.AMOUNT) AMOUNT " &_
					" FROM " &_
					" IPLASPROD.FIXEDASSET_AUC_DTL DTL " &_
					" INNER JOIN IPLASPROD.FIXEDASSET_AUC_HDR HDR ON DTL.DOCUMENTNO = HDR.DOCUMENTNO AND HDR.FCBA = DTL.FCBA " &_
					" GROUP BY HDR.FCBA, HDR.DOCUMENTNO, HDR.CREATIONDATE, HDR.FIXEDASSETCODE " &_
					" )HDR " &_
					" INNER JOIN IPLASPROD.FIXEDASSET ON HDR.FIXEDASSETCODE = IPLASPROD.FIXEDASSET.FCCODE AND HDR.FCBA = IPLASPROD.FIXEDASSET.FCBA " &_
					" INNER JOIN IPLASPROD.ASSETGROUP on IPLASPROD.FIXEDASSET.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
					" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET' " &_
					" WHERE " &_
					" TO_NUMBER(TO_CHAR(HDR.CREATIONDATE,'mm')) = " & Mid(Session("CurP_Asset"),5,2) &_ 
					" AND TO_NUMBER(TO_CHAR(HDR.CREATIONDATE,'yyyy')) = " & Mid(Session("CurP_Asset"),1,4) &_
					" AND HDR.FCBA =  '" & Session("fcBA") & "'"

	Set rsAssetAUCHDB = Server.CreateObject("ADODB.Recordset")
	rsAssetAUCHDB.Open qAssetAUCHDB, connstr
	if not rsAssetAUCHDB.eof then
		Do While Not rsAssetAUCHDB.eof	   
			 
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityDB=	"Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsAssetAUCHDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsAssetAUCHDB("JOBCODE_DB") & "','" & rsAssetAUCHDB("SOURCE_DB") & "','" & rsAssetAUCHDB("SOURCETYPE_DB") & "','" & rsAssetAUCHDB("SOURCECODE_DB") & "'," & _
								 "'','" & rsAssetAUCHDB("TARGETTYPE_DB") & "','" & rsAssetAUCHDB("TARGETCODE_DB") & "'," & _
								 "'" &  rsAssetAUCHDB("AMOUNT_DB") & "','" & rsAssetAUCHDB("DOCUMENTNO_DB") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsAssetAUCHDB("TITLE_REFERENCE_01_DB") & "' ,'"  & rsAssetAUCHDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsAssetAUCHDB("TITLE_REFERENCE_02_DB") & "' ,'"  & rsAssetAUCHDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '','', " & _
								 " '','', " & _
								 " 'Line No:', '" & rsAssetAUCHDB("LINENO") & "','Db', " & _
								 " " & rsAssetAUCHDB("ACCOUNTINGPERIOD") & "," & rsAssetAUCHDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDB)
			 connStr.execute("Commit")

		  rsAssetAUCHDB.MoveNext
	  loop
	  response.Write "Process FA AUC Debet Complete.... !!<br>"
	  response.Flush
	end if
	
	'-- End Insert Debet
	
	'-- Insert Credit
	qAssetAUCHCR =	" SELECT " &_
					" ROWNUM LINENO, " &_
					" TO_CHAR (creationdate, 'fmMM') ACCOUNTINGPERIOD, " &_
					" TO_CHAR (creationdate, 'yyyy') ACCOUNTINGYEAR, " &_			 
					" IPLASPROD.PERIODCONTROL.CURRENTDATE TDATE_CR, " &_
					" IPLASPROD.PROJECTACC.CONTROLJOBCODE JOBCODE_CR, " &_
					" 'FA' SOURCE_CR, " &_
					" 'FA' SOURCETYPE_CR, " &_
					" HDR.FIXEDASSETCODE SOURCECODE_CR, " &_
					" 'PJ' TARGETTYPE_CR, " &_
					" IPLASPROD.PROJECTACC.FCCODE TARGETCODE_CR, " &_
					" HDR.DOCUMENTNO DOCUMENTNO_CR, " &_
					" 0 QUANTITY_CR, " &_
					" 0 OUTPUT_CR, " &_
					" -1 * HDR.AMOUNT AMOUNT_CR, " &_
					" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR, " &_
					" HDR.FIXEDASSETCODE AUTO_REFERENCE_01_CR, " &_
					" 'Auto - Fixed Asset AUC No :' TITLE_REFERENCE_02_CR, " &_
					" HDR.DOCUMENTNO AUTO_REFERENCE_02_CR, " &_
					" '' TITLE_REFERENCE_03_CR, " &_
					" '' AUTO_REFERENCE_03_CR " &_
					" FROM " &_
					" ( " &_
					" Select HDR.FCBA, HDR.DOCUMENTNO, HDR.CREATIONDATE, HDR.FIXEDASSETCODE, SUM(DTL.AMOUNT) AMOUNT " &_
					" FROM " &_
					" IPLASPROD.FIXEDASSET_AUC_DTL DTL " &_
					" INNER JOIN IPLASPROD.FIXEDASSET_AUC_HDR HDR ON DTL.DOCUMENTNO = HDR.DOCUMENTNO AND HDR.FCBA = DTL.FCBA " &_
					" GROUP BY HDR.FCBA, HDR.DOCUMENTNO, HDR.CREATIONDATE, HDR.FIXEDASSETCODE " &_
					" )HDR " &_
					" INNER JOIN IPLASPROD.FIXEDASSET ON HDR.FIXEDASSETCODE = IPLASPROD.FIXEDASSET.FCCODE AND HDR.FCBA = IPLASPROD.FIXEDASSET.FCBA " &_
					" INNER JOIN IPLASPROD.ASSETGROUP ON IPLASPROD.FIXEDASSET.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
					" INNER JOIN IPLASPROD.PROJECTACC ON HDR.FIXEDASSETCODE = IPLASPROD.PROJECTACC.FCCODE AND HDR.FCBA = IPLASPROD.PROJECTACC.FCBA " &_
					" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET' " &_
					" WHERE " &_
					" TO_NUMBER(TO_CHAR(HDR.CREATIONDATE,'mm')) = " & Mid(Session("CurP_Asset"),5,2) &_ 
					" AND TO_NUMBER(TO_CHAR(HDR.CREATIONDATE,'yyyy')) = " & Mid(Session("CurP_Asset"),1,4) &_
					" AND HDR.FCBA =  '" & Session("fcBA") & "'"

	Set rsAssetAUCHCR = Server.CreateObject("ADODB.Recordset")
	rsAssetAUCHCR.Open qAssetAUCHCR, connstr
	if not rsAssetAUCHCR.eof then
		Do While Not rsAssetAUCHCR.eof	   
			 
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
							 " values (" & maxLineNo & ",to_date('" & rsAssetAUCHCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsAssetAUCHCR("JOBCODE_CR") & "','" & rsAssetAUCHCR("SOURCE_CR") & "','" & rsAssetAUCHCR("SOURCETYPE_CR") & "','" & rsAssetAUCHCR("SOURCECODE_CR") & "'," & _
							 "'','" & rsAssetAUCHCR("TARGETTYPE_CR") & "','" & rsAssetAUCHCR("TARGETCODE_CR") & "'," & _
							 "'" &  rsAssetAUCHCR("AMOUNT_CR") & "','" & rsAssetAUCHCR("DOCUMENTNO_CR") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsAssetAUCHCR("TITLE_REFERENCE_01_CR") & "' ,'"  & rsAssetAUCHCR("AUTO_REFERENCE_01_CR") & "'," & _
							 " '" & rsAssetAUCHCR("TITLE_REFERENCE_02_CR") & "' ,'"  & rsAssetAUCHCR("AUTO_REFERENCE_02_CR") & "'," & _
							 " '','', " & _
							 " '','', " & _
							 " 'Line No:', '" & rsAssetAUCHCR("LINENO") & "','Cr', " & _
							 " " & rsAssetAUCHCR("ACCOUNTINGPERIOD") & "," & rsAssetAUCHCR("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDB)
			 connStr.execute("Commit")

		  rsAssetAUCHCR.MoveNext
	  loop
	  response.Write "Process FA AUC Credit Complete.... !!<br>"
	  response.Flush
	end if
	
	'-- End Insert Credit
	'-- End Of AUCH

	'-- Begin Of FA Adjusment Addition	
	'-- Insert Debet
	qAssetFAA =" SELECT " &_
				" ROWNUM LINENO, " &_ 
				" TO_CHAR (HDR.TRANSDATE, 'fmMM') ACCOUNTINGPERIOD, " &_
				" TO_CHAR (HDR.TRANSDATE, 'yyyy') ACCOUNTINGYEAR, " &_ 
				" IPLASPROD.PERIODCONTROL.CURRENTDATE TDATE_DB, " &_
				" IPLASPROD.ASSETGROUP.CONTROLJOB JOBCODE_DB, " &_
				" 'FA' SOURCE_DB, " &_
				" 'FA' SOURCETYPE_DB, " &_
				" HDR.FIXEDASSETCODE SOURCECODE_DB, " &_
				" 'FA' TARGETTYPE_DB, " &_
				" HDR.FIXEDASSETCODE TARGETCODE_DB, " &_
				" HDR.FCCODE DOCUMENTNO_DB, " &_
				" 0 QUANTITY_DB, " &_
				" 0 OUTPUT_DB, " &_
				" HDR.ADJUSTMENT_AMOUNT AMOUNT_DB, " &_
				" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_
				" HDR.FIXEDASSETCODE AUTO_REFERENCE_01_DB, " &_
				" 'Auto - Fixed Asset Adjustment No :' TITLE_REFERENCE_02_DB, " &_
				" HDR.FCCODE AUTO_REFERENCE_02_DB, " &_
				" '' TITLE_REFERENCE_03_DB, " &_
				" '' AUTO_REFERENCE_03_DB " &_
				" FROM " &_
				" IPLASPROD.FIXEDASSET_ADJUSTMENT HDR " &_
				" INNER JOIN IPLASPROD.FIXEDASSET ON HDR.FIXEDASSETCODE = IPLASPROD.FIXEDASSET.FCCODE AND HDR.FCBA = IPLASPROD.FIXEDASSET.FCBA " &_
				" INNER JOIN IPLASPROD.ASSETGROUP on IPLASPROD.FIXEDASSET.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET' " &_
				" WHERE " &_
				" HDR.FCTYPE = 'ADDITION' " &_
				" AND TO_NUMBER(TO_CHAR(HDR.TRANSDATE,'mm')) = " & Mid(Session("CurP_Asset"),5,2) &_ 
				" AND TO_NUMBER(TO_CHAR(HDR.TRANSDATE,'yyyy')) = " & Mid(Session("CurP_Asset"),1,4) &_
				" AND HDR.FCBA =  '" & Session("fcBA") & "'"

	Set rsAssetFAA = Server.CreateObject("ADODB.Recordset")
	rsAssetFAA.Open qAssetFAA, connstr
	if not rsAssetFAA.eof then
		Do While Not rsAssetFAA.eof	   
			 
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
							 " values (" & maxLineNo & ",to_date('" & rsAssetFAA("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsAssetFAA("JOBCODE_DB") & "','" & rsAssetFAA("SOURCE_DB") & "','" & rsAssetFAA("SOURCETYPE_DB") & "','" & rsAssetFAA("SOURCECODE_DB") & "'," & _
							 "'','" & rsAssetFAA("TARGETTYPE_DB") & "','" & rsAssetFAA("TARGETCODE_DB") & "'," & _
						 "'" &  rsAssetFAA("AMOUNT_DB") & "','" & rsAssetFAA("DOCUMENTNO_DB") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
						 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
						 " '" & rsAssetFAA("TITLE_REFERENCE_01_DB") & "' ,'"  & rsAssetFAA("AUTO_REFERENCE_01_DB") & "'," & _
							 " '" & rsAssetFAA("TITLE_REFERENCE_02_DB") & "' ,'"  & rsAssetFAA("AUTO_REFERENCE_02_DB") & "'," & _
							 " '','', " & _
							 " '','', " & _
						 " 'Line No:', '" & rsAssetFAA("LINENO") & "','Db', " & _
							 " " & rsAssetFAA("ACCOUNTINGPERIOD") & "," & rsAssetFAA("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
	
				 connstr.execute(qCost_ActivityDB)
				 connStr.execute("Commit")
	
		  rsAssetFAA.MoveNext
	  loop
	  response.Write "Process FA Adjusment Additional Debet Complete.... !!<br>"
	  response.Flush
	end if
	
	'-- End Insert Debet
	
	'-- Insert Credit
	qAssetFAACR =" SELECT " &_
				" ROWNUM LINENO, " &_
				"  TO_CHAR (HDR.TRANSDATE, 'fmMM') ACCOUNTINGPERIOD, " &_
				"  TO_CHAR (HDR.TRANSDATE, 'yyyy') ACCOUNTINGYEAR, " &_
				"  IPLASPROD.PERIODCONTROL.CURRENTDATE TDATE_CR, " &_
				"  HDR.JOBCODE JOBCODE_CR, " &_
				"  'FA' SOURCE_CR, " &_
				"  'FA' SOURCETYPE_CR, " &_
				" HDR.FIXEDASSETCODE SOURCECODE_CR, " &_
				" HDR.TARGETTYPE TARGETTYPE_CR, " &_
				" HDR.TARGETCODE TARGETCODE_CR, " &_
				" HDR.FCCODE DOCUMENTNO_CR, " &_
				" 0 QUANTITY_CR, " &_
				" 0 OUTPUT_CR, " &_
				" -1*HDR.ADJUSTMENT_AMOUNT AMOUNT_CR, " &_
				" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR, " &_
				" HDR.FIXEDASSETCODE AUTO_REFERENCE_01_CR, " &_
				" 'Auto - Fixed Asset Adjustment No :' TITLE_REFERENCE_02_CR, " &_
				" HDR.FCCODE AUTO_REFERENCE_02_CR, " &_
				" '' TITLE_REFERENCE_03_CR, " &_
				" '' AUTO_REFERENCE_03_CR " &_
				" FROM " &_
				" IPLASPROD.FIXEDASSET_ADJUSTMENT HDR " &_
				" INNER JOIN IPLASPROD.FIXEDASSET ON HDR.FIXEDASSETCODE = IPLASPROD.FIXEDASSET.FCCODE AND HDR.FCBA = IPLASPROD.FIXEDASSET.FCBA " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET' " &_
				" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND HDR.JOBCODE = IPLASPROD.JOB.FCCODE " &_
				" WHERE " &_
				" HDR.FCTYPE = 'ADDITION' " &_
				" AND TO_NUMBER(TO_CHAR(HDR.TRANSDATE,'mm')) = " & Mid(Session("CurP_Asset"),5,2) &_ 
				" AND TO_NUMBER(TO_CHAR(HDR.TRANSDATE,'yyyy')) = " & Mid(Session("CurP_Asset"),1,4) &_
				" AND HDR.FCBA =  '" & Session("fcBA") & "'"
	
	Set rsAssetFAACR = Server.CreateObject("ADODB.Recordset")
	rsAssetFAACR.Open qAssetFAACR, connstr
	if not rsAssetFAACR.eof then
		Do While Not rsAssetFAACR.eof	   
			 
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
							 " values (" & maxLineNo & ",to_date('" & rsAssetFAACR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsAssetFAACR("JOBCODE_CR") & "','" & rsAssetFAACR("SOURCE_CR") & "','" & rsAssetFAACR("SOURCETYPE_CR") & "','" & rsAssetFAACR("SOURCECODE_CR") & "'," & _
							 "'','" & rsAssetFAACR("TARGETTYPE_CR") & "','" & rsAssetFAACR("TARGETCODE_CR") & "'," & _
							 "'" &  rsAssetFAACR("AMOUNT_CR") & "','" & rsAssetFAACR("DOCUMENTNO_CR") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsAssetFAACR("TITLE_REFERENCE_01_CR") & "' ,'"  & rsAssetFAACR("AUTO_REFERENCE_01_CR") & "'," & _
							 " '" & rsAssetFAACR("TITLE_REFERENCE_02_CR") & "' ,'"  & rsAssetFAACR("AUTO_REFERENCE_02_CR") & "'," & _
							 " '','', " & _
							 " '','', " & _
							 " 'Line No:', '" & rsAssetFAACR("LINENO") & "','Cr', " & _
							 " " & rsAssetFAACR("ACCOUNTINGPERIOD") & "," & rsAssetFAACR("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDB)
			 connStr.execute("Commit")

		  rsAssetFAACR.MoveNext
	 loop
	 response.Write "Process FA Adjusment Additional Complete.... !!<br>"
	 response.Flush
	 end if
	'-- End Insert Credit
	
	'-- End Of FA Adjusment Addition
	
	'-- Begin Of FA Adjusment Reduction
		'Samopai Sini Yah...  
	'-- End Of FA Adjusment Reduction
	
	'-- Begin Of FA WO Disposal 
		'Samopai Sini Yah...  
	'-- End Of FA WO Disposal
	
	'-- Begin Of FA WO Sales 
		'Samopai Sini Yah...  
	'-- End Of FA WO Sales
	
	'-- Begin Of FA Transport Task 
		'Samopai Sini Yah...  
	'-- End Of FA  Transport Task
	
	'-- Begin Of FA Receive Task
	
	'-- End Of FA  Receive Task	
	
	'-- Begin Of FA Maturity Inti ALL
		'-- Insert Debet
		qFAMatIntiALLDB =	"SELECT " &_
							" FCBA, ROWNUM LINENO," &_
							" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
							" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB, " &_
							" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB, " &_
							" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB, " &_
							" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB, " &_
							" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB  " &_
							" FROM " &_
							" ( " &_
							" SELECT " &_    
							" MAX(HDR.FCBA)FCBA, " &_
							" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD, " &_
							" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR, " &_
							" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB, " &_
							" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB) JOBCODE_DB, " &_
							" 'FA' SOURCE_DB, " &_
							" 'FA' SOURCETYPE_DB, " &_  
							" MAX(HDR.FIXEDASSETCODE) SOURCECODE_DB, " &_  
							" 'FA' TARGETTYPE_DB, " &_
							" MAX(HDR.FIXEDASSETCODE) TARGETCODE_DB, " &_  
							" MAX(HDR.DOCUMENTNO) DOCUMENTNO_DB, " &_
							" 0 QUANTITY_DB, " &_
							" 0 OUTPUT_DB, " &_
							" SUM(DTL.AMOUNT) AMOUNT_DB, " &_
							" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_
							" MAX(HDR.FIXEDASSETCODE) AUTO_REFERENCE_01_DB, " &_
							" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_DB, " &_ 
							" MAX(HDR.DOCUMENTNO) AUTO_REFERENCE_02_DB, " &_
							" '' TITLE_REFERENCE_03_DB, " &_
							" '' AUTO_REFERENCE_03_DB " &_
							" FROM " &_
							" IPLASPROD.FIXEDASSET_MATURE_DTL DTL " &_
							" INNER JOIN IPLASPROD.FIXEDASSET_MATURE HDR ON DTL.FCBA = HDR.FCBA AND DTL.DOCUMENTNO = HDR.DOCUMENTNO " &_
							" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FIXEDASSETCODE = FA.FCCODE " &_
							" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
							" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
							" WHERE " &_
							" HDR.FAMATURETYPE = 'ALL' " &_
							" AND TO_CHAR(HDR.CREATIONDATE,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
							" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
							" AND HDR.FCBA = '" & Session("fcBA") & "'" &_ 
							" GROUP BY HDR.FCBA, HDR.FIXEDASSETCODE)"
		
		Set rsFAMatIntiALLDB = Server.CreateObject("ADODB.Recordset")
		rsFAMatIntiALLDB.Open qFAMatIntiALLDB, connstr
		if not rsFAMatIntiALLDB.eof then
			Do While Not rsFAMatIntiALLDB.eof
				
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1= "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAMatIntiALLDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAMatIntiALLDB("JOBCODE_DB") & "','" & rsFAMatIntiALLDB("SOURCE_DB") & "','" & rsFAMatIntiALLDB("SOURCETYPE_DB") & "','" & rsFAMatIntiALLDB("SOURCECODE_DB") & "'," & _
									 " '','" & rsFAMatIntiALLDB("TARGETTYPE_DB") & "','" & rsFAMatIntiALLDB("TARGETCODE_DB") & "'," & _
									 " '" &  rsFAMatIntiALLDB("AMOUNT_DB") & "','" & rsFAMatIntiALLDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAMatIntiALLDB("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAMatIntiALLDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAMatIntiALLDB("AUTO_REFERENCE_01_DB") & "'," & _
									 " '" & rsFAMatIntiALLDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAMatIntiALLDB("AUTO_REFERENCE_02_DB") & "'," & _
									 " '" & rsFAMatIntiALLDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAMatIntiALLDB("AUTO_REFERENCE_03_DB") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAMatIntiALLDB("LINENO") & "','Db', " & _
									 " " & rsFAMatIntiALLDB("ACCOUNTINGPERIOD") & "," & rsFAMatIntiALLDB("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAMatIntiALLDB.MoveNext
			loop
		end if		
		'-- End Insert Debet
		
		'-- End Insert Credit
			qFAMatIntiALLCR =	"SELECT " &_
								" FCBA, ROWNUM LINENO," &_
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR," &_
								" TDATE_CR, JOBCODE_CR, SOURCE_CR, SOURCETYPE_CR, SOURCECODE_CR, TARGETTYPE_CR, TARGETCODE_CR," &_
								" DOCUMENTNO_CR, QUANTITY_CR, OUTPUT_CR, AMOUNT_CR," &_
								" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR," &_
								" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR," &_
								" TITLE_REFERENCE_03_CR, AUTO_REFERENCE_03_CR " &_
								" FROM" &_
								" (" &_
								" SELECT" &_    
								" MAX(HDR.FCBA)FCBA," &_  
								" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD," &_
								" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR," &_
								" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR," &_
								" MAX(IPLASPROD.PARAMETER.PARAM_1) JOBCODE_CR," &_
								" 'FA' SOURCE_CR," &_
								" 'FA' SOURCETYPE_CR," &_   
								" MAX(HDR.FIXEDASSETCODE) SOURCECODE_CR," &_   
								" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_CR," &_    
								" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_CR," &_
								" MAX(HDR.DOCUMENTNO) DOCUMENTNO_CR," &_
								" 0 QUANTITY_CR," &_
								" 0 OUTPUT_CR," &_
								" SUM(DTL.AMOUNT) AMOUNT_CR," &_  
								" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR," &_
								" MAX(HDR.FIXEDASSETCODE) AUTO_REFERENCE_01_CR," &_
								" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_CR," &_
								" MAX(HDR.DOCUMENTNO) AUTO_REFERENCE_02_CR," &_
								" '' TITLE_REFERENCE_03_CR," &_
								" '' AUTO_REFERENCE_03_CR" &_
								" FROM " &_
								" IPLASPROD.FIXEDASSET_MATURE_DTL DTL " &_ 
								" INNER JOIN IPLASPROD.FIXEDASSET_MATURE HDR ON DTL.FCBA = HDR.FCBA AND DTL.DOCUMENTNO = HDR.DOCUMENTNO " &_
								" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FIXEDASSETCODE = FA.FCCODE " &_
								" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
								" INNER JOIN IPLASPROD.PARAMETER ON HDR.FCBA = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'PROCREVERSED_001' " &_
								" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_1 = IPLASPROD.JOB.FCCODE" &_
								" INNER JOIN IPLASPROD.GCMASTER ON HDR.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY" &_				
								" WHERE " &_
								" HDR.FAMATURETYPE = 'ALL'" &_
								" AND TO_CHAR(HDR.CREATIONDATE,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
								" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
								" AND HDR.FCBA = '" & Session("fcBA") & "'" &_ 
								" GROUP BY HDR.FCBA, HDR.FIXEDASSETCODE)"
			
			Set rsFAMatIntiALLCR = Server.CreateObject("ADODB.Recordset")
			rsFAMatIntiALLCR.Open qFAMatIntiALLCR, connstr
			if not rsFAMatIntiALLCR.eof then
				Do While Not rsFAMatIntiALLCR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAMatIntiALLCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAMatIntiALLCR("JOBCODE_CR") & "','" & rsFAMatIntiALLCR("SOURCE_CR") & "','" & rsFAMatIntiALLCR("SOURCETYPE_CR") & "','" & rsFAMatIntiALLCR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAMatIntiALLCR("TARGETTYPE_CR") & "','" & rsFAMatIntiALLCR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAMatIntiALLCR("AMOUNT_CR")) & "','" & rsFAMatIntiALLCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAMatIntiALLCR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAMatIntiALLCR("TITLE_REFERENCE_01_CR") & "','"  & rsFAMatIntiALLCR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAMatIntiALLCR("TITLE_REFERENCE_02_CR") & "','"  & rsFAMatIntiALLCR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAMatIntiALLCR("TITLE_REFERENCE_03_CR") & "','"  & rsFAMatIntiALLCR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAMatIntiALLCR("LINENO") & "','Cr', " & _
									 " " & rsFAMatIntiALLCR("ACCOUNTINGPERIOD") & "," & rsFAMatIntiALLCR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAMatIntiALLCR.MoveNext
				loop
			end if
		'-- End Insert Credit
			response.Write "Process Fixed Asset Maturity ALL INTI Complete....<br>"
			response.Flush
		
	'-- End Of FA Maturity Inti ALL
	
	'-- Begin Of FA Maturity Inti PDE	
		'-- Insert Debet
		qFAMatIntiPDEDB =	"SELECT " &_
							" FCBA, ROWNUM LINENO," &_
							" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
							" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB, " &_
							" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB, " &_
							" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB, " &_
							" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB, " &_
							" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB  " &_
							" FROM " &_
							" ( " &_
							" SELECT " &_    
							" MAX(HDR.FCBA)FCBA, " &_
							" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD, " &_
							" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR, " &_
							" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB, " &_
							" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB) JOBCODE_DB, " &_
							" 'FA' SOURCE_DB, " &_
							" 'FA' SOURCETYPE_DB, " &_  
							" MAX(HDR.FIXEDASSETCODE) SOURCECODE_DB, " &_  
							" 'FA' TARGETTYPE_DB, " &_
							" MAX(HDR.FIXEDASSETCODE) TARGETCODE_DB, " &_  
							" MAX(HDR.DOCUMENTNO) DOCUMENTNO_DB, " &_
							" 0 QUANTITY_DB, " &_
							" 0 OUTPUT_DB, " &_
							" SUM(DTL.AMOUNT) AMOUNT_DB, " &_
							" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_
							" MAX(HDR.FIXEDASSETCODE) AUTO_REFERENCE_01_DB, " &_
							" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_DB, " &_ 
							" MAX(HDR.DOCUMENTNO) AUTO_REFERENCE_02_DB, " &_
							" '' TITLE_REFERENCE_03_DB, " &_
							" '' AUTO_REFERENCE_03_DB " &_
							" FROM " &_
							" IPLASPROD.FIXEDASSET_MATURE_DTL DTL " &_
							" INNER JOIN IPLASPROD.FIXEDASSET_MATURE HDR ON DTL.FCBA = HDR.FCBA AND DTL.DOCUMENTNO = HDR.DOCUMENTNO " &_
							" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FIXEDASSETCODE = FA.FCCODE " &_
							" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
							" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
							" WHERE " &_
							" HDR.FAMATURETYPE = 'PDE' " &_
							" AND TO_CHAR(HDR.CREATIONDATE,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
							" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
							" AND HDR.FCBA = '" & Session("fcBA") & "'" &_ 
							" GROUP BY HDR.FCBA, HDR.FIXEDASSETCODE)"
		
		Set rsFAMatIntiPDEDB = Server.CreateObject("ADODB.Recordset")
		rsFAMatIntiPDEDB.Open qFAMatIntiPDEDB, connstr
		if not rsFAMatIntiPDEDB.eof then
			Do While Not rsFAMatIntiPDEDB.eof
				
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1= "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAMatIntiPDEDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAMatIntiPDEDB("JOBCODE_DB") & "','" & rsFAMatIntiPDEDB("SOURCE_DB") & "','" & rsFAMatIntiPDEDB("SOURCETYPE_DB") & "','" & rsFAMatIntiPDEDB("SOURCECODE_DB") & "'," & _
									 " '','" & rsFAMatIntiPDEDB("TARGETTYPE_DB") & "','" & rsFAMatIntiPDEDB("TARGETCODE_DB") & "'," & _
									 " '" &  rsFAMatIntiPDEDB("AMOUNT_DB") & "','" & rsFAMatIntiPDEDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAMatIntiPDEDB("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAMatIntiPDEDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAMatIntiPDEDB("AUTO_REFERENCE_01_DB") & "'," & _
									 " '" & rsFAMatIntiPDEDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAMatIntiPDEDB("AUTO_REFERENCE_02_DB") & "'," & _
									 " '" & rsFAMatIntiPDEDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAMatIntiPDEDB("AUTO_REFERENCE_03_DB") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAMatIntiPDEDB("LINENO") & "','Db', " & _
									 " " & rsFAMatIntiPDEDB("ACCOUNTINGPERIOD") & "," & rsFAMatIntiPDEDB("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAMatIntiPDEDB.MoveNext
			loop
		end if		
		'-- End Insert Debet
		
		'-- End Insert Credit
			qFAMatIntiPDECR ="SELECT " &_
				" FCBA, ROWNUM LINENO," &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR," &_
				" TDATE_CR, JOBCODE_CR, SOURCE_CR, SOURCETYPE_CR, SOURCECODE_CR, TARGETTYPE_CR, TARGETCODE_CR," &_
				" DOCUMENTNO_CR, QUANTITY_CR, OUTPUT_CR, AMOUNT_CR," &_
				" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR," &_
				" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR," &_
				" TITLE_REFERENCE_03_CR, AUTO_REFERENCE_03_CR " &_
				" FROM" &_
				" (" &_
				" SELECT" &_    
				" MAX(HDR.FCBA)FCBA," &_  
				" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD," &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR," &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR," &_
				" MAX(IPLASPROD.PARAMETER.PARAM_1) JOBCODE_CR," &_
				" 'FA' SOURCE_CR," &_
				" 'FA' SOURCETYPE_CR," &_   
				" MAX(HDR.FIXEDASSETCODE) SOURCECODE_CR," &_   
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_CR," &_    
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_CR," &_
				" MAX(HDR.DOCUMENTNO) DOCUMENTNO_CR," &_
				" 0 QUANTITY_CR," &_
				" 0 OUTPUT_CR," &_
				" SUM(DTL.AMOUNT) AMOUNT_CR," &_  
				" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR," &_
				" MAX(HDR.FIXEDASSETCODE) AUTO_REFERENCE_01_CR," &_
				" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_CR," &_
				" MAX(HDR.DOCUMENTNO) AUTO_REFERENCE_02_CR," &_
				" '' TITLE_REFERENCE_03_CR," &_
				" '' AUTO_REFERENCE_03_CR" &_
				" FROM " &_
				" IPLASPROD.FIXEDASSET_MATURE_DTL DTL " &_ 
				" INNER JOIN IPLASPROD.FIXEDASSET_MATURE HDR ON DTL.FCBA = HDR.FCBA AND DTL.DOCUMENTNO = HDR.DOCUMENTNO " &_
				" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FIXEDASSETCODE = FA.FCCODE " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
				" INNER JOIN IPLASPROD.PARAMETER ON HDR.FCBA = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'PROCREVERSED_001' " &_
				" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_1 = IPLASPROD.JOB.FCCODE" &_
				" INNER JOIN IPLASPROD.GCMASTER ON HDR.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY" &_				
				" WHERE " &_
				" HDR.FAMATURETYPE = 'PDE'" &_
				" AND TO_CHAR(HDR.CREATIONDATE,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
				" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
				" AND HDR.FCBA = '" & Session("fcBA") & "'" &_ 
				" GROUP BY HDR.FCBA, HDR.FIXEDASSETCODE)"
			
			Set rsFAMatIntiPDECR = Server.CreateObject("ADODB.Recordset")
			rsFAMatIntiPDECR.Open qFAMatIntiPDECR, connstr
			if not rsFAMatIntiPDECR.eof then
				Do While Not rsFAMatIntiPDECR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAMatIntiPDECR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAMatIntiPDECR("JOBCODE_CR") & "','" & rsFAMatIntiPDECR("SOURCE_CR") & "','" & rsFAMatIntiPDECR("SOURCETYPE_CR") & "','" & rsFAMatIntiPDECR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAMatIntiPDECR("TARGETTYPE_CR") & "','" & rsFAMatIntiPDECR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAMatIntiPDECR("AMOUNT_CR")) & "','" & rsFAMatIntiPDECR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAMatIntiPDECR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAMatIntiPDECR("TITLE_REFERENCE_01_CR") & "','"  & rsFAMatIntiPDECR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAMatIntiPDECR("TITLE_REFERENCE_02_CR") & "','"  & rsFAMatIntiPDECR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAMatIntiPDECR("TITLE_REFERENCE_03_CR") & "','"  & rsFAMatIntiPDECR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAMatIntiPDECR("LINENO") & "','Cr', " & _
									 " " & rsFAMatIntiPDECR("ACCOUNTINGPERIOD") & "," & rsFAMatIntiPDECR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAMatIntiPDECR.MoveNext
				loop
			end if
		'-- End Insert Credit
			response.Write "Process Fixed Asset Maturity INTI PDE Complete....<br>"
			response.Flush
		
	'-- End Of FA Maturity Inti PDE

	'-- Begin Of FA Maturity Inti GC
	
	'-- Begin Insert Debet
		qFAMatIntiGCDB ="SELECT " &_
				" FCBA, ROWNUM LINENO, " &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_ 
				" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB, " &_ 
				" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB, " &_ 
				" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB, " &_ 
				" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB, " &_ 
				" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB  " &_ 
				" FROM " &_ 
				" (" &_
				" SELECT" &_ 
				" MAX(HDR.FCBA)FCBA, " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD, " &_ 
				" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR,  " &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB,  " &_
				" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB) JOBCODE_DB, " &_ 
				" 'FA' SOURCE_DB, " &_ 
				" 'FA' SOURCETYPE_DB, " &_   
				" MAX(HDR.FIXEDASSETCODE) SOURCECODE_DB, " &_   
				" 'FA' TARGETTYPE_DB, " &_ 
				" MAX(HDR.FIXEDASSETCODE) TARGETCODE_DB, " &_   
				" MAX(HDR.DOCUMENTNO) DOCUMENTNO_DB, " &_ 
				" 0 QUANTITY_DB, " &_ 
				" 0 OUTPUT_DB, " &_ 
				" SUM(DTL.AMOUNT) AMOUNT_DB, " &_ 
				" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_ 
				" MAX(HDR.FIXEDASSETCODE) AUTO_REFERENCE_01_DB, " &_ 
				" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_DB, " &_  
				" MAX(HDR.DOCUMENTNO) AUTO_REFERENCE_02_DB, " &_ 
				" '' TITLE_REFERENCE_03_DB, " &_ 
				" '' AUTO_REFERENCE_03_DB " &_ 
				" FROM " &_ 
				" IPLASPROD.FIXEDASSET_MATURE_DTL DTL " &_ 
				" INNER JOIN IPLASPROD.FIXEDASSET_MATURE HDR ON DTL.FCBA = HDR.FCBA AND DTL.DOCUMENTNO = HDR.DOCUMENTNO " &_ 
				" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FIXEDASSETCODE = FA.FCCODE " &_ 
				" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_ 
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_ 
				" WHERE " &_ 
				" HDR.FAMATURETYPE = 'GC'" &_
				" AND TO_CHAR(HDR.CREATIONDATE,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
				" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
				" AND HDR.FCBA = '" & Session("fcBA") & "'" &_ 
				" GROUP BY HDR.FCBA, HDR.FIXEDASSETCODE)"
		
		Set rsFAMatIntiGCDB = Server.CreateObject("ADODB.Recordset")
		rsFAMatIntiGCDB.Open qFAMatIntiGCDB, connstr
		if not rsFAMatIntiGCDB.eof then
			Do While Not rsFAMatIntiGCDB.eof
				
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsFAMatIntiGCDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAMatIntiGCDB("JOBCODE_DB") & "','" & rsFAMatIntiGCDB("SOURCE_DB") & "','" & rsFAMatIntiGCDB("SOURCETYPE_DB") & "','" & rsFAMatIntiGCDB("SOURCECODE_DB") & "'," & _
								 " '','" & rsFAMatIntiGCDB("TARGETTYPE_DB") & "','" & rsFAMatIntiGCDB("TARGETCODE_DB") & "'," & _
								 " '" &  rsFAMatIntiGCDB("AMOUNT_DB") & "','" & rsFAMatIntiGCDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAMatIntiGCDB("FCBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsFAMatIntiGCDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAMatIntiGCDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsFAMatIntiGCDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAMatIntiGCDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '" & rsFAMatIntiGCDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAMatIntiGCDB("AUTO_REFERENCE_03_DB") & "'," & _
								 " '','', " & _
								 " 'Line No:', '" & rsFAMatIntiGCDB("LINENO") & "','Db', " & _
								 " " & rsFAMatIntiGCDB("ACCOUNTINGPERIOD") & "," & rsFAMatIntiGCDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAMatIntiGCDB.MoveNext
			loop
		end if
	'-- End Insert Debet
	
	'-- Begin Insert Credit
	qFAMatIntiGCCR ="SELECT " &_
				" FCBA, ROWNUM LINENO, " &_ 
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_  
				" TDATE_CR, JOBCODE_CR, SOURCE_CR, SOURCETYPE_CR, SOURCECODE_CR, TARGETTYPE_CR, TARGETCODE_CR,  " &_ 
				" DOCUMENTNO_CR, QUANTITY_CR, OUTPUT_CR, AMOUNT_CR,  " &_
				" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR, " &_  
				" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR,  " &_
				" TITLE_REFERENCE_03_CR, AUTO_REFERENCE_03_CR " &_ 
				" FROM  " &_
				" (  " &_
				" SELECT   " &_ 
				" MAX(HDR.FCBA)FCBA,  " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'fmMM')) ACCOUNTINGPERIOD,  " &_
				" MAX(TO_CHAR(HDR.CREATIONDATE,'yyyy')) ACCOUNTINGYEAR,   " &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR,  " &_ 
				" MAX(IPLASPROD.PARAMETER.PARAM_1) JOBCODE_CR, " &_
				" 'FA' SOURCE_CR,  " &_ 
				" 'FA' SOURCETYPE_CR,   " &_
				" MAX(HDR.FIXEDASSETCODE) SOURCECODE_CR, " &_  
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_CR,   " &_ 
				" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_CR,  " &_ 
				" MAX(HDR.DOCUMENTNO) DOCUMENTNO_CR, " &_
				" 0 QUANTITY_CR,  " &_
				" 0 OUTPUT_CR,  " &_
				" SUM(DTL.AMOUNT) AMOUNT_CR,  " &_
				" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR," &_
				" MAX(HDR.FIXEDASSETCODE) AUTO_REFERENCE_01_CR," &_
				" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_CR," &_  
				" MAX(HDR.DOCUMENTNO) AUTO_REFERENCE_02_CR," &_
				" '' TITLE_REFERENCE_03_CR," &_
				" '' AUTO_REFERENCE_03_CR" &_  
				" FROM" &_
				" IPLASPROD.FIXEDASSET_MATURE_DTL DTL " &_ 
				" INNER JOIN IPLASPROD.FIXEDASSET_MATURE HDR ON DTL.FCBA = HDR.FCBA AND DTL.DOCUMENTNO = HDR.DOCUMENTNO " &_
				" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FIXEDASSETCODE = FA.FCCODE " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
				" INNER JOIN IPLASPROD.PARAMETER ON HDR.FCBA = IPLASPROD.PARAMETER.FCBA AND IPLASPROD.PARAMETER.PARCODE = 'PROCREVERSED_001' " &_
				" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.PARAMETER.PARAM_1 = IPLASPROD.JOB.FCCODE" &_
				" INNER JOIN IPLASPROD.GCMASTER ON HDR.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY" &_
				" WHERE " &_
				" HDR.FAMATURETYPE = 'GC'" &_
				" AND TO_CHAR(HDR.CREATIONDATE,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
				" AND TO_CHAR(HDR.CREATIONDATE,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
				" AND HDR.FCBA = '" & Session("fcBA") & "'" &_ 
				" GROUP BY HDR.FCBA, HDR.FIXEDASSETCODE)"

			Set rsFAMatIntiGCCR = Server.CreateObject("ADODB.Recordset")
			rsFAMatIntiGCCR.Open qFAMatIntiGCCR, connstr
			if not rsFAMatIntiGCCR.eof then
				Do While Not rsFAMatIntiGCCR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAMatIntiGCCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAMatIntiGCCR("JOBCODE_CR") & "','" & rsFAMatIntiGCCR("SOURCE_CR") & "','" & rsFAMatIntiGCCR("SOURCETYPE_CR") & "','" & rsFAMatIntiGCCR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAMatIntiGCCR("TARGETTYPE_CR") & "','" & rsFAMatIntiGCCR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAMatIntiGCCR("AMOUNT_CR")) & "','" & rsFAMatIntiGCCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAMatIntiGCCR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAMatIntiGCCR("TITLE_REFERENCE_01_CR") & "','"  & rsFAMatIntiGCCR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAMatIntiGCCR("TITLE_REFERENCE_02_CR") & "','"  & rsFAMatIntiGCCR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAMatIntiGCCR("TITLE_REFERENCE_03_CR") & "','"  & rsFAMatIntiGCCR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAMatIntiGCCR("LINENO") & "','Cr', " & _
									 " " & rsFAMatIntiGCCR("ACCOUNTINGPERIOD") & "," & rsFAMatIntiGCCR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAMatIntiGCCR.MoveNext
				loop
			end if
	'-- End Insert Credit
	
			response.Write "Process Fixed Asset Maturity INTI GC Complete....<br>"
			response.Flush
	'-- End Of FA Maturity Inti GC
	
	' Add By IWA 
	' Closing Asset 
	'-- Begin Of Generate Asset Ledger
		qGenerateAsset="Select ExecFADef('" & Session("fcBA") & "', " & Mid(Session("CurP_Asset"),5,2) & ", " & Mid(Session("CurP_Asset"),1,4) & ") STATUS from dual"		
		connstr.execute(qGenerateAsset)
		connStr.execute("Commit")
		Response.Write "Generate Asset Depreciation Ledger Complete" & "<br>" : response.flush()
	'-- End Of Generate Asset Ledger
	
	'-- Begin Of Depresiasi
	'-- Insert Debet
		If Session("Estate_Mill") = "HO" then		
			qFADepreciationDB =	" SELECT" &_
								" FCBA, ROWNUM LINENO, " &_
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR,  " &_
								" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB,  " &_
								" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB, " &_
								" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB,  " &_
								" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB, " &_
								" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB " &_
								" FROM " &_
								" ( " &_
								" SELECT " &_  
								" MAX(HDR.FCBA)FCBA, " &_
								" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
								" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR,  " &_
								" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB,  " &_
								" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION_HO) JOBCODE_DB," &_
								" 'FA' SOURCE_DB,  " &_
								" 'FA' SOURCETYPE_DB,  " &_
								" MAX(HDR.FCCODE) SOURCECODE_DB,  " &_
								" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB, " &_  
								" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_DB,  " &_
								" MAX(HDR.DOC_NO) DOCUMENTNO_DB," &_
								" 0 QUANTITY_DB, " &_
								" 0 OUTPUT_DB, " &_
								" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_DB, " &_
								" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_
								" MAX(HDR.FCCODE) AUTO_REFERENCE_01_DB, " &_
								" 'Auto - Fixed Asset Depreciation Ledger No :' TITLE_REFERENCE_02_DB, " &_
								" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_DB, " &_
								" '' TITLE_REFERENCE_03_DB, " &_
								" '' AUTO_REFERENCE_03_DB " &_
								" FROM " &_
								" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
								" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
								" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
								" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
								" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION_HO = IPLASPROD.JOB.FCCODE " &_  
								" INNER JOIN IPLASPROD.GCMASTER ON HDR.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY " &_
								" AND IPLASPROD.GCMASTER.FCCODE = '00999' " &_
								" WHERE " &_
								" IPLASPROD.ASSETGROUP.ISMATURITY = 'FALSE'" &_
								" AND TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
								" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _ 
								" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
								" GROUP BY HDR.FCBA, HDR.FCCODE)"
		Else
			qFADepreciationDB =	" SELECT" &_
								" FCBA, ROWNUM LINENO, " &_
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR,  " &_
								" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB,  " &_
								" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB, " &_
								" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB,  " &_
								" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB, " &_
								" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB " &_
								" FROM " &_
								" ( " &_
								" SELECT " &_  
								" MAX(HDR.FCBA)FCBA, " &_
								" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
								" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR,  " &_
								" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB,  " &_
								" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB," &_
								" 'FA' SOURCE_DB,  " &_
								" 'FA' SOURCETYPE_DB,  " &_
								" MAX(HDR.FCCODE) SOURCECODE_DB,  " &_
								" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB, " &_  
								" MAX(IPLASPROD.GCMASTER.FCCODE) TARGETCODE_DB,  " &_
								" MAX(HDR.DOC_NO) DOCUMENTNO_DB," &_
								" 0 QUANTITY_DB, " &_
								" 0 OUTPUT_DB, " &_
								" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_DB, " &_
								" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_
								" MAX(HDR.FCCODE) AUTO_REFERENCE_01_DB, " &_
								" 'Auto - Fixed Asset Depreciation Ledger No :' TITLE_REFERENCE_02_DB, " &_
								" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_DB, " &_
								" '' TITLE_REFERENCE_03_DB, " &_
								" '' AUTO_REFERENCE_03_DB " &_
								" FROM " &_
								" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
								" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
								" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
								" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
								" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION = IPLASPROD.JOB.FCCODE " &_  
								" INNER JOIN IPLASPROD.GCMASTER ON HDR.FCBA = IPLASPROD.GCMASTER.FCBA AND IPLASPROD.GCMASTER.ITEM = IPLASPROD.JOB.JOB_CATEGORY " &_
								" AND IPLASPROD.GCMASTER.FCCODE = '00999' " &_
								" WHERE " &_
								" IPLASPROD.ASSETGROUP.ISMATURITY = 'FALSE'" &_
								" AND TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
								" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _ 
								" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
								" GROUP BY HDR.FCBA, HDR.FCCODE)"
		End if
		
	Set rsFADepreciationDB = Server.CreateObject("ADODB.Recordset")
	rsFADepreciationDB.Open qFADepreciationDB, connstr
	if not rsFADepreciationDB.eof then
		Do While Not rsFADepreciationDB.eof
			
			 maxLineNo = maxLineNo + 1
			 qCost_ActivityDDB1="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsFADepreciationDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFADepreciationDB("JOBCODE_DB") & "','" & rsFADepreciationDB("SOURCE_DB") & "','" & rsFADepreciationDB("SOURCETYPE_DB") & "','" & rsFADepreciationDB("SOURCECODE_DB") & "'," & _
							 " '','" & rsFADepreciationDB("TARGETTYPE_DB") & "','" & rsFADepreciationDB("TARGETCODE_DB") & "'," & _
							 " '" &  rsFADepreciationDB("AMOUNT_DB") & "','" & rsFADepreciationDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFADepreciationDB("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsFADepreciationDB("TITLE_REFERENCE_01_DB") & "','"  & rsFADepreciationDB("AUTO_REFERENCE_01_DB") & "'," & _
							 " '" & rsFADepreciationDB("TITLE_REFERENCE_02_DB") & "','"  & rsFADepreciationDB("AUTO_REFERENCE_02_DB") & "'," & _
							 " '" & rsFADepreciationDB("TITLE_REFERENCE_03_DB") & "','"  & rsFADepreciationDB("AUTO_REFERENCE_03_DB") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsFADepreciationDB("LINENO") & "','Db', " & _
							 " " & rsFADepreciationDB("ACCOUNTINGPERIOD") & "," & rsFADepreciationDB("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "

			 connstr.execute(qCost_ActivityDDB1)
			 connStr.execute("Commit")
			 
			 
			rsFADepreciationDB.MoveNext
		loop
	end if
	'-- End Insert Debet

	'-- Insert Credit
	qFADepreciationCR =	" SELECT" &_
						" FCBA, ROWNUM LINENO, " &_
						" ACCOUNTINGPERIOD, ACCOUNTINGYEAR,  " &_
						" TDATE_CR, JOBCODE_CR, SOURCE_CR, SOURCETYPE_CR, SOURCECODE_CR, TARGETTYPE_CR, TARGETCODE_CR,  " &_
						" DOCUMENTNO_CR, QUANTITY_CR, OUTPUT_CR, AMOUNT_CR, " &_
						" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR,  " &_
						" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR, " &_
						" TITLE_REFERENCE_03_CR, AUTO_REFERENCE_03_CR " &_
						" FROM " &_
						" ( " &_
						" SELECT " &_  
						" MAX(HDR.FCBA)FCBA, " &_
						" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
						" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR,  " &_
						" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR,  " &_
						" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_ACCD_DEPRECIATION) JOBCODE_CR," &_
						" 'FA' SOURCE_CR,  " &_
						" 'FA' SOURCETYPE_CR,  " &_
						" MAX(HDR.FCCODE) SOURCECODE_CR,  " &_
						" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_CR, " &_  
						" MAX(IPLASPROD.FA.FCCODE) TARGETCODE_CR,  " &_
						" MAX(HDR.DOC_NO) DOCUMENTNO_CR," &_
						" 0 QUANTITY_CR, " &_
						" 0 OUTPUT_CR, " &_
						" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_CR, " &_
						" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR, " &_
						" MAX(HDR.FCCODE) AUTO_REFERENCE_01_CR, " &_
						" 'Auto - Fixed Asset Depreciation Ledger No :' TITLE_REFERENCE_02_CR, " &_
						" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_CR, " &_
						" '' TITLE_REFERENCE_03_CR, " &_
						" '' AUTO_REFERENCE_03_CR " &_
						" FROM " &_
						" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
						" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
						" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
						" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
						" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.ASSETGROUP.CONTROLJOB_ACCD_DEPRECIATION = IPLASPROD.JOB.FCCODE" &_
						" WHERE " &_
						" IPLASPROD.ASSETGROUP.ISMATURITY = 'FALSE'" &_
						" AND TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
						" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _ 
						" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
						" GROUP BY HDR.FCBA, HDR.FCCODE)"
						
			Set rsFADepreciationCR = Server.CreateObject("ADODB.Recordset")
			rsFADepreciationCR.Open qFADepreciationCR, connstr
			if not rsFADepreciationCR.eof then
				Do While Not rsFADepreciationCR.eof

			 maxLineNo = maxLineNo + 1
			 			 
			 qCost_ActivityCR = "Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsFADepreciationCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFADepreciationCR("JOBCODE_CR") & "','" & rsFADepreciationCR("SOURCE_CR") & "','" & rsFADepreciationCR("SOURCETYPE_CR") & "','" & rsFADepreciationCR("SOURCECODE_CR") & "'," & _
							 " '','" & rsFADepreciationCR("TARGETTYPE_CR") & "','" & rsFADepreciationCR("TARGETCODE_CR") & "'," & _
							 " '" &  -1*Cdbl(rsFADepreciationCR("AMOUNT_CR")) & "','" & rsFADepreciationCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFADepreciationCR("FCBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & rsFADepreciationCR("TITLE_REFERENCE_01_CR") & "','"  & rsFADepreciationCR("AUTO_REFERENCE_01_CR") & "'," & _
							 " '" & rsFADepreciationCR("TITLE_REFERENCE_02_CR") & "','"  & rsFADepreciationCR("AUTO_REFERENCE_02_CR") & "'," & _
							 " '" & rsFADepreciationCR("TITLE_REFERENCE_03_CR") & "','"  & rsFADepreciationCR("AUTO_REFERENCE_03_CR") & "'," & _
							 " '','', " & _
							 " 'Line No:', '" & rsFADepreciationCR("LINENO") & "','Cr', " & _
							 " " & rsFADepreciationCR("ACCOUNTINGPERIOD") & "," & rsFADepreciationCR("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
			
			 connstr.execute(qCost_ActivityCR)
			 connStr.execute("Commit")
			 
			 rsFADepreciationCR.MoveNext
		loop
	end if
	'-- End Insert Credit
	
		response.Write "Process Fixed Asset Depreciation Complete....<br>"
		response.Flush
			
	'-- End Of Depresiasi
	
	
	'-- Begin Of FA ALL, PDE, GC Allocated By Single 
	'-- Begin Of FA Amortisasi ALL
	'-- Insert Debet
		qFAAmorIntiAllDB =	"SELECT " &_
							" FCBA, ROWNUM LINENO, " &_
							" ACCOUNTINGPERIOD, ACCOUNTINGYEAR," &_
							" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB," &_
							" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB, " &_ 
							" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB," &_
							" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB," &_
							" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB " &_
							" FROM " &_ 
							" (" &_
							" SELECT " &_
							" MAX(HDR.FCBA)FCBA, " &_
							" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
							" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR, " &_
							" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB, " &_ 
							" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB, " &_ 
							" 'FA' SOURCE_DB," &_
							" 'FA' SOURCETYPE_DB," &_
							" MAX(HDR.FCCODE) SOURCECODE_DB," &_
							" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB, " &_ 
							" MAX(FAM.FIELDCODE) TARGETCODE_DB, " &_   
							" MAX(HDR.DOC_NO) DOCUMENTNO_DB, " &_ 
							" 0 QUANTITY_DB, " &_ 
							" 0 OUTPUT_DB, " &_ 
							" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_DB," &_  
							" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_ 
							" MAX(HDR.FCCODE) AUTO_REFERENCE_01_DB," &_  
							" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_DB," &_   
							" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_DB, " &_ 
							" '' TITLE_REFERENCE_03_DB, " &_
							" '' AUTO_REFERENCE_03_DB" &_  
							" FROM " &_ 
							" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
							" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
							" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_ 
							" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
							" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION = IPLASPROD.JOB.FCCODE " &_
							" INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM ON HDR.FCBA = FAM.FCBA AND HDR.FCCODE = FAM.FIXEDASSETCODE AND FAM.FAMATURETYPE = 'ALL' " &_		
							" WHERE " &_
							" TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
							" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
							" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
							" GROUP BY HDR.FCBA, HDR.FCCODE, HDR.DOC_NO, HDR.LINENO)"

		Set rsFAAmorIntiAllDB = Server.CreateObject("ADODB.Recordset")
		rsFAAmorIntiAllDB.Open qFAAmorIntiAllDB, connstr
		if not rsFAAmorIntiAllDB.eof then
			Do While Not rsFAAmorIntiAllDB.eof
				
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiAllDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiAllDB("JOBCODE_DB") & "','" & rsFAAmorIntiAllDB("SOURCE_DB") & "','" & rsFAAmorIntiAllDB("SOURCETYPE_DB") & "','" & rsFAAmorIntiAllDB("SOURCECODE_DB") & "'," & _
								 " '','" & rsFAAmorIntiAllDB("TARGETTYPE_DB") & "','" & rsFAAmorIntiAllDB("TARGETCODE_DB") & "'," & _
								 " '" &  rsFAAmorIntiAllDB("AMOUNT_DB") & "','" & rsFAAmorIntiAllDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAAmorIntiAllDB("FCBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsFAAmorIntiAllDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAAmorIntiAllDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsFAAmorIntiAllDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAAmorIntiAllDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '" & rsFAAmorIntiAllDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAAmorIntiAllDB("AUTO_REFERENCE_03_DB") & "'," & _
								 " '','', " & _
								 " 'Line No:', '" & rsFAAmorIntiAllDB("LINENO") & "','Db', " & _
								 " " & rsFAAmorIntiAllDB("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiAllDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAAmorIntiAllDB.MoveNext
			loop
		end if
		'-- End Insert Debet
		
		'-- End Insert Credit
			qFAAmorIntiALLCR =	"SELECT " &_
								" FCBA, ROWNUM LINENO, " &_
								" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
								" TDATE_CR, JOBCODE_CR, SOURCE_CR, SOURCETYPE_CR, SOURCECODE_CR, TARGETTYPE_CR, TARGETCODE_CR, " &_
								" DOCUMENTNO_CR, QUANTITY_CR, OUTPUT_CR, AMOUNT_CR, " &_
								" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR, " &_
								" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR, " &_
								" TITLE_REFERENCE_03_CR, AUTO_REFERENCE_03_CR " &_
								" FROM " &_
								" ( " &_
								" SELECT " &_
								" MAX(HDR.FCBA)FCBA, " &_
								" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
								" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR, " &_ 
								" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR, " &_
								" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_ACCD_DEPRECIATION) JOBCODE_CR, " &_
								" 'FA' SOURCE_CR, " &_
								" 'FA' SOURCETYPE_CR, " &_
								" MAX(HDR.FCCODE) SOURCECODE_CR, " &_
								" 'FA' TARGETTYPE_CR, " &_
								" MAX(HDR.FCCODE) TARGETCODE_CR, " &_
								" MAX(HDR.DOC_NO) DOCUMENTNO_CR, " &_
								" 0 QUANTITY_CR, " &_
								" 0 OUTPUT_CR, " &_
								" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_CR, " &_
								" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR, " &_
								" MAX(HDR.FCCODE) AUTO_REFERENCE_01_CR, " &_
								" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_CR, " &_
								" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_CR, " &_
								" '' TITLE_REFERENCE_03_CR, " &_
								" '' AUTO_REFERENCE_03_CR " &_
								" FROM" &_
								" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
								" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
								" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
								" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET'" &_
								" INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM ON HDR.FCBA = FAM.FCBA AND HDR.FCCODE = FAM.FIXEDASSETCODE AND FAM.FAMATURETYPE = 'ALL' " &_
								" WHERE " &_
								" TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
								" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
								" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
								" GROUP BY HDR.FCBA, HDR.FCCODE, HDR.DOC_NO, HDR.LINENO)"

			Set rsFAAmorIntiALLCR = Server.CreateObject("ADODB.Recordset")
			rsFAAmorIntiALLCR.Open qFAAmorIntiALLCR, connstr
			if not rsFAAmorIntiALLCR.eof then
				Do While Not rsFAAmorIntiALLCR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiALLCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiALLCR("JOBCODE_CR") & "','" & rsFAAmorIntiALLCR("SOURCE_CR") & "','" & rsFAAmorIntiALLCR("SOURCETYPE_CR") & "','" & rsFAAmorIntiALLCR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAAmorIntiALLCR("TARGETTYPE_CR") & "','" & rsFAAmorIntiALLCR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAAmorIntiALLCR("AMOUNT_CR")) & "','" & rsFAAmorIntiALLCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAAmorIntiALLCR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAAmorIntiALLCR("TITLE_REFERENCE_01_CR") & "','"  & rsFAAmorIntiALLCR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAAmorIntiALLCR("TITLE_REFERENCE_02_CR") & "','"  & rsFAAmorIntiALLCR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAAmorIntiALLCR("TITLE_REFERENCE_03_CR") & "','"  & rsFAAmorIntiALLCR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAAmorIntiALLCR("LINENO") & "','Cr', " & _
									 " " & rsFAAmorIntiALLCR("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiALLCR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAAmorIntiALLCR.MoveNext
				loop
			end if
		'-- End Insert Credit
			response.Write "Process Fixed Asset ALL Amortisasi Single Blok Complete....<br>"
			response.Flush
	'-- End Of FA Amortisasi ALL
	
	'-- Begin Of FA Amortisasi PDE
	'-- Insert Debet
		qFAAmorIntiPDEDB =	"SELECT " &_
							" FCBA, ROWNUM LINENO, " &_
							" ACCOUNTINGPERIOD, ACCOUNTINGYEAR," &_
							" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB," &_
							" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB, " &_ 
							" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB," &_
							" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB," &_
							" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB " &_
							" FROM " &_ 
							" (" &_
							" SELECT " &_
							" MAX(HDR.FCBA)FCBA, " &_
							" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
							" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR, " &_
							" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB, " &_ 
							" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB, " &_ 
							" 'FA' SOURCE_DB," &_
							" 'FA' SOURCETYPE_DB," &_
							" MAX(HDR.FCCODE) SOURCECODE_DB," &_
							" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB, " &_ 
							" MAX(FAM.FIELDCODE) TARGETCODE_DB, " &_   
							" MAX(HDR.DOC_NO) DOCUMENTNO_DB, " &_ 
							" 0 QUANTITY_DB, " &_ 
							" 0 OUTPUT_DB, " &_ 
							" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_DB," &_  
							" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_ 
							" MAX(HDR.FCCODE) AUTO_REFERENCE_01_DB," &_  
							" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_DB," &_   
							" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_DB, " &_ 
							" '' TITLE_REFERENCE_03_DB, " &_
							" '' AUTO_REFERENCE_03_DB" &_  
							" FROM " &_ 
							" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
							" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
							" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_ 
							" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
							" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION = IPLASPROD.JOB.FCCODE " &_
							" INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM ON HDR.FCBA = FAM.FCBA AND HDR.FCCODE = FAM.FIXEDASSETCODE AND FAM.FAMATURETYPE = 'PDE' " &_		
							" WHERE " &_
							" TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
							" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
							" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
							" GROUP BY HDR.FCBA, HDR.FCCODE, HDR.DOC_NO, HDR.LINENO)"

		Set rsFAAmorIntiPDEDB = Server.CreateObject("ADODB.Recordset")
		rsFAAmorIntiPDEDB.Open qFAAmorIntiPDEDB, connstr
		if not rsFAAmorIntiPDEDB.eof then
			Do While Not rsFAAmorIntiPDEDB.eof
				
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiPDEDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiPDEDB("JOBCODE_DB") & "','" & rsFAAmorIntiPDEDB("SOURCE_DB") & "','" & rsFAAmorIntiPDEDB("SOURCETYPE_DB") & "','" & rsFAAmorIntiPDEDB("SOURCECODE_DB") & "'," & _
								 " '','" & rsFAAmorIntiPDEDB("TARGETTYPE_DB") & "','" & rsFAAmorIntiPDEDB("TARGETCODE_DB") & "'," & _
								 " '" &  rsFAAmorIntiPDEDB("AMOUNT_DB") & "','" & rsFAAmorIntiPDEDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAAmorIntiPDEDB("FCBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsFAAmorIntiPDEDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAAmorIntiPDEDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsFAAmorIntiPDEDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAAmorIntiPDEDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '" & rsFAAmorIntiPDEDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAAmorIntiPDEDB("AUTO_REFERENCE_03_DB") & "'," & _
								 " '','', " & _
								 " 'Line No:', '" & rsFAAmorIntiPDEDB("LINENO") & "','Db', " & _
								 " " & rsFAAmorIntiPDEDB("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiPDEDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAAmorIntiPDEDB.MoveNext
			loop
		end if
		'-- End Insert Debet
		
		'-- End Insert Credit
			qFAAmorIntiPDECR ="SELECT " &_
				" FCBA, ROWNUM LINENO, " &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
				" TDATE_CR, JOBCODE_CR, SOURCE_CR, SOURCETYPE_CR, SOURCECODE_CR, TARGETTYPE_CR, TARGETCODE_CR, " &_
				" DOCUMENTNO_CR, QUANTITY_CR, OUTPUT_CR, AMOUNT_CR, " &_
				" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR, " &_
				" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR, " &_
				" TITLE_REFERENCE_03_CR, AUTO_REFERENCE_03_CR " &_
				" FROM " &_
				" ( " &_
				" SELECT " &_
				" MAX(HDR.FCBA)FCBA, " &_
				" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
				" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR, " &_ 
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR, " &_
				" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_ACCD_DEPRECIATION) JOBCODE_CR, " &_
				" 'FA' SOURCE_CR, " &_
				" 'FA' SOURCETYPE_CR, " &_
				" MAX(HDR.FCCODE) SOURCECODE_CR, " &_
				" 'FA' TARGETTYPE_CR, " &_
				" MAX(HDR.FCCODE) TARGETCODE_CR, " &_
				" MAX(HDR.DOC_NO) DOCUMENTNO_CR, " &_
				" 0 QUANTITY_CR, " &_
				" 0 OUTPUT_CR, " &_
				" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_CR, " &_
				" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR, " &_
				" MAX(HDR.FCCODE) AUTO_REFERENCE_01_CR, " &_
				" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_CR, " &_
				" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_CR, " &_
				" '' TITLE_REFERENCE_03_CR, " &_
				" '' AUTO_REFERENCE_03_CR " &_
				" FROM" &_
				" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
				" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
				" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET'" &_
				" INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM ON HDR.FCBA = FAM.FCBA AND HDR.FCCODE = FAM.FIXEDASSETCODE AND FAM.FAMATURETYPE = 'PDE' " &_
				" WHERE " &_
				" TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
				" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
				" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
				" GROUP BY HDR.FCBA, HDR.FCCODE, HDR.DOC_NO, HDR.LINENO)"

			Set rsFAAmorIntiPDECR = Server.CreateObject("ADODB.Recordset")
			rsFAAmorIntiPDECR.Open qFAAmorIntiPDECR, connstr
			if not rsFAAmorIntiPDECR.eof then
				Do While Not rsFAAmorIntiPDECR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiPDECR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiPDECR("JOBCODE_CR") & "','" & rsFAAmorIntiPDECR("SOURCE_CR") & "','" & rsFAAmorIntiPDECR("SOURCETYPE_CR") & "','" & rsFAAmorIntiPDECR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAAmorIntiPDECR("TARGETTYPE_CR") & "','" & rsFAAmorIntiPDECR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAAmorIntiPDECR("AMOUNT_CR")) & "','" & rsFAAmorIntiPDECR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAAmorIntiPDECR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAAmorIntiPDECR("TITLE_REFERENCE_01_CR") & "','"  & rsFAAmorIntiPDECR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAAmorIntiPDECR("TITLE_REFERENCE_02_CR") & "','"  & rsFAAmorIntiPDECR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAAmorIntiPDECR("TITLE_REFERENCE_03_CR") & "','"  & rsFAAmorIntiPDECR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAAmorIntiPDECR("LINENO") & "','Cr', " & _
									 " " & rsFAAmorIntiPDECR("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiPDECR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAAmorIntiPDECR.MoveNext
				loop
			end if
		'-- End Insert Credit
			response.Write "Process Fixed Asset Amortisasi PDE Single Blok Complete....<br>"
			response.Flush
	'-- End Of FA Amortisasi PDE

	'-- Begin Of FA Amortisasi GC
	'-- Begin Insert Debet
		qFAAmorIntiGCDB ="SELECT " &_
				" FCBA, ROWNUM LINENO, " &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR," &_
				" TDATE_DB, JOBCODE_DB, SOURCE_DB, SOURCETYPE_DB, SOURCECODE_DB, TARGETTYPE_DB, TARGETCODE_DB," &_
				" DOCUMENTNO_DB, QUANTITY_DB, OUTPUT_DB, AMOUNT_DB, " &_ 
				" TITLE_REFERENCE_01_DB, AUTO_REFERENCE_01_DB," &_
				" TITLE_REFERENCE_02_DB, AUTO_REFERENCE_02_DB," &_
				" TITLE_REFERENCE_03_DB, AUTO_REFERENCE_03_DB " &_
				" FROM " &_ 
				" (" &_
				" SELECT " &_
				" MAX(HDR.FCBA)FCBA, " &_
				" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
				" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR, " &_
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB, " &_ 
				" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB, " &_ 
				" 'FA' SOURCE_DB," &_
				" 'FA' SOURCETYPE_DB," &_
				" MAX(HDR.FCCODE) SOURCECODE_DB," &_
				" MAX(IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB, " &_ 
				" MAX(FAM.FIELDCODE) TARGETCODE_DB, " &_   
				" MAX(HDR.DOC_NO) DOCUMENTNO_DB, " &_ 
				" 0 QUANTITY_DB, " &_ 
				" 0 OUTPUT_DB, " &_ 
				" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_DB," &_  
				" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB, " &_ 
				" MAX(HDR.FCCODE) AUTO_REFERENCE_01_DB," &_  
				" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_DB," &_   
				" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_DB, " &_ 
				" '' TITLE_REFERENCE_03_DB, " &_
				" '' AUTO_REFERENCE_03_DB" &_  
				" FROM " &_ 
				" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
				" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
				" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_ 
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET' " &_
				" INNER JOIN IPLASPROD.JOB ON HDR.FCBA = IPLASPROD.JOB.FCBA AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION = IPLASPROD.JOB.FCCODE " &_
				" INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM ON HDR.FCBA = FAM.FCBA AND HDR.FCCODE = FAM.FIXEDASSETCODE AND FAM.FAMATURETYPE = 'GC' " &_
				" WHERE " &_
				" TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
				" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
				" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
				" GROUP BY HDR.FCBA, HDR.FCCODE, HDR.DOC_NO, HDR.LINENO)"
		
		Set rsFAAmorIntiGCDB = Server.CreateObject("ADODB.Recordset")
		rsFAAmorIntiGCDB.Open qFAAmorIntiGCDB, connstr
		if not rsFAAmorIntiGCDB.eof then
			Do While Not rsFAAmorIntiGCDB.eof
								
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiGCDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiGCDB("JOBCODE_DB") & "','" & rsFAAmorIntiGCDB("SOURCE_DB") & "','" & rsFAAmorIntiGCDB("SOURCETYPE_DB") & "','" & rsFAAmorIntiGCDB("SOURCECODE_DB") & "'," & _
								 " '','" & rsFAAmorIntiGCDB("TARGETTYPE_DB") & "','" & rsFAAmorIntiGCDB("TARGETCODE_DB") & "'," & _
								 " '" &  rsFAAmorIntiGCDB("AMOUNT_DB") & "','" & rsFAAmorIntiGCDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAAmorIntiGCDB("FCBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsFAAmorIntiGCDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAAmorIntiGCDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsFAAmorIntiGCDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAAmorIntiGCDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '" & rsFAAmorIntiGCDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAAmorIntiGCDB("AUTO_REFERENCE_03_DB") & "'," & _
								 " '','', " & _
								 " 'Line No:', '" & rsFAAmorIntiGCDB("LINENO") & "','Db', " & _
								 " " & rsFAAmorIntiGCDB("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiGCDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAAmorIntiGCDB.MoveNext
			loop
		end if
	'-- End Insert Debet
	
	'-- Begin Insert Credit
	qFAAmorIntiGCCR ="SELECT " &_
				" FCBA, ROWNUM LINENO, " &_
				" ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " &_
				" TDATE_CR, JOBCODE_CR, SOURCE_CR, SOURCETYPE_CR, SOURCECODE_CR, TARGETTYPE_CR, TARGETCODE_CR, " &_
				" DOCUMENTNO_CR, QUANTITY_CR, OUTPUT_CR, AMOUNT_CR, " &_
				" TITLE_REFERENCE_01_CR, AUTO_REFERENCE_01_CR, " &_
				" TITLE_REFERENCE_02_CR, AUTO_REFERENCE_02_CR, " &_
				" TITLE_REFERENCE_03_CR, AUTO_REFERENCE_03_CR " &_
				" FROM " &_
				" ( " &_
				" SELECT " &_
				" MAX(HDR.FCBA)FCBA, " &_
				" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'fmMM')) ACCOUNTINGPERIOD, " &_
				" MAX(TO_CHAR(HDR.DATE_TRANSACTION,'yyyy')) ACCOUNTINGYEAR, " &_ 
				" MAX(IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_CR, " &_
				" MAX(IPLASPROD.ASSETGROUP.CONTROLJOB_ACCD_DEPRECIATION) JOBCODE_CR, " &_
				" 'FA' SOURCE_CR, " &_
				" 'FA' SOURCETYPE_CR, " &_
				" MAX(HDR.FCCODE) SOURCECODE_CR, " &_
				" 'FA' TARGETTYPE_CR, " &_
				" MAX(HDR.FCCODE) TARGETCODE_CR, " &_
				" MAX(HDR.DOC_NO) DOCUMENTNO_CR, " &_
				" 0 QUANTITY_CR, " &_
				" 0 OUTPUT_CR, " &_
				" MAX(HDR.MONTHLYDEPRECIATION) AMOUNT_CR, " &_
				" 'TC - Fixed Asset Code :' TITLE_REFERENCE_01_CR, " &_
				" MAX(HDR.FCCODE) AUTO_REFERENCE_01_CR, " &_
				" 'Auto - Fixed Asset Maturity Statement No :' TITLE_REFERENCE_02_CR, " &_
				" MAX(HDR.DOC_NO) AUTO_REFERENCE_02_CR, " &_
				" '' TITLE_REFERENCE_03_CR, " &_
				" '' AUTO_REFERENCE_03_CR " &_
				" FROM" &_
				" IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR " &_
				" INNER JOIN IPLASPROD.FIXEDASSET FA ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE " &_
				" INNER JOIN IPLASPROD.ASSETGROUP ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE " &_
				" INNER JOIN IPLASPROD.PERIODCONTROL ON HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA AND UPPER(IPLASPROD.PERIODCONTROL.SYSTEM)='ASSET'" &_
				" INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM ON HDR.FCBA = FAM.FCBA AND HDR.FCCODE = FAM.FIXEDASSETCODE AND FAM.FAMATURETYPE = 'GC'	" &_
				" WHERE " &_
				" TO_CHAR(HDR.DATE_TRANSACTION,'fmMM') =  " & Mid(Session("CurP_Asset"),5,2) &_
				" AND TO_CHAR(HDR.DATE_TRANSACTION,'yyyy') = " & Mid(Session("CurP_Asset"),1,4) & _
				" AND HDR.FCBA = '" & Session("fcBA") & "'" &_
				" GROUP BY HDR.FCBA, HDR.FCCODE, HDR.DOC_NO, HDR.LINENO)"
	
			Set rsFAAmorIntiGCCR = Server.CreateObject("ADODB.Recordset")
			rsFAAmorIntiGCCR.Open qFAAmorIntiGCCR, connstr
			if not rsFAAmorIntiGCCR.eof then
				Do While Not rsFAAmorIntiGCCR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiGCCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiGCCR("JOBCODE_CR") & "','" & rsFAAmorIntiGCCR("SOURCE_CR") & "','" & rsFAAmorIntiGCCR("SOURCETYPE_CR") & "','" & rsFAAmorIntiGCCR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAAmorIntiGCCR("TARGETTYPE_CR") & "','" & rsFAAmorIntiGCCR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAAmorIntiGCCR("AMOUNT_CR")) & "','" & rsFAAmorIntiGCCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAAmorIntiGCCR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAAmorIntiGCCR("TITLE_REFERENCE_01_CR") & "','"  & rsFAAmorIntiGCCR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAAmorIntiGCCR("TITLE_REFERENCE_02_CR") & "','"  & rsFAAmorIntiGCCR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAAmorIntiGCCR("TITLE_REFERENCE_03_CR") & "','"  & rsFAAmorIntiGCCR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAAmorIntiGCCR("LINENO") & "','Cr', " & _
									 " " & rsFAAmorIntiGCCR("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiGCCR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAAmorIntiGCCR.MoveNext
				loop
			end if
	'-- End Insert Credit
	
			response.Write "Process Fixed Asset Amortisasi GC Single Blok Complete....<br>"
			response.Flush
	'-- End Of FA Amortisasi GC
	'-- End Begin Of FA ALL, PDE, GC Allocated By Single 
	
	'-- Begin Of FA ALL, PDE, GC Allocated By Multiple 
	'-- Begin Of FA Amortisasi ALL
	'-- Insert Debet
		qFAAmorIntiAllDB =	" SELECT FCBA," & _
							"        ROWNUM LINENO," & _
							"        ACCOUNTINGPERIOD," & _
							"        ACCOUNTINGYEAR," & _
							"        TDATE_DB," & _
							"        JOBCODE_DB," & _
							"        SOURCE_DB," & _
							"        SOURCETYPE_DB," & _
							"        SOURCECODE_DB," & _
							"        TARGETTYPE_DB," & _
							"        TARGETCODE_DB," & _
							"        DOCUMENTNO_DB," & _
							"        QUANTITY_DB," & _
							"        OUTPUT_DB," & _
							"        AMOUNT_DB," & _
							"        TITLE_REFERENCE_01_DB," & _
							"        AUTO_REFERENCE_01_DB," & _
							"        TITLE_REFERENCE_02_DB," & _
							"        AUTO_REFERENCE_02_DB," & _
							"        TITLE_REFERENCE_03_DB," & _
							"        AUTO_REFERENCE_03_DB" & _
							"   FROM (  SELECT MAX (HDR.FCBA) FCBA," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM')) ACCOUNTINGPERIOD," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy')) ACCOUNTINGYEAR," & _
							"                  MAX (IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB," & _
							"                  MAX (IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB," & _
							"                  'FA' SOURCE_DB," & _
							"                  'FA' SOURCETYPE_DB," & _
							"                  MAX (HDR.FCCODE) SOURCECODE_DB," & _
							"                  MAX (IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB," & _
							"                  MAX (FAM.FIELDCODE) TARGETCODE_DB," & _
							"                  MAX (HDR.DOC_NO) DOCUMENTNO_DB," & _
							"                  0 QUANTITY_DB," & _
							"                  0 OUTPUT_DB," & _
							"                  NVL (SUM (ffha.proporsiha * HDR.MONTHLYDEPRECIATION), 0) AMOUNT_DB," & _
							"                  'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB," & _
							"                  MAX (HDR.FCCODE) AUTO_REFERENCE_01_DB," & _
							"                  'Auto - Fixed Asset Maturity Statement No :'" & _
							"                     TITLE_REFERENCE_02_DB," & _
							"                  MAX (HDR.DOC_NO) AUTO_REFERENCE_02_DB," & _
							"                  '' TITLE_REFERENCE_03_DB," & _
							"                  '' AUTO_REFERENCE_03_DB" & _
							"             FROM IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET FA" & _
							"                     ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE" & _
							"                  INNER JOIN IPLASPROD.ASSETGROUP" & _
							"                     ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE" & _
							"                  INNER JOIN IPLASPROD.PERIODCONTROL" & _
							"                     ON     HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA" & _
							"                        AND UPPER (IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET'" & _
							"                  INNER JOIN IPLASPROD.JOB" & _
							"                     ON     HDR.FCBA = IPLASPROD.JOB.FCBA" & _
							"                        AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION =" & _
							"                               IPLASPROD.JOB.FCCODE" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM" & _
							"                     ON     HDR.FCBA = FAM.FCBA" & _
							"                        AND HDR.FCCODE = FAM.FIXEDASSETCODE" & _
							"                        AND FAM.FAMATURETYPE = 'ALL'" & _
							"                  RIGHT JOIN" & _
							"                  (SELECT ff.fcba," & _
							"                          ff.fccode," & _
							"                          ff.hectarageplanted," & _
							"                          fd.totalhectarageplanted," & _
							"                          (ff.hectarageplanted / fd.totalhectarageplanted)" & _
							"                             proporsiha" & _
							"                     FROM field ff" & _
							"                          INNER JOIN" & _
							"                          (  SELECT fcba," & _
							"                                    SUM (hectarageplanted) totalhectarageplanted" & _
							"                               FROM field" & _
							"                              WHERE fcba = '"& session("fcba") &"' AND DX_BLOCK_STATUS = 'TM'" & _
							"                           GROUP BY fcba) fd" & _
							"                             ON (fd.fcba = ff.fcba)" & _
							"                    WHERE ff.DX_BLOCK_STATUS = 'TM' AND ff.fcba = '"& session("fcba") &"') ffha" & _
							"                     ON (ffha.fcba = fam.fcba)" & _
							"            WHERE     TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM') = "& Mid(Session("CurP_Asset"),5,2) &"" & _
							"                  AND TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy') = "& Mid(Session("CurP_Asset"),1,4) &"" & _
							"                  AND HDR.FCBA = '"& session("fcba") &"'" & _
							"         GROUP BY HDR.FCBA," & _
							"                  HDR.FCCODE," & _
							"                  HDR.DOC_NO," & _
							"                  HDR.LINENO)"  

		Set rsFAAmorIntiAllDB = Server.CreateObject("ADODB.Recordset")
		rsFAAmorIntiAllDB.Open qFAAmorIntiAllDB, connstr
		if not rsFAAmorIntiAllDB.eof then
			Do While Not rsFAAmorIntiAllDB.eof
				
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiAllDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiAllDB("JOBCODE_DB") & "','" & rsFAAmorIntiAllDB("SOURCE_DB") & "','" & rsFAAmorIntiAllDB("SOURCETYPE_DB") & "','" & rsFAAmorIntiAllDB("SOURCECODE_DB") & "'," & _
								 " '','" & rsFAAmorIntiAllDB("TARGETTYPE_DB") & "','" & rsFAAmorIntiAllDB("TARGETCODE_DB") & "'," & _
								 " '" &  rsFAAmorIntiAllDB("AMOUNT_DB") & "','" & rsFAAmorIntiAllDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAAmorIntiAllDB("FCBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsFAAmorIntiAllDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAAmorIntiAllDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsFAAmorIntiAllDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAAmorIntiAllDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '" & rsFAAmorIntiAllDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAAmorIntiAllDB("AUTO_REFERENCE_03_DB") & "'," & _
								 " '','', " & _
								 " 'Line No:', '" & rsFAAmorIntiAllDB("LINENO") & "','Db', " & _
								 " " & rsFAAmorIntiAllDB("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiAllDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAAmorIntiAllDB.MoveNext
			loop
		end if
		'-- End Insert Debet
		
		'-- End Insert Credit
			qFAAmorIntiALLCR =	" SELECT FCBA," & _
							"        ROWNUM LINENO," & _
							"        ACCOUNTINGPERIOD," & _
							"        ACCOUNTINGYEAR," & _
							"        TDATE_DB," & _
							"        JOBCODE_DB," & _
							"        SOURCE_DB," & _
							"        SOURCETYPE_DB," & _
							"        SOURCECODE_DB," & _
							"        TARGETTYPE_DB," & _
							"        TARGETCODE_DB," & _
							"        DOCUMENTNO_DB," & _
							"        QUANTITY_DB," & _
							"        OUTPUT_DB," & _
							"        AMOUNT_DB," & _
							"        TITLE_REFERENCE_01_DB," & _
							"        AUTO_REFERENCE_01_DB," & _
							"        TITLE_REFERENCE_02_DB," & _
							"        AUTO_REFERENCE_02_DB," & _
							"        TITLE_REFERENCE_03_DB," & _
							"        AUTO_REFERENCE_03_DB" & _
							"   FROM (  SELECT MAX (HDR.FCBA) FCBA," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM')) ACCOUNTINGPERIOD," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy')) ACCOUNTINGYEAR," & _
							"                  MAX (IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB," & _
							"                  MAX (IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB," & _
							"                  'FA' SOURCE_DB," & _
							"                  'FA' SOURCETYPE_DB," & _
							"                  MAX (HDR.FCCODE) SOURCECODE_DB," & _
							"                  MAX (IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB," & _
							"                  MAX (FAM.FIELDCODE) TARGETCODE_DB," & _
							"                  MAX (HDR.DOC_NO) DOCUMENTNO_DB," & _
							"                  0 QUANTITY_DB," & _
							"                  0 OUTPUT_DB," & _
							"                  -1 * NVL (SUM (ffha.proporsiha * HDR.MONTHLYDEPRECIATION), 0) AMOUNT_DB," & _
							"                  'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB," & _
							"                  MAX (HDR.FCCODE) AUTO_REFERENCE_01_DB," & _
							"                  'Auto - Fixed Asset Maturity Statement No :'" & _
							"                     TITLE_REFERENCE_02_DB," & _
							"                  MAX (HDR.DOC_NO) AUTO_REFERENCE_02_DB," & _
							"                  '' TITLE_REFERENCE_03_DB," & _
							"                  '' AUTO_REFERENCE_03_DB" & _
							"             FROM IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET FA" & _
							"                     ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE" & _
							"                  INNER JOIN IPLASPROD.ASSETGROUP" & _
							"                     ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE" & _
							"                  INNER JOIN IPLASPROD.PERIODCONTROL" & _
							"                     ON     HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA" & _
							"                        AND UPPER (IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET'" & _
							"                  INNER JOIN IPLASPROD.JOB" & _
							"                     ON     HDR.FCBA = IPLASPROD.JOB.FCBA" & _
							"                        AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION =" & _
							"                               IPLASPROD.JOB.FCCODE" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM" & _
							"                     ON     HDR.FCBA = FAM.FCBA" & _
							"                        AND HDR.FCCODE = FAM.FIXEDASSETCODE" & _
							"                        AND FAM.FAMATURETYPE = 'ALL'" & _
							"                  RIGHT JOIN" & _
							"                  (SELECT ff.fcba," & _
							"                          ff.fccode," & _
							"                          ff.hectarageplanted," & _
							"                          fd.totalhectarageplanted," & _
							"                          (ff.hectarageplanted / fd.totalhectarageplanted)" & _
							"                             proporsiha" & _
							"                     FROM field ff" & _
							"                          INNER JOIN" & _
							"                          (  SELECT fcba," & _
							"                                    SUM (hectarageplanted) totalhectarageplanted" & _
							"                               FROM field" & _
							"                              WHERE fcba = '"& session("fcba") &"' AND DX_BLOCK_STATUS = 'TM'" & _
							"                           GROUP BY fcba) fd" & _
							"                             ON (fd.fcba = ff.fcba)" & _
							"                    WHERE ff.DX_BLOCK_STATUS = 'TM' AND ff.fcba = '"& session("fcba") &"') ffha" & _
							"                     ON (ffha.fcba = fam.fcba)" & _
							"            WHERE     TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM') = "& Mid(Session("CurP_Asset"),5,2) &"" & _
							"                  AND TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy') = "& Mid(Session("CurP_Asset"),1,4) &"" & _
							"                  AND HDR.FCBA = '"& session("fcba") &"'" & _
							"         GROUP BY HDR.FCBA," & _
							"                  HDR.FCCODE," & _
							"                  HDR.DOC_NO," & _
							"                  HDR.LINENO)"  

			Set rsFAAmorIntiALLCR = Server.CreateObject("ADODB.Recordset")
			rsFAAmorIntiALLCR.Open qFAAmorIntiALLCR, connstr
			if not rsFAAmorIntiALLCR.eof then
				Do While Not rsFAAmorIntiALLCR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiALLCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiALLCR("JOBCODE_CR") & "','" & rsFAAmorIntiALLCR("SOURCE_CR") & "','" & rsFAAmorIntiALLCR("SOURCETYPE_CR") & "','" & rsFAAmorIntiALLCR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAAmorIntiALLCR("TARGETTYPE_CR") & "','" & rsFAAmorIntiALLCR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAAmorIntiALLCR("AMOUNT_CR")) & "','" & rsFAAmorIntiALLCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAAmorIntiALLCR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAAmorIntiALLCR("TITLE_REFERENCE_01_CR") & "','"  & rsFAAmorIntiALLCR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAAmorIntiALLCR("TITLE_REFERENCE_02_CR") & "','"  & rsFAAmorIntiALLCR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAAmorIntiALLCR("TITLE_REFERENCE_03_CR") & "','"  & rsFAAmorIntiALLCR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAAmorIntiALLCR("LINENO") & "','Cr', " & _
									 " " & rsFAAmorIntiALLCR("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiALLCR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAAmorIntiALLCR.MoveNext
				loop
			end if
		'-- End Insert Credit
			response.Write "Process Fixed Asset ALL Amortisasi Multiple Blok Complete....<br>"
			response.Flush
	'-- End Of FA Amortisasi ALL
	
	'-- Begin Of FA Amortisasi PDE
	'-- Insert Debet
		qFAAmorIntiPDEDB =	" SELECT FCBA," & _
							"        ROWNUM LINENO," & _
							"        ACCOUNTINGPERIOD," & _
							"        ACCOUNTINGYEAR," & _
							"        TDATE_DB," & _
							"        JOBCODE_DB," & _
							"        SOURCE_DB," & _
							"        SOURCETYPE_DB," & _
							"        SOURCECODE_DB," & _
							"        TARGETTYPE_DB," & _
							"        TARGETCODE_DB," & _
							"        DOCUMENTNO_DB," & _
							"        QUANTITY_DB," & _
							"        OUTPUT_DB," & _
							"        AMOUNT_DB," & _
							"        TITLE_REFERENCE_01_DB," & _
							"        AUTO_REFERENCE_01_DB," & _
							"        TITLE_REFERENCE_02_DB," & _
							"        AUTO_REFERENCE_02_DB," & _
							"        TITLE_REFERENCE_03_DB," & _
							"        AUTO_REFERENCE_03_DB" & _
							"   FROM (  SELECT MAX (HDR.FCBA) FCBA," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM')) ACCOUNTINGPERIOD," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy')) ACCOUNTINGYEAR," & _
							"                  MAX (IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB," & _
							"                  MAX (IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB," & _
							"                  'FA' SOURCE_DB," & _
							"                  'FA' SOURCETYPE_DB," & _
							"                  MAX (HDR.FCCODE) SOURCECODE_DB," & _
							"                  MAX (IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB," & _
							"                  MAX (FAM.FIELDCODE) TARGETCODE_DB," & _
							"                  MAX (HDR.DOC_NO) DOCUMENTNO_DB," & _
							"                  0 QUANTITY_DB," & _
							"                  0 OUTPUT_DB," & _
							"                  NVL (SUM (ffha.proporsiha * HDR.MONTHLYDEPRECIATION), 0) AMOUNT_DB," & _
							"                  'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB," & _
							"                  MAX (HDR.FCCODE) AUTO_REFERENCE_01_DB," & _
							"                  'Auto - Fixed Asset Maturity Statement No :'" & _
							"                     TITLE_REFERENCE_02_DB," & _
							"                  MAX (HDR.DOC_NO) AUTO_REFERENCE_02_DB," & _
							"                  '' TITLE_REFERENCE_03_DB," & _
							"                  '' AUTO_REFERENCE_03_DB" & _
							"             FROM IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET FA" & _
							"                     ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE" & _
							"                  INNER JOIN IPLASPROD.ASSETGROUP" & _
							"                     ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE" & _
							"                  INNER JOIN IPLASPROD.PERIODCONTROL" & _
							"                     ON     HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA" & _
							"                        AND UPPER (IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET'" & _
							"                  INNER JOIN IPLASPROD.JOB" & _
							"                     ON     HDR.FCBA = IPLASPROD.JOB.FCBA" & _
							"                        AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION =" & _
							"                               IPLASPROD.JOB.FCCODE" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM" & _
							"                     ON     HDR.FCBA = FAM.FCBA" & _
							"                        AND HDR.FCCODE = FAM.FIXEDASSETCODE" & _
							"                        AND FAM.FAMATURETYPE = 'PDE'" & _
							"                  RIGHT JOIN" & _
							"                  (SELECT ff.fcba," & _
							"                          ff.fccode," & _
							"                          ff.hectarageplanted," & _
							"                          fd.totalhectarageplanted," & _
							"                          (ff.hectarageplanted / fd.totalhectarageplanted)" & _
							"                             proporsiha" & _
							"                     FROM field ff" & _
							"                          INNER JOIN" & _
							"                          (  SELECT fcba," & _
							"                                    SUM (hectarageplanted) totalhectarageplanted" & _
							"                               FROM field" & _
							"                              WHERE fcba = '"& session("fcba") &"' AND DX_BLOCK_STATUS = 'TM'" & _
							"                           GROUP BY fcba) fd" & _
							"                             ON (fd.fcba = ff.fcba)" & _
							"                    WHERE ff.DX_BLOCK_STATUS = 'TM' AND ff.fcba = '"& session("fcba") &"') ffha" & _
							"                     ON (ffha.fcba = fam.fcba)" & _
							"            WHERE     TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM') = "& Mid(Session("CurP_Asset"),5,2) &"" & _
							"                  AND TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy') = "& Mid(Session("CurP_Asset"),1,4) &"" & _
							"                  AND HDR.FCBA = '"& session("fcba") &"'" & _
							"         GROUP BY HDR.FCBA," & _
							"                  HDR.FCCODE," & _
							"                  HDR.DOC_NO," & _
							"                  HDR.LINENO)"  

		Set rsFAAmorIntiPDEDB = Server.CreateObject("ADODB.Recordset")
		rsFAAmorIntiPDEDB.Open qFAAmorIntiPDEDB, connstr
		if not rsFAAmorIntiPDEDB.eof then
			Do While Not rsFAAmorIntiPDEDB.eof
				
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiPDEDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiPDEDB("JOBCODE_DB") & "','" & rsFAAmorIntiPDEDB("SOURCE_DB") & "','" & rsFAAmorIntiPDEDB("SOURCETYPE_DB") & "','" & rsFAAmorIntiPDEDB("SOURCECODE_DB") & "'," & _
								 " '','" & rsFAAmorIntiPDEDB("TARGETTYPE_DB") & "','" & rsFAAmorIntiPDEDB("TARGETCODE_DB") & "'," & _
								 " '" &  rsFAAmorIntiPDEDB("AMOUNT_DB") & "','" & rsFAAmorIntiPDEDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAAmorIntiPDEDB("FCBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsFAAmorIntiPDEDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAAmorIntiPDEDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsFAAmorIntiPDEDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAAmorIntiPDEDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '" & rsFAAmorIntiPDEDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAAmorIntiPDEDB("AUTO_REFERENCE_03_DB") & "'," & _
								 " '','', " & _
								 " 'Line No:', '" & rsFAAmorIntiPDEDB("LINENO") & "','Db', " & _
								 " " & rsFAAmorIntiPDEDB("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiPDEDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAAmorIntiPDEDB.MoveNext
			loop
		end if
		'-- End Insert Debet
		
		'-- End Insert Credit
			qFAAmorIntiPDECR =	" SELECT FCBA," & _
							"        ROWNUM LINENO," & _
							"        ACCOUNTINGPERIOD," & _
							"        ACCOUNTINGYEAR," & _
							"        TDATE_DB," & _
							"        JOBCODE_DB," & _
							"        SOURCE_DB," & _
							"        SOURCETYPE_DB," & _
							"        SOURCECODE_DB," & _
							"        TARGETTYPE_DB," & _
							"        TARGETCODE_DB," & _
							"        DOCUMENTNO_DB," & _
							"        QUANTITY_DB," & _
							"        OUTPUT_DB," & _
							"        AMOUNT_DB," & _
							"        TITLE_REFERENCE_01_DB," & _
							"        AUTO_REFERENCE_01_DB," & _
							"        TITLE_REFERENCE_02_DB," & _
							"        AUTO_REFERENCE_02_DB," & _
							"        TITLE_REFERENCE_03_DB," & _
							"        AUTO_REFERENCE_03_DB" & _
							"   FROM (  SELECT MAX (HDR.FCBA) FCBA," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM')) ACCOUNTINGPERIOD," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy')) ACCOUNTINGYEAR," & _
							"                  MAX (IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB," & _
							"                  MAX (IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB," & _
							"                  'FA' SOURCE_DB," & _
							"                  'FA' SOURCETYPE_DB," & _
							"                  MAX (HDR.FCCODE) SOURCECODE_DB," & _
							"                  MAX (IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB," & _
							"                  MAX (FAM.FIELDCODE) TARGETCODE_DB," & _
							"                  MAX (HDR.DOC_NO) DOCUMENTNO_DB," & _
							"                  0 QUANTITY_DB," & _
							"                  0 OUTPUT_DB," & _
							"                  -1 * NVL (SUM (ffha.proporsiha * HDR.MONTHLYDEPRECIATION), 0) AMOUNT_DB," & _
							"                  'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB," & _
							"                  MAX (HDR.FCCODE) AUTO_REFERENCE_01_DB," & _
							"                  'Auto - Fixed Asset Maturity Statement No :'" & _
							"                     TITLE_REFERENCE_02_DB," & _
							"                  MAX (HDR.DOC_NO) AUTO_REFERENCE_02_DB," & _
							"                  '' TITLE_REFERENCE_03_DB," & _
							"                  '' AUTO_REFERENCE_03_DB" & _
							"             FROM IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET FA" & _
							"                     ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE" & _
							"                  INNER JOIN IPLASPROD.ASSETGROUP" & _
							"                     ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE" & _
							"                  INNER JOIN IPLASPROD.PERIODCONTROL" & _
							"                     ON     HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA" & _
							"                        AND UPPER (IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET'" & _
							"                  INNER JOIN IPLASPROD.JOB" & _
							"                     ON     HDR.FCBA = IPLASPROD.JOB.FCBA" & _
							"                        AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION =" & _
							"                               IPLASPROD.JOB.FCCODE" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM" & _
							"                     ON     HDR.FCBA = FAM.FCBA" & _
							"                        AND HDR.FCCODE = FAM.FIXEDASSETCODE" & _
							"                        AND FAM.FAMATURETYPE = 'ALL'" & _
							"                  RIGHT JOIN" & _
							"                  (SELECT ff.fcba," & _
							"                          ff.fccode," & _
							"                          ff.hectarageplanted," & _
							"                          fd.totalhectarageplanted," & _
							"                          (ff.hectarageplanted / fd.totalhectarageplanted)" & _
							"                             proporsiha" & _
							"                     FROM field ff" & _
							"                          INNER JOIN" & _
							"                          (  SELECT fcba," & _
							"                                    SUM (hectarageplanted) totalhectarageplanted" & _
							"                               FROM field" & _
							"                              WHERE fcba = '"& session("fcba") &"' AND DX_BLOCK_STATUS = 'TM'" & _
							"                           GROUP BY fcba) fd" & _
							"                             ON (fd.fcba = ff.fcba)" & _
							"                    WHERE ff.DX_BLOCK_STATUS = 'TM' AND ff.fcba = '"& session("fcba") &"') ffha" & _
							"                     ON (ffha.fcba = fam.fcba)" & _
							"            WHERE     TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM') = "& Mid(Session("CurP_Asset"),5,2) &"" & _
							"                  AND TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy') = "& Mid(Session("CurP_Asset"),1,4) &"" & _
							"                  AND HDR.FCBA = '"& session("fcba") &"'" & _
							"         GROUP BY HDR.FCBA," & _
							"                  HDR.FCCODE," & _
							"                  HDR.DOC_NO," & _
							"                  HDR.LINENO)"  

			Set rsFAAmorIntiPDECR = Server.CreateObject("ADODB.Recordset")
			rsFAAmorIntiPDECR.Open qFAAmorIntiPDECR, connstr
			if not rsFAAmorIntiPDECR.eof then
				Do While Not rsFAAmorIntiPDECR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _
									 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiPDECR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiPDECR("JOBCODE_CR") & "','" & rsFAAmorIntiPDECR("SOURCE_CR") & "','" & rsFAAmorIntiPDECR("SOURCETYPE_CR") & "','" & rsFAAmorIntiPDECR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAAmorIntiPDECR("TARGETTYPE_CR") & "','" & rsFAAmorIntiPDECR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAAmorIntiPDECR("AMOUNT_CR")) & "','" & rsFAAmorIntiPDECR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAAmorIntiPDECR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAAmorIntiPDECR("TITLE_REFERENCE_01_CR") & "','"  & rsFAAmorIntiPDECR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAAmorIntiPDECR("TITLE_REFERENCE_02_CR") & "','"  & rsFAAmorIntiPDECR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAAmorIntiPDECR("TITLE_REFERENCE_03_CR") & "','"  & rsFAAmorIntiPDECR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAAmorIntiPDECR("LINENO") & "','Cr', " & _
									 " " & rsFAAmorIntiPDECR("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiPDECR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAAmorIntiPDECR.MoveNext
				loop
			end if
		'-- End Insert Credit
			response.Write "Process Fixed Asset PDE Amortisasi Multiple Blok Complete....<br>"
			response.Flush
	'-- End Begin Of FA Amortisasi PDE
	
	'-- Begin Of FA Amortisasi GC
	'-- Insert Debet
		qFAAmorIntiGCB =	" SELECT FCBA," & _
							"        ROWNUM LINENO," & _
							"        ACCOUNTINGPERIOD," & _
							"        ACCOUNTINGYEAR," & _
							"        TDATE_DB," & _
							"        JOBCODE_DB," & _
							"        SOURCE_DB," & _
							"        SOURCETYPE_DB," & _
							"        SOURCECODE_DB," & _
							"        TARGETTYPE_DB," & _
							"        TARGETCODE_DB," & _
							"        DOCUMENTNO_DB," & _
							"        QUANTITY_DB," & _
							"        OUTPUT_DB," & _
							"        AMOUNT_DB," & _
							"        TITLE_REFERENCE_01_DB," & _
							"        AUTO_REFERENCE_01_DB," & _
							"        TITLE_REFERENCE_02_DB," & _
							"        AUTO_REFERENCE_02_DB," & _
							"        TITLE_REFERENCE_03_DB," & _
							"        AUTO_REFERENCE_03_DB" & _
							"   FROM (  SELECT MAX (HDR.FCBA) FCBA," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM')) ACCOUNTINGPERIOD," & _
							"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy')) ACCOUNTINGYEAR," & _
							"                  MAX (IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB," & _
							"                  MAX (IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB," & _
							"                  'FA' SOURCE_DB," & _
							"                  'FA' SOURCETYPE_DB," & _
							"                  MAX (HDR.FCCODE) SOURCECODE_DB," & _
							"                  MAX (IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB," & _
							"                  MAX (FAM.FIELDCODE) TARGETCODE_DB," & _
							"                  MAX (HDR.DOC_NO) DOCUMENTNO_DB," & _
							"                  0 QUANTITY_DB," & _
							"                  0 OUTPUT_DB," & _
							"                  NVL (SUM (ffha.proporsiha * HDR.MONTHLYDEPRECIATION), 0) AMOUNT_DB," & _
							"                  'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB," & _
							"                  MAX (HDR.FCCODE) AUTO_REFERENCE_01_DB," & _
							"                  'Auto - Fixed Asset Maturity Statement No :'" & _
							"                     TITLE_REFERENCE_02_DB," & _
							"                  MAX (HDR.DOC_NO) AUTO_REFERENCE_02_DB," & _
							"                  '' TITLE_REFERENCE_03_DB," & _
							"                  '' AUTO_REFERENCE_03_DB" & _
							"             FROM IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET FA" & _
							"                     ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE" & _
							"                  INNER JOIN IPLASPROD.ASSETGROUP" & _
							"                     ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE" & _
							"                  INNER JOIN IPLASPROD.PERIODCONTROL" & _
							"                     ON     HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA" & _
							"                        AND UPPER (IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET'" & _
							"                  INNER JOIN IPLASPROD.JOB" & _
							"                     ON     HDR.FCBA = IPLASPROD.JOB.FCBA" & _
							"                        AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION =" & _
							"                               IPLASPROD.JOB.FCCODE" & _
							"                  INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM" & _
							"                     ON     HDR.FCBA = FAM.FCBA" & _
							"                        AND HDR.FCCODE = FAM.FIXEDASSETCODE" & _
							"                        AND FAM.FAMATURETYPE = 'PDE'" & _
							"                  RIGHT JOIN" & _
							"                  (SELECT ff.fcba," & _
							"                          ff.fccode," & _
							"                          ff.hectarageplanted," & _
							"                          fd.totalhectarageplanted," & _
							"                          (ff.hectarageplanted / fd.totalhectarageplanted)" & _
							"                             proporsiha" & _
							"                     FROM field ff" & _
							"                          INNER JOIN" & _
							"                          (  SELECT fcba," & _
							"                                    SUM (hectarageplanted) totalhectarageplanted" & _
							"                               FROM field" & _
							"                              WHERE fcba = '"& session("fcba") &"' AND DX_BLOCK_STATUS = 'TM'" & _
							"                           GROUP BY fcba) fd" & _
							"                             ON (fd.fcba = ff.fcba)" & _
							"                    WHERE ff.DX_BLOCK_STATUS = 'TM' AND ff.fcba = '"& session("fcba") &"') ffha" & _
							"                     ON (ffha.fcba = fam.fcba)" & _
							"            WHERE     TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM') = "& Mid(Session("CurP_Asset"),5,2) &"" & _
							"                  AND TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy') = "& Mid(Session("CurP_Asset"),1,4) &"" & _
							"                  AND HDR.FCBA = '"& session("fcba") &"'" & _
							"         GROUP BY HDR.FCBA," & _
							"                  HDR.FCCODE," & _
							"                  HDR.DOC_NO," & _
							"                  HDR.LINENO)"  

		Set rsFAAmorIntiGCDB = Server.CreateObject("ADODB.Recordset")
		rsFAAmorIntiGCDB.Open qFAAmorIntiGCB, connstr
		if not rsFAAmorIntiGCDB.eof then
			Do While Not rsFAAmorIntiGCDB.eof
				
				 maxLineNo = maxLineNo + 1
				 qCost_ActivityDDB1="Insert into Cost_Activity " & _
								 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
								 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
								 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
								 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
								 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
								 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
								 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
								 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
								 " LASTUPDATE, LASTTIME ) " & _
								 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiGCDB("TDATE_DB") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiGCDB("JOBCODE_DB") & "','" & rsFAAmorIntiGCDB("SOURCE_DB") & "','" & rsFAAmorIntiGCDB("SOURCETYPE_DB") & "','" & rsFAAmorIntiGCDB("SOURCECODE_DB") & "'," & _
								 " '','" & rsFAAmorIntiGCDB("TARGETTYPE_DB") & "','" & rsFAAmorIntiGCDB("TARGETCODE_DB") & "'," & _
								 " '" &  rsFAAmorIntiGCDB("AMOUNT_DB") & "','" & rsFAAmorIntiGCDB("DOCUMENTNO_DB") & "'," & "0,0,'" & rsFAAmorIntiGCDB("FCBA") & "',"  & _
								 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
								 " '" & rsFAAmorIntiGCDB("TITLE_REFERENCE_01_DB") & "','"  & rsFAAmorIntiGCDB("AUTO_REFERENCE_01_DB") & "'," & _
								 " '" & rsFAAmorIntiGCDB("TITLE_REFERENCE_02_DB") & "','"  & rsFAAmorIntiGCDB("AUTO_REFERENCE_02_DB") & "'," & _
								 " '" & rsFAAmorIntiGCDB("TITLE_REFERENCE_03_DB") & "','"  & rsFAAmorIntiGCDB("AUTO_REFERENCE_03_DB") & "'," & _
								 " '','', " & _
								 " 'Line No:', '" & rsFAAmorIntiGCDB("LINENO") & "','Db', " & _
								 " " & rsFAAmorIntiGCDB("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiGCDB("ACCOUNTINGYEAR") & "," & _
								 " sysdate, to_char(sysdate,'HH24:MI')) "

				 connstr.execute(qCost_ActivityDDB1)
				 connStr.execute("Commit")
				
				rsFAAmorIntiGCDB.MoveNext
			loop
		end if
		'-- End Insert Debet
		
		'-- End Insert Credit
			qFAAmorIntiGCCR =	" SELECT FCBA," & _
								"        ROWNUM LINENO," & _
								"        ACCOUNTINGPERIOD," & _
								"        ACCOUNTINGYEAR," & _
								"        TDATE_DB," & _
								"        JOBCODE_DB," & _
								"        SOURCE_DB," & _
								"        SOURCETYPE_DB," & _
								"        SOURCECODE_DB," & _
								"        TARGETTYPE_DB," & _
								"        TARGETCODE_DB," & _
								"        DOCUMENTNO_DB," & _
								"        QUANTITY_DB," & _
								"        OUTPUT_DB," & _
								"        AMOUNT_DB," & _
								"        TITLE_REFERENCE_01_DB," & _
								"        AUTO_REFERENCE_01_DB," & _
								"        TITLE_REFERENCE_02_DB," & _
								"        AUTO_REFERENCE_02_DB," & _
								"        TITLE_REFERENCE_03_DB," & _
								"        AUTO_REFERENCE_03_DB" & _
								"   FROM (  SELECT MAX (HDR.FCBA) FCBA," & _
								"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM')) ACCOUNTINGPERIOD," & _
								"                  MAX (TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy')) ACCOUNTINGYEAR," & _
								"                  MAX (IPLASPROD.PERIODCONTROL.CURRENTDATE) TDATE_DB," & _
								"                  MAX (IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION) JOBCODE_DB," & _
								"                  'FA' SOURCE_DB," & _
								"                  'FA' SOURCETYPE_DB," & _
								"                  MAX (HDR.FCCODE) SOURCECODE_DB," & _
								"                  MAX (IPLASPROD.JOB.JOB_CATEGORY) TARGETTYPE_DB," & _
								"                  MAX (FAM.FIELDCODE) TARGETCODE_DB," & _
								"                  MAX (HDR.DOC_NO) DOCUMENTNO_DB," & _
								"                  0 QUANTITY_DB," & _
								"                  0 OUTPUT_DB," & _
								"                  -1 * NVL (SUM (ffha.proporsiha * HDR.MONTHLYDEPRECIATION), 0) AMOUNT_DB," & _
								"                  'TC - Fixed Asset Code :' TITLE_REFERENCE_01_DB," & _
								"                  MAX (HDR.FCCODE) AUTO_REFERENCE_01_DB," & _
								"                  'Auto - Fixed Asset Maturity Statement No :'" & _
								"                     TITLE_REFERENCE_02_DB," & _
								"                  MAX (HDR.DOC_NO) AUTO_REFERENCE_02_DB," & _
								"                  '' TITLE_REFERENCE_03_DB," & _
								"                  '' AUTO_REFERENCE_03_DB" & _
								"             FROM IPLASPROD.FIXEDASSET_DEPRECIATION_LEDGER HDR" & _
								"                  INNER JOIN IPLASPROD.FIXEDASSET FA" & _
								"                     ON HDR.FCBA = FA.FCBA AND HDR.FCCODE = FA.FCCODE" & _
								"                  INNER JOIN IPLASPROD.ASSETGROUP" & _
								"                     ON FA.ASSETGROUPCODE = IPLASPROD.ASSETGROUP.FCCODE" & _
								"                  INNER JOIN IPLASPROD.PERIODCONTROL" & _
								"                     ON     HDR.FCBA = IPLASPROD.PERIODCONTROL.FCBA" & _
								"                        AND UPPER (IPLASPROD.PERIODCONTROL.SYSTEM) = 'ASSET'" & _
								"                  INNER JOIN IPLASPROD.JOB" & _
								"                     ON     HDR.FCBA = IPLASPROD.JOB.FCBA" & _
								"                        AND IPLASPROD.ASSETGROUP.CONTROLJOB_DEPRECIATION =" & _
								"                               IPLASPROD.JOB.FCCODE" & _
								"                  INNER JOIN IPLASPROD.FIXEDASSET_MATURE FAM" & _
								"                     ON     HDR.FCBA = FAM.FCBA" & _
								"                        AND HDR.FCCODE = FAM.FIXEDASSETCODE" & _
								"                        AND FAM.FAMATURETYPE = 'ALL'" & _
								"                  RIGHT JOIN" & _
								"                  (SELECT ff.fcba," & _
								"                          ff.fccode," & _
								"                          ff.hectarageplanted," & _
								"                          fd.totalhectarageplanted," & _
								"                          (ff.hectarageplanted / fd.totalhectarageplanted)" & _
								"                             proporsiha" & _
								"                     FROM field ff" & _
								"                          INNER JOIN" & _
								"                          (  SELECT fcba," & _
								"                                    SUM (hectarageplanted) totalhectarageplanted" & _
								"                               FROM field" & _
								"                              WHERE fcba = '"& session("fcba") &"' AND DX_BLOCK_STATUS = 'TM'" & _
								"                           GROUP BY fcba) fd" & _
								"                             ON (fd.fcba = ff.fcba)" & _
								"                    WHERE ff.DX_BLOCK_STATUS = 'TM' AND ff.fcba = '"& session("fcba") &"') ffha" & _
								"                     ON (ffha.fcba = fam.fcba)" & _
								"            WHERE     TO_CHAR (HDR.DATE_TRANSACTION, 'fmMM') = "& Mid(Session("CurP_Asset"),5,2) &"" & _
								"                  AND TO_CHAR (HDR.DATE_TRANSACTION, 'yyyy') = "& Mid(Session("CurP_Asset"),1,4) &"" & _
								"                  AND HDR.FCBA = '"& session("fcba") &"'" & _
								"         GROUP BY HDR.FCBA," & _
								"                  HDR.FCCODE," & _
								"                  HDR.DOC_NO," & _
								"                  HDR.LINENO)"  

			Set rsFAAmorIntiGCCR = Server.CreateObject("ADODB.Recordset")
			rsFAAmorIntiGCCR.Open qFAAmorIntiGCCR, connstr
			if not rsFAAmorIntiGCCR.eof then
				Do While Not rsFAAmorIntiGCCR.eof
				
					 maxLineNo = maxLineNo + 1

					 qCost_ActivityCR = "Insert into Cost_Activity " & _
									 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									 " TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
									 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
									 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
									 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
									 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
									 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
									 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
									 " LASTUPDATE, LASTTIME ) " & _ 
									 " values (" & maxLineNo & ",to_date('" & rsFAAmorIntiGCCR("TDATE_CR") & "','mm/dd/yyyy') " & ",'" & rsFAAmorIntiGCCR("JOBCODE_CR") & "','" & rsFAAmorIntiGCCR("SOURCE_CR") & "','" & rsFAAmorIntiGCCR("SOURCETYPE_CR") & "','" & rsFAAmorIntiGCCR("SOURCECODE_CR") & "'," & _
									 " '','" & rsFAAmorIntiGCCR("TARGETTYPE_CR") & "','" & rsFAAmorIntiGCCR("TARGETCODE_CR") & "'," & _
									 " '" &  -1*Cdbl(rsFAAmorIntiGCCR("AMOUNT_CR")) & "','" & rsFAAmorIntiGCCR("DOCUMENTNO_CR") & "'," & "0,0,'" & rsFAAmorIntiGCCR("FCBA") & "',"  & _
									 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
									 " '" & rsFAAmorIntiGCCR("TITLE_REFERENCE_01_CR") & "','"  & rsFAAmorIntiGCCR("AUTO_REFERENCE_01_CR") & "'," & _
									 " '" & rsFAAmorIntiGCCR("TITLE_REFERENCE_02_CR") & "','"  & rsFAAmorIntiGCCR("AUTO_REFERENCE_02_CR") & "'," & _
									 " '" & rsFAAmorIntiGCCR("TITLE_REFERENCE_03_CR") & "','"  & rsFAAmorIntiGCCR("AUTO_REFERENCE_03_CR") & "'," & _
									 " '','', " & _
									 " 'Line No:', '" & rsFAAmorIntiGCCR("LINENO") & "','Cr', " & _
									 " " & rsFAAmorIntiGCCR("ACCOUNTINGPERIOD") & "," & rsFAAmorIntiGCCR("ACCOUNTINGYEAR") & "," & _
									 " sysdate, to_char(sysdate,'HH24:MI')) "
					
					 connstr.execute(qCost_ActivityCR)
					 connStr.execute("Commit")
					 
					 rsFAAmorIntiGCCR.MoveNext
				loop
			end if
		'-- End Insert Credit
			response.Write "Process Fixed Asset GC Amortisasi Multiple Blok Complete....<br>"
			response.Flush
	'-- End  Of FA Amortisasi GC
	'-- End  Of FA ALL, PDE, GC Allocated By Multiple 
	
	
	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Asset'" 

		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

		strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME) VALUES (" & _
					" 'Asset', " & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Asset'"
	
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.Open strLenghtime, connstr

	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Asset'"		
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
								" where UPPER(system) ='ASSET' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Asset " & "<br>"
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
				  " where SYSTEM='Asset' and fcba='" & Session("fcBA") & "'"
				ww = " Satu " 
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where SYSTEM='Asset' and fcba='" & Session("fcBA") & "'"
					  ww = " Dua " 
				ELSE
				    IF temp_NextCloseAccountingPeriod = 12 THEN
						 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
						  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear= " & temp_NextCurrentAccountingYear & " " & _  
						  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						  " where  SYSTEM='Asset' AND fcba='" & Session("fcBA") & "'"
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

			' Query ke Period Control untuk mengambli Session(v_Asset) & Session(cls_Asset)	
		
		
			txtLastDay  =cmdFindPeriod("Asset")
			txtNextMonth=cmdPeriodClose("Asset")

			Session("v_Asset")= txtLastDay
			Session("cls_Asset")= txtNextMonth
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
Asset Last Period Close <%=FormatDateTime(Session("cls_Asset"),1)%> <br> 
Asset Current Period is <%=FormatDateTime(Session("v_Asset"),1)%> <br><br>
Please Wait...<br>
<%
 Response.Write "Start ..." & Time() & "<br>" 
 
 ' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Asset','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Asset " & "<br>"
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
				<img src="images/ok.gif"/>Asset Closing is Completed!<br/>
			</font>
<%
else
		Call PostEndTime
%>
		<br/><br/>
		<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>
			<img src="images/error.gif"/>Asset Closing is Not Completed!!<br/>
		</font>
<%
end if
%>
	<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
	</script>
</body>
</html>