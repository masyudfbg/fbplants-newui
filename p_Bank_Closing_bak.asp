<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="adovbs.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>ePlaS - Bank Closing</title>
</head>
<body>
<%
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

Sub NormalisasiData
	'---- Normalisasi

	'qNormalisasiDetail = " delete from cost_activity_hist_tab " & _
	'					 " where " & _
	'					 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
	'							" where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Bank")) &  _
	'							" and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Bank")) & " and fcBA='" & Session("fcBA") & "' and Sourcetype='BA')"
	'					 
	'connstr.execute(qNormalisasiDetail)
    'connStr.execute("Commit")
	
	' qNormalisasiDetail = " update cost_activity_hist_tab set deleteflag = 'REMOVE'" & _
						 ' " where " & _
						 ' " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
						 ' " where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Bank")) &  _
						 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Bank")) & " and fcBA='" & Session("fcBA") & "' and Sourcetype='BA')"
	
    ' Add By Ali Akbar 28 Des 2015 
	
     qNormalisasiDetail = " update cost_activity_hist_tab set deleteflag = 'REMOVE'" & _
						 " where " & _
						 " lineheaderno || fcba in (select lineno||fcba from cost_activity " & _
						 " where  ACCOUNTINGPERIOD=" & Mid(Session("CurP_Bank"),5,2) &  _
						 " and ACCOUNTINGYEAR=" & Mid(Session("CurP_Bank"),1,4) & " and fcBA='" & Session("fcBA") & "' and Sourcetype='BA')"

	
	connstr.execute(qNormalisasiDetail)
    connStr.execute("Commit")
	
	qNormalisasiDetail = " delete from cost_activity_hist_tab where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

	connstr.execute(qNormalisasiDetail)
	connStr.execute("Commit")

	
	
	

	'qNormalisasi="Delete From Cost_Activity where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Bank")) &  _
	'			 " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Bank")) & " and fcBA='" & Session("fcBA") & "' and SourceType='BA'"
	'connstr.execute(qNormalisasi)
    'connStr.execute("Commit")
	
	' qNormalisasiDetail="update Cost_Activity set deleteflag = 'REMOVE' where  to_number(to_Char(TDATE,'mm'))=" & Month(Session("v_Bank")) &  _
				 ' " and to_number(to_Char(TDATE,'yyyy'))=" & Year(Session("v_Bank")) & " and fcBA='" & Session("fcBA") & "' and SourceType='BA'"
				 
    qNormalisasiDetail="update Cost_Activity set deleteflag = 'REMOVE' where  ACCOUNTINGPERIOD=" & Mid(Session("CurP_Bank"),5,2) &  _
				       " and ACCOUNTINGYEAR=" & Mid(Session("CurP_Bank"),1,4) & " and fcBA='" & Session("fcBA") & "' and SourceType='BA'"
	
	
	connstr.execute(qNormalisasiDetail)
    connStr.execute("Commit")

    qNormalisasiDetail = " delete cost_activity where deleteflag = 'REMOVE' and fcba =  '" & Session("fcBA") & "'"

	connstr.execute(qNormalisasiDetail)
	connStr.execute("Commit")
End Sub

Sub postingData
	' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	' Add by NAF to select lastperiodposted
	Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl
    
	
	
	qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted,LASTPERIODPOSTED LastEndPeriodPosted  from PERIODCONTROL  " & _
	  " where SYSTEM='Bank' and fcba='" & Session("fcBA") & "'"

	Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	rsPeriodControl.CursorType = 3
	rsPeriodControl.CursorLocation = 3
	rsPeriodControl.Open qPeriodControl, connstr

	if not rsPeriodControl.eof then
	   IF  Cint(MID(Session("CurP_Bank"),5,2)) < 11 Then
		   temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")     ' add Akbar 28 Des 2015 , for Account Period <= 12
	   ELSE
	       temp_datePeriodControl = rsPeriodControl("LASTENDPERIODPOSTED")  ' for Account Period > 12
       END IF 	   
	else
		Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Bank " & "<br>"
	end if
	' End Added
	'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


	' Remarked by NAF, as per New Concept, 7 May 2013
	' qControlJob="select ControlJob from Module where NAME='Purchase'"
	' Now we use new concept for picking Control Job, by using Process Control Tables
	' as per last concept are NOT able to pick data per Business Unit.

	tempProcessCode = "BANK_CLOSING"
	tempProcessNameCode = "p_Bank_Closing.asp"
	tempRefCodePVDetail = "PAYMENTVOUCHERDETAIL"

	' Take the Reference of Cost Activity, as per New Concept, by using this reference, we can directly refer to Original Table
	' by NAF, 12 May 2013
	' Take the Original Table of PAYMENTVOUCHERDETAIL.
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



	qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
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

	' End of Modification

	' Add By AKB  11 Des 2015, mengganti where datecrated menjadi accountingperiod dan accountingyear
	' menambahkan accountingperiod dan accountingyear ke cost activity.

	qVoucher="select p.VOUCHERCODE,p.BANKCODE,p.DATECREATED,p.accountingperiod,p.accountingyear," & _
			" pv.SOURCETYPE,pv.LOCATIONTYPE,pv.LOCATIONCODE,pv.JOBCODE,pv.REFERENCE, pv.lineno, sum(pv.AMOUNT) as Amount" & _
			" from PAYMENTVOUCHER P,PAYMENTVOUCHERDETAIL PV" & _
			" where (p.VOUCHERCODE = pv.VOUCHERCODE and p.fcba = pv.fcba)" & _
			" AND p.accountingperiod =" & Mid(Session("CurP_Bank"),5,2) &  _
			" and p.accountingyear=" & Mid(Session("CurP_Bank"),1,4) & _
			" and pv.fcBA='" & Session("fcBA") & "'" & _
			" group by p.VOUCHERCODE,p.BANKCODE,p.DATECREATED,p.accountingperiod,p.accountingyear," & _
			" pv.SOURCETYPE,pv.LOCATIONTYPE,pv.LOCATIONCODE,pv.JOBCODE,pv.REFERENCE, pv.lineno"
			
	Set rsVoucher = Server.CreateObject("ADODB.Recordset")
	rsVoucher.CursorType = 3
	rsVoucher.CursorLocation = 3
	rsVoucher.Open qVoucher, connstr
	if not rsVoucher.eof then
	   Do While Not rsVoucher.eof
		  '-- Control Job Bank     
		  txtControlJob=""
		  qPO=" select CONTROLJOB as CONTROLJOB FROM Bank" & _
			  " WHERE FCCode ='" & rsVoucher("BANKCODE")  & "'" & _
			  " and fcBA='" & Session("fcBA") & "'"
		  Set rsBank = Server.CreateObject("ADODB.Recordset")
		  rsBank.CursorType = 3
		  rsBank.CursorLocation = 3
		  rsBank.Open qPO, connstr
		  if not rsBank.eof then
			 txtControlJob = rsBank("CONTROLJOB") 

			 '---- Cek Source Type Original
			 if rsVoucher("SOURCETYPE")<>"" then
				txtOriginalST= rsVoucher("SOURCETYPE")
			 else
				txtOriginalST= "BA"
			 end if 
			 
			 ' We add NEW reference string, to make MORE simple in REFERENCE Report., by NAF, 12 May 2013
			 txtRefStringValues_PVDetail = Session("fcBA") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("VOUCHERCODE") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("LOCATIONTYPE") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("JOBCODE") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("LOCATIONCODE") & txtRefStringSeparator_PVDetail & _
										   rsVoucher("lineno") 
										   

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
							 " values (" & maxLineNo & ",to_date('" & rsVoucher("DATECREATED") & "','mm/dd/yyyy') " & ",'" & rsVoucher("JOBCODE") & "','" & txtOriginalST & "','" & "BA" & "','" & rsVoucher("BANKCODE") & "'," & _
							 "'" &  rsVoucher("BANKCODE") & "','" & rsVoucher("LOCATIONTYPE") & "','" & rsVoucher("LOCATIONCODE") & "'," & _
							 "'" &  rsVoucher("Amount") & "','" & rsVoucher("VOUCHERCODE") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " Cost_Activity_Info_Api.Get_Title_Reference_01('BA', '" & rsVoucher("LOCATIONTYPE") & "','Dr'),'"  & rsVoucher("LOCATIONCODE") & "'," & _
							 " Cost_Activity_Info_Api.Get_Title_Reference_02('BA', '" & rsVoucher("LOCATIONTYPE") & "','Dr'),'"  & rsVoucher("VOUCHERCODE") & "'," & _
							 " '','', " & _
							 " '','', " & _
							 " 'Line No:', '" & rsVoucher("lineno") & "','Dr', " & _
							 " " & rsVoucher("ACCOUNTINGPERIOD") & "," & rsVoucher("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
							 
			 connstr.execute(qCost_ActivityDB)
			 connStr.execute("Commit")
			

			' Insert History Transaction,
			
			qVoucherDetail ="select p.VOUCHERCODE,p.BANKCODE,p.DATECREATED," & _
							" pv.SOURCETYPE,pv.LOCATIONTYPE,pv.LOCATIONCODE,pv.JOBCODE,pv.REFERENCE,pv.lineno, pv.AMOUNT as Amount" & _
							" from PAYMENTVOUCHER P,PAYMENTVOUCHERDETAIL PV" & _
							" where (p.VOUCHERCODE = pv.VOUCHERCODE and p.fcba = pv.fcba)" & _
							" AND p.accountingperiod =" & Mid(Session("CurP_Bank"),5,2)&  _
							" and p.accountingyear=" & Mid(Session("CurP_Bank"),1,4) & _
							" and pv.fcBA='" & Session("fcBA") & "'" & _
							" and pv.VOUCHERCODE = '" & rsVoucher("VOUCHERCODE") & "' and " & _
							" pv.LOCATIONTYPE = '" & rsVoucher("LOCATIONTYPE") & "' and " & _
							" pv.JOBCODE = '" & rsVoucher("JOBCODE") & "' and " & _
							" pv.LOCATIONCODE = '" & rsVoucher("LOCATIONCODE")  & "'" & _
							" and pv.lineno = '" & rsVoucher("lineno")  & "'"
			
			Set rsVoucherDetail = Server.CreateObject("ADODB.Recordset")
			rsVoucherDetail.CursorType = 3
			rsVoucherDetail.CursorLocation = 3
			rsVoucherDetail.Open qVoucherDetail, connstr
			
			if not rsVoucherDetail.eof then
				maxLineNoDetail = 0
				Do While Not rsVoucherDetail.eof

					maxLineNoDetail = maxLineNoDetail + 1
					qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
									"( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
									"  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
									" QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
									maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodePVDetail & "','" & txtRefNumberKeys_PVDetail & "','" & _
									txtRefStringSeparator_PVDetail & "','" & txtRefStringKeys_PVDetail & "','" &  txtRefStringValues_PVDetail & "' , " & _
									"0,0,"& rsVoucherDetail("amount")&",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
								 
					connstr.execute(qCost_ActivityDBHistory)		
					connStr.execute("Commit")
					rsVoucherDetail.MoveNext
				loop
			end if
			rsVoucherDetail.Close
			set rsVoucherDetail = Nothing	
			
			nPurchaseControl=-1*cdbl(rsVoucher("Amount"))
			maxLineNo = maxLineNo + 1

			qCost_ActivityCR="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA , closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME) " & _
							 " values (" & maxLineNo & ", to_date('" & rsVoucher("DATECREATED") & "','mm/dd/yyyy') " & ",'" & txtControlJob & "','" & txtOriginalST & "','" & "BA" & "','" & rsVoucher("BANKCODE") & "'," & _
							 "'" &  rsVoucher("BANKCODE") & "','" & "CB" & "','" & "BS" & "'," & _
							 "'" & nPurchaseControl  & "','" & rsVoucher("VOUCHERCODE") & "'," & "0,0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & temp_datePeriodControl & "','mm/dd/yyyy') , sysdate, " & _
							 " Cost_Activity_Info_Api.Get_Title_Reference_01('BA', 'CB','Cr'),'"  & rsVoucher("LOCATIONCODE") & "'," & _
							 " Cost_Activity_Info_Api.Get_Title_Reference_02('BA', 'CB','Cr'),'"  & rsVoucher("VOUCHERCODE") & "'," & _
							 " '','', " & _
							 " '','', " & _
							 " 'Line No:', '" & rsVoucher("lineno") & "','Cr', " & _
							 " " & rsVoucher("ACCOUNTINGPERIOD") & "," & rsVoucher("ACCOUNTINGYEAR") & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI')) "
							 
							 
			 connstr.execute(qCost_ActivityCR)
			 connStr.execute("Commit")
			' Insert History Transaction,
			
			qVoucherDetail ="select p.VOUCHERCODE,p.BANKCODE,p.DATECREATED," & _
							" pv.SOURCETYPE,pv.LOCATIONTYPE,pv.LOCATIONCODE,pv.JOBCODE,pv.REFERENCE,pv.lineno, pv.AMOUNT as Amount" & _
							" from PAYMENTVOUCHER P,PAYMENTVOUCHERDETAIL PV" & _
							" where (p.VOUCHERCODE = pv.VOUCHERCODE and p.fcba = pv.fcba)" & _
							" AND to_number(to_Char(p.DATECREATED,'mm'))=" & Mid(Session("CurP_Bank"),5,2) &  _
							" and to_number(to_Char(p.DATECREATED,'yyyy'))=" & Mid(Session("CurP_Bank"),1,4) & _
							" and pv.fcBA='" & Session("fcBA") & "'" & _
							" and pv.VOUCHERCODE = '" & rsVoucher("VOUCHERCODE") & "' and " & _
							" pv.LOCATIONTYPE = '" & rsVoucher("LOCATIONTYPE") & "' and " & _
							" pv.JOBCODE = '" & rsVoucher("JOBCODE") & "' and " & _
							" pv.LOCATIONCODE = '" & rsVoucher("LOCATIONCODE")  & "'" & _
							" and pv.lineno = '" & rsVoucher("lineno")  & "'"
			
			Set rsVoucherDetail = Server.CreateObject("ADODB.Recordset")
			rsVoucherDetail.CursorType = 3
			rsVoucherDetail.CursorLocation = 3
			rsVoucherDetail.Open qVoucherDetail, connstr
			
			if not rsVoucherDetail.eof then
				maxLineNoDetail = 0
				Do While Not rsVoucherDetail.eof

					maxLineNoDetail = maxLineNoDetail + 1
					nPurchaseControl = -1*cdbl(rsVoucherDetail("Amount"))
					qCost_ActivityDBHistory = " INSERT INTO COST_ACTIVITY_HIST_TAB " & _
									"( LINEHEADERNO, LINENO, REF_PROCESS, REF_PROCESSNAME, " & _
									"  REF_TABLESOURCE, REF_NUMBERKEYS, REF_STRINGSEPARATOR, REF_STRINGKEYS, REF_STRINGVALUES, " & _
									" QUANTITY, OUTPUT, AMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
									maxLineNo & "," &  maxLineNoDetail & ",'" &  tempProcessCode & "','" & tempProcessNameCode & "','" & tempRefCodePVDetail & "','" & txtRefNumberKeys_PVDetail & "','" & _
									txtRefStringSeparator_PVDetail & "','" & txtRefStringKeys_PVDetail & "','" &  txtRefStringValues_PVDetail & "' , " & _
									"0,0,"& nPurchaseControl &",'" & Session("CodeName") & "','" & Session("CodeName") &  "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, to_char(sysdate,'HH24:MI'))"
								 
					connstr.execute(qCost_ActivityDBHistory)		
					connStr.execute("Commit")
					rsVoucherDetail.MoveNext
				loop
			end if
			rsVoucherDetail.Close
			set rsVoucherDetail = Nothing		
			
			response.Write "Process Bank... " & rsVoucher("VOUCHERCODE") & "<br>"
			response.Flush()
		  else
			 Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20011 Control Job For Bank " & rsVoucher("BANKCODE") & " is Empty" & "<br>"
		  End if
		  rsBank.Close
		  set rsBank = Nothing
		  
		  rsVoucher.MoveNext
	  loop
	End if
	
	
	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'Bank'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

	strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME) VALUES (" & _
					" 'Bank', " & _
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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Bank'"
				
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.CursorType = 3
	rsLenghtime.CursorLocation = 3
	rsLenghtime.Open strLenghtime, connstr

	upClosingLog = " UPDATE CLOSINGLOG SET" & _
					" LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'Bank'"		
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
								" where UPPER(system) ='BANK' and fcba='" & Session("fcBA") & "'"
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
					Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for Bank " & "<br>"
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
				  " where SYSTEM='Bank' and fcba='" & Session("fcBA") & "'"
				ww = " Satu " 
			ELSE
			    IF temp_NextCloseAccountingPeriod < 12 THEN
				     q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) ,CURRENTDATE=add_months(CURRENTDATE,1) " & _ 
					  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
					  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
					  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
					  " ,CurrentAccountingYear = " & temp_NextCurrentAccountingYear & " " & _  
					  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
					  " where SYSTEM='Bank' and fcba='" & Session("fcBA") & "'"
					  ww = " Dua " 
				ELSE
				    IF temp_NextCloseAccountingPeriod = 12 THEN
						 q=" Update PeriodControl Set LASTPERIODPOSTED=last_day(add_months(LASTPERIODPOSTED,1)) " & _ 
						  " ,CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear= " & temp_NextCurrentAccountingYear & " " & _  
						  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						  " where  SYSTEM='Bank' AND fcba='" & Session("fcBA") & "'"
					ELSE
					    q=" Update PeriodControl Set " & _ 
						  " CloseAccountingPeriod = " & temp_NextCloseAccountingPeriod & " " & _  
						  " ,CloseAccountingYear = " & temp_NextCloseAccountingYear & " " & _  
						  " ,CurrentAccountingPeriod = " & temp_NextCurrentAccountingPeriod & " " & _  
						  " ,CurrentAccountingYear= " & temp_NextCurrentAccountingYear & " " & _  
						  " ,lastupdate = sysdate, lasttime = to_char(sysdate, 'hh24:mi') " & _
						  " where  postingposition < 16  AND fcba='" & Session("fcBA") & "'"
					 
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

			' Query ke Period Control untuk mengambli Session(v_bank) & Session(cls_Bank)	
		
		
			txtLastDay  =cmdFindPeriod("Bank")
			txtNextMonth=cmdPeriodClose("Bank")

			Session("v_Bank")= txtLastDay
			Session("cls_Bank")= txtNextMonth
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
Bank Last Period Close <%=FormatDateTime(Session("cls_Bank"),1)%> <br> 
Bank Current Period is <%=FormatDateTime(Session("v_Bank"),1)%> <br><br>
Please Wait...<br>
<%
 Response.Write "Start ..." & Time() & "<br>" 
dim rsVoucher,rsBank,qVoucher,qPO,nPurchaseControl,txtControlJob,qCost_ActivityDB,qCost_ActivityCR,nQtyReceived,qNormalisasi,txtOriginalST

' New Concept Variable, By NAF, 12 May 2013
Dim tempProcessCode, tempProcessNameCode, tempRefCodePVDetail, qReference
Dim rsReferencePVDetail
Dim txtRefNumberKeys_PVDetail, txtRefStringSeparator_PVDetail , txtRefStringKeys_PVDetail
Dim txtRefStringValues_PVDetail
Dim donePVDetail

' End Modification.

'response.write "pc " & Cint(MID(Session("CurP_Bank"),5,2))
' Water Flow Check, By NAF
' 2 December 2013

Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('Bank','" & Session("fcBA")  & "') waterFlowCheck from dual "

Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
rsWaterFlowCheck.CursorType = 3
rsWaterFlowCheck.CursorLocation = 3
rsWaterFlowCheck.Open qWaterFlowCheck, connstr

if not rsWaterFlowCheck.eof then
    CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
else
    Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for Bank " & "<br>"
end if
rsWaterFlowCheck.close
set rsWaterFlowCheck = Nothing


if CheckWaterFlow = "TRUE" then	

	is_closing_last_month = -1
	oquery = " SELECT TO_NUMBER(TO_CHAR(SYSDATE, 'yyyymmdd')) -" & _
	         "        TO_NUMBER(TO_CHAR(LAST_DAY(TO_DATE('" & Mid(Session("CurP_Bank"),5,2) & "/1/" & Mid(Session("CurP_Bank"),1,4) & "', 'mm/dd/yyyy')), 'yyyymmdd')) result" & _
             " FROM DUAL"
	Set orecordset = Server.CreateObject("ADODB.Recordset")
	orecordset.Open oquery, connstr

	if not orecordset.eof then
		is_closing_last_month = Cdbl(orecordset("result")) ' jika >= 0 boleh, jika < 0 tidak
	else
		is_closing_last_month = -1
	end if
	orecordset.close
	set orecordset = Nothing
	
	if is_closing_last_month < 0 then
		Call NormalisasiData
		Call postingData
		Call PostEndTime
%>
			<br/><br/>
			<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>
				<img src="images/ok.gif"/>Bank Closing is Completed!<br/>
			</font>
<%
	else '<--validasi tanggal proses
	
		status_data_open = 0
		oquery = " SELECT COUNT(1) result" & _
				 " FROM recurringtransaction" & _
				 " WHERE fcba = '" & Session("fcBA") & "'" & _
				 "       AND accountingyear = '" & Mid(Session("CurP_Bank"),1,4) & "'" & _
				 "       AND accountingperiod = '" & Mid(Session("CurP_Bank"),5,2) & "'" & _
				 "       AND (UPPER(status) = UPPER('open') OR UPPER(status) = UPPER('opened'))"
		Set orecordset = Server.CreateObject("ADODB.Recordset")
		orecordset.Open oquery, connstr

		if not orecordset.eof then
			status_data_open = Cdbl(orecordset("result"))
		else
			status_data_open = 0
		end if
		orecordset.close
		set orecordset = Nothing
		
		if status_data_open > 0 then
			Call PostEndTime
%>
			<br/><br/>
			<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>
				<img src="images/error.gif"/>Bank Closing is Not Completed!<br/>
				<img src="images/error.gif"/>Masih ada Data Recurring yang outstanding! Silahkan lakukan Posting untuk Data recurring Bulan ini!<br/>
			</font>
<%
		else '<--cek status open
			
			status_data_posted = 0
			oquery = " SELECT COUNT(1) result" & _
					 " FROM recurringtransaction" & _
					 " WHERE fcba = '" & Session("fcBA") & "'" & _
					 "       AND accountingyear = '" & Mid(Session("CurP_Bank"),1,4) & "'" & _
					 "       AND accountingperiod = '" & Mid(Session("CurP_Bank"),5,2) & "'" & _
					 "       AND (UPPER(status) = UPPER('post') OR UPPER(status) = UPPER('posted'))"
			Set orecordset = Server.CreateObject("ADODB.Recordset")
			orecordset.Open oquery, connstr

			if not orecordset.eof then
				status_data_posted = Cdbl(orecordset("result"))
			else
				status_data_posted = 0
			end if
			orecordset.close
			set orecordset = Nothing
			
			max_lastupdate_payroll = ""
			oquery = " SELECT TO_CHAR(MAX(lastupdate), 'mm/dd/yyyy hh24:mi') result" & _
					 " FROM empallded" & _
					 " WHERE fcba = '" & Session("fcBA") & "'" & _
					 "       AND TO_CHAR(adhocdate, 'yyyymm') = " & Mid(Session("CurP_Bank"),1,4) & "||LPAD(" & Mid(Session("CurP_Bank"),5,2) & ", 2, 0)"
			Set orecordset = Server.CreateObject("ADODB.Recordset")
			orecordset.Open oquery, connstr

			if not orecordset.eof then
				max_lastupdate_payroll = orecordset("result")
			else
				max_lastupdate_payroll = ""
			end if
			orecordset.close
			set orecordset = Nothing
			
			if status_data_posted > 0 then
			
				arr_rt_eof = true
				oquery = " SELECT documentno, transactionno, NVL(havebeenexecute, 'FALSE') havebeenexecute" & _
						 " FROM recurringtransaction" & _
						 " WHERE fcba = '" & Session("fcBA") & "'" & _
						 "       AND accountingyear = '" & Mid(Session("CurP_Bank"),1,4) & "'" & _
						 "       AND accountingperiod = '" & Mid(Session("CurP_Bank"),5,2) & "'"
						 
				Set orecordset = Server.CreateObject("ADODB.Recordset")
				orecordset.Open oquery, connstr

				if not orecordset.eof then
					arr_rt = orecordset.getrows()
					arr_rt_eof = false
				end if
				orecordset.close
				set orecordset = Nothing
				
				if not arr_rt_eof then
					For xxx = LBound(arr_rt, 2) To UBound(arr_rt, 2)
					
						min_lastupdate_recurring = ""
						oquery = " SELECT TO_CHAR(MIN(lastupdate), 'mm/dd/yyyy hh24:mi') result" & _
								 " FROM recurringtransactiondtl" & _
								 " WHERE fcba = '" & Session("fcBA") & "'" & _
								 "       AND transactionno = '" & arr_rt(1, xxx) & "'"
						Set orecordset = Server.CreateObject("ADODB.Recordset")
						orecordset.Open oquery, connstr

						if not orecordset.eof then
							min_lastupdate_recurring = orecordset("result")
						else
							min_lastupdate_recurring = ""
						end if
						orecordset.close
						set orecordset = Nothing
						
						payroll_sign = ""
						oquery = " SELECT CASE SIGN(TO_DATE('" & max_lastupdate_payroll & "', 'mm/dd/yyyy hh24:mi')-TO_DATE('" & min_lastupdate_recurring & "', 'mm/dd/yyyy hh24:mi'))" & _
								 "        WHEN -1 THEN 'lower'" & _
								 "        WHEN  0 THEN 'equal'" & _
								 "        WHEN  1 THEN 'greater'" & _
								 "        END result" & _
								 " FROM DUAL"
						Set orecordset = Server.CreateObject("ADODB.Recordset")
						orecordset.Open oquery, connstr

						if not orecordset.eof then
							payroll_sign = orecordset("result")
						else
							payroll_sign = ""
						end if
						orecordset.close
						set orecordset = Nothing
						
						if payroll_sign = "greater" then
			
							arr_rtd_eof = true
							oquery = " SELECT DISTINCT voucherno" & _
									 " FROM recurringtransactiondtl" & _
									 " WHERE fcba = '" & Session("fcBA") & "'" & _
									 "       AND transactionno = '" & arr_rt(1, xxx) & "'"
									 
							Set orecordset = Server.CreateObject("ADODB.Recordset")
							orecordset.Open oquery, connstr

							if not orecordset.eof then
								arr_rtd = orecordset.getrows()
								arr_rtd_eof = false
							end if
							orecordset.close
							set orecordset = Nothing
							
							if not arr_rtd_eof then
								For xxxx = LBound(arr_rtd, 2) To UBound(arr_rtd, 2)
									oquery = " DELETE FROM paymentvoucherdetail" & _
											 " WHERE fcba = '" & Session("fcBA") & "'" & _
											 "       AND vouchercode = '" & arr_rtd(0, xxxx) & "'"
									connstr.execute(oquery)
									
									oquery = " DELETE FROM paymentvoucher" & _
											 " WHERE fcba = '" & Session("fcBA") & "'" & _
											 "       AND vouchercode = '" & arr_rtd(0, xxxx) & "'"
									connstr.execute(oquery)
								next
							end if
							
							oquery = " DELETE FROM recurringtransactiondtl" & _
									 " WHERE transactionno = '" & arr_rt(1, xxx) & "'" & _
									 "       AND fcba = '" & Session("fcBA") & "'" 
							connstr.execute(oquery)
							
							oquery = " UPDATE recurringtransaction" & _
									 " SET havebeenexecute = 'FALSE', executedate = NULL, status = 'OPEN'" & _
									 " WHERE transactionno = '" & arr_rt(1, xxx) & "'" & _
									 "       AND fcba = '" & Session("fcBA") & "'" 
							connstr.execute(oquery)
							
							fcba__ = Session("fcBA")
							documentno__ = arr_rt(0, xxx)
							transactionno__ = arr_rt(1, xxx)
							executedate__ = textDate(Mid(Session("CurP_Bank"),5,2) & "/" & GetLastDay(Cdbl(Mid(Session("CurP_Bank"),5,2)),Cdbl(Mid(Session("CurP_Bank"),1,4))) & "/" & Mid(Session("CurP_Bank"),1,4))
							fcentry__ = Session("CodeName")
							fcedit__ = Session("CodeName")
							fcip__ = Request.ServerVariables("REMOTE_ADDR")
							
							Set comm = Server.CreateObject("ADODB.Command")
							With comm 
								.CommandText = "{ CALL RECURRING_API.ACCRUAL_PROCESS (?,?,?,?,?,?,?,?) }"
								.CommandType = adCmdText
								.ActiveConnection = connStr
								.Prepared = True
								Set prmfcba = .CreateParameter("",adVarchar,adParamInput,255)
								Set prmdocumentno = .CreateParameter("",adVarchar,adParamInput,255)
								Set prmtransactionno = .CreateParameter("",adVarchar,adParamInput,255)
								Set prmexecutedate = .CreateParameter("",adDate,adParamInput)
								Set prmfcentry = .CreateParameter("",adVarchar,adParamInput,255)
								Set prmfcedit = .CreateParameter("",adVarchar,adParamInput,255)
								Set prmfcip = .CreateParameter("",adVarchar,adParamInput,255)
								Set prmresult = .CreateParameter("",adLongVarChar,adParamOutput,1000)
								.Parameters.Append(prmfcba)
								.Parameters.Append(prmdocumentno)
								.Parameters.Append(prmtransactionno)
								.Parameters.Append(prmexecutedate)
								.Parameters.Append(prmfcentry)
								.Parameters.Append(prmfcedit)
								.Parameters.Append(prmfcip)
								.Parameters.Append(prmresult)
							End With

							With comm 
								.Parameters(0).Value = fcba__
								.Parameters(1).Value = documentno__
								.Parameters(2).Value = transactionno__
								.Parameters(3).Value = executedate__
								.Parameters(4).Value = fcentry__
								.Parameters(5).Value = fcedit__
								.Parameters(6).Value = fcip__
								ResultCommand__ = .Parameters(7).Value
								Set rsrecurring  = .Execute
								' Response.Write(ResultCommand_)
							End With
						end if '<--Jika tanggal payroll > tanggal eksekusi recurring
					next
				end if
			end if '<--cek status posted
			
			Call NormalisasiData
			Call postingData
			Call PostEndTime
			
		end if '<--cek status open
	end if '<--validasi tanggal proses
%> 
<%
else
	Call PostEndTime
%>
	<br/><br/>
	<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>
		<img src="images/error.gif"/>Bank Closing is Not Completed!<br/>
		<img src="images/error.gif"/>Please Check Previous Closing Process, it's still not done yet<br/>
	</font>
<%
end if
%>
	<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
	</script>
</body>
</html>