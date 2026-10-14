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
				     " where ACCOUNTINGPERIOD = " & Month(Session("v_Sales")) & "" & _
				     "       and ACCOUNTINGYEAR = " & Year(Session("v_Sales")) & "" & _
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
		' qNormalisasiHPP = " DELETE FROM hargapokokpenjualan WHERE accountingperiod = '"& Mid(Session("CurP_Sales"),5,2) &"'" & _
		' 					 " AND accountingyear = '"& Mid(Session("CurP_Sales"),1,4) &"'	AND fcba =  '" & Session("fcBA") & "'"
		' response.write qNormalisasiHPP
		' connstr.execute(qNormalisasiHPP)
	' End Normalisasi Insert Table HPP Penjualan Added By SAG 08 Juni 2019
End Sub

Sub postingData
	' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX 
	' Add by NAF to select lastperiodposted
	Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl
	Dim qSalesDB,rsSalesDB,maxLineNo
    
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
	' txtMonth&"/01/"&txtYear
	if Mid(Session("CurP_Sales"),5,2) > 12 Then
	currentdate = "12"& Mid(Session("CurP_Sales"),1,4)
	else
	currentdate = Mid(Session("CurP_Sales"),5,2)&""& Mid(Session("CurP_Sales"),1,4)
	end if
	

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

	qMaxLineNo = " select trunc(currentdate,'yyyy') valid_from from periodcontrol where system = 'Sales' and fcba = '" & Session("fcBA") & "' "
					
	Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
	rsqMaxLineNo.CursorType = 3
	rsqMaxLineNo.CursorLocation = 3
	rsqMaxLineNo.Open qMaxLineNo, connstr

	if not rsqMaxLineNo.eof then
		valid_from = rsqMaxLineNo("valid_from")	
	end if
	rsqMaxLineNo.close
	set rsqMaxLineNo = Nothing
	
	
	' End of Modification

	' Add By IWA  24 Mar 2019
	
	if Session("Estate_Mill") = "E" then
	' AMH Journal Sales, dari Invoice sales TBS
	qSalesDB =  " SELECT b.salescontractno,d.productcode,a.invoicecode,b.customercode,c.description,b.amount_beforevat,b.amountvat,b.grandtotal,no_faktur,fakturdate,b.transdate," & _                                                                       
				" CASE RTrim(d.productcode) " & _
				" WHEN 'TBS' THEN '1130104' " & _
				" ELSE                      " & _  
				" '1130601' END JobDebit,   " & _ 
				" CASE RTrim(d.productcode) " & _
				" WHEN 'TBS' THEN '5110106' " & _
				" ELSE                      " & _
				" '5110105' END JobKredit,  " & _
				" '1150307' jobPPNDebit,    " & _
				" '2110305' Jobppnkredit,   " & _
				" Sum(a.quantity) qty,Sum(a.amount) amt " & _
				" FROM salesinvoicedtl a " & _
				" INNER join salesinvoicehdr b ON a.invoicecode=b.invoicecode AND a.fcba=b.fcba " & _
				" INNER JOIN customer c ON b.customercode=c.customercode AND b.fcba=c.fcba " & _
				" INNER JOIN salesagreementdet d ON b.salescontractno=d.salescontractno AND b.fcba=d.fcba " & _
				" WHERE a.state='INVOICED' " & _
				" AND b.fcba ='" & Session("fcBA") & "'" & _
			    " AND to_number(to_Char(b.transdate,'mm'))=" & Month(Session("v_Sales")) &  _
			    " AND to_number(to_Char(b.transdate,'yyyy'))=" & Year(Session("v_Sales")) & _
				" GROUP BY " & _
				"	b.salescontractno,d.productcode,a.invoicecode,b.customercode,c.description,b.amount_beforevat,b.amountvat,no_faktur,fakturdate, " & _
				"	b.transdate,grandtotal " 

    'response.write qSalesDB
    Set rsSalesDB = Server.CreateObject("ADODB.Recordset")
	rsSalesDB.Open qSalesDB, connstr
	
	if not rsSalesDB.eof then
		Do While Not rsSalesDB.eof	
	connstr.BeginTrans
			maxLineNo = maxLineNo + 1
			' Debit 
			qCost_ActivityDB ="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME, journalprefix ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("JobDebit") & "','SL','SL','" & rsSalesDB("productcode") & "'," & _
							 "'','AR','" & rsSalesDB("customercode") & "'," & _
							 "'" &  rsSalesDB("grandtotal") & "','" & rsSalesDB("invoicecode") & "'," & rsSalesDB("qty") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & "Invoice NO : " & "' ,'"  & rsSalesDB("invoicecode") & "'," & _
							 " '" & "Contact No/ Customer: " & "' ,'"  & rsSalesDB("salescontractno") & "/" & rsSalesDB("description") & "'," & _
							 " '" & "Faktur Pajak " & "','" & rsSalesDB("no_faktur")  & " " & rsSalesDB("fakturdate") &"', " & _
							 " '-','Penjualan " & rsSalesDB("productcode") & " Quantity: " & rsSalesDB("qty") &" ke " & rsSalesDB("description") & " Periode:" & Year(Session("v_Sales")) & "/" & Month(Session("v_Sales")) & "', " & _
							 " '','','Db', " & _
							 " " & Month(Session("v_Sales")) & "," & Year(Session("v_Sales")) & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI'),'SLS') "

'response.write qCost_ActivityDB & "<br>"

			 connstr.execute(qCost_ActivityDB)
			 
			 ' Kredit 
			maxLineNo = maxLineNo + 1
			qCost_ActivityDB ="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME, journalprefix ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("JobKredit") & "','SL','SL','" & rsSalesDB("productcode") & "'," & _
							 "'','AR','" & rsSalesDB("customercode") & "'," & _
							 "" &  rsSalesDB("amount_beforevat") & "*-1,'" & rsSalesDB("invoicecode") & "'," & rsSalesDB("qty") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & "Invoice NO : " & "' ,'"  & rsSalesDB("invoicecode") & "'," & _
							 " '" & "Contact No/ Customer: " & "' ,'"  & rsSalesDB("salescontractno") & "/" & rsSalesDB("description") & "'," & _
							 " '" & "Faktur Pajak " & "','" & rsSalesDB("no_faktur")  & " " & rsSalesDB("fakturdate") &"', " & _
							 " '-','Penjualan " & rsSalesDB("productcode") & " Quantity: " & rsSalesDB("qty") &" ke " & rsSalesDB("description") & " Periode:" & Year(Session("v_Sales")) & "/" & Month(Session("v_Sales")) & "', " & _
							 " '','','Cr', " & _
							 " " & Month(Session("v_Sales")) & "," & Year(Session("v_Sales")) & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI'),'SLS') "

'response.write qCost_ActivityDB & "<br>"

			 connstr.execute(qCost_ActivityDB)
			 
			 ' Debit PPN
			 maxLineNo = maxLineNo + 1
			qCost_ActivityDB ="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME, journalprefix ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("jobPPNDebit") & "','SL','SL','" & rsSalesDB("productcode") & "'," & _
							 "'','CB','BS'," & _
							 "" &  rsSalesDB("amountvat") & ",'" & rsSalesDB("invoicecode") & "'," & rsSalesDB("qty") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & "Invoice NO : " & "' ,'"  & rsSalesDB("invoicecode") & "'," & _
							 " '" & "Contact No/ Customer: " & "' ,'"  & rsSalesDB("salescontractno") & "/" & rsSalesDB("description") & "'," & _
							 " '" & "Faktur Pajak " & "','" & rsSalesDB("no_faktur")  & " " & rsSalesDB("fakturdate") &"', " & _
							 " '-','Penjualan " & rsSalesDB("productcode") & " Quantity: " & rsSalesDB("qty") &" ke " & rsSalesDB("description") & " Periode:" & Year(Session("v_Sales")) & "/" & Month(Session("v_Sales")) & "', " & _
							 " '','','Db', " & _
							 " " & Month(Session("v_Sales")) & "," & Year(Session("v_Sales")) & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI'),'SLS') "

'response.write qCost_ActivityDB & "<br>"

			' connstr.execute(qCost_ActivityDB)
			 
			 ' Kredit PPN
			maxLineNo = maxLineNo + 1
			qCost_ActivityDB ="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME, journalprefix ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("Jobppnkredit") & "','SL','SL','" & rsSalesDB("productcode") & "'," & _
							 "'','CB','BS'," & _
							 "" &  rsSalesDB("amountvat") & "*-1,'" & rsSalesDB("invoicecode") & "'," & rsSalesDB("qty") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & "Invoice NO : " & "' ,'"  & rsSalesDB("invoicecode") & "'," & _
							 " '" & "Contact No/ Customer: " & "' ,'"  & rsSalesDB("salescontractno") & "/" & rsSalesDB("description") & "'," & _
							 " '" & "Faktur Pajak " & "','" & rsSalesDB("no_faktur")  & " " & rsSalesDB("fakturdate") &"', " & _
							 " '-','Penjualan " & rsSalesDB("productcode") & " Quantity: " & rsSalesDB("qty") &" ke " & rsSalesDB("description") & " Periode:" & Year(Session("v_Sales")) & "/" & Month(Session("v_Sales")) & "', " & _
							 " '','','Cr', " & _
							 " " & Month(Session("v_Sales")) & "," & Year(Session("v_Sales")) & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI'),'SLS') "

'response.write qCost_ActivityDB & "<br>"

			 connstr.execute(qCost_ActivityDB)
		
				If Err.Number = 0 Then
					 connstr.CommitTrans
					 Response.Flush
				 Else
					 connstr.RollbackTrans
					 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_ActivityDB&"] ("&Err.Description&") <br/></font>")
					 Response.Flush
				 End If		


			
        	rsSalesDB.MoveNext
		loop
	 	end if
	
	qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT fcba,referenceno,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT fcba,referenceno,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("v_Sales")) & " AND accountingyear=" & Year(Session("v_Sales")) & " AND journalprefix='SLS' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.referenceno =  y.referenceno and x.fcba=y.fcba) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
					
     connstr.execute(qjournal)

	 response.Write "Closing Sales Estate Complete.... !!<br>"
	 response.Flush
	'-- End sales TBS
	end if
	
	if Session("Estate_Mill") = "M" then
	' AMH Journal Sales, dari Invoice sales
	' Db 1130101 Piutang Dagang CPO
	
	qSalesDB =  " SELECT b.salescontractno,d.productcode,a.invoicecode,b.customercode,c.description,b.amount_beforevat,b.amountvat,b.grandtotal,no_faktur,fakturdate,b.transdate," & _                                                                       
				" CASE RTrim(d.productcode) " & _
				" WHEN 'CPO' THEN '1130101' " & _
				" WHEN 'KER' THEN '1130102' " & _
				" WHEN 'CKG' THEN '1130103' " & _
				" ELSE                      " & _  
				" '1130601' END JobDebit,   " & _ 
				" CASE RTrim(d.productcode) " & _
				" WHEN 'CPO' THEN '5110101' " & _
				" WHEN 'KER' THEN '5110102' " & _
				" WHEN 'CKG' THEN '5110103' " & _
				" WHEN 'FBR' THEN '5110104' " & _
				" WHEN 'TBS' THEN '5110106' " & _
				" ELSE                      " & _
				" '5110105' END JobKredit,  " & _
				" '1150307' jobPPNDebit,    " & _
				" '2110305' Jobppnkredit,   " & _
				" Sum(a.quantity) qty,Sum(a.amount) amt " & _
				" FROM salesinvoicedtl a " & _
				" INNER join salesinvoicehdr b ON a.invoicecode=b.invoicecode AND a.fcba=b.fcba " & _
				" INNER JOIN customer c ON b.customercode=c.customercode AND b.fcba=c.fcba " & _
				" INNER JOIN salesagreementdet d ON b.salescontractno=d.salescontractno AND b.fcba=d.fcba " & _
				" WHERE a.state='INVOICED' " & _
				" AND b.fcba ='" & Session("fcBA") & "'" & _
			    " AND to_number(to_Char(b.transdate,'mm'))=" & Month(Session("v_Sales")) &  _
			    " AND to_number(to_Char(b.transdate,'yyyy'))=" & Year(Session("v_Sales")) & _
				" GROUP BY " & _
				"	b.salescontractno,d.productcode,a.invoicecode,b.customercode,c.description,b.amount_beforevat,b.amountvat,no_faktur,fakturdate, " & _
				"	b.transdate,grandtotal " 

    'response.write qSalesDB
    Set rsSalesDB = Server.CreateObject("ADODB.Recordset")
	rsSalesDB.Open qSalesDB, connstr
	
	if not rsSalesDB.eof then
		Do While Not rsSalesDB.eof		

		connstr.BeginTrans
			maxLineNo = maxLineNo + 1
			' Debit 
			qCost_ActivityDB ="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME, journalprefix ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("JobDebit") & "','SL','SL','" & rsSalesDB("productcode") & "'," & _
							 "'','AR','" & rsSalesDB("customercode") & "'," & _
							 "'" &  rsSalesDB("grandtotal") & "','" & rsSalesDB("invoicecode") & "'," & rsSalesDB("qty") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & "Invoice NO : " & "' ,'"  & rsSalesDB("invoicecode") & "'," & _
							 " '" & "Contact No/ Customer: " & "' ,'"  & rsSalesDB("salescontractno") & "/" & rsSalesDB("description") & "'," & _
							 " '" & "Faktur Pajak " & "','" & rsSalesDB("no_faktur")  & " " & rsSalesDB("fakturdate") &"', " & _
							  " '-','Penjualan " & rsSalesDB("productcode") & " Quantity: " & rsSalesDB("qty") &" ke " & rsSalesDB("description") & " Periode:" & Year(Session("v_Sales")) & "/" & Month(Session("v_Sales")) & "', " & _
							 " '','','Db', " & _
							 " " & Month(Session("v_Sales")) & "," & Year(Session("v_Sales")) & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI'),'SLS') "

			 connstr.execute(qCost_ActivityDB)
			 
			 ' Kredit 
			maxLineNo = maxLineNo + 1
			qCost_ActivityDB ="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME, journalprefix ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("JobKredit") & "','SL','SL','" & rsSalesDB("productcode") & "'," & _
							 "'','AR','" & rsSalesDB("customercode") & "'," & _
							 "" &  rsSalesDB("amount_beforevat") & "*-1,'" & rsSalesDB("invoicecode") & "'," & rsSalesDB("qty") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & "Invoice NO : " & "' ,'"  & rsSalesDB("invoicecode") & "'," & _
							 " '" & "Contact No/ Customer: " & "' ,'"  & rsSalesDB("salescontractno") & "/" & rsSalesDB("description") & "'," & _
							 " '" & "Faktur Pajak " & "','" & rsSalesDB("no_faktur")  & " " & rsSalesDB("fakturdate") &"', " & _
							 " '-','Penjualan " & rsSalesDB("productcode") & " Quantity: " & rsSalesDB("qty") &" ke " & rsSalesDB("description") & " Periode:" & Year(Session("v_Sales")) & "/" & Month(Session("v_Sales")) & "', " & _
							 " '','','Cr', " & _
							 " " & Month(Session("v_Sales")) & "," & Year(Session("v_Sales")) & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI'),'SLS') "

			 connstr.execute(qCost_ActivityDB)
			 
			 ' Debit PPN
			 maxLineNo = maxLineNo + 1
			qCost_ActivityDB ="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME, journalprefix ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("jobPPNDebit") & "','SL','SL','" & rsSalesDB("productcode") & "'," & _
							 "'','CB','BS'," & _
							 "" &  rsSalesDB("amountvat") & ",'" & rsSalesDB("invoicecode") & "'," & rsSalesDB("qty") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & "Invoice NO : " & "' ,'"  & rsSalesDB("invoicecode") & "'," & _
							 " '" & "Contact No/ Customer: " & "' ,'"  & rsSalesDB("salescontractno") & "/" & rsSalesDB("description") & "'," & _
							 " '" & "Faktur Pajak " & "','" & rsSalesDB("no_faktur")  & " " & rsSalesDB("fakturdate") &"', " & _
							 " '-','Penjualan " & rsSalesDB("productcode") & " Quantity: " & rsSalesDB("qty") &" ke " & rsSalesDB("description") & " Periode:" & Year(Session("v_Sales")) & "/" & Month(Session("v_Sales")) & "', " & _
							 " '','','Db', " & _
							 " " & Month(Session("v_Sales")) & "," & Year(Session("v_Sales")) & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI'),'SLS') "

'response.write qCost_ActivityDB & "<br>"

			' connstr.execute(qCost_ActivityDB)
			 
			 ' Kredit PPN
			maxLineNo = maxLineNo + 1
			qCost_ActivityDB ="Insert into Cost_Activity " & _
							 "(LINENO, TDATE, JOBCODE, SOURCE,SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
							 "TARGETTYPE, TARGETCODE, AMOUNT, REFERENCENO, QUANTITY, OUTPUT,FCBA, closing_date, process_date, " & _
							 " TITLE_REFERENCE_01, AUTO_REFERENCE_01, " & _
							 " TITLE_REFERENCE_02, AUTO_REFERENCE_02, " & _
							 " TITLE_REFERENCE_03, AUTO_REFERENCE_03, " & _
							 " TITLE_REFERENCE_04, AUTO_REFERENCE_04, " & _
							 " TITLE_REFERENCE_05, AUTO_REFERENCE_05, DEBETCREDIT, " & _
							 " ACCOUNTINGPERIOD, ACCOUNTINGYEAR, " & _
							 " LASTUPDATE, LASTTIME, journalprefix ) " & _
							 " values (" & maxLineNo & ",to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') " & ",'" & rsSalesDB("Jobppnkredit") & "','SL','SL','" & rsSalesDB("productcode") & "'," & _
							 "'','CB','BS'," & _
							 "" &  rsSalesDB("amountvat") & "*-1,'" & rsSalesDB("invoicecode") & "'," & rsSalesDB("qty") & ",0,'" & Session("fcBA") & "',"  & _
							 " to_date('" & rsSalesDB("transdate") & "','mm/dd/yyyy') , sysdate, " & _
							 " '" & "Invoice NO : " & "' ,'"  & rsSalesDB("invoicecode") & "'," & _
							 " '" & "Contact No/ Customer: " & "' ,'"  & rsSalesDB("salescontractno") & "/" & rsSalesDB("description") & "'," & _
							 " '" & "Faktur Pajak " & "','" & rsSalesDB("no_faktur")  & " " & rsSalesDB("fakturdate") &"', " & _
							 " '-','Penjualan " & rsSalesDB("productcode") & " Quantity: " & rsSalesDB("qty") &" ke " & rsSalesDB("description") & " Periode:" & Year(Session("v_Sales")) & "/" & Month(Session("v_Sales")) & "', " & _
							 " '','','Cr', " & _
							 " " & Month(Session("v_Sales")) & "," & Year(Session("v_Sales")) & "," & _
							 " sysdate, to_char(sysdate,'HH24:MI'),'SLS') "

'response.write qCost_ActivityDB & "<br>"

			 connstr.execute(qCost_ActivityDB)
		
				If Err.Number = 0 Then
					 connstr.CommitTrans
					 Response.Flush
				 Else
					 connstr.RollbackTrans
					 Response.write("<font face=Verdana size=1 color=Red>" & "Error Save ["&qCost_ActivityDB&"] ("&Err.Description&") <br/></font>")
					 Response.Flush
				 End If		


			
        	rsSalesDB.MoveNext
		loop
	 	end if
	
	qjournal=	" MERGE INTO cost_activity x " & _
					" USING " & _
					" ( " & _
					"  SELECT fcba,referenceno,idx||LPAD(ROWNUM,4,0) seq from " & _
					"  ( " & _
					"  SELECT DISTINCT fcba,referenceno,'" & session("COMPANYNO") & "/'||fcba||'/'||journalprefix||'/'||LPad(accountingperiod,2,0)||'/'||accountingyear||'/' as idx  " & _
					"  FROM cost_activity WHERE fcba='" & Session("FCBA") & "' AND accountingperiod="& Month(Session("CurP_Sales")) & " AND accountingyear=" & Year(Session("CurP_Sales")) & " AND journalprefix='SLS' " & _
					"  Order BY referenceno " & _
					"  ) " & _
					" )y ON (x.referenceno =  y.referenceno and x.fcba=y.fcba) " & _
					" WHEN MATCHED THEN " & _
					" UPDATE SET x.journalno = y.seq " 
					
     connstr.execute(qjournal)

	 response.Write "Closing Sales Mill Complete.... !!<br>"
	 response.Flush

    end if
	 ' -- END Session("Estate_Mill") = "M"

	' if Session("Estate_Mill") = "HO" then
	' '-- Begin sales Invoice	

	
	' end if

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