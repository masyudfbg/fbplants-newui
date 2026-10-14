<%@Language=VBScript%>
<% 
'Option Explicit 
%>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<% Dim textReportName : textReportName = "E R P Report"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>ePlaS - Generate <%=PCase(textReportName)%></title>
	<link rel="StyleSheet" href="style-content.css" type="text/css" />
<style>
* {font-family:Verdana,Arial,sans-serif;font-size:11px;}
</style>
</head>
<body>
<%
Dim TimeAwal,TimeAkhir,Hitung,Lama

Dim ScriptExecuted

ScriptExecuted = ""
ScriptExecuted = Request("ScriptExecuted")



TimeAwal=Time()

'--Added by Ukie at 25-Oct-2012--	
Dim Startingtime, Endingtime
Startingtime = FormatDateTime(CDate(Now))
'--End Added--

Dim v_tahun, v_Bulan

if ScriptExecuted = "ByVariable" then
	v_tahun = Request("txtYear")
	v_Bulan = Request("txtMonth")
else
	Session("CurP_Costbook") = cmdFindAccountPeriod("Costbook")
	Session("ClsP_Costbook") = cmdAccountPeriodClose("Costbook")
	v_tahun = Cdbl( Left( Session("ClsP_Costbook") , 4))
	v_Bulan = Cdbl( Right( Session("ClsP_Costbook") , 2))
end if

%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
	<br/><!--Enter-->
	Anda Sedang Memproses <strong><%=PCase(textReportName)%></strong> Bulan: <strong><%=BulanAP(v_Bulan)%></strong> Tahun: <strong><%=v_tahun%></strong>
	<br/><!--Enter-->
	<br/><!--Enter-->
	Please Wait...
	<br/><!--Enter-->
<%
 Response.Write "Start ..." & Time() & "<br><br>" 


' Start Modification to Add Generate ERP Process

'Response.Write "Hai ..." & "<br>"
'Response.Write "Generate ERP Report Account for Financial Controlller..." & "<br>"
Response.Flush()

Dim strSQLReversed, rsstrSQLReversed,  rsDate
Dim tempYear, tempMonth, strDate, tempDate
Dim qInsertData, doContinueProcess

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' Add by NAF to select lastperiodposted
Dim qPeriodControl, rsPeriodControl, temp_datePeriodControl


' Water Flow Check, By NAF
' 2 December 2013

' Dim qWaterFlowCheck, rsWaterFlowCheck, CheckWaterFlow, tempErrorWaterFlow

' qWaterFlowCheck = " select PERIODCONTROL_API.getStatusPrevClosing('ERPReport','" & Session("fcBA")  & "') waterFlowCheck from dual "

' Set rsWaterFlowCheck = Server.CreateObject("ADODB.Recordset")
' rsWaterFlowCheck.CursorType = 3
' rsWaterFlowCheck.CursorLocation = 3
' rsWaterFlowCheck.Open qWaterFlowCheck, connstr

' if not rsWaterFlowCheck.eof then
    ' CheckWaterFlow = rsWaterFlowCheck("waterFlowCheck")
' else
    ' Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No WATER FLOW CHECK for ERPReport " & "<br>"
' end if
' rsWaterFlowCheck.close
' set rsWaterFlowCheck = Nothing


'if CheckWaterFlow = "TRUE" then

Dim rsGetParameterTrue, qGetParameterTrue, tempDetail, strSQLNormalization, txtLastDay, txtNextMonth, q, qLineMax, rsqLineMax, tempMaxLine, strClosingLog, strLenghtime, rsLenghtime



	' qPeriodControl=" select  add_months(LASTPERIODPOSTED,1) LastPeriodPosted from PERIODCONTROL  " & _
	  ' " where SYSTEM='ERPReport' and fcba='" & Session("fcBA") & "'"

	' Set rsPeriodControl = Server.CreateObject("ADODB.Recordset")
	' rsPeriodControl.CursorType = 3
	' rsPeriodControl.CursorLocation = 3
	' rsPeriodControl.Open qPeriodControl, connstr

	' if not rsPeriodControl.eof then
		' temp_datePeriodControl = rsPeriodControl("LASTPERIODPOSTED")
	' else
		' Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error, No PERIODCONTROL for ERPReport " & "<br>"
	' end if
	' End Added

	doContinueProcess = "TRUE"

	' First we go for Report Detail SAP - FICO
	Response.Write "Start Normalization for Financial Controlller Data..." 

	Set rsGetParameterTrue = Server.CreateObject("ADODB.Recordset")
	rsGetParameterTrue.CursorType = 3
	rsGetParameterTrue.CursorLocation = 3

	qGetParameterTrue = "select param_6 from parameter " & _
						 " where " & _
						 " fcba = '" & Session("fcBA") & "' "  & _
						 " and parcode = 'ERPMAP'  " & _
						 " and fccode = 'COMPANY' "
						 
	rsGetParameterTrue.Open qGetParameterTrue, connstr
	tempDetail = "FALSE"
	if not rsGetParameterTrue.eof then
		tempDetail = rsGetParameterTrue ("param_6")
	else
		tempDetail = "FALSE"
	end if
	rsGetParameterTrue.close
	set rsGetParameterTrue = Nothing


	'Response.write "NAF001" accountingperiod,accountingyear 


	strSQLNormalization = " delete erp_report_sap_generated " & _
						  " where " & _
						  " fcba = '" & Session("fcBA") & "' "  & _
						  " and accountingperiod = " & v_Bulan  &  _
						  " and accountingyear= " & v_Tahun
						  
	connStr.execute(strSQLNormalization)
	Response.Write "(DONE)" & "<br>"

	Response.Write "Start Generate for Financial Controlller Data..., Year :" & v_Tahun & ", month : " &  v_Bulan  


	if tempDetail = "TRUE" then 


		' Get the Intraco Account
		Set rsGetParameterTrue = Server.CreateObject("ADODB.Recordset")
		rsGetParameterTrue.CursorType = 3
		rsGetParameterTrue.CursorLocation = 3

		qGetParameterTrue = "select param_1 from parameter " & _
							 " where " & _
							 " fcba = '" & Session("fcBA") & "' "  & _
							 " and parcode = 'ERPMAP'  " & _
							 " and fccode = 'INTRACOACCOUNT' "
							 
		rsGetParameterTrue.Open qGetParameterTrue, connstr
		tempIntracoAccount = ""
		if not rsGetParameterTrue.eof then
			tempIntracoAccount = rsGetParameterTrue ("param_1")
		else
			tempIntracoAccount = ""
		end if
		rsGetParameterTrue.close
		set rsGetParameterTrue = Nothing	
		
		if tempIntracoAccount = "" then 
			Response.Write "Error, Intraco Account NOT Defined in PARAM_1 Parcode ERPMAP, fccode INTRACOACCOUNT ... Please check your parameter"
			doContinueProcess = "FALSE"
		else
			doContinueProcess = "TRUE"
		end if
		
		if doContinueProcess = "TRUE" then 
			Set rsSAPFormated = Server.CreateObject("ADODB.Recordset")
			rsSAPFormated.CursorType = 3
			rsSAPFormated.CursorLocation = 3

			qSAPFormated = " select * from erp_report_sap " & _
								  " where " & _
								  " fcba = '" & Session("fcBA") & "' "  & _
								  " and ACCOUNTINGPERIOD = " & v_Bulan  &  _
								  " and ACCOUNTINGYEAR = " & v_Tahun




			rsSAPFormated.Open qSAPFormated, connstr
			if not rsSAPFormated.eof then
				Do While Not rsSAPFormated.eof
				
					temp_Field_Info_01 = rsSAPFormated("FIELD_INFO_01")
					if temp_Field_Info_01 = rsSAPFormated("COMPANY_CODE") or temp_Field_Info_01 = "" then 
				
						qInsertData = " insert into erp_report_sap_generated " & _
									  " ( DOCUMENT_DATE, DOCUMENT_TYPE, COMPANY_CODE, POSTING_DATE, PERIOD, " & _
									  "   CURRENCY, REFERENCE, HEADER_TEXT, POSTING_KEY, BUSINESS_AREA, GL_ACCOUNT, " & _
									  "   COST_CENTER, INTERNAL_ORDER, DOCUMENT_CCY, TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									  "   TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, ACCOUNTINGPERIOD, ACCOUNTINGYEAR, LASTUPDATE, LASTTIME ) " & _
									  "  values ( '" & _
										rsSAPFormated ("DOCUMENT_DATE") & "','" & _
										rsSAPFormated ("DOCUMENT_TYPE") & "','" & _
										rsSAPFormated ("COMPANY_CODE") & "','" & _
										rsSAPFormated ("POSTING_DATE") & "','" & _
										rsSAPFormated ("PERIOD") & "','" & _
										rsSAPFormated ("CURRENCY") & "','" & _
										rsSAPFormated ("REFERENCE") & "','" & _
										rsSAPFormated ("HEADER_TEXT") & "','" & _
										rsSAPFormated ("POSTING_KEY") & "','" & _
										rsSAPFormated ("BUSINESS_AREA") & "','" & _
										rsSAPFormated ("GL_ACCOUNT") & "','" & _
										rsSAPFormated ("COST_CENTER") & "','" & _
										rsSAPFormated ("INTERNAL_ORDER") & "','" & _
										rsSAPFormated ("DOCUMENT_CCY") & "'," & _
										"to_date('" & rsSAPFormated ("TDATE") & "','MM/DD/YYYY'),'" & _
										rsSAPFormated ("JOBCODE") & "','" & _
										rsSAPFormated ("SOURCE") & "','" & _
										rsSAPFormated ("SOURCETYPE") & "','" & _
										rsSAPFormated ("SOURCECODE") & "','" & _
										rsSAPFormated ("SOURCECODE_STATUS") & "','" & _
										rsSAPFormated ("TARGETTYPE") & "','" & _
										rsSAPFormated ("TARGETCODE") & "','" & _
										rsSAPFormated ("TARGETCODE_STATUS") & "'||'-NAFIS01','" & _
										rsSAPFormated ("REFERENCENO") & "','" & _
										rsSAPFormated ("QUANTITY") & "','" & _
										rsSAPFormated ("OUTPUT") & "','" & _
										rsSAPFormated ("AMOUNT") & "','" & _
										rsSAPFormated ("FCBA") & "'," & _
										"to_date('" & rsSAPFormated ("CLOSING_DATE") & "','MM/DD/YYYY')," & _
										"sysdate ," & _
										rsSAPFormated ("ACCOUNTINGPERIOD") & "," & _
										rsSAPFormated ("ACCOUNTINGYEAR") & "," & _
										"sysdate," & _
										"to_char(sysdate, 'HH24:MI')) "

						'response.write qInsertData
						connStr.execute(qInsertData)
					else

						temp_company_code = temp_Field_Info_01
						
						if rsSAPFormated ("BUSINESS_AREA") <> "" then 
							temp_company_code_01 = temp_company_code
						else
							temp_company_code_01 = ""
						end if
						
						if rsSAPFormated ("COST_CENTER") <> "" then 
							temp_company_code_02 = temp_company_code
						else
							temp_company_code_02 = ""
						end if
						
						if rsSAPFormated ("INTERNAL_ORDER") <> "" then 
							temp_company_code_03 = temp_company_code
						else
							temp_company_code_03 = ""
						end if
						


						
						' First Record, change to other Company with Original Account
						qInsertData = " insert into erp_report_sap_generated " & _
									  " ( DOCUMENT_DATE, DOCUMENT_TYPE, COMPANY_CODE, POSTING_DATE, PERIOD, " & _
									  "   CURRENCY, REFERENCE, HEADER_TEXT, POSTING_KEY, BUSINESS_AREA, GL_ACCOUNT, " & _
									  "   COST_CENTER, INTERNAL_ORDER, DOCUMENT_CCY, TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									  "   TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE,ACCOUNTINGPERIOD, ACCOUNTINGYEAR,  LASTUPDATE, LASTTIME ) " & _
									  "  values ( '" & _
										rsSAPFormated ("DOCUMENT_DATE") & "','" & _
										rsSAPFormated ("DOCUMENT_TYPE") & "','" & _
										temp_company_code & "','" & _
										rsSAPFormated ("POSTING_DATE") & "','" & _
										rsSAPFormated ("PERIOD") & "','" & _
										rsSAPFormated ("CURRENCY") & "','" & _
										rsSAPFormated ("REFERENCE") & "','" & _
										rsSAPFormated ("HEADER_TEXT") & "','" & _
										rsSAPFormated("POSTING_KEY") & "','" & _
										temp_company_code_01 & "'||substr('" & rsSAPFormated ("BUSINESS_AREA") & "',length('"& temp_company_code_01 & "') + 1, length('" & rsSAPFormated ("BUSINESS_AREA")  & "') - length('" & temp_company_code_01 & "')),'" & _
										rsSAPFormated ("GL_ACCOUNT") & "','" & _
										temp_company_code_02 & "'||substr('" & rsSAPFormated ("COST_CENTER") & "',length('"& temp_company_code_02 & "') + 1, length('" & rsSAPFormated ("COST_CENTER") &  "') - length('" & temp_company_code_02 & "')),'" & _
										temp_company_code_03 & "'||substr('" & rsSAPFormated ("INTERNAL_ORDER") & "',length('"& temp_company_code_03 & "') + 1, length('" & rsSAPFormated ("INTERNAL_ORDER") &  "') - length('" & temp_company_code_03 & "')),'" & _
										rsSAPFormated ("DOCUMENT_CCY") & "'," & _
										"to_date('" & rsSAPFormated ("TDATE") & "','MM/DD/YYYY'),'" & _
										rsSAPFormated ("JOBCODE") & "','" & _
										rsSAPFormated ("SOURCE") & "','" & _
										rsSAPFormated ("SOURCETYPE") & "','" & _
										rsSAPFormated ("SOURCECODE") & "','" & _
										rsSAPFormated ("SOURCECODE_STATUS") & "','" & _
										rsSAPFormated ("TARGETTYPE") & "','" & _
										rsSAPFormated ("TARGETCODE") & "','" & _
										rsSAPFormated ("TARGETCODE_STATUS") & "'||'-NAFIS02','" & _
										rsSAPFormated ("REFERENCENO") & "','" & _
										rsSAPFormated ("QUANTITY") & "','" & _
										rsSAPFormated ("OUTPUT") & "','" & _
										rsSAPFormated ("AMOUNT") & "','" & _
										rsSAPFormated ("FCBA") & "'," & _
										"to_date('" & rsSAPFormated ("CLOSING_DATE") & "','MM/DD/YYYY')," & _
										"sysdate," & _
										rsSAPFormated ("ACCOUNTINGPERIOD") & "," & _
										rsSAPFormated ("ACCOUNTINGYEAR") & "," & _
										"sysdate," & _
										"to_char(sysdate, 'HH24:MI')) "

						connStr.execute(qInsertData)
						
						'Change Posting Key Debet -> Credit or Credit to Debet
						if rsSAPFormated("POSTING_KEY") = "40" then 
							temp_Posting_Key = "50"
						else
							temp_Posting_Key = "40"
						end if	

						
						' Second Record, change to other Company with INTRACO Jobcode/Account
						qInsertData = " insert into erp_report_sap_generated " & _
									  " ( DOCUMENT_DATE, DOCUMENT_TYPE, COMPANY_CODE, POSTING_DATE, PERIOD, " & _
									  "   CURRENCY, REFERENCE, HEADER_TEXT, POSTING_KEY, BUSINESS_AREA, GL_ACCOUNT, " & _
									  "   COST_CENTER, INTERNAL_ORDER, DOCUMENT_CCY, TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									  "   TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, ACCOUNTINGPERIOD, ACCOUNTINGYEAR,  LASTUPDATE, LASTTIME ) " & _
									  "  values ( '" & _
										rsSAPFormated ("DOCUMENT_DATE") & "','" & _
										rsSAPFormated ("DOCUMENT_TYPE") & "','" & _
										temp_company_code & "','" & _
										rsSAPFormated ("POSTING_DATE") & "','" & _
										rsSAPFormated ("PERIOD") & "','" & _
										rsSAPFormated ("CURRENCY") & "','" & _
										rsSAPFormated ("REFERENCE") & "','" & _
										rsSAPFormated ("HEADER_TEXT") & "','" & _
										temp_Posting_Key & "','" & _
										temp_company_code_01 & "'||substr('" & rsSAPFormated ("BUSINESS_AREA") & "',length('"& temp_company_code_01 & "') + 1, length('" & rsSAPFormated ("BUSINESS_AREA") & "') - length('" & temp_company_code_01 & "')),"& _
										"ERP_REPORT_API.GetGLAccount ('" & rsSAPFormated ("FCBA") & "','" & tempIntracoAccount & "','" & rsSAPFormated ("SOURCETYPE") & "'),'" & _
										temp_company_code_02 & "'||substr('" & rsSAPFormated ("COST_CENTER") & "',length('"& temp_company_code_02 & "') + 1, length('" & rsSAPFormated ("COST_CENTER") & "') - length('" & temp_company_code_02 & "')),'" & _
										temp_company_code_03 & "'||substr('" & rsSAPFormated ("INTERNAL_ORDER") & "',length('"& temp_company_code_03 & "') + 1, length('" & rsSAPFormated ("INTERNAL_ORDER") & "') - length('" & temp_company_code_03 & "')),'" & _
										rsSAPFormated ("DOCUMENT_CCY") & "'," & _
										"to_date('" & rsSAPFormated ("TDATE") & "','MM/DD/YYYY'),'" & _
										tempIntracoAccount & "','" & _
										rsSAPFormated ("SOURCE") & "','" & _
										rsSAPFormated ("SOURCETYPE") & "','" & _
										rsSAPFormated ("SOURCECODE") & "','" & _
										rsSAPFormated ("SOURCECODE_STATUS") & "','" & _
										rsSAPFormated ("TARGETTYPE") & "','" & _
										rsSAPFormated ("TARGETCODE") & "','" & _
										rsSAPFormated ("TARGETCODE_STATUS") & "'||'-NAFIS03','" & _
										rsSAPFormated ("REFERENCENO") & "','" & _
										rsSAPFormated ("QUANTITY") & "','" & _
										rsSAPFormated ("OUTPUT") & "','" & _
										-1 * rsSAPFormated ("AMOUNT") & "','" & _
										rsSAPFormated ("FCBA") & "'," & _
										"to_date('" & rsSAPFormated ("CLOSING_DATE") & "','MM/DD/YYYY')," & _
										"sysdate," & _
										rsSAPFormated ("ACCOUNTINGPERIOD") & "," & _
										rsSAPFormated ("ACCOUNTINGYEAR") & "," & _
										"sysdate," & _
										"to_char(sysdate, 'HH24:MI')) "
						connStr.execute(qInsertData)
						
						'Change Posting Key Debet -> Credit or Credit to Debet
						if  temp_Posting_Key = "40" then 
							temp_Posting_Key = "50"
						else
							temp_Posting_Key = "40"
						end if					
						
						' Third Record, Original Company with INTRACO Jobcode/Account
						qInsertData = " insert into erp_report_sap_generated " & _
									  " ( DOCUMENT_DATE, DOCUMENT_TYPE, COMPANY_CODE, POSTING_DATE, PERIOD, " & _
									  "   CURRENCY, REFERENCE, HEADER_TEXT, POSTING_KEY, BUSINESS_AREA, GL_ACCOUNT, " & _
									  "   COST_CENTER, INTERNAL_ORDER, DOCUMENT_CCY, TDATE, JOBCODE, SOURCE, SOURCETYPE, SOURCECODE, SOURCECODE_STATUS, " & _
									  "   TARGETTYPE, TARGETCODE, TARGETCODE_STATUS, REFERENCENO, QUANTITY, OUTPUT, AMOUNT, FCBA, CLOSING_DATE, PROCESS_DATE, ACCOUNTINGPERIOD, ACCOUNTINGYEAR,  LASTUPDATE, LASTTIME ) " & _
									  "  values ( '" & _
										rsSAPFormated ("DOCUMENT_DATE") & "','" & _
										rsSAPFormated ("DOCUMENT_TYPE") & "','" & _
										rsSAPFormated ("COMPANY_CODE") & "','" & _
										rsSAPFormated ("POSTING_DATE") & "','" & _
										rsSAPFormated ("PERIOD") & "','" & _
										rsSAPFormated ("CURRENCY") & "','" & _
										rsSAPFormated ("REFERENCE") & "','" & _
										rsSAPFormated ("HEADER_TEXT") & "','" & _
										temp_Posting_Key & "','" & _
										rsSAPFormated ("BUSINESS_AREA") & "'," & _
										"ERP_REPORT_API.GetGLAccount ('" & rsSAPFormated ("FCBA") & "','" & tempIntracoAccount & "','" & rsSAPFormated ("SOURCETYPE") & "'),'" & _
										rsSAPFormated ("COST_CENTER") & "','" & _
										rsSAPFormated ("INTERNAL_ORDER") & "','" & _
										rsSAPFormated ("DOCUMENT_CCY") & "'," & _
										"to_date('" & rsSAPFormated ("TDATE") & "','MM/DD/YYYY'),'" & _
										tempIntracoAccount & "','" & _
										rsSAPFormated ("SOURCE") & "','" & _
										rsSAPFormated ("SOURCETYPE") & "','" & _
										rsSAPFormated ("SOURCECODE") & "','" & _
										rsSAPFormated ("SOURCECODE_STATUS") & "','" & _
										rsSAPFormated ("TARGETTYPE") & "','" & _
										rsSAPFormated ("TARGETCODE") & "','" & _
										rsSAPFormated ("TARGETCODE_STATUS") & "'||'-NAFIS04','" & _
										rsSAPFormated ("REFERENCENO") & "','" & _
										rsSAPFormated ("QUANTITY") & "','" & _
										rsSAPFormated ("OUTPUT") & "','" & _
										rsSAPFormated ("AMOUNT") & "','" & _
										rsSAPFormated ("FCBA") & "'," & _
										"to_date('" & rsSAPFormated ("CLOSING_DATE") & "','MM/DD/YYYY')," & _
										"sysdate," & _
										rsSAPFormated ("ACCOUNTINGPERIOD") & "," & _
										rsSAPFormated ("ACCOUNTINGYEAR") & "," & _
										"sysdate," & _
										"to_char(sysdate, 'HH24:MI')) "
						
						connStr.execute(qInsertData)
					
					end if
				
					rsSAPFormated.MoveNext
				loop
				rsSAPFormated.close
				set rsSAPFormated = Nothing
			end if
			Response.Write "(DONE)" & "<br>"
		else
			Response.Write "(ERROR)" & "<br>"
		end if	
	else
		qInsertData = " insert into erp_report_sap_generated " & _
							  " select * from erp_report_sap " & _
							  " where " & _
							  " fcba = '" & Session("fcBA") & "' "  & _
							  " and ACCOUNTINGPERIOD = " & v_Bulan  &  _
							  " and ACCOUNTINGYEAR = " & v_Tahun
								  



		connStr.execute(qInsertData)
		Response.Write "(DONE)" & "<br>"
	end if

	Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Report FICO is successfully...." & "<br><br></font>"


	if doContinueProcess = "TRUE" then 
		' Second we go for Detail SAP - SD 'Response.write "NAF001" accountingperiod,accountingyear 
		Response.Write "Start Normalization for Sales and Distribution Data 01 ..., Year :" & v_Tahun & ", month : " &  v_Bulan   

		strSQLNormalization = " delete erp_report_sap_sd_generated " & _
							  " where " & _
							  " fcba = '" & Session("fcBA") & "' "  & _
							  " and to_number(to_Char( to_date(delivery_date,'dd.mm.yyyy'),'mm'))= " & v_Bulan  &  _
							  " and to_number(to_Char( to_date(delivery_date,'dd.mm.yyyy'),'yyyy'))= " & v_Tahun

							  
		connStr.execute(strSQLNormalization)
		Response.Write "(DONE)" & "<br>"

		Response.Write "Start Generate for Sales and Distribution Data 01..., Year :" & v_Tahun & ", month : " &  v_Bulan  

		qInsertData = " insert into erp_report_sap_sd_generated " & _
							  " select * from erp_report_sap_sd " & _
							  " where " & _
							  " fcba = '" & Session("fcBA") & "' "  & _
							  " and to_number(to_Char( to_date(delivery_date,'dd.mm.yyyy'),'mm'))= " & v_Bulan  &  _
							  " and to_number(to_Char( to_date(delivery_date,'dd.mm.yyyy'),'yyyy'))= " & v_Tahun


		connStr.execute(qInsertData)

		Response.Write "(DONE)" & "<br>"


		' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

		Response.Write "Start Normalization for Sales and Distribution Data 02... , Year :" & v_Tahun & ", month : " &  v_Bulan  

		strSQLNormalization = " delete erp_report_sap_sd01_generated " & _
							  " where " & _
							  " fcba = '" & Session("fcBA") & "' "  & _
							  " and to_number(to_Char( to_date(post_date,'dd.mm.yyyy'),'mm'))= " & v_Bulan  &  _
							  " and to_number(to_Char( to_date(post_date,'dd.mm.yyyy'),'yyyy'))= " & v_Tahun

							  
		connStr.execute(strSQLNormalization)
		Response.Write "(DONE)" & "<br>"

		Response.Write "Start Generate for Sales and Distribution Data 02..., Year :" & v_Tahun & ", month : " &  v_Bulan   

		qInsertData = " insert into erp_report_sap_sd01_generated " & _
							  " select * from erp_report_sap_sd01 " & _
							  " where " & _
							  " fcba = '" & Session("fcBA") & "' "  & _
							  " and to_number(to_Char( to_date(post_date,'dd.mm.yyyy'),'mm'))= " & v_Bulan  &  _
							  " and to_number(to_Char( to_date(post_date,'dd.mm.yyyy'),'yyyy'))= " & v_Tahun 


		connStr.execute(qInsertData)

		Response.Write "(DONE)" & "<br>"

		' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

		Response.Write "Start Normalization for Sales and Distribution Data 03... ,  Year :" & v_Tahun & ", month : " &  v_Bulan  

		strSQLNormalization = " delete erp_report_sap_sd02_generated " & _
							  " where " & _
							  " fcba = '" & Session("fcBA") & "' "  & _
							  " and  to_number(to_Char( delivery_date,'MM'))= " & v_Bulan  & "" & _
							  " and  to_number(to_Char( delivery_date,'YYYY'))= " & v_Tahun

		'Response.Write 	strSQLNormalization				  
		connStr.execute(strSQLNormalization)
		Response.Write "(DONE)" & "<br>"

		Response.Write "Start Generate for Sales and Distribution Data 02...,  Year :" & v_Tahun & ", month : " &  v_Bulan  

		qInsertData = " insert into erp_report_sap_sd02_generated " & _
							  " select * from erp_report_sap_sd02 " & _
							  " where " & _
							  " fcba = '" & Session("fcBA") & "' "  & _
							  " and to_Char(delivery_date,'mm')= " & v_Bulan  &  _
							  " and to_Char( delivery_date,'yyyy')= " & v_Tahun

		'Response.Write qInsertData
		connStr.execute(qInsertData)

		Response.Write "(DONE)" & "<br>"
		
		Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Report SD is successfully...." & "<br><br></font>"

	end if

	%>

	<%if err.number <> 0 or doContinueProcess = "FALSE" then%>
		<br><img src="images/error.gif"><%
		response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
		cmdView	= "True"%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>	
	<%Else
		
			if ScriptExecuted <> "ByVariable" then
				Session("CurP_Costbook") = cmdFindAccountPeriod("Costbook")
				Session("ClsP_Costbook") = cmdAccountPeriodClose("Costbook")
			end if
		
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "<br>" & "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."	

	'--Added by Ukie at 25-Oct-2012--	
	Endingtime = FormatDateTime(CDate(Now))

	qLineMax = "select nvl(max(lineno),0) tempMaxLine from CLOSINGLOG where fcba = '" & session("fcBA") & "' AND SYSTEM = 'ERPReport'" 
								
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1

	strClosingLog = " INSERT INTO CLOSINGLOG(SYSTEM, FCBA, STARTINGTIME, ENDINGTIME, LINENO, LASTUPDATE, LASTTIME, fcentry) VALUES (" & _
					" 'ERPReport', " & _
					" '" & Session("fcBA") & "'," & _
					" TO_DATE('" & Startingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" TO_DATE('" & Endingtime & "', 'mm/dd/yyyy hh:mi:ssam')," & _
					" " & tempMaxLine & "" & _
					" , SYSDATE, TO_CHAR(SYSDATE, 'HH24:MI')" & _
					" ,'" & Session("CodeName") & "')" & _ 
	'connStr.execute(strClosingLog)

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
					" WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'ERPReport'"
				
	Set rsLenghtime = Server.CreateObject("ADODB.Recordset")
	rsLenghtime.CursorType = 3
	rsLenghtime.CursorLocation = 3
	rsLenghtime.Open strLenghtime, connstr

	' upClosingLog = " UPDATE CLOSINGLOG SET" & _
					' " LENGTHTIME = '" & rsLenghtime(5) & " Days " & rsLenghtime(6) & " Hours " & rsLenghtime(7) & " Minutes " & rsLenghtime(8) & " Seconds " & "'" & _
					' " WHERE FCBA = '" & Session("fcBA") & "' AND LINENO = " & tempMaxLine & " AND SYSTEM = 'ERPReport'"		
	'connStr.execute(upClosingLog)
	'--End Added--	
		
	%>
		<br>
		<img src="images/ok.gif"><font color=Blue pointsize=9 face=Verdana>CONRATULATIONS !!!, ERP Report Data have been succesfully generated, Thank You....
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
	<%End if%>
<%
'else
%>
	<!--
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<br><br><img src="images/error.gif">ERP Report is NOT completed...
	-->
	<%
    'tempErrorWaterFlow = "Please Check Previous Closing Process, its still not DONE yet"
    'response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempErrorWaterFlow & "</font>"
	%>
	<!--
	<script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>	
	<img src="images/error.gif" />
	-->
<%
'end if
%>	

</body>
</html>