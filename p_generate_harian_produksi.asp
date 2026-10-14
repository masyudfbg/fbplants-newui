<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<% Dim textReportName : textReportName = "Laporan Harian Restan dan Produksi"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>ePlaS - Generate <%=PCase(textReportName)%></title>
	<link rel="StyleSheet" href="style-content.css" type="text/css" />
<style>
* {font-family:Verdana,Arial,sans-serif;font-size:11px;}
</style>
</head>
<body>
	<div id="status_running">
		<img src="/icon/working.gif">
	</div>

<%
Dim TimeAwal, TimeAkhir, Hitung, Lama, nTotalDay
Dim oRecordset, oQuery
TimeAwal = Time()

Dim DATE_FROM : DATE_FROM = Request("DATE_FROM")
Dim DATE_TO : DATE_TO = Request("DATE_TO")
Dim PeriodeBuku : PeriodeBuku = session("v_Costbook")

Dim nIntervalDay : nIntervalDay = 1
Dim nSelisihBudgetYear : nSelisihBudgetYear = 1
Dim nSelisihDateFrom : nSelisihDateFrom = 1
Dim nSelisihDateTo : nSelisihDateTo = 1

Dim val_currentdate, val_datefrom, val_dateto
Dim val_next_currentdate, val_next_datefrom, val_next_dateto


'***** For temporary purposes we make the hardcode *****'
Dim budget_year : budget_year = 1900
Dim budget_version : budget_version = 1
Dim module_version : module_version = 1
'*******************************************************'
%>
	<br/><!--Enter-->
	Generating Data From <%=FormatDateTime(DATE_FROM,1)%>
	To <%=FormatDateTime(DATE_TO,1)%>
	<br/><!--Enter-->
	<br/><!--Enter-->
	Please Wait...
	<br/><!--Enter-->
<%Response.Write "Start ..." & TimeAwal & "<br/><!--Enter-->"%>
<%
connstr.BeginTrans

'********** GET INTERVAL DAY **********'
oQuery = " Select Parameter_Api.GetRate_1('DIR','DAYINTERVALREPORT', '" & Session("FCBA") & "') IntervalDay from dual"
		 
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open oQuery, connstr

If Not oRecordset.eof Then
	nIntervalDay = Cdbl(oRecordset("IntervalDay"))
Else
	nIntervalDay = 7
End If

oQuery = ""
oRecordset.Close
set oRecordset = Nothing
'********** END GET INTERVAL DAY **********'

'********** GET TOTAL DAY 1 to 7 **********'
oQuery = " SELECT to_date('" & DATE_TO & "', 'mm/dd/yyyy')" & _
         "        - to_date('" & DATE_FROM & "', 'mm/dd/yyyy') + 1 totalday from dual"
		 
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open oQuery, connstr

If Not oRecordset.eof Then
	nTotalDay = Cdbl(oRecordset("totalday"))
Else
	nTotalDay = 0
End If

oQuery = ""
oRecordset.Close
set oRecordset = Nothing
'********** END GET TOTAL DAY 1 to 7 **********'


'********** GET SELISIH MONTH FOR BUDGET YEAR **********'
oQuery = " SELECT to_number(to_number(to_char(to_date('" & DATE_TO & "', 'mm/dd/yyyy'), 'yyyy'))" & _
         "        - to_number(to_char(to_date('" & DATE_FROM & "', 'mm/dd/yyyy'), 'yyyy'))) BudgetYear from dual"
		 
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open oQuery, connstr

If Not oRecordset.eof Then
	nSelisihBudgetYear = Cdbl(oRecordset("BudgetYear"))
Else
	nSelisihBudgetYear = 1
End If

oQuery = ""
oRecordset.Close
set oRecordset = Nothing
'********** END SELISIH GET MONTH FOR BUDGET YEAR **********'

'********** GET MONTH FOR BUDGET YEAR **********'
oQuery = " SELECT to_number(to_char(to_date('" & DATE_TO & "', 'mm/dd/yyyy'), 'yyyy')) BudgetYear from dual"
		 
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open oQuery, connstr

If Not oRecordset.eof Then
	budget_year = Cdbl(oRecordset("BudgetYear"))
Else
	budget_year = 1980
End If

oQuery = ""
oRecordset.Close
set oRecordset = Nothing
'********** END GET MONTH FOR BUDGET YEAR **********'

'********** GET SELISIH TANGGAL DENGAN CURRENTDATE **********'
oQuery =    " select to_number(to_char(to_date('" & PeriodeBuku & "', 'mm/dd/yyyy'), 'yyyymm')) currentdate" & _
			"        , to_number(to_char(to_date('" & DATE_FROM & "', 'mm/dd/yyyy'), 'yyyymm')) datefrom" & _
			"        , to_number(to_char(to_date('" & DATE_TO & "', 'mm/dd/yyyy'), 'yyyymm')) dateto" & _
			"        , to_number(to_char(add_months(to_date('" & PeriodeBuku & "', 'mm/dd/yyyy'), 1), 'yyyymm')) next_currentdate" & _
			" from dual"

		 
Set oRecordset = Server.CreateObject("ADODB.Recordset")
oRecordset.Open oQuery, connstr

If Not oRecordset.eof Then
	val_currentdate = Cdbl(oRecordset("currentdate"))
	val_datefrom = Cdbl(oRecordset("datefrom"))
	val_dateto = Cdbl(oRecordset("dateto"))
	val_next_currentdate = Cdbl(oRecordset("next_currentdate"))
Else
	val_currentdate = 205011
	val_datefrom = 205011
	val_dateto = 205011
	val_next_currentdate = 205012
End If

oQuery = ""
oRecordset.Close
set oRecordset = Nothing
'********** END GET SELISIH TANGGAL DENGAN CURRENTDATE **********'

If nSelisihBudgetYear = 0 Then
		'If nTotalDay > nIntervalDay Or nTotalDay < 1 Then
		If nTotalDay > 7 Or nTotalDay < 1 Then
			err.number = 1001
			'err.description = "<br/>Tanggal tidak boleh lebih/ kurang dari " & nIntervalDay & " hari<br/>"
			err.description = "<br/>Tanggal tidak boleh lebih dari 7 hari<br/>"
		Else
			' Restan
			oQuery	= 	" BEGIN HARVESTING_RESTAN_REP_API.INSERTINTOTABLE(" & _
						"                                    '" & Session("FCBA") & "'" & _
						"                                    , to_date('" & DATE_FROM & "', 'mm/dd/yyyy')" & _
						"                                    , to_date('" & DATE_TO & "', 'mm/dd/yyyy')" & _
						" ); END; "
						
			connStr.execute(oQuery)
			
			' Produksi
			oQuery	= 	" BEGIN GENERATE_DAILY_REPORT_API.PRODUKSI_PROCESS(" & _
						"                                    '" & Session("FCBA") & "'" & _
						"                                    , to_date('" & DATE_FROM & "', 'mm/dd/yyyy')" & _
						"                                    , to_date('" & DATE_TO & "', 'mm/dd/yyyy')" & _
						"                                    , " & budget_year & "" & _
						"                                    , " & budget_version & "" & _
						"                                    , " & module_version & "" & _
						" ); END; "

			connStr.execute(oQuery)
			
			' Produksi di Laporan Harian Kebun
			' 1 Normalisasi
			oQuery	= 	" BEGIN PRODUCTION_SUMMARY_REP_API.Table_Normalization_StepOnly(" & _
						"                                    '" & Session("FCBA") & "'" & _
						"                                    , to_date('" & DATE_FROM & "', 'mm/dd/yyyy')" & _
						"                                    , to_date('" & DATE_TO & "', 'mm/dd/yyyy')" & _
						"                                    , 'STEP_02'" & _
						" ); END; "

			connStr.execute(oQuery)
			
			' 2 Insert Data
			oQuery	= 	" BEGIN PRODUCTION_SUMMARY_REP_API.Production_Process(" & _
						"                                    '" & Session("FCBA") & "'" & _
						"                                    , to_date('" & DATE_FROM & "', 'mm/dd/yyyy')" & _
						"                                    , to_date('" & DATE_TO & "', 'mm/dd/yyyy')" & _
						"                                    , " & budget_year & "" & _
						"                                    , " & budget_version & "" & _
						"                                    , " & module_version & "" & _
						"                                    , 'STEP_02'" & _
						" ); END; "

			connStr.execute(oQuery)
		End If
Else
	err.number = 1001
	err.description = "<br/>Tahun Tidak Sama !! antara 2 tanggal yang di pilih<br/>"
End If

If err.number = 0 Then
	connStr.CommitTrans
End If
If err.number <> 0 Then
	connStr.RollbackTrans
End If
%>

<%If err.number = 0 Then%>
	<br/><!--Enter-->
<%
TimeAkhir = Time()
Lama = TimeAkhir - TimeAwal
Response.Write "Finish ... " & TimeAkhir & "<br/><!--Enter-->"
Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds.<br/><!--Enter-->"
%>
	<br/><img src="/images/ok.gif" /> <font color="green"><strong>Generating Data Succesfully !!!</strong></font>
<%End If%>

<%If err.number <> 0 then%>
<%	
	response.Write  "<font color=""red"">" & err.description & "</font><br/><!--Enter-->"
%>
<%
TimeAkhir = Time()
Lama = TimeAkhir - TimeAwal
Response.Write " Finish ... " & TimeAkhir & "<br/><!--Enter-->"
Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds.<br/><!--Enter-->"
%>
	<br/><img src="/images/error.gif" /> <font color="red"><strong>Generating Data Failed !!!</strong></font>
<%End If%>

</body>
</html>