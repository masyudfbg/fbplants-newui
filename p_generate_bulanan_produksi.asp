<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<% Dim textReportName : textReportName = "Laporan Bulanan Produksi"%>
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

Dim month_from : month_from = ""
Dim month_until : month_until = ""


'***** For temporary purposes we make the hardcode *****'
Dim budget_year : budget_year = 1900
Dim budget_version : budget_version = 1
Dim module_version : module_version = 1
'*******************************************************'


month_from = Year(Session("v_Costbook")) & Lpad(Month(Session("v_Costbook")), "0", 2)
month_until = Year(Session("v_Costbook")) & Lpad(Month(Session("v_Costbook")), "0", 2)
budget_year = Year(Session("v_Costbook"))

%>
	<br/><!--Enter-->
	Generating Data From <%=month_from%>
	To <%=month_until%>
	<br/><!--Enter-->
	<br/><!--Enter-->
	Please Wait...
	<br/><!--Enter-->
<%Response.Write "Start ..." & TimeAwal & "<br/><!--Enter-->"%>
<%
connstr.BeginTrans

' Bulanan Produksi
oQuery	= 	" BEGIN GENERATE_MONTHLY_REPORT_API.InsertIntoTable(" & _
			"                                    '" & Session("FCBA") & "'" & _
			"                                    , '" & month_from & "'" & _
			"                                    , '" & month_until & "'" & _
			"                                    , " & budget_year & "" & _
			"                                    , " & budget_version & "" & _
			"                                    , " & module_version & "" & _
			" ); END; "
			
'Response.Write oQuery & "<br/>"

connStr.execute(oQuery)

' Produksi di Laporan Bulanan Kebun
' 1 Normalisasi
oQuery	= 	" BEGIN PRODMONTH_SUMMARY_REP_API.Table_Normalization_StepOnly(" & _
			"                                    '" & Session("FCBA") & "'" & _
			"                                    , '" & month_from & "'" & _
			"                                    , '" & month_until & "'" & _
			"                                    , 'STEP_02'" & _
			" ); END; "
			
'Response.Write oQuery & "<br/>"

connStr.execute(oQuery)

' 2 Insert Data
oQuery	= 	" BEGIN PRODMONTH_SUMMARY_REP_API.Production_Process(" & _
			"                                    '" & Session("FCBA") & "'" & _
			"                                    , '" & month_from & "'" & _
			"                                    , '" & month_until & "'" & _
			"                                    , " & budget_year & "" & _
			"                                    , " & budget_version & "" & _
			"                                    , " & module_version & "" & _
			"                                    , 'STEP_02'" & _
			" ); END; "
			
'Response.Write oQuery & "<br/>"

connStr.execute(oQuery)
			


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