<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim textReportName : textReportName = "Calendar Stats"
Dim ScriptName : ScriptName = "p_calendar_stat.asp"
Dim stringStats,objRSArray,arrRS,i,ia
Dim q_dateToDistict,arrRSHeader,objRSArrayHeader
dim bg,color1,color2

Dim NPY,AI,BNIY,Tru,Fls,color_btn1,color_btn2

ScriptName =  ScriptName

	Function GetRSArrayHeader(strSQL)
			'Response.Write(strSQL)
			set objRSArrayHeader = connStr.Execute(strSQL)
			If Not objRSArrayHeader.EOF Then
				arrRSHeader = objRSArrayHeader.GetRows()
				GetRSArrayHeader = arrRSHeader
			Else
                Exit Function
			End If
			objRSArrayHeader.Close()
			Set objRSArrayHeader = Nothing
	End Function
	
	Function GetRSArray(strSQL)
			'Response.Write(strSQL)
			set objRSArray = connStr.Execute(strSQL)
			If Not objRSArray.EOF Then
				arrRS = objRSArray.GetRows()
				GetRSArray = arrRS
			Else
                Exit Function
			End If
			objRSArray.Close()
			Set objRSArray = Nothing
	End Function
		
%>

<html>
	<head>
<title>ePlaS - <%=PCase(textReportName)%></title>
<link rel="StyleSheet" href="/css/stylereport2.css" type="text/css" />
<style>
	#calendar {
		max-width: 99%;
		margin: 0 auto;
	}
	.fc-title{font-size:110%;}
	.fc-time{font-size:110%;}
	.fc-unthemed thead th {
		background:#00A300;
		font-size:14px;
		color:#FFF;
		padding:3px;
	}
	table{margin-bottom:0px;}
</style>
<link href='../calendar/css/fullcalendar.min.css' rel='stylesheet' />
<link href='../calendar/css/fullcalendar.print.css' rel='stylesheet' media='print' />
<link rel="stylesheet" href="../css/jquery-ui.css" />

<script src='../calendar/js/moment.min.js'></script>
<script src='../calendar/js/jquery.min.js'></script>
<script src='../calendar/js/fullcalendar.js'></script>
<script src="../js/jquery-ui.js"></script>
<script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<%

dim date_now : date_now = set_format_tanggal(textDate(date()),"YYYY-MM-DD")
dim date_lusa : date_lusa = set_format_tanggal(textDate(date()-2),"YYYY-MM-DD")
dim date_show : date_show = request("txtdate") 
'response.write(date_show)
if date_show = "" then
	date_show = date_now
else
	date_show = set_format_tanggal(textDate(date_show),"YYYY-MM-DD") 
end if
'response.write(date_show)
stringStats = 	" select 'exec' exec_sign, trunc(t.fddate)fddate,cases ,count(*) total from summary_exec_daily_production t" & _
				" where to_char(t.fddate,'YYYYMM') =  to_char(to_date('" & date_show & "','YYYY-MM-DD'),'YYYYMM') " & _
				" group by  trunc(t.fddate),cases" & _
				" union" & _ 
				" select 'input' exec_sign, trunc(t.fddate)fddate,isdailyclosing cases ,count(*) total  from dailyclosingactivity t" & _
				" where to_char(t.fddate,'YYYYMM') =  to_char(to_date('" & date_show & "','YYYY-MM-DD'),'YYYYMM') and businessunit_api.Isactive(fcba)= 'TRUE' " & _
				" group by  trunc(t.fddate),isdailyclosing" & _
				" order by fddate DESC,cases"
q_dateToDistict = "select distinct (fddate) from (" & stringStats & ") order by fddate DESC"
arrRSHeader = GetRSArrayHeader(q_dateToDistict)
arrRS = GetRSArray(stringStats)

%>
<script>
$(document).ready(function() {
		$('#calendar').fullCalendar({
			header: {
				left : '',
				center: 'title',
				right: '',
			},
			defaultDate: '<%=date_show%>',
			editable: false,
			eventLimit: false, // allow "more" link when too many events
			events: [
			<% If IsArray(arrRSHeader) Then
					For i = LBound(arrRSHeader, 2) To UBound(arrRSHeader, 2) 
							NPY = "0"
							AI = "0"
							BNIY = "0"
							Tru	= "0"
							Fls	= "0"
						If IsArray(arrRS) Then
							For ia = LBound(arrRS, 2) To UBound(arrRS, 2)
							'response.write("<script>alert('" & arrRS(1, ia) & "');</script>")
								If set_format_tanggal(arrRS(1, ia),"YYYY-MM-DD") = set_format_tanggal(arrRSHeader(0, i),"YYYY-MM-DD") then
									if arrRS(2, ia) = "NOT PROCESSED YET" then
										NPY = arrRS(3, ia)
									elseif arrRS(2, ia) = "PROCESSED - ALREADY INPUTED" then
										AI = arrRS(3, ia)
									elseif arrRS(2, ia) = "PROCESSED - BUT NOT INPUTED YET" then
										BNIY = arrRS(3, ia)
									elseif arrRS(2, ia) = "TRUE" then
										Tru = arrRS(3, ia)
									elseif arrRS(2, ia) = "FALSE" then
										Fls = arrRS(3, ia)
									end if
								end if
							Next
						End If
						
						if textDate(arrRSHeader(0, i)) >= textDate(date_lusa) then
							'bg = "{start: '" & date_lusa & "',end: '" & date_now & "',rendering: 'background',color: '#FFED96'},"
							bg = ""
							color1 = "#01A9DB"
							color2 = "#F22935"
							color_btn1 = ""
							color_btn2 = ""							
						else
							bg = ""
							color1 = "#F7819F"
							color2 = "#A9D0F5"
							color_btn1 = "background-color:Green;"
							color_btn2 = "background-color:Green;"
						end if
						%>			
						{title: 'Daily Input Closing',txt_data: 'Closing: <%=Tru%> | Not Yet: <%=Fls%>',start: '<%=set_format_tanggal(textDate(arrRSHeader(0, i)),"YYYY-MM-DD")%>', url : 'p_monitoringDailyInputClosing.asp?v_txtdate=<%=arrRSHeader(0, i)%>',color: '<%=color1%>'},
						{title: 'Execute Daily Report',txt_data: 'NPY:<%=NPY%> | AI:<%=AI%> | NIY :<%=BNIY%>', url : 'p_monitoringExecDailyProd.asp?v_txtdate=<%=arrRSHeader(0, i)%>',start: '<%=set_format_tanggal(textDate(arrRSHeader(0, i)),"YYYY-MM-DD")%>',color: '<%=color2%>'},
						{x_name : 'button',title: 'LHK | LHP',url: 'p_laporan_harian_produksi_kebun.asp?txtdatefrom=<%=arrRSHeader(0, i)%>|p_laporan_harian_produksi_pabrik.asp?txtdatefrom=<%=arrRSHeader(0, i)%>',start: '<%=set_format_tanggal(textDate(arrRSHeader(0, i)),"YYYY-MM-DD")%>',color_btn1 : '<%=color_btn1%>', color_btn2 : '<%=color_btn2%>'},
						<%=bg%>
					<%
					Next
				Erase arrRS
				Erase arrRSHeader
			End If %>
			]
		});
		
	});

</script>
</head>

	<body onload="showHint('')">

	<!--#include file="menutop.asp"-->
	</br>
		<div style="Text-Align:Center; min-height:30px; "  id="header">
		   
				<h2>Laporan Agenda Monitoring</h2> 
		</div>
		<%
		Dim txtTable, txtModule
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
					<!-- <div style="border:1px #cccccc solid; padding:10px;" >-->
					<form method="POST" name="frmentry" action="<%=ScriptName%>" style="text-align:left;position:fixed;left:7px;top:5px;border:0px; padding:0px;margin:0px;z-index: 1000;" >
					<div><input class="datepicker" type="text" value="<%=set_format_tanggal(textDate(date_show),"MM/DD/YYYY")%>" name="txtdate"> <input name="submit" type="Submit" value="View"></div>
					<div Style="Background:RGBA(249,231,231,0.7); font-size:9px; padding:5px; border:1px #000 solid; margin-top:5px" >
					   NPY&nbsp;=&nbsp;&nbsp;Not Process Yet ( Kebun TBM ) </br>
					   AI&nbsp;&nbsp;&nbsp;&nbsp;=&nbsp;&nbsp;Already Input   ( kebun TM/Pabrik , sdh Process dan sdh input )  </br>
					   NIY&nbsp;&nbsp;=&nbsp;&nbsp;Not Input Yet   ( Kebun TM / Pabrik, sudah process, tapi Belum input )
					   
					</div>
					
					</form>
					
						<div id='calendar'></div>
					
					<!--</div>-->
					<div class="clear"></div>
	</body>
</html>