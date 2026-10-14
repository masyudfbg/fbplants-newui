<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim textReportName : textReportName = "Calendar Stats"
Dim ScriptName : ScriptName = "calendar_stat.asp"
Dim stringStats,objRSArray,arrRS,i,ia
Dim q_dateToDistict,arrRSHeader,objRSArrayHeader
dim bg,color1,color2

Dim NPY,AI,BNIY,Tru,Fls,color_btn1,color_btn2

ScriptName =  ScriptName

	Function GetRSArrayHeader(strSQL)
			Response.Write(strSQL)
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
			Response.Write(strSQL)
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

<script type="text/javascript" language="JavaScript1.2" src="../stlib.js" ></script>
<link rel="StyleSheet" href="../dtree.css" type="text/css" />
<link rel="StyleSheet" href="../style-content.css" type="text/css" />
<style>

	#calendar {
		max-width: 900px;
		margin: 0 auto;
	}
	.fc-title{font-size:10px;}
	.fc-time{font-size:10px;}
</style>
<link href='../calendar/css/fullcalendar.min.css' rel='stylesheet' />
<link href='../calendar/css/fullcalendar.print.css' rel='stylesheet' media='print' />

<link rel="stylesheet" href="../css/jquery-ui.css" />
<script type="text/javascript" src="../dtree.js"></script>
<script src='../calendar/js/moment.min.js'></script>
<script src='../calendar/js/jquery.min.js'></script>
<script src='../calendar/js/fullcalendar.js'></script>
<%

dim date_now : date_now = set_format_tanggal(textDate(date()),"YYYY-MM-DD")
dim date_lusa : date_lusa = set_format_tanggal(textDate(date()-2),"YYYY-MM-DD")
dim date_show : date_show = request("txtdate")

if date_show = "" then
	date_show = date_now
end if

stringStats = 	" select 'exec' exec_sign, trunc(t.fddate)fddate,cases ,count(*) total from summary_exec_daily_production t" & _
				" where to_char(t.fddate,'YYYYMM') =  to_char(to_date('" & date_show & "','YYYY-MM-DD'),'YYYYMM') " & _
				" group by  trunc(t.fddate),cases" & _
				" union" & _ 
				" select 'input' exec_sign, trunc(t.fddate)fddate,isdailyclosing cases ,count(*) total  from dailyclosingactivity t" & _
				" where to_char(t.fddate,'YYYYMM') =  to_char(to_date('" & date_show & "','YYYY-MM-DD'),'YYYYMM') " & _
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
							bg = "{start: '" & date_lusa & "',end: '" & date_now & "',rendering: 'background',color: '#FFED96'},"
							color1 = "Green"
							color2 = "#F22935"
							color_btn1 = ""
							color_btn2 = ""							
						else
							bg = ""
							color1 = "#BCBCBC"
							color2 = "#bcbcbc"
							color_btn1 = "background-color:#BCBCBC;"
							color_btn2 = "background-color:#BCBCBC;"
						end if
						%>			
						{title: 'Daily Input Closing',txt_data: 'Closing: <%=Tru%> | Not Yet: <%=Fls%>',start: '<%=set_format_tanggal(textDate(arrRSHeader(0, i)),"YYYY-MM-DD")%>', url : 'report/p_monitoringDailyInputClosing.asp?v_txtdate=<%=arrRSHeader(0, i)%>',color: '<%=color1%>'},
						{title: 'Execute Daily Report',txt_data: 'Npy:<%=NPY%> | Ai:<%=AI%> | Niy :<%=BNIY%>', url : 'report/p_monitoringExecDailyProd.asp?v_txtdate=<%=arrRSHeader(0, i)%>',start: '<%=set_format_tanggal(textDate(arrRSHeader(0, i)),"YYYY-MM-DD")%>',color: '<%=color2%>'},
						{x_name : 'button',title: 'LHK | LHP',url: 'Report/p_laporan_harian_produksi_kebun.asp?txtdatefrom=<%=arrRSHeader(0, i)%>|Report/p_laporan_harian_produksi_pabrik.asp?txtdatefrom=<%=arrRSHeader(0, i)%>',start: '<%=set_format_tanggal(textDate(arrRSHeader(0, i)),"YYYY-MM-DD")%>',color_btn1 : '<%=color_btn1%>', color_btn2 : '<%=color_btn2%>'},
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
		<!--#include file="../top.jsp"-->
		<%
		Dim txtTable, txtModule
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
		<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="../menu.asp"--> 
		</div>
		<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
				<div class="judul">
					Module : <%=PCase(textReportName)%>
				</div>
				<div class="box_both" style="overflow:scroll;">
					<!-- <div style="border:1px #cccccc solid; padding:10px;" >-->
					<form method="POST" name="frmentry" action="<%=ScriptName%>" >
					<div><input class="datepicker" value="<%=date_show%>" name="txtdate"> <input name="submit" type="Submit" value="View"></div>
					
						<div id='calendar'></div>
					</form>
					<!--</div>-->
				</div>
		</div>
		<!--#include file="../bottom.jsp"-->
	</body>
</html>