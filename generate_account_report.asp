<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim textReportName : textReportName = "Generate Account Report"
Dim ScriptName : ScriptName = "generate_account_report.asp"

ScriptName = "p_" & ScriptName
%>
<html>
	<head>
		<title>ePlaS - <%=PCase(textReportName)%></title>
		<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
	</head>

	<body>
		<!--#include file="top.jsp"-->
		<%
		Dim DATE_FROM : DATE_FROM = Request("DATE_FROM")
		Dim DATE_TO : DATE_TO = Request("DATE_TO")
		
		If Len(Trim(DATE_FROM)) = 0 Then
			DATE_FROM = textDate(Date)
		End if
		
		If Len(Trim(DATE_TO)) = 0 Then
			DATE_TO = textDate(Date)
		End if
		
		Dim txtTable, txtModule
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
		<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
		<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
				<div class="judul">
					Module : Generate Report/ <%=PCase(textReportName)%>
				</div>
				<div class="box_both" style="overflow:scroll;">
				<form action="<%=ScriptName%>" method="post" name="frmEntry" target="_blank">
					<table class = "data">
						<thead>
						</thead>
						<tbody>
							<tr>
								<td align="left" width="150">Tahun</td>
				<% 
					Dim mon, Yr, MinYr, MaxYr, i, MM, v_txtMonth, v_txtYear
					mon = Month(Session("v_Costbook"))
					If len(v_txtMonth)= 0 then v_txtMonth = mon

					Yr = Year(Session("v_Costbook"))
					If len(v_txtYear) = 0 then v_txtYear = Yr
					MinYr = Yr - 2
					MaxYr = Yr + 2
					
				%>
						<td align="left">
							<select name="txtYear">
				<%
					For i = MinYr to MaxYr
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
				<%
					Next
				%>
							</select>
						</td>
							</tr>
							<tr>
								<td align="left">Bulan</td>
						<td align="left">
							<select name="txtMonth">
				<%
					For i = 1 to 12
				%>
						<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
				<%
					Next
				%>
							</select>
						</td>
							</tr>
							<tr>
								<td align="left"><input type="submit" value="Process" ></td>
								<input Type="Hidden" Value="ByVariable" Name="ScriptExecuted">
							</tr>
						</tbody>
					</table>
				</form>
				</div>
		</div>
		<!--#include file="bottom.jsp"-->
	</body>
</html>