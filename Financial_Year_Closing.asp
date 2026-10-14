<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim textReportName : textReportName = "Financial Year Closing"
Dim ScriptName : ScriptName = "financial_year_closing.asp"

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
		Dim CLOSINGYEAR : CLOSINGYEAR = Request("CLOSINGYEAR")
		
		Dim txtTable, txtModule
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
		<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
		<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
				<div class="judul">
					Module : General Ledger / <%=PCase(textReportName)%>
				</div>
				<div class="box_both" style="overflow:scroll;">
				<form action="<%=ScriptName%>" method="post" name="frmEntry" target="_blank">
					<table class = "data">
						<thead>
							
						</thead>
						<tbody>
							<tr>
								<td align="left" width="150">Yearly Closing</td>
								<td align="left">
									<input type="text" maxlength="10" name="CLOSINGYEAR" size="10" value="<%=CLOSINGYEAR%>">
								</td>
							</tr>
							
							<tr>
								<td align="left"><input type="submit" value="Process" ></td>
							</tr>
						</tbody>
					</table>
				</form>
				</div>
		</div>
		<!--#include file="bottom.jsp"-->
	</body>
</html>