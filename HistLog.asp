<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<html>
	<head>
		<title>ePlaS - History Log</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script language="JavaScript" src='js/x.js'></script>
<script language="JavaScript" src='js/xtableheaderfixed.js'></script>
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
		Dim txtTable, txtModule, v_txtSystem
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	    Module : HistLog
	</div>
 <div class="box_both" style="overflow:scroll;">

					<table class="xthf-site data green">
						<tr>
							<th align="left">NO</td>
							<th valign="top" align="left">SITE</th>
							<th valign="top" align="left">USERNAME</th>
							<th valign="top" align="left">IP ADDRESS</th>
							<th valign="top" align="left">DATE ACTION</th>
							<th valign="top" align="left">TIME ACTION</th>
						</tr>
				<%
				Dim rsHistLog, strHistLog, TempEvenOrOdd	

				v_txtSystem	=	Request("txtSystem")

				strHistLog	=	" SELECT" & _
								"     BU FCBA" & _
								"     , USERNAME" & _
								"     , IP_ADDRESS" & _
								"     , DATE_ACT" & _
								"     , TIME_ACT" & _
								" FROM HISTORICAL_USER_TAB" & _
								" WHERE BU = '" & Session("Fcba") & "' AND USERNAME = '" & Session("CodeName") & "'"
				
				strHistLog	=	strHistLog & " ORDER BY USERNAME ASC, DATE_ACT, TIME_ACT DESC "


				Set rsHistLog = Server.CreateObject("ADODB.Recordset")
				rsHistLog.Open strHistLog, connstr, adOpenStatic
										
				if (rsHistLog.bof and rsHistLog.eof) then
				%>
						<tr bgcolor="#EEEEEE"><td align="left" colspan="5">No Available</td></tr>
				<%
				elseif not rsHistLog.eof Then
				Dim Numb : Numb = 1
					while not rsHistLog.eof
						Numb = Numb
						if Numb  mod 2 = 1 then
							TempEvenOrOdd = "#EEEEEE"
						else
							TempEvenOrOdd = "#E5E5E5"
						end if
				%>
						<tr bgcolor="<%=TempEvenOrOdd%>">
							<td align="Center"><%=Numb%></td>
							<td align="Center"><%=rsHistLog("FCBA")%></td>
							<td align="left"><%=rsHistLog("USERNAME")%></td>
							<td align="left"><%=rsHistLog("IP_ADDRESS")%></td>
							<td align="left"><%=rsHistLog("DATE_ACT")%></td>
							<td align="left"><%=rsHistLog("TIME_ACT")%></td>
						</tr>
				<%
						Numb = Numb + 1
						rsHistLog.MoveNext
					wend
					rsHistLog.Close : set rsHistLog = Nothing
				end if	
				%>
					</table>
</div></div>
  <!--#include file="bottom.jsp"-->

	</body>
</html>