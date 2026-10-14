<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<html>
	<head>
		<title>ePlaS - Account Period</title>
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
		Dim txtTable, txtModule, v_txtSystem
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : Account Period
	</div>
 <div class="box_both" style="overflow:scroll;">
					<table class="data">
						<tr>
							<form name="frmEntry" method="post">
							<td align="left">System</td>
							<td align="left">
								<input type="text" maxlength="15" name="txtSystem" size="15" value="<%=v_txtSystem%>" title="Input System">
							</td>
							<td align="left"><input type="submit" value="View" /></td>
							</form>
						</tr>
					</table>

					<table class="data green">
                    <thead>
						<tr>
							<th align="left">NO</th>
							<th valign="top" align="left">SITE</th>
							<th valign="top" align="left">SYSTEM</th>
							<th valign="top" align="left">LAST PERIOD POSTED</th>
							<th valign="top" align="left">CLOSE ACCOUNT YEAR</th>
							<th valign="top" align="left">CLOSE ACCOUNT PERIOD</th>
							<th valign="top" align="left">CURRENT DATE</th>
							<th valign="top" align="left">CURRENT ACCOUNT YEAR</th>
							<th valign="top" align="left">CURRENT ACCOUNT PERIOD</th>
						</tr>
                     </thead>
                     <tbody>
				<%
				Dim rsAccountPeriod, strAccountPeriod, TempEvenOrOdd	

				v_txtSystem	=	Request("txtSystem")

				strAccountPeriod	=	" SELECT" & _
										"     FCBA" & _
										"     , POSTINGPOSITION" & _
										"     , SYSTEM" & _
										"     , LASTPERIODPOSTED" & _
										"     , CURRENTDATE" & _
										"     , CURRENTACCOUNTINGYEAR " & _ 
										"	  , CURRENTACCOUNTINGPERIOD " & _    
										"	  , CLOSEACCOUNTINGYEAR " & _  
										"	  , CLOSEACCOUNTINGPERIOD " & _	
										" FROM PERIODCONTROL" & _
										" WHERE FCBA = '" & Session("Fcba") & "'"

				If Len(Trim(v_txtSystem)) > 0 Then
				strAccountPeriod	=	strAccountPeriod & " AND UPPER(SYSTEM) = '" & UCase(v_txtSystem) & "'"
				End If						

				strAccountPeriod	=	strAccountPeriod & " ORDER BY POSTINGPOSITION"


				Set rsAccountPeriod = Server.CreateObject("ADODB.Recordset")
				rsAccountPeriod.Open strAccountPeriod, connstr, adOpenStatic
										
				if (rsAccountPeriod.bof and rsAccountPeriod.eof) then
				%>
						<tr bgcolor="#EEEEEE"><td align="left" colspan="5">No Available</td></tr>
				<%
				elseif not rsAccountPeriod.eof Then
				Dim Numb : Numb = 1
					while not rsAccountPeriod.eof
						Numb = Numb
						if Numb  mod 2 = 1 then
							TempEvenOrOdd = "#EEEEEE"
						else
							TempEvenOrOdd = "#E5E5E5"
						end if
				%>
						<tr bgcolor="<%=TempEvenOrOdd%>">
							<td align="Center"><%=rsAccountPeriod("POSTINGPOSITION")%></td>
							<td align="Center"><%=rsAccountPeriod("FCBA")%></td>
							<td align="left"><%=rsAccountPeriod("SYSTEM")%></td>
							<td align="left"><%=rsAccountPeriod("LASTPERIODPOSTED")%></td>
							<td align="left"><%=rsAccountPeriod("CLOSEACCOUNTINGYEAR")%></td>
							<td align="left"><%=rsAccountPeriod("CLOSEACCOUNTINGPERIOD")%></td>
							<td align="left"><%=rsAccountPeriod("CURRENTDATE")%></td>
                            <td align="left"><%=rsAccountPeriod("CURRENTACCOUNTINGYEAR")%></td>
							<td align="left"><%=rsAccountPeriod("CURRENTACCOUNTINGPERIOD")%></td>							
						</tr>
				<%
						Numb = Numb + 1
						rsAccountPeriod.MoveNext
					wend
					rsAccountPeriod.Close : set rsAccountPeriod = Nothing
				end if	
				%>
                </tbody>
					</table>
                    </div>
</div></div>
  <!--#include file="bottom.jsp"-->

	</body>
</html>