<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim textReportName : textReportName = "Closing Daily"
Dim ScriptName : ScriptName = "closing_daily.asp"

ScriptName =  ScriptName
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

 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">
function showHint(str) {
if (str.length == 0) {
 str = "null";	
}
 inday = document.getElementById("inday").value;
 fcba = document.getElementById("fcba").value;
 
    if (fcba.length == 0) {
        document.getElementById("status").innerHTML = "No Status";
        return;
    } else {
		document.getElementById("status").innerHTML = "<img src='css/images/load.gif'>";
        var xmlhttp = new XMLHttpRequest();
        xmlhttp.onreadystatechange = function() {
            if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
				output = xmlhttp.responseText;
				colosing = output.trim();
                
				if (colosing == "Closing"){
					document.getElementById("status").innerHTML = "<font color='red' >" + colosing + "</font>";
					document.getElementById("button_exe").innerHTML = '<input type="submit" value="UPDATE DAILY CLOSING" onclick=showHint("doupdate") >';
				}else if(colosing == "Not Yet Closing"){
					document.getElementById("status").innerHTML = "<font color='green' >" + colosing + "</font>";
					document.getElementById("button_exe").innerHTML = '<input type="button" style="color:#a0a0a0;" value="UPDATE DAILY CLOSING" disabled>';
				}else{
					document.getElementById("status").innerHTML = colosing;
					document.getElementById("button_exe").innerHTML = '<input type="button" style="color:#a0a0a0;" value="UPDATE DAILY CLOSING" disabled>';
				}
            }
        }
		 
		txt_FCBA = fcba.trim();
		//alert("ajax/closing_daily_ajx.asp?fcba=" + txt_FCBA + "&inday=" + inday  + "&exec=" + str );
        xmlhttp.open("GET", "ajax/closing_daily_ajx.asp?fcba=" + txt_FCBA + "&inday=" + inday  + "&exec=" + str, true);
        xmlhttp.send();
    }
}
</script>
	</head>

	<body onload="showHint('')">
		<!--#include file="top.jsp"-->
		<%
		Dim INDAY : INDAY = Request("INDAY")
		Dim FCBA : FCBA = Request("FCBA")
		dim status : status = "No Status"
		
		If Len(Trim(INDAY)) = 0 Then
			INDAY = textDate(Date)
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
				<div style="border:1px #cccccc solid; padding:10px; width:400px;" >
					<table class = "data">
						<thead>
							
						</thead>
						<tbody>
							<tr>
								<td align="left" width="100">FCBA</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;<select id="fcba" name="txtFcba" onCHANGE="showHint('')" style="width:100px;">

										<%
										Dim rsFcba,strsql,v_txtFcba
										Set rsFcba = Server.CreateObject("ADODB.Recordset")
										rsFcba.CursorLocation = adUseClient
										strsql =    " select fcba from user_fcba_tab " & _
													" where " & _
													" fcusercode = '" & session("USERNAME") & "' "
										
										rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
										if not rsFcba.eof then
											DO while not rsFcba.eof
												if v_txtFcba = rsFcba("fcba") then
												%>			
													<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
												<% Else %>	
													<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
												<% End If
												
												rsFcba.movenext
											loop
										End if
										rsFcba.Close
										set rsFcba = Nothing
										%>
								</select>
								</td>
							</tr>
							<tr>
								<td align="left">DATE</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;<input id="inday" class="datepicker" type="text" maxlength="10" name="INDAY" size="10" value="<%=INDAY%>" onCHANGE="showHint('')"  onkeyup="showHint('')" style="width:100px;">
								</td>
							</tr>
							<tr>
								<td align="left">STATE</td>
								<td align="left" >
									:&nbsp;&nbsp;&nbsp;&nbsp;<span id="status" style="font-weight:bold; font-size:13px;"><%=status%></span>
									&nbsp;&nbsp;&nbsp;
								</td>
							</tr>
							<tr>
								<td colspan="2" align="left" style="padding-top:30px;"><span id="button_exe"></span></td>
							</tr>
							
						</tbody>
					</table>
				</div>
				</div>
		</div>
		<!--#include file="bottom.jsp"-->
	</body>
</html>