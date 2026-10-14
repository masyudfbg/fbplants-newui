<%Option Explicit%>
<%On Error Resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Trial Balance</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
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
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%'-------------awal-----------------------%>
<form action=p_glmasterregistermonthly.asp method=post>
<table>
<tr>
	<td><font face="verdana" size="1">YEAR</font></td><td><input type=text name=txtyear></td>
	<td><font face="verdana" size="1">MONTH</font></td><td><input type=text name=txtmonth></td>
	<td><font face="verdana" size="1">JOBGROUP</font></td><td><input type=text name=txtjobgroup></td>
	<td><font face="verdana" size="1">JOB</font></td><td><input type=text name=txtjob></td>
	<td><input type=submit value=select></td>
</tr>
</table>
</form>

 
  <%'-------------akhir----------------------%>
 </td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
</html>
