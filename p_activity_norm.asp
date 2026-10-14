<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
<title>ePlaS Reporting - Activity Norm</title>
</head>
<body>
<%
jobcode = request("jobcode")
	if jobcode <> "" then
	txtSQL = " SELECT job.fccode,job.fcname,job.job_forfieldstatus" & _
	         "        , ur.UOMJOB" & _
	         "        , ur.MANDAYSNORM" & _
			 "        , ur.SOILTYPE" & _
			 "        , ur.TERRAINTYPE" & _
             " FROM ACTIVITY_NORM ur" & _
			 " inner join job on ur.job = job.fccode and ur.fcba=job.fcba"& _			 
             " WHERE ur.fcba = '" & session("fcBA") & "'" & _
             "       AND ur.job = '" & jobcode & "'"
	else
	txtSQL = " SELECT job.fccode,job.fcname,job.job_forfieldstatus" & _
	         "        , ur.UOMJOB" & _
	         "        , ur.MANDAYSNORM" & _
			 "        , ur.SOILTYPE" & _
			 "        , ur.TERRAINTYPE" & _
             " FROM ACTIVITY_NORM ur" & _
			 " inner join job on ur.job = job.fccode and ur.fcba=job.fcba"& _
             " WHERE ur.fcba = '" & session("fcBA") & "'"
	end if

'response.Write txtsql
Set rsCat = connStr.Execute(txtSQL)
%>
	<body>
		<div id="wrapper">
		<!--#include file="print_excel.asp"-->
		<table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="5" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>ACTIVITY NORM</h2>
				</div>
			</div>
		</th>
		</tr>
			<!--<div class="printhide">-->
		<tr>
				<th>JOB CODE</th>
				<th>JOB NAME</th>
				<th>FIELD STATUS</th>
				<th>UOM JOB</th>
				<th>NORM</th>
				<th>SOILTYPE</th>
				<th>TERRAINTYPE</th>
			</tr>
	<tbody>
	<%While Not rsCat.EOF%>
	 	<tr>
		 
		<td align="left"><font face="Verdana" size="1"><%=rscat("fccode")%></td>
		<td align="left"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
		<td align="left"><font face="Verdana" size="1"><%=rscat("job_forfieldstatus")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("UOMJOB")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("MANDAYSNORM")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("SOILTYPE")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("TERRAINTYPE")%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null%>
	<form>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
	</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>