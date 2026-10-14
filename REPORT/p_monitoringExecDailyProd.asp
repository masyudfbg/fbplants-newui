<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
dim v_txtdate,rsmonitoringex,no,cClass
dim strMain,ScriptName,v_txtdate_2
Dim var_colspan : var_colspan = 1

v_txtdate = request("v_txtdate")

	if v_txtdate = "" then
		v_txtdate = date()
	end if
v_txtdate_2 = formatNumber(Month(date),0) &  "/" & formatNumber(Day(date),0)-2 & "/" & Year(date) 

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" select region,fcba, cases, totalrecord " & _
				" from iplasprod.summary_exec_daily_production " & _
				" where fddate = to_date('" & v_txtdate & "','MM/DD/YYYY')" & _
				" order by region_api.get_OrderData(region) " 
								

%>
<html>

<head>
<title>ePlaS Reporting - Monitoring Execute Daily</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form id="frmEntry" action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
	
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">Date</font>
						</td>
						<td>
							<input id="date" class="datepicker" type="text" value="<%=v_txtdate%>" name="v_txtdate"> 
						</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>

<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
	<thead style="border-left:1px #000000 solid;">
		<tr>
			<th colspan="<%=var_colspan+8%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header-wrapper">
		<div id="head_judul">
					<h4></h4> 
					<h5></h5>   
				</div>
				<div id="subheader" class="centered">
					<h2>MONITORING EXECUTE DAILY</h3>
					<br/>
				</div>
              </div
			</th>
		</tr>
	<tr>
	    <th style="background-color:#00A300;color:#fff;">No</th>		
		<th style="background-color:#00A300;color:#fff;">Region</th>
		<th style="background-color:#00A300;color:#fff;">FCBA</th>
		<th style="background-color:#00A300;color:#fff;">Cases</th>
		<th style="background-color:#00A300;color:#fff;">Total Record</th>
	
	</tr>

    </thead>
	<tbody>
<%
Set rsmonitoringex = Server.CreateObject("ADODB.Recordset")
rsmonitoringex.Open strMain, connstr, adOpenStatic

if rsmonitoringex.bof and rsmonitoringex.eof then
%>

	<tr>
		<td colspan="<%=var_colspan+4%>">NO DATA AVAILABLE</td>
	</tr>
	<script>
	
		if (confirm('Tanggal <%=v_txtdate%> tidak ada data, redirect 2 hari sebelumnya ?')) {
			document.getElementById("date").value = "<%=v_txtdate_2%>";
			if (document.getElementById("date").value == "<%=v_txtdate_2%>"){
				document.getElementById("frmEntry").submit();
			}
		}else{}
	</script>
<%
else
		No=0
		while not rsmonitoringex.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
		    No = No + 1 
%>
				<tr class="<%=cClass%>">
					<td align="center"><%=No%></td>
					<td align="left"> <%=rsmonitoringex("region")%> </td>
					<td align="center"> <%=rsmonitoringex("fcba")%> </td>
					<td align="left"> <%=rsmonitoringex("cases")%> </td>
					<td align="right"> <%=rsmonitoringex("totalrecord")%> </td>
				</tr>
<%
			rsmonitoringex.MoveNext
			Wend	
rsmonitoringex.Close
set rsmonitoringex = Nothing
%>
<%End if%>
	</tbody>
</table>

   </body>
   </html>