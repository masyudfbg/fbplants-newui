<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
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
<title>ePlaS Reporting - Supplying Declaration</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

date_from = request("date_from")
date_to = request("date_to")
documentno = request("documentno")
declarationtype = request("declarationtype")
division = request("division")
fieldcode = request("fieldcode")

txtSelect = txtSelect & " From " & date_from & " To " & date_to

txtSQL = " SELECT fcba, documentno, transdate, field_api.getvalid_afdeling(fieldcode, fcba, transdate) division, fieldcode, quantity, declarationtype, declarationtypedesc, remarks"& _
		 " FROM view_supplyingdeclaration" & _
		 " WHERE fcba = '" & Session("fcBA") & "'"
if len(trim(documentno)) > 0 then txtSQL = txtSQL & " AND documentno = '" & documentno & "'" : txtSelect = txtSelect & " Document No " & documentno
if len(trim(declarationtype)) > 0 then txtSQL = txtSQL & " AND declarationtype = '" & declarationtype & "'"
if len(trim(division)) > 0 then txtSQL = txtSQL & " AND field_api.getvalid_afdeling(fieldcode, fcba, transdate) = '" & division & "'" : txtSelect = txtSelect & " Division " & division
if len(trim(fieldcode)) > 0 then txtSQL = txtSQL & " AND fieldcode = '" & fieldcode & "'" : txtSelect = txtSelect & " Fieldcode " & fieldcode


if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

' response.Write(txtsql)
'response.Flush()
%>
 <body>
		<div id="wrapper">
 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="7" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>SUPPLYING DECLARATION</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th>Document No</th>
					<th>Document Date</th>
					<th>Division</th>
					<th>Field Code</th>
					<th>Declaration Type</th>
					<th>Quantity</th>
					<th>Remarks</th>
					
				</tr>
				</thead>
               <tbody>
<%
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
		While Not rsCat.EOF

%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("documentno")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("division")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("declarationtypedesc")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Quantity")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
<%
		rsCat.MoveNext
		Wend
	Else
		response.write "<tr><td colspan='7'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->