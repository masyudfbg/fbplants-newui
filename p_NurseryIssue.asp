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
<title>ePlaS Reporting - Nursery Issue</title>
</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
'cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
'cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
'cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtPeriodDate=REQUEST("PeriodDate")

v_txtPODATE_From=trim(request("DATEACTIVITY_From"))
v_txtPODATE_To=trim(request("DATEACTIVITY_To"))
v_NurseryCode=trim(request("NurseryCode"))
v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
v_LOCATIONCODE=trim(request("LOCATIONCODE"))

Dim v_txtNurseryNo : v_txtNurseryNo = Trim(Request("txtNurseryNo"))
Dim v_txtJobCode : v_txtJobCode = Trim(Request("txtJobCode"))

txtSQL =	" select fcba" & _
			"        , documentno" & _
			"        , NURSERYISSUEHDR_API.get_issuedate(documentno, fcba) issuedate" & _
			"        , linenokey" & _
			"        , nurserycode" & _
			"        , jobcode" & _
			"        , job_api.get_jobcode_name(jobcode, fcba) jobname" & _
			"        , locationtype" & _
			"        , locationcode" & _
			"        , reference" & _
			"        , remarks" & _
			"        , linenohist" & _
			"        , nvl(sum (nvl(quantity, 0)), 0) quantity" & _
			"        , nvl(sum(nvl(amount, 0)), 0) amount" & _
			" from nurseryissuedtl" & _
			" where fcba = '" & Session("fcBA") & "'"
			
if  len(trim(v_txtNurseryNo)) > 0 then
    txtSQL = txtSQL & " and documentno = '" & v_txtNurseryNo & "'"
end if
if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and NURSERYISSUEHDR_API.get_issuedate(documentno, fcba) >= to_date('" & v_txtPODATE_From & "', 'mm/dd/yyyy')"
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and NURSERYISSUEHDR_API.get_issuedate(documentno, fcba) <= to_date('" & v_txtPODATE_To & "', 'mm/dd/yyyy')"
end if
if  len(trim(v_NurseryCode)) > 0 then
    txtSQL = txtSQL & " and nurserycode = '" & v_NurseryCode & "'"
end if
if  len(trim(v_txtJobCode)) > 0 then
    txtSQL = txtSQL & " and jobcode = '" & v_txtJobCode & "'"
end if
if  len(trim(v_LOCATIONTYPE)) > 0 then
    txtSQL = txtSQL & " and locationtype = '" & v_LOCATIONTYPE & "'"
end if
if  len(trim(v_LOCATIONCODE)) > 0 then
    txtSQL = txtSQL & " and locationcode = '" & v_LOCATIONCODE & "'"
end if

txtSQL = txtSQL &	" group by fcba" & _
					"          , documentno" & _
					"          , NURSERYISSUEHDR_API.get_issuedate(documentno, fcba)" & _
					"          , linenokey" & _
					"          , nurserycode" & _
					"          , jobcode" & _
					"          , locationtype" & _
					"          , locationcode" & _
					"          , reference" & _
					"          , remarks" & _
					"          , linenohist"
txtSQL = txtSQL &	" order by fcba" & _
					"          , documentno" & _
					"          , NURSERYISSUEHDR_API.get_issuedate(documentno, fcba)" & _
					"          , linenokey" & _
					"          , nurserycode" & _
					"          , jobcode" & _
					"          , locationtype" & _
					"          , locationcode" & _
					"          , reference" & _
					"          , remarks" & _
					"          , linenohist"
'response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
		<div class="printhide">
			<!--#include file="print_excel.asp"-->
		</div>
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					<th colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
							<h4>PT. <%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h2>NURSERY ISSUE</h2>
						</div>
					</th>
				</tr>
				<!--<div class="printhide">-->
				<tr>
					<td colspan="13" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<!--</div>-->
	<%	
		Set rsCat = connStr.Execute(txtSQL)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1
	%>
				<tr>
					<th>DOCUMENT NUMBER</th>
					<th>ISSUE DATE</th>
					<th>LINE NO (KEY)</th>
					<th>NURSERY CODE (BATCH)</th>					
					<th>JOB</th>
					<th>JOB DESCRIPTION</th>
					<th>TARGET TYPE</th>
					<th>TARGET CODE</th>
					<th>REFERENCE</th>
					<th>REMARKS</th>
					<th>LINE NO (HISTORY)</th>
					<th>QUANTITY</th>
					<th>AMOUNT</th>
				</tr>
			</thead>
			<tbody>
	<%
    Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open txtSql, connstr
	if Not rsCat.eof then
		nCountofField=rsCat.fields.count-1
		no=1
		nQty=0
		nAmount=0
		Dim totalamount : totalamount = 0
		While Not rsCat.EOF
	%>
				<tr>		 
					<td align="Left"><font face="Verdana" size="1"><%=rscat("documentno")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("issuedate")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("linenokey")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("nurserycode")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("jobname")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("locationtype")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("locationcode")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("reference")%></td>
					<td align="Left"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
					<td align="Right"><font face="Verdana" size="1"><%=rscat("linenohist")%></td>
					<td align="Right"><font face="Verdana" size="1"><%=formatnumber(rscat("quantity"), 0)%></td>
					<td align="Right"><font face="Verdana" size="1"><%=formatnumber(rscat("amount"), 0)%></td>					
				</tr>
	<%
			nQty=nQty+trim(cdbl(rsCat("QUANTITY")))
			totalamount = totalamount + Cdbl(rscat("amount"))
			rsCat.MoveNext
			intRecordsPerPage = intRecordsPerPage - 1
			no=no+1
		Wend
	%>
				<tr bgcolor="#F4F4A0">
					<td align="Center"colspan="11"> <font face="Verdana" size="1"><strong>T o t a l</strong></td>
					<td align="Right"> <font face="Verdana" size="1"><strong><%=FormatNumber(nQty, 0)%></strong></td>
					<td align="Right"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalamount, 0)%></strong></td>
				</tr>	  
	  
	<%
				else
	%>
				<tr>
					<td align="Left" style="<%=cmdstyle%>" colspan="13"><font face="Verdana" size="1">NO DATA AVAILABLE</font></td>
				</tr>
	<%  
	end if 
	rsCat.Close
	set rsCat = Nothing%>
			</tbody>
		</TABLE>
<!--#include file="Bottom_Report.asp"-->
	</DIV>
</BODY>
</HTML>