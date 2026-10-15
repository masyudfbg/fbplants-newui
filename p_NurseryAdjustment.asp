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
<title>ePlaS Reporting - Nursery Adjustment</title>
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

Dim v_txtNurseryNo : v_txtNurseryNo = Trim(Request("txtNurseryNo"))


v_Crop=trim(request("Crop"))
v_Source=trim(request("Source"))


txtSQL = " select documentno" & _
         "        , NURSERYADJHDR_API.get_adjustmentdate(documentno, fcba) adjustmentdate" & _
         "        , linenokey" & _
		 "        , nurserycode" & _
		 "        , nursery_api.get_description(nurserycode, fcba) nurseryname" & _
         "        , reference" & _
         "        , remarks" & _
         "        , linenohist" & _
         "        , quantity" & _
         " from nurseryadjdtl" & _
         " where fcba ='" & Session("fcBA") & "'"
		 
if  len(trim(v_txtNurseryNo)) > 0 then
    txtSQL = txtSQL & " and documentno = '" & v_txtNurseryNo & "'"
end if

if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and NURSERYADJHDR_API.get_adjustmentdate(documentno, fcba) >= to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and NURSERYADJHDR_API.get_adjustmentdate(documentno, fcba) <= to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
end if
if  len(trim(v_NurseryCode)) > 0 then
    txtSQL = txtSQL & " and NurseryCode = '" & v_NurseryCode & "'"
end if

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
					<th colspan="9" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="head_judul" style="margin-bottom:-30px; position:static; text-align:left;">
							<h4>PT. <%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h2>NURSERY ADJUSTMENT</h2>
						</div>
					</th>
				</tr>
				<tr>
					<td colspan="9" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
		<%	
		Set rsCat = connStr.Execute(txtSQL)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1%>
		 
				<tr>
					<th>DOCUMENT NUMBER</th>
					<th>ADJUSTMENT DATE</th>
					<th>LINE NO (KEY)</th>
					<th>NURSERY CODE (BATCH)</th>
					<th>NURSERY NAME</th>
					<th>REFERENCE</th>
					<th>REMARKS</th>
					<th>LINE NO (HISTORY)</th>
					<th>QUANTITY</th>
				</tr>
			</thead>
			<tbody>
    <%
				Dim nQty
				Set rsCat = Server.CreateObject("ADODB.Recordset")
				rsCat.CursorType = 3
				rsCat.CursorLocation = 3
				rsCat.Open txtSql, connstr
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
					no=1
					nQty=0
					nAmount=0
					While Not rsCat.EOF
	%>
				<tr>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("documentno")%></font></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("adjustmentdate")%></font></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("linenokey")%></font></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("nurserycode")%></font></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("nurseryname")%></font></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("reference")%></font></td>
					<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></font></td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("linenohist")%> </td>
					<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("quantity"), 2)%></font></td>
				</tr>
	<%
						nQty=nQty+trim(cdbl(rsCat("QUANTITY")))
						rsCat.MoveNext
						intRecordsPerPage = intRecordsPerPage - 1
						no=no+1
					Wend
	%>
				<tr bgcolor="#F4F4A0">
					<td align="Center" style="<%=cmdstyleLeft%>" colspan="8"><font face="Verdana" size="1"><strong>Total</strong></font></td>
					<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(nQty, 2)%></strong></font></td>
				</tr>

	<%
				else
	%>
				<tr>
					<td align="Left" style="<%=cmdstyle%>" colspan="9"><font face="Verdana" size="1">NO DATA AVAILABLE</font></td>
				</tr>
	<%
				end if 
				rsCat.Close
				set rsCat = Nothing
	%>
			</tbody>
		</TABLE>
		<!--#include file="Bottom_Report.asp"-->
	</DIV>
</BODY>
</HTML>