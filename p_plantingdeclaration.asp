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
<title>ePlaS Reporting - Planting Declaration</title>
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

' v_txtCUSTOMERCODE= trim(request("periodyear"))
txtyear = request("txtyear")
documentno = request("documentno")
fieldcode = request("fieldcode")
TRANSDATE_from = request("TRANSDATE_from")
TRANSDATE_to = request("TRANSDATE_to")
fullstand = request("fullstand")

txtSQL = "	select ph.transdate, " & _
		 "  pd.fieldcode, pd.Quantity, pd.Hectare, ph.documentno, pd.remarks, pd.fullstand "& _
		 "	from plantingdeclarationdtl pd "& _
		 "  inner join plantingdeclarationhdr ph on (pd.documentno = ph.documentno and pd.fcba = ph.fcba)" & _
		 "	where pd.fcba = '"& Session("fcBA") & "'"

if cstr(""&TRANSDATE_from) = cstr(""&TRANSDATE_to) then
	txtSelect = txtSelect & " Date : " & Day(TRANSDATE_to) & " " & MonthName(Month(TRANSDATE_to)) & " " & Year(TRANSDATE_to)
else
	txtSelect = txtSelect & " From " & Day(TRANSDATE_from) & " " & MonthName(Month(TRANSDATE_from)) & " " & Year(TRANSDATE_from)
	txtSelect = txtSelect & " To " & Day(TRANSDATE_to) & " " & MonthName(Month(TRANSDATE_to)) & " " & Year(TRANSDATE_to)
end if

' if  len(trim(documentno)) > 0 then
' 	txtSelect = txtSelect & " Document No" & documentno
' end if
' if  len(trim(fieldcode)) > 0 then
' 	txtSelect = txtSelect & " Blok " & fieldcode
' end if
' if  len(trim(fullstand)) > 0 then
' 	txtSelect = txtSelect & " Fullstand " & fullstand
' end if

if len(trim(documentno)) > 0 then txtSQL = txtSQL & " AND pd.documentno = '" & documentno & "'" : txtSelect = txtSelect & " Document No " & documentno
if len(trim(fieldcode)) > 0 then txtSQL = txtSQL & " AND pd.fieldcode = '" & fieldcode & "'" : txtSelect = txtSelect & " Fieldcode " & fieldcode
if len(trim(fullstand)) > 0 then txtSQL = txtSQL & " AND pd.fullstand = '" & fullstand & "'" : txtSelect = txtSelect & " Fullstand " & fullstand


txtsql=(txtsql) 
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
         <th colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PLANTING DECLARATION</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th >Document No</th>
					<th >Document Date</th>
					<th >Field Code</th>
					<th >Quantity</th>
					<th >Hectare</th>
					<th >Fullstand</th>
					<th >Remarks</th>
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
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldcode")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Quantity")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Hectare")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Fullstand")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Remarks")%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='16'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->