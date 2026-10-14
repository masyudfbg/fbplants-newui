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
<title>ePlaS Reporting - BBC Distribution</title>
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


txtSQL = "	select fcba, " & _
		 "	SUM (DECODE (censuscode, 'BBC1', distribution1, 0)) JAN, "& _
		 "	SUM (DECODE (censuscode, 'BBC1', distribution2, 0)) FEB, "& _
		 "	SUM (DECODE (censuscode, 'BBC1', distribution3, 0)) MAR, "& _
		 "	SUM (DECODE (censuscode, 'BBC1', distribution4, 0)) APR, "& _
		 "	SUM (DECODE (censuscode, 'BBC2', distribution1, 0)) MEI, "& _
		 "	SUM (DECODE (censuscode, 'BBC2', distribution2, 0)) JUN, "& _
		 "	SUM (DECODE (censuscode, 'BBC2', distribution3, 0)) JUL, "& _
		 "	SUM (DECODE (censuscode, 'BBC2', distribution4, 0)) AGU, "& _
		 "	SUM (DECODE (censuscode, 'BBC3', distribution1, 0)) SEP, "& _
		 "	SUM (DECODE (censuscode, 'BBC3', distribution2, 0)) OKT, "& _
		 "	SUM (DECODE (censuscode, 'BBC3', distribution3, 0)) NOV, "& _
		 "	SUM (DECODE (censuscode, 'BBC3', distribution4, 0)) DES " & _
		 "	from bbcdistribution "& _
		 "	where periodyear='" & txtyear & "'" & _
		 "  group by fcba " 

if  len(trim(txtyear)) > 0 then
	txtSelect = txtSelect & " Periode " & txtyear
end if

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
					<h2>BBC DISTRIBUTION</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th rowspan="2">Tahun</th>
					<th colspan="5">BBC 1</th>
					<th colspan="5">BBC 2</th>
					<th colspan="5">BBC 3</th>
				</tr>
				<tr>
					<th>Jan</th>
					<th>Feb</th>
					<th>Mar</th>
					<th>Apr</th>
					<th>Total</th>
					
					<th>Mei</th>
					<th>Jun</th>
					<th>Jul</th>
					<th>Agu</th>
					<th>Total</th>
					
					<th>Sep</th>
					<th>Okt</th>
					<th>Nov</th>
					<th>Des</th>
					<th>Total</th>
					
				</tr>
                </thead>
               <tbody>
	<%
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	While Not rsCat.EOF
		TOTALBBC1 = cdbl(to_nol(rsCat("JAN"))) + cdbl(to_nol(rsCat("FEB"))) + cdbl(to_nol(rsCat("MAR"))) + cdbl(to_nol(rsCat("APR")))
		TOTALBBC2 = cdbl(to_nol(rsCat("MEI"))) + cdbl(to_nol(rsCat("JUN"))) + cdbl(to_nol(rsCat("JUL"))) + cdbl(to_nol(rsCat("AGU")))
		TOTALBBC3 = cdbl(to_nol(rsCat("SEP"))) + cdbl(to_nol(rsCat("OKT"))) + cdbl(to_nol(rsCat("NOV"))) + cdbl(to_nol(rsCat("DES")))

	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=txtyear%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JAN")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FEB")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MAR")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("APR")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALBBC1%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MEI")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JUN")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JUL")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("AGU")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALBBC2%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SEP")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("OKT")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NOV")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DES")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALBBC3%></td>
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