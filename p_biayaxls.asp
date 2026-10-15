<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Invoice</title>
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
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

dim tglawal,tglakhir

tglawal= trim(request("tglawal"))
tglakhir= trim(request("tglakhir"))


businessunit= trim(session("fcba"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
	set rsview = connStr.Execute(q)
	if not rsview.eof then
	   fcbaname = rsview("fcName")
	   CompanyName=rsview("FCCOMPANYNAME")
	   MgrName = rsview("FCMANAGER")
	   KTUName = rsview("FCKTU")
	   CropName=rsview("FCCROP")
	   FcbaType=rsview("fcType")
	end if

Dim sqlview,rsview 
sqlview = "Select V.BKMDATE, V.SOURCE,V.FCBA, V.ACTCODE || ' - ' || V.ACTNAME ACTNAME, JOB_API.GET_JOB_UOM(V.ACTCODE,V.FCBA) UNITOFMEASUREMENT, " & _
"V.LOCATIONTYPE, V.LOCATIONCODE,nvl(V.HK,0) HK,nvl(V.UNIT,0) UNIT, nvl(V.FISIK,0)FISIK, nvl(V.AMOUNT,0)AMOUNT,  " & _
"V.JOBGROUPCODE || ' - ' || V.JOBGROUPNAME GROUPNAME, V.ACTIVITY_HEADING " & _
"From VIEW_BKMPROSES3 V " & _
"WHERE V.BKMDATE BETWEEN TO_DATE('" & tglawal & "','MM/DD/YYYY') AND TO_DATE('" & tglakhir & "','MM/DD/YYYY') AND V.FCBA='" & businessunit & "' ORDER BY V.ACTCODE,V.LOCATIONCODE,V.BKMDATE, V.SOURCE "
set rsview = connStr.Execute(sqlview)


 'response.Write(sqlview)
'response.Flush()
%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>HARIAN BKM</h2>
				</div>
			</div>
				</th>
			</tr>
				<!--<div class="printhide">-->
						<tr>
							<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
						
						
						</tr> 
						<tr>
							
							<th>Tanggal</th>
							<th>Source</th>
							<th>Group Name</th>
							<th>Activity</th>
							<th>UOM</th>
							<th>TPYPE</th>
							<th>TCODE</th>	  
							<th>HK</th>
							<th>UNIT</th>
							<th>OUTPUT(Ha)</th>
							<th>Amount</th>
							
						</tr>
                         </thead>

						 <%
						 	dim ttlamnt
							ttlamnt = 0
							ttlhk = 0
							While Not rsview.EOF

						 %>

                         <tbody>
						<tr>
														 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("BKMDATE")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("SOURCE")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("GROUPNAME")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("ACTNAME")%></td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("UNITOFMEASUREMENT")%></td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("LOCATIONTYPE")%></td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("LOCATIONCODE")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("HK")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("UNIT")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("FISIK")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsview("AMOUNT"),2)%></td>
						</tr>
						
						</tbody>
						<%
						ttlamnt = cdbl(ttlamnt) +cdbl(rsview("AMOUNT"))
						ttlhk = cdbl(ttlhk) +cdbl(rsview("hk"))
						rsview.MoveNext
						Wend
						rsview.close()
						%>
						<tr >
							<td colspan=7><strong>Total</strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlhk)),2)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlamnt)),2)%></strong></td>
						</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	