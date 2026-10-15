<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>FB-Plants - FA Matrure</title>
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

dim fccode,fcba

fccode= trim(request("fccode"))
fcba= trim(request("fcba"))

dim fcbaname,CompanyName,documentno,fixedassetcode,period_month,period_year
Dim sqlview,q,rsview 
q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
	set rsview = connStr.Execute(q)
	if not rsview.eof then
	   fcbaname = rsview("fcName")
	   CompanyName=rsview("FCCOMPANYNAME")
	end if
rsview.Close
set rsview = Nothing

q = "Select documentno,fixedassetcode,period_month,period_year from fixedasset_mature_new where documentno='" & fccode & "' and fcba='" &  fcba &"'"
	set rsview = connStr.Execute(q)
	if not rsview.eof then
	   documentno     = rsview("documentno")
	   fixedassetcode = rsview("fixedassetcode")
	   period_month   = rsview("period_month")
	   period_year    = rsview("period_year")
	end if
rsview.Close
set rsview = Nothing

sqlview =   " SELECT b.accountingperiod,b.accountingyear, b.sourcecode,c.fccode,c.fcname,c.job_forfieldstatus,a.locationcode,round(Sum(Nvl(amountdebit,0)+Nvl(amountkredit,0)),2) amount " & _
			" FROM bkmjournaldtl a " & _
			" INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba)" & _
			" INNER JOIN jobmaster c ON (a.jobcode=c.fccode)" & _
			" INNER JOIN fixedasset_mature_new_dtl d on (a.fcba=d.fcba and a.locationcode=d.fieldcode) " & _ 
			" WHERE d.documentno='" & documentno & "' AND a.fcba='" & fcba & "'" & _
			" AND postingflg=1" & _
			" AND job_forfieldstatus IN ('TBM','LC')" & _
			" AND docdate < To_Date('" & period_month & "/01/" & period_year & "','MM/DD/YYYY') " & _ 
			" AND b.accountingperiod <= 12 " & _ 
			" GROUP BY b.accountingperiod,b.accountingyear,b.sourcecode,c.fccode,c.fcname,c.job_forfieldstatus,a.locationcode "  & _ 
			" ORDER BY a.locationcode ,c.fccode,b.sourcecode,b.accountingyear,to_number(b.accountingperiod) "  
			
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
				<div id="subheader" class="centered" style="width:50%; margin:auto;"> 
					<h3>FIXED ASSET - MATURITY</h3>
				</div>
				<div id="subheader" class="left" style="width:50%; margin:auto;"> 
					<h4>PT. <%=CompanyName%></h4>
					<h4><%=fcbaname%></h4>
					<h4>No.Doc   = <%=documentno%> </h4>
					<h4>No.Asset = <%=fixedassetcode%> </h4>
					<h4>Periode  = <%=period_month%>/<%=period_year%> </h4>
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
							<th>Acc.Period</th>
							<th>Source</th>
							<th>Job.Code</th>
							<th>Job.Name</th>
							<th>Job.Status</th>
							<th>Blok</th>
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
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("accountingperiod")%>/<%=rsview("accountingyear")%></td> 							 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("sourcecode")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("fccode")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("fcname")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("job_forfieldstatus")%></td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("locationcode")%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsview("AMOUNT"),2)%></td>
						</tr>
						
						</tbody>
						<%
						ttlamnt = cdbl(ttlamnt) +cdbl(rsview("AMOUNT"))
						rsview.MoveNext
						Wend
						rsview.close()
						%>
						<tr >
							<td colspan=6><strong>Total</strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlamnt)),2)%></strong></td>
						</tr>
						</table>
						
		</DIV>
	</BODY>
</HTML>	