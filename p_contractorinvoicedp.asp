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

v_txtINVOICEDATE_From= trim(request("CREATIONDATE_From"))
v_txtINVOICEDATE_To= trim(request("CREATIONDATE_To"))


businessunit= trim(request("businessunit"))

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
sqlview = "select h.invoicecode,h.transdate,h.invoice_by,h.agreementcode,h.contractorcode,c.fcname,h.qty,h.rate,h.amount,h.dp,h.remain " & _
" from CONTRACTORINVOICEDP h,CONTRACTOR c " & _
" where h.CONTRACTORCODE=c.FCCODE and h.FCBA = c.FCBA  " & _
" and h.transdate between TO_DATE('" & v_txtINVOICEDATE_From & "','mm/dd/yyyy') and to_date('" & v_txtINVOICEDATE_To & "','mm/dd/yyyy') " & _
" and h.fcba='" & businessunit & "'"
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
         <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>CONTRACTOR INVOICE DP</h2>
				</div>
			</div>
				</th>
			</tr>
				<!--<div class="printhide">-->
						<tr>
							<th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
						
						
						</tr> 
						<tr>
							<th colspan="1"></th>
							<th>Date</th>
							<th>Invoice Code</th>
							<th>Invoice By</th>
							<th>Agreement Code</th>
							<th>Contractor</th>	  
							<th>Qty</th>
							<th>Rate</th>
							<th>Amount</th>
							<th>DP</th>
							<th>Remain</th>
						</tr>
                         </thead>

						 <%
							While Not rsview.EOF

						 %>

                         <tbody>
						<tr>
							<td>
								<a href="printcontractorinvdp.asp?fcba=<%=businessunit%>&invoicecode=<%=rsview("invoicecode")%>">
								<img src="IMAGES/print.gif" alt="Print Data : <%=rsview(1)%>" border="0">
								</a>
							</td>
							 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("transdate")%></td> 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("invoicecode")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("invoice_by")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("agreementcode")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("fcname")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("qty")%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("rate"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("amount"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("dp"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsview("remain"))),2)%></td> 
						</tr>
						</tbody>
						<%
						rsview.MoveNext
						Wend
						rsview.close()
						%>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	