<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - RKB PDF</title>
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
 window.print();
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

dim tahun,bulan,company,businessunit,typerkb,filter

tahun = Request("tahun")
bulan = Request("bulan")
company = Request("company")
businessunit = Request("fcba")
typerkb = Request("type")
	


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
if trim(businessunit) = "-" then 
		filter = " AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & company & "') "
	else
	    filter = " AND FCBA ='" & businessunit & "' "
	end if
sqlview = "Select FCBA,RKBNO,PERIOD_YEAR||'/'||PERIOD_MONTH period,SECTIONCODE," & _
	           " case document_classification " & _ 
			   " when '501' then 'Perawatan' " & _ 
			   " when '503' then 'Panen' " & _ 
			   " when '502' then 'Vehicle' " & _ 
			   " when '504' then 'Wip' End type, sum(NVL(TOTUPAH,0)) TOTUPAH ,sum(NVL(TOTBAHAN,0)) TOTBAHAN,SUM(TOTVRA) TOTVRA ,sum(TOTWIP) TOTWIP " & _
			   " From rkbhdr " & _
			   " WHERE period_year=" & tahun & " " & _ 
			   " AND period_month=" & bulan & " " & _ 
			   " AND document_classification like '" & trim(typerkb) & "%'" & _
			   filter & _
			   " GROUP BY FCBA,RKBNO,PERIOD_YEAR||'/'||PERIOD_MONTH,SECTIONCODE,document_classification " & _ 
			   " ORDER BY FCBA,period,document_classification,RKBNO"
set rsview = connStr.Execute(sqlview)


 'response.Write(sqlview)
'response.Flush()
%>
	<body>
		<div id="wrapper">
        
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
					<h2>LISTING RKB</h2>
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
							
							<th>FCBA</th>
							<th>Type</th>
							<th>RKB No</th>
							<th>RKB Period</th>
							<th>Section</th>
							<th>RKB Upah</th>  
							<th>RKB Material</th>
													
						</tr>
                         </thead>

						 <%
						 	dim ttlupah,ttlmaterial,ttlamount
							ttlupah = 0
							ttlmaterial = 0
							
							ttlamount = 0
							While Not rsview.EOF

						 %>

                         <tbody>
						<tr>
														 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("FCBA")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("type")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("RKBNO")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("period")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("SECTIONCODE")%></td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("TOTUPAH")%></td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("TOTBAHAN")%> </td>
							
						</tr>
						
						</tbody>
						<%
							ttlupah = ttlupah +cdbl(rsview("TOTUPAH"))
							ttlmaterial = ttlmaterial+cdbl(rsview("TOTBAHAN"))
		
							ttlamount = cdbl(ttlupah)+cdbl(ttlmaterial)
						'ttlamnt = ttlamnt +cdbl(rsview("TOTAL"))
						
						rsview.MoveNext
						Wend
						rsview.close()
						%>
						<tr >
							<td colspan=5><strong>Total</strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlupah)),0)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlmaterial)),0)%></strong></td>
							
						</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	