<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - BKM PDF</title>
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

dim tglawal,tglakhir,pilihbkm,section,posting

tglawal= trim(request("tglawal"))
tglakhir= trim(request("tglakhir"))
	pilihbkm   = Request("pilihbkm")
	section = Request("section")
	posting = Request("posting")



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
sqlview = "Select h.BKMNO, h.BKMDATE CREATEDATE,h.SECTIONCODE,h.GANGCODE,nvl(h.VEHICLECODE,'-') VEHICLECODE, h.TOTUPAH TOTALBKMUPAH,h.TOTBAHAN TOTALBKMMATERIAL,h.TOTVRA TOTALBKMVEHICLE,h.TOTWIP TOTALBKMWIP,h.DOCUMENT_CLASSIFICATION " & _
	"    ,authorizestructure_api.get_template_desc(h.DOCUMENT_CLASSIFICATION) dcname,decode(h.hujan,'0','-','YA') hujan,decode(h.POSTINGFLG,'1','YA','TIDAK') POSTINGFLG" & _
	"    ,h.TOTUPAH + h.TOTBAHAN + h.TOTVRA + h.TOTWIP TOTAL " & _	
    "    From   BKMHDR h WHERE h.BKMDATE between to_date('" & tglawal & "','mm/dd/yyyy') and to_date('" & tglakhir & "','mm/dd/yyyy') and h.fcba='" & businessunit & "' " & _ 
    "    AND document_classification Like '" & pilihbkm & "%' " & _ 
	"    AND SECTIONCODE like '%" & section & "' " & _
	"    AND POSTINGFLG like '%" & posting & "' " & _
	"    ORDER BY h.BKMDATE,h.SECTIONCODE,h.GANGCODE,h.VEHICLECODE,h.DOCUMENT_CLASSIFICATION"
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
					<h2>LISTING BKM</h2>
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
							
							<th>BKM No</th>
							<th>BKM Date</th>
							<th>BKM Type</th>
							<th>Section</th>
							<th>Gang</th>
							<th>Vehicle</th>
							<th>BKM Upah</th>	  
							<th>BKM Material</th>
							<th>BKM Vehicle</th>
							<th>BKM Wip</th>
							<th>Total</th>
							
						</tr>
                         </thead>

						 <%
						 	dim ttlupah,ttlmaterial,ttlveh,ttlwip,ttlamount
							ttlupah = 0
							ttlmaterial = 0
							ttlveh = 0
							ttlwip = 0
							ttlamount = 0
							While Not rsview.EOF

						 %>

                         <tbody>
						<tr>
														 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("BKMNO")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("CREATEDATE")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("dcname")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("SECTIONCODE")%></td> 
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("GANGCODE")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("VEHICLECODE")%></td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("TOTALBKMUPAH")%></td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("TOTALBKMMATERIAL")%> </td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("TOTALBKMVEHICLE")%> </td>
							<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsview("TOTALBKMWIP")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsview("TOTAL"),2)%></td>
						</tr>
						
						</tbody>
						<%
							ttlupah = ttlupah +cdbl(rsview("TOTALBKMUPAH"))
							ttlmaterial = ttlmaterial+cdbl(rsview("TOTALBKMMATERIAL"))
							ttlveh = ttlveh+cdbl(rsview("TOTALBKMVEHICLE"))
							ttlwip = ttlwip+cdbl(rsview("TOTALBKMWIP"))
							ttlamount = cdbl(ttlupah)+cdbl(ttlmaterial)+cdbl(ttlveh)+cdbl(ttlwip)
						'ttlamnt = ttlamnt +cdbl(rsview("TOTAL"))
						
						rsview.MoveNext
						Wend
						rsview.close()
						%>
						<tr >
							<td colspan=6><strong>Total</strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlupah)),2)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlmaterial)),2)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlveh)),2)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlwip)),2)%></strong></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><strong><%=FormatNumber(cdbl(to_nol(ttlamount)),2)%></strong></td>
						</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	