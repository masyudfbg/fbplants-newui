<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting -LHD Listing</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
function() {
new xTableHeaderFixed('xthf-site', window);
}, false
);
</script>
<style type="text/css">
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
</head>

<%
Dim Columnx : Columnx = 0
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim tbl,doc,filtercon,filterag,filterfcba,pcompany,fcba,bkmfrom,bkmto,contractorname
Dim q,agreementcode
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"


   pcompany 	= Request("pcompany")
    fcba 		= Request("pfcba")
    bkmfrom 	= Request("bkmfrom")
    bkmto   	= Request("bkmto")
    contractorname   	= Request("contractorcode")
 	agreementcode   	= Request("agreementcode")

	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
	' response.write q
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
		fcbaname = rsCat("fcName")
		CompanyName=rsCat("FCCOMPANYNAME")
		MgrName = rsCat("FCMANAGER")
		KTUName = rsCat("FCKTU")
		CropName=rsCat("FCCROP")
		FcbaType=rsCat("fcType")
	end if


	if trim(agreementcode) <> "all" then
		filtercon = " AND b.contractorcode='" & contractorname & "'"
	end if

	if trim(contractorname) <> "all" then
		filterag = " AND a.agreementcode='" & agreementcode & "'"
	end if


	if trim(fcba) = "-" then 
		filterfcba = " AND a.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') "
	else
	    filterfcba = " AND a.FCBA ='" & fcba & "' "
	end if



sqlview1 =  " SELECT " & _
" b.fcba,b.contractorcode,c.fcname contractorname, " & _
" b.agreementcode,b.agreementdate,b.startdate,b.finishdate,a.jobcode,d.fcname jobname,a.lineno,nvl(a.qty,0) qtycontract,nvl(a.rate,0) ratecontract,(nvl(a.qty,0)*nvl(a.rate,0))amountcontract ,Nvl(a.remarks,'-') remarks, " & _
" e.unitkerja,e.dateplan,NVL(e.qtyplan,0)qtyplan,nvl(e.amountplan,0)amountplan, " & _
" f.datefinal,NVL(f.qtyfinal,0)qtyfinal, " & _
" g.invoicecode,g.transdate,nvl(g.rpinvoice,0) rpinvoice" & _
" FROM " & _
" contractagreementdetail a " & _
" INNER JOIN contractagreement b ON ( a.agreementcode = b.agreementcode AND a.fcba=b.fcba ) " & _
" INNER JOIN contractor c ON (b.contractorcode=c.fccode AND b.fcba = c.fcba) " & _
" INNER JOIN jobmaster d ON (a.jobcode=d.fccode) " & _
" LEFT JOIN " & _
" (SELECT fcba,agreementcode,agreementlineno,wipdate dateplan,Nvl(chargeunit,fcba) unitkerja,Sum(qty) qtyplan,Sum(wipamt) amountplan " & _
" FROM workinprogressplan " & _
" GROUP BY fcba,agreementcode,agreementlineno,wipdate,Nvl(chargeunit,fcba) " & _
" )e  ON ( a.agreementcode = e.agreementcode AND a.fcba=e.fcba AND a.lineno=e.agreementlineno) " & _
" LEFT JOIN " & _
" ( " & _
" SELECT fcba,agreementcode,agreementlineno,wipdate datefinal,Sum(qty) qtyfinal " & _
" FROM workinprogressfinal " & _
" GROUP BY fcba,agreementcode,agreementlineno,wipdate " & _
" )f ON ( a.agreementcode = f.agreementcode AND a.fcba=f.fcba AND a.lineno=f.agreementlineno) " & _
" LEFT JOIN " & _
" ( SELECT x.invoicecode,transdate,x.agreementcode,agreementlineno,x.fcba,Sum(amount) rpinvoice " & _
"  FROM contractorinvoicedtl x " & _
"  INNER JOIN  contractorinvoicehdr y ON ( x.invoicecode = y.invoicecode AND x.fcba=y.fcba AND x.agreementcode=y.agreementcode ) " & _
"  GROUP BY x.invoicecode,transdate,x.agreementcode,agreementlineno,x.fcba" & _
") g ON (f.agreementcode=g.agreementcode AND f.fcba=g.fcba AND f.agreementlineno=g.agreementlineno) " & _
" WHERE e.dateplan between to_date('" & bkmfrom & "','mm/dd/yyyy') and  to_date('" & bkmto & "','mm/dd/yyyy')" & _
filterfcba & _ 
filtercon & _ 
filterag & _ 
" ORDER BY d.fccode,a.agreementcode,a.lineno"
	Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open sqlview1, connstr
'response.Write(sqlview1)

%>

<body>
	<div id="wrapper">
        <!--#include file="print_excel_blob.asp"-->
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					
					<th colspan="18" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=CompanyName%></h4>

								<h5>Estate: <%=fcba%></h5>
							</div>
							
						</div>
					</th>
				</tr>
						
				<tr>
						
						<th align="left" colspan="18" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">MONITORING WORK IN PROGRESS PLAN</th>
						
				</tr>
				<tr>
						
						<th align="left" colspan="18" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">Periode : <%=bkmfrom%> s.d <%=bkmto%></th>
				</tr>
						


	<%	
	
	NO=1%>
	<tr>
			<th rowspan="2">No</th>
		    <th rowspan="2">FCBA</th>
			<th rowspan="2">KONTRAKTOR</th>
			<th rowspan="2">JOB</th>
			<th colspan="5">KONTRAK</th>
			<th colspan="3">OPR-WIP</th>
			<th colspan="3">INVOICE WIP</th>
			<th colspan="3">BAYAR</th>
			
	</tr>
	<tr>
			<th>Agrement Code</th>			
			<th colspan="">Tanggal</th>
			<th colspan="">Qty</th>
			<th colspan="">Harga</th>
			<th colspan="">Amount</th>
			<th colspan="">Qty</th>
			<th colspan="">Harga</th>
			<th rowspan="">Amount</th>
			<th colspan="">Invoice Code</th>
			<th rowspan="">Tgl Invoice</th>
			<th rowspan="">Amount</th>
			<th colspan="">Qty</th>
			<th colspan="">Harga</th>
			<th colspan="">Amount</th>
		</tr>
		
</thead>
<tbody>
	<%
	
	   if Not rsCat.eof then
		
			no=1
		
		While Not rsCat.EOF

			
			%>
		<tr>
			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("fcba")%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("contractorname")%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("jobcode") & " - " & rsCat("jobname")%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("agreementcode")%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("agreementdate")%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtycontract"))%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("ratecontract"))%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amountcontract"))%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("qtyplan"))%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("ratecontract"))%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("amountplan"))%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("invoicecode")%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("transdate")%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("rpinvoice"))%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">0</td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">0</td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">0</td>					
			
		</tr>
	  <%
	  
		

	  rsCat.MoveNext

	  no=no+1
	Wend
	rsCat.close()
	rsCat=null
	END IF
%>

		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyle%>" colspan= "18"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
			
		</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</BODY>
</HTML>	