<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - WIP Monitoring </title>
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
dim ttlcontract,ttldailyqty,ttldailyamount,ttlfinalqty,ttlamountfinal
dim grandqtykon,grandamountkon,grandqtydaily,grandamountdaily
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


	if trim(contractorname) <> "all" then
		filtercon = " AND contractorcode='" & contractorname & "'"
	end if

	if trim(agreementcode) <> "all" then
		filterag = " AND agreementcode='" & agreementcode & "'"
	end if


	if trim(fcba) = "-" then 
		filterfcba = " AND FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') "
	else
	    filterfcba = " AND FCBA ='" & fcba & "' "
	end if



sqlview1 =  "SELECT fcba, agreementcode, contractorcode, contractorname, agreementlineno,  jobcode, jobname,  plan_qty, " & _
" plan_unit, plan_output, rate, final_qty, final_unit, final_output, (rate * final_qty) total,agreement_qty,contractdate,remarks,wipdate " & _
" FROM ( " & _
"          SELECT wf.fcba, wf.agreementcode, ca.contractorcode, ca.contractorname, wf.agreementlineno, wf.jobcode, wf.jobname, " & _
"          NVL (wf.qty, 0) plan_qty, NVL (wf.unit, 0) plan_unit, NVL (wf.output, 0) plan_output, " & _
"          NVL (ca.rate, 0) rate, NVL (final_qty, 0) final_qty, NVL (final_unit, 0) final_unit, " & _
"          NVL (final_output, 0) final_output, NVL (ca.qty, 0) agreement_qty, NVL (final_qty, 0) progress_qty,contractdate,remarks,wipdate " & _
"          FROM (  " & _
"                    SELECT wp.fcba, wp.agreementcode, wp.agreementlineno,  wpp.jobcode,  jb.fcname jobname, " & _
"                    NVL (SUM (wp.qty), 0) qty, NVL (SUM (wpp.unit), 0) unit, " & _
"                    NVL (SUM (wpp.output), 0) output, NVL (SUM (wf.qty), 0) final_qty, " & _
"                    NVL (SUM (wf.unit), 0) final_unit, NVL (SUM (wf.output), 0) final_output,wpp.wipdate " & _
"                    FROM workinprogress wp " & _
"                    INNER JOIN workinprogressplan wpp ON ( wp.agreementcode = wpp.agreementcode AND wp.fcba = wpp.fcba " & _
"                    AND wp.agreementlineno = wpp.agreementlineno AND wp.agreementfinalno = wpp.agreementfinalno) " & _
"                    INNER JOIN workinprogressfinal wf ON ( wp.fcba = wf.fcba AND wp.agreementcode = wf.agreementcode " & _
"                    AND wp.agreementlineno = wf.agreementlineno AND wp.agreementfinalno = wf.agreementfinalno) " & _
"                    INNER JOIN job jb ON (jb.fccode = wpp.jobcode AND jb.fcba = wpp.fcba) " & _
"                    GROUP BY wp.fcba, wp.agreementcode, wp.agreementlineno, wpp.jobcode,  jb.fcname,wpp.wipdate " & _
"          ) wf " & _
"          INNER JOIN ( " & _
"                        SELECT ca.fcba, ca.agreementcode, ca.contractorcode, cr.fcname contractorname, cad.lineno, NVL (cad.rate, 0) rate, NVL ( (cad.qty), 0) qty,ca.AGREEMENTDATE contractdate, cad.remarks " & _
"                        FROM contractagreementdetail cad " & _
"                        INNER JOIN contractagreement ca ON ( ca.agreementcode = cad.agreementcode AND ca.fcba = cad.fcba) " & _
"                        INNER JOIN contractor cr ON ( cr.fccode = ca.contractorcode AND cr.fcba = ca.fcba) " & _
"                     ) ca ON ( wf.fcba = ca.fcba AND wf.agreementcode = ca.agreementcode AND wf.agreementlineno = ca.lineno)" & _
"  ) WHERE wipdate between to_date('" & bkmfrom & "','mm/dd/yyyy') and  to_date('" & bkmto & "','mm/dd/yyyy') " & _
filterfcba & _  
filtercon & _  
filterag & _  

"  ORDER BY contractorcode, agreementcode, LPAD (agreementlineno, 3, 0)"
'response.Write(sqlview1)
	Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open sqlview1, connstr


%>

<body>
	<div id="wrapper">
        <!--#include file="print_excel_blob.asp"-->
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					
					<th colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=CompanyName%></h4>

								<h5>Estate: <%=fcba%></h5>
							</div>
							
						</div>
					</th>
				</tr>
						
				<tr>
						
						<th align="left" colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">LAPORAN WORK IN PROGRESS</th>
						
				</tr>
				<tr>
						
						<th align="left" colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">Periode : <%=bkmfrom%> s.d <%=bkmto%></th>
				</tr>
						


	<%	
	
	NO=1%>
	<tr>
			<th rowspan="2">No</th>
		    <th rowspan="2">FCBA</th>
			<th rowspan="2">KONTRAKTOR</th>
			<th colspan="7">KONTRAK</th>
			<th colspan="3">DAILY-WIP</th>
			<th colspan="1">FINAL-WIP</th>
			<th rowspan="2">ACTION</th>
			
	</tr>
	<tr>
			<th>Agrement Code</th>			
			<th>Remark</th>			
			<th colspan="">Tanggal</th>
			<th >JOB</th>
			<th colspan="">Qty</th>
			<th colspan="">Harga</th>
			<th colspan="">Amount</th>
			<th colspan="">Qty</th>
			<th colspan="">Harga</th>
			<th rowspan="2">Amount</th>
			<th rowspan="">Qty Final</th>
			
		</tr>
	
</thead>
<tbody>
	<%
	
	   if Not rsCat.eof then
		
			no=1
		ttlcontract = 0 
		ttldailyqty = 0 
		ttldailyamount = 0 
		ttlfinalqty = 0 
		ttlamountfinal = 0
		While Not rsCat.EOF
				amountcontract = cdbl(rsCat("agreement_qty"))*cdbl(rsCat("rate"))
			
			%>
		<tr>
			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("fcba")%><input type="hidden" name="fcba" id="fcba<%=no%>" value="<%=rsCat("fcba")%>"></input></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("contractorname")%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><input type="hidden" name="agreementcode" id="agreementcode<%=no%>" value="<%=rsCat("agreementcode")%>"></input><%=rsCat("agreementcode")%></td>					
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("remarks")%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("contractdate")%></td>	
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("jobcode") & " - " & rsCat("jobname")%></td>							
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("agreement_qty"))%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("rate"))%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(amountcontract)%></td>					
			<td colspan="" align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("plan_qty"))%></td>					
			<td colspan="" align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("rate"))%></td>					
			<td colspan="" align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("total"))%></td>					
			<td colspan="" align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("final_qty"))%></td>					
			<td colspan="" align="center" width="100px" style="<%=cmdstyle%>"><font face="Verdana" size="1">
				
				<input type="button" value="View Invoice" id="btninvoice" onclick="getViewinv('<%=no%>')">
			</td>					

		</tr>
		

	  <%
	  
		ttldailyqty = cdbl(ttldailyqty) + cdbl(rsCat("plan_qty"))
		ttldailyamount = cdbl(ttldailyamount) + cdbl(rsCat("total"))
		ttlfinalqty = cdbl(ttlfinalqty) + cdbl(rsCat("final_qty"))
		
	  rsCat.MoveNext

	  no=no+1
	Wend
	rsCat.close()
	rsCat=null
	END IF
%>

		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyle%>" colspan= "10"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
			<td align="Center" style="<%=cmdstyle%>" colspan= ""><font face="Verdana" size="1"><strong><%=FormatNumber(ttldailyqty)%></strong></td>
			<td align="Center" style="<%=cmdstyle%>" colspan= ""><font face="Verdana" size="1"><strong></strong></td>
			<td align="Center" style="<%=cmdstyle%>" colspan= ""><font face="Verdana" size="1"><strong><%=FormatNumber(ttldailyamount)%></strong></td>
			<td align="Center" style="<%=cmdstyle%>" colspan= ""><font face="Verdana" size="1"><strong><%=FormatNumber(ttlfinalqty)%></strong></td>
			
		</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</BODY>
</HTML>	
<script>
	function getViewinv(i)
	{
		var agreementcode = $('#agreementcode' + i).val();
		var fcba = $('#fcba' + i).val();
		window.open('p_contractorinvoice_wip.asp?agreementcode=' + agreementcode + '&fcba=' + fcba, '_blank');
	}
</script>