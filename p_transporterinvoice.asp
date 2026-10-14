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
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
<title>ePlaS Reporting - Transporter Invoice </title>
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

 ' End Added by NAF
company     = trim(request("company"))
businessunit= trim(request("businessunit"))
invoice_by= trim(request("invoice_by"))
transportercode= trim(request("transportercode"))
contractno= trim(request("contractno"))
invoicecode= trim(request("invoicecode"))
creationdate_from= trim(request("creationdate_from"))
creationdate_to= trim(request("creationdate_to"))
q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
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

txtSQL = 	"   SELECT fcba_payment," & _
			"          fcba_invoice," & _
			"          invoicecode," & _
			"          salescontractno," & _
			"          creationdate," & _
			"          transportercontractno," & _
			"          transportercode," & _
			"          transportername," & _
			"          no_faktur," & _
			"          fakturdate," & _
			"          delivered_qty_netto," & _
			"          accepted_qty_netto," & _
			"          penaltyrate," & _
			"          amount," & _
			"          vat," & _
			"          pph," & _
			"          amountbeforepph," & _
			"          amountafterpph," & _
			"          vouchercode_payment," & _
			"          pvd_payment_amount," & _
			"          vouchercode_adjustment," & _
			"          pvd_adjustment_amount," & _
			"          (delivered_qty_netto - accepted_qty_netto)" & _
			"             AS var_amount" & _
			"     FROM (  SELECT MAX (fcba_payment) AS fcba_payment," & _
			"                    MAX (fcba_invoice) AS fcba_invoice," & _
			"                    invoicecode," & _
			"                    MAX (salescontractno) AS salescontractno," & _
			"                    MAX (creationdate) AS creationdate," & _
			"                    MAX (transportercontractno) AS transportercontractno," & _
			"                    MAX (transportercode) AS transportercode," & _
			"                    MAX (transportername) AS transportername," & _
			"                    MAX (no_faktur) AS no_faktur," & _
			"                    MAX (fakturdate) AS fakturdate," & _
			"                    SUM (delivered_qty_netto) AS delivered_qty_netto," & _
			"                    SUM (accepted_qty_netto) AS accepted_qty_netto," & _
			"                    MAX (penaltyrate) AS penaltyrate," & _
			"                    SUM (amount) AS amount," & _
			"                    MAX (vat) AS vat," & _
			"                    MAX (pph) AS pph," & _
			"                    SUM (amountbeforepph) AS amountbeforepph," & _
			"                    SUM (amountafterpph) AS amountafterpph," & _
			"                    MAX (vouchercode_payment) AS vouchercode_payment," & _
			"                    SUM (pvd_payment_amount) AS pvd_payment_amount," & _
			"                    MAX (vouchercode_adjustment) AS vouchercode_adjustment," & _
			"                    SUM (pvd_adjustment_amount) AS pvd_adjustment_amount" & _
			"               FROM (SELECT distinct inv.invoice_by AS fcba_payment," & _
			"                            inv.fcba AS fcba_invoice," & _
			"                            inv.invoicecode," & _
			"                            invqty.salescontractno," & _
			"                            inv.creationdate," & _
			"                            inv.transportercontractno," & _
			"                            tr.fccode AS transportercode," & _
			"                            tr.fcname AS transportername," & _
			"                            inv.no_faktur," & _
			"                            inv.fakturdate," & _
			"                            invqty.delivered_qty_netto," & _
			"                            invqty.accepted_qty_netto," & _
			"                            inv.penaltyrate," & _
			"                            inv.amount," & _
			"                            inv.vat," & _
			"                            TO_CHAR (inv.pph) pph," & _
			"                            inv.amountbeforepph," & _
			"                            inv.amountafterpph," & _
			"                            '' AS vouchercode_payment," & _
			"                            0 AS pvd_payment_amount," & _
			"                            '' AS vouchercode_adjustment," & _
			"                            0 AS pvd_adjustment_amount" & _
			"                       FROM transporterinvoice inv" & _
			"                            INNER JOIN transportercontract tc" & _
			"                               ON (    tc.fcba = inv.fcba" & _
			"                                   AND tc.transportercontractno =" & _
			"                                          inv.transportercontractno)" & _
			"                            INNER JOIN transporter tr" & _
			"                               ON (    tr.fccode = tc.transporter" & _
			"                                   AND tr.fcba = tc.fcba)" & _
			"                            INNER JOIN" & _
			"                            (  SELECT trinvdtl.invoicecode," & _
			"                                      trinvdtl.fcba," & _
			"                                      sshdr.salescontractno," & _
			"                                       (trinvdtl.deliveredqty) AS delivered_qty_netto," & _
			"                                       (trinvdtl.acceptedqtybruto) AS accepted_qty_netto" & _
			"                                 FROM transporterinvoicedtl trinvdtl" & _
			"                                      INNER JOIN salesshipmentdtl ssdtl" & _
			"                                         ON (    trinvdtl.fcba = ssdtl.fcba" & _
			"                                             AND trinvdtl.salesshipmentno =" & _
			"                                                    ssdtl.salesshipmentno)" & _
			"                                      INNER JOIN salesshipmenthdr sshdr" & _
			"                                         ON (    ssdtl.fcba = sshdr.fcba" & _
			"                                             AND ssdtl.salesshipmentno =" & _
			"                                                    sshdr.salesshipmentno)" & _
			"                             ) invqty" & _
			"                               ON (    invqty.fcba = inv.fcba" & _
			"                                   AND invqty.invoicecode = inv.invoicecode)" & _
			"                      WHERE inv.invoice_by like '" & invoice_by & "%' " & _
			"                     UNION ALL" & _
			"                       SELECT pvd.fcba AS fcba_payment," & _
			"                              '' AS fcba_invoice," & _
			"                              pvd.agreementcode AS invoicecode," & _
			"                              '' AS salescontractno," & _
			"                              NULL AS creationdate," & _
			"                              '' AS transportercontractno," & _
			"                              '' AS transportercode," & _
			"                              '' AS transportername," & _
			"                              '' AS no_faktur," & _
			"                              NULL AS fakturdate," & _
			"                              0 AS delivered_qty_netto," & _
			"                              0 AS accepted_qty_netto," & _
			"                              0 penaltyrate," & _
			"                              0 amount," & _
			"                              0 vat," & _
			"                              '' pph," & _
			"                              0 amountbeforepph," & _
			"                              0 amountafterpph," & _
			"                              pvd.vouchercode AS vouchercode_payment," & _
			"                              NVL (SUM (pvd.amount), 0) AS pvd_payment_amount," & _
			"                              '' AS vouchercode_adjustment," & _
			"                              0 AS pvd_adjustment_amount" & _
			"                         FROM paymentvoucherdetail pvd" & _
			"                              INNER JOIN paymentvoucher pv" & _
			"                                 ON (    pvd.fcba = pv.fcba" & _
			"                                     AND pvd.vouchercode = pv.vouchercode)" & _
			"                        WHERE     pvd.locationtype IN ('CA')" & _
			"                              AND pv.transtype = '02-PAYMENT'" & _
			"                     GROUP BY pvd.fcba," & _
			"                              pvd.agreementcode," & _
			"                              pvd.vouchercode," & _
			"                              pvd.locationtype" & _
			"                     UNION ALL" & _
			"                       SELECT pvd.fcba AS fcba_payment," & _
			"                              '' AS fcba_invoice," & _
			"                              pvd.agreementcode AS invoicecode," & _
			"                              '' AS salescontractno," & _
			"                              NULL AS creationdate," & _
			"                              '' AS transportercontractno," & _
			"                              '' AS transportercode," & _
			"                              '' AS transportername," & _
			"                              '' AS no_faktur," & _
			"                              NULL AS fakturdate," & _
			"                              0 AS delivered_qty_netto," & _
			"                              0 AS accepted_qty_netto," & _
			"                              0 penaltyrate," & _
			"                              0 amount," & _
			"                              0 vat," & _
			"                              '' pph," & _
			"                              0 amountbeforepph," & _
			"                              0 amountafterpph," & _
			"                              '' AS vouchercode_payment," & _
			"                              0 AS pvd_payment_amount," & _
			"                              pvd.vouchercode AS vouchercode_adjustment," & _
			"                              NVL (SUM (pvd.amount), 0) AS pvd_adjustment_amount" & _
			"                         FROM paymentvoucherdetail pvd" & _
			"                              INNER JOIN paymentvoucher pv" & _
			"                                 ON (    pvd.fcba = pv.fcba" & _
			"                                     AND pvd.vouchercode = pv.vouchercode)" & _
			"                        WHERE     pvd.locationtype IN ('CA')" & _
			"                              AND pv.transtype = '03-JOURNAL'" & _
			"                     GROUP BY pvd.fcba," & _
			"                              pvd.agreementcode," & _
			"                              pvd.vouchercode," & _
			"                              pvd.locationtype)" & _
			"           GROUP BY invoicecode)" 
if  len(trim(creationdate_from)) > 0 then
    txtSQL = txtSQL & " WHERE creationdate >=to_Date('" & creationdate_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(creationdate_from,1)
end if
if  len(trim(creationdate_to)) > 0 then
    txtSQL =txtSQL &  " and creationdate <=to_Date('" & creationdate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(creationdate_to,1)
end if
if  len(trim(invoicecode)) > 0 then
    txtSQL =txtSQL &  " and invoicecode = '"& invoicecode &"'"
end if
if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  " and fcba_invoice in (SELECT fccode from BusinessUnit where fccompanycode='" & company & "' and fccode like '%"& businessunit &"%')" 
end if 
if  len(trim(transportercode)) > 0 then
    txtSQL =txtSQL &  " and  transporter = '"& transportercode &"'"
end if
if  len(trim(contractno)) > 0 then
    txtSQL =txtSQL &  " and transportercontractno = '"& contractno &"'"
end if


txtsql=(txtsql) 

	 txtSQL =txtSQL &  " order by invoicecode, creationdate"
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

 'response.Write(txtsql)
' response.Flush()
Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open txtSQL, connstr
%>
 <body>
		<div id="wrapper">
		<!--#include file="print_excel.asp"-->
 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="31" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4><%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>TRANSPORTER INVOICE</h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="31" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr> 
					<th>Invoice Code</th>
					<th>Voucher Code</th>
					<th>Voucher Adjustment</th>
					<th>Sales Contract No</th>
					<th>Invoice By</th>
					<th>Date</th>
					<th>Transporter Contract</th>
					<th>Transporter </th>
					<th>Transporter Name</th>
					<th>No. Faktur</th>
					<th>Faktur Date</th> 
					<th>Penalty Rate</th> 
					<th>Amount</th>
					<th>VAT</th>
					<th>PPh 23</th>
					<th>Amount Before PPh 23</th>
					<th>Grand Total</th> 
					<th>Total Amount Voucher</th>
					<th>Total Amount Voucher Adj.</th>
					<th>VAR</th>
					<th>Delivery Qty Netto</th>
					<th>Accepted Qty Netto</th> 
				</tr>
                </thead>
               <tbody>
	<%
	if Not rscat.eof then
		v_deliveredqty = 0
		v_acceptedqtybruto = 0
		v_varianceqty = 0
		v_amount = 0
		v_amountbeforepph = 0
		v_amountafterpph = 0
		v_amountbeforeloss = 0
		v_amountafterloss = 0
		v_penaltyamount = 0
		v_penaltyqty = 0
		v_amountvar = 0
		v_pvd_totalamount = 0 
		v_pvd_adjustment_amount = 0
		v_delivered_qty_netto=0
		v_accepted_qty_netto= 0
	While Not rscat.EOF


	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
				<a href="javascript:launchwin('p_transporterinvoicedet.Asp?INVOICECODE=<%=rscat("INVOICECODE")%>&creationdate=<%=rscat("creationdate")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1"><%=rscat("INVOICECODE")%></font> 
				</a>
			</td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vouchercode_payment")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vouchercode_adjustment")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("salescontractno")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba_payment")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("creationdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transportercontractno")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transportercode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transportername")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("no_faktur")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fakturdate")%></td>   
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("penaltyrate"))),2)%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amount"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("vat"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("pph"))),0)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amountbeforepph"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amountafterpph"))),2)%></td> 
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pvd_payment_amount"))),2)%></td> 
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pvd_adjustment_amount"))),2)%></td> 
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("var_amount"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("delivered_qty_netto"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("accepted_qty_netto"))),2)%></td>
		</tr>
	  <%
	   ' v_deliveredqty = v_deliveredqty+ cdbl(to_nol(rscat("deliveredqty")))
	   ' v_acceptedqtybruto = v_acceptedqtybruto + cdbl(to_nol(rsCat("acceptedqtybruto")))
	   ' v_varianceqty = v_varianceqty+ cdbl(to_nol(rsCat("varianceqty")))
	   v_amount = v_amount+ cdbl(to_nol(rsCat("amount")))
	   v_amountbeforepph = v_amountbeforepph + cdbl(to_nol(rsCat("amountbeforepph")))
	   v_amountafterpph = v_amountafterpph + cdbl(to_nol(rsCat("amountafterpph")))
	   v_amountvar = v_amountvar + cdbl(to_nol(rsCat("var_amount")))
	   v_pvd_totalamount = v_pvd_totalamount+Trim(cdbl(to_nol(rsCat("pvd_payment_amount"))))
	   v_pvd_adjustment_amount = v_pvd_adjustment_amount+Trim(cdbl(to_nol(rsCat("pvd_adjustment_amount"))))
	   v_delivered_qty_netto = v_delivered_qty_netto+Trim(cdbl(to_nol(rsCat("delivered_qty_netto"))))
	   v_accepted_qty_netto = v_accepted_qty_netto+Trim(cdbl(to_nol(rsCat("accepted_qty_netto"))))
	   ' v_amountbeforeloss = v_amountbeforeloss+ cdbl(to_nol(rsCat("amountbeforeloss")))
	   ' v_amountafterloss = v_amountafterloss+ cdbl(to_nol(rsCat("amountafterloss")))
	   ' v_penaltyamount = v_penaltyamount+ cdbl(to_nol(rsCat("penaltyamount")))
	   ' v_penaltyqty = v_penaltyqty+ cdbl(to_nol(rsCat("penaltyqty")))
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='28'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="12"> <font style="font-weight: bold;" face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font style="font-weight: bold;" face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amount)),2)%></td>
		<td colspan="2"></td>			
		<td align="Right" style="<%=cmdstyle%>"> <font style="font-weight: bold;" face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amountbeforepph)),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font style="font-weight: bold;" face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amountafterpph)),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font style="font-weight: bold;" face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_pvd_totalamount)),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font style="font-weight: bold;" face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_pvd_adjustment_amount)),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font style="font-weight: bold;" face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amountvar)),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font style="font-weight: bold;" face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_delivered_qty_netto)),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font style="font-weight: bold;" face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_accepted_qty_netto)),2)%></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->	 