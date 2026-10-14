<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Invoice STO</title>
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

v_txtINVOICEDATE_From= trim(request("INVOICEDATE_From"))
v_txtINVOICEDATE_To= trim(request("INVOICEDATE_To"))
v_txtTRANSPOTERCODE= trim(request("TRANSPOTERCODE"))
v_txtINVOICECODE= trim(request("INVOICECODE"))
businessunit = trim(request("businessunit"))
company     = trim(request("company"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "

set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
   MgrName = rsCat("FCMANAGER")
   KTUName = rsCat("FCKTU")
   CropName=rsCat("FCCROP")
   FcbaType=rsCat("fcType")
else
    nSpanTotal = nSpanTotal - 1 
end if

	txtSQL= " SELECT fcba_payment," & _
			"        fcba_invoice," & _
			"        invoicecode," & _
			"        refinvoice," & _
			"        invoicedate," & _
			"        transportercode," & _
			"        transportername," & _
			"        pocode," & _
			"        stono," & _
			"        no_faktur," & _
			"        fakturdate," & _
			"        amount," & _
			"        pph," & _
			"        amountpph," & _
			"        inv_totalamount," & _
			"        vouchercode_payment," & _
			"        pvd_payment_amount," & _
			"        vouchercode_adjustment," & _
			"        pvd_adjustment_amount," & _
			"        (inv_totalamount - pvd_payment_amount - pvd_adjustment_amount)" & _
			"           AS var_amount" & _
			"   FROM (  SELECT MAX (fcba_payment) AS fcba_payment," & _
			"                  MAX (fcba_invoice) AS fcba_invoice," & _
			"                  invoicecode," & _
			"                  MAX (refinvoice) AS refinvoice ," & _
			"                  MAX (invoicedate) AS invoicedate," & _
			"                  MAX (transportercode) AS transportercode," & _
			"                  MAX (transportername) AS transportername," & _
			"                  MAX (pocode) AS pocode," & _
			"                  MAX (stono) AS stono," & _
			"                  MAX (no_faktur) AS no_faktur," & _
			"                  MAX (fakturdate) AS fakturdate," & _
			"                  SUM (amount) AS amount," & _
			"                  MAX (pph) AS pph," & _
			"                  SUM (amountpph) AS amountpph," & _
			"                  SUM (totalamount) AS inv_totalamount," & _
			"                  MAX (vouchercode_payment) AS vouchercode_payment," & _
			"                  SUM (pvd_payment_amount) AS pvd_payment_amount," & _
			"                  MAX (vouchercode_adjustment) AS vouchercode_adjustment," & _
			"                  SUM (pvd_adjustment_amount) AS pvd_adjustment_amount" & _
			"             FROM (SELECT inv.invoice_by AS fcba_payment," & _
			"                          inv.fcba AS fcba_invoice," & _
			"                          inv.invoicecode," & _
			"                          inv.refinvoice," & _
			"                          inv.invoicedate," & _
			"                          inv.transportercode," & _
			"                          tr.fcname AS transportername," & _
			"                          inv.pocode," & _
			"                          inv.stono," & _
			"                          inv.no_faktur," & _
			"                          inv.fakturdate," & _
			"                          inv.amount," & _
			"                          inv.pph," & _
			"                          inv.amountpph," & _
			"                          inv.totalamount," & _
			"                          '' AS vouchercode_payment," & _
			"                          0 AS pvd_payment_amount," & _
			"                          '' AS vouchercode_adjustment," & _
			"                          0 AS pvd_adjustment_amount" & _
			"                     FROM invoicesto inv" & _
			"                          INNER JOIN transporter tr" & _
			"                             ON (    inv.fcba = tr.fcba" & _
			"                                 AND inv.transportercode = tr.fccode)" & _
			"                    WHERE inv.fcba in (SELECT fccode from BusinessUnit where fccompanycode='" & company & "' and fccode like '%"& businessunit &"%')" & _
			"                   UNION ALL" & _
			"                     SELECT pvd.fcba AS fcba_payment," & _
			"                            '' AS fcba_invoice," & _
			"                            pvd.agreementcode AS invoicecode," & _
			"                          	 '' AS refinvoice," & _
			"                            NULL AS invoicedate," & _
			"                            '' AS transportercode," & _
			"                            '' AS transportername," & _
			"                            '' AS pocode," & _
			"                            '' AS stono," & _
			"                            '' AS no_faktur," & _
			"                            NULL AS fakturdate," & _
			"                            0 AS amount," & _
			"                            0 AS pph," & _
			"                            0 AS amountpph," & _
			"                            0 AS totalamount," & _
			"                            pvd.vouchercode AS vouchercode_payment," & _
			"                            NVL (SUM (pvd.amount), 0) AS pvd_payment_amount," & _
			"                            '' AS vouchercode_adjustment," & _
			"                            0 AS pvd_adjustment_amount" & _
			"                       FROM paymentvoucherdetail pvd" & _
			"                            INNER JOIN paymentvoucher pv" & _
			"                               ON (    pvd.fcba = pv.fcba" & _
			"                                   AND pvd.vouchercode = pv.vouchercode)" & _
			"                      WHERE     pvd.locationtype IN ('CA')" & _
			"                            AND pv.transtype = '02-PAYMENT'" & _
			"                   GROUP BY pvd.fcba," & _
			"                            pvd.agreementcode," & _
			"                            pvd.vouchercode," & _
			"                            pvd.locationtype" & _
			"                   UNION ALL" & _
			"                     SELECT pvd.fcba AS fcba_payment," & _
			"                            '' AS fcba_invoice," & _
			"                            pvd.agreementcode AS invoicecode," & _
			"                          	 '' AS refinvoice," & _
			"                            NULL AS invoicedate," & _
			"                            '' AS transportercode," & _
			"                            '' AS transportername," & _
			"                            '' AS pocode," & _
			"                            '' AS stono," & _
			"                            '' AS no_faktur," & _
			"                            NULL AS fakturdate," & _
			"                            0 AS amount," & _
			"                            0 AS pph," & _
			"                            0 AS amountpph," & _
			"                            0 AS totalamount," & _
			"                            '' AS vouchercode_payment," & _
			"                            0 AS pvd_payment_amount," & _
			"                            pvd.vouchercode AS vouchercode_adjustment," & _
			"                            NVL (SUM (pvd.amount), 0) AS pvd_adjustment_amount" & _
			"                       FROM paymentvoucherdetail pvd" & _
			"                            INNER JOIN paymentvoucher pv" & _
			"                               ON (    pvd.fcba = pv.fcba" & _
			"                                   AND pvd.vouchercode = pv.vouchercode)" & _
			"                      WHERE     pvd.locationtype IN ('CA')" & _
			"                            AND pv.transtype = '03-JOURNAL'" & _
			"                   GROUP BY pvd.fcba," & _
			"                            pvd.agreementcode," & _
			"                            pvd.vouchercode," & _
			"                            pvd.locationtype)" & _
			"         GROUP BY invoicecode)" 

 
if  len(trim(v_txtINVOICEDATE_From)) > 0 then
    txtSQL = txtSQL & " WHERE INVOICEDATE >=to_Date('" & v_txtINVOICEDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtINVOICEDATE_From,1)
end if
if  len(trim(v_txtINVOICEDATE_To)) > 0 then
    txtSQL =txtSQL &  " and INVOICEDATE <=to_Date('" & v_txtINVOICEDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " to " & FormatDateTime(v_txtINVOICEDATE_To,1)
end if
if  len(trim(v_txtINVOICECODE)) > 0 then
    txtSQL =txtSQL &  " and invoicecode='" & v_txtINVOICECODE & "'"
	txtSelect = txtSelect & " Inv. Code " & v_txtINVOICECODE
end if 
if  len(trim(v_txtTRANSPOTERCODE)) > 0 then
    txtSQL =txtSQL &  " and transportercode ='" & v_txtTRANSPOTERCODE & "'"
	txtSelect = txtSelect & " Transporter Code " & v_txtTRANSPOTERCODE
end if 

	txtSQL = txtSQL & " ORDER BY invoicecode, invoicedate"

' txtsql=(txtsql) & " group by Inv.INVOICECODE, Inv.INVOICEDATE, Inv.SupplierCode, InvDet.ITEMCODE,Stock.FcName  order by Inv.INVOICECODE,Inv.INVOICEDATE,Inv.SupplierCode,InvDet.ITEMCODE"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="19" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" class="centered" style="width:50%; margin:auto;"> 
					<h2>INVOICE STO </h2>
				</div>
			</div>
				</th>
			</tr>
				<!--<div class="printhide">-->
						<tr>
							<th colspan="19" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
							<%	
							Set rsCat = connStr.Execute(txtSQL)
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
	
						<tr>
							<th>Date</th>
							<th>Invoice Code</th>
							<th>Ref. Invoice</th>
							<th>Voucher Code Payment</th>
							<th>Voucher Code Adjustment</th>	
							<th>Invoice By</th>
							<th>Transporter Code</th>
							<th>Transporter Name</th>
							<th>PO Code</th>
							<th>STO No</th>
							<th>No. Faktur</th>
							<th>Faktur Date</th>
							<th>Amount</th>
							<th>PPh</th>
							<th>Amount PPh</th>
							<th>Total Amount Invoice</th>
							<th>Total Amount Voucher</th>
							<th>Total Amount Voucher Adj.</th>
							<th>VAR Amount</th> 
						</tr>
                         </thead>
                         <tbody>
							<% 
							v_amountpph=0
							v_amount=0
							v_totalamount=0
							v_pvd_payment_amount=0
							v_pvd_adjustment_amount=0
							v_var_amount=0 
							While Not rsCat.EOF%>
						<tr>
							 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoicedate")%></td>
							<td align="center" style="<%=cmdstyle%>">
								<a href="javascript:launchwin('p_invoicestodet.Asp?invoicecode=<%=rscat("invoicecode")%>&invoicedate=<%=rscat("invoicedate")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1"><%=rscat("invoicecode")%></font> 
								</a>
							</td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("refinvoice")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vouchercode_payment")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vouchercode_adjustment")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba_payment")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transportercode")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transportername")%> </td>    
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stono")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("no_faktur")%> </td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fakturdate")%> </td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("amount"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pph"))),0)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("amountpph"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("Inv_totalamount"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pvd_payment_amount"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pvd_adjustment_amount"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("var_amount"))),2)%></td> 
						</tr>
						  <%
						  v_amountpph=v_amountpph+trim(cdbl(to_nol(rsCat("amountpph"))))
						  v_amount =v_amount+Trim(cdbl(to_nol(rsCat("amount"))))
						  v_totalamount =v_totalamount+Trim(cdbl(to_nol(rsCat("Inv_totalamount"))))
						  v_pvd_payment_amount =v_pvd_payment_amount+Trim(cdbl(to_nol(rsCat("pvd_payment_amount"))))
						  v_pvd_adjustment_amount =v_pvd_adjustment_amount+Trim(cdbl(to_nol(rsCat("pvd_adjustment_amount"))))
						  v_var_amount =v_var_amount+Trim(cdbl(to_nol(rsCat("var_amount")))) 
						  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
						  rsCat.MoveNext
						  'response.Flush()
						  no=no+1
						Wend
						rsCat.close()
						rsCat=null%>
							<input  type="hidden" value="<%=txtTable%>" name="txtTable">
							<input  type="hidden" value="<%=txtModule%>" name="txtModule">
							</form>
							<tr bgcolor="#F4F4A0">
								<td align="Center" style="<%=cmdstyleLeft%>" colspan="12"> <font face="Verdana" size="1" style = "font-weight: bold;">T o t a l</td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style = "font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_amount)),0)%></td>
								<td></td>
								<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style = "font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_amountpph)),0)%></td>
							 	<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style = "font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_totalamount)),0)%></td>
							 	<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style = "font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_pvd_payment_amount)),0)%></td>
							 	<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style = "font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_pvd_adjustment_amount)),0)%></td>
							 	<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1" style = "font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_var_amount)),0)%></td>
							</tr>
                           </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	