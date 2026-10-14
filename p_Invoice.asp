<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>Reporting - Invoice</title>
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
v_txtSUPPLIERCODE= trim(request("SUPPLIERCODE"))
v_txtINVOICECODE= trim(request("INVOICECODE"))
v_txtITEMCODE= trim(request("ITEMCODE"))

businessunit= trim(request("businessunit"))
company     = trim(request("company"))


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
	else
	   fcbaname = "ALL UNIT"
	   CompanyName= company
	end if

if request("txtKind") = "SupplierBalance" Then

	txtSQL = 	" SELECT fcba_payment," & _
				"        fcba_invoice," & _
				"        invoicecode," & _
				"        refinvoice," & _
				"        invoicedate," & _
				"        invoicedatesupplier," & _
				"        suppliercode," & _
				"		 SUPPLIER_API.GET_SUPP_NAME(suppliercode, fcba_invoice) suppliername," & _
				"        pocode," & _
				"        amount," & _
				"        totalppn," & _
				"        amountvat," & _
				"        othercharge," & _
				"        inv_totalamount," & _
				"        vouchercode_payment," & _
				"        pvd_payment_amount," & _
				"        vouchercode_adjustment," & _
				"        pvd_adjustment_amount," & _
				"        (inv_totalamount - pvd_payment_amount - pvd_adjustment_amount)" & _
				"           AS var_amount," & _
				"        postingflg" & _
				"   FROM (  SELECT MAX (fcba_payment) AS fcba_payment," & _
				"                  MAX (fcba_invoice) AS fcba_invoice," & _
				"                  invoicecode," & _
				"                  refinvoice," & _
				"                  MAX (invoicedate) AS invoicedate," & _
				"                  MAX (invoicedatesupplier) AS invoicedatesupplier," & _
				"                  MAX (suppliercode) AS suppliercode," & _
				"                  MAX (pocode) AS pocode," & _
				"                  SUM (amount) AS amount," & _
				"                  SUM (totalppn) AS totalppn," & _
				"                  SUM (amountvat) AS amountvat," & _
				"                  SUM (othercharge) AS othercharge," & _
				"                  SUM (totalamount) AS inv_totalamount," & _
				"                  MAX (vouchercode_payment) AS vouchercode_payment," & _
				"                  SUM (pvd_payment_amount) AS pvd_payment_amount," & _
				"                  MAX (vouchercode_adjustment) AS vouchercode_adjustment," & _
				"                  SUM (pvd_adjustment_amount) AS pvd_adjustment_amount," & _
				"                  MAX (postingflg) postingflg " & _
				"             FROM (SELECT inv.invoice_by AS fcba_payment," & _
				"                          inv.fcba AS fcba_invoice," & _
				"                          inv.invoicecode," & _
				"                          inv.refinvoice," & _
				"                          TO_CHAR (inv.invoicedate, 'mm/dd/yyyy') AS invoicedate," & _
				"                          TO_CHAR (inv.invoicedatesupplier, 'mm/dd/yyyy') AS invoicedatesupplier," & _
				"                          inv.suppliercode," & _
				"                          inv.pocode," & _
				"                          inv.amount," & _
				"                          inv.totalppn," & _
				"                          inv.amountvat," & _
				"                          inv.othercharge," & _
				"                          inv.totalamount," & _
				"                          '' AS vouchercode_payment," & _
				"                          0 AS pvd_payment_amount," & _
				"                          '' AS vouchercode_adjustment," & _
				"                          0 AS pvd_adjustment_amount,inv.postingflg " & _
				"                     FROM invoice inv" & _
				"                    WHERE inv.fcba in (SELECT fccode from BusinessUnit where fccompanycode='" & company & "' and fccode like '%"& businessunit &"%')" & _
				"                   UNION ALL" & _
				"                     SELECT pvd.fcba AS fcba_payment," & _
				"                            '' AS fcba_invoice," & _
				"                            pvd.invoicecode," & _
				"							 '' AS refinvoice " & _		
				"                            '' AS invoicedate," & _
				"                            '' AS invoicedatesupplier," & _
				"                            '' AS suppliercode," & _
				"                            '' AS pocode," & _
				"                            0 AS amount," & _
				"                            0 AS totalppn," & _
				"                            0 AS amountvat," & _
				"                            0 AS othercharge," & _
				"                            0 AS totalamount," & _
				"                            pvd.vouchercode AS vouchercode_payment," & _
				"                            NVL (SUM (pvd.amount), 0) AS pvd_payment_amount," & _
				"                            '' AS vouchercode_adjustment," & _
				"                            0 AS pvd_adjustment_amount,0 postingflg" & _
				"                       FROM paymentvoucherdetail pvd" & _
				"                            INNER JOIN paymentvoucher pv" & _
				"                               ON (    pvd.fcba = pv.fcba" & _
				"                                   AND pvd.vouchercode = pv.vouchercode)" & _
				"                      WHERE     pvd.locationtype IN ('AP')" & _
				"                            AND pv.transtype = '02-PAYMENT'" & _
				"                   GROUP BY pvd.fcba," & _
				"                            pvd.invoicecode," & _
				"                            pvd.vouchercode," & _
				"                            pvd.locationtype" & _
				"                   UNION ALL" & _
				"                     SELECT pvd.fcba AS fcba_payment," & _
				"                            '' AS fcba_invoice," & _
				"                            pvd.invoicecode," & _
				"                            '' AS refinvoice," & _
				"                            '' AS invoicedate," & _
				"                            '' AS invoicedatesupplier," & _
				"                            '' AS suppliercode," & _
				"                            '' AS pocode," & _
				"                            0 AS amount," & _
				"                            0 AS totalppn," & _
				"                            0 AS amountvat," & _
				"                            0 AS othercharge," & _
				"                            0 AS totalamount," & _
				"                            '' AS vouchercode_payment," & _
				"                            0 AS pvd_payment_amount," & _
				"                            pvd.vouchercode AS vouchercode_adjustment," & _
				"                            nvl (sum (pvd.amount), 0) as pvd_adjustment_amount,0 postingflg " & _
				"                       FROM paymentvoucherdetail pvd" & _
				"                            INNER JOIN paymentvoucher pv" & _
				"                               ON (    pvd.fcba = pv.fcba" & _
				"                                   AND pvd.vouchercode = pv.vouchercode)" & _
				"                      WHERE     pvd.locationtype IN ('AP')" & _
				"                            AND pv.transtype = '03-JOURNAL'" & _
				"                   GROUP BY pvd.fcba," & _
				"                            pvd.invoicecode," & _
				"                            pvd.vouchercode," & _
				"                            pvd.locationtype)" & _
				"         GROUP BY invoicecode,refinvoice) "  
else
	txtSQL =  	" SELECT fcba_payment," & _
				"        fcba_invoice," & _
				"        invoicecode," & _
				"        refinvoice," & _
				"        invoicedate," & _
				"        invoicedatesupplier," & _
				"        suppliercode," & _
				"		 SUPPLIER_API.GET_SUPP_NAME(suppliercode, fcba_invoice) suppliername," & _
				"        pocode," & _
				"        amount," & _
				"        totalppn," & _
				"        amountvat," & _
				"        othercharge," & _
				"        inv_totalamount," & _
				"        vouchercode_payment," & _
				"        pvd_payment_amount," & _
				"        vouchercode_adjustment," & _
				"        pvd_adjustment_amount," & _
				"        (inv_totalamount - pvd_payment_amount - pvd_adjustment_amount)" & _
				"           AS var_amount," & _
				"        postingflg" & _
				"   FROM (  SELECT MAX (fcba_payment) AS fcba_payment," & _
				"                  MAX (fcba_invoice) AS fcba_invoice," & _
				"                  invoicecode," & _
				"                  refinvoice," & _
				"                  MAX (invoicedate) AS invoicedate," & _
				"                  MAX (invoicedatesupplier) AS invoicedatesupplier," & _
				"                  MAX (suppliercode) AS suppliercode," & _
				"                  MAX (pocode) AS pocode," & _
				"                  SUM (amount) AS amount," & _
				"                  SUM (totalppn) AS totalppn," & _
				"                  SUM (amountvat) AS amountvat," & _
				"                  SUM (othercharge) AS othercharge," & _
				"                  SUM (totalamount) AS inv_totalamount," & _
				"                  MAX (vouchercode_payment) AS vouchercode_payment," & _
				"                  SUM (pvd_payment_amount) AS pvd_payment_amount," & _
				"                  MAX (vouchercode_adjustment) AS vouchercode_adjustment," & _
				"                  SUM (pvd_adjustment_amount) AS pvd_adjustment_amount," & _
				"                  MAX (postingflg) AS postingflg" & _
				"             FROM (SELECT inv.invoice_by AS fcba_payment," & _
				"                          inv.fcba AS fcba_invoice," & _
				"                          inv.invoicecode," & _
				"                          inv.refinvoice," & _
				"                          inv.invoicedate," & _
				"                          inv.invoicedatesupplier," & _
				"                          inv.suppliercode," & _
				"                          inv.pocode," & _
				"                          inv.amount," & _
				"                          inv.totalppn," & _
				"                          inv.amountvat," & _
				"                          inv.othercharge," & _
				"                          inv.totalamount," & _
				"                          '' AS vouchercode_payment," & _
				"                          0 AS pvd_payment_amount," & _
				"                          '' AS vouchercode_adjustment," & _
				"                          0 AS pvd_adjustment_amount,inv.postingflg" & _
				"                     FROM invoice inv" & _
				"                    WHERE inv.fcba in (SELECT fccode from BusinessUnit where fccompanycode='" & company & "' and fccode like '%"& businessunit &"%')" & _
				"                   UNION ALL" & _
				"                     SELECT pvd.fcba AS fcba_payment," & _
				"                            '' AS fcba_invoice," & _
				"                            pvd.invoicecode," & _
				"                            '' AS refinvoice, " & _
				"                            null AS invoicedate," & _
				"                            null AS invoicedatesupplier," & _
				"                            '' AS suppliercode," & _
				"                            '' AS pocode," & _
				"                            0 AS amount," & _
				"                            0 AS totalppn," & _
				"                            0 AS amountvat," & _
				"                            0 AS othercharge," & _
				"                            0 AS totalamount," & _
				"                            pvd.vouchercode AS vouchercode_payment," & _
				"                            NVL (SUM (pvd.amount), 0) AS pvd_payment_amount," & _
				"                            '' AS vouchercode_adjustment," & _
				"                            0 AS pvd_adjustment_amount,0 as postingflg" & _
				"                       FROM paymentvoucherdetail pvd" & _
				"                            INNER JOIN paymentvoucher pv" & _
				"                               ON (    pvd.fcba = pv.fcba" & _
				"                                   AND pvd.vouchercode = pv.vouchercode)" & _
				"                      WHERE     pvd.locationtype IN ('AP')" & _
				"                            AND pv.transtype = '02-PAYMENT'" & _
				"                   GROUP BY pvd.fcba," & _
				"                            pvd.invoicecode," & _
				"                            pvd.vouchercode," & _
				"                            pvd.locationtype" & _
				"                   UNION ALL" & _
				"                     SELECT pvd.fcba AS fcba_payment," & _
				"                            '' AS fcba_invoice," & _
				"                            pvd.invoicecode," & _
				"                            null as refinvoice," & _
				"                            null AS invoicedate," & _
				"                            null AS invoicedatesupplier," & _
				"                            '' AS suppliercode," & _
				"                            '' AS pocode," & _
				"                            0 AS amount," & _
				"                            0 AS totalppn," & _
				"                            0 AS amountvat," & _
				"                            0 AS othercharge," & _
				"                            0 AS totalamount," & _
				"                            '' AS vouchercode_payment," & _
				"                            0 AS pvd_payment_amount," & _
				"                            pvd.vouchercode AS vouchercode_adjustment," & _
				"                            nvl (sum (pvd.amount), 0) as pvd_adjustment_amount,0 as postingflg " & _
				"                       FROM paymentvoucherdetail pvd" & _
				"                            INNER JOIN paymentvoucher pv" & _
				"                               ON (    pvd.fcba = pv.fcba" & _
				"                                   AND pvd.vouchercode = pv.vouchercode)" & _
				"                      WHERE     pvd.locationtype IN ('AP')" & _
				"                            AND pv.transtype = '03-JOURNAL'" & _
				"                   GROUP BY pvd.fcba," & _
				"                            pvd.invoicecode," & _
				"                            pvd.vouchercode," & _
				"                            pvd.locationtype)" & _
				"         GROUP BY invoicecode,refinvoice) "  
end if


if  len(trim(v_txtINVOICEDATE_From)) > 0 then
    txtSQL = txtSQL & " WHERE invoicedate >= to_date('" & v_txtINVOICEDATE_From & "','mm/dd/yyyy')"
	txtSelect = " From " & FormatDateTime(v_txtINVOICEDATE_From,1)
end if
if  len(trim(v_txtINVOICEDATE_To)) > 0 then
    txtSQL =txtSQL &  " AND invoicedate <= to_date('" & v_txtINVOICEDATE_To & "','mm/dd/yyyy')"
	txtSelect = txtSelect & " to " & FormatDateTime(v_txtINVOICEDATE_To,1)
end if
if  len(trim(v_txtSUPPLIERCODE)) > 0 then
    txtSQL = txtSQL & " and SupplierCode='" & v_txtSUPPLIERCODE & "'"
	txtSelect = txtSelect & " Supplier Code " & v_txtSUPPLIERCODE
end if
if  len(trim(v_txtINVOICECODE)) > 0 then
    txtSQL =txtSQL &  " and INVOICECODE='" & v_txtINVOICECODE & "'"
	txtSelect = txtSelect & " Inv. Code " & v_txtINVOICECODE
end if 
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and ITEMCODE='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Item. Code " & v_txtITEMCODE
end if 

' txtsql=(txtsql) & 	" GROUP BY Inv.INVOICECODE," &  _ 
					' "          Inv.INVOICEDATE," &  _ 
					' "          Inv.SupplierCode," &  _ 
					' "          Inv.POCODE," &  _ 
					' "          InvDet.ITEMCODE," &  _ 
					' "          pvd.vouchercode," &  _ 
					' "          st.FcName"  
					
txtsql=(txtsql) & 	" ORDER BY INVOICECODE," & _
					"          INVOICEDATE "  
					
					
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(txtsql)
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
					<h2>INVOICE</h2>
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
						    <th colspan="2"></th>
							<th>Date Entry</th>
							<th>Date Invoice Supplier</th>
							<th>Invoice Code</th>
							<th>Ref Invoice</th>
							<th>Voucher Code Payment</th>
							<th>Voucher Code Adjustment</th>
							<th>PO Code</th>
							<th>Supplier Code</th>	  
							<th>Amount + Other Charge Stock</th>
							<th>Amount PPn</th>
							<th>Amount After PPn</th>
							<th>Other Charge Non Stock</th>
							<th>Total Amount Invoice</th>
							<th>Total Amount Voucher</th>
							<th>Total Amount Voucher Adj.</th>
							<th>VAR Amount</th>
							<th>Posting</th>
						</tr>
                         </thead>
                         <tbody>
							<%
							Dim v_Quantity,v_price,v_amount
							v_Quantity=0
							v_amount=0
							v_totalppn = 0
							v_amountvat = 0
							v_othercharge = 0
							v_Inv_totalamount = 0
							v_pvd_totalamount = 0
							v_var_amount = 0 
							v_pvd_adjustment_amount = 0

							While Not rsCat.EOF%>
						<tr>
							<td align="Left" style="<%=cmdstyle%>">
							<a href="invoice.Asp?fccode=<%=rscat("INVOICECODE")%>&fcba=<%=businessunit%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(2)%>" border="0">
							</a>
							</td>								
							<td>
							<a href="javascript:launchwin('printinvoice.Asp?fcCode=<%=rscat("INVOICECODE")%>&fcba=<%=businessunit%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
							<img src="IMAGES/print.gif" alt="Print Data : <%=rscat(2)%>" border="0">
							</a>	
							</td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICEDATE")%></td> 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoicedatesupplier")%></td> 
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("INVOICECODE")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("refinvoice")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vouchercode_payment")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("vouchercode_adjustment")%></td>
							<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("POCODE")%> </td>
							<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SupplierCode") & " : " & rscat("suppliername")%> </td>  
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("amount"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("totalppn"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("amountvat"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("othercharge"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("Inv_totalamount"))),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pvd_payment_amount"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("pvd_adjustment_amount"))),2)%></td> 
							<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("var_amount"))),2)%></td> 
						    <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%if rscat("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%></td>
						</tr>
						  <%
						  ' v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
						v_amount= v_amount+Trim(cdbl(to_nol(rsCat("amount"))))
						v_totalppn = v_totalppn+Trim(cdbl(to_nol(rsCat("totalppn"))))
						v_amountvat =v_amountvat +Trim(cdbl(to_nol(rsCat("amountvat"))))
						v_othercharge =v_othercharge +Trim(cdbl(to_nol(rsCat("othercharge"))))
						v_Inv_totalamount = v_Inv_totalamount+Trim(cdbl(to_nol(rsCat("Inv_totalamount"))))
						v_pvd_totalamount = v_pvd_totalamount+Trim(cdbl(to_nol(rsCat("pvd_payment_amount"))))
						v_pvd_adjustment_amount = v_pvd_adjustment_amount+Trim(cdbl(to_nol(rsCat("pvd_adjustment_amount"))))
						v_var_amount = v_var_amount+Trim(cdbl(to_nol(rsCat("var_amount")))) 
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
							<tr bgcolor="#F4F4A0" style="font-weight: bold;">
							<td align="Center" style="<%=cmdstyleLeft%>" colspan="10"> <font face="Verdana" size="1">T o t a l</td> 
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amount)),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_totalppn)),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_amountvat)),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_othercharge)),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_Inv_totalamount)),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_pvd_totalamount)),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_pvd_adjustment_amount)),2)%></td>
							<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_var_amount)),2)%></td>
							</tr>
                           </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>	