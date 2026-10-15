<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>ePlaS Reporting - Sales Invoice</title>
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

v_txtTRANSDATE_From=trim(request("TRANSDATE_From"))
v_txtTRANSDATE_To=trim(request("TRANSDATE_To"))
v_txtSALESCONTRACTNO=trim(request("SALESCONTRACTNO"))
v_txtCUSTOMERCODE=trim(request("CUSTOMERCODE"))
v_txtSALESSHIPMENTNO=trim(request("SALESSHIPMENTNO"))
v_txtINVOICECODE=trim(request("INVOICECODE"))
businessunit=trim(request("businessunit"))
productcode=trim(request("productcode"))
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
	end if

txtSQL =  	" SELECT sih.fcba," & _
			"        sih.invoice_by," & _
			"        sih.invoicecode," & _
			"        sih.transdate," & _
			"        sih.customercode," & _
			"        customer_api.get_description (sih.customercode, sih.fcba) customername," & _
			"		 sah.productcode," & _
			"		 sum(sid.quantity) quantity," & _
			"        sih.vat," & _
			"        AVG(sih.grandtotal)grandtotal," & _
			"        sih.remarks," & _
			"        sih.salescontractno," & _
			"        AVG(sih.amount_beforevat) amount_beforevat," & _
			"        AVG(sih.amountvat)amountvat," & _
			"        sih.VALID_FROM," & _
			"        sih.valid_until," & _
			"        sih.no_faktur," & _
			"        sih.expire_date," & _
			"        sih.fakturdate," & _
			"        sih.invoicedpno" & _ 
			"   FROM salesinvoicedtl sid" & _
			"		 INNER JOIN salesinvoicehdr sih ON (sih.invoicecode = sid.invoicecode AND sih.fcba = sid.fcba)" & _
			"        INNER JOIN salesagreementdet sah" & _
          	"	ON (    sah.fcba = sih.fcba AND sah.salescontractno = sih.salescontractno) " & _
			" where sih.fcba in (SELECT fccode from BusinessUnit where fccompanycode='" & company & "' and fccode like '%"& businessunit &"%')" 

' response.write(txtSQL)
if  len(trim(v_txtTRANSDATE_From)) > 0 then
    txtSQL = txtSQL & " and TRANSDATE >=to_Date('" & v_txtTRANSDATE_From & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " From " & FormatDateTime(v_txtTRANSDATE_From,1)	
end if
if  len(trim(v_txtTRANSDATE_To)) > 0 then
    txtSQL =txtSQL &  " and TRANSDATE <=to_Date('" & v_txtTRANSDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & " To " & FormatDateTime(v_txtTRANSDATE_To,1)	
end if
if  len(trim(v_txtINVOICECODE)) > 0 then
    txtSQL = txtSQL & " and sih.INVOICECODE='" & v_txtINVOICECODE & "'"
	txtSelect = txtSelect & " Sales Contract No. " & v_txtINVOICECODE	
end if
if  len(trim(v_txtSALESCONTRACTNO)) > 0 then
    txtSQL = txtSQL & " and sih.SALESCONTRACTNO='" & v_txtSALESCONTRACTNO & "'"
	txtSelect = txtSelect & " Sales Contract No. " & v_txtSALESCONTRACTNO	
end if
if  len(trim(v_txtCUSTOMERCODE)) > 0 then
    txtSQL =txtSQL &  " and sih.customercode='" & v_txtCUSTOMERCODE & "'"
	txtSelect = txtSelect & " Delivery Order No. " & v_txtCUSTOMERCODE	
end if 
' if  len(trim(v_txtSALESSHIPMENTNO)) > 0 then
'     txtSQL =txtSQL &  " and sidt.salesshipmentno='" & v_txtSALESSHIPMENTNO & "'"
' 	txtSelect = txtSelect & " Sales Shipment No. " & v_txtSALESSHIPMENTNO
' end if
if  len(trim(productcode)) > 0 then
    txtSQL =txtSQL &  " and sah.productcode='" & productcode & "'"
	txtSelect = txtSelect & " Product Code " & productcode
end if 

' if  len(trim(v_txtRowStateId)) > 0 then
'     txtSQL =txtSQL &  " and a.ROWSTATEID='" & v_txtRowStateId & "'"
' 	txtSelect = txtSelect & " Status  " & v_txtRowStateId	
' end if
	txtSQL = txtSQL &	" GROUP BY sih.fcba," & _
						"        sih.invoice_by," & _
						"        sih.invoicecode," & _
						"        sih.transdate," & _
						"        sih.customercode," & _
						"        customer_api.get_description (sih.customercode, sih.fcba) ," & _
						"		 sah.productcode," & _ 
						"        sih.vat," & _ 
						"        sih.remarks," & _
						"        sih.salescontractno," & _ 
						"        sih.VALID_FROM," & _
						"        sih.valid_until," & _
						"        sih.no_faktur," & _
						"        sih.expire_date," & _
						"        sih.fakturdate," & _
						"        sih.invoicedpno"  
	
	txtSQL = txtSQL & " Order by transdate,invoicecode " 'ditambah ukie tgl 9 okt 2013 untuk Order by


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
         <th colspan="18" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>SALES INVOICE</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="18" style=" background-color:#FFFFFF; color:#000000; border:none"><%=txtSelect%></td>
						</tr>
						<!--</div>-->
							<tr>
									<%	
									
									Set rsCat = Server.CreateObject("ADODB.Recordset")
									rsCat.CursorType = 3
									rsCat.CursorLocation = 3
									rsCat.Open txtSQL, connstr

									
									if Not rsCat.eof then
									   nCountofField=rsCat.fields.count-1
									end if
									NO=1%>
							</tr>
							<tr>
								<th>Invoice Code</th>
								<th>Date</th> 
								<th>Invoice By</th>
								<th>Sales Contract</th> 
								<th>Expired Date</th>
								<th>Product Code</th>
								<th>No. Faktur</th>
								<th>Faktur Date</th>								
								<th>Customer Code</th>
								<th>Customer Name</th>
								<th>Quantity</th>
								<th>Amount Before VAT</th>
								<th>VAT</th>
								<th>Amount VAT</th>
								<th>Grand Total</th>
								<th>Remarks</th>
							</tr>
                            </thead>
                            <tbody>
								<%
								Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
								v_Quantity=0
								v_Amount = 0
								v_Amountvat = 0
								v_Amount_beforevat = 0
								v_grandtotal = 0
								v_amountvat = 0
								While Not rsCat.EOF%>


									<tr>
									</td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">
											<a href="javascript:launchwin('print_salesinvoice.Asp?invoicecode=<%=rscat("invoicecode")%>&transdate=<%=rscat("transdate")%>&salescontractno=<%=rscat("salescontractno")%>&businessunit=<%=rscat("fcba")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
											<font face="Verdana" size="1"><%=rscat("invoicecode")%></font>
											</a>
										</td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoice_by")%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("salescontractno")%></td>  
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("expire_date")%></td>  
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("productcode")%></td>  
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("no_faktur")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fakturdate")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("customercode")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("customername")%></td>   
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("quantity")),2)%></td>   
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("amount_beforevat")),2)%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("vat"))),2)%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("amountvat"))),2)%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("grandtotal")),2)%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td> 
									</tr>
								  <%

								  v_Quantity = v_Quantity + cdbl(rscat("quantity"))
								  ' v_amount = v_amount + cdbl(rscat("amount"))
								  v_amount_beforevat = v_amount_beforevat + cdbl(rscat("amount_beforevat"))
								  v_grandtotal = v_grandtotal + cdbl(to_nol(rscat("grandtotal")))
								  v_amountvat =  v_amountvat + cdbl(to_nol(rscat("amountvat")))
								  rsCat.MoveNext
								  'response.Flush()
								  no=no+1
								Wend
								rsCat.close()
								rsCat=null%>
								<input  type="hidden" value="<%=txtTable%>" name="txtTable">
								<input  type="hidden" value="<%=txtModule%>" name="txtModule">
								</form>
								</tbody>
								<tr  style="background-color: #F4F4A0; font-weight: bold;">
									<td colspan="10" align="center"> TOTAL </td>   
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(v_quantity),2)%></td>
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(v_amount_beforevat),2)%></td>
									<td></td>
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(v_amountvat),2)%></td>
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(v_grandtotal),2)%></td>	
									<td></td>
								</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</body>	
</HTML>