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
fcba=trim(request("fcba"))
productcode=trim(request("productcode"))
invoicecode=trim(request("invoicecode"))

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

txtSQL =  	" select invoicecode, salesshipmentno, productcode, quantity, unitprice, amount, state, postingdate from salesinvoicedtl" & _ 
			" where fcba = '"& fcba &"' and invoicecode = '"& invoicecode &"'" 

txtSQL = txtSQL & " Order by invoicecode, SALESSHIPMENTNO " 'ditambah ukie tgl 9 okt 2013 untuk Order by


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
         <th colspan="6" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>SALES INVOICE DETAIL</h2>
				</div>
			</div>
				</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<th colspan="6" style=" background-color:#FFFFFF; color:#000000; border:none"><%=txtSelect%></td>
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
								<th>Sales Shipment No</th> 
								<th>Quantity</th>
								<th>Unit Price</th>
								<th>Amount</th>
								<th>State</th>
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
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoicecode")%></td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("salesshipmentno")%> </td> 
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rscat("quantity")),0)%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("unitprice"))),0)%> </td>
										<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rscat("amount"))),0)%> </td>
										<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("state")%> </td> 
									</tr>
								  <%

								  v_Quantity = v_Quantity + cdbl(rscat("quantity"))
								  v_amount = v_amount + cdbl(rscat("amount"))
								  ' v_amount_beforevat = v_amount_beforevat + cdbl(rscat("amount_beforevat"))
								  ' v_grandtotal = v_grandtotal + cdbl(to_nol(rscat("grandtotal")))
								  ' v_amountvat =  v_amountvat + cdbl(to_nol(rscat("amountvat")))
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
									<td colspan="2" align="center"> TOTAL </td> 
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(v_Quantity),0)%></td> 
									<td></td>
									<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(v_amount),0)%></td> 
									<td align="right" style="<%=cmdstyle%>" ><font face="Verdana" size="1"></td>
								</tr>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</body>	
</HTML>