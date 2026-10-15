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

<title>ePlaS Reporting - Special Invoice </title>
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
ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID")
 use_ppn = Request("use_ppn")
 fccompanycode = trim(request("company"))
 InvoiceType =  trim(request("InvoiceType"))
 businessunit = Request("businessunit")
 purchaseby = Request("purchaseby")
 suppliername = Request("suppliername")
 supplieraddress = Request("supplieraddress")
 termofpayment = Request("termofpayment")
 SHIPPINGPOINT = Request("SHIPPINGPOINT")
 dp = Request("dp")
 ppn = Request("ppn")
 invoiceno = Request("invoiceno")
 StockName = Request("StockName")
 unitprice = Request("unitprice")
 disc = Request("disc")
 othercost = Request("othercost")
 totalamount = Request("totalamount")
 JobCode = Request("JobCode")
 targettype = Request("targettype")
 targetcode = Request("targetcode")
 grandtotal = Request("grandtotal")
 amountcharge = Request("amountcharge")
 totalppn = Request("totalppn")
 fakturno = request ("fakturno")
 description = request ("description")
 transdate_from = Request("transdate_from")
 transdate_to = Request("transdate_to")
 ponumber = Request("ponumber")

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

txtSQL = 	"	SELECT ish.company,"  & _
			"       ish.fcba,"  & _
			"       ish.invoice_type,"  & _
			"       ish.invoice_by,"  & _
			"       ish.invoiceno,"  & _
			"       ish.transdate,"  & _
			"       ish.pocode,"  & _
			"       ish.fakturdate,"  & _
			"       ish.no_faktur,"  & _
			"       ish.suppliername as suppliercode,"  & _
			"       supplier_api.get_supp_name(ish.suppliername, ish.fcba) suppliername,"  & _
			"       ish.ppn,"  & _
			"       ish.grandtotal,"  & _
			"       isd.jobcode,"  & _
			"       isd.targettype,"  & _
			"       isd.targetcode,"  & _
			"       isd.description,"  & _
			"		isd.kg, " & _
			"		isd.rate, " & _
			"		isd.amount, " & _
       		"		isd.othercost, " & _ 
			"		ish.postingflg, " & _ 
			"       supplier_api.get_controljob(ish.suppliername, ish.fcba) controljob"  & _
			"  FROM invoiceservicehdr ish"  & _
			"       INNER JOIN invoiceservicedtl isd"  & _
			"          ON (ish.invoiceno = isd.invoiceno AND ish.fcba = isd.fcba)"

if  len(trim(fccompanycode)) > 0 then
    txtSQL =txtSQL &  "  where ish.company = '"& fccompanycode &"'"
end if
if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  "  and ish.fcba = '"& businessunit &"'"
end if
if  len(trim(invoiceno)) > 0 then
    txtSQL =txtSQL &  " and ish.invoiceno = '"& invoiceno &"'"
end if
if  len(trim(ponumber)) > 0 then
    txtSQL =txtSQL &  " and ish.pocode = '"& ponumber &"'"
end if
if  len(trim(transdate_from)) > 0 then
    txtSQL = txtSQL & " and transdate >=to_Date('" & transdate_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(transdate_from,1)
end if
if  len(trim(transdate_to)) > 0 then
    txtSQL =txtSQL &  " and transdate <=to_Date('" & transdate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(transdate_to,1)
end if
if  len(trim(InvoiceType)) > 0 then
    txtSQL =txtSQL &  " and ish.invoice_type = '"& InvoiceType &"'"
end if

txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

' response.Write(txtsql)
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
         <th colspan="24" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=businessunit%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>SPECIAL INVOICE</h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="24" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
				    <th colspan="2"></th> 
					<th>Company</th>
					<th>Business Unit</th>
					<th>Invoice By</th>
					<th>Date</th>
					<th>Invoice No.</th>
					<th>Invoice Type</th>
					<th>PO No.</th>
					<th>Supplier Code</th>
					<th>Supplier Name</th>
					<th>ControlJob</th>
					<th>Faktur Date</th>
					<th>Faktur No.</th>
					<th>Jobcode</th>
					<th>Target Type</th>
					<th>Target Code</th>
					<th>Description</th>
					<th>KG</th>
					<th>Rate</th>
					<th>Ppn (%)</th>
					<th>Cost</th>
					<th>Other Cost</th>
					<th>Grand Total</th>
					<th>Posting</th>
				</tr>
                </thead>
               <tbody>
	<%
	if Not rscat.eof then
		v_amount = 0
		v_othercost = 0
		v_grandtotal = 0
		v_kg = 0
	While Not rscat.EOF


	%>
		<tr>
		   <td align="Left" style="<%=cmdstyle%>">
		    <a href="invoiceservice.Asp?invoiceno=<%=rscat("invoiceno")%>&fcba=<%=rscat("fcba")%>&txtJenis=Edit">
		    <img src="IMAGES/edit.gif" alt="Edit Data : <%=rscat(4)%>" border="0">
		    </a>
		    </td>								
			<td>
			<a href="javascript:launchwin('PrintInvoiceservice.Asp?fcCode=<%=rscat("invoiceno")%>&fcba=<%=rscat("fcba")%>&Mode=Print','newwindow','width='+(980)+',height='+(650)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
			<img src="IMAGES/print.gif" alt="Print Data : <%=rscat(4)%>" border="0">
			</a>	
			</td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("company")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoice_by")%></td>			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoiceno")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("invoice_type")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pocode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("suppliercode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("suppliername")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("controljob")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fakturdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("no_faktur")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targettype")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targetcode")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("description")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("kg"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("rate"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ppn")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("amount"),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("othercost"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("grandtotal"))),2)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%if rscat("postingflg")="1" then response.write "Ya" else response.write "Tidak" end if%></td>
		</tr>
	  <%
	   v_amount =v_amount+trim(cdbl(rsCat("amount")))
	   v_kg =v_kg+trim(cdbl(rsCat("kg")))
	   v_othercost =v_othercost+Trim(cdbl(to_nol(rsCat("othercost"))))
	   v_grandtotal =v_grandtotal+Trim(cdbl(rsCat("grandtotal")))
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='20'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="18"> <font face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_kg,2)%></td>
		<td colspan="2"></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amount,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_othercost,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_grandtotal,2)%></td>
	</tr>
	</tbody>
</table>
<table width="90%" class="data">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=MgrName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=KTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
       </tbody>
   </table>
   </body>
   </html>