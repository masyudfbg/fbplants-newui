<%'Option Explicit 
'On Error GOTO 0 
%>
<%response.buffer=True%>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print Sales Invoice</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
        <script language="JavaScript">
        xAddEventListener(window, 'load',
          function() {
            new xTableHeaderFixed('xthf-site', window);
          }, false
        );
        </script>
		<script language="JavaScript" src="/js/jquery.printPage.js"></script> 
	<script>
					function toggle(id){
						var e = document.getElementById(id);
						if(e.style.display == 'inline-block'){
							e.style.display = 'none';
						}else if(e.style.display == 'none'){
							e.style.display = 'inline-block';
						}
					}
	</script>
	<script>
	 function printDiv(divName) {
	 //var printContents = document.getElementById(divName).innerHTML;
     //var originalContents = document.body.innerHTML;

     //document.body.innerHTML = printContents;

     window.print();

     //document.body.innerHTML = originalContents;
}
</script>
<style type="text/css">
@media all {
#product-table td{padding:5px;padding-bottom: 0; border-top:0px;border-bottom:0px;
    padding-top: 0;}
.page-break	{ display: block; }
a.button_print { display:block; }	
#wrapper{background: #e2e2e2; /* Old browsers */
background: -moz-linear-gradient(-45deg,  #e2e2e2 0%, #dbdbdb 50%, #d1d1d1 51%, #fefefe 100%); /* FF3.6+ */
background: -webkit-gradient(linear, left top, right bottom, color-stop(0%,#e2e2e2), color-stop(50%,#dbdbdb), color-stop(51%,#d1d1d1), color-stop(100%,#fefefe)); /* Chrome,Safari4+ */
background: -webkit-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Chrome10+,Safari5.1+ */
background: -o-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Opera 11.10+ */
background: -ms-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* IE10+ */
background: linear-gradient(135deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* W3C */
filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#e2e2e2', endColorstr='#fefefe',GradientType=1 ); /* IE6-9 fallback on horizontal gradient */
}
a.button_print {display: inline-block;}
.letter { 
background-color:#fff; 
box-shadow:0px 0px 4px rgba(0,0,0,0.4); 
border-radius:4px;
margin: auto;
width: 98%;
padding-top:30px;
padding-bottom:30px;
padding-left:1%;
padding-right:1%;
}
#header-wrapper {display:block}
.navi_to_print {
display:block;
margin: auto;
width: 98%;
padding-top:30px;
}
}
@media print
{
table td{padding:3px;}
#header-wrapper {display:none}
.button_print {display: none;}
.printhide{display:none;}
.page-break	{ display: block; page-break-before: always; }
.product-table{width: 100%;}
#wrapper{background:#fff;padding:none;}
#product-table th {font-size:14px; }
#product-table td {font-size:14px; padding:2px;padding-bottom: 0;
    padding-top: 0;order-top:0px;border-bottom:0px;}
a.button_print {display:none;}
.letter { 
background-color:#fff; 
box-shadow:none; 
border-radius:0px;
width: 100%;
margin: none;
padding-top:0px;
padding-bottom:0px;
padding-left:0px;
padding-right:0px;
}
.navi_to_print {
display:none;
margin: auto;
width: 100%;
padding-top:30px;
}
}
</style>
</head>

<body>
<%
' Const adOpenForwardOnly = 0
' Const adLockReadOnly = 1
' Const adCmdTableDirect = &H0200
' Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleStock="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; width:100px"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
 <%
 Dim txtNoPR,txtDatePR,txtRemarks,q,nCountofField,no,i
 Dim Remarks,txtmanagedid,txtStatus,txtHeaderRemarks
 dim stock
 businessunit = Request("businessunit")
 invoicecode = request("invoicecode")
 
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
			
   Dim rsPO,intRecordsPerPage,cColor,txtSupplier
   set rsPO = server.CreateObject("Adodb.recordset")
  q = 	" SELECT sid.fcba, " &_
		    "   sid.invoicedpno, " &_
		    "	TO_CHAR (sid.transdate, 'dd') transdate_dd, " &_
		    "	TO_CHAR (sid.transdate, 'mm') transdate_mm, " &_
		    "	TO_CHAR (sid.transdate, 'yyyy') transdate_yyyy, " &_
		    "	sid.transdate, " &_
		    "	sah.customercode, " &_
		    "	TO_CHAR (sid.valid_until, 'dd') valid_until_dd, " &_
		    "	TO_CHAR (sid.valid_until, 'mm') valid_until_mm, " &_
		    "	TO_CHAR (sid.valid_until, 'yyyy') valid_until_yyyy, " &_
		    "	sp.description customername, " &_
		    "	sah.salescontractno, " &_
		    "	sad.unitofmeasure, " &_
		    "	case " &_
		    "	    when sid.lineno = 1 then sid.percentage_dp" &_
		    "	else" &_
		    "	    sid.percentage_dp - (select  sum(percentage_dp)" &_
		    "	    from " &_
		    "	    salesinvoicedp where salescontractno = sah.salescontractno " &_
		    "	    and fcba =sid.fcba and lineno < sid.lineno)" &_
		    "	end percentage_dp," &_
		    "	sp.address addresscust, " &_
		    "	MAX (sid.dp_vat) vat," &_
		    "	MAX (sad.amount) grandtotal, " &_
			" 	SUM (sad.quantity) quantity, " &_
		    "	sid.salescontractno, " &_
		    "	MAX (sid.dp_beforevat) amount_beforevat, " &_
		    "	MAX (sid.amount_dpvat) amountvat, " &_
		    "	sad.productcode, " &_
		    "	PRODUCT_API.GET_DESCRIPTION (sad.productcode, sid.fcba) productname, " &_
		    "	SUM (sid.dp_beforevat) amount, " &_
		    "	(NVL(SUM (sid.dp_beforevat),0) * NVL(SUM(sid.dp_vat),0)  / 100) ppn, " &_
		    "	bu2.address, " &_
		    "	bu2.fcname, " &_
		    "	bu2.fcfax, " &_
		    "	bu2.fctelephone " &_
		    "	FROM salesinvoicedp sid  " &_
		    "	INNER JOIN salesagreementhdr sah " &_
		    "	ON (    sid.salescontractno = sah.salescontractno AND sid.fcba = sah.fcba) " &_
		    "	INNER JOIN salesagreementdet sad ON (    sah.salescontractno = sad.salescontractno AND sid.fcba = sad.fcba) " &_
		    "	LEFT JOIN customer sp ON (sah.customercode = sp.customercode AND sid.fcba = sp.fcba) " &_
		    "	LEFT JOIN (SELECT * FROM businessunit WHERE fccode = '"& businessunit &"') bu2 ON (bu2.fccode = sid.fcba) " &_
		    "	WHERE sid.fcba = '"& businessunit &"' AND sid.invoicedpno = '"& invoicecode &"' " &_
		    "	GROUP BY sid.fcba, " &_
		    "	sid.invoicedpno, " &_
		    "	TO_CHAR (sid.transdate, 'dd'), " &_
		    "	TO_CHAR (sid.transdate, 'mm'), " &_
		    "	TO_CHAR (sid.transdate, 'yyyy'), " &_
		    "	sah.customercode, " &_
		    "	sid.transdate, " &_
		    "	TO_CHAR (sid.valid_until, 'dd'), " &_
		    "	TO_CHAR (sid.valid_until, 'mm'), " &_
		    "	TO_CHAR (sid.valid_until, 'yyyy'), " &_
		    "	sp.description, " &_
		    "	sah.salescontractno, " &_
		    "	sad.unitofmeasure, " &_
			" sid.lineno, " &_
			" sid.percentage_dp, " &_
		    "	sp.address, " &_
		    "	sid.salescontractno, " &_
		    "	sad.productcode, " &_
		    "	PRODUCT_API.GET_DESCRIPTION (sad.productcode, sid.fcba), " &_
		    "	bu2.address, " &_
		    "	bu2.fcname, " &_
		    "	bu2.fcfax, " &_
		    "	bu2.fctelephone "

   'response.Write q	
   Set rsPO = Server.CreateObject("ADODB.Recordset")
   rsPO.Open q, connstr
   if not rsPO.eof then
		fcba = rsPO("fcba")
		invoicecode = rsPO("invoicedpno")
		transdate = rsPO("transdate")
		customercode = rsPO("customercode")
		customername = rsPO("customername")
		address = rsPO("address")	
		percentage_dp = rsPO("percentage_dp")
		phone = rsPO("fctelephone")	
		addresscust = rsPO("addresscust")
		fax = rsPO("fcfax")
		amount_beforevat= rsPO("amount_beforevat")
		amountvat= rsPO("amountvat")
		amount= rsPO("amount")
		ppn= rsPO("ppn")
		total= cdbl(to_nol(rsPO("amount"))) + cdbl(to_nol(rsPO("ppn")))
		grandtotal = cdbl(to_nol(total)) - cdbl(to_nol(dp))
   End if
   rsPO.close
   set rsPO = Nothing %>
   
   <%
   Dim rsPurchaseOrder,nQty,nAmount,qprd
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   qprd  = 	" SELECT sid.fcba, " &_
		    "   sid.invoicedpno, " &_
		    "	TO_CHAR (sid.transdate, 'dd') transdate_dd, " &_
		    "	TO_CHAR (sid.transdate, 'mm') transdate_mm, " &_
		    "	TO_CHAR (sid.transdate, 'yyyy') transdate_yyyy, " &_
		    "	sid.transdate, " &_
		    "	sah.customercode, " &_
		    "	TO_CHAR (sid.valid_until, 'dd') valid_until_dd, " &_
		    "	TO_CHAR (sid.valid_until, 'mm') valid_until_mm, " &_
		    "	TO_CHAR (sid.valid_until, 'yyyy') valid_until_yyyy, " &_
		    "	sp.description customername, " &_
		    "	sah.salescontractno, " &_
		    "	sad.unitofmeasure, " &_
		    "	case " &_
		    "	    when sid.lineno = 1 then sid.percentage_dp" &_
		    "	else" &_
		    "	    sid.percentage_dp - (select  sum(percentage_dp)" &_
		    "	    from " &_
		    "	    salesinvoicedp where salescontractno = sah.salescontractno " &_
		    "	    and fcba =sid.fcba and lineno < sid.lineno)" &_
		    "	end percentage_dp," &_
		    "	sp.address addresscust, " &_
		    "	MAX (sid.dp_vat) vat," &_
		    "	MAX (sad.amount) grandtotal, " &_
		    "	sid.salescontractno, " &_
		    "	MAX (sid.dp_beforevat) amount_beforevat, " &_
		    "	MAX (sid.amount_dpvat) amountvat, " &_
		    "	sad.productcode, " &_
		    "	PRODUCT_API.GET_DESCRIPTION (sad.productcode, sid.fcba) productname, " &_
			" SUM (sad.quantity) quantity, " &_
			" MAX (sad.salesprices) unitprice, "&_
		    "	SUM (sid.dp_beforevat) amount, " &_
		    "	(NVL(SUM (sid.dp_beforevat),0) * NVL(SUM(sid.dp_vat),0)  / 100) ppn, " &_
		    "	bu2.address, " &_
		    "	bu2.fcname, " &_
		    "	bu2.fcfax, " &_
		    "	bu2.fctelephone " &_
		    "	FROM salesinvoicedp sid  " &_
		    "	INNER JOIN salesagreementhdr sah " &_
		    "	ON (    sid.salescontractno = sah.salescontractno AND sid.fcba = sah.fcba) " &_
		    "	INNER JOIN salesagreementdet sad ON (    sah.salescontractno = sad.salescontractno AND sid.fcba = sad.fcba) " &_
		    "	LEFT JOIN customer sp ON (sah.customercode = sp.customercode AND sid.fcba = sp.fcba) " &_
		    "	LEFT JOIN (SELECT * FROM businessunit WHERE fccode = '"& businessunit &"') bu2 ON (bu2.fccode = sid.fcba) " &_
		    "	WHERE sid.fcba = '"& businessunit &"' AND sid.invoicedpno = '"& invoicecode &"' " &_
		    "	GROUP BY sid.fcba, " &_
		    "	sid.invoicedpno, " &_
		    "	TO_CHAR (sid.transdate, 'dd'), " &_
		    "	TO_CHAR (sid.transdate, 'mm'), " &_
		    "	TO_CHAR (sid.transdate, 'yyyy'), " &_
		    "	sah.customercode, " &_
		    "	sid.transdate, " &_
		    "	TO_CHAR (sid.valid_until, 'dd'), " &_
		    "	TO_CHAR (sid.valid_until, 'mm'), " &_
		    "	TO_CHAR (sid.valid_until, 'yyyy'), " &_
		    "	sp.description, " &_
		    "	sah.salescontractno, " &_
		    "	sad.unitofmeasure, " &_
			" sid.lineno, " &_
			" sid.percentage_dp, " &_
		    "	sp.address, " &_
		    "	sid.salescontractno, " &_
		    "	sad.productcode, " &_
		    "	PRODUCT_API.GET_DESCRIPTION (sad.productcode, sid.fcba), " &_
		    "	bu2.address, " &_
		    "	bu2.fcname, " &_
		    "	bu2.fcfax, " &_
		    "	bu2.fctelephone "
  ' response.Write qprd	
  
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open qprd, connstr%>
<div id="wrapper">
	<div class="navi_to_print">
	<input type="button" value="Print" onclick="printDiv('note')"/>
</div>

<div class="letter">
	<div id="note">
		<style>
		@media all {
		td{padding:5px;}
		table.PAYSLIP td{padding:5px;}
		.page-break hr {display:block;border: 1px dashed red;}
		.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
		#note {margin-left:auto;padding:5px;}
		#info_note .judul { text-align:center; font-size:20px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:12px; font-weight:bold; width:100%;  margin-top:5px;}
		#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point .name { width:150px; float:left; margin-top: 10px;}
		#info_note .point .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}

		#info_note .point_right { font-size:12px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
		#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point_right .name { width:100px; float:left; margin-top: 10px;}
		#info_note .point_right .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}
		.clear { clear:both; content:'';}
		table .long {width:400px;}
		table .small {width:40px;}
		table th span { font-style:italic; font-weight:normal;}
		.line { height:1px; background:#000; content:' ';margin: auto;}
		.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
		#product-table th { background-color:#fff; color:#000;border:1px #000000 solid ;}
		#product-table td { border-color:#000; color:#000;padding-bottom: 0;order-top:0px;border-bottom:0px;
    padding-top: 0;}
		#product-table td.none { border: none;}
		#footer_note .block { font-size:10px; float:left; width:19%; height:150px; background-color:#fff; margin-left:1%; text-align:center;}
		.tembusan{font-size:10px;}
		input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
		hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
		hr{ margin-top: 5px;}
		}
		@media print
		{
		td{padding:3px;}
		table.PAYSLIP td{padding:3px;}
		.page-break hr {border:none;display: none;}
		.PAYSLIP td {font-family:verdana,sans-serif; font-size:7px; padding:0px;}
		#note {margin-left:auto;padding:0px;}
		#info_note .judul { text-align:center; font-size:20px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:14px; font-weight:bold; width:100%;  margin-top:5px;}
		#info_note .point span { font-size:12px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point .name { width:150px; float:left; margin-top: 10px;}
		#info_note .point .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}

		#info_note .point_right { font-size:14px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
		#info_note .point_right span { font-size:12px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point_right .name { width:100px; float:left; margin-top: 10px;}
		#info_note .point_right .isi  { float:left; height:36px; line-height:35px; margin-left:20px;}
		.clear { clear:both; content:'';}
		table .long {width:400px;}
		table .small {width:40px;}
		table th span { font-style:italic; font-weight:normal;}
		.line { height:1px; background:#000; content:' ';margin: auto;}
		.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
		#product-table th { background-color:#fff; color:#000; font-size:14;}
		#product-table td { border-color:#000; color:#000;padding-bottom: 0;order-top:0px;border-bottom:0px;
    padding-top: 0;}
		#product-table td.none { border: none;}
		#footer_note .block { font-size:14px; float:left; width:19%; height:150px; background-color:#fff; margin-left:1%; text-align:center;}
		.tembusan{font-size:12px;}
		input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
		hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
		hr{ margin-top: 5px;}
		}
		</style>
		
		<header id="info_note" > 
			<table style="border-style:none;">
				<tbody> 
					<tr>
						<td width="10%" rowspan = "4" style="border-right-style:none;"><img src="css/login/images/client.png" height="80px"></td>
					</tr>
					<tr>
						<td width="56.6666666667%"><p style="font-size:12pt;"><b>PT.<%=CompanyName%></b></p></td>
					</tr>
					<tr>
						<td width="56.6666666667%"><p style="font-size:10pt;"><b><%=address%></b></p></td>
					</tr>
					<tr>
						<td><p style="font-size:10pt;"><b>Phone :<%=phone%> &nbsp;&nbsp; Fax :<%=fax%></b></p></td> 
					</tr>
				</tbody>
			</table>
			<hr>
			<table width="100%">
				<!-- Left -->
				<%IF NOT rsPurchaseOrder.EOF THEN%>
				<td valign="top">
				<div class="point">
						<div class="name">
							Kepada
						</div>
						<div class="isi">
						: <%=rsPurchaseOrder("customername")%>
						</div>
					</div>
					<div class="clear"></div>
					<div class="point">
						<div class="name">
						Alamat
						</div>
						<div class="isi">
						: <%=rsPurchaseOrder("addresscust")%>
						</div>
					</div>
					<div class="clear"></div>
					<div class="point">
						<div class="name">
						NPWP
						</div>
						<div class="isi">
						: 
						</div>
					</div>
					<div class="clear"></div>
				<!-- right -->
				<td width="400px" valign="top">			
					<div class="point_right">
						<div class="name">
						No. Invoice 
						</div>
						<div class="isi">
						: <%=rsPurchaseOrder("invoicedpno")%>
						</div>
					</div>
					<div class="clear"></div>
					<div class="point_right">
						<div class="name">
						Tanggal Invoice
						</div>
						<div class="isi">
						: <%=rsPurchaseOrder("transdate_dd")%>&nbsp;<%=Bulan(rsPurchaseOrder("transdate_mm"))%>&nbsp;<%=rsPurchaseOrder("transdate_yyyy")%>
						</div>
					</div>
					<div class="clear"></div>
					<div class="point_right">
						<div class="name">
						Tanggal Jatuh Tempo 
						</div>
						<div class="isi">
						: <%=rsPurchaseOrder("valid_until_dd")%>&nbsp;<%=Bulan(rsPurchaseOrder("valid_until_mm"))%>&nbsp;<%=rsPurchaseOrder("valid_until_yyyy")%>
						</div>
					</div>
					<div class="clear"></div>
					<div class="point_right">
						<div class="name">
						Sales Kontrak 
						</div>
						<div class="isi">
						: <%=rsPurchaseOrder("salescontractno")%>
						</div>
					</div>
				</td>
				<%end if%>
			</table>
		</header>
		<div class="clear"></div>
		

		<style>
			.group_checkbox {margin-bottom:10px;margin-top:25px;}
			.group_checkbox ul {width:100%;}
			.group_checkbox ul li { 
				list-style-type:none;
				float:left;
			}
			.group_checkbox ul li p{
				font-weight:bold;
				font-size:12px;
				float:left;
				display:initial;
			}
			.group_checkbox ul li span{
				float:left;
				margin-top:0px;
				margin-left:10px;
			}
			.group_checkbox ul li span img{
				width:21px;
			}
		</style>
		 
		<table border="0" width="100%" id="product-table" style="border-bottom:1px #000000 solid;">
			<thead>
				<tr>
					<th>Nomor </th>
					<th>Kode</th>
					<th>Nama Barang</th>
					<th>Quantity</th>
					<th>Unit</th>
					<th>(%)DP</th>
					<th>Harga Satuan (Rp)</th>
					<th>Jumlah (Rp)</th> 
				</tr>
			 </thead>
			 <tbody>
			  <%
			  
			   if Not rsPurchaseOrder.eof then
					
				  nCountofField=rsPurchaseOrder.fields.count-1
				  no=0
				  nQty=0
				  nAmount=0 
				  ngrandtotal = 0
				  ' amount = 0 
				  ' ppn = 0
				  %>
				  <tr>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					  <td></td>
					</tr>
				  
				  <%
				  While Not rsPurchaseOrder.EOF
					' if len(trim(rsPurchaseOrder("stock"))) > 0 then
						' stock = rsPurchaseOrder("stock")
					' else
						' stock = "0"
					' end if
					  nQty = nQty+cdbl(rsPurchaseOrder("Quantity"))
					  nAmount = nAmount+cdbl(rsPurchaseOrder("grandtotal"))
					  ngrandtotal = ngrandtotal+cdbl(rsPurchaseOrder("grandtotal"))
					  
					  if no  mod 2 = 1 then
						 cColor = "#EEEEEE"
					  else
						 cColor = "#E5E5E5"
					  end if%>
					  
				 <tr>
					  <td align="center" style="width:5%"><%=no+1%>&nbsp;</td>
					  <td align="center" style="width:5%" title="productcode" valign="middle"><%=rsPurchaseOrder("productcode")%>&nbsp;</td>
					  <td align="center" style="width:25%" title="productname" valign="middle"><%=rsPurchaseOrder("productname")%>&nbsp;</td>
					  <td align="center"  style="width:50px;<%=cmdstyle%>" title="quantity" valign="middle"><%=formatnumber(cdbl(to_nol(rsPurchaseOrder("quantity"))),2)%>&nbsp;</td>
					  <td align="center"  style="width:50px;<%=cmdstyle%>" title="unit" valign="middle"><font face="Verdana"><%=rsPurchaseOrder("unitofmeasure")%></td>
					  <td align="center"  style="width:50px;<%=cmdstyle%>" title="unit" valign="middle"><font face="Verdana"><%=rsPurchaseOrder("percentage_dp")%></td>
					  <td align="right"  style="width:50px;<%=cmdstyle%> " title="harga satuan" valign="middle"><%=formatnumber(cdbl(to_nol(rsPurchaseOrder("unitprice"))),2)%></td>
					  <td align="right"  style="width:50px;<%=cmdstyle%> " title="approvedqty" valign="middle"><%=formatnumber(cdbl(to_nol(rsPurchaseOrder("grandtotal"))),2)%>&nbsp;</td> 
				 </tr>
				
				  <%
					  rsPurchaseOrder.MoveNext
					  intRecordsPerPage = intRecordsPerPage - 1
					  no=no+1
				  Wend  %>
				   <tr>
					<th align="right" colspan="3">Sub Total</th>
					<th align="center" ><%=formatnumber(cdbl(to_nol(nQty)),2)%></th>
					<th align="right" ></th>
					<th align="right" ></th>
					<th align="right" ></th>
					<th align="right" ><%=formatnumber(cdbl(to_nol(nAmount)),2)%></th>
				  </tr>
					<tr>
						<th colspan="6" rowspan="6" valign="top">
							<%
								Function Val( myString )
								' Val Function for VBScript (aka ParseInt Function in VBScript).
								' By Denis St-Pierre.
								' Natively VBScript has no function to extract numbers from a string.
								' Based shamelessly on MS' Helpfile example on RegExp object.
								' CAVEAT: Returns only the *last* match found
								'         (or, with objRE.Global = False, only the *first* match)
								 
									Dim colMatches, objMatch, objRE, strPattern
								 
									' Default if no numbers are found
									Val = 0
								 
									strPattern = "[-+0-9]+"       ' Numbers positive and negative; use
																  ' "ˆ[-+0-9]+" to emulate Rexx' Value()
																  ' function, which returns 0 unless the
																  ' string starts with a number or sign.
									Set objRE = New RegExp        ' Create regular expression object.
									objRE.Pattern    = strPattern ' Set pattern.
									objRE.IgnoreCase = True       ' Set case insensitivity.
									objRE.Global     = True       ' Set global applicability:
																  '   True  => return last match only,
																  '   False => return first match only.
									Set colMatches = objRE.Execute( myString )  ' Execute search.
									For Each objMatch In colMatches             ' Iterate Matches collection.
										Val = objMatch.Value
									Next
									Set objRE= Nothing
								End Function
								 
								Function Terbilang(Angka)
								   Dim n1
									Dim nangka, Strjmlhuruf, intpecahan  , strpecahan, urai, Bil1, strtot, bil2
									Dim x, Y, z, i  
								   
									For i = 1 To Len(Angka)
										If Not Mid(Angka, i, 1) = "." Then
											nangka = nangka + Mid(Angka, i, 1)
										End If
									Next
								   
									If Len(LTrim(nangka)) = 0 Then
										terbit = "Nol Rupiah"
										Exit Function
									End If
									Strjmlhuruf = LTrim(RTrim((nangka)))
									intpecahan = Val(Right(Mid(Angka, 15, 2), 2))
									If (intpecahan = 0) Then
										strpecahan = ""
									End If
									x = 0
									Y = 0
									urai = ""
								   
									While (x < Len(Strjmlhuruf))
										x = x + 1
										strtot = Mid(Strjmlhuruf, x, 1)
										Y = Y + Val(strtot)
										z = Len(Strjmlhuruf) - x + 1
								   
										Select Case Val(strtot)
										Case 1
											If (z = 1 Or z = 7 Or z = 10 Or z = 13) Then
												Bil1 = "Satu "
											ElseIf (z = 4) Then
												If (x = 1) Then
													Bil1 = "Se"
												Else
													Bil1 = "Satu "
												End If
											ElseIf (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
												x = x + 1
												strtot = Mid(Strjmlhuruf, x, 1)
												z = Len(Strjmlhuruf) - x + 1
												bil2 = ""
								   
											Select Case Val(strtot)
											Case 0
												Bil1 = "Sepuluh "
											Case 1
												Bil1 = "Sebelas "
											Case 2
												Bil1 = "Dua Belas "
											Case 3
												Bil1 = "Tiga Belas "
											Case 4
												Bil1 = "Empat Belas "
											Case 5
												Bil1 = "Lima Belas "
											Case 6
												Bil1 = "Enam Belas "
											Case 7
												Bil1 = "Tujuh Belas "
											Case 8
												Bil1 = "Delapan Belas "
											Case 9
												Bil1 = "Sembilan Belas "
											End Select
											Else
												Bil1 = "Se"
											End If
								   
										Case 2
											Bil1 = "Dua "
										Case 3
											Bil1 = "Tiga "
										Case 4
											Bil1 = "Empat "
										Case 5
											Bil1 = "Lima "
										Case 6
											Bil1 = "Enam "
										Case 7
											Bil1 = "Tujuh "
										Case 8
											Bil1 = "Delapan "
										Case 9
											Bil1 = "Sembilan "
										Case Else
											Bil1 = ""
										End Select
								   
									If (Val(strtot) > 0) Then
										If (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
											bil2 = "Puluh "
										ElseIf (z = 3 Or z = 6 Or z = 9 Or z = 12 Or z = 15) Then
											bil2 = "Ratus "
										Else
											bil2 = ""
										End If
									Else
										bil2 = ""
									End If
									If (Y > 0) Then
										Select Case z
										Case 4
											bil2 = bil2 + "Ribu "
											Y = 0
										Case 7
											bil2 = bil2 + "Juta "
											Y = 0
										Case 10
											bil2 = bil2 + "Milyar "
											Y = 0
										Case 13
											bil2 = bil2 + "Trilyun "
											Y = 0
										End Select
									End If
									urai = urai + Bil1 + bil2
								Wend
								   
								urai = urai + strpecahan
								Terbilang = urai & ""
								Terbilang = UCase(Terbilang)
								n1 = UCase(Left(Terbilang, 1))
								Terbilang = n1 & Mid(Terbilang, 2, Len(Terbilang) - 1)
								End Function


							%>
							<div class="title_terbilang" align="left">
							Terbilang :
							</div>
							<div class="terbilang" align="left">
								<%  txtSQL= "select upper(replace(terbilang_indo("& grandtotal &"),'-',' ')) terbilang from dual"
									' response.write txtSQL
									Set rsCats = Server.CreateObject("ADODB.Recordset")
									rsCats.Open txtSQL, connstr
									if Not rsCats.eof then
										v_terbilang =  rsCats("terbilang") 
									end if
									rsCats.close
									set rsCats = Nothing
								%>
								<%=v_terbilang%> RUPIAH
							</div>
						</th> 
					</tr>
					<tr>	
						<th align="left">DP <%=percentage_dp%>% :</th>
						<th align="right"><%=formatnumber(cdbl(to_nol(amount)),2)%></th>
					</tr>					
					<tr>	
						<th align="left">DPP : </th>
						<th align="right"><%=formatnumber(cdbl(to_nol(amount)),2)%></th>
					</tr> 
					<tr>	
						<th align="left">PPn :</th>
						<th align="right"><%=formatnumber(cdbl(to_nol(ppn)),2)%></th>
					</tr>
					<tr>
						 
					</tr>   
					<tr>	
						<th align="left">Grand Total : </th>
						<th align="right"><%=formatnumber(cdbl(to_nol(amount)) + (cdbl(to_nol(ppn))) ,2)%></th>
					</tr>
					<tr>
						<th valign="top" align="left" colspan="6">Bank Account :
						<p>PT.<%=CompanyName%></p>
						<p>BCA Cab. CBD Pluit</p>
						<p>A/C. 8060 287 555</p>
						</th>
						<th colspan="2" align="right" valign="top">
							<input class="isian" type="text" value="" style="font-size:12px;">
							</br></br></br></br></br></br></br>
							<input class="isian" type="text" value="" style="font-size:12px;">
							<input class="isian" type="text" value="Direktur" style="font-size:12px;">
						</th>
					</tr>
					<tr>
						<th colspan="8">Note : Pembayaran baru dianggap sah apabila cek/giro telah dicairkan, atau transfer sudah diterima di rekening Bank kami</th>
					</tr>
			<%	  
			   else
				  response.Write("<tr><td colspan='9'>Data is empty</td></tr>")	  
			   end if 
			   rsPurchaseOrder.Close
			   set rsPurchaseOrder = Nothing%> 
			    
			   </tbody>
		   </table>
		   <br/>
</body>
</html>


