<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Sales - Delivery Order</title>
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
#product-table th {font-size:10px; }
#product-table td {font-size:10px; padding:2px;padding-bottom: 0;
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
Dim oRS, oSQL, NoUrut, nTotalValue, grnno,transportername

grnno = Request("DONo")

oSQL	=	" SELECT" & _
			" a.fcba,customercode, " & _
			" customer_api.get_description(customercode,a.fcba) customername, " & _
			" a.salescontractno, " & _
			" a.deliveryorderno, " & _
			" a.productcode, " & _
			" a.quantity, " & _
			" a.transportercontractno, " & _
			" a.transporter, " & _
			" supplier_api.Get_Transporter_Name(a.transporter,a.fcba) transportername, " & _
			" a.remarks," & _
			" a.deliveryorderdate " & _
			" FROM deliveryorder a " & _ 
			" INNER JOIN salesagreementhdr b ON (a.salescontractno=b.salescontractno AND a.fcba=b.fcba) " & _
			" WHERE a.FCBA = '" & Request("fcba") & "'" & _
			"       AND a.DeliveryOrderNo = '" & grnno & "'" & _ 
			"       AND a.transporter='" &  Request("transcode") & "'" 

Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open oSQL, connstr 
transportername = oRS("transportername")

%>
<div id="wrapper">
	<div class="navi_to_print">
	<input type="button" value="Print" onclick="printDiv('note')"/>
</div>
<div class="letter">
<div class="note" >
<style>
		@media all {
		td{padding:5px;}
		table.PAYSLIP td{padding:5px;}
		.page-break hr {display:block;border: 1px dashed red;}
		.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
		#note {margin-left:auto;padding:5px;}
		#info_note .judul { text-align:center; font-size:14px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:0px;}
		#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:0px;}
		#info_note .point .name { width:100px; float:left; margin-top: 0px;}
		#info_note .point .isi  { float:left; height:5px;  margin-left:10px;}

		#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:0px; text-align:left;}
		#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:0px;}
		#info_note .point_right .name { width:100px; float:left; margin-top: 0px;}
		#info_note .point_right .isi  { float:left; height:5px;  margin-left:10px;}
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
		#info_note .judul { text-align:center; font-size:14px;}
		#info_note .left { float: left; width:35%;}
		#info_note .right {  float: right; width:35%;}
		#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:0px;}
		#info_note .point span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:0px;}
		#info_note .point .name { width:100px; float:left; margin-top: 0px;}
		#info_note .point .isi  { float:left; height:5px;  margin-left:10px;}

		#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:0px; text-align:left;}
		#info_note .point_right span { font-size:11px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:0px;}
		#info_note .point_right .name { width:100px; float:left; margin-top: 0px;}
		#info_note .point_right .isi  { float:left; height:5px;  margin-left:10px;}
		.clear { clear:both; content:'';}
		table .long {width:400px;}
		table .small {width:40px;}
		table th span { font-style:italic; font-weight:normal;}
		.line { height:1px; background:#000; content:' ';margin: auto;}
		.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
		#product-table th { background-color:#fff; color:#000; font-size:10;}
		#product-table td { border-color:#000; color:#000;padding-bottom: 0;order-top:0px;border-bottom:0px;
    padding-top: 0;}
		#product-table td.none { border: none;}
		#footer_note .block { font-size:10px; float:left; width:19%; height:150px; background-color:#fff; margin-left:1%; text-align:center;}
		.tembusan{font-size:12px;}
		input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
		hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
		hr{ margin-top: 5px;}
		}
		</style>
		
<%If Not oRS.EOF Then%>
	<header id="info_note" >
		<div class="judul"><strong>DELIVERY ORDER</strong></div>
			<table width="100%">
			<!-- Left -->
			<td valign="top">
				<div class="point">
					<div class="name">
					No.Kontrak  
					</div>
					<div class="isi">
					: <%=oRS("salescontractno")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Customer  
					</div>
					<div class="isi">
					: <%=oRS("customername")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					No.Transporter
					</div>
					<div class="isi">
					:  <%=oRS("transportercontractno")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point">
					<div class="name">
					Transporter
					</div>
					<div class="isi">
					: <%=oRS("transportername")%>
					</div>
				</div>
			</td> 
		 
			<!-- right -->
			<td width="400px" valign="top">			
				<div class="point_right">
					<div class="name">
					DO
					</div>
					<div class="isi">
					: <%=oRS("deliveryorderno")%>
					</div>
				</div>
				<div class="clear"></div>
				<div class="point_right">
					<div class="name">
					Tanggal
					</div>
					<div class="isi">
					: <%=oRS("deliveryorderdate")%> 
					</div>
				</div>
				<div class="clear"></div>
			</td>
		</table>
		

	</header>

<%End If%>

<%
oSQL	=	""
oRS.Close
set oRS	=	Nothing
%>

<div class="clear"></div>
<br/><br/>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
    <tr>
	    <th>No</th>
        <th>Produk</th>
        <th>Periode </th>
        <th>DO.Qty</th>
		<th>Satuan</th>
        <th>Keterangan</th>
        
    </tr>
</thead >

<tbody>
<%

NoUrut	=	0
oSQL	=	" SELECT" & _
			"     productcode " & _
			"     , startdate" & _
			"     , enddate" & _
			"     , quantity_do" & _
			"     , deliverycondition" & _
			"     , unitofmeasure" & _
			" FROM deliveryorder a" & _
			" WHERE a.FCBA = '" & Request("fcba") & "'" & _
			"       AND a.DeliveryOrderNo = '" & grnno & "'" & _ 
			"       AND a.transporter='" &  Request("transcode") & "'" 

Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open oSQL, connstr, adLockOptimistic, adOpenStatic
%>
<%If oRS.BOF And oRS.EOF Then%>
    <tr>
        <td Colspan="14">NO DATA AVAILABLE</td>
    </tr>
<%Else
	While Not oRS.EOF
		NoUrut	=	NoUrut + 1%>
		<tr>
			<td align="Right"><%=NoUrut%></td>
			<td align="Left"><%=oRS("productcode")%></td>
			<td align="Left"><%=oRS("startdate")%> to <%=oRS("enddate")%> </td>
			<td align="Right"><%=FormatNumber(oRS("quantity_do"), 2)%></td>
			<td align="Left"><%=oRS("unitofmeasure")%></td>
			<td align="Left"><%=oRS("deliverycondition")%></td>			
		</tr>
	<%	nTotalValue	=	nTotalValue + Cdbl(oRS("quantity_do"))
		oRS.MoveNext
	Wend%>
 
<%End If%>
<%
oSQL	=	""
oRS.Close
set oRS	=	Nothing
%>
</tbody>
</table>
<br/>
        <div id="footer_note">
    	
		<div class="block">
		Dibuat Oleh,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Commercial
        </div>
		
		<div class="block">
		Disetujui Oleh,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Commercial Dept
        </div>
        
		<div class="block">Disetujui Oleh,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        Dir.Ops
        </div>
		
		<div class="block">Disetujui Transporter,
        <br/><br/><br/><br/>
		<input class="isian" type="text" value="">
        <hr/>
        <%=transportername%>
        </div>
    
    <div class="clear"></div>
</div>




</body>
</html>
