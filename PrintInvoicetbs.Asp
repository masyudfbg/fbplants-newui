<%'Option Explicit
%>
<%'On Error GOTO CmdErrorlog 
%>
<%response.buffer=True%>

<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Print Invoice PO</title>
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

     window.print();

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

cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleStock="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; width:100px"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>

<%
 	Dim txtNoinvoice,txtSupplier,q,nCountofField,no,i
 	Dim txtPeriodeSupp,txtNetto2,txtHarga,txtTotalHarga,txtPPH,txtGrandtotal
	Dim rsInvoicetbs,rsInvoicedtl,sqldtl
	Dim txtTglInvoice,txtBank,txtNamaBank,txtNamaPemilikRek
	Dim txtdtltgl,txtdtlneto,txtdtlgrading,txtdtlnet2,txtdtlkgadj,txtdtlharga,txtdtlpph,txtdtljumlah,txthargarata
	Dim totnet1,totgrading,totnet2,totsub,totpph,totjumlahall,tothargaall,txtgrading,kg_adj,totadj
	



    Set rsInvoicetbs = Server.CreateObject("ADODB.Recordset")
 	q = "select h.invoiceno,to_char(h.transdate,'dd/mm/yyyy') transdate,c.description,to_char(h.valid_from, 'dd/mm/yyyy') || ' - '|| to_char(h.valid_until,'dd/mm/yyyy') as periode,SUM(d.kg) as Netto2, " & _
	" sum(d.amount/d.kg1) as harga, sum(d.amount) as totalharga,sum(d.kg_adj)kg_adj,h.totalpph,h.grandtotal,c.Bank,c.bankaccountno,c.bank_ownership,h.grading   " & _
	" from invoiceservicehdr h,customer c,invoiceservicedtl d " & _
	" where h.invoiceno=d.invoiceno and h.suppliername=c.customercode AND h.fcba=c.fcba and h.invoiceno='" & request("invoiceno") & "' " & _
	" and h.fcba='" & request("fcba") & "' " & _
	" group by h.invoiceno,h.transdate,c.description,h.totalpph,h.grandtotal," & _
	" to_char(h.valid_from, 'dd/mm/yyyy') || ' - '|| to_char(h.valid_until,'dd/mm/yyyy'),c.Bank,c.bankaccountno,c.bank_ownership,h.grading "
  ' response.Write q

   rsInvoicetbs.Open q, connstr

   if not rsInvoicetbs.eof then
      txtNoinvoice 		= rsInvoicetbs("invoiceno")
	  txtTglInvoice 	= rsInvoicetbs("transdate")
	  txtSupplier 		= rsInvoicetbs("description")
	  txtPeriodeSupp 	= rsInvoicetbs("periode")
	  txtNetto2 		= rsInvoicetbs("Netto2")
	  txtHarga 			= cdbl(rsInvoicetbs("totalharga"))/cdbl(rsInvoicetbs("Netto2"))
	  txtTotalHarga 	= rsInvoicetbs("totalharga")
	  txtPPH 			= rsInvoicetbs("totalpph")
	  txtGrandtotal 	= rsInvoicetbs("grandtotal")
	  txtBank 			= rsInvoicetbs("Bank")
	  txtNamaBank 		= rsInvoicetbs("bankaccountno")
	  txtNamaPemilikRek = rsInvoicetbs("bank_ownership")
	  txtgrading 		= rsInvoicetbs("grading")
	  kg_adj 			= rsInvoicetbs("kg_adj")
	  if cdbl(kg_adj)>0 then 
	  	txtTotalHarga = cdbl(kg_adj)*cdbl(txtHarga)
	else 
		txtTotalHarga = txtTotalHarga
	 end if
   End if
   rsInvoicetbs.close
   set rsInvoicetbs = Nothing
   
%>

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
		#info_note .point_right .name { width:50px; float:left; margin-top: 10px;}
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
		#info_note .point span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;padding-top:2px;margin-top:2px;}
		#info_note .point .name { width:100px; float:left; margin-top: 10px;}
		#info_note .point .isi  { float:left; height:-10px; line-height:-10px; margin-left:20px;  padding-left: 70px;}

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
				<div class="judul"><strong>INVOICE TBS</strong></div>
					<table width="100%">
					

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
			.ratakanan{
				text-align:right
			}
			.ratatengah{
				text-align:center
			}
			.posisibaris{
				padding: 5px;
			}
			.tebal{
				font-weight: bold;
			}
			.jaraktanda{
				padding-right: 65px;
			}

			.jaraktandabawah{
				padding-right: 60px;
			}
			
		</style>

		<table border="0" width="100%"  >
			
			<tr>
				<td>Kepada</td>
				<td>:</td>
				<td>Finance Department</td>
				<td>PT</td>
				<td>:</td>
				<td><%=session("company")%></td>
			</tr>
			<tr>
				<td>Tembusan</td>
				<td>:</td>
				<td>Direktur Operasional</td>
				<td>Unit</td>
				<td>:</td>
				<td><%=request("fcba")%> - <%=Session("BA")%></td>
			</tr>
			<tr>
				<td>Dari</td>
				<td>:</td>
				<td>Commercial Department</td>
				<td>No Invoice</td>
				<td>:</td>
				<td><%=txtNoinvoice%></td>
			</tr>
			<tr>
				<td>Perihal</td>
				<td>:</td>
				<td>Pembayaran TBS</td>
				<td>Tanggal</td>
				<td>:</td>
				<td><%=txtTglInvoice%></td>
			</tr>
		</table>
		<br>
		 
		<table border="0" width="100%" id="product-table" style="border-bottom:1px #000000 solid;">
			<thead>
				<tr>
					<th>Supplier</th>
					<th>Periode Supply</th>
					<th style="font-size:10px">Netto 2<div class="line"></div><span>kg</span></th>
					<th style="font-size:10px">Netto Adj<div class="line"></div><span>kg</span></th>
					<th style="font-size:10px">Harga Rp<div class="line"></div><span>kg</span></th>
					<th style="font-size:10px">Total Harga</th>					
					<th style="font-size:10px">PPH 22<div class="line"></div><span>0.25 %</span></th>
					<th style="font-size:10px">Grand Total</th>					
				</tr>
			 </thead>
				<tr>
					<td><%=txtSupplier%></td>
					<td><%=txtPeriodeSupp%></td>
					<td class="ratakanan"><%=formatnumber(cdbl(txtNetto2),0)%></td>
					<td class="ratakanan"><%=formatnumber(cdbl(kg_adj),0)%></td>
					<td class="ratakanan"><%=formatnumber(cdbl(txtHarga),0)%></td>
					<td class="ratakanan"><%=formatnumber(cdbl(txtTotalHarga),0)%></td>
					
					<td class="ratakanan"><%=formatnumber(cdbl(txtTotalHarga)*(0.25/100),0)%></td>
					<td class="ratakanan"><%=formatnumber(cdbl(txtGrandtotal),0)%></td>
				</tr>

		</table>
		Pembayaran dapat ditujukan ke rekening: <br>
		Bank		:<%=txtBank%><br>
		No Akun		:<%=txtNamaBank%><br>
		Atas Nama	:<%=txtNamaPemilikRek%><br><br>
		Mohon agar pembayaran dana pembelian TBS dimaksud dapat direalisasikan pada tanggal : <%=txtTglInvoice%><br><br>

		<div class="posisibaris ratatengah"><strong>MONITORING PEMBELIAN TBS</strong></div>
		<div class="posisibaris ratatengah"><strong><%=txtPeriodeSupp%></strong></div>
		<div class="posisibaris ratatengah"><strong><%=txtSupplier%></strong></div>
		
   	<table border="1" width="100%"  >
	   
	<tr>
		<th rowspan="2">No</th>
		<th rowspan="2">Tanggal</th>
		<th colspan="4">HI</th>
		<th rowspan="2">Harga TBS (Rp)</th>
		<th colspan="3">Pembayaran TBS</th>
	</tr>
	<tr>
		<th>Netto 1 (kg)</th>
		<th>Grading <%=txtgrading%>%</th>
		<th>Netto 2 (kg)</th>
		<th>Netto Adj (kg)</th>
		<th>Total Harga (Rp)</th>
		<th>PPh 22 (0.25%)</th>
		<th>Jumlah Tagihan Rp</th>
	</tr>
	<%
		
		sqldtl = "select description,KG1,gradding,kg,kg_adj,rate,amount " & _
		" from invoiceservicedtl" & _
		" where invoiceno='" & txtNoinvoice & "'"
		Set rsInvoicedtl = Server.CreateObject("ADODB.Recordset")
		rsInvoicedtl.Open sqldtl, connstr
		'response.Write sqldtl

		if Not rsInvoicedtl.eof then
			
			no = 1
			totnet1=0
			totgrading=0
			totnet2=0
			totadj=0
			totsub=0
			totpph=0
			totjumlahall=0
			tothargaall=0
			while not rsInvoicedtl.EOF

				txtdtltgl 		= rsInvoicedtl("description")
				txtdtlnet1		= rsInvoicedtl("KG1")
				txtdtlgrading	= rsInvoicedtl("gradding")
				txtdtlnet2		= rsInvoicedtl("kg")
				txtdtlkgadj		= rsInvoicedtl("kg_adj")
				txtdtlharga		= rsInvoicedtl("rate")
				if cdbl(txtdtlkgadj) > 0 then 
					txtdtljumlah= cdbl(txtdtlkgadj)*cdbl(txtdtlharga)
				else 
					txtdtljumlah	= rsInvoicedtl("amount")
				end if
				
				txtdtlpph		= cdbl(txtdtljumlah)*(0.25/100)
				txtdtljumlahall	= cdbl(txtdtljumlah)-cdbl(txtdtlpph)
				txthargarata	= cdbl(txtdtljumlah)/cdbl(txtdtlnet1)
				
	%>

			<tr>
				<td align="center" style="<%=cmdstyleLeft%>"><%=no%>&nbsp;</td>
				<td align="center"><%=txtdtltgl%></td>
				<td align="center" class="ratakanan"><%=formatnumber(cdbl(txtdtlnet1),0)%>&nbsp;</td>
				<td align="center" class="ratakanan"><%=formatnumber(cdbl(txtdtlgrading),0)%>&nbsp;</td>
				<td align="center" class="ratakanan"><%=formatnumber(cdbl(txtdtlnet2),0)%>&nbsp;</td>
				<td align="center" class="ratakanan"><%=formatnumber(cdbl(txtdtlkgadj),0)%>&nbsp;</td>
				<td align="center" class="ratakanan"><%=formatnumber(cdbl(txtdtlharga),0)%>&nbsp;</td>
				<td align="center" class="ratakanan"><%=formatnumber(cdbl(txtdtljumlah),0)%>&nbsp;</td>
				<td align="center" class="ratakanan"><%=formatnumber(cdbl(txtdtlpph),0)%>&nbsp;</td>
				<td align="center" class="ratakanan"><%=formatnumber(cdbl(txtdtljumlahall),0)%>&nbsp;</td>
				
			</tr>
		<%
			rsInvoicedtl.MoveNext
		  	no=no+1
			totnet1 = totnet1+cdbl(txtdtlnet1)
			totgrading = totgrading+cdbl(txtdtlgrading)
			totnet2 = totnet2+cdbl(txtdtlnet2)
			totadj = totadj+cdbl(txtdtlkgadj)
			tothargaall = cdbl(txthargarata)
			totsub = totsub+cdbl(txtdtljumlah)
			totpph = totpph+cdbl(txtdtlpph)
			totjumlahall = totjumlahall+cdbl(txtdtljumlahall)
		Wend
		%>
			
		<%
		else
			response.Write("Data is empty")	  
		end if 
			rsInvoicedtl.Close
			set rsInvoicedtl = Nothing
   		%>

		   <tr>
		   	<td class="ratakanan tebal" colspan="2">Total</td>
			<td class="ratakanan tebal" ><%=formatnumber(cdbl(totnet1),0)%>&nbsp;</td>
			<td class="ratakanan tebal" ><%=formatnumber(cdbl(totgrading),0)%>&nbsp;</td>
			<td class="ratakanan tebal" ><%=formatnumber(cdbl(totnet2),0)%>&nbsp;</td>
			<td class="ratakanan tebal" ><%=formatnumber(cdbl(totadj),0)%>&nbsp;</td>
			<td></td>
			<td class="ratakanan tebal" ><%=formatnumber(cdbl(totsub),0)%>&nbsp;</td>
			<td class="ratakanan tebal" ><%=formatnumber(cdbl(totpph),0)%>&nbsp;</td>
			<td class="ratakanan tebal" ><%=formatnumber(cdbl(totjumlahall),0)%>&nbsp;</td>
		   </tr>
		
	</table><br>

	<span class="jaraktanda">Dibuat Oleh,</span> <span class="jaraktanda">Diajukan Oleh,</span>	<span class="jaraktanda">Diperiksa Oleh</span><br><br><br>


	<span class="jaraktandabawah">Yosella</span>	<span class="jaraktandabawah">  Raymond Hernadhi.</span> <span class="jaraktanda">Accounting</span>
	
	<br/>
</body>
</html>
 