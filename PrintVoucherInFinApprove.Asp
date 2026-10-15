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
 	Dim txtNoinvoice,txtnoref,txtnogiro,txtdate,txtpaidfor,q,i,bankperusahaan,txttranstype,transfercode
	Dim rsPaymentfin
	Dim printheader
 	


    Set rsPaymentfin = Server.CreateObject("ADODB.Recordset")
 	q = "select h.vouchercode,h.bankcode,b.fcname,h.referenceno,h.cekno,h.datecreated,h.remarks,h.transtype,Nvl(i.chargecode,'-') chargecode " & _
	" from paymentvoucher h" & _
	" INNER JOIN bank b ON (h.bankcode=b.fccode and h.fcba=b.fcba) " & _
    " LEFT JOIN " & _
    " ( " & _
    "  SELECT x.vouchercode,x.fcba,x.chargecode,y.fcname FROM paymentvoucherdetail x " & _
    "  INNER JOIN bank y ON (x.chargecode=y.fccode AND x.chargeunit=y.fcba AND x.chargejob=y.controljob) " & _
    " )i ON (h.vouchercode=i.vouchercode AND h.fcba=i.fcba) " & _
	" where h.bankcode=b.fccode and h.fcba=b.fcba" & _
	" and h.fcba='" & session("FCBA") & "' AND h.vouchercode='" & request("vouchercode") & "' "
   'response.Write q

  rsPaymentfin.Open q, connstr

   if not rsPaymentfin.eof then
      txtNoinvoice 		= rsPaymentfin("vouchercode")
	  txtnoref 			= rsPaymentfin("referenceno")
	  txtnogiro 		= rsPaymentfin("cekno")
	  txtdate 			= rsPaymentfin("datecreated")
	  txtpaidfor 		= rsPaymentfin("remarks")
	  bankperusahaan 	= rsPaymentfin("fcname")
	  txttranstype      = rsPaymentfin("transtype")
	  transfercode      = rsPaymentfin("chargecode")
   End if
   rsPaymentfin.close
   set rsPaymentfin = Nothing
   
   if txttranstype = "01-RECEIVE" then printheader = "BANK RECEIVE VOUCHER" end if
   if txttranstype = "02-PAYMENT" then 
			if transfercode = "-" then
				printheader = "BANK PAYMENT VOUCHER" 
			else
				printheader = "BANK TRANSFER VOUCHER"
			end if
   end if
   
   
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
		/* ---- Perapian tampilan voucher (lembar tidak melebar penuh, label-isi berdekatan) ---- */
		#note .pv-top { display: flex; justify-content: space-between; gap: 16px; padding: 0 4px 6px; font-size: 11px; }
		#note .pv-top .pv-bank { font-weight: bold; }
		#note table.pv-info { table-layout: fixed; margin-top: 10px; font-size: 12px !important; line-height: 1.5 !important; }
		#note table.pv-info col.lb { width: 130px; } #note table.pv-info col.lb2 { width: 100px; }
		#note table.pv-info col.cl { width: 14px; } #note table.pv-info col.v2 { width: 190px; }
		#note table.pv-info td { padding: 3px 4px; vertical-align: top; }
		#note table.pv-items { margin-top: 8px; font-size: 12px !important; line-height: 1.45 !important; border-collapse: collapse; }
		#note table.pv-items > thead > tr > th { padding: 6px 5px; border-bottom: 1px solid #000; font-size: 12px !important; }
		#note table.pv-items td { padding: 5px 5px; vertical-align: top; }
		#note .title_terbilang, #note .terbilang { font-size: 12px; }
		#note table.pv-ket { font-size: 12px !important; }
		#note table.pv-ket td:first-child { width: 130px; white-space: nowrap; }
		#note table.pv-ket td { padding: 6px 4px; }
		#product-table.pv-auth { margin-top: 14px; border-collapse: collapse; font-size: 12px; }
		#product-table.pv-auth th, #product-table.pv-auth td { border: 1px solid #000 !important; padding: 5px 8px !important; }
		@media screen { #wrapper .letter, #wrapper .navi_to_print { max-width: 1100px; } }
		</style>
		
		<header id="info_note" >
			<div class="pv-top"><span>PT : <%=session("company")%> &nbsp;-&nbsp; Unit : <%=session("FCBA")%> - <%=Session("BA")%></span><span class="pv-bank"><%=bankperusahaan%></span></div>
				<div class="judul"><strong><%=printheader%></strong></div>
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
			.ratakiri{
				text-align:left
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
			.jarakjobcode{
				width: 100px;
			}
			.jarakjobname{
				width: 180px;
			}
			.jarakremark{
				width: 400px;
			}
			.jarakamountdb{
				width: 10px;
			}
			.jarakamount{
				width: 120px;
			}
			.jarakno{
				width: 20;
			}

			
		</style>

		<table border="0" width="100%" class="fontcourir pv-info">
			<colgroup><col class="lb"><col class="cl"><col><col class="lb2"><col class="cl"><col class="v2"></colgroup>
			
			<tr>
				<td>
				<%if txttranstype = "01-RECEIVE"  then %>
				Diterima Dari
				<%Else%>
				Dibayarkan Kepada
				<%End if%>
				</td>
				<td>:</td>
				<td><%=txtpaidfor%></td>
				<td>No Bukti</td>
				<td>:</td>
				<td><%=txtNoinvoice%></td>
			</tr>
			<tr>
				<td></td>
				<td></td>
				<td></td>
				<td>Tgl Voucher</td>
				<td>:</td>
				<td><%=txtdate%></td>
			</tr>
			<tr>
				<td></td>
				<td></td>
				<td></td>
				<td>No Cek/Giro</td>
				<td>:</td>
				<td><%=txtnogiro%></td>
			</tr>
						
		</table>
				 
		<table border="0" width="100%" class="pv-items" style="border-top:1px #000000 solid; font-size:10px;line-height:14px;">
			<thead>
				<tr style="border-bottom:1px #000000 solid;">
					<th class="jarakjobcode ratakiri">JobCode</th>
					<th class="jarakjobname ratakiri">JobName</th>
					<th class="jarakjobcode ratakiri">Charge Job</th>
					<th class="jarakjobname ratakiri">Charge Name</th>
					<th class="jarakremark ratakiri">Particulars</th>
					<th style="font-size:10px" class="jarakamount ratakanan">Debet</th>					
					<th style="font-size:10px" class="jarakamount ratakanan">Kredit</th>					
				</tr>
			 </thead>
			 	<%
				 Dim sqldtl,rsdtl
				 Dim txtjobcode,txtjobname,txtchargejob,txtchargename,txtbankcredit,txtcontroljob
				 Dim txtremarks,txtamount,no,txttotal,txttotal_cr,txt_totalamount
				 sqldtl = "select d.jobcode,j.fcname,d.chargejob,(select cg.fcname from job cg where cg.fccode=d.chargejob and cg.fcba=d.fcba) AS chargename, " & _
				" d.remarks,d.amount,b.controljob,b.fcname as bankcredit,h.remarks paidfor,h.totalamount  " & _
				" from paymentvoucher h, paymentvoucherdetail d,job j,bank b  " & _
				" where h.vouchercode=d.vouchercode and d.JOBCODE=j.fccode and h.bankcode=b.FCCODE and b.fcba=d.fcba  " & _
				" and d.fcba=j.fcba and d.vouchercode='" & request("vouchercode") & "' and d.fcba='" & session("FCBA") & "'"
				'response.write sqldtl
				 Set rsdtl = Server.CreateObject("ADODB.Recordset")
				 rsdtl.Open sqldtl, connstr

				if not rsdtl.eof then
					txttotal = 0
					txttotal_cr = 0
					txt_totalamount = 0
				while not rsdtl.EOF
						txtjobcode 		= rsdtl("jobcode")
						txtjobname 		= rsdtl("fcname")
						txtchargejob 	= rsdtl("chargejob")
						txtchargename 	= rsdtl("chargename")
						txtremarks 		= rsdtl("remarks")
						txtamount 		= rsdtl("amount")
						txtcontroljob 	= rsdtl("controljob")
						txtbankcredit 	= rsdtl("bankcredit")
						txtpaidfor 		= rsdtl("paidfor")
						txt_totalamount = rsdtl("totalamount")
				%>
				
				<%if txttranstype = "01-RECEIVE"  then %>
				
				<tr>
					<td class="jarakjobcode ratakiri"><%=txtjobcode%></td>
					<td class="jarakjobname ratakiri"><%=txtjobname%></td>
					<td class="jarakjobcode ratakiri"><%=txtchargejob%></td>
					<td class="jarakjobname ratakiri"><%=txtchargename%></td>
					<td class="jarakremark ratakiri"><%=txtremarks%></td>
					
					<td class="ratakanan" class="jarakamountdb ratakanan"><%=formatnumber(cdbl(txtamount),2)%></td>
				</tr>
				
				<%End if%>
				
				<%if txttranstype = "02-PAYMENT"  then %>
				<tr>
					<td class="jarakjobcode ratakiri"><%=txtjobcode%></td>
					<td class="jarakjobname ratakiri"><%=txtjobname%></td>
					<td class="jarakjobcode ratakiri"><%=txtchargejob%></td>
					<td class="jarakjobname ratakiri"><%=txtchargename%></td>
					<td class="jarakremark ratakiri"><%=txtremarks%></td>
					<% if cdbl(txtamount) > 0 then %>
						<td class="ratakanan" class="jarakamount ratakanan"><%=formatnumber(cdbl(txtamount),2)%></td>
						<td class="ratakanan" class="jarakamount ratakanan"></td>
				    <% else %>
						<td class="ratakanan" class="jarakamount ratakanan"></td>
						<td class="ratakanan" class="jarakamount ratakanan"><%=formatnumber(abs(cdbl(txtamount)),2)%></td>
					<%end if%>
					
				</tr>
				
				<%End if%>
				
				<%
					rsdtl.MoveNext
					if cdbl(txtamount) > 0 then 
					txttotal = txttotal+cdbl(txtamount)
					else
					txttotal_cr = txttotal_cr + abs(cdbl(txtamount))
					end if
				Wend
				else
					response.Write("Data is empty")	  
				end if 
					rsdtl.Close
					set rsdtl = Nothing
				%>
				
				<%if txttranstype = "01-RECEIVE"  then %>
				<tr>
					<td class="jarakjobcode ratakiri"><%=txtcontroljob%></td>
					<td class="jarakjobname ratakiri" colspan="3" ><%=txtbankcredit%></td>
					<td class="jarakjobname ratakiri" colspan="2" ><%=txtnoref%></td>
					<td class="ratakanan" class="jarakamount ratakanan"><%=formatnumber(cdbl(txt_totalamount),2)%></td>
				</tr>
				<tr style="border-top:1px #000000 solid;">
					<td colspan="5" class="ratakanan tebal">Total</td>
					<td class="ratakanan tebal"><%=formatnumber(cdbl(txttotal),2)%></td>
					<td class="ratakanan tebal"><%=formatnumber(cdbl(txttotal_cr)+cdbl(txt_totalamount),2)%></td>
				</tr>
				<%End If %>
				
				
				<%if txttranstype = "02-PAYMENT"  then %>
				<tr>
					<td class="jarakjobcode ratakiri"><%=txtcontroljob%></td>
					<td class="jarakjobname ratakiri" colspan="3"><%=txtbankcredit%></td>
					<td class="jarakjobname ratakiri" colspan="2"><%=txtnoref%></td>
					<td class="ratakanan" class="jarakamount ratakanan"><%=formatnumber(cdbl(txt_totalamount),2)%></td>
				</tr>
				<tr style="border-top:1px #000000 solid;">
					<td colspan="5" class="ratakanan tebal">Total</td>
					<td class="ratakanan tebal"><%=formatnumber(cdbl(txttotal),2)%></td>
					<td class="ratakanan tebal"><%=formatnumber(cdbl(txttotal_cr)+cdbl(txt_totalamount),2)%></td>
				</tr>
				<%End If %>
				
				
				<tr>
				<td colspan="4" class="ratakiri" style="padding:5px;">
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
				<div class="title_terbilang">
				Terbilang:
				</div>
				<div class="terbilang" >
				<%  
					Dim txtSQL,rsCats,v_terbilang
					txtSQL= "select upper(replace(terbilang_indo("& txt_totalamount &"),'-',' ')) terbilang from dual"
					Set rsCats = Server.CreateObject("ADODB.Recordset")
				''	Response.Write txtSQL
					rsCats.Open txtSQL, connstr
					if Not rsCats.eof then
						v_terbilang =  rsCats("terbilang") 
					end if
					rsCats.close
					set rsCats = Nothing
				%>
				<%=v_terbilang%> RUPIAH
				</div>
				</td>
			</tr>
		</table>
		<br>
		<table border="0" width="100%" class="pv-ket" style="border-top:1px #000000 solid;border-bottom:1px #000000 solid; font-size:10px;line-height:14px;">
			<td>
			Keterangan:
			</td>
			<td><%=txtnoref%></td>		
		</table>

		<%
		Dim rsAutorize,qprd,intRecordsPerPage
		qprd = 	"   SELECT vouchercode," & _
					"          MAX (comments) comments," & _
					"          fcba," & _
					"          fcentry," & _
					"          currentauthorizelevel," & _
					"          MAX (lastupdate) lastupdate" & _
					"     FROM (  SELECT vouchercode," & _
					"                    '' comments," & _
					"                    fcba," & _
					"                    fcentry," & _
					"                    MAX (currentauthorizelevel) currentauthorizelevel," & _
					"                    MAX (lastupdate) lastupdate" & _
					"               FROM paymentvoucher_hist" & _
					"              WHERE vouchercode = '" & request("vouchercode") & "' AND fcba = '" & session("FCBA") & "'" & _
					"           GROUP BY vouchercode, fcba, fcentry" & _
					"           UNION ALL" & _
					"             SELECT vouchercode," & _
					"                    comments," & _
					"                    fcba," & _
					"                    fcentry," & _
					"                    MAX (currentauthorizelevel) currentauthorizelevel," & _
					"                    MAX (lastupdate) lastupdate" & _
					"               FROM pv_authorization_comments" & _
					"              WHERE vouchercode = '" & request("vouchercode") & "' AND fcba = '" & session("FCBA") & "'" & _
					"           GROUP BY vouchercode," & _
					"                    fcba," & _
					"                    fcentry," & _
					"                    comments)" & _
					" GROUP BY vouchercode," & _
					"          fcba," & _
					"          fcentry," & _
					"          currentauthorizelevel" & _
					" ORDER BY currentauthorizelevel"

		' response.Write qprd	
		Set rsAutorize = Server.CreateObject("ADODB.Recordset")
		rsAutorize.Open qprd, connstr
		if not rsAutorize.eof then
			comments = rsAutorize("comments")
			fcentry = rsAutorize("fcentry")
			lastupdate = rsAutorize("lastupdate")
		End if
		rsAutorize.close
		set rsAutorize = Nothing %>
		<table border="0" width="40%" id="product-table" class="pv-auth" style="border-bottom:1px #000000 solid;">
				<thead>
					<tr>
						<th>Authorizer</th>
						<th>Date</th>
					</tr>
				</thead>
				<tbody>
					

				<%	Set rsAutorize = Server.CreateObject("ADODB.Recordset")
					rsAutorize.Open qprd, connstr
					if Not rsAutorize.eof then	
					nCountofField=rsAutorize.fields.count-1
				%> 
					
					<%
					While Not rsAutorize.EOF
						%>
						
					<tr>
						<td align="center" style="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; width:170px;"><%=rsAutorize("fcentry")%>&nbsp;</td>
						<td align="center" style="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; width:205px;"><%=rsAutorize("lastupdate")%>&nbsp;</td>
					</tr>
					
					<%
						rsAutorize.MoveNext
						intRecordsPerPage = intRecordsPerPage - 1
					Wend  %> 
					
				<%	  
				end if 
				rsAutorize.Close
				set rsAutorize = Nothing%> 
					
				</tbody>
			</table>

	<br/>
</body>
</html>
 