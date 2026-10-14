<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->

<html>
<head>
<title>ePlaS - Payment Request</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>

<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">

 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<style>
	.lebarqty{
		width: 70px;
	}
	.lebaramount{
		width: 120px;
	}
	.ratatengah{
		text-align: center;
	}
	.ratakanan{
		text-align: right;
	}
</style>
</head>
<body>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>

<%	
Dim fccode,errno 

fccode = Request("fcCode")
errno  = Request("errno")
totpo  = Request("gandtotalPO")
totdp  = Request("DPPO")

if errno <> 0 then
   Response.Write "<div class = ""error"">" & "Payment Request melebihi Amount PO !" & totdp & " > " & totpo & "</div>"
End if
	
%>

 <div class="box_both" style="overflow:scroll;">

 <form id="formentry" name="frmEntry" method="post" action="paymentrequestproses.Asp" >
 <input type="hidden" id="prosestipe" name="prosestipe"  value="1">
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Invoice No</td>
			
			<td align="left"><input type="text" maxlength="50" id="fcCode" name="fcCode" size="30" value="<%=fcCode%>" readonly="true" style="background-color:#eee;" autocomplete="off"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">DP</td>
			<td align="left"><input type="text" id="downpayment" name="downpayment" size="30" readonly="true" value="0" onkeyup="hitungamount()"  style="background-color:#eee;text-align: right"></td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Ref. Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="refinvoice" size="50" value="<%=refinvoice%>"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">Other Charge Stock</td>
			<td align="left"><input type="text" id="otherchargestock" name="otherchargestock" size="30" value="0" onkeyup="hitungPpn();hitungamount()" style="text-align: right;" style="background-color:#eee;"></td>
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<input class="datepicker" onchange="" type="text" maxlength="10" name="DATECREATED" id="DATECREATED" size="10" value="<%=DATECREATED%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left" style="width:20%;">Amount</td>
			<td align="left"><input type="text" id="amount" name="amount" size="30" value="0"  readonly="true" style="background-color:#eee;text-align: right">
				<input type="hidden" id="amounthidden" name="amounthidden" size="30" value="0" >
			</td>	
			</td>
		</tr>
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="getFCBA()" class="chosen-select" style="width: 70%">
					<option value=""></option>
					<%
					qlov = "SELECT DISTINCT bu.fccompanycode fccode FROM user_fcba_tab uft INNER JOIN businessunit bu ON bu.fccode = uft.fcba WHERE fcusercode = '" & Session("CodeName") & "'"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if fccompanycode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<td align="left" style="width:20%;">VAT</td>
			<td align="left">
			    
				<br>
				<input type="text" id="vat" name="vat" size="2" value="0" style="text-align: right;"  onkeyup="hitungPpn(this);hitungamount()">&nbsp;<input type="text" id="totalppn" name="totalppn" size="20" value="0" style="text-align: right;background-color:#eee;" >
			</td> 
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<select name="businessunit" id="businessunit" onchange="showPO()" style=width:200px >
				</select>
			</td>

			<td align="left" style="width:20%;">Amount + VAT</td>
			<td align="left"><input type="text" id="amountvat" name="amountvat" size="30" value="0"  onkeyup="hitungamount()" readonly="true" style="background-color:#eee;text-align: right;" ></td>		
		</tr>
		<tr>
			<td align="left">PO Code</td>
			<td align="left">      
				<select name="pocode" id="pocode" onchange="getViewhdr();" style=width:200px >
				</select>
			</td>
			<td align="left" style="width:20%;"> Other Charge Non Stock</td>
			<td align="left">
					<input type="text" id="othercharge" name="othercharge" size="6" value="0" onkeyup="hitungPpn()" style="text-align: right;" >&nbsp;&nbsp;Diskon&nbsp;&nbsp;<input type="text" id="diskon" name="diskon" size="6" value="0" onkeyup="hitungPpn()" style="text-align: right;" >
			</td> 
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By</td>
			<td align="left">
				<select name="invoiceby" id="invoiceby"  style=width:200px >
				</select>
			</td>
				<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="">
			<td align="left" style="width:20%;">Total Amount</td>
			<td align="left"><input type="text" id="totalamount" name="totalamount" size="30" value="0" readonly="true" style="background-color:#eee;text-align: right;">
				<input type="hidden" id="totalamounthidden" name="totalamounthidden" size="30" value="0" readonly="true" style="background-color:#eee;text-align: right;">
				</td>
			
		</tr>
		<tr>
			<td align="left" style="width:20%;">Supplier Code</td>
			<td align="left"><input type="text" id="suppliercode" name="suppliercode" size="30" value="" readonly="true" style="background-color:#eee;style="width: 70%""></td>
			<td align="left">No Faktur</td>
			<td align="left">
				<input type="text" maxlength="100" name="no_faktur" size="30" value="" autocomplete="off" >
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Supplier Name</td>
			<td align="left"><input type="text" id="suppliername" name="suppliername" size="30" value="" readonly="true" style="background-color:#eee;"></td>
			<td align="left"> Faktur Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="fakturdate" size="10" value="" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;"></td>
			<td align="left" style="width:20%;"></td>
			<td align="left"> Percentage (%)</td>
			<td align="left">
				<input class="number" type="text" maxlength="10" name="persenheader" id="persenheader" size="5" value="" autocomplete="Off"> %
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Supplier Name</td>
			<td align="left"><input type="text" id="suppliername" name="suppliername" size="30" value="" readonly="true" style="background-color:#eee;"></td>
			<td align="left"> Amount</td>
			<td align="left">
				<input class="number" type="text" maxlength="10" name="amounthdr" id="amounthdr" size="30" value="" autocomplete="Off">
				<input type="hidden" maxlength="10" name="amountheader" id="amountheader" size="30" value="" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<br><br>
				<input Type="hidden" Value="1" Name="cmdView">
				<input Type="hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Simpan">&nbsp;&nbsp;
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Purchase"
				dataSubmodule 	= "Invoice"
				dataDetail 		= "NULL"
				%>
				<input Type="Reset"  name="Add" value="Batal">&nbsp;&nbsp;
			</td>
		</tr>
    </table> <br>  
   <table class="data" border="1" width="100%">
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
			<td align="center">No</td>
			<td align="center">PR Code</td>
			<td align="center">Stock Code</td>
			<td align="center">Stock Desc.</td>
			<td align="center">Qty PO</td>
			<td align="center">Qty Payment</td>
			<td align="center">Price</td>
			<td align="center">Percent (%)</td>
			<td align="center">Amount (Rp)</td>
			<td align="center">Amount PO</td>
	   </tr>
	   <tr>
		<tbody id="bodytable">
		</tbody>
	   </tr>
   </TABLE>
   </div>
   <input type="hidden" name="jumldata" id="jumldata" value="0" />
   </form> 

</div></div>
  <!--#include file="bottom.jsp"-->
	<script src="/js/jquery.inputmask.bundle.min.js"></script>
	<script src="/js/inputmask.binding.js"></script>
	<script src="plugins/chosen/chosen.jquery.min.js"></script>
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
	<script type="text/javascript">	
		// $(".chosen-select").empty();
		
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"45%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		// $(".chosen-select").trigger("chosen:updated");

		function hanyaAngka(evt) {
			evt.value = evt.value.replace(/[^0-9,.-]/g,"");
		}

	</script>
<script>

	function getFCBA()
	{
		var company = $("#fccompanycode").val();
		showFCBA(company);
		showInvoiceBy(company);

	}

	function showFCBA(company) {
		$('#businessunit').empty();
		$('#businessunit').select2({
			ajax: {
				delay: 1000,
				url : 'getlistfcbaajax.asp?fccompanycode=' + company,
				dataType: 'json',
				type: "POST",
				allowClear: true,
				data: function (params) {
						var query = {
										keyword: params.term
									}
						return query;
				}
			}
		});
	}

	function showInvoiceBy(company) {
		$('#invoiceby').empty();
		$('#invoiceby').select2({
			ajax: {
				delay: 1000,
				url : 'getlistfcbaajax.asp?fccompanycode=' + company,
				dataType: 'json',
				type: "POST",
				allowClear: true,
				data: function (params) {
						var query = {
										keyword: params.term
									}
						return query;
				}
			}
		});
	}

	function showPO() {
		var fcba = $("#businessunit").val();
		$('#pocode').empty();
		$('#pocode').select2({
			ajax: {
				delay: 1000,
				url : 'getporeqpay.asp?fcba=' + fcba,
				dataType: 'json',
				type: "POST",
				allowClear: true,
				data: function (params) {
						var query = {
										keyword: params.term
									}
						return query;
				}
			}
		});
	}

	function getViewhdr(){
		var pocode = $("#pocode").val();
		var validasi = 1;
		$.post("getporeqpayview.asp?pocode=" + pocode + "&validasi=" + validasi, function(data) {
			var json = JSON.parse(data);
			var suppliercode 	= json.dataview[0].suppliercode;
			var suppliername 	= json.dataview[0].fcname;
			// var dp 				= json.dataview[0].dp;
			// var ppn 			= json.dataview[0].ppn;
			// var amount 			= json.dataview[0].totalamount;
			// var totalppn 		= json.dataview[0].totalppn;
			// var totalamountppn 	= json.dataview[0].totalamountppn;
			 var totalcharge 	= json.dataview[0].totalcharge;
			// var grandtotal 		= json.dataview[0].grandtotal;
			 $("#suppliercode").val(suppliercode);
			 $("#suppliername").val(suppliername);
			// $("#downpayment").val(dp);
			// $("#vat").val(ppn);
			// $("#amount").val(amount);
			// $("#totalppn").val(totalppn);
			// $("#amountvat").val(totalamountppn);
			 $("#otherchargestock").val(totalcharge);
			// $("#totalamount").val(grandtotal);
			
		});

		tampilTabel();
	}

	function numformat(val) {
		var nStr = val.toString();
		nStr = nStr.replace(/\,/g, "");
		var x = nStr.split('.');
		var x1 = x[0];
		var x2 = x.length > 1 ? '.' + x[1] : '';
		var rgx = /(\d+)(\d{3})/;
		while (rgx.test(x1)) {
			x1 = x1.replace(rgx, '$1' + '.' + '$2');
		}
		return x1 + x2;
	}

	function tampilTabel() {
			var pocode = $("#pocode").val();
			var fcba = $("#businessunit").val();
			var jumlinv = parseInt($('#jumldata').val());
			for(var n=0; n<jumlinv; n++) {
				Spry.Widget.Utils.destroyWidgets('qty' + n);
				Spry.Widget.Utils.destroyWidgets('priceqty' + n);
			}
			var validasi = 2;
			if (pocode != "") {
				$.post("getporeqpayview.asp?pocode=" + pocode + "&validasi=" + validasi + "&fcba=" + fcba, function(data) {
					$('#bodytable').empty();
					var json  = JSON.parse(data);
					var jumlview  = json.dataview.length;
					var no = 1;
					var totalAmountPO = 0;
					
					for (var i = 0; i < jumlview; i++) {
						var unitprice 	= numformat(json.dataview[i].unitprice);
						var amount 		= numformat(json.dataview[i].amount);
						var amountPO   = json.dataview[i].qtysisa * json.dataview[i].unitprice;
						var flaq = json.dataview[i].flaq;
						
						totalAmountPO += amountPO;

						if(flaq == 0){
							var qty_editable = '<td align="center" class="ratatengah lebarqty"><input type="text" class="ratatengah lebarqty"  onkeyup="getQtyAmount('+i+', this, '+json.dataview[i].qtysisa+')"  name="qty' + i + '" id="qty' + i + '" value="'+json.dataview[i].qtysisa+'"/></td>';
						}else{
							var qty_editable = '<td align="center" class="ratatengah lebarqty"><input type="text" class="ratatengah lebarqty" readonly="true" style="background-color: #eee;" onkeyup="getQtyAmount('+i+', this, '+json.dataview[i].qtysisa+')"  name="qty' + i + '" id="qty' + i + '" value="'+json.dataview[i].qtysisa+'"/></td>';
						}

						var docketlist = '';
						 docketlist += '<tr>';
						 docketlist += '<td align="center">' + no + '</td>';
						 docketlist += '<td align="left">'+ json.dataview[i].prcode + '<input type="hidden" name="prcode' + i + '" id="prcode' + i + '" value="' + json.dataview[i].prcode + '"/></td>';
						 docketlist += '<td align="left">' + json.dataview[i].itemcode + '<input type="hidden" name="itemcode' + i + '" id="itemcode' + i + '" value="' + json.dataview[i].itemcode + '"/></td>';
						 docketlist += '<td align="left">'+ json.dataview[i].fcname + '</td>';
						 docketlist += '<td align="center" class="ratatengah lebarqty">'+ json.dataview[i].qtysisa + '<input type="hidden" class="ratatengah lebarqty" readonly  name="qtypo' + i + '" id="qtypo' + i + '" value="' + json.dataview[i].qtysisa + '"/></td>';
						 docketlist += qty_editable;
						 docketlist += '<td align="right"><input type="text" class="ratatengah lebarqty" onkeyup="getAmount('+i+',this)"  name="priceqty' + i + '" id="priceqty' + i + '" value="' + json.dataview[i].unitprice + '"/><input type="hidden" name="price' + i + '" id="price' + i + '" value="' + json.dataview[i].unitprice + '"/></td>';
						 docketlist += '<td align="center" class="ratatengah lebarqty"><input type="text" class="ratatengah lebarqty" onkeyup="getPercent('+i+',this)"  name="percent' + i + '" id="percent' + i + '" value="100"/></td>';
						 docketlist += '<td align="right"><input type="text" class="ratatengah lebarqty" style="width: 100px;" onkeyup="getAmount('+i+',this)"  name="amountrp' + i + '" id="amountrp' + i + '" value="' + numformat(json.dataview[i].qtysisa*json.dataview[i].unitprice) + '"/><input type="hidden" name="amtrp' + i + '" id="amtrp' + i + '" value="' + json.dataview[i].qtysisa*json.dataview[i].unitprice + '"/></td>';
						 docketlist += '<td align="right" class="lebaramount"><span  class="lebaramount" id="amountview'+i+'">'+numformat(json.dataview[i].qtysisa*json.dataview[i].unitprice)+'</span><input class="lebaramount ratakanan" type="text" name="amountdtl' + i + '" id="amountdtl' + i + '"/></td>';
						 docketlist += '</tr>';
						 $('#bodytable').append(docketlist);
						no++;
						$("#jumldata").val(jumlview);
						new Spry.Widget.ValidationTextField("qty"+i , "real", {useCharacterMasking:true});
					}
					
					$('#amount').val(numformat(totalAmountPO));
					$('#amounthidden').val(parseInt(totalAmountPO));
					$('#totalamount').val(numformat(totalAmountPO));
					$('#totalamounthidden').val(parseInt(totalAmountPO));
					$('#amountvat').val(parseInt(totalAmountPO));

					let persen = (totalAmountPO / totalAmountPO) * 100;
					$('#persenheader').val(persen);
					$('#amounthdr').val(numformat(totalAmountPO));
					$('#amountheader').val(totalAmountPO);
				});
			} else {
				$('#bodytable').empty();
			}
	}

	function getAmount(i)
	{
		let harga 	= parseFloat($('#priceqty' + i).val());
		let qty 	= parseFloat($('#qty' + i).val());
		let amountrp = $('#amountrp' + i).val();
		let total 	= harga * qty;

		let totamount	= 0;
		let jumlview	= parseInt($('#jumldata').val());
		
		for(let i=0; i < jumlview; i++) {
			totamount += parseInt($('#amtrp'+i).val() || 0);
		}

		let amountheader = parseInt($('#amountheader'+i).val());
		let hitungPersen = (totamount/amountheader) * 100;
		$('#persenheader').val(hitungPersen);

		let amtrp 		= $('#amtrp' + i).val();
		let getPercent 	= (totamount/amtrp) * 100;
		let hasilQty 	= amountrp/harga;

		console.log(getPercent);

		$('#amountdtl' + i).val(total);
		$('#amountview' + i).text(total);
		
		$('#amountview' + i).text(numformat(total));
		$('#amountrp' + i).val(numformat(amountrp));
		parseFloat($('#percent' + i).val(getPercent));
		parseFloat($('#qty' + i).val(hasilQty));
		getPuttotal();
	}

	function getQtyAmount(i, j, k)
	{
		var harga 	= parseFloat($('#priceqty' + i).val());
		var qty 	= parseInt($('#qty' + i).val());
		var qtypo 	= parseInt($('#qtypo' + i).val());

		let totamount	= 0;
		let jumlview	= parseInt($('#jumldata').val());
		
		for(let i=0; i < jumlview; i++) {
			totamount += parseInt($('#amtrp'+i).val() || 0);
		}

		let amountheader = parseInt($('#amountheader'+i).val());
		let hitungPersen = (totamount/amountheader) * 100;
		$('#persenheader').val(hitungPersen);

		var total 	= harga * qty;

		if(qty > k){
			alert("QTY Payment melebihi QTY PO!!");
			$('#qty' + i).val(k);
			$('#amountdtl' + i).val(harga * k);
			var total_fmt	= numformat(harga * k);
			$('#amountview' + i).text(numformat(total_fmt));
			$('#amountrp' + i).val(numformat(total_fmt));
		}else{
			$('#amountdtl' + i).val(total);
			var total_fmt		= numformat(total);
			$('#amountview' + i).text(numformat(total_fmt));
			$('#amountrp' + i).val(numformat(total_fmt));
			getPuttotal();
		}

		let setPercent = (qty/qtypo) * 100;
		parseFloat($('#percent' + i).val(setPercent));
	}

	function getPercent(i)
	{
		let persen 		= parseFloat($('#percent' + i).val());
		let harga 		= parseFloat($('#priceqty' + i).val());
		var qtypo 		= parseInt($('#qtypo' + i).val());

		let totamount	= 0;
		let jumlview	= parseInt($('#jumldata').val());
		
		for(let i=0; i < jumlview; i++) {
			totamount += parseInt($('#amtrp'+i).val() || 0);
		}

		let amountheader = parseInt($('#amountheader').val());
		let hitungPersen = (totamount/amountheader) * 100;
		$('#persenheader').val(hitungPersen);

		console.log("ASU");

		let total 		= harga * qtypo;
		let hasil 		= (persen / 100) * total;
		let hasilQty 	= (persen / 100) * qtypo;
		
		$('#amountdtl' + i).val(hasil);
		$('#amountview' + i).text(numformat(hasil));
		$('#amountrp' + i).val(numformat(hasil));

		parseFloat($('#qty' + i).val(hasilQty));

		getPuttotal();
	}

	function getPuttotal() {
		let totamount	= 0;
		let jumlview	= parseInt($('#jumldata').val());
		
		for(let i=0; i < jumlview; i++) {
			totamount += parseInt($('#amountdtl'+i).val());
		}	
		let total_fmt		= numformat(totamount);
		
		$('#amount').val(total_fmt);
		$('#amounthidden').val(totamount);
		$('#totalamount').val(total_fmt);
		$('#totalamounthidden').val(totamount);
		$('#amounthdr').val(total_fmt);
		$('#amountheader').val(total_fmt);
		
		hitungPpn();
}


	function hitungPpn()
	{
		var amounthidden 		= parseInt($('#amounthidden').val());
		var vat 				= parseInt($('#vat').val());
		var otherchargestock	= parseInt($('#otherchargestock').val());
		var othercharge			= parseInt($('#othercharge').val());
		var diskon				= parseInt($('#diskon').val());
		if(otherchargestock>0){
			var ttlppn 			= (amounthidden+otherchargestock)*(vat/100);			
		}else{
			var ttlppn			= amounthidden*(vat/100);
		}
		
		var amountvat 			= ttlppn+amounthidden;
		var grandtotal 			= amountvat+otherchargestock+othercharge-diskon;
		var grandtotal_fmt 		= numformat(grandtotal);
		$('#totalppn').val(ttlppn);
		$('#amountvat').val(amountvat);
		$('#totalamount').val(grandtotal_fmt);
		$('#totalamounthidden').val(grandtotal);
	}
amount
	function hitungamount()
	{
		var dp 					= parseInt($('#downpayment').val());
		var amounthidden 		= parseInt($('#amounthidden').val());
		var amount	 			= parseInt($('#amount').val());
		var totalppn	 		= parseInt($('#totalppn').val());
		var vat 				= parseInt($('#vat').val());
		var otherchargestock	= parseInt($('#otherchargestock').val());
		var othercharge			= parseInt($('#othercharge').val());
		var diskon				= parseInt($('#diskon').val());
		if (dp>0){
				var amount 				= parseInt(dp)+parseInt(otherchargestock);
			
				//var amountvat 			= parseInt(amount+otherchargestock+totalppn)*parseInt(vat/100);
				
		}
		var amountfrm 		= numformat(amount);
		$('#amount').val(amount);
		$('#totalamount').val(amountfrm);
		$('#totalamounthidden').val(amount);
		$('#amountvat').val(amount);
		// if(dp>0){
		// 	var ttlppn 			= (amounthidden+otherchargestock)*(vat/100);			
		// }else{
		// 	var ttlppn			= amounthidden*(vat/100);
		// }
		
		// var amountvat 			= ttlppn+amounthidden;
		// var grandtotal 			= amountvat+otherchargestock+othercharge-diskon;
		// var grandtotal_fmt 		= numformat(grandtotal);
		// $('#totalppn').val(ttlppn);
		// $('#amountvat').val(amountvat);
		// $('#totalamount').val(grandtotal_fmt);
		// $('#totalamounthidden').val(grandtotal);
	}
	
	
	new Spry.Widget.ValidationTextField("DATECREATED", "none");
	new Spry.Widget.ValidationTextField("vat" , "integer", {useCharacterMasking:true});
	new Spry.Widget.ValidationTextField("otherchargestock" , "integer", {useCharacterMasking:true});
	new Spry.Widget.ValidationTextField("othercharge" , "integer", {useCharacterMasking:true});
	new Spry.Widget.ValidationTextField("diskon" , "integer", {useCharacterMasking:true});


</script>
</body>
</html>
