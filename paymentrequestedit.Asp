<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->

<html>
<head>
<title>ePlaS - Edit Payment Request</title>
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
	   Module : Payment Request / Edit
	</div>

<%	
Dim fcba,invoicecode,refno,invoicedate,company,pocode,invoiceby,supcode,supname
Dim dp,otherchargestock,amount,vatptg,vatrp,amount_vat,otherchargenonstock,ttlamount,nofaktur,fakturdate,diskon
Dim sql,rsview


fcba 		= Request("fcba")
invoicecode = Request("invoicecode")


sql = "Select I.INVOICECODE, I.INVOICEDATE, I.SUPPLIERCODE,S.FCNAME, I.COMPANY, I.FCBA, I.POCODE, I.AMOUNT, I.VAT, I.AMOUNTVAT, I.TOTALAMOUNT, " & _ 
       " I.NO_FAKTUR, I.FAKTURDATE, I.REFINVOICE, I.OTHERCHARGE, I.DP, " & _
       " I.INVOICE_BY, I.TOTALPPN, I.OTHERCHARGESTOCK, I.PAYMENT,I.DISKON " & _
" From   INVOICEDPHDR I,SUPPLIER S " & _
" WHERE I.SUPPLIERCODE=S.FCCODE AND I.FCBA = S.FCBA and I.INVOICECODE='" & invoicecode & "'" & _ 
" AND I.FCBA='" & fcba & "'" 
set rsview = connStr.Execute(sql) 
'Response.Write sql
if not rsview.eof then 
	invoicecode = rsview("INVOICECODE")
	refno = rsview("REFINVOICE")
	invoicedate = rsview("INVOICEDATE")
	company = rsview("COMPANY")
	fcba = rsview("FCBA")
	pocode = rsview("POCODE")
	invoiceby = rsview("INVOICE_BY")
	supcode = rsview("SUPPLIERCODE")
	supname = rsview("FCNAME")
	dp = rsview("DP")
	otherchargestock = rsview("OTHERCHARGESTOCK")
	diskon = rsview("DISKON")
	amount = rsview("AMOUNT")
	vatptg = rsview("VAT")
	vatrp = rsview("TOTALPPN")
	amount_vat = rsview("AMOUNTVAT")
	otherchargenonstock = rsview("OTHERCHARGE")
	ttlamount = rsview("TOTALAMOUNT")
	nofaktur = rsview("NO_FAKTUR")
	fakturdate = rsview("FAKTURDATE")
end if
	
%>

 <div class="box_both" style="overflow:scroll;">

 <form id="formentry" name="frmEntry" method="post" action="paymentrequestproses.Asp" >
 <input type="hidden" id="prosestipe" name="prosestipe"  value="2">
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Invoice No</td>
			
			<td align="left"><input type="text" maxlength="50" id="fcCode" name="fcCode" size="30" value="<%=invoicecode%>" readonly="true" style="background-color:#eee;" autocomplete="off"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">DP</td>
			<td align="left"><input type="text" id="downpayment" name="downpayment" size="30" value="<%=dp%>"  readonly="true" style="background-color:#eee;text-align: right"></td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Ref. Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="refinvoice" size="50" value="<%=refno%>"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">Other Charge Stock</td>
			<td align="left"><input type="text" id="otherchargestock" name="otherchargestock" size="30" value="<%=otherchargestock%>" onkeyup="hitungPpn()" style="text-align: right;" style="background-color:#eee;"></td>
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<input class="datepicker" onchange="" type="text" maxlength="10" name="DATECREATED" id="DATECREATED" size="10" value="<%=invoicedate%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left" style="width:20%;">Amount</td>
			<td align="left"><input type="text" id="amount" name="amount" size="30" value="<%=amount%>"  readonly="true" style="background-color:#eee;text-align: right">
				<input type="hidden" id="amounthidden" name="amounthidden" size="30" value="<%=amount%>" >
			</td>	
			</td>
		</tr>
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<%=company%>
				<input type="hidden" id="fccompanycode" name="fccompanycode" size="2" value="<%=company%>" style="text-align: right;" >
				<font color="#FF0000">*</font>
			</td>
			<td align="left" style="width:20%;">VAT</td>
			<td align="left">
			    
				<br>
				<input type="text" id="vat" name="vat" size="2" value="<%=vatptg%>" style="text-align: right;"  onkeyup="hitungPpn(this);">&nbsp;<input type="text" id="totalppn" name="totalppn" size="20" value="<%=vatrp%>" style="text-align: right;background-color:#eee;" >
			</td> 
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<%=fcba%>
				<input type="hidden" id="businessunit" name="businessunit" size="30" value="<%=fcba%>">
			</td>

			<td align="left" style="width:20%;">Amount + VAT</td>
			<td align="left"><input type="text" id="amountvat" name="amountvat" size="30" value="<%=amount_vat%>" readonly="true" style="background-color:#eee;text-align: right;" ></td>		
		</tr>
		<tr>
			<td align="left">PO Code</td>
			<td align="left">      
				<%=pocode%>
				<input type="hidden" id="pocode" name="pocode" size="30" value="<%=pocode%>">
			</td>
			<td align="left" style="width:20%;"> Other Charge Non Stock</td>
			<td align="left">
					<input type="text" id="othercharge" name="othercharge" size="6" value="<%=otherchargenonstock%>" onkeyup="hitungPpn()" style="text-align: right;" >&nbsp;&nbsp;Diskon&nbsp;&nbsp;<input type="text" id="diskon" name="diskon" size="6" value="<%=diskon%>" onkeyup="hitungPpn()" style="text-align: right;" >
			</td> 
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By</td>
			<td align="left">
				<select name="invoiceby" id="invoiceby"  style=width:200px >
					<option value="<%=invoiceby%>"><%=invoiceby%></option>
				</select>
			</td>
				<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="">
			<td align="left" style="width:20%;">Total Amount</td>
			<td align="left"><input type="text" id="totalamount" name="totalamount" size="30" value="<%=ttlamount%>" readonly="true" style="background-color:#eee;text-align: right;">
				<input type="hidden" id="totalamounthidden" name="totalamounthidden" size="30" value="<%=ttlamount%>" readonly="true" style="background-color:#eee;text-align: right;">
				</td>
			
		</tr>
		<tr>
			<td align="left" style="width:20%;">Supplier Code</td>
			<td align="left"><input type="text" id="suppliercode" name="suppliercode" size="30" value="<%=supcode%>" readonly="true" style="background-color:#eee;style="width: 70%""></td>
			<td align="left">No Faktur</td>
			<td align="left">
				<input type="text" maxlength="100" name="no_faktur" size="30" value="<%=nofaktur%>" autocomplete="off" >
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Supplier Name</td>
			<td align="left"><input type="text" id="suppliername" name="suppliername" size="30" value="<%=supname%>" readonly="true" style="background-color:#eee;"></td>
			<td align="left"> Faktur Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="fakturdate" size="10" value="<%=fakturdate%>" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Submit"  name="Add" value="Simpan">&nbsp;&nbsp;
				<input Type="Reset"  name="Add" value="Batal">&nbsp;&nbsp;
			</td>
		</tr>
    </table> <br>  
   <table class="data" border="1" width="100%">
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
			<td align="center">No</td>
			<td align="center">Stock Code</td>
			<td align="center">Stock Desc.</td>
			<td align="center">Qty Payment</td>
			<td align="center">Price</td>
			<td align="center">Amount PO</td>
	   </tr>
	   <%
	   		dim sqldetail,rsdetail,No,totalqty,totalamount,rsArray,intRow,Noloop
			sqldetail = "Select I.INVOICECODE, I.ITEMCODE,S.FCNAME, I.QUANTITY, I.PRICE, I.AMOUNT " & _
			" From   INVOICEDPDETAIL I,STOCK S " & _
			" WHERE I.ITEMCODE=S.FCCODE AND I.FCBA = S.FCBA AND I.FCBA='" & fcba & "' AND I.INVOICECODE='" & invoicecode & "'"
			set rsdetail = connStr.Execute(sqldetail) 

	   %>
		<% 
			No = 1
			Noloop = 0
			totalqty = 0
			totalamount = 0
			intRow = 0
			While Not rsdetail.EOF

		%>
	   <tr>
			<td align="center"><%=No%></td>
			<td align="center"><input style="text-align:center" type="hidden" id="itemcode<%=Noloop%>" name="itemcode<%=Noloop%>" size="5" value="<%=rsdetail("ITEMCODE")%>"><%=rsdetail("ITEMCODE")%></td>
			<td align="left"><%=rsdetail("FCNAME")%></td>
			<td align="center"><input style="text-align:center" type="text" id="qty<%=Noloop%>" onkeyup="getAmount('<%=Noloop%>',this)" name="qty<%=Noloop%>" size="5" value="<%=rsdetail("QUANTITY")%>"></td>
			<td align="right"><input style="text-align:center" type="text" id="price<%=Noloop%>" onkeyup="getAmount('<%=Noloop%>',this)" name="price<%=Noloop%>" size="5" value="<%=rsdetail("PRICE")%>"></td>
			<td align="right"><input style="text-align:center" type="hidden" id="amountdtl<%=Noloop%>" name="amountdtl<%=Noloop%>" size="5" value="<%=rsdetail("AMOUNT")%>"><span id="amountview<%=Noloop%>"><%=formatnumber(rsdetail("AMOUNT"),0)%></span></td>
	   </tr>

	   <%
	   
	    No = No + 1
	    Noloop = Noloop + 1
		totalqty = cdbl(totalqty)+cdbl(rsdetail("QUANTITY"))
		totalamount = cdbl(totalamount)+cdbl(rsdetail("AMOUNT"))
		
		
		rsdetail.MoveNext
		Wend
		rsdetail.close()
		
		%>
		<tr>
		<td align="right" colspan=3>Total</td>
		<td align="center"><input  class="ratatengah" type="text" readonly size=5 name="ttlqty" id="ttlqty" value="<%=totalqty%>"></td>
		<td align="right"></td>
		<td align="right"><input type="hidden" name="ttlamountbwhback" id="ttlamountbwhback" value="<%=totalamount%>"><input class="ratakanan" type="text" readonly size=12 name="ttlamountbwhview" id="ttlamountbwhview" value="<%=formatnumber(totalamount,0)%>"></td>
		</tr>
   </TABLE>
   </div>
   <%
  
	dim sqldd,rsdd
	sqldd = "Select count(I.INVOICECODE)INVOICECODE " & _
	" From   INVOICEDPDETAIL I,STOCK S " & _
	" WHERE I.ITEMCODE=S.FCCODE AND I.FCBA = S.FCBA AND I.INVOICECODE='" & invoicecode & "'"
	set rsdd = connStr.Execute(sqldd) 
	

	While Not rsdd.EOF %>
   <input type="hidden" name="jumldata" id="jumldata" value="<%=rsdd("INVOICECODE")%>" />
   <%

	rsdd.MoveNext
	Wend
	rsdd.close()
	%>
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

	$(function() {
      
		var company = $("#fccompanycode").val();
		showInvoiceBy(company);
    });


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
		//$('#invoiceby').empty();
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
			 $("#suppliercode").val(suppliercode);
			 $("#suppliername").val(suppliername);
					
		});

		tampilTabel();
	}

	function getAmount(i)
	{
		var harga 	= parseFloat($('#price' + i).val());
		var qty 	= parseFloat($('#qty' + i).val());
		var total 	= harga * qty;
		
		var total_fmt		= numformat(total);
		$('#amountdtl' + i).val(total);
		$('#amountview' + i).text(total_fmt);
		getPuttotal();
		//alert(total);
	}

	function getPuttotal() {
		var totamount	= 0;
		var totqty	= 0;
		var jumlview	= parseInt($('#jumldata').val());
		for(var i=0; i<jumlview; i++) {
			totamount	+= parseInt($('#amountdtl'+i).val());
			totqty	+= parseInt($('#qty'+i).val());
			//alert(totqty);
		}	
		var total_fmt		= numformat(totamount);
		$('#ttlqty').val(totqty);
		$('#amount').val(total_fmt);
		$('#amounthidden').val(totamount);
		$('#totalamount').val(total_fmt);
		$('#totalamounthidden').val(totamount);
		$('#ttlamountbwhview').val(total_fmt);
		$('#ttlamountbwhback').val(totamount);
		hitungPpn();
}


	function hitungPpn()
	{
		var amounthidden 		= parseInt($('#amounthidden').val());
		var vat 				= parseInt($('#vat').val());
		var otherchargestock	= parseInt($('#otherchargestock').val());
		var othercharge			= parseInt($('#othercharge').val());
		var diskon				= parseInt($('#diskon').val());
		var ttlppn				= (amounthidden + otherchargestock) *(vat/100);
		var amountvat 			= ttlppn+amounthidden;
		var grandtotal 			= amountvat+otherchargestock+othercharge-diskon;
		var grandtotal_fmt 		= numformat(grandtotal);
		$('#totalppn').val(ttlppn);
		$('#amountvat').val(amountvat);
		$('#totalamount').val(grandtotal_fmt);
		$('#totalamounthidden').val(grandtotal);
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
	
	new Spry.Widget.ValidationTextField("DATECREATED", "none");
	new Spry.Widget.ValidationTextField("vat" , "integer", {useCharacterMasking:true});
	new Spry.Widget.ValidationTextField("otherchargestock" , "integer", {useCharacterMasking:true});
	new Spry.Widget.ValidationTextField("othercharge" , "integer", {useCharacterMasking:true});


</script>
</body>
</html>
