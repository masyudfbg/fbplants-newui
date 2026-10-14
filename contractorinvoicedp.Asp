<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->

<html>
<head>
<title>ePlaS - Conctractor Invoice DP</title>
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
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
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

 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Contractor Invoice DP</div><% end if %>
 <form id="formentry" name="frmEntry" method="post" action="contractorinvdpproses.Asp" >
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Invoice No</td>
			
			<td align="left"><input type="text" maxlength="50" id="fcCode" name="fcCode" size="30" value="<%=fcCode%>" readonly="true" style="background-color:#eee;" autocomplete="off"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">Qty</td>
			<td align="left"><input type="text" id="qty" name="qty" size="30" value="0"  readonly="true" style="background-color:#eee;text-align: right"></td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Ref. Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="refinvoice" size="50" value="<%=refinvoice%>"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">Rate</td>
			<td align="left">
				<input type="text" id="rate" name="rate" size="30" value="0" onkeyup="hitungPpn()" style="text-align: right;" style="background-color:#eee;">
				<input type="hidden" id="rateback" name="rateback" size="30" value="0"  style="text-align: right;" style="background-color:#eee;">
			</td>
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<input class="datepicker" onchange="" type="text" maxlength="10" name="DATECREATED" id="DATECREATED" size="10" value="<%=DATECREATED%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left" style="width:20%;">Amount</td>
			<td align="left"><input type="text" id="amount" name="amount" size="30" value="0"  readonly="true" style="background-color:#eee;text-align: right">
				<input type="hidden" id="amountback" name="amountback" size="30" value="0" >
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
			<td align="left" style="width:20%;">DP</td>
			<td align="left">
			    
				<br>
				<input type="text" id="dp" name="dp" size="30" value="0" onkeyup="hitungRemain();" style="text-align: right;" >
				
			</td> 
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<select name="businessunit" id="businessunit" onchange="showPO()" style=width:200px >
				</select>
			</td>

			<td align="left" style="width:20%;">Remain</td>
			<td align="left">
				<input type="text" id="remain" name="remain" size="30" value="0" readonly="true" style="background-color:#eee;text-align: right;" >
				<input type="hidden" id="remainback" name="remainback" size="30" value="0" readonly="true" style="background-color:#eee;text-align: right;" >
				
			</td>		
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By</td>
			<td align="left">
				<select name="invoiceby" id="invoiceby"  style=width:200px >
				</select>
			</td>
				<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="">
			
			
		</tr>
		<tr>
			<td align="left">Agreement Code</td>
			<td align="left">      
				<select name="agrementcode" id="agrementcode" onchange="getViewnominal()" style=width:200px >
				</select>
			</td>
			
		</tr>
		
		<tr>
			<td align="left" style="width:20%;">Contractor Code</td>
			<td align="left"><input type="text" id="contractorcode" name="contractorcode" size="30" value="" readonly="true" style="background-color:#eee;style="width: 70%""></td>
			
		</tr>
		<tr>
			<td align="left" style="width:20%;">Contractor Name</td>
			<td align="left"><input type="text" id="contractorname" name="contractorname" size="30" value="" readonly="true" style="background-color:#eee;"></td>
			
		</tr>
		<tr>
			<td colspan="2"  align="left">
			
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
		$('#agrementcode').empty();
		$('#agrementcode').select2({
			ajax: {
				delay: 1000,
				url : 'getagreementcode.asp?fcba=' + fcba,
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

	function getViewnominal(){
		 var agrementcode = $("#agrementcode").val();
		$.post("getagreementnominal.asp?agrementcode=" + agrementcode, function(data) {
			var json = JSON.parse(data);
			var contractorcode 	= json.results[0].contractorcode;
			var fcname 	= json.results[0].fcname;
			var qty 	= json.results[0].qty;
			var rate 	= json.results[0].rate;
			var ratefrm = numformat(json.results[0].rate);
			var amount 	= json.results[0].amount;
			var amountfrm 	= numformat(json.results[0].amount);
			
			 $("#contractorcode").val(contractorcode);
			 $("#contractorname").val(fcname);
			 $("#qty").val(qty);
			 $("#rate").val(ratefrm);
			 $("#rateback").val(rate);
			 $("#amount").val(amountfrm);
			 $("#amountback").val(amount);
			
			
		});
	
	}

	function hitungRemain()
	{
		var amount 	= parseInt($('#amountback').val());
		var dp 		= parseInt($('#dp').val());
		var sisa	= amount-dp;
		var sisafrm	= numformat(sisa);
		 $('#remain').val(sisafrm);
		 $('#remainback').val(sisa);
		//alert(sisa);
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
	 new Spry.Widget.ValidationTextField("dp" , "integer", {useCharacterMasking:true});
	// new Spry.Widget.ValidationTextField("otherchargestock" , "integer", {useCharacterMasking:true});
	// new Spry.Widget.ValidationTextField("othercharge" , "integer", {useCharacterMasking:true});


</script>
</body>
</html>
