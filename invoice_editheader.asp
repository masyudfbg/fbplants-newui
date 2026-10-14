<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Invoice Non Purchase Order</title>
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

<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">

 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<script>
	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script>--> 
<script language="JavaScript">

function harusangka(evt) {
    var e = evt || window.event;
    var key = e.keyCode || e.which;

    if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
    // numbers   
    key >= 48 && key <= 57 ||
	e.ctrlKey && key == 97 ||
    // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
    // left and right arrows
    key == 37 || key == 39 ||
    // Del and Ins
	key == 46 ) {
        // input is VALID
		}
	else {
			// input is INVALID
			e.returnValue = false;
			if (e.preventDefault) e.preventDefault();
		}
}

function hitungrumus() {
		var skg = $("#kg").val();
		var srate = $("#rate").val(); 
		var newskg = skg.replace(/,/g, '');
		var newsrate = srate.replace(/,/g, '');
		
		//Untuk menghitung Total Janjang
		totalamount = (newskg * newsrate).toFixed(2); 
		if (!isNaN(totalamount)) {
    	 	$("#amount").val(totalamount);
    	}
    	else{
    		$("#amount").val();
    	}
	}
	
</script>


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
 <div class="box_both" style="overflow:scroll;"> 
 <%
 '-------------awal-----------------------
 
 Dim xCode,fcCode,cmdView,PoCode,PoDate,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim RowStateId,RowStateHdr
 Dim PRCODE, PRLINEID
 Dim tempStringError, rsTempPurchaseRequest
 Dim totalPRData
 Dim ApprovedQty, ORDEREDQTY
 Dim strQueryStatus, rsQueryStatus
 Dim sqlupdateqtypr, rsupdateqtypr
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = "" 
 Dim IsInBridge, IsFrezee,HdrState
 
 Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")



 if request("fccOde")="" then
    fcCode="" '----- Nantinya Ambil dari Format Master
 else
    fcCode = request("fccOde")
 end if
 
 fcba = Request("fcba")
If Request("fakturdate")<>"" then
	fakturdate   =Request("fakturdate")
End If

If Request("transdate")<>"" then
   transdate   =Request("transdate")
Else
   transdate   =textDate(date)
End if



 
 txtJenis = Request("txtJenis")
 xCode = Request("xCode")
 IsInBridge = Request("IsInBridge")
 Remarks = Request("Remarks")
 

 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID")
 use_ppn = Request("use_ppn")
 fccompanycode = Request("fccompanycode")

 invoice_by =Request("invoice_by")
 suppliercode = Request("suppliercode")
 selectedfcba = Request("selectedfcba")
 selectedsupplier = Request("selectedsupplier")
 selectedinvoice = Request("selectedinvoice")
 purchaseby = Request("purchaseby")
 suppliername = Request("suppliername")
 supplieraddress = Request("supplieraddress")
 termofpayment = Request("termofpayment")
 SHIPPINGPOINT = Request("SHIPPINGPOINT")
 dp = Request("dp")
 ppn = Request("ppn")
 pph = Request("pph")
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
 totalpph = Request("totalpph")
 fakturno = request ("fakturno")
 description = request ("description")
 lineno = request ("lineno")
 spolineno = request ("spolineno")
 invoice_type = request ("invoice_type")
 valid_from  = Request("valid_from")
 valid_until  = Request("valid_until")
 kg = request ("kg")
 rate = request ("rate") 


 set rsPurchase = server.CreateObject("Adodb.recordset")
	q= 	" select spo.invoiceno,spo.pocode, spo.suppliername, bu.fccompanycode, spo.fcba, spo.invoice_by, spo.valid_from, spo.valid_until," & _ 
		" spo.ppn, spo.pph, spo.totalppn, spo.totalpph, spo.grandtotal, spo.transdate, spo.no_faktur, spo.fakturdate, spo.pocode, spo.invoice_type," & _
		" spd.jobcode, spd.targetcode, spd.targettype, spd.description," & _
		" spd.kg, spd.rate, spd.amount, spd.description, spd.amount, spd.othercost, spd.totalamount, spd.lineno, spd.polinenokey" & _
		" from invoiceservicehdr spo " & _
		" LEFT join invoiceservicedtl spd " & _
		" on (spo.invoiceno = spd.invoiceno and spo.fcba = spd.fcba)" & _
		" INNER JOIN businessunit bu ON (bu.fccode = spo.fcba)" & _
		" where spo.invoiceno = '"& fcCode &"'" & _
		" and spo.fcba = '" & fcba & "'" & _ 
		" ORDER BY lineno"
	  
	  ' response.write q
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
	invoiceno = rsPurchase("invoiceno")
	transdate = rsPurchase("transdate")
	fcCode = rsPurchase("pocode")
	suppliercode = rsPurchase("suppliername")
	selectedsupplier = rsPurchase("suppliername")
	' totaldetamount = rsPurchase("totaldetamount")
	' totaldetcharge = rsPurchase("totaldetcharge")
	ppn = rsPurchase("ppn")
	pph = rsPurchase("pph")
	totalppn = rsPurchase("totalppn")
	totalpph = rsPurchase("totalpph")
	fcba  = rsPurchase("fcba")
	invoice_by  = rsPurchase("invoice_by")
	selectedinvoice  = rsPurchase("invoice_by")
	selectedfcba  = rsPurchase("fcba")
	fccompanycode  = rsPurchase("fccompanycode")
	grandtotal  = rsPurchase("grandtotal")
	fakturno  = rsPurchase("no_faktur")
	fakturdate  = rsPurchase("fakturdate")
	valid_from  = rsPurchase("valid_from")
	valid_until  = rsPurchase("valid_until")
	invoice_type  = rsPurchase("invoice_type")
	SudahClose="Y"
 End if

 if Request("edit")="Update" then

	  q= " Update invoiceservicehdr set transdate=to_Date('" & transdate & "','mm/dd/yyyy'),invoice_by='" & invoice_by & "', invoice_type='" & invoice_type & "', " & _ 
	  " suppliername='" & suppliercode & "',no_faktur='" & fakturno & "',fakturdate=to_Date('" & fakturdate & "','mm/dd/yyyy'), " & _ 
	  " valid_from=to_Date('" & valid_from & "','mm/dd/yyyy'),valid_until=to_Date('" & valid_until & "','mm/dd/yyyy'), " & _ 
	  " totalppn = replace('" & totalppn & "',',',''), totalpph = replace('" & totalpph & "',',',''), " & _
	  " grandtotal = replace('" & grandtotal & "',',','')" & _
		 ", LASTUPDATE = sysdate " & _
		 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
		 ", fcedit='" & session("CodeName") & "'" & _ 
		 " where  invoiceno='" & invoiceno & "' and fcBA='" & selectedfcba & "'"
		'response.write q
		connStr.execute(q)
		'Response.redirect("invoiceservice_edit.asp?fcCode="&invoiceno&"&fcba="&fcba&"&txtJenis=edit") 
		'response.redirect("paymentrequest.asp?fccode="&invoiceno&"&errno="&err.number&"&gandtotalPO="&gandtotalPO&"&DPPO="&DPPO)

 end if

 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="invoiceservice_edit.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Invoice No</td>
				<td align="left">
					<input type="text" maxlength="50" name="invoiceno" size="30" value="<%=invoiceno%>" autocomplete="off" readonly>
					<font color="#FF0000">*</font>
				</td>
				<td align="left">Amount Before Tax</td>
				<td align="left">
					<input type="text" maxlength="50" name="totaldetamount" size="30" value="<%=totaldetamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
				</td>
		</tr>
		<tr>
			<td align="left">Company</td>
				<td align="left">
					<select name="fccompanycode" value="<%=fccompanycode%>" id="fccompanycode"  class="chosen-select" style="width: 50%">
					<option value=""></option> 
						<% 	Dim rsComboC
							q = " SELECT DISTINCT fccompanycode " & _
							 	" FROM (SELECT businessunit.fccompanycode " & _
							    " FROM user_fcba_tab " & _
							    " INNER JOIN businessunit " & _
							    " ON user_fcba_tab.fcba = " & _
							    " businessunit.fccode " & _
							    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"')"
							' response.write q
							Set rsComboC = Server.CreateObject("ADODB.Recordset")
							rsComboC.CursorType = 3
							rsComboC.CursorLocation = 3
							rsComboC.Open q, connstr
							if not rsComboC.eof then%>
							<%
							Do While Not rsComboC.eof
							%>
								<option value="<%=rsComboC("fccompanycode")%>" 
									<%if cstr(""&fccompanycode)=cstr(""&rsComboC("fccompanycode")) then %> selected 
									<%End If%>><%=rsComboC("fccompanycode")%>
								</option> 
							<%rsComboC.movenext
							loop
							End if
							rsComboC.Close
							set rsComboC = Nothing
							q=""
						%>
				</select> 
					<font color="#FF0000">*</font>
				</td>
				<td align="left">Other Charge</td>
				<td align="left">
						<input type="text" maxlength="50" name="totaldetcharge" size="30" value="<%=totaldetcharge%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
				</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<div id="fcbalist"></div>
				<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
				<select name="fcba" value="<%=fcba%>" id="fcba"  class="chosen-select" style="width: 50%">
				
						<% 	Dim rsCombounit
							q = " SELECT DISTINCT fccode,fcname " & _
							    " FROM BUSINESSUNIT "
							' response.write q
							Set rsCombounit = Server.CreateObject("ADODB.Recordset")
							rsCombounit.CursorType = 3
							rsCombounit.CursorLocation = 3
							rsCombounit.Open q, connstr
							if not rsCombounit.eof then%>
							<%
							Do While Not rsCombounit.eof
							%>
								<option value="<%=rsCombounit("fccode")%>" 
									<%if cstr(""&fcba)=cstr(""&rsCombounit("fccode")) then %> selected 
									<%End If%>><%=rsCombounit("fccode")%>
								</option> 
							<%rsCombounit.movenext
							loop
							End if
							rsCombounit.Close
							set rsCombounit = Nothing
							
						%>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<td align="left">Ppn (%)</td>
			<td align="left">
				<input type="text" size="9" name="ppn" value="<%=ppn%>" onkeypress="return harusangka(event);">
				<input type="text" size="15" name="totalppn" value="<%=totalppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
			</td>
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="transdate" size="10" value="<%=transdate%>">
			</td>
			<td align="left">Pph (%)</td>
			<td align="left">
				<input type="text" size="9" name="pph" value="<%=pph%>" onkeypress="return harusangka(event);">
				<input type="text" size="15" name="totalpph" value="<%=totalpph%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
			</td>
		</tr>
		<tr>
			<td align="left">Invoice Type</td>	
			<td align="left">
				<select id="invoice_type" name="invoice_type" value="<%=invoice_type%>" onchange="javascript:getlistsupplier();" class="chosen-select" style="width: 40%">
					<option value="<%=invoice_type%>"></option>
					<option value="General" <% If cstr(""&invoice_type) = "General" Then %> Selected <%End if%>>General</option> 
					<option value="TBS" <% If cstr(""&invoice_type) = "TBS" Then %> Selected <%End if%>>TBS</option> 
				</select>
			</td>
			<td align="left">Grand Total</td>
				<td align="left">
						<input type="text" maxlength="50" name="grandtotal" size="30" value="<%=grandtotal%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By</td>
			<td align="left">
				<select name="invoice_by" value="<%=invoice_by%>" id="invoice_by"  class="chosen-select" style="width: 50%">
				
						<% 	Dim rsCombounitby
							q = " SELECT DISTINCT fccode,fcname " & _
							    " FROM BUSINESSUNIT "
							' response.write q
							Set rsCombounitby = Server.CreateObject("ADODB.Recordset")
							rsCombounitby.CursorType = 3
							rsCombounitby.CursorLocation = 3
							rsCombounitby.Open q, connstr
							if not rsCombounitby.eof then%>
							<%
							Do While Not rsCombounitby.eof
							%>
								<option value="<%=rsCombounitby("fccode")%>" 
									<%if cstr(""&selectedinvoice)=cstr(""&rsCombounitby("fccode")) then %> selected 
									<%End If%>><%=rsCombounitby("fccode")%>
								</option> 
							<%rsCombounitby.movenext
							loop
							End if
							rsCombounitby.Close
							set rsCombounitby = Nothing
							
						%>
				</select>
			</td>
				<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="<%=selectedinvoice%>">
			<td align="left">No Faktur</td>
			<td align="left">
				<input type="text" maxlength="50" name="fakturno"  size="30" value="<%=fakturno%>" autocomplete="off" >
			</td>	
		<tr>
			<td align="left">Supplier</td>
			<td align="left">
				<select name="suppliercode" value="<%=suppliercode%>" id="suppliercode"  class="chosen-select" style="width: 50%">
				
						<% 	Dim rsCombosup
							q = " SELECT DISTINCT fccode,fccode || ' - ' || fcname supname " & _
							    " FROM supplier where fcba='" & fcba & "' "
							' response.write q
							Set rsCombosup = Server.CreateObject("ADODB.Recordset")
							rsCombosup.CursorType = 3
							rsCombosup.CursorLocation = 3
							rsCombosup.Open q, connstr
							if not rsCombosup.eof then%>
							<%
							Do While Not rsCombosup.eof
							%>
								<option value="<%=rsCombosup("fccode")%>" 
									<%if cstr(""&suppliercode)=cstr(""&rsCombosup("fccode")) then %> selected 
									<%End If%>><%=rsCombosup("supname")%>
								</option> 
							<%rsCombosup.movenext
							loop
							End if
							rsCombosup.Close
							set rsCombosup = Nothing
							
						%>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<td align="left"> Faktur Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" id="fakturdate" name="fakturdate" size="10" value="<%=fakturdate%>">
			</td>
		</tr>
		<tr>
			<td align="left">Valid From</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" id="valid_from" name="valid_from" size="10" value="<%=valid_from%>">
			</td>
			<td align="left">Valid Until</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" id="valid_until" name="valid_until" size="10" value="<%=valid_until%>">
			</td>
		</tr>
		</tr>
		<%if businessunit = "KKP" OR businessunit = "KPK" then%>
		<tr>
		<td align="left">Ppn (10%)</td>
		<td align="left">
			<input type="checkbox" id="use_ppn" class="css-checkbox" name="use_ppn" value="<%=use_ppn%>" onclick='handleClick(this);' <%if use_ppn = "Y" then%> checked <%end if%>>
			<label for="use_ppn" class="css-label"></label>
		</td>
		</tr>
		<%end if%>
		
		<tr><td colspan="2"  align="left">

		<input Type="Submit"  name="edit" value="Update">&nbsp;&nbsp;
		
		</td>
		</tr>
    </table>  
 </form>

</div></div>
  <!--#include file="bottom.jsp"-->
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	
	// $(".chosen-select").empty();
	
	var config = {
	  '.chosen-select'           : {},
	  '.chosen-select-deselect'  : {allow_single_deselect:true},
	  '.chosen-select-no-single' : {disable_search_threshold:10},
	  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
	  '.chosen-select-width'     : {width:"95%"}
	}
	for (var selector in config) {
	  $(selector).chosen(config[selector]);
	}
	// $(".chosen-select").trigger("chosen:updated");

</script>

</body>
</html>
