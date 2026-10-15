<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Purchase Invoice Product</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
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
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<script>
	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}
</script>
<script>
	function hitungrumus() {
		var squantity = $("#quantity").val();
		var sunitprice = $("#unitprice").val();
		var newsquantity = squantity.replace(/,/g, '');
		var newsunitprice = sunitprice.replace(/,/g, '');
		
		//Untuk menghitung Total Quantity Tolerance
		totalamount_det = newsquantity * newsunitprice; //a kali b
		if (!isNaN(totalamount_det)) {
    	 	$("#amount").val(totalamount_det);
    	}
    	else{
    		$("#amount").val();
    	}		
	}
</script>
<script language="javascript" type="text/javascript">
/** XHConn - Simple XMLHTTP Interface - bfults@gmail.com - 2005-04-08        **
 ** Code licensed under Creative Commons Attribution-ShareAlike License      **
 ** http://creativecommons.org/licenses/by-sa/2.0/                           **/
function XHConn()
{
  var xmlhttp, bComplete = false;
  try { xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); }
  catch (e) { try { xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); }
  catch (e) { try { xmlhttp = new XMLHttpRequest(); }
  catch (e) { xmlhttp = false; }}}
  if (!xmlhttp) return null;
  this.connect = function(sURL, sMethod, sVars, fnDone)
  {
    if (!xmlhttp) return false;
    bComplete = false;
    sMethod = sMethod.toUpperCase();
    try {
      if (sMethod == "GET")
      {
        xmlhttp.open(sMethod, sURL+"?"+sVars, true);
        sVars = "";
      }
      else
      {
        xmlhttp.open(sMethod, sURL, true);
        xmlhttp.setRequestHeader("Method", "POST "+sURL+" HTTP/1.1");
        xmlhttp.setRequestHeader("Content-Type",
          "application/x-www-form-urlencoded");
      }
      xmlhttp.onreadystatechange = function(){
        if (xmlhttp.readyState == 4 && !bComplete)
        {
          bComplete = true;
          fnDone(xmlhttp);
        }};
      xmlhttp.send(sVars);
    }
    catch(z) { return false; }
    return true;
  };
  return this;
}

  	// doAJAXCall : Generic AJAX Handler, used with XHConn
// Author : Bryce Christensen (www.esonica.com)
// PageURL : the server side page we are calling
// ReqType : either POST or GET, typically POST
// PostStr : parameter passed in a query string format 'param1=foo&param2=bar'
// FunctionName : the JS function that will handle the response

var doAJAXCall = function (PageURL, ReqType, PostStr, FunctionName) {

	// create the new object for doing the XMLHTTP Request
	var myConn = new XHConn();

	// check if the browser supports it
	if (myConn)	{
	    
	    // XMLHTTPRequest is supported by the browser, continue with the request
	    myConn.connect('' + PageURL + '', '' + ReqType + '', '' + PostStr + '', FunctionName);    
	} 
	else {
	    // Not support by this browser, alert the user
	    alert("XMLHTTP not available. Try a newer/better browser, this application will not work!");   
	}
}

// launched from button click 
var getSalesContract = function () {
	// build up the post string when passing variables to the server side page
	var salescontractno = document.getElementById("salescontractno").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'salescontractno='+salescontractno+'&fcba='+selectedfcba;
	// use the generic function to make the request
	doAJAXCall('ajax/GetPurchaseAgreementInv.asp', 'POST', PostStr, showMessageResponse);
	
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // update the Div to show the result from the server
	document.getElementById("vat").value = s.vat;
}

var GetShipmentInv = function () {
	// build up the post string when passing variables to the server side page
	var selectedfcba = document.getElementById("selectedfcba").value;
	var salesshipmentno = document.getElementById("salesshipmentno").value;
	var PostStr = 'salesshipmentno='+salesshipmentno+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetShipmentInvPurchase.asp', 'POST', PostStr, ShowShipmentResponse);
}

// The function for handling the response from the server
var ShowShipmentResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    //alert (s.postingdate);
    
    // update the Div to show the result from the server
	document.getElementById("product").value = s.product;
	document.getElementById("unitprice").value = s.unitprice;
	document.getElementById("quantity").value = s.quantity;
}

// launched from button click 
var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbapurchaseinvoice.asp', 'POST', PostStr, showlistfcba);
}

// The function for handling the response from the server
var showlistfcba = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("fcbalist").innerHTML = "";
	document.getElementById("fcbalist").innerHTML = response;

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
}

var getlistinvoice = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedinvoice = document.getElementById("selectedinvoice").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&invoice_by='+selectedinvoice;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getinvoiceby.asp', 'POST', PostStr, showlistinvoice);
}

// The function for handling the response from the server
var showlistinvoice = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("invoicelist").innerHTML = "";
	document.getElementById("invoicelist").innerHTML = response;

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
}

var getlistcustomer = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedcustomercode = document.getElementById("selectedcustomercode").value;
	var PostStr = 'fcba='+selectedfcba+'&customercode='+selectedcustomercode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistsuppliercodesalesbasedonfcba.asp', 'POST', PostStr, showlistcustomer);
}

// The function for handling the response from the server
var showlistcustomer = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("customerlist").innerHTML = "";
	document.getElementById("customerlist").innerHTML = response;
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
}

var getlistinvoicedp = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedinvoicedpno = document.getElementById("selectedinvoicedpno").value;
	var PostStr = 'fcba='+selectedfcba+'&invoicedpno='+selectedinvoicedpno;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistinvoicedppurchase.asp', 'POST', PostStr, showlistinvoicedp);
}

var getlistlinenosales = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedinvoicedpno = document.getElementById("selectedinvoicedpno").value;
	var selectedlinenosales = document.getElementById("selectedlinenosales").value;
	var PostStr = 'fcba='+selectedfcba+'&invoicedpno='+selectedinvoicedpno+'&linenoinvoicedp='+selectedlinenosales;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistlinenopurchaseinvoice.asp', 'POST', PostStr, showlistlinenosales);
}

// The function for handling the response from the server
var showlistlinenosales = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("linenosalesinvoice").innerHTML = "";
	document.getElementById("linenosalesinvoice").innerHTML = response;

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
}

// The function for handling the response from the server
var showlistinvoicedp = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("invoicedplist").innerHTML = "";
	document.getElementById("invoicedplist").innerHTML = response;
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
}

var getlistsalescontract = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedcustomercode = document.getElementById("selectedcustomercode").value;
	var selectedsalescontractno = document.getElementById("selectedsalescontractno").value;
	var PostStr = 'fcba='+selectedfcba+'&customercode='+selectedcustomercode+'&salescontractno='+selectedsalescontractno;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistpurchasecontract.asp', 'POST', PostStr, showlistsalescontract);
}

// The function for handling the response from the server
var showlistsalescontract = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("salescontractlist").innerHTML = "";
	document.getElementById("salescontractlist").innerHTML = response;
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
}

var Getdetinvoicedp = function () {
	// build up the post string when passing variables to the server side page
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedinvoicedpno = document.getElementById("selectedinvoicedpno").value;
	var selectedlinenosales = document.getElementById("selectedlinenosales").value;
	var PostStr = 'invoicedpno='+selectedinvoicedpno+'&fcba='+selectedfcba+'&linenoinvoicedp='+selectedlinenosales;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getdetinvoicedppurchase.asp', 'POST', PostStr, Showdetinvoicedp);
}

// The function for handling the response from the server
var Showdetinvoicedp = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    //alert (s.dp_aftervat);
    
    // update the Div to show the result from the server
	document.getElementById("dp_aftervat").value = s.dp_aftervat; 
}

</script>
<script language="JavaScript">					
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>

<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
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
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal!------------->
 <%

Dim QUANTITY_,RATE_,AMOUNT_

invoicecode=request("invoicecode")
salescontractno =request("salescontractno")
transdate =request("transdate")
customercode =request("customercode")
vat =request("vat")
netbalance=request("netbalance")
remarks=request("remarks")
deliveryorderno=request("deliveryorderno")
txtlineno = request("txtlineno")
productcode =request("productcode")
unitprice =request("unitprice")
deliveryquantity =request("deliveryquantity")
quantity =request("quantity")
grandtotal  =request("grandtotal")
salesshipmentno  =request("salesshipmentno")
amount  =request("amount")
amount_beforevat  =request("amount_beforevat")
amountvat  =request("amountvat")
dp_aftervat  =request("dp_aftervat")
state  =request("state")
invoice_by = Request("invoice_by")
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
selectedcustomercode = Request("selectedcustomercode")
selectedsalescontractno = Request("selectedsalescontractno")
selectedinvoice = Request("selectedinvoice")
selectedinvoicedpno = Request("selectedinvoicedpno")
selectedlinenosales= Request("selectedlinenosales")
linenoinvoicedp = Request("linenoinvoicedp")
valid_from = request("valid_from")
valid_until = request("valid_until")
expire_date = request("expire_date")
fakturdate = request("fakturdate")
no_faktur = request("no_faktur")
invoicedpno =request("invoicedpno")

Dim SudahClose 
SudahClose = "N"


Function CheckTicket(initchit, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total from deliverynote" & _
					" where wbt_chitno = '"& initchit &"'" & _
					" and fcba = '"& initialFcba & "'" 
	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable =  Cdbl(to_nol(rsAvailable("total")))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckTicket = initialAvailable
	
End Function

Function CheckQtyDo(initdeliveryorderno, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total from deliveryorder" & _
					" where deliveryorderno = '"& initdeliveryorderno &"'" & _
					" and fcba = '"& initialFcba & "'" 
	response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable =  Cdbl(to_nol(rsAvailable("total")))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckQtyDo = initialAvailable
	
End Function



If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, transdate)
		If Cek_IsDailyClosing = "FALSE" then

			strSuppyingDeclaration	=	" SELECT count(1) totaldata FROM purchaseinvoicedtl" & _
										" WHERE invoicecode = '" & invoicecode & "' and fcba = '"& fcba &"'"  
			' response.Write strSuppyingDeclaration
			set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
			rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
			if Not rsSupplyingDeclaration.Eof then
				while Not rsSupplyingDeclaration.Eof
					totaldata =  cdbl(to_nol(rsSupplyingDeclaration("totaldata")))
					rsSupplyingDeclaration.movenext
				wend
			End if
			rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing
			
				if totaldata > 0 then
					err.number = 1001
					err.description = "Informasi purchase Invoice """ & invoicecode & """ masih ada, mohon hapus purchase Invoice Detail terlebih dahulu ! "
				else
					q = " delete from purchaseinvoicehdr " & _
						" where fcBA = '" & fcba & "'" & _
						" and invoicecode = '" & invoicecode & "'"
						' Response.Write q
						connStr.execute(q)
				end if 
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then
			
			q = " delete from purchaseinvoicedtl " & _
				" where fcBA = '" & fcba & "'" & _
				" and invoicecode = '"& invoicecode & "'" & _
				" and purchaseshipmentno = '"& salesshipmentno &"'"
				'Response.Write q
				connStr.execute(q)
			
			' Untuk Menghitung Amount Before Vat dan Amount After Vat By SAG 25 Juni 2019
				amount = 0
				strSuppyingDeclaration	=	" SELECT NVL(amount,0) amount FROM purchaseinvoicedtl" & _
											" WHERE invoicecode = '" & invoicecode & "' and fcba = '"& fcba &"'"  
				' response.Write strSuppyingDeclaration
				set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
				rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
				if Not rsSupplyingDeclaration.Eof then
					while Not rsSupplyingDeclaration.Eof
						amount = amount + cdbl(to_nol(rsSupplyingDeclaration("amount")))
						rsSupplyingDeclaration.movenext
					wend
				End if
				rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing

				strSuppyingDeclaration	=	" SELECT NVL(vat,0) vat FROM purchaseinvoicehdr" & _
											" WHERE invoicecode = '" & invoicecode & "' and fcba = '"& fcba &"'"  
				' response.Write strSuppyingDeclaration
				set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
				rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
				if Not rsSupplyingDeclaration.Eof then
					while Not rsSupplyingDeclaration.Eof
						vat = cdbl(to_nol(rsSupplyingDeclaration("vat")))
						rsSupplyingDeclaration.movenext
					wend
				End if
				rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing

				amount_beforevat = amount 
				amountvat = amount * cdbl(to_nol(vat))/100
				grandtotal = amount_beforevat + amountvat

				q = " Update purchaseinvoicehdr  set amount_beforevat = REPLACE('" & amount_beforevat & "',',','')"  & _
					", amountvat = REPLACE('" & amountvat & "',',','')" & _ 
					", grandtotal = REPLACE('" & grandtotal & "',',','')" & _ 
					", LASTUPDATE = sysdate " & _
					", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					" where invoicecode='" & invoicecode & "'" & _
					" and fcba='" & fcba & "'"
				' response.Write(q)
				connStr.execute(q)
			' End Added
			
			cmdView	= "True"

	elseif Request("txtJenis")="New" then
		Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
		rsSupplyingDeclarationHdr.Open " select * from purchaseinvoicehdr where invoicecode = '" & invoicecode & "' and fcba ='" & fcba & "' ", connstr
		If rsSupplyingDeclarationHdr.EOF Then
		
			q = " INSERT into purchaseinvoicehdr" & _
				" (invoicecode, invoice_by, linenoinvoicedp, invoicedpno, transdate, valid_from, valid_until, expire_date, no_faktur, fakturdate, suppliercode, vat, purchasecontractno, remarks, fcentry, fcedit, fcip, fcba, lastupdate, lasttime,suppliercontractno) values (" & _
				" '" & invoicecode & "'" & _
				" , '" & invoice_by & "'" & _
				" , '" & linenoinvoicedp & "'" & _
				" , '" & invoicedpno & "'" & _
				" , to_date('" & transdate & "','mm/dd/yyyy')" & _ 
				" , to_date('" & valid_from & "','mm/dd/yyyy')" & _ 
				" , to_date('" & valid_until & "','mm/dd/yyyy')" & _ 
				" , to_date('" & expire_date & "','mm/dd/yyyy')" & _ 
				" , '" & no_faktur & "'" & _ 
				" , to_date('" & fakturdate & "','mm/dd/yyyy')" & _ 
				" , '" & customercode & "'" & _ 
				" , replace('" & vat & "',',','')" & _
				" , '" & salescontractno & "'" & _
				" , '" & remarks & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				" , '" & fcba & "'" & _
				" , sysdate" & _
				" , to_char(sysdate, 'hh24:mi')" & _
				" , '" & salescontractno & "'" & _
				" )"								
			' Response.Write q
			connStr.execute(q)
			
		End If
	' rsSupplyingDeclarationHdr.Close
	' set rsSupplyingDeclarationHdr = Nothing
	
			' q = " insert into " & _
			' 	" (invoicecode, salesshipmentno, lineno, productcode, quantity, unitprice, amount, state, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
			' 	" '" & invoicecode & "'" & _
			' 	" , '" & salesshipmentno & "'" & _
			' 	" , (SELECT NVL(MAX(lineno),0)  FROM salesinvoicedtl WHERE fcba = '" & fcba & "') +1" & _
			' 	" , '" & productcode & "'" & _
			' 	" , '" & quantity & "'" & _
			' 	" , replace('" & unitprice & "',',','')" & _
			' 	" , replace('" & amount & "',',','')" & _
			' 	" , 'Invoiced'" & _
			' 	" , '" & Session("CodeName") & "'" & _
			' 	" , '" & Session("CodeName") & "'" & _
			' 	" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			' 	" , '" & fcba & "'" & _
			' 	" , sysdate" & _
			' 	" , to_char(sysdate, 'hh24:mi')" & _
			' 	" )"
			' 	' Response.Write q
			' 	connStr.execute(q)
			
			' ' Untuk Menghitung Amount Before Vat dan Amount After Vat By SAG 25 Juni 2019
			' 	strAvailable =	" select sum(amount) amount from salesinvoicedtl" & _
			' 					" where invoicecode = '"& invoicecode &"'" & _
			' 					" and fcba = '"& fcba & "'" 
			' 	' response.write stravailable
			' 	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
			' 	rsAvailable.Open strAvailable, connstr, 3, 3
				
			' 	If Not rsAvailable.EOF Then
			' 		amount =  Cdbl(to_nol(rsAvailable("amount")))
			' 	End If
			' 	rsAvailable.Close
			' 	set rsAvailable = Nothing
				
			' 	amount_beforevat = amount
			' 	amountvat = amount * vat / 100
			' 	grandtotal = amount_beforevat + amountvat 

			' 	q =	" Update salesinvoicehdr" & _
			' 		" set 	amountvat = replace('" & amountvat & "',',','')" & _
			' 		" 	  , amount_beforevat = replace('" & amount & "',',','')" & _
			' 		" 	  , grandtotal = replace('" & grandtotal & "',',','')" & _
			' 		"     , fcentry = '" & Session("CodeName") & "'" & _
			' 		"     , fcedit = '" & Session("CodeName") & "'" & _
			' 		" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			' 		" 		where invoicecode = '" & invoicecode & "'" & _
			' 		"       and fcBA = '" & fcba & "'"
	 	' 		' response.write q
			' 	connStr.execute(q)
			' End Added
			if cdbl(to_nol(ratemarket)) = 0 then
			For n=1 to (cdbl(request("no"))-1) 
				if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				            QUANTITY_ = cdbl(to_nol(request("acceptquantity_"&TRIM(n))))
							RATE_     = cdbl(to_nol(request("salesprices_"&TRIM(n))))
							AMOUNT_   = QUANTITY_ * RATE_														
				
						q = " Insert into  purchaseinvoicedtl (invoicecode, postingdate, purchaseshipmentno, productcode, quantity, unitprice, amount, lineno, state, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
							" values ('" & invoicecode & "'" & _
							" ,to_date('" & request("postingdate_"&TRIM(n)) & "','mm/dd/yyyy')" & _
							" ,'" & request("salesshipmentno_"&TRIM(n)) & "'" & _
							" ,'" & request("productcode_"&TRIM(n)) & "'" & _
							" , REPLACE('" & QUANTITY_ & "',',', '')" & _
							" , REPLACE('" & RATE_ & "',',', '')" & _
							" , REPLACE('" & AMOUNT_ & "',',', '')" & _
							" ,(SELECT NVL(MAX(lineno),0) FROM purchaseinvoicedtl WHERE invoicecode = '" & invoicecode & "' AND fcba = '"& fcba &"') +1" & _
							" ,'INVOICED'" & _
							" ,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & request("fcba_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						'response.write q
						connStr.execute(q)
						
					' Untuk Menghitung Amount Before Vat dan Amount After Vat By SAG 25 Juni 2019
						amount = 0
						strSuppyingDeclaration	=	" SELECT NVL(amount,0) amount FROM purchaseinvoicedtl" & _
													" WHERE invoicecode = '" & invoicecode & "' and fcba = '"& fcba &"'" 
						' response.Write strSuppyingDeclaration
						set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
						rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
						if Not rsSupplyingDeclaration.Eof then
							while Not rsSupplyingDeclaration.Eof
								amount = amount + cdbl(to_nol(rsSupplyingDeclaration("amount")))
								rsSupplyingDeclaration.movenext
							wend
						End if
						rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing
						amount_beforevat = amount 
						amountvat = amount * cdbl(to_nol(vat))/100
						grandtotal = amount_beforevat + amountvat

						q = " Update purchaseinvoicehdr  set amount_beforevat = REPLACE('" & amount_beforevat & "',',','')"  & _
							", amountvat = REPLACE('" & amountvat & "',',','')" & _ 
							", grandtotal = REPLACE('" & grandtotal & "',',','')" & _ 
							", LASTUPDATE = sysdate " & _
							", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							" where invoicecode='" & invoicecode & "'" & _
							" and fcba='" & fcba & "'"
						' response.Write(q)
						connStr.execute(q)
					' End Added
				End if	 
			Next
		Else
			For n=1 to (cdbl(request("no"))-1) 
				if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				            QUANTITY_ = cdbl(to_nol(request("acceptquantity_"&TRIM(n))))
							RATE_     = cdbl(to_nol(request("salesprices_"&TRIM(n))))
							AMOUNT_   = QUANTITY_ * RATE_
							
						q = " Insert into  purchaseinvoicedtl ( invoicecode, postingdate, purchaseshipmentno, productcode, quantity, unitprice, amount, lineno, state, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
							" values ( '" & invoicecode & "'" & _
							" ,to_date('" & request("postingdate_"&TRIM(n)) & "','mm/dd/yyyy')" & _
							" ,'" & request("salesshipmentno_"&TRIM(n)) & "'" & _
							" ,'" & request("productcode_"&TRIM(n)) & "'" & _
							" , REPLACE('" & QUANTITY_ & "',',', '')" & _
							" , REPLACE('" & RATE_ & "',',', '')" & _
							" , REPLACE('" & AMOUNT_ & "',',', '')" & _
							" ,(SELECT NVL(MAX(lineno),0) FROM purchaseinvoicedtl WHERE invoicecode = '" & invoicecode & "' AND fcba = '"& fcba &"') +1" & _
							" ,'INVOICED'" & _
							" ,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & request("fcba_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						' response.write q
						connStr.execute(q)
						
						' Untuk Menghitung Amount Before Vat dan Amount After Vat By SAG 25 Juni 2019
						amount = 0
						strSuppyingDeclaration	=	" SELECT NVL(amount,0) amount FROM purchaseinvoicedtl" & _
													" WHERE invoicecode = '" & invoicecode & "' and fcba = '"& fcba &"'"  
						' response.Write strSuppyingDeclaration
						set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
						rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
						if Not rsSupplyingDeclaration.Eof then
							while Not rsSupplyingDeclaration.Eof
								amount = amount + cdbl(to_nol(rsSupplyingDeclaration("amount")))
								rsSupplyingDeclaration.movenext
							wend
						End if
						rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing
						amount_beforevat = amount 
						amountvat = amount * cdbl(to_nol(vat))/100
						grandtotal = amount_beforevat + amountvat

						q = " Update purchaseinvoicehdr  set amount_beforevat = REPLACE('" & amount_beforevat & "',',','')"  & _
							", amountvat = REPLACE('" & amountvat & "',',','')" & _ 
							", grandtotal = REPLACE('" & grandtotal & "',',','')" & _ 
							", LASTUPDATE = sysdate " & _
							", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							" where invoicecode='" & invoicecode & "'" & _
							" and fcba='" & fcba & "'"
						' response.Write(q)
						connStr.execute(q)
					' End Added
				End if	 
			Next
		End If 
' End If
cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			' If Len(Trim(chitno)) = 0 Then
			' 	Err.Number = 1001
			' 	Err.Description = "No Tikcet Harus diisi, tidak boleh kosong"
			' Else
			' Dim val_Tickets : val_Tickets = 0
			' val_Tickets = CheckTicket(wbt_chitno, session("fcBA"))
			' If val_Tickets = 0 Then
			' 	Err.Number = 1001
			' 	Err.Description = "No Ticket """ & wbt_chitno & """ tidak ada, Koreksi kembali No Ticket yang anda masukkan."
			' Else
					q =	" Update purchaseinvoicedtl" & _
						" set purchaseshipmentno = '" & salesshipmentno & "'" & _
						"     , productcode = '" & productcode & "'" & _
						"     , quantity = '" & quantity & "'" & _
						" 	  , unitprice = replace('" & unitprice & "',',','')" & _
						" 	  , amount = replace('" & amount & "',',','')" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where invoicecode = '" & invoicecode & "'" & _
						"       and fcBA = '" & fcba & "'"
		 			' response.write q
					connStr.execute(q)
					End If
				End If		
		' 	End If	
		' End If		
	

cmdView	= "True"	
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

'untuk menampilkan datepicker default
If Len(Trim(transdate)) = 0 Then	transdate = textDate(date)
If Len(Trim(valid_from)) = 0 Then	valid_from = textDate(date)
If Len(Trim(valid_until)) = 0 Then	valid_until = textDate(date)
If Len(Trim(expire_date)) = 0 Then	expire_date = textDate(date) 

Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select hdr.*, bu.fccompanycode,dp.dp_aftervat,(SELECT Count(invoicecode) cnt FROM purchaseinvoicedtl WHERE invoicecode=hdr.invoicecode) cnt " & _
							" from purchaseinvoicehdr hdr" & _
							" left join businessunit bu on bu.fccode = hdr.fcba" & _
							" left join purchaseinvoicedp dp on hdr.invoicedpno=dp.invoicedpno and hdr.fcba=dp.fcba" & _ 
							" where hdr.invoicecode = '" & invoicecode & "'" 
' response.write strSuppyingDeclaration							
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	Do While Not rsSupplyingDeclaration.eof
		transdate = rsSupplyingDeclaration("transdate")
		valid_from = rsSupplyingDeclaration("valid_from")
		valid_until = rsSupplyingDeclaration("valid_until")
		expire_date = rsSupplyingDeclaration("expire_date")
		fakturdate = rsSupplyingDeclaration("fakturdate")
		no_faktur = rsSupplyingDeclaration("no_faktur")
		remarks = rsSupplyingDeclaration("remarks")	
		customercode = rsSupplyingDeclaration("suppliercode")
		salescontractno = rsSupplyingDeclaration("purchasecontractno")
		vat = rsSupplyingDeclaration("vat")	
		grandtotal = rsSupplyingDeclaration("grandtotal")
		amount_beforevat = rsSupplyingDeclaration("amount_beforevat")
		amountvat = rsSupplyingDeclaration("amountvat")
		invoice_by = rsSupplyingDeclaration("invoice_by")
		selectedinvoice = rsSupplyingDeclaration("invoice_by")
		fccompanycode = rsSupplyingDeclaration("fccompanycode")
		fcba = rsSupplyingDeclaration("fcba")
		selectedfcba = rsSupplyingDeclaration("fcba")
		selectedcustomercode = rsSupplyingDeclaration("suppliercode")
		selectedsalescontractno = rsSupplyingDeclaration("purchasecontractno")
		selectedinvoicedpno = rsSupplyingDeclaration("invoicedpno")
		invoicedpno = rsSupplyingDeclaration("invoicedpno")
		linenoinvoicedp = rsSupplyingDeclaration("linenoinvoicedp")
		selectedlinenosales = rsSupplyingDeclaration("linenoinvoicedp")
		' state =  rsSupplyingDeclaration("state")
		dp_aftervat = rsSupplyingDeclaration("dp_aftervat") 
		if rsSupplyingDeclaration("cnt") <> "0" then
		SudahClose = "Y"
		else
		SudahClose = "N"
		end if
		
		rsSupplyingDeclaration.movenext
	loop
else
	amount_beforevat = 0
	amountvat = 0
	grandtotal = 0
End if
'response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('costbook')"
			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if

'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing
%>

 <!-------------ISI Supplying Declaration Form!------------->
 
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Purchase Invoice</div><% end if %>
	<form name="frmEntry" method="post" action="purchaseinvoice.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left">Invoice No.</td>
			<td align="left">
			<input type="text" maxlength="50" name="invoicecode" style="background:#FFFFCC;" size="30" value="<%=invoicecode%>" autocomplete="Off">
			<font color="#FF0000">*</font>
			</td>
			<td align="left" width="20%">purchase Invoice DP No</td>
			<td align="left">
				<div id="invoicedplist"></div>
				<input type="Hidden" id="selectedinvoicedpno" name="selectedinvoicedpno" value="<%=selectedinvoicedpno%>">
			</td>
		</tr>
		<tr>
			<td align="left">Creation Date </td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="transdate" value="<%=transdate%>">
			</td>
			<td align="left" width="20%">Line No purchase Invoice DP</td>
			<td align="left">
				<div id="linenosalesinvoice"></div>
				<input type="Hidden" id="selectedlinenosales" name="selectedlinenosales" value="<%=selectedlinenosales%>">
			</td>
		</tr>
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba(); javascript:getlistinvoice();" class="chosen-select" style="width: 50%">
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
			<td align="left" width="20%">DP After VAT</td>
			<td align="left">
				<input type="text" id="dp_aftervat" name="dp_aftervat" value="<%=dp_aftervat%>" readOnly style="background-color:#eee; text-align:right;" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
			</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<div id="fcbalist"></div>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			<td align="left">Valid From</td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="valid_from" value="<%=valid_from%>">
			</td> 
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By
				<td align="left">
					<div id="invoicelist"></div>
				</td>
					<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="<%=selectedinvoice%>">
			</td>
			<td align="left">Valid Until</td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="valid_until" value="<%=valid_until%>">
			</td> 	
		</tr>
		<tr>
			<td align="left">Supplier Code</td>
			<td align="left">
				<div id="customerlist"></div>
				<input type="Hidden" id="selectedcustomercode" name="selectedcustomercode" value="<%=selectedcustomercode%>">
			</td>
			<td align="left">Amount Before VAT</td>
			<td align="left">
				<input type="text" size="29" style="background:#EEE;" name="amount_beforevat" value="<%=amount_beforevat%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>		
		</tr>
		<tr>
			<td align="left">purchase Contract No.</td>
			<td align="left" >
				<div id="salescontractlist"></div>
				<input type="Hidden" id="selectedsalescontractno" name="selectedsalescontractno" value="<%=selectedsalescontractno%>">
			</td>
			<td align="left">VAT</td>
			<td align="left">
				<input type="text" size="5"  id="vat" name="vat" value="<%=vat%>" >
				<font color="#FF0000">%</font>
				<input type="text" size="15" style="background:#EEE;" id="amountvat" name="amountvat" value="<%=amountvat%>" readonly  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"> 
			</td>
		</tr> 
		<tr>
			<td align="left">No. Faktur</td>
			<td align="left" >
				<input type="text" size="30" id="no_faktur" name="no_faktur" value="<%=no_faktur%>">
			</td>
			<td align="left">Grand Total</td>
			<td align="left">
				<input type="text" size="29" style="background:#EEE;" name="grandtotal" value="<%=grandtotal%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Remarks</td>
			<td align="left">
			<textarea name="remarks" cols="40"><%=remarks%></textarea>
			</td>
			<td align="left">Expired Date</td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="expire_date" value="<%=expire_date%>">
			</td>
		</tr>
		<tr>
			<td align="left">Faktur Date </td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="fakturdate" value="<%=fakturdate%>">
			</td>
			
		</tr>
		<tr>
			<td colspan="4"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				
					<a href="purchaseinvoice.asp?invoicecode=<%=invoicecode%>&salescontractno=<%=salescontractno%>&fcba=<%=fcba%>&txtJenis=HeaderDelete">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0" onClick="return confirm('Are you sure to delete');">
					</a>
				
				<!-- HADIR -->
				<!-- <%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Nursery"
				dataSubmodule 	= "Culling"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a> -->
				<!-- END:HADIR -->
			</td>
		</tr>
		</table>
	</form>
<%if cmdView = "True" and len(trim(invoicecode)) > 0  Then

	Dim rsVH, intRecordsPerPage, cColor
	if SudahClose="Y" Then
	'----------- Query get Shipment Invoice lama
	q	=	"  SELECT sidt.fcba," & _
			"         sidt.purchaseshipmentno," & _
			"         sah.purchasecontractno," & _
			"         sidt.postingdate," & _
			"         SUM (sidt.quantity) acceptquantity," & _
			"         sah.suppliercode," & _
			"         sad.productcode," & _
			"         sidt.unitprice salesprices," & _
			"         pd.productdescription," & _
			"         sidt.state " & _
			"    FROM purchaseinvoicedtl sidt " & _
			" 	 INNER JOIN purchaseinvoicehdr sihr ON (sidt.invoicecode=sihr.invoicecode AND sidt.fcba=sihr.fcba) " & _
			" 	 LEFT JOIN purchaseagreementhdr sah ON ( sah.purchasecontractno = sihr.purchasecontractno AND sah.fcba = sihr.fcba) " & _
			" 	 LEFT JOIN purchaseagreementdet sad ON ( sah.purchasecontractno = sad.purchasecontractno AND sah.fcba = sad.fcba) " & _
			" 	 LEFT JOIN product pd ON (pd.productcode = sad.productcode AND pd.fcba = sad.fcba) " & _
			"    WHERE     sah.purchasecontractno = '"& salescontractno &"'" & _
			"		  AND sidt.invoicecode = '"& invoicecode &"'" & _
			"         AND sidt.fcba = '"& fcba &"'" & _
			"        /* AND sidt.postingdate >= TO_DATE ('"& valid_from &"', 'mm/dd/yyyy') */" & _
			"       /*  AND sidt.postingdate <= TO_DATE ('"& valid_until &"', 'mm/dd/yyyy') */" & _
			"		  AND sidt.state is not null " & _
			"	GROUP BY " & _
			"   sidt.fcba, " & _
			"	sidt.purchaseshipmentno, " & _
			"	sah.purchasecontractno, " & _
			"	sidt.postingdate, " & _
			"	sah.suppliercode, " & _
			"	sad.productcode, " & _
			"	sidt.unitprice, " & _
			"	pd.productdescription, " & _
			"	sidt.state " & _
			" Order BY sidt.postingdate, sidt.fcba "
	
	' q	=	"  SELECT shd.fcba," & _
			' "         shd.salesshipmentno," & _
			' "         sah.salescontractno," & _
			' "         ssd.postingdate," & _
			' "         SUM (ssd.acceptquantity) acceptquantity," & _
			' "         sah.customercode," & _
			' "         sad.productcode," & _
			' "         sad.salesprices," & _
			' "         pd.productdescription," & _
			' "         sidt.state" & _
			' "    FROM salesshipmenthdr shd" & _
			' "         LEFT JOIN salesshipmentdtl ssd" & _
			' "            ON (    shd.salesshipmentno = ssd.salesshipmentno" & _
			' "                AND shd.fcba = ssd.fcba)" & _
			' "         LEFT JOIN salesagreementhdr sah" & _
			' "            ON (    sah.salescontractno = shd.salescontractno" & _
			' "                AND sah.fcba = shd.fcba)" & _
			' "         LEFT JOIN salesagreementdet sad" & _
			' "            ON (    sah.salescontractno = sad.salescontractno" & _
			' "                AND sah.fcba = sad.fcba)" & _
			' "         LEFT JOIN product pd" & _
			' "            ON (pd.productcode = sad.productcode AND pd.fcba = sad.fcba)" & _
			' "         LEFT JOIN salesinvoicedtl sidt" & _
			' "            ON (    sidt.salesshipmentno = ssd.salesshipmentno" & _
			' "                AND sidt.fcba = ssd.fcba)" & _
			' "   WHERE     shd.salescontractno = '"& salescontractno &"'" & _
			' "		  AND sidt.invoicecode = '"& invoicecode &"'" & _
			' "         AND ssd.fcba = '"& fcba &"'" & _
			' "         AND ssd.postingdate >= TO_DATE ('"& valid_from &"', 'mm/dd/yyyy')" & _
			' "         AND ssd.postingdate <= TO_DATE ('"& valid_until &"', 'mm/dd/yyyy')" & _
			' "		  AND sidt.state is not null" & _
			' " GROUP BY shd.fcba," & _
			' "         shd.salesshipmentno," & _
			' "         sah.salescontractno," & _
			' "         ssd.postingdate," & _
			' "         sah.customercode," & _
			' "         sad.productcode," & _
			' "         sad.salesprices," & _
			' "         pd.productdescription," & _
			' "         sidt.state" & _ 
			' " Order BY ssd.postingdate, shd.fcba"

	'response.Write ("muncul 1:" & q & "<br>")
	set rsVH = server.CreateObject("Adodb.recordset")
	rsVH.Open q, connstr

	if Not rsVH.eof then
		while not rsVH.Eof
		salescontractno = rsVH("purchasecontractno")
		quantity = rsVH("acceptquantity")
		postingdate = rsVH("postingdate")
		productcode = rsVH("productcode")
		customercode = rsVH("suppliercode")
		salesprices = rsVH("salesprices")
		productdescription = ("productdescription")
		fcBA = rsVH("fcBA")
		state = rsVH("state")
		rsVH.movenext
		wend
	end if
	rsVH.Close : Set rsVH = Nothing

	strAvailable =	"  SELECT mpp.fcba," & _
					"         mpp.purchaseshipmentno," & _
					"         mpp.postingdate," & _
					"         mp.productcode," & _
					"         mpp.productdescription," & _
					"         mpp.suppliercode," & _
					"         SUM (mpp.acceptquantity) acceptquantity," & _
					"         NVL (mp.rate, 0) salesprices," & _
					"         mpp.state" & _
					"    FROM marketprice mp" & _
					"         INNER JOIN" & _
					"         (  SELECT MAX (valid_from) valid_from," & _
					"                   sad.productcode," & _
					"                   ssd.acceptquantity," & _
					"                   pd.productdescription," & _
					"                   sah.suppliercode," & _
					"                   ssh.purchaseshipmentno," & _
					"                   ssh.fcba," & _
					"                   ssd.postingdate," & _
					"                   sidt.state" & _
					"              FROM marketprice mp" & _
					"                   LEFT JOIN purchaseagreementdet sad" & _
					"                      ON (sad.productcode = mp.productcode AND sad.fcba = mp.fcba)" & _
					"                   LEFT JOIN purchaseshipmenthdr ssh" & _
					"                      ON (    ssh.purchasecontractno = sad.purchasecontractno" & _
					"                          AND ssh.fcba = sad.fcba)" & _
					"                   LEFT JOIN purchaseagreementhdr sah" & _
					"                      ON (    sah.purchasecontractno = ssh.purchasecontractno" & _
					"                          AND sah.fcba = ssh.fcba)" & _
					"                   LEFT JOIN purchaseshipmentdtl ssd" & _
					"                      ON (    ssd.purchaseshipmentno = ssh.purchaseshipmentno" & _
					"                          AND ssd.fcba = ssh.fcba)" & _
					"                   LEFT JOIN product pd" & _
					"                      ON (pd.productcode = sad.productcode AND pd.fcba = sad.fcba)" & _
					"                   LEFT JOIN purchaseinvoicedtl sidt" & _
					"                      ON (    sidt.purchaseshipmentno = ssd.purchaseshipmentno" & _
					"                          AND sidt.fcba = ssd.fcba)" & _
					"             WHERE     mp.productcode = '"& productcode &"'" & _
					"                   AND valid_from <= ssd.postingdate" & _
					"                   AND sah.suppliercode = '"& customercode &"'" & _
					"                   AND mp.fcba = '"& fcba &"'" & _
					"                   AND sah.purchasecontractno = '"& salescontractno &"'" & _
					"                   AND ssd.postingdate >= TO_DATE ('"& valid_from &"', 'mm/dd/yyyy')" & _
					"                   AND ssd.postingdate <= TO_DATE ('"& valid_until &"', 'mm/dd/yyyy')" & _
					"					AND sidt.state is not null" & _
					"          GROUP BY sad.productcode," & _
					"                   ssd.acceptquantity," & _
					"                   pd.productdescription," & _
					"                   sah.suppliercode," & _
					"                   ssh.purchaseshipmentno," & _
					"                   ssd.postingdate," & _
					"                   ssh.fcba," & _
					"                   sidt.state) mpp" & _
					"            ON (    mpp.valid_from = mp.valid_from" & _
					"                AND mp.productcode = '"& productcode &"'" & _
					"                AND mp.customercode = '"& customercode &"'" & _
					"                AND mp.fcba = mpp.fcba)" & _
					" GROUP BY mpp.fcba," & _
					"         mpp.purchaseshipmentno," & _
					"         mpp.postingdate," & _
					"         mp.productcode," & _
					"         mpp.productdescription," & _
					"         mpp.suppliercode," & _
					"         NVL (mp.rate, 0)," & _
					"         mpp.state" & _
					" ORDER BY mpp.postingdate, mpp.fcba"

	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr
	
	If Not rsAvailable.EOF Then
		while Not rsAvailable.Eof
			salesshipmentnomarket = rsAvailable("purchaseshipmentno")
			ratemarket = Cdbl(to_nol(rsAvailable("salesprices")))
			qtymarket =  Cdbl(to_nol(rsAvailable("acceptquantity")))
			customermarket = rsAvailable("suppliercode")
			postingdatemarket = rsAvailable("postingdate")
			productcodemarket = rsAvailable("productcode")
			productdescmarket = rsAvailable("productdescription")
			fcbamarket = rsAvailable("fcba")
			postingdate = rsAvailable("postingdate")
			state = rsAvailable("state")
			rsAvailable.movenext
		wend
	End if
	rsAvailable.Close : Set rsAvailable = Nothing
	
	' if cdbl(to_nol(initialAvailable)) = 0 then
	' 	unitprice = salesprices 
	' Else
	' 	unitprice = initialAvailable
	' End If
	' response.write "ratemarket" & ratemarket
	Else
	'----------- Query get Shipment Invoice Baru
		q	=	"  SELECT shd.fcba," & _
				"         shd.purchaseshipmentno," & _
				"         sah.purchasecontractno," & _
				"         ssd.postingdate," & _
				"         SUM (ssd.acceptquantity) acceptquantity," & _
				"         sah.suppliercode," & _
				"         sad.productcode," & _
				"         sad.prices salesprices," & _
				"         pd.productdescription," & _
				"         sidt.state" & _
				"    FROM purchaseshipmenthdr shd" & _
				"         LEFT JOIN purchaseshipmentdtl ssd" & _
				"            ON (    shd.purchaseshipmentno = ssd.purchaseshipmentno" & _
				"                AND shd.fcba = ssd.fcba)" & _
				"         LEFT JOIN purchaseagreementhdr sah" & _
				"            ON (    sah.purchasecontractno = shd.purchasecontractno" & _
				"                AND sah.fcba = shd.fcba)" & _
				"         LEFT JOIN purchaseagreementdet sad" & _
				"            ON (    sah.purchasecontractno = sad.purchasecontractno" & _
				"                AND sah.fcba = sad.fcba)" & _
				"         LEFT JOIN product pd" & _
				"            ON (pd.productcode = sad.productcode AND pd.fcba = sad.fcba)" & _
				"         LEFT JOIN purchaseinvoicedtl sidt" & _
				"            ON (    sidt.purchaseshipmentno = ssd.purchaseshipmentno" & _
				"                AND sidt.fcba = ssd.fcba)" & _
				"   WHERE     shd.purchasecontractno = '"& salescontractno &"'" & _
				"         AND ssd.fcba = '"& fcba &"'" & _
				"         AND ssd.postingdate >= TO_DATE ('"& valid_from &"', 'mm/dd/yyyy')" & _
				"         AND ssd.postingdate <= TO_DATE ('"& valid_until &"', 'mm/dd/yyyy')" & _
				"		  AND sidt.state is null" & _
				" GROUP BY shd.fcba," & _
				"         shd.purchaseshipmentno," & _
				"         sah.purchasecontractno," & _
				"         ssd.postingdate," & _
				"         sah.suppliercode," & _
				"         sad.productcode," & _
				"         sad.prices," & _
				"         pd.productdescription," & _
				"         sidt.state" & _ 
				" ORDER BY ssd.postingdate, shd.fcba"

	'response.Write ("muncul 2:" & q & "<br>")
	set rsVH = server.CreateObject("Adodb.recordset")
	rsVH.Open q, connstr

	if Not rsVH.eof then
		while not rsVH.Eof
		salescontractno = rsVH("purchasecontractno")
		quantity = rsVH("acceptquantity")
		postingdate = rsVH("postingdate")
		productcode = rsVH("productcode")
		customercode = rsVH("suppliercode")
		salesprices = rsVH("salesprices")
		productdescription = ("productdescription")
		fcBA = rsVH("fcBA")
		state = rsVH("state")
		rsVH.movenext
		wend
	end if
	rsVH.Close : Set rsVH = Nothing

	strAvailable =	"  SELECT mpp.fcba," & _
					"         mpp.purchaseshipmentno," & _
					"         mpp.postingdate," & _
					"         mp.productcode," & _
					"         mpp.productdescription," & _
					"         mpp.suppliercode customercode," & _
					"         SUM (mpp.acceptquantity) acceptquantity," & _
					"         NVL (mp.rate, 0) salesprices," & _
					"         mpp.state" & _
					"    FROM marketprice mp" & _
					"         INNER JOIN" & _
					"         (  SELECT MAX (valid_from) valid_from," & _
					"                   sad.productcode," & _
					"                   ssd.acceptquantity," & _
					"                   pd.productdescription," & _
					"                   sah.suppliercode," & _
					"                   ssh.purchaseshipmentno," & _
					"                   ssh.fcba," & _
					"                   ssd.postingdate," & _
					"                   sidt.state" & _
					"              FROM marketprice mp" & _
					"                   LEFT JOIN purchaseagreementdet sad" & _
					"                      ON (sad.productcode = mp.productcode AND sad.fcba = mp.fcba)" & _
					"                   LEFT JOIN purchaseshipmenthdr ssh" & _
					"                      ON (    ssh.purchasecontractno = sad.purchasecontractno" & _
					"                          AND ssh.fcba = sad.fcba)" & _
					"                   LEFT JOIN purchaseagreementhdr sah" & _
					"                      ON (    sah.purchasecontractno = ssh.purchasecontractno" & _
					"                          AND sah.fcba = ssh.fcba)" & _
					"                   LEFT JOIN purchaseshipmentdtl ssd" & _
					"                      ON (    ssd.purchaseshipmentno = ssh.purchaseshipmentno" & _
					"                          AND ssd.fcba = ssh.fcba)" & _
					"                   LEFT JOIN product pd" & _
					"                      ON (pd.productcode = sad.productcode AND pd.fcba = sad.fcba)" & _
					"                   LEFT JOIN purchaseinvoicedtl sidt" & _
					"                      ON (    sidt.purchaseshipmentno = ssd.purchaseshipmentno" & _
					"                          AND sidt.fcba = ssd.fcba)" & _
					"             WHERE     mp.productcode = '"& productcode &"'" & _
					"                   AND valid_from <= ssd.postingdate" & _
					"                   AND sah.suppliercode = '"& customercode &"'" & _
					"                   AND mp.fcba = '"& fcba &"'" & _
					"                   AND sah.purchasecontractno = '"& salescontractno &"'" & _
					"                   AND ssd.postingdate >= TO_DATE ('"& valid_from &"', 'mm/dd/yyyy')" & _
					"                   AND ssd.postingdate <= TO_DATE ('"& valid_until &"', 'mm/dd/yyyy')" & _
					"                   AND sidt.state IS NULL" & _
					"          GROUP BY sad.productcode," & _
					"                   ssd.acceptquantity," & _
					"                   pd.productdescription," & _
					"                   sah.suppliercode," & _
					"                   ssh.purchaseshipmentno," & _
					"                   ssd.postingdate," & _
					"                   ssh.fcba," & _
					"                   sidt.state) mpp" & _
					"            ON (    mpp.valid_from = mp.valid_from" & _
					"                AND mp.productcode = '"& productcode &"'" & _
					"                AND mp.customercode = '"& customercode &"'" & _
					"                AND mp.fcba = mpp.fcba)" & _
					" GROUP BY mpp.fcba," & _
					"         mpp.purchaseshipmentno," & _
					"         mpp.postingdate," & _
					"         mp.productcode," & _
					"         mpp.productdescription," & _
					"         mpp.suppliercode," & _
					"         NVL (mp.rate, 0)," & _
					"         mpp.state" & _ 
					"	ORDER BY mpp.postingdate, mpp.fcba"

	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr
	
	If Not rsAvailable.EOF Then
		while Not rsAvailable.Eof
			salesshipmentnomarket = rsAvailable("purchaseshipmentno")
			ratemarket = Cdbl(to_nol(rsAvailable("salesprices")))
			qtymarket =  Cdbl(to_nol(rsAvailable("acceptquantity")))
			customermarket = rsAvailable("customercode")
			postingdatemarket = rsAvailable("postingdate")
			productcodemarket = rsAvailable("productcode")
			productdescmarket = rsAvailable("productdescription")
			fcbamarket = rsAvailable("fcba")
			postingdate = rsAvailable("postingdate")
			state = rsAvailable("state")
			rsAvailable.movenext
		wend
	End if
	rsAvailable.Close : Set rsAvailable = Nothing	 
	End If
%>
<% if cdbl(to_nol(ratemarket)) = 0 then %>
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
	<form action="purchaseinvoice.asp" method="post" name="eForm">
	   <table class="data" border="1" width="100%">
	   <%if SudahClose <> "Y" Then%>
		   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
			   <td align="center">No</td> 
			   <td align="center">Business Unit</td>
			   <td align="center">Shipment No</td>
			   <td align="center">Date</td>
			   <td align="center">Product</td>
			   <td align="center">Quantity</td>
			   <td align="center">Unit Price </td>
			   <td align="center">Amount</td>
			   <td align="center"><input type="checkbox" name="select-all" id="select-all"/> Check/Uncheck All</td>
			</tr>
		<%ELSE%>
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
			   <td align="center">No</td>
			   <td align="center"></td>
			   <td align="center">Business Unit</td>
			   <td align="center">Shipment No</td>
			   <td align="center">Date</td>
			   <td align="center">Product</td>
			   <td align="center">Quantity</td>
			   <td align="center">Unit Price </td>
			   <td align="center">Amount</td>
			</tr>
		<%End If%>
	   <% set rsVH = server.CreateObject("Adodb.recordset")
	      'response.Write ("muncul 3:" & q & "<br>")
		  rsVH.Open q, connstr
	      if Not rsVH.eof then
	      nCountofField=rsVH.fields.count-1
		  no=1
		  totqty=0
		  totamount=0
		  While Not rsVH.EOF
		  	  amount = cdbl(to_nol(rsVH("salesprices"))) * cdbl(to_nol(rsVH("acceptquantity")))
			  totqty =   totqty + Cdbl(to_nol(rsVH("acceptquantity")))
			  totamount = totamount + amount
			  postingdate = rsVH("postingdate")
			  if no  mod 2 = 1 then
	             cColor = "#EEEEEE"
			  else
			     cColor = "#E5E5E5"
			  end if%>
			  <tr bgcolor="<%=cColor%>">
			  <td align="center" valign="top"><%=NO%></td>
				<%IF SudahClose="Y" then%>
					<!-- validate period control in here -->
						<td>
							<a href="purchaseinvoice.asp?invoicecode=<%=invoicecode%>&salesshipmentno=<%=rsVH("purchaseshipmentno")%>&salescontractno=<%=purchasecontractno%>&customercode=<%=rsVH("suppliercode")%>&fcba=<%=fcba%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
						</td> 
					<!-- end validate period control in here -->
						<td align="center" title="" valign="top"><%=rsVH("fcba")%></td>
						<td align="center" title="" valign="top"><%=rsVH("purchaseshipmentno")%></td>
						<td align="center" title="" valign="top"><%=rsVH("postingdate")%></td>
						<td align="center" title="" valign="top"><%=rsVH("productdescription")%></td>
				        <td align="right" title="" valign="top"><%=FbpNum2(rsVH("acceptquantity"))%></td>
				        
						<td align="center">
							<input style="text-align:right;background-color:#eee;" type="text" name="salesprices_<%=no%>" class="amount" value = "<%=FormatNumber(salesprices,0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
						</td>
						<td align="center">
							<input style="text-align:right;background-color:#eee;" type="text" name="amount_<%=no%>" class="amount" value = "<%=FormatNumber(amount,0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
						</td>
				<%End if%>
				<%if SudahClose<>"Y" Then%>
						<td align="center" title="" valign="top"><%=rsVH("fcba")%></td>
						<td align="center" title="" valign="top"><%=rsVH("purchaseshipmentno")%></td>
						<td align="center" title="" valign="top"><%=rsVH("postingdate")%></td>
						<td align="center" title="" valign="top"><%=rsVH("productdescription")%></td>
				        <td align="right" title="" valign="top"><%=FbpNum2(rsVH("acceptquantity"))%></td>
				        <td align="center">
					    <input style="text-align:right;" type="text" name="salesprices_<%=no%>" value = "<%=salesprices%>">
						</td>
						
						<td align="center">
					    <input style="text-align:right;background-color:#eee;" type="text" name="amount_<%=no%>" class="amount" value = "<%=FormatNumber(amount,0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
						</td>
						<td align="center">
							<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
						</td>
						<input type="hidden" name="fcba_<%=no%>" value="<%=rsVH("fcba")%>">
						<input type="hidden" name="salesshipmentno_<%=no%>" value="<%=rsVH("purchaseshipmentno")%>">
						<input type="hidden" name="postingdate_<%=no%>" value="<%=rsVH("postingdate")%>">
						<input type="hidden" name="productcode_<%=no%>" value="<%=rsVH("productcode")%>">
						<input type="hidden" name="acceptquantity_<%=no%>" value="<%=rsVH("acceptquantity")%>">
					 </td>
				<%End if%>
			  </tr>
			  <%
			  no=no+1
			  rsVH.MoveNext
			  intRecordsPerPage = intRecordsPerPage - 1
		  Wend
	   else
	      response.Write("Data is empty")	  
	   end if 
	   rsVH.Close
	   set rsVH = Nothing%>
	   
	      <tr>
		  <%if SudahClose <> "Y" Then%>
		    <td colspan="5" align="left"></td>
		  <%Else%>
		    <td colspan="6" align="left"></td>
		  <%End if%>
		  
			<td align="right" title="" valign="top"><%=FormatNumber(totqty,2)%></td>
			<td align="center" title="" valign="top"></td>
			<td align="right" title="" valign="top"><%=FormatNumber(totamount,2)%></td>
			
			
		  </tr>
		  <tr><td colspan="10" align="left">
			<!-- diambil dari form yang diatas -->
		    <input type="hidden" name="fcCode" value="<%=fcCode%>">
		    <input type="hidden" name="selectedinvoice" value="<%=selectedinvoice%>">
		    <input type="hidden" name="invoice_by" value="<%=invoice_by%>">
			<input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
			<input type="hidden" name="No" value="<%=No%>">
			<input type="hidden" name="customercode" value="<%=customercode%>">
			<input type="hidden" name="subtotal" value="<%=subtotal%>">  
			<input type="hidden" name="pph" value="<%=pph%>">
			<input type="hidden" name="pphamount" value="<%=pphamount%>">  
			<input type="hidden" name="grandtotal" value="<%=grandtotal%>">  
			<input type="hidden" name="invoicecode" value="<%=invoicecode%>">
			<input type="hidden" name="fcba" value="<%=fcba%>">
			<input type="hidden" name="invoicedpno" value="<%=invoicedpno%>">
			<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
			<input type="hidden" name="salescontractno" value="<%=salescontractno%>">
			<input type="hidden" name="selectedfccompanycode" value="<%=selectedfccompanycode%>">
			<input type="hidden" name="selectedcustomercode" value="<%=selectedcustomercode%>">
			<input type="hidden" name="selectedsalescontractno" value="<%=selectedsalescontractno%>">
			<input type="hidden" name="selectedlinenosales" value="<%=selectedlinenosales%>">
			<input type="hidden" name="linenoinvoicedp" value="<%=linenoinvoicedp%>">
			<input type="hidden" name="vat" value="<%=vat%>">
			<input type="hidden" name="amountvat" value="<%=amountvat%>">
			<input type="hidden" name="totalamount" value="<%=totalamount%>">
			<input type="hidden" name="no_faktur" value="<%=no_faktur%>">
			<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
			<input type="hidden" name="valid_from" value="<%=valid_from%>">
			<input type="hidden" name="valid_until" value="<%=valid_until%>">
			<input type="hidden" name="expire_date" value="<%=expire_date%>">
			<input Type="Hidden" Value="<%=request("transdate")%>" Name="transdate">

		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <!-- validasi period control in here -->
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Submit">
		  <%End if%>
		  <!-- end validasi period control in here -->
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		  </tr>
	   </TABLE>
   </form>
<%else%>
   <form action="purchaseinvoice.asp" method="post" name="eForm">
	   <table class="data" border="1" width="100%">
	   <%if SudahClose <> "Y" Then%>
		   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
			   <td align="center">No</td>
			   <td align="center">Business Unit</td>
			   <td align="center">Shipment No</td>
			   <td align="center">Date</td>
			   <td align="center">Product</td>
			   <td align="center">Quantity</td>
			   <td align="center">Unit Price </td>
			   <td align="center">Amount</td>
			   <td align="center"><input type="checkbox" name="select-all" id="select-all"/> Check/Uncheck All</td>
			</tr>
		<%ELSE%>
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
			   <td align="center">No</td>
			   <td align="center">Business Unit</td>
			   <td align="center">Shipment No</td>
			   <td align="center">Date</td>
			   <td align="center">Product</td>
			   <td align="center">Quantity</td>
			   <td align="center">Unit Price </td>
			   <td align="center">Amount</td>
			</tr>
		<%End If%>

	   <%   Set rsAvailable = Server.CreateObject("ADODB.Recordset")
			rsAvailable.Open strAvailable, connstr
		    if Not rsAvailable.eof then
		      nCountofField=rsAvailable.fields.count-1
			  no=1
			  totqty=0
			  totamount=0
			While Not rsAvailable.EOF
				' Untuk menghitung Amount Market dari market price
			  amount =   Cdbl(to_nol(rsAvailable("acceptquantity"))) * Cdbl(to_nol(rsAvailable("salesprices")))
			  totqty =   totqty + Cdbl(to_nol(rsAvailable("acceptquantity")))
			  totamount = totamount + amount
			  postingdate = rsAvailable("postingdate")
			  if no  mod 2 = 1 then
	             cColor = "#EEEEEE"
			  else
			     cColor = "#E5E5E5"
			  end if%>
			  <tr bgcolor="<%=cColor%>">
			  <td align="center" valign="top"><%=NO%></td>
				<%IF SudahClose="Y" then%>
					<!-- validate period control in here -->
						<td>
							<a href="purchaseinvoice.asp?invoicecode=<%=invoicecode%>&salesshipmentno=<%=rsVH("purchaseshipmentno")%>&salescontractno=<%=salescontractno%>&customercode=<%=rsVH("suppliercode")%>&fcba=<%=fcba%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
						</td> 
					<!-- end validate period control in here -->
						<td align="center" title="" valign="top"><%=rsAvailable("fcba")%></td>
						<td align="center" title="" valign="top"><%=rsAvailable("purchaseshipmentno")%></td>
						<td align="center" title="" valign="top"><%=rsAvailable("postingdate")%></td>
				        <td align="center" title="" valign="top"><%=rsAvailable("productdescription")%></td>
				        <td align="right" title="" valign="top"><%=FbpNum2(rsAvailable("acceptquantity"))%></td>
						
						<td align="center">
							<input style="text-align:right;background-color:#eee;" type="text" name="salesprices<%=no%>" class="amount" value = "<%=FormatNumber(salesprices,0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
						</td>
						<td align="center">
							<input style="text-align:right;background-color:#eee;" type="text" name="amount_<%=no%>" class="amount" value = "<%=FormatNumber(amount,0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
						</td>
				<%End If%>
				<%if SudahClose<>"Y" Then%>
						<td align="center" title="" valign="top"><%=rsAvailable("fcba")%></td>
						<td align="center" title="" valign="top"><%=rsAvailable("purchaseshipmentno")%></td>
						<td align="center" title="" valign="top"><%=rsAvailable("postingdate")%></td>
				        <td align="center" title="" valign="top"><%=rsAvailable("productdescription")%></td>
				        <td align="right" title="" valign="top"><%=FbpNum2(rsAvailable("acceptquantity"))%></td>
						<td align="center">
						    <input style="text-align:right;" type="text" name="salesprices_<%=no%>" value = "<%=salesprices%>">
						</td>
						<td align="center">
							<input style="text-align:right;background-color:#eee;" type="text" name="amount_<%=no%>" class="amount" value = "<%=FormatNumber(amount,0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
						</td>
						<td align="center">
							<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
						</td>
						<input type="hidden" name="fcba_<%=no%>" value="<%=rsAvailable("fcba")%>">
						<input type="hidden" name="salesshipmentno_<%=no%>" value="<%=rsAvailable("purchaseshipmentno")%>">
						<input type="hidden" name="postingdate_<%=no%>" value="<%=rsAvailable("postingdate")%>">
						<input type="hidden" name="salesprices_<%=no%>" value="<%=rsAvailable("salesprices")%>">
						<input type="hidden" name="productcode_<%=no%>" value="<%=rsAvailable("productcode")%>">
						<input type="hidden" name="acceptquantity_<%=no%>" value="<%=rsAvailable("acceptquantity")%>">
					 </td>
				<%End If%>
			  </tr>
			  <%
			  no=no+1
			  rsAvailable.MoveNext
			  intRecordsPerPage = intRecordsPerPage - 1
		  Wend
	   else
	      response.Write("Data is empty")	  
	   end if 
	   rsAvailable.Close
	   set rsAvailable = Nothing%>
		  <tr>
		    <td colspan="6" align="left"></td>
			<td align="right" title="" valign="top"><%=FormatNumber(totqty,2)%></td>
			<td align="center" title="" valign="top"></td>
			<td align="right" title="" valign="top"><%=FormatNumber(totamount,2)%></td>
			
			
		  </tr>
		  <tr><td colspan="10" align="left">
			<!-- diambil dari form yang diatas -->
		    <input type="hidden" name="fcCode" value="<%=fcCode%>">
		    <input type="hidden" name="selectedinvoice" value="<%=selectedinvoice%>">
		    <input type="hidden" name="invoice_by" value="<%=invoice_by%>">
			<input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
			<input type="hidden" name="No" value="<%=No%>">
			<input type="hidden" name="customercode" value="<%=customercode%>">
			<input type="hidden" name="subtotal" value="<%=subtotal%>">  
			<input type="hidden" name="pph" value="<%=pph%>">
			<input type="hidden" name="pphamount" value="<%=pphamount%>">  
			<input type="hidden" name="grandtotal" value="<%=grandtotal%>">  
			<input type="hidden" name="invoicecode" value="<%=invoicecode%>">
			<input type="hidden" name="fcba" value="<%=fcba%>">
			<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
			<input type="hidden" name="salescontractno" value="<%=salescontractno%>">
			<input type="hidden" name="selectedfccompanycode" value="<%=selectedfccompanycode%>">
			<input type="hidden" name="selectedcustomercode" value="<%=selectedcustomercode%>">
			<input type="hidden" name="selectedsalescontractno" value="<%=selectedsalescontractno%>">
			<input type="hidden" name="selectedinvoicedpno" value="<%=selectedinvoicedpno%>">
			<input type="hidden" name="selectedlinenosales" value="<%=selectedlinenosales%>">
			<input type="hidden" name="linenoinvoicedp" value="<%=linenoinvoicedp%>">
			<input type="hidden" name="invoicedpno" value="<%=invoicedpno%>">
			<input type="hidden" name="vat" value="<%=vat%>">
			<input type="hidden" name="amountvat" value="<%=amountvat%>">
			<input type="hidden" name="dp_aftervat" value="<%=dp_aftervat%>">
			<input type="hidden" name="totalamount" value="<%=totalamount%>">
			<input type="hidden" name="no_faktur" value="<%=no_faktur%>">
			<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
			<input type="hidden" name="valid_from" value="<%=valid_from%>">
			<input type="hidden" name="valid_until" value="<%=valid_until%>">
			<input type="hidden" name="expire_date" value="<%=expire_date%>">
			<input Type="Hidden"  Name="transdate" Value="<%=request("transdate")%>">

		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Submit">
		  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		  </tr>
	   </TABLE>
   </form>
<%End if%> 
<%End if%>  
 
  <!-------------akhir!------------->
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
	  '.chosen-select-width'     : {width:"45%"}
	}
	for (var selector in config) {
	  $(selector).chosen(config[selector]);
	}
	// $(".chosen-select").trigger("chosen:updated");

</script>
<script type="text/javascript">
	function handleClick(cb){
		if(cb.checked == 1){
			cb.value = 1;
		} else {
			cb.value = 0;
		}
	}
	
	function CheckCheckboxes1(chk){
		var txt = chk.parentNode.parentNode.cells[6].getElementsByTagName('input')[0];
		
		if(chk.checked == true)
		{
			txt.readOnly = true;
			txt.style.backgroundColor = "#eee";
		}
		else
		{
			txt.readOnly = true;
			txt.style.backgroundColor = "#eee";
		}
	}
</script>
<script type="text/javascript">	
	$('#select-all').click( function () {
		$('.item').prop('checked', this.checked);
		$('.item').val(this.checked ? 1 : 0);
		if($('#select-all').is(":checked")){
			$('.amount').css("background-color", "#eee");
			$('.amount').prop('readOnly', true);
		}else{
			$('.amount').css("background-color", "#eee");
			$('.amount').prop('readOnly', true);
		}
	})
</script>
<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
}

function getselectedcustomer() {
	var scustomercode = document.getElementById("selectedcustomercode");
	scustomercode.value = document.getElementById("customercode").value;
}

function getselectedsalescontractno() {
	var ssalescontractno = document.getElementById("selectedsalescontractno");
	ssalescontractno.value = document.getElementById("salescontractno").value;
}

function getselectedinvoice() {
	var sfcba = document.getElementById("selectedinvoice");
	sfcba.value = document.getElementById("invoice_by").value;
}

function getselectedinvoicedpno() {
	var sinvoicedpno = document.getElementById("selectedinvoicedpno");
	sinvoicedpno.value = document.getElementById("invoicedpno").value;
}

function getselectedlinenosales() {
	var slinenosales = document.getElementById("selectedlinenosales");
	slinenosales.value = document.getElementById("linenoinvoicedp").value;
}

Document.prototype.ready = function(callback) {
  if(callback && typeof callback === 'function') {
    document.addEventListener("DOMContentLoaded", function() {
      if(document.readyState === "interactive" || document.readyState === "complete") {
        return callback();
      }
    });
  }
};

// usage
document.ready(function() {
	getlistfcba();
	getlistinvoice();
	getlistcustomer();
	getlistsalescontract();	
	getlistinvoicedp();
	getlistlinenosales();
	getdetinvoicedp();
	getselectedinvoicedpno();
	getselectedlinenosales();
});
</script>
</body>
</html>
