<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Purchase Order</title>
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

function validateFloatKeyPress(el, evt) {
    var charCode = (evt.which) ? evt.which : event.keyCode;
    var number = el.value.split('.');
    if (charCode != 46 && charCode > 31 && (charCode < 48 || charCode > 57)) {
        return false;
    }
    //just one dot
    if(number.length>1 && charCode == 46){
         return false;
    }
    //get the carat position
    var caratPos = getSelectionStart(el);
    var dotPos = el.value.indexOf(".");
    if( caratPos > dotPos && dotPos>-1 && (number[1].length > 1)){
        return false;
    }
    return true;
}

//thanks: http://javascript.nwbox.com/cursor_position/
function getSelectionStart(o) {
	if (o.createTextRange) {
		var r = document.selection.createRange().duplicate()
		r.moveEnd('character', o.value.length)
		if (r.text == '') return o.value.length
		return o.value.lastIndexOf(r.text)
	} else return o.selectionStart
}
</script>
<script type="text/javascript">
	function hitungrumus() {
		var squantity = $("#quantity").val();
		var sunitprice = $("#unitprice").val(); 
		var sotcstock1 = $("#otcstock1").val();
		var sotcstock2 = $("#otcstock2").val();
		var sotcstock3 = $("#otcstock3").val();
		var sotcstock4 = $("#otcstock4").val(); 
		var newsquantity = squantity.replace(/,/g, '');
		var newsunitprice = sunitprice.replace(/,/g, '');
		var newsotcstock1 = sotcstock1.replace(/,/g, '');
		var newsotcstock2 = sotcstock2.replace(/,/g, '');
		var newsotcstock3 = sotcstock3.replace(/,/g, '');
		var newsotcstock4 = sotcstock4.replace(/,/g, '');
		
		//Untuk menghitung Total Quantity Tolerance
		v_amount = newsquantity * newsunitprice; //a kali b
		if (!isNaN(v_amount)) {
    	 	$("#amount1").val(v_amount);
    	}
    	else{
    		$("#amount1").val();
    	}

    	//Untuk menghitung Other Charge Stock
		//
		//totalothercharge = (parseInt(newsotcstock1) + parseInt(newsotcstock2) + parseInt(newsotcstock3) + parseInt(newsotcstock4)).toFixed(2);
		totalothercharge = (parseFloat(newsotcstock1) + parseFloat(newsotcstock2) + parseFloat(newsotcstock3) + parseFloat(newsotcstock4));
		
		if (!isNaN(totalothercharge)) {
    	 	$("#otherchargestock").val(totalothercharge); 
    	}
    	else{
    		$("#otherchargestock").val();
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
var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbapo.asp', 'POST', PostStr, showlistfcba);
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
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getPurchaseBy = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedpurchaseby = document.getElementById("selectedpurchaseby").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&purchaseby='+selectedpurchaseby;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistpurchaseby.asp', 'POST', PostStr, showPurchaseBy);
}

// The function for handling the response from the server
var showPurchaseBy = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("purchasebylist").innerHTML = "";
	document.getElementById("purchasebylist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getSupplierPO = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedsupplier = document.getElementById("selectedsupplier").value;
	var selectedpurchaseby = document.getElementById("selectedpurchaseby").value;
	var PostStr = 'fcba='+selectedpurchaseby+'&fccode='+selectedsupplier;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getSupplierPO.asp', 'POST', PostStr, showSupplierPO);
}

// The function for handling the response from the server
var showSupplierPO = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("supplierlist").innerHTML = "";
	document.getElementById("supplierlist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getdeliverydest = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selecteddeliverydestination = document.getElementById("selecteddeliverydestination").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&deliverydestination='+selecteddeliverydestination;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getdeliverydestpo.asp', 'POST', PostStr, showdelivery);
}

// The function for handling the response from the server
var showdelivery = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("deliverydestinationlist").innerHTML = "";
	document.getElementById("deliverydestinationlist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getprline = function () {
	// build up the post string when passing variables to the server side page
	
	var prcode = document.getElementById("prcode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'prcode='+prcode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getprline.asp', 'POST', PostStr, showprline);
}

// The function for handling the response from the server
var showprline = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    // alert("2");
    // update the Div to show the result from the server
	document.getElementById("prlinelist").innerHTML = "";
	document.getElementById("prlinelist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}


// launched from button click 
var getStockCode = function () {
	// build up the post string when passing variables to the server side page
	var prlineid = document.getElementById("prlineid").value;
	var prcode = document.getElementById("prcode").value;
	var fcba = document.getElementById("selectedfcba").value;
	var PostStr = 'prcode='+prcode+'&prlineid='+prlineid+'&fcba='+fcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getstockcode.asp', 'POST', PostStr, showStockCodeResponse);
	
}

// The function for handling the response from the server
var showStockCodeResponse = function (oXML) { 
     // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(response);
    // update the Div to show the result from the server
    document.getElementById("itemcode").value = s.itemcode;
	document.getElementById("stockname").value = s.stockname;
	document.getElementById("requestedqty").value = s.approvedqty;
	document.getElementById("remainsqty").value = s.remainsqty;
	// document.getElementById("descpr").value = s.description;
	   
}

// launched from button click 
var getSupp = function () {
	// build up the post string when passing variables to the server side page
	var selectedsupplier = document.getElementById("selectedsupplier").value;
	var selectedpurchaseby = document.getElementById("selectedpurchaseby").value;
	var PostStr = 'suppliercode='+selectedsupplier+'&purchaseby='+selectedpurchaseby;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getSupplier2PO.asp', 'POST', PostStr, showMessageResponse);
	
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.supplieraddress);
    // update the Div to show the result from the server
    // document.getElementById("suppliername").value = s.suppliername;
	document.getElementById("supplieraddress").value = s.supplieraddress;
	// document.getElementById("address").innerHTML = document.getElementById("supplieraddress").value;
	// alert(s.suppliername);
    
}

var reloadPoCode = function () {
	var podate =$( ".datepicker" ).val();
	var selectedfcba =$( "#selectedfcba" ).val();
	var fccompanycode = $( "#fccompanycode" ).val();
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Procurement&txtTable=PurchaseOrderDetails&podateurl="+podate+"&selectedfcba="+selectedfcba+"&fccompanycode="+fccompanycode );
}

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

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 ' if request("fccOde")="" then
 '    fcCode="" '----- Nantinya Ambil dari Format Master
 ' else
 '    fcCode = request("fccOde")
 ' end if

 fcCode = request("fccOde")
 If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if request("podateurl") <>"" then
   PODATE   = request("podateurl")
 elseif Request("PODATE")<>"" then
   PODATE   =Request("PODATE")
 else
   PODATE   =textDate(date)
 end if
 
  PODATE = FormatDateTime(PODATE)
 
 xCode = Request("xCode")
 IsInBridge = Request("IsInBridge")
 Remarks = Request("Remarks")
 
 '---- Data Details
 ' Add by NAF, 2012-12-1
 LINENOKEY=Request("LineNoKey")
 ' End Added by NAF
 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 amount1 =Request("amount1")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID")
 use_ppn = Request("use_ppn")
 fccompanycode = Request("fccompanycode")
 businessunit = Request("businessunit")
 purchaseby = Request("purchaseby")
 suppliername = Request("suppliername")
 supplieraddress = Request("supplieraddress")
 termofpayment = Request("termofpayment")
 SHIPPINGPOINT = Request("SHIPPINGPOINT")
 dp = Request("dp")
 ppn = Request("ppn")
 deliverydestination = Request("deliverydestination")
 StockName = Request("StockName")
 requestedqty = Request("requestedqty")
 unitprice = Request("unitprice")
 disc = Request("disc")
 othercharge = Request("othercharge")
 otherchargestock = Request("otherchargestock")
 totalamount = Request("totalamount")
 totalamountppn = Request("totalamountppn")
 totalcharge = Request("totalcharge")
 grandtotal = Request("grandtotal")
 amountcharge = Request("amountcharge")
 totalppn = Request("totalppn")
 approvedqty = request ("approvedqty")
 document_classification = request("document_classification")
 authorizelevelcode = request("authorizelevelcode")
 remainsqty = request ("remainsqty") 
 otcstock1 = request ("otcstock1") 
 otcstock2 = request ("otcstock2") 
 otcstock3 = request ("otcstock3") 
 otcstock4 = request ("otcstock4") 
 '----
 'CurrencyCode = request ("CurrencyCode")
 if Request("CurrencyCode")<>"" then
   CurrencyCode   =Request("CurrencyCode")
 else
   CurrencyCode   ="IDR"
 end if
 
 if Request("ExchangeRate")<>"" then
   ExchangeRate   =Request("ExchangeRate")
 else
   ExchangeRate   ="1"
 end if
 'ExchangeRate = request ("ExchangeRate")
 '----
selectedpurchaseby = request("selectedpurchaseby")
selectedfcba = request("selectedfcba")
selectedsupplier = request("selectedsupplier")
selecteddeliverydestination = request("selecteddeliverydestination")

if "" & use_ppn = "" then use_ppn = "N" 
 
 'response.write "Request;" & Request("txtJenis")
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "Ordered" or Request("txtJenis")= "CompletedDetails"  or  Request("txtJenis")= "CanceledDetails"  or  Request("txtJenis")= "UnRelease" or Request("txtJenis")= "StartAuthorize" Then
	
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
       connstr.BeginTrans
		
				sqlupdateqtypr = " select prcode, prlineid from purchaseorderdetails" & _
                                 " where pocode = '" & fcCode & "'" & _
                                 "       and fcba = '" & businessunit & "'" & _
                                 "       and prcode is not null" & _
                                 "       and prlineid is not null"
								 ' response.write sqlupdateqtypr
			  	Set rsupdateqtypr = Server.CreateObject("ADODB.Recordset")
				rsupdateqtypr.Open sqlupdateqtypr, connstr, adLockOptimistic, adOpenStatic
				
				if not rsupdateqtypr.EOF then
					While Not rsupdateqtypr.EOF
							q = " update purchaserequestdetails" & _
								" set orderedqty = 0" & _
								" where  prcode = '" & rsupdateqtypr("prcode") & "'" & _
								"        and lineid = '" & rsupdateqtypr("prlineid") & "'" & _
								"        and fcba = '" & businessunit & "'"

							connStr.execute(q)
						rsupdateqtypr.MoveNext
					wend
				end if
							
				rsupdateqtypr.Close
				set rsupdateqtypr = Nothing 
				
				q="Delete from PurchaseOrderDetails where  POCODE='" & fcCode & "' and fcBA='" & businessunit & "'"
				connStr.execute(q)

				q="Delete from PurchaseOrder where  POCODE='" & fcCode & "' and fcBA='" & businessunit & "'"
				connStr.execute(q)
				
				  q=   	" Update purchaserequestdetails set orderedqty=(SELECT SUM (pod.quantity) qty" & _
						"     FROM purchaseorderdetails pod" & _
						"          INNER JOIN purchaserequestdetails prd" & _
						"             ON (pod.prcode = prd.prcode AND pod.prlineid = prd.lineid)" & _
						"    WHERE pod.prcode = '"& prcode &"' and prd.lineid = '"& prlineid &"'" & _
						" GROUP BY prd.prcode, prd.lineid, pod.itemcode) " & _
						", LASTUPDATE = sysdate " & _
						", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						" where  prcode='" & prcode & "' and ItemCode='" & ITEMCODE & "' and lineid='" & prlineid & "'"
					connStr.execute(q)
					' response.write q
					
				if err.number  = 0 then
					connStr.CommitTrans
				end if
	   cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans
			   q="Delete from PurchaseOrderDetails where  POCODE='" & fcCode & "' and ItemCode='" & request("ItemCode") & "' and LineNoKey='" & request("LineNoKey") & "' and fcBA='" & businessunit & "'"
				' response.write q
			  connStr.execute(q)

					dtl_totalamount = 0
					dtl_charge = 0
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select *  from PURCHASEORDERDETAILS " & _
					  " where poCode='" & fcCode & "' and " & _
					  " fcBA='" & businessunit & "'"
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
						dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("amount")))
						dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercharge")))
						rsTempPurchaseRequest.movenext
					wend 
						
					End if
					
					rsTempPurchaseRequest.Close
					set rsTempPurchaseRequest = Nothing
					
					oSql = 	" SELECT prd.prcode," & _
							"          prd.lineid," & _
							"          pod.itemcode," & _
							"          SUM (pod.quantity) qty" & _
							"     FROM purchaseorderdetails pod" & _
							"          INNER JOIN purchaserequestdetails prd" & _
							"             ON (pod.prcode = prd.prcode AND pod.prlineid = prd.lineid)" & _
							"    WHERE pod.prcode = '"& prcode &"' and prd.lineid = '"& prlineid &"'" & _
							" GROUP BY prd.prcode, prd.lineid, pod.itemcode"  
					' Response.Write oSql
						   
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.Open oSql, connstr 
					If Not oRs.EOF Then 
						ttl_qty =  cdbl(to_nol(oRs("qty")))  
					Else	
						ttl_qty = 0
					End If 
					oRs.Close
					Set oRs = Nothing  
					
					totalamount = dtl_totalamount
					totalppn = totalamount  * cdbl(to_nol(ppn))/100
					totalamountppn = totalamount + totalppn
					totalcharge = dtl_charge 
					grandtotal = totalamountppn + totalcharge
					v_total_qty = cdbl(to_nol(ttl_qty)) 

				  q= " Update purchaseorder set totalamountppn=replace('" & totalamountppn & "',',','')" & _
					 ", totalcharge = replace('" & totalcharge & "',',','')" & _
					 ", totalppn = replace('" & totalppn & "',',','')" & _
					 ", totalamount = replace('" & totalamount & "',',','')" & _
					 ", grandtotal = replace('" & grandtotal & "',',','')" & _
					 ", LASTUPDATE = sysdate " & _
					 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					 " where  POCode='" & fcCode & "' and fcBA='" & businessunit & "'"
					'response.write "AA:" & q
					connStr.execute(q)
					
				' q=   "Update purchaserequestdetails set orderedqty='" & v_total_qty & "'" & _
				' 	 ", LASTUPDATE = sysdate " & _
				' 	 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
				' 	 " where  prcode='" & prcode & "' and ItemCode='" & ITEMCODE & "' and lineid='" & prlineid & "'"
				' 		connStr.execute(q)
				' 		' response.write q

				q=   	" Update purchaserequestdetails set orderedqty= NVL( orderedqty - (SELECT NVL (SUM (pod.quantity), 0) qty" & _
						"     FROM purchaseorderdetails pod" & _
						"          INNER JOIN purchaserequestdetails prd" & _
						"             ON (pod.prcode = prd.prcode AND pod.prlineid = prd.lineid)" & _
						"    WHERE pod.prcode = '"& prcode &"' and prd.lineid = '"& prlineid &"' and pod.rowstateid <> 'Cancelled'" & _
						" GROUP BY prd.prcode, prd.lineid, pod.itemcode),0) " & _
						", LASTUPDATE = sysdate " & _
						", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						" where  prcode='" & prcode & "' and ItemCode='" & ITEMCODE & "' and lineid='" & prlineid & "'"
					' response.write q
					connStr.execute(q)
			   if err.number  = 0 then
				  connStr.CommitTrans
			   end if
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
	connstr.BeginTrans
		'--Add by Ukie to handle Auto Number of PR
			If Ucase(fcCode) = Ucase("(NEW)") Then
				oSql = " select nvl(controlnumber_api.get_nextnumber(upper('PURCHASEORDER')" & _
					   "                                         , " & year(textDate(PoDate)) & "" & _
					   "                                         , " & month(textDate(podate)) & "" & _
					   "                                         , '" & purchaseby & "'" & _
					   " ), 'FALSE') pocode from dual"
				
				' Response.Write oSql
					   
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.Open oSql, connstr, adLockOptimistic, adOpenStatic
				If Not oRs.EOF Then
					txtPOCode = oRs("poCode")
				End If
				oRs.Close
				Set oRs = Nothing
					fcCode = txtPOCode
			End If
			
			If txtPOCode = "FALSE" Then
				err.Number = 1001
				err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
			End If
			
			'Response.Write PRCODE & txtPRCode
		'--End Added

				 '  Dim rsPO, rsPurchases,rsQty
				 '  set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					' q="select APPROVEDQTY, ORDEREDQTY  from PURCHASEREQUESTDETAILS  " & _
					'   " where PrCode='" & PRCODE & "' and " & _
					'   " lineid = '" & PRLINEID & "' and " & _
					'   " itemcode = '" & ITEMCODE & "' and " & _
					'   " fcBA='" & businessunit & "'"

					' rsTempPurchaseRequest.Open q, connstr
					' if Not rsTempPurchaseRequest.Eof then
					' 	ApprovedQty = cdbl(to_nol(rsTempPurchaseRequest("APPROVEDQTY")))
					' 	OrderedQty = cdbl(to_nol(rsTempPurchaseRequest("ORDEREDQTY")))
						
					' 	totalPRData= 1
					' else
					' 	totalPRData = 0
					' 	ApprovedQty = 0
					' 	OrderedQty = 0
					' End if
					
					' rsTempPurchaseRequest.Close
					' set rsTempPurchaseRequest = Nothing
	  Set rsPO = Server.CreateObject("ADODB.Recordset")
			  rsPO.Open "Select poCode from PurchaseOrder where poCode='" & fcCode & "' AND FCBA = '" & businessunit & "' AND company='"& fccompanycode &"'" , connstr
			  IF rsPO.eof then	   
				 ' Modified by NAF
				 ' To handle first time Input for Purchase Order

				 q = " INSERT INTO purchaseorder ( pocode, podate, suppliercode, document_classification, remarks, use_ppn, ppn, totalamount, company, supplieraddress, purchaseby, deliverydestination, termofpayment, dp, totalamountppn, totalcharge, totalppn, grandtotal, incoterm1, fcentry, fcedit, fcip, fcba, rowstateid, lastupdate, lasttime, currencycode, exchangerate ) " & _
				     " VALUES ( '" & fcCode & "'" & _
					 "         , TO_DATE('" & PoDate & "', 'mm/dd/yyyy')" & _
					 "         , '" & xCode & "'" & _
					 "         , '" & document_classification & "'" & _
					 "         , '" & remarks & "'" & _
					 "         , '" & use_ppn & "'" & _
					 "         , '" & ppn & "'" & _
					 " 		   , replace('" & totalamount & "',',','')" & _
					 "         , '" & fccompanycode & "'" & _
					 "         , '" & supplieraddress & "'" & _
					 "         , '" & purchaseby & "'" & _
					 "         , '" & deliverydestination & "'" & _
					 "         , '" & termofpayment & "'" & _
					 " 		   , replace('" & dp & "',',','')" & _
					 " 		   , replace('" & totalamountppn & "',',','')" & _
					 " 		   , replace('" & totalcharge & "',',','')" & _
					 " 		   , replace('" & totalppn & "',',','')" & _
					 " 		   , replace('" & grandtotal & "',',','')" & _
					 "         , '" & SHIPPINGPOINT & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					 "         , '" & businessunit & "'" & _
					 "         , 'Planned'" & _
					 "         , SYSDATE" & _
					 "         , TO_CHAR(SYSDATE,'hh24:mi')" & _
					 "         , '" & currencycode & "'" & _
					 " 		   , replace('" & exchangerate & "',',','')" & _
					 ")"
				 connStr.execute(q)
				 ' response.write q 
				 ' end if
				Else
			
			  End if
			' end if
			  rsPO.Close
				' Added by NAF, 2012-12-01
				 strLineMax = " select nvl(max(linenokey),0) maxLineNoKey from purchaseorderdetails where fcBA='" & businessunit & "'"

				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
				end if
							
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 
				' End of Added by NAF, 2012-12-01
				 
				 ' if len(trim(PRCODE)) <> 0 or len(trim(PRLINEID)) <> 0 then 
					
					' if totalPRData <> 0 then 
						'Untuk menghhitung jumlah total di header
						' dtl_delivery = 0
						' delivery_quantity = 0
						set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
						ttl_quantity = 0
						q=" select nvl(sum(quantity),0) total from purchaseorderdetails " & _
						  " where " & _
						  " fcBA='" & businessunit & "' and prcode = '"& PRCODE &"' and prlineid ='"& PRLINEID &"' and rowstateid ='Ordered'"
						  ' response.write q

						  	' response.write ApprovedQty
							' response.write OrderedQty
							' response.write Quantity
							' response.write requestedqty
						  
						  rsTempPurchaseRequest.Open q, connstr
						if Not rsTempPurchaseRequest.Eof then
							while Not rsTempPurchaseRequest.Eof 
								ttl_quantity = cdbl(to_nol(rsTempPurchaseRequest("total"))) 
								rsTempPurchaseRequest.movenext
							wend 
						else
							ttl_quantity = 0
						End if
						
						rsTempPurchaseRequest.Close
						set rsTempPurchaseRequest = Nothing
						' delivery_quantity = dtl_delivery 
						' delivery_quantity = dtl_delivery_quantity + cdbl(to_nol(delivery_quantity))
						 
					' 'end add
							' oSql = 	" SELECT prd.prcode," & _
									' "          prd.lineid," & _
									' "          pod.itemcode," & _
									' "          SUM (pod.quantity) qty" & _
									' "     FROM purchaseorderdetails pod" & _
									' "          INNER JOIN purchaserequestdetails prd" & _
									' "             ON (pod.prcode = prd.prcode AND pod.prlineid = prd.lineid)" & _
									' "    WHERE pod.prcode = '"& prcode &"' and prd.lineid = '"& prlineid &"'" & _
									' " GROUP BY prd.prcode, prd.lineid, pod.itemcode"  
							' Response.Write oSql
								   
							' Set oRs = Server.CreateObject("ADODB.Recordset")
							' oRs.Open oSql, connstr 
							' If Not oRs.EOF Then 
								' ttl_qty =  cdbl(to_nol(oRs("qty")))  
							' Else	
								' ttl_qty = 0
							' End If 
							' oRs.Close
							' Set oRs = Nothing  
							
						if  cdbl(to_nol(quantity)) <= cdbl(to_nol(remainsqty)) then 
						' if cdbl(to_nol(requestedqty)) - cdbl(to_nol(quantity)) - cdbl(to_nol(remainsqty)) >= cdbl(to_nol(ttl_quantity)) then 
							 
							' response.write ttl_qty
							amount = 0
						 	' amount1 = cdbl(to_nol(quantity)) * cdbl(to_nol(unitprice))
							amount2 = amount1 * cdbl(to_nol(disc))/100
							amount3 = amount1 - amount2
							' amount4 = amount3 * cdbl(to_nol(ppn))/100
							' amount5 = amount3 + amount4
							AMOUNT = amount3 + cdbl(to_nol(otherchargestock))
							AMOUNTCHARGE = amount3 + cdbl(to_nol(othercharge))+ cdbl(to_nol(otherchargestock))
							v_total_qty = cdbl(to_nol(ttl_qty))  
							UNITPRICE_CUR = cdbl(exchangerate) * cdbl(UNITPRICE)
							PPN_OTCSTOCK3    = cdbl(to_nol(OTCSTOCK3)) *  cdbl(to_nol(ppn))/100
							q="Insert into  PURCHASEORDERDETAILS (POCODE, ITEMCODE, QUANTITY, AMOUNT, OTHERITEMDESC, " & _ 
							  "RECEIVED, UNITPRICE, STOCKNAME, REQUESTEDQTY, DISC, "& _ 
							  "OTHERCHARGE, OTHERCHARGESTOCK, OTCSTOCK1, OTCSTOCK2, OTCSTOCK3, OTCSTOCK4, AMOUNTCHARGE, FCENTRY, FCEDIT, FCIP, FCBA, LINENOKEY, PRCODE, PRLINEID, LASTUPDATE, LASTTIME, ROWSTATEID,UNITPRICECURRENCY,PPN_OTHERCHARGE ) " & _
							 "values ('" & fcCode & "'" & _
							 "	,'" & ITEMCODE & "'" & _
							 "	,'"  &  QUANTITY & "'" & _
							 "	,replace('" & amount3 & "',',','')" & _
							 "	,'" & OTHERITEMDESC & "'" & _
							 "	," & "0" & "" & _
							 "	,replace('" & UNITPRICE & "',',','')" & _
							 "	,'" & STOCKNAME & "'" & _
							 "	,replace('" & REQUESTEDQTY & "',',','')" & _
							 "	,replace('" & DISC & "',',','')" & _
							 "	,replace('" & OTHERCHARGE & "',',','')" & _
							 "	,replace('" & OTHERCHARGESTOCK & "',',','')" & _
							 "	,replace('" & OTCSTOCK1 & "',',','')" & _
							 "	,replace('" & OTCSTOCK2 & "',',','')" & _
							 "	,replace('" & OTCSTOCK3 & "',',','')" & _
							 "	,replace('" & OTCSTOCK4 & "',',','')" & _
							 "	,replace('" & AMOUNTCHARGE & "',',','')" & _
							 "	,'" & Session("CodeName")  & "'" & _
							 "	,'" & Session("CodeName") & "'" & _
							 "	,'" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
							 "	,'" & businessunit & "'" & _
							 "	,"& tempMaxLineNo & "" & _
							 "	,'" & PRCODE & "'" & _
							 "	,'"& PRLINEID &"'" & _
							 "	, sysdate, To_Char(sysdate,'HH24:MI'), 'Planned'"& _
							 "	,replace('" & UNITPRICE_CUR & "',',','')" & _
							 "	,replace('" & PPN_OTCSTOCK3 & "',',','')" & _
							 ")"
							 
							connStr.execute(q)				
							' response.write q 
							 
							'--- UPDATE Qty Order di PR
							'--Modif Aam Feb 21

							q = "select lineid from purchaserequestdetails"
							
                            q = "MERGE INTO purchaserequestdetails t1 " & _
								"USING " & _
								"( " & _
								"SELECT prcode,itemcode,fcba,prlineid,sum(quantity) qty FROM purchaseorderdetails " & _
								"where prcode = '"& prcode &"' and ItemCode='" & ITEMCODE & "' and fcba = '" & businessunit & "' and rowstateid <> 'Cancelled' " & _
								"group by prcode,itemcode,fcba,prlineid " & _
								")t2 " & _
								"ON(t1.prcode = t2.prcode and t1.itemcode=t2.itemcode and t1.fcba=t2.fcba and t1.lineid=t2.prlineid) " & _
								"WHEN MATCHED THEN UPDATE SET " & _
								"t1.orderedqty = t2.qty"


							
							 	'" Update purchaserequestdetails set orderedqty=(SELECT SUM (pod.quantity) qty" & _
								'	"     FROM purchaseorderdetails pod" & _
								'	"          INNER JOIN purchaserequestdetails prd" & _
								'	"             ON (pod.prcode = prd.prcode AND pod.prlineid = prd.lineid)" & _
								'	"    WHERE pod.prcode = '"& prcode &"' and prd.lineid = '"& prlineid &"' and pod.rowstateid <> 'Cancelled'" & _
								'	" GROUP BY prd.prcode, prd.lineid, pod.itemcode) " & _
								'	", LASTUPDATE = sysdate " & _
								'	", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
								'	" where  prcode='" & prcode & "' and ItemCode='" & ITEMCODE & "' and lineid='" & prlineid & "'"
								
								connStr.execute(q)
								' response.write q
							
							else
								err.number=1001
								err.description = "Your Order Quantity is over from Requested Quantity, Please Check other PO that has been ordered"		
							end if
							'response.write quantity&"</br>"
							' else
							' 	err.number=1001
							' 	err.description = "PR Information does not match, Check the PR Number and PR line"
								
							' End if	
						' end if
						'  else
					 ' 	amount = 0
					 ' 	amount1 = cdbl(to_nol(quantity)) * cdbl(to_nol(unitprice))
						' amount2 = amount1 * cdbl(to_nol(disc))/100
						' amount3 = amount1 - amount2
						' ' amount4 = amount3 * cdbl(to_nol(ppn))/100
						' ' amount5 = amount3 + amount4
						' AMOUNT = amount + amount3
						' AMOUNTCHARGE = amount3 + cdbl(to_nol(othercharge))
						
						' q="Insert into  PURCHASEORDERDETAILS (POCODE, ITEMCODE, QUANTITY, AMOUNT, OTHERITEMDESC, RECEIVED, UNITPRICE, STOCKNAME, REQUESTEDQTY, DISC, OTHERCHARGE, AMOUNTCHARGE, FCENTRY, FCEDIT, FCIP, FCBA, LINENOKEY, PRCODE, PRLINEID, LASTUPDATE, LASTTIME, ROWSTATEID ) " & _
						' 		 "values ('" & fcCode & "','" & ITEMCODE & "','"  &  QUANTITY & "',replace('" & AMOUNT & "',',','')" & _
						' 		 ",'" & OTHERITEMDESC & "'," & "0" & ",replace('" & UNITPRICE & "',',',''),'" & STOCKNAME & "','" & REQUESTEDQTY & "',replace('" & DISC & "',',',''),replace('" & OTHERCHARGE & "',',',''),replace('" & AMOUNTCHARGE & "',',',''),'" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
						' 		 Request.ServerVariables("REMOTE_ADDR") &"','" & businessunit & "',"& tempMaxLineNo & ",'" & PRCODE & "','"& PRLINEID &"', sysdate, To_Char(sysdate,'HH24:MI'), 'Planned')"
						 
						' connStr.execute(q)				
						' ' response.write q
						
						' tempStringError = "PO Saved Successfully without PR Info, thank you."				 
						' response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
					 ' end if
				 ' end if
			  if err.number  = 0 then
				 connStr.CommitTrans
			  end if
      cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	   connstr.BeginTrans
	   			'Untuk menghhitung jumlah total di header
						' dtl_delivery = 0
						' delivery_quantity = 0
						
						update_qty = 0
						q = " SELECT NVL(quantity,0) total from purchaseorderdetails " & _
						    " WHERE fcba = '" & businessunit & "' AND pocode = '" & fcCode & "' AND linenokey = '" & linenokey & "'"
						  ' response.write q
						  set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
						  rsTempPurchaseRequest.Open q, connstr
						if Not rsTempPurchaseRequest.Eof then
							update_qty = cdbl(to_nol(rsTempPurchaseRequest("total"))) 
						else
							update_qty = 0
						End if
						rsTempPurchaseRequest.Close : set rsTempPurchaseRequest = Nothing
						
						ttl_quantity = 0
						set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
						q="select sum(quantity) total from purchaseorderdetails " & _
						  " where " & _
						  " fcBA='" & businessunit & "' and prcode = '"& PRCODE &"' and prlineid ='"& PRLINEID &"'"
						  ' response.write q

						  	' response.write ApprovedQty
						  
						  rsTempPurchaseRequest.Open q, connstr
						if Not rsTempPurchaseRequest.Eof then
							while Not rsTempPurchaseRequest.Eof 
								ttl_quantity = cdbl(to_nol(rsTempPurchaseRequest("total"))) 
								rsTempPurchaseRequest.movenext
							wend 
						else
							ttl_quantity = 0
						End if
						
						rsTempPurchaseRequest.Close
						set rsTempPurchaseRequest = Nothing
						' delivery_quantity = dtl_delivery 
						' delivery_quantity = dtl_delivery_quantity + cdbl(to_nol(delivery_quantity))
					' 'end add 
						
				' if ( cdbl(to_nol(ttl_quantity)) -cdbl(to_nol(update_qty)) ) +cdbl(to_nol(Quantity)) <= cdbl(to_nol(REQUESTEDQTY)) then 
	  				if ( cdbl(to_nol(ttl_quantity)) -cdbl(to_nol(update_qty)) - cdbl(to_nol(remainsqty)) ) +cdbl(to_nol(Quantity)) <= cdbl(to_nol(REQUESTEDQTY)) then 
	  					' response.write ttl_quantity
						' response.write Quantity
						' response.write REQUESTEDQTY
						' oSql = 	" SELECT prd.prcode," & _
								' "          prd.lineid," & _
								' "          pod.itemcode," & _
								' "          SUM (pod.quantity) qty" & _
								' "     FROM purchaseorderdetails pod" & _
								' "          INNER JOIN purchaserequestdetails prd" & _
								' "             ON (pod.prcode = prd.prcode AND pod.prlineid = prd.lineid)" & _
								' "    WHERE pod.prcode = '"& prcode &"' and prd.lineid = '"& prlineid &"'" & _
								' " GROUP BY prd.prcode, prd.lineid, pod.itemcode"  
						' Response.Write oSql
							   
						' Set oRs = Server.CreateObject("ADODB.Recordset")
						' oRs.Open oSql, connstr 
						' If Not oRs.EOF Then 
							' ttl_qty =  cdbl(to_nol(oRs("qty")))  
						' Else	
							' ttl_qty = 0
						' End If 
						' oRs.Close
						' Set oRs = Nothing  
						
	  					amount = 0
					 	amount1 = cdbl(to_nol(quantity)) * cdbl(to_nol(unitprice))
						amount2 = amount1 * cdbl(to_nol(disc))/100
						amount3 = amount1 - amount2
						' amount4 = amount3 * cdbl(to_nol(ppn))/100
						' amount5 = amount3 + amount4
						AMOUNT = amount3 + cdbl(to_nol(otherchargestock))
						AMOUNTCHARGE = amount3 + cdbl(to_nol(othercharge))+ cdbl(to_nol(otherchargestock))
						' v_total_qty = cdbl(to_nol(ttl_qty)) 
						UNITPRICE_CUR = cdbl(exchangerate) * cdbl(UNITPRICE)
						PPN_OTCSTOCK3    = cdbl(to_nol(OTCSTOCK3)) *  cdbl(to_nol(ppn))/100
							
					   q="Update PurchaseOrderDetails " & _
					     "set Quantity='" & Quantity & "'" & _ 
						 ",Amount=replace('" & amount3 & "',',','')" & _ 
						 ",Amountcharge=replace('" & AMOUNTcharge & "',',','')" & _
						 ",OTHERITEMDESC='" & OTHERITEMDESC  & "'," & _
						 "PRCODE= '"& PRCODE & "'," & _
						 "PRLINEID = '" & PRLINEID & "'" & _
						 ", unitprice = replace('" & unitprice & "',',','')" & _
						 ", othercharge = replace('" & othercharge & "',',','')" & _
						 ", otcstock1 = replace('" & otcstock1 & "',',','')" & _
						 ", otcstock2 = replace('" & otcstock2 & "',',','')" & _
						 ", otcstock3 = replace('" & otcstock3 & "',',','')" & _
						 ", otcstock4 = replace('" & otcstock4 & "',',','')" & _
						 ", otherchargestock = replace('" & otherchargestock & "',',','')" & _
						 ", disc = replace('" & disc & "',',','')" & _
						 ", StockName = '" & StockName & "'" & _
						 ", LASTUPDATE = sysdate " & _
						 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						 ", UNITPRICECURRENCY = replace('" & UNITPRICE_CUR & "',',','')" & _
						 ", PPN_OTHERCHARGE = replace('" & PPN_OTCSTOCK3 & "',',','')" & _
						 " where  POCode='" & fcCode & "' and ItemCode='" & ITEMCODE & "' and LineNoKey='" & LineNoKey & "' and fcBA='" & businessunit & "'"
						connStr.execute(q)
						' response.write q
						
						'--UPDATE Qty Order di PR
						'--Modif Aam Feb 21
						q =    "MERGE INTO purchaserequestdetails t1 " & _
								"USING " & _
								"( " & _
								"SELECT prcode,itemcode,fcba,sum(quantity) qty FROM purchaseorderdetails " & _
								"where prcode = '"& prcode &"' and ItemCode='" & ITEMCODE & "' and fcba = '" & businessunit & "' and rowstateid <> 'Cancelled' " & _
								"group by prcode,itemcode,fcba " & _
								")t2 " & _
								"ON(t1.prcode = t2.prcode and t1.itemcode=t2.itemcode and t1.fcba=t2.fcba) " & _
								"WHEN MATCHED THEN UPDATE SET " & _
								"t1.orderedqty = t2.qty"

						    '	" Update purchaserequestdetails set orderedqty=(SELECT SUM (pod.quantity) qty" & _
							'	"     FROM purchaseorderdetails pod" & _
							'	"          INNER JOIN purchaserequestdetails prd" & _
							'	"             ON (pod.prcode = prd.prcode AND pod.prlineid = prd.lineid)" & _
							'	"    WHERE pod.prcode = '"& prcode &"' and prd.lineid = '"& prlineid &"'" & _
							'	" GROUP BY prd.prcode, prd.lineid, pod.itemcode) " & _
							'	", LASTUPDATE = sysdate " & _
							'	", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							'	" where  prcode='" & prcode & "' and ItemCode='" & ITEMCODE & "' and lineid='" & prlineid & "'"
							
							connStr.execute(q)
							' response.write q
				else
					err.number=1001
					err.description = "Your Order Quantity is over from Requested Quantity, Please Check other PO that has been ordered"
				end if
					dtl_totalamount = 0
					dtl_charge = 0
					dtl_transport = 0 
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select *  from PURCHASEORDERDETAILS " & _
					  " where poCode='" & fcCode & "' and " & _
					  " fcBA='" & businessunit & "'"
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
						dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("amount"))) 
						'Add Aam Jan 2021
						dtl_transport = dtl_transport + cdbl(to_nol(rsTempPurchaseRequest("otcstock3")))
						dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercharge")))+cdbl(to_nol(rsTempPurchaseRequest("otherchargestock")))
						rsTempPurchaseRequest.movenext
					wend 
						
					End if
					
					rsTempPurchaseRequest.Close
					set rsTempPurchaseRequest = Nothing
					totalamount=0
					totaltransport=0
					totalppn=0
					totalamountppn=0
					totalcharge=0
					grandtotal=0
					
					totalamount = dtl_totalamount
					totaltransport = dtl_transport
					'Update Aam Jan 21
					'totalppn = totalamount  * cdbl(to_nol(ppn))/100
					'totalamountppn = totalamount + totalppn
					totalppn = (totalamount + totaltransport)  * cdbl(to_nol(ppn))/100
					totalamountppn = (totalamount  + totalppn)
					totalcharge = dtl_charge 
					grandtotal = totalamountppn + totalcharge

				  q= " Update purchaseorder set totalamountppn=replace('" & totalamountppn & "',',','')" & _
					 ", totalcharge = replace('" & totalcharge & "',',','')" & _
					 ", totalppn = replace('" & totalppn & "',',','')" & _
					 ", totalamount = replace('" & totalamount & "',',','')" & _
					 ", grandtotal = replace('" & grandtotal & "',',','')" & _
					 ", LASTUPDATE = sysdate " & _
					 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					 " where  POCode='" & fcCode & "' and fcBA='" & businessunit & "'"
					' response.write "AB:" & q
					connStr.execute(q)
				
				if err.number  = 0 then
				   connStr.CommitTrans
				end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Confirmed" then
	   connstr.BeginTrans
	   
			set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			q="select sum(totalppn) totalppn  from PURCHASEORDER " & _
			  " where poCode='" & fcCode & "'" & _
			  " and fcBA='" & businessunit & "'"
			  ' response.write q
			rsTempPurchaseRequest.Open q, connstr
			if Not rsTempPurchaseRequest.Eof then
				while Not rsTempPurchaseRequest.Eof 
				totalppn = cdbl(to_nol(rsTempPurchaseRequest("totalppn")))
				rsTempPurchaseRequest.movenext
			wend 
				
			End if
			' response.write dtl_totalamount
			rsTempPurchaseRequest.Close
			set rsTempPurchaseRequest = Nothing
			
		if cdbl(to_nol(totalppn)) <> 0 then
			'Untuk menghhitung jumlah total di header
			'Remark Aam Jan 21
			'dtl_totalamount = 0
			'dtl_charge = 0
			'set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			'q="select sum(totalppn) totalppn  from PURCHASEORDER " & _
			'  " where poCode='" & fcCode & "'" & _
			'  " and fcBA='" & businessunit & "'"
			'  ' response.write q
			'rsTempPurchaseRequest.Open q, connstr
			'if Not rsTempPurchaseRequest.Eof then
		   	'while Not rsTempPurchaseRequest.Eof 
			'	totalppn = cdbl(to_nol(rsTempPurchaseRequest("totalppn")))
			'	rsTempPurchaseRequest.movenext
			'wend 
			'End if
			'' response.write dtl_totalamount
			'rsTempPurchaseRequest.Close
			'set rsTempPurchaseRequest = Nothing
			
			set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			totalamount=0
			totaltransport=0
			totalppn=0
			totalamountppn=0
			totalcharge=0
			grandtotal=0

			
			q="select sum(amount) amount, sum(othercharge) othercharge, sum(otherchargestock) otherchargestock, sum(otcstock3) transport from PURCHASEORDERDETAILS " & _
			  " where poCode='" & fcCode & "'" & _
			  " and fcBA='" & businessunit & "'"
			  ' response.write q
			rsTempPurchaseRequest.Open q, connstr
			if Not rsTempPurchaseRequest.Eof then
				while Not rsTempPurchaseRequest.Eof 
				totalamount = cdbl(to_nol(rsTempPurchaseRequest("amount")))
				totalothercharge = cdbl(to_nol(rsTempPurchaseRequest("othercharge"))) 
				totalotherchargestock = cdbl(to_nol(rsTempPurchaseRequest("otherchargestock"))) 
				totaltransport = cdbl(to_nol(rsTempPurchaseRequest("transport"))) 
				rsTempPurchaseRequest.movenext
			wend 
				
			End if
			' response.write dtl_totalamount
			rsTempPurchaseRequest.Close
			set rsTempPurchaseRequest = Nothing
			
			set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			q="select nvl(ppn,0) ppn  from PURCHASEORDER " & _
			  " where poCode='" & fcCode & "'" & _
			  " and fcBA='" & businessunit & "'"
			  ' response.write q
			rsTempPurchaseRequest.Open q, connstr
			if Not rsTempPurchaseRequest.Eof then
				while Not rsTempPurchaseRequest.Eof 
				ppn = cdbl(to_nol(rsTempPurchaseRequest("ppn")))
				rsTempPurchaseRequest.movenext
			wend 
				
			End if
			' response.write dtl_totalamount
			rsTempPurchaseRequest.Close
			set rsTempPurchaseRequest = Nothing 
			
			'Update Aam Jan 21
			'totalcharge = totalothercharge  + totalotherchargestock
			'totalamount = totalamount  - totalotherchargestock
			'totalamountppn = totalamount + totalppn
			'grandtotal = totalamountppn + totalcharge
			
			totalcharge = totalothercharge  + totalotherchargestock
			totalppn = (totalamount + totaltransport)  * cdbl(to_nol(ppn))/100
			totalamountppn = (totalamount  + totalppn)
			grandtotal = totalamountppn + totalcharge
			
			'end add
	        'response.write "?" & totalamount
		
		q= " Update purchaseorder set totalamountppn=replace('" & totalamountppn & "',',','')" & _
			 ", totalcharge = replace('" & totalcharge & "',',','')" & _
			 ", totalppn = replace('" & totalppn & "',',','')" & _
			 ", totalamount = replace('" & totalamount & "',',','')" & _
			 ", grandtotal = replace('" & grandtotal & "',',','')" & _
			 ", LASTUPDATE = sysdate " & _
			 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
			 " where  POCode='" & fcCode & "' and fcBA='" & businessunit & "'"
	'		 response.write "AC:" & q
			connStr.execute(q)
		else
			dtl_totalamount = 0
			dtl_charge = 0
			set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			q="select sum(amount) amount, sum(othercharge) othercharge, sum(otherchargestock) otherchargestock, sum(otcstock3) transport from PURCHASEORDERDETAILS " & _
			  " where poCode='" & fcCode & "'" & _
			  " and fcBA='" & businessunit & "'"
			  ' response.write q
			rsTempPurchaseRequest.Open q, connstr
			if Not rsTempPurchaseRequest.Eof then
				while Not rsTempPurchaseRequest.Eof 
				totalamount = cdbl(to_nol(rsTempPurchaseRequest("amount")))
				totalothercharge = cdbl(to_nol(rsTempPurchaseRequest("othercharge"))) 
				totalotherchargestock = cdbl(to_nol(rsTempPurchaseRequest("otherchargestock"))) 
				totaltransport = cdbl(to_nol(rsTempPurchaseRequest("transport"))) 
				rsTempPurchaseRequest.movenext
			wend 
				
			End if
			' response.write dtl_totalamount
			rsTempPurchaseRequest.Close
			set rsTempPurchaseRequest = Nothing 
			'end add
			
			set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			q="select nvl(ppn,0)ppn  from PURCHASEORDER " & _
			  " where poCode='" & fcCode & "'" & _
			  " and fcBA='" & businessunit & "'"
			  ' response.write q
			rsTempPurchaseRequest.Open q, connstr
			if Not rsTempPurchaseRequest.Eof then
				while Not rsTempPurchaseRequest.Eof 
				ppn = cdbl(to_nol(rsTempPurchaseRequest("ppn")))
				rsTempPurchaseRequest.movenext
			wend 
				
			End if
			' response.write dtl_totalamount
			rsTempPurchaseRequest.Close
			set rsTempPurchaseRequest = Nothing 
			
			' Update Aam Jan 21
			'totalcharge = totalothercharge  + totalotherchargestock 
			'totalamount = totalamount  - totalotherchargestock
			'totalppn = totalamount  * cdbl(to_nol(ppn))/100
			'totalamountppn = totalamount + totalppn
			'grandtotal = totalamountppn + totalcharge

			totalcharge = totalothercharge  + totalotherchargestock
			totalppn = (totalamount + totaltransport)  * cdbl(to_nol(ppn))/100
			totalamountppn = (totalamount  + totalppn)
			grandtotal = totalamountppn + totalcharge


			'end add

		  q= " Update purchaseorder set totalamountppn=replace('" & totalamountppn & "',',','')" & _
			 ", totalcharge = replace('" & totalcharge & "',',','')" & _
			 ", totalamount = replace('" & totalamount & "',',','')" & _
			 ", totalppn = replace('" & totalppn & "',',','')" & _
			 ", grandtotal = replace('" & grandtotal & "',',','')" & _
			 ", LASTUPDATE = sysdate " & _
			 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
			 " where  POCode='" & fcCode & "' and fcBA='" & businessunit & "'"
	'	    response.write"AD:" &  q
			connStr.execute(q)
		end if
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'Confirm' ); END; "
	   connStr.execute(strQueryStatus)	
       cmdView	= "True"
	elseif Request("txtJenis")= "RePlanned" then
	   connstr.BeginTrans
	   q= " Update purchaseorder set totalamountppn=0" & _
			 ", totalcharge = 0" & _
			 ", totalamount =0" & _
			 ", totalppn = 0" & _
			 ", grandtotal = 0" & _
			 ", LASTUPDATE = sysdate " & _
			 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
			 " where  POCode='" & fcCode & "' and fcBA='" & businessunit & "'"
			 'response.write "AE:" & q
			connStr.execute(q)
		strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'Replan' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "Ordered" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'Order' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
    elseif Request("txtJenis")= "UnRelease"	then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.UnReleasePO('" & fcCode & "' ); END; "
	   response.Write strQueryStatus
	   connStr.execute(strQueryStatus)	 
	   cmdView	= "True"
	elseif Request("txtJenis")= "StartAuthorize" then
	   connstr.BeginTrans
	   
	   q = " SELECT al.fccode, al.startlevel, al.description " & _
			"  FROM authorizelevel al" & _
			"	   INNER JOIN authorizelevel_users alu" & _
			"		  ON (al.fccode = alu.authorizelevelcode)" & _
			" WHERE al.startlevel = 'TRUE' AND alu.fcusercode = '"& session("CodeName") &"'" 
			
		set rsPurchase = server.CreateObject("Adodb.recordset")
		rsPurchase.Open q, connstr
		IF not rsPurchase.eof then
			authorizelevelcode = rsPurchase("fccode")
		end if

	    q = " INSERT INTO purchaseorderdetails_hist ( pocode" & _
			"                                           , currentauthorizelevel" & _
			"                                           , fcentry" & _
			"                                           , fcedit" & _
			"                                           , fcip" & _
			"                                           , fcba" & _
			"                                           , lastupdate" & _
			"                                           , lasttime)"& _
			" VALUES ( '" & fcCode & "'" & _
			"          , '" & authorizelevelcode & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"          , '" & businessunit & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
		' response.write q
		connStr.execute(q)
	   
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'StartAuthorize' ); END; "
	   connStr.execute(strQueryStatus)	 
	   
       cmdView	= "True"
	elseif Request("txtJenis")= "CompletedDetails" then
		connstr.BeginTrans
		strQueryStatus = " BEGIN" & _
						 " PURCHASEORDERDETAILS_API.NextState('" & businessunit & "'" & _
	                     "                                    , '" & fcCode & "'" & _
					     "                                    , '" & txtLineNoKey & "'" & _
	                     "                                    , 'Complete'" & _
	                     "                                    );" & _
	                     " END; "
		'response.Write strQueryStatus
		connStr.execute(strQueryStatus)	 

		cmdView	= "True"
	
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if

 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select POCODE, PODATE, SUPPLIERCODE, REMARKS, FCENTRY, " & _
    "   FCEDIT, FCIP, FCBA, ROWSTATEID, LASTUPDATE, " & _
    "   LASTTIME, USE_PPN, ISINBRIDGE, INCOTERM1, INCOTERM2, " & _
    "   PPN, NVL(TOTALAMOUNTPPN,0)TOTALAMOUNTPPN, COMPANY, SUPPLIERADDRESS, PURCHASEBY, " & _
    "   DELIVERYDESTINATION, TERMOFPAYMENT, DP, TOTALCHARGE, NVL(GRANDTOTAL,0)GRANDTOTAL, " & _
    "   NVL(TOTALAMOUNT,0)TOTALAMOUNT, NVL(TOTALPPN,0)TOTALPPN, DOCUMENT_CLASSIFICATION, CURRENCYCODE, EXCHANGERATE, " & _
    "   POTYPE from PURCHASEORDER  " & _
  " where PoCode='" & fcCode & "'"
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    fccompanycode=rsPurchase("company")
    businessunit=rsPurchase("fcba")
	PoDate=rsPurchase("PoDate")
	purchaseby=rsPurchase("purchaseby")
	SHIPPINGPOINT = rsPurchase("INCOTERM1")
	xCode=rsPurchase("SUPPLIERCODE")
	supplieraddress=rsPurchase("supplieraddress")
	deliverydestination  = rsPurchase("deliverydestination")
	termofpayment  = rsPurchase("termofpayment")
	dp  = rsPurchase("dp")
	ppn  = rsPurchase("ppn")
	totalppn  = formatnumber(cdbl(to_nol(rsPurchase("totalppn"))),0)
	totalamount  = formatnumber(rsPurchase("totalamount"),0)
	totalamountppn = formatnumber(rsPurchase("totalamountppn"),0)
	totalcharge = formatnumber(rsPurchase("totalcharge"),0)
	grandtotal  = formatnumber(rsPurchase("grandtotal"),0)
	RowStateId = rsPurchase("RowStateId")
	RowStateHdr = rsPurchase("RowStateId")
	Remarks=rsPurchase("Remarks")
	selectedfcba =rsPurchase("fcba")
	selectedpurchaseby =rsPurchase("purchaseby")
	selectedsupplier =rsPurchase("suppliercode")
	selecteddeliverydestination =rsPurchase("deliverydestination")
	document_classification = rsPurchase("document_classification")
    currencycode = rsPurchase("currencycode")
	exchangerate = rsPurchase("exchangerate")
 else
 ' 	PoDate=rsPurchase("PoDate")
	' PODATE = FormatDateTime(PODATE)
	' xCode=rsPurchase("SUPPLIERCODE")
	' SHIPPINGPOINT = 0
	' Remarks= 0
	RowStateId = ""
	RowStateHdr =  ""
	use_ppn = 0
	IsInBridge  = 0
	totalamountppn = 0
	totalcharge = 0
	totalamount  = 0
	' totalppn  = 0
	grandtotal  = 0
 End if
 
 
 

 dim rsFrez
 set rsFrez = server.CreateObject("Adodb.recordset")
	q=" select distinct frezee  from Purchaseorderdetails " & _
	  " where PoCode='" & fcCode & "' and " & _
	  " fcBA='" & businessunit & "'"
 rsFrez.Open q, connstr
 if Not rsFrez.Eof then
	IsFrezee  = rsFrez("frezee")
 else
    IsFrezee ="0" 
 End if
 rsFrez.Close
 set rsFrez = Nothing
 
set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT fcentry, NVL(SUM(grandtotal),0) total" & _
	" FROM purchaseorder" & _
	" WHERE pocode = '" & fcCode & "'" & _
	"       AND fcba = '" & businessunit & "'" & _
	" GROUP BY fcentry"
' response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	xfcentry = rsFrez("fcentry")
	po_totalamount = cdbl(to_nol(rsFrez("total")))
else
	xfcentry = ""
	po_totalamount = 0
End if
rsFrez.Close : set rsFrez = Nothing 
 
set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'PO'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND "&po_totalamount&" BETWEEN at.amountrangefrom and at.amountrangeuntil" & _
	"       AND '"&xfcentry&"' = '" & session("codename") & "'" & _
	"       AND aus.fcusercode = '" & session("codename") & "'" & _
	"       AND aus.fcba = '" & businessunit & "'"
' response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	can_startauthorize =  cdbl(to_nol(rsFrez("total")))
else
	can_startauthorize = 0
End if
rsFrez.Close : set rsFrez = Nothing 

set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'PO'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND case when "&po_totalamount&" < 0 then 1 else "&po_totalamount&" end BETWEEN at.amountrangefrom and at.amountrangeuntil" & _
	"       AND aus.fcusercode = '" & session("codename") & "'" & _
	"       AND aus.fcba = '" & businessunit & "'"
' response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	is_final =  cdbl(to_nol(rsFrez("total")))
else
	is_final = 0
End if
rsFrez.Close : set rsFrez = Nothing 

 
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<div class='error'>" & err.description & "</div>"
End if
 '-------------ISI Payment Form
%>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Purchase Order</div><% end if %>
 <form name="frmEntry" method="post" action="PurchaseOrderDetails.Asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Purchase Order No</td>
			<%
			if fcCode = "(NEW)" then
				q = "select company_api.get_companyno(company_api.get_companycode('" & selectedfcba & "')) ||'/'||'" & selectedfcba & "/'|| MAX(cn.PREFIX) || LPAD(NVL(MAX(regexp_substr(po.POCODE, '[^/]+', 1, 5)),0000)+0001,4,0) DOCNUMBER " & _
					" from " & _
					" controlnumber cn " & _
					" left join purchaseorder po on cn.fcba = po.fcba and cn.accountingyear = to_number(to_char(po.PODATE,'yyyy')) " & _
					" where cn.fcba ='" & selectedfcba & "' " & _
					" and cn.LU_NAME ='PURCHASEORDER' "&_
					" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& PODATE &"','mm/dd/yyyy'),'yyyy')) " & _
					" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& PODATE &"','mm/dd/yyyy'),'mm')) "
		'		Response.write q
				set rsPONo = server.CreateObject("Adodb.recordset")
				rsPONo.Open q, connstr
				if Not rsPONo.Eof then
					fcCode = rsPONo("DOCNUMBER")
				End if
			End if
			%>
				<td align="left">
					<input type="text" style="width: 70%" maxlength="50" id="fcCode" name="fcCode" size="30" value="<%=fcCode%>" autocomplete="off">
					<font color="#FF0000">*</font>
				</td>
			
				<td align="left">Term Of Payment</td>
				<td align="left">
					<select name="termofpayment" value="<%=termofpayment%>" class="chosen-select" style="width: 40%">
						<option value=""></option> 
							<% 	q = " SELECT DISTINCT fccode, fcname FROM parameter WHERE parcode='TERM_OF_PAYMENT' ORDER BY fccode, fcname"
									' response.write q
									Set rsComboP = Server.CreateObject("ADODB.Recordset")
									rsComboP.CursorType = 3
									rsComboP.CursorLocation = 3
									rsComboP.Open q, connstr
									if not rsComboP.eof then%>
									<%
									Do While Not rsComboP.eof
									%>
										<option value="<%=rsComboP("fccode")%>" 
											<%if cstr(""&termofpayment)=cstr(""&rsComboP("fccode")) then %> selected 
											<%End If%>><%=rsComboP("fcname")%> 
										</option> 
									<%rsComboP.movenext
									loop
									End if
									rsComboP.Close
									set rsComboP = Nothing
									q=""
							%>
					</select>
 				</td>
		</tr>
		<tr>
			<td align="left">Company</td>
				<td align="left">
					<select name="fccompanycode" value="<%=fccompanycode%>" id="fccompanycode" onchange="javascript:getlistfcba();javascript:getPurchaseBy();javascript:getdeliverydest();" class="chosen-select" style="width: 70%">
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
			<td align="left">DP</td>
				<td align="left">
					<input type="text" maxlength="50" name="dp" size="30" value="<%=dp%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
				</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
				<td align="left">
					<div id="fcbalist"></div>
					<font color="#FF0000">*</font>
				</td>
				<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			<td align="left">PPn (%)</td>
					<td align="left">
						<input type="text" size="5" name="ppn" value="<%=ppn%>" onkeypress="return harusangka(event);">
						<input type="text" size="18" name="totalppn" value="<%=totalppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
					</td>
		</tr>
		<tr>
			<td align="left">Date</td>
				<td align="left">
					<input class="datepicker" type="text" onchange="javascript:reloadPoCode();" maxlength="10" name="PODATE" size="10" value="<%=PODATE%>">
				</td>
			<td align="left">Total Amount</td>
					<td align="left">
						<input type="text" maxlength="50" name="totalamount" size="15" value="<%=totalamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Purchase By</td>
				<td align="left">      
					<div id="purchasebylist"></div>
				</td>
				<input type="Hidden" id="selectedpurchaseby" name="selectedpurchaseby" value="<%=selectedpurchaseby%>">
				<td align="left">Total Amount + PPn</td>
					<td align="left">
						<input type="text" maxlength="50" name="totalamountppn" size="15" value="<%=totalamountppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly >
					</td>
			<tr>
				<td align="left">Inco Terms</td>	
				<td align="left">
					<select name="SHIPPINGPOINT" value="<%=SHIPPINGPOINT%>" class="chosen-select" style="width: 70%">
						<option value="<%=SHIPPINGPOINT%>"></option>
						<option value="FRC" <% If cstr(""&SHIPPINGPOINT) = "FRC" Then %> Selected <%End if%>>FRC</option> 
						<option value="FRT" <% If cstr(""&SHIPPINGPOINT) = "FRT" Then %> Selected <%End if%>>FRT</option> 
						<option value="LCO" <% If cstr(""&SHIPPINGPOINT) = "LCO" Then %> Selected <%End if%>>LCO</option> 
					</select>
				</td>	
				<td align="left">Total Charge</td>
				<td align="left"><input type="text" size="15" name="totalcharge" value="<%=totalcharge%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
				</td>	
			</tr>
		</tr>
		<tr>
		<td align="left">Classification</td>
		<td align="left">
			<select name="document_classification" class="chosen-select" style="width: 70%">
			<option></option>
			<%
			qLov = " SELECT fccode, fcname FROM document_classification where document_type= 'PO' and ISACTIVE='TRUE' ORDER BY fccode"
   
			Set rsLov = Server.CreateObject("ADODB.Recordset")
			rsLov.Open qLov, connstr
			if not rsLov.eof then
				while not rsLov.eof
			%>
				<option value="<%=rsLov("fccode")%>" <% if cstr(""&rsLov("fccode")) = cstr(""&document_classification) then%> selected <%end if%>><%=rsLov("fcname")%></option>
			<%
					rsLov.movenext
				wend
			end if
			%>
			</select>
			<font color="#FF0000">*</font>
		</td>	
			<td align="left">Grand Total</td>
			<td align="left">
				<input type="text" maxlength="50" name="grandtotal" size="15" value="<%=grandtotal%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
			</td>	
		</tr>
		<tr>
			<td align="left">Supplier</td>
				<td align="left">      
					<div id="supplierlist"></div>
				</td>
				<input type="Hidden" id="selectedsupplier" name="selectedsupplier" value="<%=selectedsupplier%>">
			<td align="left">Status</td>
			<td align="left">
				<input type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>" style="background:#eee" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Supplier Address</td>
				<td align="left">
						<textarea id="supplieraddress"  name="supplieraddress" style="width: 70%" style="background:#eee" readonly><%=supplieraddress%></textarea>
				</td>
			<td align="left">Remarks</td>
				<td align="left">
					<textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea>
				</td>
		</tr>
		<tr>
			<td align="left">Delivery Destination</td>
				<td align="left">
					<div id="deliverydestinationlist"></div>
				</td>	
				<input type="Hidden" id="selecteddeliverydestination" name="selecteddeliverydestination" value="<%=selecteddeliverydestination%>">
				
		</tr>
		</tr>
		<tr>
			<td align="left">Currency</td>
				<td align="left">
					<select name="CurrencyCode" value="<%=CurrencyCode%>" id="CurrencyCode"  class="chosen-select" style="width: 70%">
					<option value=""></option> 
						<% 	Dim rsComboCur
							q = " SELECT DISTINCT ALPHABETIC_CODE as currency" & _
							 	" FROM Currency " & _
							    " WHERE isactive = 'Y'"
							' response.write q
							Set rsComboCur = Server.CreateObject("ADODB.Recordset")
							rsComboCur.CursorType = 3
							rsComboCur.CursorLocation = 3
							rsComboCur.Open q, connstr
							if not rsComboCur.eof then%>
							<%
							Do While Not rsComboCur.eof
							%>
								<option value="<%=rsComboCur("currency")%>" 
									<%if cstr(""&CurrencyCode)=cstr(""&rsComboCur("currency")) then %> selected 
									<%End If%>><%=rsComboCur("currency")%>
								</option> 
							<%rsComboCur.movenext
							loop
							End if
							rsComboCur.Close
							set rsComboCur = Nothing
							q=""
						%>
				</select> 
					<font color="#FF0000">*</font>
				</td>
			<td align="left">Exchange Rate</td>
				<td align="left">
					<input type="text" maxlength="50" name="ExchangeRate" size="15" value="<%=ExchangeRate%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '1'">
				</td>
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
		
		<tr><td colspan="4"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		
		<%if len(trim(businessunit)) = "KKP" OR len(trim(businessunit)) = "KPK" then%>
		<a href="printingpo.Asp?pocode=<%=fcCode%>" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp; 
		<%end if%>
		<% if RowStateId <> "Planned" then%>
			<a href="printingpo.asp?pocode=<%=fcCode%>&fcba=<%=businessunit%>&purchaseby=<%=purchaseby%>&deliverydestination=<%=deliverydestination%>" target="_blank">
				<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
			</a>
		<%end if%>
		<% if RowStateId <> "AuthorizeOnProgress" then%>
		<a href="PurchaseOrderDetails.Asp?fcCode=<%=fcCode%>&PODATE=<%=PODATE%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<%end if%>
		<% if RowStateId = "Planned" And cint(IsInBridge) = 1 then %>
		    <a href="PurchaseOrderDetails.Asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=Ordered" onClick="return confirm('Are you sure to Order this data :  <%=fcCode%> ?');"> Ordered </a>
		<% elseif RowStateId = "Planned" And cint(IsInBridge) = 0 then %>
		    <a href="PurchaseOrderDetails.Asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to Confirm this data :  <%=fcCode%> ?');"> Confirmed </a>
		<% elseif RowStateId = "Confirmed" then %>	
		    <a href="PurchaseOrderDetails.Asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=fcCode%> ?');"> Re-Planned </a>
			&nbsp;&nbsp;
			<%if can_startauthorize = 0 then%>
				<font color="#FF0000">You Have No Permission to Start Authorize</font>
			<%else%>
			<a href="PurchaseOrderDetails.Asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=StartAuthorize" onClick="return confirm('Are you sure to start authorize this data :  <%=fcCode%> ?');">Start Authorize </a>
			<%end if%>
		<% elseif RowStateId = "AuthorizeOnProgress" then %>	
		    <%if is_final <> 0 then%>
			<a href="PurchaseOrderDetails.Asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=fcCode%> ?');"> Re-Planned </a>
			<%end if%>
			&nbsp;&nbsp;
		<% elseif  RowStateId = "Ordered" And cint(IsInBridge) = 1 And Cint(IsFrezee) =0 then %>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="PurchaseOrderDetails.Asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=UnRelease" onClick="return confirm('Are you sure to Un Release this data :  <%=fcCode%> ?');"> UnRelease </a>
		<% end if%>
		<%if RowStateId <> "" then %>
		<a href="PurchaseOrderHeaderEdit.asp?fcCode=<%=fcCode%>&businessunit=<%=businessunit%>" target="_blank">
			Edit Header
		</a>
		<% end if%>
		
		
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" AND LEN(TRIM(fcCode)) > 0 AND LEN(TRIM(fccompanycode)) > 0 AND LEN(TRIM(businessunit)) > 0  Then
	Dim rsPurchaseOrder,intRecordsPerPage,cColor

	q = " select *  from purchaseorderdetails pod" & _
		" inner join stock st on (st.fccode = pod.itemcode and st.fcba = pod.fcba)" & _
		" where pod.PoCode = '" & fcCode & "'" & _
	    "       and pod.fcBA = '" & businessunit & "'"
	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.CursorType = 3
	rsPurchaseOrder.CursorLocation = 3
	rsPurchaseOrder.Open q, connstr

	If rsPurchaseOrder.eof then
		otcstock1 = 0
		otcstock2 = 0
		otcstock3 = 0
		otcstock4 = 0
	End If 

	If Request("txtJenis")= "Edit" Then
		q = " select prd.itemcode, decode(prd.description, null, s.fcname, s.fcname||' - '||prd.description) description, prd.approvedqty, " & _
			" (prd.approvedqty - prd.orderedqty) remainsqty from purchaserequestdetails prd " & _
			" inner join purchaserequest pr ON (pr.prcode = prd.prcode and pr.fcba = prd.fcba)" & _
			" left join stock s ON (s.fccode = prd.itemcode and s.fcba = prd.fcba)" & _
			" where pr.prcode ='"& prcode &"' and prd.lineid = '"& prlineid &"'"

		' response.Write q

		set rsrecurtrans = server.CreateObject("Adodb.recordset")
		rsrecurtrans.Open q, connstr

		if Not rsrecurtrans.eof then
			remainsqty = rsrecurtrans("remainsqty")
		end if
		rsrecurtrans.Close
		set rsrecurtrans = Nothing
	End If
%>
   


   
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="PurchaseOrderDetails.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
   		<tr>
			<td align="left">PR. Number.</td>
			<td align="left">
			<%IF txtJenis="Edit" Then%>
				 <input type="text" size="15" name="PRCODE" value="<%=PRCODE%>" readonly="True" style="background:#eee">&nbsp;<font color="#FF0000">*
			<%else%>
				<select name="PRCODE" value="<%=PRCODE%>" id="prcode" onchange="javascript:getprline();" class="chosen-select" style="width: 40%">
					<option value=""></option> 
						<% 	q = " select distinct pr.prcode from (select prd.PrCode from purchaserequestdetails prd" & _
								" inner join purchaserequest pr on (pr.prcode=prd.prcode and pr.fcba=prd.fcba) " & _
								" inner join businessunit bu ON(bu.fccode = prd.fcba)" & _
								" inner join company co  ON(co.companycode = bu.fccompanycode)" & _
								" where pr.fcba = '"& businessunit &"' and prd.orderedqty < prd.approvedqty AND to_char(pr.PRDATE,'yyyy')>='2023'" & _
								" and pr.rowstateid = 'Approved' ) pr order by pr.prcode DESC"
								' response.write q
								Set rsComboP = Server.CreateObject("ADODB.Recordset")
								rsComboP.CursorType = 3
								rsComboP.CursorLocation = 3
								rsComboP.Open q, connstr
								if not rsComboP.eof then%>
								<%
								Do While Not rsComboP.eof
								%>
									<option value="<%=rsComboP("PrCode")%>" 
										<%if cstr(""&prcode)=cstr(""&rsComboP("PrCode")) then %> selected 
										<%End If%>><%=rsComboP("PrCode")%> 
									</option> 
								<%rsComboP.movenext
								loop
								End if
								rsComboP.Close
								set rsComboP = Nothing
								q=""
						%>
				</select> 
				<font color="#FF0000">*</font>
			<%End if%>
				
			</td>
		</tr>
		<tr>
			<td align="left">PR. Line Number.</td>
			<td align="left">
			<%IF txtJenis="Edit" Then%>
				 <input type="text" size="15" name="prlineid" value="<%=prlineid%>" readonly="True" style="background:#eee">&nbsp;<font color="#FF0000">*
			<%else%>
				<div id="prlinelist"></div>
			<%end if%>
			</td>
		</tr>
		<tr>
		   <td  align="left">Stock Code</td>
		   <td align="left">
				 <input type="text" size="20" id="itemcode" name="ITEMCODE" value="<%=ITEMCODE%>" style="background:#eee" readonly >
			</td>
			</tr>
		<tr>
			<td align="left">Stock Name</td>
			<td align="left"> 
				<input type="text" size="50" id="stockname" name="StockName" value="<%=StockName%>" style="background:#eee" readonly >
			</td>
		</tr>
		<tr>
			<td align="left">Request Quantity</td>
			<td align="left"><input type="text" size="15" id="requestedqty" name="requestedqty" value="<%=requestedqty%>" style="background:#eee" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Remain Quantity</td>
			<td align="left"><input type="text" size="15" id="remainsqty" name="remainsqty" value="<%=remainsqty%>" style="background:#eee" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Ordered Quantity</td>
			<td align="left"><input type="text" size="15" id="quantity" name="quantity" value="<%=quantity%>" style="text-align: right;" id="aninput" onkeypress="return validateFloatKeyPress(this,event);" autocomplete="off"  onkeyup="hitungrumus();" >&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		<tr>
			<td align="left">Unit Price</td>
			<td align="left"><input type="text" size="15" id="unitprice" name="unitprice" value="<%=unitprice%>" onkeypress="return harusangka(event);" onkeyup="hitungrumus();" style = "text-align: right;">&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		<tr>
			<td align="left">Amount</td>
			<td align="left"><input type="text" size="15" id="amount1" name="amount1" value="<%=amount1%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right; background:#eee;" >
			</td>
		</tr>
		<tr>
			<td align="left">Disc (%)</td>
			<td align="left"><input type="text" size="15" name="disc" value="<%=disc%>" onkeypress="return harusangka(event);" autocomplete="off">&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		<tr>
			<td align="left">PBBKB (7.5%)</td>
			<td align="left">
				<input type="text" size="15" onkeypress="return harusangka(event);" onkeyup="hitungrumus();" id="otcstock1" name="otcstock1" value="<%=otcstock1%>" style="text-align: right;" autocomplete="off">
			</td>	
		</tr>
		<tr>
			<td align="left">PBBKB (10%)</td>
			<td align="left">
				<input type="text" size="15" onkeypress="return harusangka(event);" onkeyup="hitungrumus();" id="otcstock2" name="otcstock2" value="<%=otcstock2%>" style="text-align: right;" autocomplete="off">
			</td>	
		</tr>
		<tr>
			<td align="left">Transportasi</td>
			<td align="left">
				<input type="text" size="15" onkeypress="return harusangka(event);" onkeyup="hitungrumus();" id="otcstock3" name="otcstock3" value="<%=otcstock3%>" style="text-align: right;" autocomplete="off">
			</td>	
		</tr>
		<tr>
			<td align="left">Others</td>
			<td align="left">
				<input type="text" size="15" onkeypress="return harusangka(event);" onkeyup="hitungrumus();" id="otcstock4" name="otcstock4" value="<%=otcstock4%>" style="text-align: right;" autocomplete="off">
			</td>	
		</tr>
		<tr>
			<td align="left">Other Charge Stock</td>
				<td align="left"><input type="text" size="15" id="otherchargestock" name="otherchargestock" value="<%=otherchargestock%>" style="text-align: right; background-color: #eee;" readonly autocomplete="off">
				</td>	
		</tr>
		<tr>
			<td align="left">Other Charge Non Stock</td>
				<td align="left"><input type="text" size="15" name="othercharge" value="<%=othercharge%>" style="text-align: right;" autocomplete="off">
				</td>	
		</tr>
		<tr>
			<td align="left">Amount Charge</td>
			<td align="left"><input type="text" size="15" id="amountcharge" name="AMOUNTCHARGE" value="<%=AMOUNTCHARGE%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readony style ="background-color: #eee;" >
			</td>
		</tr>
		<tr>
			<td align="left">Received</td>
			<td align="left"><input type="text" size="15" name="received" value="<%=received%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee;" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Description</td>
			<td align="left">
				<textarea name="OTHERITEMDESC" cols="50"><%=OTHERITEMDESC%></textarea>
			</td>
		</tr>
			
		
		
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="PODATE" value="<%=PODATE%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="LineNokey" value="<%=LINENOKEY%>">
		  <input type="hidden" name="use_ppn" value="<%=use_ppn%>">
		  <input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		  <input type="hidden" name="businessunit" value="<%=businessunit%>">
		  <input type="hidden" name="purchaseby" value="<%=purchaseby%>">
		  <input type="hidden" name="supplieraddress" value="<%=supplieraddress%>">
		  <input type="hidden" name="deliverydestination" value="<%=deliverydestination%>">
		  <input type="hidden" name="termofpayment" value="<%=termofpayment%>">
		  <input type="hidden" name="SHIPPINGPOINT" value="<%=SHIPPINGPOINT%>">
		  <input type="hidden" name="dp" value="<%=dp%>">
		  <input type="hidden" name="ppn" value="<%=ppn%>">
		  <input type="hidden" name="totalamount" value="<%=totalamount%>">
		  <input type="hidden" name="totalamountppn" value="<%=totalamountppn%>">
		  <input type="hidden" name="totalcharge" value="<%=totalcharge%>">
		  <input type="hidden" name="grandtotal" value="<%=grandtotal%>">
		  <input type="hidden" name="totalppn" value="<%=totalppn%>">
		  <input type="hidden" name="document_classification" value="<%=document_classification%>">
		  <input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		  <input type="hidden" name="selectedpurchaseby" value="<%=selectedpurchaseby%>">
		  <input type="hidden" name="selectedsupplier" value="<%=selectedsupplier%>">
		  <input type="hidden" name="selecteddeliverydestination" value="<%=selecteddeliverydestination%>">

		  <input type="hidden" name="RowStatehdr" value="<%=RowStatehdr%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <% If HdrState="Planned" or len(trim(HdrState))=0  Then %>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else  %>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" name="Add" value="Add">
		  <%End if%>
		  <%End if%>
		  <input Type="Reset" name="Reset" value="Reset">
		  <input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="middle" colspan="2"> ACTION</TD>
	<TD align="left">PO CODE</TD>
	<TD align="left">LINENOKEY</TD>
	<TD align="left">ITEM CODE</TD>
	<TD align="left">QUANTITY</TD>
	<TD align="left">UNIT PRICE</TD>
	<!-- Add Aam Jan 21 -->
	<TD align="left">ITEM AMOUNT</TD>
	<TD align="left">DISC</TD>
	<TD align="left">PBBKB 7.5%</TD>
	<TD align="left">PBBKB 10%</TD>
	<TD align="left">TRANSPORTASI</TD>
	<TD align="left">OTHER CHARGE </TD>
	<TD align="left">OTHER CHARGE STOCK</TD>
	<TD align="left">OTHER CHARGE NON STOCK</TD>
	<TD align="left">TOT.AMOUNT</TD>
	<TD align="left">OTHER ITEM DESC</TD>
	<TD align="left">RECEIVED</TD>
	<TD align="left">PRCODE</TD>
	<TD align="left">PRLINEID</TD>
	<TD align="left">LINE STATUS</TD>
	<TD align="left">FREZEE</TD>
	</tr>
	</thead>
	<tbody>
   <%
   
     Set rsPO = Server.CreateObject("ADODB.Recordset")
	 rsPO.Open " select purchaseorder_api.GetPoState('" & businessunit & "' ,'" & fcCode & "') HdrState from dual" , connstr
	 IF Not rsPO.eof then
		HdrState =  rsPO("HdrState")
	 End if
	 rsPO.Close
	 set rsPO = Nothing
   
	  if Not rsPurchaseOrder.eof then
      nCountofField=rsPurchaseOrder.fields.count-1
	  no=1
	  While Not rsPurchaseOrder.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<% If HdrState="Planned" Then %>			
			    <td valign="middle" width="10">
				   <a href="PurchaseOrderDetails.Asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("ItemCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&Amountcharge=<%=rsPurchaseOrder("Amountcharge")%>&otherchargestock=<%=rsPurchaseOrder("otherchargestock")%>&OTHERITEMDESC=<%=rsPurchaseOrder("OTHERITEMDESC")%>&PRCODE=<%=rsPurchaseOrder("PRCODE")%>&PRLINEID=<%=rsPurchaseOrder("PRLINEID")%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&stockname=<%=rsPurchaseOrder("stockname")%>&unitprice=<%=rsPurchaseOrder("unitprice")%>&disc=<%=rsPurchaseOrder("disc")%>&othercharge=<%=rsPurchaseOrder("othercharge")%>&otcstock1=<%=rsPurchaseOrder("otcstock1")%>&otcstock2=<%=rsPurchaseOrder("otcstock2")%>&otcstock3=<%=rsPurchaseOrder("otcstock3")%>&otcstock4=<%=rsPurchaseOrder("otcstock4")%>&requestedqty=<%=rsPurchaseOrder("requestedqty")%>&txtJenis=Edit">
				   <img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(1)%>" border="0"></a>
				</td>
				<td valign="middle" width="10">
				   <a href="PurchaseOrderDetails.Asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("ItemCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&Amountcharge=<%=rsPurchaseOrder("Amountcharge")%>&otherchargestock=<%=rsPurchaseOrder("otherchargestock")%>&OTHERITEMDESC=<%=rsPurchaseOrder("OTHERITEMDESC")%>&PRCODE=<%=rsPurchaseOrder("PRCODE")%>&PRLINEID=<%=rsPurchaseOrder("PRLINEID")%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&stockname=<%=rsPurchaseOrder("stockname")%>&unitprice=<%=rsPurchaseOrder("unitprice")%>&disc=<%=rsPurchaseOrder("disc")%>&othercharge=<%=rsPurchaseOrder("othercharge")%>&requestedqty=<%=rsPurchaseOrder("requestedqty")%>&txtJenis=Details">
				   <img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');">
				   </a>
				</td>
			<% Else %>	
               	<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>		
            <% End If %>	
				
				<td align="left" title="<%=rsPurchaseOrder("pocode").Name%>" valign="top">
					<%=rsPurchaseOrder("pocode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("linenokey").Name%>" valign="top">
					<%=rsPurchaseOrder("linenokey")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("itemcode").Name%>" valign="top">
					<%=rsPurchaseOrder("itemcode")%> - <%=rsPurchaseOrder("fcname")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("quantity").Name%>" valign="top">
					<%=rsPurchaseOrder("quantity")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("unitprice").Name%>" valign="top">
					<%=rsPurchaseOrder("unitprice")%>
				</td>
					<!-- Add Aam Jan 21 -->
				<td align="left" title="<%=rsPurchaseOrder("amount").Name%>" valign="top">
					<%=rsPurchaseOrder("amount")%>
				</td>
	
				<td align="left" title="<%=rsPurchaseOrder("disc").Name%>" valign="top">
					<%=rsPurchaseOrder("disc")%> %
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otcstock1").Name%>" valign="top">
					<%=rsPurchaseOrder("otcstock1")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otcstock2").Name%>" valign="top">
					<%=rsPurchaseOrder("otcstock2")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otcstock3").Name%>" valign="top">
					<%=rsPurchaseOrder("otcstock3")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otcstock4").Name%>" valign="top">
					<%=rsPurchaseOrder("otcstock4")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otherchargestock").Name%>" valign="top">
					<%=rsPurchaseOrder("otherchargestock")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("othercharge").Name%>" valign="top">
					<%=rsPurchaseOrder("othercharge")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("amountcharge").Name%>" valign="top">
					<%=formatnumber(rsPurchaseOrder("amountcharge"),0)%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otheritemdesc").Name%>" valign="top">
					<%=rsPurchaseOrder("otheritemdesc")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("received").Name%>" valign="top">
					<%=rsPurchaseOrder("received")%>
				</td>
				
				<td align="left" title="<%=rsPurchaseOrder("prcode").Name%>" valign="top">
					<%=rsPurchaseOrder("prcode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("prlineid").Name%>" valign="top">
					<%=rsPurchaseOrder("prlineid")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("rowstateid").Name%>" valign="top">
					<%=rsPurchaseOrder("rowstateid")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("FREZEE").Name%>" valign="top">
					<%=rsPurchaseOrder("FREZEE")%>
				</td>
			</tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("<tr><td colspan=""22"" class=""fbp-doc-empty"">Data is empty</td></tr>")
   end if
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </tbody>
   </TABLE>
   </div>
<% if Request("oldui") <> "1" then %>
<script type="text/javascript">
// Status baris jadi badge (tampilan saja)
(function () {
	var cells = document.querySelectorAll('table.fbp-doc-lines tbody td[title]');
	for (var i = 0; i < cells.length; i++) {
		if (cells[i].getAttribute('title').toUpperCase() !== 'ROWSTATEID') continue;
		var s = (cells[i].textContent || '').replace(/\s+/g, ' ').trim();
		if (!s) continue;
		cells[i].parentNode.setAttribute('data-state', s);
		cells[i].innerHTML = '';
		var b = document.createElement('span');
		b.className = 'fbp-state';
		b.textContent = s;
		cells[i].appendChild(b);
	}
})();
</script>
<% end if %>
   </form>
<%End if%>
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
				cb.value = "Y";
			} else {
				cb.value = "N";
			}
        }
	</script>

<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
}

function getselectedpurchaseby() {
	var scustomercode = document.getElementById("selectedpurchaseby");
	scustomercode.value = document.getElementById("purchaseby").value;
}

function getselectedsupplier() {
	var ssupplier = document.getElementById("selectedsupplier");
	ssupplier.value = document.getElementById("suppliercode").value;
}

function getdeliverydestination() {
	var sdeliverydestination = document.getElementById("selecteddeliverydestination");
	sdeliverydestination.value = document.getElementById("deliverydestination").value;
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
	getPurchaseBy();
	getSupplierPO();
	getdeliverydest();
	getSupp();
	getprline();
	getStockCode();
	getselectedfcba();
	getselectedpurchaseby();
	getdeliverydestination();
	getselectedsupplier();
	hitungrumus(); 
});
</script>
		
</body>
</html>
