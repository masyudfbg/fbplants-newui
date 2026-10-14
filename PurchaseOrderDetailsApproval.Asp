<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="email/po_mail_queue.inc"-->
<html>
<head>
<title>ePlaS - Purchase Order Approval</title>
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
var getSupp = function () {
	// build up the post string when passing variables to the server side page
	var suppliercode = document.getElementById("suppliercode").value;
	var purchaseby = document.getElementById("purchaseby").value;
	var PostStr = 'suppliercode='+suppliercode+'&purchaseby='+purchaseby;
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
    document.getElementById("suppliername").value = s.suppliername;
	document.getElementById("supplieraddress").value = s.supplieraddress;
	// document.getElementById("address").innerHTML = document.getElementById("supplieraddress").value;
	// alert(s.suppliername);
    
}

// launched from button click 
var getStockCode = function () {
	// alert("1");
	// build up the post string when passing variables to the server side page
	// var PrCode = document.getElementById("prcode").value;
	// var PRLINEID = document.getElementById("prlineid").value;
	var PostStr = 'PrCode='+document.getElementById("prcode").value+'&PRLINEID='+document.getElementById("prlineid").value;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getStockCode.asp', 'POST', PostStr, showStockCodeResponse);
	
}

// The function for handling the response from the server
var showStockCodeResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.itemcode);
    // update the Div to show the result from the server
    document.getElementById("itemcode").value = s.itemcode;
	document.getElementById("stockname").value = s.stockname;
	document.getElementById("requestedqty").value = s.approvedqty;
	
    
};

</script>
<script language="JavaScript">
function getFcba(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("businessunit").innerHTML="";
	  document.getElementById("purchaseby").innerHTML="";
	  document.getElementById("deliverydestination").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {// code for IE7+, Firefox, Chrome, Opera, Safari
	  xmlhttp=new XMLHttpRequest();
	  StockName=new XMLHttpRequest();
	  }
	else
	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{
		// alert(xmlhttp.responseText);
		document.getElementById("businessunit").innerHTML=xmlhttp.responseText;
		document.getElementById("purchaseby").innerHTML=xmlhttp.responseText;
		document.getElementById("deliverydestination").innerHTML=xmlhttp.responseText;
		// alert(businessunit);
		}
	  }
	xmlhttp.open("GET","ajax/getFcbaPO.asp?fccompanycode="+str+"",true);
	xmlhttp.send();
}

function getSupplierPO(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("purchaseby").innerHTML="";
	  // document.getElementById("purchaseby").innerHTML="";
	  // document.getElementById("receviedtask").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {// code for IE7+, Firefox, Chrome, Opera, Safari
	  xmlhttp=new XMLHttpRequest();
	  StockName=new XMLHttpRequest();
	  }
	else
	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{
		// alert(xmlhttp.responseText);
		document.getElementById("suppliercode").innerHTML=xmlhttp.responseText;
		// document.getElementById("purchaseby").innerHTML=xmlhttp.responseText;
		// document.getElementById("receviedtask").innerHTML="";
		// alert(businessunit);
		}
	  }
	// xmlhttp.open("GET","ajax/getSupplierPO.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
	xmlhttp.open("GET","ajax/getSupplierPO.asp?fcba="+str+"",true);
	xmlhttp.send();
}

// function getIncoTerm(str){

// 	var xmlhttp;
// 	if (str=="")
// 	  {
// 	  document.getElementById("SHIPPINGPOINT").innerHTML="";
// 	  // document.getElementById("purchaseby").innerHTML="";
// 	  // document.getElementById("receviedtask").innerHTML="";
// 	  return;
// 	  }
// 	if (window.XMLHttpRequest)
// 	  {// code for IE7+, Firefox, Chrome, Opera, Safari
// 	  xmlhttp=new XMLHttpRequest();
// 	  StockName=new XMLHttpRequest();
// 	  }
// 	else
// 	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
// 	xmlhttp.onreadystatechange=function()
// 	  {
// 	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
// 		{
// 		alert(xmlhttp.responseText);
// 		document.getElementById("deliverydestination").innerHTML=xmlhttp.responseText;
// 		// document.getElementById("purchaseby").innerHTML=xmlhttp.responseText;
// 		// document.getElementById("receviedtask").innerHTML="";
// 		// alert(businessunit);
// 		}
// 	  }
// 	// xmlhttp.open("GET","ajax/getSupplierPO.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
// 	xmlhttp.open("GET","SHIPPINGPOINT="+str+"",true);
// 	xmlhttp.send();
// }

function getPRLine(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("prcode").innerHTML="";
	  document.getElementById("businessunit").innerHTML="";
	  // document.getElementById("receviedtask").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {// code for IE7+, Firefox, Chrome, Opera, Safari
	  xmlhttp=new XMLHttpRequest();
	  StockName=new XMLHttpRequest();
	  }
	else
	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{
		// alert(xmlhttp.responseText);
		document.getElementById("prlineid").innerHTML=xmlhttp.responseText;
		document.getElementById("requestedqty").innerHTML=xmlhttp.responseText;
		// document.getElementById("receviedtask").innerHTML="";
		// alert(PRLINEID);
		}
	  }
	// xmlhttp.open("GET","ajax/getSupplierPO.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
	xmlhttp.open("GET","ajax/getPRLine.asp?prcode="+str+"&businessunit="+document.getElementById("businessunit").value+"",true);
	xmlhttp.send();
}
					
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<style>
.button {
    position: relative;
    overflow: visible;
    display: inline-block;
    padding: 0.5em 1em;
    border: 1px solid #d4d4d4;
    margin: 0;
    text-decoration: none;
    text-align: center;
    text-shadow: 1px 1px 0 #fff;
    font:11px/normal sans-serif;
    color: #333;
    white-space: nowrap;
    cursor: pointer;
    outline: none;
    background-color: #ececec;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#f4f4f4), to(#ececec));
    background-image: -moz-linear-gradient(#f4f4f4, #ececec);
    background-image: -ms-linear-gradient(#f4f4f4, #ececec);
    background-image: -o-linear-gradient(#f4f4f4, #ececec);
    background-image: linear-gradient(#f4f4f4, #ececec);
    -moz-background-clip: padding; /* for Firefox 3.6 */
    background-clip: padding-box;
    border-radius: 0.2em;
    /* IE hacks */
    zoom: 1;
    *display: inline;
}

.button:hover,
.button:focus,
.button:active,
.button.active {
    border-color: #3072b3;
    border-bottom-color: #2a65a0;
    text-decoration: none;
    text-shadow: -1px -1px 0 rgba(0,0,0,0.3);
    color: #fff;
    background-color: #3c8dde;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#599bdc), to(#3072b3));
    background-image: -moz-linear-gradient(#599bdc, #3072b3);
    background-image: -o-linear-gradient(#599bdc, #3072b3);
    background-image: linear-gradient(#599bdc, #3072b3);
}

.button:active,
.button.active {
    border-color: #2a65a0;
    border-bottom-color: #3884cd;
    background-color: #3072b3;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#3072b3), to(#599bdc));
    background-image: -moz-linear-gradient(#3072b3, #599bdc);
    background-image: -ms-linear-gradient(#3072b3, #599bdc);
    background-image: -o-linear-gradient(#3072b3, #599bdc);
    background-image: linear-gradient(#3072b3, #599bdc);
}

/* overrides extra padding on button elements in Firefox */
.button::-moz-focus-inner {
    padding: 0;
    border: 0;
}


/* =============================================================================
   Button icons
   ========================================================================== */

.button.icon:before {
    content: "";
    position: relative;
    top: 1px;
    float:left;
    width: 12px;
    height: 12px;
    margin: 0 0.75em 0 -0.25em;
    background: url(gh-icons.png) 0 99px no-repeat;
}

.button.arrowup.icon:before { background-position: 0 0; }
.button.arrowup.icon:hover:before,
.button.arrowup.icon:focus:before,
.button.arrowup.icon:active:before { background-position: -12px 0; }

.button.arrowdown.icon:before { background-position: 0 -12px; }
.button.arrowdown.icon:hover:before,
.button.arrowdown.icon:focus:before,
.button.arrowdown.icon:active:before { background-position: -12px -12px; }

.button.arrowleft.icon:before { background-position: 0 -24px; }
.button.arrowleft.icon:hover:before,
.button.arrowleft.icon:focus:before,
.button.arrowleft.icon:active:before { background-position: -12px -24px; }

.button.arrowright.icon:before { float:right; margin: 0 -0.25em 0 0.5em; background-position: 0 -36px; }
.button.arrowright.icon:hover:before,
.button.arrowright.icon:focus:before,
.button.arrowright.icon:active:before { background-position: -12px -36px; }

.button.approve.icon:before { background-position: 0 -48px; }
.button.approve.icon:hover:before,
.button.approve.icon:focus:before,
.button.approve.icon:active:before { background-position: -12px -48px; }

.button.add.icon:before { background-position: 0 -288px; }
.button.add.icon:hover:before,
.button.add.icon:focus:before,
.button.add.icon:active:before { background-position: -12px -288px; }

.button.remove.icon:before { background-position: 0 -60px; }
.button.remove.icon:hover:before,
.button.remove.icon:focus:before,
.button.remove.icon:active:before { background-position: -12px -60px; }

.button.log.icon:before { background-position: 0 -72px; }
.button.log.icon:hover:before,
.button.log.icon:focus:before,
.button.log.icon:active:before { background-position: -12px -72px; }

.button.calendar.icon:before { background-position: 0 -84px; }
.button.calendar.icon:hover:before,
.button.calendar.icon:focus:before,
.button.calendar.icon:active:before { background-position: -12px -84px; }

.button.chat.icon:before { background-position: 0 -96px; }
.button.chat.icon:hover:before,
.button.chat.icon:focus:before,
.button.chat.icon:active:before { background-position: -12px -96px; }

.button.clock.icon:before { background-position: 0 -108px; }
.button.clock.icon:hover:before,
.button.clock.icon:focus:before,
.button.clock.icon:active:before { background-position: -12px -108px; }

.button.settings.icon:before { background-position: 0 -120px; }
.button.settings.icon:hover:before,
.button.settings.icon:focus:before,
.button.settings.icon:active:before { background-position: -12px -120px; }

.button.comment.icon:before { background-position: 0 -132px; }
.button.comment.icon:hover:before,
.button.comment.icon:focus:before,
.button.comment.icon:active:before { background-position: -12px -132px; }

.button.fork.icon:before { background-position: 0 -144px; }
.button.fork.icon:hover:before,
.button.fork.icon:focus:before,
.button.fork.icon:active:before { background-position: -12px -144px; }

.button.like.icon:before { background-position: 0 -156px; }
.button.like.icon:hover:before,
.button.like.icon:focus:before,
.button.like.icon:active:before { background-position: -12px -156px; }

.button.favorite.icon:before { background-position: 0 -348px; }
.button.favorite.icon:hover:before,
.button.favorite.icon:focus:before,
.button.favorite.icon:active:before { background-position: -12px -348px; }

.button.home.icon:before { background-position: 0 -168px; }
.button.home.icon:hover:before,
.button.home.icon:focus:before,
.button.home.icon:active:before { background-position: -12px -168px; }

.button.key.icon:before { background-position: 0 -180px; }
.button.key.icon:hover:before,
.button.key.icon:focus:before,
.button.key.icon:active:before { background-position: -12px -180px; }

.button.lock.icon:before { background-position: 0 -192px; }
.button.lock.icon:hover:before,
.button.lock.icon:focus:before,
.button.lock.icon:active:before { background-position: -12px -192px; }

.button.unlock.icon:before { background-position: 0 -204px; }
.button.unlock.icon:hover:before,
.button.unlock.icon:focus:before,
.button.unlock.icon:active:before { background-position: -12px -204px; }

.button.loop.icon:before { background-position: 0 -216px; }
.button.loop.icon:hover:before,
.button.loop.icon:focus:before,
.button.loop.icon:active:before { background-position: -12px -216px; }

.button.search.icon:before { background-position: 0 -228px; }
.button.search.icon:hover:before,
.button.search.icon:focus:before,
.button.search.icon:active:before { background-position: -12px -228px; }

.button.mail.icon:before { background-position: 0 -240px; }
.button.mail.icon:hover:before,
.button.mail.icon:focus:before,
.button.mail.icon:active:before { background-position: -12px -240px; }

.button.move.icon:before { background-position: 0 -252px; }
.button.move.icon:hover:before,
.button.move.icon:focus:before,
.button.move.icon:active:before { background-position: -12px -252px; }

.button.edit.icon:before { background-position: 0 -264px; }
.button.edit.icon:hover:before,
.button.edit.icon:focus:before,
.button.edit.icon:active:before { background-position: -12px -264px; }

.button.pin.icon:before { background-position: 0 -276px; }
.button.pin.icon:hover:before,
.button.pin.icon:focus:before,
.button.pin.icon:active:before { background-position: -12px -276px; }

.button.reload.icon:before { background-position: 0 -300px; }
.button.reload.icon:hover:before,
.button.reload.icon:focus:before,
.button.reload.icon:active:before { background-position: -12px -300px; }

.button.rss.icon:before { background-position: 0 -312px; }
.button.rss.icon:hover:before,
.button.rss.icon:focus:before,
.button.rss.icon:active:before { background-position: -12px -312px; }

.button.tag.icon:before { background-position: 0 -324px; }
.button.tag.icon:hover:before,
.button.tag.icon:focus:before,
.button.tag.icon:active:before { background-position: -12px -324px; }

.button.trash.icon:before { background-position: 0 -336px; }
.button.trash.icon:hover:before,
.button.trash.icon:focus:before,
.button.trash.icon:active:before { background-position: -12px -336px; }

.button.user.icon:before { background-position: 0 -360px; }
.button.user.icon:hover:before,
.button.user.icon:focus:before,
.button.user.icon:active:before { background-position: -12px -360px; }


/* =============================================================================
   Button extensions
   ========================================================================== */

/* Primary button
   ========================================================================== */

.button.primary {
    font-weight: bold;
}

/* Danger button
   ========================================================================== */

.button.danger {
    color: #900;
}

.button.danger:hover,
.button.danger:focus,
.button.danger:active {
    border-color: #b53f3a;
    border-bottom-color: #a0302a;
    color: #fff;
    background-color: #dc5f59;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#dc5f59), to(#b33630));
    background-image: -moz-linear-gradient(#dc5f59, #b33630);
    background-image: -ms-linear-gradient(#dc5f59, #b33630);
    background-image: -o-linear-gradient(#dc5f59, #b33630);
    background-image: linear-gradient(#dc5f59, #b33630);
}

.button.danger:active,
.button.danger.active {
    border-color: #a0302a;
    border-bottom-color: #bf4843;
    background-color: #b33630;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#b33630), to(#dc5f59));
    background-image: -moz-linear-gradient(#b33630, #dc5f59);
    background-image: -ms-linear-gradient(#b33630, #dc5f59);
    background-image: -o-linear-gradient(#b33630, #dc5f59);
    background-image: linear-gradient(#b33630, #dc5f59);
}

/* Pill button
   ========================================================================== */

.button.pill {
    border-radius: 50em;
}

/* Disabled button
   ========================================================================== */

.button.disable {
    opacity: 0.5;
}

/* Big button
   ========================================================================== */

.button.big {
    font-size: 14px;
}

.button.big.icon:before {
    top: 0;
}


/* =============================================================================
   Button groups
   ========================================================================== */

/* Standard group
   ========================================================================== */

.button-group {
    display: inline-block;
    list-style: none;
    padding: 0;
    margin: 0;
    /* IE hacks */
    zoom: 1;
    *display: inline;
}

.button + .button,
.button + .button-group,
.button-group + .button,
.button-group + .button-group {
    margin-left: 5px;
}

.button-group li {
    float: left;
    padding: 0;
    margin: 0;
}

.button-group .button {
    float: left;
    margin-left: -1px;
}

.button-group > .button:not(:first-child):not(:last-child),
.button-group li:not(:first-child):not(:last-child) .button {
    border-radius: 0;
}

.button-group > .button:first-child,
.button-group li:first-child .button {
    margin-left: 0;
    border-top-right-radius: 0;
    border-bottom-right-radius: 0;
}

.button-group > .button:last-child,
.button-group li:last-child > .button {
    border-top-left-radius: 0;
    border-bottom-left-radius: 0;
}

/* Minor group
   ========================================================================== */

.button-group.minor-group .button {
    border: 1px solid #d4d4d4;
    text-shadow: none;
    background-image: none;
    background-color: #fff;
}

.button-group.minor-group .button:hover,
.button-group.minor-group .button:focus {
    background-color: #599bdc;
}

.button-group.minor-group .button:active,
.button-group.minor-group .button.active {
    background-color: #3072b3;
}

.button-group.minor-group .button.icon:before {
    opacity: 0.8;
}

/* =============================================================================
   Button container (mixing buttons and groups, e.g., nav bar)
   ========================================================================== */

.button-container .button,
.button-container .button-group {
    vertical-align: top;
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
    fcCode = request("fccOde")
 ' end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("PODATE")<>"" then
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
 totalamount = Request("totalamount")
 totalamountppn = Request("totalamountppn")
 totalcharge = Request("totalcharge")
 grandtotal = Request("grandtotal")
 amountcharge = Request("amountcharge")
 totalppn = Request("totalppn")
 approvedqty = request ("approvedqty")
 comments = request ("comments")

if "" & use_ppn = "" then use_ppn = "N" 

fcba = request("fcba")
authorizelevelcode = request("authorizelevelcode")
authorizetemplatecode = request("authorizetemplatecode")
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "ApprovalAction" or Request("txtJenis")= "RejectAction" Then

	if Request("txtJenis")="ApprovalAction" then
		connstr.BeginTrans
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
			"          , '" & fcba & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
		' response.write q
		connStr.execute(q)
		
		q = " INSERT INTO po_authorization_comments ( pocode" & _
			"                                         , currentauthorizelevel" & _
			"                                         , comments" & _
			"                                         , fcentry" & _
			"                                         , fcedit" & _
			"                                         , fcip" & _
			"                                         , fcba" & _
			"                                         , lastupdate" & _
			"                                         , lasttime)"& _
			" VALUES ( '" & fcCode & "'" & _
			"          , '" & authorizelevelcode & "'" & _
			"          , '" & comments & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"          , '" & fcba & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
		' response.write q
		connStr.execute(q)
		
		is_final = 0
		q = " SELECT COUNT(1) total FROM authorizelevel al" & _
			" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
			" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
			" WHERE at.document_type = 'PO'" & _
			"       AND DECODE(at.document_type, 'PR', at.document_classification, '0') = '0'" & _
			"       AND al.isactive = 'TRUE'" & _
			"       AND al.finallevel = 'TRUE'" & _
			"       AND al.fccode = '" & authorizelevelcode & "'" & _
			"       AND aus.authorizetemplatecode = '" & authorizetemplatecode & "'" & _
			"       AND aus.fcba = '" & fcba & "'"
		'response.write q
		set rsPurchase = server.CreateObject("Adodb.recordset")
		rsPurchase.Open q, connstr
		IF not rsPurchase.eof then
			is_final = cdbl(to_nol(rsPurchase("total")))
		end if
		
		if is_final > 0 then
			set rsDone = server.CreateObject("Adodb.recordset")
			qDone = " select purchaseorder_api.getApproved('" & fcCode & "','" & fcba & "') tempDone from dual"
		   
			rsDone.Open qDone, connstr
			if Not rsDone.Eof then
				doneApproved = rsDone ("tempDone")
			End if
			rsDone.Close : Set rsDone = nothing
			
			' IF doneApproved = "TRUE" THEN 
			   strQueryStatus = " begin purchaseorder_api.setApproved('" & fcCode & "','" & fcba & "'); end;"
			   connStr.execute(strQueryStatus)
			' else
				' err.number = 1001
				' err.description = "PO "&fcCode&", business unit "&fcba&" has errors according to PR data!"
			' END IF

			strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & fcba & "','" & fccode & "', 'Approve' ); END; "
			connStr.execute(strQueryStatus)
			
			strQueryStatus = "UPDATE purchaseorderdetails SET rowstateid = 'Ordered' WHERE fcba = '" & fcba & "' AND pocode = '" & fcCode & "'"
			connStr.execute(strQueryStatus)
			
			strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & fcba & "','" & fcCode & "', 'Print' ); END; "
			connStr.execute(strQueryStatus)
			
			strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & fcba & "','" & fcCode & "', 'Order' ); END; "
			connStr.execute(strQueryStatus)
			
			sqlupdateqtypr = " select prcode, prlineid, quantity from purchaseorderdetails" & _
							 " where pocode = '" & fcCode & "'" & _
							 "       and fcba = '" & fcba & "'" & _
							 "       and prcode is not null" & _
							 "       and prlineid is not null"
							 
			Set rsupdateqtypr = Server.CreateObject("ADODB.Recordset")
			rsupdateqtypr.Open sqlupdateqtypr, connstr, adLockOptimistic, adOpenStatic
			
			if not rsupdateqtypr.EOF then
				While Not rsupdateqtypr.EOF
						prlineid = rsupdateqtypr ("prlineid")
						quantity = rsupdateqtypr ("quantity")
						prcode = rsupdateqtypr ("prcode")
					rsupdateqtypr.MoveNext
				wend
			end if
						
			rsupdateqtypr.Close
			set rsupdateqtypr = Nothing 
			
			q = " update purchaserequestdetails" & _
				" set orderedqty = '" & quantity & "'" & _
				" where  prcode = '" & prcode & "'" & _
				"        and lineid = '" & prlineid & "'" & _
				"        and fcba = '" & fcba & "'"
			' Response.write q
			connStr.execute(q)
			
			' Approver terakhir: antrikan email PO ke supplier (PDF + QR) di dalam transaksi yang sama,
			' sehingga ikut batal bila approval di-rollback. Lihat email/po_mail_queue.inc
			If Err.Number = 0 Then Call PoMail_Queue(fcba, fcCode)
		
		end if
		if err.number  = 0 then
			  connStr.CommitTrans
		end if
		' Kirim email PO yang baru diantrikan (cadangan: Task Scheduler email/run_po_outbox.vbs)
		If err.number = 0 Then Call PoMail_Kick()
		cmdView	= "True"
	elseif Request("txtJenis")="RejectAction" then
	connstr.BeginTrans
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
			"          , '" & fcba & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
		'response.write q
		connStr.execute(q)
		
		q = " INSERT INTO po_authorization_comments ( pocode" & _
			"                                         , currentauthorizelevel" & _
			"                                         , comments" & _
			"                                         , fcentry" & _
			"                                         , fcedit" & _
			"                                         , fcip" & _
			"                                         , fcba" & _
			"                                         , lastupdate" & _
			"                                         , lasttime)"& _
			" VALUES ( '" & fcCode & "'" & _
			"          , '" & authorizelevelcode & "'" & _
			"          , '" & comments & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"          , '" & fcba & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
		'response.write q
		connStr.execute(q)
		
		
		strQueryStatus = "UPDATE purchaseorderdetails SET rowstateid = 'Rejected' WHERE fcba = '" & fcba & "' AND pocode = '" & fcCode & "'"
		connStr.execute(strQueryStatus)
		
	    strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & fcba & "','" & fcCode & "', 'Reject' ); END; "
		connStr.execute(strQueryStatus)
		
		For n=1 to (cdbl(request("no"))-1)  
			q=   	" Update purchaserequestdetails set orderedqty=(SELECT SUM (pod.quantity) qty" & _
					"     FROM purchaseorderdetails pod" & _
					"          INNER JOIN purchaserequestdetails prd" & _
					"             ON (pod.prcode = prd.prcode AND pod.prlineid = prd.lineid)" & _
					"    WHERE pod.prcode = '"& request("prcode_"&TRIM(n)) &"' and prd.lineid = '"& request("prlineid_"&TRIM(n)) &"'" & _
					" GROUP BY prd.prcode, prd.lineid, pod.itemcode) " & _
					", LASTUPDATE = sysdate " & _
					", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					" where  prcode='" & request("prcode_"&TRIM(n)) & "' and ItemCode='" & request("itemcode_"&TRIM(n)) & "' and lineid='" & request("prlineid_"&TRIM(n)) & "'"
				connStr.execute(q)
				' response.write q
		Next
			
		if err.number  = 0 then
			connStr.CommitTrans
		end if	
		cmdView	= "True"
    else
        
	end if			
		
	cmdView	= "True"
 else
    cmdView	= request("cmdView")
 End if

 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q= "select * from PURCHASEORDER po" & _
	" left join supplier sp on (sp.fccode = po.suppliercode and sp.fcba = po.fcba) " & _
	" where po.PoCode='" & fcCode & "'"
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
	suppliername  = rsPurchase("fcname")
	termofpayment  = rsPurchase("termofpayment")
	dp  = rsPurchase("dp")
	ppn  = rsPurchase("ppn")
	totalppn  = rsPurchase("totalppn")
	totalamount  = rsPurchase("totalamount")
	totalamountppn = rsPurchase("totalamountppn")
	totalcharge = rsPurchase("totalcharge")
	grandtotal  = rsPurchase("grandtotal")
	RowStateId = rsPurchase("RowStateId")
	RowStateHdr = rsPurchase("RowStateId")
	Remarks=rsPurchase("Remarks")
 else
 ' 	PoDate=rsPurchase("PoDate")
	' PODATE = FormatDateTime(PODATE)
	' xCode=rsPurchase("SUPPLIERCODE")
	' SHIPPINGPOINT = 0
	' Remarks= 0
	RowStateId = 0
	RowStateHdr = 0
	use_ppn = 0
	IsInBridge  = 0
	' dp  = 0
	totalamountppn = 0
	totalcharge = 0
	totalamount  = 0
	totalppn  = 0
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
 
  'response.write "IsInBridge--> " & IsInBridge
has_approvedby = 0
q = " SELECT COUNT(1) total FROM purchaseorderdetails_hist WHERE pocode = '" & fcCode & "' AND fcba = '" & fcba & "' AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approvedby = cdbl(to_nol(rsPurchase("total")))
End if

q = " SELECT * FROM po_authorization_comments" & _
    " WHERE pocode = '" & fcCode & "'" & _
    "       AND fcba = '" & fcba & "'" & _
	"       AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	comments = rsPurchase("comments")
End if


if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="PurchaseOrderDetailsApproval.Asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Purchase Order No</td>
				<td align="left">
					<input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>" autocomplete="off">
					<font color="#FF0000">*</font>
				</td>
			
				<td align="left">Term Of Payment</td>
				<td align="left">
					<select name="termofpayment" value="<%=termofpayment%>">
						<option value="<%=termofpayment%>"></option>
						<option value="TOP7" <% If cstr(""&termofpayment) = "TOP7" Then %> Selected <%End if%>>TOP7</option> 
						<option value="TOP14" <% If cstr(""&termofpayment) = "TOP14" Then %> Selected <%End if%>>TOP14</option> 
						<option value="TOP30" <% If cstr(""&termofpayment) = "TOP30" Then %> Selected <%End if%>>TOP30</option> 
					</select>
				</td>
		</tr>
		<tr>
			<td align="left">Company</td>
				<td align="left">
					<select name="fccompanycode" value="<%=fccompanycode%>" id="fccompanycode" onchange="getFcba(this.value);">
					<option value=""></option> 
						<% 	Dim rsComboC
							q = " SELECT DISTINCT fccompanycode " & _
							 	" FROM (SELECT iplasprod.businessunit.fccompanycode " & _
							    " FROM iplasprod.user_fcba_tab " & _
							    " INNER JOIN iplasprod.businessunit " & _
							    " ON iplasprod.user_fcba_tab.fcba = " & _
							    " iplasprod.businessunit.fccode " & _
							    " WHERE iplasprod.businessunit.isactive = 'TRUE' AND iplasprod.user_fcba_tab.fcusercode = '"&  Session("CodeName") &"')"
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
					<select name="businessunit" id="businessunit" >
						<option value=""></option> 
						<% 	Dim rsComboB
							q = "  SELECT  * FROM businessunit WHERE fccompanycode = '"& fccompanycode &"'"
							' response.write q
							Set rsComboB = Server.CreateObject("ADODB.Recordset")
							rsComboB.CursorType = 3
							rsComboB.CursorLocation = 3
							rsComboB.Open q, connstr
							if not rsComboB.eof then%>
							<%
							Do While Not rsComboB.eof
							%>
								<option value="<%=rsComboB("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboB("fccode")) then %> selected 
									<%End If%>><%=rsComboB("fccode")%>
								</option> 
							<%rsComboB.movenext
							loop
							End if
							rsComboB.Close
							set rsComboB = Nothing
							q=""
						%>
					</select>
					<font color="#FF0000">*</font>
				</td>
			<td align="left">PPn (%)</td>
					<td align="left">
						<input type="text" size="5" name="ppn" value="<%=ppn%>" onkeypress="return harusangka(event);">
						<input type="text" size="5" name="totalppn" value="<%=totalppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Date</td>
				<td align="left">
					<input class="datepicker" type="text" maxlength="10" name="PODATE" size="10" value="<%=PODATE%>">
				</td>
			<td align="left">Total Amount</td>
					<td align="left">
						<input type="text" maxlength="50" name="totalamount" size="15" value="<%=totalamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Purchase By</td>
				<td align="left">      
					<select name="purchaseby" id="purchaseby" onchange="getSupplierPO(this.value);" >
						<option value="<%=purchaseby%>" <% If cstr(""&purchaseby) = Request("purchaseby") Then %> Selected <%End if%>><%=purchaseby%></option> 
					</select>
				</td>
				<td align="left">Total Amount + PPn</td>
					<td align="left">
						<input type="text" maxlength="50" name="totalamountppn" size="15" value="<%=totalamountppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
			<tr>
				<td align="left">Inco Terms</td>	
				<td align="left">
					<select name="SHIPPINGPOINT" value="<%=SHIPPINGPOINT%>">
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
			<td align="left">Supplier</td>
				<td align="left">      
					<select name="xCode" id="suppliercode" onchange="javascript:getSupp();"  >
						<option value="<%=xCode%>" <% If cstr(""&xCode) = Request("xCode") Then %> Selected <%End if%>><%=xCode%></option> 
					</select>
					<input type="text" maxlength="50" id="suppliername" name="suppliername" size="35" value="<%=suppliername%>" style="background:#eee" readonly>
				</td>
				<td align="left">Grand Total</td>
					<td align="left">
						<input type="text" maxlength="50" name="grandtotal" size="15" value="<%=grandtotal%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Supplier Address</td>
				<td align="left">
						<textarea id="supplieraddress"  name="supplieraddress" style="background:#eee" readonly><%=supplieraddress%></textarea>
				</td>
			
			<td align="left">Status</td>
					<td align="left">
						<input type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Delivery Destination</td>
				<td align="left">
					<select name="deliverydestination" id="deliverydestination" >
						<option value="<%=deliverydestination%>" <% If cstr(""&deliverydestination) = Request("deliverydestination") Then %> Selected <%End if%>><%=deliverydestination%></option> 
					</select>
				</td>	
				<td align="left">Remarks</td>
				<td align="left">
					<textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea>
				</td>
		</tr>
		<tr>
		<td align="left">Comments</td>
		<td align="left"><input type="text" maxlength="200" name="comments" size="50" value="<%=comments%>">
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
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Hidden" Value="<%=fcba%>" Name="fcba">
		<input Type="Hidden" Value="<%=authorizelevelcode%>" Name="authorizelevelcode">
		<input Type="Hidden" Value="<%=authorizetemplatecode%>" Name="authorizetemplatecode">
		<%if RowStateId <> "Rejected" then%>
			<%if has_approvedby=0 then%>
				<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Are you sure to Approve this data : <%=fcCode%> ?');"><i class="fa fa-check"></i> Approve</button>
				<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Are you sure to Reject this data : <%=fcCode%> ?');"><i class="fa fa-times"></i> Reject</button>
			<%else%>
				<b>This PO number : <%=fcCode%> has been Approved</b>
			<%end if%>
		<%else%>
			<b>This PO number : <%=fcCode%> has been Rejected</b>
		<%end if%>
		
		
		
		</td>
		</tr>
    </table>
<%

	q = " select *  from purchaseorderdetails pod" & _
		"  inner join stock st on (st.fcba = pod.fcba and st.fccode = pod.itemcode) " & _
		" left join purchaserequestdetails prd on (prd.prcode = pod.prcode and prd.lineid = pod.prlineid)" & _
		" where pod.PoCode = '" & fcCode & "'" & _
	    "       and pod.fcBA = '" & businessunit & "'"
	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.Open q, connstr
%>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="left">PO CODE</TD>
	<TD align="left">LINENOKEY</TD>
	<TD align="left">ITEM CODE</TD>
	<TD align="left">QUANTITY</TD>
	<TD align="left">APPROVED QUANTITY PR</TD>
	<TD align="left">UNIT PRICE</TD>
	<TD align="left">DISC</TD>
	<TD align="left">OTHER CHARGE</TD>
	<TD align="left">AMOUNT</TD>
	<TD align="left">OTHER ITEM DESC</TD>
	<TD align="left">RECEIVED</TD>
	<TD align="left">PRCODE</TD>
	<TD align="left">PRLINEID</TD>
	<TD align="left">LINE STATUS</TD>
	</tr>   
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
				<td align=" centers" title="<%=rsPurchaseOrder("pocode").Name%>" valign="top">
					<%=rsPurchaseOrder("pocode")%>
				</td>
				<td align="center" title="<%=rsPurchaseOrder("linenokey").Name%>" valign="top">
					<%=rsPurchaseOrder("linenokey")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("itemcode").Name%>" valign="top">
					<%=rsPurchaseOrder("itemcode")%> - <%=rsPurchaseOrder("fcname")%>
				</td>
				<td align="right" title="<%=rsPurchaseOrder("quantity").Name%>" valign="top">
					<%=rsPurchaseOrder("quantity")%>
				</td>
				<td align="right" title="<%=rsPurchaseOrder("approvedqty").Name%>" valign="top">
					<%=rsPurchaseOrder("approvedqty")%>
				</td>
				<td align="right" title="<%=rsPurchaseOrder("unitprice").Name%>" valign="top">
					<%=rsPurchaseOrder("unitprice")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("disc").Name%>" valign="top">
					<%=rsPurchaseOrder("disc")%> %
				</td>
				<td align="right" title="<%=rsPurchaseOrder("othercharge").Name%>" valign="top">
					<%=rsPurchaseOrder("othercharge")%>
				</td>
				<td align="right" title="<%=rsPurchaseOrder("amountcharge").Name%>" valign="top">
					<%=FormatNumber(rsPurchaseOrder("amountcharge"),0)%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otheritemdesc").Name%>" valign="top">
					<%=rsPurchaseOrder("otheritemdesc")%>
				</td>
				<td align="right" title="<%=rsPurchaseOrder("received").Name%>" valign="top">
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
					<%prcode = "0":if has_approveddata = 0 then prcode = rsPurchaseOrder("prcode") else prcode = rsPurchaseOrder("prcode")%>
					<input type="hidden" style="text-align:right;" name="prcode_<%=no%>" width="2" value="<%=prcode%>" >
					
					<%prlineid = "0":if has_approveddata = 0 then prlineid = rsPurchaseOrder("prlineid") else prcode = rsPurchaseOrder("prlineid")%>
					<input type="hidden" style="text-align:right;" name="prlineid_<%=no%>" width="2" value="<%=prlineid%>" >
					
					<%itemcode = "0":if has_approveddata = 0 then itemcode = rsPurchaseOrder("itemcode") else itemcode = rsPurchaseOrder("itemcode")%>
					<input type="hidden" style="text-align:right;" name="itemcode_<%=no%>" width="2" value="<%=itemcode%>" >
			</tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </div>
   </TABLE>
   </form>
</div></div>
  <!--#include file="bottom.jsp"-->
 	<script type="text/javascript">
		function handleClick(cb){
			if(cb.checked == 1){
				cb.value = "Y";
			} else {
				cb.value = "N";
			}
        }
	</script>
		<script src="/js/jquery.inputmask.bundle.min.js"></script>
		<script src="/js/inputmask.binding.js"></script>
</body>
</html>
