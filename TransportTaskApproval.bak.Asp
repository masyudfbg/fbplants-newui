<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Transport Task Approval</title>
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
    // document.getElementById("suppliername").value = s.suppliername;
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

    orderno = request("orderno")
    orderdate = request("orderdate")
    remarks = request("remarks")
    rowstateid = request("rowstateid")
 
 
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
 JobCode=Request("JobCode")
 LocationType="CB"
 LocationCode="BS"
 Remarks = Request("Remarks")
 Reference = Request("Reference")
 ITEMCODE=Request("ITEMCODE")
 STORECODE = Request("STORECODE")
 LineNo=Request("LineNo")
 QuantityTransfered=Request("QuantityTransfered")
 QuantityReceived=Request("QuantityReceived")
 Price = Request("Price")
 AMOUNT =Request("AMOUNT")
 Linenohist = Request("linenohist")
 FcbaDistination = Request("FcbaDistination")
 rowstateid = Request("rowstateid")
 document_classification = Request("document_classification")
 comments = request ("comments")

if "" & use_ppn = "" then use_ppn = "N" 

fcba = request("fcba")
authorizelevelcode = request("authorizelevelcode")
authorizetemplatecode = request("authorizetemplatecode")
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "ApprovalAction" or Request("txtJenis")= "RejectAction" Then

	if Request("txtJenis")="ApprovalAction" then

		q = " INSERT INTO transporttask_hist ( orderno" & _
			"                                           , currentauthorizelevel" & _
			"                                           , fcentry" & _
			"                                           , fcedit" & _
			"                                           , fcip" & _
			"                                           , fcba" & _
			"                                           , lastupdate" & _
			"                                           , lasttime)"& _
			" VALUES ( '" & orderno & "'" & _
			"          , '" & authorizelevelcode & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"          , '" & fcba & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
		' response.write q
		connStr.execute(q)
		
		q = " INSERT INTO tt_authorization_comments ( orderno" & _
			"                                         , currentauthorizelevel" & _
			"                                         , comments" & _
			"                                         , fcentry" & _
			"                                         , fcedit" & _
			"                                         , fcip" & _
			"                                         , fcba" & _
			"                                         , lastupdate" & _
			"                                         , lasttime)"& _
			" VALUES ( '" & orderno & "'" & _
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

		is_final = 0
		q = " SELECT count (distinct authorizelevelcode) total FROM authorizelevel al" & _
			" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
			" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
			" WHERE at.document_type = 'TT'" & _
			"       AND DECODE(at.document_type, 'PR', at.document_classification, '0') = '0'" & _
			"       AND al.isactive = 'TRUE'" & _
			"       AND al.finallevel = 'TRUE'" & _
			"       AND al.fccode = '" & authorizelevelcode & "'" & _
			"       AND aus.authorizetemplatecode = '" & authorizetemplatecode & "'" & _
			"       AND aus.fcba = '" & fcba & "'"
		' response.write q
		set rsPurchase = server.CreateObject("Adodb.recordset")
		rsPurchase.Open q, connstr
		IF not rsPurchase.eof then
			is_final = cdbl(to_nol(rsPurchase("total")))
		end if
		
		if is_final = 1 then
			q = "delete TransportTaskDtl_tab " & _
				"where ORDERNO = '" & ORDERNO & "' and " & _
				"fcba = '" & fcba &"' "
				' response.write q
				connStr.execute(q)
			For n=1 to (cdbl(request("no"))-1)
				if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 Then

		' Insert STock On Hand dan Stock Trans Hist
			' set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
		 '   q="select ORDERNO,LINENO, LINENOHIST,STOCKCODE,stock_api.get_stockName(Stockcode,fcba) StockName, JOBCODE ,LOCATIONTYPE, LOCATIONCODE, QUANTITYTRANSFERED, STORECODE, FCBADISTINATION, QUANTITYRECEIVED," & _
		 '     " Price, AMOUNT, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA  from TransportTaskDtl_tab  " & _
		 '     " where OrderNo='" & OrderNo & "' and " & _
		 '     " fcBA='" & fcba & "'"
		 '   response.Write q	
		 '   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
		 '   rsPurchaseOrder.Open q, connstr
		 '    if Not rsPurchaseOrder.Eof then
			'     ORDERNO = rsPurchaseOrder("ORDERNO")
			' 	LINENO = rsPurchaseOrder("LINENO")
			' 	LINENOHIST = rsPurchaseOrder("LINENOHIST")
			' 	ITEMCODE = rsPurchaseOrder("STOCKCODE")
			' 	StockName = rsPurchaseOrder("StockName")
			' 	JOBCODE = rsPurchaseOrder("JOBCODE")
			' 	LOCATIONTYPE = rsPurchaseOrder("LOCATIONTYPE")
			' 	LOCATIONCODE = rsPurchaseOrder("LOCATIONCODE")
			' 	QUANTITYTRANSFERED = cdbl(to_nol(rsPurchaseOrder("QUANTITYTRANSFERED")))
			' 	STORECODE = rsPurchaseOrder("STORECODE")
			' 	FCBADISTINATION = rsPurchaseOrder("FCBADISTINATION")
			' 	QUANTITYRECEIVED = rsPurchaseOrder("QUANTITYRECEIVED")
			' 	Price = rsPurchaseOrder("Price")
			' 	AMOUNT = rsPurchaseOrder("AMOUNT")
			' 	REMARKS = rsPurchaseOrder("REMARKS")
			' End if

			' Add by NAF to handle Inventory Stock On Hand
				' Insert the History of Received Inventory Transaction
		
				qLessZero = "select islesszero from stock " & _	   
					"where " & _
					"fccode = '" & request("itemcode_"&TRIM(n)) & "' and " & _
					"fcba = '" & request("fcba_"&TRIM(n)) & "'" 
				
				set rsLessZero = Server.CreateObject("ADODB.Recordset")
				rsLessZero.Open qLessZero, connstr
				if not rsLessZero.eof then
					tempislesszero = rsLessZero("islesszero")
				else
					tempislesszero = "FALSE"
				end if 
					
				' We have to Update the Quantity by Adding them with the New one.
				
				
				qQuantityOnHand = "select quantity from stockonhand " & _
								  "where " & _
								  "storecode = '" & request("storecode_"&TRIM(n)) & "' and " & _
								  "stockcode  = '" & request("itemcode_"&TRIM(n)) & "' and " & _ 
								  "fcba = '" & request("fcba_"&TRIM(n)) &"' "
	
            	' response.write 	qQuantityOnHand		
				
				set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
				rsQuantityOnHand.Open qQuantityOnHand, connstr
				if not rsQuantityOnHand.eof then
					tempQuantityOnHand = cdbl(to_nol(rsQuantityOnHand("quantity")))
				else
					tempQuantityOnHand = 0
				end if 		

				
				if (tempQuantityOnHand - QuantityTransfered >= 0) or tempislesszero = "TRUE"  then 
				
					qLineMax =  " select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
								" where " & _
								" fcba = '" & request("fcba_"&TRIM(n)) & "'" 
					
					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
					else
						tempMaxLine = 0
					end if 

					tempMaxLine = tempMaxLine + 1
					
					q = "Insert into StockTransHist ( " & _
						" LINENO, FCBA, DOCUMENTNO, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, Quantity, LASTUPDATE, LASTTIME) values ( "& _
						  tempMaxLine & ",'" & _ 
						  request("fcba_"&TRIM(n)) & "','" & _
						  request("orderno_"&TRIM(n)) & "','" & _
						  request("itemcode_"&TRIM(n)) &"'," & _
						  "to_date('" & OrderDate & "','mm/dd/yyyy') ,'" & _
						  "ST_TransportTask" & "','" & _
						  request("storecode_"&TRIM(n)) & "'," & _
						  "'ST','" & _
						  request("storecode_"&TRIM(n)) & "','" & _
						  request("locationtype_"&TRIM(n)) & "','" & _
						  request("locationcode_"&TRIM(n)) & "'," & _
						  -1 * request("quantitytransfered_"&TRIM(n)) & ", sysdate, To_Char(sysdate,'HH24:MI') ) "
						' response.write "Stock Transhist </br>" & q
						connStr.execute(q)

					' q = "update TransportTaskDtl_tab " & _
					' 	"set LINENOHIST = " & tempMaxLine & " " & _
					' 	", LASTUPDATE = sysdate " & _
					' 	", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					' 	"where " & _
					' 	"ORDERNO = '" & ORDERNO & "' and " & _
					' 	"fcba = '" & fcba &"' "
							 
					' 	connStr.execute(q)
					strLineMax = " select nvl(max(LineNo),0) maxLineNo from TransportTaskDtl_tab " & _
							   " where " & _
							   " fcba = '" & request("fcba_"&TRIM(n)) & "'"
					
					' response.write strLineMax		   
					Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
					rsstrLineMax.CursorType = 3
					rsstrLineMax.CursorLocation = 3
					rsstrLineMax.Open strLineMax, connstr
					
					if not rsstrLineMax.eof then
						tempMaxLineNo = cdbl(rsstrLineMax("maxLineNo")) + 1
					end if
								
					rsstrLineMax.Close
					set rsstrLineMax = Nothing 

					q=  "Insert into  TransportTaskDtl_tab (ORDERNO, STOCKCODE,LINENO,LINENOHIST, LOCATIONCODE,LOCATIONTYPE, QUANTITYTRANSFERED,STORECODE,JOBCODE,FCBADISTINATION,QUANTITYRECEIVED,PRICE,AMOUNT, " & _
						"REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
						"values ('" & request("orderno_"&TRIM(n)) & "'" & _
						" ,'" & request("itemcode_"&TRIM(n)) & "'" & _
						" ,'" & tempMaxLineNo & "'" & _
						" ," & tempMaxLine &""& _
						" ,'"  & request("locationcode_"&TRIM(n)) & "'"& _
						" ,'" & request("locationtype_"&TRIM(n)) & "'" & _
						" ," & request("quantitytransfered_"&TRIM(n)) & "" & _
						" ,'" & request("storecode_"&TRIM(n)) & "'" & _
						" ,'" & request("jobcode_"&TRIM(n)) & "'" & _
						" ,'" & request("fcbadistination_"&TRIM(n)) & "'" & _
						" ,0" & _
						" ," &  request("price_"&TRIM(n)) & "" & _
						" ," &  request("amount_"&TRIM(n)) & "" & _
						" ,'" & request("remarks_"&TRIM(n)) & "'" & _
						" ,'" & Session("CodeName")  & "'" & _
						" ,'" & Session("CodeName") & "'" & _
						" ,'" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						" ,'" & request("fcba_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI') )"
						' response.Write(q)
					connStr.execute(q)

					qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
										 "where " & _
										 "storecode = '" & request("storecode_"&TRIM(n)) & "' and " & _												 
										 "stockcode = '" & request("itemcode_"&TRIM(n)) & "' and " & _
										 "fcba = '" & request("fcba_"&TRIM(n)) & "'"

					set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
					rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
					if not rsTotalRecordOnHand.eof then
						tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
					else
						tempTotalRecordOnHand = 0
					end if 
							
					if tempTotalRecordOnHand > 0 then
					
						q = "update stockonhand " & _
							"set quantity = quantity - " & request("QUANTITYTRANSFERED_"&TRIM(n)) & " " & _
							", LASTUPDATE = sysdate " & _
							", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							"where " & _
							"storecode = '" & request("storecode_"&TRIM(n)) & "' and " & _
							"stockcode  = '" & request("itemcode_"&TRIM(n)) & "' and " & _ 
							"fcba = '" & request("fcba_"&TRIM(n)) &"' "
							 
						connStr.execute(q)
					else
						' Just Insert the New Record with the following/received Quantity
						q = " insert into stockonhand (storecode" & _
						    "                          , stockcode" & _
						    "                          , fcba" & _
						    "                          , quantity" & _
						    "                          , lastupdate" & _
						    "                          , lasttime )" & _
						    " values ('" & request("storecode_"&TRIM(n)) & "'" & _
						    "         , '" & request("itemcode_"&TRIM(n)) & "'" & _
						    "         , '" & request("fcba_"&TRIM(n)) & "'" & _
						    "         , '0'" & _
						    "         , sysdate" & _
						    "         , To_Char(sysdate,'HH24:MI') ) "
						    ' response.write q
						connStr.execute(q)
						
						q = " update stockonhand " & _
							" set quantity = quantity - " & request("QUANTITYTRANSFERED_"&TRIM(n)) & " " & _
							"     , LASTUPDATE = sysdate " & _
							"     , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							" where storecode = '" & request("storecode_"&TRIM(n)) & "'" & _
							"       and stockcode  = '" & request("itemcode_"&TRIM(n)) & "'" & _ 
							"       and fcba = '" & request("fcba_"&TRIM(n)) & "'"
							 
						connStr.execute(q)
					end if
				else
					err.number = 1001
					err.description = "Nilai Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
				end if
			
		' End Insert

			strQueryStatus = "UPDATE transporttask_tab SET rowstateid = 'Approved' WHERE fcba = '" & request("fcba_"&TRIM(n)) & "' AND orderno = '" & request("orderno_"&TRIM(n)) & "'"
			connStr.execute(strQueryStatus)	

			strQueryStatus = "BEGIN TRANSPORTTASK_API.NextState('" & request("fcba_"&TRIM(n)) & "','" & request("orderno_"&TRIM(n)) & "', 'Approve' ); END; "
			connStr.execute(strQueryStatus)
			end if
		next
	end if
	
		cmdView	= "True"
	elseif Request("txtJenis")="RejectAction" then
		q = " INSERT INTO tt_authorization_comments ( orderno" & _
			"                                         , currentauthorizelevel" & _
			"                                         , comments" & _
			"                                         , fcentry" & _
			"                                         , fcedit" & _
			"                                         , fcip" & _
			"                                         , fcba" & _
			"                                         , lastupdate" & _
			"                                         , lasttime)"& _
			" VALUES ( '" & orderno & "'" & _
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
		
		
	    strQueryStatus = "BEGIN TRANSPORTTASK_API.NextState('" & fcba & "','" & orderno & "', 'Reject' ); END; "
		connStr.execute(strQueryStatus)
		
		cmdView	= "True"
    else
        
	end if								 
	cmdView	= "True"
 else
    cmdView	= request("cmdView")
 End if

 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select * from transporttask_tab  " & _
  " where orderno='" & orderno & "' AND fcba = '"&fcba&"'"
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    orderdate=rsPurchase("orderdate")
    remarks=rsPurchase("remarks")
    rowstateid=rsPurchase("rowstateid")
    document_classification=rsPurchase("document_classification")
 End if
 
 
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
'	connStr.RollbackTrans
    response.Write  "<div class='error'>" & err.description & "</div>"
End if

has_approvedby = 0
q = " SELECT COUNT(1) total FROM transporttask_hist WHERE orderno = '" & orderno & "' AND fcba = '" & fcba & "' AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approvedby = cdbl(to_nol(rsPurchase("total")))
End if

q = " SELECT * FROM tt_authorization_comments" & _
    " WHERE orderno = '" & orderno & "'" & _
    "       AND fcba = '" & fcba & "'" & _
	"       AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	comments = rsPurchase("comments")
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="TransportTaskApproval.Asp">
	<table class="data" width="100%">
		<tr>
		  <td align="left">Order No</td>
		  <td align="left"><input type="text" maxlength="15" name="OrderNo" size="30" value="<%=OrderNo%>"><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
		   <td align="left">Date</td>
		   <td align="left"><input class="datepicker" type="text" maxlength="10" name="OrderDate" size="10" value="<%=OrderDate%>"><font color="#FF0000">&nbsp;*</font>
		</td>
		</tr>
		<tr>
			<td align="left">Classification</td>
				<td align="left">
					<select name="document_classification" class="chosen-select" style="width: 50%">
					<option></option>
					<%
					qLov = " SELECT fccode, fcname FROM document_classification ORDER BY fccode"
		   
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
		</tr>
		<tr> 
		    <td align="left">REMARKS</td>
			<td align="left"><textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea></td>
		</tr>
		<tr>
		  <td align="left">Status</td>
		  <td align="left"><input type="text" maxlength="50" name="rowstateid" size="30" value="<%=rowstateid%>" style="background-color:#eee;" readonly><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
		<td align="left">Comments</td>
		<td align="left"><input type="text" maxlength="200" name="comments" size="50" value="<%=comments%>">
		</td>
		</tr>
    </table>
<%

	q = " select ORDERNO,LINENO, LINENOHIST,STOCKCODE,stock_api.get_stockName(Stockcode,fcba) StockName, JOBCODE ,LOCATIONTYPE, LOCATIONCODE, QUANTITYTRANSFERED, STORECODE, FCBADISTINATION, QUANTITYRECEIVED," & _
     " Price, AMOUNT, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA  from TransportTaskDtl_tab" & _
		" where orderno = '" & orderno & "'" & _
	    "       and fcBA = '" & fcba & "'"
	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.Open q, connstr
%>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="left">Order No</TD>
	<TD align="left">Line No</TD>	
	<TD align="left">Line No Hist</TD>	
	<TD align="left">Stock Code</TD>
	<TD align="left">Stock Name</TD>
	<TD align="left">Job Code</TD>
	<TD align="left">Loc. Type</TD>	
	<TD align="left">Loc. Code</TD>
	<TD align="left">Quantity Transfered</TD>	
	<TD align="left">Store Code</TD>	
	<TD align="left">Fcba Distination</TD>	
	<TD align="left">Quantity Received</TD>	
	<TD align="left">Unit Price</TD>
	<TD align="left">Amount</TD>	
	<TD align="left">Remarks</TD>
	<TD align="left"></TD>
	</tr>   
   <%
   
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
				<td align="left" title="<%=rsPurchaseOrder("ORDERNO").Name%>" valign="top"><%=rsPurchaseOrder("ORDERNO")%></td>
				<td align="left" title="<%=rsPurchaseOrder("LINENO").Name%>" valign="top"><%=rsPurchaseOrder("LINENO")%></td>
				<td align="left" title="<%=rsPurchaseOrder("LINENOHIST").Name%>" valign="top"><%=rsPurchaseOrder("LINENOHIST")%></td>
				<td align="left" title="<%=rsPurchaseOrder("STOCKCODE").Name%>" valign="top"><%=rsPurchaseOrder("STOCKCODE")%></td>
				<td align="left" title="<%=rsPurchaseOrder("StockName").Name%>" valign="top"><%=rsPurchaseOrder("StockName")%></td>
				<td align="left" title="<%=rsPurchaseOrder("JOBCODE").Name%>" valign="top"><%=rsPurchaseOrder("JOBCODE")%></td>
				<td align="left" title="<%=rsPurchaseOrder("LOCATIONTYPE").Name%>" valign="top"><%=rsPurchaseOrder("LOCATIONTYPE")%></td>
				<td align="left" title="<%=rsPurchaseOrder("LOCATIONCODE").Name%>" valign="top"><%=rsPurchaseOrder("LOCATIONCODE")%></td>
				<td align="left" title="<%=rsPurchaseOrder("QUANTITYTRANSFERED").Name%>" valign="top"><%=rsPurchaseOrder("QUANTITYTRANSFERED")%></td>
				<td align="left" title="<%=rsPurchaseOrder("STORECODE").Name%>" valign="top"><%=rsPurchaseOrder("STORECODE")%></td>
				<td align="left" title="<%=rsPurchaseOrder("FCBADISTINATION").Name%>" valign="top"><%=rsPurchaseOrder("FCBADISTINATION")%></td>
				<td align="left" title="<%=rsPurchaseOrder("QUANTITYRECEIVED").Name%>" valign="top"><%=rsPurchaseOrder("QUANTITYRECEIVED")%></td>
				<td align="left" title="<%=rsPurchaseOrder("Price").Name%>" valign="top"><%=rsPurchaseOrder("Price")%></td>
				<td align="left" title="<%=rsPurchaseOrder("AMOUNT").Name%>" valign="top"><%=rsPurchaseOrder("AMOUNT")%></td>
				<td align="left" title="<%=rsPurchaseOrder("REMARKS").Name%>" valign="top"><%=rsPurchaseOrder("REMARKS")%></td>
				<td valign="middle" align="center">
					<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "1" onclick='handleClick(this);' checked>
					<input Type="Hidden" Value="True" Name="cmdView">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					<input type="hidden" name="ORDERNO_<%=NO%>" value="<%=rsPurchaseOrder("ORDERNO")%>">
					<input type="hidden" name="LINENO_<%=NO%>" value="<%=rsPurchaseOrder("LINENO")%>">
					<input type="hidden" name="LINENOHIST_<%=NO%>" value="<%=rsPurchaseOrder("LINENOHIST")%>">
					<input type="hidden" name="ITEMCODE_<%=NO%>" value="<%=rsPurchaseOrder("STOCKCODE")%>">
					<input type="hidden" name="StockName_<%=NO%>" value="<%=rsPurchaseOrder("StockName")%>">
					<input type="hidden" name="JOBCODE_<%=NO%>" value="<%=rsPurchaseOrder("JOBCODE")%>">
					<input type="hidden" name="LOCATIONTYPE_<%=NO%>" value="<%=rsPurchaseOrder("LOCATIONTYPE")%>">
					<input type="hidden" name="LOCATIONCODE_<%=NO%>" value="<%=rsPurchaseOrder("LOCATIONCODE")%>">
					<input type="hidden" name="fcba_<%=NO%>" value="<%=rsPurchaseOrder("fcba")%>">
					<input type="hidden" name="QUANTITYTRANSFERED_<%=NO%>" value="<%=rsPurchaseOrder("QUANTITYTRANSFERED")%>">
					<input type="hidden" name="STORECODE_<%=NO%>" value="<%=rsPurchaseOrder("STORECODE")%>">
					<input type="hidden" name="FCBADISTINATION_<%=NO%>" value="<%=rsPurchaseOrder("FCBADISTINATION")%>">
					<input type="hidden" name="Price_<%=NO%>" value="<%=rsPurchaseOrder("Price")%>">
					<input type="hidden" name="AMOUNT_<%=NO%>" value="<%=rsPurchaseOrder("AMOUNT")%>">
					<input type="hidden" name="REMARKS_<%=NO%>" value="<%=rsPurchaseOrder("REMARKS")%>">
				</td>
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
   	<tr>
		<input Type="Hidden" Value="<%=no%>" Name="no">
		<input Type="Hidden" Value="<%=authorizelevelcode%>" Name="authorizelevelcode">
		<input Type="Hidden" Value="<%=authorizetemplatecode%>" Name="authorizetemplatecode">
		<input Type="Hidden" Value="<%=fcba%>" Name="fcba">
   		<%if RowStateId <> "Rejected" then%>
			<%if has_approvedby=0 then%>
			<td align="left">
				<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Are you sure to Approve this data : <%=orderno%> ?');"><i class="fa fa-check"></i> Approve</button>
			</td>
			<td align="left">
				<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Are you sure to Reject this data : <%=orderno%> ?');"><i class="fa fa-times"></i> Reject</button>
			</td>
			<%else%>
				<b>This Order number : <%=orderno%> has been Approved</b>
			<%end if%>
		<%else%>
			<b>This Order number : <%=orderno%> has been Rejected</b>
		<%end if%>
	</tr>
   </div>
   </TABLE>
   </form>
</div></div>
  <!--#include file="bottom.jsp"-->
<script type="text/javascript">
	function handleClick(cb){
		if(cb.checked == 0){
			cb.value = 1;
		} else {
			cb.value = 0;
		}
	}
	

</script>
<script type="text/javascript">	
		$('#select-all').click( function () {
			$('.item').prop('checked', this.checked);
			$('.item').val(this.checked ? 1 : 0);
		})
</script>

		<script src="/js/jquery.inputmask.bundle.min.js"></script>
		<script src="/js/inputmask.binding.js"></script>
</body>
</html>
