<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Spb Harvesting Update</title>
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
    if (myConn) {
        
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

	// launched from button click 
	var GetEmpName = function () {
		// build up the post string when passing variables to the server side page
		// var suppliercode = document.getElementById("suppliercode").value;
		var driver = document.getElementById("driver").value;
		var PostStr = 'validate=empcd&q='+driver;
		//alert(PostStr);
		// use the generic function to make the request
		doAJAXCall('ajax/validate.asp', 'POST', PostStr, showEmpnameResponse);
		
	}

	// The function for handling the response from the server
	var showEmpnameResponse = function (oXML) { 
	    
	    // get the response text, into a variable
	    var response = oXML.responseText;
	    // alert(s.supplieraddress);
	    // update the Div to show the result from the server
	    // document.getElementById("suppliername").value = s.suppliername;
		// document.getElementById("supplieraddress").value = s.supplieraddress;
		document.getElementById("nonscript_empcd").innerHTML = response;
		// alert(s.suppliername);
	    
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

//  var xmlhttp;
//  if (str=="")
//    {
//    document.getElementById("SHIPPINGPOINT").innerHTML="";
//    // document.getElementById("purchaseby").innerHTML="";
//    // document.getElementById("receviedtask").innerHTML="";
//    return;
//    }
//  if (window.XMLHttpRequest)
//    {// code for IE7+, Firefox, Chrome, Opera, Safari
//    xmlhttp=new XMLHttpRequest();
//    StockName=new XMLHttpRequest();
//    }
//  else
//    {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
//  xmlhttp.onreadystatechange=function()
//    {
//    if (xmlhttp.readyState==4 && xmlhttp.status==200)
//      {
//      alert(xmlhttp.responseText);
//      document.getElementById("deliverydestination").innerHTML=xmlhttp.responseText;
//      // document.getElementById("purchaseby").innerHTML=xmlhttp.responseText;
//      // document.getElementById("receviedtask").innerHTML="";
//      // alert(businessunit);
//      }
//    }
//  // xmlhttp.open("GET","ajax/getSupplierPO.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
//  xmlhttp.open("GET","SHIPPINGPOINT="+str+"",true);
//  xmlhttp.send();
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
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = "" 
 Dim IsInBridge, IsFrezee,HdrState
 
 Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

CHITNO=request("CHITNO")
CROPCODE=request("CROPCODE")
PRODUCTCODE=request("ProductCODE")
receptiondate=request("receptiondate")
VEHICLE=request("VEHICLE")
OWN=request("OWN")
TRANSPORTTYPE = Request("TRANSPORTTYPE")
CONTRACTNUMBER = Request("CONTRACTNUMBER")
QUANTITY=request("QUANTITY")
MILL_WEIGHT_BRUTO=request("MILL_WEIGHT_BRUTO")
MILL_WEIGHT=request("MILL_WEIGHT")
SPB_NO = request ("SPB_NO")
SPB_TYPE = request ("SPB_TYPE")
AGREEMENTCODE = request ("AGREEMENTCODE")
DRIVER = request ("DRIVER")
MILL = request ("MILL")
lineno = request("lineno")

' Detail
QUANTITY_DTL=request("QUANTITY_DTL")
FIELDCODE=request("FIELDCODE") 
pressemester_abw=request("pressemester_abw")
bucket=request("bucket")
bunch_estateweight=request("bunch_estateweight")
bunch=request("bunch")
harvestdate = request("harvestdate")
total_bunch=request("total_bunch")
total_bucket = request("total_bucket")
total_estateweight = request ("total_estateweight")

if len(trim(""&spb_type)) = 0 then spb_type = 0
If Len(Trim(""&receptiondate)) = 0 Then	receptiondate = textDate(date)
If Len(Trim(harvestdate)) = 0 Then	harvestdate = textDate(date)

' End Added
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "EditHeader"   Then

    if Request("txtJenis")="EditHeader" then
     
      q =	" Update harvestingspb" & _
			" set receptiondate = TO_DATE('" & receptiondate & "', 'mm/dd/yyyy')" & _
			"     , transporttype = '" & transporttype & "'" & _
			"     , cropcode = '" & cropcode & "'" & _
			"     , productcode = '" & productcode & "'" & _
			"     , own = '" & own & "'" & _
			"     , agreementcode = '" & agreementcode & "'" & _
			"     , driver = '" & driver & "'" & _
			"     , vehicle = '" & vehicle & "'" & _
			"     , mill = '" & mill & "'" & _
			"     , spb_type = '" & spb_type & "'" & _
			"     , fcentry = '" & Session("CodeName") & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , lastupdate = sysdate " & _
			"     , lasttime = to_char(sysdate,'hh24:mi') " & _
			" where fcBA='" & session("fcba") & "' and spbno = '"& spb_no &"'"
		' response.write q
		connStr.execute(q)
     end if                              
        cmdView = "True"   
     else
        cmdView = request("cmdView")
     End if


 
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
    connStr.RollbackTrans
    response.Write  "<div class='error'>" & err.description & "</div>"
End if

 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="spbharvestingupdate.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">SPB No</td>
			<td align="left">
				<input type="text" maxlength="50" name="SPB_NO" size="30" value="<%=SPB_NO%>" autocomplete="Off">
				<font color="#FF0000"> *</font>
			</td>
			<td align="left" width="20%">Vehicle Status </td>
			<td align="left">
				<select id="own" name="OWN" >
				<%
					Dim rsVehicleStatus, strVehicleStatus
					strVehicleStatus = " SELECT  FCNAME FROM HARVESTINGVEHICLESTATUS WHERE FCCODE <> '03' ORDER BY FCCODE"
					
					Set rsVehicleStatus = Server.CreateObject("ADODB.Recordset")
					rsVehicleStatus.Open strVehicleStatus, connstr
					
					If Not rsVehicleStatus.EOF then
						Do While Not rsVehicleStatus.EOF
				%>
					<option value = "<%=rsVehicleStatus("FCNAME")%>" <%If OWN = rsVehicleStatus("FCNAME") Then%> selected <%End If%>>
						<%=rsVehicleStatus("FCNAME")%>
					</option> 
				<%			rsVehicleStatus.movenext
						Loop
					End if
					rsVehicleStatus.Close
					set rsVehicleStatus = Nothing
				%>
				</select>
			</td>  
		</tr>
		<tr>
			<td align="left" width="20%">SPB Date</td>
			<td align="left"><input class="datepicker" type="text" size="15" name="receptiondate" id="receptiondate" value="<%=receptiondate%>"><font color="#FF0000"> *</font></td>
			 
			<td align="left" width="20%">Vehicle Code</td>
			<td align="left"><input type="text" size="15" name="VEHICLE" value="<%=VEHICLE%>"  >
			 <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Vehicle&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
			<font face="Verdana" size="1">Vehicle Code</font>
			</td> 
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Transport Type</font></td>
			<td align="left">
				<select name="TRANSPORTTYPE">
				<%
					Dim rsTransportType, strTransportType
					strTransportType = " SELECT FCCODE, FCNAME FROM HARVESTINGTRANSPORTTYPE ORDER BY FCCODE"
					
					Set rsTransportType = Server.CreateObject("ADODB.Recordset")
					rsTransportType.Open strTransportType, connstr
					
					If Not rsTransportType.EOF then
						Do While Not rsTransportType.EOF
				%>
					<option value = "<%=rsTransportType("FCCODE")%>" <%If TRANSPORTTYPE = rsTransportType("FCCODE") Then%> selected <%End If%>>
						<%=rsTransportType("FCNAME")%>
					</option> 
				<%			rsTransportType.movenext
						Loop
					End if
					rsTransportType.Close
					set rsTransportType = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
			<%if cstr(""&OWN) = "CONTRACT" then%>
			<td align="left" width="20%"> Driver</td>
			<td align="left">
				<input id="driver" type="text" size="15" name="driver" value="<%=driver%>" onchange="GetEmpName(this.value);" onkeypress="return hanyaAngka(event);" >
				<span id="nonscript_empcd" style="color:black;">&nbsp;*&nbsp;
				<a id ="nonscriptempcd_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
				<font face="Verdana"   size="1">Employee</font></a></span>
			</td>
			<%else%>
			<td align="left" width="20%"> Driver</td>
			<td align="left">
				<input id="driver" type="text" size="15" name="driver" value="<%=driver%>" onchange="GetEmpName(this.value);" onkeypress="return hanyaAngka(event);" >
				<span id="nonscript_empcd" style="color:black;">&nbsp;*&nbsp;
				<a id ="nonscriptempcd_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
				<font face="Verdana"   size="1">Employee</font></a></span>
			<%end if%>
		</tr>
		<tr> 
			<td align="left" width="20%">Crop Code</td>
			<td align="left">
				<select name="CROPCODE">
				<%
					Dim rsCrop, strCrop
					strCrop =	" SELECT P.FCCODE,P.FCNAME FROM  PARAMETER P  WHERE P.PARCODE= 'F0005' AND FCBA = '" &  Session("fcBA") & "'" & _
								" AND P.FCCODE = 'OP' ORDER BY P.FCNAME DESC"
					
					Set rsCrop = Server.CreateObject("ADODB.Recordset")
					rsCrop.Open strCrop, connstr
					
					If Not rsCrop.EOF then
						Do While Not rsCrop.EOF
				%>
					<option value="<%=rsCrop("FCCODE")%>" <%If CROPCODE = rsCrop("FCCODE") Then%> selected <%End If%>><%=rsCrop("FCNAME")%></option> 
				<%			rsCrop.movenext
						Loop
					End if
					rsCrop.Close
					set rsCrop = Nothing
				%>
				</select>
			</td>
			<td align="left" width="20%"> Mill</td>
			<td align="left"><input type="text" size="30" name="MILL" value="<%=MILL%>"></td>
		</tr>
		<tr>
			<td align="left" width="20%">Product Code</td>
			<td align="left">
				<select name="ProductCODE">
				<%
					Dim rsProductCode, strProductCode
					strProductCode = " SELECT PRODUCTCODE, PRODUCTDESCRIPTION FROM PRODUCT WHERE FCBA = '" &  Session("fcBA") & "' AND PRODUCTCODE = 'TBS' ORDER BY PRODUCTCODE"
					
					Set rsProductCode = Server.CreateObject("ADODB.Recordset")
					rsProductCode.Open strProductCode, connstr
					
					If Not rsProductCode.EOF then
						Do While Not rsProductCode.EOF
				%>
					<option value="<%=rsProductCode("PRODUCTCODE")%>" <%If PRODUCTCODE = rsProductCode("PRODUCTCODE") Then%> selected <%End If%>><%=rsProductCode("PRODUCTCODE")%> : <%=rsProductCode("PRODUCTDESCRIPTION")%></option> 
				<%			rsProductCode.movenext
						Loop
					End if
					rsProductCode.Close
					set rsProductCode = Nothing
				%>
				</select>
			</td> 
			<td align="left" width="20%">Weighbridge</td>
			<td align="left">
				<input type="checkbox" id="spb_type" name="spb_type" value="<%=spb_type%>" <%if cstr(""&spb_type) = "1" then%> checked <%end if%> onchange="handleClick(this);"> Available
			</td>
		</tr>  
		<tr>
			<%if cstr(""&OWN) <> "CONTRACT" then%>	
			<td align="left" width="20%"> Contract Code</td>
			<td align="left"><input type="text" size="30" id="agreementcode" name="AGREEMENTCODE" value="<%=AGREEMENTCODE%>" ></td>
			<%else%>
			<td align="left" width="20%"> Contract Code</td>
			<td align="left"><input type="text" size="30" name="AGREEMENTCODE" value="<%=AGREEMENTCODE%>"></td>
			<%end if%>	
		</tr>
		
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<button type="submit" class="button" name="txtJenis" value="EditHeader" onClick="return confirm('Are you sure to Update this data : <%=spb_no%> ?');"><i class="fa fa-check"></i> Save</button>
				
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
  <!--#include file="bottom.jsp"-->
 <script type="text/javascript">
	function handleClick(avail){
		if(avail.checked == 1){
			avail.value = 1;
		} else {
			avail.value = 0;
		}
	}
	
	function cekspbtype() {
		if($('#spb_type').is(":checked")){
			$('#bunch_estateweight').css("background-color", "#fff");
			$('#bunch_estateweight').prop('readOnly', false);
		}else{
			$('#bunch_estateweight').css("background-color", "#eee");
			$('#bunch_estateweight').prop('readOnly', true);
		}
	}
</script>
<script type="text/javascript">	

	// $('#spb_type').click( function () {
	// 	//$('.item').prop('checked', this.checked);
	// 	//$('.item').val(this.checked ? 1 : 0);
	// 	if($('#spb_type').is(":checked")){
	// 		$('#bunch_estateweight').css("background-color", "#fff");
	// 		$('#bunch_estateweight').prop('readOnly', false);
	// 	}else{
	// 		$('#bunch_estateweight').css("background-color", "#eee");
	// 		$('#bunch_estateweight').prop('readOnly', true);
	// 	}
	// })
	
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
		// document.getElementById("agreementcode").style.backgroundColor="#eee";
		// $( "#agreementcode" ).prop( "disabled", true );
		// $( "#vehicle" ).prop( "disabled", false );
		// $( "#driver" ).prop( "disabled", false );
		// $(document).ready(function() {
				$('#agreementcode').css("background-color", "#eee");
				$('#agreementcode').prop('readOnly', true);
				$('#agreementcode').val(null);
				$('#vehicle').css("background-color", "#fff");
				$('#vehicle').prop('readOnly', false);
				$('#driver').css("background-color", "#fff");
				$('#driver').prop('readOnly', false);

		 	$('#own').click(function() {
		  	if( $(this).val() === 'OWN') {
		   		$('#agreementcode').css("background-color", "#eee");
				$('#agreementcode').prop('readOnly', true);
				// $('#vehicle').css("background-color", "#fff");
				// $('#vehicle').prop('readOnly', false);
				$('#driver').css("background-color", "#fff");
				$('#driver').prop('readOnly', false);
		      	// $('#vehicle').val(null);
		      	$('#driver').val(null);
				// $( "#agreementcode" ).prop( "disabled", true );
				// $( "#vehicle" ).prop( "disabled", false );
				// $( "#driver" ).prop( "disabled", false );
		    } else {      
		    	$('#agreementcode').css("background-color", "#fff");
				$('#agreementcode').prop('readOnly', false);
				// $('#vehicle').css("background-color", "#eee");
				// $('#vehicle').prop('readOnly', true);
				$('#driver').css("background-color", "#eee");
				$('#driver').prop('readOnly', true);
				$('#agreementcode').val(null);
		  //   	$( "#agreementcode" ).prop( "disabled", false );
		  //   	$( "#vehicle" ).prop( "disabled", true );
		  // 	$( "#driver" ).prop( "disabled", true );
		    }
		  });
		 
		// });
		hitungrumus();
		getOvertime();
		cekspbtype();

	});
</script>
</body>
</html>
