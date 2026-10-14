<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Edit Header Sales Agreement</title>
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
<script>
  function hitungrumusheader() {
    var sdp_beforevat = $("#dp_beforevat").val();
    var sdp_vat = $("#dp_vat").val();
    var newsdp_beforevat = sdp_beforevat.replace(/,/g, '');
    
      //Untuk menghitung Amount DP Vat 
    totalamountdpvat = sdp_vat * newsdp_beforevat / 100;
    if (!isNaN(totalamountdpvat)) {
        $("#amount_dpvat").val(totalamountdpvat);
      }
      else{
        $("#amount_dpvat").val();
      }

      //Untuk menghitung Amount DP After Vat 
    totalamountdpaftervat = parseInt(newsdp_beforevat) + parseInt(totalamountdpvat) ;
    if (!isNaN(totalamountdpaftervat)) {
        $("#dp_aftervat").val(totalamountdpaftervat);
      }
      else{
        $("#dp_aftervat").val();
      }
  }
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

  '---- Data Header
 if Request("DATECREATED")<>"" then
   DATECREATED = Request("DATECREATED")
 else
   DATECREATED = textDate(date)
 end if
 xCode = Request("xCode")
 fcCode = Request("fcCode")
 remarks = Request("remarks")
 
 '---- Data Details

salescontractno =request("salescontractno")
customercode =request("customercode")
deliveryaddress =request("deliveryaddress")
curr =request("curr")
deliverytype =request("deliverytype")
dp_aftervat =request("dp_aftervat")
dp_beforevat =request("dp_beforevat")
dp_vat =request("dp_vat")
percentage_dp =request("percentage_dp")
totalamount =request("totalamount")
amount_dpvat =request("amount_dpvat")
deliverydate =request("deliverydate")
creationdate =request("creationdate")
startdate =request("startdate")
enddate =request("enddate")
quantity=request("quantity")
remarks=request("remarks")
paymentcondition=request("paymentcondition")
deliverycondition = Request("deliverycondition")
txtlineno = request("txtlineno")
productcode =request("productcode")
unitofmeasure =request("unitofmeasure")
salesprices =request("salesprices")
amount =request("amount")
vat =request("vat")
amountvat =request("amountvat")
totalamountvat =request("totalamountvat")
quantitytolerance =request("quantitytolerance")
quantityaftertolerance =request("quantityaftertolerance")
quantitytolerancetotal =request("quantitytolerancetotal")
invoicedpno =request("invoicedpno")
mill =request("mill")

fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
selectedcustomercode = Request("selectedcustomercode")
selecteddeliverytype = Request("selecteddeliverytype")

' End Added
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "EditHeader"   Then

  if Request("txtJenis")="EditHeader" then

    v_totalamount = 0
    qLineMax = " select nvl(SUM(totalamountvat),0) totalamountvat" & _
           " from salesagreementdet" & _
           " where salescontractno = '" & salescontractno & "'" & _
           " and fcba = '" & fcba & "'"       

    set rsqLineMax = Server.CreateObject("ADODB.Recordset")
    rsqLineMax.Open qLineMax, connstr
    if not rsqLineMax.eof then
      totalamountvat = cdbl(rsqLineMax("totalamountvat"))
    end if 

    qLineMax = " select nvl(SUM(dp_aftervat),0) dp_aftervat" & _
           " from salesagreementhdr" & _
           " where salescontractno = '" & salescontractno & "'" & _
           " and fcba = '" & fcba & "'"       

    set rsqLineMax = Server.CreateObject("ADODB.Recordset")
    rsqLineMax.Open qLineMax, connstr
    if not rsqLineMax.eof then
      v_dp_aftervat = cdbl(rsqLineMax("dp_aftervat"))
    end if 
    
    v_totalamount = v_totalamount + totalamountvat
    if cdbl(to_nol(v_totalamount))  = 0 then v_percentage_dp = 0 else v_percentage_dp = (cdbl(to_nol(dp_aftervat)) / v_totalamount) * 100
    ' response.write v_percentage_dp

    If v_dp_aftervat > v_totalamount Then
      Err.Number = 1001
      Err.Description = "DP Tidak boleh lebih dari Total Amount !"
      
      q = " delete from salesagreementdet " & _
        " where fcBA = '" & fcba & "'" & _
        " and salescontractno = '"& salescontractno & "'" 
        ' Response.Write q
        connStr.execute(q)
    Else      
      q = " Update salesagreementhdr" & _
          " set totalamount = replace('" & v_totalamount & "',',','')" & _
          "     , dp_beforevat =  replace('" & dp_beforevat & "',',','')" & _
          "     , dp_aftervat =  replace('" & dp_aftervat & "',',','')" & _
          "     , dp_vat = '" & dp_vat & "'" & _
		  "     , amount_dpvat = replace('" & amount_dpvat & "',',','')" & _
          "     , percentage_dp = '" & v_percentage_dp & "'" & _
		  "     , remarks = '" & remarks & "'" & _
          "     where salescontractno = '" & salescontractno & "'" & _
          "       and fcBA = '" & fcba & "'"
      ' response.write q
      connStr.execute(q)
    End if

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
 <form name="frmEntry" method="post" action="salesagreementheaderedit.Asp">
    <table class="data" width="100%">
        <tr> 
            <td align="left" width="20%">Down Payment Before Vat</td>
            <td align="left">
                <input type="text" maxlength="15" id="dp_beforevat" name="dp_beforevat" size="30" value="<%=dp_beforevat%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" onkeyup="hitungrumusheader();">
            </td>
        </tr>
        <tr>
            <td align="left" width="20%">Down Payment VAT (%)</td>
            <td align="left">
                <input type="text" maxlength="10" id="dp_vat" name="dp_vat" size="5" value="<%=dp_vat%>" autocomplete="Off" onkeyup="hitungrumusheader();">
                <input type="text" maxlength="30" id="amount_dpvat" name="amount_dpvat" size="18" value="<%=amount_dpvat%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly style="background-color:#eee;" onkeyup="hitungrumusheader();">
            </td>
        </tr>
        <tr>
            <td align="left" width="20%">Down Payment + VAT</td>
            <td align="left">
               <input type="text" maxlength="15" id="dp_aftervat" name="dp_aftervat" size="30" value="<%=dp_aftervat%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly style="background-color:#eee;">
            </td>
        </tr>
        <tr>
          <td align="left" width="20%">Total Amount Sales Contract</td>
          <td align="left">
              <input type="text" maxlength="15" name="totalamount" size="30" value="<%=totalamount%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly style="background-color:#eee;">
          </td>
        </tr>
        <tr>
            <td align="left" width="20%">% Down Payment</td>
            <td align="left">
                <input type="text" maxlength="15" name="percentage_dp" size="30" value="<%=percentage_dp%>" autocomplete="Off" readonly style="background-color:#eee;">
            </td>
        </tr>
		
		<tr>
            <td align="left" width="20%">Remarks</td>
            <td align="left">
                <textarea name="remarks" cols="50"><%=remarks%></textarea>
            </td>
        </tr>
		
        <tr><td colspan="2"  align="left">
        <input Type="Hidden" Value="True" Name="cmdView">
        <input type="hidden" name="fcba" value="<%=fcba%>">
        <input type="hidden" name="salescontractno" value="<%=salescontractno%>">
        <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
        <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
        <input Type="Hidden" Value="<%=fcCode%>" Name="fcCode">
       
        <button type="submit" class="button" name="txtJenis" value="EditHeader" onClick="return confirm('Are you sure to Update this data : <%=salescontractno%> ?');"><i class="fa fa-check"></i> Save</button>
        
        
        </td>
        </tr>
    </table>
</div>
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
