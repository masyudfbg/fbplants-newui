<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS -  Payment Voucher Approval Detail</title>
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
<script language="JavaScript" src="/js/date-picker.js"></script> 
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
var getMessage = function () {
	// build up the post string when passing variables to the server side page
	
	var fccode = document.getElementById("fccode").value;
	var PostStr = 'fccode='+fccode;
	// use the generic function to make the request
	doAJAXCall('ajax/getstockinfobyfcba.asp', 'POST', PostStr, showMessageResponse);
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);

    // update the Div to show the result from the server
	document.getElementById("fcname").value = s.fcname;
	document.getElementById("unitofmeasurement").value = s.unitofmeasurement;
}

function checkInput(element){ element.value = element.value.replace(/[^0-9\.]/g,'');}
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

.button.save {
    color: #000;
}

.button.save:hover,
.button.save:focus,
.button.save:active {
    border-color: #30b378;
    border-bottom-color: #30b378;
    color: #fff;
    background-color: #2ecc71;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#2ecc71), to(#30b378));
    background-image: -moz-linear-gradient(#2ecc71, #30b378);
    background-image: -ms-linear-gradient(#2ecc71, #30b378);
    background-image: -o-linear-gradient(#2ecc71, #30b378);
    background-image: linear-gradient(#2ecc71, #30b378);
}

.button.save:active,
.button.save.active {
    border-color: #30b378;
    border-bottom-color: #30b378;
    background-color: #30b378;
    background-image: -webkit-gradient(linear, 0 0, 0 100%, from(#30b378), to(#2ecc71));
    background-image: -moz-linear-gradient(#30b378, #2ecc71);
    background-image: -ms-linear-gradient(#30b378, #2ecc71);
    background-image: -o-linear-gradient(#30b378, #2ecc71);
    background-image: linear-gradient(#30b378, #2ecc71);
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
 <!-------------awal!------------->
 <%
 Dim cmdView,PrCode,PrDate
 Dim Description, Remarks, RequestorId
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim ManagedId, PurchaseGroupId, RowStateId
 Dim LineId, ItemCode, RequestedQty, ApprovedQty, OrderedQty, DescriptionDetail
 Dim UnitOfMeasurement, WantedDeliveryDate
 Dim strItemCodeInfo, rsstrItemCodeInfo
 Dim rsPurchase, rsPurchaseTemp
 Dim strQueryStatus, rsQueryStatus
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 dim txtActivation

Dim txtPRCode : txtPRCode = 0
Dim txtValidStock : txtValidStock = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0

fcba = request("fcba")
authorizelevelcode = request("authorizelevelcode")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 
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
 vouchercode = request ("vouchercode")
 cekno = request ("cekno")
 referenceno = request ("referenceno")
 jobcode = request ("jobcode")
 locationtype = request ("locationtype")
 locationcode = request ("locationcode")
 invoicecode = request ("invoicecode")
 agreementcode = request ("agreementcode")
 padifor = request ("padifor")
 chargejob = request ("chargejob")
 chargetype = request ("chargetype")
 chargecode = request ("chargecode")
 

fcba = request("fcba")
authorizelevelcode = request("authorizelevelcode")
authorizetemplatecode = request("authorizetemplatecode")

Function CheckAvailableJobZerorise(initialJob, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from parameter" & _
	                " where upper(parcode) = upper('JOBCODE_ZERORISE')" & _
					"       and fccode = '" & Trim(initialJob) & "'" & _
					"       and fcba = '" & Trim(initialFcba) & "'"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckAvailableJobZerorise = initialAvailable
End Function


Function f_i_check(in_fcba, in_vouchercode, in_datecreated, in_bankcode, in_jobcode, in_targettype, in_targetcode, in_documentno, in_amount)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT validasi_voucher_api.check_data( '"&in_fcba&"', '"&in_vouchercode&"', TO_DATE('"&in_datecreated&"', 'mm/dd/yyyy'), '"&in_bankcode&"', '"&in_jobcode&"', '"&in_targettype&"', '"&in_targetcode&"', '"&in_documentno&"', REPLACE('"&in_amount&"', ',', '') ) result" & _
	                " FROM DUAL"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("result")
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	f_i_check = initialAvailable
End Function


Function CheckCostJobCr(statusxAmount ,initialJob, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from Costjobcr" & _
	                " where  jobcode = '" & Trim(initialJob) & "'" & _
					" and fcba = '" & Trim(initialFcba) & "'"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	IF statusxAmount = "MINUS"  THEN
			If Not rsAvailable.EOF Then
				initialAvailable = Cdbl(rsAvailable("Total"))
			Else
				initialAvailable = 0
			End If
	 Else
	    initialAvailable = 0
	 END IF
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckCostJobCr = initialAvailable
End Function

Function CheckCharge(initType, initCode, initJob, initFcba, initDate, initSource)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT F_CekTarget_byDate ('"& initType &"'," & _
					"                           '"& initCode &"'," & _
					"                           '"& initJob &"'," & _
					"                           '"& initFcba &"'," & _
					"                           TO_DATE ('"& initDate &"', 'MM/DD/YYYY')," & _
					"                           'BA') Total" & _
					"  FROM DUAL"

	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("Total")
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckCharge = initialAvailable
End Function

 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "ApprovalAction" or Request("txtJenis")= "SaveData" or Request("txtJenis")= "RejectAction" Then
	check_data = f_i_check(Session("fcBA"), fcCode, DATECREATED, xCode, Jobcode, locationtype, locationcode, documentno, amount) 
	
	if Request("txtJenis")="ApprovalAction" then
		if check_data = "Success" then 
			invoicecode = ""
			agreementcode = ""
		
			if locationtype = "AP" then
				invoicecode = documentno
			elseif locationtype = "CA" then
				agreementcode = documentno
			else
				invoicecode = ""
				agreementcode = ""
			end if 
      
      qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& jobcode &"' group by fcba"     
      ' response.write qLineMax
      set rsqLineMax = Server.CreateObject("ADODB.Recordset")
      rsqLineMax.Open qLineMax, connstr
      if not rsqLineMax.eof then
        code = cdbl(to_nol(rsqLineMax("fccode")))
        businessunit = rsqLineMax("fcba")
      end if 
      
      strLineMax = " select nvl(max(lineno),0) maxLineNoKey from paymentvoucher_hist where fcBA='" & fcba & "' and vouchercode = '"& vouchercode &"'"

      Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
      rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
      
      if not rsstrLineMax.eof then
        tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
      end if
            
      rsstrLineMax.Close
      set rsstrLineMax = Nothing 


			if code = 0 then
				if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 ) then
					err.number = 1001
					err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
				else 
					For n=1 to (cdbl(request("no"))-1)  

						q = " UPDATE paymentvoucherdetail" & _
						    " SET jobcode = '"& request("jobcode_"&TRIM(n)) & "'" & _
						    "	 ,locationtype = '"& request("locationtype_"&TRIM(n)) & "'" & _
						    "	 ,locationcode = '"& request("locationcode_"&TRIM(n)) & "'" & _
						    "	 ,chargejob = '"& request("chargejob_"&TRIM(n)) & "'" & _
						    "	 ,chargetype = '"& request("chargetype_"&TRIM(n)) & "'" & _
						    "	 ,chargecode = '"& request("chargecode_"&TRIM(n)) & "'" & _
							"  ,amount = '"& request("amount_"&TRIM(n)) & "'" & _
							"  ,invoicecode = '"& request("invoicecode_"&TRIM(n)) & "'" & _
              "  ,agreementcode = '"& request("agreementcode_"&TRIM(n)) & "'" & _
							"  ,remarks = '"& request("remarks_"&TRIM(n)) & "'" & _
						  	" WHERE vouchercode = '" & vouchercode & "'" & _
							"       AND lineno = '"& request("lineno_"&TRIM(n)) & "'" & _ 
						    "       AND fcba = '"& fcba & "'"
						' response.write q
						connStr.execute(q)

            set rsPurchase = server.CreateObject("Adodb.recordset")
               q= " select pvd.vouchercode,pv.referenceno, pv.cekno, sum(pvd.amount) amount, pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification from paymentvoucherdetail pvd" & _
                  " inner join paymentvoucher pv on (pvd.vouchercode = pv.vouchercode and pvd.fcba = pv.fcba)" & _
                  " where pvd.vouchercode='" & vouchercode & "' AND pvd.fcba = '"&fcba&"'" & _
                  " group by pvd.vouchercode,pv.referenceno, pv.cekno,  pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification   "

                ' response.write q
               rsPurchase.Open q, connstr
               if Not rsPurchase.Eof then
                vouchercode=rsPurchase("vouchercode")
                  referenceno=rsPurchase("referenceno")
                  cekno=rsPurchase("cekno")
                  amount=rsPurchase("amount")
                  datecreated=rsPurchase("datecreated")
                  remarks=rsPurchase("remarks")
                  rowstateid=rsPurchase("rowstateid")
                  document_classification=rsPurchase("document_classification")
               End if

            q = " UPDATE paymentvoucher " & _
                " SET  totalamount = '"& amount & "'" & _  
                " WHERE vouchercode = '" & vouchercode & "'" & _ 
                "       AND fcba = '"& fcba & "'"
            ' response.write q
            connStr.execute(q)

					next 
				end if
			else 
				if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 ) then
					err.number = 1001
					err.description = "Anda harus input Charge Job, Charge Type dan Charge Code! "
				else
					Dim val_Charge : val_Charge = 0
					val_Charge = CheckCharge(chargetype, chargecode, chargejob, businessunit, DATECREATED, sourcetype)
					If val_Charge <> "Success" Then
						Err.Number = 1001
						Err.Description = val_Charge
					Else
						For n=1 to (cdbl(request("no"))-1) 

						q = " UPDATE paymentvoucherdetail" & _
						    " SET jobcode = '"& request("jobcode_"&TRIM(n)) & "'" & _
						    "	 ,locationtype = '"& request("locationtype_"&TRIM(n)) & "'" & _
						    "	 ,locationcode = '"& request("locationcode_"&TRIM(n)) & "'" & _
						    "	 ,chargejob = '"& request("chargejob_"&TRIM(n)) & "'" & _
						    "	 ,chargetype = '"& request("chargetype_"&TRIM(n)) & "'" & _
						    "	 ,chargecode = '"& request("chargecode_"&TRIM(n)) & "'" & _
							"  ,amount = '"& request("amount_"&TRIM(n)) & "'" & _
							"  ,invoicecode = '"& request("invoicecode_"&TRIM(n)) & "'" & _
              "  ,agreementcode = '"& request("agreementcode_"&TRIM(n)) & "'" & _
							"  ,remarks = '"& request("remarks_"&TRIM(n)) & "'" & _
							" WHERE vouchercode = '" & vouchercode & "'" & _
							"       AND lineno = '"& request("lineno_"&TRIM(n)) & "'" & _ 
						    "       AND fcba = '"& fcba & "'"
						' response.write q
						connStr.execute(q)

            set rsPurchase = server.CreateObject("Adodb.recordset")
               q= " select pvd.vouchercode,pv.referenceno, pv.cekno, sum(pvd.amount) amount, pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification from paymentvoucherdetail pvd" & _
                  " inner join paymentvoucher pv on (pvd.vouchercode = pv.vouchercode and pvd.fcba = pv.fcba)" & _
                  " where pvd.vouchercode='" & vouchercode & "' AND pvd.fcba = '"&fcba&"'" & _
                  " group by pvd.vouchercode,pv.referenceno, pv.cekno,  pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification   "

                ' response.write q
               rsPurchase.Open q, connstr
               if Not rsPurchase.Eof then
                vouchercode=rsPurchase("vouchercode")
                  referenceno=rsPurchase("referenceno")
                  cekno=rsPurchase("cekno")
                  amount=rsPurchase("amount")
                  datecreated=rsPurchase("datecreated")
                  remarks=rsPurchase("remarks")
                  rowstateid=rsPurchase("rowstateid")
                  document_classification=rsPurchase("document_classification")
               End if

            q = " UPDATE paymentvoucher " & _
                " SET  totalamount = '"& amount & "'" & _  
                " WHERE vouchercode = '" & vouchercode & "'" & _ 
                "       AND fcba = '"& fcba & "'"
            ' response.write q
            connStr.execute(q)
						next
					End If
				End If
			End If
						
						q = " INSERT INTO paymentvoucher_hist ( vouchercode" & _
  							"                                           , currentauthorizelevel" & _ 
  							"                                           , fcentry" & _
  							"                                           , fcedit" & _
  							"                                           , fcip" & _
  							"                                           , fcba" & _ 
  							"                                           , lastupdate" & _
  							"                                           , lasttime)"& _
  							" VALUES ( '" & vouchercode & "'" & _
  							"          , '" & authorizelevelcode & "'" & _ 
  							"          , '" & Session("CodeName") & "'" & _
  							"          , '" & Session("CodeName") & "'" & _
  							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
  							"          , '" & fcba & "'" & _ 
  							"          , SYSDATE" & _
  							"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
  						' response.write q
  						connStr.execute(q) 
						
						q = " INSERT INTO pv_authorization_comments ( vouchercode" & _
							"                                         , currentauthorizelevel" & _
							"                                         , comments" & _
							"                                         , fcentry" & _
							"                                         , fcedit" & _
							"                                         , fcip" & _
							"                                         , fcba" & _
							"                                         , lastupdate" & _
							"                                         , lasttime)"& _
							" VALUES ( '" & vouchercode & "'" & _
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
						q = " SELECT COUNT(1) total FROM authorizelevel al" & _
							" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
							" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
							" WHERE at.document_type = 'PV'" & _
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
						
						if is_final > 0 then
							
							strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & fcba & "','" & vouchercode & "', 'Approve' ); END; "
							connStr.execute(strQueryStatus)
							
							strQueryStatus = "UPDATE paymentvoucher SET rowstateid = 'Approved' WHERE fcba = '" & fcba & "' AND vouchercode = '" & vouchercode & "'"
							connStr.execute(strQueryStatus)
							
						end if
			else
				err.number = 1001
				err.description = check_data
			End If
		cmdView	= "True"

    elseif Request("txtJenis")= "SaveData" then
    connstr.BeginTrans
        if check_data = "Success" then 
      invoicecode = ""
      agreementcode = ""
    
      if locationtype = "AP" then
        invoicecode = documentno
      elseif locationtype = "CA" then
        agreementcode = documentno
      else
        invoicecode = ""
        agreementcode = ""
      end if 
      
      qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& jobcode &"' group by fcba"     
      ' response.write qLineMax
      set rsqLineMax = Server.CreateObject("ADODB.Recordset")
      rsqLineMax.Open qLineMax, connstr
      if not rsqLineMax.eof then
        code = cdbl(to_nol(rsqLineMax("fccode")))
        businessunit = rsqLineMax("fcba")
      end if 
      
      strLineMax = " select nvl(max(lineno),0) maxLineNoKey from paymentvoucher_hist where fcBA='" & fcba & "' and vouchercode = '"& vouchercode &"'"

      Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
      rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
      
      if not rsstrLineMax.eof then
        tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
      end if
            
      rsstrLineMax.Close
      set rsstrLineMax = Nothing 


      if code = 0 then
        if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 ) then
          err.number = 1001
          err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
        else 
          For n=1 to (cdbl(request("no"))-1)  

            q = " UPDATE paymentvoucherdetail" & _
                " SET jobcode = '"& request("jobcode_"&TRIM(n)) & "'" & _
                "  ,locationtype = '"& request("locationtype_"&TRIM(n)) & "'" & _
                "  ,locationcode = '"& request("locationcode_"&TRIM(n)) & "'" & _
                "  ,chargejob = '"& request("chargejob_"&TRIM(n)) & "'" & _
                "  ,chargetype = '"& request("chargetype_"&TRIM(n)) & "'" & _
                "  ,chargecode = '"& request("chargecode_"&TRIM(n)) & "'" & _
                "  ,amount = '"& request("amount_"&TRIM(n)) & "'" & _
                "  ,invoicecode = '"& request("invoicecode_"&TRIM(n)) & "'" & _
                "  ,agreementcode = '"& request("agreementcode_"&TRIM(n)) & "'" & _
                "  ,remarks = '"& request("remarks_"&TRIM(n)) & "'" & _
                " WHERE vouchercode = '" & vouchercode & "'" & _
                "       AND lineno = '"& request("lineno_"&TRIM(n)) & "'" & _ 
                "       AND fcba = '"& fcba & "'"
            ' response.write q
            connStr.execute(q)

            set rsPurchase = server.CreateObject("Adodb.recordset")
               q= " select pvd.vouchercode,pv.referenceno, pv.cekno, sum(pvd.amount) amount, pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification from paymentvoucherdetail pvd" & _
                  " inner join paymentvoucher pv on (pvd.vouchercode = pv.vouchercode and pvd.fcba = pv.fcba)" & _
                  " where pvd.vouchercode='" & vouchercode & "' AND pvd.fcba = '"&fcba&"'" & _
                  " group by pvd.vouchercode,pv.referenceno, pv.cekno,  pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification   "

                ' response.write q
               rsPurchase.Open q, connstr
               if Not rsPurchase.Eof then
                vouchercode=rsPurchase("vouchercode")
                  referenceno=rsPurchase("referenceno")
                  cekno=rsPurchase("cekno")
                  amount=rsPurchase("amount")
                  datecreated=rsPurchase("datecreated")
                  remarks=rsPurchase("remarks")
                  rowstateid=rsPurchase("rowstateid")
                  document_classification=rsPurchase("document_classification")
               End if

            q = " UPDATE paymentvoucher " & _
                " SET  totalamount = '"& amount & "'" & _  
                " WHERE vouchercode = '" & vouchercode & "'" & _ 
                "       AND fcba = '"& fcba & "'"
            ' response.write q
            connStr.execute(q)

          next 
        end if
      else 
        if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 ) then
          err.number = 1001
          err.description = "Anda harus input Charge Job, Charge Type dan Charge Code! "
        else 
          val_Charge = CheckCharge(chargetype, chargecode, chargejob, businessunit, DATECREATED, sourcetype)
          If val_Charge <> "Success" Then
            Err.Number = 1001
            Err.Description = val_Charge
          Else
            For n=1 to (cdbl(request("no"))-1) 

            q = " UPDATE paymentvoucherdetail" & _
                " SET jobcode = '"& request("jobcode_"&TRIM(n)) & "'" & _
                "  ,locationtype = '"& request("locationtype_"&TRIM(n)) & "'" & _
                "  ,locationcode = '"& request("locationcode_"&TRIM(n)) & "'" & _
                "  ,chargejob = '"& request("chargejob_"&TRIM(n)) & "'" & _
                "  ,chargetype = '"& request("chargetype_"&TRIM(n)) & "'" & _
                "  ,chargecode = '"& request("chargecode_"&TRIM(n)) & "'" & _
                "  ,amount = '"& request("amount_"&TRIM(n)) & "'" & _
                "  ,invoicecode = '"& request("invoicecode_"&TRIM(n)) & "'" & _
                "  ,agreementcode = '"& request("agreementcode_"&TRIM(n)) & "'" & _
                "  ,remarks = '"& request("remarks_"&TRIM(n)) & "'" & _
              " WHERE vouchercode = '" & vouchercode & "'" & _
              "       AND lineno = '"& request("lineno_"&TRIM(n)) & "'" & _ 
                "       AND fcba = '"& fcba & "'"
            ' response.write q
            connStr.execute(q)

            set rsPurchase = server.CreateObject("Adodb.recordset")
               q= " select pvd.vouchercode,pv.referenceno, pv.cekno, sum(pvd.amount) amount, pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification from paymentvoucherdetail pvd" & _
                  " inner join paymentvoucher pv on (pvd.vouchercode = pv.vouchercode and pvd.fcba = pv.fcba)" & _
                  " where pvd.vouchercode='" & vouchercode & "' AND pvd.fcba = '"&fcba&"'" & _
                  " group by pvd.vouchercode,pv.referenceno, pv.cekno,  pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification   "

                ' response.write q
               rsPurchase.Open q, connstr
               if Not rsPurchase.Eof then
                vouchercode=rsPurchase("vouchercode")
                  referenceno=rsPurchase("referenceno")
                  cekno=rsPurchase("cekno")
                  amount=rsPurchase("amount")
                  datecreated=rsPurchase("datecreated")
                  remarks=rsPurchase("remarks")
                  rowstateid=rsPurchase("rowstateid")
                  document_classification=rsPurchase("document_classification")
               End if

            q = " UPDATE paymentvoucher " & _
                " SET  totalamount = '"& amount & "'" & _  
                " WHERE vouchercode = '" & vouchercode & "'" & _ 
                "       AND fcba = '"& fcba & "'"
            ' response.write q
            connStr.execute(q)
            next
          End If
        End If
      End If
    else
      
    End If
    
    if err.number  = 0 then
      connStr.CommitTrans
    end if
      
    cmdView = "True"

	elseif Request("txtJenis")="RejectAction" then
			q = " INSERT INTO pv_authorization_comments ( vouchercode" & _
			"                                         , currentauthorizelevel" & _
			"                                         , comments" & _
			"                                         , fcentry" & _
			"                                         , fcedit" & _
			"                                         , fcip" & _
			"                                         , fcba" & _
			"                                         , lastupdate" & _
			"                                         , lasttime)"& _
			" VALUES ( '" & vouchercode & "'" & _
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
		
		
	    strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & fcba & "','" & vouchercode & "', 'Reject' ); END; "
		connStr.execute(strQueryStatus)
		
		cmdView	= "True"
    else
        
	end if								 
	cmdView	= "True"
 else
    cmdView	= request("cmdView")
 End if

 set rsPurchase = server.CreateObject("Adodb.recordset")
 q=	" select pvd.vouchercode,pv.referenceno, pv.cekno, sum(pvd.amount) amount, pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification " & _
    " , accountingyear||LPAD(accountingperiod,2,0) accountingyearperiod" & _
	" , accountingperiod" & _
	" , accountingyear" & _
	" from paymentvoucherdetail pvd" & _
	" inner join paymentvoucher pv on (pvd.vouchercode = pv.vouchercode and pvd.fcba = pv.fcba)" & _
    " where pvd.vouchercode='" & vouchercode & "' AND pvd.fcba = '"&fcba&"'" & _
    " group by pvd.vouchercode,pv.referenceno, pv.cekno,  pv.datecreated, pv.remarks, pv.rowstateid,pv.document_classification, accountingyear||LPAD(accountingperiod,2,0) ,accountingyear, accountingperiod  "

  ' response.write q
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
 	vouchercode=rsPurchase("vouchercode")
    referenceno=rsPurchase("referenceno")
    cekno=rsPurchase("cekno")
    amount=rsPurchase("amount")
    datecreated=rsPurchase("datecreated")
    remarks=rsPurchase("remarks")
    rowstateid=rsPurchase("rowstateid")
    document_classification=rsPurchase("document_classification")
	ACCOUNTINGYEAR =rsPurchase("ACCOUNTINGYEAR")
	ACCOUNTINGPERIOD =rsPurchase("ACCOUNTINGPERIOD")
	ACCOUNTINGYEARPERIOD = rsPurchase("accountingyearperiod")
 End if
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<div class='error'>" & err.description & "</div>"
End if

has_approveddata = 0
q = " SELECT COUNT(1) total FROM paymentvoucher_hist WHERE vouchercode = '" & vouchercode & "' AND fcba = '" & fcba & "'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approveddata = cdbl(to_nol(rsPurchase("total")))
End if

has_approvedby = 0
q = " SELECT COUNT(1) total FROM paymentvoucher_hist WHERE vouchercode = '" & vouchercode & "' AND fcba = '" & fcba & "' AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approvedby = cdbl(to_nol(rsPurchase("total")))
End if

q = " SELECT * FROM pv_authorization_comments" & _
    " WHERE vouchercode = '" & vouchercode & "'" & _
    "       AND fcba = '" & fcba & "'" & _
	"       AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	comments = rsPurchase("comments")
End if
 '-------------ISI Payment Form
%>
  <form name="frmEntry" method="post" action="PaymentVoucherApproval.Asp">
	<table class="data" width="100%">
		<tr>
		  <td align="left">Voucher No</td>
		  <td align="left"><input type="text" maxlength="15" name="vouchercode" size="30" value="<%=vouchercode%>"><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
		  <td align="left">Reference No</td>
		  <td align="left"><input type="text" maxlength="15" name="referenceno" size="30" value="<%=referenceno%>"></td>
		</tr>
		<tr>
		  <td align="left">Cek No</td>
		  <td align="left"><input type="text" maxlength="15" name="cekno" size="30" value="<%=cekno%>"></td>
		</tr>
		<tr>
		   <td align="left">Date</td>
		   <td align="left"><input class="datepicker" type="text" maxlength="10" name="datecreated" size="10" value="<%=datecreated%>"><font color="#FF0000">&nbsp;*</font>
		</td>
		</tr>
		<tr>
			<td align="left">Accounting Period</td>
			<td align="left">
				<select name="ACCOUNTINGYEARPERIOD" class="chosen-select" style="width: 50%">
				<option></option>
				<%
				qLov = 	"   SELECT pv.accountingyear || LPAD (pv.accountingperiod, 2, 0)" & _
						"             ACCOUNTINGYEARPERIOD," & _
						"          description" & _
						"     FROM paymentvoucher pv" & _
						"          LEFT JOIN accountingyearperiod ap" & _
						"             ON (    ap.accountingyear = pv.accountingyear" & _
						"                 AND ap.accountingperiod = pv.accountingperiod" & _
						"                 AND ap.fcba = pv.fcba)" & _
						"    WHERE vouchercode = '"& vouchercode &"' AND pv.fcba = '"& fcba &"'" & _
						" ORDER BY vouchercode"  
						
				Set rsLov = Server.CreateObject("ADODB.Recordset")
				rsLov.Open qLov, connstr
				if not rsLov.eof then
					while not rsLov.eof
						If Request("ACCOUNTINGYEARPERIOD") = rsLov("ACCOUNTINGYEARPERIOD") Then 
							Data_ACCOUNTINGYEARPERIOD = "true"
						else
							Data_ACCOUNTINGYEARPERIOD = "false"
						End if
				%>
					<option value="<%=rsLov("ACCOUNTINGYEARPERIOD")%>" <% If Request("ACCOUNTINGYEARPERIOD") = rsLov("ACCOUNTINGYEARPERIOD") Then %> Selected <%End if%>><%=rsLov("Description")%></option> 
				<%
						rsLov.movenext
					wend
				if Data_ACCOUNTINGYEARPERIOD = "false" then %>
					<option value="<%=ACCOUNTINGYEAR&ACCOUNTINGPERIOD%>" Selected>Tahun <%=ACCOUNTINGYEAR%> Bulan <%=ACCOUNTINGPERIOD%></option>
				<% end if %>
				
				<%
					End if
					rsLov.Close
					set rsLov = Nothing
					q=""
					%>
				</select>
				<font color="#FF0000">*</font>
			</td>	
		</tr>
		<tr>
		  <td align="left">Paid For</td>
		  <td align="left"><textarea name="padifor	" cols="50"><%=remarks%></textarea></td>
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
		  <td align="left">Amount</td>
		  <td align="left"><input type="text" maxlength="15" name="amount" size="30" value="<%=amount%>"></td>
		</tr>
		<tr>
		  <td align="left">Status</td>
		  <td align="left"><input type="text" maxlength="50" name="rowstateid" size="30" value="<%=rowstateid%>" style="background-color:#eee;" readonly><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
		<td align="left">Comments</td>
		<td align="left"><textarea name="comments" cols="50"><%=comments%></textarea>
		</td>
		</tr>
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Hidden" Value="<%=fcba%>" Name="fcba">
		<input Type="Hidden" Value="<%=authorizelevelcode%>" Name="authorizelevelcode">
		<input Type="Hidden" Value="<%=authorizetemplatecode%>" Name="authorizetemplatecode">
		<%if RowStateId <> "Rejected" then%>
			<%if has_approvedby=0 then%>
        <button type="submit" class="button save" name="txtJenis" value="SaveData" onClick="return confirm('Are you sure to Save this data : <%=vouchercode%> ?');"><i class="fa fa-save"></i> Save</button>
				<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Are you sure to Approve this data : <%=vouchercode%> ?');"><i class="fa fa-check"></i> Approve</button>
				<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Are you sure to Reject this data : <%=vouchercode%> ?');"><i class="fa fa-times"></i> Reject</button>
			<%else%>
				<b>This Voucher number : <%=vouchercode%> has been Approved</b>
			<%end if%>
		<%else%>
			<b>This Voucher number : <%=vouchercode%> has been Rejected</b>
		<%end if%>
		
		</td>
		</tr>
    </table>
<%

	q = " SELECT pvd.vouchercode," & _
		"       pvd.jobcode," & _
		"       jb.fcname jobname," & _
		"       pvd.locationtype," & _
		"       pvd.locationcode," & _
		"       pvd.chargejob," & _
		"       pvd.chargetype," & _
		"       pvd.chargecode," & _
		"       pvd.invoicecode," & _
		"       pvd.agreementcode," & _
    "       pvd.amount," & _
		"       pvd.remarks," & _ 
		"       pvd.lineno" & _
		"  FROM paymentvoucherdetail pvd" & _
		"       INNER JOIN job jb ON (jb.fccode = pvd.jobcode AND jb.fcba = pvd.fcba)" & _
		" WHERE pvd.vouchercode = '"& vouchercode &"' AND pvd.fcBA = '"& fcba &"' ORDER BY lineno " 

	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.Open q, connstr
%>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="left">Voucher No</TD> 
	<%if RowStateId <> "Rejected" then%>
		<%if has_approvedby=0 then%>
			<TD align="left">Jobcode Original</TD>	
			<TD align="left">Jobcode</TD>	 
			<TD align="left">Location Type</TD>	
			<TD align="left">Location Code</TD>
			<TD align="left">Charge Job</TD>	 
			<TD align="left">Charge Type</TD>	
			<TD align="left">Charge Code</TD>	
      <TD align="left">Amount</TD> 
      <TD align="left">Invoice Code</TD>
      <TD align="left">Agreement Code</TD> 
      <TD align="left">Remarks</TD> 
		<%else%>
			<TD align="left">Jobcode</TD>	 
			<TD align="left">Location Type</TD>	
			<TD align="left">Location Code</TD>	
			<TD align="left">Charge Job</TD>	 
			<TD align="left">Charge Type</TD>	
			<TD align="left">Charge Code</TD>
      <TD align="left">Amount</TD> 
      <TD align="left">Invoice Code</TD>
      <TD align="left">Agreement Code</TD>  
      <TD align="left">Remarks</TD> 
		<%end if%>
	<%end if%> 
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
				<td align="left" title="<%=rsPurchaseOrder("vouchercode").Name%>" valign="top"><%=rsPurchaseOrder("vouchercode")%></td> 
				<%if RowStateId <> "Rejected" then%>
  				<%if has_approvedby=0 then%>
  					<td align="left" title="<%=rsPurchaseOrder("jobcode").Name%>" valign="top"><%=rsPurchaseOrder("jobcode")%> : <%=rsPurchaseOrder("jobname")%></td>
  					<td align="right" title="Jobcode" valign="top">
  						<%jobcode = "0":if has_approveddata = 0 then jobcode = rsPurchaseOrder("JobCode") else jobcode = rsPurchaseOrder("JobCode")%>
  						<input type="text" style="text-align:right;" name="jobcode_<%=no%>" width="2" value="<%=jobcode%>" > 
  						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
  					<font face="Verdana" size="1">Activity</font>
  					</td> 
  					<td align="right" title="locationtype" valign="top">
  						<%locationtype = "0":if has_approveddata = 0 then locationtype = rsPurchaseOrder("locationtype") else locationtype = rsPurchaseOrder("locationtype")%>
  						<input type="text" style="text-align:right;" name="locationtype_<%=no%>" width="2" value="<%=locationtype%>" > 
  						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                              <font face="Verdana" size="1">Location Type</font></a>
  					</td>
  					<td align="right" title="locationcode" valign="top">
  						<%locationcode = "0":if has_approveddata = 0 then locationcode = rsPurchaseOrder("locationcode") else locationcode = rsPurchaseOrder("locationcode")%>
  						<input type="text" style="text-align:right;" name="locationcode_<%=no%>" width="2" value="<%=locationcode%>" > 
  					</td>

  					<td align="right" title="chargejob" valign="top">
  						<%chargejob = "0":if has_approveddata = 0 then chargejob = rsPurchaseOrder("chargejob") else chargejob = rsPurchaseOrder("chargejob")%>
  						<input type="text" style="text-align:right;" name="chargejob_<%=no%>" width="2" value="<%=chargejob%>" > 
  						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
  					<font face="Verdana" size="1">Activity</font>
  					</td> 
  					<td align="right" title="chargetype" valign="top">
  						<%chargetype = "0":if has_approveddata = 0 then chargetype = rsPurchaseOrder("chargetype") else chargetype = rsPurchaseOrder("chargetype")%>
  						<input type="text" style="text-align:right;" name="chargetype_<%=no%>" width="2" value="<%=chargetype%>" > 
  						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                              <font face="Verdana" size="1">Location Type</font></a>
  					</td>
  					<td align="right" title="chargecode" valign="top">
  						<%chargecode = "0":if has_approveddata = 0 then chargecode = rsPurchaseOrder("chargecode") else chargecode = rsPurchaseOrder("chargecode")%>
  						<input type="text" style="text-align:right;" name="chargecode_<%=no%>" width="2" value="<%=chargecode%>" > 
  					</td>
            <td align="right" valign="top" >
              <%amount = "0":if has_approveddata = 0 then amount = rsPurchaseOrder("amount") else amount = rsPurchaseOrder("amount")%>
              <input type="text" style="text-align:right;" name="amount_<%=no%>" width="2" value="<%=amount%>" > 
            </td>
            <td align="right" valign="top" >
              <%invoicecode = "0":if has_approveddata = 0 then invoicecode = rsPurchaseOrder("invoicecode") else invoicecode = rsPurchaseOrder("invoicecode")%>
              <input type="text" style="text-align:right;" name="invoicecode_<%=no%>" width="2" value="<%=invoicecode%>" > 
            </td>
            <td align="right" valign="top" >
              <%agreementcode = "0":if has_approveddata = 0 then agreementcode = rsPurchaseOrder("agreementcode") else agreementcode = rsPurchaseOrder("agreementcode")%>
              <input type="text" style="text-align:right;" name="agreementcode_<%=no%>" width="2" value="<%=agreementcode%>" > 
            </td>
            <td align="right" valign="top" >
              <%remarks = "0":if has_approveddata = 0 then remarks = rsPurchaseOrder("remarks") else remarks = rsPurchaseOrder("remarks")%>
              <input type="text" style="text-align:right;" name="remarks_<%=no%>" width="2" value="<%=remarks%>" > 
            </td>
  					<input type="hidden" name="lineno_<%=no%>" value="<%=rsPurchaseOrder("lineno")%>"> 
  				<%else%>
  					<td align="left" title="<%=rsPurchaseOrder("jobcode").Name%>" valign="top"><%=rsPurchaseOrder("jobcode")%> : <%=rsPurchaseOrder("jobname")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("locationtype").Name%>" valign="top"><%=rsPurchaseOrder("locationtype")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("locationcode").Name%>" valign="top"><%=rsPurchaseOrder("locationcode")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("chargejob").Name%>" valign="top"><%=rsPurchaseOrder("chargejob")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("chargetype").Name%>" valign="top"><%=rsPurchaseOrder("chargetype")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("chargecode").Name%>" valign="top"><%=rsPurchaseOrder("chargecode")%></td>
            <td align="left" title="<%=rsPurchaseOrder("amount").Name%>" valign="top"><%=rsPurchaseOrder("amount")%></td>
            <td align="left" title="<%=rsPurchaseOrder("invoicecode").Name%>" valign="top"><%=rsPurchaseOrder("invoicecode")%></td>
            <td align="left" title="<%=rsPurchaseOrder("agreementcode").Name%>" valign="top"><%=rsPurchaseOrder("agreementcode")%></td>
            <td align="left" title="<%=rsPurchaseOrder("remarks").Name%>" valign="top"><%=rsPurchaseOrder("remarks")%></td>
  				<%end if%>
				<%end if%> 
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
 </br>
  <tr>
      <input type="hidden" name="no" value="<%=no%>">
      <input type="hidden" name="jobcode" value="<%=jobcode%>">
      <input type="hidden" name="chargejob" value="<%=chargejob%>">
      <input type="hidden" name="chargetype" value="<%=chargetype%>">
      <input type="hidden" name="chargecode" value="<%=chargecode%>"> 
  </tr>
  </form>

</div></div>
<!--#include file="bottom.jsp"-->
<script>

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
	getMessage();
});
</script>
</body>
</html>
