<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS -  Attendance & Activity Approval Detail</title>
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

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 '---- Data Header
 if Request("fddate")<>"" then
   fddate   =Request("fddate")
 else
   fddate   =textDate(date)
 end if
 
GANGCODE=request("GANGCODE")
fddate =request("fdDate")
EMPLOYEECODE=request("EMPLOYEECODE")
SUPERVISION_1=request("supervision_1")
SUPERVISION_2=request("SUPERVISION_2")
SUPERVISION_3=request("SUPERVISION_3")
SUPERVISION_4=request("SUPERVISION_4")
ATTENDANCE=request("ATTENDANCE")
JOBCODE=request("JOBCODE")
locationtype=request("locationtype")
LOCATIONCODE=request("LOCATIONCODE")
MANDAYS=To_Nol(request("MANDAYS"))
UNIT=To_Nol(request("UNIT"))
UNITHARVESTING=To_Nol(request("UNITHARVESTING"))
OUTPUT=To_Nol(request("OUTPUT"))
OTHRS=To_Nol(request("OTHRS"))
RATE=To_Nol(request("RATE"))
REFERENCE=request("REFERENCE")
REMARKS=request("REMARKS")
GangType=Request("GangType")
TOTMANDAYS="0"
recordAda="F"
BUCKET=Request("BUCKET")
chitno=Request("chitno")
kg_janjang=Request("kg_janjang")
kg_brondolan=Request("kg_brondolan")
spbno=Request("spbno")
RowState=Request("RowState")

' Added Variable Harvestig Quality By Sendy 09/24/2018
UNDER_RIPE=To_Nol(request("UNDER_RIPE"))
OVER_RIPE=To_Nol(request("OVER_RIPE"))
ABNORMAL=To_Nol(request("ABNORMAL"))
LONG_STALK=To_Nol(request("LONG_STALK"))
EATEN_BY_RAT=To_Nol(request("EATEN_BY_RAT"))
UNHARVEST_FFB=To_Nol(request("UNHARVEST_FFB"))
UNCOLLECT_LF_CIRCLE=To_Nol(request("UNCOLLECT_LF_CIRCLE"))
UNCOLLECT_LF_PIECE=To_Nol(request("UNCOLLECT_LF_PIECE"))
UNARRANGE_FFB=To_Nol(request("UNARRANGE_FFB"))
UNPRUNE_FROND=To_Nol(request("UNPRUNE_FROND"))
QE_1=To_Nol(request("QE_1"))
QE_2=To_Nol(request("QE_2"))
QE_3=To_Nol(request("QE_3"))
QE_4=To_Nol(request("QE_4"))
QE_5=To_Nol(request("QE_5"))
QE_6=To_Nol(request("QE_6"))
QE_7=To_Nol(request("QE_7"))
QE_8=To_Nol(request("QE_8"))
QE_9=To_Nol(request("QE_9"))
QE_10=To_Nol(request("QE_10"))

' Added Variable ChargeJob, ChargeType, ChargeCode
chargejob=Request("chargejob")
chargetype=Request("chargetype")
chargecode=Request("chargecode")
businessunit=Request("businessunit")
' End Added
 

fcba = request("fcba")
authorizelevelcode = request("authorizelevelcode")
authorizetemplatecode = request("authorizetemplatecode")

Function f_c_output_harian(fcba__, fddate__, employeecode__, attendance__, jobcode__, targettype__, targetcode__, output__)
  Dim Result : Result = "Success"
  Dim rsAvailable, strAvailable
  
  strAvailable =  " SELECT validasi_att_api.max_output('"&fcba__&"',  TO_DATE('"&fddate__&"', 'mm/dd/yyyy'), '"&employeecode__&"', '"&attendance__&"', '"&jobcode__&"', '"&targettype__&"', '"&targetcode__&"', '"&output__&"') result" & _
          " FROM DUAL"

  Set rsAvailable = Server.CreateObject("ADODB.Recordset")
  rsAvailable.Open strAvailable, connstr, 3
  
  If Not rsAvailable.EOF Then
    Result = rsAvailable("result")
  Else
    Result = "Success"
  End If
  
  strAvailable =  ""
  rsAvailable.Close
  set rsAvailable = Nothing

  f_c_output_harian = Result
End Function

Function f_c_output_rotasi(fcba__, fddate__, employeecode__, attendance__, jobcode__, targettype__, targetcode__, output__)
  Dim Result : Result = "Success"
  Dim rsAvailable, strAvailable
  
  strAvailable =  " SELECT validasi_att_api.max_totaloutput('"&fcba__&"',  TO_DATE('"&fddate__&"', 'mm/dd/yyyy'), '"&employeecode__&"', '"&attendance__&"', '"&jobcode__&"', '"&targettype__&"', '"&targetcode__&"', '"&output__&"') result" & _
          " FROM DUAL"

  Set rsAvailable = Server.CreateObject("ADODB.Recordset")
  rsAvailable.Open strAvailable, connstr, 3
  
  If Not rsAvailable.EOF Then
    Result = rsAvailable("result")
  Else
    Result = "Success"
  End If
  
  strAvailable =  ""
  rsAvailable.Close
  set rsAvailable = Nothing

  f_c_output_rotasi = Result
End Function

'--Edit

Function f_u_output_harian(fcba__, linenokey__, attendance__, output__)
  Dim Result : Result = "Success"
  Dim rsAvailable, strAvailable
  
  strAvailable =  " SELECT validasi_att_api.max_output('"&fcba__&"', '"&linenokey__&"', '"&attendance__&"', '"&output__&"') result" & _
          " FROM DUAL"

  Set rsAvailable = Server.CreateObject("ADODB.Recordset")
  rsAvailable.Open strAvailable, connstr, 3
  
  If Not rsAvailable.EOF Then
    Result = rsAvailable("result")
  Else
    Result = "Success"
  End If
  
  strAvailable =  ""
  rsAvailable.Close
  set rsAvailable = Nothing

  f_u_output_harian = Result
End Function

Function f_u_output_rotasi(fcba__, linenokey__, attendance__, output__)
  Dim Result : Result = "Success"
  Dim rsAvailable, strAvailable
  
  strAvailable =  " SELECT validasi_att_api.max_totaloutput('"&fcba__&"', '"&linenokey__&"', '"&attendance__&"', '"&output__&"') result" & _
          " FROM DUAL"

  Set rsAvailable = Server.CreateObject("ADODB.Recordset")
  rsAvailable.Open strAvailable, connstr, 3
  
  If Not rsAvailable.EOF Then
    Result = rsAvailable("result")
  Else
    Result = "Success"
  End If
  
  strAvailable =  ""
  rsAvailable.Close
  set rsAvailable = Nothing

  f_u_output_rotasi = Result
End Function

'-- End Fungsi Validasi Panen --'


Function NewRecordDailyValidation(init_fcba, init_fddate, init_employeecode, init_attendance, init_mandays)
  Dim Result : Result = "NOENTRY"
  Dim rsAvailable, strAvailable
  
  strAvailable =  " select attendance_gad_api.dailyvalidation('"&init_fcba&"', to_date('"&init_fddate&"', 'mm/dd/yyyy'), '"&init_employeecode&"', '"&init_attendance&"', "&init_mandays&") result" & _
          " from dual"

  Set rsAvailable = Server.CreateObject("ADODB.Recordset")
  rsAvailable.Open strAvailable, connstr, 3
  
  If Not rsAvailable.EOF Then
    Result = rsAvailable("result")
  Else
    Result = "NOENTRY"
  End If
  
  strAvailable =  ""
  rsAvailable.Close
  set rsAvailable = Nothing

  NewRecordDailyValidation = Result
End Function

Function EditRecordDailyValidation(init_fcba, init_fddate, init_employeecode, init_attendance, init_mandays, init_linenokey)
  Dim Result : Result = "NOENTRY"
  Dim rsAvailable, strAvailable
  
  strAvailable =  " select attendance_gad_api.dailyvalidation('"&init_fcba&"', to_date('"&init_fddate&"', 'mm/dd/yyyy'), '"&init_employeecode&"', '"&init_attendance&"', "&init_mandays&", "&init_linenokey&") result" & _
          " from dual"

  Set rsAvailable = Server.CreateObject("ADODB.Recordset")
  rsAvailable.Open strAvailable, connstr, 3
  
  If Not rsAvailable.EOF Then
    Result = rsAvailable("result")
  Else
    Result = "NOENTRY"
  End If
  
  strAvailable =  ""
  rsAvailable.Close
  set rsAvailable = Nothing

  EditRecordDailyValidation = Result
End Function

' Function FP_CheckAttendance(in_employeecode, in_attendance, in_fddate, in_othrs, in_fcba)
'   Dim Result : Result = ""
'   Dim rsAvailable, strAvailable
  
'   strAvailable =  " select FiNGERVALIDATIONabc_API.checkingattendance('"&in_employeecode&"', '"&in_attendance&"', to_date('"&in_fddate&"', 'mm/dd/yyyy'), "&in_othrs&", '"&in_fcba&"') result" & _
'           " from dual"
          
'   ' response.write strAvailable

'   Set rsAvailable = Server.CreateObject("ADODB.Recordset")
'   rsAvailable.Open strAvailable, connstr, 3
  
'   If Not rsAvailable.EOF Then
'     Result = rsAvailable("result")
'   Else
'     Result = ""
'   End If
  
'   strAvailable =  ""
'   rsAvailable.Close
'   set rsAvailable = Nothing

'   FP_CheckAttendance = Result
' End Function

Function CheckCharge(initType, initCode, initJob, initFcba, initDate, initSource)
  Dim initialAvailable
  Dim rsAvailable, strAvailable
  
  strAvailable =  "SELECT F_CekTarget_byDate ('"& initType &"'," & _
          "                           '"& initCode &"'," & _
          "                           '"& initJob &"'," & _
          "                           '"& initFcba &"'," & _
          "                           TO_DATE ('"& initDate &"', 'MM/DD/YYYY')," & _
          "                           'EM') Total" & _
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
  

  '******* END ADD BY UKIE 08 Juni 2014 '*******

  
  'Add Akbar 4 des 2012 begin
  
   ' q="select sum(mandays)  mandays1 from Attendance_GAD where  to_char(fdDate,'"& "mm/dd/yyyy" & "')='" & fdDate & _
      ' "' and EmployeeCode='" & EmployeeCode & "'  and fcBA='" & Session("fcBA") & "'"
   ' set rs = connStr.execute(q)
   ' if not rs.eof then
        ' IF NOT IsNull(rs("mandays1")) then
       ' TOTMANDAYS = cdbl(rs("mandays1"))
      ' end if
   ' end if
   ' rs.close()
     ' rs=null 
  
  
  
  ' q=  "select sum(mandays)  mandays1 from Attendance_GAD " & _
      ' " where  to_char(fdDate,'"& "mm/dd/yyyy" & "')='" & fdDate & _
     ' "' and EmployeeCode='" & EmployeeCode & "' And locationtype='" & locationtype & "' and " & _
      ' " LOCATIONCODE='" & LOCATIONCODE & "' and JOBCODE='" & JOBCODE & "' and fcBA='" & Session("fcBA") & "'"
   ' set rs = connStr.execute(q)
   ' if not rs.eof then
        ' IF NOT IsNull(rs("mandays1")) then
       ' OLDMANDAYS = cdbl(rs("mandays1"))
      ' end if
   ' end if
   ' rs.close()
     ' rs=null 
   
  'Add Akbar 4 des 2012 end
  
  
	q = " select sum(mandays)  mandays1 from ATTENDANCE_GAD_TEMP where  fdDate = to_DATE('" & fdDate & "','mm/dd/yyyy')" & _
		" and EmployeeCode=trim('" & request("employeecode_"&TRIM(n)) & "')  and fcBA=trim('" & Session("fcba") & "')"
		' response.write q
	set rs = connStr.execute(q)
	if not rs.eof then
		IF NOT IsNull(rs("mandays1")) then
			TOTMANDAYS = cdbl(rs("mandays1"))
		end if
	end if
	rs.close()
	rs=null 
   
   ' Add By Akbar 10 Feb 2014
   
  q = " select 1 from ATTENDANCE_GAD_TEMP where FDDATE = to_date('" & fdDate & "', 'mm/dd/yyyy')" & _
      " and EmployeeCode = trim('" & request("employeecode_"&TRIM(n)) & "')" & _        
      " And locationtype = trim('" & request("locationtype_"&TRIM(n)) & "')" & _
      " and LOCATIONCODE = trim('" & request("locationcode_"&TRIM(n)) & "')" & _
      " and JOBCODE = trim('" & request("jobcode_"&TRIM(n)) & "')" & _
      " and fcBA= trim('" & Session("fcba") & "')"
      
	set rs = connStr.execute(q)
	if not rs.eof then
		recordAda = "T"
	end if
	rs.close()
	rs=null 
   
   ' End Add By Akbar 10 Feb 2014
   
  q = " select sum(mandays)  mandays1 from ATTENDANCE_GAD_TEMP " & _
      " where  fdDate = to_DATE('" & fdDate & "','mm/dd/yyyy')" & _
      " and EmployeeCode=trim('" & request("employeecode_"&TRIM(n)) & "') And locationtype=trim('" & request("locationtype_"&TRIM(n)) & "') and " & _
      " LOCATIONCODE=trim('" & request("locationcode_"&TRIM(n)) & "') and JOBCODE=trim('" & request("jobcode_"&TRIM(n)) & "') and fcBA=trim('" & Session("fcba") & "')"
	set rs = connStr.execute(q)
	if not rs.eof then
		IF NOT IsNull(rs("mandays1")) then
			OLDMANDAYS = cdbl(rs("mandays1"))
		end if
	end if
	rs.close()
	rs=null 
  
  
  
  ' Start Modification to Pick Validation whether User have Mandays or Not. by NAF, 24-4-2013
  ' its depended on GANGTYPE and ATTENDANCE ==> Should be KJ
  ' Jika Total Dia bekerja bukan sebagai Gang Mandor dan KJ, maka Mandays Harus Lebih besar dari Nol
  
	q = " Select gangtypecode from Gang " & _
		" where " & _
		" fccode = trim('" & GANGCODE & "') and " & _
		" fcba = trim('" & Session("fcba") & "')"
    
	set rs = connStr.execute(q)
	if not rs.eof then
		GangType = rs("gangtypecode")
	End if
	rs.close()
	rs=null     

  
  ' End of Modification 24-4-2013 
  
  ' Start modification to pick the default Application by NAF, 11-10-2012
  ' IsMandatory will be put with TRUE or FALSE
  ' DefaultJob will be put with Job Activity
  
  
	q = " SELECT param_1 as IsMandatory, param_2 as DefaultJob FROM PARAMETER " & _
		" WHERE " & _
		" parcode = 'E0102' AND " & _
		" fccode = trim('" & request("attendance_"&TRIM(n)) & "') AND " & _
		" fcba = '"& fcba  &"' "
  
	set rs = connStr.execute(q)
	if not rs.eof then
		IsMandatory = rs("IsMandatory")
		DefaultJob = rs("DefaultJob")
	End if
	rs.close()
	rs=null 
    
  ' End of Modification by NAF, 11-10-2012
  
	done = 1
	if ATTENDANCE <> "" then 
		if IsMandatory = "" then 
			done = 0
		end if
	end if

Sub DeleteAllowance
  if (len(trim(""&ALLOWANCE_CODE))) > 0 OR cdbl(to_nol(Amount)) > 0 then
    q2  = " delete from attendance_empallded where linenokey='"& LineNoKey &"' and fcba ='"& Session("fcba") &"'"
    ' response.Write(q2)
    connStr.execute(q2)
  Else
    q2  = " delete from attendance_empallded where linenokey='"& LineNoKey &"' and fcba ='"& Session("fcba") &"'"
    ' response.Write(q2)
    connStr.execute(q2)
  End If
End Sub

sub InsertAllowance
  qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& request("jobcode_"&TRIM(n)) &"' group by fcba"      
  ' response.write qLineMax
  set rsqLineMax = Server.CreateObject("ADODB.Recordset")
  rsqLineMax.Open qLineMax, connstr
  if not rsqLineMax.eof then
    code = cdbl(to_nol(rsqLineMax("fccode")))
    businessunit = rsqLineMax("fcba")
  end if 

  if code = 0 then
    if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 ) then
      err.number = 1001
      err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
    else
      '-- Get Data Max Linenokey
      querymaxlineno = " select nvl(max(linenokey), 0) linenokey" & _ 
					   " from ATTENDANCE_GAD " & _ 
					   " where fcba = '" & Session("fcba") & "'"

      Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
      rsmaxlineno.CursorType = 3
      rsmaxlineno.CursorLocation = 3
      
      'response.write querymaxlineno
      
      rsmaxlineno.Open querymaxlineno, connstr
      if not rsmaxlineno.eof then
         maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
      else
         maxlinenokey = 0
      end if
      rsmaxlineno.close
      set rsmaxlineno = Nothing
      if (len(trim(""&ALLOWANCE_CODE))) = 0 then
      Else
      if cdbl(to_nol(Amount)) = 0 then
        tempStringError = "Attendance Saved Successfully without Allowded, thank you."         
        response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
      Else
        q2  = " Insert into attendance_empallded (" & _
            " LINENOKEY" & _
            " , ALLOWANCE_CODE" & _
            " , AMOUNT" & _
            " , FCENTRY" & _
            " , FCEDIT" & _
            " , FCIP" & _
            " , FCBA" & _
            " , LASTUPDATE" & _
            " , LASTTIME" & _
            " ) " & _
            " values ("& _
            " '" & maxlinenokey & "'" & _
            " , '" & ALLOWANCE_CODE & "'" & _
            " , replace('" & AMOUNT & "',',','')" & _
            " , '" & Session("CodeName") & "'" & _
            " , '" & Session("CodeName") & "'" & _
            " , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
            " , '" & Session("fcBA") & "'" & _
            " , sysdate" & _
            " , To_Char(sysdate,'HH24:MI')" & _
            " )"    
            ' response.Write(q2)
            connStr.execute(q2)
        End if
      End if
    End If
  else
  if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 ) then
    err.number = 1001
    err.description = "Anda harus input Charge Job, Charge Type dan Charge Code! "
  else
    Dim val_Charge : val_Charge = 0
    val_Charge = CheckCharge(chargetype, chargecode, chargejob, businessunit, fddate, sourcetype)
  If val_Charge <> "Success" Then
    Err.Number = 1001
    Err.Description = val_Charge
  Else
            '-- Get Data Max Linenokey
      querymaxlineno = " select nvl(max(linenokey), 0) linenokey" & _ 
               " from ATTENDANCE_GAD_TEMP" & _ 
               " where fcba = '" & Session("fcBA") & "'"

      Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
      rsmaxlineno.CursorType = 3
      rsmaxlineno.CursorLocation = 3
      
      'response.write querymaxlineno
      
      rsmaxlineno.Open querymaxlineno, connstr
      if not rsmaxlineno.eof then
         maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
      else
         maxlinenokey = 0
      end if
      rsmaxlineno.close
      set rsmaxlineno = Nothing
    if (len(trim(""&ALLOWANCE_CODE))) = 0 then
      Else
      if cdbl(to_nol(Amount)) = 0 then
        tempStringError = "Attendance Saved Successfully without Allowded, thank you."         
        response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
      Else
        q2  = " Insert into attendance_empallded (" & _
            " LINENOKEY" & _
            " , ALLOWANCE_CODE" & _
            " , AMOUNT" & _
            " , FCENTRY" & _
            " , FCEDIT" & _
            " , FCIP" & _
            " , FCBA" & _
            " , LASTUPDATE" & _
            " , LASTTIME" & _
            " ) " & _
            " values ("& _
            " '" & maxlinenokey & "'" & _
            " , '" & ALLOWANCE_CODE & "'" & _
            " , replace('" & AMOUNT & "',',','')" & _
            " , '" & Session("CodeName") & "'" & _
            " , '" & Session("CodeName") & "'" & _
            " , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
            " , '" & Session("fcBA") & "'" & _
            " , sysdate" & _
            " , To_Char(sysdate,'HH24:MI')" & _
            " )"    
            ' response.Write(q2)
            connStr.execute(q2)
        End If
      End If
    End If
  End If
End If
End Sub

sub UpdateAllowance
  if (len(trim(""&ALLOWANCE_CODE))) = 0 OR cdbl(to_nol(Amount)) = 0 then
    Call DeleteAllowance
    Call InsertAllowance
  Else
  if (len(trim(""&ALLOWANCE_CODE))) <> 0 then
    Call DeleteAllowance
    ' Call DeleteDataCharge
    ' Call InsertDataCharge
    Call InsertAllowance
    ' q2  = " Insert into attendance_empallded (" & _
    '     " LINENOKEY" & _
    '     " , ALLOWANCE_CODE" & _
    '     " , AMOUNT" & _
    '     " , FCENTRY" & _
    '     " , FCEDIT" & _
    '     " , FCIP" & _
    '     " , FCBA" & _
    '     " , LASTUPDATE" & _
    '     " , LASTTIME" & _
    '     " ) " & _
    '     " values ("& _
    '     " '" & linenokey & "'" & _
    '     " , '" & ALLOWANCE_CODE & "'" & _
    '     " , replace('" & AMOUNT & "',',','')" & _
    '     " , '" & Session("CodeName") & "'" & _
    '     " , '" & Session("CodeName") & "'" & _
    '     " , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
    '     " , '" & Session("fcBA") & "'" & _
    '     " , sysdate" & _
    '     " , To_Char(sysdate,'HH24:MI')" & _
    '     " )"    
    '     response.Write(q2)
    '     connStr.execute(q2)
    End If
  End If
End Sub

Sub DeleteDataCharge
  q = " delete from ATTENDANCE_GAD_TEMP" & _
    " where fcba = '" & Session("fcba") & "'" & _
    "       and linenokey = '" & request("linenokey_"&TRIM(n)) & "'"
    response.write q
  connStr.execute(q)
End Sub

Sub InsertDataCharge

  qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& request("jobcode_"&TRIM(n)) &"' group by fcba"      
  ' response.write qLineMax
  set rsqLineMax = Server.CreateObject("ADODB.Recordset")
  rsqLineMax.Open qLineMax, connstr
  if not rsqLineMax.eof then
    code = cdbl(to_nol(rsqLineMax("fccode")))
    businessunit = rsqLineMax("fcba")
  end if 

if code = 0 then
    if ( len(trim(""&request("chargetype_"&TRIM(n)))) <> 0 or len(trim(""&request("chargecode_"&TRIM(n)))) <> 0 or len(trim(""&request("chargejob_"&TRIM(n)))) <> 0 ) then
      err.number = 1001
      err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
    else
		'-- Get Data Max Linenokey
		querymaxlineno = " select nvl(max(linenokey), 0) linenokey" & _ 
						" from ATTENDANCE_GAD_TEMP" & _ 
						" where fcba = '" & Session("fcba") & "'"

		Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
		rsmaxlineno.CursorType = 3
		rsmaxlineno.CursorLocation = 3

		'response.write querymaxlineno

		rsmaxlineno.Open querymaxlineno, connstr
		if not rsmaxlineno.eof then
		 maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
		else
		 maxlinenokey = 0
		end if
		rsmaxlineno.close
		set rsmaxlineno = Nothing

		maxlinenokey = maxlinenokey + 1
      
		q = " SELECT COUNT (*) jobcodespbharvesting" & _
			"  FROM parameterdetail" & _
			" WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
			"       AND fcba = '"& Session("fcba") &"'" & _
			"       AND fccode = '"& request("jobcode_"&TRIM(n)) &"'"

      ' response.Write q

      set rsrecurtrans = server.CreateObject("Adodb.recordset")
      rsrecurtrans.Open q, connstr

      if Not rsrecurtrans.eof then
        jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
      end if
      rsrecurtrans.Close
      set rsrecurtrans = Nothing

      q = " SELECT FCCODE FROM DOCUMENT_CLASSIFICATION WHERE DOCUMENT_TYPE = 'AA'"

      ' response.Write q

      set rsrecurtrans = server.CreateObject("Adodb.recordset")
      rsrecurtrans.Open q, connstr

      if Not rsrecurtrans.eof then
        document_classification = rsrecurtrans("FCCODE")
      end if
      rsrecurtrans.Close
      set rsrecurtrans = Nothing
      
      if cdbl(to_nol(jobcodespbharvesting)) > 0 then 

      q = " Insert into ATTENDANCE_GAD  (GANGCODE" & _
          "                             , FDDATE" & _
          "                             , SUPERVISION_1" & _
          "                             , SUPERVISION_2" & _
          "                             , SUPERVISION_3" & _
          "                             , SUPERVISION_4" & _
          "                             , EMPLOYEECODE" & _
          "                             , ATTENDANCE" & _
          "                             , JOBCODE" & _
          "                             , LOCATIONTYPE" & _
          "                             , LOCATIONCODE" & _
          "                             , CHARGEJOB" & _
          "                             , CHARGETYPE" & _
          "                             , CHARGECODE" & _
          "                             , MANDAYS" & _
          "                             , OTHRS" & _
          "                             , RATE" & _
          "                             , UNIT" & _
          "                             , OUTPUT" & _
          "                             , REFERENCE" & _
          "                             , REMARKS" & _
          "                             , BUCKET" & _
          "                             , FCENTRY" & _
          "                             , FCEDIT" & _
          "                             , FCIP" & _
          "                             , FCBA" & _
          "                             , LASTUPDATE" & _
          "                             , LASTTIME" & _
          "                             , linenokey" & _
          "                             , RowState" & _
          "                             , document_classification" & _
          "                             , TYPE_OVERTIME" & _
          "                             , OVERTIME_HOURS" & _
          "                             , KG_BRONDOLAN" & _
          "                             , SPBNO" & _
          "                             , KG_JANJANG" & _
          " ) values (trim('" & GANGCODE & "')" & _
          "           , to_date('" & fddate & "','mm/dd/yyyy')" & _
          "           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
          "           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
          "           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
          "           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
          "           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
          "           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
          "           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
          "           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
          "           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
          "           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
          "           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
          "           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
          "           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
          "           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
          "           , trim('" & request("rate_"&TRIM(n)) & "')" & _
          "           , trim('" & request("unit_"&TRIM(n)) & "')" & _
          "           , trim('" & request("output_"&TRIM(n)) & "')" & _
          "           , '" & request("reference_"&TRIM(n)) & "'" & _
          "           , '" & request("remarks_"&TRIM(n)) & "'" & _
          "           , '" & request("bucket_"&TRIM(n)) & "'" & _
          "           , '" & Session("CodeName") & "'" & _
          "           , '" & Session("CodeName") & "'" & _
          "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
          "           , '" & Session("fcba") & "'" & _
          "           , sysdate" & _
          "           , To_Char(sysdate,'HH24:MI')" & _
          "           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD  where fcba = '"& Session("fcba") &"')" & _
          "           , '"& rowstate &"'" & _
          "           , '" & request("document_classification_"&TRIM(n)) & "'" & _
          "           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
          "           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
          "           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
          "           , '" & request("spbno_"&TRIM(n)) & "'" & _
          "           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
          " )"
        ' response.Write(q)
        connStr.execute(q)
      else
        q = " Insert into ATTENDANCE_GAD  (GANGCODE" & _
            "                             , FDDATE" & _
            "                             , SUPERVISION_1" & _
            "                             , SUPERVISION_2" & _
            "                             , SUPERVISION_3" & _
            "                             , SUPERVISION_4" & _
            "                             , EMPLOYEECODE" & _
            "                             , ATTENDANCE" & _
            "                             , JOBCODE" & _
            "                             , LOCATIONTYPE" & _
            "                             , LOCATIONCODE" & _
            "                             , CHARGEJOB" & _
            "                             , CHARGETYPE" & _
            "                             , CHARGECODE" & _
            "                             , MANDAYS" & _
            "                             , OTHRS" & _
            "                             , RATE" & _
            "                             , UNIT" & _
            "                             , OUTPUT" & _
            "                             , REFERENCE" & _
            "                             , REMARKS" & _
            "                             , BUCKET" & _
            "                             , FCENTRY" & _
            "                             , FCEDIT" & _
            "                             , FCIP" & _
            "                             , FCBA" & _
            "                             , LASTUPDATE" & _
            "                             , LASTTIME" & _
            "                             , linenokey" & _
            "                             , RowState" & _
            "                             , document_classification" & _
            "                             , TYPE_OVERTIME" & _
            "                             , OVERTIME_HOURS" & _
            "                             , KG_BRONDOLAN" & _
            "                             , SPBNO" & _
            "                             , KG_JANJANG" & _
            " ) values (trim('" & GANGCODE & "')" & _
            "           , to_date('" & fddate & "','mm/dd/yyyy')" & _
            "           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
            "           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
            "           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
            "           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
            "           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
            "           , trim('" & request("rate_"&TRIM(n)) & "')" & _
            "           , trim('" & request("unit_"&TRIM(n)) & "')" & _
            "           , trim('" & request("output_"&TRIM(n)) & "')" & _
            "           , '" & request("reference_"&TRIM(n)) & "'" & _
            "           , '" & request("remarks_"&TRIM(n)) & "'" & _
            "           , '" & request("bucket_"&TRIM(n)) & "'" & _
            "           , '" & Session("CodeName") & "'" & _
            "           , '" & Session("CodeName") & "'" & _
            "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
            "           , '" & Session("fcba") & "'" & _
            "           , sysdate" & _
            "           , To_Char(sysdate,'HH24:MI')" & _
            "           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD  where fcba = '"& Session("fcba") &"')" & _
            "           , '"& rowstate &"'" & _
            "           , '" & request("document_classification_"&TRIM(n)) & "'" & _
            "           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
            "           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
            "           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
            "           , '" & request("spbno_"&TRIM(n)) & "'" & _
            "           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
            " )"
          ' response.Write(q)
          connStr.execute(q)
      end if 
    End If
  else
    if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 ) then
      err.number = 1001
      err.description = "Anda harus input Charge Job, Charge Type dan Charge Code! "
    else
      Dim val_Charge : val_Charge = 0
      val_Charge = CheckCharge(chargetype, chargecode, chargejob, businessunit, fddate, sourcetype)
      If val_Charge <> "Success" Then
        Err.Number = 1001
        Err.Description = val_Charge
      Else
        '-- Get Data Max Linenokey
        querymaxlineno = " select nvl(max(linenokey), 0) linenokey" & _ 
                 " from ATTENDANCE_GAD " & _ 
                 " where fcba = '" & Session("fcba") & "'"

        Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
        rsmaxlineno.CursorType = 3
        rsmaxlineno.CursorLocation = 3
        
        'response.write querymaxlineno
        
        rsmaxlineno.Open querymaxlineno, connstr
        if not rsmaxlineno.eof then
           maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
        else
           maxlinenokey = 0
        end if
        rsmaxlineno.close
        set rsmaxlineno = Nothing
        
        maxlinenokey = maxlinenokey + 1
        q = " SELECT COUNT (*) jobcodespbharvesting" & _
        "  FROM parameterdetail" & _
        " WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
        "       AND fcba = '"& Session("fcba") &"'" & _
        "       AND fccode = '"& request("jobcode_"&TRIM(n)) &"'"

      ' response.Write q

      set rsrecurtrans = server.CreateObject("Adodb.recordset")
      rsrecurtrans.Open q, connstr

      if Not rsrecurtrans.eof then
        jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
      end if
      rsrecurtrans.Close
      set rsrecurtrans = Nothing
      
      if cdbl(to_nol(jobcodespbharvesting)) > 0 then
        q = " Insert into ATTENDANCE_GAD  (GANGCODE" & _
            "                             , FDDATE" & _
            "                             , SUPERVISION_1" & _
            "                             , SUPERVISION_2" & _
            "                             , SUPERVISION_3" & _
            "                             , SUPERVISION_4" & _
            "                             , EMPLOYEECODE" & _
            "                             , ATTENDANCE" & _
            "                             , JOBCODE" & _
            "                             , LOCATIONTYPE" & _
            "                             , LOCATIONCODE" & _
            "                             , CHARGEJOB" & _
            "                             , CHARGETYPE" & _
            "                             , CHARGECODE" & _
            "                             , MANDAYS" & _
            "                             , OTHRS" & _
            "                             , RATE" & _
            "                             , UNIT" & _
            "                             , OUTPUT" & _
            "                             , REFERENCE" & _
            "                             , REMARKS" & _
            "                             , BUCKET" & _
            "                             , FCENTRY" & _
            "                             , FCEDIT" & _
            "                             , FCIP" & _
            "                             , FCBA" & _
            "                             , LASTUPDATE" & _
            "                             , LASTTIME" & _
            "                             , linenokey" & _
            "                             , RowState" & _
            "                             , document_classification" & _
            "                             , TYPE_OVERTIME" & _
            "                             , OVERTIME_HOURS" & _
            "                             , KG_BRONDOLAN" & _
            "                             , SPBNO" & _
            "                             , KG_JANJANG" & _
            " ) values (trim('" & GANGCODE & "')" & _
            "           , to_date('" & fddate & "','mm/dd/yyyy')" & _
            "           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
            "           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
            "           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
            "           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
            "           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
            "           , trim('" & request("rate_"&TRIM(n)) & "')" & _
            "           , trim('" & request("unit_"&TRIM(n)) & "')" & _
            "           , trim('" & request("output_"&TRIM(n)) & "')" & _
            "           , '" & request("reference_"&TRIM(n)) & "'" & _
            "           , '" & request("remarks_"&TRIM(n)) & "'" & _
            "           , '" & request("bucket_"&TRIM(n)) & "'" & _
            "           , '" & Session("CodeName") & "'" & _
            "           , '" & Session("CodeName") & "'" & _
            "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
            "           , '" & Session("fcba") & "'" & _
            "           , sysdate" & _
            "           , To_Char(sysdate,'HH24:MI')" & _
            "           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD  where fcba = '"& Session("fcba") &"')" & _
            "           , '"& rowstate &"'" & _
            "           , '" & request("document_classification_"&TRIM(n)) & "'" & _
            "           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
            "           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
            "           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
            "           , '" & request("spbno_"&TRIM(n)) & "'" & _
            "           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
            " )"
          ' response.Write(q)
          connStr.execute(q)
      else
        q = " Insert into ATTENDANCE_GAD  (GANGCODE" & _
            "                             , FDDATE" & _
            "                             , SUPERVISION_1" & _
            "                             , SUPERVISION_2" & _
            "                             , SUPERVISION_3" & _
            "                             , SUPERVISION_4" & _
            "                             , EMPLOYEECODE" & _
            "                             , ATTENDANCE" & _
            "                             , JOBCODE" & _
            "                             , LOCATIONTYPE" & _
            "                             , LOCATIONCODE" & _
            "                             , CHARGEJOB" & _
            "                             , CHARGETYPE" & _
            "                             , CHARGECODE" & _
            "                             , MANDAYS" & _
            "                             , OTHRS" & _
            "                             , RATE" & _
            "                             , UNIT" & _
            "                             , OUTPUT" & _
            "                             , REFERENCE" & _
            "                             , REMARKS" & _
            "                             , BUCKET" & _
            "                             , FCENTRY" & _
            "                             , FCEDIT" & _
            "                             , FCIP" & _
            "                             , FCBA" & _
            "                             , LASTUPDATE" & _
            "                             , LASTTIME" & _
            "                             , linenokey" & _
            "                             , RowState" & _
            "                             , document_classification" & _
            "                             , TYPE_OVERTIME" & _
            "                             , OVERTIME_HOURS" & _
            "                             , KG_BRONDOLAN" & _
            "                             , SPBNO" & _
            "                             , KG_JANJANG" & _
            " ) values (trim('" & GANGCODE & "')" & _
            "           , to_date('" & fddate & "','mm/dd/yyyy')" & _
            "           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
            "           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
            "           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
            "           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
            "           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
            "           , trim('" & request("rate_"&TRIM(n)) & "')" & _
            "           , trim('" & request("unit_"&TRIM(n)) & "')" & _
            "           , trim('" & request("output_"&TRIM(n)) & "')" & _
            "           , '" & request("reference_"&TRIM(n)) & "'" & _
            "           , '" & request("remarks_"&TRIM(n)) & "'" & _
            "           , '" & request("bucket_"&TRIM(n)) & "'" & _
            "           , '" & Session("CodeName") & "'" & _
            "           , '" & Session("CodeName") & "'" & _
            "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
            "           , '" & Session("fcba") & "'" & _
            "           , sysdate" & _
            "           , To_Char(sysdate,'HH24:MI')" & _
            "           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD  where fcba = '"& Session("fcba") &"')" & _
            "           , '"& rowstate &"'" & _
            "           , '" & request("document_classification_"&TRIM(n)) & "'" & _
            "           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
            "           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
            "           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
            "           , '" & request("spbno_"&TRIM(n)) & "'" & _
            "           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
            " )"
          ' response.Write(q)
          connStr.execute(q) 
      end if 
      End If
    End If
  End If
End Sub 

Sub UpdateDataCharge
  qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& request("jobcode_"&TRIM(n)) &"' group by fcba"      
  ' response.write qLineMax
  set rsqLineMax = Server.CreateObject("ADODB.Recordset")
  rsqLineMax.Open qLineMax, connstr
  if not rsqLineMax.eof then
    code = cdbl(to_nol(rsqLineMax("fccode")))
    businessunit = rsqLineMax("fcba")
  end if 

if code = 0 then
    if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 ) then
      err.number = 1001
      err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
    else
    Call DeleteDataCharge
    '-- Get Data Max Linenokey
      querymaxlineno = " select nvl(max(linenokey), 0) linenokey" & _ 
               " from ATTENDANCE_GAD_TEMP" & _ 
               " where fcba = '" & Session("fcba") & "'"

      Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
      rsmaxlineno.CursorType = 3
      rsmaxlineno.CursorLocation = 3
      
      'response.write querymaxlineno
      
      rsmaxlineno.Open querymaxlineno, connstr
      if not rsmaxlineno.eof then
         maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
      else
         maxlinenokey = 0
      end if
      rsmaxlineno.close
      set rsmaxlineno = Nothing
      
      maxlinenokey = maxlinenokey + 1
      q = " SELECT COUNT (*) jobcodespbharvesting" & _
        "  FROM parameterdetail" & _
        " WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
        "       AND fcba = '"& Session("fcba") &"'" & _
        "       AND fccode = '"& request("jobcode_"&TRIM(n)) &"'"

      ' response.Write q

      set rsrecurtrans = server.CreateObject("Adodb.recordset")
      rsrecurtrans.Open q, connstr

      if Not rsrecurtrans.eof then
        jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
      end if
      rsrecurtrans.Close
      set rsrecurtrans = Nothing
      
      if cdbl(to_nol(jobcodespbharvesting)) > 0 then
      
          q = " Insert into ATTENDANCE_GAD  (GANGCODE" & _
              "                             , FDDATE" & _
              "                             , SUPERVISION_1" & _
              "                             , SUPERVISION_2" & _
              "                             , SUPERVISION_3" & _
              "                             , SUPERVISION_4" & _
              "                             , EMPLOYEECODE" & _
              "                             , ATTENDANCE" & _
              "                             , JOBCODE" & _
              "                             , LOCATIONTYPE" & _
              "                             , LOCATIONCODE" & _
              "                             , CHARGEJOB" & _
              "                             , CHARGETYPE" & _
              "                             , CHARGECODE" & _
              "                             , MANDAYS" & _
              "                             , OTHRS" & _
              "                             , RATE" & _
              "                             , UNIT" & _
              "                             , OUTPUT" & _
              "                             , REFERENCE" & _
              "                             , REMARKS" & _
              "                             , BUCKET" & _
              "                             , FCENTRY" & _
              "                             , FCEDIT" & _
              "                             , FCIP" & _
              "                             , FCBA" & _
              "                             , LASTUPDATE" & _
              "                             , LASTTIME" & _
              "                             , linenokey" & _
              "                             , RowState" & _
              "                             , document_classification" & _
              "                             , TYPE_OVERTIME" & _
              "                             , OVERTIME_HOURS" & _
              "                             , KG_BRONDOLAN" & _
              "                             , SPBNO" & _
              "                             , KG_JANJANG" & _
              " ) values (trim('" & GANGCODE & "')" & _
              "           , to_date('" & fddate & "','mm/dd/yyyy')" & _
              "           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
              "           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
              "           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
              "           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
              "           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
              "           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
              "           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
              "           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
              "           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
              "           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
              "           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
              "           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
              "           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
              "           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
              "           , trim('" & request("rate_"&TRIM(n)) & "')" & _
              "           , trim('" & request("unit_"&TRIM(n)) & "')" & _
              "           , trim('" & request("output_"&TRIM(n)) & "')" & _
              "           , '" & request("reference_"&TRIM(n)) & "'" & _
              "           , '" & request("remarks_"&TRIM(n)) & "'" & _
              "           , '" & request("bucket_"&TRIM(n)) & "'" & _
              "           , '" & Session("CodeName") & "'" & _
              "           , '" & Session("CodeName") & "'" & _
              "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
              "           , '" & Session("fcba") & "'" & _
              "           , sysdate" & _
              "           , To_Char(sysdate,'HH24:MI')" & _
              "           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD  where fcba = '"& Session("fcba") &"')" & _
              "           , '"& rowstate &"'" & _
              "           , '" & request("document_classification_"&TRIM(n)) & "'" & _
              "           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
              "           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
              "           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
              "           , '" & request("spbno_"&TRIM(n)) & "'" & _
              "           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
              " )" 
          response.Write(q)
          connStr.execute(q)
      else
      OTHRS = request ("OTHRS")
        q = " Insert into ATTENDANCE_GAD  (GANGCODE" & _
            "                             , FDDATE" & _
            "                             , SUPERVISION_1" & _
            "                             , SUPERVISION_2" & _
            "                             , SUPERVISION_3" & _
            "                             , SUPERVISION_4" & _
            "                             , EMPLOYEECODE" & _
            "                             , ATTENDANCE" & _
            "                             , JOBCODE" & _
            "                             , LOCATIONTYPE" & _
            "                             , LOCATIONCODE" & _
            "                             , CHARGEJOB" & _
            "                             , CHARGETYPE" & _
            "                             , CHARGECODE" & _
            "                             , MANDAYS" & _
            "                             , OTHRS" & _
            "                             , RATE" & _
            "                             , UNIT" & _
            "                             , OUTPUT" & _
            "                             , REFERENCE" & _
            "                             , REMARKS" & _
            "                             , BUCKET" & _
            "                             , FCENTRY" & _
            "                             , FCEDIT" & _
            "                             , FCIP" & _
            "                             , FCBA" & _
            "                             , LASTUPDATE" & _
            "                             , LASTTIME" & _
            "                             , linenokey" & _
            "                             , RowState" & _
            "                             , document_classification" & _
            "                             , TYPE_OVERTIME" & _
            "                             , OVERTIME_HOURS" & _
            "                             , KG_BRONDOLAN" & _
            "                             , SPBNO" & _
            "                             , KG_JANJANG" & _
            " ) values (trim('" & GANGCODE & "')" & _
            "           , to_date('" & fddate & "','mm/dd/yyyy')" & _
            "           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
            "           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
            "           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
            "           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
            "           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
            "           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
            "           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
            "           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
            "           , trim('" & request("rate_"&TRIM(n)) & "')" & _
            "           , trim('" & request("unit_"&TRIM(n)) & "')" & _
            "           , trim('" & request("output_"&TRIM(n)) & "')" & _
            "           , '" & request("reference_"&TRIM(n)) & "'" & _
            "           , '" & request("remarks_"&TRIM(n)) & "'" & _
            "           , '" & request("bucket_"&TRIM(n)) & "'" & _
            "           , '" & Session("CodeName") & "'" & _
            "           , '" & Session("CodeName") & "'" & _
            "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
            "           , '" & Session("fcba") & "'" & _
            "           , sysdate" & _
            "           , To_Char(sysdate,'HH24:MI')" & _
            "           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD  where fcba = '"& Session("fcba") &"')" & _
            "           , '"& rowstate &"'" & _
            "           , '" & request("document_classification_"&TRIM(n)) & "'" & _
            "           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
            "           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
            "           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
            "           , '" & request("spbno_"&TRIM(n)) & "'" & _
            "           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
            " )" 
        response.Write(q)
        connStr.execute(q)
      end if 
    End If
  Else
    Call DeleteDataCharge
    Call InsertDataCharge
  end if
End Sub

 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "ApprovalAction" or Request("txtJenis")= "SaveData" or Request("txtJenis")= "RejectAction" Then

    If Request("txtJenis")="ApprovalAction" then
		q = " INSERT INTO attendance_gad_hist ( fddate, gangcode" & _
  			"                                           , currentauthorizelevel" & _ 
  			"                                           , fcentry" & _
  			"                                           , fcedit" & _
  			"                                           , fcip" & _
  			"                                           , fcba" & _ 
  			"                                           , lastupdate" & _
  			"                                           , lasttime)"& _
  			" VALUES ( to_date('" & fddate & "','mm/dd/yyyy'), '"& gangcode &"'" & _
  			"          , '" & authorizelevelcode & "'" & _ 
  			"          , '" & Session("CodeName") & "'" & _
  			"          , '" & Session("CodeName") & "'" & _
  			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
  			"          , '" & Session("fcba") & "'" & _ 
  			"          , SYSDATE" & _
  			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
  		' response.write q
  		connStr.execute(q) 
		
		q = " INSERT INTO ag_authorization_comments ( fddate, gangcode" & _
  			"                                         , currentauthorizelevel" & _
  			"                                         , comments" & _
  			"                                         , fcentry" & _
  			"                                         , fcedit" & _
  			"                                         , fcip" & _
  			"                                         , fcba" & _
  			"                                         , lastupdate" & _
  			"                                         , lasttime)"& _
  			" VALUES ( to_date('" & fddate & "','mm/dd/yyyy'), '"& gangcode &"'" & _
  			"          , '" & authorizelevelcode & "'" & _
  			"          , '" & comments & "'" & _
  			"          , '" & Session("CodeName") & "'" & _
  			"          , '" & Session("CodeName") & "'" & _
  			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
  			"          , '" & Session("fcba") & "'" & _
  			"          , SYSDATE" & _
  			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
  		'response.write q
  		connStr.execute(q)

      response.write "<b> Approved"
		is_final = 0
		q = " SELECT COUNT(1) total FROM authorizelevel al" & _
			" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
			" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
			" WHERE at.document_type = 'AA'" & _
			"       AND DECODE(at.document_type, 'PR', at.document_classification, '0') = '0'" & _
			"       AND al.isactive = 'TRUE'" & _
			"       AND al.finallevel = 'TRUE'" & _
			"       AND al.fccode = '" & authorizelevelcode & "'" & _
			"       AND aus.authorizetemplatecode = '" & authorizetemplatecode & "'" & _
			"       AND aus.fcba = '" & Session("fcba") & "'"
		' response.write q
		set rsPurchase = server.CreateObject("Adodb.recordset")
		rsPurchase.Open q, connstr
		IF not rsPurchase.eof then
			is_final = cdbl(to_nol(rsPurchase("total")))
		end if
		
		if is_final > 0 then
			
			For n=1 to (cdbl(request("no"))-1)   
			'-- Check for Empallded PreProcess
			qCheckStatus =  " select EMPALLDEDSTATUS_API.GetCheckPreAttendanceUpdate (to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'YYYY'))," & _ 
							" to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'MM')),'" & Session("fcba") & "') closingStatus from dual "

			Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
			rsCheckStatus.CursorType = 3
			rsCheckStatus.CursorLocation = 3
			'response.write qCheckStatus
			rsCheckStatus.Open qCheckStatus, connstr
			if not rsCheckStatus.eof then
				txtClosingStatus = rsCheckStatus("closingStatus")
			else
				Response.Write  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
				txtClosingStatus="FALSE"
			end if
			rsCheckStatus.close
			set rsCheckStatus = Nothing

			IF txtClosingStatus = "TRUE" THEN
				Cek_IsDailyClosing = IsDailyClosing(fcba, fdDate)
				If Cek_IsDailyClosing = "FALSE" then
					strParameter = 	" select nvl(param_1, 'FALSE') IsMandays" & _
									"        , nvl(param_2, 'FALSE') IsUnit" & _
									"        , nvl(param_3, 'FALSE') IsOutput" & _
									"        , nvl(param_4, 'FALSE') IsOthrs" & _
									"        , nvl(param_5, 'FALSE') IsRate" & _
									" from parameter" & _
									" where fcba = '" & Session("fcba") & "'" & _
									"       and parcode = 'E0102'" & _
									"       and fccode = '" & request("attendance_"&TRIM(n)) & "'"
							  ' response.write strParameter
									   
							  set rsParameter = server.CreateObject("Adodb.recordset")
							  rsParameter.Open strParameter , connstr, adOpenStatic

					If Len(Trim(request("attendance_"&TRIM(n)))) > 0 Then
						If Not rsParameter.EOF Then
							ValidateIsMandays = rsParameter("IsMandays")
							ValidateIsUnit = rsParameter("IsUnit")
							ValidateIsOutput = rsParameter("IsOutput")
							ValidateIsOthrs = rsParameter("IsOthrs")
							ValidateIsRate = rsParameter("IsRate")
						Else
							Err.Number = 1001
							Err.Description = " Parameter belum di set! Contact your Administrator"
						End If
						rsParameter.Close()
						Set rsParameter = Nothing
					End If
						  
					If ValidateIsMandays = "TRUE" Then
						If Cdbl(request("mandays_"&TRIM(n))) <= 0 Then
							Err.Number = 1001
							Err.Description = " Mandays Harus di isi, Range Mandays dimulai dari 0.01 s/d 1 !"
						ElseIf Cdbl(request("mandays_"&TRIM(n))) > 1 Then
							Err.Number = 1001
							Err.Description = " Mandays tidak boleh melebihi angka 1 !"
						Else
							DoneIsMandays = 1
					End If
					ElseIf ValidateIsMandays = "FALSE" Then
						If Cdbl(request("mandays_"&TRIM(n))) <> 0 Then
							Err.Number = 1001
							Err.Description = " Mandays di kosongkan saja atau masukkan angka ""0"" !"
						Else
							DoneIsMandays = 1
						End If
					End If

					'response.write(UNIT)
					If ValidateIsUnit = "TRUE" Then
						If Cdbl(request("unit_"&TRIM(n))) < 0 Then
							Err.Number = 1001
							Err.Description = " Unit Harus di isi!. Range Unit dimulai dari 0 !"
						Else
							DoneIsUnit = 1
						End If
					ElseIf ValidateIsUnit = "FALSE" Then
						If Cdbl(request("unit_"&TRIM(n))) <> 0 Then
							Err.Number = 1001
							Err.Description = " Unit di kosongkan saja atau masukkan angka ""0"" !"
						Else
							DoneIsUnit = 1
						End If
					End If

					If ValidateIsOutput = "TRUE" Then
						If Cdbl(request("output_"&TRIM(n))) < 0 Then
							Err.Number = 1001
							Err.Description = " Output Harus di isi!. Range Output dimulai dari 0 !"
						Else
							DoneIsOutput = 1
						End If
					ElseIf ValidateIsOutput = "FALSE" Then
						If Cdbl(request("output_"&TRIM(n))) <> 0 Then
							Err.Number = 1001
							Err.Description = " Output di kosongkan saja atau masukkan angka ""0"" !"
						Else
							DoneIsOutput = 1
						End If
					End If

					If ValidateIsOthrs = "TRUE" Then
						If Cdbl(request("othrs_"&TRIM(n))) < 0 Then
							Err.Number = 1001
							Err.Description = " Overtime Harus di isi!. Range Overtime dimulai dari 0 !"
						Else
							DoneIsOthrs = 1
						End If
					ElseIf ValidateIsOthrs = "FALSE" Then
						If Cdbl(request("othrs_"&TRIM(n))) <> 0 Then
							Err.Number = 1001
							Err.Description = " Overtime di kosongkan saja atau masukkan angka ""0"" !"
						Else
							DoneIsOthrs = 1
						End If
					End If

					If ValidateIsRate = "TRUE" Then
						If Cdbl(request("rate_"&TRIM(n))) < 0 Then
							Err.Number = 1001
							Err.Description = " Rate Harus di isi!. Range Rate dimulai dari 0 !"
						Else
							DoneIsRate = 1
						End If
					ElseIf ValidateIsRate = "FALSE" Then
						If Cdbl(request("rate_"&TRIM(n))) <> 0 Then
							Err.Number = 1001
							Err.Description = " Rate di kosongkan saja atau masukkan angka ""0"" !"
						Else
							DoneIsRate = 1
						End If
					End If
				
					If DoneIsMandays = 1 And DoneIsUnit = 1 And DoneIsOutput = 1 And DoneIsOthrs = 1 And DoneIsRate = 1 Then	 
						' Add by NAF, 11-10-2012
						If IsMandatory = "FALSE" then 
							if JOBCODE = "" then 
								JOBCODE = defaultJob
							End If  
						End If
						' End of Modification
					  
						'Jika Total Dia bekerja bukan sebagai Gang Mandor dan KJ, maka Mandays Harus Lebih besar dari Nol                  
						IF ( cdbl(request("mandays_"&TRIM(n))) = 0 ) and ( gangtype <> "SUPERVISION" ) and ( request("attendance_"&TRIM(n)) = "KJ" ) THEN 
							response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Gang Type is not a Supervision Gang and its a KJ attendance, the mandays should not be ZERO " & "</font>"
							cmdView = "True"
						ELSE       
							IF ( cdbl(TOTMANDAYS) + request("mandays_"&TRIM(n)) ) <= 1 then     'Add Akbar 4 des 2012                      
								'-- Get Data Entry Validate if KHL not allowed over20 Days, else free
								queryentryvalidate =   " select nvl(ATTENDANCE_GAD_API.ValidasiHKKHL('" & Session("fcba") & "'" & _
													   "                                             , to_date('" & fddate & "', 'mm/dd/yyyy')" & _
													   "                                             , '" & request("employeecode_"&TRIM(n)) & "'" & _
													   "                                             , '" & request("attendance_"&TRIM(n)) & "'" & _
													   "                                             , " & Cdbl(request("mandays_"&TRIM(n))) & "), 'FALSE') entryvalidate" & _ 
													   " from dual"

								Set rsentryvalidate = Server.CreateObject("ADODB.Recordset")
								rsentryvalidate.CursorType = 3
								rsentryvalidate.CursorLocation = 3
								
								'response.write queryentryvalidate
								
								rsentryvalidate.Open queryentryvalidate, connstr
								if not rsentryvalidate.eof then
								   tempentryvalidate = rsentryvalidate("entryvalidate")
								else
								   tempentryvalidate = "FALSE"
								end if
								rsentryvalidate.close
								set rsentryvalidate = Nothing
								
									'response.write tempentryvalidate
							
								if tempentryvalidate = "TRUE" or tempentryvalidate = "EARLY_WARNING" then
									if tempentryvalidate = "EARLY_WARNING" then
										message_warning = "Data Tersimpan, tetapi Karyawan tersebut sudah mencapai >= 21 HK selama 3 Bulan berturut-turut, mohon di proses untuk perubahan status. Jika status masih belum berubah maka tidak dapat di input di bulan depan!" & "<br/>" & _
										"Untuk detailnya silahkan tanya kepada manajemen anda, Terima Kasih!" & " <br/>"
										Response.Write "<div class = ""warning"">" & message_warning & "</div>"
									end if
						  
									' Validasi Maksimal hanya 1 HK perhari
									attInRecord = NewRecordDailyValidation(fcba, fddate, EMPLOYEECODE, ATTENDANCE, MANDAYS)
									'Response.Write attInRecord
						  
									if attInRecord = "ENTRY" then                          
										'Validasi terkait data di fingerprint
										' fp_att_cek = FP_CheckAttendance(EMPLOYEECODE, ATTENDANCE, fddate, overtime_hours, Session("fcBA"))
										' response.write " Record = " & fp_att_cek & " | length = " & len(trim(fp_att_cek))
										' if fp_att_cek = "Validation Passed" then

										c_output_harian = f_c_output_harian(fcba, fddate, EMPLOYEECODE, ATTENDANCE, JOBCODE, locationtype, LOCATIONCODE, OUTPUT)
										c_output_rotasi = f_c_output_rotasi(fcba, fddate, EMPLOYEECODE, ATTENDANCE, JOBCODE, locationtype, LOCATIONCODE, OUTPUT)
										response.write c_output_harian
										if c_output_harian = "Success" then
											if c_output_rotasi = "Success" then
												Call InsertDataCharge  
											else
												err.number = 1001
												err.description = c_output_rotasi & " <br/>"
											end if
										else
											err.number = 1001
											err.description = c_output_harian & " <br/>"
										end if
										' else
										'   err.number = 1001
										'   err.description = fp_att_cek & " <br/>"
										' end if
							Else
								err.number = 1001
								err.description = "Anda sudah memasukkan data kehadiran/ Maksimal 1 HK!!!" & "<br/>" & _
												  "Silahkan cek kembali data yang anda ingin/sudah di masukkan ke dalam sistem!" & " <br/>"
							End if
						Else
							if tempentryvalidate = "NO_INPUT" then
								err.number = 1001
								err.description = "Karyawan tersebut belum di proses perubahan status-nya sehingga hk tidak bisa di input!" & "<br/>" & _
												  "Untuk detailnya silahkan tanya kepada manajemen anda, Terima Kasih!" & " <br/>"
							else
								err.number = 1001
								err.description = "Maksimal HK yang di perbolehkan untuk KHL adalah 20 HK, Cek kembali inputan anda!" & "<br/>" & _
												  "Untuk detailnya silahkan tanya kepada manajemen anda, Terima Kasih!" & " <br/>"
							end if
						End if
					ELSE
						'Add Akbar 4 des 2012 begin
						response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Total mandays perhari 1 ." & "</font>"
						response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Total mandays yang sudah di input tgl "& fdDate &" adalah " & cdbl(TOTMANDAYS) & " sisa mandays yg bisa di input adalah " &  1 -  (cdbl(TOTMANDAYS)) & "</font>"
						cmdView = "True"
						'Add Akbar 4 des 2012 end
					END IF  
						' ELSE
						  ' 'Add Akbar 10 Des 2014 begin
							' response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Record Sudah Ada" & "</font>"
							' cmdView = "True"
						  ' 'Add Akbar 10 Des 2014 end
						' END IF    
						END IF
					End If
				Else
					err.number = 1001
					err.description = "Tanggal """ & fdDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
			Else
			  response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Unable to Update/Delete/Insert," & "</font>"
			  response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Please OPEN your payroll Status </font>"
			  cmdView = "True"      
			End if
		next  
			
			strQueryStatus = "BEGIN ATTENDANCEGAD_API.NextState('" & Session("fcba") & "',to_date('" & fddate & "','mm/dd/yyyy'),'"& gangcode &"','Approve' ); END; "
			connStr.execute(strQueryStatus)
			
			strQueryStatus = " UPDATE attendance_gad SET rowstate = 'Approved' WHERE fcba = '" & Session("fcba") & "'" & _
							 " AND fddate = to_date('" & fddate & "','mm/dd/yyyy') AND gangcode = '"& gangcode &"'"
			connStr.execute(strQueryStatus)

			strQueryStatus = " UPDATE attendance_gad_temp SET rowstate = 'Approved' WHERE fcba = '" & Session("fcba") & "'" & _
							 " AND fddate = to_date('" & fddate & "','mm/dd/yyyy') AND gangcode = '"& gangcode &"'"
			connStr.execute(strQueryStatus)
			
		end if
		' else
			' err.number = 1001
			' err.description = check_data
		' End If
		cmdView	= "True"

    elseif Request("txtJenis")= "SaveData" then
    connstr.BeginTrans
    For n=1 to (cdbl(request("no"))-1)  
      ' Call UpdateDataCharge
		  qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& request("jobcode_"&TRIM(n)) &"' group by fcba"      
		  ' response.write qLineMax
		  set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		  rsqLineMax.Open qLineMax, connstr
		  if not rsqLineMax.eof then
			code = cdbl(to_nol(rsqLineMax("fccode")))
			businessunit = rsqLineMax("fcba")
		  end if 

		if code = 0 then
			if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 ) then
			  err.number = 1001
			  err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
			Else
				q = " delete from ATTENDANCE_GAD_TEMP" & _
					" where fcba = '" & Session("fcba") & "'" & _
					"       and linenokey = '" & request("linenokey_"&TRIM(n)) & "'"
					response.write q
				  connStr.execute(q)
					'-- Get Data Max Linenokey
					querymaxlineno =   " select nvl(max(linenokey), 0) linenokey" & _ 
									   " from ATTENDANCE_GAD_TEMP" & _ 
									   " where fcba = '" & Session("fcba") & "'"

					Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
					rsmaxlineno.CursorType = 3
					rsmaxlineno.CursorLocation = 3

					'response.write querymaxlineno

					rsmaxlineno.Open querymaxlineno, connstr
					if not rsmaxlineno.eof then
					 maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
					else
					 maxlinenokey = 0
					end if
					rsmaxlineno.close
					set rsmaxlineno = Nothing

					maxlinenokey = maxlinenokey + 1
					q = " SELECT COUNT (*) jobcodespbharvesting" & _
						"  FROM parameterdetail" & _
						" WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
						"       AND fcba = '"& Session("fcba") &"'" & _
						"       AND fccode = '"& request("jobcode_"&TRIM(n)) &"'"

					' response.Write q

					set rsrecurtrans = server.CreateObject("Adodb.recordset")
					rsrecurtrans.Open q, connstr

					if Not rsrecurtrans.eof then
					jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
					end if
					rsrecurtrans.Close
					set rsrecurtrans = Nothing

				If cdbl(to_nol(jobcodespbharvesting)) > 0 then

					  q = " Insert into ATTENDANCE_GAD_TEMP  (GANGCODE" & _
						  "                             , FDDATE" & _
						  "                             , SUPERVISION_1" & _
						  "                             , SUPERVISION_2" & _
						  "                             , SUPERVISION_3" & _
						  "                             , SUPERVISION_4" & _
						  "                             , EMPLOYEECODE" & _
						  "                             , ATTENDANCE" & _
						  "                             , JOBCODE" & _
						  "                             , LOCATIONTYPE" & _
						  "                             , LOCATIONCODE" & _
						  "                             , CHARGEJOB" & _
						  "                             , CHARGETYPE" & _
						  "                             , CHARGECODE" & _
						  "                             , MANDAYS" & _
						  "                             , OTHRS" & _
						  "                             , RATE" & _
						  "                             , UNIT" & _
						  "                             , OUTPUT" & _
						  "                             , REFERENCE" & _
						  "                             , REMARKS" & _
						  "                             , BUCKET" & _
						  "                             , FCENTRY" & _
						  "                             , FCEDIT" & _
						  "                             , FCIP" & _
						  "                             , FCBA" & _
						  "                             , LASTUPDATE" & _
						  "                             , LASTTIME" & _
						  "                             , linenokey" & _
						  "                             , RowState" & _
						  "                             , document_classification" & _
						  "                             , TYPE_OVERTIME" & _
						  "                             , OVERTIME_HOURS" & _
						  "                             , KG_BRONDOLAN" & _
						  "                             , SPBNO" & _
						  "                             , KG_JANJANG" & _
						  " ) values (trim('" & GANGCODE & "')" & _
						  "           , to_date('" & fddate & "','mm/dd/yyyy')" & _
						  "           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("rate_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("unit_"&TRIM(n)) & "')" & _
						  "           , trim('" & request("output_"&TRIM(n)) & "')" & _
						  "           , '" & request("reference_"&TRIM(n)) & "'" & _
						  "           , '" & request("remarks_"&TRIM(n)) & "'" & _
						  "           , '" & request("bucket_"&TRIM(n)) & "'" & _
						  "           , '" & Session("CodeName") & "'" & _
						  "           , '" & Session("CodeName") & "'" & _
						  "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						  "           , '" & Session("fcba") & "'" & _
						  "           , sysdate" & _
						  "           , To_Char(sysdate,'HH24:MI')" & _
						  "           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD_TEMP  where fcba = '"& Session("fcba") &"')" & _
						  "           , '"& rowstate &"'" & _
						  "           , '" & request("document_classification_"&TRIM(n)) & "'" & _
						  "           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
						  "           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
						  "           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
						  "           , '" & request("spbno_"&TRIM(n)) & "'" & _
						  "           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
						  " )" 
					  ' response.Write(q)
					  connStr.execute(q)
				Else
					OTHRS = request ("OTHRS")
					q = " Insert into ATTENDANCE_GAD_TEMP  (GANGCODE" & _
						"                             , FDDATE" & _
						"                             , SUPERVISION_1" & _
						"                             , SUPERVISION_2" & _
						"                             , SUPERVISION_3" & _
						"                             , SUPERVISION_4" & _
						"                             , EMPLOYEECODE" & _
						"                             , ATTENDANCE" & _
						"                             , JOBCODE" & _
						"                             , LOCATIONTYPE" & _
						"                             , LOCATIONCODE" & _
						"                             , CHARGEJOB" & _
						"                             , CHARGETYPE" & _
						"                             , CHARGECODE" & _
						"                             , MANDAYS" & _
						"                             , OTHRS" & _
						"                             , RATE" & _
						"                             , UNIT" & _
						"                             , OUTPUT" & _
						"                             , REFERENCE" & _
						"                             , REMARKS" & _
						"                             , BUCKET" & _
						"                             , FCENTRY" & _
						"                             , FCEDIT" & _
						"                             , FCIP" & _
						"                             , FCBA" & _
						"                             , LASTUPDATE" & _
						"                             , LASTTIME" & _
						"                             , linenokey" & _
						"                             , RowState" & _
						"                             , document_classification" & _
						"                             , TYPE_OVERTIME" & _
						"                             , OVERTIME_HOURS" & _
						"                             , KG_BRONDOLAN" & _
						"                             , SPBNO" & _
						"                             , KG_JANJANG" & _
						" ) values (trim('" & GANGCODE & "')" & _
						"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
						"           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
						"           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
						"           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
						"           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
						"           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
						"           , trim('" & request("rate_"&TRIM(n)) & "')" & _
						"           , trim('" & request("unit_"&TRIM(n)) & "')" & _
						"           , trim('" & request("output_"&TRIM(n)) & "')" & _
						"           , '" & request("reference_"&TRIM(n)) & "'" & _
						"           , '" & request("remarks_"&TRIM(n)) & "'" & _
						"           , '" & request("bucket_"&TRIM(n)) & "'" & _
						"           , '" & Session("CodeName") & "'" & _
						"           , '" & Session("CodeName") & "'" & _
						"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						"           , '" & Session("fcba") & "'" & _
						"           , sysdate" & _
						"           , To_Char(sysdate,'HH24:MI')" & _
						"           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD_TEMP  where fcba = '"& Session("fcba") &"')" & _
						"           , '"& rowstate &"'" & _
						"           , '" & request("document_classification_"&TRIM(n)) & "'" & _
						"           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
						"           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
						"           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
						"           , '" & request("spbno_"&TRIM(n)) & "'" & _
						"           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
						" )" 
					' response.Write(q)
					connStr.execute(q)
				End If 
			End If
		Else
			q = " delete from ATTENDANCE_GAD_TEMP" & _
				" where fcba = '" & Session("fcba") & "'" & _
				"       and linenokey = '" & request("linenokey_"&TRIM(n)) & "'"
				response.write q
			  connStr.execute(q)
			' Call InsertDataCharge
			  qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& request("jobcode_"&TRIM(n)) &"' group by fcba"      
			  ' response.write qLineMax
			  set rsqLineMax = Server.CreateObject("ADODB.Recordset")
			  rsqLineMax.Open qLineMax, connstr
			  if not rsqLineMax.eof then
				code = cdbl(to_nol(rsqLineMax("fccode")))
				businessunit = rsqLineMax("fcba")
			  end if 

			if code = 0 then
				if ( len(trim(""&request("chargetype_"&TRIM(n)))) <> 0 or len(trim(""&request("chargecode_"&TRIM(n)))) <> 0 or len(trim(""&request("chargejob_"&TRIM(n)))) <> 0 ) then
				  err.number = 1001
				  err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
				else
					'-- Get Data Max Linenokey
					querymaxlineno = " select nvl(max(linenokey), 0) linenokey" & _ 
									" from ATTENDANCE_GAD_TEMP" & _ 
									" where fcba = '" & Session("fcba") & "'"

					Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
					rsmaxlineno.CursorType = 3
					rsmaxlineno.CursorLocation = 3

					'response.write querymaxlineno

					rsmaxlineno.Open querymaxlineno, connstr
					if not rsmaxlineno.eof then
					 maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
					else
					 maxlinenokey = 0
					end if
					rsmaxlineno.close
					set rsmaxlineno = Nothing

					maxlinenokey = maxlinenokey + 1
				  
					q = " SELECT COUNT (*) jobcodespbharvesting" & _
						"  FROM parameterdetail" & _
						" WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
						"       AND fcba = '"& Session("fcba") &"'" & _
						"       AND fccode = '"& request("jobcode_"&TRIM(n)) &"'"

				  ' response.Write q

				  set rsrecurtrans = server.CreateObject("Adodb.recordset")
				  rsrecurtrans.Open q, connstr

				  if Not rsrecurtrans.eof then
					jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
				  end if
				  rsrecurtrans.Close
				  set rsrecurtrans = Nothing

				  q = " SELECT FCCODE FROM DOCUMENT_CLASSIFICATION WHERE DOCUMENT_TYPE = 'AA'"

				  ' response.Write q

				  set rsrecurtrans = server.CreateObject("Adodb.recordset")
				  rsrecurtrans.Open q, connstr

				  if Not rsrecurtrans.eof then
					document_classification = rsrecurtrans("FCCODE")
				  end if
				  rsrecurtrans.Close
				  set rsrecurtrans = Nothing
				  
				  if cdbl(to_nol(jobcodespbharvesting)) > 0 then 

				  q = " Insert into ATTENDANCE_GAD_TEMP  (GANGCODE" & _
					  "                             , FDDATE" & _
					  "                             , SUPERVISION_1" & _
					  "                             , SUPERVISION_2" & _
					  "                             , SUPERVISION_3" & _
					  "                             , SUPERVISION_4" & _
					  "                             , EMPLOYEECODE" & _
					  "                             , ATTENDANCE" & _
					  "                             , JOBCODE" & _
					  "                             , LOCATIONTYPE" & _
					  "                             , LOCATIONCODE" & _
					  "                             , CHARGEJOB" & _
					  "                             , CHARGETYPE" & _
					  "                             , CHARGECODE" & _
					  "                             , MANDAYS" & _
					  "                             , OTHRS" & _
					  "                             , RATE" & _
					  "                             , UNIT" & _
					  "                             , OUTPUT" & _
					  "                             , REFERENCE" & _
					  "                             , REMARKS" & _
					  "                             , BUCKET" & _
					  "                             , FCENTRY" & _
					  "                             , FCEDIT" & _
					  "                             , FCIP" & _
					  "                             , FCBA" & _
					  "                             , LASTUPDATE" & _
					  "                             , LASTTIME" & _
					  "                             , linenokey" & _
					  "                             , RowState" & _
					  "                             , document_classification" & _
					  "                             , TYPE_OVERTIME" & _
					  "                             , OVERTIME_HOURS" & _
					  "                             , KG_BRONDOLAN" & _
					  "                             , SPBNO" & _
					  "                             , KG_JANJANG" & _
					  " ) values (trim('" & GANGCODE & "')" & _
					  "           , to_date('" & fddate & "','mm/dd/yyyy')" & _
					  "           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("rate_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("unit_"&TRIM(n)) & "')" & _
					  "           , trim('" & request("output_"&TRIM(n)) & "')" & _
					  "           , '" & request("reference_"&TRIM(n)) & "'" & _
					  "           , '" & request("remarks_"&TRIM(n)) & "'" & _
					  "           , '" & request("bucket_"&TRIM(n)) & "'" & _
					  "           , '" & Session("CodeName") & "'" & _
					  "           , '" & Session("CodeName") & "'" & _
					  "           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					  "           , '" & Session("fcba") & "'" & _
					  "           , sysdate" & _
					  "           , To_Char(sysdate,'HH24:MI')" & _
					  "           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD_TEMP  where fcba = '"& Session("fcba") &"')" & _
					  "           , '"& rowstate &"'" & _
					  "           , '" & request("document_classification_"&TRIM(n)) & "'" & _
					  "           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
					  "           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
					  "           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
					  "           , '" & request("spbno_"&TRIM(n)) & "'" & _
					  "           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
					  " )"
					' response.Write(q)
					connStr.execute(q)
				  else
					q = " Insert into ATTENDANCE_GAD_TEMP  (GANGCODE" & _
						"                             , FDDATE" & _
						"                             , SUPERVISION_1" & _
						"                             , SUPERVISION_2" & _
						"                             , SUPERVISION_3" & _
						"                             , SUPERVISION_4" & _
						"                             , EMPLOYEECODE" & _
						"                             , ATTENDANCE" & _
						"                             , JOBCODE" & _
						"                             , LOCATIONTYPE" & _
						"                             , LOCATIONCODE" & _
						"                             , CHARGEJOB" & _
						"                             , CHARGETYPE" & _
						"                             , CHARGECODE" & _
						"                             , MANDAYS" & _
						"                             , OTHRS" & _
						"                             , RATE" & _
						"                             , UNIT" & _
						"                             , OUTPUT" & _
						"                             , REFERENCE" & _
						"                             , REMARKS" & _
						"                             , BUCKET" & _
						"                             , FCENTRY" & _
						"                             , FCEDIT" & _
						"                             , FCIP" & _
						"                             , FCBA" & _
						"                             , LASTUPDATE" & _
						"                             , LASTTIME" & _
						"                             , linenokey" & _
						"                             , RowState" & _
						"                             , document_classification" & _
						"                             , TYPE_OVERTIME" & _
						"                             , OVERTIME_HOURS" & _
						"                             , KG_BRONDOLAN" & _
						"                             , SPBNO" & _
						"                             , KG_JANJANG" & _
						" ) values (trim('" & GANGCODE & "')" & _
						"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
						"           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
						"           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
						"           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
						"           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
						"           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
						"           , trim('" & request("rate_"&TRIM(n)) & "')" & _
						"           , trim('" & request("unit_"&TRIM(n)) & "')" & _
						"           , trim('" & request("output_"&TRIM(n)) & "')" & _
						"           , '" & request("reference_"&TRIM(n)) & "'" & _
						"           , '" & request("remarks_"&TRIM(n)) & "'" & _
						"           , '" & request("bucket_"&TRIM(n)) & "'" & _
						"           , '" & Session("CodeName") & "'" & _
						"           , '" & Session("CodeName") & "'" & _
						"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						"           , '" & Session("fcba") & "'" & _
						"           , sysdate" & _
						"           , To_Char(sysdate,'HH24:MI')" & _
						"           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD_TEMP  where fcba = '"& Session("fcba") &"')" & _
						"           , '"& rowstate &"'" & _
						"           , '" & request("document_classification_"&TRIM(n)) & "'" & _
						"           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
						"           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
						"           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
						"           , '" & request("spbno_"&TRIM(n)) & "'" & _
						"           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
						" )"
					  ' response.Write(q)
					  connStr.execute(q)
				  end if 
				End If
			  else
				if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 ) then
				  err.number = 1001
				  err.description = "Anda harus input Charge Job, Charge Type dan Charge Code! "
				else
				  Dim val_Charge : val_Charge = 0
				  val_Charge = CheckCharge(chargetype, chargecode, chargejob, businessunit, fddate, sourcetype)
				  If val_Charge <> "Success" Then
					Err.Number = 1001
					Err.Description = val_Charge
				  Else
					'-- Get Data Max Linenokey
					querymaxlineno = " select nvl(max(linenokey), 0) linenokey" & _ 
							 " from ATTENDANCE_GAD_TEMP " & _ 
							 " where fcba = '" & Session("fcba") & "'"

					Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
					rsmaxlineno.CursorType = 3
					rsmaxlineno.CursorLocation = 3
					
					'response.write querymaxlineno
					
					rsmaxlineno.Open querymaxlineno, connstr
					if not rsmaxlineno.eof then
					   maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
					else
					   maxlinenokey = 0
					end if
					rsmaxlineno.close
					set rsmaxlineno = Nothing
					
					maxlinenokey = maxlinenokey + 1
					q = " SELECT COUNT (*) jobcodespbharvesting" & _
					"  FROM parameterdetail" & _
					" WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
					"       AND fcba = '"& Session("fcba") &"'" & _
					"       AND fccode = '"& request("jobcode_"&TRIM(n)) &"'"

				  ' response.Write q

				  set rsrecurtrans = server.CreateObject("Adodb.recordset")
				  rsrecurtrans.Open q, connstr

				  if Not rsrecurtrans.eof then
					jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
				  end if
				  rsrecurtrans.Close
				  set rsrecurtrans = Nothing
				  
				  if cdbl(to_nol(jobcodespbharvesting)) > 0 then
					q = " Insert into ATTENDANCE_GAD_TEMP  (GANGCODE" & _
						"                             , FDDATE" & _
						"                             , SUPERVISION_1" & _
						"                             , SUPERVISION_2" & _
						"                             , SUPERVISION_3" & _
						"                             , SUPERVISION_4" & _
						"                             , EMPLOYEECODE" & _
						"                             , ATTENDANCE" & _
						"                             , JOBCODE" & _
						"                             , LOCATIONTYPE" & _
						"                             , LOCATIONCODE" & _
						"                             , CHARGEJOB" & _
						"                             , CHARGETYPE" & _
						"                             , CHARGECODE" & _
						"                             , MANDAYS" & _
						"                             , OTHRS" & _
						"                             , RATE" & _
						"                             , UNIT" & _
						"                             , OUTPUT" & _
						"                             , REFERENCE" & _
						"                             , REMARKS" & _
						"                             , BUCKET" & _
						"                             , FCENTRY" & _
						"                             , FCEDIT" & _
						"                             , FCIP" & _
						"                             , FCBA" & _
						"                             , LASTUPDATE" & _
						"                             , LASTTIME" & _
						"                             , linenokey" & _
						"                             , RowState" & _
						"                             , document_classification" & _
						"                             , TYPE_OVERTIME" & _
						"                             , OVERTIME_HOURS" & _
						"                             , KG_BRONDOLAN" & _
						"                             , SPBNO" & _
						"                             , KG_JANJANG" & _
						" ) values (trim('" & GANGCODE & "')" & _
						"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
						"           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
						"           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
						"           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
						"           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
						"           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
						"           , trim('" & request("rate_"&TRIM(n)) & "')" & _
						"           , trim('" & request("unit_"&TRIM(n)) & "')" & _
						"           , trim('" & request("output_"&TRIM(n)) & "')" & _
						"           , '" & request("reference_"&TRIM(n)) & "'" & _
						"           , '" & request("remarks_"&TRIM(n)) & "'" & _
						"           , '" & request("bucket_"&TRIM(n)) & "'" & _
						"           , '" & Session("CodeName") & "'" & _
						"           , '" & Session("CodeName") & "'" & _
						"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						"           , '" & Session("fcba") & "'" & _
						"           , sysdate" & _
						"           , To_Char(sysdate,'HH24:MI')" & _
						"           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD_TEMP  where fcba = '"& Session("fcba") &"')" & _
						"           , '"& rowstate &"'" & _
						"           , '" & request("document_classification_"&TRIM(n)) & "'" & _
						"           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
						"           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
						"           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
						"           , '" & request("spbno_"&TRIM(n)) & "'" & _
						"           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
						" )"
					  ' response.Write(q)
					  connStr.execute(q)
				  else
					q = " Insert into ATTENDANCE_GAD_TEMP  (GANGCODE" & _
						"                             , FDDATE" & _
						"                             , SUPERVISION_1" & _
						"                             , SUPERVISION_2" & _
						"                             , SUPERVISION_3" & _
						"                             , SUPERVISION_4" & _
						"                             , EMPLOYEECODE" & _
						"                             , ATTENDANCE" & _
						"                             , JOBCODE" & _
						"                             , LOCATIONTYPE" & _
						"                             , LOCATIONCODE" & _
						"                             , CHARGEJOB" & _
						"                             , CHARGETYPE" & _
						"                             , CHARGECODE" & _
						"                             , MANDAYS" & _
						"                             , OTHRS" & _
						"                             , RATE" & _
						"                             , UNIT" & _
						"                             , OUTPUT" & _
						"                             , REFERENCE" & _
						"                             , REMARKS" & _
						"                             , BUCKET" & _
						"                             , FCENTRY" & _
						"                             , FCEDIT" & _
						"                             , FCIP" & _
						"                             , FCBA" & _
						"                             , LASTUPDATE" & _
						"                             , LASTTIME" & _
						"                             , linenokey" & _
						"                             , RowState" & _
						"                             , document_classification" & _
						"                             , TYPE_OVERTIME" & _
						"                             , OVERTIME_HOURS" & _
						"                             , KG_BRONDOLAN" & _
						"                             , SPBNO" & _
						"                             , KG_JANJANG" & _
						" ) values (trim('" & GANGCODE & "')" & _
						"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
						"           , trim('" & request("supervision_1_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_2_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_3_"&TRIM(n)) & "')" & _
						"           , trim('" & request("supervision_4_"&TRIM(n)) & "')" & _
						"           , trim('" & request("employeecode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("attendance_"&TRIM(n)) & "')" & _
						"           , trim('" & request("jobcode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("locationtype_"&TRIM(n)) & "')" & _
						"           , trim('" & request("locationcode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargejob_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargetype_"&TRIM(n)) & "')" & _
						"           , trim('" & request("chargecode_"&TRIM(n)) & "')" & _
						"           , trim('" & request("mandays_"&TRIM(n)) & "')" & _
						"           , trim('" & request("othrs_"&TRIM(n)) & "')" & _
						"           , trim('" & request("rate_"&TRIM(n)) & "')" & _
						"           , trim('" & request("unit_"&TRIM(n)) & "')" & _
						"           , trim('" & request("output_"&TRIM(n)) & "')" & _
						"           , '" & request("reference_"&TRIM(n)) & "'" & _
						"           , '" & request("remarks_"&TRIM(n)) & "'" & _
						"           , '" & request("bucket_"&TRIM(n)) & "'" & _
						"           , '" & Session("CodeName") & "'" & _
						"           , '" & Session("CodeName") & "'" & _
						"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
						"           , '" & Session("fcba") & "'" & _
						"           , sysdate" & _
						"           , To_Char(sysdate,'HH24:MI')" & _
						"           , (select nvl(max(linenokey), 0) + 1 linenokey from ATTENDANCE_GAD_TEMP  where fcba = '"& Session("fcba") &"')" & _
						"           , '"& rowstate &"'" & _
						"           , '" & request("document_classification_"&TRIM(n)) & "'" & _
						"           , '" & request("type_overtime_"&TRIM(n)) & "'" & _
						"           , '" & request("overtime_hours_"&TRIM(n)) & "'" & _
						"           , '" & request("kg_brondolan_"&TRIM(n)) & "'" & _
						"           , '" & request("spbno_"&TRIM(n)) & "'" & _
						"           , '" & request("kg_janjang_"&TRIM(n)) & "'" & _
						" )"
					  ' response.Write(q)
					  connStr.execute(q) 
				  end if 
				  End If
				End If
			  End If
		End If
    next
	
    if err.number  = 0 then
      connStr.CommitTrans
    end if
      
    cmdView = "True"

  elseif Request("txtJenis")="RejectAction" then
			q = " INSERT INTO ag_authorization_comments ( fddate" & _
    			"                                         , gangcode" & _
          "                                         , currentauthorizelevel" & _
    			"                                         , comments" & _
    			"                                         , fcentry" & _
    			"                                         , fcedit" & _
    			"                                         , fcip" & _
    			"                                         , fcba" & _
    			"                                         , lastupdate" & _
    			"                                         , lasttime)"& _
    			" VALUES ( to_date('" & fdDate& "','mm/dd/yyyy')" & _
    			"          , '" & gangcode & "'" & _
          "          , '" & authorizelevelcode & "'" & _
    			"          , '" & comments & "'" & _
    			"          , '" & Session("CodeName") & "'" & _
    			"          , '" & Session("CodeName") & "'" & _
    			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
    			"          , '" & Session("fcba") & "'" & _
    			"          , SYSDATE" & _
    			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
    		'response.write q
    		connStr.execute(q)
		
		
    	  strQueryStatus = "BEGIN ATTENDANCEGAD_API.NextState('" & Session("fcba") & "',to_date('" & fddate & "','mm/dd/yyyy'),'"& gangcode &"','Reject' ); END; "
        connStr.execute(strQueryStatus)
        
        strQueryStatus = " UPDATE attendance_gad SET rowstate = 'Rejected' WHERE fcba = '" & Session("fcba") & "'" & _
                 " AND fddate = to_date('" & fddate & "','mm/dd/yyyy') AND gangcode = '"& gangcode &"'"
        connStr.execute(strQueryStatus)

        strQueryStatus = " UPDATE attendance_gad_temp SET rowstate = 'Rejected' WHERE fcba = '" & Session("fcba") & "'" & _
                 " AND fddate = to_date('" & fddate & "','mm/dd/yyyy') AND gangcode = '"& gangcode &"'"
        connStr.execute(strQueryStatus)
    		
  		cmdView	= "True"
    else
        
	end if								 
	cmdView	= "True"
 else
    cmdView	= request("cmdView")
 End if

 set rsPurchase = server.CreateObject("Adodb.recordset")
 q=	" select distinct fddate, gangcode, rowstate FROM attendance_gad_temp WHERE fddate = to_date('"& fddate &"','mm/dd/yyyy') and gangcode = '"& gangcode &"' and fcba = '"& Session("fcba") &"'"

  ' response.write q
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
 	  fddate=rsPurchase("fddate")
    gangcode=rsPurchase("gangcode")
    rowstate=rsPurchase("rowstate")
 End if
  'response.write "IsInBridge--> " & IsInBridge
  ' response.write err.number
if err.number <> 0 then
	connStr.RollbackTrans
  response.Write  "<div class='error'>" & err.description & "</div>"
End if

has_approveddata = 0
q = " SELECT COUNT(1) total FROM attendance_gad_hist WHERE fddate = to_date('" & vouchercode & "','mm/dd/yyyy')  AND gangcode = '" & gangcode & "' AND fcba = '" & Session("fcba") & "'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approveddata = cdbl(to_nol(rsPurchase("total")))
End if

has_approvedby = 0
q = "  SELECT COUNT(1) total FROM attendance_gad_hist WHERE fddate = to_date('" & vouchercode & "','mm/dd/yyyy')  AND gangcode = '" & gangcode & "' AND fcba = '" & Session("fcba") & "'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approvedby = cdbl(to_nol(rsPurchase("total")))
End if

q = " SELECT * FROM ag_authorization_comments" & _
    " WHERE fddate = to_date('" & vouchercode & "','mm/dd/yyyy')" & _
    "       AND gangcode = '" & gangcode & "'" & _
    "       AND fcba = '" & Session("fcba") & "'" & _
	  "       AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	comments = rsPurchase("comments")
End if
 '-------------ISI Payment Form
%>
  <form name="frmEntry" method="post" action="AttendanceGADApproval.asp">
	<table class="data" width="100%">
		<tr>
		  <td align="left">Date</td>
		  <td align="left"><input type="text" class="datepicker" maxlength="15" name="fddate" size="30" value="<%=fddate%>"><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
		  <td align="left">Gang</td>
		  <td align="left">
          <select name="GangCode">
            <%cmdCallCombo "Gang",GangCode,"",""%>
          </select>
      </td>
		</tr>
		<tr>
		  <td align="left">Status</td>
		  <td align="left"><input type="text" maxlength="50" name="rowstate" size="30" value="<%=rowstate%>" style="background-color:#eee;" readonly><font color="#FF0000">&nbsp;*</font></td>
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
        <button type="submit" class="button save" name="txtJenis" value="SaveData" onClick="return confirm('Are you sure to Save this data : <%=fddate%>,<%=gangcode%> ?');"><i class="fa fa-save"></i> Save</button>
				<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Are you sure to Approve this data : <%=fddate%>,<%=gangcode%> ?');"><i class="fa fa-check"></i> Approve</button>
				<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Are you sure to Reject this data : <%=fddate%>,<%=gangcode%> ?');"><i class="fa fa-times"></i> Reject</button>
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

	q = "   SELECT DISTINCT stage1.businessunit fcba," & _
		"          stage1.gangcode," & _
		"          stage1.fddate," & _
		"          gad.supervision_1," & _
		"          gad.supervision_2," & _
		"          gad.supervision_3," & _
		"          gad.supervision_4," & _
		"          gad.employeecode," & _
		"          gad.attendance," & _
		"          gad.jobcode," & _
		"          gad.locationtype," & _
		"          gad.locationcode," & _
		"          gad.chargejob," & _
		"          gad.chargetype," & _
		"          gad.chargecode," & _
		"          gad.mandays," & _
		"          gad.unit," & _
		"          gad.bucket," & _
		"          gad.output," & _
		"          gad.rate," & _
		"          gad.kg_janjang," & _
		"          gad.kg_brondolan," & _
		"          gad.spbno," & _
		"          gad.overtime_hours," & _
		"          gad.type_overtime," & _
		"          gad.othrs AS overtime_convertion," & _
		"          gad.LineNoKey," & _
		"          gad.document_classification," & _
		"          stage1.fccode authorizetemplatecode," & _
		"          stage1.document_classification," & _
		"          stage1.currentauthorizelevel," & _
		"          stage1.authorizelevelname," & _
		"          stage1.lastauthorizer," & _
		"          stage1.lastactiondate," & _
		"          stage1.authorizelevelcode," & _
		"          aust.fcusercode," & _
		"          stage1.rowstate" & _
		"     FROM (SELECT abc.businessunit," & _
		"                  abc.gangcode," & _
		"                  abc.fddate," & _
		"                  at.fccode," & _
		"                  abc.document_classification," & _
		"                  abc.currentauthorizelevel," & _
		"                  abc.authorizelevelname," & _
		"                  abc.lastauthorizer," & _
		"                  abc.lastactiondate," & _
		"                  authorizestructure_api.get_nextlevelag (" & _
		"                     'AA'," & _
		"                     abc.document_classification," & _
		"                     abc.businessunit," & _
		"                     NVL (abc.currentauthorizelevel, 0)," & _
		"                     abc.fddate," & _
		"                     abc.gangcode," & _
		"                     abc.lastauthorizer)" & _
		"                     AS authorizelevelcode," & _
		"                  rowstate" & _
		"             FROM (  SELECT temp.fcba AS businessunit," & _
		"                            temp.gangcode AS gangcode," & _
		"                            temp.fddate AS fddate," & _
		"                            temp.document_classification," & _
		"                            authorizelevel.fccode AS currentauthorizelevel," & _
		"                            authorizelevel.fcname AS authorizelevelname," & _
		"                            attendance_gad_hist.fcentry AS lastauthorizer," & _
		"                            attendance_gad_hist.lastupdate AS lastactiondate," & _
		"                            temp.rowstate" & _
		"                       FROM attendance_gad_hist" & _
		"                            INNER JOIN" & _
		"                            (SELECT attendance_gad_temp.fcba," & _
		"                                    attendance_gad_temp.gangcode," & _
		"                                    attendance_gad_temp.fddate," & _
		"                                    attendance_gad_temp.rowstate," & _
		"                                    attendance_gad_temp.document_classification" & _
		"                               FROM attendance_gad_temp" & _
		"                              WHERE attendance_gad_temp.fcba = '"& session("fcba") &"') temp" & _
		"                               ON (    attendance_gad_hist.fcba = temp.fcba" & _
		"                                   AND attendance_gad_hist.gangcode =" & _
		"                                          temp.gangcode" & _
		"                                   AND attendance_gad_hist.fddate = temp.fddate)" & _
		"                            INNER JOIN authorizelevel" & _
		"                               ON attendance_gad_hist.currentauthorizelevel =" & _
		"                                     authorizelevel.fccode" & _
		"                            INNER JOIN" & _
		"                            (  SELECT attendance_gad_hist.fcba," & _
		"                                      attendance_gad_hist.gangcode," & _
		"                                      attendance_gad_hist.fddate," & _
		"                                      MAX (" & _
		"                                         attendance_gad_hist.currentauthorizelevel)" & _
		"                                         alevel" & _
		"                                 FROM attendance_gad_hist" & _
		"                             GROUP BY attendance_gad_hist.fcba," & _
		"                                      attendance_gad_hist.gangcode," & _
		"                                      attendance_gad_hist.fddate) aa" & _
		"                               ON (    attendance_gad_hist.fcba = aa.fcba" & _
		"                                   AND attendance_gad_hist.currentauthorizelevel =" & _
		"                                          aa.alevel" & _
		"                                   AND attendance_gad_hist.fddate = aa.fddate" & _
		"                                   AND attendance_gad_hist.gangcode = aa.gangcode)" & _
		"                      WHERE temp.fcba = '"& session("fcba") &"'" & _
		"                   GROUP BY temp.fcba," & _
		"                            temp.gangcode," & _
		"                            temp.fddate," & _
		"                            temp.document_classification," & _
		"                            authorizelevel.fccode," & _
		"                            authorizelevel.fcname," & _
		"                            attendance_gad_hist.fcentry," & _
		"                            attendance_gad_hist.lastupdate," & _
		"                            temp.rowstate) abc" & _
		"                  LEFT JOIN iplasprod.authorizetemplate at" & _
		"                     ON abc.document_classification = at.document_classification" & _
		"                  LEFT JOIN iplasprod.authorizestructure aus" & _
		"                     ON (    at.fccode = aus.authorizetemplatecode" & _
		"                         AND abc.currentauthorizelevel = aus.authorizelevelcode" & _
		"                         AND abc.businessunit = aus.fcba" & _
		"                         AND aus.fcusercode = abc.lastauthorizer)) stage1" & _
		"          INNER JOIN iplasprod.authorizestructure aust" & _
		"             ON (    stage1.businessunit = aust.fcba" & _
		"                 AND stage1.authorizelevelcode = aust.authorizelevelcode" & _
		"                 AND stage1.fccode = aust.authorizetemplatecode)" & _
		"          INNER JOIN attendance_gad_temp gad" & _
		"             ON (    stage1.businessunit = gad.fcba" & _
		"                 AND stage1.gangcode = gad.gangcode" & _
		"                 AND stage1.fddate = gad.fddate)" & _
		"    WHERE stage1.rowstate = 'AuthorizeOnProgress'" & _
		"    AND gad.fddate = to_date('"& fddate &"','mm/dd/yyyy') AND gad.gangcode = '"& gangcode &"' AND gad.fcba = '"& Session("fcba") &"'" & _
		" ORDER BY stage1.businessunit," & _
		"          stage1.FDDATE," & _
		"          stage1.GANGCODE," & _
		"          gad.employeecode," & _
		"          gad.jobcode," & _
		"          gad.locationtype," & _
		"          gad.locationcode"  

	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.Open q, connstr
%>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" > 
	<%if RowStateId <> "Rejected" then%>
		<%if has_approvedby=0 then%>
			<TD align="left">NO</TD> 
			<TD align="left">LINE NO</TD> 
			<TD align="left">SUP_1</TD> 
			<TD align="left">SUP_2</TD> 
			<TD align="left">SUP_3</TD> 
			<TD align="left">SUP_4</TD> 
			<TD align="left">EMPLOYEE CODE</TD> 
			<TD align="left">ATTENDANCE</TD>  
			<TD align="left">JOB CODE</TD>  
			<TD align="left">LOC TYPE</TD>  
			<TD align="left">LOC CODE</TD>  
			<TD align="left">CHARGE JOB</TD>  
			<TD align="left">CHARGE TYPE</TD> 
			<TD align="left">CHARGE CODE</TD> 
			<TD align="left">MANDAYS</TD>   
			<TD align="left">UNIT</TD>    
			<TD align="left">BUCKET</TD>
			<TD align="left">SPBNO</TD>
			<TD align="left">KG BRONDOLAN</TD>
			<TD align="left">KG JANJANG</TD>
			<TD align="left">OUTPUT</TD>    
			<TD align="left">RATE</TD>   
			<TD align="left">OVERTIME HOURS</TD>
			<TD align="left">TYPE OVERTIME</TD>
			<TD align="left">OVERTIME CONVERTION</TD>      
		<%else%>
			<TD align="left">NO</TD> 
			<TD align="left">LINE NO</TD>
			<TD align="left">SUP_1</TD> 
			<TD align="left">SUP_2</TD> 
			<TD align="left">SUP_3</TD> 
			<TD align="left">SUP_4</TD> 
			<TD align="left">EMPLOYEE CODE</TD> 
			<TD align="left">ATTENDANCE</TD>  
			<TD align="left">JOB CODE</TD>  
			<TD align="left">LOC TYPE</TD>  
			<TD align="left">LOC CODE</TD>  
			<TD align="left">CHARGE JOB</TD>  
			<TD align="left">CHARGE TYPE</TD> 
			<TD align="left">CHARGE CODE</TD> 
			<TD align="left">MANDAYS</TD>  
			<TD align="left">UNIT</TD>    
			<TD align="left">BUCKET</TD>
			<TD align="left">SPBNO</TD>
			<TD align="left">KG BRONDOLAN</TD>
			<TD align="left">KG JANJANG</TD>
			<TD align="left">OUTPUT</TD>    
			<TD align="left">RATE</TD>  
			<TD align="left">OVERTIME HOURS</TD>
			<TD align="left">TYPE OVERTIME</TD>
			<TD align="left">OVERTIME CONVERTION</TD>
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
			<%if RowStateId <> "Rejected" then%>
  				<%if has_approvedby=0 then%> 
					<td align=" centers" title="<%=rsPurchaseOrder("linenokey").Name%>" valign="top">
						<%linenokey = "0":if has_approveddata = 0 then linenokey = rsPurchaseOrder("linenokey") else linenokey = rsPurchaseOrder("linenokey")%>
  						<input type="hidden" style="text-align:right;" name="linenokey_<%=no%>" width="2" value="<%=linenokey%>" >
						<%=rsPurchaseOrder("linenokey")%>
					</td>
  					<td align="right" title="Supervision 1" valign="top">
  						<%supervision_1 = "0":if has_approveddata = 0 then supervision_1 = rsPurchaseOrder("supervision_1") else supervision_1 = rsPurchaseOrder("supervision_1")%>
  						<input type="text" style="text-align:right;" name="supervision_1_<%=no%>" width="2" value="<%=supervision_1%>" >  
  					</td> 
  					<td align="right" title="Supervision 2" valign="top">
						<%supervision_2 = "0":if has_approveddata = 0 then supervision_2 = rsPurchaseOrder("supervision_2") else supervision_2 = rsPurchaseOrder("supervision_2")%>
					  <input type="text" style="text-align:right;" name="supervision_2_<%=no%>" width="2" value="<%=supervision_2%>" >  
					</td> 
					<td align="right" title="Supervision 3" valign="top">
						<%supervision_3 = "0":if has_approveddata = 0 then supervision_3 = rsPurchaseOrder("supervision_3") else supervision_3 = rsPurchaseOrder("supervision_3")%>
						<input type="text" style="text-align:right;" name="supervision_3_<%=no%>" width="2" value="<%=supervision_3%>" >  
					</td> 
					<td align="right" title="Supervision 4" valign="top">
						<%supervision_4 = "0":if has_approveddata = 0 then supervision_4 = rsPurchaseOrder("supervision_4") else supervision_4 = rsPurchaseOrder("supervision_4")%>
						<input type="text" style="text-align:right;" name="supervision_4_<%=no%>" width="2" value="<%=supervision_4%>" >  
					</td> 
					<td align="right" title="Employeecode" valign="top">
						<%employeecode = "0":if has_approveddata = 0 then employeecode = rsPurchaseOrder("employeecode") else employeecode = rsPurchaseOrder("employeecode")%>
						<input type="text" style="text-align:right;" name="employeecode_<%=no%>" width="2" value="<%=employeecode%>" >  
					</td> 
  					<td align="right" title="attendance" valign="top">
  						<%attendance = "0":if has_approveddata = 0 then attendance = rsPurchaseOrder("attendance") else attendance = rsPurchaseOrder("attendance")%>
  						<input type="text" style="text-align:right;" name="attendance_<%=no%>" width="2" value="<%=attendance%>" > 
  					</td>
  					<td align="right" title="Jobcode" valign="top">
  						<%jobcode = "0":if has_approveddata = 0 then jobcode = rsPurchaseOrder("jobcode") else jobcode = rsPurchaseOrder("jobcode")%>
  						<input type="text" style="text-align:right;" name="jobcode_<%=no%>" width="2" value="<%=jobcode%>" > 
  						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
  					<font face="Verdana" size="1">Activity</font>
  					</td> 
  					<td align="right" title="LocationType" valign="top">
  						<%locationtype = "0":if has_approveddata = 0 then locationtype = rsPurchaseOrder("locationtype") else locationtype = rsPurchaseOrder("locationtype")%>
  						<input type="text" style="text-align:right;" name="locationtype_<%=no%>" width="2" value="<%=locationtype%>" > 
  						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
						<font face="Verdana" size="1">Location Type</font></a>
  					</td>
  					<td align="right" title="LocationCode" valign="top">
  						<%locationcode = "0":if has_approveddata = 0 then locationcode = rsPurchaseOrder("locationcode") else locationcode = rsPurchaseOrder("locationcode")%>
  						<input type="text" style="text-align:right;" name="locationcode_<%=no%>" width="2" value="<%=locationcode%>" > 
  					</td> 
					<td align="right" title="ChargeJob" valign="top">
						<%chargejob = "0":if has_approveddata = 0 then chargejob = rsPurchaseOrder("chargejob") else chargejob = rsPurchaseOrder("chargejob")%>
						<input type="text" style="text-align:right;" name="chargejob_<%=no%>" width="2" value="<%=chargejob%>" > 
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
						<font face="Verdana" size="1">Activity</font>
					</td> 
					<td align="right" title="ChargeType" valign="top">
						<%chargetype = "0":if has_approveddata = 0 then chargetype = rsPurchaseOrder("chargetype") else chargetype = rsPurchaseOrder("chargetype")%>
						<input type="text" style="text-align:right;" name="chargetype_<%=no%>" width="2" value="<%=chargetype%>" > 
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
						<font face="Verdana" size="1">Location Type</font></a>
					</td>
					<td align="right" title="ChargeCode" valign="top">
						<%chargecode = "0":if has_approveddata = 0 then chargecode = rsPurchaseOrder("chargecode") else chargecode = rsPurchaseOrder("chargecode")%>
						<input type="text" style="text-align:right;" name="chargecode_<%=no%>" width="2" value="<%=chargecode%>" > 
					</td> 
					<td align="right" valign="top" >
						<%mandays = "0":if has_approveddata = 0 then mandays = rsPurchaseOrder("mandays") else mandays = rsPurchaseOrder("mandays")%>
						<input type="text" style="text-align:right;" name="mandays_<%=no%>" width="2" value="<%=mandays%>" > 
					</td>
					<td align="right" valign="top" >
						<%unit = "0":if has_approveddata = 0 then unit = rsPurchaseOrder("unit") else unit = rsPurchaseOrder("unit")%>
						<input type="text" style="text-align:right;" name="unit_<%=no%>" width="2" value="<%=unit%>" > 
					</td>
					<td align="right" valign="top" >
						<%bucket = "0":if has_approveddata = 0 then bucket = rsPurchaseOrder("bucket") else bucket = rsPurchaseOrder("bucket")%>
						<input type="text" style="text-align:right;" name="bucket_<%=no%>" width="2" value="<%=bucket%>" > 
					</td>
					<td align="right" valign="top" >
						<%spbno = "0":if has_approveddata = 0 then spbno = rsPurchaseOrder("spbno") else spbno = rsPurchaseOrder("spbno")%>
						<input type="text" style="text-align:right;" name="spbno_<%=no%>" width="2" value="<%=spbno%>" > 
					</td>
					<td align="right" valign="top" >
						<%kg_janjang = "0":if has_approveddata = 0 then kg_janjang = rsPurchaseOrder("kg_janjang") else kg_janjang = rsPurchaseOrder("kg_janjang")%>
						<input type="text" style="text-align:right;" name="kg_janjang_<%=kg_janjang%>" width="2" value="<%=kg_janjang%>" > 
					</td>
					<td align="right" valign="top" >
						<%kg_brondolan = "0":if has_approveddata = 0 then kg_brondolan = rsPurchaseOrder("kg_brondolan") else kg_brondolan = rsPurchaseOrder("kg_brondolan")%>
						<input type="text" style="text-align:right;" name="kg_brondolan_<%=kg_brondolan%>" width="2" value="<%=kg_brondolan%>" > 
					</td>
					<td align="right" valign="top" >
						<%output = "0":if has_approveddata = 0 then output = rsPurchaseOrder("output") else output = rsPurchaseOrder("output")%>
						<input type="text" style="text-align:right;" name="output_<%=no%>" width="2" value="<%=output%>" > 
					</td>
					<td align="right" valign="top" >
						<%rate = "0":if has_approveddata = 0 then rate = rsPurchaseOrder("rate") else rate = rsPurchaseOrder("rate")%>
						<input type="text" style="text-align:right;" name="rate_<%=no%>" width="2" value="<%=rate%>" > 
					</td>
					<td align="right" valign="top" >
						<%overtime_hours = "0":if has_approveddata = 0 then overtime_hours = rsPurchaseOrder("overtime_hours") else overtime_hours = rsPurchaseOrder("overtime_hours")%>
						<input type="text" style="text-align:right;" name="overtime_hours_<%=no%>" width="2" value="<%=overtime_hours%>" > 
					</td>
					<td align="right" valign="top" >
						<%type_overtime = "0":if has_approveddata = 0 then type_overtime = rsPurchaseOrder("type_overtime") else type_overtime = rsPurchaseOrder("type_overtime")%>
						<input type="text" style="text-align:right;" name="type_overtime_<%=no%>" width="2" value="<%=type_overtime%>" > 
					</td>
					<td align="right" valign="top" >
						<%overtime_convertion = "0":if has_approveddata = 0 then overtime_convertion = rsPurchaseOrder("overtime_convertion") else overtime_convertion = rsPurchaseOrder("overtime_convertion")%>
						<input type="text" style="text-align:right;" name="overtime_convertion_<%=no%>" width="2" value="<%=overtime_convertion%>" > 
					</td>
					<%document_classification = "0":if has_approveddata = 0 then document_classification = rsPurchaseOrder("document_classification") else document_classification = rsPurchaseOrder("document_classification")%>
  					<input type="hidden" style="text-align:right;" name="document_classification_<%=no%>" width="2" value="<%=document_classification%>" >
  				<%else%>
  					<td align="left" title="<%=rsPurchaseOrder("jobcode").Name%>" valign="top"><%=rsPurchaseOrder("jobcode")%> : <%=rsPurchaseOrder("jobname")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("locationtype").Name%>" valign="top"><%=rsPurchaseOrder("locationtype")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("locationcode").Name%>" valign="top"><%=rsPurchaseOrder("locationcode")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("chargejob").Name%>" valign="top"><%=rsPurchaseOrder("chargejob")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("chargetype").Name%>" valign="top"><%=rsPurchaseOrder("chargetype")%></td>
  					<td align="left" title="<%=rsPurchaseOrder("chargecode").Name%>" valign="top"><%=rsPurchaseOrder("chargecode")%></td>
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
