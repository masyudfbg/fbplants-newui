<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Purchase Request Approval</title>
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
 
 PRCODE = Request("PRCODE")
 
 If Len(Trim(PRCODE)) = 0 Then PRCODE = "(NEW)"
 
 '---- Data Header
 if Request("PRDATE")<>"" then
   PRDATE   =Request("PRDATE")
 else
   PRDATE   =textDate(date)
 end if
 PRDATE = FormatDateTime(PRDATE)
 'response.write "1." & PRDATE
 'response.write "2." & FormatDateTime(PRDATE)
 
 Description = Request("Description")
 Remarks = Request("Remarks")
 RequestorId = Request("RequestorId")
 ManagedId = Request("ManagedId")
 PurchaseGroupId = Request("PurchaseGroupId")
 RowStateId = Request("RowStateId")
 document_classification = Request("document_classification")
 
 
 
 '---- Data Details
 
 LINEID=Request("LineId")
 ITEMCODE=Request("ITEMCODE")
 REQUESTEDQTY=Request("REQUESTEDQTY")
 APPROVEDQTY=Request("APPROVEDQTY")
 ORDEREDQTY=Request("ORDEREDQTY")
 DESCRIPTIONDETAIL = Request("DESCRIPTIONDETAIL")
 WANTEDDELIVERYDATE = Request ("WANTEDDELIVERYDATE")
 UnitOfMeasurement = Request("UnitOfMeasurement")
 REMARKS=request("REMARKS")
 comments = request("comments")
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")= "ApprovalAction" or Request("txtJenis")= "RejectAction" Then

	if Request("txtJenis")="ApprovalAction" then
		For n=1 to (cdbl(request("no"))-1)
			q = " INSERT INTO purchaserequestdetails_hist ( prcode" & _
			    "                                           , lineid" & _
			    "                                           , originalapprovedqty" & _
			    "                                           , approvedqty" & _
			    "                                           , currentauthorizelevel" & _
			    "                                           , fcentry" & _
			    "                                           , fcedit" & _
			    "                                           , fcip" & _
			    "                                           , fcba" & _
			    "                                           , lastupdate" & _
			    "                                           , lasttime)"& _
				" VALUES ( '" & prcode & "'" & _
			    "          , '" & request("lineid_"&TRIM(n)) & "'" & _
			    "          , '" & request("originalapprovedqty_"&TRIM(n)) & "'" & _
			    "          , '" & request("suggestedqty_"&TRIM(n)) & "'" & _
			    "          , '" & authorizelevelcode & "'" & _
			    "          , '" & Session("CodeName") & "'" & _
			    "          , '" & Session("CodeName") & "'" & _
			    "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			    "          , '" & fcba & "'" & _
			    "          , SYSDATE" & _
			    "          , TO_CHAR(SYSDATE, 'hh24:mi'))"
			'response.write q
			connStr.execute(q)
			
			q = " UPDATE purchaserequestdetails" & _
			    " SET approvedqty = '"& request("suggestedqty_"&TRIM(n)) & "'" & _
				" WHERE prcode = '" & prcode & "'" & _
			    "       AND lineid = '"& request("lineid_"&TRIM(n)) & "'" & _
			    "       AND fcba = '"& fcba & "'"
			'response.write q
			connStr.execute(q)
		next
		
		q = " INSERT INTO pr_authorization_comments ( prcode" & _
			"                                         , currentauthorizelevel" & _
			"                                         , comments" & _
			"                                         , fcentry" & _
			"                                         , fcedit" & _
			"                                         , fcip" & _
			"                                         , fcba" & _
			"                                         , lastupdate" & _
			"                                         , lasttime)"& _
			" VALUES ( '" & prcode & "'" & _
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
			" WHERE at.document_type = 'PR'" & _
			"       AND at.document_classification = '" & document_classification & "'" & _
			"       AND al.isactive = 'TRUE'" & _
			"       AND al.finallevel = 'TRUE'" & _
			"       AND al.fccode = '" & authorizelevelcode & "'" & _
			"       AND aus.fcba = '" & fcba & "'"

		set rsPurchase = server.CreateObject("Adodb.recordset")
		rsPurchase.Open q, connstr
		IF not rsPurchase.eof then
			is_final = cdbl(to_nol(rsPurchase("total")))
		end if
		'response.write q
		'response.write is_final 
		'response.end
		
		
		if is_final > 1 then
			strQueryStatus = "BEGIN PURCHASEREQUEST_API.NextState('" & fcba & "','" & prcode & "', 'Approve' ); END; "
			connStr.execute(strQueryStatus)
			
			strQueryStatus = "UPDATE purchaserequestdetails SET rowstateid = 'Approved' WHERE fcba = '" & fcba & "' AND prcode = '" & prcode & "'"
			connStr.execute(strQueryStatus)
		end if
		cmdView	= "True"
	elseif Request("txtJenis")="RejectAction" then
			q = " INSERT INTO pr_authorization_comments ( prcode" & _
			    "                                         , currentauthorizelevel" & _
			    "                                         , comments" & _
			    "                                         , fcentry" & _
			    "                                         , fcedit" & _
			    "                                         , fcip" & _
			    "                                         , fcba" & _
			    "                                         , lastupdate" & _
			    "                                         , lasttime)"& _
				" VALUES ( '" & prcode & "'" & _
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
	
		strQueryStatus = "UPDATE purchaserequestdetails SET rowstateid = 'Rejected' WHERE fcba = '" & fcba & "' AND prcode = '" & prcode & "'"
		connStr.execute(strQueryStatus)
		
	    strQueryStatus = "BEGIN PURCHASEREQUEST_API.NextState('" & fcba & "','" & prcode & "', 'Reject' ); END; "
		connStr.execute(strQueryStatus)
		
		cmdView	= "True"
    else
        
	end if								 
	cmdView	= "True"
 else
    cmdView	= request("cmdView")
 End if

 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select * from PURCHASEREQUEST  " & _
  " where PrCode='" & PrCode & "' and " & _
  " fcBA='" & fcba & "'"
  
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    PrDate=rsPurchase("PrDate")
    Description=rsPurchase("Description")
    document_classification=rsPurchase("document_classification")
	RequestorId = rsPurchase("RequestorId")
	PurchaseGroupId = rsPurchase("PurchaseGroupId")
	ManagedId = rsPurchase("ManagedId")	
	Remarks=rsPurchase("Remarks")
	RowStateId = rsPurchase("RowStateId")
 End if
if err.number <> 0 then
    Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

has_approveddata = 0
q = " SELECT COUNT(1) total FROM purchaserequestdetails_hist WHERE prcode = '" & PrCode & "' AND fcba = '" & fcba & "'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approveddata = cdbl(to_nol(rsPurchase("total")))
End if

has_approvedby = 0
q = " SELECT COUNT(1) total FROM purchaserequestdetails_hist WHERE prcode = '" & PrCode & "' AND fcba = '" & fcba & "' AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approvedby = cdbl(to_nol(rsPurchase("total")))
End if

q = " SELECT * FROM pr_authorization_comments" & _
    " WHERE prcode = '" & PrCode & "'" & _
    "       AND fcba = '" & fcba & "'" & _
	"       AND currentauthorizelevel = '"&authorizelevelcode&"'"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	comments = rsPurchase("comments")
End if

 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="PurchaseRequestDetailsApproval.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Purchase Request No</td>
		<td align="left"><input type="text" maxlength="50" name="PrCode" size="30" value="<%=PrCode%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="PRDATE" size="10" value="<%=PRDATE%>">
		<!--<a href="javascript:show_calendar('frmEntry.PRDATE');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
		</td>
		</tr>
		
		<tr>
		<td align="left">Description</td>
		<td align="left"><textarea  name="Description"  cols="50" rows="3"><%=Description%></textarea></td>		
		</tr>
		
		<tr>
		<td align="left">Classification</td>
		<td align="left">
			<select name="document_classification">
			<%
			'qLov = " SELECT fccode, fcname FROM document_classification ORDER BY fccode"
			'Update Aam MAr 21
			qLov = " SELECT fccode, fcname FROM document_classification where document_type= 'PR' and isactive='TRUE' ORDER BY fccode"
   
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
		<td align="left">Requestor</td>
		<td align="left">      
		<select name="RequestorId"  >
	    <%cmdCallCombo "REQUESTOR",RequestorId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>


		
		<tr>
		<td align="left">Managed By</td>
		<td align="left">      
		<select name="ManagedId" >
	    <%cmdCallCombo "PURCHASEMANAGER",ManagedId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>	
		</tr>
		
		<tr>
		<td align="left">Purchase Group By</td>
		<td align="left">      
		<select name="PurchaseGroupId" >
	    <%cmdCallCombo "PURCHASEGROUP",PurchaseGroupId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>	
		</tr>		
		
		
		<tr>
		<td align="left">Remarks</td>
		<td align="left"><textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea></td>		
		</tr>	

		<tr>
		<td align="left">Status</td>
		<td align="left"><input disabled type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr>
		<td align="left">Comments</td>
		<td align="left"><input type="text" maxlength="200" name="comments" size="50" value="<%=comments%>">
		</td>
		</tr>

		
		
		</td>
		</tr>

	

		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Hidden" Value="<%=fcba%>" Name="fcba">
				<input Type="Hidden" Value="<%=authorizelevelcode%>" Name="authorizelevelcode">
				<%if RowStateId <> "Rejected" then%>
					<%if has_approvedby=0 then%>
						<button type="submit" class="button" name="txtJenis" value="ApprovalAction" onClick="return confirm('Are you sure to Approve this data : <%=PrCode%> ?');"><i class="fa fa-check"></i> Approve</button>
						<button type="submit" class="button danger" name="txtJenis" value="RejectAction" onClick="return confirm('Are you sure to Reject this data : <%=PrCode%> ?');"><i class="fa fa-times"></i> Reject</button>
					<%else%>
						<b>This PR number : <%=prcode%> has been Approved</b>
					<%end if%>
				<%else%>
					<b>This PR number : <%=prcode%> has been Rejected</b>
				<%end if%>
			</td>     
		</tr>
    </table>
<%
	q = " SELECT prcode, lineid, itemcode,stock_api.get_stockname(itemcode,fcba) itemname, requestedqty, approvedqty, orderedqty, unitofmeasurement, wanteddeliverydate, description, fcentry, fcedit, fcip, fcba" & _
		" FROM purchaserequestdetails" & _
		" WHERE PrCode='" & PrCode & "'" & _
		"       AND fcBA='" & fcba & "'" & _
		" ORDER BY lineid"
   
   Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
   rsPurchaseRequest.Open q, connstr
%>
	<table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">No</TD>
	<TD align="left">PR Code</TD>	
	<TD align="left">Line ID</TD>	
	<TD align="left">Item Code</TD>	
	<TD align="left">Item Name</TD>	
	<TD align="left">Requested Qty</TD>	
	<TD align="left">Approved Qty</TD>
	<%if RowStateId <> "Rejected" then%>
	<%if has_approvedby=0 then%>
	<TD align="left">Suggested Approved Qty</TD>
	<%end if%>
	<%end if%>
	<TD align="left">Uom</TD>	
	<TD align="left">Wanted Delivery Date</TD>	
	<TD align="left">Description</TD>
	</tr>   
<%
   if Not rsPurchaseRequest.eof then
      nCountofField=rsPurchaseRequest.fields.count-1
	  no=1
	  While Not rsPurchaseRequest.EOF

		if no  mod 2 = 1 then
			cColor = "#EEEEEE"
		else
			cColor = "#E5E5E5"
		end if%>
		<tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=no%></td>
			<td align="left" title="<%=rsPurchaseRequest("prcode").Name%>" valign="top"><%=rsPurchaseRequest("prcode")%></td>
			<td align="left" title="<%=rsPurchaseRequest("lineid").Name%>" valign="top"><%=rsPurchaseRequest("lineid")%></td>
			<td align="left" title="<%=rsPurchaseRequest("itemcode").Name%>" valign="top"><%=rsPurchaseRequest("itemcode")%></td>
			<td align="left" title="<%=rsPurchaseRequest("itemname").Name%>" valign="top"><%=rsPurchaseRequest("itemname")%></td>
			<td align="right" title="<%=rsPurchaseRequest("requestedqty").Name%>" valign="top"><%=rsPurchaseRequest("requestedqty")%></td>
			<td align="right" title="<%=rsPurchaseRequest("approvedqty").Name%>" valign="top"><%=rsPurchaseRequest("approvedqty")%></td>
			<%if RowStateId <> "Rejected" then%>
			<%if has_approvedby=0 then%>
			<td align="right" title="Suggested Approved Qty" valign="top">
				<%suggestedqty = "0":if has_approveddata = 0 then suggestedqty = rsPurchaseRequest("requestedqty") else suggestedqty = rsPurchaseRequest("approvedqty")%>
				<input type="text" style="text-align:right;" name="suggestedqty_<%=no%>" width="2" value="<%=suggestedqty%>" onkeyup="checkInput(this);">
				<input type="hidden" name="lineid_<%=no%>" value="<%=rsPurchaseRequest("lineid")%>">
				<input type="hidden" name="originalapprovedqty_<%=no%>" value="<%=rsPurchaseRequest("approvedqty")%>">
			</td>
			<%end if%>
			<%end if%>
			<td align="left" title="<%=rsPurchaseRequest("unitofmeasurement").Name%>" valign="top"><%=rsPurchaseRequest("unitofmeasurement")%></td>
			<td align="left" title="<%=rsPurchaseRequest("wanteddeliverydate").Name%>" valign="top"><%=rsPurchaseRequest("wanteddeliverydate")%></td>
			<td align="left" title="<%=rsPurchaseRequest("description").Name%>" valign="top"><%=rsPurchaseRequest("description")%></td>
		</tr>
<%
		  rsPurchaseRequest.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseRequest.Close
   set rsPurchaseRequest = Nothing%>
   </div>
   </TABLE>
	<input type="hidden" name="no" value="<%=no%>">
	
 <!-- Add Aam Mar 21 Detail Approval Hist -->
 <%
  q = "select DISTINCT to_char(a.lastupdate,'DD-MON-YY HH24:MM') lastupdate,a.lasttime,b.fcname,a.fcentry " & _
	  "from purchaserequestdetails_hist  a  " & _
      "inner join authorizelevel b on a.currentauthorizelevel=b.fccode " & _
      "where PrCode='" & PrCode & "'" & _
	  "       AND fcBA='" & fcba & "'" 
	 
   Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
   rsPurchaseRequest.Open q, connstr
 %>	
	<table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">No</TD>
	<TD align="left">Date Approved</TD>	
	<TD align="left">Level</TD>	
	<TD align="left">Approval By</TD>	
	</tr>   
<%
   if Not rsPurchaseRequest.eof then
      nCountofField=rsPurchaseRequest.fields.count-1
	  no=1
	  While Not rsPurchaseRequest.EOF

		if no  mod 2 = 1 then
			cColor = "#EEEEEE"
		else
			cColor = "#E5E5E5"
		end if%>
		<tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=no%></td>
			<td align="left" title="1" valign="top"><%=rsPurchaseRequest("Lastupdate")%></td>
			<td align="left" title="3" valign="top"><%=rsPurchaseRequest("fcname")%></td>
			<td align="left" title="4" valign="top"><%=rsPurchaseRequest("fcentry")%></td>
		</tr>
<%
		  rsPurchaseRequest.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseRequest.Close
   set rsPurchaseRequest = Nothing%>
   </div>
	</Table>
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
