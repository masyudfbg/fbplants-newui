<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Cancel Purchase Request </title>

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


var getTransporter = function () {
	// build up the post string when passing variables to the server side page

	var lineno = document.getElementById("lineno").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var transportercontractno = document.getElementById("transportercontractno").value;
	var PostStr = 'lineno='+lineno+'&fcba='+selectedfcba+'&transportercontractno='+transportercontractno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetTransporter.asp', 'POST', PostStr, ShowTransporeterResponse);
}

// The function for handling the response from the server
var ShowTransporeterResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.lineno);
    
    // update the Div to show the result from the server
	document.getElementById("transporter").value = s.transporter;
	document.getElementById("losstolerance").value = s.losstolerance;
	document.getElementById("rate").value = s.rate;
	document.getElementById("transportername").value = s.transportername;

}

var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbadosales.asp', 'POST', PostStr, showlistfcba);
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

// launched from button click 
var getTransporterLine = function () {
	// build up the post string when passing variables to the server side page
	
	var transportercontractno = document.getElementById("transportercontractno").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedtransporterline = document.getElementById("selectedtransporterline").value;
	var PostStr = 'transportercontractno='+transportercontractno+'&fcba='+selectedfcba+'&lineno='+selectedtransporterline;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetTransporterLine.asp', 'POST', PostStr, showTransporterLine);
}

// The function for handling the response from the server
var showTransporterLine = function (oXML) { 
    // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("transporterlinelist").innerHTML = "";
	document.getElementById("transporterlinelist").innerHTML = response;

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

<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<link rel="StyleSheet" href="css/fbp-table.css?v=20261010c" type="text/css" />
<link rel="StyleSheet" href="css/fbp-approval.css?v=20261007b" type="text/css" />
<script type="text/javascript" src="js/fbp-table.js?v=20261009c"></script>
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc page-approval"<% end if %>>
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
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

prcode =request("prcode")
prdate_from =request("prdate_from")
prdate_to =request("prdate_to")
suppliercode =request("suppliercode")
fcname =request("fcname")
deliverydestination =request("deliverydestination")  
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")  

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis") 

If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then 
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, creationdate)
			If Cek_IsDailyClosing = "FALSE" then
				q = " delete from deliveryorder " & _
					" where fcBA = '" & fcba & "'" & _
					" and salescontractno = '" & salescontractno & "'"
					' Response.Write q
					connStr.execute(q)
				
				if err.number = 0 then
				   connStr.CommitTrans
				end if
			else
				err.number = 1001
				err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
			
	        cmdView	= "True"

	elseif Request("txtJenis")= "Delete" then 
		q = " delete from deliveryorder " & _
			" where fcBA = '" & fcba & "'" & _
			" and salescontractno = '"& salescontractno & "'" & _
			" and deliveryorderno = '"& deliveryorderno & "'" 
			' Response.Write q
			connStr.execute(q) 

			cmdView	= "True"

	elseif Request("txtJenis")="New" then	 
		Set rsDeliveryOrder = Server.CreateObject("ADODB.Recordset")
			rsDeliveryOrder.Open " select salescontractno from salesagreementhdr where salescontractno = '" & salescontractno & "' and fcba ='" & fcba & "' ", connstr
		If rsDeliveryOrder.EOF Then 
		
		End If
			rsDeliveryOrder.Close
			set rsDeliveryOrder = Nothing    
				For n=1 to (cdbl(request("no"))-1)
					if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then 
						q =	" Update purchaserequestdetails" & _
								" set orderedqty = 0, approveqty=0 " & _
								"       where prcode = '"& request("prcode_"&TRIM(n)) & "' and fcba = '"& request("fcba_"&TRIM(n)) & "'"   
						' response.write q
						
						q =	" Update purchaserequest" & _
  							" set rowstateid = 'Cancelled'" & _
  							"       where prcode = '"& request("prcode_"&TRIM(n)) & "' and fcba = '"& request("fcba_"&TRIM(n)) & "'" 
			 			' response.write q
						connStr.execute(q)
						
						q =	" Update purchaserequestdetails" & _
  							" set rowstateid = 'Cancelled'" & _
  							"       where prcode = '"& request("prcode_"&TRIM(n)) & "' and fcba = '"& request("fcba_"&TRIM(n)) & "'" 
			 			' response.write q
						connStr.execute(q)
					end if
				next
			cmdView	= "True"  
		Else
			cmdView	= "True"
	End If
Else
	cmdView	= request("cmdView")
End If
 
'untuk menampilkan datepicker default
If Len(Trim(prdate_from)) = 0 Then	prdate_from = textDate(date) 
If Len(Trim(prdate_to)) = 0 Then	prdate_to = textDate(date) 

If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & prdate_from & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & prdate_to & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
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
 
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-ban" aria-hidden="true"></i>Header Cancel Purchase Request</div><% end if %>
<form name="frmEntry" method="post" action="cancel_pr.asp">
	<table class="data" width="60%">  
		<tr>
			<td align="left">
				Date From
			</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="prdate_from" size="10" value="<%=prdate_from%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				Date To
			</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="prdate_to" size="10" value="<%=prdate_to%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				Company
			</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();" class="chosen-select" style="width: 30%">
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
		</tr> 
		<tr>
			<td align="left">
				Business Unit
			</td>
			<td align="left">
				<div id="fcbalist"></div>
				<font color="#FF0000">*</font>
			</td>
				<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
		</tr> 
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				
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
<%if cmdView = "True" AND len(trim(fcba)) > 0 Then
   Dim rsGRN,intRecordsPerPage,cColor 
        q =   "   SELECT fcba," & _
              "          prcode," & _
              "          prdate," & _
              "          purchasegroupid," & _
              "          requestorid," & _
              "          managedid," & _
              "          rowstateid," & _
              "          remarks" & _
              "     FROM (  SELECT a.fcba," & _
              "                    a.prcode," & _
              "                    b.prdate," & _
              "                    b.purchasegroupid," & _
              "                    b.requestorid," & _
              "                    b.managedid," & _
              "                    b.rowstateid," & _
              "                    b.remarks," & _
              "                    SUM (a.orderedqty) AS orderedqty" & _
              "               FROM purchaserequestdetails a" & _
              "                    INNER JOIN purchaserequest b" & _
              "                       ON (a.fcba = b.fcba AND a.prcode = b.prcode)" & _
              "           GROUP BY a.fcba," & _
              "                    a.prcode," & _
              "                    b.prdate," & _
              "                    b.purchasegroupid," & _
              "                    b.requestorid," & _
              "                    b.managedid," & _
              "                    b.rowstateid," & _
              "                    b.remarks)" & _
              "    WHERE     orderedqty = 0" & _
              "          AND fcba = '"& fcba &"'" & _
              "          AND prdate >= TO_DATE ('"& prdate_from &"', 'mm/dd/yyyy')" & _
              "          AND prdate <= TO_DATE ('"& prdate_to &"', 'mm/dd/yyyy')" & _
              "          AND rowstateid in ('Approved','AuthorizeOnProgress')" & _
              " ORDER BY fcba, prcode, prdate" 
 
   ' response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.Open q, connstr%>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar PR<small>PR berstatus Approved / AuthorizeOnProgress yang belum di-order</small></div><% end if %>
<form action="cancel_pr.asp" method="post" name="eForm" >
   <div class="pra-card">
   <table class="data pra-table" id="praTable" style="width:100%;">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
  	<TD align="center">No</TD> 
  	<TD align="center">PR Code</TD>
  	<TD align="center">PR Date</TD>
    <TD align="center">Purchase Group</TD>
  	<TD align="center">Requestor</TD>
  	<TD align="center">Managed</TD>
    <TD align="center">Remarks</TD>
	<TD align="center" data-nosort title="Pilih semua PR yang tampil di halaman ini"><input type='checkbox' name='select-all' id='select-all' aria-label="Pilih semua PR yang tampil"/></TD>	
	</tr>
	</thead>
	<tbody> 
	
   <%
   if Not rsGRN.eof then
      nCountofField=rsGRN.fields.count-1
	  no=1
	  While Not rsGRN.EOF
		if no  mod 2 = 1 then
			cColor = "#EEEEEE"
		else
			cColor = "#E5E5E5"
		end if%>
		<tr bgcolor="<%=cColor%>">
			<td align="center" valign="top"><%=NO%></td> 
			<td align="center" title="<%=rsGRN("prcode").Name%>" valign="top"><%=rsGRN("prcode")%></td>
			<td align="center" title="<%=rsGRN("prdate").Name%>" valign="top"><%=rsGRN("prdate")%></td>
			<td align="center" title="<%=rsGRN("purchasegroupid").Name%>" valign="top"><%=rsGRN("purchasegroupid")%></td>
			<td align="center" title="<%=rsGRN("requestorid").Name%>" valign="top"><%=rsGRN("requestorid")%></td>
			<td align="center" title="<%=rsGRN("managedid").Name%>" valign="top"><%=rsGRN("managedid")%></td>
      <td align="center" title="<%=rsGRN("remarks").Name%>" valign="top"><%=rsGRN("remarks")%></td>
			<td valign="middle" align="center">
				<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
				<input type="hidden" name="fcba_<%=NO%>" value="<%=rsGRN("fcba")%>">
				<input type="hidden" name="prcode_<%=NO%>" value="<%=rsGRN("prcode")%>"> 
			</td>
		</tr>
		  <%
		  rsGRN.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write"<tr><td colspan='8' align='left' class='pra-empty'>Data is empty</td></tr>"
   end if 
   rsGRN.Close : set rsGRN = Nothing
%>
	</tbody>
   </table>
   </div>
   <div class="pra-actions">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Hidden" Value="<%=no%>" Name="no">
				<input Type="Hidden" Value="<%=prdate_from%>" Name="prdate_from">
				<input Type="Hidden" Value="<%=prdate_to%>" Name="prdate_to">
				<input type="hidden" value="<%=selectedfcba%>" name="selectedfcba">
				<input type="hidden" value="<%=fccompanycode%>" name="fccompanycode">
					<span class="pra-selected" id="praSelected" aria-live="polite"></span>
					<button type="submit" class="button" name="txtJenis" value="New" onClick="return confirm('Anda yakin untuk melakukan pembatalan terhadap PR yang sudah dipilih?')">
						<i class="fa fa-check"></i> Save
					</button>
   </div>
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
<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
}

function getselectedpocode() {
	var spocode = document.getElementById("selectedpocode");
	spocode.value = document.getElementById("pocode").value;
}

function getselectedinvoice() {
	var sfcba = document.getElementById("selectedinvoice");
	sfcba.value = document.getElementById("invoice_by").value;
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
	getlistpocode();
	getselectedfcba();
	getselectedinvoice()
	getpoinfo();
});
</script>
<script type="text/javascript">
	function handleClick(cb){
		if(cb.checked == 1){
			cb.value = 1;
		} else {
			cb.value = 0;
		}
		if (window.praSync) praSync();
	}
	

</script>
<script type="text/javascript">	
		// "Pilih semua" hanya untuk baris yang sedang tampil (hasil cari / halaman aktif),
		// supaya PR yang tersembunyi tidak ikut terproses tanpa terlihat.
		$('#select-all').click( function () {
			var rows = $('#praTable tbody tr:visible .item');
			rows.prop('checked', this.checked);
			rows.val(this.checked ? 1 : 0);
			if (window.praSync) praSync();
		})
	</script>
<script type="text/javascript">
	// Tampilan: jumlah PR terpilih, tombol aktif bila ada pilihan, tabel cari/urut/halaman
	(function () {
		var table = document.getElementById('praTable');
		var info = document.getElementById('praSelected');
		var buttons = document.querySelectorAll('.pra-actions button[type="submit"]');
		if (!table) return;
		window.praSync = function () {
			var items = table.querySelectorAll('tbody .item');
			var n = 0, visible = 0, visibleChecked = 0;
			for (var i = 0; i < items.length; i++) {
				var shown = items[i].offsetParent !== null;
				if (items[i].checked) n++;
				if (shown) { visible++; if (items[i].checked) visibleChecked++; }
				var tr = items[i].closest ? items[i].closest('tr') : null;
				if (tr) tr.classList.toggle('pra-checked', items[i].checked);
			}
			if (info) info.innerHTML = items.length ? ('<b>' + n + '</b> dari ' + items.length + ' PR dipilih') : 'Tidak ada PR yang dapat diproses';
			for (var b = 0; b < buttons.length; b++) buttons[b].disabled = (n === 0);
			var all = document.getElementById('select-all');
			if (all) { all.checked = visible > 0 && visibleChecked === visible; all.indeterminate = visibleChecked > 0 && visibleChecked < visible; }
		};
		if (window.FbpTable) FbpTable.init(table, { pageSize: 10, onChange: function () { praSync(); } });
		praSync();
	})();
</script>
</body>
</html>
