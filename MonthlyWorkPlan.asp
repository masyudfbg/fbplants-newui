<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - Monthly Work Plan</title><link rel="StyleSheet" href="dtree.css" type="text/css" />
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
	
	function hitungTotalAmount(evt) {
		var ratemandays = parseFloat(document.getElementById("ratemandays").value);
		var mandays = parseFloat(document.getElementById("mandays").value);
		var total = mandays * ratemandays;
		
		document.getElementById("totalratemandays").value = total.toFixed(2);
		//alert (total);
	}
	
	function hitungMandays(evt) {
		var normamandays = parseFloat(document.getElementById("normamandays").value);
		var volume = parseFloat(document.getElementById("volume").value);
		var rotasi = parseFloat(document.getElementById("rotasi").value);
		
		var total = rotasi * (volume * normamandays);
		//alert (total);
		document.getElementById("mandays").value = total;
		
	}
	
	function hitungTotalPremi(evt) {
		var volume = parseFloat(document.getElementById("volume").value);
		var total = parseFloat(evt.value) * volume;
		//alert (total);
		if (!isNaN(total)) {
    	 	$("#boronganpremi").val(total);
    	}
    	else{
    		$("#boronganpremi").val();
    	}
		
		//document.getElementById("boronganpremi").value = total.toFixed(2);
		
	}
	
	function hitungTotalPremi2(evt) {
		var volume = parseFloat(document.getElementById("volume").value);
		var normapremi = parseFloat(document.getElementById("normapremi").value);
		var total = normapremi * volume;
		
		if (!isNaN(total)) {
    	 	$("#boronganpremi").val(total);
    	}
    	else{
    		$("#boronganpremi").val();
    	}
		
	//	document.getElementById("boronganpremi").value = total.toFixed(2);
		
	}
	
	function hitungTotalQuantity(evt) {
		var materialnorm = parseFloat(document.getElementById("materialnorm").value);
		var volume = parseFloat(document.getElementById("volume").value);
		var total = volume * materialnorm;
		var new_total = total.toFixed(2);
	
		document.getElementById("quantity").value = new_total;
		//alert (total);
	}
	
	function hitungTotalAmountMatByQTy(evt) {
		var ratematerial = parseFloat(document.getElementById("ratematerial").value);
		var total = parseFloat(document.getElementById("quantity").value) * ratematerial;
		var new_total = total.toFixed(2);
		
		document.getElementById("totalratematerial").value = new_total;
		//alert (total);
	}
	
	function hitungTotalAmountMatByRate(rate) {
		var qty = parseFloat(document.getElementById("volume").value);
		var total = parseFloat(rate.value) * qty;
		var new_total = total.toFixed(2);
		
		document.getElementById("totalratematerial").value = new_total;
		//alert (total);
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

var getyear = function () {
	// build up the post string when passing variables to the server side page
	
	//alert(PostStr);
	var year = document.getElementById("selectedyear").value;
	var PostStr = 'year='+year;
	// use the generic function to make the request
	doAJAXCall('ajax/getyear.asp', 'POST', PostStr, showyear);
}

// The function for handling the response from the server
var showyear = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("yearlist").innerHTML = "";
	document.getElementById("yearlist").innerHTML = response;

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

var getmonth = function () {
	// build up the post string when passing variables to the server side page
	
	//alert(PostStr);
	var month = document.getElementById("selectedmonth").value;
	var PostStr = 'month='+month;
	// use the generic function to make the request
	doAJAXCall('ajax/getmonth.asp', 'POST', PostStr, showmonth);
}

// The function for handling the response from the server
var showmonth = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("monthlist").innerHTML = "";
	document.getElementById("monthlist").innerHTML = response;

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

var getactivity = function () {
	// build up the post string when passing variables to the server side page
	
	//alert(PostStr);
	var fcba = document.getElementById("fcba").value;
	var activity = document.getElementById("selectedactivity").value;
	var PostStr = 'fcba='+fcba+'&activity='+activity;
	// use the generic function to make the request
	doAJAXCall('ajax/getactivitynorm.asp', 'POST', PostStr, showactivity);
}

// The function for handling the response from the server
var showactivity = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("activitylist").innerHTML = "";
	document.getElementById("activitylist").innerHTML = response;

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

var getlistafdeling = function () {
	// build up the post string when passing variables to the server side page
	
	var fcba = document.getElementById("fcba").value;
	var selectedafdeling = document.getElementById("selectedafdeling").value;
	
	var PostStr = 'fcba='+fcba+'&afdeling='+selectedafdeling;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistdivisionrkb.asp', 'POST', PostStr, showlistafdeling);
}

// The function for handling the response from the server
var showlistafdeling = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("afdelinglist").innerHTML = "";
	document.getElementById("afdelinglist").innerHTML = response;

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

var gettargetcode = function () {
	// build up the post string when passing variables to the server side page
	
	var fcba = document.getElementById("fcba").value;
	var fieldcode = document.getElementById("selectedfield").value;
	var activity = document.getElementById("selectedactivity").value;
	var PostStr = 'fcba='+fcba+'&fieldcode='+fieldcode+'&activity='+activity;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/gettargetcode.asp', 'POST', PostStr, showlistfield);
}

// The function for handling the response from the server
var showlistfield = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("fieldlist").innerHTML = "";
	document.getElementById("fieldlist").innerHTML = response;

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

var getHKRate = function () {

	var fcba = document.getElementById("fcba").value;
	var PostStr = 'fcba='+fcba;
	
	doAJAXCall('ajax/GetHKRate.asp', 'POST', PostStr, ShowHKRateResponse);
	
}

var ShowHKRateResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.lineno);
    
    // update the Div to show the result from the server
	document.getElementById("ratemandays").value = s.ratemandays;
}



var getNormaHK = function () {

	var fcba = document.getElementById("fcba").value;
	var activitycode = document.getElementById("activitycode").value;
	
	var PostStr = 'fcba='+fcba+'&activitycode='+activitycode;
	
	doAJAXCall('ajax/GetNormaHK.asp', 'POST', PostStr, ShowNormaHKResponse);
	
}

var ShowNormaHKResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.lineno);
    
    // update the Div to show the result from the server
	document.getElementById("normamandays").value = s.normamandays;
}

var getMaterialNorm = function () {

	var fcba = document.getElementById("fcba").value;
	var stock = document.getElementById("stock").value;
	var activitycode = document.getElementById("activitycode").value;
	
	var PostStr = 'fcba='+fcba+'&stock='+stock+'&activitycode='+activitycode;
	
	doAJAXCall('ajax/GetMaterialNorm.asp', 'POST', PostStr, ShowMaterialNormResponse);
	
}

var ShowMaterialNormResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
   // alert (s.materialnorm);
    
    // update the Div to show the result from the server
	document.getElementById("materialnorm").value = s.materialnorm;
}

var getratematerial = function () {

	var fcba = document.getElementById("fcba").value;
	var stock = document.getElementById("stock").value;
	var txtYear = document.getElementById("txtYear").value;
	var txtMonth = document.getElementById("txtMonth").value;
	
	var PostStr = 'fcba='+fcba+'&stock='+stock+'&txtYear='+txtYear+'&txtMonth='+txtMonth;
	// alert(PostStr);
	doAJAXCall('ajax/getStockLedgerRKB.asp', 'POST', PostStr, ShowratematerialResponse);
	
}

var ShowratematerialResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.amtperunit);
    
    // update the Div to show the result from the server
	document.getElementById("ratematerial").value = s.amtperunit;
}

var getNoRKB = function () {
	// build up the post string when passing variables to the server side page

	var fcba = document.getElementById("fcba").value;
	var txtYear = document.getElementById("txtYear").value;
	var txtMonth = document.getElementById("txtMonth").value;
	var activitycode = document.getElementById("activitycode").value;
	var divisi = document.getElementById("divisi").value;
	var fieldcode = document.getElementById("fieldcode").value;
	var PostStr = 'fcba='+fcba +'&txtYear='+ txtYear +'&txtMonth='+ txtMonth +'&activitycode='+ activitycode +'&divisi='+ divisi +'&fieldcode='+ fieldcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetNoRKB.asp', 'POST', PostStr, ShowNoRKBResponse);
}

// The function for handling the response from the server
var ShowNoRKBResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.lineno);
    
    // update the Div to show the result from the server
	document.getElementById("transno").value = s.transno;
}

function getselectedyear() {
	var selectedyear = document.getElementById("selectedyear");
	selectedyear.value = document.getElementById("txtYear").value;
}

function getselectedmonth() {
	var selectedmonth = document.getElementById("selectedmonth");
	selectedmonth.value = document.getElementById("txtMonth").value;
}

function getselectedactivity() {
	var selectedactivity = document.getElementById("selectedactivity");
	selectedactivity.value = document.getElementById("activitycode").value;
}

function getselectedafdeling() {
	var selectedafdeling = document.getElementById("selectedafdeling");
	selectedafdeling.value = document.getElementById("divisi").value;
}

function getselectedfield() {
	var selectedfield = document.getElementById("selectedfield");
	selectedfield.value = document.getElementById("fieldcode").value;
}

function clearheader() {
	transno.value = null;
	selectedyear.value = null;
	selectedmonth.value = null;
	selectedactivity.value = null;
	selectedafdeling.value = null;
	selectedfield.value = null;
}

var getFieldInfo = function () {

	var fcba = document.getElementById("fcba").value;
	var txtYear = document.getElementById("txtYear").value;
	var txtMonth = document.getElementById("txtMonth").value;
	var fieldcode = document.getElementById("fieldcode").value;
	var PostStr = 'fcba='+fcba+'&fieldcode='+fieldcode+'&txtMonth='+txtMonth+'&txtYear='+txtYear;
	//alert(PostStr);
	doAJAXCall('ajax/GetFieldRKB.asp', 'POST', PostStr, ShowFieldInfoResponse);
	
}

var ShowFieldInfoResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.lineno);
    
    // update the Div to show the result from the server
	document.getElementById("haactual").value = s.hectarageplanted; 
}

</script>
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
 Dim cmdView
 
 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 selectedyear =request("selectedyear")
 selectedmonth =request("selectedmonth")
 selectedactivity =request("selectedactivity")
 selectedafdeling =request("selectedafdeling")
 selectedfield =request("selectedfield")

 fcba = request("fcbas")
 stock =request("stock")
 transno =request("transno")
 ratemandays =request("ratemandays")
 normamandays =request("normamandays")
 normapremi =request("normapremi")
 totalratemandays =request("totalratemandays")
 mandays =request("mandays")  
 volume =request("volume")
 boronganpremi =request("boronganpremi")
 divisi = request("divisi")
 fieldcode = request("fieldcode")
 months = request("txtMonth")
 years = request("txtYear")
 activity = request("activitycode")
 quantity =request("quantity")
 materialnorm =request("materialnorm")
 haactual = request("haactual")
 haworkplan =request("haworkplan")
 ratematerial =request("ratematerial")
 totalratematerial =request("totalratematerial")
 remarks =request("remarks")
 lineno =request("lineno")
 janjang =request("janjang")
 kilogram =request("kilogram")
 rotasi =request("rotasi")
 'response.write Request("txtJenis") & cmdView
 
 '---- Control Job
	'qRateMandays = " select rate_1 from allowded where fcba ='" & fcba & "' and fccode ='A01'"
	'response.write qRateMandays
	'Set rsRateMandays = Server.CreateObject("ADODB.Recordset")
	'rsRateMandays.Open qRateMandays, connstr
	'	if not rsRateMandays.eof then
	'	   ratemandays=rsRateMandays("rate_1")
	'	else
	'	   ratemandays=""
	'	end if
'---- End Control Job


 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis") = "" then
	 '---- Control Job
	qHeader = "select monthlyworkplan.fcba,period_year,period_month,activitycode,rate,fieldcode,nvl(mandays,0) mandays,nvl(volume,0) volume,nvl(boronganpremi,0) boronganpremi,rkbno, totalamount,field.division,nvl(field.hectarageplanted,0) hectarageplanted,nvl(normamandays,0) normamandays, " & _ 
	          " nvl(normapremi,0) normapremi ,nvl(mandays,0) / (nvl(normamandays,0) * nvl(volume,0) ) rotasi" & _
			  " from monthlyworkplan " & _
			  " inner join field on monthlyworkplan.fieldcode = field.fccode and field.fcba = monthlyworkplan.fcba " & _
			  " where rkbno ='" & transno & "'"
	'response.write qHeader
	Set rsHeader = Server.CreateObject("ADODB.Recordset")
	rsHeader.Open qHeader, connstr
		if not rsHeader.eof then
			fcba = rsHeader("fcba")
			selectedyear = rsHeader("period_year")
			selectedmonth = rsHeader("period_month")
			selectedactivity = rsHeader("activitycode")
			rate = rsHeader("rate")
			selectedafdeling = rsHeader("division")
			selectedfield = rsHeader("fieldcode")
            normamandays = 	rsHeader("normamandays")		
			volume = rsHeader("volume")
			rotasi = rsHeader("rotasi")
			mandays = rsHeader("mandays")
			normapremi = rsHeader("normapremi")
			boronganpremi = rsHeader("boronganpremi")
			transno = rsHeader("rkbno")
			totalratemandays = rsHeader("totalamount")
			haactual = rsHeader("hectarageplanted")
			
		end if
	'---- End Control Job
	
   rsHeader.Close
   set rsHeader = Nothing
	
	if Cdbl(volume) > Cdbl(haactual) then
		Err.Description = "ha plan (" & volume & ") cannot be greater than ha actual (" & haactual & ") !"
		response.Write  "<div class = ""error"">" & err.description & "</div>"
    end if			

 end if
 
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
 
	'response.write Request("txtJenis")
	if Request("txtJenis")= "Header" then
	   connstr.BeginTrans
		If err.number=0 then
			q="Delete from MONTHLYWORKPLAN where  fcba='" & fcba & "' and RKBNO='" & transno & "'"
			
			'Response.Write q
			
			connStr.execute(q)
			if err.number  = 0 then
				connStr.CommitTrans
			end if
		end if
       cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	
	    connstr.BeginTrans		
									
		q = "Delete from MONTHLYWORKPLANDET where  fcba='" & fcba & "' and RKBNO='" & transno & "' and lineno = '"& lineno &"'"
		connStr.execute(q)
		
		'Response.Write q
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans	 
		'if mandays = "0" or mandays = "" then
			'Err.Number = 1001
			'Err.Description = "mandays cannot be empty and must be greater than zero !"
		'Else
		
			Set rsRKB = Server.CreateObject("ADODB.Recordset")
			rsRKB.Open "Select RKBNO from MONTHLYWORKPLAN where RKBNO='" & transno & "' AND FCBA = '" & fcba & "'" , connstr
			IF rsRKB.eof then
				q = " INSERT into MONTHLYWORKPLAN" & _
					" (RKBNO, UNITKERJA, PERIOD_YEAR, PERIOD_MONTH, FIELDCODE, ACTIVITYCODE, MANDAYS, VOLUME, BORONGANPREMI, RATE, NORMAMANDAYS, NORMAPREMI, TOTALAMOUNT, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
					" '" & transno & "'" & _
					" ,'" & divisi & "'" & _
					" , " & years & "" & _
					" , " & months & "" & _
					" , '" & fieldcode & "'" & _
					" , '" & activity & "'" & _
					" , '" & mandays & "'" & _
					" , '" & volume & "'" & _
					" , '" & boronganpremi & "'" & _
					" , '" & ratemandays & "'" & _
					" , '" & normamandays & "'" & _
					" , '" & normapremi & "'" & _
					" , '" & totalratemandays & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					" , '" & fcba & "'" & _
					" , sysdate" & _
					" , To_Char(sysdate,'HH24:MI')" & _
					" )"								
				'Response.Write q
				connStr.execute(q)
			End If
			rsRKB.Close
			
		'End if
		
		'if mandays = "0" or mandays = "" then
		'	Err.Number = 1001
		'	Err.Description = "mandays cannot be empty and must be greater than zero !"
		'Else
			'response.write haactual & " "& haworkplan
			if quantity = "" then quantity = 0 end if
			if materialnorm = "" then materialnorm = 0 end if
			if haworkplan = "" then haworkplan = 0 end if
			if ratematerial = "" then ratematerial = 0 end if
			if totalratematerial = "" then totalratematerial = 0 end if
			if kilogram = "" then kilogram = 0 end if
			if janjang = "" then janjang = 0 end if
			if rotasi = "" then rotasi = 0 end if
				
			if Cdbl(haworkplan) > Cdbl(haactual) then
				Err.Number = 1001
				Err.Description = "ha plan cannot be greater than ha actual !"
			Else			
				
				
				Set rsRKBDet = Server.CreateObject("ADODB.Recordset")
					rsRKBDet.Open "Select RKBNO from MONTHLYWORKPLANDET where RKBNO='" & transno & "' AND STOCKCODE = '"& stock &"' AND FCBA = '" & fcba & "'" , connstr
				IF rsRKBDet.eof then
					
					'get line no				
					qLineno = "Select nvl(max(lineno),0) lineno from MONTHLYWORKPLANDET where RKBNO='" & transno & "' AND FCBA = '" & fcba & "'"
										
					set rsRKBDetline= Server.CreateObject("ADODB.Recordset")
					rsRKBDetline.Open qLineno, connstr

					if not rsRKBDetline.eof then
						lineno = cdbl(rsRKBDetline("lineno")) + 1
					end if 
					rsRKBDetline.Close
					
					qdetail = 	" INSERT into MONTHLYWORKPLANDET" & _
								" (LINENO, RKBNO, STOCKCODE, QTY, ACTIVEMATERIALNORM, HAPLAN, RATESTOCK, TOTALAMOUNT, KG, JJG, ROTASI, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) values (" & _
								" '" & lineno & "'" & _
								" , '" & transno & "'" & _
								" , '" & stock & "'" & _
								" , " & quantity & "" & _
								" , " & materialnorm & "" & _
								" , " & haworkplan & "" & _
								" , " & ratematerial & "" & _
								" , " & totalratematerial & "" & _
								" , " & kilogram & "" & _
								" , " & janjang & "" & _
								" , " & rotasi & "" & _								
								" , '" & remarks & "'" & _
								" , '" & Session("CodeName") & "'" & _
								" , '" & Session("CodeName") & "'" & _
								" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
								" , '" & fcba & "'" & _
								" , sysdate" & _
								" , To_Char(sysdate,'HH24:MI')" & _
							   " )"

					'Response.Write qdetail
					connStr.execute(qdetail)
				Else
					Err.Number = 1001
					Err.Description = "Material '"& stock &"' for field '"& fieldcode &"'  already exist!"
				End If
				if err.number  = 0 then
					connStr.CommitTrans
				end if
				
				rsRKBDet.Close
			End If			
		'End If
		
        cmdView	= "True"
		
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	    connstr.BeginTrans
		if quantity = "" then quantity = 0 end if
		if materialnorm = "" then materialnorm = 0 end if
		if haworkplan = "" then haworkplan = 0 end if
		if ratematerial = "" then ratematerial = 0 end if
		if totalratematerial = "" then totalratematerial = 0 end if
		if kilogram = "" then kilogram = 0 end if
		if janjang = "" then janjang = 0 end if
		if rotasi = "" then rotasi = 0 end if
		
		'if mandays = "0" or mandays = "" then
		'	Err.Number = 1001
		'	Err.Description = "mandays cannot be empty and must be greater than zero !"
		'Else
			if Cdbl(haworkplan) > Cdbl(haactual) then
				Err.Number = 1001
				Err.Description = "ha plan cannot be greater than ha actual !"
			Else
				q = " Update MonthlyWorkPlanDet set" & _
					"     stockcode = '" & stock & "'"& _
					"     , qty = " & quantity & "" & _
					"     , activematerialnorm = " & materialnorm & "" & _					
					"     , kg = " & kilogram & "" & _
					"     , jjg = " & janjang & "" & _
					"     , rotasi = " & rotasi & "" & _
					"     , haplan = " & haworkplan & "" & _
					"     , ratestock = " & ratematerial & "" & _
					"     , totalamount = " & totalratematerial & "" & _
					"     , REMARKS = '" & Remarks & "'" & _
					"     , LASTUPDATE = sysdate " & _
					"     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
					" where  fcba='" & fcba & "' and RKBNO='" & transno & "' and lineno = "& lineno &""
					
					'response.write q
					connStr.execute(q)
					if err.number  = 0 then
						connStr.CommitTrans
						stock = ""
						quantity = ""
						materialnorm = ""
						kilogram = ""
						janjang = ""
						rotasi = ""
						haworkplan = ""
						ratematerial = ""
						totalratematerial = ""
						Remarks = ""
					end if
			End If
		'End If
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if
 
 'get line no
  qhaactual = " SELECT field_api.getvalid_ha(fccode, fcba, ADD_MONTHS((LAST_DAY(SYSDATE)+1),-1)) rshaactual" & _
			  " FROM field" & _
			  " WHERE fccode = '"& fieldcode &"' AND fcba = '" & fcba & "'"
  
  'response.write qhaactual
  set rshaactual= Server.CreateObject("ADODB.Recordset")
  rshaactual.Open qhaactual, connstr

  if not rshaactual.eof then
	haactual = cdbl(rshaactual("rshaactual"))
  end if
  rshaactual.Close
 
 
 if err.number <> 0 then
   connStr.RollbackTrans   
   response.Write  "<div class = ""error"">" & err.description & "</div>"
 End if
 '-------------ISI Payment Form

 %>
 <form name="frmEntry" method="post" action="MonthlyWorkPlan.Asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Business Unit</td>
			<td align="left" width="35%">
				<select id="fcba" name="fcbas" onchange="javascript:clearheader(); javascript:getactivity(); javascript:getlistafdeling(); javascript:getHKRate(); javascript:getFieldInfo(); javascript:getNoRKB(); javascript:getratematerial();" class="chosen-select" style="width: 60%">
					<option value=""></option>
					<%
					qlov = " SELECT fcba, fcname FROM user_fcba_tab uft INNER JOIN businessunit bu " & _
						   " ON bu.fccode = uft.fcba WHERE uft.fcusercode = '" & Session("CodeName") & "' and bu.fctype='E'"   
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fcba")%>" <%if fcba = rslov("fcba") then%>selected<%end if%>><%=rslov("fcba")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<td align="left">Unit Kerja</td>
			<td align="left" width="35%">
				<div id ="afdelinglist"></div>
				<input type="hidden" id= "selectedafdeling" name="selectedafdeling" size="20" value="<%=selectedafdeling%>" >
			</td>
		</tr>
		<tr>
			<td align="left" width="15%" >Year</td>
			<td align="left">
				<div id ="yearlist"></div>
				<input type="hidden" id= "selectedyear" name="selectedyear" size="20" value="<%=selectedyear%>" >
			</td>
			<td align="left" width="15%">Field</td> 
			<td align="left">
				<div id ="fieldlist"></div>
				<input type="hidden" id= "selectedfield" name="selectedfield" size="20" value="<%=selectedfield%>" >
			</td>
		</tr>
		<tr>
			<td align="left" >Month</td>
			<td align="left">
				<div id ="monthlist"></div>
				<input type="hidden" id= "selectedmonth" name="selectedmonth" size="20" value="<%=selectedmonth%>" >
			</td>
			<td align="left" style="width:20%;">Ha Actual</td>
			<td align="left">
				<input type="text" id="haactual" name="haactual" value="<%=haactual%>" style ="background-color: #eee;">
			</td>	
		</tr>
		<tr>
			<td align="left" >Activity</td>
			<td align="left">
				<div id ="activitylist"></div>
				<input type="hidden" id= "selectedactivity" name="selectedactivity" size="20" value="<%=selectedactivity%>" >
			</td>
			<td align="left">Mandays</td>
			<td align="left">
				<input type="text" maxlength="15" id="mandays" name="mandays" size="10" value="<%=mandays%>" autocomplete="Off" readonly style ="background-color: #eee;">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left" >Rate</td>
			<td align="left">
				<input type="text" maxlength="50" id="ratemandays" name="ratemandays" size="20" value="<%=ratemandays%>" autocomplete="Off" readonly style ="background-color: #eee;">
			</td>
			<%If len(trim(transno)) > 0 Then %>
				<td align="left" >RKB No 
				<td align="left">
					<input type="text" maxlength="50" id="transno" name="transno" size="30" value="<%=transno%>" autocomplete="Off" readonly style ="background-color: #eee;">
					<font color="#FF0000">*</font>
				</td>
			<%Else%>
				<td align="left" >RKB No 
				<td align="left">
					<input type="text" maxlength="50" id="transno" name="transno" size="30" value="<%=transno%>" autocomplete="Off" >
					<font color="#FF0000">*</font>
				</td>
			<%End If%>		
		</tr>
		<tr>
			<td align="left" >Mandays Norm</td>
			<td align="left">
				<input type="text" maxlength="50" id="normamandays" name="normamandays" size="20" value="<%=normamandays%>" autocomplete="Off" readonly style ="background-color: #eee;">
			</td>
			<td align="left" >Total Amount Mandays</td>
			<td align="left">
				<input type="text" maxlength="50" id="totalratemandays" name="totalratemandays" size="20" value="<%=totalratemandays%>" autocomplete="Off" readonly style ="background-color: #eee;">
			</td>			
		</tr>
		<tr>
			<td align="left" >Luas Dikerjakan</td>
			<td align="left">
				<input type="text" maxlength="50" id="volume" name="volume" size="20" value="<%=volume%>" autocomplete="Off" onkeyup="hanyaAngka(this); hitungMandays(this); hitungTotalAmount(this); hitungTotalPremi2(this); hitungTotalQuantity(this);">&nbsp;Ha
				<font color="#FF0000">*</font>
			</td>
			<td align="left" >Borongan Rate</td>
			<td align="left">
				<input type="text" maxlength="50" id="normapremi" name="normapremi" size="20" value="<%=normapremi%>" onkeyup="hanyaAngka(this); hitungTotalPremi(this);" autocomplete="Off" >
			</td>	
		</tr>
		<tr>
			<td align="left">Rotasi</td>
			<td align="left"><input type="text" id="rotasi" name="rotasi" value="<%=rotasi%>" onkeyup="hitungMandays(this); hitungTotalAmount(this);hitungTotalPremi2(this); ">
			<font color="#FF0000">*</font>
			</td>				
			<td align="left" >Amount Borongan</td>
			<td align="left">
				<input type="text" maxlength="50" id="boronganpremi" name="boronganpremi" size="20" value="<%=boronganpremi%>" autocomplete="Off" onkeyup="hanyaAngka(this);">
			</td>
						
		</tr>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="MonthlyWorkPlan.Asp?transno=<%=transno%>&fcbas=<%=fcba%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>
 
<%

'response.write ratemandays

if cmdView = "True" AND LEN(TRIM(transno)) > 4 and len(trim(fcba)) > 0 and cdbl(to_nol(volume)) <= cdbl(to_nol(haactual)) Then
   Dim rsDet,intRecordsPerPage,cColor
   set rsDet = server.CreateObject("Adodb.recordset")
   q	=	" SELECT mwd.lineno, mwd.fcba, mwd.RKBNO, mwd.STOCKCODE, REPLACE(mwd.QTY,',','.') Qty, REPLACE(mwd.activematerialnorm,',','.') activematerialnorm, REPLACE(mwh.volume ,',','.') volume, REPLACE(mwd.RATESTOCK,',','.') RATESTOCK, REPLACE(mwd.TOTALAMOUNT,',','.') TOTALAMOUNT, mwd.KG, mwd.JJG, mwd.ROTASI, mwd.REMARKS, stock.fccode, stock.fcname from MONTHLYWORKPLANDET mwd"&_
			" inner join monthlyworkplan mwh on mwd.fcba = mwh.fcba and mwd.rkbno = mwh.rkbno"&_
			" left join activematerial stock on  mwd.stockcode = stock.fccode"&_
			" where mwd.fcba ='"& fcba &"' and mwd.RKBNO = '"& transno &"'"
   'response.Write q	
   Set rsDet = Server.CreateObject("ADODB.Recordset")
   rsDet.CursorType = 3
   rsDet.CursorLocation = 3
   rsDet.Open q, connstr
   %>
   <form action="MonthlyWorkPlan.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td align="left">Material</td>
					<td align="left" width="35%">
						<select id="stock" name="stock" class="chosen-select" style="width: 80%" onchange="javascript:getMaterialNorm(); javascript:getratematerial();">
							<option value=""></option>
							<%
							'qlov = "select  itemcode as fccode, stock_api.Get_stockName(itemcode, fcba)fcname from stockledger where fcba ='"& fcba &"' and month = '"& months &"' and year = '"& years &"' and stock_api.get_StockGroup(itemcode, fcba) IN ('301','302','305','307') ORDER BY itemcode"   
							qlov =  " SELECT actmat,fcname,uomactivematerial FROM  materialactivity_norm a INNER JOIN activematerial b ON a.actmat=b.fccode " & _
									" WHERE a.fcba='"& fcba &"' and job='" & selectedactivity & "'" & _
									" ORDER BY fcname"
									
							 
							Set rslov = Server.CreateObject("ADODB.Recordset")
							rslov.Open qlov, connstr 
							if Not rslov.eof then
								While Not rslov.EOF
							%>
							<option value="<%=rslov("actmat")%>"<%if stock = rslov("actmat") then%>selected<%end if%>><%=rslov("actmat")%> : <%=rslov("fcname")%></option>
							<%
									rslov.MoveNext
								Wend
							end if
							%>
						</select>
					</td>
					<td align="left">Rate Material</td>
					<td align="left"><input type="text" id="ratematerial" name="ratematerial" value="<%=ratematerial%>" onkeyup="hanyaAngka(this); hitungTotalQuantity(this);hitungTotalAmountMatByRate(this); hitungTotalAmountMatByQTy(this);"></td>
				</tr>
				<% if activity = "6502101" then %>
				<tr>					
						<td align="left">Janjang</td>
						<td align="left" width="35%"><input type="text" id="janjang" name="janjang" value="<%=janjang%>" ></td>
						<td align="left" style="width:15%;">Kilogram</td>
						<td align="left" width="35%"><input type="text" id="kilogram" name="kilogram" value="<%=kilogram%>" ></td>
				</tr>
				<% End if %>				
				<tr> 					
					<td align="left">Material Norm</td>
					<td align="left">
						<input type="text" maxlength="15" id="materialnorm" name="materialnorm" size="10" onkeyup="hanyaAngka(this); hitungTotalQuantity(this);" value="<%=materialnorm%>" autocomplete="Off" readonly style ="background-color: #eee;">
					</td>
					<td align="left">Total Amount</td>
					<td align="left"><input type="text" id="totalratematerial" name="totalratematerial" value="<%=totalratematerial%>" style ="background-color: #eee;"></td>		
				</tr> 
				<tr>
					<td align="left">Quantity</td>
					<td align="left">
						<input type="text" maxlength="15" id="quantity" name="quantity" size="10" onkeyup="hitungTotalAmountMatByQTy(this);" value="<%=quantity%>" readonly style ="background-color: #eee;" autocomplete="Off">
					</td>	 
					<td align="left">Remarks</td>
					<td align="left" ><input type="text" id="remarks" size="50" name="remarks" value="<%=remarks%>" ></td>
				</tr>
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
		  <input type="hidden" name="lineno" value="<%=lineno%>">
	      <input type="hidden" name="transno" value="<%=transno%>">
		  <input type="hidden" name="mandays" value="<%=mandays%>">
		  <input type="hidden" name="volume" value="<%=volume%>">
		  <input type="hidden" name="rotasi" value="<%=rotasi%>">
		  <input type="hidden" name="boronganpremi" value="<%=boronganpremi%>">
		  <input type="hidden" name="haactual" value="<%=haactual%>">
		  <input type="hidden" name="ratemandays" value="<%=ratemandays%>">
		  <input type="hidden" name="normamandays" value="<%=normamandays%>">
		  <input type="hidden" name="normapremi" value="<%=normapremi%>">
		  <input type="hidden" name="totalratemandays" value="<%=totalratemandays%>">
		  <input type="hidden" name="fcbas" value="<%=fcba%>">
		  <input type="hidden" name="txtYear" value="<%=years%>">
		  <input type="hidden" name="txtMonth" value="<%=months%>">
		  <input type="hidden" name="activitycode" value="<%=activity%>">		 
		  <input type="hidden" name="divisi" value="<%=divisi%>">
		  <input type="hidden" name="fieldcode" value="<%=fieldcode%>">
		  <input type="hidden" name="selectedyear" value="<%=years%>">
		  <input type="hidden" name="selectedmonth" value="<%=months%>">
		  <input type="hidden" name="selectedactivity" value="<%=activity%>">		 
		  <input type="hidden" name="selectedafdeling" value="<%=divisi%>">
		  <input type="hidden" name="selectedfield" value="<%=fieldcode%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset" name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="2" width="30">&nbsp;</td>
		<% if activity = "6502101" then %>
			<td align="left">Janjang</td>
			<td align="left">Kilogram</td>
			<td align="left">Rotasi</td>
		<% End if %>
		<td align="left">Material</td>
		<td align="left">Norm</td>
		<td align="left">Quantity</td>
		<td align="left">Rate Material</td>
		<td align="left">Total Amount</td>		
		<td align="left">Luas Dikerjakan</td>
		<td align="left">Remarks</td>
	</tr>   
   <%
   if Not rsDet.eof then
      nCountofField=rsDet.fields.count-1
	  no=1
	  While Not rsDet.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if
	%>
		  <tr bgcolor="<%=cColor%>">
				<td align="left" valign="top"><%=rsDet("lineno")%></td>
			<!--  Validation Period Control In Here -->
				<td valign="middle">
					<a href="MonthlyWorkPlan.asp?txtYear=<%=years%>&txtMonth=<%=months%>&activitycode=<%=activity%>&divisi=<%=divisi%>&fieldcode=<%=fieldcode%>&stock=<%=rsDet("stockcode") %>&quantity=<%=rsDet("qty") %>&materialnorm=<%=rsDet("activematerialnorm") %>&kilogram=<%=rsDet("kg") %>&janjang=<%=rsDet("jjg") %>&rotasi=<%=rsDet("rotasi") %>&ratematerial=<%=rsDet("ratestock") %>&totalratematerial=<%=rsDet("totalamount") %>&remarks=<%=rsDet("remarks") %>&lineno=<%=rsDet("lineno")%>&transno=<%=rsDet("RKBNO")%>&fcbas=<%=rsDet("fcba")%>&selectedyear=<%=selectedyear%>&selectedmonth=<%=selectedmonth%>&selectedactivity=<%=selectedactivity%>&selectedafdeling=<%=selectedafdeling%>&selectedfield=<%=selectedfield%>&mandays=<%=mandays%>&volume=<%=volume%>&boronganpremi=<%=boronganpremi%>&ratemandays=<%=ratemandays%>&normamandays=<%=normamandays%>&normapremi=<%=normapremi%>&totalratemandays=<%=totalratemandays%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtYear")%>" border="0"></a>
				</td>
				<td valign="middle">
					<a href="MonthlyWorkPlan.asp?lineno=<%=rsDet("lineno")%>&transno=<%=rsDet("RKBNO")%>&fcbas=<%=rsDet("fcba")%>&selectedyear=<%=selectedyear%>&selectedmonth=<%=selectedmonth%>&selectedactivity=<%=selectedactivity%>&selectedafdeling=<%=selectedafdeling%>&selectedfield=<%=selectedfield%>&mandays=<%=mandays%>&volume=<%=volume%>&boronganpremi=<%=boronganpremi%>&ratemandays=<%=ratemandays%>&normamandays=<%=normamandays%>&normapremi=<%=normapremi%>&totalratemandays=<%=totalratemandays%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("transno")%>" border="0"></a>
				</td>
			<!-- END Validation Period Control In Here -->
			<% if activity = "6502101" then %>		  	
				<td align="left" title="<%=rsDet("kg").Name%>" valign="top"><%=rsDet("kg")%></td>
				<td align="left" title="<%=rsDet("jjg").Name%>" valign="top"><%=rsDet("jjg")%></td>
				<td align="left" title="<%=rsDet("rotasi").Name%>" valign="top"><%=rsDet("rotasi")%></td>
			<% End if %>
				<td align="left" title="<%=rsDet("fccode").Name%>" valign="top"><%=rsDet("fccode")%>-<%=rsDet("fcname")%></td>
				<td align="left" title="<%=rsDet("activematerialnorm").Name%>" valign="top"><%=rsDet("activematerialnorm")%></td>			
				<td align="left" title="<%=rsDet("qty").Name%>" valign="top"><%=rsDet("qty")%></td>			
				<td align="left" title="<%=rsDet("ratestock").Name%>" valign="top"><%=rsDet("ratestock")%></td>			
				<td align="left" title="<%=rsDet("totalamount").Name%>" valign="top"><%=rsDet("totalamount")%></td>							
				<td align="left" title="<%=rsDet("Volume").Name%>" valign="top"><%=rsDet("Volume")%></td>
				<td align="left" title="<%=rsDet("remarks").Name%>" valign="top"><%=rsDet("remarks")%></td>
		  </tr>
		  <%
		  rsDet.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsDet.Close
   set rsDet = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
<%'-------------akhir----------------------%>
</div></div>
<script src="plugins/tabcontent/tabcontent.js" type="text/javascript"></script>
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
	 Document.prototype.ready = function(callback) {
	  if(callback && typeof callback === 'function') {
		document.addEventListener("DOMContentLoaded", function() {
		  if(document.readyState === "interactive" || document.readyState === "complete") {
			return callback();
		  }
		});
	  }
	};
	document.ready(function() {
		getyear();
		getmonth();
		getactivity();
		getlistafdeling();
		getHKRate();
		gettargetcode();
		getFieldInfo();
		getratematerial();
	});
 </script>
</html>
