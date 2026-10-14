<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include file="DailyPayrollValidation.inc"-->
<html> 
<head>
<title>ePlaS - Gang Activity and Attendance</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">

<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script language="javascript" type="text/javascript" src="ajax/validate_ajax.js"></script>

<script language="javascript" type="text/javascript">

                function showHint(str)
					{
				var xmlhttp;
				var unit;
				var output;
				var oth;
				var jobname;
					if (str=="")
					  {
					  document.getElementById("mandays").innerHTML="";
					  document.getElementById("unit").innerHTML="";
					  document.getElementById("output").innerHTML="";
					  document.getElementById("oth").innerHTML="";
					  document.getElementById("jobname").innerHTML="";
				
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  xmlhttp=new XMLHttpRequest();
					  unit=new XMLHttpRequest();
					  output=new XMLHttpRequest();
					  oth=new XMLHttpRequest();
					  jobname=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
					  unit=new ActiveXObject("Microsoft.XMLHTTP");
					  output=new ActiveXObject("Microsoft.XMLHTTP");
					  oth=new ActiveXObject("Microsoft.XMLHTTP");
					  jobname=new ActiveXObject("Microsoft.XMLHTTP");
					  }
					  
					  
						xmlhttp.onreadystatechange=function()
						  {
						  if (xmlhttp.readyState==4 && xmlhttp.status==200)
							{
							document.getElementById("mandays").innerHTML=xmlhttp.responseText;
							}
						  }
						xmlhttp.open("GET","ajax/MANDAYS.asp?ajaxUoM=mandays&q="+str,true);
						xmlhttp.send();
						
						
						unit.onreadystatechange=function()
						  {
						  if (unit.readyState==4 && unit.status==200)
							{
							document.getElementById("unit").innerHTML=unit.responseText;
							}
						  }
						unit.open("GET","ajax/MANDAYS.asp?ajaxUoM=unit&q="+str,true);
						unit.send();
						
						output.onreadystatechange=function()
						  {
						  if (output.readyState==4 && output.status==200)
							{
							document.getElementById("output").innerHTML=output.responseText;
							}
						  }
						output.open("GET","ajax/MANDAYS.asp?ajaxUoM=output&q="+str,true);
						output.send();
						
						oth.onreadystatechange=function()
						  {
						  if (oth.readyState==4 && oth.status==200)
							{
							document.getElementById("oth").innerHTML=oth.responseText;
							}
						  }
						oth.open("GET","ajax/MANDAYS.asp?ajaxUoM=oth&q="+str,true);
						oth.send();
						
						jobname.onreadystatechange=function()
						  {
						  if (jobname.readyState==4 && jobname.status==200)
							{
							document.getElementById("jobname").innerHTML=jobname.responseText;
							}
						  }
						jobname.open("GET","ajax/MANDAYS.asp?ajaxUoM=jobname&q="+str,true);
						jobname.send();
					}
					
					function showtimefinger(str)
					{
				var attin;
				var attout;
				var dt= document.getElementById("fddate").value;
								
					if (str=="")
					  {
					  document.getElementById("Att_in").innerHTML="";
					  document.getElementById("Att_out").innerHTML="";
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  attin=new XMLHttpRequest();
					  attout=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  attin=new ActiveXObject("Microsoft.XMLHTTP");
					  attout=new ActiveXObject("Microsoft.XMLHTTP");
					  }
					  
					  
						
						attin.onreadystatechange=function()
						  {
						  if (attin.readyState==4 && attin.status==200)
							{
							document.getElementById("Att_in").innerHTML=attin.responseText;
							}
						  }
						attin.open("GET","ajax/MANDAYS.asp?ajaxUoM=timein&q="+str+"&d="+dt,true);
						attin.send();
						
						attout.onreadystatechange=function()
						  {
						  if (attout.readyState==4 && attout.status==200)
							{
							document.getElementById("Att_out").innerHTML=attout.responseText;
							}
						  }
						attout.open("GET","ajax/MANDAYS.asp?ajaxUoM=timeout&q="+str+"&d="+dt,true);
						attout.send();
						
					}		
</script>
<script language="JavaScript">
        function PopupTType(_fcba) {
		var txtttype = document.getElementById("locationtype").value;
		var txtfcba_ = document.getElementById("chargeunit_");
		var value_ = txtfcba_.options[txtfcba_.selectedIndex].value;
		
		if (_fcba == "" && value_ != ""){
		    _fcba = value_
			 txtttype = document.getElementById("chargetype_").value;
		}
			
		
		var newwin   = window.open("PrintData.Asp?txtModule=Data Master&txtTable=LocationCode&txtSearch="+txtttype+"&txtFcba="+_fcba+"&Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
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
var getOvertime = function () {
	// build up the post string when passing variables to the server side page
	
	var typex = document.getElementById("typex").value;
	var hours = document.getElementById("hours").value;
	var PostStr = 'typex='+typex+'&hours='+hours;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getOvertimeGAD.asp', 'POST', PostStr, showMessageResponse);
}



// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.convertion);
    // update the Div to show the result from the server
	document.getElementById("convertion").value = s.convertion;
	// document.getElementById("salescontractno").value = s.salescontractno;
}

var getHarvestingQualityData = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("jobcode").value;
	var fddate = document.getElementById("fddate").value;
	var empcode = document.getElementById("empcd").value;
	var fieldcode = document.getElementById("targetcode").value;
	var PostStr = 'jobcode='+jobcode+'&empcode='+empcode+'&fddate='+fddate+'&fieldcode='+fieldcode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getHarvestingQualityData.asp', 'POST', PostStr, showHarvestingQualityData);
}



// The function for handling the response from the server
var showHarvestingQualityData = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.convertion);
    // update the Div to show the result from the server
	// document.getElementById("salescontractno").value = s.salescontractno;
	//alert(s.isjobpanen);
	
	if (s.isjobpanen > 0) { 
		document.getElementById("kg_brondolan").style.display="block";
		document.getElementById("harvestingquality").style.display="block";
		$('#lblunit').text("BUNCH");
		document.getElementById("panenunits").style.display="block";
		document.getElementById("under_ripe").value=s.under_ripe;
		document.getElementById("long_stalk").value=s.long_stalk;
		document.getElementById("qe_6").value=s.qe_6;
		document.getElementById("qe_7").value=s.qe_7;
		document.getElementById("over_ripe").value=s.over_ripe;
		document.getElementById("qe_3").value=s.qe_3;
		document.getElementById("abnormal").value=s.abnormal;
		document.getElementById("unharvest_ffb").value=s.unharvest_ffb;
		document.getElementById("uncollect_lf_circle").value=s.uncollect_lf_circle;
		document.getElementById("unprune_frond").value=s.unprune_frond;
		document.getElementById("qe_5").value=s.qe_5;
		document.getElementById("qe_1").value=s.qe_1;
		document.getElementById("qe_2").value=s.qe_2;
		document.getElementById("eaten_by_rat").value=s.eaten_by_rat;

	} else {
		document.getElementById("kg_brondolan").style.display="none";
		document.getElementById("harvestingquality").style.display="none";
		$('#lblunit').text("UNIT");
		document.getElementById("panenunits").style.display="none";
		document.getElementById("under_ripe").value="";
		document.getElementById("long_stalk").value="";
		document.getElementById("qe_6").value="";
		document.getElementById("qe_7").value="";
		document.getElementById("over_ripe").value="";
		document.getElementById("qe_3").value="";
		document.getElementById("abnormal").value="";
		document.getElementById("unharvest_ffb").value="";
		document.getElementById("uncollect_lf_circle").value="";
		document.getElementById("unprune_frond").value="";
		document.getElementById("qe_5").value="";
		document.getElementById("qe_1").value="";
		document.getElementById("qe_2").value="";
		document.getElementById("eaten_by_rat").value="";
	}
}

var getJobcodeIntraco = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("jobcode").value;
	// var fddate = document.getElementById("fddate").value;
	// var empcode = document.getElementById("empcd").value;
	// var fieldcode = document.getElementById("targetcode").value;
	var PostStr = 'jobcode='+jobcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getChargeIntraco.asp', 'POST', PostStr, showJobcodeIntraco);
}

// The function for handling the response from the server
var showJobcodeIntraco = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.convertion);
    // update the Div to show the result from the server
	// document.getElementById("salescontractno").value = s.salescontractno;
	// alert(s.jobcodeintraco);
	// document.getElementById("jobcodeintraco").value=s.jobcodeintraco;

	if (s.jobcodeintraco > 0) { 
		document.getElementById("jobcodeintraco").style.display="block";
        document.getElementById("chargeunit").style.display="block";
		document.getElementById("chargejob").style.display="block";
		document.getElementById("chargetype").style.display="block";
		document.getElementById("chargecode").style.display="block";
	} else {
		document.getElementById("jobcodeintraco").style.display="none";
        document.getElementById("chargeunit").style.display="none";
		document.getElementById("chargejob").style.display="none";
		document.getElementById("chargetype").style.display="none";
		document.getElementById("chargecode").style.display="none";
	}
}

var getJobcodeSpbHarvesting = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("jobcode").value;
	// var fddate = document.getElementById("fddate").value;
	// var empcode = document.getElementById("empcd").value;
	// var fieldcode = document.getElementById("targetcode").value;
	var PostStr = 'jobcode='+jobcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getJobcodeHarvesting.asp', 'POST', PostStr, showJobcodeSpbHarvesting);
}

// The function for handling the response from the server
var showJobcodeSpbHarvesting = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.convertion);
    // update the Div to show the result from the server
	// document.getElementById("salescontractno").value = s.salescontractno;
	// alert(s.jobcodeintraco);
	// document.getElementById("jobcodeintraco").value=s.jobcodeintraco;

	if (s.jobcodespbharvesting > 0) { 
		document.getElementById("jobcodespbharvesting").style.display="block";
		document.getElementById("kg_janjang").style.display="block";
		document.getElementById("chitno").style.display="block";
		document.getElementById("vunit").style.display="block";
		document.getElementById("vunits").style.display="none";
		// document.getElementById("kg_brondolan").style.display="block";
	} else {
		document.getElementById("jobcodespbharvesting").style.display="none";
		document.getElementById("kg_janjang").style.display="none";
		document.getElementById("chitno").style.display="none";
		document.getElementById("vunit").style.display="none";
		document.getElementById("vunits").style.display="block";
		// document.getElementById("kg_brondolan").style.display="none";
	}
}

var getBjrSpbHarvesting = function () {
	// build up the post string when passing variables to the server side page
	var chitnoharvesting = document.getElementById("chitnoharvesting").value;
	var PostStr = 'chitnoharvesting='+chitnoharvesting;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getBjrSpbHarvesting.asp', 'POST', PostStr, showBjrSpbHarvesting);
}

// The function for handling the response from the server
var showBjrSpbHarvesting = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.bjr_per_chitno);
    // update the Div to show the result from the server
	document.getElementById("bjr_per_chitno").value = s.bjr_per_chitno;
	document.getElementById("kg_per_bucket").value = s.kg_per_bucket;
	// document.getElementById("salescontractno").value = s.salescontractno;
}

// launched from button click 
var getListChitno = function () {
	// build up the post string when passing variables to the server side page
	
	var fddate = document.getElementById("fddate").value;
	var targetcode = document.getElementById("targetcode").value;
	var PostStr = 'targetcode='+targetcode+'&fddate='+fddate;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getChitnoGAD.asp', 'POST', PostStr, showListChitno);
}

// The function for handling the response from the server
var showListChitno = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("listchitno").innerHTML = "";
	document.getElementById("listchitno").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"63%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getTragettype = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("jobcode").value;
	var PostStr = 'jobcode='+jobcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetPVInfo.asp', 'POST', PostStr, showTragettype);
	
}

// The function for handling the response from the server
var showTragettype = function (oXML) { 
     // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    //alert(s.target);
    // update the Div to show the result from the server
    document.getElementById("locationtype").value = s.target;
   		
	// document.getElementById("descpr").value = s.description; 
}

var getChargetype = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("chargejob_").value;
	var PostStr = 'jobcode='+jobcode;
	 //alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetPVInfo.asp', 'POST', PostStr, showChargetype);
	
}

// The function for handling the response from the server
var showChargetype = function (oXML) { 
     // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    //alert(s.target);
    // update the Div to show the result from the server
    document.getElementById("chargetype_").value = s.target;
   		
	// document.getElementById("descpr").value = s.description; 
}

// launched from button click 
var getGangcode = function () {
	// build up the post string when passing variables to the server side page
	
	var sectionname = document.getElementById("sectionname").value;
	var selectedgangcode = document.getElementById("selectedgangcode_").value;
	var PostStr = 'sectionname='+sectionname+'&gangcode='+selectedgangcode;
	 //alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistgangcode_bkm.asp', 'POST', PostStr, showGangcode);
}

// The function for handling the response from the server
var showGangcode = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("gangcodelist").innerHTML = "";
	document.getElementById("gangcodelist").innerHTML = response;

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

</script>
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
function periksa(form)
 { if(form.mandays.value > 1)
    {alert("mandays diisi paling besar 1 !");
     form.mandays.focus();
     form.mandays.select();
     return(false);
    }
 
  return(true);
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
<script>
	function hitungrumus() {
		var sbjr_per_chitno = $("#bjr_per_chitno").val();
		var skg_janjang_harvesting = $("#kg_janjang_harvesting").val();
		var skg_per_bucket = $("#kg_per_bucket").val();
		var sbucket = $("#bucket").val();
		// var newssalesprices = ssalesprices.replace(/,/g, '');
		
		//Untuk menghitung Total Estate Weight
		var totalunit = Math.round(skg_janjang_harvesting * sbjr_per_chitno) ; //a kali b
		// alert(totalunit);
		if (!isNaN(totalunit)) {
    	 	$("#unitharvesting").val(totalunit);
    	}
    	else{
    		$("#unitharvesting").val();
    	}
		
		//Untuk menghitung Total Kg Brondolan
		var totalkgbrd = Math.round(skg_per_bucket * sbucket) ; //a kali b
		//alert(totalkgbrd);
		if (!isNaN(totalkgbrd)) {
    	 	$("#hasilkg_brondolan").val(totalkgbrd);
    	}
    	else{
    		$("#hasilkg_brondolan").val();
    	}
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
 Dim GANGCODE, FDDATE, SUPERVISION_1, SUPERVISION_2, SUPERVISION_3, SUPERVISION_4, EMPLOYEECODE, ATTENDANCE, JOBCODE,BKMCODE
 Dim LOCATIONTYPE
 Dim LOCATIONCODE, MANDAYS, OTHRS, RATE, UNIT, OUTPUT, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA
 Dim cmdView
 Dim strSQL, rsSQL, IsMandatory, DefaultJob, Done, GangType
 Dim TOTMANDAYS,qq,OLDMANDAYS	
 Dim tempStringJob, JobCodeDescription, rsStringJob
 Dim tempStringMandays, MandaysDescription, rsStringMandays
 Dim tempStringUnit, UnitDescription, rsStringUnit
 Dim tempStringOutput, OutputDescription, rsStringOutput
 Dim tempStringOth, OthDescription, rsStringOth
 Dim qCheckStatus, txtClosingStatus, rsCheckStatus
 Dim recordAda
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

Dim LineNoKey : LineNoKey = Request("LineNoKey")
Dim querymaxlineno, rsmaxlineno
Dim queryentryvalidate, rsentryvalidate
Dim maxlinenokey : maxlinenokey = 0
Dim tempentryvalidate : tempentryvalidate = "FALSE"

Dim message_warning

'-- Fungsi Validasi Panen --'
'--Insert

Function f_c_output_harian(fcba__, fddate__, employeecode__, attendance__, jobcode__, targettype__, targetcode__, output__)
	Dim Result : Result = "Success"
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT validasi_att_temp_api.max_output('"&fcba__&"',  TO_DATE('"&fddate__&"', 'mm/dd/yyyy'), '"&employeecode__&"', '"&attendance__&"', '"&jobcode__&"', '"&targettype__&"', '"&targetcode__&"', '"&output__&"') result" & _
					" FROM DUAL"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = "Success"
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	f_c_output_harian = Result
End Function

Function f_c_output_rotasi(fcba__, fddate__, employeecode__, attendance__, jobcode__, targettype__, targetcode__, output__)
	Dim Result : Result = "Success"
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT validasi_att_temp_api.max_totaloutput('"&fcba__&"',  TO_DATE('"&fddate__&"', 'mm/dd/yyyy'), '"&employeecode__&"', '"&attendance__&"', '"&jobcode__&"', '"&targettype__&"', '"&targetcode__&"', '"&output__&"') result" & _
					" FROM DUAL"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = "Success"
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	f_c_output_rotasi = Result
End Function

'--Edit

Function f_u_output_harian(fcba__, linenokey__, attendance__, output__)
	Dim Result : Result = "Success"
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT validasi_att_temp_api.max_output('"&fcba__&"', '"&linenokey__&"', '"&attendance__&"', '"&output__&"') result" & _
					" FROM DUAL"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = "Success"
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	f_u_output_harian = Result
End Function

Function f_u_output_rotasi(fcba__, linenokey__, attendance__, output__)
	Dim Result : Result = "Success"
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT validasi_att_temp_api.max_totaloutput('"&fcba__&"', '"&linenokey__&"', '"&attendance__&"', '"&output__&"') result" & _
					" FROM DUAL"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = "Success"
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	f_u_output_rotasi = Result
End Function

'-- End Fungsi Validasi Panen --'


Function NewRecordDailyValidation(init_fcba, init_fddate, init_employeecode, init_attendance, init_mandays)
	Dim Result : Result = "NOENTRY"
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select ATTENDANCE_GAD_TEMP_API.dailyvalidation('"&init_fcba&"', to_date('"&init_fddate&"', 'mm/dd/yyyy'), '"&init_employeecode&"', '"&init_attendance&"', "&init_mandays&") result" & _
					" from dual"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = "NOENTRY"
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	NewRecordDailyValidation = Result
End Function

Function EditRecordDailyValidation(init_fcba, init_fddate, init_employeecode, init_attendance, init_mandays, init_linenokey)
	Dim Result : Result = "NOENTRY"
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select ATTENDANCE_GAD_TEMP_API.dailyvalidation('"&init_fcba&"', to_date('"&init_fddate&"', 'mm/dd/yyyy'), '"&init_employeecode&"', '"&init_attendance&"', "&init_mandays&", "&init_linenokey&") result" & _
					" from dual"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = "NOENTRY"
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	EditRecordDailyValidation = Result
End Function

Function FP_CheckUseFinger(in_fddate, in_fcba)
	Dim Result : Result = ""
 	Dim rsAvailable, strAvailable
	
 	strAvailable =	" select FiNGERVALIDATIONabc_API.Get_Validation('"&in_fcba&"',to_date('"&in_fddate&"', 'mm/dd/yyyy')) result" & _
 					" from dual"
					
				response.write strAvailable
 	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
 	rsAvailable.Open strAvailable, connstr, 3
	
 	If Not rsAvailable.EOF Then
 		Result = rsAvailable("result")
 	Else
 		Result = ""
 	End If
	
	strAvailable =	""
 	rsAvailable.Close
 	set rsAvailable = Nothing

 	FP_CheckUseFinger = Result
End Function

Function FP_CheckAttendance(in_employeecode, in_attendance, in_fddate, in_othrs, in_fcba)
 	Dim Result : Result = ""
 	Dim rsAvailable, strAvailable
	
 	strAvailable =	" select FiNGERVALIDATIONabc_API.checkingattendance('"&in_employeecode&"', '"&in_attendance&"', to_date('"&in_fddate&"', 'mm/dd/yyyy'), "&in_othrs&", '"&in_fcba&"') result" & _
 					" from dual"
					
 	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
 	rsAvailable.Open strAvailable, connstr, 3
	
 	If Not rsAvailable.EOF Then
 		Result = rsAvailable("result")
 	Else
 		Result = ""
 	End If
	
	strAvailable =	""
 	rsAvailable.Close
 	set rsAvailable = Nothing

 	FP_CheckAttendance = Result
 End Function

Function CheckCharge(initType, initCode, initJob, initFcba, initDate, initSource)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	"SELECT F_CekTarget_byDate ('"& initType &"'," & _
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

Sub DeleteAllowance
	if (len(trim(""&ALLOWANCE_CODE))) > 0 OR cdbl(to_nol(Amount)) > 0 then
		q2	=	" delete from attendance_empallded where linenokey='"& LineNoKey &"' and fcba ='"& session("fcba") &"'"
		' response.Write(q2)
		connStr.execute(q2)
	Else
		q2	=	" delete from attendance_empallded where linenokey='"& LineNoKey &"' and fcba ='"& session("fcba") &"'"
		' response.Write(q2)
		connStr.execute(q2)
	End If
End Sub

sub InsertAllowance
	qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& jobcode &"' and status='TRUE' group by fcba"			
	' response.write qLineMax
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		code = cdbl(to_nol(rsqLineMax("fccode")))
		'businessunit = rsqLineMax("fcba")
		businessunit = chargeunit 
	end if 
    rsqLineMax.Close
	set rsqLineMax = Nothing
	
	if code = 0 then
		if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 or len(trim(""&chargeunit)) <> 0 ) then
			err.number = 1001
			err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
		else
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
				q2	=	" Insert into attendance_empallded (" & _
						" LINENOKEY" & _
						" , ALLOWANCE_CODE" & _
						" , AMOUNT" & _
						" , FCENTRY" & _
						" , FCEDIT" & _
						" , FCIP" & _
						" , FCBA" & _
						" , LASTUPDATE" & _
						" , LASTTIME" & _
						" , REMARKSALLOWDED" & _
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
						" , '" & REMARKSALLOWDED & "'" & _
						" )"		
						' response.Write(q2)
						connStr.execute(q2)
				End if
			End if
		End If
	else
	if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 or len(trim(""&chargeunit)) = 0) then
		err.number = 1001
		err.description = "Anda harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
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
				q2	=	" Insert into attendance_empallded (" & _
						" LINENOKEY" & _
						" , ALLOWANCE_CODE" & _
						" , AMOUNT" & _
						" , FCENTRY" & _
						" , FCEDIT" & _
						" , FCIP" & _
						" , FCBA" & _
						" , LASTUPDATE" & _
						" , LASTTIME" & _
						" , REMARKSALLOWDED" & _
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
						" , '" & REMARKSALLOWDED & "'" & _
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
		' q2	=	" Insert into attendance_empallded (" & _
		' 		" LINENOKEY" & _
		' 		" , ALLOWANCE_CODE" & _
		' 		" , AMOUNT" & _
		' 		" , FCENTRY" & _
		' 		" , FCEDIT" & _
		' 		" , FCIP" & _
		' 		" , FCBA" & _
		' 		" , LASTUPDATE" & _
		' 		" , LASTTIME" & _
		' 		" ) " & _
		' 		" values ("& _
		' 		" '" & linenokey & "'" & _
		' 		" , '" & ALLOWANCE_CODE & "'" & _
		' 		" , replace('" & AMOUNT & "',',','')" & _
		' 		" , '" & Session("CodeName") & "'" & _
		' 		" , '" & Session("CodeName") & "'" & _
		' 		" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
		' 		" , '" & Session("fcBA") & "'" & _
		' 		" , sysdate" & _
		' 		" , To_Char(sysdate,'HH24:MI')" & _
		' 		" )"		
		' 		response.Write(q2)
		' 		connStr.execute(q2)
		End If
	End If
End Sub

Sub DeleteDataCharge
	q = " delete from ATTENDANCE_GAD_TEMP" & _
		" where fcba = '" & Session("fcBA") & "'" & _
		"       and linenokey = '" & LineNoKey & "'"
		' response.write q
	connStr.execute(q)
End Sub

Sub SetBKMCODE
    q = " UPDATE ATTENDANCE_GAD_TEMP" & _
		" SET BKMCODE =  'BKM/'||fcba||'/'||To_Char(fddate,'yyyymm')||'/'||SubStr(standard_hash((gangcode||To_Char(fddate,'yyyymmdd')||fcba),'MD5'),0,10) " & _
	    " WHERE fcba = '" & Session("fcBA") & "'" & _ 
	    " AND GangCode = '" & GANGCODE & "'" & _
		" AND FDDATE = to_date('" & fddate & "','mm/dd/yyyy')"  & _
		" AND BKMCODE is NULL " 
		' response.write q
	connStr.execute(q)

End Sub

Sub InsertDataCharge

	qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& jobcode &"' and status='TRUE' group by fcba"			
	' response.write qLineMax
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		code = cdbl(to_nol(rsqLineMax("fccode")))
		'businessunit = rsqLineMax("fcba")
		businessunit = chargeunit 
	end if 
	rsqLineMax.Close
	set rsqLineMax = Nothing

	if code = 0 then
		if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 or len(trim(""&chargeunit)) <> 0 ) then
			err.number = 1001
			err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
		else
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
			
			maxlinenokey = maxlinenokey + 1
			
			q = " SELECT COUNT (*) jobcodespbharvesting" & _
				"  FROM parameterdetail" & _
				" WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
				"       AND fcba = '"& session("fcba") &"'" & _
				"       AND fccode = '"& jobcode &"'"

			' response.Write q

			set rsrecurtrans = server.CreateObject("Adodb.recordset")
			rsrecurtrans.Open q, connstr

			if Not rsrecurtrans.eof then
				jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
			end if
			rsrecurtrans.Close
			set rsrecurtrans = Nothing

			' q = " SELECT FCCODE FROM DOCUMENT_CLASSIFICATION WHERE DOCUMENT_TYPE = 'AA'"

			' ' response.Write q

			' set rsrecurtrans = server.CreateObject("Adodb.recordset")
			' rsrecurtrans.Open q, connstr

			' if Not rsrecurtrans.eof then
			' 	document_classification = rsrecurtrans("FCCODE")
			' end if
			' rsrecurtrans.Close
			' set rsrecurtrans = Nothing
			
			if cdbl(to_nol(jobcodespbharvesting)) > 0 then 

			q = " Insert into ATTENDANCE_GAD_TEMP (GANGCODE" & _
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
				"                             , CHARGEUNIT" & _
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
				"                             , ABW" & _
				"                             , BUNCH_WEIGHT" & _
				" ) values (trim('" & GANGCODE & "')" & _
				"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
				"           , regexp_replace(trim('" & SUPERVISION_1 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_2 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_3 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_4 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & EMPLOYEECODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & ATTENDANCE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & JOBCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & locationtype & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & LOCATIONCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , trim('" & CHARGEUNIT & "')" & _
				"           , trim('" & CHARGEJOB & "')" & _
				"           , trim('" & CHARGETYPE & "')" & _
				"           , trim('" & CHARGECODE & "')" & _
				"           , trim('" & MANDAYS & "')" & _
				"           , trim('" & OTHRS & "')" & _
				"           , trim('" & RATE  & "')" & _
				"           , trim('" & UNITHARVESTING & "')" & _
				"           , trim('" & OUTPUT & "')" & _
				"           , '" & REFERENCE & "'" & _
				"           , '" & REMARKS & "'" & _
				"           , '" & BUCKET & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				"           , '" & Session("fcBA") & "'" & _
				"           , sysdate" & _
				"           , To_Char(sysdate,'HH24:MI')" & _
				"           , '" & maxlinenokey & "'" & _
				"           , 'Planned'" & _
				"           , '" & document_classification & "'" & _
				"           , '" & TYPE_OVERTIME & "'" & _
				"           , '" & OVERTIME_HOURS & "'" & _
				"           , '" & KG_BRONDOLAN & "'" & _
				"           , '" & SPBNO & "'" & _
				"           , '" & KG_JANJANG & "'" & _
				"           , '" & ABW & "'" & _
				"           , '" & BUNCH_WEIGHT & "'" & _
				" )"
			' response.Write(q)
			connStr.execute(q)
			else
			q = " Insert into ATTENDANCE_GAD_TEMP (GANGCODE" & _
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
				"                             , CHARGEUNIT" & _
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
				"                             , ABW" & _				
				"                             , BUNCH_WEIGHT" & _
				" ) values (trim('" & GANGCODE & "')" & _
				"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
				"           , regexp_replace(trim('" & SUPERVISION_1 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_2 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_3 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_4 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & EMPLOYEECODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & ATTENDANCE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & JOBCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & locationtype & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & LOCATIONCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , trim('" & CHARGEUNIT & "')" & _
				"           , trim('" & CHARGEJOB & "')" & _
				"           , trim('" & CHARGETYPE & "')" & _
				"           , trim('" & CHARGECODE & "')" & _
				"           , trim('" & MANDAYS & "')" & _
				"           , trim('" & OTHRS & "')" & _
				"           , trim('" & RATE  & "')" & _
				"           , trim('" & UNIT & "')" & _
				"           , trim('" & OUTPUT & "')" & _
				"           , '" & REFERENCE & "'" & _
				"           , '" & REMARKS & "'" & _
				"           , '" & BUCKET & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				"           , '" & Session("fcBA") & "'" & _
				"           , sysdate" & _
				"           , To_Char(sysdate,'HH24:MI')" & _
				"           , '" & maxlinenokey & "'" & _
				"           , 'Planned'" & _
				"           , '" & document_classification & "'" & _
				"           , '" & TYPE_OVERTIME & "'" & _
				"           , '" & OVERTIME_HOURS & "'" & _
				"           , '" & KG_BRONDOLAN & "'" & _
				"           , '" & SPBNO & "'" & _
				"           , '" & KG_JANJANG & "'" & _
				"           , '" & ABW & "'" & _
				"           , '" & BUNCH_WEIGHT & "'" & _
				" )"
			' response.Write(q)
			connStr.execute(q)
			end if 
		End If
	else
		if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 or len(trim(""&chargeunit)) = 0 ) then
			err.number = 1001
			err.description = "Anda harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
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
				
				maxlinenokey = maxlinenokey + 1
				q = " SELECT COUNT (*) jobcodespbharvesting" & _
				"  FROM parameterdetail" & _
				" WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
				"       AND fcba = '"& session("fcba") &"'" & _
				"       AND fccode = '"& jobcode &"'"

			' response.Write q

			set rsrecurtrans = server.CreateObject("Adodb.recordset")
			rsrecurtrans.Open q, connstr

			if Not rsrecurtrans.eof then
				jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
			end if
			rsrecurtrans.Close
			set rsrecurtrans = Nothing
			
			if cdbl(to_nol(jobcodespbharvesting)) > 0 then
			
			q = " Insert into ATTENDANCE_GAD_TEMP (GANGCODE" & _
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
				"                             , ABW" & _
				"                             , BUNCH_WEIGHT" & _
				" ) values (trim('" & GANGCODE & "')" & _
				"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
				"           , regexp_replace(trim('" & SUPERVISION_1 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_2 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_3 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_4 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & EMPLOYEECODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & ATTENDANCE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & JOBCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & locationtype & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & LOCATIONCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , trim('" & CHARGEUNIT & "')" & _
				"           , trim('" & CHARGEJOB & "')" & _
				"           , trim('" & CHARGETYPE & "')" & _
				"           , trim('" & CHARGECODE & "')" & _
				"           , trim('" & MANDAYS & "')" & _
				"           , trim('" & OTHRS & "')" & _
				"           , trim('" & RATE  & "')" & _
				"           , trim('" & UNITHARVESTING & "')" & _
				"           , trim('" & OUTPUT & "')" & _
				"           , '" & REFERENCE & "'" & _
				"           , '" & REMARKS & "'" & _
				"           , '" & BUCKET & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				"           , '" & Session("fcBA") & "'" & _
				"           , sysdate" & _
				"           , To_Char(sysdate,'HH24:MI')" & _
				"           , '" & maxlinenokey & "'" & _
				"           , 'Planned'" & _
				"           , '" & document_classification & "'" & _
				"           , '" & TYPE_OVERTIME & "'" & _
				"           , '" & OVERTIME_HOURS & "'" & _
				"           , '" & KG_BRONDOLAN & "'" & _
				"           , '" & SPBNO & "'" & _
				"           , '" & KG_JANJANG & "'" & _
				"           , '" & ABW & "'" & _
				"           , '" & BUNCH_WEIGHT & "'" & _
				" )"
			' response.Write(q)
			connStr.execute(q)
			else
				q = " Insert into ATTENDANCE_GAD_TEMP (GANGCODE" & _
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
				"                             , CHARGEUNIT" & _
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
				"                             , ABW" & _
				"                             , BUNCH_WEIGHT" & _
				" ) values (trim('" & GANGCODE & "')" & _
				"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
				"           , regexp_replace(trim('" & SUPERVISION_1 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_2 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_3 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_4 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & EMPLOYEECODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & ATTENDANCE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & JOBCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & locationtype & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & LOCATIONCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , trim('" & CHARGEUNIT & "')" & _
				"           , trim('" & CHARGEJOB & "')" & _
				"           , trim('" & CHARGETYPE & "')" & _
				"           , trim('" & CHARGECODE & "')" & _
				"           , trim('" & MANDAYS & "')" & _
				"           , trim('" & OTHRS & "')" & _
				"           , trim('" & RATE  & "')" & _
				"           , trim('" & UNIT & "')" & _
				"           , trim('" & OUTPUT & "')" & _
				"           , '" & REFERENCE & "'" & _
				"           , '" & REMARKS & "'" & _
				"           , '" & BUCKET & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				"           , '" & Session("fcBA") & "'" & _
				"           , sysdate" & _
				"           , To_Char(sysdate,'HH24:MI')" & _
				"           , '" & maxlinenokey & "'" & _
				"           , 'Planned'" & _
				"           , '" & document_classification & "'" & _
				"           , '" & TYPE_OVERTIME & "'" & _
				"           , '" & OVERTIME_HOURS & "'" & _
				"           , '" & KG_BRONDOLAN & "'" & _
				"           , '" & SPBNO & "'" & _
				"           , '" & KG_JANJANG & "'" & _
				"           , '" & ABW & "'" & _
				"           , '" & BUNCH_WEIGHT & "'" & _
				" )"
			' response.Write(q)
			connStr.execute(q)
			end if 
			End If
		End If
	End If
	Call SetBKMCODE
End Sub

Sub UpdateDataCharge
	qLineMax = " select count(1) fccode, fcba from intraco_tab where fccode = '"& jobcode &"' and status='TRUE' group by fcba"			
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		code = cdbl(to_nol(rsqLineMax("fccode")))
		'businessunit = rsqLineMax("fcba")
		businessunit = chargeunit 
	end if 
	rsqLineMax.Close
	set rsqLineMax = Nothing

if code = 0 then
		if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 or len(trim(""&chargeunit)) <> 0 ) then
			err.number = 1001
			err.description = "Anda tidak harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
		else
		Call DeleteDataCharge
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
			
			maxlinenokey = maxlinenokey + 1
			q = " SELECT COUNT (*) jobcodespbharvesting" & _
				"  FROM parameterdetail" & _
				" WHERE     parcode = 'JOBCODE_PENGANGKUTAN'" & _
				"       AND fcba = '"& session("fcba") &"'" & _
				"       AND fccode = '"& jobcode &"'"

			' response.Write q

			set rsrecurtrans = server.CreateObject("Adodb.recordset")
			rsrecurtrans.Open q, connstr

			if Not rsrecurtrans.eof then
				jobcodespbharvesting = rsrecurtrans("jobcodespbharvesting")
			end if
			rsrecurtrans.Close
			set rsrecurtrans = Nothing
			
			if cdbl(to_nol(jobcodespbharvesting)) > 0 then
			
			q = " Insert into ATTENDANCE_GAD_TEMP (GANGCODE" & _
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
				"                             , CHARGEUNIT" & _
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
				"                             , ABW" & _
				"                             , BUNCH_WEIGHT" & _
				" ) values (trim('" & GANGCODE & "')" & _
				"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
				"           , regexp_replace(trim('" & SUPERVISION_1 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_2 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_3 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_4 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & EMPLOYEECODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & ATTENDANCE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & JOBCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & locationtype & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & LOCATIONCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , trim('" & CHARGEUNIT & "')" & _
				"           , trim('" & CHARGEJOB & "')" & _
				"           , trim('" & CHARGETYPE & "')" & _
				"           , trim('" & CHARGECODE & "')" & _
				"           , trim('" & MANDAYS & "')" & _
				"           , trim('" & OTHRS & "')" & _
				"           , trim('" & RATE  & "')" & _
				"           , trim('" & UNITHARVESTING & "')" & _
				"           , trim('" & OUTPUT & "')" & _
				"           , '" & REFERENCE & "'" & _
				"           , '" & REMARKS & "'" & _
				"           , '" & BUCKET & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				"           , '" & Session("fcBA") & "'" & _
				"           , sysdate" & _
				"           , To_Char(sysdate,'HH24:MI')" & _
				"           , '" & maxlinenokey & "'" & _
				"           , 'Planned'" & _
				"           , '" & document_classification & "'" & _
				"           , '" & TYPE_OVERTIME & "'" & _
				"           , '" & OVERTIME_HOURS & "'" & _
				"           , '" & KG_BRONDOLAN & "'" & _
				"           , '" & SPBNO & "'" & _
				"           , '" & KG_JANJANG & "'" & _
				"           , '" & ABW & "'" & _
				"           , '" & BUNCH_WEIGHT & "'" & _
				" )"
			' response.Write(q)
			connStr.execute(q)
			else
			OTHRS = request ("OTHRS")
				q = " Insert into ATTENDANCE_GAD_TEMP (GANGCODE" & _
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
				"                             , CHARGEUNIT" & _
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
				"                             , ABW" & _
				"                             , BUNCH_WEIGHT" & _
				" ) values (trim('" & GANGCODE & "')" & _
				"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
				"           , regexp_replace(trim('" & SUPERVISION_1 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_2 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_3 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & SUPERVISION_4 & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & EMPLOYEECODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & ATTENDANCE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & JOBCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & locationtype & "'),'[^A-Z0-9 ]','')" & _
				"           , regexp_replace(trim('" & LOCATIONCODE & "'),'[^A-Z0-9 ]','')" & _
				"           , trim('" & CHARGEUNIT & "')" & _
				"           , trim('" & CHARGEJOB & "')" & _
				"           , trim('" & CHARGETYPE & "')" & _
				"           , trim('" & CHARGECODE & "')" & _
				"           , trim('" & MANDAYS & "')" & _
				"           , trim('" & OTHRS & "')" & _
				"           , trim('" & RATE  & "')" & _
				"           , trim('" & UNIT & "')" & _
				"           , trim('" & OUTPUT & "')" & _
				"           , '" & REFERENCE & "'" & _
				"           , '" & REMARKS & "'" & _
				"           , '" & BUCKET & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Session("CodeName") & "'" & _
				"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				"           , '" & Session("fcBA") & "'" & _
				"           , sysdate" & _
				"           , To_Char(sysdate,'HH24:MI')" & _
				"           , '" & maxlinenokey & "'" & _
				"           , 'Planned'" & _
				"           , '" & document_classification & "'" & _
				"           , '" & TYPE_OVERTIME & "'" & _
				"           , '" & OVERTIME_HOURS & "'" & _
				"           , '" & KG_BRONDOLAN & "'" & _
				"           , '" & SPBNO & "'" & _
				"           , '" & KG_JANJANG & "'" & _
				"           , '" & ABW & "'" & _
				"           , '" & BUNCH_WEIGHT & "'" & _
				" )"
			' response.Write(q)
			connStr.execute(q) 
			end if 
		End If
	Else
		Call DeleteDataCharge
		Call InsertDataCharge
	end if
	CAll SetBKMCODE
End Sub


Dim attInRecord : attInRecord = "NOENTRY"
Dim fp_att_cek : fp_att_cek = 0
Dim use_fp : use_fp = "FALSE"

' Added Variable ChargeJob, ChargeType, ChargeCode
chargeunit=Request("chargeunit")
chargejob=Request("chargejob")
chargetype=Request("chargetype")
chargecode=Request("chargecode")
'businessunit=Request("businessunit")
businessunit=Request("chargeunit")

' End Added

TYPE_OVERTIME= request("TYPE_OVERTIME")
if len(trim(""&TYPE_OVERTIME))= 0 then TYPE_OVERTIME = 1


 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "StartAuthorize" Then
	GANGCODE=request("GangCode")
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
	document_classification=Request("document_classification")
	selectedgangcode=Request("selectedgangcode")
	sectionname=Request("sectionname")
	remarksallowded=Request("remarksallowded")
    BKMCODE = Request("BKMCODE")
	
	ABW = Request("ABW")
	BUNCH_WEIGHT = Request("BUNCH_WEIGHT")
	
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
	' End Added
	
	' Added Variable Allowded dan Overtime By Sendy 12/08/2018
	
	OVERTIME_HOURS=To_Nol(request("OVERTIME_HOURS"))
	ALLOWANCE_CODE=To_Nol(request("ALLOWANCE_CODE"))
	AMOUNT=To_Nol(request("AMOUNT"))	
	'End Added 
	
	'********* ADD BY UKIE 08 Juni 2014 '*********
	' Nama Table = PARAMETER
	' dengan kondisi FCBA = xxx PARCODE = E0102
	' PARAM_1 = IsMandays
	' PARAM_2 = IsUnit
	' PARAM_3 = IsOutput
	' PARAM_4 = IsOthrs
	' PARAM_5 = IsRate
	

	Dim ValidateIsMandays : ValidateIsMandays = ""
	Dim ValidateIsUnit : ValidateIsUnit = ""
	Dim ValidateIsOutput : ValidateIsOutput = ""
	Dim ValidateIsOthrs : ValidateIsOthrs = ""
	Dim ValidateIsRate : ValidateIsRate = ""
	
	Dim DoneIsMandays : DoneIsMandays = 0
	Dim DoneIsUnit : DoneIsUnit = 0
	Dim DoneIsOutput : DoneIsOutput = 0
	Dim DoneIsOthrs : DoneIsOthrs = 0
	Dim DoneIsRate : DoneIsRate = 0
	
	Dim rsParameter, strParameter
	strParameter = " select nvl(param_1, 'FALSE') IsMandays" & _
	               "        , nvl(param_2, 'FALSE') IsUnit" & _
	               "        , nvl(param_3, 'FALSE') IsOutput" & _
	               "        , nvl(param_4, 'FALSE') IsOthrs" & _
	               "        , nvl(param_5, 'FALSE') IsRate" & _
				   " from parameter" & _
				   " where fcba = '" & Session("fcBA") & "'" & _
				   "       and parcode = 'E0102'" & _
				   "       and fccode = '" & ATTENDANCE & "'"
	'response.write strParameter
				   
	set rsParameter = server.CreateObject("Adodb.recordset")
	rsParameter.Open strParameter , connstr, adOpenStatic

	If Len(Trim(ATTENDANCE)) > 0 Then
		If Not rsParameter.EOF Then
			ValidateIsMandays = rsParameter("IsMandays")
			ValidateIsUnit = rsParameter("IsUnit")
			ValidateIsOutput = rsParameter("IsOutput")
			ValidateIsOthrs = rsParameter("IsOthrs")
			ValidateIsRate = rsParameter("IsRate")
		Else
			Err.Number = 1001
			Err.Description = " Parameter belum di set! Contact your Administrator"
		end if
		rsParameter.Close()
		Set rsParameter = Nothing
	End If
	
	If ValidateIsMandays = "TRUE" Then
		If Cdbl(MANDAYS) <= 0 Then
			Err.Number = 1001
			Err.Description = " Mandays Harus di isi, Range Mandays dimulai dari 0.01 s/d 1 !"
		ElseIf Cdbl(MANDAYS) > 1 Then
			Err.Number = 1001
			Err.Description = " Mandays tidak boleh melebihi angka 1 !"
		Else
			DoneIsMandays = 1
		End If
	ElseIf ValidateIsMandays = "FALSE" Then
		If Cdbl(MANDAYS) <> 0 Then
			Err.Number = 1001
			Err.Description = " Mandays di kosongkan saja atau masukkan angka ""0"" !"
		Else
			DoneIsMandays = 1
		End If
	End If
	
	'response.write(UNIT)
	If ValidateIsUnit = "TRUE" Then
	 	If (Cdbl(UNIT) < 0 AND Trim(JOBCODE) <> "6502101") or (Cdbl(UNIT) <= 0 AND Cdbl(OUTPUT) <= 0 AND Trim(JOBCODE) = "6502101" AND ATTENDANCE<> "MK") Then
			Err.Number = 1001
			Err.Description = " Unit Harus di isi! / untuk Panen Unit & Output Ha Harus diisi !"
		Else
			DoneIsUnit = 1
		End If
	ElseIf ValidateIsUnit = "FALSE" Then
		If Cdbl(UNIT) <> 0 Then
			Err.Number = 1001
			Err.Description = " Unit di kosongkan saja atau masukkan angka ""0"" !"
		Else
			DoneIsUnit = 1
		End If
	End If
	
	If ValidateIsOutput = "TRUE" Then
		If Cdbl(OUTPUT) < 0 Then
			Err.Number = 1001
			Err.Description = " Output Harus di isi!. Range Output dimulai dari 0 !"
		Else
			DoneIsOutput = 1
		End If
	ElseIf ValidateIsOutput = "FALSE" Then
		If Cdbl(OUTPUT) <> 0 Then
			Err.Number = 1001
			Err.Description = " Output di kosongkan saja atau masukkan angka ""0"" !"
		Else
			DoneIsOutput = 1
		End If
	End If
	
	If ValidateIsOthrs = "TRUE" Then
		If Cdbl(OTHRS) < 0 Then
			Err.Number = 1001
			Err.Description = " Overtime Harus di isi!. Range Overtime dimulai dari 0 !"
		Else
			DoneIsOthrs = 1
		End If
	ElseIf ValidateIsOthrs = "FALSE" Then
		If Cdbl(OTHRS) <> 0 Then
			Err.Number = 1001
			Err.Description = " Overtime di kosongkan saja atau masukkan angka ""0"" !"
		Else
			DoneIsOthrs = 1
		End If
	End If
	
	If ValidateIsRate = "TRUE" Then
		If Cdbl(RATE) < 0 Then
			Err.Number = 1001
			Err.Description = " Rate Harus di isi!. Range Rate dimulai dari 0 !"
		Else
			DoneIsRate = 1
		End If
	ElseIf ValidateIsRate = "FALSE" Then
		If Cdbl(RATE) <> 0 Then
			Err.Number = 1001
			Err.Description = " Rate di kosongkan saja atau masukkan angka ""0"" !"
		Else
			DoneIsRate = 1
		End If
	End If
	

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
	
	
	
	q	=	" select sum(mandays)  mandays1 from ATTENDANCE_GAD_TEMP where  fdDate = to_DATE('" & fdDate & "','mm/dd/yyyy')" & _
			" and EmployeeCode=trim('" & EmployeeCode & "')  and fcBA=trim('" & Session("fcBA") & "')"
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
	 
	q	= " select 1 from ATTENDANCE_GAD_TEMP where FDDATE = to_date('" & fdDate & "', 'mm/dd/yyyy')" & _
		  " and EmployeeCode = trim('" & EmployeeCode & "')" & _				
		  " And locationtype = trim('" & locationtype & "')" & _
		  " and LOCATIONCODE = trim('" & LOCATIONCODE & "')" & _
		  " and JOBCODE = trim('" & JOBCODE & "')" & _
		  " and fcBA= trim('" & Session("fcBA") & "')"
		  
	set rs = connStr.execute(q)
	 if not rs.eof then
	      recordAda = "T"
	 end if
	 rs.close()
     rs=null 
	 
	 ' End Add By Akbar 10 Feb 2014
	 
	q	=	" select sum(mandays)  mandays1 from ATTENDANCE_GAD_TEMP " & _
			" where  fdDate = to_DATE('" & fdDate & "','mm/dd/yyyy')" & _
			" and EmployeeCode=trim('" & EmployeeCode & "') And locationtype=trim('" & locationtype & "') and " & _
			" LOCATIONCODE=trim('" & LOCATIONCODE & "') and JOBCODE=trim('" & JOBCODE & "') and fcBA=trim('" & Session("fcBA") & "')"
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
		" fcba = trim('" & Session("fcBA") & "')"
		
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
				" fccode = trim('" & ATTENDANCE & "') AND " & _
				" fcba = '"& Session("fcBA") &"' "
	
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
	
	if done = 1 then 
		if Request("txtJenis")= "Delete" then
		

			'-- Check for Empallded PreProcess
			qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreAttendanceUpdate (to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'YYYY'))," & _ 
																				  " to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'MM')),'" & _ 
																					Session("fcBA") & "') closingStatus from dual "

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
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fdDate)
				If Cek_IsDailyClosing = "FALSE" then
					q = " delete from ATTENDANCE_GAD_TEMP" & _
						" where fcba = '" & Session("fcBA") & "'" & _
						"       and linenokey = '" & LineNoKey & "'"
					connStr.execute(q)
					
					Call DeleteAllowance

					q	=	" Delete from HarvestingQuality where" & _
							" EMPCODE = '" & EMPLOYEECODE & "'" & _
							" And FIELDCODE = '" & LOCATIONCODE & "'" & _
							" and fddate = to_Date('" & fdDate & "','MM/DD/YYYY')" & _
							" and fcBA = '" & Session("fcBA") & "'"
					
					'response.Write(q)
					connStr.execute(q)
					cmdView	= "True"
				else
					err.number = 1001
					err.description = "Tanggal """ & fdDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
			else
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Unable to Update/Delete/Insert," & "</font>"
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Please OPEN your payroll Status </font>"
				cmdView	= "True"			
			End if
			
		elseif Request("txtJenis")="New" then
		
			'-- Check for Empallded PreProcess
			qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreAttendanceUpdate (to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'YYYY'))," & _ 
																				  " to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'MM')),'" & _ 
																					Session("fcBA") & "') closingStatus from dual "

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
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fdDate)
				If Cek_IsDailyClosing = "FALSE" then
					If DoneIsMandays = 1 And DoneIsUnit = 1 And DoneIsOutput = 1 And DoneIsOthrs = 1 And DoneIsRate = 1 Then
			
						' Add by NAF, 11-10-2012
						if IsMandatory = "FALSE" then 
							if JOBCODE = "" then 
							   JOBCODE = defaultJob
							end if	
						end if
						' End of Modification
						
						'Jika Total Dia bekerja bukan sebagai Gang Mandor dan KJ, maka Mandays Harus Lebih besar dari Nol
						
						IF ( cdbl(mandays) = 0 ) and ( gangtype <> "SUPERVISION" ) and ( ATTENDANCE = "KJ" ) THEN 
							response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Gang Type is not a Supervision Gang and its a KJ attendance, the mandays should not be ZERO " & "</font>"
							cmdView	= "True"
						ELSE			
							' IF recordAda = "F" Then
								IF ( cdbl(TOTMANDAYS) + MANDAYS ) <= 1 then     'Add Akbar 4 des 2012
								
									'-- Get Data Entry Validate if KHL not allowed over20 Days, else free
									queryentryvalidate = " select nvl(ATTENDANCE_GAD_TEMP_API.ValidasiHKKHL('" & Session("fcBA") & "'" & _
														 "                                             , to_date('" & fddate & "', 'mm/dd/yyyy')" & _
														 "                                             , '" & EMPLOYEECODE & "'" & _
														 "                                             , '" & ATTENDANCE & "'" & _
														 "                                             , " & Cdbl(MANDAYS) & "), 'FALSE') entryvalidate" & _ 
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
										attInRecord = NewRecordDailyValidation(Session("fcBA"), fddate, EMPLOYEECODE, ATTENDANCE, MANDAYS)
										
										'Response.Write attInRecord
										
										if attInRecord = "ENTRY" then
										
											'Validasi terkait data di fingerprint
											 fp_att_cek = FP_CheckAttendance(EMPLOYEECODE, ATTENDANCE, fddate, 0, Session("fcBA"))
											' response.write " Record = " & fp_att_cek & " | length = " & len(trim(fp_att_cek))
											 if fp_att_cek = "Validation Passed" then
											
												c_output_harian = f_c_output_harian(Session("fcBA"), fddate, EMPLOYEECODE, ATTENDANCE, JOBCODE, locationtype, LOCATIONCODE, OUTPUT)
												c_output_rotasi = f_c_output_rotasi(Session("fcBA"), fddate, EMPLOYEECODE, ATTENDANCE, JOBCODE, locationtype, LOCATIONCODE, OUTPUT)
												
												if c_output_harian = "Success" then
													if c_output_rotasi = "Success" then
									
														' '-- Get Data Max Linenokey
														' querymaxlineno = " select nvl(max(linenokey), 0) linenokey" & _ 
														' 				 " from attendance_gad" & _ 
														' 				 " where fcba = '" & Session("fcBA") & "'"

														' Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
														' rsmaxlineno.CursorType = 3
														' rsmaxlineno.CursorLocation = 3
														
														' 'response.write querymaxlineno
														
														' rsmaxlineno.Open querymaxlineno, connstr
														' if not rsmaxlineno.eof then
														'    maxlinenokey = Cdbl(rsmaxlineno("linenokey"))
														' else
														'    maxlinenokey = 0
														' end if
														' rsmaxlineno.close
														' set rsmaxlineno = Nothing
														
														' maxlinenokey = maxlinenokey + 1
														
														' q = " Insert into Attendance_GAD (GANGCODE" & _
														' 	"                             , FDDATE" & _
														' 	"                             , SUPERVISION_1" & _
														' 	"                             , SUPERVISION_2" & _
														' 	"                             , SUPERVISION_3" & _
														' 	"                             , SUPERVISION_4" & _
														' 	"                             , EMPLOYEECODE" & _
														' 	"                             , ATTENDANCE" & _
														' 	"                             , JOBCODE" & _
														' 	"                             , LOCATIONTYPE" & _
														' 	"                             , LOCATIONCODE" & _
														' 	"                             , MANDAYS" & _
														' 	"                             , OTHRS" & _
														' 	"                             , RATE" & _
														' 	"                             , UNIT" & _
														' 	"                             , OUTPUT" & _
														' 	"                             , REFERENCE" & _
														' 	"                             , REMARKS" & _
														' 	"                             , FCENTRY" & _
														' 	"                             , FCEDIT" & _
														' 	"                             , FCIP" & _
														' 	"                             , FCBA" & _
														' 	"                             , LASTUPDATE" & _
														' 	"                             , LASTTIME" & _
														' 	"                             , linenokey" & _
														' 	"                             , TYPE_OVERTIME" & _
														' 	"                             , OVERTIME_HOURS" & _
														' 	" ) values (trim('" & GANGCODE & "')" & _
														' 	"           , to_date('" & fddate & "','mm/dd/yyyy')" & _
														' 	"           , trim('" & SUPERVISION_1 & "')" & _
														' 	"           , trim('" & SUPERVISION_2 & "')" & _
														' 	"           , trim('" & SUPERVISION_3 & "')" & _
														' 	"           , trim('" & SUPERVISION_4 & "')" & _
														' 	"           , trim('" & EMPLOYEECODE & "')" & _
														' 	"           , trim('" & ATTENDANCE & "')" & _
														' 	"           , trim('" & JOBCODE & "')" & _
														' 	"           , trim('" & locationtype & "')" & _
														' 	"           , trim('" & LOCATIONCODE & "')" & _
														' 	"           , trim('" & MANDAYS & "')" & _
														' 	"           , trim('" & OTHRS & "')" & _
														' 	"           , trim('" & RATE  & "')" & _
														' 	"           , trim('" & UNIT & "')" & _
														' 	"           , trim('" & OUTPUT & "')" & _
														' 	"           , '" & REFERENCE & "'" & _
														' 	"           , '" & REMARKS & "'" & _
														' 	"           , '" & Session("CodeName") & "'" & _
														' 	"           , '" & Session("CodeName") & "'" & _
														' 	"           , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
														' 	"           , '" & Session("fcBA") & "'" & _
														' 	"           , sysdate" & _
														' 	"           , To_Char(sysdate,'HH24:MI')" & _
														' 	"           , '" & maxlinenokey & "'" & _
														' 	"           , '" & TYPE_OVERTIME & "'" & _
														' 	"           , '" & OVERTIME_HOURS & "'" & _
														' 	" )"
														' ' response.Write(q)
														' connStr.execute(q)
														if len(trim(""&ALLOWANCE_CODE)) = "A99" then
															err.number = 1001
															err.description = "Anda Harus Mengisi Kolom AMOUNT ALLOWED !"
															response.write allowance_code
														Else
															Call InsertDataCharge
															call InsertAllowance
														End if
														' response.Write (request(JOBCODE))

														strFieldDetail = " SELECT pr.parcode, pr.parname, prd.fccode, prd.parheadcode, prd.fcname" & _
																		 "	FROM parameter pr " & _
																		 "	INNER JOIN parameterdetail prd ON (pr.parcode = prd.parcode and pr.fcba = prd.fcba)" & _
																	 	 "	WHERE pr.parcode = 'HARDCODE' AND pr.fcba = '"& session("fcBA") &"'"				
																	 	 	' Response.Write strFieldDetail

														Set rsQty = Server.CreateObject("ADODB.Recordset")
														rsQty.Open strFieldDetail, connstr							
														if not rsQty.Eof then
														do while not rsQty.eof
																JobPanen = rsQty("fccode")
																rsQty.MoveNext
															loop
														End If

														If (len(trim(""&JobPanen))) <> 0 then
															if (cdbl(to_nol(UNDER_RIPE)) <> 0 OR cdbl(to_nol(OVER_RIPE)) <> 0 OR cdbl(to_nol(ABNORMAL)) <> 0 OR cdbl(to_nol(LONG_STALK)) <> 0 OR cdbl(to_nol(EATEN_BY_RAT)) <> 0 OR cdbl(to_nol(UNHARVEST_FFB)) <> 0 OR cdbl(to_nol(UNCOLLECT_LF_CIRCLE)) <> 0 OR cdbl(to_nol(UNPRUNE_FROND)) <> 0 OR cdbl(to_nol(QE_1)) <> 0 OR cdbl(to_nol(QE_2)) <> 0 OR cdbl(to_nol(QE_3)) <> 0 OR cdbl(to_nol(QE_5)) <> 0 OR cdbl(to_nol(QE_6)) <> 0 OR cdbl(to_nol(QE_7)) <> 0 )then
														
															' err.number = 1001
															' err.description = "Anda tidak menginput Harvesting Quality"
															' ' response.Write(jobcode)
															' Call InsertDataCharge
															' call InsertAllowance

															q1	=	" Insert into HarvestingQuality (" & _
																	" EMPCODE" & _
																	" , FDDATE" & _
																	" , FIELDCODE" & _
																	" , UNDER_RIPE" & _
																	" , OVER_RIPE" & _
																	" , ABNORMAL" & _
																	" , LONG_STALK" & _
																	" , EATEN_BY_RAT" & _
																	" , UNHARVEST_FFB" & _
																	" , UNCOLLECT_LF_CIRCLE" & _
																	" , UNCOLLECT_LF_PIECE" & _
																	" , UNARRANGE_FFB" & _
																	" , UNPRUNE_FROND" & _
																	" , QE_1" & _
																	" , QE_2" & _
																	" , QE_3" & _
																	" , QE_4" & _
																	" , QE_5" & _
																	" , QE_6" & _
																	" , QE_7" & _
																	" , QE_8" & _
																	" , QE_9" & _
																	" , QE_10" & _
																	" , FCENTRY" & _
																	" , FCEDIT" & _
																	" , FCIP" & _
																	" , FCBA" & _
																	" , LASTUPDATE" & _
																	" , LASTTIME" & _
																	" ) " & _
																	" values ("& _
																	" '" & EMPLOYEECODE & "'," & _
																	" to_date('" & fddate & "','mm/dd/yyyy')" & _
																	" , '" & LOCATIONCODE & "'" & _
																	" , '" & UNDER_RIPE & "'" & _
																	" , '" & OVER_RIPE & "'" & _
																	" , '" & ABNORMAL & "'" & _
																	" , '" & LONG_STALK & "'" & _
																	" , '" & EATEN_BY_RAT & "'" & _
																	" , '" & UNHARVEST_FFB & "'" & _
																	" , '" & UNCOLLECT_LF_CIRCLE & "'" & _
																	" , '" & UNCOLLECT_LF_PIECE & "'" & _
																	" , '" & UNARRANGE_FFB & "'" & _
																	" , '" & UNPRUNE_FROND & "'" & _
																	" , '" & QE_1 & "'" & _
																	" , '" & QE_2 & "'" & _
																	" , '" & QE_3 & "'" & _
																	" , '" & QE_4 & "'" & _
																	" , '" & QE_5 & "'" & _
																	" , '" & QE_6 & "'" & _
																	" , '" & QE_7 & "'" & _
																	" , '" & QE_8 & "'" & _
																	" , '" & QE_9 & "'" & _
																	" , '" & QE_10 & "'" & _
																	" , '" & Session("CodeName") & "'" & _
																	" , '" & Session("CodeName") & "'" & _
																	" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
																	" , '" & Session("fcBA") & "'" & _
																	" , sysdate" & _
																	" , To_Char(sysdate,'HH24:MI')" & _
																	" )"
																	
																' response.Write(q1)
																' response.Write(jobcode)
																response.Flush()
																connStr.execute(q1)
															' err.number = 1001
															' err.description = "Anda memilih Jobcode Potong Buah, Silahkan isi data harvesting Quality"
															' response.Write(jobcode)
														end if
													End if
														cmdView	= "True"
														' end if
													else
														err.number = 1001
														err.description = c_output_rotasi & " <br/>"
													end if
												else
													err.number = 1001
													err.description = c_output_harian & " <br/>"
												end if
											 else
											    ' No Finger
											 	err.number = 1001
											 	err.description = fp_att_cek & " <br/>"
											 end if
										else
											err.number = 1001
											err.description = "Anda sudah memasukkan data kehadiran/ Maksimal 1 HK!!!" & "<br/>" & _
															  "Silahkan cek kembali data yang anda ingin/sudah di masukkan ke dalam sistem!" & " <br/>"
										end if
									else
										if tempentryvalidate = "NO_INPUT" then
											err.number = 1001
											err.description = "Karyawan tersebut belum di proses perubahan status-nya sehingga hk tidak bisa di input!" & "<br/>" & _
															  "Untuk detailnya silahkan tanya kepada manajemen anda, Terima Kasih!" & " <br/>"
										else
											err.number = 1001
											err.description = "Maksimal HK yang di perbolehkan untuk KHL adalah 20 HK, Cek kembali inputan anda!" & "<br/>" & _
															  "Untuk detailnya silahkan tanya kepada manajemen anda, Terima Kasih!" & " <br/>"
										end if
									end if
								ELSE
									'Add Akbar 4 des 2012 begin
									response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Total mandays perhari 1 ." & "</font>"
									response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Total mandays yang sudah di input tgl "& fdDate &" adalah " & cdbl(TOTMANDAYS) & " sisa mandays yg bisa di input adalah " &  1 -  (cdbl(TOTMANDAYS)) & "</font>"
									cmdView	= "True"
									'Add Akbar 4 des 2012 end
								END IF	
							' ELSE
								' 'Add Akbar 10 Des 2014 begin
									' response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Record Sudah Ada" & "</font>"
									' cmdView	= "True"
								' 'Add Akbar 10 Des 2014 end
							' END IF		
						END IF
					End If
				else
					err.number = 1001
					err.description = "Tanggal """ & fdDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
			else
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Unable to Update/Delete/Insert," & "</font>"
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Please OPEN your payroll Status </font>"
				cmdView	= "True"			
			end if
			
			cmdView	= "True"
		   
		elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		    
			
			'-- Check for Empallded PreProcess
			qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreAttendanceUpdate (to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'YYYY'))," & _ 
																				  " to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'MM')),'" & _ 
																					Session("fcBA") & "') closingStatus from dual "

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
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fdDate)
				If Cek_IsDailyClosing = "FALSE" then
					If DoneIsMandays = 1 And DoneIsUnit = 1 And DoneIsOutput = 1 And DoneIsOthrs = 1 And DoneIsRate = 1 Then
					
						' Add by NAF, 11-10-2012
						if IsMandatory = "FALSE" then 
							if JOBCODE = "" then 
								JOBCODE = defaultJob
							end if	
						end if
						' End of Modification
						
						
						IF ( cdbl(mandays) = 0 ) and ( gangtype <> "SUPERVISION" ) and ( ATTENDANCE = "KJ" ) THEN 
							response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Gang Type is not a Supervision Gang and its a KJ attendance, the mandays should not be ZERO " & "</font>"
							cmdView	= "True"
						ELSE
						
							q = " select sum(mandays) mandays1" & _
								" from ATTENDANCE_GAD_TEMP" & _
								" where fdDate = to_DATE('" & fdDate & "','mm/dd/yyyy')" & _
								"       and EmployeeCode = '" & EmployeeCode & "'" & _
								"       and fcBA = '" & Session("fcBA") & "'"
							set rs = connStr.execute(q)
							if not rs.eof then
								IF NOT IsNull(rs("mandays1")) then
									TOTMANDAYS = cdbl(rs("mandays1"))
								end if
							end if
							rs.close()
							rs=null
							
							q = " select nvl(sum(mandays), 0) oldmandays" & _
								" from ATTENDANCE_GAD_TEMP" & _
								" where fcBA = '" & Session("fcBA") & "'" & _
								"       and linenokey = '" & LineNoKey & "'"
							'response.write "q : " & q & "<br/>"
							set rs = connStr.execute(q)
							if not rs.eof then
								OLDMANDAYS = cdbl(rs("oldmandays"))
							else
								OLDMANDAYS = 0
							end if
							rs.close()
							rs=null
							
							'response.write "Old Mandays : " & OLDMANDAYS & "<br/>"
							
							If (Cdbl(TOTMANDAYS) - Cdbl(OLDMANDAYS)) + Cdbl(MANDAYS) <= 1 Then
							
								'-- Get Data Entry Validate if KHL not allowed over20 Days, else free
								queryentryvalidate = " select nvl(ATTENDANCE_GAD_TEMP_API.ValidasiHKKHL('" & Session("fcBA") & "'" & _
													 "                                             , to_date('" & fddate & "', 'mm/dd/yyyy')" & _
													 "                                             , '" & EMPLOYEECODE & "'" & _
													 "                                             , '" & attendance & "'" & _
													 "                                             , " & Cdbl(MANDAYS) & "" & _
													 "                                             , " & LineNoKey & "), 'FALSE') entryvalidate" & _ 
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
								
								if tempentryvalidate = "TRUE" or tempentryvalidate = "EARLY_WARNING" then
								
									if tempentryvalidate = "EARLY_WARNING" then
										message_warning = "Data Tersimpan, tetapi Karyawan tersebut sudah mencapai >= 21 HK selama 3 Bulan berturut-turut, mohon di proses untuk perubahan status. Jika status masih belum berubah maka tidak dapat di input di bulan depan!" & "<br/>" & _
														  "Untuk detailnya silahkan tanya kepada manajemen anda, Terima Kasih!" & " <br/>"
										Response.Write "<div class = ""warning"">" & message_warning & "</div>"
									end if
								
									attInRecord = EditRecordDailyValidation(Session("fcBA"), fddate, EMPLOYEECODE, ATTENDANCE, MANDAYS, LineNoKey)
										
									'Response.Write attInRecord
										
									if attInRecord = "ENTRY" then
										strFieldDetail = " select SUM(overtime_hours) overtime_hours from ATTENDANCE_GAD_TEMP " & _
												         " where fcba =  '"& session("fcba") &"'" & _
												         " and employeecode = '"& employeecode &"'" & _
												         " and fddate = to_date('"& fddate &"','mm/dd/yyyy')" & _
												         " and linenokey = '"& linenokey &"'"				
														' Response.Write strFieldDetail

													Set rsQty = Server.CreateObject("ADODB.Recordset")
													rsQty.Open strFieldDetail, connstr							
													if not rsQty.Eof then 
														total_overtime_hours = cdbl(to_nol(rsQty("overtime_hours")))  
													End If
  
										' if   total_overtime_hours <= cdbl(to_nol(overtime_hours))  then  
										' 	'Validasi terkait data di fingerprint
										 	fp_att_cek = FP_CheckAttendance(EMPLOYEECODE, ATTENDANCE, fddate, 0, Session("fcBA"))
										' 	' response.write " Record = " & fp_att_cek & " | length = " & len(trim(fp_att_cek))
										' else
										' 	fp_att_cek = "Validation Passed"
										' end if

										 if cstr(""&fp_att_cek) = "Validation Passed" then
										
											u_output_harian = f_u_output_harian(Session("fcBA"), LineNoKey, ATTENDANCE, OUTPUT)
											u_output_rotasi = f_u_output_rotasi(Session("fcBA"), LineNoKey, ATTENDANCE, OUTPUT)
											
											if u_output_harian = "Success" then
												if u_output_rotasi = "Success" then

													' q = " Update Attendance_GAD" & _
													' 	" set SUPERVISION_1 = trim('" & SUPERVISION_1 & "')" & _
													' 	"     , Supervision_2 = trim('" & SUPERVISION_2 & "')" & _
													' 	"     , supervision_3 = trim('" & SUPERVISION_3 & "')" & _
													' 	"     , supervision_4 = trim('" & SUPERVISION_4 & "')" & _
													' 	"     , EMPLOYEECODE = trim('" & EMPLOYEECODE & "')" & _
													' 	"     , ATTENDANCE = trim('" & attendance & "')" & _
													' 	"     , Jobcode = trim('" & JOBCODE & "')" & _
													' 	"     , locationtype = trim('" & locationtype & "')" & _
													' 	"     , locationcode = trim('" & LOCATIONCODE & "')" & _
													' 	"     , mandays = trim('" & MANDAYS & "')" & _
													' 	"     , othrs = trim('" & OTHRS & "')" & _
													' 	"     , rate = trim('" & RATE & "')" & _
													' 	"     , unit = trim('" & UNIT & "')" & _
													' 	"     , output = trim('" & OUTPUT & "')" & _
													' 	"     , reference = '" & REFERENCE & "'" & _
													' 	"     , remarks = '" & REMARKS & "'" & _
													' 	"     , TYPE_OVERTIME = '" & TYPE_OVERTIME & "'" & _
													' 	"     , OVERTIME_HOURS = '" & OVERTIME_HOURS & "'" & _
													' 	"     , LASTUPDATE = sysdate " & _
													' 	"     , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
													' 	" where fcBA = '" & Session("fcBA") & "'" & _
													' 	"       and linenokey = '" & LineNoKey & "'"												

													' connStr.execute(q)
													Call UpdateDataCharge
													call UpdateAllowance
													' if cdbl(to_nol(Amount)) = 0 then
													' 	Err.Number = 1001
													' 	Err.Description = "Anda harus input angka lebih dari 0"
													' Else
													' q = " Update attendance_empallded" & _
													' 	" set ALLOWANCE_CODE = trim('" & ALLOWANCE_CODE & "')" & _
													' 	"     , AMOUNT = trim('" & AMOUNT & "')" & _
													' 	"     , LASTUPDATE = sysdate " & _
													' 	"     , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
													' 	" where fcBA = '" & Session("fcBA") & "'" & _
													' 	"       and linenokey = '" & LineNoKey & "'"												
													' response.Write(q)
													' connStr.execute(q)
													' End If
													strFieldDetail = " SELECT pr.parcode, pr.parname, prd.fccode, prd.parheadcode, prd.fcname" & _
																	 "	FROM parameter pr " & _
																	 "	INNER JOIN parameterdetail prd ON (pr.parcode = prd.parcode and pr.fcba = prd.fcba)" & _
																 	 "	WHERE pr.parcode = 'HARDCODE' AND pr.fcba = '"& session("fcBA") &"'"				
																 	 	' Response.Write strFieldDetail

													Set rsQty = Server.CreateObject("ADODB.Recordset")
													rsQty.Open strFieldDetail, connstr							
													if not rsQty.Eof then
													do while not rsQty.eof
															JobPanen = rsQty("fccode")
															rsQty.MoveNext
														loop
													End If

													If (len(trim(""&JobPanen))) <> 0 then
													q	=	" Delete from HarvestingQuality where" & _
															" EMPCODE = '" & EMPLOYEECODE & "'" & _
															" And FIELDCODE = '" & LOCATIONCODE & "'" & _
															" and fddate = to_Date('" & fdDate & "','MM/DD/YYYY')" & _
															" and fcBA = '" & Session("fcBA") & "'"
													
															' response.Write(q)
															connStr.execute(q)

													' If cdbl(to_nol(empcode)) > 0 Then
														if (cdbl(to_nol(UNDER_RIPE)) <> 0 OR cdbl(to_nol(OVER_RIPE)) <> 0 OR cdbl(to_nol(ABNORMAL)) <> 0 OR cdbl(to_nol(LONG_STALK)) <> 0 OR cdbl(to_nol(EATEN_BY_RAT)) <> 0 OR cdbl(to_nol(UNHARVEST_FFB)) <> 0 OR cdbl(to_nol(UNCOLLECT_LF_CIRCLE)) <> 0 OR cdbl(to_nol(UNPRUNE_FROND)) <> 0 OR cdbl(to_nol(QE_1)) <> 0 OR cdbl(to_nol(QE_2)) <> 0 OR cdbl(to_nol(QE_3)) <> 0 OR cdbl(to_nol(QE_5)) <> 0 OR cdbl(to_nol(QE_6)) <> 0 OR cdbl(to_nol(QE_7)) <> 0 )then
															
															' Call UpdateDataCharge
															' call UpdateAllowance
															
															q1	=	" Insert into HarvestingQuality (" & _
																	" EMPCODE" & _
																	" , FDDATE" & _
																	" , FIELDCODE" & _
																	" , UNDER_RIPE" & _
																	" , OVER_RIPE" & _
																	" , ABNORMAL" & _
																	" , LONG_STALK" & _
																	" , EATEN_BY_RAT" & _
																	" , UNHARVEST_FFB" & _
																	" , UNCOLLECT_LF_CIRCLE" & _
																	" , UNCOLLECT_LF_PIECE" & _
																	" , UNARRANGE_FFB" & _
																	" , UNPRUNE_FROND" & _
																	" , QE_1" & _
																	" , QE_2" & _
																	" , QE_3" & _
																	" , QE_4" & _
																	" , QE_5" & _
																	" , QE_6" & _
																	" , QE_7" & _
																	" , QE_8" & _
																	" , QE_9" & _
																	" , QE_10" & _
																	" , FCENTRY" & _
																	" , FCEDIT" & _
																	" , FCIP" & _
																	" , FCBA" & _
																	" , LASTUPDATE" & _
																	" , LASTTIME" & _
																	" ) " & _
																	" values ("& _
																	" '" & EMPLOYEECODE & "'," & _
																	" to_date('" & fddate & "','mm/dd/yyyy')" & _
																	" , '" & LOCATIONCODE & "'" & _
																	" , '" & UNDER_RIPE & "'" & _
																	" , '" & OVER_RIPE & "'" & _
																	" , '" & ABNORMAL & "'" & _
																	" , '" & LONG_STALK & "'" & _
																	" , '" & EATEN_BY_RAT & "'" & _
																	" , '" & UNHARVEST_FFB & "'" & _
																	" , '" & UNCOLLECT_LF_CIRCLE & "'" & _
																	" , '" & UNCOLLECT_LF_PIECE & "'" & _
																	" , '" & UNARRANGE_FFB & "'" & _
																	" , '" & UNPRUNE_FROND & "'" & _
																	" , '" & QE_1 & "'" & _
																	" , '" & QE_2 & "'" & _
																	" , '" & QE_3 & "'" & _
																	" , '" & QE_4 & "'" & _
																	" , '" & QE_5 & "'" & _
																	" , '" & QE_6 & "'" & _
																	" , '" & QE_7 & "'" & _
																	" , '" & QE_8 & "'" & _
																	" , '" & QE_9 & "'" & _
																	" , '" & QE_10 & "'" & _
																	" , '" & Session("CodeName") & "'" & _
																	" , '" & Session("CodeName") & "'" & _
																	" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
																	" , '" & Session("fcBA") & "'" & _
																	" , sysdate" & _
																	" , To_Char(sysdate,'HH24:MI')" & _
																	" )"
																	
																' response.Write(q1)
																' response.Write(jobcode)
																response.Flush()
																connStr.execute(q1)
															' end if
														end if
													else

													end if
												else
													err.number = 1001
													err.description = u_output_rotasi & " <br/>"
												end if
											else
												err.number = 1001
												err.description = u_output_harian & " <br/>"
											end if
										else
										    'No FInger
										 	err.number = 1001
										 	err.description = fp_att_cek & " <br/>"
										end if
									else
										err.number = 1001
										err.description = "Anda sudah memasukkan data kehadiran/ Maksimal 1 HK!!!" & "<br/>" & _
														  "Silahkan cek kembali data yang anda ingin/sudah di masukkan ke dalam sistem!" & " <br/>"
									end if
								else
									if tempentryvalidate = "NO_INPUT" then
										err.number = 1001
										err.description = "Karyawan tersebut belum di proses perubahan status-nya sehingga hk tidak bisa di input!" & "<br/>" & _
														  "Untuk detailnya silahkan tanya kepada manajemen anda, Terima Kasih!" & " <br/>"
									else
										err.number = 1001
										err.description = "Maksimal HK yang di perbolehkan untuk KHL adalah 20 HK, Cek kembali inputan anda!" & "<br/>" & _
														  "Untuk detailnya silahkan tanya kepada manajemen anda, Terima Kasih!" & " <br/>"
									end if
								end if
							else
								err.Number = 1001
								err.description = "Total mandays perhari 1, Cek kembali mandays yang anda input!" & "<br/>" & _
												  "sisa mandays yg bisa diedit adalah " & 1 - (Cdbl(TOTMANDAYS) - Cdbl(OLDMANDAYS)) & " <br/>"
							end if

							 IF ( cdbl(TOTMANDAYS) ) <= 1 then    'Add Akbar 4 des 2012 			
								connStr.execute("commit")
								cmdView	= "True"
							 ELSE
								'Add Akbar 4 des 2012 begin
								connStr.execute("rollback")
								response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Total mandays perhari 1 ." & "</font>"
								response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Total mandays yang sudah di input tgl "& fdDate &" adalah " & cdbl(TOTMANDAYS) & " sisa mandays yg bisa(edit) di input adalah " &  1 -  (cdbl(TOTMANDAYS) - cdbl(OLDMANDAYS)) & "</font>"
								cmdView	= "True"
								'Add Akbar 4 des 2012 end
							 END IF	
						END IF
					End If
				else
					err.number = 1001
					err.description = "Tanggal """ & fdDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
			else
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Unable to Update/Delete/Insert," & "</font>"
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Please OPEN your payroll Status </font>"
				cmdView	= "True"
			end if

			elseif Request("txtJenis")= "Confirmed" then
			   connstr.BeginTrans
			   strQueryStatus = "BEGIN ATTENDANCEGAD_TEMP_API.NextState('" & Session("fcBA") & "',to_date('" & fdDate & "','mm/dd/yyyy'),'"& gangcode &"','Confirm' ); END; "
			   connStr.execute(strQueryStatus)

			   if err.number  = 0 then
					  connStr.CommitTrans
			   end if       
			   
		       cmdView	= "True"

		    elseif Request("txtJenis")= "RePlanned" then
			   connstr.BeginTrans   
			   strQueryStatus = "BEGIN ATTENDANCEGAD_TEMP_API.NextState('" & Session("fcBA") & "',to_date('" & fdDate & "','mm/dd/yyyy'),'"& gangcode &"', 'Replan' ); END; "
			   ' response.write strQueryStatus
			   connStr.execute(strQueryStatus)	
			   
			   strQueryStatus2 = "DELETE FROM  ATTENDANCE_GAD_HIST WHERE FCBA='" & Session("fcBA") & "' AND FDDATE= to_date('" & fdDate & "','mm/dd/yyyy') AND GANGCODE='"& gangcode &"'"
			  '  response.write strQueryStatus2
			   connStr.execute(strQueryStatus2)	
			   
		      
			   if err.number  = 0 then
					connStr.CommitTrans
			   end if      	  
               
			   

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
				
			   q = 	" INSERT INTO attendance_gad_hist ( fdDate" & _
					"                                           , gangcode" & _
					"                                           , currentauthorizelevel" & _
					"                                           , fcentry" & _
					"                                           , fcedit" & _
					"                                           , fcip" & _
					"                                           , fcba" & _
					"                                           , lastupdate" & _
					"                                           , lasttime)"& _
					" SELECT fddate" & _
					"        , gangcode" & _
					"        , '" & authorizelevelcode & "' currentauthorizelevel" & _
					"        , '" & Session("CodeName") & "' fcentry" & _
					"        , '" & Session("CodeName") & "' fcedit" & _
					"        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
					"        , fcba" & _
					"        , SYSDATE lastupdate" & _
					"        , TO_CHAR(SYSDATE, 'hh24:mi') lasttime" & _
					" FROM ATTENDANCE_GAD_TEMP" & _
					" WHERE fdDate = to_date('" & fdDate & "','mm/dd/yyyy')" & _
					"       AND fcba = '"&  session("fcba") & "'" & _
					"		AND GangCode = '"& gangcode &"'"
				' response.write q
				connStr.execute(q)
			   strQueryStatus = "BEGIN ATTENDANCEGAD_TEMP_API.NextState('" & Session("fcBA") & "',to_date('" & fdDate & "','mm/dd/yyyy'),'"& gangcode &"','StartAuthorize' ); END; "
			   'response.write strQueryStatus
			   connStr.execute(strQueryStatus)	
		      
			   	if err.number  = 0 then
					connStr.CommitTrans
			  	end if
			   
			cmdView	= "True"
		else
			cmdView	= "True"
		end if								 
		'Response.Write "sql " & q
	end if
else
    cmdView	= request("cmdView")
	done = 1
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Request("fddate")<>"" then
   fddate   =Request("fddate")
else
   fddate   =textDate(date)
end if
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
GangCode=request("GangCode")
txtJenis = request("txtJenis")
if done = 0 then
   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Anda Harus melakukan setup di master Parameter E0102 untuk Default Mandatory JOB Activity " & "</font>" 
end if 
' if err.number <> 0 then
   ' response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Error Number : " & err.number & " because : " & err.description & "</font>"
' End if

if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

 '-------------ISI Payment Form

q = "SELECT a.linenokey," & _
	"       a.gangcode," & _
	"       a.document_classification," & _
	"       to_char(a.fddate,'mm/dd/yyyy') fddate," & _
	"       a.supervision_1," & _
	"       a.supervision_2," & _
	"       a.supervision_3," & _
	"       a.supervision_4," & _
	"       a.employeecode," & _
	"       c1.fcname spvname1," & _
	"       c2.fcname spvname2," & _
	"       c3.fcname spvname3," & _
	"       c4.fcname spvname4," & _
	"       c.fcname empname," & _
	"       gang_api.get_remarks(a.gangcode,a.fcba) sectionname," & _
	"       a.attendance," & _
	"       a.jobcode," & _
	"       a.locationtype," & _
	"       a.locationcode," & _
	"       a.chargeunit," & _
	"       a.chargejob," & _
	"       a.chargetype," & _
	"       a.chargecode," & _
	"       a.mandays," & _
	"       a.othrs," & _
	"       a.rate," & _
	"       a.unit," & _
	"       a.output," & _
	"       a.reference," & _
	"       a.remarks," & _
	"       a.BUCKET," & _
	"       a.KG_BRONDOLAN," & _
	"       a.kg_janjang," & _
	"       a.spbno," & _
	"       a.type_overtime," & _
	"       a.overtime_hours," & _
	"       a.fcentry," & _
	"       a.fcedit," & _
	"       a.fcip," & _
	"       a.fcba," & _
	"       a.RowState," & _
	"       b.allowance_code," & _
	"       b.amount," & _
	"       d.under_ripe," & _
	"       d.long_stalk," & _
	"       d.qe_6," & _
	"       d.qe_7," & _
	"       d.over_ripe," & _
	"       d.abnormal," & _
	"       d.qe_3," & _
	"       d.unharvest_ffb," & _
	"       d.uncollect_lf_circle," & _
	"       d.unprune_frond," & _
	"       d.qe_1," & _
	"       d.qe_2," & _
	"       d.eaten_by_rat," & _
	"       d.qe_5," & _
	"       a.bkmcode," & _
	"       a.abw," & _
	"       a.bunch_weight" & _
	"  FROM ATTENDANCE_GAD_TEMP a" & _
	"  LEFT JOIN attendance_empallded b ON (a.linenokey = b.linenokey and a.fcba = b.fcba)" & _
	"  LEFT JOIN employee c ON (a.fcba = c.fcba and a.employeecode = c.fccode)" & _
	"  LEFT JOIN employee c1 ON (a.fcba = c1.fcba and a.SUPERVISION_1 = c1.fccode)" & _
	"  LEFT JOIN employee c2 ON (a.fcba = c2.fcba and a.SUPERVISION_2 = c2.fccode)" & _
	"  LEFT JOIN employee c3 ON (a.fcba = c3.fcba and a.SUPERVISION_3 = c3.fccode)" & _
	"  LEFT JOIN employee c4 ON (a.fcba = c4.fcba and a.SUPERVISION_4 = c4.fccode)" & _
	"  LEFT JOIN harvestingquality d ON (a.fcba = d.fcba and a.employeecode = d.empcode and a.fddate = d.fdDate and a.locationcode = d.fieldcode)" & _
	" where a.gangcode = '" & gangcode & "'" & _
    "       and a.fddate = to_date('" & fdDate & "', 'mm/dd/yyyy')" & _
    "       and a.fcba = '" & session("fcBA") & "' ORDER BY a.LASTUPDATE"
' response.write q
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
	fdDate=rsVoucher("fdDate")
	gangcode=rsVoucher("gangcode")
	sectionname=rsVoucher("sectionname")
	selectedgangcode=rsVoucher("gangcode")
	rowstate = rsVoucher("rowstate")
	document_classification = rsVoucher("document_classification")
	bkmcode = rsVoucher("bkmcode")
	'SudahClose= "Y"
	' response.write selectedgangcode
End if

set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'AA'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _ 
	"       AND '"& session("CodeName") &"' = '" & session("CodeName") & "'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & session("fcba") & "'"
' response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	can_startauthorize =  cdbl(to_nol(rsFrez("total")))
else
	can_startauthorize = 0
End if
rsFrez.Close : set rsFrez = Nothing 

q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'AA'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'FALSE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _ 
	"       AND alu.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & session("fcba") & "'"

set ArsFrez = server.CreateObject("Adodb.recordset")
ArsFrez.Open q, connstr
if Not ArsFrez.Eof then
	can_restartauthorize =  cdbl(to_nol(ArsFrez("total")))
else
	can_restartauthorize = 0
End if
ArsFrez.Close : set ArsFrez = Nothing 


Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & fdDate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & fdDate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Employee')"
			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
'response.Write oSQLQuery
oRecordSet.Close
Set oRecordSet = Nothing

%>
<form name="frmEntry" method="post" action="Attendance_GAD.asp">
	<table class="data" style="width: 100%">
		<tr>
			<td align="left" style="width: 10%">Date</td>
			<td align="left" >
				<input class="datepicker" type="text" maxlength="10" id="fddate" name="fdDate" size="10" value="<%=fdDate%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left">
				Code : <input type="text" maxlength="50" name="bkmcode" size="30" value="<%=bkmcode%>" style="background-color:#eee;" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Section Name</td>
			<td align="left">
				<select name="sectionname" value="<%=sectionname%>" id="sectionname" onchange="javascript:getGangcode();" class="chosen-select" style="width:70%">
						<option value=""></option> 
							<% 	
								q = " SELECT DISTINCT emp.sectionname, sc.fcname" & _
									"    FROM employee emp" & _
									"         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
									"   WHERE fcba = '"& session("fcba") &"' ORDER BY emp.sectionname"
								' response.write q
								Set rsComboC = Server.CreateObject("ADODB.Recordset")
								rsComboC.CursorType = 3
								rsComboC.CursorLocation = 3
								rsComboC.Open q, connstr
								if not rsComboC.eof then%>
								<%
								Do While Not rsComboC.eof
								%>
									<option value="<%=rsComboC("sectionname")%>" 
										<%if request("sectionname") = cstr(""&rsComboC("sectionname")) then %> selected 
										<%End If%>><%=rsComboC("sectionname")%>
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
		</tr>
		<tr>
			<td  align="left">Gang</td>
			<td align="left">
				<div id="gangcodelist">
				<font color="#FF0000">*</font>
			</td>
				<input type="Hidden" id="selectedgangcode_" name="selectedgangcode" value="<%=request("selectedgangcode")%>">
		</tr>  
		<%If SudahClose = "Y" Then%>
			<tr>
				<td align="left">Classification</td>
				<td align="left">
					<input type="text" name="document_classification" value="<%=document_classification%>" readonly style="background-color: #eee;">
				</td>
			</tr>
		<%Else%>
			<tr>
				<td align="left">Classification</td>
				<td align="left">
					<select name="document_classification" value="<%=document_classification%>" id="document_classification" class="chosen-select" style="width:63%">
							<!--<option value=""></option> -->
								<% 	
									q = " SELECT fccode, fcname FROM document_classification where document_type= 'AA' and isactive='TRUE' ORDER BY fccode"
									' response.write q
									Set rsComboC = Server.CreateObject("ADODB.Recordset")
									rsComboC.CursorType = 3
									rsComboC.CursorLocation = 3
									rsComboC.Open q, connstr
									if not rsComboC.eof then%>
									<%
									Do While Not rsComboC.eof
									%>
										<option value="<%=rsComboC("fccode")%>" 
											<%if request("document_classification") = cstr(""&rsComboC("fccode")) then %> selected 
											<%End If%>><%=rsComboC("fcname")%>
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
			</tr>
		<%End If%>
		<tr>
			<td align="left">
				Status
			</td>
			<td align="left">
				<input type="text" maxlength="50" name="rowstate" size="30" value="<%=rowstate%>" style="background-color:#eee;" readonly>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="PrintGAD.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fdDate%>&GangCode=<%=GangCode%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
				<%
				dim dataModule,dataSubmodule,dataDetail,hiddenHadir
				dataModule 		= "Employee"
				dataSubmodule 	= "Attendance and Activity"
				if GangCode = "" then
					dataDetail 	= ""
					hiddenHadir = "visibility: hidden;"
				else
					dataDetail 	= GangCode
					hiddenHadir = "visibility: visible;"
				
				end if
				%>
				<a class="jqplot-image-button" style="<%=hiddenHadir%>" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
				<% if RowState = "Planned" then %>
				    <a href="?fdDate=<%=fddate%>&gangcode=<%=gangcode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to confirm this data :  <%=fddate%>,<%=gangcode%> ?');"> Confirmed </a>  
				<% Elseif RowState = "Confirmed" then %> 
					<a href="?fdDate=<%=fddate%>&gangcode=<%=gangcode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to replan this data :  <%=fddate%>,<%=gangcode%> ?');"> Re-Planned </a>
					&nbsp;&nbsp;
						<%if can_startauthorize = 0 then%>
							<font color="#FF0000">You Have No Permission to Start Authorize</font>
						<%else%>
							<a href="?fdDate=<%=fddate%>&gangcode=<%=gangcode%>&txtJenis=StartAuthorize" onClick="return confirm('Are you sure to start authorize this data :  <%=fcCode%> ?');">Start Authorize</a>
						<% end if%>
				<% Elseif RowState = "AuthorizeOnProgress" then %> 
						<%if can_restartauthorize <> 0 then%>
							<a href="?fdDate=<%=fddate%>&gangcode=<%=gangcode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to Re-Planned this data :  <%=fcCode%> ?');">Re-Planned</a>
						<% end if%>
				
				<% end if%>
				
			</td>
		</tr>
	</table>  
</form>
<%
'response.write request("document_classification")&"ssss"
'if cmdView = "True" and len(trim(GangCode)) > 0 and len(trim(document_classification)) > 0 Then
if cmdView = "True" and len(trim(GangCode)) > 0 and len(trim(request("document_classification"))) > 0 Then
	Dim rsGAD,intRecordsPerPage,cColor
	set rsGAD = server.CreateObject("Adodb.recordset")
	q = "SELECT a.linenokey," & _
		"       a.gangcode," & _
		"       a.fddate," & _
		"       a.supervision_1," & _
		"       a.supervision_2," & _
		"       a.supervision_3," & _
		"       a.supervision_4," & _
		"       a.employeecode," & _
		"       c1.fcname spvname1," & _
		"       c2.fcname spvname2," & _
		"       c3.fcname spvname3," & _
		"       c4.fcname spvname4," & _
		"       c.fcname empname," & _
		"       c.sectionname," & _
		"       a.attendance," & _
		"       a.jobcode," & _
		"       a.locationtype," & _
		"       a.locationcode," & _
		"       a.chargeunit," & _
		"       a.chargejob," & _
		"       a.chargetype," & _
		"       a.chargecode," & _
		"       a.mandays," & _
		"       a.othrs," & _
		"       a.rate," & _
		"       a.unit," & _
		"       a.output," & _
		"       a.reference," & _
		"       a.remarks," & _
		"       a.BUCKET," & _
		"       a.KG_BRONDOLAN," & _
		"       a.kg_janjang," & _
		"       a.spbno," & _
		"       a.type_overtime," & _
		"       a.overtime_hours," & _
		"       a.fcentry," & _
		"       a.fcedit," & _
		"       a.fcip," & _
		"       a.fcba," & _
		"       a.RowState," & _
		"       b.allowance_code," & _
		"       b.amount," & _
		"       d.under_ripe," & _
		"       d.long_stalk," & _
		"       d.qe_6," & _
		"       d.qe_7," & _
		"       d.over_ripe," & _
		"       d.abnormal," & _
		"       d.qe_3," & _
		"       d.unharvest_ffb," & _
		"       d.uncollect_lf_circle," & _
		"       d.unprune_frond," & _
		"       d.qe_1," & _
		"       d.qe_2," & _
		"       d.eaten_by_rat," & _
		"       d.qe_5," & _
		"       a.jobcode||' '||job_api.get_jobcode_name(a.jobcode,a.fcba) JobDet," & _
		"       a.chargejob||' '||job_api.get_jobcode_name(a.chargejob,a.fcba) CJobDet," & _
		"       a.abw," & _
		"       a.bunch_weight " & _
		"  FROM ATTENDANCE_GAD_TEMP a" & _
		"  LEFT JOIN attendance_empallded b ON (a.linenokey = b.linenokey and a.fcba = b.fcba)" & _
		"  LEFT JOIN employee c ON (a.fcba = c.fcba and a.employeecode = c.fccode)" & _
		"  LEFT JOIN employee c1 ON (a.fcba = c1.fcba and a.SUPERVISION_1 = c1.fccode)" & _
		"  LEFT JOIN employee c2 ON (a.fcba = c2.fcba and a.SUPERVISION_2 = c2.fccode)" & _
		"  LEFT JOIN employee c3 ON (a.fcba = c3.fcba and a.SUPERVISION_3 = c3.fccode)" & _
		"  LEFT JOIN employee c4 ON (a.fcba = c4.fcba and a.SUPERVISION_4 = c4.fccode)" & _
		"  LEFT JOIN harvestingquality d ON (a.fcba = d.fcba and a.employeecode = d.empcode and a.fddate = d.fdDate and a.locationcode = d.fieldcode)" & _
		" where a.gangcode = '" & gangcode & "'" & _
		"       and a.fddate = to_date('" & fdDate & "', 'mm/dd/yyyy')" & _
        "       and a.fcba = '" & session("fcBA") & "' ORDER BY a.LASTUPDATE"
	' response.Write q	
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.CursorType = 3
	rsGAD.CursorLocation = 3
	rsGAD.Open q, connstr

	' if txtJenis="Edit" then
	' 	q = "SELECT * FROM HarvestingQuality" & _
	' 	" where fddate = to_date('" & fdDate & "', 'mm/dd/yyyy')" & _
 '        "       and fcba = '" & session("fcBA") & "'"
	' response.Write q	
	' Set rsQ = Server.CreateObject("ADODB.Recordset")
	' rsQ.CursorType = 3
	' rsQ.CursorLocation = 3
	' rsQ.Open q, connstr
 
%>
   
   <form action="Attendance_GAD.Asp" method="post" name="eForm" >
   
	<%
	' Start Modification to Pick Validation whether User have Mandays or Not. by NAF, 24-4-2013
	' its depended on GANGTYPE and ATTENDANCE ==> Should be KJ
	' Jika Total Dia bekerja bukan sebagai Gang Mandor dan KJ, maka Mandays Harus Lebih besar dari Nol
	
	q = " Select gangtypecode from Gang " & _
		" where " & _
		" fccode = trim('" & GangCode & "') and " & _
		" fcba = trim('" & Session("fcBA") & "')"
		
	set rs = connStr.execute(q)
	if not rs.eof then
		GangType = rs("gangtypecode")
	End if
	rs.close()
	rs=null 
	%>
   
   
    <%
		tempStringJob = "select JOB_API.GET_JOBCODE_NAME( '" &  JOBCODE & "','" & Session("fcBA") & "') JobName from dual"
		set rsStringJob = connStr.execute(tempStringJob)
		if not rsStringJob.eof then
			JobCodeDescription = rsStringJob("JobName")
		else
			JobCodeDescription = "..."
		End if
		rsStringJob.close()
		rsStringJob=null 	
		
		tempStringMandays = "select PARAMETER_COMB_UOM_API.GetUom_Mandays('ATTENDANCE_GAD','" &  JOBCODE & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') MandaysName from dual"
		set rsStringMandays = connStr.execute(tempStringMandays)
		if not rsStringMandays.eof then
			MandaysDescription = rsStringMandays("MandaysName")
		else
			MandaysDescription = "..."
		End if
		rsStringMandays.close()
		rsStringMandays=null 

		tempStringUnit = "select PARAMETER_COMB_UOM_API.GetUom_Unit('ATTENDANCE_GAD','" &  JOBCODE & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') UnitName from dual"
		set rsStringUnit = connStr.execute(tempStringUnit)
		if not rsStringUnit.eof then
			UnitDescription = rsStringUnit("UnitName")
		else
			UnitDescription = "..."
		End if
		rsStringUnit.close()
		rsStringUnit=null
		
		tempStringOutput = "select PARAMETER_COMB_UOM_API.GetUom_Output('ATTENDANCE_GAD','" &  JOBCODE & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') OutputName from dual"
		set rsStringOutput = connStr.execute(tempStringOutput)
		if not rsStringOutput.eof then
			OutputDescription = rsStringOutput("OutputName")
		else
			OutputDescription = "..."
		End if
		rsStringOutput.close()
		rsStringOutput=null
		
		tempStringOth = "select PARAMETER_COMB_UOM_API.GetUom_Othrs('ATTENDANCE_GAD','" &  JOBCODE & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'" & Session("fcBA") & "') OthName from dual"
		set rsStringOth = connStr.execute(tempStringOth)
		if not rsStringOth.eof then
			OthDescription = rsStringOth("OthName")
		else
			OthDescription = "..."
		End if
		rsStringOth.close()
		rsStringOth=null
	%>
   
   
	<table cellpadding="0" cellspacing="0" width="100%" id="myForm">
		<tr style="display: block; padding-left: 10px">
			<td align="right">
				<font face="Verdana" size="1">SUPERVISION 1</font>
			</td>
			<td align="left">
				<input id="s1" class="" type="text" size="15" name="SUPERVISION_1" value="<%=SUPERVISION_1%>" onblur="validate_(this.value)" onkeypress="return harusangka(event);">
				<span id="nonscript1">&nbsp;&nbsp;&nbsp;</span>
				<a id ="nonscript1_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Employee</font></a>
		</tr>
		<tr style="display: block; padding-left: 10px">
			<td align="right">
				<font face="Verdana" size="1">SUPERVISION 2</font>
			</td>
			<td align="left">
				<input id="s2" type="text" size="15" name="SUPERVISION_2" value="<%=SUPERVISION_2%>" onblur="validate_2(this.value)" onkeypress="return harusangka(event);">
				<span id="nonscript2">&nbsp;&nbsp;&nbsp;</span>
				<a id ="nonscript2_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Employee</font></a>
			</td>
		</tr>
		<tr style="display: block; padding-left: 10px">
			<td align="right">
				<font face="Verdana" size="1">SUPERVISION 3</font>
			</td>
			<td align="left">
				<input id="s3" type="text" size="15" name="SUPERVISION_3" value="<%=SUPERVISION_3%>" onblur="validate_3(this.value)" onkeypress="return harusangka(event);">
				<span id="nonscript3">&nbsp;&nbsp;&nbsp;</span>
				<a id ="nonscript3_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Employee</font></a></td>
		</tr>
		<tr style="display: block; padding-left: 10px">
			<td align="right">
				<font face="Verdana" size="1">SUPERVISION 4</font>
			</td>
			<td align="left">
				<input id="s4" type="text" size="15" name="SUPERVISION_4" value="<%=SUPERVISION_4%>" onblur="validate_4(this.value)" onkeypress="return harusangka(event);">
				<span id="nonscript4" >&nbsp;&nbsp;&nbsp;</span>
				<a id ="nonscript4_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Employee</font></a></td>
		</tr>
		<%IF txtJenis="Edit" Then%>
			<tr style="display: block;">
				<td align="right">
					<font face="Verdana"   size="1">EMPLOYEECODE</font>
				</td>
				<td align="left">
					<input id="empcd" type="text" size="15" name="EMPLOYEECODE" value="<%=EMPLOYEECODE%>" onblur="showtimefinger(this.value);validate_empcd(this.value);" readonly="True" style="background:#999999" onkeypress="return harusangka(event);">
					<span id="nonscript_empcd">&nbsp;*&nbsp;</span>
				</td>
			</tr>
			<input type="hidden" name="LineNoKey" value="<%=LineNoKey%>">
			<tr style="display: block; padding-left:15px">
				<td align="right">
					<font face="Verdana"   size="1">ATTENDANCE</font>
				</td>
				
			 
			<%
				Dim rsAttCode, qAttCode
				
				qAttCode = "select FCCODE, FCNAME from  PARAMETER  where PARCODE= 'E0102'  and fcba='" & Session("fcBA") & "'  order by FCNAME ASC"
				
				Set rsAttCode = Server.CreateObject("ADODB.Recordset")
				rsAttCode.CursorType = 3
				rsAttCode.CursorLocation = 3
				rsAttCode.Open qAttCode, connstr
				if not rsAttCode.eof then
			%>
				<td align="left">
					<select name="ATTENDANCE">
					<%Do While Not rsAttCode.eof%>
					<%if rsAttCode("FCCODE")=request("ATTENDANCE") Then%>			
						<option value="<%=rsAttCode("FCCODE")%>" selected><%=rsAttCode("FCCODE")%> : <%=rsAttCode("FCNAME")%></option> 
					<%Else%>
						<option value="<%=rsAttCode("FCCODE")%>"><%=rsAttCode("FCCODE")%> : <%=rsAttCode("FCNAME")%></option> 
					<%End if%>
			<%
						rsAttCode.movenext
					loop
				End if
				rsAttCode.Close
				set rsAttCode = Nothing
				qAttCode = ""
			%>
					</select>
					<font color="#FF0000">*</font><br>
					<font face="Verdana"   size="1">Finger:&nbsp;
					<span class="satuan_item" id="Att_in"><%= Att_in %></span>&nbsp;-&nbsp;<span class="satuan_item" id="Att_out"><%= Att_out %></span>
					</font>
				</td>
			</tr>
			 
			<tr style="display: block; padding-left:30px" id="job">
				<td align="right">
					<font face="Verdana"   size="1">JOBCODE</font>
				</td>
				<td align="left">
					<input onkeyup="showHint(this.value)" type="text" size="15" id="jobcode" name="JOBCODE" value="<%=JOBCODE%>"   onkeypress="return harusangka(event);" onchange="javascript:getHarvestingQualityData();javascript:getJobcodeIntraco();javascript:getJobcodeSpbHarvesting();" onblur="javascript:getHarvestingQualityData();javascript:getJobcodeIntraco(); javascript:getJobcodeSpbHarvesting();">
				    <font color="#FF0000">&nbsp;*&nbsp;</font>
<!--					<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a> -->
					<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=BKM&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>
					<span class="satuan_item" id="jobname">&nbsp;<%= JobCodeDescription %></span>
				</td>
			</tr>
			<tr style="display: block; padding-left:22px">
				<td align="right">
					<font face="Verdana"   size="1">TARGET TYPE</font>
				</td>
				<td align="left">
					<input type="text" size="15" name="LOCATIONTYPE" value="<%=locationtype%>" >
					<font color="#FF0000">&nbsp;*&nbsp;</font>
					<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a>
				</td>
			</tr>
			<tr style="display: block; padding-left:22px">
				<td align="right">
					<font face="Verdana"   size="1">TARGET CODE</font>
				</td>
				<td align="left">
					<input type="text" size="15" id="targetcode" name="LOCATIONCODE" value="<%=LOCATIONCODE%>" >
					<font color="#FF0000">&nbsp;*&nbsp;</font>
					<a href="javascript:PopupTType('<%=session("fcBA") %>')">
                            <font face="Verdana" size="1">Location Code</font></a>
				</td>
			</tr>
			<tr id="jobcodeintraco">
			
			<tr style="display: block; padding-left:22px" id="chargeunit">
				<td align="right">
					<font face="Verdana" size="1">CHARGE UNIT</font>
				</td>
                <td align="left">
                    <select  name="chargeunit" id = "chargeunit_" >
							<option value=""></option>
							<%
							qlov = 	 " select fccode,fcname " & _
									 "	from businessunit " & _
									 "	where fccompanycode = businessunit_api.get_CompanyCode('" & session("fcba") & "') " & _
									 "	and businessunit_api.Isactive(fccode)='TRUE' " & _ 
									 "  and fccode <> '" & session("fcba") & "'"
									 
							Set rslov = Server.CreateObject("ADODB.Recordset")
							rslov.Open qlov, connstr 
							if Not rslov.eof then
								While Not rslov.EOF
							%>
							<option value="<%=rslov("fccode")%>" <%if chargeunit = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%> </option>
							<%
									rslov.MoveNext
								Wend
							end if
							%>
					</select>
					
                </td>
            </tr>
			
			<tr style="display: block; padding-left:25px" id="chargejob">
				<td align="right">
					<font face="Verdana" size="1">CHARGE JOB</font>
				</td>
                <td align="left">
                    <input type="text" size="11" name="chargejob" id="chargejob_" onchange="javascript:getChargetype();" onkeypress="return harusangka(event);" value="<%=chargejob%>">&nbsp;<font color="#FF0000">*</font>
                   <!-- <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">-->
   				    <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=BKM&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>

                </td>
            </tr>
            <tr style="display: block; padding-left:22px" id="chargetype">
				<td align="right">
					<font face="Verdana"   size="1">CHARGE TYPE</font>
				</td>
                <td align="left">
                    <input type="text" size="11" name="chargetype" id="chargetype_" value="<%=chargetype%>">&nbsp;<font
                        color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                            <font face="Verdana" size="1">Location Type</font></a>
                </td>
            </tr>
            <tr style="display: block; padding-left:22px" id="chargecode">
				<td align="right">
					<font face="Verdana"   size="1">CHARGE CODE</font>
				</td>
                <td align="left">
                    <input type="text" size="11"  name="chargecode" value="<%=chargecode%>">&nbsp;<font
                        color="#FF0000">*</font>
						<a href="javascript:PopupTType('')">
                            <font face="Verdana" size="1">Location Code</font></a></font>
                </td>
            </tr>
            </tr>
			<% IF gangtype = "SUPERVISION"  THEN %>
			<tr style="display: block; padding-left:30px">
				<td align="right">
					<font face="Verdana" size="1">MANDAYS</font>
				</td>
				<td align="left">
					<input type="text" size="2" name="MANDAYS" value="<%=MANDAYS%>">
					<span class="satuan_item" id="mandays" onkeypress="return harusangka(event);">&nbsp;<%= MandaysDescription %></span>
				</td>
			</tr>
			<%else%>
			<tr style="display: block; padding-left:30px">
				<td align="right">
					<font face="Verdana"   size="1">MANDAYS</font>
				</td>
				<td align="left">
					<input type="text" size="2" name="MANDAYS" value="<%=MANDAYS%>">
					<span class="satuan_item" id="mandays" onkeypress="return harusangka(event);">&nbsp;<%= MandaysDescription %></span>
				</td>
			</tr>
			<%end if%>
			 
		<%else%>
			<tr style="display: block; padding-left: 10px">
				<td align="right">
					<font face="Verdana" size="1">EMPLOYEECODE</font>
				</td>
				<td align="left">
					<input id="empcd" type="text" size="15" name="EMPLOYEECODE" value="<%=EMPLOYEECODE%>" onblur="showtimefinger(this.value);validate_empcd(this.value);" onkeypress="return harusangka(event);">
					<span id="nonscript_empcd" style="color:red;">&nbsp;*&nbsp;</span>
					<a id ="nonscriptempcd_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
					<font face="Verdana"   size="1">Employee</font></a>
				</td>	
				<!--
				<td>
							<select name="txtempcode" class="chosen-select" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select fccode, fcname from employee where fcba = '"& Session("fcBA") &"' order by fcCode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%> </option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
				</td>
				-->
			</tr>
			<tr style="display: block; padding-left:23px ">
				<td align="right">
					<font face="Verdana"   size="1">ATTENDANCE</font>
				</td>
			<%
				Dim rsAttCodeA, qAttCodeA
				qAttCodeA = "select FCCODE, FCNAME from  PARAMETER  where PARCODE= 'E0102'  and fcba='" & Session("fcBA") & "'  order by FCNAME ASC"
				
				Set rsAttCodeA = Server.CreateObject("ADODB.Recordset")
				rsAttCodeA.CursorType = 3
				rsAttCodeA.CursorLocation = 3
				rsAttCodeA.Open qAttCodeA, connstr
				if not rsAttCodeA.eof then
			%>
				<td align="left">
					<select name="ATTENDANCE">
						<option value=""></option>
					<%Do While Not rsAttCodeA.eof%>
						<%if rsAttCodeA("FCCODE")=request("ATTENDANCE") Then%>			
						<option value="<%=rsAttCodeA("FCCODE")%>" selected><%=rsAttCodeA("FCCODE")%> : <%=rsAttCodeA("FCNAME")%></option> 
						<%Else%>
						<option value="<%=rsAttCodeA("FCCODE")%>"><%=rsAttCodeA("FCCODE")%> : <%=rsAttCodeA("FCNAME")%></option> 
						<%End if%>
			<%			rsAttCodeA.movenext
					loop
				End if
				rsAttCodeA.Close
				set rsAttCodeA = Nothing
				qAttCodeA = ""
			%>
					</select>
					<font color="#FF0000">*</font></br>
					<font face="Verdana"   size="1">Finger:&nbsp;
					<span class="satuan_item" id="Att_in"><%= Att_in %></span>&nbsp;-&nbsp;<span class="satuan_item" id="Att_out"><%= Att_out %></span>
					</font>
				</td>
			</tr>
			<tr style="display: block; padding-left:42px"  id="job">
				<td align="right">
					<font face="Verdana"   size="1">JOBCODE</font>
				</td>
				<td align="left">
					<input onkeyup="showHint(this.value)" type="text" size="15" id="jobcode" onchange="javascript:getTragettype();" name="JOBCODE" value="<%=JOBCODE%>" onkeypress="return harusangka(event);" onchange="javascript:getHarvestingQualityData(); javascript:getJobcodeIntraco();javascript:getJobcodeSpbHarvesting(); javascript:getListChitno();" onblur="javascript:getHarvestingQualityData(); javascript:getJobcodeIntraco();javascript:getJobcodeSpbHarvesting(); javascript:getListChitno();">
					<font color="#FF0000">&nbsp;*&nbsp;</font>
					<!--<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>-->
					<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=BKM&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>

					<span class="satuan_item" id="jobname">&nbsp;<%= JobCodeDescription %></span>
				</td>
			</tr>
			<tr style="display: block; padding-left:22px">
				<td align="right">
					<font face="Verdana"   size="1">TARGET TYPE</font>
				</td>
				<td align="left">
					<input type="text" size="15" id="locationtype" name="LOCATIONTYPE" value="<%=locationtype%>"style="background-color: #eee;">
					<font color="#FF0000">&nbsp;*&nbsp;</font>
					<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a>
				</td>
			<tr style="display: block; padding-left:22px">
				<td align="right">
					<font face="Verdana"   size="1">TARGET CODE</font>
				</td>
				<td align="left">
					<input type="text" size="15" id="targetcode" name="LOCATIONCODE" value="<%=LOCATIONCODE%>" onchange="javascript:getJobcodeSpbHarvesting(); javascript:getListChitno();">
					<font color="#FF0000"> *</font><a href="javascript:PopupTType('<%=session("fcBA") %>')">
                            <font face="Verdana" size="1">Location Code</font></a>
				</td>
			</tr>
			<tr id="jobcodeintraco">
			<tr style="display: block; padding-left:25px" id="chargeunit">
				<td align="right">
					<font face="Verdana" size="1">CHARGE UNIT</font>
				</td>
                <td align="left">
                    <select name="chargeunit" id = "chargeunit_">
							<option value=""></option>
							<%
							qlov = 	 " select fccode,fcname " & _
									 "	from businessunit " & _
									 "	where fccompanycode = businessunit_api.get_CompanyCode('" & session("fcba") & "') " & _
									 "	and businessunit_api.Isactive(fccode)='TRUE' " & _ 
									 "  and fccode <> '" & session("fcba") & "'"
									 
							Set rslov = Server.CreateObject("ADODB.Recordset")
							rslov.Open qlov, connstr 
							if Not rslov.eof then
								While Not rslov.EOF
							%>
							<option value="<%=rslov("fccode")%>" <%if chargeunit = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%></option>
							<%
									rslov.MoveNext
								Wend
							end if
							%>
					</select>
                </td>
            </tr>
			<tr style="display: block; padding-left:25px" id="chargejob">
				<td align="right">
					<font face="Verdana" size="1">CHARGE JOB</font>
				</td>
                <td align="left">
                    <input type="text" size="11" name="chargejob" id="chargejob_" onchange="javascript:getChargetype();" onkeypress="return harusangka(event);"  value="<%=chargejob%>">&nbsp;<font color="#FF0000">*</font>
                    <!--<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">-->
				    <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=BKM&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>

					
                </td>
            </tr>
            <tr style="display: block; padding-left:22px" id="chargetype">
				<td align="right">
					<font face="Verdana"   size="1">CHARGE TYPE</font>
				</td>
                <td align="left">
                    <input type="text" size="11" name="chargetype" id="chargetype_" value="<%=chargetype%>">&nbsp;<font
                        color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                            <font face="Verdana" size="1">Location Type</font></a>
                </td>
            </tr>
            <tr style="display: block; padding-left:19px" id="chargecode">
				<td align="right">
					<font face="Verdana"   size="1">CHARGE CODE</font>
				</td>
                <td align="left">
                    <input type="text" size="11"  name="chargecode" value="<%=chargecode%>">&nbsp;<font
                        color="#FF0000">*</font><a href="javascript:PopupTType('')">
                            <font face="Verdana" size="1">Location Code</font></a>
                </td>
            </tr>
            </tr>
            <tr style="display: block; padding-left:43px">
				<td align="right">
					<font face="Verdana" size="1">MANDAYS</font>
				</td>
				<td align="left">
					<input type="text" size="2" name="MANDAYS" value="<%=MANDAYS%>" onkeypress="return harusangka(event);">
					<span class="satuan_item" id="mandays">&nbsp;<%= MandaysDescription %></span>
				</td>
			</tr>
		<%End if%>

		<!--<tr><td align="right"><font face="Verdana"   size="1">MANDAYS</td><td align="left"><input type="text" size="2" name="MANDAYS" value="<%=MANDAYS%>"></td></tr>-->
		<tr id="jobcodespbharvesting">
			<tr style="display: block; padding-left:53px" id="chitno">
				<td align="right">
					<font face="Verdana"   size="1">SPB NO</font>
				</td>
				<td align="left" >
					<div id="listchitno" style="width: 200px;"></div>
					<input type="hidden" size="50" id="bjr_per_chitno" name="bjr_per_chitno" value="<%=bjr_per_chitno%>" onkeyup="hitungrumus();" onchange=" hitungrumus();" onblur="hitungrumus();">
					<input type="hidden" size="50" id="kg_per_bucket" name="kg_per_bucket" value="<%=kg_per_bucket%>" onkeyup="hitungrumus();" onchange=" hitungrumus();" onblur="hitungrumus();">
				</td>
			</tr>
			<tr style="display: block; padding-left:5px" id="kg_janjang" >
				<td align="right">
					<font face="Verdana" size="1">JUMLAH JANJANG</font>
				</td>
				<td align="left">
					<input type="text" size="5" id="kg_janjang_harvesting" name="kg_janjang" value="<%=kg_janjang_harvesting%>" onkeypress="return harusangka(event);" onkeyup="hitungrumus();" onchange=" hitungrumus();" onblur="hitungrumus();">
				</td>
			</tr>
		</tr>
		<tr id="jobcodespbharvesting">
			<tr style="display: block; padding-left:67px"  id="vunit">
				<td align="right">
					<font face="Verdana"   size="1">UNIT</font>
				</td>
				<td align="left">
					<input type="text" size="5" id="unitharvesting" name="unitharvesting" value="<%=unitharvesting%>" onkeypress="return harusangka(event);" >
					<span class="satuan_item" id="unit">&nbsp;<%= UnitDescription %></span>
				</td>
			</tr> 
		</tr>
		<tr style="display: block; padding-left:67px" id="vunits">
			<td align="right">
				<font face="Verdana"  size="1"><label id ="lblunit" text="UNIT"></label></font>
			</td>
			<td align="left">
				<input type="text" size="5"  name="UNIT" value="<%=UNIT%>" onkeypress="return harusangka(event);" >
				<span class="satuan_item" id="unit">&nbsp;<%= UnitDescription %></span>
			</td>
		</tr> 
		<tr style="display: none; padding-left:67px" id="panenunits">
			<td align="right">
				<font face="Verdana"  size="1">ABW</font>&nbsp;<input type="text" size="5"  name="ABW" value="<%=ABW%>" onkeypress="return harusangka(event);">
			</td>
			<td align="left">
			    <font face="Verdana"  size="1">BUNCH WEIGHT</font>&nbsp;<input type="text" size="5"  name="BUNCH_WEIGHT" value="<%=BUNCH_WEIGHT%>" onkeypress="return harusangka(event);">
			</td>
		</tr> 
		<tr style="display: block; padding-left:53px">
			<td align="right">
				<font face="Verdana"   size="1">OUTPUT</font>
			</td>
			<td align="left">
				<input type="text" size="5" name="OUTPUT" value="<%=OUTPUT%>" onkeypress="return harusangka(event);">
				<span class="satuan_item" id="output">&nbsp;<%= OutputDescription %></span>
			</td>
		</tr>
		<tr style="display: none; padding-left:10px" id="kg_brondolan" >
				<td align="right">
					<font face="Verdana" size="1">KG BRONDOLAN</font>
				</td>
				<td align="left">
					<input type="text" size="5" name="kg_brondolan" value="<%=kg_brondolan%>" onkeypress="return harusangka(event);">
				</td>
			</tr>
		<tr style="display: block;">
			<td align="right">
				<font face="Verdana"   size="1">OVERTIME HOURS</font>
			</td>
			<td align="left">
				<input type="text" size="5" name="OVERTIME_HOURS" id="hours" value="<%=OVERTIME_HOURS%>" onkeypress="return harusangka(event);" onchange="javascript:getOvertime();" onkeyup="javascript:getOvertime();" onblur="javascript:getOvertime();">
				<span class="satuan_item" id="oth">&nbsp;<%= OthDescription %></span>
			</td>
		</tr>
		<tr style="display: block; padding-left:10px" >
			<td align="right">
				<font face="Verdana" size="1">TYPE OVERTIME</font>
			</td>
			<td align="left">
				<select id="typex" name="TYPE_OVERTIME" onchange="javascript:getOvertime();">
						<% 	Dim rsComboC
							q = "select distinct type, fcname from overtime_convertion where fcba = '"& session("fcba") &"' order by type" 
							' response.write q
							Set rsComboC = Server.CreateObject("ADODB.Recordset")
							rsComboC.CursorType = 3
							rsComboC.CursorLocation = 3
							rsComboC.Open q, connstr
							if not rsComboC.eof then%>
							<%
							Do While Not rsComboC.eof
							%>
								<option value="<%=rsComboC("type")%>" 
									<%if cstr(""&TYPE_OVERTIME)=cstr(""&rsComboC("type")) then %> selected 
									<%End If%>><%=rsComboC("type")%> : <%=rsComboC("FCNAME")%>
								</option> 
							<%rsComboC.movenext
							loop
							End if
							rsComboC.Close
							set rsComboC = Nothing
							q=""
						%>
				</select>
			</td>
		</tr>
		<tr style="display: block;">
			<td align="right">
				<font face="Verdana" size="1">OTH CONVERTION</font>
			</td>
			<td align="left">
				<input type="text" size="5" id="convertion" name="OTHRS" value="<%=OTHRS%>" style="background-color: #eee" readonly onkeypress="return harusangka(event);">
			</td>
		</tr>
		<tr style="display: block; padding-left:64px">
			<td align="right">
				<font face="Verdana"   size="1">RATE</font>
			</td>
			<td align="left">
				<input type="text" size="5" name="RATE" value="<%=RATE%>" onkeypress="return harusangka(event);">
			</td>
		</tr>
		<tr style="display: block; padding-left:29px">
			<td align="right">
				<font face="Verdana"   size="1">REFERENCE</font>
			</td>
			<td align="left">
				<input type="text" size="50" name="REFERENCE" value="<%=REFERENCE%>">
			</td>
		</tr>
		<tr style="display: block; padding-left:37px">
			<td align="right">
				<font face="Verdana"   size="1">REMARKS</font>
			</td>
			<td align="left">
				<input type="text" size="50" name="REMARKS" value="<%=REMARKS%>">
			</td>
		</tr>
		<tr bgcolor="#00a300" style="width:1235px;" >
			<td colspan="4">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Isian Premi (Pilih Kode Premi dan Masukan Total Amount Premi)</strong></font>
				</div>
			</td>
		</tr>
		<tr style="display: block; padding-left:32px">
			<td><font face="Verdana"   size="1">ALLOWDED</font></td>
			<td align="left">
				<select name="ALLOWANCE_CODE" id="allowded" <%if ( len(trim(""&ALLOWANCE_CODE))) > 0 then%> value="<%=ALLOWANCE_CODE%>" selected <%end if%> >
					<option value="">--Choose Allowded--</option>
					<%
					strallowded = " select fccode, fcname" & _
								  " from allowded" & _
								  " where fcba = '" & Session("fcBA") & "'" & _
								  "       and man_auto = 'M'" & _
								  "       and allded = 'A'" & _
								  "       and fccode in ('A99')" & _
								  "       and includedpayrollprocessing = 'TRUE'" & _
								  " order by fccode asc"
					
					Set rsallowded = Server.CreateObject("ADODB.Recordset")
					rsallowded.Open strallowded, connstr, adOpenStatic 
					
					if not rsallowded.eof then
					Do While Not rsallowded.eof
					%>
						<option value="<%=rsallowded("fccode")%>" <%if cstr(""&ALLOWANCE_CODE) = rsallowded("fccode") then%> Selected <%end if%>><%=rsallowded("fccode")%> : <%=rsallowded("fcname")%></option>
					<%
						rsallowded.movenext
					loop
					end if
					rsallowded.Close
					set rsallowded = Nothing
					%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr style="display: block; padding-left:43px" id="foo">
			<td align="right" >
				<font face="Verdana" size="1">AMOUNT</font>
			</td>
			<td align="left">
				<input type="text" size="50" id="amount" name="AMOUNT" value="<%=AMOUNT%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
			</td>
		</tr>
		<tr style="display: block; padding-left:43px" id="foo2">
			<td align="right" >
				<font face="Verdana" size="1">REMARKS</font>
			</td>
			<td align="left">
				<textarea id="remarksallowded" name="remarksallowded" cols = "50"><%=remarksallowded%></textarea>
			</td>
		</tr>
		<tr id="harvestingquality" style="width:100%">
		<td>
		<table style="width:100%">
		<tr bgcolor="#00a300" style="width:1235px;" id="HQ41"> 
			<td colspan="4">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Harvesting Penalty</strong></font>
				</div>
			</td>
		</tr>
		<tr><td></td></tr>
		<tr bgcolor="#00a300" style="width:1235px;" id="HQ42">
			<td colspan="4">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Kualitas TBS (Data bersumber dari Buku Penerimaan TBS dan Pemeriksaan Mutu Buah) </strong></font>
				</div>
			</td>
		</tr>
		<tr align="left" id="HQ44">
			<td colspan="4">
				<div align="left">
					<h4 face="Verdana"  style="color:#FF0000"  size="1">Penalty</h4>
				</div>
			</td>
		</tr>
		<tr>
			<td align="left" id="HQ5">
				<font face="Verdana" size="1">Buah Mentah (A)</font>
			</td>
			<td align="left" id="HQ6">
				<input type="text" size="15" id="under_ripe" name="UNDER_RIPE" value="<%=UNDER_RIPE%>" onkeyup="checkInput(this);">
			</td>
			<td align="left" id="HQ31">
				<font face="Verdana"   size="1">Tangkai Panjang (TP)</font>
			</td>
			<td align="left" id="HQ32">
				<input type="text" size="15" id="long_stalk" name="LONG_STALK" value="<%=LONG_STALK%>" onkeyup="checkInput(this);">
			</td>
		</tr>
		<tr>
			<td align="left" id="HQ23">
				<font face="Verdana"   size="1">Brondolan Tidak di Alas (BA)</font>
			</td>
			<td align="left" id="HQ24">
				<input type="text" size="15" id="qe_6" name="QE_6" value="<%=QE_6%>" onkeyup="checkInput(this);">
			</td>
			<td align="left" id="HQ35">
				<font face="Verdana"   size="1">Brondolan Kotor/ Sampah (bk)</font>
			</td>
			<td align="left" id="HQ36">
				<input type="text" size="15" id="qe_7" name="QE_7" value="<%=QE_7%>" onkeyup="checkInput(this);">
			</td>
			
		</tr>
		<tr align="left" id="HQ45">
			<td colspan="4">
				<div align="left">
					<h4 face="Verdana"  style="color:#FF0000"  size="1">Lain - Lain</h4>
				</div>
			</td>
		</tr>
		<tr>
			<td align="left" id="HQ9">
				<font face="Verdana"   size="1">Buah Terlalu Matang (O)</font>
			</td>
			<td align="left" id="HQ10">
				<input type="text" size="15" id="over_ripe" name="OVER_RIPE" value="<%=OVER_RIPE%>" onkeyup="checkInput(this);">
			</td>
			<td align="left" id="HQ11">
				<font face="Verdana"   size="1">Buah Busuk (E)</font>
			</td>
			<td align="left" id="HQ12">
				<input type="text" size="15" id="qe_3" name="QE_3" value="<%=QE_3%>" onkeyup="checkInput(this);">
			</td>
		</tr>
		<tr>
			<td align="left" id="HQ13">
				<font face="Verdana"   size="1">Buah Abnormal (AB)</font>
			</td>
			<td align="left" id="HQ14">
				<input type="text" size="15" id="abnormal" name="ABNORMAL" value="<%=ABNORMAL%>" onkeyup="checkInput(this);">
			</td>
			<td align="left" id="HQ37">
			</td>
			<td align="left" id="HQ38">
			</td>
			
		</tr>
		<tr>
			<td align="left" id="HQ17">
				
			</td>
			<td align="left" id="HQ18">

			</td>
			<td align="left" id="HQ27">
			
			</td>
			<td align="left" id="HQ28">
			</td>
			
		</tr>
		<tr>
			<td align="left" id="HQ39">
			
			</td>
			<td align="left" id="HQ40">
				
			</td>

		</tr>
		<tr bgcolor="#00a300" style="width:1235px;" id="HQ43">
			<td colspan="4">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Kulitas Ancak (Data bersumber dari Buku Pemeriksaan Ancak Panen)</strong></font>
				</div>
			</td>
		</tr>
		<tr align="left" id="HQ46">
			<td colspan="4">
				<div align="left">
					<h4 face="Verdana"  style="color:#FF0000"  size="1">Penalty</h4>
				</div>
			</td>
		</tr>
		<tr>
			<td align="left" id="HQ33">
				<font face="Verdana"   size="1">Buah Tinggal (BT)</font>
			</td>
			<td align="left" id="HQ34">
				<input type="text" size="15" id="unharvest_ffb" name="UNHARVEST_FFB" value="<%=UNHARVEST_FFB%>" onkeyup="checkInput(this);">
			</td>
			<td align="left" id="HQ29">
				<font face="Verdana"   size="1">Brondolan Tinggal (BRD)</font>
			</td>
			<td align="left" id="HQ30">
				<input type="text" size="15" id="uncollect_lf_circle" name="UNCOLLECT_LF_CIRCLE" value="<%=UNCOLLECT_LF_CIRCLE%>" onkeyup="checkInput(this);">
			</td>
		</tr>
		<tr>
			<td align="left" id="HQ1">
				<font face="Verdana"   size="1">Pelepah Sengkleh (PS)</font>
			</td>
			<td align="left" id="HQ2">
				<input type="text" size="15" id="unprune_frond" name="UNPRUNE_FROND" value="<%=UNPRUNE_FROND%>" onkeyup="checkInput(this);">
			</td>
			<td align="left" id="HQ19">
				<font face="Verdana"   size="1">Over Pruning (OP)</font>
			</td>
			<td align="left" id="HQ20">
				<input type="text" size="15" id="qe_5" name="QE_5" value="<%=QE_5%>" onkeyup="checkInput(this);">
			</td>
			
		</tr>
		<tr>
			<td align="left" id="HQ3">
				<font face="Verdana"   size="1">Pelepah Tidak Disusun Rapi (SP)</font>
			</td>
			<td align="left" id="HQ4">
				<input type="text" size="15" id="qe_1" name="QE_1" value="<%=QE_1%>" onkeyup="checkInput(this);">
			</td>
			<td align="left" id="HQ7">
				<font face="Verdana"   size="1">Buah Matahari (bm)</font>
			</td>
			<td align="left" id="HQ8">
				<input type="text" size="15" id="qe_2" name="QE_2" value="<%=QE_2%>" onkeyup="checkInput(this);">
			</td>
			<td align="left" id="HQ25">				
			</td>
			<td align="left" id="HQ26">
			</td>
		</tr>
		<tr>
			<td align="left" id="HQ15">
			</td>
			<td align="left" id="HQ16">
			</td>
		</tr>
		<tr align="left" id="HQ47">
			<td colspan="4">
				<div align="left">
					<h4 face="Verdana"  style="color:#FF0000"  size="1">Lain - Lain</h4>
				</div>
			</td>
		</tr>
		<tr>
			<td align="left" id="HQ21">
				<font face="Verdana"   size="1">Buah Dimakan Tikus (bdt)</font>
			</td>
			<td align="left" id="HQ22">
				<input type="text" size="15" id="eaten_by_rat" name="eaten_by_rat" value="<%=EATEN_BY_RAT%>" onkeyup="checkInput(this);">
			</td>
		</td>
		</tr>
		</table>
		</tr>
	    <TR>
			<td colspan="2">
				<input type="hidden" name="gangcode" value="<%=gangcode%>">
				<input type="hidden" name="sectionname" value="<%=request("sectionname")%>">
				<input type="hidden" name="fdDate" value="<%=fdDate%>">
				<input type="hidden" name="rowstate" value="<%=rowstate%>">
				<input type="hidden" name="document_classification" value="<%=request("document_classification")%>">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				
				<%If SudahClose <> "Y" Then %>				
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
						<input Type="Submit"  name="Add" value="Update"  OnClick ="periksa(this.form)" >
					<%Else%>
						<input type="hidden" name="txtJenis" value="New">
						<input Type="Submit"  name="Add" value="Add" OnClick ="periksa(this.form); submit();">
					<%End if%>
					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				<%End If%>
			</td>
		</tr>
	</table>

	<table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">LINENOKEY</TD>	
	<TD align="left">GANG</TD>	
	<TD align="left">DATE</TD>	
	<TD align="left">SUP_1</TD>	
	<TD align="left">SUP_2</TD>	
	<TD align="left">SUP_3</TD>	
	<TD align="left">SUP_4</TD>	
	<TD align="left">EMPLOYEE CODE</TD>	
	<TD align="left">ATTENDANCE</TD>	
	<TD align="left">JOB CODE</TD>	
	<TD align="left">LOC TYPE</TD>	
	<TD align="left">LOC CODE</TD>	
	<TD align="left">CHARGE UNIT</TD>	
	<TD align="left">CHARGE JOB</TD>	
	<TD align="left">CHARGE TYPE</TD>	
	<TD align="left">CHARGE CODE</TD>	
	<TD align="left">MANDAYS</TD>	
	<TD align="left">OVERTIME</TD>	
	<TD align="left">RATE</TD>	
	<TD align="left">UNIT/BUNCH</TD>	
	<TD align="left">BUNCH WEIGHT</TD>	
	<TD align="left">SPB NO</TD>	
	<TD align="left">KG BRONDOLAN</TD>	
	<TD align="left">JANJANG</TD>	
	<TD align="left">BUCKET</TD>
	<TD align="left">OUTPUT</TD>	
	<TD align="left">REFERENCES</TD>
	<TD align="left">ALLOWANCE CODE</TD>	
	<TD align="left">AMOUNT ALLOWANCE</TD>
	<TD align="left">Buah Mentah (A)</TD>	
	<TD align="left">Brondolan Tidak di Alas (BA)</TD>	
	<TD align="left">Tangkai Panjang (TP)</TD>	
	<TD align="left">Brondolan Kotor / Sampah (bk)</TD>	
	<TD align="left">Buah Terlalu Matang (O)</TD>	
	<TD align="left">Buah Abnormal (AB)</TD>	
	<TD align="left">Buah Busuk (E)</TD>	
	<TD align="left">Buah Tinggal (BT)</TD>	
	<TD align="left">Pelepah Sengkleh (PS)</TD>	
	<TD align="left">Pelepah Tidak Disusun Rapih (SP)</TD>	
	<TD align="left">Brondolan Tinggal (BRD)</TD>	
	<TD align="left">Over Prunning (OP)</TD>	
	<TD align="left">Buah Matahari (bm)</TD>	
	<TD align="left">Buah Dimakan Tikus (bdt)</TD>	
	</tr>
   <%
   'response.write Session("v_Employee")
   
   if Not rsGAD.eof then
		NO = 0
		While Not rsGAD.EOF
			NO = NO +1
			if no  mod 2 = 1 then cColor = "#EEEEEE" else cColor = "#E5E5E5"
				%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			
			<!-- Validasi Period Control In Here -->
				<%If rowstate = "Approved" Or rowstate = "AuthorizeOnProgress"  Then%>
					<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>		
				<%Else%>
				<td valign="middle" width="30">
					<a href="Attendance_GAD.Asp?sectionname=<%=sectionname%>&GangCode=<%=rsGAD("GangCode")%>&selectedgangcode=<%=rsGAD("GangCode")%>&document_classification=<%=document_classification%>&fdDate=<%=fdDate%>&LineNoKey=<%=rsGAD("linenokey")%>&EMPLOYEECODE=<%=rsGAD("EMPLOYEECODE")%>
					&JOBCODE=<%=rsGAD("JOBCODE")%>&locationtype=<%=rsGAD("locationtype")%>&locationcode=<%=rsGAD("locationcode")%>
					&supervision_1=<%=rsGAD("supervision_1")%>&supervision_2=<%=rsGAD("supervision_2")%>
					&supervision_3=<%=rsGAD("supervision_3")%>&supervision_4=<%=rsGAD("supervision_4")%>&ATTENDANCE=<%=rsGAD("ATTENDANCE")%>
					&MANDAYS=<%=rsGAD("MANDAYS")%>&UNIT=<%=rsGAD("UNIT")%>&UNITHARVESTING=<%=rsGAD("UNIT")%>&SPBNO=<%=rsGAD("SPBNO")%>&bucket=<%=rsGAD("bucket")%>&kg_brondolan=<%=rsGAD("kg_brondolan")%>
					&OUTPUT=<%=rsGAD("OUTPUT")%>&OTHRS=<%=rsGAD("OTHRS")%>
					&RATE=<%=rsGAD("RATE")%>&REFERENCE=<%=rsGAD("REFERENCE")%>&ALLOWANCE_CODE=<%=rsGAD("ALLOWANCE_CODE")%>&AMOUNT=<%=rsGAD("AMOUNT")%>
					&REMARKS=<%=rsGAD("REMARKS")%>&TYPE_OVERTIME=<%=rsGAD("TYPE_OVERTIME")%>&OVERTIME_HOURS=<%=rsGAD("OVERTIME_HOURS")%>&CHARGEUNIT=<%=rsGAD("CHARGEUNIT")%>&CHARGEJOB=<%=rsGAD("CHARGEJOB")%>&CHARGETYPE=<%=rsGAD("CHARGETYPE")%>&CHARGECODE=<%=rsGAD("CHARGECODE")%>
					&UNDER_RIPE=<%=UNDER_RIPE%>&OVER_RIPE=<%=OVER_RIPE%>&ABNORMAL=<%=ABNORMAL%>&LONG_STALK=<%=LONG_STALK%>&EATEN_BY_RAT=<%=EATEN_BY_RAT%>&UNHARVEST_FFB=<%=UNHARVEST_FFB%>&UNCOLLECT_LF_CIRCLE=<%=UNCOLLECT_LF_CIRCLE%>&UNCOLLECT_LF_PIECE=<%=UNCOLLECT_LF_PIECE%>&UNARRANGE_FFB=<%=UNARRANGE_FFB%>&UNPRUNE_FROND=<%=UNPRUNE_FROND%>&QE_1=<%=QE_1%>&QE_2=<%=QE_2%>&QE_3=<%=QE_3%>&QE_4=<%=QE_4%>&QE_5=<%=QE_5%>&QE_6=<%=QE_6%>&QE_7=<%=QE_7%>&QE_8=<%=QE_8%>&QE_9=<%=QE_9%>&QE_10=<%=QE_10%>
					&ABW=<%=request("ABW")%>&BUNCH_WEIGHT=<%=request("BUNCH_WEIGHT")%>
					&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit"  onkeyup="showHint(this.value)">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
				</td>
				<td>
					<a href="Attendance_GAD.Asp?sectionname=<%=sectionname%>&GangCode=<%=rsGAD("GangCode")%>&selectedgangcode=<%=rsGAD("GangCode")%>&document_classification=<%=document_classification%>&fdDate=<%=fdDate%>&LineNoKey=<%=rsGAD("linenokey")%>&EMPLOYEECODE=<%=rsGAD("EMPLOYEECODE")%>
					&JOBCODE=<%=rsGAD("JOBCODE")%>&locationtype=<%=rsGAD("locationtype")%>&locationcode=<%=rsGAD("locationcode")%>
					&supervision_1=<%=rsGAD("supervision_1")%>&supervision_2=<%=rsGAD("supervision_2")%>
					&supervision_3=<%=rsGAD("supervision_3")%>&supervision_4=<%=rsGAD("supervision_4")%>&ATTENDANCE=<%=rsGAD("ATTENDANCE")%>
					&MANDAYS=<%=rsGAD("MANDAYS")%>&UNIT=<%=rsGAD("UNIT")%>&UNITHARVESTING=<%=rsGAD("UNIT")%>&SPBNO=<%=rsGAD("SPBNO")%>&bucket=<%=rsGAD("bucket")%>&kg_brondolan=<%=rsGAD("kg_brondolan")%>
					&OUTPUT=<%=rsGAD("OUTPUT")%>&OTHRS=<%=rsGAD("OTHRS")%>
					&RATE=<%=rsGAD("RATE")%>&REFERENCE=<%=rsGAD("REFERENCE")%>&ALLOWANCE_CODE=<%=rsGAD("ALLOWANCE_CODE")%>&AMOUNT=<%=rsGAD("AMOUNT")%>
					&REMARKS=<%=rsGAD("REMARKS")%>&TYPE_OVERTIME=<%=rsGAD("TYPE_OVERTIME")%>&OVERTIME_HOURS=<%=rsGAD("OVERTIME_HOURS")%>&CHARGEUNIT=<%=rsGAD("CHARGEUNIT")%>&CHARGEJOB=<%=rsGAD("CHARGEJOB")%>&CHARGETYPE=<%=rsGAD("CHARGETYPE")%>&CHARGECODE=<%=rsGAD("CHARGECODE")%>
					&UNDER_RIPE=<%=UNDER_RIPE%>&OVER_RIPE=<%=OVER_RIPE%>&ABNORMAL=<%=ABNORMAL%>&LONG_STALK=<%=LONG_STALK%>&EATEN_BY_RAT=<%=EATEN_BY_RAT%>&UNHARVEST_FFB=<%=UNHARVEST_FFB%>&UNCOLLECT_LF_CIRCLE=<%=UNCOLLECT_LF_CIRCLE%>&UNCOLLECT_LF_PIECE=<%=UNCOLLECT_LF_PIECE%>&UNARRANGE_FFB=<%=UNARRANGE_FFB%>&UNPRUNE_FROND=<%=UNPRUNE_FROND%>&QE_1=<%=QE_1%>&QE_2=<%=QE_2%>&QE_3=<%=QE_3%>&QE_4=<%=QE_4%>&QE_5=<%=QE_5%>&QE_6=<%=QE_6%>&QE_7=<%=QE_7%>&QE_8=<%=QE_8%>&QE_9=<%=QE_9%>&QE_10=<%=QE_10%>
					&ABW=<%=request("ABW")%>&BUNCH_WEIGHT=<%=request("BUNCH_WEIGHT")%>
					&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete" onkeyup="showHint(this.value)">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				</TD>
				<%End If%>
			<!-- End -->	 
			</td>
		   		<td align="center" title="<%=rsGAD("linenokey").Name%>"><%=rsGAD("linenokey")%></td>
				<td align="center" title="<%=rsGAD("GangCode").Name%>"><%=rsGAD("GangCode")%></td>
				<td align="center" title="<%=rsGAD("fddate").Name%>"><%=rsGAD("fddate")%></td>
				<td align="center" title="<%=rsGAD("SUPERVISION_1").Name%>"><%=rsGAD("SUPERVISION_1")%> - <%=rsGAD("spvname1")%></td>
				<td align="center" title="<%=rsGAD("SUPERVISION_2").Name%>"><%=rsGAD("SUPERVISION_2")%> - <%=rsGAD("spvname2")%></td>
				<td align="center" title="<%=rsGAD("SUPERVISION_3").Name%>"><%=rsGAD("SUPERVISION_3")%> - <%=rsGAD("spvname3")%></td>
				<td align="center" title="<%=rsGAD("SUPERVISION_4").Name%>"><%=rsGAD("SUPERVISION_4")%> - <%=rsGAD("spvname4")%></td>
				<td align="center" title="<%=rsGAD("employeecode").Name%>"><%=rsGAD("employeecode")%> - <%=rsGAD("empname")%></td>
				<td align="center" title="<%=rsGAD("attendance").Name%>"><%=rsGAD("attendance")%></td>
				<td align="center" title="<%=rsGAD("jobcode").Name%>"><%=rsGAD("JobDet")%></td>
				<td align="center" title="<%=rsGAD("locationtype").Name%>"><%=rsGAD("locationtype")%></td>
				<td align="center" title="<%=rsGAD("locationcode").Name%>"><%=rsGAD("locationcode")%></td>
				<td align="center" title="<%=rsGAD("chargeunit").Name%>"><%=rsGAD("chargeunit")%></td>
				<td align="center" title="<%=rsGAD("chargejob").Name%>"><%=rsGAD("CJobDet")%></td>
				<td align="center" title="<%=rsGAD("chargetype").Name%>"><%=rsGAD("chargetype")%></td>
				<td align="center" title="<%=rsGAD("chargecode").Name%>"><%=rsGAD("chargecode")%></td>
				<td align="center" title="<%=rsGAD("mandays").Name%>"><%=rsGAD("mandays")%></td>
				<td align="center" title="<%=rsGAD("othrs").Name%>"><%=cdbl(to_nol(rsGAD("othrs")))%></td>
				<td align="center" title="<%=rsGAD("rate").Name%>"><%=rsGAD("rate")%></td>
				<td align="center" title="<%=rsGAD("unit").Name%>"><%=rsGAD("unit")%></td>
				<td align="center" title="<%=rsGAD("bunch_weight").Name%>"><%=rsGAD("bunch_weight")%></td>				
				<td align="center" title="<%=rsGAD("spbno").Name%>"><%=rsGAD("spbno")%></td>
				<td align="center" title="<%=rsGAD("kg_brondolan").Name%>"><%=rsGAD("kg_brondolan")%></td>
				<td align="center" title="<%=rsGAD("kg_janjang").Name%>"><%=rsGAD("kg_janjang")%></td>
				<td align="center" title="<%=rsGAD("BUCKET").Name%>"><%=rsGAD("BUCKET")%></td>
				<td align="center" title="<%=rsGAD("output").Name%>"><%=rsGAD("output")%></td>
				<td align="center" title="<%=rsGAD("reference").Name%>"><%=rsGAD("reference")%></td>
				<td align="center" title="<%=rsGAD("allowance_code").Name%>"><%=rsGAD("allowance_code")%></td>
				<td align="center" title="<%=rsGAD("amount").Name%>"><%=rsGAD("amount")%></td>
				<td align="center" title="<%=rsGAD("under_ripe").Name%>"><%=rsGAD("under_ripe")%></td>
				<td align="center" title="<%=rsGAD("qe_6").Name%>"><%=rsGAD("qe_6")%></td>
				<td align="center" title="<%=rsGAD("long_stalk").Name%>"><%=rsGAD("long_stalk")%></td>
				<td align="center" title="<%=rsGAD("QE_7").Name%>"><%=rsGAD("QE_7")%></td>
				<td align="center" title="<%=rsGAD("over_ripe").Name%>"><%=rsGAD("over_ripe")%></td>
				<td align="center" title="<%=rsGAD("abnormal").Name%>"><%=rsGAD("abnormal")%></td>
				<td align="center" title="<%=rsGAD("qe_3").Name%>"><%=rsGAD("qe_3")%></td>
				<td align="center" title="<%=rsGAD("unharvest_ffb").Name%>"><%=rsGAD("unharvest_ffb")%></td>
				<td align="center" title="<%=rsGAD("unprune_frond").Name%>"><%=rsGAD("unprune_frond")%></td>
				<td align="center" title="<%=rsGAD("qe_1").Name%>"><%=rsGAD("qe_1")%></td>
				<td align="center" title="<%=rsGAD("uncollect_lf_circle").Name%>"><%=rsGAD("uncollect_lf_circle")%></td>
				<td align="center" title="<%=rsGAD("qe_5").Name%>"><%=rsGAD("qe_5")%></td>
				<td align="center" title="<%=rsGAD("QE_2").Name%>"><%=rsGAD("QE_2")%></td>
				<td align="center" title="<%=rsGAD("eaten_by_rat").Name%>"><%=rsGAD("eaten_by_rat")%></td>
		  </tr>
		 <%
			rsGAD.MoveNext
		Wend
	end if 
	rsGAD.Close : set rsGAD = Nothing
%>
   </div>
	</TABLE>

</form>
<script language="JavaScript" type="text/javascript">
//You should create the validator only after the definition of the HTML form
  var frmvalidator  = new Validator("eForm");
  frmvalidator.addValidation("EMPLOYEECODE","req","Please enter Emp.Code");
  frmvalidator.addValidation("JOBCODE","req","Please enter Job.Code");
  frmvalidator.addValidation("LOCATIONTYPE","req","Please enter Location Type");
  frmvalidator.addValidation("LOCATIONCODE","req","Please enter Location Code");
  frmvalidator.addValidation("MANDAYS","ltelmnt=1.001","Total Mandays perhari lebih dari 1");
  frmvalidator.addValidation("MANDAYS","gtelmnt=0","Total Mandays perhari lebih dari 1");

</script>
   
<%End if%>  
 
  <%'-------------akhir----------------------
  %>
</div></div>
  <!--#include file="bottom.jsp"-->
<script>
function submit() {
  /*Put all the data posting code here*/
 document.getElementById("myForm").reset();
}
</script>
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
	function handleClick(avail){
		//var xx=document.getElementById("tipehidden");
		if(avail.checked == 1){
			avail.value = 2;
			//xx.value = 2;
		} else {
			avail.value = 1;
			//xx.value = 1;
		}
	}

function getselectedgangcode() {
	var sgangcode = document.getElementById("selectedgangcode_");
	sgangcode.value = document.getElementById("gangcode").value;
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

$("#allowded").change(function() {
	// foo is the id of the other select box 
	if ($("#allowded").val() != "") {
		$("#foo").show();
		$("#foo2").show();
	}else{
		$("#amount").val(0);
		$("#foo").hide();
		$("#foo2").hide();
		$("#remarksallowded").hide();
	} 
});


document.ready(function() {
	getGangcode();
	getselectedgangcode();
	getOvertime();
	$("#foo").hide();
	 if ($("#allowded").val() != "") {
	        $("#foo").show();
			$("#foo2").show();
	    }else{
	        $("#foo").hide();
			$("#foo2").hide();
	    } 
	document.getElementById("kg_brondolan").style.display="none";
	document.getElementById("harvestingquality").style.display="none";
	getHarvestingQualityData();
	document.getElementById("jobcodeintraco").style.display="none";
	getJobcodeIntraco();
	document.getElementById("jobcodespbharvesting").style.display="none";
	getJobcodeSpbHarvesting();
	getListChitno();
	getBjrSpbHarvesting();
	hitungrumus();
	getTragettype();
	
});
</script> 
	
</body>
</html>
