<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<title>Fixed Asset | FB-Plants</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

function CheckIsActive()
{
 var cIsActive = document.getElementById('ISACTIVE');
 if (cIsActive.checked != true )
	{
		alert("Is Active Belum di Click");
	}
}

function calcParam()
{
	var cIsActive = document.getElementById('ISACTIVE').checked;
	if(cIsActive)
	{
		document.getElementById("hide_ISACTIVE").value ="1";
	}else
	{
		document.getElementById("hide_ISACTIVE").value ="0";
	}
	
	var cIsBlock = document.getElementById('ISBLOCK').checked;
	if(cIsBlock)
	{
		document.getElementById("hide_ISBLOCK").value ="1";
	}else
	{
		document.getElementById("hide_ISBLOCK").value ="0";
	}
		
}

function populateSelect(target, min, max){
	if (!target){
		return false;
	}
	else {
		var min = min || 0,
			max = max || min + 0;

		select = document.getElementById(target);

		for (var i = min; i<=max; i++){
			var opt = document.createElement('option');
			opt.innerHTML = "";
			opt.value = i;
			opt.innerHTML = i;
			select.appendChild(opt);
		}
	}
}

function removeOptions(selectbox)
{
    var i;
    for(i = selectbox.options.length - 1 ; i >= 0 ; i--)
    {
        selectbox.remove(i);
    }
}
//using the function:
	function hanyaAngka(evt) {
	  // var charCode = (evt.which) ? evt.which : event.keyCode
	  //  if (charCode > 31 && (charCode < 48 || charCode > 57))

	  //   return false;
	  // return true;
	  evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}
</script>
<script type="text/javascript">
	function hitungrumus() {
		var samount = $("#amount").val();
		var sadditionalvalue = $("#additionalvalue").val();

		var newsacquisitionvalue = samount.replace(/,/g, '');
		var newsadditionalvalue = sadditionalvalue.replace(/,/g, '');
		
		//Untuk menghitung Total 
		totalvalue = parseInt(newsacquisitionvalue) + parseInt(newsadditionalvalue);
		// alert(totalvalue);
		// if (!isNaN(totalvalue)) {
    	 	$("#total").val(totalvalue);
    	// }
    	// else{
    		// $("#total").val();
    	// }
	}
	
	function hitungrumusDepr() {
		var stotal = $("#total").val();
		var sdepreciationrange = $("#depreciationrange").val();
		
		var newstotal = stotal.replace(/,/g, '');
  
		totalvalue = parseInt(newstotal) / parseInt(sdepreciationrange);
		totalvalueyear = parseInt(totalvalue) * 12;
		
		// alert(newstotal);
		// if (!isNaN(totalvalue)) {
    	 	$("#monthlydepreciation").val(totalvalue);
			$("#yearlydepreciation").val(totalvalueyear);
			$("#bookvalue").val(newstotal);
			
    	// }
    	// else{
    		// $("#total").val();
    	// }
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
var getAssetTaxInfo = function () {
	// build up the post string when passing variables to the server side page
	
	var assettaxcode = document.getElementById("assettaxcode").value;
	var PostStr = 'assettaxcode='+assettaxcode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getassettaxgroupinfo.asp', 'POST', PostStr, showAssetTaxInfo);
}



// The function for handling the response from the server
var showAssetTaxInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.salescontractno);
    // update the Div to show the result from the server
	document.getElementById("assettaxname").value = s.assettaxname;
	// document.getElementById("assetgroupname").value = s.assetgroupname;
	// document.getElementById("unitofmeasurement").value = s.unitofmeasurement;
	// document.getElementById("depreciationrange").value = s.depreciationrange;
}

// launched from button click 
var getAssetInfo = function () {
	// build up the post string when passing variables to the server side page
	
	var assetcode = document.getElementById("assetcode").value;
	var PostStr = 'assetcode='+assetcode;
//	alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getassetandassetgroupinfo.asp', 'POST', PostStr, showAssetInfo);
}

// The function for handling the response from the server
var showAssetInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.assetname);
    // update the Div to show the result from the server
	document.getElementById("assetgroupcode").value = s.assetgroupcode;
	document.getElementById("assetgroupname").value = s.assetgroupname;
	document.getElementById("assettaxcode").value = s.assettaxcode;
	document.getElementById("assettaxname").value = s.assettaxname;
	document.getElementById("unitofmeasurement").value = s.unitofmeasurement;
	document.getElementById("assetname").value = s.assetname;
}

// launched from button click 
var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbaasset.asp', 'POST', PostStr, showlistfcba);
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
var assetgroupcodes = function () {
	// build up the post string when passing variables to the server side page
	
	var assetgroupcode = document.getElementById("assetgroupcode").value;
	var PostStr = 'assetgroupcode='+assetgroupcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getDepreciationmethodassetgroup.asp', 'POST', PostStr, showassetgroupcodes);
}

// The function for handling the response from the server
var showassetgroupcodes = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	document.getElementById("depreciationrange").value = s.depreciationrange;
}

// launched from button click 
var assettaxcodes = function () {
	// build up the post string when passing variables to the server side page
	
	var assettaxcode = document.getElementById("assettaxcode").value;
	var PostStr = 'assettaxcode='+assettaxcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getDepreciationmethodassettaxgroup.asp', 'POST', PostStr, showassettaxcodes);
}

// The function for handling the response from the server
var showassettaxcodes = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	document.getElementById("depreciationrange").value = s.depreciationrange;
}

// launched from button click 
var getAcquisitionValue = function () {
	// build up the post string when passing variables to the server side page
	
	var fccode = document.getElementById("fccode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fixedassetcode='+fccode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfixedassettransinfo.asp', 'POST', PostStr, showAcquisitionValue);
}

// The function for handling the response from the server
var showAcquisitionValue = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.salescontractno);
    // update the Div to show the result from the server
	document.getElementById("amount").value = s.amount;
	// document.getElementById("assetgroupname").value = s.assetgroupname;
	// document.getElementById("unitofmeasurement").value = s.unitofmeasurement;
	// document.getElementById("depreciationrange").value = s.depreciationrange;
}

function GetDepreciationRange(){
	var depreciationmethodby = document.getElementById("depreciationmethodby").value;
	if (depreciationmethodby === 'ASSET GROUP') {
		assetgroupcodes();
		// alert("1");
	} else {
		assettaxcodes();
		// alert("2");
	}
}

var reloadFACode = function () {
	var podate =$( ".datepicker" ).val();
	var selectedfcba =$( "#selectedfcba" ).val();
	var fccompanycode = $( "#fccompanycode" ).val();
	var selectedassetcode = $( "#assetcode" ).val();
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Finance and Accounting&txtTable=fixedasset_trans&transdate="+podate+"&selectedfcba="+selectedfcba+"&fccompanycode="+fccompanycode+"&selectedassetcode="+selectedassetcode );
}

</script>

<script language="JavaScript">
        function PopupIssue(_fcba) {
		var selectedfcba =$( "#selectedfcba" ).val();
		var newwin   = window.open("popup_issue_fa.Asp?fcba="+selectedfcba+"&Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
               
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
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")


searchdata = request("searchdata")
fccode = request("fccode")
fcname = request("fcname")
transdate = request("transdate")

assetcode = request("assetcode")
selectedassetcode = request("selectedassetcode")
assetname = request("assetname")
assetgroupcode = request("assetgroupcode")
assetgroupname = request("assetgroupname")
assettaxcode = request("assettaxcode")
assettaxname = request("assettaxname")

unitofmeasurement = request("unitofmeasurement")
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = request("selectedfcba")
depreciationmethod_by = request("depreciationmethod_by")
depreciationmonth = request("depreciationmonth")
total = request("total")
acquisitionvalue = request("acquisitionvalue")
additionalvalue = request("additionalvalue")
description = request("description")
date_transaction = request("date_transaction")
fctype = request("fctype")
fctype_adjustment = request("fctype_adjustment")
monthlydepreciation = request("monthlydepreciation")
yearlydepreciation = request("yearlydepreciation")
isactive = request("isactive")
faadjustmentcode = request("faadjustmentcode")
adjustment_amount = request("adjustment_amount")
adjustment_monthdepreciation = request("adjustment_monthdepreciation")
doc_no = request("doc_no")
monthtobe = request("monthtobe")
depreciationmonth_ledger = request("depreciationmonth_ledger")
lastdepreciationmonth = request("lastdepreciationmonth")
facodeadjustment = request("facodeadjustment")
valid_until = request("valid_until")
totaldepreciation = request("totaldepreciation")
bookvalue         = request("bookvalue")
sivcode           = request("sivcode")
kodebarang        = request("kodebarang")
namabarang        = request("namabarang")


if ""&transdate = "" then transdate = TextDate(date)
if ""&valid_until = "" then valid_until = TextDate(date)

fcCode = request("fccOde")
If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"

if Request("txtJenis")="New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
		
		q	= "DELETE FROM Fixedasset_Depreciation_Ledger WHERE fccode = '"& fcCode &"' AND lineno = '"& lineno &"' AND fcba = '"& fcba &"'"
		ExecuteAndLog ScriptName, q
		
		q   = "DELETE FROM fixedasset WHERE fccode = '" & fccode & "' AND fcba = '"& fcba &"'"
		ExecuteAndLog ScriptName, q
		
		cmdView	= "True"

	elseif Request("txtJenis")= "Delete" then
		strFieldDetail = " SELECT lineno FROM Fixedasset_Depreciation_Ledger " & _
                         " WHERE fccode = '" & fccode & "'" & _
                         "       AND fcBA = '" & fcba & "'"								
		' Response.Write strFieldDetail
		Set rsQty = Server.CreateObject("ADODB.Recordset")
		rsQty.Open strFieldDetail, connstr							
		if not rsQty.Eof then
			lineno = cdbl (to_nol(rsQty("lineno")))
		End If

		q	=	"DELETE FROM Fixedasset_Depreciation_Ledger WHERE fccode = '"& fcCode &"' AND lineno = '"& lineno &"' AND fcba = '"& fcba &"'"
		' response.write q
		ExecuteAndLog ScriptName, q
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select fccode from fixedasset where fccode = '" & fccode & "' and fcba ='" & fcba & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then
						q = " INSERT INTO fixedasset ( fccode" & _
							"                           , fcname" & _
							"                           , transdate" & _
							"                           , assetcode" & _
							"                           , assetname" & _
							"                           , assetgroupcode" & _
							"                           , assetgroupname" & _
							"                           , assettaxcode" & _
							"                           , assettaxname" & _
							"                           , uom" & _
							"                           , depreciationmethod_by" & _
							"                           , depreciationmonth" & _
							"                           , acquisitionvalue" & _
							"                           , additionalvalue" & _
							"                           , total" & _
							"                           , description" & _
							"                           , isactive" & _
							"                           , valid_until" & _
							"                           , fcentry" & _
							"                           , fcedit" & _
							"                           , fcip" & _
							"                           , fcba" & _
							"                           , lastupdate" & _
							"                           , lasttime" & _
							"                           , amountdepreciation" & _
							"                           , totaldepreciation" & _
							"                           , bookvalue" & _
							"                           , lastdepreciationmonth" & _
							"                           , saldoawal_totaldepreciation" & _
							"                           , saldoawal_bookvalue" & _
							"                           , saldoawal_depreciationmonth" & _
							"                           , saldoawal_month" & _
							"                           , saldoawal_year" & _
							"                           , sivcode" & _
							"                           , stockcode" & _
							")" & _
							" VALUES ( '" & fccode & "'" & _
							"          , '" & fcname & "'" & _
							"          , TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
							"          , '" & assetcode & "'" & _
							"          , '" & assetname & "'" & _
							"          , '" & assetgroupcode & "'" & _
							"          , '" & assetgroupname & "'" & _
							"          , '" & assettaxcode & "'" & _
							"          , '" & assettaxname & "'" & _
							"          , '" & unitofmeasurement & "'" & _
							"          , '" & depreciationmethod_by & "'" & _
							"          , '" & depreciationmonth & "'" & _
							"          , replace('" & acquisitionvalue & "',',','')" & _
							"          , replace('" & additionalvalue & "',',','')" & _
							"          , replace('" & total & "',',','')" & _
							"          , '" & description & "'" & _
							"          , '" & isactive & "'" & _
							"          , TO_DATE('" & valid_until & "', 'mm/dd/yyyy')" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & fcba & "'" & _
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE, 'hh24:mi')" & _
							"          , replace('" & monthlydepreciation & "',',','')" & _
							"          , replace('" & totaldepreciation & "',',','')" & _ 
							"          , replace('" & bookvalue & "',',','')" & _ 
							"          , replace('" & monthtobe & "',',','')" & _ 
							"          , replace('" & sa_tot_depre & "',',','')" & _ 
							"          , replace('" & sa_book_value & "',',','')" & _ 
							"          , '" & sa_month_depre & "'" & _ 
							"          , '" & sa_month & "'" & _ 
							"          , '" & sa_year & "'" & _ 
							"          , '" & sivcode & "'" & _
							"          , '" & kodebarang & "'" & _
							")"			 

						' response.write q
						ExecuteAndLog ScriptName, q
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing
	
	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q = " UPDATE fixedasset_transdet" & _
            " SET  date_transaction = TO_DATE('" & date_transaction & "', 'mm/dd/yyyy')" & _
			"     , acquisitionvalue_type = '" & acquisitionvalue_type & "'" & _
			"     , monthlydepreciation = '" & monthlydepreciation & "'" & _
			"     , yearlydepreciation = '" & yearlydepreciation & "'" & _
			"	  , amountdepreciation = replace('" & monthlydepreciation & "',',','')" & _
			"	  , totaldepreciation = replace('" & totaldepreciation & "',',','')" & _ 
			"	  , bookvalue = replace('" & bookvalue & "',',','')" & _ 
			"	  , lastdepreciationmonth = replace('" & monthtobe & "',',','')" & _ 
			"	  , saldoawal_totaldepreciation = replace('" & sa_tot_depre & "',',','')" & _ 
			"	  , saldoawal_bookvalue = replace('" & sa_book_value & "',',','')" & _ 
			"	  , saldoawal_depreciationmonth = '" & sa_month_depre & "'" & _ 
			"	  , saldoawal_month = '" & sa_month & "'" & _ 
			"	  , saldoawal_year = '" & sa_year & "'" & _ 
			"	  , sivcode = '" & sivcode & "'" & _
			"	  , stockcode = '" & kodebarang & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , fcba = '" & fcba & "'" & _
			"     , lastupdate = SYSDATE" & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
			" WHERE fccode = '" & fccode & "' AND fcba = '"& fcba &"' AND lineno = '"& lineno &"'" 
		' response.write q
		ExecuteAndLog ScriptName, q
		cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")


'Remarks By Iwa Edit date_transaction di rubah sesuai creation date

' If Len(Trim(date_transaction)) = 0 Then	date_transaction = textDate(date)

If Len(Trim(date_transaction)) = 0 Then
	
		date_transaction = transdate
		
End if

' End Remarks By Iwa Edit date_transactiondi rubah sesuai creation date

Dim rsFaHeader, strFaHeader
strFaHeader	=	" SELECT fa.fccode," & _
				"       bu.fccompanycode," & _
				"       fa.fcba," & _
				"       fa.fcname," & _
				"       fa.transdate," & _
				"       fa.assetcode," & _
				"       fa.assetname," & _
				"       fa.assetgroupcode," & _
				"       fa.assetgroupname," & _
				"       fa.assettaxcode," & _
				"       fa.assettaxname," & _
				"       fa.uom," & _
				"       fa.depreciationmethod_by," & _
				"       fa.depreciationmonth," & _
				"       fa.acquisitionvalue," & _
				"       fa.additionalvalue," & _
				"       fa.total," & _
				"       fa.description," & _
				"       fa.isactive," & _
				"       fa.valid_until," & _
				"       fa.sivcode," & _
				"       fa.stockcode," & _
				"       stock_api.get_stockname(fa.stockcode, fa.fcba) stockname " & _ 
				"  FROM fixedasset fa" & _
				"       LEFT JOIN asset ass ON (ass.fccode = fa.assetcode)" & _
				"       LEFT JOIN assetgroup asg ON (asg.fccode = fa.assetgroupcode)" & _
				"       LEFT JOIN fixedasset_taxgroup atg ON (atg.fccode = fa.assettaxcode)" & _
				"       LEFT JOIN businessunit bu ON (bu.fccode = fa.fcba)" & _
				" WHERE fa.fccode = '"& fcCode &"' AND fa.fcba = '"& fcba &"'"

						' response.write strFaHeader
set rsFaHeader= server.CreateObject("Adodb.recordset")
rsFaHeader.Open strFaHeader, connstr
if Not rsFaHeader.Eof then
	fccode = rsFaHeader("fccode")
	fcname = rsFaHeader("fcname")
	fccompanycode = rsFaHeader("fccompanycode")
	fcba = rsFaHeader("fcba")
	selectedfcba = rsFaHeader("fcba")
	transdate = rsFaHeader("transdate")
	assetcode = rsFaHeader("assetcode")
	assetname = rsFaHeader("assetname")
	assetgroupcode = rsFaHeader("assetgroupcode")
	assetgroupname = rsFaHeader("assetgroupname")
	description = rsFaHeader("description")
	unitofmeasurement = rsFaHeader("uom")
	isactive = rsFaHeader("isactive")
	valid_until = rsFaHeader("valid_until")
	depreciationmethod_by = rsFaHeader("depreciationmethod_by")
	depreciationmonth = rsFaHeader("depreciationmonth")
	acquisitionvalue = rsFaHeader("acquisitionvalue")
	total = rsFaHeader("total")
	assettaxcode = rsFaHeader("assettaxcode")
	assettaxname = rsFaHeader("assettaxname")
	acquisitionvalue = rsFaHeader("acquisitionvalue")
	additionalvalue = rsFaHeader("additionalvalue")
	selectedassetcode = rsFaHeader("assetcode")
End if
'response.Write strFaHeader
rsFaHeader.Close
Set rsFaHeader = Nothing

if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if


oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('bank')"
			
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
 '-------------ISI Payment Form
  %>
  
<form name="frmEntry" method="post" action="fixedasset_trans.asp">
	<table class="data" width="100%">
	<tr>
		<td width="20%" align="left">Fixed Asset Code</td>
		<td align="left">
		<%
			if fcCode = "(NEW)" then
				q = " select '" & selectedfcba & "'||'/'||'" & selectedassetcode & "'||'/'||TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'yyyy')||'/'||TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'mm')||'/'||LPAD(NVL(MAX(regexp_substr(cn.fccode, '[^/]+', 1, 5)),0000)+0001,4,0) DOCNUMBER " & _
					" from " & _
					" fixedasset cn " & _
					" where cn.fcba ='" & selectedfcba & "' " & _
					" and cn.assetcode ='" & selectedassetcode & "' "&_
					" and TO_NUMBER(TO_CHAR(cn.transdate,'yyyy')) =  TO_NUMBER(TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'yyyy')) " & _
					" and TO_NUMBER(TO_CHAR(cn.transdate,'mm')) = TO_NUMBER(TO_CHAR(TO_DATE('"& transdate &"','mm/dd/yyyy'),'mm')) "
				'Response.write q
				set rsPONo = server.CreateObject("Adodb.recordset")
				rsPONo.Open q, connstr
				if Not rsPONo.Eof then
					fcCode = rsPONo("DOCNUMBER")
				End if
			End if
		%>
			<input type="text" size="30" maxlength="50" id="fccode" name="fccode" value="<%=fccode%>"  
			<%IF Request("txtJenis")="Edit" then%> readonly="True" style="background-color:#ccc;"<%end if%>>
			<font color="#FF0000">*</font>
		</td>
		
		
		<td align="left">Asset Tax Group Code</td>
		<td align="left"><input type="text" size="50" id="assettaxcode" name="assettaxcode" value="<%=assettaxcode%>" readonly="true"  style="background-color:#ccc;"></td>
	</tr>
	<tr>
	    <td align="left">Fixed Asset Name</td>
		<td align="left"><input type="text" size="50" maxlength="150" id="fcname" name="fcname" value="<%=fcname%>"><font color="#FF0000">*</font></td>
		
		
		<td align="left">Asset Tax Group Name</td>
		<td align="left"><input type="text" size="50" id="assettaxname" name="assettaxname" value="<%=assettaxname%>" readonly="true"  style="background-color:#ccc;"></td>
		
	</tr>
	<tr>
		<td align="left">Company</td>
		<td align="left">
			<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();" class="chosen-select" style="width: 40%">
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
		<td align="left">Asset Group Code</td>
		<td align="left"><input type="text" size="50" id="assetgroupcode" name="assetgroupcode" value="<%=assetgroupcode%>" readonly="true"  style="background-color:#ccc;"></td>
	</tr>
	<tr>
		<td align="left">Business Unit</td>
		<td align="left">
			<div id="fcbalist"></div>
			<font color="#FF0000">*</font>
		</td>
		<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
		<td align="left">Asset Group Name</td>
		<td align="left"><input type="text" size="50" id="assetgroupname" name="assetgroupname" value="<%=assetgroupname%>" readonly="true"  style="background-color:#ccc;"></td>
	</tr>		
	<tr>
		<td align="left">Creation Date</td>
		<td align="left">
			<input class="datepicker" type="text" maxlength="10" name="transdate" size="10" value="<%=transdate%>" onchange="javascript:reloadFACode();" autocomplete="off">
			<font color="#FF0000">&nbsp;*</font>
		</td>
		<td align="left">Uom</td>
		<td align="left">
			<input type="text" size="50" id="unitofmeasurement" name="unitofmeasurement" value="<%=unitofmeasurement%>" readonly="true"  style="background-color:#ccc;">
		</td>
	</tr>
	<tr>
		
		<td align="left">Asset Code</td>
			<td align="left">
				<select id="assetcode" name="assetcode" onchange="javascript:reloadFACode()" class="chosen-select" style="width: 70%">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM asset ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if selectedassetcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedassetcode" name="selectedassetcode" value="<%=selectedassetcode%>">
		
		
		<td align="left">Acquisition Value</td>
		<td align="left">
			<input type="text" size="50" id="amount" name="acquisitionvalue"    value="<%=acquisitionvalue%>" onkeyup="hanyaAngka(this); hitungrumus();hitungrumusDepr();" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
		</td>
	</tr>	
	<tr>
		<td align="left">Asset Name</td>
		<td align="left"><input type="text" size="50" id="assetname" name="assetname" value="<%=assetname%>" readonly="true"  style="background-color:#ccc;"></td>
		<td align="left">Adjustment Value</td>
		<td align="left">
			<input type="text" size="50" id="additionalvalue" name="additionalvalue"   value="<%=additionalvalue%>" onkeyup="hanyaAngka(this); hitungrumus();hitungrumusDepr();" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
		</td>
	</tr>
	<tr>
		<td align="left">Depreciation Method By</td>
		<td align="left">
			<select id="depreciationmethodby" name="depreciationmethod_by" onchange="GetDepreciationRange();" class="chosen-select" style="width: 70%">
				<option></option>
				<option value="ASSET GROUP" <%if depreciationmethod_by = "ASSET GROUP" then%>selected<%end if%>>FROM ASSET GROUP</option>
				<option value="ASSET TAX GROUP" <%if depreciationmethod_by = "ASSET TAX GROUP" then%>selected<%end if%>>FROM ASSET TAX GROUP</option>
			</select>
			<font color="#FF0000">*</font>
		</td>
		<td align="left">Total Value</td>
			<td align="left">
			<input type="text" size="50" id="total" name="total" value="<%=total%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly>
		</td>
		</tr>
		<tr>
			<td align="left">Depreciation Month</td>
				<td align="left">
					<input type="text" size="50" id="depreciationrange" name="depreciationmonth" value="<%=depreciationmonth%>" readonly="true"  style="background-color:#ccc;" onkeyup="hanyaAngka(this); hitungrumus();">
				</td>
			</td>
			<td align="left">Description</td>
			<td align="left">
				<textarea name="description" cols="50" id="description"><%=description%></textarea>
			</td>
		</tr>	
		<tr>
			<td align="left">Valid Until</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="valid_until" size="10" value="<%=valid_until%>" autocomplete="off">
				<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left">Is Active</td>
				<td align="left">
				<select name="isactive" class="chosen-select" style="width: 20%">
					<option value="1" <%if isactive = "1" then%>selected<%end if%>>TRUE</option>
					<option value="0" <%if isactive = "0" then%>selected<%end if%>>FALSE</option>
				</select>
				<font color="#FF0000">*</font>
			</td>	
		</tr>
		<tr>
			<td colspan="4"><hr></td>
		<tr>
		<tr>
			<td align="left">Load From Issue</td>
				<td align="left">
					<input type="text" size="30" id="sivcode" name="sivcode" value="<%=sivcode%>" readonly="true"  style="background-color:#ccc;" >
					<a href="javascript:PopupIssue('<%=fcba%>')">
                    <font face="Verdana" size="1">Search</font>
					</a>
				</td>
			<td align="left">Asset Stock</td>
				<td align="left">
					<input type="text" size="15" id="kodebarang" name="kodebarang" value="<%=kodebarang%>" readonly="true"  style="background-color:#ccc;" >
					<input type="text" size="40" id="namabarang" name="namabarang" value="<%=namabarang%>" readonly="true"  style="background-color:#ccc;" >
			</td>	
		</tr>
		<tr>
			<td colspan="4"><hr></td>
		<tr>
		
		
<tr>
	<td colspan="2"  align="left">
	<input Type="Hidden" Value="True" Name="cmdView">
	<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
	<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
	<input Type="Submit"  name="Add" value="Select">&nbsp;
	
	<a href="fixedasset_trans.asp?fcCode=<%=fccode%>&transdate=<%=transdate%>&fcba=<%=fcba%>&txtJenis=HeaderDelete" onClick="return confirm('Are you sure to delete');">
		<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0">
	</a>
	<a href="FixedAsset_HeaderEdit.asp?fcba=<%=fcba%>&fcCode=<%=fcCode%>&fcname=<%=fcname%>&description=<%=description%>&transdate=<%=transdate%>" target="_blank">
		Edit Header
	</a>
	<a href="Printfixedasset_trans.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
	</td>
</tr>
</table>  
</form>
<%
if cmdView = "True" AND len(trim(fcCode)) > 0 AND len(trim(fcba)) > 0 Then
	q =	" SELECT bu.fccompanycode," & _
		"		fa.fcba," & _
		"		fa.fccode," & _
		"       fa.fcname," & _
		"		fad.fccode faadjustmentcode," & _
        "		fad.fctype fctype_adjustment," & _
        "		fad.adjustment_amount," & _
        "		fad.adjustment_monthdepreciation," & _
		"       fa.transdate," & _
		"       fa.assetcode," & _
		"       fa.assetname," & _
		"       fa.assetgroupcode," & _
		"       fa.assetgroupname," & _
		"       fa.assettaxcode," & _
		"       fa.assettaxname," & _
		"       fa.uom," & _
		"       fa.depreciationmethod_by," & _
		"       fa.depreciationmonth," & _
		"       fa.acquisitionvalue," & _
		"       fa.additionalvalue," & _
		"       fa.total," & _
		"       fa.description," & _
		"       fa.isactive," & _
		"       fal.date_transaction," & _
		"       fal.fctype," & _
		"       fal.doc_no," & _
		"       fa.amountdepreciation," & _
		"       depreciationmonth_ledger," & _
		"       fal.monthtobe," & _
		"       fal.lastdepreciationmonth," & _
		"       fal.monthlydepreciation," & _
		"       fal.yearlydepreciation," & _
		"       fal.lineno," & _
		"       fal.totaldepreciation," & _
		"       fal.bookvalue," & _		
		"       fa.saldoawal_totaldepreciation," & _		
		"       fa.saldoawal_bookvalue," & _		
		"       fa.saldoawal_depreciationmonth," & _
		"       fa.saldoawal_month," & _		
		"       fa.saldoawal_year," & _	
		"       fa.sivcode," & _	
		"       fa.stockcode" & _	
		"  FROM fixedasset fa" & _
		"       LEFT JOIN Fixedasset_Depreciation_Ledger fal" & _
		"          ON (fa.fccode = fal.fccode AND fa.fcba = fal.fcba)" & _
		" LEFT JOIN businessunit bu ON (bu.fcCode = fa.fcba)" & _
		" LEFT JOIN fixedasset_adjustment fad ON (fa.fccode = fad.fixedassetcode AND fa.fcba = fad.fcba)" & _
		"          WHERE fa.fccode = '"& fcCode &"'" & _
		"          AND fa.fcba = '"& fcba &"'" 

	'response.write q
	Set rsFixedasset = Server.CreateObject("ADODB.Recordset")
	rsFixedasset.Open q, connstr
	if Not rsFixedasset.eof then
		while not rsFixedasset.Eof
			faadjustmentcode = rsFixedasset("faadjustmentcode")
			fctype_adjustment = rsFixedasset("fctype_adjustment")
			adjustment_amount = rsFixedasset("adjustment_amount")
			adjustment_monthdepreciation = rsFixedasset("adjustment_monthdepreciation") 
			fccompanycode = rsFixedasset("fccompanycode")
			fcba = rsFixedasset("fcba")
			fixedassetcode = rsFixedasset("fccode")
			doc_no = rsFixedasset("doc_no")
			fcname = rsFixedasset("fcname")
			assetcode = rsFixedasset("assetcode")
			assetname = rsFixedasset("assetname")
			assetgroupcode = rsFixedasset("assetgroupcode")
			assetgroupname = rsFixedasset("assetgroupname")
			assettaxcode = rsFixedasset("assettaxcode")
			assettaxname = rsFixedasset("assettaxname")
			unitofmeasurement = rsFixedasset("uom")
			depreciationmethod_by = rsFixedasset("depreciationmethod_by")
			depreciationmonth = rsFixedasset("depreciationmonth")
			acquisitionvalue = rsFixedasset("acquisitionvalue")
			additionalvalue = rsFixedasset("additionalvalue")
			totalvalue = rsFixedasset("total")
			description = rsFixedasset("description")
			isactive = rsFixedasset("isactive")
			date_transaction = rsFixedasset("date_transaction")
			fctype = rsFixedasset("fctype")
			amountdepreciation = rsFixedasset("amountdepreciation")
			depreciationmonth_ledger = rsFixedasset("depreciationmonth_ledger")
			monthlydepreciation = rsFixedasset("monthlydepreciation")
			yearlydepreciation = rsFixedasset("yearlydepreciation")
			lineno = rsFixedasset("lineno")
			totaldepreciation = rsFixedasset("totaldepreciation")
			bookvalue = rsFixedasset("bookvalue")
			sa_tot_depre = rsFixedasset("saldoawal_totaldepreciation")
			sa_book_value= rsFixedasset("saldoawal_bookvalue")
			sa_month_depre = rsFixedasset("saldoawal_depreciationmonth")
			sa_month     = rsFixedasset("saldoawal_month")
			sa_year      = rsFixedasset("saldoawal_year")
			sivcode      = rsFixedasset("sivcode")
			kodebarang	 = rsFixedasset("stockcode")
			
			rsFixedasset.MoveNext
		wend
		isNew = "FALSE"
	else
		isNew = "TRUE"
	end if
	
	
	if isNEw = "TRUE" and cdbl(to_nol(total)) > 0 then
	    totalvalue = cdbl(to_nol(total))
	    bookvalue  = cdbl(to_nol(total))
		monthlydepreciation = cdbl(to_nol(total)) / cdbl(to_nol(depreciationmonth))
		yearlydepreciation = cdbl(to_nol(monthlydepreciation)) * 12
		sa_tot_depre  = cdbl(monthlydepreciation)
		sa_book_value = cdbl(to_nol(total))
	end if
	

	qLineMax = " select nvl(max(monthtobe),0) monthtobe" & _
			   " from Fixedasset_Depreciation_Ledger" & _
			   " where fcCode = '" & fccode & "'" & _
			   " and fcba = '" & fcba & "'" 			
	' response.write qLineMax
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		monthtobe = cdbl(rsqLineMax("monthtobe"))
	else
		monthtobe = 0
	end if 

	if monthtobe=0 then
		date_transaction =transdate
		monthtobe = monthtobe + 1
		txtmonthtobe = monthtobe
	    sa_month_depre = monthtobe
		
		sa_month = month(date_transaction)
		sa_year  = year(date_transaction)
	else
		monthtobe = monthtobe + 1
		txtmonthtobe = monthtobe
		date_transaction = DateAdd("m",monthtobe,transdate)
	
	end if
	
    totaldepreciation = cdbl(to_nol(monthlydepreciation)) * cdbl(to_nol(monthtobe))
	bookvalue = cdbl(to_nol(totalvalue)) - cdbl(to_nol(totaldepreciation))
	
	
	if len(trim(faadjustmentcode)) > 0 then
		lastdepreciationmonth = cdbl(to_nol(adjustment_monthdepreciation)) - cdbl(to_nol(txtmonthtobe))
	else
		lastdepreciationmonth = cdbl(to_nol(depreciationmonth)) - cdbl(to_nol(txtmonthtobe))
	end if
	
	
%>
   
   <!-- Form Entry And Edit --> 
   
   <form action="fixedasset_trans.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
			<tr>
				<td align="left">
					Fixed Asset Code
				</td>
				<td align="left">
					<input type="text" size="50" id="fixedassetcode" name="fixedassetcode" value="<%=fccode%>" readonly="true"  style="background-color:#ccc;">
				</td>
			</tr>
			<tr>
				<td align="left">Date</td>
				<td align="left">
					<input  type="text" maxlength="10" name="date_transaction" size="10" value="<%=date_transaction%>" autocomplete="off"  style="background-color:#ccc;" >
				</td>
			</tr>
			<%	if len(trim(faadjustmentcode)) > 0 then %>
				<tr>
				<td align="left">Type</td>
					<td align="left">
						<input type="text" size="50" name="fctype" value="FROM ADJUSTMENT" readonly="true"  style="background-color:#ccc;">
					</td>
				</tr>
				<tr>
					<td align="left">
						No Doc
					</td>
					<td align="left">
						<input type="text" size="50" name="doc_no" value="<%=faadjustmentcode%>" >
					</td>
				</tr>
				<tr>
					<td align="left">Value</td>
						<td align="left">
							<input type="text" size="50"  name="amountdepreciation" value="<%=adjustment_amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
						</td>
					</td>
				</tr>
				<tr>
					<td align="left">Depreciation Month</td>
						<td align="left">
							<input type="text" size="50"  name="depreciationmonth_ledger" value="<%=adjustment_monthdepreciation%>" style="background:#ccc"  >
						</td>
					</td>
				</tr>
			<%else%>
				<tr>
				<td align="left">Type</td>
					<td align="left">
						<input type="text" size="50" name="fctype" value="FROM STRAIGHT LINE" readonly="true"  style="background-color:#ccc;">
					</td>
				</tr>
				<tr>
					<td align="left">
						No Doc
					</td>
					<td align="left">
						<input type="text" size="50" id="doc_no" name="doc_no" value="<%=fccode%>" readonly="true"  style="background-color:#ccc;">
					</td>
				</tr>
				<tr>
					<td align="left">Value</td>
						<td align="left">
							<input type="text" size="50" id="amountdepreciation" name="amountdepreciation" value="<%=acquisitionvalue%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
						</td>
					</td>
				</tr>
				<tr>
					<td align="left">Depreciation Month</td>
						<td align="left">
							<input type="text" size="50" id="depreciationmonth_ledger" name="depreciationmonth_ledger" value="<%=depreciationmonth%>" style="background:#ccc"  >
						</td>
					</td>
				</tr>
			<%end if%>
			
			<tr>
				<td align="left">Start Month</td>
					<td align="left">
						<input type="text" size="50" id="monthtobe" name="monthtobe" value="<%=txtmonthtobe%>"  >
					</td>
				</td>
			</tr>
			<tr>
				<td align="left">End Month</td>
					<td align="left">
						<input type="text" size="50" id="lastdepreciationmonth" name="lastdepreciationmonth" value="<%=lastdepreciationmonth%>" style="background:#ccc" >
					</td>
				</td>
			</tr>
			<tr>
				<td align="left">Value Depreciation/Month</td>
					<td align="left">
						<input type="text" size="50" id="monthlydepreciation" name="monthlydepreciation" value="<%=monthlydepreciation%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly>
					</td>
				</td>
			</tr>
			<tr>
				<td align="left">Value Depreciation/Year</td>
					<td align="left">
						<input type="text" size="50" id="yearlydepreciation" name="yearlydepreciation" value="<%=yearlydepreciation%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly>
					</td>
				</td>
			</tr>
			
			<tr>
				<td align="left">Total.Depreciation </td>
					<td align="left">
						<input type="text" size="50" id="totaldepreciation" name="totaldepreciation" value="<%=totaldepreciation%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
					</td>
				</td>
			</tr>
			<tr>
				<td align="left">Book Value</td>
					<td align="left">
						<input type="text" size="50" id="bookvalue" name="bookvalue" value="<%=bookvalue%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
					</td>
				</td>
			</tr>
			
			<tr>
				<td align="left">Begining Tot.Depreciation</td>
					<td align="left">
						<input type="text" size="50" id="sa_tot_depre" name="sa_tot_depre" value="<%=sa_tot_depre%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
					</td>
				</td>
			</tr>
			<tr>
				<td align="left">Begining Book Value</td>
					<td align="left">
						<input type="text" size="50" id="sa_book_value" name="sa_book_value" value="<%=sa_book_value%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
					</td>
				</td>
			</tr>
			<tr>
				<td align="left">Begining Tot.Month</td>
					<td align="left">
						<input type="text" size="50" id="sa_month_depre" name="sa_month_depre" value="<%=sa_month_depre%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  >
					</td>
				</td>
			</tr>
			
			<tr>
				<td align="left">Begining Periode</td>
					<td align="left">
						<input type="text" size="15" maxlength="2" id="sa_month" name="sa_month" value="<%=sa_month%>"  >
						<input type="text" size="15" maxlength="4" id="sa_year" name="sa_year"   value="<%=sa_year%>"   >
					</td>
				</td>
			</tr>
			
		<tr>
			<td colspan="2" align="left">
				<input type="Hidden" name="searchdata" value="<%=searchdata%>">
			  	<input Type="Hidden" Value="<%=request("transdate")%>" Name="transdate">
				<input Type="Hidden" Value="<%=request("fcba")%>" Name="fcba">
				<input Type="Hidden" Value="<%=request("selectedfcba")%>" Name="selectedfcba">
				<input Type="Hidden" Value="<%=request("fccompanycode")%>" Name="fccompanycode">
				<input Type="Hidden" Value="<%=request("fcCode")%>" Name="fcCode">
				<input Type="Hidden" Value="<%=request("fcname")%>" Name="fcname">
				<input Type="Hidden" Value="<%=request("assettaxcode")%>" Name="assettaxcode">
				<input Type="Hidden" Value="<%=request("assettaxname")%>" Name="assettaxname">
				<input Type="Hidden" Value="<%=request("depreciationmethod_by")%>" Name="depreciationmethod_by">
				<input Type="Hidden" Value="<%=request("depreciationmonth")%>" Name="depreciationmonth">
				<input Type="Hidden" Value="<%=request("assetcode")%>" Name="assetcode">
				<input Type="Hidden" Value="<%=request("assetname")%>" Name="assetname">
				<input Type="Hidden" Value="<%=request("assetgroupcode")%>" Name="assetgroupcode">
				<input Type="Hidden" Value="<%=request("assetgroupname")%>" Name="assetgroupname">
				<input Type="Hidden" Value="<%=request("unitofmeasurement")%>" Name="unitofmeasurement">
				<input Type="Hidden" Value="<%=request("acquisitionvalue")%>" Name="acquisitionvalue">
				<input Type="Hidden" Value="<%=request("additionalvalue")%>" Name="additionalvalue">
				<input Type="Hidden" Value="<%=request("total")%>" Name="total">
				<input Type="Hidden" Value="<%=request("valid_until")%>" Name="valid_until">
				<input Type="Hidden" Value="<%=request("description")%>" Name="description">
				<input Type="Hidden" Value="<%=request("isactive")%>" Name="isactive">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				
				<input Type="Hidden" Value="<%=request("sivcode")%>" Name="sivcode">
				<input Type="Hidden" Value="<%=request("kodebarang")%>" Name="kodebarang">
				<input Type="Hidden" Value="<%=request("namabarang")%>" Name="namabarang">
				
				
				<!-- Start Val Is True Vaidaton -->
				<%IF Request("txtJenis")="Edit" then%>
					   <input Type="Hidden" Value="<%=lineno%>" Name="lineno">
					   <input type="hidden" name="txtJenis" value="Edit">
					   <input Type="Submit"  name="Add" value="Update">
				  <%Else%>
					   <input type="hidden" name="txtJenis" value="New">
					   <input Type="Submit"  name="Add" value="Add">
				  <%End if%>
				  <input Type="Reset"  name="Reset" value="Reset">
				  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				<!-- End Val Is True Vaidaton -->
			</td>
		</tr>
   </table>
   </form>   
   <table class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<TD align="left">No</TD>
		<TD align="center"></TD>
		<TD align="left">Company</TD>
		<TD align="left">Business Unit</TD>
		<TD align="left">Fixed Asset Code</TD>
		<TD align="left">Fixed Asset Name</TD>
		<TD align="left">Doc No</TD>
		<TD align="left">Creation Date</TD>
		<TD align="left">Asset Code</TD>
		<TD align="left">Asset Name</TD>
		<TD align="left">Asset Group Code</TD>
		<TD align="left">Asset Group Name</TD>
		<TD align="left">Asset Tax Code</TD>
		<TD align="left">Asset Tax Name</TD>
		<TD align="left">Uom</TD>
		<TD align="left">Depreciation Method By</TD>
		<TD align="left">Depreciation Range</TD>
		<TD align="left">Month To Be</TD>
		<TD align="left">Last Depreciation Month</TD>
		<TD align="left">Amount To Be Depreciate</TD>
		<TD align="left">Monthly Depreciation</TD>
		<TD align="left">Total Depreciation</TD>
		<TD align="left">Book Value</TD>
		<TD align="left">Yearly Depreciation </TD>
		<TD align="left">Acquisition Value</TD>
		<TD align="left">Additional Value</TD>
		<TD align="left">Total Value</TD>
		<TD align="left">Description</TD>
		<TD align="left">Is Active</TD>
	</tr>
<%	Set rsGAD = Server.CreateObject("ADODB.Recordset")
    rsGAD.Open q, connstr
	if Not rsGAD.eof then
		nCountofField = rsGAD.fields.count - 1
		no = 1
		While Not rsGAD.EOF
			if no  mod 2 = 1 then cColor = "#EEEEEE" else cColor = "#E5E5E5"
%>
			<tr bgcolor="<%=cColor%>">
				<td align="center"><%=NO%></td>
				
				<td valign="middle">
					<a href="fixedasset_trans.asp?fccode=<%=rsGAD("fccode")%>&lineno=<%=rsGAD("lineno")%>&date_transaction=<%=rsGAD("date_transaction")%>&fcba=<%=rsGAD("fcba")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD("fccode")%>" border="0">
					</a>
				</td>
		   		
				<td align="center" title="<%=rsGAD("fccompanycode").Name%>"><%=rsGAD("fccompanycode")%></td>
				<td align="center" title="<%=rsGAD("fcba").Name%>"><%=rsGAD("fcba")%></td>
				<td align="center" title="<%=rsGAD("fccode").Name%>"><%=rsGAD("fccode")%></td>
				<td align="left" title="<%=rsGAD("fcname").Name%>"><%=rsGAD("fcname")%></td>
				<td align="left" title="<%=rsGAD("doc_no").Name%>"><%=rsGAD("doc_no")%></td>
				<td align="center" title="<%=rsGAD("date_transaction").Name%>"><%=rsGAD("date_transaction")%></td>
				<td align="center" title="<%=rsGAD("assetcode").Name%>"><%=rsGAD("assetcode")%></td>
				<td align="left" title="<%=rsGAD("assetname").Name%>"><%=rsGAD("assetname")%></td>
				<td align="center" title="<%=rsGAD("assetgroupcode").Name%>"><%=rsGAD("assetgroupcode")%></td>
				<td align="left" title="<%=rsGAD("assetgroupname").Name%>"><%=rsGAD("assetgroupname")%></td>
				<td align="center" title="<%=rsGAD("assettaxcode").Name%>"><%=rsGAD("assettaxcode")%></td>
				<td align="left" title="<%=rsGAD("assettaxname").Name%>"><%=rsGAD("assettaxname")%></td>
				<td align="center" title="<%=rsGAD("uom").Name%>"><%=rsGAD("uom")%></td>
				<td align="center" title="<%=rsGAD("depreciationmethod_by").Name%>"><%=rsGAD("depreciationmethod_by")%></td>
				<td align="center" title="<%=rsGAD("depreciationmonth_ledger").Name%>"><%=rsGAD("depreciationmonth_ledger")%></td>
				<td align="center" title="<%=rsGAD("monthtobe").Name%>"><%=rsGAD("monthtobe")%></td>
				<td align="center" title="<%=rsGAD("lastdepreciationmonth").Name%>"><%=rsGAD("lastdepreciationmonth")%></td>
				<td align="right" title="<%=rsGAD("amountdepreciation").Name%>"><%=FormatNumber(cdbl(to_nol(rsGAD("amountdepreciation"))),2)%></td>
				<td align="right" title="<%=rsGAD("monthlydepreciation").Name%>"><%=FormatNumber(cdbl(to_nol(rsGAD("monthlydepreciation"))),2)%></td>
				<td align="right" title="<%=rsGAD("totaldepreciation").Name%>"><%=FormatNumber(cdbl(to_nol(rsGAD("totaldepreciation"))),2)%></td>
				<td align="right" title="<%=rsGAD("bookvalue").Name%>"><%=FormatNumber(cdbl(to_nol(rsGAD("bookvalue"))),2)%></td>				
				<td align="right" title="<%=rsGAD("yearlydepreciation").Name%>"><%=FormatNumber(cdbl(to_nol(rsGAD("yearlydepreciation"))),2)%></td>
				<td align="right" title="<%=rsGAD("acquisitionvalue").Name%>"><%=FormatNumber(cdbl(to_nol(rsGAD("acquisitionvalue"))),2)%></td>
				<td align="right" title="<%=rsGAD("additionalvalue").Name%>"><%=FormatNumber(cdbl(to_nol(rsGAD("additionalvalue"))),2)%></td>
				<td align="right" title="Total Value"><%=FormatNumber( cdbl(to_nol(rsGAD("acquisitionvalue")))+cdbl(to_nol(rsGAD("additionalvalue"))), 2)%></td>
				<td align="left" title="<%=rsGAD("description").Name%>"><%=rsGAD("description")%></td>
				<td align="center" title="<%=rsGAD("isactive").Name%>"><%=rsGAD("isactive")%></td>

			</tr>
<%
			rsGAD.MoveNext
			intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
		Wend
	end if 
	rsGAD.Close : set rsGAD = Nothing
%>
   </div>
   </TABLE>
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

function getselectedassetcode() {
	var sfcba = document.getElementById("selectedassetcode");
	sfcba.value = document.getElementById("assetcode").value;
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
	getselectedassetcode();
	getAssetInfo();
	getlistfcba();
	getselectedfcba();
	hitungrumus();
	//getAcquisitionValue();
	hitungrumusDepr();
	// document.getElementById('monthlydepreciation').value = 0;
});
</script>
</body>
</html>
