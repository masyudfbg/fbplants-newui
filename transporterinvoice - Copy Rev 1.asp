<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Transporter Invoice</title>

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

var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbatransporterinvoice.asp', 'POST', PostStr, showlistfcba);
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

var getlisttransportercontract = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedtransporter = document.getElementById("selectedtransporter").value;
	var PostStr = 'fcba='+selectedfcba+'&transportercontractno='+selectedtransporter;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlisttransportercontract.asp', 'POST', PostStr, showlisttransportercontract);
}

// The function for handling the response from the server
var showlisttransportercontract = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("transportercontractlist").innerHTML = "";
	document.getElementById("transportercontractlist").innerHTML = response;

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

var getlistinvoice = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedinvoice = document.getElementById("selectedinvoice").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&invoice_by='+selectedinvoice;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getinvoiceby.asp', 'POST', PostStr, showlistinvoice);
}

// The function for handling the response from the server
var showlistinvoice = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("invoicelist").innerHTML = "";
	document.getElementById("invoicelist").innerHTML = response;

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

var getlisttickettransporter = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var salesshipmentno = document.getElementById("salesshipmentno").value;
	var PostStr = 'fcba='+selectedfcba+'&salesshipmentno='+salesshipmentno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlisttickettransporterinvoice.asp', 'POST', PostStr, showlisttickettransporter);
}

// The function for handling the response from the server
var showlisttickettransporter = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("ticketlist").innerHTML = "";
	document.getElementById("ticketlist").innerHTML = response;

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
var getTransporterContractInfo = function () {
	// build up the post string when passing variables to the server side page
	var transportercontractno = document.getElementById("transportercontractno").value;
	var fcba = document.getElementById("selectedfcba").value;
	var PostStr = 'transportercontractno='+transportercontractno+'&fcba='+fcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetTransporterContractInfo.asp', 'POST', PostStr, showTransporterContractInfo);
	
}

// The function for handling the response from the server
var showTransporterContractInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // update the Div to show the result from the server
    document.getElementById("losstolerance").value = "";
    document.getElementById("losstolerance").value = s.losstolerance;
	document.getElementById("transportercontractrate").value = "";
	document.getElementById("transportercontractrate").value = s.rate;
}

var getSalesShipmentInfo = function () {
	// build up the post string when passing variables to the server side page
	var fcba = document.getElementById("selectedfcba").value;
	var wbt_chitno = document.getElementById("wbt_chitno").value;
	var salesshipmentno = document.getElementById("salesshipmentno").value;
	var PostStr = 'fcba='+fcba+'&wbt_chitno='+wbt_chitno+'&salesshipmentno='+salesshipmentno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetSalesShipmentInfo.asp', 'POST', PostStr, showSalesShipmentInfo);
	
}

// The function for handling the response from the server
var showSalesShipmentInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // update the Div to show the result from the server
	document.getElementById("acceptedqtybruto").value = "";
	document.getElementById("deliveredqty").value = "";
	document.getElementById("transportationcost").value = "";
	document.getElementById("losstolerance").value = "";
	document.getElementById("losstolerance").value = s.losstolerance;
	document.getElementById("transportationcost").value = s.transportationcost;
	document.getElementById("acceptedqtybruto").value = s.acceptquantity;
	document.getElementById("deliveredqty").value = s.quantity;
	// alert(s.quantity);
	var deliveredqty = s.quantity;
	var acceptedqtybruto = s.acceptquantity;
	var transportationcost = s.transportationcost;
	var losstolerance = s.losstolerance;
	var newacceptedqtybruto = acceptedqtybruto.replace(/,/g, '');
	var newdeliveredqty = deliveredqty.replace(/,/g, '');

	//untuk Menghitung variance qty
	var varianceqty;
	if ( newdeliveredqty == 0 ) {
		document.getElementById("varianceqty").value = newacceptedqtybruto;
	} else {
		varianceqty = newdeliveredqty - newacceptedqtybruto;
	}
		document.getElementById("varianceqty").value = varianceqty;
	
	//untuk menghitung max tolerance/variance tolerance qty
	// var losstolerance = document.getElementById("losstolerance").value;
	var maxtolerance = newdeliveredqty * losstolerance /100;
	document.getElementById("maxtolerance").value = maxtolerance;
	
	var penaltyqty;
	if ( varianceqty < maxtolerance ) {
		penaltyqty = 0;
		document.getElementById("penaltyqty").value = penaltyqty;
	} else {
		penaltyqty = varianceqty - maxtolerance;
		document.getElementById("penaltyqty").value = penaltyqty;
	}
	// alert(penaltyqty);

	var vpenaltyrate = document.getElementById("penaltyrate").value;
	var penaltyrate = vpenaltyrate.replace(",","");
	var penaltyamount = penaltyqty * penaltyrate;
	document.getElementById("penaltyamount").value = penaltyamount;
	
	//untuk menghitung loss percentage
	var losspercentage;
	if ( newdeliveredqty == 0 ) {
		losspercentage = 0;
	} else {
		losspercentage = Math.round( ( 100 - (newacceptedqtybruto /newdeliveredqty *100) )) ;
	}
	// alert(losspercentage);
	document.getElementById("losspercentage").value = losspercentage;

	var amountbeforeloss;
	amountbeforeloss = newdeliveredqty * transportationcost;
	document.getElementById("amountbeforeloss").value = amountbeforeloss;
	
	document.getElementById("amountafterloss").value = amountbeforeloss - penaltyamount;
}

</script>
<script language="JavaScript">					
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
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
 <!-------------awal!------------->
 <%


invoicecode = request("invoicecode")
creationdate = request("creationdate")
transportercontractno = request("transportercontractno")
penaltyrate = request("penaltyrate")
vat = request("vat")
description = request("description")
status = request("status")
salesshipmentno = request("salesshipmentno")
wbt_chitno = request("wbt_chitno")
acceptedqtybruto = request("acceptedqtybruto")
deliveredqty = request("deliveredqty")
varianceqty = request("varianceqty")
amountbeforeloss = request("amountbeforeloss")
losspercentage = request("losspercentage")
maxtolerance = request("maxtolerance")
penaltyqty = request("penaltyqty")
penaltyamount = request("penaltyamount")
amountafterloss = request("amountafterloss")
remarks = request("remarks")
selectedtransporter = request("selectedtransporter")
amount = request("amount")
amountvat = request("amountvat")
totalamountvat = request("totalamountvat")
lineno = request("lineno")
transportationcost =  request("transportationcost")
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
losstolerance = request("losstolerance")
no_faktur = Request("no_faktur")
fakturdate = Request("fakturdate")
pph = Request("pph")
amountpph = Request("amountpph")
amountbeforepph = Request("amountbeforepph")
amountafterpph = Request("amountafterpph")
invoice_by = Request("invoice_by")
selectedinvoice	 = Request("selectedinvoice")
customer_chitno = request("customer_chitno")

if len(trim(""&creationdate)) = 0 then creationdate = textDate(date)
if len(trim(""&fakturdate)) = 0 then fakturdate = textDate(date)

If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then	    
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, creationdate)
		If Cek_IsDailyClosing = "FALSE" then
			q = " delete from transporterinvoice " & _
				" where fcBA = '" & fcba & "'" & _
				" and invoicecode = '" & invoicecode & "'"
				' Response.Write q
				connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
		
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then
			
			q = " DELETE from transporterinvoicedtl " & _
				" WHERE invoicecode = '"& invoicecode & "'" & _
				"       AND fcba = '" & fcba & "' AND lineno = '"&lineno&"'" 
				'Response.Write q
				connStr.execute(q)

			q = " UPDATE salesshipmentdtl" & _
	            " SET   transporter_invoice = ''" & _
				" WHERE wbt_chitno = '" & wbt_chitno & "' AND fcba = '"& fcba &"' AND salesshipmentno = '"& salesshipmentno &"'" 
			response.write q
			connStr.execute(q)	
			
			cmdView	= "True"

	elseif Request("txtJenis")="New" then
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select * from transporterinvoice where invoicecode = '" & invoicecode & "' and fcba ='" & fcba & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then
			
				q = " INSERT INTO transporterinvoice ( invoicecode" & _
					"                                  , creationdate" & _
					"                                  , transportercontractno" & _
					"                                  , penaltyrate" & _
					"                                  , amount" & _
					"                                  , vat" & _
					"                                  , pph" & _
					"                                  , description" & _
					"                                  , status" & _
					"                                  , no_faktur" & _
					"                                  , fakturdate" & _
					"                                  , fcentry" & _
					"                                  , fcedit" & _
					"                                  , fcip" & _
					"                                  , fcba" & _
					"                                  , invoice_by" & _
					"                                  , lastupdate" & _
					"                                  , lasttime)" & _
					" VALUES ( '" & invoicecode & "'" & _
					"          , TO_DATE('" & creationdate & "','mm/dd/yyyy')" & _ 
					"          , '" & transportercontractno & "'" & _ 
					"          , REPLACE('" & penaltyrate & "',',','')" & _
					"          , REPLACE('" & amount & "',',','')" & _
					"          , REPLACE('" & vat & "',',','')" & _
					"          , REPLACE('" & pph & "',',','')" & _
					"          , '" & description & "'" & _
					"          , 'INVOICED'" & _
					"          , '" & no_faktur & "'" & _
					"          , TO_DATE('" & fakturdate & "', 'mm/dd/yyyy')" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"          , '" & fcba & "'" & _
					"          , '" & invoice_by & "'" & _
					"          , SYSDATE" & _
					"          , TO_CHAR(SYSDATE, 'hh24:mi') )"								
				' Response.Write q
				connStr.execute(q)
			end if
			
			if len(trim(wbt_chitno)) > 0 then
				qMaxLineNo = " SELECT COUNT(1) total " & _
							 " FROM salesshipmentdtl " & _
							 " where fcba = '" & fcba & "' AND wbt_chitno = '" & wbt_chitno & "'"
													
				Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
				rsqMaxLineNo.Open qMaxLineNo, connstr

				if not rsqMaxLineNo.eof then
					has_datawbt_chitno = cdbl(to_nol(rsqMaxLineNo("total")))
				else
					has_datawbt_chitno = 0
				end if
				rsqMaxLineNo.close : set rsqMaxLineNo = Nothing
				
				if has_datawbt_chitno > 0 then
					qMaxLineNo = " SELECT COUNT(1) total " & _
								 " FROM transporterinvoicedtl " & _
								 " where fcba = '" & fcba & "' AND wbt_chitno = '" & wbt_chitno & "' AND salesshipmentno = '"& salesshipmentno &"'"
														
					Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
					rsqMaxLineNo.Open qMaxLineNo, connstr

					if not rsqMaxLineNo.eof then
						has_invoiced = cdbl(to_nol(rsqMaxLineNo("total")))
					else
						has_invoiced = 0
					end if
					rsqMaxLineNo.close : set rsqMaxLineNo = Nothing
					
					if has_invoiced = 0 then
						q = " INSERT INTO transporterinvoicedtl ( invoicecode" & _
							"                                     , salesshipmentno" & _
							"                                     , lineno" & _
							"                                     , wbt_chitno" & _
							"                                     , acceptedqtybruto" & _
							"                                     , deliveredqty" & _
							"                                     , varianceqty" & _
							"                                     , amountbeforeloss" & _
							"                                     , losspercentage" & _
							"                                     , maxtolerance" & _
							"                                     , penaltyqty" & _
							"                                     , penaltyamount" & _
							"                                     , amountafterloss" & _
							"                                     , transportationcost" & _
							"                                     , losstolerance" & _
							"                                     , remarks" & _
							"                                     , fcentry" & _
							"                                     , fcedit" & _
							"                                     , fcip" & _
							"                                     , fcba" & _
							"                                     , lastupdate" & _
							"                                     , lasttime)" & _
							" VALUES ( '" & invoicecode & "'" & _
							"          , '" & salesshipmentno & "'" & _
							"          , (SELECT NVL(MAX(lineno),0) FROM transporterinvoicedtl WHERE invoicecode = '" & invoicecode & "' AND fcba = '" & fcba & "') +1" & _
							"          , '" & wbt_chitno & "'" & _
							"          , REPLACE('" & acceptedqtybruto & "',',','')" & _
							"          , REPLACE('" & deliveredqty & "',',','')" & _
							"          , REPLACE('" & varianceqty & "',',','')" & _
							"          , REPLACE('" & amountbeforeloss & "',',','')" & _
							"          , REPLACE('" & losspercentage & "',',','')" & _
							"          , REPLACE('" & maxtolerance & "',',','')" & _
							"          , REPLACE('" & penaltyqty & "',',','')" & _
							"          , REPLACE('" & penaltyamount & "',',','')" & _
							"          , REPLACE('" & amountafterloss & "',',','')" & _
							"          , REPLACE('" & transportationcost & "',',','')" & _
							"          , REPLACE('" & losstolerance & "',',','')" & _
							"          , '" & remarks & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & fcba & "'" & _
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE, 'hh24:mi') )"
							' Response.Write q
							connStr.execute(q)
						
						amount = 0
						strSuppyingDeclaration	=	" SELECT NVL(amountafterloss,0) amount FROM TRANSPORTERINVOICEDTL" & _
													" WHERE invoicecode = '" & invoicecode & "'" 
						'response.Write strSuppyingDeclaration
						set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
						rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
						if Not rsSupplyingDeclaration.Eof then
							while Not rsSupplyingDeclaration.Eof
								amount = amount + cdbl(to_nol(rsSupplyingDeclaration("amount")))
								rsSupplyingDeclaration.movenext
							wend
						End if
						rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing
						' response.Write amount
						amountvat = amount * cdbl(to_nol(vat))/100
						totalamountvat = amount +amountvat
						amountbeforepph = amount
						amountpph = amountbeforepph * cdbl(to_nol(pph))/100
						amountafterpph = amountbeforepph - amountpph


						q = " UPDATE transporterinvoice" & _
				            " SET   amount  = REPLACE('" & amount & "',',','')" & _
				            " , amountbeforepph  = REPLACE('" & amountbeforepph & "',',','')" & _
				            " , amountpph  = REPLACE('" & amountpph & "',',','')" & _
				            " , amountafterpph  = REPLACE('" & amountafterpph & "',',','')" & _
							" WHERE invoicecode = '" & invoicecode & "' AND fcba = '"& fcba &"'" 
						' response.write q
						connStr.execute(q)

						q = " UPDATE salesshipmentdtl" & _
				            " SET   transporter_invoice = 'INVOICED'" & _
							" WHERE wbt_chitno = '" & wbt_chitno & "' AND fcba = '"& fcba &"' AND salesshipmentno = '"& salesshipmentno &"'" 
						' response.write q
						connStr.execute(q)	
					else
						err.number = 1001
						err.description = "Chit No : "&wbt_chitno&", has invoiced! Please use other chit no"
					end if
				else
					err.number = 1001
					err.description = " Please use correct chit no"
				end if
			else
				err.number = 1001
				err.description = "Please Enter Chit No"
			end if
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		cmdView	= "True"
	End If
else
	cmdView	= request("cmdView")
End If
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

'untuk menampilkan datepicker default
If Len(Trim(transdate)) = 0 Then	transdate = textDate(date)

Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select hdr.*, bu.fccompanycode" & _
							" from transporterinvoice hdr" & _
							" left join businessunit bu on bu.fccode = hdr.fcba" & _
							" where hdr.invoicecode = '" & invoicecode & "'" 
							
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	creationdate = rsSupplyingDeclaration("creationdate")
	transportercontractno = rsSupplyingDeclaration("transportercontractno")	
	penaltyrate = rsSupplyingDeclaration("penaltyrate")
	vat = rsSupplyingDeclaration("vat")	
	description = rsSupplyingDeclaration("description")
	status = rsSupplyingDeclaration("status")
	invoice_by = rsSupplyingDeclaration("invoice_by")
	fccompanycode = rsSupplyingDeclaration("fccompanycode")
	fcba = rsSupplyingDeclaration("fcba")
	selectedfcba = rsSupplyingDeclaration("fcba")
	selectedtransporter = rsSupplyingDeclaration("transportercontractno")
	selectedinvoice = rsSupplyingDeclaration("invoice_by")
	no_faktur = rsSupplyingDeclaration("no_faktur")
	fakturdate = rsSupplyingDeclaration("fakturdate")
	amountbeforepph = rsSupplyingDeclaration("amountbeforepph")
	amountpph = rsSupplyingDeclaration("amountpph")
	amountafterpph = rsSupplyingDeclaration("amountafterpph")
	pph = rsSupplyingDeclaration("pph")
End if
'response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

amount = 0
strSuppyingDeclaration	=	" SELECT NVL(amountafterloss,0) amount FROM TRANSPORTERINVOICEDTL" & _
							" WHERE invoicecode = '" & invoicecode & "'" 
'response.Write strSuppyingDeclaration
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	while Not rsSupplyingDeclaration.Eof
		amount = amount + cdbl(to_nol(rsSupplyingDeclaration("amount")))
		rsSupplyingDeclaration.movenext
	wend
End if
rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing
' response.Write amount
amountvat = amount * cdbl(to_nol(vat))/100
totalamountvat = amount +amountvat
amountbeforepph = amount
amountpph = amountbeforepph * cdbl(to_nol(pph))/100
amountafterpph = amountbeforepph - amountpph

If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & creationdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & creationdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
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
 
	<form name="frmEntry" method="post" action="transporterinvoice.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Invoice No.</td>
			<td align="left">
			<input type="text" maxlength="50" name="invoicecode" style="background:#FFFFCC;" size="30" value="<%=invoicecode%>" autocomplete="Off">
			<font color="#FF0000">*</font>
			</td>
			<td align="left">Penalty Rate</td>
			<td align="left">
				<input type="text" size="30" id="penaltyrate" name="penaltyrate" value="<%=penaltyrate%>" onkeyup="hanyaAngka(this);" >
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>

			<td align="left" width="20%">Creation Date 
			<td align="left">
				<input class="datepicker" type="text" size="15" name="creationdate" value="<%=creationdate%>">
				<font color="#FF0000">*</font>
			</td>
			<td align="left">PPn</td>
			<td align="left">
				<input type="text" size="5" maxlength="3" id="vat" name="vat" value="<%=vat%>" >
				% <font color="#FF0000"> *</font>
				<input type="text" size="15" style="background:#EEE;" name="amountvat" value="<%=amountvat%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
		</tr>
		
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba(); javascript:getlistinvoice();" class="chosen-select" style="width: 50%;">
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

			<td align="left" width="20%">PPh 23</td>
			<td align="left">
				<input type="text" size="5" maxlength="3" name="pph" value="<%=pph%>"> %
				<input type="text" size="15" style="background:#EEE;" name="amountpph" value="<%=amountpph%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
			
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<div id="fcbalist"></div>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			<td align="left" width="20%">Amount</td>
			<td align="left">
				<input type="text" size="15" style="background:#EEE;" name="amount" value="<%=amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
		</tr>
		
		<tr>
			<td align="left">Transporter Contract No</td>
			<td align="left">
				<div id= "transportercontractlist"></div>
				<font color="#FF0000">*</font>
			</td>
				<input type="Hidden" id="selectedtransporter" name="selectedtransporter" value="<%=selectedtransporter%>">
			<td align="left" >Amount + VAT</td>
			<td align="left">
				<input type="text" size="15" style="background:#EEE;" name="totalamountvat" value="<%=totalamountvat%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By
				<td align="left">
					<div id="invoicelist"></div>
				</td>
					<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="<%=selectedinvoice%>">
			</td>
			<td align="left" >Amount Before PPh 23</td>
			<td align="left">
				<input type="text" size="15" style="background:#EEE;" name="amountbeforepph" value="<%=amountbeforepph%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">No Faktur</td>
			<td align="left">
				<input type="text" maxlength="100" name="no_faktur" size="30" value="<%=no_faktur%>" autocomplete="off" >
			</td>
			<td align="left" >Amount After PPh</td>
			<td align="left">
				<input type="text" size="15" style="background:#EEE;" name="amountafterpph" value="<%=amountafterpph%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
		</tr>
		<tr>
			<td align="left"> Faktur Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="fakturdate" size="10" value="<%=fakturdate%>" autocomplete="Off">
			</td>
			<td align="left" width="20%">Status</td>
			<td align="left">
				<input type="text" size="30" name="status" value="<%=status%>" style="background:#EEE;" readonly>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Description</td>
			<td align="left">
				<textarea name="description" id="description" cols="30"><%=description%></textarea>
			</td>		
			
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				<a href="print_transporter_invoice_new.html?fcba=<%=fcba%>&invoicecode=<%=invoicecode%>" target="_blank">
					<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
				</a>&nbsp;&nbsp; 
				
					<a href="transporterinvoice.asp?invoicecode=<%=invoicecode%>&fcba=<%=fcba%>&txtJenis=HeaderDelete">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0" onClick="return confirm('Are you sure to delete');">
					</a>
				
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
<%
if cmdView = "True" and len(trim(invoicecode)) > 0 and len(trim(transportercontractno)) > 0 and len(trim(penaltyrate)) > 0 and len(trim(vat)) > 0 Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'query Untuk menampilkan isi di Field
	q	= " SELECT *" & _
		  " FROM transporterinvoicedtl " & _
		  " where invoicecode = '" & invoicecode & "'"

	' response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.Open q, connstr

%>
	<form action="transporterinvoice.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td align="left" width="20%">Shipment No</td>
			<td align="left">
				<select id="salesshipmentno" name="salesshipmentno" onchange="javascript:getlisttickettransporter();" class="chosen-select" style="width: 50%;">
						<option value=""></option>
						<%
						qlov =  " SELECT DISTINCT ssh.salesshipmentno" & _
								"  FROM salesshipmenthdr ssh" & _
								"       INNER JOIN salesshipmentdtl ssd" & _
								"          ON (    ssh.salesshipmentno = ssd.salesshipmentno" & _
								"              AND ssh.fcba = ssd.fcba)" & _
								"       LEFT JOIN deliveryorder do" & _
								"          ON (do.deliveryorderno = ssh.deliveryorderno AND do.fcba = ssh.fcba)" & _
								" WHERE     ssh.fcba = '"& fcba &"'" & _
								"       AND do.transportercontractno = '"& transportercontractno &"'" & _
								"       AND transporter_invoice IS NULL"
								  
						' response.write qlov
						Set rslov = Server.CreateObject("ADODB.Recordset")
						rslov.Open qlov, connstr 
						if Not rslov.eof then
							While Not rslov.EOF
						%>
						<option value="<%=rslov("salesshipmentno")%>" <%if salesshipmentno = rslov("salesshipmentno") then%>selected<%end if%>><%=rslov("salesshipmentno")%></option>
						<%
								rslov.MoveNext
							Wend
						end if
						%>
					</select>
			</td>
			<td align="left" width="20%">Loss Percentage</td>
			<td align="left">
			<input type="text" id="losspercentage" name="losspercentage" style="background:#eee;" size="30" value="<%=losspercentage%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly >
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Ticket No.</td>
			<td align="left">
				<div id="ticketlist"></div>
				<font color="#FF0000">*</font>
			</td>
			<td align="left" width="20%">Max Tolerance</td> 
			<td align="left">
			<input type="text" maxlength="15" id="maxtolerance" name="maxtolerance" style="background:#eee;" size="30" value="<%=maxtolerance%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Accepted Qty (Netto)</td>
			<td align="left">
			<input type="text" id="acceptedqtybruto" name="acceptedqtybruto" style="background:#eee;" size="30" value="<%=acceptedqtybruto%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly>
			</td>
			
			<td align="left" width="20%">Penalty Qty</td>
			<td align="left">
			<input type="text" id="penaltyqty" name="penaltyqty" style="background:#eee;" size="30" value="<%=penaltyqty%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly >
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Delivered Qty</td>
			<td align="left">
			<input type="text" id="deliveredqty" name="deliveredqty" style="background:#eee;" size="30" value="<%=deliveredqty%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits':2, 'digitsOptional': false, 'placeholder': '0'" readonly >
			</td>
			
			<td align="left" width="20%">Penalty Amount</td>
			<td align="left">
			<input type="text" id="penaltyamount" name="penaltyamount" style="background:#eee;" size="30" value="<%=penaltyamount%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly >
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Variance Qty</td>
			<td align="left">
			<input type="text" id="varianceqty" name="varianceqty" style="background:#eee;" size="30" value="<%=varianceqty%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly >
			</td>
			<td align="left" width="20%">Amount Before % Loss</td>
			<td align="left">
			<input type="text" id="amountbeforeloss" name="amountbeforeloss" style="background:#eee;" size="30" value="<%=amountbeforeloss%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly >
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Transporter Rate</td>
			<td align="left">
			<input type="text" id="transportationcost" name="transportationcost" style="background:#eee;" size="30" value="<%=transportationcost%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly >
			</td>
			<td align="left" width="20%">Amount After % Loss</td>
			<td align="left">
			<input type="text" id="amountafterloss" name="amountafterloss" style="background:#eee;" size="30" value="<%=amountafterloss%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly >
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Loss Tolerance</td>
			<td align="left">
				<input type="text" maxlength="15" id="losstolerance" name="losstolerance" style="background:#eee; text-align: right" size="30" value="<%=losstolerance%>" autocomplete="Off" readonly>
			</td>
			<td align="left" width="20%">Remarks</td>
			<td align="left">
				<textarea name="remarks" id="textarea" cols="30"><%=remarks%></textarea>
			</td>
		</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="invoicecode" value="<%=invoicecode%>">
					<input type="hidden" name="creationdate" value="<%=creationdate%>">
					<input type="hidden" name="transportercontractno" value="<%=transportercontractno%>">
					<input type="hidden" name="penaltyrate" value="<%=penaltyrate%>">
					<input type="hidden" name="vat" value="<%=vat%>">
					<input type="hidden" name="description" value="<%=description%>">
					<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					<input type="hidden" name="amount" value="<%=amount%>">
					<input type="hidden" name="pph" value="<%=pph%>">
					<input type="hidden" name="amountpph" value="<%=amountpph%>">
					<input type="hidden" name="amountbeforepph" value="<%=amountbeforepph%>">
					<input type="hidden" name="amountafterpph" value="<%=amountafterpph%>">
					<input type="hidden" name="invoice_by" value="<%=invoice_by%>">
					<input type="hidden" name="selectedinvoice" value="<%=selectedinvoice%>">
					<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
					<input type="hidden" name="selectedtransporter" value="<%=selectedtransporter%>">
					<input type="hidden" name="no_faktur" value="<%=no_faktur%>">
					<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
		
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					
						<%IF Request("txtJenis")="Edit" then%>
							<input type="hidden" name="txtJenis" value="Edit">
							<input type="hidden" name="chitno" value="<%=chitno%>">
							<input Type="Submit"  name="Add" value="Update">
						<%Else%>
							<input type="hidden" name="txtJenis" value="New">
							<input Type="Submit"  name="Add" value="Add">
						<%End if%>
						<input Type="Reset"  name="Reset" value="Reset">
						<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
					
				</td>
			</tr>
		</table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">No</td>
		<td>&nbsp;</td>
		<td align="left">Invoice No</td>
		<td align="left">Chit No</td>
		<td align="left">Accepted Qty Bruto</td>
		<td align="left">Delivered Qty</td>
		<td align="left">Variance Qty</td>
		<td align="left">Max Tolerance</td>
		<td align="left">Loss Percentage</td>
		<td align="left">Penalty Qty</td>
		<td align="left">Penalty Amount</td>
		<td align="left">Amount Before Loss</td>
		<td align="left">Amount After Loss</td>
		<td align="left">Remarks</td>
	</tr>
   
   <%
	if Not rsVH.eof then
		nCountofField = rsVH.fields.count - 1
		no = 1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=NO%></td>
			<td valign="middle">
				<a href="transporterinvoice.asp?invoicecode=<%=rsVH("invoicecode")%>&lineno=<%=rsVH("lineno")%>&fcba=<%=rsVH("fcba")%>&wbt_chitno=<%=rsVH("wbt_chitno")%>&salesshipmentno=<%=rsVH("salesshipmentno")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
			    <td align="left" title="<%=rsVH("invoicecode").Name%>" valign="top"><%=rsVH("invoicecode")%></td>
			    <td align="left" title="<%=rsVH("wbt_chitno").Name%>" valign="top"><%=rsVH("wbt_chitno")%></td>
			    <td align="right" title="<%=rsVH("acceptedqtybruto").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("acceptedqtybruto"))),0)%></td>
			    <td align="right" title="<%=rsVH("deliveredqty").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("deliveredqty"))),0)%></td>
			    <td align="right" title="<%=rsVH("varianceqty").Name%>" valign="top"><%=rsVH("varianceqty")%></td>
			    <td align="right" title="<%=rsVH("maxtolerance").Name%>" valign="top"><%=rsVH("maxtolerance")%></td>
			    <td align="right" title="<%=rsVH("losspercentage").Name%>" valign="top"><%=rsVH("losspercentage")%></td>
			    <td align="right" title="<%=rsVH("penaltyqty").Name%>" valign="top"><%=rsVH("penaltyqty")%></td>
			    <td align="right" title="<%=rsVH("penaltyamount").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("penaltyamount"))),0)%></td>
			    <td align="right" title="<%=rsVH("amountbeforeloss").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("amountbeforeloss"))),0)%></td>
			    <td align="right" title="<%=rsVH("amountafterloss").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("amountafterloss"))),0)%></td>
			    <td align="left" title="<%=rsVH("remarks").Name%>" valign="top"><%=rsVH("remarks")%></td>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>
 
  <!-------------akhir!------------->
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

function getselectedinvoice() {
	var sfcba = document.getElementById("selectedinvoice");
	sfcba.value = document.getElementById("invoice_by").value;
}

function getselectedtransportercontract() {
	var sselectedtransporter = document.getElementById("selectedtransporter");
	sselectedtransporter.value = document.getElementById("transportercontractno").value;
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
	getlisttransportercontract();
	getlisttickettransporter();	
	getSalesShipmentInfo();
	getTransporterContractInfo();
	getselectedfcba();
	getselectedtransportercontract();
});
</script>
</body>
</html>
