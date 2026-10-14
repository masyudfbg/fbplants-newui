<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Monthly Workplan Harv</title>

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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
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
<script>
	function hitungrumus() {
		var skg_budget = $("#kg_budget").val();
		var sbjr_budget = $("#harvestingbased_abw_budget").val(); 
		var snormamandays = $("#normamandays").val(); 
		var sratemandays = $("#ratemandays").val(); 
		
		//Untuk menghitung Total Janjang
		totaljjg = (skg_budget / sbjr_budget).toFixed(2); 
		if (!isNaN(totaljjg)) {
    	 	$("#bunch").val(totaljjg);
    	}
    	else{
    		$("#bunch").val();
    	}

		//Untuk menghitung Total Mandays
		totalmandays = skg_budget / 1000 / snormamandays; 
		if (!isNaN(totalmandays)) {
    	 	$("#mandays").val(totalmandays);
    	}
    	else{
    		$("#mandays").val();
    	}
		
		//Untuk menghitung Total RP/HK
		totalrphk = (totalmandays/sratemandays).toFixed(2); 
		if (!isNaN(totalrphk)) {
    	 	$("#amountpermandays").val(totalrphk);
    	}
    	else{
    		$("#amountpermandays").val();
    	}
		
		//Untuk menghitung Total Output/JJG
		totaloutputjjg = (totaljjg * 70/100).toFixed(2); 
		if (!isNaN(totaloutputjjg)) {
    	 	$("#outputperbunch").val(totaloutputjjg);
    	}
    	else{
    		$("#outputperbunch").val();
    	}
		
		//Untuk menghitung Total HK Basis
		totalhkbasis = (parseInt(totaljjg) - parseInt(totaloutputjjg)).toFixed(2); 
		if (!isNaN(totalhkbasis)) {
    	 	$("#basismandays").val(totalhkbasis);
    	}
    	else{
    		$("#basismandays").val();
    	}
		
		//Untuk menghitung Total Lebih Basis JJG
		totallebihbasis = ((totalmandays * 7/100) * 35).toFixed(2); 
		if (!isNaN(totallebihbasis)) {
    	 	$("#extra_basis_bunch").val(totallebihbasis);
    	}
    	else{
    		$("#extra_basis_bunch").val();
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

var getNoRKB = function () {
	// build up the post string when passing variables to the server side page

	var fcba = document.getElementById("fcba").value;
	var txtYear = document.getElementById("txtYear").value;
	var txtMonth = document.getElementById("txtMonth").value;
	var activitycode = document.getElementById("activitycode").value;	
	var selectedsection = document.getElementById("selectedsection").value;	
	var selectedfieldcode = document.getElementById("selectedfieldcode").value;	
	var PostStr = 'fcba='+fcba +'&txtYear='+ txtYear +'&txtMonth='+ txtMonth + '&activitycode='+ activitycode + '&unitkerja='+ selectedsection + '&fieldcode='+ selectedfieldcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetNoRKBPanen.asp', 'POST', PostStr, ShowNoRKBResponse);
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

var getFieldInfo = function () {

	var fcba = document.getElementById("fcba").value;
	var txtYear = document.getElementById("txtYear").value;
	var txtMonth = document.getElementById("txtMonth").value;
	var fieldcode = document.getElementById("fieldcode").value;
	var PostStr = 'fcba='+fcba+'&fieldcode='+fieldcode+'&txtMonth='+txtMonth+'&txtYear='+txtYear;
	// alert(PostStr);
	doAJAXCall('ajax/GetFieldRKB.asp', 'POST', PostStr, ShowFieldInfoResponse);
	
}

var ShowFieldInfoResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.lineno);
    
    // update the Div to show the result from the server
	document.getElementById("hectarageplanted").value = s.hectarageplanted;
	document.getElementById("plantingpoint").value = s.plantingpoint;
	document.getElementById("plantingyear").value = s.plantingyear;
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

// launched from button click 
var getListJob = function () {
	// build up the post string when passing variables to the server side page
	
	var fcba = document.getElementById("fcba").value;
	var selectedactivity = document.getElementById("selectedactivity").value;
	var PostStr = 'fcba='+fcba+'&activitycode='+selectedactivity;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistjobrkbharv.asp', 'POST', PostStr, showListJob);
}

// The function for handling the response from the server
var showListJob = function (oXML) { 
    // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("jobcodelist").innerHTML = "";
	document.getElementById("jobcodelist").innerHTML = response;

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
var getListsection = function () {
	// build up the post string when passing variables to the server side page
	
	var fcba = document.getElementById("fcba").value;
	var selectedsection = document.getElementById("selectedsection").value;
	var PostStr = 'fcba='+fcba+'&unitkerja='+selectedsection;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getListSectionRKBHarv.asp', 'POST', PostStr, showListsection);
}

// The function for handling the response from the server
var showListsection = function (oXML) { 
    // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("sectionlist").innerHTML = "";
	document.getElementById("sectionlist").innerHTML = response;

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
var getListfieldcode = function () {
	// build up the post string when passing variables to the server side page
	
	var fcba = document.getElementById("fcba").value;
	var selectedsection = document.getElementById("selectedsection").value;
	var selectedfieldcode = document.getElementById("selectedfieldcode").value;
	var PostStr = 'fcba='+fcba+'&unitkerja='+selectedsection+'&fieldcode='+selectedfieldcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfieldrkb.asp', 'POST', PostStr, showLListfieldcode);
}

// The function for handling the response from the server
var showLListfieldcode = function (oXML) { 
    // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("fieldcodelist").innerHTML = "";
	document.getElementById("fieldcodelist").innerHTML = response;

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
};

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


' Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, LineNoHist, ITEMCODE
' DIM v_NURSERYCODE,TRANSDATE
' DIM v_QUANTITY
' Dim cmdView
' Dim qSupplyingDeclaration, rsNurseryInfo, tempStoreCode, tempStockCode, doneError
' Dim qLineMax, rsqLineMax, tempMaxLine
' Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
' Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
' Dim qLessZero, rsLessZero, tempislesszero
' Dim done, qLineUpper, rsqLineUpper, totalUpper, qQuantity, rsqQuantity, totalQuantity, QuantityOriginal
' Dim qDeleteHist
' Dim rsReceiveDetail, originalQuantity, v_CrossCekDetail
Dim qSupplyingDeclaration, rsSupplyingDeclarationInfo, tempQuantity, tempDecType
Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, v_FCCODE, v_FCNAME
Dim qSupplyingInfo, rsSupplyingInfo, qDeclarationType, rsDeclarationType
Dim salescontractno, parcode, totalstand, customercode, deliveryaddress, deliverytype, downpayment
Dim FIELDCODE, productcode,  unitofmeasure
Dim QUANTITY, curr, salesprices, amount, vat, totalamountvat, amountvat
Dim REMARKS, paymentcondition, deliverycondition
Dim creationdate, startdate, enddate, deliverydate
Dim cmdView
Dim txtDocumentNo ,txtlineno
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
Dim v_txtMonth,v_txtYear,selectedmonth,selectedyear

v_txtYear		= Request("txtYear")
v_txtMonth		= Request("txtMonth")
creationdate	= Request("creationdate")

If Len(Trim(v_txtMonth)) <> 0 Then selectedmonth = v_txtMonth 
If Len(Trim(v_txtYear)) <> 0 Then selectedyear = v_txtYear 

If Len(Trim(creationdate)) = 0 Then	creationdate = textDate(date)
If Len(Trim(v_txtYear)) = 0 Then v_txtYear = Year(textDate(date))
If Len(Trim(v_txtMonth)) = 0 Then v_txtMonth = Month(textDate(date))

activitycode = request("activitycode")
transno = request("transno")
ratemandays = request("ratemandays")
fcba = request("fcba")

unitkerja = request("unitkerja")
fieldcode = request("fieldcode")
hectarageplanted = request("hectarageplanted")
plantingpoint = request("plantingpoint")
plantingyear = request("plantingyear")
harvestingbased_abw_budget = request("harvestingbased_abw_budget")
rotasi = request("rotasi")
kg_budget = request("kg_budget")
bunch = request("bunch")
normamandays = request("normamandays")
mandays = request("mandays")
amountpermandays = request("amountpermandays")
outputperbunch = request("outputperbunch")
premiluasan = request("premiluasan")
basismandays = request("basismandays")
premi_basis = request("premi_basis")
extra_basis_bunch = request("extra_basis_bunch")
premi_extra_basis = request("premi_extra_basis")
total_premi = request("total_premi")
total_amount_panen = request("total_amount_panen")
total_amountperkg = request("total_amountperkg")
lineno = request("lineno")
selectedactivity = request("selectedactivity")
percent_sb = request("percent_sb")
percent_lb = request("percent_lb")
selectedsection = request("selectedsection")
selectedfieldcode = request("selectedfieldcode")

Function CheckFieldCode(initialFieldCode, initialFcba, initialDate)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from dual " & _
					" where field_api.GetValid_Activation('" & initialFieldCode & "', '" & initialFcba & "', to_date('"& initialDate &"', 'mm/dd/yyyy')) = 'Y'"
	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckFieldCode = initialAvailable
	'Select count(1) from dual where field_api.GetValid_Activation('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')) = 'Y'
	'select to_char (field_api.GetValid_Plantingdate('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')),'yyyy') from dual
End Function

if Request("txtJenis") = "" then
	 '---- Control Job
	qHeader = "select a.fcba,period_year,period_month,activitycode,rate,b.fieldcode,nvl(b.mandays,0) mandays,a.rkbno" & _ 
	          " ,unitkerja" & _
			  " from monthlyworkplan_panen a" & _
			  " inner join monthlyworkplan_panendet b on a.rkbno = b.rkbno and a.fcba = b.fcba " & _
			  " where a.rkbno ='" & transno & "'"
	'response.write qHeader
	Set rsHeader = Server.CreateObject("ADODB.Recordset")
	rsHeader.Open qHeader, connstr
		if not rsHeader.eof then
			fcba = rsHeader("fcba")
			v_txtYear = rsHeader("period_year")
			v_txtMonth = rsHeader("period_month")
			selectedactivity = rsHeader("activitycode")
			rate = rsHeader("rate")
			selectedafdeling = rsHeader("unitkerja")
			selectedfield = rsHeader("fieldcode")
        	mandays = rsHeader("mandays")
			transno = rsHeader("rkbno")
			selectedmonth = rsHeader("period_month")
			selectedyear  = rsHeader("period_year")
			
		end if
	'---- End Control Job
	
   rsHeader.Close
   set rsHeader = Nothing
	
	'if Cdbl(volume) > Cdbl(haactual) then
	'	Err.Description = "ha plan (" & volume & ") cannot be greater than ha actual (" & haactual & ") !"
	'	response.Write  "<div class = ""error"">" & err.description & "</div>"
    'end if			

 end if


If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
 	v_txtMonth = request("selectedmonth")
	v_txtYear = request("selectedyear")
	
	if Request("txtJenis")= "HeaderDelete" then
		connstr.BeginTrans
		
		strFieldDetail = " SELECT COUNT(1) total FROM salesagreementdet " & _
                         " WHERE salescontractno = '" & salescontractno & "'" & _
                         "       AND fcBA = '" & fcba & "'"								
		'Response.Write strFieldDetail
		Set rsQty = Server.CreateObject("ADODB.Recordset")
		rsQty.Open strFieldDetail, connstr							
		if not rsQty.Eof then
			has_detail = cdbl (to_nol(rsQty("total")))
		End If
		
		if has_detail = 0 then
			q = " delete from salesagreementhdr " & _
				" where fcBA = '" & fcba & "'" & _
				" and salescontractno = '" & salescontractno & "'"
				'Response.Write q
			connStr.execute(q)
		else
			err.number = 1001
			err.description = " Sales Contract No """ & salescontractno & """ has detail, please delete all detail first!"
		End If
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			' strFieldDetail =	"select quantity" & _
			' 					"		from salesagreementdet " & _
			' 					" 		where lineno = '" & txtlineno & "'" & _
			' 					"       and salescontractno = '" & salescontractno & "'" & _
			' 					"       and fcBA = '" & fcba & "'"								
			' 							' Response.Write strFieldDetail
			' 							Set rsQty = Server.CreateObject("ADODB.Recordset")
			' 							rsQty.Open strFieldDetail, connstr							
			' if not rsQty.Eof then
			' do while not rsQty.eof
			' 		editquantity = cdbl (to_nol(rsQty("quantity")))
			' 		rsQty.MoveNext
			' 	loop
			' End If
			
			q = " delete from monthlyworkplan_panendet " & _
				" 		where rkbno= '" & transno & "'" & _
				"       and lineno = '" & lineno & "'" & _
				"       and fcBA = '" & fcba & "'"
				' Response.Write q
				connStr.execute(q)
			
			' v_totalamount = 0
			' qLineMax = " select nvl(SUM(totalamountvat),0) totalamountvat" & _
			' 		   " from salesagreementdet" & _
			' 		   " where salescontractno = '" & salescontractno & "'" & _
			' 		   " and fcba = '" & fcba & "'" 			

			' set rsqLineMax = Server.CreateObject("ADODB.Recordset")
			' rsqLineMax.Open qLineMax, connstr
			' if not rsqLineMax.eof then
			' 	totalamountvat = cdbl(rsqLineMax("totalamountvat"))
			' end if 

			' qLineMax = " select nvl(SUM(dp_aftervat),0) dp_aftervat" & _
			' 		   " from salesagreementhdr" & _
			' 		   " where salescontractno = '" & salescontractno & "'" & _
			' 		   " and fcba = '" & fcba & "'" 			

			' set rsqLineMax = Server.CreateObject("ADODB.Recordset")
			' rsqLineMax.Open qLineMax, connstr
			' if not rsqLineMax.eof then
			' 	dp_aftervat = cdbl(rsqLineMax("dp_aftervat"))
			' end if 

			' v_totalamount = v_totalamount + totalamountvat
			' if cdbl(to_nol(v_totalamount))  = 0 then v_percentage_dp = 0 else v_percentage_dp = (dp_aftervat / v_totalamount) * 100

			' q =	" Update salesagreementhdr" & _
			' 	" set totalamount = replace('" & v_totalamount & "',',','')" & _
			' 	"     , percentage_dp = replace('" & v_percentage_dp & "',',','.')" & _
			' 	" 		where salescontractno = '" & salescontractno & "'" & _
			' 	"       and fcBA = '" & fcba & "'"
			' ' response.write q
			' connStr.execute(q)
			
			cmdView	= "True"

	elseif Request("txtJenis")="New" then
			Dim rsSupplyingDeclarationHdr
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select rkbno from MONTHLYWORKPLAN_PANEN where rkbno = '" & transno & "' and fcba ='" & fcba & "' ", connstr
			
			
			If rsSupplyingDeclarationHdr.EOF Then
				q = " INSERT into MONTHLYWORKPLAN_PANEN" & _
					" (rkbno, period_year, period_month, activitycode, rate, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" '" & transno & "'" & _ 
					" , '" & v_txtYear & "'" & _
					" , '" & v_txtMonth & "'" & _
					" , '" & activitycode & "'" & _ 
					" , replace('" & ratemandays & "',',','')" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & fcba & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"								

				connStr.execute(q)
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing
					
			' If Len(Trim(unitkerja)) = 0 Then
			' 	Err.Number = 1001
			' 	Err.Description = "Unit Kerja Harus dipilih, tidak boleh kosong"
			' Else
			' 	If Len(Trim(fieldcode)) = 0 Then
			' 		Err.Number = 1001
			' 		Err.Description = "Fieldcode Harus dipilih, tidak boleh kosong"
			' 	Else 
					qLineMax = " select nvl(max(lineno),0) tempMaxLine" & _
							   " from MONTHLYWORKPLAN_PANENDET" & _
							   " where rkbno = '" & transno & "'" & _
							   " and fcba = '" & fcba & "'" 			

					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
					else
						tempMaxLine = 0
					end if 

					tempMaxLine = tempMaxLine + 1
					txtLineNoHist = tempMaxLine
					
					q = " insert into MONTHLYWORKPLAN_PANENDET" & _
						" (rkbno, unitkerja, fieldcode, plantingyear, hectarageplanted, plantingpoint, " & _ 
						" harvestingbased_abw_budget, rotation, kg_budget, bunch, normamandays, " & _
						" mandays, amountpermandays, outputperbunch, premi_luasan, basismandays, premi_basis, " & _ 
						" extra_basis_bunch, premi_extra_basis, total_premi, total_amount_panen," & _ 
						" total_amountperkg, lineno, percent_lb, percent_sb, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
						" '" & transno & "'" & _
						" , '" & unitkerja & "'" & _
						" , '" & fieldcode & "'" & _
						" , '" & plantingyear & "'" & _
						" , '" & hectarageplanted & "'" & _
						" , '" & plantingpoint & "'" & _ 
						" , replace('" & harvestingbased_abw_budget & "',',','')" & _
						" , replace('" & rotasi & "',',','')" & _
						" , replace('" & kg_budget & "',',','')" & _
						" , replace('" & bunch & "',',','')" & _
						" , replace('" & normamandays & "',',','')" & _
						" , replace('" & mandays & "',',','')" & _
						" , replace('" & amountpermandays & "',',','')" & _
						" , replace('" & outputperbunch & "',',','')" & _
						" , replace('" & premiluasan & "',',','')" & _
						" , replace('" & basismandays & "',',','')" & _
						" , replace('" & premi_basis & "',',','')" & _
						" , replace('" & extra_basis_bunch & "',',','')" & _
						" , replace('" & premi_extra_basis & "',',','')" & _
						" , replace('" & total_premi & "',',','')" & _
						" , replace('" & total_amount_panen & "',',','')" & _
						" , replace('" & total_amountperkg & "',',','')" & _
						" , '" & txtLineNoHist &"'" & _
						" , replace('" & percent_lb & "',',','')" & _
						" , replace('" & percent_sb & "',',','')" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						" , '" & fcba & "'" & _
						" , sysdate" & _
						" , to_char(sysdate, 'hh24:mi')" & _
						" )"
						' Response.Write q
						connStr.execute(q)
			' 	End If
			' End If
	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q =	" Update monthlyworkplan_panendet" & _
			" set unitkerja = '" & unitkerja & "'" & _
			"     , fieldcode = '" & fieldcode & "'" & _
			"     , harvestingbased_abw_budget= '" & harvestingbased_abw_budget & "'" & _
			"     , kg_budget = '" & kg_budget & "'" & _
			"     , percent_lb = replace('" & percent_lb & "',',','')" & _
			"     , percent_sb = replace('" & percent_sb & "',',','')" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , lastupdate = sysdate " & _
			"     , lasttime = to_char(sysdate,'hh24:mi') " & _
			" 		where rkbno= '" & transno & "'" & _
			"       and lineno = '" & lineno & "'" & _
			"       and fcBA = '" & fcba & "'"
		' response.write q
		connStr.execute(q) 		
		End If
	End If
' End If
cmdView	= "True"	
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

'untuk menampilkan datepicker default
If Len(Trim(deliverydate)) = 0 Then	deliverydate = textDate(date)
If Len(Trim(startdate)) = 0 Then startdate = textDate(date)
If Len(Trim(enddate)) = 0 Then	enddate = textDate(date)



Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" SELECT mwp.rkbno," & _
							"        mwp.period_year," & _
							"        mwp.period_month," & _
							"        mwp.activitycode," & _
							"        mwp.rate," & _
							"        mwd.unitkerja," & _
							"        mwd.fieldcode," & _
							"        mwd.plantingyear," & _
							"        mwd.hectarageplanted," & _
							"        mwd.plantingpoint," & _
							"        mwd.harvestingbased_abw_budget," & _
							"        mwd.rotation," & _
							"        mwd.kg_budget," & _
							"        mwd.bunch," & _
							"        mwd.normamandays," & _
							"        mwd.mandays," & _
							"        mwd.amountpermandays," & _
							"        mwd.outputperbunch," & _
							"        mwd.premi_luasan," & _
							"        mwd.basismandays," & _
							"        mwd.premi_basis," & _
							"        mwd.extra_basis_bunch," & _
							"        mwd.premi_extra_basis," & _
							"        mwd.total_premi," & _
							"        mwd.total_amount_panen," & _
							"        mwd.total_amountperkg," & _
							"        nvl(mwd.percent_lb,0) percent_lb," & _
							"        nvl(mwd.percent_sb,0) percent_sb," & _
							"        mwd.lineno" & _
							"   FROM MONTHLYWORKPLAN_PANENDET mwd" & _
							"        INNER JOIN MONTHLYWORKPLAN_PANEN mwp" & _
							"           ON (mwd.fcba = mwp.fcba AND mwd.rkbno = mwp.rkbno)" & _
							"  WHERE mwd.fcba = '"& fcba &"' AND mwd.rkbno = '"& transno &"'" 
						
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	rkbno = rsSupplyingDeclaration("rkbno")
	activitycode = rsSupplyingDeclaration("activitycode")
	selectedactivity = rsSupplyingDeclaration("activitycode")
	period_month = rsSupplyingDeclaration("period_month")
	period_year = rsSupplyingDeclaration("period_year")
	percent_lb = rsSupplyingDeclaration("percent_lb")
	percent_sb = rsSupplyingDeclaration("percent_sb")
	ratemandays = rsSupplyingDeclaration("rate")
	unitkerja = rsSupplyingDeclaration("unitkerja")
	selectedsection = rsSupplyingDeclaration("unitkerja")
	fieldcode = rsSupplyingDeclaration("fieldcode")
	selectedfieldcode = rsSupplyingDeclaration("fieldcode")
	plantingpoint = rsSupplyingDeclaration("plantingpoint")
	hectarageplanted = rsSupplyingDeclaration("hectarageplanted")
	plantingyear = rsSupplyingDeclaration("plantingyear")
	harvestingbased_abw_budget = rsSupplyingDeclaration("harvestingbased_abw_budget")
	rotasi = rsSupplyingDeclaration("rotation")
	kg_budget = rsSupplyingDeclaration("kg_budget")	
	bunch = rsSupplyingDeclaration("bunch")
	normamandays = rsSupplyingDeclaration("normamandays")
	mandays = rsSupplyingDeclaration("mandays")
	amountpermandays = rsSupplyingDeclaration("amountpermandays")
	outputperbunch = rsSupplyingDeclaration("outputperbunch")
	premi_luasan = rsSupplyingDeclaration("premi_luasan")
	basismandays = rsSupplyingDeclaration("basismandays")
	premi_basis = rsSupplyingDeclaration("premi_basis")
	extra_basis_bunch = rsSupplyingDeclaration("extra_basis_bunch")
	premi_extra_basis = rsSupplyingDeclaration("premi_extra_basis")
	total_premi = rsSupplyingDeclaration("total_premi")
	total_amount_panen = rsSupplyingDeclaration("total_amount_panen")
	total_amountperkg = rsSupplyingDeclaration("total_amountperkg")
End if
' response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

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
 
	<form name="frmEntry" method="post" action="MonthlyWorkPlanHarv.asp">
		<table class="data" width="100%">
		
		<tr>
			<td align="left" width="20%">Creation Date</td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="creationdate" value="<%=creationdate%>">
			</td>
			<td align="left" width="20%">Unit Kerja </td>
			<td align="left">
				<div id="sectionlist"></div>
				<font color="#FF0000">*</font>
			</td> 
			<input type="hidden" id ="selectedsection" name="selectedsection" value="<%=selectedsection%>">
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<select id="fcba" name="fcba" class="chosen-select" onchange=" javascript:getListsection(); javascript:getListJob(); javascript:getListfieldcode(); javascript:getNoRKB(); javascript:getHKRate(); javascript:getNormaHK();">
					<option value=""></option>
					<%
					qlov = 	" SELECT DISTINCT bu.fccode, bu.fcname FROM user_fcba_tab uft INNER JOIN businessunit bu ON bu.fccode = uft.fcba" & _
							" WHERE fcusercode = '" & Session("CodeName") & "' AND isactive = 'TRUE' ORDER BY bu.fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if fcba = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%> : <%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<td align="left" width="20%">Field </td>
			<td align="left">
				<div id="fieldcodelist"></div>
				<font color="#FF0000">*</font>
			</td>
			<input type="hidden" id ="selectedfieldcode" name="selectedfieldcode" value="<%=selectedfieldcode%>"> 
		</tr>
		<tr>
			<td align="left">Year</td> 
			<%  
				Yr = cint(v_txtYear)
				MinYr = 2015
				MaxYr = Yr  + 2
			%>
			<td align="left">
				<select id ="txtYear" name="txtYear" class="chosen-select" onchange="javascript:getNoRKB(); javascript:getFieldInfo();">
					<option value=""></option>
			<%
				For i = MinYr to MaxYr
			%>
				<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
			<%
				Next
			%>
				</select>
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
			<td align="left">Month</td> 
			<td align="left">
				<select id = "txtMonth" name="txtMonth" class="chosen-select" onchange="javascript:getNoRKB(); javascript:getFieldInfo();">
					<option value=""></option>
			<% 
				For i = 1 to 12
					
					if i <= 9 then 
						iBulan = 0&i
					else
						iBulan = i
					end if
					i	= iBulan
			%>
				<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Monthname(i)%></option>
			<%
				Next
			%>
				</select>
			</td>
			<td align="left" >Rate</td>
			<td align="left">
				<input type="text" maxlength="50" id="ratemandays" name="ratemandays" size="20" value="<%=ratemandays%>" autocomplete="Off" readonly style ="background-color: #eee;">
			</td>
		</tr>
		<tr>
			<td align="left" >Activity Code</td>
			<td align="left">
				<div id="jobcodelist"></div> 
			</td>
			<input type="hidden" id ="selectedactivity" name="selectedactivity" value="<%=selectedactivity%>">
		</tr>
		<tr>
			
		</tr>
		<tr>
			
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
					<a href="salesagreementheaderedit.asp?salescontractno=<%=salescontractno%>&creationdate=<%=creationdate%>&deliverydate=<%=deliverydate%>&fcba=<%=fcba%>&dp_beforevat=<%=dp_beforevat%>&dp_vat=<%=dp_vat%>&dp_aftervat=<%=dp_aftervat%>&totalamount=<%=totalamount%>&percentage_dp=<%=percentage_dp%>&amount_dpvat=<%=amount_dpvat%>" target="_blank"> Edit Header </a>
					<a href="MonthlyWorkPlanHarv.asp?salescontractno=<%=salescontractno%>&creationdate=<%=creationdate%>&deliverydate=<%=deliverydate%>&fcba=<%=fcba%>&txtJenis=HeaderDelete" onClick="return confirm('Are you sure to delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0">
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
if cmdView = "True" and len(trim(transno)) > 0  Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= 	" SELECT mwp.rkbno," & _
			"        mwp.fcba," & _
			"        mwp.period_year," & _
			"        mwp.period_month," & _
			"        mwp.activitycode," & _
			"        job_api.get_jobcode_name(mwp.activitycode,mwp.fcba) activityname," & _
			"        mwp.rate," & _
			"        mwd.unitkerja," & _
			"        mwd.fieldcode," & _
			"        mwd.plantingyear," & _
			"        mwd.hectarageplanted," & _
			"        mwd.plantingpoint," & _
			"        mwd.harvestingbased_abw_budget," & _
			"        mwd.rotation," & _
			"        mwd.kg_budget," & _
			"        mwd.bunch," & _
			"        mwd.normamandays," & _
			"        mwd.mandays," & _
			"        mwd.amountpermandays," & _
			"        mwd.outputperbunch," & _
			"        mwd.premi_luasan," & _
			"        mwd.basismandays," & _
			"        mwd.premi_basis," & _
			"        mwd.extra_basis_bunch," & _
			"        mwd.premi_extra_basis," & _
			"        mwd.total_premi," & _
			"        mwd.total_amount_panen," & _
			"        mwd.total_amountperkg," & _
			"        mwd.percent_lb," & _
			"        mwd.percent_sb," & _
			"        mwd.lineno" & _
			"   FROM MONTHLYWORKPLAN_PANENDET mwd" & _
			"        INNER JOIN MONTHLYWORKPLAN_PANEN mwp" & _
			"           ON (mwd.fcba = mwp.fcba AND mwd.rkbno = mwp.rkbno)" & _
			"  WHERE mwd.fcba = '"& fcba &"' AND mwd.rkbno = '"& transno &"'"

	'response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<form action="MonthlyWorkPlanHarv.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">  
			<tr>
				<td align="left" >BJR Budget</td>
				<td align="left">
					<input type="text" maxlength="50" id="" name="harvestingbased_abw_budget" size="25" value="<%=harvestingbased_abw_budget%>" onkeyup="hanyaAngka(this);hitungrumus();">
				</td> 
			</tr> 
			<tr>
				<td align="left" >KG Budget</td>
				<td align="left">
					<input type="text" maxlength="50" id="" name="kg_budget" size="25" value="<%=kg_budget%>" onkeyup="hanyaAngka(this);hitungrumus();">
				</td>   
			</tr>
			<tr>
				<td align="left" >SB(%)</td>
				<td align="left">
					<input type="text" maxlength="50" id="percent_sb" name="percent_sb" size="25" value="<%=percent_sb%>" onkeyup="hanyaAngka(this);hitungrumus();">
				</td> 
			</tr>
			<tr>
				<td align="left" >LB(%)</td>
				<td align="left">
					<input type="text" maxlength="50" id="percent_lb" name="percent_lb" size="25" value="<%=percent_lb%>" onkeyup="hanyaAngka(this);hitungrumus();">
				</td> 
			</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="transno" value="<%=transno%>">
					<input type="hidden" name="txtyear" value="<%=txtyear%>">
					<input type="hidden" name="txtMonth" value="<%=txtMonth%>">
					<input type="hidden" name="activitycode" value="<%=activitycode%>">
					<input type="hidden" name="ratemandays" value="<%=ratemandays%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					<input type="hidden" name="unitkerja" value="<%=unitkerja%>">
					<input type="hidden" name="fieldcode" value="<%=fieldcode%>">
					<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
					<input type="hidden" name="selectedfieldcode" value="<%=selectedfieldcode%>">
					<input type="hidden" name="selectedcustomercode" value="<%=selectedcustomercode%>">
					<input type="hidden" name="selecteddeliverytype" value="<%=selecteddeliverytype%>">
					<input type="hidden" name="selectedactivity" value="<%=selectedactivity%>">
					<input type="hidden" name="selectedmonth" value="<%=selectedmonth%>">
					<input type="hidden" name="selectedyear" value="<%=selectedyear%>">
					
					
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					
						<%IF Request("txtJenis")="Edit" then%>
							<input type="hidden" name="txtJenis" value="Edit">
							<input type="hidden" name="lineno" value="<%=lineno%>">
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
		<td align="center">NO</td>
		<td colspan="2">&nbsp;</td>
		<td align="center">RKB NO</td>
		<td align="center">YEAR</td>
		<td align="center">MONTH</td>
		<td align="center">ACTIVITY CODE</td>
		<td align="center">ACTIVITY NAME</td>
		<td align="center">RATE</td> 
		<td align="center">UNIT KERJA</td>
		<td align="center">FIELD CODE</td>
		<td align="center">BJR BUDGET</td> 
		<td align="center">KG BUDGET</td> 
		<td align="center">SB(%)</td> 
		<td align="center">LB(%)</td>   
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
			<td align="center" valign="top"><%=NO%></td>
			<td>
				<a href="MonthlyWorkPlanHarv.asp?lineno=<%=rsVH("lineno")%>&transno=<%=rsVH("rkbno")%>&txtMonth=<%=rsVH("period_month")%>&txtyear=<%=rsVH("period_year")%>&unitkerja=<%=rsVH("unitkerja")%>&fieldcode=<%=rsVH("fieldcode")%>&percent_lb=<%=rsVH("percent_lb")%>&percent_sb=<%=rsVH("percent_sb")%>&harvestingbased_abw_budget=<%=rsVH("harvestingbased_abw_budget")%>&kg_budget=<%=rsVH("kg_budget")%>&fcba=<%=rsVH("fcba")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
			</td>
			<td>
				<a href="MonthlyWorkPlanHarv.asp?lineno=<%=rsVH("lineno")%>&transno=<%=rsVH("rkbno")%>&txtMonth=<%=rsVH("period_month")%>&txtyear=<%=rsVH("period_year")%>&unitkerja=<%=rsVH("unitkerja")%>&fieldcode=<%=rsVH("fieldcode")%>&percent_lb=<%=rsVH("percent_lb")%>&percent_sb=<%=rsVH("percent_sb")%>&harvestingbased_abw_budget=<%=rsVH("harvestingbased_abw_budget")%>&kg_budget=<%=rsVH("kg_budget")%>&fcba=<%=rsVH("fcba")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
			    <td align="center" title="<%=rsVH("rkbno").Name%>" valign="top"><%=rsVH("rkbno")%></td>
			    <td align="center" title="<%=rsVH("period_year").Name%>" valign="top"><%=rsVH("period_year")%></td>
			    <td align="center" title="<%=rsVH("period_month").Name%>" valign="top"><%=rsVH("period_month")%></td>
			    <td align="center" title="<%=rsVH("activitycode").Name%>" valign="top"><%=rsVH("activitycode")%></td>
			    <td align="center" title="<%=rsVH("activityname").Name%>" valign="top"><%=rsVH("activityname")%></td>
			    <td align="center" title="<%=rsVH("rate").Name%>" valign="top"><%=rsVH("rate")%></td>
			    <td align="center" title="<%=rsVH("unitkerja").Name%>" valign="top"><%=rsVH("unitkerja")%></td>
			    <td align="center" title="<%=rsVH("fieldcode").Name%>" valign="top"><%=rsVH("fieldcode")%></td>
			    <td align="right" title="<%=rsVH("harvestingbased_abw_budget").Name%>" valign="top"><%=formatnumber(rsVH("harvestingbased_abw_budget"),2)%></td>
			    <td align="right" title="<%=rsVH("kg_budget").Name%>" valign="top"><%=formatnumber(rsVH("kg_budget"),0)%></td>
			    <td align="right" title="<%=rsVH("percent_sb").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("percent_sb"))),2)%></td>
			    <td align="right" title="<%=rsVH("percent_lb").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("percent_lb"))),2)%></td>
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

function getselectedcustomer() {
	var scustomercode = document.getElementById("selectedcustomercode");
	scustomercode.value = document.getElementById("customercode").value;
}

function getselecteddeliverytype() {
	var sdeliverytype = document.getElementById("selecteddeliverytype");
	sdeliverytype.value = document.getElementById("deliverytype").value;
}

function getselectedactivity() {
	var sactivitycode = document.getElementById("selectedactivity");
	sactivitycode.value = document.getElementById("activitycode").value;
}

function getselectedsection() {
	var sunitkerja = document.getElementById("selectedsection");
	sunitkerja.value = document.getElementById("unitkerja").value;
}

function getselectedfieldcode() {
	var sfieldcode = document.getElementById("selectedfieldcode");
	sfieldcode.value = document.getElementById("fieldcode").value;
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
	getListsection();
	getListJob();
	getListfieldcode();
	getNoRKB();
	getHKRate();
	getNormaHK();
	getFieldInfo();
	hitungrumus();
	getselectedsection();
	getselectedfieldcode();
	getselectedactivity();
});
</script>
</body>
</html>
