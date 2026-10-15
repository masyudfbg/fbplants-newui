<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Sales Agreement</title>

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
		var squantity = $("#quantity").val();
		var squantitytolerance = $("#quantitytolerance").val();
		var squantityaftertolerance = $("#quantityaftertolerance").val();
		var ssalesprices = $("#salesprices").val();
		var svat = $("#vat").val();
		var newssalesprices = ssalesprices.replace(/,/g, '');
		
		//Untuk menghitung Total Quantity Tolerance
		totalquantitytolerance = squantity * squantitytolerance / 100; //a kali b
		if (!isNaN(totalquantitytolerance)) {
    	 	$("#quantitytolerancetotal").val(totalquantitytolerance);
    	}
    	else{
    		$("#quantitytolerancetotal").val();
    	}
		
		//Untuk menghitung Total Quantity After Tolerance
		totalquantityaftertolerance = parseInt(squantity) + parseInt(totalquantitytolerance);
		if (!isNaN(totalquantityaftertolerance)) {
    	 	$("#quantityaftertolerance").val(totalquantityaftertolerance);
    	}
    	else{
    		$("#quantityaftertolerance").val();
    	}
		
		//Untuk menghitung Amount
		totalamount = squantity * newssalesprices;
		if (!isNaN(totalamount)) {
    	 	$("#amount").val(totalamount);
    	}
    	else{
    		$("#amount").val();
    	}
		
		//Untuk menghitung Amount Vat
		totalvat = totalamount * svat / 100;
		if (!isNaN(totalvat)) {
    	 	$("#amountvat").val(totalvat);
    	}
    	else{
    		$("#amountvat").val();
    	}
		
		//Untuk menghitung Total Amount 
		total = parseInt(totalamount) + parseInt(totalvat);
		if (!isNaN(total)) {
    	 	$("#totalamountvat").val(total);
    	}
    	else{
    		$("#totalamountvat").val();
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
	var fcba = document.getElementById("fcba").value;

	var productcode = document.getElementById("productcode").value;
	var PostStr = 'productcode='+productcode+'&fcba='+fcba;
	
	// use the generic function to make the request
	doAJAXCall('ajax/GetProductUom.asp', 'POST', PostStr, showMessageResponse);
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    
    // update the Div to show the result from the server
	document.getElementById("unitofmeasure").value = response;
}

// launched from button click 
var getLastNumberSalesContractno = function () {
	// build up the post string when passing variables to the server side page
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getLastSalesContractno.asp', 'POST', PostStr, showLastNumberSalesContractno);
}

// The function for handling the response from the server
var showLastNumberSalesContractno = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	//alert(s.lastnumber);
    // update the Div to show the result from the server
	document.getElementById("lastnumber").value = s.lastnumber;
}

// launched from button click 
var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbasalescontract.asp', 'POST', PostStr, showlistfcba);
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
var getlistcustomer = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedcustomercode = document.getElementById("selectedcustomercode").value;
	var PostStr = 'fcba='+selectedfcba+'&customercode='+selectedcustomercode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistcustomercodesalesbasedonfcba.asp', 'POST', PostStr, showlistcustomer);
}

// The function for handling the response from the server
var showlistcustomer = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("customercodelist").innerHTML = "";
	document.getElementById("customercodelist").innerHTML = response;

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

var getlistdeliverytype = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selecteddeliverytype = document.getElementById("selecteddeliverytype").value;
	var PostStr = 'fcba='+selectedfcba+'&deliverytype='+selecteddeliverytype;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getdeliverytypebasedonfcba.asp', 'POST', PostStr, showlistdeliverytype);
}

// The function for handling the response from the server
var showlistdeliverytype = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("deliverytypelist").innerHTML = "";
	document.getElementById("deliverytypelist").innerHTML = response;

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
var getIndikator = function () {
	// build up the post string when passing variables to the server side page
	var selectedfcba = document.getElementById("selectedfcba").value;
	var indikator = document.getElementById("indikator").value;
	var PostStr = 'fcba='+selectedfcba+'&indikator='+indikator;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getIndikatorSales.asp', 'POST', PostStr, showIndikator);
}

// The function for handling the response from the server
var showIndikator = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	//alert(s.lastnumber);
    // update the Div to show the result from the server
	document.getElementById("ffa").value = s.ffa;
	document.getElementById("moisture").value = s.moisture;
	document.getElementById("dirt").value = s.dirt;
	document.getElementById("dobi").value = s.dobi;
}


</script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
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
Dim txtDocumentNo ,txtlineno,ScriptName
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

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
ffa =request("ffa")
moisture =request("moisture")
dirt =request("dirt")
dobi =request("dobi")

fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
selectedcustomercode = Request("selectedcustomercode")
selecteddeliverytype = Request("selecteddeliverytype")

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

'response.write Request("txtJenis")

If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
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
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q
		else
			err.number = 1001
			err.description = " Sales Contract No """ & salescontractno & """ has detail, please delete all detail first!"
		End If
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
 	End if
	
	if Request("txtJenis")= "Delete" then

			strFieldDetail =	"select quantity" & _
								"		from salesagreementdet " & _
								" 		where lineno = '" & txtlineno & "'" & _
								"       and salescontractno = '" & salescontractno & "'" & _
								"       and fcBA = '" & fcba & "'"								
										' Response.Write strFieldDetail
										Set rsQty = Server.CreateObject("ADODB.Recordset")
										rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
			do while not rsQty.eof
					editquantity = cdbl (to_nol(rsQty("quantity")))
					rsQty.MoveNext
				loop
			End If
			
			q = " delete from salesagreementdet " & _
				" where fcBA = '" & fcba & "'" & _
				" and salescontractno = '"& salescontractno & "'" & _
				" and lineno = '"& txtlineno & "'"
				' Response.Write q
				connStr.execute(q)
				WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q
			
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
				dp_aftervat = cdbl(rsqLineMax("dp_aftervat"))
			end if 

			v_totalamount = v_totalamount + totalamountvat
			if cdbl(to_nol(v_totalamount))  = 0 then v_percentage_dp = 0 else v_percentage_dp = (dp_aftervat / v_totalamount) * 100

			q =	" Update salesagreementhdr" & _
				" set totalamount = replace('" & v_totalamount & "',',','')" & _
				"     , percentage_dp = replace('" & v_percentage_dp & "',',','.')" & _
				" 		where salescontractno = '" & salescontractno & "'" & _
				"       and fcBA = '" & fcba & "'"
			' response.write q
			connStr.execute(q)
			WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q
			
	End If 
	
	if Request("txtJenis")="New" then
	    	Dim rsSupplyingDeclarationHdr
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select salescontractno from salesagreementhdr where salescontractno = '" & salescontractno & "' and fcba ='" & fcba & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then
				q = " INSERT into salesagreementhdr" & _
					" (salescontractno, wbregistered, invoicedpno, creationdate, deliverydate, customercode, deliveryaddress, deliverytype, dp_vat, dp_beforevat, dp_aftervat, amount_dpvat, currency, paymentcondition, deliverycondition, remarks, mill,fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" '" & salescontractno & "'" & _
					" , 'FALSE'" & _
					" , '" & invoicedpno & "'" & _
					" , to_date('" & creationdate & "','mm/dd/yyyy')" & _ 
					" , to_date('" & deliverydate & "','mm/dd/yyyy')" & _ 
					" , '" & customercode & "'" & _
					" , '" & deliveryaddress & "'" & _
					" , '" & deliverytype & "'" & _
					" , replace('" & dp_vat & "',',','')" & _
					" , replace('" & dp_beforevat & "',',','')" & _
					" , replace('" & dp_aftervat & "',',','')" & _
					" , replace('" & amount_dpvat & "',',','')" & _
					" , '" & curr & "'" & _
					" , '" & paymentcondition & "'" & _
					" , '" & deliverycondition & "'" & _
					" , '" & remarks & "'" & _
					" , '" & mill & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & fcba & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"								
		
				connStr.execute(q)
				WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing
					
			If Len(Trim(productcode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Product ID Harus diisi, tidak boleh kosong"
			Else
			'If cdbl(to_nol(quantity)) = 0 Then
			'	Err.Number = 1001
			'	Err.Description = "Quantity Harus diisi, tidak boleh kosong"
			'Else
			If cdbl(to_nol(salesprices)) = 0 Then
				Err.Number = 1001
				Err.Description = "Sales Price Harus diisi, tidak boleh kosong"
			Else
			
					qLineMax = " select nvl(max(lineno),0) tempMaxLine" & _
							   " from salesagreementdet" & _
							   " where salescontractno = '" & salescontractno & "'" & _
							   " and fcba = '" & fcba & "'" 			

					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
					else
						tempMaxLine = 0
					end if 
					rsqLineMax.Close
					set rsqLineMax = Nothing

					tempMaxLine = tempMaxLine + 1
					txtLineNoHist = tempMaxLine
					
					' quantitytolerance = cdbl(to_nol(quantity)) *  cdbl(to_nol(quantity)) 
					' amount = cdbl(to_nol(quantity)) * cdbl(to_nol(salesprices))
					' amountvat = cdbl(to_nol(vat)) * cdbl(to_nol(amount))/100
					' totalamountvat = amount +amountvat
					
					q = " insert into salesagreementdet" & _
						" (productcode, salescontractno, unitofmeasure, ffa, moisture, dirt, dobi, startdate, enddate, quantity, salesprices, amount, vat, amountvat, totalamountvat, quantitytolerance, quantitytolerancetotal, quantityaftertolerance, lineno, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
						" '" & productcode & "'" & _
						" , '" & salescontractno & "'" & _
						" , '" & unitofmeasure & "'" & _
						" , '" & to_nol(ffa) & "'" & _
						" , '" & to_nol(moisture) & "'" & _
						" , '" & to_nol(dirt) & "'" & _
						" , '" & to_nol(dobi) & "'" & _
						" , to_date('" & startdate & "','mm/dd/yyyy')" & _ 
						" , to_date('" & enddate & "','mm/dd/yyyy')" & _ 
						" , replace('" & to_nol(quantity) & "',',','')" & _
						" , replace('" & to_nol(salesprices) & "',',','')" & _
						" , replace('" & to_nol(amount) & "',',','')" & _
						" , replace('" & to_nol(vat) & "',',','')" & _
						" , replace('" & to_nol(amountvat) & "',',','')" & _
						" , replace('" & to_nol(totalamountvat) & "',',','')" & _
						" , replace('" & to_nol(quantitytolerance) & "',',','')" & _
						" , replace('" & to_nol(quantitytolerancetotal) & "',',','')" & _
						" , replace('" & to_nol(quantityaftertolerance) & "',',','')" & _
						" , '" & tempMaxLine & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						" , '" & fcba & "'" & _
						" , sysdate" & _
						" , to_char(sysdate, 'hh24:mi')" & _
						" )"
						' response.write q
						connStr.execute(q)
						WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q

					v_totalamount = 0
					qLineMax = " select nvl(SUM(totalamountvat),0) totalamountvat" & _
							   " from salesagreementdet" & _
							   " where salescontractno = '" & salescontractno & "'" & _
							   " and fcba = '" & fcba & "'" 			

					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						totalamountvat = cdbl(rsqLineMax("totalamountvat"))
					else	
						totalamountvat = 0
					end if 
					rsqLineMax.Close
					set rsqLineMax = Nothing

					qLineMax = " select nvl(SUM(dp_aftervat),0) dp_aftervat" & _
							   " from salesagreementhdr" & _
							   " where salescontractno = '" & salescontractno & "'" & _
							   " and fcba = '" & fcba & "'" 			

					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						dp_aftervat = cdbl(rsqLineMax("dp_aftervat"))
					else
						dp_aftervat = 0
					end if 
					rsqLineMax.Close
					set rsqLineMax = Nothing
					
					v_totalamount = v_totalamount + totalamountvat

					if cdbl(v_totalamount) = 0 then
						v_percentage_dp = 0
					else
						v_percentage_dp = (cdbl(dp_aftervat) / cdbl(v_totalamount)) * 100
					end if

					If dp_aftervat > v_totalamount Then
						Err.Number = 1001
						Err.Description = "DP Tidak boleh lebih dari Total Amount !"

						' q = " delete from salesagreementdet " & _
							' " where fcBA = '" & fcba & "'" & _
							' " and salescontractno = '"& salescontractno & "'" 
							' Response.Write q
							' connStr.execute(q)
				
					Else			
						q =	" Update salesagreementhdr" & _
							" set totalamount = replace('" & v_totalamount & "',',','')" & _
							"     , percentage_dp = '" & v_percentage_dp & "'" & _
							" 		where salescontractno = '" & salescontractno & "'" & _
							"       and fcBA = '" & fcba & "'"
						connStr.execute(q)
						WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q
					End if
			End If			
			'End if
	
	End if 
	
	if Request("txtJenis")="Edit" and Request("Add")="Update" then
 			'If cdbl(to_nol(quantity)) = 0 Then
			'	Err.Number = 1001
			'	Err.Description = "Quantity Harus diisi, tidak boleh kosong"
			'Else
			' If cdbl(to_nol(amount)) = 0 Then
			' 	Err.Number = 1001
			' 	Err.Description = "Amount Harus diisi, tidak boleh kosong"
			' Else
			If cdbl(to_nol(salesprices)) = 0 Then
				Err.Number = 1001
				Err.Description = "Sales Price Harus diisi, tidak boleh kosong"
			Else
			' If cdbl(to_nol(VAT)) = 0 Then
			' 	Err.Number = 1001
			' 	Err.Description = "VAT Harus diisi, tidak boleh kosong"
			' Else
					' amount = cdbl(to_nol(quantity)) * cdbl(to_nol(salesprices))
					' amountvat = cdbl(to_nol(vat)) * cdbl(to_nol(amount))/100
					' totalamountvat = amount +amountvat
					
					q =	" Update salesagreementdet" & _
						" set productcode = '" & productcode & "'" & _
						"     , salescontractno = '" & salescontractno & "'" & _
						"     , unitofmeasure = '" & unitofmeasure & "'" & _
						"     , ffa = '" & ffa & "'" & _
						"     , moisture = '" & moisture & "'" & _
						"     , dobi = '" & dobi & "'" & _
						"     , dirt = '" & dirt & "'" & _ 
						"     , startdate = to_date('"& startdate &"','mm/dd/yyyy')" & _
						"	  , enddate = to_date('"& enddate &"','mm/dd/yyyy')" & _
						"     , quantity = replace('" & quantity & "',',','')" & _
						"     , salesprices = replace('" & salesprices & "',',','')" & _
						"     , amount = replace('" & amount & "',',','')" & _
						"     , vat = replace('" & vat & "',',','')" & _
						"     , amountvat = replace('" & amountvat & "',',','')" & _
						"     , totalamountvat = replace('" & totalamountvat & "',',','')" & _
						"     , quantitytolerance = replace('" & quantitytolerance & "',',','')" & _
						"     , quantityaftertolerance = replace('" & quantityaftertolerance & "',',','')" & _
						"     , quantitytolerancetotal = replace('" & quantitytolerancetotal & "',',','')" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where salescontractno = '" & salescontractno & "'" & _
						"       and lineno = '" & txtlineno & "'" & _
						"       and fcBA = '" & fcba & "'"
		 			'response.write q
					connStr.execute(q)
					WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q

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
						dp_aftervat = cdbl(rsqLineMax("dp_aftervat"))
					end if 
					
					v_totalamount = v_totalamount + totalamountvat
					if cdbl(to_nol(v_totalamount))  = 0 then v_percentage_dp = 0 else v_percentage_dp = (dp_aftervat / v_totalamount) * 100
					' response.write v_percentage_dp

						If dp_aftervat > v_totalamount Then
							Err.Number = 1001
							Err.Description = "DP Tidak boleh lebih dari Total Amount !"
							
							q = " delete from salesagreementdet " & _
								" where fcBA = '" & fcba & "'" & _
								" and salescontractno = '"& salescontractno & "'" 
								' Response.Write q
								connStr.execute(q)
								WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q
						Else			
							q =	" Update salesagreementhdr" & _
								" set totalamount = replace('" & v_totalamount & "',',','')" & _
								"     , percentage_dp = '" & v_percentage_dp & "'" & _
								" 		where salescontractno = '" & salescontractno & "'" & _
								"       and fcBA = '" & fcba & "'"
							' response.write q
							connStr.execute(q)
							WriteLog session("username") & " Execution at: " & Now() &  " " & ScriptName & " Query executed: " & q
						End if
					End If
			End If		
'			End If			
'		End If
' End If

    End If
	End If
	cmdView	= "True"	

'End If


Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

'untuk menampilkan datepicker default
If Len(Trim(creationdate)) = 0 Then	creationdate = textDate(date)
If Len(Trim(deliverydate)) = 0 Then	deliverydate = textDate(date)
If Len(Trim(startdate)) = 0 Then startdate = textDate(date)
If Len(Trim(enddate)) = 0 Then	enddate = textDate(date)



Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select sa.*, bu.fccompanycode" & _
							" from salesagreementhdr sa left join businessunit bu on bu.fccode = sa.fcba" & _
							" where sa.salescontractno = '" & salescontractno & "'"
						
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	creationdate = rsSupplyingDeclaration("creationdate")
	deliverydate = rsSupplyingDeclaration("deliverydate")
	customercode = rsSupplyingDeclaration("customercode")
	deliveryaddress = rsSupplyingDeclaration("deliveryaddress")
	deliverytype = rsSupplyingDeclaration("deliverytype")
	invoicedpno = rsSupplyingDeclaration("invoicedpno")
	dp_vat = rsSupplyingDeclaration("dp_vat")
	dp_beforevat = rsSupplyingDeclaration("dp_beforevat")
	dp_aftervat = rsSupplyingDeclaration("dp_aftervat")
	amount_dpvat = rsSupplyingDeclaration("amount_dpvat")
	percentage_dp = rsSupplyingDeclaration("percentage_dp")
	totalamount = rsSupplyingDeclaration("totalamount")
	curr = rsSupplyingDeclaration("currency")	
	paymentcondition = rsSupplyingDeclaration("paymentcondition")
	deliverycondition = rsSupplyingDeclaration("deliverycondition")
	remarks = rsSupplyingDeclaration("remarks")
	fccompanycode = rsSupplyingDeclaration("fccompanycode")
	fcba = rsSupplyingDeclaration("fcba")
	selectedfcba = rsSupplyingDeclaration("fcba")
	selectedcustomercode = rsSupplyingDeclaration("customercode")
	selecteddeliverytype = rsSupplyingDeclaration("deliverytype")
	mill = rsSupplyingDeclaration("mill")
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
 
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-briefcase" aria-hidden="true"></i>Header Sales Agreement</div><% end if %>
	<form name="frmEntry" method="post" action="salesagreement.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Sales Contract No</td>
			<td align="left">
			<input type="text" maxlength="50" name="salescontractno" id="salescontractno" style="background:#FFFFCC;" size="30" value="<%=salescontractno%>" autocomplete="Off">
			<font color="#FF0000">*</font>
			</td>
			<td align="left" width="20%">Mill</td>
			<td align="left">
			<input type="text" maxlength="15" name="mill" size="30" value="<%=mill%>" autocomplete="Off">
			<input type="hidden" maxlength="15" name="mill" id="sessionusername" size="30" value="<%=session("username")%>">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Last Number Sales Contract
				<td align="left">
				<input type="text" id="lastnumber" name="lastnumber" maxlength="15" style="background:#eee;" size="30" value="<%=lastnumber%>" readonly="true">
				</td>
			</td>
			<td align="left" width="20%">Currency</td>
			<td align="left">
				<select id="curr" name="curr" class="chosen-select" style="width: 50%">
				<option></option>
				<%
					qlov = "SELECT alphabetic_code fccode, currency fcname FROM currency WHERE upper(entity) = upper('indonesia')"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if curr = rslov("fccode") then%>selected<%end if%>><%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Creation Date</td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="creationdate" id="datecreation" value="<%=creationdate%>">
			</td>
			<td align="left" width="20%">Payment Condition</td>
			<td align="left">
			<textarea name="paymentcondition" size="30" ><%=paymentcondition%></textarea>
			</td>
		</tr>
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();" class="chosen-select" style="width: 70%">
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
			<td align="left" width="20%">Delivery Address</td>
			<td align="left">
			<textarea name="deliveryaddress"><%=deliveryaddress%></textarea>
			</td>

		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<div id="fcbalist"> </div>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			<td align="left"><font face="Verdana" size="1">Delivery Type</font></td>
			<td align="left">
			<div id="deliverytypelist"></div>
			<input type="Hidden" id="selecteddeliverytype" name="selecteddeliverytype" value="<%=selecteddeliverytype%>">
			</td> 
		</tr>		
		<tr>
			<td align="left" width="20%">Customer Code</td>
			<td align="left">
			<div id= "customercodelist"> </div>
				<input type="Hidden" id="selectedcustomercode" name="selectedcustomercode" value="<%=selectedcustomercode%>">
			</td>
			<td align="left" width="20%">Delivery Condition</td>
			<td align="left">
			<input type="text" maxlength="15" name="deliverycondition" size="30" value="<%=deliverycondition%>" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Delivery Date</td>
			<td align="left">
			<input class="datepicker" type="text" size="15" name="deliverydate" value="<%=deliverydate%>">
			</td>  
			<td align="left" width="20%">Total Amount Sales Contract</td>
			<td align="left">
			<input type="text" maxlength="15" name="totalamount" size="30" value="<%=totalamount%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly style="background-color:#eee;">
			</td>
		</tr> 
		<tr>
			
			<td align="left" width="20%">Remarks</td>
			<td align="left">
			<textarea name="remarks" cols="50"><%=remarks%></textarea>
			</td>
		</tr>
		<tr>
			<td colspan="4"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
					<a href="salesagreementheaderedit.asp?salescontractno=<%=salescontractno%>&creationdate=<%=creationdate%>&deliverydate=<%=deliverydate%>&fcba=<%=fcba%>&dp_beforevat=<%=dp_beforevat%>&dp_vat=<%=dp_vat%>&dp_aftervat=<%=dp_aftervat%>&totalamount=<%=totalamount%>&percentage_dp=<%=percentage_dp%>&amount_dpvat=<%=amount_dpvat%>" target="_blank"> Edit Header </a>
					<a href="salesagreement.asp?salescontractno=<%=salescontractno%>&creationdate=<%=creationdate%>&deliverydate=<%=deliverydate%>&fcba=<%=fcba%>&txtJenis=HeaderDelete" onClick="return confirm('Are you sure to delete');">
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


if cmdView = "True" and len(trim(salescontractno)) > 0  Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= " SELECT sad.salescontractno, sah.mill, sah.customercode, cs.description customername, sah.creationdate, sah.deliverydate, sah.deliverytype, sad.lineno, sah.remarks, sad.productcode, sad.unitofmeasure, " & _
		  " sad.salesprices, sad.startdate, sad.enddate, sad.amount, sad.amountvat, sad.totalamountvat, sad.quantity, sad.ffa, sad.moisture, sad.dirt, sad.dobi, sad.vat, pr.FCNAME as deliverytype," & _
		  " sad.quantitytolerance, sad.quantityaftertolerance, sad.quantitytolerancetotal" & _
		  " FROM salesagreementdet sad" & _
		  " INNER JOIN salesagreementhdr sah"& _
		  " ON (sad.salescontractno = sah.salescontractno AND sad.fcba = sah.fcba)" & _
		  " LEFT JOIN parameter pr ON (sad.fcba = pr.fcba AND sah.deliverytype = pr.fccode AND pr.parcode = 'DELIVERY_TYPE')" & _
		  " LEFT JOIN customer cs ON (cs.fcba = sah.fcba AND cs.customercode = sah.customercode)" & _
		  " WHERE sad.fcba = '" & fcba & "'" & _
		  " AND sad.salescontractno = '" & salescontractno & "'" 

	'response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
	<form action="salesagreement.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr>
			<td align="left" width="20%">Product 
				<td align="left">
				<input type="hidden" maxlength="50" name="fcbasession" size="20" value="<%=Session("fcba")%>" autocomplete="Off">
				<select name="productcode" id="productcode" style="background:#FFFFCC;" class="chosen-select" title="Delivery Type CODE Character a s/d z or 1 s/d 9" onchange="javascript:getMessage();">
					<option value="<%=PRODUCTCODE%>"></option> 
						<% 	Dim rsComboP
							q = "SELECT * FROM PRODUCT WHERE fcba = '" & fcba & "' ORDER BY productcode"
							Set rsComboP = Server.CreateObject("ADODB.Recordset")
							rsComboP.CursorType = 3
							rsComboP.CursorLocation = 3
							rsComboP.Open q, connstr
							if not rsComboP.eof then%>
							<%
							Do While Not rsComboP.eof
							%>
								<option value="<%=rsComboP("PRODUCTCODE")%>" 
									<%if cstr(""&productcode)=cstr(""&rsComboP("PRODUCTCODE")) then %> selected 
									<%End If%>><%=rsComboP("PRODUCTDESCRIPTION")%>
								</option> 
							<%rsComboP.movenext
							loop
							End if
							rsComboP.Close
							set rsComboP = Nothing
							q=""
						%>
				</select> 
				<font color="#FF0000">*</font>
				</td>
			</td>

			<td align="left" width="20%">Unit Of Measure
				<td align="left">
					<input type="text" id="unitofmeasure" name="unitofmeasure" maxlength="15" style="background:#eee;" size="30" value="<%=unitofmeasure%>" readonly="true">
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Start Date 
			<td align="left">
				<input class="datepicker" type="text" size="15" name="startdate" value="<%=startdate%>">
			</td>

			<td align="left" width="20%">End Date 
				<td align="left">
					<input type="text" class="datepicker" maxlength="15" name="enddate" size="30" value="<%=enddate%>" autocomplete="Off">
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Indikator
				<td align="left">
					<select name="indikator" id="indikator" style="background:#FFFFCC;" class="chosen-select" title="Indikator" onchange="javascript:getIndikator();">
					<option value="<%=indikator%>"></option> 
						<% 	q = "SELECT fccode, fcname FROM millquality WHERE fcba = '" & fcba & "' ORDER BY fccode"
							Set rsComboP = Server.CreateObject("ADODB.Recordset")
							rsComboP.CursorType = 3
							rsComboP.CursorLocation = 3
							rsComboP.Open q, connstr
							if not rsComboP.eof then%>
							<%
							Do While Not rsComboP.eof
							%>
								<option value="<%=rsComboP("fccode")%>" 
									<%if cstr(""&indikator)=cstr(""&rsComboP("fccode")) then %> selected 
									<%End If%>><%=rsComboP("fccode")%>
								</option> 
							<%rsComboP.movenext
							loop
							End if
							rsComboP.Close
							set rsComboP = Nothing
							q=""
						%>
				</select> 
				</td>
			</td>
			<td align="left" width="20%">FFA
				<td align="left">
					<input type="text" maxlength="15" id="ffa" name="ffa" style="background:#FFFFCC;" size="30" value="<%=ffa%>" autocomplete="Off" onkeyup="hanyaAngka(this);hitungrumus();">
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Moisture & Impurity (Max) %
				<td align="left">
					<input type="text" maxlength="15" id="moisture" name="moisture" style="background:#FFFFCC;" size="30" value="<%=moisture%>" autocomplete="Off" onkeyup="hanyaAngka(this);hitungrumus();">
				</td>
			</td>
			<td align="left" width="20%">Dirt
				<td align="left">
				<input type="text" maxlength="15" id="dirt" name="dirt" style="background:#FFFFCC;" size="30" value="<%=dirt%>" autocomplete="Off" onkeyup="hanyaAngka(this);hitungrumus();" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Quantity
				<td align="left">
				<input type="text" maxlength="15" id="quantity" name="quantity" style="background:#FFFFCC;" size="30" value="<%=quantity%>" autocomplete="Off" onkeyup="hanyaAngka(this);hitungrumus();">
				<font color="#FF0000">*</font>
				</td>
			</td>
			<td align="left" width="20%">DOBI
				<td align="left">
				<input type="text" maxlength="15" id="dobi" name="dobi" style="background:#FFFFCC;" size="30" value="<%=dobi%>" autocomplete="Off" onkeyup="hanyaAngka(this);hitungrumus();" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Quantity Tolerance
				<td align="left">
				<input type="text" maxlength="15" id="quantitytolerance" name="quantitytolerance" style="background:#FFFFCC;" size="5" value="<%=quantitytolerance%>" autocomplete="Off" onkeyup="hanyaAngka(this);hitungrumus();">
				<font color="#FF0000">%</font>
				<input type="text" maxlength="15" id="quantitytolerancetotal" name="quantitytolerancetotal" style="background:#eee;" size="16" value="<%=quantitytolerancetotal%>"  autocomplete="Off" onkeyup="hitungrumus();" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
				</td>
			</td>
			<td align="left" width="20%">Price
				<td align="left">
				<input type="text" maxlength="15" id="salesprices" name="salesprices" style="background:#FFFFCC;" size="30" value="<%=salesprices%>" autocomplete="Off" onkeyup="hanyaAngka(this);hitungrumus();" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
				<font color="#FF0000">*</font>
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Quantity After Tolerance 
				<td align="left">
					<input type="text" maxlength="15" id="quantityaftertolerance" name="quantityaftertolerance" style="background:#eee;" size="30" value="<%=quantityaftertolerance%>" readonly autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
				</td>
			</td>
			<td align="left" width="20%">VAT 
				<td align="left">
				<input type="text" maxlength="15" id="vat" name="vat" size="5" style="background:#FFFFCC;" value="<%=vat%>" autocomplete="Off" onkeyup="hanyaAngka(this);hitungrumus();">
				<font color="#FF0000">%</font>
				<input type="text" maxlength="15" style="background:#eee;" id="amountvat" name="amountvat" size="16" value="<%=amountvat%>" autocomplete="Off" onkeyup="hitungrumus();" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Amount
				<td align="left">
				<input type="text" maxlength="15" id="amount" name="amount" style="background:#eee;" size="30" value="<%=amount%>" autocomplete="Off" onkeyup="hitungrumus();" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
				</td>
			</td>  
			<td align="left" width="20%">Total 
				<td align="left">
				<input type="text" maxlength="15" style="background:#eee;" id="totalamountvat" name="totalamountvat" size="30" value="<%=totalamountvat%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
				</td>
			</td>
		</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="salescontractno" value="<%=salescontractno%>">
					<input type="hidden" name="invoicedpno" value="<%=invoicedpno%>">
					<input type="hidden" name="creationdate" value="<%=creationdate%>">
					<input type="hidden" name="deliverydate" value="<%=deliverydate%>">
					<input type="hidden" name="customercode" value="<%=customercode%>">
					<input type="hidden" name="deliverytype" value="<%=deliverytype%>">
					<input type="hidden" name="deliveryaddress" value="<%=deliveryaddress%>">
					<input type="hidden" name="dp_aftervat" value="<%=dp_aftervat%>">
					<input type="hidden" name="dp_beforevat" value="<%=dp_beforevat%>">
					<input type="hidden" name="dp_vat" value="<%=dp_vat%>">
					<input type="hidden" name="percentage_dp" value="<%=percentage_dp%>">
					<input type="hidden" name="totalamount" value="<%=totalamount%>">
					<input type="hidden" name="amount_dpvat" value="<%=amount_dpvat%>">
					<input type="hidden" name="curr" value="<%=curr%>">
					<input type="hidden" name="paymentcondition" value="<%=paymentcondition%>">
					<input type="hidden" name="deliverycondition" value="<%=deliverycondition%>">
					<input type="hidden" name="remarks" value="<%=remarks%>">
					<input type="hidden" name="mill" value="<%=mill%>">
					<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
					<input type="hidden" name="selectedcustomercode" value="<%=selectedcustomercode%>">
					<input type="hidden" name="selecteddeliverytype" value="<%=selecteddeliverytype%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					
						<%IF Request("txtJenis")="Edit" then%>
							<input type="hidden" name="txtJenis" value="Edit">
							<input type="hidden" name="txtlineno" value="<%=txtlineno%>">
							<input Type="Submit"  name="Add" value="Update" id="insert">
						<%Else%>
							<input type="hidden" name="txtJenis" value="New">
							<input Type="Submit"  name="Add" value="Add" id="insert">
						<%End if%>
						<input Type="Reset"  name="Reset" value="Reset">
						<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
					
				</td>
			</tr>
		</table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="center">NO</td>
		<td colspan="2">&nbsp;</td>
		<td align="center">SALES CONTRACT NO</td>
		<td align="center">CUSTOMER CODE</td>
		<td align="center">MILL</td>
		<td align="center">DELIVERY TYPE</td>
		<td align="center">START DATE</td>
		<td align="center">END DATE</td>
		<td align="center">PRODUCT NAME</td>
		<td align="center">QUANTITY</td>
		<td align="center">FFA</td>
		<td align="center">MOISTURE</td>
		<td align="center">DIRT</td>
		<td align="center">DOBI</td>
		<td align="center">QUANTITY TOLERANCE (%)</td>
		<td align="center">QUANTITY TOLERANCE</td>
		<td align="center">QUANTITY AFTER TOLERANCE</td>
		<td align="center">UOM</td>
		<td align="center">SALES PRICES</td>
		<td align="center">AMOUNT</td>
		<td align="center">VAT</td>
		<td align="center">AMOUNT VAT</td>
		<td align="center">TOTAL AMOUNT</td>	
		<td align="center">REMARKS</td>	
	</tr>
	</thead>
	<tbody>
   
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
				<a href="salesagreement.asp?txtlineno=<%=rsVH("lineno")%>&salescontractno=<%=salescontractno%>&customercode=<%=rsVH("customercode")%>&deliverytype=<%=deliverytype%>&productcode=<%=rsVH("productcode")%>&ffa=<%=rsVH("ffa")%>&moisture=<%=rsVH("moisture")%>&dirt=<%=rsVH("dirt")%>&dobi=<%=rsVH("dobi")%>&startdate=<%=rsVH("startdate")%>&enddate=<%=rsVH("enddate")%>&remarks=<%=rsVH("remarks")%>&dp_aftervat=<%=dp_aftervat%>&deliveryaddress=<%=deliveryaddress%>&curr=<%=curr%>&paymentcondition=<%=paymentcondition%>&deliverycondition=<%=deliverycondition%>&unitofmeasure=<%=rsVH("unitofmeasure")%>&salesprices=<%=rsVH("salesprices")%>&amount=<%=rsVH("amount")%>&quantity=<%=rsVH("quantity")%>&quantitytolerance=<%=rsVH("quantitytolerance")%>&quantitytolerancetotal=<%=rsVH("quantitytolerancetotal")%>&quantityaftertolerance=<%=rsVH("quantityaftertolerance")%>&vat=<%=rsVH("vat")%>&amountvat=<%=rsVH("amountvat")%>&totalamountvat=<%=rsVH("totalamountvat")%>&fcba=<%=fcba%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
			</td>
			<td>
				<a href="salesagreement.asp?txtlineno=<%=rsVH("lineno")%>&salescontractno=<%=salescontractno%>&customercode=<%=rsVH("customercode")%>&deliverytype=<%=deliverytype%>&productcode=<%=rsVH("productcode")%>&ffa=<%=rsVH("ffa")%>&moisture=<%=rsVH("moisture")%>&dirt=<%=rsVH("dirt")%>&dobi=<%=rsVH("dobi")%>&startdate=<%=rsVH("startdate")%>&enddate=<%=rsVH("enddate")%>&remarks=<%=rsVH("remarks")%>&dp_aftervat=<%=dp_aftervat%>&deliveryaddress=<%=deliveryaddress%>&curr=<%=curr%>&paymentcondition=<%=paymentcondition%>&deliverycondition=<%=deliverycondition%>&unitofmeasure=<%=rsVH("unitofmeasure")%>&salesprices=<%=rsVH("salesprices")%>&amount=<%=rsVH("amount")%>&quantity=<%=rsVH("quantity")%>&quantitytolerance=<%=rsVH("quantitytolerance")%>&quantitytolerancetotal=<%=rsVH("quantitytolerancetotal")%>&quantityaftertolerance=<%=rsVH("quantityaftertolerance")%>&vat=<%=rsVH("vat")%>&amountvat=<%=rsVH("amountvat")%>&totalamountvat=<%=rsVH("totalamountvat")%>&fcba=<%=fcba%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
			    <td align="center" title="<%=rsVH("salescontractno").Name%>" valign="top"><%=rsVH("salescontractno")%></td>
			    <td align="center" title="<%=rsVH("customercode").Name%>" valign="top"><%=rsVH("customercode")%> : <%=rsVH("customername")%></td>
			    <td align="center" title="<%=rsVH("mill").Name%>" valign="top"><%=rsVH("mill")%></td>
			    <td align="center" title="<%=rsVH("deliverytype").Name%>" valign="top"><%=rsVH("deliverytype")%></td>
			    <td align="center" title="<%=rsVH("startdate").Name%>" valign="top"><%=rsVH("startdate")%></td>
			    <td align="center" title="<%=rsVH("enddate").Name%>" valign="top"><%=rsVH("enddate")%></td>
			    <td align="center" title="<%=rsVH("productcode").Name%>" valign="top"><%=rsVH("productcode")%></td>
			    <td align="right" title="<%=rsVH("quantity").Name%>" valign="top"><%=formatnumber(rsVH("quantity"),2)%></td>
			    <td align="right" title="<%=rsVH("ffa").Name%>" valign="top"><%=formatnumber(rsVH("ffa"),2)%></td>
			    <td align="right" title="<%=rsVH("moisture").Name%>" valign="top"><%=formatnumber(rsVH("moisture"),2)%></td>
			    <td align="right" title="<%=rsVH("dirt").Name%>" valign="top"><%=formatnumber(rsVH("dirt"),2)%></td>
			    <td align="right" title="<%=rsVH("dobi").Name%>" valign="top"><%=formatnumber(rsVH("dobi"),2)%></td>
				<td align="right" title="<%=rsVH("quantitytolerance").Name%>" valign="top"><%=rsVH("quantitytolerance")%></td>
				<td align="right" title="<%=rsVH("quantitytolerancetotal").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("quantitytolerancetotal"))),2)%></td>
				<td align="right" title="<%=rsVH("quantityaftertolerance").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("quantityaftertolerance"))),2)%></td>
			    <td align="center" title="<%=rsVH("unitofmeasure").Name%>" valign="top"><%=rsVH("unitofmeasure")%></td>
			    <td align="right" title="<%=rsVH("salesprices").Name%>" valign="top"><%=formatnumber(rsVH("salesprices"),2)%></td>
			    <td align="right" title="<%=rsVH("amount").Name%>" valign="top"><%=formatnumber(rsVH("amount"),2)%></td>
			    <td align="right" title="<%=rsVH("vat").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("vat"))),2)%></td>
			    <td align="right" title="<%=rsVH("amountvat").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("amountvat"))),2)%></td>
			    <td align="right" title="<%=rsVH("totalamountvat").Name%>" valign="top"><%=formatnumber(rsVH("totalamountvat"),2)%></td>
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
   </tbody>
   </TABLE>
   </div>
<% if Request("oldui") <> "1" then %>
<script type="text/javascript">
// Status baris jadi badge (tampilan saja)
(function () {
	var cells = document.querySelectorAll('table.fbp-doc-lines tbody td[title]');
	for (var i = 0; i < cells.length; i++) {
		if (cells[i].getAttribute('title').toUpperCase() !== 'ROWSTATEID') continue;
		var s = (cells[i].textContent || '').replace(/\s+/g, ' ').trim();
		if (!s) continue;
		cells[i].parentNode.setAttribute('data-state', s);
		cells[i].innerHTML = '';
		var b = document.createElement('span');
		b.className = 'fbp-state';
		b.textContent = s;
		cells[i].appendChild(b);
	}
})();
</script>
<% end if %>
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

	$("#insert").on('click', function(){
		const fcbasession = document.querySelector('input[name="fcbasession"]').value;

		var settings = {
			"url": "http://localhost:8080/api/contracts",
			"method": "POST",
			"timeout": 0,
			"headers": {
				"Content-Type": "application/json"
			},
			"data": JSON.stringify({
				"contract_date": $("#datecreation").val(),
				"contract_no": $("#salescontractno").val(),
				"product_id": $("#productcode").val(),
				"contract_qty": $("#quantity").val(),
				"price_per_kg": $("#salesprices").val(),
				"currency_id": "IDR",
				"customer_id": $("#customercode").val(),
				"top": "Net 30",
				"dp_ptg": 0,
				"dp_ptg2": 0,
				"tax_ptg": parseInt($("#vat").val()),
				"invoice_flg": 0,
				"is_no_dp": 0,
				"add_price_per_kg": 0,
				"kwitansi_no": null,
				"kwitansi_date": $("#datecreation").val(),
				"user_id": $("#sessionusername").val(),
				"fcba": fcbasession
			}),
		};

		$.ajax(settings).done(function (response) {
			console.log(response);
		});
	});
	
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
	getlistcustomer();
	getlistdeliverytype();
	getLastNumberSalesContractno();
	getIndikator();
	getselectedfcba();
	getselectedcustomer();
	getselecteddeliverytype();
	hitungrumus();
});
</script>
</body>
</html>
