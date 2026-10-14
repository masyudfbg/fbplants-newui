<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Purchase Shipment</title>

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

// launched from button click 
var getDOList = function () {
	// build up the post string when passing variables to the server side page
	// var fccompanycode = document.getElementById("fccompanycode").value;
	var fcba = document.getElementById("fcba").value;
	var selecteddo = document.getElementById("selecteddo").value;
	var PostStr = 'deliveryorderno='+selecteddo+'&fcba='+fcba;
	//alert(PostStr);
	// use the generic function to make the request
	if (!$('#extracarry').is(":checked")){
			doAJAXCall('ajax/getPurchaseDOListShipment.asp', 'POST', PostStr, showdolist);
		}else{
			doAJAXCall('ajax/getPurchaseDOListShipmentExtraCarry.asp', 'POST', PostStr, showdolist);
		}
}

// The function for handling the response from the server
var showdolist = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("dolist").innerHTML = "";
	document.getElementById("dolist").innerHTML = response;

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
var getDOLine = function () {
	// build up the post string when passing variables to the server side page
	var fcba = document.getElementById("fcba").value;
	var selecteddo = document.getElementById("selecteddo").value;
	var selectedlineno = document.getElementById("selectedlineno").value;
	var PostStr = 'deliveryorderno='+selecteddo+'&fcba='+fcba+'&doline='+selectedlineno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getPurchaseDOlineList.asp', 'POST', PostStr, showDoLine);
}

// The function for handling the response from the server
var showDoLine = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("dolinelist").innerHTML = "";
	document.getElementById("dolinelist").innerHTML = response;

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
var getMessage = function () {
	// build up the post string when passing variables to the server side page
	var fcba = document.getElementById("fcba").value;
	var selecteddo = document.getElementById("selecteddo").value;
	var selectedlineno = document.getElementById("selectedlineno").value;
	var PostStr = 'deliveryorderno='+selecteddo+'&fcba='+fcba+'&doline='+selectedlineno;
	// alert(PostStr);
	// use the generic function to make the request
	if (!$('#extracarry').is(":checked")){
			doAJAXCall('ajax/GetPurchaseDeliveryOrder.asp', 'POST', PostStr, showMessageResponse);
		} else {
			doAJAXCall('ajax/GetPurchaseDeliveryOrderExtraCarry.asp', 'POST', PostStr, showMessageResponse);
		} 
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s);
    // update the Div to show the result from the server
    document.getElementById("quantity_do").value = s.quantity_do;
	document.getElementById("salescontractno").value = s.salescontractno;
	document.getElementById("productcode").value = s.productcode;
	document.getElementById("customercode_").value = s.customercode;
	document.getElementById("transportercode_").value = s.transportercode;
	
	// document.getElementById("address").innerHTML = document.getElementById("supplieraddress").value;
	 //alert(s.salescontractno);
    
} 

var getTicket = function () {
	// build up the post string when passing variables to the server side page

	var chitno = document.getElementById("chitno").value;
	var PostStr = 'chitno='+chitno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetTicket.asp', 'POST', PostStr, ShowTicketResponse);
}

// The function for handling the response from the server
var ShowTicketResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.vehicleno);
    // alert(s.Netto);
    
    // update the Div to show the result from the server
	// document.getElementById("postingdate").value = s.postingdate;
	// document.getElementById("transporter").value = s.transporter;
	document.getElementById("vehicleno").value = s.vehicleno;
	document.getElementById("Netto").value = s.Netto;

};

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
Dim salescontractno, parcode, totalstand, customercode, customercode_,deliveryaddress, deliverytype, downpayment
Dim FIELDCODE, productcode,  unitofmeasure, wbt_chitno
Dim QUANTITY, curr, salesprices, amount, vat, totalamountvat, amountvat
Dim REMARKS, paymentcondition, deliverycondition
Dim creationdate, startdate, enddate, deliverydate, deliveryorderdate
Dim cmdView
Dim txtDocumentNo ,txtlineno
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

purchaseshipmentno =request("purchaseshipmentno")
salescontractno =request("salescontractno") 
deliverydate =request("deliverydate")
creationdate =request("creationdate")
deliveryorderno =request("deliveryorderno")
deliveryorderno2 =request("deliveryorderno2")
deliveryorderdate =request("deliveryorderdate")
quantity_do=request("quantity_do") 
shipmenttype=request("shipmenttype")
destinationcountry=request("destinationcountry")
address = Request("address")
remarks = Request("remarks")
txtlineno = request("txtlineno")
productcode =request("productcode")
transportercode_=request("transportercode")
customercode_ = request("customercode_")
chitno =request("chitno")
wbt_chitno =request("wbt_chitno")
customer_chitno =request("customer_chitno")
postingdate =request("postingdate")
transporter =request("transporter")
vehicleno =request("vehicleno")
quantity =request("quantity")
acceptquantity =request("acceptquantity")
acceptquantitybruto =request("acceptquantitybruto")
transporter =request("transporter")
transportercode =request("transportercode")
deliveredqty =request("deliveredqty")
shipmentfrom =request("shipmentfrom")
selectedfcba =request("selectedfcba") 
fcba =request("fcba") 
selecteddo =request("selecteddo") 
extracarry =request("extracarry")
doline =request("doline")
selectedlineno =request("selectedlineno")


Function CheckTicket(initchit, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total from deliverynote" & _
					" where wbt_chitno = '"& initchit &"'" & _
					" and fcba = '"& initialFcba & "'" 
	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable =  Cdbl(to_nol(rsAvailable("total")))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckTicket = initialAvailable
	
End Function

Function CheckQtyDo(initdeliveryorderno, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total from deliveryorder" & _
					" where deliveryorderno = '"& initdeliveryorderno &"'" & _
					" and fcba = '"& initialFcba & "'" 
	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable =  Cdbl(to_nol(rsAvailable("total")))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckQtyDo = initialAvailable
	
End Function



If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, creationdate)
		If Cek_IsDailyClosing = "FALSE" then
			' strFieldDetail = " SELECT COUNT(1) total FROM purchaseshipmentdtl " & _
	  '                        " where fcBA = '" & fcba & "'" & _
			' 				 " and purchaseshipmentno = '" & purchaseshipmentno & "'"							
			' ' Response.Write strFieldDetail
			' Set rsQty = Server.CreateObject("ADODB.Recordset")
			' rsQty.Open strFieldDetail, connstr							
			' if not rsQty.Eof then
			' 	has_detail = cdbl (to_nol(rsQty("total")))
			' End If
			
			' 	if has_detail = 0 then
					q = " delete from purchaseshipmenthdr " & _
						" where fcBA = '" & fcba & "'" & _
						" and purchaseshipmentno = '" & purchaseshipmentno & "'"
						' Response.Write q
						connStr.execute(q)

					q = " delete from purchaseshipmentdtl " & _
						" where fcBA = '" & fcba & "'" & _
						" and purchaseshipmentno = '"& purchaseshipmentno & "'" 
						' Response.Write q
						connStr.execute(q)

					q = " delete from purchaseshipmentdtl_extracarry " & _
						" where fcBA = '" & fcba & "'" & _
						" and purchaseshipmentno = '"& purchaseshipmentno & "'" 
						' Response.Write q
						connStr.execute(q)
				' Else
				' 	err.number = 1001
				' 	err.description = " Sales Shipment """ & purchaseshipmentno & """ has detail, please delete all detail first!"
				' End If
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & Creationdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then
			
			 q = " delete from purchaseshipmentdtl " & _
				" where fcBA = '" & fcba & "'" & _
			 	" and purchaseshipmentno = '"& purchaseshipmentno & "'" & _
			 	" and wbt_chitno = '"& wbt_chitno & "'" 
			 	 Response.Write q
			 	connStr.execute(q)

			
			cmdView	= "True"

	elseif Request("txtJenis")="New" then
			Dim rsSupplyingDeclarationHdr
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select purchasecontractno from purchaseshipmenthdr where purchaseshipmentno = '" & purchaseshipmentno & "' and fcba ='" & fcba & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then 
				q = " INSERT into purchaseshipmenthdr" & _
					" (purchaseshipmentno, creationdate, purchasecontractno, deliverydate, deliveryorderno, deliveryorderdate, quantity_do, shipmenttype, destinationcountry, address, remarks, shipmentfrom, extracarry, doline, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" '" & purchaseshipmentno & "'" & _
					" , to_date('" & creationdate & "','mm/dd/yyyy')" & _ 
					" , '" & salescontractno & "'" & _
					" , to_date('" & deliverydate & "','mm/dd/yyyy')" & _ 
					" , '" & deliveryorderno & "'" & _
					" , to_date('" & deliveryorderdate & "','mm/dd/yyyy')" & _ 
					" , '" & quantity_do & "'" & _
					" , '" & shipmenttype & "'" & _
					" , '" & destinationcountry & "'" & _
					" , '" & address & "'" & _
					" , '" & remarks & "'" & _
					" , '" & shipmentfrom & "'" & _
					" , '" & extracarry & "'" & _
					" , '" & doline & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & fcba & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"								
				' Response.Write q
				connStr.execute(q) 
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing
			
			
			if cstr(""&shipmentfrom) = "WEIGHBRIDGE" then 	
      			
				If cdbl(to_nol(extracarry)) > 0 Then  
					For n=1 to (cdbl(request("no"))-1) 
						if request("acceptquantity_"&TRIM(n))<>"" then
							strFieldDetail = " select SUM(quantityaftertolerance) quantityaftertolerance " & _
											 " from purchaseagreementdet  " & _ 
											 " where purchasecontractno ='"& salescontractno &"'" & _
										 	 " and fcba ='" & fcba & "'"
											' Response.Write strFieldDetail
											Set rsQty = Server.CreateObject("ADODB.Recordset")
											rsQty.Open strFieldDetail, connstr							
							if not rsQty.Eof then
							do while not rsQty.eof
									quantityaftertolerance = cdbl (to_nol(rsQty("quantityaftertolerance")))
									rsQty.MoveNext
								loop
							End If
							
							strFieldDetail = " select SUM(ssd.acceptquantity) acceptquantity " & _
											 " from purchaseshipmentdtl ssd " & _ 
											 " inner join purchaseshipmenthdr ssh" & _
											 " on (ssd.fcba = ssh.fcba and ssd.purchaseshipmentno = ssh.purchaseshipmentno)" & _
											 " where ssh.purchasecontractno ='"& salescontractno &"'" & _
										 	 " and ssh.fcba ='" & fcba & "'"
											' Response.Write strFieldDetail
											Set rsQty = Server.CreateObject("ADODB.Recordset")
											rsQty.Open strFieldDetail, connstr							
							if not rsQty.Eof then
							do while not rsQty.eof
									v_acceptquantity = cdbl (to_nol(rsQty("acceptquantity")))
									rsQty.MoveNext
								loop
							End If
							
							total = v_acceptquantity + cdbl(to_nol(request("acceptquantity_"&TRIM(n))))
							
							
							If (cdbl(to_nol(total)) > cdbl(to_nol(quantityaftertolerance))) AND cstr(""&productcode) <> "TBS" Then
								Err.Number = 1001
								Err.Description = "Accept Quantity tidak boleh lebih dari Quantity After Tolerance Purchase Contract"
							Else
								' If  cdbl(to_nol(request("acceptquantity_original_"&TRIM(n)))) = 0 Then deliveredqty = 0 Else deliveredqty = round((  cdbl(to_nol(request("acceptquantity_"&TRIM(n)))) / cdbl(to_nol(request("acceptquantity_original_"&TRIM(n)))) ) * cdbl(to_nol(request("deliveredqty_original_"&TRIM(n)))),0)
								' response.write deliveredqty
								q = " insert into purchaseshipmentdtl" & _
									" (wbt_chitno," & _
									"	customer_chitno, " & _
									"   purchaseshipmentno, " & _ 
									"	postingdate, "& _ 
									"	transporter, " & _
									" 	vehicleno, " & _
									"	quantity, " & _ 
									"	acceptquantitynetto_ori, " & _ 
									"	deliveredqtynetto_ori, " & _ 
									"	acceptquantity, " & _ 
									"	acceptquantitybruto, " & _ 
									"	validationextracarry, " & _ 
									"	fcentry, " & _ 
									"	fcedit, " & _ 
									" 	fcip," & _ 
									"	fcba, " & _ 
									" 	lastupdate," & _ 
									"	lasttime) values (" & _
									"  '" & request("wbt_chitno_"&TRIM(n)) & "'" & _
									" , '" & request("customer_chitno_"&TRIM(n)) & "'" & _
									" , '" & purchaseshipmentno & "'" & _ 
									" ,to_date('" & request("postingdate_"&TRIM(n)) & "','mm/dd/yyyy')" & _
									" , '" & transportercode_ & "'" & _
									" , '" & request("vehicleno_"&TRIM(n)) & "'" & _
									" , replace('" & request("deliveredqty_original_"&TRIM(n)) & "',',','')" & _
									" , replace('" & request("acceptquantity_original_"&TRIM(n)) & "',',','')" & _
									" , replace('" & request("deliveredqty_original_"&TRIM(n)) & "',',','')" & _
									" , replace('" & request("acceptquantity_"&TRIM(n)) & "',',','')" & _
									" , replace('" & request("acceptquantitybruto_"&TRIM(n)) & "',',','')" & _ 
									" , '" & request("cb_"&TRIM(n)) & "'" & _
									" , '" & Session("CodeName") & "'" & _
									" , '" & Session("CodeName") & "'" & _
									" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									" , '" & fcba & "'" & _
									" , sysdate" & _
									" , to_char(sysdate, 'hh24:mi')" & _
									" )"
									 Response.Write q
									connStr.execute(q) 

								q1 = 	" INSERT INTO purchaseshipmentdtl_extracarry SELECT purchaseshipmentno," & _
										"        deliveryorderno," & _
										"        wbt_chitno," & _
										"        postingdate," & _
										"        transporter," & _
										"        vehicleno," & _
										"        quantity," & _
										"        acceptquantity," & _
										"        acceptquantitybruto," & _
										"        transporter_invoice," & _
										"        fcentry," & _
										"        fcedit," & _
										"        fcip," & _
										"        fcba," & _
										"        lastupdate," & _
										"        lasttime," & _
										"        customer_chitno" & _
										"   FROM (SELECT ssd.purchaseshipmentno," & _
										"                ssh.deliveryorderno," & _
										"                ssd.wbt_chitno," & _
										"                ssd.postingdate," & _
										"                ssd.transporter," & _
										"                ssd.vehicleno," & _
										"                ssd.quantity," & _
										"                ssd.acceptquantity," & _
										"                ssd.acceptquantitybruto," & _
										"                ssd.transporter_invoice," & _
										"                ssd.fcentry," & _
										"                ssd.fcedit," & _
										"                ssd.fcip," & _
										"                ssd.fcba," & _
										"                ssd.lastupdate," & _
										"                ssd.lasttime," & _
										"                ssd.customer_chitno" & _
										"           FROM purchaseshipmentdtl ssd" & _
										"                INNER JOIN purchaseshipmenthdr ssh" & _
										"                   ON (    ssh.fcba = ssd.fcba" & _
										"                       AND ssh.purchaseshipmentno = ssd.purchaseshipmentno)" & _
										"          WHERE     ssh.purchaseshipmentno = '"& purchaseshipmentno &"'" & _
										"                AND ssh.fcba = '"& fcba &"'" & _
										"                AND ssh.extracarry = 1)"  
								' Response.Write q1
								connStr.execute(q1)
							End if	
						End If 
					Next

				Else 
Response.Write "Masuk Not extra WEIGHBRIDGE"
			
					For n=1 to (cdbl(request("no"))-1) 
						if request("acceptquantity_"&TRIM(n))<>"" then
							strFieldDetail = " select SUM(quantityaftertolerance) quantityaftertolerance " & _
											 " from purchaseagreementdet  " & _ 
											 " where purchasecontractno ='"& salescontractno &"'" & _
										 	 " and fcba ='" & fcba & "'"
											' Response.Write strFieldDetail
											Set rsQty = Server.CreateObject("ADODB.Recordset")
											rsQty.Open strFieldDetail, connstr							
							if not rsQty.Eof then
							do while not rsQty.eof
									quantityaftertolerance = cdbl (to_nol(rsQty("quantityaftertolerance")))
									rsQty.MoveNext
								loop
							End If
							
							strFieldDetail = " select SUM(ssd.acceptquantity) acceptquantity " & _
											 " from purchaseshipmentdtl ssd " & _ 
											 " inner join purchaseshipmenthdr ssh" & _
											 " on (ssd.fcba = ssh.fcba and ssd.purchaseshipmentno = ssh.purchaseshipmentno)" & _
											 " where ssh.purchasecontractno ='"& salescontractno &"'" & _
										 	 " and ssh.fcba ='" & fcba & "'"
											' Response.Write strFieldDetail
											Set rsQty = Server.CreateObject("ADODB.Recordset")
											rsQty.Open strFieldDetail, connstr							
							if not rsQty.Eof then
							do while not rsQty.eof
									v_acceptquantity = cdbl (to_nol(rsQty("acceptquantity")))
									rsQty.MoveNext
								loop
							End If
							
							total = v_acceptquantity + cdbl(to_nol(request("acceptquantity_"&TRIM(n))))
							
									response.write total & "|" & quantityaftertolerance & "|" & productcode
					
							
							If (cdbl(to_nol(total)) > cdbl(to_nol(quantityaftertolerance))) AND  cstr(""&productcode) <> "TBS" Then
								Err.Number = 1001
								Err.Description = "Accept Quantity tidak boleh lebih dari Quantity After Tolerance Purchase Contract"
							Else
								' If cdbl(to_nol(request("acceptquantity_original_"&TRIM(n)))) = 0 Then deliveredqty = 0 Else deliveredqty = round( (  cdbl(to_nol(request("acceptquantity_"&TRIM(n)))) / cdbl(to_nol(request("acceptquantity_original_"&TRIM(n)))) ) * cdbl(to_nol(request("deliveredqty_original_"&TRIM(n)))),0)
								' response.write deliveredqty
								q = " insert into purchaseshipmentdtl" & _
									" (wbt_chitno," & _
									"	customer_chitno, " & _
									"   purchaseshipmentno, " & _ 
									"	postingdate, "& _ 
									"	transporter, " & _
									" 	vehicleno, " & _
									"	quantity, " & _ 
									"	acceptquantitynetto_ori, " & _ 
									"	deliveredqtynetto_ori, " & _ 
									"	acceptquantity, " & _ 
									"	acceptquantitybruto, " & _ 
									"	validationextracarry, " & _ 
									"	fcentry, " & _ 
									"	fcedit, " & _ 
									" 	fcip," & _ 
									"	fcba, " & _ 
									" 	lastupdate," & _ 
									"	lasttime) values (" & _
									"  '" & request("wbt_chitno_"&TRIM(n)) & "'" & _
									" , '" & request("customer_chitno_"&TRIM(n)) & "'" & _
									" , '" & purchaseshipmentno & "'" & _ 
									" ,to_date('" & request("postingdate_"&TRIM(n)) & "','mm/dd/yyyy')" & _
									" , '" & transportercode_ & "'" & _
									" , '" & request("vehicleno_"&TRIM(n)) & "'" & _
									" , replace('" & request("deliveredqty_original_"&TRIM(n)) & "',',','')" & _
									" , replace('" & request("acceptquantity_original_"&TRIM(n)) & "',',','')" & _
									" , replace('" & request("deliveredqty_original_"&TRIM(n)) & "',',','')" & _
									" , replace('" & request("acceptquantity_"&TRIM(n)) & "',',','')" & _
									" , replace('" & request("acceptquantitybruto_"&TRIM(n)) & "',',','')" & _ 
									" , '" & request("cb_"&TRIM(n)) & "'" & _
									" , '" & Session("CodeName") & "'" & _
									" , '" & Session("CodeName") & "'" & _
									" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									" , '" & fcba & "'" & _
									" , sysdate" & _
									" , to_char(sysdate, 'hh24:mi')" & _
									" )"
									' Response.Write q
									connStr.execute(q) 
							End if	
						End If 
					Next
				ENd If
		Else
			If Len(Trim(wbt_chitno)) = 0 Then
				Err.Number = 1001
				Err.Description = "No Tikcet Harus diisi, tidak boleh kosong"
			Else
			strFieldDetail =	" select count(1) total" & _
				                " from purchaseshipmentdtl " & _
								" where wbt_chitno ='" & wbt_chitno & "'" & _
								" and fcba ='" & fcba & "'"
							' Response.Write strFieldDetail
							Set rsQty = Server.CreateObject("ADODB.Recordset")
							rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
			do while not rsQty.eof
					ticket = cdbl (to_nol(rsQty("total")))
					rsQty.MoveNext
				loop
			End If

			If ticket = 1 Then
				Err.Number = 1001
				Err.Description = "Ticket """ & wbt_chitno & """ sudah ada, mohon masukan Ticket yang lain."
			Else

			If (cdbl(to_nol(acceptquantity)) > cdbl(to_nol(quantity))) AND cstr(""&productcode) <> "TBS" Then
				Err.Number = 1001
				Err.Description = "Accept Quantity tidak boleh lebih dari Quantity Delivery Order"
			Else
			If Len(Trim(transporter)) = 0 Then
				Err.Number = 1001
				Err.Description = "Transporter tidak boleh kosong"
			Else 
			If Len(Trim(acceptquantity)) = 0 Then
				Err.Number = 1001
				Err.Description = "Accept Quantity tidak boleh kosong"
			Else  
						q = " insert into purchaseshipmentdtl" & _
						" (wbt_chitno, customer_chitno, purchaseshipmentno, postingdate, transporter, vehicleno, quantity, acceptquantity, acceptquantitybruto, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
						" '" & wbt_chitno & "'" & _
						" , '" & customer_chitno & "'" & _
						" , '" & purchaseshipmentno & "'" & _
						" , to_date('" & postingdate & "','mm/dd/yyyy')" & _
						" , '" & transporter & "'" & _
						" , '" & vehicleno & "'" & _
						" , '" & quantity & "'" & _
						" , '" & acceptquantity & "'" & _
						" , '" & acceptquantitybruto & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						" , '" & fcba & "'" & _
						" , sysdate" & _
						" , to_char(sysdate, 'hh24:mi')" & _
						" )"
						' Response.Write q
						connStr.execute(q)
					End If
				End If
			End If
		End If
	End If
End If
cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			If Len(Trim(wbt_chitno)) = 0 Then
				Err.Number = 1001
				Err.Description = "No Tikcet Harus diisi, tidak boleh kosong"
			Else
			' strFieldDetail =	" select count(1) total" & _
			' 	                " from purchaseshipmentdtl " & _
			' 					" where wbt_chitno ='" & wbt_chitno & "'" & _
			' 					" and fcba ='" & Session("fcba") & "'"
			' 				' Response.Write strFieldDetail
			' 				Set rsQty = Server.CreateObject("ADODB.Recordset")
			' 				rsQty.Open strFieldDetail, connstr							
			' if not rsQty.Eof then
			' do while not rsQty.eof
			' 		ticket = cdbl (to_nol(rsQty("total")))
			' 		rsQty.MoveNext
			' 	loop
			' End If

			' If ticket = 1 Then
			' 	Err.Number = 1001
			' 	Err.Description = "Ticket """ & wbt_chitno & """ sudah ada, mohon masukan Ticket yang lain."
			' Else

			' strFieldDetail = " select wb.weightin, wb.weightout, gr.weight, (wb.weightin-wb.weightout-gr.weight) as Netto " & _
			' 				 " from weighbridgeticket wb" & _
			' 				 " left join gradingrecord gr ON (wb.chitno = gr.wbt_chitno and wb.fcba = gr.fcba)" & _
			' 				 " where wb.chitno='029105'" & _
			' 			 	 " and wb.fcba ='" & Session("fcBA") & "'"
			' 				' Response.Write strFieldDetail
			' 				Set rsQty = Server.CreateObject("ADODB.Recordset")
			' 				rsQty.Open strFieldDetail, connstr							
			' if not rsQty.Eof then
			' do while not rsQty.eof
			' 		netto = cdbl (to_nol(rsQty("Netto")))
			' 		rsQty.MoveNext
			' 	loop
			' End If
			' Response.Write acceptquantity
			' Response.Write netto
			If (cdbl(to_nol(acceptquantity)) > cdbl(to_nol(quantity))) AND cstr(""&productcode) <> "TBS" Then
				Err.Number = 1001
				Err.Description = "Accept Quantity tidak boleh lebih dari Quantity Delivery Order"
			Else
					q =	" Update purchaseshipmentdtl" & _
						" set wbt_chitno = '"& wbt_chitno &"'" & _
						"     , customer_chitno = '" & customer_chitno & "'" & _
						"     , vehicleno = '" & vehicleno & "'" & _
						"     , acceptquantity = '" & acceptquantity & "'" & _
						"     , acceptquantitybruto = '" & acceptquantitybruto & "'" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where purchaseshipmentno = '" & purchaseshipmentno & "'" & _
						"       and fcBA = '" & fcba & "'"
		 			' response.write q
					connStr.execute(q)
					End If
				End If		
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
If Len(Trim(creationdate)) = 0 Then	creationdate = textDate(date)
If Len(Trim(deliverydate)) = 0 Then	deliverydate = textDate(date)
If Len(Trim(deliveryorderdate)) = 0 Then deliveryorderdate = textDate(date)
If Len(Trim(postingdate)) = 0 Then postingdate = textDate(date)



Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select *" & _
							" from purchaseshipmenthdr hdr" & _
							" where hdr.purchaseshipmentno = '" & purchaseshipmentno & "'" 

'response.write 	strSuppyingDeclaration					
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	creationdate = rsSupplyingDeclaration("creationdate")
	salescontractno = rsSupplyingDeclaration("purchasecontractno")
	salescontractno2 = rsSupplyingDeclaration("purchasecontractno")
	deliverydate = rsSupplyingDeclaration("deliverydate")
	deliveryorderno = rsSupplyingDeclaration("deliveryorderno")
	deliveryorderdate = rsSupplyingDeclaration("deliveryorderdate")
	quantity_do = rsSupplyingDeclaration("quantity_do")
	quantity_do2 = rsSupplyingDeclaration("quantity_do")
	shipmenttype = rsSupplyingDeclaration("shipmenttype")
	destinationcountry = rsSupplyingDeclaration("destinationcountry")
	address = rsSupplyingDeclaration("address")
	remarks = rsSupplyingDeclaration("remarks")
	doline = rsSupplyingDeclaration("doline")
	fcba = rsSupplyingDeclaration("fcba")
	fcba2 = rsSupplyingDeclaration("fcba")
	selectedfcba = rsSupplyingDeclaration("fcba")
	selectedfcba2 = rsSupplyingDeclaration("fcba")
	selecteddo = rsSupplyingDeclaration("deliveryorderno")
	selecteddo2 = rsSupplyingDeclaration("deliveryorderno")
	selectedlineno = rsSupplyingDeclaration("doline")
	HeaderData = "Y"
	shipmentfrom = rsSupplyingDeclaration("shipmentfrom")
End if
' response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

strSuppyingDeclarations	=	" select *" & _
							" from purchaseshipmentdtl " & _
							" where purchaseshipmentno = '" & purchaseshipmentno & "'" 
						
set rsSupplyingDeclarations = server.CreateObject("Adodb.recordset")
rsSupplyingDeclarations.Open strSuppyingDeclarations, connstr
if Not rsSupplyingDeclarations.Eof then
	postingdate = rsSupplyingDeclarations("postingdate")
	purchaseshipmentno = rsSupplyingDeclarations("purchaseshipmentno")
	wbt_chitno = rsSupplyingDeclarations("wbt_chitno")
	transporter = rsSupplyingDeclarations("transporter")
	vehicleno = rsSupplyingDeclarations("vehicleno")
	quantity = rsSupplyingDeclarations("quantity")
	acceptquantity = rsSupplyingDeclarations("acceptquantity")
	customer_chitno = rsSupplyingDeclarations("customer_chitno")
	acceptquantitybruto = rsSupplyingDeclarations("acceptquantitybruto")
	SudahClose = "Y"
End if
' response.Write strSuppyingDeclarations
rsSupplyingDeclarations.Close
Set rsSupplyingDeclarations = Nothing

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
' response.write oSQLQuery			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if

'response.Write strNurseryCulling
oRecordSet.Close : Set oRecordSet = Nothing
 
%>

 <!-------------ISI Supplying Declaration Form!------------->
 
	<form name="frmEntry" method="post" action="purchaseshipment.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Shipment No
				<td align="left">
				<input type="text" maxlength="50" name="purchaseshipmentno" style="background:#FFFFCC;" size="30" value="<%=purchaseshipmentno%>" autocomplete="Off">
				<font color="#FF0000">*</font>
				</td>
			</td>
			<td align="left" width="20%">Shipment Type 
				<td align="left">
				<select name="shipmenttype" title="Shipment Type CODE Character a s/d z or 1 s/d 9">
					<option value=""></option> 
						<% 	Dim rsComboT
							q = "SELECT * FROM PARAMETER WHERE PARCODE = 'SHIPMENT_TYPE' and fcba = '" & Session("fcba") & "'"
							Set rsComboT = Server.CreateObject("ADODB.Recordset")
							rsComboT.CursorType = 3
							rsComboT.CursorLocation = 3
							rsComboT.Open q, connstr
							if not rsComboT.eof then%>
							<%
							Do While Not rsComboT.eof
							%>
								<option value="<%=rsComboT("FCNAME")%>" 
									<%if shipmenttype=cstr(""&rsComboT("FCNAME")) then %> selected 
									<%End If%>><%=rsComboT("FCNAME")%>
								</option> 
							<%rsComboT.movenext
							loop
							End if
							rsComboT.Close
							set rsComboT = Nothing
							q=""
						%>
				</select> 
				<input type="text" maxlength="5" name="productcode" id="productcode" size="5" style="background:#eee;" value="<%=productcode%>" autocomplete="Off" readonly>
                <input type="text" maxlength="5" name="transportercode_" id="transportercode_" size="5" style="background:#eee;" value="<%=transportercode_%>" autocomplete="Off" readonly>			
				</td>
				
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Creation Date 
			<td align="left">
				<input class="datepicker" type="text" size="15" name="creationdate" value="<%=creationdate%>">
				Extra Carry
				<input type="checkbox" id="extracarry" name="extracarry" value="<%=extracarry%>" <%if cstr(""&extracarry) = "1" then%> checked <%end if%> onchange="CheckCheckboxes1(this);"> 
			</td>
			<td align="left" width="20%">Destination Country
				<td align="left">
				<input type="text" maxlength="15" name="destinationcountry" size="30" value="<%=destinationcountry%>" autocomplete="Off">
				</td>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Business Unit</td>
				<td align="left">
					<select name="fcba" id="fcba" title="Business Unit Character a s/d z or 1 s/d 9" onchange="javascript:getDOList();" class="chosen-select" style="width: 70%">
							<option value=""></option> 
								<% 	q = " SELECT fcba, fcname FROM user_fcba_tab uft" & _
										" INNER JOIN businessunit bu ON bu.fccode = uft.fcba " & _
										" WHERE uft.fcusercode = '" & Session("CodeName") & "' /*AND fctype = 'M'*/"
									' response.write q
									Set rsComboB = Server.CreateObject("ADODB.Recordset")
									rsComboB.CursorType = 3
									rsComboB.CursorLocation = 3
									rsComboB.Open q, connstr
									if not rsComboB.eof then%>
									<%
									Do While Not rsComboB.eof
									%>
										<option value="<%=rsComboB("fcba")%>" 
											<%if fcba =cstr(""&rsComboB("fcba")) then %> selected 
											<%End If%>><%=rsComboB("fcba") & " : " & rsComboB("fcname")%> 
										</option> 
									<%rsComboB.movenext
									loop
									End if
									rsComboB.Close
									set rsComboB = Nothing
									q=""
								%>
						</select>
				</td> 
					<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>"> 
			<td align="left" width="20%">Address
				<td align="left">
				<textarea name="address"><%=address%></textarea>
				</td>
			</td>
			
		</tr>
		<tr>
			<%if HeaderData = "Y" then %>
			<td align="left" width="20%">Delivery Order No
				<td align="left">
					<input type="text" id="selecteddo" name="selecteddo" value="<%=selecteddo%>" readonly style="background-color: #eee;">
					<font color="#FF0000">*</font>
				</td>
			<%else%>
				<td align="left" width="20%">Delivery Order No
				<td align="left">
					<div id= "dolist"></div>
					<font color="#FF0000">*</font>
				</td> 
				<input type="Hidden" id="selecteddo" name="selecteddo" value="<%=selecteddo%>">
			<%end if%>
			</td>
			<td align="left" width="20%">Remarks
				<td align="left">
				<textarea name="remarks"><%=remarks%></textarea>
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Delivery Order Line No 
			<td align="left">
				<div id="dolinelist"></div>
			</td> 
			<input type="Hidden" id="selectedlineno" name="selectedlineno" value="<%=selectedlineno%>">
			</td>
			<td align="left" width="20%">Shipment From
				<td align="left">
					<select name="shipmentfrom" title="Shipment From Character a s/d z or 1 s/d 9" >
						<option value="WEIGHBRIDGE" <% If cstr(""&shipmentfrom) = "WEIGHBRIDGE" Then %> Selected <%End if%>>WEIGHBRIDGE</option> 
						<option value="NON_WEIGHBRIDGE" <% If cstr(""&shipmentfrom) = "NON_WEIGHBRIDGE" Then %> Selected <%End if%>>NON WEIGHBRIDGE</option> 
					</select>
				</td>
			</td>
		</tr>
    	<tr>
			<td align="left" width="20%">Delivery Order Quantity 
			<td align="left">
				<input type="text" maxlength="15" name="quantity_do" id="quantity_do" size="30" value="<%=quantity_do%>" autocomplete="Off" style="background:#eee;" size="30" readonly="true">
			</td> 
			</td>
			<td align="left" width="20%">
				Purchase Contract No
			</td>
			<td align="left">
			    <input type="text" maxlength="5" name="customercode_" id="customercode_" size="5" style="background:#eee;" value="<%=customercode_%>" autocomplete="Off" readonly>
				<input type="text" maxlength="45" name="salescontractno" id="salescontractno" size="25" style="background:#eee;" value="<%=salescontractno%>" autocomplete="Off" readonly>
			</td>  
		</tr>

		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				
					<a href="purchaseshipment.asp?purchaseshipmentno=<%=purchaseshipmentno%>&fcba=<%=fcba%>&salescontractno=<%=salescontractno%>&deliveryorderno=<%=deliveryorderno%>&txtJenis=HeaderDelete">
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
<% if cmdView = "True" and len(trim(purchaseshipmentno)) > 0  Then
	if cstr(""&shipmentfrom) = "WEIGHBRIDGE" then 
		if SudahClose = "Y" Then
			q = " SELECT sdt.wbt_chitno AS wbt_chitno," & _
				"        sdt.postingdate," & _
				"        sdt.transporter," & _
				"        tr.fcname AS transportername," & _
				"        sdt.quantity," & _
				"        sdt.deliveredqtynetto_ori," & _
				"        sdt.vehicleno," & _
				"        sdt.customer_chitno," & _
				"        sdt.acceptquantity," & _
				"        sdt.acceptquantitynetto_ori," & _
				"        sdt.acceptquantitybruto" & _
				"   FROM   purchaseshipmentdtl sdt " & _
				"        INNER JOIN purchaseshipmenthdr shd" & _
				"           ON (    sdt.purchaseshipmentno = shd.purchaseshipmentno" & _
				"               AND sdt.fcba = shd.fcba" & _
				"               AND sdt.wbt_chitno = sdt.wbt_chitno)" & _ 
				"        INNER JOIN transporter tr" & _
				"           ON (tr.fccode = sdt.transporter AND tr.fcba = sdt.fcba)" & _
				" WHERE     shd.fcba = '"& fcba &"'" & _
				"       AND shd.purchasecontractno = '"& salescontractno &"'" & _
				"		AND shd.deliveryorderno ='"& deliveryorderno &"'" & _
				" 		AND shd.purchaseshipmentno = '"& purchaseshipmentno &"'"
	   	Else	 
	   		If cdbl(to_nol(extracarry)) > 0 Then 
				q = " SELECT sse.purchaseshipmentno," & _
					"        sse.wbt_chitno," & _
					"        sse.postingdate," & _
					"        sse.transporter transportercode," & _
					"        tr.fcname  AS transportername," & _
					"        sse.vehicleno," & _
					"        sse.quantity," & _
					"		 sse.deliveredqtynetto_ori," & _
					"        sse.acceptquantity," & _
					"		 sse.acceptquantitynetto_ori," & _
					"        sse.acceptquantitybruto," & _
					"        sse.customer_chitno" & _
					"   FROM purchaseshipmentdtl sse" & _
					"        INNER JOIN purchaseshipmenthdr ssh" & _
					"           ON (    ssh.purchaseshipmentno = sse.purchaseshipmentno" & _
					"               AND ssh.fcba = sse.fcba)" & _
					"        INNER JOIN transporter tr" & _
					"           ON (tr.fccode = sse.transporter AND tr.fcba = sse.fcba)" & _
					"  WHERE     sse.fcba = '"& fcba &"'" & _
					"        AND sse.validationextracarry = 1" & _
					"        AND ssh.deliveryorderno = '"& deliveryorderno &"'" & _
					"        AND ssh.purchasecontractno = '"& salescontractno &"'"  
			Else
				IF cstr(""&productcode) <> "TBS" then
				q = " SELECT wb.chitno AS wbt_chitno," & _
					"        wb.dateout AS postingdate," & _
					"        dn.transportco AS transportercode," & _
					"        tr.fcname AS transportername," & _
					"        (wb.weightout - wb.weightin) AS deliveredqty," & _
					"        wb.vehicleno" & _
					"   FROM weighbridgeticket wb" & _
					"        INNER JOIN deliverynote dn" & _
					"           ON (wb.fcba = dn.fcba AND wb.chitno = dn.wbt_chitno)" & _
					" 		 INNER JOIN deliveryorder do ON (do.fcba = dn.fcba AND do.deliveryorderno = dn.deliverynotecode AND do.transporter = dn.transportco)" & _
					"		 INNEr JOIN transporter tr ON (dn.transportco = tr.fccode AND dn.fcba = tr.fcba)" & _
					"  WHERE     wb.fcba = '"& fcba &"'" & _
					"        AND dn.contractcode = '"& salescontractno &"'" & _
					"        AND dn.deliverynotecode = '"& deliveryorderno &"'" & _
					"        AND do.lineno = '"& doline &"'" & _
					"        AND wb.chitno IN" & _
					"               (SELECT wbt_chitno" & _
					"                  FROM (SELECT FCBA, wbt_chitno" & _
					"                          FROM iplasprod.deliverynote" & _
					"                         WHERE     fcba = '"& fcba &"'" & _
					"                               AND deliverynotecode = '"& deliveryorderno &"'" & _
					"                        MINUS" & _
					"                        SELECT sdt.fcba, sdt.wbt_chitno" & _
					"                          FROM purchaseshipmentdtl sdt" & _
					"                               INNER JOIN purchaseshipmenthdr ssh" & _
					"                                  ON (    ssh.fcba = sdt.fcba" & _
					"                                      AND ssh.purchaseshipmentno =" & _
					"                                             sdt.purchaseshipmentno)" & _
					"                         WHERE     sdt.fcba = '"& fcba &"'" & _
					"                               AND ssh.deliveryorderno = '"& deliveryorderno &"'))"
				Else
				    q =  " SELECT wb.docketid AS wbt_chitno," & _
					"        b.docketdate AS postingdate," & _
					"        '' AS transportercode," & _
					"        '' AS transportername," & _
					"        (wb.weightout - wb.weightin) AS deliveredqty," & _
					"        b.noreg vehicleno" & _
					" FROM lhpffbwbtrn wb " & _
					" left JOIN lhptikethdr b ON wb.docketid=b.docketid AND Trim(wb.fcba)=Trim(b.fcba) " & _
					" left JOIN lhppksmst c ON  Trim(b.idpksmst)=Trim(c.idpks) AND Trim(wb.fcba)= Trim(c.fcba) " & _
					" left JOIN customer d ON Trim(c.eplascode)=Trim(d.customercode) AND Trim(c.fcba)=Trim(d.fcba) " & _
					"  WHERE     wb.fcba = '"& fcba &"'" & _
					"        AND c.eplascode = '"& customercode_ &"'" & _
					"        AND b.docketdate = to_date('" & creationdate & "','mm/dd/yyyy')" 
				
				END IF
			End If    
	   End if
       'response.Write q	
	   Set rsGRN = Server.CreateObject("ADODB.Recordset")
	   rsGRN.Open q, connstr
	   %>
	   <form action="purchaseshipment.asp" method="post" name="eForm">
	   <table class="data" border="1" width="100%">
	   <%if SudahClose<>"Y" Then%>
		   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		   <td align="center">No</td>
		   <td align="center">Chitno</td>
		   <td align="center">Date</td>
		   <td align="center">Transporter</td> 
		   <td align="center">Delivered Qty Netto </td>
		   <td align="center">Vehicle No</td>
		   <td align="center">Customer Ticket</td> 
		   <td align="center">Accept Quantity Netto</td>
		   <td align="center">Validation Extra Carry</td>
		   </tr>
	   <%ELSE%>
		   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		   <td align="center">No</td>
		   <!-- <td align="center">&nbsp;</td> -->
		   <td align="center">Chitno</td>
		   <td align="center">Date</td>
		   <td align="center">Transporter</td> 
		   <td align="center">Delivered Qty Netto </td>
		   <td align="center">Vehicle No</td>
		   <td align="center">Customer Ticket</td> 
		   <td align="center">Accept Quantity Netto</td> 
		   </tr>
	   <%End if
		   if Not rsGRN.eof then
			  nCountofField=rsGRN.fields.count-1
			  no=1
			  v_deliveryqty_ori = 0
			  v_deliveryqty_prop = 0
			  v_acceptqtybruto = 0
			  v_acceptqty_ori = 0
			  v_acceptqty_prop = 0
			  While Not rsGRN.EOF

				  if no  mod 2 = 1 then
					 cColor = "#EEEEEE"
				  else
					 cColor = "#E5E5E5"
				  end if
				  
				  %>
			  <tr bgcolor="<%=cColor%>">
			  <td align="center" valign="top"><%=NO%></td>
				<%if SudahClose = "Y" then%> 
						<td align="center" title="" valign="top"><%=rsGRN("wbt_chitno")%></td>
						<td align="center" title="" valign="top"><%=rsGRN("postingdate")%></td>
						<td align="center" title="" valign="top"><%=rsGRN("transporter")%> : <%=rsGRN("transportername")%></td> 
						<td align="right" title="" valign="top"><%=FormatNumber(rsGRN("quantity"),0)%></td>
						<td align="center" title="" valign="top"><%=rsGRN("vehicleno")%></td>
						<td align="center" title="" valign="top"><%=rsGRN("customer_chitno")%></td>
						<td align="right" title="" valign="top"><%=FormatNumber(cdbl(to_nol(rsGRN("acceptquantity"))),0)%></td>
				<%End if%>

				<%if SudahClose<>"Y" Then%>
	                    <td align="center" title="" valign="top"><%=rsGRN("wbt_chitno")%></td>
				        <td align="center" title="" valign="top"><%=rsGRN("postingdate")%></td>
				        <td align="center" title="" valign="top"><%=rsGRN("transportercode")%> : <%=rsGRN("transportername")%></td>  
						<%If cdbl(to_nol(extracarry)) > 0 Then %>
							<td align="right" title="" valign="top"></td>
						<%Else%>
							<td align="right" title="" valign="top"><%=FormatNumber(rsGRN("deliveredqty"),0)%></td>
						<%End If%>
				        <td align="center" title="" valign="top"><%=rsGRN("vehicleno")%></td>
						<td align="center">
							<input style="text-align:left;" type="text" name="customer_chitno_<%=no%>" class="amount" value = "<%=request("customer_chitno_"&TRIM(no))%>" autocomplete="off">
						</td> 
						<%If cdbl(to_nol(extracarry)) > 0 Then %>
							<td align="center">
								<input style="text-align:right;" type="text" name="acceptquantity_<%=no%>" class="amount" value = "<%=request("acceptquantity_"&TRIM(no))%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
							</td>
						<%Else%> 
							<td align="center">
								<input style="text-align:right;" type="text" name="acceptquantity_<%=no%>" class="amount" value = "<%=FormatNumber(rsGRN("deliveredqty"),0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
							</td>
						<%End If%>
						<td align="center">
							<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
						</td>
						<input type="hidden" name="wbt_chitno_<%=no%>" value="<%=rsGRN("wbt_chitno")%>">
						<input type="hidden" name="postingdate_<%=no%>" value="<%=rsGRN("postingdate")%>">
						<input type="hidden" name="transportercode_<%=no%>" value="<%=rsGRN("transportercode")%>">
						<%If cdbl(to_nol(extracarry)) > 0 Then %> 
							<input type="hidden" name="deliveredqty_original_<%=no%>" value="<%=rsGRN("deliveredqtynetto_ori")%>">
						<%Else%>
							<input type="hidden" name="deliveredqty_original_<%=no%>" value="<%=rsGRN("deliveredqty")%>">
						<%End If%>
						<input type="hidden" name="vehicleno_<%=no%>" value="<%=rsGRN("vehicleno")%>">
					 </td>
				<%End if%>
			  </tr>
			  <%
				If SudahClose = "Y" Then 
					v_deliveryqty_ori = v_deliveryqty_ori + cdbl(to_nol(rsGRN("deliveredqtynetto_ori")))
					v_deliveryqty_prop = v_deliveryqty_prop + cdbl(to_nol(rsGRN("quantity")))
					v_acceptqtybruto = v_acceptqtybruto + cdbl(to_nol(rsGRN("acceptquantitybruto")))
					v_acceptqty_ori = v_acceptqty_ori + cdbl(to_nol(rsGRN("acceptquantitynetto_ori")))
					v_acceptqty_prop = v_acceptqty_prop + cdbl(to_nol(rsGRN("acceptquantity")))
				Else
					If cdbl(to_nol(extracarry)) > 0 Then
						v_deliveryqty_ori = v_deliveryqty_ori + cdbl(to_nol(rsGRN("deliveredqtynetto_ori")))
						v_deliveryqty_prop = v_deliveryqty_prop + cdbl(to_nol(rsGRN("quantity")))
						' v_acceptqtybruto = v_acceptqtybruto + cdbl(to_nol(rsGRN("acceptquantitybruto")))
						v_acceptqty_ori = v_acceptqty_ori + cdbl(to_nol(rsGRN("acceptquantitynetto_ori")))
						v_acceptqty_prop = v_acceptqty_prop + cdbl(to_nol(rsGRN("acceptquantity")))
					Else
						v_deliveryqty_ori = v_deliveryqty_ori + cdbl(to_nol(rsGRN("deliveredqty")))
						v_deliveryqty_prop = v_deliveryqty_prop + cdbl(to_nol(rsGRN("deliveredqty")))
						' v_acceptqtybruto = v_acceptqtybruto + cdbl(to_nol(rsGRN("acceptquantitybruto")))
						v_acceptqty_ori = v_acceptqty_ori + cdbl(to_nol(rsGRN("deliveredqty")))
						v_acceptqty_prop = v_acceptqty_prop + cdbl(to_nol(rsGRN("deliveredqty")))
					End If
				End If
			   rsGRN.MoveNext 
		  no=no+1
		  intRecordsPerPage = intRecordsPerPage - 1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsGRN.Close
   set rsGRN = Nothing%>
   
		<%IF SudahClose="Y" then%>
			<td align="right" title="" valign="top" colspan="3"><font face="Verdana" size="1" style="font-weight: bold">TOTAL</font></td>
			<td></td>
			<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_deliveryqty_prop),0)%> </font></td> 
			<td colspan = "2"></td>  
			<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_acceptqty_prop),0)%> </font></td>
		<%Else%>
			<%IF cdbl(to_nol(extracarry)) > 0 then%>
				<td align="right" title="" valign="top" colspan="3"><font face="Verdana" size="1" style="font-weight: bold">TOTAL</font></td>
				<td colspan = "4"></td>  
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_acceptqty_prop),0)%> </font></td>
			<%Else%>
				<td align="right" title="" valign="top" colspan="3"><font face="Verdana" size="1" style="font-weight: bold">TOTAL</font></td>
				<td colspan = "4"></td> 
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_acceptqty_prop),0)%> </font></td>
			<%End if%>
		<%End if%>
		<tr>
			<td colspan="12" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="no" value="<%=No%>">
		    <input type="hidden" name="purchaseshipmentno" value="<%=purchaseshipmentno%>">
			<input type="hidden" name="deliveryorderno" value="<%=deliveryorderno%>">
			<input type="hidden" name="creationdate" value="<%=creationdate%>">
			<input type="hidden" name="deliverydate" value="<%=deliverydate%>">
			<input type="hidden" name="deliveryorderdate" value="<%=deliveryorderdate%>">
			<input type="hidden" name="salescontractno" value="<%=salescontractno%>">
			<input type="hidden" name="quantity_do" value="<%=quantity_do%>">
			<input type="hidden" name="shipmenttype" value="<%=shipmenttype%>">
			<input type="hidden" name="destinationcountry" value="<%=destinationcountry%>">
			<input type="hidden" name="address" value="<%=address%>">
			<input type="hidden" name="remarks" value="<%=remarks%>">
			<input type="hidden" name="fcba" value="<%=fcba%>">
			<input type="hidden" name="shipmentfrom" value="<%=shipmentfrom%>">
			<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
			<input type="hidden" name="selecteddo" value="<%=selecteddo%>">
			<input type="hidden" name="extracarry" value="<%=extracarry%>">
			<input type="hidden" name="selectedlineno" value="<%=selectedlineno%>">
			<input type="hidden" name="doline" value="<%=doline%>">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<input type="hidden" name="productcode" value="<%=productcode%>">
			
			  <!-- Period Control In Here -->
			  <%IF Request("txtJenis")="Edit" then%>
			       <input type="hidden" name="txtJenis" value="Edit">
			       <input Type="Submit"  name="Add" value="Update">
			  <%Else%>
			       <input type="hidden" name="txtJenis" value="New">
			       <input Type="Submit"  name="Add" value="Submit">
			  <%End if%>
			  <!-- End Period Control In Here -->
			  <input Type="Reset"  name="Reset" value="Reset">
			  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		  </tr>
	   </TABLE>
	   </div>
	   </form>
	<%else%>
		<% Dim rsVH, intRecordsPerPage, cColor
		
		'Query untuk menampilkan table Delivery Order
		q	= " SELECT transporter" & _
			  " FROM deliveryorder" & _
			  " WHERE fcba = '" & fcba & "'" & _
			  " AND deliveryorderno = '" & deliveryorderno & "' AND lineno = '"& doline &"'" 

		' response.Write q
		Set rsVH = Server.CreateObject("ADODB.Recordset")
		rsVH.CursorType = 3
		rsVH.CursorLocation = 3
		rsVH.Open q, connstr
		if Not rsVH.eof then
			transporter =  rsVH("transporter")
		end if
		'query Untuk menampilkan isi di Field
		q	= " SELECT ssd.wbt_chitno, ssd.postingdate, ssd.vehicleno, ssd.transporter, ssd.quantity, ssd.acceptquantity, " & _
			  " ssh.purchaseshipmentno, ssh.purchasecontractno, ssh.deliveryorderno, ssd.customer_chitno, ssd.acceptquantitybruto,ssh.fcba " & _
			  " FROM purchaseshipmentdtl ssd" & _
			  " INNER JOIN purchaseshipmenthdr ssh"& _
			  " ON (ssd.purchaseshipmentno = ssh.purchaseshipmentno AND ssd.fcba = ssh.fcba)" & _
			  " WHERE ssh.fcba = '" & fcba & "'" & _
			  " /*AND ssh.purchasecontractno = '" & salescontractno & "'" & _
			  " AND ssh.deliveryorderno = '" & deliveryorderno & "'*/" & _
			  " AND ssh.purchaseshipmentno = '" & purchaseshipmentno & "'"

		' response.Write q
		Set rsVH = Server.CreateObject("ADODB.Recordset")
		rsVH.CursorType = 3
		rsVH.CursorLocation = 3
		rsVH.Open q, connstr
	%>
		<form action="purchaseshipment.asp" method="post" name="eForm">
			<table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td align="left" width="20%">Document No
				<td align="left">
					<input type="text" maxlength="30" id="chitno" name="wbt_chitno" size="30" value="<%=wbt_chitno%>" style="background:#FFFFCC;" autocomplete="off" onchange="javascript:getTicket();">
					<font color="#FF0000">*</font>
				</td>
				<td align="left" width="20%">Delivered Quantity (Netto)
					<td align="left">
					<input type="text" id="Netto" name="quantity" size="30" value="<%=quantity%>" autocomplete="Off" style="background:#FFFFCC;" onkeypress="return harusangka(event);">
					</td>
				</td>
			   <td align="left" width="20%">Customer Ticket 
				<td align="left">
					<input type="text" maxlength="30" id="customer_chitno" name="customer_chitno" size="30" value="<%=customer_chitno%>" style="background:#FFFFCC;" autocomplete="off" >
					<font color="#FF0000">*</font>
				</td> 
			</tr>
			<tr>
				<td align="left" width="20%">Date 
				<td align="left">
					<input type="text" class="datepicker" size="15" name="postingdate" value="<%=postingdate%>" >
				</td> 
				<td align="left" width="20%">Accepted Quantity (Netto)
					<td align="left">
					<input type="text" name="acceptquantity" size="30" value="<%=acceptquantity%>" autocomplete="Off" style="background:#FFFFCC;" onkeypress="return harusangka(event);">
					<font color="#FF0000">*</font>
					</td>
				</td>
				
				<td align="left" width="20%">Vehicle No
					<td align="left">
					<input type="text" maxlength="10" id="vehicleno" name="vehicleno" size="30" value="<%=vehicleno%>" autocomplete="Off" style="background:#FFFFCC;"  >
					</td>
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Transporter 
					<td align="left">
					<%If len(trim(cstr(""&transporter))) = 0 Then%>
						<input type="text" maxlength="50" id="transporter" name="transporter" size="30" value="<%=transporter%>" style="background:#FFFFCC;" autocomplete="Off" >	
					<%Else%>
						<input type="text" maxlength="50" id="transporter" name="transporter" size="30" value="<%=transporter%>" style="background:#eee;" autocomplete="Off" readonly >
					<%End if%>
						<font color="#FF0000">*</font>
					</td>
				</td> 
			</tr>
				<tr>
					<td align="left" colspan="2">
						<input type="hidden" name="purchaseshipmentno" value="<%=purchaseshipmentno%>">
						<input type="hidden" name="deliveryorderno" value="<%=deliveryorderno%>">
						<input type="hidden" name="creationdate" value="<%=creationdate%>">
						<input type="hidden" name="deliverydate" value="<%=deliverydate%>">
						<input type="hidden" name="deliveryorderdate" value="<%=deliveryorderdate%>">
						<input type="hidden" name="salescontractno" value="<%=salescontractno%>">
						<input type="hidden" name="quantity_do" value="<%=quantity_do%>">
						<input type="hidden" name="shipment_type" value="<%=shipment_type%>">
						<input type="hidden" name="destinationcountry" value="<%=destinationcountry%>">
						<input type="hidden" name="address" value="<%=address%>">
						<input type="hidden" name="remarks" value="<%=remarks%>">
						<input type="hidden" name="fcba" value="<%=fcba%>">
						<input type="hidden" name="shipmentfrom" value="<%=shipmentfrom%>">
						<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
						<input type="hidden" name="selecteddo" value="<%=selecteddo%>">
						<input type="hidden" name="selectedlineno" value="<%=selectedlineno%>">
						<input type="hidden" name="doline" value="<%=doline%>">
						<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
						<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
						<input type="hidden" name="productcode" value="<%=productcode%>">
						<!-- Period Control In Here -->
							<%IF Request("txtJenis")="Edit" then%>
								<input type="hidden" name="txtJenis" value="Edit">
								<input type="hidden" name="chitno" value="<%=chitno%>">
								<input Type="Submit"  name="Add" value="Update">
							<%Else%>
								<input type="hidden" name="txtJenis" value="New">
								<input Type="Submit"  name="Add" value="Add">
							<%End if%>
						<!-- End Period Control In Here -->
							<input Type="Reset"  name="Reset" value="Reset">
							<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
						
					</td>
				</tr>
			</table>

	   <table class="data">
		<BR>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
			<td align="left">NO</td>
			<td colspan="2" width="30">&nbsp;</td>
			<td align="left">NO TICKET</td>
			<td align="left">CUSTOMER TICKET</td>
			<td align="left">SALES SHIPMENT</td>
			<td align="left">POSTING DATE</td>
			<td align="left">TRANSPORTER</td>
			<td align="left">VEHICLE NO</td>
			<td align="left">DELIVERY QUANTITY NETTO</td>
			<td align="left">ACCEPT QUANTITY NETTO</td>
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
				
				
				<td valign="middle" width="50" colspan="2">
					<a href="purchaseshipment.asp?wbt_chitno=<%=rsVH("wbt_chitno")%>&deliveryorderno=<%=deliveryorderno%>&salescontractno=<%=salescontractno%>&purchaseshipmentno=<%=purchaseshipmentno%>&postingdate=<%=postingdate%>&transporter=<%=rsVH("transporter")%>&quantity_do=<%=quantity_do%>&vehicleno=<%=rsVH("vehicleno")%>&quantity=<%=rsVH("quantity")%>&acceptquantity=<%=rsVH("acceptquantity")%>&fcba=<%=rsVH("fcba")%>&acceptquantitybruto=<%=rsVH("acceptquantitybruto")%>&customer_chitno=<%=rsVH("customer_chitno")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
			
					<a href="purchaseshipment.asp?wbt_chitno=<%=rsVH("wbt_chitno")%>&deliveryorderno=<%=deliveryorderno%>&salescontractno=<%=salescontractno%>&purchaseshipmentno=<%=purchaseshipmentno%>&postingdate=<%=postingdate%>&transporter=<%=rsVH("transporter")%>&quantity_do=<%=quantity_do%>&vehicleno=<%=rsVH("vehicleno")%>&quantity=<%=rsVH("quantity")%>&acceptquantity=<%=rsVH("acceptquantity")%>&fcba=<%=rsVH("fcba")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				</td>
				    <td align="left" title="<%=rsVH("wbt_chitno").Name%>" valign="top"><%=rsVH("wbt_chitno")%></td>
				    <td align="left" title="<%=rsVH("customer_chitno").Name%>" valign="top"><%=rsVH("customer_chitno")%></td>
				    <td align="left" title="<%=rsVH("purchaseshipmentno").Name%>" valign="top"><%=rsVH("purchaseshipmentno")%></td>
				    <td align="left" title="<%=rsVH("postingdate").Name%>" valign="top"><%=rsVH("postingdate")%></td>
				    <td align="left" title="<%=rsVH("transporter").Name%>" valign="top"><%=rsVH("transporter")%></td>
				    <td align="left" title="<%=rsVH("vehicleno").Name%>" valign="top"><%=rsVH("vehicleno")%></td>
				    <td align="left" title="<%=rsVH("quantity").Name%>" valign="top"><%=rsVH("quantity")%></td> 
				    <td align="left" title="<%=rsVH("acceptquantity").Name%>" valign="top"><%=rsVH("acceptquantity")%></td>
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
 <script type="text/javascript">
	function handleClick(cb){
		if(cb.checked == 1){
			cb.value = 1;
		} else {
			cb.value = 0;
		}
	}

	function CheckCheckboxes1(chk){
		if(chk.checked == 1){
			chk.value = 1;
		} else {
			chk.value = 0;
		}
	}
	
	// function CheckCheckboxes1(chk){
	// 	var txt = chk.parentNode.parentNode.cells[8].getElementsByTagName('input')[0];
	// 	var txts = chk.parentNode.parentNode.cells[5].getElementsByTagName('input')[0];
		
	// 	if(chk.checked == true)
	// 	{
	// 		txt.readOnly = false;
	// 		txt.style.backgroundColor = "#fff";
	// 	}
	// 	else
	// 	{
	// 		txt.readOnly = true;
	// 		txt.style.backgroundColor = "#eee";
	// 	}
	// 	if(chk.checked == true)
	// 	{
	// 		txts.readOnly = false;
	// 		txts.style.backgroundColor = "#fff";
	// 	}
	// 	else
	// 	{
	// 		txts.readOnly = true;
	// 		txts.style.backgroundColor = "#eee";
	// 	}
	// }
</script> 

	<script>

	function getselectedfcba() {
		var sfcba = document.getElementById("selectedfcba");
		sfcba.value = document.getElementById("fcba").value;
	}

	function getselectedfcba2() {
		var sfcba = document.getElementById("selectedfcba2");
		sfcba.value = document.getElementById("fcba2").value;
	}

	function getselecteddo() {
		var scustomercode = document.getElementById("selecteddo");
		scustomercode.value = document.getElementById("deliveryorderno").value;
	}

	function getselecteddo2() {
		var scustomercode = document.getElementById("selecteddo2");
		scustomercode.value = document.getElementById("deliveryorderno2").value;
	}

	function getselectedrentby() {
		var ssupplier = document.getElementById("selectedrentby");
		ssupplier.value = document.getElementById("rentby").value;
	}

	function getselectedtransportercode() {
		var sdeliverydestination = document.getElementById("selectedtransportercode");
		sdeliverydestination.value = document.getElementById("transportercode").value;
	}
	
	function getselecteddoline() {
		var slineno = document.getElementById("selectedlineno");
		slineno.value = document.getElementById("doline").value;
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
		getDOList(); 
		getDOLine();
		getMessage();
		//getDOListExtraCarry();
		//getMessageExtraCarry(); 
		getselectedfcba();
		getselectedfcba2();
		getselecteddo();
		getselecteddo2(); 
		getselecteddoline();
	});
	/*
	$(function () {
        $("#extracarry").click(function () {
            if ($(this).is(":checked")) {
                $("#fcba_extracarry1").css("display", "none");
				$("#fcba_extracarry2").css("display", "block");
                $("#do_extracarry1").css("display", "none");
                $("#do_extracarry2").css("display", "block");
                $("#do_qty_extracarry1").css("display", "none");
                $("#do_qty_extracarry2").css("display", "block");
                $("#sc_extracarry1").css("display", "none");
                $("#sc_extracarry2").css("display", "block");
            } else {
                $("#fcba_extracarry1").css("display", "block");
                $("#fcba_extracarry2").css("display", "none");
                $("#do_extracarry1").css("display", "block");
                $("#do_extracarry2").css("display", "none");
                $("#do_qty_extracarry1").css("display", "block");
                $("#do_qty_extracarry2").css("display", "none");
                $("#sc_extracarry1").css("display", "block");
                $("#sc_extracarry2").css("display", "none");
            }
        });
    });*/
	</script>
</body>
</html>
