<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Delivery Order</title>
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
var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbaDO.asp', 'POST', PostStr, showlistfcba);
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
var getlistpo = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedpo = document.getElementById("selectedpo").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fcba='+selectedfcba+'&fccode='+selectedpo;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistPODO.asp', 'POST', PostStr, showlistpo);
}

// The function for handling the response from the server
var showlistpo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("polist").innerHTML = "";
	document.getElementById("polist").innerHTML = response;

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
var getliststono = function () {
	// build up the post string when passing variables to the server side page
	
	var stono = document.getElementById("stono").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fcba='+selectedfcba+'&stono='+stono;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getliststono.asp', 'POST', PostStr, showliststono);
}

// The function for handling the response from the server
var showliststono = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("stonolist").innerHTML = "";
	document.getElementById("stonolist").innerHTML = response;

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
var getPODO = function () {
	// build up the post string when passing variables to the server side page
	var stono = document.getElementById("stono").value;
	var stolineno = document.getElementById("stolineno").value;
	var PostStr = 'stono='+stono+'&stolineno='+stolineno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getSTODet.asp', 'POST', PostStr, showMessageResponse);
	
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.supplieraddress);
    // update the Div to show the result from the server
    document.getElementById("itemcode").value = s.itemcode;
	document.getElementById("stockname").value = s.stockname;
	document.getElementById("quantityorder").value = s.quantityorder;
	// document.getElementById("address").innerHTML = document.getElementById("supplieraddress").value;
	// alert(s.suppliername);
    
}

// launched from button click 
var getStockCodeSTO = function () {
	// build up the post string when passing variables to the server side page
	// var linenokey = document.getElementById("linenokey").value;
	// var rentby = document.getElementById("linenokey").value;
	var PostStr = 'linenokey='+document.getElementById("linenokey").value+'&businessunit='+document.getElementById("businessunit").value+'&fccode='+document.getElementById("fccode").value;
	// use the generic function to make the request
	// alert(PostStr);
	doAJAXCall('ajax/getStockCodeSTO.asp', 'POST', PostStr, showStockCodeSTOResponse);
	
}

// The function for handling the response from the server
var showStockCodeSTOResponse = function (oXML) { 

    // get the response text, into a variable
    var response = oXML.responseText;
    // alert(response);
    var s = JSON.parse(response);
    // update the Div to show the result from the server
    document.getElementById("quantity").value = s.quantity;
    document.getElementById("itemcode").value = s.itemcode;
	document.getElementById("stockname").value = s.stockname;
	
    
};

</script>
<!-- <script language="JavaScript">
function getFcba(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("businessunit").innerHTML="";
	  // document.getElementById("rentby").innerHTML="";
	  // document.getElementById("receviedtask").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {// code for IE7+, Firefox, Chrome, Opera, Safari
	  xmlhttp=new XMLHttpRequest();
	  StockName=new XMLHttpRequest();
	  }
	else
	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{
		// alert(xmlhttp.responseText);
		document.getElementById("businessunit").innerHTML=xmlhttp.responseText;
		// document.getElementById("rentby").innerHTML=xmlhttp.responseText;
		// document.getElementById("receviedtask").innerHTML="";
		// alert(businessunit);
		}
	  }
	xmlhttp.open("GET","ajax/getFcbaDO.asp?fccompanycode="+str+"",true);
	xmlhttp.send();
}

function getPODO(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("businessunit").innerHTML="";
	  // document.getElementById("purchaseby").innerHTML="";
	  // document.getElementById("receviedtask").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {// code for IE7+, Firefox, Chrome, Opera, Safari
	  xmlhttp=new XMLHttpRequest();
	  StockName=new XMLHttpRequest();
	  }
	else
	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{
		// alert(xmlhttp.responseText);
		document.getElementById("fccode").innerHTML=xmlhttp.responseText;
		// document.getElementById("purchaseby").innerHTML=xmlhttp.responseText;
		// document.getElementById("receviedtask").innerHTML="";
		// alert(businessunit);
		}
	  }
	// xmlhttp.open("GET","ajax/getRentBy.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
	xmlhttp.open("GET","ajax/getPODO.asp?fcba="+str+"&rowstateid=Ordered",true);
	xmlhttp.send();
}

function getSTOLine(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("stono").innerHTML="";
	  document.getElementById("businessunit").innerHTML="";
	  // document.getElementById("receviedtask").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {// code for IE7+, Firefox, Chrome, Opera, Safari
	  xmlhttp=new XMLHttpRequest();
	  StockName=new XMLHttpRequest();
	  }
	else
	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{
		// alert(xmlhttp.responseText);
		document.getElementById("stolineno").innerHTML=xmlhttp.responseText;
		// document.getElementById("requestedqty").innerHTML=xmlhttp.responseText;
		// document.getElementById("receviedtask").innerHTML="";
		// alert(PRLINEID);
		}
	  }
	// xmlhttp.open("GET","ajax/getRentBy.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
	xmlhttp.open("GET","ajax/getSTOLine.asp?stono="+str+"&businessunit="+document.getElementById("businessunit").value+"",true);
	xmlhttp.send();
}


function validasi(form){
 
	if (form.stono.value == ""){
		alert("STONO No Tidak Boleh Kosong.");
		form.stono.focus();
		return (false);
	}

	
	var sto_quantity = form.sto_quantity.value;
	var delivery_quantity = form.delivery_quantity.value;
	var dtl_delivery = form.dtl_delivery.value;
	var nHitungDoWeight = parseFloat(dtl_delivery)+parseFloat(delivery_quantity);
	//var nSelisihDoWeight = parseFloat(sto_quantity)+parseFloat(dtl_delivery);
	
	if (nHitungDoWeight > sto_quantity)	{
		 alert("Total Do Weight ( Detail:"+delivery_quantity+"+"+dtl_delivery+"="+nHitungDoWeight+" ) Lebih Besar Dari Total Weight ( Header="+sto_quantity+" )");
		 form.delivery_quantity.focus();
		 return(false);
	}
		
	return(true);
}

					
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script> -->

<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261015a"></script>
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
 <%
 '-------------awal-----------------------
 
 Dim xCode,fccode,cmdView,PoCode,DO_DATE,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim RowStateId,RowStateHdr
 Dim PRCODE, PRLINEID
 Dim tempStringError, rsTempPurchaseRequest
 Dim totalPRData
 Dim ApprovedQty, ORDEREDQTY
 Dim strQueryStatus, rsQueryStatus
 Dim sqlupdateqtypr, rsupdateqtypr
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = "" 
 Dim IsInBridge, IsFrezee,HdrState
 
 Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 ' if request("fccode")="" then
 '    fccode="" '----- Nantinya Ambil dari Format Master
 ' else
 '    fccode = request("fccode")
 ' end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("DO_DATE")<>"" then
   DO_DATE   =Request("DO_DATE")
 else
   DO_DATE   =textDate(date)
 end if
 
  DO_DATE = FormatDateTime(DO_DATE)
 
 xCode = Request("xCode")
 
 '---- Data Details
 ' Add by NAF, 2012-12-1
 LINENOKEY=Request("LineNoKey")
 ' End Added by NAF
 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID")
 fccompanycode = Request("fccompanycode")
 businessunit = Request("businessunit")
 rentby = Request("rentby")
 sto_quantity = Request("sto_quantity")
 remarks = Request("remarks")
 delivery_quantity = Request("delivery_quantity")
 do_no = Request("do_no")
 stolineno = Request("stolineno")
 stono = Request("stono")
 StockName = Request("StockName")
 doquantity = Request("doquantity")
 ponumber = Request("ponumber")
 DOLINE = Request("DOLINE")
 stonumber = Request("stonumber")
 do_quantity = Request("do_quantity")
 fccode = Request("fccode")
 selectedfcba = Request("selectedfcba")
 selectedpo = Request("selectedpo")
if "" & use_ppn = "" then use_ppn = "N" 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
		connstr.BeginTrans
		
		strFieldDetail = " SELECT COUNT(1) total FROM deliveryorderpurchasedtl " & _
                         " WHERE do_no = '" & do_no & "'" & _
                         "       AND fcBA = '" & businessunit & "'"								
		' Response.Write strFieldDetail
		Set rsQty = Server.CreateObject("ADODB.Recordset")
		rsQty.Open strFieldDetail, connstr							
		if not rsQty.Eof then
			has_detail = cdbl (to_nol(rsQty("total")))
		End If
		
		if has_detail = 0 then
			q = "Delete from deliveryorderpurchasehdr where do_no = '"& do_no &"' and  POCODE='" & fccode & "' and fcBA='" & businessunit & "'"
				' Response.Write q
			connStr.execute(q)
		else
			tempStringError = "Delivery Order No '"& do_no &"' Has Detail, Please Delete All Data Detail First!"				 
			response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
		End If
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then
		connstr.BeginTrans
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
			
			q = " delete from deliveryorderpurchasedtl " & _
				" where fcBA = '" & businessunit & "'" & _
				" and do_no = '"& do_no & "'" & _
				" and stono = '"& stonumber & "'"
				' Response.Write q
				connStr.execute(q)

			if err.number = 0 then
			   connStr.CommitTrans
			end if
			cmdView	= "True"

	elseif Request("txtJenis")="New" then
			Dim rsSupplyingDeclarationHdr
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open "Select poCode, do_no from deliveryorderpurchasehdr where do_no = '"& do_no &"' and poCode='" & fccode & "' AND FCBA = '" & businessunit & "'", connstr
			If rsSupplyingDeclarationHdr.EOF Then
				q = " INSERT INTO deliveryorderpurchasehdr ( company, do_no, poCode, DO_DATE, description, fcentry, fcedit, fcip, fcba, lastupdate, lasttime ) " & _
				     " VALUES (  '" & fccompanycode & "'" & _
				     "         , '" & do_no & "'" & _
				     "         , '" & fccode & "'" & _
					 "         , TO_DATE('" & DO_DATE & "', 'mm/dd/yyyy')" & _
					 "         , '" & description & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					 "         , '" & businessunit & "'" & _
					 "         , SYSDATE" & _
					 "         , TO_CHAR(SYSDATE,'hh24:mi'))"
				 ' response.write q
				 connStr.execute(q)
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing
					
			' Added by NAF, 2012-12-01
				 strLineMax = " select nvl(max(DOLINE),0) maxLineNoKey from deliveryorderpurchasedtl where fcBA='" & businessunit & "'"

				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
				end if
							
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 
				' ' End of Added by NAF, 2012-12-01
				set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select sum(delivery_quantity) total from deliveryorderpurchasedtl " & _
					  " where " & _
					  " fcBA='" & businessunit & "' and itemcode='"& ItemCode &"' and stono ='"& stonumber &"'"
					  ' response.write q
					  ' response.write sto_quantity
					  ' response.write delivery_quantity
					 
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
						dtl_delivery = cdbl(to_nol(rsTempPurchaseRequest("total"))) 
						rsTempPurchaseRequest.movenext
					wend 
						
					End if
					
					rsTempPurchaseRequest.Close
					set rsTempPurchaseRequest = Nothing
					If cdbl (to_nol(sto_quantity)) - cdbl (to_nol(delivery_quantity)) >= cdbl (to_nol(dtl_delivery))  Then
						
						q="Insert into  deliveryorderpurchasedtl (DO_NO, STONO, STOLINENO, ITEMCODE, STOCKNAME, STO_QUANTITY, DELIVERY_QUANTITY, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, rowstateid, DOLINE ) " & _
						 "values ('" & DO_NO & "','" & stonumber & "','" & STOLINENO & "','"& ITEMCODE &"','"& STOCKNAME &"','"& STO_QUANTITY &"','"& DELIVERY_QUANTITY &"','" & OTHERITEMDESC & "','" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
						 Request.ServerVariables("REMOTE_ADDR") &"','" & businessunit & "', sysdate, TO_CHAR(SYSDATE,'hh24:mi'),'','"& tempMaxLineNo &"')"
						 ' response.write q
						 ' response.write stono
						connStr.execute(q)				
						
						 q=	 " Update stocktransferorderdtl set DO_QUANTITY='" & DELIVERY_QUANTITY & "'" & _
							 " where  STONO='" & stonumber & "' and STOLINENO='" & STOLINENO & "' and fcBA='" & businessunit & "'"
							 ' response.write q
							connStr.execute(q)

						tempStringError = "DO Saved Successfully with PO Info, if you want to modify just click Edit Button, Thank you."				 
						response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"

						'  q= " Update deliveryorderpurchasedtl set delivery_quantity='" & delivery_quantity & "'" & _
						'  ", LASTUPDATE = sysdate " & _
						'  ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						'  " where  do_no='" & do_no & "' and fcBA='" & businessunit & "' and itemcode='"& ItemCode &"' and stono = '"& stonumber &"'"
						'  ' response.write q
						'  response.write dtl_delivery
						'  response.write delivery_quantity
						'   response.write ttl_delivery_quantity
						' connStr.execute(q)	

						' if err.number  = 0 then
						' 	 connStr.CommitTrans
						' end if
					else
						tempStringError = "Anda melebihi Quantity Stock Transfer Order !"				 
						response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"	
					end if
						' End if
				' End If			
			' End if
		' End If			
	' End If
	' End If
	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		If cdbl (to_nol(sto_quantity)) - cdbl (to_nol(delivery_quantity)) >= cdbl (to_nol(dtl_delivery))  Then
				q =	" Update deliveryorderpurchasedtl" & _
					" set delivery_quantity = '" & delivery_quantity & "'" & _
					"     , remarks = '" & otheritemdesc & "'" & _
					"     , fcentry = '" & Session("CodeName") & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"     , lastupdate = sysdate " & _
					"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					" 		where do_no = '" & do_no & "'" & _
					"       and stono = '" & stonumber & "'" & _
					"       and fcBA = '" & businessunit & "'"
 			' response.write q
 			connStr.execute(q)
 		else
			tempStringError = "Anda melebihi Quantity Stock Transfer Order !"				 
			response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"	
		end if
			'  if err.number  = 0 then
			' 	 connStr.CommitTrans
			' end if
			End If			
		End If
	' End If
' End If
cmdView	= "True"	

' Untuk menampilkan value header jika ada data.
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")
 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select * from deliveryorderpurchasehdr  " & _
  " where do_no='" & do_no & "'"
  ' response.write q
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    DO_DATE=rsPurchase("DO_DATE")
	DO_DATE = FormatDateTime(DO_DATE)
	do_no=rsPurchase("do_no")
	businessunit = rsPurchase("fcba")
	fccode =  rsPurchase("poCode")
	selectedfcba = rsPurchase("fcba")
	selectedpo = rsPurchase("poCode")
	fccompanycode = rsPurchase("company")
	remarks = rsPurchase("description") 
 End if

 val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & businessunit & "'" & _
	        "       and to_date ('" & DO_DATE & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & DO_DATE & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Purchase')"
'response.Write oSQLQuery
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
oRecordSet.Close : Set oRecordSet = Nothing

 
 dim rsFrez
 set rsFrez = server.CreateObject("Adodb.recordset")
	q=" select distinct frezee  from Purchaseorderdetails " & _
	  " where PoCode='" & fccode & "' and " & _
	  " fcBA='" & businessunit & "'"
 rsFrez.Open q, connstr
 if Not rsFrez.Eof then
	IsFrezee  = rsFrez("frezee")
 else
    IsFrezee ="0" 
 End if
 rsFrez.Close
 set rsFrez = Nothing 
 
  'response.write "IsInBridge--> " & IsInBridge
' if err.number <> 0 then
' 	connStr.RollbackTrans
'     response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
' End if
 '-------------ISI Payment Form
%>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-truck" aria-hidden="true"></i>Header Delivery Order</div><% end if %>
 <form name="frmEntry" method="post" action="deliveryorder_purchase.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Company</td>
				<td align="left">
					<select name="fccompanycode" value="<%=fccompanycode%>" id="fccompanycode" onchange="javascript:getlistfcba();javascript:getlistrentby();javascript:getdeliverydest();" class="chosen-select" style="width: 70%">
						<option value=""></option> 
							<% 	Dim rsComboC
								q = " SELECT DISTINCT fccompanycode " & _
								 	" FROM (SELECT businessunit.fccompanycode " & _
								    " FROM user_fcba_tab " & _
								    " INNER JOIN businessunit " & _
								    " ON user_fcba_tab.fcba = " & _
								    " businessunit.fccode " & _
								    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"')"
								' response.write q
								Set rsComboC = Server.CreateObject("ADODB.Recordset")
								rsComboC.CursorType = 3
								rsComboC.CursorLocation = 3
								rsComboC.Open q, connstr
								if not rsComboC.eof then%>
								<%
								Do While Not rsComboC.eof
								%>
									<option value="<%=rsComboC("fccompanycode")%>" 
										<%if cstr(""&fccompanycode)=cstr(""&rsComboC("fccompanycode")) then %> selected 
										<%End If%>><%=rsComboC("fccompanycode")%>
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
				<td align="left">Delivery Order No</td>
				<td align="left">
					<input type="text" maxlength="50" name="do_no" size="30" value="<%=do_no%>" autocomplete="off">
					<font color="#FF0000">*</font>
				</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
				<td align="left">
					<div id="fcbalist"></div>
					<font color="#FF0000">*</font>
				</td>
				<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			<td align="left">Purchase Order No</td>
			<td align="left">
				<div id="polist"></div>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedpo" name="selectedpo" value="<%=selectedpo%>">
		</tr>
		<tr>
			<td align="left">Date</td>
				<td align="left">
					<input class="datepicker" type="text" maxlength="10" name="DO_DATE" size="10" value="<%=DO_DATE%>">
				</td>
				<td align="left">Remarks</td>
				<td align="left">
					<textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea>
				</td>
			</tr>
		
		</tr>
		</tr>
		
		<tr><td colspan="4"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		
		<a href="deliveryorder_purchase.asp?do_no=<%=do_no%>&fccode=<%=fccode%>&DO_DATE=<%=DO_DATE%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&txtJenis=HeaderDelete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" AND LEN(TRIM(do_no)) > 0 Then
	Dim rsPurchaseOrder,intRecordsPerPage,cColor

	q = " select *  from deliveryorderpurchasedtl dod" & _
		" inner join deliveryorderpurchasehdr doh ON (dod.do_no = doh.do_no and dod.fcba = doh.fcba)" & _
		" where dod.do_no = '" & do_no & "'" 
	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.CursorType = 3
	rsPurchaseOrder.CursorLocation = 3
	rsPurchaseOrder.Open q, connstr
%>
   

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="deliveryorder_purchase.asp" method="post" name="eForm" onSubmit="javascript:return validasi(this)">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
   		<tr>
			<td align="left">STO. Number.</td>
			<td align="left">
			<%IF txtJenis="Edit" Then%>
				 <input type="text" size="15" name="stonumber" id="stono" value="<%=stonumber%>" readonly="True" style="background:#eee">&nbsp;<font color="#FF0000">*
			<%else%>
				 <select name="stonumber" value="<%=stonumber%>" id="stono" onchange="javascript:getliststono();" class="chosen-select" style="width: 40%">
					<option value=""></option> 
						<% 	Dim rsComboS
							q = " select sth.stono as stono from stocktransferorderdtl std" & _
								" inner join stocktransferorderhdr sth on (sth.pocode=std.pocode and sth.fcba=std.fcba and sth.stono = std.stono) " & _
								" where sth.fcba = '"& businessunit &"' and sth.pocode= '"& fccode &"' order by std.STOLINENO "  
							' response.write q
							Set rsComboS = Server.CreateObject("ADODB.Recordset")
							rsComboS.CursorType = 3
							rsComboS.CursorLocation = 3
							rsComboS.Open q, connstr
							if not rsComboS.eof then%>
							<%
							Do While Not rsComboS.eof
							%>
								<option value="<%=rsComboS("stono")%>" 
									<%if cstr(""&stonumber)=cstr(""&rsComboS("stono")) then %> selected 
									<%End If%>><%=rsComboS("stono")%>
								</option> 
							<%rsComboS.movenext
							loop
							End if
							rsComboS.Close
							set rsComboS = Nothing
							q=""
						%>
				</select> 
				<font color="#FF0000">*</font>
			<%End if%>
				
			</td>
		</tr>
   		<tr>
			<td align="left">STO. Line No</td>
			<td align="left">
			<%IF txtJenis="Edit" Then%>
				 <input type="text" size="15" name="stolineno" value="<%=stolineno%>" readonly="True" style="background:#eee">&nbsp;<font color="#FF0000">*
			<%else%>
				<div id="stonolist"></div>
			<%End if%>
				
			</td>
		</tr>
		   <td  align="left">Stock Code</td>
		   <td align="left">
				 <input type="text" size="20" id="itemcode" name="ITEMCODE" value="<%=ITEMCODE%>" style="background:#eee" readonly >
			</td>
			</tr>
		<tr>
			<td align="left">Stock Name</td>
			<td align="left"><input type="text" size="50" id="stockname" name="StockName" value="<%=StockName%>" style="background:#eee" readonly >
			</td>
		</tr>
		<tr>
			<td align="left">STO Quantity</td>
			<td align="left"><input type="text" size="15" id="quantityorder" name="sto_quantity" value="<%=sto_quantity%>" style="background:#eee" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Delivery Quantity</td>
			<td align="left"><input type="text" size="15" name="delivery_quantity" value="<%=delivery_quantity%>" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td align="left">Description</td>
			<td align="left">
				<textarea  name="OTHERITEMDESC"  cols="50" rows="3"><%=OTHERITEMDESC%>
				</textarea>
			</td>
		</tr>

		<tr>
		   <td align="left"></td>
		   <td></td>
		</tr>
			
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
		  <input type="hidden" name="dtl_delivery" value="<%=dtl_delivery%>">
	      <input type="hidden" name="fccode" value="<%=fccode%>">
	      <input type="hidden" name="do_no" value="<%=do_no%>">
		  <input type="hidden" name="DO_DATE" value="<%=DO_DATE%>">
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		  <input type="hidden" name="businessunit" value="<%=businessunit%>">
		  <input type="hidden" name="RowStatehdr" value="<%=RowStatehdr%>">
		  <input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		  <input type="hidden" name="selectedpo" value="<%=selectedpo%>">
		  <input Type="Hidden" Value="True" Name="cmdView">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else  %>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset" name="Reset" value="Reset">
		  <input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="middle" colspan="2"> ACTION</TD>
	<TD align="left">DO NO</TD>
	<TD align="left">STO NO</TD>
	<TD align="left">PO CODE</TD>
	<TD align="left">ITEM CODE</TD>
	<TD align="left">ITEM NAME</TD>
	<TD align="left">STO QUANTITY</TD>
	<TD align="left">DELIVERY QUANTITY</TD>
	<TD align="left">DESCRIPTION</TD>
	</tr>
	</thead>
	<tbody>
   <%
   
     Set rsPO = Server.CreateObject("ADODB.Recordset")
	 rsPO.Open " select purchaseorder_api.GetPoState('" & businessunit & "' ,'" & fccode & "') HdrState from dual" , connstr
	 IF Not rsPO.eof then
		HdrState =  rsPO("HdrState")
	 End if
	 rsPO.Close
	 set rsPO = Nothing
   
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
			  <!-- Validasi Period Control taro sini-->	
			    <td valign="middle" width="10">
				   <a href="deliveryorder_purchase.asp?stonumber=<%=rsPurchaseOrder("stono")%>&stolineno=<%=rsPurchaseOrder("stolineno")%>&itemcode=<%=rsPurchaseOrder("itemcode")%>&stockname=<%=rsPurchaseOrder("stockname")%>&otheritemdesc=<%=rsPurchaseOrder("remarks")%>&do_no=<%=rsPurchaseOrder("do_no")%>&fccode=<%=rsPurchaseOrder("pocode")%>&do_date=<%=rsPurchaseOrder("do_date")%>&stono=<%=rsPurchaseOrder("stono")%>&businessunit=<%=rsPurchaseOrder("fcba")%>&sto_quantity=<%=rsPurchaseOrder("sto_quantity")%>&delivery_quantity=<%=rsPurchaseOrder("delivery_quantity")%>&txtJenis=Edit">
				   	<img src="IMAGES/edit.gif" alt="Edit Data : <%=request("do_no")%>" border="0"></a>
				</td>
				<td valign="middle" width="10">
				   <a href="deliveryorder_purchase.asp?stonumber=<%=rsPurchaseOrder("stono")%>&stolineno=<%=rsPurchaseOrder("stolineno")%>&itemcode=<%=rsPurchaseOrder("itemcode")%>&stockname=<%=rsPurchaseOrder("stockname")%>&otheritemdesc=<%=rsPurchaseOrder("remarks")%>&do_no=<%=rsPurchaseOrder("do_no")%>&fccode=<%=rsPurchaseOrder("pocode")%>&do_date=<%=rsPurchaseOrder("do_date")%>&stono=<%=rsPurchaseOrder("stono")%>&businessunit=<%=rsPurchaseOrder("fcba")%>&sto_quantity=<%=rsPurchaseOrder("sto_quantity")%>&delivery_quantity=<%=rsPurchaseOrder("delivery_quantity")%>&txtJenis=Delete">
				   <img src="IMAGES/delete.gif" alt="Delete Data : <%=request("do_no")%>" border="0" onClick="return confirm('Are you sure to delete');">
				   </a>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("do_no").Name%>" valign="top">
					<%=rsPurchaseOrder("do_no")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("stono").Name%>" valign="top">
					<%=rsPurchaseOrder("stono")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("pocode").Name%>" valign="top">
					<%=rsPurchaseOrder("pocode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("itemcode").Name%>" valign="top">
					<%=rsPurchaseOrder("itemcode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("stockname").Name%>" valign="top">
					<%=rsPurchaseOrder("stockname")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("sto_quantity").Name%>" valign="top">
					<%=rsPurchaseOrder("sto_quantity")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("delivery_quantity").Name%>" valign="top">
					<%=rsPurchaseOrder("delivery_quantity")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("remarks").Name%>" valign="top">
					<%=rsPurchaseOrder("remarks")%>
				</td>
			</tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("<tr><td colspan=""12"" class=""fbp-doc-empty"">Data is empty</td></tr>")
   end if
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </tbody>
   </TABLE>
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
	 	<script type="text/javascript">
			function handleClick(cb){
				if(cb.checked == 1){
					cb.value = "Y";
				} else {
					cb.value = "N";
				}
	        }
		</script>

		<script>

		function getselectedfcba() {
			var sfcba = document.getElementById("selectedfcba");
			sfcba.value = document.getElementById("fcba").value;
		}

		function getselectedpo() {
			var scustomercode = document.getElementById("selectedpo");
			scustomercode.value = document.getElementById("fccode").value;
		}

		function getselectedrentby() {
			var ssupplier = document.getElementById("selectedrentby");
			ssupplier.value = document.getElementById("rentby").value;
		}

		function getselectedtransportercode() {
			var sdeliverydestination = document.getElementById("selectedtransportercode");
			sdeliverydestination.value = document.getElementById("transportercode").value;
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
			getlistpo();
			getliststono();
			getselectedfcba();
			getselectedpo();
			
		});
		</script>
</body>
</html>
