<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Stock Transfer Order</title>
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

function validateFloatKeyPress(el, evt) {
    var charCode = (evt.which) ? evt.which : event.keyCode;
    var number = el.value.split('.');
    if (charCode != 46 && charCode > 31 && (charCode < 48 || charCode > 57)) {
        return false;
    }
    //just one dot
    if(number.length>1 && charCode == 46){
         return false;
    }
    //get the carat position
    var caratPos = getSelectionStart(el);
    var dotPos = el.value.indexOf(".");
    if( caratPos > dotPos && dotPos>-1 && (number[1].length > 1)){
        return false;
    }
    return true;
}

//thanks: http://javascript.nwbox.com/cursor_position/
function getSelectionStart(o) {
	if (o.createTextRange) {
		var r = document.selection.createRange().duplicate()
		r.moveEnd('character', o.value.length)
		if (r.text == '') return o.value.length
		return o.value.lastIndexOf(r.text)
	} else return o.selectionStart
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
	doAJAXCall('ajax/getlistfcbasto.asp', 'POST', PostStr, showlistfcba);
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
	var PostStr = 'fccode='+selectedpo+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistposto.asp', 'POST', PostStr, showlistpo);
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
var getlistrentby = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedrentby = document.getElementById("selectedrentby").value;
	var fccompanycode = document.getElementById("fccompanycode").value;
	var PostStr = 'rentby='+selectedrentby+'&fccompanycode='+fccompanycode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistrentby.asp', 'POST', PostStr, showlistrentby);
}

// The function for handling the response from the server
var showlistrentby = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("rentbylist").innerHTML = "";
	document.getElementById("rentbylist").innerHTML = response;

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
var getlisttransportercode = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedtransportercode = document.getElementById("selectedtransportercode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'transportercode='+selectedtransportercode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlisttransporter.asp', 'POST', PostStr, showlisttransportercode);
}

// The function for handling the response from the server
var showlisttransportercode = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("transportercodelist").innerHTML = "";
	document.getElementById("transportercodelist").innerHTML = response;

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
var getRent = function () {
	// build up the post string when passing variables to the server side page
	var transportercode = document.getElementById("transportercode").value;
	var rentby = document.getElementById("rentby").value;
	var PostStr = 'transportercode='+transportercode+'&rentby='+rentby;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getRentBy2.asp', 'POST', PostStr, showMessageResponse);
	
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.supplieraddress);
    // update the Div to show the result from the server
    document.getElementById("transportername").value = s.transportername;
	document.getElementById("transporteraddress").value = s.transporteraddress;
	// document.getElementById("address").innerHTML = document.getElementById("supplieraddress").value;
	// alert(s.suppliername);
    
}

// launched from button click 
var getStockCodeSTO = function () {
	// build up the post string when passing variables to the server side page
	var linenokey = document.getElementById("linenokey").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedpo = document.getElementById("selectedpo").value;
	var PostStr = 'linenokey='+linenokey+'&fccode='+selectedpo+'&fcba='+selectedfcba;
	// var PostStr = linenokey='+document.getElementById("linenokey").value+'&businessunit='+document.getElementById("businessunit").value+'&fccode='+document.getElementById("fccode").value;
	// use the generic function to make the request
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
	
    
}

var reloadStoCode = function () {
	var stodate =$( ".datepicker" ).val();
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Procurement&txtTable=StockTransferOrder&stodateurl="+stodate );
}

</script>
<script>
	function hitungrumus() {
		var squantity = $("#quantity").val();
		var sunitprice = $("#unitprice").val();
		// var squantityaftertolerance = $("#quantityaftertolerance").val();
		// var ssalesprices = $("#salesprices").val();
		// var svat = $("#vat").val();
		var newsunitprice = sunitprice.replace(/,/g, '');
		
		//Untuk menghitung Total Quantity Tolerance
		totalamount = squantity * newsunitprice 
		// alert(totalamount);
		if (!isNaN(totalamount)) {
    	 	$("#amount").val(totalamount);
    	}
    	else{
    		$("#amount").val();
    	}
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
 <%
 '-------------awal-----------------------
 
 Dim xCode,fcCode,cmdView,PoCode,STODATE,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
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

 if request("fccOde")="" then
    fcCode="" '----- Nantinya Ambil dari Format Master
 else
    fcCode = request("fccOde")
 end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if request("stodateurl") <>"" then
   STODATE   = request("stodateurl")
 elseif Request("STODATE")<>"" then
   STODATE   =Request("STODATE")
 else
   STODATE   =textDate(date)
 end if
 
  STODATE = FormatDateTime(STODATE)
 
 xCode = Request("xCode")
 IsInBridge = Request("IsInBridge")
 Remarks = Request("Remarks")
 
 '---- Data Details
 ' Add by NAF, 2012-12-1
 ' End Added by NAF
 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 quantity = Request("quantity") 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID")
 fccompanycode = Request("fccompanycode")
 businessunit = Request("businessunit")
 rentby = Request("rentby")
 transportercode = Request("transportercode")
 transportername = Request("transportername")
 transporteraddress = Request("transporteraddress")
 termofpayment = Request("termofpayment")
 dp = Request("dp")
 ppn = Request("ppn")
 stono = Request("stono")
 StockName = Request("StockName")
 unitprice = Request("unitprice")
 ppnptg = Request("ppnptg")
 ppnrp = Request("ppnrp")
 totalamount = Request("totalamount")
 doquantity = Request("doquantity")
 ponumber = Request("ponumber")
 stolineno = Request("stolineno")
 selectedfcba = Request("selectedfcba")
 selectedpo = Request("selectedpo")
 selectedtransportercode = Request("selectedtransportercode")
 selectedrentby = Request("selectedrentby")
 linenokey = request("linenokey")

if "" & use_ppn = "" then use_ppn = "N" 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "Ordered" or Request("txtJenis")= "CompletedDetails"  or  Request("txtJenis")= "CanceledDetails"  or  Request("txtJenis")= "UnRelease" Then
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
       connstr.BeginTrans

				
		q="Delete from stocktransferorderdtl where stono = '"& stono &"' and  POCODE='" & fcCode & "' and fcBA='" & businessunit & "'"
		connStr.execute(q)

		q="Delete from stocktransferorderhdr where stono = '"& stono &"' and  POCODE='" & fcCode & "' and fcBA='" & businessunit & "'"
		' response.write q
		connStr.execute(q)

		if err.number  = 0 then
			connStr.CommitTrans
		end if

	   cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans
	   q =" DELETE FROM stocktransferorderdtl WHERE stono = '"& stono &"' AND stolineno ='" & stolineno & "' AND fcBA='" & businessunit & "'"
	   'response.write q
		connStr.execute(q)

			   'Untuk menghhitung jumlah total di header
				dtl_totalamount = 0
				dtl_charge = 0
				totalamount = 0
				set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
				q="select *  from stocktransferorderdtl " & _
				  " where stono='"& stono &"' and fcBA='" & businessunit & "'"
				rsTempPurchaseRequest.Open q, connstr
				if Not rsTempPurchaseRequest.Eof then
					while Not rsTempPurchaseRequest.Eof 
						dtl_totalamount = dtl_totalamount + cdbl(rsTempPurchaseRequest("amount")) 
						rsTempPurchaseRequest.movenext
					wend 
					
				End if
				
				rsTempPurchaseRequest.Close
				set rsTempPurchaseRequest = Nothing
				dim total_all
				totalamount = totalamount + cdbl(to_nol(amount))
				 totalppn = dtl_totalamount  * cdbl(to_nol(ppnptg))/100
				 totalamountppn = dtl_totalamount + totalppn
				'' total_all = totalppn+totalamountppn

			  q= " Update stocktransferorderhdr set totalamount=replace('" & totalamountppn & "',',','')" & _
				 ", LASTUPDATE = sysdate " & _
				 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
				 " where stono = '" & stono & "' and fcBA = '" & businessunit & "'"
				connStr.execute(q)
			   if err.number  = 0 then
				  connStr.CommitTrans
			   end if
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans
	  Dim rsPO, rsPurchases,rsQty
	  
		Set rsPO = Server.CreateObject("ADODB.Recordset")
			  rsPO.Open "Select poCode, stono from stocktransferorderhdr where stono = '"& stono &"' and poCode='" & fcCode & "' AND FCBA = '" & businessunit & "'" , connstr
			  IF rsPO.eof then	   
				 ' Modified by NAF
				 ' To handle first time Input for Purchase Order

				 q = " INSERT INTO stocktransferorderhdr ( company, stono, poCode, stodate, rentby, transportercode, transportername, transporteraddress, termofpayment, dp, ppn, remarks, totalamount, rowstateid, fcentry, fcedit, fcip, fcba, lastupdate, lasttime,ppnptg,ppnrp ) " & _
				     " VALUES (  '" & fccompanycode & "'" & _
				     "         , '" & stono & "'" & _
				     "         , '" & fcCode & "'" & _
					 "         , TO_DATE('" & STODATE & "', 'mm/dd/yyyy')" & _
					 "         , '" & rentby & "'" & _
					 "         , '" & transportercode & "'" & _
					 "         , '" & transportername & "'" & _
					 "         , '" & transporteraddress & "'" & _
					 "         , '" & termofpayment & "'" & _
					 "         , replace('" & dp & "',',','')" & _
					 "         , replace('" & ppn & "',',','')" & _
					 "         , '" & remarks & "'" & _
					 "         , replace('" & totalamount & "',',','')" & _
					 "         , 'Approved'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					 "         , '" & businessunit & "'" & _
					 "         , SYSDATE" & _
					 "         , TO_CHAR(SYSDATE,'hh24:mi') " & _   
					 "         , replace('" & ppnptg & "',',','')" & _
					 "         , replace('" & ppnrp & "',',',''))"
				 ' response.write q
				 connStr.execute(q)
			
			  End if
			  rsPO.Close
				' Added by NAF, 2012-12-01
				 strLineMax = " select nvl(max(STOLINENO),0) STOLINE from stocktransferorderdtl where fcBA='" & businessunit & "'"

				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					tempMaxLineNo = cdbl(rsstrLineMax("STOLINE")) + 1
				end if
							
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 
				' ' End of Added by NAF, 2012-12-01
				 
					' if totalPRData <> 0 then 
						' if ApprovedQty - OrderedQty - Quantity >= 0 then 
							amount = 0
						 	amount1 = cdbl(to_nol(quantity)) * cdbl(to_nol(unitprice))
							amount2 = amount1 * cdbl(to_nol(ppn))/100
							Amount = amount1 - amount2
							' amount4 = amount3 * cdbl(to_nol(ppn))/100
							' amount5 = amount3 + amount4
							' AMOUNT = amount + amount3
							' AMOUNTCHARGE = amount3 + cdbl(to_nol(othercharge))
							
							q="Insert into  stocktransferorderdtl (STONO, LINENOKEY, POCODE, ITEMCODE, STOCKNAME, QUANTITYORDER, UNITPRICE, AMOUNT, DO_QUANTITY, DESCRIPTION, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, rowstateid, STOLINENO ) " & _
							 "values ('" & STONO & "','" & PONUMBER & "','" & fcCode & "','"& ITEMCODE &"','"& STOCKNAME &"','"& QUANTITY &"',replace('" & UNITPRICE & "',',','')" & _
							 ",replace('"& AMOUNT &"',',','')" & _
							 ",'" & DOQUANTITY & "','" & OTHERITEMDESC & "','" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
							 Request.ServerVariables("REMOTE_ADDR") &"','" & businessunit & "', sysdate, TO_CHAR(SYSDATE,'hh24:mi'),'','"& tempMaxLineNo &"')"
							 ' response.write q
							connStr.execute(q)

							'Untuk menghhitung jumlah total di header
							dtl_totalamount = 0
							dtl_charge = 0
							totalamount = 0
							set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
							q="select *  from stocktransferorderdtl " & _
							  " where stono='"& stono &"' and " & _
							  " fcBA='" & businessunit & "'"
							 ' response.write q
							rsTempPurchaseRequest.Open q, connstr
							if Not rsTempPurchaseRequest.Eof then
								while Not rsTempPurchaseRequest.Eof 
									dtl_totalamount = dtl_totalamount + cdbl(rsTempPurchaseRequest("amount")) 
									rsTempPurchaseRequest.movenext
								wend 

								
								totalppn = dtl_totalamount  * cdbl(to_nol(ppnptg))/100
								totalamountppn = dtl_totalamount + totalppn
								'total_all = totalppn+totalamountppn
								
							End if
							
							rsTempPurchaseRequest.Close
							set rsTempPurchaseRequest = Nothing

						  q= " Update stocktransferorderhdr set totalamount=replace('" & totalamountppn & "',',','')" & _
							 ", LASTUPDATE = sysdate " & _
							 ", LASTTIME = To_Char(sysdate,'HH24:MI'),ppnrp=replace('" & totalppn & "',',','') "& _
							 " where  stono ='" & stono & "' and fcBA='" & businessunit & "'"
							connStr.execute(q)
							if err.number  = 0 then
								 connStr.CommitTrans
							end if
	      cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	   connstr.BeginTrans
	   		amount = 0
		 	amount1 = cdbl(to_nol(quantity)) * cdbl(to_nol(unitprice))
			amount2 = amount1 * cdbl(to_nol(ppn))/100
			Amount = amount1 - amount2

	   		q =	" Update stocktransferorderdtl" & _
				" set description = '" & description & "'" & _
				"     , amount = replace('" & amount & "',',','')" & _
				"     , unitprice = replace('" & unitprice & "',',','')" & _
				"     , fcentry = '" & Session("CodeName") & "'" & _
				"     , fcedit = '" & Session("CodeName") & "'" & _
				" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"     , lastupdate = sysdate " & _
				"     , lasttime = to_char(sysdate,'hh24:mi') " & _
				" 		where STONO = '" & STONO & "'" & _
				"       and LINENOKEY = '" & LINENOKEY & "'" & _
				"       and fcBA = '" & businessunit & "'"
 			' response.write q
			connStr.execute(q)
			'Untuk menghhitung jumlah total di header
			dtl_totalamount = 0
			dtl_charge = 0
			totalamount = 0
			set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			q="select *  from stocktransferorderdtl " & _
			  " where stono='"& stono &"' and " & _
			  " fcBA='" & businessunit & "'"
			 ' response.write q
			rsTempPurchaseRequest.Open q, connstr
			if Not rsTempPurchaseRequest.Eof then
				while Not rsTempPurchaseRequest.Eof 
					dtl_totalamount = dtl_totalamount + cdbl(rsTempPurchaseRequest("amount")) 
					rsTempPurchaseRequest.movenext
				wend 
				totalppn = dtl_totalamount  * cdbl(to_nol(ppnptg))/100
				totalamountppn = dtl_totalamount + totalppn
			End if
			
			rsTempPurchaseRequest.Close
			set rsTempPurchaseRequest = Nothing

		  q= " Update stocktransferorderhdr set totalamount=replace('" & totalamountppn & "',',','')" & _
			 ", LASTUPDATE = sysdate " & _
			 ", LASTTIME = To_Char(sysdate,'HH24:MI'),ppnrp=replace('" & totalppn & "',',','') "& _
			 " where  stono ='" & stono & "' and fcBA='" & businessunit & "'"
			connStr.execute(q)
			if err.number  = 0 then
				 connStr.CommitTrans
			end if
		' if err.number  = 0 then
		'    connStr.CommitTrans
		' end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Confirmed" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'Confirm' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "RePlanned" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'RePlan' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "Ordered" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'Order' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
    elseif Request("txtJenis")= "UnRelease"	then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.UnReleasePO('" & fcCode & "' ); END; "
	   ' response.Write strQueryStatus
	   connStr.execute(strQueryStatus)	 
	   cmdView	= "True"
	elseif Request("txtJenis")= "CompletedDetails" then
		connstr.BeginTrans
		strQueryStatus = " BEGIN" & _
						 " PURCHASEORDERDETAILS_API.NextState('" & businessunit & "'" & _
	                     "                                    , '" & fcCode & "'" & _
					     "                                    , '" & txtLineNoKey & "'" & _
	                     "                                    , 'Complete'" & _
	                     "                                    );" & _
	                     " END; "
		'response.Write strQueryStatus
		connStr.execute(strQueryStatus)	 

		cmdView	= "True"
	
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

' Untuk menampilkan value header jika ada data.
 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q = " SELECT * FROM stocktransferorderhdr sto LEFT JOIN businessunit bu ON bu.fccode = sto.fcba " & _
     " WHERE sto.stono = '" & stono & "'"
 ' response.write q
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    STODATE = rsPurchase("STODATE")
	fccompanycode = rsPurchase("fccompanycode")
	businessunit = rsPurchase("fcba")
	fccode = rsPurchase("pocode")
	rentby = rsPurchase("rentby")
	Remarks = rsPurchase("Remarks")
	transportercode = rsPurchase("transportercode")
	transportername = rsPurchase("transportername")
	transporteraddress = rsPurchase("transporteraddress")
	termofpayment = rsPurchase("termofpayment")
	dp = rsPurchase("dp")
	ppn = rsPurchase("ppn")
	RowStateId = rsPurchase("RowStateId")
	RowStateHdr = rsPurchase("RowStateId")
	totalamount = rsPurchase("totalamount")
	selectedtransportercode = rsPurchase("transportercode")
	selectedfcba = rsPurchase("fcba")
	selectedrentby = rsPurchase("rentby")
	selectedpo = rsPurchase("pocode")
	ppnptg = rsPurchase("ppnptg")
	ppnrp = rsPurchase("ppnrp")
 End if
 

' amount = 0
' strSuppyingDeclaration	=	" SELECT NVL(total,0) amount FROM stocktransferorderhdr" & _
' 							" WHERE stono = '" & stono & "' and fcba = '"& businessunit &"'" 
' response.Write strSuppyingDeclaration
' set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
' rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
' if Not rsSupplyingDeclaration.Eof then
' 	while Not rsSupplyingDeclaration.Eof
' 		amount = amount + cdbl(to_nol(rsSupplyingDeclaration("amount")))
' 		rsSupplyingDeclaration.movenext
' 	wend
' End if
' rsSupplyingDeclaration.Close : Set rsSupplyingDeclaration = Nothing
' ' response.Write amount
' totalppn = amount * cdbl(to_nol(pph))/100

 dim rsFrez
 set rsFrez = server.CreateObject("Adodb.recordset")
	q=" select distinct frezee  from Purchaseorderdetails " & _
	  " where PoCode='" & fcCode & "' and " & _
	  " fcBA='" & businessunit & "'"
 rsFrez.Open q, connstr
 if Not rsFrez.Eof then
	IsFrezee  = rsFrez("frezee")
 else
    IsFrezee ="0" 
 End if
 rsFrez.Close
 set rsFrez = Nothing
 
 val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & businessunit & "'" & _
	        "       and to_date ('" & STODATE & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & STODATE & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
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

has_do = 0

oSQLQuery = " SELECT COUNT(1) total" & _
            " FROM deliveryorderpurchasedtl" & _
	        " WHERE fcba = '" & businessunit & "'" & _
	        "       AND stono = '"&stono&"'"
'response.Write oSQLQuery
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	has_do = cdbl(to_nol(oRecordSet("total")))
Else
	has_do = 0
End if
oRecordSet.Close : Set oRecordSet = Nothing
 
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<div class='error'>" & err.description & "</div>"
End if
 '-------------ISI Payment Form
%>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-exchange" aria-hidden="true"></i>Header Stock Transfer Order</div><% end if %>
 <form name="frmEntry" method="post" action="stocktransferorder.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">STO No</td>
			<%
			if stono = "" then
				q = "select MAX(cn.PREFIX) || LPAD(NVL(MAX(regexp_substr(sto.STONO, '[^/]+', 1, 4)),0000)+0001,4,0) DOCNUMBER " & _
					" from " & _
					" controlnumber cn " & _
					" left join stocktransferorderhdr sto on cn.fcba = sto.fcba and cn.accountingyear = to_number(to_char(sto.STODATE,'yyyy')) " & _
					" where cn.fcba ='" & session("fcBA") & "' " & _
					" and cn.LU_NAME ='STOCKTRANSFER' "&_
					" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& STODATE &"','mm/dd/yyyy'),'yyyy')) " & _
					" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& STODATE &"','mm/dd/yyyy'),'mm')) "
				'Response.write q
				set rsSTONo = server.CreateObject("Adodb.recordset")
				rsSTONo.Open q, connstr
				if Not rsSTONo.Eof then
					stono = rsSTONo("DOCNUMBER")
				End if
			End if
			%>
			<td align="left">
				<input type="text" maxlength="50" name="stono" size="30" value="<%=stono%>" autocomplete="off">
				<font color="#FF0000">*</font>
			</td>
			<td align="left" style="width:20%;">Transporter Code</td>
				<td align="left">
					<div id="transportercodelist"></div>
					<font color="#FF0000">*</font>
				</td>
				<input type="Hidden" id="selectedtransportercode" name="selectedtransportercode" value="<%=selectedtransportercode%>">
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<input class="datepicker" type="text" onchange="javascript:reloadStoCode();" maxlength="10" name="STODATE" size="10" value="<%=STODATE%>">
			</td>
			
			<td align="left">Transporter</td>
			<td align="left">
					<textarea id="transportername"  name="transportername" style="background:#eee" readonly><%=transportername%></textarea>
			</td>
		</tr>
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

			<td align="left">Transporter Address</td>
				<td align="left">
						<textarea id="transporteraddress"  name="transporteraddress" style="background:#eee" readonly><%=transporteraddress%></textarea>
				</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
				<td align="left">
					<div id="fcbalist"></div>
					<font color="#FF0000">*</font>
				</td>
				<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">

			<td align="left">PPN (%)</td>
			<td align="left">
				<input type="text" size="5" name="ppnptg" value="<%=ppnptg%>" onkeypress="return harusangka(event);">&nbsp;<input type="text" maxlength="50" name="ppnrp" size="18" value="<%=ppnrp%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
			</td>
		</tr>
		<tr>
			<td align="left">Purchase Order No</td>
				<td align="left">
					<div id="polist"></div>
					<font color="#FF0000">*</font>
				</td>
				<input type="Hidden" id="selectedpo" name="selectedpo" value="<%=selectedpo%>">
			<td align="left">DP</td>
			<td align="left">
				<input type="text" maxlength="50" name="dp" size="30" value="<%=dp%>" style="background:#eee" readonly data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
			</td>

		</tr>
		<tr>
			<td align="left">Rent By</td>
				<td align="left">
					<div id="rentbylist"></div>
					<font color="#FF0000">*</font>
				</td>
				<input type="Hidden" id="selectedrentby" name="selectedrentby" value="<%=selectedrentby%>">

			<td align="left">PPh(%)</td>
			<td align="left">
				<input type="text" size="5" name="ppn" value="<%=ppn%>" readonly style="background:#eee" onkeypress="return harusangka(event);">
			</td>
		</tr>
		<tr>
			<td align="left">Term Of Payment</td>
			<td align="left">
				<select name="termofpayment" value="<%=termofpayment%>" class="chosen-select" style="width: 50%">
					<option value="<%=termofpayment%>"></option>
					<option value="TOP7" <% If cstr(""&termofpayment) = "TOP7" Then %> Selected <%End if%>>TOP7</option> 
					<option value="TOP14" <% If cstr(""&termofpayment) = "TOP14" Then %> Selected <%End if%>>TOP14</option> 
					<option value="TOP30" <% If cstr(""&termofpayment) = "TOP30" Then %> Selected <%End if%>>TOP30</option> 
				</select>
			</td>
			<td align="left">Total Amount</td>
			<td align="left">
				<input type="text" maxlength="50" name="totalamount" size="30" value="<%=totalamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
			</td>
			
		</tr>

		<tr>
			<td align="left">Remarks</td>
			<td align="left">
				<textarea name="Remarks"><%=Remarks%></textarea>
			</td>
		</tr>



		</tr>
		<tr><td colspan="4"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		 <!-- Validasi Period Control taro sini-->	
			<%if has_do = 0 then%>
			<a href="stocktransferorder.asp?stono=<%=stono%>&fcCode=<%=fcCode%>&STODATE=<%=STODATE%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&txtJenis=Header"onClick="return confirm('Are you sure to delete');">
				<img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0">
			</a>
			<%Else%>
			<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
			<%end if%>
		 
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" AND LEN(TRIM(fcCode)) > 0 AND LEN(TRIM(stono)) > 0 AND LEN(TRIM(fccompanycode)) > 0 AND LEN(TRIM(businessunit)) > 0 AND LEN(TRIM(rentby)) > 0  Then
	Dim rsPurchaseOrder,intRecordsPerPage,cColor

	q = " select *  from stocktransferorderdtl" & _
		" where stono = '" & stono & "'" & _
	    " and fcBA = '" & businessunit & "'"
	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.CursorType = 3
	rsPurchaseOrder.CursorLocation = 3
	rsPurchaseOrder.Open q, connstr
%>
   


   
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="stocktransferorder.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
   		<tr>
			<td align="left">PO. Line No </td>
			<td align="left">
			<%IF txtJenis="Edit" Then%>
				 <input type="text" size="15" name="ponumber" value="<%=linenokey%>" id="linenokey" readonly="True" style="background:#eee">&nbsp;<font color="#FF0000">*
			<%else%>
				 <select id="linenokey" name="ponumber" onchange="javascript:getStockCodeSTO();" onblur="javascript:getStockCodeSTO();">
					<option value=""></option> 
						<% 	Dim rsComboS
							q = "  SELECT linenokey" & _
								"    FROM (SELECT pod.linenokey" & _
								"            FROM purchaseorderdetails pod" & _
								"                 INNER JOIN purchaseorder po" & _
								"                    ON (po.pocode = pod.pocode AND po.fcba = pod.fcba)" & _
								"           WHERE po.fcba = '"& businessunit &"' AND po.pocode = '"& fcCode &"'" & _
								"          MINUS" & _
								"          SELECT std.linenokey" & _
								"            FROM stocktransferorderdtl std" & _
								"                 INNER JOIN stocktransferorderhdr sth" & _
								"                    ON (std.stono = sth.stono AND std.fcba = sth.fcba)" & _
								"           WHERE std.fcba = '"& businessunit &"' AND std.pocode = '"& fcCode &"')" & _
								" ORDER BY linenokey"  

							' response.write q
							Set rsComboS = Server.CreateObject("ADODB.Recordset")
							rsComboS.CursorType = 3
							rsComboS.CursorLocation = 3
							rsComboS.Open q, connstr
							if not rsComboS.eof then%>
							<%
							Do While Not rsComboS.eof
							%>
								<option value="<%=rsComboS("linenokey")%>" 
									<%if cstr(""&linenokey)=cstr(""&rsComboS("linenokey")) then %> selected 
									<%End If%>><%=rsComboS("linenokey")%>
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
			<td align="left">Ordered PO Quantity</td>
			<td align="left"><input type="text" size="15" id="quantity" name="quantity" value="<%=quantity%>" style="background:#eee" readonly onkeyup="hitungrumus();" >
			</td>
		</tr>
		<tr>
			<td align="left">Unit Price</td>
			<td align="left"><input type="text" size="15" id="unitprice" name="unitprice" value="<%=unitprice%>" onkeypress="return validateFloatKeyPress(this,event);" onkeyup="harusangka();hitungrumus();">&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		<tr>
			<td align="left">Amount</td>
			<td align="left"><input type="text" size="15" id="amount" name="AMOUNT" value="<%=AMOUNT%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  style="background:#eee;" readonly >
			</td>
		</tr>
		<tr>
			<td align="left">DO Quantity</td>
			<td align="left"><input type="text" size="15" name="doquantity" value="<%=doquantity%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee;" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Description</td>
			<td align="left">
				<textarea  name="OTHERITEMDESC"  cols="50" rows="3"><%=OTHERITEMDESC%></textarea>
			</td>
		</tr>
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
	      <input type="hidden" name="stono" value="<%=stono%>">
		  <input type="hidden" name="STODATE" value="<%=STODATE%>">
		  <input type="hidden" name="transportercode" value="<%=transportercode%>">
		  <input type="hidden" name="transportername" value="<%=transportername%>">
		  <input type="hidden" name="transporteraddresss" value="<%=transporteraddresss%>">
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="LineNokey" value="<%=LINENOKEY%>">
		  <input type="hidden" name="use_ppn" value="<%=use_ppn%>">
		  <input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		  <input type="hidden" name="businessunit" value="<%=businessunit%>">
		  <input type="hidden" name="termofpayment" value="<%=termofpayment%>">
		  <input type="hidden" name="dp" value="<%=dp%>">
		  <input type="hidden" name="ppn" value="<%=ppn%>">
		  <input type="hidden" name="rentby" value="<%=rentby%>">
		  <input type="hidden" name="ppnptg" value="<%=ppnptg%>">
		  <input type="hidden" name="ppnrp" value="<%=ppnrp%>">
		  <input type="hidden" name="totalamount" value="<%=totalamount%>">
		  <input type="hidden" name="RowStatehdr" value="<%=RowStatehdr%>">
		  <input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		  <input type="hidden" name="selectedrentby" value="<%=selectedrentby%>">
		  <input type="hidden" name="selectedpo" value="<%=selectedpo%>">
		  <input type="hidden" name="selectedtransportercode" value="<%=selectedtransportercode%>">
		  <input Type="Hidden" Value="True" Name="cmdView">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		   <!-- Validasi Period Control taro sini-->		  
			  <%IF Request("txtJenis")="Edit" then%>
				   <input type="hidden" name="txtJenis" value="Edit">
				   <input Type="Submit"  name="Add" value="Update">
			  <%Else  %>
				   <input type="hidden" name="txtJenis" value="New">
				   <input Type="Submit" name="Add" value="Add">
			  <%End if%>
			  <input Type="Reset" name="Reset" value="Reset">
			  <input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)">
		   
		  </td>
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="middle" colspan="2">&nbsp;</TD>
	<TD align="left">STO NO</TD>
	<TD align="left">PO CODE</TD>
	<TD align="left">ITEM CODE</TD>
	<TD align="left">ITEM NAME</TD>
	<TD align="left">UNIT PRICE</TD>
	<TD align="left">QUANTITY</TD>
	<TD align="left">AMOUNT</TD>
	<TD align="left">DESCRIPTION</TD>
	</tr>
	</thead>
	<tbody>   
   <%
   
     Set rsPO = Server.CreateObject("ADODB.Recordset")
	 rsPO.Open " select purchaseorder_api.GetPoState('" & businessunit & "' ,'" & fcCode & "') HdrState from dual" , connstr
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
			<%if val_IsTrue = "Y" then%>
				<%if has_do = 0 then%>
				<td colspan="2">
					<a href="stocktransferorder.asp?stono=<%=rsPurchaseOrder("stono")%>&pocode=<%=rsPurchaseOrder("pocode")%>&stolineno=<%=rsPurchaseOrder("stolineno")%>&businessunit=<%=rsPurchaseOrder("fcba")%>&linenokey=<%=rsPurchaseOrder("linenokey")%>&itemcode=<%=rsPurchaseOrder("itemcode")%>&stockname=<%=rsPurchaseOrder("stockname")%>&quantity=<%=rsPurchaseOrder("quantityorder")%>&unitprice=<%=rsPurchaseOrder("unitprice")%>&amount=<%=rsPurchaseOrder("amount")%>&description=<%=rsPurchaseOrder("description")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
				
				   <a href="stocktransferorder.asp?stono=<%=rsPurchaseOrder("stono")%>&pocode=<%=rsPurchaseOrder("pocode")%>&stolineno=<%=rsPurchaseOrder("stolineno")%>&businessunit=<%=rsPurchaseOrder("fcba")%>&linenokey=<%=rsPurchaseOrder("linenokey")%>&itemcode=<%=rsPurchaseOrder("itemcode")%>&stockname=<%=rsPurchaseOrder("stockname")%>&quantity=<%=rsPurchaseOrder("quantityorder")%>&unitprice=<%=rsPurchaseOrder("unitprice")%>&amount=<%=rsPurchaseOrder("amount")%>&description=<%=rsPurchaseOrder("description")%>&txtJenis=Details">
				   <img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');">
				   </a>
				</td>
				<% Else %>
					<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>		
				<% End If %>
			<% Else %>	
               	<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>		
            <% End If %>	
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
				<td align="left" title="<%=rsPurchaseOrder("unitprice").Name%>" valign="top">
					<%=rsPurchaseOrder("unitprice")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("QUANTITYORDER").Name%>" valign="top">
					<%=rsPurchaseOrder("QUANTITYORDER")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("amount").Name%>" valign="top">
					<%=rsPurchaseOrder("amount")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("description").Name%>" valign="top">
					<%=rsPurchaseOrder("description")%>
				</td>
			</tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("<tr><td colspan=""11"" class=""fbp-doc-empty"">Data is empty</td></tr>")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
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
		getlistrentby();
		getlisttransportercode();
		hitungrumus();
		getselectedfcba();
		getselectedpo();
		getselectedrentby();
		getselectedtransportercode();
		
	});
	</script>
</body>
</html>
