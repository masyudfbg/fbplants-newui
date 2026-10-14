<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Invoice TBS</title>
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
<script>
	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}
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

function hitungrumus() {
		var skg = $("#kg").val();
		var srate = $("#rate").val(); 
		var newskg = skg.replace(/,/g, '');
		var newsrate = srate.replace(/,/g, '');
		
		//Untuk menghitung Total Janjang
		totalamount = (newskg * newsrate).toFixed(2); 
		if (!isNaN(totalamount)) {
    	 	$("#amount").val(totalamount);
    	}
    	else{
    		$("#amount").val();
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
	doAJAXCall('ajax/getlistfcbainvoicetbs.asp', 'POST', PostStr, showlistfcba);
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
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getlistsupplier = function () {
	// build up the post string when passing variables to the server side page
	var selectedsupplier = document.getElementById("selectedsupplier").value;
	var selectedfcba = document.getElementById("selectedfcba").value; 
	var PostStr = 'fcba='+selectedfcba+'&fccode='+selectedsupplier;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getSupplierTBS.asp', 'POST', PostStr, showSupplier);
	
}

// The function for handling the response from the server
var showSupplier = function (oXML) { 
  
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("supplierlist").innerHTML = "";
	document.getElementById("supplierlist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
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

// launched from button click 
var getStockCode = function () {
	// alert("1");
	// build up the post string when passing variables to the server side page
	// var PrCode = document.getElementById("prcode").value;
	// var PRLINEID = document.getElementById("prlineid").value;
	var PostStr = 'PrCode='+document.getElementById("prcode").value+'&PRLINEID='+document.getElementById("prlineid").value;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getStockCode.asp', 'POST', PostStr, showStockCodeResponse);
	
}

// The function for handling the response from the server
var showStockCodeResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.itemcode);
    // update the Div to show the result from the server
    document.getElementById("itemcode").value = s.itemcode;
	document.getElementById("stockname").value = s.stockname;
	document.getElementById("requestedqty").value = s.requestedqty;
	
    
}

</script>
<!-- <script language="JavaScript">
function getFcba(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("businessunit").innerHTML="";
	  document.getElementById("purchaseby").innerHTML="";
	  document.getElementById("deliverydestination").innerHTML="";
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
		document.getElementById("purchaseby").innerHTML=xmlhttp.responseText;
		document.getElementById("deliverydestination").innerHTML=xmlhttp.responseText;
		// alert(businessunit);
		}
	  }
	xmlhttp.open("GET","ajax/getFcbaPO.asp?fccompanycode="+str+"",true);
	xmlhttp.send();
}

function getSupp(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("suppliername").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {
	  // code for IE7+, Firefox, Chrome, Opera, Safari
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
		document.getElementById("suppliername").innerHTML=xmlhttp.responseText;
		}
	  }
	xmlhttp.open("GET","ajax/getSupplierFCBA.asp?fcba="+str+"",true);
	xmlhttp.send();
}
					
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 -->
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
 
 Dim xCode,fcCode,cmdView,PoCode,PoDate,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
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
 if Request("PODATE")<>"" then
   PODATE   =Request("PODATE")
 else
   PODATE   =textDate(date)
 end if

If Request("fakturdate")<>"" then
	fakturdate   =Request("fakturdate")
End If

If Request("transdate")<>"" then
   transdate   =Request("transdate")
Else
   transdate   =textDate(date)
End if

If Request("valid_from")<>"" then
	valid_from   =Request("valid_from")
Else
	valid_from =textDate(date)
End If

If Request("valid_until")<>"" then
	valid_until   =Request("valid_until")
Else
	valid_until =textDate(date)
End If 
  PODATE = FormatDateTime(PODATE)

 
 xCode = Request("xCode")
 IsInBridge = Request("IsInBridge")
 Remarks = Request("Remarks")
 
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
 use_ppn = Request("use_ppn")
 fccompanycode = Request("fccompanycode")
 fcba = Request("fcba")
 invoice_by =Request("invoice_by")
 suppliercode = Request("suppliercode")
 selectedfcba = Request("selectedfcba")
 selectedsupplier = Request("selectedsupplier")
 selectedinvoice = Request("selectedinvoice")
 purchaseby = Request("purchaseby")
 suppliername = Request("suppliername")
 supplieraddress = Request("supplieraddress")
 termofpayment = Request("termofpayment")
 SHIPPINGPOINT = Request("SHIPPINGPOINT")
 dp = Request("dp")
 ppn = cdbl(to_nol(Request("ppn")))
 pph = cdbl(to_nol(Request("pph")))
 invoiceno = Request("invoiceno")
 StockName = Request("StockName")
 unitprice = Request("unitprice")
 disc = Request("disc")
 othercost = cdbl(to_nol(Request("othercost")))
 totalamount = cdbl(to_nol(Request("totalamount")))
 If request ("JobCode") = "" Then
	Jobcode = "1114500"
 Else
	JobCode = Request("JobCode")
 End If

 If request ("targettype") = "" Then
	targettype = "CB"
 Else
	targettype = Request("targettype")
 End If 
 
 If request ("targetcode") = "" Then
	targetcode = "BS"
 Else
	targetcode = Request("targetcode")
 End If 
 grandtotal = cdbl(to_nol(Request("grandtotal")))
 amountcharge = cdbl(to_nol(Request("amountcharge")))
 totalppn = cdbl(to_nol(Request("totalppn")))
 totalpph = cdbl(to_nol(Request("totalpph")))
 fakturno = request ("fakturno")
 description = request ("description")
 lineno = request ("lineno")
 spolineno = request ("spolineno")
 invoice_type = request ("invoice_type")
 kg = request ("kg")
 rate = request ("rate") 
 quantity_adjustment = request ("quantity_adjustment") 
 If request ("grading") = "" Then
 	grading = 3
 Else
 	grading = request ("grading")
 End If
 
private function cek_trigger(tt_, tc_, jobcode_, invoice_by_, date_, sourcetype_)
	result_ = ""
	q = " SELECT f_cektarget_bydate ('"&tt_&"', '"&tc_&"', '"&jobcode_&"', '"&invoice_by_&"', TO_DATE('"&date_&"', 'mm/dd/yyyy'), '"&sourcetype_&"') outputstr FROM DUAL"
	' response.write q
	set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
	rsTempPurchaseRequest.Open q, connstr
	if Not rsTempPurchaseRequest.Eof then
		result_ = rsTempPurchaseRequest("outputstr")
	else
		result_ = "No Trigger Setup!"
	End if
	rsTempPurchaseRequest.Close : set rsTempPurchaseRequest = Nothing
	
	cek_trigger = result_
end function

if "" & use_ppn = "" then use_ppn = "N" 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "Ordered" or Request("txtJenis")= "CompletedDetails"  or  Request("txtJenis")= "CanceledDetails"  or  Request("txtJenis")= "UnRelease" Then
	
	 
	
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
       connstr.BeginTrans
		
				
		q="Delete from invoiceservicedtl where invoiceno='" & invoiceno & "' and fcBA='" & fcba & "'"
		connStr.execute(q)

		q="Delete from invoiceservicehdr where invoiceno='" & invoiceno & "' and fcBA='" & fcba & "'"
		connStr.execute(q)
			
		if err.number  = 0 then
			connStr.CommitTrans
		end if
	   cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans 
		'Untuk menghhitung jumlah total di header
		dtl_totalamount = 0
		dtl_amount = 0
		dtl_charge = 0
		set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
		q="select *  from invoiceservicedtl " & _
		  " where invoiceno='" & invoiceno & "' and " & _
		  " fcBA='" & fcba & "'"
		  ' response.write q
		rsTempPurchaseRequest.Open q, connstr
		if Not rsTempPurchaseRequest.Eof then
			while Not rsTempPurchaseRequest.Eof 
				dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("totalamount")))
				dtl_amount = dtl_amount + cdbl(to_nol(rsTempPurchaseRequest("amount")))
				dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercost")))
				rsTempPurchaseRequest.movenext
			wend 						
		End if					
		rsTempPurchaseRequest.Close : set rsTempPurchaseRequest = Nothing

		totalppn = dtl_amount *cdbl(to_nol(ppn))/100
		totalpph = dtl_amount *cdbl(to_nol(pph))/100
		grandtotal = (dtl_amount+totalppn-totalpph)+dtl_charge

	  q= " Update invoiceservicehdr set totalppn = replace('" & totalppn & "',',',''), totalpph = replace('" & totalpph & "',',','')" & _
		 ", grandtotal = replace('" & grandtotal & "',',','')" & _
		 ", LASTUPDATE = sysdate " & _
		 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
		 " where  invoiceno='" & invoiceno & "' and fcBA='" & fcba & "'"
		'response.write q
		connStr.execute(q)
		
		 q="Delete from invoiceservicedtl where invoiceno='" & invoiceno & "' and fcBA='" & fcba & "' and lineno = '"&lineno&"'"
		   ' response.write q
		   connStr.execute(q)

	   if err.number  = 0 then
		  connStr.CommitTrans
	   end if
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans 
		
			
				max_amount = 0 : max_charge = 0 : max_amountcharge = 0
			
				Dim rsPO, rsPurchases,rsQty
			  
				Set rsPO = Server.CreateObject("ADODB.Recordset")
				rsPO.Open "Select poCode from invoiceservicehdr where invoiceno='" & invoiceno & "' AND FCBA = '" & fcba & "'" , connstr
				IF rsPO.eof then	   
					' Modified by NAF
					' To handle first time Input for Purchase Order

					q =  " INSERT INTO invoiceservicehdr ( invoiceno, invoice_type, pocode, transdate, ppn, pph, company, suppliername, no_faktur, fakturdate, grading, totalppn, totalpph, grandtotal, fcentry, fcedit, fcip, fcba, invoice_by, valid_from, valid_until, lastupdate, lasttime ) " & _
						 " VALUES ( '" & invoiceno & "'" & _
						 "         , 'TBS'" & _
						 "         , '" & fcCode & "'" & _
						 "         , TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
						 "         , '" & ppn & "'" & _
						 "         , '" & pph & "'" & _
						 "         , '" & fccompanycode & "'" & _
						 "         , '" & suppliercode & "'" & _
						 "         , '" & fakturno & "'" & _
						 "         , TO_DATE('" & fakturdate & "', 'mm/dd/yyyy')" & _					 
						 " 		   , replace('" & grading & "',',','')" & _
						 " 		   , replace('" & totalppn & "',',','')" & _
						 " 		   , replace('" & totalpph & "',',','')" & _
						 " 		   , replace('" & grandtotal & "',',','')" & _
						 "         , '" & Session("CodeName") & "'" & _
						 "         , '" & Session("CodeName") & "'" & _
						 "         , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						 "         , '" & fcba & "'" & _
						 "         , '" & invoice_by & "'" & _
						 "         , TO_DATE('" & valid_from & "', 'mm/dd/yyyy')" & _
						 "         , TO_DATE('" & valid_until & "', 'mm/dd/yyyy')" & _
						 "         , SYSDATE" & _
						 "         , TO_CHAR(SYSDATE,'hh24:mi'))"
						' response.write q
						connStr.execute(q)
				End if
				rsPO.Close
				
				in_data_amount = 0 : in_data_charge = 0 : in_data_amountcharge = 0
				strLineMax = 	" SELECT NVL(SUM(amount),0) amount, NVL(SUM(othercost),0) othercharge, NVL(SUM(totalamount),0) amountcharge" & _
								" FROM invoiceservicedtl" & _
								" WHERE pocode = '"&fcCode&"'" & _
								"       AND fcba = '" & fcba & "'" & _
								"       AND polinenokey = '" & spolineno & "'"
				
				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					in_data_amount = cdbl(to_nol(rsstrLineMax("amount")))
					in_data_charge = cdbl(to_nol(rsstrLineMax("othercharge")))
					in_data_amountcharge = cdbl(to_nol(rsstrLineMax("amountcharge")))
				else
					in_data_amount = 0 : in_data_charge = 0 : in_data_amountcharge = 0
				end if							
				rsstrLineMax.Close : set rsstrLineMax = Nothing 
					
					' TOTALAMOUNT = cdbl(to_nol(replace(AMOUNT,",",""))) + cdbl(to_nol(replace(OTHERCOST,",","")))
							
					For n=1 to (cdbl(request("no"))-1) 
						if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
							If ( len(trim(""&request("jobcode_"&TRIM(n)))) > 0 or len(trim(""&request("targettype_"&TRIM(n)))) > 0 or len(trim(""&request("targetcode_"&TRIM(n)))) > 0 )then
								If cek_trigger(request("targettype_"&TRIM(n)), request("targetcode_"&TRIM(n)), request("jobcode_"&TRIM(n)), invoice_by, transdate, "PR") = "Success" then
									If cdbl(to_nol(request("quantity_adjustment_"&TRIM(n))))  = 0 Then 
									    
										quantity1 =  cdbl(to_nol(request("quantity_"&TRIM(n))))
										amount = ( cdbl(to_nol(request("quantity_"&TRIM(n)))) - (cdbl(to_nol(request("quantity_"&TRIM(n)))) * (cdbl(to_nol(request("grading")))/ 100 )) ) * cdbl(to_nol(request("rate_"&TRIM(n))))
										quantity = ( cdbl(to_nol(request("quantity_"&TRIM(n)))) - (cdbl(to_nol(request("quantity_"&TRIM(n)))) * (cdbl(to_nol(request("grading")))/ 100 )) )
										quantity_adj = 0
									Else
									    quantity1 = cdbl(to_nol(request("quantity_"&TRIM(n))))
										quantity_adj = round((cdbl(to_nol(request("quantity_"&TRIM(n)))) * (cdbl(to_nol(request("grading")))/ 100 )),0)
										quantity = quantity1 - quantity_adj
										amount = quantity  * cdbl(to_nol(request("rate_"&TRIM(n))))

									End If 
										tempMaxLineNo = 0
										strLineMax = " SELECT NVL(MAX(lineno),0) maxlineno FROM invoiceservicedtl WHERE fcBA='" & fcba & "' AND invoiceno = '"&INVOICENO&"'"
										Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
										rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
										
										if not rsstrLineMax.eof then
											tempMaxLineNo = cdbl(to_nol(rsstrLineMax("maxlineno")))
										end if
													
										rsstrLineMax.Close : set rsstrLineMax = Nothing
												
										tempMaxLineNo = tempMaxLineNo +1
										
										q = " Insert into  invoiceservicedtl (invoiceno, jobcode, targettype, targetcode, description, kg, rate, amount, othercost, totalamount, lineno, fcentry, fcedit, fcip, fcba, lastupdate, lasttime, kg1, kg_adj) " & _
											" values ('" & invoiceno & "'" & _ 
											" ,'" & request("jobcode_"&TRIM(n)) & "'" & _
											" ,'" & request("targettype_"&TRIM(n)) & "'" & _
											" ,'" & request("targetcode_"&TRIM(n)) & "'" & _
											" ,'INVOICE TBS-'" &"||" & "to_char(to_date('" & request("description_"&TRIM(n)) & "','mm/dd/yyyy'),'mm/dd/yyyy')" & _ 
											" , REPLACE('" & quantity & "',',', '')" & _
											" , REPLACE('" & request("rate_"&TRIM(n)) & "',',', '')" & _
											" , REPLACE('" & amount & "',',', '')" & _ 
											" , 0" & _
											" , REPLACE('" & amount & "',',', '')" & _
											" , '" & tempMaxLineNo & "'" & _
											" ,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
											Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI') " & _
											" , REPLACE('" & quantity1 & "',',', '')" & _
											" , REPLACE('" & quantity_adj & "',',', '')" & _
											")"
										 'response.write q
										connStr.execute(q) 
								Else
									err.number = 1001
									err.description = cek_trigger(request("targettype_"&TRIM(n)), request("targetcode_"&TRIM(n)), request("jobcode_"&TRIM(n)), invoice_by, transdate, "PR")
								End if
							Else
								err.number = 1001
								err.description = "Jobcode, Target Type dan Target Code is Required"
							End If
						End if	 
					Next 
				
				' 	'Untuk menghhitung jumlah total di header
					dtl_totalamount = 0
					dtl_amount = 0
					dtl_charge = 0
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select *  from invoiceservicedtl " & _
					  " where invoiceno='" & invoiceno & "' and " & _
					  " fcBA='" & fcba & "'"
					  ' response.write q
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
							dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("totalamount")))
							dtl_amount = dtl_amount + cdbl(to_nol(rsTempPurchaseRequest("amount")))
							dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercost")))
							rsTempPurchaseRequest.movenext
						wend
					End if								
					rsTempPurchaseRequest.Close : set rsTempPurchaseRequest = Nothing 
					totalppn = dtl_amount *cdbl(to_nol(ppn))/100
					totalpph = dtl_amount *cdbl(to_nol(pph))/100
					grandtotal = (dtl_amount+totalppn-totalpph)+dtl_charge		
					
					q =  " Update invoiceservicehdr set totalppn = replace('" & totalppn & "',',','')" & _
						 ", totalpph = replace('" & totalpph & "',',','')" & _ 
						 ", grandtotal = replace('" & grandtotal & "',',','')" & _
						 ", LASTUPDATE = sysdate " & _
						 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						 " where  invoiceno='" & invoiceno & "' and fcBA='" & fcba & "'"
						'response.write q
						connStr.execute(q)
							
					amount_spo = 0
					strLineMax = " SELECT NVL(grandtotal,0) grandtotal FROM servicepurchaseorder WHERE fcba = '" & fcba & "' AND pocode = '"&fcCode&"'"
					Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
					rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
					
					if not rsstrLineMax.eof then
						amount_spo = cdbl(to_nol(rsstrLineMax("grandtotal")))
					else
						amount_spo = 0
					end if									
					rsstrLineMax.Close : set rsstrLineMax = Nothing   
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
      cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		connstr.BeginTrans
		
		max_amount = 0 : max_charge = 0 : max_amountcharge = 0
	   
		in_data_amount = 0 : in_data_charge = 0 : in_data_amountcharge = 0
		strLineMax = 	" SELECT NVL(SUM(amount),0) amount, NVL(SUM(othercost),0) othercharge, NVL(SUM(totalamount),0) amountcharge" & _
						" FROM invoiceservicedtl" & _
						" WHERE pocode = '"&fcCode&"'" & _
						"       AND fcba = '" & fcba & "'" & _
						"       AND polinenokey = '" & spolineno & "'"

		Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
		rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic

		if not rsstrLineMax.eof then
			in_data_amount = cdbl(to_nol(rsstrLineMax("amount")))
			in_data_charge = cdbl(to_nol(rsstrLineMax("othercharge")))
			in_data_amountcharge = cdbl(to_nol(rsstrLineMax("amountcharge")))
		else
			in_data_amount = 0 : in_data_charge = 0 : in_data_amountcharge = 0
		end if							
		rsstrLineMax.Close : set rsstrLineMax = Nothing
		
		edit_data_amount = 0 : edit_data_charge = 0 : edit_data_amountcharge = 0
		strLineMax = 	" SELECT NVL(SUM(amount),0) amount, NVL(SUM(othercost),0) othercharge, NVL(SUM(totalamount),0) amountcharge" & _
						" FROM invoiceservicedtl" & _
						" WHERE invoiceno = '"&invoiceno&"'" & _
						"       AND fcba = '" & fcba & "'" & _
						"       AND lineno = '" & lineno & "'"

		Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
		rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic

		if not rsstrLineMax.eof then
			edit_data_amount = cdbl(to_nol(rsstrLineMax("amount")))
			edit_data_charge = cdbl(to_nol(rsstrLineMax("othercharge")))
			edit_data_amountcharge = cdbl(to_nol(rsstrLineMax("amountcharge")))
		else
			edit_data_amount = 0 : edit_data_charge = 0 : edit_data_amountcharge = 0
		end if							
		rsstrLineMax.Close : set rsstrLineMax = Nothing
	   
			   q = " UPDATE invoiceservicedtl" & _
			       " SET jobcode = '" & jobcode & "'" & _
			       "     , targetcode = '" & targetcode  & "'" & _
			       "     , targettype = '"& targettype & "'" & _
			       "     , kg = REPLACE('" & kg & "',',','')" & _
			       "     , rate = REPLACE('" & rate & "',',','')" & _
			       "     , amount = REPLACE('" & AMOUNT & "',',','')" & _
			       "     , othercost = REPLACE('" & OTHERCOST & "',',','')" & _
			       "     , totalamount = REPLACE('" & AMOUNT & "',',','') +REPLACE('" & OTHERCOST & "',',','')" & _
			       "     , LASTUPDATE = SYSDATE " & _
			       "     , LASTTIME = TO_CHAR(SYSDATE,'HH24:MI') "& _
			       " WHERE invoiceno = '" & invoiceno & "'" & _
			       "       AND fcBA='" & fcba & "'" & _
			       "       AND lineno = '"&lineno&"'"
				connStr.execute(q)
				
					'Untuk menghhitung jumlah total di header
					dtl_totalamount = 0
					dtl_amount = 0
					dtl_charge = 0
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select *  from invoiceservicedtl " & _
					  " where invoiceno='" & invoiceno & "' and " & _
					  " fcBA='" & fcba & "'"
					  ' response.write q
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
							dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("totalamount")))
							dtl_amount = dtl_amount + cdbl(to_nol(rsTempPurchaseRequest("amount")))
							dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercost")))
							rsTempPurchaseRequest.movenext
						wend 
						
					End if
					
					rsTempPurchaseRequest.Close
					set rsTempPurchaseRequest = Nothing
					totalamount = dtl_totalamount
					totalppn = dtl_amount *cdbl(to_nol(ppn))/100
					totalpph = dtl_amount *cdbl(to_nol(pph))/100
					grandtotal = (dtl_amount+totalppn-totalpph)+dtl_charge		
					
				  q= " Update invoiceservicehdr set totalppn = replace('" & totalppn & "',',',''), totalpph = replace('" & totalpph & "',',','')" & _
					 ", grandtotal = replace('" & grandtotal & "',',','')" & _
					 ", LASTUPDATE = sysdate " & _
					 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					 " where  invoiceno='" & invoiceno & "' and fcBA='" & fcba & "'"
					'response.write q
					connStr.execute(q)
				
		if err.number  = 0 then
		   connStr.CommitTrans
		end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Confirmed" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & fcba & "','" & fcCode & "', 'Confirm' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "RePlanned" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & fcba & "','" & fcCode & "', 'RePlan' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "Ordered" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & fcba & "','" & fcCode & "', 'Order' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
    elseif Request("txtJenis")= "UnRelease"	then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.UnReleasePO('" & fcCode & "' ); END; "
	   response.Write strQueryStatus
	   connStr.execute(strQueryStatus)	 
	   cmdView	= "True"
	elseif Request("txtJenis")= "CompletedDetails" then
		connstr.BeginTrans
		strQueryStatus = " BEGIN" & _
						 " PURCHASEORDERDETAILS_API.NextState('" & fcba & "'" & _
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

 Dim rsPurchase
' untuk menampilkan data field header

 set rsPurchase = server.CreateObject("Adodb.recordset")
	q= 	"   SELECT spo.invoiceno," & _
		"          spo.pocode," & _
		"          spo.suppliername," & _
		"          bu.fccompanycode," & _
		"          spo.fcba," & _
		"          spo.invoice_by," & _
		"          spo.valid_from," & _
		"          spo.valid_until," & _
		"          spo.transdate," & _
		"          spo.no_faktur," & _
		"          spo.fakturdate," & _
		"          spo.pocode," & _
		"          spo.invoice_type, " & _
		"          MAX (spo.ppn) ppn," & _
		"          MAX (spo.pph) pph," & _
		"          MAX (spo.totalppn) totalppn," & _
		"          MAX (spo.totalpph) totalpph," & _
		"          MAX (spo.grandtotal) grandtotal," & _
		"          SUM (spd.kg) kg," & _
		"          MAX (spd.rate) rate," & _
		"          SUM (spd.amount) amount," & _
		"          SUM (spd.othercost) othercost," & _
		"          SUM (spd.totalamount)" & _
		"     FROM invoiceservicehdr spo" & _
		"          INNER JOIN invoiceservicedtl spd" & _
		"             ON (spo.invoiceno = spd.invoiceno AND spo.fcba = spd.fcba)" & _
		"          INNER JOIN businessunit bu ON (bu.fccode = spo.fcba)" & _
		"    WHERE spo.invoiceno = '"& invoiceno &"'" & _
		" GROUP BY spo.invoiceno," & _
		"          spo.pocode," & _
		"          spo.suppliername," & _
		"          bu.fccompanycode," & _
		"          spo.fcba," & _
		"          spo.invoice_by," & _
		"          spo.valid_from," & _
		"          spo.valid_until," & _
		"          spo.transdate," & _
		"          spo.no_faktur," & _
		"          spo.fakturdate," & _
		"          spo.pocode," & _
		"          spo.invoice_type " 
	  
	  ' response.write q
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
	transdate = rsPurchase("transdate")
	fcCode = rsPurchase("pocode")
	suppliercode = rsPurchase("suppliername")
	totaldetamount = rsPurchase("amount")
	' totaldetcharge = rsPurchase("totaldetcharge")
	ppn = rsPurchase("ppn")
	pph = rsPurchase("pph")
	totalppn = rsPurchase("totalppn")
	totalpph = rsPurchase("totalpph")
	fcba  = rsPurchase("fcba")
	invoice_by  = rsPurchase("invoice_by")
	selectedinvoice  = rsPurchase("invoice_by")
	selectedfcba  = rsPurchase("fcba")
	fccompanycode  = rsPurchase("fccompanycode")
	othercost  = rsPurchase("othercost")
	grandtotal  = rsPurchase("grandtotal")
	fakturno  = rsPurchase("no_faktur")
	fakturdate  = rsPurchase("fakturdate")
	valid_from  = rsPurchase("valid_from")
	valid_until  = rsPurchase("valid_until")
	invoice_type  = rsPurchase("invoice_type")
	selectedsupplier = rsPurchase("suppliername")
	SudahClose="Y"
 End if

 ' dim rsFrez
 set rsFrez = server.CreateObject("Adodb.recordset")
	q=" select distinct frezee  from Purchaseorderdetails " & _
	  " where PoCode='" & fcCode & "' and " & _
	  " fcBA='" & fcba & "'"
 rsFrez.Open q, connstr
 if Not rsFrez.Eof then
	IsFrezee  = rsFrez("frezee")
 else
    IsFrezee ="0" 
 End if
 rsFrez.Close
 set rsFrez = Nothing 
 
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="invoicetbs.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Invoice No</td>
				<td align="left">
					<input type="text" maxlength="50" name="invoiceno" size="30" value="<%=invoiceno%>" autocomplete="off">
					<font color="#FF0000">*</font>
				</td>
				<td align="left">Amount Before Tax</td>
				<td align="left">
					<input type="text" maxlength="50" name="totaldetamount" size="30" value="<%=totaldetamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
				</td>
		</tr>
		<tr>
			<td align="left">Company</td>
				<td align="left">
					<select name="fccompanycode" value="<%=fccompanycode%>" id="fccompanycode" onchange="javascript:getlistfcba();javascript:getlistinvoice();" class="chosen-select" style="width: 50%">
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
				<td align="left">Other Charge</td>
				<td align="left">
						<input type="text" maxlength="50" name="totaldetcharge" size="30" value="<%=totaldetcharge%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
				</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<div id="fcbalist"></div>
				<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
				<font color="#FF0000">*</font>
			</td>
			<td align="left">Ppn (%)</td>
			<td align="left">
				<input type="text" size="9" name="ppn" value="<%=ppn%>" onkeypress="return harusangka(event);">
				<input type="text" size="15" name="totalppn" value="<%=totalppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="transdate" size="10" value="<%=transdate%>">
			</td>
			<td align="left">Pph (%)</td>
			<td align="left">
				<input type="text" size="9" name="pph" value="<%=pph%>" onkeypress="return harusangka(event);">
				<input type="text" size="15" name="totalpph" value="<%=totalpph%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By</td>
			<td align="left">
				<div id="invoicelist"></div>
			</td>
				<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="<%=selectedinvoice%>">
			<td align="left">Grand Total</td>
			<td align="left">
				<input type="text" maxlength="50" name="grandtotal" size="30" value="<%=grandtotal%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Supplier</td>
			<td align="left">
				<div id="supplierlist"></div>
					<input type="Hidden" id="selectedsupplier" name="selectedsupplier" value="<%=selectedsupplier%>">
				<font color="#FF0000">*</font>
			</td>
			<td align="left">No Faktur</td>
			<td align="left">
				<input type="text" maxlength="50" name="fakturno"  size="30" value="<%=fakturno%>" autocomplete="off" >
			</td>	
		<tr>
			<td align="left">Grading (%) </td>
			<td align="left">
				<input type="text" maxlength="10" name="grading" size="10" value="<%=grading%>">
			</td>
			<td align="left">Faktur Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="fakturdate" size="10" value="<%=fakturdate%>">
			</td>
		</tr>
		<tr>
			<td align="left">Valid From</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" id="valid_from" name="valid_from" size="10" value="<%=valid_from%>">
			</td>
			<td align="left">Valid Until</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" id="valid_until" name="valid_until" size="10" value="<%=valid_until%>">
			</td>
		</tr>
		</tr>
		<%if businessunit = "KKP" OR businessunit = "KPK" then%>
		<tr>
		<td align="left">Ppn (10%)</td>
		<td align="left">
			<input type="checkbox" id="use_ppn" class="css-checkbox" name="use_ppn" value="<%=use_ppn%>" onclick='handleClick(this);' <%if use_ppn = "Y" then%> checked <%end if%>>
			<label for="use_ppn" class="css-label"></label>
		</td>
		</tr>
		<%end if%>
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		
		<%if len(trim(businessunit)) = "KKP" OR len(trim(businessunit)) = "KPK" then%>
		<a href="#" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp;
		<%else%>
	    <a href="#" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp;
		<%end if%>
		<a href="invoicetbs.asp?invoiceno=<%=invoiceno%>&fcCode=<%=fcCode%>&transdate=<%=transdate%>&fccompanycode=<%=fccompanycode%>&fcba=<%=fcba%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" And cint(IsInBridge) = 1 then %>
		    <a href="invoicetbs.asp?fccompanycode=<%=fccompanycode%>&fcba=<%=fcba%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=Ordered" onClick="return confirm('Are you sure to Order this data :  <%=fcCode%> ?');"> Ordered </a>
		<% elseif RowStateId = "Planned" And cint(IsInBridge) = 0 then %>
		    <a href="invoicetbs.asp?fccompanycode=<%=fccompanycode%>&fcba=<%=fcba%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to Confirm this data :  <%=fcCode%> ?');"> Confirmed </a>
		<% elseif RowStateId = "Confirmed" then %>	
		    <a href="invoicetbs.asp?fccompanycode=<%=fccompanycode%>&fcba=<%=fcba%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=fcCode%> ?');"> Re-Planned </a>
		<% elseif  RowStateId = "Ordered" And cint(IsInBridge) = 1 And Cint(IsFrezee) =0 then %>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="invoicetbs.asp?fccompanycode=<%=fccompanycode%>&fcba=<%=fcba%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=UnRelease" onClick="return confirm('Are you sure to Un Release this data :  <%=fcCode%> ?');"> UnRelease </a>
		<% end if%>
		</td>
		</tr>
    </table>  
 </form>
<% If cmdView = "True" and len(trim(invoiceno)) > 0  Then

If SudahClose <> "Y" Then  
	q1 = " SELECT count(*) total FROM invoiceservicehdr WHERE fcba = '"& fcba &"'" & _
		 " AND suppliername = '"& suppliercode &"'" & _
		 " AND valid_from >= TO_DATE ('"& valid_from &"', 'mm/dd/yyyy')" & _
		 " AND valid_until <= TO_DATE ('"& valid_until &"', 'mm/dd/yyyy')" 
	' response.write q1
	Set rsTotal = Server.CreateObject("ADODB.Recordset")
   	rsTotal.Open q1, connstr

   	If Not rsTotal.Eof Then 
   		total = cdbl(to_nol(rsTotal("total")))
   	End If

   	If total = 0 Then 
		q = "   SELECT fcba," & _
			"          field_estate_estatecode," & _
			"          cust_customercode," & _
			"          postingdate," & _
			"          SUM (jjg) jjg," & _
			"          SUM (qtybeforegrading) qtybeforegrading," & _
			"          ROUND ( ( (SUM (qty_grading) / SUM (qtybeforegrading)) * 100), 2) percent_grading," & _
			"          SUM (qty_grading) qty_grading," & _
			"          SUM (qtyaftergrading) quantity" & _
			"     FROM (SELECT wb.fcba," & _
			"                  wb.field_estate_estatecode," & _
			"                  wb.cust_customercode," & _
			"                  wb.postingdate," & _
			"                  wb.spbno," & _
			"                  wb.chitno," & _
			"                  wb.vehicleno," & _
			"                  wb.tbstype," & _
			"                  wb.jjg," & _
			"                  (wb.weightin - wb.weightout) AS qtybeforegrading," & _
			"                  ROUND ( ( (gr.weight / (wb.weightin - wb.weightout)) * 100)," & _
			"                         2)" & _
			"                     percent_grading," & _
			"                  gr.weight AS qty_grading," & _
			"                  (wb.weightin - wb.weightout - gr.weight) AS qtyaftergrading" & _
			"             FROM weighbridgeticket wb" & _
			"                  INNER JOIN gradingrecord gr" & _
			"                     ON (wb.fcba = gr.fcba AND wb.chitno = gr.wbt_chitno)" & _
			"            WHERE     wb.tbstype = 'TBS'" & _
			"                  AND wb.fcba = '"& fcba &"'" & _
			"                  AND datein >= TO_DATE ('"& valid_from &"', 'mm/dd/yyyy')" & _
			"                  AND dateout <= TO_DATE ('"& valid_until &"', 'mm/dd/yyyy')" & _
			"                  AND cust_customercode = '"& suppliercode &"')" & _
			" GROUP BY postingdate, fcba," & _
			"          field_estate_estatecode," & _
			"          cust_customercode " & _
			" ORDER BY postingdate, fcba," & _
			"          field_estate_estatecode," & _
			"          cust_customercode "
	Else
		q = " select spo.invoiceno, spo.transdate, spo.no_faktur, spo.fakturdate, spo.pocode, spo.invoice_type," & _
			" spd.jobcode, spd.targetcode, spd.targettype, spd.description," & _
			" spd.kg, spd.rate, spd.amount, spd.description, spd.amount, spd.othercost, spd.totalamount, spd.lineno, spd.polinenokey" & _
			" from invoiceservicehdr spo " & _
			" inner join invoiceservicedtl spd " & _
			" on (spo.invoiceno = spd.invoiceno and spo.fcba = spd.fcba)" & _
			" where spo.invoiceno = '"& invoiceno &"'" & _
			" ORDER BY lineno"
	End If
Else	 
	q = " select spo.invoiceno, spo.transdate, spo.no_faktur, spo.fakturdate, spo.pocode, spo.invoice_type," & _
		" spd.jobcode, spd.targetcode, spd.targettype, spd.description," & _
		" spd.kg, spd.rate, spd.amount, spd.description, spd.amount, spd.othercost, spd.totalamount, spd.lineno, spd.polinenokey" & _
		" ,spd.kg1,spd.kg_adj " & _
		" from invoiceservicehdr spo " & _
		" inner join invoiceservicedtl spd " & _
		" on (spo.invoiceno = spd.invoiceno and spo.fcba = spd.fcba)" & _
		" where spo.invoiceno = '"& invoiceno &"'" & _
		" ORDER BY lineno"
		 
End if
   ' response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.Open q, connstr
 %>
<form action="invoicetbs.Asp" method="post" name="eForm">
   <table class="data" border="1" width="100%">
   <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">Date</td>
	   <td align="center">Jobcode</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <td align="center">Netto1 (Kg)</td>
	   <td align="center">Adjustment (Kg)</td>
	   <td align="center">Netto2 (Kg)</td>
	   <td align="center">Rate (Rp)</td>
	   <td align="center">Amount Invoice (Rp)</td>
	   <td align="center"><input type="checkbox" name="select-all" id="select-all"/> Check/Uncheck All</td>
	   </tr>
   <%ELSE%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">&nbsp;</td>
	   <td align="center">Description</td> 
	   <td align="center">Jobcode</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <td align="center">Netto1 (Kg)</td> 
	   <td align="center">Adjustment (Kg)</td>
	   <td align="center">Netto2 (Kg)</td> 
	   <td align="center">Rate (Rp)</td>
	   <td align="center">Amount Invoice (Rp)</td>
	   </tr>
   <%End if
   if Not rsGRN.eof then
      nCountofField=rsGRN.fields.count-1
	  no=1
	  v_Quantity = 0 
	  v_amount = 0
	  While Not rsGRN.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="ringht" valign="top"><%=NO%></td>
			<%IF SudahClose="Y" then%>
				<td valign="middle">
					<a href="Invoicetbs.Asp?invoiceno=<%=rsGRN("invoiceno")%>&fcba=<%=fcba%>&lineno=<%=rsGRN("lineno")%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0"></a></td>
					<td align="center" title="Description" valign="top"><%=rsGRN("description")%></td>
					<td align="center" title="Supplier" valign="top"><%=rsGRN("jobcode")%></td>
					<td align="center" title="Item Code" valign="top"><%=rsGRN("targettype")%></td>
					<td align="center" title="Item Name" valign="top"><%=rsGRN("targetcode")%></td>
					<td align="right" title="Netto1" valign="top"><%=FormatNumber(cdbl(rsGRN("kg1")),2)%></td>
					<td align="right" title="Adjustment" valign="top"><%=FormatNumber(cdbl(rsGRN("kg_adj")),2)%></td>
					<td align="right" title="Netto2" valign="top"><%=FormatNumber(cdbl(rsGRN("kg")),2)%></td>
					<td align="right" title="Rate" valign="top"><%=FormatNumber(cdbl(rsGRN("rate")),2)%></td>
					<td align="right" title="totalamount" valign="top"><%=FormatNumber(cdbl(rsGRN("totalamount")),2)%></td> 
			<%End if%>

			<%if SudahClose<>"Y" Then%>
                    <td align="center" title="" valign="top"><%=rsGRN("postingdate")%></td>
			        <td align="center" title="" valign="top">
						<input style="text-align:left;" type="text" name="jobcode_<%=no%>" value ="<%=jobcode%>">
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>
						<span class="satuan_item" id="jobname">&nbsp;<%= JobCodeDescription %></span>
					</td>
			        <td align="center" title="" valign="top">
						<input style="text-align:left;" type="text" name="targettype_<%=no%>" value = "<%=targettype%>">
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Target Type</font></a>
					</td>
			        <td align="left" title="" valign="top">
						<input style="text-align:left;" type="text" name="targetcode_<%=no%>" value = "<%=targetcode%>">
					</td>
			        <td align="right" title="" valign="top">
			        	<input style="text-align:right;background-color:#eee;" type="text" name="Quantity_<%=no%>" value = "<%=rsGRN("qtybeforegrading")%>" readonly>
			        </td>
                    <td align="right" title="" valign="top">
<!--						<input style="text-align:right;" type="text" name="quantity_adjustment_<%=no%>" value = "<%=quantity_adjustment%>"> -->
						<input style="text-align:right;" type="text" name="quantity_adjustment_<%=no%>" value = "<%=rsGRN("qty_grading")%>">

					</td>
					
					 <td align="right" title="" valign="top">
						<input style="text-align:right;" type="text" name="quantity2_<%=no%>" value = "<%=rsGRN("quantity")%>">
					</td>
					
					<td align="right" title="" valign="top">
						<input style="text-align:right;" type="text" name="rate_<%=no%>" value = "<%=rate%>">
					</td>
					<td align="center">
						<input style="text-align:right;background-color:#eee;" type="text" name="amount_<%=no%>" class="amount" value = "<%=amount%>"  autocomplete="off" readonly>
					</td> 
		            <td align="center">
						<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
					</td> 
					<input type="hidden" name="description_<%=no%>" value="<%=rsGRN("postingdate")%>">
				 </td> 
			<%End if%>
		  </tr>
		  <%
		  IF SudahClose="Y" then
		    v_Quantity1 = v_Quantity1 + Trim(cdbl(rsGRN("kg1"))) 
			v_Quantity_adj = v_Quantity_adj + Trim(cdbl(rsGRN("kg_adj"))) 
			v_Quantity = v_Quantity + Trim(cdbl(rsGRN("kg")))
			v_amount = v_amount + Trim(cdbl(rsGRN("totalamount")))
		  ELSE
		    v_Quantity1 = v_Quantity1 + Trim(cdbl(rsGRN("qtybeforegrading"))) 
			v_Quantity_adj = v_Quantity_adj + Trim(cdbl(rsGRN("qtybeforegrading"))) 
			v_Quantity = v_Quantity + Trim(cdbl(rsGRN("quantity")))
		  END IF
		  rsGRN.MoveNext 
		  no=no+1
		  intRecordsPerPage = intRecordsPerPage - 1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsGRN.Close
   set rsGRN = Nothing%>
		<tr>
			<%IF SudahClose<>"Y" then%>
				<td align="right" title="" valign="top" colspan="5"><font face="Verdana" size="1" style="font-weight: bold">TOTAL</font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_Quantity1),0)%> </font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_Quantity_adj),0)%> </font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_Quantity),0)%> </font></td>
				<td colspan="4"></td> 
			<%Else%>
				<td align="right" title="" valign="top" colspan="6"><font face="Verdana" size="1" style="font-weight: bold">TOTAL</font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_Quantity1),0)%> </font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_Quantity_adj),0)%> </font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_Quantity),0)%> </font></td>				
				<td></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_amount),2)%> </font></td>
			<%End if%>
		</tr>
	  <tr>
		<td colspan="11" align="left">
		<!-- diambil dari form yang diatas -->
			<input type="hidden" name="No" value="<%=No%>">
			<input type="hidden" name="fcCode" value="<%=fcCode%>">
			<input type="hidden" name="transdate" value="<%=transdate%>">
			<input type="hidden" name="valid_from" value="<%=valid_from%>">
			<input type="hidden" name="valid_until" value="<%=valid_until%>">
			<input type="hidden" name="fakturno" value="<%=fakturno%>">
			<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
			<input type="hidden" name="suppliercode" value="<%=suppliercode%>">
			<input type="hidden" name="selectedsupplier" value="<%=selectedsupplier%>">
			<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
			<input type="hidden" name="fcba" value="<%=fcba%>">
			<input type="hidden" name="invoice_by" value="<%=invoice_by%>">
			<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
			<input type="hidden" name="selectedinvoice" value="<%=selectedinvoice%>">
			<input type="hidden" name="invoiceno" value="<%=invoiceno%>">
			<input type="hidden" name="dp" value="<%=dp%>">
			<input type="hidden" name="ppn" value="<%=ppn%>">
			<input type="hidden" name="pph" value="<%=pph%>">
			<input type="hidden" name="grandtotal" value="<%=grandtotal%>">
			<input type="hidden" name="totalppn" value="<%=totalppn%>">
			<input type="hidden" name="totalpph" value="<%=totalpph%>">
			<input type="hidden" name="grading" value="<%=grading%>">
			<input type="hidden" name="RowStatehdr" value="<%=RowStatehdr%>">
			<input type="hidden" name="lineno" value="<%=lineno%>">
			<input type="hidden" name="invoice_type" value="<%=invoice_type%>">

			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<!-- Validasi Period Disini -->
			<%IF Request("txtJenis")="Edit" then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit"  name="Add" value="Update">
			<%Else%>
				<input type="hidden" name="txtJenis" value="New">
				<input Type="Submit"  name="Add" value="Submit">
			<%End if%>
				<input Type="Reset"  name="Reset" value="Reset">
				<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
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
	  '.chosen-select-width'     : {width:"95%"}
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

function getselectedsupplier() {
	var spocode = document.getElementById("selectedsupplier");
	spocode.value = document.getElementById("suppliercode").value;
}

function getselectedinvoice() {
	var spocode = document.getElementById("selectedinvoice");
	spocode.value = document.getElementById("invoice_by").value;
}


	$(function() {
		$('#valid_date').hide(); 
		$('#invoice_type').change(function(){
			if($('#invoice_type').val() == 'TBS') {
				$('#valid_date').show(); 
			} else {
				$('#valid_date').hide(); 
			} 
		});
	});

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
	getlistsupplier();
	getlistinvoice();
	getselectedfcba();
	getselectedsupplier();
	getselectedinvoice();
	hitungrumus();
});
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
		var txt = chk.parentNode.parentNode.cells[9].getElementsByTagName('input')[0];
		
		if(chk.checked == true)
		{
			txt.readOnly = true;
			txt.style.backgroundColor = "#eee";
		}
		else
		{
			txt.readOnly = true;
			txt.style.backgroundColor = "#eee";
		}
	}
</script>
<script type="text/javascript">	
	$('#select-all').click( function () {
		$('.item').prop('checked', this.checked);
		$('.item').val(this.checked ? 1 : 0);
		if($('#select-all').is(":checked")){
			$('.amount').css("background-color", "#eee");
			$('.amount').prop('readOnly', true);
		}else{
			$('.amount').css("background-color", "#eee");
			$('.amount').prop('readOnly', true);
		}
	})
</script>
</body>
</html>
