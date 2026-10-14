<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>FBPlants - Edit Header Contractor Invoice</title>
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

	// function hitungrumus(){
 //    	var totalamount_afterretention = document.getElementById("totalamount_afterretention");
 //    	var ppn = document.getElementById("ppn");
    	
 //    	var totalamountppn = Math.round( (totalamount_afterretention.value * ppn.value / (100) );
 //    	alert(totalamountppn);
 //    	if (!isNaN(totalamountppn)) {
 //    	 document.getElementById('amountppn').value = totalamountppn;
 //    	}
	// }
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
</script>

<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.]/g,'');}
</script>

<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">

<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

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

var getlistvoucher = function () {
	// build up the post string when passing variables to the server side page
	
	// var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedagreement = document.getElementById("selectedagreement").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	// var transdate = document.getElementById("transdate").value;
	var PostStr = 'fcba='+selectedfcba+'&agreementcode='+selectedagreement;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getwipinfobasedonfcba.asp', 'POST', PostStr, showlistvoucher);
}

// The function for handling the response from the server
var showlistvoucher = function (oXML) { 
    // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("agreementlist").innerHTML = "";
	document.getElementById("agreementlist").innerHTML = response;

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

var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfcbacibasedoncompany.asp', 'POST', PostStr, showlistfcba);
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


var getretention = function () {
	// build up the post string when passing variables to the server side page

	var selectedfcba = document.getElementById("selectedfcba").value;
	var agreementcode = document.getElementById("agreementcode").value;
	var PostStr = 'fcba='+selectedfcba+'&agreementcode='+agreementcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getRetentionCI.asp', 'POST', PostStr, showretention);
}

// The function for handling the response from the server
var showretention = function (oXML) { 
    // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);

    // update the Div to show the result from the server
	document.getElementById("retention").value = s.retention;
	document.getElementById("contractorcode").value = s.contractorcode;
	document.getElementById("contractorname").value = s.contractorname;
	// document.getElementById("vat").value = s.ppn;
}

var getamountinfo = function () {
	// build up the post string when passing variables to the server side page

	var selectedfcba = document.getElementById("selectedfcba").value;
	var agreementcode = document.getElementById("agreementcode").value;
	var jobcode = document.getElementById("jobcode").value;
	var PostStr = 'fcba='+selectedfcba+'&agreementcode='+agreementcode+'&jobcode='+jobcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getinfoamountci.asp', 'POST', PostStr, showamountinfo);
}

// The function for handling the response from the server
var showamountinfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    //alert (s.postingdate);
    
    // update the Div to show the result from the server
	document.getElementById("quantity").value = s.quantity;
	document.getElementById("rate").value = s.rate;
	document.getElementById("retention").value = s.retention;
	document.getElementById("retention").value = "";
	document.getElementById("quantity").value = "";
	document.getElementById("rate").value = "";

	var rate = s.rate;
	var quantity = s.quantity;
	var retention = s.retention;
	var amount = rate * quantity;
	var amountretention = amount * retention / 100
	var amountafter_retention = amount - amountretention
	document.getElementById("amount").value = amount;
	document.getElementById("amountretention").value = amountretention;
	document.getElementById("amountafter_retention").value = amountafter_retention;
	// document.getElementById("totalamount_afterretention").value = amountafter_retention;
	document.getElementById("retention").value = retention;
	document.getElementById("quantity").value = s.quantity;
	document.getElementById("rate").value = s.rate;
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
 
 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks,invoicecode
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks,txtNoPO,v_CrossCekPO
 Dim recordTotalAmount, rsrecordTotalAmount, varTotalAmount, updateNetBalance,dp,selecteddp
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""


 Dim Jobcode, LocationType, LocationCode, Reference, Remark
dim ScriptName,modulFile,executeQry
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")



 

 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

 'if Request("invoicecode") <> "" then
	
    'invoicecode="" '----- Nantinya Ambil dari Format Master
 'end if
'Response.Write invoicecode
 
 '---- Data Header
 
 if Request("transdate")<>"" then
   transdate   =Request("transdate")
 else
   transdate   =textDate(date)
 end if
 
  transdate = FormatDateTime(transdate)

 ' If Len(Trim(invoicecode)) = 0 Then invoicecode = "(NEW)"

 if Request("fakturdate")<>"" then
   fakturdate   =Request("fakturdate")
 else
   fakturdate   =textDate(date)
 end if

 xCode = Request("xCode")
 HeaderRemarks = Request("HeaderRemarks")
 
 '---- Data Details

 JobCode=Request("JobCode")
 LocationType=Request("LocationType")
 LocationCode=Request("LocationCode")
 Reference=Request("Reference")
 Remarks=Request("Remarks")
 Amount=Request("Amount")
 agreementlineno =Request("agreementlineno")
 fcba =Request("fcba")
 agreementcode =Request("agreementcode")
 selectedagreement =Request("selectedagreement")
 dp =Request("dp")
 selecteddp =Request("selecteddp")
 
 fccompanycode = request("fccompanycode")
 selectedfcba = Request("selectedfcba")
 contractorcode = Request("contractorcode")
 retention = Request("retention")
 isretention_invoice = Request("isretention_invoice")
 totalamount_afterretention = Request("totalamount_afterretention")
 ppn = Request("ppn")
 amountppn = Request("amountppn")
 pph = Request("pph")
 invoicecode = Request("invoicecode")
 amountpph = Request("amountpph")
 grandtotal_aftertax = Request("grandtotal_aftertax")
 description = Request("description")
 rowstate = Request("rowstate")
 totalfinal = Request("totalfinal")
 rate = Request("rate")
 amount = Request("amount")
 amountretention = Request("amountretention")
 amountafter_retention = Request("amountafter_retention")
 remarks = Request("remarks")
 lineno = Request("lineno")
 invoice_by = Request("invoice_by")
 selectedinvoice = Request("selectedinvoice")
 contractorname = Request("contractorname")
 agreementlinenotxt = Request("agreementlinenotxt")
  
sub post_amount
    set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select ppn, amountppn, pph, amountpph,dp from contractorinvoicehdr " & _
					  " where invoicecode='" & invoicecode & "' and " & _
					  " fcBA='" & fcBA & "'"

	rsTempPurchaseRequest.Open q, connstr
	if Not rsTempPurchaseRequest.Eof then
		ppn = cdbl(to_nol(rsTempPurchaseRequest("ppn")))
		pph = cdbl(to_nol(rsTempPurchaseRequest("pph")))
		selecteddp = cdbl(to_nol(rsTempPurchaseRequest("dp")))
	Else					
		ppn = 0
		pph = 0
		selecteddp = 0
	End if
	rsTempPurchaseRequest.Close
	set rsTempPurchaseRequest = Nothing
					
    set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
	q=	"select NVL(SUM(amountafter_retention),0) amountafter_retention from contractorinvoicedtl " & _
		" where invoicecode='" & invoicecode & "' and " & _
		" fcBA='" & fcBA & "'"

	rsTempPurchaseRequest.Open q, connstr
	if Not rsTempPurchaseRequest.Eof then
		amountafter_retention = cdbl(to_nol(rsTempPurchaseRequest("amountafter_retention")))
	Else
	    amountafter_retention = 0
	End if
					
	rsTempPurchaseRequest.Close
	set rsTempPurchaseRequest = Nothing

	dtl_amountppn = amountafter_retention * ppn / 100
	dtl_amountpph = amountafter_retention * pph / 100
	dtl_grandtotal_aftertax = amountafter_retention + dtl_amountppn - (dtl_amountpph+selecteddp)
					'end add
	q= " Update contractorinvoicehdr set totalamount_afterretention = replace('" & amountafter_retention & "',',',''), " & _
	   " amountppn = replace('" & dtl_amountppn & "',',',''), " & _
	   " amountpph = replace('" & dtl_amountpph & "',',',''), " & _
	   " grandtotal_aftertax = replace('" & dtl_grandtotal_aftertax & "',',','')" & _
	   " where  invoicecode='" & invoicecode & "' and fcBA='" & fcba & "'"
	connStr.execute(q)
	modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
	executeQry	= "Query executed:" & q
	ExecuteAndLog modulFile,executeQry
	
end sub
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 
 
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")="Posting" Then
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		if cstr(""&isretention_invoice) = "FALSE" then
			strFieldDetail = " SELECT DISTINCT agreementcode,to_char(wipdate,'MM/DD/YYYY') wipdate FROM contractorinvoicedtl " & _
							 " WHERE invoicecode = '" & invoicecode & "'" & _
							 "       AND fcBA = '" & fcba & "'"								
			' Response.Write strFieldDetail
			Set rsQty = Server.CreateObject("ADODB.Recordset")
			rsQty.Open strFieldDetail, connstr	
			while Not rsQty.Eof 		
			    
				agreementcode = rsQty("agreementcode")
				wipdate       = rsQty("wipdate")
				q ="update workinprogressfinal set isinvoice=0 where agreementcode='" & agreementcode & "' and fcba='" & fcba & "' and wipdate= to_date('" & wipdate & "','MM/DD/YYYY')" 
				connStr.execute(q)
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & q
				ExecuteAndLog modulFile,executeQry
				rsQty.movenext
			wend 
		end if
					
		
		' if has_detail = 0 then
			q = " DELETE FROM contractorinvoicedtl WHERE invoicecode = '" & invoicecode & "' AND fcba = '"& fcba &"'"
			connStr.execute(q)
			modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
			executeQry	= "Query executed:" & q
			ExecuteAndLog modulFile,executeQry
			
			q = " DELETE FROM contractorinvoicehdr WHERE invoicecode = '" & invoicecode & "' AND fcba = '"& fcba &"'"
			connStr.execute(q)
			modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
			executeQry	= "Query executed:" & q
			ExecuteAndLog modulFile,executeQry

		' else
		' 	err.number = 1001
		' 	err.description = " Contractor Invoice """ & invoicecode & """ has detail, please delete all detail first!"
		' End If
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
        cmdView	= "True"
    elseif Request("txtJenis")= "Posting" then
					
					strQueryStatus = "UPDATE contractorinvoicehdr SET POSTINGFLG='1' WHERE fcba = '" & fcba & "' AND invoicecode = '" & invoicecode & "'"
					connStr.execute(strQueryStatus)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & strQueryStatus
					ExecuteAndLog modulFile,executeQry
					
					strQueryStatus = "BEGIN fac_journal_inv_ico_api.ProsesGenerate('" & invoicecode & "', '" & fcba & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & strQueryStatus
					ExecuteAndLog modulFile,executeQry
					
					
					strQueryStatus = "update BKMJOURNALHDR set postingflg=1 where fcba='" & fcba & "' AND docno='" & invoicecode & "'"
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & strQueryStatus
					ExecuteAndLog modulFile,executeQry
	
					strQueryStatus = "BEGIN fac_closing_accounting_api.ProsesClosing(TO_NUMBER(TO_CHAR(TO_DATE('" & transdate  & "','MM/DD/YYYY'),'YYYY')), " & _ 
			                         "TO_NUMBER(TO_CHAR(TO_DATE('" & transdate  & "','MM/DD/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus	& "<br>"			 
					connStr.execute(strQueryStatus)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & strQueryStatus
					ExecuteAndLog modulFile,executeQry
		      
					cmdView	= "True"	
    elseif Request("txtJenis")= "Details" then
	 
    connstr.BeginTrans
		q	=	"DELETE FROM contractorinvoicedtl WHERE invoicecode = '"& invoicecode &"' AND lineno = '"& lineno &"' AND fcba = '"& fcba &"'"
		connStr.execute(q)
		modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
		executeQry	= "Query executed:" & q
		ExecuteAndLog modulFile,executeQry

		if err.number  = 0 then
			 connStr.CommitTrans
		  end if
		cmdView	= "True"

	elseif Request("txtJenis")= "New" then
	
	connstr.BeginTrans
	Dim rsSupplyingDeclarationHdr
		
		'--End Added
		Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select invoicecode from contractorinvoicehdr where invoicecode = '"& invoicecode &"' and fcba ='"& fcba &"'" , connstr
			If rsSupplyingDeclarationHdr.EOF Then

			if selecteddp = "" then 
				selecteddp= 0
			end if
					
			q = " INSERT INTO contractorinvoicehdr ( invoicecode" & _
				"                           , transdate" & _
				"                           , agreementcode" & _
				"                           , contractorcode" & _
				"                           , isretention_invoice" & _
				"                           , retention" & _
				"                           , totalamount_afterretention" & _
				"                           , ppn" & _
				"                           , amountppn" & _
				"                           , pph" & _
				"                           , amountpph" & _
				"                           , grandtotal_aftertax" & _
				"                           , description" & _
				"                           , rowstate" & _
				"                           , invoice_by" & _
				"                           , fcentry" & _
				"                           , fcedit" & _
				"                           , fcip" & _
				"                           , fcba" & _
				"                           , lastupdate" & _
				"                           , lasttime " & _
				"                           , payment " & _
				"                           , dp)" & _
				" VALUES ( '" & invoicecode & "'" & _
				"          , TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
				"          , '" & selectedagreement & "'" & _
				"          , '" & contractorcode & "'" & _
				"          , '" & isretention_invoice & "'" & _
				"          , '" & retention & "'" & _
				"          , replace('" & totalamount_afterretention & "',',','')" & _
				"          , '" & ppn & "'" & _
				"          , replace('" & amountppn & "',',','')" & _
				"          , '" & pph & "'" & _
				"          , replace('" & amountpph & "',',','')" & _
				"          , replace('" & grandtotal_aftertax & "',',','')" & _
				"          , '" & description & "'" & _
				"          , 'INVOICED'" & _
				"          , '" & invoice_by & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , '" & fcba & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'),'0','" & selecteddp & "')"			 

			connStr.execute(q)
			modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
			executeQry	= "Query executed:" & q
			ExecuteAndLog modulFile,executeQry
			Else
				
			End If
			rsSupplyingDeclarationHdr.Close
			' set rsSupplyingDeclarationHdr = Nothing
		
				' Added by NAF, 2012-12-01
				 strLineMax = " select nvl(max(lineno),0) maxLineNoKey from contractorinvoicedtl where fcBA='" & fcba & "'"

				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr
				
				if not rsstrLineMax.eof then
					tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
				end if
							
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 
				' End of Added by NAF, 2012-12-01

		
			if cstr(""&isretention_invoice) = "FALSE" then
				For n=1 to (cdbl(request("no"))-1) 
					if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
							q = " Insert into  contractorinvoicedtl (invoicecode, agreementcode, jobcode, totalfinal, rate, amount, amountretention, amountafter_retention, lineno, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, WIPDATE,AGREEMENTLINENO) " & _
								" values ('" & invoicecode & "'" & _
								" ,'" & agreementcode & "'" & _
								" ,'" & request("jobcode_"&TRIM(n)) & "'" & _
								" ,REPLACE('" & request("qty_final_"&TRIM(n)) & "', ',', '')" & _
								" ,REPLACE('" & request("rate_"&TRIM(n)) & "', ',', '')" & _
								" ,REPLACE('" & request("amount_"&TRIM(n)) & "', ',', '')" & _
								" ,REPLACE('" & request("amountretention_"&TRIM(n)) & "', ',', '')" & _
								" ,REPLACE('" & request("amountafter_retention_"&TRIM(n)) & "', ',', '')" & _
								" ,(SELECT NVL(MAX(lineno),0) FROM contractorinvoicedtl WHERE invoicecode = '" & invoicecode & "' AND fcba = '"& fcba &"' and agreementcode = '"& agreementcode &"') +1" & _
								" ,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI')" & _
								" ,TO_DATE('" & request("wipdate_"&TRIM(n))  & "','MM/DD/YYYY'),'" & request("agreementlinenotxt"&TRIM(n)) & "'  " & _ 
								")"
							connStr.execute(q)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & q
							ExecuteAndLog modulFile,executeQry

							'untuk update kolom isinvoice
							q ="update workinprogressfinal set isinvoice=1 where agreementcode='" & agreementcode & "' and fcba='" & fcba & "' and wipdate= to_date('" & request("wipdate_"&TRIM(n)) & "','MM/DD/YYYY') and AGREEMENTFINALNO= " & request("agreementfinalno_"&TRIM(n)) & "" 
							connStr.execute(q)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & q
							ExecuteAndLog modulFile,executeQry

					' Untuk menghhitung jumlah total di header
					dtl_amountppn = 0
					dtl_amountpph = 0
					dtl_grandtotal_aftertax = 0
					' dtl_charge = 0
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select ppn, amountppn, pph, amountpph from contractorinvoicehdr " & _
					  " where invoicecode='" & invoicecode & "' and " & _
					  " fcBA='" & fcBA & "' and agreementcode = '"& agreementcode &"'"

					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
							amountppn = cdbl(to_nol(rsTempPurchaseRequest("amountppn"))) 
							amountpph = cdbl(to_nol(rsTempPurchaseRequest("amountpph")))
							ppn = cdbl(to_nol(rsTempPurchaseRequest("ppn")))
							pph = cdbl(to_nol(rsTempPurchaseRequest("pph")))
						rsTempPurchaseRequest.movenext
					wend 
						
					End if

					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select SUM(amountafter_retention) amountafter_retention from contractorinvoicedtl " & _
					  " where invoicecode='" & invoicecode & "' and " & _
					  " fcBA='" & fcBA & "' and agreementcode = '"& agreementcode &"'"

					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
							amountafter_retention = cdbl(to_nol(rsTempPurchaseRequest("amountafter_retention")))
						rsTempPurchaseRequest.movenext
					wend 
						
					End if
					
					rsTempPurchaseRequest.Close
					set rsTempPurchaseRequest = Nothing

					dtl_amountppn = amountafter_retention * ppn / 100
					dtl_amountpph = amountafter_retention * pph / 100
					dtl_grandtotal_aftertax = amountafter_retention + dtl_amountppn - dtl_amountpph
					'end add
				   q= " Update contractorinvoicehdr set totalamount_afterretention = replace('" & amountafter_retention & "',',',''), " & _
				   	  " amountppn = replace('" & dtl_amountppn & "',',',''), " & _
				   	  " amountpph = replace('" & dtl_amountpph & "',',',''), " & _
				   	  " grandtotal_aftertax = replace('" & dtl_grandtotal_aftertax & "',',','')" & _
					  " where  invoicecode='" & invoicecode & "' and agreementcode= '"& agreementcode &"' and fcBA='" & fcba & "'"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry
					 
				End if	 
				Next
			Else
				For n=1 to (cdbl(request("no"))-1) 
					if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
							q = " Insert into  contractorinvoicedtl (invoicecode, agreementcode, jobcode, totalfinal, rate, amount, amountretention, amountafter_retention, lineno, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, WIPDATE,AGREEMENTLINENO,Remarks) " & _
								" values ('" & invoicecode & "'" & _
								" ,'" & agreementcode & "'" & _
								" ,'" & request("jobcode_"&TRIM(n)) & "'" & _
								" ,0" & _
								" ,0" & _
								" ,REPLACE('" & request("amountretention_"&TRIM(n)) & "', ',', '')" & _
								" ,0" & _
								" ,REPLACE('" & request("amountretention_"&TRIM(n)) & "', ',', '')" & _
								" ,(SELECT NVL(MAX(lineno),0) FROM contractorinvoicedtl WHERE invoicecode = '" & invoicecode & "' AND fcba = '"& fcba &"' and agreementcode = '"& agreementcode &"') +1" & _
								" ,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI')" & _
						        " ,TO_DATE('" & request("wipdate_"&TRIM(n))  & "','MM/DD/YYYY'),'" & request("agreementlinenotxt"&TRIM(n)) & "' " & _
								" ,'" & request("bappno_"&TRIM(n)) & "'" & _ 
								")"

							'response.write q + "<br>"
							connStr.execute(q)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & q
							ExecuteAndLog modulFile,executeQry

							'untuk update kolom isinvoice
							q ="update workinprogressfinal set isinvoice=1 where agreementcode='" & agreementcode & "' and fcba='" & fcba & "' and wipdate= to_date('" & request("wipdate_"&TRIM(n)) & "','MM/DD/YYYY') and AGREEMENTFINALNO= " & request("agreementfinalno_"&TRIM(n)) & "" 
							'response.write q + "<br>"
							connStr.execute(q)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & q
							ExecuteAndLog modulFile,executeQry
							
							'untuk update kolom isinvoice
							q ="update workinprogressfinal_new set isinvoice=1 where agreementcode='" & agreementcode & "' and fcba='" & fcba & "' and wipdate= to_date('" & request("wipdate_"&TRIM(n)) & "','MM/DD/YYYY') and AGREEMENTFINALNO= " & request("agreementfinalno_"&TRIM(n)) & "" 
							'response.write q + "<br>"
							connStr.execute(q)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & q
							ExecuteAndLog modulFile,executeQry

						' Untuk menghhitung jumlah total di header
						dtl_amountppn = 0
						dtl_amountpph = 0
						dtl_grandtotal_aftertax = 0
						' dtl_charge = 0
						set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
						q="select ppn, amountppn, pph, amountpph from contractorinvoicehdr " & _
						  " where invoicecode='" & invoicecode & "' and " & _
						  " fcBA='" & fcBA & "' and agreementcode = '"& agreementcode &"'"


						rsTempPurchaseRequest.Open q, connstr
						if Not rsTempPurchaseRequest.Eof then
							while Not rsTempPurchaseRequest.Eof 
								amountppn = cdbl(to_nol(rsTempPurchaseRequest("amountppn"))) 
								amountpph = cdbl(to_nol(rsTempPurchaseRequest("amountpph")))
								ppn = cdbl(to_nol(rsTempPurchaseRequest("ppn")))
								pph = cdbl(to_nol(rsTempPurchaseRequest("pph")))
							rsTempPurchaseRequest.movenext
						wend 
							
						End if
						
						rsTempPurchaseRequest.Close
						set rsTempPurchaseRequest = Nothing

						set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
						q="select SUM(amountafter_retention) amountafter_retention from contractorinvoicedtl " & _
						  " where invoicecode='" & invoicecode & "' and " & _
						  " fcBA='" & fcBA & "' and agreementcode = '"& agreementcode &"'"

						rsTempPurchaseRequest.Open q, connstr
						if Not rsTempPurchaseRequest.Eof then
							while Not rsTempPurchaseRequest.Eof 
								amountafter_retention = cdbl(to_nol(rsTempPurchaseRequest("amountafter_retention")))
							rsTempPurchaseRequest.movenext
						wend 
							
						End if
						
						rsTempPurchaseRequest.Close
						set rsTempPurchaseRequest = Nothing

						dtl_amountppn = amountafter_retention * ppn / 100
						dtl_amountpph = amountafter_retention * pph / 100
						dtl_grandtotal_aftertax = amountafter_retention + dtl_amountppn - dtl_amountpph
						'end add
					   q= " Update contractorinvoicehdr set totalamount_afterretention = replace('" & amountafter_retention & "',',',''), " & _
					   	  " amountppn = replace('" & dtl_amountppn & "',',',''), " & _
					   	  " amountpph = replace('" & dtl_amountpph & "',',',''), " & _
					   	  " grandtotal_aftertax = replace('" & dtl_grandtotal_aftertax & "',',','')" & _
						  " where  invoicecode='" & invoicecode & "' and agreementcode= '"& agreementcode &"' and fcBA='" & fcba & "'"
						connStr.execute(q)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry
						 
					end if
					Next
				end if
		 if err.number  = 0 then
			 connStr.CommitTrans
		  end if
		
	    Call post_amount
		cmdView	= "True"
		
	elseif Request("txtJenis")="Update" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, transdate)
		If Cek_IsDailyClosing = "FALSE" then
			
		
			response.write("Masuk update")
			 Call post_amount
			
		else
			err.number = 1001
			err.description = "Tanggal """ & transdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		cmdView	= "True"
	else
		cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
	cmdView	= request("cmdView")
End if

' Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

'untuk menampilkan datepicker default
' If Len(Trim(transdate)) = 0 Then	transdate = textDate(date)

Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	= "	SELECT hdr.invoicecode, hdr.fcba, hdr.transdate ,hdr.agreementcode, bu.fccompanycode " & _
						  " ,hdr.contractorcode, hdr.isretention_invoice, hdr.retention, hdr.totalamount_afterretention, hdr.ppn" & _
						  " ,hdr.amountppn, hdr.pph, hdr.amountpph, hdr.grandtotal_aftertax, hdr.description, hdr.rowstate, hdr.invoice_by, cr.fcname, hdr.dp" & _
						  " FROM contractorinvoicehdr hdr" & _
						  "       LEFT JOIN businessunit bu ON bu.fccode = hdr.fcba " & _
						  " 	  LEFT JOIN contractor cr ON (cr.fccode = hdr.contractorcode AND cr.fcba = hdr.fcba)" & _
						  " WHERE hdr.invoicecode = '"& invoicecode &"'"
							
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	invoicecode = rsSupplyingDeclaration("invoicecode")
	invoice_by = rsSupplyingDeclaration("invoice_by")
	selectedinvoice = rsSupplyingDeclaration("invoice_by")
	transdate  = rsSupplyingDeclaration("transdate")
	fccompanycode = rsSupplyingDeclaration("fccompanycode")	
	fcba = rsSupplyingDeclaration("fcba")
	agreementcode = rsSupplyingDeclaration("agreementcode")	
	selectedagreement = rsSupplyingDeclaration("agreementcode")	
	contractorcode = rsSupplyingDeclaration("contractorcode")
	contractorname = rsSupplyingDeclaration("fcname")
	isretention_invoice = rsSupplyingDeclaration("isretention_invoice")
	retention = rsSupplyingDeclaration("retention")
	description = rsSupplyingDeclaration("description")
	rowstate = rsSupplyingDeclaration("rowstate")
	selectedfcba = rsSupplyingDeclaration("fcba")
	totalamount_afterretention = formatnumber(cdbl(to_nol(rsSupplyingDeclaration("totalamount_afterretention"))),0)
	ppn = rsSupplyingDeclaration("ppn")
	amountppn = formatnumber(cdbl(to_nol(rsSupplyingDeclaration("amountppn"))),0)
	pph = rsSupplyingDeclaration("pph")
	amountpph = formatnumber(cdbl(to_nol(rsSupplyingDeclaration("amountpph"))),0)
	grandtotal_aftertax = formatnumber(cdbl(to_nol(rsSupplyingDeclaration("grandtotal_aftertax"))),0)
	dp  = rsSupplyingDeclaration("dp")
	selecteddp =  rsSupplyingDeclaration("dp")
	SudahClose = "Y"
else
	'invoicecode = invoicecode
	totalamount_afterretention = 0
	amountppn = 0
	amountpph = 0
	grandtotal_aftertax = 0
	rowstate = ""
End if
 'response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

 if err.number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
 End if
 '-------------ISI Payment Form
 
 Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcba") & "'" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "      /* and to_date ('" & transdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1)) */" & _
	        "       and upper(system) = upper('Costbook')"
	  ' response.write oSQLQuery
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
'response.Write "periode::" & val_IsTrue
oRecordSet.Close
Set oRecordSet = Nothing
 %>
 <form name="frmEntry" method="post" action="contractorinvoice.asp">
  
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="invoicecode" size="30" onkeyup="hurufbesar(this);" value="<%=invoicecode%>" autocomplete="off"><font color="#FF0000">&nbsp;*</font></td>
			<td align="left">Total Amount After Retention Deduction</td>
				<td align="left">
					<input type="text" size="20" id="totalamount_afterretention" name="totalamount_afterretention" value="<%=totalamount_afterretention%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly >
				</td>
		</tr>
		<tr>
			<td align="left">Creation Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" id="transdate" name="transdate" size="10" value="<%=transdate%>">
			<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left" style="width:20%;">Ppn (%)</td>
				<td align="left">
					<input type="text" maxlength="50" id="ppn" name="ppn" size="10" value="<%=ppn%>" autocomplete="off" onkeypress="return hanyaAngka(event);">
					<font color="#FF0000">&nbsp;%</font>
				</td>
		</tr>
		<tr>
			<td align="left" width="20%">Company
				<td align="left">
					<select id="fccompanycode" name="fccompanycode" class="chosen-select" style="width:30%;" onchange="javascript:getlistfcba();javascript:getlistinvoice();">
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
				</td>
			</td>
			<td align="left">Amount Ppn</td>
				<td align="left">
					<input type="text" size="20" id="amountppn" name="amountppn" value="<%=amountppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly>
				</td>
		</tr>
		<tr>
			<td align="left" width="20%">Business Unit
				<td align="left">
					<div id="fcbalist"></div>
				</td>
					<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			</td>
			<td align="left" style="width:20%;">Pph 23</td>
				<td align="left">
					<input type="text" maxlength="50" id="pph" name="pph" size="10" value="<%=pph%>" autocomplete="off" onkeypress="return hanyaAngka(event);">
					<font color="#FF0000">&nbsp;%</font>
				</td>
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By
				<td align="left">
					<div id="invoicelist"></div>
				</td>
					<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="<%=selectedinvoice%>">
			</td>
			<td align="left">Amount Pph 23</td>
				<td align="left">
					<input type="text" size="20" id="amountpph" name="amountpph" value="<%=amountpph%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly>
				</td>
		</tr>
		<tr>
			<td align="left" width="20%">Agreement Code
				<td align="left">
				 	<div id="agreementlist"></div>
				</td>
					<input type="hidden" id="selectedagreement" name="selectedagreement" value="<%=selectedagreement%>">
			</td>
			<td align="left">Grand Total After Tax</td>
				<td align="left">
					<input type="text" size="20" id="grandtotal_aftertax" name="grandtotal_aftertax" value="<%=grandtotal_aftertax%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly>
				</td>
		</tr>
		<tr>
			<td align="left" width="20%">Contractor Code
				<td align="left">
				 	<input type="text" size="10" id="contractorcode" name="contractorcode" value="<%=contractorcode%>" style="background:#ccc" readonly>
				 	<input type="text" size="20" id="contractorname" name="contractorname" value="<%=contractorname%>" style="background:#ccc" readonly>
				</td>
			</td>
			<td align="left" width="20%">
				Down Payment
			</td>	
			<td align="left">
			
				<select name="dp" id="dp" style=width:100px onchange="getSelectdp();">
				   <option value="<%=dp%>"><%=selecteddp%></option>
				</select><% response.write dp %>
			</td>
		</tr>
		
		<tr>
			<td align="left" width="20%">Is Retention Invoice
				<td align="left">
					<select id="isretention_invoice" name="isretention_invoice" class="chosen-select" style="width:30%;" >
						<option value="FALSE" <% If cstr(""&isretention_invoice) = "FALSE" Then %> Selected <%End if%>>FALSE</option> 
						<option value="TRUE" <% If cstr(""&isretention_invoice) = "TRUE" Then %> Selected <%End if%>>TRUE</option> 
					</select>
				</td>
			</td>
			<td align="left" width="20%">
				Grand Total
			</td>	
			<td align="left">
				<input type="text" name="grandtotal" id="grandtotal" value=0 data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly><%=grandtotal%></input>
			</td>
			
		</tr>
		<tr>
			<td align="left" style="width:20%;">Retention</td>
				<td align="left">
					<input type="text" maxlength="50" id="retention" name="retention" size="10" value="<%=retention%>" autocomplete="off">
					<font color="#FF0000">&nbsp;%</font>
				</td>
			<td align="left" width="20%">
				Description
			</td>	
			<td align="left">
				<textarea name="description" id="description"><%=description%></textarea>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="javascript:launchwin('Printcontractorinvoice.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=invoicecode%>&DATECREATED=<%=DATECREATED%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" Then%>
				<a href="contractorinvoice.asp?invoicecode=<%=invoicecode%>&agreementcode=<%=agreementcode%>&fcba=<%=fcba%>&transdate=<%=transdate%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete :  <%=invoicecode%>');"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End if%>
				<a href="contractorinvoice.asp?invoicecode=<%=invoicecode%>&agreementcode=<%=agreementcode%>&fcba=<%=fcba%>&transdate=<%=transdate%>&txtJenis=Posting" onClick="return confirm('Are you sure to Posting Journal this data :  <%=invoicecode%> ?');">Posting</a>&nbsp;&nbsp;
				
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Purchase"
				dataSubmodule 	= "Invoice"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>&nbsp;&nbsp;
				<!-- END:HADIR -->
				<a href="contractorinv_edit.asp?invoicecode=<%=invoicecode%>&fcba=<%=fcba%>" target="_blank"><font face="Verdana" size="1">Edit Header</font></a>&nbsp;&nbsp;
			</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(invoicecode)) Then
   Dim rsGRN,intRecordsPerPage,cColor
   if SudahClose="Y" Then
      q =   " SELECT b.fcba, " & _
	        "        b.invoicecode, " & _
	        "        b.agreementcode, " & _
			"        b.contractorcode, " & _ 
			"        to_char(a.wipdate,'MM/DD/YYYY') wipdate, " & _
			"        a.jobcode, " & _
			"        jb.fcname jobname," & _
			"        a.totalfinal qty_final, " & _
			"        a.rate, " & _
			"        a.amount, " & _
			"        a.amountretention, " & _
			"        a.amountafter_retention, " & _ 
			"        a.Remarks bappno, " & _ 
			"        0 AGREEMENTFINALNO " & _ 
			" FROM contractorinvoicedtl a " & _
			" INNER JOIN contractorinvoicehdr b on (a.fcba=b.fcba and a.invoicecode=b.invoicecode)" & _
			" LEFT JOIN contractor cr ON (cr.fccode = b.contractorcode AND cr.fcba = b.fcba)" & _
			" LEFT JOIN job jb ON (jb.fccode = a.jobcode AND jb.fcba = a.fcba)" & _ 
		    " WHERE b.fcba = '"& fcba &"' AND b.agreementcode = '"& agreementcode &"' AND b.invoicecode='" & TRIM(invoicecode)& "'"
	
   Else	 
	
        if cstr(""&isretention_invoice) = "FALSE" then
		' q = " select wf.AGREEMENTFINALNO,cad.LINENO, ca.agreementcode, to_char(wf.wipdate,'MM/DD/YYYY') wipdate,jb.fccode JobCode, jb.fcname jobname, ca.retention, ca.contractorcode, " & _
			' " cad.rate,SUM (wf.qty) qty_final, SUM (wf.unit) unit_final, SUM (wf.output) output_final " & _
			' " from workinprogressfinal wf " & _
			' " INNER JOIN contractagreement ca on (ca.agreementcode = wf.agreementcode AND ca.fcba = wf.fcba)" & _
			' " INNER JOIN contractagreementdetail cad ON (ca.agreementcode = cad.agreementcode AND ca.fcba = cad.fcba and cad.lineno=wf.AGREEMENTLINENO) " & _
			' " inner join job jb ON (jb.fccode = cad.jobcode AND jb.fcba = cad.fcba) " & _
			' " where wf.fcba = '"& fcba &"' AND wf.agreementcode = '" & agreementcode & "' AND wf.isinvoice=0 " & _
			' " group by wf.AGREEMENTFINALNO,cad.LINENO,jb.fccode, jb.fcname,ca.agreementcode, to_char(wf.wipdate,'MM/DD/YYYY'), ca.retention, ca.contractorcode,cad.rate " & _
			' " order by wf.AGREEMENTFINALNO,to_char(wf.wipdate,'MM/DD/YYYY')asc"
		
		'Update AMH Mar 24
		q = " select wf.AGREEMENTFINALNO,cad.LINENO, ca.agreementcode, to_char(wf.wipdate,'MM/DD/YYYY') wipdate,jb.fccode JobCode, jb.fcname jobname, ca.retention, ca.contractorcode, " & _
			" wf.rate_final rate,SUM (wf.qty) qty_final, SUM (wf.unit) unit_final, SUM (wf.output) output_final ,wf.bappno " & _
			" from workinprogressfinal wf " & _
			" INNER JOIN contractagreement ca on (ca.agreementcode = wf.agreementcode AND ca.fcba = wf.fcba)" & _
			" INNER JOIN contractagreementdetail cad ON (ca.agreementcode = cad.agreementcode AND ca.fcba = cad.fcba and cad.lineno=wf.AGREEMENTLINENO) " & _
			" INNER JOIN jobmaster jb ON (jb.fccode = cad.jobcode ) " & _
			" where wf.fcba = '"& fcba &"' AND wf.agreementcode = '" & agreementcode & "' AND wf.isinvoice=0 " & _
			" group by wf.AGREEMENTFINALNO,cad.LINENO,jb.fccode, jb.fcname,ca.agreementcode, to_char(wf.wipdate,'MM/DD/YYYY'), ca.retention, ca.contractorcode,wf.rate_final,wf.bappno  " & _
			" order by wf.AGREEMENTFINALNO,to_char(wf.wipdate,'MM/DD/YYYY')asc"
		
		else
		q = " SELECT cid.invoicecode,cad.LINENO,cid.agreementcode,cid.wipdate,cid.jobcode,jb.fcname jobname,ca.contractorcode,cid.TOTALFINAL qty_final,cid.RATE,cid.amountretention,cid.Remarks bappno,0 AGREEMENTFINALNO " & _
			" FROM CONTRACTORINVOICEDTL cid " & _
			" INNER JOIN contractagreement ca on (ca.agreementcode = cid.agreementcode AND ca.fcba = cid.fcba)" & _
			" INNER JOIN contractagreementdetail cad ON (ca.agreementcode = cad.agreementcode AND ca.fcba = cad.fcba and cad.lineno=cid.AGREEMENTLINENO) " & _
			" inner join job jb ON (jb.fccode = cid.jobcode AND jb.fcba = cid.fcba) " & _
			" WHERE cid.FCBA='"& fcba &"' AND cid.agreementcode= '" & agreementcode & "' " & _
			" ORDER BY cid.INVOICECODE asc"
		end if

   End if
 '  response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.Open q, connstr
'    if not rsGRN.eof then
'    While Not rsGRN.EOF
' 		agreementcode = rsGRN("agreementcode")
' 		contractorcode = rsGRN("contractorcode")
' 		JobCode = rsGRN("JobCode")
' 		qty_final = rsGRN("qty_final")
' 		rate = rsGRN("rate")
' 		rsGRN.movenext
' 	wend
' end if
   %>
   <form action="contractorinvoice.asp" method="post" name="eForm">
   <table class="data" border="1" width="100%">
   <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<td align="center">No</td>
		<%  if cstr(""&isretention_invoice) = "TRUE" then %>
		<td align="center">Invoice No</td>
		<%END IF%>
		<td align="center">BAPP No</td>
		<td align="center">Agreement Code</td>
		<td align="center">Tgl.WIP</td>
		<td align="center">Jobcode</td>
		<td align="center">JobName</td>
		<td align="center">Total Quantity Final</td>
		<td align="center">Rate</td>
		<td align="center">Amount</td>
		<td align="center">Amount Retention</td>
		<td align="center">Amount After Retention Deduction</td>
	    <td align="center"><input type="checkbox" name="select-all" id="select-all"/> Check/Uncheck All</td>
	   </tr>
   <%ELSE%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	    <td align="center">No</td>
		<td align="center">Invoice No</td>
		<td align="center">BAPP No</td>
		<td align="center">Agreement Code</td>
		<td align="center">Tgl.WIP</td>
		<td align="center">Jobcode</td>
		<td align="center">JobName</td>
		<td align="center">Total Quantity Final</td>
		<td align="center">Rate</td>
		<td align="center">Amount</td>
		<td align="center">Amount Retention</td>
		<td align="center">Amount After Retention Deduction</td>
	   </tr>
   <%End if
   if Not rsGRN.eof then
      nCountofField=rsGRN.fields.count-1
	  no=1
	  While Not rsGRN.EOF
	  		
	  		agreementcode = rsGRN("agreementcode")
	  		contractorcode = rsGRN("contractorcode")
	  		JobCode = rsGRN("JobCode")
	  		qty_final = rsGRN("qty_final")
	  		
	  		' amount = rsGRN("amount")
	  		rate = rsGRN("rate")
			
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if

		  if SudahClose <> "Y" Then
		  if cstr(""&isretention_invoice) = "FALSE" then
		  		amount = cdbl(to_nol(qty_final)) * cdbl(to_nol(rate))
		  		amountretention = cdbl(to_nol(amount)) * cdbl(to_nol(retention)) / 100
		  		amountafter_retention = cdbl(to_nol(amount)) - cdbl(to_nol(amountretention))
		  else
		  		amount = 0
			  
				
				amount = cdbl(to_nol(qty_final)) * cdbl(to_nol(rate))
		  		'amountretention = cdbl(to_nol(amount)) * cdbl(to_nol(retention)) / 100
				amountretention = rsGRN("amountretention")
				amountafter_retention = cdbl(to_nol(amount)) - cdbl(to_nol(amountretention))
				'amountafter_retention=0
				
		  end if
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="ringht" valign="top"><%=NO%></td>
			<%IF SudahClose="Y" then%>
				
				<% if cstr(""&isretention_invoice) = "FALSE" then%>
					<td align="left" title="invoicecode" valign="top"><%=rsGRN("invoicecode")%></td>
					<td align="left" title="bappno" valign="top"><%=rsGRN("bappno")%></td>
					<td align="left" title="agreementcode" valign="top"><%=rsGRN("agreementcode")%></td>
					<td align="left" title="agreementdate" valign="top"><%=rsGRN("wipdate")%></td>
					<td align="left" title="JobCode" valign="top"><%=rsGRN("JobCode")%></td>
					<td align="left" title="JobName" valign="top"><%=rsGRN("JobName")%></td>
					<td align="right" title="qty_final" valign="top"><%=rsGRN("qty_final")%></td>
					<td align="right" title="rate" valign="top"><%=FormatNumber(rsGRN("rate"),0)%></td>				
					<td align="right" title="Amount" valign="top"><%=FormatNumber(rsGRN("amount"),0)%></td>
					<td align="right" title="amountretention" valign="top"><%=FormatNumber(rsGRN("amountretention"),0)%></td>
				<%else%>
					<td align="left" title="invoicecode" valign="top"><%=rsGRN("invoicecode")%></td>
					<td align="left" title="bappno" valign="top"><%=rsGRN("bappno")%></td>
					<td align="left" title="agreementcode" valign="top"><%=rsGRN("agreementcode")%></td>
					<td align="left" title="agreementdate" valign="top"><%=rsGRN("wipdate")%></td>
					<td align="left" title="JobCode" valign="top"><%=rsGRN("JobCode")%></td>
					<td align="left" title="JobName" valign="top"><%=rsGRN("JobName")%></td>
					<td align="right" title="qty_final" valign="top">0</td>
					<td align="right" title="rate" valign="top">0</td>
					<td align="right" title="Amount" valign="top">0</td>
					<td align="right" title="amountretention" valign="top"><%=FormatNumber(rsGRN("amountafter_retention"),0)%></td>
				<%End if%>
				<td align="center">
					<input style="text-align:right;background-color:#eee;" type="text" name="amountafter_retention_<%=no%>" class="amount" value = "<%=FormatNumber(amountafter_retention,0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
				</td>
			<%End if%>

			<%if SudahClose<>"Y" Then%>
				<% if cstr(""&isretention_invoice) = "FALSE" then%>
					<td align="left" title="bappno" valign="top"><%=rsGRN("bappno")%></td>
					<td align="left" title="agreementcode" valign="top"><%=rsGRN("agreementcode")%></td>
					<td align="left" title="agreementdate" valign="top"><%=rsGRN("wipdate")%></td>
					<td align="left" title="JobCode" valign="top"><%=rsGRN("JobCode")%></td>
					<td align="left" title="JobName" valign="top"><%=rsGRN("JobName")%></td>
					<td align="right" title="qty_final" valign="top"><%=rsGRN("qty_final")%></td>
					<td align="right" title="rate" valign="top"><%=FormatNumber(rsGRN("rate"),0)%></td>				
					<td align="right" title="Amount" valign="top"><%=FormatNumber(amount,0)%></td>
					<td align="right" title="amountretention" valign="top"><%=FormatNumber(amountretention,0)%></td>
					<td align="center">
					<input style="text-align:right;background-color:#eee;" type="text" name="amountafter_retention_<%=no%>" class="amount" value = "<%=FormatNumber(amountafter_retention,0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
					</td>
				<%else%>
					<td align="left" title="invoicecode" valign="top"><%=rsGRN("invoicecode")%></td>
					<td align="left" title="bappno" valign="top"><%=rsGRN("bappno")%></td>
					<td align="left" title="agreementcode" valign="top"><%=rsGRN("agreementcode")%></td>
					<td align="left" title="agreementdate" valign="top"><%=rsGRN("wipdate")%></td>
					<td align="left" title="JobCode" valign="top"><%=rsGRN("JobCode")%></td>
					<td align="left" title="JobName" valign="top"><%=rsGRN("JobName")%></td>
					<td align="right" title="qty_final" valign="top">0</td>
					<td align="right" title="rate" valign="top">0</td>
					<td align="right" title="Amount" valign="top">0</td>
					<td align="right" title="amountretention" valign="top"><%=FormatNumber(rsGRN("amountretention"),0)%></td>
					<td align="center">
					<input style="text-align:right;background-color:#eee;" type="text" name="amountafter_retention_<%=no%>" class="amount" value = "0" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
					</td>
				<%End if%>
				
				<td align="center">
					<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
				</td>
					<input type="hidden" name="jobcode_<%=no%>" value="<%=rsGRN("jobcode")%>">
					<input type="hidden" name="qty_final_<%=no%>" value="<%=rsGRN("qty_final")%>">
					<input type="hidden" name="rate_<%=no%>" value="<%=rsGRN("rate")%>">
					<input type="hidden" name="amount_<%=no%>" value="<%=amount%>">
					<input type="hidden" name="amountretention_<%=no%>" value="<%=amountretention%>">
					<input type="hidden" id="wipdate_<%=no%>" name="wipdate_<%=no%>" value="<%=rsGRN("wipdate")%>">
					<input type="hidden" id="agreementlinenotxt<%=no%>" name="agreementlinenotxt<%=no%>" value="<%=rsGRN("LINENO")%>">
					<input type="hidden" name="bappno_<%=no%>" value="<%=rsGRN("bappno")%>">
					<input type="hidden" name="agreementfinalno_<%=no%>" value="<%=rsGRN("AGREEMENTFINALNO")%>">
					
					
				 </td>
				
			<%End if%>
		  </tr>
		  <%
		  no=no+1
		  rsGRN.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsGRN.Close
   set rsGRN = Nothing%>
	  <tr><td colspan="10" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
	      <input type="hidden" name="invoicecode" value="<%=invoicecode%>">
	      <input type="hidden" name="agreementcode" value="<%=agreementcode%>">
	      <input type="hidden" name="selecteddp" id="selecteddp" value="<%=dp%>">
	      <input type="hidden" name="selectedagreement" value="<%=selectedagreement%>">
	      <input type="hidden" name="pph" value="<%=pph%>">
		  <input type="hidden" name="ppn" value="<%=ppn%>">
		  <input type="hidden" name="grandtotal_aftertax" value="<%=grandtotal_aftertax%>">
		  <input type="hidden" name="No" value="<%=No%>">
		  <input type="hidden" name="selectedinvoice" value="<%=selectedinvoice%>">
		  <input type="hidden" name="invoice_by" value="<%=invoice_by%>">
		  <input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
	  	  <input type="hidden" name="fcba" value="<%=fcba%>">
		  <input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		  <input type="hidden" name="amountppn" value="<%=amountppn%>">	
		  <input type="hidden" name="amountpph" value="<%=amountpph%>">
		  <input type="hidden" name="rowstate" value="<%=rowstate%>">
		  <input type="hidden" name="isretention_invoice" value="<%=isretention_invoice%>">
		  <input type="hidden" name="retention" value="<%=retention%>">
		  <input type="hidden" name="description" value="<%=description%>">
		  <input type="hidden" name="totalamount_afterretention" value="<%=totalamount_afterretention%>">
		  <input type="hidden" name="contractorcode" value="<%=contractorcode%>">
		  <input type="hidden" name="contractorname" value="<%=contractorname%>">
		  <input type="hidden" name="transdate" value="<%=transdate%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%if SudahClose<>"Y" Then%>
		  <%If val_IsTrue = "Y" Then%>
			  <%IF Request("txtJenis")="Edit" then%>
				   <input type="hidden" name="txtJenis" value="Edit">
				   <input Type="Submit"  name="Add" value="Update">
			  <%Else%>
				   <input type="hidden" name="txtJenis" value="New">
				   <input Type="Submit"  name="Add" value="Submit">
			  <%End if%>
		  <%End if%>
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
	<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
	<script src="/js/select2/select2.min.js" type="text/javascript"></script>
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
<!-- <script type="text/javascript">	
	$('#select-all').click( function () {
		$('.item').prop('checked', this.checked);
		$('.item').val(this.checked ? 1 : 0);
	})
</script> -->
<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
}

function getselectedinvoice() {
	var sfcba = document.getElementById("selectedinvoice");
	sfcba.value = document.getElementById("invoice_by").value;
}

function getselectedagreement() {
	var spocode = document.getElementById("selectedagreement");
	spocode.value = document.getElementById("agreementcode").value;
	// alert("vouchersource");
}




Document.prototype.ready = function(callback) {
	
  if(callback && typeof callback === 'function') {
    document.addEventListener("DOMContentLoaded", function() {
		
      if(document.readyState === "interactive" || document.readyState === "complete") {
		//alert('tes');
		
		getDP();
		// var agreementcode = $('#selectedagreement').val();
		// //var agreementcode = document.getElementById("selectedagreement");
		// alert(agreementcode);
        return callback();
      }
    });
  }
};

// usage
document.ready(function() {
	getlistfcba();
	getlistinvoice();
	getlistvoucher();
    getselectedagreement();
    getretention();
    hitungrumus();
    getselectedinvoice();
	//getSelectdp();
	
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
		var txt = chk.parentNode.parentNode.cells[8].getElementsByTagName('input')[0];
		
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


	function getDP()
	{
		var agreementcode = $('#selectedagreement').val();
		$('#dp').empty();
				$('#dp').select2({
					ajax: {
						delay: 1000,
						url : 'getlistdpwip.asp?agreementcode=' + agreementcode,
						dataType: 'json',
						type: "POST",
						allowClear: true,
						data: function (params) {
								var query = {
												keyword: params.term
											}
								return query;
						}
					}
				});
	}

	function getSelectdp()
	{
		var dp = $('#dp').val();
		 $('#selecteddp').val(dp);
	}

	function hurufbesar(obj) {
		var strval	= $(obj).val();
		strval		= strval.toUpperCase();
		strval		= strval.replace(/\s/g,"");
		$(obj).val(strval);
	}

	
</script>
</body>
</html>
