<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<%
function f_title(transtype_, banktype_)
	result = ""
	if Ucase(banktype_) = "JOURNAL" then result = "Journal Voucher" else result = PCase(transtype_) & " Voucher"
	
	f_title = result
end function

%>
<html>
<head>
    <title><%=f_title(Request("transtype"), Request("banktype")) & " | ePlaS"%></title>
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
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />
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
<script language="javascript" type="text/javascript">
function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
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

// launched from button click 
var getamountinvoice = function () {
	// var fccompanycode = document.getElementById("fccompanycode").value;
	var jobcode = document.getElementById("jobcode").value;
	var locationcode = document.getElementById("locationcode").value;
	var documentno = document.getElementById("documentno").value;
	var PostStr = 'jobcode='+jobcode+'&locationcode='+locationcode+'&documentno='+documentno;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getAmountInvoicePayment.asp', 'POST', PostStr, showAmountInvoice);
	
}

// The function for handling the response from the server
var showAmountInvoice = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.itemcode);
    // update the Div to show the result from the server
    document.getElementById("amountinvoice").value = s.amountinvoice;
	// document.getElementById("stockname").value = s.stockname;
	// document.getElementById("requestedqty").value = s.approvedqty; 
}

// launched from button click 
var getLastNumber = function () {
	// var fccompanycode = document.getElementById("fccompanycode").value;
	var transtype = document.getElementById("transtype").value;
	var bankcode = document.getElementById("bankcode").value;
	var PostStr = 'transtype='+transtype+'&bankcode='+bankcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getLastVoucher.asp', 'POST', PostStr, showLastNumber);
	
}

// The function for handling the response from the server
var showLastNumber = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.lastvouchercode);
    // update the Div to show the result from the server
    document.getElementById("lastvouchercode").value = s.lastvouchercode;
}

// launched from button click 
var getJobCode = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedjobcode = document.getElementById("selectedjobcode").value;
	var PostStr = 'fccode='+selectedjobcode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getJobCode.asp', 'POST', PostStr, showJobcode);
}

// The function for handling the response from the server
var showJobcode = function (oXML) { 
    
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
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
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
		// document.getElementById("jobcodeintraco").style.display="block";
		document.getElementById("chargeunit").style.display="table-row";
		document.getElementById("chargejob").style.display="table-row";
		document.getElementById("chargetype").style.display="table-row";
		document.getElementById("chargecode").style.display="table-row";
	} else {
		// document.getElementById("jobcodeintraco").style.display="none";
        document.getElementById("chargeunit").style.display="none";
		document.getElementById("chargejob").style.display="none";
		document.getElementById("chargetype").style.display="none";
		document.getElementById("chargecode").style.display="none";
	}
}

$.urlParam = function(name){
    var results = new RegExp('[\?&]' + name + '=([^&#]*)').exec(window.location.href);
    if (results==null) {
       return null;
    }
    return decodeURI(results[1]) || 0;
}

var reloadCode = function (PostStr) {
	
	var transdate =$( ".datepicker" ).val();
	var bankcode =$( "#bankcode" ).val();
	var txtTable = $.urlParam('txtTable');
	var txtBanktype = $.urlParam('banktype');
	var vtranstype = $.urlParam('transtype');
	//alert(bankcode);
	
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Finance and Accounting&txtTable="+txtTable+"&transtype="+vtranstype+"&banktype="+txtBanktype+"&transdate="+transdate+"&bankcodeurl="+bankcode+"&"+PostStr );
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
    // alert(response);
    // update the Div to show the result from the server
    document.getElementById("locationtype").value = s.target; 
	// document.getElementById("descpr").value = s.description; 

}

var getChargetype = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("chargejob_").value;
	var PostStr = 'jobcode='+jobcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetPVInfo.asp', 'POST', PostStr, showChargettype);
	
}

// The function for handling the response from the server
var showChargettype = function (oXML) { 
     // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(response);
    // update the Div to show the result from the server
    document.getElementById("chargetype_").value = s.target; 
	// document.getElementById("descpr").value = s.description; 
}

// tampil nominal invoice 
var getNominal = function () { 
	 var jnsinvoice = document.getElementById("jnsinvoice").value;
	 var selectedinvoice = document.getElementById("noinvoice").value;
	 var PostStr = 'selectnoinv='+selectedinvoice+'&jnsinvoice='+jnsinvoice;
	 //alert(PostStr);
	
	// // use the generic function to make the request
	doAJAXCall('ajax/getnominalviewinv.asp', 'POST', PostStr, shownominalinv);
	
}

var shownominalinv = function (oXML) { 
    
    var response = oXML.responseText;
	var s = JSON.parse(response);
    
    // update the Div to show the result from the server
	document.getElementById("amountinvoice").value = s.nominal;

}

</script>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

    <script language="JavaScript">
        function showDetails(winurl, p_item) {
            window.open(winurl + '&p_item=' + p_item, width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1);
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
</head>

<style>
	
	.lebarsup {
		position: absolute;
        right: 380;
        width: 200px;

	}
	.ltinggi{
		padding-top:-50px;
	}
</style>


<%if request("Add")="" then%>
<body onLoad="document.frmEntry.fcCode.focus()">
    <%Else%>
    <body onload="document.eForm.selectedjobcode.focus()">
        <%End if%>
        <!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
                    <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	  Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">   
    <!-----------------ATAS ISI!----------------->

<%
 '-------------awal-----------------------
 
Function CheckAvailableJobZerorise(initialJob, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from parameter" & _
	                " where upper(parcode) = upper('JOBCODE_ZERORISE')" & _
					"       and fccode = '" & Trim(initialJob) & "'" & _
					"       and fcba = '" & Trim(initialFcba) & "'"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckAvailableJobZerorise = initialAvailable
End Function


Function f_i_check(in_fcba, in_vouchercode, in_datecreated, in_bankcode, in_jobcode, in_targettype, in_targetcode, in_documentno, in_amount)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT validasi_voucher_api.check_data( '"&in_fcba&"', '"&in_vouchercode&"', TO_DATE('"&in_datecreated&"', 'mm/dd/yyyy'), '"&in_bankcode&"', '"&in_jobcode&"', '"&in_targettype&"', '"&in_targetcode&"', '"&in_documentno&"', REPLACE('"&in_amount&"', ',', '') ) result" & _
	                " FROM DUAL"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("result")
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	f_i_check = initialAvailable
End Function


Function CheckCostJobCr(statusxAmount ,initialJob, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from Costjobcr" & _
	                " where  jobcode = '" & Trim(initialJob) & "'" & _
					" and fcba = '" & Trim(initialFcba) & "'"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	IF statusxAmount = "MINUS"  THEN
			If Not rsAvailable.EOF Then
				initialAvailable = Cdbl(rsAvailable("Total"))
			Else
				initialAvailable = 0
			End If
	 Else
	    initialAvailable = 0
	 END IF
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckCostJobCr = initialAvailable
End Function

Function CheckCharge(initType, initCode, initJob, initFcba, initDate, initSource)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT F_CekTarget_byDate ('"& initType &"'," & _
					"                           '"& initCode &"'," & _
					"                           '"& initJob &"'," & _
					"                           '"& initFcba &"'," & _
					"                           TO_DATE ('"& initDate &"', 'MM/DD/YYYY')," & _
					"                           'BA') Total" & _
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

 Dim check

 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,SourceType
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks,TxtTType
 Dim rsInvoiceRecord, varTotalAmount, invoiceUpdate, recordcheck, rsrecordlineno, recordWIPLINE,invoiceRecord
 Dim rsPaymentVoucherDetail, varamountOriginal, recordinvoicepaid, vartotal, varError, ivp,varmaxlineno,recordlineno
 Dim Jobcode, LocationType, LocationCode, Reference, Remark,InvoiceCode,AGREEMENTCODE,AGREEMENTLINENO ,recordwipcheck
 Dim varSupplierCode, varStringError, varlineno, WIPPAIDRecord, rsWIPPAIDRecord, WIPUpdate ,rsrecordWIPLINE ,W, rsrecordwipcheck, WIPDATE, varContract
 Dim InsertIntoCostActivity, rsqMaxLineNo, qMaxLineNo, MaxLineNo, IsCostbook
 Dim MaxLineNopvd
 Dim rlineno, dataModule, dataSubmodule, dataDetail, qq, recPaymentVoucherDetail
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD
 Dim xACCOUNTINGYEARPERIOD
 DIM xAmount, StatusXAmount
 Dim sqljnsinv,rsjnsinv,sqlupdatepay,txtnoinvoice,jenisinv,amountinvoice

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 ' if request("fccOde")="" then
 '    fcCode="" '----- Nantinya Ambil dari Format Master
 ' else
 '    fcCode = trim(request("fccOde"))
 ' end if
 

 'fcCode = request("fccOde")
 'If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if request("transdate") <>"" then
   DATECREATED   = request("transdate")
 elseif Request("DATECREATED")<>"" then
   DATECREATED = Request("DATECREATED")
 else
   DATECREATED = textDate(date)
 end if
 
 
 if request("bankcodeurl") <>"" then
   xCode   = Request("bankcodeurl")
   'response.write xCode
 elseif Request("xCode")<>"" then
   xCode = Request("xCode")
 else
   xCode = Request("xCode")
 end if
 
 
 HeaderRemarks = Request("HeaderRemarks")
 
 '---- Data Details
'JobCode=Request("JobCode")
JobCode=Request("selectedjobcode")
LocationType=Request("LocationType")
LocationCode=Request("LocationCode")
Reference=Request("Reference")
Remarks=Request("Remarks")
Amount=Request("Amount")
InvoiceCode=Request("InvoiceCode")
AGREEMENTCODE=Request("AGREEMENTCODE")
AGREEMENTLINENO=Request("AGREEMENTLINENO")
WIPDATE=Request("WIPDATE")
rlineno = Request("txtlineno")
documentno = Request("documentno")
document_classification = Request("document_classification")
banktype = Request("banktype")
transtype = Request("transtype")

' Added Variable ChargeJob, ChargeType, ChargeCode
chargeunit=Request("chargeunit")
chargejob=Request("chargejob")
chargetype=Request("chargetype")
chargecode=Request("chargecode")
'businessunit=Request("businessunit")
businessunit=Request("chargeunit")
referenceno=Request("referenceno")
referenceno=Request("referenceno")
cekno=Request("cekno")
fcCode = Request("fcCode")
txtnoinvoice = Request("noinvoice")
jenisinv = Request("jnsinvoice")
'amountinvoice = Request("amountinvoice")

' End Added


If Request("Amount") <> ""  then
   xAmount = Replace(Amount, ",", "")
   IF Cdbl(xAmount) > 0 then
      StatusXAmount = "PLUS"
   Else
      StatusXAmount = "MINUS"       
   End If
End If
'response.write "Amount nya " & xAmount


 ' ACCOUNTINGYEARPERIOD = Request("ACCOUNTINGYEARPERIOD")
 ' if "" & ACCOUNTINGYEARPERIOD <> "" then
    ' ACCOUNTINGYEAR = MID(ACCOUNTINGYEARPERIOD,1,4)
	' ACCOUNTINGPERIOD = MID(ACCOUNTINGYEARPERIOD,5,2)
 ' end if
 
 'response.write "ACCOUNTINGYEAR " & ACCOUNTINGYEAR  & "ACCOUNTINGPERIOD " & ACCOUNTINGPERIOD 

  
 if Request("SourceType")="" then
    SourceType="BA" 'Default if Entry is Blank
 else
    SourceType=Request("SourceType")
 End if
dim tt

Sub DeleteDataPaymentVoucherDetail
	q = " DELETE FROM paymentvoucherdetail" & _
		" WHERE fcBA = '" & Session("fcBA") & "'" & _
		"       AND vouchercode = '" & fcCode & "'" & _
		"       AND lineno = '" & rlineno & "'"
	connStr.execute(q)
End Sub

Sub InsertDataPaymentVoucherDetail

	qLineMax = " select count(1) fccode, fcba from intraco_tab where status='TRUE' and fccode = '"& jobcode &"' group by fcba"			
	' response.write qLineMax
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		code = cdbl(to_nol(rsqLineMax("fccode")))
		businessunit = rsqLineMax("fcba")
		businessunit = chargeunit
	end if 

	if code = 0 then
		if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 or len(trim(""&chargeunit)) <> 0 ) then
			err.number = 1001
			err.description = "Anda tidak harus input Charge Job, Charge Type dan Charge Code! "
		else
		qq = " INSERT INTO paymentvoucherdetail ( vouchercode" & _
		     "                                    , sourcetype" & _
		     "                                    , locationtype" & _
		     "                                    , locationcode" & _
		     "                                    , jobcode" & _
		     "                                    , chargetype" & _
		     "                                    , chargecode" & _
		     "                                    , chargejob" & _
			 "                                    , chargeunit" & _
		     "                                    , amount" & _
		     "                                    , reference" & _
		     "                                    , remarks" & _
		     "                                    , fcentry" & _
		     "                                    , fcedit" & _
		     "                                    , fcip" & _
		     "                                    , fcba" & _
		     "                                    , invoicecode" & _
		     "                                    , agreementcode" & _
		     "                                    , agreementlineno" & _
		     "                                    , wipdate" & _
		     "                                    , lineno" & _
		     "                                    , lastupdate" & _
		     "                                    , lasttime " & _
			 "                                    , jenisinv) " & _
		     " VALUES ( '" & fcCode & "'" & _
	         "          , '" & SourceType & "'" & _
	         "          , '" & LocationType & "'" & _ 
	         "          , '" & LocationCode &  "'" & _
	         "          , '" & Jobcode & "'" & _
	         "          , '" & chargetype & "'" & _ 
	         "          , '" & chargecode &  "'" & _
	         "          , '" & chargejob & "'" & _
			 "          , '" & chargeunit & "'" & _
	         "          , REPLACE('" & Amount & "', ',', '')" & _
	         "          , '" & Reference & "'" & _ 
	         "          , '" & Remarks & "'" & _
	         "          , '" & Session("CodeName") & "'" & _
	         "          , '" & Session("CodeName") & "'" & _
	         "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
	         "          , '" & Session("fcBA") & "'" & _
	         "          , '" & InvoiceCode & "'" & _
	         "          , '" & AGREEMENTCODE & "'"& _
	         "          , '" & AGREEMENTLINENO & "'" & _
	         "          , TO_DATE('" & WIPDATE & "', 'mm/dd/yyyy')" & _
	         "          , '" & rlineno & "'" & _
	         "          , SYSDATE" & _
	         "          , TO_CHAR(SYSDATE,'HH24:MI') " & _ 
			 "			, '" & jenisinv & "')"
		' response.write qq
		connStr.execute(qq)
			' Untuk Update Total Amount Di Header
				q = " SELECT NVL(SUM(amount), 0) totalamount" & _
					" FROM paymentvoucherdetail" & _
					" WHERE VoucherCode='" & fcCode & "'" & _
					"       AND fcBA = '" & session("fcBA") & "'"
				set rsVoucher = server.CreateObject("Adodb.recordset")
				rsVoucher.Open q, connstr
				if Not rsVoucher.Eof then
					TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
				End if

				q =	" Update paymentvoucher" & _
					" set totalamount = '"& TOTALAMOUNT &"'" & _
					"     , fcentry = '" & Session("CodeName") & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"     , lastupdate = sysdate " & _
					"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					" 		where vouchercode = '" & fccode & "'" & _
					"       and fcBA = '" & Session("fcBA") & "'"
	 			' response.write q
				connStr.execute(q)
			' End Added
		End If
	else
		
		if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 or len(trim(""&chargeunit)) = 0 ) then
			err.number = 1001
			err.description = "Anda harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
		else
			Dim val_Charge : val_Charge = 0
			val_Charge = CheckCharge(chargetype, chargecode, chargejob, businessunit, DATECREATED, sourcetype)
			If val_Charge <> "Success" Then
				Err.Number = 1001
				Err.Description = val_Charge
			Else
			
				qq = " INSERT INTO paymentvoucherdetail ( vouchercode" & _
				     "                                    , sourcetype" & _
				     "                                    , locationtype" & _
				     "                                    , locationcode" & _
				     "                                    , jobcode" & _
				     "                                    , chargetype" & _
				     "                                    , chargecode" & _
				     "                                    , chargejob" & _
					 "                                    , chargeunit" & _
				     "                                    , amount" & _
				     "                                    , reference" & _
				     "                                    , remarks" & _
				     "                                    , fcentry" & _
				     "                                    , fcedit" & _
				     "                                    , fcip" & _
				     "                                    , fcba" & _
				     "                                    , invoicecode" & _
				     "                                    , agreementcode" & _
				     "                                    , agreementlineno" & _
				     "                                    , wipdate" & _
				     "                                    , lineno" & _
				     "                                    , lastupdate" & _
				     "                                    , lasttime" & _ 
					 " 									  , jenisinv) " & _
				     " VALUES ( '" & fcCode & "'" & _
			         "          , '" & SourceType & "'" & _
			         "          , '" & LocationType & "'" & _ 
			         "          , '" & LocationCode &  "'" & _
			         "          , '" & Jobcode & "'" & _
			         "          , '" & chargetype & "'" & _ 
			         "          , '" & chargecode &  "'" & _
			         "          , '" & chargejob & "'" & _
					 "          , '" & chargeunit & "'" & _
			         "          , REPLACE('" & Amount & "', ',', '')" & _
			         "          , '" & Reference & "'" & _ 
			         "          , '" & Remarks & "'" & _
			         "          , '" & Session("CodeName") & "'" & _
			         "          , '" & Session("CodeName") & "'" & _
			         "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
			         "          , '" & Session("fcBA") & "'" & _
			         "          , '" & InvoiceCode & "'" & _
			         "          , '" & AGREEMENTCODE & "'"& _
			         "          , '" & AGREEMENTLINENO & "'" & _
			         "          , TO_DATE('" & WIPDATE & "', 'mm/dd/yyyy')" & _
			         "          , '" & rlineno & "'" & _
			         "          , SYSDATE" & _
			         "          , TO_CHAR(SYSDATE,'HH24:MI')" & _ 
					 " 			, '" & jenisinv & "')"
				' response.write qq
				connStr.execute(qq)
				
				' Untuk Update Total Amount Di Header
					q = " SELECT NVL(SUM(amount), 0) totalamount" & _
						" FROM paymentvoucherdetail" & _
						" WHERE VoucherCode='" & fcCode & "'" & _
						"       AND fcBA = '" & session("fcBA") & "'"
					set rsVoucher = server.CreateObject("Adodb.recordset")
					rsVoucher.Open q, connstr
					if Not rsVoucher.Eof then
						TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
					End if

					q =	" Update paymentvoucher" & _
						" set totalamount = '"& TOTALAMOUNT &"'" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where vouchercode = '" & fccode & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
		 			' response.write q
					connStr.execute(q)
				' End Added
			End If
		End If
	End If
End Sub

Sub UpdateDataPaymentVoucherDetail
	qLineMax = " select count(1) fccode, fcba from intraco_tab where status='TRUE' and fccode = '"& jobcode &"' group by fcba"			
	' response.write qLineMax
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		code = cdbl(to_nol(rsqLineMax("fccode")))
		'businessunit = rsqLineMax("fcba")
		 businessunit = chargeunit
	end if 
if code = 0 then
		if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 or len(trim(""&chargeunit)) <> 0 ) then
			err.number = 1001
			err.description = "Anda tidak harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
		else
		Call DeleteDataPaymentVoucherDetail
		qq = " INSERT INTO paymentvoucherdetail ( vouchercode" & _
		     "                                    , sourcetype" & _
		     "                                    , locationtype" & _
		     "                                    , locationcode" & _
		     "                                    , jobcode" & _
		     "                                    , chargetype" & _
		     "                                    , chargecode" & _
		     "                                    , chargejob" & _
			 "                                    , chargeunit" & _
		     "                                    , amount" & _
		     "                                    , reference" & _
		     "                                    , remarks" & _
		     "                                    , fcentry" & _
		     "                                    , fcedit" & _
		     "                                    , fcip" & _
		     "                                    , fcba" & _
		     "                                    , invoicecode" & _
		     "                                    , agreementcode" & _
		     "                                    , agreementlineno" & _
		     "                                    , wipdate" & _
		     "                                    , lineno" & _
		     "                                    , lastupdate" & _
		     "                                    , lasttime " & _
		     "                                    , jenisinv) " & _
		     " VALUES ( '" & fcCode & "'" & _
	         "          , '" & SourceType & "'" & _
	         "          , '" & LocationType & "'" & _ 
	         "          , '" & LocationCode &  "'" & _
	         "          , '" & Jobcode & "'" & _
	         "          , '" & chargetype & "'" & _ 
	         "          , '" & chargecode &  "'" & _
	         "          , '" & chargejob & "'" & _
			 "          , '" & chargeunit & "'" & _
	         "          , REPLACE('" & Amount & "', ',', '')" & _
	         "          , '" & Reference & "'" & _ 
	         "          , '" & Remarks & "'" & _
	         "          , '" & Session("CodeName") & "'" & _
	         "          , '" & Session("CodeName") & "'" & _
	         "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
	         "          , '" & Session("fcBA") & "'" & _
	         "          , '" & InvoiceCode & "'" & _
	         "          , '" & AGREEMENTCODE & "'"& _
	         "          , '" & AGREEMENTLINENO & "'" & _
	         "          , TO_DATE('" & WIPDATE & "', 'mm/dd/yyyy')" & _
	         "          , '" & rlineno & "'" & _
	         "          , SYSDATE" & _
	         "          , TO_CHAR(SYSDATE,'HH24:MI')" & _
	         "          , '" & jenisinv & "' )"
			 
        'response.write(qq)		 
		connStr.execute(qq)

		' Untuk Update Total Amount Di Header
			q = " SELECT NVL(SUM(amount), 0) totalamount" & _
				" FROM paymentvoucherdetail" & _
				" WHERE VoucherCode='" & fcCode & "'" & _
				"       AND fcBA = '" & session("fcBA") & "'"
			set rsVoucher = server.CreateObject("Adodb.recordset")
			rsVoucher.Open q, connstr
			if Not rsVoucher.Eof then
				TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
			End if

			q =	" Update paymentvoucher" & _
				" set totalamount = '"& TOTALAMOUNT &"'" & _
				"     , fcentry = '" & Session("CodeName") & "'" & _
				"     , fcedit = '" & Session("CodeName") & "'" & _
				" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"     , lastupdate = sysdate " & _
				"     , lasttime = to_char(sysdate,'hh24:mi') " & _
				" 		where vouchercode = '" & fccode & "'" & _
				"       and fcBA = '" & Session("fcBA") & "'"
 			' response.write q
			connStr.execute(q)
		' End Added
		End If
	Else
		Call DeleteDataPaymentVoucherDetail
		Call InsertDataPaymentVoucherDetail
	end if
End Sub
		
'--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "StartAuthorize" or Request("txtJenis")= "Approved" Then	
	If CheckCostJobCr(StatusXAmount,Jobcode, Session("FCBA")) = 0 then
	
			if Request("txtJenis")= "Header" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then

					q = " Delete from PaymentVoucherDetail where  VoucherCode = '" & fcCode & "' and fcBA = '" & Session("fcBA") & "'"
					connStr.execute(q)
					
					q = " Delete from PaymentVoucher where  VoucherCode = '" & fcCode & "' and fcBA = '" & Session("fcBA") & "'"
					connStr.execute(q)
					
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If  
				
				cmdView	= "True"
			   
			elseif Request("txtJenis")= "Details" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then

					call DeleteDataPaymentVoucherDetail
					
					oSQLQuery = " select nvl(bank_api.get_trigger_system(bankcode, fcba), 'Bank') trigger_system" & _
								" from paymentvoucher" & _
								" where fcba = '" & Session("fcBA") & "'" & _
								"       and vouchercode = '" & fcCode & "'"
					set oRecordSet = server.CreateObject("Adodb.recordset")
					oRecordSet.Open oSQLQuery, connstr
					If Not oRecordSet.Eof then
						IsCostbook = oRecordSet("trigger_system")
					Else
						IsCostbook = "Bank"
					End if
					oRecordSet.Close
					Set oRecordSet = Nothing
					 
					If UCase(IsCostbook) = UCase("Costbook") Then
					
						arr_lineno_eof = true
						oSQLQuery = " SELECT lineno" & _
									" FROM cost_activity" & _
									" WHERE fcba = '" & Session("fcBA") & "'" & _
									"       AND sourcetype = 'BA'" & _
									"       AND referenceno = '" & fcCode & "'" & _
									"       AND auto_reference_05 = '" & rlineno & "'"
						set oRecordSet = server.CreateObject("Adodb.recordset")
						oRecordSet.Open oSQLQuery, connstr
						If Not oRecordSet.Eof then
							arr_lineno = oRecordSet.getrows()
							arr_lineno_eof = false
						End if
						oRecordSet.Close
						Set oRecordSet = Nothing
						
						if not arr_lineno_eof then
							for xi=0 to ubound(arr_lineno,2)
					
								q = " DELETE FROM cost_activity_hist_tab " & _
									" WHERE fcba = '" & Session("fcBA") & "'" & _
									"       AND lineheaderno = '" & arr_lineno(0,xi) & "'"

								connStr.execute(q)
								
							next
						end if
						
						q = " DELETE FROM cost_activity" & _
							" WHERE fcBA = '" & Session("fcBA") & "'" & _
							"       AND sourcetype = 'BA'" & _
							"       AND referenceno = '" & fcCode & "'" & _
							"       AND auto_reference_05 = '" & rlineno & "'"
							
						connStr.execute(q)
					
					End If
					
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If 
				
				cmdView	= "True"
			
			elseif Request("txtJenis")="New" then
				Dim rsPO
				 'response.write "datanya " & Request("ACCOUNTINGYEARPERIOD") 
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then
					
					check_data = f_i_check(Session("fcBA"), fcCode, DATECREATED, xCode, Jobcode, locationtype, locationcode, documentno, amount)
					'--Add by Ukie to handle Auto Number of PR
							If Ucase(fcCode) = Ucase("(NEW)") Then
								oSql = " select nvl(controlnumber_api.get_nextnumber(upper('"& transtype &"')" & _
									   "                                         , " & year(textDate(datecreated)) & "" & _
									   "                                         , " & month(textDate(datecreated)) & "" & _
									   "                                         , '" & session("fcba") & "'" & _
									   " ), 'FALSE') transtype from dual"
								
								' Response.Write oSql
									   
								Set oRs = Server.CreateObject("ADODB.Recordset")
								oRs.Open oSql, connstr
								If Not oRs.EOF Then
									txtvouchercode = oRs("transtype")
								End If
								oRs.Close
								Set oRs = Nothing
									fcCode = txtvouchercode
							End If
							
							If txtvouchercode = "FALSE" Then
								'err.Number = 1001
								'err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
							End If
						
						'Response.Write PRCODE & txtPRCode
						check_data = "Success"
					'--End Added
					if check_data = "Success" then
					
						invoicecode = ""
						agreementcode = ""
					
						if locationtype = "AP" then
							invoicecode = documentno
						elseif locationtype = "CA" then
							agreementcode = documentno
						else
							invoicecode = ""
							agreementcode = ""
						end if
				
						Set rsPO = Server.CreateObject("ADODB.Recordset")
						rsPO.Open " SELECT vouchercode FROM paymentvoucher WHERE vouchercode = '" & fcCode & "' AND fcba = '" & Session("FCBA") & "'", connstr
						IF rsPO.eof then
							
							' q	=	" INSERT INTO paymentvoucher ( vouchercode" & _
									' "                              , referenceno" & _
									' "                              , cekno" & _
									' "                              , bankcode" & _
									' "                              , datecreated" & _
									' "                              , document_classification" & _
									' "                              , accountingyear" & _
									' "                              , accountingperiod" & _
									' "                              , totalamount" & _
									' "                              , suppcontcode" & _
									' "                              , remarks" & _
									' "							   , rowstateid	" & _
									' "							   , automaticapproved	" & _
									' "							   , transtype	" & _
									' "                              , fcentry" & _
									' "                              , fcedit" & _
									' "                              , fcip" & _
									' "                              , fcba" & _
									' "                              , lastupdate" & _
									' "                              , lasttime)" & _
									' " VALUES ( '" & fcCode & "'" & _
									' "          , '" & referenceno & "'" & _
									' "          , '" & cekno & "'" & _
									' "          , '" & xCode & "'" & _
									' "          , TO_DATE('" & DATECREATED & "','mm/dd/yyyy')" & _
									' "          , '" & document_classification & "'" & _
									' "          ," & ACCOUNTINGYEAR & " " &_
									' "          ,TO_NUMBER(" & ACCOUNTINGPERIOD & ")" &_
									' "          , 0" & _
									' "          , ''" & _
									' "          , '" &  request("HeaderRemarks") & "'" & _
									' "          , 'Planned'" & _
									' "          , '1'" & _
									' "          , UPPER('" & transtype & "')" & _
									' "          , '" & Session("CodeName") & "'" & _
									' "          , '" & Session("CodeName") & "'" & _
									' "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									' "          , '" & Session("fcBA") & "'" & _
									' "          , SYSDATE" & _
									' "          , TO_CHAR(SYSDATE,'HH24:MI') )"
									' ' response.Write("q " & q )
									' connStr.execute(q)
									
									q	=	" INSERT INTO paymentvoucher ( vouchercode" & _
									"                              , referenceno" & _
									"                              , cekno" & _
									"                              , bankcode" & _
									"                              , datecreated" & _
									"                              , document_classification" & _
									"                              , accountingyear" & _
									"                              , accountingperiod" & _
									"                              , totalamount" & _
									"                              , suppcontcode" & _
									"                              , remarks" & _
									"							   , rowstateid	" & _
									"							   , automaticapproved	" & _
									"							   , transtype	" & _
									"                              , fcentry" & _
									"                              , fcedit" & _
									"                              , fcip" & _
									"                              , fcba" & _
									"                              , lastupdate" & _
									"                              , lasttime)" & _
									" VALUES ( '" & fcCode & "'" & _
									"          , '" & referenceno & "'" & _
									"          , '" & cekno & "'" & _
									"          , '" & xCode & "'" & _
									"          , TO_DATE('" & DATECREATED & "','mm/dd/yyyy')" & _
									"          , '" & document_classification & "'" & _
									"          ,TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED & "','mm/dd/yyyy'),'yyyy'))" &_
									"          ,TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED & "','mm/dd/yyyy'),'mm'))" &_
									"          , 0" & _
									"          , ''" & _
									"          , '" &  request("HeaderRemarks") & "'" & _
									"          , 'Planned'" & _
									"          , '1'" & _
									"          , UPPER('" & transtype & "')" & _
									"          , '" & Session("CodeName") & "'" & _
									"          , '" & Session("CodeName") & "'" & _
									"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									"          , '" & Session("fcBA") & "'" & _
									"          , SYSDATE" & _
									"          , TO_CHAR(SYSDATE,'HH24:MI') )"
									
									'response.Write("q " & q )
									connStr.execute(q)
							 
						End if
						rsPO.Close
						set rsPO = Nothing

						' update ke tabel invoice untuk kolom payment
						if jenisinv = "1" then
							sqlupdatepay = "update invoice set payment=payment + '" & xAmount & "' " & _ 
							" where invoicecode='" & txtnoinvoice & "'"
							connStr.execute(sqlupdatepay)
							'Response.Write sqlupdatepay
						else
							sqlupdatepay = " update invoiceservicehdr set payment=payment + '" & xAmount & "' " & _ 
							" where invoiceno='" & txtnoinvoice & "'"
							connStr.execute(sqlupdatepay)
						end if
						

						oSQLQuery = " select nvl(trigger_system, 'Bank') trigger_system" & _
									" from bank" & _
									" where fcba = '" & Session("fcBA") & "'" & _
									"       and fccode = '" & xCode & "'"
						set oRecordSet = server.CreateObject("Adodb.recordset")
						oRecordSet.Open oSQLQuery, connstr
						If Not oRecordSet.Eof then
							IsCostbook = oRecordSet("trigger_system")
						Else
							IsCostbook = "Bank"
						End if
						oRecordSet.Close
						Set oRecordSet = Nothing
						
						If UCase(IsCostbook) = UCase("Costbook") Then
							If CheckAvailableJobZerorise(Jobcode, Session("FCBA")) = 1 then
							
								qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
												" from paymentvoucherdetail " & _
												" where " & _
												" fcba = '" & Session("fcBA") & "' "
												
								Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
								rsqMaxLineNo.CursorType = 3
								rsqMaxLineNo.CursorLocation = 3
								rsqMaxLineNo.Open qMaxLineNo, connstr

								if not rsqMaxLineNo.eof then
									maxLineNopvd = cdbl(rsqMaxLineNo("maxLineNo"))
								else
									maxLineNopvd = 0
								end if
								rsqMaxLineNo.close
								set rsqMaxLineNo = Nothing
								
								maxLineNopvd = maxLineNopvd + 1
								
								rlineno = maxLineNopvd

								Call InsertDataPaymentVoucherDetail

								qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
												" from cost_activity " & _
												" where " & _
												" fcba = '" & Session("fcBA") & "' "
												
								Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
								rsqMaxLineNo.CursorType = 3
								rsqMaxLineNo.CursorLocation = 3
								rsqMaxLineNo.Open qMaxLineNo, connstr

								if not rsqMaxLineNo.eof then
									maxLineNo = cdbl(rsqMaxLineNo("maxLineNo"))
								else
									maxLineNo = 0
								end if
								rsqMaxLineNo.close
								set rsqMaxLineNo = Nothing
								
								maxLineNo = maxLineNo + 1
								
								InsertIntoCostActivity = " insert into cost_activity (tdate" & _
														 " 							  , ACCOUNTINGYEAR " &_
														 " 							  , ACCOUNTINGPERIOD " &_
														 "                            , jobcode" & _
														 "                            , source" & _
														 "                            , sourcetype" & _
														 "                            , sourcecode" & _
														 "                            , sourcecode_status" & _
														 "                            , targettype" & _
														 "                            , targetcode" & _
														 "                            , targetcode_status" & _
														 "                            , referenceno" & _
														 "                            , quantity" & _
														 "                            , output" & _
														 "                            , amount" & _
														 "                            , fcba" & _
														 "                            , closing_date" & _
														 "                            , process_date" & _
														 "                            , lastupdate" & _
														 "                            , lasttime" & _
														 "                            , lineno" & _
														 "                            , title_reference_01" & _
														 "                            , auto_reference_01" & _
														 "                            , title_reference_02" & _
														 "                            , auto_reference_02" & _
														 "                            , title_reference_03" & _
														 "                            , auto_reference_03" & _
														 "                            , title_reference_04" & _
														 "                            , auto_reference_04" & _
														 "                            , title_reference_05" & _
														 "                            , auto_reference_05" & _
														 "                            , debetcredit" & _
														 " ) values (to_date('" & DATECREATED & "', 'mm/dd/yyyy')" & _
														 "          ,TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED & "','mm/dd/yyyy'),'yyyy'))" &_
														 "          ,TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED & "','mm/dd/yyyy'),'mm'))" &_
														 "           , '" & Jobcode & "'" & _
														 "           , '" & SourceType & "'" & _
														 "           , 'BA'" & _
														 "           , '" & xCode & "'" & _
														 "           , '" & xCode & "'" & _
														 "           , '" & LocationType & "'" & _
														 "           , '" & LocationCode & "'" & _
														 "           , null" & _
														 "           , '" & fcCode & "'" & _
														 "           , 0" & _
														 "           , 0" & _
														 "           , REPLACE('" & Amount & "', ',', '')" & _
														 "           , '" & Session("fcBA") & "'" & _
														 "           , last_day(to_date('" & DATECREATED & "', 'mm/dd/yyyy'))" & _
														 "           , sysdate" & _
														 "           , sysdate" & _
														 "           , to_char(sysdate, 'hh24:mi')" & _
														 "           , '" & maxLineNo & "'" & _
														 "           , Cost_Activity_Info_Api.Get_Title_Reference_01('BA', 'CB','Cr')" & _
														 "           , '" & LocationCode & "'" & _
														 "           , Cost_Activity_Info_Api.Get_Title_Reference_02('BA', 'CB','Cr')" & _
														 "           , '" & fcCode & "'" & _
														 "           , null" & _
														 "           , null" & _
														 "           , null" & _
														 "           , null" & _
														 "           , 'Line No:'" & _
														 "           , '" & rlineno & "'" & _
														 "           , Decode(GreaterThan(REPLACE('" & Amount & "', ',', ''), '0'), 'TRUE', 'Db', 'Cr')" & _
														 " )"

								connStr.execute(InsertIntoCostActivity)
							Else
								err.number = 1001
								err.description = "Jobcode """ & Jobcode & """ tidak boleh di gunakan untuk zerorise, silahkan pilih jobcode yang benar !"
							End If
							
						Else
							
								qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
												" from paymentvoucherdetail " & _
												" where " & _
												" fcba = '" & Session("fcBA") & "' "
												
								Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
								rsqMaxLineNo.CursorType = 3
								rsqMaxLineNo.CursorLocation = 3
								rsqMaxLineNo.Open qMaxLineNo, connstr

								if not rsqMaxLineNo.eof then
									maxLineNopvd = cdbl(rsqMaxLineNo("maxLineNo"))
								else
									maxLineNopvd = 0
								end if
								rsqMaxLineNo.close
								set rsqMaxLineNo = Nothing
								
								maxLineNopvd = maxLineNopvd + 1
								
								rlineno = maxLineNopvd
						
								Call InsertDataPaymentVoucherDetail

						End If
								 
					else
						err.number = 1001
						err.description = check_data
					End If
								 
								 
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
				cmdView	= "True"
			elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			    'response.write("masuka")
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				'response.write("masukb" & Cek_IsDailyClosing)
				If Cek_IsDailyClosing = "FALSE" then
					
					check_data = f_i_check(Session("fcBA"), fcCode, DATECREATED, xCode, Jobcode, locationtype, locationcode, documentno, amount)

					if check_data = "Success" then
				
						invoicecode = ""
						agreementcode = ""
					
						if locationtype = "AP" then
							invoicecode = documentno
						elseif locationtype = "CA" then
							agreementcode = documentno
						else
							invoicecode = ""
							agreementcode = ""
						end if

						oSQLQuery = " select nvl(bank_api.get_trigger_system(bankcode, fcba), 'Bank') trigger_system" & _
									" from paymentvoucher" & _
									" where fcba = '" & Session("fcBA") & "'" & _
									"       and vouchercode = '" & fcCode & "'"
						set oRecordSet = server.CreateObject("Adodb.recordset")
						oRecordSet.Open oSQLQuery, connstr
						If Not oRecordSet.Eof then
							IsCostbook = oRecordSet("trigger_system")
						Else
							IsCostbook = "Bank"
						End if
						oRecordSet.Close
						Set oRecordSet = Nothing
						
						'response.write("masuk1")
						 
					If UCase(IsCostbook) = UCase("Costbook") Then
						
						If CheckAvailableJobZerorise(Jobcode, Session("FCBA")) = 1 then
						
							Call UpdateDataPaymentVoucherDetail
						
							arr_lineno_eof = true
							oSQLQuery = " SELECT lineno" & _
										" FROM cost_activity" & _
										" WHERE fcba = '" & Session("fcBA") & "'" & _
										"       AND sourcetype = 'BA'" & _
										"       AND referenceno = '" & fcCode & "'" & _
										"       AND auto_reference_05 = '" & rlineno & "'"
							set oRecordSet = server.CreateObject("Adodb.recordset")
							oRecordSet.Open oSQLQuery, connstr
							If Not oRecordSet.Eof then
								arr_lineno = oRecordSet.getrows()
								arr_lineno_eof = false
							End if
							oRecordSet.Close
							Set oRecordSet = Nothing
							
							if not arr_lineno_eof then
								for xi=0 to ubound(arr_lineno,2)
						
									q = " DELETE FROM cost_activity_hist_tab " & _
										" WHERE fcba = '" & Session("fcBA") & "'" & _
										"       AND lineheaderno = '" & arr_lineno(0,xi) & "'"

									connStr.execute(q)
									
								next
							end if
								
							q = " DELETE FROM cost_activity" & _
								" WHERE fcBA = '" & Session("fcBA") & "'" & _
								"       AND sourcetype = 'BA'" & _
								"       AND referenceno = '" & fcCode & "'" & _
								"       AND auto_reference_05 = '" & rlineno & "'"
								
							connStr.execute(q)
						
							qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
											" from cost_activity " & _
											" where " & _
											" fcba = '" & Session("fcBA") & "' "
											
							Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
							rsqMaxLineNo.CursorType = 3
							rsqMaxLineNo.CursorLocation = 3
							rsqMaxLineNo.Open qMaxLineNo, connstr

							if not rsqMaxLineNo.eof then
								maxLineNo = cdbl(rsqMaxLineNo("maxLineNo"))
							else
								maxLineNo = 0
							end if
							rsqMaxLineNo.close
							set rsqMaxLineNo = Nothing
							
							maxLineNo = maxLineNo + 1
							
							InsertIntoCostActivity = " insert into cost_activity (tdate" & _
													 " 							  , ACCOUNTINGYEAR " &_
													 " 							  , ACCOUNTINGPERIOD " &_
													 "                            , jobcode" & _
													 "                            , source" & _
													 "                            , sourcetype" & _
													 "                            , sourcecode" & _
													 "                            , sourcecode_status" & _
													 "                            , targettype" & _
													 "                            , targetcode" & _
													 "                            , targetcode_status" & _
													 "                            , referenceno" & _
													 "                            , quantity" & _
													 "                            , output" & _
													 "                            , amount" & _
													 "                            , fcba" & _
													 "                            , closing_date" & _
													 "                            , process_date" & _
													 "                            , lastupdate" & _
													 "                            , lasttime" & _
													 "                            , lineno" & _
													 "                            , title_reference_01" & _
													 "                            , auto_reference_01" & _
													 "                            , title_reference_02" & _
													 "                            , auto_reference_02" & _
													 "                            , title_reference_03" & _
													 "                            , auto_reference_03" & _
													 "                            , title_reference_04" & _
													 "                            , auto_reference_04" & _
													 "                            , title_reference_05" & _
													 "                            , auto_reference_05" & _
													 "                            , debetcredit" & _
													 " ) values (to_date('" & DATECREATED & "', 'mm/dd/yyyy')" & _
													 "          ,TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED & "','mm/dd/yyyy'),'yyyy'))" &_
													 "          ,TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED & "','mm/dd/yyyy'),'mm'))" &_
													 "           , '" & Jobcode & "'" & _
													 "           , '" & SourceType & "'" & _
													 "           , 'BA'" & _
													 "           , '" & xCode & "'" & _
													 "           , '" & xCode & "'" & _
													 "           , '" & LocationType & "'" & _
													 "           , '" & LocationCode & "'" & _
													 "           , null" & _
													 "           , '" & fcCode & "'" & _
													 "           , 0" & _
													 "           , 0" & _
													 "           , REPLACE('" & Amount & "', ',', '')" & _
													 "           , '" & Session("fcBA") & "'" & _
													 "           , last_day(to_date('" & DATECREATED & "', 'mm/dd/yyyy'))" & _
													 "           , sysdate" & _
													 "           , sysdate" & _
													 "           , to_char(sysdate, 'hh24:mi')" & _
													 "           , '" & maxLineNo & "'" & _
													 "           , Cost_Activity_Info_Api.Get_Title_Reference_01('BA', 'CB','Cr')" & _
													 "           , '" & LocationCode & "'" & _
													 "           , Cost_Activity_Info_Api.Get_Title_Reference_02('BA', 'CB','Cr')" & _
													 "           , '" & fcCode & "'" & _
													 "           , null" & _
													 "           , null" & _
													 "           , null" & _
													 "           , null" & _
													 "           , 'Line No:'" & _
													 "           , '" & rlineno & "'" & _
													 "           , Decode(GreaterThan(REPLACE('" & Amount & "', ',', ''), '0'), 'TRUE', 'Db', 'Cr')" & _
													 " )"

							connStr.execute(InsertIntoCostActivity)
							
						Else
							err.number = 1001
							err.description = "Jobcode """ & Jobcode & """ tidak boleh di gunakan untuk zerorise, silahkan pilih jobcode yang benar !"
						End If
					Else
						Call UpdateDataPaymentVoucherDetail
					End If
					Else
						err.number = 1001
						err.description = check_data
					End If
					cmdView	= "True"
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If 
			elseif Request("txtJenis")= "Confirmed" then
			   connstr.BeginTrans
			   strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & Session("fcBA") & "','" & fccode & "', 'Confirm' ); END; "
			   connStr.execute(strQueryStatus)

			   if err.number  = 0 then
					  connStr.CommitTrans
			   end if       
			   
		       cmdView	= "True"
			elseif Request("txtJenis")= "Approved" then
			   connstr.BeginTrans
				q =	" Update paymentvoucher" & _
					" set rowstateid = 'Approved'" & _
					"     , fcentry = '" & Session("CodeName") & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"     , lastupdate = sysdate " & _
					"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					" 		where vouchercode = '" & fccode & "'" & _
					"       and fcBA = '" & Session("fcBA") & "'"
				' response.write q
				connStr.execute(q)

			   if err.number  = 0 then
					connStr.CommitTrans
			   end if       
			   
		       cmdView	= "True"
			elseif Request("txtJenis")= "RePlanned" then
			   connstr.BeginTrans   
			   strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & Session("fcBA") & "','" & fccode & "', 'Replan' ); END; "
			   connStr.execute(strQueryStatus)	
		      
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
				
			   q = " INSERT INTO paymentvoucher_hist ( vouchercode" & _
					"                                           , currentauthorizelevel" & _
					"                                           , fcentry" & _
					"                                           , fcedit" & _
					"                                           , fcip" & _
					"                                           , fcba" & _
					"                                           , lastupdate" & _
					"                                           , lasttime)"& _
					" SELECT vouchercode" & _
					"        , '" & authorizelevelcode & "' currentauthorizelevel" & _
					"        , '" & Session("CodeName") & "' fcentry" & _
					"        , '" & Session("CodeName") & "' fcedit" & _
					"        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
					"        , fcba" & _
					"        , SYSDATE lastupdate" & _
					"        , TO_CHAR(SYSDATE, 'hh24:mi') lasttime" & _
					" FROM paymentvoucher" & _
					" WHERE vouchercode = '" & fccode & "'" & _
					"       AND fcba = '"&  session("fcba") & "'"
				' response.write q
				connStr.execute(q)
			   strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & Session("fcBA") & "','" & fccode & "', 'StartAuthorize' ); END; "
			   'response.write strQueryStatus
			   connStr.execute(strQueryStatus)	
		      
			   if err.number  = 0 then
					  connStr.CommitTrans
			   end if
			   
			   cmdView	= "True"
		    else
		        cmdView	= "True"
			end if								 	
    else
		err.number = 1001
		err.description = "Jobcode """ & Jobcode & """ adalah Account Biaya ,tidak boleh digunakan untuk mengkredit !!!."
	End If 
		'Response.Write "<br/>q-->" & q
		'Response.Write "<br/>qq-->" & qq
 else
    cmdView	= request("cmdView")
 End if

Dim rsVoucher
' Remark by iwa 
'q = " SELECT bankcode" & _
'	"        , datecreated" & _
'	"        , document_classification" & _
'	"        , totalamount" & _
'	"        , remarks" & _
'	"        , rowstateid" & _
'	"        , accountingperiod" & _
'	"        , accountingyear" & _
'	"        , accountingyear||LPAD(accountingperiod,2,0) accountingyearperiod" & _
'	" FROM PaymentVoucher" & _
'	" WHERE VoucherCode like '%" & fcCode & "%'" & _
'	"       AND fcBA = '" & session("fcBA") & "'"
' End Remark by iwa 	

q = " SELECT bankcode" & _
	"        , to_char(datecreated,'mm/dd/yyyy') datecreated" & _
	"        , document_classification" & _
	"        , totalamount" & _
	"        , remarks" & _
	"        , rowstateid" & _
	"        , accountingperiod" & _
	"        , accountingyear" & _
	"        , accountingyear||LPAD(accountingperiod,2,0) accountingyearperiod" & _
	" FROM PaymentVoucher" & _
	" WHERE VoucherCode = '" & fcCode & "'" & _
	"       AND fcBA = '" & session("fcBA") & "'"
'response.write q
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
	BankCode=rsVoucher("BankCode")
	DATECREATED=rsVoucher("DateCreated")
	TOTALAMOUNT = rsVoucher("TOTALAMOUNT")
	HeaderRemarks=rsVoucher("Remarks")	
	document_classification=rsVoucher("document_classification")
	rowstateid=rsVoucher("rowstateid")	
	ACCOUNTINGYEAR =rsVoucher("ACCOUNTINGYEAR")
	ACCOUNTINGPERIOD =rsVoucher("ACCOUNTINGPERIOD")
	
End if

q = " SELECT NVL(SUM(amount), 0) totalamount" & _
	" FROM paymentvoucherdetail" & _
	" WHERE VoucherCode='" & fcCode & "'" & _
	"       AND fcBA = '" & session("fcBA") & "'"
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
	TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
End if
 
if Request("txtJenis") = "Edit" then
	q = " SELECT vouchercode" & _
		"        , sourcetype" & _
		"        , locationtype" & _
		"        , locationcode" & _
		"        , jobcode" & _
		"        , chargetype" & _
		"        , chargecode" & _
		"        , chargejob" & _
		"        , chargeunit" & _
		"        , amount" & _
		"        , reference" & _
		"        , remarks" & _
		"        , invoicecode" & _
		"        , agreementcode" & _
		"        , jenisinv" & _
		" FROM paymentvoucherdetail" & _
		" WHERE fcba = '" & session("fcBA") & "'" & _
		"       AND vouchercode like '%" & fcCode & "%'" & _
		"       AND lineno = '" & rlineno & "'" 

	set rsVoucher = server.CreateObject("Adodb.recordset")
	rsVoucher.Open q, connstr
	
	if Not rsVoucher.Eof then
		SourceType = rsVoucher("sourcetype")
		Jobcode = rsVoucher("jobcode")
		selectedjobcode = rsVoucher("jobcode")
		LocationType = rsVoucher("locationtype")
		LocationCode = rsVoucher("locationcode")
		chargetype = rsVoucher("chargetype")
		chargecode = rsVoucher("chargecode")
		chargejob = rsVoucher("chargejob")	
		chargeunit = rsVoucher("chargeunit")	
		Amount = rsVoucher("amount")
		Reference = rsVoucher("reference")
		Remarks = rsVoucher("remarks")
		invoicecode = rsVoucher("invoicecode")
		agreementcode = rsVoucher("agreementcode")
		jenisinv = rsVoucher("jenisinv")
	End if
	
	if "" & agreementcode <> "" then
		documentno = agreementcode
	elseif "" & invoicecode <> "" then
		documentno = invoicecode
	else
		documentno = ""
	end if
end if


 
rsVoucher.Close
set rsVoucher = Nothing

if err.number <> 0 then
   
   'connStr.RollbackTrans
   response.Write  "<div class='error'>" & err.description & "</div>"
   cmdView	= "True"
End if

Dim val_trigger_system : val_trigger_system = "Bank"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select nvl(trigger_system, 'Bank') trigger_system" & _
            " from bank" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and fccode = '" & xCode & "'"
	   
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_trigger_system = oRecordSet("trigger_system")
Else
	val_trigger_system = "Bank"
End if
'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing

' Untuk Cek Start Authorize By Sendy 04/29/2019
set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT NVL(SUM(amount),0) total" & _
	" FROM paymentvoucherdetail" & _
	" WHERE vouchercode = '" & fcCode & "'" & _
	"       AND fcba = '" & session("fcba") & "'" 
' response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	pv_totalamount = cdbl(to_nol(rsFrez("total")))
else
	pv_totalamount = 0
End if
rsFrez.Close : set rsFrez = Nothing 
 
set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'PV'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND "&pv_totalamount&" BETWEEN at.amountrangefrom and at.amountrangeuntil" & _
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
' End Added

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 2))" & _
	        "       and upper(system) = upper('" & val_trigger_system & "')"
			
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
                   
                    <form name="frmEntry" method="post" action="PaymentVoucherDetail.Asp">
					
                    <font color="red">
                        <%= varStringError  %></font>
                    <table class="data" width="100%">
						<tr>
                            <td align="left">
                                Date
                            </td>
                            <td align="left">
                                <input class="datepicker" type="text" maxlength="10" onchange="javascript:reloadCode('#');" name="DATECREATED" size="10" value="<%=DATECREATED%>">
                                <!--<a href="javascript:show_calendar('frmEntry.DATECREATED');" onMouseOver="window.status='Date Picker';return true;"
                                    onmouseout="window.status='';return true;">
                                    <img src="Images/show-calendar.gif" width="24" height="22" border="0" align="top"></a>-->
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                        </tr>							
						<tr>
						<%
							 q = 	" select count(module_01) module_01 from users_tab" & _
									" where fccode ='"& session("CodeName") &"'"
					 
								 ' response.write q
								 set rsAmount = server.CreateObject("Adodb.recordset")
								 rsAmount.Open q, connstr
								 if Not rsAmount.Eof then
									while not rsAmount.eof
										module_01 = rsAmount("module_01")
										rsAmount.movenext
									wend
								End if
						%>
                            
						<% if cdbl(to_nol(module_01)) = 0 then %>
							<td align="left">
                                Bank Code
                            </td>
                            <td align="left">
								<input type="hidden" id="transtype" name="transtype" value="<%=transtype%>">
								<input type="hidden" name="banktype" value="<%=banktype%>">
                                <select name="xCode" id ="bankcode" onchange="javascript:reloadCode('#');">
								<%
								q = " SELECT '' fccode, '--Pilih Bank--' fcname from dual " & _
								    " UNION " & _
								    " SELECT a.fccode, a.fcname" & _
									"  FROM bank a" & _
									" WHERE a.fcba = '"& session ("fcBA") &"'" & _
									"       AND a.activation = 'Y'" & _
								    "       AND UPPER(a.banktype) <> ('JOURNALRECLASS')" 
								
								if UCase(cstr(""&banktype)) = "JOURNAL" then
									q = q & "       AND UPPER(banktype) = UPPER('JOURNAL')"
								else
									q = q & "       AND UPPER(banktype) IN ('CASH', 'BANK')"
								end if
								    q = q & " ORDER BY fcCode"
								' response.write q
								Set rsComboA = Server.CreateObject("ADODB.Recordset")
								
								rsComboA.Open q, connstr
								if not rsComboA.eof then
									Do While Not rsComboA.eof
								%>
									<option value="<%=rsComboA("fccode")%>" <% If cstr(""&xCode) = cstr(""&rsComboA("fccode")) Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
								<% 
										rsComboA.movenext
									loop
								End if
								rsComboA.Close : set rsComboA = Nothing : q = ""
								%>
                                </select>
								<font color="#FF0000">&nbsp;*</font>
                            </td>
						<%else%>
							<td align="left">
                                Bank Code
                            </td>
							<td align="left">
								<input type="hidden" id="transtype" name="transtype" value="<%=transtype%>">
								<input type="hidden" name="banktype" value="<%=banktype%>">
                                <select name="xCode" id ="bankcode" name="bankcode">
								<%
								q = " SELECT a.fccode, a.fcname, b.lu_name, c.module_01 " & _
									"  FROM module_tab b" & _
									"       RIGHT JOIN bank a" & _
									"          ON (a.fccode = b.code)" & _
									"          LEFT JOIN users_tab c ON (c.module_01 = b.fccode)" & _
									" WHERE a.fcba = '"& session ("fcBA") &"'" & _
									"       AND a.activation = 'Y'" & _
								    "       AND UPPER(a.banktype) <> ('JOURNALRECLASS')" & _
									" AND c.fccode = '"& session("CodeName") &"'" 

								if UCase(cstr(""&banktype)) = "JOURNAL" then
									q = q & "       AND UPPER(banktype) = UPPER('JOURNAL')"
								else
									q = q & "       AND UPPER(banktype) IN ('CASH', 'BANK')"
								end if
								' response.write q
								Set rsComboA = Server.CreateObject("ADODB.Recordset")
								
								rsComboA.Open q, connstr
								if not rsComboA.eof then
									Do While Not rsComboA.eof
								%>
									<option value="<%=rsComboA("fccode")%>" <% If cstr(""&xCode) = cstr(""&rsComboA("fccode")) Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
								<% 
										rsComboA.movenext
									loop
								End if
								rsComboA.Close : set rsComboA = Nothing : q = ""
								%>
                                </select>
								<font color="#FF0000">&nbsp;*</font>
                            </td>
						<%end if%>
						</tr>						
                        <tr>
                            <td align="left">
                                Voucher No
                            </td>
                            <td align="left">
							<%
									if  Request("txtTable")="02-paymentvoucher" then
										var_txtTable = "PAYMENTVOUCHER"
										var_tempx = "K"
									elseif Request("txtTable")="01-receivevoucher" then
										var_txtTable = "RECEIVEVOUCHER"
										var_tempx = "M"
									else
										var_txtTable = "JOURNALVOUCHER"
										var_tempx = "J"
									end if
									'response.write var_txtTable
									
									if fcCode = "" then
									'response.write xCode
								' MODIFY AMH - Apr 21 penomoran sesuai bank
										q = "select MAX(cn.PREFIX) || '"& var_tempx &"/' || '"& xCode &"/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'YY') ||'/'|| LPAD(NVL(MAX(regexp_substr(det.VOUCHERCODE, '[^/]+', 1, 5)),0000)+0001,4,0) DOCNUMBER " & _
											" from " & _
											" controlnumber cn " & _
											" left join paymentvoucher det on cn.fcba = det.fcba and cn.accountingyear = to_number(to_char(det.DATECREATED,'yyyy')) " & _
											" where cn.fcba ='" & session("fcBA") & "' " & _
											" and cn.LU_NAME ='"& var_txtTable &"' "&_
											" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yyyy')) " & _
											" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'mm')) " & _
											" and LPAD(det.vouchercode,Length(rtrim(det.vouchercode))-4) = cn.PREFIX || '"& var_tempx &"/' || '"& xCode &"/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yy') ||'/' "
										response.write q 
										set rsQuery = server.CreateObject("Adodb.recordset")
										rsQuery.Open q, connstr
										if Not rsQuery.Eof then
										
											if request("bankcodeurl") <>"" then
												fcCode = rsQuery("DOCNUMBER")
											else
												fcCode = "(NEW)"
											end if
										End if										
									End if
									
									'response.write request("bankcodeurl") 
							%>
                                <input type="text" maxlength="30" name="fcCode" size="30" value="<%=fcCode%>" title="">
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                        </tr>
                        <tr>
                        	<td align="left">
                                Last Voucher No
                            </td>
                            <td align="left">
                                <input type="text" maxlength="30" name="lastnumber" size="30" id="lastvouchercode" value="<%=lastvouchercode%>" readonly style ="background-color: #eee;">
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                Reference No
                            </td>
                            <td align="left">
                                <input type="text" maxlength="30" id="referenceno" name="referenceno" size="30" value="<%=referenceno%>" title="" autocomplete="off">
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                Cek No
                            </td>
                            <td align="left">
                                <input type="text" maxlength="30" name="cekno" size="30" value="<%=cekno%>" title="" autocomplete="off">
                            </td>
                        </tr>
                        <tr>
							<td align="left">Classification</td>
								<td align="left">
									<select name="document_classification" class="chosen-select" style="width: 50%">
									<option></option>
									<%
									qLov = " SELECT fccode, fcname FROM document_classification where document_type = 'PV' ORDER BY fccode"
						   
									Set rsLov = Server.CreateObject("ADODB.Recordset")
									rsLov.Open qLov, connstr
									if not rsLov.eof then
										while not rsLov.eof
									%>
										<option value="<%=rsLov("fccode")%>" <% if cstr(""&rsLov("fccode")) = cstr(""&document_classification) then%> selected <%end if%>><%=rsLov("fcname")%></option>
									<%
											rsLov.movenext
										wend
									end if
									%>
									</select>
									<font color="#FF0000">*</font>
								</td>	
						</tr>
				
                            <tr>
                                <td align="left">
                                    Amount Total
                                </td>
                                <td align="left">
                                    <input type="text" maxlength="50" name="TOTALAMOUNT" size="10" value="<%=TOTALAMOUNT%>"
                                        readonly="" style="background:#FFFFCC;">
                                </td>
                            </tr>
                            <%if  Request("txttable") = "02-paymentvoucher" or Request("txtJenis") = "Edit" then%>
                            <tr>
                                <td align="left">
                                   Paid For
                                </td>
                                <td align="left">
                                    <textarea name="HeaderRemarks" cols="50" rows="3"><%=HeaderRemarks%></textarea>
                                </td>
                            </tr>
                            <%else if Request("txttable") = "01-receivevoucher" or Request("txtJenis") = "Edit" then%>
                            <tr>
                                <td align="left">
                                   Receive From
                                </td>
                                <td align="left">
                                    <textarea name="HeaderRemarks" cols="50" rows="3"><%=HeaderRemarks%></textarea>
                                </td>
                            </tr>
                            <%else if Request("txttable") = "03-journalvoucher" or Request("txtJenis") = "Edit" then%>
                            <tr>
                                <td align="left">
                                   Journal By
                                </td>
                                <td align="left">
                                    <textarea name="HeaderRemarks" cols="50" rows="3"><%=HeaderRemarks%></textarea>
                                </td>
                            </tr>
                    <%end if
					end if
					end if%>
					<tr>
						<td align="left">
							Status
						</td>
						<td align="left">
							<input type="text" maxlength="50" name="rowstateid" size="30" value="<%=rowstateid%>" style="background-color:#eee;" readonly>
						</td>
					</tr>
					<tr>
						<td colspan="2" align="left">
							<input type="Hidden" value="True" name="cmdView">
							<input type="Hidden" value="<%=request("txtTable")%>" name="txtTable">
							<input type="Hidden" value="<%=request("txtModule")%>" name="txtModule">
							<input type="Submit" name="Add" value="Select">&nbsp;&nbsp;							
							<!--<a href="javascript:launchwin('PrintVoucher.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&transtype=<%=transtype%>&banktype=<%=banktype%>&transdate=<%=request("datepicker")%>&bankcodeurl=<%=request("bankcode")%>&DATECREATED=<%=DATECREATED%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
								<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana"
									size="1">Print</font></a>&nbsp;&nbsp;-->							
							<%
							txtFileReportName = "print_bank_payment_voucher"
							%>
							<a href="<%=txtFileReportName%>.html?fcba=<%=Session("fcba")%>&vouchercode=<%=fcCode%>&transtype=<%=transtype%>" target="_blank">
								<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
							</a>&nbsp;&nbsp;			
							<!-- Period Control In Here -->
								<% if RowStateId = "Planned" OR RowStateId = "" then %>
								<a href="PaymentVoucherHeaderEdit.asp?fcCode=<%=fcCode%>&xcode=<%=xcode%>&headerremarks=<%=headerremarks%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>" target="_blank">
									Edit Header
								</a>
								<a href="PaymentVoucherDetail.Asp?fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&txtJenis=Header">
									<img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');">
								</a>
								<%Else%>
								<a href="PaymentVoucherHeaderEdit.asp?fcCode=<%=fcCode%>&xcode=<%=xcode%>&headerremarks=<%=headerremarks%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>" target="_blank">
									Edit Header
								</a>
								<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
								<%End If%>
							<!-- End Period Control In Here -->
							
							<% if RowStateId = "Planned" then %>
							     <!-- Modify AMH Mar 21 validasi balance -->
								 <%if Request("txttable") = "03-journalvoucher" and TOTALAMOUNT  <> 0 then %>
								 <font color="#FF0000">Pending Confirmation,please check your balance</font>
								 <%else%>
								 <a href="?fcCode=<%=fcCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to confirm this data :  <%=fcCode%> ?');"> Confirmed </a>  
								 <% end if%>
								 
							<% Elseif RowStateId = "Confirmed" then %> 
								<% If Request("txttable") = "03-journalvoucher" Then %>
									<a href="?fcCode=<%=fcCode%>&txtJenis=Approved" onClick="return confirm('Are you sure to Approved this data :  <%=fcCode%> ?');">Approved</a>  
								<%else%>
									<a href="?fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to replan this data :  <%=fcCode%> ?');"> Re-Planned </a>
									&nbsp;&nbsp;
										<%if can_startauthorize = 0 then%>
											<font color="#FF0000">You Have No Permission to Start Authorize</font>
										<%else%>
											<a href="?fcCode=<%=fcCode%>&txtJenis=StartAuthorize" onClick="return confirm('Are you sure to start authorize this data :  <%=fcCode%> ?');">Start Authorize</a>
										<% end if%>
								<% end if%>
							<% end if%> 
							<!-- HADIR -->
							<%
							dataModule 		= "Bank"
							dataSubmodule 	= "Payment Voucher"
							dataDetail 		= "NULL"
							%>
							<!-- <a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a> -->
							<!-- END:HADIR -->
						</td>
					</tr>
			</table>
        </form>

<%if cmdView = "True" and "" & fcCode <> "" and "" & xCode <> ""  and "" & document_classification <> "" then 
	Dim rsPaymentVoucher, intRecordsPerPage, cColor

	q = " select pv.vouchercode" & _
		"        , pv.referenceno" & _
		"        , pv.cekno" & _
		"        , pvd.lineno" & _
	    "        , pvd.sourcetype" & _
	    "        , pvd.locationtype" & _
	    "        , pvd.locationcode" & _
	    "        , pvd.jobcode||' ('||jb.fcname||')' jobcode" & _
	    "        , pvd.amount" & _
	    "        , pvd.reference" & _
	    "        , pvd.remarks" & _
	    "        , pvd.invoicecode" & _
	    "        , pvd.agreementcode" & _
		"        , pvd.Chargeunit" & _
		"        , pvd.Chargejob" & _
	    " from paymentvoucher pv" & _
	    " inner join PaymentVoucherDetail pvd ON (pv.VoucherCode = pvd.vouchercode and pv.fcba = pvd.fcba)" & _
		" inner join job jb ON (pvd.jobcode = jb.fccode and pvd.fcba=jb.fcba)" & _
		" where pv.vouchercode = '" & fcCode & "'" & _
		"       and pv.fcba = '" & session("fcBA") & "'" & _
		" order by pv.fcba, pv.vouchercode, pvd.lineno"
	'response.Write q	
	Set rsPaymentVoucher = Server.CreateObject("ADODB.Recordset")
	rsPaymentVoucher.CursorType = 3
	rsPaymentVoucher.CursorLocation = 3
	rsPaymentVoucher.Open q, connstr
%>
        <form action="PaymentVoucherDetail.Asp" method="post" name="eForm" style="display:table;">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" border="0">
            <tr style="display: table-row;">
                <td align="left">
                    Original Source Type
                </td>
                <td align="left">
					<input type="hidden" size="11" name="txtlineno" value="<%=rlineno%>">
                    <input type="text" size="11" name="SourceType" value="<%=SourceType%>" >&nbsp;<font color="#FF0000">*</font>
                </td>
            </tr>
            <tr style="display: table-row;">
                <td align="left">
                    Job Code
                </td>
                <td align="left">
				    <div id="jobcodelist"></div>
					<input type="Hidden" id="selectedjobcode" name="selectedjobcode" value="<%=selectedjobcode%>">
                </td>
            </tr>
            <tr style="display: table-row;">
                <td align="left">
                    Target Type
                </td>
                <td align="left">
                    <input type="text" size="11" id="locationtype" name="locationtype" value="<%=LocationType%>" style="background-color: #eee;" readonly >&nbsp;<font
                        color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                            <font face="Verdana" size="1">Location Type</font></a>
                </td>
            </tr>
            <tr>
                <td align="left">
                    Target Code 
                </td>
                <td align="left">
                    <input type="text" size="11" id="locationcode" name="LocationCode" value="<%=LocationCode%>" onkeyup="javascript:getamountinvoice();">&nbsp;<font
                        color="#FF0000">*<a href="javascript:PopupTType('<%=session("fcBA") %>')">
                            <font face="Verdana" size="1">Location Code</font></a>
                </td>
            </tr>
			<tr>
                <td align="left" id="jdljnspemb">
                    Jenis Pembayaran
                </td>
                <td align="left">
                   <select name="jnsinvoice" value="<%=jenisinv%>" id="jnsinvoice" onchange="getJnsInvoice(this);" style="width: 20%">	
							<option value=0>--Pilih Jenis Pemb.--</option>
							<%

							sqljnsinv = "SELECT id, jenisinvoice FROM jenisinvoicemst ORDER BY jenisinvoice asc"
							Set rsjnsinv = Server.CreateObject("ADODB.Recordset")
							rsjnsinv.Open sqljnsinv, connstr
							if not rsjnsinv.eof then
								while not rsjnsinv.eof
							%>
								<option value="<%=rsjnsinv("id")%>" 
									<%if cstr(""&jnsinvoice)=cstr(""&rsjnsinv("id")) then %> selected 
										<%End If%>><%=rsjnsinv("jenisinvoice")%></option>
							<%
									rsjnsinv.movenext
								wend
							end if
							rsjnsinv.Close
							set rsjnsinv = Nothing
							%>
									
					</select> 
                </td>
            </tr>
			<tr id="pilihsup" style="display: table-row;">
			<td align="left"  id="">Pilih Supplier</td>
				<td class="" align="left">
				 <select name="supplierID" id="supplierID" style=width:200px onchange="getInvoice();">
        		</select>
				</td>
			</tr>
			<tr id="pilihinv" style="display: table-row;">
			<td align="left" width="20%" id="">Pilih Invoice</td>
			<td align="left">
				 <select name="noinvoice" id="noinvoice" style=width:200px onchange="getNominal();">
        		</select>
			</td>
				
			</tr>
			<tr id="chargeunit" style="display: table-row;">
                <td align="left">
                    Charge Unit
                </td>
                <td align="left">
                   <select  name="chargeunit"  id = "chargeunit_" >
							<option value=""></option>
							<%
							qlov = 	 " select fccode,fcname " & _
									 "	from businessunit " & _
									 "	where fccompanycode = businessunit_api.get_CompanyCode('" & session("fcba") & "') " & _
									 "	and businessunit_api.Isactive(fccode)='TRUE' " & _ 
									 "  /*and fccode <> '" & session("fcba") & "'*/"
									 
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
					</select>&nbsp;<font color="#FF0000">*</font>
                </td>
            </tr>
            <tr id="chargejob" style="display: table-row;">
                <td align="left">
                    Charge Job
                </td>
                <td align="left">
                    <input type="text" size="11" name="chargejob" id="chargejob_" onchange="javascript:getChargetype();" value="<%=chargejob%>">&nbsp;<font color="#FF0000">*</font>
                    <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
					<font face="Verdana" size="1">Activity</font>
                </td>
            </tr>
            <tr id="chargetype">
                <td align="left">
                    Charge Type
                </td>
                <td align="left">
                    <input type="text" size="11" name="chargetype" id="chargetype_"  value="<%=chargetype%>">&nbsp;<font
                        color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                            <font face="Verdana" size="1">Location Type</font></a>
                </td>
            </tr>
            <tr id="chargecode" style="display: table-row;">
                <td align="left">
                    Charge Code
                </td>
                <td align="left">
                    <input type="text" size="11" name="chargecode" value="<%=chargecode%>">&nbsp;<font
                        color="#FF0000">*
						<a href="javascript:PopupTType('')">
                            <font face="Verdana" size="1">Location Code</font></a></font>
                </td>
            </tr>
		
			
			<tr style="display: table-row;">
                
                <tr>
                    <td class="wip" align="left" style=" border-right:none;">
                        <div>Document References</div>
                    </td>
                    <td class="wip" align="left" style=" border-left:none;">
                        <input type="text" size="50" placeholder="Invoice/Contract" id="documentno" name="documentno" value="<%=documentno%>" onkeyup="javascript:getamountinvoice();">
                    </td>
                </tr>
            </tr>
            <tr style="display: table-row;">
                <td align="left">
                    Amount
                </td>
                <td align="left">
                    <input type="text" size="20" name="Amount" value="<%=Amount%>"  onkeyup="hanyaAngka(this);" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">&nbsp;<font color="#FF0000">*
                        (+) Debet (-) Credit</font>
                </td>
            </tr>
            <tr style="display: table-row;">
                <td align="left">
                    Amount Invoice
                </td>
                <td align="left">
                    <input type="text" size="20" id="amountinvoice" name="amountinvoice" value="<%=amountinvoice%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right; background: #ccc;" readonly>
                </td>
            </tr>
            <tr style="display: table-row;">
                <td align="left">
                    Reference
                </td>
                <td align="left">
                    <textarea name="Reference" cols="50"><%=Reference%></textarea>
                </td>
            </tr>
            <tr style="display: table-row;">
                <td align="left">
                    Particular
                </td>
                <td align="left">
                    <textarea name="Remarks" cols="50" rows="3"><%=Remarks%></textarea>
                </td>
            </tr>
            </td></tr>
            <tr>
                <td colspan="2" align="left">
                    <!-- diambil dari form yang diatas -->
                    <input type="hidden" name="fcCode" value="<%=fcCode%>">
                    <input type="hidden" name="cekno" value="<%=cekno%>">
                    <input type="hidden" name="referenceno" value="<%=referenceno%>">
					<input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
                    <input type="hidden" name="xCode" value="<%=xCode%>">
                    <input type="hidden" name="HeaderRemarks" value="<%=HeaderRemarks%>">
					<input type="hidden" name="banktype" value="<%=banktype%>">
                    <input type="hidden" name="transtype" value="<%=transtype%>">
                    <input type="hidden" name="rowstateid" value="<%=rowstateid%>">
                    <input type="hidden" name="document_classification" value="<%=document_classification%>">
					<input type="hidden" name="jobcode" value="<%=jobcode%>">
                    <input type="Hidden" value="<%=request("txtTable")%>" name="txtTable">
                    <input type="Hidden" value="<%=request("txtModule")%>" name="txtModule">
					
					<!-- Period Control In Here -->
						<%IF Request("txtJenis")="Edit" then%>
							<input type="hidden" name="txtJenis" value="Edit">
							<input Type="Submit" name="Add" value="Update">
						<%Else%>
							<% if RowStateId = "" OR RowStateId = "Planned" then %>
							   <input type="hidden" name="txtJenis" value="New">
							   <input Type="Submit" name="Add" value="Add">
							<%Else%>
						<%End if%>
						<!-- End Period Control In Here -->
						  <input Type="Reset" name="Reset" value="Reset">
						  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
					<%End if%>
                </td>
            </tr>
        </table>
        <table class="data xthf-site">
            <br>
            <thead>
            <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
                <th align="left">
                    NO
                </th>
                <th colspan="2">&nbsp;
                    
                </th>
                <th align="left">
                    CODE
                </th>
                <th align="left">
                    REFERENCE NO
                </th>
                 <th align="left">
                    CEK NO
                </th>
				<th align="left">
                    LINE NO
                </th>
                <th align="left">
                    SOURCE TYPE
                </th>
                <th align="left">
                    LOCATION TYPE
                </th>
                <th align="left">
                    LOCATION CODE
                </th>
                <th align="left">
                    JOB CODE
                </th>
                <th align="left">
                    AMOUNT
                </th>
                <th align="left">
                    REFERENCES
                </th>
                <th align="left">
                    REMARKS
                </th>
                <th align="left">
                    INVOICE CODE
                </th>
                <th align="left">
                    AGREEMENT CODE
                </th>
				<th align="left">
                    CHARGE UNIT
                </th>
				<th align="left">
                    CHARGE JOB
                </th>
                <!--<th align="left">
                    JENIS INVOICE 
                </th>-->
            </tr>
            </thead>
            <tbody>
            <%
   if Not rsPaymentVoucher.eof then
      nCountofField=rsPaymentVoucher.fields.count-1
	  no=1
	  While Not rsPaymentVoucher.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
            <tr bgcolor="<%=cColor%>">
                <td align="left" valign="top">
                    <%=NO%>
                </td>
				<!-- Cek Data Close or Not -->
				<!-- Period Control In Here -->
					<% if RowStateId = "Planned" then %>
						<td valign="middle" width="10">
							<a href="PaymentVoucherDetail.Asp?fcCode=<%=rsPaymentVoucher("VoucherCode")%>&txtlineno=<%=rsPaymentVoucher("lineno")%>&transtype=<%=transtype%>&banktype=<%=banktype%>&xCode=<%=xCode%>&txtJenis=Edit">
								<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPaymentVoucher(1)%>" border="0">
							</a>
						</td>
						<td>
							<a href="PaymentVoucherDetail.Asp?fcCode=<%=rsPaymentVoucher("VoucherCode")%>&txtlineno=<%=rsPaymentVoucher("lineno")%>&transtype=<%=transtype%>&banktype=<%=banktype%>&txtJenis=Details">
								<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPaymentVoucher(1)%>" border="0" onclick="return confirm('Are you sure to delete');">
							</a>
						</td>
					<%Else%>
						<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
					<%End if%>
               <!-- End Period Control In Here -->
                <%for i = 0 to nCountofField%>
					<td align="center" title="<%=rsPaymentVoucher(i).Name%>" valign="top">
						<%=rsPaymentVoucher(i)%>
					</td>
                <%next%>
                
            </tr>
            <%
		  rsPaymentVoucher.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPaymentVoucher.Close
   set rsPaymentVoucher = Nothing%>
   </tbody>
        </table>
        
        </form>
        <%End if%>
				<!-------------akhir!------------->
			</div>
		</div>
		<!--#include file="bottom.jsp"-->

<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="plugins/chosen/chosen.jquery.min.js"></script>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script>
function getselectedjobcode() {
	var ssupplier = document.getElementById("selectedjobcode");
	ssupplier.value = document.getElementById("jobcode").value;
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
	//getLastNumber();
	document.getElementById("chargeunit").style.display="none";
	document.getElementById("chargejob").style.display="none";
	document.getElementById("chargetype").style.display="none";
	document.getElementById("chargecode").style.display="none";
	getJobCode();
	getselectedjobcode();
	getamountinvoice();	
	getJobcodeIntraco();
	getTragettype();
	getHidden();

});


</script>
		
		<script type="text/javascript">	
	
			function getJnsInvoice()
			{
				document.getElementById('locationcode').value = '';
				$("#amountinvoice").val(0);
				var jnsinvoice = $("#jnsinvoice").val();
				if (jnsinvoice=="0")
				{
					getHidden();	
				}else{
					
					$("#supplierID").empty();
					$("#noinvoice").empty();
					$("#amountinvoice").val(0);
					getSup(jnsinvoice);
					getShowHidden();
				}

				
			}

			function getInvoice()
			{
				document.getElementById('noinvoice').value = '';
				var supplierID = $('#supplierID').val();
				var jnsinvoice = $('#jnsinvoice').val();
				
				if(jnsinvoice == "1"){
					getInvPurch();
					$('#locationcode').val(supplierID);	
				}else{
					getInvTBS();
					$('#locationcode').val(supplierID);	
				}
					
			}

			function getHidden() {
				$("#pilihsup").hide();							
				$("#pilihinv").hide();							
			}

			function getShowHidden () {
				
				$("#pilihsup").show();							
				$("#pilihinv").show();						
			}


			function getSup(jnsinvoice) {
				$('#supplierID').select2({
					ajax: {
						delay: 1000,
						url : 'getlistsupplierinvajax.asp?jnsinvoice=' + jnsinvoice,
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

			function getInvTBS() {
				var SupplierID	 = $('#supplierID').val();
				$('#noinvoice').select2({
					ajax: {
						delay: 1000,
						url : 'getlistinvajax.asp?SupplierID=' + SupplierID,
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

			function getInvPurch() {
				var SupplierID	 = $('#supplierID').val();
				$('#noinvoice').select2({
					ajax: {
						delay: 1000,
						url : 'getlistinvpurchajax.asp?SupplierID=' + SupplierID,
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

		</script>
    </body>
</html>