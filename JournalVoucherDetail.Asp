<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
    <title><%="JOURNAL VOUCHER | ePlaS"%></title>
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
	//var jobcode = document.getElementById("jobcode").value;
	var jobcode = document.getElementById("selectedjobcode").value;
	// var fddate = document.getElementById("fddate").value;
	// var empcode = document.getElementById("empcd").value;
	// var fieldcode = document.getElementById("targetcode").value;
	var PostStr = 'jobcode='+jobcode;
	//alert("masuk intraco");
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
	var fcba =$( "#fcba" ).val();
	
	var txtTable = $.urlParam('txtTable');
	var txtBanktype = $.urlParam('banktype');
	var vtranstype = $.urlParam('transtype');
	//alert(bankcode);
	
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Finance and Accounting&txtTable="+txtTable+"&transtype="+vtranstype+"&banktype="+txtBanktype+"&transdate="+transdate+"&fcba="+fcba+"&bankcodeurl="+bankcode+"&"+PostStr );
}

// launched from button click 
var getTragettype = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("selectedjobcode").value;
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
		var txtdate  = document.getElementById("DATECREATED").value;
		var txtfcba_ = document.getElementById("chargeunit_");
		var value_ = txtfcba_.options[txtfcba_.selectedIndex].value;
		
		if (_fcba == "" && value_ != ""){
		    _fcba = value_
			 txtttype = document.getElementById("chargetype_").value;
		}
			
		
		//var newwin   = window.open("PrintData.Asp?txtModule=Data Master&txtTable=LocationCode&txtSearch="+txtttype+"&txtFcba="+_fcba+"&Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
		var newwin   = window.open("popup_loccode.Asp?txtSearch="+txtttype+"&txtdate="+txtdate+"&Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
               
	   }

		
    </script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261015a"></script>
<% end if %>
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
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %> onLoad="document.frmEntry.fcCode.focus()">
    <%Else%>
    <body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %> onload="document.eForm.Amount.focus()">
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
 
' Menjalankan SQL dan mengembalikan pesan error ("" bila sukses), agar error tidak tertelan.
Function ExecuteChecked(sqlQuery)
	On Error Resume Next
	connStr.execute(sqlQuery)
	If Err.Number <> 0 Then
		ExecuteChecked = Err.Description
		If ExecuteChecked = "" Then ExecuteChecked = "Error " & Err.Number
	Else
		ExecuteChecked = ""
	End If
	On Error GoTo 0
End Function

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
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
 Dim rsInvoiceRecord, varTotalAmount, invoiceUpdate, recordcheck, rsrecordlineno, recordWIPLINE,invoiceRecord
 Dim rsPaymentVoucherDetail, varamountOriginal, recordinvoicepaid, vartotal, varError, ivp,varmaxlineno,recordlineno
 Dim Jobcode, LocationType, LocationCode, Reference, Remark,InvoiceCode,AGREEMENTCODE,AGREEMENTLINENO ,recordwipcheck
 Dim fisik_qty, fisik_unit, fisik_output
 Dim varSupplierCode, varStringError, varlineno, WIPPAIDRecord, rsWIPPAIDRecord, WIPUpdate ,rsrecordWIPLINE ,W, rsrecordwipcheck, WIPDATE, varContract
 Dim InsertIntoCostActivity, rsqMaxLineNo, qMaxLineNo, MaxLineNo, IsCostbook
 Dim MaxLineNopvd
 Dim rlineno, dataModule, dataSubmodule, dataDetail, qq, recPaymentVoucherDetail
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim Cek_IsAccClosing : Cek_IsAccClosing = ""
 
 Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD
 Dim xACCOUNTINGYEARPERIOD
 DIM xAmount, StatusXAmount
 Dim sqljnsinv,rsjnsinv,sqlupdatepay,txtnoinvoice,jenisinv,amountinvoice,ScriptName
 ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 
 if Request("fcba") <> "" then
	fcba = Request("fcba")
 else	
	fcba = session("fcba")  ' perbaikan: dulu "fcba = fcba" sehingga kosong bila URL tanpa fcba
 end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if request("transdate") <>"" then
   DATECREATED   = request("transdate")
 elseif Request("DATECREATED")<>"" then
   DATECREATED = Request("DATECREATED")
 else
   DATECREATED = textDate(date)
 end if
 
 if KonversiTanggalKeDbl(DATECREATED) > KonversiTanggalKeDbl(textDate(date)) then
	DATECREATED = textDate(date)
 end if
 'response.write DATECREATED
 
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
transtype = "03-JOURNAL" 'Request("transtype")

' Added Variable ChargeJob, ChargeType, ChargeCode
chargeunit=Request("chargeunit")
chargejob=Request("chargejob")
chargetype=Request("chargetype")
chargecode=Request("chargecode")
'businessunit=Request("businessunit")
businessunit=Request("chargeunit")
referenceno=Request("referenceno")
remarksdtl=Request("referenceno")
cekno=Request("cekno")
fcCode = Request("fcCode")
txtnoinvoice = Request("noinvoice")
jenisinv = "0"
'amountinvoice = Request("amountinvoice")
fisik_qty    = Request("fisik_qty")
fisik_unit   = Request("fisik_unit")
fisik_output = Request("fisik_output")

' End Added


If Request("Amount") <> ""  then
   xAmount = Replace(Amount, ",", "")
   IF Cdbl(xAmount) > 0 then
      StatusXAmount = "PLUS"
   Else
      StatusXAmount = "MINUS"       
   End If
End If

 ACCOUNTINGYEARPERIOD = Request("ACCOUNTINGYEARPERIOD")
 
 if "" & ACCOUNTINGYEARPERIOD = "" then
 ACCOUNTINGYEARPERIOD = Right(DATECREATED,4) & MID(DATECREATED,1,2)
 else
 ACCOUNTINGYEARPERIOD = Request("ACCOUNTINGYEARPERIOD")
 end if
 
 if "" & ACCOUNTINGYEARPERIOD <> "" then
    ACCOUNTINGYEAR = MID(ACCOUNTINGYEARPERIOD,1,4)
	ACCOUNTINGPERIOD = MID(ACCOUNTINGYEARPERIOD,5,2)
 end if

'response.write ACCOUNTINGYEARPERIOD & "|" & ACCOUNTINGYEAR & "|" & ACCOUNTINGPERIOD

 if Request("SourceType")="" then
    SourceType="BA" 'Default if Entry is Blank
 else
    SourceType=Request("SourceType")
 End if
dim tt

Sub DeleteDataPaymentVoucherDetail
	q = " DELETE FROM paymentvoucherdetail" & _
		" WHERE fcBA = '" & fcba & "'" & _
		"       AND vouchercode = '" & fcCode & "'" & _
		"       AND lineno = '" & rlineno & "'"
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q
	Call UpdateAmountPaymentVoucherHeader
End Sub

Sub UpdateAmountPaymentVoucherHeader
' Untuk Update Total Amount Di Header
				q = " SELECT NVL(SUM(amount), 0) totalamount" & _
					" FROM paymentvoucherdetail" & _
					" WHERE VoucherCode='" & fcCode & "'" & _
					"       AND fcBA = '" & fcba & "'"
				set rsVoucher = server.CreateObject("Adodb.recordset")
				rsVoucher.Open q, connstr
				if Not rsVoucher.Eof then
					TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
				End if

				q =	" Update paymentvoucher" & _
					" set totalamount = '"& TOTALAMOUNT &"'" & _
					" where vouchercode = '" & fccode & "'" & _
					" and fcBA = '" & fcba & "'"
	 			'connStr.execute(q)
				ExecuteAndLog ScriptName, q
End Sub

Sub InsertDataPaymentVoucherDetail
if fisik_qty = "" then
	fisik_qty = 0 
end if
if fisik_unit = "" then
	fisik_unit = 0 
end if
if fisik_output = "" then
	fisik_output = 0 
end if
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
			 "                                    , jenisinv, qty, unit, output) " & _
		     " VALUES ( '" & fcCode & "'" & _
	         "          , '" & SourceType & "'" & _
	         "          , '" & LocationType & "'" & _ 
	         "          , '" & LocationCode &  "'" & _
	         "          , '" & jobcode & "'" & _
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
	         "          , '" & fcba & "'" & _
	         "          , '" & "" & "'" & _
	         "          , '" & "" & "'"& _
	         "          , '" & "" & "'" & _
	         "          , TO_DATE('" & "01/01/1900" & "', 'mm/dd/yyyy')" & _
	         "          , '" & rlineno & "'" & _
	         "          , SYSDATE" & _
	         "          , TO_CHAR(SYSDATE,'HH24:MI') " & _ 
			 "			, '" & jenisinv & "'" & _ 
			 "			, " & fisik_qty & "" & _ 
			 "			, " & fisik_unit & "" & _ 
			 "			, " & fisik_output & "" & _ 
			 " )"
		 'response.write qq
		'connStr.execute(qq)
		' Perbaikan: dulu ExecuteAndLog menelan error sehingga baris gagal disimpan tanpa pesan
		Dim errInsertDetail : errInsertDetail = ExecuteChecked(qq)
		If errInsertDetail <> "" Then
			err.number = 1001
			err.description = "Detail jurnal gagal disimpan: " & errInsertDetail
			Exit Sub
		End If
			' Untuk Update Total Amount Di Header
			Call UpdateAmountPaymentVoucherHeader
			' End Added
		
	
End Sub

Sub UpdateDataPaymentVoucherDetail
	' Perbaikan: dulu DELETE lalu INSERT tanpa transaksi; bila INSERT gagal baris jurnal hilang.
	connStr.BeginTrans
	Dim errDeleteDetail : errDeleteDetail = ExecuteChecked(" DELETE FROM paymentvoucherdetail" & _
		" WHERE fcBA = '" & fcba & "'" & _
		"       AND vouchercode = '" & fcCode & "'" & _
		"       AND lineno = '" & rlineno & "'")
	If errDeleteDetail = "" Then Call InsertDataPaymentVoucherDetail
	If errDeleteDetail = "" And Err.Number = 0 Then
		connStr.CommitTrans
	Else
		connStr.RollbackTrans
		If errDeleteDetail <> "" Then
			err.number = 1001
			err.description = "Detail jurnal gagal diubah: " & errDeleteDetail
		End If
	End If
End Sub
		
Cek_IsAccClosing = StatusAccClosing(ACCOUNTINGYEAR,ACCOUNTINGPERIOD,fcba)
'response.write "Final =|" & Cek_IsAccClosing & "|" & fcba & "|" & DATECREATED

if 	Cek_IsAccClosing = "FALSE" then 
'--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "StartAuthorize" or Request("txtJenis")= "Approved" or Request("txtJenis")= "Posting" or Request("txtJenis")= "UnPosting" Then	
	If CheckCostJobCr(StatusXAmount,Jobcode, fcba) = 0 then
	
			if Request("txtJenis")= "Header" then
				Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then

					q = " Delete from PaymentVoucherDetail where  VoucherCode = '" & fcCode & "' and fcBA = '" & fcba & "'"
					'connStr.execute(q)
					ExecuteAndLog ScriptName, q
					
					q = " Delete from PaymentVoucher where  VoucherCode = '" & fcCode & "' and fcBA = '" & fcba & "'"
					'connStr.execute(q)
					ExecuteAndLog ScriptName, q
					
					q = "DELETE BKMJOURNALDTL where SOURCE='JRV' AND fcba='" & fcba & "' AND docno='" & fccode & "'"
					ExecuteAndLog ScriptName, q
				
					q = "DELETE BKMJOURNALHDR where SOURCECODE='FAC-JRV' AND fcba='" & fcba & "' AND docno='" & fccode & "'"
					ExecuteAndLog ScriptName, q
					
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If  
				
				cmdView	= "True"
			   
			elseif Request("txtJenis")= "Details" then
				Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then

					call DeleteDataPaymentVoucherDetail
					
					oSQLQuery = " select nvl(bank_api.get_trigger_system(bankcode, fcba), 'Bank') trigger_system" & _
								" from paymentvoucher" & _
								" where fcba = '" & fcba & "'" & _
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
									" WHERE fcba = '" & fcba & "'" & _
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
									" WHERE fcba = '" & fcba & "'" & _
									"       AND lineheaderno = '" & arr_lineno(0,xi) & "'"

								connStr.execute(q)
								
							next
						end if
						
						q = " DELETE FROM cost_activity" & _
							" WHERE fcBA = '" & fcba & "'" & _
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
				Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then
					
					check_data = f_i_check(fcba, fcCode, DATECREATED, xCode, Jobcode, locationtype, locationcode, documentno, amount)
					'--Add by Ukie to handle Auto Number of PR
							If Ucase(fcCode) = Ucase("(NEW)") Then
								oSql = " select nvl(controlnumber_api.get_nextnumber(upper('"& transtype &"')" & _
									   "                                         , " & year(textDate(datecreated)) & "" & _
									   "                                         , " & month(textDate(datecreated)) & "" & _
									   "                                         , '" & fcba & "'" & _
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
					
						'if locationtype = "AP" then
						'	invoicecode = documentno
						'elseif locationtype = "CA" then
						'	agreementcode = documentno
						'else
							invoicecode = ""
							agreementcode = ""
						'end if
				
						Set rsPO = Server.CreateObject("ADODB.Recordset")
						rsPO.Open " SELECT vouchercode FROM paymentvoucher WHERE vouchercode = '" & fcCode & "' AND fcba = '" & fcba & "'", connstr
						IF rsPO.eof then
							
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
									"          ," & ACCOUNTINGYEAR & " " &_
									"          ,TO_NUMBER(" & ACCOUNTINGPERIOD & ")" &_
									"          , 0" & _
									"          , ''" & _
									"          , '" &  request("HeaderRemarks") & "'" & _
									"          , 'Planned'" & _
									"          , '1'" & _
									"          , UPPER('" & transtype & "')" & _
									"          , '" & Session("CodeName") & "'" & _
									"          , '" & Session("CodeName") & "'" & _
									"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									"          , '" & fcba & "'" & _
									"          , SYSDATE" & _
									"          , TO_CHAR(SYSDATE,'HH24:MI') )"
									' response.Write("q " & q )
									'connStr.execute(q)
									' Perbaikan: dulu ExecuteAndLog menelan error, detail tetap tersimpan tanpa header
									Dim errInsertHeader : errInsertHeader = ExecuteChecked(q)
									If errInsertHeader <> "" Then
										err.number = 1001
										err.description = "Header jurnal gagal disimpan: " & errInsertHeader
									End If
									
							 
						End if
						rsPO.Close
						set rsPO = Nothing

						
						If err.number = 0 Then  ' Perbaikan: jangan simpan detail bila header gagal dibuat
						oSQLQuery = " select nvl(trigger_system, 'Bank') trigger_system" & _
									" from bank" & _
									" where fcba = '" & fcba & "'" & _
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
							If CheckAvailableJobZerorise(Jobcode, fcba) = 1 then
							
								qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
												" from paymentvoucherdetail " & _
												" where " & _
												" fcba = '" & fcba & "' and vouchercode='" & fcCode & "'"
												
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
												" fcba = '" & fcba & "' "
												
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
														 "          ,TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yyyy'))" &_
														 "          ,TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'mm'))" &_
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
														 "           , '" & fcba & "'" & _
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

								'connStr.execute(InsertIntoCostActivity)
							Else
								err.number = 1001
								err.description = "Jobcode """ & Jobcode & """ tidak boleh di gunakan untuk zerorise, silahkan pilih jobcode yang benar !"
							End If
							
						Else
							
								qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
												" from paymentvoucherdetail " & _
												" where " & _
												" fcba = '" & fcba & "' and VoucherCode='" & fccode & "'"
												
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
						End If  ' akhir: detail hanya disimpan bila header berhasil
								 
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
				Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then
					
					check_data = f_i_check(fcba, fcCode, DATECREATED, xCode, Jobcode, locationtype, locationcode, documentno, amount)
					
                    check_data = "Success"
					if check_data = "Success" then
							Call UpdateDataPaymentVoucherDetail
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
			   strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & fcba & "','" & fccode & "', 'Confirm' ); END; "
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
					"       and fcBA = '" & fcba & "'"
				' response.write q + "<br>"
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
			
				

			   if err.number  = 0 then
					connStr.CommitTrans
			   end if       
			   
		       cmdView	= "True"
			elseif Request("txtJenis")= "Posting" then
			   connstr.BeginTrans
			
				strQueryStatus = "BEGIN FAC_JOURNAL_FAC_JRV_API.FAC_BNK_JRV_DOCNO('" & fccode & "'," & ACCOUNTINGYEAR & ", " & ACCOUNTINGPERIOD & ", '" & fcba & "','System'); " & _ 
								 "END;"
				'response.write strQueryStatus + "<br>"				 
				'connStr.execute(strQueryStatus)	
				ExecuteAndLog ScriptName, strQueryStatus
				
					
				strQueryStatus = "update BKMJOURNALHDR set postingflg=1,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND docno='" & fccode & "'"
				'response.write strQueryStatus + "<br>"	
				'connStr.execute(strQueryStatus)
				ExecuteAndLog ScriptName, strQueryStatus
				
				strQueryStatus = "BEGIN fac_closing_accounting_api.ProsesClosing(" & ACCOUNTINGYEAR & ", " & ACCOUNTINGPERIOD & ",'" & fcba & "','System'); " & _ 
							     "END;"			
				'response.write strQueryStatus + "<br>"					 
				'connStr.execute(strQueryStatus)
				ExecuteAndLog ScriptName, strQueryStatus
				

			   if err.number  = 0 then
					connStr.CommitTrans
			   end if       
			   
		       cmdView	= "True"
			elseif Request("txtJenis")= "UnPosting" then
			   connstr.BeginTrans
			    strQueryStatus = "DELETE BKMJOURNALDTL where SOURCE='JRV' AND fcba='" & fcba & "' AND docno='" & fccode & "'"
				ExecuteAndLog ScriptName, strQueryStatus
				
				strQueryStatus = "DELETE BKMJOURNALHDR where SOURCECODE='FAC-JRV' AND fcba='" & fcba & "' AND docno='" & fccode & "'"
				ExecuteAndLog ScriptName, strQueryStatus

				strQueryStatus = "DELETE PAYMENTVOUCHER_HIST where fcba='" & fcba & "' AND vouchercode='" & fccode & "'"
				ExecuteAndLog ScriptName, strQueryStatus

				strQueryStatus = "UPDATE PAYMENTVOUCHER set postingflg='0',ROWSTATEID='Confirmed',lastupdate=sysdate,fcedit='" & session("CodeName") & "'  where fcba='" & fcba & "' AND vouchercode='" & fccode & "'"
				ExecuteAndLog ScriptName, strQueryStatus
			
				strQueryStatus = "BEGIN fac_closing_accounting_api.ProsesClosing(" & ACCOUNTINGYEAR & ", " & ACCOUNTINGPERIOD & ",'" & fcba & "','System'); " & _ 
							     "END;"			
				'response.write strQueryStatus + "<br>"					 
				ExecuteAndLog ScriptName, strQueryStatus
				

			   if err.number  = 0 then
					connStr.CommitTrans
			   end if       
			   
		       cmdView	= "True"
			elseif Request("txtJenis")= "RePlanned" then
			   connstr.BeginTrans   
			   strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & fcba & "','" & fccode & "', 'Replan' ); END; "
			   connStr.execute(strQueryStatus)	
			   
			   q =	" Update paymentvoucher" & _
					" set   fcedit = '" & Session("CodeName") & "'" & _
					" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"     , lastupdate = sysdate " & _
					"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					"     , postingflg = '0' " & _
					" 		where vouchercode = '" & fccode & "'" & _
					"       and fcBA = '" & fcba & "'"
				' response.write q
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
				
				strQueryStatus = "BEGIN FAC_JOURNAL_FAC_JRV_API.FAC_BNK_JRV_DOCNO('" & fccode & "'," & ACCOUNTINGYEAR & ", " & ACCOUNTINGPERIOD & ", '" & fcba & "','System'); " & _ 
								 "END;"
				'connStr.execute(strQueryStatus)	
				ExecuteAndLog ScriptName, strQueryStatus
				
					
				strQueryStatus = "update BKMJOURNALHDR set postingflg=0,lastupdate=sysdate,fcedit='" & session("CodeName") & "' where fcba='" & fcba & "' AND docno='" & fccode & "'"
				'connStr.execute(strQueryStatus)
				ExecuteAndLog ScriptName, strQueryStatus
				
				strQueryStatus = "BEGIN fac_closing_accounting_api.ProsesClosing(" & ACCOUNTINGYEAR & ", " & ACCOUNTINGPERIOD & ",'" & fcba & "','System'); " & _ 
							     "END;"			
				'connStr.execute(strQueryStatus)
				ExecuteAndLog ScriptName, strQueryStatus
			   
		      
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
					"       AND fcba = '"&  fcba & "'"
				' response.write q
				connStr.execute(q)
			   strQueryStatus = "BEGIN PAYMENTVOUCHER_API.NextState('" & fcba & "','" & fccode & "', 'StartAuthorize' ); END; "
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
else
    err.number = 1001
	err.description = "Accounting Period  " & ACCOUNTINGPERIOD & "/" & ACCOUNTINGYEAR & "  LOCKED !!!"
	cmdView	= request("cmdView")
End if

Dim rsVoucher
	

q = " SELECT bankcode" & _
	"        , to_char(datecreated,'mm/dd/yyyy') datecreated" & _
	"        , document_classification" & _
	"        , totalamount" & _
	"        , referenceno" & _
	"        , remarks" & _
	"        , rowstateid" & _
	"        , accountingperiod" & _
	"        , accountingyear" & _
	"        , accountingyear||LPAD(accountingperiod,2,0) accountingyearperiod" & _
	" FROM PaymentVoucher" & _
	" WHERE VoucherCode = '" & fcCode & "'" & _
	"       AND fcBA = '" & fcba & "'"
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
	ACCOUNTINGYEARPERIOD = rsVoucher("accountingyearperiod")
	referenceno = rsVoucher("referenceno")
End if

q = " SELECT NVL(SUM(amount), 0) totalamount" & _
	" FROM paymentvoucherdetail" & _
	" WHERE VoucherCode='" & fcCode & "'" & _
	"       AND fcBA = '" & fcba & "'"
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
	TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
End if
 
if Request("txtJenis") = "Edit" then
   if trim(rlineno) = "" Then
		q = " UPDATE paymentvoucherdetail " & _
		    " SET lineno = (select MAX(lineno)+1 FROM paymentvoucherdetail WHERE fcba='" & fcba & "' AND vouchercode='" & fcCode & "')" & _
			" WHERE fcba = '" & fcba & "'" & _
			"       AND vouchercode = '" & fcCode & "'" & _
			"       AND jobcode = '" & jobcode & "'" & _
			"       AND locationcode = '" & LocationCode & "'" & _
			"       AND locationtype = '" & LocationType & "'" 
		ExecuteAndLog ScriptName, q	
			
   end If
 
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
		"        , qty,unit,output,lineno" & _
		" FROM paymentvoucherdetail" & _
		" WHERE fcba = '" & fcba & "'" & _
		"       AND vouchercode = '" & fcCode & "'" & _
		"       AND jobcode = '" & jobcode & "'" & _
		"       AND locationcode = '" & LocationCode & "'" & _
		"       AND locationtype = '" & LocationType & "'" & _
		"       AND lineno = '" & rlineno & "'" 
	'Response.Write q
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
		fisik_qty = rsVoucher("qty")
		fisik_unit = rsVoucher("unit")
		fisik_output = rsVoucher("output")
		rlineno = rsVoucher("lineno")
	End if
	
	'if "" & agreementcode <> "" then
	'	documentno = agreementcode
	'elseif "" & invoicecode <> "" then
	'	documentno = invoicecode
	'else
		documentno = ""
	'end if
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
	        " where fcba = '" & fcba & "'" & _
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
	"       AND fcba = '" & fcba & "'" 
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
	"       AND aus.fcba = '" & fcba & "'"
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
	        " where fcba = '" & fcba & "'" & _
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
                   
                    <form name="frmEntry" method="post" action="JournalVoucherDetail.Asp">
					
                    <font color="red">
                        <%= varStringError  %></font>
                    <table class="data" width="100%">
						<tr>
                            <td align="left">
                                Date
                            </td>
                            <td align="left">
                                <input class="datepicker" type="text" maxlength="10" onchange="javascript:reloadCode('#');" name="DATECREATED" id="DATECREATED" size="10" value="<%=DATECREATED%>">
								<font color="#FF0000">&nbsp;*</font>
								<input type="text" maxlength="30" name="fcba" id="fcba" size="5" value="<%=fcba%>" style="background-color:#eee;" readonly>
                            </td>
                        </tr>							
						<tr>
                            <td align="left">
                                Accounting Period
                            </td>
							
							<% 	Dim rsComboA,Data_ACCOUNTINGYEARPERIOD

'								q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from ACCOUNTINGYEARPERIOD where fcBA='" &  fcba & "'  order by accountingyear||lpad(accountingperiod,2,0)"
								q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from  ACCOUNTINGYEARPERIOD where fcBA='" &  fcba & "'  and to_number(to_char(valid_until ,'YYYY')) >= to_number('"& Year(Session("v_Costbook")) &"')-1  and to_number(to_char(valid_until ,'MM')) >= to_number('"& Month(Session("v_Costbook")) &"') -1  order by accountingyear||lpad(accountingperiod,2,0)"
							
								Set rsComboA = Server.CreateObject("ADODB.Recordset")
								rsComboA.CursorType = 3
								rsComboA.CursorLocation = 3
								rsComboA.Open q, connstr
								if not rsComboA.eof then%>
								<td align="left"> 
									<select name="ACCOUNTINGYEARPERIOD">    
										<% Do While Not rsComboA.eof %>
										<% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then 
											Data_ACCOUNTINGYEARPERIOD = "true"
										else
											Data_ACCOUNTINGYEARPERIOD = "false"
										End if%>
										<option value="<%=rsComboA("ACCOUNTINGYEARPERIOD")%>" <% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then %> Selected <%End if%>><%=rsComboA("Description")%></option> 
										<%  rsComboA.movenext
										   loop
									if Data_ACCOUNTINGYEARPERIOD = "false" then %>
										<option value="<%=ACCOUNTINGYEAR&ACCOUNTINGPERIOD%>" Selected>Tahun <%=ACCOUNTINGYEAR%> Bulan <%=ACCOUNTINGPERIOD%></option>
									<% end if %>
									
									<%
										End if
										rsComboA.Close
										set rsComboA = Nothing
										q=""
										%>
									</select> <font color="#FF0000">*</font>
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
								    "       AND UPPER(a.banktype) <> ('JOURNALRECLASS')" & _ 
									"       AND UPPER(banktype) = UPPER('JOURNAL')" & _
								    " ORDER BY fcCode"
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
									" AND c.fccode = '"& session("CodeName") &"'" & _ 
									"       AND UPPER(banktype) = UPPER('JOURNAL')"
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
										var_txtTable = "JOURNALVOUCHER"
										var_tempx = "J"
									if fcCode = "" then
									'response.write xCode
								    'MODIFY AMH - Apr 21 penomoran sesuai bank
										q = "select NVL(MAX(cn.PREFIX),'" & fcba & "'||'/') || '"& var_tempx &"/' || '"& xCode &"/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'YY') ||'/'|| LPAD(NVL(MAX(regexp_substr(det.VOUCHERCODE, '[^/]+', 1, 5)),0000)+0001,4,0) DOCNUMBER " & _
											" from " & _
											" controlnumber cn " & _
											" left join paymentvoucher det on cn.fcba = det.fcba and cn.accountingyear = to_number(to_char(det.DATECREATED,'yyyy')) " & _
											" where cn.fcba ='" & fcba & "' " & _
											" and cn.LU_NAME ='"& var_txtTable &"' "& _ 
											" AND transtype LIKE '03%' " & _ 
											" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yyyy')) " & _
											" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'mm')) " & _
											" and LPAD(det.vouchercode,Length(rtrim(det.vouchercode))-4) = cn.PREFIX || '"& var_tempx &"/' || '"& xCode &"/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yy') ||'/' "
										'response.write q 
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
							<td align="left">Classification</td>
								<td align="left">
									<select name="document_classification" class="chosen-select" >
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
                            <tr>
                                <td align="left">
                                   Journal By
                                </td>
                                <td align="left">
                                    <input type="text" maxlength="200" name="HeaderRemarks" size="30" value="<%=HeaderRemarks%>">  
                                </td>
                            </tr>
                    <tr>
						<td align="left">
							Status
						</td>
						<td align="left">
							<input type="text" maxlength="50" name="rowstateid" size="30" value="<%=rowstateid%>" style="background-color:#eee;" readonly>
						</td>
					</tr>
					<tr>
                            <td align="left">
                                Remarks
                            </td>
                            <td align="left">
							    <textarea id="referenceno" name="referenceno"  cols="50" rows="3"><%=referenceno%></textarea>
                            </td>
                    </tr>
					<tr>
						<td colspan="2" align="left">
							<input type="Hidden" value="True" name="cmdView">
							<input type="Hidden" value="<%=request("txtTable")%>" name="txtTable">
							<input type="Hidden" value="<%=request("txtModule")%>" name="txtModule">
							<input type="Submit" name="Add" value="Select">&nbsp;&nbsp;							
							<%
							txtFileReportName = "print_bank_payment_voucher"
							%>
							<a href="<%=txtFileReportName%>.html?fcba=<%=fcba%>&vouchercode=<%=fcCode%>&transtype=<%=transtype%>" target="_blank">
								<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
							</a>&nbsp;&nbsp;			
							<!-- Period Control In Here -->
								<% if RowStateId = "Planned" OR RowStateId = "" then %>
								<a href="JournalVoucherHeaderEdit.asp?fcCode=<%=fcCode%>&xcode=<%=xcode%>&headerremarks=<%=headerremarks%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&ACCOUNTINGYEARPERIOD=<%=ACCOUNTINGYEARPERIOD%>&fcba=<%=fcba%>&referenceno=<%=referenceno%>" target="_blank">
									Edit Header
								</a>
								<!--<a href="JournalVoucherDetail.Asp?fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&fcba=<%=fcba%>&txtJenis=Header">-->
								<!--	<img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');">-->
								<!--</a>-->
								<%Else%>
								<a href="JournalVoucherHeaderEdit.asp?fcCode=<%=fcCode%>&xcode=<%=xcode%>&headerremarks=<%=headerremarks%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&ACCOUNTINGYEARPERIOD=<%=ACCOUNTINGYEARPERIOD%>&fcba=<%=fcba%>" target="_blank">
									Edit Header
								</a>
								<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
								<%End If%>
							<!-- End Period Control In Here -->
							
							<% if RowStateId = "Planned" then %>
							     <!-- Modify AMH Mar 21 validasi balance -->
								 <%if TOTALAMOUNT  <> 0 then %>
								 <font color="#FF0000">Pending Confirmation,please check your balance</font>
								 <%else%>
								 <a href="?fcCode=<%=fcCode%>&fcba=<%=fcba%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to confirm this data :  <%=fcCode%> ?');"> Confirmed </a>  
								 <% end if%>
							<% end if%>
							
							<% if RowStateId = "Confirmed" then %> 
									<a href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&ACCOUNTINGYEARPERIOD=<%=ACCOUNTINGYEARPERIOD%>&fcba=<%=fcba%>&txtJenis=Approved" onClick="return confirm('Are you sure to Approved this data :  <%=fcCode%> ?');">Approved</a>  
									<a href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&ACCOUNTINGYEARPERIOD=<%=ACCOUNTINGYEARPERIOD%>&fcba=<%=fcba%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to Re-Planned this data :  <%=fcCode%> ?');">Re-Planned</a>  
							<% end if%>
							
							<% if RowStateId = "Approved" then %> 
									<a href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&ACCOUNTINGYEARPERIOD=<%=ACCOUNTINGYEARPERIOD%>&fcba=<%=fcba%>&txtJenis=Posting" onClick="return confirm('Are you sure to Journal Posting this data :  <%=fcCode%> ?');">Posting</a>  
									<a href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&ACCOUNTINGYEARPERIOD=<%=ACCOUNTINGYEARPERIOD%>&fcba=<%=fcba%>&txtJenis=UnPosting" onClick="return confirm('Are you sure to Un-Posting this data :  <%=fcCode%> ?');">Un-Posting</a>  
									
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

<%if cmdView = "True" and "" & fcCode <> "" and "" & xCode <> "" and "" & document_classification <> "" then 
	Dim rsPaymentVoucher, intRecordsPerPage, cColor

	q = " select pv.vouchercode" & _
		"        , pv.referenceno" & _
		"        , pvd.lineno" & _
	    "        , pvd.sourcetype" & _
	    "        , pvd.locationtype" & _
	    "        , pvd.locationcode" & _
	    "        , pvd.jobcode" & _
		"        , job_api.get_jobcode_name(pvd.jobcode,pvd.fcba) jobname" & _
		"        , pvd.chargeunit" & _
	    "        , pvd.amount" & _
	    "        , pvd.reference" & _
	    "        , pvd.remarks" & _		
		"        , pvd.qty" & _		
		"        , pvd.unit" & _		
		"        , pvd.output" & _		
	    " from paymentvoucher pv" & _
	    " inner join PaymentVoucherDetail pvd ON (pv.VoucherCode = pvd.vouchercode and pv.fcba = pvd.fcba)" & _
	    " where pv.vouchercode = '" & fcCode & "'" & _
		"       and pv.fcba = '" & fcba & "'" & _
		" order by pv.fcba, pv.vouchercode, pvd.lineno"
	'response.Write q	
	Set rsPaymentVoucher = Server.CreateObject("ADODB.Recordset")
	rsPaymentVoucher.CursorType = 3
	rsPaymentVoucher.CursorLocation = 3
	rsPaymentVoucher.Open q, connstr
%>
        <form action="JournalVoucherDetail.Asp" method="post" name="eForm" style="display:table;">
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
				</td>
				<input type="Hidden" id="selectedjobcode" name="selectedjobcode" value="<%=selectedjobcode%>">
            </tr>
            <tr style="display: table-row;">
                <td align="left">
                    Target Type
                </td>
                <td align="left">
                    <input type="text" size="11" id="locationtype" name="LocationType" value="<%=LocationType%>" style="background-color: #eee;" readonly >&nbsp;<font
                        color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                            <font face="Verdana" size="1">Location Type</font></a>
                </td>
            </tr>
            <tr>
                <td align="left">
                    Target Code
                </td>
                <td align="left">
                    <input type="text" size="11" id="locationcode" id="locationcode" name="LocationCode" value="<%=LocationCode%>" >&nbsp;<font
                        color="#FF0000">*</font><a href="javascript:PopupTType('<%=fcba %>')">
                            <font face="Verdana" size="1">Location Code</font><span class="satuan_item" id="locationname">&nbsp;</span></a>
                </td>
            </tr>
			
			<tr id="chargeunit" style="display: table-row;">
                <td align="left">
                    Charge Unit
                </td>
                <td align="left">
                   <select  name="chargeunit"  id = "chargeunit_">
							<option value=""></option>
							<%
							qlov = 	 " select fccode,fcname " & _
									 "	from businessunit " & _
									 "	where fccompanycode = businessunit_api.get_CompanyCode('" & fcba & "') " & _
									 "	and businessunit_api.Isactive(fccode)='TRUE' " & _ 
									 "  /*and fccode <> '" & fcba & "'*/"
									 
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
					</select>
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
                    <input type="text" size="11" name="chargetype" id="chargetype_" value="<%=chargetype%>">&nbsp;<font
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
                        color="#FF0000">*</font><a href="javascript:PopupTType('')">
                            <font face="Verdana" size="1">Location Code</font></a></font>
                </td>
            </tr>
		
			<tr style="display: table-row;">
                <td align="left">
                    Fisik Qty-Unit-Output
                </td>
                <td align="left">
                    <input type="text" size="11" name="fisik_qty" value="<%=fisik_qty%>">&nbsp;<input type="text" size="11" name="fisik_unit" value="<%=fisik_unit%>">&nbsp;<input type="text" size="11" name="fisik_output" value="<%=fisik_output%>">
                </td>
            </tr>
			
			<tr style="display: table-row;">
                
                <tr>
                    <td class="wip" align="left" style=" border-right:none;">
                        <div>Document References</div>
                    </td>
                    <td class="wip" align="left" style=" border-left:none;">
                        <input type="text" size="50"  id="documentno" name="documentno" value="<%=documentno%>" >
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
					<input type="hidden" name="ACCOUNTINGYEARPERIOD" value="<%=ACCOUNTINGYEARPERIOD%>">
                    <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					
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
                    JOB NAME
                </th>
				<th align="left">
                    CHARGE UNIT
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
                    Qty
                </th>
				<th align="left">
                    Unit
                </th>
				<th align="left">
                    Output
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
							<a href="JournalVoucherDetail.Asp?fcCode=<%=rsPaymentVoucher("VoucherCode")%>&selectedjobcode=<%=rsPaymentVoucher("jobcode")%>&LocationType=<%=rsPaymentVoucher("locationtype")%>&LocationCode=<%=rsPaymentVoucher("locationcode")%>&txtlineno=<%=rsPaymentVoucher("lineno")%>&transtype=<%=transtype%>&banktype=<%=banktype%>&xCode=<%=xCode%>&fcba=<%=fcba%>&txtJenis=Edit">
								<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPaymentVoucher(1)%>" border="0">
							</a>
						</td>
						<td>
							<a href="JournalVoucherDetail.Asp?fcCode=<%=rsPaymentVoucher("VoucherCode")%>&txtlineno=<%=rsPaymentVoucher("lineno")%>&transtype=<%=transtype%>&banktype=<%=banktype%>&fcba=<%=fcba%>&txtJenis=Details">
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
	getJobcodeIntraco();
	getTragettype();
	
});

function getViewcharge(){
				var kodeintraco = $("#selectedjobcode").val();
			}


</script>
		
    </body>
</html>