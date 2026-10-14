<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Invoice</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">

function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
}

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

var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbainvoice.asp', 'POST', PostStr, showlistfcba);
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

var getlistpocode = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedpocode = document.getElementById("selectedpocode").value;
	var PostStr = 'fcba='+selectedfcba+'&pocode='+selectedpocode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getpocodenotinvoicedbasedonfcba.asp', 'POST', PostStr, showlistpocode);
}

// The function for handling the response from the server
var showlistpocode = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("pocodelist").innerHTML = "";
	document.getElementById("pocodelist").innerHTML = response;

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

var getpoinfo = function () {
	// build up the post string when passing variables to the server side page

	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedpocode = document.getElementById("selectedpocode").value;
	var PostStr = 'fcba='+selectedfcba+'&pocode='+selectedpocode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetPOInfo.asp', 'POST', PostStr, showpoinfo);
}

// The function for handling the response from the server
var showpoinfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var data = JSON.parse(response);
    
    // update the Div to show the result from the server
	document.getElementById("suppliercode").value = data.suppliercode;
	document.getElementById("suppliername").value = data.suppliername;
	document.getElementById("downpayment").value = data.dp;
	document.getElementById("vat").value = data.ppn;
	document.getElementById("totalppn").value = data.totalppn;
	document.getElementById("othercharge").value = data.othercharge;
	document.getElementById("otherchargestock").value = data.otherchargestock;

};

var reloadInvCode = function () {
	var invdate =$( ".datepicker" ).val();
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Procurement&txtTable=Invoice&invdateurl="+invdate );
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
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
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
 
 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks,otherchargestock
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks,txtNoPO,v_CrossCekPO
 Dim recordTotalAmount, rsrecordTotalAmount, varTotalAmount, updateNetBalance
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""


 Dim Jobcode, LocationType, LocationCode, Reference, Remark

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 ' if request("fccOde")="" then
    ' fcCode="" '----- Nantinya Ambil dari Format Master
 ' else
    ' fcCode = request("fccOde")
 ' end if
 
 fcCode = request("fccOde")
If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if request("invdateurl") <>"" then
   DATECREATED   = request("invdateurl")
 elseif Request("DATECREATED")<>"" then
   DATECREATED   =Request("DATECREATED")
 else
   DATECREATED   =textDate(date)
 end if
 
 if KonversiTanggalKeDbl(DATECREATED) > KonversiTanggalKeDbl(textDate(date)) then
	DATECREATED = textDate(date)
 end if
 
 if Request("fakturdate")<>"" then
   fakturdate   =Request("fakturdate")
 else
   fakturdate   =textDate(date)
 end if


 if Request("invoicedatesupplier")<>"" then
   invoicedatesupplier   =Request("invoicedatesupplier")
 else
   invoicedatesupplier   =textDate(date)
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
pocode = request("pocode")
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
selectedpocode = Request("selectedpocode")
invoice_by = Request("invoice_by")
selectedinvoice = Request("selectedinvoice")
vat = Request("vat")
amountvat = Request("amountvat")
totalamount = Request("totalamount")
refinvoice = Request("refinvoice")
suppliercode = Request("suppliercode")
suppliername = Request("suppliername")
no_faktur = Request("no_faktur")
fakturdate = Request("fakturdate")
totalppn = Request("totalppn")
downpayment = Request("downpayment")
othercharge = Request("othercharge")
use_own_ppn = Request("use_own_ppn")
otherchargestock = Request("otherchargestock")
invoicedatesupplier = Request("invoicedatesupplier")
top = Request("top")
dphapus = Request("dphapus")

 If dphapus = "" Then
        dphapus = 0
End If
 If downpayment = "" Then
        downpayment = 0
End If
 If othercharge = "" Then
        othercharge = 0
End If
 If use_own_ppn = "" Then
        use_own_ppn = 0
End If
 If otherchargestock = "" Then
        otherchargestock = 0
End If
 If vat = "" Then
        vat = 0
End If
 If amountvat = "" Then
        amountvat = 0
End If
 If totalppn = "" Then
        totalppn = 0
End If


sub post_amount
	p_amountbeforevat = 0 : p_othercharge = 0 : p_totalamount = 0
	recordTotalAmount = " SELECT NVL(SUM(ind.amountbeforevat),0) amountbeforevat" & _
						"        , NVL(SUM(inv.othercharge), 0) othercharge" & _
						"        , NVL(SUM(inv.otherchargestock), 0) otherchargestock" & _
						"        , NVL(SUM(ind.amount), 0) totalAmount" & _
						" FROM invoicedetail ind" & _
						" INNER JOIN invoice inv ON (inv.invoicecode = ind.invoicecode AND inv.fcba = ind.fcba)" & _
						" WHERE inv.invoicecode = '" & fcCode & "'" & _ 
						"       AND inv.fcba = '" & fcba & "'"
	'Response.Write recordTotalAmount
	Set rsrecordTotalAmount = Server.CreateObject("ADODB.Recordset")
	rsrecordTotalAmount.Open recordTotalAmount, connstr

	if Not rsrecordTotalAmount.Eof then
		p_amountbeforevat = cdbl(to_nol(rsrecordTotalAmount("amountbeforevat")))
		p_othercharge = cdbl(to_nol(rsrecordTotalAmount("othercharge")))
		otherchargestock = cdbl(to_nol(rsrecordTotalAmount("otherchargestock")))
		p_totalamount = cdbl(to_nol(rsrecordTotalAmount("totalAmount")))
	Else
		p_amountbeforevat = 0
		otherchargestock = request("otherchargestock")
	End if
	
	If use_own_ppn <> "1" Then 
		vat = cdbl(vat)
		totalppn  = (p_amountbeforevat + otherchargestock) * vat /100
		amountvat = p_amountbeforevat + totalppn
		
		'totalamount = p_amountbeforevat + p_othercharge
		'Update Aam Mar 21
		downpayment = cdbl(to_nol(Replace(Request("downpayment"), ",", ""))) ' cdbl(to_nol(Request("downpayment")))
		totalamount = amountvat + p_othercharge + otherchargestock - downpayment
			'  response.write downpayment
		' response.write totalppn

		' ini proses edit
		q = " UPDATE invoice" & _
			" SET amount = '" & p_amountbeforevat & "'" & _
			"     , totalppn = '" & totalppn & "'" & _
			"     , amountvat = '" & amountvat & "'" & _
			"     , othercharge = replace('" & request("othercharge") & "',',','')" & _ 
            "     , otherchargestock = replace('" & otherchargestock & "',',','')" & _ 	
			"     , totalamount = '" & totalamount & "'" & _ 
			"     , lastupdate = SYSDATE " & _
			"     , lasttime = TO_CHAR(SYSDATE,'HH24:MI') "& _	
            "     , dp =  '" & downpayment & "'" & _	
            "     , fcedit='" & Session("CodeName") & "'" & _ 			
			" WHERE invoicecode = '" & fcCode & "'" & _
			"       AND fcba = '" & fcba & "'"
		' response.write q
		connStr.execute(q)
	'	Response.end
	Else
		' totalppn = p_amountbeforevat * vat /100
		' amountvat = p_amountbeforevat + totalppn

		'Update Aam Mar 21
		'totalamount = p_amountbeforevat + p_othercharge
		downpayment = cdbl(to_nol(Replace(Request("downpayment"), ",", ""))) 'cdbl(to_nol(Request("downpayment")))
		totalamount = p_amountbeforevat + p_othercharge + amountvat - downpayment

		q = " UPDATE invoice" & _
			" SET amount = '" & p_amountbeforevat & "'" & _
			"     , amountvat = '" & amountvat & "'" & _
			"     , othercharge = replace('" & p_othercharge & "',',','')" & _ 
            "     , otherchargestock = replace('" & request("otherchargestock") & "',',','')" & _ 	
			"     , totalamount = '" & totalamount & "'" & _ 
			"     , lastupdate = SYSDATE " & _
			"     , lasttime = TO_CHAR(SYSDATE,'HH24:MI') "& _		
            "     , dp =  '" & downpayment & "'" & _
			"     , fcedit='" & Session("CodeName") & "'" & _ 				
			" WHERE invoicecode = '" & fcCode & "'" & _
			"       AND fcba = '" & fcba & "'"
		' response.write q
		connStr.execute(q)
	End If 
end sub

sub post_amount2
	p_amountbeforevat = 0 : x_othercharge = 0 : p_totalamount = 0
	recordTotalAmount = " SELECT NVL(SUM(ind.amountbeforevat),0) amountbeforevat" & _
						"        , NVL(SUM(inv.otherchargestock), 0) otherchargestock" & _
						"        , NVL(SUM(inv.othercharge), 0) othercharge" & _
						"        , NVL(SUM(ind.amount), 0) totalAmount" & _
						"        , NVL(SUM(inv.totalppn), 0) totalppn" & _
						" FROM invoicedetail ind" & _
						" INNER JOIN invoice inv ON (inv.invoicecode = ind.invoicecode AND inv.fcba = ind.fcba)" & _
						" WHERE inv.invoicecode = '" & fcCode & "'" & _ 
						"       AND inv.fcba = '" & fcba & "'" & _
						" GROUP BY inv.invoicecode, inv.fcba"
	' response.Write recordTotalAmount
	Set rsrecordTotalAmount = Server.CreateObject("ADODB.Recordset")
	rsrecordTotalAmount.Open recordTotalAmount, connstr

	if Not rsrecordTotalAmount.Eof then
		p_amountbeforevat = cdbl(to_nol(rsrecordTotalAmount("amountbeforevat")))
		x_othercharge = cdbl(to_nol(rsrecordTotalAmount("othercharge")))
		otherchargestock = cdbl(to_nol(rsrecordTotalAmount("otherchargestock")))
		p_totalamount = cdbl(to_nol(rsrecordTotalAmount("totalAmount")))
		p_totalppn = cdbl(to_nol(rsrecordTotalAmount("totalppn")))
	Else
		p_amountbeforevat = 0
	End if
	
	'response.write "<br> use_own_ppn" & use_own_ppn
	
	If use_own_ppn <> "1" Then 
		'totalppn = p_amountbeforevat * vat /100
		'test Aam Apr 21'
		
		if request("otherchargestock") <> 0 then
		totalppn = request("totalppn")
		else
		totalppn = p_amountbeforevat * vat /100
		end if
		amountvat = p_amountbeforevat + totalppn
		'totalamount = p_amountbeforevat + p_othercharge  Replace(Amount, ",", "")
		downpayment = cdbl(to_nol(replace(Request("downpayment"), ",","")))
		totalamount = amountvat + x_othercharge + request("otherchargestock") - downpayment

		q = " UPDATE invoice" & _
			" SET amount = '" & p_amountbeforevat & "'" & _
			"     , totalppn = '" & totalppn & "'" & _
			"     , amountvat = '" & amountvat & "'" & _
			"     , totalamount = '" & totalamount & "'" & _ 
			"     , lastupdate = SYSDATE " & _
			"     , lasttime = TO_CHAR(SYSDATE,'HH24:MI') "& _	
			"     , othercharge = replace('" & request("othercharge") & "',',','')" & _ 
            "     , otherchargestock = replace('" & otherchargestock & "',',','')" & _ 	
			"     , dp =  '" & downpayment & "'" & _
			"     , fcedit='" & Session("CodeName") & "'" & _ 	
			" WHERE invoicecode = '" & fcCode & "'" & _
			"       AND fcba = '" & fcba & "'"
		' response.write q
		connStr.execute(q)
	Else
		' totalppn = round(p_amountbeforevat * vat /100,0)
		totalppn = request("totalppn")
		amountvat = p_amountbeforevat + totalppn
		downpayment = cdbl(to_nol(replace(Request("downpayment"), ",","")))
		totalamount = amountvat + x_othercharge + request("otherchargestock") - downpayment
		
		'  response.write totalamount
		' response.write totalppn
		' response.write p_amountbeforevat
		
		q = " UPDATE invoice" & _
			" SET amount = '" & p_amountbeforevat & "'" & _
			"     , totalamount = '" & totalamount & "'" & _
			"     , amountvat = '" & amountvat & "'" & _		
			"     , lastupdate = SYSDATE " & _
			"     , lasttime = TO_CHAR(SYSDATE,'HH24:MI') "& _	
			"     , othercharge = replace('" & request("othercharge") & "',',','')" & _ 
			"     , otherchargestock = replace('" & otherchargestock & "',',','')" & _ 	
			"     , dp =  '" & downpayment & "'" & _
			"     , fcedit='" & Session("CodeName") & "'" & _ 	
			" WHERE invoicecode = '" & fcCode & "'" & _
			"       AND fcba = '" & fcba & "'"
		' response.write q
		connStr.execute(q)
	End If
end sub
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")="Posting" Then
	
	REMARKS=request("REMARKS")
	
	if Request("txtJenis")= "Header" then
		
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
			arr_dtl_eof = true
			q = " SELECT receivenotecode, pocode, purchaselinenokey, amount FROM invoicedetail WHERE invoicecode = '" & fcCode & "' AND fcba = '" & fcba & "'"
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open q, connstr
			if not rsPO.eof then
				arr_dtl = rsPO.getrows()
				arr_dtl_eof = false
			end if
			rsPO.close : set rsPO = nothing
			
			if not arr_dtl_eof then
				for xyz=0 to ubound(arr_dtl,2)
					q = " DELETE FROM invoicedetail" & _
						" WHERE invoicecode = '" & fcCode & "'" & _
						"       AND receivenotecode = '" & arr_dtl(0,xyz) & "'" & _
						"       AND pocode = '" & arr_dtl(1,xyz) & "'" & _
						"       AND purchaselinenokey = '" & arr_dtl(2,xyz) & "'" & _
						"       AND fcBA = '" & fcba & "'"
					' response.write q
					connStr.execute(q)
					
					q = " UPDATE receivenotedetail SET amountinvoice = amountinvoice - '" & arr_dtl(3,xyz) & "' " & _
						" WHERE receivenotecode = '" & arr_dtl(0,xyz) & "'" & _
						"       AND purchaselinenokey = '" & arr_dtl(2,xyz) & "'" & _
						"       AND fcBA = '" & fcba & "'"
					' response.write q
					connStr.execute(q)
				next
				erase arr_dtl
			end if
			arr_dtl_eof = true

			q = " DELETE FROM invoice WHERE invoicecode = '" & fcCode & "' AND fcba = '" & fcba & "'"
			connStr.execute(q)
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		
		cmdView	= "True"
    elseif Request("txtJenis")= "Posting" then
					
					strQueryStatus = "UPDATE invoice SET POSTINGFLG='1' WHERE fcba = '" & fcba & "' AND invoicecode = '" & fcCode & "'"
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
					
					strQueryStatus = "BEGIN fac_journal_invpurch_api.prosesbyNotran('" & fcCode & "', '" & fcba & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
					
					
					strQueryStatus = "update BKMJOURNALHDR set postingflg=1 where fcba='" & fcba & "' AND docno='" & fcCode & "'"
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
	
					strQueryStatus = "BEGIN fac_closing_accounting_api.ProsesClosing(TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED  & "','MM/DD/YYYY'),'YYYY')), " & _ 
			                         "TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED  & "','MM/DD/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
									 "END;"			
					'response.write strQueryStatus	& "<br>"			 
					connStr.execute(strQueryStatus)
		      
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then

			q = " DELETE FROM invoicedetail" & _
			    " WHERE invoicecode = '" & fcCode & "'" & _
			    "       AND receivenotecode = '" & Request("receivenotecode") & "'" & _
			    "       AND pocode = '" & pocode & "'" & _
			    "       AND purchaselinenokey = '" & Request("purchaselinenokey") & "'" & _
			    "       and fcBA='" & fcba & "'"
			'response.write q
			connStr.execute(q)
			
		

			qq = 	" Update ReceiveNoteDetail  set AMOUNTINVOICE= NVL( amountinvoice - (SELECT SUM (NVL (inh.amount, 0)) amount" & _
					"	  FROM  invoice inh" & _ 
					"	  INNER JOIN invoicedetail ind ON (inh.fcba = ind.fcba AND inh.invoicecode = ind.invoicecode)" & _
					"	 WHERE inh.invoicecode = '"& fccode &"' AND inh.fcba = '"& fcba &"' and ind.itemcode = '"& request("itemcode") &"'), 0)"  & _
					", LASTUPDATE = sysdate " & _
					", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					" where RECEIVENOTECODE='" & request("receivenotecode") & "' and purchaseitemcode = '" & request("itemcode") & "'" & _
					" and fcBA='" & fcba & "'"
				' response.Write (qq) 
				connStr.execute (qq) 
			
			'response.write "b" & "post_amount"
			Call post_amount

		
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
			
		cmdView	= "True"

		'response.Write(q)
	elseif Request("txtJenis")="New" then
			
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
		'--Add by Ukie to handle Auto Number of PR
				If Ucase(fcCode) = Ucase("(NEW)") Then
					oSql = " select nvl(controlnumber_api.get_nextnumber(upper('INVOICE')" & _
						   "                                         , " & year(textDate(DATECREATED)) & "" & _
						   "                                         , " & month(textDate(DATECREATED)) & "" & _
						   "                                         , '" & fcba & "'" & _
						   " ), 'FALSE') INVOICE from dual"
					
					 Response.Write oSql
						   
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.Open oSql, connstr
					If Not oRs.EOF Then
						txtINVOICE = oRs("INVOICE")
					End If
					oRs.Close
					Set oRs = Nothing
						fcCode = txtINVOICE
				End If
				
				If txtINVOICE = "FALSE" Then
					err.Number = 1001
					err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
				End If
				
				'Response.Write PRCODE & txtPRCode
			'--End Added
			
			Dim rsPO
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open 	" Select INVOICECODE, suppliercode from INVOICE where INVOICECODE = '" & fcCode & "'" & _
						" and fcBA = '" & fcba & "'", connstr
			

			IF rsPO.eof then	    
				Dim txtCodeSupplier,txtCodePO
				q = " Select distinct POCODE,SUPPLIERCODE From RECEIVENOTE where pocode = '" & pocode & "'"  & _
					" and fcBA = '" & fcba & "'"
				'response.write q
				Dim rsLPO  
				set rsLPO = Server.CreateObject("ADODB.Recordset")
				rsLPO.Open q, connstr
				if not rsLPO.eof then
					txtCodePO=rsLPO("POCODE")
					txtCodeSupplier=rsLPO("SUPPLIERCODE")			
				End if 
			
				rsLPO.Close
				set rsLPO = Nothing  

				response.write cdbl(to_nol(totalppn))
				
				if cdbl(to_nol(totalppn)) > 0 then
					'Response.Write othercharge
					'response.End
					q = " Insert into Invoice (INVOICECODE,USE_OWN_PPN, INVOICE_BY, REFINVOICE, INVOICEDATE, SUPPLIERCODE, PAID, OTHERCHARGE,OTHERCHARGESTOCK, totalppn, REMARKS, NETBALANCE, pocode, no_faktur, fakturdate, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, vat, dp,INVOICEDATESUPPLIER,top,ceklistdp) " & _
					" values ('" & fcCode & "','" & use_own_ppn & "','" & invoice_by & "','" & refinvoice & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & txtCodeSupplier& "',0,REPLACE('" & request("othercharge") & "', ',', ''),REPLACE('" & request("otherchargestock") & "', ',', ''),REPLACE('" & request("totalppn") & "', ',', ''),'',0,'" & pocode& "','" & no_faktur & "', TO_DATE('" & fakturdate & "', 'mm/dd/yyyy'),'" & _
					Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'), '" & vat & "', REPLACE('" & downpayment & "',',',''),to_date('" & invoicedatesupplier & "','mm/dd/yyyy'),'" & top & "','" & dphapus & "')"
					' response.write q
					connStr.execute(q)
				else
					q = " Insert into Invoice (INVOICECODE, USE_OWN_PPN, INVOICE_BY, REFINVOICE, INVOICEDATE, SUPPLIERCODE, PAID, OTHERCHARGE,OTHERCHARGESTOCK, REMARKS, NETBALANCE, pocode, no_faktur, fakturdate, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, vat, dp,INVOICEDATESUPPLIER,top,ceklistdp) " & _
						" values ('" & fcCode & "','" & use_own_ppn & "','" & invoice_by & "','" & refinvoice & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & txtCodeSupplier& "',0,REPLACE('" & request("othercharge") & "', ',', ''),REPLACE('" & request("otherchargestock") & "', ',', ''),'',0,'" & pocode& "','" & no_faktur & "', TO_DATE('" & fakturdate & "', 'mm/dd/yyyy'),'" & _
						Session("CodeName") & "','" & Session("CodeName") & "','" & _
						Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'), '" & vat & "', REPLACE('" & downpayment & "',',',''),to_date('" & invoicedatesupplier & "','mm/dd/yyyy'),'" & top & "','" & dphapus & "')"
				'	response.write q
					connStr.execute(q)
				end if
			End if
			rsPO.Close	
			Dim n
			v_amount = 0
			v_quantity = 0
			' Validasi untuk amount dan quantity GRN terhadap invoice Added By SAG 18 Juli 2019 
				 q =  " SELECT hdr.pocode" & _
					  "        , dtl.receivenotecode" & _
					  "        , dtl.purchaselinenokey" & _
					  "        , dtl.purchaseitemcode" & _
					  "        , dtl.fcba" & _
					  "        , s.fcname" & _
					  "        , NVL(dtl.quantity,0) quantity" & _
					  "        , NVL(dtl.amount,0) amount" & _
					  "        , NVL(dtl.amountpo_receive,0) amountreceived" & _
					  " FROM receivenotedetail dtl" & _
					  " INNER JOIN receivenote hdr ON (hdr.receivenotecode = dtl.receivenotecode AND hdr.fcba = dtl.fcba)" & _
					  " LEFT JOIN stock s ON (s.fcba = dtl.fcba AND s.fccode = dtl.purchaseitemcode)" & _
					  " WHERE hdr.pocode = '" & pocode & "'" & _
					  "       AND dtl.fcBA='" & fcba & "'" & _
					  "       AND NVL(dtl.amountinvoice,0)  < amountreceived" & _
					  " ORDER BY dtl.receivenotecode, dtl.purchaseitemcode"

				' response.write q
				 
				set rsLPO = Server.CreateObject("ADODB.Recordset")
				rsLPO.Open q, connstr
				if not rsLPO.eof then  
					purchaseitemcode= rsLPO("purchaseitemcode") 	
					fcba= rsLPO("fcba") 	
					receivenotecode=rsLPO("receivenotecode") 
				End if 
				 
				 q =   	"   SELECT ind.invoicecode, SUM (NVL (ind.amount, 0)) amount, SUM (NVL (ind.quantity, 0)) quantity,  " & _
						"   SUM (NVL (rnd.amountpo_receive, 0)) amountgrn," & _
						"	SUM (NVL (rnd.quantity, 0)) quantitygrn" & _
						"     FROM invoicedetail ind" & _
						"          INNER JOIN receivenotedetail rnd" & _
						"             ON (    ind.receivenotecode = rnd.receivenotecode" & _
						"                 AND ind.fcba = rnd.fcba)" & _
						"    WHERE ind.receivenotecode = '"& receivenotecode &"' AND ind.fcBA = '"& fcba &"'" & _
						" 	AND itemcode= '"& purchaseitemcode &"'" & _
						" GROUP BY ind.invoicecode"  

				' response.write q
				 
				set rsLPO = Server.CreateObject("ADODB.Recordset")
				rsLPO.Open q, connstr
				if not rsLPO.eof then 
					amount=cdbl(to_nol(rsLPO("amount")))	
					quantity=cdbl(to_nol(rsLPO("quantity"))) 
					amountgrn=cdbl(to_nol(rsLPO("amountgrn")))	
					quantitygrn=cdbl(to_nol(rsLPO("quantitygrn")))  
				End if  
			' End Added By SAG 18 Juli 2019
			'response.Write "No " & (cdbl(request("no"))-1) & " amount : " &  request("AmountInvoice_"&TRIM(n))
			if cdbl(to_nol(totalppn)) > 0 then
				For n=1 to (cdbl(request("no"))-1) 
					if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then 
						' v_amount = amount + request("AmountInvoice_"&TRIM(n))
						' v_quantity = quantity + request("Quantity_"&TRIM(n)) 
						
						' response.write amountgrn
						
						' if cdbl(to_nol(v_amount)) > cdbl(to_nol(amountgrn)) AND  cdbl(to_nol(v_quantity)) > cdbl(to_nol(quantitygrn)) then 
						' 	err.Number = 1001
						' 	err.description = "Quantity Dan Amount Sudah Lebih Dari Receivenote"
						' Else 			
							p_amountbeforevat = 0
							p_amountbeforevat = Cdbl(to_nol(REPLACE(request("AmountInvoice_"&TRIM(n)), ",", ""))) -Cdbl(to_nol(REPLACE(request("othercharge_"&TRIM(n)), ",", "")))
								q = " Insert into  INVOICEDETAIL (INVOICECODE, ITEMCODE, QUANTITY, AMOUNT, POCODE, receivenotecode, purchaselinenokey, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, amountbeforevat, othercharge) " & _
									" values ('" & fcCode & "','" & request("Stock_"&TRIM(n)) & "','" & request("Quantity_"&TRIM(n)) &  _
									"',REPLACE('" & request("AmountInvoice_"&TRIM(n)) & "', ',', ''),'" & pocode & "','" & request("receivenotecode_"&TRIM(n)) & "','" & request("linenokey_"&TRIM(n)) & "'," & _
									"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'),'" & p_amountbeforevat & "','" & request("othercharge_"&TRIM(n)) & "')"
								connStr.execute(q) 
								
								'response.Write "e " & q
								
								' q =   "   SELECT ind.fcba, rnd.receivenotecode, ind.invoicecode, ind.itemcode, SUM (NVL (ind.amount, 0)) amount" & _
								' 		"     FROM invoicedetail ind" & _
								' 		"          INNER JOIN receivenotedetail rnd" & _
								' 		"             ON (    ind.receivenotecode = rnd.receivenotecode" & _
								' 		"                 AND ind.fcba = rnd.fcba)" & _
								' 		"    WHERE     ind.receivenotecode = '"&  request("receivenotecode_"&TRIM(n)) &"'" & _
								' 		"          AND ind.fcBA = '"& fcba &"'" & _
								' 		"          AND ind.itemcode = '"& request("Stock_"&TRIM(n)) &"'" & _
								' 		" GROUP BY ind.fcba, ind.invoicecode, ind.itemcode, rnd.receivenotecode"  

								' 'response.write q
								 
								' set rsLPO = Server.CreateObject("ADODB.Recordset")
								' rsLPO.Open q, connstr
								' if not rsLPO.eof then
								' 	while not rsLPO.eof 
								' 		q = " Update ReceiveNoteDetail  set AMOUNTINVOICE= AMOUNTINVOICE+ REPLACE('" & rsLPO("amount") & "', ',', '')"  & _
								' 			", LASTUPDATE = sysdate " & _
								' 			", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
								' 			" where RECEIVENOTECODE='" & rsLPO("receivenotecode") & "' and purchaseitemcode = '" & rsLPO("itemcode") & "'" & _
								' 			" and fcBA='" & fcba & "'"
								' 		'response.Write(q)
								' 		connStr.execute(q)
								' 	rsLPO.movenext
								' 	wend
								' End if 
								q = " Update ReceiveNoteDetail  set AMOUNTINVOICE= (SELECT SUM (NVL (ind.amount, 0)) amount" & _
									"     FROM invoicedetail ind" & _
									"          INNER JOIN receivenotedetail rnd" & _
									"             ON (    ind.receivenotecode = rnd.receivenotecode" & _
									"                 AND ind.fcba = rnd.fcba)" & _
									"    WHERE     ind.receivenotecode = '"&  request("receivenotecode_"&TRIM(n)) &"'" & _
									"          AND ind.fcBA = '"& fcba &"')"  & _
									", LASTUPDATE = sysdate " & _
									", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
									" where RECEIVENOTECODE='" & request("receivenotecode_"&TRIM(n)) & "' and purchaseitemcode = '" & request("Stock_"&TRIM(n)) & "'" & _
									" and fcBA='" & fcba & "'"
								'response.Write "d " & q
								connStr.execute(q)
						End If  
					' End if	 
				Next
				
				'response.write "<br>c " & "post_amount2"
				' ini fungsi untuk menghitung sub total disaat proses input baru
				Call post_amount2
			Else
				For n=1 to (cdbl(request("no"))-1) 
					if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then 
						p_amountbeforevat = 0
						p_amountbeforevat = Cdbl(to_nol(REPLACE(request("AmountInvoice_"&TRIM(n)), ",", ""))) -Cdbl(to_nol(REPLACE(request("othercharge_"&TRIM(n)), ",", "")))
							q = " Insert into  INVOICEDETAIL (INVOICECODE, ITEMCODE, QUANTITY, AMOUNT, POCODE, receivenotecode, purchaselinenokey, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, amountbeforevat, othercharge) " & _
								" values ('" & fcCode & "','" & request("Stock_"&TRIM(n)) & "','" & request("Quantity_"&TRIM(n)) &  _
								"',REPLACE('" & request("AmountInvoice_"&TRIM(n)) & "', ',', ''),'" & pocode & "','" & request("receivenotecode_"&TRIM(n)) & "','" & request("linenokey_"&TRIM(n)) & "'," & _
								"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'),'" & p_amountbeforevat & "','" & request("othercharge_"&TRIM(n)) & "')"
							' response.Write(q)
							connStr.execute(q)
							
						'	response.Write "e " & q
							
							' q =   	"   SELECT ind.fcba, rnd.receivenotecode, ind.invoicecode, ind.itemcode, SUM (NVL (ind.amount, 0)) amount" & _
							' 		"     FROM invoicedetail ind" & _
							' 		"          INNER JOIN receivenotedetail rnd" & _
							' 		"             ON (    ind.receivenotecode = rnd.receivenotecode" & _
							' 		"                 AND ind.fcba = rnd.fcba)" & _
							' 		"    WHERE     ind.receivenotecode = '"&  request("receivenotecode_"&TRIM(n)) &"'" & _
							' 		"          AND ind.fcBA = '"& fcba &"'" & _
							' 		"          AND ind.itemcode = '"& request("Stock_"&TRIM(n)) &"'" & _
							' 		" GROUP BY ind.fcba, ind.invoicecode, ind.itemcode, rnd.receivenotecode"  

							' 'response.write q
							 
							' set rsLPO = Server.CreateObject("ADODB.Recordset")
							' rsLPO.Open q, connstr
							' if not rsLPO.eof then
							' 	while not rsLPO.eof 
							' 		q = " Update ReceiveNoteDetail  set AMOUNTINVOICE=AMOUNTINVOICE+ REPLACE('" & rsLPO("amount") & "', ',', '')"  & _
							' 			", LASTUPDATE = sysdate " & _
							' 			", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							' 			" where RECEIVENOTECODE='" & rsLPO("receivenotecode") & "' and purchaseitemcode = '" & rsLPO("itemcode") & "'" & _
							' 			" and fcBA='" & fcba & "'"
							' 		' response.Write(q)
							' 		connStr.execute(q)
							' 	rsLPO.movenext
							' 	wend
							' End if  
						q = " Update ReceiveNoteDetail  set AMOUNTINVOICE= (SELECT SUM (NVL (ind.amount, 0)) amount" & _
							"     FROM invoicedetail ind" & _
							"          INNER JOIN receivenotedetail rnd" & _
							"             ON (    ind.receivenotecode = rnd.receivenotecode" & _
							"                 AND ind.fcba = rnd.fcba)" & _
							"    WHERE     ind.receivenotecode = '"&  request("receivenotecode_"&TRIM(n)) &"'" & _
							"          AND ind.fcBA = '"& fcba &"')"  & _
							", LASTUPDATE = sysdate " & _
							", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							" where RECEIVENOTECODE='" & request("receivenotecode_"&TRIM(n)) & "' and purchaseitemcode = '" & request("Stock_"&TRIM(n)) & "'" & _
							" and fcBA='" & fcba & "'"
						'response.Write(q)
						connStr.execute(q)
						
					'	response.Write "f " & q
					End if	 
				Next
				
				Call post_amount
			End If
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
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

 Dim rsVoucher,SudahClose
 q = " select hdr.invoicedate, hdr.suppliercode, hdr.fcba, hdr.totalppn, hdr.invoice_by, " & _
	 " hdr.pocode, hdr.refinvoice, hdr.dp, hdr.amount, hdr.vat, hdr.amountvat, hdr.othercharge,hdr.otherchargestock," & _
	 " hdr.totalAmount, hdr.no_faktur, hdr.fakturdate, bu.fccompanycode, " & _
	 " supplier_api.get_supp_name(hdr.suppliercode, hdr.fcba) suppliername,hdr.invoicedatesupplier,hdr.top,hdr.ceklistdp " & _
     " from Invoice hdr" & _
     " left join businessunit bu on bu.fccode = hdr.fcba" & _
     " where hdr.invoicecode='" & fcCode & "' and hdr.fcba='" & fcba & "'" 
 
 set rsVoucher = server.CreateObject("Adodb.recordset")
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED = rsVoucher("INVOICEDATE")
	supppliercode = rsVoucher("suppliercode") 
	suppliername = rsVoucher("suppliername") 
    fccompanycode = rsVoucher("fccompanycode")
	fcba = rsVoucher("fcba")
	totalppn = formatnumber(rsVoucher("totalppn"),0)
	selectedfcba = rsVoucher("fcba")
	invoice_by = rsVoucher("invoice_by")
	pocode = rsVoucher("pocode")
	refinvoice = rsVoucher("refinvoice")
	selectedpocode = rsVoucher("pocode") 
	selectedinvoice = rsVoucher("invoice_by")  
	downpayment = rsVoucher("dp") 
	amount = rsVoucher("amount")
	vat = rsVoucher("vat")
	amountvat = formatnumber(rsVoucher("amountvat"),0)
	otherchargestock = rsVoucher("otherchargestock")
	othercharge = rsVoucher("othercharge")
	totalamount = formatnumber(rsVoucher("totalamount"),0)
	no_faktur = rsVoucher("no_faktur")
	fakturdate = rsVoucher("fakturdate")  
	invoicedatesupplier = rsVoucher("invoicedatesupplier")  
	top = rsVoucher("top")  
	dphapus = rsVoucher("ceklistdp")  
	
	HeaderClose = "Y"
 End if
rsVoucher.Close
Set rsVoucher = Nothing

strSuppyingDeclaration	=	" select *" & _
							" from invoicedetail " & _
							" where invoicecode = '" & invoicecode & "'" 
						
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	invoicecode = rsSupplyingDeclaration("invoicecode")
	itemcode = rsSupplyingDeclaration("itemcode")
	quantity = rsSupplyingDeclaration("quantity")
	amount = rsSupplyingDeclaration("amount")
	pocode = rsSupplyingDeclaration("pocode")
	receivenotecode = rsSupplyingDeclaration("receivenotecode")
	purchaselinenokey = rsSupplyingDeclaration("purchaselinenokey")
	amountbeforevat = rsSupplyingDeclaration("amountbeforevat") 
	SudahClose = "Y"
End if
' response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

 '-- Validasi untuk other charge Add By SAG
 q = 	" SELECT NVL(SUM (inv.othercharge),0) othercharge" & _
		"   FROM invoice inv" & _
		"  WHERE inv.fcba = '"& fcba &"' AND inv.pocode = '"& pocode &"'"  
 ' response.write q
 set rsVoucher = server.CreateObject("Adodb.recordset")
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    v_othercharge = rsVoucher("othercharge")
 End if
 rsVoucher.Close : Set rsVoucher = Nothing
 
 q = 	" select" & _
		" (nvl(po.otch_nonstock_po,0)-nvl(inv.otch_nonstock_inv,0)) as totalcharge," & _
		" nvl(otch_stock_po,0) otch_stock_po " & _
		" from (" & _
		"     select pocode,sum(othercharge) as otch_nonstock_inv" & _
		"     from invoice" & _
		"     group by pocode" & _
		" ) inv" & _
		" right join (" & _
		"     select pocode,sum(othercharge) as otch_nonstock_po,nvl(sum(otherchargestock),0) otch_stock_po from purchaseorderdetails" & _
		"     group by pocode" & _
		" ) po" & _
		" on (inv.pocode = po.pocode)" & _
		" where po.pocode = '"& pocode &"'"   

 ' response.write q
 set rsVoucher = server.CreateObject("Adodb.recordset")
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
	v_totalcharge = rsVoucher("totalcharge")
 End if
 rsVoucher.Close : Set rsVoucher = Nothing
 '-- End Added
 
 '-- Validasi untuk  VAT Add By SAG
 q = 	" SELECT NVL(SUM (po.totalppn),0) totalppn" & _
		"   FROM purchaseorder po" & _
		"  WHERE po.fcba = '"& fcba &"' AND po.pocode = '"& pocode &"'"  
 ' response.write q
 set rsVoucher = server.CreateObject("Adodb.recordset")
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    v_totalppn = rsVoucher("totalppn")
 End if
 rsVoucher.Close : Set rsVoucher = Nothing
 '-- End Added
 
 if err.number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
 End if
 '-------------ISI Payment Form
 
Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Purchase')"
	   
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if

oRecordSet.Close
Set oRecordSet = Nothing
 %>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Invoice</div><% end if %>
 <form id="formentry" name="frmEntry" method="post" action="Invoice.Asp" >
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Invoice No</td>
			<%
			if fcCode = "(NEW)" and Len(Trim(fcba))  <> 0  then
				q = " select MAX(cn.PREFIX) || LPAD(NVL(MAX(regexp_substr(inv.INVOICECODE, '[^/]+', 1, 4)),0000)+0001,4,0) DOCNUMBER  " & _
					" from  " & _
					" controlnumber cn  " & _
					" left join invoice inv on cn.fcba = inv.fcba and cn.accountingyear = to_number(to_char(inv.INVOICEDATE,'yyyy'))  " & _
					" where cn.fcba ='" & fcba & "' " & _
					" and cn.LU_NAME ='INVOICE' "&_
					" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yyyy')) " & _
					" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'mm')) "
				'Response.write q
				set rsPONo = server.CreateObject("Adodb.recordset")
				rsPONo.Open q, connstr
				if Not rsPONo.Eof then
					fcCode = rsPONo("DOCNUMBER")
				End if
			End if
			%>
			<td align="left"><input type="text" maxlength="50" id="fcCode" name="fcCode" size="30" value="<%=fcCode%>" autocomplete="off"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">DP</td>
			<td align="right">
			
				<input type="text" id="downpayment" name="downpayment" size="30" value="<%=downpayment%>" onkeyup="hanyaAngka(this);" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
				<input type="hidden"  name="dphapus" size="30" value="0">
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Ref. Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="refinvoice" size="50" value="<%=refinvoice%>"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">Other Charge Stock</td>
			<td align="left"><input type="text" id="otherchargestock" name="otherchargestock" size="30" value="<%=otherchargestock%>"   style="text-align: right;" style="background-color:#eee;"></td>
		</tr>
		<tr>
			<td align="left">Entry Date</td>
			<td align="left">
				<input onchange="javascript:reloadInvCode();" type="text" maxlength="10" name="DATECREATED" size="10" value="<%=DATECREATED%>" readonly="true" style="background-color:#eee;">
				<font color="#FF0000">&nbsp;*</font>TOP <input type="text" id="top" name="top" size="5" value="<%=top%>">
			</td>
			<td align="left" style="width:20%;">Amount</td>
			<td align="left"><input type="text" id="amount" name="amount" size="30" value="<%=amount%>"  readonly="true" style="background-color:#eee;"></td>	
			</td>
		</tr>
		<tr>
			<td align="left">Invoice Date</td>
			<td align="left">
				<input class="datepicker"  type="text" maxlength="10" name="invoicedatesupplier" size="10" value="<%=invoicedatesupplier%>">
				<font color="#FF0000">&nbsp;*</font> 
			</td>
			<td align="left" style="width:20%;"></td>
			<td align="left"></td>	
			</td>
		</tr>
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();javascript:getlistinvoice();" class="chosen-select" style="width: 70%">
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
			<td align="left" style="width:20%;">VAT</td>
			<td align="left">
			    <input type="checkbox" id="use_own_ppn" name="use_own_ppn" value="<%=use_own_ppn%>" <%if cstr(""&use_own_ppn) = "1" then%> checked <%end if%> onchange="handleClick(this);"> OWN PPN
				<br>
				<input type="text" id="vat" name="vat" size="2" value="<%=vat%>" >&nbsp;<input type="text" id="totalppn" name="totalppn" size="20" value="<%=totalppn%>" style="text-align: right;">
			</td> 
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<div id="fcbalist"></div>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			<td align="left" style="width:20%;">Amount + VAT</td>
			<td align="left"><input type="text" id="amountvat" name="amountvat" size="30" value="<%=amountvat%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;" ></td>		
		</tr>
		<tr>
			<td align="left">PO Code</td>
			<td align="left">      
			<%if HeaderClose="Y" then%>
				<input type="text" id="pocode" name="pocode" value="<%=pocode%>" readonly style="background-color: #eee;"> 
			<%else%>
				<div id="pocodelist"></div>
			<%End if%>
			<input type="Hidden" id="selectedpocode" name="selectedpocode" value="<%=selectedpocode%>">
			<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left" style="width:20%;"> Other Charge Non Stock</td>
			<td align="left">
				<%if HeaderClose<>"Y" then%>
					<input type="text" name="othercharge" id="othercharge" size="30" value="<%=othercharge%>" autocomplete="off" style="text-align: right;" >
				<%else
				
				dim sqlothercarge,rsothercarge,txtothercarge
				sqlothercarge = "select sum(othercharge) othercharge from PURCHASEORDERDETAILS where POCODE='"& pocode &"' AND fcBA = '"& fcba &"'"
				Set rsothercarge = Server.CreateObject("ADODB.Recordset")
   				rsothercarge.Open sqlothercarge, connstr
				if not rsothercarge.eof then  
					othercharge = rsothercarge("othercharge")
					'response.Write othercharge
				end if
				
				%>
					<input type="text" id="othercharge" name="othercharge" size="30" value="<%=othercharge%>" autocomplete="off" style="text-align: right;" >
				<%End if%>
			</td> 
		</tr>
		<tr>
			<td align="left" width="20%">Invoice By</td>
			<td align="left">
				<div id="invoicelist"></div>
			</td>
				<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="<%=selectedinvoice%>">
			<td align="left" style="width:20%;">Total Amount</td>
			<td align="left"><input type="text" id="totalamount" name="totalamount" size="30" value="<%=totalamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 2, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>
			
		</tr>
		<tr>
			<td align="left" style="width:20%;">Supplier Code</td>
			<td align="left"><input type="text" id="suppliercode" name="suppliercode" size="30" value="<%=suppliercode%>" readonly="true" style="background-color:#eee;style="width: 70%""></td>
			<td align="left">No Faktur</td>
			<td align="left">
				<input type="text" maxlength="100" name="no_faktur" size="30" value="<%=no_faktur%>" autocomplete="off" >
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Supplier Name</td>
			<td align="left"><input type="text" id="suppliername" name="suppliername" size="30" value="<%=suppliername%>" readonly="true" style="background-color:#eee;"></td>
			<td align="left"> Faktur Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="fakturdate" size="10" value="<%=fakturdate%>" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td colspan="4"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="printinvoice.asp?fcCode=<%=fcCode%>&fcba=<%=fcba%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<a href="Invoice.asp?fcCode=<%=fcCode%>&fcba=<%=fcba%>&DATECREATED=<%=DATECREATED%>&txtJenis=Posting" onClick="return confirm('Are you sure to Posting Journal this data :  <%=fccode%> ?');">Posting</a>&nbsp;&nbsp;
				<a href="Invoice.Asp?fcCode=<%=fcCode%>&fcba=<%=fcba%>&DATECREATED=<%=DATECREATED%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0"></a>
				<a href="invoice_editheader.asp?fcCode=<%=fcCode%>&fcba=<%=fcba%>" target="_blank"><font face="Verdana" size="1">Edit Header</font></a>&nbsp;&nbsp;
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Purchase"
				dataSubmodule 	= "Invoice"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td>
		</tr>
    </table>  
 </form>
<% if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 and len(trim(pocode)) > 0 and len(trim(invoice_by)) > 0 Then  
  Dim rsGRN,intRecordsPerPage,cColor
If HeaderClose = "Y" Then
		q = " select sp.fcname suppdesc, dtl.itemcode, s.fcname stockdesc, dtl.quantity, dtl.amount, dtl.receivenotecode, dtl.purchaselinenokey" & _
			" FROM invoicedetail dtl" & _
			" INNER JOIN invoice hdr ON (dtl.invoicecode = hdr.invoicecode AND dtl.fcba = hdr.fcba)" & _
			" LEFT JOIN stock s ON (dtl.itemcode = s.fccode AND dtl.fcba = s.fcba)" & _
			" LEFT JOIN supplier sp ON (hdr.suppliercode = sp.fccode AND hdr.fcba = sp.fcba)" & _
			" WHERE dtl.INVOICECODE = '" & fcCode & "'" & _
			" AND dtl.fcBA = '" & fcba & "'"
		  
	 '  response.Write q	
	   Set rsGRN = Server.CreateObject("ADODB.Recordset")
	   rsGRN.Open q, connstr 
	   
%>
 <form action="Invoice.Asp" method="post" name="eForm">
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines" border="1">
	<thead>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">&nbsp;</td>
	   <td align="center">GRN Code</td>
	   <td align="center">Stock Code</td>
	   <td align="center">Stock Desc.</td>
	   <td align="center">Quantity</td>
	   <td align="center">Amount</td>
	   </tr>
	</thead>
	<tbody>
   <% if Not rsGRN.eof then
      nCountofField=rsGRN.fields.count-1
	  no=1
	  v_totalamountbeforesubmit = 0
	  v_totalamountaftersubmit = 0 
	  While Not rsGRN.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="ringht" valign="top"><%=NO%></td> 
				<td valign="middle">
					<a href="Invoice.Asp?fcCode=<%=fcCode%>&pocode=<%=pocode%>&fcba=<%=fcba%>&receivenotecode=<%=rsGRN("receivenotecode")%>&itemcode=<%=rsGRN("itemcode")%>&purchaselinenokey=<%=rsGRN("purchaselinenokey")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0"></a></td>
					<td align="left" title="Supplier" valign="top"><%=rsGRN("receivenotecode")%></td>
					<td align="left" title="Item Code" valign="top"><%=rsGRN("ITEMCODE")%></td>
					<td align="left" title="Item Name" valign="top"><%=rsGRN("StockDesc")%></td>
					<td align="right" title="Quantity" valign="top"><%=rsGRN("Quantity")%></td>
					<td align="right" title="Amount" valign="top"><%=rsGRN("amount")%></td>
			 
		  </tr>
		  <% 
			v_totalamountaftersubmit = v_totalamountaftersubmit + Trim(cdbl(rsGRN("amount")))
		  
		  rsGRN.MoveNext 
		  no=no+1
		  intRecordsPerPage = intRecordsPerPage - 1
	  Wend
   else
      response.Write("<tr><td colspan=""7"" class=""fbp-doc-empty"">Data is empty</td></tr>")	  
   end if 
   rsGRN.Close
   set rsGRN = Nothing%>
		<tr> 
				<td align="right" title="" valign="top" colspan="6"><font face="Verdana" size="1" style="font-weight: bold">TOTAL</font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_totalamountaftersubmit),2)%> </font></td> 
		 
		</tr>
	  <tr><td colspan="11" align="left">
		<!-- diambil dari form yang diatas -->
	    <input type="hidden" name="fcCode" value="<%=fcCode%>">
		<input type="hidden" name="refinvoice" value="<%=refinvoice%>">
		<input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		<input type="hidden" name="pocode" value="<%=pocode%>">
		<input type="hidden" name="No" value="<%=No%>">
		<input type="hidden" name="totalppn" value="<%=totalppn%>"> 
		<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		<input type="hidden" name="fcba" value="<%=fcba%>">
		<input type="hidden" name="invoice_by" value="<%=invoice_by%>">
		<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		<input type="hidden" name="selectedpocode" value="<%=selectedpocode%>">
		<input type="hidden" name="selectedinvoice" value="<%=selectedinvoice%>">
		<input type="hidden" name="othercharge" value="<%=othercharge%>">
		<input type="hidden" name="downpayment" value="<%=downpayment%>">
		<input type="hidden" name="amount" value="<%=amount%>">
		<input type="hidden" name="vat" value="<%=vat%>">
		<input type="hidden" name="amountvat" value="<%=amountvat%>">
		<input type="hidden" name="totalamount" value="<%=totalamount%>"> 
		<input type="hidden" name="suppliercode" value="<%=suppliercode%>">
		<input type="hidden" name="suppliername" value="<%=suppliername%>">
		<input type="hidden" name="no_faktur" value="<%=no_faktur%>">
		<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
		<input type="hidden" name="invoicedatesupplier" value="<%=invoicedatesupplier%>">
		<input type="hidden" name="top" value="<%=top%>">
		  
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
   </tbody>
   </TABLE>
   </div>
   
   </form>
<%Else%>
 
<% 
if cdbl(to_nol(v_othercharge)) + cdbl(to_nol(othercharge)) <= cdbl(to_nol(v_totalcharge)) then
	if cdbl(to_nol(totalppn)) <=  cdbl(to_nol(v_totalppn)) then
		if SudahClose="Y" Then 
				' response.write "<br/>" & v_othercharge
				' response.write "<br/>" & othercharge
				' response.write "<br/>" & v_totalcharge 
				  q = " select sp.fcname suppdesc, dtl.itemcode, s.fcname stockdesc, dtl.quantity, dtl.amount, dtl.receivenotecode, dtl.purchaselinenokey" & _
					  " FROM invoicedetail dtl" & _
					  " INNER JOIN invoice hdr ON (dtl.invoicecode = hdr.invoicecode AND dtl.fcba = hdr.fcba)" & _
					  " LEFT JOIN stock s ON (dtl.itemcode = s.fccode AND dtl.fcba = s.fcba)" & _
					  " LEFT JOIN supplier sp ON (hdr.suppliercode = sp.fccode AND hdr.fcba = sp.fcba)" & _
					  " WHERE dtl.INVOICECODE = '" & fcCode & "'" & _
					  "       AND dtl.fcBA = '" & fcba & "'"
		Else	 
			q =	"   SELECT hdr.pocode," & _
				"          dtl.receivenotecode," & _
				"          dtl.purchaselinenokey," & _
				"          dtl.purchaseitemcode," & _
				"          s.fcname," & _
				"          CASE" & _
				"             WHEN (NVL (dtl.quantity, 0) - NVL (ind.quantity, 0)) = 0" & _
				"             THEN" & _
				"                0" & _
				"             ELSE" & _
				"                (NVL (dtl.quantity, 0) - NVL (ind.quantity, 0))" & _
				"          END" & _
				"             quantity," & _
				"          NVL (dtl.amount, 0) amount," & _
				"          CASE" & _
				"             WHEN (NVL (dtl.amountpo_receive, 0) - NVL (ind.amount, 0)) = 0" & _
				"             THEN" & _
				"                0" & _
				"             ELSE" & _
				"                ROUND((NVL (dtl.amountpo_receive, 0) - NVL (ind.amount, 0)))" & _
				"          END" & _
				"             amountreceived" & _
				"     FROM receivenotedetail dtl" & _
				"          INNER JOIN receivenote hdr" & _
				"             ON (    hdr.receivenotecode = dtl.receivenotecode" & _
				"                 AND hdr.fcba = dtl.fcba)" & _
				"          INNER JOIN stock s" & _
				"             ON (s.fcba = dtl.fcba AND s.fccode = dtl.purchaseitemcode)" & _
				"          LEFT JOIN (  SELECT fcba," & _
				"                              receivenotecode," & _
				"                              purchaselinenokey," & _
				"                              itemcode," & _
				"                              SUM (quantity) AS quantity," & _
				"                              SUM (amount) AS amount" & _
				"                         FROM invoicedetail" & _
				"                        WHERE pocode = '"& pocode &"'" & _
				"                     GROUP BY fcba," & _
				"                              receivenotecode," & _
				"                              purchaselinenokey," & _
				"                              itemcode) ind" & _
				"             ON (    ind.receivenotecode = dtl.receivenotecode" & _
				"                 AND ind.fcba = dtl.fcba" & _
				"                 AND dtl.purchaselinenokey = ind.purchaselinenokey)" & _
				"    WHERE hdr.pocode = '"& pocode &"' AND dtl.fcBA = '"& fcba &"' AND hdr.rowstateid='Approved'" & _
				" ORDER BY dtl.receivenotecode, dtl.purchaseitemcode" 
			
		
				

	   End if
   ' response.Write q	
   'response.Write "Event awal saat select po: " & q & "<br>"
   
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.Open q, connstr
 %>
   <form action="Invoice.Asp" method="post" name="eForm">
   <table class="data" border="1" width="100%">
   <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">PO Code</td>
	   <td align="center">GRN Code</td>
	   <td align="center">Stock Code</td>
	   <td align="center">Stock Desc.</td>
	   <td align="center">Qty Received</td>
	   <td align="center">Amount PO Receive</td>
	   <td align="center"><input type="checkbox" name="select-all" id="select-all"/> Check/Uncheck All</td>
	   <td align="center">Amount Invoice</td>
	   </tr>
   <%ELSE%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">&nbsp;</td>
	   <td align="center">GRN Code</td>
	   <td align="center">Stock Code</td>
	   <td align="center">Stock Desc.</td>
	   <td align="center">Quantity</td>
	   <td align="center">Amount</td>
	   </tr>
   <%End if
   if Not rsGRN.eof then
      nCountofField=rsGRN.fields.count-1
	  no=1
	  v_totalamountbeforesubmit = 0
	  v_totalamountaftersubmit = 0
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
					<a href="Invoice.Asp?fcCode=<%=fcCode%>&pocode=<%=pocode%>&fcba=<%=fcba%>&receivenotecode=<%=rsGRN("receivenotecode")%>&itemcode=<%=rsGRN("itemcode")%>&purchaselinenokey=<%=rsGRN("purchaselinenokey")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0"></a></td>
					<td align="center" title="Supplier" valign="top"><%=rsGRN("receivenotecode")%></td>
					<td align="center" title="Item Code" valign="top"><%=rsGRN("ITEMCODE")%></td>
					<td align="left" title="Item Name" valign="top"><%=rsGRN("StockDesc")%></td>
					<td align="center" title="Quantity" valign="top"><%=rsGRN("Quantity")%></td>
					<td align="right" title="Amount" valign="top"><%=rsGRN("amount")%></td>
			<%End if%>

			<%if SudahClose<>"Y" Then%>
                    <td align="center" title="" valign="top"><%=rsGRN("POCODE")%></td>
			        <td align="center" title="" valign="top"><%=rsGRN("RECEIVENOTECODE")%></td>
			        <td align="center" title="" valign="top"><%=rsGRN("PURCHASEITEMCODE")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("fcname")%></td>
			        <td align="right" title="" valign="top">
			        	<input style="text-align:right;background-color:#eee;" type="text" name="Quantity_<%=no%>" class="quantity" value = "<%=rsGRN("Quantity")%>" readonly>
			        </td>
                    <td align="right" title="" valign="top"><%=rsGRN("amountreceived")%></td>
		            <td align="center">
						<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
					</td>
					<td align="center">
						<input style="text-align:right;background-color:#eee;" type="text" name="AmountInvoice_<%=no%>" class="amount" value = "<%=rsGRN("amountreceived")%>"  autocomplete="off" readonly>
					</td>
					<input type="hidden" name="receivenotecode_<%=no%>" value="<%=rsGRN("receivenotecode")%>">
					<input type="hidden" name="linenokey_<%=no%>" value="<%=rsGRN("purchaselinenokey")%>">
					<input type="hidden" name="Stock_<%=no%>" value="<%=rsGRN("purchaseitemcode")%>">
					<input type="hidden" name="amount_<%=no%>" value="<%=rsGRN("amount")%>">
				 </td> 
			<%End if%>
		  </tr>
		  <%
		  IF SudahClose="Y" then
			v_totalamountaftersubmit = v_totalamountaftersubmit + Trim(cdbl(rsGRN("amount")))
		  ELSE
			v_totalamountbeforesubmit = v_totalamountbeforesubmit + Trim(cdbl(rsGRN("amountreceived")))
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
			<%IF SudahClose="Y" then%>
				<td align="right" title="" valign="top" colspan="6"><font face="Verdana" size="1" style="font-weight: bold">TOTAL</font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_totalamountaftersubmit),2)%> </font></td>
			<%Else%>
				<td align="right" title="" valign="top" colspan="8"><font face="Verdana" size="1" style="font-weight: bold">TOTAL</font></td>
				<td align="right" title="" valign="top"><font face="Verdana" size="1" style="font-weight: bold"><%=FormatNumber(cdbl(v_totalamountbeforesubmit),0)%> </font></td>
			<%End if%>
		</tr>
	  <tr><td colspan="11" align="left">
		<!-- diambil dari form yang diatas -->
	    <input type="hidden" name="fcCode" value="<%=fcCode%>">
		<input type="hidden" name="refinvoice" value="<%=refinvoice%>">
		<input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		<input type="hidden" name="pocode" value="<%=pocode%>">
		<input type="hidden" name="No" value="<%=No%>">
		<input type="hidden" name="totalppn" value="<%=totalppn%>"> 
		<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		<input type="hidden" name="fcba" value="<%=fcba%>">
		<input type="hidden" name="invoice_by" value="<%=invoice_by%>">
		<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		<input type="hidden" name="selectedpocode" value="<%=selectedpocode%>">
		<input type="hidden" name="selectedinvoice" value="<%=selectedinvoice%>">
		<input type="hidden" name="othercharge" value="<%=othercharge%>">
		<input type="hidden" name="downpayment" value="<%=downpayment%>">
		<input type="hidden" name="amount" value="<%=amount%>">
		<input type="hidden" name="vat" value="<%=vat%>">
		<input type="hidden" name="amountvat" value="<%=amountvat%>">
		<input type="hidden" name="totalamount" value="<%=totalamount%>"> 
		<input type="hidden" name="suppliercode" value="<%=suppliercode%>">
		<input type="hidden" name="suppliername" value="<%=suppliername%>">
		<input type="hidden" name="no_faktur" value="<%=no_faktur%>">
		<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
		<input type="hidden" name="use_own_ppn" value="<%=use_own_ppn%>">
		<input type="hidden" name="otherchargestock" value="<%=otherchargestock%>">
		<input type="hidden" name="invoicedatesupplier" value="<%=invoicedatesupplier%>">
		<input type="hidden" name="top" value="<%=top%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <!-- Validasi Period Disini -->
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="text" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Submit">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </TABLE>
   
   </form>
   <%
		else
				tempStringError = "Total Amount VAT Lebih Dari Total Amount VAT Purchase Order !"				 
				response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
			End If
   else
		tempStringError = "Other Charge Non Stock yang diinput Lebih besar Dari Other Charge Non Stock di Purchase Order !"				 
		response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
	End If%>
<%End if%>  
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
<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = $("#fcba").val();
}

function getselectedpocode() {
	var spocode = document.getElementById("selectedpocode");
	spocode.value = document.getElementById("pocode").value;
}

function getselectedinvoice() {
	var sfcba = document.getElementById("selectedinvoice");
	sfcba.value = $("#invoice_by").val();
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
	getlistpocode();
	//getpoinfo();
	getselectedfcba();
	getselectedinvoice();
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
		var txts = chk.parentNode.parentNode.cells[5].getElementsByTagName('input')[0];
		
		if(chk.checked == true)
		{
			txt.readOnly = false;
			txt.style.backgroundColor = "#fff";
		}
		else
		{
			txt.readOnly = true;
			txt.style.backgroundColor = "#eee";
		}
		if(chk.checked == true)
		{
			txts.readOnly = false;
			txts.style.backgroundColor = "#fff";
		}
		else
		{
			txts.readOnly = true;
			txts.style.backgroundColor = "#eee";
		}
	}

	function handleClickDP()
	{
		var downpayment = $("#downpayment");
		if (!downpayment.data("original")) {
        	downpayment.data("original", downpayment.val());
    	}

	
		if(dphapus.checked == 1){
			$("#downpayment").val(0);
		} else {
			 downpayment.val(downpayment.data("original"));
		}
	}
</script>
<script type="text/javascript">	
	$('#select-all').click( function () {
		$('.item').prop('checked', this.checked);
		$('.item').val(this.checked ? 1 : 0);
		if($('#select-all').is(":checked")){
			$('.amount').css("background-color", "#fff");
			$('.amount').prop('readOnly', false);
			$('.quantity').css("background-color", "#fff");
			$('.quantity').prop('readOnly', false);
		}else{
			$('.amount').css("background-color", "#eee");
			$('.amount').prop('readOnly', true);
			$('.quantity').css("background-color", "#eee");
			$('.quantity').prop('readOnly', true);
		}
	})
</script>
</body>
</html>
