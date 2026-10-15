<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Invoice STO</title>
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
	doAJAXCall('ajax/getlistfcbainvoicesto.asp', 'POST', PostStr, showlistfcba);
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
	doAJAXCall('ajax/getpostoinvoice.asp', 'POST', PostStr, showlistpocode);
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

var getliststono = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedpocode = document.getElementById("selectedpocode").value;
	var selectedstono = document.getElementById("selectedstono").value;
	var PostStr = 'fcba='+selectedfcba+'&pocode='+selectedpocode+'&stono='+selectedstono;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getliststonoinvoice.asp', 'POST', PostStr, showliststono);
}

// The function for handling the response from the server
var showliststono = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("stolist").innerHTML = "";
	document.getElementById("stolist").innerHTML = response;

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

var getstoinfo = function () {
	// build up the post string when passing variables to the server side page

	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedstono = document.getElementById("selectedstono").value;
	var PostStr = 'fcba='+selectedfcba+'&stono='+selectedstono;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetSTOInfo.asp', 'POST', PostStr, showstoinfo);
}

// The function for handling the response from the server
var showstoinfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    //alert (s.transportercode);
    
    // update the Div to show the result from the server
	document.getElementById("transportercode").value = s.transportercode;
	document.getElementById("transportername").value = s.transportername;
	document.getElementById("dp").value = s.dp;
	document.getElementById("pph").value = s.pph;
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
 
 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks
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
 if Request("DATECREATED")<>"" then
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
transportercode = Request("transportercode")
downpayment = Request("downpayment")
othercharge = Request("othercharge")
stono = Request("stono")
do_no = Request("do_no")
selectedstono = Request("selectedstono")
pph = Request("pph")
amountpph = Request("amountpph")
transportername = Request("transportername")
invoicecode = Request("invoicecode")
rowstate = Request("rowstate")

sub post_amount
	p_amountbeforevat = 0 : p_othercharge = 0 : p_totalamount = 0
	recordTotalAmount = " SELECT NVL(SUM(amount), 0) totalAmount" & _
						" FROM invoicestodetail" & _
						" WHERE invoicecode = '" & fcCode & "'" & _ 
						"       AND fcba = '" & fcba & "'"
	' response.write recordTotalAmount
	Set rsrecordTotalAmount = Server.CreateObject("ADODB.Recordset")
	rsrecordTotalAmount.Open recordTotalAmount, connstr

	if Not rsrecordTotalAmount.Eof then
		p_totalamount = cdbl(to_nol(rsrecordTotalAmount("totalAmount")))
	End if
	' response.write pph
	' amount = p_totalamount
	totalamountpph = p_totalamount *cdbl(to_nol(pph))/100
	totalamount = p_totalamount - totalamountpph

	q = " UPDATE invoicesto" & _
		" SET amountpph = '" & totalamountpph & "'" & _
		"     , totalamount = '" & totalamount & "'" & _
		"     , amount = '" & p_totalamount & "'" & _
		"     , lastupdate = SYSDATE " & _
		"     , lasttime = TO_CHAR(SYSDATE,'HH24:MI') "& _	
		"	  , fcedit='" & session("CodeName") & "'" & _ 		
		" WHERE invoicecode = '" & fcCode & "'" & _
		"       AND fcba = '" & fcba & "'"
		' response.write q
	connStr.execute(q)

end sub
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	
	REMARKS=request("REMARKS")
	
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
			' arr_dtl_eof = true
			' q = " SELECT receivenotecode, pocode, stono FROM invoicestodetail WHERE invoicecode = '" & fcCode & "' AND fcba = '" & fcba & "'"
			' Set rsPO = Server.CreateObject("ADODB.Recordset")
			' rsPO.Open q, connstr
			' if not rsPO.eof then
			' 	arr_dtl = rsPO.getrows()
			' 	arr_dtl_eof = false
			' end if
			' rsPO.close : set rsPO = nothing
			
			' if not arr_dtl_eof then
			' 	for xyz=0 to ubound(arr_dtl,2)
			' 		q = " DELETE FROM invoicedetail" & _
			' 			" WHERE invoicecode = '" & fcCode & "'" & _
			' 			"       AND receivenotecode = '" & arr_dtl(0,xyz) & "'" & _
			' 			"       AND pocode = '" & arr_dtl(1,xyz) & "'" & _
			' 			"       AND purchaselinenokey = '" & arr_dtl(2,xyz) & "'" & _
			' 			"       AND fcBA = '" & fcba & "'"
			' 		'response.write q
			' 		connStr.execute(q)
					
			' 		q = " UPDATE receivenotedetail SET amountinvoice = 0" & _
			' 			" WHERE receivenotecode = '" & arr_dtl(0,xyz) & "'" & _
			' 			"       AND purchaselinenokey = '" & arr_dtl(2,xyz) & "'" & _
			' 			"       AND fcBA = '" & fcba & "'"
			' 		'response.write q
			' 		connStr.execute(q)
			' 	next
			' 	erase arr_dtl
			' end if
			' arr_dtl_eof = true

			' q = " DELETE FROM invoice WHERE invoicecode = '" & fcCode & "' AND fcba = '" & fcba & "'"
			' connStr.execute(q)
			' strFieldDetail = " SELECT COUNT(1) total FROM invoicestodetail " & _
	  '                        " WHERE invoicecode = '" & fccode & "'" & _
	  '                        "       AND fcBA = '" & fcba & "'"								
			' ' Response.Write strFieldDetail
			' Set rsQty = Server.CreateObject("ADODB.Recordset")
			' rsQty.Open strFieldDetail, connstr							
			' if not rsQty.Eof then
			' 	has_detail = cdbl (to_nol(rsQty("total")))
			' End If
			
			' if has_detail = 0 then
				q = " DELETE FROM invoicesto WHERE invoicecode = '" & fccode & "' AND fcba = '"& fcba &"' and pocode = '"& pocode &"' and stono = '"& stono &"'"
					' Response.Write q
				connStr.execute(q)

				q = " DELETE FROM invoicestodetail" & _
				    " WHERE invoicecode = '" & fcCode & "'" & _
				    "       AND stono = '" & stono & "'" & _
				    "       AND pocode = '" & pocode & "'" & _
				    "       and fcBA='" & fcba & "'"
				' response.write q
				connStr.execute(q)
			' else
			' 	err.number = 1001
			' 	err.description = " STO Invoice """ & fcCode & """ has detail, please delete all detail first!"
			' End If
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		
		cmdView	= "True"

	elseif Request("txtJenis")= "Details" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then

			q = " DELETE FROM invoicestodetail" & _
			    " WHERE invoicecode = '" & fcCode & "'" & _
			    "       AND stono = '" & stono & "'" & _
			    "       AND pocode = '" & pocode & "'" & _
			    "       and fcBA='" & fcba & "'"
			' response.write q
			connStr.execute(q)
			
			' q = " UPDATE receivenotedetail SET amountinvoice = 0" & _
			'     " WHERE receivenotecode = '" & Request("receivenotecode") & "'" & _
			'     "       AND purchaselinenokey = '" & Request("purchaselinenokey") & "'" & _
			'     "       and fcBA='" & fcba & "'"
			' 'response.write q
			' connStr.execute(q)
			
			' Call post_amount
			
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
				If Ucase(fccode) = Ucase("(NEW)") Then
					oSql = " select nvl(controlnumber_api.get_nextnumber(upper('INVOICESTO')" & _
						   "                                         , " & year(textDate(DATECREATED)) & "" & _
						   "                                         , " & month(textDate(DATECREATED)) & "" & _
						   "                                         , '" & fcba & "'" & _
						   " ), 'FALSE') INVOICESTO from dual"
					
					' Response.Write oSql
						   
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.Open oSql, connstr
					If Not oRs.EOF Then
						txtINVOICE = oRs("INVOICESTO")
					End If
					oRs.Close
					Set oRs = Nothing
						fccode = txtINVOICE
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
				q = " Insert into Invoicesto (INVOICECODE, INVOICE_BY, REFINVOICE, INVOICEDATE, TRANSPORTERCODE, PAID, REMARKS, NETBALANCE, pocode, stono, no_faktur, fakturdate, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, pph, dp, amount, amountpph, totalamount, rowstate) " & _
					" values ('" & fcCode & "'"& _
					" ,'" & invoice_by & "'"& _
					" ,'" & refinvoice & "'"& _
					" ," & " to_date('" & DATECREATED & "','mm/dd/yyyy')"& _
					" ,'" & TRANSPORTERCODE& "'"& _
					" ,0"& _
					" ,'"& Remarks &"'" & _
					" ,0"& _
					" ,'" & pocode& "'" & _
					" ,'" & stono & "'" & _
					" ,'" & no_faktur & "'" & _
					" ," & " TO_DATE('" & fakturdate & "','mm/dd/yyyy')" & _
					" ,'" & Session("CodeName") & "'" & _
					" ,'" & Session("CodeName") & "'" & _
					" ,'" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" ,'" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI') "& _
					" ,'" & pph & "'"& _
					" , REPLACE('" & downpayment & "',',','')" & _
					" , REPLACE('" & amount & "',',','')" & _
					" , REPLACE('" & amountpph & "',',','')" & _
					" , REPLACE('" & totalamount & "',',','')" & _
					" ,'INVOICED')"
				' response.write q
				connStr.execute(q)
			End if
			rsPO.Close	
			Dim n

			' Added by NAF, 2012-12-01
			 strLineMax = " select nvl(max(lineno),0) lineno from invoicestodetail where fcBA='" & fcba & "' and invoicecode = '"& fcCode &"' and stono='"& stono &"'"
			' response.write strLineMax
			Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
			rsstrLineMax.Open strLineMax, connstr
			
			if not rsstrLineMax.eof then
				tempMaxLineNo = cdbl(rsstrLineMax("lineno")) 
			end if
						
			rsstrLineMax.Close
			set rsstrLineMax = Nothing 
			' ' End of Added by NAF, 2012-12-01

			'response.Write "No " & (cdbl(request("no"))-1) & " amount : " &  request("AmountInvoice_"&TRIM(n))
			For n=1 to (cdbl(request("no"))-1) 
				if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
					tempMaxLineNo = tempMaxLineNo + 1
					' p_amountbeforevat = 0
					' p_amountbeforevat = Cdbl(to_nol(REPLACE(request("AmountInvoice_"&TRIM(n)), ",", ""))) -Cdbl(to_nol(REPLACE(request("othercharge_"&TRIM(n)), ",", "")))
						q = " Insert into  INVOICESTODETAIL (INVOICECODE, POCODE, STONO, DO_NO, RECEIVENOTECODE, ITEMCODE, QUANTITY, AMOUNT, UNITPRICE, LINENO, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
							" values ('" & fcCode & "'" & _
							" ,'" & pocode & "'" & _
							" ,'" & stono & "'"& _
							" ,'" & request("do_no_"&TRIM(n)) & "'" & _
							" ,'" & request("receivenotecode_"&TRIM(n)) & "'" & _
							" ,'" & request("itemcode_"&TRIM(n)) &"'" &  _
							" ,REPLACE('" & request("quantity_"&TRIM(n)) & "', ',', '')" &  _
							" ,REPLACE('" & request("amount_"&TRIM(n)) & "', ',', '')" & _
							" ,REPLACE('" & request("unitprice_"&TRIM(n)) & "',',','')" & _
							" ,'" & tempMaxLineNo & "'"& _
							" ,'" & Session("CodeName") & "'" & _
							" ,'" & Session("CodeName") & "'" & _
							" ,'" & Request.ServerVariables("REMOTE_ADDR") &"'"& _
							" ,'" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						' response.write q
						connStr.execute(q)
						
						q = " Update ReceiveNoteDetail  set AMOUNTSTO_INVOICE =AMOUNTSTO_INVOICE+REPLACE('" & request("amount_"&TRIM(n)) & "', ',', '')"  & _
							", LASTUPDATE = sysdate " & _
							", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							" where RECEIVENOTECODE='" & request("receivenotecode_"&TRIM(n)) & "'" & _
							" and fcBA='" & fcba & "'"
						' response.Write(q)
						connStr.execute(q)
				End if	 
			Next
			
			Call post_amount
			
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

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

 Dim rsVoucher,SudahClose
 q = " select hdr.*, bu.fccompanycode, tr.fcname" & _
     " from Invoicesto hdr" & _
     " left join invoicestodetail isd on (isd.invoicecode = hdr.invoicecode and isd.fcba = hdr.fcba)" & _
     " left join businessunit bu on bu.fccode = hdr.fcba" & _
     " left join transporter tr ON tr.fcba = hdr.fcba and tr.fccode = hdr.transportercode " & _
     " where hdr.invoicecode='" & fcCode & "'"
 ' response.write q
 set rsVoucher = server.CreateObject("Adodb.recordset")
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED = rsVoucher("INVOICEDATE")
	TRANSPORTERCODE = rsVoucher("TRANSPORTERCODE")
    fccompanycode = rsVoucher("fccompanycode")
	fcba = rsVoucher("fcba")
	selectedfcba = rsVoucher("fcba")
	invoice_by = rsVoucher("invoice_by")
	pocode = rsVoucher("pocode")
	refinvoice = rsVoucher("refinvoice")
	selectedpocode = rsVoucher("pocode")
	selectedstono = rsVoucher("stono")
	selectedinvoice = rsVoucher("invoice_by")
	stono = rsVoucher("stono")
	transportername = rsVoucher("fcname")
	downpayment = rsVoucher("dp")
	amount = rsVoucher("amount")
	pph = rsVoucher("pph")
	amountpph = rsVoucher("amountpph")
	totalamount = rsVoucher("totalamount")
	no_faktur = rsVoucher("no_faktur")
	fakturdate = rsVoucher("fakturdate")

	SudahClose = "Y"
else
	pph = 0
	amount = 0
	amountvat = 0
	totalcharge = 0
	totalAmount = 0
End if
 rsVoucher.Close : Set rsVoucher = Nothing
 
' Untuk Validasi DO yang sudah di Invoice
' q = " SELECT COUNT (*) total" & _
'   	" FROM invoicestodetail" & _
'  	" WHERE fcba = '"& fcba &"' AND do_no = '"& do_no &"'"
'  response.write q
'  set rsCek = server.CreateObject("Adodb.recordset")
'  rsCek.Open q, connstr
'  if Not rsCek.Eof then
'     totalDO = rsCek("total")
'  End if
 ' rsCek.Close : Set rsCek = Nothing
' End Add

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
'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing
 %>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Invoice STO</div><% end if %>
 <form name="frmEntry" method="post" action="invoicesto.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>" autocomplete= "off"><font color="#FF0000">&nbsp;*</font></td>
			<td align="left" width="20%">Invoice By
				<td align="left">
					<div id="invoicelist"></div>
				</td>
					<input type="Hidden" id="selectedinvoice" name="selectedinvoice" value="<%=selectedinvoice%>">
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Ref. Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="refinvoice" size="30" value="<%=refinvoice%>" autocomplete="off"><font color="#FF0000">&nbsp;*</font></td>
			<td align="left" style="width:20%;">DP</td>
			<td align="left"><input type="text" id="downpayment" name="downpayment" size="30" value="<%=downpayment%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>			
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="DATECREATED" size="10" value="<%=DATECREATED%>">
			<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left" style="width:20%;">Amount</td>
			<td align="left"><input type="text" id="amount" name="amount" size="30" value="<%=amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>
		</tr>
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();javascript:getlistinvoice();" class="chosen-select" style="width: 50%">
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
			<td align="left" style="width:20%;">PPh 23</td>
			<td align="left"><input type="text" id="pph" name="pph" size="30" value="<%=pph%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;">
			</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<div id="fcbalist"></div>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			<td align="left" style="width:20%;">Amount + PPh</td>
			<td align="left">
				<input type="text" id="amountpph" name="amountpph" size="30" value="<%=amountpph%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;">
			</td>
		</tr>
		<tr>
			<td align="left">PO Code</td>
			<td align="left">      
			<%if SudahClose="Y" then%>
				<div id="pocodelist"></div>
			<%else%>
				<div id="pocodelist"></div>
			<%End if%>
			<input type="Hidden" id="selectedpocode" name="selectedpocode" value="<%=selectedpocode%>">
			<font color="#FF0000">&nbsp;*</font>
			</td>
			<td align="left" style="width:20%;">Total Amount</td>
			<td align="left"><input type="text" id="totalamount" name="totalamount" size="30" value="<%=totalamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">STO No
				<td align="left">
					<div id="stolist"></div>
				</td>
					<input type="Hidden" id="selectedstono" name="selectedstono" value="<%=selectedstono%>">
			</td>
			<td align="left">No Faktur</td>
			<td align="left">
				<input type="text" maxlength="100" name="no_faktur" size="30" value="<%=no_faktur%>" autocomplete="off" >
			</td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Transporter Code</td>
			<td align="left"><input type="text" id="transportercode" name="transportercode" size="30" value="<%=transportercode%>" readonly="true" style="background-color:#eee;"></td>
			<td align="left"> Faktur Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="fakturdate" size="10" value="<%=fakturdate%>" autocomplete="Off">
			</td>	
		</tr>
		<tr>
			<td align="left" style="width:20%;">Transporter Name</td>
			<td align="left"><input type="text" id="transportername" name="transportername" size="30" value="<%=transportername%>" readonly="true" style="background-color:#eee;"></td>
			
		</tr>
		<tr>
			<td colspan="4"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="printinvoicesto.asp?fccode=<%=fcCode%>&fcba=<%=fcba%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<!-- Validate Period Control In Here -->
					<a href="invoicesto.asp?fcCode=<%=fcCode%>&fcba=<%=fcba%>&DATECREATED=<%=DATECREATED%>&stono=<%=stono%>&pocode=<%=pocode%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0"></a>
				<!-- End Validate Period Control In Here -->
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
<%if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 Then
   Dim rsGRN,intRecordsPerPage,cColor
   if SudahClose="Y" Then
       q = 	" SELECT isd.invoicecode," & _
			"        isd.pocode," & _
			"        isd.stono," & _
			"        isd.do_no," & _
			"        isd.receivenotecode," & _
			"        isd.itemcode," & _
			"        dod.stockname," & _
			"        isd.unitprice," & _
			"        isd.quantity," & _
			"        isd.amount" & _
			"   FROM invoicestodetail isd" & _
			"        INNER JOIN deliveryorderpurchasedtl dod" & _
			"           ON (    isd.stono = dod.stono" & _
			"               AND isd.fcba = dod.fcba" & _
			"               AND isd.itemcode = dod.itemcode" & _
			"               AND isd.do_no = dod.do_no)" & _ 
			" WHERE isd.stono = '"& stono &"' AND isd.fcba = '"& fcba &"'"
   Else	 
      q = 	" SELECT sth.stono," & _
			"       sth.pocode," & _
			"       rn.receivenotecode," & _
			"       do.do_no," & _
			"       dod.itemcode," & _
			"       dod.stockname," & _
			"       std.unitprice," & _
			"       rnd.quantity," & _
			"       rnd.amountsto_receive amount" & _
			"  FROM stocktransferorderdtl std" & _
			"       INNER JOIN stocktransferorderhdr sth" & _
			"          ON (sth.fcba = std.fcba AND sth.stono = std.stono)" & _
			"       INNER JOIN deliveryorderpurchasehdr do" & _
			"          ON (do.pocode = sth.pocode AND do.fcba = sth.fcba)" & _
			"       INNER JOIN deliveryorderpurchasedtl dod" & _
			"          ON (do.do_no = dod.do_no AND do.fcba = dod.fcba)" & _
			"       INNER JOIN receivenote rn" & _
			"          ON (rn.fcba = do.fcba AND rn.deliveryordercode = do.do_no)" & _
			"       INNER JOIN receivenotedetail rnd" & _
			"          ON (rn.receivenotecode = rnd.receivenotecode AND rn.fcba = rnd.fcba)" & _
			"		LEFT JOIN invoicesto ist" & _
          	"		   ON (ist.stono = sth.stono AND ist.fcba = sth.fcba)" & _
       		"		LEFT JOIN invoicestodetail isd" & _
          	"		   ON (isd.invoicecode = ist.invoicecode AND isd.fcba = ist.fcba)" & _
			" WHERE sth.stono = '"& stono &"' AND sth.fcba = '"& fcba &"' AND ist.rowstate is null"

   End if
   ' response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.Open q, connstr%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="invoicesto.asp" method="post" name="eForm">
   <table class="data" border="1" width="100%">
   <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">PO Code</td>
	   <td align="center">STO No</td>
	   <td align="center">DO No</td>
	   <td align="center">GRN Code</td>
	   <td align="center">Stock Code</td>
	   <td align="center">Stock Desc.</td>
	   <td align="center">Qty Received</td>
	   <td align="center">Unit Price</td>
	   <td align="center"><input type="checkbox" name="select-all" id="select-all"/> Check/Uncheck All</td>
	   <td align="center">Amount Invoice</td>
	   </tr>
   <%ELSE%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">PO Code</td>
	   <td align="center">STO No</td>
	   <td align="center">DO No</td>
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
	  While Not rsGRN.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="ringht" valign="top"><%=NO%></td>
			<%IF SudahClose="Y" then%>
				<!-- Validate Period Control In Here -->
					<td align="left" title="POCODE" valign="top"><%=rsGRN("pocode")%></td>
					<td align="left" title="STONO" valign="top"><%=rsGRN("stono")%></td>
					<td align="left" title="DO NO" valign="top"><%=rsGRN("do_no")%></td>
					<td align="left" title="GRN CODE" valign="top"><%=rsGRN("receivenotecode")%></td>
					<td align="left" title="ITEMCODE" valign="top"><%=rsGRN("ITEMCODE")%></td>
					<td align="left" title="ITEMNAME" valign="top"><%=rsGRN("stockname")%></td>
					<td align="right" title="QUANTITY" valign="top"><%=FbpNum2(rsGRN("Quantity"))%></td>
					<td align="right" title="AMOUNT" valign="top"><%=FormatNumber(rsGRN("Amount"),2)%></td>
			<%End if%>

			<%if SudahClose<>"Y" Then%>
                    <td align="left" title="POCODE" valign="top"><%=rsGRN("POCODE")%></td>
                    <td align="left" title="STONO" valign="top"><%=rsGRN("STONO")%></td>
			        <td align="left" title="DO NO" valign="top"><%=rsGRN("DO_NO")%></td>
			        <td align="left" title="GRN CODE" valign="top"><%=rsGRN("RECEIVENOTECODE")%></td>
			        <td align="left" title="ITEMCODE" valign="top"><%=rsGRN("ITEMCODE")%></td>
			        <td align="left" title="ITEMNAME" valign="top"><%=rsGRN("STOCKNAME")%></td>
			        <td align="right" title="QUANTITY" valign="top"><%=FbpNum2(rsGRN("Quantity"))%></td>
			        <td align="right" title="UNITPRICE" valign="top"><%=FormatNumber(cdbl(to_nol(rsGRN("unitprice"))),2)%></td>
                    <td align="center">
						<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
					</td>
					<td align="center">
						<input style="text-align:right;background-color:#eee;" type="text" name="Amount_<%=no%>" class="amount" value = "<%=FormatNumber(rsGRN("amount"),2)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
					</td>
					<input type="hidden" name="receivenotecode_<%=no%>" value="<%=rsGRN("receivenotecode")%>">
					<input type="hidden" name="pocode_<%=no%>" value="<%=rsGRN("pocode")%>">
					<input type="hidden" name="stono_<%=no%>" value="<%=rsGRN("stono")%>">
					<input type="hidden" name="do_no_<%=no%>" value="<%=rsGRN("do_no")%>">
					<input type="hidden" name="itemcode_<%=no%>" value="<%=rsGRN("itemcode")%>">
					<input type="hidden" name="quantity_<%=no%>" value="<%=rsGRN("quantity")%>">
					<input type="hidden" name="unitprice_<%=no%>" value="<%=rsGRN("unitprice")%>">
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
	  <tr><td colspan="11" align="left">
		<!-- diambil dari form yang diatas -->
	    <input type="hidden" name="fcCode" value="<%=fcCode%>">
		<input type="hidden" name="refinvoice" value="<%=refinvoice%>">
		<input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		<input type="hidden" name="pocode" value="<%=pocode%>">
		<input type="hidden" name="No" value="<%=No%>">
		<input type="hidden" name="stono" value="<%=stono%>">  
		<input type="hidden" name="transportercode" value="<%=transportercode%>">  
		<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		<input type="hidden" name="fcba" value="<%=fcba%>">
		<input type="hidden" name="invoice_by" value="<%=invoice_by%>">
		<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		<input type="hidden" name="selectedpocode" value="<%=selectedpocode%>">
		<input type="hidden" name="selectedinvoice" value="<%=selectedinvoice%>">
		<input type="hidden" name="rowstate" value="<%=rowstate%>">		
		<input type="hidden" name="downpayment" value="<%=downpayment%>">
		<input type="hidden" name="amount" value="<%=amount%>">
		<input type="hidden" name="pph" value="<%=pph%>">
		<input type="hidden" name="amountpph" value="<%=amountpph%>">
		<input type="hidden" name="totalamount" value="<%=totalamount%>">
		
		<input type="hidden" name="no_faktur" value="<%=no_faktur%>">
		<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
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
<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
}

function getselectedpocode() {
	var spocode = document.getElementById("selectedpocode");
	spocode.value = document.getElementById("pocode").value;
}

function getselectedstono() {
	var spocode = document.getElementById("selectedstono");
	spocode.value = document.getElementById("stono").value;
}

function getselectedinvoice() {
	var sfcba = document.getElementById("selectedinvoice");
	sfcba.value = document.getElementById("invoice_by").value;
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
	getliststono();
	getselectedfcba();
	getselectedinvoice()
	getselectedstono();
	getstoinfo();
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
			txt.readOnly = false;
			txt.style.backgroundColor = "#fff";
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
			$('.amount').css("background-color", "#fff");
			$('.amount').prop('readOnly', false);
		}else{
			$('.amount').css("background-color", "#eee");
			$('.amount').prop('readOnly', true);
		}
	})
</script>
</body>
</html>
