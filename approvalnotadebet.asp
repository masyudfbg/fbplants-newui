<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Approval Nota Debet</title>
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
	var selectedvoucher = document.getElementById("selectedvoucher").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var transdate = document.getElementById("transdate").value;
	var PostStr = 'fcbasource='+selectedfcba+'&transdate='+transdate+'&vouchersource='+selectedvoucher;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getVoucherSourcebasedonfcba.asp', 'POST', PostStr, showlistvoucher);
}

// The function for handling the response from the server
var showlistvoucher = function (oXML) { 
    // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("voucherlist").innerHTML = "";
	document.getElementById("voucherlist").innerHTML = response;

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
	document.getElementById("pocode").innerHTML = "";
	document.getElementById("pocode").innerHTML = response;
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
    var s = JSON.parse(response);
    //alert (s.postingdate);
    
    // update the Div to show the result from the server
	document.getElementById("suppliercode").value = s.suppliercode;
	document.getElementById("suppliername").value = s.suppliername;
	document.getElementById("vat").value = s.ppn;
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
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks,txtNoPO,v_CrossCekPO
 Dim recordTotalAmount, rsrecordTotalAmount, varTotalAmount, updateNetBalance
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""


 Dim Jobcode, LocationType, LocationCode, Reference, Remark

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 if request("fccOde")="" then
    fcCode="" '----- Nantinya Ambil dari Format Master
 else
    fcCode = request("fccOde")
 end if
 
 ' fcCode = request("fccOde")
 ' If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("transdate")<>"" then
   transdate   =Request("transdate")
 else
   transdate   =textDate(date)
 end if

If "" & txtYear = "" Then txtYear = Year(transdate)
If "" & txtMonth = "" Then txtMonth = Month(transdate)

txttdate = textDate(txtMonth&"/01/"&txtYear)

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

 chargejob =Request("chargejob")
 chargeType =Request("chargeType")
 chargeCode =Request("chargeCode")
 vouchercode =Request("vouchercode")
 fcbasource =Request("fcbasource")
 vouchersource =Request("vouchersource")
 selectedvoucher =Request("selectedvoucher")
 
pocode = request("pocode")
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
selectedpocode = Request("selectedpocode")

vat = Request("vat")
amountvat = Request("amountvat")
totalamount = Request("totalamount")
refinvoice = Request("refinvoice")
no_faktur = Request("no_faktur")

 if Request("SourceType")="" then
    SourceType="BA" 'Default if Entry is Blank
 else
    SourceType=Request("SourceType")
 End if

Sub InsertDataPaymentVoucherDetail
	For n=1 to (cdbl(request("no"))-1) 
		if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
			qMaxLineNo =	" select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
							" from paymentvoucherdetail " & _
							" where " & _
							" fcba = '" & Session("fcBA") & "'" 
			
			Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
			rsqMaxLineNo.CursorType = 3
			rsqMaxLineNo.CursorLocation = 3
			rsqMaxLineNo.Open qMaxLineNo, connstr

			if not rsqMaxLineNo.eof then
				maxLineNo = cdbl(to_nol(rsqMaxLineNo("maxLineNo")))
			else
				maxLineNo = 0
			end if
			rsqMaxLineNo.close
			set rsqMaxLineNo = Nothing
			
			maxLineNo = maxLineNo + 1

		qq = 	" INSERT INTO paymentvoucherdetail ( vouchercode" & _
				"                                    , sourcetype" & _
				"                                    , locationtype" & _
				"                                    , locationcode" & _
				"                                    , jobcode" & _
				"                                    , chargetype" & _
				"                                    , chargecode" & _
				"                                    , chargejob" & _
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
				"                                    , lasttime) " & _
				" VALUES ( '" & vouchercode & "'" & _
				"          , '" & SourceType & "'" & _
				"          , '" & request("chargetype_"&TRIM(n)) & "'" & _ 
				"          , '" & request("chargecode_"&TRIM(n)) &  "'" & _
				"          , '" & request("chargejob_"&TRIM(n)) & "'" & _
				"          , '" & chargetype & "'" & _ 
				"          , '" & chargecode &  "'" & _
				"          , '" & chargejob & "'" & _
				"          , REPLACE('" & request("amount_"&TRIM(n)) & "', ',', '')" & _
				"          , '" & fcbasource & "'" & _ 
				"          , '" & Remarks & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
				"          , '" & Session("fcBA") & "'" & _
				"          , '" & InvoiceCode & "'" & _
				"          , '" & AGREEMENTCODE & "'"& _
				"          , '" & AGREEMENTLINENO & "'" & _
				"          , TO_DATE('" & WIPDATE & "', 'mm/dd/yyyy')" & _
				"          , '" & maxLineNo & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE,'HH24:MI') )"
		' response.write qq
		connStr.execute(qq)
		End if
	Next
End Sub

Sub InsertDataPaymentVoucherDetail2
	qMaxLineNo =  " SELECT fccode" & _
				  "  FROM intraco_tab" & _
				  "	 WHERE fcba = '"& fcbasource &"'" 
				
	Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
	rsqMaxLineNo.CursorType = 3
	rsqMaxLineNo.CursorLocation = 3
	rsqMaxLineNo.Open qMaxLineNo, connstr

	if not rsqMaxLineNo.eof then
		jobintraco = cdbl(rsqMaxLineNo("fccode"))
	else
		jobintraco = 0
	end if
	rsqMaxLineNo.close
	set rsqMaxLineNo = Nothing

	rjobintraco = jobintraco

	qMaxLineNo =	" select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
					" from paymentvoucherdetail " & _
					" where " & _
					" fcba = '" & Session("fcBA") & "'" 
	' response.write qMaxLineNo
	Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
	rsqMaxLineNo.CursorType = 3
	rsqMaxLineNo.CursorLocation = 3
	rsqMaxLineNo.Open qMaxLineNo, connstr

	if not rsqMaxLineNo.eof then
		maxLineNo = cdbl(to_nol(rsqMaxLineNo("maxLineNo")))
	else
		maxLineNo = 0
	end if
	rsqMaxLineNo.close
	set rsqMaxLineNo = Nothing
	
	maxLineNo = maxLineNo + 1


	qMaxLineNo =  " SELECT SUM(NVL(dtl.amount *-1,0)) AS amount" & _
				  " FROM approvalnotadebetdtl dtl" & _
				  " WHERE dtl.fcba = '" & Session("fcba") & "'" & _
				  " AND vouchercode = '"& vouchercode &"'" 
				
	Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
	rsqMaxLineNo.CursorType = 3
	rsqMaxLineNo.CursorLocation = 3
	rsqMaxLineNo.Open qMaxLineNo, connstr

	if not rsqMaxLineNo.eof then
		amount = cdbl(rsqMaxLineNo("amount"))
	else
		amount = 0
	end if
	rsqMaxLineNo.close
	set rsqMaxLineNo = Nothing

	ramount = amount

qq = 	" INSERT INTO paymentvoucherdetail ( vouchercode" & _
		"                                    , sourcetype" & _
		"                                    , locationtype" & _
		"                                    , locationcode" & _
		"                                    , jobcode" & _
		"                                    , chargetype" & _
		"                                    , chargecode" & _
		"                                    , chargejob" & _
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
		"                                    , lasttime) " & _
		" VALUES ( '" & vouchercode & "'" & _
		"          , '" & SourceType & "'" & _
		"          , 'CB'" & _ 
		"          , 'BS'" & _
		"          , '" & rjobintraco & "'" & _
		"          , '" & chargetype & "'" & _ 
		"          , '" & chargecode &  "'" & _
		"          , '" & chargejob & "'" & _
		"          , REPLACE('" & ramount & "', ',', '')" & _
		"          , '" & fcbasource & "'" & _ 
		"          , '" & Remarks & "'" & _
		"          , '" & Session("CodeName") & "'" & _
		"          , '" & Session("CodeName") & "'" & _
		"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
		"          , '" & Session("fcBA") & "'" & _
		"          , '" & InvoiceCode & "'" & _
		"          , '" & AGREEMENTCODE & "'"& _
		"          , '" & AGREEMENTLINENO & "'" & _
		"          , TO_DATE('" & WIPDATE & "', 'mm/dd/yyyy')" & _
		"          , '" & maxLineNo & "'" & _
		"          , SYSDATE" & _
		"          , TO_CHAR(SYSDATE,'HH24:MI') )"

connStr.execute(qq)
End Sub

Sub DeleteDataPaymentVoucher
		q = " DELETE FROM paymentvoucherdetail" & _
		    " WHERE vouchercode = '" & vouchercode & "'" & _
		    "       and fcBA='" & Session("fcBA") & "'"
		' response.write q
		connStr.execute(q)
End Sub
sub post_amount
	recordTotalAmount = " SELECT NVL(SUM(amount),0) totalAmount" & _
						" FROM invoicedetail" & _
						" WHERE invoicecode = '" & fcCode & "'" & _ 
						"       AND fcba = '" & fcba & "'"

	Set rsrecordTotalAmount = Server.CreateObject("ADODB.Recordset")
	rsrecordTotalAmount.Open recordTotalAmount, connstr

	if Not rsrecordTotalAmount.Eof then
		allamount = cdbl(to_nol(rsrecordTotalAmount("totalAmount")))
	Else
		allamount = 0
	End if
	
	amountvat = allamount *cdbl(to_nol(vat))/100
	totalamount = allamount +cdbl(to_nol(amountvat))


	q = " UPDATE invoice" & _
		" SET amount = '" & allamount & "'" & _
		"     , amountvat = '" & amountvat & "'" & _
		"     , totalamount = '" & totalamount & "'" & _
		"     , lastupdate = SYSDATE " & _
		"     , lasttime = TO_CHAR(SYSDATE,'HH24:MI') "& _	
        "     , fcedit = '" & session("CodeName") & "' " & _		
		" WHERE invoicecode = '" & fcCode & "'" & _
		"       AND fcba = '" & fcba & "'"

	connStr.execute(q)

end sub
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	
	REMARKS=request("REMARKS")
	
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
			arr_dtl_eof = true
			q = " SELECT * FROM approvalnotadebet WHERE vouchercode = '" & vouchercode & "' AND fcba = '" & session("fcba") & "'"
			' response.write q
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open q, connstr
			if not rsPO.eof then
				arr_dtl = rsPO.getrows()
				arr_dtl_eof = false
			end if
			rsPO.close : set rsPO = nothing
			
			if not arr_dtl_eof then
				for xyz=0 to ubound(arr_dtl,2)
					q = " DELETE FROM approvalnotadebet" & _
					    " WHERE vouchercode = '" & vouchercode & "'" & _
					    "       AND fcbasource = '" & Request("fcbasource") & "'" & _
					    "       AND vouchersource = '" & vouchersource & "'" & _
					    "       AND transdate = to_date('" & transdate & "','mm/dd/yyyy')" & _
					    "       and fcBA='" & Session("fcBA") & "'"
					' response.write q
					connStr.execute(q)

					q = " DELETE FROM approvalnotadebetdtl" & _
					    " WHERE vouchercode = '" & vouchercode & "'" & _
					    "       and fcBA='" & Session("fcBA") & "'"
					' response.write q
					connStr.execute(q) 
					
					call DeleteDataPaymentVoucher

					' recordTotalAmount = " select count(vouchercode) total from paymentvoucherdetail" & _
					' 					"	where fcba = '"& session("fcba") &"'" & _
					' 					"	and vouchercode = '"& vouchercode &"'"
					' response.write recordTotalAmount
					' Set rsrecordTotalAmount = Server.CreateObject("ADODB.Recordset")
					' rsrecordTotalAmount.Open recordTotalAmount, connstr

					' if Not rsrecordTotalAmount.Eof then
					' 	total = cdbl(to_nol(rsrecordTotalAmount("total")))
					' End if

					' If Len(Trim(total)) = 0 Then
						q = " DELETE FROM paymentvoucher" & _
					    	" WHERE vouchercode = '" & vouchercode & "'" & _
					    	"       AND datecreated = to_date('" & transdate & "','mm/dd/yyyy')" & _
					    	"       and fcBA='" & Session("fcBA") & "'"
							' response.write q
							connStr.execute(q)
					' Else

					' End If
					' q = " DELETE FROM invoicedetail" & _
					' 	" WHERE invoicecode = '" & fcCode & "'" & _
					' 	"       AND receivenotecode = '" & arr_dtl(0,xyz) & "'" & _
					' 	"       AND pocode = '" & arr_dtl(1,xyz) & "'" & _
					' 	"       AND purchaselinenokey = '" & arr_dtl(2,xyz) & "'" & _
					' 	"       AND fcBA = '" & fcba & "'"
					' 'response.write q
					' connStr.execute(q)
					
					' q = " UPDATE receivenotedetail SET amountinvoice = 0" & _
					' 	" WHERE receivenotecode = '" & arr_dtl(0,xyz) & "'" & _
					' 	"       AND purchaselinenokey = '" & arr_dtl(2,xyz) & "'" & _
					' 	"       AND fcBA = '" & fcba & "'"
					' 'response.write q
					' connStr.execute(q)
				next
			end if
			erase arr_dtl : arr_dtl_eof = true

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

	elseif Request("txtJenis")= "Details" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, transdate)
		If Cek_IsDailyClosing = "FALSE" then

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
		Cek_IsDailyClosing = IsDailyClosing(fcba, transdate)
		If Cek_IsDailyClosing = "FALSE" then
			'--Add by Ukie to handle Auto Number of PR
				' If Ucase(fcCode) = Ucase("(NEW)") Then
				' 	oSql = " select nvl(controlnumber_api.get_nextnumber(upper('INVOICE')" & _
				' 		   "                                         , " & year(textDate(DATECREATED)) & "" & _
				' 		   "                                         , " & month(textDate(DATECREATED)) & "" & _
				' 		   "                                         , '" & fcba & "'" & _
				' 		   " ), 'FALSE') INVOICE from dual"
					
				' 	' Response.Write oSql
						   
				' 	Set oRs = Server.CreateObject("ADODB.Recordset")
				' 	oRs.Open oSql, connstr
				' 	If Not oRs.EOF Then
				' 		txtINVOICE = oRs("INVOICE")
				' 	End If
				' 	oRs.Close
				' 	Set oRs = Nothing
				' 		fcCode = txtINVOICE
				' End If
				
				' If txtINVOICE = "FALSE" Then
				' 	err.Number = 1001
				' 	err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
				' End If
				
				'Response.Write PRCODE & txtPRCode
			'--End Added
			
			' Dim rsPO
			' Set rsPO = Server.CreateObject("ADODB.Recordset")
			' rsPO.Open 	" Select INVOICECODE, suppliercode from INVOICE where INVOICECODE = '" & fcCode & "'" & _
			' 			" and fcBA = '" & fcba & "'", connstr
			

			' IF rsPO.eof then	    
			' 	Dim txtCodeSupplier,txtCodePO
			' 	q = " Select distinct POCODE,SUPPLIERCODE From RECEIVENOTE where pocode = '" & pocode & "'"  & _
			' 		" and fcBA = '" & fcba & "'"
			' 	'response.write q
			' 	Dim rsLPO  
			' 	set rsLPO = Server.CreateObject("ADODB.Recordset")
			' 	rsLPO.Open q, connstr
			' 	if not rsLPO.eof then
			' 		txtCodePO=rsLPO("POCODE")
			' 		txtCodeSupplier=rsLPO("SUPPLIERCODE")			
			' 	End if 
			
			' 	rsLPO.Close
			' 	set rsLPO = Nothing
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select * from approvalnotadebet where vouchercode = '" & vouchercode & "' and fcba ='" & session("fcba") & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then

				q = " Insert into approvalnotadebet (VoucherCode, transdate, fcbasource, vouchersource, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
					" values ('" & vouchercode & "',to_date('" & transdate & "','mm/dd/yyyy'),'" & fcbasource & "','" & vouchersource & "','" & _
					Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & session("fcba") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
				' response.write q
				connStr.execute(q)
			End if
			' rsPO.Close	
			' Dim n
			'response.Write "No " & (cdbl(request("no"))-1) & " amount : " &  request("AmountInvoice_"&TRIM(n))
			For n=1 to (cdbl(request("no"))-1) 
				if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
						q = " Insert into  approvalnotadebetdtl (vouchercode, voucherlineno, amount, jobcode, targetcode, targettype, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
							" values ('" & vouchercode & "'" & _
							" ,'" & request("lineno_"&TRIM(n)) & "'" & _
							" , REPLACE('" & request("amount_"&TRIM(n)) & "',',', '')" & _
							" ,'" & request("chargejob_"&TRIM(n)) & "'" & _
							" ,'" & request("chargecode_"&TRIM(n)) & "','" & request("chargetype_"&TRIM(n)) & "'" & _
							" ,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						' response.write q
						connStr.execute(q)
						' q = " Update ReceiveNoteDetail  set AMOUNTINVOICE=AMOUNTINVOICE+REPLACE('" & request("AmountInvoice_"&TRIM(n)) & "', ',', '')"  & _
						' 	", LASTUPDATE = sysdate " & _
						' 	", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						' 	" where RECEIVENOTECODE='" & request("receivenotecode_"&TRIM(n)) & "' and purchaselinenokey = '" & request("linenokey_"&TRIM(n)) & "'" & _
						' 	" and fcBA='" & fcba & "'"
						' 'response.Write(q)
						' connStr.execute(q)
						
				End if	 
			Next
			
			' Call post_amount
						qMaxLineNo = " select fccode from bank" & _ 
									 "	where banktype= 'JOURNAL'" & _
									 " and upper(trigger_system ) = upper('Bank')" & _
									 " and fcba = '"& Session ("fcba") &"'"
									
						Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
						rsqMaxLineNo.CursorType = 3
						rsqMaxLineNo.CursorLocation = 3
						rsqMaxLineNo.Open qMaxLineNo, connstr

						if not rsqMaxLineNo.eof then
							bankcode = cdbl(rsqMaxLineNo("fccode"))
						else
							bankcode = 0
						end if
						rsqMaxLineNo.close
						set rsqMaxLineNo = Nothing

						rbankcode = bankcode

						qMaxLineNo = "select accountingyear, accountingperiod from  ACCOUNTINGYEARPERIOD where fcBA='" &  Session("fcBA") & "'  and to_char(valid_until ,'YYYY') = '"& Year(Session("v_Costbook")) &"'  and to_number(to_char(valid_until ,'MM')) > to_number('"& Month(Session("v_Costbook")) &"') -1  order by accountingyear||lpad(accountingperiod,2,0)"

						Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
						rsqMaxLineNo.CursorType = 3
						rsqMaxLineNo.CursorLocation = 3
						rsqMaxLineNo.Open qMaxLineNo, connstr

						if not rsqMaxLineNo.eof then
							accountingyear = cdbl(rsqMaxLineNo("accountingyear"))
							accountingperiod = cdbl(rsqMaxLineNo("accountingperiod"))
						else
							accountingyear = 0
							accountingperiod = 0
						end if
						rsqMaxLineNo.close
						set rsqMaxLineNo = Nothing

						raccountingyear = accountingyear
						raccountingperiod = accountingperiod


						q = " SELECT NVL(SUM(amount), 0) totalamount" & _
							" FROM approvalnotadebetdtl" & _
							" WHERE VoucherCode='" & vouchercode & "'" & _
							"       AND fcBA = '" & session("fcBA") & "'"

						set rsVoucher = server.CreateObject("Adodb.recordset")
						rsVoucher.Open q, connstr
						if Not rsVoucher.Eof then
							TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
						End if

						q	=	" INSERT INTO paymentvoucher ( vouchercode" & _
								"                              , referenceno" & _
								"                              , bankcode" & _
								"                              , datecreated" & _
								"                              , accountingyear" & _
								"                              , accountingperiod" & _
								"                              , totalamount" & _
								"                              , suppcontcode" & _
								"                              , remarks" & _
								"							   , rowstateid	" & _
								"							   , automaticapproved	" & _
								"							   , transtype	" & _
								"                              , vouchertype" & _
								"                              , fcentry" & _
								"                              , fcedit" & _
								"                              , fcip" & _
								"                              , fcba" & _
								"                              , lastupdate" & _
								"                              , lasttime)" & _
								" VALUES ( '" & vouchercode & "'" & _
								"          , '" & referenceno & "'" & _
								"          , '" & rbankcode & "'" & _
								"          , TO_DATE('" & transdate & "','mm/dd/yyyy')" & _
								"          ," & txtyear & " " &_
								"          ,TO_NUMBER(" & txtMonth & ")" &_
								"          , '"& TOTALAMOUNT &"'" & _
								"          , ''" & _
								"          , '" &  request("HeaderRemarks") & "'" & _
								"          , 'Approved'" & _
								"          , '1'" & _
								"          ,'JOURNAL'" & _
								"          , ''" & _
								"          , '" & Session("CodeName") & "'" & _
								"          , '" & Session("CodeName") & "'" & _
								"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
								"          , '" & Session("fcBA") & "'" & _
								"          , SYSDATE" & _
								"          , TO_CHAR(SYSDATE,'HH24:MI') )"
								' response.Write("q " & q )
								connStr.execute(q) 

						call InsertDataPaymentVoucherDetail
						call InsertDataPaymentVoucherDetail2
			
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

 Dim rsVoucher,SudahClose
 q = " select * " & _
     " from approvalnotadebet " & _
     " where vouchercode='" & vouchercode & "' and fcba = '"& Session("fcBA") &"'"
 ' response.write q
 set rsVoucher = server.CreateObject("Adodb.recordset")
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    vouchercode = rsVoucher("vouchercode")	
    transdate = rsVoucher("transdate")
	fcbasource = rsVoucher("fcbasource")
	selectedfcba = rsVoucher("fcbasource")
	vouchersource = rsVoucher("vouchersource")
	selectedvoucher = rsVoucher("vouchersource")
	' selectedpocode = rsVoucher("pocode")
	
	SudahClose = "Y"
	
	' vat = rsVoucher("vat")
	' amountvat = rsVoucher("amountvat")
	' totalamount = rsVoucher("totalamount")
	
	' no_faktur = rsVoucher("no_faktur")
	' fakturdate = rsVoucher("fakturdate")
 End if
 rsVoucher.Close : Set rsVoucher = Nothing
 

 if err.number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
 End if
 '-------------ISI Payment Form
 
 Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Bank')"
	  ' response.write oSQLQuery
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
 <form name="frmEntry" method="post" action="approvalnotadebet.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Voucher No</td>
			<td align="left"><input type="text" maxlength="50" name="vouchercode" size="30" value="<%=vouchercode%>" autocomplete="off"><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" id="transdate" name="transdate" size="10" value="<%=transdate%>">
			<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Fcba Source
				<td align="left">
					<% 	Dim rsCropA
						q =  " SELECT DISTINCT bu.fccode fccode, bu.fcname fcname FROM user_company_tab uft INNER JOIN businessunit bu ON bu.fccompanycode = uft.fccompanycode " & _
							 " WHERE fcusercode = '" & Session("CodeName") & "' AND isdefault = 'TRUE' ORDER BY fccode"
						Set rsCropA = Server.CreateObject("ADODB.Recordset")
						rsCropA.CursorType = 3
						rsCropA.CursorLocation = 3
						rsCropA.Open q, connstr
						if not rsCropA.eof then%>
						<select id="fcbasource" name="fcbasource" class="chosen-select" style="width:30%;" onchange="javascript:getselectedfcba();javascript:getlistvoucher();">         
						<option value=""></option> 
						<%
						Do While Not rsCropA.eoF%>
						<%IF cstr(""&fcbasource) =rsCropA("fccOde") Then%>
							<option value="<%=rsCropA("fccOde")%>" selected><%=rsCropA("fccOde")%> : <%=rsCropA("fcname")%></option> 
						<%Else%>
							<option value="<%=rsCropA("fccOde")%>"><%=rsCropA("fccOde")%> : <%=rsCropA("fcname")%></option> 
						<%End if%>
						<%rsCropA.movenext
						loop
						End if
						rsCropA.Close
						set rsCropA = Nothing
						q=""
						%>
				</select>
				</td>
				<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Voucher Source
				<td align="left">
				 	<div id="voucherlist"></div>
				</td>
					<input type="Hidden" id="selectedvoucher" name="selectedvoucher" value="<%=selectedvoucher%>">
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="javascript:launchwin('Printapprovalnotadebet.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" Then%>
					<a href="approvalnotadebet.asp?vouchercode=<%=vouchercode%>&fcbasource=<%=fcbasource%>&transdate=<%=transdate%>&vouchersource=<%=vouchersource%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" border="0"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End if%>
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
<%if cmdView = "True" AND LEN(TRIM(vouchercode)) >  0 Then
   Dim rsGRN,intRecordsPerPage,cColor
   if SudahClose="Y" Then
      q = 	" SELECT distinct pv.vouchercode, pvd.chargejob, pvd.chargetype, pvd.chargecode, pvd.lineno, pvd.amount" & _
			"  FROM paymentvoucher pv" & _
			"       INNER JOIN paymentvoucherdetail pvd" & _
			"          ON (pv.vouchercode = pvd.vouchercode AND pv.fcba = pvd.fcba)" & _
			"       INNER JOIN intraco_tab it" & _
			"          ON (it.fccode = pvd.jobcode)" & _
			" WHERE     pv.fcba = '"& fcbasource &"' " & _
			"              AND it.fcba = '"& Session("fcba") &"' " & _
			"              AND pv.vouchercode = '"& vouchersource &"'" 
   Else	 
      q = 	" SELECT distinct pv.vouchercode, pvd.chargejob, pvd.chargetype, pvd.chargecode, pvd.lineno , pvd.amount" & _
			"  FROM paymentvoucher pv" & _
			"       INNER JOIN paymentvoucherdetail pvd" & _
			"          ON (pv.vouchercode = pvd.vouchercode AND pv.fcba = pvd.fcba)" & _
			"       INNER JOIN intraco_tab it" & _
			"          ON (it.fccode = pvd.jobcode)" & _
			" WHERE     pv.fcba = '"& fcbasource &"' " & _
			"              AND it.fcba = '"& Session("fcba") &"' " & _
			"              AND pv.vouchercode = '"& vouchersource &"'" 

   End if
   ' response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.Open q, connstr%>
   <form action="approvalnotadebet.asp" method="post" name="eForm">
   <table class="data" border="1" width="100%">
   <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		   <td align="center">No</td>
		   <td align="center">Voucher Line No</td>
		   <td align="center">Jobcode</td>
		   <td align="center">Target Type</td>
		   <td align="center">Taget Code</td>
		   <td align="center"><input type="checkbox" name="select-all" id="select-all"/> Check/Uncheck All</td>
		   <td align="center">Amount </td>
	   </tr>
   <%ELSE%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		   <td align="center">No</td>
		   <td align="center">Voucher Line No</td>
		   <td align="center">Jobcode</td>
		   <td align="center">Target Type</td>
		   <td align="center">Taget Code</td>
		   <td align="center">Amount </td>
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
                    <td align="left" title="" valign="top"><%=rsGRN("lineno")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("chargejob")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("chargetype")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("chargecode")%></td>
			        <td align="center">
						<input style="text-align:right;background-color:#eee;" type="text" name="Amount_<%=no%>" class="amount" value = "<%=FormatNumber(rsGRN("amount"),0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
					</td>
			<%End if%>

			<%if SudahClose<>"Y" Then%>
                    <td align="center" title="" valign="top"><%=rsGRN("lineno")%></td>
			        <td align="center" title="" valign="top"><%=rsGRN("chargejob")%></td>
			        <td align="center" title="" valign="top"><%=rsGRN("chargetype")%></td>
			        <td align="center" title="" valign="top"><%=rsGRN("chargecode")%></td>
		            <td align="center">
						<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
					</td>
					<td align="center">
						<input style="text-align:right;background-color:#eee;" type="text" name="Amount_<%=no%>" class="amount" value = "<%=FormatNumber(rsGRN("amount"),0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
					</td>
					<input type="hidden" name="lineno_<%=no%>" value="<%=rsGRN("lineno")%>">
					<input type="hidden" name="chargejob_<%=no%>" value="<%=rsGRN("chargejob")%>">
					<input type="hidden" name="chargecode_<%=no%>" value="<%=rsGRN("chargecode")%>">
					<input type="hidden" name="chargetype_<%=no%>" value="<%=rsGRN("chargetype")%>">

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
	  <tr><td colspan="9" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
	      <input type="hidden" name="refinvoice" value="<%=refinvoice%>">
		  <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		  <input type="hidden" name="pocode" value="<%=pocode%>">
		  <input type="hidden" name="No" value="<%=No%>">
		  
		<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		<input type="hidden" name="fcba" value="<%=fcba%>">
		<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		<input type="hidden" name="selectedpocode" value="<%=selectedpocode%>">
		
		<input type="hidden" name="vat" value="<%=vat%>">
		<input type="hidden" name="amountvat" value="<%=amountvat%>">
		<input type="hidden" name="totalamount" value="<%=totalamount%>">
		
		<input type="hidden" name="no_faktur" value="<%=no_faktur%>">
		<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
		  
  		<input Type="Hidden" Value="<%=request("vouchercode")%>" Name="vouchercode">
		<input Type="Hidden" Value="<%=request("vouchersource")%>" Name="vouchersource">
		<input Type="Hidden" Value="<%=request("fcbasource")%>" Name="fcbasource">
		<input Type="Hidden" Value="<%=request("transdate")%>" Name="transdate">

		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Submit">
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
	sfcba.value = document.getElementById("fcbasource").value;
}

function getselectedvoucher() {
	var spocode = document.getElementById("selectedvoucher");
	spocode.value = document.getElementById("vouchersource").value;
	// alert("vouchersource");
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
	getlistvoucher();
	getselectedfcba();
    getselectedvoucher();
	// getpoinfo();
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
		var txt = chk.parentNode.parentNode.cells[6].getElementsByTagName('input')[0];
		
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
<script type="text/javascript">
	

</script>
</body>
</html>
