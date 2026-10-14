<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Invoice Purchase Order</title>
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
var getSupp = function () {
	// build up the post string when passing variables to the server side page
	var fccode = document.getElementById("fccode").value;
	var businessunit = document.getElementById("businessunit").value;
	var PostStr = 'fccode='+fccode+'&businessunit='+businessunit;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getDetailSPO.asp', 'POST', PostStr, showMessageResponse);
	
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.supplieraddress);
    // update the Div to show the result from the server
    // document.getElementById("suppliername").value = s.suppliername;
	document.getElementById("suppliername").value = s.suppliername;
	// document.getElementById("address").innerHTML = document.getElementById("supplieraddress").value;
	// alert(s.suppliername);
    
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

var getSPODetailsInfo = function () {
	// build up the post string when passing variables to the server side page
	var fccode = document.getElementById("fccode").value;
	var businessunit = document.getElementById("businessunit").value;
	var spolineno = document.getElementById("spolineno").value;
	var PostStr = 'fccode='+fccode+'&businessunit='+businessunit+'&spolineno='+spolineno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getSPODetailsInfo.asp', 'POST', PostStr, showSPODetailsInfo);
	
}

// The function for handling the response from the server
var showSPODetailsInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // update the Div to show the result from the server
	document.getElementById("spo_amount").value = s.amount;
	document.getElementById("spo_charge").value = s.othercharge;
	document.getElementById("spo_totalamount").value = s.amountcharge;
	
	if ( document.getElementById("amount").value == 0 || document.getElementById("amount").value == null) {
		document.getElementById("amount").value = s.amount;
	}
	
	if ( document.getElementById("othercost").value == 0 || document.getElementById("othercost").value == null) {
		document.getElementById("othercost").value = s.othercharge;
	}
	
	document.getElementById("descriptionx").value = s.descriptionx;
    
};

</script>
<script language="JavaScript">
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

function getPO(str){

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
	xmlhttp.open("GET","ajax/getSPO.asp?businessunit="+str+"",true);
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

  if Request("fakturdate")<>"" then
   fakturdate   =Request("fakturdate")
 else
   fakturdate   =textDate(date)
 end if

 if Request("transdate")<>"" then
   transdate   =Request("transdate")
 else
   transdate   =textDate(date)
 end if
 
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
 businessunit = Request("businessunit")
 purchaseby = Request("purchaseby")
 suppliername = Request("suppliername")
 supplieraddress = Request("supplieraddress")
 termofpayment = Request("termofpayment")
 SHIPPINGPOINT = Request("SHIPPINGPOINT")
 dp = Request("dp")
 ppn = Request("ppn")
 invoiceno = Request("invoiceno")
 StockName = Request("StockName")
 unitprice = Request("unitprice")
 disc = Request("disc")
 othercost = Request("othercost")
 totalamount = Request("totalamount")
 JobCode = Request("JobCode")
 targettype = Request("targettype")
 targetcode = Request("targetcode")
 grandtotal = Request("grandtotal")
 amountcharge = Request("amountcharge")
 totalppn = Request("totalppn")
 fakturno = request ("fakturno")
 descriptionx = request ("descriptionx")
 lineno = request ("lineno")
 spolineno = request ("spolineno")
 
private function cek_trigger(tt_, tc_, jobcode_, fcba_, date_, sourcetype_)
	result_ = ""
	q = " SELECT f_cektarget_bydate ('"&tt_&"', '"&tc_&"', '"&jobcode_&"', '"&fcba_&"', TO_DATE('"&date_&"', 'mm/dd/yyyy'), '"&sourcetype_&"') outputstr FROM DUAL"
	set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
	rsTempPurchaseRequest.Open q, connstr
	' response.write q
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
		
				
		q="Delete from invoiceservicedtl where invoiceno='" & invoiceno & "' and fcBA='" & businessunit & "'"
		connStr.execute(q)

		q="Delete from invoiceservicehdr where invoiceno='" & invoiceno & "' and fcBA='" & businessunit & "'"
		connStr.execute(q)
			
		if err.number  = 0 then
			connStr.CommitTrans
		end if
	   cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans
	   q="Delete from invoiceservicedtl where invoiceno='" & invoiceno & "' and fcBA='" & businessunit & "' and lineno = '"&lineno&"'"
	   ' response.write q
	   connStr.execute(q)

		'Untuk menghhitung jumlah total di header
		dtl_totalamount = 0
		set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
		q="select *  from invoiceservicedtl " & _
		  " where invoiceno='" & invoiceno & "' and " & _
		  " fcBA='" & businessunit & "'"
		  ' response.write q
		rsTempPurchaseRequest.Open q, connstr
		if Not rsTempPurchaseRequest.Eof then
			while Not rsTempPurchaseRequest.Eof 
				dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("totalamount")))
				rsTempPurchaseRequest.movenext
			wend 						
		End if					
		rsTempPurchaseRequest.Close : set rsTempPurchaseRequest = Nothing

		totalppn = dtl_totalamount *cdbl(to_nol(ppn))/100
		grandtotal = totalppn +dtl_totalamount

	  q= " Update invoiceservicehdr set totalppn = replace('" & totalppn & "',',','')" & _
		 ", grandtotal = replace('" & grandtotal & "',',','')" & _
		 ", LASTUPDATE = sysdate " & _
		 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
		 " where  invoiceno='" & invoiceno & "' and fcBA='" & businessunit & "'"
		' response.write q
		connStr.execute(q)
	   if err.number  = 0 then
		  connStr.CommitTrans
	   end if
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans
		
		if ( len(trim(""&JOBCODE)) > 0 or len(trim(""&TARGETTYPE)) > 0 or len(trim(""&TARGETCODE)) > 0 )then 
		
			if UCase(cek_trigger(TARGETTYPE, TARGETCODE, JOBCODE, businessunit, transdate, "PR")) = UCase("Success") then
				max_amount = 0 : max_charge = 0 : max_amountcharge = 0
				strLineMax = 	" SELECT NVL(SUM(amount),0) amount, NVL(SUM(othercharge),0) othercharge, NVL(SUM(amountcharge),0) amountcharge" & _
								" FROM servicepurchaseorderdetails" & _
								" WHERE pocode = '"&fcCode&"'" & _
								"       AND fcba = '" & businessunit & "'" & _
								"       AND linenokey = '" & spolineno & "'"

				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					max_amount = cdbl(to_nol(rsstrLineMax("amount")))
					max_charge = cdbl(to_nol(rsstrLineMax("othercharge")))
					max_amountcharge = cdbl(to_nol(rsstrLineMax("amountcharge")))
				else
					max_amount = 0 : max_charge = 0 : max_amountcharge = 0
				end if							
				rsstrLineMax.Close : set rsstrLineMax = Nothing
			
				Dim rsPO, rsPurchases,rsQty
			  
				Set rsPO = Server.CreateObject("ADODB.Recordset")
				rsPO.Open "Select poCode from invoiceservicehdr where invoiceno='" & invoiceno & "' AND FCBA = '" & businessunit & "'" , connstr
				IF rsPO.eof then	   
					' Modified by NAF
					' To handle first time Input for Purchase Order

					q = " INSERT INTO invoiceservicehdr ( invoiceno, pocode, transdate, ppn, company, suppliername, no_faktur, fakturdate, totalppn, grandtotal, fcentry, fcedit, fcip, fcba,  lastupdate, lasttime ) " & _
					 " VALUES ( '" & invoiceno & "'" & _
					 "         , '" & fcCode & "'" & _
					 "         , TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
					 "         , '" & ppn & "'" & _
					 "         , '" & fccompanycode & "'" & _
					 "         , '" & suppliername & "'" & _
					 "         , '" & fakturno & "'" & _
					 "         , TO_DATE('" & fakturdate & "', 'mm/dd/yyyy')" & _					 
					 " 		   , replace('" & totalppn & "',',','')" & _
					 " 		   , replace('" & grandtotal & "',',','')" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					 "         , '" & businessunit & "'" & _
					 "         , SYSDATE" & _
					 "         , TO_CHAR(SYSDATE,'hh24:mi'))"
					connStr.execute(q)
					' response.write q
				End if
				rsPO.Close
				
				in_data_amount = 0 : in_data_charge = 0 : in_data_amountcharge = 0
				strLineMax = 	" SELECT NVL(SUM(amount),0) amount, NVL(SUM(othercost),0) othercharge, NVL(SUM(totalamount),0) amountcharge" & _
								" FROM invoiceservicedtl" & _
								" WHERE pocode = '"&fcCode&"'" & _
								"       AND fcba = '" & businessunit & "'" & _
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
				
				if in_data_amountcharge + cdbl(to_nol(replace(AMOUNT,",",""))) + cdbl(to_nol(replace(OTHERCOST,",",""))) <= max_amountcharge then

					tempMaxLineNo = 0
					strLineMax = " SELECT NVL(MAX(lineno),0) maxlineno FROM invoiceservicedtl WHERE fcBA='" & businessunit & "' AND invoiceno = '"&INVOICENO&"'"
					Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
					rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
					
					if not rsstrLineMax.eof then
						tempMaxLineNo = cdbl(to_nol(rsstrLineMax("maxlineno")))
					end if
								
					rsstrLineMax.Close : set rsstrLineMax = Nothing
							
					tempMaxLineNo = tempMaxLineNo +1
					
					TOTALAMOUNT = cdbl(to_nol(replace(AMOUNT,",",""))) + cdbl(to_nol(replace(OTHERCOST,",","")))
							
					q = " INSERT INTO invoiceservicedtl (invoiceno, jobcode, targettype, targetcode, description, amount, othercost, totalamount, fcentry, fcedit, fcip, fcba, lastupdate, lasttime, lineno, pocode, polinenokey) " & _
							 "values ('" & INVOICENO & "','" & JOBCODE & "','" & TARGETTYPE & "','"  &  TARGETCODE & "','"  &  descriptionx & "',replace('" & AMOUNT & "',',','')" & _
							 ",replace('" & OTHERCOST & "',',',''),replace('" & TOTALAMOUNT & "',',',''),'" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
							 Request.ServerVariables("REMOTE_ADDR") &"','" & businessunit & "', sysdate, To_Char(sysdate,'HH24:MI'), '"&tempMaxLineNo&"', '"&fcCode&"', '"&spolineno&"')"
					' response.write q
					 
					connStr.execute(q)
							
				' 	'Untuk menghhitung jumlah total di header
					dtl_totalamount = 0
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select *  from invoiceservicedtl " & _
					  " where invoiceno='" & invoiceno & "' and " & _
					  " fcBA='" & businessunit & "'"
					  ' response.write q
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
							dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("totalamount")))
							rsTempPurchaseRequest.movenext
						wend
					End if								
					rsTempPurchaseRequest.Close : set rsTempPurchaseRequest = Nothing
					
					totalppn = dtl_totalamount  * cdbl(to_nol(ppn))/100
					grandtotal = totalppn +dtl_totalamount

					q = " Update invoiceservicehdr set totalppn = replace('" & totalppn & "',',','')" & _
					 ", grandtotal = replace('" & grandtotal & "',',','')" & _
					 ", LASTUPDATE = sysdate " & _
					 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					 " where  invoiceno='" & invoiceno & "' and fcBA='" & businessunit & "'"
					' response.write q
					connStr.execute(q)
							
					amount_spo = 0
					strLineMax = " SELECT NVL(grandtotal,0) grandtotal FROM servicepurchaseorder WHERE fcba = '" & businessunit & "' AND pocode = '"&fcCode&"'"
					Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
					rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
					
					if not rsstrLineMax.eof then
						amount_spo = cdbl(to_nol(rsstrLineMax("grandtotal")))
					else
						amount_spo = 0
					end if									
					rsstrLineMax.Close : set rsstrLineMax = Nothing
				else
					err.number = 1001
					err.description = "Maksimal Amount + Charge is "&FormatNumber(max_amountcharge,0)&", has Invoiced "&FormatNumber(in_data_amountcharge,0)
				end if
			
			else
				err.number = 1001
				err.description = cek_trigger(TARGETTYPE, TARGETCODE, JOBCODE, businessunit, transdate, "PR")
			end if
		
		else
			err.number = 1001
			err.description = "Jobcode, Target Type dan Target Code is Required"
		end if
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
      cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		connstr.BeginTrans
		
		max_amount = 0 : max_charge = 0 : max_amountcharge = 0
		strLineMax = 	" SELECT NVL(SUM(amount),0) amount, NVL(SUM(othercharge),0) othercharge, NVL(SUM(amountcharge),0) amountcharge" & _
						" FROM servicepurchaseorderdetails" & _
						" WHERE pocode = '"&fcCode&"'" & _
						"       AND fcba = '" & businessunit & "'" & _
						"       AND linenokey = '" & spolineno & "'"

		Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
		rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
		
		if not rsstrLineMax.eof then
			max_amount = cdbl(to_nol(rsstrLineMax("amount")))
			max_charge = cdbl(to_nol(rsstrLineMax("othercharge")))
			max_amountcharge = cdbl(to_nol(rsstrLineMax("amountcharge")))
		else
			max_amount = 0 : max_charge = 0 : max_amountcharge = 0
		end if							
		rsstrLineMax.Close : set rsstrLineMax = Nothing
	   
		in_data_amount = 0 : in_data_charge = 0 : in_data_amountcharge = 0
		strLineMax = 	" SELECT NVL(SUM(amount),0) amount, NVL(SUM(othercost),0) othercharge, NVL(SUM(totalamount),0) amountcharge" & _
						" FROM invoiceservicedtl" & _
						" WHERE pocode = '"&fcCode&"'" & _
						"       AND fcba = '" & businessunit & "'" & _
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
						"       AND fcba = '" & businessunit & "'" & _
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
		
		if ( in_data_amountcharge - edit_data_amountcharge ) + cdbl(to_nol(replace(AMOUNT,",",""))) + cdbl(to_nol(replace(OTHERCOST,",",""))) <= max_amountcharge then
	   
			   q = " UPDATE invoiceservicedtl" & _
			       " SET jobcode = '" & jobcode & "'" & _
			       "     , targetcode = '" & targetcode  & "'" & _
			       "     , targettype = '"& targettype & "'" & _
			       "     , amount = REPLACE('" & AMOUNT & "',',','')" & _
			       "     , othercost = REPLACE('" & OTHERCOST & "',',','')" & _
			       "     , totalamount = REPLACE('" & AMOUNT & "',',','') +REPLACE('" & OTHERCOST & "',',','')" & _
			       "     , LASTUPDATE = SYSDATE " & _
			       "     , LASTTIME = TO_CHAR(SYSDATE,'HH24:MI') "& _
			       " WHERE invoiceno = '" & invoiceno & "'" & _
			       "       AND fcBA='" & businessunit & "'" & _
			       "       AND lineno = '"&lineno&"'"
				connStr.execute(q)
				
					'Untuk menghhitung jumlah total di header
					dtl_totalamount = 0
					dtl_charge = 0
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select *  from invoiceservicedtl " & _
					  " where invoiceno='" & invoiceno & "' and " & _
					  " fcBA='" & businessunit & "'"
					  ' response.write q
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
							dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("totalamount")))
							rsTempPurchaseRequest.movenext
						wend 
						
					End if
					
					rsTempPurchaseRequest.Close
					set rsTempPurchaseRequest = Nothing
					totalamount = dtl_totalamount
					totalppn = totalamount  * cdbl(to_nol(ppn))/100
					totalcharge = dtl_charge 
					grandtotal = totalppn + totalamount

				  q= " Update invoiceservicehdr set totalppn = replace('" & totalppn & "',',','')" & _
					 ", grandtotal = replace('" & grandtotal & "',',','')" & _
					 ", LASTUPDATE = sysdate " & _
					 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					 " where  invoiceno='" & invoiceno & "' and fcBA='" & businessunit & "'"
					' response.write q
					connStr.execute(q)
					
		else
			err.number = 1001
			err.description = "Maksimal Amount + Charge is "&FormatNumber(max_amountcharge,0)&", has Invoiced "&FormatNumber(( in_data_amountcharge - edit_data_amountcharge ),0)
		end if
				
		if err.number  = 0 then
		   connStr.CommitTrans
		end if
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
	   response.Write strQueryStatus
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

 Dim rsPurchase
' untuk menampilkan data field header

 set rsPurchase = server.CreateObject("Adodb.recordset")
 q=	  " select * from invoiceservicehdr hdr" & _
	  " left join businessunit bu on (bu.fccode = hdr.fcba)" & _
	  " where hdr.invoiceno='" & invoiceno & "'"
	  'response.write q
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
	transdate = rsPurchase("transdate")
	fcCode = rsPurchase("pocode")
	suppliername = rsPurchase("suppliername")
	ppn = rsPurchase("ppn")
	totalppn = rsPurchase("totalppn")
	businessunit  = rsPurchase("fcba")
	fccompanycode  = rsPurchase("fccompanycode")
	grandtotal  = rsPurchase("grandtotal")
	fakturno  = rsPurchase("no_faktur")
	fakturdate  = rsPurchase("fakturdate")
 End if
 

 ' dim rsFrez
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
 
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="invoiceservice.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Invoice No</td>
				<td align="left">
					<input type="text" maxlength="50" name="invoiceno" size="30" value="<%=invoiceno%>" autocomplete="off">
					<font color="#FF0000">*</font>
				</td>
				</td>
			<td align="left">No Faktur</td>
					<td align="left">
						<input type="text" maxlength="50" name="fakturno"  size="30" value="<%=fakturno%>" autocomplete="off" >
				</td>	
		</tr>
		<tr>
			<td align="left">Company</td>
				<td align="left">
					<select name="fccompanycode" value="<%=fccompanycode%>" id="fccompanycode" onchange="getFcba(this.value);">
					<option value=""></option> 
						<% 	Dim rsComboC
							q = " SELECT DISTINCT fccompanycode " & _
							 	" FROM (SELECT iplasprod.businessunit.fccompanycode " & _
							    " FROM iplasprod.user_fcba_tab " & _
							    " INNER JOIN iplasprod.businessunit " & _
							    " ON iplasprod.user_fcba_tab.fcba = " & _
							    " iplasprod.businessunit.fccode " & _
							    " WHERE iplasprod.businessunit.isactive = 'TRUE' AND iplasprod.user_fcba_tab.fcusercode = '"&  Session("CodeName") &"')"
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
			<td align="left"> Faktur Date</td>
				<td align="left">
					<input class="datepicker" type="text" maxlength="10" name="fakturdate" size="10" value="<%=fakturdate%>">
				</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
				<td align="left">
					<select name="businessunit" id="businessunit"  onchange="getPO(this.value);" >
						<option value=""></option> 
						<% 	Dim rsComboB
							q = "  SELECT  * FROM businessunit WHERE fccompanycode = '"& fccompanycode &"'"
							' response.write q
							Set rsComboB = Server.CreateObject("ADODB.Recordset")
							rsComboB.CursorType = 3
							rsComboB.CursorLocation = 3
							rsComboB.Open q, connstr
							if not rsComboB.eof then%>
							<%
							Do While Not rsComboB.eof
							%>
								<option value="<%=rsComboB("fccode")%>" 
									<%if cstr(""&businessunit)=cstr(""&rsComboB("fccode")) then %> selected 
									<%End If%>><%=rsComboB("fccode")%>
								</option> 
							<%rsComboB.movenext
							loop
							End if
							rsComboB.Close
							set rsComboB = Nothing
							q=""
						%>
					</select>
					<font color="#FF0000">*</font>
				</td>
			<td align="left">Tax (%)</td>
					<td align="left">
						<input type="text" size="5" name="ppn" value="<%=ppn%>" onkeypress="return harusangka(event);">
						<input type="text" size="5" name="totalppn" value="<%=totalppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Date</td>
				<td align="left">
					<input class="datepicker" type="text" maxlength="10" name="transdate" size="10" value="<%=transdate%>">
				</td>
			<td align="left">Grand Total</td>
					<td align="left">
						<input type="text" maxlength="50" name="grandtotal" size="15" value="<%=grandtotal%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Purchase Order No</td>
				<td align="left">
					<select name="fcCode" id="fccode" onchange="javascript:getSupp();">
						<option value="<%=fcCode%>" <% If cstr(""&fcCode) = Request("fcCode") Then %> Selected <%End if%>><%=fcCode%></option> 
					</select>
					<font color="#FF0000">*</font>
		</tr>
		<tr>
			<td align="left">Supplier</td>
					<td align="left">
						<input type="text" maxlength="50" id="suppliername" name="suppliername" style="background:#eee"  size="15" value="<%=suppliername%>" readonly>
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
		<a href="printingpo.Asp?pocode=<%=fcCode%>" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp;
		<%else%>
	    <a href="PrintPO.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&PoDate=<%=PoDate%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&Mode=Print" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp;
		<%end if%>
		<a href="invoiceservice.asp?invoiceno=<%=invoiceno%>&fcCode=<%=fcCode%>&transdate=<%=transdate%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" And cint(IsInBridge) = 1 then %>
		    <a href="invoiceservice.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=Ordered" onClick="return confirm('Are you sure to Order this data :  <%=fcCode%> ?');"> Ordered </a>
		<% elseif RowStateId = "Planned" And cint(IsInBridge) = 0 then %>
		    <a href="invoiceservice.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to Confirm this data :  <%=fcCode%> ?');"> Confirmed </a>
		<% elseif RowStateId = "Confirmed" then %>	
		    <a href="invoiceservice.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=fcCode%> ?');"> Re-Planned </a>
		<% elseif  RowStateId = "Ordered" And cint(IsInBridge) = 1 And Cint(IsFrezee) =0 then %>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="invoiceservice.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=UnRelease" onClick="return confirm('Are you sure to Un Release this data :  <%=fcCode%> ?');"> UnRelease </a>
		<% end if%>
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" AND LEN(TRIM(invoiceno)) > 0 AND LEN(TRIM(fccompanycode)) > 0 AND LEN(TRIM(businessunit)) > 0 AND LEN(TRIM(fcCode)) > 0  Then
	Dim rsPurchaseOrder,intRecordsPerPage,cColor
	

	q = " select spo.invoiceno, spo.transdate, spo.no_faktur, spo.fakturdate, spo.pocode," & _
		" spd.jobcode, spd.targetcode, spd.targettype, spd.description," & _
		" spd.amount, spd.description, spd.amount, spd.othercost, spd.totalamount, spd.lineno, spd.polinenokey" & _
		" from invoiceservicehdr spo " & _
		" inner join invoiceservicedtl spd " & _
		" on (spo.invoiceno = spd.invoiceno and spo.fcba = spd.fcba)" & _
		" where spo.invoiceno = '"& invoiceno &"'" & _
		" ORDER BY lineno"
	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.CursorType = 3
	rsPurchaseOrder.CursorLocation = 3
	rsPurchaseOrder.Open q, connstr
%>
   


   
   <form action="invoiceservice.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td align="left">Description</td>
			<td align="left">
				<select name="spolineno" id="spolineno" onchange="javascript:getSPODetailsInfo();" >
				<option value=""></option> 
				<% 	q = "  SELECT linenokey fccode, stockname fcname FROM servicepurchaseorderdetails WHERE pocode = '"& fcCode &"'"
					' response.write q
					Set rsComboB = Server.CreateObject("ADODB.Recordset")
					rsComboB.Open q, connstr
					if not rsComboB.eof then%>
					<%Do While Not rsComboB.eof%>
						<option value="<%=rsComboB("fccode")%>" <%if cstr(""&spolineno)=cstr(""&rsComboB("fccode")) then %> selected <%End If%>><%=rsComboB("fcname")%></option> 
					<%	rsComboB.movenext
					loop
					End if
					rsComboB.Close : set rsComboB = Nothing : q=""
				%>
			</select>
			<font color="#FF0000">*</font>
			<input type="Hidden" id="descriptionx" name="descriptionx" value="<%=descriptionx%>">
			</td>
		</tr>
   		<tr>
			<td align="left">Job Code</td>
                <td align="left">
                    <input type="text" size="11" name="JobCode" value="<%=JobCode%>" autocomplete="off">&nbsp;<font color="#FF0000">*</font>
                    	<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
					<font face="Verdana" size="1">Activity</font>
                </td>
			
		</tr>
		<tr>
			<td align="left">Target Type</td>
                <td align="left">
                    <input type="text" size="11" name="targettype" value="<%=targettype%>" autocomplete="off">&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                    <font face="Verdana" size="1">Location Type</font></a>
                </td>
                
		</tr>
		<tr>
		    <td align="left">Target Code</td>
                <td align="left">
                    <input type="text" size="11" name="targetcode" value="<%=targetcode%>" autocomplete="off">&nbsp;<font
                        color="#FF0000">*</font>
                </td>
			</tr>
		<tr>
			<td align="left">SPO Amount</td>
                <td align="left">
                   <input type="text" size="15" id="spo_amount" name="spo_amount" value="<%=spo_amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  readonly style="background: #eee;">
                </td>
			<td align="left">Post Amount</td>
				<td align="left">
					<input type="text" size="15" id="amount" name="amount" value="<%=amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" autocomplete="off" ><font color="#FF0000"> *</font>
				</td>
		</tr>
		<tr>
			<td align="left">SPO Charge</td>
                <td align="left">
                    <input type="text" size="15" id="spo_charge" name="spo_charge" value="<%=spo_charge%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  readonly style="background: #eee;">
                </td>
                <td align="left">Post Charge</td>
					<td align="left">
						<input type="text" size="15" id="othercost" name="othercost" value="<%=othercost%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" ><font color="#FF0000"> *</font>
				</td>
		</tr>
		<tr>
		   
                <td align="left">Total SPO Amount</td>
					<td align="left">
						<input type="text" size="15" id="spo_totalamount" name="spo_totalamount" value="<%=spo_totalamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly style="background: #eee;">
				</td>
			</tr>
		
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="transdate" value="<%=transdate%>">
		  <input type="hidden" name="fakturno" value="<%=fakturno%>">
		  <input type="hidden" name="fakturdate" value="<%=fakturdate%>">
		  <input type="hidden" name="suppliername" value="<%=suppliername%>">
		  <input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		  <input type="hidden" name="businessunit" value="<%=businessunit%>">
		  <input type="hidden" name="invoiceno" value="<%=invoiceno%>">
		  <input type="hidden" name="dp" value="<%=dp%>">
		  <input type="hidden" name="ppn" value="<%=ppn%>">
		  <input type="hidden" name="grandtotal" value="<%=grandtotal%>">
		  <input type="hidden" name="totalppn" value="<%=totalppn%>">
		  <input type="hidden" name="RowStatehdr" value="<%=RowStatehdr%>">
		  <input type="hidden" name="lineno" value="<%=lineno%>">

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

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="middle" colspan="2"> ACTION</TD>
	
	<TD align="left">JOBCODE</TD>
	<TD align="left">TARGET TYPE</TD>
	<TD align="left">TARGET CODE</TD>	
	<TD align="left">DESCRIPTION</TD>
	<TD align="left">AMOUNT</TD>
	<TD align="left">OTHER COST</TD>
	<TD align="left">TOTAL AMOUNT</TD>
	</tr>   
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
				<td valign="middle">
				<a href="invoiceservice.asp?invoiceno=<%=rsPurchaseOrder("invoiceno")%>&jobcode=<%=rsPurchaseOrder("jobcode")%>&targetcode=<%=rsPurchaseOrder("targetcode")%>&targettype=<%=rsPurchaseOrder("targettype")%>&description=<%=rsPurchaseOrder("description")%>&amount=<%=rsPurchaseOrder("amount")%>&othercost=<%=rsPurchaseOrder("othercost")%>&totalamount=<%=rsPurchaseOrder("totalamount")%>&fccompanycode=<%=fccompanycode%>&fccode=<%=fccode%>&businessunit=<%=businessunit%>&lineno=<%=rsPurchaseOrder("lineno")%>&spolineno=<%=rsPurchaseOrder("polinenokey")%>&txtJenis=Edit"> <img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(1)%>" border="0"></a>
				</td>
				<td>
				<a href="invoiceservice.asp?invoiceno=<%=rsPurchaseOrder("invoiceno")%>&jobcode=<%=rsPurchaseOrder("jobcode")%>&targetcode=<%=rsPurchaseOrder("targetcode")%>&targettype=<%=rsPurchaseOrder("targettype")%>&description=<%=rsPurchaseOrder("description")%>&amount=<%=rsPurchaseOrder("amount")%>&othercost=<%=rsPurchaseOrder("othercost")%>&fccompanycode=<%=fccompanycode%>&fccode=<%=fccode%>&businessunit=<%=businessunit%>&lineno=<%=rsPurchaseOrder("lineno")%>&ppn=<%=ppn%>&txtJenis=Details"> <img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');">
				</td>
				
				<td align="left" title="<%=rsPurchaseOrder("jobcode").Name%>" valign="top">
					<%=rsPurchaseOrder("jobcode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("targettype").Name%>" valign="top">
					<%=rsPurchaseOrder("targettype")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("targetcode").Name%>" valign="top">
					<%=rsPurchaseOrder("targetcode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("description").Name%>" valign="top">
					<%=rsPurchaseOrder("description")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("Amount").Name%>" valign="top">
					<%=rsPurchaseOrder("Amount")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("othercost").Name%>" valign="top">
					<%=rsPurchaseOrder("othercost")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("totalamount").Name%>" valign="top">
					<%=rsPurchaseOrder("totalamount")%>
				</td>
			</tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>
</div></div>
  <!--#include file="bottom.jsp"-->
		<script src="/js/jquery.inputmask.bundle.min.js"></script>
		<script src="/js/inputmask.binding.js"></script>
<script>
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
	getSPODetailsInfo();
});
</script>
</body>
</html>
