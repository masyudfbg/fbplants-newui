<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Service Purchase Order</title>
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
	var suppliercode = document.getElementById("suppliercode").value;
	var purchaseby = document.getElementById("purchaseby").value;
	var PostStr = 'suppliercode='+suppliercode+'&purchaseby='+purchaseby;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getSupplierSPO.asp', 'POST', PostStr, showMessageResponse);
	
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

function getSupplierPO(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("purchaseby").innerHTML="";
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
		document.getElementById("suppliercode").innerHTML=xmlhttp.responseText;
		// document.getElementById("purchaseby").innerHTML=xmlhttp.responseText;
		// document.getElementById("receviedtask").innerHTML="";
		// alert(businessunit);
		}
	  }
	// xmlhttp.open("GET","ajax/getSupplierPO.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
	xmlhttp.open("GET","ajax/getSupplierPO.asp?fcba="+str+"",true);
	xmlhttp.send();
}


function getPRLine(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("prcode").innerHTML="";
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
		document.getElementById("prlineid").innerHTML=xmlhttp.responseText;
		document.getElementById("requestedqty").innerHTML=xmlhttp.responseText;
		// document.getElementById("receviedtask").innerHTML="";
		// alert(PRLINEID);
		}
	  }
	// xmlhttp.open("GET","ajax/getSupplierPO.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
	xmlhttp.open("GET","ajax/getPRLine.asp?prcode="+str+"&businessunit="+document.getElementById("businessunit").value+"",true);
	xmlhttp.send();
}
					
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);

function getFCBA()
{
	var fcba = $('#businessunit').val();
	$('#selectedfcba').val(fcba);
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

 ' if request("fccOde")="" then
 '    fcCode="" '----- Nantinya Ambil dari Format Master
 ' else
 '    fcCode = request("fccOde")
 ' end if
 
 fcCode = request("fccOde")
 If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"

 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("PODATE")<>"" then
   PODATE   =Request("PODATE")
 else
   PODATE   =textDate(date)
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
 deliverydestination = Request("deliverydestination")
 StockName = Request("StockName")
 requestedqty = Request("requestedqty")
 unitprice = Request("unitprice")
 disc = Request("disc")
 othercharge = Request("othercharge")
 totalamount = Request("totalamount")
 totalamountppn = Request("totalamountppn")
 totalcharge = Request("totalcharge")
 grandtotal = Request("grandtotal")
 amountcharge = Request("amountcharge")
 totalppn = Request("totalppn")
 approvedqty = request ("approvedqty")

if "" & use_ppn = "" then use_ppn = "N" 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "Ordered" or Request("txtJenis")= "CompletedDetails"  or  Request("txtJenis")= "CanceledDetails"  or  Request("txtJenis")= "UnRelease" Then
	
	 
	
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
       connstr.BeginTrans
				
		q="Delete from SERVICEPURCHASEORDERDETAILS where  POCODE='" & fcCode & "' and fcBA='" & businessunit & "'"
		connStr.execute(q)

		q="Delete from SERVICEPURCHASEORDER where  POCODE='" & fcCode & "' and fcBA='" & businessunit & "'"
		connStr.execute(q)
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans
			   q="Delete from SERVICEPURCHASEORDERDETAILS where  POCODE='" & fcCode & "' and LineNoKey='" & request("LineNoKey") & "' and fcBA='" & businessunit & "'"
			   connStr.execute(q)

			   'Untuk menghhitung jumlah total di header
				dtl_totalamount = 0
				dtl_charge = 0
				set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
				q="select *  from SERVICEPURCHASEORDERDETAILS " & _
				  " where poCode='" & fcCode & "' and " & _
				  " fcBA='" & businessunit & "'"
				rsTempPurchaseRequest.Open q, connstr
				if Not rsTempPurchaseRequest.Eof then
					while Not rsTempPurchaseRequest.Eof 
						dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("amount")))
						dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercharge")))
						rsTempPurchaseRequest.movenext
					wend
				End if
				
				rsTempPurchaseRequest.Close
				set rsTempPurchaseRequest = Nothing
				
				totalamount = dtl_totalamount
				totalppn = totalamount  * cdbl(to_nol(ppn))/100
				totalamountppn = totalamount + totalppn
				totalcharge = dtl_charge 
				grandtotal = totalamountppn + totalcharge
				'end add

			  q= " Update SERVICEPURCHASEORDER set totalamountppn=replace('" & totalamountppn & "',',','')" & _
				 ", totalcharge = replace('" & totalcharge & "',',','')" & _
				 ", totalppn = replace('" & totalppn & "',',','')" & _
				 ", totalamount = replace('" & totalamount & "',',','')" & _
				 ", grandtotal = replace('" & grandtotal & "',',','')" & _
				 ", LASTUPDATE = sysdate " & _
				 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
				 " where  POCode='" & fcCode & "' and fcBA='" & businessunit & "'"
				connStr.execute(q)
			  
			   if err.number  = 0 then
				  connStr.CommitTrans
			   end if
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
	connstr.BeginTrans
	  Dim rsPO, rsPurchases,rsQty

			'--Add by Ukie to handle Auto Number of PR
			If Ucase(fcCode) = Ucase("(NEW)") Then
				oSql = " select nvl(controlnumber_api.get_nextnumber(upper('SERVICEPURCHASEORDER')" & _
					   "                                         , " & year(textDate(PoDate)) & "" & _
					   "                                         , " & month(textDate(podate)) & "" & _
					   "                                         , '" & businessunit & "'" & _
					   " ), 'FALSE') pocode from dual"
				
				' Response.Write oSql
					   
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.Open oSql, connstr, adLockOptimistic, adOpenStatic
				If Not oRs.EOF Then
					txtPOCode = oRs("poCode")
				End If
				oRs.Close
				Set oRs = Nothing
					fcCode = txtPOCode
			End If
			
			If txtPOCode = "FALSE" Then
				err.Number = 1001
				err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
			End If
			
			'Response.Write PRCODE & txtPRCode
		'--End Added
	  Set rsPO = Server.CreateObject("ADODB.Recordset")
			  rsPO.Open "Select poCode from SERVICEPURCHASEORDER where poCode='" & fcCode & "' AND FCBA = '" & businessunit & "' AND company='"& fccompanycode &"'" , connstr
			  IF rsPO.eof then	   
				 ' Modified by NAF
				 ' To handle first time Input for Purchase Order

				 q = " INSERT INTO Servicepurchaseorder ( pocode, podate, suppliercode, use_ppn, ppn, totalamount, company, suppliername, purchaseby, termofpayment, dp, totalamountppn, totalcharge, totalppn, grandtotal, fcentry, fcedit, fcip, fcba, rowstateid, lastupdate, lasttime ) " & _
				     " VALUES ( '" & fcCode & "'" & _
					 "         , TO_DATE('" & PoDate & "', 'mm/dd/yyyy')" & _
					 "         , '" & xCode & "'" & _
					 "         , '" & use_ppn & "'" & _
					 "         , '" & ppn & "'" & _
					 " 		   , replace('" & totalamount & "',',','')" & _
					 "         , '" & fccompanycode & "'" & _
					 "         , '" & suppliername & "'" & _
					 "         , '" & purchaseby & "'" & _
					 "         , '" & termofpayment & "'" & _
					 " 		   , replace('" & dp & "',',','')" & _
					 " 		   , replace('" & totalamountppn & "',',','')" & _
					 " 		   , replace('" & totalcharge & "',',','')" & _
					 " 		   , replace('" & totalppn & "',',','')" & _
					 " 		   , replace('" & grandtotal & "',',','')" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					 "         , '" & businessunit & "'" & _
					 "         , 'Planned'" & _
					 "         , SYSDATE" & _
					 "         , TO_CHAR(SYSDATE,'hh24:mi'))"
				 connStr.execute(q)			
			  End if
			  rsPO.Close
			  
				' Added by NAF, 2012-12-01
				 strLineMax = " select nvl(max(linenokey),0) maxLineNoKey from SERVICEPURCHASEORDERDETAILS where fcBA='" & businessunit & "'"

				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
				end if
							
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 
				' End of Added by NAF, 2012-12-01
				 
				 
							amount = 0
						 	amount1 = cdbl(to_nol(quantity)) * cdbl(to_nol(unitprice))
							amount2 = amount1 * cdbl(to_nol(disc))/100
							amount3 = amount1 - amount2

							AMOUNT = amount + amount3
							AMOUNTCHARGE = amount3 + cdbl(to_nol(othercharge))
							
							q="Insert into  SERVICEPURCHASEORDERDETAILS (POCODE, QUANTITY, AMOUNT, OTHERITEMDESC, RECEIVED, UNITPRICE, STOCKNAME, DISC, OTHERCHARGE, AMOUNTCHARGE, FCENTRY, FCEDIT, FCIP, FCBA, LINENOKEY, LASTUPDATE, LASTTIME, ROWSTATEID ) " & _
							 "values ('" & fcCode & "','"  &  QUANTITY & "',replace('" & AMOUNT & "',',','')" & _
							 ",'" & OTHERITEMDESC & "'," & "0" & ",replace('" & UNITPRICE & "',',',''),'" & STOCKNAME & "',replace('" & DISC & "',',',''),replace('" & OTHERCHARGE & "',',',''),replace('" & AMOUNTCHARGE & "',',',''),'" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
							 Request.ServerVariables("REMOTE_ADDR") &"','" & businessunit & "',"& tempMaxLineNo & ", sysdate, To_Char(sysdate,'HH24:MI'), 'Planned')"
							 
							connStr.execute(q)				
							' response.write q
							
							tempStringError = "Service PO Saved Successfully, if you want to modify just click Edit Button, Thank you."				 
							response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
							'Untuk menghhitung jumlah total di header
							dtl_totalamount = 0
							dtl_charge = 0
							set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
							q="select *  from SERVICEPURCHASEORDERDETAILS " & _
							  " where poCode='" & fcCode & "' and " & _
							  " fcBA='" & businessunit & "'"
							rsTempPurchaseRequest.Open q, connstr
							if Not rsTempPurchaseRequest.Eof then
								while Not rsTempPurchaseRequest.Eof 
									dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("amount")))
									dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercharge")))
									rsTempPurchaseRequest.movenext
								wend
							End if
							
							rsTempPurchaseRequest.Close
							set rsTempPurchaseRequest = Nothing
							totalamount = dtl_totalamount
							totalppn = totalamount  * cdbl(to_nol(ppn))/100
							totalamountppn = totalamount + totalppn
							totalcharge = dtl_charge 
							grandtotal = totalamountppn + totalcharge
							'end add

						  q= " Update SERVICEPURCHASEORDER set totalamountppn=replace('" & totalamountppn & "',',','')" & _
							 ", totalcharge = replace('" & totalcharge & "',',','')" & _
							 ", totalppn = replace('" & totalppn & "',',','')" & _
							 ", totalamount = replace('" & totalamount & "',',','')" & _
							 ", grandtotal = replace('" & grandtotal & "',',','')" & _
							 ", LASTUPDATE = sysdate " & _
							 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							 " where  POCode='" & fcCode & "' and fcBA='" & businessunit & "'"
							connStr.execute(q)
							
			  if err.number  = 0 then
				 connStr.CommitTrans
			  end if
      cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		connstr.BeginTrans
		
		amount = 0
		amount1 = cdbl(to_nol(quantity)) * cdbl(to_nol(unitprice))
		amount2 = amount1 * cdbl(to_nol(disc))/100
		amount3 = amount1 - amount2
		AMOUNT = amount + amount3
		AMOUNTCHARGE = amount3 + cdbl(to_nol(othercharge))
			   q="Update servicepurchaseorderdetails set Quantity='" & Quantity & "',Amount = replace('" & Amount & "',',','')" & _
				 ",OTHERITEMDESC='" & OTHERITEMDESC  & "'" & _
				 ", unitprice = replace('" & unitprice & "',',','')" & _
				 ", othercharge = replace('" & othercharge & "',',','')" & _
				 ", disc = replace('" & disc & "',',','')" & _
				 ", StockName = '" & StockName & "'" & _
				 ", LASTUPDATE = sysdate " & _
				 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
				 " where  POCode='" & fcCode & "' AND LineNoKey='" & LineNoKey & "' and fcBA='" & businessunit & "'"
				connStr.execute(q)
				
					dtl_totalamount = 0
					dtl_charge = 0
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select *  from SERVICEPURCHASEORDERDETAILS " & _
					  " where poCode='" & fcCode & "' and " & _
					  " fcBA='" & businessunit & "'"
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						while Not rsTempPurchaseRequest.Eof 
							dtl_totalamount = dtl_totalamount + cdbl(to_nol(rsTempPurchaseRequest("amount")))
							dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercharge")))
							rsTempPurchaseRequest.movenext
						wend
					End if
					
					rsTempPurchaseRequest.Close
					set rsTempPurchaseRequest = Nothing
					totalamount = dtl_totalamount
					totalppn = totalamount  * cdbl(to_nol(ppn))/100
					totalamountppn = totalamount + totalppn
					totalcharge = dtl_charge 
					grandtotal = totalamountppn + totalcharge
					'end add

				  q= " Update SERVICEPURCHASEORDER set totalamountppn=replace('" & totalamountppn & "',',','')" & _
					 ", totalcharge = replace('" & totalcharge & "',',','')" & _
					 ", totalppn = replace('" & totalppn & "',',','')" & _
					 ", totalamount = replace('" & totalamount & "',',','')" & _
					 ", grandtotal = replace('" & grandtotal & "',',','')" & _
					 ", LASTUPDATE = sysdate " & _
					 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					 " where  POCode='" & fcCode & "' and fcBA='" & businessunit & "'"
					connStr.execute(q)
				
				if err.number  = 0 then
				   connStr.CommitTrans
				end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Confirmed" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN SERVICEPURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'Confirm' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "RePlanned" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN SERVICEPURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'Replan' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "Ordered" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN SERVICEPURCHASEORDER_API.NextState('" & businessunit & "','" & fcCode & "', 'Order' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
  
	elseif Request("txtJenis")= "CompletedDetails" then
		connstr.BeginTrans
		strQueryStatus = " BEGIN" & _
						 " SERVICEPURCHASEORDER_API.NextState('" & businessunit & "'" & _
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
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q= " select * from SERVICEPURCHASEORDER spo  " & _
 	" left join businessunit bu on (bu.fccode = spo.fcba)" & _
  " where spo.PoCode='" & fcCode & "'"
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    PoDate=rsPurchase("PoDate")
	PODATE = FormatDateTime(PODATE)
	xCode=rsPurchase("SUPPLIERCODE")
	RowStateId = rsPurchase("RowStateId")
	RowStateHdr = rsPurchase("RowStateId")
	use_ppn = rsPurchase("use_ppn")
	suppliername  = rsPurchase("suppliername")
	termofpayment  = rsPurchase("termofpayment")
	ppn  = rsPurchase("ppn")
	dp  = rsPurchase("dp")
	fcba  = rsPurchase("fcba")
	fccompanycode = rsPurchase("fccompanycode")
	businessunit  = rsPurchase("fcba")
	purchaseby = rsPurchase("purchaseby")
	totalamountppn = rsPurchase("totalamountppn")
	totalcharge = rsPurchase("totalcharge")
	totalamount  = rsPurchase("totalamount")
	totalppn  = rsPurchase("totalppn")
	grandtotal  = rsPurchase("grandtotal")
 End if
 

 dim rsFrez
 set rsFrez = server.CreateObject("Adodb.recordset")
	q=" select distinct frezee  from SERVICEPURCHASEORDERDETAILS " & _
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
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-wrench" aria-hidden="true"></i>Header Service Purchase Order</div><% end if %>
 <form name="frmEntry" method="post" action="servicepurchaseorder.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Purchase Order No</td>
			<%
			if fcCode = "(NEW)" then
				q = "select '" & Session("fcBA") & "/'|| MAX(cn.PREFIX) || LPAD(NVL(MAX(regexp_substr(po.POCODE, '[^/]+', 1, 5)),0000)+0001,4,0) DOCNUMBER " & _
					" from " & _
					" controlnumber cn " & _
					" left join SERVICEPURCHASEORDER po on cn.fcba = po.fcba and cn.accountingyear = to_number(to_char(po.PODATE,'yyyy')) " & _
					" where cn.fcba ='" & Session("fcBA") & "' " & _
					" and cn.LU_NAME ='SERVICEPURCHASEORDER' "&_
					" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& PODATE &"','mm/dd/yyyy'),'yyyy')) " & _
					" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& PODATE &"','mm/dd/yyyy'),'mm')) "
				'Response.write q
				set rsPONo = server.CreateObject("Adodb.recordset")
				rsPONo.Open q, connstr
				if Not rsPONo.Eof then
					fcCode = rsPONo("DOCNUMBER")
				End if
			End if
			%>
				<td align="left">
					<input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>" autocomplete="off">
					<font color="#FF0000">*</font>
				</td>
			
				<td align="left">Term Of Payment</td>
				<td align="left">
					<select name="termofpayment" value="<%=termofpayment%>">
						<option value="<%=termofpayment%>"></option>
						<option value="TOP7" <% If cstr(""&termofpayment) = "TOP7" Then %> Selected <%End if%>>TOP7</option> 
						<option value="TOP14" <% If cstr(""&termofpayment) = "TOP14" Then %> Selected <%End if%>>TOP14</option> 
						<option value="TOP30" <% If cstr(""&termofpayment) = "TOP30" Then %> Selected <%End if%>>TOP30</option> 
					</select>
				</td>
		</tr>
		<tr>
			<td align="left">Company</td>
				<td align="left">
					<select name="fccompanycode" value="<%=fccompanycode%>" id="fccompanycode" onchange="getFcba(this.value);">
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
			<td align="left">DP</td>
				<td align="left">
					<input type="text" maxlength="50" name="dp" size="30" value="<%=dp%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
				</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
				<td align="left">
					<select name="businessunit" id="businessunit" onchange="">
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
					<input type="hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
				</td>
			<td align="left">PPn (%)</td>
					<td align="left">
						<input type="text" size="5" name="ppn" value="<%=ppn%>" onkeypress="return harusangka(event);">
						<input type="text" size="5" name="totalppn" value="<%=totalppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Date</td>
				<td align="left">
					<input class="datepicker" type="text" maxlength="10" name="PODATE" size="10" value="<%=PODATE%>">
				</td>
			<td align="left">Total Amount</td>
					<td align="left">
						<input type="text" maxlength="50" name="totalamount" size="15" value="<%=totalamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Purchase By</td>
				<td align="left">      
					<select name="purchaseby" id="purchaseby" onchange="getSupplierPO(this.value);" >
						<option value="<%=purchaseby%>" <% If cstr(""&purchaseby) = Request("purchaseby") Then %> Selected <%End if%>><%=purchaseby%></option> 
					</select>
				</td>
				<td align="left">Total Amount + PPn</td>
					<td align="left">
						<input type="text" maxlength="50" name="totalamountppn" size="15" value="<%=totalamountppn%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
			<tr>
				<td align="left">Vendor</td>
				<td align="left">      
					<select name="xCode" id="suppliercode" class="chosen-select" onchange="javascript:getSupp();"  >
						<option value="<%=xCode%>" <% If cstr(""&xCode) = Request("xCode") Then %> Selected <%End if%>><%=xCode%></option> 
					</select>
				</td>
				<td align="left">Total Charge</td>
					<td align="left">
						<input type="text" size="15" name="totalcharge" value="<%=totalcharge%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>	
			</tr>
		</tr>
		<tr>
			<td align="left">Supplier Name</td>
				<td align="left">
						<textarea id="suppliername"  name="suppliername" style="background:#eee" readonly><%=suppliername%></textarea>
				</td>
				<td align="left">Grand Total</td>
					<td align="left">
						<input type="text" maxlength="50" name="grandtotal" size="15" value="<%=grandtotal%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee" readonly>
					</td>
		</tr>
		<tr>
			<td align="left">Status</td>
					<td align="left">
						<input type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>" style="background:#eee" readonly>
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
		
		<tr><td colspan="4"  align="left" class="fbp-doc-actions">
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
		<a href="servicepurchaseorder.asp?fcCode=<%=fcCode%>&PODATE=<%=PODATE%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" And cint(IsInBridge) = 1 then %>
		    <a href="servicepurchaseorder.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=Ordered" onClick="return confirm('Are you sure to Order this data :  <%=fcCode%> ?');"> Ordered </a>
		<% elseif RowStateId = "Planned" And cint(IsInBridge) = 0 then %>
		    <a href="servicepurchaseorder.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to Confirm this data :  <%=fcCode%> ?');"> Confirmed </a>
		<% elseif RowStateId = "Confirmed" then %>	
		    <a href="servicepurchaseorder.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=fcCode%> ?');"> Re-Planned </a>
		<% elseif  RowStateId = "Ordered" And cint(IsInBridge) = 1 And Cint(IsFrezee) =0 then %>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="servicepurchaseorder.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=fcCode%>&txtJenis=UnRelease" onClick="return confirm('Are you sure to Un Release this data :  <%=fcCode%> ?');"> UnRelease </a>
		<% end if%>
		
		
		
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" AND LEN(TRIM(fcCode)) Then
	Dim rsPurchaseOrder,intRecordsPerPage,cColor

	q = " select *  from SERVICEPURCHASEORDERDETAILS" & _
		" where PoCode = '" & fcCode & "'" & _
	    "       and fcBA = '" & businessunit & "'" & _
		" ORDER BY linenokey"
	' response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.CursorType = 3
	rsPurchaseOrder.CursorLocation = 3
	rsPurchaseOrder.Open q, connstr
%>
   


   
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Jasa</div><% end if %>
   <form action="servicepurchaseorder.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
   		<tr>
			<td align="left">Description</td>
			<td align="left"><input type="text" size="50" id="stockname" name="StockName" value="<%=StockName%>" autocomplete ="off">
			</td>
			<td align="left">Disc (%)</td>
			<td align="left"><input type="text" size="15" name="disc" value="<%=disc%>" onkeypress="return harusangka(event);" autocomplete="off">&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		<tr>
			<td align="left">Ordered Quantity</td>
				<td align="left"><input type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>" onkeypress="return harusangka(event);" autocomplete="off" required>
			</td>
			<td align="left">Amount</td>
			<td align="left"><input type="text" size="15" name="AMOUNTCHARGE" value="<%=AMOUNTCHARGE%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'"  style="background:#eee;" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Unit Price</td>
			<td align="left"><input type="text" size="15" name="unitprice" value="<%=unitprice%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" autocomplete="off" required>
			</td>
			<td align="left">Received</td>
			<td align="left"><input type="text" size="15" name="received" value="<%=received%>"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#eee;" readonly>
			</td>
		</tr>
		<tr>
			<td align="left">Other Charge</td>
				<td align="left"><input type="text" size="15" name="othercharge" value="<%=othercharge%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" autocomplete="off">
				</td>
			<td align="left">Remarks</td>
			<td align="left">
				<textarea  name="OTHERITEMDESC"  cols="40" rows="3"><%=OTHERITEMDESC%>
				</textarea>
			</td>
		</tr>
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="PODATE" value="<%=PODATE%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="LineNokey" value="<%=LINENOKEY%>">
		  <input type="hidden" name="use_ppn" value="<%=use_ppn%>">
		  <input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		  <input type="hidden" name="businessunit" value="<%=businessunit%>">
		  <input type="hidden" name="purchaseby" value="<%=purchaseby%>">
		  <input type="hidden" name="suppliername" value="<%=suppliername%>">
		  <input type="hidden" name="termofpayment" value="<%=termofpayment%>">
		  <input type="hidden" name="SHIPPINGPOINT" value="<%=SHIPPINGPOINT%>">
		  <input type="hidden" name="dp" value="<%=dp%>">
		  <input type="hidden" name="ppn" value="<%=ppn%>">
		  <input type="hidden" name="totalamount" value="<%=totalamount%>">
		  <input type="hidden" name="totalamountppn" value="<%=totalamountppn%>">
		  <input type="hidden" name="totalcharge" value="<%=totalcharge%>">
		  <input type="hidden" name="grandtotal" value="<%=grandtotal%>">
		  <input type="hidden" name="totalppn" value="<%=totalppn%>">


		  <input type="hidden" name="RowStatehdr" value="<%=RowStatehdr%>">
		  
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

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Jasa</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="middle" colspan="2"> ACTION</TD>
	<TD align="left">PO CODE</TD>
	<TD align="left">LINENOKEY</TD>
	<TD align="left">DESCRIPTION</TD>
	<TD align="left">QUANTITY</TD>
	<TD align="left">UNIT PRICE</TD>
	<TD align="left">DISC</TD>
	<TD align="left">OTHER CHARGE</TD>
	<TD align="left">AMOUNT</TD>
	<TD align="left">OTHER ITEM DESC</TD>
	<TD align="left">RECEIVED</TD>
	<TD align="left">LINE STATUS</TD>
	<TD align="left">FREZEE</TD>
	</tr>
	</thead>
	<tbody>
   <%

     Set rsPO = Server.CreateObject("ADODB.Recordset")
	 rsPO.Open " select rowstateid from servicepurchaseorder where fcba = '"& businessunit &"' and poCode = '"& fcCode &"'" , connstr
	 IF Not rsPO.eof then
		HdrState =  rsPO("rowstateid")
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
			<% If HdrState="Planned" Then %>			
			    <td valign="middle" width="10">
				   <a href="servicepurchaseorder.asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&OTHERITEMDESC=<%=rsPurchaseOrder("OTHERITEMDESC")%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&stockname=<%=rsPurchaseOrder("stockname")%>&unitprice=<%=rsPurchaseOrder("unitprice")%>&amount=<%=rsPurchaseOrder("amount")%>&amountcharge=<%=rsPurchaseOrder("amountcharge")%>&disc=<%=rsPurchaseOrder("disc")%>&received=<%=rsPurchaseOrder("received")%>&othercharge=<%=rsPurchaseOrder("othercharge")%>&txtJenis=Edit">
				   <img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(1)%>" border="0"></a>
				</td>
				<td valign="middle" width="10">
				   <a href="servicepurchaseorder.asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&OTHERITEMDESC=<%=rsPurchaseOrder("OTHERITEMDESC")%>&fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&stockname=<%=rsPurchaseOrder("stockname")%>&unitprice=<%=rsPurchaseOrder("unitprice")%>&amount=<%=rsPurchaseOrder("amount")%>&amountcharge=<%=rsPurchaseOrder("amountcharge")%>&disc=<%=rsPurchaseOrder("disc")%>&received=<%=rsPurchaseOrder("received")%>&othercharge=<%=rsPurchaseOrder("othercharge")%>&ppn=<%=ppn%>&txtJenis=Details">
				   <img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');">
				   </a>
				</td>
            <% ElseIf rsPurchaseOrder("ROWSTATEID")="Ordered"  And cint(IsInBridge) = 0  then %>	
			    <td valign="middle" width="10">
				   <a href="servicepurchaseorder.asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("ItemCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&OTHERITEMDESC=<%=rsPurchaseOrder("OTHERITEMDESC")%>&PRCODE=<%=rsPurchaseOrder("PRCODE")%>&PRLINEID=<%=rsPurchaseOrder("PRLINEID")%>&txtJenis=CanceledDetails" onClick="return confirm('Are you sure to Cancel This data ?');">Cancel</a>
				</td>
				<td valign="middle" width="10">
				   <a href="servicepurchaseorder.asp?fccompanycode=<%=fccompanycode%>&businessunit=<%=businessunit%>&purchaseby=<%=purchaseby%>&fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("ItemCode")%>&PODATE=<%=PODATE%>&txtJenis=CompletedDetails" onClick="return confirm('Are you sure to Close This data ?');">Close</a>
				</td> 
			<% Else %>	
               	<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>		
            <% End If %>	
				
				<td align="left" title="<%=rsPurchaseOrder("pocode").Name%>" valign="top">
					<%=rsPurchaseOrder("pocode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("linenokey").Name%>" valign="top">
					<%=rsPurchaseOrder("linenokey")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("StockName").Name%>" valign="top">
					<%=rsPurchaseOrder("StockName")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("quantity").Name%>" valign="top">
					<%=rsPurchaseOrder("quantity")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("unitprice").Name%>" valign="top">
					<%=rsPurchaseOrder("unitprice")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("disc").Name%>" valign="top">
					<%=rsPurchaseOrder("disc")%> %
				</td>
				<td align="left" title="<%=rsPurchaseOrder("othercharge").Name%>" valign="top">
					<%=rsPurchaseOrder("othercharge")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("amountcharge").Name%>" valign="top">
					<%=rsPurchaseOrder("amountcharge")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otheritemdesc").Name%>" valign="top">
					<%=rsPurchaseOrder("otheritemdesc")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("received").Name%>" valign="top">
					<%=rsPurchaseOrder("received")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("rowstateid").Name%>" valign="top">
					<%=rsPurchaseOrder("rowstateid")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("FREZEE").Name%>" valign="top">
					<%=rsPurchaseOrder("FREZEE")%>
				</td>
			</tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("<tr><td colspan=""15"" class=""fbp-doc-empty"">Data is empty</td></tr>")
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
 	<script type="text/javascript">
		function handleClick(cb){
			if(cb.checked == 1){
				cb.value = "Y";
			} else {
				cb.value = "N";
			}
        }
	</script>
		<script src="/js/jquery.inputmask.bundle.min.js"></script>
		<script src="/js/inputmask.binding.js"></script>
</body>
</html>
