<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Sales Invoice DP</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
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

<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<script>
	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}
</script>
<script>
	function hitungrumus() {
		var sdp_beforevat = $("#dp_beforevat").val();
		var sdp_aftervat = $("#dp_aftervat").val();
		var stotaltodate_dp = $("#totaltodate_dp").val();
		var stotaltodate_perc_dp = $("#totaltodate_perc_dp").val();
		var sdp_vat = $("#dp_vat").val();
		var stotalamount = $("#totalamount").val();
		var stodate_dp_aftervat = $("#todate_dp_aftervat").val();
		var sremainsdp = $("#remainsdp").val();
		var newsdp_beforevat = sdp_beforevat.replace(/,/g, '');
		var newsdp_aftervat = sdp_aftervat.replace(/,/g, '');
		var newstotaltodate_dp = stotaltodate_dp.replace(/,/g, '');
		var newstotaltodate_perc_dp = stotaltodate_perc_dp.replace(/,/g, '');
		var newstotalamount = stotalamount.replace(/,/g, '');
		var newstodate_dp_aftervat = stodate_dp_aftervat.replace(/,/g, '');
		var newsremainsdp = sremainsdp.replace(/,/g, '');
		
    	//Untuk menghitung Amount DP Vat 
		totalamountdpvat = sdp_vat * newsdp_beforevat / 100;
		if (!isNaN(totalamountdpvat)) {
    	 	$("#amount_dpvat").val(totalamountdpvat);
    	}
    	else{
    		$("#amount_dpvat").val();
    	}

    	//Untuk menghitung Amount DP After Vat 
		totalamountdpaftervat = parseInt(newsdp_beforevat) + parseInt(totalamountdpvat) ;
		if (!isNaN(totalamountdpaftervat)) {
    	 	$("#dp_aftervat").val(totalamountdpaftervat);
			newsdp_aftervat = totalamountdpaftervat;
    	}
    	else{
    		$("#dp_aftervat").val();
    	}
		
		//Untuk menghitung Todate DP Aftervat
		totaltodate = (parseInt(newsdp_aftervat) + parseInt(newstotaltodate_dp)).toFixed(0);
		if (!isNaN(totaltodate)) {
    	 	$("#todate_dp_aftervat").val(totaltodate);
			newstodate_dp_aftervat = totaltodate;
    	}
    	else{
    		$("#todate_dp_aftervat").val();
    	}

    	//Untuk menghitung Percentage DP  
		totalpercent = ( (newstodate_dp_aftervat / newstotalamount)* 100).toFixed(0);
		if (!isNaN(totalpercent)) {
    	 	$("#percentage_dp").val(totalpercent);
    	 	newstotaltodate_perc_dp = totalpercent;
    	}
    	else{
    		$("#percentage_dp").val();
    	}
		
		//Untuk menghitung Todate Percentage DP Aftervat
		totaltodateperc = (parseInt(totalpercent) + parseInt(newstotaltodate_perc_dp) - parseInt(newstotaltodate_perc_dp)).toFixed(0);
		if (!isNaN(totaltodateperc)) {
    	 	$("#todate_percentage_dp").val(totaltodateperc); 
    	}
    	else{
    		$("#todate_percentage_dp").val();
    	}

		//Untuk menghitung Remain DP  
		totalremaindp = (parseInt(newstotalamount) - parseInt(newsremainsdp)).toFixed(0);
		if (!isNaN(totalremaindp)) {
    	 	$("#remaindp").val(totalremaindp);
    	}
    	else{
    		$("#remaindp").val();
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
var getMessage = function () {
	// build up the post string when passing variables to the server side page
	var fcba = document.getElementById("fcba").value;

	var productcode = document.getElementById("productcode").value;
	var PostStr = 'productcode='+productcode+'&fcba='+fcba;
	
	// use the generic function to make the request
	doAJAXCall('ajax/GetProductUom.asp', 'POST', PostStr, showMessageResponse);
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    
    // update the Div to show the result from the server
	document.getElementById("unitofmeasure").value = response;
}

// launched from button click 
var getLastNumberSalesContractno = function () {
	// build up the post string when passing variables to the server side page
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getLastSalesContractno.asp', 'POST', PostStr, showLastNumberSalesContractno);
}

// The function for handling the response from the server
var showLastNumberSalesContractno = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	//alert(s.lastnumber);
    // update the Div to show the result from the server
	document.getElementById("lastnumber").value = s.lastnumber;
}

// launched from button click 
var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbasalesinvoicedp.asp', 'POST', PostStr, showlistfcba);
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
var getSalesInvoiceDp = function () {
	// build up the post string when passing variables to the server side page
	var fcba = document.getElementById("fcba").value;
	var salescontractno = document.getElementById("salescontractno").value;
	var PostStr = 'salescontractno='+salescontractno+'&fcba='+fcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getsalesinvoicedp.asp', 'POST', PostStr, showMessageResponse);
	
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) {     
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	document.getElementById("totalamount").value = s.totalamount;
}

// launched from button click 
var gettodatedp = function () {
	// build up the post string when passing variables to the server side page
	var fcba = document.getElementById("fcba").value;
	var salescontractno = document.getElementById("salescontractno").value;
	var PostStr = 'salescontractno='+salescontractno+'&fcba='+fcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/gettodateamountsales.asp', 'POST', PostStr, showToDateDP);
	
}

// The function for handling the response from the server
var showToDateDP = function (oXML) {     
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	document.getElementById("totaltodate_dp").value = s.todate_dp_aftervat;
}

// launched from button click 
var getremaindp = function () {
	// build up the post string when passing variables to the server side page
	var fcba = document.getElementById("fcba").value;
	var salescontractno = document.getElementById("salescontractno").value; 
	var PostStr = 'salescontractno='+salescontractno+'&fcba='+fcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getremainssalesinvoicedp.asp', 'POST', PostStr, showremaindp);
	
}

// The function for handling the response from the server
var showremaindp = function (oXML) {     
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	document.getElementById("remainsdp").value = s.remainsdp;
}

// launched from button click 
var gettodateperdp = function () {
	// build up the post string when passing variables to the server side page
	var fcba = document.getElementById("fcba").value;
	var salescontractno = document.getElementById("salescontractno").value;
	var PostStr = 'salescontractno='+salescontractno+'&fcba='+fcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/gettodatepercentdp.asp', 'POST', PostStr, showtodateperdp);
	
}

// The function for handling the response from the server
var showtodateperdp = function (oXML) {     
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
	document.getElementById("totaltodate_perc_dp").value = s.todate_percentage_dp;
}; 

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
 <!-------------awal!------------->
 <%
 
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

invoicedpno = request("invoicedpno")
salescontractno = request("salescontractno")
transdate = request("transdate")
totalamount = request("totalamount")
dp_beforevat = request("dp_beforevat")
dp_aftervat = request("dp_aftervat")
dp_vat = request("dp_vat")
amount_dpvat = request("amount_dpvat")
todate_dp_aftervat= request("todate_dp_aftervat")
percentage_dp = request("percentage_dp")
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
txtlineno = Request("txtlineno")
valid_until = Request("valid_until")
todate_percentage_dp = Request("todate_percentage_dp")

Function CheckFieldCode(initialFieldCode, initialFcba, initialDate)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from dual " & _
					" where field_api.GetValid_Activation('" & initialFieldCode & "', '" & initialFcba & "', to_date('"& initialDate &"', 'mm/dd/yyyy')) = 'Y'"
	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckFieldCode = initialAvailable
	'Select count(1) from dual where field_api.GetValid_Activation('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')) = 'Y'
	'select to_char (field_api.GetValid_Plantingdate('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')),'yyyy') from dual
End Function



If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
		connstr.BeginTrans
		
		strFieldDetail = " SELECT COUNT(1) total FROM salesagreementdet " & _
                         " WHERE salescontractno = '" & salescontractno & "'" & _
                         "       AND fcBA = '" & fcba & "'"								
		'Response.Write strFieldDetail
		Set rsQty = Server.CreateObject("ADODB.Recordset")
		rsQty.Open strFieldDetail, connstr							
		if not rsQty.Eof then
			has_detail = cdbl (to_nol(rsQty("total")))
		End If
		
		if has_detail = 0 then
			q = " delete from salesagreementhdr " & _
				" where fcBA = '" & fcba & "'" & _
				" and salescontractno = '" & salescontractno & "'"
				'Response.Write q
			connStr.execute(q)
		else
			err.number = 1001
			err.description = " Sales Contract No """ & salescontractno & """ has detail, please delete all detail first!"
		End If
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			' strFieldDetail =	"select count(*) has_total" & _
								' "		from salesinvoicedp " & _
								' " 		where invoicedpno = '" & invoicedpno & "'" & _
								' "       and fcBA = '" & fcba & "'"								
			' Response.Write strFieldDetail
			' Set rsQty = Server.CreateObject("ADODB.Recordset")
			' rsQty.Open strFieldDetail, connstr							
			' if not rsQty.Eof then
				' has_total = cdbl (to_nol(rsQty("has_total")))
			' End If
			
			' If has_total > 0 Then 
				' err.number = 1001
				' err.description = "You Must Delete Data Detail First"
			' Else
				q = " delete from salesinvoicedp " & _
					" where fcBA = '" & fcba & "'" & _
					" and invoicedpno = '"& invoicedpno & "'" & _
					" and lineno = '"& txtlineno &"'"
					' Response.Write q
					connStr.execute(q)
			' End If
			cmdView	= "True"

	elseif Request("txtJenis")="New" then 
					
		If Len(Trim(salescontractno)) = 0 Then
			Err.Number = 1001
			Err.Description = "Sales Contract Harus dipilih, tidak boleh kosong"
		Else
			qLineMax = " select nvl(max(lineno),0) tempMaxLine" & _
					   " from salesinvoicedp" & _
					   " where salescontractno = '" & salescontractno & "'" & _
					   " and fcba = '" & fcba & "'" 			

			set rsqLineMax = Server.CreateObject("ADODB.Recordset")
			rsqLineMax.Open qLineMax, connstr
			if not rsqLineMax.eof then
				tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
			else
				tempMaxLine = 0
			end if 

			tempMaxLine = tempMaxLine + 1
			txtLineNoHist = tempMaxLine
			
			' quantitytolerance = cdbl(to_nol(quantity)) *  cdbl(to_nol(quantity)) 
			' amount = cdbl(to_nol(quantity)) * cdbl(to_nol(salesprices))
			' amountvat = cdbl(to_nol(vat)) * cdbl(to_nol(amount))/100
			' totalamountvat = amount +amountvat
			
			If cdbl(to_nol(percentage_dp)) > 100 Then
				Err.Number = 1001
				Err.Description = "Maksimal Percentage DP 100%"
			Else				
				q = " INSERT INTO salesinvoicedp" & _
					" (invoicedpno, transdate, salescontractno, totalamount, dp_beforevat, dp_vat, amount_dpvat, dp_aftervat, todate_dp_aftervat, percentage_dp, todate_percentage_dp, valid_until, lineno, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" '" & invoicedpno & "'" & _
					" , to_date('" & transdate & "','mm/dd/yyyy')" & _
					" , '" & salescontractno & "'" & _  
					" , replace('" & totalamount & "',',','')" & _
					" , replace('" & dp_beforevat & "',',','')" & _
					" , replace('" & dp_vat & "',',','')" & _
					" , replace('" & amount_dpvat & "',',','')" & _
					" , replace('" & dp_aftervat & "',',','')" & _
					" , replace('" & todate_dp_aftervat & "',',','')" & _
					" , replace('" & percentage_dp & "',',','')" & _
					" , replace('" & todate_percentage_dp & "',',','')" & _
					" , to_date('" & valid_until & "','mm/dd/yyyy')" & _ 
					" , '" & tempMaxLine & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & fcba & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"
					' Response.Write q
					connStr.execute(q)
			End If
		End If
	' End If
	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		If Len(Trim(salescontractno)) = 0 Then
			Err.Number = 1001
			Err.Description = "Sales Contract Harus dipilih, tidak boleh kosong"
		Else
			If cdbl(to_nol(percentage_dp)) > 100 Then
				Err.Number = 1001
				Err.Description = "Maksimal Percentage DP 100%"
			Else
				q =	" Update salesinvoicedp" & _
					" set salescontractno = '" & salescontractno & "'" & _
					"     , totalamount = replace('" & totalamount & "',',','')" & _
					"     , dp_beforevat = replace('" & dp_beforevat & "',',','')" & _
					"     , dp_vat = replace('" & dp_vat & "',',','')" & _
					"     , amount_dpvat = replace('" & amount_dpvat & "',',','')" & _
					"     , dp_aftervat = replace('" & dp_aftervat & "',',','')" & _
					"     , todate_dp_aftervat = replace('" & todate_dp_aftervat & "',',','')" & _
					"     , percentage_dp = replace('" & percentage_dp & "',',','')" & _
					"     , todate_percentage_dp = replace('" & todate_percentage_dp & "',',','')" & _
					"     , valid_until = to_date('" & valid_until & "','mm/dd/yyyy')" & _
					"     , fcentry = '" & Session("CodeName") & "'" & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"     , lastupdate = sysdate " & _
					"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					" 		where invoicedpno = '" & invoicedpno & "'" & _
					"       and lineno = '" & txtlineno & "'" & _
					"       and fcBA = '" & fcba & "'"
				' response.write q
				connStr.execute(q)
			End If		
		End If		
	End If			
End If
	' End If
' End If
cmdView	= "True"	
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

'untuk menampilkan datepicker default
If Len(Trim(transdate)) = 0 Then	transdate = textDate(date)
If Len(Trim(valid_until)) = 0 Then	valid_until = textDate(date)
If Len(Trim(startdate)) = 0 Then startdate = textDate(date)
If Len(Trim(enddate)) = 0 Then	enddate = textDate(date)



Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select sa.*, bu.fccompanycode" & _
							" from salesinvoicedp sa left join businessunit bu on bu.fccode = sa.fcba" & _
							" where sa.invoicedpno = '" & invoicedpno & "'"
						
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	transdate = rsSupplyingDeclaration("transdate")
	invoicedpno = rsSupplyingDeclaration("invoicedpno")
	fccompanycode = rsSupplyingDeclaration("fccompanycode")
	fcba = rsSupplyingDeclaration("fcba")
	selectedfcba = rsSupplyingDeclaration("fcba")
End if
' response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('costbook')"
			
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

 <!-------------ISI Supplying Declaration Form!------------->
 
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Sales Invoice DP</div><% end if %>
	<form name="frmEntry" method="post" action="salesinvoicedp.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Sales Invoice DP No</td>
			<td align="left">
			<input type="text" maxlength="50" name="invoicedpno" size="30" value="<%=invoicedpno%>" autocomplete="Off" >
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Creation Date</td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="transdate" value="<%=transdate%>">
			</td>
		</tr>
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();" class="chosen-select" style="width: 30%">
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
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<div id="fcbalist"> </div>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
		</tr>		
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;		
				
				<!-- HADIR -->
				<!-- <%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Nursery"
				dataSubmodule 	= "Culling"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a> -->
				<!-- END:HADIR -->
			</td>
		</tr>
		</table>
	</form>
<%
if cmdView = "True" and len(trim(invoicedpno)) > 0  Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " SELECT *" & _
		"  FROM salesinvoicedp" & _ 
        " WHERE invoicedpno = '"& invoicedpno &"' and fcba = '"& selectedfcba &"'" 

	' response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
	
	If Request("txtJenis")="Edit" Then
		q = " SELECT invoicedpno, totalamount FROM salesagreementhdr WHERE fcba = '"& fcba &"' AND salescontractno = '"& salescontractno &"'"  

		' response.Write q
		set rsrecurtrans = server.CreateObject("Adodb.recordset")
		rsrecurtrans.Open q, connstr

		if Not rsrecurtrans.eof then 
			totalamount= rsrecurtrans("totalamount")
		end if
		rsrecurtrans.Close
		set rsrecurtrans = Nothing
	End If
%>
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
	<form action="salesinvoicedp.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td align="left">Sales Contract No</td>
				<td align="left">
					<select id="salescontractno" name="salescontractno" onchange="javascript:getSalesInvoiceDp(); javascript:gettodatedp(); javascript:gettodateperdp(); hitungrumus(); javascript:getremaindp();" class="chosen-select" style="width: 30%">
						<option value=""></option>
						<%
						qlov = " SELECT salescontractno FROM salesagreementhdr WHERE fcba = '"& fcba &"'"
						Set rslov = Server.CreateObject("ADODB.Recordset")
						rslov.Open qlov, connstr 
						if Not rslov.eof then
							While Not rslov.EOF
						%>
						<option value="<%=rslov("salescontractno")%>" <%if salescontractno = rslov("salescontractno") then%>selected<%end if%>><%=rslov("salescontractno")%></option>
						<%
								rslov.MoveNext
							Wend
						end if
						%>
					</select>
					<font color="#FF0000">*</font>
				</td>
			</tr>
			<tr>
				<td align="left">Total Amount</td>
				<td align="left"><input type="text" size="30" maxlength="50" id="totalamount" name="totalamount" value="<%=totalamount%>" onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background-color: #eee;" readonly></td>
			</tr>
			<tr>
				<td align="left">Remain DP</td>
				<td align="left">
					<input type="hidden" name = "remainsdp" id ="remainsdp" value="<%=remainsdp%>" onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();" >
					<input type="text" size="30" maxlength="50" id="remaindp" name="remaindp" value="<%=remaindp%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background-color: #eee;" readonly>
				</td>
			</tr>
			<tr>
				<td align="left">DP Before VAT</td>
				<td align="left"><input type="text" size="30" maxlength="50" id="dp_beforevat" name="dp_beforevat" value="<%=dp_beforevat%>" onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" ></td>
			</tr>
			<tr>
				<td align="left">DP VAT</td>
				<td align="left"><input type="text" size="30" maxlength="50" id="dp_vat" name="dp_vat" value="<%=dp_vat%>" onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" ></td>
			</tr>
			<tr>
				<td align="left">Amount DP VAT</td>
				<td align="left"><input type="text" size="30" maxlength="50" id="amount_dpvat" name="amount_dpvat" value="<%=amount_dpvat%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background-color: #eee;" readonly></td>
			</tr>
			<tr>
				<td align="left">DP After VAT</td>
				<td align="left">
					<input type="text" size="30" maxlength="50" id="dp_aftervat" name="dp_aftervat" value="<%=dp_aftervat%>"  onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background-color: #eee;" readonly >
				</td>
			</tr>
			<tr>
				<td align="left">To Date Amount DP VAT</td>
				<td align="left">
					<input type="hidden" name = "totaltodate_dp" id ="totaltodate_dp" value="<%=totaltodate_dp%>" onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();" >
					<input type="text" size="30" maxlength="50" id="todate_dp_aftervat" name="todate_dp_aftervat" value="<%=todate_dp_aftervat%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background-color: #eee;" readonly>
				</td>
			</tr>
			<tr>
				<td align="left">Percentage DP</td>
				<td align="left"><input type="text" size="30" maxlength="50" id="percentage_dp" name="percentage_dp" value="<%=percentage_dp%>" style="background-color: #eee; text-align: right;" readonly> %</td>
			</tr>
			<tr>
				<td align="left">To Date Percentage DP</td>
				<td align="left">
					<input type="hidden" name = "totaltodate_perc_dp" id ="totaltodate_perc_dp" value="<%=totaltodate_perc_dp%>" onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();" >
					<input type="text" size="30" maxlength="50" id="todate_percentage_dp" name="todate_percentage_dp" value="<%=todate_percentage_dp%>" style="background-color: #eee; text-align: right;" readonly> %
				</td>
			</tr>
			<tr>
				<td align="left" width="20%">Due Date</td>
				<td align="left">
					<input class="datepicker" type="text" size="15" name="valid_until" value="<%=valid_until%>">
				</td>
			</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="invoicedpno" value="<%=invoicedpno%>">
					<input type="hidden" name="transdate" value="<%=transdate%>">
					<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
						<input type="hidden" name="txtlineno" value="<%=txtlineno%>">
						<input Type="Submit"  name="Add" value="Update">
					<%Else%>
						<input type="hidden" name="txtJenis" value="New">
						<input Type="Submit"  name="Add" value="Add">
					<%End if%>
					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">					
				</td>
			</tr>
		</table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="center">NO</td>
		<td colspan="2">&nbsp;</td>
		<td align="center">INVOICE DP NO</td>
		<td align="center">SALES CONTRACT NO</td>
		<td align="center">TOTAL AMOUNT</td>
		<td align="center">DP BEFORE VAT</td>
		<td align="center">DP VAT</td>
		<td align="center">AMOUNT DP VAT</td>
		<td align="center">DP AFTER VAT</td>
		<td align="center">TO DATE AMOUNT DP VAT</td>
		<td align="center">PERCENTAGE DP(%)</td>
		<td align="center">TO DATE PERCENTAGE DP(%)</td>
		<td align="center">VALID UNTIL</td>	
	</tr>
	</thead>
	<tbody>
   <%
	if Not rsVH.eof then
		nCountofField = rsVH.fields.count - 1
		no = 1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		<tr bgcolor="<%=cColor%>">
			<td align="center" valign="top"><%=NO%></td>
			
			
			<td>
				<a href="salesinvoicedp.asp?txtlineno=<%=rsVH("lineno")%>&salescontractno=<%=rsVH("salescontractno")%>&invoicedpno=<%=rsVH("invoicedpno")%>&totalamount=<%=rsVH("totalamount")%>&dp_beforevat=<%=rsVH("dp_beforevat")%>&dp_vat=<%=rsVH("dp_vat")%>&amount_dpvat=<%=rsVH("amount_dpvat")%>&dp_aftervat=<%=rsVH("dp_aftervat")%>&todate_dp_aftervat=<%=rsVH("todate_dp_aftervat")%>&percentage_dp=<%=rsVH("percentage_dp")%>&todate_percentage_dp=<%=rsVH("todate_percentage_dp")%>&valid_until=<%=rsVH("valid_until")%>&fcba=<%=rsVH("fcba")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
			</td>
			<td>
				<a href="salesinvoicedp.asp?txtlineno=<%=rsVH("lineno")%>&salescontractno=<%=rsVH("salescontractno")%>&invoicedpno=<%=rsVH("invoicedpno")%>&totalamount=<%=rsVH("totalamount")%>&dp_beforevat=<%=rsVH("dp_beforevat")%>&dp_vat=<%=rsVH("dp_vat")%>&amount_dpvat=<%=rsVH("amount_dpvat")%>&dp_aftervat=<%=rsVH("dp_aftervat")%>&todate_dp_aftervat=<%=rsVH("todate_dp_aftervat")%>&percentage_dp=<%=rsVH("percentage_dp")%>&fcba=<%=rsVH("fcba")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
			    <td align="center" title="<%=rsVH("invoicedpno").Name%>" valign="top"><%=rsVH("invoicedpno")%></td>
			    <td align="center" title="<%=rsVH("salescontractno").Name%>" valign="top"><%=rsVH("salescontractno")%></td>
				<td align="right" title="<%=rsVH("totalamount").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("totalamount"))),0)%></td>
				<td align="right" title="<%=rsVH("dp_beforevat").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("dp_beforevat"))),0)%></td>
				<td align="right" title="<%=rsVH("dp_vat").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("dp_vat"))),0)%></td>
			    <td align="right" title="<%=rsVH("amount_dpvat").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("amount_dpvat"))),0)%></td>
			    <td align="right" title="<%=rsVH("dp_aftervat").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("dp_aftervat"))),0)%></td>
			    <td align="right" title="<%=rsVH("todate_dp_aftervat").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("todate_dp_aftervat"))),0)%></td>
			    <td align="right" title="<%=rsVH("percentage_dp").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("percentage_dp"))),0)%></td>
			    <td align="right" title="<%=rsVH("todate_percentage_dp").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("todate_percentage_dp"))),0)%></td>
			    <td align="right" title="<%=rsVH("valid_until").Name%>" valign="top"><%=rsVH("valid_until")%></td>
		</tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </tbody>
   </TABLE>
   </div>
   </form>
<%End if%>
 
  <!-------------akhir!------------->
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

function getselectedcustomer() {
	var scustomercode = document.getElementById("selectedcustomercode");
	scustomercode.value = document.getElementById("customercode").value;
}

function getselecteddeliverytype() {
	var sdeliverytype = document.getElementById("selecteddeliverytype");
	sdeliverytype.value = document.getElementById("deliverytype").value;
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
	hitungrumus();
	getremaindp();
	getSalesInvoiceDp();
	gettodatedp();
	gettodateperdp(); 
	getselectedfcba();
});
</script>
</body>
</html>
