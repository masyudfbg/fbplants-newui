<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ePlaS - Electronic Plantation Solutions</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
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
var getlistbank = function () {
	// build up the post string when passing variables to the server side page
	
	var businessunit = document.getElementById("businessunit").value; 
	var PostStr = 'businessunit='+businessunit;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getListBankVoucher.asp', 'POST', PostStr, showlistbank);
}

var getlistjournal = function () {
	// build up the post string when passing variables to the server side page
	
	var businessunit = document.getElementById("businessunit").value; 
	var PostStr = 'businessunit='+businessunit;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getListJournalVoucher.asp', 'POST', PostStr, showlistbank);
}

// The function for handling the response from the server
var showlistbank = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("banklist").innerHTML = "";
	document.getElementById("banklist").innerHTML = response;

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
};

</script>
<%
' Tampilan baru (css/fbp-shell.css + css/fbp-forms.css) aktif untuk semua modul.
' Tambahkan ?oldui=1 di URL untuk melihat tampilan lama (pembanding / darurat).
' useCustomForm: modul yang sudah punya desain sendiri di Select_Form_Show.asp,
' jadi tidak dibungkus kartu .lf.
Dim useNewLayout, useCustomForm
useNewLayout = (Request("oldui") <> "1")
useCustomForm = (Request("txtTable") = "search_stock" Or LCase(Request("txtTable")) = "mutasistockharian")

' Judul halaman dari nama tabel, mis. "PurchaseRequestDetails" -> "Purchase Request Details"
Function fbpPageTitle(s)
	Dim re, t
	t = Replace(Replace(s & "", "_", " "), "-", " ")
	Set re = New RegExp
	re.Global = True
	re.Pattern = "([a-z])([A-Z])"
	t = re.Replace(t, "$1 $2")
	re.Pattern = "^\d+\s*"
	t = re.Replace(t, "")
	fbpPageTitle = Trim(t)
End Function

if useNewLayout then %>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if useNewLayout then %> class="dash page-show"<% end if %>>
<!--#include file="top.jsp"-->

<%
Dim txtTable,txtModule
txtTable = Request("txtTable")
txtModule = Request("txtModule")
%>
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <% if Session("fcBa") <> "" then %>
     <!--#include file="menu.asp"-->
 <% end if %>
</div>
   <div class="judul_head"  valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Moduled : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;"> 

    <%' -----------------ATAS ISI
	%> 

	<%if request("txtModule")="Data Master" or request("txtModule")="Data Setup" Then%>
	     <form action="PrintData.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%Elseif request("txtModule")="Administrator" Then%>
	     <form action="PrintData.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
	<%Elseif request("txtModule")="Report" Then%>

    	<%	
			if request("txttable")= "AllReport" then 
				response.redirect "/Report/Dashboard.asp"
			else
				response.redirect "/" & request("txttable") & ".asp" 
			end if
		
		%>
	<%
    elseif request("txtTable")="Financial_Year_Closing" and request("txtModule")="General_Ledger" then
	        response.redirect "/Financial_Year_Closing.asp"
	elseif request("txttable")="Account_Period" Then
			response.redirect "/account_period.asp"
	elseif request("txttable")="MasterMenu" Then
			response.redirect "/MasterMenu.asp"
	elseif request("txttable")="HistLog" Then
			response.redirect "/HistLog.asp"
	elseif request("txttable")="LogActivityUser" Then
			response.redirect "/log_activity.asp"
	Else
		if Request("txttable") = "02-paymentvoucher" or Request("txttable") = "01-receivevoucher"  then
			transtype = ""
			banktype = ""
			
			transtype = Replace(Request("txttable") , "voucher", "")
			%>
		 <form action="p_paymentvoucherdetail.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
			<input type="hidden" name="transtype" value="<%=transtype%>">
			<input type="hidden" name="banktype" value="<%=banktype%>">
		<%elseif Request("txttable") = "SpbHarvesting" then%>
		<form action="p_Harvesting.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
		<%else%>
	     <form action="p_<%=request("txtTable")%>.asp" method="post" target="_blank" style="screenY:0,directories:0,fullscreen:0,location:0,menubar:0,scrollbars:1,status:0,toolbar:1" name="frmEntry">
		<%End if%>
	<%End if%>
	<% if useNewLayout then %>
	<div class="page-form">
	<% if not useCustomForm then %>
	<div class="page-title">
		<div class="eyebrow"><%=Server.HTMLEncode(Request("txtModule") & "")%></div>
		<h2><%=Server.HTMLEncode(fbpPageTitle(Request("txtTable")))%></h2>
	</div>
	<div class="lf">
	<% end if %>
	<% else %>
	<table style="font:Arial, Helvetica, sans-serif" style="font-size:9px" class="sample" cellspacing="0" width="100%"> 
		<tr>
	<% end if %>
            <%
			Dim txtLink
			if txtModule="Data Master" then
			    if txtTable = "ExtendStock" then
                                txtLink = "ExtendStock"
                            else
			        txtLink = "TableMaster"
                            end if	
			elseif (txtModule="Data Setup") and txtTable = "Process" then
			    txtLink = txtTable	   
			elseif (txtModule="Data Setup") and txtTable <> "AccessRights" then
			   txtLink = "TableStatistik"					   
			elseif txtTable = "Roles" then
			   txtLink = "TableMaster"				
			elseif txtTable = "Users" then
			   txtLink = "TableStatistik"				
			else
			   txtLink = txtTable
			end if
			'response.write txtLink
			'response.write txtTable
			%>
			<!-- Tampilan select form -->

            <!--#include file="Select_Form_Show.asp"-->

			<!-- Akhir Tampilan Select form -->
		<% if useNewLayout then %>
		<% if not useCustomForm then %></div><% end if %>
		</div>
		<% else %>
			</a></td>
		</tr>
		<% end if %>
		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		</form>
</div>
		
	<% if not useNewLayout then %></table><% end if %>
    </div></div>
<%'------------------ BAWAH ISI%>

  <!--#include file="bottom.jsp"-->
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
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
//document.ready(function() {
//	getlistbank();
	
//});
</script>
</body>
</html>