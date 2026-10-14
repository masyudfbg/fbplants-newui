<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<title>ePlaS | Fixed Asset Adjustment</title>
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

function CheckIsActive()
{
 var cIsActive = document.getElementById('ISACTIVE');
 if (cIsActive.checked != true )
	{
		alert("Is Active Belum di Click");
	}
}

function calcParam()
{
	var cIsActive = document.getElementById('ISACTIVE').checked;
	if(cIsActive)
	{
		document.getElementById("hide_ISACTIVE").value ="1";
	}else
	{
		document.getElementById("hide_ISACTIVE").value ="0";
	}
	
	var cIsBlock = document.getElementById('ISBLOCK').checked;
	if(cIsBlock)
	{
		document.getElementById("hide_ISBLOCK").value ="1";
	}else
	{
		document.getElementById("hide_ISBLOCK").value ="0";
	}
		
}

function populateSelect(target, min, max){
	if (!target){
		return false;
	}
	else {
		var min = min || 0,
			max = max || min + 0;

		select = document.getElementById(target);

		for (var i = min; i<=max; i++){
			var opt = document.createElement('option');
			opt.innerHTML = "";
			opt.value = i;
			opt.innerHTML = i;
			select.appendChild(opt);
		}
	}
}

function hanyaAngka(evt) {
  // var charCode = (evt.which) ? evt.which : event.keyCode
  //  if (charCode > 31 && (charCode < 48 || charCode > 57))

  //   return false;
  // return true;
  evt.value = evt.value.replace(/[^0-9,.-]/g,"");
}

function hitungrumus(){
    	var total = document.getElementById("total").value;
    	var totalyearly = document.getElementById("totalyearly").value;
    	var ar_others = document.getElementById("ar_others").value;
    	// var acquisitionvalue_type = document.getElementById("acquisitionvalue_type").value;
    	// var yearlydepreciation = document.getElementById("yearlydepreciation").value;
		// alert(acquisitionvalue_type);
    	// var totalvalue = document.getElementById("totalvalue").value;
    	// var myStr = 'this,is,a,test';
		var newtotal = total.replace(/,/g, '');
		var newtotalyearly = totalyearly.replace(/,/g, '');
		var newar_others = ar_others.replace(/,/g, '');

    	// var total = parseInt(newacquisitionvalue) + parseInt(newadditionalvalue);
    	// if (!isNaN(total)) {
    	//  	document.getElementById('totalvalue').value = total;
    	// }
    	// else{
    	// 	document.getElementById('totalvalue').value = 0;
    	// }

    	
		if (fctype = 'SALES') {
			var totalgainloss =  newtotal -  newtotalyearly - newar_others;
			// alert(newar_others);
			if (!isNaN(totalgainloss)) {
			 	document.getElementById('gain_loss').value = totalgainloss;
			}
			else {
				document.getElementById('gain_loss').value = 0;
			}    	
		} 
    }

</script>
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

var getlistline = function () {
	// build up the post string when passing variables to the server side page
	
	var fixassetcode = document.getElementById("fixassetcode").value;
	var lineno = document.getElementById("selectedline").value;
	var PostStr = 'fixassetcode='+fixassetcode+'&lineno='+selectedline;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlinenofa.asp', 'POST', PostStr, showlistline);
}

// The function for handling the response from the server
var showlistline = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("linelist").innerHTML = "";
	document.getElementById("linelist").innerHTML = response;

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
var getAssetWriteOffInfo = function () {
	// build up the post string when passing variables to the server side page
	
	var fixassetcode = document.getElementById("fixassetcode").value;
	var lineno = document.getElementById("lineno").value;
	var PostStr = 'fixassetcode='+fixassetcode+'&lineno='+lineno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfixedassetwriteoffinfo.asp', 'POST', PostStr, showAssetWriteOffInfo);
}


// The function for handling the response from the server
var showAssetWriteOffInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.salescontractno);
    // update the Div to show the result from the server
	document.getElementById("fixassetname").value = s.fixassetname;
	document.getElementById("total").value = s.total;
	document.getElementById("total").value = "";
	document.getElementById("totalyearly").value = s.totalyearly;
	document.getElementById("totalyearly").value = "";
	
	// document.getElementById("ar_others").value = s.ar_others;
	// document.getElementById("ar_others").value = "";

	var total = s.total;
	var totalyearly = s.totalyearly;
	var total_ar_others = total - totalyearly;
	var fctype = document.getElementById("fctype").value ;
	
	document.getElementById("total").value = s.total;
	document.getElementById("totalyearly").value = s.totalyearly;
	

	if (fctype != 'SALES') {
		document.getElementById("ar_others").value = total_ar_others;
	} else{
		document.getElementById("ar_others").value = 0; 
	}  
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
searchdata = request("searchdata")
fccode = request("fccode")
fctype = request("fctype")
transdate = request("transdate")

fixassetcode = request("fixassetcode")
fixassetname = request("fixassetname")
targettype_depreciation = request("targettype_depreciation")
targetcode_depreciation = request("targetcode_depreciation")
jobcode_depreciation = request("jobcode_depreciation")
accumulated_depreciation = request("accumulated_depreciation")
targettype_others = request("targettype_others")
targetcode_others = request("targetcode_others")
jobcode_others = request("jobcode_others")
ar_others = request("ar_others")
targettype_gain = request("targettype_gain")
targetcode_gain = request("targetcode_gain")
jobcode_gain = request("jobcode_gain")
gain_loss = request("gain_loss")
remarks = request("remarks")
description = request("description")
fixassetamount = request("fixassetamount")
lineno = request("lineno")


if ""&transdate = "" then transdate = TextDate(date)

if Request("txtJenis")="New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
		connstr.BeginTrans
		
		strFieldDetail = " SELECT COUNT(1) total FROM fixedasset_writeoff_det " & _
                         " WHERE fccode = '" & fccode & "'" & _
                         "       AND fcBA = '" & Session("fcba") & "'"								
		' Response.Write strFieldDetail
		Set rsQty = Server.CreateObject("ADODB.Recordset")
		rsQty.Open strFieldDetail, connstr							
		if not rsQty.Eof then
			has_detail = cdbl (to_nol(rsQty("total")))
		End If
		
		if has_detail = 0 then
			q = " DELETE FROM fixedasset_writeoff WHERE fccode = '" & fccode & "' AND fcba = '"& Session("fcBA") &"'"
				'Response.Write q
			connStr.execute(q)
		else
			err.number = 1001
			err.description = " Fixed Asset Write Off Trans """ & fcCode & """ has detail, please delete all detail first!"
		End If
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
        cmdView	= "True"


	elseif Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM fixedasset_writeoff_det WHERE fixassetcode = '"& fixassetcode &"' AND fccode = '"& fccode &"' AND fcba = '"& Session("fcBA") &"' AND lineno = '"& lineno &"'"
		' response.write q
		connStr.execute(q)

		q = " UPDATE fixedasset" & _
            " SET   isactive = '1'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , lastupdate = SYSDATE" & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
			" WHERE fccode = '" & fixassetcode & "' AND fcba = '"& Session("fcBA") &"'" 
		' response.write q
		connStr.execute(q)

		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		hasdata = 0
		q = " SELECT COUNT(1) hasdata FROM fixedasset_writeoff_det WHERE fccode = '"& fccode &"' and fcba = '"& Session("fcba") &"'"
		' response.write q
		Set rscekpk = Server.CreateObject("ADODB.Recordset")
		rscekpk.Open q, connstr
		
		if Not rscekpk.eof then
			hasdata = cdbl(to_nol(rscekpk("hasdata")))
		else
			hasdata = 0
		end if
		
		if hasdata = 1 then
			err.number = 1001
			err.description = "Fixed Asset Code : " & fixassetcode & " , already exist"
		else
			
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select fccode from fixedasset_writeoff where fccode = '" & fccode & "' and fcba ='" & fcba & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then
						q = " INSERT INTO fixedasset_writeoff ( fccode" & _
							"                           , fctype" & _
							"                           , transdate" & _
							"                           , description" & _
							"                           , fcentry" & _
							"                           , fcedit" & _
							"                           , fcip" & _
							"                           , fcba" & _
							"                           , lastupdate" & _
							"                           , lasttime)" & _
							" VALUES ( '" & fccode & "'" & _
							"          , '" & fctype & "'" & _
							"          , TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
							"          , '" & description & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & Session("fcba") & "'" & _
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

						' response.write q
						connStr.execute(q)
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing

			q = " INSERT INTO fixedasset_writeoff_det ( fccode " & _
				"							, fixassetcode" & _
				"                           , fixassetname" & _
				"                           , fixassetamount" & _
				"                           , targettype_depreciation" & _
				"                           , targetcode_depreciation" & _
				"                           , jobcode_depreciation" & _
				"                           , accumulated_depreciation" & _
				"                           , targettype_others" & _
				"                           , targetcode_others" & _
				"                           , jobcode_others" & _
				"                           , ar_others" & _
				"                           , targettype_gain" & _
				"                           , targetcode_gain" & _
				"                           , jobcode_gain" & _
				"                           , gain_loss" & _
				"                           , remarks" & _
				"                           , lineno" & _
				"                           , fcentry" & _
				"                           , fcedit" & _
				"                           , fcip" & _
				"                           , fcba" & _
				"                           , lastupdate" & _
				"                           , lasttime)" & _
				" VALUES ( '" & fccode & "'" & _
				"          , '" & fixassetcode & "'" & _
				"          , '" & fixassetname & "'" & _
				"          ,replace('" & fixassetamount & "',',','')" & _
				"          , '" & targettype_depreciation & "'" & _
				"          , '" & targetcode_depreciation & "'" & _
				"          , '" & jobcode_depreciation & "'" & _
				"          ,replace('" & accumulated_depreciation & "',',','')" & _
				"          , '" & targettype_others & "'" & _
				"          , '" & targetcode_others & "'" & _
				"          , '" & jobcode_others & "'" & _
				"          ,replace('" & ar_others & "',',','')" & _
				"          , '" & targettype_gain & "'" & _
				"          , '" & targetcode_gain & "'" & _
				"          , '" & jobcode_gain & "'" & _
				"          ,replace('" & gain_loss & "',',','')" & _
				"          , '" & remarks & "'" & _
				"          , '" & lineno & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , '" & Session("fcba") & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

			' response.write q
			connStr.execute(q)

			q = " UPDATE fixedasset" & _
	            " SET   isactive = '0'" & _
				"     , valid_until = to_date( ( select trunc(transdate + 1) from fixedasset_writeoff where fccode = '"& fcCode &"' ),'mm/dd/yyyy')" & _
				"     , fcedit = '" & Session("CodeName") & "'" & _
				"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"     , lastupdate = SYSDATE" & _
				"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
				" WHERE fccode = '" & fixassetcode & "' AND fcba = '"& Session("fcBA") &"'" 
			' response.write q
			connStr.execute(q)
		end if
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q = " UPDATE fixedasset_writeoff_det" & _
            " SET   fixassetname = '" & fixassetname & "'" & _
			"     , fixassetamount = replace('" & fixassetamount & "',',','')" & _
			"     , jobcode_depreciation = '" & jobcode_depreciation & "'" & _
			"     , targettype_depreciation = '" & targettype_depreciation & "'" & _
			"     , targetcode_depreciation = '" & targetcode_depreciation & "'" & _
			"     , accumulated_depreciation = replace('" & accumulated_depreciation & "',',','')" & _
			"     , targettype_others = '" & targettype_others & "'" & _
			"     , targetcode_others = '" & targetcode_others & "'" & _
			"     , jobcode_others = '" & jobcode_others & "'" & _
			"     , ar_others = replace('" & ar_others & "',',','')" & _
			"     , targettype_gain = '" & targettype_gain & "'" & _
			"     , targetcode_gain = '" & targetcode_gain & "'" & _
			"     , jobcode_gain = '" & jobcode_gain & "'" & _
			"     , gain_loss = replace('" & gain_loss & "',',','')" & _
			"     , remarks = '" & remarks & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , lastupdate = SYSDATE" & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
			" WHERE fixassetcode = '" & fixassetcode & "' AND fccode = '"& fccode &"' AND fcba = '"& Session("fcBA") &"' AND lineno = '"& lineno &"'" 
		' response.write q
		connStr.execute(q)	
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")


If Len(Trim(date_transaction)) = 0 Then	date_transaction = textDate(date)

Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select * from fixedasset_writeoff where fcCode = '" & fccode & "' and fcba = '"& Session("fcBA") &"'"
						
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	transdate = rsSupplyingDeclaration("transdate")
	fcCode = rsSupplyingDeclaration("fcCode")
	fctype = rsSupplyingDeclaration("fctype")
	Description = rsSupplyingDeclaration("Description")
End if
' response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

 '-------------ISI Payment Form
  %>
  
<form name="frmEntry" method="post" action="fixedasset_writeoff.asp">
	<table class="data" width="100%">
	<tr>
		<td align="left">Fixed Asset Write Off</td>
			<td align="left">
				<input type="text" size="50" id="fcCode" name="fcCode" value="<%=fcCode%>" autocomplete="off">
				<font color="#FF0000">&nbsp;*</font>
			</td>
	</tr>
	<tr>
		<td align="left">Fixed Asset Write Off Type</td>
			<td align="left">
				<select id="fctype" name="fctype" onchange="GetDepreciationRange();" class="chosen-select" style="width: 40%;">
					<option></option>
					<option value="SALES" <%if fctype = "SALES" then%>selected<%end if%>>SALES</option>
					<option value="DISPOSAL" <%if fctype = "DISPOSAL" then%>selected<%end if%>>DISPOSAL</option>
				</select>
			</td>
	</tr>
	<tr>
		<td align="left">Creation Date</td>
		<td align="left">
			<input class="datepicker" type="text" maxlength="10" name="transdate" size="10" value="<%=transdate%>" autocomplete="off">
			<font color="#FF0000">&nbsp;*</font>
		</td>
	</tr>
	<tr>
		<td align="left">Description</td>
		<td align="left">
			<textarea name="description"><%=description%></textarea>
		</td>
	</tr>	
<tr>
	<td colspan="2"  align="left">
	<input Type="Hidden" Value="True" Name="cmdView">
	<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
	<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
	<input Type="Submit"  name="Add" value="Select">&nbsp;
	<a href="fixedasset_writeoff.asp?fcCode=<%=fccode%>&txtJenis=HeaderDelete" onClick="return confirm('Are you sure to delete');">
		<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0">
	</a>
	<a href="Printfixedasset_writeoff.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
	</td>
</tr>
</table>  
</form>
<%
if cmdView = "True" AND len(trim(fcCode)) > 0 Then
	q = " SELECT fwo.fccode," & _
		"       fwo.fctype," & _
		"       fwo.transdate," & _
		"       fwo.description," & _
		"       fwd.lineno," & _
		"       fwd.fixassetcode," & _
		"       fwd.fixassetname," & _
		"       fwd.fixassetamount," & _
		"       fwd.targettype_depreciation," & _
		"       fwd.targetcode_depreciation," & _
		"       fwd.jobcode_depreciation," & _
		"       fwd.accumulated_depreciation," & _
		"       fwd.targettype_others," & _
		"       fwd.targetcode_others," & _
		"       fwd.jobcode_others," & _
		"       fwd.ar_others," & _
		"       fwd.targettype_gain," & _
		"       fwd.targetcode_gain," & _
		"       fwd.jobcode_gain," & _
		"       fwd.gain_loss," & _
		"       fwd.remarks" & _
		"  FROM fixedasset_writeoff fwo" & _
		"       INNER JOIN fixedasset_writeoff_det fwd" & _
		"          ON (fwo.fccode = fwd.fccode AND fwo.fcba = fwd.fcba)" & _
		" WHERE     fwo.fccode = '"& fccode &"'" & _
		"       AND fwo.fcba = '"& Session("fcba") &"'" 

	' if ""&searchdata <> "" then q = q & " AND (UPPER(fa.fccode) LIKE UPPER('%"&searchdata&"%') OR UPPER(fa.fcname) LIKE ('%"&searchdata&"%'))"
	' if Request("txtJenis")="Edit" then q = q & " AND fa.fccode = '" & fccode & "'"
	
	' response.write q
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.Open q, connstr
	
	if Request("txtJenis")="Edit" then
		if Not rsGAD.eof then
			fcname = rsGAD("fcCode")
			fctype = rsGAD("fctype")
			lineno = rsGAD("lineno")
			description = rsGAD("description")
			transdate = rsGAD("transdate")
			fixassetcode = rsGAD("fixassetcode")
			fixassetname = rsGAD("fixassetname")
			fixassetamount = rsGAD("fixassetamount")
			targettype_depreciation = rsGAD("targettype_depreciation")
			targetcode_depreciation = rsGAD("targetcode_depreciation")
			jobcode_depreciation = rsGAD("jobcode_depreciation")
			accumulated_depreciation = rsGAD("accumulated_depreciation")
			targettype_others = rsGAD("targettype_others")
			targetcode_others = rsGAD("targetcode_others")
			jobcode_others = rsGAD("jobcode_others")
			ar_others = rsGAD("ar_others")
			targettype_gain = rsGAD("targettype_gain")
			targetcode_gain = rsGAD("targetcode_gain")
			jobcode_gain = rsGAD("jobcode_gain")
			gain_loss = rsGAD("gain_loss")
			remarks = rsGAD("remarks")
			fixassetamount = rsGAD("fixassetamount")
		end if
	end if

%>
   
   <!-- Form Entry And Edit --> 
   
   <form action="fixedasset_writeoff.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
			<tr>
				<%if Request("txtJenis")="Edit" then%>
				<td align="left">Fixed Asset Code</td>
				<td align="left">
        			<input type="text" size="50" id="fixassetcode" name="fixassetcode" value="<%=fixassetcode%>" readonly style="background-color:#ccc;">
        		</td>
        		<%Else%>
				<td align="left">Fixed Asset Code</td>
				<td align="left">
					<select id="fixassetcode" name="fixassetcode" onchange="getlistline();" class="chosen-select" style="width: 70%;">
						<option value=""></option>
						<%
						qlov = 	"  SELECT fat.fccode, fat.fcname" & _
								"  FROM fixedasset fat" & _
								"       LEFT JOIN user_fcba_tab uft ON (fat.fcba = uft.fcba)" & _
								" WHERE uft.fcusercode = '"& Session("CodeName") &"' AND fat.fcba = '"& Session("fcba") &"'" 

						' response.write qlov
						Set rslov = Server.CreateObject("ADODB.Recordset")
						rslov.Open qlov, connstr 
						if Not rslov.eof then
							While Not rslov.EOF
						%>
						<option value="<%=rslov("fccode")%>" <%if cstr(""&fixassetcode) = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
						<%
								rslov.MoveNext
							Wend
						end if
						%>
					</select>
				</td>
				<%End If%>
				<td align="left">Jobcode A/R Others </td>
					<td align="left">
						<input type="text" size="25" id="jobcode_others" name="jobcode_others" value="<%=jobcode_others%>">
						<font color="#FF0000">&nbsp;*&nbsp;</font>
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>
					</td>
			</tr>
			<tr>
				<%if Request("txtJenis")="Edit" then%>
				<td align="left">Line No</td>
				<td align="left">
        			<input type="text" size="25" id="lineno" name="lineno" value="<%=lineno%>" readonly style="background-color:#ccc;">
        		</td>
        		<%Else%>
				<td align="left">Line No</td>
				<td align="left">
					<div id="linelist" onchange="hitungrumus();" onblur="hitungrumus();" onkeypress="hitungrumus();" onkeyup="hitungrumus();"></div>
				</td>
					<input type="hidden" size="25" id="selectedline" name="selectedline" value="<%=selectedline%>" >
				<%End if%>
				<td align="left">Target Type A/R Others </td>
					<td align="left">
						<input type="text" size="25" id="targettype_others" name="targettype_others" value="<%=targettype_others%>" >
						<font color="#FF0000">&nbsp;*&nbsp;</font>
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a>
					</td>
			</tr>
			<tr>
				<td align="left">Fixed Asset Name</td>
					<td align="left">
						<input type="text" size="50" id="fixassetname" name="fixassetname" value="<%=fixassetname%>" readonly="true"  style="background-color:#ccc;">
					</td>
				</td>
				<td align="left">Target Code A/R Others</td>
					<td align="left">
						<input type="text" size="25" id="targetcode_others" name="targetcode_others" value="<%=targetcode_others%>" >
					</td>
			</tr>
			<tr>
				<td align="left">Fixed Asset Amount</td>
					<td align="left">
						<input type="text" size="25" id="total" name="fixassetamount" value="<%=fixassetamount%>" onchange="hitungrumus();" onblur="hitungrumus();" onkeypress="hitungrumus();" onkeyup="hitungrumus();" style="background:#ccc; text-align: right;" readonly >
					</td>
				<td align="left">A/R Others / Biaya Amount</td>
					<td align="left">
						<input type="text" size="25" id="ar_others" name="ar_others" onchange="hitungrumus();" onblur="hitungrumus();" onkeypress="hitungrumus();" onkeyup="hitungrumus();" value="<%=ar_others%>" style="text-align: right;" autocomplete="off">
					</td>
			<tr>
				<td align="left">Jobcode Acc. Depreciation</td>
					<td align="left">
						<input type="text" size="25" id="jobcode_depreciation" name="jobcode_depreciation" value="<%=jobcode_depreciation%>" >
						<font color="#FF0000">&nbsp;*</font>
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>
					</td>
				<%if fctype = "SALES" then%>
					<td align="left">Jobcode Gain/Loss</td>
						<td align="left">
							<input type="text" size="25" id="jobcode_gain" name="jobcode_gain" value="<%=jobcode_gain%>">
							<font color="#FF0000">&nbsp;*&nbsp;</font>
							<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>
					</td>
				<%Else%>
					<td align="left">Jobcode Gain/Loss</td>
						<td align="left">
							<input type="text" size="25" id="jobcode_gain" name="jobcode_gain" value="<%=jobcode_gain%>" style="background-color:#ccc;"  readonly>
							<font color="#FF0000">&nbsp;*&nbsp;</font>
							<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>
					</td>
				<%End If%>
			</tr>
			<tr>
				<td align="left">Target Type Acc. Depreciation</td>
					<td align="left">
						<input type="text" size="25" id="targettype_depreciation" name="targettype_depreciation" value="<%=targettype_depreciation%>">
						<font color="#FF0000">&nbsp;*&nbsp;</font>
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a>
					</td>
				<%if fctype = "SALES" then%>
					<td align="left">Target Type Gain/Loss</td>
						<td align="left">
							<input type="text" size="25" id="targettype_gain" name="targettype_gain" value="<%=targettype_gain%>" >
							<font color="#FF0000">&nbsp;*&nbsp;</font>
							<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a>
						</td>
				<%Else%>
					<td align="left">Target Type Gain/Loss</td>
						<td align="left">
							<input type="text" size="25" id="targettype_gain" name="targettype_gain" value="<%=targettype_gain%>" style="background-color:#ccc;"  readonly >
							<font color="#FF0000">&nbsp;*&nbsp;</font>
							<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a>
						</td>
				<%End If%>
			</tr>
			<tr>
				<td align="left">Target Code Acc. Depreciation</td>
					<td align="left">
						<input type="text" size="25" id="targetcode_depreciation" name="targetcode_depreciation" value="<%=targetcode_depreciation%>">
					</td>
				<%if fctype = "SALES" then%>
					<td align="left">Target Code Gain/Loss</td>
						<td align="left">
							<input type="text" size="25" id="targetcode_gain" name="targetcode_gain" value="<%=targetcode_gain%>" >
					</td>
				<%Else%>
					<td align="left">Target Code Gain/Loss</td>
						<td align="left">
							<input type="text" size="25" id="targetcode_gain" name="targetcode_gain" value="<%=targetcode_gain%>" style="background-color:#ccc;"  readonly>
					</td>
				<%End If%>
			</tr>
			<tr>
				<td align="left">Acc. Depreciation Fixed Asset Amount</td>
					<td align="left">
						<input type="text" size="25" id="totalyearly" name="accumulated_depreciation" value="<%=accumulated_depreciation%>" onchange="hitungrumus();" onblur="hitungrumus();" onkeypress="hitungrumus();" onkeyup="hitungrumus();" style="background:#ccc; text-align: right;" readonly >
					</td>
				<%if fctype = "SALES" then%>
				<td align="left">Gain/Loss Amount</td>
					<td align="left">
						<input type="text" size="25" id="gain_loss" name="gain_loss" value="<%=gain_loss%>" style="text-align: right; background-color:#ccc;" autocomplete="off" readonly>
					</td>
				</td>
				<%else%>
				<td align="left">Gain/Loss Amount</td>
					<td align="left">
						<input type="text" size="25" id="gain_loss" name="gain_loss" value="<%=gain_loss%>" readonly="true"  style="background-color:#ccc; text-align: right;" >
					</td>
				</td>
				<%end if%>
			</tr>
			<td></td>
			<td></td>
			<td align="left">Remarks</td>
					<td align="left">
						<textarea name="remarks" cols="50"></textarea>
					</td>
				</td>
		<tr>
			<td colspan="2" align="left">
				<input type="Hidden" name="searchdata" value="<%=searchdata%>">
			  	<input Type="Hidden" Value="<%=request("transdate")%>" Name="transdate">
				<input Type="Hidden" Value="<%=request("fcCode")%>" Name="fcCode">
				<input Type="Hidden" Value="<%=request("fctype")%>" Name="fctype">
				<input Type="Hidden" Value="<%=request("description")%>" Name="description">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				  <%IF Request("txtJenis")="Edit" then%>
					   <input type="hidden" name="txtJenis" value="Edit">
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
   </form>   
   <table class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<TD align="left">No</TD>
		<TD align="center" colspan="2"></TD>
		<TD align="left">Fixed Asset Write Off</TD>
		<TD align="left">Fixed Asset Write Type</TD>
		<TD align="left">Creation Date</TD>
		<TD align="left">Description</TD>
		<TD align="left">Fixed Asset Code</TD>
		<TD align="left">Fixed Asset Name</TD>
		<TD align="left">Fixed Asset Amount</TD>
		<TD align="left">Accumulated Depreciation Fixed Asset</TD>
		<TD align="left">Jobcode Others</TD>
		<TD align="left">Target Type Others</TD>
		<TD align="left">Target Code Others</TD>
		<TD align="left">A/R Others</TD>
		<TD align="left">Gain / Loss</TD>
		<TD align="left">Remarks</TD>
	</tr>
<%
	if Not rsGAD.eof then
		NO = 0
		While Not rsGAD.EOF
			NO = NO +1
			if no  mod 2 = 1 then cColor = "#EEEEEE" else cColor = "#E5E5E5"
%>
			<tr bgcolor="<%=cColor%>">
				<td align="center"><%=NO%></td>
				<td valign="middle">
					<a href="fixedasset_writeoff.asp?fccode=<%=rsGAD("fccode")%>&fixassetcode=<%=rsGAD("fixassetcode")%>&lineno=<%=rsGAD("lineno")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD("fccode")%>" border="0"></a>		
				</td>
				<td valign="middle">
					<a href="fixedasset_writeoff.asp?fccode=<%=rsGAD("fccode")%>&fixassetcode=<%=rsGAD("fixassetcode")%>&lineno=<%=rsGAD("lineno")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD("fccode")%>" border="0">
					</a>
				</td>
		   
				<td align="left" title="<%=rsGAD("fccode").Name%>"><%=rsGAD("fccode")%></td>
				<td align="left" title="<%=rsGAD("fctype").Name%>"><%=rsGAD("fctype")%></td>
				<td align="left" title="<%=rsGAD("transdate").Name%>"><%=rsGAD("transdate")%></td>
				<td align="left" title="<%=rsGAD("description").Name%>"><%=rsGAD("description")%></td>
				<td align="left" title="<%=rsGAD("fixassetcode").Name%>"><%=rsGAD("fixassetcode")%></td>
				<td align="left" title="<%=rsGAD("fixassetname").Name%>"><%=rsGAD("fixassetname")%></td>
				<td align="left" title="<%=rsGAD("fixassetamount").Name%>"><%=formatnumber(cdbl(to_nol(rsGAD("fixassetamount"))),2)%></td>
				<td align="left" title="<%=rsGAD("accumulated_depreciation").Name%>"><%=formatnumber(cdbl(to_nol(rsGAD("accumulated_depreciation"))),2)%></td>
				<td align="left" title="<%=rsGAD("jobcode_others").Name%>"><%=rsGAD("jobcode_others")%></td>
				<td align="left" title="<%=rsGAD("targettype_others").Name%>"><%=rsGAD("targettype_others")%></td>
				<td align="left" title="<%=rsGAD("targetcode_others").Name%>"><%=rsGAD("targetcode_others")%></td>
				<td align="left" title="<%=rsGAD("ar_others").Name%>"><%=rsGAD("ar_others")%></td>
				<td align="left" title="<%=rsGAD("gain_loss").Name%>"><%=formatnumber(cdbl(to_nol(rsGAD("gain_loss"))),2)%></td>
				<td align="left" title="<%=rsGAD("remarks").Name%>"><%=rsGAD("remarks")%></td>
			</tr>
<%
			rsGAD.MoveNext
		Wend
	end if 
	rsGAD.Close : set rsGAD = Nothing
%>
   </div>
   </TABLE>
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

// function getselectedfcba() {
// 	var sfcba = document.getElementById("selectedfcba");
// 	sfcba.value = document.getElementById("fcba").value;
// }

function getselectedline() {
	var sline = document.getElementById("selectedline");
	sline.value = document.getElementById("lineno").value;
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
	hitungrumus();
	getlistline();
	getselectedline();
});
</script>
</body>
</html>
