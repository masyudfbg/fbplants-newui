<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<title>Fixed Asset Received Task | iPlaS</title>
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

function removeOptions(selectbox)
{
    var i;
    for(i = selectbox.options.length - 1 ; i >= 0 ; i--)
    {
        selectbox.remove(i);
    }
}
//using the function:
function hitungrumus(){
    	var accumulated_depreciation = document.getElementById("accumulated_depreciation").value;
    	var total = document.getElementById("total").value;
    	// var totalvalue = document.getElementById("totalvalue").value;
    	var totalintraco = parseInt(total) - parseInt(accumulated_depreciation);
    	if (!isNaN(total)) {
    	 document.getElementById('intracoamount').value = totalintraco;
    	}
    }

	function hanyaAngka(evt) {
	  // var charCode = (evt.which) ? evt.which : event.keyCode
	  //  if (charCode > 31 && (charCode < 48 || charCode > 57))

	  //   return false;
	  // return true;
	  evt.value = evt.value.replace(/[^0-9,.-]/g,"");
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
var getAssetWriteOffInfo = function () {
	// build up the post string when passing variables to the server side page
	
	var fixassetcode = document.getElementById("fixassetcode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fixassetcode='+fixassetcode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfixedassetreceivedtask.asp', 'POST', PostStr, showAssetWriteOffInfo);
}

// The function for handling the response from the server
var showAssetWriteOffInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.total);
    // alert(s.assetcode);
    // update the Div to show the result from the server
	document.getElementById("fixassetname").value = s.fixassetname;
	document.getElementById("fixassetamount").value = s.fixassetamount;
	document.getElementById("assetcode").value = s.assetcode;
	document.getElementById("assetname").value = s.assetname;
	document.getElementById("assetgroupcode").value = s.assetgroupcode;
	document.getElementById("assetgroupname").value = s.assetgroupname;
	document.getElementById("assettaxcode").value = s.assettaxcode;
	document.getElementById("assettaxname").value = s.assettaxname;
	document.getElementById("uom").value = s.uom;
	document.getElementById("depreciationmonth").value = s.depreciationmonth;
	document.getElementById("targetcode_fixasset").value = s.targetcode_fixasset;
	document.getElementById("intracoamount").value = s.intracoamount;
	document.getElementById("accumulated_depreciation").value = s.accumulated_depreciation;

}

// launched from button click 
var getdatetransportask = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getdatefixassettransportask.asp', 'POST', PostStr, showDateTransporTaskInfo);
}

// The function for handling the response from the server
var showDateTransporTaskInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.total);
    // alert(s.assetcode);
    // update the Div to show the result from the server
	document.getElementById("fa_transportaskdate").value = s.fa_transportaskdate;
}

// launched from button click 
var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbabasedoncompany.asp', 'POST', PostStr, showlistfcba);
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
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var gettransportaskno = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedtransport = document.getElementById("selectedtransport").value;
	var PostStr = 'fcba='+selectedfcba+'&fa_transportask='+selectedtransport;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfixedassettranstransportaskbasedonfcba.asp', 'POST', PostStr, showlisttransportask);
}

// The function for handling the response from the server
var showlisttransportask = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    // alert("fccode");
    // update the Div to show the result from the server
	document.getElementById("fatransporlist").innerHTML = "";
	document.getElementById("fatransporlist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
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
fa_transportaskdate = request("fa_transportaskdate")

fixassetcode = request("fixassetcode")
fixassetname = request("fixassetname")
assetcode = request("assetcode")
assetname = request("assetname")
assetgroupcode = request("assetgroupcode")
assetgroupname = request("assetgroupname")
assettaxcode = request("assettaxcode")
assettaxname = request("assettaxname")
targettype_depreciation = request("targettype_depreciation")
targetcode_depreciation = request("targetcode_depreciation")
jobcode_depreciation = request("jobcode_depreciation")
accumulated_depreciation = request("accumulated_depreciation")
targettype_others = request("targettype_others")
targetcode_others = request("targetcode_others")
jobcode_intraco = request("jobcode_intraco")
targettype_intraco = request("targettype_intraco")
targetcode_intraco = request("targetcode_intraco")
jobcode_intraco  = request("jobcode_intraco")
intracoamount = request("intracoamount")
remarks = request("remarks")
description = request("description")
fixassetamount = request("fixassetamount")
amountreceived = request("amountreceived")
depreciationmonth = request("depreciationmonth")
uom = request("uom")
fcba_destination = request("fcba_destination")
targettype_fixasset = request("targettype_fixasset")
targetcode_fixasset = request("targetcode_fixasset")
fa_transportask = request("fa_transportask")
selectedtransport = request("fa_transportask")

fccompanycode = request("fccompanycode")
fcBA = request("fcBA")
selectedfcba = request("fcBA")

jobcode_intraco = request("jobcode_intraco")
sourcetype_intraco = request("sourcetype_intraco")
sourcecode_intraco = request("sourcecode_intraco")


if ""&transdate = "" then transdate = TextDate(date)

Sub InsertFixedAssetTrans

	set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
	q="select fcba_destination from fixedasset_transportask_det " & _
	  " where fccode='" & fixassetcode & "'"
	  ' response.write q
	rsTempPurchaseRequest.Open q, connstr
	if Not rsTempPurchaseRequest.Eof then 
		fcba_destination = rsTempPurchaseRequest("fcba_destination")  
	End if
	
	rsTempPurchaseRequest.Close
	set rsTempPurchaseRequest = Nothing

	Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select fccode from fixedasset where fccode = '" & fixassetcode & "' and fcba ='" & fcba & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then 

				q = " INSERT INTO fixedasset ( fccode" & _
					"                           , fcname" & _
					"                           , transdate" & _
					"                           , assetcode" & _
					"                           , assetname" & _
					"                           , assetgroupcode" & _
					"                           , assetgroupname" & _
					"                           , assettaxcode" & _
					"                           , assettaxname" & _
					"                           , uom" & _
					"                           , depreciationmethod_by" & _
					"                           , depreciationmonth" & _
					"                           , acquisitionvalue" & _
					"                           , additionalvalue" & _
					"                           , total" & _
					"                           , description" & _
					"                           , isactive" & _
					"                           , fcentry" & _
					"                           , fcedit" & _
					"                           , fcip" & _
					"                           , fcba" & _
					"                           , lastupdate" & _
					"                           , lasttime)" & _
					" VALUES ( '" & fixassetcode & "'" & _
					"          , '" & fixassetname & "'" & _
					"          , TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
					"          , '" & assetcode & "'" & _
					"          , '" & assetname & "'" & _
					"          , '" & assetgroupcode & "'" & _
					"          , '" & assetgroupname & "'" & _
					"          , '" & assettaxcode & "'" & _
					"          , '" & assettaxname & "'" & _
					"          , '" & uom & "'" & _
					"          , '" & depreciationmethod_by & "'" & _
					"          , '" & depreciationmonth & "'" & _
					"          , replace('" & fixassetamount & "',',','')" & _
					"          , replace('" & additionalvalue & "',',','')" & _
					"          , replace('" & totalacquisition & "',',','')" & _
					"          , '" & description & "'" & _
					"          , '1'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"          , '" & fcba_destination & "'" & _
					"          , SYSDATE" & _
					"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

				' response.write q
				connStr.execute(q)
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing

			'Untuk menghhitung jumlah total di header
				dtl_acquisitionvalue = 0
				' dtl_charge = 0
				set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
				q="select *  from fixedasset " & _
				  " where fccode='" & fixassetcode & "' and " & _
				  " fcBA='" & fcba & "'"
				  ' response.write q
				rsTempPurchaseRequest.Open q, connstr
				if Not rsTempPurchaseRequest.Eof then
					while Not rsTempPurchaseRequest.Eof 
					totalacquisition = cdbl(to_nol(rsTempPurchaseRequest("acquisitionvalue")))  + cdbl(to_nol(rsTempPurchaseRequest("additionalvalue")))
					rsTempPurchaseRequest.movenext
				wend 
					
				End if
				
				rsTempPurchaseRequest.Close
				set rsTempPurchaseRequest = Nothing
				totalacquisitions = totalacquisition
			'end add

			q="Update fixedasset set total =replace('" & totalacquisitions & "',',','')" & _
			 " where  fccode='" & fixassetcode & "' and fcBA='" & fcba_destination & "'"
			connStr.execute(q)
			' response.write q
		
		qLineMax = " select nvl(max(lineno),0) tempMaxLine" & _
				   " from fixedasset_depreciation_ledger" & _
				   " where fcCode = '" & fixassetcode & "'" & _
				   " and fcba = '" & fcba_destination & "'" 			
		' response.write qLineMax
		set rsqLineMax = Server.CreateObject("ADODB.Recordset")
		rsqLineMax.Open qLineMax, connstr
		if not rsqLineMax.eof then
			tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
		else
			tempMaxLine = 0
		end if 

		tempMaxLine = tempMaxLine + 1
		txtLineNoHist = tempMaxLine


		hasdata = 0
		q = " SELECT COUNT(1) hasdata FROM fixedasset_depreciation_ledger WHERE fccode = '" & fixassetcode & "' and fcba = '"& fcba_destination &"' and lineno = '"& tempMaxLine &"'"
		Set rscekpk = Server.CreateObject("ADODB.Recordset")
		rscekpk.Open q, connstr
		
		if Not rscekpk.eof then
			hasdata = cdbl(to_nol(rscekpk("hasdata")))
		else
			hasdata = 0
		end if
		
		if hasdata = 1 then
			err.number = 1001
			err.description = "Date Transaction : " & date_transaction & " , already exist"
		else
			'Untuk menghhitung jumlah total di header
			dtl_acquisitionvalue = 0
			' dtl_charge = 0
			set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			q="select *  from fixedasset " & _
			  " where fccode='" & fixassetcode & "' and " & _
			  " fcBA='" & fcba & "'"

			rsTempPurchaseRequest.Open q, connstr
			if Not rsTempPurchaseRequest.Eof then
				while Not rsTempPurchaseRequest.Eof 
					acquisitionvalue = Cdbl(to_nol(rsTempPurchaseRequest("acquisitionvalue")))
					depreciationmonth = Cdbl(to_nol(rsTempPurchaseRequest("depreciationmonth")))
				rsTempPurchaseRequest.movenext
			wend 
				
			End if
			
			rsTempPurchaseRequest.Close
			set rsTempPurchaseRequest = Nothing
			' totalacquisitionvalue = dtl_acquisitionvalue
			'end add

			if Cdbl(to_nol(depreciationmonth)) = 0 then dtl_depremonth = 0 else dtl_depremonth = Cdbl(to_nol(acquisitionvalue)) / Cdbl(to_nol(depreciationmonth))

				monthtobe = 1
				lastdepreciationmonth = depreciationmonth - monthtobe

			q = " INSERT INTO fixedasset_depreciation_ledger ( FCCODE," & _
				" DATE_TRANSACTION," & _
				" FCTYPE ," & _
				" DOC_NO," & _
				" AMOUNTDEPRECIATION    ," & _
				" DEPRECIATIONMONTH_LEDGER    ," & _
				" MONTHTOBE  ," & _
				" LASTDEPRECIATIONMONTH ," & _
				" MONTHLYDEPRECIATION  ," & _
				" YEARLYDEPRECIATION   ," & _
				" LINENO ," & _
				" BEGINING_BALANCE ," & _
				" FCENTRY," & _
				" FCEDIT    ," & _
				" FCIP   ," & _
				" FCBA   ," & _
				" LASTUPDATE  ," & _
				" LASTTIME )" & _
				" VALUES ( '" & fixassetcode & "'" & _
				"          , TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
				"          , 'FROM STRAIGHT LINE'" & _
				"          , '" & fixassetcode & "'" & _
				"          , replace('" & acquisitionvalue & "',',','')" & _
				"          , '" & depreciationmonth & "'" & _
				"          , '" & monthtobe & "'" & _
				"          , '" & lastdepreciationmonth & "'" & _
				"          , '" & dtl_depremonth & "'" & _
				"          , '" & dtl_depremonth & "'" & _
				"          , '" & tempMaxLine & "'" & _
				"          , 'FALSE'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , '" & fcba_destination & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

			'response.write q
			connStr.execute(q)

			'Untuk menghhitung jumlah total di header
			dtl_yearlydepreciation = 0
			' dtl_charge = 0
			set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
			q="select *  from fixedasset_depreciation_ledger " & _
			  " where fccode='" & fixassetcode & "' and " & _
			  " fcBA='" & fcBA & "'"

			rsTempPurchaseRequest.Open q, connstr
			if Not rsTempPurchaseRequest.Eof then
				while Not rsTempPurchaseRequest.Eof 
				dtl_yearlydepreciation = dtl_yearlydepreciation + cdbl(to_nol(rsTempPurchaseRequest("monthlydepreciation")))
				' dtl_charge = dtl_charge + cdbl(to_nol(rsTempPurchaseRequest("othercharge")))
				rsTempPurchaseRequest.movenext
			wend 
				
			End if
			
			rsTempPurchaseRequest.Close
			set rsTempPurchaseRequest = Nothing
			totalyearlydepreciation = dtl_yearlydepreciation
			'end add
			
		 	qLineMax = " select nvl(lineno,0) lineno" & _
					   " from fixedasset_depreciation_ledger" & _
					   " where fcCode = '" & fixassetcode & "'" & _
					   " and fcba = '" & fcba_destination & "'" 			
				' response.write qLineMax
				set rsqLineMax = Server.CreateObject("ADODB.Recordset")
				rsqLineMax.Open qLineMax, connstr
				if not rsqLineMax.eof then
					lineno = cdbl(rsqLineMax("lineno"))
				else
					lineno = 0
				end if 

		   q="Update fixedasset_depreciation_ledger set yearlydepreciation=replace('" & totalyearlydepreciation & "',',','')" & _
			 " where  fccode='" & fccode & "' and date_transaction= TO_DATE('" & date_transaction & "','mm/dd/yyyy') and lineno = '"& lineno &"' and fcBA='" & fcba_destination & "'"
			connStr.execute(q)
			' response.write q
		end if
End Sub

Sub UpdateAmountTransporTask
	set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
	q="select fcba_destination from fixedasset_transportask_det " & _
	  " where fccode='" & fixassetcode & "'"
	  ' response.write q
	rsTempPurchaseRequest.Open q, connstr
	if Not rsTempPurchaseRequest.Eof then 
		fcba_destination = rsTempPurchaseRequest("fcba_destination")  
	End if
	
	rsTempPurchaseRequest.Close
	set rsTempPurchaseRequest = Nothing

	q = " UPDATE fixedasset_transportask_det SET amountreceived = replace('"& fixassetamount &"',',','')" & _
		" WHERE fcBA = '" & fcba & "'" & _
		"       AND fixassetcode = '" & fixassetcode & "' AND fccode = '"& fa_transportask &"'" & _
		" AND fcba_destination = '"& fcba_destination &"'"
		' response.write q
	connStr.execute(q)
End Sub

Sub DeleteFixedAssetTransport
	set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
	q="select fcba_destination from fixedasset_transportask_det " & _
	  " where fccode='" & fixassetcode & "'"
	  ' response.write q
	rsTempPurchaseRequest.Open q, connstr
	if Not rsTempPurchaseRequest.Eof then 
		fcba_destination = rsTempPurchaseRequest("fcba_destination")  
	End if
	
	rsTempPurchaseRequest.Close
	set rsTempPurchaseRequest = Nothing

	Dim rsSupplyingDeclaration, strSuppyingDeclaration
	strSuppyingDeclaration	=	" select fr.transdate, fr.fccode, fr.description, fr.fa_transportask, fr.fa_transportaskdate, fr.fcba_source, bu.fccompanycode " & _
								" from fixedasset_receivedtask fr " & _ 
								" inner join businessunit bu on (bu.fccode = fr.fcba) where fr.fcCode = '" & fccode & "'"
							
	set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
	rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
	if Not rsSupplyingDeclaration.Eof then
		transdate = rsSupplyingDeclaration("transdate")
		fcCode = rsSupplyingDeclaration("fcCode")
		Description = rsSupplyingDeclaration("Description")
		fccompanycode = rsSupplyingDeclaration("fccompanycode")
		fa_transportask = rsSupplyingDeclaration("fa_transportask")
		fa_transportaskdate = rsSupplyingDeclaration("fa_transportaskdate")
		fcba = rsSupplyingDeclaration("fcba_source")
		selectedfcba = rsSupplyingDeclaration("fcba_source")
		selectedtransport = rsSupplyingDeclaration("fa_transportask")
	End if
	' response.Write strSuppyingDeclaration
	rsSupplyingDeclaration.Close
	Set rsSupplyingDeclaration = Nothing

	q	=	"DELETE FROM fixedasset WHERE fccode = '"& fixassetcode &"' AND fcba = '"& fcba_destination &"'"
	' response.write q
	connStr.execute(q)

	q	=	"DELETE FROM fixedasset_depreciation_ledger WHERE fccode = '"& fixassetcode &"' AND fcba = '"& fcba_destination &"'"
	' response.write q
	connStr.execute(q)

	q = " UPDATE fixedasset_transportask_det SET amountreceived = '0'" & _
		" WHERE fcBA = '" & fcba & "'" & _
		"       AND fixassetcode = '" & fixassetcode & "' AND fccode = '"& fa_transportask &"'" & _
		" AND fcba_destination = '"& fcba_destination &"'"
	' response.write q
	connStr.execute(q)
End Sub

if Request("txtJenis")="New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
		connstr.BeginTrans
		
		set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
		q="select fcba_destination from fixedasset_transportask_det " & _
		  " where fccode='" & fixassetcode & "'"
		  ' response.write q
		rsTempPurchaseRequest.Open q, connstr
		if Not rsTempPurchaseRequest.Eof then 
			fcba_destination = rsTempPurchaseRequest("fcba_destination")  
		End if
		
		rsTempPurchaseRequest.Close
		set rsTempPurchaseRequest = Nothing

		strFieldDetail = " SELECT COUNT(1) total FROM fixedasset_receivedtask_det " & _
                         " WHERE fccode = '" & fccode & "'" & _
                         "       AND fcBA = '" & fcba_destination & "'"								
		' Response.Write strFieldDetail
		Set rsQty = Server.CreateObject("ADODB.Recordset")
		rsQty.Open strFieldDetail, connstr							
		if not rsQty.Eof then
			has_detail = cdbl (to_nol(rsQty("total")))
		End If
		
		if has_detail = 0 then
			q = " DELETE FROM fixedasset_receivedtask WHERE fccode = '" & fccode & "' AND fcba = '"& fcba_destination &"'"
				'Response.Write q
			connStr.execute(q)
		else
			err.number = 1001
			err.description = " Fixed Asset Received Task """ & fcCode & """ has detail, please delete all detail first!"
		End If
		
		if err.number = 0 then
		   connStr.CommitTrans
		end if
        cmdView	= "True"


	elseif Request("txtJenis")= "Delete" then

		set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
		q="select fcba_destination from fixedasset_transportask_det " & _
		  " where fccode='" & fixassetcode & "'"
		  ' response.write q
		rsTempPurchaseRequest.Open q, connstr
		if Not rsTempPurchaseRequest.Eof then 
			fcba_destination = rsTempPurchaseRequest("fcba_destination")  
		End if
		
		rsTempPurchaseRequest.Close
		set rsTempPurchaseRequest = Nothing

		q	=	"DELETE FROM fixedasset_receivedtask_det WHERE fixassetcode = '"& fixassetcode &"' AND fccode = '"& fccode &"' AND fcba = '"& fcba_destination &"'"
		' response.write q
		connStr.execute(q)

		call DeleteFixedAssetTransport
		
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
		q="select fcba_destination from fixedasset_transportask_det " & _
		  " where fccode='" & fixassetcode & "'"
		  ' response.write q
		rsTempPurchaseRequest.Open q, connstr
		if Not rsTempPurchaseRequest.Eof then 
			fcba_destination = rsTempPurchaseRequest("fcba_destination")  
		End if
		
		rsTempPurchaseRequest.Close
		set rsTempPurchaseRequest = Nothing

		hasdata = 0
		q = " SELECT COUNT(1) hasdata FROM fixedasset_receivedtask_det WHERE fccode = '"& fccode &"' and fcba = '"& fcba_destination &"'"
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
			err.description = "Fixed Asset Code : " & fcCode & " , already exist"
		else
			
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select fccode from fixedasset_receivedtask where fccode = '" & fccode & "' and fcba ='" & fcba_destination & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then
						q = " INSERT INTO fixedasset_receivedtask ( fccode" & _
							"                           , transdate" & _
							"                           , description" & _
							"                           , fcba_source" & _
							"                           , fa_transportask" & _
							"                           , fa_transportaskdate" & _
							"                           , fcentry" & _
							"                           , fcedit" & _
							"                           , fcip" & _
							"                           , fcba" & _
							"                           , lastupdate" & _
							"                           , lasttime)" & _
							" VALUES ( '" & fccode & "'" & _
							"          , TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
							"          , '" & description & "'" & _
							"          , '" & fcba & "'" & _
							"          , '" & fa_transportask & "'" & _
							"          , TO_DATE('" & fa_transportaskdate & "', 'mm/dd/yyyy')" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & fcba_destination & "'" & _
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

						' response.write q
						connStr.execute(q)
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing

			q = " INSERT INTO fixedasset_receivedtask_det (fccode" & _
				"							, fixassetcode" & _
				"							, fixassetname" & _
				"                           , assetcode" & _
				"                           , assetname" & _
				"                           , assetgroupcode" & _
				"                           , assetgroupname" & _
				"                           , assettaxcode" & _
				"                           , assettaxname" & _
				"                           , uom" & _
				"                           , depreciationmonth" & _
				"                           , fixassetamount" & _
				"                           , targettype_fixasset" & _
				"                           , targetcode_fixasset" & _
				"                           , sourcetype_depreciation" & _
				"                           , sourcecode_depreciation" & _
				"                           , jobcode_depreciation" & _
				"                           , accumulated_depreciation" & _
				"                           , jobcode_intraco" & _
				"                           , sourcetype_intraco" & _
				"                           , sourcecode_intraco" & _
				"                           , intraco_amount" & _
				"                           , remarks" & _
				"                           , fcentry" & _
				"                           , fcedit" & _
				"                           , fcip" & _
				"                           , fcba" & _
				"                           , lastupdate" & _
				"                           , lasttime)" & _
				" VALUES ( '" & fccode & "'" & _
				"          , '" & fixassetcode & "'" & _
				"          , '" & fixassetname & "'" & _
				"          , '" & assetcode & "'" & _
				"          , '" & assetname & "'" & _
				"          , '" & assetgroupcode & "'" & _
				"          , '" & assetgroupname & "'" & _
				"          , '" & assettaxcode & "'" & _
				"          , '" & assettaxname & "'" & _
				"          , '" & uom & "'" & _
				"          , '" & depreciationmonth & "'" & _
				"          ,replace('" & fixassetamount & "',',','')" & _
				"          , '" & targettype_fixasset & "'" & _
				"          , '" & targetcode_fixasset & "'" & _
				"          , '" & targettype_depreciation & "'" & _
				"          , '" & targetcode_depreciation & "'" & _
				"          , '" & jobcode_depreciation & "'" & _
				"          ,replace('" & accumulated_depreciation & "',',','')" & _
				"          , '" & jobcode_intraco & "'" & _
				"          , '" & sourcetype_intraco & "'" & _
				"          , '" & sourcecode_intraco & "'" & _
				"          ,replace('" & intracoamount & "',',','')" & _
				"          , '" & remarks & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , '" & fcba_destination & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

			' response.write q
			connStr.execute(q)

			call InsertFixedAssetTrans
			call UpdateAmountTransporTask

		end if
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
		q="select fcba_destination from fixedasset_transportask_det " & _
		  " where fccode='" & fixassetcode & "'"
		  ' response.write q
		rsTempPurchaseRequest.Open q, connstr
		if Not rsTempPurchaseRequest.Eof then 
			fcba_destination = rsTempPurchaseRequest("fcba_destination")  
		End if
		
		rsTempPurchaseRequest.Close
		set rsTempPurchaseRequest = Nothing

		q = " UPDATE fixedasset_receivedtask_det" & _
            " SET   fixassetname = '" & fixassetname & "'" & _
			"     , assetcode = replace('" & assetcode & "',',','')" & _
			"     , assetname = '" & assetname & "'" & _
			"     , assetgroupcode = replace('" & assetgroupcode & "',',','')" & _
			"     , assetgroupname = '" & assetgroupname & "'" & _
			"     , assettaxcode = replace('" & assettaxcode & "',',','')" & _
			"     , assettaxname = '" & assettaxname & "'" & _
			"     , uom = '" & uom & "'" & _
			"     , depreciationmonth = '" & depreciationmonth & "'" & _
			"     , accumulated_depreciation = replace('" & accumulated_depreciation & "',',','')" & _
			"     , fixassetamount = replace('" & fixassetamount & "',',','')" & _
			"     , jobcode_depreciation = '" & jobcode_depreciation & "'" & _
			"     , sourcetype_depreciation = '" & targettype_depreciation & "'" & _
			"     , sourcecode_depreciation = '" & targetcode_depreciation & "'" & _
			"     , sourcetype_intraco = '" & sourcetype_intraco & "'" & _
			"     , sourcecode_intraco = '" & sourcecode_intraco & "'" & _
			"     , jobcode_intraco = '" & jobcode_intraco & "'" & _
			"     , intraco_amount = replace('" & intracoamount & "',',','')" & _
			"     , remarks = '" & remarks & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , lastupdate = SYSDATE" & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
			" WHERE fixassetcode = '" & fixassetcode & "' AND fccode = '"& fccode &"' AND fcba = '"& fcba_destination &"'" 
		' response.write q
		connStr.execute(q)	

		call DeleteFixedAssetTransport
		call InsertFixedAssetTrans
		call UpdateAmountTransporTask
        
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


' If Len(Trim(date_transaction)) = 0 Then	date_transaction = textDate(date)
If Len(Trim(fa_transportaskdate)) = 0 Then	fa_transportaskdate = textDate(date)

Dim rsSupplyingDeclaration, strSuppyingDeclaration
strSuppyingDeclaration	=	" select fr.transdate, fr.fccode, fr.description, fr.fa_transportask, fr.fa_transportaskdate, fr.fcba_source, bu.fccompanycode " & _
							" from fixedasset_receivedtask fr " & _ 
							" inner join businessunit bu on (bu.fccode = fr.fcba) where fr.fcCode = '" & fccode & "'"
						
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	transdate = rsSupplyingDeclaration("transdate")
	fcCode = rsSupplyingDeclaration("fcCode")
	Description = rsSupplyingDeclaration("Description")
	fccompanycode = rsSupplyingDeclaration("fccompanycode")
	fa_transportask = rsSupplyingDeclaration("fa_transportask")
	fa_transportaskdate = rsSupplyingDeclaration("fa_transportaskdate")
	fcba = rsSupplyingDeclaration("fcba_source")
	selectedfcba = rsSupplyingDeclaration("fcba_source")
	selectedtransport = rsSupplyingDeclaration("fa_transportask")
End if
' response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcba") & "'" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & transdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('bank')"
			
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
  
<form name="frmEntry" method="post" action="fixedasset_receivedtask.asp">
	<table class="data" width="80%">
	<tr>
		<td align="left">Company</td>
		<td align="left">
			<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();" class="chosen-select" style="width: 40%">
				<option value=""></option>
				<%
				qlov = "SELECT DISTINCT bu.fccompanycode fccode FROM user_fcba_tab uft INNER JOIN businessunit bu ON bu.fccode = uft.fcba WHERE fcusercode = '" & Session("CodeName") & "'"
				Set rslov = Server.CreateObject("ADODB.Recordset")
				rslov.Open qlov, connstr 
				if Not rslov.eof then
					While Not rslov.EOF
				%>
				<option value="<%=rslov("fccode")%>" <%if cstr(""&fccompanycode) = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%></option>
				<%
						rslov.MoveNext
					Wend
				end if
				%>
			</select>
			<font color="#FF0000">*</font>
	</tr>
	<tr>
		<td align="left">Business Unit Source</td>
		<td align="left">
			<div id="fcbalist"></div>
			<font color="#FF0000">*</font>
		</td>
		<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
	</tr>
	<tr>
		<td align="left">Fixed Asset Received Task No</td>
		<td align="left">
			<input type="text" name="fcCode" value="<%=fcCode%>" maxlength="20" size="20" autocomplete="off">
			<font color="#FF0000">&nbsp;*</font>
		</td>
	</tr>
	<tr>
		<td align="left">Fixed Asset Transport Task No Source</td>
		<td align="left">
			<div id="fatransporlist"></div>
			<font color="#FF0000">*</font>
			<input type="Hidden" id="selectedtransport" name="selectedtransport" value="<%=selectedtransport%>">
		</td>
	</tr>
	<tr>
		<td align="left">Fixed Asset Creation Date</td>
		<td align="left">
			<input type="text" maxlength="10" id="fa_transportaskdate" name="fa_transportaskdate" size="10" value="<%=fa_transportaskdate%>" style="background-color: #eee" >
			<font color="#FF0000">&nbsp;*</font>
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
	<%If val_IsTrue = "Y" Then%>
	<a href="fixedasset_receivedtask.asp?fcCode=<%=fccode%>&txtJenis=HeaderDelete" onClick="return confirm('Are you sure to delete');">
		<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0">
	</a>
	<%Else%>
		 <img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> 
	<%End If%>
	<a href="Printfixedasset_receivedtask.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
	</td>
</tr>
</table>  
</form>
<%
if cmdView = "True" AND len(trim(fcCode)) > 0 Then
	q = "SELECT bu.fccompanycode," & _
		"       ftt.fccode," & _
		"       ftt.transdate," & _
		"       ftt.description," & _
		"       ftd.fixassetcode," & _
		"       ftd.fixassetname," & _
		"       ftd.assetcode," & _
		"       ftd.assetname," & _
		"       ftd.assetgroupcode," & _
		"       ftd.assetgroupname," & _
		"       ftd.assettaxcode," & _
		"       ftd.assettaxname," & _
		"       ftd.uom," & _
		"       ftd.depreciationmonth," & _
		"       ftd.fixassetamount," & _
		"       ftd.targettype_fixasset," & _
		"       ftd.targetcode_fixasset," & _
		"       ftd.sourcetype_depreciation," & _
		"       ftd.sourcecode_depreciation," & _
		"       ftd.jobcode_depreciation," & _
		"       ftd.accumulated_depreciation," & _
		"       ftd.jobcode_intraco," & _
		"       ftd.sourcetype_intraco," & _
		"       ftd.sourcecode_intraco," & _
		"       ftd.intraco_amount," & _
		"       ftd.remarks" & _
		"  FROM fixedasset_receivedtask ftt" & _
		"       INNER JOIN fixedasset_receivedtask_det ftd" & _
		"          ON (ftt.fccode = ftd.fccode AND ftt.fcba = ftd.fcba)" & _
		"		LEFT JOIN businessunit bu ON (bu.fccode = ftt.fcba)" & _
		" WHERE     ftt.fccode = '"& fcCode &"'" & _
		"       AND ftt.transdate = TO_DATE ('"& transdate &"', 'MM/DD/YYYY')" & _
		"       AND ftt.fcba = '"& Session("fcBA") &"'" 

	' if ""&searchdata <> "" then q = q & " AND (UPPER(fa.fccode) LIKE UPPER('%"&searchdata&"%') OR UPPER(fa.fcname) LIKE ('%"&searchdata&"%'))"
	' if Request("txtJenis")="Edit" then q = q & " AND fa.fccode = '" & fccode & "'"
	
	' response.write q
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.Open q, connstr
	
	if Request("txtJenis")="Edit" then
		if Not rsGAD.eof then
			fccompanycode = rsGAD("fccompanycode")
			fcCode = rsGAD("fcCode")
			description = rsGAD("description")
			transdate = rsGAD("transdate")
			fixassetcode = rsGAD("fixassetcode")
			fixassetname = rsGAD("fixassetname")
			assetcode = rsGAD("assetcode")
			assetname = rsGAD("assetname")
			assetgroupcode = rsGAD("assetgroupcode")
			assetgroupname = rsGAD("assetgroupname")
			targettype_fixasset = rsGAD("targettype_fixasset")
			targetcode_fixasset = rsGAD("targetcode_fixasset")
			assettaxcode = rsGAD("assettaxcode")
			assettaxname = rsGAD("assettaxname")
			uom = rsGAD("uom")
			depreciationmonth = rsGAD("depreciationmonth")
			fixassetamount = rsGAD("fixassetamount")
			targettype_depreciation = rsGAD("sourcetype_depreciation")
			targetcode_depreciation = rsGAD("sourcecode_depreciation")
			jobcode_depreciation = rsGAD("jobcode_depreciation")
			accumulated_depreciation = rsGAD("accumulated_depreciation")
			sourcetype_intraco = rsGAD("sourcetype_intraco")
			sourcecode_intraco = rsGAD("sourcecode_intraco")
			jobcode_intraco = rsGAD("jobcode_intraco")
			intracoamount = rsGAD("intraco_amount")
			remarks = rsGAD("remarks")
		end if
	end if

%>
   
   <!-- Form Entry And Edit --> 
   
   <form action="fixedasset_receivedtask.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
			<tr>
				<%if Request("txtJenis")="Edit" then%>
				<td align="left">Fixed Asset Code</td>
				<td align="left">
        			<input type="text" size="50" id="fixassetcode" name="fixassetcode" value="<%=fixassetcode%>" onchange="javascript:getAssetWriteOffInfo();" readonly style="background-color:#ccc;">
        		</td>
        		<%Else%>
				<td align="left">Fixed Asset Code</td>
				<td align="left">
					<select id="fixassetcode" name="fixassetcode" onchange="javascript:getAssetWriteOffInfo();" class="chosen-select" style="width: 70%">
						<option value=""></option>
						<%
						qlov =  "SELECT fatd.fixassetcode, fatd.fixassetname" & _
								"  FROM fixedasset_transportask fat" & _
								"       INNER JOIN fixedasset_transportask_det fatd" & _
								"          ON (fat.fccode = fatd.fccode AND fat.fcba = fatd.fcba)" & _
								"       LEFT JOIN user_fcba_tab uft ON (fat.fcba = uft.fcba)" & _
								" WHERE uft.fcusercode = '"& Session("CodeName") &"' AND fat.fcba = '"& fcba &"' AND fat.fccode = '"& fa_transportask &"'" 

						' response.write qlov
						Set rslov = Server.CreateObject("ADODB.Recordset")
						rslov.Open qlov, connstr 
						if Not rslov.eof then
							While Not rslov.EOF
						%>
						<option value="<%=rslov("fixassetcode")%>" <%if fixassetcode = rslov("fixassetcode") then%>selected<%end if%>><%=rslov("fixassetcode") & ": " & rslov("fixassetname")%></option>
						<%
								rslov.MoveNext
							Wend
						end if
						%>
					</select>
				</td>
				<%End if%>
				<td align="left">Target Type Fixed Asset</td>
					<td align="left">
						<input type="text" size="50" id="targettype_fixasset" name="targettype_fixasset" value="<%=targettype_fixasset%>">
					</td>
			</tr>
			<tr>
				<td align="left">Fixed Asset Name</td>
					<td align="left">
						<input type="text" size="50" id="fixassetname" name="fixassetname" value="<%=fixassetname%>" readonly="true"  style="background-color:#ccc;">
					</td>
				</td>
				<td align="left">Target Code Fixed Code</td>
					<td align="left">
						<input type="text" size="50" id="targetcode_fixasset" name="targetcode_fixasset" value="<%=targetcode_fixasset%>" readonly="true"  style="background-color:#ccc;" >
					</td>
			</tr>
			<tr>
				<td align="left">Asset Code</td>
					<td align="left">
						<input type="text" size="50" id="assetcode" name="assetcode" value="<%=assetcode%>" readonly="true"  style="background-color:#ccc;">
					</td>
				<td align="left">Jobcode Depreciation</td>
					<td align="left">
						<input type="text" size="50" id="jobcode_depreciation" name="jobcode_depreciation" value="<%=jobcode_depreciation%>" >
					</td>
			<tr>
				<td align="left">Asset Name</td>
					<td align="left">
						<input type="text" size="50" id="assetname" name="assetname" value="<%=assetname%>" readonly="true"  style="background-color:#ccc;">
					</td>
				<td align="left">Source Type Depreciation</td>
					<td align="left">
						<input type="text" size="50" id="targettype_depreciation" name="targettype_depreciation" value="<%=targettype_depreciation%>">
					</td>
			<tr>
				<td align="left">Asset Group Code</td>
					<td align="left">
						<input type="text" size="50" id="assetgroupcode" name="assetgroupcode" value="<%=assetgroupcode%>" readonly="true"  style="background-color:#ccc;">
					</td>
				<td align="left">Source Code Depreciation</td>
					<td align="left">
						<input type="text" size="50" id="targetcode_depreciation" name="targetcode_depreciation" value="<%=targetcode_depreciation%>">
					</td>
			<tr>
				<td align="left">Asset Group Name</td>
					<td align="left">
						<input type="text" size="50" id="assetgroupname" name="assetgroupname" value="<%=assetgroupname%>" readonly="true"  style="background-color:#ccc;">
					</td>
				<td align="left">Accumulated Depreciation Fixed Asset</td>
					<td align="left">
						<input type="text" size="50" id="accumulated_depreciation" name="accumulated_depreciation" value="<%=accumulated_depreciation%>" readonly="true" onkeyup="hitungrumus();" style="background-color:#ccc;" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
					</td>
			<tr>
				<td align="left">Asset Tax Group Code</td>
					<td align="left">
						<input type="text" size="50" id="assettaxcode" name="assettaxcode" value="<%=assettaxcode%>" readonly="true"  style="background-color:#ccc;">
					</td>
				<td align="left">Jobcode Intraco Sender</td>
					<td align="left">
						<input type="text" size="50" id="jobcode_intraco" name="jobcode_intraco" value="<%=jobcode_intraco%>" >
					</td>
			<tr>
				<td align="left">Asset Tax Group Name</td>
					<td align="left">
						<input type="text" size="50" id="assettaxname" name="assettaxname" value="<%=assettaxname%>" readonly="true"  style="background-color:#ccc;">
					</td>
				<td align="left">Source Type Intraco Sender</td>
					<td align="left">
						<input type="text" size="50" id="targettype_intraco" name="sourcetype_intraco" value="<%=sourcetype_intraco%>" >
					</td>
			<tr>
				<td align="left">Unit Of Measure</td>
					<td align="left">
						<input type="text" size="50" id="uom" name="uom" value="<%=uom%>" readonly="true"  style="background-color:#ccc;" >
					</td>
				<td align="left">Source Code Intraco Sender</td>
					<td align="left">
						<input type="text" size="50" id="targetcode_intraco" name="sourcecode_intraco" value="<%=sourcecode_intraco%>" >
					</td>
			</tr>
			<tr>
				<td align="left">Depreciation Month</td>
					<td align="left">
						<input type="text" size="50" id="depreciationmonth" name="depreciationmonth" value="<%=depreciationmonth%>" readonly="true"  style="background-color:#ccc;">
					</td>
				<td align="left">Intraco Amount</td>
					<td align="left">
						<input type="text" size="50" id="intracoamount" name="intracoamount" value="<%=intracoamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly>
					</td>
			</tr>
			<tr>
				<td align="left">Fix Asset Amount</td>
					<td align="left">
						<input type="text" size="50" id="fixassetamount" name="fixassetamount" value="<%=fixassetamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="background:#ccc" readonly onkeyup="hitungrumus();">
					</td>
				<td align="left">Remarks</td>
					<td align="left">
						<textarea id="remarks" name="remarks" ><%=remarks%></textarea>
					</td>
			</tr>
		<tr>
			<td colspan="2" align="left">
				<input type="Hidden" name="searchdata" value="<%=searchdata%>">
			  	<input Type="Hidden" Value="<%=request("transdate")%>" Name="transdate">
			  	<input Type="Hidden" Value="<%=request("fa_transportaskdate")%>" Name="fa_transportaskdate">
			  	<input Type="Hidden" Value="<%=request("selectedtransport")%>" Name="selectedtransport">
			  	<input Type="Hidden" Value="<%=request("fa_transportask")%>" Name="fa_transportask">
				<input Type="Hidden" Value="<%=request("fcCode")%>" Name="fcCode">
				<input Type="Hidden" Value="<%=request("fcba")%>" Name="fcba">
				<input Type="Hidden" Value="<%=request("selectedfcba")%>" Name="selectedfcba">
				<input Type="Hidden" Value="<%=request("description")%>" Name="description">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%If val_IsTrue = "Y" Then%>
				  <%IF Request("txtJenis")="Edit" then%>
					   <input type="hidden" name="txtJenis" value="Edit">
					   <input Type="Submit"  name="Add" value="Update">
				  <%Else%>
					   <input type="hidden" name="txtJenis" value="New">
					   <input Type="Submit"  name="Add" value="Add">
				  <%End if%>
				  <input Type="Reset"  name="Reset" value="Reset">
				  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				<%End if%>
			</td>
		</tr>
   </table>
   </form>   
   <table class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<TD align="left">No</TD>
		<TD align="center" colspan="2"></TD>
		<TD align="left">Company</TD>
		<TD align="left">Fixed Asset Transportask</TD>
		<TD align="left">Creation Date</TD>
		<TD align="left">Description</TD>
		<TD align="left">Fixed Asset Code</TD>
		<TD align="left">Fixed Asset Name</TD>
		<TD align="left">Asset Code</TD>
		<TD align="left">Asset Name</TD>
		<TD align="left">Asset Group Code</TD>
		<TD align="left">Asset Group Name</TD>
		<TD align="left">Asset Tax Code</TD>
		<TD align="left">Asset Tax Name</TD>
		<TD align="left">UoM</TD>
		<TD align="left">Depreciation Month</TD>
		<TD align="left">Fix Asset Amount</TD>
		<TD align="left">Target Type Fixed Asset</TD>
		<TD align="left">Target Code Fixed Asset</TD>
		<TD align="left">Jobcode Depreciation</TD>
		<TD align="left">Target Type Depreciation</TD>
		<TD align="left">Target Code Depreciation</TD>
		<TD align="left">Accumulated Depreciation</TD>
		<TD align="left">Jobcode Intraco</TD>
		<TD align="left">Target Type Intraco</TD>
		<TD align="left">Target Code Intraco</TD>
		<TD align="left">Intraco Amount</TD>
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
					<a href="fixedasset_receivedtask.asp?fccode=<%=rsGAD("fccode")%>&fixassetcode=<%=rsGAD("fixassetcode")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD("fccode")%>" border="0"></a>		
				</td>
				<td valign="middle">
					<a href="fixedasset_receivedtask.asp?fccode=<%=rsGAD("fccode")%>&fixassetcode=<%=rsGAD("fixassetcode")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD("fccode")%>" border="0">
					</a>
				</td>
		   
				<td align="left" title="<%=rsGAD("fccompanycode").Name%>"><%=rsGAD("fccompanycode")%></td>
				<td align="left" title="<%=rsGAD("fcCode").Name%>"><%=rsGAD("fcCode")%></td>
				<td align="left" title="<%=rsGAD("transdate").Name%>"><%=rsGAD("transdate")%></td>
				<td align="left" title="<%=rsGAD("description").Name%>"><%=rsGAD("description")%></td>
				<td align="left" title="<%=rsGAD("fixassetcode").Name%>"><%=rsGAD("fixassetcode")%></td>
				<td align="left" title="<%=rsGAD("fixassetname").Name%>"><%=rsGAD("fixassetname")%></td>
				<td align="left" title="<%=rsGAD("assetcode").Name%>"><%=rsGAD("assetcode")%></td>
				<td align="left" title="<%=rsGAD("assetname").Name%>"><%=rsGAD("assetname")%></td>
				<td align="left" title="<%=rsGAD("assetgroupcode").Name%>"><%=rsGAD("assetgroupcode")%></td>
				<td align="left" title="<%=rsGAD("assetgroupname").Name%>"><%=rsGAD("assetgroupname")%></td>
				<td align="left" title="<%=rsGAD("assettaxcode").Name%>"><%=rsGAD("assettaxcode")%></td>
				<td align="left" title="<%=rsGAD("assettaxname").Name%>"><%=rsGAD("assettaxname")%></td>
				<td align="left" title="<%=rsGAD("uom").Name%>"><%=rsGAD("uom")%></td>
				<td align="left" title="<%=rsGAD("depreciationmonth").Name%>"><%=rsGAD("depreciationmonth")%></td>
				<td align="left" title="<%=rsGAD("fixassetamount").Name%>"><%=rsGAD("fixassetamount")%></td>
				<td align="left" title="<%=rsGAD("targettype_fixasset").Name%>"><%=rsGAD("targettype_fixasset")%></td>
				<td align="left" title="<%=rsGAD("targetcode_fixasset").Name%>"><%=rsGAD("targetcode_fixasset")%></td>
				<td align="left" title="<%=rsGAD("jobcode_depreciation").Name%>"><%=rsGAD("jobcode_depreciation")%></td>
				<td align="left" title="<%=rsGAD("sourcetype_depreciation").Name%>"><%=rsGAD("sourcetype_depreciation")%></td>
				<td align="left" title="<%=rsGAD("sourcecode_depreciation").Name%>"><%=rsGAD("sourcecode_depreciation")%></td>
				<td align="left" title="<%=rsGAD("accumulated_depreciation").Name%>"><%=rsGAD("accumulated_depreciation")%></td>
				<td align="left" title="<%=rsGAD("jobcode_intraco").Name%>"><%=rsGAD("jobcode_intraco")%></td>
				<td align="left" title="<%=rsGAD("sourcetype_intraco").Name%>"><%=rsGAD("sourcetype_intraco")%></td>
				<td align="left" title="<%=rsGAD("sourcecode_intraco").Name%>"><%=rsGAD("sourcecode_intraco")%></td>
				<td align="left" title="<%=rsGAD("intraco_amount").Name%>"><%=rsGAD("intraco_amount")%></td>
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
	  '.chosen-select-width'     : {width:"95%"}
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
function getselectedtransportask() {
	var stransportask = document.getElementById("selectedtransport");
	stransportask.value = document.getElementById("fa_transportask").value;
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
	gettransportaskno();
	getdatetransportask();
	getselectedtransportask();
	// getAssetInfo();
	// getAssetWriteOffInfo();
	
});
</script>
</body>
</html>
