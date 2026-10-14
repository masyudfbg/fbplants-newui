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

function loaddepreciationrange(){
	var method = document.getElementById("depreciationmethod").value;
	if (method === 'STRAIGHT LINE') {
		removeOptions(document.getElementById('depreciationrange'));
		populateSelect('depreciationrange', 1, 20);
	} else {
		removeOptions(document.getElementById('depreciationrange'));
		populateSelect('depreciationrange', 0, 0);
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
// Untuk Disable Enable Field Text Based On Select Option
$(document).ready(function() {
 	$('#fctype').change(function() {
  	if( $(this).val() === 'ADDITION') {
   		document.getElementById("targetcode").val() ;
   		document.getElementById("jobcode").val();
   		document.getElementById("targettype").value = 'FA';
   		document.getElementById("amount").val();
    } else {       
      	$('#jobcode').val(null);
      	$('#targetcode').val(null);
      	$('#targettype').val(null);
      	$('#amount').val(null);
    }
  });
 
});


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
var getAdjustmentInfo = function () {
	// build up the post string when passing variables to the server side page
	
	var fixedassetcode = document.getElementById("fixedassetcode").value;
	var PostStr = 'fixedassetcode='+fixedassetcode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfixedadjustmentinfo.asp', 'POST', PostStr, showAdjustmentInfo);
}



// The function for handling the response from the server
var showAdjustmentInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.salescontractno);
    // update the Div to show the result from the server
	document.getElementById("amount").value = s.amount;

	var fctype = document.getElementById("fctype").value;
	if (fctype === 'ADDITION') {
		document.getElementById("amount").value = s.amount;
	} else {
		document.getElementById("amount").value = "";
	}
	return null;
}

// launched from button click 
var getFixedAssetInfo = function () {
	// build up the post string when passing variables to the server side page
	
	var fixedassetcode = document.getElementById("fixedassetcode").value;
	var PostStr = 'fixedassetcode='+fixedassetcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfixedassetinfo.asp', 'POST', PostStr, showAssetInfo);
}

// The function for handling the response from the server
var showAssetInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.fixedassetname);
    // update the Div to show the result from the server
    document.getElementById("fixedassetname").value = s.fixedassetname;
    // document.getElementById("targetcode").value = s.targetcode;
    // document.getElementById("jobcode").value = s.jobcode;

    var fctype = document.getElementById("fctype").value;
	if (fctype === 'ADDITION') {
		document.getElementById("targetcode").value = s.targetcode;
   		document.getElementById("jobcode").value = s.jobcode;
   		document.getElementById("targettype").value = 'FA';
	} else {
		document.getElementById("targetcode").value = "";
    	document.getElementById("jobcode").value = "";
    	document.getElementById("targettype").value = "";
	}
	return null;
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
fixedassetcode = request("fixedassetcode")
transdate = request("transdate")
fixedassetname = request("fixedassetname")
jobcode = request("jobcode")
targettype = request("targettype")
targetcode = request("targetcode")
adjustment_amount = request("adjustment_amount")
fcba = request("fcba")
description = request("description")
fctype = request("fctype")
assettaxname = request("assettaxname")
acquisitionvalue = request("acquisitionvalue")
additionalvalue = request("additionalvalue")
totalvalue = request("totalvalue")
assetgroupcode = request("assetgroupcode")
assetgroupname = request("assetgroupname")
unitofmeasurement = request("unitofmeasurement")
adjustment_monthdepreciation = request ("adjustment_monthdepreciation")
remarks = request("remarks")

if ""&transdate = "" then transdate = TextDate(date)

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM fixedasset_adjustment WHERE fccode = '" & fccode & "' and fcba = '"& Session("fcba") &"'"
		connStr.execute(q)
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		hasdata = 0
		q = " SELECT COUNT(1) hasdata FROM fixedasset_adjustment WHERE fccode = '" & fccode & "'"
		Set rscekpk = Server.CreateObject("ADODB.Recordset")
		rscekpk.Open q, connstr
		
		if Not rscekpk.eof then
			hasdata = cdbl(to_nol(rscekpk("hasdata")))
		else
			hasdata = 0
		end if
		
		if hasdata = 1 then
			err.number = 1001
			err.description = "Fixed Asset : " & fccode & " , already exist"
		else
			q = " INSERT INTO fixedasset_adjustment ( fccode" & _
				"                           , fctype" & _
				"                           , transdate" & _
				"                           , description" & _
				"                           , fixedassetcode" & _
				"                           , fixedassetname" & _
				"                           , jobcode" & _
				"                           , targettype" & _
				"                           , targetcode" & _
				"                           , adjustment_amount" & _
				"                           , remarks" & _
				"                           , adjustment_monthdepreciation" & _
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
				"          , '" & fixedassetcode & "'" & _
				"          , '" & fixedassetname & "'" & _
				"          , '" & jobcode & "'" & _
				"          , '" & targettype & "'" & _
				"          , '" & targetcode & "'" & _
				"		   , replace('" & adjustment_amount & "',',','')" & _
				"          , '" & remarks & "'" & _
				"          , '" & adjustment_monthdepreciation & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , '" & Session("fcba") & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

			' response.write q
			connStr.execute(q)
		end if
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q = " UPDATE fixedasset_adjustment" & _
            " SET  transdate = TO_DATE('" & transdate & "', 'mm/dd/yyyy')" & _
			"     , description = '" & description & "'" & _
			"     , fixedassetcode = '" & fixedassetcode & "'" & _
			"     , fixedassetname = '" & fixedassetname & "'" & _
			"     , jobcode = '" & jobcode & "'" & _
			"     , targettype =  replace('" & targettype & "',',','')" & _
			"     , targetcode = '" & targetcode & "'" & _
			"     , adjustment_amount =  replace('" & adjustment_amount & "',',','')" & _
			"     , remarks = '" & remarks & "'" & _
			"     , adjustment_monthdepreciation = '" & adjustment_monthdepreciation & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , lastupdate = SYSDATE" & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
			" WHERE fcba = '"& Session("fcba") &"' and fccode = '" & fccode & "'" 
			
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

if err.number <> 0 then
   response.Write  "<div class='error'>" & err.description & "</div>"
End if

 '-------------ISI Payment Form
  %>
  
<form name="frmEntry" method="post" action="fixedasset_adjustment.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
	<TR> 
		<td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE FIXED ASSET ADJUSTMENT</strong></font></div></td>
	</TR>
	<tr>
		<td align="left" width="20%">Search by Code or Name</td>
		<td align="left"><input type="text" size="30" name="searchdata" value="<%=searchdata%>"  ></td>
	</tr>	
	<tr>
		<td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;
		<a href="Printfixedasset_adjustment.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
	</tr>
	</table>  
</form>
<%
if cmdView = "True" Then
	q = " SELECT fad.fccode," & _
		"       bu.fccompanycode," & _
		"       fad.fcba," & _
		"       fad.transdate," & _
		"       fad.jobcode," & _
		"       fad.targettype," & _
		"       fad.targetcode," & _
		"       fad.adjustment_monthdepreciation," & _
		"       fad.fctype," & _
		"       fad.fixedassetcode," & _
		"       fad.fixedassetname," & _
		"       fad.adjustment_amount," & _
		"       fad.remarks," & _
		"       fad.description" & _
		"  FROM fixedasset_adjustment fad" & _
		"       LEFT JOIN businessunit bu ON fad.fcba = bu.fccode" & _
        " WHERE 1=1"

	if ""&searchdata <> "" then q = q & " AND (UPPER(fad.fccode) LIKE UPPER('%"&searchdata&"%'))"
	if Request("txtJenis")="Edit" then q = q & " AND fad.fccode = '" & fccode & "' and fcba= '"& Session("fcba") &"'"
	
	' response.write q
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.Open q, connstr
	
	if Request("txtJenis")="Edit" then
		if Not rsGAD.eof then
			fixedassetname = rsGAD("fixedassetname")
			fcba = rsGAD("fcba")
			selectedfcba = rsGAD("fcba")
			transdate = rsGAD("transdate")
			jobcode = rsGAD("jobcode")
			targettype = rsGAD("targettype")
			targetcode = rsGAD("targetcode")
			adjustment_monthdepreciation = rsGAD("adjustment_monthdepreciation")
			fctype = rsGAD("fctype")
			fixedassetcode = rsGAD("fixedassetcode")
			adjustment_amount = rsGAD("adjustment_amount")
			description = rsGAD("description")
			remarks = rsGAD("remarks")
		end if
	end if

%>
   
   <!-- Form Entry And Edit --> 
   
   <form action="fixedasset_adjustment.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
		
		<tr>
			<td width="20%" align="left">Fixed Asset Adjustement Code</td>
			<td align="left">
				<input type="text" size="25" maxlength="50" name="fccode" value="<%=fccode%>" <%IF Request("txtJenis")="Edit" then%> readonly="True" style="background-color:#ccc;"<%end if%>><font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="left">Fixed Asset Adjustment Type</td>
			<td align="left">
				<select id ="fctype" name="fctype" class="chosen-select" style="width: 20%" id="fctype">
					<option></option>
					<option value="ADDITION" <%if fctype = "ADDITION" then%>selected<%end if%>>ADDITION</option>
					<option value="REDUCTION" <%if fctype = "REDUCTION" then%>selected<%end if%>>REDUCTION</option>
				</select>
				<font color="#FF0000">*</font>
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
			<td align="left" width="100%" >
				<textarea name="description" cols="50"><%=description%></textarea>
			</td>		
		</tr>
		<tr>
			<td align="left">Fixed Asset Code</td>
			<td align="left">
				<select id="fixedassetcode" name="fixedassetcode" onchange="javascript:getFixedAssetInfo();javascript:getAdjustmentInfo();" class="chosen-select" style="width: 30%">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM fixedasset where fcba = '"& Session("fcba") &"' ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if fixedassetcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%></option>
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
			<td align="left">Fixed Asset Name</td>
			<td align="left"><input type="text" size="50" maxlength="150" id="fixedassetname" name="fixedassetname" value="<%=fixedassetname%>" style="background-color: #ccc;" readonly></td>		
		</tr>
		<tr>
			<td align="left">Jobcode</td>
			<td align="left"><input type="text" size="50" id="jobcode" name="jobcode" value="<%=jobcode%>"></td>
		</tr>		
		<tr>
			<td align="left">Target Type</td>
			<td align="left"><input type="text" size="50" id="targettype" name="targettype" value="<%=targettype%>"></td>
		</tr>
		<tr>
			<td align="left">Target Code</td>
			<td align="left"><input type="text" size="50" id="targetcode" name="targetcode" value="<%=targetcode%>"></td>
		</tr>		
		<tr>
			<td align="left">Adjustment Amount</td>
			<td align="left">
				<input type="text" size="50" id="amount" name="adjustment_amount" value="<%=adjustment_amount%>">
			</td>
		</tr>
		<tr>
			<td align="left">Adjustment Month Depreciation</td>
			<td align="left">
				<input type="text" size="50" id="adjustment_monthdepreciation" name="adjustment_monthdepreciation" value="<%=adjustment_monthdepreciation%>">
			</td>
		</tr>		
		<tr>
			<td colspan="2" align="left">
				  <input type="Hidden" name="searchdata" value="<%=searchdata%>">
				
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
   <table class="data">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<TD align="left">No</TD>
		<TD align="center" colspan="2"></TD>
		<TD align="center">Company</TD>
		<TD align="center">Business Unit</TD>
		<TD align="center">Fixed Asset Adjustment No</TD>
		<TD align="center">Fixed Asset Code</TD>
		<TD align="center">Fixed Asset Name</TD>
		<TD align="center">Description</TD>
		<TD align="center">Creation Date</TD>
		<TD align="center">FA Adjustment Type</TD>
		<TD align="center">Jobcode</TD>
		<TD align="center">Target Type</TD>
		<TD align="center">Target Code</TD>
		<TD align="center">Adjustment Amount</TD>
		<TD align="center">Adjustment Month Depreciation</TD>
		<TD align="center">Remarks</TD>
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
					<a href="fixedasset_adjustment.asp?fccode=<%=rsGAD("fccode")%>&fixedassetcode=<%=rsGAD("fixedassetcode")%>&searchdata=<%=searchdata%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD("fccode")%>" border="0"></a>		
				</td>
				<td valign="middle">
					<a href="fixedasset_adjustment.asp?fccode=<%=rsGAD("fccode")%>&fixedassetcode=<%=rsGAD("fixedassetcode")%>&searchdata=<%=searchdata%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD("fccode")%>" border="0">
					</a>
				</td>
		   		<td align="center" title="<%=rsGAD("fccompanycode").Name%>"><%=rsGAD("fccompanycode")%></td>
				<td align="center" title="<%=rsGAD("fcba").Name%>"><%=rsGAD("fcba")%></td>
				<td align="center" title="<%=rsGAD("fccode").Name%>"><%=rsGAD("fccode")%></td>
				<td align="center" title="<%=rsGAD("fixedassetcode").Name%>"><%=rsGAD("fixedassetcode")%></td>
				<td align="left" title="<%=rsGAD("fixedassetname").Name%>"><%=rsGAD("fixedassetname")%></td>
				<td align="left" title="<%=rsGAD("description").Name%>"><%=rsGAD("description")%></td>
				<td align="center" title="<%=rsGAD("transdate").Name%>"><%=rsGAD("transdate")%></td>
				<td align="center" title="<%=rsGAD("fctype").Name%>"><%=rsGAD("fctype")%></td>
				<td align="center" title="<%=rsGAD("jobcode").Name%>"><%=rsGAD("jobcode")%></td>
				<td align="center" title="<%=rsGAD("targettype").Name%>"><%=rsGAD("targettype")%></td>
				<td align="center" title="<%=rsGAD("targetcode").Name%>"><%=rsGAD("targetcode")%></td>
				<td align="right" title="<%=rsGAD("adjustment_amount").Name%>"><%=FormatNumber(cdbl(to_nol(rsGAD("adjustment_amount"))),0)%></td>>
				<td align="center" title="<%=rsGAD("adjustment_monthdepreciation").Name%>"><%=rsGAD("adjustment_monthdepreciation")%></td>
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

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
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
	getFixedAssetInfo();
	adjusmenttype();
	getAdjustmentInfo();
});
</script>
</body>
</html>
