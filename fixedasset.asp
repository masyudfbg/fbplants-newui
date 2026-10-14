<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<title>FIxed Asset | iPlaS</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
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
var getAssetTaxInfo = function () {
	// build up the post string when passing variables to the server side page
	
	var assettaxcode = document.getElementById("assettaxcode").value;
	var PostStr = 'assettaxcode='+assettaxcode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getassettaxgroupinfo.asp', 'POST', PostStr, showAssetTaxInfo);
}



// The function for handling the response from the server
var showAssetTaxInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.salescontractno);
    // update the Div to show the result from the server
	document.getElementById("assettaxname").value = s.assettaxname;
	// document.getElementById("assetgroupname").value = s.assetgroupname;
	// document.getElementById("unitofmeasurement").value = s.unitofmeasurement;
	// document.getElementById("depreciationrange").value = s.depreciationrange;
}

// launched from button click 
var getAssetInfo = function () {
	// build up the post string when passing variables to the server side page
	
	var assetcode = document.getElementById("assetcode").value;
	var PostStr = 'assetcode='+assetcode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getassetandassetgroupinfo.asp', 'POST', PostStr, showAssetInfo);
}



// The function for handling the response from the server
var showAssetInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.salescontractno);
    // update the Div to show the result from the server
	document.getElementById("assetgroupcode").value = s.assetgroupcode;
	document.getElementById("assetgroupname").value = s.assetgroupname;
	document.getElementById("unitofmeasurement").value = s.unitofmeasurement;
	document.getElementById("depreciationrange").value = s.depreciationrange;
}

// launched from button click 
var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfcbabasedoncompany.asp', 'POST', PostStr, showlistfcba);
}



// The function for handling the response from the server
var showlistfcba = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("fcba").innerHTML = "";
	document.getElementById("fcba").innerHTML = response;
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
fcname = request("fcname")
creationdate = request("creationdate")
assetcode = request("assetcode")
description = request("description")
isactive = request("isactive")
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = request("selectedfcba")

assettaxcode = request("assettaxcode")
assettaxname = request("assettaxname")

acquisitionvalue = request("acquisitionvalue")
additionalvalue = request("additionalvalue")
totalvalue = request("totalvalue")

if ""&creationdate = "" then creationdate = TextDate(date)

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM fixedasset WHERE fccode = '" & fccode & "'"
		connStr.execute(q)
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		hasdata = 0
		q = " SELECT COUNT(1) hasdata FROM fixedasset WHERE fccode = '" & fccode & "'"
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
			q = " INSERT INTO fixedasset ( fccode" & _
				"                          , fcname" & _
				"                           , creationdate" & _
				"                           , assetcode" & _
				"                           , description" & _
				"                           , isactive" & _
				"                           , assettaxcode" & _
				"                           , assettaxname" & _
				"                           , fcentry" & _
				"                           , fcedit" & _
				"                           , fcip" & _
				"                           , fcba" & _
				"                           , lastupdate" & _
				"                           , lasttime)" & _
				" VALUES ( '" & fccode & "'" & _
				"          , '" & fcname & "'" & _
				"          , TO_DATE('" & creationdate & "', 'mm/dd/yyyy')" & _
				"          , '" & assetcode & "'" & _
				"          , '" & description & "'" & _
				"          , '" & isactive & "'" & _
				"          , '" & assettaxcode & "'" & _
				"          , '" & assettaxname & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , '" & fcba & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

			'response.write q
			connStr.execute(q)
		end if
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q = " UPDATE fixedasset" & _
            " SET fcname = '" & fcname & "'" & _
			"     , creationdate = TO_DATE('" & creationdate & "', 'mm/dd/yyyy')" & _
			"     , assetcode = '" & assetcode & "'" & _
			"     , description = '" & description & "'" & _
			"     , isactive = '" & isactive & "'" & _
			"     , assettaxcode = '" & assettaxcode & "'" & _
			"     , assettaxname = '" & assettaxname & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , fcba = '" & fcba & "'" & _
			"     , lastupdate = SYSDATE" & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
			" WHERE fccode = '" & fccode & "'" 
			
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
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

 '-------------ISI Payment Form
  %>
  
<form name="frmEntry" method="post" action="fixedasset.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
	<TR> 
		<td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE FIXED ASSET</strong></font></div></td>
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
		<a href="Printfixedasset.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
	</tr>
	</table>  
</form>
<%
if cmdView = "True" Then
	q = " SELECT fa.fccode" & _
	    "        , fa.fcname" & _
	    "        , bu.fccompanycode" & _
	    "        , fa.fcba" & _
	    "        , fa.creationdate" & _
	    "        , fa.assetcode" & _
	    "        , fa.assettaxcode" & _
	    "        , fa.assettaxname" & _
	    "        , ass.fcname assetname" & _
	    "        , ass.assetgroupcode" & _
	    "        , asg.fcname assetgroupname" & _
	    "        , ass.unitofmeasurement" & _
	    "        , asg.depreciationrange" & _
	    "        , NVL( (SELECT NVL(ft.amount,0) amount FROM fixedasset_transaction ft WHERE ft.fixedassetcode = fa.fccode AND ft.transactiontype = 'ACQUISITION'), 0) acquisitionvalue" & _
	    "        , NVL( (SELECT NVL(SUM(ft.amount),0) amount FROM fixedasset_transaction ft WHERE ft.fixedassetcode = fa.fccode AND ft.transactiontype = 'ADDITIONAL'), 0) additionalvalue" & _
	    "        , fa.description" & _
	    "        , fa.isactive" & _
	    "        , DECODE(fa.isactive,'1','TRUE', 'FALSE') isactivedesc" & _
	    " FROM fixedasset fa" & _
	    " LEFT JOIN asset ass ON ass.fccode = fa.assetcode" & _
		" LEFT JOIN assetgroup asg ON asg.fccode = ass.assetgroupcode" & _
		" LEFT JOIN businessunit bu ON fa.fcba = bu.fccode" & _
        " WHERE 1=1"

	if ""&searchdata <> "" then q = q & " AND (UPPER(fa.fccode) LIKE UPPER('%"&searchdata&"%') OR UPPER(fa.fcname) LIKE ('%"&searchdata&"%'))"
	if Request("txtJenis")="Edit" then q = q & " AND fa.fccode = '" & fccode & "'"
	
	' response.write q
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.Open q, connstr
	
	if Request("txtJenis")="Edit" then
		if Not rsGAD.eof then
			fcname = rsGAD("fcname")
			fccompanycode = rsGAD("fccompanycode")
			fcba = rsGAD("fcba")
			selectedfcba = rsGAD("fcba")
			creationdate = rsGAD("creationdate")
			assetcode = rsGAD("assetcode")
			description = rsGAD("description")
			isactive = rsGAD("isactive")
			assettaxcode = rsGAD("assettaxcode")
			assettaxname = rsGAD("assettaxname")
			acquisitionvalue = rsGAD("acquisitionvalue")
			additionalvalue = rsGAD("additionalvalue")
			totalvalue = cdbl(to_nol(rsGAD("acquisitionvalue"))) +cdbl(to_nol(rsGAD("additionalvalue")))
		end if
	end if

%>
   
   <!-- Form Entry And Edit --> 
   
   <form action="fixedasset.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
		
		<tr>
			<td width="20%" align="left">Fixed Asset Code</td>
			<td align="left">
				<input type="text" size="15" maxlength="50" name="fccode" value="<%=fccode%>" <%IF Request("txtJenis")="Edit" then%> readonly="True" style="background-color:#ccc;"<%end if%>><font color="#FF0000">*</font></td>
		</tr>
		
		<tr>
			<td align="left">Fixed Asset Name</td>
			<td align="left"><input type="text" size="50" maxlength="150" name="fcname" value="<%=fcname%>"><font color="#FF0000">*</font></td>		
		</tr>
		
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();">
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
				<select id="fcba" name="fcba" onchange="javascript:getselectedfcba();">
					<option value=""></option>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
		</tr>
		
		<tr>
			<td align="left">Creation Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="creationdate" size="10" value="<%=creationdate%>" autocomplete="off">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Asset Tax Group Code</td>
			<td align="left">
				<select id="assettaxcode" name="assettaxcode" onchange="javascript:getAssetTaxInfo();">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM fixedasset_taxgroup ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if assetcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
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
			<td align="left">Asset Tax Group Name</td>
			<td align="left"><input type="text" size="50" id="assettaxname" name="assettaxname" value="<%=assetgroupcode%>" readonly="true"  style="background-color:#ccc;"></td>
		</tr>
		<tr>
		
		<tr>
			<td align="left">Asset Code</td>
			<td align="left">
				<select id="assetcode" name="assetcode" onchange="javascript:getAssetInfo();">
					<option value=""></option>
					<%
					qlov = "SELECT fccode, fcname FROM asset ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if assetcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
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
			<td align="left">Asset Group Code</td>
			<td align="left"><input type="text" size="50" id="assetgroupcode" name="assetgroupcode" value="<%=assetgroupcode%>" readonly="true"  style="background-color:#ccc;"></td>
		</tr>
		<tr>
			<td align="left">Asset Group Name</td>
			<td align="left"><input type="text" size="50" id="assetgroupname" name="assetgroupname" value="<%=assetgroupname%>" readonly="true"  style="background-color:#ccc;"></td>
		</tr>		
		<tr>
			<td align="left">Uom</td>
			<td align="left">
				<input type="text" size="50" id="unitofmeasurement" name="unitofmeasurement" value="<%=unitofmeasurement%>" readonly="true"  style="background-color:#ccc;">
			</td>
		</tr>		
		<tr>
			<td align="left">Depreciation Range</td>
			<td align="left">
				<input type="text" size="50" id="depreciationrange" name="depreciationrange" value="<%=depreciationrange%>" readonly="true"  style="background-color:#ccc;">
			</td>
		</tr>
		<tr>
			<td align="left">Acquisition Value</td>
			<td align="left">
				<input type="text" size="50" name="acquisitionvalue" value="<%=FormatNumber(acquisitionvalue,0)%>" readonly="true"  style="background-color:#ccc;">
			</td>
		</tr>		
		<tr>
			<td align="left">Additional Value</td>
			<td align="left">
				<input type="text" size="50" name="additionalvalue" value="<%=FormatNumber(additionalvalue,0)%>" readonly="true"  style="background-color:#ccc;">
			</td>
		</tr>
		<tr>
			<td align="left">Total Value</td>
			<td align="left">
				<input type="text" size="50" name="totalvalue" value="<%=FormatNumber(totalvalue,0)%>" readonly="true"  style="background-color:#ccc;">
			</td>
		</tr>
		<tr>	
			<td align="left">Description</td>
			<td align="left" width="100%" ><input type="text" size="50" maxlength="100" name="description" value="<%=description%>"></td>		
		</tr>
		
		<tr>
			<td align="left">Is Active</td>
			<td align="left">
				<select name="isactive">
					<option value="1" <%if isactive = "1" then%>selected<%end if%>>TRUE</option>
					<option value="0" <%if isactive = "0" then%>selected<%end if%>>FALSE</option>
				</select>
				<font color="#FF0000">*</font>
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
   <table class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<TD align="left">No</TD>
		<TD align="center" colspan="2"></TD>
		<TD align="left">Fixed Asset Code</TD>
		<TD align="left">Fixed Asset Name</TD>
		<TD align="left">Company</TD>
		<TD align="left">Business Unit</TD>
		<TD align="left">Creation Date</TD>
		<TD align="left">Asset Tax Code</TD>
		<TD align="left">Asset Tax Name</TD>
		<TD align="left">Asset Code</TD>
		<TD align="left">Asset Name</TD>
		<TD align="left">Asset Group Code</TD>
		<TD align="left">Asset Group Name</TD>
		<TD align="left">Uom</TD>
		<TD align="left">Depreciation Range</TD>
		<TD align="left">Acquisition Value</TD>
		<TD align="left">Additional Value</TD>
		<TD align="left">Total Value</TD>
		<TD align="left">Description</TD>
		<TD align="left">Is Active</TD>
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
					<a href="fixedasset.asp?fccode=<%=rsGAD("fccode")%>&searchdata=<%=searchdata%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD("fccode")%>" border="0"></a>		
				</td>
				<td valign="middle">
					<a href="fixedasset.asp?fccode=<%=rsGAD("fccode")%>&searchdata=<%=searchdata%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD("fccode")%>" border="0">
					</a>
				</td>
		   
				<td align="left" title="<%=rsGAD("fccode").Name%>"><%=rsGAD("fccode")%></td>
				<td align="left" title="<%=rsGAD("fcname").Name%>"><%=rsGAD("fcname")%></td>
				<td align="left" title="<%=rsGAD("fccompanycode").Name%>"><%=rsGAD("fccompanycode")%></td>
				<td align="left" title="<%=rsGAD("fcba").Name%>"><%=rsGAD("fcba")%></td>
				<td align="left" title="<%=rsGAD("creationdate").Name%>"><%=rsGAD("creationdate")%></td>
				<td align="left" title="<%=rsGAD("assettaxcode").Name%>"><%=rsGAD("assettaxcode")%></td>
				<td align="left" title="<%=rsGAD("assettaxname").Name%>"><%=rsGAD("assettaxname")%></td>
				<td align="left" title="<%=rsGAD("assetcode").Name%>"><%=rsGAD("assetcode")%></td>
				<td align="left" title="<%=rsGAD("assetname").Name%>"><%=rsGAD("assetname")%></td>
				<td align="left" title="<%=rsGAD("assetgroupcode").Name%>"><%=rsGAD("assetgroupcode")%></td>
				<td align="left" title="<%=rsGAD("assetgroupname").Name%>"><%=rsGAD("assetgroupname")%></td>
				<td align="left" title="<%=rsGAD("unitofmeasurement").Name%>"><%=rsGAD("unitofmeasurement")%></td>
				<td align="left" title="<%=rsGAD("depreciationrange").Name%>"><%=rsGAD("depreciationrange")%></td>
				<td align="left" title="<%=rsGAD("acquisitionvalue").Name%>"><%=FormatNumber(rsGAD("acquisitionvalue"),0)%></td>
				<td align="left" title="<%=rsGAD("additionalvalue").Name%>"><%=FormatNumber(rsGAD("additionalvalue"),0)%></td>
				<td align="left" title="Total Value"><%=FormatNumber( cdbl(to_nol(rsGAD("acquisitionvalue")))+cdbl(to_nol(rsGAD("additionalvalue"))), 0)%></td>
				<td align="left" title="<%=rsGAD("description").Name%>"><%=rsGAD("description")%></td>
				<td align="left" title="<%=rsGAD("isactivedesc").Name%>"><%=rsGAD("isactivedesc")%></td>

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
	getlistfcba();
	getAssetInfo();
});
</script>
</body>
</html>
