<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
	<title>ePlaS - Authorize Level</title>
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

<style>
.close-button {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 30px;
  height: 30px;
  position: relative;
  padding: 0 1px 0 0;
  border: none;
  -webkit-border-radius: 1em;
  border-radius: 1em;
  font: normal 8em/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #6d6d6d;
  cursor:pointer;
}

.close-button:hover {background: #f72c2c;}

.close-button::before {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 5px;
  height: 20px;
  position: absolute;
  content: "";
  top: 5px;
  left: 13px;
  border: none;
  font: normal 100%/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #ffffff;
  text-shadow: none;
  -webkit-transform: rotateZ(45deg)   ;
  transform: rotateZ(45deg)   ;
}

.close-button::after {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 5px;
  height: 20px;
  position: absolute;
  content: "";
  top: 5px;
  left: 13px;
  border: none;
  font: normal 100%/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #ffffff;
  text-shadow: none;
  -webkit-transform: rotateZ(-45deg)   ;
  transform: rotateZ(-45deg)   ;
}
</style>

<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %>>
	<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;">
<%
'--Awal--

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

fccode = Request("fccode")
fcname = Request("fcname")
document_type = Request("document_type")
document_classification = Request("document_classification")
startlevel = Request("startlevel")
finallevel = Request("finallevel")
description = Request("description")
isactive = Request("isactive")

fcusercode = Request("fcusercode")
isactive_user = Request("isactive_user")
lineno = Request("lineno")

txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

if "" & status = "" then status = "Planned"
	
			
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" OR Request("txtJenis") = "Confirmed" OR Request("txtJenis") = "Cancelled" OR Request("txtJenis") = "Executed" OR Request("txtJenis") = "RePlanned" Then
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		
		'Delete data from table recurringsettlementpayment
		q = " DELETE FROM authorizelevel_users" & _
			" WHERE authorizelevelcode = '" & fccode & "'"

		'Response.Write q
		connStr.execute(q)
		
		'Delete data from table recurringsettlement
		q = " DELETE FROM authorizelevel" & _
			" WHERE fccode = '" & fccode & "'"

		'Response.Write q
		connStr.execute(q)
		
		if err.number = 0 then connStr.CommitTrans
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		connstr.BeginTrans
		q = " DELETE FROM authorizelevel_users" & _
			" WHERE authorizelevelcode = '" & fccode & "'" & _
			"       AND lineno = '" & lineno & "'"
		'Response.Write q
		connStr.execute(q)
		if err.number = 0 then connStr.CommitTrans
		cmdView	= "True"				
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans
		hasheader = 0
		oSql = " SELECT COUNT(1) hasheader" & _
			   " FROM authorizelevel" & _
			   " WHERE fccode = '" & fccode & "'"
		set oRs = server.CreateObject("Adodb.recordset")
		oRs.Open oSql, connstr
		if Not oRs.Eof then
			hasheader = cdbl(to_nol(oRs("hasheader")))
		End if
		oRs.Close
		Set oRs = Nothing
		
		if hasheader = 0 then
			q = " INSERT INTO authorizelevel ( fccode" & _
			    "                              , fcname" & _
			    "                              , startlevel" & _
			    "                              , finallevel" & _
			    "                              , description" & _
			    "                              , isactive" & _
			    "                              , fcentry" & _
			    "                              , fcedit" & _
			    "                              , fcip" & _
			    "                              , lastupdate" & _
			    "                              , lasttime)" & _
			    " VALUES ( '" & fccode & "'" & _
				"          , '" & fcname & "'" & _
				"          , '" & startlevel & "'" & _
				"          , '" & finallevel & "'" & _
				"          , '" & description & "'" & _
				"          , '" & isactive & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"

			'Response.Write q
			connStr.execute(q)
		end if
		
		isduplicateuser = 0
		oSql = " SELECT COUNT(1) isduplicateuser" & _
			   " FROM authorizelevel_users" & _
			   " WHERE authorizelevelcode = '" & fccode & "'" & _
			   "       AND fcusercode = '" & fcusercode & "'"
		set oRs = server.CreateObject("Adodb.recordset")
		oRs.Open oSql, connstr
		if Not oRs.Eof then
			isduplicateuser = cdbl(to_nol(oRs("isduplicateuser")))
		End if
		oRs.Close
		Set oRs = Nothing
		
		if isduplicateuser = 0 then
		
			q = " INSERT INTO authorizelevel_users ( authorizelevelcode" & _
				"                                    , lineno" & _
				"                                    , fcusercode" & _
				"                                    , isactive" & _
				"                                    , fcentry" & _
				"                                    , fcedit" & _
				"                                    , fcip" & _
				"                                    , lastupdate" & _
				"                                    , lasttime)" & _
				" VALUES ( '" & fccode & "'" & _
				"          , ( SELECT NVL(MAX(lineno),0) FROM authorizelevel_users WHERE authorizelevelcode = '"&fccode&"' ) +1" & _
				"          , '" & fcusercode & "'" & _
				"          , '" & isactive_user & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"

			'Response.Write q
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "user exist, please check your input"
		end if
		
		if err.number = 0 then connStr.CommitTrans
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		connstr.BeginTrans
		q = " UPDATE authorizelevel_users" & _
		    " SET fcusercode = '" & fcusercode & "'" & _
			"     , isactive = '" & isactive_user & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , lastupdate = SYSDATE" & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
			" WHERE authorizelevelcode = '" & fccode & "'" & _
			"       AND lineno = '" & lineno & "'"

		'response.Write(q)
		connStr.execute(q)
		
		if err.number = 0 then connStr.CommitTrans
		cmdView	= "True"
    else
	
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")		    
End if

set oRs = server.CreateObject("Adodb.recordset")
oSql = " SELECT *" & _
       " FROM authorizelevel" & _
       " WHERE fccode = '" & fccode & "'"

oRs.Open oSql, connstr
if Not oRs.Eof then
	fcname = oRs("fcname")
	startlevel = oRs("startlevel")
	finallevel = oRs("finallevel")
	description = oRs("description")
	isactive = oRs("isactive")
End if
oRs.Close
Set oRs = Nothing

arr_user_eof = true
q = " SELECT fccode FROM users_tab WHERE TRUNC(SYSDATE) BETWEEN TRUNC(valid_from) AND TRUNC(valid_until) AND isactive = 'TRUE' ORDER BY fccode"
'response.Write q
set rswipfinal = server.CreateObject("Adodb.recordset")
rswipfinal.Open q, connstr
if not rswipfinal.eof then
	arr_user = rswipfinal.getrows()
	arr_user_eof = false
end if
rswipfinal.close : set rswipfinal = nothing


if err.number <> 0 then
	connStr.RollbackTrans
	response.Write  "<div class = ""error"">" & err.description & "</div>"
end if
%>
<form name="frmEntry" method="post" action="authorizelevel.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Authorize Level Code</td>
			<td align="left">
				<input type="text" maxlength="50" name="fccode" size="50" value="<%=fccode%>" autocomplete="off">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Authorize Level Name</td>
			<td align="left">
				<input type="text" maxlength="150" name="fcname" size="50" value="<%=fcname%>" autocomplete="off">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Start Level</td>
			<td align="left">
				<select id="startlevel" name="startlevel">
					<option value="FALSE" <%if startlevel = "FALSE" then%> selected <%end if%>>FALSE</option>
					<option value="TRUE" <%if startlevel = "TRUE" then%> selected <%end if%>>TRUE</option>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Final Level</td>
			<td align="left">
				<select id="finallevel" name="finallevel">
					<option value="FALSE" <%if finallevel = "FALSE" then%> selected <%end if%>>FALSE</option>
					<option value="TRUE" <%if finallevel = "TRUE" then%> selected <%end if%>>TRUE</option>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Description</td>
			<td align="left">
				<input type="text" maxlength="150" name="description" size="50" value="<%=description%>" autocomplete="off">
			</td>
		</tr>
		<tr>
			<td align="left">Is Active</td>
			<td align="left">
				<select id="isactive" name="isactive">
					<option value="FALSE" <%if isactive = "FALSE" then%> selected <%end if%>>FALSE</option>
					<option value="TRUE" <%if isactive = "TRUE" then%> selected <%end if%>>TRUE</option>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">
				&nbsp;&nbsp;
				<a href="authorizelevel.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fccode=<%=fccode%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=documentno%>" border="0">
				</a>
					<!--
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
					-->
			</td>
		</tr>
	</table>  
</form>
<%
If cmdView = "True" and "" & fccode <> "" and "" & fcname <> "" and "" & startlevel <> "" and "" & finallevel <> "" and "" & isactive <> "" Then
	arr_authuser_eof = true
	q = " SELECT lineno, fcusercode, isactive FROM authorizelevel_users WHERE authorizelevelcode = '" & fccode & "' ORDER BY lineno"
	'response.Write q
	set rswipfinal = server.CreateObject("Adodb.recordset")
	rswipfinal.Open q, connstr
	if not rswipfinal.eof then
		arr_authuser = rswipfinal.getrows()
		arr_authuser_eof = false
	end if
	rswipfinal.close : set rswipfinal = nothing
%>
<form action="authorizelevel.asp" method="post" name="eForm">
	<table class="data" width="100%">
	<tr>
		<td align="left" style="width:20%;">User</td>
		<td align="left">
			<%if txtJenis <> "Edit" then%>
			<select id="fcusercode" name="fcusercode" class="chosen-select" style="width:75%;">
				<%if not arr_user_eof then%>
				<%for xyyy=0 to ubound(arr_user,2)%>
				<option value="<%=arr_user(0,xyyy)%>" <%if fcusercode = arr_user(0,xyyy) then%> selected <%end if%>><%=arr_user(0,xyyy)%></option>
				<%next%>
				<%end if%>
			<select>
			<%else%>
			<input type="text" size="30" name="fcusercode" value=<%=fcusercode%> readonly="true" style="background-color:#ccc;">
			<%end if%>
			<font color="#FF0000">&nbsp;*</font>
		</td>
	</tr>
	<tr>
		<td align="left">Is Active</td>
		<td align="left">
			<select id="isactive_user" name="isactive_user">
				<option value="FALSE" <%if isactive_user = "FALSE" then%> selected <%end if%>>FALSE</option>
				<option value="TRUE" <%if isactive_user = "TRUE" then%> selected <%end if%>>TRUE</option>
			<select>
			<font color="#FF0000">&nbsp;*</font>
		</td>
	</tr>
	<tr>
			<td colspan="11" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="fccode" value="<%=fccode%>">
			<input type="hidden" name="fcname" value="<%=fcname%>">
			<input type="hidden" name="startlevel" value="<%=startlevel%>">
			<input type="hidden" name="finallevel" value="<%=finallevel%>">
			<input type="hidden" name="description" value="<%=description%>">
			<input type="hidden" name="isactive" value="<%=isactive%>">
			<input type="hidden" name="lineno" value="<%=lineno%>">
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
			<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
	</table>
	<table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">No</td>
		<td align="left" colspan="2">&nbsp;</td>
		<td align="left">User</td>
		<td align="left">Is Active</td>
	</tr>
	<%if not arr_authuser_eof then%>
		<%
		no = 0
		%>
		<%for xyyy=0 to ubound(arr_authuser,2)%>
		<%if no  mod 2 = 1 then cColor = "#EEEEEE" else cColor = "#E5E5E5"%>
		<%no = no +1%>
			<tr bgcolor="<%=cColor%>">
				<td align="left" valign="top"><%=no%></td>
				<td valign="middle">
				   <a href="authorizelevel.asp?fccode=<%=fccode%>&lineno=<%=arr_authuser(0,xyyy)%>&fcusercode=<%=arr_authuser(1,xyyy)%>&isactive_user=<%=arr_authuser(2,xyyy)%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=arr_authuser(1,xyyy)%>" border="0">
				   </a>
			   </td>
			   <td valign="middle" width="10">
					<a href="authorizelevel.asp?fccode=<%=fccode%>&lineno=<%=arr_authuser(0,xyyy)%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=arr_authuser(1,xyyy)%>" border="0">
					</a>
			   </td>
			   <td align="left" valign="top"><%=arr_authuser(1,xyyy)%></td>
			   <td align="left" valign="top"><%=arr_authuser(2,xyyy)%></td>
			</tr>
		<%next%>
	<%end if%>
				
</form>
<%End if%>
<!---------------akhir!--------------->
</div>
</div>
<!--#include file="bottom.jsp"-->	
	<script src="plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
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
		
	</script>
	
	<script src="/js/jquery.inputmask.bundle.min.js"></script>
	<script src="/js/inputmask.binding.js"></script>
	<script src="plugins/tabcontent/tabcontent.js" type="text/javascript"></script>
	
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
		// var PostStr = "variable=value&variable2=value2";
		var PostStr = "documentno="+document.getElementById("documentno").value;
		
		// use the generic function to make the request
		doAJAXCall('ajax/recurringdetail.asp', 'POST', '' + PostStr + '', showMessageResponse);
	}

	// The function for handling the response from the server
	var showMessageResponse = function (oXML) { 
		
		// get the response text, into a variable
		var response = JSON.parse(oXML.responseText);
		//alert(response.finishdate);
		
		
		document.getElementById("finishdate").value = response.finishdate;
		document.getElementById("accountingyear").value = response.accountingyear;
		
		
		//document.getElementById("finishdate").value = JSON.stringify(response.finishdate);
	};

	
	function getnumberofruns(){
	
	var start_date = new Date(document.getElementById("startingdate").value);
	var run_date = document.getElementById("rundate")
	var numberofruns = document.getElementById("numberofrun");
	var end_date = new Date(document.getElementById("finishdate").value);
	total_months = (end_date.getFullYear() - start_date.getFullYear())*12 + (end_date.getMonth() - start_date.getMonth());
	if (total_months < 0){
		numberofruns.value = 0;
	} else {
		numberofruns.value = total_months +1;
	}
	
	
	run_date.value = document.getElementById("startingdate").value;
	
	//alert(numberofruns);
}
	</script>
	
	<script type="text/javascript">
		function saveValue(e){
            var id = e.id;  // get the sender's id to save it . 
            var val = e.value; // get the value. 
            sessionStorage.setItem(id, val);// Every time user writing something, the localStorage's value will override . 
        }
		
		function addpayment(){
			var counter = document.getElementById("maxcounter");
			var x = 1;
			num = x + parseInt(counter.value);
			
			option = "'option"+num+"'";
			$("#addpayment").append('<div id="option'+num+'">'+
									  '<div class="section group">'+
										' <div class="col span_5_of_12">'+
											' <select id="cashjob'+num+'" name="cashjob'+num+'" class="chosen-select" style="width:100%;" onchange="saveValue(this);">'+
												'<option value=""></option>'+
												<%if not arr_job_eof then%>
													<%For xx = LBound(arr_job, 2) To UBound(arr_job, 2)%>
													'<option value="<%=arr_job(0, xx)%>" <% If request("") = arr_job(0, xx) then%> Selected <%End if%>><%=arr_job(0, xx) & " : " & arr_job(1, xx)%></option>' +
													<%Next%>
												<%End if%>
											' </select>'+
										' </div>'+
										' <div class="col span_5_of_12">'+
											' <input style="width:100%;"type="text" id="cashamount'+num+'" name="cashamount'+num+'" value="" data-inputmask="&#39;alias&#39;: &#39;numeric&#39;, &#39;groupSeparator&#39;: &#39;,&#39;, &#39;autoGroup&#39;: true, &#39;digits&#39;: 0, &#39;digitsOptional&#39;: false, &#39;placeholder&#39;: &#39;0&#39;" style="text-align: right;" onkeyup="saveValue(this);">' +
										' </div>'+
										' <div class="col span_2_of_12">'+
											' <div class="close-button" onclick="delete_element(' + option + ');"></div>'+
										' </div>'+
									  ' </div>'+
									' </div>');
			counter.value = num;
			
			var script = document.createElement('script');
			script.src = '/js/inputmask.binding.js';
			document.body.appendChild(script);
		
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
		
		function delete_element(id){
			Element.prototype.remove = function() {
				this.parentElement.removeChild(this);
			}
			NodeList.prototype.remove = HTMLCollection.prototype.remove = function() {
				for(var i = this.length - 1; i >= 0; i--) {
					if(this[i] && this[i].parentElement) {
						this[i].parentElement.removeChild(this[i]);
					}
				}
			}
			document.getElementById(id).remove();
		}
	</script>
</body>
</html>
