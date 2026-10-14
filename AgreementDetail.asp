<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Contract Agreement</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
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

var getlistcontractor = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedcontractorcode = document.getElementById("selectedcontractorcode").value;
	var PostStr = 'fcba='+selectedfcba+'&contractorcode='+selectedcontractorcode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getcontractorbasedonfcba.asp', 'POST', PostStr, showlistcontractor);
}

// The function for handling the response from the server
var showlistcontractor = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("contractorcode").innerHTML = "";
	document.getElementById("contractorcode").innerHTML = response;
}

function JobCodeHint(str)
					{
					var jobname;
					
					if (str=="")
					  {
					  document.getElementById("jobname").innerHTML="";
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  jobname=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  jobname=new ActiveXObject("Microsoft.XMLHTTP");
					  }
				    jobname.onreadystatechange=function()
						  {
						  if (jobname.readyState==4 && jobname.status==200)
							{
							document.getElementById("jobname").innerHTML=jobname.responseText;
							}
						  }
						jobname.open("GET","ajax/MANDAYS.asp?ajaxUoM=jobname&q="+str,true);
						jobname.send();
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
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
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
 Dim CONTRACTORCODE, CONTRACTORDESC, FCCODE,RETENTION,QTY,RATE,LINEno,DUEDATE
 Dim FDDATE,PeriodDate, JOB
 Dim LOCATIONTYPE
 Dim LOCATIONCode,UNIT, OUTPUT, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA
 Dim cmdView,ScriptName
 Dim feecontractor

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
jobcode=request("jobcode")
QTY=request("QTY")
RATE=request("RATE")
LINENO=request("LINENO")

REFERENCE=request("REFERENCE")
REMARKS=request("REMARKS")

agreementcode = Replace(request("agreementcode")," ","")
agreementdate = request("agreementdate")

if request("agreementdate") <>"" then
   agreementdate = request("agreementdate")
 else
   agreementdate = textDate(date)
end if
 
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
selectedcontractorcode = Request("selectedcontractorcode")

CONTRACTORCODE=request("CONTRACTORCODE")
description=request("description")
RETENTION=request("RETENTION")
startdate=request("startdate")
finishdate=request("finishdate")
tax=request("tax")
feecontractor=request("feecontractor")
		
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q = " DELETE FROM contractagreementdetail" & _
			" WHERE agreementcode = '" & Replace(agreementcode," ","") & "'" & _
			"       AND fcba = '" & fcba & "'" & _
			"       AND lineno = '" & lineno & "'"
				
		'Response.Write q
		'ConnStr.execute(q)
		ExecuteAndLog ScriptName, q
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		has_job_in_master = 0
		q = " SELECT COUNT(1) totalrecord FROM job WHERE fccode = '" & jobcode & "' AND fcBA = '" & fcba & "'"
		set rsLPO = Server.CreateObject("ADODB.Recordset")
		rsLPO.Open q, connstr
		if not rsLPO.eof then
			has_job_in_master = cdbl(to_nol(rsLPO("totalrecord")))
		else
			has_job_in_master = 0
		End if
		rsLPO.Close : set rsLPO = Nothing
		
		if has_job_in_master = 0 then
			err.number = 1001
			err.description = "Jobcode """ & jobcode & """ not found, check your data !"
		else

			if feecontractor = "" then 
				feecontractor = 0
			else 
				feecontractor = 1
			end if
	
			q = " SELECT * FROM contractagreement WHERE agreementcode = '" & Replace(agreementcode," ","") & "' AND fcBA = '" & fcba & "'"
			set rsLPO = Server.CreateObject("ADODB.Recordset")
			rsLPO.Open q, connstr
			if rsLPO.eof then
				q = " INSERT INTO contractagreement ( agreementcode" & _
					"                                 , agreementdate" & _
					"                                 , contractorcode" & _
					"                                 , description" & _
					"                                 , retention" & _
					"                                 , startdate" & _
					"                                 , finishdate" & _
					"                                 , tax" & _
					"                                 , fcentry" & _
					"                                 , fcedit" & _
					"                                 , fcip" & _
					"                                 , fcba" & _
					"                                 , lastupdate" & _
					"                                 , lasttime" & _
					"                                 , feecontractor)" & _
					" VALUES ( '" & Replace(agreementcode," ","") & "'" & _
					"          , TO_DATE('" & agreementdate & "', 'mm/dd/yyyy')" & _
					"          , '" & contractorcode & "'" & _
					"          , '" & description & "'" & _
					"          , '" & retention & "'" & _
					"          , TO_DATE('" & startdate & "', 'mm/dd/yyyy')" & _
					"          , TO_DATE('" & finishdate & "', 'mm/dd/yyyy')" & _
					"          , '" & tax & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					"          , '" & fcba & "'" & _
					"          , SYSDATE" & _
					"          , TO_CHAR(SYSDATE, 'hh24:mi') " & _
					"          , '" & feecontractor & "' " & _
					" )"
						
				Response.Write q
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
			End if
			rsLPO.Close : set rsLPO = Nothing
			
				lineno = 0
				q = " SELECT NVL(MAX(lineno),0) lineno FROM contractagreementdetail WHERE agreementcode = '" & Replace(agreementcode," ","") & "' AND fcBA = '" & fcba & "'"
				set rsLPO = Server.CreateObject("ADODB.Recordset")
				rsLPO.Open q, connstr
				if not rsLPO.eof then
					lineno = cdbl(to_nol(rsLPO("lineno")))
				else
					lineno = 0
				End if
				rsLPO.Close : set rsLPO = Nothing
				
				lineno = lineno +1		
				q = " INSERT INTO contractagreementdetail ( agreementcode" & _
					"                                       , lineno" & _
					"                                       , jobcode" & _
					"                                       , qty" & _
					"                                       , rate" & _
					"                                       , reference" & _
					"                                       , remarks" & _
					"                                       , fcentry" & _
					"                                       , fcedit" & _
					"                                       , fcip" & _
					"                                       , fcba" & _
					"                                       , lastupdate" & _
					"                                       , lasttime)" & _
					" VALUES ( '" & Replace(agreementcode," ","") & "'" & _
					"          , '" & lineno & "'" & _
					"          , '" & jobcode & "'" & _
					"          , '" & qty & "'" & _
					"          , '" & rate & "'" & _
					"          , '" & reference & "'" & _
					"          , '" & remarks & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					"          , '" & fcba & "'" & _
					"          , SYSDATE" & _
					"          , TO_CHAR(SYSDATE, 'hh24:mi')" & _
					" )"
						
			''	Response.Write q
				'connStr.execute(q)
				ExecuteAndLog ScriptName, q
			
		end if
			
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q = " UPDATE contractagreementdetail" & _
		    " SET jobcode = '" & jobcode & "'" & _
			"     , qty = '" & QTY & "'" & _
			"     , rate = '" & RATE & "'" & _
			"     , reference = '" & REFERENCE & "'" & _
			"     , remarks = '" & REMARKS & "'" & _
			"     , lastupdate = SYSDATE " & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi') "& _
			"     , fcedit = '" & Session("CodeName") & "' "& _
			" WHERE agreementcode = '" & Replace(agreementcode," ","") & "'" & _
			"       AND fcba = '" & fcba & "'" & _
			"       AND lineno = '" & lineno & "'"
				
		'Response.Write q
		'connStr.execute(q)
		ExecuteAndLog ScriptName, q
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if


txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

q = " SELECT hdr.*, bu.fccompanycode" & _
    " FROM contractagreement hdr" & _
    " LEFT JOIN businessunit bu on bu.fccode = hdr.fcba" & _
    " WHERE hdr.agreementcode = '" & agreementcode & "'"
 
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
    agreementdate = rsVoucher("agreementdate")
	
    fccompanycode = rsVoucher("fccompanycode")
	fcba = rsVoucher("fcba")
	selectedfcba = rsVoucher("fcba")
	
	contractorcode = rsVoucher("contractorcode")
	selectedcontractorcode = rsVoucher("contractorcode")
	
	description = rsVoucher("description")
	retention = rsVoucher("retention")
	startdate = rsVoucher("startdate")
	finishdate = rsVoucher("finishdate")
	tax = rsVoucher("tax")
	feecontractor = rsVoucher("feecontractor")
End if
rsVoucher.Close : Set rsVoucher = Nothing

if err.number <> 0 then Response.Write "<div class = ""error"">" & Err.Description & "</div>"
%>

 <!-------------ISI Payment Form!------------->
 
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text" aria-hidden="true"></i>Header Agreement Detail</div><% end if %>
<form name="frmEntry" method="post" action="AgreementDetail.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Agreement Code</td>
			<td align="left"><input type="text" maxlength="50" name="agreementcode" size="50" value="<%=agreementcode%>"><font color="#FF0000"> *</font></td>
			
			<td align="left">Description</td>
			<td align="left"><input type="text" name="description" size="50" value="<%=description%>"></td>
		</tr>
		
		<tr>
			<td align="left">Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="agreementdate" size="10" value="<%=agreementdate%>"><font color="#FF0000"> *</font></td>
			
			<td align="left">Retention</td>
			<td align="left"><input type="text" size="15" name="retention" value="<%=retention%>"></td>
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
			
			<td align="left">Start Date</td>
			<td align="left"><input class="datepicker" type="text" size="15" name="startdate" value="<%=startdate%>"><font color="#FF0000"> *</font></td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<select id="fcba" name="fcba" onchange="javascript:getselectedfcba();javascript:getlistcontractor();">
					<option value=""></option>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			
			<td align="left">Finish Date</td>
			<td align="left"><input class="datepicker" type="text" size="15" name="finishdate" value="<%=finishdate%>"><font color="#FF0000"> *</font></td>
		</tr>
		<tr>
			<td align="left">Contractor Code</td>
			<td align="left">
				<select id= "contractorcode" name= "contractorcode" onchange="javascript:getselectedcontractorcode();">
				</select>
				<font color="#FF0000"> *</font>
				<input type="Hidden" id="selectedcontractorcode" name="selectedcontractorcode" value="<%=selectedcontractorcode%>">
			</td>
			
			<td align="left">Tax</td>
			<td align="left"><input type="text" size="15" name="tax" value="<%=tax%>" style="vertical-align: middle;">&nbsp;&nbsp; <input type="checkbox" name="feecontractor" id="feecontractor" style="vertical-align: middle;" <%if cstr(""&feecontractor) = "1" then%> checked <%end if%> onchange="CheckCheckboxes1(this);">  <span style="vertical-align: middle;">Fee Contractor</span> </td>
		</tr>
		<tr>
			<td colspan="4" align="left" class="fbp-doc-actions">
			<input Type="Hidden" Value="True" Name="cmdView">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
			<a href="agreementheaderedit.asp?agreementcode=<%=agreementcode%>&agreementdate=<%=agreementdate%>&fccompanycode=<%=fccompanycode%>&fcba=<%=selectedfcba%>&contractorcode=<%=contractorcode%>&description=<%=description%>&retention=<%=retention%>&startdate=<%=startdate%>&finishdate=<%=finishdate%>&tax=<%=tax%>&txtJenis=Editheader" target="_blank"> Edit Header </a>
			<a href="PrintAgreementDetail.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=PeriodDate%>&fcCode=<%=fcCode%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
			</td>
		</tr>
	</table>  
</form>
<%
if cmdView = "True" and len(trim(""&agreementcode)) > 0 and len(trim(""&fccompanycode)) > 0 and len(trim(""&fcba)) > 0 and len(trim(""&startdate)) > 0 and len(trim(""&finishdate)) > 0 Then
   
   q = " SELECT agreementcode, jobcode, qty, rate, reference, remarks, lineno, fcba" & _
	   " FROM CONTRACTAGREEMENTDETAIL  " & _
	   " WHERE agreementcode = '" & agreementcode & "'" & _ 
	   "       AND fcBA = '" & fcba & "'"

   'response.Write q	
   Set rsWC = Server.CreateObject("ADODB.Recordset")
   rsWC.Open q, connstr
   %>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="AgreementDetail.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<!-------------Yang tidak bisa di edit adalah PK Table!------------->
		 <tr>
			<td align="left">Jobcode</td>
			<td align="left"><input type="text" size="15" name="jobcode" id="jobcode"  value="<%=jobcode%>"  ><font color="#FF0000"> *</font>
				<a href="javascript:launchwin('popup_activity.asp','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Activity</a>
			
			</td>
		</tr>
		<tr>
			<td align="left">Quantity</td>
			<td align="left"><input type="text" size="15" name="QTY" value="<%=QTY%>"><font color="#FF0000"> *</font></td>
		</tr>
		<tr>
			<td align="left">Rate</td>
			<td align="left"><input type="text" size="15" name="RATE" value="<%=RATE%>"><font color="#FF0000"> *</font></td>
		</tr>
		<tr>
			<td align="left">Reference</td>
			<td align="left"><input type="text" size="50" maxlength = "100" readonly name="REFERENCE" id="jobname" value="<%=REFERENCE%>"></td>
		</tr>
		<tr>
			<td align="left">Remarks</td>
			<td align="left"><textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea></td>
		</tr>
	    <TR>
			<td colspan="2" align="left">
				<input type="hidden" name="locationtypecode" id="locationtypecode">
				<input type="hidden" name="agreementcode" value="<%=agreementcode%>">
				<input type="hidden" name="agreementdate" value="<%=agreementdate%>">
				<input type="hidden" name="contractorcode" value="<%=contractorcode%>">
				
				<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
				<input type="hidden" name="fcba" value="<%=fcba%>">
				<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
				<input type="hidden" name="selectedcontractorcode" value="<%=selectedcontractorcode%>">
				<input type="hidden" name="description" value="<%=description%>">
				<input type="hidden" name="retention" value="<%=retention%>">
				<input type="hidden" name="startdate" value="<%=startdate%>">
				<input type="hidden" name="finishdate" value="<%=finishdate%>">
				<input type="hidden" name="tax" value="<%=tax%>">
				<input type="hidden" name="lineno" value="<%=lineno%>">
				<input type="hidden" name="feecontractor" value="<%=feecontractor%>">
				
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

    <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
    <div class="fbp-doc-scroll">
    <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">No</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">Agreement Code</TD>	
	<TD align="left">Line No</TD>	
	<TD align="left">Jobcode</TD>
	<TD align="left">Quantity</TD>	
	<TD align="left">Rate</TD>
	<TD align="left">Reference</TD>	
	<TD align="left">Remarks</TD>	
	</tr>
	</thead>
	<tbody>   
	</tr>   
   <%
   if Not rsWC.eof then
      nCountofField=rsWC.fields.count-1
	  no=1
	  While Not rsWC.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle">
			<a href="AgreementDetail.Asp?agreementcode=<%=rsWC("AGREEMENTCODE")%>&LINENO=<%=rsWC("LINENO")%>&jobcode=<%=rsWC("jobcode")%>&QTY=<%=rsWC("QTY")%>&RATE=<%=rsWC("RATE")%>&REFERENCE=<%=rsWC("REFERENCE")%>&REMARKS=<%=rsWC("REMARKS")%>&fcba=<%=fcba%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsWC(1)%>" border="0"></a>
			</TD>
			<td valign="middle">
			<a href="AgreementDetail.Asp?agreementcode=<%=rsWC("AGREEMENTCODE")%>&LINENO=<%=rsWC("LINENO")%>&jobcode=<%=rsWC("jobcode")%>&fcba=<%=fcba%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsWC(1)%>" border="0"></a>
			</td>
			    <td align="left" title="<%=rsWC("agreementcode").Name%>" valign="top"><%=rsWC("agreementcode")%></td>
			    <td align="left" title="<%=rsWC("lineno").Name%>" valign="top"><%=rsWC("lineno")%></td>
			    <td align="left" title="<%=rsWC("jobcode").Name%>" valign="top"><%=rsWC("jobcode")%></td>
			    <td align="left" title="<%=rsWC("qty").Name%>" valign="top"><%=rsWC("qty")%></td>
			    <td align="left" title="<%=rsWC("rate").Name%>" valign="top"><%=rsWC("rate")%></td>
			    <td align="left" title="<%=rsWC("reference").Name%>" valign="top"><%=rsWC("reference")%></td>
			    <td align="left" title="<%=rsWC("remarks").Name%>" valign="top"><%=rsWC("remarks")%></td>
		  </tr>
		  <%
		  rsWC.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsWC.Close
   set rsWC = Nothing%>
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
<script>
function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
}

function getselectedcontractorcode() {
	var scontractorcode = document.getElementById("selectedcontractorcode");
	scontractorcode.value = document.getElementById("contractorcode").value;
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
	getlistcontractor();
});

function CheckCheckboxes1(chk){
		if(chk.checked == 1){
			chk.value = 1;
		} else {
			chk.value = 0;
		}
	}
</script>
</body>
</html>
