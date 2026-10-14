<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Edit Contract Agreement</title>
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
</script>

<script language="JavaScript">
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
 Dim CONTRACTORCODE, CONTRACTORDESC, FCCODE,RETENTION,QTY,RATE,LINEno,DUEDATE
 Dim FDDATE,PeriodDate, JOB
 Dim LOCATIONTYPE
 Dim LOCATIONCode,UNIT, OUTPUT, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA,txtJenis
 Dim cmdView



agreementcode = request("agreementcode")
agreementdate = request("agreementdate")
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
FCBA=request("fcba")
txtJenis = request("txtJenis")

if Request("btnsimpan")="Simpan" then 

	q = " UPDATE CONTRACTAGREEMENT " & _
		    " SET AGREEMENTDATE = TO_DATE('" & agreementdate & "', 'mm/dd/yyyy')   " & _
			"     , CONTRACTORCODE = '" & CONTRACTORCODE & "'" & _
			"     , DESCRIPTION = '" & description & "'" & _
			"     , RETENTION = '" & RETENTION & "'" & _
			"     , FCBA = '" & FCBA & "'" & _
			"     , FINISHDATE = to_date('" & finishdate & "', 'mm/dd/yyyy') " & _
			"     , STARTDATE = to_date('" & startdate & "', 'mm/dd/yyyy') " & _
			"     , TAX = '" & tax & "' " & _
			"     , lastupdate = SYSDATE " & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi') "& _
			"     , fcedit = '" & Session("CodeName") & "' "& _
			" WHERE agreementcode = '" & agreementcode & "'" 
				
	'	Response.Write q
		connStr.execute(q)

		q = "UPDATE CONTRACTAGREEMENTDETAIL SET FCBA='" & FCBA & "' WHERE agreementcode = '" & agreementcode & "'"
		connStr.execute(q)

end if
%>
 
<form name="frmEntry" method="post" action="Agreementheaderedit.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Agreement Code</td>
			<td align="left"><input type="text" maxlength="50" readonly style="background:#eee" name="agreementcode" size="50" value="<%=agreementcode%>"><font color="#FF0000"> *</font></td>
			
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
				<div id="fcbalist"></div>
				<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
				<select name="fcba" value="<%=fcba%>" id="fcba"  class="chosen-select" style="width: 50%">
				
						<% 	Dim rsCombounit
							q = " SELECT DISTINCT fccode,fcname " & _
							    " FROM BUSINESSUNIT	"
							' response.write q
							Set rsCombounit = Server.CreateObject("ADODB.Recordset")
							rsCombounit.CursorType = 3
							rsCombounit.CursorLocation = 3
							rsCombounit.Open q, connstr
							if not rsCombounit.eof then%>
							<%
							Do While Not rsCombounit.eof
							%>
								<option value="<%=rsCombounit("fccode")%>" 
									<%if cstr(""&fcba)=cstr(""&rsCombounit("fccode")) then %> selected 
									<%End If%>><%=rsCombounit("fccode")%>
								</option> 
							<%rsCombounit.movenext
							loop
							End if
							rsCombounit.Close
							set rsCombounit = Nothing
							
						%>
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
			<input type="hidden" id="selectedcontractorcode" name="selectedcontractorcode" value="<%=selectedcontractorcode%>">
				<select id= "contractorcode" name= "contractorcode" onchange="javascript:getselectedcontractorcode();">
				<%
					qlov = "select distinct fccode,fcname " & _
					" from CONTRACTOR where fccode='" & contractorcode & "' and fcba='" & FCBA & "'"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
						
						<option value="<%=rslov("fccode")%>" <%if contractorcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
				<font color="#FF0000"> *</font>
				
			</td>
			
			<td align="left">Tax</td>
			<td align="left"><input type="text" size="15" name="tax" value="<%=tax%>"></td>
		</tr>
		<tr>
			<td colspan="2" align="left">
			
			<input Type="Submit"  name="btnsimpan" value="Simpan">&nbsp;&nbsp;
			
			</td>
		</tr>
	</table>  
</form>
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
	//getlistfcba();
	//getlistcontractor();
});
</script>
</body>
</html>
