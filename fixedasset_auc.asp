<%
	response.buffer = True
	validatesession
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
	<title>Fixed Asset Settlement AUC</title>
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

documentno = Request("documentno")
creationdate = Request("creationdate")
fccompanycode = Request("fccompanycode")
fcba = Request("fcba")
selectedfcba = Request("selectedfcba")
fixedassetcode = Request("fixedassetcode")
description = Request("description")
lineno = Request("lineno")
projectcode = Request("projectcode")


If ""&creationdate = "" then creationdate = TextDate(date)

txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
			
			'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Header" then
			connstr.BeginTrans
			
			oSql = " SELECT *" & _
				   " FROM fixedasset_auc_hdr " & _
				   " where documentno = '" & documentno & "'"

			set oRs = server.CreateObject("Adodb.recordset")
			oRs.Open oSql, connstr
			if not oRs.Eof then
				fixedassetcode = oRs("fixedassetcode")
			end if
			oRs.close : set oRs = nothing
			
			q = " DELETE FROM fixedasset_transaction" & _
					" WHERE fixedassetcode = '" & fixedassetcode & "'" & _
					"       AND transactiontype = 'ACQUISITION'"

				'Response.Write q
				connStr.execute(q)
			
			oSql = " SELECT *" & _
				   " FROM fixedasset_auc_dtl " & _
				   " where documentno = '" & documentno & "'"

			set oRs = server.CreateObject("Adodb.recordset")
			oRs.Open oSql, connstr
			if not oRs.Eof then
				while not oRs.eof
					q = " delete from fixedasset_auc_dtl" & _
						" where documentno = '" & documentno & "'" & _
						"       and lineno = '" & oRs("lineno") & "'"

					'Response.Write q
					connStr.execute(q)
					oRs.movenext
				wend
			end if
			oRs.close : set oRs = nothing
		
			q = " DELETE FROM fixedasset_auc_hdr" & _
				" WHERE documentno = '" & documentno & "'"

			'Response.Write q
			connStr.execute(q)

			if err.number = 0 then
				connStr.CommitTrans
			end if
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans
		hasdata = 0
		q = " SELECT COUNT(1) hasdata FROM fixedasset_auc_hdr WHERE documentno = '" & documentno & "'"
		Set rscekpk = Server.CreateObject("ADODB.Recordset")
		rscekpk.Open q, connstr
		
		if Not rscekpk.eof then
			hasdata = cdbl(to_nol(rscekpk("hasdata")))
		else
			hasdata = 0
		end if
		rscekpk.close : set rscekpk = nothing
		
		if hasdata = 1 then
			err.number = 1001
			err.description = "Document No : " & fccode & " , already exist!"
		else
			hasmaster = 0
			q = " SELECT COUNT(1) hasdata FROM fixedasset WHERE fccode = '" & fixedassetcode & "'"
			Set rscekpk = Server.CreateObject("ADODB.Recordset")
			rscekpk.Open q, connstr
			
			if Not rscekpk.eof then
				hasmaster = cdbl(to_nol(rscekpk("hasdata")))
			else
				hasmaster = 0
			end if
			rscekpk.close : set rscekpk = nothing
			if hasmaster = 0 then
				err.number = 1001
				err.description = "Asset : " & fixedassetcode & ", does not exist in master!"
			else
				hastransaction = 0
				q = " SELECT COUNT(1) hasdata FROM fixedasset_transaction WHERE fixedassetcode = '" & fixedassetcode & "'"
				Set rscekpk = Server.CreateObject("ADODB.Recordset")
				rscekpk.Open q, connstr
				
				if Not rscekpk.eof then
					hastransaction = cdbl(to_nol(rscekpk("hasdata")))
				else
					hastransaction = 0
				end if
				rscekpk.close : set rscekpk = nothing
				if hastransaction = 1 then
					err.number = 1001
					err.description = "Asset : " & fixedassetcode & ", has been acquired!"
				else
					q = " INSERT INTO fixedasset_auc_hdr ( documentno" & _
						"                                  , creationdate" & _
						"                                  , fixedassetcode" & _
						"                                  , projectcode" & _
						"                                  , description" & _
						"                                  , fcentry" & _
						"                                  , fcedit" & _
						"                                  , fcip" & _
						"                                  , fcba" & _
						"                                  , lastupdate" & _
						"                                  , lasttime)" & _
						" VALUES ( '" & documentno & "'" & _
						"          , TO_DATE('" & creationdate & "', 'mm/dd/yyyy')" & _
						"          , '" & fixedassetcode & "'" & _
						"          , '" & projectcode & "'" & _
						"          , '" & description & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"          , '" & fcba & "'" & _
						"          , SYSDATE" & _
						"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

					'response.write q
					connStr.execute(q)
					
					lineno = 0
					q = " SELECT NVL(MAX(lineno),0) lineno FROM fixedasset_auc_dtl WHERE documentno = '" & documentno & "'"
					Set rscekpk = Server.CreateObject("ADODB.Recordset")
					rscekpk.Open q, connstr
					
					if Not rscekpk.eof then
						lineno = cdbl(to_nol(rscekpk("lineno")))
					else
						lineno = 0
					end if
					rscekpk.close : set rscekpk = nothing
					
					total_value = 0
					For n = 1 to (cdbl(request("no")) - 1)
						total_value = total_value +cdbl(to_nol(request("amount_"&n)))
						q = " INSERT INTO fixedasset_auc_dtl ( documentno" & _
							"                                  , lineno" & _
							"                                  , accountingyear" & _
							"                                  , accountingperiod" & _
							"                                  , sourcetype" & _
							"                                  , jobcode" & _
							"                                  , amount" & _
							"                                  , fcentry" & _
							"                                  , fcedit" & _
							"                                  , fcip" & _
							"                                  , fcba" & _
							"                                  , lastupdate" & _
							"                                  , lasttime)" & _
							" VALUES ( '" & documentno & "'" & _
							"          , '" & (lineno + n) & "'" & _
							"          , '" & request("accountingyear_"&n) & "'" & _
							"          , '" & request("accountingperiod_"&n) & "'" & _
							"          , '" & request("sourcetype_"&n) & "'" & _
							"          , '" & request("jobcode_"&n) & "'" & _
							"          , '" & request("amount_"&n) & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							"          , '" & fcba & "'" & _
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE, 'hh24:mi'))"

						'response.write q
						connStr.execute(q)
					next
					
					lineno = 0
					q = " SELECT NVL(MAX(lineno),0) lineno FROM fixedasset_transaction WHERE fixedassetcode = '" & fixedassetcode & "'"
					Set rscekpk = Server.CreateObject("ADODB.Recordset")
					rscekpk.Open q, connstr
					
					if Not rscekpk.eof then
						lineno = cdbl(to_nol(rscekpk("lineno")))
					else
						lineno = 0
					end if
					rscekpk.close : set rscekpk = nothing
					
					q = " INSERT INTO fixedasset_transaction ( fixedassetcode" & _
						"                                      , lineno" & _
						"                                      , transactiontype" & _
						"                                      , amount" & _
						"                                      , ytd_depreciation" & _
						"                                      , reference" & _
						"                                      , fcentry" & _
						"                                      , fcedit" & _
						"                                      , fcip" & _
						"                                      , fcba" & _
						"                                      , lastupdate" & _
						"                                      , lasttime)" & _
						" VALUES ( '" & fixedassetcode & "'" & _
						"          , '" & (lineno +1) & "'" & _
						"          , 'ACQUISITION'" & _
						"          , '" & total_value & "'" & _
						"          , '0'" & _
						"          , '" & documentno & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"          , '" & fcba & "'" & _
						"          , SYSDATE" & _
						"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
						
					'response.write q
					connStr.execute(q)
					
				End if
			End if
		End if
		if err.number=0 then
			connStr.CommitTrans
		End if
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		cmdView	= "True"
    else	
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")		    
End if

SudahClose = "N"
set oRs = server.CreateObject("Adodb.recordset")
oSql = " SELECT fa.*, bu.fccompanycode" & _
       " FROM fixedasset_auc_hdr fa" & _
       " LEFT JOIN businessunit bu ON fa.fcba = bu.fccode" & _
	   " WHERE fa.documentno = '" & documentno & "'"

oRs.Open oSql, connstr
if Not oRs.Eof then
	creationdate = oRs("creationdate")
	fccompanycode = oRs("fccompanycode")
	selectedfcba = oRs("fcba")
	fcba = oRs("fcba")
	description = oRs("description")
	fixedassetcode = oRs("fixedassetcode")
	SudahClose = "Y"
End if
oRs.Close : Set oRs = Nothing

has_acquisition = 0
set oRs = server.CreateObject("Adodb.recordset")
oSql = " SELECT COUNT(1) hasdata" & _
       " FROM fixedasset_auc_hdr fa" & _
	   " WHERE fa.fixedassetcode = '" & fixedassetcode & "'"

oRs.Open oSql, connstr
if Not oRs.Eof then
	has_acquisition = cdbl(to_nol(oRs("hasdata")))
else
	has_acquisition = 0
End if
oRs.Close : Set oRs = Nothing

has_depreciation = 0
set oRs = server.CreateObject("Adodb.recordset")
oSql = " SELECT NVL(SUM(ytd_depreciation),0) amount" & _
       " FROM fixedasset_transaction fa" & _
	   " WHERE fa.fixedassetcode = '" & fixedassetcode & "'"

oRs.Open oSql, connstr
if Not oRs.Eof then
	has_depreciation = cdbl(to_nol(oRs("amount")))
else
	has_depreciation = 0
End if
oRs.Close : Set oRs = Nothing


If err.number <> 0 Then
	connStr.RollbackTrans
	response.Write  "<div class = ""error"">" & err.description & "</div>"
End If


Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " SELECT 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & creationdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & creationdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Costbook')"

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
<form name="frmEntry" method="post" action="fixedasset_auc.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Document No</td>
			<td align="left">
				<input type="text" maxlength="50" name="documentno" size="30" value="<%=documentno%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Creation Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="creationdate" size="10" value="<%=creationdate%>" autocomplete="off">
				<font color="#FF0000">&nbsp;*</font>
			</td>
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
			<td align="left">Fixed Asset Code</td>
			<td align="left">
				<input type="text" maxlength="50" name="fixedassetcode" size="30" value="<%=fixedassetcode%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Project Code</td>
			<td align="left">
				<input type="text" maxlength="50" name="projectcode" size="30" value="<%=projectcode%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		</tr>	
			<td align="left">Description</td>
			<td align="left"><input type="text" size="50" maxlength="100" name="description" value="<%=description%>"></td>		
		</tr>
		<tr>
			<td colspan="2"  align="left"> 
				<input Type="Hidden" Value="True" Name="cmdView">  
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">   	 
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  
				<a href="print_fixedasset_auc.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&documentno=<%=documentno%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" and has_depreciation = 0 Then%>
					<a href="fixedasset_auc.asp?documentno=<%=documentno%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=documentno%>" border="0" >
					</a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				<!-- HADIR -->
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Contract"
				dataSubmodule 	= "Work in Progress Final"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td> 
		</tr> 
	</table>    
</form>
<%
If cmdView = "True" and ""&documentno <> "" and ""&creationdate <> "" and ""&fccompanycode <> "" and ""&fcba <> "" and ""&fixedassetcode <> "" Then
	Dim rswipfinal, intRecordsPerPage, cColor
	set rswipfinal = server.CreateObject("Adodb.recordset")
	If SudahClose <> "Y" Then
		q = " SELECT sourcetype, jobcode, SUM(amount) amount" & _
		    " FROM cost_activity" & _
			" WHERE fcba = '" & selectedfcba & "'" & _
			"       AND targettype = 'PJ'" & _
			"       AND targetcode = '" & projectcode & "'" & _
			" GROUP BY sourcetype, jobcode" & _
			" ORDER BY sourcetype, jobcode"

	Else	 
		q = " SELECT *" & _
            " FROM fixedasset_auc_dtl" & _
            " WHERE documentno = '" & documentno & "'" & _
			" ORDER BY lineno"
   End if
	  
   'response.Write q
   rswipfinal.Open q, connstr%>
   <form action="fixedasset_auc.asp" method="post" name="eForm">
   <table class="data" border="1" width="100%" >
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
		   <td align="center">Year</td>
		   <td align="center">Month</td>
		   <td align="center">Source Type</td>
		   <td align="center">Jobcode</td>
		   <td align="center">Amount</td>
   	   </tr>
   <%
   if Not rswipfinal.eof then
	  no = 1
	  gt_amount = 0
	  While Not rswipfinal.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if
		  gt_amount = gt_amount + cdbl(to_nol(rswipfinal("amount")))
%>
			<tr bgcolor="<%=cColor%>">
				<%IF SudahClose = "Y" then%>
					<td align="center" title="Accounting Year" valign="top"><%=rswipfinal("accountingyear")%></td>
					<td align="right" title="Accounting Period" valign="top"><%=rswipfinal("accountingperiod")%></td>
					<td align="center" title="Sourcetype" valign="top"><%=rswipfinal("sourcetype")%></td>
					<td align="left" title="Jobcode" valign="top"><%=rswipfinal("jobcode")%></td>
					<td align="right" title="Amount" valign="top"><%=FormatNumber(rswipfinal("amount"),2)%></td>
					
				<%else%>
					<td align="center" title="Accounting Year" valign="top"><%=rswipfinal("accountingyear")%></td>
					<td align="right" title="Accounting Period" valign="top"><%=rswipfinal("accountingperiod")%></td>
					<td align="center" title="Sourcetype" valign="top"><%=rswipfinal("sourcetype")%></td>
					<td align="left" title="Jobcode" valign="top"><%=rswipfinal("jobcode")%></td>
					<td align="right" title="Amount" valign="top"><%=FormatNumber(rswipfinal("amount"),2)%></td>
					
					<input type="hidden" name="accountingyear_<%=no%>" value="<%=rswipfinal("accountingyear")%>">
					<input type="hidden" name="accountingperiod_<%=no%>" value="<%=rswipfinal("accountingperiod")%>">
					<input type="hidden" name="sourcetype_<%=no%>" value="<%=rswipfinal("sourcetype")%>">
					<input type="hidden" name="jobcode_<%=no%>" value="<%=rswipfinal("jobcode")%>">
					<input type="hidden" name="amount_<%=no%>" value="<%=rswipfinal("amount")%>">
				<%End if%>
		  </tr>
		  <%
		  no = no + 1
		  rswipfinal.MoveNext
	  Wend
			response.write "<tr style='background-color:#F4F4A0;'>"
			response.write "	<td align='right' colspan='4'><b>Total</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(gt_amount,2)&"</b></td>"
			response.write "</tr>"
   else
      response.Write("Data is empty")	  
   end if 
   rswipfinal.Close
   set rswipfinal = Nothing%>
		<tr>
			<td colspan="11" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="documentno" value="<%=documentno%>">
			<input type="hidden" name="creationdate" value="<%=creationdate%>">
			<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
			<input type="hidden" name="fcba" value="<%=fcba%>">
			<input type="hidden" name="projectcode" value="<%=projectcode%>">
			<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
			<input type="hidden" name="fixedassetcode" value="<%=fixedassetcode%>">
			<input type="hidden" name="description" value="<%=description%>">
			
			<input type="hidden" name="No" value="<%=No%>">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<%If val_IsTrue = "Y" then%>
				<%If SudahClose <> "Y" then%>
					<%if has_acquisition = 0 then%>
					<input type="hidden" name="txtJenis" value="New">
					<input Type="Submit" name="Add" value="Submit">
					<input Type="Reset" name="Reset" value="Reset">
					<input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)">
					<%Else%>
					<input type="hidden" name="txtJenis" value="New">
					<input Type="Submit" name="Add" value="Submit" disabled>
					<input Type="Reset" name="Reset" value="Reset" disabled>
					<input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)" disabled>
					<%End if%>
				<%End if%>
			<%End if%>
			</td>
		</tr>
	
	</TABLE>
</div>
</form>
<%End if%>
<!---------------akhir!--------------->
</div>
</div>
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
});
</script>
</body>
</html>
