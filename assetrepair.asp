<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Asset Repair</title>

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
</script>

<script>
	function hanyaAngka(evt) {
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
var getMessage = function () {
	// build up the post string when passing variables to the server side page
	
	var fccode = document.getElementById("fccode").value;
	var PostStr = 'fccode='+fccode;

	// use the generic function to make the request
	doAJAXCall('ajax/GetFAGroup.asp', 'POST', PostStr, showMessageResponse);
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    
    // update the Div to show the result from the server
	document.getElementById("fcname").value = response;
};

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
 <!-------------awal!------------->
 <%

referencenumber =request("referencenumber")
referencedate =request("referencedate")
sourcetype =request("sourcetype")
sourcecode =request("sourcecode")
sourcecodedesc =request("sourcecodedesc")
itemcode =request("itemcode")
assetname =request("assetname")
startdepreciation =request("startdepreciation")
acquisitionvalue =request("acquisitionvalue")
accum_depre_val =request("accum_depre_val")
netbook =request("netbook")
curr_depre_value =request("curr_depre_value")
estimatelifeyear=request("estimatelifeyear")
effectivedate =request("effectivedate")
lifetimeremain =request("lifetimeremain")
repaircost =request("repaircost")
additionlifetime =request("additionlifetime")
lifetimeremain =request("lifetimeremain")
newlifetime =request("newlifetime")
new_deprate_rate =request("new_deprate_rate")
description =request("description")
currencycode =request("currencycode")
remarks =request("remarks")
currencyrate =request("currencyrate")
amount =request("amount")

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
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), creationdate)
		If Cek_IsDailyClosing = "FALSE" then
			q = " delete from assetrepairhdr " & _
				" where fcBA = '" & Session("fcBA") & "'" & _
				" and referencenumber = '" & referencenumber & "'"
				' Response.Write q
				connStr.execute(q)
			
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & TRANSDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			
			q = " delete from assetrepairdtl " & _
				" where fcBA = '" & Session("fcBA") & "'" & _
				" and referencenumber = '"& referencenumber & "'" 
				' Response.Write q
				connStr.execute(q)

			
			cmdView	= "True"

	elseif Request("txtJenis")="New" then		
			Dim rsSupplyingDeclarationHdr
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select salescontractno from salesagreementhdr where salescontractno = '" & salescontractno & "' and fcba ='" & Session("fcba") & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then
				q = " INSERT into assetrepairhdr" & _
					" (referencenumber, referencedate, itemcode, assetname, startdepreciation, acquisitionvalue, accum_depre_val, netbook, curr_depre_value, estimatelifeyear, effectivedate, lifetimeremain, repaircost, additionlifetime, newlifetime, new_deprate_rate, description, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" '" & referencenumber & "'" & _
					" , to_date('" & referencedate & "','mm/dd/yyyy')" & _
					" , '" & itemcode & "'" & _
					" , '" & assetname & "'" & _
					" , to_date('" & startdepreciation & "','mm/dd/yyyy')" & _
					" , '" & acquisitionvalue & "'" & _
					" , '" & accum_depre_val & "'" & _
					" , '" & netbook & "'" & _
					" , '" & curr_depre_value & "'" & _
					" , '" & estimatelifeyear & "'" & _
					" , to_date('" & effectivedate & "','mm/dd/yyyy')" & _
					" , '" & lifetimeremain & "'" & _
					" , '" & repaircost & "'" & _
					" , '" & additionlifetime & "'" & _
					" , '" & newlifetime & "'" & _
					" , '" & new_deprate_rate & "'" & _
					" , '" & description & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & Session("fcba") & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"								
				' Response.Write q
				connStr.execute(q)
				
			End If
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing
					
			If Len(Trim(sourcetype)) = 0 Then
				Err.Number = 1001
				Err.Description = "Source Type Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(sourcecode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Source Code Harus diisi, tidak boleh kosong"
			Else
			
					q = " insert into assetrepairdtl" & _
						" (sourcetype, sourcecode, sourcecodedesc, referencenumber, referencedate, remarks, currencycode, acquisitionvalue, currencyrate, amount, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
						" '" & sourcetype & "'" & _
						" , '" & sourcecode & "'" & _
						" , '" & sourcecodedesc & "'" & _
						" , '" & referencenumber & "'" & _
						" , to_date('" & referencedate & "','mm/dd/yyyy')" & _ 
						" , '" & remarks & "'" & _
						" , '" & currencycode & "'" & _
						" , '" & acquisitionvalue & "'" & _
						" , '" & currencyrate & "'" & _
						" , '" & amount & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						" , '" & Session("fcba") & "'" & _
						" , sysdate" & _
						" , to_char(sysdate, 'hh24:mi')" & _
						" )"
						' Response.Write q
						connStr.execute(q)
				End If
			End If			

cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			If Len(Trim(sourcetype)) = 0 Then
				Err.Number = 1001
				Err.Description = "Source Type Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(sourcecode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Source Code Harus diisi, tidak boleh kosong"
			Else
					
					q =	" Update assetrepairdtl" & _
						" set  sourcetype = '" & sourcetype & "'" & _
						"     , sourcecode = '" & sourcecode & "'" & _
						"     , sourcecodedesc = '" & sourcecodedesc & "'" & _
						"     , remarks = '" & remarks & "'" & _
						"     , currencycode = '" & currencycode & "'" & _
						"     , currencyrate = '" & currencyrate & "'" & _
						"     , amount = '" & amount & "'" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where referencenumber = '" & referencenumber & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
		 			' response.write q
					connStr.execute(q)
					End If
				End If			
			End if
		End If			
cmdView	= "True"	
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

'untuk menampilkan datepicker default
If Len(Trim(referencedate)) = 0 Then	referencedate = textDate(date)
If Len(Trim(effectivedate)) = 0 Then	effectivedate = textDate(date)
If Len(Trim(startdepreciation)) = 0 Then	startdepreciation = textDate(date)


If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & TRANSDATE & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & TRANSDATE & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
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
 
	<form name="frmEntry" method="post" action="assetrepair.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Reference Number 
				<td align="left">
				<input type="text" maxlength="50" name="referencenumber" size="30" value="<%=referencenumber%>" style="background:#FFFFCC;" autocomplete="Off" onkeyup="hanyaAngka(this);">
				<font color="#FF0000">*</font>
				</td>
			</td>
			<td align="left" width="20%">Repair Cost
				<td align="left">
				<input type="text" maxlength="15" name="repaircost" size="30" value="<%=repaircost%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >

				</td>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Reference Date
				<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="referencedate" size="15" value="<%=referencedate%>" autocomplete="Off">
				</td>
			</td>
			<td align="left" width="20%">Curr. Depre. Value
				<td align="left">
				<input type="text" maxlength="15" name="curr_depre_value" size="30" value="<%=curr_depre_value%>" autocomplete="Off" >
				</td>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Item Code
				<td align="left">
				<input type="text" maxlength="15" name="itemcode" size="30" value="<%=itemcode%>" style="background:#FFFFCC;" autocomplete="Off" >
				<font color="#FF0000">*</font>
				</td>
			</td>
			<td align="left" width="20%">Accum. Depre. Value
				<td align="left">
				<input type="text" maxlength="15" name="accum_depre_val" size="30" value="<%=accum_depre_val%>" autocomplete="Off" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Asset Name
				<td align="left">
				<input type="text" maxlength="15" name="assetname" size="30" value="<%=assetname%>" autocomplete="Off">

				</td>
			</td>

			<td align="left" width="20%">Net Book
				<td align="left">
				<input type="text" maxlength="15" name="netbook" size="30" value="<%=netbook%>" autocomplete="Off" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Acquisition Value
				<td align="left">
				<input type="text" maxlength="15" name="acquisitionvalue" size="30" value="<%=acquisitionvalue%>" autocomplete="Off" >

				</td>
			</td>


			 <td align="left" width="20%">Estimate Life Year
				<td align="left">
				<input type="text" maxlength="15" name="estimatelifeyear" size="30" value="<%=estimatelifeyear%>" autocomplete="Off" >

				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Effective Date
				<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="effectivedate" size="15" value="<%=effectivedate%>" autocomplete="Off">
				</td>
			</td>

			 	<td align="left" width="20%">Life Time Remain
				<td align="left">
				<input type="text" maxlength="15" name="lifetimeremain" size="30" value="<%=lifetimeremain%>" autocomplete="Off" >

				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Start Depreciation
				<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="startdepreciation" size="15" value="<%=startdepreciation%>" autocomplete="Off">
				</td>
			</td>
			<td align="left" width="20%">Addition Life Time
				<td align="left">
				<input type="text" maxlength="15" name="additionlifetime" size="30" value="<%=additionlifetime%>" autocomplete="Off" >

				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">New Life Time
				<td align="left">
				<input type="text" maxlength="15" name="newlifetime" size="30" value="<%=newlifetime%>" autocomplete="Off" >

				</td>
			</td>
			<td align="left" width="20%">New Deprate Rate
				<td align="left">
				<input type="text" maxlength="15" name="new_deprate_rate" size="30" value="<%=new_deprate_rate%>" autocomplete="Off" >

				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Description
				<td align="left">
				<input type="text" maxlength="15" name="description" size="30" value="<%=description%>" autocomplete="Off" >

				</td>
			</td>
			
		</tr>

		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				
					<a href="assetrepair.asp?itemcode=<%=itemcode%>&referencenumber=<%=referencenumber%>&txtJenis=HeaderDelete">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0" onClick="return confirm('Are you sure to delete');">
					</a>
				
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
if cmdView = "True" and len(trim(referencenumber)) > 0 and len(trim(itemcode)) > 0  Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= " SELECT DISTINCT ARH.REFERENCENUMBER, ARH.REFERENCEDATE, ARH.ITEMCODE, ARH.ASSETNAME, ARH.STARTDEPRECIATION, ARH.ACQUISITIONVALUE, " & _
		  "	ARH.ACCUM_DEPRE_VAL, ARH.NETBOOK, ARH.CURR_DEPRE_VALUE, ARH.ESTIMATELIFEYEAR, ARH.EFFECTIVEDATE, ARH.LIFETIMEREMAIN, ARH.REPAIRCOST," & _
		  "	ARH.ADDITIONLIFETIME, ARH.NEWLIFETIME, ARH.NEW_DEPRATE_RATE, ARH.DESCRIPTION, ARD.SOURCECODE, ARD.SOURCETYPE, ARD.SOURCECODEDESC, ARD.REMARKS, " & _
	 	  "	ARD.CURRENCYCODE, ARD.CURRENCYRATE, ARD.AMOUNT" & _
		  "	FROM ASSETREPAIRHDR ARH" & _
		  "	INNER JOIN ASSETREPAIRDTL ARD ON (ARH.REFERENCENUMBER = ARD.REFERENCENUMBER AND ARH.FCBA = ARD.FCBA)" & _
		  "	WHERE ARH.FCBA = '"& Session("fcba") &"'" & _
		  "	AND ARH.REFERENCENUMBER = '"& referencenumber &"'" 

	' response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<form action="assetrepair.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
		
		<tr>
			<td align="left" width="20%">Source Type
				<td align="left">
				<input type="text" maxlength="15" name="sourcetype" size="30" value="<%=sourcetype%>" autocomplete="Off" >
				</td>
			</td>
			<td align="left" width="20%">Currency
				<td align="left">
				<input type="text" maxlength="15" name="currencycode" size="30" value="<%=currencycode%>" autocomplete="Off">
				</td>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Source Code
				<td align="left">
				<input type="text" maxlength="15" name="sourcecode" id="fcname" size="30" value="<%=sourcecode%>">
				</td>
			</td>
			<td align="left" width="20%">Currency Rate
				<td align="left">
				<input type="text" maxlength="15" name="currencyrate" size="30" value="<%=currencyrate%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Source Code Desc.
				<td align="left">
				<textarea name="sourcecodedesc"><%=sourcecodedesc%></textarea>
			
				</td>
			</td>
			<td align="left" width="20%">Amount
				<td align="left">
				<input type="text" maxlength="15" name="amount" size="30" value="<%=amount%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
				</td>
			</td>	
		</tr>
		<tr>
			<td align="left" width="20%">Remarks
				<td align="left">
				<textarea name="remarks"><%=remarks%></textarea>
			
				</td>
			</td>
			
		</tr>
	
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="referencenumber" value="<%=referencenumber%>">
					<input type="hidden" name="referencedate" value="<%=referencedate%>">
					<input type="hidden" name="itemcode" value="<%=itemcode%>">
					<input type="hidden" name="assetname" value="<%=assetname%>">
					<input type="hidden" name="acquisitionvalue" value="<%=acquisitionvalue%>">
					<input type="hidden" name="effectivedate" value="<%=effectivedate%>">
					<input type="hidden" name="startdepreciation" value="<%=startdepreciation%>">
					<input type="hidden" name="newlifetime" value="<%=newlifetime%>">
					<input type="hidden" name="description" value="<%=description%>">
					<input type="hidden" name="repaircost" value="<%=repaircost%>">
					<input type="hidden" name="curr_depre_value" value="<%=curr_depre_value%>">
					<input type="hidden" name="accum_depre_val" value="<%=accum_depre_val%>">
					<input type="hidden" name="netbook" value="<%=netbook%>">
					<input type="hidden" name="estimatelifeyear" value="<%=estimatelifeyear%>">
					<input type="hidden" name="lifetimeremain" value="<%=lifetimeremain%>">
					<input type="hidden" name="additionlifetime" value="<%=additionlifetime%>">
					<input type="hidden" name="new_deprate_rate" value="<%=new_deprate_rate%>">
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

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="2" width="30">&nbsp;</td>
		<td align="left">ITEM CODE</td>
		<td align="left">ASSET NAME</td>
		<td align="left">ACQUISITION VALUE</td>
		<td align="left">REPAIR COST </td>
		<td align="left">EFFECTIVE DATE</td>
		<td align="left">START DEPRECIATION</td>
		<td align="left">ACCUM. DEPRE. VALUE</td>
		<td align="left">LIFE TIME REMAIN</td>
		<td align="left">ADDITION LIFE TIME</td>
		<td align="left">ESTIMATE LIFE YEAR</td>
	</tr>
   
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
			<td align="left" valign="top"><%=NO%></td>
			
			
			<td valign="middle" width="50" colspan="2">
				<a href="assetrepair.asp?itemcode=<%=rsVH("itemcode")%>&referencenumber=<%=rsVH("referencenumber")%>&referencedate=<%=rsVH("referencedate")%>&assetname=<%=rsVH("assetname")%>&startdepreciation=<%=rsVH("startdepreciation")%>&estimatelifeyear=<%=rsVH("estimatelifeyear")%>&effectivedate=<%=rsVH("effectivedate")%>&lifetimeremain=<%=rsVH("lifetimeremain")%>&repaircost=<%=rsVH("repaircost")%>&additionlifetime=<%=rsVH("additionlifetime")%>&newlifetime=<%=rsVH("newlifetime")%>&new_deprate_rate=<%=rsVH("new_deprate_rate")%>&description=<%=rsVH("description")%>&sourcecode=<%=rsVH("sourcecode")%>&sourcetype=<%=rsVH("sourcetype")%>&sourcecodedesc=<%=rsVH("sourcecodedesc")%>&remarks=<%=rsVH("remarks")%>&currencycode=<%=rsVH("currencycode")%>&currencyrate=<%=rsVH("currencyrate")%>&amount=<%=rsVH("amount")%>&curr_depre_value=<%=rsVH("curr_depre_value")%>&accum_depre_val=<%=rsVH("accum_depre_val")%>&netbook=<%=rsVH("netbook")%>&acquisitionvalue=<%=rsVH("acquisitionvalue")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
		
				<a href="assetrepair.asp?itemcode=<%=rsVH("itemcode")%>&referencenumber=<%=rsVH("referencenumber")%>&referencedate=<%=rsVH("referencedate")%>&assetname=<%=rsVH("assetname")%>&startdepreciation=<%=rsVH("startdepreciation")%>&estimatelifeyear=<%=rsVH("estimatelifeyear")%>&effectivedate=<%=rsVH("effectivedate")%>&lifetimeremain=<%=rsVH("lifetimeremain")%>&repaircost=<%=rsVH("repaircost")%>&additionlifetime=<%=rsVH("additionlifetime")%>&newlifetime=<%=rsVH("newlifetime")%>&new_deprate_rate=<%=rsVH("new_deprate_rate")%>&description=<%=rsVH("description")%>&sourcecode=<%=rsVH("sourcecode")%>&sourcetype=<%=rsVH("sourcetype")%>&sourcecodedesc=<%=rsVH("sourcecodedesc")%>&remarks=<%=rsVH("remarks")%>&currencycode=<%=rsVH("currencycode")%>&currencyrate=<%=rsVH("currencyrate")%>&amount=<%=rsVH("amount")%>&curr_depre_value=<%=rsVH("curr_depre_value")%>&accum_depre_val=<%=rsVH("accum_depre_val")%>&netbook=<%=rsVH("netbook")%>&acquisitionvalue=<%=rsVH("acquisitionvalue")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
			    <td align="left" title="<%=rsVH("itemcode").Name%>" valign="top"><%=rsVH("itemcode")%></td>
			    <td align="left" title="<%=rsVH("assetname").Name%>" valign="top"><%=rsVH("assetname")%></td>
			    <td align="left" title="<%=rsVH("acquisitionvalue").Name%>" valign="top"><%=rsVH("acquisitionvalue")%></td>
			    <td align="left" title="<%=rsVH("repaircost").Name%>" valign="top"><%=rsVH("repaircost")%></td>
			    <td align="left" title="<%=rsVH("effectivedate").Name%>" valign="top"><%=rsVH("effectivedate")%></td>
			    <td align="left" title="<%=rsVH("startdepreciation").Name%>" valign="top"><%=rsVH("startdepreciation")%></td>
			    <td align="left" title="<%=rsVH("accum_depre_val").Name%>" valign="top"><%=rsVH("accum_depre_val")%></td>
			    <td align="left" title="<%=rsVH("lifetimeremain").Name%>" valign="top"><%=rsVH("lifetimeremain")%></td>
			    <td align="left" title="<%=rsVH("additionlifetime").Name%>" valign="top"><%=rsVH("additionlifetime")%></td>
			    <td align="left" title="<%=rsVH("estimatelifeyear").Name%>" valign="top"><%=rsVH("estimatelifeyear")%></td>
			  
			</tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

		<script src="/js/jquery.inputmask.bundle.min.js"></script>
		<script src="/js/inputmask.binding.js"></script>


</body>
</html>
