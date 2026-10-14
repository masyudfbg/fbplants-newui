<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Asset Additional</title>

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

po_code =request("po_code")
itemcode =request("itemcode")
itemdescription =request("itemdescription")
suppliername =request("suppliername")
acquisitionvalue =request("acquisitionvalue")
currencycode =request("currencycode")
currencyrate =request("currencyrate")
qty_hdr =request("qty_hdr")
qty_dtl =request("qty_dtl")
receivenotecode =request("receivenotecode")
receivenotedate =request("receivenotedate")
assetname =request("assetname")
fa_groupid =request("fa_groupid")
fa_groupname =request("fa_groupname")
location_id =request("location_id")
installdate =request("installdate")
additionvalue =request("additionvalue")
residualvalue =request("residualvalue")
vouchernumber =request("vouchernumber")
specification =request("specification")
amount =request("amount")
amountdepreciation =request("amountdepreciation")
tagnumber =request("tagnumber")
serialnumber =request("serialnumber")
startdepreciation =request("startdepreciation")
estimatelifeyear =request("estimatelifeyear")
assettype =request("assettype")
jobcode =request("jobcode")
costcenter =request("costcenter")
brand =request("brand")
madein =request("madein")
taxgroupcode =request("taxgroupcode")
taxgroupname =request("taxgroupname")




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
			q = " delete from assetadditionalhdr " & _
				" where fcBA = '" & Session("fcBA") & "'" & _
				" and itemcode = '" & itemcode & "'"
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

			
			q = " delete from assetadditionaldtl " & _
				" where fcBA = '" & Session("fcBA") & "'" & _
				" and itemcode = '"& itemcode & "'" 
				' Response.Write q
				connStr.execute(q)

			
			cmdView	= "True"

	elseif Request("txtJenis")="New" then		
			Dim rsSupplyingDeclarationHdr
			Set rsSupplyingDeclarationHdr = Server.CreateObject("ADODB.Recordset")
			rsSupplyingDeclarationHdr.Open " select salescontractno from salesagreementhdr where salescontractno = '" & salescontractno & "' and fcba ='" & Session("fcba") & "' ", connstr
			If rsSupplyingDeclarationHdr.EOF Then
			If cdbl(to_nol(qty_dtl)) > cdbl(to_nol(qty_hdr)) Then
				Err.Number = 1001
				Err.Description = "Anda melebihi jumlah quantity yang dibeli"
			Else
				q = " INSERT into assetadditionalhdr" & _
					" (po_code, itemcode, itemdescription, suppliername, acquisitionvalue, currencycode, currencyrate, qty_hdr, receivenotecode, receivenotedate, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" '" & po_code & "'" & _
					" , '" & itemcode & "'" & _
					" , '" & itemdescription & "'" & _
					" , '" & suppliername & "'" & _
					" , '" & acquisitionvalue & "'" & _
					" , '" & currencycode & "'" & _
					" , '" & currencyrate & "'" & _
					" , '" & qty_hdr & "'" & _
					" , '" & receivenotecode & "'" & _
					" , to_date('" & receivenotedate & "','mm/dd/yyyy')" & _
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
			rsSupplyingDeclarationHdr.Close
			set rsSupplyingDeclarationHdr = Nothing
					
			If Len(Trim(fa_groupid)) = 0 Then
				Err.Number = 1001
				Err.Description = "FA Group ID Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(qty_dtl)) = 0 Then
				Err.Number = 1001
				Err.Description = "Quantity Fix Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(qty_dtl)) > cdbl(to_nol(qty_hdr)) Then
				Err.Number = 1001
				Err.Description = "Anda melebihi jumlah quantity yang dibeli"
			Else
			If cdbl(to_nol(vouchernumber)) = 0 Then
				Err.Number = 1001
				Err.Description = "Voucher Number Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(serialnumber)) = 0 Then
				Err.Number = 1001
				Err.Description = "Serial Number Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(amount)) = 0 Then
				Err.Number = 1001
				Err.Description = "Amount Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(amountdepreciation)) = 0 Then
				Err.Number = 1001
				Err.Description = "Amount Depreciation Harus diisi, tidak boleh kosong"
			Else
					
					q = " insert into assetadditionaldtl" & _
						" (assetname, fa_groupid, fa_groupname, itemcode, location_id, qty_dtl, installdate, acquisitionvalue, additionvalue, residualvalue, po_code, vouchernumber, specification, amount, amountdepreciation, tagnumber, serialnumber, startdepreciation, estimatelifeyear, jobcode, costcenter, brand, madein, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
						" '" & itemdescription & "'" & _
						" , '" & fa_groupid & "'" & _
						" , '" & fa_groupname & "'" & _
						" , '" & itemcode & "'" & _
						" , '" & location_id & "'" & _
						" , '" & qty_dtl & "'" & _
						" , to_date('" & installdate & "','mm/dd/yyyy')" & _ 
						" , '" & acquisitionvalue & "'" & _
						" , '" & additionvalue & "'" & _
						" , '" & residualvalue & "'" & _
						" , '" & po_code & "'" & _
						" , '" & vouchernumber & "'" & _
						" , '" & specification & "'" & _
						" , '" & amount & "'" & _
						" , '" & amountdepreciation & "'" & _
						" , '" & tagnumber & "'" & _
						" , '" & serialnumber & "'" & _
						" , to_date('" & startdepreciation & "','mm/dd/yyyy')" & _ 
						" , '" & estimatelifeyear & "'" & _
						" , '" & jobcode & "'" & _
						" , '" & costcenter & "'" & _
						" , '" & brand & "'" & _
						" , '" & madein & "'" & _
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
			End if
		End If			
	End If
End If
End If
cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			If Len(Trim(fa_groupid)) = 0 Then
				Err.Number = 1001
				Err.Description = "FA Group ID Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(qty_dtl)) = 0 Then
				Err.Number = 1001
				Err.Description = "Quantity Fix Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(qty_dtl)) > cdbl(to_nol(qty_hdr)) Then
				Err.Number = 1001
				Err.Description = "Anda melebihi jumlah quantity yang dibeli"
			Else
			If cdbl(to_nol(vouchernumber)) = 0 Then
				Err.Number = 1001
				Err.Description = "Voucher Number Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(serialnumber)) = 0 Then
				Err.Number = 1001
				Err.Description = "Serial Number Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(amount)) = 0 Then
				Err.Number = 1001
				Err.Description = "Amount Harus diisi, tidak boleh kosong"
			Else
			If cdbl(to_nol(amountdepreciation)) = 0 Then
				Err.Number = 1001
				Err.Description = "Amount Depreciation Harus diisi, tidak boleh kosong"
			Else
					
					q =	" Update assetadditionaldtl" & _
						" set  fa_groupid = '" & fa_groupid & "'" & _
						"     , fa_groupname = '" & fa_groupname & "'" & _
						"     , location_id = '" & location_id & "'" & _
						"     , qty_dtl = '" & qty_dtl & "'" & _
						"     , amount = '" & amount & "'" & _
						"     , brand = '" & brand & "'" & _
						"     , costcenter = '" & costcenter & "'" & _
						"     , additionvalue = '" & additionvalue & "'" & _
						"     , vouchernumber = '" & vouchernumber & "'" & _
						"     , serialnumber = '" & serialnumber & "'" & _
						"     , specification = '" & specification & "'" & _
						"     , residualvalue = '" & residualvalue & "'" & _
						"     , amountdepreciation = '" & amountdepreciation & "'" & _
						"     , madein = '" & madein & "'" & _
						"     , estimatelifeyear = '" & estimatelifeyear & "'" & _
						"     , tagnumber = '" & tagnumber & "'" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where itemcode = '" & itemcode & "'" & _
						"       and fcBA = '" & Session("fcBA") & "'"
		 			' response.write q
					connStr.execute(q)
					End If
				End If			
			End if
		End If			
	End If
End If
End If
End If
End If
cmdView	= "True"	
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

'untuk menampilkan datepicker default
If Len(Trim(receivenotedate)) = 0 Then	receivenotedate = textDate(date)
If Len(Trim(installdate)) = 0 Then	installdate = textDate(date)
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
 
	<form name="frmEntry" method="post" action="assetadditional.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">PO Code 
				<td align="left">
				<input type="text" maxlength="50" name="po_code" style="background:#FFFFCC;" size="30" value="<%=po_code%>" autocomplete="Off">
				<font color="#FF0000">*</font>
				</td>
			</td>
			<td align="left" width="20%">Receive Note Code
				<td align="left">
				<input type="text" maxlength="15" name="receivenotecode" style="background:#FFFFCC;" size="30" value="<%=receivenotecode%>" autocomplete="Off">
				<font color="#FF0000">*</font>
				</td>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Item Code
				<td align="left">
				<input type="text" maxlength="15" name="itemcode"  style="background:#FFFFCC;" size="30" value="<%=itemcode%>" autocomplete="Off">
				<font color="#FF0000">*</font>
				</td>
			</td>
			<td align="left" width="20%">Receive Note Date
				<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="receivenotedate" size="15" value="<%=receivenotedate%>" autocomplete="Off">
				</td>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Item Description
				<td align="left">
				<textarea name="itemdescription"><%=itemdescription%></textarea>
				</td>
			</td>
			<td align="left" width="20%">Supplier Name
				<td align="left">
				<input type="text" maxlength="15" name="suppliername" size="30" value="<%=suppliername%>" autocomplete="Off" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Acquistion Value
				<td align="left">
				<input type="text" maxlength="15" name="acquisitionvalue" size="30" value="<%=acquisitionvalue%>" autocomplete="Off">

				</td>
			</td>

			<td align="left" width="20%">Quantity 
				<td align="left">
				<input type="text" maxlength="15" name="qty_hdr" size="30" value="<%=qty_hdr%>" autocomplete="Off" onkeyup="hanyaAngka(this);">
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Currency 
				<td align="left">
				<input type="text" maxlength="15" name="currencycode" size="30" value="<%=currencycode%>" autocomplete="Off">

				</td>
			</td>

			 	<td align="left" width="20%">Currency Rate
				<td align="left">
				<input type="text" maxlength="15" name="currencyrate" size="30" value="<%=currencyrate%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">

				</td>
			</td>
		</tr>

		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				
					<a href="assetadditional.asp?itemcode=<%=itemcode%>&po_code=<%=po_code%>&receivenotecode=<%=receivenotecode%>&txtJenis=HeaderDelete">
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
if cmdView = "True" and len(trim(po_code)) > 0 and len(trim(receivenotecode)) > 0 and len(trim(itemcode)) > 0  Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= " SELECT DISTINCT ASH.ITEMCODE, ASH.PO_CODE, ASH.ACQUISITIONVALUE, ASH.CURRENCYCODE, ASH.ITEMDESCRIPTION, ASH.QTY_HDR, ASH.RECEIVENOTECODE, ASH.RECEIVENOTEDATE," & _
		  " ASH.SUPPLIERNAME, ASD.ADDITIONVALUE, ASD.AMOUNT, ASD.AMOUNTDEPRECIATION, ASD.ASSETNAME, ASD.BRAND, ASD.COSTCENTER, ASD.ESTIMATELIFEYEAR, ASD.FA_GROUPID, " & _
		  "	ASD.FA_GROUPNAME, ASD.JOBCODE, ASD.LOCATION_ID, ASD.MADEIN, ASD.QTY_DTL, ASD.RESIDUALVALUE, ASD.SERIALNUMBER, ASD.SPECIFICATION, ASD.STARTDEPRECIATION, " & _
		  "	ASD.TAGNUMBER, ASD.VOUCHERNUMBER, ASD.INSTALLDATE, ASH.CURRENCYRATE " & _
		  "	FROM ASSETADDITIONALHDR ASH " & _
		  "	INNER JOIN ASSETADDITIONALDTL ASD ON (ASH.ITEMCODE = ASD.ITEMCODE AND ASH.FCBA =  ASD.FCBA) " & _
		  "	WHERE ASH.FCBA = '"& Session("FCBA") &"'" & _
		  "	AND ASH.ITEMCODE = '"& itemcode &"'"

	' response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<form action="assetadditional.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
	
		<tr>
			<td align="left" width="20%">FA Group Id 
				<td align="left">
				<select name="fa_groupid" id="fccode" title="Delivery Type CODE Character a s/d z or 1 s/d 9" onchange="javascript:getMessage();">
					<option value="<%=fa_groupid%>"></option> 
						<% 	Dim rsComboP
							q = "SELECT * FROM assetgroup WHERE fcba = '" & Session("fcba") & "'"
							Set rsComboP = Server.CreateObject("ADODB.Recordset")
							rsComboP.CursorType = 3
							rsComboP.CursorLocation = 3
							rsComboP.Open q, connstr
							if not rsComboP.eof then%>
							<%
							Do While Not rsComboP.eof
							%>
								<option value="<%=rsComboP("fccode")%>" 
									<%if cstr(""&fa_groupid)=cstr(""&rsComboP("fccode")) then %> selected 
									<%End If%>><%=rsComboP("fccode")%>
								</option> 
							<%rsComboP.movenext
							loop
							End if
							rsComboP.Close
							set rsComboP = Nothing
							q=""
						%>
				</select> 
				
				</td>
			</td>
			<td align="left" width="20%">Additional Value
				<td align="left">
				<input type="text" maxlength="15" name="additionvalue" size="30" value="<%=additionvalue%>" autocomplete="Off" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">FA Group Name
				<td align="left">
				<input type="text" maxlength="15" name="fa_groupname" id="fcname" style="background:#eee;" size="30" value="<%=fa_groupname%>" readonly="true">
				</td>
			</td>

			<td align="left" width="20%">Voucher Number
				<td align="left">
				<input type="text" maxlength="15" name="vouchernumber" size="30" value="<%=vouchernumber%>" autocomplete="Off" onkeyup="hanyaAngka(this);" >
				
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Location
				<td align="left">
				<input type="text" maxlength="15" name="location_id" size="30" value="<%=location_id%>" autocomplete="Off">
			
				</td>
			</td>

			<td align="left" width="20%">Serial Number 
				<td align="left">
				<input type="text" maxlength="15" name="serialnumber" size="30" value="<%=serialnumber%>" autocomplete="Off" onkeyup="hanyaAngka(this);">
				</td>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Quantity Fix
				<td align="left">
				<input type="text" maxlength="15" name="qty_dtl" size="30" value="<%=qty_dtl%>" autocomplete="Off" onkeyup="hanyaAngka(this);" >
				</td>
			</td>
			<td align="left" width="20%">Specification
				<td align="left">
				<textarea name="specification"><%=specification%></textarea>
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Install Date
				<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="installdate" size="15" value="<%=installdate%>" autocomplete="Off" >
				</td>
			</td>
			<td align="left" width="20%">Residual Value 
				<td align="left">
				<input type="text" maxlength="15" name="residualvalue" size="30" value="<%=residualvalue%>" autocomplete="Off" >
				</td>
			</td>
		</tr>
		
		<tr>
			<td align="left" width="20%">Start Depreciation
				<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="startdepreciation" size="15" value="<%=startdepreciation%>" autocomplete="Off" >
				</td>
			</td>
			<td align="left" width="20%">Brand
				<td align="left">
				<input type="text" maxlength="15" name="brand" size="30" value="<%=brand%>" autocomplete="Off" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Amount Depreciation
				<td align="left">
				<input type="text" maxlength="15" name="amountdepreciation" size="30" value="<%=amountdepreciation%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
				</td>
			</td>
			
			<td align="left" width="20%">Made In
				<td align="left">
				<input type="text" maxlength="15" name="madein" size="30" value="<%=madein%>" autocomplete="Off" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Amount
				<td align="left">
				<input type="text" maxlength="15" name="amount" size="30" value="<%=amount%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
				</td>
			</td>
			
			<td align="left" width="20%">Estimate Life Year
				<td align="left">
				<input type="text" maxlength="15" name="estimatelifeyear" size="30" value="<%=estimatelifeyear%>" autocomplete="Off" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Cost Center
				<td align="left">
				<input type="text" maxlength="15" name="costcenter" size="30" value="<%=costcenter%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
				</td>
			</td>
			<td align="left" width="20%">Tag Number
				<td align="left">
				<input type="text" maxlength="15" name="tagnumber" size="30" value="<%=tagnumber%>" autocomplete="Off" onkeyup="hanyaAngka(this);" >
				</td>
			</td>			
		</tr>
	
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="po_code" value="<%=po_code%>">
					<input type="hidden" name="itemcode" value="<%=itemcode%>">
					<input type="hidden" name="itemdescription" value="<%=itemdescription%>">
					<input type="hidden" name="acquisitionvalue" value="<%=acquisitionvalue%>">
					<input type="hidden" name="currencycode" value="<%=currencycode%>">
					<input type="hidden" name="receivenotecode" value="<%=receivenotecode%>">
					<input type="hidden" name="receivenotedate" value="<%=receivenotedate%>">
					<input type="hidden" name="suppliername" value="<%=suppliername%>">
					<input type="hidden" name="qty_hdr" value="<%=qty_hdr%>">
					<input type="hidden" name="currencyrate" value="<%=currencyrate%>">
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
		<td align="left">PO CODE</td>
		<td align="left">ASSET NAME</td>
		<td align="left">ACQUISITION VALUE</td>
		<td align="left">SUPPLIER NAME </td>
		<td align="left">QUANTITY</td>
		<td align="left">FA GROUP ID</td>
		<td align="left">FA GROUP NAME</td>
		<td align="left">QUANTITY FIX</td>
		<td align="left">AMOUNT DEPRECIATION</td>
		<td align="left">ESTIMATE LIFE YEAR</td>
		<td align="left">INSTALL DATE</td>	
		<td align="left">START DEPRECIATION</td>	
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
				<a href="assetadditional.asp?itemcode=<%=rsVH("itemcode")%>&receivenotedate=<%=rsVH("receivenotedate")%>&receivenotecode=<%=rsVH("receivenotecode")%>&po_code=<%=rsVH("po_code")%>&vouchernumber=<%=rsVH("vouchernumber")%>&serialnumber=<%=rsVH("serialnumber")%>&amountdepreciation=<%=rsVH("amountdepreciation")%>&estimatelifeyear=<%=rsVH("estimatelifeyear")%>&qty_dtl=<%=rsVH("qty_dtl")%>&qty_hdr=<%=rsVH("qty_hdr")%>&fa_groupid=<%=rsVH("fa_groupid")%>&fa_groupname=<%=rsVH("fa_groupname")%>&acquisitionvalue=<%=rsVH("acquisitionvalue")%>&amount=<%=rsVH("amount")%>&currencycode=<%=rsVH("currencycode")%>&suppliername=<%=rsVH("suppliername")%>&itemdescription=<%=rsVH("itemdescription")%>&additionvalue=<%=rsVH("additionvalue")%>&installdate=<%=rsVH("installdate")%>&startdepreciation=<%=rsVH("startdepreciation")%>&location_id=<%=rsVH("location_id")%>&costcenter=<%=rsVH("costcenter")%>&specification=<%=rsVH("specification")%>&residualvalue=<%=rsVH("residualvalue")%>&brand=<%=rsVH("brand")%>&madein=<%=rsVH("madein")%>&tagnumber=<%=rsVH("tagnumber")%>&currencyrate=<%=rsVH("currencyrate")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
		
				<a href="assetadditional.asp?itemcode=<%=rsVH("itemcode")%>&receivenotedate=<%=rsVH("receivenotedate")%>&receivenotecode=<%=rsVH("receivenotecode")%>&po_code=<%=rsVH("po_code")%>&vouchernumber=<%=rsVH("vouchernumber")%>&serialnumber=<%=rsVH("serialnumber")%>&amountdepreciation=<%=rsVH("amountdepreciation")%>&estimatelifeyear=<%=rsVH("estimatelifeyear")%>&qty_dtl=<%=rsVH("qty_dtl")%>&qty_hdr=<%=rsVH("qty_hdr")%>&fa_groupid=<%=rsVH("fa_groupid")%>&fa_groupname=<%=rsVH("fa_groupname")%>&acquisitionvalue=<%=rsVH("acquisitionvalue")%>&amount=<%=rsVH("amount")%>&currencycode=<%=rsVH("currencycode")%>&suppliername=<%=rsVH("suppliername")%>&itemdescription=<%=rsVH("itemdescription")%>&additionvalue=<%=rsVH("additionvalue")%>&installdate=<%=rsVH("installdate")%>&startdepreciation=<%=rsVH("startdepreciation")%>&location_id=<%=rsVH("location_id")%>&costcenter=<%=rsVH("costcenter")%>&specification=<%=rsVH("specification")%>&residualvalue=<%=rsVH("residualvalue")%>&brand=<%=rsVH("brand")%>&madein=<%=rsVH("madein")%>&tagnumber=<%=rsVH("tagnumber")%>&currencyrate=<%=rsVH("currencyrate")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
			    <td align="left" title="<%=rsVH("itemcode").Name%>" valign="top"><%=rsVH("itemcode")%></td>
			    <td align="left" title="<%=rsVH("po_code").Name%>" valign="top"><%=rsVH("po_code")%></td>
			    <td align="left" title="<%=rsVH("itemdescription").Name%>" valign="top"><%=rsVH("itemdescription")%></td>
			    <td align="left" title="<%=rsVH("acquisitionvalue").Name%>" valign="top"><%=rsVH("acquisitionvalue")%></td>
			    <td align="left" title="<%=rsVH("suppliername").Name%>" valign="top"><%=rsVH("suppliername")%></td>
			    <td align="left" title="<%=rsVH("qty_hdr").Name%>" valign="top"><%=rsVH("qty_hdr")%></td>
			    <td align="left" title="<%=rsVH("fa_groupid").Name%>" valign="top"><%=rsVH("fa_groupid")%></td>
			    <td align="left" title="<%=rsVH("fa_groupname").Name%>" valign="top"><%=rsVH("fa_groupname")%></td>
			    <td align="left" title="<%=rsVH("qty_dtl").Name%>" valign="top"><%=rsVH("qty_dtl")%></td>
			    <td align="left" title="<%=rsVH("amountdepreciation").Name%>" valign="top"><%=rsVH("amountdepreciation")%></td>
			    <td align="left" title="<%=rsVH("estimatelifeyear").Name%>" valign="top"><%=rsVH("estimatelifeyear")%></td>
			    <td align="left" title="<%=rsVH("installdate").Name%>" valign="top"><%=rsVH("installdate")%></td>
			    <td align="left" title="<%=rsVH("startdepreciation").Name%>" valign="top"><%=rsVH("startdepreciation")%></td>
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
