<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Invoice</title>
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

var getlistpocode = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedpocode = document.getElementById("selectedpocode").value;
	var PostStr = 'fcba='+selectedfcba+'&pocode='+selectedpocode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getpocodenotinvoicedbasedonfcba.asp', 'POST', PostStr, showlistpocode);
}

// The function for handling the response from the server
var showlistpocode = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("pocode").innerHTML = "";
	document.getElementById("pocode").innerHTML = response;
}

var getpoinfo = function () {
	// build up the post string when passing variables to the server side page

	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedpocode = document.getElementById("selectedpocode").value;
	var PostStr = 'fcba='+selectedfcba+'&pocode='+selectedpocode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetPOInfo.asp', 'POST', PostStr, showpoinfo);
}

// The function for handling the response from the server
var showpoinfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    //alert (s.postingdate);
    
    // update the Div to show the result from the server
	document.getElementById("suppliercode").value = s.suppliercode;
	document.getElementById("suppliername").value = s.suppliername;
	document.getElementById("downpayment").value = s.dp;
	document.getElementById("vat").value = s.ppn;
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
 
 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks,txtNoPO,v_CrossCekPO
 Dim recordTotalAmount, rsrecordTotalAmount, varTotalAmount, updateNetBalance
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""


 Dim Jobcode, LocationType, LocationCode, Reference, Remark

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 ' if request("fccOde")="" then
    ' fcCode="" '----- Nantinya Ambil dari Format Master
 ' else
    ' fcCode = request("fccOde")
 ' end if
 
 fcCode = request("fccOde")
If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("DATECREATED")<>"" then
   DATECREATED   =Request("DATECREATED")
 else
   DATECREATED   =textDate(date)
 end if
 if Request("fakturdate")<>"" then
   fakturdate   =Request("fakturdate")
 else
   fakturdate   =textDate(date)
 end if
 xCode = Request("xCode")
 HeaderRemarks = Request("HeaderRemarks")
 
 '---- Data Details
 JobCode=Request("JobCode")
 LocationType=Request("LocationType")
 LocationCode=Request("LocationCode")
 Reference=Request("Reference")
 Remarks=Request("Remarks")
 Amount=Request("Amount")
 
pocode = request("pocode")
fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
selectedpocode = Request("selectedpocode")

vat = Request("vat")
amountvat = Request("amountvat")
totalamount = Request("totalamount")
refinvoice = Request("refinvoice")

no_faktur = Request("no_faktur")
fakturdate = Request("fakturdate")

downpayment = Request("downpayment")
othercharge = Request("othercharge")

sub post_amount
	p_amountbeforevat = 0 : p_othercharge = 0 : p_totalamount = 0
	recordTotalAmount = " SELECT NVL(SUM(amountbeforevat),0) amountbeforevat" & _
						"        , NVL(SUM(othercharge), 0) othercharge" & _
						"        , NVL(SUM(amount), 0) totalAmount" & _
						" FROM invoicedetail" & _
						" WHERE invoicecode = '" & fcCode & "'" & _ 
						"       AND fcba = '" & fcba & "'"

	Set rsrecordTotalAmount = Server.CreateObject("ADODB.Recordset")
	rsrecordTotalAmount.Open recordTotalAmount, connstr

	if Not rsrecordTotalAmount.Eof then
		p_amountbeforevat = cdbl(to_nol(rsrecordTotalAmount("amountbeforevat")))
		p_othercharge = cdbl(to_nol(rsrecordTotalAmount("othercharge")))
		p_totalamount = cdbl(to_nol(rsrecordTotalAmount("totalAmount")))
	Else
		p_amountbeforevat = 0
	End if
	
	amountvat = p_amountbeforevat + (p_amountbeforevat *cdbl(to_nol(vat))/100)
	totalamount = amountvat +p_othercharge


	q = " UPDATE invoice" & _
		" SET amount = '" & p_amountbeforevat & "'" & _
		"     , amountvat = '" & amountvat & "'" & _
		"     , totalamount = '" & totalamount & "'" & _
		"     , othercharge = '" & p_othercharge & "'" & _
		"     , lastupdate = SYSDATE " & _
		"     , lasttime = TO_CHAR(SYSDATE,'HH24:MI') "& _												
		" WHERE invoicecode = '" & fcCode & "'" & _
		"       AND fcba = '" & fcba & "'"

	connStr.execute(q)

end sub
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	
	REMARKS=request("REMARKS")
	
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
			arr_dtl_eof = true
			q = " SELECT receivenotecode, pocode, purchaselinenokey FROM invoicedetail WHERE invoicecode = '" & fcCode & "' AND fcba = '" & fcba & "'"
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open q, connstr
			if not rsPO.eof then
				arr_dtl = rsPO.getrows()
				arr_dtl_eof = false
			end if
			rsPO.close : set rsPO = nothing
			
			if not arr_dtl_eof then
				for xyz=0 to ubound(arr_dtl,2)
					q = " DELETE FROM invoicedetail" & _
						" WHERE invoicecode = '" & fcCode & "'" & _
						"       AND receivenotecode = '" & arr_dtl(0,xyz) & "'" & _
						"       AND pocode = '" & arr_dtl(1,xyz) & "'" & _
						"       AND purchaselinenokey = '" & arr_dtl(2,xyz) & "'" & _
						"       AND fcBA = '" & fcba & "'"
					'response.write q
					connStr.execute(q)
					
					q = " UPDATE receivenotedetail SET amountinvoice = 0" & _
						" WHERE receivenotecode = '" & arr_dtl(0,xyz) & "'" & _
						"       AND purchaselinenokey = '" & arr_dtl(2,xyz) & "'" & _
						"       AND fcBA = '" & fcba & "'"
					'response.write q
					connStr.execute(q)
				next
				erase arr_dtl
			end if
			arr_dtl_eof = true

			q = " DELETE FROM invoice WHERE invoicecode = '" & fcCode & "' AND fcba = '" & fcba & "'"
			connStr.execute(q)
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		
		cmdView	= "True"

	elseif Request("txtJenis")= "Details" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then

			q = " DELETE FROM invoicedetail" & _
			    " WHERE invoicecode = '" & fcCode & "'" & _
			    "       AND receivenotecode = '" & Request("receivenotecode") & "'" & _
			    "       AND pocode = '" & pocode & "'" & _
			    "       AND purchaselinenokey = '" & Request("purchaselinenokey") & "'" & _
			    "       and fcBA='" & fcba & "'"
			'response.write q
			connStr.execute(q)
			
			q = " UPDATE receivenotedetail SET amountinvoice = 0" & _
			    " WHERE receivenotecode = '" & Request("receivenotecode") & "'" & _
			    "       AND purchaselinenokey = '" & Request("purchaselinenokey") & "'" & _
			    "       and fcBA='" & fcba & "'"
			'response.write q
			connStr.execute(q)
			
			Call post_amount
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
			
		cmdView	= "True"

		'response.Write(q)
	elseif Request("txtJenis")="New" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
		If Cek_IsDailyClosing = "FALSE" then
		'--Add by Ukie to handle Auto Number of PR
				If Ucase(fcCode) = Ucase("(NEW)") Then
					oSql = " select nvl(controlnumber_api.get_nextnumber(upper('INVOICE')" & _
						   "                                         , " & year(textDate(DATECREATED)) & "" & _
						   "                                         , " & month(textDate(DATECREATED)) & "" & _
						   "                                         , '" & fcba & "'" & _
						   " ), 'FALSE') INVOICE from dual"
					
					' Response.Write oSql
						   
					Set oRs = Server.CreateObject("ADODB.Recordset")
					oRs.Open oSql, connstr
					If Not oRs.EOF Then
						txtINVOICE = oRs("INVOICE")
					End If
					oRs.Close
					Set oRs = Nothing
						fcCode = txtINVOICE
				End If
				
				If txtINVOICE = "FALSE" Then
					err.Number = 1001
					err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
				End If
				
				'Response.Write PRCODE & txtPRCode
			'--End Added
			
			Dim rsPO
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			rsPO.Open 	" Select INVOICECODE, suppliercode from INVOICE where INVOICECODE = '" & fcCode & "'" & _
						" and fcBA = '" & fcba & "'", connstr
			

			IF rsPO.eof then	    
				Dim txtCodeSupplier,txtCodePO
				q = " Select distinct POCODE,SUPPLIERCODE From RECEIVENOTE where pocode = '" & pocode & "'"  & _
					" and fcBA = '" & fcba & "'"
				'response.write q
				Dim rsLPO  
				set rsLPO = Server.CreateObject("ADODB.Recordset")
				rsLPO.Open q, connstr
				if not rsLPO.eof then
					txtCodePO=rsLPO("POCODE")
					txtCodeSupplier=rsLPO("SUPPLIERCODE")			
				End if 
			
				rsLPO.Close
				set rsLPO = Nothing
				q = " Insert into Invoice (INVOICECODE, REFINVOICE, INVOICEDATE, SUPPLIERCODE, PAID, REMARKS, NETBALANCE, pocode, no_faktur, fakturdate, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, vat, dp) " & _
					" values ('" & fcCode & "','" & refinvoice & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & txtCodeSupplier& "',0,'',0,'" & pocode& "','" & no_faktur & "', TO_DATE('" & fakturdate & "', 'mm/dd/yyyy'),'" & _
					Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'), '" & vat & "', REPLACE('" & downpayment & "',',',''))"
				'response.write q
				connStr.execute(q)
			End if
			rsPO.Close	
			Dim n
			'response.Write "No " & (cdbl(request("no"))-1) & " amount : " &  request("AmountInvoice_"&TRIM(n))
			For n=1 to (cdbl(request("no"))-1) 
				if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
						q = " Insert into  INVOICEDETAIL (INVOICECODE, ITEMCODE, QUANTITY, AMOUNT, POCODE, receivenotecode, purchaselinenokey, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME, amountbeforevat, othercharge) " & _
							" values ('" & fcCode & "','" & request("Stock_"&TRIM(n)) & "','" & request("Quantity_"&TRIM(n)) &  _
							"',REPLACE('" & request("AmountInvoice_"&TRIM(n)) & "', ',', ''),'" & pocode & "','" & request("receivenotecode_"&TRIM(n)) & "','" & request("linenokey_"&TRIM(n)) & "'," & _
							"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'),'" & request("amount_"&TRIM(n)) & "','" & request("othercharge_"&TRIM(n)) & "')"
						connStr.execute(q)
						
						q = " Update ReceiveNoteDetail  set AMOUNTINVOICE=AMOUNTINVOICE+REPLACE('" & request("AmountInvoice_"&TRIM(n)) & "', ',', '')"  & _
							", LASTUPDATE = sysdate " & _
							", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							" where RECEIVENOTECODE='" & request("receivenotecode_"&TRIM(n)) & "' and purchaselinenokey = '" & request("linenokey_"&TRIM(n)) & "'" & _
							" and fcBA='" & fcba & "'"
						'response.Write(q)
						connStr.execute(q)
						
				End if	 
			Next
			
			Call post_amount
			
		else
			err.number = 1001
			err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		cmdView	= "True"
	else
		cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
	cmdView	= request("cmdView")
End if

 Dim rsVoucher,SudahClose
 q = " select hdr.*, bu.fccompanycode" & _
     " from Invoice hdr" & _
     " left join businessunit bu on bu.fccode = hdr.fcba" & _
     " where hdr.invoicecode='" & fcCode & "'"
 
 set rsVoucher = server.CreateObject("Adodb.recordset")
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED = rsVoucher("INVOICEDATE")
	
    fccompanycode = rsVoucher("fccompanycode")
	fcba = rsVoucher("fcba")
	selectedfcba = rsVoucher("fcba")
	
	pocode = rsVoucher("pocode")
	refinvoice = rsVoucher("refinvoice")
	selectedpocode = rsVoucher("pocode")
	
	SudahClose = "Y"
	
	downpayment = rsVoucher("dp")
	amount = rsVoucher("amount")
	vat = rsVoucher("vat")
	amountvat = rsVoucher("amountvat")
	othercharge = rsVoucher("othercharge")
	totalamount = rsVoucher("totalamount")
	
	no_faktur = rsVoucher("no_faktur")
	fakturdate = rsVoucher("fakturdate")
 End if
 rsVoucher.Close : Set rsVoucher = Nothing
 

 if err.number <> 0 then
	connStr.RollbackTrans
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
 End if
 '-------------ISI Payment Form
 
 Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Purchase')"
	   
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
 <form name="frmEntry" method="post" action="Invoice.Asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">DP</td>
			<td align="left"><input type="text" id="downpayment" name="downpayment" size="30" value="<%=downpayment%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>
		</tr>
		<tr>
			<td align="left" style="width:20%;">Ref. Invoice No</td>
			<td align="left"><input type="text" maxlength="50" name="refinvoice" size="30" value="<%=refinvoice%>"><font color="#FF0000">&nbsp;*</font></td>
			
			<td align="left" style="width:20%;">Amount</td>
			<td align="left"><input type="text" id="amount" name="amount" size="30" value="<%=amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="DATECREATED" size="10" value="<%=DATECREATED%>">
			<font color="#FF0000">&nbsp;*</font>
			</td>
			
			<td align="left" style="width:20%;">VAT</td>
			<td align="left"><input type="text" id="vat" name="vat" size="30" value="<%=vat%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>
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
			
			<td align="left" style="width:20%;">Amount + VAT</td>
			<td align="left"><input type="text" id="amountvat" name="amountvat" size="30" value="<%=amountvat%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<select id="fcba" name="fcba" onchange="javascript:getselectedfcba();javascript:getlistpocode();">
					<option value=""></option>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
			
			<td align="left" style="width:20%;">Total Charge</td>
			<td align="left"><input type="text" id="othercharge" name="othercharge" size="30" value="<%=othercharge%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>
		</tr>
		<tr>
			<td align="left">PO Code</td>
			<td align="left">      
			<%if SudahClose="Y" then%>
				<select name="pocode" onchange="javascript:getpoinfo();"><option value="<%=pocode%>" Selected><%=pocode%></option></select>
				
			<%else%>
				<select id="pocode" name="pocode" onchange="javascript:getselectedpocode();javascript:getpoinfo();">
					<option value=""></option>
				</select>
			<%End if%>
			<input type="Hidden" id="selectedpocode" name="selectedpocode" value="<%=selectedpocode%>">
			<font color="#FF0000">&nbsp;*</font>
			</td>
			
			<td align="left" style="width:20%;">Total Amount</td>
			<td align="left"><input type="text" id="totalamount" name="totalamount" size="30" value="<%=totalamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly="true" style="background-color:#eee;"></td>
			
		</tr>
		<tr>
			<td align="left" style="width:20%;">Supplier Code</td>
			<td align="left"><input type="text" id="suppliercode" name="suppliercode" size="30" value="<%=suppliercode%>" readonly="true" style="background-color:#eee;"></td>
			
			
		</tr>
		
		<tr>
			<td align="left" style="width:20%;">Supplier Name</td>
			<td align="left"><input type="text" id="suppliername" name="suppliername" size="30" value="<%=suppliername%>" readonly="true" style="background-color:#eee;"></td>
		</tr>
		
		<tr>
			<td align="left">No Faktur</td>
			<td align="left">
				<input type="text" maxlength="100" name="no_faktur" size="30" value="<%=no_faktur%>" autocomplete="off" >
			</td>
			<td align="left"> Faktur Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="fakturdate" size="10" value="<%=fakturdate%>" autocomplete="Off">
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="print_invoice.html?invoicecode=<%=fcCode%>&fcba=<%=fcba%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" Then%>
					<a href="Invoice.Asp?fcCode=<%=fcCode%>&fcba=<%=fcba%>&DATECREATED=<%=DATECREATED%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End if%>
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Purchase"
				dataSubmodule 	= "Invoice"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 and len(trim(pocode)) > 0 Then
   Dim rsGRN,intRecordsPerPage,cColor
   if SudahClose="Y" Then
      q = " select sp.fcname suppdesc, dtl.itemcode, s.fcname stockdesc, dtl.quantity, dtl.amount, dtl.receivenotecode, dtl.purchaselinenokey" & _
          " FROM invoicedetail dtl" & _
          " INNER JOIN invoice hdr ON (dtl.invoicecode = hdr.invoicecode AND dtl.fcba = hdr.fcba)" & _
          " LEFT JOIN stock s ON (dtl.itemcode = s.fccode AND dtl.fcba = s.fcba)" & _
          " LEFT JOIN supplier sp ON (hdr.suppliercode = sp.fccode AND hdr.fcba = sp.fcba)" & _
          " WHERE dtl.INVOICECODE = '" & fcCode & "'" & _
          "       AND dtl.fcBA = '" & fcba & "'"
   Else	 
      q = " SELECT hdr.pocode" & _
	      "        , dtl.receivenotecode" & _
	      "        , dtl.purchaselinenokey" & _
	      "        , dtl.purchaseitemcode" & _
	      "        , s.fcname" & _
	      "        , NVL(dtl.quantity,0) quantity" & _
	      "        , NVL(dtl.amount,0) amount" & _
	      "        , NVL(dtl.othercharge,0) othercharge" & _
	      "        , NVL(dtl.amountreceived,0) amountreceived" & _
	      " FROM receivenotedetail dtl" & _
	      " INNER JOIN receivenote hdr ON (hdr.receivenotecode = dtl.receivenotecode AND hdr.fcba = dtl.fcba)" & _
	      " LEFT JOIN stock s ON (s.fcba = dtl.fcba AND s.fccode = dtl.purchaseitemcode)" & _
	      " WHERE hdr.pocode = '" & pocode & "'" & _
	      "       AND dtl.fcBA='" & fcba & "'" & _
		  "       AND NVL(dtl.amountinvoice,0) = 0" & _
		  " ORDER BY dtl.receivenotecode, dtl.purchaseitemcode"
   End if
   'response.Write q	
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.Open q, connstr%>
   <form action="Invoice.Asp" method="post" name="eForm">
   <table class="data" border="1" width="100%">
   <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">PO Code</td>
	   <td align="center">GRN Code</td>
	   <td align="center">Stock Code</td>
	   <td align="center">Stock Desc.</td>
	   <td align="center">Qty Received</td>
	   <td align="center">Amount</td>
	   <td align="center">Other Charge</td>
	   <td align="center">Total</td>
	   <td align="center"><input type="checkbox" name="select-all" id="select-all"/> Check/Uncheck All</td>
	   <td align="center">Amount Invoice</td>
	   </tr>
   <%ELSE%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">&nbsp;</td>
	   <td align="center">GRN Code</td>
	   <td align="center">Stock Code</td>
	   <td align="center">Stock Desc.</td>
	   <td align="center">Quantity</td>
	   <td align="center">Amount</td>
	   </tr>
   <%End if
   if Not rsGRN.eof then
      nCountofField=rsGRN.fields.count-1
	  no=1
	  While Not rsGRN.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="ringht" valign="top"><%=NO%></td>
			<%IF SudahClose="Y" then%>
				<td valign="middle">
				<%If val_IsTrue = "Y" Then%>
					<a href="Invoice.Asp?fcCode=<%=fcCode%>&pocode=<%=pocode%>&fcba=<%=fcba%>&receivenotecode=<%=rsGRN("receivenotecode")%>&purchaselinenokey=<%=rsGRN("purchaselinenokey")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0"></a></td>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End if%>
					<td align="left" title="Supplier" valign="top"><%=rsGRN("receivenotecode")%></td>
					<td align="left" title="Item Code" valign="top"><%=rsGRN("ITEMCODE")%></td>
					<td align="left" title="Item Name" valign="top"><%=rsGRN("StockDesc")%></td>
					<td align="right" title="Quantity" valign="top"><%=rsGRN("Quantity")%></td>
					<td align="right" title="Amount" valign="top"><%=FormatNumber(rsGRN("Amount"),0)%></td>
			<%End if%>

			<%if SudahClose<>"Y" Then%>
                    <td align="left" title="" valign="top"><%=rsGRN("POCODE")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("RECEIVENOTECODE")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("PURCHASEITEMCODE")%></td>
			        <td align="left" title="" valign="top"><%=rsGRN("fcname")%></td>
			        <td align="right" title="" valign="top"><%=rsGRN("Quantity")%></td>
                    <td align="right" title="" valign="top"><%=FormatNumber(rsGRN("amount"),0)%></td>
                    <td align="right" title="" valign="top"><%=FormatNumber(rsGRN("othercharge"),0)%></td>
                    <td align="right" title="" valign="top"><%=FormatNumber(rsGRN("amountreceived"),0)%></td>
		            <td align="center">
						<input type="checkbox" id="cb_<%=no%>" name="cb_<%=no%>" class="item" value = "" onclick="handleClick(this);" onchange="CheckCheckboxes1(this)">
					</td>
					<td align="center">
						<input style="text-align:right;background-color:#eee;" type="text" name="AmountInvoice_<%=no%>" class="amount" value = "<%=FormatNumber(rsGRN("amountreceived"),0)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readonly>
					</td>
					<input type="hidden" name="receivenotecode_<%=no%>" value="<%=rsGRN("receivenotecode")%>">
					<input type="hidden" name="linenokey_<%=no%>" value="<%=rsGRN("purchaselinenokey")%>">
					<input type="hidden" name="Stock_<%=no%>" value="<%=rsGRN("purchaseitemcode")%>">
					<input type="hidden" name="Quantity_<%=no%>" value="<%=rsGRN("quantity")%>">
					<input type="hidden" name="amount_<%=no%>" value="<%=rsGRN("amount")%>">
					<input type="hidden" name="othercharge_<%=no%>" value="<%=rsGRN("othercharge")%>">
				 </td>
			<%End if%>
		  </tr>
		  <%
		  no=no+1
		  rsGRN.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsGRN.Close
   set rsGRN = Nothing%>
	  <tr><td colspan="11" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="refinvoice" value="<%=refinvoice%>">
		  <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		  <input type="hidden" name="pocode" value="<%=pocode%>">
		  <input type="hidden" name="No" value="<%=No%>">
		  
		  <input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
		<input type="hidden" name="fcba" value="<%=fcba%>">
		<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
		<input type="hidden" name="selectedpocode" value="<%=selectedpocode%>">
		
		<input type="hidden" name="downpayment" value="<%=downpayment%>">
		<input type="hidden" name="amount" value="<%=amount%>">
		<input type="hidden" name="vat" value="<%=vat%>">
		<input type="hidden" name="amountvat" value="<%=amountvat%>">
		<input type="hidden" name="totalamount" value="<%=totalamount%>">
		
		<input type="hidden" name="no_faktur" value="<%=no_faktur%>">
		<input type="hidden" name="fakturdate" value="<%=fakturdate%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Submit">
		  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </TABLE>
   </div>
   </form>
<%End if%>  
</div></div>
  <!--#include file="bottom.jsp"-->
		<script src="/js/jquery.inputmask.bundle.min.js"></script>
		<script src="/js/inputmask.binding.js"></script>
<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
}

function getselectedpocode() {
	var spocode = document.getElementById("selectedpocode");
	spocode.value = document.getElementById("pocode").value;
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
	getlistpocode();
	getpoinfo();
});
</script>
<script type="text/javascript">
	function handleClick(cb){
		if(cb.checked == 1){
			cb.value = 1;
		} else {
			cb.value = 0;
		}
	}
	
	function CheckCheckboxes1(chk){
		var txt = chk.parentNode.parentNode.cells[8].getElementsByTagName('input')[0];
		
		if(chk.checked == true)
		{
			txt.readOnly = false;
			txt.style.backgroundColor = "#fff";
		}
		else
		{
			txt.readOnly = true;
			txt.style.backgroundColor = "#eee";
		}
	}
</script>
<script type="text/javascript">	
	$('#select-all').click( function () {
		$('.item').prop('checked', this.checked);
		$('.item').val(this.checked ? 1 : 0);
		if($('#select-all').is(":checked")){
			$('.amount').css("background-color", "#fff");
			$('.amount').prop('readOnly', false);
		}else{
			$('.amount').css("background-color", "#eee");
			$('.amount').prop('readOnly', true);
		}
	})
</script>
</body>
</html>
