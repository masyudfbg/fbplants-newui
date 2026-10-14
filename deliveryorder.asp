<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Delivery Order </title>

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
<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">
<script language="JavaScript">
	window.addEventListener('load', function() {
		if (typeof xTableHeaderFixed !== 'undefined') {
			new xTableHeaderFixed('xthf-site', window);
			new xTableHeaderFixed('xthf-green', window);
			new xTableHeaderFixed('xthf-blue', window);
		}
	});
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


var getTransporter = function () {
	// build up the post string when passing variables to the server side page

	var lineno = document.getElementById("lineno").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var transportercontractno = document.getElementById("transportercontractno").value;
	var PostStr = 'lineno='+lineno+'&fcba='+selectedfcba+'&transportercontractno='+transportercontractno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetTransporter.asp', 'POST', PostStr, ShowTransporeterResponse);
}

// The function for handling the response from the server
var ShowTransporeterResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert (s.lineno);
    
    // update the Div to show the result from the server
	document.getElementById("transporter").value = s.transporter;
	document.getElementById("losstolerance").value = s.losstolerance;
	document.getElementById("rate").value = s.rate;
	document.getElementById("transportername").value = s.transportername;

}

var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistfcbadosales.asp', 'POST', PostStr, showlistfcba);
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
		  '.chosen-select-width'     : {width:"45%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getTransporterLine = function () {
	// build up the post string when passing variables to the server side page
	
	var transportercontractno = document.getElementById("transportercontractno").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var selectedtransporterline = document.getElementById("selectedtransporterline").value;
	var PostStr = 'transportercontractno='+transportercontractno+'&fcba='+selectedfcba+'&lineno='+selectedtransporterline;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetTransporterLine.asp', 'POST', PostStr, showTransporterLine);
}

// The function for handling the response from the server
var showTransporterLine = function (oXML) { 
    // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("transporterlinelist").innerHTML = "";
	document.getElementById("transporterlinelist").innerHTML = response;

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
 <!-------------awal!------------->
 <%


' Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, LineNoHist, ITEMCODE
' DIM v_NURSERYCODE,TRANSDATE
' DIM v_QUANTITY
' Dim cmdView
' Dim qSupplyingDeclaration, rsNurseryInfo, tempStoreCode, tempStockCode, doneError
' Dim qLineMax, rsqLineMax, tempMaxLine
' Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
' Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
' Dim qLessZero, rsLessZero, tempislesszero
' Dim done, qLineUpper, rsqLineUpper, totalUpper, qQuantity, rsqQuantity, totalQuantity, QuantityOriginal
' Dim qDeleteHist
' Dim rsReceiveDetail, originalQuantity, v_CrossCekDetail
Dim qSupplyingDeclaration, rsSupplyingDeclarationInfo, tempQuantity, tempDecType
Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, v_FCCODE, v_FCNAME
Dim qSupplyingInfo, rsSupplyingInfo, qDeclarationType, rsDeclarationType
Dim salescontractno, parcode
Dim productcode,  unitofmeasure, deliveryorderno
Dim quantity, quantity_do, transporter, transportercontractno, transportationcost, losstolerance
Dim REMARKS, paymentcondition, deliverycondition
Dim creationdate, startdate, enddate, deliverydate, deliveryorderdate
Dim cmdView
Dim txtDocumentNo ,txtlineno
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

salescontractno =request("salescontractno")
deliveryorderno =request("deliveryorderno")
quantity_do =request("quantity_do")
transportercontractno =request("transportercontractno")
transportationcost =request("transportationcost")
losstolerance =request("losstolerance")
startdate =request("startdate")
enddate =request("enddate")
quantity=request("quantity")
remarks=request("remarks")
transporter=request("transporter")
deliverycondition = Request("deliverycondition")
txtlineno = request("txtlineno")
productcode =request("productcode")
unitofmeasure =request("unitofmeasure")
productname = request("productname")
deliveryorderdate = request("deliveryorderdate")
creationdate = request("creationdate")
transporterline = request("transporterline")
lineno = request("lineno")

fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")

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

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")


If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, creationdate)
		If Cek_IsDailyClosing = "FALSE" then
			q = " delete from deliveryorder " & _
				" where fcBA = '" & fcba & "'" & _
				" and salescontractno = '" & salescontractno & "'"
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
			
			q = " delete from deliveryorder " & _
				" where fcBA = '" & fcba & "'" & _
				" and salescontractno = '"& salescontractno & "'" & _
				" and deliveryorderno = '"& deliveryorderno & "'" & _
				" and lineno = '"& lineno & "'" 
				' Response.Write q
				connStr.execute(q)

			
			cmdView	= "True"

	elseif Request("txtJenis")="New" then	

			Dim rsDeliveryOrder
			Set rsDeliveryOrder = Server.CreateObject("ADODB.Recordset")
			rsDeliveryOrder.Open " select salescontractno from salesagreementhdr where salescontractno = '" & salescontractno & "' and fcba ='" & fcba & "' ", connstr
			If rsDeliveryOrder.EOF Then
				
				
			End If
			rsDeliveryOrder.Close
			set rsDeliveryOrder = Nothing
					
			'If Len(Trim(deliveryorderno)) = 0 Then
			'	Err.Number = 1001
			'	Err.Description = "Delivery Order No Harus diisi, tidak boleh kosong"
			'Else
			'If cdbl(to_nol(quantity_do)) = 0 Then
			'	Err.Number = 1001
			'	Err.Description = "Quantity Order Harus diisi, tidak boleh kosong"
			'Else

			' If cdbl(to_nol(quantity_do)) > cdbl(to_nol(quantity)) Then
			' 	Err.Number = 1001
			' 	Err.Description = "Quantity Delivery Order tidak boleh lebih dari Quantity Contract"
			' Else
			
			strFieldDetail= " SELECT creationdate  FROM salesagreementhdr" & _
							" WHERE FCBA = '"& fcba &"' AND SALESCONTRACTNO = '"& salescontractno &"'" 
							' Response.Write strFieldDetail
							Set rsQty = Server.CreateObject("ADODB.Recordset")
							rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then 
				creationdate = rsQty("creationdate")  
			End If
			
			strFieldDetail= " SELECT count(1) total FROM DUAL WHERE TO_DATE('"& creationdate &"','mm/dd/yyyy') <=  TO_DATE('"& deliveryorderdate &"','mm/dd/yyyy')  " 
							' Response.Write strFieldDetail
							Set rsQty = Server.CreateObject("ADODB.Recordset")
							rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
			do while not rsQty.eof
					checkdo = cdbl (to_nol(rsQty("total")))
					rsQty.MoveNext
				loop
			End If
			If  cdbl(to_nol(checkdo)) = 0 Then
				Err.Number = 1001
				Err.Description = "Tanggal Delivery Order harus lebih besar sama dengan Tanggal Creation Date Sales Contract"
			Else
			
			strFieldDetail= " select nvl(sum(quantity_do),0) total from deliveryorder where fcba = '" & fcba & "' and salescontractno = '"& salescontractno &"' " 
			'Response.Write strFieldDetail
			Set rsQty = Server.CreateObject("ADODB.Recordset")
			rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
				checkqty = cdbl (to_nol(rsQty("total")))
			else
				checkqty = 0
			End If
			If ( cdbl(to_nol(checkqty)) +cdbl(to_nol(quantity_do)) ) > cdbl(to_nol(quantity)) Then
				Err.Number = 1001
				Err.Description = "Quantity Delivery Order tidak boleh lebih dari Quantity Contract"
			Else

			' strFieldDetail= " select sum(quantity_do) total from deliveryorder where fcba = '" & session("fcBA") & "' and salescontractno = '"& salescontractno &"' " 
			' 				' Response.Write strFieldDetail
			' 				Set rsQty = Server.CreateObject("ADODB.Recordset")
			' 				rsQty.Open strFieldDetail, connstr							
			' if not rsQty.Eof then
			' do while not rsQty.eof
			' 		checkqty = cdbl (to_nol(rsQty("total")))
			' 		rsQty.MoveNext
			' 	loop
			' End If
			' ' Response.Write checkqty
			' ' Response.Write quantity_do
			' If cdbl(to_nol(checkqty)) >= cdbl(to_nol(quantity)) Then
			' 	Err.Number = 1001
			' 	Err.Description = "Quantity Delivery Order tidak boleh lebih dari Quantity Contract"
			' Else
					qLineMax = " select nvl(max(lineno),0) doline" & _
												   " from deliveryorder" & _
												   " where salescontractno = '" & salescontractno & "'" & _
												   " and fcba = '" & fcba & "'" 			

					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						tempMaxLineDO = cdbl(rsqLineMax("doline"))
					else
						tempMaxLineDO = 0
					end if 

					tempMaxLineDO = tempMaxLineDO + 1
					txtLineNoHist = tempMaxLineDO

					q = " INSERT into deliveryorder" & _
					" (salescontractno, deliveryorderno, deliveryorderdate, productcode, quantity, unitofmeasure, startdate, enddate, quantity_do, deliverycondition, remarks, transportercontractno, transporter, transportationcost, losstolerance, fcentry, fcedit, fcip, fcba, lastupdate, lasttime, transporterline, lineno) values (" & _
					" '" & salescontractno & "'" & _
					" , '" & deliveryorderno & "'" & _
					" , to_date('" & deliveryorderdate & "','mm/dd/yyyy')" & _
					" , '" & productcode & "'" & _
					" , REPLACE('" & quantity & "', ',', '')" & _
					" , '" & unitofmeasure & "'" & _
					" , to_date('" & startdate & "','mm/dd/yyyy')" & _
					" , to_date('" & enddate & "','mm/dd/yyyy')" & _ 
					" , REPLACE('" & quantity_do & "', ',', '')" & _
					" , '" & deliverycondition & "'" & _
					" , '" & remarks & "'" & _
					" , '" & transportercontractno & "'" & _
					" , '" & transporter & "'" & _
					" , REPLACE('" & transportationcost & "', ',', '')" & _
					" , REPLACE('" & losstolerance & "', ',', '')" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & fcba & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" , '" & transporterline & "'" & _
					" , '" & tempMaxLineDO & "'" & _
					" )"								
				' Response.Write q
				connStr.execute(q)
				
					End If			
				End if
	'		End If
	'	End If	
	' End If		
	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			'If Len(Trim(deliveryorderno)) = 0 Then
			'	Err.Number = 1001
			'	Err.Description = "Delivery Order No Harus diisi, tidak boleh kosong"
			'Else
			'If cdbl(to_nol(quantity_do)) = 0 Then
			'	Err.Number = 1001
			'	Err.Description = "Quantity Order Harus diisi, tidak boleh kosong"
			'Else
			strFieldDetail= " SELECT count(1) total FROM DUAL" & _
							" WHERE TO_DATE('" & deliveryorderdate & "','mm/dd/yyyy') BETWEEN TO_DATE('" & startdate & "','mm/dd/yyyy') AND TO_DATE('" & enddate & "','mm/dd/yyyy')" 
							' Response.Write strFieldDetail
							Set rsQty = Server.CreateObject("ADODB.Recordset")
							rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
			do while not rsQty.eof
					checkdo = cdbl (to_nol(rsQty("total")))
					rsQty.MoveNext
				loop
			End If
			If cdbl(to_nol(checkdo)) = 0 Then
				Err.Number = 1001
				Err.Description = "Tanggal Delivery Order harus lebih besar sama dengan Tanggal Creation Date Sales Contract"
			Else
			
			edit_checkqty = 0
			strFieldDetail= " select nvl(sum(quantity_do),0) total from deliveryorder where fcba = '" & fcba & "' and salescontractno = '"& salescontractno &"' and deliveryorderno = '"& deliveryorderno &"' " 
			'Response.Write strFieldDetail
			Set rsQty = Server.CreateObject("ADODB.Recordset")
			rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
				edit_checkqty = cdbl (to_nol(rsQty("total")))
			else
				edit_checkqty = 0
			End If
			
			strFieldDetail= " select nvl(sum(quantity_do),0) total from deliveryorder where fcba = '" & fcba & "' and salescontractno = '"& salescontractno &"' " 
			'Response.Write strFieldDetail
			Set rsQty = Server.CreateObject("ADODB.Recordset")
			rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
				checkqty = cdbl (to_nol(rsQty("total")))
			else
				checkqty = 0
			End If
			
			If ( cdbl(to_nol(checkqty)) -cdbl(to_nol(edit_checkqty)) +cdbl(to_nol(quantity_do)) ) > cdbl(to_nol(quantity)) Then
				Err.Number = 1001
				Err.Description = "Quantity Delivery Order tidak boleh lebih dari Quantity Contract"
			Else
					q =	" Update deliveryorder" & _
						" set quantity_do = REPLACE('" & quantity_do & "', ',', '')" & _
						"     , deliverycondition = '" & deliverycondition & "'" & _
						"     , remarks = '" & remarks & "'" & _
						"     , deliveryorderdate = to_date('" & deliveryorderdate & "','mm/dd/yyyy')" & _
						"     , startdate = to_date('" & startdate & "','mm/dd/yyyy')" & _
						"     , enddate = to_date('" & enddate & "','mm/dd/yyyy')" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" 		where salescontractno = '" & salescontractno & "'" & _
						"       and fcBA = '" & fcba & "'" & _
						" 		and deliveryorderno = '"& deliveryorderno & "'" & _ 
						" 		and lineno = '"& request("lineno") & "'" 
		 			' response.write q
					connStr.execute(q)
				End If
		'	End If				
		'End If
	End If
		cmdView	= "True"
		
	else
		cmdView	= "True"
End If
else
	cmdView	= request("cmdView")
End If

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis



'untuk menampilkan datepicker default
If Len(Trim(creationdate)) = 0 Then	creationdate = textDate(date)
If Len(Trim(deliverydate)) = 0 Then	deliverydate = textDate(date)
If Len(Trim(deliveryorderdate)) = 0 Then	deliveryorderdate = textDate(date)
If Len(Trim(startdate)) = 0 Then startdate = textDate(date)
If Len(Trim(enddate)) = 0 Then	enddate = textDate(date)

oSQLQuery = " select sa.*, bu.fccompanycode" & _
            " from salesagreementhdr sa left join businessunit bu on bu.fccode = sa.fcba" & _
	        " where sa.salescontractno = '" & salescontractno & "' and sa.fcba = '"& fcba &"'"
			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	fccompanycode = oRecordSet("fccompanycode")
	fcba = oRecordSet("fcba")
	selectedfcba = oRecordSet("fcba")
End if

If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
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
 
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-truck" aria-hidden="true"></i>Header Delivery Order</div><% end if %>
	<form name="frmEntry" method="post" action="deliveryorder.asp">
		<table class="data" width="60%">
	
		<tr>
			<td align="left" width="20%">Sales Contract No 
				<td align="left">
				<input type="text" maxlength="50" name="salescontractno" size="20" value="<%=salescontractno%>" autocomplete="Off">
				<input type="hidden" maxlength="50" name="username" size="20" value="<%=Session("username")%>" autocomplete="Off">
				<input type="hidden" maxlength="50" name="fcbasession" size="20" value="<%=Session("fcba")%>" autocomplete="Off">
				<font color="#FF0000">*</font>
				</td>
			</td>
        </tr>
		
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();" class="chosen-select" style="width: 30%">
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
				<div id="fcbalist"></div>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
		</tr>

		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				
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
if cmdView = "True" and len(trim(salescontractno)) > 0  Then
	Dim rsVH, intRecordsPerPage, cColor
	
'query Untuk menampilkan isi di Field
	q	= " SELECT sad.salescontractno, sad.productcode, sad.unitofmeasure, do.quantity_do, do.deliverycondition, do.remarks, do.deliveryorderdate, do.deliveryorderno, do.lineno, " & _
		  " sad.salesprices, sad.startdate, sad.enddate, sad.amount, sad.amountvat, sad.totalamountvat, sad.quantity, sad.vat, " & _
		  " sah.customercode, sah.creationdate, sah.deliverydate, sah.deliverytype, " & _
		  " prod.productdescription as productname, prod.unitofmeasure, " & _
		  " tc.transportercontractno, tc.transporter, tc.losstolerance, do.transportationcost, tc.shippingaddress,supplier_api.Get_Transporter_Name(tc.transporter,tc.fcba) AS transportername " & _
		  " FROM salesagreementdet sad" & _
		  " INNER JOIN salesagreementhdr sah ON (sad.salescontractno = sah.salescontractno AND sad.fcba = sah.fcba)" & _
		  " LEFT JOIN deliveryorder do ON (sad.salescontractno = do.salescontractno AND sad.fcba = do.fcba)" & _
 		  " LEFT JOIN product prod ON (prod.productcode = sad.productcode AND sad.fcba = prod.fcba)" & _
		  " LEFT JOIN transportercontract tc ON (do.transportercontractno = tc.transportercontractno AND do.fcba = tc.fcba)" & _
		  " WHERE sad.fcba = '" & fcba & "'" & _
		  " AND sad.salescontractno = '" & salescontractno & "'" 

	' response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr

	if Not rsVH.eof then
		productcode =  rsVH("productcode")
		productname =  rsVH("productcode") & " : " & rsVH("productname")
		quantity =  rsVH("quantity")
		customercode =  rsVH("customercode")
		unitofmeasure =  rsVH("unitofmeasure")
		' lineno =  rsVH("lineno")
		startdate =  rsVH("startdate")
		enddate =  rsVH("enddate")
		shippingaddress =  rsVH("shippingaddress")
	end if


'Query untuk menampilkan table Delivery Order
	q	= " SELECT a.*,supplier_api.Get_Transporter_Name(a.transporter,a.fcba) AS transportername FROM deliveryorder a " & _
		  " WHERE fcba = '" & fcba & "'" & _
		  " AND salescontractno = '" & salescontractno & "'" 

	' response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr

%>
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
	<form action="deliveryorder.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr>
			<td align="left" width="20%">Delivery Order No 
				<td align="left">
				<input type="text" maxlength="50" style="background:#FFFFCC;" name="deliveryorderno" size="30" value="<%=deliveryorderno%>" autocomplete="Off">
				<input type="hidden" maxlength="50" style="background:#FFFFCC;" name="customerid" size="30" value="<%=customercode%>" autocomplete="Off">
				<font color="#FF0000">*</font>
				</td>
			</td>

			<td align="left" width="20%">Quantity Delivery Order
				<td align="left">
				<input type="text" maxlength="15" name="quantity_do" size="30" value="<%=quantity_do%>" autocomplete="Off" onkeyup="hanyaAngka(this);">
				<font color="#FF0000">*</font>
				</td>
			</td>
			
		</tr>
		<tr>
			<td align="left" width="20%">Delivery Order Date 
				<td align="left">
				<input class="datepicker" type="hidden" size="15" name="creationdate" value="<%=creationdate%>">
				<input class="datepicker" type="text" size="15" name="deliveryorderdate" value="<%=deliveryorderdate%>">
			</td>

			<td align="left" width="20%">Delivery Condition
				<td align="left">
				<textarea name="deliverycondition" size="30"><%=deliverycondition%></textarea>
			
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Product 
				<td align="left">
				<input type="hidden" maxlength="15" name="productcode" style="background:#eee;" size="30" value="<%=productcode%>" autocomplete="Off">
				<input type="text" maxlength="15" name="productname" style="background:#eee;" size="30" value="<%=productname%>" autocomplete="Off" readonly="true">
				</td>
			</td>

			<td align="left" width="20%">Remarks
				<td align="left">
					<textarea name="remarks"><%=remarks%></textarea>
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Quantity
				<td align="left">
				<input type="text" maxlength="15" name="Quantity" style="background:#eee;" size="30" value="<%=Quantity%>" autocomplete="Off" readonly="true">
				</td>
			</td>

			<td align="left" width="20%">Transporter Contract
				<td align="left">
				<%IF txtJenis="Edit" Then%>
					<input type="text" size="15" name="transportercontractno" value="<%=transportercontractno%>" readonly style="background:#eee">
				<%else%>

					<select name="transportercontractno" id="transportercontractno" style="background:#FFFFCC;"  title="TRANSPORTER CODE Character a s/d z or 1 s/d 9" onchange="javascript:getTransporterLine();" class="chosen-select">
						<option value=""></option> 
							<% 	Dim rsComboP
								q = "SELECT DISTINCT TRANSPORTERCONTRACTNO FROM TRANSPORTERCONTRACT WHERE fcba = '" & fcba & "'"
								Set rsComboP = Server.CreateObject("ADODB.Recordset")
								rsComboP.CursorType = 3
								rsComboP.CursorLocation = 3
								rsComboP.Open q, connstr
								if not rsComboP.eof then%>
								<%
								Do While Not rsComboP.eof
								%>
									<option value="<%=rsComboP("TRANSPORTERCONTRACTNO")%>" 
										<%if cstr(""&transportercontractno)=cstr(""&rsComboP("TRANSPORTERCONTRACTNO")) then %> selected 
										<%End If%>><%=rsComboP("TRANSPORTERCONTRACTNO")%> 
									</option> 
								<%rsComboP.movenext
								loop
								End if
								rsComboP.Close
								set rsComboP = Nothing
								q=""
							%>
					</select>
					
					
				<%end if%>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Unit Of Measure
				<td align="left">
				<input type="text" maxlength="15" name="unitofmeasure" style="background:#eee;" size="30" value="<%=unitofmeasure%>" autocomplete="Off" readonly="true">
				</td>
			</td>
			<td align="left" width="20%">Line No
				<td align="left">
				<%IF txtJenis="Edit" Then%>
					<input type="text" size="5" name="transporterline" value="<%=transporterline%>" readonly style="background:#eee">
					<input type="text" size="10" name="shippingaddress" value="<%=shippingaddress%>" readonly style="background:#eee">
				<%else%>
				<div id="transporterlinelist"></div>
				<input type="Hidden" id="selectedtransporterline" name="selectedtransporterline" value="<%=selectedtransporterline%>">
				<%end if%>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%"> Start Date
				<td align="left">
				<input type="text" class="datepicker"  size="15" name="startdate" value="<%=startdate%>" >
				</td>
			</td>

			<td align="left" width="20%">Transporter 
				<td align="left">
				<input type="text" maxlength="15" id="transporter" name="transporter" style="background:#eee;" size="5" value="<%=transporter%>" autocomplete="Off" readonly="true" >
				<input type="text" maxlength="15" id="transportername" name="transportername" style="background:#eee;" size="18" value="<%=transportername%>" autocomplete="Off" readonly="true" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">End Date 
				<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="enddate" size="15" value="<%=enddate%>">
				</td>
			</td>
		
			<td align="left" width="20%">Transportation Cost 
				<td align="left">
				<input type="text" maxlength="15" id="rate" name="transportationcost" size="30" style="background:#eee;" value="<%=transportationcost%>" autocomplete="Off" readonly="true">
				</td>
			</td>
		</tr>
		<tr>
			<td></td>
			<td></td>
			<td align="left" width="20%">Loss Tolerance 
				<td align="left">
				<input type="text" maxlength="15" style="background:#eee;" id="losstolerance" name="losstolerance" size="30" value="<%=losstolerance%>" autocomplete="Off" readonly="true" >
				<font color="#FF0000">%</font>
				</td>
			</td>
		</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="salescontractno" value="<%=salescontractno%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
						<%IF Request("txtJenis")="Edit" then%>
							<input type="hidden" name="txtJenis" value="Edit">
							<input type="hidden" name="lineno" value="<%=lineno%>">
							<input Type="Submit"  name="Add" value="Update" id="hitApi">
						<%Else%>
							<input type="hidden" name="txtJenis" value="New">
							<input Type="Submit"  name="Add" value="Add" id="hitApi">
						<%End if%>
						<input Type="Reset"  name="Reset" value="Reset">
						<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
					
				</td>
			</tr>
		</table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="2" width="30">&nbsp;</td>
		<td align="left">DO No</td>
		<td align="left">PRODUCT</td>
		<td align="left">QUANTITY</td>
		<td align="left">QTY DELIVERY ORDER</td>
		<td align="left">UOM</td>
		<td align="left">DELIVERY DATE</td>
		<td align="left">DELIVERY CONDITION</td>
		<td align="left">REMARKS</td>
		<td align="left">START DATE</td>
		<td align="left">END DATE</td>
		<td align="left">TRANSPORTER</td>
		<td align="left">TRANSPORTER LINE</td>
		<td align="left">TRANSPORTATION COST</td>
		<td align="left">LOSS TOLERANCE</td>
	</tr>
	</thead>
	<tbody>
   
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
			
			
			<td valign="middle">
				<a href="deliveryorder.asp?lineno=<%=rsVH("lineno")%>&transporterline=<%=rsVH("transporterline")%>&startdate=<%=rsVH("startdate")%>&enddate=<%=rsVH("enddate")%>&transportercontractno=<%=rsVH("transportercontractno")%>&salescontractno=<%=salescontractno%>&deliveryorderno=<%=rsVH("deliveryorderno")%>&deliveryorderdate=<%=rsVH("deliveryorderdate")%>&remarks=<%=rsVH("remarks")%>&quantity_do=<%=rsVH("quantity_do")%>&quantity=<%=rsVH("quantity")%>&deliverycondition=<%=rsVH("deliverycondition")%>&productcode=<%=rsVH("productcode")%>&startdate=<%=rsVH("startdate")%>&enddate=<%=rsVH("enddate")%>&unitofmeasure=<%=rsVH("unitofmeasure")%>&transporter=<%=rsVH("transporter")%>&transportationcost=<%=rsVH("transportationcost")%>&losstolerance=<%=rsVH("losstolerance")%>&fcba=<%=fcba%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
			</td>
			<td valign="middle">
				<a href="deliveryorder.asp?lineno=<%=rsVH("lineno")%>&transportercontractno=<%=rsVH("transportercontractno")%>&startdate=<%=rsVH("startdate")%>&enddate=<%=rsVH("enddate")%>&salescontractno=<%=salescontractno%>&deliveryorderno=<%=rsVH("deliveryorderno")%>&deliveryorderdate=<%=rsVH("deliveryorderdate")%>&remarks=<%=rsVH("remarks")%>&quantity_do=<%=rsVH("quantity_do")%>&quantity=<%=rsVH("quantity")%>&deliverycondition=<%=rsVH("deliverycondition")%>&productcode=<%=rsVH("productcode")%>&startdate=<%=rsVH("startdate")%>&enddate=<%=rsVH("enddate")%>&unitofmeasure=<%=rsVH("unitofmeasure")%>&transporter=<%=rsVH("transporter")%>&transportationcost=<%=rsVH("transportationcost")%>&losstolerance=<%=rsVH("losstolerance")%>&fcba=<%=fcba%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0"></a>
			</td>
			    <td align="left" title="<%=rsVH("deliveryorderno").Name%>" valign="top"><%=rsVH("deliveryorderno")%></td>
			    <td align="left" title="<%=rsVH("productcode").Name%>" valign="top"><%=rsVH("productcode")%></td>
			    <td align="left" title="<%=rsVH("Quantity").Name%>" valign="top"><%=rsVH("Quantity")%></td>
			    <td align="left" title="<%=rsVH("quantity_do").Name%>" valign="top"><%=rsVH("quantity_do")%></td>
			    <td align="left" title="<%=rsVH("unitofmeasure").Name%>" valign="top"><%=rsVH("unitofmeasure")%></td>
			    <td align="left" title="<%=rsVH("deliveryorderdate").Name%>" valign="top"><%=rsVH("deliveryorderdate")%></td>
			    <td align="left" title="<%=rsVH("deliverycondition").Name%>" valign="top"><%=rsVH("deliverycondition")%></td>
			    <td align="left" title="<%=rsVH("remarks").Name%>" valign="top"><%=rsVH("remarks")%></td>
			    <td align="left" title="<%=rsVH("startdate").Name%>" valign="top"><%=rsVH("startdate")%></td>
			    <td align="left" title="<%=rsVH("enddate").Name%>" valign="top"><%=rsVH("enddate")%></td>
			    <td align="left" title="<%=rsVH("transporter").Name%>" valign="top"><%=rsVH("transporter")%>: <%=rsVH("transportername")%></td>
			    <td align="left" title="<%=rsVH("transporterline").Name%>" valign="top"><%=rsVH("transporterline")%></td>
			    <td align="left" title="<%=rsVH("transportationcost").Name%>" valign="top"><%=rsVH("transportationcost")%></td>
			    <td align="left" title="<%=rsVH("losstolerance").Name%>" valign="top"><%=rsVH("losstolerance")%></td>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </tbody>
   </TABLE>
   </div>
<% if Request("oldui") <> "1" then %>
<script type="text/javascript">
// Status baris jadi badge (tampilan saja)
(function () {
	var cells = document.querySelectorAll('table.fbp-doc-lines tbody td[title]');
	for (var i = 0; i < cells.length; i++) {
		if (cells[i].getAttribute('title').toUpperCase() !== 'ROWSTATEID') continue;
		var s = (cells[i].textContent || '').replace(/\s+/g, ' ').trim();
		if (!s) continue;
		cells[i].parentNode.setAttribute('data-state', s);
		cells[i].innerHTML = '';
		var b = document.createElement('span');
		b.className = 'fbp-state';
		b.textContent = s;
		cells[i].appendChild(b);
	}
})();
</script>
<% end if %>
   </form>
<%End if%>
 
  <!-------------akhir!------------->
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

	function formatDate(tgl) {
		let parts = tgl.split("/"); // MM/DD/YYYY
		return parts[2] + "-" + parts[0].padStart(2, '0') + "-" + parts[1].padStart(2, '0');
	}
	
	$("#hitApi").on("click", function(e){
		e.preventDefault(); // ← WAJIB, cegah form submit dulu

		const productCode = document.querySelector('input[name="productcode"]').value;
		const deliveryorderno = document.querySelector('input[name="deliveryorderno"]').value;
		const deliveryorderdate = document.querySelector('input[name="deliveryorderdate"]').value;
		const startdate = document.querySelector('input[name="startdate"]').value;
		const enddate = document.querySelector('input[name="enddate"]').value;
		const qty = document.querySelector('input[name="Quantity"]').value;
		const quantity_do = document.querySelector('input[name="quantity_do"]').value;
		const salescontractno = document.querySelector('input[name="salescontractno"]').value;
		const deliverycondition = document.querySelector('textarea[name="deliverycondition"]').value; // ← textarea bukan input
		const remarks = document.querySelector('textarea[name="remarks"]').value;                     // ← textarea bukan input
		const username = document.querySelector('input[name="username"]').value;
		const fcbasession = document.querySelector('input[name="fcbasession"]').value;
		const customerid = document.querySelector('input[name="customerid"]').value;

		let dariTgl = formatDate(startdate);
		let keTgl = formatDate(enddate);

		var settings = {
			"url": "http://117.53.45.125:8080/api/deliveryorder",
			"method": "POST",
			"timeout": 0,
			"headers": {
				"Content-Type": "application/json"
			},
			"data": JSON.stringify({
				"productid": productCode,
				"doorderid": deliveryorderno,
				"dofromdate": dariTgl,
				"dotodate": keTgl,
				"contractqty": parseFloat(qty) || 0,           // ← parseFloat
				// "usedqty": parseFloat(quantity_do) || 0,
				"usedqty": parseFloat(0),
				"priceperkg": 0,
				"currencyid": "IDR",
				"customerid": customerid,
				"contractid": salescontractno,
				"contractorid": $("#transporter").val(),
				"lastspbno": null,
				"transportprice": 0,
				"doflg": 0,
				"invoiceflg": 0,
				"exportflg": 0,
				"note": deliverycondition + " | " + remarks,
				"userid": username,
				"fcba": fcbasession
			}),
		};

		$.ajax(settings)
		.done(function (response) {
			console.log(response);
			if (response.status === 200) {
				document.eForm.submit(); // ← submit form setelah API sukses
			} else {
				alert("Gagal: " + response.message);
			}
		})
		.fail(function (err) {
			console.error(err);
			alert("Gagal hit API: " + err.responseText);
		});
	});

</script>
<script>

function getselectedfcba() {
    var sfcba = document.getElementById("selectedfcba");
    var fcbaEl = document.getElementById("fcba");
    // Tambah pengecekan null
    if (sfcba && fcbaEl) {
        sfcba.value = fcbaEl.value;
    }
}

function getselectedtransporterline() {
    var slineno = document.getElementById("selectedtransporterline");
    var linenoEl = document.getElementById("lineno");
    // Tambah pengecekan null
    if (slineno && linenoEl) {
        slineno.value = linenoEl.value;
    }
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
$(document).ready(function() {
    getlistfcba();
    
    // Hanya jalankan jika element ada (mode New, bukan Edit)
    if (document.getElementById("transportercontractno")) {
        getTransporterLine();
    }
    
    if (document.getElementById("transporter")) {
        getTransporter();
    }
    
    getselectedfcba();
    getselectedtransporterline();
});
</script>
</body>
</html>