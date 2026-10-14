<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - SPB Harvesting</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script language="javascript" type="text/javascript" src="ajax/validate_ajax.js"></script>
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
    var e = evt || window.event;
    var key = e.keyCode || e.which;

    if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
    // numbers   
    key >= 48 && key <= 57 ||
	e.ctrlKey && key == 97 ||
    // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
    // left and right arrows
    key == 37 || key == 39 ||
    // Del and Ins
	key == 46 ) {
        // input is VALID
		}
	else {
			// input is INVALID
			e.returnValue = false;
			if (e.preventDefault) e.preventDefault();
		}
}

// Untuk Disable Enable Field Text Based On Select Option
// $(document).ready(function() {
//  	$('#own').click(function() {
//   	if( $(this).val() === 'OWN') {
//   //  		$('#agreementcode').css("background-color", "#eee");
// 		// $('#agreementcode').prop('readOnly', true);
// 		$( "#agreementcode" ).prop( "disabled", true );
// 		$( "#vehicle" ).prop( "disabled", false );
// 		$( "#driver" ).prop( "disabled", false );
//     } else {      

//     	$( "#agreementcode" ).prop( "disabled", false );
//     	$( "#vehicle" ).prop( "disabled", true );
// 		$( "#driver" ).prop( "disabled", true );
//     }
//   });
 
// });
</script>

<script>
	function hitungrumus() {
		var spressemester_abw = $("#pressemester_abw").val();
		var sbunch = $("#bunch").val();
		// var newssalesprices = ssalesprices.replace(/,/g, '');
		
		//Untuk menghitung Total Estate Weight
		totalestateweight = sbunch * spressemester_abw ; //a kali b
		// alert(totalestateweight);
		if (!isNaN(totalestateweight)) {
    	 	$("#bunch_estateweight").val(totalestateweight);
    	}
    	else{
    		$("#bunch_estateweight").val();
    	}
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
	var GetEmpName = function () {
		// build up the post string when passing variables to the server side page
		// var suppliercode = document.getElementById("suppliercode").value;
		var driver = document.getElementById("driver").value;
		var PostStr = 'validate=empcd&q='+driver;
		//alert(PostStr);
		// use the generic function to make the request
		doAJAXCall('ajax/validate.asp', 'POST', PostStr, showEmpnameResponse);
		
	}

	// The function for handling the response from the server
	var showEmpnameResponse = function (oXML) { 
	    
	    // get the response text, into a variable
	    var response = oXML.responseText;
	    // alert(s.supplieraddress);
	    // update the Div to show the result from the server
	    // document.getElementById("suppliername").value = s.suppliername;
		// document.getElementById("supplieraddress").value = s.supplieraddress;
		document.getElementById("nonscript_empcd").innerHTML = response;
		// alert(s.suppliername);
	    
	}

// launched from button click 
var getBJR = function () {
	// build up the post string when passing variables to the server side page
	var fieldcode = document.getElementById("fieldcode").value;
	var receptiondate = document.getElementById("receptiondate").value;
	var PostStr = 'fieldcode='+fieldcode+'&receptiondate='+receptiondate;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetBJR.asp', 'POST', PostStr, showMessageResponse);
	
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.harvestingbased_abw);
    // update the Div to show the result from the server
    // document.getElementById("suppliername").value = s.suppliername;
	document.getElementById("pressemester_abw").value = s.pressemester_abw;
	// document.getElementById("address").innerHTML = document.getElementById("supplieraddress").value;
	// alert(s.suppliername);
    
};

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
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
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
 Dim cmdView
CHITNO=request("CHITNO")
CROPCODE=request("CROPCODE")
PRODUCTCODE=request("ProductCODE")
receptiondate=request("receptiondate")
VEHICLE=request("VEHICLE")
OWN=request("OWN")
TRANSPORTTYPE = Request("TRANSPORTTYPE")
CONTRACTNUMBER = Request("CONTRACTNUMBER")
QUANTITY=request("QUANTITY")
MILL_WEIGHT_BRUTO=request("MILL_WEIGHT_BRUTO")
MILL_WEIGHT=request("MILL_WEIGHT")
SPB_NO = request ("SPB_NO")
SPB_TYPE = request ("SPB_TYPE")
AGREEMENTCODE = request ("AGREEMENTCODE")
DRIVER = request ("DRIVER")
MILL = request ("MILL")
lineno = request("lineno")

' Detail
QUANTITY_DTL=request("QUANTITY_DTL")
FIELDCODE=request("FIELDCODE") 
pressemester_abw=request("pressemester_abw")
bucket=request("bucket")
bunch_estateweight=request("bunch_estateweight")
bunch=request("bunch")
harvestdate = request("harvestdate")
total_bunch=request("total_bunch")
total_bucket = request("total_bucket")
total_estateweight = request ("total_estateweight")
brondolan = request ("brondolan")

if len(trim(""&spb_type)) = 0 then spb_type = 0
If Len(Trim(""&receptiondate)) = 0 Then	receptiondate = textDate(date)
If Len(Trim(harvestdate)) = 0 Then	harvestdate = textDate(date)

Function CheckPrimary(initialFcba, initialDocument)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	" select count(1) total" & _
	                " from supplyingdeclarationhdr " & _
					" where documentno = '"& initialDocument &"'" & _
					" and fcba = '" & initialFcba & "'"

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

	CheckPrimary = initialAvailable
End Function

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

Function CheckVehicleCode(initVehicleCode, initFcba)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select VEHICLE_API.GET_VEHICLECODE_NAME('"& initVehicleCode &"','"& initFcba &"') total from dual" & _
				" where VEHICLE_API.GET_ACTIVATION('"& initVehicleCode &"','"& initFcba &"') = 'Y'"
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(to_nol(rsAvail("Total")))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	CheckVehicleCode = initialAvail
	'Select count(1) from dual where field_api.GetValid_Activation('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')) = 'Y'
	'select to_char (field_api.GetValid_Plantingdate('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')),'yyyy') from dual
End Function

Function F_CurrentTotalStand(initFieldCode, initFcba, initDate)
	Dim initialAvail
	Dim rsAvail, strAvail
	
	strAvail =	" select field_api.getValid_POKOK('" & initFieldCode & "', '" & initFcba & "', to_date('"& initDate &"', 'mm/dd/yyyy')) total" & _
	            " from dual " 
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		initialAvail = Cdbl(rsAvail("Total"))
	Else
		initialAvail = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	F_CurrentTotalStand = initialAvail
End Function

Function F_round (initNumber)
	
	result = 0
	strAvail =	" select round ('" & initNumber & "',0) total" & _
	            " from dual " 
	' response.write stravail
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	rsAvail.Open strAvail, connstr, 3, 3
	
	If Not rsAvail.EOF Then
		result = Cdbl(to_nol(rsAvail("Total")))
	Else
		result = 0
	End If
	strAvail = ""
	rsAvail.Close
	set rsAvail = Nothing

	F_round = result
End Function

Function CheckQtyPartial(initfieldcode, initialFcba, inityear)
	Dim initialAvailable
	Dim rsAvailable, strAvailable

	strAvailable =	" SELECT sum(sdd.quantity) total " & _
		  " FROM supplyingdeclarationdtl sdd" & _
		  " INNER JOIN supplyingdeclarationhdr sdh"& _
		  " ON (sdd.documentno = sdh.documentno AND sdd.fcba = sdh.fcba)" & _
		  " WHERE sdd.fcba = '" & initialfcba & "'" & _
		  " AND sdd.fieldcode = '" & initfieldcode & "'" & _
		  " AND to_char(sdh.transdate,'yyyy') = '" & inityear & "'"

	 ' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(to_nol(rsAvailable("Total")))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckQtyPartial = initialAvailable
End Function

Private Sub ProcessFormulaAndInsertHarvesting
	' strFieldDetail = " SELECT COUNT(*) total FROM harvestingdtl hd" & _
	' 				 " INNER JOIN harvestinghdr hh ON(hh.chitno = hd.chitno AND hh.fcba = hd.fcba)" & _
 '                     " where hd.fcBA = '" & Session("fcBA") & "'" & _
	' 				 " and hd.chitno = '"& chitno &"' and hd.fieldcode = '"& fieldcode &"'"							
	' Response.Write strFieldDetail
	' Set rsQty = Server.CreateObject("ADODB.Recordset")
	' rsQty.Open strFieldDetail, connstr							
	' if not rsQty.Eof then
	' 	while not rsQty.Eof
	' 		fieldcodetotal = cdbl(to_nol(rsQty("total")))
	' 		rsQty.movenext
	' 	wend
	' End If
	
	' If cdbl(to_nol(fieldcodetotal)) > 1 Then
	' 		err.number = 1001
	' 		Err.Description = "Anda Sudah Memasukan Blok Ini, Silahkan Pilih Blok Yang Lain !"
	' Else
	z = " DELETE from harvesting " & _
		" WHERE chitno = '" & CHITNO & "'" & _
		"		AND fcBA = '" & Session("fcBA") & "' AND fieldcode = '"& fieldcode &"'"
	' Response.Write "<br>"&z
	connStr.execute(z)

	'Untuk menghhitung jumlah total
	estate_weight_total = 0
	' estate_weight_dtl = 0
	
	q="select *  from harvestingdtl " & _
	  " where chitno ='"& CHITNO &"'" & _
	  " and fcBA='" & session("fcba") & "'"
	  ' response.write q
	set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
	rsTempPurchaseRequest.Open q, connstr
	if Not rsTempPurchaseRequest.Eof then
		while Not rsTempPurchaseRequest.Eof 
			estate_weight_dtl =  F_round( ( cdbl(to_nol(rsTempPurchaseRequest("quantity_dtl"))) *  cdbl(to_nol(rsTempPurchaseRequest("pressemester_abw"))) ) + cdbl(to_nol(rsTempPurchaseRequest("brondolan"))) )
			estate_weight_total = estate_weight_total + estate_weight_dtl
			
			rsTempPurchaseRequest.movenext
		wend 
	End if
	rsTempPurchaseRequest.close : set rsTempPurchaseRequest = nothing

	'response.write q
	set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
	rsTempPurchaseRequest.Open q, connstr
	if Not rsTempPurchaseRequest.Eof then
		while Not rsTempPurchaseRequest.Eof
			
			estate_weight_dtl =  F_round( ( cdbl(to_nol(rsTempPurchaseRequest("quantity_dtl"))) *  cdbl(to_nol(rsTempPurchaseRequest("pressemester_abw"))) ) + cdbl(to_nol(rsTempPurchaseRequest("brondolan"))) )
		
			if Cdbl(to_nol(estate_weight_total)) = 0 then mill_weight_bruto_dtl = 0 else mill_weight_bruto_dtl = F_round(estate_weight_dtl / Estate_weight_total * Cdbl(to_nol(mill_weight_bruto)))
			if Cdbl(to_nol(estate_weight_total)) = 0 then mill_weight_dtl = 0 else mill_weight_dtl = F_round(estate_weight_dtl / Estate_weight_total * Cdbl(to_nol(mill_weight)))
			
			if cstr(""&spb_type) = "1" then
				if Cdbl(to_nol(estate_weight_total)) = 0 then estate_weight_dtl = 0 else estate_weight_dtl = F_round( estate_weight_dtl / Estate_weight_total * cdbl(to_nol(ESTATE_WEIGHT)) )
			else
				if Cdbl(to_nol(estate_weight_total)) = 0 then estate_weight_dtl = 0 else estate_weight_dtl = F_round(estate_weight_dtl / Estate_weight_total * Cdbl(to_nol(estate_weight_total)))
			end if
		
		  z= " Update harvestingdtl set mill_weight_bruto_dtl = round(replace('" & mill_weight_bruto_dtl & "',',',''),0)" & _
			 ", estate_weight_dtl = round(replace('" & estate_weight_dtl & "',',',''),0)" & _
			 ", mill_weight_dtl = round(replace('" & mill_weight_dtl & "',',',''),0)" & _
			 " where chitno ='"& CHITNO &"'" & _
			 " and fcBA='" & session("fcba") & "' and lineno = '"& rsTempPurchaseRequest("lineno") &"'"
			 ' response.write z
			connStr.execute(z)

			z= " Update harvesting set mill_weight_bruto = round(replace('" & mill_weight_bruto_dtl & "',',',''),0)" & _
			 ", estate_weight = round(replace('" & estate_weight_dtl & "',',',''),0)" & _
			 ", mill_weight = round(replace('" & mill_weight_dtl & "',',',''),0)" & _
			 " where chitno ='"& CHITNO &"'" & _
			 " and fcBA='" & session("fcba") & "' and lineno = '"& rsTempPurchaseRequest("lineno") &"'"
			 ' response.write z
			connStr.execute(z)
			rsTempPurchaseRequest.movenext
		wend 
		
	End if

	rsTempPurchaseRequest.Close : set rsTempPurchaseRequest = Nothing
		
	z = "INSERT INTO harvesting (CHITNO," & _
		"                        FIELDCODE," & _
		"                        CROPCODE," & _
		"                        PRODUCTCODE," & _
		"                        RECEPTIONDATE," & _
		"                        VEHICLE," & _
		"                        OWN," & _
		"                        QUANTITY," & _
		"                        ESTATE_WEIGHT," & _
		"                        MILL_WEIGHT," & _
		"                        FCENTRY," & _
		"                        FCEDIT," & _
		"                        FCIP," & _
		"                        FCBA," & _
		"                        LASTUPDATE," & _
		"                        LASTTIME," & _
		"                        HARVESTINGTRANPORTTYPECODE," & _
		"                        AGREEMENTCODE," & _
		"                        CONTRACTORCODE," & _
		"                        LINENO," & _
		"                        MONTH," & _
		"                        YEAR," & _
		"                        MILL_WEIGHT_BRUTO)" & _
		"   SELECT hh.chitno CHITNO," & _
		"          hd.fieldcode FIELDCODE," & _
		"          hh.cropcode CROPCODE," & _
		"          hh.productcode PRODUCTCODE," & _
		"          hh.receptiondate RECEPTIONDATE," & _
		"          hh.vehicle VEHICLE," & _
		"          hh.own OWN," & _
		"          hd.quantity_dtl QUANTITY," & _
		"          hd.ESTATE_WEIGHT_DTL ESTATE_WEIGHT," & _
		"          hd.MILL_WEIGHT_DTL MILL_WEIGHT," & _
		"          hh.fcentry FCENTRY," & _
		"          hh.fcedit FCEDIT," & _
		"          hh.fcip FCIP," & _
		"          hh.fcba FCBA," & _
		"          hh.lastupdate LASTUPDATE," & _
		"          hh.lasttime LASTTIME," & _
		"          HARVESTINGTRANPORTTYPECODE," & _
		"          AGREEMENTCODE," & _
		"          NULL CONTRACTORCODE," & _
		"          LINENO," & _
		"          hh.month MONTH," & _
		"          hh.year YEAR," & _
		"          hd.MILL_WEIGHT_BRUTO_DTL MILL_WEIGHT_BRUTO" & _
		"     FROM harvestingdtl hd" & _
		"          INNER JOIN harvestinghdr hh" & _
		"             ON (hh.chitno = hd.chitno AND hh.fcba = hd.fcba)" & _
		"    WHERE hh.chitno = '"& chitno &"' AND hh.fcba = '"& session("fcba") &"' AND hd.fieldcode ='"& fieldcode &"'" 

	' Response.Write "<br>"&z
	connStr.execute(z)
	'end add
	
	' Untuk menghitung Selisih Mill Weight
	qLineMax =  "  SELECT MAX (lineno) lineno" & _
	            "         , hh.mill_weight - SUM (mill_weight_dtl) selisihnetto" & _
	            "         , hh.mill_weight_bruto - SUM (mill_weight_bruto_dtl) selisihbruto " & _
	            "         , CASE WHEN hh.spb_type = '1' THEN hh.estate_weight - SUM (estate_weight_dtl) ELSE SUM (estate_weight_dtl) - SUM (estate_weight_dtl) END selisih_ew " & _
				"  FROM harvestingdtl hd INNER JOIN harvestinghdr hh ON (hh.chitno = hd.chitno and hh.fcba = hd.fcba)" & _
				"  where hh.chitno = '" & CHITNO & "'" & _
				"  and hh.fcba = '" & session("fcBA") & "'" & _		
				"  GROUP BY hh.mill_weight, hh.mill_weight_bruto, hh.spb_type, hh.estate_weight" 
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		lineno = cdbl(to_nol(rsqLineMax("lineno")))
		selisihnetto = cdbl(to_nol(rsqLineMax("selisihnetto")))
		selisihbruto = cdbl(to_nol(rsqLineMax("selisihbruto")))
		selisih_ew = cdbl(to_nol(rsqLineMax("selisih_ew")))
	end if 
	
	If selisihnetto <> 0 Then
		 z=  " Update harvestingdtl set mill_weight_dtl = mill_weight_dtl + '"& selisihnetto &"'" & _
			 " where chitno ='"& CHITNO &"'" & _
			 " and fcBA='" & session("fcba") & "' and lineno = '"& lineno &"'"
			 ' response.write z
         connStr.execute(z)
         z=  " Update harvesting set mill_weight = mill_weight_dtl + '"& selisihnetto &"'" & _
			 " where chitno ='"& CHITNO &"'" & _
			 " and fcBA='" & session("fcba") & "' and lineno = '"& lineno &"'"
			 ' response.write z
         connStr.execute(z)
	End If
	If selisihbruto <> 0 Then
		 z=  " Update harvestingdtl set mill_weight_bruto_dtl = mill_weight_bruto_dtl + '"& selisihbruto &"'" & _
			 " where chitno ='"& CHITNO &"'" & _
			 " and fcBA='" & session("fcba") & "' and lineno = '"& lineno &"'"
			 ' response.write z
         connStr.execute(z)
         z=  " Update harvesting set mill_weight_bruto = mill_weight_bruto_dtl + '"& selisihbruto &"'" & _
			 " where chitno ='"& CHITNO &"'" & _
			 " and fcBA='" & session("fcba") & "' and lineno = '"& lineno &"'"
			 ' response.write z
         connStr.execute(z)
	End If
	If selisih_ew <> 0 Then
		 z=  " Update harvestingdtl set estate_weight_dtl = estate_weight_dtl + '"& selisih_ew &"'" & _
			 " where chitno ='"& CHITNO &"'" & _
			 " and fcBA='" & session("fcba") & "' and lineno = '"& lineno &"'"
			 ' response.write z
		 connStr.execute(z)
		 z=  " Update harvesting set estate_weight = estate_weight_dtl + '"& selisih_ew &"'" & _
			 " where chitno ='"& CHITNO &"'" & _
			 " and fcBA='" & session("fcba") & "' and lineno = '"& lineno &"'"
			 ' response.write z
		 connStr.execute(z)
	End If
' End If
			
End Sub


'Insert/Update/Delete
If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			q = " DELETE from harvestingspb " & _
				" WHERE spbno = '" & spb_no & "'" & _
				"       AND fieldcode = '" & fieldcode & "'" & _
				"       AND fcBA = '" & Session("fcBA") & "'"
			' Response.Write q
			connStr.execute(q)
			
			cmdView	= "True"
			
	elseif Request("txtJenis")="New" then
			' Added By Sendy 03/20/2019
			' Validasi untuk Blok yang di Input tidak boleh sama dalam 1 chitno
			strFieldDetail = " SELECT COUNT(*) total FROM harvestingspb hd" & _
							 " where hd.fcBA = '" & Session("fcBA") & "'" & _
							 " and hd.spbno = '"& spb_no &"' and hd.fieldcode = '"& fieldcode &"'" & _
							 " and hd.harvestdate = to_date('"& harvestdate &"','mm/dd/yyyy')"
			' Response.Write strFieldDetail
			Set rsQty = Server.CreateObject("ADODB.Recordset")
			rsQty.Open strFieldDetail, connstr							
			if not rsQty.Eof then
				while not rsQty.Eof
					fieldcodetotal = cdbl(to_nol(rsQty("total")))
					rsQty.movenext
				wend
			End If
			
			If cdbl(to_nol(fieldcodetotal)) > 0 Then
				err.number = 1001
				Err.Description = "Anda Sudah Memasukan Blok """ & fieldcode & """, Silahkan Pilih Blok Yang Lain !"
			Else	
			' If cstr(""&own) = "OWN" Then
			' 	val_VehicleCode = checkvehiclecode(vehicle, session("fcBA"))
			' 	If cdbl(to_nol(val_VehicleCode)) = 0 Then
			' 		Err.Number = 1001
			' 		Err.Description = "Vehicle Code Harus diisi, tidak boleh kosong"
			' 	Else
			' 		If cdbl(to_nol(checkemployee)) = 0 Then
			' 			Err.Number = 1001
			' 			Err.Description = "Employee Code Harus diisi, tidak boleh kosong"
			' 		End If
			' 	End If
			' Else
			' If cstr(""&own) = "CONTRACT" Then
			' 	If cdbl(to_nol(agreementcode)) = 0 Then
			' 		Err.Number = 1001
			' 		Err.Description = "Contract Code Harus diisi, tidak boleh kosong"
			' 	End If
			' Else
			' strFieldDetail = " SELECT COUNT(*) total FROM employee " & _
			' 				 " where fcBA = '" & Session("fcBA") & "'" & _
			' 				 " and dateterminate is null and fccode = '"& driver &"'"							
			' ' Response.Write strFieldDetail
			' Set rsQty = Server.CreateObject("ADODB.Recordset")
			' rsQty.Open strFieldDetail, connstr							
			' if not rsQty.Eof then
			' 	while not rsQty.Eof
			' 		checkemployee = cdbl(to_nol(rsQty("total")))
			' 		rsQty.movenext
			' 	wend
			' End If
			' If cdbl(to_nol(checkemployee)) = 0 Then
			' 	Err.Number = 1001
			' 	Err.Description = "Employee Code """ & driver & """ tidak ada di master Employee iPlaS atau Employee Code """ & driver & """ Sudah Tidak Aktif."
			' Else

			If Len(Trim(fieldcode)) = 0 Then
				Err.Number = 1001
				Err.Description = "Fieldcode Harus diisi, tidak boleh kosong"
			Else
			' If cdbl(to_nol(bunch)) = 0 Then
				' Err.Number = 1001
				' Err.Description = "Bunch Harus diisi, tidak boleh kosong"
			' Else
			Dim val_SupplyingCode : val_SupplyingCode = 0
			val_SupplyingCode = CheckFieldCode(fieldcode, session("fcBA"), receptiondate)
			If val_SupplyingCode = 0 Then
				Err.Number = 1001
				Err.Description = "Field Code """ & fieldcode & """ tidak ada di master Field iPlaS, Koreksi kembali Kode Blok yang anda masukkan."
			Else
			' Dim val_VehicleCode : val_VehicleCode = 0
			' val_VehicleCode = checkvehiclecode(vehicle, session("fcBA"))


			If request("own") = "OWN" Then 

				strFieldDetail =	" SELECT count(1) total" & _
									"   FROM employee" & _
									"  WHERE fcba = '"& session("fcba") &"' AND fccode = '"& driver &"'"  
						
				' Response.Write strFieldDetail
				Set rsQty = Server.CreateObject("ADODB.Recordset")
				rsQty.Open strFieldDetail, connstr							
				if not rsQty.Eof then
					while not rsQty.Eof
						checkemployee = cdbl(to_nol(rsQty("total")))
						rsQty.movenext
					wend
				End If
				
				strFieldDetail =	" SELECT count(1) total" & _
									"   FROM vehicle" & _
									"  WHERE fcba = '"& session("fcba") &"' AND fccode = '"& vehicle &"' AND activation ='Y'"  
						
				' Response.Write strFieldDetail
				Set rsQty = Server.CreateObject("ADODB.Recordset")
				rsQty.Open strFieldDetail, connstr							
				if not rsQty.Eof then
					while not rsQty.Eof
						checkvehicle = cdbl(to_nol(rsQty("total")))
						rsQty.movenext
					wend
				End If
				
				If checkvehicle = 0 Then
					Err.Number = 1001
					Err.Description = "Vehicle Code """ & vehicle & """ tidak ada di master Vehicle iPlaS, Koreksi kembali Vehicle Code yang anda masukkan."
				Else 
					If checkemployee = 0 Then
						Err.Number = 1001
						Err.Description = "Employee Code """ & driver & """ tidak ada di master Employee iPlaS, Koreksi kembali Employee Code yang anda masukkan."
					Else
						q = " insert into harvestingspb" & _
							" (spbno, receptiondate, harvestdate, cropcode, fieldcode, productcode, own, vehicle, driver, mill, agreementcode, transporttype, spb_type, bunch, bucket, pressemester_abw, bunch_estateweight, brondolan, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
							" 	'" & spb_no & "'" & _
							" , TO_DATE('" & receptiondate & "', 'mm/dd/yyyy')" & _
							" , TO_DATE('" & harvestdate & "', 'mm/dd/yyyy')" & _
							" ,	'" & cropcode & "'" & _
							" ,	'" & fieldcode & "'" & _
							" ,	'" & productcode & "'" & _
							" ,	'" & own & "'" & _
							" ,	'" & vehicle & "'" & _
							" ,	'" & driver & "'" & _
							" ,	'" & mill & "'" & _
							" ,	'" & agreementcode & "'" & _
							" ,	'" & transporttype & "'" & _
							" ,	'" & spb_type & "'" & _
							" , '" & bunch & "'" & _
							" , '" & bucket & "'" & _
							" , replace('" & pressemester_abw & "',',','.')" & _
							" , replace('" & bunch_estateweight & "',',','')" & _
							" , replace('" & brondolan & "',',','')" & _
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
			Else
				q = " insert into harvestingspb" & _
					" (spbno, receptiondate, harvestdate, cropcode, fieldcode, productcode, own, vehicle, driver, mill, agreementcode, transporttype, spb_type, bunch, bucket, pressemester_abw, bunch_estateweight, brondolan, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" 	'" & spb_no & "'" & _
					" , TO_DATE('" & receptiondate & "', 'mm/dd/yyyy')" & _
					" , TO_DATE('" & harvestdate & "', 'mm/dd/yyyy')" & _
					" ,	'" & cropcode & "'" & _
					" ,	'" & fieldcode & "'" & _
					" ,	'" & productcode & "'" & _
					" ,	'" & own & "'" & _
					" ,	'" & vehicle & "'" & _ 
					" ,	'" & driver & "'" & _ 
					" ,	'" & mill & "'" & _
					" ,	'" & agreementcode & "'" & _
					" ,	'" & transporttype & "'" & _
					" ,	'" & spb_type & "'" & _
					" , '" & bunch & "'" & _
					" , '" & bucket & "'" & _
					" , replace('" & pressemester_abw & "',',','.')" & _
					" , replace('" & bunch_estateweight & "',',','')" & _
					" , replace('" & brondolan & "',',','')" & _
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
	End If
End If
cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		If Len(Trim(fieldcode)) = 0 Then
			Err.Number = 1001
			Err.Description = "Fieldcode Harus diisi, tidak boleh kosong"
		Else
			v_total = cdbl(to_nol(bunch)) + cdbl(to_nol(bucket))
			If cdbl(to_nol(v_total)) = 0 Then
				Err.Number = 1001
				Err.Description = "Bunch Atau Bucket Harus diisi, tidak boleh kosong"
			Else
				' Dim val_SupplyingCode : val_SupplyingCode = 0
				val_SupplyingCode = CheckFieldCode(fieldcode, session("fcBA"), receptiondate)
				If val_SupplyingCode = 0 Then
					Err.Number = 1001
					Err.Description = "Field Code """ & fieldcode & """ tidak ada di master Field iPlaS, Koreksi kembali Kode Blok yang anda masukkan."
				Else
					q =	" Update harvestingspb" & _
						" set fieldcode = '"& fieldcode &"'" & _
						"     , pressemester_abw = replace('" & pressemester_abw & "',',','.')" & _
						"     , receptiondate = TO_DATE('" & receptiondate & "', 'mm/dd/yyyy')" & _
						"     , harvestdate = TO_DATE('" & harvestdate & "', 'mm/dd/yyyy')" & _
						"     , transporttype = '" & transporttype & "'" & _
						"     , cropcode = '" & cropcode & "'" & _
						"     , productcode = '" & productcode & "'" & _
						"     , own = '" & own & "'" & _
						"     , agreementcode = '" & agreementcode & "'" & _
						"     , driver = '" & driver & "'" & _
						"     , mill = '" & mill & "'" & _
						"     , spb_type = '" & spb_type & "'" & _
						"     , bunch = '" & bunch & "'" & _
						"     , bucket = '" & bucket & "'" & _
						"	  ,	bunch_estateweight = replace('" & bunch_estateweight & "',',','')" & _
						"	  ,	brondolan = replace('" & brondolan & "',',','')" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" where fcBA='" & session("fcba") & "' and spbno = '"& spb_no &"' and fieldcode = '"& fieldcode &"'"
					'response.write q
					connStr.execute(q)
					' call ProcessFormulaAndInsertHarvesting
				End If
			End If
		End If
		
		cmdView	= "True"
	else
		cmdView	= "True"
 	End if
	Else
	 	cmdview = request("cmdview")
 	End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")


If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if


' Select Header
	
Dim rsSupplyingDeclaration, strSuppyingDeclaration

total_bunch = 0
total_bucket = 0 
total_estateweight = 0
strSuppyingDeclaration	=	" select * " & _
							" from harvestingspb" & _
							" where spbno = '"& spb_no &"'"& _
							" and fcba = '" & session("fcBA") & "'"
' response.write strSuppyingDeclaration	
set rsSupplyingDeclaration = server.CreateObject("Adodb.recordset")
rsSupplyingDeclaration.Open strSuppyingDeclaration, connstr
if Not rsSupplyingDeclaration.Eof then
	while not rsSupplyingDeclaration.Eof 
		spb_no				=	rsSupplyingDeclaration("spbno")
		receptiondate		=	rsSupplyingDeclaration("RECEPTIONDATE")
		cropcode			=	rsSupplyingDeclaration("cropcode")	
		productcode			=	rsSupplyingDeclaration("productcode")	
		own					=	rsSupplyingDeclaration("own")	
		vehicle				=	rsSupplyingDeclaration("vehicle")	
		agreementcode		=	rsSupplyingDeclaration("agreementcode")	
		driver				=	rsSupplyingDeclaration("driver")	
		mill				=	rsSupplyingDeclaration("mill")	
		transporttype		=	rsSupplyingDeclaration("transporttype")
		spb_type			=	rsSupplyingDeclaration("spb_type")	

		total_bunch = total_bunch + cdbl(to_nol(rsSupplyingDeclaration("bunch")))	
		total_bucket = total_bucket + cdbl(to_nol(rsSupplyingDeclaration("bucket")))
		total_estateweight = total_estateweight + cdbl(to_nol(rsSupplyingDeclaration("bunch_estateweight"))) 	
		rsSupplyingDeclaration.movenext
	wend
else
		total_bunch = 0
		total_bucket = 0
		total_estateweight = 0
End if
'response.Write strSuppyingDeclaration
rsSupplyingDeclaration.Close
Set rsSupplyingDeclaration = Nothing

' End Select

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & receptiondate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & receptiondate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('costbook')"
			'response.Write strNurseryCulling
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
 
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-truck" aria-hidden="true"></i>Header SPB Harvesting</div><% end if %>
	<form name="frmEntry" method="post" action="spbharvesting.asp">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">SPB No</td>
			<td align="left">
				<input type="text" maxlength="50" name="SPB_NO" size="30" value="<%=SPB_NO%>" autocomplete="Off">
				<font color="#FF0000"> *</font>
			</td>
			<%if cstr(""&OWN) = "CONTRACT" then%>
			<td align="left" width="20%"> Driver</td>
			<td align="left">
				<input id="driver" type="text" size="15" name="driver" value="<%=driver%>" onchange="GetEmpName(this.value);" onkeypress="return hanyaAngka(event);" >
				<span id="nonscript_empcd" style="color:black;">&nbsp;*&nbsp;
				<a id ="nonscriptempcd_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
				<font face="Verdana"   size="1">Employee</font></a></span>
			</td>
			<%else%>
			<td align="left" width="20%"> Driver</td>
			<td align="left">
				<input id="driver" type="text" size="15" name="driver" value="<%=driver%>" onchange="GetEmpName(this.value);" >
				<span id="nonscript_empcd" style="color:black;">&nbsp;*&nbsp;
				<a id ="nonscriptempcd_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
				<font face="Verdana"   size="1">Employee</font></a></span>
			<%end if%>
		</tr>
		<tr>
			<td align="left" width="20%">SPB Date</td>
			<td align="left"><input class="datepicker" type="text" size="15" name="receptiondate" id="receptiondate" value="<%=receptiondate%>"><font color="#FF0000"> *</font></td>
			<td align="left" width="20%"> Mill</td>
			<td align="left"><input type="text" size="30" name="MILL" value="<%=MILL%>"></td>
		</tr>
		<tr>
			<td align="left"><font face="Verdana" size="1">Transport Type</font></td>
			<td align="left">
				<select name="TRANSPORTTYPE">
				<%
					Dim rsTransportType, strTransportType
					strTransportType = " SELECT FCCODE, FCNAME FROM HARVESTINGTRANSPORTTYPE ORDER BY FCCODE"
					
					Set rsTransportType = Server.CreateObject("ADODB.Recordset")
					rsTransportType.Open strTransportType, connstr
					
					If Not rsTransportType.EOF then
						Do While Not rsTransportType.EOF
				%>
					<option value = "<%=rsTransportType("FCCODE")%>" <%If TRANSPORTTYPE = rsTransportType("FCCODE") Then%> selected <%End If%>>
						<%=rsTransportType("FCNAME")%>
					</option> 
				<%			rsTransportType.movenext
						Loop
					End if
					rsTransportType.Close
					set rsTransportType = Nothing
				%>
				</select><font color="#FF0000"> *</font>
			</td>
			<td align="left" width="20%">Weighbridge</td>
			<td align="left">
				<input type="checkbox" id="spb_type" name="spb_type" value="<%=spb_type%>" <%if cstr(""&spb_type) = "1" then%> checked <%end if%> onchange="handleClick(this);"> Available
			</td>
		</tr>
		<tr>
		
			<td align="left" width="20%">Crop Code</td>
			<td align="left">
				<select name="CROPCODE">
				<%
					Dim rsCrop, strCrop
					strCrop =	" SELECT P.FCCODE,P.FCNAME FROM  PARAMETER P  WHERE P.PARCODE= 'F0005' AND FCBA = '" &  Session("fcBA") & "'" & _
								" AND P.FCCODE = 'OP' ORDER BY P.FCNAME DESC"
					
					Set rsCrop = Server.CreateObject("ADODB.Recordset")
					rsCrop.Open strCrop, connstr
					
					If Not rsCrop.EOF then
						Do While Not rsCrop.EOF
				%>
					<option value="<%=rsCrop("FCCODE")%>" <%If CROPCODE = rsCrop("FCCODE") Then%> selected <%End If%>><%=rsCrop("FCNAME")%></option> 
				<%			rsCrop.movenext
						Loop
					End if
					rsCrop.Close
					set rsCrop = Nothing
				%>
				</select>
			</td>
			<td align="left" width="20%">Total Bunch</td>
			<td align="left">
				<input type="text" maxlength="15" name="total_bunch" size="15" value="<%=total_bunch%>" autocomplete="Off" readonly style= "background-color:#eee;">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Product Code</td>
			<td align="left">
				<select name="ProductCODE">
				<%
					Dim rsProductCode, strProductCode
					strProductCode = " SELECT PRODUCTCODE, PRODUCTDESCRIPTION FROM PRODUCT WHERE FCBA = '" &  Session("fcBA") & "' AND PRODUCTCODE = 'TBS' ORDER BY PRODUCTCODE"
					
					Set rsProductCode = Server.CreateObject("ADODB.Recordset")
					rsProductCode.Open strProductCode, connstr
					
					If Not rsProductCode.EOF then
						Do While Not rsProductCode.EOF
				%>
					<option value="<%=rsProductCode("PRODUCTCODE")%>" <%If PRODUCTCODE = rsProductCode("PRODUCTCODE") Then%> selected <%End If%>><%=rsProductCode("PRODUCTCODE")%> : <%=rsProductCode("PRODUCTDESCRIPTION")%></option> 
				<%			rsProductCode.movenext
						Loop
					End if
					rsProductCode.Close
					set rsProductCode = Nothing
				%>
				</select>
			</td>
			<td align="left" width="20%">Total Brondolan (Bucket)</td>
			<td align="left">
				<input type="text" name="total_bucket" size="15" value="<%=total_bucket%>" autocomplete="Off" readonly style= "background-color:#eee;">
			</td>
		</tr>
		<tr>
		
			<td align="left" width="20%">Vehicle Status </td>
			<td align="left">
				<select id="own" name="OWN" >
				<%
					Dim rsVehicleStatus, strVehicleStatus
					strVehicleStatus = " SELECT  FCNAME FROM HARVESTINGVEHICLESTATUS ORDER BY FCCODE"
					
					Set rsVehicleStatus = Server.CreateObject("ADODB.Recordset")
					rsVehicleStatus.Open strVehicleStatus, connstr
					
					If Not rsVehicleStatus.EOF then
						Do While Not rsVehicleStatus.EOF
				%>
					<option value = "<%=rsVehicleStatus("FCNAME")%>" <%If OWN = rsVehicleStatus("FCNAME") Then%> selected <%End If%>>
						<%=rsVehicleStatus("FCNAME")%>
					</option> 
				<%			rsVehicleStatus.movenext
						Loop
					End if
					rsVehicleStatus.Close
					set rsVehicleStatus = Nothing
				%>
				</select>
			</td>
			<td align="left" width="20%">Total Estate Weight</td>
			<td align="left">
				<input type="text" maxlength="15" name="total_estateweight" size="15" value="<%=total_estateweight%>" autocomplete="Off" onkeyup="hanyaAngka(this);" readonly style= "background-color:#eee;">
			</td>
		</tr>
		<tr> 
			<td align="left" width="20%">Vehicle Code</td>
			<td align="left"><input type="text" size="15" id="vehicle" name="VEHICLE" value="<%=VEHICLE%>" >
			 <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Vehicle&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
			<font face="Verdana" size="1">Vehicle Code</font>
			</td> 	
		</tr>
		<tr>
			<%if cstr(""&OWN) <> "CONTRACT" then%>	
				<td align="left" width="20%"> Contract Code</td>
				<td align="left"><input type="text" size="30" id="agreementcode" name="AGREEMENTCODE" value="<%=AGREEMENTCODE%>" ></td>
			<%else%>
				<td align="left" width="20%"> Contract Code</td>
				<td align="left"><input type="text" size="30" name="AGREEMENTCODE" value="<%=AGREEMENTCODE%>"></td>
			<%end if%>	
		</tr>
		
		<tr>
			<td colspan="4"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;
				<a href="spbharvestingupdate.asp?spb_no=<%=spb_no%>&productcode=<%=productcode%>&cropcode=<%=cropcode%>&driver=<%=driver%>&mill=<%=mill%>&agreementcode=<%=agreementcode%>&own=<%=own%>&vehicle=<%=vehicle%>&spb_type=<%=spb_type%>&receptiondate=<%=receptiondate%>" target="_blank"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("spb_no")%>" border="0"></a>
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
if cmdView = "True" and len(trim(spb_no)) > 0 and len(trim(vehicle)) > 0  Then 
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= 	" select * " & _
			" from harvestingspb" & _
			" where spbno = '"& spb_no &"'"& _
			" and fcba = '" & session("fcBA") & "'"

	'response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
	<form action="spbharvesting.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td align="left" width="20%">Field Code</td>
				<td align="left">
					<input type="text" size="15" name="FIELDCODE" id="fieldcode" value="<%=fieldcode%>" maxlength="32" autocomplete="off" onblur="javascript:getBJR();" <%if Request("txtJenis")="Edit" then %> style="background:#eee;" readonly <%else%> style="background:#FFFFCC;" <%end if%>>
					<font color="#FF0000">*</font>
				</td>
				
				<td align="left" width="20%">BJR Census This Semester</td>
				<td align="left">
					<input type="text" size="15" name="pressemester_abw" id="pressemester_abw" value="<%=pressemester_abw%>" maxlength="32" style="background:#EEEEEE;" readonly onkeyup="return hanyaAngka(this); hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();">
				</td>
				
			</tr>
			<tr>
			
				<td align="left" width="20%">Harvest Date</td>
				<td align="left">
					<input class="datepicker" type="text" size="15" name="harvestdate" id="harvestdate" value="<%=harvestdate%>">
				</td>
				<%if cdbl(to_nol(spb_type)) > 0 then %>	
				<td align="left" width="20%">Estate Weight</td>
				<td align="left">
					<input type="text" size="15" name="bunch_estateweight" value="<%=bunch_estateweight%>" autocomplete="off" maxlength="32" onkeypress="return hanyaAngka(event);" >
				</td>
				<%else%>
				<td align="left" width="20%">Estate Weight</td>
				<td align="left">
					<input type="text" size="15" id="bunch_estateweight" name="bunch_estateweight" value="<%=bunch_estateweight%>" autocomplete="off" maxlength="32" onkeypress="return hanyaAngka(event);" onkeyup="hitungrumus();" onchange=" hitungrumus();" onblur="hitungrumus();" style="background:#EEEEEE;" readonly>
				</td>
				<%end if%>
				
			</tr>
			<tr>
				<td align="left" width="20%">Bunch</td>
				<td align="left">
					<input type="text" size="15" id="bunch" name="bunch" style="background:#FFFFCC;" value="<%=bunch%>" autocomplete="off" maxlength="32" onkeypress="return hanyaAngka(event);" onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();">
					<font color="#FF0000">*</font>
				</td>
				<td align="left" width="20%">Brondolan</td>
				<td align="left">
					<input type="text" size="15" id="brondolan" name="brondolan" style="background:#FFFFCC;" value="<%=brondolan%>" autocomplete="off" maxlength="32" onkeypress="return hanyaAngka(event);" onkeyup="hitungrumus();" onchange="hitungrumus();" onblur="hitungrumus();">
					<font color="#FF0000">*</font>
				</td>
			</tr>
			 
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="SPB_NO" value="<%=SPB_NO%>">
					<input type="hidden" name="RECEPTIONDATE" value="<%=RECEPTIONDATE%>">
					<input type="hidden" name="CROPCODE" value="<%=CROPCODE%>">
					<input type="hidden" name="PRODUCTCODE" value="<%=PRODUCTCODE%>">
					<input type="hidden" name="OWN" value="<%=OWN%>">
					<input type="hidden" name="VEHICLE" value="<%=VEHICLE%>">
					<input type="hidden" name="AGREEMENTCODE" value="<%=AGREEMENTCODE%>">
					<input type="hidden" name="DRIVER" value="<%=DRIVER%>">
					<input type="hidden" name="MILL" value="<%=MILL%>">
					<input type="hidden" name="TRANSPORTTYPE" value="<%=TRANSPORTTYPE%>">
					<input type="hidden" name="SPB_TYPE" value="<%=SPB_TYPE%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					<!-- Validasi Period Control In Here -->
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
					<!-- End Validasi Period Control In Here -->
				</td>
			</tr>
		</table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="2">&nbsp;</td>
		<td align="left">SPB NO</td>
		<td align="left">RECEPTIONDATE</td>
		<td align="left">HARVEST DATE</td>
		<td align="left">CROP</td>
		<td align="left">PRODUCT</td>
		<td align="left">VEHICLE STATUS</td>
		<td align="left">VEHICLE CODE</td>
		<td align="left">CONTRACT</td>
		<td align="left">DRIVER</td>
		<td align="left">MILL</td>
		<td align="left">TRANSPORT TYPE</td>
		<td align="left">SPB TYPE</td>
		<td align="left">FIELDCODE</td>
		<td align="left">BJR CENSUS THIS SEMESTER</td>		
		<td align="left">BUNCH</td>
		<td align="left">BRONDOLAN</td>
		<td align="left">ESTATE WIGHT</td>
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
				<!-- Validasi Period Control In Here -->
				<td>
					<a href="spbharvesting.asp?fcBA=<%=rsVh("fcBA")%>&spb_no=<%=rsVh("spbno")%>&fieldcode=<%=rsVH("fieldcode")%>&productcode=<%=rsVH("productcode")%>&cropcode=<%=rsVH("cropcode")%>&harvestdate=<%=rsVH("harvestdate")%>&driver=<%=rsVH("driver")%>&mill=<%=rsVH("mill")%>&agreementcode=<%=rsVH("agreementcode")%>&own=<%=rsVH("own")%>&vehicle=<%=rsVH("vehicle")%>&spb_type=<%=rsVH("spb_type")%>&bunch=<%=rsVH("bunch")%>&brondolan=<%=rsVH("brondolan")%>&bunch_estateweight=<%=rsVH("bunch_estateweight")%>&pressemester_abw=<%=rsVH("pressemester_abw")%>&receptiondate=<%=rsVH("receptiondate")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("spb_no")%>" border="0"></a>
				</td>
				<td>
					<a href="spbharvesting.asp?fcBA=<%=rsVh("fcBA")%>&spb_no=<%=rsVh("spbno")%>&fieldcode=<%=rsVH("fieldcode")%>&productcode=<%=rsVH("productcode")%>&cropcode=<%=rsVH("cropcode")%>&harvestdate=<%=rsVH("harvestdate")%>&driver=<%=rsVH("driver")%>&mill=<%=rsVH("mill")%>&agreementcode=<%=rsVH("agreementcode")%>&own=<%=rsVH("own")%>&vehicle=<%=rsVH("vehicle")%>&spb_type=<%=rsVH("spb_type")%>&bunch=<%=rsVH("bunch")%>&brondolan=<%=rsVH("brondolan")%>&bunch_estateweight=<%=rsVH("bunch_estateweight")%>&pressemester_abw=<%=rsVH("pressemester_abw")%>&receptiondate=<%=rsVH("receptiondate")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("spb_no")%>" border="0"></a>
				</td>
				<!-- End Validasi Period Control In Here -->
			    <td align="left" title="<%=rsVH("spbno").Name%>" valign="top"><%=rsVH("spbno")%></td>
			    <td align="left" title="<%=rsVH("receptiondate").Name%>" valign="top"><%=rsVH("receptiondate")%></td>
			    <td align="left" title="<%=rsVH("harvestdate").Name%>" valign="top"><%=rsVH("harvestdate")%></td>
			    <td align="left" title="<%=rsVH("cropcode").Name%>" valign="top"><%=rsVH("cropcode")%></td>
			    <td align="left" title="<%=rsVH("productcode").Name%>" valign="top"><%=rsVH("productcode")%></td>
			    <td align="left" title="<%=rsVH("own").Name%>" valign="top"><%=rsVH("own")%></td>
			    <td align="left" title="<%=rsVH("vehicle").Name%>" valign="top"><%=rsVH("vehicle")%></td>
			    <td align="left" title="<%=rsVH("agreementcode").Name%>" valign="top"><%=rsVH("agreementcode")%></td>
			    <td align="left" title="<%=rsVH("driver").Name%>" valign="top"><%=rsVH("driver")%></td>
			    <td align="left" title="<%=rsVH("mill").Name%>" valign="top"><%=rsVH("mill")%></td>
			    <td align="left" title="<%=rsVH("transporttype").Name%>" valign="top"><%=rsVH("transporttype")%></td>
			    <td align="left" title="<%=rsVH("spb_type").Name%>" valign="top"><%=rsVH("spb_type")%></td>
			    <td align="left" title="<%=rsVH("fieldcode").Name%>" valign="top"><%=rsVH("fieldcode")%></td>
			    <td align="left" title="<%=rsVH("pressemester_abw").Name%>" valign="top"><%=rsVH("pressemester_abw")%></td>
			    <td align="left" title="<%=rsVH("bunch").Name%>" valign="top"><%=rsVH("bunch")%></td>
			    <td align="left" title="<%=rsVH("brondolan").Name%>" valign="top"><%=rsVH("brondolan")%></td>
			    <td align="left" title="<%=rsVH("bunch_estateweight").Name%>" valign="top"><%=rsVH("bunch_estateweight")%></td>
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
   </form>
<% else
		tempStringError = "Jika Vehicle Status Diisi Own, maka Vehicle Code Diisi dengan Kode Kendaraan. Jika Vehicle Status Diisi Contract, maka Vehicle Code Diisi dengan Nomor Polisi."				 
		response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>" %>
<%End if%>
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->
 <script language="JavaScript" type="text/javascript">
//You should create the validator only after the definition of the HTML form
  var frmvalidator  = new Validator("eForm");
  frmvalidator.addValidation("EMPLOYEECODE","req","Please enter Emp.Code"); 

</script>
 <script type="text/javascript">
	function handleClick(avail){
		if(avail.checked == 1){
			avail.value = 1;
		} else {
			avail.value = 0;
		}
	}
	
	function cekspbtype() {
		if($('#spb_type').is(":checked")){
			$('#bunch_estateweight').css("background-color", "#fff");
			$('#bunch_estateweight').prop('readOnly', false);
		}else{
			$('#bunch_estateweight').css("background-color", "#eee");
			$('#bunch_estateweight').prop('readOnly', true);
		}
	}
</script>
<script type="text/javascript">	

	// $('#spb_type').click( function () {
	// 	//$('.item').prop('checked', this.checked);
	// 	//$('.item').val(this.checked ? 1 : 0);
	// 	if($('#spb_type').is(":checked")){
	// 		$('#bunch_estateweight').css("background-color", "#fff");
	// 		$('#bunch_estateweight').prop('readOnly', false);
	// 	}else{
	// 		$('#bunch_estateweight').css("background-color", "#eee");
	// 		$('#bunch_estateweight').prop('readOnly', true);
	// 	}
	// })
	
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
		// document.getElementById("agreementcode").style.backgroundColor="#eee";
		// $( "#agreementcode" ).prop( "disabled", true );
		// $( "#vehicle" ).prop( "disabled", false );
		// $( "#driver" ).prop( "disabled", false );
		// $(document).ready(function() {
				$('#agreementcode').css("background-color", "#eee");
				$('#agreementcode').prop('readOnly', true);
				$('#agreementcode').val(null);
				$('#vehicle').css("background-color", "#fff");
				$('#vehicle').prop('readOnly', false);
				$('#driver').css("background-color", "#fff");
				$('#driver').prop('readOnly', false);

		 	$('#own').click(function() {
		  	if( $(this).val() === 'OWN') {
		   		$('#agreementcode').css("background-color", "#eee");
				$('#agreementcode').prop('readOnly', true);
				// $('#vehicle').css("background-color", "#fff");
				// $('#vehicle').prop('readOnly', false);
				$('#driver').css("background-color", "#fff");
				$('#driver').prop('readOnly', false);
		      	// $('#vehicle').val(null);
		      	$('#driver').val(null);
				// $( "#agreementcode" ).prop( "disabled", true );
				// $( "#vehicle" ).prop( "disabled", false );
				// $( "#driver" ).prop( "disabled", false );
		    } else if( $(this).val() === 'CONTRACT') {      
		    	$('#agreementcode').css("background-color", "#fff");
				$('#agreementcode').prop('readOnly', false);
				// $('#vehicle').css("background-color", "#eee");
				// $('#vehicle').prop('readOnly', true);
				$('#driver').css("background-color", "#eee");
				$('#driver').prop('readOnly', true);
				$('#agreementcode').val(null);
		  //   	$( "#agreementcode" ).prop( "disabled", false );
		  //   	$( "#vehicle" ).prop( "disabled", true );
		  // 	$( "#driver" ).prop( "disabled", true );
		    } else {      
		    	$('#agreementcode').css("background-color", "#fff");
				$('#agreementcode').prop('readOnly', false);
				// $('#vehicle').css("background-color", "#eee");
				// $('#vehicle').prop('readOnly', true);
				$('#driver').css("background-color", "#fff");
				$('#driver').prop('readOnly', false);
				$('#agreementcode').val(null);
		  //   	$( "#agreementcode" ).prop( "disabled", false );
		  //   	$( "#vehicle" ).prop( "disabled", true );
		  // 	$( "#driver" ).prop( "disabled", true );
		    }
		  });
		 
		// });
		hitungrumus();
		getOvertime();
		cekspbtype();

	});
</script>
</body>
</html>
