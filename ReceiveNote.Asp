<%
	response.buffer = True
	validatesession
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="qtyOnHandFunction.inc"-->

<html>
<head>
	<title>ePlaS - Receive Note</title>
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
var getSupplierRG = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedsupplier = document.getElementById("selectedsupplier").value;
	var PostStr = 'fccode='+selectedsupplier;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getSupplierRG.asp', 'POST', PostStr, showSupplierRG);
}

// The function for handling the response from the server
var showSupplierRG = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("supplierlist").innerHTML = "";
	document.getElementById("supplierlist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getPOgrn = function () {
	// build up the post string when passing variables to the server side page
	
	var xcode = document.getElementById("selectedpo").value;
	var selectedsupplier = document.getElementById("selectedsupplier").value;
	var PostStr = 'xcode='+xcode+'&selectedsupplier='+selectedsupplier;;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getPOgrn.asp', 'POST', PostStr, showPogrn);
}

// The function for handling the response from the server
var showPogrn = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("polist").innerHTML = "";
	document.getElementById("polist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}


// launched from button click 
var getliststo = function () {
	// build up the post string when passing variables to the server side page
	
	var xcode = document.getElementById("selectedpo").value;
	var selectedstogrn = document.getElementById("selectedstogrn").value;
	var PostStr = 'xcode='+xcode+'&stocktransno2='+selectedstogrn;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getstonogrn.asp', 'POST', PostStr, showliststo);
}

// The function for handling the response from the server
var showliststo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("stolist").innerHTML = "";
	document.getElementById("stolist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}

// launched from button click 
var getlistdo = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedstogrn = document.getElementById("selectedstogrn").value;
	var selecteddogrn = document.getElementById("selecteddogrn").value;
	var PostStr = 'stocktransno2='+selectedstogrn+'&dono2='+selecteddogrn;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getdonogrn.asp', 'POST', PostStr, showlistdo);
}

// The function for handling the response from the server
var showlistdo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("dolist").innerHTML = "";
	document.getElementById("dolist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}



</script>
<script language="JavaScript">
        function EditHeader() {
		var fccode = document.getElementById("fcCode").value;
		var rgdate = document.getElementById("DATECREATED").value;
		//alert(rgdate);
		window.location.href = "ReceiveNote.Asp?fcCode="+fccode+"&DATECREATED="+rgdate+"&txtJenis=EditHeader";
}		
</script>  

 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

function harusangka(evt) {
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

function getValidasiQty(i)
{
	var QuantityOrder_ = parseInt($("#QuantityOrder_" + i).val());
	var QtyReceived_ = parseInt($("#QtyReceived_" + i).val());
	if (QtyReceived_>QuantityOrder_){
		alert("QTY RECEIVED LEBIH BESAR DARI QTY ORDER");
		$("#QtyReceived_" + i).val(0)
	}
}

var reloadGrnCode = function () {
	var grndate =$( ".datepicker" ).val();
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Procurement&txtTable=ReceiveNote&grndateurl="+grndate );
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
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
	<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if session("fcba") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;">
	
<%
dim ScriptName,modulFile,executeQry
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
Function cmdCheckHistNursery(initfcba, initNURSERYCODE, initLineNoHist)
	dim qCheckHistNursery,rsCheckHistNursery,qTotalDataNursery
	
	qCheckHistNursery = " select count(1) totalData" & _
						" from stocktranshist" & _
						" where fcba = '" & initfcba & "'" & _
						" and targetcode = '" & initNURSERYCODE & "'" & _
						" and lineno > " & initLineNoHist & ""
						
	set rsCheckHistNursery = Server.CreateObject("ADODB.Recordset")
	rsCheckHistNursery.Open qCheckHistNursery, connstr
	if not rsCheckHistNursery.eof then
		qTotalDataNursery = cdbl(rsCheckHistNursery("totalData"))
	else
		qTotalDataNursery = 0
	end if 	

	cmdCheckHistNursery = qTotalDataNursery
End Function
%>
		<%
Dim xcode,xstoreCode,txtstockCode,txtKolomtransaksi,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,v_CrossCekPO,fcba
Dim Remarks,txtLokasiGudang
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
Dim Jobcode, LocationType, LocationCode, Reference, Remark
Dim qLineMax, rsqLineMax, tempMaxLine
Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
Dim sqlSQLCheck, rsSQL, qtyRemained, qtyDoRemaind
Dim qCheckStore, rsqCheckStore
Dim txtLocationType	

Dim txtGRNCode : txtGRNCode = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
Dim txtLocationCode : txtLocationCode = ""
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
Dim Cek_IsAccClosing : Cek_IsAccClosing = ""
Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD,postingflg,postingflgtxt
txtKolomtransaksi 	= "RN"

dim stocktransno2,dono2,IsUseExtSys

	txtTable =Request("txtTable")
	txtModule=Request("txtModule")
	txtSearch=Request("txtSearch")
	txtField =Request("txtField")
			
	if Request("fcba") <> "" then
	fcba = Request("fcba")
	else	
	fcba = session("fcba") 
	end if
 
	txtJenis = request("txtJenis")
			
	'---- Data Header
	fcCode = request("fccOde")
 
	If Len(Trim(fcCode)) = 0 Then fcCode = "(NEW)"
	
	if request("grndateurl") <>"" then
		DATECREATED   = request("grndateurl")
	elseif  Request("DATECREATED") <> "" then
	    DATECREATED	=	Request("DATECREATED")
	else
		DATECREATED	=	textDate(date)
	end if
	
	if KonversiTanggalKeDbl(DATECREATED) > KonversiTanggalKeDbl(textDate(date)) then
		DATECREATED = textDate(date)
	end if
 
	ACCOUNTINGPERIOD  = left(DATECREATED,2)
	ACCOUNTINGYEAR    = right(DATECREATED,4)
	
	DATECREATED = FormatDateTime(DATECREATED)
			
	xcode = Request("xcode")
	xstoreCode = Request("xstoreCode")
	stocktransno2 = Request("stocktransno2")
	dono2 = Request("dono2")
			
	HeaderRemarks = Request("HeaderRemarks")
			
	'---- Data Details
	JobCode=Request("JobCode")
	LocationType=Request("LocationType")
	LocationCode=Request("LocationCode")
	Reference=Request("Reference")
	Remarks=Request("Remarks")
	Amount=Request("Amount")
	PurchaseLineNoKey = Request("PurchaseLineNoKey")
	selectedsupplier = Request("selectedsupplier")
	selectedPO = Request("selectedPO")
	selectedstogrn = Request("selectedstogrn")
	selecteddogrn = Request("selecteddogrn")
	
Cek_IsAccClosing = StatusAccClosingTgl(DATECREATED,fcba) 
if 	Cek_IsAccClosing = "FALSE" then  			
	if  Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="EditHeader" or Request("txtJenis")="Approved" Then
		
		REMARKS	= request("REMARKS")
			
		Set rsPO = Server.CreateObject("ADODB.Recordset")
		rsPO.Open " select company_api.IsUseExtSys(businessunit_api.get_CompanyCode('" & fcba & "')) IsUseExtSys, PURCHASEORDER_API.GetIsInBridge('" & xcode & "','" & fcba & "') IsInBridge from dual" , connstr
		IF  Not rsPO.eof then
			IsUseExtSys =  rsPO("IsUseExtSys")
			IsInBridge = rsPO("IsInBridge")
		End if
		rsPO.Close
		set rsPO = Nothing
				
		if  Request("txtJenis")= "Header" then
		    Set rsPO = Server.CreateObject("ADODB.Recordset")
				rsPO.Open " SELECT a.invoicecode,count(b.receivenotecode) grnno FROM invoice a INNER JOIN invoicedetail b ON (a.invoicecode=b.invoicecode AND a.fcba=b.fcba) "& _
					      " WHERE  b.receivenotecode ='" & fcCode & "' group by a.invoicecode", connstr
				If  Not rsPO.eof then
					grnno = cdbl(rsPO("grnno"))
					invno = rsPO("invoicecode")
				Else
					grnno = 0
					invno = "-"
				End if
			rsPO.Close
			set rsPO = Nothing
			
			connstr.BeginTrans
		    
			IF grnno = 0 THEN
			
				Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
				If  Cek_IsDailyClosing = "FALSE" then

					Set rsPO = Server.CreateObject("ADODB.Recordset")
					rsPO.Open   " select pocode" & _
								" from receivenote" & _
								" where receivenotecode = '" & fcCode & "'" & _
								"       and fcBA = '" & fcba & "'", connstr
					IF  Not rsPO.eof then
						txtNoPO =  rsPO("pocode")
					End if
					rsPO.Close
					set rsPO = Nothing

					q =  	" MERGE INTO purchaseorderdetails a " & _
							" USING" & _
							" (" & _
							" 	SELECT y.pocode,y.fcba,x.purchaseitemcode,x.purchaselinenokey,x.quantity QtyRcv FROM  receivenotedetail  x" & _
							" 	INNER JOIN receivenote y ON (x.receivenotecode=y.receivenotecode AND x.fcba=y.fcba)" & _
							" 	WHERE " & _
							" 	x.receivenotecode='" & fcCode & "' AND x.fcba='" & fcba & "'" & _
							" ) b ON (a.pocode=b.pocode AND a.fcba=b.fcba AND a.itemcode=b.purchaseitemcode AND a.linenokey=b.purchaselinenokey)" & _
							" WHEN MATCHED THEN" & _
							" 	UPDATE SET a.received = a.received-b.QtyRcv " & _
							" 	WHERE a.received > 0" 
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry

					q = " delete from receivenotedetail" & _
						" where receivenotecode = '" & fcCode & "'" & _
						"       and fcba = '" & fcba & "'"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry

					q = " delete from receivenote" & _
						" where receivenotecode = '" & fcCode & "'" & _
						"       and fcba = '" & fcba & "'"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry

					fcCode = "(NEW)"
					
					if  err.number = 0 then
						connStr.CommitTrans
					end if
					
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
			
			Else
				err.number = 1001
				err.description = "MAAF Sudah ter-INVOICE # """ & invno & """, silakan hapus terlebih dahulu transaksi invoice tersebut !!! "
			
			End if
					
			cmdView	= "True"
					
		elseif  Request("txtJenis")= "EditHeader" then		
			connstr.BeginTrans
					
			Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
			If  Cek_IsDailyClosing = "FALSE" then
					
				q = " Update receivenote" & _
					" Set RNDATE = to_date('" & DATECREATED & "','mm/dd/yyyy')" & _ 
					" where receivenotecode = '" & fcCode & "'" & _
					" and fcba = '" & fcba & "'"
				connStr.execute(q)
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & q
				ExecuteAndLog modulFile,executeQry
				
				if  err.number = 0 then
					connStr.CommitTrans
				end if
					
			else
				err.number = 1001
				err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
					
			cmdView	= "True"
		
		Elseif Request("txtJenis")= "Approved" then
			' ============================================================
			' APPROVED: hanya update status & postingflg di ReceiveNote
			' (StockOnHand, StockLedger, purchaseorderdetails
			'  sudah diproses saat txtJenis=New)
			' ============================================================
			strQueryStatus = "UPDATE ReceiveNote SET POSTINGFLG='1',rowstateid='Approved',lastupdate=sysdate,fcedit='" & session("CodeName") & "' WHERE fcba = '" & fcba & "' AND receivenotecode = '" & fccode & "'"
			connStr.execute(strQueryStatus)
			modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
			executeQry	= "Query executed:" & strQueryStatus
			ExecuteAndLog modulFile,executeQry
					
			cmdView	= "True"
	
		elseif  Request("txtJenis")= "Details" then
			Set rsPO = Server.CreateObject("ADODB.Recordset")
				rsPO.Open " SELECT a.invoicecode,count(b.receivenotecode) grnno FROM invoice a INNER JOIN invoicedetail b ON (a.invoicecode=b.invoicecode AND a.fcba=b.fcba) "& _
					      " WHERE  b.receivenotecode ='" & fcCode & "' group by a.invoicecode", connstr
				If  Not rsPO.eof then
					grnno = cdbl(rsPO("grnno"))
					invno = rsPO("invoicecode")
				Else
					grnno = 0
					invno = "-"
				End if
			rsPO.Close
			set rsPO = Nothing
		    
			connstr.BeginTrans
			IF grnno = 0 THEN
		
				Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
				If  Cek_IsDailyClosing = "FALSE" then
						
					Set rsPO = Server.CreateObject("ADODB.Recordset")
					rsPO.Open " Select pocode,DELIVERYORDERCODE,StockOrderNo, STORECODE from ReceiveNote where RECEIVENOTECODE='" & fcCode & "'" & _
						      " and fcBA='" & fcba & "'", connstr
					If  Not rsPO.eof then
							txtNoPO =  rsPO("pocode")
							stocktransno2 =  rsPO("StockOrderNo")
							dono2 =  rsPO("DELIVERYORDERCODE")
							txtStoreCode = rsPO("STORECODE")
					End if
					rsPO.Close
					set rsPO = Nothing
						
					q = " Select  LineNoHist" & _
						"       , quantity" & _
						"       , PURCHASEITEMCODE" & _
						"       , PURCHASELINENOKEY" & _
						"       , LOCATIONTYPE" & _
						"       , LOCATIONCODE" & _
						" from ReceiveNoteDetail" & _
						" where RECEIVENOTECODE = '" & fcCode & "'" & _
						"       and PURCHASELINENOKEY = '" & Request("PURCHASELINENOKEY") & "'" & _
						"       and fcBA = '" & fcba & "'"
						
					Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					rsReceiveDetail.Open q, connstr
							
					IF  Not rsReceiveDetail.eof then
							txtLineNoHist = rsReceiveDetail("LINENOHIST")
							originalQuantity = CDbl(rsReceiveDetail("QUANTITY"))
							txtItemCodeDelete = rsReceiveDetail("PURCHASEITEMCODE")
							txtItemLineKeyDelete = rsReceiveDetail("PURCHASELINENOKEY")
							txtLocationType = rsReceiveDetail("LOCATIONTYPE")
							txtLocationCode = rsReceiveDetail("LOCATIONCODE")
					End if
					rsReceiveDetail.Close
					set rsReceiveDetail = Nothing					

					qCheckStore = " select LOCATIONTYPE_API.get_item('" & txtLocationType & "') item from dual "

					set rsqCheckStore = Server.CreateObject("ADODB.Recordset")
						rsqCheckStore.Open qCheckStore, connstr
					if  Not rsqCheckStore.eof then
							tempItem = rsqCheckStore("item")
					else
							tempItem = ""
					end if 
					rsqCheckStore.Close
					Set rsqCheckStore = Nothing
								
					if  tempItem = "MATERIAL" then
						
						If  txtLocationType = "MN" Then
								v_CrossCekDetail = cmdCheckHistNursery(fcba, txtLocationCode, txtLineNoHist)
						
								qLessZero = " select islesszero" & _
											" from stock" & _
											" where fccode = '" & Request("ItemCode") & "'" & _
											"       and fcba = '" & fcba & "'" 
										
								set rsLessZero = Server.CreateObject("ADODB.Recordset")
								rsLessZero.Open qLessZero, connstr
								If  not rsLessZero.eof then
									tempislesszero = rsLessZero("islesszero")
								Else
									tempislesszero = "FALSE"
								End if
										
								qQuantityOnHand =   " select quantity" & _
													" from stockonhandnursery" & _
													" where storecode = '" & txtStoreCode & "'" & _
													"       and stockcode  = '" & Request("ItemCode") & "'" & _
													"       and nurserycode  = '" & txtLocationCode & "'" & _
													"       and fcba = '" & fcba &"' "
													  
								set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
								rsQuantityOnHand.Open qQuantityOnHand, connstr
								If  Not rsQuantityOnHand.eof then
										QuantityOnHandNursery = cdbl(rsQuantityOnHand("quantity"))
								Else
										QuantityOnHandNursery = 0
								End if
										
								If  tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandNursery - originalQuantity) >= 0)) Then
										
									q =	" Delete from stocktranshist" & _
										" where lineno = " & txtLineNoHist & "" & _
										"       and fcba = '" & fcba & "'"
									connStr.execute(q)
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry

									q = " update stockonhandnursery " & _
										" set quantity = quantity - " & originalQuantity & "" & _
										"     , lastupdate = sysdate " & _
										"     , lasttime = to_char(sysdate,'hh24:mi') " & _
										" where storecode = '" & txtStoreCode & "'" & _
										"       and stockcode  = '" & Request("ItemCode") & "'" & _ 
										"       and nurserycode  = '" & txtLocationCode & "'" & _ 
										"       and fcba = '" & fcba &"'"
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry

									q = " delete from receivenotedetail" & _
										" where receivenotecode = '" & fcCode & "'" & _
										"       and purchaselinenokey = '" & PurchaseLineNoKey & "'" & _
										"       and fcba = '" & fcba & "'"
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry

									q = " update purchaseorderdetails" & _
										" set rowstateid = 'Opened'" & _
										" , received = received - "& originalQuantity &"" & _
										" where pocode = '" & txtNoPo & "'" & _
										"       and linenokey = '" & txtItemLineKeyDelete & "'" & _
										"       and fcBA = '" & fcba & "'"
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry

									v_CrossCekPO =  cmdCrossCekPOUpdate (txtNoPo, txtItemLineKeyDelete,  originalQuantity)
											
									if  err.number=0 then
										connStr.CommitTrans
									End if
											
								Else
									err.number = 1001
									Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
								End If
								
						Else
								v_CrossCekDetail = cmdCheckHist(txtLineNoHist, Request("ItemCode"))
								
								qLessZero = " select islesszero" & _
											" from stock" & _
											" where fccode = '" & Request("ItemCode") & "'" & _
											"       and fcba = '" & fcba & "'" 
								
								set rsLessZero = Server.CreateObject("ADODB.Recordset")
								rsLessZero.Open qLessZero, connstr
								if not rsLessZero.eof then
									tempislesszero = rsLessZero("islesszero")
								else
									tempislesszero = "FALSE"
								end if
									
								If  tempislesszero = "TRUE" Or (tempislesszero = "FALSE") Then

									q = "Delete from stocktranshist " & _
										" where " & _
										" lineno = " & txtLineNoHist & " and " & _
										" fcba = '" & fcba & "'"
								
									connStr.execute(q)
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry
																	
									q = " Update STOCKTRANSFERORDERDTL sto  set sto.dorcvqty = sto.dorcvqty  - " & originalQuantity & _
											", rowstateid = '', LASTUPDATE = sysdate " & _
											", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
											" where sto.pocode = '" & txtNoPo & "' " & _ 
											"	    and sto.linenokey = '" & txtItemLineKeyDelete & "'" & _ 
											"	    and sto.stono =   '" & stocktransno2 & "' " & _  
											"	    and sto.fcba =  '" & fcba & "'" 
									connStr.execute(q)	
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry

									q = " Delete from RECEIVENOTEDETAIL where  RECEIVENOTECODE='" & fcCode & "' and PURCHASELINENOKEY='" & PurchaseLineNoKey & _
										    " ' and fcBA='" & fcba & "'"
									connStr.execute(q)
									modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
									executeQry	= "Query executed:" & q
									ExecuteAndLog modulFile,executeQry

									if err.number=0 then
									    If  IsUseExtSys = "Y"  Then
											strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.DeleteGRNLineNoToBridge('" & fcba & "','" & fcCode & "','" & PurchaseLineNoKey & "'); END; "
											connStr.execute(strQueryStatus)
										End if	
										connStr.CommitTrans
									End if
									
								Else
									err.number = 1001
									Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
								End If
								
						End If
							
					Else
						
						q = " Update STOCKTRANSFERORDER sto  set sto.dorcvqty = sto.dorcvqty - " & originalQuantity & _
							"   , rowstateid = '', LASTUPDATE = sysdate " & _
							"   , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
							" where sto.trackno = '" & txtNoPo & "' " & _ 
							"	 and sto.tracklineno = '" & txtItemLineKeyDelete & "'" & _ 
							"	 and sto.stono =   '" & stocktransno2 & "' " & _  
							"	 and sto.dono =  '" & dono2 & "' " & _  
							"	 and sto.fcba =  '" & fcba & "'" 
						connStr.execute(q)	
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry

						q = " Update deliveryorderpurchasedtl  " & _
							" set receivedquantity = receivedquantity -  " & originalQuantity & _
							"	 where stono =   '" & stocktransno2 & "' " & _  
							"	 and do_no =  '" & dono2 & "' " & _  
							"	 and fcba =  '" & fcba & "'" 
						connStr.execute(q)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry

						q = " Delete from RECEIVENOTEDETAIL where  RECEIVENOTECODE='" & fcCode & "' and PURCHASELINENOKEY='" & PurchaseLineNoKey & _
							"' and fcBA='" & fcba & "'"
								
						connStr.execute(q)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry

						If  err.number=0 then
							connStr.CommitTrans
						End if
							
					End if
						
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
			
			Else
				err.number = 1001
				err.description = "MAAF Sudah ter-INVOICE # """ & invno & """, silakan hapus terlebih dahulu transaksi invoice tersebut !!! "
			
			End if
			
			cmdView	= "True"
		
		ElseIf  Request("txtJenis")="New" then
				
				connstr.BeginTrans
				Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
			
				If Cek_IsDailyClosing = "FALSE" then
					
					If  Ucase(fcCode) = Ucase("(NEW)") Then
						oSql = " select nvl(controlnumber_api.get_nextnumber(upper('receivenote')" & _
							   "                                         , " & year(textDate(DATECREATED)) & "" & _
							   "                                         , " & month(textDate(DATECREATED)) & "" & _
							   "                                         , '" & fcba & "'" & _
							   " ), 'FALSE') grncode from dual"
							   
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.Open oSql, connstr, 3, 3
						If Not oRs.EOF Then
							txtGRNCode = oRs("grncode")
						End If
						oRs.Close
						Set oRs = Nothing
						fcCode = txtGRNCode
					End If

					If txtGRNCode = "FALSE" Then
						err.Number = 1001
						err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
					End If
					
					Dim rsPO
					Set rsPO = Server.CreateObject("ADODB.Recordset")
					rsPO.Open " Select RECEIVENOTECODE from ReceiveNote where RECEIVENOTECODE='" & fcCode & "'" & _
							  " and fcBA='" & fcba & "'", connstr
					If  rsPO.eof then	    
						
						Dim txtCodeSupplier
						q = "Select SUPPLIERCODE From PurchaseOrder where pocode ='" & xcode & "'"  & _
						    " and fcBA='" & fcba & "'"
						Dim rsLPO  
						set rsLPO = Server.CreateObject("ADODB.Recordset")
						rsLPO.Open q,connstr,1,1
						If  Not rsLPO.eof then
							txtCodeSupplier=rsLPO("SupplierCode")
						End if 
						rsLPO.Close
						set rsLPO = Nothing

						q = " Insert into ReceiveNote (RECEIVENOTECODE, RNDATE, pocode,STOCKORDERNO,DELIVERYORDERCODE,SupplierCode, Storecode, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME,ROWSTATEID) " & _
						    " values ('" & fcCode & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & xcode & "','" & stocktransno2 & "','" & dono2 & "','" & txtCodeSupplier & "','" & _
						    xstoreCode & "','" & _
						    Session("CodeName") & "','" & Session("CodeName") & "','" & _
						    Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'),'Planned')"
						connStr.execute(q)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry
					End if
					rsPO.Close	
					Set rsPO = Nothing

					Dim n
	  
					For n=1 to (cdbl(request("no"))-1) 
						If  request("LT_"&TRIM(n))<>"" and request("LC_"&TRIM(n))<>"" and request("QtyReceived_"&TRIM(n))<>"" then
							If  request("JobCode_"&TRIM(n))="" Then
								Dim txtControlJOb
								q = "select sg.CONTROLJOB as CONTROLJOB  from stockgroup sg,stock s" & _
									" where (sg.FCCODE = s.STOCKGROUPCODE and sg.fcba = s.fcba) " & _
									" and s.FCCODE='" & request("Stock_"&TRIM(n)) & "'" & _
									" and s.fcBA='" & fcba & "'" & _
									" and ROWNUM =1 "
								set rsStock = Server.CreateObject("ADODB.Recordset")
								rsStock.Open q, connstr
								if not rsStock.eof then
									txtControlJob=rsStock("CONTROLJOB")
								End if 
							Else
								txtControlJob=request("JobCode_"&TRIM(n))
							End if
								 
							If  request("AMOUNTReceived_"&TRIM(n)) = "" Then
								Dim txtamountreceived,txtQtyPO
									
								q = " Select p.linenokey" & _
									"      , p.itemcode" & _
									"      , s.fcname" & _
									"      , sum(p.quantity) as quantity" & _
									"      , sum(amount) as amount" & _
									" From purchaseorderdetails p" & _
									"    , stock s" & _
									" where (p.itemcode = s.fccode and p.fcba = s.fcba)" & _
									"       and pocode = '" & xcode & "'" & _
									"       and s.fccode = '" & request("Stock_"&TRIM(n)) & "'" & _
									"       and p.linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
									"       and p.fcba = '" & fcba & "'" & _
									" group by p.linenokey" & _
									"          , p.itemcode" & _
									"          , s.fcname"
								set rsamr = Server.CreateObject("ADODB.Recordset")
								rsamr.Open q, connstr
									
								if  not rsamr.eof then
										txtamountreceivedpo = Round((cdbl(rsamr("Amount"))/cdbl(rsamr("Quantity")))*request("QtyReceived_"&TRIM(n)), 2)
										PurchaseLineNoKey = rsamr("lineNoKey")
										txtQtyPO = cdbl(rsamr("Quantity"))
								End if 
								rsamr.close
								set rsamr = nothing	
									
								q = " select s.stono, " & _
									"		 s.pocode," & _
									"		 s.linenokey," & _
									"		 s.itemcode, " & _
									"		 s.do_quantity," & _
									"		 s.amount " & _
									"  from stocktransferorderdtl s " & _
									"  where s.pocode = '" & xcode & "'" & _
									"        and s.lineNoKey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
									"	     and s.ItemCode = '" & request("Stock_"&TRIM(n)) & "'" & _
									"	     and s.stono =   '" & stocktransno2 & "' " & _  
									"	 and s.fcba =  '" & fcba & "'" 

								set rsamr = Server.CreateObject("ADODB.Recordset")
								rsamr.Open q, connstr
								if  not rsamr.eof then
										txtamountreceivedsto =  Round((cdbl(rsamr("amount"))/cdbl(txtQtyPO))*cdbl(request("QtyReceived_"&TRIM(n))), 2)
								End if 
								rsamr.close
								set rsamr = nothing	
							End if
								 
							sqlSQLCheck =   " select (quantity - received) as qtyRemained from purchaseorderdetails " & _ 
											" where pocode='" & xcode & "' and LineNoKey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _
											" and fcBA='" & fcba & "'"
												
							Set rsSQL = Server.CreateObject("ADODB.Recordset")
							rsSQL.Open sqlSQLCheck, connstr, adOpenStatic 

							If  not rsSQL.eof then
									qtyRemained = cdbl(rsSQL("qtyRemained"))
							End if
							rsSQL.Close
							set rsSQL = nothing	
                                
							sqlSQLCheck = "  Select  (do_quantity - dorcvqty)  qtydoremaind " & _ 
										  "	 from STOCKTRANSFERORDERDTL  " & _ 
										  "	 where pocode = '" & xcode & "' " & _ 
										  "	 and linenokey = '" & request("LineNoKey_"&TRIM(n)) & "'" & _ 
										  "	 and stono =   '" & stocktransno2 & "' " & _  
										  "	 and fcba =  '" & fcba & "'" 
							Set rsSQL = Server.CreateObject("ADODB.Recordset")
							rsSQL.Open sqlSQLCheck, connstr, adOpenStatic 
							
							if not rsSQL.eof then
								qtyDoRemaind = cdbl(rsSQL("qtydoremaind"))
							End if
							rsSQL.Close
							set rsSQL = nothing	
							
							If  cdbl(qtyDoRemaind) > 0 and  cdbl(qtyDoRemaind) <= cdbl(qtyRemained) then
							    qtyRemained = qtyDoRemaind
							End if									
																	
							If  ( qtyRemained - cdbl(request("QtyReceived_"&TRIM(n))) ) < 0 then
									err.Number = 1001
									err.description = "Total QTY yang diterima tidak boleh lebih besar dari total QTY dalam item PO atau DO"
							Else
	
									qCheckStore = " select LOCATIONTYPE_API.get_item('" & request("LT_"&TRIM(n)) & "') item from dual "
									
									set rsqCheckStore = Server.CreateObject("ADODB.Recordset") 
									rsqCheckStore.Open qCheckStore, connstr
									if  Not rsqCheckStore.eof then
										tempItem = rsqCheckStore("item")
									else
										tempItem = ""
									end if 
									rsqCheckStore.Close
									Set rsqCheckStore = Nothing
									
									If  tempItem = "MATERIAL" then 

										txtstockCode =  request("Stock_"&TRIM(n))
										Quantity	 =  request("QtyReceived_"&TRIM(n))
										txtKolomtransaksi 	= "RN"
										
										qLineMax = "select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
												   "where " & _
												   "fcba = '" & fcba & "'" 
											
										set rsqLineMax = Server.CreateObject("ADODB.Recordset")
										rsqLineMax.Open qLineMax, connstr
										if not rsqLineMax.eof then
											tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
										else
											tempMaxLine = 0
										end if 

										tempMaxLine = tempMaxLine + 1
											
										q = " Insert into StockTransHist ( " & _
											" LINENO, FCBA, DOCUMENTNO, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME) values ( "& _
											tempMaxLine & ",'" & _ 
											fcba & "','" & _
											fccode & "','" & _
											request("Stock_"&TRIM(n)) &"'," & _
											"to_date('" & DATECREATED & "','mm/dd/yyyy') ,'" & _
											"RECEIVED" & "','" & _
											xstoreCode & "'," & _
											"'PR','" & _
											xcode & "','" & _
											request("LT_"&TRIM(n)) & "','" & _
											request("LC_"&TRIM(n)) & "'," & _
											request("QtyReceived_"&TRIM(n)) & ", sysdate, To_Char(sysdate,'HH24:MI') ) "
										connStr.execute(q)
										modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
										executeQry	= "Query executed:" & q
										ExecuteAndLog modulFile,executeQry
									
										If  request("LT_"&TRIM(n))<>"" and request("LC_"&TRIM(n))<>"" and request("Stock_"&TRIM(n)) <>"" and txtControlJOb <> "" then
											q = "Insert into  RECEIVENOTEDETAIL (RECEIVENOTECODE, LOCATIONTYPE, LOCATIONCODE, PURCHASEITEMCODE, " & _
												"JOBCODE, QUANTITY, QTYRETURN, AMOUNTRECEIVED, AMOUNTPO_RECEIVE, AMOUNTSTO_RECEIVE, AMOUNTRETURN, FCENTRY, FCEDIT, FCIP, FCBA, linenohist, purchaselinenokey, LASTUPDATE, LASTTIME) " & _
												"values ('" & fcCode & "','" & request("LT_"&TRIM(n)) & "','" & request("LC_"&TRIM(n)) &  _
												"','" & request("Stock_"&TRIM(n)) & "','" & txtControlJOb & "','" & request("QtyReceived_"&TRIM(n)) & _
												"',0,'"& txtamountreceivedpo + txtamountreceivedsto & "','"& txtamountreceivedpo & "','"& txtamountreceivedsto & "',0," & _
												"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "',"& tempMaxLine & ",'" & request("LineNoKey_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI'))"
											
											connStr.execute(q)
											modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
											executeQry	= "Query executed:" & q
											ExecuteAndLog modulFile,executeQry
										end if

										' ============================================================
										' TAMBAHAN: Update StockOnHand saat New
										' ============================================================
										qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
															 "where storecode = '" & xstoreCode & "' and " & _
															 "stockcode = '" & request("Stock_"&TRIM(n)) & "' and " & _
															 "fcba = '" & fcba & "'"

										set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
										rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
										if not rsTotalRecordOnHand.eof then
											tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
										else
											tempTotalRecordOnHand = 0
										end if
										rsTotalRecordOnHand.Close : set rsTotalRecordOnHand = Nothing

										If tempTotalRecordOnHand = 0 then
											qryinsertstockOnhand = cmdInsertStockOnHand(xstoreCode, request("Stock_"&TRIM(n)), request("QtyReceived_"&TRIM(n)), txtKolomtransaksi)
										Else
											qryUpdatestockOnhand = cmdUpdateStockOnHand(xstoreCode, request("Stock_"&TRIM(n)), request("QtyReceived_"&TRIM(n)), txtKolomtransaksi, "Insert")
										End If

										' ============================================================
										' TAMBAHAN: Update StockLedger saat New
										' ============================================================
										dim qryStockLedger_new, rsStockLedger_new, totalRecordStockLedger_new
										qryStockLedger_new = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _
															 " WHERE STORECODE = '" & xstoreCode & "' and " & _
															 " ITEMCODE  = '" & request("Stock_"&TRIM(n)) & "' and " & _
															 " year  = to_char(to_date('" & DATECREATED & "','mm/dd/yyyy'),'yyyy') and " & _
															 " month = to_char(to_date('" & DATECREATED & "','mm/dd/yyyy'),'mm') and " & _
															 " FCBA  = '" & fcba & "'"

										set rsStockLedger_new = Server.CreateObject("ADODB.Recordset")
										rsStockLedger_new.Open qryStockLedger_new, connstr
										if not rsStockLedger_new.eof then
											totalRecordStockLedger_new = cdbl(rsStockLedger_new("totalRecordStockLedger"))
										else
											totalRecordStockLedger_new = 0
										end if
										rsStockLedger_new.Close : set rsStockLedger_new = Nothing

										If totalRecordStockLedger_new = 0 then
											qryinsertstockledger = cmdInsertStockLedger(xstoreCode, request("Stock_"&TRIM(n)), DATECREATED, request("QtyReceived_"&TRIM(n)), "RN")
										Else
											qryupdateStockLedger = cmdUpdateStockLedger(xstoreCode, request("Stock_"&TRIM(n)), DATECREATED, request("QtyReceived_"&TRIM(n)), "RN", "Insert")
										End If

										' ============================================================
										' TAMBAHAN: Update received di purchaseorderdetails saat New
										' ============================================================
										q = "Update purchaseorderdetails set received = received + '" & request("QtyReceived_"&TRIM(n)) & "' " & _
											" WHERE POCODE='" & xcode & "' AND LINENOKEY='" & request("LineNoKey_"&TRIM(n)) & "' AND FCBA='" & fcba & "'"
										connStr.execute(q)
										modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName
										executeQry	= "Query executed:" & q
										ExecuteAndLog modulFile, executeQry
										' ============================================================
										' AKHIR TAMBAHAN
										' ============================================================
									
									Else
										
										qLineMax = "select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
												   "where " & _
												   "fcba = '" & fcba & "'" 
										set rsqLineMax = Server.CreateObject("ADODB.Recordset")
										rsqLineMax.Open qLineMax, connstr
										if not rsqLineMax.eof then
											tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
										else
											tempMaxLine = 0
										end if 

										tempMaxLine = tempMaxLine + 1
										If  request("LT_"&TRIM(n))<>"" and request("LC_"&TRIM(n))<>"" and request("Stock_"&TRIM(n)) <>"" and txtControlJOb <> "" then
											q = "Insert into  RECEIVENOTEDETAIL (RECEIVENOTECODE, LOCATIONTYPE, LOCATIONCODE, PURCHASEITEMCODE, " & _
												"JOBCODE, QUANTITY, QTYRETURN, AMOUNTRECEIVED, AMOUNTPO_RECEIVE, AMOUNTSTO_RECEIVE, AMOUNTRETURN, FCENTRY, FCEDIT, FCIP, FCBA, linenohist, purchaselinenokey, LASTUPDATE, LASTTIME) " & _
												"values ('" & fcCode & "','" & request("LT_"&TRIM(n)) & "','" & request("LC_"&TRIM(n)) &  _
												"','" & request("Stock_"&TRIM(n)) & "','" & txtControlJOb & "','" & request("QtyReceived_"&TRIM(n)) & _
												"',0,'"& txtamountreceivedpo + txtamountreceivedsto & "','"& txtamountreceivedpo & "','"& txtamountreceivedsto & "',0," & _
												"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
												Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "',"& tempMaxLine & ",'" & request("LineNoKey_"&TRIM(n)) & "', sysdate, To_Char(sysdate,'HH24:MI'))"
												connStr.execute(q)
												modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
												executeQry	= "Query executed:" & q
												ExecuteAndLog modulFile,executeQry

										end if

										' ============================================================
										' TAMBAHAN: Update StockOnHand saat New (non-MATERIAL)
										' ============================================================
										qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
															 "where storecode = '" & xstoreCode & "' and " & _
															 "stockcode = '" & request("Stock_"&TRIM(n)) & "' and " & _
															 "fcba = '" & fcba & "'"

										set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
										rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
										if not rsTotalRecordOnHand.eof then
											tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
										else
											tempTotalRecordOnHand = 0
										end if
										rsTotalRecordOnHand.Close : set rsTotalRecordOnHand = Nothing

										If tempTotalRecordOnHand = 0 then
											qryinsertstockOnhand = cmdInsertStockOnHand(xstoreCode, request("Stock_"&TRIM(n)), request("QtyReceived_"&TRIM(n)), txtKolomtransaksi)
										Else
											qryUpdatestockOnhand = cmdUpdateStockOnHand(xstoreCode, request("Stock_"&TRIM(n)), request("QtyReceived_"&TRIM(n)), txtKolomtransaksi, "Insert")
										End If

										' ============================================================
										' TAMBAHAN: Update StockLedger saat New (non-MATERIAL)
										' ============================================================
										dim qryStockLedger_new2, rsStockLedger_new2, totalRecordStockLedger_new2
										qryStockLedger_new2 = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _
															  " WHERE STORECODE = '" & xstoreCode & "' and " & _
															  " ITEMCODE  = '" & request("Stock_"&TRIM(n)) & "' and " & _
															  " year  = to_char(to_date('" & DATECREATED & "','mm/dd/yyyy'),'yyyy') and " & _
															  " month = to_char(to_date('" & DATECREATED & "','mm/dd/yyyy'),'mm') and " & _
															  " FCBA  = '" & fcba & "'"

										set rsStockLedger_new2 = Server.CreateObject("ADODB.Recordset")
										rsStockLedger_new2.Open qryStockLedger_new2, connstr
										if not rsStockLedger_new2.eof then
											totalRecordStockLedger_new2 = cdbl(rsStockLedger_new2("totalRecordStockLedger"))
										else
											totalRecordStockLedger_new2 = 0
										end if
										rsStockLedger_new2.Close : set rsStockLedger_new2 = Nothing

										If totalRecordStockLedger_new2 = 0 then
											qryinsertstockledger = cmdInsertStockLedger(xstoreCode, request("Stock_"&TRIM(n)), DATECREATED, request("QtyReceived_"&TRIM(n)), "RN")
										Else
											qryupdateStockLedger = cmdUpdateStockLedger(xstoreCode, request("Stock_"&TRIM(n)), DATECREATED, request("QtyReceived_"&TRIM(n)), "RN", "Insert")
										End If

										' ============================================================
										' TAMBAHAN: Update received di purchaseorderdetails saat New (non-MATERIAL)
										' ============================================================
										q = "Update purchaseorderdetails set received = received + '" & request("QtyReceived_"&TRIM(n)) & "' " & _
											" WHERE POCODE='" & xcode & "' AND LINENOKEY='" & request("LineNoKey_"&TRIM(n)) & "' AND FCBA='" & fcba & "'"
										connStr.execute(q)
										modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName
										executeQry	= "Query executed:" & q
										ExecuteAndLog modulFile, executeQry
										' ============================================================
										' AKHIR TAMBAHAN
										' ============================================================

									End if
								
							End If	
						End If
					Next

					If  err.number=0 then
						connStr.CommitTrans	
					End if
												
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
				cmdView	= "True"
		else
				cmdView	= "True"
		end if								 
	
		else
			cmdView	= request("cmdView")		    
	End if
else
    err.number = 1001
	err.description = "Accounting Period  " & ACCOUNTINGPERIOD & "/" & ACCOUNTINGYEAR & "  LOCKED !!!"
	cmdView	= request("cmdView")
End if

	
 Dim rsVoucher,SudahClose
 set rsVoucher = server.CreateObject("Adodb.recordset")
 q="select a.*,trim(a.suppliercode)||' : '||trim(b.fcname) supplierDecription from RECEIVENOTE a INNER JOIN Supplier b on a.suppliercode=b.fccode and a.fcba=b.fcba  " & _
  " where RECEIVENOTECODE='" & fcCode & "' and " & _
  " a.fcBA='" & fcba & "'"
 SudahClose="N"
 rsVoucher.Open q, connstr
 postingflg = "0"
 if Not rsVoucher.Eof then
    DATECREATED = FormatDateTime(rsVoucher("RNDATE"))
	suppliercode = rsVoucher("supplierDecription")
	xcode=rsVoucher("pocode")
	stocktransno2 =  rsVoucher("StockOrderNo")
	selectedsupplier=rsVoucher("suppliercode")
	selectedPO = rsVoucher("pocode")
	selectedstogrn =  rsVoucher("StockOrderNo")
	selecteddogrn = rsVoucher("DELIVERYORDERCODE") 
	dono2 =  rsVoucher("DELIVERYORDERCODE")
	xstoreCode = rsVoucher("STORECODE")
	SudahClose="Y"
	postingflg = rsVoucher("postingflg")
	rowstate = rsVoucher("rowstateid")
 End if
 rsVoucher.Close
 Set rsVoucher = Nothing

Cek_IsAccClosing = StatusAccClosingTgl(DATECREATED,fcba) 
IF postingflg = "1" THEN 
   postingflgtxt = "YA"
ELSE
   postingflgtxt	= "TIDAK"
END IF
IF rowstate = "" OR rowstate = "Planned" THEN 
   rowstate = "Planned"
ELSEIF rowstate = "Review Accounting" THEN 
   rowstate = "Review Accounting"
ELSE
   rowstate	= "Approved"
END IF


Dim IsAvlSto : IsAvlSto = "N"
oSQLQuery = "  Select nvl(sum(s.do_quantity - s.dorcvqty),0) qtyOSTD  from STOCKTRANSFERORDERDTL s " & _
			"	Where s.do_quantity - s.dorcvqty > 0 " & _
			"	  and s.pocode = '" & xcode& "'  " 
			
set rs = server.CreateObject("Adodb.recordset")
rs.Open oSQLQuery, connstr
If Not rs.Eof then
    If cdbl(rs("qtyOSTD")) > 0 then
	    IsAvlSto = "Y"
	end if	
End if
rs.Close
Set rs = Nothing


Dim IsAvlLCO : IsAvlLCO = ""
oSQLQuery = "  Select incoterm1  from PURCHASEORDER s " & _
			"	Where  s.pocode = '" & xcode& "'  " 
			
set rs = server.CreateObject("Adodb.recordset")
rs.Open oSQLQuery, connstr
If Not rs.Eof then
	IsAvlLCO = rs("incoterm1")
End if
rs.Close
Set rs = Nothing



Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
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
oRecordSet.Close
Set oRecordSet = Nothing


dim isApproved,rsApproved
set rsApproved = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'RN' and at.fccode='508' " & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND '"& session("CodeName") &"' = '" & session("CodeName") & "'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & session("fcba") & "'"
rsApproved.Open q, connstr
if Not rsApproved.Eof then
	isApproved =  cdbl(to_nol(rsApproved("total")))
else
	isApproved = 0
End if
rsApproved.Close : set rsApproved = Nothing 



 %>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-download" aria-hidden="true"></i>Header Receive Note</div><% end if %>
 <form name="frmEntry" method="post" action="ReceiveNote.Asp">
	<input type="hidden" value="<%=rowstate%>" id="statusnya" />
	<table class="data" width="100%">
		<tr>
		<td align="left">GRN No</td>
		<%
		    if fcCode = "(NEW)" then
				q = "select MAX(cn.PREFIX) || LPAD(NVL(MAX(regexp_substr(grn.RECEIVENOTECODE, '[^/]+', 1, 4)),0000)+0001,4,0) DOCNUMBER " & _
					" from " & _
					" controlnumber cn " & _
					" left join receivenote grn on cn.fcba = grn.fcba and cn.accountingyear = to_number(to_char(grn.RNDATE,'yyyy')) " & _
					" where cn.fcba ='" & fcba & "' " & _
					" and cn.LU_NAME ='RECEIVENOTE' "&_
					" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yyyy')) " & _
					" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'mm')) "
				set rsGRNNo = server.CreateObject("Adodb.recordset")
				rsGRNNo.Open q, connstr
				if Not rsGRNNo.Eof then
					fcCode = rsGRNNo("DOCNUMBER")
				End if
			End if
			
		%>
		<td align="left"><input type="text" maxlength="50" id="fcCode" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">&nbsp;*</font></td></tr><tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker"  type="text" maxlength="10" name="DATECREATED" id ="DATECREATED" size="10" value="<%=DATECREATED%>" onchange="ValidationClosing();get_Allowentry();">
		<font color="#FF0000">&nbsp;*</font>
		</td></tr>
		<tr>
			<td align="left">Supplier</td>
			<td align="left">      
				<%if SudahClose="Y" then  
				%>
				<select name="suppliercode" id="suppliercode" ><option value="<%=suppliercode%>" Selected><%=suppliercode%></option></select>				
				<%else%>		
					<div id="supplierlist"></div>
					<input type="Hidden" id="selectedsupplier" name="selectedsupplier" value="<%=selectedsupplier%>">
				<% End if %>	  
			</td>
		</tr>
		<tr>
			<td align="left">PO Number</td>
			<td align="left">     
	            <%if SudahClose="Y" then  
				%>
				<select name="xcode" id="xcode" ><option value="<%=xcode%>" Selected><%=xcode%></option></select>				
				<%else%>		
				<div id="polist"></div>
				<input type="Hidden" id="selectedpo" name="selectedpo" value="<%=selectedpo%>">
			 <% End if %>	  
			</td>
		</tr>
		<tr>
			<td align="left">STO No</td>
			<td align="left" > 
			<%if SudahClose="Y" then  
					%>
				<select name="stocktransno2" id="stocktransno2" ><option value="<%=stocktransno2%>" Selected><%=stocktransno2%></option></select>
				
			<%else%>		
				<div id="stolist"></div>
			
				<input type="Hidden" id="selectedstogrn" name="selectedstogrn" value="<%=selectedstogrn%>">
			 <% End if %>	  
			</td>
		</tr>
		<tr>
			<td align="left">DO No</td>
			<td align="left" > 
			<%if SudahClose="Y" then  
			%>
				<select name="dono2" id="dono2"  ><option value="<%=dono2%>" Selected><%=dono2%></option></select>
			<%else	%>	
				<div id="dolist"></div>
			
				<input type="Hidden" id="selecteddogrn" name="selecteddogrn" value="<%=selecteddogrn%>">
				  
			 <% End if %>	  
			</td>
		</tr>
		
		<tr>
		<td align="left">Store Code</td>
		<td align="left">        
				<select name="xstoreCode" value="<%=xstoreCode%>">
				<option></option>
				<%
				qLov = " select fccode, fcname from store where fcba = '"& fcba &"' and activation = 'Y' and iscentral='Y' "
	   
				Set rsLov = Server.CreateObject("ADODB.Recordset")
				rsLov.Open qLov, connstr
				if not rsLov.eof then
					while not rsLov.eof
				%>
					<option value="<%=rsLov("fccode")%>" <% if cstr(""&rsLov("fccode")) = cstr(""&xstoreCode) then%> selected <%end if%>><%=rsLov("fccode")%> : <%=rsLov("fcname")%></option>
				<%
						rsLov.movenext
					wend
				end if
				%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
        </td>
		<tr>
			<td align="left" >Status/Posting</td>
			<td align="left"><%=rowstate & "/ " & postingflgtxt%></td>
		</tr>
		</tr>
		
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions"> 
				<input Type="hidden" Value="True" Name="cmdView">  
				<input Type="hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="hidden" Value="<%=request("txtModule")%>" Name="txtModule">  
				<input Type="hidden" Value="<%=postingflg%>" id="postingflg_">				
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  
				<a href="goods_received_note.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&grnno=<%=fcCode%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"></a>
				<% if rowstate = "Planned" OR rowstate = "" Or rowstate = "Review Accounting" then %>
					<a id="editbtn" href="javascript:EditHeader();"><img src="IMAGES/edit.gif" alt="Edit Header : <%=fccode%>" border="0" onClick="return confirm('Are you sure to update header <%=fccode%>');"></a>
					<a id="delhdr" href="ReceiveNote.Asp?fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete <%=fccode%> ');"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				<% if rowstate = "Planned" Or rowstate = "Review Accounting" then %>
					<%If isApproved <> 0 Then %>
							<a id="poshdr" href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&fcba=<%=fcba%>&txtJenis=Approved" onClick="return confirm('Are you sure to Approved Receive Note this data :  <%=fcCode%> ?');">Approved</a>
					<%Else%>
							
					<%End If%>
				<%End If%>

				<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Purchase"
		dataSubmodule 	= "Receive Note"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>&nbsp;&nbsp;<span id="msg" >-</span>
		</td> 
		</tr> 
    </table>    
 </form>
<%
If cmdView = "True" and LEN(TRIM(fcCode)) > 4 and len(trim(xstoreCode)) > 0 Then
	Dim rsGRN,intRecordsPerPage,cColor
	Dim Menu : Menu = "N"
	Dim oRecordSetMenu, oSQLQueryMenu
	oSQLQueryMenu = " SELECT 'Y' istrue FROM users_tab " & _
				" WHERE fccode = '" & Session("CodeName") & "' AND MENU_GRN = 'TRUE'" 
	set oRecordSetMenu = server.CreateObject("Adodb.recordset")
	oRecordSetMenu.Open oSQLQueryMenu, connstr
	If Not oRecordSetMenu.Eof then
		Menu = oRecordSetMenu("istrue")
	Else
		Menu = "N"
	End if

	set rsGRN = server.CreateObject("Adodb.recordset")
	If SudahClose = "Y" Then
		q	= " select purchaseitemcode  " & _
		      "      , s.fcname itemname " & _
		      "      , purchaselinenokey " & _
			  "      , linenohist " & _
			  "      , jobcode " & _
			  "      , locationtype " & _
			  "      , locationcode " & _
			  "      , nvl(sum(nvl(quantity, 0)), 0) as quantity " & _
			  "      , nvl(sum(nvl(amountreceived, 0)), 0) as amountreceived " & _			  
			  " from receivenotedetail p " & _
			  "      , stock s " & _
			  " where (p.purchaseitemcode = s.fccode and p.fcba = s.fcba) " & _
			  "       and receivenotecode = '" & fcCode & "'" & _
			  "       and p.fcBA = '" & fcba & "'" & _
			  " group by purchaseitemcode" & _
			  "        , s.fcname" & _
			  "        , purchaselinenokey" & _
			  "        , linenohist" & _
			  "        , jobcode" & _
			  "        , locationtype" & _
			  "        , locationcode" & _
			  " order by purchaseitemcode" & _
			  "        , purchaselinenokey"
	end if	

    If 	IsAvlLCO = "LCO" and stocktransno2 = ""  then
	    Response.Write " This PO Must Be Stock Transfer Order ( STO )"
		q	= " select * from purchaseorderdetails p" & _	
		      " where pocode='TidakAda' "
	Else
		If SudahClose = "N" AND IsAvlSto ="N" then	 
				q	= " select p.itemcode,s.fcname,p.linenokey purchaselinenokey,p.quantity quantityorder,p.received quantityreceived,(p.quantity-p.received) as sisa, " & _
				" stockgroup_api.get_controljob(s.stockgroupcode,s.fcba) jobcode,'ST' LT,'" & xstoreCode & "' LC " & _ 
				" from purchaseorderdetails p inner join stock s on  p.itemcode = s.fccode and p.fcba = s.fcba  " & _
				" where " & _
				" p.pocode = '" & xcode & "' " & _
				" and p.fcba = '" & fcba & "' " & _
				" and quantity-p.received>0"
		 Elseif SudahClose = "N" AND IsAvlSto ="Y" then	 
				q = "   SELECT p.itemcode," & _
						"          stock_api.get_stockName (p.itemcode, p.fcba) fcname," & _
						"          p.linenokey purchaselinenokey," & _
						"          NVL (SUM (NVL (p.quantity, 0)), 0) quantityorder, NVL (SUM (NVL (p.quantity, 0)), 0) - NVL (SUM (NVL (p.received, 0)), 0) AS quantityreceived," & _
						"          SUM (s.do_quantity) doqty," & _
						"          SUM (s.do_quantity) - SUM (s.dorcvqty) doreceiveqty," & _
						"          ROUND (( (NVL (SUM (NVL (p.amount, 0)), 0) / SUM (p.quantity)) * SUM (s.do_quantity)) + ( (SUM (NVL (s.amount, 0)) / SUM (p.quantity) ) * SUM (s.do_quantity))) AS amount" & _
						"     FROM purchaseorderdetails p" & _
						"          RIGHT JOIN STOCKTRANSFERORDERDTL s" & _
						"             ON (    p.pocode = s.pocode" & _
						"                 AND p.linenokey = s.linenokey" & _
						"                 AND p.fcba = s.fcba)" & _
						"    WHERE     p.pocode = '" & xcode& "'" & _
						"          AND s.stono = '" & stocktransno2 & "' " & _
						"          AND s.do_quantity - s.dorcvqty > 0" & _
						"          AND p.fcba = '"& fcba &"'" & _
						" GROUP BY p.itemcode, p.fcba, p.linenokey"
		End if
	End If		
   Set rsGRN = Server.CreateObject("ADODB.Recordset")
   rsGRN.CursorType = 3
   rsGRN.CursorLocation = 3
   rsGRN.Open q, connstr%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="ReceiveNote.Asp" method="post" name="eForm" >
		<div style="overflow-x: auto">
		
		<table class="data" border="1" width="100%" >
		<%
			if SudahClose<>"Y" AND IsAvlSto ="Y" Then%>
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
			<td align="center">No</td>
			<td align="center">LineNoKey</td>
			<td align="center">Code</td>
			<td align="center">Desc.</td>
			<td align="center">Qty PO</td>
			<td align="center">Sisa Qty PO</td>
			<td align="center">Qty DO</td>
			<td align="center">Sisa Qty DO</td>
			<%If Menu = "Y" Then%>
			<td align="center">Amount</td>
			<%End If%>
			<td align="center">Job Code</td>
			<td align="center">Target Type</td>
			<td align="center">Target Code</td>
			<td align="center">Quantity</td>
			</tr>
		<%ElseIf  SudahClose<>"Y" AND IsAvlSto ="N" Then	%>
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
			<td align="center">No</td>
			<td align="center">LineNoKey</td>
			<td align="center">Code</td>
			<td align="center">Desc.</td>
			<td align="center">Order (Qty)</td>
			<td align="center">Receive (Qty)</td>
			<td align="center">Sisa order</td>
			<td align="center">Job Code</td>
			<td align="center">Target Type</td>
			<td align="center">Target Code</td>
			<td align="center">Quantity</td>
			</tr>
			<%Else	%>
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
			<td align="center">No</td>
			<td align="center">&nbsp;</td>
			<td align="center">Stock Code</td>
			<td align="center">Stock Name</td>
			<td align="center">LineNoKey</td>
			<td align="center">Line No Hist</td>
			<td align="center">Control Job</td>
			<td align="center">Target Type</td>
			<td align="center">Target Code</td>
			<td align="center">Quantity</td>
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
						<td align="left" valign="top"><%=NO%></td>
						<%IF SudahClose = "Y" then%>
							<% if rowstate = "Planned" OR rowstate = "" then %>
								<td valign="middle">
									<a  id="deldtl" href="ReceiveNote.Asp?fcCode=<%=fcCode%>&ItemCode=<%=rsGRN("PURCHASEITEMCODE")%>&purchaseLineNoKey=<%=rsGRN("PurchaseLineNoKey")%>&linenomorhist=<%=rsGRN("linenohist")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details">
									<img id="deldtl" src="IMAGES/delete.gif" alt="Delete Data : <%=rsGRN(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
								</td>
							<%Else%>
								<td valign="middle">
									<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
								</td>
							<%End If%>
												
							<%for i = 0 to nCountofField-1%>
								<td align="left" title="<%=rsGRN(i).Name%>" valign="top"><%=rsGRN(i)%></td>
							<%next%>
						<%End if%>
						
						<%if SudahClose<>"Y"  AND IsAvlSto ="N"  Then%>
							<td align="Right" title="LineNoKey" valign="top"><%=rsGRN("purchaselinenokey")%></td>
							<td align="left" title="Item Code" valign="top"><%=rsGRN("itemcode")%></td>
							<td align="left" title="Item Name"  style="width: 200px;" valign="top"><%=rsGRN("fcname")%></td>
							<td align="Right" title="QuantityOrder" valign="top"><%=FbpNum2(rsGRN("quantityorder"))%></td>
							<td align="Right" title="QuantityReceived"  style="width: 40px;" valign="top"><%=FbpNum2(rsGRN("quantityreceived"))%></td>
							<td align="Right" title="sisa"  style="width: 40px;" valign="top"><%=FbpNum2(rsGRN("sisa"))%></td>
							<td align="left" width="2"><input type="text" style="width: 100px;" name="Jobcode_<%=no%>" width="2" value="<%=rsGRN("Jobcode")%>"></td>
							<td align="left" width="2"><input type="text" style="width: 100px;" name="LT_<%=no%>" width="2" value="<%=rsGRN("LT")%>"></td>
							<td align="left" width="10"><input type="text" style="width: 100px;" name="LC_<%=no%>" width="10" value="<%=rsGRN("LC")%>"></td>
							<td align="right" width="10"><input type="text" name="QtyReceived_<%=no%>" id="QtyReceived_<%=no%>" style="width: 40px;" value="<%=request("QtyReceived_"&TRIM(no))%>" onkeyup='getValidasiQty("<%=no%>")'  onkeypress="return harusangka(event);" >
							<input type="hidden" name="AmountReceived_<%=no%>" width="10" value="<%=txtamountreceivedpo%>">
							<input type="hidden" name="LineNoKey_<%=no%>" width="10" value="<%=rsGRN("PurchaseLineNoKey")%>">
							<input type="hidden" name="Stock_<%=no%>" value="<%=rsGRN("ITEMCODE")%>">
							<input type="hidden" name="QuantityOrder_<%=no%>" id="QuantityOrder_<%=no%>" value="<%=rsGRN("quantityorder")%>"></td>
						<%elseif SudahClose<>"Y"  AND IsAvlSto ="Y"  Then%>
							<td align="Right" title="LineNoKey" valign="top"><%=rsGRN("purchaselinenokey")%></td>
							<td align="left" title="Item Code" valign="top"><%=rsGRN("itemcode")%></td>
							<td align="left" title="Item Name" style="width: 200px;" valign="top"><%=rsGRN("fcname")%></td>
							<td align="Right" title="QuantityOrder" valign="top"><%=FbpNum2(rsGRN("quantityorder"))%></td>
							<td align="Right" title="QuantityReceived" valign="top"><%=FbpNum2(rsGRN("quantityreceived"))%></td>
							<td align="Right" title="QuantityReceived" valign="top"><%=FbpNum2(rsGRN("doqty"))%></td>
							<td align="Right" title="QuantityReceived" style="width: 40px;" valign="top"><%=FbpNum2(rsGRN("doreceiveqty"))%></td>
							<td align="left" width="2"><input type="text" style="width: 100px;" name="Jobcode_<%=no%>" width="2" value="<%=request("Jobcode_"&TRIM(no))%>"><a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a></td>
							<td align="left" width="2"><input type="text" style="width: 100px;" name="LT_<%=no%>" width="2" value="<%=request("LT_"&TRIM(no))%>"<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Target Type</font></a><font color="#FF0000">&nbsp;*</font></td>
							<td align="left" width="10"><input type="text" style="width: 100px;" name="LC_<%=no%>" width="10" value="<%=request("LC_"&TRIM(no))%>"></td>
							<td align="right" width="10"><input type="text"  style="width: 40px;" name="QtyReceived_<%=no%>" id="QtyReceived_<%=no%>" width="10" value="<%=request("QtyReceived_"&TRIM(no))%>" onkeyup='getValidasiQty("<%=no%>")' onkeypress="return harusangka(event);" >
							<input type="hidden" name="AmountReceived_<%=no%>" width="10" value="<%=txtamountreceived%>">
							<input type="hidden" name="LineNoKey_<%=no%>" width="10" value="<%=rsGRN("PurchaseLineNoKey")%>">
							<input type="hidden" name="Stock_<%=no%>" value="<%=rsGRN("ITEMCODE")%>">
							<input type="hidden" name="QuantityOrder_<%=no%>" id="QuantityOrder_<%=no%>" value="<%=rsGRN("quantityorder")%>">
							</td>
						<%End if%>
					</tr>
				<%
				rsGRN.MoveNext
				intRecordsPerPage = intRecordsPerPage - 1
				no=no+1
			Wend
		else
			response.Write("<tr><td colspan=""13"" class=""fbp-doc-empty"">Data is empty</td></tr>")	  
		end if 
		rsGRN.Close
		set rsGRN = Nothing%>
			<tr>
				<td colspan="13" align="left">
				<input type="hidden" name="fcCode" value="<%=fcCode%>">
				<input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
				<input type="hidden" id="_RGDate" value="<%=DATECREATED%>">
				<input type="hidden" name="xcode" value="<%=xcode%>">
				<input type="hidden" name="stocktransno2" value="<%=stocktransno2%>">
				<input type="hidden" name="dono2" value="<%=dono2%>">
				<input type="hidden" name="xstoreCode" value="<%=xstoreCode%>">
				<input type="hidden" name="No" value="<%=No%>">
				<input type="hidden" name="selecteddogrn" value="<%=selecteddogrn%>">
				<input type="hidden" name="selectedstogrn" value="<%=selectedstogrn%>">
				<input type="hidden" name="selectedpo" value="<%=selectedpo%>">
				<input type="hidden" name="selectedsupplier" value="<%=selectedsupplier%>">
				
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<% if postingflg ="0" and Cek_IsAccClosing="FALSE" then%>
				<%IF Request("txtJenis")="Edit" then%>
					<input type="hidden" name="txtJenis" value="Edit">
					<input Type="Submit"  name="Add" value="Update" id="add">
				<%Else%>
					<input type="hidden" name="txtJenis" value="New">
					<input type="submit" name="Add" value="Submit" id="add" onclick="return validateForm(event);">
				<%End if%>
				<%End if%>
				
				<input Type="Reset"  name="Reset" value="Reset">
				<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				
			</tr>
			
		</TABLE>
		<div style="overflow-x: auto; margin-top: 30px">
			<h3>HISTORI RECEIVE NOTE BERDASARKAN NOMOR PO</h3>
			<%
				sqlhist = " select " & _  
						" a.lineno, a.fcba, a.stockcode, b.FCNAME, a.transdate, a.transtype, a.storecode, a.sourcetype, a.sourcecode, a.targettype, a.targetcode, a.quantity, a.documentno " & _
						" from StockTransHist a " & _
						" left join stock b on a.STOCKCODE = b.fccode " & _
						" where sourcecode = '" & xcode& "' " & _
						" group by a.lineno, a.fcba, a.stockcode, b.FCNAME, a.transdate, a.transtype, a.storecode, a.sourcetype, a.sourcecode, a.targettype, a.targetcode, a.quantity, a.documentno " & _
						" order by a.quantity, a.transdate desc "
				Set rsHist = Server.CreateObject("ADODB.Recordset")
				rsHist.CursorType = 3
				rsHist.CursorLocation = 3
				rsHist.Open sqlhist, connStr
			%>

			<table class="data" border="1" width="100%">
				<thead>
					<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
						<td align="center">Line No</td>
						<td align="center">FCBA</td>
						<td align="center">ITEM CODE</td>
						<td align="center">ITEM NAME</td>
						<td align="center">TRANS DATE</td>
						<td align="center">TRANS TYPE</td>
						<td align="center">STORE CODE</td>
						<td align="center">SOURCE TYPE</td>
						<td align="center">DOC NO</td>
						<td align="center">GRN NO</td>
						<td align="center">TARGET TYPE</td>
						<td align="center">TARGET CODE</td>
						<td align="center">QUANTITY</td>
					</tr>
				</thead>
				<tbody>
					<% If Not rsHist.EOF Then
						Do While Not rsHist.EOF
					%>
						<tr>
							<td align="center"><%=rsHist("LINENO")%></td>
							<td align="center"><%=rsHist("fcba")%></td>
							<td align="center"><%=rsHist("stockcode")%></td>
							<td align="center"><%=rsHist("FCNAME")%></td>
							<td align="center"><%=rsHist("transdate")%></td>
							<td align="center"><%=rsHist("transtype")%></td>
							<td align="center"><%=rsHist("storecode")%></td>
							<td align="center"><%=rsHist("sourcetype")%></td>
							<td align="center"><%=rsHist("sourcecode")%></td>
							<td align="center"><%=rsHist("documentno")%></td>
							<td align="center"><%=rsHist("targettype")%></td>
							<td align="center"><%=rsHist("targetcode")%></td>
							<td align="center"><%=rsHist("quantity")%></td>
						</tr>
					<%
						rsHist.MoveNext
						Loop
					Else
					%>
						<tr>
							<td colspan="13" align="center">Tidak ada data</td>
						</tr>
					<% End If %>
				</tbody>
			</table>
		</div>
   	</div>
   </form>
<%
End if  
%>
 </div> </div>
  <!--#include file="bottom.jsp"-->
<script src="/js/jquery.inputmask.bundle.min.js"></script>
	<script src="/js/inputmask.binding.js"></script>
  	<script src="plugins/chosen/chosen.jquery.min.js"></script>
	<script>
		function validateForm(e) {
			if (e) e.preventDefault();

			let selectedpo = document.querySelector('[name="xcode"]').value;
			let isValid = false;

			$.ajax({
				url: "bkm_viewdata.asp?idview=42&pocode=" + encodeURIComponent(selectedpo),
				method: "GET",
				dataType: "json",
				async: false,
				success: function(data) {
					console.log("Server Response:", data);
					if (data.count > 0) {
						alert(data.message);
						isValid = false;
					} else {
						isValid = true;
					}
				},
				error: function(xhr, status, err) {
					alert("Terjadi kesalahan saat validasi data.");
					isValid = false;
				}
			});

			if (isValid) {
				document.forms["eForm"].submit();
			}

			return false;
		}
	</script>

	<script type="text/javascript">	
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
	</script>
 	<script type="text/javascript">
		function handleClick(cb){
			if(cb.checked == 1){
				cb.value = "Y";
			} else {
				cb.value = "N";
			}
        }
	</script>

	<script>

	function getselectedsupplier() {
		var sselectedsupplier = document.getElementById("selectedsupplier");
		sselectedsupplier.value = document.getElementById("suppliercode").value;
	}
	
	function getselectedpo() {
		var sselectedpo = document.getElementById("selectedpo");
		sselectedpo.value = document.getElementById("xcode").value;
	}

	
	function getselectedstogrn() {
		var sfcba = document.getElementById("selectedstogrn");
		sfcba.value = document.getElementById("stocktransno2").value;
	}

	function getselecteddo() {
		var scustomercode = document.getElementById("selecteddogrn");
		scustomercode.value = document.getElementById("dono2").value;
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

	document.ready(function() {
	    getSupplierRG();
		getPOgrn();
		getliststo();
		getlistdo();
		getselectedpo();
		getselectedstogrn();
		getselecteddo();
		
	});
	</script>
	
	<script>
$(function() {	

	ValidationClosing();
	get_Allowentry();

	
});

function ValidationClosing()
{
	var IssueDate = $('#DATECREATED').val();
	var postingflg =  $('#postingflg_').val();
	if (IssueDate != "") {
		$.post("acc_allow_entry.asp?tglbkm=" + IssueDate, function(data) {
			$('#msg').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			
			if(closingStatus=="FALSE" && $('#statusnya').val() != "Review Accounting"){
				$('#msg').html('<label color:green>OPEN PERIODE </label>')
				if (postingflg == 1) 
					{
					$('#delhdr').hide();
					$('#poshdr').hide();
					$('#editbtn').hide();
					$('#deldtl').hide();
					$(':input[type="text"]').prop('disabled', true); 
					$(':input[type="Submit"]').prop('disabled', true);
					}
					else
					{
					$('#delhdr').show();
					$('#poshdr').show();
					$('#editbtn').show();
					$('#deldtl').show();
					
					}
				
			}elseif(closingStatus=="FALSE" && $('#statusnya').val() == "Review Accounting"){
				$('#msg').html('<label color:green>OPEN PERIODE </label>')
				if (postingflg == 1) 
					{
					$('#delhdr').hide();
					$('#poshdr').hide();
					$('#editbtn').hide();
					$('#deldtl').hide();
					$(':input[type="text"]').prop('disabled', true); 
					$(':input[type="Submit"]').prop('disabled', true);
					}
					else
					{
					$('#delhdr').show();
					$('#poshdr').show();
					$('#editbtn').show();
					$('#deldtl').show();
					
					}
				
			}else{
				alert('PERIODE SUDAH DI CLOSE !!!');
				$('#msg').html('<label color:red>PERIODE SUDAH DI CLOSE !!!</label>')
				$('#delhdr').hide();
				$('#poshdr').hide();
				$('#editbtn').hide();
				$('#deldtl').hide();
				$(':input[type="text"]').prop('disabled', true); 
				$(':input[type="Submit"]').prop('disabled', true);
				
			}
						
		});
		
	} 
}

function get_Allowentry()
{
	var tglbkm = $('#DATECREATED').val();
	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE"){
				$('#add').show();
				$('#deldtl').show();
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$('#add').hide();
				$('#deldtl').hide();
				
			}
			
		});
		
	} 
}

</script>
</body>
</html>
