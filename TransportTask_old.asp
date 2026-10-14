<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="qtyOnHandFunction.inc"-->
<html>
<head>
<title>FBPlants - Transport Task</title>
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
var getPrice = function () { 
	// var fccompanycode = document.getElementById("fccompanycode").value;
	
	if ($("#priceby").val() === "OPB"){
		var orderdate = document.getElementById("orderdate").value; 
		var itemcode = document.getElementById("itemcode").value; 
		var storecode = document.getElementById("storecode").value; 
		var priceby = document.getElementById("priceby").value; 
		var PostStr = 'orderdate='+orderdate+'&storecode='+storecode+'&itemcode='+itemcode+'&priceby='+priceby;
		// alert(PostStr); 
		// use the generic function to make the request
		doAJAXCall('ajax/getAmountTransporttask.asp', 'POST', PostStr, showMessageResponse);
	}else if ($("#priceby").val() === "GRN"){
		// var fccompanycode = document.getElementById("fccompanycode").value;
		var lineby = document.getElementById("lineby").value; 
		var documentno = document.getElementById("documentno").value;  
		var itemcode = document.getElementById("itemcode").value;  
		var storecode = document.getElementById("storecode").value;
		var priceby = document.getElementById("priceby").value; 
		var PostStr = 'lineby='+lineby+'&documentno='+documentno+'&itemcode='+itemcode+'&storecode='+storecode+'&priceby='+priceby;
		// alert(PostStr); 
		// use the generic function to make the request
		doAJAXCall('ajax/getpricegrntt.asp', 'POST', PostStr, showMessageResponse);
	}else if ($("#priceby").val() === "RCT"){
		// var fccompanycode = document.getElementById("fccompanycode").value;
		var lineby = document.getElementById("lineby").value; 
		var documentno = document.getElementById("documentno").value;  
		var itemcode = document.getElementById("itemcode").value;  
		var storecode = document.getElementById("storecode").value;
		var priceby = document.getElementById("priceby").value; 
		var PostStr = 'lineby='+lineby+'&documentno='+documentno+'&itemcode='+itemcode+'&storecode='+storecode+'&priceby='+priceby;
		// alert(PostStr); 
		// use the generic function to make the request
		doAJAXCall('ajax/getpricegrntt.asp', 'POST', PostStr, showMessageResponse);
	}else {
		var lineby = document.getElementById("lineby").value; 
		var documentno = document.getElementById("documentno").value;  
		var itemcode = document.getElementById("itemcode").value;  
		var storecode = document.getElementById("storecode").value;
		var priceby = document.getElementById("priceby").value; 
		var PostStr = 'lineby='+lineby+'&documentno='+documentno+'&itemcode='+itemcode+'&storecode='+storecode+'&priceby='+priceby;
		// alert(PostStr); 
		// use the generic function to make the request
		doAJAXCall('ajax/getpricegrntt.asp', 'POST', PostStr, showMessageResponse);
	}
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) {  
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    //alert(s.price);
	
    // update the Div to show the result from the server
    //$("#price").val() = 0000;
	document.getElementById("price").value = s.price;
	// document.getElementById("stockname").value = s.stockname;
	// document.getElementById("requestedqty").value = s.approvedqty; 
}




</script>
<script language="javascript" type="text/javascript"> 

    function showHint(str){
		var xmlhttp;
		var StockName;
		var stockitem;

		if (str==""){
		  document.getElementById("announ").innerHTML="";
		  document.getElementById("StockName").innerHTML="";
		  document.getElementById("stockitem").innerHTML="";
		  return;
		}
		if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
		  xmlhttp=new XMLHttpRequest();
		  StockName=new XMLHttpRequest();
		}else{// code for IE6, IE5
		  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
		  StockName=new ActiveXObject("Microsoft.XMLHTTP");
		  stockitem=new ActiveXObject("Microsoft.XMLHTTP");
		  
		}
		xmlhttp.onreadystatechange=function(){
		  	if (xmlhttp.readyState==4 && xmlhttp.status==200){
				document.getElementById("announ").innerHTML=xmlhttp.responseText;
			}
		}
		xmlhttp.open("GET","ajax/qualityAjax.asp?ajaxUoM=quality&q="+str,true);
		xmlhttp.send();
		
		StockName.onreadystatechange=function(){
			if (StockName.readyState==4 && StockName.status==200){
				document.getElementById("StockName").innerHTML = StockName.responseText;
			}
		}

		StockName.open("GET","ajax/qualityAjax.asp?ajaxUoM=StockName&q="+str,true);
		StockName.send();
			
	}

	function showHintStock(str){
		var xmlhttp;
		var StokBarang;

		if (str==""){
		  document.getElementById("announ").innerHTML="";
		  document.getElementById("StokBarang").innerHTML="";
		  return;
		}

		if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
		  xmlhttp=new XMLHttpRequest();
		  StokBarang=new XMLHttpRequest();
		}else{// code for IE6, IE5
		  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
		  StokBarang=new ActiveXObject("Microsoft.XMLHTTP");
		  
		}
		
		let getDate = $('#orderdate').val();
		let setDet  = new Date(getDate)
		let bln 	= setDet.getMonth() + 1;
		let thn 	= setDet.getFullYear();

		StokBarang.onreadystatechange = function(){
			if (StokBarang.readyState == 4 && stockitem.status == 200){
				document.getElementById("StokBarang").innerHTML = StokBarang.responseText;
			}
		}

		StokBarang.open("GET","ajax/qualityAjax.asp?ajaxUoM=StockItem&q="+str+"&bln="+bln+"&thn="+thn,true);
		StokBarang.send();
			
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

function getLinenoGRN(str){

	var xmlhttp;
	if (str=="")
	  {
	  document.getElementById("orderdate").innerHTML="";
	  document.getElementById("documentno").innerHTML="";
	  document.getElementById("itemcode").innerHTML="";
	  document.getElementById("storecode").innerHTML="";
	  // document.getElementById("receviedtask").innerHTML="";
	  return;
	  }
	if (window.XMLHttpRequest)
	  {// code for IE7+, Firefox, Chrome, Opera, Safari
	  xmlhttp=new XMLHttpRequest();
	  StockName=new XMLHttpRequest();
	  }
	else
	  {xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");}
	xmlhttp.onreadystatechange=function()
	  {
	  if (xmlhttp.readyState==4 && xmlhttp.status==200)
		{
		// alert(xmlhttp.responseText);
		document.getElementById("lineby").innerHTML=xmlhttp.responseText;
		// document.getElementById("requestedqty").innerHTML=xmlhttp.responseText;
		// document.getElementById("receviedtask").innerHTML="";
		// alert(PRLINEID);
		}
	  }
	   if ($("#priceby").val() === "RCT"){
			xmlhttp.open("GET","ajax/getLinenoRCT.asp?documentno="+str+"&itemcode="+document.getElementById("itemcode").value+"&storecode="+document.getElementById("storecode").value+"&orderdate="+document.getElementById("orderdate").value+"",true);
			xmlhttp.send();
		} else if ($("#priceby").val() === "GRN"){
			// xmlhttp.open("GET","ajax/getSupplierPO.asp?(lihat cara pemanggilan di ajax)="+str+"",true);  
			xmlhttp.open("GET","ajax/getlinenoamounttransporttask.asp?documentno="+str+"&itemcode="+document.getElementById("itemcode").value+"&storecode="+document.getElementById("storecode").value+"&orderdate="+document.getElementById("orderdate").value+"",true);
			xmlhttp.send();
		} else {
			xmlhttp.open("GET","ajax/getLinenoWRO.asp?documentno="+str+"&itemcode="+document.getElementById("itemcode").value+"&storecode="+document.getElementById("storecode").value+"&orderdate="+document.getElementById("orderdate").value+"",true);
			xmlhttp.send();
		}
	// alert(xmlhttp.open("GET","ajax/getSupplierPO.asp?(lihat cara pemanggilan di ajax)="+str+"",true));
}


					
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);

</script>
<script>
	function hitungrumus() {
		var squantity 	= parseFloat($("#quantity").val().replace(/,/g, ''));
		var sprice 		= $("#price").val();
		var nquantity 	= squantity;
		var nprice 		= sprice.replace(/,/g, '');
		var stok 		= parseFloat($('#StokBarang').text().replace(/,/g, ''));

		var cekInput 	= $('#Add').val();

		console.log('INPUT : ', squantity);
		console.log('STOK : ', stok);
		console.log('HASIL : ', stok - squantity);

		if(cekInput == "Select" && squantity > 0 && stok > 0){
			if (squantity > stok) {
				alert("Jumlah melebihi stok (" + stok + ")");
				$("#quantity").val(0);
			}else{
				//Untuk menghitung Amount 
				totalamount = nquantity * nprice;
				if (!isNaN(totalamount)) {
					$("#amount").val(totalamount.toFixed(2));
				}
				else{
					$("#amount").val();
				}
			}
		}else if(cekInput == "Select" && squantity > 0 && stok == 0){
			$("#quantity").val(0);
		}
	}
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<script>
function validateFloatKeyPress(el, evt) {
    var charCode = (evt.which) ? evt.which : event.keyCode;
    var number = el.value.split('.');
    if (charCode != 46 && charCode > 31 && (charCode < 48 || charCode > 57)) {
        return false;
    }
    //just one dot
    if(number.length>1 && charCode == 46){
         return false;
    }
    //get the carat position
    var caratPos = getSelectionStart(el);
    var dotPos = el.value.indexOf(".");
    if( caratPos > dotPos && dotPos>-1 && (number[1].length > 1)){
        return false;
    }
    return true;
}

//thanks: http://javascript.nwbox.com/cursor_position/
function getSelectionStart(o) {
	if (o.createTextRange) {
		var r = document.selection.createRange().duplicate()
		r.moveEnd('character', o.value.length)
		if (r.text == '') return o.value.length
		return o.value.lastIndexOf(r.text)
	} else return o.selectionStart
}
</script>
<script language="JavaScript"> 
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

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
 Dim STORECODE,OrderNo,cmdView,PoCode,OrderDate,ITEMCODE,QuantityTransfered,Price,AMOUNT, LineNoHist,fcba
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim Jobcode, LocationType, LocationCode, Remarks, Reference
 Dim qLineMax, rsqLineMax, tempMaxLine
 Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
 Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand, QuantityOriginal
 Dim done, qq  
 Dim qLessZero, rsLessZero, tempislesszero
 Dim qLineUpper, rsqLineUpper, totalUpper
 Dim qQuantity, rsqQuantity, totalQuantity,LineNo
 
 Dim rsTransportTaskDetail
 Dim v_CrossCekDetail
 Dim qDeleteHist
 Dim strLineMax, rsstrLineMax, tempMaxLineNo 
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim Cek_IsAccClosing : Cek_IsAccClosing = ""
 Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD
 Dim FcbaDistination,QuantityReceived
 Dim CompanyCode,rsCompany,rsRealPrice,realprice,realAmount
 Dim txtLineNoHist, originalQuantity,txtItemCodeDelete, txtLineNo, txtStoreCode
 dim txtKolomtransaksi,qryUpdatestockOnhand

 Dim detail_remarks : detail_remarks = request("detail_remarks")
 dim ScriptName,modulFile,executeQry
 ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
txtKolomtransaksi = "TRT"

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 ' if request("OrderNo")="" then
    ' OrderNo="" '----- Nantinya Ambil dari Format Master
 ' else
    ' OrderNo = request("OrderNo")
 ' end if
 
  OrderNo = request("OrderNo")
 If Len(Trim(OrderNo)) = 0 Then OrderNo = "(NEW)"
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("OrderDate")<>"" then
   OrderDate   =Request("OrderDate")
 else
   OrderDate   =textDate(date)
 end if
 
 if KonversiTanggalKeDbl(OrderDate) > KonversiTanggalKeDbl(textDate(date)) then
	OrderDate = textDate(date)
 end if
 
 if Request("fcba") <> "" then
	fcba = Request("fcba")
 else	
	fcba = session("fcba") 
 end if
 
 ACCOUNTINGPERIOD  = left(OrderDate,2)
 ACCOUNTINGYEAR    = right(OrderDate,4)
 
 '---- Data Details
 JobCode=Request("JobCode")
 LocationType="CB"
 LocationCode="BS"
 Remarks = Request("Remarks")
 Reference = Request("Reference")
 ITEMCODE=Request("ITEMCODE")
 STORECODE = Request("STORECODE")
 LineNo=Request("LineNo")
 QuantityTransfered=Request("QuantityTransfered")
 QuantityReceived=Request("QuantityReceived")
 Price = Request("Price")
 AMOUNT =Request("AMOUNT")
 Linenohist = Request("linenohist")
 FcbaDistination = Request("FcbaDistination")
 rowstateid = Request("rowstateid")
 document_classification = "206"
 priceby = Request("priceby")
 documentno = Request("documentno")
 lineby = Request("lineby")
 
 'response.write JobCode & "-" & Linenohist
                
 
 ' Get Companycode  add by Akbar 5 Sept 2016
 
 q = "select t.fccompanycode from Businessunit t  where t.fccode='" & fcba & "'"
 
 Set rsCompany = Server.CreateObject("ADODB.Recordset")
 rsCompany.Open q, connstr
 IF Not rsCompany.eof then
	CompanyCode = rsCompany("fccompanycode")
 Else
    CompanyCode = "NONE"
 End if
 rsCompany.Close
 set rsCompany = Nothing	
 
 
  IF AMOUNT= "" or AMOUNT="0" then
	 AMOUNT = 0
	 PRICE = 0
  END if

Dim rsTRT,postingflgtxt
 set rsTRT = server.CreateObject("Adodb.recordset")
 q="select to_char(ORDERDATE,'mm/dd/yyyy') ORDERDATE,postingflg from TRANSPORTTASK_TAB  " & _
  " where ORDERNO='" & OrderNo & "' and " & _
  " fcBA='" & fcba & "'"

 rsTRT.Open q, connstr

 if Not rsTRT.Eof then
    OrderDate=rsTRT("ORDERDATE")
	postingflg = rsTRT("postingflg")
 End if

 IF postingflg = "1" THEN 
   postingflgtxt = "YA"
ELSE
   postingflgtxt	= "TIDAK"
END IF
 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
Cek_IsAccClosing =  StatusAccClosingTgl(OrderDate,fcba)  
if 	Cek_IsAccClosing = "FALSE" then  
 
 	if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "StartAuthorize" or Request("txtJenis")= "Posting" or Request("txtJenis")= "Approved" Then

		if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(fcba, OrderDate)
			If Cek_IsDailyClosing = "FALSE" then
				If err.number=0 then
					q="Delete from TransportTask_Tab where  OrderNo='" & OrderNo & "' and fcBA='" & fcba & "'"
					connStr.execute(q)
					if err.number  = 0 then
						connStr.CommitTrans
						else
						connStr.RollbackTrans   
						response.Write  "<div class = ""error"">" & err.description & "</div>"
					end if
				end if
				
					strQueryStatus = "BEGIN fac_journal_mam_api.MAM_STT_DOCNO('" & OrderNo & "', TO_NUMBER(TO_CHAR(TO_DATE('" & OrderDate  & "','MM/DD/YYYY'),'YYYY')), " & _ 
							"TO_NUMBER(TO_CHAR(TO_DATE('" & OrderDate  & "','MM/DD/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
							"END;"			
					'response.write strQueryStatus & "<br>"
					connStr.execute(strQueryStatus)
			else
				err.number = 1001
				err.description = "Tanggal """ & OrderDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
		cmdView	= "True"
		elseif Request("txtJenis")= "Details" then
		
			connstr.BeginTrans
			
			Cek_IsDailyClosing = IsDailyClosing(fcba, OrderDate)
			If Cek_IsDailyClosing = "FALSE" then

					'ambil qty original di issuedetail
					q = "Select STOCKCODE,STORECODE, LINENO, QUANTITYTRANSFERED, lineNohist from TransportTaskDtl_tab where ORDERNO='" & ORDERNO & _
					"' and LINENO='" & LineNo & _
					"' and fcBA='" & fcba & "'"

					'Response.Write q

					Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					rsReceiveDetail.Open q, connstr
											
					IF Not rsReceiveDetail.eof then
						txtLineNoHist = rsReceiveDetail("lineNohist")
						originalQuantity = rsReceiveDetail("QUANTITYTRANSFERED")
						txtItemCodeDelete = rsReceiveDetail("stockcode")
						txtLineNoKey = rsReceiveDetail("LINENO")
						txtStoreCode = rsReceiveDetail("STORECODE")
					End if
					rsReceiveDetail.Close
					set rsReceiveDetail = Nothing		

					q = "Select count(orderno) as total from Transporttask_tab where orderno='" & OrderNo & _
				"' and fcBA='" & fcba & "'"
				
			Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
			rsReceiveDetail.Open q, connstr
					
			IF Not rsReceiveDetail.eof then
				totalbaris 		= rsReceiveDetail("total")
			End if
			rsReceiveDetail.Close
			set rsReceiveDetail = Nothing		

					' 'function untuk mengembalikan qty ketika ada proses delete
					' qrylessstockonhand	= cmdUpdateStockOnHand(txtStoreCode,txtItemCodeDelete,originalQuantity,"TRT","Delete")
					' qrylessstockledger 	= cmdUpdateStockLedger(txtStoreCode,txtItemCodeDelete,OrderDate,originalQuantity,"TRT","Delete")

													
					q="Delete from TransportTaskDtl_tab where  OrderNo='" & OrderNo & "' and LineNo=" & LineNo & " and fcBA='" & fcba & "'"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry

					'jika total baris di tabel detail sisa satu, delete tabel issue header
				if totalbaris = "1" then
					q="Delete from Transporttask_tab where  orderno='" & OrderNo & "' and fcBA='" & fcba & "'"
					ExecuteAndLog "Transporttask.asp", q
				end if
					
					if err.number  = 0 then
					connStr.CommitTrans
					else
					connStr.RollbackTrans   
					response.Write  "<div class = ""error"">" & err.description & "</div>"
					end if
					
			else
				err.number = 1001
				err.description = "Tanggal """ & OrderDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If

			cmdView	= "True"
		elseif Request("txtJenis")="New" then
		connstr.BeginTrans	  
			'--Add by Ukie to handle Auto Number of PR
					If Ucase(OrderNo) = Ucase("(NEW)") Then
						oSql = " select nvl(controlnumber_api.get_nextnumber(upper('TRANSPORTTASK')" & _
							"                                         , " & year(textDate(OrderDate)) & "" & _
							"                                         , " & month(textDate(OrderDate)) & "" & _
							"                                         , '" & fcba & "'" & _
							" ), 'FALSE') OrderNo from dual"
						
						' Response.Write oSql
							
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.Open oSql, connstr
						If Not oRs.EOF Then
							txtOrderNo = oRs("OrderNo")
						End If
						oRs.Close
						Set oRs = Nothing
							OrderNo = txtOrderNo
					End If
					
					If txtOrderNo = "FALSE" Then
						err.Number = 1001
						err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
					End If
					
					'Response.Write PRCODE & txtPRCode
				'--End Added
			Cek_IsDailyClosing = IsDailyClosing(fcba, OrderDate)
			If Cek_IsDailyClosing = "FALSE" then
			if Request("txtJenis") <> "Approved" then
				dim sisaqty
				sisaqty 	= getSisaStock(STORECODE,ITEMCODE,OrderDate)
			
				if cdbl(QuantityTransfered) <= cdbl(sisaqty) then

						Dim rsPO
						Set rsPO = Server.CreateObject("ADODB.Recordset")
						rsPO.Open "Select OrderNo from TransportTask_Tab where OrderNo='" & OrderNo & "'AND FCBA = '" & fcba & "'", connstr
						IF rsPO.eof then	    
						
							q="Insert into TransportTask_Tab (OrderNo, OrderDate,REMARKS, ROWSTATEID, document_classification, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
							"values ('" & OrderNo & "', to_date('" & OrderDate & "','mm/dd/yyyy') " & _
							",'" &  REMARKS & "','Planned','206','" &  Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'))"
							connStr.execute(q)
							modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
							executeQry	= "Query executed:" & q
							ExecuteAndLog modulFile,executeQry
						End if
						rsPO.Close

				
					' Insert Transport Task Before Approved
						strLineMax =   " select nvl(max(LineNo),0) maxLineNo from TransportTaskDtl_tab " & _
									" where " & _
									" fcba = '" & fcba & "'"
						
						' response.write strLineMax		   
						Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
						rsstrLineMax.CursorType = 3
						rsstrLineMax.CursorLocation = 3
						rsstrLineMax.Open strLineMax, connstr
						
						if not rsstrLineMax.eof then
							tempMaxLineNo = cdbl(rsstrLineMax("maxLineNo")) + 1
						end if
									
						rsstrLineMax.Close
						set rsstrLineMax = Nothing 

						strLineMax =   " select sum(quantity) quantity from stockonhand " & _
									" where " & _
									" fcba = '" & fcba & "' and fcba = '"& ITEMCODE &"'"
						
						' response.write strLineMax		   
						Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
						rsstrLineMax.CursorType = 3
						rsstrLineMax.CursorLocation = 3
						rsstrLineMax.Open strLineMax, connstr
						
						if not rsstrLineMax.eof then
							quantity = cdbl(to_nol(rsstrLineMax("quantity")))  
						end if
									
						rsstrLineMax.Close
						set rsstrLineMax = Nothing 

						'-- AMH Apr 21 Temp di remark dulu aagr bisa kejar input

								q=  "Insert into  TransportTaskDtl_tab (ORDERNO, PRICEBY, DOCUMENTNO, LINEBY, STOCKCODE,LINENO,LINENOHIST, LOCATIONCODE,LOCATIONTYPE, QUANTITYTRANSFERED,STORECODE,JOBCODE,FCBADISTINATION,QUANTITYRECEIVED,PRICE,AMOUNT, " & _
									" REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
									" values ('" & OrderNo & "','" & PRICEBY & "','" & DOCUMENTNO & "','" & LINEBY & "','" & ITEMCODE & "','" & tempMaxLineNo & "','','" & LOCATIONCODE & "','" & LOCATIONTYPE & "'" & _
									", " & QuantityTransfered & ",'" & STORECODE & "','" & Jobcode & "','" & FCBADISTINATION & "',0,'" &  round(price,2) & "'," & round(amount,2) & ",'" & detail_remarks & "','" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI') )"
								connStr.execute(q)
								modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
								executeQry	= "Query executed:" & q
								ExecuteAndLog modulFile,executeQry
				else 
					err.number = 1001
					err.description = "Stock Kode Barang ini """ & ITEMCODE & """, Kosong, Silahkan tambah stock dulu, atau input dengan kodebarang lain"

				END IF
						
				end if
			else
				err.number = 1001
				err.description = "Tanggal """ & OrderDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
			
			if err.number  = 0 then
				connStr.CommitTrans
			else
			connStr.RollbackTrans   
			response.Write  "<div class = ""error"">" & err.description & "</div>"
			end if
			cmdView	= "True"
			
		elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			'Response.Write QuantityTransfered
			'Response.end
			connstr.BeginTrans
			
			Cek_IsDailyClosing = IsDailyClosing(fcba, OrderDate)
			If Cek_IsDailyClosing = "FALSE" then

				q = "update TransportTaskDtl_tab " & _
						"set JobCode = '" & JobCode & "'" & _
						", storecode = '"& storecode &"'" & _
						", QuantityTransfered = '" & QuantityTransfered & "'" & _
						", Remarks = '"& Remarks &"'" & _
						", documentno = '"& documentno &"'" & _
						", lineby = '"& lineby &"'" & _
						", priceby = '"& priceby &"'" & _
						", amount = '"& round(amount,2) &"'" & _
						", LASTUPDATE = sysdate " & _
						", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						"where " & _
						"ORDERNO = '" & ORDERNO & "' and " & _
						"fcba = '" & fcba &"' and lineno = '"& LineNo &"'"
					connStr.execute(q)	
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry
				
				
				q = "update TransportTask_tab " & _
						"set OrderDate = to_date('" & OrderDate & "','mm/dd/yyyy') " & _
						", REMARKS = '"& REMARKS &"'" & _
						", FCEDIT ='" & Session("CodeName") & "'" & _ 
						", LASTUPDATE = sysdate " & _
						", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						"where " & _
						"ORDERNO = '" & ORDERNO & "' and " & _
						"fcba = '" & fcba & "'"
				connStr.execute(q)		
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & q
				ExecuteAndLog modulFile,executeQry

			else
				err.number = 1001
				err.description = "Tanggal """ & OrderDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
			
			if err.number  = 0 then
				connStr.CommitTrans
			else
			connStr.RollbackTrans   
			response.Write  "<div class = ""error"">" & err.description & "</div>"
			end if
				
			cmdView	= "True"
	
		elseif Request("txtJenis")= "Approved" then 
			' modify jun (19 juni 2024)
			'query 
			dim arr_dtltrt_eof
			arr_dtltrt_eof =  true
			q = "SELECT T.LINENO, T.ORDERNO, T.FCBA, T.STOCKCODE, T.LOCATIONTYPE, " & _
			" T.LOCATIONCODE, T.JOBCODE, T.QUANTITYTRANSFERED, T.STORECODE, T.FCBADISTINATION, " & _
			" T.QUANTITYRECEIVED, T.PRICE, T.AMOUNT, T.REMARKS, T.LINENOHIST, " & _
			" T.PRICEBY, T.DOCUMENTNO, T.LINEBY " & _ 
			" FROM TRANSPORTTASKDTL_TAB T " & _ 
			" WHERE T.ORDERNO='" & ORDERNO & "'"
			set rsRNDetail = Server.CreateObject("ADODB.Recordset")
			rsRNDetail.Open q, connstr
			if not rsRNDetail.eof then 
				arr_dtltrt = rsRNDetail.getrows()
				arr_dtltrt_eof = false
			End IF

			rsRNDetail.close : set rsRNDetail = nothing
			if not arr_dtltrt_eof then 
				for x=0 to ubound(arr_dtltrt,2)

					ITEMCODE 			= arr_dtltrt(3,x)
					QuantityTransfered 	= arr_dtltrt(7,x)
					STORECODE 			= arr_dtltrt(8,x)
					
					
					qTotalRecordOnHand  =   "select count(*) totalRecordOnHand from stockonhand " & _
											"where " & _
											"storecode = '" & STORECODE & "' and " & _												 
											"stockcode = '" & ITEMCODE & "' and " & _
											"fcba = '" &  fcba & "'"
					'Response.Write qTotalRecordOnHand
					set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
					rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
					if not rsTotalRecordOnHand.eof then
							tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
					else
							tempTotalRecordOnHand = 0
					end if 
					
						
					If tempTotalRecordOnHand = 0 then
						
							' Tambahkan item stock dan qty di tabel stock onhand
							qryinsertstockOnhand 	= cmdInsertStockOnHand(STORECODE,ITEMCODE,QuantityTransfered,txtKolomtransaksi)
							
					Else	 
							
							' Update qty di tabel stock onhand
							qryUpdatestockOnhand 	= cmdUpdateStockOnHand(STORECODE,ITEMCODE,QuantityTransfered,txtKolomtransaksi,"Insert")
							
					End If	
					

					'CEK ITEM DAHULU DI TABEL STOCKLEDGER
					'JIKA TIDAK ADA INSERT BARU, SELAIN ITU UPDATE
					dim qryStockLedger, rsStockLedger,temp,totalRecordStockLedger
					qryStockLedger = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _ 
					" WHERE " & _ 
					" STORECODE = '" & STORECODE & "' and " & _
					" ITEMCODE  = '" & ITEMCODE & "' and " & _ 
					" year  	= to_char(to_date('" & OrderDate & "','mm/dd/yyyy'),'yyyy') and " & _ 
					" month 	= to_char(to_date('" & OrderDate & "','mm/dd/yyyy'),'mm') and " & _ 
					" FCBA 		= '" & fcba & "' "
					
					set rsStockLedger = Server.CreateObject("ADODB.Recordset")
					rsStockLedger.Open qryStockLedger, connstr
					if not rsStockLedger.eof then
							totalRecordStockLedger = cdbl(rsStockLedger("totalRecordStockLedger"))
					else
							totalRecordStockLedger = 0
					end if 
				

					'function ini digunakan untuk insert/update data di tabel stockledger
					' modify jun (19 juni 2024)
					If totalRecordStockLedger = 0 then
					
						'insert stockledger
						qryinsertstockledger 	= cmdInsertStockLedger(STORECODE,ITEMCODE,OrderDate,QuantityTransfered, "TRT")
					ELSE
					
						'update stockledger
						qryupdateStockLedger 	= cmdUpdateStockLedger(STORECODE,ITEMCODE,OrderDate,QuantityTransfered, "TRT", "Insert")
					END IF
				next
					strQueryStatus = "UPDATE TransportTask_Tab SET POSTINGFLG='1',rowstateid='Approved',lastupdate=sysdate WHERE fcba = '" & fcba & "' AND orderno = '" & OrderNo & "'"
					connStr.execute(strQueryStatus)	
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & strQueryStatus
					ExecuteAndLog modulFile,executeQry
					cmdView	= request("cmdView")
			end if

		end if
 
 	else
    	cmdView	= request("cmdView")
 	End if
else
    err.number = 1001
	err.description = "Accounting Period  " & ACCOUNTINGPERIOD & "/" & ACCOUNTINGYEAR & "  LOCKED !!!"
	cmdView	= request("cmdView")
End if 

 Dim rsTransportTask
 set rsTransportTask = server.CreateObject("Adodb.recordset")
 q="select OrderNo,fcba,to_char(OrderDate,'MM/DD/YYYY') OrderDate,REMARKS,rowstateid, document_classification from TransportTask_Tab  " & _
  " where OrderNo='" & OrderNo & "' and " & _
  " fcBA='" & fcba & "'"
 rsTransportTask.Open q, connstr
 if Not rsTransportTask.Eof then
    OrderDate = rsTransportTask("OrderDate")
	REMARKS = rsTransportTask("REMARKS")
	rowstateid = rsTransportTask("rowstateid")
	document_classification = rsTransportTask("document_classification")
	cmdView	= "True"
 End if
 

 
set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" WHERE at.document_type = 'TT' and at.fccode='TRANSPORT_TASK_LOCAL_STOCK'  " & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND aus.fcusercode = '" & Session("CodeName") & "'" & _
	"       AND aus.fcba = '" & fcba & "'"
 'response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	can_startauthorize =  cdbl(to_nol(rsFrez("total")))
else
	can_startauthorize = 0
End if
rsFrez.Close : set rsFrez = Nothing


 '-------------ISI Payment Form
 
Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        " and to_date ('" & OrderDate & "', 'mm/dd/yyyy') >= currentdate" & _
	        " and to_date ('" & OrderDate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        " and upper(system) = upper('Stores')"
			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
'response.Write oSQLQuery
oRecordSet.Close
Set oRecordSet = Nothing
 %>
 <form name="frmEntry" id="frmEntry" method="post" action="TransportTask.Asp">
	<table class="data" width="100%">
		<tr>
		  <td align="left">Order No</td>
		  <td align="left"><input type="text" maxlength="20" name="OrderNo" size="30" value="<%=OrderNo%>" autocomplete = "Off"><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
		   <td align="left">Date</td>
		   <td align="left"><input class="datepicker" onchange="get_Allowentry();" type="text" maxlength="10" id="orderdate" name="orderdate" size="10" value="<%=orderdate%>" >
		   <font color="#FF0000">&nbsp;*</font>
		</td>
		</tr>
		
		<tr> 
		    <td align="left">REMARKS</td>
			<td align="left"><textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea></td>
		</tr>
		<tr>
		  <td align="left">Status</td>
		  <td align="left"><input type="text" maxlength="50" name="rowstateid" size="30" id="statusnya" value="<%=rowstateid%>" style="background-color:#eee;" readonly><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
		  <td align="left">Posting</td>
		  <td align="left"><%=postingflgtxt%><input Type="Hidden" Value="<%=postingflg%>" id="postingflg_"></td>
		</tr>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"   name="Add" id="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('Print_transporttask.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&OrderNo=<%=OrderNo%>&OrderDate=<%=OrderDate%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<!-- Period Control In Here -->
			<% if RowStateId = "Planned" OR RowStateId = "" then %>
				<a href="TransportTask.Asp?OrderNo=<%=OrderNo%>&OrderDate=<%=OrderDate%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=OrderNo%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			<%Else%>
				<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
			<%End If%>
		<!-- End Period Control In Here -->
		
		<% if RowStateId = "Planned" then %>
			<%if can_startauthorize = 0 then%>
				
			<%else%>
				<a href="?OrderNo=<%=OrderNo%>&OrderDate=<%=OrderDate%>&fcba=<%=fcba%>&txtJenis=Approved" onClick="return confirm('Are you sure to start authorize this data :  <%=OrderNo%> ?');">Approved</a>
			<% end if%>
		<% end if%>

		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(OrderNo)) > 4 Then
   Dim rsPurchaseOrder,intRecordsPerPage,cColor
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q=" select ORDERNO,LINENO, LINENOHIST,STOCKCODE,stock_api.get_stockName(Stockcode,fcba) StockName, JOBCODE ,LOCATIONTYPE, LOCATIONCODE, " & _
   	 " PRICEBY, DOCUMENTNO, LINEBY, QUANTITYTRANSFERED, STORECODE, FCBADISTINATION, QUANTITYRECEIVED," & _
     " Price, AMOUNT, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA  from TransportTaskDtl_tab  " & _
     " where OrderNo='" & OrderNo & "' and " & _
     " fcBA='" & fcba & "' order by orderno, lineno"
   'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   <form action="TransportTask.Asp" method="post" name="eForm" id="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<% if request("txtJenis")= "Edit" then %>
		<tr>
		    <td align="left">Item Code</td>
		    <td align="left">
				<input type="text" onkeyup="showHint(this.value)" size="15" id="itemcode" name="itemcode" value="<%=ITEMCODE%>" style="background-color: #dddddd;" readonly onchange="javascript:getPrice(); " value="<%=itemcode%>">&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('popup_stock_new.asp?fcba=<%=fcba%>&IssueDate=<%=OrderDate%>','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font><span class="satuan_item" id="StockName">&nbsp;...</span></a>
			</td>
			<tr>
				<td align="left">Item Stock</td><td align="left">
					<input id="stockitem" value="" onkeyup="showHintStock(this.value)" type="text" size="15" name="stockitem" style="background:#dddddd; text-align: right" readonly>
					<span class="satuan_item" id="StokBarang" style="color: white">0</span>
			</tr>
		</tr>	
		<%else%>
		<tr>
		    <td align="left">Item Code</td>
		    <td align="left">
		     <input  type="text" onkeyup="showHint(this.value)" size="15" id="itemcode" name="itemcode" style="background-color: #dddddd;" readonly onchange="javascript:getPrice(); " value="<%=itemcode%>">&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('popup_stock_new.asp?fcba=<%=fcba%>&IssueDate=<%=OrderDate%>','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font><span class="satuan_item" id="StockName">&nbsp;...</span></a>
			</td>
			<tr>
				<td align="left">Item Stock</td><td align="left">
					<input id="stockitem" value="" onkeyup="showHintStock(this.value)" type="text" size="15" name="stockitem" style="background:#dddddd; text-align: right" readonly>
					<span class="satuan_item" id="StokBarang" style="color: white">0</span>
			</tr>
		</tr>
		<%end if%>
		<tr>
		    <td align="left">Job Code</td>
			
			<% 	Dim rsCombo,Data_JobsAccrual
				if Session("Estate_Mill") = "HO" then
				    q = " SELECT fccode jobcode, fcname Jobname, fcba" & _
						"	FROM INTRACO_TAB" & _
						"	WHERE " & _
						"	companycode = '"& companycode &"'" & _
						"	AND fcba <> '"& fcba &"'" & _
						"	AND status = 'TRUE'" & _
						"	union all" & _
						"	SELECT fccode jobcode, fcname Jobname, fcba" & _
						"	FROM INTRACO_TAB" & _
						"	WHERE " & _
						"	companycode = '"& companycode &"'" & _
						"	AND fcba = '"& fcba &"'" & _
						"	AND intraco_type ='SBU'" & _ 
						"	AND status = 'TRUE'" 
				else
					q = " SELECT fccode jobcode, fcname Jobname, fcba,companycode " & _
						"	FROM INTRACO_TAB" & _
						"	WHERE " & _
						"	companycode = '"& companycode &"'" & _
						"	AND fcba <> '"& fcba &"'" & _
						"	AND status = 'TRUE'" & _
						"	union all" & _
						"	SELECT afiliasi_db jobcode, 'AFILIASI '||companycode Jobname, fcba,companycode " & _
						"	FROM INTRACO_TAB" & _
						"	WHERE " & _
						"	companycode <> '"& companycode &"'" & _
						"	AND status = 'TRUE'" & _
						"   ORDER BY companycode,fcba" 
				End if
					set rsCombo = Server.CreateObject("ADODB.Recordset")
					'response.write q
					rsCombo.CursorType = 3
					rsCombo.CursorLocation = 3
					rsCombo.Open q, connstr
					if not rsCombo.eof then
				%>
				
				<td align="left"> 
					<select name="JobCode" onchange="javascript:transfer_value(this.options[this.selectedIndex].getAttribute('data-fcba'));">    
					<option value="" selected></option>
								<% Do While Not rsCombo.eof %>
									  <option data-fcba="<%=rsCombo("fcba")%>" value="<%=rsCombo("jobcode")%>" <% If Request("JobCode") = rsCombo("jobcode") Then %> Selected <%End if%>><%=rsCombo("fcba")%> : <%=rsCombo("jobcode")%> : <%=rsCombo("jobname")%></option> 
								<%    rsCombo.movenext
								   loop

								
					End if
					
					rsCombo.Close
					set rsCombo = Nothing
					q=""
				%>
					</select> <font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
		    <td align="left">Target Type</td>
			<td align="left">
			<input type="text" size="15" name="LocationType" value=<%=LocationType%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		<tr>
		    <td align="left">Target Code</td>
			<td align="left">
			<input type="text" size="15" name="LocationCode" value=<%=LocationCode%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		<tr>
			<td align="left">Store
			</td>
			<td align="left">      
				<select name="storecode" id="storecode" style="width: 18%" onchange="javascript:getPrice();">
				<option></option>
				<%
				qLov = " select fccode, fcname from store where fcba = '"& fcba &"' and iscentral='Y' and activation = 'Y' and nvl(sectionname,'-') = '-' "
	   
				Set rsLov = Server.CreateObject("ADODB.Recordset")
				rsLov.Open qLov, connstr
				if not rsLov.eof then
					while not rsLov.eof
				%>
					<option value="<%=rsLov("fccode")%>" <% if cstr(""&rsLov("fccode")) = cstr(""&storecode) then%> selected <%end if%>><%=rsLov("fccode")%> : <%=rsLov("fcname")%></option>
				<%
						rsLov.movenext
					wend
				end if
				%>
				</select><font color="#FF0000">&nbsp;*</font>
			</td>	
		</tr>	
		<tr>
			<td align="left">Price By</td>	
			<td align="left">
				<select id="priceby" name="priceby" value="<%=priceby%>" style="width: 25%" >
					<option value="OPB" <% If cstr(""&priceby) = "OPB" Then %> Selected <%End if%>>Opening Balance</option> 
					<!--
					<option value="GRN" <% If cstr(""&priceby) = "GRN" Then %> Selected <%End if%>>Received Note</option> 
					<option value="RCT" <% If cstr(""&priceby) = "RCT" Then %> Selected <%End if%>>Received Task</option>
					<option value="WRO" <% If cstr(""&priceby) = "WRO" Then %> Selected <%End if%>>Write Off/Adjustment</option> 
					-->
				</select>
			</td>	 
		</tr>
		<tr id="show_documentno" style="display: table-row;"  >
		    <td align="left">Document No</td>
		    <td align="left">
		     <input type="text" size="32" id="documentno" name="documentno" value="<%=documentno%>" onchange="getLinenoGRN(this.value);"> 
			</td>
		</tr>
		<tr id="show_lineby" style="display: table-row;">
			<td align="left">Line No.</td>
			<td align="left"> 
				<select name="lineby" id="lineby" onchange="javascript:getPrice(); " >
					<option></option> 
				</select>
			</td>	
		</tr> 		
			<input type="hidden" size="15" name="LineNo" value="<%=LineNo%>" readonly="True" style="background:#999999">
			<input type="hidden" size="15" name="Linenohist" value="<%=Linenohist%>" readonly="True" style="background:#999999">
		<tr>
		    <td align="left">Quantity Transfered</td>
			<td valign="center" align="left">
				<%
					if QuantityTransfered = "" Then
						kuantitas = 0
					else
						kuantitas = QuantityTransfered
					end if
				%>
				<input type="number" size="10" name="QuantityTransfered" step="0.01" min="0" max="999999" value="<%=kuantitas%>" id="quantity" onkeyup="hitungrumus();" onkeypress="return validateFloatKeyPress(this,event);">&nbsp;<font color="#FF0000">*
			</td>
	    </tr>
		<tr>
		  <td align="left">Fcba Distination</td>
		  <td align="left"><input id="FcbaDistination" type="text" maxlength="5" name="FcbaDistination" size="5" value="<%=FcbaDistination%>" readonly="True" style="background:#999999" ></td>
			<script>
				var FcbaDistination = document.getElementById("FcbaDistination");
				function transfer_value(fcba){
					//alert(fcba);
					FcbaDistination.value = fcba;
				}
			</script>
		</tr>
		<tr>
		  <td align="left">Quantity Received</td>
		  <td align="left"><input type="text" maxlength="10" name="QuantityReceived" size="5" value="<%=QuantityReceived%>" readonly="True" style="background:#999999" ></td>
		</tr>
		<tr>
		    <td align="left">Price</td>
			<td valign="center" align="left">
			  <input type="text" size="10" id="price" name="price" value="<%=price%>" onkeyup="hitungrumus();" readonly style="background:#999999" >
			</td>
	    </tr>
		<tr>
		    <td align="left">Amount Received</td>
			<td valign="center" align="left">
			  <input type="text" size="10" readonly="True" id="amount" name="AMOUNT" value="<%=AMOUNT%>" readonly="True" style="background:#999999"  >
			</td>
	    </tr>
		<tr> 
		    <td align="left">Remarks</td>
			<td align="left"><textarea  name="detail_remarks" maxlength = "100" cols="50" rows="3"><%=detail_remarks%></textarea></td>
		</tr>
		
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="OrderNo" value="<%=OrderNo%>">
		  <input type="hidden" name="OrderDate" value="<%=OrderDate%>">
		  <input type="hidden" name="REMARKS" value="<%=REMARKS%>">
		  <input type="hidden" name="rowstateid" value="<%=rowstateid%>">
		  <input type="hidden" name="document_classification" value="<%=document_classification%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <!-- Period COntrol In Here -->
			<%IF Request("txtJenis")="Edit" then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit" name="Add" value="Update" id="submitdata">
			<%Else%>
				<% if RowStateId = "" OR RowStateId = "Planned" then %>
				   <input type="hidden" name="txtJenis" value="New">
				   <input Type="Submit" name="Add" value="Add" id="submitdata">
				<%Else%>
			<%End if%>

			<!-- End Period COntrol In Here -->
			  <input Type="Reset" name="Reset" value="Reset">
			  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		<%End if%>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>	
	<TD align="left">Order No</TD>
	<TD align="left">Line No</TD>	
	<TD align="left">Line No Hist</TD>	
	<TD align="left">Stock Code</TD>
	<TD align="left">Stock Name</TD>
	<TD align="left">Job Code</TD>
	<TD align="left">Loc. Type</TD>	
	<TD align="left">Loc. Code</TD>
	<TD align="left">Price By</TD>
	<TD align="left">Document No</TD>
	<TD align="left">Line By</TD>
	<TD align="left">Quantity Transfered</TD>	
	<TD align="left">Store Code</TD>	
	<TD align="left">Fcba Distination</TD>	
	<TD align="left">Quantity Received</TD>	
	<TD align="left">Unit Price</TD>
	<TD align="left">Amount</TD>	
	<TD align="left">Remarks</TD>	
	</tr>   
   <%
   if Not rsPurchaseOrder.eof then
      nCountofField=rsPurchaseOrder.fields.count-1
	  no=1
	  While Not rsPurchaseOrder.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if
		  IF cdbl(rsPurchaseOrder("QUANTITYRECEIVED"))=cdbl(rsPurchaseOrder("QuantityTransfered")) then
		      val_IsTrue ="N" 
		  End If
		  
		  %>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<!-- Period COntrol In Here -->
				<% if RowStateId = "Planned" Or RowStateId = "Review Accounting" then %>
					<%if cdbl(to_nol(rsPurchaseOrder("QUANTITYRECEIVED"))) = 0 then%>
						<td valign="middle">
							<a href="TransportTask.Asp?OrderNo=<%=rsPurchaseOrder("OrderNo")%>&LineNo=<%=rsPurchaseOrder("LineNo")%>&Linenohist=<%=rsPurchaseOrder("Linenohist")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&QuantityTransfered=<%=rsPurchaseOrder("QuantityTransfered")%>&JobCode=<%=rsPurchaseOrder("JobCode")%>&LocationType=<%=rsPurchaseOrder("LocationType")%>&LocationCode=<%=rsPurchaseOrder("LocationCode")%>&FCBADISTINATION=<%=rsPurchaseOrder("FCBADISTINATION")%>&QUANTITYRECEIVED=<%=rsPurchaseOrder("QUANTITYRECEIVED")%>&Price=<%=rsPurchaseOrder("Price")%>&AMOUNT=<%=rsPurchaseOrder("AMOUNT")%>&StoreCode=<%=rsPurchaseOrder("StoreCode")%>&priceby=<%=rsPurchaseOrder("priceby")%>&documentno=<%=rsPurchaseOrder("documentno")%>&lineby=<%=rsPurchaseOrder("lineby")%>&REMARKS=<%=REMARKS%>&rowstateid=<%=rowstateid%>&detail_remarks=<%=rsPurchaseOrder("Remarks")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(1)%>" border="0"></a>
						</td>
						<td valign="middle">
							<a href="TransportTask.Asp?OrderNo=<%=rsPurchaseOrder("OrderNo")%>&LineNo=<%=rsPurchaseOrder("LineNo")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&StoreCode=<%=rsPurchaseOrder("StoreCode")%>&Linenohist=<%=rsPurchaseOrder("Linenohist")%>&OrderDate=<%=OrderDate%>&REMARKS=<%=REMARKS%>&rowstateid=<%=rowstateid%>&txtJenis=Details">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
						</td>
					<%Else%>
						<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
					<%End if%>
				<%Else%>
					<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
				<%End if%>
			<!-- End Period COntrol In Here -->
		    <%for i = 0 to nCountofField-4%>
			    <td align="left" title="<%=rsPurchaseOrder(i).Name%>" valign="top"><%=rsPurchaseOrder(i)%></td>
		    <%next%>	    
		  </tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
<%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->
<script type="text/javascript">	 

	$('#submitdata').click(function(){
		let val 	= $('#quantity').val();
		console.log(val);
		if (val == 0) {
			alert("Harap masukan QUANTITY!");

			return false
		}
	});

	function getselectedlineby() {
		var sfcba = document.getElementById("selectedlineby");
		sfcba.value = document.getElementById("lineby").value;
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
	// document.ready(function() {
	// 	document.getElementById("show_documentno").style.display="none";
	// 	document.getElementById("show_lineby").style.display="none";
	// 		$("#priceby").click(function() {
	// 		if( $(this).val() === "OPB") {
	// 			document.getElementById("show_documentno").style.display="none";
	// 			document.getElementById("show_lineby").style.display="none"; 
	// 		} else {      
	// 			document.getElementById("show_documentno").style.display="table-row";
	// 			document.getElementById("show_lineby").style.display="table-row"; 
	// 			document.getElementById("amount").value = 0;
	// 			document.getElementById("documentno").value ="";
	// 			document.getElementById("lineby").value ="";
	// 		}
	// 	}); 
	// 	getPrice(); 
	// 	hitungrumus();
	// }); 

	document.ready(function() {
		document.getElementById("show_documentno").style.display = "none";
		document.getElementById("show_lineby").style.display = "none";

		$("#priceby").click(function() {
			if ($(this).val() === "OPB") {
				document.getElementById("show_documentno").style.display = "none";
				document.getElementById("show_lineby").style.display = "none";
			} else {
				document.getElementById("show_documentno").style.display = "table-row";
				document.getElementById("show_lineby").style.display = "table-row";
				document.getElementById("amount").value = 0;
				document.getElementById("documentno").value = "";
				document.getElementById("lineby").value = "";
			}
		});

		getPrice();
		hitungrumus();

		// ← TAMBAHAN: auto-select & disable JobCode jika tabel sudah ada data
		autoSelectJobCodeFromTable();
	});

	$(function() {	
		get_Allowentry();
		
});

function autoSelectJobCodeFromTable() {
    // Ambil semua baris data dari tabel (skip baris header yang bgcolor="#0C5B24")
    // Kolom FcbaDistination ada di index ke-16 (0-based) pada <td> data baris.
    // Kita identifikasi via teks kolom header "Fcba Distination".

    var tableRows = document.querySelectorAll('table.data tr');
    var fcbaDistination = null;

    // Cari baris data pertama (bukan header)
    for (var i = 0; i < tableRows.length; i++) {
        var row = tableRows[i];
        // Baris header memiliki bgcolor="#0C5B24", skip
        var bg = row.getAttribute('bgcolor');
        if (bg && bg.toLowerCase() === '#0c5b24') continue;

        // Ambil semua <td> di baris ini
        var tds = row.querySelectorAll('td');
        // Berdasarkan urutan kolom di tabel:
        // NO(0), edit(1), del(2), OrderNo(3), LineNo(4), LineNoHist(5),
        // StockCode(6), StockName(7), JobCode(8), LocType(9), LocCode(10),
        // PriceBy(11), DocumentNo(12), LineBy(13), QtyTransfered(14),
        // StoreCode(15), FcbaDistination(16), QtyReceived(17),
        // UnitPrice(18), Amount(19), Remarks(20)
        if (tds.length >= 17) {
            var candidate = tds[16].innerText.trim();
            if (candidate !== '') {
                fcbaDistination = candidate;
                break;
            }
        }
    }

    if (!fcbaDistination) return; // Tidak ada data, tidak perlu apa-apa

    // Sekarang auto-select option di <select name="JobCode"> 
    // yang data-fcba-nya === fcbaDistination
    var jobCodeSelect = document.querySelector('select[name="JobCode"]');
    if (!jobCodeSelect) return;

    var options = jobCodeSelect.options;
    var matched = false;
    for (var j = 0; j < options.length; j++) {
        var opt = options[j];
        var dataFcba = opt.getAttribute('data-fcba');
        if (dataFcba && dataFcba.trim() === fcbaDistination) {
            jobCodeSelect.selectedIndex = j;
            // Trigger transfer_value agar FcbaDistination field juga terupdate
            transfer_value(dataFcba.trim());
            matched = true;
            break;
        }
    }

    // Disable dropdown JobCode agar user tidak bisa mengubah
    if (matched) {
        jobCodeSelect.disabled = true;
        // Tambahkan hidden input agar value tetap terkirim saat form submit
        // (disabled field tidak ikut tersubmit)
        var existingHidden = document.getElementById('hidden_JobCode');
        if (!existingHidden) {
            var hiddenInput = document.createElement('input');
            hiddenInput.type = 'hidden';
            hiddenInput.name = 'JobCode';
            hiddenInput.id   = 'hidden_JobCode';
            hiddenInput.value = jobCodeSelect.value;
            jobCodeSelect.parentNode.appendChild(hiddenInput);
        } else {
            existingHidden.value = jobCodeSelect.value;
        }
    }
}

function get_Allowentry()
{
	var orderdate 	= $('#orderdate').val();
	var postingflg 	=  $('#postingflg_').val();
	if (orderdate != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + orderdate, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE" || closingStatus == "TRUE" && $("#statusnya").val() != "Review Accounting"){
				$('#eForm').show();
				$('#txtposting').show();
				
			}else if(closingStatus == "FALSE" || closingStatus == "TRUE" && $("#statusnya").val() == "Review Accounting"){
				$('#eForm').show();
				$('#delhdr').show();
				//$('#poshdr').show();
				$('#Add').show();
			}else{
				alert("PERIODE SUDAH DI CLOSE !!! ");
				$('#eForm').hide();
				$('#txtposting').hide();
				
			}
			
		});
		
	} 
}

</script>
</body>
</html>