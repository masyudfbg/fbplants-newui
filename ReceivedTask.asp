<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="qtyOnHandFunction.inc"-->
<html>
<head>
	<title>ePlaS - Received Task </title>
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
var getlisttransporttask = function () {
	// build up the post string when passing variables to the server side page
	
	var xcode = document.getElementById("xcode").value;
	var selectedtransporttask = document.getElementById("selectedtransporttask").value;
	var PostStr = 'xcode='+xcode+'&transporttaskno='+selectedtransporttask;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/gettransporttaskreceived.asp', 'POST', PostStr, showlisttransporttask);
}

// The function for handling the response from the server
var showlisttransporttask = function (oXML) { 
    var response = oXML.responseText;
    document.getElementById("transporttasklist").innerHTML = "";
    document.getElementById("transporttasklist").innerHTML = response;

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

    // Validasi langsung saat user memilih Transport Task No
    $("#transporttaskno").on("change", function() {
        
        // Update hidden input selectedtransporttask
        var selectedVal = $(this).val(); // contoh: "ARM/TT/20/001 | Tanggal : 03/31/2020"
        $('#selectedtransporttask').val(selectedVal);

        if (!selectedVal) return;

        var orderDateStr = $('#orderdate').val(); // format: mm/dd/yyyy
        if (!orderDateStr) {
            alert("Peringatan!\nSilakan isi Order Date terlebih dahulu.");
            $(this).val("").trigger("chosen:updated");
            $('#selectedtransporttask').val("");
            return;
        }

        // Parse tanggal dari label Transport Task
        // Format: "ARM/TT/20/001 | Tanggal : 03/31/2020"
        var matchTanggal = selectedVal.match(/Tanggal\s*:\s*(\d{2})\/(\d{2})\/(\d{4})/i);
        if (!matchTanggal) return;

        var ttMonth = parseInt(matchTanggal[1], 10);
        var ttDay   = parseInt(matchTanggal[2], 10);
        var ttYear  = parseInt(matchTanggal[3], 10);

        // Parse Order Date (mm/dd/yyyy)
        var odParts = orderDateStr.split("/");
        var odMonth = parseInt(odParts[0], 10);
        var odDay   = parseInt(odParts[1], 10);
        var odYear  = parseInt(odParts[2], 10);

        var orderDate = new Date(odYear, odMonth - 1, odDay);
        var ttDate    = new Date(ttYear, ttMonth - 1, ttDay);

        // Kondisi 1: Bulan TT berbeda dengan bulan Order Date
        if (ttMonth !== odMonth || ttYear !== odYear) {
            alert("Peringatan!\n\n" +
                  "Bulan Transport Task No : " + ttMonth + "/" + ttYear + "\n" +
                  "Bulan Order Date        : " + odMonth + "/" + odYear + "\n\n" +
                  "Bulan Transport Task No harus sama dengan bulan Order Date!");
            $(this).val("").trigger("chosen:updated");
            $('#selectedtransporttask').val("");
            return;
        }

        // Kondisi 2: Order Date lebih kecil dari tanggal Transport Task No
        if (orderDate < ttDate) {
            alert("Peringatan!\n\n" +
                  "Order Date (" + orderDateStr + ") " +
                  "tidak boleh lebih kecil dari\n" +
                  "tanggal Transport Task No (" + 
                  matchTanggal[1] + "/" + matchTanggal[2] + "/" + matchTanggal[3] + ")!");
            $(this).val("").trigger("chosen:updated");
            $('#selectedtransporttask').val("");
            return;
        }
    });
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
	       <!--#include file="menu.asp"-->
    </div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;">
	

<%
'--Awal--
Dim xCode,xstoreCode,TempOrderNo,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,v_CrossCekPO,fcba
Dim Remarks
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
Dim Jobcode, LocationType, LocationCode, Reference, Remark
Dim qLineMax, rsqLineMax, tempMaxLine
Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
Dim sqlSQLCheck, rsSQL, qtyRemained
Dim qCheckStore, rsqCheckStore
Dim txtLocationType	
Dim TransportTaskNo,fcbasource

Dim txtOrderNo : txtOrderNo = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
Dim txtLocationCode : txtLocationCode = ""
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
Dim Cek_IsAccClosing : Cek_IsAccClosing = ""
Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD
Dim modulFile,executeQry,linenosource

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
selectedtransporttask = Request("selectedtransporttask")			
			
if request("TempOrderNo")="" then
	TempOrderNo="" '----- Nantinya Ambil dari Format Master
else
	TempOrderNo = request("TempOrderNo")
end if

'response.write TempOrderNo
 if Request("fcba") <> "" then
	fcba = Request("fcba")
 else	
	fcba = session("fcba") 
 end if

 Dim rsTRT,postingflgtxt,rowstate
 set rsTRT = server.CreateObject("Adodb.recordset")
 q="select to_char(ORDERDATE,'mm/dd/yyyy') ORDERDATE,postingflg,rowstate from RECEIVEDTASK_TAB  " & _
  " where ORDERNO='" & TempOrderNo & "' and " & _
  " fcBA='" & fcba & "'"

 rsTRT.Open q, connstr
 postingflg = "0"
 if Not rsTRT.Eof then
    OrderDate=rsTRT("ORDERDATE")
	postingflg = rsTRT("postingflg")
	rowstate = rsTRT("rowstate")
 End if


 IF postingflg = "1" THEN 
   postingflgtxt = "YA"
ELSE
   postingflgtxt	= "TIDAK"
END IF
 
	
If  Len(Trim(TempOrderNo)) = 0 Then TempOrderNo = "(NEW)"			
	txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
	'---- Data Header
	if Request("DATECREATED") <> "" then
		DATECREATED	=	Request("DATECREATED")
	else
		DATECREATED	=	textDate(date)
	end if
	
	if KonversiTanggalKeDbl(DATECREATED) > KonversiTanggalKeDbl(textDate(date)) then
		DATECREATED = textDate(date)
	end if
	
	ACCOUNTINGPERIOD  = left(OrderDate,2)
	ACCOUNTINGYEAR    = right(OrderDate,4)
	xCode = Request("xCode")  '-- fcba source
	HeaderRemarks = Request("HeaderRemarks")
			
	'---- Data Details

	TransportTaskNo = Request("TransportTaskNo")

	Cek_IsAccClosing = StatusAccClosingTgl(DATECREATED,fcba)  
	if 	Cek_IsAccClosing = "FALSE" then  
			
	'--- Action ke Database di Oracle [ insert,edit,delete ]
	If  Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")="Posting" or Request("txtJenis")="Approved" Then
		REMARKS	=	request("REMARKS")

		if Request("txtJenis")= "Header" then
			'Simpan Order No nya terlebih dahulu sebelum di delete
			connstr.BeginTrans
		
			Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
			If Cek_IsDailyClosing = "FALSE" then
				Set rsMNLRCV = Server.CreateObject("ADODB.Recordset")
				rsMNLRCV.Open " select OrderNo" & _
							" from RECEIVEDTASK_TAB" & _
							" where OrderNo = '" & TempOrderNo & "'" & _
							"       and fcBA = '" & fcba & "'", connstr
				IF Not rsMNLRCV.eof then
					txtOrderNo =  rsMNLRCV("OrderNo")
				End if
				rsMNLRCV.Close
				set rsMNLRCV = Nothing

				q = " delete from RECEIVEDTASK_TAB" & _
					" where OrderNo = '" & TempOrderNo & "'" & _
					"       and fcba = '" & fcba & "'"

				'connStr.execute(q) CCCC
				ExecuteAndLog "ReceivedTask.asp", q
				TempOrderNo = "(NEW)"

				q = "Delete from RECEIVEDTASKDTL_TAB where  OrderNo='" & txtOrderNo & "'"
				connStr.execute(q)
				modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
				executeQry	= "Query executed:" & q
				ExecuteAndLog modulFile,executeQry		
				
				if err.number = 0 then
					connStr.CommitTrans
				end if
				
				strQueryStatus = "BEGIN fac_journal_mam_api.MAM_SRT_DOCNO('" & TempOrderNo & "', TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED  & "','MM/DD/YYYY'),'YYYY')), " & _ 
								"TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED  & "','MM/DD/YYYY'),'MM')) , '" & fcba & "','System'); " & _ 
								"END;"			
				'response.write strQueryStatus & "<br>"
				connStr.execute(strQueryStatus)
				' connStr.CommitTrans
				
			else
				err.number = 1001
				err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
			
			cmdView	= "True"
	
		elseif Request("txtJenis")= "Details" then
			'Simpan Order No nya terlebih dahulu sebelum di delete
			connstr.BeginTrans
			
			Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
			If Cek_IsDailyClosing = "FALSE" then

				' update validasi jika rows yang didelte sisa 1, lakukan delete header --------
				q = "Select count(OrderNo) as total from RECEIVEDTASKDTL_TAB where OrderNo='" & TempOrderNo & _
				"' and fcBA='" & fcba & "'"
				
				Set rsMNLRCVDetail = Server.CreateObject("ADODB.Recordset")
				rsMNLRCVDetail.Open q, connstr
						
				IF Not rsMNLRCVDetail.eof then
					totalbaris 		= rsMNLRCVDetail("total")
				End if
				rsMNLRCVDetail.Close
				set rsMNLRCVDetail = Nothing	

				' -------------- BATAS -----------------------------
				
				Set rsMNLRCV = Server.CreateObject("ADODB.Recordset")
				rsMNLRCV.Open "Select OrderNo from RECEIVEDTASK_TAB where  OrderNo='" & TempOrderNo & "'" & _
				" and fcBA='" & fcba & "'", connstr
				IF Not rsMNLRCV.eof then
					txtOrderNo =  rsMNLRCV("OrderNo")
				End if
				rsMNLRCV.Close
				set rsMNLRCV = Nothing
				
				q = " Select LineNoHist" & _
					"       , QUANTITYRECEIVED" & _
					"       , STOCKCODE" & _
					"       , FCBASOURCE" & _
					"       , ORDERSOURCE" & _ 
					"       , LINENOSOURCE" & _
					"       , LOCATIONTYPE" & _
					"       , LOCATIONCODE" & _
					" from RECEIVEDTASKDTL_TAB" & _
					" where OrderNo = '" & TempOrderNo & "'" & _
					"       and LINENOSOURCE = '" & Request("linenosource") & "'" & _
					"       and fcBA = '" & fcba & "'"
			'	response.write q
				Set rsMNLRCVDetail = Server.CreateObject("ADODB.Recordset")
				rsMNLRCVDetail.Open q, connstr

						
				IF Not rsMNLRCVDetail.eof then
					txtLineNoHist = rsMNLRCVDetail("LINENOHIST")
					originalQuantity = CDbl(rsMNLRCVDetail("QUANTITYRECEIVED"))
					txtItemCodeDelete = rsMNLRCVDetail("STOCKCODE")
					txtItemLineKeyDelete = rsMNLRCVDetail("LINENOSOURCE")
					txtOrderSource = rsMNLRCVDetail("ORDERSOURCE")
					txtFcbaSource = rsMNLRCVDetail("FCBASOURCE")
					txtLocationType = rsMNLRCVDetail("LOCATIONTYPE")
					txtLocationCode = rsMNLRCVDetail("LOCATIONCODE")
				End if
				rsMNLRCVDetail.Close
				set rsMNLRCVDetail = Nothing					
				
					
					q = "Delete from stocktranshist " & _
						"where   " & _
						"lineno = " & txtLineNoHist & " and " & _
						"fcba = '" & fcba & "'"
					'	response.write q
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry
						
					q = "Delete from RECEIVEDTASKDTL_TAB where  OrderNo='" & TempOrderNo & "' and LINENOHIST='" & txtLineNoHist & _
						"' and fcBA='" & fcba & "'"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry		
					
					'jika total baris di tabel detail sisa satu, delete tabel issue header
					if totalbaris = "1" then
						q="Delete from RECEIVEDTASK_TAB where  ORDERNO='" & TempOrderNo & "' and fcBA='" & fcba & "'"
						connStr.execute(q)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry	
					end if
					
					if err.number=0 then
						connStr.CommitTrans
					End if
			else
				err.number = 1001
				err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
			
			cmdView	= "True"   	
					
        elseif Request("txtJenis")="Edit"  then	
		    
			cmdView	= "True"
		   	
        elseif Request("txtJenis")="New" then
			connstr.BeginTrans
					
			Cek_IsDailyClosing = IsDailyClosing(fcba, DATECREATED)
			If  Cek_IsDailyClosing = "FALSE" then
			
			    Dim AdaHeader
				q= "select * from RECEIVEDTASK_TAB  " & _
				   " where OrderNo='" & TempOrderNo & "' and " & _
				   " fcBA='" & fcba & "'"									
				Set rsCek = Server.CreateObject("ADODB.Recordset")
				rsCek.open q , connstr
				If not rsCek.eof then
					AdaHeader ="Y"
				Else
					AdaHeader ="N"
				end if 
				rsCek.Close
				Set rsCek = Nothing
				
				
				If AdaHeader ="N"	Then	
				    '--handle Auto Number of RECEIVEDTASK
					If  Ucase(TempOrderNo) = Ucase("(NEW)") Then
						oSql =  " select nvl(controlnumber_api.get_nextnumber(upper('RECEIVEDTASK')" & _
								"                                         , " & year(textDate(DATECREATED)) & "" & _
								"                                         , " & month(textDate(DATECREATED)) & "" & _
								"                                         , '" & fcba & "'" & _
								" ), 'FALSE') OrderNo from dual"
									
						'Response.Write oSql
										   
						Set oRs = Server.CreateObject("ADODB.Recordset")
						oRs.Open oSql, connstr, 3, 3
						If  Not oRs.EOF Then
							txtOrderNo = oRs("OrderNo")
						End If
						oRs.Close
						Set oRs = Nothing
						TempOrderNo = txtOrderNo
					End If
					
					'response.write OrderNo
							
					If  txtOrderNo = "FALSE" Then
						err.Number = 1001
						err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
					End If
					
					
					q = "Insert into RECEIVEDTASK_TAB (OrderNo, OrderDate,FcbaSource,OrderSource, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME,ROWSTATE) " & _
						"values ('" & TempOrderNo & "'," & " to_date('" & DATECREATED & "','mm/dd/yyyy'),'" & xCode & "','" & TransportTaskNo & "','" & _
						 Session("CodeName") & "','" & Session("CodeName") & "','" & _
						 Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'), 'Planned')"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry
					
				End if
				
				
				Dim TranportTaskDate_		
				Dim n, ItemPrice, ItemAmount 
		  
				For n=1 to (cdbl(request("no"))-1) 
								
						' Insert the History of Received Inventory Transaction
						dim QuantityReceived,ITEMCODE
						QuantityReceived 	= 	request("quantityreceived_"&TRIM(n))		
						ITEMCODE 			= 	request("STOCKCODE_"&TRIM(n))		
						STORECODE 			= 	request("xstoreCode_"&TRIM(n))

						qLineMax =  " Select nvl(max(lineno),0) tempMaxLine from stocktranshist " & _
									" Where " & _
									" Fcba = '" & fcba & "'" 
													
						set rsqLineMax = Server.CreateObject("ADODB.Recordset")
						rsqLineMax.Open qLineMax, connstr
						if not rsqLineMax.eof then
							tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
						else
							tempMaxLine = 0
						end if 

						tempMaxLine = tempMaxLine + 1
						
						
						q= "select last_day(TRANSPORTTASK_APi.GET_RCVDATE('" & xCode & "', '" & request("Ordersource_"&TRIM(n)) & "' )) TranportTaskDate from dual "
						'response.Write(q)				
						Set rsTranportTask = Server.CreateObject("ADODB.Recordset")
						rsTranportTask.open q , connstr
						If not rsTranportTask.eof then
							TranportTaskDate_ =rsTranportTask("TranportTaskDate")
						end if 
						rsTranportTask.Close
						Set rsTranportTask = Nothing
						
						q = " SELECT round(NVL(price,0),2) price, round(NVL(amount,0),2) amount FROM TRANSPORTTASKDTL_TAB a" & _
						    " WHERE a.fcba = '" & xCode & "' AND a.orderno = '" & request("Ordersource_"&TRIM(n)) & "' AND a.lineno = " & request("LineNoKey_"&TRIM(n)) 
 
                        'response.write q 						
						set rsitemprice = Server.CreateObject("ADODB.Recordset")
						rsitemprice.open q , connstr
						If not rsitemprice.eof then
							ItemPrice   = cdbl(to_nol(rsitemprice("price")))
							ItemAmount  = cdbl(to_nol(rsitemprice("amount")))
						Else
							ItemPrice   = 0
							ItemAmount  = 0
						end if 
						rsitemprice.Close
						Set rsitemprice = Nothing
																		
						qa = "Insert into StockTransHist ( " & _
							" LINENO, FCBA, DOCUMENTNO, STOCKCODE, TRANSDATE, TRANSTYPE, STORECODE, SOURCETYPE, SOURCECODE, TARGETTYPE, TARGETCODE, QUANTITY, LASTUPDATE, LASTTIME) values ( "& _
							tempMaxLine & ",'" & _ 
							fcba & "','" & _
							TempOrderNo & "','" & _
							request("STOCKCODE_"&TRIM(n)) &"'," & _
							"to_date('" & DATECREATED & "','mm/dd/yyyy') ,'" & _
							"MANUAL RECEIVED" & "','" & _
							request("xstoreCode_"&TRIM(n)) & "'," & _
							"'CB','BS','" & _
							request("locationtype_"&TRIM(n)) & "','" & _
							request("xstoreCode_"&TRIM(n)) & "'," & _
							" TO_NUMBER(REPLACE('"&request("quantityreceived_"&TRIM(n))&"', ',', '.')), sysdate, To_Char(sysdate,'HH24:MI') ) "
						' response.write qa
						connStr.execute(qa)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & qa
						ExecuteAndLog modulFile,executeQry
									
						q = "Insert into  RECEIVEDTASKDTL_TAB (OrderNo, LOCATIONTYPE, LOCATIONCODE, STOCKCODE, JOBCODE, QUANTITYRECEIVED, PRICERECEIVED, AMOUNTRECEIVED, " & _
							"  FCBASOURCE, ORDERSOURCE, LINENOSOURCE, QUANTITYSOURCE, LINENOHIST, remarks, FCENTRY, FCEDIT, FCIP, FCBA,LASTUPDATE, LASTTIME) " & _
							" values ('" & TempOrderNo & "','" & request("locationtype_"&TRIM(n)) & "','" & request("xstoreCode_"&TRIM(n)) &  _
							"','" & request("STOCKCODE_"&TRIM(n)) & "','" & request("Jobcode_"&TRIM(n)) & "', TO_NUMBER(REPLACE('"&request("quantityreceived_"&TRIM(n))&"', ',', '.'))," & _
							"round(" & ItemPrice & ",2) , round(" & ItemAmount & ",2),'" & xCode & "','" & request("Ordersource_"&TRIM(n)) & "'," & request("LineNoKey_"&TRIM(n)) & ", TO_NUMBER(REPLACE('" & request("quantityreceived_"&TRIM(n)) & "', ',', '.')), "& tempMaxLine & ", '" & request("remarks_"&TRIM(n)) & "', " & _
							"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'))"
						
						connStr.execute(q)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry

						' UPDATE TRANSPORT TASK DETAIL
						qLineMaxUpdate = "UPDATE transporttaskdtl_tab SET QUANTITYRECEIVED = TO_NUMBER(REPLACE('"&request("quantityreceived_"&TRIM(n))&"', ',', '.')) WHERE orderno = '" & request("Ordersource_"&TRIM(n)) & "'"
						' response.write qLineMaxUpdate
						connStr.execute(qLineMaxUpdate)
				Next
							
				if  err.number=0 then
					connStr.CommitTrans
				End if
						
			Else
				err.number = 1001
				err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
		   cmdView	= "True"
        
		elseif Request("txtJenis")="Approved" then

			dim arr_dtlrc_eof
			arr_dtlrc_eof =  true
			q = "SELECT R.ORDERNO, R.LOCATIONTYPE, R.LOCATIONCODE, R.STOCKCODE, R.JOBCODE, " & _ 
			" R.QUANTITYRECEIVED, R.PRICERECEIVED, R.AMOUNTRECEIVED, R.FCBASOURCE, R.ORDERSOURCE, " & _ 
			" R.LINENOSOURCE, R.QUANTITYSOURCE, R.LINENOHIST " & _ 
			" FROM RECEIVEDTASKDTL_TAB R " & _ 
			" WHERE R.ORDERNO='" & TempOrderNo & "'"
			set rsRCDetail = Server.CreateObject("ADODB.Recordset")
			rsRCDetail.Open q, connstr
			if not rsRCDetail.eof then 
				arr_dtlrc = rsRCDetail.getrows()
				arr_dtlrc_eof = false
			End IF
		
			rsRCDetail.close : set rsRCDetail = nothing
			if not arr_dtlrc_eof then 
				for x=0 to ubound(arr_dtlrc,2)
					STORECODE 			= arr_dtlrc(2,x) ' kode gudang
					ITEMCODE 			= arr_dtlrc(3,x) ' kodebarang
					QuantityReceived 	= arr_dtlrc(5,x) ' qty received
					linenosource	 	= arr_dtlrc(10,x) ' line no dari tabel transporttask detail
					transporttaskno	 	= arr_dtlrc(9,x) ' notransport task dari tabel transporttask detail
					fcbasource	 		= arr_dtlrc(8,x) ' fcba sumber pengirim barang

						qTotalRecordOnHand = "select count(*) totalRecordOnHand from stockonhand " & _
												"where " & _
												"storecode = '" & STORECODE  & "' and " & _												 
												"stockcode = '" & ITEMCODE & "' and " & _
												"fcba = '" &  fcba & "'"

						set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
						rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
						if not rsTotalRecordOnHand.eof then
							tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
						else
							tempTotalRecordOnHand = 0
						end if 
												
						if  tempTotalRecordOnHand = 0 then			
							
							' Tambahkan item stock dan qty di tabel stock onhand
							qryinsertstockOnhand 	= cmdInsertStockOnHand(STORECODE,ITEMCODE,QuantityReceived,"RCV")
													
						Else	
							' Update qty di tabel stock onhand
							qryUpdatestockOnhand 	= cmdUpdateStockOnHand(STORECODE,ITEMCODE,QuantityReceived,"RCV","Insert")  

						End If

						'CEK ITEM DAHULU DI TABEL STOCKLEDGER
						'JIKA TIDAK ADA INSERT BARU, SELAIN ITU UPDATE
						dim qryStockLedger, rsStockLedger,temp,totalRecordStockLedger
						qryStockLedger = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _ 
						" WHERE " & _ 
						" STORECODE = '" & STORECODE & "' and " & _
						" ITEMCODE  = '" & ITEMCODE & "' and " & _ 
						" year  	= to_char(to_date('" & DATECREATED & "','mm/dd/yyyy'),'yyyy') and " & _ 
						" month 	= to_char(to_date('" & DATECREATED & "','mm/dd/yyyy'),'mm') and " & _ 
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
							qryinsertstockledger 	= cmdInsertStockLedger(STORECODE,ITEMCODE,DATECREATED,QuantityReceived, "RCV")
						ELSE
						
							'update stockledger
							qryupdateStockLedger 	= cmdUpdateStockLedger(STORECODE,ITEMCODE,DATECREATED,QuantityReceived, "RCV", "Insert")
						END IF

						'update kolom qtyreceive pada table transporttask detail
						q = " Update TRANSPORTTASKDTL_TAB  a set a.QUANTITYRECEIVED = " &  QuantityReceived & " " & _
						" where a.fcba = '" & fcbasource & "' and a.orderno = '" & transporttaskno & "'  and a.lineno = '" & linenosource & "'"  
						connStr.execute(q)
						modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
						executeQry	= "Query executed:" & q
						ExecuteAndLog modulFile,executeQry
				NEXT
			END IF

			strQueryStatus = "UPDATE RECEIVEDTASK_TAB SET POSTINGFLG='1',rowstate='Approved',lastupdate=sysdate,fcedit='" & Session("CodeName") & "' WHERE fcba = '" & fcba & "' AND orderno = '" & TempOrderNo & "'"
			connStr.execute(strQueryStatus)
			modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
			executeQry	= "Query executed:" & strQueryStatus
			ExecuteAndLog modulFile,executeQry
	
		    cmdView	= "True"
		End if
		cmdView	= "True"
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

 q="select to_char(orderdate,'mm/dd/yyyy')orderdate,fcbasource,Ordersource,postingflg,rowstate from RECEIVEDTASK_TAB  " & _
  " where OrderNo='" & TempOrderNo & "' and " & _
  " fcBA='" & fcba & "'"
SudahClose="N"
' response.write q

 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED=rsVoucher("OrderDate")
	xCode=rsVoucher("fcbasource")
	TransportTaskNo = rsVoucher("Ordersource")
	postingflg = rsVoucher("postingflg")
	rowstate = rsVoucher("rowstate")
	SudahClose="Y"
 End if
 rsVoucher.Close
 Set rsVoucher = Nothing
 
 'response.write xCode & TransportTaskNo
 
' If Err.Number <> 0 then
' 	connStr.RollbackTrans
' 	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
' End if


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
'response.Write oSQLQuery
oRecordSet.Close
Set oRecordSet = Nothing


set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" WHERE at.document_type = 'RT' and at.fccode='507'  " & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'FALSE'" & _
	"       AND aus.fcusercode = '" & Session("CodeName") & "'" & _
	"       AND aus.fcba = '" & fcba & "'"

rsFrez.Open q, connstr
if Not rsFrez.Eof then
	isApproved =  cdbl(to_nol(rsFrez("total")))
else
	isApproved = 0
End if
rsFrez.Close : set rsFrez = Nothing 
'Response.Write isApproved


 %>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-inbox" aria-hidden="true"></i>Header Received Task</div><% end if %>
 <form name="frmEntry" id="frmEntry" method="post" action="ReceivedTask.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Order No</td>
		<td align="left"><input type="text" maxlength="50" name="TempOrderNo" size="30" value="<%=TempOrderNo%>"><font color="#FF0000">&nbsp;*</font></td></tr><tr>
		<td align="left">Order Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="DATECREATED" id="orderdate" onchange="javascript:get_Allowentry();" size="10" value="<%=DATECREATED%>">
		<font color="#FF0000">&nbsp;*</font>
		</td></tr>
		<tr>
		<td align="left">Fcba Source</td>
		<td align="left" readonly="TRUE" > 
        <%if SudahClose="Y" then%>
		   <input type="text" name="xCode" value="<%=xcode%>" style="background-color: #eee;" readonly>
		<%else%>   		
		   <select id="xcode" name="xcode" onchange="javascript:get_Allowentry();getlisttransporttask();" class="chosen-select" style="width: 40%">
				<option value=""></option>
				<%
				qlov = 	 " SELECT DISTINCT dtl.fcba fccode,businessunit_api.get_fcname(dtl.fcba) fcname " & _
						 " FROM transporttaskdtl_tab dtl " & _
	                     " INNER JOIN transporttask_tab hdr ON (hdr.fcba = dtl.fcba AND hdr.orderno = dtl.orderno) " & _
	                     " WHERE " & _
	    	             " dtl.fcbadistination = '" & fcba & "' " & _
						 " AND hdr.rowstateid = 'Approved' " & _
						 " AND (dtl.quantitytransfered - dtl.quantityreceived ) > 0 " 
			  
				Set rslov = Server.CreateObject("ADODB.Recordset")
				rslov.Open qlov, connstr 
				if Not rslov.eof then
					While Not rslov.EOF
				%>
				<option value="<%=rslov("fccode")%>" <%if xcode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%> : <%=rslov("fcname")%></option>
				<%
						rslov.MoveNext
					Wend
				end if
				%>
			</select>
		<%End if%>
		<font color="#FF0000">*</font>
        </td>
		</tr>
		<tr>
		<td align="left">Transport Task No</td>
		<td align="left"> 
		<%if SudahClose="Y" then  %>
		    <input type="text" readonly style="background-color: #eee;" value="<%=TransportTaskNo%>" Selected> 
		<%else%>   
			<div id="transporttasklist"></div>
			 <input type="Hidden" id="selectedtransporttask" name="selectedtransporttask" value="<%=selectedtransporttask%>">
		<% End if %>	  
		</td>
		</tr>
		<tr>
		  <td align="left">Posting / Status</td>
		  <td align="left">
		  	<%=postingflgtxt & " / " & rowstate%><input Type="Hidden" Value="<%=postingflg%>" id="postingflg_">
			<input type="hidden" value="<%=postingflgtxt%>" id="statusnya" />
		  </td>
		</tr>
		
		
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions"> 
				<input Type="Hidden" Value="True" Name="cmdView">  
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">   	 
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  
				<a href="ReceivedTask.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&TempOrderNo=<%=TempOrderNo%>&TransportTaskNo=<%=TransportTaskNo%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
					
			   	<% if rowstate = "Planned" OR rowstate = "" Or rowstate = "Review Accounting" then %>
				<a id="txtdelete1" href="ReceivedTask.Asp?TempOrderNo=<%=TempOrderNo%>&DATECREATED=<%=DATECREATED%>&TransportTaskNo=<%=TransportTaskNo%>&txtJenis=Header"><img id="txtdelete2" src="IMAGES/delete.gif" alt="Delete Data : <%=TempOrderNo%>" border="0" onClick="return confirm('Are you sure to delete');"></a>	
				<%Else%>
				<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				
				<% if RowState = "Planned" Or rowstate = "Review Accounting" then %>
					<%if isApproved = 0 then%>
						
					<%else%>
						<a id="txtposting" href="?TempOrderNo=<%=TempOrderNo%>&DATECREATED=<%=DATECREATED%>&fcba=<%=fcba%>&txtJenis=Approved" onClick="return confirm('Are you sure to Posting Journal this data :  <%=TempOrderNo%> ?');">Approved</a>
					<% end if%>
				<% end if%>
				<a id="txtposting" href="?TempOrderNo=<%=TempOrderNo%>&DATECREATED=<%=DATECREATED%>&fcba=<%=fcba%>&txtJenis=Approved" onClick="return confirm('Are you sure to Posting Journal this data :  <%=TempOrderNo%> ?');">Approved</a>
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Purchase"
		dataSubmodule 	= "Manual Receive "
		dataDetail 		= "NULL"
		%>
		&nbsp;&nbsp;&nbsp;<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td> 
		</tr> 
    </table>    
 </form>
<%
If cmdView = "True" and LEN(TRIM(TempOrderNo)) > 4 and len(trim(TempOrderNo)) > 0 Then
	Dim rsMaRcv,intRecordsPerPage,cColor
	set rsMaRcv = server.CreateObject("Adodb.recordset")
	If SudahClose = "Y" Then
		q	= " select a.Linenosource, " & _
			  "		   a.ordersource, " & _
			  "		   a.stockcode, " & _
			  "		   stock_api.get_stockName(a.stockcode,a.fcba) Stockname, " & _
			  "		   a.quantityreceived, " & _
			  "		   round(a.PriceReceived,2) PRICE, " & _
			  "		   round(a.amountreceived,2) amount, " & _
			  "		   a.jobcode, " & _
			  "        job_api.get_jobcode_name(a.jobcode,a.fcba) Jobname,"& _
			  " 	   a.locationtype, " & _
			  "		   a.locationcode  " & _
			  "		   , remarks" & _
			  "	from RECEIVEDTASKdtl_tab a" & _
			  "	where fcba='" & fcba & "' and a.orderno ='" & TempOrderNo & "' " & _
			  "	order by a.Linenosource"
	Else
		
		q   = "SELECT " & _
				"	a.lineno            AS Linenosource, " & _
				"	a.orderno           AS Ordersource, " & _
				"	a.stockcode, " & _
				"	s.fcname            AS Stockname, " & _
				"	a.quantitytransfered AS QuantitySource, " & _
				"	a.quantitytransfered AS QuantityReceived, " & _
				"	ROUND(a.price,2)    AS Price, " & _
				"	ROUND(a.amount,2)   AS Amount, " & _
				"	sg.controljob       AS JobCode, " & _
				"	sg.fcname           AS JobName, " & _
				"	'ST'                AS LocationType, " & _
				"	'ST1'               AS LocationCode,  " & _
				"	a.remarks " & _
				" FROM transporttaskdtl_tab a " & _
				" JOIN stock s " & _
				"	ON s.fccode = a.stockcode and s.fcba=a.fcba " & _
				" JOIN stockgroup sg " & _
				"	ON sg.fccode = s.stockgroupcode and sg.fcba=s.fcba " & _
				" WHERE a.fcba = '" & xCode & "' " & _
				" AND a.orderno = '" & TransportTaskNo & "' " & _
				" AND (a.quantitytransfered - a.quantityreceived) > 0 " & _
				" ORDER BY a.lineno"
    End If	
	
 ''   response.Write q	& dataSubmodule
   Set rsMaRcv = Server.CreateObject("ADODB.Recordset")
   rsMaRcv.CursorType = 3
   rsMaRcv.CursorLocation = 3
   rsMaRcv.Open q, connstr%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form  id="receviedtask"  action="ReceivedTask.Asp" method="post" name="eForm" id="eForm">
   <table class="data" border="1" width="100%" >
    <%if SudahClose<>"Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">LineNoKey</td>
	   <td align="center">Order No Source</td>
	   <td align="center">Item Code</td>
	   <td align="center">Item Name</td>
	   <td align="center">Qty Source</td>
	   <td align="center">Price Source</td>
	   <td align="center">Amount Source</td>
	   <td align="center">Job Code</td>
	   <td align="center">Job Name</td>
	   <td align="center">Remarks</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <td align="center">Quantity Received</td>
	   <!--<td align="center">Amount</td>-->
   	   </tr>
	 <%Else%>
       <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	   <td align="center">No</td>
	   <td align="center">&nbsp;</td>
	   <td align="center">LineNoKey</td>
	   <td align="center">Order No Source</td>
	   <td align="center">Item Code</td>
	   <td align="center">Item Name</td>
	   <td align="center">Quantity Received</td>
	   <td align="center">Price Received</td>
	   <td align="center">Amount Received</td>
	   <td align="center">Job Code</td>
	   <td align="center">Job Name</td>
	   <td align="center">Target Type</td>
	   <td align="center">Target Code</td>
	   <td align="center">Remarks</td>
	   <!--<td align="center">Amount</td>-->
   	   </tr>
     	 
   <% End if
   
    Dim qkt : qkt = "quantityreceived_"
	
   if Not rsMaRcv.eof then
      nCountofField=rsMaRcv.fields.count-1
	  no=1
	  While Not rsMaRcv.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
			<tr bgcolor="<%=cColor%>">
				<td align="left" valign="top"><%=NO%></td>
				<%IF SudahClose = "Y" then%>
					<td valign="middle" id="txtdelete">
						<% if RowState = "Planned" Or RowState = "Review Accounting" then %>
							<a id="txtdeletex" href="ReceivedTask.Asp?TempOrderNo=<%=TempOrderNo%>&ItemCode=<%=rsMaRcv("STOCKCODE")%>&linenosource=<%=rsMaRcv("linenosource")%>&DATECREATED=<%=DATECREATED%>&txtJenis=Details">
							<img id="txtdelete" src="IMAGES/delete.gif" alt="Delete Data : <%=rsMaRcv(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
						<%else%>
							<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
						<% end if%>
					
						

						
					</td>
					<%for i = 0 to nCountofField%>
						<td align="left" title="<%=rsMaRcv(i).Name%>" valign="top"><%=rsMaRcv(i)%></td>
					<%next%>
				<%End if%>
				
				<%if SudahClose<>"Y" Then%>
					<td align="Right" title="LineNoKey" valign="top"><%=rsMaRcv("Linenosource")%></td>
					<td align="left" title="Order Source" valign="top"><%=rsMaRcv("Ordersource")%></td>
					<td align="left" title="Item Code" valign="top"><%=rsMaRcv("STOCKCODE")%></td>
					<td align="left" title="Item Name" valign="top"><%=rsMaRcv("Stockname")%></td>
					<td align="left" title="QuantitySource" valign="top"><%=rsMaRcv("QuantitySource")%></td>
					<td align="Right" title="Price" valign="top"><%=FormatNumber(rsMaRcv("Price"),2)%></td>
					<td align="Right" title="Amount" valign="top"><%=FormatNumber(rsMaRcv("Amount"),2)%></td>
                    <td align="left" title="Jobcode" valign="top"><%=rsMaRcv("Jobcode")%></td> 
					<td align="left" title="Jobname" valign="top"><%=rsMaRcv("Jobname")%></td>
					<td align="left" title="remarks" valign="top"><%=rsMaRcv("remarks")%></td>
					<td align="left" title="locationtype" valign="top"><%=rsMaRcv("locationtype")%></td>
                    <td align="left">        
							
							<select name="xstoreCode_<%=no%>" value="<%=request("xstoreCode_"&TRIM(no))%>"> 
								<%cmdCallCombo "Store","xstoreCode_"&TRIM(no),"ACTIVATION", "Y"%>
							</select>
							<font color="#FF0000"></font>
					</td>
					<td align="Right" title="quantityreceived" valign="top"><%=FbpNum2(rsMaRcv("quantityreceived"))%></td>
					
					<input type="hidden" name="LineNoKey_<%=no%>" value="<%=rsMaRcv("Linenosource")%>"></td>
					<input type="hidden" name="Ordersource_<%=no%>" value="<%=rsMaRcv("Ordersource")%>"></td>
					<input type="hidden" name="STOCKCODE_<%=no%>" value="<%=rsMaRcv("STOCKCODE")%>"></td>
					<input type="hidden" name="QuantitySource_<%=no%>" value="<%=rsMaRcv("QuantitySource")%>"></td>
					<input type="hidden" name="Jobcode_<%=no%>" value="<%=rsMaRcv("Jobcode")%>"></td>
					<input type="hidden" name="locationtype_<%=no%>" value="<%=rsMaRcv("locationtype")%>"></td>
					<input type="hidden" name="quantityreceived_<%=no%>" value="<%=rsMaRcv("quantityreceived")%>"></td>
					<input type="hidden" name="remarks_<%=no%>" value="<%=rsMaRcv("remarks")%>"></td>
					
			<%End if%>
		  </tr>
		  <%
		  rsMaRcv.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("<tr><td colspan=""13"" class=""fbp-doc-empty"">Data is empty</td></tr>")	  
   end if 
   rsMaRcv.Close
   set rsMaRcv = Nothing
   %>
	  <tr><td colspan="14" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="TempOrderNo" value="<%=TempOrderNo%>">
		  <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="TransportTaskNo" value="<%=TransportTaskNo%>">
		  <input type="hidden" name="selectedtransporttask" value="<%=selectedtransporttask%>">
		  <input type="hidden" name="No" value="<%=No%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" id="btnsubmit" name="Add" value="Update">
		  <%Else%>
		  <%if rowstate="Planned" or rowstate="" then%>
		       <input type="hidden" name="txtJenis" value="New">
			   <input Type="Submit" id="btnsubmit" name="Add" value="Submit" onClick="if(!validateTransportTaskDate()) return false; return confirm('Apakah Anda sudah Cek jumlah dan qualitas dari tiap Barang yg diterima ?');">
		   <%end if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
     
   </TABLE>
   </form>
<%
End if  
 
'-------------akhir----------------------
  %>
 </div> </div>
  <!--#include file="bottom.jsp"-->
  <script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	
	// $(".chosen-select").empty();
	
	// var config = {
	//   '.chosen-select'           : {},
	//   '.chosen-select-deselect'  : {allow_single_deselect:true},
	//   '.chosen-select-no-single' : {disable_search_threshold:10},
	//   '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
	//   '.chosen-select-width'     : {width:"45%"}
	// }
	// for (var selector in config) {
	//   $(selector).chosen(config[selector]);
	// }
	// $(".chosen-select").trigger("chosen:updated");
</script>
<script>

function getselectedtransporttask() {
    var sfcba = document.getElementById("selectedtransporttask");
    sfcba.value = document.getElementById("transporttaskno").value;

    var selectEl     = document.getElementById("transporttaskno");
    var selectedText = selectEl.options[selectEl.selectedIndex].text;

    if (!selectedText || selectedText.trim() === "") return;

    var orderDateStr = $('#orderdate').val();
    
    // ✅ KUNCI: jika orderdate kosong, JANGAN lakukan validasi, langsung return
    if (!orderDateStr || orderDateStr.trim() === "") return;

    var matchTanggal = selectedText.match(/Tanggal\s*:\s*(\d{2})\/(\d{2})\/(\d{4})/i);
    if (!matchTanggal) return;

    var ttMonth = parseInt(matchTanggal[1], 10);
    var ttDay   = parseInt(matchTanggal[2], 10);
    var ttYear  = parseInt(matchTanggal[3], 10);

    var odParts = orderDateStr.split("/");
    if (odParts.length < 3) return;

    var odMonth = parseInt(odParts[0], 10);
    var odDay   = parseInt(odParts[1], 10);
    var odYear  = parseInt(odParts[2], 10);

    // ✅ Tahun harus sudah 4 digit
    if (odYear < 1000) return;

    var orderDate = new Date(odYear, odMonth - 1, odDay);
    var ttDate    = new Date(ttYear, ttMonth - 1, ttDay);

    // Kondisi 1: Bulan atau Tahun berbeda
    if (ttMonth !== odMonth || ttYear !== odYear) {
        alert("Peringatan!\n\n" +
              "Bulan Transport Task No : " + ttMonth + "/" + ttYear + "\n" +
              "Bulan Order Date        : " + odMonth + "/" + odYear + "\n\n" +
              "Bulan Transport Task No harus sama dengan bulan Order Date!");
        selectEl.value = "";
        $(selectEl).trigger("chosen:updated");
        sfcba.value = "";
        return;
    }

    // Kondisi 2: Order Date lebih kecil dari tanggal Transport Task No
    if (orderDate < ttDate) {
        alert("Peringatan!\n\n" +
              "Order Date (" + orderDateStr + ") " +
              "tidak boleh lebih kecil dari\n" +
              "tanggal Transport Task No (" +
              matchTanggal[1] + "/" + matchTanggal[2] + "/" + matchTanggal[3] + ")!");
        selectEl.value = "";
        $(selectEl).trigger("chosen:updated");
        sfcba.value = "";
        return;
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
document.ready(function() {
	getlisttransporttask();
	getselectedtransporttask();

});

$(function() {	
	get_Allowentry();
});

//validasi lock inputan dri periode akunting
function get_Acc_Allowentry()
{
	var datecreated = $('#DATECREATED').val();
	if (datecreated != "") {
		$.post("acc_allow_entry.asp?tglbkm=" + datecreated, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE" && $("#statusnya").val() != "Review Accounting"){
				
				$("input[name='Add']").show();
				$("#txtdeletex").show();
				$("#txtdelete1").show();
				$("#txtdelete2").show();
				$("#txtposting").show();
			
			}else if(closingStatus=="FALSE" && $("#statusnya").val() == "Review Accounting"){
				
				$("input[name='Add']").show();
				$("#txtdeletex").show();
				$("#txtdelete1").show();
				$("#txtdelete2").show();
				$("#txtposting").show();
			
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				
				$("input[name='Add']").hide();
				$("#txtdeletex").hide();
				$("#txtdelete1").hide();
				$("#txtdelete2").hide();
				$("#txtposting").hide();
				
			}
			
		});
		
	} 
}

//validasi lock inputan dri H+3
function get_Allowentry()
{
    var tglbkm = $('#orderdate').val();
    var postingflg_ = $('#postingflg_').val();
    if (tglbkm != "") {
        $.post("datadashboard.asp?tglbkm=" + tglbkm + "&idview=3", function(data) {
            $('#bodytable').empty();
            var json  = JSON.parse(data);
            var closingStatus = json.results[0].stock_closing;
			console.log(json.results[0]);
            if(closingStatus=="FALSE"){
                if(postingflg_==1){
                    $("input[name='Add']").hide();
                    $("#txtdeletex").hide();
                    $("#txtdelete1").hide();
                    $("#txtdelete2").hide();
                    $("#txtposting").hide();
                }else{
                    $("input[name='Add']").show();
                    $("#txtdeletex").show();
                    $("#txtdelete1").show();
                    $("#txtdelete2").show();
                    $("#txtposting").show();
                }
            }else{
                alert("PERIODE SUDAH DI CLOSE !!!");
                $("input[name='Add']").hide();
                $("#txtdeletex").hide();
                $("#txtdelete1").hide();
                $("#txtdelete2").hide();
                $("#txtposting").hide();
            }

            // ✅ Validasi hanya jika orderdate sudah diisi oleh user (bukan saat page load)
            if (!tglbkm || tglbkm === "") return;

            var selectEl     = document.getElementById("transporttaskno");
            var selectedText = selectEl ? selectEl.options[selectEl.selectedIndex].text : "";

            // ✅ Validasi hanya jika Transport Task No sudah dipilih
            if (!selectedText || selectedText.trim() === "") return;

            var matchTanggal = selectedText.match(/Tanggal\s*:\s*(\d{2})\/(\d{2})\/(\d{4})/i);
            if (!matchTanggal) return;

            var ttMonth = parseInt(matchTanggal[1], 10);
            var ttDay   = parseInt(matchTanggal[2], 10);
            var ttYear  = parseInt(matchTanggal[3], 10);

            var odParts = tglbkm.split("/");
            if (odParts.length < 3) return;

            var odMonth = parseInt(odParts[0], 10);
            var odDay   = parseInt(odParts[1], 10);
            var odYear  = parseInt(odParts[2], 10);

            // ✅ Validasi hanya jika tahun sudah 4 digit (user selesai mengetik)
            if (odYear < 1000) return;

            var orderDate = new Date(odYear, odMonth - 1, odDay);
            var ttDate    = new Date(ttYear, ttMonth - 1, ttDay);

            // Kondisi 1: Bulan atau Tahun berbeda
            if (ttMonth !== odMonth || ttYear !== odYear) {
                alert("Peringatan!\n\n" +
                      "Bulan Transport Task No : " + ttMonth + "/" + ttYear + "\n" +
                      "Bulan Order Date        : " + odMonth + "/" + odYear + "\n\n" +
                      "Bulan Order Date harus sama dengan bulan Transport Task No!");
                $('#orderdate').val("");
                return;
            }

            // Kondisi 2: Order Date lebih kecil dari tanggal Transport Task No
            if (orderDate < ttDate) {
                alert("Peringatan!\n\n" +
                      "Order Date (" + tglbkm + ") " +
                      "tidak boleh lebih kecil dari\n" +
                      "tanggal Transport Task No (" +
                      matchTanggal[1] + "/" + matchTanggal[2] + "/" + matchTanggal[3] + ")!");
                $('#orderdate').val("");
				location.reload();
                return;
            }

        });
    } 
}
</script>
</body>
</html>
