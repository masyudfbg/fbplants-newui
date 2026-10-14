<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="qtyOnHandFunction.inc"-->
<html>
<head>
<title>ePlaS - Transport Task Divisi</title>
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

function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}
</script>
<script language="javascript" type="text/javascript"> 
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

    function showHint(str)
		{
	var xmlhttp;
	var StockName;

		if (str=="")
		  {
		  document.getElementById("announ").innerHTML="";
		  document.getElementById("StockName").innerHTML="";
		  return;
		  }
		if (window.XMLHttpRequest)
		  {// code for IE7+, Firefox, Chrome, Opera, Safari
		  xmlhttp=new XMLHttpRequest();
		  StockName=new XMLHttpRequest();
		  }
		else
		  {// code for IE6, IE5
		  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
		  StockName=new ActiveXObject("Microsoft.XMLHTTP");
		  
		  }
		xmlhttp.onreadystatechange=function()
		  {
		  if (xmlhttp.readyState==4 && xmlhttp.status==200)
			{
			document.getElementById("announ").innerHTML=xmlhttp.responseText;
			}
		  }
		xmlhttp.open("GET","ajax/qualityAjax.asp?ajaxUoM=quality&q="+str,true);
		xmlhttp.send();
		
		StockName.onreadystatechange=function()
			  {
			  if (StockName.readyState==4 && StockName.status==200)
				{
				document.getElementById("StockName").innerHTML=StockName.responseText;
				}
			  }
			StockName.open("GET","ajax/qualityAjax.asp?ajaxUoM=StockName&q="+str,true);
			StockName.send();
			
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
 '-------------awal-----------------------
 Dim STORECODE,OrderNo,cmdView,PoCode,OrderDate,ITEMCODE,QuantityTransfered,Price,AMOUNT, LineNoHist
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

Dim detail_remarks : detail_remarks = request("detail_remarks")


 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 
  OrderNo = request("OrderNo")
 If Len(Trim(OrderNo)) = 0 Then OrderNo = "(NEW)"
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("OrderDate")<>"" then
   OrderDate   =Request("OrderDate")
 else
   OrderDate   =textDate(date)
 end if
 
 
 '---- Data Details
 Remarks = Request("Remarks")
 REQUESTORBY = Request("REQUESTORBY")
 Reference = Request("Reference")
 ITEMCODE=Request("ITEMCODE")
 STORECODEFROM = Request("STORECODEFROM")
 STORECODETO   = Request("STORECODETO")
 LineNo=Request("LineNo")
 QuantityTransfered=Request("QuantityTransfered")
 Linenohist = Request("linenohist")
 
 
 q = "select t.fccompanycode from Businessunit t  where t.fccode='" & Session("fcBA") & "'"
 
 Set rsCompany = Server.CreateObject("ADODB.Recordset")
 rsCompany.Open q, connstr
 IF Not rsCompany.eof then
	CompanyCode = rsCompany("fccompanycode")
 Else
    CompanyCode = "NONE"
 End if
 rsCompany.Close
 set rsCompany = Nothing	

 'QUERY GET AUTHORIZE
dim isApproved,rsApproved
set rsApproved = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'TD' and at.fccode='510' " & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND '"& session("CodeName") &"' = '" & session("CodeName") & "'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & session("fcba") & "'"
' response.write q
rsApproved.Open q, connstr
if Not rsApproved.Eof then
	isApproved =  cdbl(to_nol(rsApproved("total")))
else
	isApproved = 0
End if
rsApproved.Close : set rsApproved = Nothing 

q = "Select ORDERNO, FCBA, ORDERDATE, REMARKS, ROWSTATEID, " & _
" DOCUMENT_CLASSIFICATION, POSTINGFLG " & _
" From   TRANSPORTTASK_TAB " & _
" WHERE ORDERNO='" & OrderNo & "' "
 
  
 '--- Action ke Database di Oracle [ insert,edit,delete ]
Cek_IsAccClosing = StatusAccClosingTgl(OrderDate,fcba)  
if 	Cek_IsAccClosing = "FALSE" then  
 	if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "StartAuthorize" or Request("txtJenis")= "Approved" Then
		if Request("txtJenis")= "Header" then
			connstr.BeginTrans
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), OrderDate)
			If Cek_IsDailyClosing = "FALSE" then
				If err.number=0 then
					q="Delete from TransportTaskDiv_Tab where  OrderNo='" & OrderNo & "' and fcBA='" & Session("fcBA") & "'"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry
					if err.number  = 0 then
						connStr.CommitTrans
					end if
				end if
			else
				err.number = 1001
				err.description = "Tanggal """ & OrderDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If
			cmdView	= "True"
		elseif Request("txtJenis")= "Details" then
			connstr.BeginTrans
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), OrderDate)
			If Cek_IsDailyClosing = "FALSE" then

					
					q="Delete from TransportTaskDivDtl_tab where  OrderNo='" & OrderNo & "' and LineNo=" & LineNo & " and fcBA='" & Session("fcBA") & "'"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry
					if err.number  = 0 then
					connStr.CommitTrans
					end if
			else
				err.number = 1001
				err.description = "Tanggal """ & OrderDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If

			cmdView	= "True"
		elseif Request("txtJenis")="New" then
					if STORECODEFROM = STORECODETO Then
						err.number = 1001
						err.description = "<br/>Lokasi Store tidak boleh sama !!"
					end if
					
					connstr.BeginTrans	  
		
					If Ucase(OrderNo) = Ucase("(NEW)") Then
						oSql = " select nvl(controlnumber_api.get_nextnumber(upper('TRANSPORTTASKDIV')" & _
							"                                         , " & year(textDate(OrderDate)) & "" & _
							"                                         , " & month(textDate(OrderDate)) & "" & _
							"                                         , '" & Session("FCBA") & "'" & _
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
					
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), OrderDate)
			If Cek_IsDailyClosing = "FALSE" then
				Dim rsPO
				Set rsPO = Server.CreateObject("ADODB.Recordset")
				rsPO.Open "Select OrderNo from TransportTaskDiv_Tab where OrderNo='" & OrderNo & "'AND FCBA = '" & Session("FCBA") & "'", connstr
				IF rsPO.eof then	    
				
					q="Insert into TransportTaskDiv_Tab (OrderNo, OrderDate,requestorby, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME,rowstate) " & _
					"values ('" & OrderNo & "', to_date('" & OrderDate & "','mm/dd/yyyy') " & _
					",'" &  REQUESTORBY & "','" &  REMARKS & "','" &  Session("CodeName") & "','" & Session("CodeName") & "','" & _
					Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'),'Planned')"
					connStr.execute(q)
					modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
					executeQry	= "Query executed:" & q
					ExecuteAndLog modulFile,executeQry
				End if
				rsPO.Close

							' Insert Transport Task Before Approved
							strLineMax =   " select nvl(max(LineNo),0) maxLineNo from TransportTaskDivDTL_Tab " & _
											" where " & _
											" fcba = '" & Session("fcBA") & "'"
							
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

							q=  "Insert into  TransportTaskDivDTL_Tab (ORDERNO, DOCUMENTNO, STOCKCODE,LINENO,LINENOHIST, QUANTITYTRANSFERED,STORECODEFROM,STORECODETO ," & _
								" REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
								" values ('" & OrderNo & "','" & DOCUMENTNO & "','" & ITEMCODE & "','" & tempMaxLineNo & "',''" & _
								", " & QuantityTransfered & ",'" & STORECODEFROM & "','" & STORECODETO & "','" & detail_remarks & "','" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
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
			end if
			cmdView	= "True"
			
		
		
		elseif Request("txtJenis")="Approved" then
			OrderNo = request("OrderNo")
			'ambil qty original di trtdetail
			q = "select h.orderno,h.ORDERDATE,d.STORECODEFROM,d.STORECODETO,d.STOCKCODE, d.QUANTITYTRANSFERED,d.LINENOHIST,d.LINENO " & _ 
			" from TRANSPORTTASKDIV_TAB h " & _ 
			" inner join TRANSPORTTASKDIVDTL_TAB d on h.ORDERNO = d.ORDERNO and h.FCBA = d.FCBA " & _ 
			" where h.ORDERNO='" & OrderNo & "' " & _ 
			" and h.fcBA='" & Session("FCBA") & "'"
			'Response.Write q

			Set rsTRTDivDetail = Server.CreateObject("ADODB.Recordset")
			rsTRTDivDetail.Open q, connstr

			dim arr_dtlrc_eof
			arr_dtlrc_eof =  true
			if not rsTRTDivDetail.eof then 
				arr_dtlrc = rsTRTDivDetail.getrows()
				arr_dtlrc_eof = false
			End IF
		
			rsTRTDivDetail.close : set rsTRTDivDetail = nothing
			if not arr_dtlrc_eof then 
				for x=0 to ubound(arr_dtlrc,2)

						txtLineNoHist 		= arr_dtlrc(6,x)
						QuantityTransfered 	= arr_dtlrc(5,x)
						ITEMCODE 			= arr_dtlrc(4,x)
						txtLineNoKey 		= arr_dtlrc(7,x)
						storecodefrom 		= arr_dtlrc(2,x)
						storecodeto 		= arr_dtlrc(3,x)
					
						' modify jun (19 juni 2024)
						dim txtKolomtransaksi,qryUpdatestockOnhand
						txtKolomtransaksi = "TRTDIV"
						qTotalRecordOnHand  =   "select count(*) totalRecordOnHand from stockonhand " & _
												"where " & _
												"storecode = '" & STORECODEFROM & "' and " & _
												"stockcode = '" & ITEMCODE & "' and " & _
												"fcba = '" &  Session("fcBA") & "'"
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
								qryinsertstockOnhand 	= cmdInsertStockOnHand(STORECODEFROM,ITEMCODE,QuantityTransfered,txtKolomtransaksi)
						Else	 
								' Update qty di tabel stock onhand
								qryUpdatestockOnhand 	= cmdUpdateStockOnHand(STORECODEFROM,ITEMCODE,QuantityTransfered,txtKolomtransaksi,"Insert")
								
						End If	

						' CEK DULU DI TABEL ONHAND UNTUK GUDANG DIVISI (STORECODE TO)
						qTotalRecordOnHand  =   "select count(*) totalRecordOnHandTo from stockonhand " & _
												"where " & _
												"storecode = '" & STORECODETO & "' and " & _												 
												"stockcode = '" & ITEMCODE & "' and " & _
												"fcba = '" &  Session("fcBA") & "'"
						'Response.Write qTotalRecordOnHand
						set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
						rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
						if not rsTotalRecordOnHand.eof then
								tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHandTo"))
						else
								tempTotalRecordOnHand = 0
						end if 
						
							
						If tempTotalRecordOnHand = 0 then
							
								' Tambahkan item stock dan qty di tabel stock onhand
								qryinsertstockOnhand 	= cmdInsertStockOnHandStoreTo(STORECODETO,ITEMCODE,QuantityTransfered,txtKolomtransaksi)
						Else	 
								' Update qty di tabel stock onhand
								qryUpdatestockOnhand 	= cmdUpdateStockOnHandStoreTo(STORECODETO,ITEMCODE,QuantityTransfered,txtKolomtransaksi,"Insert")
						End If	

						'CEK ITEM DAHULU DI TABEL STOCKLEDGER
						'JIKA TIDAK ADA INSERT BARU, SELAIN ITU UPDATE
						dim qryStockLedger, rsStockLedger,temp,totalRecordStockLedger
						qryStockLedger = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _ 
						" WHERE " & _ 
						" STORECODE = '" & STORECODEFROM & "' and " & _
						" ITEMCODE  = '" & ITEMCODE & "' and " & _ 
						" year  	= to_char(to_date('" & OrderDate & "','mm/dd/yyyy'),'yyyy') and " & _ 
						" month 	= to_char(to_date('" & OrderDate & "','mm/dd/yyyy'),'mm') and " & _ 
						" FCBA 		= '" & Session("fcBA") & "' "
						
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
							qryinsertstockledger 	= cmdInsertStockLedger(STORECODEFROM,ITEMCODE,OrderDate,QuantityTransfered, "TRTDIV")
						ELSE
						
							'update stockledger
							qryupdateStockLedger 	= cmdUpdateStockLedger(STORECODEFROM,ITEMCODE,OrderDate,QuantityTransfered, "TRTDIV", "Insert")
						END IF

						'CEK DULU DI TABEL STOCKLEDGER UNTUK KODE GUDANG DIVISI (STORE FROM TO)
						qryStockLedger = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _ 
						" WHERE " & _ 
						" STORECODE = '" & STORECODETO & "' and " & _
						" ITEMCODE  = '" & ITEMCODE & "' and " & _ 
						" year  	= to_char(to_date('" & OrderDate & "','mm/dd/yyyy'),'yyyy') and " & _ 
						" month 	= to_char(to_date('" & OrderDate & "','mm/dd/yyyy'),'mm') and " & _ 
						" FCBA 		= '" & Session("fcBA") & "' "
						
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
							qryinsertstockledger 	= cmdInsertStockLedgerStoreTo(STORECODETO,ITEMCODE,OrderDate,QuantityTransfered, "TRTDIV")
						ELSE
						
							'update stockledger
							qryupdateStockLedger 	= cmdUpdateStockLedgerStoreTo(STORECODETO,ITEMCODE,OrderDate,QuantityTransfered, "TRTDIV", "Insert")
						END IF

						
				next
			end if

			'update kolom rowstate dan postingflg
			strQueryStatus = "UPDATE TransportTaskDiv_Tab SET POSTINGFLG='1',rowstate='Approved',lastupdate=sysdate,fcedit='" & session("CodeName") & "' WHERE fcba = '" & Session("FCBA") & "' AND OrderNo = '" & OrderNo & "'"
			connStr.execute(strQueryStatus)
			modulFile 	= session("username") & " Execution at: " & Now() &  ScriptName 
			executeQry	= "Query executed:" & strQueryStatus
			ExecuteAndLog modulFile,executeQry

			cmdView	= "True"
	 	else
    		cmdView	= request("cmdView")
		end if 
	else
			cmdView	= request("cmdView")		    
	end if
else
    err.number = 1001
	err.description = "Accounting Period  " & ACCOUNTINGPERIOD & "/" & ACCOUNTINGYEAR & "  LOCKED !!!"
	cmdView	= request("cmdView")
End if 

 Dim rsTransportTask,txtrowstate
 set rsTransportTask = server.CreateObject("Adodb.recordset")
 q="select OrderNo,fcba,to_char(OrderDate,'MM/DD/YYYY') OrderDate,REMARKS,rowstate from TransportTaskDiv_Tab  " & _
  " where OrderNo='" & OrderNo & "' and " & _
  " fcBA='" & session("fcBA") & "'"
  
 rsTransportTask.Open q, connstr
 if Not rsTransportTask.Eof then
    OrderDate = rsTransportTask("OrderDate")
	REMARKS = rsTransportTask("REMARKS")
	txtrowstate = rsTransportTask("rowstate")
	Statusnya = rsTransportTask("rowstate")
 End if

 

 %>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-truck" aria-hidden="true"></i>Header Transport Task Divisi</div><% end if %>
 <form name="frmEntry" method="post" action="TransportTaskDiv.Asp">
	<table class="data" width="100%">
		<tr>
		  <td align="left">Order No</td>
		  <td align="left"><input type="text" maxlength="20" name="OrderNo" size="30" value="<%=OrderNo%>" autocomplete = "Off"><font color="#FF0000">&nbsp;*</font></td>
		</tr>
		<tr>
		   <td align="left">Date</td>
		   <td align="left"><input class="datepicker" type="text" maxlength="10" id="orderdate" name="orderdate" onchange="get_Allowentry()" size="10" value="<%=orderdate%>" >
		   <font color="#FF0000">&nbsp;*</font>
		</td>
		</tr>
		
		<tr> 
		    <td align="left">Requestor</td>
			<td align="left"><input type="text" maxlength="50" id="requestorby" name="requestorby" size="10" value="<%=requestorby%>" >
			<font color="#FF0000">&nbsp;*</font>
		</tr>
		<tr> 
		    <td align="left">Remarks</td>
			<td align="left">
				<textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea>
				<input type="hidden" value="<%=Statusnya%>" id="statusnya" />
			</td>
		</tr>
		
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"   name="Add" id="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('Print_transporttaskDiv.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&OrderNo=<%=OrderNo%>&OrderDate=<%=OrderDate%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		
		<% if txtrowstate = "Planned" OR txtrowstate = "" Or txtrowstate = "Review Accounting" then %>
			<a id="txtdelete1" href="TransportTaskDiv.Asp?OrderNo=<%=OrderNo%>&OrderDate=<%=OrderDate%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=OrderNo%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			<%Else%>
			<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
			<%End If%>

		<% if txtrowstate = "Planned" Or txtrowstate = "Review Accounting" then %>
		<% if isApproved <> 0 then %>
		      <a id="poshdr" href="?fcCode=<%=fcCode%>&TransportTaskDate=<%=OrderDate%>&fcba=<%=fcba%>&txtJenis=Approved&orderno=<%=OrderNo%>&txtrowstate=<%=txtrowstate%>" onClick="return confirm('Are you sure to Approved transport Task Divisi this data :  <%=fcCode%> ?');">Approved</a>
		<%Else%>
			<%IF txtrowstate = "Approved" then %>
				<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
			<%End If%>
		<%End If%>
		<%End If%>
		
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(OrderNo)) > 4 Then
   Dim rsPurchaseOrder,intRecordsPerPage,cColor
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q=" select ORDERNO,LINENO, STOCKCODE,stock_api.get_stockName(Stockcode,fcba) StockName, " & _
   	 " STORECODEFROM,STORECODETO,QUANTITYTRANSFERED,REMARKS,DOCUMENTNO,  " & _
     " LINENOHIST,FCENTRY, FCEDIT, FCIP, FCBA  from TransportTaskDivDTL_Tab  " & _
     " where OrderNo='" & OrderNo & "' and " & _
     " fcBA='" & session("fcBA") & "' order by orderno, lineno"
   'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="TransportTaskDiv.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<% if request("txtJenis")= "Edit" then %>
		<tr>
		    <td align="left">Item Code</td>
		    <td align="left">
		     <input type="text" onkeyup="showHint(this.value)" size="15" name="ITEMCODE" value=<%=ITEMCODE%> readonly style="background:#dddddd" >&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('popup_stock_new.asp?fcba=<%=session("fcBA")%>&IssueDate=<%=OrderDate%>','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font><span class="satuan_item" id="StockName">&nbsp;...</span></a>
			</td>
			<tr>
				<td align="left">Item Stock</td><td align="left">
					<input id="stockitem" onkeyup="showHintStock(this.value)" type="text" size="15" name="stockitem" style="background:#dddddd; text-align: right" readonly>
					<span class="satuan_item" id="StokBarang" style="color: white">0</span>
			</tr>
		</tr>	
		<%else%>
		<tr>
		    <td align="left">Item Code</td>
		    <td align="left">
		     <input  type="text" onkeyup="showHint(this.value)" size="15" id="itemcode" name="itemcode" readonly onchange="javascript:getPrice(); " value="<%=itemcode%>" style="background:#dddddd" >&nbsp;<font color="#FF0000">* <a href="javascript:launchwin('popup_stock_new.asp?fcba=<%=session("fcBA")%>&IssueDate=<%=OrderDate%>','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font><span class="satuan_item" id="StockName">&nbsp;...</span></a>
			</td>
			<tr>
				<td align="left">Item Stock</td><td align="left">
					<input id="stockitem" onkeyup="showHintStock(this.value)" type="text" size="15" name="stockitem" style="background:#dddddd; text-align: right" readonly>
					<span class="satuan_item" id="StokBarang" style="color: white">0</span>
			</tr>
		</tr>
		<%end if%>
		
		<tr>
			<td align="left">Store From
			</td>
			<td align="left">      
				<select name="storecodefrom" id="storecodefrom" style="width: 18%" onchange="getValidasiGudang()">
				<option></option>
				<%
				qLov = " select fccode, fcname from store where fcba = '"& session("fcba") &"' and activation = 'Y'"
	   
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
			<td align="left">Store To
			</td>
			<td align="left">      
				<select name="storecodeto" id="storecodeto" style="width: 18%" onchange="getValidasiGudang()">
				<option></option>
				<%
				qLov = " select fccode, fcname from store where fcba = '"& session("fcba") &"' and activation = 'Y'"
	   
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
		
		<tr id="show_documentno" style="display: table-row;"  >
		    <td align="left">Document No</td>
		    <td align="left">
		     <input type="text" size="32" id="documentno" name="documentno" value="<%=documentno%>" "> 
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
					kuantitas = QUANTITY
				end if
			%>
			  <input type="number" size="10" name="QuantityTransfered" step="0.01" min="0" max="999999" value="<%=kuantitas%>" id="quantity" onkeyup="hitungrumus();" onkeypress="return validateFloatKeyPress(this,event);">&nbsp;<font color="#FF0000">*
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
		 
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <!-- Period COntrol In Here -->
			<%IF Request("txtJenis")="Edit" then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit" name="Add" value="Update" id="submitdata">
			<%Else%>
				
				   <input type="hidden" name="txtJenis" value="New">
				   <input Type="Submit" name="Add" value="Add" id="submitdata">
			<%End if%>
			<!-- End Period COntrol In Here -->
			  <input Type="Reset" name="Reset" value="Reset">
			  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="">&nbsp;</TD>	
	<TD align="left">Order No</TD>
	<TD align="left">Line No</TD>	
	<TD align="left">Stock Code</TD>
	<TD align="left">Stock Name</TD>
    <TD align="left">Store Code From</TD>	
	<TD align="left">Store Code To</TD>	
	<TD align="left">Quantity Transfered</TD>	
	<TD align="left">Remarks</TD>	
	</tr>
	</thead>
	<tbody>   
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
		
		  
		  %>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
				<% if txtrowstate = "Planned" Or txtrowstate = "Review Accounting" then %>
					<%if cdbl(to_nol(rsPurchaseOrder("QuantityTransfered")))<> 0 then%>
						
						<td valign="middle">
							<a id="txtdelete2" href="TransportTaskDiv.Asp?OrderNo=<%=rsPurchaseOrder("OrderNo")%>&LineNo=<%=rsPurchaseOrder("LineNo")%>&ItemCode=<%=rsPurchaseOrder("StockCode")%>&StoreCodeFrom=<%=rsPurchaseOrder("StoreCodeFrom")%>&StoreCodeto=<%=rsPurchaseOrder("StoreCodeto")%>&Linenohist=<%=rsPurchaseOrder("Linenohist")%>&OrderDate=<%=OrderDate%>&REMARKS=<%=REMARKS%>&rowstateid=<%=rowstateid%>&txtJenis=Details">
							<img id="txtdelete3" src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
						</td>
					<%End if%>
				<%Else%>
					<td colspan="1"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
				<%End if%>
		    <%for i = 0 to nCountofField-6%>
			    <td align="left" title="<%=rsPurchaseOrder(i).Name%>" valign="top"><%=rsPurchaseOrder(i)%></td>
		    <%next%>	    
		  </tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("<tr><td colspan=""10"" class=""fbp-doc-empty"">Data is empty</td></tr>")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
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
document.ready(function() {
	document.getElementById("show_documentno").style.display="none";
	document.getElementById("show_lineby").style.display="none";
	 	$("#priceby").click(function() {
	  	if( $(this).val() === "OPB") {
	   		document.getElementById("show_documentno").style.display="none";
			document.getElementById("show_lineby").style.display="none"; 
	    } else {      
	    	document.getElementById("show_documentno").style.display="table-row";
			document.getElementById("show_lineby").style.display="table-row"; 
			document.getElementById("amount").value = 0;
			document.getElementById("documentno").value ="";
			document.getElementById("lineby").value ="";
	    }
	  }); 

}); 

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

function hitungrumus() {
	var squantity 	= parseFloat($("#quantity").val().replace(/,/g, ''));
	var stok 		= parseFloat($('#StokBarang').text().replace(/,/g, ''));

	console.log('INPUT : ', squantity);
	console.log('STOK : ', stok);
	console.log('HASIL : ', stok - squantity);

	var cekInput 	= $('#Add').val();

	if(cekInput == "Select" && squantity > 0 && stok > 0){
		if (squantity > stok) {
			alert("Jumlah melebihi stok (" + stok + ")");
			$("#quantity").val(0);
		}
	}else if(cekInput == "Select" && squantity > 0 && stok == 0){
		$("#quantity").val(0);
	}
}

$(function() {	
	get_Allowentry();
	
});

//function ini digunakan untuk lock inputan berdasarkan periode akunting
function get_Allowentry_Acc()
{
	var orderdate = $('#orderdate').val();
	if (orderdate != "") {
		$.post("acc_allow_entry.asp?tglbkm=" + orderdate, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE" && $("#statusnya").val() != "Review Accounting"){
				$("input[name='Add']").show();
				$("#txtdelete1").show();
				$("#txtdelete2").show();
				$("#txtdelete3").show();
				$("#txtdelete4").show();
				
			}else if(closingStatus=="FALSE" && $("#statusnya").val() == "Review Accounting"){
				$("input[name='Add']").show();
				$("#txtdelete1").show();
				$("#txtdelete2").show();
				$("#txtdelete3").show();
				$("#txtdelete4").show();
				
			}else{
				alert("PERIODE SUDAH DI CLOSE !!! ");
				$("input[name='Add']").hide();
				$("#txtdelete1").hide();
				$("#txtdelete2").hide();
				$("#txtdelete3").hide();
				$("#txtdelete4").hide();
				

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
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE" && $("#statusnya").val() != "Review Accounting"){
				if(postingflg_==1){
					$("input[name='Add']").hide();
					$("#txtdelete1").hide();
					$("#txtdelete2").hide();
					$("#txtdelete3").hide();
					$("#txtdelete4").hide();
					$("#poshdr").hide();
				}else{
					$("input[name='Add']").show();
					$("#txtdelete1").show();
					$("#txtdelete2").show();
					$("#txtdelete3").show();
					$("#txtdelete4").show();
					$("#poshdr").show();
				}
			}else if(closingStatus=="TRUE" || closingStatus=="TRUE" && $("#statusnya").val() == "Review Accounting"){
				$("input[name='Add']").show();
				$("#txtdelete1").show();
				$("#txtdelete2").show();
				$("#txtdelete3").show();
				$("#txtdelete4").show();
				$("#poshdr").show();
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$("input[name='Add']").hide();
				$("#txtdelete1").hide();
				$("#txtdelete2").hide();
				$("#txtdelete3").hide();
				$("#txtdelete4").hide();
				$("#poshdr").hide();
			}
			
		});
		
	} 
}

function getValidasiGudang()
{
	var storecodefrom = $('#storecodefrom').val();
	var storecodeto = $('#storecodeto').val();
	if(storecodefrom==storecodeto){
		alert("STORE FROM  DAN STORE TO, TIDAK BOLEH SAMA..!!");
		$('#storecodeto').val("");

	}
}
</script>
</body>
</html>