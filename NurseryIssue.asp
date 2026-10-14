<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Nursery Issue</title>
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
<script language="javascript" type="text/javascript">

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

// launched from button click 
var getTragettype = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("jobcode").value;
	var PostStr = 'jobcode='+jobcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetPVInfo.asp', 'POST', PostStr, showTragettype);
	
}

// The function for handling the response from the server
var showTragettype = function (oXML) { 
     // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(response);
    // update the Div to show the result from the server
    document.getElementById("locationtype").value = s.target; 
	// document.getElementById("descpr").value = s.description; 
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
 <!-------------awal!------------->
<%


Function cmdCheckHistNursery(initfcba, initNURSERYCODE, initLineNoHist)
	dim qCheckHistNursery,rsCheckHistNursery,qTotalDataNursery
	
	qCheckHistNursery = " select count(1) totalData" & _
						" from stocktranshist" & _
						" where fcba = '" & initfcba & "'" & _
						" and sourcecode = '" & initNURSERYCODE & "'" & _
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

Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, ITEMCODE
DIM v_NURSERYCODE,v_ISSUEDATE, v_JOBCODE, v_LOCATIONTYPE, v_LOCATIONCODE, v_REFERENCE, v_REMARKS, Linenohist, PeriodDate
DIM v_QUANTITY
Dim cmdView
Dim qNurseryInfo, rsNurseryInfo, tempStoreCode, tempStockCode, doneError
Dim qLineMax, rsqLineMax, tempMaxLine
Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
Dim qLessZero, rsLessZero, tempislesszero
Dim done
Dim qLineUpper, rsqLineUpper, totalUpper, qQuantity, rsqQuantity, totalQuantity, QuantityOriginal,qDeleteHist
Dim rsReceiveDetail, txtLineNoHist, originalQuantity, v_CrossCekDetail

Dim txtDocumentNo : txtDocumentNo = Request("txtDocumentNo")
Dim CekLineHistNursery : CekLineHistNursery = 100
Dim rsMaxLineNoKey, strMaxLineNoKey, tempMaxLineNoKey
Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")
Dim QuantityOnHandNursery : QuantityOnHandNursery = 0
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""


v_NURSERYCODE = request("v_NURSERYCODE")
v_ISSUEDATE = request("v_ISSUEDATE")
v_JOBCODE = request("v_JOBCODE")
v_QUANTITY = request("v_QUANTITY")
v_LOCATIONTYPE = request("v_LOCATIONTYPE")
v_LOCATIONCODE = request("v_LOCATIONCODE")
v_REFERENCE = request("v_REFERENCE")
v_REMARKS = request("v_REMARKS")
Linenohist = Request ("linenohist")

if Request("txtJenis")="New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
		connstr.BeginTrans
		
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), v_ISSUEDATE)
		If Cek_IsDailyClosing = "FALSE" then

			q	=	" delete from nurseryissuehdr" & _
					" where documentno = '" & txtDocumentNo & "'" & _
					"       and fcba = '" & Session("fcBA") & "'"

			'Response.Write q
			connStr.execute(q)
			
			if err.number = 0 then
			   connStr.CommitTrans
			end if		
		else
			err.number = 1001
			err.description = "Tanggal """ & v_ISSUEDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then
	    connstr.BeginTrans
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), v_ISSUEDATE)
		If Cek_IsDailyClosing = "FALSE" then
			qNurseryInfo =  " select" & _
							" storecode" & _
							" , stockcode" & _
							" from nursery where" & _
							" nurserycode = '" &  v_NURSERYCODE & "'" & _
							" and fcba = '"& Session("fcBA") & "'"
					
			set rsNurseryInfo = Server.CreateObject("ADODB.Recordset")
			rsNurseryInfo.Open qNurseryInfo, connstr
			if not rsNurseryInfo.eof then
				tempStoreCode = rsNurseryInfo("storecode")
				tempStockCode = rsNurseryInfo("stockcode")				
			else
				tempStoreCode = ""
				tempStockCode = ""
			end if 
			
			If tempStoreCode = "" Then
				err.number = 1001
				err.description = "Storecode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
			Else
				If tempStockCode = "" Then
					err.number = 1001
					err.description = "Stockcode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
				Else
					q =	" Select linenohist" & _
						"        , quantity" & _
						" from nurseryissuedtl" & _
						" where documentno = '" & txtDocumentNo & "'" & _
						"       and NURSERYCODE = '" & v_NURSERYCODE & "'" & _
						"       and linenokey = " & txtLineNoKey & "" & _
						"       and fcBA = '" & Session("fcBA") & "'"
			
					
					Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					rsReceiveDetail.Open q, connstr
					
					IF Not rsReceiveDetail.eof then
						txtLineNoHist = rsReceiveDetail("LINENOHIST")
						originalQuantity = cdbl(rsReceiveDetail("QUANTITY"))
					End if
					rsReceiveDetail.Close
					set rsReceiveDetail = Nothing	
					
					
					' xxxxxxxxxxxxxxxxxxxxxxx
					CekLineHistNursery = cmdCheckHistNursery(Session("fcBA"), v_NURSERYCODE, txtLineNoHist)
					
						qLessZero = " select islesszero" & _
									" from stock" & _
									" where fccode = '" & tempStockCode & "'" & _
									"       and fcba = '" & session("fcBA") & "'" 
						
						set rsLessZero = Server.CreateObject("ADODB.Recordset")
						rsLessZero.Open qLessZero, connstr
						if not rsLessZero.eof then
							tempislesszero = rsLessZero("islesszero")
						else
							tempislesszero = "FALSE"
						end if
						
						qQuantityOnHand = " select quantity" & _
										   " from stockonhandnursery" & _
										   " where storecode = '" & tempStoreCode & "'" & _
										   "       and stockcode  = '" & tempStockCode & "'" & _
										   "       and nurserycode  = '" & v_NURSERYCODE & "'" & _
										   "       and fcba = '" & session("fcBA") &"' "
										  
						set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
						rsQuantityOnHand.Open qQuantityOnHand, connstr
						if not rsQuantityOnHand.eof then
							QuantityOnHandNursery = cdbl(rsQuantityOnHand("quantity"))
						else
							QuantityOnHandNursery = 0
						end if
						
						If tempislesszero = "TRUE" Or (tempislesszero = "FALSE" And ((QuantityOnHandNursery + originalQuantity) >= 0)) Then

							q =	" Delete from stocktranshist" & _
								" where lineno = " & txtLineNoHist & "" & _
								"       and fcba = '" & Session("fcBA") & "'"
						
							connStr.execute(q)
							
							q = " update stockonhand" & _
								" set quantity = quantity + " & originalQuantity & "" & _
								"     , LASTUPDATE = sysdate " & _
								"     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode = '" & tempStockCode & "'" & _
								"       and fcba = '" & session("fcBA") &"'"

							connStr.execute(q)
							
							q = " update stockonhandnursery " & _
								" set quantity = quantity + " & originalQuantity & "" & _
								"     , LASTUPDATE = sysdate " & _
								"     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode  = '" & tempStockCode & "'" & _ 
								"       and nurserycode  = '" & v_NURSERYCODE & "'" & _ 
								"       and fcba = '" & session("fcBA") &"'"

							connStr.execute(q)
							
							q = " Delete from nurseryissuedtl" & _
								" where documentno = '" & txtDocumentNo & "'" & _
								"       and NURSERYCODE = '" & v_NURSERYCODE & "'" & _
								"       and linenokey = " & txtLineNoKey & "" & _
								"       and fcBA = '" & Session("fcBA") & "'"
							connStr.execute(q)
						Else
							err.number = 1001
							Err.Description = "Quantity tidak boleh Negatif atau Saldo belum didefinisikan"
						End If
				End If
			End If
			
			if err.number = 0 then
				connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & v_ISSUEDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
	elseif Request("txtJenis")="New" then
		' After Insert the Nursury Culling then system will setup Inventory Transaction History and Stock On Hand 
		' the stock will be picked from nursery information
        connstr.BeginTrans
		
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), v_ISSUEDATE)
		If Cek_IsDailyClosing = "FALSE" then
		
			Dim rsNurseryIssueHdr
			Set rsNurseryIssueHdr = Server.CreateObject("ADODB.Recordset")
			rsNurseryIssueHdr.Open " select documentno from nurseryissuehdr where documentno = '" & txtDocumentNo & "' and fcba = '" & Session("FCBA") & "'", connstr
			If rsNurseryIssueHdr.EOF Then
				q = " insert into nurseryissuehdr" & _
					" (documentno, issuedate, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) values (" & _
					" '" & txtDocumentNo & "'" & _
					" , to_date('" & v_ISSUEDATE & "', 'mm/dd/yyyy')" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & Session("fcba") & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"
				connStr.execute(q)
			End If
			rsNurseryIssueHdr.Close
			set rsNurseryIssueHdr = Nothing

			qNurseryInfo = 	" select storecode" & _
							"        , stockcode" & _
							" from nursery" & _
							" where nurserycode = '" &  v_NURSERYCODE & "'" & _
							"       and fcba = '"& Session("fcBA") & "'"
				
			set rsNurseryInfo = Server.CreateObject("ADODB.Recordset")
			rsNurseryInfo.Open qNurseryInfo, connstr
			if not rsNurseryInfo.eof then
				tempStoreCode = rsNurseryInfo("storecode")
				tempStockCode = rsNurseryInfo("stockcode")
			else
				tempStoreCode = ""
				tempStockCode = ""
			end if 
			
			If tempStoreCode = "" Then
				err.number = 1001
				err.description = "Storecode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
			Else
				If tempStockCode = "" Then
					err.number = 1001
					err.description = "Stockcode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
				Else
				
					qLessZero =	" select islesszero" & _
								" from stock" & _
								" where fccode = '" & tempStockCode & "'" & _
								"       and fcba = '" & session("fcBA") & "'" 
					
					set rsLessZero = Server.CreateObject("ADODB.Recordset")
					rsLessZero.Open qLessZero, connstr
					if not rsLessZero.eof then
						tempislesszero = rsLessZero("islesszero")
					else
						tempislesszero = "FALSE"
					end if
						
					qQuantityOnHand =	" select quantity" & _
										" from stockonhandnursery" & _
										" where storecode = '" & tempStoreCode & "'" & _
										"       and stockcode = '" & tempStockCode & "'" & _
										"       and nurserycode = '" & v_NURSERYCODE & "'" & _
										"       and fcba = '" &  session("fcBA") & "'"
									  
					set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
					rsQuantityOnHand.Open qQuantityOnHand, connstr
					if not rsQuantityOnHand.eof then
						tempQuantityOnHand = cdbl(rsQuantityOnHand("quantity"))
					else
						tempQuantityOnHand = 0
					end if
						
					' We have to Update the Quantity by Adding them with the New one.
					if (tempislesszero = "FALSE" And ((tempQuantityOnHand -  v_QUANTITY) >= 0)) or tempislesszero = "TRUE"  then 

						qLineMax = " select nvl(max(lineno),0) tempMaxLine" & _
								   " from stocktranshist" & _
								   " where fcba = '" & session("fcBA") & "'" 
							
						set rsqLineMax = Server.CreateObject("ADODB.Recordset")
						rsqLineMax.Open qLineMax, connstr
						if not rsqLineMax.eof then
							tempMaxLine = cdbl(rsqLineMax("tempMaxLine"))
						else
							tempMaxLine = 0
						end if 

						tempMaxLine = tempMaxLine + 1

						q = " Insert into StockTransHist (" & _
							" LINENO" & _
							" , FCBA" & _
							" , STOCKCODE" & _
							" , TRANSDATE" & _
							" , TRANSTYPE" & _
							" , STORECODE" & _
							" , SOURCETYPE" & _
							" , SOURCECODE" & _
							" , TARGETTYPE" & _
							" , TARGETCODE" & _
							" , QUANTITY" & _
							" , LASTUPDATE" & _
							" , LASTTIME" & _
							" ) values ("& _
							"" & tempMaxLine & "" & _
							" , '" & session("fcBA") & "'" & _
							" , '" & tempStockCode &"'" & _
							" , to_date('" & v_ISSUEDATE & "','mm/dd/yyyy')" & _
							" , 'MN_ISSUE'" & _
							" , '" & tempStoreCode & "'" & _
							" , 'MN'" & _
							" , '" & v_NURSERYCODE & "'" & _
							" , '" & v_LOCATIONTYPE & "'" & _
							" , '" & v_LOCATIONCODE & "'" & _
							" , " & -1 * v_QUANTITY & "" & _
							" , sysdate" & _
							" , To_Char(sysdate,'HH24:MI')" & _
							" )"
							  
						connStr.execute(q)
						
						strMaxLineNoKey = " select nvl(max(linenokey),0) tempMaxLineNoKey" & _
								   " from nurseryissuedtl" & _
								   " where fcba = '" & session("fcBA") & "'" 
							
						set rsMaxLineNoKey = Server.CreateObject("ADODB.Recordset")
						rsMaxLineNoKey.Open strMaxLineNoKey, connstr
						if not rsMaxLineNoKey.eof then
							tempMaxLineNoKey = cdbl(rsMaxLineNoKey("tempMaxLineNoKey"))
						else
							tempMaxLineNoKey = 0
						end if 

						tempMaxLineNoKey = tempMaxLineNoKey + 1
						
						q = " Insert into nurseryissuedtl" & _
							" (documentno, linenokey, nurserycode, jobcode, locationtype" & _
							" , locationcode, quantity, amount, reference, remarks" & _
							" , fcentry, fcedit, fcip, fcba, linenohist, lastupdate, lasttime) values (" & _
							" '" & txtDocumentNo & "'" & _
							" , '" & tempMaxLineNoKey & "'" & _
							" , '" & v_NURSERYCODE & "'" & _
							" , '" & v_Jobcode & "'" & _
							" , '" & v_LOCATIONTYPE & "'" & _
							" , '" & v_LOCATIONCODE & "'" & _
							" , '" & v_QUANTITY & "'" & _
							" , 0" & _
							" , '" & v_Reference & "'" & _
							" , '" & v_Remarks & "'" & _
							" , '" & Session("CodeName") & "'" & _
							" , '" & Session("CodeName") & "'" & _
							" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
							" , '" & Session("fcBA") & "'" & _
							" , " & tempMaxLine & "" & _
							" , sysdate" & _
							" , To_Char(sysdate,'HH24:MI')" & _
							" )"
						'response.Write q
						'response.Flush()
						connStr.execute(q)
						
						'Lihat Record di tabel stockonhand
						qTotalRecordOnHand =	" select count(*) totalRecordOnHand" & _
												" from stockonhand" & _
												" where storecode = '" & tempStoreCode & "'" & _
												"       and stockcode = '" & tempStockCode & "'" & _
												"       and fcba = '" &  session("fcBA") & "'"

						set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
						rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
						if not rsTotalRecordOnHand.eof then
							tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
						else
							tempTotalRecordOnHand = 0
						end if
						
						'Jika tidak ada
						If tempTotalRecordOnHand = 0 Then
							' Just Insert the New Record with the following/received Quantity
							q	=	" insert into stockonhand" & _
									" (storecode, stockcode, fcba, quantity, lastupdate, lasttime)" & _
									" values (" & _
									" '" & tempStoreCode & "'" & _
									" , '" & tempStockCode & "'" & _
									" , '" & session("fcBA") & "'" & _
									" , '" & (v_Quantity * -1) & "'" & _
									" , sysdate" & _
									" , To_Char(sysdate,'HH24:MI')" & _
									" )"

							connStr.execute(q)
						Else					
							q = " update stockonhand" & _
								" set quantity = quantity + " & (v_Quantity * -1) & "" & _
								"     , LASTUPDATE = sysdate " & _
								"     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode = '" & tempStockCode & "'" & _
								"       and fcba = '" &  session("fcBA") & "'"
								 
							connStr.execute(q)
						End If
						
						qTotalRecordOnHand =	" select count(*) totalRecordOnHand" & _
												" from stockonhandnursery" & _
												" where storecode = '" & tempStoreCode & "'" & _
												"       and stockcode = '" & tempStockCode & "'" & _
												"       and nurserycode = '" & v_NURSERYCODE & "'" & _
												"       and fcba = '" &  session("fcBA") & "'"

						set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
						rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
						if not rsTotalRecordOnHand.eof then
							tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
						else
							tempTotalRecordOnHand = 0
						end if
						
						If tempTotalRecordOnHand = 0 Then
							'Lihat Record di tabel stockonhandnursery
							q	=	" Insert into stockonhandnursery(" & _
									" STORECODE" & _
									" , STOCKCODE" & _
									" , NURSERYCODE" & _
									" , FCBA" & _
									" , QUANTITY" & _
									" , LASTUPDATE" & _
									" , LASTTIME" & _
									" ) values (" & _
									" '" & tempStoreCode & "'" & _
									" , '" & tempStockCode & "'" & _
									" , '" & v_NURSERYCODE & "'" & _
									" , '" & session("fcBA") & "'" & _
									" , '" & (v_Quantity * -1) & "'" & _
									" , sysdate" & _
									" , To_Char(sysdate,'HH24:MI')" & _
									" )"

							connStr.execute(q)
						Else
							q = " update stockonhandnursery" & _
								" set quantity = quantity + " & (v_Quantity * -1) & "" & _
								"     , LASTUPDATE = sysdate " & _
								"     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode = '" & tempStockCode & "'" & _
								"       and nurserycode = '" & v_NURSERYCODE & "'" & _
								"       and fcba = '" &  session("fcBA") & "'"
								 
							connStr.execute(q)
						End If
					Else
						err.number = 1001
						err.description = "Quantity pada batch """ & v_NURSERYCODE & """ saat ini hanya """ & tempQuantityOnHand & """ sehingga nilai quantity tidak boleh negatif"
					end if
				End If
			End If
				
			if err.number = 0 then
			   connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & v_ISSUEDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	
        connstr.BeginTrans	
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), v_ISSUEDATE)
		If Cek_IsDailyClosing = "FALSE" then
			qNurseryInfo =	" select storecode" & _
							"        , stockcode" & _
							" from nursery" & _
							" where nurserycode = '" &  v_NURSERYCODE & "'" & _
							"       and fcba = '"& Session("fcBA") & "'"
				
				
			set rsNurseryInfo = Server.CreateObject("ADODB.Recordset")
			rsNurseryInfo.Open qNurseryInfo, connstr
			if not rsNurseryInfo.eof then
				tempStoreCode = rsNurseryInfo("storecode")
				tempStockCode = rsNurseryInfo("stockcode")				
			else
				tempStoreCode = ""
				tempStockCode = ""
			end if 
			
			If tempStoreCode = "" Then
				err.number = 1001
				err.description = "Storecode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
			Else
				If tempStockCode = "" Then
					err.number = 1001
					err.description = "Stockcode pada tabel nursery masih kosong, mohon di definisikan terlebih dahulu"
				Else
				
					q =	" select linenohist" & _
						"        , quantity" & _
						" from nurseryissuedtl" & _
						" where documentno = '" & txtDocumentNo & "'" & _
						"       and nurserycode = '" & v_NURSERYCODE & "'" & _
						"       and linenokey = '" & txtLineNoKey & "'" & _
						"       and fcba = '" & Session("fcBA") & "'"
					'Response.Write q
					Set rsReceiveDetail = Server.CreateObject("ADODB.Recordset")
					rsReceiveDetail.Open q, connstr
					
					IF Not rsReceiveDetail.eof then
						txtLineNoHist = rsReceiveDetail("LINENOHIST")
						originalQuantity = Cdbl(rsReceiveDetail("QUANTITY"))
					End if
					rsReceiveDetail.Close
					set rsReceiveDetail = Nothing
				
					totalUpper = cmdCheckHistNursery(Session("fcBA"), v_NURSERYCODE, txtLineNoHist)

						qQuantity =	" select quantity" & _
									" from stocktranshist" & _
									" where lineno = " & txtLineNoHist & "" & _
									"       and fcba = '" & Session("fcBA") & "'"
						
						
						set rsqQuantity= Server.CreateObject("ADODB.Recordset")
						rsqQuantity.Open qQuantity, connstr

						if not rsqQuantity.eof then
							totalQuantity = cdbl(rsqQuantity("quantity"))
						else
							totalQuantity = 0
						end if 
						
						
						' Add by NAF to handle Inventory Stock On Hand
						' Insert the History of Received Inventory Transaction
						

						qLessZero =	" select islesszero" & _
									" from stock" & _
									" where fccode = '" & tempStockCode & "'" & _
									"       and fcba = '" & session("fcBA") & "'" 
						
						
						set rsLessZero = Server.CreateObject("ADODB.Recordset")
						rsLessZero.Open qLessZero, connstr
						if not rsLessZero.eof then
							tempislesszero = rsLessZero("islesszero")
						else
							tempislesszero = "FALSE"
						end if 

						
						qQuantityOnHand = " select quantity" & _
										  " from stockonhandnursery" & _
										  " where storecode = '" & tempStoreCode & "'" & _
										  "       and stockcode = '" & tempStockCode & "'" & _
										  "       and nurserycode = '" & v_NURSERYCODE & "'" & _
										  "       and fcba = '" &  session("fcBA") & "'"
							
						set rsQuantityOnHand = Server.CreateObject("ADODB.Recordset")
						rsQuantityOnHand.Open qQuantityOnHand, connstr
						if not rsQuantityOnHand.eof then
							QuantityOriginal = cdbl(rsQuantityOnHand("quantity"))
						else
							QuantityOriginal = 0
						end if 	
						
						
						if tempislesszero = "TRUE" or ( tempislesszero = "FALSE" And ((QuantityOriginal - totalQuantity - v_QUANTITY) >= 0)) then 
						
							qDeleteHist =	" delete stocktranshist" & _
											" where lineno = " & txtLineNoHist & "" & _
											"       and fcba = '" & Session("fcBA") & "'"

							connStr.execute(qDeleteHist)

							q = " Insert into StockTransHist (" & _
								" LINENO" & _
								" , FCBA" & _
								" , STOCKCODE" & _
								" , TRANSDATE" & _
								" , TRANSTYPE" & _
								" , STORECODE" & _
								" , SOURCETYPE" & _
								" , SOURCECODE" & _
								" , TARGETTYPE" & _
								" , TARGETCODE" & _
								" , QUANTITY" & _
								" , LASTUPDATE" & _
								" , LASTTIME" & _
								" ) values ("& _
								"" & txtLineNoHist & "" & _
								" , '" & session("fcBA") & "'" & _
								" , '" & tempStockCode &"'" & _
								" , to_date('" & v_ISSUEDATE & "','mm/dd/yyyy')" & _
								" , 'MN_ISSUE'" & _
								" , '" & tempStoreCode & "'" & _
								" , 'MN'" & _
								" , '" & v_NURSERYCODE & "'" & _
								" , '" & v_LOCATIONTYPE & "'" & _
								" , '" & v_LOCATIONCODE & "'" & _
								" , " & -1 * v_QUANTITY & "" & _
								" , sysdate" & _
								" , To_Char(sysdate,'HH24:MI')" & _
								" )"
								  
							connStr.execute(q)					

							q =	" Update nurseryissuedtl" & _
								" set QUANTITY = " & v_QUANTITY & "" & _
								"     , Reference = '" & v_Reference & "'" & _
								"     , Remarks = '" & v_Remarks & "'" & _
								"     , fcedit = '" & Session("CodeName") & "'" & _
								"     , lastupdate = sysdate " & _
								"     , lasttime = to_char(sysdate,'hh24:mi') " & _
								" where documentno = '" & txtDocumentNo & "'" & _
								"       and NURSERYCODE = '" & v_NURSERYCODE & "'" & _
								"       and linenokey = " & txtLineNoKey & "" & _
								"       and fcBA = '" & Session("fcBA") & "'"
							
							'Response.Write q
							connStr.execute(q)
							
							q = " update stockonhand" & _
								" set quantity = ( quantity - " & totalQuantity & " )" & - v_QUANTITY & "" & _
								"     , LASTUPDATE = sysdate " & _
								"     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode = '" & tempStockCode & "'" & _
								"       and fcba = '" &  session("fcBA") & "'"
								 
							connStr.execute(q)				
							
							q = " update stockonhandnursery" & _
								" set quantity = " & QuantityOriginal - totalQuantity - v_QUANTITY & _ 
								"     , LASTUPDATE = sysdate " & _
								"     , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
								" where storecode = '" & tempStoreCode & "'" & _
								"       and stockcode = '" & tempStockCode & "'" & _
								"       and nurserycode = '" & v_NURSERYCODE & "'" & _
								"       and fcba = '" &  session("fcBA") & "'"
								 
							connStr.execute(q)
										
							if err.number  = 0 then
								 connstr.CommitTrans	
							end if
						else
							err.number = 1001
							err.description = "Quantity pada batch """ & v_NURSERYCODE & """ saat ini hanya """ & QuantityOriginal - totalQuantity & """ sehingga nilai quantity tidak boleh negatif"
						end if
				End If
			End If
		else
			err.number = 1001
			err.description = "Tanggal """ & v_ISSUEDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"

    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Len(Trim(v_ISSUEDATE)) = 0 then
	v_ISSUEDATE = textDate(date)
else
   v_ISSUEDATE = Request("v_ISSUEDATE")
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")

Dim rsNurseryHeader, strNurseryHeader
set rsNurseryHeader = server.CreateObject("Adodb.recordset")
strNurseryHeader	=	" select documentno, to_char(issuedate,'mm/dd/yyyy') issuedate" & _
						" from nurseryissuehdr" & _
						" where documentno = '" & txtDocumentNo & "'" & _
						"       and fcba = '" & session("fcBA") & "'"

rsNurseryHeader.Open strNurseryHeader, connstr
if Not rsNurseryHeader.Eof then
	txtDocumentNo	=	rsNurseryHeader("documentno")
	v_ISSUEDATE	=	rsNurseryHeader("ISSUEDATE")
End if

rsNurseryHeader.Close
Set rsNurseryHeader = Nothing


if done = 2 then 
	connStr.RollbackTrans   
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>Terdapat Transaksi setelah data ini, lakukan delete terlebih dahulu untuk transaksi setelah ini</font>"
end if

if done = 1 then 
	connStr.RollbackTrans   
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>Nilai Quantity tidak boleh Negatif atau Saldo belum didefinisikan</font>"
end if
if err.number <> 0 then
	connStr.RollbackTrans   
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if


Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & v_ISSUEDATE & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & v_ISSUEDATE & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Nursery')"
	   
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

 <!-------------ISI Payment Form!------------->
 
	<form name="frmEntry" method="post" action="NurseryIssue.asp">
		<table class="data" width="100%">
			<tr>
				<td align="left">Issue No</td>
				<td align="left"><input type="text" maxlength="15" name="txtDocumentNo" size="30" value="<%=txtDocumentNo%>"></td>
			</tr>
			<tr>
				<td align="left" width="20%">Date Issue</td>
				<td align="left"><input class="datepicker" type="text" size="15" name="v_ISSUEDATE" value="<%=v_ISSUEDATE%>"></td>
			</tr>
			<tr>
				<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="PrintNurseryIssue.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&v_NURSERYCODE=<%=v_NURSERYCODE%>&v_ISSUEDATE=<%=v_ISSUEDATE%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" Then%>
					<a href="NurseryIssue.Asp?txtDocumentNo=<%=txtDocumentNo%>&v_ISSUEDATE=<%=v_ISSUEDATE%>&txtJenis=HeaderDelete">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=txtDocumentNo%>" border="0" onClick="return confirm('Are you sure to delete');">
					</a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				<!-- HADIR -->
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Nursery"
				dataSubmodule 	= "Issue"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
				</td>
			</tr>
		</table>
	</form>
<%
if cmdView = "True" and len(trim(txtDocumentNo)) > 0 Then
	Dim rsVH, intRecordsPerPage, cColor
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	= " select documentno" & _
		  "        , linenokey" & _
		  "        , nurserycode" & _
		  "        , jobcode" & _
		  "        , locationtype" & _
		  "        , locationcode" & _
		  "        , nvl(quantity, 0) quantity" & _
		  "        , nvl(amount, 0) amount" & _
		  "        , reference" & _
		  "        , remarks" & _
		  "        , linenohist" & _
		  " from nurseryissuedtl" & _
		  " where documentno = '" & txtDocumentNo & "'" & _
		  "       and fcba = '" & session("fcBA") & "'" & _
		  " order by linenokey"
	'response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
   <form action="NurseryIssue.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td align="left" width="20%">Nursery Code</td>			
			<td align="left">
				<%If Request("txtJenis")="Edit" Then%>
				<input type="text" maxlength="10" name="v_NURSERYCODE" size="10" value="<%=v_NURSERYCODE%>" readonly="True" style="background:#FFFFCC;">
				<%Else%>
				<input type="text" maxlength="10" name="v_NURSERYCODE" size="10" value="<%=v_NURSERYCODE%>">
				<%End If%>
			</td>			
		</tr>
		<tr>
			<td align="left" width="20%">Job Code</td>
			<td align="left">
				<%If Request("txtJenis")="Edit" Then%>
				<input type="text" size="15" id="jobcode" onchange="javascript:getTragettype();" name="v_JOBCODE" value="<%=v_JOBCODE%>" readonly="True" style="background:#FFFFCC;">&nbsp;<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Activity</font></a><font color="#FF0000"> *</font>
				<%Else%>
				<input type="text" size="15" id="jobcode" onchange="javascript:getTragettype();" name="v_JOBCODE" value="<%=v_JOBCODE%>">&nbsp;<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Activity</font></a><font color="#FF0000"> *</font>
				<%End If%>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Location Type</td>
			<td align="left">
				<%If Request("txtJenis")="Edit" Then%>
				<input type="text" size="15" id="locationtype" name="v_LOCATIONTYPE" value="<%=v_LOCATIONTYPE%>" readonly="True" style="background:#FFFFCC;" style="background-color: #eee;" readonly>&nbsp;<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a><font color="#FF0000"> *</font>
				<%Else%>
				<input type="text" size="15" id="locationtype" name="v_LOCATIONTYPE" value="<%=v_LOCATIONTYPE%>"style="background-color: #eee;" readonly>&nbsp;<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a><font color="#FF0000"> *</font>
				<%End If%>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Location Code</td>
			<td align="left">
				<%If Request("txtJenis")="Edit" Then%>
				<input type="text" size="15" name="v_LOCATIONCODE" value="<%=v_LOCATIONCODE%>" readonly="True" style="background:#FFFFCC;">
				<%Else%>
				<input type="text" size="15" name="v_LOCATIONCODE" value="<%=v_LOCATIONCODE%>">
				<%End If%>
			</td>
		</tr>
			<%If Request("txtJenis") = "Edit" Then%>
				<input type="hidden" size="15" name="txtLineNoKey" value="<%=txtLineNoKey%>" readonly="True" style="background:#FFFFCC">
			<%End if%>
		<tr>
			<td align="left" width="20%">Quantity</td>
			<td align="left">
				<input type="text" size="15" name="v_QUANTITY" value="<%=v_QUANTITY%>" maxlength="32">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Reference</td>
			<td align="left">
				<input type="text" size="15" name="v_Reference" value="<%=v_Reference%>" maxlength="50">
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Remarks</td>
			<td align="left">
				<input type="text" size="15" name="v_Remarks" value="<%=v_Remarks%>" maxlength="50">
			</td>
		</tr>
		<tr>
		<td align="left" colspan="2">
		  <input type="hidden" name="txtDocumentNo" value="<%=txtDocumentNo%>">
		  <input type="hidden" name="v_ISSUEDATE" value="<%=v_ISSUEDATE%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit" />
		       <input Type="Submit"  name="Add" value="Update" />
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New" />
		       <input Type="Submit"  name="Add" value="Add" />
		  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset" />
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)" />
		</td>
	  </tr>
   </table>
   
   <table class="data">
 	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">LINE NO (KEY)</TD>
	<TD align="left">NURSERY CODE</TD>
	<TD align="left">JOBCODE</TD>
	<TD align="left">TARGET TYPE</TD>
	<TD align="left">TARGET CODE</TD>
	<TD align="left">QUANTITY</TD>
	<TD align="left">AMOUNT</TD>
	<TD align="left">REFERENCE</TD>
	<TD align="left">REMARKS</TD>
	<TD align="left">LINE NO (HISTORY)</TD>
	</tr>

	<%
	Dim ValueIsAllowInput : ValueIsAllowInput = 0
	if Not rsVH.eof then
		nCountofField = rsVH.fields.count - 1
		no = 1
		While Not rsVH.EOF
			if no  mod 2 = 1 then
				cColor = "#EEEEEE"
			else
				cColor = "#E5E5E5"
			end if
	%>
			<tr bgcolor="<%=cColor%>">
				<td align="left" valign="top"><%=NO%></td>
				<%If val_IsTrue = "Y" Then%>
				<td valign="middle" width="30">
					<a href="NurseryIssue.Asp?txtDocumentNo=<%=rsVH("documentno")%>&v_NURSERYCODE=<%=rsVH("NURSERYCODE")%>&v_QUANTITY=<%=rsVH("QUANTITY")%>&v_JOBCODE=<%=rsVH("JOBCODE")%>&v_LOCATIONTYPE=<%=rsVH("LOCATIONTYPE")%>&v_LOCATIONCODE=<%=rsVH("LOCATIONCODE")%>&v_reference=<%=rsVH("Reference")%>&v_Remarks=<%=rsVH("Remarks")%>&txtLineNoKey=<%=rsVH("linenokey")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("v_NURSERYCODE")%>" border="0"></a>
				</td>
				<td>
					<a href="NurseryIssue.Asp?txtDocumentNo=<%=rsVH("documentno")%>&v_NURSERYCODE=<%=rsVH("NURSERYCODE")%>&v_QUANTITY=<%=rsVH("QUANTITY")%>&v_JOBCODE=<%=rsVH("JOBCODE")%>&v_LOCATIONTYPE=<%=rsVH("LOCATIONTYPE")%>&v_LOCATIONCODE=<%=rsVH("LOCATIONCODE")%>&txtLineNoKey=<%=rsVH("linenokey")%>&v_ISSUEDATE=<%=v_ISSUEDATE%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("v_NURSERYCODE")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				</td>
				<%Else%>
				<td colspan="2">
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				</td>
				<%End If%>
			    <td align="left" title="<%=rsVH("linenokey").Name%>" valign="top"><%=rsVH("linenokey")%></td>
				<td align="left" title="<%=rsVH("NURSERYCODE").Name%>" valign="top"><%=rsVH("NURSERYCODE")%></td>
				<td align="left" title="<%=rsVH("JOBCODE").Name%>" valign="top"><%=rsVH("JOBCODE")%></td>
				<td align="left" title="<%=rsVH("LOCATIONTYPE").Name%>" valign="top"><%=rsVH("LOCATIONTYPE")%></td>
				<td align="left" title="<%=rsVH("LOCATIONCODE").Name%>" valign="top"><%=rsVH("LOCATIONCODE")%></td>
				<td align="left" title="<%=rsVH("quantity").Name%>" valign="top"><%=rsVH("quantity")%></td>
				<td align="left" title="<%=rsVH("amount").Name%>" valign="top"><%=rsVH("amount")%></td>
				<td align="left" title="<%=rsVH("reference").Name%>" valign="top"><%=rsVH("reference")%></td>
				<td align="left" title="<%=rsVH("remarks").Name%>" valign="top"><%=rsVH("remarks")%></td>
				<td align="left" title="<%=rsVH("LineNoHist").Name%>" valign="top"><%=rsVH("LineNoHist")%></td>
		  </tr>
	<%
			rsVH.MoveNext
			intRecordsPerPage = intRecordsPerPage - 1
			no = no + 1
		Wend
	end if 
	rsVH.Close
	set rsVH = Nothing
   %>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->
<script> 

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
	getTragettype();
});
</script>
</body>
</html>
