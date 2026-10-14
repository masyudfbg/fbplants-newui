<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>Purchase Request</title>
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
<script language="JavaScript" src="/js/date-picker.js"></script> 
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
	doAJAXCall('ajax/getstockinfobyfcba.asp', 'POST', PostStr, showMessageResponse);
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);

    // update the Div to show the result from the server
	document.getElementById("fcname").value = s.fcname;
	document.getElementById("unitofmeasurement").value = s.unitofmeasurement;
	document.getElementById("stockqty").value = s.stockqty;
}

var reloadPrCode = function () {
		var prdate =$( ".datepicker" ).val();
		window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Procurement&txtTable=PurchaseRequestDetails&prdateurl="+prdate );
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
 <!-------------awal!------------->
 <%
 Dim cmdView,PrCode,PrDate,fcba
 Dim Description, Remarks, RequestorId
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim ManagedId, PurchaseGroupId, RowStateId
 Dim LineId, ItemCode, RequestedQty, ApprovedQty, OrderedQty, DescriptionDetail
 Dim UnitOfMeasurement, WantedDeliveryDate
 Dim strItemCodeInfo, rsstrItemCodeInfo
 Dim rsPurchase, rsPurchaseTemp
 Dim strQueryStatus, rsQueryStatus
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 dim txtActivation

Dim txtPRCode : txtPRCode = 0
Dim txtValidStock : txtValidStock = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
 
 Function CheckPRLineRelatedPOInSys(in_fcba,in_PrNo, in_PrLineNo)
	Dim Result : Result = ""
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select BridgingSAP_PRPOGRN_API.CheckPRLineRelatedPOInSys ('"& in_fcba &"','"& in_PrNo &"','"& in_PrLineNo &"') result" & _
					" from dual"
					
	response.write strAvailable

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = ""
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckPRLineRelatedPOInSys = Result
End Function

Function CheckPRLineRelatedPOExtSys(in_fcba,in_PrNo, in_PrLineNo)
	Dim Result : Result = ""
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select BridgingSAP_PRPOGRN_API.CheckPRLineRelatedPOExtSys ('"& in_fcba &"','"& in_PrNo &"','"& in_PrLineNo &"') result" & _
					" from dual"
					
	response.write strAvailable

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		Result = rsAvailable("result")
	Else
		Result = ""
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckPRLineRelatedPOExtSys = Result
End Function


 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 PRCODE = Request("PRCODE")
 
 if Request("fcba") <> "" then
	fcba = Request("fcba")
	else	
	fcba = session("fcba") 
 end if
 
 If Len(Trim(PRCODE)) = 0 Then PRCODE = "(NEW)"
 
 '---- Data Header
 if request("prdateurl") <>"" then
	PRDATE   = request("prdateurl")
 elseif Request("PRDATE")<>"" then
	PRDATE   = Request("PRDATE")
 else
   PRDATE   =textDate(date)
 end if
'  PRDATE = FormatDateTime(PRDATE)
 PRDATE = PRDATE
 'response.write "1." & PRDATE
 'response.write "2." & FormatDateTime(PRDATE)
 
 Description = Request("Description")
 Remarks = Request("Remarks")
 RequestorId = Request("RequestorId")
 ManagedId = Request("ManagedId")
 PurchaseGroupId = Request("PurchaseGroupId")
 RowStateId = Request("RowStateId")
 document_classification = Request("document_classification")
 
 
 
 '---- Data Details
 
 LINEID=Request("LineId")
 ITEMCODE=Request("ITEMCODE")
 REQUESTEDQTY=Request("REQUESTEDQTY")
 APPROVEDQTY=Request("APPROVEDQTY")
 ORDEREDQTY=Request("ORDEREDQTY")
 DESCRIPTIONDETAIL = Request("DESCRIPTIONDETAIL")
 WANTEDDELIVERYDATE = Request ("WANTEDDELIVERYDATE")
 UnitOfMeasurement = Request("UnitOfMeasurement")
 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")= "Cancel" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "ReleaseOnProgress" or Request("txtJenis")= "RefreshPR" or Request("txtJenis")= "CancelOnProgressDet" or Request("txtJenis")= "CloseOnProgressDet" or Request("txtJenis")= "StartAuthorize" Then
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		q = " DELETE FROM purchaserequestdetails WHERE prcode = '" & PRCODE & "' AND fcba = '" & fcba & "'"
		connStr.execute(q)
		
		q = " DELETE FROM purchaserequest WHERE prcode = '" & PRCODE & "' AND fcba = '" & fcba & "'"
		connStr.execute(q)
		
		PRCODE = "(NEW)"
		
		if err.number = 0 then
			connStr.CommitTrans
		end if

		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans
		   q="Delete from PurchaseRequestDetails where  PRCODE='" & PRCODE & "' and LineId='" & request("LineId") & "' and fcBA='" & fcba & "'"
		   connStr.execute(q)
		   if err.number  = 0 then
			  connStr.CommitTrans
		   end if
       cmdView	= "True"	   
	elseif Request("txtJenis")= "Confirmed" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEREQUEST_API.NextState('" & fcba & "','" & PRCODE & "', 'Confirm' ); END; "
	   connStr.execute(strQueryStatus)

       q="Update PurchaseRequestDetails  set rowstateId='Confirmed' where  PRCODE='" & PRCODE & "' and rowstateId='Planned' and fcBA='" & fcba & "'"
	   connStr.execute(q)
	   if err.number  = 0 then
			  connStr.CommitTrans
	   end if       
	   
       cmdView	= "True"
	elseif Request("txtJenis")= "Cancel" then
       q="Update purchaserequest  set ROWSTATEID='Cancel' where  PRCODE='" & PRCODE & "'"
	   connStr.execute(q)
	   if err.number  = 0 then
			  connStr.CommitTrans
	   end if
	   
       cmdView	= "True"
	elseif Request("txtJenis")= "RePlanned" then
	   connstr.BeginTrans   
	   q="Update PurchaseRequest  set rowstateId='Planned' where  PRCODE='" & PRCODE & "' and fcBA='" & fcba & "'"
	   connStr.execute(q)	
      
	   q="Update PurchaseRequestDetails  set rowstateId='Planned' where  PRCODE='" & PRCODE & "' and fcBA='" & fcba & "'"
	   connStr.execute(q)
	   if err.number  = 0 then
			  connStr.CommitTrans
	   end if
       	  

       cmdView	= "True"
	elseif Request("txtJenis")= "StartAuthorize" then
	   connstr.BeginTrans   
		q = " SELECT al.fccode, al.startlevel, al.description " & _
			"  FROM authorizelevel al" & _
			"	   INNER JOIN authorizelevel_users alu" & _
			"		  ON (al.fccode = alu.authorizelevelcode)" & _
			" WHERE al.startlevel = 'TRUE' AND alu.fcusercode = '"& session("CodeName") &"'" 
			
		set rsPurchase = server.CreateObject("Adodb.recordset")
		rsPurchase.Open q, connstr
		IF not rsPurchase.eof then
			authorizelevelcode = rsPurchase("fccode")
		end if
			
	   q = 	" INSERT INTO pr_authorization_comments ( prcode" & _
			"                                         , currentauthorizelevel" & _
			"                                         , comments" & _
			"                                         , fcentry" & _
			"                                         , fcedit" & _
			"                                         , fcip" & _
			"                                         , fcba" & _
			"                                         , lastupdate" & _
			"                                         , lasttime)"& _
			" VALUES ( '" & prcode & "'" & _
			"          , '" & authorizelevelcode & "'" & _
			"          , '" & comments & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"          , '" & fcba & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"
		'response.write q
		connStr.execute(q)
		
		q="Update PurchaseRequestDetails  set rowstateId='AuthorizeOnProgress' where  PRCODE='" & PRCODE & "' and rowstateId='Confirmed' and fcBA='" & fcba & "'"
		connStr.execute(q)
		
	   strQueryStatus = "BEGIN PURCHASEREQUEST_API.NextState('" & fcba & "','" & PRCODE & "', 'StartAuthorize' ); END; "
	   'response.write strQueryStatus
	   connStr.execute(strQueryStatus)	
      
	   if err.number  = 0 then
			  connStr.CommitTrans
	   end if
       	  

       cmdView	= "True"
	
	elseif Request("txtJenis") =  "ReleaseOnProgress" then
	    connstr.BeginTrans
	    
        'Response.Write "BEGIN BridgingSAP_PRPOGRN_API.ReleasePr('" & PRCODE & "','" & fcba & "'); END; "
	   
	    strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.ReleasePR('" & PRCODE & "','" & fcba & "'); END; "
	    connStr.execute(strQueryStatus)	 
        
        cmdView	= "True"    

    elseif Request("txtJenis") =  "RefreshPR" then
	     
		strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.UpdateStatusPRFcba('" & PRCODE & "','" & fcba & "'); END; "
	    connStr.execute(strQueryStatus)
		
        cmdView	= "True"   	
    elseif Request("txtJenis") =  "CancelOnProgressDet" then
	    
		Dim  Msgbox_ 
		Msgbox_ =  CheckPRLineRelatedPOIntSys(fcba,PRCODE, lineid)

	    If Msgbox_ = "TRUE" Then	
            strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.CancelPRLine('" & PRCODE & "','" & lineid &"', '" & fcba & "'); END; "
	        connStr.execute(strQueryStatus)
        else   		
 		    err.Number = 1001
		    err.description = "<br/>" & Msgbox_ 
		End If
		
        cmdView	= "True"
		
    elseif Request("txtJenis") =  "CloseOnProgressDet" then
	    
		Msgbox_ =  CheckPRLineRelatedPOExtSys(fcba,PRCODE, lineid)

	    If Msgbox_ = "TRUE" Then	
            strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.ClosePRLine('" & PRCODE & "','" & lineid &"', '" & fcba & "'); END; "
	        connStr.execute(strQueryStatus)
        else   		
 		    err.Number = 1001
		    err.description = "<br/>" & Msgbox_ 
		End If
		 
        cmdView	= "True"   	 		
	elseif Request("txtJenis")="New" then
	  Dim rsPO
	  Set rsPO = Server.CreateObject("ADODB.Recordset")
	  connstr.BeginTrans
			'--Add by Ukie to handle Auto Number of PR
			If Ucase(PRCODE) = Ucase("(NEW)") Then
				oSql = " select nvl(controlnumber_api.get_nextnumber(upper('purchaserequest')" & _
					   "                                         , " & year(textDate(PrDate)) & "" & _
					   "                                         , " & month(textDate(PrDate)) & "" & _
					   "                                         , '" & fcba & "'" & _
					   " ), 'FALSE') prcode from dual"
				
				Response.Write oSql
					   
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.Open oSql, connstr, adLockOptimistic, adOpenStatic
				If Not oRs.EOF Then
					txtPRCode = oRs("prcode")
				End If
				oRs.Close
				Set oRs = Nothing
				PRCODE = txtPRCode
			End If
			
			If txtPRCode = "FALSE" Then
				err.Number = 1001
				err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
			End If
			
			If ITEMCODE ="" Then
			    err.Number = 1001
				err.description = "<br/> Stock Code Is Empty ( Stock Code Tidak Boleh Kosong )"
			End if
			
			'Response.Write PRCODE & txtPRCode
			'--End Added
			
		  rsPO.Open "Select prCode from PurchaseRequest where prCode='" & PRCODE & "' and fcba = '" & fcba & "'", connstr
		  IF rsPO.eof then	   

			 q="Insert into PURCHASEREQUEST " & _
			 "( PRCODE, DESCRIPTION, PRDATE, REQUESTORID, PURCHASEGROUPID, MANAGEDID, REMARKS, ROWSTATEID, document_classification, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
			 "values ('" & PRCODE & "','" & Description & "', to_date('" & PrDate & "','mm/dd/yyyy') " & ",'"  &  RequestorId & _
			 "','" & PurchaseGroupId  & "','" & ManagedId & "','" & Remarks & "','Planned', '" & document_classification & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			 Request.ServerVariables("REMOTE_ADDR") &"','" & fcba & "', sysdate, To_Char(sysdate,'HH24:MI'))"
			 'response.write q
			 connStr.execute(q)
		  End if
		  rsPO.Close
		  if err.number = 0 then	
		  
			' Added by NAF, 2012-12-01
			 strLineMax = " select max(to_number(nvl(lineid,0))) maxLineIdKey from PurchaseRequestDetails where " & _
							" prcode = '" & PRCODE & "' " & _ 
							" and fcBA='" & fcba & "'"

		  
			Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
			rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
			
			if not rsstrLineMax.eof then
				if len(trim(rsstrLineMax("maxLineIdKey"))) <> 0 then 
					tempMaxLineNo = cdbl(rsstrLineMax("maxLineIdKey")) + 1
				else
					tempMaxLineNo = 1
				end if
			end if
						
			rsstrLineMax.Close
			set rsstrLineMax = Nothing 
			' End of Added by NAF, 2012-12-01
		  
			' Get the Information of Material
			strItemCodeInfo = " select * from Stock where " & _
							" fccode = '" & ItemCode & "' " & _ 
							" and fcBA='" & fcba & "'"

		  
			Set rsstrItemCodeInfo = Server.CreateObject("ADODB.Recordset")
			rsstrItemCodeInfo.Open strItemCodeInfo, connstr, adLockOptimistic, adOpenStatic
			
			if not rsstrItemCodeInfo.eof then
				UnitOfMeasurement = rsstrItemCodeInfo("UnitOfMeasurement")
			end if
						
			rsstrItemCodeInfo.Close
			set rsstrItemCodeInfo = Nothing 
			' End of Added by NAF, 2012-12-01
			
            '--Add by Akbar validation stock code with valid_until  4 Des 2015
			
				oSql = " select stock_api.get_ValidMaterial( '" & ItemCode & "'" & _
					   "                                         , to_date('" & PrDate & "','mm/dd/yyyy') " & _
					   "                                         , '" & fcba & "'" & _
					   " ) ValidStock from dual"
				
				'Response.Write oSql
					   
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.Open oSql, connstr, adLockOptimistic, adOpenStatic
				If Not oRs.EOF Then
					txtValidStock = oRs("ValidStock")
				End If
				oRs.Close
				Set oRs = Nothing
				'PRCODE = txtPRCode
		     
			
			If txtValidStock = "FALSE" Then
				err.Number = 1001
				err.description = "<br/> STOCK CODE " & ItemCode & "  WAS DISCONTINUED"
			End If
			
			oSql = " select stock_api.get_Isactive( '" & ItemCode & "'" & _
					   "                                         , '" & fcba & "'" & _
					   " ) ValidActivation from dual"
				
				'Response.Write oSql
					   
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.Open oSql, connstr, adLockOptimistic, adOpenStatic
				If Not oRs.EOF Then
					txtActivation = oRs("ValidActivation")
				End If
				oRs.Close
				Set oRs = Nothing
				'PRCODE = txtPRCode
		     
			
			If txtActivation = "N" Then
				err.Number = 1001
				err.description = "<br/> STOCK CODE " & ItemCode & "  WAS NON ACTIVE"
			End If
			
			
			'--Response.Write ItemCode & " Status "  & txtValidStock
			'--End Added
			
		  

		  
			q= "Insert into  PurchaseRequestDetails " & _
				"(PRCODE, LINEID, lineid_reference, ITEMCODE, REQUESTEDQTY, APPROVEDQTY, ORDEREDQTY, UNITOFMEASUREMENT, WANTEDDELIVERYDATE, DESCRIPTION, RowStateId, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
				"values ('" & _
				  PRCODE & "','" & _
				  tempMaxLineNo & "','" & _
				  tempMaxLineNo & "','" & _
				  ITEMCODE & "'" & _
				  ",replace('" & REQUESTEDQTY & "',',',''),0,0,'" & _
				  UNITOFMEASUREMENT & "'," & _
				  "to_date('" & WantedDeliveryDate & "','mm/dd/yyyy'),'" & _
				  DESCRIPTIONDETAIL  & "'," & _
				  "'Planned','" & _
				  Session("CodeName")  & "','" & _ 
				  Session("CodeName") & "','" & _
				  Request.ServerVariables("REMOTE_ADDR") &"','" & _
				  fcba & "', sysdate, To_Char(sysdate,'HH24:MI'))"
				  
				  
			 
			 'Response.write q
			 connStr.execute(q)
			 
			 
			 
			 if err.number  = 0 then
				APPROVEDQTY = 0
				ORDEREDQTY = 0
				connStr.CommitTrans
			 end if
		  end if
      cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	    
		'--Add by Akbar validation stock code with valid_until  4 Des 2015
			
				oSql = " select stock_api.get_ValidMaterial( '" & ItemCode & "'" & _
					   "                                         , to_date('" & PrDate & "','mm/dd/yyyy') " & _
					   "                                         , '" & fcba & "'" & _
					   " ) ValidStock from dual"
				
				'Response.Write oSql
					   
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.Open oSql, connstr, adLockOptimistic, adOpenStatic
				If Not oRs.EOF Then
					txtValidStock = oRs("ValidStock")
				End If
				oRs.Close
				Set oRs = Nothing
				'PRCODE = txtPRCode
		
			
			If txtValidStock = "FALSE" Then
				err.Number = 1001
				err.description = "<br/> STOCK CODE " & ItemCode & "  WAS DISCONTINUED"
			End If
			
			'--Response.Write ItemCode & " Status "  & txtValidStock
			'--End Added

			connstr.BeginTrans
			   q="Update PurchaseRequestDetails set ApprovedQty = REPLACE('" & request("RequestedQty") & "', ',', '') " & _
				 ", WantedDeliveryDate = to_date('" & WantedDeliveryDate & "','mm/dd/yyyy') " & _
				 ", Description = '" & DescriptionDetail & "' " & _
				 ", fcedit = '" & Session("CodeName") & "' " & _
				 ", LASTUPDATE = sysdate " & _
				 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
				 " where  PRCode='" & PrCode & "' and LineId='" & LineId & "' and fcBA='" & fcba & "'"
				' Response.Write q
				connStr.execute(q)
				if err.number  = 0 then
				   connStr.CommitTrans
				   
					set rsPurchaseTemp = server.CreateObject("Adodb.recordset")
					q=" select * from PURCHASEREQUESTDETAILS  " & _
					  " where  PRCode='" & PrCode & "' and LineId='" & LineId & "' and fcBA='" & fcba & "'"
					
					rsPurchaseTemp.Open q, connstr
					
					if Not rsPurchaseTemp.Eof then
						APPROVEDQTY=rsPurchaseTemp("APPROVEDQTY")
						ORDEREDQTY = rsPurchaseTemp("ORDEREDQTY")

					End if
					rsPurchaseTemp.Close
					set rsPurchaseTemp = Nothing 
				   
				end if
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if

 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select P.PRCODE, P.DESCRIPTION, TO_CHAR(P.PRDATE,'MM/DD/YYYY')PRDATE, P.REQUESTORID, P.PURCHASEGROUPID, " & _
      " P.MANAGEDID, P.REMARKS, P.ROWSTATEID, P.FCENTRY, P.FCEDIT, " & _
      " P.FCIP, P.FCBA, P.LASTUPDATE, P.LASTTIME, P.ISINBRIDGE, " & _
      " P.DOCUMENT_CLASSIFICATION from PURCHASEREQUEST P " & _
  " where P.PrCode='" & PrCode & "' and " & _
  " P.fcBA='" & fcba & "'"
  
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    PrDate=rsPurchase("PrDate")
    Description=rsPurchase("Description")
    document_classification=rsPurchase("document_classification")
	RequestorId = rsPurchase("RequestorId")
	PurchaseGroupId = rsPurchase("PurchaseGroupId")
	ManagedId = rsPurchase("ManagedId")	
	Remarks=rsPurchase("Remarks")
	RowStateId = rsPurchase("RowStateId")
 End if
if err.number <> 0 then
	connStr.RollbackTrans
    Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" WHERE at.document_type = 'PR'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'TRUE'" & _
	"       AND at.document_classification = '" & document_classification & "'" & _
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

set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" WHERE at.document_type = 'PR'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'FALSE'" & _
	"       AND at.document_classification = '" & document_classification & "'" & _
	"       AND aus.fcusercode = '" & Session("CodeName") & "'" & _
	"       AND aus.fcba = '" & fcba & "'"
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	can_approved =  cdbl(to_nol(rsFrez("total")))
else
	can_approved= 0
End if
rsFrez.Close : set rsFrez = Nothing 

set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" WHERE at.document_type = 'PR'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"       AND at.document_classification = '" & document_classification & "'" & _
	"       AND aus.fcusercode = '" & Session("CodeName") & "'" & _
	"       AND aus.fcba = '" & fcba & "'"
' response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	is_final =  cdbl(to_nol(rsFrez("total")))
else
	is_final = 0
End if
rsFrez.Close : set rsFrez = Nothing 

has_approveddata = 0
 
q = " SELECT COUNT(1) total FROM purchaserequestdetails WHERE prcode = '" & PrCode & "' AND fcba = '" & fcba & "' AND approvedqty > 0"
set rsPurchase = server.CreateObject("Adodb.recordset")
rsPurchase.Open q, connstr
if Not rsPurchase.Eof then
	has_approveddata = cdbl(to_nol(rsPurchase("total")))
End if

 '-------------ISI Payment Form
%>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Purchase Request</div><% end if %>
 <form name="frmEntry" method="post" action="PurchaseRequestDetails.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Purchase Request No</td>
		<%
			'response.write PrCode
			if PrCode = "(NEW)" then
				q = "select MAX(cn.PREFIX) || LPAD(NVL(MAX(to_number(regexp_substr(pr.PRCODE, '[^/]+', 1, 4))),0000)+0001,4,0) DOCNUMBER " & _
					" from " & _
					" controlnumber cn " & _
					" left join purchaserequest pr on cn.fcba = pr.fcba and cn.accountingyear = to_number(to_char(pr.PRDATE,'yyyy')) " & _
					" where cn.fcba ='" & fcba & "' " & _
					" and cn.LU_NAME ='PURCHASEREQUEST' "&_
					" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& PRDATE &"','mm/dd/yyyy'),'yyyy')) " & _
					" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& PRDATE &"','mm/dd/yyyy'),'mm')) "
				'Response.write q
				set rsPRNo = server.CreateObject("Adodb.recordset")
				rsPRNo.Open q, connstr
				if Not rsPRNo.Eof then
					PrCode = rsPRNo("DOCNUMBER")
				End if
			End if
		%>
		<td align="left"><input type="text" maxlength="50" id="PrCode" name="PrCode" size="30" value="<%=PrCode%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" onchange="javascript:reloadPrCode();"   type="text" maxlength="10" name="PRDATE" size="10" value="<%=PRDATE%>">
		<!--<a href="javascript:show_calendar('frmEntry.PRDATE');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
		</td>
		</tr>
		
		<tr>
		<td align="left">Description</td>
		<td align="left"><textarea  name="Description"  cols="50" rows="3"><%=Description%></textarea></td>		
		</tr>
		
		<tr>
		<td align="left">Classification</td>
		<td align="left">
			<select name="document_classification">
			<%
			qLov = " SELECT fccode, fcname FROM document_classification where document_type= 'PR' and isactive='TRUE' ORDER BY fccode"
   
			Set rsLov = Server.CreateObject("ADODB.Recordset")
			rsLov.Open qLov, connstr
			if not rsLov.eof then
				while not rsLov.eof
			%>
				<option value="<%=rsLov("fccode")%>" <% if cstr(""&rsLov("fccode")) = cstr(""&document_classification) then%> selected <%end if%>><%=rsLov("fcname")%></option>
			<%
					rsLov.movenext
				wend
			end if
			%>
			</select>
			<font color="#FF0000">*</font>
		</td>		
		</tr>
		
	
		<tr>
		<td align="left">Requestor</td>
		<td align="left">      
		<select name="RequestorId"  >
	    <%cmdCallCombo "REQUESTOR",RequestorId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>


		
		<tr>
		<td align="left">Managed By</td>
		<td align="left">      
		<select name="ManagedId" >
	    <%cmdCallCombo "PURCHASEMANAGER",ManagedId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>	
		</tr>
		
		<tr>
		<td align="left">Purchase Group By</td>
		<td align="left">      
		<select name="PurchaseGroupId" >
	    <%cmdCallCombo "PURCHASEGROUP",PurchaseGroupId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>	
		</tr>		
		
		
		<tr>
		<td align="left">Remarks</td>
		<td align="left"><textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea></td>		
		</tr>	

		<tr>
		<td align="left">Status</td>
		<td align="left"><input disabled type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>"><font color="#FF0000">*</font>
		</td>
		</tr>

		
		
		</td>
		</tr>

	

		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		
		
		
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('PrintPR.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&PrCode=<%=PrCode%>&fcba=<%=fcba%>&PrDate=<%=PrDate%>&Decription=<%=Description%>&RequestorId=<%=RequestorId%>&PurchaseGroupId=<%=PurchaseGroupId%>&ManagedId=<%=ManagedId%>&Remarks=<%=Remarks%>&RowStateId=<%=RowStateId%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1" style="font-size: 14px">Print</font></a>&nbsp;&nbsp;
		
		<%if RowStateId = "Approved" then%> 
			<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> 
		<%else%>
		    <a href="PurchaseRequestDetails.Asp?PrCode=<%=PrCode%>&PrDate=<%=PrDate%>&Decription=<%=PrDate%>&RequestorId=<%=RequestorId%>&PurchaseGroupId=<%=PurchaseGroupId%>&ManagedId=<%=ManagedId%>&Remarks=<%=Remarks%>&RowStateId=<%=RowStateId%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=PrCode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
		<%end if%> 
		
		<%if has_approveddata = 0 then%>

		<% if RowStateId = "Planned" then %>
		    <a href="PurchaseRequestDetails.Asp?PrCode=<%=PrCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to confirm this data :  <%=PrCode%> ?');" style="font-size: 14px"> Confirmed </a>
			<a href="PurchaseRequestDetails.Asp?PrCode=<%=PrCode%>&txtJenis=Cancel" onClick="return cancel('Are you sure to cancel this data :  <%=PrCode%> ?');" 
			style="background-color: #c41212; color: #fff; padding: 5px; border-radius: 3px; margin-left: 15px; font-size: 14px"> 
				Cancel 
			</a>
		<% elseif RowStateId = "Confirmed" then %>
		    <a href="PurchaseRequestDetails.Asp?PrCode=<%=PrCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to replan this data :  <%=PrCode%> ?');"> Re-Planned </a>
		    &nbsp;&nbsp;
			<%if can_startauthorize = 0 then%>
				<font color="#FF0000">You Have No Permission to Start Authorize</font>
			<%else%>
				<a href="PurchaseRequestDetails.Asp?PrCode=<%=PrCode%>&txtJenis=StartAuthorize" onClick="return confirm('Are you sure to start authorize this data :  <%=PrCode%> ?');">Start Authorize</a>
			<% end if%>
		<% elseif RowStateId = "AuthorizeOnProgress" AND ( can_approved > 0 or is_final > 0) then %>
			<a href="PurchaseRequestDetails.Asp?PrCode=<%=PrCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to replan this data :  <%=PrCode%> ?');"> Re-Planned </a>
		<% end if%>
		
		<% end if%>
		
		</td>     
		</tr>
    </table>       
	
 </form>
<%if cmdView = "True" AND LEN(TRIM(PrCode)) > 4 Then
   Dim rsPurchaseRequest,intRecordsPerPage,cColor
   set rsPurchaseRequest = server.CreateObject("Adodb.recordset")
   q="select PRCODE, LINEID, ITEMCODE,stock_api.get_stockName(itemcode,fcba) itemname, REQUESTEDQTY, APPROVEDQTY, ORDEREDQTY, UNITOFMEASUREMENT, WANTEDDELIVERYDATE, DESCRIPTION, ROWSTATEID, FCENTRY, FCEDIT, FCIP, FCBA from PURCHASEREQUESTDETAILS  " & _
     " where PrCode='" & PrCode & "' and " & _
     " fcBA='" & fcba & "'"
   'response.write q
   Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
   rsPurchaseRequest.CursorType = 3
   rsPurchaseRequest.CursorLocation = 3
   rsPurchaseRequest.Open q, connstr%>
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="PurchaseRequestDetails.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr><td  align="left">Stock Code</td>
		<td align="left">
		<%IF txtJenis="Edit" Then%>
		     <input type="text" size="15" id="fccode" name="ITEMCODE" value="<%=ITEMCODE%>" readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">*</font>
		<%else%>
		     <input type="text" size="15" id="fccode" name="ITEMCODE" value="<%=ITEMCODE%>" readonly="True" style="background:#eee">&nbsp;<font color="#FF0000">*</font>
		<%End if%>
		<!--<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Stock&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font></a></td>-->
		<a href="javascript:launchwin('popup_allstock_new.Asp?fcba=<%=fcba%>&prdate=<%=prdate%>','newwindow','width='+(800)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font></a></td>
		
		</tr>
		
		<tr>
		     <td align="left">Stock Name </td>
			 <td align="left">
			 <input type="text" size="50" id="fcname" name="stockname" value="<%=stockname%>" readonly style="background-color:#eee;">
			 &nbsp;<font color="#FF0000">*</font>
			 </td>
		</tr>
		
	    <tr>
		     <td align="left">Unit of Measure </td>
			 <td align="left">
			 <input type="text" size="15" id="unitofmeasurement" name="UNITOFMEASUREMENT" value="<%=UNITOFMEASUREMENT%>" readonly style="background-color:#eee;">
			 &nbsp;<font color="#FF0000">*</font>
			 </td>
		</tr>
		<tr>
		     <td align="left">Stock On Hand Qty</td>
			 <td align="left">
			 <input type="text" size="15" id="stockqty" name="stockqty" value="<%=stockqty%>" readonly style="background-color:#eee;" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">&nbsp;</font>
		</tr>
		<tr>
		     <td align="left">Requested Qty</td>
			 <td align="left">
			 <input type="text" size="15" name="REQUESTEDQTY" value="<%=REQUESTEDQTY%>"data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" id="aninput" onkeypress="return validateFloatKeyPress(this,event);">&nbsp;<font color="#FF0000">*</font>
		</tr>

		
		<tr>
		
		<td align="left">Wanted Delivery Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="WANTEDDELIVERYDATE" size="10" value="<%=WANTEDDELIVERYDATE%>">
		&nbsp;<font color="#FF0000">*</font>
		</td>
		</tr>
		

	    <tr>
		     <td align="left">Approved Qty</td>
			 <td align="left">
			 <input type="text" size="15" name="APPROVEDQTY" value="<%=APPROVEDQTY%>" readonly style="background-color:#eee;" >&nbsp;<font color="#FF0000">*</font>
			 </td>
		</tr>
		
	    <tr>
		     <td align="left">Ordered Qty</td>
			 <td align="left">
			 <input type="text" size="15" name="ORDEREDQTY" value="<%=ORDEREDQTY%>"  readonly style="background-color:#eee;" >&nbsp;<font color="#FF0000">*</font>
			 </td>
		</tr>
		
		<tr>
			<td align="left">Description</td>
			<td align="left"><textarea  name="DESCRIPTIONDETAIL"  cols="50" rows="3" value=<%=DESCRIPTIONDETAIL%> ><%=DESCRIPTIONDETAIL%></textarea></td>		
		</tr>			

		
		
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="PRCODE" value="<%=PRCODE%>">
		  <input type="hidden" name="PRDATE" value="<%=PRDATE%>">
		  <input type="hidden" name="Description" value="<%=Description%>">
		  <input type="hidden" name="document_classification" value="<%=document_classification%>">
		  <input type="hidden" name="RequestorId" value="<%=RequestorId%>">
		  <input type="hidden" name="ManagedId" value="<%=ManagedId%>">
		  <input type="hidden" name="PurchaseGroupId" value="<%=PurchaseGroupId%>">		  
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="LineId" value="<%=LINEID%>">
		  
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <% if rowstateId = "Planned" or rowstateId = "" then %> 
		      <%IF Request("txtJenis")="Edit" then%>		  
				   <input type="hidden" name="txtJenis" value="Edit">
				   <input Type="Submit"  name="Add" value="Update">
			  <%Else%>
				   <input type="hidden" name="txtJenis" value="New">
				   <input Type="Submit"  name="Add" value="Add">
			   <% end if %>
		  <%End if%>
		  
		  <% if rowstateId = "Approved" and is_final > 0 then %> 
		      <%IF Request("txtJenis")="Edit" then%>		  
				   <input type="hidden" name="txtJenis" value="Edit">
				   <input Type="Submit"  name="Add" value="Update">
			  <% end if %>
		  <%End if%>
		  
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="middle" colspan="2"> ACTION</TD>
	<TD align="left">PR CODE</TD>	
	<TD align="left">LINE ID</TD>	
	<TD align="left">ITEM CODE</TD>	
	<TD align="left">ITEM NAME</TD>	
	<TD align="left">REQUESTED QTY</TD>	
	<TD align="left">APPROVED QTY</TD>	
	<TD align="left">ORDERED QTY</TD>	
	<TD align="left">UNIT OF MEASURE</TD>	
	<TD align="left">WANTED DELIVERY DATE</TD>	
	<!--<TD align="left">INSERTED</TD>	
	<TD align="left">UPDATED</TD>	
	<TD align="left">IP ADDRESS</TD>	
	<TD align="left">BA</TD>-->	
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">LINE STATUS</TD>
	</tr>
	</thead>
	<tbody>
   <%
   if Not rsPurchaseRequest.eof then
      nCountofField=rsPurchaseRequest.fields.count-1
	  no=1
	  While Not rsPurchaseRequest.EOF

		if no  mod 2 = 1 then
			cColor = "#EEEEEE"
		else
			cColor = "#E5E5E5"
		end if%>
		<tr bgcolor="<%=cColor%>">
		<td align="left" valign="top"><%=NO%></td>
			
            
            <% If rsPurchaseRequest("ROWSTATEID")="Planned" Then %>			
			    <td valign="middle" width="10">
				   <a href="PurchaseRequestDetails.Asp?PrCode=<%=rsPurchaseRequest("PrCode")%>&LineId=<%=rsPurchaseRequest("LineId")%>&ItemCode=<%=rsPurchaseRequest("ItemCode")%>&fcba=<%=fcba%>&WantedDeliveryDate=<%=rsPurchaseRequest("WantedDeliveryDate")%>&RequestedQty=<%=rsPurchaseRequest("RequestedQty")%>&RequestorId=<%=RequestorId%>&UnitOfMeasurement=<%=rsPurchaseRequest("UnitOfMeasurement")%>&ApprovedQty=<%=rsPurchaseRequest("ApprovedQty")%>&OrderedQty=<%=rsPurchaseRequest("OrderedQty")%>&DescriptionDetail=<%=rsPurchaseRequest("Description")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseRequest(1)%>" border="0">
				   </a>
			   </td> 
			   <td valign="middle" width="10">
					<a href="PurchaseRequestDetails.Asp?PrCode=<%=rsPurchaseRequest("PrCode")%>&LineId=<%=rsPurchaseRequest("LineId")%>&fcba=<%=fcba%>&RequestedQty=<%=rsPurchaseRequest("RequestedQty")%>&PRDATE=<%=PRDATE%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseRequest(1)%>" border="0">
					</a>
			   </td>
            <% ElseIf rsPurchaseRequest("ROWSTATEID")="Released"  And cdbl(rsPurchaseRequest("ORDEREDQTY")) = 0 then %>	
			    <td valign="middle" >
				   <a href="PurchaseRequestDetails.Asp?PrCode=<%=rsPurchaseRequest("PrCode")%>&LineId=<%=rsPurchaseRequest("LineId")%>&fcba=<%=fcba%>&RequestedQty=<%=rsPurchaseRequest("RequestedQty")%>&ORDEREDQTY=<%=rsPurchaseRequest("ORDEREDQTY")%>&PRDATE=<%=PRDATE%>&txtJenis=CancelOnProgressDet" onClick="return confirm('Are you sure to Cancel This data ?');">Cancel</a>
			    </td>
				<td valign="middle" >
				   <a href="PurchaseRequestDetails.Asp?PrCode=<%=rsPurchaseRequest("PrCode")%>&LineId=<%=rsPurchaseRequest("LineId")%>&fcba=<%=fcba%>&RequestedQty=<%=rsPurchaseRequest("RequestedQty")%>&ORDEREDQTY=<%=rsPurchaseRequest("ORDEREDQTY")%>&PRDATE=<%=PRDATE%>&txtJenis=CloseOnProgressDet" onClick="return confirm('Are you sure to Close This data ?');">Close</a>
			    </td>
			<% ElseIf rsPurchaseRequest("ROWSTATEID")="Released"  And cdbl(rsPurchaseRequest("ORDEREDQTY")) > 0 then %>	
			    
				<td valign="middle" colspan="2">
				   <a href="PurchaseRequestDetails.Asp?PrCode=<%=rsPurchaseRequest("PrCode")%>&LineId=<%=rsPurchaseRequest("LineId")%>&fcba=<%=fcba%>&RequestedQty=<%=rsPurchaseRequest("RequestedQty")%>&ORDEREDQTY=<%=rsPurchaseRequest("ORDEREDQTY")%>&PRDATE=<%=PRDATE%>&txtJenis=CloseOnProgressDet" onClick="return confirm('Are you sure to Close This data ?');">Close</a>
			    </td>	
			<% ElseIf rsPurchaseRequest("ROWSTATEID")="Approved" And cdbl(rsPurchaseRequest("APPROVEDQTY")) > 0  and is_final > 0 Then %>			
			    <td valign="middle" colspan="2" width="10">
				   <a href="PurchaseRequestDetails.Asp?PrCode=<%=rsPurchaseRequest("PrCode")%>&LineId=<%=rsPurchaseRequest("LineId")%>&ItemCode=<%=rsPurchaseRequest("ItemCode")%>&fcba=<%=fcba%>&WantedDeliveryDate=<%=rsPurchaseRequest("WantedDeliveryDate")%>&RequestedQty=<%=rsPurchaseRequest("RequestedQty")%>&RequestorId=<%=RequestorId%>&UnitOfMeasurement=<%=rsPurchaseRequest("UnitOfMeasurement")%>&ApprovedQty=<%=rsPurchaseRequest("ApprovedQty")%>&OrderedQty=<%=rsPurchaseRequest("OrderedQty")%>&DescriptionDetail=<%=rsPurchaseRequest("Description")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseRequest(1)%>" border="0">
				   </a>
			   </td> 
			<% Else %>	
               	<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>		
            <% End If %>	

			
		    <%for i = 0 to nCountofField -4%>
			    <td align="left" title="<%=rsPurchaseRequest(i).Name%>" valign="top"><%=rsPurchaseRequest(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsPurchaseRequest.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("<tr><td colspan=""14"" class=""fbp-doc-empty"">Data is empty</td></tr>")
   end if
   rsPurchaseRequest.Close
   set rsPurchaseRequest = Nothing%>
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
</div></div>
<!--#include file="bottom.jsp"-->
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
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
	getMessage();
});
</script>
</body>
</html>
