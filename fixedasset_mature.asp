<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
    <title><%="FIXED ASSET MATURITY | FB-PLANTS"%></title>
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
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />
<script language="JavaScript">
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<script language="javascript" type="text/javascript">
function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
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
var getJobCode = function () {
	// build up the post string when passing variables to the server side page
	
	var selectedjobcode = document.getElementById("selectedjobcode").value;
	var PostStr = 'fccode='+selectedjobcode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getJobCode.asp', 'POST', PostStr, showJobcode);
}

// The function for handling the response from the server
var showJobcode = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("jobcodelist").innerHTML = "";
	document.getElementById("jobcodelist").innerHTML = response;

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


var getJobcodeIntraco = function () {
	// build up the post string when passing variables to the server side page
	//var jobcode = document.getElementById("jobcode").value;
	var jobcode = document.getElementById("selectedjobcode").value;
	// var fddate = document.getElementById("fddate").value;
	// var empcode = document.getElementById("empcd").value;
	// var fieldcode = document.getElementById("targetcode").value;
	var PostStr = 'jobcode='+jobcode;
	//alert("masuk intraco");
	// use the generic function to make the request
	doAJAXCall('ajax/getChargeIntraco.asp', 'POST', PostStr, showJobcodeIntraco);
}

// The function for handling the response from the server
var showJobcodeIntraco = function (oXML) { 
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.convertion);
    // update the Div to show the result from the server
	// document.getElementById("salescontractno").value = s.salescontractno;
	// alert(s.jobcodeintraco);
	// document.getElementById("jobcodeintraco").value=s.jobcodeintraco;

	if (s.jobcodeintraco > 0) { 
		// document.getElementById("jobcodeintraco").style.display="block";
        document.getElementById("chargeunit").style.display="table-row";
		document.getElementById("chargejob").style.display="table-row";
		document.getElementById("chargetype").style.display="table-row";
		document.getElementById("chargecode").style.display="table-row";
	} else {
		// document.getElementById("jobcodeintraco").style.display="none";
		document.getElementById("chargeunit").style.display="none";
		document.getElementById("chargejob").style.display="none";
		document.getElementById("chargetype").style.display="none";
		document.getElementById("chargecode").style.display="none";
	}
}

$.urlParam = function(name){
    var results = new RegExp('[\?&]' + name + '=([^&#]*)').exec(window.location.href);
    if (results==null) {
       return null;
    }
    return decodeURI(results[1]) || 0;
}

var reloadCode = function (PostStr) {
	
	var transdate =$( ".datepicker" ).val();
	//var bankcode =$( "#bankcode" ).val();
	var fcba =$( "#fcba" ).val();
	var accperiod =$( "#ACCOUNTINGYEARPERIOD" ).val();
	var txtTable = $.urlParam('txtTable');
	//var txtBanktype = $.urlParam('banktype');
	//var vtranstype = $.urlParam('transtype');
	//alert(bankcode);
	
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Finance and Accounting&txtTable="+txtTable+"&transdate="+transdate+"&fcba="+fcba+"&ACCOUNTINGYEARPERIOD="+accperiod+"&"+PostStr );
}

// launched from button click 
var getTragettype = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("selectedjobcode").value;
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

var getChargetype = function () {
	// build up the post string when passing variables to the server side page
	var jobcode = document.getElementById("chargejob_").value;
	var PostStr = 'jobcode='+jobcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetPVInfo.asp', 'POST', PostStr, showChargettype);
	
}

// The function for handling the response from the server
var showChargettype = function (oXML) { 
     // alert("1");
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(response);
    // update the Div to show the result from the server
    document.getElementById("chargetype_").value = s.target; 
	// document.getElementById("descpr").value = s.description; 
}


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


</script>
    <script language="JavaScript">
        function showDetails(winurl, p_item) {
            window.open(winurl + '&p_item=' + p_item, width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1);
        }

		
    </script>
	 <script language="JavaScript">
        function PopupTType(_fcba) {
		var txtttype = "FF";
		var txtmonth = document.getElementById("ACCOUNTINGPERIOD").value;
		var txtyear = document.getElementById("ACCOUNTINGYEAR").value;
		
		var txtdate  = document.getElementById("DATECREATED").value;
		
		//var newwin   = window.open("PrintData.Asp?txtModule=Data Master&txtTable=LocationCode&txtSearch="+txtttype+"&txtFcba="+_fcba+"&Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
		var newwin   = window.open("popup_loccode_fa.Asp?txtSearch="+txtttype+"&txtmonth="+txtmonth+"&txtyear="+txtyear+"&Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
               
	   }

		
    </script>
</head>

<style>
	
	.lebarsup {
		position: absolute;
        right: 380;
        width: 200px;

	}
	.ltinggi{
		padding-top:-50px;
	}
</style>


<%if request("Add")="" then%>
<body onLoad="document.frmEntry.fcCode.focus()">
 <%End if%>
        <!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
        <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	  Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">   
    <!-----------------ATAS ISI!----------------->

<%
 '-------------awal-----------------------
 
Function CheckAvailableJobZerorise(initialJob, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from parameter" & _
	                " where upper(parcode) = upper('JOBCODE_ZERORISE')" & _
					"       and fccode = '" & Trim(initialJob) & "'" & _
					"       and fcba = '" & Trim(initialFcba) & "'"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckAvailableJobZerorise = initialAvailable
End Function



Function CheckCostJobCr(statusxAmount ,initialJob, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from Costjobcr" & _
	                " where  jobcode = '" & Trim(initialJob) & "'" & _
					" and fcba = '" & Trim(initialFcba) & "'"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	IF statusxAmount = "MINUS"  THEN
			If Not rsAvailable.EOF Then
				initialAvailable = Cdbl(rsAvailable("Total"))
			Else
				initialAvailable = 0
			End If
	 Else
	    initialAvailable = 0
	 END IF
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckCostJobCr = initialAvailable
End Function

Function CheckCharge(initType, initCode, initJob, initFcba, initDate, initSource)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT F_CekTarget_byDate ('"& initType &"'," & _
					"                           '"& initCode &"'," & _
					"                           '"& initJob &"'," & _
					"                           '"& initFcba &"'," & _
					"                           TO_DATE ('"& initDate &"', 'MM/DD/YYYY')," & _
					"                           'BA') Total" & _
					"  FROM DUAL"

	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("Total")
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckCharge = initialAvailable
End Function

 Dim check

 Dim xCode,fcCode,cmdView,DATECREATED,documentno,QUANTITY,AMOUNT,OTHERITEMDESC,SourceType,assetcode
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
 Dim rsInvoiceRecord, varTotalAmount, invoiceUpdate, recordcheck, rsrecordlineno, recordWIPLINE,invoiceRecord
 Dim rsfixedasset_mature_new_dtl, varamountOriginal, recordinvoicepaid, vartotal, varError, ivp,varmaxlineno,recordlineno
 Dim Jobcode, LocationType, LocationCode, Reference, Remark,InvoiceCode,AGREEMENTCODE,AGREEMENTLINENO ,recordwipcheck
 Dim fisik_qty, fisik_unit, fisik_output
 Dim varSupplierCode, varStringError, varlineno, WIPPAIDRecord, rsWIPPAIDRecord, WIPUpdate ,rsrecordWIPLINE ,W, rsrecordwipcheck, WIPDATE, varContract
 Dim InsertIntoCostActivity, rsqMaxLineNo, qMaxLineNo, MaxLineNo, IsCostbook
 Dim MaxLineNopvd
 Dim rlineno, dataModule, dataSubmodule, dataDetail, qq, recfixedasset_mature_new_dtl
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim Cek_IsAccClosing : Cek_IsAccClosing = ""
 
 Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD
 Dim xACCOUNTINGYEARPERIOD
 DIM xAmount, StatusXAmount
 Dim sqljnsinv,rsjnsinv,sqlupdatepay,txtnoinvoice,jenisinv,amountinvoice,ScriptName
 ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 
 if Request("fcba") <> "" then
	fcba = Request("fcba")
 else	
	fcba = session("fcba") 
 end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if request("transdate") <>"" then
   DATECREATED   = request("transdate")
 elseif Request("DATECREATED")<>"" then
   DATECREATED = Request("DATECREATED")
 else
   DATECREATED = textDate(date)
 end if
 
 if KonversiTanggalKeDbl(DATECREATED) > KonversiTanggalKeDbl(textDate(date)) then
	DATECREATED = textDate(date)
 end if
 'response.write DATECREATED
 
 if request("bankcodeurl") <>"" then
   xCode   = Request("bankcodeurl")
   'response.write xCode
 elseif Request("xCode")<>"" then
   xCode = Request("xCode")
 else
   xCode = Request("xCode")
 end if
 
 
 HeaderRemarks = Request("HeaderRemarks")
 
 '---- Data Details


LocationCode=Request("LocationCode")
Reference=Request("Reference")
Remarks=Request("Remarks")
Amount=Request("Amount")
AGREEMENTCODE=Request("AGREEMENTCODE")
AGREEMENTLINENO=Request("AGREEMENTLINENO")
rlineno = Request("txtlineno")
documentno = Request("documentno")

' Added Variable ChargeJob, ChargeType, ChargeCode
referenceno=Request("referenceno")
remarksdtl=Request("referenceno")
fcCode = Request("fcCode")
assetcode = Request("assetcode")

' End Added


If Request("Amount") <> ""  then
   xAmount = Replace(Amount, ",", "")
   IF Cdbl(xAmount) > 0 then
      StatusXAmount = "PLUS"
   Else
      StatusXAmount = "MINUS"       
   End If
End If

 ACCOUNTINGYEARPERIOD = Request("ACCOUNTINGYEARPERIOD")
 
 if "" & ACCOUNTINGYEARPERIOD = "" then
 ACCOUNTINGYEARPERIOD = Right(DATECREATED,4) & MID(DATECREATED,1,2)
 else
 ACCOUNTINGYEARPERIOD = Request("ACCOUNTINGYEARPERIOD")
 end if
 
 if "" & ACCOUNTINGYEARPERIOD <> "" then
    ACCOUNTINGYEAR = MID(ACCOUNTINGYEARPERIOD,1,4)
	ACCOUNTINGPERIOD = MID(ACCOUNTINGYEARPERIOD,5,2)
 end if

'response.write ACCOUNTINGYEARPERIOD & "|" & ACCOUNTINGYEAR & "|" & ACCOUNTINGPERIOD

 if Request("SourceType")="" then
    SourceType="BA" 'Default if Entry is Blank
 else
    SourceType=Request("SourceType")
 End if
dim tt

Sub DeleteDatafixedasset_mature_new_dtl
	q = " DELETE FROM fixedasset_mature_new_dtl" & _
		" WHERE fcBA = '" & fcba & "'" & _
		"       AND documentno = '" & fcCode & "'" & _
		"       AND lineno = '" & rlineno & "'"
	'connStr.execute(q)
	ExecuteAndLog ScriptName, q
	Call UpdateAmountfixedasset_mature_newHeader
End Sub

Sub UpdateAmountfixedasset_mature_newHeader
' Untuk Update Total Amount Di Header
				q = " SELECT NVL(SUM(amount), 0) totalamount" & _
					" FROM fixedasset_mature_new_dtl" & _
					" WHERE documentno='" & fcCode & "'" & _
					"       AND fcBA = '" & fcba & "'"
				set rsVoucher = server.CreateObject("Adodb.recordset")
				rsVoucher.Open q, connstr
				if Not rsVoucher.Eof then
					TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
				End if

				q =	" Update fixedasset_mature_new" & _
					" set totalamount = '"& TOTALAMOUNT &"'" & _
					" where documentno = '" & fccode & "'" & _
					" and fcBA = '" & fcba & "'"
	 			'connStr.execute(q)
				ExecuteAndLog ScriptName, q
End Sub

function Amount_LocationCode(Blok_)
Dim qryAmt,rsAmt
qryAmt =   " SELECT round(Sum(Nvl(amountdebit,0)+Nvl(amountkredit,0)),2)  amount " & _
		" FROM bkmjournaldtl a " & _
		" INNER JOIN bkmjournalhdr b ON (a.docno=b.docno AND a.fcba=b.fcba)" & _
		" INNER JOIN jobmaster c ON (a.jobcode=c.fccode)" & _
		" WHERE a.locationcode='" & Blok_ & "' AND a.fcba='" & fcba & "'" & _
		" AND postingflg=1" & _
		" AND job_forfieldstatus IN ('TBM','LC')" & _
		" AND docdate < To_Date('" & ACCOUNTINGPERIOD & "/01/" & ACCOUNTINGYEAR & "','MM/DD/YYYY') " & _ 
		" AND b.accountingperiod <= 12 " 
			
set rsAmt = server.CreateObject("Adodb.recordset")
rsAmt.Open qryAmt, connstr
	if Not rsAmt.Eof then
	Amount_LocationCode = cdbl(to_nol(rsAmt("amount")))
	else
	Amount_LocationCode = 0
	End if
End function

Sub InsertDatafixedasset_mature_new_dtl
Dim Amt 
Amt = Amount_LocationCode(LocationCode)
    
	qq = " INSERT INTO fixedasset_mature_new_dtl ( documentno" & _
		     "                                    , fieldcode" & _
		     "                                    , amount" & _
		     "                                    , remarks" & _
		     "                                    , fcentry" & _
		     "                                    , fcedit" & _
		     "                                    , fcip" & _
		     "                                    , fcba" & _
		     "                                    , lastupdate" & _
		     "                                    , lasttime " & _
			 "                                    ) " & _
		     "   VALUES ( '" & fcCode & "'" & _
	         "          , '" & LocationCode &  "'" & _
	         "          , REPLACE('" & Amt & "', ',', '')" & _
	         "          , '" & Reference & "'" & _
	         "          , '" & Session("CodeName") & "'" & _
	         "          , '" & Session("CodeName") & "'" & _
	         "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
	         "          , '" & fcba & "'" & _
	         "          , SYSDATE" & _
	         "          , TO_CHAR(SYSDATE,'HH24:MI') " & _ 
			 " )"
		ExecuteAndLog ScriptName, qq
End Sub

Sub UpdateDatafixedasset_mature_new_dtl
		Call DeleteDatafixedasset_mature_new_dtl
		Call InsertDatafixedasset_mature_new_dtl
End Sub
		
Cek_IsAccClosing = StatusAccClosing(ACCOUNTINGYEAR,ACCOUNTINGPERIOD,fcba)
'response.write "Final =|" & Cek_IsAccClosing & "|" & fcba & "|" & DATECREATED

if 	Cek_IsAccClosing = "FALSE" then 
'--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "StartAuthorize" or Request("txtJenis")= "Approved" or Request("txtJenis")= "Posting" or Request("txtJenis")= "UnPosting" Then	
			if Request("txtJenis")= "Header" then
					
					q = " Delete from fixedasset_mature_new_dtl where  documentno = '" & fcCode & "' and fcBA = '" & fcba & "'"
					ExecuteAndLog ScriptName, q
					
					q = " Delete from fixedasset_mature_new where  documentno = '" & fcCode & "' and fcBA = '" & fcba & "'"
					ExecuteAndLog ScriptName, q
					
					q = " Delete fixedasset where  fcba='" & fcba & "' AND fccode='" & fccode & "'"
					ExecuteAndLog ScriptName, q
				
				cmdView	= "True"
			   
			elseif Request("txtJenis")= "Details" then
				
					call DeleteDatafixedasset_mature_new_dtl
				
				cmdView	= "True"
			
			elseif Request("txtJenis")="New" then
				
							q	=	" MERGE INTO fixedasset_mature_new a " & _ 
							        " USING ( " & _ 
									" SELECT   '" & fcCode & "' documentno" & _
									"          , '" & assetcode & "' fixedassetcode" & _
									"          , TO_DATE('" & DATECREATED & "','mm/dd/yyyy') creationdate" & _
									"          , TO_NUMBER(" & ACCOUNTINGPERIOD & ") period_month" &_
									"          , TO_NUMBER(" & ACCOUNTINGYEAR & ") period_year" &_
									"          , '" & request("description") & "' description" & _
									"          , '" & Session("CodeName") & "' fcentry" & _
									"          , '" & Session("CodeName") & "' fcedit" & _
									"          , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
									"          , '" & fcba & "' fcba" & _
									"          , SYSDATE lastupdate" & _
									"          , TO_CHAR(SYSDATE,'HH24:MI') lasttime" & _ 
									"          , 'Planned' rowstateid" & _
									"          , 0 postingflg FROM DUAL ) b ON (a.documentno=b.documentno and a.fixedassetcode=b.fixedassetcode and a.period_month=b.period_month and a.period_year=b.period_year and a.fcba=b.fcba) " & _
									" WHEN MATCHED THEN UPDATE " & _
									"  	SET a.creationdate = b.creationdate,a.fcedit=b.fcedit,a.lastupdate=b.lastupdate,a.lasttime=b.lasttime " & _
									" WHEN NOT MATCHED THEN " & _
									" INSERT ( a.documentno, a.fixedassetcode, a.creationdate, a.period_month, a.period_year, a.description , a.fcentry, a.fcedit, a.fcip, a.fcba, a.lastupdate , a.lasttime, a.rowstateid, a.postingflg )" & _
									" VALUES ( b.documentno, b.fixedassetcode, b.creationdate, b.period_month, b.period_year, b.description , b.fcentry, b.fcedit, b.fcip, b.fcba, b.lastupdate , b.lasttime, b.rowstateid, b.postingflg )" 
									
									ExecuteAndLog ScriptName, q
									
								Call InsertDatafixedasset_mature_new_dtl
						
				cmdView	= "True"
			elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
					
					Call UpdateDatafixedasset_mature_new_dtl
					cmdView	= "True"
				 
			elseif Request("txtJenis")= "Approved" then
			   q =	" Update fixedasset_mature_new" & _
					" set rowstateid = 'Approved', postingflg=1 " & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"     , lastupdate = sysdate " & _
					"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					" 		where documentno = '" & fccode & "'" & _
					"       and fcBA = '" & fcba & "'"
				ExecuteAndLog ScriptName, q
			
		       cmdView	= "True"
			
			elseif Request("txtJenis")= "RePlanned" then
			   
			   q =	" Update fixedasset_mature_new" & _
					" set rowstateid = 'Planned', postingflg=0 " & _
					"     , fcedit = '" & Session("CodeName") & "'" & _
					" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					"     , lastupdate = sysdate " & _
					"     , lasttime = to_char(sysdate,'hh24:mi') " & _
					" 		where documentno = '" & fccode & "'" & _
					"       and fcBA = '" & fcba & "'"
				ExecuteAndLog ScriptName, q
			
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

Dim rsVoucher
	

q = " SELECT " & _
	"         to_char(creationdate,'mm/dd/yyyy') datecreated" & _
	"        , fixedassetcode" & _
	"        , description" & _
	"        , rowstateid" & _
	"        , period_month" & _
	"        , period_year" & _
	"        , period_year||LPAD(period_month,2,0) accountingyearperiod" & _
	" FROM fixedasset_mature_new" & _
	" WHERE documentno = '" & fcCode & "'" & _
	"       AND fcBA = '" & fcba & "'"
'response.write q
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
	DATECREATED=rsVoucher("DateCreated")
	HeaderRemarks=rsVoucher("description")	
	rowstateid=rsVoucher("rowstateid")	
	ACCOUNTINGYEAR =rsVoucher("period_year")
	ACCOUNTINGPERIOD =rsVoucher("period_month")
	ACCOUNTINGYEARPERIOD = rsVoucher("accountingyearperiod")
	assetcode = rsVoucher("fixedassetcode")
End if

if Request("txtJenis") = "Edit" then
   
   q = " SELECT documentno" & _
		"        , fieldcode" & _
		"        , amount" & _
		" FROM fixedasset_mature_new_dtl" & _
		" WHERE fcba = '" & fcba & "'" & _
		"       AND documentno = '" & fcCode & "'" 
		
	set rsVoucher = server.CreateObject("Adodb.recordset")
	rsVoucher.Open q, connstr
	
	if Not rsVoucher.Eof then
		LocationCode = rsVoucher("fieldcode")
		Amount = rsVoucher("amount")
		End if
end if

rsVoucher.Close
set rsVoucher = Nothing

if err.number <> 0 then
   response.Write  "<div class='error'>" & err.description & "</div>"
   cmdView	= "True"
End if

' Untuk Cek Start Authorize By Sendy 04/29/2019
set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'FA'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND aus.fcusercode  = alu.fcusercode " & _
	"       AND alu.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & fcba & "'"
' response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	can_startauthorize =  cdbl(to_nol(rsFrez("total")))
else
	can_startauthorize = 0
End if
rsFrez.Close : set rsFrez = Nothing 
' End Added

Dim val_IsTrue : val_IsTrue = "N"
val_IsTrue = "Y"

 '-------------ISI Payment Form
%>
                   
                    <form name="frmEntry" method="post" action="fixedasset_mature.Asp">
					
                    <font color="red">
                        <%= varStringError  %></font>
                    <table class="data" width="100%">
						<tr>
                            <td align="left">
                                Date
                            </td>
                            <td align="left">
                                <input class="datepicker" type="text" maxlength="10" onchange="javascript:reloadCode('#');" name="DATECREATED" id="DATECREATED" size="10" value="<%=DATECREATED%>">
								<font color="#FF0000">&nbsp;*</font>
								<input type="text" maxlength="30" name="fcba" id="fcba" size="5" value="<%=fcba%>" style="background-color:#eee;" readonly>
                            </td>
                        </tr>							
						<tr>
                            <td align="left">
                                Accounting Period
                            </td>
							
							<% 	Dim rsComboA,Data_ACCOUNTINGYEARPERIOD

								q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from  ACCOUNTINGYEARPERIOD where accountingperiod <=12 and fcBA='" &  fcba & "'  and accountingyear >= to_number('"& Year(Date) &"')-1   order by accountingyear||lpad(accountingperiod,2,0)"
							
								Set rsComboA = Server.CreateObject("ADODB.Recordset")
								rsComboA.CursorType = 3
								rsComboA.CursorLocation = 3
								rsComboA.Open q, connstr
								if not rsComboA.eof then%>
								<td align="left"> 
									<select name="ACCOUNTINGYEARPERIOD" id="ACCOUNTINGYEARPERIOD" onchange="javascript:reloadCode('#');">    
										<% Do While Not rsComboA.eof %>
										<% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then 
											Data_ACCOUNTINGYEARPERIOD = "true"
										else
											Data_ACCOUNTINGYEARPERIOD = "false"
										End if%>
										<option value="<%=rsComboA("ACCOUNTINGYEARPERIOD")%>" <% If Request("ACCOUNTINGYEARPERIOD") = rsComboA("ACCOUNTINGYEARPERIOD") Then %> Selected <%End if%>><%=rsComboA("Description")%></option> 
										<%  rsComboA.movenext
										   loop
									if Data_ACCOUNTINGYEARPERIOD = "false" then %>
										<option value="<%=ACCOUNTINGYEAR&ACCOUNTINGPERIOD%>" Selected>Tahun <%=ACCOUNTINGYEAR%> Bulan <%=ACCOUNTINGPERIOD%></option>
									<% end if %>
									
									<%
										End if
										rsComboA.Close
										set rsComboA = Nothing
										q=""
										%>
									</select> <font color="#FF0000">*</font>
								</td>
                       	</tr>
						
                        <tr>
                            <td align="left">
                                Document.No
                            </td>
                            <td align="left">
							<%
								    var_txtTable = "ASSETMATURITY"
									if fcCode = "" then
									'response.write xCode
								    'MODIFY AMH - Apr 21 penomoran sesuai bank
										q = " select   NVL(MAX(cn.PREFIX),'" & fcba & "'||'/FA-MTR/'||'" & ACCOUNTINGPERIOD & "'||'/') || '" & ACCOUNTINGYEAR & "' ||'/'|| LPAD(NVL(MAX(regexp_substr(det.documentno, '[^/]+', 1, 5)),0000)+0001,4,0) DOCNUMBER " & _
											" from " & _
											" controlnumber cn " & _
											" left join fixedasset_mature_new det on (cn.fcba = det.fcba and cn.accountingyear = det.period_year) " & _
											" where cn.fcba ='" & fcba & "' " & _
											" and cn.LU_NAME ='"& var_txtTable &"' "& _ 
											" and cn.accountingyear =  " & ACCOUNTINGYEAR & " " & _
											" and cn.accountingperiod = " & ACCOUNTINGPERIOD & "" & _
											" and LPAD(det.documentno,Length(rtrim(det.documentno))-4) = cn.PREFIX || '" & ACCOUNTINGYEAR & "' ||'/' "
										'response.write q 
										set rsQuery = server.CreateObject("Adodb.recordset")
										rsQuery.Open q, connstr
										if Not rsQuery.Eof then
										
											if request("ACCOUNTINGYEARPERIOD") <>"" then
												fcCode = rsQuery("DOCNUMBER")
												assetcode= fcCode
											else
												fcCode = "(NEW)"
											end if
										End if										
									End if
									
									'response.write request("bankcodeurl") 
							%>
                                <input type="text" maxlength="30" name="fcCode" size="30" value="<%=fcCode%>" title="">
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                        </tr>
                        
                        <tr>
							<td align="left">Fixed Asset Code</td>
								<td align="left">
									<input type="text" maxlength="30" id="assetcode" name="assetcode" size="30" value="<%=assetcode%>" style="background-color:#eee;" readonly >
									<font color="#FF0000">*</font>
								</td>	
						</tr>
						<tr>
								<td align="left">
									Status
								</td>
								<td align="left">
									<input type="text" maxlength="50" name="rowstateid" size="30" value="<%=rowstateid%>" style="background-color:#eee;" readonly>
								</td>
						</tr>
						<tr>
                            <td align="left">
                                Description
                            </td>
                            <td align="left">
							    <textarea id="Description" name="Description"  cols="50" rows="3"><%=Description%></textarea>
                            </td>
                    </tr>
					<tr>
						<td colspan="2" align="left">
							<input type="Hidden" value="<%=ACCOUNTINGYEAR%>"   id="ACCOUNTINGYEAR" name="ACCOUNTINGYEAR">
							<input type="Hidden" value="<%=ACCOUNTINGPERIOD%>" id="ACCOUNTINGPERIOD" name="ACCOUNTINGPERIOD">
							<input type="Hidden" value="True" name="cmdView">
							<input type="Hidden" value="<%=request("txtTable")%>" name="txtTable">
							<input type="Hidden" value="<%=request("txtModule")%>" name="txtModule">
							<input type="Submit" name="Add" value="Select">&nbsp;&nbsp;							
							<%
							txtFileReportName = "Print_FixedAsset_Mature"
							%>
							
							<% if RowStateId = "Planned" and can_startauthorize <>  1 then %>
							     <!-- Modify AMH Mar 21 validasi balance -->
								 <a href="?fcCode=<%=fcCode%>&fcba=<%=fcba%>&txtJenis=Approved" onClick="return confirm('Are you sure to Approved this data :  <%=fcCode%> ?');"> Approved </a>  
							<% end if%>
							
							<% if RowStateId = "Approved" and can_startauthorize <> 1 then %> 
							     <a href="<%=txtFileReportName%>.asp?fcba=<%=fcba%>&fccode=<%=fcCode%>" target="_blank">
								<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
								</a>&nbsp;&nbsp;
								 <a href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&ACCOUNTINGYEARPERIOD=<%=ACCOUNTINGYEARPERIOD%>&fcba=<%=fcba%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to Re-Planned this data :  <%=fcCode%> ?');">Re-Planned</a>  
							
							<% end if%>
							
						</td>
					</tr>
			</table>
        </form>

<%if cmdView = "True" and "" & fcCode <> "" then 
	Dim rsfixedasset_mature_new, intRecordsPerPage, cColor,totAmt
	totAmt = 0

	q = " select pv.documentno" & _
		"        , pv.fixedassetcode" & _
		"        , pv.description" & _
		"        , creationdate" & _
		"        , period_month" & _
		"        , period_year" & _		
		"        , pv.fcba" & _		
		"        , pvd.fieldcode" & _
	    "        , pvd.amount" & _
		"        , pvd.remarks" & _
		" from fixedasset_mature_new pv" & _
	    " inner join fixedasset_mature_new_dtl pvd ON (pv.documentno = pvd.documentno and pv.fcba = pvd.fcba)" & _
	    " where pv.documentno = '" & fcCode & "'" & _
		"       and pv.fcba = '" & fcba & "'" & _
		" order by pv.fcba, pv.documentno, pvd.fieldcode"
	'response.Write q	
	Set rsfixedasset_mature_new = Server.CreateObject("ADODB.Recordset")
	rsfixedasset_mature_new.CursorType = 3
	rsfixedasset_mature_new.CursorLocation = 3
	rsfixedasset_mature_new.Open q, connstr
%>
        <form action="fixedasset_mature.Asp" method="post" name="eForm" style="display:table;">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" border="0">
            <tr>
                <td align="left">
                    Block
                </td>
                <td align="left">
                    <input type="text" size="11" id="locationcode" id="locationcode" name="LocationCode" value="<%=LocationCode%>" >&nbsp;<font
                        color="#FF0000">*</font><a href="javascript:PopupTType('<%=fcba %>')">
                            <font face="Verdana" size="1">Search</font><span class="satuan_item" id="locationname">&nbsp;</span></a>
                </td>
            </tr>
			
			
            <tr style="display: table-row;">
                <td align="left">
                    Remark
                </td>
                <td align="left">
                    <textarea name="Reference" cols="50"><%=Remarks%></textarea>
                </td>
            </tr>
            
            </td></tr>
            <tr>
                <td colspan="2" align="left">
                    <!-- diambil dari form yang diatas -->
                    <input type="hidden" name="fcCode" value="<%=fcCode%>">
                    <input type="hidden" name="cekno" value="<%=cekno%>">
                    <input type="hidden" name="referenceno" value="<%=referenceno%>">
					<input type="hidden" name="ACCOUNTINGYEARPERIOD" value="<%=ACCOUNTINGYEARPERIOD%>">
                    <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					<input type="hidden" name="assetcode" value="<%=assetcode%>">
					
                    <input type="hidden" name="xCode" value="<%=xCode%>">
                    <input type="hidden" name="HeaderRemarks" value="<%=HeaderRemarks%>">
					<input type="hidden" name="rowstateid" value="<%=rowstateid%>">
                    <input type="hidden" name="document_classification" value="<%=document_classification%>">
					<input type="Hidden" value="<%=request("txtTable")%>" name="txtTable">
                    <input type="Hidden" value="<%=request("txtModule")%>" name="txtModule">
					
				
					<!-- Period Control In Here -->
					<%IF Cek_IsAccClosing= "FALSE" then %>
						<%IF Request("txtJenis")="Edit" then%>
							<input type="hidden" name="txtJenis" value="Edit">
							<input Type="Submit" name="Add" value="Update">
						<%Else%>
							<% if RowStateId = "" OR RowStateId = "Planned" then %>
							   <input type="hidden" name="txtJenis" value="New">
							   <input Type="Submit" name="Add" value="Add">
							<%Else%>
						<%End if%>
					<%End if%>
							
						<!-- End Period Control In Here -->
						  <input Type="Reset" name="Reset" value="Reset">
						  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
					<%End if%>
                </td>
            </tr>
        </table>
        <table class="data xthf-site">
            <br>
            <thead>
            <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
                <th align="left">
                    NO
                </th>
                <th colspan="2">&nbsp;
                </th>
                <th align="left">
                    BLOCK
                </th>
                <th align="left">
                    AMOUNT
                </th>
                
		        <th align="left">
                    REMARKS
                </th>
                
				<!--<th align="left">
                    JENIS INVOICE 
                </th>-->
            </tr>
            </thead>
            <tbody>
            <%
   if Not rsfixedasset_mature_new.eof then
      nCountofField=rsfixedasset_mature_new.fields.count-1
	  no=1
	  While Not rsfixedasset_mature_new.EOF
          
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
            <tr bgcolor="<%=cColor%>">
                <td align="left" valign="top">
                    <%=NO%>
                </td>
				<!-- Cek Data Close or Not -->
				<!-- Period Control In Here -->
					<% if RowStateId = "Planned" then %>
						<td valign="middle" width="10">
							<a href="fixedasset_mature.Asp?fcCode=<%=rsfixedasset_mature_new("documentno")%>&LocationCode=<%=rsfixedasset_mature_new("fieldcode")%>&fcba=<%=fcba%>&txtJenis=Edit">
								<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsfixedasset_mature_new(1)%>" border="0">
							</a>
						</td>
						<td>
							<a href="fixedasset_mature.Asp?fcCode=<%=rsfixedasset_mature_new("documentno")%>&LocationCode=<%=rsfixedasset_mature_new("fieldcode")%>&fcba=<%=fcba%>&txtJenis=Details">
								<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsfixedasset_mature_new(1)%>" border="0" onclick="return confirm('Are you sure to delete');">
							</a>
						</td>
					<%Else%>
						<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
					<%End if%>
               <!-- End Period Control In Here -->
			    <td align="left" ><%=rsfixedasset_mature_new("fieldcode")%></td>
				<td align="right"><%=FormatNumber(rsfixedasset_mature_new("amount"),2)%></td>
				<td align="left" ><%=rsfixedasset_mature_new("remarks")%></td>
				
            </tr>
            <%
	      totAmt = totAmt + cdbl(rsfixedasset_mature_new("amount")) 		
		  rsfixedasset_mature_new.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsfixedasset_mature_new.Close
   set rsfixedasset_mature_new = Nothing%>
   <tr bgcolor="<%=cColor%>">
   <td colspan="4" align="left" >Total</td>
   <td align="right"><%=FormatNumber(totAmt,2)%></td>
   <td align="left" ></td>
   </tr>
   </tbody>
        </table>
        
        </form>
        <%End if%>
				<!-------------akhir!------------->
			</div>
		</div>
		<!--#include file="bottom.jsp"-->

<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="plugins/chosen/chosen.jquery.min.js"></script>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script>

function getselectedjobcode() {
	var ssupplier = document.getElementById("selectedjobcode");
	ssupplier.value = document.getElementById("jobcode").value;
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
	//getLastNumber();
	document.getElementById("chargeunit").style.display="none";
	document.getElementById("chargejob").style.display="none";
	document.getElementById("chargetype").style.display="none";
	document.getElementById("chargecode").style.display="none";
	getJobCode();
	getselectedjobcode();
	getJobcodeIntraco();
	getTragettype();
	
});

function getViewcharge(){
				var kodeintraco = $("#selectedjobcode").val();
			}


</script>
		
    </body>
</html>