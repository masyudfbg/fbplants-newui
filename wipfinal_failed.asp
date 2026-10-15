<%
	validatesession
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
	<title>ePlaS - Final Work in Progress</title>
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
 <script>
$(function() {
	//$( ".datepicker" ).datepicker();
	$("#tglawal").datepicker();
	$("#tglakhir").datepicker();
	$("#txtwipdate").datepicker();
	
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

<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.]/g,'');}
	
	
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

var getlistkontrak = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("Kontraktorlist").value;
	var selectedkontrak = document.getElementById("selectedkontrak").value;
	var selectedfcba = document.getElementById("fcba").value;
	
	var PostStr = 'contractor='+fccompanycode+'&txtcari='+selectedkontrak+'&fcBA='+selectedfcba;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistagreementcode.asp', 'POST', PostStr, showlistkontrak);
}

// The function for handling the response from the server
var showlistkontrak = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("Kontraklist").innerHTML = "";
	document.getElementById("Kontraklist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
}



</script>

</head>
<body>
	<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;">
<%
'--Awal--
Dim v_txtagreementcode, v_txtbappcode ,v_txtmonth, v_txtyear, MinYr, MaxYr, mon, Yr, n, SudahClose, v_txtagreementfinalno, v_txttglawal,v_txttglakhir
Dim xCode,xstoreCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,v_CrossCekPO
Dim Remarks
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
Dim Jobcode, LocationType, LocationCode, Reference, Remark
Dim qLineMax, rsqLineMax, maxagreementfinalno
Dim qTotalRecordOnHand, rsTotalRecordOnHand, tempTotalRecordOnHand
Dim qQuantityOnHand, rsQuantityOnHand, tempQuantityOnHand
Dim sqlSQLCheck, rsSQL, qtyRemained
Dim qCheckStore, rsqCheckStore
Dim txtLocationType	
Dim rsCheckRate, checkRate, varRate
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
Dim wipdate 
Dim isNew

Dim stringduedate : stringduedate = 0

Dim IsOverDuedate, varRemainDueDate, varDueDate, CompleteStringError, varErrorDueDate, stringremainqty, varRemainQty, IsOverRemainQty, varErrorRemainQty, varEditedQty

Dim qtyplanedit : qtyplanedit = Request("qtyplanedit")
Dim unitplanedit : unitplanedit = Request("unitplanedit")
Dim outputplanedit : outputplanedit = Request("outputplanedit")
Dim qtyfinaledit : qtyfinaledit = Request("qtyfinaledit")
Dim unitfinaledit : unitfinaledit = Request("unitfinaledit")
Dim outputfinaledit : outputfinaledit = Request("outputfinaledit")
Dim txtagreementlineno : txtagreementlineno = Request("txtagreementlineno")

Dim txtGRNCode : txtGRNCode = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
Dim txtLocationCode : txtLocationCode = ""

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")


v_bappno           = Request("txtbappno")
v_txtagreementcode = Request("txtagreementcode")
v_txttglawal = Request("tglawal")
v_txttglakhir = Request("tglakhir")
v_txtmonth = Request("txtmonth")
v_txtyear = Request("txtyear")
v_txtagreementfinalno = Request("txtagreementfinalno")
txtKontraktorcode = Request("Kontraktorlist")
selectedkontrak = Request("selectedkontrak")
kontrakcode = Request("kontrakcode")

response.write "<script type='text/javascript'>alert('2')</script>"

if v_txtagreementcode = "" then
v_txtagreementcode = selectedkontrak
end if 


 if request("txtwipdate") <>"" then
   wipdate   = request("txtwipdate")
 else
   wipdate = textDate(date)
 end if

Yr = Year(wipdate)
mon = Month(wipdate)
MinYr = Yr - 2
MaxYr = Yr  + 2

If len(v_txtyear)= 0 then v_txtyear = Yr
If len(v_txtmonth) = 0 then v_txtmonth = mon

'wipdate = v_txtmonth & "/" & GetLastDay(v_txtmonth, v_txtyear) & "/" & v_txtyear

txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
	
Sub insertdatawipfinal
	' Insert Data to table workinprogressfinal
	q = " insert into workinprogressfinal (agreementfinalno" & _
	    "                                  , agreementcode" & _
		"                                  , wipdate" & _
		"                                  , agreementlineno" & _
		"                                  , year" & _
		"                                  , month" & _
		"                                  , qty" & _
		"                                  , unit" & _
		"                                  , output" & _
		"                                  , isprocessing" & _
		"                                  , fcentry" & _
		"                                  , fcedit" & _
		"                                  , fcip" & _
		"                                  , fcba" & _
		"                                  , lastupdate" & _
		"                                  , lasttime" & _
		"                                  , chargeunit)" & _
		" values ('" & maxagreementfinalno & "'" & _
        "         , '" & v_txtagreementcode & "'" & _
        "         , to_date('" & request("wipdate_"&TRIM(n)) & "', 'mm/dd/yyyy')" & _
        "         , '" & request("agreementlineno_"&TRIM(n)) & "'" & _
        "         , '" & v_txtyear & "'" & _
        "         , '" & v_txtmonth & "'" & _
        "         , '" & Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) & "'" & _
		"         , '" & Cdbl(To_Nol(request("unitfinal_"&TRIM(n)))) & "'" & _
        "         , '" & Cdbl(To_Nol(request("outputfinal_"&TRIM(n)))) & "'" & _
        "         , 'TRUE'" & _
        "         , '" & Session("CodeName") & "'" & _
        "         , '" & Session("CodeName") & "'" & _
        "         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
        "         , '" & Session("fcBA") & "'" & _
        "         , sysdate " & _
        "         , to_char(sysdate, 'hh24:mi')" & _
		"         , '" & request("chargeunit_"&TRIM(n)) & "'" & _
        " )"
			
	'Response.Write q
	connStr.execute(q)
	
	' Insert Data to table workinprogressplan
	q = " insert into workinprogressplan (agreementcode" & _
	    "                                 , wipdate" & _
	    "                                 , agreementlineno" & _
	    "                                 , lineno" & _
	    "                                 , qty" & _
	    "                                 , unit" & _
		"                                 , output" & _
	    "                                 , postedflag" & _
	    "                                 , isgenerated" & _
	    "                                 , fcentry" & _
	    "                                 , fcedit" & _
	    "                                 , fcip" & _
	    "                                 , fcba" & _
	    "                                 , wipamt" & _
	    "                                 , reference" & _
	    "                                 , remarks" & _
	    "                                 , lastupdate" & _
	    "                                 , lasttime" & _
	    "                                 , agreementfinalno" & _
	    "                                 , jobcode" & _
	    "                                 , locationtype" & _
	    "                                 , locationcode" & _
		"                                 , chargeunit)" & _		
		" values ('" & v_txtagreementcode & "'" & _
        "         , to_date('" & request("wipdate_"&TRIM(n)) & "', 'mm/dd/yyyy')" & _
        "         , '" & request("agreementlineno_"&TRIM(n)) & "'" & _
		"         , 99999" & _
        "         , '" & ( Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) - Cdbl(To_Nol(request("qtyplan_"&TRIM(n)))) ) & "'" & _
		"         , '" & ( Cdbl(To_Nol(request("unitfinal_"&TRIM(n)))) - Cdbl(To_Nol(request("unitplan_"&TRIM(n)))) ) & "'" & _
		"         , '" & ( Cdbl(To_Nol(request("outputfinal_"&TRIM(n)))) - Cdbl(To_Nol(request("outputplan_"&TRIM(n)))) ) & "'" & _
		"         , 0" & _
		"         , 'TRUE'" & _
        "         , '" & Session("CodeName") & "'" & _
        "         , '" & Session("CodeName") & "'" & _
        "         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
        "         , '" & Session("fcBA") & "'" & _
		"         , '" & ( Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) - Cdbl(To_Nol(request("qtyplan_"&TRIM(n)))) ) * varRate & "'" & _
		"         , null" & _
		"         , null" & _
        "         , sysdate " & _
        "         , to_char(sysdate, 'hh24:mi')" & _
		"         , '" & maxagreementfinalno & "'" & _
		"         , '" & request("jobcode_"&TRIM(n)) & "'" & _
		"         , '" & request("targettype_"&TRIM(n)) & "'" & _
		"         , '" & request("targetcode_"&TRIM(n)) & "'" & _
		"         , '" & request("chargeunit_"&TRIM(n)) & "'" & _
        " )"
			
	'Response.Write q
	connStr.execute(q)
	
	' Insert Data to table workinprogress
	qq = " insert into workinprogress (agreementcode" & _
	    "                             , wipdate" & _
	    "                             , agreementlineno" & _
	    "                             , qty" & _
	    "                             , postedflag" & _
	    "                             , progresspaymentdate" & _
	    "                             , fcentry" & _
	    "                             , fcedit" & _
	    "                             , fcip" & _
	    "                             , fcba" & _
	    "                             , wipamt" & _
	    "                             , paid" & _
	    "                             , netbalance" & _
	    "                             , reference" & _
	    "                             , remarks" & _
	    "                             , lastupdate" & _
	    "                             , lasttime" & _
	    "                             , unit" & _
		"                             , output" & _
	    "                             , agreementfinalno" & _
		"                             , chargeunit)" & _
		" values ('" & v_txtagreementcode & "'" & _
        "         , to_date('" & request("wipdate_"&TRIM(n)) & "', 'mm/dd/yyyy')" & _
        "         , '" & request("agreementlineno_"&TRIM(n)) & "'" & _
        "         , '" & Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) & "'" & _
		"         , 0" & _
		"         , to_date('" & request("wipdate_"&TRIM(n)) & "', 'mm/dd/yyyy')" & _
        "         , '" & Session("CodeName") & "'" & _
        "         , '" & Session("CodeName") & "'" & _
        "         , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
        "         , '" & Session("fcBA") & "'" & _
		"         , '" & Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) * varRate & "'" & _
		"         , 0" & _
		"         , 0" & _
		"         , null" & _
		"         , null" & _
        "         , sysdate " & _
        "         , to_char(sysdate, 'hh24:mi')" & _
        "         , '" & Cdbl(To_Nol(request("unitfinal_"&TRIM(n)))) & "'" & _
		"         , '" & Cdbl(To_Nol(request("outputfinal_"&TRIM(n)))) & "'" & _
		"         , '" & maxagreementfinalno & "'" & _
		"         , '" & request("chargeunit_"&TRIM(n)) & "'" & _
        " )"

		'Response.Write qq
		connStr.execute(qq)


		' update is final pada tabel workinprogressplan 
		q = "update workinprogressplan  set isfinal=1 where agreementcode='" & v_txtagreementcode & "' and fcba='" & Session("fcBA") & "' and jobcode='" & request("jobcode_"&TRIM(n)) & "' AND locationcode='" & request("targetcode_"&TRIM(n)) & "'"
			
	'Response.Write q
	connStr.execute(q)
End Sub

Sub Updatedatawipfinal
	oSql = " select rate from contractagreementdetail " & _
				" where fcba = '" & Session("fcBA") & "'" & _
				"       and agreementcode = '" & v_txtagreementcode & "'" & _
				"       and lineno = '" & txtagreementlineno & "'"
				
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open oSql, connstr
	if Not oRs.Eof then
		varRate = cdbl(to_nol(oRs("rate")))
	Else
		varRate = 0
	End if	

	oRs.Close
	set oRs = Nothing
	
	'Update data from table workinprogress
	q = " Update workinprogress" & _
	    " set qty = '" & Cdbl(To_Nol(qtyfinaledit)) & "'" & _
		"     , unit = '" & Cdbl(To_Nol(unitfinaledit)) & "'" & _
		"     , output = '" & Cdbl(To_Nol(outputfinaledit)) & "'" & _
		"     , wipamt = '" & Cdbl(To_Nol(qtyfinaledit)) * varRate & "'" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	connStr.execute(q)
	
	'Update data from table workinprogressplan
	q = " Update workinprogressplan" & _
	    " set qty = '" & Cdbl(To_Nol(qtyfinaledit)) - Cdbl(To_Nol(qtyplanedit)) & "'" & _
		"     , unit = '" & Cdbl(To_Nol(unitfinaledit)) - Cdbl(To_Nol(unitplanedit)) & "'" & _
		"     , output = '" & Cdbl(To_Nol(outputfinaledit)) - Cdbl(To_Nol(outputplanedit)) & "'" & _
		"     , wipamt = '" & ( Cdbl(To_Nol(qtyfinaledit)) - Cdbl(To_Nol(qtyplanedit)) ) * varRate & "'" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	connStr.execute(q)
	
	'Update data from table workinprogressfinal
	q = " Update workinprogressfinal" & _
	    " set qty = '" & Cdbl(To_Nol(qtyfinaledit)) & "'" & _
		"     , unit = '" & Cdbl(To_Nol(unitfinaledit)) & "'" & _
		"     , output = '" & Cdbl(To_Nol(outputfinaledit)) & "'" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	connStr.execute(q)
End Sub

Sub deletedatawipfinal
	'Delete data from table workinprogress
	q = " delete from workinprogress" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	connStr.execute(q)
	
	'Delete data from table workinprogressplan
	q = " delete from workinprogressplan" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	connStr.execute(q)
	
	'Delete data from table workinprogressfinal
	q = " delete from workinprogressfinal" & _
        " where fcba = '" & Session("fcBA") & "'" & _
        "       and agreementfinalno = '" & v_txtagreementfinalno & "'"

	'Response.Write q
	connStr.execute(q)

	q = " update workinprogressplan set isfinal=0 where agreementcode='" & v_txtagreementcode & "'"
	connStr.execute(q)
End Sub

			
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Header" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), wipdate)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			
			set oRs = server.CreateObject("Adodb.recordset")
			oSql = " select agreementfinalno" & _
				   " from workinprogressfinal " & _
				   " where fcba = '" & Session("fcBA") & "'" & _
				   "       and agreementcode = '" & v_txtagreementcode & "'" & _
				   "  AND wipdate between to_date('" & v_txttglawal & "','mm/dd/yyyy') and to_date('" & v_txttglakhir & "','mm/dd/yyyy')"
			

			oRs.Open oSql, connstr
			do while not oRs.Eof
				'Delete data from table workinprogress
				q = " delete from workinprogress" & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and agreementfinalno = '" & ors("agreementfinalno") & "'"

			'	Response.Write q
				connStr.execute(q)
				
				'Delete data from table workinprogressplan
				q = " delete from workinprogressplan" & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and agreementfinalno = '" & ors("agreementfinalno") & "'"

				'Response.Write q
				connStr.execute(q)
				
				'Delete data from table workinprogressfinal
				q = " delete from workinprogressfinal" & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and agreementfinalno = '" & ors("agreementfinalno") & "'"

				'Response.Write q
				connStr.execute(q)

				'Delete data from table workinprogressfinal
				q = " update workinprogressplan set isfinal=0 " & _
					" where fcba = '" & Session("fcBA") & "'" & _
					"       and agreementcode = '" & v_txtagreementcode & "'"

				'Response.Write q
				connStr.execute(q)
				
				oRs.movenext
			loop
			oRs.Close
			Set oRs = Nothing

			if err.number = 0 then
				connStr.CommitTrans
			end if
		else
			CompleteStringError = "Tanggal """ & wipdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), wipdate)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			
			Call deletedatawipfinal
		
			if err.number=0 then
				connStr.CommitTrans
			End if
		else
			CompleteStringError = "Tanggal """ & wipdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), wipdate)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			
			stringduedate = Cstr("0")
			stringremainqty = Cstr("0")
			varErrorDueDate = ""
			
			For n = 1 to (cdbl(request("no")) - 1)
			
				'Cek Remain Qty
				oSql = " select fcba" & _
					   "        , agreementcode" & _
					   "        , lineno" & _
					   "        , sum (totalagreement) totalagreement" & _
					   "        , sum (totalwip) totalwip" & _
					   "        , sum (totalagreement) - sum (totalwip) selisih" & _
					   " from (select fcba" & _
					   "              , agreementcode" & _
					   "              , lineno" & _
					   "              , nvl(qty,0) totalagreement" & _
					   "              , 0 totalwip" & _
					   "       from contractagreementdetail" & _
					   "       where fcba = '" & session("fcBA") & "'" & _
					   "             and agreementcode = '" & v_txtagreementcode & "'" & _
					   "             and lineno = '" & request("agreementlineno_"&TRIM(n)) & "'" & _
					   "       union all" & _
					   "       select fcba" & _
					   "              , agreementcode" & _
					   "              , agreementlineno lineno" & _
					   "              , 0 totalagreement" & _
					   "              , nvl(qty,0) totalwip" & _
					   "       from workinprogress" & _
					   "       where fcba = '" & session("fcBA") & "'" & _
					   "             and agreementcode = '" & v_txtagreementcode & "'" & _
					   "             and agreementlineno = '" & request("agreementlineno_"&TRIM(n)) & "'" & _
					   " ) dataall" & _
					   " where 1=1" & _
					   " group by fcba" & _
					   "          , agreementcode" & _
					   "          , lineno"
				
				Set oRs = Server.CreateObject("ADODB.Recordset")
				oRs.Open oSql, connstr
				if Not oRs.Eof then
					varRemainQty = cdbl(to_nol(oRs("selisih")))
				Else
					varRemainQty = 0		   
				End if
				oRs.Close
				Set oRs = Nothing
				
				if ( varRemainQty - Cdbl(To_Nol(request("qtyfinal_"&TRIM(n)))) ) >= 0 then
					stringremainqty = stringremainqty & Cstr("0")
				else
					stringremainqty = stringremainqty & Cstr("1")
					if len(trim(varErrorRemainQty)) = 0 then
						varErrorRemainQty = "&nbsp;&nbsp;&nbsp;&nbsp;agreementcode-->" & v_txtagreementcode & ", lineno-->" & request("agreementlineno_"&TRIM(n)) & ", Quantity Maximal-->" & varRemainQty & "<br/>"
					else
						varErrorRemainQty = varErrorRemainQty & "&nbsp;&nbsp;&nbsp;&nbsp;agreementcode-->" & v_txtagreementcode & ", lineno-->" & request("agreementlineno_"&TRIM(n)) & ", Quantity Maximal-->" & varRemainQty & "<br/>"
					end if
				end if
				
				set oRs = server.CreateObject("Adodb.recordset")
				oSql = " SELECT dtl.rate, hdr.finishdate FROM contractagreement hdr INNER JOIN contractagreementdetail dtl ON (dtl.fcba = hdr.fcba AND dtl.agreementcode = hdr.agreementcode)" & _
					   " WHERE dtl.fcba = '" & Session("fcBA") & "'" & _
					   "       and dtl.agreementcode = '" & v_txtagreementcode & "'" & _
					   "       and dtl.lineno = '" & request("agreementlineno_"&TRIM(n)) & "'"

				oRs.Open oSql, connstr
				if Not oRs.Eof then
					varRate = cdbl(oRs("rate"))
					varDueDate = oRs("finishdate")
				Else
					varRate = 0
				End if
				oRs.Close
				Set oRs = Nothing
				
				set oRs = server.CreateObject("Adodb.recordset")
				oSql = "select nvl(to_number(to_date('" & wipdate & "','mm/dd/yyyy') - to_date('" &  varDueDate & "','mm/dd/yyyy')),0) as remaindate from dual "

				oRs.Open oSql, connstr
				if Not oRs.Eof then
					varRemainDueDate = cdbl(oRs("RemainDate"))
				Else
					varRemainDueDate = 0
				End if
				oRs.Close
				Set oRs = Nothing
				
				if varRemainDueDate <= 0 then
					stringduedate = stringduedate & Cstr("0")
				else
					stringduedate = stringduedate & Cstr("1")
					if len(trim(varErrorDueDate)) = 0 then
						varErrorDueDate = "&nbsp;&nbsp;&nbsp;&nbsp;agreementcode-->" & v_txtagreementcode & ", WIPDate" & textdate(wipdate) & " lineno-->" & request("agreementlineno_"&TRIM(n)) & ", Due Date-->" & textdate(varDueDate) & "<br/>"
					else
						varErrorDueDate = varErrorDueDate & "&nbsp;&nbsp;&nbsp;&nbsp;agreementcode-->" & v_txtagreementcode & ", lineno-->" & request("agreementlineno_"&TRIM(n)) & ", Due Date-->" & textdate(varDueDate) & "<br/>"
					end if
				end if
				
			Next
			
			set oRs = server.CreateObject("Adodb.recordset")
			oSql = "select count(1) total from dual where instr('" & stringduedate & "', '1') <> 0"
			oRs.Open oSql, connstr
			if Not oRs.Eof then
				IsOverDuedate = cdbl(oRs("total"))
			Else
				IsOverDuedate = 1
			End if
			oRs.Close
			Set oRs = Nothing
			
			set oRs = server.CreateObject("Adodb.recordset")
			oSql = "select count(1) total from dual where instr('" & stringremainqty & "', '1') <> 0"
			oRs.Open oSql, connstr
			if Not oRs.Eof then
				IsOverRemainQty = cdbl(oRs("total"))
			Else
				IsOverRemainQty = 1
			End if
			oRs.Close
			Set oRs = Nothing
			
			CompleteStringError = ""
			if (IsOverRemainQty = 0 And IsOverDuedate = 0) then
					set oRs = server.CreateObject("Adodb.recordset")
					oSql = " select nvl(max(agreementfinalno), 0) maxagreementfinalno" & _
						   " from workinprogressfinal " & _
						   " where fcba = '" & session("fcBA") & "'"

					oRs.Open oSql, connstr
					if Not oRs.Eof then
						maxagreementfinalno = cdbl(oRs("maxagreementfinalno"))
					else
						maxagreementfinalno = 0
					End if
					oRs.Close
					Set oRs = Nothing
					
					For n = 1 to (cdbl(request("no")) - 1)
						maxagreementfinalno = maxagreementfinalno + 1
						Call insertdatawipfinal
					Next
			elseif (IsOverRemainQty = 0 And IsOverDuedate = 1) then
				cmdView	= "True"
				CompleteStringError = "Tanggal WIP harus lebih kecil dari Tanggal DUEDATE pada Contract Aggreement" & "<br/>"
				CompleteStringError = CompleteStringError & varErrorDueDate
			elseif (IsOverRemainQty = 1 And IsOverDuedate = 0) then
				cmdView	= "True"
				CompleteStringError = "Total Quantity yang anda masukkan melebihi dari kontrak yang disepakati" & "<br/>"
				CompleteStringError = CompleteStringError & varErrorRemainQty
			else
				cmdView	= "True"
				CompleteStringError = "Total Quantity yang anda masukkan melebihi dari kontrak yang disepakati" & "<br/>"
				CompleteStringError = CompleteStringError & varErrorRemainQty
				CompleteStringError = CompleteStringError & "<br/>" & "Tanggal WIP harus lebih kecil dari Tanggal DUEDATE pada Contract Aggreement" & "<br/>"
				CompleteStringError = CompleteStringError & varErrorDueDate
			end if

			if err.number=0 then
				connStr.CommitTrans
			End if
		else
			CompleteStringError = "Tanggal """ & wipdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), wipdate)
		If Cek_IsDailyClosing = "FALSE" then
			'Cek Remain Qty
			oSql = " select fcba" & _
				   "        , agreementcode" & _
				   "        , lineno" & _
				   "        , sum (totalagreement) totalagreement" & _
				   "        , sum (totalwip) totalwip" & _
				   "        , sum (totalagreement) - sum (totalwip) selisih" & _
				   " from (select fcba" & _
				   "              , agreementcode" & _
				   "              , lineno" & _
				   "              , nvl(qty, 0) totalagreement" & _
				   "              , 0 totalwip" & _
				   "       from contractagreementdetail" & _
				   "       where fcba = '" & session("fcBA") & "'" & _
				   "             and agreementcode = '" & v_txtagreementcode & "'" & _
				   "             and lineno = '" & txtagreementlineno & "'" & _
				   "       union all" & _
				   "       select fcba" & _
				   "              , agreementcode" & _
				   "              , agreementlineno lineno" & _
				   "              , 0 totalagreement" & _
				   "              , nvl(qty, 0) totalwip" & _
				   "       from workinprogress" & _
				   "       where fcba = '" & session("fcBA") & "'" & _
				   "             and agreementcode = '" & v_txtagreementcode & "'" & _
				   "             and agreementlineno = '" & txtagreementlineno & "'" & _
				   " ) dataall" & _
				   " where 1=1" & _
				   " group by fcba" & _
				   "          , agreementcode" & _
				   "          , lineno"
			
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open oSql, connstr
			if Not oRs.Eof then
				varRemainQty = cdbl(oRs("selisih"))
			Else
				varRemainQty = 0		   
			End if
			oRs.Close
			Set oRs = Nothing
			
			'Cek Edited Qty
			oSql = " select nvl(sum (qty), 0) totalwipedit" & _
				   " from workinprogress" & _
				   " where fcba = '" & session("fcBA") & "'" & _
				   "       and agreementfinalno = '" & v_txtagreementfinalno & "'"
			
			Set oRs = Server.CreateObject("ADODB.Recordset")
			oRs.Open oSql, connstr
			if Not oRs.Eof then
				varEditedQty = cdbl(oRs("totalwipedit"))
			Else
				varEditedQty = 0		   
			End if
			oRs.Close
			Set oRs = Nothing
			
			if ((varRemainQty + varEditedQty) - Cdbl(To_Nol(qtyfinaledit)) ) >=  0 then
				connstr.BeginTrans
				Call Updatedatawipfinal
				if err.number=0 then
					connStr.CommitTrans
				End if
				cmdView	= "True"
				Response.Redirect "/wipfinal.asp?txtagreementcode=" & v_txtagreementcode & "&txtyear=" & v_txtyear & "&txtmonth=" & v_txtmonth & "&cmdView=True&SudahClose=Y"
			else
				cmdView	= "True"
				CompleteStringError = "Total Quantity yang anda masukan melebihi dari kontrak yang disepakati" & "<br/>"
				CompleteStringError = CompleteStringError & "&nbsp;&nbsp;&nbsp;&nbsp;Quantity Maximal-->" & (varRemainQty + varEditedQty)
			end if
		else
			CompleteStringError = "Tanggal """ & wipdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			'Response.Redirect "/wipfinal.asp?txtagreementcode=" & v_txtagreementcode & "&txtyear=" & v_txtyear & "&txtmonth=" & v_txtmonth & "&cmdView=True&SudahClose=Y"
		End If
		cmdView	= "True"
    else
	
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")		    
End if

SudahClose = "N"

set oRs = server.CreateObject("Adodb.recordset")
oSql = " select count(1) id" & _
       " from workinprogressfinal" & _
       " where fcba = '" & Session("fcba") & "'" & _
       "       and agreementcode = '" & v_txtagreementcode & "'" & _
       " AND wipdate between to_date('" & v_txttglawal & "','mm/dd/yyyy') and to_date('" & v_txttglakhir & "','mm/dd/yyyy')"

response.write "x-" & oSql 
oRs.Open oSql, connstr
if Not oRs.Eof then 
	if cdbl(oRs("id")) > 0 then 
	SudahClose = "Y"
	else
	SudahClose = "N"
	end if
End if
oRs.Close
Set oRs = Nothing

If Len(Trim(CompleteStringError)) > 0 Then
	response.Write  "<div class = ""error"">" & CompleteStringError & "</div>"
End If


Dim val_IsTrue : val_IsTrue = "Y"
' Dim oRecordSet, oSQLQuery
' oSQLQuery = " select 'Y' istrue" & _
            ' " from periodcontrol" & _
	        ' " where fcba = '" & Session("fcBA") & "' " & _
	        ' "       and to_date ('" & v_txtmonth & "/01/" & v_txtyear & "', 'mm/dd/yyyy') >= currentdate " & _
	        ' "       and upper(system) = upper('Contract')"
	   
' set oRecordSet = server.CreateObject("Adodb.recordset")
' oRecordSet.Open oSQLQuery, connstr
' If Not oRecordSet.Eof then
	' val_IsTrue = oRecordSet("istrue")
' Else
	' val_IsTrue = "N"
' End if
'response.Write strNurseryCulling
'oRecordSet.Close
'Set oRecordSet = Nothing
 %>
<form name="frmEntry" method="post" action="wipfinal.asp">
	<table class="data" width="100%">
	    <tr>
			<td align="left">No.BAPP</td>
			<td align="left">
				<input type="text" maxlength="150" name="txtbappno" size="30" value="<%=v_bappno%>" required>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Tgl.BAPP</td>
			<td align="left">
			    <input type="text" class="datepicker" name="txtwipdate" id="txtwipdate" value="<%=wipdate%>"> 
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Kontraktor </td>
			<td align="left">
			<select name="Kontraktorlist" value="<%=txtKontraktorcode%>" id="Kontraktorlist" style="width:200px" onchange="javascript:getlistkontrak();"  >
									<option></option>
									<%
									qLov = " SELECT fccode, fcname FROM contractor where fcba = '" & Session("fcba") & "' ORDER BY fcname"
						   
									Set rsLov = Server.CreateObject("ADODB.Recordset")
									rsLov.Open qLov, connstr
									if not rsLov.eof then
										while not rsLov.eof
									%>
									<option value="<%=rsLov("fccode")%>" <% if cstr(""&rsLov("fccode")) = cstr(""&txtKontraktorcode) then%> selected <%end if%>><%=rsLov("fcname")%></option>
									<%
									rsLov.movenext
									wend
									end if
									%>
				</select>	
			</td>
		</tr>
		
		<tr>
			<td align="left">No.Kontrak</td>
			<td align="left">
			    <select name="Kontraklist" id="Kontraklist" style="width:300px"  >
				</select>
				<input type="Hidden" id="selectedkontrak" name="selectedkontrak" value="<%=selectedkontrak%>">
				<input type="Hidden" id="fcba" name="fcba" value="<%=session("fcba")%>">
				
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Periode Daily</td>
			<td align="left">
				<input type="text" class="datepicker" name="tglawal" id="tglawal" value="<%=v_txttglawal%>"> s.d <input type="text" class="datepicker" name="tglakhir" id="tglakhir" value="<%=v_txttglakhir%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left"> 
				<input Type="Hidden" Value="True" Name="cmdView">  
				<input Type="Hidden" Value="<%=request("v_txtagreementcode")%>" Name="txtagreementcode_">  
				
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">  
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">   	 
				<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;  
				<a href="print_wipfinal.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtagreementcode=<%=v_txtagreementcode%>" target="_blank"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
				<%If val_IsTrue = "Y" Then%>
					<a href="wipfinal.asp?txtagreementcode=<%=v_txtagreementcode%>&txtyear=<%=v_txtyear%>&txtmonth=<%=v_txtmonth%>&wipdate=<%=wipdate%>&txtJenis=Header&tglawal=<%=v_txttglawal%>&tglakhir=<%=v_txttglakhir%>"><img src="IMAGES/delete.gif" alt="Delete Data : <%=v_txtagreementcode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End If%>
				<!-- HADIR -->
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Contract"
				dataSubmodule 	= "Work in Progress Final"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td> 
		</tr> 
	</table>    
</form>
<%
If cmdView = "True" and Len(Trim(v_txtagreementcode)) > 0 Then
	Dim rswipfinal, intRecordsPerPage, cColor
	set rswipfinal = server.CreateObject("Adodb.recordset")
	If SudahClose <> "Y" Then
		q	= " select wip.fcba" & _
              "        , wip.agreementcode" & _
			  "        , to_char(wip.wipdate,'mm/dd/yyyy') wipdate" & _
              "        , wip.agreementlineno" & _
              "        , wip.jobcode" & _
              "        , job.fcname jobname" & _
			  "        , wip.chargeunit" & _
              "        , wip.locationtype targettype" & _
              "        , wip.locationcode targetcode" & _
              "        , nvl(sum(wip.qty), 0) qtyplan" & _
			  "        , nvl(sum(wip.unit), 0) unitplan" & _
              "        , nvl(sum(wip.output), 0) outputplan" & _
              " from workinprogressplan wip" & _
              " left join job on (job.fccode = wip.jobcode and job.fcba = wip.fcba)" & _
              " where wip.fcba = '" & Session("fcba") & "'" & _
              "       and wip.agreementcode = '" & v_txtagreementcode & "'" & _
              " AND wip.wipdate between to_date('" & v_txttglawal & "','mm/dd/yyyy') and to_date('" & v_txttglakhir & "','mm/dd/yyyy')  AND wip.isfinal=0 " & _
              " group by wip.fcba" & _
              "          , wip.agreementcode" & _
              "          , wip.agreementlineno" & _
			  "          , wip.jobcode" & _
              "          , job.fcname" & _
			  "          , wip.chargeunit" & _
              "          , wip.locationtype" & _
              "          , wip.locationcode" & _
              "          , wip.wipdate" & _
			  " ORDER BY wip.fcba,wip.agreementlineno,wip.wipdate "

	Else	 
		q	= " select wipf.fcba" & _
              "        , wipf.agreementfinalno" & _
			  "        , wipf.agreementcode" & _
              "        , to_char(wipf.wipdate,'dd/mm/yyyy') wipdate" & _
			  "        , wipf.agreementlineno" & _
			  "        , wipf.year" & _
			  "        , wipf.month" & _
              "        , wipp.jobcode" & _
              "        , job.fcname jobname" & _
			  "        , wipp.chargeunit" & _
              "        , wipp.locationtype targettype" & _
              "        , wipp.locationcode targetcode" & _
              "        , nvl(sum(wipf.qty), 0) qtyfinal" & _
			  "        , nvl(sum(wipf.unit), 0) unitfinal" & _
              "        , nvl(sum(wipf.output), 0) outputfinal" & _
              " from workinprogressfinal wipf" & _
              " left join workinprogressplan wipp on (wipf.fcba = wipp.fcba and wipf.agreementfinalno = wipp.agreementfinalno)" & _
              " left join job on (job.fcba = wipp.fcba and job.fccode = wipp.jobcode)" & _
              " where wipf.fcba = '" & Session("fcba") & "'" & _
              "       and wipf.agreementcode = '" & v_txtagreementcode & "'" & _
              "       and wipf.wipdate between to_date('" & v_txttglawal & "','mm/dd/yyyy') and to_date('" & v_txttglakhir & "','mm/dd/yyyy') " & _
              " group by wipf.fcba" & _
              "          , wipf.agreementfinalno" & _
			  "          , wipf.agreementcode" & _
              "          , wipf.wipdate" & _
			  "          , wipf.agreementlineno" & _
			  "          , wipf.year" & _
			  "          , wipf.month" & _
			  "          , wipp.jobcode" & _
			  "          , wipp.chargeunit" & _
			  "          , job.fcname" & _
			  "          , wipp.locationtype" & _
              "          , wipp.locationcode" & _
			  " order by wipf.fcba" & _
			  "          , wipf.agreementcode" & _
              "          , wipf.year" & _
			  "          , wipf.month" & _
			  "          , wipf.agreementlineno,wipf.wipdate"
   End if

   If Request("txtJenis")="Edit" Then
		q = " select a.fcba" & _
			"                      , a.agreementfinalno" & _
			"                      , a.agreementcode" & _
			"                      , a.wipdate" & _
			"                      , a.agreementlineno" & _
			"                      , a.year" & _
			"                      , a.month" & _
			"                      , b.jobcode" & _
			"                      , b.jobname" & _
			"                      , b.chargeunit" & _
			"                      , b.locationtype targettype" & _
			"                      , b.locationcode targetcode" & _
			"                      , nvl(sum(b.qtyplan), 0) qtyplan" & _
			"                      , nvl(sum(b.unitplan), 0) unitplan" & _
			"                      , nvl(sum(b.outputplan), 0) outputplan" & _
			"                      , nvl(sum(a.qty), 0) qtyfinal" & _
			"                      , nvl(sum(a.unit), 0) unitfinal" & _
			"                      , nvl(sum(a.output), 0) outputfinal" & _
			"               from workinprogressfinal a" & _
			"                    , (select wipp.fcba" & _
			"                              , wipp.agreementcode" & _
			"                              , to_number(to_char(wipp.wipdate, 'yyyy')) year" & _
			"                              , to_number(to_char(wipp.wipdate, 'mm')) month" & _
			"                              , wipp.agreementlineno" & _
			"                              , wipp.jobcode" & _
			"                              , job.fcname jobname" & _
			"                              , wipp.chargeunit" & _
			"                              , wipp.locationtype" & _
			"                              , wipp.locationcode" & _
			"                              , nvl(sum(wipp.qty),0) qtyplan" & _
			"                              , nvl(sum(wipp.unit),0) unitplan" & _
			"                              , nvl(sum(wipp.output),0) outputplan" & _
			"                       from workinprogressplan wipp" & _
			"                       left join job on (job.fcba = wipp.fcba and job.fccode = wipp.jobcode)" & _
			"                       where wipp.agreementfinalno is null" & _
			"                             and wipp.fcba = '" & Session("fcba") & "'" & _
			"                             and wipp.agreementcode = '" & v_txtagreementcode & "' AND wipp.isfinal=1" & _
			"                        group by wipp.fcba" & _
			"                                 , wipp.agreementcode" & _
			"                                 , to_number(to_char(wipp.wipdate, 'yyyy'))" & _
			"                                 , to_number(to_char(wipp.wipdate, 'mm'))" & _
			"                                 , wipp.agreementlineno" & _
			"                                 , wipp.chargeunit" & _
			"                                 , wipp.jobcode" & _
			"                                 , job.fcname" & _
			"                                 , wipp.locationtype" & _
			"                                 , wipp.locationcode) b" & _
			"               where a. fcba = b.fcba" & _
			"                     and a.agreementcode = b.agreementcode" & _
			"                     and a.year = b.year" & _
			"                     and a.month = b.month" & _
			"                     and a.agreementlineno = b.agreementlineno" & _
			"                     and a.fcba = '" & Session("fcba") & "'" & _
			"                     and a.agreementcode = '" & v_txtagreementcode & "'" & _
			"                     and a.agreementfinalno = '" & v_txtagreementfinalno & "'" & _
			"               group by a.fcba" & _
			"                        , a.agreementfinalno" & _
			"                        , a.agreementcode" & _
			"                        , a.wipdate" & _
			"                        , a.agreementlineno" & _
			"                        , a.year" & _
			"                        , a.month" & _
			"                        , b.jobcode" & _
			"                        , b.jobname" & _
			"                        , b.chargeunit" & _
			"                        , b.locationtype" & _
			"                        , b.locationcode" & _
			"               order by a.fcba" & _
			"                        , a.agreementcode" & _
			"                        , a.year" & _
			"                        , a.month" & _
			"                        , a.agreementlineno"

	End If
			  
  response.Write q	
   rswipfinal.Open q, connstr%>
   <form action="wipfinal.asp" method="post" name="eForm">
   <table class="data" border="1" width="100%" >
   <%if SudahClose <> "Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
		   <td align="center">Line No</td>
		   <td align="center">Tanggal</td>
		   <td align="center">Charge Unit</td>
		   <td align="center">Jobcode</td>
		   <td align="center">Jobname</td>		   
		   <td align="center">Target Type</td>
		   <td align="center">Target Code</td>
		   <td align="center">Quantity (Plan)</td>
		   <td align="center">Unit (Plan)</td>
		   <td align="center">Output (Plan)</td>
		   <td align="center">Quantity (Final)</td>
		   <td align="center">Unit (Final)</td>
		   <td align="center">Output (Final)</td>
   	   </tr>
   <%Else%>
		<%if Request("txtJenis")="Edit" Then%>
   
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
				<td align="center">Line No</td>
				<td align="center">Tanggal</td>
				<td align="center">Charge Unit</td>
				<td align="center">Jobcode</td>
				<td align="center">Jobname</td>
				<td align="center">Target Type</td>
				<td align="center">Target Code</td>
				<td align="center">Quantity (Plan)</td>
				<td align="center">Unit (Plan)</td>
				<td align="center">Output (Plan)</td>
				<td align="center">Quantity (Final)</td>
				<td align="center">Unit (Final)</td>
				<td align="center">Output (Final)</td>
		   </tr>
		<%else%>
	   
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
				<td align="center" colspan="2">&nbsp;</td>
				<td align="center">Line No</td>
				<td align="center">Tanggal</td>
				<td align="center">Charge Unit</td>
				<td align="center">Jobcode</td>
				<td align="center">Jobname</td>
				<td align="center">Target Type</td>
				<td align="center">Target Code</td>
				<td align="center">Quantity (Final)</td>
				<td align="center">Unit (Final)</td>
				<td align="center">Output (Final)</td>
			</tr>
		<%End if%>

   <%End if
   if Not rswipfinal.eof then
      nCountofField=rswipfinal.fields.count-1
	  no = 1
	  While Not rswipfinal.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
			<tr bgcolor="<%=cColor%>">
				<%IF SudahClose = "Y" then%>

					<%if Request("txtJenis")="Edit" Then%>
						<td align="left" title="Agreement Line No" valign="top"><%=rswipfinal("agreementlineno")%></td>
						<td align="left" title="Tanggal" valign="top"><%=rswipfinal("wipdate")%></td>
						<td align="left" title="Chargeunit" valign="top"><%=rswipfinal("chargeunit")%></td>
						<td align="left" title="Job Code" valign="top"><%=rswipfinal("jobcode")%></td>
						<td align="left" title="Job Name" valign="top"><%=rswipfinal("jobname")%></td>
						<td align="left" title="Target Type" valign="top"><%=rswipfinal("targettype")%></td>
						<td align="left" title="Target Code" valign="top"><%=rswipfinal("targetcode")%></td>
						<td align="Right" title="Qty Plan" valign="top"><%=FbpNum2(rswipfinal("qtyplan"))%></td>
						<td align="Right" title="Unit Plan" valign="top"><%=FbpNum2(rswipfinal("unitplan"))%></td>
						<td align="Right" title="Output Plan" valign="top"><%=FbpNum2(rswipfinal("outputplan"))%></td>
						<td align="left" width="2">
							<input type="text" name="qtyfinaledit" width="2" value="<%=rswipfinal("qtyfinal")%>" onkeyup="checkInput(this);">
						</td>
						<td align="left" width="2">
							<input type="text" name="unitfinaledit" width="2" value="<%=rswipfinal("unitfinal")%>" onkeyup="checkInput(this);">
						</td>
						<td align="left" width="2">
							<input type="text" name="outputfinaledit" width="2" value="<%=rswipfinal("outputfinal")%>" onkeyup="checkInput(this);">
						</td>
						<input type="hidden" name="txtagreementfinalno" value="<%=rswipfinal("agreementfinalno")%>">
						<input type="hidden" name="txtagreementlineno" value="<%=rswipfinal("agreementlineno")%>">
						<input type="hidden" name="qtyplanedit" value="<%=rswipfinal("qtyplan")%>">
						<input type="hidden" name="unitplanedit" value="<%=rswipfinal("unitplan")%>">
						<input type="hidden" name="outputplanedit" value="<%=rswipfinal("outputplan")%>">
						<input type="hidden" name="chargeunitedit" value="<%=rswipfinal("chargeunit")%>">
						<input type="hidden" name="jobcodeedit" value="<%=rswipfinal("jobcode")%>">
						<input type="hidden" name="targettypeedit" value="<%=rswipfinal("targettype")%>">
						<input type="hidden" name="targetcodeedit" value="<%=rswipfinal("targetcode")%>">
					<%Else%>
				
						<%If val_IsTrue = "Y" Then%>
						<td valign="middle" width="10">
							<a href="wipfinal.asp?txtagreementcode=<%=rswipfinal("agreementcode")%>&txtagreementfinalno=<%=rswipfinal("agreementfinalno")%>&txtyear=<%=rswipfinal("year")%>&txtmonth=<%=rswipfinal("month")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rswipfinal(1)%>" border="0"></a>
						</td>
						<td valign="middle" width="10">
							<a href="wipfinal.asp?txtagreementcode=<%=rswipfinal("agreementcode")%>&txtagreementfinalno=<%=rswipfinal("agreementfinalno")%>&txtyear=<%=rswipfinal("year")%>&txtmonth=<%=rswipfinal("month")%>&wipdate=<%=wipdate%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rswipfinal(1)%>" border="0"></a>
						</td>
						<%Else%>
						<td colspan="2">
							<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
						</td>
						<%End If%>
						
						<td align="left" title="<%=rswipfinal(4).Name%>" valign="top"><%=rswipfinal(4)%></td>
						<td align="left" title="<%=rswipfinal(3).Name%>" valign="top"><%=rswipfinal(3)%></td>
						<td align="left" title="<%=rswipfinal(9).Name%>" valign="top"><%=rswipfinal(9)%></td>
						<td align="left" title="<%=rswipfinal(7).Name%>" valign="top"><%=rswipfinal(7)%></td>
						<td align="left" title="<%=rswipfinal(8).Name%>" valign="top"><%=rswipfinal(8)%></td>
						<td align="left" title="<%=rswipfinal(10).Name%>" valign="top"><%=rswipfinal(10)%></td>
						<td align="left" title="<%=rswipfinal(11).Name%>" valign="top"><%=rswipfinal(11)%></td>
						<td align="left" title="<%=rswipfinal(12).Name%>" valign="top"><%=rswipfinal(12)%></td>
						<td align="left" title="<%=rswipfinal(13).Name%>" valign="top"><%=rswipfinal(13)%></td>
						<td align="left" title="<%=rswipfinal(14).Name%>" valign="top"><%=rswipfinal(14)%></td>
					
					<%End If%>
					
				<%End if%>
				
				<%if SudahClose <> "Y" Then%>
					<td align="left" title="Agreement Line No" valign="top"><%=rswipfinal("agreementlineno")%></td>
					<td align="left" title="Tanggal" valign="top"><%=rswipfinal("wipdate")%></td>
					<td align="left" title="Chargeunit" valign="top"><%=rswipfinal("chargeunit")%></td>
					<td align="left" title="Job Code" valign="top"><%=rswipfinal("jobcode")%></td>
					<td align="left" title="Job Name" valign="top"><%=rswipfinal("jobname")%></td>
					<td align="left" title="Target Type" valign="top"><%=rswipfinal("targettype")%></td>
					<td align="left" title="Target Code" valign="top"><%=rswipfinal("targetcode")%></td>
					<td align="Right" title="Quantity Plan" valign="top"><%=FbpNum2(rswipfinal("qtyplan"))%></td>
					<td align="Right" title="Unit Plan" valign="top"><%=FbpNum2(rswipfinal("unitplan"))%></td>
					<td align="Right" title="Output Plan" valign="top"><%=FbpNum2(rswipfinal("outputplan"))%></td>
					<td align="left" width="2">
						<input type="text" name="qtyfinal_<%=no%>" width="2" value="<%=rswipfinal("qtyplan")%>" onkeyup="checkInput(this);">
					</td>
					<td align="left" width="2">
						<input type="text" name="unitfinal_<%=no%>" width="2" value="<%=rswipfinal("unitplan")%>" onkeyup="checkInput(this);">
					</td>
					<td align="left" width="2">
						<input type="text" name="outputfinal_<%=no%>" width="2" value="<%=rswipfinal("outputplan")%>" onkeyup="checkInput(this);">
					</td>
						<input type="hidden" name="wipdate_<%=no%>" value="<%=rswipfinal("wipdate")%>">
						<input type="hidden" name="agreementlineno_<%=no%>" value="<%=rswipfinal("agreementlineno")%>">
						<input type="hidden" name="qtyplan_<%=no%>" value="<%=rswipfinal("qtyplan")%>">
						<input type="hidden" name="unitplan_<%=no%>" value="<%=rswipfinal("unitplan")%>">
						<input type="hidden" name="outputplan_<%=no%>" value="<%=rswipfinal("outputplan")%>">
						<input type="hidden" name="chargeunit_<%=no%>" value="<%=rswipfinal("chargeunit")%>">
						<input type="hidden" name="jobcode_<%=no%>" value="<%=rswipfinal("jobcode")%>">
						<input type="hidden" name="targettype_<%=no%>" value="<%=rswipfinal("targettype")%>">
						<input type="hidden" name="targetcode_<%=no%>" value="<%=rswipfinal("targetcode")%>">
				<%End if%>
		  </tr>
		  <%
		  rswipfinal.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no = no + 1
	  Wend
   else
      response.Write("Data is empty") 	  
   end if 
   rswipfinal.Close
   set rswipfinal = Nothing%>
		<tr>
			<td colspan="12" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="txtagreementcode" value="<%=v_txtagreementcode%>">
			<input type="hidden" name="txtyear" value="<%=v_txtyear%>">
			<input type="hidden" name="tglawal" value="<%=v_txttglawal%>">
			<input type="hidden" name="tglakhir" value="<%=v_txttglakhir%>">
			<input type="hidden" name="txtmonth" value="<%=v_txtmonth%>">
			<input type="hidden" name="wipdate" value="<%=wipdate%>">
			<input type="hidden" name="selectedkontrak" value="<%=selectedkontrak%>">
			<input type="hidden" name="agreementcode" value="<%=agreementcode%>">
			
			
			<input type="hidden" name="No" value="<%=No%>">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<%If SudahClose = "Y" Then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit"  name="Add" value="Update">
			<%End if%>
			
			<%If SudahClose <> "Y" Then%>
				<input type="hidden" name="txtJenis" value="New">
				<input Type="Submit"  name="Add" value="Submit">
			<%End if%>
			<input Type="Reset"  name="Reset" value="Reset">
			<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
	
	</TABLE>
</div>
</form>
<%End if%>
<!---------------akhir!--------------->
</div>
</div>
<!--#include file="bottom.jsp"-->
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	

function getselectedkontrak() {
	var spocode = document.getElementById("selectedkontrak");
	//spocode.value = document.getElementById("agreementcode").value;
	alert(document.getElementById("agreementcode").value);
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
	alert("1");

	getlistkontrak();
	getselectedkontrak();
});
</script>


</body>
</html>
