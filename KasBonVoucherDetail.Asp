<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
    <title><%="KASBON | FB-Plans"%></title>
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


$.urlParam = function(name){
    var results = new RegExp('[\?&]' + name + '=([^&#]*)').exec(window.location.href);
    if (results==null) {
       return null;
    }
    return decodeURI(results[1]) || 0;
}

var reloadCode = function (PostStr) {
	
	var transdate =$( ".datepicker" ).val();
	var txtTable = $.urlParam('txtTable');
	//var vtranstype = $.urlParam('transtype');
	var txtTable = document.getElementById("txtTable").value;
	//var txtBanktype = document.getElementById("banktype").value;
	var fcba =$( "#fcba" ).val();
	
	//alert(vtranstype);
	
	window.location.href = window.location.href.replace( /[\?#].*|$/, "?txtModule=Finance and Accounting&txtTable="+txtTable+"&transdate="+transdate+"&fcba="+fcba+"&"+PostStr );
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
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261015a"></script>
<% end if %>
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

<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %>>  <!-- perbaikan: tag body sebelumnya tidak ada -->
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


Function f_i_check(in_fcba, in_vouchercode, in_datecreated, in_bankcode, in_jobcode, in_targettype, in_targetcode, in_documentno, in_amount)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT validasi_voucher_api.check_data( '"&in_fcba&"', '"&in_vouchercode&"', TO_DATE('"&in_datecreated&"', 'mm/dd/yyyy'), '"&in_bankcode&"', '"&in_jobcode&"', '"&in_targettype&"', '"&in_targetcode&"', '"&in_documentno&"', REPLACE('"&in_amount&"', ',', '') ) result" & _
	                " FROM DUAL"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("result")
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	f_i_check = initialAvailable
End Function


 Dim check
 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,SourceType,fcBA
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks,TxtTType
 Dim rsInvoiceRecord, varTotalAmount, invoiceUpdate, recordcheck, rsrecordlineno, recordWIPLINE,invoiceRecord
 Dim rskasbonvoucherdetail, varamountOriginal, recordinvoicepaid, vartotal, varError, ivp,varmaxlineno,recordlineno
 Dim Jobcode, LocationType, LocationCode, Reference, Remark,InvoiceCode,AGREEMENTCODE,AGREEMENTLINENO ,recordwipcheck
 Dim varSupplierCode, varStringError, varlineno, WIPPAIDRecord, rsWIPPAIDRecord, WIPUpdate ,rsrecordWIPLINE ,W, rsrecordwipcheck, WIPDATE, varContract
 Dim InsertIntoCostActivity, rsqMaxLineNo, qMaxLineNo, MaxLineNo, IsCostbook
 Dim MaxLineNopvd
 Dim rlineno, dataModule, dataSubmodule, dataDetail, qq, reckasbonvoucherdetail
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim Cek_IsAccClosing : Cek_IsAccClosing = ""
 Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD
 Dim xACCOUNTINGYEARPERIOD
 DIM xAmount, StatusXAmount
 Dim sqljnsinv,rsjnsinv,sqlupdatepay,txtnoinvoice,jenisinv,amountinvoice
 Dim is_final
 Dim isNew

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 

 
 if Request("fcba") <> "" then
	fcba = Request("fcba")
 else	
	fcba = session("fcba") 
 end if

 ' if request("fccOde")="" then
 '    fcCode="" '----- Nantinya Ambil dari Format Master
 ' else
 '    fcCode = trim(request("fccOde"))
 ' end if
 

 'fcCode = request("fccOde")
 'If Len(Trim(fcCode)) = 0 Then 
'	fcCode = ""
 'end if
 
 
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
 
 if request("bankcodeurl") <>"" then
   xCode   = Request("bankcodeurl")
   'response.write xCode
 elseif Request("xCode")<>"" then
   xCode = Request("xCode")
 else
   xCode = Request("xCode")
 end if
 
 if Request("isNew") = "" then
 isNew    = "TRUE"
 else
 isNew    = "FALSE" 
 end if
 
 HeaderRemarks = Request("HeaderRemarks")
 
 '---- Data Details
'JobCode=Request("JobCode")
Reference=Request("Reference")
Remarks=Request("Remarks")
Amount=Request("Amount")
rlineno = Request("txtlineno")
documentno = Request("documentno")
document_classification = Request("document_classification")

' Added Variable ChargeJob, ChargeType, ChargeCode
'businessunit=Request("businessunit")
referenceno=Request("referenceno")
remarksdtl=Request("referenceno")
fcCode = Request("fcCode")

' End Added


If Request("Amount") <> ""  then
   xAmount = Replace(Amount, ",", "")
   IF Cdbl(xAmount) > 0 then
      StatusXAmount = "PLUS"
   Else
      StatusXAmount = "MINUS"       
   End If
End If
'response.write "Amount nya " & xAmount


 ' ACCOUNTINGYEARPERIOD = Request("ACCOUNTINGYEARPERIOD")
 ' if "" & ACCOUNTINGYEARPERIOD <> "" then
    ' ACCOUNTINGYEAR = MID(ACCOUNTINGYEARPERIOD,1,4)
	' ACCOUNTINGPERIOD = MID(ACCOUNTINGYEARPERIOD,5,2)
 ' end if
 
 'response.write "ACCOUNTINGYEAR " & ACCOUNTINGYEAR  & "ACCOUNTINGPERIOD " & ACCOUNTINGPERIOD 

  
dim tt

Sub DeleteDatakasbonvoucherdetail
    q = " DELETE FROM kasbonvoucherdetail" & _
		" WHERE fcBA = '" & fcba & "'" & _
		"       AND vouchercode = '" & fcCode & "'" & _
		"       AND lineno = '" & rlineno & "'"
	connStr.execute(q)
	
	Call UpdateAmountkasbonvoucherHeader
End Sub

Sub UpdateAmountkasbonvoucherHeader
' Untuk Update Total Amount Di Header
				q = " SELECT NVL(SUM(amount), 0) totalamount" & _
					" FROM kasbonvoucherdetail" & _
					" WHERE VoucherCode='" & fcCode & "'" & _
					"       AND fcBA = '" & fcba & "'"
				set rsVoucher = server.CreateObject("Adodb.recordset")
				rsVoucher.Open q, connstr
				if Not rsVoucher.Eof then
					TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
				End if

				q =	" Update kasbonvoucher" & _
					" set totalamount = '"& TOTALAMOUNT &"'" & _
					" where vouchercode = '" & fccode & "'" & _
					" and fcBA = '" & fcba & "'"
	 			connStr.execute(q)
End Sub

Sub InsertDatakasbonvoucherdetail
		qq = " INSERT INTO kasbonvoucherdetail ( vouchercode" & _
		     "                                    , amount" & _
		     "                                    , reference" & _
		     "                                    , remarks" & _
		     "                                    , fcentry" & _
		     "                                    , fcedit" & _
		     "                                    , fcip" & _
		     "                                    , fcba" & _
		     "                                    , lineno" & _
			 "                                    ) " & _
		     " VALUES ( '" & fcCode & "'" & _
	         "          , REPLACE('" & Amount & "', ',', '')" & _
	         "          , '" & Reference & "'" & _ 
	         "          , '" & Remarks & "'" & _
	         "          , '" & Session("CodeName") & "'" & _
	         "          , '" & Session("CodeName") & "'" & _
	         "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
	         "          , '" & fcba & "'" & _
	         "          , '" & rlineno & "'" & _
	         "			)"
		' response.write qq
		connStr.execute(qq)
			' Untuk Update Total Amount Di Header
			Call UpdateAmountkasbonvoucherHeader
			' End Added
End Sub

Sub UpdateDatakasbonvoucherdetail
	
	connstr.BeginTrans
		Call DeleteDatakasbonvoucherdetail
		Call InsertDatakasbonvoucherdetail
	
		If Err.Number = 0 Then
			connstr.CommitTrans
			'Response.write("<font face=Verdana size=1 color=Green>Normalisasi employee_transit_tab, Success! " & " Time: " & time() & " ..<br/><br/></font>")
			'Response.Flush
		Else
			connstr.RollbackTrans
			err.number = 1001
			err.description = "Error Save ["&ExecuteSQL&"] ("&Err.Description&")"
			'Response.write("Error Save ["&ExecuteSQL&"] ("&Err.Description&")")
			'Response.Flush
		End If
End Sub

'response.write "A|"	& Request("txtJenis")
	
Cek_IsAccClosing = StatusAccClosingTgl(DATECREATED,fcba)
'response.write "Final =|" & Cek_IsAccClosing & "|" & fcba & "|" & DATECREATED

if 	Cek_IsAccClosing = "FALSE" then 	
'--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "StartAuthorize" or Request("txtJenis")= "Approved" or Request("txtJenis")= "Posting" Then	
			if Request("txtJenis")= "Header" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then

				strAvailable =	" select count(1) total from kasbonvoucherdetail where  VoucherCode = '" & fcCode & "' and fcBA = '" & fcba & "'"
				Set rsAvailable = Server.CreateObject("ADODB.Recordset")
				rsAvailable.Open strAvailable, connstr, 3
				IF Not rsAvailable.EOF Then
					rowAvailable = Cdbl(rsAvailable("Total"))
				Else
					rowAvailable = 0
				End If
				strAvailable =	""
				rsAvailable.Close
				set rsAvailable = Nothing
				
				IF rowAvailable = 0 Then
				'	q = " Delete from kasbonvoucherdetail where  VoucherCode = '" & fcCode & "' and fcBA = '" & fcba & "'"
				'	connStr.execute(q)
					
					q = " Delete from kasbonvoucher where  VoucherCode = '" & fcCode & "' and fcBA = '" & fcba & "'"
					connStr.execute(q)
				Else
					err.number = 1001
					err.description = "Silahkan delete detail voucher dahulu, terima kasih "
				END If
					
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If  
				
				cmdView	= "True"
			   
			elseif Request("txtJenis")= "Details" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then

					call DeleteDatakasbonvoucherdetail
					
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If 
				
				cmdView	= "True"
			
			elseif Request("txtJenis")="New" then
				Dim rsPO
				 'response.write "datanya " & Request("ACCOUNTINGYEARPERIOD") 
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then
					
					'--Add by Ukie to handle Auto Number of PR
							If Ucase(fcCode) = Ucase("(NEW)") Then
								oSql = " select nvl(controlnumber_api.get_nextnumber(upper('KASBONVOUCHER')" & _
									   "                                         , " & year(textDate(datecreated)) & "" & _
									   "                                         , " & month(textDate(datecreated)) & "" & _
									   "                                         , '" & fcba & "'" & _
									   " ), 'FALSE') transtype from dual"
								
								' Response.Write oSql
									   
								Set oRs = Server.CreateObject("ADODB.Recordset")
								oRs.Open oSql, connstr
								If Not oRs.EOF Then
									txtvouchercode = oRs("transtype")
								End If
								oRs.Close
								Set oRs = Nothing
								'response.write "X|" & txtvouchercode  ' dinonaktifkan: debug tampil di halaman
									fcCode = txtvouchercode
							End If
							
							If txtvouchercode = "FALSE" Then
								'err.Number = 1001
								'err.description = "<br/>Automatic Number Not Have Been Set, Call your Administrator!"
							End If
							
							' AMH cek if voucher number already exists 
							oSql = " SELECT 1 FROM kasbonvoucher where VoucherCode='" & fccode & "' and fcba='" & fcba & "'"
							Set oRs = Server.CreateObject("ADODB.Recordset")
								oRs.Open oSql, connstr
								If Not oRs.EOF Then
									txtvouchercode = "TRUE"
								else
									txtvouchercode = "FALSE"
								End If
								oRs.Close
								Set oRs = Nothing
						'Response.write "1-" & txtvouchercode & "|" & trim(isNew) & "<br>"
						
							if txtvouchercode = "TRUE" and trim(isNew) = "TRUE" then
								q = "select Decode( Nvl(Max(cn.PREFIX),'-'), '-', '" & fcba & "'||'/'|| 'KASBON/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'YYMM') ||'/'|| LPAD(NVL(MAX(regexp_substr(det.VOUCHERCODE, '[^/]+', 1, 4)),0000)+0001,4,0), " & _
										    " MAX(cn.PREFIX) || 'KASBON/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'YYMM') ||'/'|| LPAD(NVL(MAX(regexp_substr(det.VOUCHERCODE, '[^/]+', 1, 4)),0000)+0001,4,0)) DOCNUMBER " & _
											" from " & _
											" controlnumber cn " & _
											" left join kasbonvoucher det on cn.fcba = det.fcba and cn.accountingyear = to_number(to_char(det.DATECREATED,'yyyy')) " & _
											" where cn.fcba ='" & fcba & "' " & _
											" and cn.LU_NAME ='KASBONVOUCHER' "&_
											" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yyyy')) " & _
											" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'mm')) " & _
											" and LPAD(det.vouchercode,Length(rtrim(det.vouchercode))-4) = cn.PREFIX || 'KASBON/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'YYMM') ||'/' "
										'response.write q  ' dinonaktifkan: dulu mencetak SQL mentah ke halaman
										set rsQuery = server.CreateObject("Adodb.recordset")
										rsQuery.Open q, connstr
										if Not rsQuery.Eof then
												fcCode = rsQuery("DOCNUMBER")
										end if
										rsQuery.Close
										Set rsQuery = Nothing
							isNew = "FALSE"
							End if
						
						'Response.write "2-" & txtvouchercode & "|" & trim(isNew) & "<br>"
						
						'Response.Write PRCODE & txtPRCode
						check_data = "Success"
						
					'--End Added
					if check_data = "Success" then
					
						Set rsPO = Server.CreateObject("ADODB.Recordset")
						rsPO.Open " SELECT vouchercode FROM kasbonvoucher WHERE vouchercode = '" & fcCode & "' AND fcba = '" & fcba & "'", connstr
						IF rsPO.eof then
							
									
									q	=	" INSERT INTO kasbonvoucher ( vouchercode" & _
									"                              , referenceno" & _
									"                              , datecreated" & _
									"                              , document_classification" & _
									"                              , accountingyear" & _
									"                              , accountingperiod" & _
									"                              , totalamount" & _
									"                              , remarks" & _
									"							   , rowstateid	" & _
									"                              , fcentry" & _
									"                              , fcedit" & _
									"                              , fcip" & _
									"                              , fcba" & _
									"                              , lastupdate" & _
									"                              , lasttime)" & _
									" VALUES ( '" & fcCode & "'" & _
									"          , '" & referenceno & "'" & _
									"          , TO_DATE('" & DATECREATED & "','mm/dd/yyyy')" & _
									"          , '" & document_classification & "'" & _
									"          ,TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED & "','mm/dd/yyyy'),'yyyy'))" &_
									"          ,TO_NUMBER(TO_CHAR(TO_DATE('" & DATECREATED & "','mm/dd/yyyy'),'mm'))" &_
									"          , 0" & _
									"          , '" &  request("HeaderRemarks") & "'" & _
									"          , 'Planned'" & _
									"          , '" & Session("CodeName") & "'" & _
									"          , '" & Session("CodeName") & "'" & _
									"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									"          , '" & fcba & "'" & _
									"          , SYSDATE" & _
									"          , TO_CHAR(SYSDATE,'HH24:MI') )"
									
									'response.Write("q " & q )
									connStr.execute(q)
							 
						End if
						rsPO.Close
						set rsPO = Nothing

							
								qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
												" from kasbonvoucherdetail " & _
												" where " & _
												" fcba = '" & fcba & "' " & _
												" and VoucherCode = '" & fcCode & "'"
								'response.write 	qMaxLineNo			
								Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
								rsqMaxLineNo.CursorType = 3
								rsqMaxLineNo.CursorLocation = 3
								rsqMaxLineNo.Open qMaxLineNo, connstr

								if not rsqMaxLineNo.eof then
									maxLineNopvd = cdbl(rsqMaxLineNo("maxLineNo"))
								else
									maxLineNopvd = 0
								end if
								rsqMaxLineNo.close
								set rsqMaxLineNo = Nothing
								
								maxLineNopvd = maxLineNopvd + 1
								
								rlineno = maxLineNopvd
						
								Call InsertDatakasbonvoucherdetail

					End If
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
				cmdView	= "True"
			elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			    'response.write("masuka")
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				'response.write("masukb" & Cek_IsDailyClosing)
				If Cek_IsDailyClosing = "FALSE" then
					
					Call UpdateDatakasbonvoucherdetail
					cmdView	= "True"
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If 
			elseif Request("txtJenis")= "Confirmed" then
			   connstr.BeginTrans
			   strQueryStatus = 	" Update kasbonvoucher" & _
									" set rowstateid = 'Approved'" & _
									"     , fcedit = '" & Session("CodeName") & "'" & _
									" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									"     , lastupdate = sysdate " & _
									"     , lasttime = to_char(sysdate,'hh24:mi') " & _
									" 		where vouchercode = '" & fccode & "'" & _
									"       and fcBA = '" & fcba & "'"
			   connStr.execute(strQueryStatus)

			   if err.number  = 0 then
					  connStr.CommitTrans
			   end if       
			   Call UpdateAmountkasbonvoucherHeader
			   
		       cmdView	= "True"
			elseif Request("txtJenis")= "RePlanned" then
			   connstr.BeginTrans   
			   strQueryStatus = 	" Update kasbonvoucher" & _
									" set rowstateid = 'Planned'" & _
									"     , fcedit = '" & Session("CodeName") & "'" & _
									" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									"     , lastupdate = sysdate " & _
									"     , lasttime = to_char(sysdate,'hh24:mi') " & _
									" 		where vouchercode = '" & fccode & "'" & _
									"       and fcBA = '" & fcba & "'"
			   connStr.execute(strQueryStatus)
			  
			   if err.number  = 0 then
					  connStr.CommitTrans
			   end if      	  
               
			   Call UpdateAmountkasbonvoucherHeader
		       cmdView	= "True"
			else
		        cmdView	= "True"
			end if								 	
  else
    cmdView	= request("cmdView")
 End if
else
    err.number = 1001
	err.description = "Accounting Period untuk tanggal " & DATECREATED & " (BA " & fcba & ") sudah LOCKED !!!"
	cmdView	= request("cmdView")
End if

Dim rsVoucher
' Remark by iwa 
'q = " SELECT bankcode" & _
'	"        , datecreated" & _
'	"        , document_classification" & _
'	"        , totalamount" & _
'	"        , remarks" & _
'	"        , rowstateid" & _
'	"        , accountingperiod" & _
'	"        , accountingyear" & _
'	"        , accountingyear||LPAD(accountingperiod,2,0) accountingyearperiod" & _
'	" FROM kasbonvoucher" & _
'	" WHERE VoucherCode like '%" & fcCode & "%'" & _
'	"       AND fcBA = '" & fcba & "'"
' End Remark by iwa 	

q = " SELECT " & _
	"          to_char(datecreated,'mm/dd/yyyy') datecreated" & _
	"        , document_classification" & _
	"        , totalamount" & _
	"        , referenceno" & _ 
	"        , remarks" & _
	"        , rowstateid" & _
	"        , accountingperiod" & _
	"        , accountingyear" & _
	"        , accountingyear||LPAD(accountingperiod,2,0) accountingyearperiod" & _
	" FROM kasbonvoucher" & _
	" WHERE VoucherCode = '" & fcCode & "'" & _
	"       AND fcBA = '" & fcba & "'"
'response.write q
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
	DATECREATED=rsVoucher("DateCreated")
	TOTALAMOUNT = rsVoucher("TOTALAMOUNT")
	HeaderRemarks=rsVoucher("Remarks")	
	document_classification=rsVoucher("document_classification")
	rowstateid=rsVoucher("rowstateid")	
	ACCOUNTINGYEAR =rsVoucher("ACCOUNTINGYEAR")
	ACCOUNTINGPERIOD =rsVoucher("ACCOUNTINGPERIOD")
	referenceno = rsVoucher("referenceno")
	isNew = "FALSE"
else	
	isNew = "TRUE"
End if

q = " SELECT NVL(SUM(amount), 0) totalamount" & _
	" FROM kasbonvoucherdetail" & _
	" WHERE VoucherCode='" & fcCode & "'" & _
	"       AND fcBA = '" & fcba & "'"
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
	TOTALAMOUNT = cdbl(to_nol(rsVoucher("TOTALAMOUNT")))
End if
 
if Request("txtJenis") = "Edit" then
	q = " SELECT vouchercode" & _
		"        , amount" & _
		"        , reference" & _
		"        , remarks" & _
		"        , lineno" & _
		" FROM kasbonvoucherdetail" & _
		" WHERE fcba = '" & fcba & "'" & _
		"       AND vouchercode like '%" & fcCode & "%'" & _
		"       AND lineno = '" & rlineno & "'" 

	set rsVoucher = server.CreateObject("Adodb.recordset")
	rsVoucher.Open q, connstr
	
	if Not rsVoucher.Eof then
		Amount = rsVoucher("amount")
		Reference = rsVoucher("reference")
		Remarks = rsVoucher("remarks")
		rlineno = rsVoucher("lineno")
	End if
	
	
		documentno = ""
	
end if


 
rsVoucher.Close
set rsVoucher = Nothing

if err.number <> 0 then
   
   'connStr.RollbackTrans
   response.Write  "<div class='error'>" & err.description & "</div>"
   cmdView	= "True"
End if

Dim val_trigger_system : val_trigger_system = "Bank"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select nvl(trigger_system, 'Bank') trigger_system" & _
            " from bank" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and fccode = '" & xCode & "'"
	   
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_trigger_system = oRecordSet("trigger_system")
Else
	val_trigger_system = "Bank"
End if
'response.Write strNurseryCulling
oRecordSet.Close
Set oRecordSet = Nothing

' Untuk Cek Start Authorize By Sendy 04/29/2019
set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT NVL(SUM(amount),0) total" & _
	" FROM kasbonvoucherdetail" & _
	" WHERE vouchercode = '" & fcCode & "'" & _
	"       AND fcba = '" & fcba & "'" 
' response.write q
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	pv_totalamount = cdbl(to_nol(rsFrez("total")))
else
	pv_totalamount = 0
End if
rsFrez.Close : set rsFrez = Nothing 
 
set rsFrez = server.CreateObject("Adodb.recordset")
q = " SELECT COUNT(1) total" & _
	" FROM authorizelevel al" & _
	" INNER JOIN authorizestructure aus ON aus.authorizelevelcode = al.fccode" & _
	" INNER JOIN authorizetemplate at ON aus.authorizetemplatecode = at.fccode" & _
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'PV'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.startlevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND case when "&pv_totalamount&" < 0 then 1 else "&pv_totalamount&" end BETWEEN at.amountrangefrom and at.amountrangeuntil" & _
	"       AND '"& session("CodeName") &"' = '" & session("CodeName") & "'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & fcba & "'"
' response.write q
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
	" INNER JOIN authorizelevel_users alu ON (alu.authorizelevelcode = al.fccode)" & _
	" WHERE at.document_type = 'PV'" & _
	"       AND al.isactive = 'TRUE'" & _
	"       AND al.finallevel = 'TRUE'" & _
	"		AND alu.isactive = 'TRUE'" & _
	"       AND case when "&pv_totalamount&" < 0 then 1 else "&pv_totalamount&" end BETWEEN at.amountrangefrom and at.amountrangeuntil" & _
	"       AND '"& session("CodeName") &"' = '" & session("CodeName") & "'" & _
	"       AND aus.fcusercode = '" & session("CodeName") & "'" & _
	"       AND aus.fcba = '" & fcba & "'"
rsFrez.Open q, connstr
if Not rsFrez.Eof then
	is_final =  cdbl(to_nol(rsFrez("total")))
else
	is_final = 0
End if
rsFrez.Close : set rsFrez = Nothing 

' End Added

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 2))" & _
	        "       and upper(system) = upper('" & val_trigger_system & "')"
			
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

 '-------------ISI Payment Form
%>
                   
                    <form name="frmEntry" method="post" action="kasbonvoucherdetail.Asp" >
					
                    <font color="red">
                        <%= varStringError  %></font>
                    <table class="data" width="100%">
						<tr>
                            <td align="left">
                                Date
                            </td>
                            <td align="left">
                                <input class="datepicker" type="text" maxlength="10" onchange="javascript:reloadCode('#');" name="DATECREATED" size="10" value="<%=DATECREATED%>">
                                <!--<a href="javascript:show_calendar('frmEntry.DATECREATED');" onMouseOver="window.status='Date Picker';return true;"
                                    onmouseout="window.status='';return true;">
                                    <img src="Images/show-calendar.gif" width="24" height="22" border="0" align="top"></a>-->
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                        </tr>							
											
                        <tr>
                            <td align="left">
                                Voucher No
                            </td>
                            <td align="left">
							<%
									
								if isNew = "TRUE" then
									'response.write xCode
								' MODIFY AMH - Apr 21 penomoran sesuai bank
										q = "select Decode( Nvl(Max(cn.PREFIX),'-'), '-', '" & fcba & "'||'/'|| 'KASBON/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'YYMM') ||'/'|| LPAD(NVL(MAX(regexp_substr(det.VOUCHERCODE, '[^/]+', 1, 4)),0000)+0001,4,0), " & _
										    " MAX(cn.PREFIX) || 'KASBON/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'YYMM') ||'/'|| LPAD(NVL(MAX(regexp_substr(det.VOUCHERCODE, '[^/]+', 1, 4)),0000)+0001,4,0)) DOCNUMBER " & _
											" from " & _
											" controlnumber cn " & _
											" left join kasbonvoucher det on cn.fcba = det.fcba and cn.accountingyear = to_number(to_char(det.DATECREATED,'yyyy')) " & _
											" where cn.fcba ='" & fcba & "' " & _
											" and cn.LU_NAME ='KASBONVOUCHER' "&_
											" and cn.accountingyear =  TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'yyyy')) " & _
											" and cn.accountingperiod = TO_NUMBER(TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'mm')) " & _
											" and LPAD(det.vouchercode,Length(rtrim(det.vouchercode))-4) = cn.PREFIX || 'KASBON/' || TO_CHAR(TO_DATE('"& DATECREATED &"','mm/dd/yyyy'),'YYMM') ||'/' "
										'response.write q 
										set rsQuery = server.CreateObject("Adodb.recordset")
										rsQuery.Open q, connstr
										if Not rsQuery.Eof then
												fcCode = rsQuery("DOCNUMBER")
											else
												fcCode = "(NEW)"
										End if										
									End if
									
							%>
                                <input type="text" maxlength="30" name="fcCode" size="30" value="<%=fcCode%>" title="" Required>
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                        </tr>
                       
                        <tr>
							<td align="left">Classification</td>
								<td align="left">
									<select name="document_classification" class="chosen-select"  >
									<option></option>
									<%
									qLov = " SELECT fccode, fcname FROM document_classification where document_type = 'PV' ORDER BY fccode"
						   
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
                                <td align="left">
                                    Amount Total
                                </td>
                                <td align="left">
                                    <input type="text" maxlength="50" name="TOTALAMOUNT" size="30" value="<%=TOTALAMOUNT%>"
                                        readonly="" style="background:#FFFFCC;">
                                </td>
                            </tr>
                            <tr>
                                <td align="left">
                                   Requestor
                                </td>
                                <td align="left">
								    <input type="text" maxlength="200" name="HeaderRemarks" size="30" value="<%=HeaderRemarks%>" >                                 
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
                                Remarks
                            </td>
                            <td align="left">
							    <textarea id="referenceno" name="referenceno"  cols="50" rows="3"><%=referenceno%></textarea>
                            </td>
                    </tr>
					<tr>
						<td colspan="2" align="left">
							<input type="Hidden" value="True" name="cmdView">
							<input type="Hidden" value="<%=request("txtTable")%>" name="txtTable" id="txtTable">
							<input type="Hidden" value="<%=request("txtModule")%>" name="txtModule">
							<input type="Submit" name="Add" value="Select">&nbsp;&nbsp;							
							<!--<a href="javascript:launchwin('PrintVoucher.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&transtype=<%=transtype%>&banktype=<%=banktype%>&transdate=<%=request("datepicker")%>&bankcodeurl=<%=request("bankcode")%>&DATECREATED=<%=DATECREATED%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
								<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana"
									size="1">Print</font></a>&nbsp;&nbsp;-->							
							<%
							'if  (RowStateId = "Confirmed" or RowStateId = "Planned")  then 
							txtFileReportName = "PrintKasbon"
							'else 
							'txtFileReportName = "printvoucherfinapprove"
							'end if
							
							%>
							<% if RowStateId <> "Planned"  then %>
							<a href="<%=txtFileReportName%>.asp?fcba=<%=fcba%>&fccode=<%=fcCode%>" target="_blank">
								<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
							</a>&nbsp;&nbsp;			
							<%End If%>
							
							
							<% if RowStateId = "Planned" then %>
							     <!-- Modify AMH Mar 21 validasi balance -->
								 <%if Request("txttable") = "03-journalvoucher" and TOTALAMOUNT  <> 0 then %>
								 <font color="#FF0000">Pending Confirmation,please check your balance</font>
								 <%else%>
								 <a href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&document_classification=<%=document_classification%>&fcba=<%=fcba%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to confirm this data :  <%=fcCode%> ?');"> Confirmed </a>  
								 <% end if%>
								 
							<% Elseif RowStateId = "Confirmed" then %> 
									<a href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&document_classification=<%=document_classification%>&fcba=<%=fcba%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to replan this data :  <%=fcCode%> ?');"> Re-Planned </a>
									&nbsp;&nbsp;
										<%if can_startauthorize = 0 then%>
											<font color="#FF0000">You Have No Permission to Start Authorize</font>
										<%else%>
											<a href="?fcCode=<%=fcCode%>&xcode=<%=xcode%>&DATECREATED=<%=DATECREATED%>&transtype=<%=transtype%>&banktype=<%=banktype%>&document_classification=<%=document_classification%>&fcba=<%=fcba%>&txtJenis=StartAuthorize" onClick="return confirm('Are you sure to start authorize this data :  <%=fcCode%> ?');">Start Authorize</a>
										<% end if%>
							<% end if%>
							
							
							
						</td>
					</tr>
			</table>
        </form>

<%if cmdView = "True" and "" & fcCode <> ""  and "" & document_classification <> "" then 
	Dim rskasbonvoucher, intRecordsPerPage, cColor
    isNew = "FALSE"
	q = " select pv.vouchercode" & _
		"        , pv.referenceno" & _
		"        , pvd.lineno" & _
	    "        , pvd.amount" & _
	    "        , pvd.reference" & _
	    "        , pvd.remarks" & _
	    " from kasbonvoucher pv" & _
	    " inner join kasbonvoucherdetail pvd ON (pv.VoucherCode = pvd.vouchercode and pv.fcba = pvd.fcba)" & _
		" where pv.vouchercode = '" & fcCode & "'" & _
		"       and pv.fcba = '" & fcba & "'" & _
		" order by pv.fcba, pv.vouchercode, pvd.lineno"
	'response.Write q	
	Set rskasbonvoucher = Server.CreateObject("ADODB.Recordset")
	rskasbonvoucher.CursorType = 3
	rskasbonvoucher.CursorLocation = 3
	rskasbonvoucher.Open q, connstr
%>
        <form action="kasbonvoucherdetail.Asp?isNew=FALSE" method="post" name="eForm" style="display:table;">
								
        <table class="data" cellpadding="0" cellspacing="0" width="100%" border="0">
            
			<tr style="display: table-row;">
             
            <tr style="display: table-row;">
                <td align="left">
                    Amount
                </td>
                <td align="left">
				<% if RowStateId = "Planned" or RowStateId = "" then %>
                    <input type="text" size="20" name="Amount" value="<%=Amount%>"  onkeyup="hanyaAngka(this);" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">&nbsp;
				<% Else %>
                    <input type="text" size="20" name="Amount" value="<%=Amount%>"  readonly="" style="background:#FFFFCC;" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">&nbsp;
				<% End If %>
			    </td>
            </tr>
           
            <tr style="display: table-row;">
                <td align="left">
                    Reference
                </td>
                <td align="left">
                    <textarea name="Reference" cols="50"><%=Reference%></textarea>
                </td>
            </tr>
            <tr style="display: table-row;">
                <td align="left">
                    Particular
                </td>
                <td align="left">
                    <textarea name="Remarks" cols="50" rows="3"><%=Remarks%></textarea>
                </td>
            </tr>
            </td></tr>
            <tr>
                <td colspan="2" align="left">
                    <!-- diambil dari form yang diatas -->
					<input type="hidden" size="11" name="txtlineno" value="<%=rlineno%>">
                    <input type="hidden" name="fcCode" value="<%=fcCode%>">
                    <input type="hidden" name="referenceno" value="<%=referenceno%>">
					<input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
					<input type="hidden" name="fcba" id="fcba" value="<%=fcba%>">
                    <input type="hidden" name="xCode" value="<%=xCode%>">
                    <input type="hidden" name="HeaderRemarks" value="<%=HeaderRemarks%>">
					<input type="hidden" name="rowstateid" value="<%=rowstateid%>">
                    <input type="hidden" name="document_classification" value="<%=document_classification%>">
					
					<!-- Period Control In Here -->
						<%IF Request("txtJenis")="Edit" then%>
							<input type="hidden" name="txtJenis" value="Edit">
							<input Type="Submit" name="Add" value="Update">
						<%Else%>
							<% if RowStateId = "" OR RowStateId = "Planned" then %>
							   <input type="hidden" name="txtJenis" value="New">
							   <input type="hidden" name="isNew" id="isNew" value="<%=isNew%>">
							   <input Type="Submit" name="Add" value="Add">
							<%Else%>
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
                    CODE
                </th>
                <th align="left">
                     REMARK KAS BANK
                </th>
             	<th align="left">
                    LINE NO
                </th>
                <th align="left">
                    AMOUNT
                </th>
                <th align="left">
                    REFERENCES
                </th>
                <th align="left">
                    REMARKS
                </th>
            </tr>
            </thead>
            <tbody>
            <%
   if Not rskasbonvoucher.eof then
      nCountofField=rskasbonvoucher.fields.count-1
	  no=1
	  While Not rskasbonvoucher.EOF
          isNew = "FALSE"
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
							<a href="kasbonvoucherdetail.Asp?fcCode=<%=rskasbonvoucher("VoucherCode")%>&txtlineno=<%=rskasbonvoucher("lineno")%>&fcba=<%=fcba%>&txtJenis=Edit">
								<img src="IMAGES/edit.gif" alt="Edit Data : <%=rskasbonvoucher(1)%>" border="0">
							</a>
						</td>
						<td>
							<a href="kasbonvoucherdetail.Asp?fcCode=<%=rskasbonvoucher("VoucherCode")%>&txtlineno=<%=rskasbonvoucher("lineno")%>&fcba=<%=fcba%>&txtJenis=Details">
								<img src="IMAGES/delete.gif" alt="Delete Data : <%=rskasbonvoucher(1)%>" border="0" onclick="return confirm('Are you sure to delete');">
							</a>
						</td>
					<%Else%>
						<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
					<%End if%>
               <!-- End Period Control In Here -->
                <%for i = 0 to nCountofField%>
					<td align="center" title="<%=rskasbonvoucher(i).Name%>" valign="top">
						<%=rskasbonvoucher(i)%>
					</td>
                <%next%>
                
            </tr>
            <%
		  rskasbonvoucher.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rskasbonvoucher.Close
   set rskasbonvoucher = Nothing%>
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
</body>
</html>