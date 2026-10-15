<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
    <title>Journal Reclass - iPlaS</title>
	<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-3.1.0.min.js"></script>
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
<%if request("Add")="" then%>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %> onLoad="document.frmEntry.fcCode.focus()">
    <%Else%>
    <body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %> onload="document.eForm.JobCode.focus()">
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
 
Function JobcodeWhitelist(in_fcba, in_vouchercode, in_jobcode_db, in_jobcode_cr)
	strAvailable = " SELECT reclassjournal_api.check_data( '"&in_fcba&"', '"&in_vouchercode&"', '"&in_jobcode_db&"', '"&in_jobcode_cr&"') resultx FROM DUAL"
	'response.write strAvailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("resultx")
	Else
		initialAvailable = "Reclass tidak diizinkan, Silahkan perbaiki kembali jobcode debet dan credit yang di masukkan"
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	JobcodeWhitelist = initialAvailable
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

Cek_IsDailyClosing = ""

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

if request("fccOde")="" then
	fcCode="" '----- Nantinya Ambil dari Format Master
else
	fcCode = trim(request("fccOde"))
end if
 

 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("DATECREATED")<>"" then
   DATECREATED = Request("DATECREATED")
 else
   DATECREATED = textDate(date)
 end if
 xCode = Request("xCode")
 HeaderRemarks = Request("HeaderRemarks")
 
 '---- Data Details
JobCode_db=Request("JobCode_db")
LocationType_db=Request("LocationType_db")
LocationCode_db=Request("LocationCode_db")

JobCode_cr=Request("JobCode_cr")
LocationType_cr=Request("LocationType_cr")
LocationCode_cr=Request("LocationCode_cr")

documentno = Request("documentno")
Amount=Request("Amount")
Reference=Request("Reference")
Remarks=Request("Remarks")

rlineno = Request("txtlineno")
grouplineno = Request("grouplineno")

If "" & Amount <> ""  then
   xAmount = Replace(Amount, ",", "")
   IF Cdbl(xAmount) > 0 then
      StatusXAmount = "PLUS"
   Else
      StatusXAmount = "MINUS"       
   End If
End If
'response.write "Amount nya " & xAmount


 ACCOUNTINGYEARPERIOD = Request("ACCOUNTINGYEARPERIOD")
 if "" & ACCOUNTINGYEARPERIOD = "" then
	ACCOUNTINGYEARPERIOD = Year(DATECREATED) & Lpad(Month(DATECREATED), 0, 2)
 else
    ACCOUNTINGYEAR = MID(ACCOUNTINGYEARPERIOD,1,4)
	ACCOUNTINGPERIOD = MID(ACCOUNTINGYEARPERIOD,5,2)
 end if
 
 'response.write "ACCOUNTINGYEAR " & ACCOUNTINGYEAR  & "ACCOUNTINGPERIOD " & ACCOUNTINGPERIOD 

  
 if Request("SourceType")="" then
    SourceType="BA" 'Default if Entry is Blank
 else
    SourceType=Request("SourceType")
 End if
dim tt

Sub DeleteDataPaymentVoucherDetail
	arr_del_eof = true
	oSQLQuery = " SELECT lineno" & _
				" FROM paymentvoucherdetail" & _
				" WHERE fcba = '" & Session("fcBA") & "'" & _
				"       AND vouchercode = '" & fcCode & "'" & _
				"       AND grouplineno = '" & grouplineno & "'"
	'response.write "<br/>" & oSQLQuery
	set oRecordSet = server.CreateObject("Adodb.recordset")
	oRecordSet.Open oSQLQuery, connstr
	If Not oRecordSet.Eof then
		arr_del = oRecordSet.getrows()
		arr_del_eof = false
	End if
	oRecordSet.Close
	Set oRecordSet = Nothing
	
	if not arr_del_eof then
		for yy=0 to ubound(arr_del,2)
		
			q = " DELETE FROM paymentvoucherdetail" & _
				" WHERE fcBA = '" & Session("fcBA") & "'" & _
				"       AND vouchercode = '" & fcCode & "'" & _
				"       AND lineno = '" & arr_del(0,yy) & "'"
			connStr.execute(q)
		
		next
	end if
End Sub

Sub InsertDataPaymentVoucherDetail
	qq = " INSERT INTO paymentvoucherdetail ( vouchercode" & _
	     "                                    , sourcetype" & _
	     "                                    , locationtype" & _
	     "                                    , locationcode" & _
	     "                                    , jobcode" & _
	     "                                    , amount" & _
	     "                                    , reference" & _
	     "                                    , remarks" & _
	     "                                    , fcentry" & _
	     "                                    , fcedit" & _
	     "                                    , fcip" & _
	     "                                    , fcba" & _
	     "                                    , invoicecode" & _
	     "                                    , agreementcode" & _
	     "                                    , agreementlineno" & _
	     "                                    , wipdate" & _
	     "                                    , lineno" & _
	     "                                    , lastupdate" & _
	     "                                    , lasttime" & _
		 "                                    , grouplineno) " & _
	     " VALUES ( '" & fcCode & "'" & _
         "          , 'BA'" & _
         "          , '" & LocationType_db & "'" & _ 
         "          , '" & LocationCode_db &  "'" & _
         "          , '" & Jobcode_db & "'" & _
         "          , REPLACE('" & Amount & "', ',', '')" & _
         "          , '" & Reference & "'" & _ 
         "          , '" & Remarks & "'" & _
         "          , '" & Session("CodeName") & "'" & _
         "          , '" & Session("CodeName") & "'" & _
         "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
         "          , '" & Session("fcBA") & "'" & _
         "          , NULL" & _
         "          , NULL" & _
         "          , NULL" & _
         "          , NULL" & _
         "          , '" & rlineno & "'" & _
         "          , SYSDATE" & _
         "          , TO_CHAR(SYSDATE,'HH24:MI')" & _
         "          , '" & grouplineno & "' )"
	 
	connStr.execute(qq)
	
	rlineno = rlineno +1
	qq = " INSERT INTO paymentvoucherdetail ( vouchercode" & _
	     "                                    , sourcetype" & _
	     "                                    , locationtype" & _
	     "                                    , locationcode" & _
	     "                                    , jobcode" & _
	     "                                    , amount" & _
	     "                                    , reference" & _
	     "                                    , remarks" & _
	     "                                    , fcentry" & _
	     "                                    , fcedit" & _
	     "                                    , fcip" & _
	     "                                    , fcba" & _
	     "                                    , invoicecode" & _
	     "                                    , agreementcode" & _
	     "                                    , agreementlineno" & _
	     "                                    , wipdate" & _
	     "                                    , lineno" & _
	     "                                    , lastupdate" & _
	     "                                    , lasttime" & _
		 "                                    , grouplineno) " & _
	     " VALUES ( '" & fcCode & "'" & _
         "          , 'BA'" & _
         "          , '" & LocationType_cr & "'" & _ 
         "          , '" & LocationCode_cr &  "'" & _
         "          , '" & Jobcode_cr & "'" & _
         "          , REPLACE('" & Amount & "', ',', '') *-1" & _
         "          , '" & Reference & "'" & _ 
         "          , '" & Remarks & "'" & _
         "          , '" & Session("CodeName") & "'" & _
         "          , '" & Session("CodeName") & "'" & _
         "          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _ 
         "          , '" & Session("fcBA") & "'" & _
         "          , NULL" & _
         "          , NULL" & _
         "          , NULL" & _
         "          , NULL" & _
         "          , '" & rlineno & "'" & _
         "          , SYSDATE" & _
         "          , TO_CHAR(SYSDATE,'HH24:MI')" & _
         "          , '" & grouplineno & "' )"
	 
	connStr.execute(qq)
End Sub

Sub UpdateDataPaymentVoucherDetail
	Call DeleteDataPaymentVoucherDetail
	Call InsertDataPaymentVoucherDetail
End Sub
		
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then		
			if Request("txtJenis")= "Header" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then

					q = " Delete from PaymentVoucherDetail where  VoucherCode = '" & fcCode & "' and fcBA = '" & Session("fcBA") & "'"
					connStr.execute(q)
					
					q = " Delete from PaymentVoucher where  VoucherCode = '" & fcCode & "' and fcBA = '" & Session("fcBA") & "'"
					connStr.execute(q)
					
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If  
				
				cmdView	= "True"
			   
			elseif Request("txtJenis")= "Details" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then

					call DeleteDataPaymentVoucherDetail
					
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
				
					Cek_JobcodeWhitelist = JobcodeWhitelist(Session("FCBA"), fcCode, JobCode_db, JobCode_cr)
					
					If "" & Cek_JobcodeWhitelist = "Success" THEN
						' Perbaikan: header + baris debet + baris kredit dalam satu transaksi agar jurnal tidak timpang
						connStr.BeginTrans

						Set rsPO = Server.CreateObject("ADODB.Recordset")
						rsPO.Open " SELECT vouchercode FROM paymentvoucher WHERE vouchercode = '" & fcCode & "' AND fcba = '" & Session("FCBA") & "'", connstr
						IF rsPO.eof then	    
							q	=	" INSERT INTO paymentvoucher ( vouchercode" & _
									"                              , bankcode" & _
									"                              , datecreated" & _
									"                              , accountingyear" & _
									"                              , accountingperiod" & _
									"                              , totalamount" & _
									"                              , suppcontcode" & _
									"                              , vouchertype" & _
									"                              , remarks" & _
									"							   , rowstateid	" & _
									"							   , automaticapproved	" & _
									"                              , fcentry" & _
									"                              , fcedit" & _
									"                              , fcip" & _
									"                              , fcba" & _
									"                              , lastupdate" & _
									"                              , lasttime)" & _
									" VALUES ( '" & fcCode & "'" & _
									"          , '88'" & _
									"          , TO_DATE('" & DATECREATED & "','mm/dd/yyyy')" & _
									"          , " & ACCOUNTINGYEAR & " " &_
									"          , TO_NUMBER(" & ACCOUNTINGPERIOD & ")" &_
									"          , 0" & _
									"          , ''" & _
									"          , ''" & _
									"          , '" &  request("HeaderRemarks") & "'" & _
									"          , 'Approved'" & _
									"          , '1'" & _
									"          , '" & Session("CodeName") & "'" & _
									"          , '" & Session("CodeName") & "'" & _
									"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
									"          , '" & Session("fcBA") & "'" & _
									"          , SYSDATE" & _
									"          , TO_CHAR(SYSDATE,'HH24:MI') )"
									'response.Write("q " & q )
									connStr.execute(q)
							 
						End if
						rsPO.Close
						set rsPO = Nothing
						
						grouplineno = 0
						qMaxLineNo = " SELECT NVL(MAX(grouplineno),0) maxLineNo " & _
									 " FROM paymentvoucherdetail " & _
									 " WHERE fcba = '" & Session("fcBA") & "' "
										
						Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
						rsqMaxLineNo.CursorType = 3
						rsqMaxLineNo.CursorLocation = 3
						rsqMaxLineNo.Open qMaxLineNo, connstr

						if not rsqMaxLineNo.eof then
							grouplineno = cdbl(rsqMaxLineNo("maxLineNo"))
						else
							grouplineno = 0
						end if
						rsqMaxLineNo.close
						set rsqMaxLineNo = Nothing
						
						grouplineno = grouplineno +1
						
						qMaxLineNo = " SELECT NVL(MAX(lineno),0) maxLineNo " & _
									 " FROM paymentvoucherdetail " & _
									 " WHERE fcba = '" & Session("fcBA") & "' "
										
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
				
						Call InsertDataPaymentVoucherDetail
						connStr.CommitTrans
					else
						err.number = 1001
						err.description = Cek_JobcodeWhitelist
					end if
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
				cmdView	= "True"
			elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then
					
					' Perbaikan: hapus + insert ulang pasangan debet/kredit dalam satu transaksi agar tidak hilang bila gagal
					connStr.BeginTrans
					Call DeleteDataPaymentVoucherDetail
					
					qMaxLineNo = " SELECT NVL(MAX(lineno),0) maxLineNo " & _
								 " FROM paymentvoucherdetail " & _
								 " WHERE fcba = '" & Session("fcBA") & "' "
									
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
			
					Call InsertDataPaymentVoucherDetail				
					connStr.CommitTrans
					

					cmdView	= "True"
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If 
			else
				cmdView	= "True"
			end if
	'Response.Write "<br/>q-->" & q
	'Response.Write "<br/>qq-->" & qq
 else
    cmdView	= request("cmdView")
 End if

Dim rsVoucher
q = " SELECT bankcode" & _
	"        , datecreated" & _
	"        , totalamount" & _
	"        , remarks" & _
	"        , accountingperiod" & _
	"        , accountingyear" & _
	"        , accountingyear||LPAD(accountingperiod,2,0) accountingyearperiod" & _
	" FROM PaymentVoucher" & _
	" WHERE VoucherCode='" & fcCode & "'" & _
	"       AND fcBA = '" & session("fcBA") & "'"
set rsVoucher = server.CreateObject("Adodb.recordset")
rsVoucher.Open q, connstr
if Not rsVoucher.Eof then
	xCode=rsVoucher("BankCode")
	DATECREATED=rsVoucher("DateCreated")
	TOTALAMOUNT = rsVoucher("TOTALAMOUNT")
	HeaderRemarks=rsVoucher("Remarks")	
	ACCOUNTINGYEAR =rsVoucher("ACCOUNTINGYEAR")
	ACCOUNTINGPERIOD =rsVoucher("ACCOUNTINGPERIOD")
	ACCOUNTINGYEARPERIOD = rsVoucher("accountingyearperiod")
End if

 
if Request("txtJenis") = "Edit" then
	q = " SELECT locationtype" & _
		"        , locationcode" & _
		"        , jobcode" & _
		"        , amount" & _
		"        , reference" & _
		"        , remarks" & _
		" FROM paymentvoucherdetail" & _
		" WHERE fcba = '" & session("fcBA") & "'" & _
		"       AND vouchercode = '" & fcCode & "'" & _
		"       AND grouplineno = '" & grouplineno & "'" 

	set rsVoucher = server.CreateObject("Adodb.recordset")
	rsVoucher.Open q, connstr
	if Not rsVoucher.Eof then
		While Not rsVoucher.Eof
			if cdbl(to_nol(rsVoucher("amount"))) < 0 then
				Jobcode_cr = rsVoucher("jobcode")
				LocationType_cr = rsVoucher("locationtype")
				LocationCode_cr = rsVoucher("locationcode")
			else
				Jobcode_db = rsVoucher("jobcode")
				LocationType_db = rsVoucher("locationtype")
				LocationCode_db = rsVoucher("locationcode")
				Amount = rsVoucher("amount")
				Reference = rsVoucher("reference")
				Remarks = rsVoucher("remarks")
			end if
			rsVoucher.MoveNext
		Wend
	End if
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
	        " where fcba = '" & Session("fcBA") & "'" & _
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

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & DATECREATED & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('" & val_trigger_system & "')"
'response.write oSQLQuery			
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
                   
                    <form name="frmEntry" method="post" action="JournalReclass.asp">
					
                    <font color="red">
                        <%= varStringError  %></font>
                    <table class="data" width="100%">
                        <tr>
                            <td align="left" style="width:20%;">
                                Voucher No
                            </td>
                            <td align="left">
                                <input type="text" maxlength="30" name="fcCode" size="30" value="<%=fcCode%>" title="">
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                Date
                            </td>
                            <td align="left">
                                <input class="datepicker" type="text" maxlength="10" name="DATECREATED" size="10" value="<%=DATECREATED%>">
                                <!--<a href="javascript:show_calendar('frmEntry.DATECREATED');" onMouseOver="window.status='Date Picker';return true;"
                                    onmouseout="window.status='';return true;">
                                    <img src="Images/show-calendar.gif" width="24" height="22" border="0" align="top"></a>-->
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                        </tr>
						 <tr>
                            <td align="left">
                                Accounting Period
                            </td>
							
							<% 	Dim rsComboA,Data_ACCOUNTINGYEARPERIOD

								q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from  ACCOUNTINGYEARPERIOD where fcBA='" &  Session("fcBA") & "'  and to_char(valid_until ,'YYYY') = '"& Year(Session("v_Costbook")) &"'  and to_number(to_char(valid_until ,'MM')) > to_number('"& Month(Session("v_Costbook")) &"') -1  order by accountingyear||lpad(accountingperiod,2,0)"

								Set rsComboA = Server.CreateObject("ADODB.Recordset")
								'response.write q
								rsComboA.CursorType = 3
								rsComboA.CursorLocation = 3
								rsComboA.Open q, connstr
								if not rsComboA.eof then%>
								<td align="left"> 
									<select name="ACCOUNTINGYEARPERIOD">    
										<% Do While Not rsComboA.eof %>
										<% If Cstr(""&ACCOUNTINGYEARPERIOD) = Cstr(rsComboA("ACCOUNTINGYEARPERIOD")) Then 
											Data_ACCOUNTINGYEARPERIOD = "true"
										else
											Data_ACCOUNTINGYEARPERIOD = "false"
										End if%>
										<option value="<%=rsComboA("ACCOUNTINGYEARPERIOD")%>" <% If Cstr("" & ACCOUNTINGYEARPERIOD) = Cstr(rsComboA("ACCOUNTINGYEARPERIOD")) Then %> Selected <%End if%>><%=rsComboA("Description")%></option> 
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
								Remarks
							</td>
							<td align="left">
								<textarea name="HeaderRemarks" cols="50" rows="3"><%=HeaderRemarks%></textarea>
							</td>
						</tr>
					<tr>
						<td colspan="2" align="left">
							<input type="Hidden" value="True" name="cmdView">
							<input type="Hidden" value="<%=request("txtTable")%>" name="txtTable">
							<input type="Hidden" value="<%=request("txtModule")%>" name="txtModule">
							<input type="Submit" name="Add" value="Select">&nbsp;&nbsp;
							<a href="javascript:launchwin('PrintVoucher.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
								<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana"
									size="1">Print</font></a>&nbsp;&nbsp;
							<%If val_IsTrue = "Y" Then%>
							<a href="JournalReclass.asp?fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&txtJenis=Header">
								<img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');">
							</a>
							<%Else%>
							<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
							<%End if%>
							<!-- HADIR -->
							<%
							dataModule 		= "Bank"
							dataSubmodule 	= "Payment Voucher"
							dataDetail 		= "NULL"
							%>
							<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
							<!-- END:HADIR -->
						</td>
					</tr>
			</table>
        </form>
<%if cmdView = "True" and "" & fcCode <> "" and "" & ACCOUNTINGYEARPERIOD <> "" then
	Dim rsPaymentVoucher, intRecordsPerPage, cColor

	q = " SELECT vouchercode" & _
		"        , lineno" & _
	    "        , sourcetype" & _
	    "        , locationtype" & _
	    "        , locationcode" & _
	    "        , jobcode" & _
	    "        , amount" & _
	    "        , reference" & _
	    "        , remarks" & _
	    "        , invoicecode" & _
	    "        , agreementcode" & _
	    "        , fcentry" & _
	    "        , fcedit" & _
	    "        , fcip" & _
	    "        , fcba" & _
	    "        , grouplineno" & _
	    " FROM paymentvoucherdetail" & _
	    " WHERE vouchercode = '" & fcCode & "'" & _
		"       AND fcba = '" & session("fcBA") & "'" & _
		" ORDER BY grouplineno, lineno"
	'response.Write q	
	Set rsPaymentVoucher = Server.CreateObject("ADODB.Recordset")
	rsPaymentVoucher.CursorType = 3
	rsPaymentVoucher.CursorLocation = 3
	rsPaymentVoucher.Open q, connstr
%>
        <form action="JournalReclass.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%">
            <tr>
                <td align="left" style="width:20%;">Debet</td>
				<td align="left">
							<input type="text" size="11" id="JobCode_db" name="JobCode_db" value="<%=JobCode_db%>" placeholder="Jobcode">
							<input type="text" size="11" id="LocationType_db" name="LocationType_db" value="<%=LocationType_db%>" placeholder="Target Type">
							<input type="text" size="11" id="LocationCode_db" name="LocationCode_db" value="<%=LocationCode_db%>" placeholder="Target Code">
				</td>
            </tr>
            <tr>
				<td align="left">Credit</td>
                <td align="left">
							<input type="text" size="11" id="JobCode_cr" name="JobCode_cr" value="<%=JobCode_cr%>" placeholder="Jobcode">
							<input type="text" size="11" id="LocationType_cr" name="LocationType_cr" value="<%=LocationType_cr%>" placeholder="Target Type">
							<input type="text" size="11" id="LocationCode_cr" name="LocationCode_cr" value="<%=LocationCode_cr%>" placeholder="Target Code">
				</td>
            </tr>
           

            <tr>
                <td align="left">
                    Amount
                </td>
                <td align="left">
                    <input type="text" size="20" id="Amount" name="Amount" value="<%=Amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
                </td>
            </tr>
			<tr>
                <tr>
                    <td class="wip" align="left" style=" border-right:none;">
                        <div>Document References</div>
                    </td>
                    <td class="wip" align="left" style=" border-left:none;">
                        <input type="text" size="59" placeholder="No Berita Acara" id="reference" name="reference" value="<%=reference%>">
                    </td>
                </tr>
            </tr>
            <tr>
                <td align="left">
                    Remarks
                </td>
                <td align="left">
                    <textarea name="Remarks" cols="50" rows="3"><%=Remarks%></textarea>
                </td>
            </tr>
			<%if request("txtJenis")= "Edit" then%>
            <tr>
				<td> <input type="hidden" name="grouplineno" value="<%=grouplineno%>"></td>
			</tr>
			<%end if%>
            <tr>
                <td colspan="2" align="left">
                    <!-- diambil dari form yang diatas -->
                    <input type="hidden" name="fcCode" value="<%=fcCode%>">
					<input type="hidden" name="ACCOUNTINGYEARPERIOD" value="<%=ACCOUNTINGYEARPERIOD%>">
                    <input type="hidden" name="DATECREATED" value="<%=DATECREATED%>">
                    <input type="hidden" name="xCode" value="<%=xCode%>">
                    <input type="hidden" name="HeaderRemarks" value="<%=HeaderRemarks%>">
                    <input type="Hidden" value="<%=request("txtTable")%>" name="txtTable">
                    <input type="Hidden" value="<%=request("txtModule")%>" name="txtModule">
					<%If val_IsTrue = "Y" Then%>
						<%IF Request("txtJenis")="Edit" then%>
							<input type="hidden" name="txtJenis" value="Edit">
							<input type="Submit" name="Add" value="Update" onClick="return validasiform();">
						<%Else%>
							<input type="hidden" name="txtJenis" value="New">
							<input type="Submit" name="Add" value="Add" onClick="return validasiform();">
						<%End if%>
					<%End if%>
                    <input type="Reset"  name="Reset" value="Reset">
                    <input type="Button" value="Back" name="B2" onClick="javascript:history.back(1)">
                    <font color="#FF0000">* (Must be Filled)</font>
                </td>
            </tr>
        </table>
        <table class="data xthf-site">
            <br>
            <thead>
            <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
                <th align="left">No</th>
                <th colspan="2">&nbsp;</th>
                <th align="left">Debet Credit</th>
                <th align="left">Jobcode</th>
                <th align="left">Target Type</th>
                <th align="left">Target Code</th>
                <th align="left">Amount</th>
                <th align="left">Reference</th>
                <th align="left">Remarks</th>
            </tr>
            </thead>
            <tbody>
            <%
   if Not rsPaymentVoucher.eof then
      currGL = ""
	  prevGL = ""
	  no = 0
	  While Not rsPaymentVoucher.EOF
		currGL = cdbl(to_nol(rsPaymentVoucher("grouplineno")))
		if currGL <> prevGL then
			no = no +1
		end if

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
            <tr bgcolor="<%=cColor%>">
				<%if currGL <> prevGL then%>
                <td align="left" valign="top" rowspan="2">
                    <%=NO%>
                </td>
				
				<!-- Cek Data Close or Not -->
				<%If val_IsTrue = "Y" Then%>
					<td valign="top" rowspan="2">
						<a href="JournalReclass.asp?fcCode=<%=rsPaymentVoucher("VoucherCode")%>&grouplineno=<%=rsPaymentVoucher("grouplineno")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPaymentVoucher(1)%>" border="0">
						</a>
					</td>
					<td valign="top" rowspan="2">
						<a href="JournalReclass.asp?fcCode=<%=rsPaymentVoucher("VoucherCode")%>&grouplineno=<%=rsPaymentVoucher("grouplineno")%>&txtJenis=Details">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPaymentVoucher(1)%>" border="0" onclick="return confirm('Are you sure to delete');">
						</a>
					</td>
                <%Else%>
					<td colspan="2" rowspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
                <%End if%>
				<%end if%>
				<%if cdbl(to_nol(rsPaymentVoucher("amount"))) >= 0 then%>
					<td align="center">Db</td>
				<%else%>
					<td align="center">&nbsp;&nbsp;&nbsp;&nbsp;Cr</td>
				<%end if%>
				<td align="left" title="<%=rsPaymentVoucher("jobcode").Name%>" valign="top"><%=rsPaymentVoucher("jobcode")%></td>
				<td align="left" title="<%=rsPaymentVoucher("locationtype").Name%>" valign="top"><%=rsPaymentVoucher("locationtype")%></td>
				<td align="left" title="<%=rsPaymentVoucher("locationcode").Name%>" valign="top"><%=rsPaymentVoucher("locationcode")%></td>
				<td align="left" title="<%=rsPaymentVoucher("amount").Name%>" valign="top"><%=rsPaymentVoucher("amount")%></td>
				<td align="left" title="<%=rsPaymentVoucher("reference").Name%>" valign="top"><%=rsPaymentVoucher("reference")%></td>
				<td align="left" title="<%=rsPaymentVoucher("remarks").Name%>" valign="top"><%=rsPaymentVoucher("remarks")%></td>
                
            </tr>
<%
		prevGL = currGL
		rsPaymentVoucher.MoveNext
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPaymentVoucher.Close
   set rsPaymentVoucher = Nothing%>
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
		<script>
			function validasiform (){
				var jc_db = document.getElementById("JobCode_db");
				var tt_db = document.getElementById("LocationType_db");
				var tc_db = document.getElementById("LocationCode_db");
				var jc_cr = document.getElementById("JobCode_cr");
				var tt_cr = document.getElementById("LocationType_cr");
				var tc_cr = document.getElementById("LocationCode_cr");
				var amt = document.getElementById("Amount").value;
				
				if (jc_db.value === '') {
					alert ("Jobcode Debet harus diisi");
					jc_db.focus();
					return false;
				}
				
				if (tt_db.value === '') {
					alert ("Target Type Debet harus diisi");
					tt_db.focus();
					return false;
				}
				
				if (tc_db.value === '') {
					alert ("Target Code Debet harus diisi");
					tc_db.focus();
					return false;
				}
				
				if (jc_cr.value === '') {
					alert ("Jobcode Kredit harus diisi");
					jc_cr.focus();
					return false;
				}
				
				if (tt_cr.value === '') {
					alert ("Target Type Kredit harus diisi");
					tt_cr.focus();
					return false;
				}
				
				if (tc_cr.value === '') {
					alert ("Target Code Kredit harus diisi");
					tc_cr.focus();
					return false;
				}
				
				var xamt = amt.replace(/,/g , "");
				
				if (amt === 0 || amt === '') {
					alert ("Amount harus diisi");
					return false;
				} else {
					if (xamt < 0) {
						alert ("Amount harus positif");
						return false;
					}
				}
			}
		</script>
		
		
    </body>
</html>