<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
    <title>ePlaS -  Reclass Voucher</title>
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
</script>
    <script language="JavaScript">
        function showDetails(winurl, p_item) {
            window.open(winurl + '&p_item=' + p_item, width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1);
        }
    </script>
</head>
<%if request("Add")="" then%>
<body onLoad="document.frmEntry.fcCode.focus()">
    <%Else%>
    <body onload="document.eForm.JobCode.focus()">
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
	                " from trigger_reclass" & _
	                " where " & _
					"  jobcode = '" & Trim(initialJob) & "'" & _
					"  and fcba = '" & Trim(initialFcba) & "'"
					
    ' response.write strAvailable
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

 Dim check

 Dim xCode,fcCode,cmdView,DATECREATED,VoucherCode,QUANTITY,AMOUNT,OTHERITEMDESC,SourceType
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis,TOTALAMOUNT,HeaderRemarks
 Dim rsInvoiceRecord, varTotalAmount, invoiceUpdate, recordcheck, rsrecordlineno, recordWIPLINE,invoiceRecord
 Dim rsRECLASSVOUCHERDETAIL, varamountOriginal, recordinvoicepaid, vartotal, varError, ivp,varmaxlineno,recordlineno
 Dim Jobcode, LocationType, LocationCode, Reference, Remark,InvoiceCode,AGREEMENTCODE,AGREEMENTLINENO ,recordwipcheck
 Dim varSupplierCode, varStringError, varlineno, WIPPAIDRecord, rsWIPPAIDRecord, WIPUpdate ,rsrecordWIPLINE ,W, rsrecordwipcheck, WIPDATE, varContract
 Dim InsertIntoCostActivity, rsqMaxLineNo, qMaxLineNo, MaxLineNo, IsCostbook
 Dim MaxLineNopvd
 Dim rlineno, dataModule, dataSubmodule, dataDetail, qq, recRECLASSVOUCHERDETAIL
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD
 Dim xACCOUNTINGYEARPERIOD
 Dim ReclassJob, RECLASSREFERENCE,YEARREFERENCE,MONTHREFERENCE

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

LocationType=Request("LocationType")
LocationCode=Request("LocationCode")

Amount=CDBL(Request("Amount"))
InvoiceCode=Request("InvoiceCode")
AGREEMENTCODE=Request("AGREEMENTCODE")
AGREEMENTLINENO=Request("AGREEMENTLINENO")
WIPDATE=Request("WIPDATE")
rlineno = Request("txtlineno")
ReclassJob = Request("ReclassJob")
RECLASSREFERENCE = Request("RECLASSREFERENCE")

IF  RECLASSREFERENCE <> "" then
	YEARREFERENCE  = MID(RECLASSREFERENCE,1,4)
	MONTHREFERENCE = MID(RECLASSREFERENCE,5,2)
END IF
ACCOUNTINGYEARPERIOD = Request("ACCOUNTINGYEARPERIOD")
 
 if ACCOUNTINGYEARPERIOD <> "" then
    ACCOUNTINGYEAR = MID(ACCOUNTINGYEARPERIOD,1,4)
	ACCOUNTINGPERIOD = MID(ACCOUNTINGYEARPERIOD,5,2)
 end if
 
 'response.write "ACCOUNTINGYEAR " & ACCOUNTINGYEAR  & "ACCOUNTINGPERIOD " & ACCOUNTINGPERIOD 

  
 IF Request("SourceType")="" then
    SourceType="BA" 'Default if Entry is Blank
 else
    SourceType=Request("SourceType")
 End if
dim tt

Sub DeleteDataRECLASSVOUCHERDETAIL
	q = " delete from RECLASSVOUCHERDETAIL" & _
		" where fcBA = '" & Session("fcBA") & "'" & _
		"       and vouchercode = '" & fcCode & "'" & _
		"       and lineno = '" & rlineno & "'"
	connStr.execute(q)
End Sub

Sub InsertDataRECLASSVOUCHERDETAIL
	qq = "Insert into  RECLASSVOUCHERDETAIL ( " & _
					"VOUCHERCODE, " & _
					"SOURCETYPE, " & _
					"LOCATIONTYPE, " & _
					"LOCATIONCODE, " & _
					"JOBCODE, " & _
					"AMOUNT, " & _
					"REFERENCEMONTH, " & _
					"FCENTRY, " & _
					"FCEDIT, " & _
					"FCIP, " & _
					"FCBA, " & _
					"lineno, " & _
					"LASTUPDATE, " & _
					"LASTTIME) " & _
				"values ('" & fcCode & "','" & _
							SourceType & "','" & _
							LocationType & "','" & _ 
							LocationCode &  "','" & _
							ReclassJob & "','" & _
							Amount & "','" & _
							RECLASSREFERENCE & "','" & _ 
							Session("CodeName") & "','" & _
							Session("CodeName") & "','" & _
							Request.ServerVariables("REMOTE_ADDR") &"','" & _ 
							Session("fcBA") & "'," & _
							" '" & rlineno & "'," & _
							" sysdate, " & _
							" To_Char(sysdate,'HH24:MI')" & _
							" )"
	 'response.write qq
	connStr.execute(qq)
End Sub

Sub UpdateDataRECLASSVOUCHERDETAIL
	Call DeleteDataRECLASSVOUCHERDETAIL
	Call InsertDataRECLASSVOUCHERDETAIL
	
End Sub
		
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then	
		if Request("txtJenis")= "Header" then
			Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
			If Cek_IsDailyClosing = "FALSE" then
				connstr.BeginTrans
				q = " Delete from RECLASSVOUCHERDETAIL where  VoucherCode = '" & fcCode & "' and fcBA = '" & Session("fcBA") & "'"
				connStr.execute(q)
				
				q = " Delete from RECLASSVOUCHER where  VoucherCode = '" & fcCode & "' and fcBA = '" & Session("fcBA") & "'"
				connStr.execute(q)
				
				if err.number  = 0 then
					connStr.CommitTrans
				else
					connStr.RollbackTrans
				end if
			else
				err.number = 1001
				err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
			End If  
			
			cmdView	= "True"
	   
		elseif Request("txtJenis")= "Details" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then
					connstr.BeginTrans
					call DeleteDataRECLASSVOUCHERDETAIL
					
					oSQLQuery = " select nvl(bank_api.get_trigger_system(bankcode, fcba), 'Bank') trigger_system" & _
								" from RECLASSVOUCHER" & _
								" where fcba = '" & Session("fcBA") & "'" & _
								"       and vouchercode = '" & fcCode & "'"
					set oRecordSet = server.CreateObject("Adodb.recordset")
					oRecordSet.Open oSQLQuery, connstr
					If Not oRecordSet.Eof then
						IsCostbook = oRecordSet("trigger_system")
					Else
						IsCostbook = "Bank"
					End if
					oRecordSet.Close
					Set oRecordSet = Nothing
			
					If err.number  = 0 then
						connStr.CommitTrans
					Else
						connStr.RollbackTrans
					End if
				Else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If 
				cmdView	= "True"
	
		elseif Request("txtJenis")="New" then
				
				 'response.write "datanya " & Request("ACCOUNTINGYEARPERIOD") 
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then
				    Dim rsPO
					connstr.BeginTrans
					Set rsPO = Server.CreateObject("ADODB.Recordset")
					rsPO.Open "Select VoucherCode from RECLASSVOUCHER where VoucherCode='" & fcCode & "'AND FCBA = '" & Session("FCBA") & "'", connstr
					IF rsPO.eof then	    
						q	=	" Insert into RECLASSVOUCHER (" & _
								"   VOUCHERCODE" & _
								" , BANKCODE" & _
								" , DATECREATED" & _
								" , ACCOUNTINGYEAR " &_
								" , ACCOUNTINGPERIOD " &_
								" , TOTALAMOUNT" & _
								" , SUPPCONTCODE" & _
								" , VOUCHERTYPE" & _
								" , REMARKS" & _
								" , FCENTRY" & _
								" , FCEDIT" & _
								" , FCIP" & _
								" , FCBA" & _
								" , LASTUPDATE" & _
								" , LASTTIME" & _
								" )" & _
								" Values (" & _
								" '" & fcCode & "'" & _
								" , '" & xCode & "'" & _
								" , To_Date('" & DATECREATED & "','mm/dd/yyyy')" & _
								" ," & ACCOUNTINGYEAR & " " &_
								" ,to_number(" & ACCOUNTINGPERIOD & ")" &_
								" , 0" & _
								" , ''" & _
								" , ''" & _
								" , '" &  request("HeaderRemarks") & "'" & _
								" , '" & Session("CodeName") & "'" & _
								" , '" & Session("CodeName") & "'" & _
								" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
								" , '" & Session("fcBA") & "'" & _
								" , sysdate" & _
								" , To_Char(sysdate,'HH24:MI')" & _
								" )"
								'response.Write("q " & q )
								connStr.execute(q)
						 
					End if
					rsPO.Close
					set rsPO = Nothing	
			  
					IF err.number = 0 then
						If CheckAvailableJobZerorise(ReclassJob, Session("FCBA")) = 1 then					
							qMaxLineNo = " select nvl(max(nvl(lineno,0)),0) maxLineNo " & _
										" from RECLASSVOUCHERDETAIL " & _
										" where " & _
										" fcba = '" & Session("fcBA") & "' "
													
							Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
							rsqMaxLineNo.CursorType = 3
							rsqMaxLineNo.CursorLocation = 3
							rsqMaxLineNo.Open qMaxLineNo, connstr

							if not rsqMaxLineNo.eof then
							   maxLineNopvd = cdbl(rsqMaxLineNo("maxLineNo")) + 1
							else
							   maxLineNopvd = 1
							end if
							rsqMaxLineNo.close
							set rsqMaxLineNo = Nothing
									
							rlineno = maxLineNopvd
							 
							Call InsertDataRECLASSVOUCHERDETAIL
                            If err.number  = 0 then
								connStr.CommitTrans
							Else
								connStr.RollbackTrans
							End if
							
						Else
							err.number = 1001
							err.description = "Jobcode """ & ReclassJob & """ tidak boleh di gunakan untuk Reclass, silahkan pilih jobcode yang benar !"
						End If
					End if	
			
				Else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
		
		elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), DATECREATED)
				If Cek_IsDailyClosing = "FALSE" then
				    connstr.BeginTrans
				    If  CheckAvailableJobZerorise(ReclassJob, Session("FCBA")) = 1 then
					   Call UpdateDataRECLASSVOUCHERDETAIL
					   
					Else
						err.number = 1001
						err.description = "Jobcode """ & ReclassJob & """ tidak boleh di gunakan untuk zerorise, silahkan pilih jobcode yang benar !"
					End If
				else
					err.number = 1001
					err.description = "Tanggal """ & DATECREATED & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				end if	
                cmdView	= "True"				
		else
			cmdView	= "True"
		End if				
End if

cmdView	= "True"	
			

 Dim rsVoucher
 set rsVoucher = server.CreateObject("Adodb.recordset")

 q="select * from RECLASSVOUCHER  " & _
  " where VoucherCode='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
  
 rsVoucher.Open q, connstr
 if Not rsVoucher.Eof then
    DATECREATED=rsVoucher("DateCreated")
	xCode=rsVoucher("BankCode")
	HeaderRemarks=rsVoucher("Remarks")
	TOTALAMOUNT=rsVoucher("TOTALAMOUNT")
	ACCOUNTINGYEAR =rsVoucher("ACCOUNTINGYEAR")
	ACCOUNTINGPERIOD =rsVoucher("ACCOUNTINGPERIOD")
 End if


 
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

' response.Write  oSQLQuery			
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
                   
                    <form name="frmEntry" method="post" action="ReclassVoucherDetail.Asp">
					
                    <font color="red">
                        <%= varStringError  %></font>
                    <table class="data" width="100%">
                        <tr>
                            <td align="left">
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
								q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from  ACCOUNTINGYEARPERIOD where fcBA='" &  Session("fcBA") & "'  and to_char(valid_until ,'YYYY') = '"& Year(Session("v_Bank")) &"'  and to_number(to_char(valid_until ,'MM')) >= to_number('"& Month(Session("v_Bank")) &"') order by accountingyear||lpad(accountingperiod,2,0)"

								Set rsComboA = Server.CreateObject("ADODB.Recordset")
								
								rsComboA.CursorType = 3
								rsComboA.CursorLocation = 3
								rsComboA.Open q, connstr
								if not rsComboA.eof then%>
								<td align="left"> 
									<select name="ACCOUNTINGYEARPERIOD">    
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
                                Bank Code
                            </td>
                            <td align="left">
                                <select name="xCode">
                                    <%cmdCallCombo "Bank",xCode,"FCNAME","JURNAL"%>
                                </select>
								<font color="#FF0000">&nbsp;*</font>
                            </td>
                            <tr>
                                <td align="left">
                                    Amount Total
                                </td>
                                <td align="left">
                                    <input type="text" maxlength="50" name="TOTALAMOUNT" size="10" value="<%=TOTALAMOUNT%>"
                                        readonly="" style="background:#FFFFCC;">
                                </td>
                            </tr>
                           
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
							<a href="ReclassVoucherDetail.Asp?fcCode=<%=fcCode%>&DATECREATED=<%=DATECREATED%>&txtJenis=Header">
								<img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');">
							</a>
							<%Else%>
							<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
							<%End if%>
							<!-- HADIR -->
							<%
							dataModule 		= "Bank"
							dataSubmodule 	= "Reclass Voucher"
							dataDetail 		= "NULL"
							%>
							<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
							<!-- END:HADIR -->
						</td>
					</tr>
			</table>
        </form>
<%   
    ' response.write " cmdView " & cmdView & " , trim(xcode) " & trim(xcode)
    if cmdView = "True" AND LEN(TRIM(fcCode)) > 3 and len(trim(xcode)) > 0 Then
		Dim rsRECLASSVOUCHER, intRecordsPerPage, cColor

		q = " select vouchercode" & _
			"        , lineno" & _
			"        , sourcetype" & _
			"        , locationtype" & _
			"        , locationcode" & _
			"        , jobcode" & _
			"        , amount" & _
			"        , referencemonth" & _
			"        , fcentry" & _
			"        , fcedit" & _
			"        , fcip" & _
			"        , fcba" & _
			" from RECLASSVOUCHERDETAIL" & _
			" where vouchercode = '" & fcCode & "'" & _
			"       and fcba = '" & session("fcBA") & "'" & _
			" order by fcba, vouchercode, lineno"
		'response.Write q	
		Set rsRECLASSVOUCHER = Server.CreateObject("ADODB.Recordset")
		rsRECLASSVOUCHER.CursorType = 3
		rsRECLASSVOUCHER.CursorLocation = 3
		rsRECLASSVOUCHER.Open q, connstr
%>
        <form action="ReclassVoucherDetail.Asp" method="post" name="eForm" style="display:table;">
        <table class="data" cellpadding="0" cellspacing="0" width="100%">
           
               
					<input type="hidden" size="11" name="txtlineno" value="<%=rlineno%>">
                    
               
            <tr>
                <td align="left">
                    Job Code
                </td>
				<% 	Dim rsCombo,Data_JobsAccrual
				    q = "select jobcode, job_api.get_jobcode_name(jobcode,fcba) Jobname   from trigger_reclass  where fcba='" &  Session("fcBA") & "'  order by jobcode"
					set rsCombo = Server.CreateObject("ADODB.Recordset")
					'response.write q
					rsCombo.CursorType = 3
					rsCombo.CursorLocation = 3
					rsCombo.Open q, connstr
					if not rsCombo.eof then
				%>
				<td align="left"> 
							<select name="ReclassJob">    
								<% Do While Not rsCombo.eof %>
									
									  <option value="<%=rsCombo("jobcode")%>" <% If Request("ReclassJob") = rsCombo("jobcode") Then %> Selected <%End if%>><%=rsCombo("jobcode")%> : <%=rsCombo("jobname")%></option> 
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
                <td align="left">
                    Target Type
                </td>
                <td align="left">
                    <input type="text" size="11" name="LocationType" value="<%=LocationType%>">&nbsp;<font
                        color="#FF0000">* <a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
                            <font face="Verdana" size="1">Location Type</font></a>
                </td>
            </tr>
            <tr>
                <td align="left">
                    Target Code
                </td>
                <td align="left">
                    <input type="text" size="11" name="LocationCode" value="<%=LocationCode%>">&nbsp;<font
                        color="#FF0000">*</font>
                </td>
            </tr>
             <tr>
                <td align="left">
                    Source Type
                </td>
                <td align="left">
                    <input type="text" size="11" name="SourceType" value="<%=SourceType%>">&nbsp;<font color="#FF0000">*</font>
                </td>
            </tr> 
            
            <tr>
                <td align="left">
                    Reference Month
                </td>
               	<% 	            Dim rsComboB,Data_Reference
								q = "select accountingyear||lpad(accountingperiod,2,0) ACCOUNTINGYEARPERIOD,Description from  ACCOUNTINGYEARPERIOD where fcBA='" &  Session("fcBA") & "'  and to_char(valid_until ,'YYYY') = '"& Year(Session("v_Bank")) &"'  and to_number(to_char(valid_until ,'MM')) <= to_number('"& Month(Session("v_Bank")) &"') order by accountingyear||lpad(accountingperiod,2,0)"

								Set rsComboB= Server.CreateObject("ADODB.Recordset")
								
								rsComboB.CursorType = 3
								rsComboB.CursorLocation = 3
								rsComboB.Open q, connstr
								if not rsComboB.eof then%>
								<td align="left"> 
									<select name="RECLASSREFERENCE">    
										<% Do While Not rsComboB.eof %>
										<% If Request("MONTHREFERENCE") = rsComboB("ACCOUNTINGYEARPERIOD") Then 
											Data_Reference = "true"
										else
											Data_Reference = "false"
										End if%>
										<option value="<%=rsComboB("ACCOUNTINGYEARPERIOD")%>" <% If Request("RECLASSREFERENCE") = rsComboB("ACCOUNTINGYEARPERIOD") Then %> Selected <%End if%>><%=rsComboB("Description")%></option> 
										<%  rsComboB.movenext
										   loop
									if Data_Reference = "false" then %>
										<option value="<%=YEARREFERENCE&MONTHREFERENCE%>" Selected>Tahun <%=YEARREFERENCE%> Bulan <%=MONTHREFERENCE%></option>
									<% end if %>
									
									<%
										End if
										rsComboB.Close
										set rsComboB = Nothing
										q=""
										%>
									</select> <font color="#FF0000">*</font>
								</td>
            </tr>
            
			<tr>
                <td align="left">
                    Amount
                </td>
                <td align="left">
                    <input type="text" size="11" name="Amount" value="<%=Amount%>">&nbsp;<font color="#FF0000">*
                        (+) Debet (-) Credit</font>
                </td>
            </tr>
			
            </td></tr>
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
							<input type="Submit" name="Add" value="Update">
						<%Else%>
							<input type="hidden" name="txtJenis" value="New">
							<input type="Submit" name="Add" value="Add">
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
                <th align="left">
                    NO
                </th>
                <th colspan="2">&nbsp;
                    
                </th>
                <th align="left">
                    CODE
                </th>
				<th align="left">
                    LINE NO
                </th>
                <th align="left">
                    SOURCE TYPE
                </th>
                <th align="left">
                    LOCATION TYPE
                </th>
                <th align="left">
                    LOCATION CODE
                </th>
                <th align="left">
                    JOB CODE
                </th>
                <th align="left">
                    AMOUNT
                </th>
                <th align="left">
                    REFERENCES MONTH
                </th>  
            </tr>
            </thead>
            <tbody>
            <%
			
   if Not rsRECLASSVOUCHER.eof then
      nCountofField=rsRECLASSVOUCHER.fields.count-1
	  no=1
	  While Not rsRECLASSVOUCHER.EOF

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
				<%If val_IsTrue = "Y" Then%>
					<td valign="middle" width="10">
						<a href="ReclassVoucherDetail.Asp?fcCode=<%=rsRECLASSVOUCHER("VoucherCode")%>&txtlineno=<%=rsRECLASSVOUCHER("lineno")%>&ReclassJob=<%=rsRECLASSVOUCHER("JobCode")%>&SourceType=<%=rsRECLASSVOUCHER("SourceType")%>&LocationType=<%=rsRECLASSVOUCHER("LocationType")%>&LocationCode=<%=rsRECLASSVOUCHER("LocationCode")%>&Amount=<%=rsRECLASSVOUCHER("Amount")%>&RECLASSREFERENCE=<%=rsRECLASSVOUCHER("Referencemonth")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsRECLASSVOUCHER(1)%>" border="0">
						</a>
					</td>
					<td>
						<a href="ReclassVoucherDetail.Asp?fcCode=<%=rsRECLASSVOUCHER("VoucherCode")%>&txtlineno=<%=rsRECLASSVOUCHER("lineno")%>&DATECREATED=<%=DATECREATED%>&ReclassJob=<%=rsRECLASSVOUCHER("JobCode")%>&LocationType=<%=rsRECLASSVOUCHER("LocationType")%>&LocationCode=<%=rsRECLASSVOUCHER("LocationCode")%>&Amount=<%=rsRECLASSVOUCHER("Amount")%>&RECLASSREFERENCE=<%=rsRECLASSVOUCHER("Referencemonth")%>&txtJenis=Details">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsRECLASSVOUCHER(1)%>" border="0" onclick="return confirm('Are you sure to delete');">
						</a>
					</td>
                <%Else%>
					<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
                <%End if%>
                <%for i = 0 to nCountofField-4%>
					<td align="left" title="<%=rsRECLASSVOUCHER(i).Name%>" valign="top">
						<%=rsRECLASSVOUCHER(i)%>
					</td>
                <%next%>
                
            </tr>
            <%
		  rsRECLASSVOUCHER.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsRECLASSVOUCHER.Close
   set rsRECLASSVOUCHER = Nothing%>
   </tbody>
        </table>
        
        </form>
        <%End if%>
		 
				<!-------------akhir!------------->
			</div>
		</div>
		<!--#include file="bottom.jsp"-->

       <%ON ERROR GOTO 0%>
    </body>
</html>