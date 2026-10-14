<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
	<title>ePlaS - Recurring Settlement</title>
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
</script>

<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">

<style>
.close-button {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 30px;
  height: 30px;
  position: relative;
  padding: 0 1px 0 0;
  border: none;
  -webkit-border-radius: 1em;
  border-radius: 1em;
  font: normal 8em/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #6d6d6d;
  cursor:pointer;
}

.close-button:hover {background: #f72c2c;}

.close-button::before {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 5px;
  height: 20px;
  position: absolute;
  content: "";
  top: 5px;
  left: 13px;
  border: none;
  font: normal 100%/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #ffffff;
  text-shadow: none;
  -webkit-transform: rotateZ(45deg)   ;
  transform: rotateZ(45deg)   ;
}

.close-button::after {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 5px;
  height: 20px;
  position: absolute;
  content: "";
  top: 5px;
  left: 13px;
  border: none;
  font: normal 100%/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #ffffff;
  text-shadow: none;
  -webkit-transform: rotateZ(-45deg)   ;
  transform: rotateZ(-45deg)   ;
}
</style>

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
Dim v_txtagreementcode, v_txtmonth, v_txtyear, MinYr, MaxYr, mon, Yr, n, SudahClose, v_txtagreementfinalno
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
Dim wipdate : wipdate = ""

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

settlementno = Request("settlementno")
documentno = Request("documentno")
accountingyear = Request("accountingyear")
startingdate = Request("startingdate")
finishdate = Request("finishdate")
rundate = Request("rundate")
numberofrun = Request("numberofrun")
amount = Request("amount")
prepaymentjob = Request("prepaymentjob")
temporarycashjob = Request("temporarycashjob")
status = Request("status")

maxcounter = request("maxcounter")
if "" & maxcounter = "" then maxcounter = 1


jobcode_edit = request("jobcode_edit")
targettype_edit = request("targettype_edit")
amount_edit = request("amount_edit")

txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

if "" & status = "" then status = "Planned"
	
			
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" OR Request("txtJenis") = "Confirmed" OR Request("txtJenis") = "Cancelled" OR Request("txtJenis") = "Executed" OR Request("txtJenis") = "RePlanned" Then
	if Request("txtJenis")= "Header" then
		'Delete data from table recurringsettlementdtl
		q = " DELETE FROM recurringsettlementdtl" & _
			" WHERE settlementno = '" & settlementno & "'" & _
			"       AND fcba = '" & Session("fcBA") & "'"

		'Response.Write q
		connStr.execute(q)
		
		'Delete data from table recurringsettlementpayment
		q = " DELETE FROM recurringsettlementpayment" & _
			" WHERE settlementno = '" & settlementno & "'" & _
			"       AND fcba = '" & Session("fcBA") & "'"

		'Response.Write q
		connStr.execute(q)
		
		'Delete data from table recurringsettlement
		q = " DELETE FROM recurringsettlement" & _
			" WHERE settlementno = '" & settlementno & "'" & _
			"       AND fcba = '" & Session("fcBA") & "'"

		'Response.Write q
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
		amt_cash = 0
		amt_sett = 0
		for ii = 1 to cdbl(maxcounter)
			if "" & request("cashamount"&ii) <> "" then
				amt_cash = amt_cash + Cdbl(To_Nol(request("cashamount"&ii)))
			end if
		next
		
		for n=1 to (cdbl(request("no"))-1)
			if "" & request("amount_"&TRIM(n))<>"" then
				amt_sett = amt_sett + Cdbl(To_Nol(request("amount_"&TRIM(n))))
			end if
		next
		
		if amt_cash <> amt_sett then		
			CompleteStringError = " Angka Settlement dan Cash tidak sama, Silahkan koreksi kembali data yang dimasukkan!"
		else
		
			if Session("CurP_Costbook") <> Year(rundate) & Lpad(Month(rundate), "0", "2") then
				CompleteStringError = " Periode rundate tidak sesuai dengan costbook!"
			else
		
				Set rsPO = Server.CreateObject("ADODB.Recordset")
				rsPO.Open " SELECT settlementno FROM recurringsettlement WHERE settlementno = '" & settlementno & "' AND fcba = '" & Session("fcBA") & "'", connstr
				IF rsPO.eof then	    
					q = " INSERT INTO recurringsettlement ( settlementno" & _
						"                                   , documentno" & _
						"                                   , accountingyear" & _
						"                                   , startingdate" & _
						"                                   , finishdate" & _
						"                                   , rundate" & _
						"                                   , numberofrun" & _
						"                                   , amount" & _
						"                                   , fcba" & _
						"                                   , prepaymentjob" & _
						"                                   , temporarycashjob" & _
						"                                   , status" & _
						"                                   , fcentry" & _
						"                                   , fcedit" & _
						"                                   , fcip" & _
						"                                   , lastupdate" & _
						"                                   , lasttime) " & _
						" VALUES ( '" & settlementno & "'" & _
						"          , '" & documentno & "'" & _
						"          , '" & accountingyear & "'" & _
						"          , TO_DATE('" & startingdate & "', 'mm/dd/yyyy')" & _
						"          , TO_DATE('" & finishdate & "', 'mm/dd/yyyy')" & _
						"          , TO_DATE('" & rundate & "', 'mm/dd/yyyy')" & _
						"          , '" & numberofrun & "'" & _
						"          , '" & amount & "'" & _
						"          , '" & Session("fcBA") & "'" & _
						"          , '" & prepaymentjob & "'" & _
						"          , '" & temporarycashjob & "'" & _
						"          , '" & status & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _					
						"          , SYSDATE" & _
						"          , TO_CHAR(SYSDATE,'HH24:MI'))"
					'response.Write(q)
					connStr.execute(q)
				End if
				rsPO.Close
				
				for n=1 to (cdbl(request("no"))-1) 
					if "" & request("amount_"&TRIM(n))<>"" then
						q = " INSERT INTO recurringsettlementdtl ( settlementno" & _
							"                                      , accountingyear" & _
							"                                      , debetcredit" & _
							"                                      , jobcode" & _
							"                                      , targettype" & _
							"                                      , amount_accrued" & _
							"                                      , amount_settlement" & _
							"                                      , intervalmonth" & _
							"                                      , amountmonth" & _
							"                                      , amountendinstallment" & _
							"                                      , fcba" & _
							"                                      , status" & _
							"                                      , fcentry" & _
							"                                      , fcedit" & _
							"                                      , fcip" & _
							"                                      , lastupdate" & _
							"                                      , lasttime) " & _
							" VALUES ( '" & settlementno & "'" & _
							"          , '" & accountingyear & "'" & _
							"          , '" & request("debetcredit_"&TRIM(n)) & "'" & _
							"          , '" & request("jobcode_"&TRIM(n)) & "'" & _
							"          , '" & request("targettype_"&TRIM(n)) & "'" & _
							"          , REPLACE('" & request("amt_accrued_"&TRIM(n)) & "', ',', '')" & _
							"          , REPLACE('" & request("amount_"&TRIM(n)) & "', ',', '')" & _
							"          , '" & numberofrun & "'" & _
							"          , '0'" & _
							"          , '0'" & _
							"          , '" & Session("fcBA") & "'" & _
							"          , 'Planned'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _					
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE,'HH24:MI'))"
						'response.Write(q)
						connStr.execute(q)
					end if
				next
				
				q = " DELETE recurringsettlementpayment" & _
				    " WHERE settlementno = '" & settlementno & "'" & _
					"       AND fcba = '" & Session("fcBA") & "'"
				connStr.execute(q)
				
				for ii = 1 to cdbl(maxcounter)
					if "" & request("cashamount"&ii) <> "" then
						q = " INSERT INTO recurringsettlementpayment ( settlementno" & _
							"                                          , accountingyear" & _
							"                                          , debetcredit" & _
							"                                          , jobcode" & _
							"                                          , amount" & _
							"                                          , voucherno" & _
							"                                          , fcba" & _
							"                                          , status" & _
							"                                          , fcentry" & _
							"                                          , fcedit" & _
							"                                          , fcip" & _
							"                                          , lastupdate" & _
							"                                          , lasttime) " & _
							" VALUES ( '" & settlementno & "'" & _
							"          , '" & accountingyear & "'" & _
							"          , 'Cr'" & _
							"          , '" & request("cashjob"&ii) & "'" & _
							"          , REPLACE('" & request("cashamount"&ii) & "', ',', '')" & _
							"          , NULL" & _
							"          , '" & Session("fcBA") & "'" & _
							"          , 'OPEN'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _					
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE,'HH24:MI'))"
						'response.Write(q)
						connStr.execute(q)
					end if
				next
				
			end if '-- Cek Tanggal
		end if '-- Cek amount settlement vs cash

		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		connstr.BeginTrans
		q = " UPDATE recurringsettlementdtl" & _
		    " SET amount = REPLACE('" & amount_edit & "',',','')" & _
		    "     , fcedit = '" & Session("CodeName") & "'" & _
		    "     , lastupdate = SYSDATE" & _
		    "     , lasttime = TO_CHAR(SYSDATE,'HH24:MI')" & _
			" WHERE fcba = '" & Session("fcBA") & "'" & _
			"       AND settlementno = '" & settlementno & "'" & _
			"       AND jobcode = '" & jobcode_edit & "'" & _
			"       AND targettype = '" & targettype_edit & "'"
		response.Write(q)
		connStr.execute(q)
		if err.number=0 then
			connStr.CommitTrans
		End if
		cmdView	= "True"
		Response.Redirect "/recurringsettlement.asp?settlementno=" & settlementno & "&cmdView=True&SudahClose=Y"
	elseif Request("txtJenis") = "RePlanned" THEN
		q = " BEGIN recurringsettlement_api.NextState('" & Session("fcBA") & "', '" & settlementno & "', 'RePlan' ); END; "
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis") = "Confirmed" THEN
		q = " BEGIN recurringsettlement_api.NextState('" & Session("fcBA") & "', '" & settlementno & "', 'Confirm' ); END; "
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis") = "Cancelled" THEN
		q = " BEGIN recurringsettlement_api.NextState('" & Session("fcBA") & "', '" & settlementno & "', 'Cancel' ); END; "
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis") = "Executed" THEN
		q = " BEGIN recurringsettlement_api.validatesettlement('" & Session("fcBA") & "', '" & settlementno & "', '" & Session("CodeName") & "'); END; "
		connStr.execute(q)
		
		q = " BEGIN recurringsettlement_api.NextState('" & Session("fcBA") & "', '" & settlementno & "', 'Execute' ); END; "
		connStr.execute(q)
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
oSql = " select *" & _
       " from recurringsettlement" & _
       " where fcba = '" & Session("fcba") & "'" & _
       "       and settlementno = '" & settlementno & "'"

oRs.Open oSql, connstr
if Not oRs.Eof then
	SudahClose = "Y"
	documentno = oRs("documentno")
	accountingyear = oRs("accountingyear")
	startingdate = oRs("startingdate")
	finishdate = oRs("finishdate")
	rundate = oRs("rundate")
	numberofrun = oRs("numberofrun")
	amount = oRs("amount")
	prepaymentjob = oRs("prepaymentjob")
	temporarycashjob = oRs("temporarycashjob")
	status = oRs("status")
End if
oRs.Close
Set oRs = Nothing


If Len(Trim(CompleteStringError)) > 0 Then
	response.Write  "<div class = ""error"">" & CompleteStringError & "</div>"
End If

val_IsTrue = "Y"

if SudahClose = "Y" then
	val_IsTrue = "N"
else
	val_IsTrue = "Y"
end if
 %>
<form name="frmEntry" method="post" action="recurringsettlement.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Settlement No</td>
			<td align="left">
				<input type="text" maxlength = "35" name="settlementno" size="30" value="<%=settlementno%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Document No</td>
			<td align="left">
				<select id="documentno" name="documentno" style="width:50%;" class="chosen-select" onchange="javascript:getMessage();">
					<option value=""></option>
				<%
					if SudahClose <> "Y" then
					qaccrualjob = " SELECT documentno FROM (" & _
                                  "     SELECT documentno FROM recurring WHERE fcba = '" & Session("fcBA") & "' AND isusesettlement = 'TRUE' AND status = 'Executed'" & _
                                  "     MINUS" & _
                                  "     SELECT documentno FROM recurringsettlement WHERE fcba = '" & Session("fcBA") & "'" & _
                                  " ) ORDER BY documentno"
					else
					qaccrualjob = " SELECT documentno FROM recurringsettlement WHERE fcba = '" & Session("fcBA") & "'" & _
								  " ORDER BY documentno"
					end if
					Set rsaccrualjob = Server.CreateObject("ADODB.Recordset")
					rsaccrualjob.Open qaccrualjob, connstr
					IF NOT rsaccrualjob.EOF THEN
						WHILE NOT rsaccrualjob.EOF
				%>
				<option value="<%=rsaccrualjob("documentno")%>" <%IF documentno = rsaccrualjob("documentno") THEN%> selected <%END IF%>>
					<%=rsaccrualjob("documentno")%>
				</option>
				<%
							rsaccrualjob.MOVENEXT
						WEND
					END IF
				%>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Accounting Year</td>
			<td align="left">
				<input type="text" maxlength = "100" id="accountingyear" name="accountingyear" size="30" value="<%=accountingyear%>" readonly="true" style="background:#ccc;">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Starting Date</td>
			<td align="left">
				<input id="startingdate" class="datepicker" type="text" maxlength="10" name="startingdate" size="10" value="<%=startingdate%>" onchange="getnumberofruns();">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Finished Date</td>
			<td align="left">
				<input type="text" maxlength="10" id="finishdate" name="finishdate" size="10" value="<%=finishdate%>" readonly="true" style="background:#ccc;" >
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Run Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" id="rundate" name="rundate" size="10" value="<%=rundate%>" readonly="true" style="background:#ccc;">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Number of Run</td>
			<td align="left">
				<input id="numberofrun" type="number" name="numberofrun" size="30" value="<%=numberofrun%>" readonly="true">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Amount</td>
			<td align="left">
				<input disabled style="background:#ccc;" type="text" name="amount" size="30" value="<%=amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Prepayment Job</td>
			<td align="left">
				<select name="prepaymentjob" class="chosen-select" style="width:50%;">
				<%
					qaccrualjob = " SELECT fccode, fcname FROM job WHERE fcba = '" & Session("fcBA") & "' AND jobgroupcode = '1140000' ORDER BY fccode"
					Set rsaccrualjob = Server.CreateObject("ADODB.Recordset")
					rsaccrualjob.Open qaccrualjob, connstr
					IF NOT rsaccrualjob.EOF THEN
						WHILE NOT rsaccrualjob.EOF
				%>
				<option value="<%=rsaccrualjob("fccode")%>" <%IF prepaymentjob = rsaccrualjob("fccode") THEN%> selected <%END IF%>>
					<%=rsaccrualjob("fccode")%> : <%=rsaccrualjob("fcname")%>
				</option>
				<%
							rsaccrualjob.MOVENEXT
						WEND
					END IF
				%>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Status</td>
			<td align="left">
				<input disabled type="text" maxlength="10" name="status" size="30" value="<%=status%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">
				<%If status = "Planned" Then%>
				<a href="recurringsettlement.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&settlementno=<%=settlementno%>&txtJenis=Header">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=documentno%>" border="0" onClick="return confirm('Are you sure to delete');">
				</a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End if%>
				<% if status = "Planned" then %>
					&nbsp;&nbsp;
					<a href="recurringsettlement.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&settlementno=<%=settlementno%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to confirm this data :  <%=settlementno%> ?');">
						Confirm
					</a>
					&nbsp;&nbsp;
				<% elseif status = "Confirmed" then %>
					&nbsp;&nbsp;
					<a href="recurringsettlement.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&settlementno=<%=settlementno%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=settlementno%> ?');">
						Re-Plan
					</a>
					&nbsp;&nbsp;
					<a href="recurringsettlement.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&settlementno=<%=settlementno%>&txtJenis=Executed" onClick="return confirm('Are you sure to Execute this data, Please check your detail data before once Executed!, Settlement No:  <%=settlementno%> ?');">
						Execute
					</a>
					&nbsp;&nbsp;
					<a href="recurringsettlement.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&settlementno=<%=settlementno%>&state=<%=state%>&txtJenis=Cancelled" onClick="return confirm('Are you sure to Cancel this data : <%=settlementno%> ?');">
						Cancel
					</a>
					&nbsp;&nbsp;
		<% end if%>
			</td>
		</tr>
	</table>  
</form>
<%
If cmdView = "True" and "" & settlementno <> "" Then
	Dim rswipfinal, intRecordsPerPage, cColor
	set rswipfinal = server.CreateObject("Adodb.recordset")
	arr_pt_eof = true
	rswipfinal_eof = true
	If SudahClose <> "Y" Then
		q	= " SELECT red.debetcredit" & _
			"        , red.jobcode" & _
			"        , red.targettype" & _
			"        , red.amount created" & _
			"        , rtd.amount accrued" & _
			" FROM recurring re" & _
			" INNER JOIN recurringdtl red ON (re.fcba = red.fcba AND re.documentno = red.documentno)" & _
			" LEFT JOIN ( SELECT rtd.fcba, rtd.jobcode, rtd.targettype, SUM(rtd.amount) amount" & _
			"             FROM recurringtransactiondtl rtd" & _
			"             INNER JOIN recurringtransaction rt ON (rt.fcba = rtd.fcba AND rt.transactionno = rtd.transactionno)" & _
			"             WHERE rtd.fcba = '" & Session("fcba") & "'" & _
			"                   AND rt.documentno = '" & documentno & "'" & _
			"                   AND rtd.debetkredit = 'Db'" & _
			"                   GROUP BY rtd.fcba, rtd.jobcode, rtd.targettype) rtd" & _
			" ON (red.fcba = rtd.fcba AND red.jobcode = rtd.jobcode AND red.targettype = rtd.targettype)" & _
			" WHERE red.fcba = '" & Session("fcba") & "'" & _
			"       AND red.documentno = '" & documentno & "'" & _
			" ORDER BY red.jobcode, red.targettype"

	Else	 
		q = " SELECT red.debetcredit" & _
			"        , red.jobcode" & _
			"        , red.targettype" & _
			"        , red.amount created" & _
			"        , rtd.amount accrued" & _
			"        , rsd.amount settlemented" & _
			" FROM recurring re" & _
			" INNER JOIN recurringdtl red ON (re.fcba = red.fcba AND re.documentno = red.documentno)" & _
			" LEFT JOIN ( SELECT rtd.fcba, rtd.jobcode, rtd.targettype, SUM(rtd.amount) amount" & _
			"             FROM recurringtransactiondtl rtd" & _
			"             INNER JOIN recurringtransaction rt ON (rt.fcba = rtd.fcba AND rt.transactionno = rtd.transactionno)" & _
			"             WHERE rtd.fcba = '" & Session("fcba") & "'" & _
			"                   AND rt.documentno = '" & documentno & "'" & _
			"                   AND rtd.debetkredit = 'Db'" & _
			"                   GROUP BY rtd.fcba, rtd.jobcode, rtd.targettype) rtd" & _
			" ON (red.fcba = rtd.fcba AND red.jobcode = rtd.jobcode AND red.targettype = rtd.targettype)" & _
			" LEFT JOIN ( SELECT rsd.fcba, rs.documentno, rsd.jobcode, rsd.targettype, SUM(rsd.amount_settlement) amount" & _
			"             FROM recurringsettlementdtl rsd" & _
			"             INNER JOIN recurringsettlement rs ON (rs.fcba = rsd.fcba AND rs.settlementno = rsd.settlementno)" & _
			"             WHERE rsd.fcba = '" & Session("fcba") & "'" & _
			"                   AND rs.settlementno = '" & settlementno & "'" & _
			"                   AND rsd.debetcredit = 'Db'" & _
			"                   GROUP BY rsd.fcba, rs.documentno, rsd.jobcode, rsd.targettype) rsd" & _
			" ON (red.fcba = rsd.fcba AND red.documentno = rsd.documentno AND red.jobcode = rsd.jobcode AND red.targettype = rsd.targettype)" & _
			" WHERE red.fcba = '" & Session("fcba") & "'" & _
			"       AND red.documentno = '" & documentno & "'" & _
			" ORDER BY red.jobcode, red.targettype"
			
		q2 = " SELECT rsp.jobcode, jj.fcname jobname, rsp.amount" & _
		     " FROM recurringsettlementpayment rsp" & _
		     " LEFT JOIN job jj ON (rsp.fcba = jj.fcba AND rsp.jobcode = jj.fccode)" & _
		     " WHERE rsp.fcba = '" & Session("fcba") & "'" & _
			 "       AND rsp.settlementno = '" & settlementno & "'"
			 
		set rspayment = server.CreateObject("Adodb.recordset")
		rspayment.Open q2, connstr

		if not rspayment.eof then
			arr_pt = rspayment.getrows()
			arr_pt_eof = false
		end if
   End if

   If Request("txtJenis")="Edit" Then
		q = " SELECT red.debetcredit" & _
			"        , red.jobcode" & _
			"        , red.targettype" & _
			"        , red.amount created" & _
			"        , rtd.amount accrued" & _
			"        , rsd.amount settlemented" & _
			" FROM recurring re" & _
			" INNER JOIN recurringdtl red ON (re.fcba = red.fcba AND re.documentno = red.documentno)" & _
			" LEFT JOIN ( SELECT rtd.fcba, rtd.jobcode, rtd.targettype, SUM(rtd.amount) amount" & _
			"             FROM recurringtransactiondtl rtd" & _
			"             INNER JOIN recurringtransaction rt ON (rt.fcba = rtd.fcba AND rt.transactionno = rtd.transactionno)" & _
			"             WHERE rtd.fcba = '" & Session("fcba") & "'" & _
			"                   AND rt.documentno = '" & documentno & "'" & _
			"                   AND rtd.debetkredit = 'Db'" & _
			"                   GROUP BY rtd.fcba, rtd.jobcode, rtd.targettype) rtd" & _
			" ON (red.fcba = rtd.fcba AND red.jobcode = rtd.jobcode AND red.targettype = rtd.targettype)" & _
			" LEFT JOIN ( SELECT rsd.fcba, rs.documentno, rsd.jobcode, rsd.targettype, SUM(rsd.amount_settlement) amount" & _
			"             FROM recurringsettlementdtl rsd" & _
			"             INNER JOIN recurringsettlement rs ON (rs.fcba = rsd.fcba AND rs.settlementno = rsd.settlementno)" & _
			"             WHERE rsd.fcba = '" & Session("fcba") & "'" & _
			"                   AND rs.settlementno = '" & settlementno & "'" & _
			"                   AND rsd.debetcredit = 'Db'" & _
			"                   GROUP BY rsd.fcba, rs.documentno, rsd.jobcode, rsd.targettype) rsd" & _
			" ON (red.fcba = rsd.fcba AND red.documentno = rsd.documentno AND red.jobcode = rsd.jobcode AND red.targettype = rsd.targettype)" & _
			" WHERE red.fcba = '" & Session("fcba") & "'" & _
			"       AND red.documentno = '" & documentno & "'" & _
			"       AND red.jobcode = '" & jobcode_edit & "'" & _
			"       AND red.targettype = '" & targettype_edit & "'" & _
			" ORDER BY red.jobcode, red.targettype"

	End If
			  
   'response.Write q	
   rswipfinal.CursorType = 3
   rswipfinal.CursorLocation = 3
   rswipfinal.Open q, connstr
   
%>
   <form action="recurringsettlement.asp" method="post" name="eForm">
<ul class="tabs">
    <li><a href="#view1">Detail Settlement</a></li>
    <li><a href="#view2">Payment Terms</a></li>
</ul>
<div class="tabcontents">
    <div id="view1">
<table class="data" border="1" width="100%" >
   <%if SudahClose <> "Y" Then%>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
		   <td align="center" rowspan="2">Debet Credit</td>
		   <td align="center" rowspan="2">Jobcode</td>
		   <td align="center" rowspan="2">Target Type</td>
		   <td align="center" colspan="3">Amount</td>
   	   </tr>
	   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
		   <td align="center">Created</td>
		   <td align="center">Accrued</td>
		   <td align="center">Settlement</td>
   	   </tr>
   <%Else%>
		<%if Request("txtJenis")="Edit" Then%>
		   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
			   <td align="center" rowspan="2">Debet Credit</td>
			   <td align="center" rowspan="2">Jobcode</td>
			   <td align="center" rowspan="2">Target Type</td>
			   <td align="center" colspan="3">Amount</td>
		   </tr>
		   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
			   <td align="center">Created</td>
			   <td align="center">Accrued</td>
			   <td align="center">Settlement</td>
		   </tr>
		<%else%>
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
			   <td align="center" rowspan="2" colspan="2">&nbsp;</td>
			   <td align="center" rowspan="2">Debet Credit</td>
			   <td align="center" rowspan="2">Jobcode</td>
			   <td align="center" rowspan="2">Target Type</td>
			   <td align="center" colspan="3">Amount</td>
		   </tr>
		   <tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
			   <td align="center">Created</td>
			   <td align="center">Accrued</td>
			   <td align="center">Settlement</td>
		   </tr>
		<%End if%>

   <%End if
   if Not rswipfinal.eof then
      nCountofField=rswipfinal.fields.count-1
	  no = 1
	  rswipfinal_eof = false
	  While Not rswipfinal.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
			<tr bgcolor="<%=cColor%>">
				<%IF SudahClose = "Y" then%>

					<%if Request("txtJenis")="Edit" Then%>
						<td align="left" title="<%=rswipfinal("debetcredit").Name%>" valign="top"><%=rswipfinal("debetcredit")%></td>
						<td align="left" title="<%=rswipfinal("jobcode").Name%>" valign="top"><%=rswipfinal("jobcode")%></td>
						<td align="left" title="<%=rswipfinal("targettype").Name%>" valign="top"><%=rswipfinal("targettype")%></td>
						<td align="right" title="<%=rswipfinal("created").Name%>" valign="top"><%=FormatNumber(rswipfinal("created"), 0)%></td>
						<td align="right" title="<%=rswipfinal("accrued").Name%>" valign="top"><%=FormatNumber(rswipfinal("accrued"), 0)%></td>
						<td align="left" width="2">
							<input type="text" name="amount_edit" width="2" value="<%=rswipfinal("settlemented")%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
						</td>
						<input type="hidden" name="jobcode_edit" value="<%=rswipfinal("jobcode")%>">
						<input type="hidden" name="targettype_edit" value="<%=rswipfinal("targettype")%>">
					<%Else%>
				
						<%If val_IsTrue = "Y" Then%>
						<td valign="middle" width="10">
							<a href="recurringsettlement.asp?settlementno=<%=settlementno%>&jobcode_edit=<%=rswipfinal("jobcode")%>&targettype_edit=<%=rswipfinal("targettype")%>&amount_edit=<%=rswipfinal("settlemented")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rswipfinal(1)%>" border="0"></a>
						</td>
						<td valign="middle" width="10">
							<a href="recurringsettlement.asp?txtJenis=Details" onClick="return confirm('Are you sure to delete');">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rswipfinal(1)%>" border="0"></a>
						</td>
						<%Else%>
						<td colspan="2" align="center">
							<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
						</td>
						<%End If%>
						
						<td align="left" title="<%=rswipfinal("debetcredit").Name%>" valign="top"><%=rswipfinal("debetcredit")%></td>
						<td align="left" title="<%=rswipfinal("jobcode").Name%>" valign="top"><%=rswipfinal("jobcode")%></td>
						<td align="left" title="<%=rswipfinal("targettype").Name%>" valign="top"><%=rswipfinal("targettype")%></td>
						<td align="right" title="<%=rswipfinal("created").Name%>" valign="top"><%=FormatNumber(Cdbl(To_Nol(rswipfinal("created"))), 0)%></td>
						<td align="right" title="<%=rswipfinal("accrued").Name%>" valign="top"><%=FormatNumber(Cdbl(To_Nol(rswipfinal("accrued"))), 0)%></td>
						<td align="right" title="<%=rswipfinal("settlemented").Name%>" valign="top"><%=FormatNumber(Cdbl(To_Nol(rswipfinal("settlemented"))), 0)%></td>
						
						<%
					      t_created = t_created +cdbl(To_Nol(rswipfinal("created")))
						  t_accrued = t_accrued +cdbl(To_Nol(rswipfinal("accrued")))
						  t_settlemented = t_settlemented +cdbl(To_Nol(rswipfinal("settlemented")))
						%>
						<input type="hidden" name="amount_<%=no%>" value="<%=Cdbl(To_Nol(rswipfinal("settlemented")))%>">
					
					<%End If%>
					
				<%End if%>
				
				<%if SudahClose <> "Y" Then%>
					<td align="left" title="Debet Credit" valign="top"><%=rswipfinal("debetcredit")%></td>
					<td align="left" title="Job Code" valign="top"><%=rswipfinal("jobcode")%></td>
					<td align="left" title="Target Type" valign="top"><%=rswipfinal("targettype")%></td>
					<td align="Right" title="Created" valign="top"><%=FormatNumber(rswipfinal("created"), 0)%></td>
					<td align="Right" title="Accrued" valign="top"><%=FormatNumber(rswipfinal("accrued"), 0)%></td>
					<td align="left" width="2">
						<input type="text" name="amount_<%=no%>" width="2" value="<%=request("amount_"&TRIM(no))%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
					</td>
					<input type="hidden" name="debetcredit_<%=no%>" value="<%=rswipfinal("debetcredit")%>">
					<input type="hidden" name="jobcode_<%=no%>" value="<%=rswipfinal("jobcode")%>">
					<input type="hidden" name="targettype_<%=no%>" value="<%=rswipfinal("targettype")%>">
					<input type="hidden" name="amt_accrued_<%=no%>" value="<%=rswipfinal("accrued")%>">
				<%End if%>
		  </tr>
		  <%
		  rswipfinal.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no = no + 1
		  if SudahClose = "Y" then
			  
		  end if
	  Wend
		if SudahClose = "Y" then
			if Request("txtJenis")<>"Edit" Then
				response.write "<tr bgcolor='#F4F4A0'>"
				response.write "	<td align='center' colspan='5'>Total</td>"
				response.write "	<td align='right'>"&FormatNumber(t_created,0)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(t_accrued,0)&"</td>"
				response.write "	<td align='right'>"&FormatNumber(t_settlemented,0)&"</td>"
				response.write "</tr>"
			end if
		end if
   else
      response.Write("<tr><td colspan='6' align='left'>Data is empty</td></tr>")	  
   end if 
   rswipfinal.Close
   set rswipfinal = Nothing%>
	</TABLE>
    </div>
    <div id="view2">
	<%
	qaccrualjob = " SELECT fccode, fcname" & _
                  " FROM (" & _
                  "     SELECT jo.fccode, jo.fcname" & _
                  "     FROM job jo" & _
                  "     INNER JOIN bank ba ON (ba.fcba = jo.fcba AND ba.controljob = jo.fccode)" & _
                  "     WHERE ba.fcba = '" & Session("fcBA") & "' AND jobgroupcode IN ('1111000', '1112000')" & _
                  "     UNION ALL" & _
                  "     SELECT fccode, fcname FROM job" & _
                  "     WHERE fcba = '" & Session("fcBA") & "' AND jobgroupcode = '1120000' AND fccode <> '1120400'" & _
                  " ) ORDER BY fccode"
	Set rsaccrualjob = Server.CreateObject("ADODB.Recordset")
	rsaccrualjob.Open qaccrualjob, connstr
	arr_job_eof = true
	if not rsaccrualjob.eof then
		arr_job = rsaccrualjob.getrows()
		arr_job_eof = false
	end if
	%>
		<div id="addpayment">
			<%if not arr_pt_eof then
				indx = 1
				maxcounter = 0
				
				For dtlxx = 0 To UBound(arr_pt, 2)
					maxcounter = maxcounter +1
				next
				For dtlxx = 0 To UBound(arr_pt, 2)
				%>
					<%if indx > 1 then%>
					<div id="option<%=indx%>">
					<%end if%>
					<div class="section group">
						<div class="col span_5_of_12">
						<select name="cashjob<%=indx%>" class="chosen-select" style="width:100%;">
							<option value=""></option>
						<%if not arrjob_eof then%>
							<%For xx = 0 To UBound(arr_job, 2)%>
							<option value="<%=arr_job(0, xx)%>" <% If arr_pt(0, dtlxx) = arr_job(0, xx) then%> Selected <%End if%>><%=arr_job(0, xx) & " : " & arr_job(1, xx)%></option> 
							<%Next%>
						<%End if%>
						</select>
					</div>
					<div class="col span_5_of_12">
						<input  style="width:100%;"type="text" name="cashamount<%=indx%>" value="<%=arr_pt(2, dtlxx)%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
					</div>
					<div class="col span_2_of_7">
						<%if indx = 1 then%>
							<%If val_IsTrue = "Y" Then%>
								<a class="button_menu_top" onclick="addpayment();"> + </a>
							<%end if%>
						<%else%>
							<%If val_IsTrue = "Y" Then%>
								<div class="close-button" onclick="delete_element('option<%=indx%>');"></div>
							<%end if%>
						<%end if%>
						<input Type="hidden" value="<%=maxcounter%>" id="maxcounter" name="maxcounter">
					</div>
					<%if indx = 1 then%>
					</div>
					<%end if%>
					<%if indx > 1 then%>
					</div>
					<%end if%>
					<%indx = indx +1%>
				<%next%>
			<%else%>
			<div class="option1">
				<div class="section group">
					<div class="col span_5_of_12">
						<select name="cashjob1" class="chosen-select" style="width:100%;">
							<option value=""></option>
							<%if not arr_job_eof then%>
							<%for ji=0 to ubound(arr_job,2)%>
							<option value="<%=arr_job(0, ji)%>" <%IF request("cashjob1") = arr_job(0, ji) THEN%> selected <%END IF%>>
								<%=arr_job(0, ji)%> : <%=arr_job(1, ji)%>
							</option>
							<%next%>
							<%end if%>
						<select>
					</div>
					<div class="col span_5_of_12">
						<input  style="width:100%;"type="text" name="cashamount1" value="<%=request("cashamount1")%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
					</div>
					<div class="col span_2_of_12">
						<a class="button_menu_top" onclick="addpayment();"> + </a>
						<input Type="hidden" value="<%=maxcounter%>" id="maxcounter" name="maxcounter">
					</div>
				</div>
			</div>
			<%end if%>
		</div>
    </div>
</div>
<table>
	<tr>
			<td colspan="11" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="settlementno" value="<%=settlementno%>">
			<input type="hidden" name="documentno" value="<%=documentno%>">
			<input type="hidden" name="accountingyear" value="<%=accountingyear%>">
			<input type="hidden" name="startingdate" value="<%=startingdate%>">
			<input type="hidden" name="finishdate" value="<%=finishdate%>">
			<input type="hidden" name="rundate" value="<%=rundate%>">
			<input type="hidden" name="numberofrun" value="<%=numberofrun%>">
			<input type="hidden" name="amount" value="<%=amount%>">
			<input type="hidden" name="prepaymentjob" value="<%=prepaymentjob%>">
			<input type="hidden" name="temporarycashjob" value="<%=temporarycashjob%>">
			<input type="hidden" name="status" value="<%=status%>">
			<input type="hidden" name="No" value="<%=No%>">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<%If val_IsTrue = "Y" Then%>
			<%IF Request("txtJenis")="Edit" then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit"  name="Add" value="Update">
			<%Else%>
				<input type="hidden" name="txtJenis" value="New">
				<input Type="Submit"  name="Add" value="Submit">
			<%End if%>
			<%End if%>
			<input Type="Reset"  name="Reset" value="Reset">
			<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
	</table>
</div>
</form>
<%End if%>
<!---------------akhir!--------------->
</div>
</div>
<!--#include file="bottom.jsp"-->	
	<script src="plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
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
		
	</script>
	
	<script src="/js/jquery.inputmask.bundle.min.js"></script>
	<script src="/js/inputmask.binding.js"></script>
	<script src="plugins/tabcontent/tabcontent.js" type="text/javascript"></script>
	
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
		// var PostStr = "variable=value&variable2=value2";
		var PostStr = "documentno="+document.getElementById("documentno").value;
		
		// use the generic function to make the request
		doAJAXCall('ajax/recurringdetail.asp', 'POST', '' + PostStr + '', showMessageResponse);
	}

	// The function for handling the response from the server
	var showMessageResponse = function (oXML) { 
		
		// get the response text, into a variable
		var response = JSON.parse(oXML.responseText);
		//alert(response.finishdate);
		
		
		document.getElementById("finishdate").value = response.finishdate;
		document.getElementById("accountingyear").value = response.accountingyear;
		
		
		//document.getElementById("finishdate").value = JSON.stringify(response.finishdate);
	};

	
	function getnumberofruns(){
	
	var start_date = new Date(document.getElementById("startingdate").value);
	var run_date = document.getElementById("rundate")
	var numberofruns = document.getElementById("numberofrun");
	var end_date = new Date(document.getElementById("finishdate").value);
	total_months = (end_date.getFullYear() - start_date.getFullYear())*12 + (end_date.getMonth() - start_date.getMonth());
	if (total_months < 0){
		numberofruns.value = 0;
	} else {
		numberofruns.value = total_months +1;
	}
	
	
	run_date.value = document.getElementById("startingdate").value;
	
	//alert(numberofruns);
}
	</script>
	
	<script type="text/javascript">
		function saveValue(e){
            var id = e.id;  // get the sender's id to save it . 
            var val = e.value; // get the value. 
            sessionStorage.setItem(id, val);// Every time user writing something, the localStorage's value will override . 
        }
		
		function addpayment(){
			var counter = document.getElementById("maxcounter");
			var x = 1;
			num = x + parseInt(counter.value);
			
			option = "'option"+num+"'";
			$("#addpayment").append('<div id="option'+num+'">'+
									  '<div class="section group">'+
										' <div class="col span_5_of_12">'+
											' <select id="cashjob'+num+'" name="cashjob'+num+'" class="chosen-select" style="width:100%;" onchange="saveValue(this);">'+
												'<option value=""></option>'+
												<%if not arr_job_eof then%>
													<%For xx = LBound(arr_job, 2) To UBound(arr_job, 2)%>
													'<option value="<%=arr_job(0, xx)%>" <% If request("") = arr_job(0, xx) then%> Selected <%End if%>><%=arr_job(0, xx) & " : " & arr_job(1, xx)%></option>' +
													<%Next%>
												<%End if%>
											' </select>'+
										' </div>'+
										' <div class="col span_5_of_12">'+
											' <input style="width:100%;"type="text" id="cashamount'+num+'" name="cashamount'+num+'" value="" data-inputmask="&#39;alias&#39;: &#39;numeric&#39;, &#39;groupSeparator&#39;: &#39;,&#39;, &#39;autoGroup&#39;: true, &#39;digits&#39;: 0, &#39;digitsOptional&#39;: false, &#39;placeholder&#39;: &#39;0&#39;" style="text-align: right;" onkeyup="saveValue(this);">' +
										' </div>'+
										' <div class="col span_2_of_12">'+
											' <div class="close-button" onclick="delete_element(' + option + ');"></div>'+
										' </div>'+
									  ' </div>'+
									' </div>');
			counter.value = num;
			
			var script = document.createElement('script');
			script.src = '/js/inputmask.binding.js';
			document.body.appendChild(script);
		
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
		
		function delete_element(id){
			Element.prototype.remove = function() {
				this.parentElement.removeChild(this);
			}
			NodeList.prototype.remove = HTMLCollection.prototype.remove = function() {
				for(var i = this.length - 1; i >= 0; i--) {
					if(this[i] && this[i].parentElement) {
						this[i].parentElement.removeChild(this[i]);
					}
				}
			}
			document.getElementById(id).remove();
		}
	</script>
</body>
</html>
