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
		Response.Redirect "/userdefinition.asp?settlementno=" & settlementno & "&cmdView=True&SudahClose=Y"
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
<form name="frmEntry" method="post" action="userdefinition.asp">
	<table class="data" border="0" width="100%" >
    <%if Request("txtJenis") <> "Edit" then%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                USER AND PROFILE</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Password</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCPASSWORD" title="PASSWORD Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Full Name</font></td>
            <td align="left"><input type="text" size="25" maxlength="25" name="FULLNAME" title="FULLNAME Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Title</font></td>
            <td align="left">
			<select name="TITLE">
			
			<%
			dim rstitle
			q = "select fcCode,description from usertitle Order by FCCODE"
			Set rstitle = Server.CreateObject("ADODB.Recordset")
			rstitle.CursorType = 3
			rstitle.CursorLocation = 3
			rstitle.Open q, connstr
			if not rstitle.eof then
			%>
			<OPTION value="">-- Title --</option>
			<%
			 Do While Not rstitle.eof
			%>
			<option value="<%=rstitle("fcCode")%>"><%=rstitle("fcCode")%> : <%=rstitle("description")%></option>
			<%rstitle.movenext
			 loop
			End if
			rstitle.Close
			set rstitle = Nothing
			q=""
			%>
			</SELECT><font color="#FF0000">*</font>
			
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Department</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="DEPARTMENT"  title="DEPARTMENT Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Access Rights</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ACCESSRIGHTS" title="ACCESSRIGHTS Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Active</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ISACTIVE" title="ACTIVE Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">First Address</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="FIRSTADDRESS" title="ACTIVE Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Valid From</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="VALID_FROM" title="VALID FROM Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Valid Until</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="VALID_UNTIL" title="VALID UNTIL Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>		  
		  
   <%Else%>
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE 
                USER AND PROFILE</strong></font></div></td>
          </TR>
          <Tr>
            <td align="left"><font face="Verdana" size="1">Code</font></td>
            <td align="left"><input type="text" size="30" maxlength="30"  name="FCCODE" value="<%=rsSetup("FCCODE")%>" readonly="True" title="CODE Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Password</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="FCPASSWORD" title="PASSWORD Character a s/d z or 1 s/d 9"> 
              <font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Full Name</font></td>
            <td align="left"><input type="text" size="25" maxlength="25" name="FULLNAME" value="<%=rsSetup("FULLNAME")%>" title="FULLNAME Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Title</font></td>
            <td align="left">
			<select name="TITLE">
			
			<%
			
			q = "select fcCode,description from usertitle Order by FCCODE"
			Set rstitle = Server.CreateObject("ADODB.Recordset")
			rstitle.CursorType = 3
			rstitle.CursorLocation = 3
			rstitle.Open q, connstr
			if not rstitle.eof then
			%>
			<OPTION value="">-- Title --</option>
			<%
			 Do While Not rstitle.eof
			%>
			<option value="<%=rstitle("fcCode")%>" <%if rsSetup("TITLE") = rstitle("fcCode") then%> Selected <%End if%>><%=rstitle("fcCode")%> : <%=rstitle("description")%></option>
			<%rstitle.movenext
			 loop
			End if
			rstitle.Close
			set rstitle = Nothing
			q=""
			%>
			</SELECT><font color="#FF0000">*</font>
			
			</td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Department</font></td>
            <td align="left"><input type="text" size="15" maxlength="15"  name="DEPARTMENT" value="<%=rsSetup("DEPARTMENT")%>" title="DEPARTMENT Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Access Rights</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ACCESSRIGHTS" value="<%=rsSetup("ACCESSRIGHTS")%>" title="ACCESSRIGHTS Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
          <Tr> 
            <td align="left"><font face="Verdana" size="1">Active</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="ISACTIVE" value="<%=rsSetup("ISACTIVE")%>"  title="ACTIVE Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">First Address</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="FIRSTADDRESS"  value="<%=rsSetup("FIRSTADDRESS")%>" title="ACTIVE Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Valid From</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="VALID_FROM" value="<%=rsSetup("VALID_FROM")%>"  title="VALID FROM Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>
		  <Tr> 
            <td align="left"><font face="Verdana" size="1">Valid Until</font></td>
            <td align="left"><input type="text" size="50" maxlength="200"  name="VALID_UNTIL" value="<%=rsSetup("VALID_UNTIL")%>"  title="VALID UNTIL Character a s/d z or 1 s/d 9">
			<font color="#FF0000">*</font></td>
          </Tr>			
   <%End IF%>
	</TABLE>  
</form>
<%

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
<form action="userdefinition.asp" method="post" name="eForm">
<ul class="tabs">
    <li><a href="#view1">User And Access Control</a></li>
    <li><a href="#view2">User And Fcba</a></li>
    <li><a href="#view3">User And Company</a></li>
    <li><a href="#view4">User And Roles</a></li>
</ul>
<div class="tabcontents">
    <div id="view1">
    	<table class="data" border="0" width="100%" >
			<%
			Dim rsGRN,intRecordsPerPages,cColors
			set rsGRN = server.CreateObject("Adodb.recordset")

				q	= " select distinct column_name, document_type from ac_module"
			   Set rsGRN = Server.CreateObject("ADODB.Recordset")
			   rsGRN.CursorType = 3
			   rsGRN.CursorLocation = 3
			   rsGRN.Open q, connstr%>
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
				<TD align="left">NO</TD>
				<TD>&nbsp;</TD>
				<Td align="left" valign="top">Document Type</Td>
			
				<% 
				  if Not rsGRN.eof then
			      nCountofField=rsGRN.fields.count-1
				  no=1
				  While Not rsGRN.EOF
			  	%>

					<td align="left" title="" valign="top"><%=rsGRN("document_type")%></td>
	
			
				<td align="left" valign="top" colspan="2"><%=NO%></td>
				<TD align="left"><%=rsGRN("document_type")%></TD>
				<TD align="left">
					<input style="float: left;" type="checkbox"  name="fullstand" value="<%=fullstand%>" >
				</TD>
			</tr>
			<%
				  rsGRN.MoveNext
				  intRecordsPerPage = intRecordsPerPage - 1
				  no=no+1
			  Wend
		   else
		      ' response.Write("Data is empty")	  
		   end if 
		   rsGRN.Close
		   set rsGRN = Nothing%>
		</table>
    </div>
    <div id="view2">

    </div>
    <div id="view3">

    </div>
    <div id="view4">

    </div>
	
</div>
</form>
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
