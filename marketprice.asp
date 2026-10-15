<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
	<title>ePlaS - Market Price </title>
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

valid_from = Request("valid_from")
productcode = Request("productcode")
customercode = Request("customercode")
rate = Request("rate")
fcba = Request("fcba")
' fcCode = Request("fcba")
fcCode = Request("fcCode")


' maxcounter = request("maxcounter")
' if "" & maxcounter = "" then maxcounter = 1


' jobcode_edit = request("jobcode_edit")
' targettype_edit = request("targettype_edit")
' amount_edit = request("amount_edit")

txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

if "" & status = "" then status = "Planned"
	
			
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")="Delete" OR Request("txtJenis") = "Confirmed" OR Request("txtJenis") = "Cancelled" OR Request("txtJenis") = "Executed" OR Request("txtJenis") = "RePlanned" Then
	if Request("txtJenis")= "Header" then
		'Delete data from table recurringsettlementdtl
		
		' 'Delete data from table recurringsettlementpayment
		' q = " DELETE FROM recurringsettlementpayment" & _
		' 	" WHERE settlementno = '" & settlementno & "'" & _
		' 	"       AND fcba = '" & Session("fcBA") & "'"

		' 'Response.Write q
		' connStr.execute(q)
		
		' 'Delete data from table recurringsettlement
		' q = " DELETE FROM recurringsettlement" & _
		' 	" WHERE settlementno = '" & settlementno & "'" & _
		' 	"       AND fcba = '" & Session("fcBA") & "'"

		' 'Response.Write q
		' connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then
			
			' For n=1 to (cdbl(request("no"))-1)
			' if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				q = " DELETE FROM marketprice" & _
					" WHERE productcode = '" & productcode & "'" & _
					"       AND customercode = '" & customercode & "'" & _
					"       AND valid_from = to_date('" & valid_from & "','mm/dd/yyyy')" & _
					"       AND fcba = '" & fcba & "'"

					' Response.Write q
					connStr.execute(q)
		' 	end if
		' next

		cmdView	= "True"

	elseif Request("txtJenis")= "Details" then
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
' response.write "No="&no
		For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then

						q =	" INSERT INTO marketprice ( productcode" & _
							"                           , customercode" & _
							"                           , rate" & _
							"                           , valid_from" & _
							"                           , fcba" & _
							"                           , fcentry" & _
							"                           , fcedit" & _
							"                           , fcip" & _
							"                           , lastupdate" & _
							"                           , lasttime) " & _
							" VALUES ( '" & productcode & "'" & _
							"          , '" & customercode & "'" & _
							"          , replace('" & rate & "',',','')" & _
							"          , TO_DATE('" & valid_from & "', 'mm/dd/yyyy')" & _
							"          , '" & request("fcba_"&TRIM(n)) & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Session("CodeName") & "'" & _
							"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _					
							"          , SYSDATE" & _
							"          , TO_CHAR(SYSDATE,'HH24:MI'))"
						' response.write q
						connStr.execute(q)
			end if
		next

		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
				' q	= " SELECT mp.productcode, " & _
				'    	  "	mp.customercode," & _
				'    	  " mp.rate," & _
				'       " mp.valid_from," & _
				'       " mp.fcba," & _
				'       " bu.fcname" & _
				' 		  " FROM marketprice mp LEFT JOIN businessunit bu ON (bu.fccode = mp.fcba) " 

				' ' response.Write q
				' Set rsVH = Server.CreateObject("ADODB.Recordset")
				' rsVH.CursorType = 3
				' rsVH.CursorLocation = 3
				' rsVH.Open q, connstr
				' if Not rsVH.eof then
				' 	productcode =  rsVH("productcode")
				' 	customercode =  rsVH("customercode")
				' 	rate =  rsVH("rate")
				' 	valid_from =  rsVH("valid_from")
				' 	fcname =  rsVH("fcname")
				' 	fcba =  rsVH("fcba")
				' end if

				' q = " DELETE FROM marketprice" & _
				' 	" WHERE productcode = '" & productcode & "'" & _
				' 	"       AND customercode = '" & customercode & "'" & _
				' 	"       AND valid_from = to_date('" & valid_from & "','mm/dd/yyyy')" & _
				' 	"       AND fcba = '" & fcba & "'"

				' 	Response.Write q
				' 	connStr.execute(q)
				q = " DELETE FROM marketprice"
					' Response.Write q
					connStr.execute(q)

	For n=1 to (cdbl(request("no"))-1)
			if cdbl(to_nol(request("cb_"&TRIM(n)))) = 1 then
				

				q =	" INSERT INTO marketprice ( productcode" & _
					"                           , customercode" & _
					"                           , rate" & _
					"                           , valid_from" & _
					"                           , fcba" & _
					"                           , fcentry" & _
					"                           , fcedit" & _
					"                           , fcip" & _
					"                           , lastupdate" & _
					"                           , lasttime) " & _
					" VALUES ( '" & productcode & "'" & _
					"          , '" & customercode & "'" & _
					"          , replace('" & rate & "',',','')" & _
					"          , TO_DATE('" & valid_from & "', 'mm/dd/yyyy')" & _
					"          , '" & request("fcba_"&TRIM(n)) & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _					
					"          , SYSDATE" & _
					"          , TO_CHAR(SYSDATE,'HH24:MI'))"
				' response.write q
				connStr.execute(q)
				' connstr.BeginTrans
				' q = " UPDATE marketprice" & _
				'     " SET rate = REPLACE('" & rate & "',',','')" & _
				'     "     , productcode = '" & productcode & "'" & _
				'     "     , customercode = '" & customercode & "'" & _
				'     "     , fcba = '" & request("fcba_"&TRIM(n)) & "'" & _
				'     "     , fcedit = '" & Session("CodeName") & "'" & _
				'     "     , lastupdate = SYSDATE" & _
				'     "     , lasttime = TO_CHAR(SYSDATE,'HH24:MI')" & _
				' 	" WHERE productcode = '" & productcode & "'" & _
				' 	"       AND customercode = '" & customercode & "'" & _
				' 	" 		AND fcba = '"&  request("fcba_"&TRIM(n)) &"'"

				' response.Write(q)
				' connStr.execute(q)
			End if
		next
		' if err.number=0 then
		' 	connStr.CommitTrans
		' End if
		cmdView	= "True"
		end if			

	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")		    
End if

SudahClose = "N"
If Len(Trim(valid_from)) = 0 Then	valid_from = textDate(date)

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
<form action="marketprice.asp" method="post" name="eForm">
<ul class="tabs">
    <li><a href="#view1">Market Price</a></li>
    <li><a href="#view2">Market Price Extend</a></li>
</ul>
<div class="tabcontents">
    <div id="view1">
    	<table class="data" border="0" width="100%" >
    		<%IF Request("txtJenis")="Edit" then%>
    			<tr>
					<td align="left" width="20%">Valid From
						<td align="left">
						<input class="datepicker" type="text" maxlength="15" size="10" name="valid_from"  value="<%=valid_from%>" readony style="background-color: #eee;" >
						</td>
					</td>
				</tr>
				<tr>
					<td align="left" width="20%">Product
						<td align="left">
							<input type="text" maxlength="15" size="10" name="productcode"  value="<%=productcode%>" readony style="background-color: #eee;" >
						</td>
					</td>
				</tr>
				<tr>
					<td align="left" width="20%">Customer
						<td align="left">
							<input type="text" maxlength="15" size="10" name="customercode"  value="<%=customercode%>" readony style="background-color: #eee;" >
						</td>
					</td>
				</tr>
				<tr>
					<td align="left" width="20%">Rate
						<td align="left">
						<input type="text" maxlength="15" name="rate" size="30" value="<%=rate%>" autocomplete="Off"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" readony style="background-color: #eee;" >
						</td>
					</td>
				</tr>
    		<%Else%>
	    		<tr>
					<td align="left" width="20%">Valid From
						<td align="left">
						<input class="datepicker" type="text" maxlength="15" size="10" name="valid_from"  value="<%=valid_from%>" >
						</td>
					</td>
				</tr>
				<tr>
					<td align="left" width="20%">Product
						<td align="left">
							<% 	Dim rsCropA
								q = "select distinct productcode, productdescription from product order by productcode"
								Set rsCropA = Server.CreateObject("ADODB.Recordset")
								rsCropA.CursorType = 3
								rsCropA.CursorLocation = 3
								rsCropA.Open q, connstr
								if not rsCropA.eof then%>
								<select id="productcode" name="productcode" class="chosen-select" style="width:30%;">         
								<option value=""></option> 
								<%
								Do While Not rsCropA.eoF%>
								<%IF Request("productcode")=rsCropA("productcode") Then%>
									<option value="<%=rsCropA("productcode")%>" selected><%=rsCropA("productcode")%> : <%=rsCropA("productdescription")%></option> 
								<%Else%>
									<option value="<%=rsCropA("productcode")%>"><%=rsCropA("productcode")%> : <%=rsCropA("productdescription")%></option> 
								<%End if%>
								<%rsCropA.movenext
								loop
								End if
								rsCropA.Close
								set rsCropA = Nothing
								q=""
								%>
						<select>
						</td>
					</td>
				</tr>
				<tr>
					<td align="left" width="20%">Customer
						<td align="left">
							<% 	Dim rsCropB
								q = "select distinct customercode, description from Customer order by customercode"
								Set rsCropB = Server.CreateObject("ADODB.Recordset")
								rsCropB.CursorType = 3
								rsCropB.CursorLocation = 3
								rsCropB.Open q, connstr
								if not rsCropB.eof then%>
								<select id="customercode" name="customercode" class="chosen-select" style="width:70%;">         
								<option value=""></option> 
								<%
								Do While Not rsCropB.eoF%>
								<%IF Request("customercode")=rsCropB("customercode") Then%>
									<option value="<%=rsCropB("customercode")%>" selected><%=rsCropB("customercode")%> : <%=rsCropB("description")%></option> 
								<%Else%>
									<option value="<%=rsCropB("customercode")%>"><%=rsCropB("customercode")%> : <%=rsCropB("description")%></option> 
								<%End if%>
								<%rsCropB.movenext
								loop
								End if
								rsCropB.Close
								set rsCropB = Nothing
								q=""
								%>
						<select>
						</td>
					</td>
				</tr>
				<tr>
					<td align="left" width="20%">Rate
						<td align="left">
						<input type="text" maxlength="15" name="rate" size="30" value="<%=rate%>" autocomplete="Off"  data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" >
						</td>
					</td>
				</tr>
    		<%End If%>
		</table>
    </div>
    <div id="view2">
    	<%
    		q = " SELECT DISTINCT fccode, fcname " & _
			 	" FROM (SELECT businessunit.fccode, businessunit.fcname " & _
			    " FROM user_fcba_tab " & _
			    " INNER JOIN businessunit " & _
			    " ON user_fcba_tab.fcba = " & _
			    " businessunit.fccode " & _
			    " WHERE businessunit.isactive = 'TRUE' AND user_fcba_tab.fcusercode = '"&  Session("CodeName") &"') order by fccode"
			' response.write q
	Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
	rsPurchaseRequest.Open q, connstr
    	%>
    	<table class="data" style="width:50%;">
			<BR>
			<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
			<TD align="center">No</TD>
			<TD align="center">Kebun/Pabrik</TD>
			<TD align="center">Description</TD>
			<TD align="center"><input type='checkbox' name='select-all' id='select-all'/></TD>	
			</tr> 
			
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
					<td align="center" valign="top" width="20%"><%=NO%></td>
					<td align="center" title="<%=rsPurchaseRequest("fccode").Name%>" valign="top"><%=rsPurchaseRequest("fccode")%></td>	
					<td align="center" title="<%=rsPurchaseRequest("fcname").Name%>" valign="top"><%=rsPurchaseRequest("fcname")%></td>	
				<%IF Request("txtJenis")="Edit" then%>
					<td valign="middle" align="center">
						<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" value= "<%=fcba%>" checked>
						<input type="hidden" name="fcba_<%=NO%>" value="<%=rsPurchaseRequest("fccode")%>">
					</td>
				<%Else%>
					<td valign="middle" align="center">
						<input type="checkbox" name="cb_<%=NO%>" id="cb_<%=NO%>" class='item' value= "" onclick='handleClick(this);'>
						<input type="hidden" name="fcba_<%=NO%>" value="<%=rsPurchaseRequest("fccode")%>">
					</td>
				<%End if%>
				  </tr>
				  <%
				  rsPurchaseRequest.MoveNext
				  intRecordsPerPage = intRecordsPerPage - 1
				  no=no+1
			  Wend
		   else
		      response.Write"<tr><td colspan='"&d_span&"' align='left'>Data is empty</td></tr>"
		   end if 
		   rsPurchaseRequest.Close : set rsPurchaseRequest = Nothing
		%>
		</table>
    </div>
</div>
<br>
	<tr>
		<td align="left" colspan="2">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<input Type="Hidden" Value="<%=NO%>" Name="no">
				<%IF Request("txtJenis")="Edit" then%>
					<input type="hidden" name="txtJenis" value="Edit">
					<input type="hidden" name="fcba_<%=NO%>" value="<%=fcba%>">
					<input type="hidden" name="txtlineno" value="<%=txtlineno%>">
					<input Type="Submit"  name="Add" value="Update">
				<%Else%>
					<input type="hidden" name="txtJenis" value="New">
					<input Type="Submit"  name="Add" value="Add">
				<%End if%>
				<input Type="Reset"  name="Reset" value="Reset">
				<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
			
		</td>
	</tr>
</form>
<%	
	Dim rsVH, intRecordsPerPage, cColor
	
'query Untuk menampilkan isi di Field
	q	= " SELECT mp.productcode, " & _
       	  "	mp.customercode," & _
       	  "	cust.description," & _
       	  " mp.rate," & _
          " mp.valid_from," & _
          " mp.fcba," & _
          " bu.fcname" & _
  		  " FROM marketprice mp LEFT JOIN businessunit bu ON (bu.fccode = mp.fcba) " & _
  		  " INNER JOIN customer cust ON (cust.customercode = mp.customercode and cust.fcba = mp.fcba) order by mp.fcba, valid_from" 

	' response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr

	if Not rsVH.eof then
		productcode =  rsVH("productcode")
		customercode =  rsVH("customercode")
		customername =  rsVH("description")
		description =  rsVH("description")
		rate =  rsVH("rate")
		valid_from =  rsVH("valid_from")
		fcname =  rsVH("fcname")
		fcba =  rsVH("fcba")
	end if


' 'Query untuk menampilkan table Delivery Order
' 	q	= " SELECT * FROM deliveryorder " & _
' 		  " WHERE fcba = '" & fcba & "'" & _
' 		  " AND salescontractno = '" & salescontractno & "'" 

' 	' response.Write q
' 	Set rsVH = Server.CreateObject("ADODB.Recordset")
' 	rsVH.CursorType = 3
' 	rsVH.CursorLocation = 3
' 	rsVH.Open q, connstr
%>
 <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td  width="30">&nbsp;</td>
		<td align="left">Kebun / Pabrik</td>
		<td align="left">Valid From</td>
		<td align="left">Product</td>
		<td align="left">Customer</td>
		<td align="left">Rate</td>
   
   <%
	if Not rsVH.eof then
		nCountofField = rsVH.fields.count - 1
		no = 1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=NO%></td>
			<td valign="middle">
				<a href="marketprice.asp?fcba=<%=rsVH("fcba")%>&productcode=<%=rsVH("productcode")%>&valid_from=<%=rsVH("valid_from")%>&rate=<%=rate%>&customercode=<%=rsVH("customercode")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtDocumentNo")%>" border="0"></a>
			</td>
			    <td align="left" title="<%=rsVH("fcname").Name%>" valign="top"><%=rsVH("fcname")%></td>
			    <td align="left" title="<%=rsVH("valid_from").Name%>" valign="top"><%=rsVH("valid_from")%></td>
			    <td align="left" title="<%=rsVH("productcode").Name%>" valign="top"><%=rsVH("productcode")%></td>
			    <td align="left" title="<%=rsVH("customercode").Name%>" valign="top"><%=rsVH("customercode")%> : <%=rsVH("description")%></td>
		  		<td align="left" title="<%=rsVH("rate").Name%>" valign="top"><%=formatnumber(cdbl(to_nol(rsVH("rate"))),2)%></td>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>

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
	<script src="plugins/tabcontent/tabcontent.js" type="text/javascript"></script>
	
	<script type="text/javascript">
	function handleClick(cb){
		if(cb.checked == 1){
			cb.value = 1;
		} else {
			cb.value = 0;
		}
	}

</script>
<script type="text/javascript">	
	$('#select-all').click( function () {
		$('.item').prop('checked', this.checked);
		$('.item').val(this.checked ? 1 : 0);
	})
</script>
		<script src="/js/jquery.inputmask.bundle.min.js"></script>
		<script src="/js/inputmask.binding.js"></script>
</body>
</html>
