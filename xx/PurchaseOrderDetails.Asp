<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Purchase Order</title>
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
 <%
 '-------------awal-----------------------
 
 Dim xCode,fcCode,cmdView,PoCode,PoDate,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim RowStateId,RowStateHdr
 Dim PRCODE, PRLINEID
 Dim tempStringError, rsTempPurchaseRequest
 Dim totalPRData
 Dim ApprovedQty, ORDEREDQTY
 Dim strQueryStatus, rsQueryStatus
 Dim sqlupdateqtypr, rsupdateqtypr
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = "" 
 Dim IsInBridge, IsFrezee
 
 Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 if request("fccOde")="" then
    fcCode="" '----- Nantinya Ambil dari Format Master
 else
    fcCode = request("fccOde")
 end if
 
 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 '---- Data Header
 if Request("PODATE")<>"" then
   PODATE   =Request("PODATE")
 else
   PODATE   =textDate(date)
 end if
 
  PODATE = FormatDateTime(PODATE)
 
 xCode = Request("xCode")
 IsInBridge = Request("IsInBridge")
 Remarks = Request("Remarks")
 
 '---- Data Details
 ' Add by NAF, 2012-12-1
 LINENOKEY=Request("LineNoKey")
 ' End Added by NAF
 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID")
 use_ppn = Request("use_ppn")
if "" & use_ppn = "" then use_ppn = "N" 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "Ordered" or Request("txtJenis")= "CompletedDetails"  or  Request("txtJenis")= "CanceledDetails"  or  Request("txtJenis")= "UnRelease" Then
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
       connstr.BeginTrans
		
				sqlupdateqtypr = " select prcode, prlineid from purchaseorderdetails" & _
                                 " where pocode = '" & fcCode & "'" & _
                                 "       and fcba = '" & Session("fcBA") & "'" & _
                                 "       and prcode is not null" & _
                                 "       and prlineid is not null"
								 
				Set rsupdateqtypr = Server.CreateObject("ADODB.Recordset")
				rsupdateqtypr.Open sqlupdateqtypr, connstr, adLockOptimistic, adOpenStatic
				
				if not rsupdateqtypr.EOF then
					While Not rsupdateqtypr.EOF
							q = " update purchaserequestdetails" & _
								" set orderedqty = 0" & _
								" where  prcode = '" & rsupdateqtypr("prcode") & "'" & _
								"        and lineid = '" & rsupdateqtypr("prlineid") & "'" & _
								"        and fcba = '" & Session("fcBA") & "'"

							connStr.execute(q)
						rsupdateqtypr.MoveNext
					wend
				end if
							
				rsupdateqtypr.Close
				set rsupdateqtypr = Nothing 
				
				q="Delete from PurchaseOrderDetails where  POCODE='" & fcCode & "' and fcBA='" & Session("fcBA") & "'"
				connStr.execute(q)
			   
				q="Delete from PurchaseOrder where  POCODE='" & fcCode & "' and fcBA='" & Session("fcBA") & "'"
				connStr.execute(q)
				
				if err.number  = 0 then
					connStr.CommitTrans
				end if
	   cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans
			   q="Delete from PurchaseOrderDetails where  POCODE='" & fcCode & "' and ItemCode='" & request("ItemCode") & "' and LineNoKey='" & request("LineNoKey") & "' and fcBA='" & Session("fcBA") & "'"
			   connStr.execute(q)
			   if err.number  = 0 then
				  connStr.CommitTrans
			   end if
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
	  Dim rsPO
	  Set rsPO = Server.CreateObject("ADODB.Recordset")
	  connstr.BeginTrans
			  rsPO.Open "Select poCode from PurchaseOrder where poCode='" & fcCode & "' AND FCBA = '" & Session("FCBA") & "'" , connstr
			  IF rsPO.eof then	   
				 ' Modified by NAF
				 ' To handle first time Input for Purchase Order
				 q = " INSERT INTO purchaseorder ( pocode, podate, suppliercode, remarks, use_ppn, fcentry, fcedit, fcip, fcba, rowstateid, lastupdate, lasttime ) " & _
				     " VALUES ( '" & fcCode & "'" & _
					 "         , TO_DATE('" & PoDate & "', 'mm/dd/yyyy')" & _
					 "         , '" & xCode & "'" & _
					 "         , '" & Remarks & "'" & _
					 "         , '" & use_ppn & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Session("CodeName") & "'" & _
					 "         , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					 "         , '" & Session("fcBA") & "'" & _
					 "         , 'Planned'" & _
					 "         , SYSDATE" & _
					 "         , TO_CHAR(SYSDATE,'hh24:mi'))"
				 connStr.execute(q)
			  End if
			  rsPO.Close
			  if err.number = 0 then	
			  
				' Added by NAF, 2012-12-01
				 strLineMax = " select nvl(max(linenokey),0) maxLineNoKey from purchaseorderdetails where fcBA='" & session("fcBA") & "'"

			  
				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
				end if
							
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 
				' End of Added by NAF, 2012-12-01
			  
			  
			  
			  

				 
				 
				 if len(trim(PRCODE)) <> 0 or len(trim(PRLINEID)) <> 0 then 
					set rsTempPurchaseRequest = server.CreateObject("Adodb.recordset")
					q="select APPROVEDQTY, ORDEREDQTY  from PURCHASEREQUESTDETAILS  " & _
					  " where PrCode='" & PRCODE & "' and " & _
					  " lineid = '" & PRLINEID & "' and " & _
					  " itemcode = '" & ITEMCODE & "' and " & _
					  " fcBA='" & session("fcBA") & "'"
					rsTempPurchaseRequest.Open q, connstr
					if Not rsTempPurchaseRequest.Eof then
						ApprovedQty = cdbl(rsTempPurchaseRequest("APPROVEDQTY"))
						OrderedQty = cdbl(rsTempPurchaseRequest("ORDEREDQTY"))
						
						totalPRData= 1
					else
						totalPRData = 0
						ApprovedQty = 0
						OrderedQty = 0
						
					End if
					
					rsTempPurchaseRequest.Close
					set rsTempPurchaseRequest = Nothing
					
					if totalPRData <> 0 then 
						if err.number  = 0 then
						
							if ApprovedQty - OrderedQty - Quantity >= 0 then 
								q="Insert into  PURCHASEORDERDETAILS (POCODE, ITEMCODE, QUANTITY, AMOUNT, OTHERITEMDESC, RECEIVED, FCENTRY, FCEDIT, FCIP, FCBA, LINENOKEY, PRCODE, PRLINEID, LASTUPDATE, LASTTIME, ROWSTATEID ) " & _
								 "values ('" & fcCode & "','" & ITEMCODE & "','"  &  Quantity & "','" & AMOUNT & "'" & _
								 ",'" & OTHERITEMDESC & "'," & "0" & ",'" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
								 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLineNo & ",'" & PRCODE & "','"& PRLINEID &"', sysdate, To_Char(sysdate,'HH24:MI'), 'Opened')"
								 
								connStr.execute(q)				
							
								connStr.CommitTrans
								tempStringError = "PO Saved Successfully with PR Info, if you want to modify just click Edit Button, Thank you."				 
								response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
							else
								connStr.RollbackTrans
								tempStringError = "Your Quantity is not Satisfied for Available Quantity, Please Check them in PR Form."
								response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
							end if
						end if
					else
						connStr.RollbackTrans
						tempStringError = "PR Information are not match with the Current PR Data, Please Check them in PR Form."
						response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
					End if	
				 else
					 if err.number  = 0 then

						q="Insert into  PURCHASEORDERDETAILS (POCODE, ITEMCODE, QUANTITY, AMOUNT, OTHERITEMDESC, RECEIVED, FCENTRY, FCEDIT, FCIP, FCBA, LINENOKEY, PRCODE, PRLINEID, LASTUPDATE, LASTTIME, ROWSTATEID ) " & _
						 "values ('" & fcCode & "','" & ITEMCODE & "','"  &  QUANTITY & "','" & AMOUNT & "'" & _
						 ",'" & OTHERITEMDESC & "'," & "0" & ",'" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
						 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLineNo & ",'" & PRCODE & "','"& PRLINEID &"', sysdate, To_Char(sysdate,'HH24:MI'), 'Opened')"
						 
						connStr.execute(q)				

						connStr.CommitTrans
						tempStringError = "PO Saved Successfully without PR Info, thank you."				 
						response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
					 else
						connStr.RollbackTrans
						tempStringError = "Some Error happened, Please Check data in PO Form."
						response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
					 
					 end if
				 end if

			  end if
      cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	   connstr.BeginTrans
			   q="Update PurchaseOrderDetails set Quantity='" & Quantity & "',Amount='" & Amount & "'" & _
				 ",OTHERITEMDESC='" & OTHERITEMDESC  & "'," & _
				 "PRCODE= '"& PRCODE & "'," & _
				 "PRLINEID = '" & PRLINEID & "'" & _
				 ", LASTUPDATE = sysdate " & _
				 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
				 " where  POCode='" & fcCode & "' and ItemCode='" & ITEMCODE & "' and LineNoKey='" & LineNoKey & "' and fcBA='" & Session("fcBA") & "'"
				connStr.execute(q)
				
				if err.number  = 0 then
				   connStr.CommitTrans
				end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Confirmed" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & Session("fcBA") & "','" & fcCode & "', 'Confirm' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "RePlanned" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & Session("fcBA") & "','" & fcCode & "', 'RePlan' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	elseif Request("txtJenis")= "Ordered" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & Session("fcBA") & "','" & fcCode & "', 'Order' ); END; "
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
    elseif Request("txtJenis")= "UnRelease"	then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.UnReleasePO('" & fcCode & "' ); END; "
	   response.Write strQueryStatus
	   connStr.execute(strQueryStatus)	 
	   cmdView	= "True"
	elseif Request("txtJenis")= "CompletedDetails" then
		connstr.BeginTrans
		strQueryStatus = " BEGIN" & _
						 " PURCHASEORDERDETAILS_API.NextState('" & Session("fcBA") & "'" & _
	                     "                                    , '" & fcCode & "'" & _
					     "                                    , '" & txtLineNoKey & "'" & _
	                     "                                    , 'Complete'" & _
	                     "                                    );" & _
	                     " END; "
		'response.Write strQueryStatus
		connStr.execute(strQueryStatus)	 

		cmdView	= "True"
	
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if

 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select * from PURCHASEORDER  " & _
  " where PoCode='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    PoDate=rsPurchase("PoDate")
	PODATE = FormatDateTime(PODATE)
	xCode=rsPurchase("SUPPLIERCODE")
	SHIPPINGPOINT = rsPurchase("INCOTERM1")
	Remarks=rsPurchase("Remarks")
	RowStateId = rsPurchase("RowStateId")
	RowStateHdr = rsPurchase("RowStateId")
	use_ppn = rsPurchase("use_ppn")
	IsInBridge  = rsPurchase("IsInBridge")
	
 End if
 
 dim rsFrez
 set rsFrez = server.CreateObject("Adodb.recordset")
	q=" select distinct frezee  from Purchaseorderdetails " & _
	  " where PoCode='" & fcCode & "' and " & _
	  " fcBA='" & session("fcBA") & "'"
 rsFrez.Open q, connstr
 if Not rsFrez.Eof then
	IsFrezee  = rsFrez("frezee")
 else
    IsFrezee ="0" 
 End if
 rsFrez.Close
 set rsFrez = Nothing 
 
  'response.write "IsInBridge--> " & IsInBridge
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="PurchaseOrderDetails.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Purchase Order No</td>
		<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">*</font></td>
		</tr>
		
		<tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="PODATE" size="10" value="<%=PODATE%>">
		</td>
		</tr>
		<tr>
			<td align="left">Supplier</td>
			<td align="left">      
			<select name="xCode">
			<%cmdCallCombo "Supplier",xCode,"ACTIVATION","Y"%>
			</select><font color="#FF0000">*</font>
			</td>
			<tr>
			<td align="left">Remarks</td>
			<td align="left"><textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea></td>		
			</tr>
		</tr>
		<tr>
			<td align="left">Inco Terms</td>	
			<td align="left"><textarea readonly name="SHIPPINGPOINT"  cols="5" ><%=SHIPPINGPOINT%></textarea></td>		
		</tr>
		</tr>
		<%if Session("fcBA") = "KKP" OR Session("fcBA") = "KPK" then%>
		<tr>
		<td align="left">Ppn (10%)</td>
		<td align="left">
			<input type="checkbox" id="use_ppn" class="css-checkbox" name="use_ppn" value="<%=use_ppn%>" onclick='handleClick(this);' <%if use_ppn = "Y" then%> checked <%end if%>>
			<label for="use_ppn" class="css-label"></label>
		</td>
		</tr>
		<%end if%>
		
		<tr>
		<td align="left">Status</td>
		<td align="left"><input disabled type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		
		<%if Session("fcBA") = "KKP" OR Session("fcBA") = "KPK" then%>
		<a href="printingpo.Asp?pocode=<%=fcCode%>" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp;
		<%else%>
	    <a href="PrintPO.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&PoDate=<%=PoDate%>&Mode=Print" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp;
		<%end if%>
		<a href="PurchaseOrderDetails.Asp?fcCode=<%=fcCode%>&PODATE=<%=PODATE%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" And cint(IsInBridge) = 1 then %>
		    <a href="PurchaseOrderDetails.Asp?fcCode=<%=fcCode%>&txtJenis=Ordered" onClick="return confirm('Are you sure to Order this data :  <%=fcCode%> ?');"> Ordered </a>
		<% elseif RowStateId = "Planned" And cint(IsInBridge) = 0 then %>
		    <a href="PurchaseOrderDetails.Asp?fcCode=<%=fcCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to Confirm this data :  <%=fcCode%> ?');"> Confirmed </a>
		<% elseif RowStateId = "Confirmed" then %>	
		    <a href="PurchaseOrderDetails.Asp?fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=fcCode%> ?');"> Re-Planned </a>
		<% elseif  RowStateId = "Ordered" And cint(IsInBridge) = 1 And Cint(IsFrezee) =0 then %>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="PurchaseOrderDetails.Asp?fcCode=<%=fcCode%>&txtJenis=UnRelease" onClick="return confirm('Are you sure to Un Release this data :  <%=fcCode%> ?');"> UnRelease </a>
		<% end if%>
		
		
		
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 Then
	Dim rsPurchaseOrder,intRecordsPerPage,cColor

	q = " select pocode" & _
	    "        , rowstateid" & _
		"        , itemcode" & _
	    "        , quantity" & _
	    "        , amount" & _
	    "        , otheritemdesc" & _
	    "        , received" & _
	    "        , linenokey" & _
	    "        , prcode" & _
	    "        , prlineid" & _
		"        , FREZEE " & _
	    " from purchaseorderdetails" & _
		" where PoCode = '" & fcCode & "'" & _
	    "       and fcBA = '" & session("fcBA") & "'"
	'response.Write q	
	Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
	rsPurchaseOrder.CursorType = 3
	rsPurchaseOrder.CursorLocation = 3
	rsPurchaseOrder.Open q, connstr
%>
   


   
   <form action="PurchaseOrderDetails.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr>
		   <td  align="left">Stock Code</td>
		   <td align="left">
			<%IF txtJenis="Edit" Then%>
				 <input type="text" size="15" name="ITEMCODE" value=<%=ITEMCODE%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">*
			<%else%>
				 <input type="text" size="15" name="ITEMCODE" value=<%=ITEMCODE%>>&nbsp;<font color="#FF0000">*
			<%End if%>
			<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Stock&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font></a></td>
		</tr>

		<tr>
			<td align="left">Quantity</td>
			<td align="left"><input type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>" onkeypress="return harusangka(event);">&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		
		<tr>
			<td align="left">Amount</td>
			<td align="left"><input type="text" size="15" name="AMOUNT" value="<%=AMOUNT%>"  onkeypress="return harusangka(event);">&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		
		<tr>
			<td align="left">Description</td>
			<td align="left">
				<textarea  name="OTHERITEMDESC"  cols="50" rows="3"><%=OTHERITEMDESC%>
				</textarea>Please Do Not write (comma) or "
			</td>
		</tr>

		<tr>
		   <td align="left">Please write PR# Information below, if any....</td>
		   <td></td>
		</tr>

		
		
		<tr>
			<td align="left">PR. Number.</td>
			<td align="left"><input type="text" size="15" name="PRCODE" value=<%=PRCODE%>>&nbsp;<font color="#FF0000">*
			</td>
		</tr>


		<tr>
			<td align="left">PR. Line Number.</td>
			<td align="left"><input type="text" size="15" name="PRLINEID" value=<%=PRLINEID%>>&nbsp;<font color="#FF0000">*
			</td>
		</tr>	
		
		
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="PODATE" value="<%=PODATE%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="LineNokey" value="<%=LINENOKEY%>">
		  <input type="hidden" name="use_ppn" value="<%=use_ppn%>">
		  <input type="hidden" name="RowStatehdr" value="<%=RowStatehdr%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else  %>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset" name="Reset" value="Reset">
		  <input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="middle" colspan="2"> ACTION</TD>
	<TD align="left">PO CODE</TD>
	<TD align="left">LINENOKEY</TD>
	<TD align="left">ITEM CODE</TD>
	<TD align="left">QUANTITY</TD>
	<TD align="left">AMOUNT</TD>
	<TD align="left">OTHER ITEM DESC</TD>
	<TD align="left">RECEIVED</TD>
	<TD align="left">PRCODE</TD>
	<TD align="left">PRLINEID</TD>
	<TD align="left">LINE STATUS</TD>
	<TD align="left">FREZEE</TD>
	</tr>   
   <%
   if Not rsPurchaseOrder.eof then
      nCountofField=rsPurchaseOrder.fields.count-1
	  no=1
	  While Not rsPurchaseOrder.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<% If rsPurchaseOrder("ROWSTATEID")="Planned" Then %>			
			    <td valign="middle" width="10">
				   <a href="PurchaseOrderDetails.Asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("ItemCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&OTHERITEMDESC=<%=rsPurchaseOrder("OTHERITEMDESC")%>&PRCODE=<%=rsPurchaseOrder("PRCODE")%>&PRLINEID=<%=rsPurchaseOrder("PRLINEID")%>&txtJenis=Edit">
				   <img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(1)%>" border="0"></a>
				</td>
				<td valign="middle" width="10">
				   <a href="PurchaseOrderDetails.Asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("ItemCode")%>&PODATE=<%=PODATE%>&txtJenis=Details">
				   <img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');">
				   </a>
				</td>
            <% ElseIf rsPurchaseOrder("ROWSTATEID")="Ordered"  And cint(IsInBridge) = 0 then %>	
			    <td valign="middle" width="10">
				   <a href="PurchaseOrderDetails.Asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("ItemCode")%>&Quantity=<%=rsPurchaseOrder("Quantity")%>&Amount=<%=rsPurchaseOrder("Amount")%>&OTHERITEMDESC=<%=rsPurchaseOrder("OTHERITEMDESC")%>&PRCODE=<%=rsPurchaseOrder("PRCODE")%>&PRLINEID=<%=rsPurchaseOrder("PRLINEID")%>&txtJenis=CanceledDetails" onClick="return confirm('Are you sure to Cancel This data ?');">Cancel</a>
				</td>
				<td valign="middle" width="10">
				   <a href="PurchaseOrderDetails.Asp?fcCode=<%=rsPurchaseOrder("POCode")%>&LineNoKey=<%=rsPurchaseOrder("LineNoKey")%>&ItemCode=<%=rsPurchaseOrder("ItemCode")%>&PODATE=<%=PODATE%>&txtJenis=CompletedDetails" onClick="return confirm('Are you sure to Close This data ?');">Close</a>
				</td> 
			<% Else %>	
               	<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>		
            <% End If %>	
				
				<td align="left" title="<%=rsPurchaseOrder("pocode").Name%>" valign="top">
					<%=rsPurchaseOrder("pocode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("linenokey").Name%>" valign="top">
					<%=rsPurchaseOrder("linenokey")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("itemcode").Name%>" valign="top">
					<%=rsPurchaseOrder("itemcode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("quantity").Name%>" valign="top">
					<%=rsPurchaseOrder("quantity")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("amount").Name%>" valign="top">
					<%=rsPurchaseOrder("amount")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("otheritemdesc").Name%>" valign="top">
					<%=rsPurchaseOrder("otheritemdesc")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("received").Name%>" valign="top">
					<%=rsPurchaseOrder("received")%>
				</td>
				
				<td align="left" title="<%=rsPurchaseOrder("prcode").Name%>" valign="top">
					<%=rsPurchaseOrder("prcode")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("prlineid").Name%>" valign="top">
					<%=rsPurchaseOrder("prlineid")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("rowstateid").Name%>" valign="top">
					<%=rsPurchaseOrder("rowstateid")%>
				</td>
				<td align="left" title="<%=rsPurchaseOrder("FREZEE").Name%>" valign="top">
					<%=rsPurchaseOrder("FREZEE")%>
				</td>
			</tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>
</div></div>
  <!--#include file="bottom.jsp"-->
 	<script type="text/javascript">
		function handleClick(cb){
			if(cb.checked == 1){
				cb.value = "Y";
			} else {
				cb.value = "N";
			}
        }
	</script>
</body>
</html>
