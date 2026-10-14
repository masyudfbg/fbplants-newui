<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Purchase Order Approval</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
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
 Dim RowStateId
 Dim PRCODE, PRLINEID
 Dim doneApproved, rsDone, qDone
 Dim tempStringError
 Dim strQueryStatus, rsQueryStatus
 Dim sqlupdateqtypr, rsupdateqtypr
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""


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
 xCode = Request("xCode")
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
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "Approved" or Request("txtJenis")= "Printed"  or  Request("txtJenis")= "Ordered" or Request("txtJenis")= "Rejected" Then
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
				 q="Insert into PURCHASEORDER " & _
				 " ( POCODE, PODATE, SUPPLIERCODE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, ROWSTATEID, LASTUPDATE, LASTTIME ) " & _
				 "values ('" & fcCode & "'," & " to_date('" & PoDate & "','mm/dd/yyyy') " & ",'"  &  xCode & _
				 "','" & Remarks & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','Planned', sysdate, To_Char(sysdate,'HH24:MI'))"
				 connStr.execute(q)
			  End if
			  rsPO.Close
			  if err.number = 0 then	
			  
				' Added by NAF, 2012-12-01
				 strLineMax = " select nvl(max(linenokey),0) maxLineNoKey from PurchaseOrderApproval where fcBA='" & session("fcBA") & "'"

			  
				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					tempMaxLineNo = cdbl(rsstrLineMax("maxLineNoKey")) + 1
				end if
							
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 
				' End of Added by NAF, 2012-12-01
			  
			  
			  
			  
				 q="Insert into  PurchaseOrderDetails (POCODE, ITEMCODE, QUANTITY, AMOUNT, OTHERITEMDESC, RECEIVED, FCENTRY, FCEDIT, FCIP, FCBA, LINENOKEY, PRCODE, PRLINEID, LASTUPDATE, LASTTIME ) " & _
				 "values ('" & fcCode & "','" & ITEMCODE & "','"  &  QUANTITY & "','" & AMOUNT & "'" & _
				 ",'" & OTHERITEMDESC & "'," & "0" & ",'" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
				 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLineNo & ",'" & PRCODE & "','"& PRLINEID &"', sysdate, To_Char(sysdate,'HH24:MI'))"
				 
				 connStr.execute(q)
				 if err.number  = 0 then
					connStr.CommitTrans
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
	elseif Request("txtJenis")= "Rejected" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & Session("fcBA") & "','" & fcCode & "', 'Reject' ); END; "
	   connStr.execute(strQueryStatus)	 


       cmdView	= "True"
	   
	elseif Request("txtJenis")= "Approved" then
		connstr.BeginTrans
	  
	    doneApproved = "TRUE"
		set rsDone = server.CreateObject("Adodb.recordset")
		qDone = " select purchaseorder_api.getApproved('" & fcCode & "','" & Session("fcBA") & "') tempDone from dual"
	   
		rsDone.Open qDone, connstr
		if Not rsDone.Eof then
			doneApproved = rsDone ("tempDone")
		End if

        rsDone.Close
		Set rsDone = nothing
		
		IF doneApproved = "TRUE" THEN 
		
		   set rsDone = server.CreateObject("Adodb.recordset")
		   qDone = " begin purchaseorder_api.setApproved('" & fcCode & "','" & Session("fcBA") & "'); end;"
		   rsDone.Open qDone, connstr
		   if err.number  = 0 then
			  connStr.CommitTrans	

			  tempStringError = "Conratulations, PO Approved Successfully, Thank you."				 
			  response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
			  
		   end if
		ELSE
		    tempStringError = "Failed, Unable to be Approved, Please check your data, Thank you."				 
			response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
	    END IF
        cmdView	= "True"	   
	   
	elseif Request("txtJenis")= "Printed" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & Session("fcBA") & "','" & fcCode & "', 'Print' ); END; "
	   connStr.execute(strQueryStatus)	 


       cmdView	= "True"	   
		
	elseif Request("txtJenis")= "Ordered" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDER_API.NextState('" & Session("fcBA") & "','" & fcCode & "', 'Order' ); END; "
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
	xCode=rsPurchase("SUPPLIERCODE")
	Remarks=rsPurchase("Remarks")
	RowStateId = rsPurchase("RowStateId")
	use_ppn = rsPurchase("use_ppn")
 End if
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="PurchaseOrderApproval.Asp">
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
		<input Type="Submit" name="Add" value="Select">&nbsp;&nbsp;
	    <%if Session("fcBA") = "KKP" OR Session("fcBA") = "KPK" then%>
		<a href="printingpo.Asp?pocode=<%=fcCode%>" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp;
		<%else%>
	    <a href="PrintPO.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&PoDate=<%=PoDate%>&Mode=Print" target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font>
		</a>&nbsp;&nbsp;
		<%end if%>
		<a href="PurchaseOrderApproval.Asp?fcCode=<%=fcCode%>&PODATE=<%=PODATE%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" then %>
		    <a href="PurchaseOrderApproval.Asp?fcCode=<%=fcCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to Confirm this data :  <%=fcCode%> ?');"> Confirmed </a>&nbsp;&nbsp;
		<% elseif RowStateId = "Confirmed" then %>	
		    <a href="PurchaseOrderApproval.Asp?fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=fcCode%> ?');"> Re-Planned </a>&nbsp;&nbsp;
		    <a href="PurchaseOrderApproval.Asp?fcCode=<%=fcCode%>&txtJenis=Rejected" onClick="return confirm('Are you sure to Reject this data :  <%=fcCode%> ?');"> Rejected </a>&nbsp;&nbsp;
		    <a href="PurchaseOrderApproval.Asp?fcCode=<%=fcCode%>&txtJenis=Approved" onClick="return confirm('Are you sure to Approve this data :  <%=fcCode%> ?');"> Approved </a>		&nbsp;&nbsp;
		<% elseif RowStateId = "Approved" then %>	
		    <a href="PurchaseOrderApproval.Asp?fcCode=<%=fcCode%>&txtJenis=Printed" onClick="return confirm('Are you sure to Print the Status of this data :  <%=fcCode%> ?');"> Printed </a>&nbsp;&nbsp;
		<% elseif RowStateId = "Printed" then %>	
		    <a href="PurchaseOrderApproval.Asp?fcCode=<%=fcCode%>&txtJenis=Ordered" onClick="return confirm('Are you sure to Send the Order to Supplier of this data :  <%=fcCode%> ?');"> Ordered </a>&nbsp;&nbsp;
		<% end if%>
		
		
		
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(fcCode)) > 4 Then
   Dim rsPurchaseOrder,intRecordsPerPage,cColor
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q="select POCODE, ITEMCODE, QUANTITY, AMOUNT, OTHERITEMDESC, RECEIVED, LINENOKEY, PRCODE, PRLINEID, FCENTRY, FCEDIT, FCIP, FCBA from PurchaseOrderDetails  " & _
     " where PoCode='" & fcCode & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   


   
   <form action="PurchaseOrderApproval.Asp" method="post" name="eForm">
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
			<td align="left"><input type="text" size="15" name="QUANTITY" value=<%=QUANTITY%>>&nbsp;<font color="#FF0000">*
			</td>
		</tr>
		
		<tr>
			<td align="left">Amount</td>
			<td align="left"><input type="text" size="15" name="AMOUNT" value=<%=AMOUNT%>>&nbsp;<font color="#FF0000">*
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
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <% if rowstateId = "Confirmed" then %> 
			  <%IF Request("txtJenis")="Edit" then%>
				   <input type="hidden" name="txtJenis" value="Edit">
				   <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
				   bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
			  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">PO CODE</TD>	
	<TD align="left">ITEM CODE</TD>	
	<TD align="left">QUANTITY</TD>	
	<TD align="left">AMOUNT</TD>	
	<TD align="left">OTHER ITEM DESC</TD>	
	<TD align="left">RECEIVED</TD>	
	<!--<TD align="left">FCENTRY</TD>	
	<TD align="left">FCEDIT</TD>	
	<TD align="left">FCIP</TD>	
	<TD align="left">FCBA</TD>-->	
	<TD align="left">LINENOKEY</TD>	
	<TD align="left">PRCODE</TD>	
	<TD align="left">PRLINEID</TD>	
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
			<td valign="middle" width="10">
			<%Dim rsCekClose,qCekClose
			Set rsCekClose = Server.CreateObject("ADODB.Recordset")
            rsCekClose.CursorType = 3
            rsCekClose.CursorLocation = 3
			qCekClose="Select POCode from PurchaseOrder where POCode='" & rsPurchaseOrder("POCode") & "'" & " and rowstateId = 'Planned' and PODATE >=to_Date('" & Session("cls_Purchase") & "','mm/dd/yyyy')"
			' response.Write(qCekClose)
			' response.Flush()
			rsCekClose.Open qCekClose, connstr
			if Not rsCekClose.Eof then%>
			    <td colspan="1"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%Else%>	 
			    <td colspan="1"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if
			rsCekClose.close
			set rsCekClose = Nothing%>	 
		    <%for i = 0 to nCountofField - 4%>
			    <td align="left" title="<%=rsPurchaseOrder(i).Name%>" valign="top"><%=rsPurchaseOrder(i)%></td>
		    <%next%>
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
 
<!-------------akhir---------------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
