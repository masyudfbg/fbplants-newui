<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<html>
<head>
<title>ePlaS - Upload Purchase Order From SAP</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>--> 
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
 
 Dim xCode,fcCode,cmdView,UPLOADDOCCODE,UPLOADDOCDATE,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim RowStateId
 Dim PRCODE, PRLINEID
 Dim tempStringError, rsTempPurchaseRequest
 Dim totalPRData
 Dim ApprovedQty, ORDEREDQTY
 Dim strQueryStatus, rsQueryStatus 

Dim purchase_doc,item,vendor,matl_group,material,purchase_organization,puchase_group,plant,sloc,doc_date,order_quantity,order_unit,LineNo
Dim net_order_value,ccy,condition_type,amount_cond_type,transporter,pr_no,pr_item,po_release,error_flag,error_desc  
Dim fcfilename

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
 if Request("UPLOADDOCDATE")<>"" then
   UPLOADDOCDATE   =Request("UPLOADDOCDATE")
 else
   UPLOADDOCDATE   =textDate(date)
 end if
 xCode = Request("xCode")
 Remarks = Request("Remarks")
 
 '---- Data Details
 ' Add by NAF, 2012-12-1
 LINENOKEY=Request("LineNoKey")
 ' End Added by NAF

 purchase_doc 			= request("purchase_doc")
 LineNo					= Request("LineNo")
 Item     				= request("Item")
 vendor       			= request("vendor")
 matl_group   			= request("matl_group")
 material     			= request("material")
 purchase_organization  = request("purchase_organization")
 puchase_group          = request("puchase_group")
 plant                  = request("plant")
 sloc                   = request("sloc")
 doc_date               = request("doc_date")
 order_quantity         = request("order_quantity")
 order_unit             = request("order_unit")
 net_order_value		= request("net_order_value")
 ccy					= request("ccy")
 condition_type			= request("condition_type")
 amount_cond_type		= request("amount_cond_type")
 transporter			= request("transporter")
 pr_no					= request("pr_no")
 pr_item				= request("pr_item")
 po_release				= request("po_release")
 error_flag				= request("error_flag")
 error_desc 			= request("error_desc")
 cmdView				= request("cmdView")
 
 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID") 
 'response.write Request("txtJenis")
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Validate" or Request("txtJenis")= "Transfer"  Then
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
       connstr.BeginTrans
	   q="Delete from purchaseorderdetail_sap where  UPLOADDOCCODE='" & fcCode & "' and fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       if err.number = 0 then
	      q="Delete from purchaseorderheader_sap where  UPLOADDOCCODE='" & fcCode & "' and fcBA='" & Session("fcBA") & "'"
   	      connStr.execute(q)
		  if err.number  = 0 then
		     connStr.CommitTrans
		  end if
       end if
	   cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans
	   q="Delete from purchaseorderdetail_sap where  UPLOADDOCCODE='" & fcCode & "' and lineno='" & LineNo & "' and fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
	   if err.number  = 0 then
		  connStr.CommitTrans
	   end if
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
	  Dim rsPO
	  Set rsPO = Server.CreateObject("ADODB.Recordset")
	  connstr.BeginTrans
	  rsPO.Open "Select UPLOADDOCCODE from purchaseorderheader_sap where UPLOADDOCCODE='" & fcCode & "' AND FCBA = '" & Session("FCBA") & "'" , connstr
	  IF rsPO.eof then	   
		 ' Modified by NAF
		 ' To handle first time Input for Purchase Order
	     q="Insert into purchaseorderheader_sap " & _
		 " ( UPLOADDOCCODE, UPLOADDOCDATE, fcfilename, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, ROWSTATEID, LASTUPDATE, LASTTIME ) " & _
		 "values ('" & fcCode & "'," & " to_date('" & UPLOADDOCDATE & "','mm/dd/yyyy') " & ",'"  &  xCode & _
		 "','" & Remarks & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','Planned', sysdate, To_Char(sysdate,'HH24:MI'))"
	     connStr.execute(q)
	  End if
	  rsPO.Close
	  if err.number = 0 then	
	  
	    ' Added by NAF, 2012-12-01
	     strLineMax = " select nvl(max(linenokey),0) maxLineNoKey from purchaseorderdetail_sap where fcBA='" & session("fcBA") & "'"

	  
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
						q="Insert into  purchaseorderdetail_sap (UPLOADDOCCODE, ITEMCODE, QUANTITY, AMOUNT, OTHERITEMDESC, RECEIVED, FCENTRY, FCEDIT, FCIP, FCBA, LINENOKEY, PRCODE, PRLINEID, LASTUPDATE, LASTTIME ) " & _
						 "values ('" & fcCode & "','" & ITEMCODE & "','"  &  Quantity & "','" & AMOUNT & "'" & _
						 ",'" & OTHERITEMDESC & "'," & "0" & ",'" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
						 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLineNo & ",'" & PRCODE & "','"& PRLINEID &"', sysdate, To_Char(sysdate,'HH24:MI'))"
						 
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

				q="Insert into  purchaseorderdetail_sap (UPLOADDOCCODE, ITEMCODE, QUANTITY, AMOUNT, OTHERITEMDESC, RECEIVED, FCENTRY, FCEDIT, FCIP, FCBA, LINENOKEY, PRCODE, PRLINEID, LASTUPDATE, LASTTIME ) " & _
				 "values ('" & fcCode & "','" & ITEMCODE & "','"  &  QUANTITY & "','" & AMOUNT & "'" & _
				 ",'" & OTHERITEMDESC & "'," & "0" & ",'" & Session("CodeName")  & "','" & Session("CodeName") & "','" & _
				 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "',"& tempMaxLineNo & ",'" & PRCODE & "','"& PRLINEID &"', sysdate, To_Char(sysdate,'HH24:MI'))"
				 
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
       q="Update purchaseorderdetail_sap set purchase_doc ='" & purchase_doc & "'" & _										   
		 ", item='" & item & "'" & _
		 ", vendor= '"& vendor & "'" & _
		 ", matl_group = '" & matl_group & "'" & _
		 ", material = '" & material & "'" & _
		 ", purchase_organization = '" & purchase_organization & "'" & _
		 ", puchase_group = '" & puchase_group & "'" & _
		 ", plant = '" & plant & "'" & _
		 ", sloc = '" & sloc & "'" & _
		 ", doc_date = '" & doc_date & "'" & _
		 ", order_quantity = '" & order_quantity & "'" & _
		 ", order_unit = '" & order_unit & "'" & _
		 ", net_order_value = '" & net_order_value & "'" & _
		 ", ccy = '" & ccy & "'" & _
		 ", condition_type = '" & condition_type & "'" & _
		 ", amount_cond_type = '" & amount_cond_type & "'" & _
		 ", transporter = '" & transporter & "'" & _
		 ", pr_no = '" & pr_no & "'" & _
		 ", pr_item = '" & pr_item & "'" & _
		 ", LASTUPDATE = sysdate " & _
		 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
	     " where  UPLOADDOCCODE='" & fcCode & "' and LineNo='" & LineNo & "' and fcBA='" & Session("fcBA") & "'"

		 connStr.execute(q)
		
		if err.number  = 0 then
		   connStr.CommitTrans
		end if
        cmdView	= "True"
	elseif Request("txtJenis")= "Validate" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDERHEADER_SAP_API.ValidateDetail('" & fcCode & "','" & Session("fcBA") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "' ); END; "
	   'response.write strQueryStatus
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	   
	elseif Request("txtJenis")= "Transfer" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEORDERHEADER_SAP_API.TransferDetail('" & fcCode & "','" & Session("fcBA") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "' ); END; "
	   'response.write strQueryStatus
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
	
 else
    cmdView	= request("cmdView")
	'Response.Write " else"
 End if

 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select * from purchaseorderheader_sap  " & _
  " where UPLOADDOCCODE='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    UPLOADDOCDATE=rsPurchase("UPLOADDOCDATE")
	xCode=rsPurchase("filename")
	Remarks=rsPurchase("Remarks")
	RowStateId = rsPurchase("RowStateId")
	fcfilename=xCode
 End if
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if
 '-------------ISI Payment Form 
 %>
 <form name="frmEntry" method="post" action="POUpload.asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Upload Doc Code</td>
		<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">*</font></td>
		</tr>
		
		<tr>
		<td align="left">Upload Doc Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="uploaddocdate" size="10" value="<%=uploaddocdate%>">
		</td>
		</tr>
		
		<tr>
			<td align="left">File Name</td>
			 <td align="left">
			 <input type="text" maxlength="50" name="fcfilename" size="30" value="<%=fcfilename%>">
			 <input Type="Button"  name="Add" value="Upload" onClick="window.location.href='UploadPoFromSap.asp'"></td>
			<tr>
			<td align="left">Remarks</td>
			<td align="left"><textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea></td>		
			</tr>
		</tr>
		
		
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
	    <a href="javascript:launchwin('PrintPO.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&UPLOADDOCDATE=<%=UPLOADDOCDATE%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<a href="POUpload.asp?fcCode=<%=fcCode%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" then %>
		    <a href="POUpload.asp?fcCode=<%=fcCode%>&txtJenis=Validate" onClick="return confirm('Are you sure to Validate this data :  <%=fcCode%> ?');"> Validate </a>
		<% elseif RowStateId = "Validated" then %>	
		    <a href="POUpload.asp?fcCode=<%=fcCode%>&txtJenis=Validate" onClick="return confirm('Are you sure to Validate this data :  <%=fcCode%> ?');"> Validate</a>
		<% elseif RowStateId = "Ready" then %>	
		    <a href="POUpload.asp?fcCode=<%=fcCode%>&txtJenis=Transfer" onClick="return confirm('Are you sure to Transfer this data :  <%=fcCode%> ?');"> Transfer</a>
		<% end if%>
		
		
		
		</td>
		</tr>
    </table>  
 </form>
 
<%if cmdView = "True" AND LEN(TRIM(fcCode)) > 10 Then
   Dim rsPurchaseOrder,intRecordsPerPage,cColor
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
  q="select UPLOADDOCCODE,LineNo,purchase_doc,item,vendor,matl_group,material,purchase_organization,puchase_group,plant,sloc,doc_date,order_quantity,order_unit ," & _
     " net_order_value,ccy,condition_type,amount_cond_type,transporter,pr_no,pr_item,error_flag,error_desc from purchaseorderdetail_sap  " & _
     " where UPLOADDOCCODE='" & fcCode & "' and " & _
     " fcBA='" & session("fcBA") & "'"
  'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   


   
   <form action="POUpload.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
    <%
	
	DIM DISABLED_
	IF Request("txtJenis")="Edit" then
	DISABLED_ = "DISABLED"
	Else
	DISABLED_ = ""
	END IF%>
		  
		<tr>
		   <td  align="left">Purchase Doc</td>
		   <td align="left">
			
				 <input type="text" size="15" name="purchase_doc" value="<%=purchase_doc%>" <%=DISABLED_%>>&nbsp;<font color="#FF0000">
			</td>	 
		</tr>


		<tr>
			<td align="left">Item</td>
			<td align="left"><input type="text" size="15" name="Item" value=<%=Item%>><font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Vendor</td>
			<td align="left"><input type="text" size="15" name="vendor" value=<%=vendor%>><font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Matrial Group</td>
			<td align="left"><input type="text" size="15" name="matl_group" value=<%=matl_group%>>&nbsp;<font color="#FF0000">
		</tr>
       
		<tr>
			<td align="left">Matrial</td>
			<td align="left"><input type="text" size="15" name="material" value=<%=material%>>&nbsp;<font color="#FF0000">
		</tr>

		<tr>
			<td align="left">Purchase Organization</td>
			<td align="left"><input type="text" size="15" name="purchase_organization" value=<%=purchase_organization%>>&nbsp;<font color="#FF0000">
		</tr>
		
		<tr>
			<td align="left">Plant</td>
			<td align="left"><input type="text" size="15" name="plant" value=<%=plant%>>&nbsp;<font color="#FF0000">
		</tr>
		
		<tr>
			<td align="left">S Loc</td>
			<td align="left"><input type="text" size="15" name="sloc" value=<%=sloc%>>&nbsp;<font color="#FF0000">
		</tr>
		
		<tr>
			<td align="left">Document Date</td>
			<td align="left"><input type="text" size="15" name="doc_date" value="<%=doc_date%>" <%=DISABLED_%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>


		<tr>
			<td align="left">Order Qty</td>
			<td align="left"><input type="text" size="15" name="order_quantity" value=<%=order_quantity%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>	
		
		<tr>
			<td align="left">Order Unit</td>
			<td align="left"><input type="text" size="15" name="order_unit" value=<%=order_unit%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>	
		
		
		
		<tr>
			<td align="left">Net Order Value</td>
			<td align="left"><input type="text" size="15" name="net_order_value" value=<%=net_order_value%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>	
		
		<tr>
			<td align="left">Currency </td>
			<td align="left"><input type="text" size="15" name="ccy" value=<%=ccy%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Condition Type</td>
			<td align="left"><input type="text" size="15" name="condition_type" value=<%=condition_type%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Amount Cond Type</td>
			<td align="left"><input type="text" size="15" name="amount_cond_type" value=<%=amount_cond_type%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Transporter</td>
			<td align="left"><input type="text" size="15" name="transporter" value=<%=transporter%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
	     <tr>
			<td align="left">PR. Number.</td>
			<td align="left"><input type="text" size="15" name="pr_no" value=<%=pr_no%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>


		<tr>
			<td align="left">PR. Line Number.</td>
			<td align="left"><input type="text" size="15" name="pr_item" value=<%=pr_item%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>	
		
		
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="doc_date" value="<%=doc_date%>">
		  <input type="hidden" name="purchase_doc" value="<%=purchase_doc%>">
		  <input type="hidden" name="UPLOADDOCDATE" value="<%=UPLOADDOCDATE%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="LineNo" value="<%=LineNo%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       
		  <%End if%>
		  <input Type="Reset" name="Reset" value="Reset">
		  <input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">PURCHASE DOC</TD>	
	<TD align="left">ITEM</TD>	
	<TD align="left">VENDOR</TD>	
	<TD align="left">MATL_GROUP</TD>	
	<TD align="left">MATERIAL</TD>	
	<TD align="left">PURCH_ORG</TD>	
	<TD align="left">PUCHASE_GROUP</TD>
	<TD align="left">PLANT</TD>
	<TD align="left">SLOC</TD>
	<TD align="left">DOC DATE</TD>
	<TD align="left">ORD_QTY</TD>
	<TD align="left">ORD_UNIT</TD>
	<TD align="left">NET_ORD_VAL</TD>
	<TD align="left">CCY</TD>
	<TD align="left">COND_TYPE</TD>
	<TD align="left">AMT_COND_TYPE</TD>
	<TD align="left">TRANSPORTER</TD>
	<TD align="left">PRCODE</TD>	
	<TD align="left">PRLINEID</TD>
	<TD align="left">ERROR_FLAG</TD>
	<TD align="left">ERROR_DESC</TD>
	
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
			qCekClose="Select UPLOADDOCCODE from purchaseorderheader_sap where UPLOADDOCCODE='" & rsPurchaseOrder("UPLOADDOCCODE") & "'" & " and rowstateId = 'Planned' and UPLOADDOCDATE >=to_Date('" & Session("cls_Purchase") & "','mm/dd/yyyy')"
			' response.Write(qCekClose)
			' response.Flush()
			rsCekClose.Open qCekClose, connstr
			if Not rsCekClose.Eof then%>											                                                                                                                                                                                                                                                                                                                                                                                                                                                              
				                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
				<a href="POUpload.asp?fcCode=<%=rsPurchaseOrder("UPLOADDOCCODE")%>&purchase_doc=<%=rsPurchaseOrder("purchase_doc")%>&item=<%=rsPurchaseOrder("LineNo")%>&vendor=<%=rsPurchaseOrder("vendor")%>&material=<%=rsPurchaseOrder("material")%>&matl_group=<%=rsPurchaseOrder("matl_group")%>&purchase_organization=<%=rsPurchaseOrder("purchase_organization")%>&puchase_group=<%=rsPurchaseOrder("puchase_group")%>&plant=<%=rsPurchaseOrder("plant")%>&sloc=<%=rsPurchaseOrder("sloc")%>&doc_date=<%=rsPurchaseOrder("doc_date")%>&order_quantity=<%=rsPurchaseOrder("order_quantity")%>&order_unit=<%=rsPurchaseOrder("order_unit")%>&net_order_value=<%=rsPurchaseOrder("net_order_value")%>&ccy=<%=rsPurchaseOrder("ccy")%>&condition_type=<%=rsPurchaseOrder("condition_type")%>&amount_cond_type=<%=rsPurchaseOrder("amount_cond_type")%>&transporter=<%=rsPurchaseOrder("transporter")%>&pr_no=<%=rsPurchaseOrder("pr_no")%>&pr_item=<%=rsPurchaseOrder("pr_item")%>&error_flag=<%=rsPurchaseOrder("error_flag")%>&error_desc=<%=rsPurchaseOrder("error_desc")%>&LineNo=<%=rsPurchaseOrder("LineNo")%>&txtJenis=Edit">
				<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseOrder(2)%>" border="0"></a></td>
				
				<td valign="middle" width="10"><a href="POUpload.asp?fcCode=<%=rsPurchaseOrder("UPLOADDOCCODE")%>&LineNo=<%=rsPurchaseOrder("LineNo")%>&txtJenis=Details">
				<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseOrder(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a></td>
			<%Else%>	 
			    <td colspan="1"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if
			rsCekClose.close
			set rsCekClose = Nothing%>	 
		    <%for i = 2 to nCountofField%>
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
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->
 
</body>
</html>
