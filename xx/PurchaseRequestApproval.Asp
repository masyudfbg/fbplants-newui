<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Purchase Request Approval</title>
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
 Dim cmdView,PrCode,PrDate
 Dim Description, Remarks, RequestorId
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim ManagedId, PurchaseGroupId, RowStateId
 Dim LineId, ItemCode, RequestedQty, ApprovedQty, OrderedQty
 Dim UnitOfMeasurement, WantedDeliveryDate
 Dim strItemCodeInfo, rsstrItemCodeInfo
 Dim rsPurchase, rsPurchaseTemp
 Dim strQueryStatus, rsQueryStatus
 Dim strQueryAutoNumber
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim IsUseExtSys
 
 


 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 
 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 PRCODE = Request("PRCODE")
 
 '---- Data Header
 if Request("PRDATE")<>"" then
   PRDATE   =Request("PRDATE")
 else
   PRDATE   =textDate(date)
 end if
 Description = Request("Description")
 Remarks = Request("Remarks")
 RequestorId = Request("RequestorId")
 ManagedId = Request("ManagedId")
 PurchaseGroupId = Request("PurchaseGroupId")
 RowStateId = Request("RowStateId")
 
 
 
 '---- Data Details
 
 LINEID=Request("LineId")
 ITEMCODE=Request("ITEMCODE")
 REQUESTEDQTY=Request("REQUESTEDQTY")
 APPROVEDQTY=Request("APPROVEDQTY")
 IF APPROVEDQTY="" then
    APPROVEDQTY="0"
 end if
 IF APPROVEDQTY="0" then
     APPROVEDQTY=REQUESTEDQTY
 END IF

 ORDEREDQTY=Request("ORDEREDQTY")
 WANTEDDELIVERYDATE = Request ("WANTEDDELIVERYDATE")
 UnitOfMeasurement = Request("UnitOfMeasurement")
 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" or Request("txtJenis")= "Approved" or Request("txtJenis")= "ReleaseOnProgress" or Request("txtJenis")= "RefreshPR" 	Then
	REMARKS=request("REMARKS")
	
	Set RsPR = Server.CreateObject("ADODB.Recordset")
	RsPR.Open " select company_api.IsUseExtSys(businessunit_api.get_CompanyCode('" & Session("fcBA") & "')) IsUseExtSys from dual" , connstr
	IF Not RsPR.eof then
	   IsUseExtSys =  RsPR("IsUseExtSys")
	End if
	RsPR.Close
	set RsPR = Nothing	
	
	if Request("txtJenis")= "Header" then
       connstr.BeginTrans
			   q="Delete from PurchaseRequestApproval where  PRCODE='" & PRCODE & "' and fcBA='" & Session("fcBA") & "'"
			   connStr.execute(q)
			   if err.number = 0 then
				  q="Delete from PurchaseRequest where  PRCODE='" & PRCODE & "' and fcBA='" & Session("fcBA") & "'"
				  connStr.execute(q)
				  if err.number  = 0 then
					 connStr.CommitTrans
				  end if
			   end if
	   cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
	   connstr.BeginTrans
			   q="Delete from PurchaseRequestApproval where  PRCODE='" & PRCODE & "' and LineId='" & request("LineId") & "' and fcBA='" & Session("fcBA") & "'"
			   connStr.execute(q)
			   if err.number  = 0 then
				  connStr.CommitTrans
			   end if
       cmdView	= "True"	   
	elseif Request("txtJenis")= "Confirmed" then
	
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEREQUEST_API.NextState('" & Session("fcBA") & "','" & PRCODE & "', 'Confirm' ); END; "
	   connStr.execute(strQueryStatus)	 

       q="Update PurchaseRequestDetails  set rowstateId='Confirmed' where  PRCODE='" & PRCODE & "' and rowstateId='Planned' and fcBA='" & Session("fcBA") & "'"
	   connStr.execute(q)
	   if err.number  = 0 then
			  connStr.CommitTrans
	   end if    
	   
       cmdView	= "True"

	elseif Request("txtJenis")= "RePlanned" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEREQUEST_API.NextState('" & Session("fcBA") & "','" & PRCODE & "', 'RePlan' ); END; "
	   connStr.execute(strQueryStatus)	 
       
	   q="Update PurchaseRequestDetails  set rowstateId='Planned' where  PRCODE='" & PRCODE & "' and rowstateId='Confirmed' and fcBA='" & Session("fcBA") & "'"
	   connStr.execute(q)
	   if err.number  = 0 then
			  connStr.CommitTrans
	   end if

       cmdView	= "True"
	elseif Request("txtJenis")= "Rejected" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN PURCHASEREQUEST_API.NextState('" & Session("fcBA") & "','" & PRCODE & "', 'Reject' ); END; "
	   connStr.execute(strQueryStatus)	 

	   q="Update PurchaseRequestDetails  set rowstateId='Reject' where  PRCODE='" & PRCODE & "' and rowstateId='Confirmed' and fcBA='" & Session("fcBA") & "'"
	   connStr.execute(q)
	   if err.number  = 0 then
			  connStr.CommitTrans
	   end if

       cmdView	= "True"
	   
	elseif Request("txtJenis")= "Approved" then
	   connstr.BeginTrans
	   
		strQueryAutoNumber = " update purchaserequestdetails aa" & _
		                     " set aa.lineid_reference = (select lineno_reference from (select rownum lineno_reference, prcode, lineid, fcba from (select prcode, lineid, fcba from purchaserequestdetails where prcode = '" & PRCODE & "' and fcba = '" & Session("fcBA") & "' order by prcode, lpad(lineid, 2, 0))) cc where cc.prcode = aa.prcode and cc.lineid = aa.lineid and cc.fcba = aa.fcba)" & _
							 " where aa.prcode||aa.lineid||aa.fcba in (select bb.prcode||bb.lineid||bb.fcba from purchaserequestdetails bb where bb.prcode = '" & PRCODE & "' and bb.fcba = '" & Session("fcBA") & "')"
		'Response.Write strQueryAutoNumber
		connStr.execute(strQueryAutoNumber) 
		 Dim rsCekQtyApp,check1
		 set rsCekQtyApp = server.CreateObject("Adodb.recordset")
		 q="select PURCHASEREQUEST_API.CheckBeforeApproved('" & Session("fcBA") & "','" & PRCODE & "') Check1 from Dual  " 
		  
		 rsCekQtyApp.Open q, connstr
		 if Not rsCekQtyApp.Eof then
			Check1=rsCekQtyApp("Check1")
		 End if
		
		If Check1="FALSE" Then
			strQueryStatus = "BEGIN PURCHASEREQUEST_API.NextState('" & Session("fcBA") & "','" & PRCODE & "', 'Approve' ); END; "
			connStr.execute(strQueryStatus)
			
			
			'q="Update PurchaseRequestDetails  set rowstateId='Approved' where  PRCODE='" & PRCODE & "' and rowstateId='Confirmed' and fcBA='" & Session("fcBA") & "'"
		    'connStr.execute(q)
			
			
		    if err.number  = 0 then
				  connStr.CommitTrans
		    end if
			
			
			
	    Else		
	        err.Number = 1001
			err.description = "<br/> Attention :" & Check1 
        End If


       cmdView	= "True"	
	elseif Request("txtJenis") =  "ReleaseOnProgress" then
	    connstr.BeginTrans
	    
        'Response.Write "BEGIN BridgingSAP_PRPOGRN_API.ReleasePr('" & PRCODE & "','" & Session("fcBA") & "'); END; "
	   
	    strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.ReleasePR('" & PRCODE & "','" & Session("fcBA") & "'); END; "
	    connStr.execute(strQueryStatus)	 
        
        cmdView	= "True"    

    elseif Request("txtJenis") =  "RefreshPR" then
	     
		strQueryStatus = "BEGIN BridgingSAP_PRPOGRN_API.UpdateStatusPRFcba('" & PRCODE & "','" & Session("fcBA") & "'); END; "
	    connStr.execute(strQueryStatus)
		
        cmdView	= "True"
	
	elseif Request("txtJenis")="New" then
	  Dim rsPO
	  Set rsPO = Server.CreateObject("ADODB.Recordset")
	  connstr.BeginTrans
			  rsPO.Open "Select prCode from PurchaseRequest where prCode='" & PRCODE & "' and fcba = '" & Session("fcBA") & "'", connstr
			  IF rsPO.eof then	   

				 q="Insert into PURCHASEREQUEST " & _
				 "( PRCODE, DESCRIPTION, PRDATE, REQUESTORID, PURCHASEGROUPID, MANAGEDID, REMARKS, ROWSTATEID, FCENTRY, FCEDIT, FCIP, FCBA ) " & _
				 "values ('" & PRCODE & "','" & Description & "', to_date('" & PrDate & "','mm/dd/yyyy') " & ",'"  &  RequestorId & _
				 "','" & PurchaseGroupId  & "','" & ManagedId & "','" & Remarks & "','Planned','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				 connStr.execute(q)
			  End if
			  rsPO.Close
			  if err.number = 0 then	
			  
				' Added by NAF, 2012-12-01
				 strLineMax = " select max(to_number(nvl(lineid,0))) maxLineIdKey from PurchaseRequestApproval where " & _
								" prcode = '" & PRCODE & "' " & _ 
								" and fcBA='" & session("fcBA") & "'"

			  
				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrLineMax.eof then
					if len(trim(rsstrLineMax("maxLineIdKey"))) <> 0 then 
						tempMaxLineNo = cdbl(rsstrLineMax("maxLineIdKey")) + 1
					else
						tempMaxLineNo = 1
					end if
				end if
							
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 
				' End of Added by NAF, 2012-12-01
			  
				' Get the Information of Material
				strItemCodeInfo = " select * from Stock where " & _
								" fccode = '" & ItemCode & "' " & _ 
								" and fcBA='" & session("fcBA") & "'"

			  
				Set rsstrItemCodeInfo = Server.CreateObject("ADODB.Recordset")
				rsstrItemCodeInfo.Open strItemCodeInfo, connstr, adLockOptimistic, adOpenStatic
				
				if not rsstrItemCodeInfo.eof then
					UnitOfMeasurement = rsstrItemCodeInfo("UnitOfMeasurement")
				end if
							
				rsstrItemCodeInfo.Close
				set rsstrItemCodeInfo = Nothing 
				' End of Added by NAF, 2012-12-01	  
				
			  

			  
				q= "Insert into  PurchaseRequestApproval " & _
					"(PRCODE, LINEID, lineid_reference, ITEMCODE, REQUESTEDQTY, APPROVEDQTY, ORDEREDQTY, UNITOFMEASUREMENT, WANTEDDELIVERYDATE, FCENTRY, FCEDIT, FCIP, FCBA ) " & _
					"values ('" & _
					  PRCODE & "','" & _
					  tempMaxLineNo & "','" & _
					  tempMaxLineNo & "','" & _
					  ITEMCODE & "',"  &  _ 
					  REQUESTEDQTY & ",0,0,'" & _
					  UNITOFMEASUREMENT & "'," & _
					  "to_date('" & WantedDeliveryDate & "','mm/dd/yyyy'),'" & _
					  Session("CodeName")  & "','" & _ 
					  Session("CodeName") & "','" & _
					  Request.ServerVariables("REMOTE_ADDR") &"','" & _
					  Session("fcBA") & "')"
					  
					  
				 
				 'Response.write q
				 connStr.execute(q)
				 
				 
				 
				 if err.number  = 0 then
					APPROVEDQTY = 0
					ORDEREDQTY = 0
					connStr.CommitTrans
				 end if
			  end if
      cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	   connstr.BeginTrans
	            IF APPROVEDQTY="" then
					APPROVEDQTY="0"
				 end if
			   q="Update PurchaseRequestDetails set ApprovedQty =" & ApprovedQty & _
				 ", LASTUPDATE = sysdate " & _
				 ", LASTTIME = To_Char(sysdate,'HH24:MI') "& _
				 " where  PRCode='" & PrCode & "' and LineId='" & LineId & "' and fcBA='" & Session("fcBA") & "'"
				
				'response.write q
				connStr.execute(q)
				if err.number  = 0 then
				   connStr.CommitTrans
				   
					set rsPurchaseTemp = server.CreateObject("Adodb.recordset")
					q=" select * from PurchaseRequestDetails  " & _
					  " where  PRCode='" & PrCode & "' and LineId='" & LineId & "' and fcBA='" & Session("fcBA") & "'"
					
					rsPurchaseTemp.Open q, connstr
					
					if Not rsPurchaseTemp.Eof then
						APPROVEDQTY=rsPurchaseTemp("APPROVEDQTY")
						ORDEREDQTY = rsPurchaseTemp("ORDEREDQTY")
						LINEID=rsPurchaseTemp("LineId")
						ITEMCODE=rsPurchaseTemp("ITEMCODE")
						REQUESTEDQTY=rsPurchaseTemp("REQUESTEDQTY")
						WANTEDDELIVERYDATE = rsPurchaseTemp ("WANTEDDELIVERYDATE")
						UnitOfMeasurement = rsPurchaseTemp("UnitOfMeasurement")
						

					End if
					rsPurchaseTemp.Close
					set rsPurchaseTemp = Nothing 
				   
				end if
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if

 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select * from PURCHASEREQUEST  " & _
  " where PrCode='" & PrCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
  
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    PrDate=rsPurchase("PrDate")
	RequestorId = rsPurchase("RequestorId")
	PurchaseGroupId = rsPurchase("PurchaseGroupId")
	ManagedId = rsPurchase("ManagedId")	
	Remarks=rsPurchase("Remarks")
	RowStateId = rsPurchase("RowStateId")
 End if
' response.write ManagedId
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="PurchaseRequestApproval.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Purchase Request No</td>
		<td align="left"><input type="text" maxlength="50" name="PrCode" size="30" value="<%=PrCode%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="PRDATE" size="10" value="<%=PRDATE%>">
		
		</td>
		</tr>
		
		<tr>
		<td align="left">Description</td>
		<td align="left"><textarea  name="Description"  cols="50" rows="3"><%=Description%></textarea></td>		
		</tr>
		
	
		<tr>
		<td align="left">Requestor</td>
		<td align="left">      
		<select name="RequestorId"  >
	    <%cmdCallCombo "REQUESTOR",RequestorId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>


		
		<tr>
		<td align="left">Managed By</td>
		<td align="left">      
		<select name="ManagedId" >
	    <%cmdCallCombo "PURCHASEMANAGER",ManagedId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>	
		</tr>
		
		<tr>
		<td align="left">Purchase Group By</td>
		<td align="left">      
		<select name="PurchaseGroupId" >
	    <%cmdCallCombo "PURCHASEGROUP",PurchaseGroupId,"",""%>
        </select><font color="#FF0000">*</font>
        </td>	
		</tr>		
		
		
		<tr>
		<td align="left">Remarks</td>
		<td align="left"><textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea></td>		
		</tr>	

		<tr>
		<td align="left">Status</td>
		<td align="left"><input disabled type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>"><font color="#FF0000">*</font>
		</td>
		</tr>

		
		
		</td>
		</tr>

	

		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		
		
		
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('PrintPO.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&PrCode=<%=PrCode%>&PrDate=<%=PrDate%>&Decription=<%=Description%>&RequestorId=<%=RequestorId%>&PurchaseGroupId=<%=PurchaseGroupId%>&ManagedId=<%=ManagedId%>&Remarks=<%=Remarks%>&RowStateId=<%=RowStateId%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" then %>
		    <a href="PurchaseRequestApproval.Asp?PrCode=<%=PrCode%>&PrDate=<%=PrDate%>&Decription=<%=PrDate%>&RequestorId=<%=RequestorId%>&PurchaseGroupId=<%=PurchaseGroupId%>&ManagedId=<%=ManagedId%>&Remarks=<%=Remarks%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to confirm this data :  <%=PrCode%> ?');"> Confirmed </a>&nbsp;&nbsp;
		<% elseif RowStateId = "Confirmed" then %>	
		    <a href="PurchaseRequestApproval.Asp?PrCode=<%=PrCode%>&PrDate=<%=PrDate%>&Decription=<%=PrDate%>&RequestorId=<%=RequestorId%>&PurchaseGroupId=<%=PurchaseGroupId%>&ManagedId=<%=ManagedId%>&Remarks=<%=Remarks%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=PrCode%> ?');"> Re-Planned </a>&nbsp;&nbsp;
		    <a href="PurchaseRequestApproval.Asp?PrCode=<%=PrCode%>&PrDate=<%=PrDate%>&Decription=<%=PrDate%>&RequestorId=<%=RequestorId%>&PurchaseGroupId=<%=PurchaseGroupId%>&ManagedId=<%=ManagedId%>&Remarks=<%=Remarks%>&txtJenis=Rejected" onClick="return confirm('Are you sure to Reject this data :  <%=PrCode%> ?');"> Rejected </a>&nbsp;&nbsp;
		    <a href="PurchaseRequestApproval.Asp?PrCode=<%=PrCode%>&PrDate=<%=PrDate%>&Decription=<%=PrDate%>&RequestorId=<%=RequestorId%>&PurchaseGroupId=<%=PurchaseGroupId%>&ManagedId=<%=ManagedId%>&Remarks=<%=Remarks%>&txtJenis=Approved" onClick="return confirm('Are you sure to Approve this data, Please check your detail data before once Approved..., PR#: <%=PrCode%> ?');"> Approved </a>&nbsp;&nbsp;
		<% elseif RowStateId = "Approved" and ManagedId <> "LOCAL" And IsUseExtSys="Y"  then %>	
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		    <a href="PurchaseRequestApproval.Asp?PrCode=<%=PrCode%>&PrDate=<%=PrDate%>&Decription=<%=PrDate%>&RequestorId=<%=RequestorId%>&PurchaseGroupId=<%=PurchaseGroupId%>&ManagedId=<%=ManagedId%>&Remarks=<%=Remarks%>&txtJenis=ReleaseOnProgress" onClick="return confirm('Are you sure to Release this data :  <%=PrCode%> ?');">Release</a>
		<% elseif RowStateId = "Released" or RowStateId = "ReleaseOnProgress" And IsUseExtSys="Y" then %>	
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		    <a href="PurchaseRequestApproval.Asp?PrCode=<%=PrCode%>&txtJenis=RefreshPR" onClick="return confirm('Refresh this PR No :  <%=PrCode%> ?');">Refresh PR</a>&nbsp;&nbsp;
		<% end if%>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(PrCode)) > 4 Then
   Dim rsPurchaseRequest,intRecordsPerPage,cColor
   set rsPurchaseRequest = server.CreateObject("Adodb.recordset")
   q="select PRCODE, LINEID, ITEMCODE,stock_api.get_stockName(itemcode,fcba) itemname, REQUESTEDQTY, APPROVEDQTY, ORDEREDQTY, UNITOFMEASUREMENT, WANTEDDELIVERYDATE, DESCRIPTION, FCENTRY, FCEDIT, FCIP, FCBA from PurchaseRequestDetails  " & _
     " where PrCode='" & PrCode & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   
   Set rsPurchaseRequest = Server.CreateObject("ADODB.Recordset")
   rsPurchaseRequest.CursorType = 3
   rsPurchaseRequest.CursorLocation = 3
   rsPurchaseRequest.Open q, connstr%>
   <form action="PurchaseRequestApproval.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr><td  align="left">Stock Code</td>
		<td align="left">
		<%IF txtJenis="Edit" Then%>
		     <input disabled type="text" size="15" name="ITEMCODE" value=<%=ITEMCODE%> readonly="True" style="background:#999999">&nbsp;<font color="#FF0000">*
		<%else%>
		     <input disabled type="text" size="15" name="ITEMCODE" value=<%=ITEMCODE%>>&nbsp;<font color="#FF0000">*
		<%End if%>
		<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Stock&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Stock</font></a></td>
		</tr>
		
		
		<tr>
		     <td align="left">Requested Qty</td>
			 <td align="left">
			 <input disabled type="text" size="15" name="REQUESTEDQTY" value= "<%=REQUESTEDQTY%>"  onkeypress="return harusangka(event);">&nbsp;<font color="#FF0000">*</td>
		</tr>
		
	    <tr>
		     <td align="left">Unit of Measure </td>
			 <td align="left">
			 <input  disabled type="text" size="15" name="UNITOFMEASUREMENT" value=<%=UNITOFMEASUREMENT%>   >&nbsp;<font color="#FF0000">*
			 </td>
		</tr>

		
		<tr>
		<td align="left">Wanted Delivery Date</td>
		<td align="left"><input class="datepicker" disabled type="text" maxlength="10" name="WANTEDDELIVERYDATE" size="10" value="<%=WANTEDDELIVERYDATE%>">
		
		</td>
		</tr>
		

	    <tr>
		     <td align="left">Approved Qty</td>
			 <td align="left">
			 <input  type="text" size="15" name="APPROVEDQTY" value="<%=APPROVEDQTY%>"  onkeypress="return harusangka(event);"   >&nbsp;<font color="#FF0000">*
			 </td>
		</tr>
		
	    <tr>
		     <td align="left">Ordered Qty</td>
			 <td align="left">
			 <input  disabled type="text" size="15" name="ORDEREDQTY" value=<%=ORDEREDQTY%>   >&nbsp;<font color="#FF0000">*
			 </td>
		</tr>

	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="PRCODE" value="<%=PRCODE%>">
		  <input type="hidden" name="PRDATE" value="<%=PRDATE%>">
		  <input type="hidden" name="RequestorId" value="<%=RequestorId%>">
		  <input type="hidden" name="PurchaseGroupId" value="<%=PurchaseGroupId%>">
		  <input type="hidden" name="ManagedId" value="<%=ManagedId%>">
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="LineId" value="<%=LINEID%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <% if rowstateId = "Confirmed" then %> 
		      <%IF Request("txtJenis")="Edit" then%>		  
				   <input type="hidden" name="txtJenis" value="Edit">
				   <input Type="Submit" name="Add" value="Update">
			   <% end if %>
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
	<TD align="left">PR CODE</TD>	
	<TD align="left">LINE ID</TD>	
	<TD align="left">ITEM CODE</TD>	
	<TD align="left">ITEM NAME</TD>	
	<TD align="left">REQUESTED QTY</TD>	
	<TD align="left">APPROVED QTY</TD>	
	<TD align="left">ORDERED QTY</TD>	
	<TD align="left">UNIT OF MEASURE</TD>	
	<TD align="left">WANTED DELIVERY DATE</TD>	
	<!--<TD align="left">INSERTED</TD>	
	<TD align="left">UPDATED</TD>	
	<TD align="left">IP ADDRESS</TD>	
	<TD align="left">BA</TD>-->	
	<TD align="left">DESCRIPTION</TD>	
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
		<td align="left" valign="top"><%=NO%></td>
		<%Dim rsCekClose,qCekClose
			Set rsCekClose = Server.CreateObject("ADODB.Recordset")
            rsCekClose.CursorType = 3
            rsCekClose.CursorLocation = 3
			qCekClose="Select PrCode from PurchaseRequest where PrCode='" & rsPurchaseRequest("PrCode") & "'" & " and RowStateId = 'Confirmed' and fcba = '" & session("fcba") & "'"

			rsCekClose.Open qCekClose, connstr
			if Not rsCekClose.Eof then%>
			<td valign="middle" width="10">
				<a href="PurchaseRequestApproval.Asp?PrCode=<%=rsPurchaseRequest("PrCode")%>&LineId=<%=rsPurchaseRequest("LineId")%>&ItemCode=<%=rsPurchaseRequest("ItemCode")%>&fcba=<%=session("fcba")%>&WantedDeliveryDate=<%=rsPurchaseRequest("WantedDeliveryDate")%>&RequestedQty=<%=rsPurchaseRequest("RequestedQty")%>&RequestorId=<%=RequestorId%>&UnitOfMeasurement=<%=rsPurchaseRequest("UnitOfMeasurement")%>&ApprovedQty=<%=rsPurchaseRequest("ApprovedQty")%>&OrderedQty=<%=rsPurchaseRequest("OrderedQty")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsPurchaseRequest(1)%>" border="0">
				</a>
			</td>
			<td valign="middle" width="10">
				<a href="PurchaseRequestApproval.Asp?PrCode=<%=rsPurchaseRequest("PrCode")%>&LineId=<%=rsPurchaseRequest("LineId")%>&fcba=<%=session("fcba")%>&RequestedQty=<%=rsPurchaseRequest("RequestedQty")%>&PRDATE=<%=PRDATE%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsPurchaseRequest(1)%>" border="0">
				</a>
			</td>
			<%Else%>	 
			    <td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if
			rsCekClose.close
			set rsCekClose = Nothing%>	 
		    <%for i = 0 to nCountofField -4%>
			    <td align="left" title="<%=rsPurchaseRequest(i).Name%>" valign="top"><%=rsPurchaseRequest(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsPurchaseRequest.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseRequest.Close
   set rsPurchaseRequest = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
