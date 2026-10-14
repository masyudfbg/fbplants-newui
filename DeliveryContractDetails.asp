<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Delivery Contract Details</title>
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
 Dim Xcontractcode
 Dim Remarks
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax
 Dim RowStateId
 Dim PRCODE, PRLINEID
 Dim tempStringError, rsTempPurchaseRequest
 Dim totalPRData
 Dim ApprovedQty, ORDEREDQTY
 Dim strQueryStatus, rsQueryStatus
 Dim sqlupdateqtypr, rsupdateqtypr
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = "" 

 DIM DELIVERYORDER,DATEFROM,QUANTITYDO,RowStateIdDTL,TRANSPORTER
 DIM TemptQtyContract,TemptTotalQtyDO,TemptQtyRemain

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
 vcontractcode = request("contractcode")
 vstartdate = request("startdate")
 vcustomercode = request("customercode")
 vcustomername = request("customername") 
 vproductcode  = request("productcode")
 vquantitycontract = request("quantitycontract")

 
 '---- Data Details
 ' Add by NAF, 2012-12-1
 LINENOKEY=Request("LineNoKey")
 ' End Added by NAF
 vContractHdr = vcontractcode
 vTRANSPORTER=request("TRANSPORTER")
 vDELIVERYORDER=Request("DELIVERYORDER")
 vDATEFROM=Request("DATEFROM")
 vQUANTITYDO =Request("QUANTITYDO")
 vOTHERITEMDESC=request("OTHERITEMDESC")
 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned"  or Request("txtJenis")= "CompletedDetails"  or Request("txtJenis")= "Delete" Then
	REMARKS=request("REMARKS")
	if Request("txtJenis")="New" then
	    Dim rsPO
	    Set rsPO = Server.CreateObject("ADODB.Recordset")
	    connstr.BeginTrans
		strLineMax = " select contractcode,CUST_CUSTOMERCODE,customer_api.get_description(cust_customercode,fcba) Custname ," & _
					 "		   totalweight , deliverycontract_api.get_TotalQtyDO(fieldid) TotalQtyDO, " & _
					 "		   totalweight - deliverycontract_api.get_TotalQtyDO(fieldid) QtyRemain" & _
					 "		   from DELIVERYCONTRACT " & _
					 "		   where fieldid='" &  Session("XfieldId") & "'"
					 
        'response.write strLineMax
		Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
		rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
		TemptQtyRemain = 0
		TemptQtyContract = 0
		TemptTotalQtyDO = 0
		if not rsstrLineMax.eof then
		   TemptQtyContract = cdbl(rsstrLineMax("totalweight"))
		   TemptTotalQtyDO = cdbl(rsstrLineMax("TotalQtyDO"))
		   TemptQtyRemain = cdbl(rsstrLineMax("QtyRemain"))
		end if
							
		rsstrLineMax.Close
		set rsstrLineMax = Nothing	
        response.Write " TemptQtyRemain = " & TemptQtyRemain & " , QUANTITYDO = " & vQUANTITYDO
        if cdbl(TemptQtyRemain) >= cdbl(vQUANTITYDO) then 
		    q = "Insert into  DELIVERYCONTRACT_DETAIL (DELIVERYCONTRACT , CONTRACTCODE  ,TRANSPORTER ,DELIVERYORDER ,DATEFROM ,QUANTITY, rowstateid ) " & _
			    "values ('" & Session("XfieldId") & "','" & vContractHdr & "','" & vTRANSPORTER & "','"  &  vDELIVERYORDER & "',to_date('" & vSTARTDATE & "','mm/dd/yyyy')" & _
			    "," & cdbl(vQUANTITYDO) & ", 'Opened')"					 
			'response.Write  q
			connStr.execute(q)				
			connStr.CommitTrans
			tempStringError = "Delivery Order Saved Successfully with Contract No "	& vContractHdr								
			response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>" 	
		Else
		    connStr.RollbackTrans
			tempStringError = "Quantity yang di input melebihi Sisa Contract No" & vContractHdr & " Sebesar " & TemptQtyRemain
			response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
		end if
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	   connstr.BeginTrans
			   q="Update DELIVERYCONTRACT_DETAIL set Quantity='" & vQuantityDO & "',DATEFROM= to_date('" & vSTARTDATE & "','mm/dd/yyyy')" & _
				 "		   where DELIVERYCONTRACT= '" &  Session("XfieldId") & "' And DELIVERYORDER ='"  &  vDELIVERYORDER & "'"				
				 connStr.execute(q)
				
				if err.number  = 0 then
				   connStr.CommitTrans
				end if
        cmdView	= "True"
	elseif Request("txtJenis")="Delete" then
	   connstr.BeginTrans
			    q = " Delete DELIVERYCONTRACT_DETAIL " & _
				    " Where DELIVERYCONTRACT='" &  Session("XfieldId") & "' And DELIVERYORDER ='"  &  vDELIVERYORDER & "'"				
				connStr.execute(q)		
				if err.number  = 0 then
				     DELIVERYORDER=""
					 QUANTITYDO = ""
				   connStr.CommitTrans
				end if
        cmdView	= "True"	
	elseif Request("txtJenis")= "Confirmed" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN DELIVERYCONTRACT_API.NextState('" & Session("fcBA") & "','" & Session("XfieldId") & "', 'Confirm' ); END; "
	   connStr.execute(strQueryStatus)	  
       cmdView	= "True"
	elseif Request("txtJenis")= "RePlanned" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN DELIVERYCONTRACT_API.NextState('" & Session("fcBA") & "','" & Session("XfieldId") & "', 'RePlan' ); END; "
	   connStr.execute(strQueryStatus)	 


       cmdView	= "True"
	   
	elseif Request("txtJenis")= "CompletedDetails" then
		connstr.BeginTrans
		strQueryStatus = " BEGIN" & _
						 " DELIVERYCONTRACT_DETAIL_API.NextState('" & Session("fcBA") & "'" & _
	                     "                                    , '" & Session("XfieldId")  & "'" & _
	                     "                                    , 'Approve'" & _
	                     "                                    );" & _
	                     " END; "
		response.Write strQueryStatus
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
 q= "select CONTRACTCODE, CUST_CUSTOMERCODE,customer_api.get_description(cust_customercode,fcba) Custname, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, " & _
    "PRODUCT_PRODUCTCODE, TOTALWEIGHT, RowStateId, REFERENCE, FCENTRY, FCEDIT, FCIP, FCBA,FIELDID from DELIVERYCONTRACT  " & _
    " where contractcode='" & vcontractcode & "' and " & _
    " fcBA='" & session("fcBA") & "'"
 'response.write q	
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
     vcontractcode = rsPurchase("contractcode")
	 vstartdate = rsPurchase("startdate")
	 vcustomercode = rsPurchase("CUST_CUSTOMERCODE")
	 vcustomername = rsPurchase("Custname") 
	 vproductcode  = rsPurchase("PRODUCT_PRODUCTCODE")
	 vquantitycontract = rsPurchase("TOTALWEIGHT")
	 Session("XfieldId")= rsPurchase("FIELDID")
	 cmdView = "True"
 End if
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="DeliveryContractDetails.Asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Contract No</td>
			<td align="left"><input  type="text" maxlength="50" name="contractcode" size="30" value="<%=vcontractcode%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">Start Date</td>
			<td align="left"><input disabled class="datepicker" type="text" maxlength="10" name="STARTDATE" size="10" value="<%=vSTARTDATE%>">
			</td>
		</tr>
		
		<tr>
			<td align="left">Customer Code</td>
			<td align="left"><input disabled type="text" maxlength="20" name="customercode" size="15" value="<%=vcustomercode%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">Customer Name</td>
			<td align="left"><input disabled type="text" maxlength="50" name="customername" size="50" value="<%=vcustomername%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">Product Code</td>
			<td align="left"><input disabled type="text" maxlength="5" name="productcode" size="5" value="<%=vproductcode%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">Quantity Contract</td>
			<td align="left"><input disabled type="text" maxlength="10" name="quantitycontract" size="10" value="<%=vquantitycontract%>"><align="Right" font color="#FF0000"> Kg</font></td>
		</tr>
		
		<tr>
		<td align="left">Status</td>
		<td align="left"><input disabled type="text" maxlength="30" name="RowStateId" size="30" value="<%=RowStateId%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" AND LEN(TRIM(vcontractcode)) > 4 Then
	Dim rsDeliveryContract,intRecordsPerPage,cColor

	q = " select deliverycontract," & _
	    "        TRANSPORTER, " & _
		" 		 contractcode," & _
		" 		 deliveryorder," & _
		" 		 datefrom," & _
		" 		 quantity," & _
		" 		 RowStateId" & _
	    " from DELIVERYCONTRACT_DETAIL" & _
		" where deliverycontract = '" & Session("XfieldId") & "'" 
	'response.Write q	
	Set rsDeliveryContract = Server.CreateObject("ADODB.Recordset")
	rsDeliveryContract.CursorType = 3
	rsDeliveryContract.CursorLocation = 3
	rsDeliveryContract.Open q, connstr
%>
   


   
   <form action="DeliveryContractDetails.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		
		<tr>
			<td align="left">Contract Hdr</td>
			<td align="left"><input disabled type="text" size="25" name="CONTRACTHDR" value=<%=vCONTRACTHDR%>>&nbsp;<font color="#FF0000"></font>
			</td>
		</tr>

		<tr>
			<td align="left">Transporter </td>
			<%
				Dim rsAttCodeA, qAttCodeA
				qAttCodeA = "select CUSTOMERCODE, DESCRIPTION from  CUSTOMER  where supplier ='Y' and istransporter ='1' and fcba='" & Session("fcBA") & "'  order by CUSTOMERCODE ASC"
				
				Set rsAttCodeA = Server.CreateObject("ADODB.Recordset")
				rsAttCodeA.CursorType = 3
				rsAttCodeA.CursorLocation = 3
				rsAttCodeA.Open qAttCodeA, connstr
				if not rsAttCodeA.eof then
				%>
					<td align="left">
						<select name="TRANSPORTER">
							<option value=""></option>
						<%Do While Not rsAttCodeA.eof%>
							<%if rsAttCodeA("CUSTOMERCODE")=request("TRANSPORTER") Then%>			
							<option value="<%=rsAttCodeA("CUSTOMERCODE")%>" selected><%=rsAttCodeA("CUSTOMERCODE")%> - <%=rsAttCodeA("DESCRIPTION")%></option> 
							<%Else%>
							<option value="<%=rsAttCodeA("CUSTOMERCODE")%>"><%=rsAttCodeA("CUSTOMERCODE")%> - <%=rsAttCodeA("DESCRIPTION")%></option> 
							<%End if%>
				<%			rsAttCodeA.movenext
						loop
					End if
					rsAttCodeA.Close
					set rsAttCodeA = Nothing
					qAttCodeA = ""
				%>
						</select>
						<font color="#FF0000">*</font>
		    </td>
			
		</tr>
		
		<tr>
			<td align="left">Delivery Order</td>
			<td align="left"><input type="text" size="15" name="DELIVERYORDER" value=<%=vDELIVERYORDER%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Date From</td>
			<td align="left"><input class="datepicker" type="text" size="10" name="DATEFROM" value=<%=vDATEFROM%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Quantity DO</td>
			<td align="left"><input type="text" size="15" name="QUANTITYDO" value=<%=vQUANTITYDO%>>&nbsp;<font color="#FF0000">Kg
			</td>
		</tr>
		
		<tr>
			<td align="left">Status</td>
			<td align="left"><input disabled type="text" size="15" name="RowStateIdDTL" value=<%=request("RowStateIdDTL")%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		
		
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	      <input type="hidden" name="contractcode" value="<%=vcontractcode%>">
		  <input type="hidden" name="STARTDATE" value="<%=vSTARTDATE%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="LineNokey" value="<%=LINENOKEY%>">
		  <input type="hidden" name="use_ppn" value="<%=use_ppn%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
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
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">Action</TD>
	<TD align="left">Contract</TD>
	<TD align="left">Transporter</TD>
	<TD align="left">Delivery Order</TD>
	<TD align="left">Date From</TD>
	<TD align="left">Qauntity</TD>
	<TD align="left">Status</TD>
	
	</tr>   
   <%
   if Not rsDeliveryContract.eof then
      nCountofField=rsDeliveryContract.fields.count-1
	  no=1
	  While Not rsDeliveryContract.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="10">
			<%
			if UCase(rsDeliveryContract("RowStateId")) = UCase("Opened") then%>
				<a href="DeliveryContractDetails.Asp?contractcode=<%=rsDeliveryContract("contractcode")%>&Transporter=<%=rsDeliveryContract("Transporter")%>&deliveryorder=<%=rsDeliveryContract("deliveryorder")%>&datefrom=<%=rsDeliveryContract("datefrom")%>&QUANTITYDO=<%=rsDeliveryContract("quantity")%>&RowStateIdDTL=<%=rsDeliveryContract("RowStateId")%>&txtJenis=Edit">
				<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsDeliveryContract(1)%>" border="0"></a></td>
				<td valign="middle" width="10"><a href="DeliveryContractDetails.Asp?contractcode=<%=rsDeliveryContract("contractcode")%>&deliveryorder=<%=rsDeliveryContract("deliveryorder")%>&datefrom=<%=rsDeliveryContract("datefrom")%>&QUANTITYDO=<%=rsDeliveryContract("quantity")%>&RowStateIdDTL=<%=rsDeliveryContract("RowStateId")%>&txtJenis=Delete">
				<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsDeliveryContract(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a></td>
			<%Else%>	 
			    <td colspan="1"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if
			
			%>
			
				<%  response.write UCase(rsDeliveryContract("RowStateId"))
				
				if UCase(rsDeliveryContract("RowStateId")) = UCase("Opened") then %>
				<td align="left" title="<%=rsDeliveryContract("rowstateid").Name%>" valign="top">
				<a href="DeliveryContractDetails.Asp?contractcode=<%=rsDeliveryContract("contractcode")%>&txtJenis=CompletedDetails" onClick="return confirm('Are you sure to Complete this data : Contract Code: <%=contractcode%>;">Approve</a>
				</td>
				<% else%>	
				<td align="left"  valign="top">
					
				</td>
				<% end if%>
				
				<td align="left" title="<%=rsDeliveryContract("contractcode").Name%>" valign="top">
					<%=rsDeliveryContract("contractcode")%>
				</td>
				<td align="left" title="<%=rsDeliveryContract("Transporter").Name%>" valign="top">
					<%=rsDeliveryContract("Transporter")%>
				</td>
				<td align="left" title="<%=rsDeliveryContract("deliveryorder").Name%>" valign="top">
					<%=rsDeliveryContract("deliveryorder")%>
				</td>
				<td align="left" title="<%=rsDeliveryContract("datefrom").Name%>" valign="top">
					<%=rsDeliveryContract("datefrom")%>
				</td>
				<td align="left" title="<%=rsDeliveryContract("quantity").Name%>" valign="top">
					<%=rsDeliveryContract("quantity")%>
				</td>
				<td align="left" title="<%=rsDeliveryContract("RowStateId").Name%>" valign="top">
					<%=rsDeliveryContract("RowStateId")%>
				</td>
			</tr>
		  <%
		  rsDeliveryContract.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsDeliveryContract.Close
   set rsDeliveryContract = Nothing%>
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
