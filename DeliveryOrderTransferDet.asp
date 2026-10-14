<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Delivery Order Transfer Details</title>
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

function validasi(form){
 
	<!-- Start additional Validation by AST 28 Juni 2018 -->
	if (form.DELIVERYORDERNODET.value == ""){
		alert("Delivery Order No Tidak Boleh Kosong.");
		form.DELIVERYORDERNODET.focus();
		return (false);
	}
	
	if (form.TRANSPORTER.value == ""){
		alert("Anda belum mengisikan Transporter.");
		form.TRANSPORTER.focus();
		return (false);
	}
	
	if (form.DOTRANSDET.value == ""){
		alert("Anda belum mengisikan Do Trans.");
		form.DOTRANSDET.focus();
		return (false);
	}
	

	if (form.STARTDATEDET.value == ""){
		alert("Anda belum mengisikan Start Date.");
		form.STARTDATEDET.focus();
		return (false);
	}
	
	if (form.DOWEIGHT.value == ""){
		alert("Anda belum mengisikan Do Weight.");
		form.DOWEIGHT.focus();
		return (false);
	}

	if (form.REFERENCEDET.value == ""){
		alert("Anda belum mengisikan Reference.");
		form.REFERENCEDET.focus();
		return (false);
	}
	
	var nTotalWeight = form.TOTALWEIGHT.value;
	var nDoWeight = form.DOWEIGHT.value;
	var nSumDoWeight = form.nsum_DoWeight.value;
	var nHitungDoWeight = parseFloat(nSumDoWeight)+parseFloat(nDoWeight);
	var nSelisihDoWeight = parseFloat(nTotalWeight)+parseFloat(nSumDoWeight);
	
	if (nHitungDoWeight > nTotalWeight)	{
		 alert("Total Do Weight ( Detail:"+nSumDoWeight+"+"+nDoWeight+"="+nHitungDoWeight+" ) Lebih Besar Dari Total Weight ( Header="+nTotalWeight+" )");
		 form.DOWEIGHT.focus();
		 return(false);
	}
		
	return(true);
}

/*
function Hitung_DO_Weight(){
	
	var nTotalWeight = document.getElementById('TOTALWEIGHT').value;
	var nDoWeight = document.getElementById('DOWEIGHT').value;
	var nSumDoWeight = document.getElementById('nsum_DoWeight').value;
	var nHitungDoWeight = parseFloat(nSumDoWeight)+parseFloat(nDoWeight);
		
	alert(nHitungDoWeight);
	
	if (nHitungDoWeight > nTotalWeight)	{
		 alert("Total Do Weight ( Detail ) Lebih Besar Dari Total Weight ( Header )");
		 document.getElementById('DOWEIGHT').focus();
		 return(false);
	}
	
	return(true);
 }
*/

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
 Dim xDELIVERYORDERNO
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
 
 '' Header
 Dim vDELIVERYORDERNO, vINTDESTID, vSTARTDATE, vPRODUCTCODE, vTOTALWEIGHT, vREFERENCE, vRowstate
 ''Detail
 Dim vDELIVERYORDERNODET, vTRANSPORTER, vDOTRANSDET, vDOWEIGHT, vSTARTDATEDET, vREFERENCEDET

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
 ' if Request("PODATE")<>"" then
   ' PODATE   =Request("PODATE")
 ' else
   ' PODATE   =textDate(date)
 ' end if
 ' contractcode = request("contractcode")
 

 ' if Session("xDELIVERYORDERNO") = "" then
    ' Session("xDELIVERYORDERNO")= DELIVERYORDERNO
 ' end if	
 
 ' startdate = request("startdate")
 ' customercode = request("customercode")
 ' customername = request("customername") 
 ' productcode  = request("productcode")
 ' quantitycontract = request("quantitycontract")

 '-----Data Header
 vDELIVERYORDERNO = Request("DELIVERYORDERNO")
 
 '---- Data Details
 vDELIVERYORDERNODET = Request("DELIVERYORDERNODET")
 vTRANSPORTER = Request("TRANSPORTER")
 vDOTRANSDET = Request("DOTRANSDET")
 vSTARTDATEDET = Request("STARTDATEDET")
 vDOWEIGHT = Request("DOWEIGHT")
 vREFERENCEDET = Request("REFERENCEDET")
 
 ' LINENOKEY=Request("LineNoKey")
 ' ContractHdr = Session("Xcontractcode")
 ' TRANSPORTER=request("TRANSPORTER")
 ' DELIVERYORDER=Request("DELIVERYORDER")
 ' DATEFROM=Request("DATEFROM")
 ' QUANTITYDO =Request("QUANTITYDO")
 ' OTHERITEMDESC=request("OTHERITEMDESC")
 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")="Edit" or Request("txtJenis")= "Delete" Then
	
	REMARKS=request("REMARKS")
	if Request("txtJenis")="New" then
			
			q	=	" INSERT INTO DeliveryOrderTransferDet (" & _
					" DOHEADER, " & _
					" TRANSPORTER, " & _
					" DOTRANS, " & _
					" VALIDUNTIL, " & _
					" DOWEIGHT, " & _
					" REFERENCE, " & _
					" FCENTRY, " & _
					" FCEDIT, " & _
					" FCIP, " & _
					" FCBA, " & _
					" LASTUPDATE, " & _
					" LASTTIME, " & _
					" ROWSTATEID " & _				
					" ) values (" & _
					"   '" & vDELIVERYORDERNO & "'" & _			
					" , '" & vTRANSPORTER & "'" & _
					" , '" & vDOTRANSDET & "'" & _
					" , to_date('" & Request("STARTDATEDET") & "','mm/dd/yyyy') " & _
					" , '" & vDOWEIGHT & "'" & _
					" , '" & vREFERENCEDET & "'" & _
					" , '" & Session("CodeName") & "'" & _					
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _					
					" , '" & Session("fcBA") & "'" & _					
					" , sysdate " & _										
					" , To_Char(sysdate,'HH24:MI') " & _	
					" , '" & vRowstate & "'" & _					
					" )"			 

			'response.write q
			'response.Flush()
			connStr.execute(q)
			cmdView	= "True"
			
			
			' Dim rsPO
			' Set rsPO = Server.CreateObject("ADODB.Recordset")
			' connstr.BeginTrans
			' strLineMax = " select contractcode,CUST_CUSTOMERCODE,customer_api.get_description(cust_customercode,fcba) Custname ," & _
						 ' "		   totalweight , deliverycontract_api.get_TotalQtyDO(fieldid) TotalQtyDO, " & _
						 ' "		   totalweight - deliverycontract_api.get_TotalQtyDO(fieldid) QtyRemain" & _
						 ' "		   from DELIVERYCONTRACT " & _
						 ' "		   where fieldid='" &  Session("XfieldId") & "'"

			' Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
			' rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
			' TemptQtyRemain = 0
			' TemptQtyContract = 0
			' TemptTotalQtyDO = 0
			' if not rsstrLineMax.eof then
			   ' TemptQtyContract = cdbl(rsstrLineMax("totalweight"))
			   ' TemptTotalQtyDO = cdbl(rsstrLineMax("TotalQtyDO"))
			   ' TemptQtyRemain = cdbl(rsstrLineMax("QtyRemain"))
			' end if
								
			' rsstrLineMax.Close
			' set rsstrLineMax = Nothing	
		   '''response.Write " TemptQtyRemain = " & TemptQtyRemain & " , QUANTITYDO = " & QUANTITYDO
			' if cdbl(TemptQtyRemain) >= cdbl(QUANTITYDO) then 
				' q = "Insert into  DELIVERYCONTRACT_DETAIL (DELIVERYCONTRACT , CONTRACTCODE  ,TRANSPORTER ,DELIVERYORDER ,DATEFROM ,QUANTITY, rowstateid ) " & _
					' "values ('" & Session("XfieldId") & "','" & ContractHdr & "','" & TRANSPORTER & "','"  &  DELIVERYORDER & "',to_date('" & STARTDATE & "','mm/dd/yyyy')" & _
					' "," & cdbl(QUANTITYDO) & ", 'Opened')"					 
				'''response.Write  q
				' connStr.execute(q)				
				' connStr.CommitTrans
				' tempStringError = "Delivery Order Saved Successfully with Contract No "	& ContractHdr								
				' response.Write  "<font color=#000000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>" 	
			' Else
				' connStr.RollbackTrans
				' tempStringError = "Quantity yang di input melebihi Sisa Contract No" & ContractHdr & " Sebesar " & TemptQtyRemain
				' response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & tempStringError & "</font>"
			' end if
			'cmdView	= "True"
			
		elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			q  ="  UPDATE DeliveryOrderTransferDet SET " & _
				"  VALIDUNTIL =to_date('" & Request("STARTDATEDET") & "','mm/dd/yyyy') " & " " & _
				", DOWEIGHT ='" & vDOWEIGHT & "'" & _
				", REFERENCE ='" & vREFERENCEDET & "'" & _
				", LASTUPDATE = sysdate " & _
				", LASTTIME = To_Char(sysdate,'HH24:MI') " & _			
				"  WHERE  DOHEADER='" & vDELIVERYORDERNO &"'"  & _
				"  AND fcBA='" & Session("fcBA") & "'" & _
				"  AND TRANSPORTER='" & vTRANSPORTER & "'" & _
				"  AND DOTRANS='" & vDOTRANSDET & "'" 
				

				'response.Write(q)
				connStr.execute(q)
				
				'response.Flush()
				connStr.execute("commit")
				cmdView	= "True"
		
		   ' connstr.BeginTrans
				   ' q="Update DELIVERYCONTRACT_DETAIL set Quantity='" & QuantityDO & "',DATEFROM= to_date('" & STARTDATE & "','mm/dd/yyyy')" & _
					 ' "		   where DELIVERYCONTRACT= '" &  Session("XfieldId") & "' And DELIVERYORDER ='"  &  DELIVERYORDER & "'"				
					 ' connStr.execute(q)
					
					' if err.number  = 0 then
					   ' connStr.CommitTrans
					' end if
			' cmdView	= "True"
		elseif Request("txtJenis")="Delete" then
			q=	" DELETE FROM DeliveryOrderTransferDet " & _
				" WHERE DOHEADER='" & vDELIVERYORDERNO & "'" & _
				" AND TRANSPORTER='" & vTRANSPORTER & "'" & _
				" AND DOTRANS='" & vDOTRANSDET & "'" & _
				" AND fcBA='" & Session("fcBA") & "'"
				  
				''response.Write(q)
				connStr.execute(q)
				cmdView	= "True"
				
		   ' connstr.BeginTrans
					' q = " Delete DELIVERYCONTRACT_DETAIL " & _
						' " Where DELIVERYCONTRACT='" &  Session("XfieldId") & "' And DELIVERYORDER ='"  &  DELIVERYORDER & "'"				
					' connStr.execute(q)		
					' if err.number  = 0 then
						 ' DELIVERYORDER=""
						 ' QUANTITYDO = ""
					   ' connStr.CommitTrans
					' end if
			' cmdView	= "True"	
			
		' elseif Request("txtJenis")= "Confirmed" then
		   ' connstr.BeginTrans
		   ' strQueryStatus = "BEGIN DELIVERYCONTRACT_API.NextState('" & Session("fcBA") & "','" & Session("XfieldId") & "', 'Confirm' ); END; "
		   ' connStr.execute(strQueryStatus)	  
		   ' cmdView	= "True"
		   
		' elseif Request("txtJenis")= "RePlanned" then
		   ' connstr.BeginTrans
		   ' strQueryStatus = "BEGIN DELIVERYCONTRACT_API.NextState('" & Session("fcBA") & "','" & Session("XfieldId") & "', 'RePlan' ); END; "
		   ' connStr.execute(strQueryStatus)	 
		   ' cmdView	= "True"
		   
		' elseif Request("txtJenis")= "CompletedDetails" then
			' connstr.BeginTrans
			' strQueryStatus = " BEGIN" & _
							 ' " DELIVERYCONTRACT_DETAIL_API.NextState('" & Session("fcBA") & "'" & _
							 ' "                                    , '" & Session("XfieldId")  & "'" & _
							 ' "                                    , 'Complete'" & _
							 ' "                                    );" & _
							 ' " END; "
			' connStr.execute(strQueryStatus)	 
			' cmdView	= "True"
		else
			cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
 else
    cmdView	= request("cmdView")
 End if

 ''---- Load Data Header
 Dim rsSQL
 set rsSQL = server.CreateObject("Adodb.recordset")
 	q	=	" SELECT " & _
			" a.DELIVERYORDERNO, a.INTDESTID, a.STARTDATE, " & _
			" a.PRODUCTCODE, a.TOTALWEIGHT, a.REFERENCE " & _			
			" FROM DeliveryOrderTransfer a " & _
			" WHERE a.fcba = '" & Session("fcBA") & "'" & _
			" AND a.DELIVERYORDERNO='" & vDELIVERYORDERNO &"'"

		'response.Write(q)
		rsSQL.Open q, connstr
 
 if Not rsSQL.Eof then
     vDELIVERYORDERNO = rsSQL("DELIVERYORDERNO")
	 vINTDESTID = rsSQL("INTDESTID")
	 vSTARTDATE = rsSQL("STARTDATE")
	 vPRODUCTCODE = rsSQL("PRODUCTCODE") 
	 vTOTALWEIGHT  = rsSQL("TOTALWEIGHT")
	 vREFERENCE = rsSQL("REFERENCE")
 End if
 
 
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="DeliveryOrderTransferDet.Asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Delivery Order No</td>
			<td align="left"><input readonly type="text" maxlength="50" name="DELIVERYORDERNO" size="30" value="<%=vDELIVERYORDERNO%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">Internal Destination Id</td>
			<td align="left"><input disabled type="text" maxlength="50" name="INTDESTID" size="30" value="<%=vINTDESTID%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">Start Date</td>
			<td align="left"><input disabled class="datepicker" type="text" maxlength="10" name="STARTDATE" size="10" value="<%=vSTARTDATE%>">
			</td>
		</tr>
		
		<tr>
			<td align="left">Product Code</td>
			<td align="left"><input disabled type="text" maxlength="50" name="PRODUCTCODE" size="30" value="<%=vPRODUCTCODE%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">Total Weight</td>
			<td align="left"><input disabled type="text" maxlength="50" name="TOTALWEIGHT" size="30" value="<%=vTOTALWEIGHT%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">Reference</td>
			<td align="left"><input disabled type="text" maxlength="50" name="REFERENCE" size="30" value="<%=vREFERENCE%>"><font color="#FF0000"></font></td>
		</tr>
			
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
			</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True"  Then  ''AND LEN(TRIM(contractcode)) > 4
	
	''--- Load Data Detail on Table Detail
	
	Dim rsDeliveryContract,intRecordsPerPage,cColor
	q = " SELECT DOHEADER," & _
	    "        TRANSPORTER, " & _  
		"        DOTRANS, " & _ 
		" 		 VALIDUNTIL," & _
		" 		 DOWEIGHT," & _
		" 		 REFERENCE," & _
		" 		 ROWSTATEID " & _		
	    " FROM DeliveryOrderTransferDet" & _	
		" WHERE DOHEADER = '" & vDELIVERYORDERNO & "'" 

	''Response.Write q	
	Set rsDeliveryContract = Server.CreateObject("ADODB.Recordset")
	rsDeliveryContract.Open q, connstr
%>
      
   <form action="DeliveryOrderTransferDet.Asp" method="post" name="eForm" onSubmit="javascript:return validasi(this)">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
	
		<tr>
			<td align="left">Delivery Order No</td>
			<td align="left">
				<input disabled type="text" size="25" name="DELIVERYORDERNODET" value=<%=vDELIVERYORDERNO%>>&nbsp;<font color="#FF0000"></font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Transporter</td>
			<% 							
				'' " AND fcba='" & Session("fcBA") & "'" & _
				Dim rsAttCodeA, qAttCodeA
				qAttCodeA = "SELECT " & _
							" CUSTOMERCODE, DESCRIPTION " & _
							" FROM CUSTOMER " & _
							" WHERE supplier ='Y' " & _
							" AND istransporter ='1' " & _
							" AND fcba='" & Session("fcBA") & "'" & _
							" ORDER BY CUSTOMERCODE ASC "
				
				Set rsAttCodeA = Server.CreateObject("ADODB.Recordset")
				rsAttCodeA.CursorType = 3
				rsAttCodeA.CursorLocation = 3
				rsAttCodeA.Open qAttCodeA, connstr
				
				%>
					<td align="left">
						<select name="TRANSPORTER">
						<option value=""></option>
						<% if not rsAttCodeA.eof then %>
							<%Do While Not rsAttCodeA.eof%>
								<%if rsAttCodeA("CUSTOMERCODE")=request("TRANSPORTER") Then%>			
								<option value="<%=rsAttCodeA("CUSTOMERCODE")%>" selected><%=rsAttCodeA("CUSTOMERCODE")%> - <%=rsAttCodeA("DESCRIPTION")%></option> 
								<%Else%>
								<option value="<%=rsAttCodeA("CUSTOMERCODE")%>"><%=rsAttCodeA("CUSTOMERCODE")%> - <%=rsAttCodeA("DESCRIPTION")%></option> 
								<%End if%>
								<%			
								rsAttCodeA.movenext
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
			<td align="left">Do Trans</td>
			<td align="left">
				<input type="text" size="25" name="DOTRANSDET" value=<%=vDOTRANSDET%> <%IF txtJenis="Edit" Then%> readonly="True" style="background:#999999"<%End if%>>&nbsp;<font color="#FF0000"></font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Valid Until</td>
			<td align="left"><input class="datepicker" type="text" size="10" name="STARTDATEDET" value=<%=vSTARTDATEDET%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Do Weight</td>
			<td align="left"><input type="text" size="15" name="DOWEIGHT" value=<%=vDOWEIGHT%>>&nbsp;<font color="#FF0000">Kg
			</td>
		</tr>
		
		<tr>
			<td align="left">Reference</td>
			<td align="left">
				<input type="text" size="25" name="REFERENCEDET" value=<%=vREFERENCEDET%>>&nbsp;<font color="#FF0000"></font>
			</td>
		</tr>

	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
		  <input type="hidden" name="DELIVERYORDERNO" value="<%=vDELIVERYORDERNO%>">
		  <input type="hidden" name="TOTALWEIGHT" id="TOTALWEIGHT" value="<%=vTOTALWEIGHT%>">
		  
		  <!--
			  <input type="hidden" name="contractcode" value="<%=contractcode%>">
			  <input type="hidden" name="STARTDATE" value="<%=STARTDATE%>">
			  <input type="hidden" name="xCode" value="<%=xCode%>">
			  <input type="hidden" name="Remarks" value="<%=Remarks%>">
			  <input type="hidden" name="LineNokey" value="<%=LINENOKEY%>">
			  <input type="hidden" name="use_ppn" value="<%=use_ppn%>">
		  -->
		  
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
	<TD align="left">No</TD>
	<TD colspan="2">Action</TD>
	<TD align="left">Delivery Order No </TD>
	<TD align="left">Transporter</TD>
	<TD align="left">DO Trans</TD>
	<TD align="left">Valid Until</TD>
	<TD align="left">Do Weight</TD>
	<TD align="left">Reference</TD>	
	</tr>   
   <%
   dim sum_DoWeight
   if Not rsDeliveryContract.eof then
      nCountofField=rsDeliveryContract.fields.count-1
	  	  
	  sum_DoWeight=0
	  no=1
	  While Not rsDeliveryContract.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if
	%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
	
			<td valign="middle" width="10">
				<a href="DeliveryOrderTransferDet.asp?DELIVERYORDERNO=<%=rsDeliveryContract("DOHEADER")%>&TRANSPORTER=<%=rsDeliveryContract("TRANSPORTER")%>&DOTRANSDET=<%=rsDeliveryContract("DOTRANS")%>&STARTDATEDET=<%=rsDeliveryContract("VALIDUNTIL")%>&DOWEIGHT=<%=rsDeliveryContract("DOWEIGHT")%>&REFERENCEDET=<%=rsDeliveryContract("REFERENCE")%>&txttable=<%=request("txttable")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsDeliveryContract(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
				<a href="DeliveryOrderTransferDet.asp?DELIVERYORDERNO=<%=rsDeliveryContract("DOHEADER")%>&TRANSPORTER=<%=rsDeliveryContract("TRANSPORTER")%>&DOTRANSDET=<%=rsDeliveryContract("DOTRANS")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsDeliveryContract(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
			
			<%for i = 0 to nCountofField-1%>
				<td width="597" align="left" valign="top" 
					title="<%=rsDeliveryContract(i).Name%>"><%=rsDeliveryContract(i)%>
				</td>				
		    <%next%>
		  </tr>
		  <%
		  
		  ''sum_DoWeight = sum_DoWeight + Cdbl(rsDeliveryContract("DOWEIGHT"))
		  sum_DoWeight=sum_DoWeight+trim(cdbl(rsDeliveryContract("DOWEIGHT")))
		  'response.write rsDeliveryContract("DOWEIGHT")
		  'response.write sum_DoWeight
		  
		  rsDeliveryContract.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1

	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsDeliveryContract.Close
   set rsDeliveryContract = Nothing
   %>
   <!-- Sum Do Weight Here, Created By AST, Function: Need For Calculate -->
   <input type="hidden" name="nsum_DoWeight" id="nsum_DoWeight" value="<%=FormatNumber(sum_DoWeight,0)%>">
     
   </div>
   </TABLE>
   </form>
<%End if%>
</div></div>
  <!--#include file="bottom.jsp"-->	
</body>
</html>
