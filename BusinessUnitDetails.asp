<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Business Unit Details</title>
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
 
 Dim xCode,vfcCode, cmdView,PoCode,PoDate,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
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

 Dim FCCODE, vFCNAME, vFCBA, vEXTSYSTEM, vEXTCOMP, vEXTBU, vDOCTYPE, vPLANT, vSTORLOC, vPURCHORGID, vPURCHORG, vDESCRIPTION
 Dim vUNIQUEKEY

 DIM DELIVERYORDER,DATEFROM,QUANTITYDO,RowStateIdDTL,TRANSPORTER
 DIM TemptQtyContract,TemptTotalQtyDO,TemptQtyRemain

 Dim txtLineNoKey : txtLineNoKey = Request("txtLineNoKey")

 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")

 
 ''Response.Write("Hasil :"+txtJenis)

  '---- Data header
 vFCCODE = request("FCCODE")
 vFCNAME = request("FCNAME")
 
 cmdView = request("cmdView")
 
 '---- Data Details
 vUNIQUEKEY=request("UNIQUEKEY") '' PK
 vFCBA = Request("FCCODE")
 vEXTSYSTEM = Request("EXTSYSTEM")
 vEXTCOMP = Request("EXTCOMP")
 vEXTBU = Request("EXTBU")
 vDOCTYPE = Request("DOCTYPE")
 vPLANT = Request("PLANT")
 vSTORLOC = Request("STORLOC")
 vPURCHORGID = Request("PURCHORGID")
 vPURCHORG = Request("PURCHORG")
 vDESCRIPTION = Request("DESCRIPTION")
 
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
  if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned"  or Request("txtJenis")= "CompletedDetails"  or Request("txtJenis")= "Delete" Then
		if Request("txtJenis")="New" then
			q = " Insert into  BusinessUnitExt(" & _
				" Fcba, " & _
				" ExtSystem, " & _ 
				" ExtComp, " & _
				" ExtBU, " & _
				" doctype, " & _
				" Plant, " & _
				" Storloc, " & _ 
				" PURCHORGID, " & _ 
				" PURCHORG, " & _ 
				" Description, " & _ 
				" UNIQUEKEY, " & _ 
				" FCENTRY, " & _
				" FCEDIT, " & _
				" FCIP, " & _
				" LASTUPDATE, " & _
				" LASTTIME " & _
				" ) " & _
				" values ( " & _
				" 	'" & vFCBA & "' " & _
				"  ,'" & vEXTSYSTEM & "' " & _
				"  ,'" & vEXTCOMP & "' " & _
				"  ,'" & vEXTBU & "' " & _
				"  ,'" & vDOCTYPE & "' " & _
				"  ,'" & vPLANT & "' " & _
				"  ,'" & vSTORLOC & "' " & _
				"  ,'" & vPURCHORGID & "' " & _
				"  ,'" & vPURCHORG & "' " & _
				"  ,'" & vDESCRIPTION & "' " & _
				"  ,'" & vFCBA+vEXTSYSTEM+vEXTCOMP+vEXTBU+vPLANT+vSTORLOC+vPURCHORGID & "' " & _
				" , '" & Session("CodeName") & "'" & _					
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				" , sysdate " & _										
				" , To_Char(sysdate,'HH24:MI') " & _
				"  )" 
				 
			 ''response.Write  q
			 connStr.execute(q)				
		
			cmdView	= "True"
		elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			   q=	"  UPDATE BusinessUnitExt SET" & _ 
					"  EXTSYSTEM ='" & vEXTSYSTEM & "'" & _
					", EXTCOMP ='" & vEXTCOMP & "'" & _
					", EXTBU ='" & vEXTBU & "'" & _
					", DOCTYPE ='" & vDOCTYPE & "'" & _
					", PLANT ='" & vPLANT &"'" & _
					", STORLOC ='" & vSTORLOC & "'" & _
					", PURCHORGID ='" & vPURCHORGID & "'" & _
					", PURCHORG ='" & vPURCHORG & "'" & _
					", DESCRIPTION ='" & vDESCRIPTION & "'" & _
					"  WHERE  UNIQUEKEY='" & vUNIQUEKEY &"'"
					
				connStr.execute(q)
				connStr.execute("commit")
				''response.Write(q)
					
				cmdView	= "True"
		elseif Request("txtJenis")="Delete" then
				q	=	" DELETE FROM BusinessUnitExt WHERE UNIQUEKEY='" & vUNIQUEKEY & "'"
	
				connStr.execute(q)
				''response.Write q
				cmdView	= "True"
		else
			cmdView	= "True"
		end if								 
	' 'Response.Write "sql " & q
  else
     cmdView	= request("cmdView")
  End if

 txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
 ''============================================
 '' Event GetInitEdit / Mengisi Data
 ''============================================
	if request("txtJenis")="Edit" then
		Dim rsDeliveryContract
		q = " select Fcba," & _
			"        ExtSystem, " & _
			" 		 ExtComp," & _
			" 		 ExtBU," & _
			" 		 DocType," & _
			" 		 Plant," & _
			" 		 Storloc," & _
			" 		 PURCHORGID," & _
			" 		 PURCHORG," & _
			" 		 Description," & _
			" 		 fcentry," & _
			" 		 fcedit," & _
			" 		 fcip," & _
			" 		 lastupdate," & _
			" 		 lasttime" & _
			" from BusinessUnitExt" & _
			" where UNIQUEKEY = '" & vUNIQUEKEY & "'"
		
		''response.Write q		
		
		Set rsDeliveryContract = Server.CreateObject("ADODB.Recordset")
		rsDeliveryContract.Open q, connstr
		
		if not rsDeliveryContract.eof then 
			While not rsDeliveryContract.eof
		
			vFCBA = rsDeliveryContract("FCBA")
			vEXTSYSTEM = rsDeliveryContract("EXTSYSTEM")
			vEXTCOMP = rsDeliveryContract("EXTCOMP")
			vEXTBU = rsDeliveryContract("EXTBU")
			vDOCTYPE = rsDeliveryContract("DOCTYPE")
			vPLANT = rsDeliveryContract("PLANT")
			vSTORLOC = rsDeliveryContract("STORLOC")
			vPURCHORGID = rsDeliveryContract("PURCHORGID")
			vPURCHORG = rsDeliveryContract("PURCHORG")
			vDESCRIPTION = rsDeliveryContract("DESCRIPTION")
		 
			rsDeliveryContract.movenext
			Wend
			end if
		end if
	
 '-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="BusinessUnitDetails.Asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">FC CODE</td>
			<td align="left"><input type="text" readonly="true" maxlength="50" name="FCCODE" size="30" value="<%=vFCCODE%>"><font color="#FF0000"></font></td>
		</tr>
		
		<tr>
			<td align="left">FC NAME</td>
			<td align="left"><input type="text" readonly="true" maxlength="50" name="FCNAME" size="30" value="<%=vFCNAME%>"><font color="#FF0000"></font></td>
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
 
 
<% if cmdView = "True" Then

'' Query Menampilkan Data di Table Bawah 
	Dim rsVH,intRecordsPerPage,cColor
		q = " select Fcba, UNIQUEKEY, " & _
			"        ExtSystem, " & _
			" 		 ExtComp," & _
			" 		 ExtBU," & _
			" 		 DocType," & _
			" 		 Plant," & _
			" 		 Storloc," & _
			" 		 PURCHORGID," & _
			" 		 PURCHORG," & _
			" 		 Description," & _
			" 		 fcentry," & _
			" 		 fcedit," & _
			" 		 fcip," & _
			" 		 lastupdate," & _
			" 		 lasttime" & _
			" from BusinessUnitExt" & _
			" where Fcba='" & request("FCCODE") & "'"
			
			' if  len(trim(request("FCCODE"))) <> "" then
				' q = q & " and Fcba='" & request("FCCODE") & "'"
			' end if
				
			''response.Write(q)				
			Set rsVH = Server.CreateObject("ADODB.Recordset")
			rsVH.Open q, connstr   
%>   
   <form action="BusinessUnitDetails.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr> 
			<td width="180" align="left"><input type="hidden" name="UNIQUEKEY" id="UNIQUEKEY" value="<%=vUNIQUEKEY%>" readonly="True" style="background:#999999"></td>
		</tr>
		<tr>
			<td align="left">FCBA</td>
			<td align="left"><input type="text" readonly="true" size="15" name="FCBA" value=<%=vFCBA%>>&nbsp;<font color="#FF0000"></font>
			</td>
		</tr>
	
		<tr>
			<td align="left">External System</td>
			<td align="left">
				<select name="ExtSystem">
					<option value=""></option>
					<% 
						DIM cffbT,rscffbT,SELECTED
						cffbT="select fccode, fcDesc from ExternalSystem"
						Set rscffbT = Server.CreateObject("ADODB.Recordset")
						'response.Write(cffbT)
						'response.Flush()
						
						rscffbT.CursorType = 3
						rscffbT.CursorLocation = 3
						rscffbT.Open cffbT, connstr 
						
						if Not rscffbT.eof then
							While Not rscffbT.EOF
								IF vExtSystem = rscffbT("fccode") THEN
									SELECTED = "SELECTED"
								ELSE
									SELECTED = ""
								END IF
					%>
							<option value="<%=rscffbT("fccode")%>" <%=SELECTED%>><%=rscffbT("fcDesc")%></option>
					<%
							rscffbT.MoveNext
							Wend
						end if 
						rscffbT.Close
						set rscffbT = Nothing
					%>
				</select>
			</td>
			
		</tr>
		
		<tr>
			<td align="left">External Comp</td>
			<td align="left"><input type="text" size="15" name="EXTCOMP" value=<%=vEXTCOMP%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">External BU</td>
			<td align="left"><input type="text" size="15" name="EXTBU" value=<%=vEXTBU%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>

		<tr>
			<td align="left">Doc Type</td>
			<td align="left"><input type="text" size="15" name="doctype" value=<%=vDOCTYPE%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>	
		
		<tr>
			<td align="left">Plant</td>
			<td align="left"><input type="text" size="15" name="PLANT" value=<%=vPLANT%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>	

		<tr>
			<td align="left">Storloc</td>
			<td align="left"><input type="text" size="15" name="STORLOC" value=<%=vSTORLOC%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>

		<tr>
			<td align="left">Purch Org ID</td>
			<td align="left"><input type="text" size="15" name="PURCHORGID" value=<%=vPURCHORGID%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Purch Org</td>
			<td align="left"><input type="text" size="15" name="PURCHORG" value=<%=vPURCHORG%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>

		<tr>
			<td align="left">Description</td>
			<td align="left"><input type="text" size="15" name="DESCRIPTION" value=<%=vDESCRIPTION%>>&nbsp;<font color="#FF0000">
			</td>
		</tr>		
		
		
	    <tr><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
		  
	      <input type="Hidden" readonly="true"  name="FCCODE" value="<%=vFCCODE%>">
		  <input type="Hidden" readonly="true"  name="FCNAME" value="<%=vFCNAME%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="Hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="Hidden" name="txtJenis" value="New">
		       <input Type="Submit" name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset" name="Reset" value="Reset">
		  <input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<br>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
			<td align="left">NO</td>
			<td colspan="2">Action</td>
			<td align="left">Fcba</td>
			<td align="left">UNIQUEKEY</td>
			<td align="left">External System</td>
			<td align="left">External Comp</td>
			<td align="left">External BU</td>
			<td align="left">Doc Type</td>
			<td align="left">Plant</td>
			<td align="left">Storloc</td>
			<td align="left">Purch Org ID</td>			
			<td align="left">Purch Org</td>			
			<td align="left">Description</td>
		</tr>   
   <%
   
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-6
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			
			<%
			if Not rsVH.Eof then%>
				<td valign="middle" width="10">
					<a href="BusinessUnitDetails.Asp?UNIQUEKEY=<%=rsVH("UNIQUEKEY")%>&FCCODE=<%=vFCCODE%>&FCNAME=<%=vFCNAME%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
				</td>
				<td valign="middle" width="10">
					<a href="BusinessUnitDetails.Asp?UNIQUEKEY=<%=rsVH("UNIQUEKEY")%>&FCCODE=<%=vFCCODE%>&FCNAME=<%=vFCNAME%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
				</td>
			<%Else%>	 
			    <td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if%>
			
			<!-- Pengisian Data -->
			<%for i = 0 to nCountofField%>
				<% if rsVH(i).Name="UNIQUEKEY1" Then %>  <!-- Untuk ID sebagai PK, tidak perlu ditampilkan -->
				<%Else %>
					<td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
				<%End If%>
			<%next%>
										
			</tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      'response	.Write("Data is empty")	  
	  response.Write("<td align=""center"" colspan=""9"" title="" valign=""top""> Data Is Empty</td>")	 
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
   
   
<%End if%>
</div></div>
  <!--#include file="bottom.jsp"-->
</body>
</html>
