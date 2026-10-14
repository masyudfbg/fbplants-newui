<%Option Explicit%>
<%
'On Error resume Next
%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Delivery Order Transfer</title>
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

function validasi(form){
 
	<!-- Start additional Validation by AST 28 Juni 2018 -->
	if (form.DELIVERYORDERNO.value == ""){
		alert("Anda belum mengisikan Delivery Order No.");
		form.DELIVERYORDERNO.focus();
		return (false);
	}
	
	if (form.INTDESTID.value == ""){
	alert("Anda belum mengisikan Internal Destination Id.");
	form.INTDESTID.focus();
	return (false);
	}
	
	if (form.STARTDATE.value == ""){
	alert("Anda belum mengisikan Start Date.");
	form.STARTDATE.focus();
	return (false);
	}

	if (form.PRODUCTCODE.value == ""){
	alert("Anda belum mengisikan Product Code.");
	form.PRODUCTCODE.focus();
	return (false);
	}	
	
	if (form.TOTALWEIGHT.value == ""){
	alert("Anda belum mengisikan Total Weight.");
	form.TOTALWEIGHT.focus();
	return (false);
	}	

	if (form.REFERENCE.value == ""){
	alert("Anda belum mengisikan Reference.");
	form.REFERENCE.focus();
	return (false);
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
 %>
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
''Dim CONTRACTCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, TOTALWEIGHT, txtSTATUS, REFERENCE, CUST_CUSTOMERCODE, PRODUCT_PRODUCTCODE
Dim vDELIVERYORDERNO, vINTDESTID, vSTARTDATE, vPRODUCTCODE, vTOTALWEIGHT, vREFERENCE, vRowstate
Dim txt_DeliveryOrderNo

Dim cmdView '',txt_CONTRACTCODE, TEXTFIELDID
txttable=request("txttable")
response.write Request("txtJenis")

	if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
		
		txt_DeliveryOrderNo = Request("DELIVERYORDERNO")
		vDELIVERYORDERNO = Request("DELIVERYORDERNO")
		vINTDESTID = Request("INTDESTID")
		vSTARTDATE = Request("STARTDATE")
		vPRODUCTCODE = Request("PRODUCTCODE")
		vTOTALWEIGHT = Request("TOTALWEIGHT")
		vREFERENCE = Request("REFERENCE")
		vRowstate = Request("REFERENCE")
		
		if Request("txtJenis")= "Delete" then
			q=" DELETE FROM DeliveryOrderTransfer " & _
			  " WHERE DELIVERYORDERNO='" & vDELIVERYORDERNO & "'" & _
			  " AND INTDESTID='" & vINTDESTID & "'" & _
			  "	AND fcBA='" & Session("fcBA") & "'"
			connStr.execute(q)
			cmdView	= "True"
		   
		elseif Request("txtJenis")="New" then	
		q	=	" INSERT INTO DeliveryOrderTransfer (" & _
				" DELIVERYORDERNO, " & _
				" INTDESTID, " & _
				" STARTDATE, " & _
				" PRODUCTCODE, " & _
				" TOTALWEIGHT, " & _
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
				" , '" & vINTDESTID & "'" & _
				" , to_date('" & Request("STARTDATE") & "','mm/dd/yyyy') " & _
				" , '" & vPRODUCTCODE & "'" & _
				" , '" & vTOTALWEIGHT & "'" & _
				" , '" & vREFERENCE & "'" & _
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
					
		elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		  q="  UPDATE DeliveryOrderTransfer SET " & _
			"  INTDESTID ='" & vINTDESTID & "'" & _
			" ,STARTDATE =to_date('" & Request("STARTDATE") & "','mm/dd/yyyy') " & " " & _
			", PRODUCTCODE ='" & vPRODUCTCODE & "'" & _
			", TOTALWEIGHT ='" & vTOTALWEIGHT & "'" & _
			", REFERENCE ='" & vREFERENCE & "'" & _
			", LASTUPDATE = sysdate " & _
			", LASTTIME = To_Char(sysdate,'HH24:MI') " & _			
			"  WHERE  DELIVERYORDERNO='" & vDELIVERYORDERNO &"'"  & _
			"  AND fcBA='" & Session("fcBA") & "'" 

			'response.Write(q)
			connStr.execute(q)
			
			'response.Flush()
			connStr.execute("commit")
			cmdView	= "True"
		else
			cmdView	= "True"
		end if								 

	else
		cmdView	= request("cmdView")
	End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")

If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if

'' Event GetInitEdit / Mengisi Data
''============================================
if request("txtJenis")="Edit" then
 Dim rsVHx
 
 	q	=	" SELECT " & _
			" a.DELIVERYORDERNO, a.INTDESTID, c.INTDESTNAME, a.STARTDATE, " & _
			" a.PRODUCTCODE, b.PRODUCTDESCRIPTION, a.TOTALWEIGHT, a.REFERENCE " & _			
			" FROM DeliveryOrderTransfer a " & _
			" INNER JOIN PRODUCT b ON b.PRODUCTCODE = a.PRODUCTCODE " & _
			" INNER JOIN IntDestination c ON c.INTDESTID = a.INTDESTID " & _			
			" WHERE a.fcba = '" & Session("fcBA") & "'" & _
			" AND a.DELIVERYORDERNO='" & vDELIVERYORDERNO &"'"
			
''response.Write(q)				
Set rsVHx = Server.CreateObject("ADODB.Recordset")
rsVHx.Open q, connstr

	if not rsVHx.eof then 
		While not rsVHx.eof
			vDELIVERYORDERNO =rsVHx("DELIVERYORDERNO")
			vINTDESTID =rsVHx("INTDESTNAME")
			vSTARTDATE =rsVHx("STARTDATE")
			vPRODUCTCODE =rsVHx("PRODUCTDESCRIPTION")
			vTOTALWEIGHT =rsVHx("TOTALWEIGHT")
			vREFERENCE =rsVHx("REFERENCE")
		rsVHx.movenext
		Wend
		end if
	end if		
 '-------------ISI Payment Form
 %>
 
<form name="frmEntry" method="post" action="DeliveryOrderTransfer.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
           <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE DELIVERY ORDER TRANSFER</strong></font></div></td>
          </TR>
		<tr>
			<td align="left">Delivery Order No</td><td align="left"><input type="text" size="25" name="txt_DeliveryOrderNo" value="<%=REQUEST("txt_DeliveryOrderNo")%>"  ></td>
		</tr>	
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;
				<a href="PrintDeliveryOrderTransfer.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_DeliveryOrderNo=<%=REQUEST("txt_DeliveryOrderNo")%>" 
				title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
			
			</td>
		</tr>
    </table>  
 </form>
  <%
  if cmdView = "True" Then
	Dim rsGAD,intRecordsPerPage,cColor
	''set rsGAD = server.CreateObject("Adodb.recordset")
    ''Field Numeric Jangan di Sum karena ditampilkan mesti level detail

	'' Query Menampilkan Data di Table Bawah  
	q	=	" SELECT " & _
			" DELIVERYORDERNO, INTDESTID, TO_CHAR(STARTDATE, 'mm/dd/yyyy') as STARTDATE, " & _
			" PRODUCTCODE, TOTALWEIGHT, REFERENCE, ROWSTATEID " & _			
			" FROM DeliveryOrderTransfer " & _
			" WHERE fcba = '" & Session("fcBA") & "'"
	
	if  len(trim(request("txt_DeliveryOrderNo"))) > 0 then
		q = q & " AND DELIVERYORDERNO='" & request("txt_DeliveryOrderNo") & "'"
	end if
	q = q		
			
    Set rsGAD = Server.CreateObject("ADODB.Recordset")
      
   ''response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr 
   %>
   
   <form action="DeliveryOrderTransfer.asp" method="post" name="eForm" onSubmit="javascript:return validasi(this)">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">

		<tr>
			<%IF Request("txtJenis")="Edit" then%>
				<td align="left">Delivery Order No</td><td align="left"><input type="text" size="25" maxlength="25" name="DELIVERYORDERNO" value="<%=vDELIVERYORDERNO%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
			<%else%>
				<td align="left">Delivery Order No</td><td align="left"><input type="text" size="25" maxlength="25" name="DELIVERYORDERNO" value="<%=vDELIVERYORDERNO%>"  ><font color="#FF0000">*</font></td>
			<%end if%>
		</tr>
		<tr>
			<%IF txtJenis="Edit" Then%>
				<td align="left">Internal Destination Id</td> 
				<td align="left"> <select name="INTDESTID">
				<%cmdCallComboSetup "INTDESTINATION",request("INTDESTID"),"fcba",session("fcba")%>
				</select> <font color="#FF0000">*</font></td>
			<%Else%>
				<td align="left">Internal Destination Id</td>
				<td align="left"> <select name="INTDESTID">
				<%cmdCallComboSetup "INTDESTINATION","INTDESTIDk","fcba",session("fcba")%>
				</select> <font color="#FF0000">*</font></td>
			<%End if%>
		</tr>
		
		<tr>
			<td align="left">Start Date</td>
			<td align="left"><input class="datepicker" type="text" maxlength="10" name="STARTDATE" size="10" value="<%=vSTARTDATE%>">
			<font color="#FF0000">*</font>
			</td>
		</tr>

		<tr>
			<%IF txtJenis="Edit" Then%>
				<td align="left">Product Code</td>
				<td align="left"> <select name="PRODUCTCODE">
				<%cmdCallComboSetup "PRODUCT",request("PRODUCTCODE"),"fcba",session("fcba")%>
				</select> <font color="#FF0000">*</font> </td>
			<%Else%>
				<td align="left">Product Code</td>
				<td align="left"> <select name="PRODUCTCODE">
				<%cmdCallComboSetup "PRODUCT","PRODUCTCODE","fcba",session("fcba")%>
				</select> <font color="#FF0000">*</font> </td>
			<%End if%>
		</tr>		
	
		<tr>
			<td align="left">Total Weight</td><td align="left"><input type="text" size="15" name="TOTALWEIGHT" value="<%=vTOTALWEIGHT%>"><font color="#FF0000">*</font></td>
		</tr>
				
		<tr>
			<td align="left">Reference</td><td align="left"><input type="text" size="25" maxlength="25" name="REFERENCE" value="<%=vREFERENCE%>" ><font color="#FF0000">*</font></td>
		</tr>		
		
	    <tr>
			<td colspan="2" align="left">
			  <input type="hidden" name="txt_DeliveryOrderNo" value="<%=REQUEST("txt_DeliveryOrderNo")%>">				
			  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			  <%IF Request("txtJenis")="Edit" then%>
				   <input type="hidden" name="txtJenis" value="Edit">
				   <input Type="Submit"  name="Add" value="Update">
			  <%Else%>
				   <input type="hidden" name="txtJenis" value="New">
				   <input Type="Submit"  name="Add" value="Add">
			  <%End if%>
			  <input Type="Reset"  name="Reset" value="Reset">
			  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
   </table>
   </form>
  <!--#include file="validation_show.Asp"--> 
<%
	
	' Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	' howmanyfields=rsGAD.fields.count-1
	
	' JumRec=1
	' intRecordsPerPage = 20
	
	' rsGAD.PageSize = intRecordsPerPage
	' intNumberOfPages = 0
	' intNumberOfPages =  rsGAD.PageCount
	' intRecords = rsGAD.RecordCount
	''Set our AbsolutePage property.
	' if not rsGAD.EOF then
		' If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  ' rsGAD.AbsolutePage = 1
		' Else
		  ' rsGAD.AbsolutePage = Request("iPageNo")
		' End If  
	' end if
	' if Request("iPageNo") = "" then
		' nPage = 1
	' Else
		' nPage = Request("iPageNo")
	' End if
   ' rsGAD.PageSize = intRecordsPerPage
   ' intNumberOfPages = 0
   ' intNumberOfPages =  rsGAD.PageCount
   ' intRecords = rsGAD.RecordCount   
   %>
   <table width="737" class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">ACTION</TD>
	<TD align="left">DELIVERY ORDER NO</TD>	
	<TD align="left">INT DEST ID</TD>	
	<TD align="left">START DATE</TD>	
	<TD align="left">PRODUCT CODE</TD>	
	<TD align="left">TOTAL WEIGHT</TD>	
	<TD align="left">REFERENCE</TD>	
	</tr>

        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-2
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">
			<a href="DeliveryOrderTransfer.asp?DELIVERYORDERNO=<%=rsGAD("DELIVERYORDERNO")%>&INTDESTID=<%=rsGAD("INTDESTID")%>&PRODUCTCODE=<%=rsGAD("PRODUCTCODE")%>&STARTDATE=<%=rsGAD("STARTDATE")%>&txttable=<%=request("txttable")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="DeliveryOrderTransfer.asp?DELIVERYORDERNO=<%=rsGAD("DELIVERYORDERNO")%>&INTDESTID=<%=rsGAD("INTDESTID")%>&PRODUCTCODE=<%=rsGAD("PRODUCTCODE")%>&STARTDATE=<%=rsGAD("STARTDATE")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <% If i = 0 then %>
						<td valign="middle" width="597">
							<a target="_blank" href="DeliveryOrderTransferDet.Asp?DELIVERYORDERNO=<%=rsGAD("DELIVERYORDERNO")%>
							"<%=rsGAD(i).Name%>"><%=rsGAD(i)%>
							</a>						
	                    </td>
				<% else %>
						<td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%>
				</td>
				<% End If%>
		    <%next%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
	</div> 
   </TABLE>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
