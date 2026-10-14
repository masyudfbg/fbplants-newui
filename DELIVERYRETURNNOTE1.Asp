<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Delivery Return Note</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="gen_validatorv2.js" type="text/javascript"></script>
<script language="JavaScript" src="date-picker.js"></script> 
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
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%'-------------awal-----------------------%>
 
 <%
	Dim FCENTRY, FCEDIT, FCIP, FCBA
	dim txt_RETURNCODE,Perioddate
	dim RETURNDATE, RETURN_CODE, DNOTE_DELIVERYNOTECODE, WBT_CHITNO, STORAGE_CODE, QUANTITY, REFERENCE, REMARKS
	dim cmdView,rsDN
	
	txt_RETURNCODE=request("txt_RETURNCODE")
	Perioddate=request ("Perioddate")
	RETURNDATE=Request ("RETURNDATE")
	RETURN_CODE=Request ("RETURN_CODE")
	DNOTE_DELIVERYNOTECODE=Request ("DNOTE_DELIVERYNOTECODE")
	WBT_CHITNO=Request ("WBT_CHITNO")
	STORAGE_CODE=Request ("STORAGE_CODE")
	QUANTITY=Request ("QUANTITY")
	REFERENCE=Request ("REFERENCE")
	REMARKS=Request ("REMARKS")	
	
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	
	if Request("txtJenis")= "Delete" then
       q="Delete from DELIVERYRETURNNOTE where RETURN_CODE='" & RETURN_CODE & "' And RETURNDATE=to_Date('" & RETURNDATE & "','MM/DD/YYYY') And  fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
	
	  q="Insert into DELIVERYRETURNNOTE ( RETURNDATE, RETURN_CODE, DNOTE_DELIVERYNOTECODE, WBT_CHITNO, STORAGE_CODE, QUANTITY, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values (to_Date('" & RETURNDATE & "','MM/DD/YYYY'),'" & RETURN_CODE & "','" & DNOTE_DELIVERYNOTECODE & "','" &  WBT_CHITNO & "'," &_
		"'" & STORAGE_CODE & "','" & QUANTITY & "','" & REFERENCE & "','" & REMARKS & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		connStr.execute(q)
		cmdView	= "True"
		
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	
       q="Update DELIVERYRETURNNOTE set QUANTITY ='" & QUANTITY & "',REFERENCE='" & REFERENCE & "', WBT_CHITNO='" & WBT_CHITNO & "',REMARKS ='" & REMARKS & "'," & _
		"STORAGECODE ='" & STORAGECODE & "' where RETURN_CODE='" & RETURN_CODE & "' And RETURNDATE=to_Date('" & RETURNDATE & "','MM/DD/YYYY') And  fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		connStr.execute("commit")
		else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Request("Perioddate")<>"" then
   Perioddate   =Request("Perioddate")
else
   Perioddate   =textDate(date)
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if
'-------------ISI Payment Form%>
 <form name="frmEntry" method="post" action="DELIVERYRETURNNOTE.asp">
	<table class="data" width="100%" cellpadding="0" cellspacing="0">
		<TR> 
		<td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE DELIVERY RETURN NOTE</strong></font></div></td>
		</TR>

		<tr><td align="left">Date</td>
		<td align="left"><input type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		<a href="javascript:show_calendar('frmEntry.Perioddate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<TR>
		<td align="left">Return Code</td><td align="left"><input type="text" size="15" maxlength="15" name="txt_RETURNCODE" value="<%=txt_RETURNCODE%>"><font color="#FF0000">*</font></td>
		</TR>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;
		<a href="PrintDeliveryReturnNote.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_RETURNCODE=<%=txt_RETURNCODE%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
 
 <%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
    q="select RETURNDATE, RETURN_CODE, DNOTE_DELIVERYNOTECODE, WBT_CHITNO, STORAGE_CODE, QUANTITY, " & _
     "REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA from DELIVERYRETURNNOTE where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(request("Perioddate"))) > 0 then
		q = q & " and to_char(RETURNDATE,'mm/yyyy')=to_char(to_date('" & request("Perioddate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_RETURNCODE"))) > 0 then
		q = q & " and RETURN_CODE='" & request("txt_RETURNCODE") & "'"
	end if
	q = q

   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   
  <form action="DELIVERYRETURNNOTE.asp" method="post" name="eForm">
   <table class="data" width="100%" cellpadding="0" cellspacing="0">
		<tr>
		<td align="left">Date</td>
			<td align="left"><input type="text" maxlength="10" name="RETURNDATE" size="10" value="<%=RETURNDATE%>">
			<a href="javascript:show_calendar('eForm.RETURNDATE');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
		
		</tr>
		<TR>
		<td align="left">RETURN CODE</td><td align="left"><input type="text" size="15" maxlength="15" name="RETURN_CODE" value="<%=RETURN_CODE%>"><font color="#FF0000">*</font></td>
		</TR>	
		
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">DELIVERY NOTE CODE</font></td>
            <td align="left"> <select name="DNOTE_DELIVERYNOTECODE">
                <%cmdCallComboSetup "DELIVERYNOTE",request("DELIVERYNOTECODE"),"",""%>
              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">DELIVERY NOTE CODE</font></td>
				<td align="left"> <select name="DNOTE_DELIVERYNOTECODE">
				<%cmdCallComboSetup "DELIVERYNOTE","DELIVERYNOTECODE","",""%>
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
			
		</tr>
		<tr>
		<td align="left">STORAGE CODE</td><td align="left"><input type="text" size="30" maxlength="50" name="STORAGE_CODE" value="<%=STORAGE_CODE%>"></td>	
		</tr>
		<TR>
		<td align="left">Weighbridge ChitNo</td><td align="left"><input type="text" size="15" maxlength="15" name="WBT_CHITNO" value="<%=WBT_CHITNO%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Quantity (KG)</td><td align="left"><input type="text" size="30" maxlength="50" name="QUANTITY" value="<%=QUANTITY%>"></td>			
		</tr>
		<tr>
		<td align="left">REFERENCE</td><td align="left"><input type="text" size="15" name="REFERENCE" value="<%=REFERENCE%>"></td>
		</tr>
		<tr>
		<td align="left">REMARKS</td><td align="left"><input type="text" size="15" name="REMARKS" value="<%=REMARKS%>"></td>			
		</tr>
			
		 <TR><td colspan="2">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_RETURNCODE" value="<%=txt_RETURNCODE%>">

		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset" style="background-color: #298CF7; color: #F7C610; font-weight:
		  bold; font-size: 8pt; font-family:Verdana" name="Reset" value="Reset">
		  <input Type="Button" style="background-color: #298CF7; color: #F7C610; font-weight: 
		  bold; font-size: 8pt; font-family:Verdana" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
   </form>
  <!--#include file="validation_show.Asp"--> 
  <%
  
	
	Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 20
	
	rsGAD.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  rsGAD.PageCount
	intRecords = rsGAD.RecordCount
	'Set our AbsolutePage property.
	if not rsGAD.EOF then
		If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  rsGAD.AbsolutePage = 1
		Else
		  rsGAD.AbsolutePage = Request("iPageNo")
		End If  
	end if
	if Request("iPageNo") = "" then
		nPage = 1
	Else
		nPage = Request("iPageNo")
	End if

   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
   %>
   <table width="737" class="data" >
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">RETURN CODE</TD>	
	<TD align="left">DELIVERYNOTECODE CODE</TD>	
	<TD align="left">WBT CHIT NO</TD>	
	<TD align="left">STORAGE CODE</TD>	
	<TD align="left">QUATITY</TD>	
	<TD align="left">REFERENCE</TD>	
	<TD align="left">REMARKS</TD>
	</tr>
     <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-5
	  no=1
	  While Not rsGAD.EOF and intRecordsPerPage > 0 
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">

			<a href="DELIVERYRETURNNOTE.asp?RETURN_CODE=<%=rsGAD("RETURN_CODE")%>&txttable=<%=request("txttable")%>&RETURNDATE=<%=rsGAD("RETURNDATE")%>&QUANTITY=<%=rsGAD("QUANTITY")%>&WBT_CHITNO=<%=rsGAD("WBT_CHITNO")%>&STORAGE_CODE=<%=rsGAD("STORAGE_CODE")%>&REFERENCE=<%=rsGAD("REFERENCE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="DELIVERYRETURNNOTE.asp?RETURN_CODE=<%=rsGAD("RETURN_CODE")%>&txttable=<%=request("txttable")%>&RETURNDATE=<%=rsGAD("RETURNDATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
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
	<tr>
	<td colspan="10">
	   <%if intNumberOfPages<>1 then
		  For i = 1 to intNumberOfPages
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""DELIVERYRETURNNOTE.asp?txtModule=" &_
			  txtModule & "&iPageNo=" &  i & "&txtTable=" & request("txtTable")& "&txtField=" & request("txtField")  & _
			  "&txtSearch=" & request("txtSearch") & _
			  """>" & i & "  "
		  Next
	   End if%>
	   </a><br>
	   <%if Request("iPageNo") = "" then%>
			<%nPage = 1%>
	   <%Else%><%'End if%>  
 
  <%'-------------akhir----------------------%>
 </td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
			<%nPage = Request("iPageNo")%>
	   <%End if%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Page : <%=nPage%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%></td>
	</tr>
   </TABLE>
<%End if%>
</html>

