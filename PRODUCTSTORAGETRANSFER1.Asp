<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Product Storage Transfer</title>
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
Dim TRANSFERNO, TRANSFERDATE, PSTORAGE_STORAGECODE, PSTORAGE_STORAGECODE_TO_TANK, QUANTITY
Dim cmdView,rsPT
dim Perioddate,txt_TRANSFERNO
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	Perioddate=request("Perioddate")
	txt_TRANSFERNO=request("txt_TRANSFERNO")
	TRANSFERNO=request("TRANSFERNO")
	TRANSFERDATE=request("TRANSFERDATE")
	PSTORAGE_STORAGECODE=request("PSTORAGE_STORAGECODE")
	PSTORAGE_STORAGECODE_TO_TANK=request("PSTORAGE_STORAGECODE_TO_TANK")
	QUANTITY=request("QUANTITY")

	if Request("txtJenis")= "Delete" then
       q="Delete from PRODUCTSTORAGETRANSFER where TRANSFERNO='" & TRANSFERNO & "' And TRANSFERDATE=to_Date('" & TRANSFERDATE & "','MM/DD/YYYY') And PSTORAGE_STORAGECODE='" & PSTORAGE_STORAGECODE & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
		  q="Insert into PRODUCTSTORAGETRANSFER (TRANSFERNO, TRANSFERDATE, PSTORAGE_STORAGECODE, PSTORAGE_STORAGECODE_TO_TANK, QUANTITY,FCENTRY, FCEDIT, FCIP, FCBA) " & _
				"values ('" & TRANSFERNO & "',to_date('" & Request("TRANSFERDATE") & "','mm/dd/yyyy') " & ",'" &  PSTORAGE_STORAGECODE & "','" &  PSTORAGE_STORAGECODE_TO_TANK & "','" &  QUANTITY & "'," & _
				"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				'response.Write(q)
				'response.Flush()
				connStr.execute(q)
				cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then

       q="Update PRODUCTSTORAGETRANSFER set TRANSFERNO ='" & TRANSFERNO & "',TRANSFERDATE =to_date('" & TRANSFERDATE & "','mm/dd/yyyy')," & _
		"PSTORAGE_STORAGECODE='" & PSTORAGE_STORAGECODE & "'," & _
		"PSTORAGE_STORAGECODE_TO_TANK ='" & PSTORAGE_STORAGECODE_TO_TANK & "',QUANTITY ='" & QUANTITY & "'" & _
		"where TRANSFERNO='" & TRANSFERNO & "' And TRANSFERDATE=to_Date('" & TRANSFERDATE & "','MM/DD/YYYY') And PSTORAGE_STORAGECODE='" & PSTORAGE_STORAGECODE & "' And fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		'response.Write(q)
		'response.Flush()
		connStr.execute("commit")
		'RESPONSE.Write "SQLnsddsdya : " & Q
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
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

 '-------------ISI Payment Form%>
 <form name="frmEntry" method="post" action="PRODUCTSTORAGETRANSFER.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
           <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE PRODUCT STORAGE TRANSFER</strong></font></div></td>
          </TR>
		<tr><td align="left">Date</td>
		<td align="left"><input type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		<a href="javascript:show_calendar('frmEntry.Perioddate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>	
		<tr>
		<td align="left">Transfer No</td><td align="left"><input type="text" size="15" maxlength="15" name="txt_TRANSFERNO" value="<%=REQUEST("txt_TRANSFERNO")%>"></td>
		</tr>

		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintPRODUCTSTORAGETRANSFER.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_STORAGECODE=<%=REQUEST("txt_STORAGECODE")%>&txt_ADJUSTMENTNO=<%=REQUEST("txt_ADJUSTMENTNO")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>

		</td>
		</tr>
    </table>  
 </form>
 <%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select TRANSFERDATE, TRANSFERNO, PSTORAGE_STORAGECODE, PSTORAGE_STORAGECODE_TO_TANK, QUANTITY, FCENTRY, FCEDIT, FCIP, FCBA from PRODUCTSTORAGETRANSFER where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(request("Perioddate"))) > 0 then
		q = q & " and to_char(TRANSFERDATE,'mm/yyyy')=to_char(to_date('" & request("Perioddate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_TRANSFERNO"))) > 0 then
		q = q & " and TRANSFERNO='" & request("txt_TRANSFERNO") & "'"
	end if
	q = q
  
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <form action="PRODUCTSTORAGETRANSFER.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">
		  <tr>
		<td align="left">Transfer Date</td>
		<td align="left"><input type="text" maxlength="10" name="TRANSFERDATE" size="10" value="<%=TRANSFERDATE%>">
		<a href="javascript:show_calendar('eForm.TRANSFERDATE');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
		</td></tr>
		<tr>
		<td align="left">Transfer No</td><td align="left"><input type="text" size="15" maxlength="15" name="TRANSFERNO" value="<%=TRANSFERNO%>"><font color="#FF0000">*</font></td>
		</tr>
		<TR>
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">Storage Code</font></td>
            <td align="left"> <select name="PSTORAGE_STORAGECODE">
                <%cmdCallComboSetup "PRODUCTSTORAGE",request("PSTORAGE_STORAGECODE"),"",""%>
              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">Storage Code</font></td>
				<td align="left"> <select name="PSTORAGE_STORAGECODE">
				<%cmdCallComboSetup "PRODUCTSTORAGE","STORAGECODE","",""%>
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
		</tr>
		<TR>
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">Storage Code To Tank</font></td>
            <td align="left"> <select name="PSTORAGE_STORAGECODE_TO_TANK">
                <%cmdCallComboSetup "PRODUCTSTORAGE",request("PSTORAGE_STORAGECODE_TO_TANK"),"",""%>
              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">Storage Code To Tank</font></td>
				<td align="left"> <select name="PSTORAGE_STORAGECODE_TO_TANK">
				<%cmdCallComboSetup "PRODUCTSTORAGE","STORAGECODE","",""%>
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
		</tr>
		<tr>
		<td align="left">Quantity</td><td align="left"><input type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>"></td>
		</tr>
	    <TR><td colspan="2">
		  <input Type="Hidden" Value="<%=request("Perioddate")%>" Name="Perioddate">
		  <input Type="Hidden" Value="<%=request("txt_TRANSFERNO")%>" Name="txt_TRANSFERNO">
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
	<TD align="left">TRANSFER NO</TD>	
	<TD align="left">STORAGE CODE</TD>	
	<TD align="left">STORAGE CODE TO TANK</TD>	
	<TD align="left">QUANTITY</TD>	
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
			<a href="PRODUCTSTORAGETRANSFER.asp?TRANSFERNO=<%=rsGAD("TRANSFERNO")%>&txttable=<%=request("txttable")%>&TRANSFERDATE=<%=rsGAD("TRANSFERDATE")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&PSTORAGE_STORAGECODE_TO_TANK=<%=rsGAD("PSTORAGE_STORAGECODE_TO_TANK")%>&QUANTITY=<%=rsGAD("QUANTITY")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="PRODUCTSTORAGETRANSFER.asp?TRANSFERNO=<%=rsGAD("TRANSFERNO")%>&txttable=<%=request("txttable")%>&TRANSFERDATE=<%=rsGAD("TRANSFERDATE")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
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
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""PRODUCTSTORAGETRANSFER.asp?txtModule=" &_
			  txtModule & "&iPageNo=" &  i & "&txtTable=" & request("txtTable")& "&txtField=" & request("txtField")  & _
			  "&txtSearch=" & request("txtSearch") & _
			  """>" & i & "  "
		  Next
	   End if%>
	   </a><br>
	   <%if Request("iPageNo") = "" then%>
			<%nPage = 1%>
	   <%Else%>
			<%nPage = Request("iPageNo")%>
	   <%End if%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Page : <%=nPage%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%></td>
	</tr>
   </TABLE>
<%End if%>  
 
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
</html>
