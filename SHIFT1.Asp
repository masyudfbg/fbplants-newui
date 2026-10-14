<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Shift</title>
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
Dim SHIFTCODE, STARTTIME, ENDTIME, txtDESCRIPTION
Dim cmdView
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 SHIFTCODE=request("SHIFTCODE")
 STARTTIME=request("STARTTIME")
 ENDTIME=request("ENDTIME")
 txtDESCRIPTION=request("txtDESCRIPTION")

	if Request("txtJenis")= "Delete" then
       q="Delete from Shift where SHIFTCODE='" & SHIFTCODE & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into Shift (SHIFTCODE, STARTTIME, ENDTIME, DESCRIPTION,FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" & SHIFTCODE & "','" & STARTTIME & "','" & ENDTIME & "','" &  txtDESCRIPTION & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		'response.Flush()
	    connStr.execute(q)
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update Shift set SHIFTCODE ='" & SHIFTCODE & "'," & _
		"STARTTIME ='" & STARTTIME & "',ENDTIME='" & ENDTIME & "'," & _
		"DESCRIPTION ='" & txtDESCRIPTION & "'" & _
		"where ShiftCODE='" & ShiftCODE & "' And fcBA='" & Session("fcBA") & "'"
		'response.Write(q)
		'response.Flush()
		connStr.execute(q)
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
<% 
'   Dim rsVH
'   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
'	q="select ShiftCODE, ShiftDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA from  Shift"
'response.Write q
'   Set rsVH = Server.CreateObject("ADODB.Recordset")
'   rsVH.CursorType = 3
'   rsVH.CursorLocation = 3
'   rsVH.Open q, connstr%>
   <form action="Shift.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE SHIFT</strong></font></div></td>
          </TR>		
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Shift Code</td><td align="left"><input type="text" size="15" name="SHIFTCODE" value="<%=SHIFTCODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font> </td>
		<%else%>
			<td align="left">Shift Code</td><td align="left"><input type="text" size="15" name="SHIFTCODE" value="<%=SHIFTCODE%>"  ><font color="#FF0000">*</font> </td>
		<%end if%>
		</tr>
		<tr><td align="left">Time Start</td>
		<td align="left"><input type="text" maxlength="10" name="STARTTIME" size="10" value="<%=STARTTIME%>">
		</td></tr>	
		<tr><td align="left">Time End</td>
		<td align="left"><input type="text" maxlength="10" name="ENDTIME" size="10" value="<%=ENDTIME%>"> 
		</td></tr>	
		<tr>
		<td align="left">Description</td><td align="left"><input type="text" size="15" name="txtDESCRIPTION" value="<%=txtDESCRIPTION%>"></td>
		</tr>
	    <TR><td colspan="2">
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
   Dim rsGAD
	
   q="select * from " & txtTable & " where fcBA='" & Session("fcBA") & "'"   
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr
   JumRec=1
   intRecordsPerPage = 20
	
	Dim intRecordsPerPage,intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
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
	Dim cColor

   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
   %>
   <table width="737" class="data" >
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">SHIFT CODE</TD>	
	<TD align="left">START TIME</TD>	
	<TD align="left">END TIME</TD>	
	<TD align="left">DESCRIPTION</TD>	
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
			<a href="Shift.asp?SHIFTCODE=<%=rsGAD("SHIFTCODE")%>&txttable=<%=request("txttable")%>&STARTTIME=<%=rsGAD("STARTTIME")%>&ENDTIME=<%=rsGAD("ENDTIME")%>&txtDESCRIPTION=<%=rsGAD("DESCRIPTION")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="Shift.asp?SHIFTCODE=<%=rsGAD("SHIFTCODE")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
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
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""Shift.asp?txtModule=" &_
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
<%'End if%>  
 
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
