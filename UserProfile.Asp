<%Option Explicit%>
<%' On Error GOTO CmdErrorlog %>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - User and Profiles</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
</head>
<body onLoad="document.eForm.fcCode();">
<!--#include file="top.jsp"-->
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
 <%'-------------awal-----------------------%>
 <%
 Dim FCCODE, TDATE, EMPCODE, SUPERVISION_1, SUPERVISION_2, SUPERVISION_3, SUPERVISION_4, ATTENDANCE,cmdView
 Dim JOBCODE, LOCATIONTYPECODE, LOCATIONCODE, MANDAYS, UNITS, OUTPUT, OTHRS, RATE, REFERENCENO, REMARKS, FCBA
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" Then
	RATE=request("RATE")
	REFERENCENO=request("REFERENCENO")
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Delete" then
       q="Delete from Attendance_GAD where  to_char(tdate,'"& "mm/dd/yyyy" & "')='" & fdDate & _
	     "' and empcode='" & empcode & "' And LOCATIONTYPECODE='" & LOCATIONTYPECODE & "' and " & _
		 " LOCATIONCODE='" & LOCATIONCODE & "' and JOBCODE='" & JOBCODE & "'"
	else
       q="Insert into Attendance_GAD values ('" & FCCODE & "'," & " to_date('" & fddate & "','mm/dd/yyyy') " & ",'"  &  EMPCODE & _
	                               "','" &  SUPERVISION_1 & "','" & SUPERVISION_2 & "','" & _
	                                 SUPERVISION_3  & "','"  &  SUPERVISION_4 & "','" &  ATTENDANCE & "','" & JOBCODE & "','" & _
	                                 LOCATIONTYPECODE  & "','"  &  LOCATIONCODE & "','" &  MANDAYS & "','" & UNITS & "','" & _
	                                 OUTPUT  & "','"  &  OTHRS & "','" &  RATE & "','" & REFERENCENO & "','" & _
	                                 REMARKS  & "','"  &  session("BA") & "')"
	end if								 
	'Response.Write q
	'connStr.execute(q)
    cmdView	= "True"
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,fddate,txtJenis
fddate   =Request("fddate")
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
fcCode=request("fcCode")
txtJenis = request("txtJenis")

Dim rsSetup 
set rsSetup = server.CreateObject("Adodb.recordset")
q="select * from " & txtTable
response.Write "sql : " & q
Set rsSetup = Server.CreateObject("ADODB.Recordset")
rsSetup.CursorType = 3
rsSetup.CursorLocation = 3
rsSetup.Open q, connstr%>
 <%'-------------isi Form%>
 <form name="frmEntry" method="post" action="Attendance_GAD.asp">
	<table class="data">
        <%if not rsSetup.eof Then
		    nCountofField=rsSetup.fields.count-1
		  End if%>
		  <tr>
		  <%for i = 0 to  nCountofField%>
		      <td><%=rsSetup(i).name%></td>
		  <%next%>
		  </tr>
		  <tr>
		  <%for i = 0 to  nCountofField%>
		      <td><input type="text" value="<%=request(rsSetup(i).name)%>" name="<%=rsSetup(i).name%>"</td>
		  <%next%>
		  </tr>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintData.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fdDate%>&fcCode=<%=fcCode%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%'if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   q="select * from " & txtTable
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr%>
   <table class="data">
   <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-1
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="UserProfile.Asp?fcCode=<%=fcCode%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete">
			<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0"></a></td>
		    <%for i = 0 to nCountofField%>
			    <td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%></td>
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
