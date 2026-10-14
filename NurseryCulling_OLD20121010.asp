<%'Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Nursery Culling</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
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


Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA
DIM v_NURSERYCODE,v_CULLINGDATE
DIM v_QUANTITY
Dim cmdView
v_NURSERYCODE =request("v_NURSERYCODE")
v_CULLINGDATE =request("v_CULLINGDATE")
v_QUANTITY=request("v_QUANTITY")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
       q="Delete from NurseryCulling " & _
	     " where  NURSERYCODE='" & v_NURSERYCODE & _
		 "' and CULLINGDATE= to_date('" & v_CULLINGDATE & "','mm/dd/yyyy') " & _
		  " and fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	elseif Request("txtJenis")="New" then
	
	  q="Insert into NurseryCulling (NURSERYCODE, CULLINGDATE, QUANTITY, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" & v_NURSERYCODE & "'," & " to_date('" & v_CULLINGDATE & "','mm/dd/yyyy') " & "," & v_QUANTITY & _
		",'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
	    connStr.execute(q)
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update NurseryCulling set QUANTITY =" & v_QUANTITY& _
         " where  NURSERYCODE='" & v_NURSERYCODE & _
		 "' and CULLINGDATE=" & " to_date('" & v_CULLINGDATE & "','mm/dd/yyyy') " & _
		  " and fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		connStr.execute("commit")
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Request("PeriodDate")<>"" then
   PeriodDate   =Request("PeriodDate")
else
   PeriodDate   =textDate(date)
end if

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
 
 <form name="frmEntry" method="post" action="NurseryCulling.asp">
	<table class="data" width="100%">
		<tr><td align="left">Nursery Code</td>
		<td align="left"><input type="text" maxlength="10" name="v_NURSERYCODE" size="10" value="<%=v_NURSERYCODE%>"></td></tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintNurseryCulling.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&v_NURSERYCODE=<%=v_NURSERYCODE%>&v_CULLINGDATE=<%=v_CULLINGDATE%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" and len(trim(v_NURSERYCODE)) > 0 Then
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q="select * from  NurseryCulling where Nurserycode='" & v_NurseryCode & "' and fcBA='" & session("fcBA") & "'"
'response.Write q
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="NurseryCulling.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
	  <%IF Request("txtJenis")="Edit" then%>
		<tr><td align="left">Date Culling</td><td align="left">
		<input type="text" size="15" name="v_CULLINGDATE" value="<%=v_CULLINGDATE%>" readonly="True" style="background:#999999"></td></tr>
      <%else%>
		<tr><td align="left">Date Culling</td><td align="left"><input type="text" size="15" name="v_CULLINGDATE" value="<%=v_CULLINGDATE%>">
		<a href="javascript:show_calendar('eForm.v_CULLINGDATE');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a></td></tr>
	  <%End if%>
		<tr><td align="left">Quantity</td><td align="left"><input type="text" size="15" name="v_QUANTITY" value="<%=v_QUANTITY%>"  maxlength="32">  (+) Adjust Negative ; (-) Adjust Positive</td></tr>
		<tr>
		<td align="left">
		  <input type="hidden" name="v_NURSERYCODE" value="<%=v_NURSERYCODE%>">
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

   <table class="data">
	<BR>
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">NURSERY CODE</TD>	
	<TD align="left">CULLING DATE</TD>	
	<TD align="left">QUANTITY</TD>	
	</tr>
   
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">

			<a href="NurseryCulling.Asp?v_NURSERYCODE=<%=rsVH("NURSERYCODE")%>&v_QUANTITY=<%=rsVH("QUANTITY")%>&v_CULLINGDATE=<%=rsVH("CULLINGDATE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("v_NURSERYCODE")%>" border="0"></a></td><td>
			<a href="NurseryCulling.Asp?v_NURSERYCODE=<%=request("v_NURSERYCODE")%>&v_CULLINGDATE=<%=rsVH("CULLINGDATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("v_NURSERYCODE")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
		    <%for i = 0 to nCountofField-4%>
			    <td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </TABLE>
   </form>
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
