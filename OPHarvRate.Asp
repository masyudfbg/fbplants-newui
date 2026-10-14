<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - OP Harv Rate</title>
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
Dim CROPCODE, ABW_CENSUS, TERRAIN_TYPE, ABW_MIN, ABW_MAX, BASIS_QTY, BASIS_RATE, EXTRA_BASIS_RATE
Dim cmdView
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 CROPCODE=request("CROPCODE")
 ABW_CENSUS=request("ABW_CENSUS")
 TERRAIN_TYPE=request("TERRAIN_TYPE")
 ABW_MIN=request("ABW_MIN")
 ABW_MAX=request("ABW_MAX")
 BASIS_QTY=request("BASIS_QTY")
 BASIS_RATE=request("BASIS_RATE")
 EXTRA_BASIS_RATE=request("EXTRA_BASIS_RATE")

	if Request("txtJenis")= "Delete" then

       q="Delete from OPHarvRate where  CROPCODE='" & CROPCODE & "'" & _
	   " And ABW_CENSUS='" & ABW_CENSUS & "' And TERRAIN_TYPE='" & TERRAIN_TYPE & "'" & _
	   " And ABW_MIN='" & ABW_MIN & "' And ABW_MAX='" & ABW_MAX & "'" & _
	   " And BASIS_QTY='" & BASIS_QTY & "' And BASIS_RATE='" & BASIS_RATE & "'" & _
	   " And EXTRA_BASIS_RATE='" & EXTRA_BASIS_RATE & "' And fcBA='" & Session("fcBA") & "'"
	   'RESPONSE.Write(q)
	   'response.Flush()
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
	  q="Insert into OPHarvRate (CROPCODE, ABW_CENSUS, TERRAIN_TYPE, ABW_MIN, ABW_MAX, BASIS_QTY, BASIS_RATE, EXTRA_BASIS_RATE, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" & CROPCODE & "','" & ABW_CENSUS & "','" & TERRAIN_TYPE & "','" & ABW_MIN & "','" & ABW_MAX & "','" & BASIS_QTY & "','" & BASIS_RATE & "','" & EXTRA_BASIS_RATE & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		'response.Flush()
	    connStr.execute(q)
        cmdView	= "True"
			
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update OPHarvRate set CROPCODE ='" & CROPCODE & "'," & _
		"ABW_CENSUS ='" & ABW_CENSUS & "',TERRAIN_TYPE='" & TERRAIN_TYPE & "',ABW_MIN='" & ABW_MIN & "',ABW_MAX='" & ABW_MAX & "'," & _
		"BASIS_QTY ='" & BASIS_QTY & "',BASIS_RATE ='" & BASIS_RATE & "',EXTRA_BASIS_RATE ='" & EXTRA_BASIS_RATE & "'" & _
		"where CROPCODE='" & CROPCODE & "'" & _
	   " And ABW_CENSUS='" & ABW_CENSUS & "' And TERRAIN_TYPE='" & TERRAIN_TYPE & "'" & _
	   " And ABW_MIN='" & ABW_MIN & "' And ABW_MAX='" & ABW_MAX & "'" & _
	   " And BASIS_QTY='" & BASIS_QTY & "' And BASIS_RATE='" & BASIS_RATE & "'" & _
	   " And EXTRA_BASIS_RATE='" & EXTRA_BASIS_RATE & "' And fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		'RESPONSE.Write(q)
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
<% 
'   Dim rsVH
'   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
'	q="select OPHarvRateCODE, OPHarvRateDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA from  OPHarvRate"
'response.Write q
'   Set rsVH = Server.CreateObject("ADODB.Recordset")
'   rsVH.CursorType = 3
'   rsVH.CursorLocation = 3
'   rsVH.Open q, connstr%>
   <form action="OPHarvRate.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE OPHARVESTING RATE</strong></font></div></td>
          </TR>
		
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Crop Code</td>
					<% 	Dim rsCropA
					q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' order by P.FCNAME desc"
					Set rsCropA = Server.CreateObject("ADODB.Recordset")
					rsCropA.CursorType = 3
					rsCropA.CursorLocation = 3
					rsCropA.Open q, connstr
					if not rsCropA.eof then%>
					<td align="left"> <select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
					<option value=""></option> 
					<%
					Do While Not rsCropA.eoF%>
					<%IF Request("CROPCODE")=rsCropA("FCCODE") Then%>
						<option value="<%=rsCropA("FCCODE")%>" selected><%=rsCropA("FCNAME")%></option> 
					<%Else%>
						<option value="<%=rsCropA("FCCODE")%>"><%=rsCropA("FCNAME")%></option> 
					<%End if%>
					<%rsCropA.movenext
					loop
					End if
					rsCropA.Close
					set rsCropA = Nothing
					q=""
					%>
					</select> <font color="#FF0000">*</font></td>			
		<%else%>
			<td align="left">Crop Code</td>
					<% 	Dim rsCrop
					q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' order by P.FCNAME desc"
					Set rsCrop = Server.CreateObject("ADODB.Recordset")
					rsCrop.CursorType = 3
					rsCrop.CursorLocation = 3
					rsCrop.Open q, connstr
					if not rsCrop.eof then%>
					<td align="left"> <select name="CROPCODE" title="CROPCODE Character a s/d z or 1 s/d 9">         
					<option value=""></option> 
					<%
					Do While Not rsCrop.eof
					%>
					<option value="<%=rsCrop("FCCODE")%>"><%=rsCrop("FCNAME")%></option> 
					<%rsCrop.movenext
					loop
					End if
					rsCrop.Close
					set rsCrop = Nothing
					q=""
					%>
					</select> <font color="#FF0000">*</font></td>
			
		<%end if%>
		</tr>
		<tr>
		<td align="left">ABW Average</td><td align="left"><input type="text" size="10" name="ABW_CENSUS" value="<%=ABW_CENSUS%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Terrain Type</td><td align="left"><input type="text" size="15" maxlength="15" name="TERRAIN_TYPE" value="<%=TERRAIN_TYPE%>"><font color="#FF0000">*</font></td>
		</TR>
		<TR>
		<td align="left">ABW Minimum</td><td align="left"><input type="text" size="10" name="ABW_MIN" value="<%=ABW_MIN%>"></td>
		</tr>
		<tr>
		<td align="left">ABW Maximum</td><td align="left"><input type="text" size="10" name="ABW_MAX" value="<%=ABW_MAX%>"></td>
		</tr>
		<tr>
		<td align="left">Basis Quantity</td><td align="left"><input type="text" size="10" name="BASIS_QTY" value="<%=BASIS_QTY%>"></td>
		</tr>
		<tr>
		<td align="left">Basis Rate</td><td align="left"><input type="text" size="10" name="BASIS_RATE" value="<%=BASIS_RATE%>"></td>
		</tr>
		<tr>
		<td align="left">Extra Basis Rate</td><td align="left"><input type="text" size="10" name="EXTRA_BASIS_RATE" value="<%=EXTRA_BASIS_RATE%>"></td>
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
	
   q="select * from " & txtTable & ""   
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
   <table class="data" >
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CROP CODE</TD>	
	<TD align="left">ABW AVERAGE</TD>	
	<TD align="left">TERRAIN TYPE</TD>	
	<TD align="left">ABW MIN</TD>	
	<TD align="left">ABW MAX</TD>	
	<TD align="left">BASIS QTY</TD>	
	<TD align="left">BASIS RATE</TD>	
	<TD align="left">EXTRA BASIS RATE</TD>	
	</TR>
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
		  <td width="65" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="22">
				<a href="OPHarvRate.asp?CROPCODE=<%=rsGAD("CROPCODE")%>&txttable=<%=request("txttable")%>&ABW_CENSUS=<%=rsGAD("ABW_CENSUS")%>&TERRAIN_TYPE=<%=rsGAD("TERRAIN_TYPE")%>&ABW_MIN=<%=rsGAD("ABW_MIN")%>&ABW_MAX=<%=rsGAD("ABW_MAX")%>&BASIS_QTY=<%=rsGAD("BASIS_QTY")%>&BASIS_RATE=<%=rsGAD("BASIS_RATE")%>&EXTRA_BASIS_RATE=<%=rsGAD("EXTRA_BASIS_RATE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="26">
				<a href="OPHarvRate.asp?CROPCODE=<%=rsGAD("CROPCODE")%>&txttable=<%=request("txttable")%>&ABW_CENSUS=<%=rsGAD("ABW_CENSUS")%>&TERRAIN_TYPE=<%=rsGAD("TERRAIN_TYPE")%>&ABW_MIN=<%=rsGAD("ABW_MIN")%>&ABW_MAX=<%=rsGAD("ABW_MAX")%>&BASIS_QTY=<%=rsGAD("BASIS_QTY")%>&BASIS_RATE=<%=rsGAD("BASIS_RATE")%>&EXTRA_BASIS_RATE=<%=rsGAD("EXTRA_BASIS_RATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>	
		    <%for i = 0 to nCountofField%>
			    <td width="221" align="left" valign="top"><%=rsGAD(i)%></td>
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
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""OPHarvRate.asp?txtModule=" &_
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
