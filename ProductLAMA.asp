<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Product</title>
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
dim txt_PRODUCTCODE
Dim PRODUCTCODE, PRODUCTDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA
Dim cmdView
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 	txt_PRODUCTCODE=request("txt_PRODUCTCODE")
	PRODUCTCODE=request("PRODUCTCODE")
	PRODUCTDESCRIPTION=request("PRODUCTDESCRIPTION")
	CROPCODE=request("CROPCODE")
	UNITOFMEASURE=request("UNITOFMEASURE")
	FORMULA=request("FORMULA")

	if Request("txtJenis")= "Delete" then
       q="Delete from Product where PRODUCTCODE='" & PRODUCTCODE & "' And PRODUCTDESCRIPTION='" & PRODUCTDESCRIPTION & "' And CROPCODE='" & CROPCODE & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into Product (PRODUCTCODE, PRODUCTDESCRIPTION, CROPCODE, UNITOFMEASURE,FORMULA,FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" & PRODUCTCODE & "','" & PRODUCTDESCRIPTION & "','" &  CROPCODE & "','" &  FORMULA & "','" &  UNITOFMEASURE & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		response.Flush()
	    connStr.execute(q)
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update Product set UNITOFMEASURE ='" & UNITOFMEASURE & "'," & _
		"FORMULA ='" & FORMULA & "' where PRODUCTCODE='" & PRODUCTCODE & "' And PRODUCTDESCRIPTION='" & PRODUCTDESCRIPTION & "' And CROPCODE='" & CROPCODE & "' And fcBA='" & Session("fcBA") & "'"
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
 <form name="frmEntry" method="post" action="Product.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE PRODUCT</strong></font></div></td>
          </TR>
		<tr>
			<td align="left">Product Code</td><td align="left"><input type="text" size="15" name="txt_PRODUCTCODE" value="<%=REQUEST("txt_PRODUCTCODE")%>"></td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintProduct.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_PRODUCTCODE=<%=REQUEST("txt_PRODUCTCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select PRODUCTCODE, PRODUCTDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA, FCENTRY, FCEDIT, FCIP, FCBA from PRODUCT where " & _
     " fcBA='" & session("fcBA") & "'"
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("txt_PRODUCTCODE"))) > 0 then
		q = q & " and PRODUCTCODE='" & request("txt_PRODUCTCODE") & "'"
	end if
	q = q
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <form action="Product.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Product Code</td><td align="left"><input type="text" size="15" name="PRODUCTCODE" value="<%=PRODUCTCODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
		<%else%>
			<td align="left">Product Code</td><td align="left"><input type="text" size="15" name="PRODUCTCODE" value="<%=PRODUCTCODE%>"  ><font color="#FF0000">*</font></td>
		<%end if%>
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Product Description</td><td align="left"><input type="text" size="50" name="PRODUCTDESCRIPTION" value="<%=PRODUCTDESCRIPTION%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left">Product Description</td><td align="left"><input type="text" size="50" name="PRODUCTDESCRIPTION" value="<%=PRODUCTDESCRIPTION%>"><font color="#FF0000">*</font></td>
		<%End if%>
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
				<td align="left"><font face="Verdana" size="1">Crop Code</font></td>
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
		<%Else%>
				<td align="left"><font face="Verdana" size="1">Crop Code</font></td>
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
		<%End if%>
		</TR>
		<TR>
		<td align="left">Unit Of Measure</td><td align="left"><input type="text" size="15" name="UNITOFMEASURE" value="<%=UNITOFMEASURE%>"></td>
		</tr>
		<tr>
		<td align="left">Formula</td><td align="left"><input type="text" size="15" name="FORMULA" value="<%=FORMULA%>"></td>
		</tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="txt_PRODUCTCODE" value="<%=REQUEST("txt_PRODUCTCODE")%>">		
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
	<TD align="left">PRODUCT CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">CROP CODE</TD>	
	<TD align="left">UNIT OF MEASURE</TD>	
	<TD align="left">FORMULA</TD>	
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
			<a href="Product.asp?PRODUCTCODE=<%=rsGAD("PRODUCTCODE")%>&txttable=<%=request("txttable")%>&PRODUCTDESCRIPTION=<%=rsGAD("PRODUCTDESCRIPTION")%>&CROPCODE=<%=rsGAD("CROPCODE")%>&UNITOFMEASURE=<%=rsGAD("UNITOFMEASURE")%>&FORMULA=<%=rsGAD("FORMULA")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="Product.asp?PRODUCTCODE=<%=rsGAD("PRODUCTCODE")%>&txttable=<%=request("txttable")%>&PRODUCTDESCRIPTION=<%=rsGAD("PRODUCTDESCRIPTION")%>&CROPCODE=<%=rsGAD("CROPCODE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
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
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""Product.asp?txtModule=" &_
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
