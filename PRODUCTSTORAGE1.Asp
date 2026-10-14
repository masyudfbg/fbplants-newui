<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS -  Product Storage</title>
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
Dim STORAGECODE, MAXCAPACITY, txtDESCRIPTION, PRODUCT_PRODUCTCODE, DEADSTOCK_CAPACITY,REMARKS
Dim cmdView,txt_STORAGECODE
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 	txt_STORAGECODE=request("txt_STORAGECODE")
	STORAGECODE=request("STORAGECODE")
	MAXCAPACITY=request("MAXCAPACITY")
	txtDESCRIPTION=request("txtDESCRIPTION")
	PRODUCT_PRODUCTCODE=request("PRODUCT_PRODUCTCODE")
	DEADSTOCK_CAPACITY=request("DEADSTOCK_CAPACITY")
	REMARKS=request("REMARKS")
	

	if Request("txtJenis")= "Delete" then
       q="Delete from PRODUCTSTORAGE where STORAGECODE='" & STORAGECODE & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into PRODUCTSTORAGE (STORAGECODE, MAXCAPACITY, DESCRIPTION, PRODUCT_PRODUCTCODE, DEADSTOCK_CAPACITY,REMARKS,FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" & STORAGECODE & "','" & MAXCAPACITY & "','" &  txtDESCRIPTION & "','" &  PRODUCT_PRODUCTCODE & "','" &  DEADSTOCK_CAPACITY & "','" &  REMARKS & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		response.Flush()
	    connStr.execute(q)
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update PRODUCTSTORAGE set STORAGECODE ='" & STORAGECODE & "'," & _
		"MAXCAPACITY ='" & MAXCAPACITY & "',DESCRIPTION='" & txtDESCRIPTION & "'," & _
		"PRODUCT_PRODUCTCODE ='" & PRODUCT_PRODUCTCODE & "'," & _
		"DEADSTOCK_CAPACITY ='" & DEADSTOCK_CAPACITY & "',REMARKS ='" & REMARKS & "' where STORAGECODE='" & STORAGECODE & "' And fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		connStr.execute("commit")
		'RESPONSE.Write(Q)
		'RESPONSE.Flush()
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
<form name="frmEntry" method="post" action="PRODUCTSTORAGE.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE PRODUCT STORAGE</strong></font></div></td>
          </TR>
		<tr>
			<td align="left">Storage Code</td><td align="left"><input type="text" size="15" name="txt_STORAGECODE" value="<%=REQUEST("txt_STORAGECODE")%>"  ></td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;
		<a href="PrintPRODUCTSTORAGE.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_STORAGECODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		
		</td>
		</tr>
    </table>  
 </form>
 <%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select STORAGECODE, DESCRIPTION, PRODUCT_PRODUCTCODE, MAXCAPACITY, DEADSTOCK_CAPACITY, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA from PRODUCTSTORAGE where " & _
     " fcBA='" & session("fcBA") & "'"
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("txt_STORAGECODE"))) > 0 then
		q = q & " and STORAGECODE='" & request("txt_STORAGECODE") & "'"
	end if
	q = q
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <form action="PRODUCTSTORAGE.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">

		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Storage Code</td><td align="left"><input type="text" size="15" maxlength="15" name="STORAGECODE" value="<%=STORAGECODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
		<%else%>
			<td align="left">Storage Code</td><td align="left"><input type="text" size="15" maxlength="15" name="STORAGECODE" value="<%=STORAGECODE%>"  ><font color="#FF0000">*</font></td>
		<%end if%>
		</tr>
		<tr>
		<td align="left">Description</td><td align="left"><input type="text" size="30" maxlength="50" name="txtDESCRIPTION" value="<%=txtDESCRIPTION%>"><font color="#FF0000">*</font></td>
		</TR>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Product Storage Code</td>
			<% 	Dim rsProductA
			q = "select P.PRODUCTCODE,P.PRODUCTDESCRIPTION from  PRODUCT P"
			Set rsProductA = Server.CreateObject("ADODB.Recordset")
			rsProductA.CursorType = 3
			rsProductA.CursorLocation = 3
			rsProductA.Open q, connstr
			if not rsProductA.eof then%>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE" title="PRODUCT_PRODUCTCODE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%
			Do While Not rsProductA.eoF%>
			<%IF Request("PRODUCT_PRODUCTCODE")=rsProductA("PRODUCTCODE") Then%>
				<option value="<%=rsProductA("PRODUCTCODE")%>" selected><%=rsProductA("PRODUCTDESCRIPTION")%></option> 
			<%Else%>
				<option value="<%=rsProductA("PRODUCTCODE")%>"><%=rsProductA("PRODUCTDESCRIPTION")%></option> 
			<%End if%>
			<%rsProductA.movenext
			loop
			End if
			rsProductA.Close
			set rsProductA = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>			
			<td align="left">Product Storage Code</td>
			<% 	Dim rsProductc
			q = "select P.PRODUCTCODE,P.PRODUCTDESCRIPTION from  PRODUCT P"
			Set rsProductc = Server.CreateObject("ADODB.Recordset")
			rsProductc.CursorType = 3
			rsProductc.CursorLocation = 3
			rsProductc.Open q, connstr
			if not rsProductc.eof then%>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE" title="PRODUCT_PRODUCTCODE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%
			Do While Not rsProductc.eoF%>
			<%IF Request("PRODUCT_PRODUCTCODE")=rsProductc("PRODUCTCODE") Then%>
				<option value="<%=rsProductc("PRODUCTCODE")%>" selected><%=rsProductc("PRODUCTDESCRIPTION")%></option> 
			<%Else%>
				<option value="<%=rsProductc("PRODUCTCODE")%>"><%=rsProductc("PRODUCTDESCRIPTION")%></option> 
			<%End if%>
			<%rsProductc.movenext
			loop
			End if
			rsProductc.Close
			set rsProductc = Nothing
			q=""
			%>
			</select> <font color="#FF0000">*</font></td>
		<%End if%>
		</tr>
		<tr>
		<td align="left">Maximum Capacity (Kg)</td><td align="left"><input type="text" size="15" name="MAXCAPACITY" value="<%=MAXCAPACITY%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Deadstock Capacity (Kg)</td><td align="left"><input type="text" size="15" name="DEADSTOCK_CAPACITY" value="<%=DEADSTOCK_CAPACITY%>"></td>
		</tr>
		<tr>
		<td align="left">Remarks</td><td align="left"><input type="text" size="15" maxlength="100" name="REMARKS" value="<%=REMARKS%>"></td>
		</tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="txt_STORAGECODE" value="<%=REQUEST("txt_STORAGECODE")%>">		
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
	<TD align="left">STORAGE CODE</TD>
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">PRODUCT CODE</TD>		
	<TD align="left">MAX CAPACITY</TD>		
	<TD align="left">DEADSTOCK CAPACITY</TD>	
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

			<a href="PRODUCTSTORAGE.asp?STORAGECODE=<%=rsGAD("STORAGECODE")%>&txttable=<%=request("txttable")%>&MAXCAPACITY=<%=rsGAD("MAXCAPACITY")%>&txtDESCRIPTION=<%=rsGAD("DESCRIPTION")%>&PRODUCT_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&DEADSTOCK_CAPACITY=<%=rsGAD("DEADSTOCK_CAPACITY")%>&REMARKS=<%=rsGAD("REMARKS")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="PRODUCTSTORAGE.asp?STORAGECODE=<%=rsGAD("STORAGECODE")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
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
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""PRODUCTSTORAGE.asp?txtModule=" &_
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
