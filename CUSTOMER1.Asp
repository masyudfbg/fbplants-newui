<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Customer</title>
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
Dim CUSTOMERCODE, txtDESCRIPTION, SUPPLIER, ICNO, BANK, BANKACCOUNTNO, ADDRESS, COMPANYNO
Dim cmdView,txt_CUSTOMERCODE
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 txt_CUSTOMERCODE=request("txt_CUSTOMERCODE")
CUSTOMERCODE=request("CUSTOMERCODE")
txtDESCRIPTION=request("txtDESCRIPTION")
SUPPLIER=request("SUPPLIER")
ICNO=request("ICNO")
BANK=request("BANK")
BANKACCOUNTNO=request("BANKACCOUNTNO")
ADDRESS=request("ADDRESS")
COMPANYNO=request("COMPANYNO")


	if Request("txtJenis")= "Delete" then
       q="Delete from CUSTOMER where  CUSTOMERCODE='" & CUSTOMERCODE & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into CUSTOMER (CUSTOMERCODE, DESCRIPTION, SUPPLIER, ICNO, BANK, BANKACCOUNTNO, ADDRESS, COMPANYNO, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" & CUSTOMERCODE & "','" & txtDESCRIPTION & "','" &  SUPPLIER & "','" &  ICNO & "','" &  BANK & "'," & _
		"'" &  BANKACCOUNTNO & "','" &  ADDRESS & "','" &  COMPANYNO & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		response.Flush()
	    connStr.execute(q)
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update CUSTOMER set DESCRIPTION ='" & txtDESCRIPTION & "',ICNO ='" & ICNO & "',BANK ='" & BANK & "'," & _
		"BANKACCOUNTNO ='" & BANKACCOUNTNO & "',ADDRESS ='" & ADDRESS & "',COMPANYNO ='" & COMPANYNO & "'," & _
		"SUPPLIER ='" & SUPPLIER & "'" & _
		" where  CUSTOMERCODE='" & CUSTOMERCODE & "' And fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		connStr.execute("commit")
		'response.Write(q)
		'response.Flush()
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
<form name="frmEntry" method="post" action="CUSTOMER.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE CUSTOMER INFORMATION</strong></font></div></td>
          </TR>
		<tr>
			<td align="left">Customer Code</td><td align="left"><input type="text" size="15" name="txt_CUSTOMERCODE" value="<%=request("txt_CUSTOMERCODE")%>"  ></td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;
		<a href="PrintCustomer.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		
		</td>
		</tr>
    </table>  
 </form>
 <%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select CUSTOMERCODE, DESCRIPTION, SUPPLIER, ICNO, BANK, BANKACCOUNTNO, ADDRESS, COMPANYNO, FCENTRY, FCEDIT, FCIP, FCBA from CUSTOMER where " & _
     " fcBA='" & session("fcBA") & "'"
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("txt_CUSTOMERCODE"))) > 0 then
		q = q & " and CUSTOMERCODE='" & request("txt_CUSTOMERCODE") & "'"
	end if
	q = q
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <form action="CUSTOMER.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="94%">
		
		<%IF Request("txtJenis")="Edit" then%>
			<td width="17%" align="left">Customer Code</td><td width="64%" align="left"><input type="text" size="15" maxlength="15" name="CUSTOMERCODE" value="<%=CUSTOMERCODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
		<%else%>
			<td width="8%" align="left">Customer Code</td><td width="11%" align="left"><input type="text" size="15" maxlength="15" name="CUSTOMERCODE" value="<%=CUSTOMERCODE%>"  ><font color="#FF0000">*</font></td>
		<%end if%>
		</tr>

		<tr>
		<td align="left">Description</td><td align="left"><input type="text" size="70" maxlength="255" name="txtDESCRIPTION" value="<%=txtDESCRIPTION%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Supplier</td>
			<td align="left"><select name="SUPPLIER">
                <%IF request("SUPPLIER")="Y" THEN %>
                <Option value="Y" selected>Yes</Option>
                <Option value="N">No</Option>
                <%ELSE%>
                <Option value="Y">Yes</Option>
                <Option value="N" selected>No</Option>
              </select> </td>
				<%End If%>
		<%Else%>
			<td align="left">Supplier</td>
			<td align="left"><select name="SUPPLIER">
			<Option value="Y">Yes</Option>
			<Option value="N">No</Option>
			</select></td>
		<%End If%>
		</tr>
		<tr>
		<td align="left">ICNO</td><td align="left"><input type="text" size="70"  maxlength="240" name="ICNO" value="<%=ICNO%>"></td>
		</TR>
		<TR>
		<td align="left">Bank</td><td align="left"><input type="text" size="15" maxlength="15" name="BANK" value="<%=BANK%>"></td>
		</tr>
		<tr>
		<td align="left">Bank Account No</td><td align="left"><input type="text" size="40"  maxlength="40" name="BANKACCOUNTNO" value="<%=BANKACCOUNTNO%>"></td>
		</tr>
		<tr>
		<td align="left">Address</td><td align="left"><input type="text" size="60" maxlength="100" name="ADDRESS" value="<%=ADDRESS%>"></td>
		</tr>
		<tr>
		<td align="left">Company No</td><td align="left"><input type="text" size="30" maxlength="30" name="COMPANYNO" value="<%=COMPANYNO%>"></td>
		</tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="txt_CUSTOMERCODE" value="<%=REQUEST("txt_CUSTOMERCODE")%>">		
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
   <table class="data" >
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CUSTOMER CODE</TD>	
	<TD align="left">DESCRIPTION</TD>	
	<TD align="left">SUPPLIER</TD>	
	<TD align="left">ICNO</TD>	
	<TD align="left">BANK</TD>	
	<TD align="left">BANK ACCOUNT NO</TD>	
	<TD align="left">ADDRESS</TD>	
	<TD align="left">COMPANY NO</TD>	
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
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="CUSTOMER.asp?CUSTOMERCODE=<%=rsGAD("CUSTOMERCODE")%>&txttable=<%=request("txttable")%>&txtDESCRIPTION=<%=rsGAD("DESCRIPTION")%>&ICNO=<%=rsGAD("ICNO")%>&BANK=<%=rsGAD("BANK")%>&BANKACCOUNTNO=<%=rsGAD("BANKACCOUNTNO")%>&ADDRESS=<%=rsGAD("ADDRESS")%>&COMPANYNO=<%=rsGAD("COMPANYNO")%>&SUPPLIER=<%=rsGAD("SUPPLIER")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>		
			</td>
			<td valign="middle" width="30">
			<a href="CUSTOMER.asp?CUSTOMERCODE=<%=rsGAD("CUSTOMERCODE")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
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
	<tr>
	<td colspan="10">
	   <%if intNumberOfPages<>1 then
		  For i = 1 to intNumberOfPages
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""CUSTOMER.asp?txtModule=" &_
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
