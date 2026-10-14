<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Weighbridge Ticket</title>
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
dim txt_PRODUCTCODE,Perioddate

Dim TRANSACTIONTYPE, CHITNO, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE, FIELD_ESTATE_ESTATECODE
Dim FIELD_FIELDCODE, DATEIN, WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT, WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT
Dim cmdView
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 txt_PRODUCTCODE=request("txt_PRODUCTCODE")
 Perioddate=request("Perioddate")
TRANSACTIONTYPE=request("TRANSACTIONTYPE")
CHITNO=request("CHITNO")
PRODUCT_PRODUCTCODE=request("PRODUCT_PRODUCTCODE")
VEHICLENO=request("VEHICLENO")
CUST_CUSTOMERCODE=request("CUST_CUSTOMERCODE")
FIELD_ESTATE_ESTATECODE=request("FIELD_ESTATE_ESTATECODE")
FIELD_FIELDCODE=request("FIELD_FIELDCODE")
DATEIN=request("DATEIN")
WBRIDGE_WEIGHBRIDGECODEIN=request("WBRIDGE_WEIGHBRIDGECODEIN")
WEIGHTIN=request("WEIGHTIN")
DATEOUT=request("DATEOUT")
WBRIDGE_WEIGHBRIDGECODEOUT=request("WBRIDGE_WEIGHBRIDGECODEOUT")
WEIGHTOUT=request("WEIGHTOUT")

	if Request("txtJenis")= "Delete" then
       q="Delete from WEIGHBRIDGETICKET where PRODUCT_PRODUCTCODE='" & PRODUCT_PRODUCTCODE & "' And CHITNO='" & CHITNO & "' And CUST_CUSTOMERCODE='" & CUST_CUSTOMERCODE & "' And DATEIN=to_Date('" & DATEIN & "','MM/DD/YYYY') And DATEOUT=to_Date('" & DATEOUT & "','MM/DD/YYYY') And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
	   'response.Write(q)
	   'response.Flush()
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
	  q="Insert into WEIGHBRIDGETICKET (TRANSACTIONTYPE, CHITNO, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE, FIELD_ESTATE_ESTATECODE, FIELD_FIELDCODE, DATEIN, WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT, WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" &  TRANSACTIONTYPE & "','" &  CHITNO & "','" &  PRODUCT_PRODUCTCODE & "','" &  VEHICLENO & "','" &  CUST_CUSTOMERCODE & "','" &  FIELD_ESTATE_ESTATECODE & "'," & _
		"'" &  FIELD_FIELDCODE & "',to_Date('" & DATEIN & "','MM/DD/YYYY'),'" &  WBRIDGE_WEIGHBRIDGECODEIN & "'," &  WEIGHTIN & ",to_Date('" & DATEOUT & "','MM/DD/YYYY')," & _
		"'" &  WBRIDGE_WEIGHBRIDGECODEOUT & "'," &  WEIGHTOUT & "," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		'response.Flush()
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update WEIGHBRIDGETICKET set TRANSACTIONTYPE ='" & TRANSACTIONTYPE & "',VEHICLENO ='" & VEHICLENO & "',FIELD_ESTATE_ESTATECODE ='" & FIELD_ESTATE_ESTATECODE & "'," & _
	   	"FIELD_FIELDCODE ='" & FIELD_FIELDCODE & "',WBRIDGE_WEIGHBRIDGECODEIN ='" & WBRIDGE_WEIGHBRIDGECODEIN & "',WEIGHTIN =" & WEIGHTIN & "," & _
		"WBRIDGE_WEIGHBRIDGECODEOUT ='" & WBRIDGE_WEIGHBRIDGECODEOUT & "',WEIGHTOUT=" & WEIGHTOUT & "" & _
		" where PRODUCT_PRODUCTCODE='" & PRODUCT_PRODUCTCODE & "' And CHITNO='" & CHITNO & "' And CUST_CUSTOMERCODE='" & CUST_CUSTOMERCODE & "' And DATEIN=to_Date('" & DATEIN & "','MM/DD/YYYY') And DATEOUT=to_Date('" & DATEOUT & "','MM/DD/YYYY') And fcBA='" & Session("fcBA") & "'"
		'response.Write(q)
		'response.Flush()
		connStr.execute(q)
		connStr.execute("commit")
		'RESPONSE.Write "SQLnsddsdya : " & Q
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & 
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
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if

 '-------------ISI Payment Form%>

 <form name="frmEntry" method="post" action="WEIGHBRIDGETICKET.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<TR> 
		<td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE WEIGHBRIDGETICKET</strong></font></div></td>
		</TR>

		<tr><td align="left">Date</td>
		<td align="left"><input type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		<a href="javascript:show_calendar('frmEntry.Perioddate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<TR>
			<td align="left">Product Code</td>
			<td align="left"> <select name="txt_PRODUCTCODE">
			<%cmdCallComboSetup "PRODUCT",REQUEST("txt_PRODUCTCODE"),"",""%>
			</select></td>
		</TR>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintWEIGHBRIDGETICKET.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_PRODUCTCODE=<%=REQUEST("txt_PRODUCTCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>

<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select TRANSACTIONTYPE, CHITNO, PRODUCT_PRODUCTCODE, VEHICLENO, CUST_CUSTOMERCODE, FIELD_ESTATE_ESTATECODE, FIELD_FIELDCODE, DATEIN, " & _ 
     "WBRIDGE_WEIGHBRIDGECODEIN, WEIGHTIN, DATEOUT, WBRIDGE_WEIGHBRIDGECODEOUT, WEIGHTOUT, FCENTRY, FCEDIT, FCIP, FCBA from WEIGHBRIDGETICKET where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(request("Perioddate"))) > 0 then
		q = q & " and to_char(DATEIN,'mm/yyyy')=to_char(to_date('" & request("Perioddate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_PRODUCTCODE"))) > 0 then
		q = q & " and PRODUCT_PRODUCTCODE='" & request("txt_PRODUCTCODE") & "'"
	end if
	q = q 
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>

   <form action="WEIGHBRIDGETICKET.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		<tr>
		<%IF txtJenis="Edit" Then%>
		<td align="left">Transaction Type</td>
			<% 	Dim rsType
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'W0001' and P.FCBA = '" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsType = Server.CreateObject("ADODB.Recordset")
			rsType.CursorType = 3
			rsType.CursorLocation = 3
			rsType.Open q, connstr
			if not rsType.eof then%>
     	    <td align="left"> <select name="TRANSACTIONTYPE" title="TRANSACTIONTYPE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%			Do While Not rsType.eof%>
			<%if rsType("FCCODE")=request("TRANSACTIONTYPE") Then%>			
				<option value="<%=rsType("FCCODE")%>" selected><%=rsType("FCNAME")%></option> 
			<%Else%>
				<option value="<%=rsType("FCCODE")%>"><%=rsType("FCNAME")%></option> 
			<%End if%>
			<%rsType.movenext
			loop
			End if
			rsType.Close
			set rsType = Nothing
			q=""
			%>
			</select></td>
		<%Else%>
		<td align="left">Transaction Type</td>
			<% 	Dim rsTypeA
			q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'W0001' and P.FCBA = '" & Session("fcBA") & "' order by P.FCNAME desc"
			Set rsTypeA = Server.CreateObject("ADODB.Recordset")
			rsTypeA.CursorType = 3
			rsTypeA.CursorLocation = 3
			rsTypeA.Open q, connstr
			if not rsTypeA.eof then%>
     	    <td align="left"> <select name="TRANSACTIONTYPE" title="TRANSACTIONTYPE Character a s/d z or 1 s/d 9">         
			<option value=""></option> 
			<%	Do While Not rsTypeA.eof%>
				<option value="<%=rsTypeA("FCCODE")%>"><%=rsTypeA("FCNAME")%></option> 
			<%rsTypeA.movenext
			loop
			End if
			rsTypeA.Close
			set rsTypeA = Nothing
			q=""
			%>
			</select></td>
		<%End if%>
		</tr>
		<tr>
		<td align="left">Chit No</td><td align="left"><input type="text" size="15" maxlength="15" name="CHITNO" value="<%=CHITNO%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
			<td align="left">Product Code</td>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE">
			<%cmdCallComboSetup "PRODUCT",request("PRODUCT_PRODUCTCODE"),"",""%>
			</select> <font color="#FF0000">*</font> </td>
		<%Else%>
			<td align="left">Product Code</td>
			<td align="left"> <select name="PRODUCT_PRODUCTCODE">
			<%cmdCallComboSetup "PRODUCT","ProductCODE","",""%>
			</select> <font color="#FF0000">*</font> </td>
		<%End if%>		
		</tr>
			<tr>
            <td align="left"><font face="Verdana" size="1">Vehicle</font></td><td align="left"><input type="text" size="15" maxlength="50" name="VEHICLENO" value="<%=VEHICLENO%>"><font color="#FF0000">*</font></td>
			</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
			<td align="left">Customer Code</td>
			<td align="left"> <select name="CUST_CUSTOMERCODE">
			<%cmdCallComboSetup "CUSTOMER",request("CUST_CUSTOMERCODE"),"",""%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left">Customer Code</td>
			<td align="left"> <select name="CUST_CUSTOMERCODE">
			<%cmdCallComboSetup "CUSTOMER","CUSTOMERCODE","",""%>
			</select> <font color="#FF0000">*</font></td>
		<%End if%>
		<tr>
		<tr>
		<td align="left">Field Estate Code</td><td align="left"><input type="text" size="15" name="FIELD_ESTATE_ESTATECODE" value="<%=FIELD_ESTATE_ESTATECODE%>"></td>
		</tr>
		<tr>
		<td align="left"><font face="Verdana" size="1">Field Code</font></td><td align="left"><input type="text" size="15" maxlength="30" name="FIELD_FIELDCODE" value="<%=FIELD_FIELDCODE%>"></td>
		</tr>
		<tr>
		
		<td align="left">Date In</td>
			<%IF txtJenis="Edit" Then%>
			<td align="left"><input type="text" maxlength="10" name="DATEIN" size="10" value="<%=DATEIN%>">
			<%Else%>
			<td align="left"><input type="text" maxlength="10" name="DATEIN" size="10" value="<%=DATE()%>">
			<%End if%>
			<a href="javascript:show_calendar('eForm.DATEIN');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
		<td align="left">Weighbridge Code In</td>
		<%IF txtJenis="Edit" Then%>
			<td align="left"> <select name="WBRIDGE_WEIGHBRIDGECODEIN">
			<%cmdCallComboSetup "WEIGHBRIDGE",REQUEST("WBRIDGE_WEIGHBRIDGECODEIN"),"",""%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left"> <select name="WBRIDGE_WEIGHBRIDGECODEIN">
			<%cmdCallComboSetup "WEIGHBRIDGE","WBRIDGE_WEIGHBRIDGECODEIN","",""%>
			</select> <font color="#FF0000">*</font></td>		
		<%End if%>
		</tr>
		<tr>
		<td align="left">Weight In (Kg)</td><td align="left"><input type="text" size="15" maxlength="15" name="WEIGHTIN" value="<%=WEIGHTIN%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Date Out</td>
			<%IF txtJenis="Edit" Then%>
				<td align="left"><input type="text" maxlength="10" name="DATEOUT" size="10" value="<%=DATEOUT%>">
			<%Else%>
				<td align="left"><input type="text" maxlength="10" name="DATEOUT" size="10" value="<%=DATE()%>">
			<%End IF%>
			<a href="javascript:show_calendar('eForm.DATEOUT');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
		<td align="left">Weighbridge Code Out</td>
		<%IF txtJenis="Edit" Then%>
			<td align="left"> <select name="WBRIDGE_WEIGHBRIDGECODEOUT">
			<%cmdCallComboSetup "WEIGHBRIDGE",REQUEST("WBRIDGE_WEIGHBRIDGECODEOUT"),"",""%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left"> <select name="WBRIDGE_WEIGHBRIDGECODEOUT">
			<%cmdCallComboSetup "WEIGHBRIDGE","WBRIDGE_WEIGHBRIDGECODEOUT","",""%>
			</select> <font color="#FF0000">*</font></td>		
		<%End if%>

		</tr>
		<tr>
		<td align="left">Weight Out (Kg)</td><td align="left"><input type="text" size="15" maxlength="15" name="WEIGHTOUT" value="<%=WEIGHTOUT%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_PRODUCTCODE" value="<%=txt_PRODUCTCODE%>">
	
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
  
	JumRec=1
	intRecordsPerPage = 20
	Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
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
          <TD align="left">TRANS TYPE</TD>
          <TD align="left">CHITNO</TD>
          <TD align="left">PRODUCT CODE</TD>
          <TD align="left">VEHICLE</TD>
          <TD align="left">CUSTOMER</TD>
          <TD align="left">FIELD ESTATE</TD>
          <TD align="left">FIELD CODE</TD>
          <TD align="left">DATE IN</TD>
          <TD align="left">WEIGHBRIDGE CODE IN</TD>
          <TD align="left">WEIGHT IN</TD>
          <TD align="left">DATE OUT</TD>
          <TD align="left">WEIGHBRIDGE CODE OUT</TD>
          <TD align="left">WEIGHT OUT</TD>
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
          <td valign="middle" width="21"> <a href="WEIGHBRIDGETICKET.asp?PRODUCT_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&txttable=<%=request("txttable")%>&CHITNO=<%=rsGAD("CHITNO")%>&VEHICLENO=<%=rsGAD("VEHICLENO")%>&WEIGHTIN=<%=rsGAD("WEIGHTIN")%>&WEIGHTOUT=<%=rsGAD("WEIGHTOUT")%>&CUST_CUSTOMERCODE=<%=rsGAD("CUST_CUSTOMERCODE")%>&DATEIN=<%=rsGAD("DATEIN")%>&DATEOUT=<%=rsGAD("DATEOUT")%>&WBRIDGE_WEIGHBRIDGECODEIN=<%=rsGAD("WBRIDGE_WEIGHBRIDGECODEIN")%>&TRANSACTIONTYPE=<%=rsGAD("TRANSACTIONTYPE")%>&FIELD_ESTATE_ESTATECODE=<%=rsGAD("FIELD_ESTATE_ESTATECODE")%>&FIELD_FIELDCODE=<%=rsGAD("FIELD_FIELDCODE")%>&WBRIDGE_WEIGHBRIDGECODEOUT=<%=rsGAD("WBRIDGE_WEIGHBRIDGECODEOUT")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a> 
          </td>
          <td valign="middle" width="20"> <a href="WEIGHBRIDGETICKET.asp?PRODUCT_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&txttable=<%=request("txttable")%>&CHITNO=<%=rsGAD("CHITNO")%>&CUST_CUSTOMERCODE=<%=rsGAD("CUST_CUSTOMERCODE")%>&DATEIN=<%=rsGAD("DATEIN")%>&DATEOUT=<%=rsGAD("DATEOUT")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a> 
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
          <td colspan="10"> <%if intNumberOfPages<>1 then
		  For i = 1 to intNumberOfPages
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""WEIGHBRIDGETICKET.asp?txtModule=" &_
			  txtModule & "&iPageNo=" &  i & "&txtTable=" & request("txtTable")& "&txtField=" & request("txtField")  & _
			  "&txtSearch=" & request("txtSearch") & _
			  """>" & i & "  "
		  Next
	   End if%> </a><br> <%if Request("iPageNo") = "" then%> <%nPage = 1%> <%Else%> <%nPage = Request("iPageNo")%> <%End if%> <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Page 
            : <%=nPage%> <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total 
            Records : <%=intRecords%></td>
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
