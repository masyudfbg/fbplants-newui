<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Delivery Contract</title>
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
Dim CONTRACTCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, TOTALWEIGHT, txtSTATUS, REFERENCE, CUST_CUSTOMERCODE, PRODUCT_PRODUCTCODE
Dim cmdView,txt_CONTRACTCODE
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 	txt_CONTRACTCODE=request("txt_CONTRACTCODE")
 	CONTRACTCODE=request("CONTRACTCODE")
	CUST_CUSTOMERCODE=request("CUST_CUSTOMERCODE")
	STARTDATE=request("STARTDATE")
	ENDDATE=request("ENDDATE")
	ACTUALSTART=request("ACTUALSTART")
	ACTUALEND=request("ACTUALEND")
	PRODUCT_PRODUCTCODE=request("PRODUCT_PRODUCTCODE")
	TOTALWEIGHT=request("TOTALWEIGHT")
	txtSTATUS=request("txtSTATUS")
	REFERENCE=request("REFERENCE")
	if Request("txtJenis")= "Delete" then
       q="Delete from DELIVERYCONTRACT where CONTRACTCODE='" & CONTRACTCODE & "' And STARTDATE=to_Date('" & STARTDATE & "','MM/DD/YYYY') And ENDDATE=to_Date('" & ENDDATE & "','MM/DD/YYYY') And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into DELIVERYCONTRACT (CONTRACTCODE, CUST_CUSTOMERCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, PRODUCT_PRODUCTCODE, TOTALWEIGHT, STATUS, REFERENCE, FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values ('" & CONTRACTCODE & "','" &  CUST_CUSTOMERCODE & "',to_date('" & Request("STARTDATE") & "','mm/dd/yyyy') " & ",to_date('" & Request("ENDDATE") & "','mm/dd/yyyy') " & "," & _
		" to_date('" & Request("ACTUALSTART") & "','mm/dd/yyyy') " & ",to_date('" & Request("ACTUALEND") & "','mm/dd/yyyy') " & "," & _
		"'" &  PRODUCT_PRODUCTCODE & "','" &  TOTALWEIGHT & "','" &  txtSTATUS & "','" &  REFERENCE & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		'response.Flush()
	    connStr.execute(q)
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update DELIVERYCONTRACT set CONTRACTCODE ='" & CONTRACTCODE & "'," & _
		"STARTDATE =to_date('" & Request("STARTDATE") & "','mm/dd/yyyy') " & ",ENDDATE=to_date('" & Request("ENDDATE") & "','mm/dd/yyyy') " & ",ACTUALSTART=to_date('" & Request("ACTUALSTART") & "','mm/dd/yyyy') " & ",ACTUALEND=to_date('" & Request("ACTUALEND") & "','mm/dd/yyyy') " & ",TOTALWEIGHT='" & TOTALWEIGHT & "'," & _
		"STATUS ='" & txtSTATUS & "',REFERENCE ='" & REFERENCE & "',CUST_CUSTOMERCODE ='" & CUST_CUSTOMERCODE & "',PRODUCT_PRODUCTCODE ='" & PRODUCT_PRODUCTCODE & "'" & _
		"where CONTRACTCODE='" & CONTRACTCODE & "' And STARTDATE=to_Date('" & STARTDATE & "','MM/DD/YYYY') And ENDDATE=to_Date('" & ENDDATE & "','MM/DD/YYYY') And fcBA='" & Session("fcBA") & "'"
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
<form name="frmEntry" method="post" action="DELIVERYCONTRACT.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
           <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE DELIVERY CONTRACT</strong></font></div></td>
          </TR>
		<tr>
			<td align="left">Contract Code</td><td align="left"><input type="text" size="15" name="txt_CONTRACTCODE" value="<%=REQUEST("txt_CONTRACTCODE")%>"  ></td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;
		<a href="PrintDELIVERYCONTRACT.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CONTRACTCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		
		</td>
		</tr>
    </table>  
 </form>
  <%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select CONTRACTCODE, CUST_CUSTOMERCODE, STARTDATE, ENDDATE, ACTUALSTART, ACTUALEND, " & _
     "PRODUCT_PRODUCTCODE, TOTALWEIGHT, STATUS, REFERENCE, FCENTRY, FCEDIT, FCIP, FCBA from DELIVERYCONTRACT where " & _
     " fcBA='" & session("fcBA") & "'"
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("txt_CONTRACTCODE"))) > 0 then
		q = q & " and CONTRACTCODE='" & request("txt_CONTRACTCODE") & "'"
	end if
	q = q
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <form action="DELIVERYCONTRACT.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">


		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Contract Code</td><td align="left"><input type="text" size="15" maxlength="15" name="CONTRACTCODE" value="<%=CONTRACTCODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
		<%else%>
			<td align="left">Contract Code</td><td align="left"><input type="text" size="15" maxlength="15" name="CONTRACTCODE" value="<%=CONTRACTCODE%>"  ><font color="#FF0000">*</font></td>
		<%end if%>
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
		</tr>
		<tr>
			<td align="left">Start Date</td>
			<td align="left"><input type="text" maxlength="10" name="STARTDATE" size="10" value="<%=STARTDATE%>">
			<a href="javascript:show_calendar('eForm.STARTDATE');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">End Date</td>
			<td align="left"><input type="text" maxlength="10" name="ENDDATE" size="10" value="<%=ENDDATE%>">
			<a href="javascript:show_calendar('eForm.ENDDATE');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Actual Start</td>
			<td align="left"><input type="text" maxlength="10" name="ACTUALSTART" size="10" value="<%=ACTUALSTART%>">
			<a href="javascript:show_calendar('eForm.ACTUALSTART');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Actual End</td>
			<td align="left"><input type="text" maxlength="10" name="ACTUALEND" size="10" value="<%=ACTUALEND%>">
			<a href="javascript:show_calendar('eForm.ACTUALEND');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
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
			<td align="left">Total Weight</td><td align="left"><input type="text" size="15" name="TOTALWEIGHT" value="<%=TOTALWEIGHT%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="left">Status</td><td align="left"><input type="text" size="10" maxlength="10" name="txtSTATUS" value="<%=txtSTATUS%>"><font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="left">Reference</td><td align="left"><input type="text" size="50" maxlength="100" name="REFERENCE" value="<%=REFERENCE%>"></td>
		</tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="txt_CONTRACTCODE" value="<%=REQUEST("txt_CONTRACTCODE")%>">				
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
	<TD align="left">CONTRACT</TD>	
	<TD align="left">CUSTOMER</TD>	
	<TD align="left">START DATE</TD>	
	<TD align="left">END DATE</TD>	
	<TD align="left">ACTUAL START</TD>	
	<TD align="left">ACTUAL END</TD>	
	<TD align="left">PRODUCT</TD>	
	<TD align="left">TOTAL WEIGHT</TD>	
	<TD align="left">STATUS</TD>	
	<TD align="left">REFERENCES</TD>	
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
			<a href="DELIVERYCONTRACT.asp?CONTRACTCODE=<%=rsGAD("CONTRACTCODE")%>&txttable=<%=request("txttable")%>&STARTDATE=<%=rsGAD("STARTDATE")%>&ENDDATE=<%=rsGAD("ENDDATE")%>&ACTUALSTART=<%=rsGAD("ACTUALSTART")%>&ACTUALEND=<%=rsGAD("ACTUALEND")%>&TOTALWEIGHT=<%=rsGAD("TOTALWEIGHT")%>&txtSTATUS=<%=rsGAD("STATUS")%>&REFERENCE=<%=rsGAD("REFERENCE")%>&CUST_CUSTOMERCODE=<%=rsGAD("CUST_CUSTOMERCODE")%>&PRODUCT_PRODUCTCODE=<%=rsGAD("PRODUCT_PRODUCTCODE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="DELIVERYCONTRACT.asp?CONTRACTCODE=<%=rsGAD("CONTRACTCODE")%>&txttable=<%=request("txttable")%>&STARTDATE=<%=rsGAD("STARTDATE")%>&ENDDATE=<%=rsGAD("ENDDATE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
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
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""DELIVERYCONTRACT.asp?txtModule=" &_
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
