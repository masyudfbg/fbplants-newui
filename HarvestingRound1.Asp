<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Harvesting Rounds</title>
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
DIM txt_HRNDYEAR, txt_HRNDMONTH, txt_FIELDCODE
Dim HRNDYEAR, HRNDMONTH, FIELDCODE, CROPCODE, NUMBEROFROUNDS, HECTARAGE, STANDPERHECTHARVESTED

Dim cmdView
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
txt_HRNDYEAR=request("txt_HRNDYEAR")
txt_HRNDMONTH=request("txt_HRNDMONTH")
txt_FIELDCODE=request("txt_FIELDCODE")
HRNDYEAR=request("HRNDYEAR")
HRNDMONTH=request("HRNDMONTH")
FIELDCODE=request("FIELDCODE")
CROPCODE=request("CROPCODE")
NUMBEROFROUNDS=request("NUMBEROFROUNDS")
HECTARAGE=request("HECTARAGE")
STANDPERHECTHARVESTED=request("STANDPERHECTHARVESTED")
	if Request("txtJenis")= "Delete" then

       q="Delete from HarvestingRound where  HRNDYEAR='" & HRNDYEAR & "' And HRNDMONTH='" & HRNDMONTH & "' And FIELDCODE='" & FIELDCODE & "' And CROPCODE='" & CROPCODE & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
	   'response.Write(q)
	   'response.Flush()
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then

	  q="Insert into HarvestingRound (HRNDYEAR, HRNDMONTH, FIELDCODE, CROPCODE, NUMBEROFROUNDS, HECTARAGE, STANDPERHECTHARVESTED,FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values (" & HRNDYEAR & "," & HRNDMONTH & ",'" &  FIELDCODE & "','" &  CROPCODE & "'," &  NUMBEROFROUNDS & "," &  HECTARAGE & "," &  STANDPERHECTHARVESTED & "," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		'response.Flush()
	    connStr.execute(q)
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update HarvestingRound set NUMBEROFROUNDS =" & NUMBEROFROUNDS & ",HECTARAGE=" & HECTARAGE & ",STANDPERHECTHARVESTED=" & STANDPERHECTHARVESTED & " " & _
		"where  HRNDYEAR='" & HRNDYEAR & "' And HRNDMONTH='" & HRNDMONTH & "' And FIELDCODE='" & FIELDCODE & "' And CROPCODE='" & CROPCODE & "' And fcBA='" & Session("fcBA") & "'"
'		response.Write(q)
'		response.Flush()
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
 <form name="frmEntry" method="post" action="HarvestingRound.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<TR> 
		<td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>HARVESTING ROUNDS</strong></font></div></td>
		</TR>
		<tr>
	    	<td align="left">Year</td><td align="left"><input type="text" size="10" maxlength="4" name="txt_HRNDYEAR" value="<%=REQUEST("txt_HRNDYEAR")%>"></td>
		</tr>
		<tr>
	    	<td align="left">Month</td><td align="left"><input type="text" size="10" maxlength="4" name="txt_HRNDMONTH" value="<%=REQUEST("txt_HRNDMONTH")%>"></td>
		</tr>
		<TR>
		<td align="left"><font face="Verdana" size="1">Field Code</font></td>
		<td align="left"> <select name="Txt_Fieldcode">
		<%cmdCallCombo "Field",request("Txt_Fieldcode"),"",""%>
		</select></td>
		</tr>		
			
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintHarvestingRound.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_HRNDYEAR=<%=REQUEST("txt_HRNDYEAR")%>&txt_HRNDMONTH=<%=REQUEST("txt_HRNDMONTH")%>&Txt_Fieldcode=<%=REQUEST("Txt_Fieldcode")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>

<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select HRNDYEAR, HRNDMONTH, FIELDCODE, CROPCODE, NUMBEROFROUNDS, HECTARAGE, STANDPERHECTHARVESTED, FCENTRY, FCEDIT, FCIP, FCBA from HarvestingRound where " & _
     " fcBA='" & session("fcBA") & "'"
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	if  len(trim(request("Txt_Fieldcode"))) > 0 then
		q = q & " and FIELDCODE='" & request("Txt_Fieldcode") & "'"
	end if
	if  len(trim(request("txt_HRNDYEAR"))) > 0 then
		q = q & " and HRNDYEAR='" & request("txt_HRNDYEAR") & "'"
	end if
	if  len(trim(request("txt_HRNDMONTH"))) > 0 then
		q = q & " and HRNDMONTH='" & request("txt_HRNDMONTH") & "'"
	end if
	q = q
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <form action="HarvestingRound.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
	    	<td align="left">Year</td><td align="left"><input type="text" size="10" maxlength="4" name="HRNDYEAR" readonly="True" style="background:#999999" value="<%=HRNDYEAR%>"> <font color="#FF0000">*</font></td>
		<%Else%>
	    	<td align="left">Year</td><td align="left"><input type="text" size="10" maxlength="4" name="HRNDYEAR" value="<%=HRNDYEAR%>"> <font color="#FF0000">*</font></td>
		<%End If%>
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Month</td><td align="left"><input type="text" size="10" maxlength="2" name="HRNDMONTH" readonly="True" style="background:#999999" value="<%=HRNDMONTH%>"> <font color="#FF0000">*</font></td>
		<%Else%>
			<td align="left">Month</td><td align="left"><input type="text" size="10" maxlength="2" name="HRNDMONTH" value="<%=HRNDMONTH%>"> <font color="#FF0000">*</font></td>
		<%End if%>
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left"><font face="Verdana" size="1">Field Code</font></td>
			<td align="left"> <select name="FIELDCODE">
			<%cmdCallCombo "Field",request("FIELDCODE"),"",""%>
			</select> <font color="#FF0000">*</font> </td>
		<%Else%>
				<td align="left"><font face="Verdana" size="1">Field Code</font></td>
				<td align="left"> <select name="FIELDCODE">
				<%cmdCallCombo "Field","fccode","",""%>
				</select> <font color="#FF0000">*</font> </td>
		<%End if%>
		</TR>
		<TR>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left"><font face="Verdana" size="1">Crop Code</font></td>
				<% 	Dim rsCrop
				q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' order by P.FCNAME desc"
				Set rsCrop = Server.CreateObject("ADODB.Recordset")
				rsCrop.CursorType = 3
				rsCrop.CursorLocation = 3
				rsCrop.Open q, connstr
				if not rsCrop.eof then%>
				<td align="left"> <select name="CropCode" title="Crop Character a s/d z or 1 s/d 9">         
				<option value=""></option> 
				<%Do While Not rsCrop.eof
				If rsCrop("FCCODE") = request("CropCode") Then%>
					<option value="<%=rsCrop("FCCODE")%>" selected><%=rsCrop("FCNAME")%></option> 
				<%Else%>
					<option value="<%=rsCrop("FCCODE")%>"><%=rsCrop("FCNAME")%></option> 
				<%End if%>				
				<%rsCrop.movenext
				loop
				End if
				rsCrop.Close
				set rsCrop = Nothing
				q=""
				%>
			</select> <font color="#FF0000">*</font></td>
		<%Else%>
				<td align="left"><font face="Verdana" size="1">Crop Code</font></td>
				<% 	Dim rsCropA
				q = "select P.FCCODE,P.FCNAME from  PARAMETER P  where P.PARCODE= 'F0005' order by P.FCNAME desc"
				Set rsCropA = Server.CreateObject("ADODB.Recordset")
				rsCropA.CursorType = 3
				rsCropA.CursorLocation = 3
				rsCropA.Open q, connstr
				if not rsCropA.eof then%>
				<td align="left"> <select name="CropCode" title="Crop Character a s/d z or 1 s/d 9">         
				<option value=""></option> 
				<%Do While Not rsCropA.eof%>
					<option value="<%=rsCropA("FCCODE")%>"><%=rsCropA("FCNAME")%></option> 
				<%rsCropA.movenext
				loop
				End if
				rsCropA.Close
				set rsCropA = Nothing
				q=""
				%>
			</select> <font color="#FF0000">*</font></td>
		<%End if%>
		</tr>
		<tr>
		<td align="left">Number Of Rounds</td><td align="left"><input type="text" size="15" name="NUMBEROFROUNDS" value="<%=NUMBEROFROUNDS%>"></td>
		</tr>
		<tr>
		<td align="left">Hectarage</td><td align="left"><input type="text" size="15" name="HECTARAGE" value="<%=HECTARAGE%>"></td>
		</tr>
		<tr>
		<td align="left">Standart Perhectharvested</td><td align="left"><input type="text" size="15" name="STANDPERHECTHARVESTED" value="<%=STANDPERHECTHARVESTED%>"></td>
		</tr>					  
	    <TR><td colspan="2">
		  <input type="hidden" name="txt_HRNDYEAR" value="<%=REQUEST("txt_HRNDYEAR")%>">
		  <input type="hidden" name="txt_HRNDMONTH" value="<%=REQUEST("txt_HRNDMONTH")%>">
		  <input type="hidden" name="Txt_Fieldcode" value="<%=REQUEST("Txt_Fieldcode")%>">
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
'   Dim rsGAD
	
'   q="select * from " & txtTable & ""   
'   Set rsGAD = Server.CreateObject("ADODB.Recordset")
	
'   rsGAD.CursorType = 3
'   rsGAD.CursorLocation = 3
'   rsGAD.Open q, connstr
   JumRec=1
   intRecordsPerPage = 20
	
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
   <table class="data">
	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">YEAR</TD>	
	<TD align="left">MONTH</TD>	
	<TD align="left">FIELD CODE</TD>	
	<TD align="left">CROP CODE</TD>	
	<TD align="left">NUMBER OF ROUND</TD>	
	<TD align="left">HECTARAGE</TD>	
	<TD align="left">STANDART PERHECTHARVESTED</TD>	
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
		  <td width="75" align="left"><%=NO%></td>
			<td width="25" align="left">
			<a href="HarvestingRound.Asp?HRNDYEAR=<%=rsGAD("HRNDYEAR")%>&txttable=<%=request("txttable")%>&HRNDMONTH=<%=rsGAD("HRNDMONTH")%>&FIELDCODE=<%=rsGAD("FIELDCODE")%>&CROPCODE=<%=rsGAD("CROPCODE")%>&NUMBEROFROUNDS=<%=rsGAD("NUMBEROFROUNDS")%>&HECTARAGE=<%=rsGAD("HECTARAGE")%>&STANDPERHECTHARVESTED=<%=rsGAD("STANDPERHECTHARVESTED")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td width="28" align="left">
			<a href="HarvestingRound.Asp?HRNDYEAR=<%=rsGAD("HRNDYEAR")%>&txttable=<%=request("txttable")%>&HRNDMONTH=<%=rsGAD("HRNDMONTH")%>&FIELDCODE=<%=rsGAD("FIELDCODE")%>&CROPCODE=<%=rsGAD("CROPCODE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="658" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
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
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""HarvestingRound.Asp?txtModule=" &_
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
