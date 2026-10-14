<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Shift Report</title>
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
dim txt_SHIFTCODE,Perioddate
Dim PRODUCTIONDATE, SHIFT_SHIFTCODE, ACTUALSTART, ACTUALEND, CAGESPROCESSED, CAGEWEIGHT, RAMPBALANCE, CAGESBALANCE, PLATFORMBALANCE, DOWNTIME, CAUSE
Dim cmdView
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
Perioddate=request("Perioddate")
txt_SHIFTCODE=request("txt_SHIFTCODE")
PRODUCTIONDATE=request("PRODUCTIONDATE")
SHIFT_SHIFTCODE=request("SHIFT_SHIFTCODE")
ACTUALSTART=request("ACTUALSTART")
ACTUALEND=request("ACTUALEND")
CAGESPROCESSED=request("CAGESPROCESSED")
CAGEWEIGHT=request("CAGEWEIGHT")
RAMPBALANCE=request("RAMPBALANCE")
CAGESBALANCE=request("CAGESBALANCE")
PLATFORMBALANCE=request("PLATFORMBALANCE")
DOWNTIME=request("DOWNTIME")
CAUSE=request("CAUSE")
	if Request("txtJenis")= "Delete" then
       q="Delete from SHIFTREPORT where PRODUCTIONDATE=to_Date('" & PRODUCTIONDATE & "','MM/DD/YYYY') And ACTUALSTART='" & ACTUALSTART & "' And ACTUALEND='" & ACTUALEND & "' And fcBA='" & Session("fcBA") & "'"
	   'response.Write(q)
	   'response.Flush()
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
	  q="Insert into SHIFTREPORT (PRODUCTIONDATE, SHIFT_SHIFTCODE, ACTUALSTART, ACTUALEND, CAGESPROCESSED, CAGEWEIGHT, RAMPBALANCE, CAGESBALANCE, PLATFORMBALANCE, DOWNTIME, CAUSE,FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values (to_Date('" & PRODUCTIONDATE & "','MM/DD/YYYY'),'" &  SHIFT_SHIFTCODE & "','" & ACTUALSTART & "','" & ACTUALEND & "'," &  CAGESPROCESSED & "," & _
		"" &  CAGEWEIGHT & "," &  RAMPBALANCE & "," &  CAGESBALANCE & "," &  PLATFORMBALANCE & ",'" & DOWNTIME & "','" &  CAUSE & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		'Response.Flush()
	    connStr.execute(q)
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then 
       q="Update SHIFTREPORT set CAGESPROCESSED =" & CAGESPROCESSED & ",RAMPBALANCE =" & RAMPBALANCE & ",CAGESBALANCE =" & CAGESBALANCE & "," & _
		"PLATFORMBALANCE =" & PLATFORMBALANCE & ",DOWNTIME ='" & DOWNTIME & "',CAUSE ='" & CAUSE & "',SHIFT_SHIFTCODE ='" & SHIFT_SHIFTCODE & "',CAGEWEIGHT =" & CAGEWEIGHT & "" & _
		"where PRODUCTIONDATE=to_Date('" & PRODUCTIONDATE & "','MM/DD/YYYY') " & " And ACTUALSTART='" & ACTUALSTART & "' And ACTUALEND='" & ACTUALEND & "' And fcBA='" & Session("fcBA") & "'"
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

 <form name="frmEntry" method="post" action="SHIFTREPORT.asp">
	<table class="data" width="100%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE SHIFT REPORT</strong></font></div></td>
          </TR>

		<tr><td align="left">Date</td>
		<td align="left"><input type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		<a href="javascript:show_calendar('frmEntry.Perioddate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<TR>
			<td align="left">Shift Code</td>
			<td align="left"> <select name="txt_SHIFTCODE">
			<%cmdCallComboSetup "SHIFT",REQUEST("txt_SHIFTCODE"),"",""%>
			</select></td>
		</tr>		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintSHIFTREPORT.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_SHIFTCODE=<%=REQUEST("txt_SHIFTCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select PRODUCTIONDATE, SHIFT_SHIFTCODE, ACTUALSTART, ACTUALEND, CAGESPROCESSED, CAGEWEIGHT, " & _
      "RAMPBALANCE, CAGESBALANCE, PLATFORMBALANCE, DOWNTIME, CAUSE, FCENTRY, FCEDIT, FCIP, FCBA from SHIFTREPORT where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(request("Perioddate"))) > 0 then
		q = q & " and to_char(PRODUCTIONDATE,'mm/yyyy')=to_char(to_date('" & request("Perioddate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_SHIFTCODE"))) > 0 then
		q = q & " and SHIFT_SHIFTCODE='" & request("txt_SHIFTCODE") & "'"
	end if
	q = q 
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>

   <form action="SHIFTREPORT.asp" method="post" name="eForm">
   <table class="data" width="100%" cellpadding="0" cellspacing="0">
		<td align="left">Date</td>
		<%IF txtJenis="Edit" Then%>
				<td align="left"><input type="text" maxlength="10" name="PRODUCTIONDATE" size="10" value="<%=PRODUCTIONDATE%>">
		<%Else%>
				<td align="left"><input type="text" maxlength="10" name="PRODUCTIONDATE" size="10" value="<%=DATE()%>">
		<%End if%>
				<a href="javascript:show_calendar('eForm.PRODUCTIONDATE');" onMouseOver="window.status='Date Picker';return true;" 
				onmouseout="window.status='';return true;">
				<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
			<td align="left">Shift Code</td>
			<td align="left"> <select name="SHIFT_SHIFTCODE">
			<%cmdCallComboSetup "SHIFT",request("SHIFT_SHIFTCODE"),"",""%>
			</select> <font color="#FF0000">*</font> </td>
		<%Else%>
			<td align="left">Shift Code</td>
			<td align="left"> <select name="SHIFT_SHIFTCODE">
			<%cmdCallComboSetup "SHIFT","SHIFT_SHIFTCODE","",""%>
			</select> <font color="#FF0000">*</font> </td>
		<%End if%>
		</tr>
		<tr>
			<td align="left">Actual Start</td>
			<td align="left"><input type="text" maxlength="10" name="ACTUALSTART" size="10" value="<%=ACTUALSTART%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="left">Actual End</td>
			<td align="left"><input type="text" maxlength="10" name="ACTUALEND" size="10" value="<%=ACTUALEND%>"><font color="#FF0000">*</font></td>
		</TR>
		<TR>
		<td align="left">Cages Processed</td><td align="left"><input type="text" size="15" name="CAGESPROCESSED" value="<%=CAGESPROCESSED%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Cages Weight (Kg)</td><td align="left"><input type="text" size="15" name="CAGEWEIGHT" value="<%=CAGEWEIGHT%>"></td>
		</tr>

		<tr>
		<td align="left">Ramp Balance (Kg)</td><td align="left"><input type="text" size="15" name="RAMPBALANCE" value="<%=RAMPBALANCE%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Cages Balance (Kg)</td><td align="left"><input type="text" size="15" name="CAGESBALANCE" value="<%=CAGESBALANCE%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Plat Form Balance (Kg)</td><td align="left"><input type="text" size="15" name="PLATFORMBALANCE" value="<%=PLATFORMBALANCE%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Down Time</td><td align="left"><input type="text" size="15" name="DOWNTIME" value="<%=DOWNTIME%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Cause</td><td align="left"><input type="text" size="50" maxlength="250" name="CAUSE" value="<%=CAUSE%>"></td>
		</tr>
	    <TR><td colspan="2">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_SHIFTCODE" value="<%=txt_SHIFTCODE%>">
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
	<TD align="left">DATE</TD>	
	<TD align="left">SHIFT CODE</TD>	
	<TD align="left">ACTUAL START</TD>	
	<TD align="left">ACTUAL END</TD>	
	<TD align="left">CAGES PROCESSED</TD>	
	<TD align="left">CAGES WEIGHT</TD>	
	<TD align="left">RAMP BALANCES</TD>	
	<TD align="left">CAGE BALANCES</TD>	
	<TD align="left">PLAT FROM BALANCES</TD>	
	<TD align="left">DOWN TIME</TD>	
	<TD align="left">CAUSE</TD>	
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
			
			<a href="SHIFTREPORT.asp?PRODUCTIONDATE=<%=rsGAD("PRODUCTIONDATE")%>&txttable=<%=request("txttable")%>&ACTUALSTART=<%=rsGAD("ACTUALSTART")%>&ACTUALEND=<%=rsGAD("ACTUALEND")%>&CAGESPROCESSED=<%=rsGAD("CAGESPROCESSED")%>&RAMPBALANCE=<%=rsGAD("RAMPBALANCE")%>&CAGESBALANCE=<%=rsGAD("CAGESBALANCE")%>&PLATFORMBALANCE=<%=rsGAD("PLATFORMBALANCE")%>&DOWNTIME=<%=rsGAD("DOWNTIME")%>&CAUSE=<%=rsGAD("CAUSE")%>&SHIFT_SHIFTCODE=<%=rsGAD("SHIFT_SHIFTCODE")%>&CAGEWEIGHT=<%=rsGAD("CAGEWEIGHT")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="SHIFTREPORT.asp?PRODUCTIONDATE=<%=rsGAD("PRODUCTIONDATE")%>&txttable=<%=request("txttable")%>&ACTUALSTART=<%=rsGAD("ACTUALSTART")%>&ACTUALEND=<%=rsGAD("ACTUALEND")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
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
			  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""SHIFTREPORT.asp?txtModule=" &_
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
