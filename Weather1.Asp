<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Weather</title>
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
DIM PeriodDate, txt_DIVISION
Dim WIDATE, WIDIVISION, TOTRNINCHES, TOTRNDAYS, AVGTEMP, AVGWIND
Dim SEALNO
Dim cmdView
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	PeriodDate =request("PeriodDate")
	txt_DIVISION =request("txt_DIVISION")
	WIDATE=request("WIDATE")
	WIDIVISION=request("WIDIVISION")
	TOTRNINCHES=request("TOTRNINCHES")
	TOTRNDAYS=request("TOTRNDAYS")
	AVGTEMP=request("AVGTEMP")
	AVGWIND=request("AVGWIND")

	if Request("txtJenis")= "Delete" then
       q="Delete from Weather where WIDATE=to_Date('" & WIDATE & "','MM/DD/YYYY') And WIDIVISION='" & WIDIVISION & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into Weather (WIDATE, WIDIVISION, TOTRNINCHES, TOTRNDAYS, AVGTEMP, AVGWIND,FCENTRY, FCEDIT, FCIP, FCBA) " & _
		"values (to_Date('" & WIDATE & "','MM/DD/YYYY'),'" &  WIDIVISION & "','" & TOTRNINCHES & "','" &  TOTRNDAYS & "','" &  AVGTEMP & "'," & _
		"'" & AVGWIND & "'," & _		
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
		'response.Write(q)
		'response.Flush()
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q="Update Weather set WIDATE =to_Date('" & WIDATE & "','MM/DD/YYYY'),WIDIVISION ='" & WIDIVISION & "'," & _
		"TOTRNINCHES ='" & TOTRNINCHES & "',TOTRNDAYS='" & TOTRNDAYS & "',AVGTEMP='" & AVGTEMP & "',AVGWIND ='" & AVGWIND & "'" & _
		"where WIDATE=to_Date('" & WIDATE & "','MM/DD/YYYY') And WIDIVISION='" & WIDIVISION & "' And fcBA='" & Session("fcBA") & "'"
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
  <form name="frmEntry" method="post" action="Weather.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<TR> 
		<td colspan="2"><div align="center"><font size="2" face="Verdana"><strong>MAINTENANCE WEATHER</strong></font></div></td>
		</TR>
		<tr><td align="left">Periode</td>
		<td align="left"><input type="text" maxlength="10" name="PeriodDate" size="10" value="<%=PeriodDate%>">
		<a href="javascript:show_calendar('frmEntry.PeriodDate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>
		</td></tr>
		<tr>
	    	<td align="left">Division</td><td align="left"><input type="text" size="10" maxlength="15" name="txt_DIVISION" value="<%=REQUEST("txt_DIVISION")%>"></td>
		</tr>		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" style="background-color: #298CF7; color: #F7C610;
		font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintWeather.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&DIVISION=<%=REQUEST("txt_DIVISION")%>&PeriodDate=<%=REQUEST("PeriodDate")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>

<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select WIDATE, WIDIVISION, TOTRNINCHES, TOTRNDAYS, AVGTEMP, AVGWIND, FCENTRY, FCEDIT, FCIP, FCBA from WEATHER where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(request("PeriodDate"))) > 0 then
		q = q & " and to_char(WIDATE,'mm/yyyy')=to_char(to_date('" & request("PeriodDate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')"
	end if
	if  len(trim(request("txt_DIVISION"))) > 0 then
		q = q & " and WIDIVISION='" & request("txt_DIVISION") & "'"
	end if
	q = q	 
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
   <form action="Weather.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		<tr>
		<td align="left">Date</td>
			<td align="left"><input type="text" maxlength="10" name="WIDATE" size="10" value="<%=WIDATE%>">
			<a href="javascript:show_calendar('eForm.WIDATE');" onMouseOver="window.status='Date Picker';return true;" 
			onmouseout="window.status='';return true;">
			<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a><font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
		<td align="left">Division</td><td align="left"><input type="text" size="15" maxlength="15" name="WIDIVISION" value="<%=WIDIVISION%>"><font color="#FF0000">*</font></td>
		</tr>
		<TR>
		<td align="left">Rain Inches</td><td align="left"><input type="text" size="15" name="TOTRNINCHES" value="<%=TOTRNINCHES%>"></td>
		</tr>
		<tr>
		<td align="left">Rain Days</td><td align="left"><input type="text" size="15" name="TOTRNDAYS" value="<%=TOTRNDAYS%>"></td>
		</tr>
		<tr>
		<td align="left">Temperature</td><td align="left"><input type="text" size="15" name="AVGTEMP" value="<%=AVGTEMP%>"></td>
		</tr>
		<tr>
		<td align="left">Wind Speed</td><td align="left"><input type="text" size="15" name="AVGWIND" value="<%=AVGWIND%>"></td>
		</tr>
	    <TR><td colspan="2">

		  <input type="hidden" name="PeriodDate" value="<%=PeriodDate%>">
		  <input type="hidden" name="Txt_Division" value="<%=Txt_Division%>">
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
   <table width="737" class="data" >

	<tr bgcolor="#6666FF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">DEVISION</TD>	
	<TD align="left">RAIN INCHES</TD>	
	<TD align="left">RAIN DAYS</TD>	
	<TD align="left">TEMPERATURE</TD>	
	<TD align="left">WIND SPEED</TD>	
	</TR>

        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">
	
			<a href="Weather.asp?WIDATE=<%=rsGAD("WIDATE")%>&txttable=<%=request("txttable")%>&WIDIVISION=<%=rsGAD("WIDIVISION")%>&TOTRNINCHES=<%=rsGAD("TOTRNINCHES")%>&TOTRNDAYS=<%=rsGAD("TOTRNDAYS")%>&AVGTEMP=<%=rsGAD("AVGTEMP")%>&AVGWIND=<%=rsGAD("AVGWIND")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="Weather.asp?WIDATE=<%=rsGAD("WIDATE")%>&txttable=<%=request("txttable")%>&WIDIVISION=<%=rsGAD("WIDIVISION")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
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
