<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Harvesting Rounds</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
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
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal!------------->
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
DIM header_year, header_month
Dim FIELDCODE, CROPCODE, NUMBEROFROUNDS, HECTARAGE, STANDPERHECTHARVESTED
Dim cmdView

header_year=request("header_year")
header_month=request("header_month")
FIELDCODE=request("FIELDCODE")
CropCode=request("CropCode")
NUMBEROFROUNDS=request("NUMBEROFROUNDS")
HECTARAGE=request("HECTARAGE")
STANDPERHECTHARVESTED=request("STANDPERHECTHARVESTED")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	" DELETE FROM HARVESTINGROUND WHERE" & _
				" HRNDYEAR = '" & header_year & "'" & _
				" And HRNDMONTH = '" & header_month & "'" & _
				" And FIELDCODE = '" & FIELDCODE & "'" & _
				" And CROPCODE = '" & CROPCODE & "'" & _
				" And fcBA = '" & Session("fcBA") & "'"
				
		'response.Write(q)
		'response.Flush()
		connStr.execute(q)
		cmdView	= "True"

	elseif Request("txtJenis")="New" then
		If CheckIsTM(FIELDCODE, Session("fcBA")) > 0 Then
			q	=	" INSERT INTO HARVESTINGROUND (" & _
					" HRNDYEAR" & _
					" , HRNDMONTH" & _
					" , FIELDCODE" & _
					" , CROPCODE" & _
					" , NUMBEROFROUNDS" & _
					" , HECTARAGE" & _
					" , STANDPERHECTHARVESTED" & _
					" , FCENTRY" & _
					" , FCEDIT" & _
					" , FCIP" & _
					" , FCBA" & _
					" , LASTUPDATE" & _
					" , LASTTIME" & _
					" ) values (" & _
					"" & header_year & "" & _
					" , " & header_month & "" & _
					" , '" & FIELDCODE & "'" & _
					" , '" & CROPCODE & "'" & _
					" , " & NUMBEROFROUNDS & "" & _
					" , " & HECTARAGE & "" & _
					" , " & STANDPERHECTHARVESTED & "" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					" , '" & Session("fcBA") & "'" & _
					" , sysdate" & _
					" , To_Char(sysdate,'HH24:MI')" & _
					" )"
				
			'response.Write(q)
			'response.Flush()
			connStr.execute(q)
		Else
			err.description = "Blok " & FIELDCODE & " berstatus TBM"
			err.number = 10001
		End If
		
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q	=	" Update HarvestingRound set" & _
				" NUMBEROFROUNDS = " & NUMBEROFROUNDS & "" & _
				" , HECTARAGE = " & HECTARAGE & "" & _
				" , STANDPERHECTHARVESTED =" & STANDPERHECTHARVESTED & "" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
				" where HRNDYEAR = '" & header_year & "'" & _
				" And HRNDMONTH = '" & header_month & "'" & _
				" And FIELDCODE = '" & FIELDCODE & "'" & _
				" And CROPCODE = '" & CROPCODE & "'" & _
				" And fcBA = '" & Session("fcBA") & "'"
				
		'response.Write(q)
		'response.Flush()
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

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

txtJenis = request("txtJenis")
If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if
%>
 <!-------------ISI Payment Form!------------->
 <form name="frmEntry" method="post" action="HarvestingRound.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
		<tr>
			<td colspan="2">
				<div align="left"><font size="2" face="Verdana"><strong>HARVESTING ROUNDS</strong></font></div>
			</td>
			
		</tr>
		<tr>
	    	<td align="left">Year</td>
			<td align="left">
				<input type="text" size="10" maxlength="4" name="header_year" value="<%=header_year%>">
			</td>
		</tr>
		<tr>
	    	<td align="left">Month</td>
			<td align="left">
				<input type="text" size="10" maxlength="4" name="header_month" value="<%=header_month%>">
			</td>
		</tr>	
			
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintHarvestingRound.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&header_year=<%=header_year%>&header_month=<%=header_month%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		<!-- HADIR -->
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Harvesting"
		dataSubmodule 	= "Harvesting Rounds"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>

<%if cmdView = "True" And len(trim(header_year)) > 0 And len(trim(header_month)) > 0 Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	=	" select" & _
			" HRNDYEAR" & _
			" , HRNDMONTH" & _
			" , FIELDCODE" & _
			" , CROPCODE" & _
			" , NUMBEROFROUNDS" & _
			" , HECTARAGE" & _
			" , STANDPERHECTHARVESTED" & _
			" FROM HARVESTINGROUND" & _
			" where" & _
			" fcBA = '" & session("fcBA") & "'" & _
			" and HRNDYEAR = '" & header_year & "'" & _
			" and HRNDMONTH = '" & header_month & "'" & _
			" Order By HRNDYEAR, HRNDMONTH, FIELDCODE"
			
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>
	<form action="HarvestingRound.Asp" method="post" name="eForm">
	<table class="data" cellpadding="0" cellspacing="0" width="70%">
		<tr>
			<td align="left">Field Code</td>
			<td align="left">
				<select name="FIELDCODE" <%IF Request("txtJenis")="Edit" then%> disabled="disabled"> <%End If%>>
					<%cmdCallCombo "Field",FIELDCODE,"",""%>
				</select><font color="#FF0000">*</font>
				<%IF Request("txtJenis")="Edit" then%> 
				<input type="hidden" name="FIELDCODE" value="<%=FIELDCODE%>" />
				<%End If%>
			</td>
		</tr>
		
		<tr>
			<td align="left">Crop Code</td>
			<td align="left">
				<select name="CropCode" <%IF Request("txtJenis")="Edit" then%> disabled="disabled"> <%End If%>>
					<option value=""></option> 
				<%
					Dim rsCrop, strCrop
					strCrop =	" SELECT P.FCCODE,P.FCNAME FROM  PARAMETER P  WHERE P.PARCODE= 'F0005' AND FCBA = '" &  Session("fcBA") & "'" & _
								" ORDER BY P.FCNAME DESC"
					
					Set rsCrop = Server.CreateObject("ADODB.Recordset")
					rsCrop.Open strCrop, connstr
					
					If Not rsCrop.EOF then
						Do While Not rsCrop.EOF
				%>
					<option value="<%=rsCrop("FCCODE")%>" <%If CropCode = rsCrop("FCCODE") Then%> selected <%End If%>><%=rsCrop("FCNAME")%></option> 
				<%			rsCrop.movenext
						Loop
					End if
					rsCrop.Close
					set rsCrop = Nothing
				%>
				</select><font color="#FF0000">*</font>
				
				<%IF Request("txtJenis")="Edit" then%> 
				<input type="hidden" name="CropCode" value="<%=CropCode%>" />
				<%End If%>
			</td>
		</tr>
		
		<tr>
			<td align="left">Number Of Rounds</td>
			<td align="left">
				<input type="text" size="15" name="NUMBEROFROUNDS" value="<%=NUMBEROFROUNDS%>">
			</td>
		</tr>
		
		<tr>
			<td align="left">Hectarage</td>
			<td align="left">
				<input type="text" size="15" name="HECTARAGE" value="<%=HECTARAGE%>">
			</td>
		</tr>
		
		<tr>
			<td align="left">Standard Perhectharvested</td>
			<td align="left">
				<input type="text" size="15" name="STANDPERHECTHARVESTED" value="<%=STANDPERHECTHARVESTED%>">
			</td>
		</tr>					  
	    <tr><td colspan="2">
		  <input type="hidden" name="header_year" value="<%=header_year%>">
		  <input type="hidden" name="header_month" value="<%=header_month%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
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
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<td align="left">NO</td>
	<td colspan="2">&nbsp;</td>
	<td align="left">YEAR</td>	
	<td align="left">MONTH</td>	
	<td align="left">FIELD CODE</td>	
	<td align="left">CROP CODE</td>	
	<td align="left">NUMBER OF ROUND</td>	
	<td align="left">HECTARAGE</td>	
	<td align="left">STANDARD PERHECTHARVESTED</td>	
	</tr>
   <%
   if Not rsGAD.eof then
      nCountofField	=	rsGAD.fields.count-1
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="75" align="left"><%=NO%></td>
			<td width="25" align="left">
			<a href="HarvestingRound.Asp?header_year=<%=rsGAD("HRNDYEAR")%>&header_month=<%=rsGAD("HRNDMONTH")%>&FIELDCODE=<%=rsGAD("FIELDCODE")%>&CROPCODE=<%=rsGAD("CROPCODE")%>&NUMBEROFROUNDS=<%=rsGAD("NUMBEROFROUNDS")%>&HECTARAGE=<%=rsGAD("HECTARAGE")%>&STANDPERHECTHARVESTED=<%=rsGAD("STANDPERHECTHARVESTED")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td width="28" align="left">
			<a href="HarvestingRound.Asp?header_year=<%=rsGAD("HRNDYEAR")%>&header_month=<%=rsGAD("HRNDMONTH")%>&FIELDCODE=<%=rsGAD("FIELDCODE")%>&CROPCODE=<%=rsGAD("CROPCODE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
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
   </div>
   </TABLE>
<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
