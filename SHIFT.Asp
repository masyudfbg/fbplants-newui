<%Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Shift</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
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
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <%'-------------awal-----------------------%>
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim SHIFTCODE, STARTTIME, ENDTIME, txtDESCRIPTION
Dim cmdView
txttable=request("txttable")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 SHIFTCODE=request("SHIFTCODE")
 STARTTIME=request("STARTTIME")
 ENDTIME=request("ENDTIME")
 txtDESCRIPTION=request("txtDESCRIPTION")

	if Request("txtJenis")= "Delete" then
       q="Delete from Shift where SHIFTCODE='" & SHIFTCODE & "' And fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
			
	  q="Insert into Shift (SHIFTCODE, STARTTIME, ENDTIME, DESCRIPTION,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME ) " & _
		"values ('" & SHIFTCODE & "','" & STARTTIME & "','" & ENDTIME & "','" &  txtDESCRIPTION & "'," & _
		"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		'response.Write(q)
		'response.Flush()
	    connStr.execute(q)
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q=" Update Shift set " & _
	    "  SHIFTCODE ='" & SHIFTCODE & "'" & _
		" , STARTTIME ='" & STARTTIME & "'" & _
		" , ENDTIME='" & ENDTIME & "'" & _
		" , DESCRIPTION ='" & txtDESCRIPTION & "'" & _
		" , LASTUPDATE = sysdate " & _
		" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
		" where ShiftCODE='" & ShiftCODE & "'" & _
		" And fcBA='" & Session("fcBA") & "'"
		'response.Write(q)
		'response.Flush()
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
<% 
'   Dim rsVH
'   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
'	q="select ShiftCODE, ShiftDESCRIPTION, CROPCODE, UNITOFMEASURE, FORMULA from  Shift"
'response.Write q
'   Set rsVH = Server.CreateObject("ADODB.Recordset")
'   rsVH.CursorType = 3
'   rsVH.CursorLocation = 3
'   rsVH.Open q, connstr%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-clock-o" aria-hidden="true"></i>Shift</div><% end if %>
   <form action="Shift.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="50%">
          <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE SHIFT</strong></font></div></td>
          </TR>		
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Shift Code</td><td align="left"><input type="text" size="15" name="SHIFTCODE" value="<%=SHIFTCODE%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font> </td>
		<%else%>
			<td align="left">Shift Code</td><td align="left"><input type="text" size="15" name="SHIFTCODE" value="<%=SHIFTCODE%>"  ><font color="#FF0000">*</font> </td>
		<%end if%>
		</tr>
		<tr><td align="left">Time Start</td>
		<td align="left"><input type="text" maxlength="10" name="STARTTIME" size="10" value="<%=STARTTIME%>">
		</td></tr>	
		<tr><td align="left">Time End</td>
		<td align="left"><input type="text" maxlength="10" name="ENDTIME" size="10" value="<%=ENDTIME%>"> 
		</td></tr>	
		<tr>
		<td align="left">Description</td><td align="left"><input type="text" size="15" name="txtDESCRIPTION" value="<%=txtDESCRIPTION%>"></td>
		</tr>
	    <TR><td colspan="2" align="left">
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
   Dim rsGAD
	
   q="select * from " & txtTable & " where fcBA='" & Session("fcBA") & "'"   
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
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines" >
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">SHIFT CODE</TD>	
	<TD align="left">START TIME</TD>	
	<TD align="left">END TIME</TD>	
	<TD align="left">DESCRIPTION</TD>	
	</tr>
	</thead>
	<tbody>
   
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
			<a href="Shift.asp?SHIFTCODE=<%=rsGAD("SHIFTCODE")%>&txttable=<%=request("txttable")%>&STARTTIME=<%=rsGAD("STARTTIME")%>&ENDTIME=<%=rsGAD("ENDTIME")%>&txtDESCRIPTION=<%=rsGAD("DESCRIPTION")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="Shift.asp?SHIFTCODE=<%=rsGAD("SHIFTCODE")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
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
   </tbody>
   </TABLE>
   </div>
<%'End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
