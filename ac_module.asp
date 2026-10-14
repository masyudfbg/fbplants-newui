<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Access Control Module</title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
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


' Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA
' DIM v_NURSERYCODE,v_DATEPLANTED
' DIM v_DESCRIPTION,  v_CROP, v_PLANTINGMATERIAL, v_CLONE, v_SOURCE, v_TYPE, v_STATUS, v_StoreCode, v_StockCode
' Dim cmdView
DOCUMENT_TYPE =request("DOCUMENT_TYPE")
COLUMN_NAME =request("COLUMN_NAME")
' v_DESCRIPTION=request("v_DESCRIPTION")
' v_DATEPLANTED=request("v_DATEPLANTED")
' v_CROP=request("v_CROP")
' v_PLANTINGMATERIAL=request("v_PLANTINGMATERIAL")
' v_CLONE=request("v_CLONE")
' v_SOURCE=request("v_SOURCE")
' v_TYPE=request("v_TYPE")
' v_StoreCode=request("v_StoreCode")
' v_StockCode=request("v_StockCode")
' v_STATUS=request("v_STATUS")

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	" Delete from ac_module where" & _
				" DOCUMENT_TYPE = '" & DOCUMENT_TYPE & "'"
				
		'Response.Write q
		connStr.execute(q)
		cmdView	= "True"
	elseif Request("txtJenis")="New" then	
		q	=	" Insert into ac_module (" & _
				" DOCUMENT_TYPE" & _
				" , COLUMN_NAME" & _
				" , FCENTRY" & _
				" , FCEDIT" & _
				" , FCIP" & _
				" , LASTUPDATE" & _
				" , LASTTIME" & _
				" )" & _
				" values (" & _
				" '" & DOCUMENT_TYPE & "'" & _
				" , '" & COLUMN_NAME & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _		
				" , sysdate" & _
				" , To_Char(sysdate,'HH24:MI')" & _				
				" )"
				
		'Response.Write q
	    connStr.execute(q)		
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q	=	" Update ac_module set" & _
				" COLUMN_NAME = '" & COLUMN_NAME & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
				" where" & _
				" DOCUMENT_TYPE = '" & DOCUMENT_TYPE & "'"
		'Response.Write q
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
 
 <form name="frmEntry" method="post" action="ac_module.asp">
	<table class="data" width="100%">
		<tr><td align="left" width="20%">Document Type</td>
		<td align="left"><input type="text" maxlength="10" name="DOCUMENT_TYPE" size="10" value="<%=DOCUMENT_TYPE%>" autocomplete = "Off"></td></tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="ac_module.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&v_NURSERYCODE=<%=v_NURSERYCODE%>&v_DATEPLANTED=<%=v_DATEPLANTED%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" and len(trim(Document_type)) > 0 Then
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	=	" select" & _
			" DOCUMENT_TYPE" & _
			" , COLUMN_NAME" & _
			" , FCENTRY" & _
			" , FCEDIT" & _
			" , FCIP" & _
			" from  ac_module where" & _
			" DOCUMENT_TYPE ='" & DOCUMENT_TYPE & "'"
			
' If Len(Trim(v_NurseryCode)) > 0 Then
' 	q	=	q & " And Nurserycode = '" & v_NurseryCode & "'"
' End If

	' q	=	q & " Order By Nurserycode"
   'response.Write q
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="ac_module.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<%if Request("txtJenis")="Edit" then%>
			<td align="left" width="20%">Document Type</td><td align="left"><input type="text" size="10" name="DOCUMENT_TYPE" value="<%=DOCUMENT_TYPE%>" maxlength="50" readonly style="background: #eee;"></td>
			<%else%>
			<td align="left" width="20%">Document Type</td><td align="left"><input type="text" size="10" name="DOCUMENT_TYPE" maxlength="50"></td>
			<%End if%>
		</tr>
		<tr><td align="left" width="20%">Column Name</td><td align="left"><input type="text" size="40" name="COLUMN_NAME" value="<%=COLUMN_NAME%>"></td></tr>
		<tr>
		<td colspan="2" align="left">
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
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
		</td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DOCUMENT TYPE</TD>	
	<TD align="left">COLUMN NAME</TD>	
	</tr>
   
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">

			<a href="ac_module.asp?DOCUMENT_TYPE=<%=rsVH("DOCUMENT_TYPE")%>&COLUMN_NAME=<%=rsVH("COLUMN_NAME")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("DOCUMENT_TYPE")%>" border="0"></a></td><td>
			<a href="ac_module.asp?DOCUMENT_TYPE=<%=rsVH("DOCUMENT_TYPE")%>&COLUMN_NAME=<%=rsVH("COLUMN_NAME")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("DOCUMENT_TYPE")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
		    <%for i = 0 to nCountofField-3%>
			    <td align="center" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
		    <%next%>
		
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
