<%'Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>iPS - Access Rights</title>
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
<script language="JavaScript" src="/js/date-picker.js"></script> 
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
 <%'-------------awal-----------------------%>
 <%


Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA
DIM v_FCCODE
Dim cmdView
v_FCCODE =request("v_FCCODE")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
       q="Delete from AccessRights where  fccode='" & v_FCCODE & "' and fcBA='" & Session("fcBA") & "'"
   	   connStr.execute(q)
       cmdView	= "True"
	elseif Request("txtJenis")="New" or Request("txtJenis")="Edit" then
	
        q="Delete from AccessRights where  fccode='" & v_FCCODE & "' and fcBA='" & Session("fcBA") & "'"
   	    connStr.execute(q)

	    Dim rsModule,qModule,qInsert
		qModule="select module.code as ModuleCode, module.NAME as ModuleName, sub.code as SubModuleCode, sub.NAME as SubModuleName" & _
          " FROM module module, submodule sub " & _
          " WHERE ((module.code = sub.module)) order by module.code,sub.code"
		Set rsModule = Server.CreateObject("ADODB.Recordset")
		rsModule.CursorType = 3
		rsModule.CursorLocation = 3
		rsModule.Open qModule, connstr
		if Not rsModule.eof then
		   q="Delete from AccessRights where  fccode='" & v_FCCODE & "' and fcBA='" & Session("fcBA") & "'"
		   connStr.execute(q)
		   connStr.execute("commit")
		   RESPONSE.Write (qInsert) & "<BR>"
		   RESPONSE.Flush()
   	       While Not rsModule.EOF
		      if request(trim(rsModule("ModuleCode")&rsModule("SubModuleCode")))="on" Then
				qInsert="Insert into ACCESSRIGHTS (FCCODE, MODULECODE, SUBMODULECODE, AUTHORIZED, FCENRTY, FCEDIT, FCIP,FCBA) " & _
				"values ('" & v_FCCODE & "','" & rsModule("ModuleCode") & "','" & rsModule("SubModuleCode") & "','" &  "Y" & "'," & _
				"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "')"
				'RESPONSE.Write rsModule("ModuleCode") & rsModule("SubModuleCode")  & "<BR>"
				'RESPONSE.Flush()
				connStr.execute(qInsert) 
				connStr.execute("commit") 
		      End if
		      rsModule.MoveNext
		   wend
		end if
		response.Write "<font face=Verdana size=1 color=blue>Accessright Completed..</font>"
	    rsModule.Close
		set rsModule = Nothing
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
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
 
 <form name="frmEntry" method="post" action="AccessRights.asp">
	<table class="data" width="500">
		<tr><td align="left">User Name</td>
		<td align="left"><input type="text" maxlength="15" name="v_FCCODE" size="15" value="<%=v_FCCODE%>"></td></tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		</td>
		</tr>
    </table>  
 </form>
<%
dim qCekUser,rsCekUser
qCekUser="select fccode from USERPROFILE where  fccode='" & v_FCCODE & "' and fcBA='" & Session("fcBA") & "'"
Set rsCekUser = Server.CreateObject("ADODB.Recordset")
rsCekUser.CursorType = 3
rsCekUser.CursorLocation = 3
rsCekUser.Open qCekUser, connstr
if rsCekUser.eof then
   txtCekUser="False"
else
   txtCekUser="True"
end if
rsCekUser.close
set rsCekUser = Nothing

if cmdView = "True" and txtCekUser="True" Then 
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q="select module.code as ModuleCode, module.NAME as ModuleName, sub.code as SubModuleCode, sub.NAME as SubModuleName" & _
      " FROM module module, submodule sub " & _
      " WHERE ((module.code = sub.module)) order by module.code,sub.code"
   'response.Write q
   'RESPONSE.Flush()
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="AccessRights.Asp" method="post" name="eForm">
   <table class="entry" cellpadding="1" cellspacing="1" width="620">
	<tr bgcolor="#EEEEEE">
	<td>&nbsp;</td>
	<td align="center" title="Module Code" valign="top" colspan="2">Module</td>
	<td align="center" title="Sub Module Name" valign="top" colspan="2">Sub Module</td>
	<td>&nbsp;</td>
	</tr>
	<tr bgcolor="#EEEEEE">
	<td align="center" valign="top" width="20">No.</td>
	<td align="center" title="Module Code" valign="top" width="50">Code</td>
	<td align="center" title="Module Name" valign="top" width="200">Name</td>
	<td align="center" title="Sub Module Code" valign="top" width="50">Code</td>
	<td align="center" title="Sub Module Name" valign="top" width="200">Name</td>
	<td align="center" title="Can Accessor No" valign="top" width="100">Active</td>
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
		  end if
		  v_ModuleCode=rsVH("ModuleCode")
		  v_SubModuleCode=rsVH("SubModuleCode")
		  v_name = rsVH("ModuleCode")&rsVH("SubModuleCode")
		  %>
		  <tr bgcolor="<%=cColor%>">
		  <td align="Right" valign="top"><%=NO%></td>
	      <td align="Center" title="Module Code" valign="top"><%=rsVH("ModuleCode")%></td>
	      <td align="Left" title="Module Name" valign="top"><%=rsVH("ModuleName")%></td>
	      <td align="Center" title="Sub Module Code" valign="top"><%=rsVH("SubModuleCode")%></td>
	      <td align="Left" title="Sub Module Name" valign="top"><%=rsVH("SubModuleName")%></td>
		  <%
		    Dim rsCek,txtq,v_ModuleCode,v_SubModuleCode
			Set rsCek = server.CreateObject("Adodb.recordset")
			txtq="select AUTHORIZED from ACCESSRIGHTS where FCCODE='" & v_fcCode  & "' and MODULECODE='" & v_ModuleCode & "' and SUBMODULECODE='" & v_SubModuleCode & "'  and fcBA='" & Session("fcBA") & "'"
			'response.Write q
			'RESPONSE.Flush()
			Set rsCek = Server.CreateObject("ADODB.Recordset")
			rsCek.CursorType = 3
			rsCek.CursorLocation = 3
			rsCek.Open txtq, connstr
		  %>
		  <%if rsCek.Eof Then%>
	           <td align="left" title="Can Accessor No" valign="top">Active Or No<input name="<%=v_name%>" type="checkbox"></td>
		  <%Else%>
	           <td align="left" title="Can Accessor No" valign="top">Active Or No<input name="<%=v_name%>" type="checkbox" checked></td>
		  <%End if%>
		  <%
		  rsCek.Close
		  set rsCek = Nothing
		  %>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
	<tr bgcolor="#EEEEEE">
	<td colspan="6" align="center">
	  <input type="hidden" name="v_FCCODE" value="<%=v_FCCODE%>">
	  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
	  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
	  <%IF Request("txtJenis")="Edit" then%>
		   <input type="hidden" name="txtJenis" value="Edit">
		   <input Type="Submit"  name="Add" value="Update">
	  <%Else%>
		   <input type="hidden" name="txtJenis" value="New">
		   <input Type="Submit"  name="Add" value="Save">
	  <%End if%>
	  <input Type="Reset"  name="Reset" value="Reset">
	  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
  </tr>

   </TABLE>
   </div>
   </form>
<%elseif cmdView = "True" and txtCekUser="False" Then %>
   <P><font face="Verdana" size="1" color="#FF0000">User Name Not Found.. Please entry from Data Setup</P>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
