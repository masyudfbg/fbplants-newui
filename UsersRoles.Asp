<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include virtual="/adovbs.inc"-->
<html>
	<head>
		<title>ePlaS - Combination Users and Roles</title>
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
	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
	
	<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %>>
		<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
		<% if Session("fcBa") <> "" then %>
         <!--#include file="menu.asp"-->
        <% end if %>
	</div>
    <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
    <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
    <div class="box_both" style="overflow:scroll;"> 
				<!---------------awal----------------------->
<%


Dim FCCODE, FCROLESCODE, FCUSERCODE, VALID_FROM, VALID_UNTIL, ISACTIVE, FCENTRY, FCEDIT, FCIP
Dim cmdView, qInsert, qUpdate, qDelete

FCCODE		= Request("FCCODE")
FCROLESCODE	= Request("FCROLESCODE")
FCUSERCODE	= Request("FCUSERCODE")
VALID_FROM	= Request("VALID_FROM")
VALID_UNTIL	= Request("VALID_UNTIL")
ISACTIVE		= Request("ISACTIVE")
FCENTRY		= Session("CodeName")
FCEDIT 		= Session("CodeName")
FCIP 			= Request.ServerVariables("REMOTE_ADDR")


if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
	    qDelete = 	" DELETE FROM USER_ROLES_TAB " & _
					" WHERE " & _
					" FCUSERCODE = '" & FCUSERCODE & "' AND " & _
					" FCROLESCODE = '" & FCROLESCODE & "' AND " & _
					" VALID_FROM = TO_DATE('" & VALID_FROM & "', 'MM/DD/YYYY')"
		
		'Response.Write qDelete
		connStr.execute(qDelete)
		cmdView	= "True"
		   
	elseif Request("txtJenis")="New" then
	    qInsert = 	" INSERT INTO USER_ROLES_TAB (" & _
					" FCROLESCODE " & _
					" , FCUSERCODE " & _
					" , VALID_FROM " & _
					" , VALID_UNTIL " & _
					" , ISACTIVE " & _
					" , FCENTRY " & _
					" , FCEDIT " & _
					" , FCIP " & _
					" , LASTUPDATE " & _
					" , LASTTIME " & _
					" ) VALUES ( " & _
					" '" & FCROLESCODE & "'" & _
					" , '" & FCUSERCODE & "'" & _
					" , TO_DATE('" & VALID_FROM & "', 'MM/DD/YYYY')" & _
					" , TO_DATE('" & VALID_UNTIL & "', 'MM/DD/YYYY')" & _
					" , '" & ISACTIVE & "'" & _
					" , '" & FCENTRY & "'" & _
					" , '" & FCEDIT & "'" & _
					" , '" & FCIP & "'" & _
					" , sysdate, To_Char(sysdate,'HH24:MI')" & _
					" )"

		'Response.Write qInsert
		connStr.execute(qInsert)
		cmdView	= "True"
		   
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		qUpdate = 	" UPDATE USER_ROLES_TAB SET" & _
					" VALID_UNTIL =  TO_DATE('" & VALID_UNTIL & "', 'MM/DD/YYYY')" & _
					" , ISACTIVE = '" & ISACTIVE & "'" & _
					" , FCENTRY = '" & FCENTRY & "'" & _
					" , FCEDIT = '" & FCEDIT & "'" & _
					" , FCIP = '" & FCIP & "'" & _
					" , LASTUPDATE = sysdate " & _
				    " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
					" WHERE " & _
					" FCROLESCODE = '" & FCROLESCODE & "' AND" & _
					" FCUSERCODE = '" & FCUSERCODE & "' AND" & _
					" VALID_FROM = TO_DATE('" & VALID_FROM & "', 'MM/DD/YYYY')"
					
		' Response.Write qUpdate
		connStr.execute(qUpdate)
		cmdView	= "True"
	else
		cmdView	= "True"
	end if	
else
	cmdView	= Request("cmdView")
end if 
		   
Dim txtModule, txtTable, txtSearch, txtField, i, q, nCountofField, NO, txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = Request("txtJenis")


'-------------ISI Payment Form
%>
 <form name="frmEntry" method="post" action="UsersRoles.asp">
	<table class="data" cellpadding="0" cellspacing="0" width="75%">
		<tr>
			<td colspan="2">
				<div align="left">
					<font size="2" face="verdana"><strong>Maintenance Users and Roles</strong></font>
				</div>
			</td>
		</tr>
		<tr>
			<td align="left">UserName</td>
			<td align="left">
				<input type="text" maxlength="30" name="FCCODE" size="30" value="<%=FCCODE%>">
			</td>
		</tr>	
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=Request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=Request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">
			</td>
		</tr>
    </table>  
 </form>

<%
if cmdView = "True" Then
	Dim rsUsersRoles, intRecordsPerPage, cColor
	set rsUsersRoles = server.CreateObject("Adodb.recordset")
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	=	" SELECT FCUSERCODE," & _
			"        FCROLESCODE," & _
			"        TO_CHAR(VALID_FROM, 'MM/DD/YYYY') VALID_FROM," & _
			"        TO_CHAR(VALID_UNTIL, 'MM/DD/YYYY') VALID_UNTIL," & _
			"        ISACTIVE" & _
			"   FROM USER_ROLES_TAB"
			
	if len(trim(FCCODE)) > 0 then
	q	=	q &	" WHERE FCUSERCODE LIKE '%" & trim(FCCODE) & "%'"
	else
	q	=	q &	" "
	end if

	rsUsersRoles.CursorType = 3
	rsUsersRoles.CursorLocation = 3
	rsUsersRoles.Open q, connstr
%>
	<form action="UsersRoles.Asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="75%">
			<tr>
				<td align="left">User Name</td>
				<td align="left">
					<%if Request("txtJenis")="Edit" then%>
					<input type="text" size="30" name="FCUSERCODE" value="<%=FCUSERCODE%>" readonly="True" style="background:#999999">
					<%else%>
					<input type="text" size="30" name="FCUSERCODE" value="<%=FCUSERCODE%>">
					<%end if%>
				</td>
			</tr>
			
			<%if Request("txtJenis")="Edit" then%>
			<tr>
				<td align="left">Roles</td>
				<td align="left">
					<select name="FCROLESCODE" readonly style="background:#999999;color:#000;align:left;">
						<option value=""></option>
						<%
						Dim rsRoles, strRoles
						strRoles	=	" SELECT FCCODE, FCNAME FROM ROLES_TAB ORDER BY FCCODE"
						
						Set rsRoles = Server.CreateObject("ADODB.Recordset")
						rsRoles.Open strRoles, connstr, adOpenStatic
						
						if not rsRoles.eof then
						Do While Not rsRoles.eof
						%>
						<option value="<%=rsRoles("FCCODE")%>" <%if FCROLESCODE=rsRoles("FCCODE") then%> selected <%End if%>><%=rsRoles("FCCODE")%> : <%=rsRoles("FCNAME")%></option>
						<%
							rsRoles.movenext
						loop
						end if
						rsRoles.Close
						set rsRoles = Nothing
						%>
					</select>
				</td>
			</tr>
			<%else%>
			<tr>
				<td align="left">Roles</td>
				<td align="left">
					<select name="FCROLESCODE" >
						<option value=""></option>
						<%
						strRoles	=	" SELECT FCCODE, FCNAME FROM ROLES_TAB ORDER BY FCCODE"
						
						Set rsRoles = Server.CreateObject("ADODB.Recordset")
						rsRoles.Open strRoles, connstr, adOpenStatic
						
						if not rsRoles.eof then
						Do While Not rsRoles.eof
						%>
						<option value="<%=rsRoles("FCCODE")%>" <%if FCROLESCODE=rsRoles("FCCODE") then%> selected <%End if%>><%=rsRoles("FCCODE")%> : <%=rsRoles("FCNAME")%></option>
						<%
							rsRoles.movenext
						loop
						end if
						rsRoles.Close
						set rsRoles = Nothing
						%>
					</select>
				</td>
			</tr>
			<%end if%>
			
			<tr>
				<td align="left">Valid From</td>
				<td align="left">
					<%if Request("txtJenis")="Edit" then%>
					<input type="text" size="15" name="VALID_FROM" value="<%=VALID_FROM%>" readonly="True" style="background:#999999">
					<%else%>
					<input class="datepicker" type="text" size="15" name="VALID_FROM" value="<%=VALID_FROM%>">
					
					<%end if%>
				</td>
			</tr>

			<tr>
				<td align="left">Valid Until</td>
				<td align="left">
					<input class="datepicker" type="text" size="15" name="VALID_UNTIL" value="<%=VALID_UNTIL%>">
					
				</td>
			</tr>
			
			<tr>
				<td align="left">Is Active</td>
				<td align="left">
					<select name="ISACTIVE" title="Activation">
						<%if ISACTIVE="TRUE" Then%>
						<Option value=""></Option>
						<Option value="TRUE" selected>TRUE</Option>
						<Option value="FALSE">FALSE</Option>
						<%Elseif ISACTIVE="FALSE" Then%>
						<Option value=""></Option>
						<Option value="TRUE">TRUE</Option>
						<Option value="FALSE" selected>FALSE</Option>
						<%else%>
						<Option value="" selected></Option>
						<Option value="TRUE">TRUE</Option>
						<Option value="FALSE">FALSE</Option>
						<%End If%>
					</select>
				</td>
			</tr>
		
			<tr>
				<td align="left">
					<input type="hidden" name="FCCODE" value="<%=FCCODE%>">
					<input Type="Hidden" Value="<%=Request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=Request("txtModule")%>" Name="txtModule">
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

   <table class="data">
   <br/>
	<tr bgcolor="#0C5B24" style="color:#fff;" >
		<td align="left">NO</td>
		<td colspan="2">&nbsp;</td>
		<td align="left">USERNAME</td>	
		<td align="left">ROLES</td>	
		<td align="left">VALID FROM</td>	
		<td align="left">VALID UNTIL</td>	
		<td align="left">IS ACTIVE</td>	
	</tr>

   <%
   if Not rsUsersRoles.eof then
      nCountofField=rsUsersRoles.fields.count-1
	  no=1
	  While Not rsUsersRoles.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="UsersRoles.Asp?FCUSERCODE=<%=rsUsersRoles("FCUSERCODE")%>&FCROLESCODE=<%=rsUsersRoles("FCROLESCODE")%>&VALID_FROM=<%=rsUsersRoles("VALID_FROM")%>&VALID_UNTIL=<%=rsUsersRoles("VALID_UNTIL")%>&ISACTIVE=<%=rsUsersRoles("ISACTIVE")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=Request("FCUSERCODE")%>" border="0"></a></td><td>
			<a href="UsersRoles.Asp?FCUSERCODE=<%=rsUsersRoles("FCUSERCODE")%>&FCROLESCODE=<%=rsUsersRoles("FCROLESCODE")%>&VALID_FROM=<%=rsUsersRoles("VALID_FROM")%>&VALID_UNTIL=<%=rsUsersRoles("VALID_UNTIL")%>&ISACTIVE=<%=rsUsersRoles("ISACTIVE")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=Request("FCUSERCODE")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td align="left" title="<%=rsUsersRoles(i).Name%>" valign="top"><%=rsUsersRoles(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsUsersRoles.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsUsersRoles.Close
   set rsUsersRoles = Nothing%>
   </TABLE>
   </div>
   </form>
<%End if%>  
 
  <!-------------akhir---------------------->
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
