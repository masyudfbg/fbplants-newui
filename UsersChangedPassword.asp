<%'Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/md5.asp"-->
<html>
<head>
<title>ePlaS - Change Password</title>
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


function SetCheckBox(CbCode, CbOriCode){    

    var StringAll;
	var value_;


	var CID = document.getElementById(CbOriCode).checked;
			
	if(document.getElementById(CbOriCode).checked == 1) {
		StringAll = "<input id='data_" + CbCode + "' name='data_" + CbCode + "' type='hidden' value='on'>";
	}
	else {
		StringAll = "<input id='data_" + CbCode + "' name='data_" + CbCode + "' type='hidden' value='off'>";
	}
	
    document.getElementById("div_" + CbCode ).innerHTML=StringAll;
}

function SetCheckBoxDefault(CbCode, CbOriCode){    

    var StringAll;
	var value_;


	var CID = document.getElementById(CbOriCode).checked;
			
	if(document.getElementById(CbOriCode).checked == 1) {
		StringAll = "<input id='data_default_" + CbCode + "' name='data_default_" + CbCode + "' type='hidden' value='on'>";
	}
	else {
		StringAll = "<input id='data_default_" + CbCode + "' name='data_default_" + CbCode + "' type='hidden' value='off'>";
	}
	
    document.getElementById("div_default_" + CbCode ).innerHTML=StringAll;
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
 <%'-------------awal-----------------------%>
 <%


Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA
DIM v_FCCODE,OLDPASSWORD,FCPASSWORD, RFCPASSWORD
Dim cmdView
Dim defaultCompany, qUserChgPass
Dim tempCounterDefault, doneContinue
Dim rsModule,qModule,qInsert, rsModuleCheck, QUpdate
		


v_FCCODE = Session("CodeName")
'Response.Write  Request("txtJenis") 
 if Request("txtJenis")="Edit" Then
   	dim qCekUsers,rsCekUsers,txtCekUsers
	qCekUsers= " Select 1 from users_tab where  FCPASSWORD = '" & md5(request("OLDPASSWORD")) & "' and fcCode='" & v_FCCODE & "'"	
	Set rsCekUsers = Server.CreateObject("ADODB.Recordset")
	rsCekUsers.CursorType = 3
	rsCekUsers.CursorLocation = 3
	rsCekUsers.Open qCekUsers, connstr
	if rsCekUsers.eof then
	   Response.Write  " Current Password No Match !!!"
	else
	
	  IF request("FCPASSWORD") = request("RFCPASSWORD")  THEN
		   qUserChgPass  = " update users_tab " & _
							" set " & _
							" FCPASSWORD = '" & md5(request("FCPASSWORD")) & "', " & _
							" FCIP  = '" & Request.ServerVariables("REMOTE_ADDR") &"'," & _
							" FCEDIT  = '" & Session("CodeName") & "'," & _
							" LASTUPDATE = sysdate ," & _
							" LASTTIME = To_Char(sysdate,'HH24:MI') " & _
							" where  FCPASSWORD = '" & md5(request("OLDPASSWORD")) & "' and fcCode='" & v_FCCODE & "'"	
				  
			Set rsModuleCheck = Server.CreateObject("ADODB.Recordset")
			rsModuleCheck.CursorType = 3
			rsModuleCheck.CursorLocation = 3
			rsModuleCheck.Open qUserChgPass, connstr
		   
			Response.Write  " Password Changed Success "
	  ELSE
            Response.Write  " New Password not Match then Re-type New Password "
      END IF	  
	   
	end if
	rsCekUsers.close
	set rsCekUsers = Nothing
	
    cmdView	= "True"							 

End if

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
 
 <form name="frmEntry" method="post" action="UsersChangedPassword.asp">
	<table class="data" width="500">
		<tr><td align="left">User Code</td>
		<td align="left"><input readonly="True" type="text" maxlength="25" name="v_FCCODE" size="25" value="<%=v_FCCODE%>"></td></tr>	
		<tr><td align="left">Current Password</td>
		<td align="left"><input type="text" maxlength="25" name="OLDPASSWORD" size="25" value="<%=OLDPASSWORD%>"></td></tr>
		<tr><td align="left">New Password</td>
		<td align="left"><input type="text" maxlength="25" name="FCPASSWORD" size="25" value="<%=FCPASSWORD%>"></td></tr>	
		<tr><td align="left">Re-type New Password</td>
		<td align="left"><input type="text" maxlength="25" name="RFCPASSWORD" size="25" value="<%=RFCPASSWORD%>"></td></tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input type="hidden" name="txtJenis" value="Edit">
		<input Type="Submit" name="Edit" value="Update">&nbsp;&nbsp;
		</td>
		</tr>
    </table>  
 </form>

 
  <%'-------------akhir----------------------%>
  <!--#include file="bottom.jsp"-->

</body>
</html>
