<%'Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="include/ba_session.asp"-->
<html>
<head>
        <title>ePlaS - Business Unit Setup for Default Access </title>
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
		<script language="JavaScript"> 
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
DIM v_FCCODE
Dim cmdView
Dim defaultCompany, qCompanyAll
Dim tempCounterDefault, doneContinue
Dim rsModule,qModule,qInsert, rsModuleCheck, QUpdate
		


v_FCCODE = Session("CodeName")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
	    q = " delete from user_fcba_tab " & _
			" where " & _
			" fcusercode = '" & v_FCCODE & "'"	

		   connStr.execute(q)
		   cmdView	= "True"
	elseif Request("txtJenis")="New" or Request("txtJenis")="Edit" then
	
		qCompanyAll = " select b.fccode, b.fcname, A.ISDEFAULT from  " & _
						" user_fcba_tab a, " & _
						" businessunit b " & _
						" where b.isactive='TRUE' and " & _
						" a.fcba = b.fccode and " & _
						" a.fcusercode = '" & v_FCCODE  & "'" 
			  
	    Set rsModuleCheck = Server.CreateObject("ADODB.Recordset")
		rsModuleCheck.CursorType = 3
		rsModuleCheck.CursorLocation = 3
		rsModuleCheck.Open qCompanyAll, connstr
		
		doneContinue = "TRUE"
		tempCounterDefault = 0
		if Not rsModuleCheck.eof then 
		    While Not rsModuleCheck.EOF
				if request(trim("data_default_" & rsModuleCheck("fccode")))="on" Then
					tempCounterDefault = tempCounterDefault + 1
				end if
				rsModuleCheck.MoveNext
		    wend
			if tempCounterDefault > 1 then 
			    Response.Write  "<font color=""#FF0000"" pointsize=""9"" face=""Arial, Helvetica, sans-serif"">Error : One User can only have One Default Business Unit, Please Check your data</font>"
				doneContinue = "FALSE"
			else
				if tempCounterDefault = 0 then 
					Response.Write  "<font color=""#FF0000"" pointsize=""9"" face=""Arial, Helvetica, sans-serif"">Error : One User, at least have one Default Business Unit, Please Check your data</font>"
					doneContinue = "FALSE"
				else
					doneContinue = "TRUE"
				end if
			end if
			
		end if	
		
		rsModuleCheck.Close
		set rsModuleCheck = Nothing
		
		if doneContinue = "TRUE" then 
		
			q = " update user_fcba_tab " & _
			    " set isdefault = 'FALSE' " & _
				" where " & _
				" fcusercode = '" & v_FCCODE & "'"	
			
			connStr.execute(q)		
			
			
			Set rsModule = Server.CreateObject("ADODB.Recordset")
			rsModule.CursorType = 3
			rsModule.CursorLocation = 3
			rsModule.Open qCompanyAll, connstr
		
			if Not rsModule.eof then
			
			   q = " update user_fcba_tab " & _
					" set isdefault = 'FALSE' " & _
					" where " & _
					" fcusercode = '" & v_FCCODE & "'"	
			
			   connStr.execute(q)
			   connStr.execute("commit")

			   While Not rsModule.EOF
					if request(trim("data_default_" & rsModule("fccode")))="on" Then
						defaultCompany = "TRUE"
						
						qQueryData = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where isactive='TRUE' and fcCode='" & rsModule("fccode") & "'"
						set rsQueryData = connStr.Execute(qQueryData)
						if not rsQueryData.eof then
							' session BA (fcBa, BA, PT, periode tiap modul, dst.) - include/ba_session.asp
							FbpSetSessionBA rsQueryData
						
							if request("persistent") = "y" Then
								  response.Cookies("myCookies")("txtBA") = rsModule("fccode")
							End if
							q="Insert into HISTORICAL_USER_TAB " & _
							      " ( USERNAME, IP_ADDRESS, BU, DATE_ACT, TIME_ACT, LASTUPDATE, LASTTIME ) " & _
									"values ('" & Session("CodeName") & "','" & _
									Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','"& FormatDateTime(Date, 1) &"','"& FormatDateTime(Date, 3) &"', sysdate, to_char(sysdate, 'HH24:MI'))"
							connStr.execute(q)
							
							qUpdateCompany = " update user_company_tab " & _
											 " set " & _
											 " isdefault = 'FALSE'" & _
											 " where " & _
											 " fcusercode = '" & v_FCCODE & "'" 
														
							connStr.execute(qUpdateCompany) 
							
							qUpdateCompany = " update user_company_tab " & _
											 " set " & _
											 " isdefault = 'TRUE'" & _
											 " where " & _
											 " fcusercode = '" & v_FCCODE & "' and " & _
											 " fccompanycode = '" & Session("COMPANYCODE") & "'"
							connStr.execute(qUpdateCompany) 
						
						end if
					else 
						defaultCompany = "FALSE"
					end if
				  
					qUpdate = " update user_fcba_tab " & _
						" set isdefault = '" & defaultCompany & "' " & _
						" where " & _
						" fcba = '" & rsModule("fccode") & "' and " & _
						" fcusercode = '" & v_FCCODE & "'"
													
					connStr.execute(qUpdate) 
					
					
					connStr.execute("commit") 
					
				    rsModule.MoveNext
			   wend
			end if
			response.Write "<font face=Verdana size=1 color=blue>Accessright Completed..</font>"
			rsModule.Close
			set rsModule = Nothing
		else
		    cmdView = "True"
		end if
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
 
 <form name="frmEntry" method="post" action="UsersChangedBusinessUnit.asp">
	<table class="data" width="500">
		<tr><td align="left">User Code</td>
		<td align="left"><input readonly="True" type="text" maxlength="25" name="v_FCCODE" size="25" value="<%=v_FCCODE%>"></td></tr>	
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
dim qCekUsers,rsCekUsers,txtCekUsers
qCekUsers= " select fccode from users_tab " & _
			" where  " & _
			" fccode = '" & v_FCCODE & "' and " & _
			" accessrights = 'Y' and  " & _
			" isactive = 'TRUE' and  " & _
			" sysdate between valid_from and valid_until  " 

Set rsCekUsers = Server.CreateObject("ADODB.Recordset")
rsCekUsers.CursorType = 3
rsCekUsers.CursorLocation = 3
rsCekUsers.Open qCekUsers, connstr
if rsCekUsers.eof then
   txtCekUsers="False"
else
   txtCekUsers="True"
end if
rsCekUsers.close
set rsCekUsers = Nothing

if cmdView = "True" and txtCekUsers="True" Then 
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
  
	  
	q = " select b.FCCOMPANYCODE, b.fccode, b.fcname, A.ISDEFAULT from  " & _
						" user_fcba_tab a, " & _
						" businessunit b " & _
						" where b.isactive='TRUE' and " & _
						" a.fcba = b.fccode and " & _
						" a.fcusercode = '" & v_FCCODE  & "'  order by b.FCCOMPANYCODE,b.fccode" 	  
						
	  
	  
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="UsersChangedBusinessUnit.Asp" method="post" name="eForm">
   <table class="entry" cellpadding="1" cellspacing="1" width="620">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
	<td align="center" valign="top" width="20">No.</td>
	<td align="center" title="Company" valign="top" width="50">Company</td>
	<td align="center" title="Business Unit Code" valign="top" width="50">Business Unit Code</td>
	<td align="center" title="Name" valign="top" width="200">Business Unit Name</td>
	<td readonly="True" align="center" title="Active" valign="top" width="100">Activated</td>
	<td align="center" title="Default" valign="top" width="100">Is Default</td>
	</tr>
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count
	  no=1
	  While Not rsVH.EOF
	      
	  
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if
		  v_CompanyCode=rsVH("fccode")
		  v_CompanyName=rsVH("fcname")
		  v_IsDefault = rsVH("ISDEFAULT")


		  %>
		  <tr bgcolor="<%=cColor%>">
			  <td align="Right" valign="top"><%=NO%></td>
			  <td align="Left" title="Business Unit Code" valign="top"><%=rsVH("FCCOMPANYCODE")%></td>
			  <td align="Left" title="Business Unit Code" valign="top"><%=rsVH("fccode")%></td>
			  <td align="Left" title="Business Unit Name" valign="top"><%=rsVH("fcname")%></td>
			  <td align="Left" title="Active" valign="top">Active</td>
			  
				  

			  <% if v_IsDefault = "TRUE" Then %>
				   <td align="left" title="Default Or No" valign="top">Default Or No 
						<input name="cb_default_<%=v_CompanyCode%>" id="cb_default_<%=v_CompanyCode%>" type="checkbox" value="on" onClick="SetCheckBoxDefault('<%=v_CompanyCode%>','cb_default_<%=v_CompanyCode%>')" checked>
						<div   id="div_default_<%=v_CompanyCode%>" name="div_default_<%=v_CompanyCode%>">
							 <input id="data_default_<%=v_CompanyCode%>" name="data_default_<%=v_CompanyCode%>" type="hidden" value="on" >
						</div>		
					</td>
			  <% else %>
				   <td align="left" title="Default Or No" valign="top">Default Or No 
						<input name="cb_default_<%=v_CompanyCode%>" id="cb_default_<%=v_CompanyCode%>" type="checkbox" value="off" onClick="SetCheckBoxDefault('<%=v_CompanyCode%>','cb_default_<%=v_CompanyCode%>')">
						<div   id="div_default_<%=v_CompanyCode%>" name="div_default_<%=v_CompanyCode%>">
							 <input id="data_default_<%=v_CompanyCode%>" name="data_default_<%=v_CompanyCode%>" type="hidden" value="off" >
						</div>		
					</td>
			  <%End if%>			  

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
<%elseif cmdView = "True" and txtCekRoles="False" Then %>
   <P><font face="Verdana" size="1" color="#FF0000">User Name Not Found.. Please entry from Data Setup</P>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
