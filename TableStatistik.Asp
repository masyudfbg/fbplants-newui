<%
'Option Explicit
'On Error Resume Next
'response.buffer=True
validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/md5.asp"-->
<html>
<head>
<title>ePlaS - Electronic Plantation Solutions</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript">
// Perbaikan: x.js tidak dimuat di halaman ini -> dulu ReferenceError; header sticky kini dari FbpTable
if (window.xAddEventListener) xAddEventListener(window, 'load',
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
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <%
 '-------------awal-----------------------

Dim cmdView
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtJenis=Request("txtJenis")
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("Add")= "Update" Then
	if Request("txtJenis")= "Delete" then
		If txtTable ="BusinessUnit" OR txtTable ="Parameter" then
			q="Delete from " & txtTable & " where  fcCode='" & request("fcCode") & "'"
		ElseIf txtTable ="Pattern" then
			q="Delete from " & txtTable & " where  TABLENAME='" & request("fcCode") & "'"
		ElseIf txtTable ="Trigger_Control" then
			q="Delete from " & txtTable & " where  ACTIVITY='" & request("fcCode") & "'"
		ElseIf txtTable ="Users" then
			q="Delete from Users_Tab where  fccode='" & request("fcCode") & "'"
		ElseIf txtTable ="Jobmaster" then
			q="Delete from Jobmaster where  fccode='" & request("fcCode") & "'"
		ElseIf txtTable ="JobGroupMaster" then
			q="Delete from Jobgroupmaster where  fccode='" & request("fcCode") & "'" 	
		    q2= "Delete from Jobgroup where  fccode='" & request("fcCode") & "'" 
			connStr.execute(q2)
		ElseIf txtTable ="Jobbsmaster" then
			q="Delete from Jobbsmaster where  fccode='" & request("fcCode") & "'"
		ElseIf txtTable ="Jobismaster" then
			q="Delete from Jobismaster where  fccode='" & request("fcCode") & "'"
		ElseIf txtTable ="Jobcfmaster" then
			q="Delete from Jobcfmaster where  fccode='" & request("fcCode") & "'"
		ElseIf txtTable ="Jobbtmaster" then
			q="Delete from Jobbtmaster where  fccode='" & request("fcCode") & "'"
		else	
        	q="Delete from " & txtTable & " where  fcCode='" & request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"
		end if
        connStr.execute(q)
        cmdView	= "True"
	elseif Request("txtJenis")= "New" then%>
	   <!-------------------------------- jika insert baru ---------------------------------->	  

	  <%If txtTable ="BusinessUnit" then
	    'Save Default User Admin Password : SACPLANT
       q= "Insert into BusinessUnit (FCCODE, FCNAME, FCCOMPANYCODE, FCCOMPANYNAME, FCAREA, FCOWNERSHIP, FCMANAGER, FCTELEPHONE, FCFAX, FCEMAIL, FCKTU, FCTYPE, FCCROP, LASTUPDATE, LASTTIME,ISACTIVE, ADDRESS)" & _
		  "values ('" &  Request("FCCODE") & "','" & Request("FCNAME") & "','"  &  Request("FCCOMPANYCODE") & "'," & _
		  "'" & Request("FCCOMPANYNAME") & "','" & Request("FCAREA") & "','" & Request("FCOWNERSHIP") & "'," & _
		  "'" & Request("FCMANAGER") & "','" & Request("FCTELEPHONE") & "','" & Request("FCFAX") & "','" & Request("FCEMAIL") & "','" & Request("FCKTU") & "','" & Request("FCTYPE") & "','" & Request("FCCROP") & "', sysdate, To_Char(sysdate,'HH24:MI'),'" & Request("ISACTIVE") & "','" & Request("ADDRESS") & "' )" 
       connStr.execute(q)
			q= "Insert into UserProfile (FCCODE, FCPASSWORD, FULLNAME, DEPARTMENT, ACCESSRIGHTS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
			"values ('" &  "ADMIN" & "','" & md5("ADMIN"+Request("FCCODE")) & "','" & "ADMINISTRATOR"+Request("FCCODE") & "','" & "ADMINISTRATOR" & "','" & "Y" & "','" & _
			Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Request("FCCODE") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
			connStr.execute(q)
        q= "Insert into ACCESSRIGHTS (FCCODE, MODULECODE, SUBMODULECODE, AUTHORIZED, FCBA, FCENRTY, FCEDIT, FCIP, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  "ADMIN" & "','" & "1" & "','" & "2" & "','" & "Y" & "','" & Request("FCCODE") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
		  connStr.execute(q)
          cmdView	= "True"

	  ElseIf txtTable ="UserProfile" then
       q= "Insert into UserProfile " & _
		  "values ('" &  Request("FCCODE") & "','" & md5(Request("fcPassword")) & "','" & Request("fullname") & "','" & Request("department") & "','" & Request("ACCESSRIGHTS") & "','" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	      connStr.execute(q)
          cmdView	= "True"
		  
	  ElseIf txtTable ="Users" then
       q= "Insert into Users_Tab (FCCODE, FCPASSWORD, FULLNAME,TITLE, DEPARTMENT, ACCESSRIGHTS,FIRSTADDRESS, ISACTIVE, VALID_FROM, VALID_UNTIL, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Request("FCCODE") & "','" & md5(Request("fcPassword")) & "','" & Request("fullname") & "','" & Request("tilte") & "','" & Request("department") & "','" & Request("ACCESSRIGHTS") & "','" & Request("FIRSTADDRESS") & "','" & _
		  Request("ISACTIVE") & "', to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy') ,to_date('" & Request("VALID_UNTIL") & "','mm/dd/yyyy') ,'" & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI') )"
	      connStr.execute(q)
          cmdView	= "True"		  
	  ElseIf txtTable ="Pattern" then
	       q= "Insert into Pattern " & _
		  "values ('" &  Request("TABLENAME") & "','" & Request("DISPLAYPATTERN") & "','"  &  Request("EDITPATTERN") & "', sysdate, To_Char(sysdate,'HH24:MI') )" 
	      connStr.execute(q)
          cmdView	= "True"
	  ElseIf txtTable ="Parameter" then
	       q= "Insert into Parameter " & _
		  "values ('" &  Request("FCCODE") & "','" & Request("PARAMETERNAME") & "', sysdate, To_Char(sysdate,'HH24:MI') )" 
	      connStr.execute(q)
          cmdView	= "True"
	  ElseIf txtTable ="Trigger_Control" then
	       q= "Insert into Trigger_Control " & _
		  "values ('" &  Request("ACTIVITY") & "','" & Request("SOURCE") & "','" & Request("TARGET") & "', sysdate, To_Char(sysdate,'HH24:MI') )" 
	      connStr.execute(q)
          cmdView	= "True"
	  ElseIf txtTable ="Job" then
       q= "Insert into Job " & _
		  "SELECT '" &  Request("FCCODE") & "','" & Request("fcName") & "','" & Request("JobGroupCode") & "','" & Request("JOB_FORCROPTYPE") & "','" & Request("JOB_FORFIELDSTATUS") & "','"  & _
		  Request("JOB_CATEGORY") & "','" & Request("UNITOFMEASUREMENT") & "','" & Request("ACCOUNTCODE") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"',fcCode,'', sysdate, To_Char(sysdate,'HH24:MI'),'" & Request("JOB_FORBALANCESHEETCODE") & "','" & Request("JOB_FORBALANCESHEETNAME") & "','" & Request("JOB_FORPROFITANDLOSSCODE") & "','" & Request("JOB_FORPROFITANDLOSSNAME") & "','Y',,'" & Request("JOB_FORCASHFLOWCODE") & "','" & Request("JOB_FORCASHFLOWNAME") & "','" & Request("JOB_FORBUDGETCODE") & "','" & Request("JOB_FORBUDGETNAME") & "','" & Request("BUDGETQTY") & "' "  & _ 
		  " FROM businessunit WHERE isactive='TRUE' " 
		  connStr.execute(q)
		 ' Response.Write q
       q= " Insert into Jobmaster " & _
		  " SELECT '" &  Request("FCCODE") & "','" & Request("fcName") & "','" & Request("JobGroupCode") & "','" & Request("JOB_FORCROPTYPE") & "','" & Request("JOB_FORFIELDSTATUS") & "','"  & _
		  Request("JOB_CATEGORY") & "','" & Request("UNITOFMEASUREMENT") & "','" & Request("ACCOUNTCODE") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','', sysdate, To_Char(sysdate,'HH24:MI'),'" & Request("JOB_FORBALANCESHEETCODE") & "','" & Request("JOB_FORBALANCESHEETNAME") & "','" & Request("JOB_FORPROFITANDLOSSCODE") & "','" & Request("JOB_FORPROFITANDLOSSNAME") & "','Y',,'" & Request("JOB_FORCASHFLOWCODE") & "','" & Request("JOB_FORCASHFLOWNAME") & "','" & Request("JOB_FORBUDGETCODE") & "','" & Request("JOB_FORBUDGETNAME") & "','" & Request("BUDGETQTY") & "' "  & _  
		  " FROM DUAL " 
		  connStr.execute(q)
          cmdView	= "True"		
      ElseIf txtTable ="Jobmaster" then
      q= " Insert into Jobmaster " & _
		  " SELECT '" &  Request("FCCODE") & "','" & Request("fcName") & "','" & Request("JobGroupCode") & "','" & Request("JOB_FORCROPTYPE") & "','" & Request("JOB_FORFIELDSTATUS") & "','"  & _
		  Request("JOB_CATEGORY") & "','" & Request("UNITOFMEASUREMENT") & "','" & Request("ACCOUNTCODE") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','', sysdate, To_Char(sysdate,'HH24:MI'),'" & Request("JOB_FORBALANCESHEETCODE") & "','" & Request("JOB_FORBALANCESHEETNAME") & "', " & _  
		  " '" & Request("JOB_FORPROFITANDLOSSCODE") & "','" & Request("JOB_FORPROFITANDLOSSNAME") & "','Y','" & Request("JOB_FORCASHFLOWCODE") & "','" & Request("JOB_FORCASHFLOWNAME") & "','" & Request("JOB_FORBUDGETCODE") & "','" & Request("JOB_FORBUDGETNAME") & "','" & Request("BUDGETQTY") & "' ,'" & Request("UNITOFMEASUREMENT") & "'"  & _ 
		  " FROM DUAL " 
		  connStr.execute(q)
	  q= "Insert into Job " & _
		  "SELECT '" &  Request("FCCODE") & "','" & Request("fcName") & "','" & Request("JobGroupCode") & "','" & Request("JOB_FORCROPTYPE") & "','" & Request("JOB_FORFIELDSTATUS") & "','"  & _
		  Request("JOB_CATEGORY") & "','" & Request("UNITOFMEASUREMENT") & "','" & Request("ACCOUNTCODE") & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"',fcCode,'', sysdate, To_Char(sysdate,'HH24:MI'),'" & Request("JOB_FORBALANCESHEETCODE") & "','" & Request("JOB_FORBALANCESHEETNAME") & "','" & Request("JOB_FORPROFITANDLOSSCODE") & "','" & Request("JOB_FORPROFITANDLOSSNAME") & "','Y','" & Request("JOB_FORCASHFLOWCODE") & "','" & Request("JOB_FORCASHFLOWNAME") & "','" & Request("JOB_FORBUDGETCODE") & "','" & Request("JOB_FORBUDGETNAME") & "','" & Request("BUDGETQTY") & "' ,'" & Request("UNITOFMEASUREMENT") & "'"  & _  
		  " FROM businessunit WHERE isactive='TRUE' " 
		  connStr.execute(q)
          cmdView	= "True"
	  ElseIf txtTable ="JobGroup" then
       q= "Insert into JobGroup " & _
		  "values ('" &  Request("FCCODE") & "','" & Request("fcName") & "','" & Request("PARENTCODE") & "','" & Request("PARENTDESC") & "','"  & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','', sysdate, To_Char(sysdate,'HH24:MI') )"
	      connStr.execute(q)
          cmdView	= "True"
	  ElseIf txtTable ="JobGroupMaster" then
       q= "Insert into JobGroupMaster(FCCODE,FCNAME,PARENTCODE,PARENTDESC,FCENTRY,FCEDIT,FCIP,ACTIVITY_HEADING,LASTUPDATE,LASTTIME,FCNAME_ENG,PARENTDESC_ENG) " & _
		  "values ('" &  Request("FCCODE") & "','" & Request("fcName") & "','" & Request("PARENTCODE") & "','" & Request("PARENTDESC") & "','"  & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','', sysdate, To_Char(sysdate,'HH24:MI') ,'" & Request("fcName") & "','"  & Request("PARENTDESC_ENG") & "' )"
		'  Response.Write q
	      connStr.execute(q)
		q= "Insert into JobGroup " & _
		  "SELECT '" &  Request("FCCODE") & "','" & Request("fcName") & "','" & Request("PARENTCODE") & "','" & Request("PARENTDESC") & "','"  & _
		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"',fccode,'', sysdate, To_Char(sysdate,'HH24:MI') ,'" & Request("fcName") & "','"  & Request("PARENTDESC_ENG") & "' " & _ 
		  " FROM businessunit WHERE isactive='TRUE' " 
		  connStr.execute(q)
	      cmdView	= "True"
	  
	  ElseIf txtTable ="GCMaster" then
	       q= "Insert into GCMaster (ITEM, FCCODE, FCNAME, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME)" & _
		  "values ('" &  Request("ITEM") & "','" & Request("FCCODE") & "','"  &  Request("FCNAME") & "','" & _
  		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
	      connStr.execute(q)
          cmdView	= "True"
	 ElseIf txtTable ="Jobbsmaster" then
	       q= "Insert into Jobbsmaster (FCCODE, FCNAME, FCNAME_ENG, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME)" & _
		  "values ('" & Request("FCCODE") & "','"  &  Request("FCNAME") & "','" &  Request("FCNAME_ENG") & "','" & _
  		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI') )"
	      connStr.execute(q)
          cmdView	= "True"
		
	ElseIf txtTable ="Jobismaster" then
	       q= "Insert into Jobismaster (FCCODE, FCNAME, FCNAME_ENG,  FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME)" & _
		  "values ('" & Request("FCCODE") & "','"  &  Request("FCNAME") & "','" &  Request("FCNAME_ENG") & "','" & _
  		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI') )"
	      connStr.execute(q)
          cmdView	= "True"
	ElseIf txtTable ="Jobcfmaster" then
	       q= "Insert into Jobcfmaster (FCCODE, FCNAME, FCNAME_ENG, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME)" & _
		  "values ('" & Request("FCCODE") & "','"  &  Request("FCNAME") & "','" &  Request("FCNAME_ENG") & "','" & _
  		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"',sysdate, To_Char(sysdate,'HH24:MI') )"
	      connStr.execute(q)
          cmdView	= "True"
	ElseIf txtTable ="Jobbtmaster" then
	       q= "Insert into Jobbtmaster (FCCODE, FCNAME, FCNAME_ENG, FCENTRY, FCEDIT, FCIP, LASTUPDATE, LASTTIME)" & _
		  "values ('"  & Request("FCCODE") & "','"  &  Request("FCNAME") & "','" &  Request("FCNAME_ENG") & "','" & _
  		  Session("CodeName") & "','" & Session("CodeName") & "','" & _
		  Request.ServerVariables("REMOTE_ADDR") &"', sysdate, To_Char(sysdate,'HH24:MI') )"
	      connStr.execute(q)
          cmdView	= "True"		  
	  end if%>
	  
	  
    <!-------------------------------- jika edit ---------------------------------->
		  
	<%elseif Request("Add") ="Update"  then
    	If txtTable ="BusinessUnit" then
			q="Update BusinessUnit set FCNAME ='" & Request("FCNAME") & "'" & _
				",FCCOMPANYCODE='" & Request("FCCOMPANYCODE") & "',FCCOMPANYNAME='" & Request("FCCOMPANYNAME") & "'" & _
				",FCAREA='" & REQUEST("FCAREA") & "',FCOWNERSHIP='" & REQUEST("FCOWNERSHIP") & "',FCFAX='" & REQUEST("FCFAX") & "',FCMANAGER='" & REQUEST("FCMANAGER") & "'" & _
				",fcType='" & REQUEST("fcType") & "',fcCrop='" & REQUEST("fcCrop") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				",FCTELEPHONE='" & REQUEST("FCTELEPHONE") & "',FCEMAIL='" & REQUEST("FCEMAIL") & "',FCKTU ='" & REQUEST("FCKTU") & "',ISACTIVE = '" & REQUEST("ISACTIVE") & "' ,ADDRESS = '" & REQUEST("ADDRESS") & "'" & _
				" where  fcCode='" & Request("fcCode") & "'"								
     	ElseIf txtTable ="UserProfile" then 
			q="Update UserProfile set FCPASSWORD ='" & md5(Request("FCPASSWORD")) & "'" & _
				",FULLNAME='" & Request("FULLNAME") & "',DEPARTMENT='" & Request("DEPARTMENT") & "'" & _
				",ACCESSRIGHTS='" & REQUEST("ACCESSRIGHTS") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fcCode='" & Request("fcCode") & "' and fcBA='" & Session("fcBA") & "'"								
     	ElseIf txtTable ="Users" then 
						If len(trim(request("FCPASSWORD")))	 > 1 then
							q= " update users_tab " & _
								" set " & _
									" FCPASSWORD = '" & md5(request("FCPASSWORD")) & "', " & _
									" FULLNAME  = '" & request("FULLNAME") & "', " & _
									" TITLE  = '" & request("TITLE") & "', " & _
									" DEPARTMENT  = '" & request("DEPARTMENT") & "', " & _
									" ACCESSRIGHTS  = '" & request("ACCESSRIGHTS") & "', " & _
									" ISACTIVE  = '" & request("ISACTIVE") & "', " & _
									" FIRSTADDRESS = '" & request("FIRSTADDRESS") & "', " & _
									" VALID_FROM  = to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy') ," & _
									" VALID_UNTIL  = to_date('" & Request("VALID_UNTIL") & "','mm/dd/yyyy') ," & _
									" FCIP  = '" & Request.ServerVariables("REMOTE_ADDR") &"'," & _
									" FCEDIT  = '" & Session("CodeName") & "'," & _
									"  LASTUPDATE = sysdate ," & _
									"   LASTTIME = To_Char(sysdate,'HH24:MI') " & _
									" where  fcCode='" & Request("fcCode") & "'"
					  else
					      q= " update users_tab " & _
								" set " & _
									" FULLNAME  = '" & request("FULLNAME") & "', " & _
									" TITLE  = '" & request("TITLE") & "', " & _
									" DEPARTMENT  = '" & request("DEPARTMENT") & "', " & _
									" ACCESSRIGHTS  = '" & request("ACCESSRIGHTS") & "', " & _
									" ISACTIVE  = '" & request("ISACTIVE") & "', " & _
									" FIRSTADDRESS = '" & request("FIRSTADDRESS") & "', " & _
									" VALID_FROM  = to_date('" & Request("VALID_FROM") & "','mm/dd/yyyy') ," & _
									" VALID_UNTIL  = to_date('" & Request("VALID_UNTIL") & "','mm/dd/yyyy') ," & _
									" FCIP  = '" & Request.ServerVariables("REMOTE_ADDR") &"'," & _
									" FCEDIT  = '" & Session("CodeName") & "'," & _
									"  LASTUPDATE = sysdate ," & _
									"   LASTTIME = To_Char(sysdate,'HH24:MI') " & _
									" where  fcCode='" & Request("fcCode") & "'"
                      end if					  
			ElseIf txtTable ="UsersChangedPassword" then 
		    q= " update users_tab " & _
				" set " & _
					" FCPASSWORD = '" & md5(request("FCPASSWORD")) & "', " & _
					" FCIP  = '" & Request.ServerVariables("REMOTE_ADDR") &"'," & _
					" FCEDIT  = '" & Session("CodeName") & "'," & _
					"  LASTUPDATE = sysdate ," & _
					"   LASTTIME = To_Char(sysdate,'HH24:MI') " & _
					" where  fcCode='" & Request("fcCode") & "'"		
     	ElseIf txtTable ="Trigger_Control" then 
			q="Update Trigger_Control set SOURCE='" & Request("SOURCE") & "',TARGET='" & Request("TARGET") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  ACTIVITY='" & Request("ACTIVITY") & "'"
    	ElseIf txtTable ="Pattern" then 
			q="Update Pattern set DISPLAYPATTERN ='" & Request("DISPLAYPATTERN") & "',EDITPATTERN ='" & Request("EDITPATTERN") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  TABLENAME='" & Request("fccode") & "'"								
    	ElseIf txtTable ="Parameter" then
			q="Update Parameter set PARAMETERNAME ='" & Request("PARAMETERNAME") & "'" & _
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
				" where  fccode='" & Request("fccode") & "'"								
	  ElseIf txtTable ="Job" then
		   q= "Update Job set " & _
			  "fcName='" & Request("fcName") & "',JobGroupCode='" & Request("JobGroupCode") & "',JOB_FORCROPTYPE='" & Request("JOB_FORCROPTYPE") & "',JOB_FORFIELDSTATUS='" & Request("JOB_FORFIELDSTATUS") & "'"  & _
			  " ,JOB_CATEGORY='" & Request("JOB_CATEGORY") & "',JOB_FORBALANCESHEETCODE='" & Request("JOB_FORBALANCESHEETCODE") & "'" & _
			  " ,JOB_FORBALANCESHEETNAME='" & Request("JOB_FORBALANCESHEETNAME") & "',JOB_FORPROFITANDLOSSCODE='" & Request("JOB_FORPROFITANDLOSSCODE") & "',JOB_FORPROFITANDLOSSNAME='" & Request("JOB_FORPROFITANDLOSSNAME") & "',UNITOFMEASUREMENT='" & Request("UNITOFMEASUREMENT") & "',ACCOUNTCODE='" & Request("ACCOUNTCODE") & "'" & _
			   " , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI'),BUDGETQTY='" & Request("BUDGETQTY") & "' " & _	
    			" where  fcCode='" & Request("fcCode") & "'"	
      ElseIf txtTable ="Jobmaster" then
		   q= "Update Jobmaster set " & _
			  "fcName='" & Request("fcName") & "',JobGroupCode='" & Request("JobGroupCode") & "',JOB_FORCROPTYPE='" & Request("JOB_FORCROPTYPE") & "',JOB_FORFIELDSTATUS='" & Request("JOB_FORFIELDSTATUS") & "'"  & _
			  " ,JOB_CATEGORY='" & Request("JOB_CATEGORY") & "',JOB_FORBALANCESHEETCODE='" & Request("JOB_FORBALANCESHEETCODE") & "'" & _
			  " ,JOB_FORBALANCESHEETNAME='" & Request("JOB_FORBALANCESHEETNAME") & "',JOB_FORPROFITANDLOSSCODE='" & Request("JOB_FORPROFITANDLOSSCODE") & "',JOB_FORPROFITANDLOSSNAME='" & Request("JOB_FORPROFITANDLOSSNAME") & "',UNITOFMEASUREMENT='" & Request("UNITOFMEASUREMENT") & "',ACCOUNTCODE='" & Request("ACCOUNTCODE") & "'" & _
			  " ,JOB_FORCASHFLOWCODE='" & Request("JOB_FORCASHFLOWCODE") & "',JOB_FORCASHFLOWNAME='" & Request("JOB_FORCASHFLOWNAME") & "'" & _
			  " ,JOB_FORBUDGETCODE='" & Request("JOB_FORBUDGETCODE") & "',JOB_FORBUDGETNAME='" & Request("JOB_FORBUDGETNAME") & "' " & _
			  " , LASTUPDATE = sysdate " & _
			  " , LASTTIME = To_Char(sysdate,'HH24:MI'),BUDGETQTY='" & Request("BUDGETQTY") & "' " & _	
    		  " where  fcCode='" & Request("fcCode") & "'"					
	  ElseIf txtTable ="JobGroup" then
		q=    "Update JobGroup set " & _
			  "fcName='" & Request("fcName") & "',PARENTCODE='" & Request("PARENTCODE") & "',PARENTDESC='" & Request("PARENTDESC") & "'"  & _
			  " , PARENTDESC_ENG='" & Request("PARENTDESC_ENG") & "'"  & _
			  " , LASTUPDATE = sysdate " & _
			  " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
			  " where  fcCode='" & Request("fcCode") & "'"	
	  ElseIf txtTable ="JobGroupMaster" then
		q=    "Update JobGroupMaster set " & _
			  "fcName='" & Request("fcName") & "',PARENTCODE='" & Request("PARENTCODE") & "',PARENTDESC='" & Request("PARENTDESC") & "'"  & _
              " , PARENTDESC_ENG='" & Request("PARENTDESC_ENG") & "'"  & _
			  " , LASTUPDATE = sysdate " & _
			  " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
			  " where  fcCode='" & Request("fcCode") & "'"	
			  
	  ElseIf txtTable ="GCMaster" then
		q=    "Update GCMaster set " & _
			  "FCNAME='" & Request("FCNAME") & "'"  & _
			  " , LASTUPDATE = sysdate " & _
			  " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
			  " where  ITEM='" & Request("ITEM") & "' and fcCode='" & Request("fcCode") & "'"			
       ElseIf txtTable ="Jobbsmaster" then
		q=    "Update Jobbsmaster set " & _
			  " FCNAME='" & Request("FCNAME") & "'"  & _
			  " , FCNAME_ENG='" & Request("FCNAME_ENG") & "'"  & _
			  " , LASTUPDATE = sysdate " & _
			  " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
			  " where fcCode='" & Request("fcCode") & "'"	
		ElseIf txtTable ="Jobismaster" then
		q=    "Update Jobismaster set " & _
			  " FCNAME='" & Request("FCNAME") & "'"  & _
			  " , FCNAME_ENG='" & Request("FCNAME_ENG") & "'"  & _
			  " , LASTUPDATE = sysdate " & _
			  " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
			  " where fcCode='" & Request("fcCode") & "'"	
		ElseIf txtTable ="Jobcfmaster" then
		q=    "Update Jobcfmaster set " & _
			  " FCNAME='" & Request("FCNAME") & "'"  & _
			  " , FCNAME_ENG='" & Request("FCNAME_ENG") & "'"  & _
			  " , LASTUPDATE = sysdate " & _
			  " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
			  " where fcCode='" & Request("fcCode") & "'"	
		ElseIf txtTable ="Jobbtmaster" then
		q=    "Update Jobbtmaster set " & _
			  " FCNAME='" & Request("FCNAME") & "'"  & _
			  " , FCNAME_ENG='" & Request("FCNAME_ENG") & "'"  & _
			  " , LASTUPDATE = sysdate " & _
			  " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
			  " where fcCode='" & Request("fcCode") & "'"	
			  
	  ' Special for Administrator, Added by NAF, 2-May-2013
	  ElseIf txtTable ="Roles" then
       q= "Update Roles_Tab set " & _
	      "FCNAME='" & Request("FCNAME") & "'"  & _
		  " , LASTUPDATE = sysdate " & _
		  " , LASTTIME = To_Char(sysdate,'HH24:MI') " & _	
		  " where  fcCode='" & Request("fcCode") & "'"	
	  
 	   end if
	  ' response.write q
	   connStr.execute(q)
       cmdView	= "True"
	end if
	'Response.Write "Sql is " & q	
	'Response.Flush()							 
else
    cmdView	= request("cmdView")
End if
Dim txtModule,txtTable,txtSearch,txtField,Ni,q,nCountofField,NO,fddate,txtJenis,fcCode
fddate   =Request("fddate")
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
fcCode=   Request("fcCode")
txtJenis =Request("txtJenis")

if err.number <> 0 then
   response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & err.description & "</font>"
   cmdView	= "True"
End if

Dim rsSetup 
set rsSetup = server.CreateObject("Adodb.recordset")
if Request("txtJenis")="Edit" then
	if txtTable="Pattern" Then
	   q = "select * from " & txtTable & " where  TABLENAME='" & request("fcCode") & "'"
	ELSEif txtTable="Trigger_Control" Then
	   q = "select * from " & txtTable & " where  ACTIVITY='" & request("fcCode") & "'" 
	ELSEif txtTable="Users" Then
	   q = "select FCCODE, FCPASSWORD, FULLNAME,TITLE, DEPARTMENT, ACCESSRIGHTS, ISACTIVE,FIRSTADDRESS, VALID_FROM, VALID_UNTIL from Users_Tab where  fccode='" & request("fcCode") & "'"
	else
	   q = "select * from " & txtTable & " where  fcCode='" & request("fcCode") & "'"
	end if
ELSE
	if txtTable = "Users" then 
	    q = "select FCCODE, FCPASSWORD, FULLNAME,TITLE, DEPARTMENT, ACCESSRIGHTS, ISACTIVE,FIRSTADDRESS, VALID_FROM, VALID_UNTIL from Users_Tab "
    else
    	q = "select * from " & txtTable
	end if
END IF
Set rsSetup = Server.CreateObject("ADODB.Recordset")
rsSetup.CursorType = 3
rsSetup.CursorLocation = 3
'response.write q
rsSetup.Open q, connstr%>
 <%'-------------isi Form
 %>
<%if Request("txttable") ="BusinessUnit" then%>
	<script>
	function validasi(form){

	  if (form.FCCOMPANYCODE.value == ""){

		alert("Anda belum mengisikan Company Code.");
		
		form.FCCOMPANYCODE.focus();

		return (false);

	  }
	if (form.FCCOMPANYNAME.value == ""){

		alert("Anda belum mengisikan Company Name.");

		form.FCCOMPANYNAME.focus();

		return (false);

	  }
	if (form.FCCODE.value == ""){

		alert("Anda belum mengisikan Business Unit Code.");

		form.FCCODE.focus();

		return (false);

	  }  
	if (form.FCNAME.value == ""){

		alert("Anda belum mengisikan Business Unit Name.");

		form.FCNAME.focus();

		return (false);

	  }  
	if (form.FCAREA.value == ""){

		alert("Anda belum mengisikan Area.");

		form.FCAREA.focus();

		return (false);

	  }  
	if (form.FCOWNERSHIP.value == ""){

		alert("Anda belum mengisikan Ownership.");

		form.FCOWNERSHIP.focus();

		return (false);
	}
	if (form.fcType.value == ""){

		alert("Anda belum mengisikan Type.");

		form.fcType.focus();

		return (false);

	  }    
	 if (form.fcCrop.value == ""){

		alert("Anda belum mengisikan Crop.");

		form.fcCrop.focus();

		return (false);

	  }    
	 }
	</script>
<form name="frmEntry" method="post" action="TableStatistik.asp" onSubmit="javascript:return validasi(this)">	
<%else%>
 <form name="frmEntry" method="post" action="TableStatistik.asp">
<%end if%>
   <% IF txtTable <> "HistLog"   THEN %>
	    <table class="data" >
        <%'if not rsSetup.eof Then
		    nCountofField=rsSetup.fields.count-1
		  'End if
         Dim txtBoxColor,txtTitle%>
		  
	<!----------------------Atur Tampilan Form -->

	<!--#include file="Select_Form_Show_Setup.asp"-->
	
	<% If txtTable <> "PeriodControl"  or txtTable <> "HistLog"   then%>
          <tr> 
            <td colspan="2"  align="left"> <input Type="Hidden" Value="True" Name="cmdView"> 
              <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable"> 
              <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule"> 
              <%IF Request("txtJenis")="Edit" then%> <input type="hidden" name="txtJenis" value="Edit"> 
              <input Type="Submit"  name="Add" value="Update"> 
              <%Else%> <input type="hidden" name="txtJenis" value="New"> 
              <input Type="Submit"  name="Add" value="Add"> 
              <%End if%> <input Type="Reset"  name="Reset" value="Cancel">
		<%IF txtModule="Data Master" or txtModule="Data Setup" Then%>
			<a href="PrintData.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fdDate%>&fcCode=<%=fcCode%>&txtSearch=<%=request("fcCode")%>&txtField=<%="fcCode"%>" 
			title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a> * Mandatory Field 
		<%Else%>
			<a href="PrintData.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fdDate%>&fcCode=<%=fcCode%>" 
			title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a> * Mandatory Field 
		<%End If%>
 
            </td>
          </tr>
	 <%End IF%>
        </table>  
	<% END IF %>	
 </form>
 
 <% IF txtTable <> "HistLog"   THEN %>
	<!--#include file="validation_show.Asp"--> 
<% END IF %>
	

<%'if cmdView = "True" Then
	Dim rsGAD
	set rsGAD = server.CreateObject("Adodb.recordset")
	if Request("txttable")="Pattern" or Request("txttable")="Parameter" then
		q	=	" select * from " & txtTable & ""
	ElseIf Request("txttable")="Trigger_Control" Then
		q	=	" SELECT ACTIVITY, JOB_API.get_jobcode_name(ACTIVITY,'" & session("fcBA") & "') as jobname,SOURCE, TARGET FROM " & txtTable & " ORDER BY ACTIVITY, SOURCE, TARGET"
	elseif Request("txttable")="BusinessUnit" Then
		q	=	" select" & _
				" FCCOMPANYCODE" & _
				" , FCCOMPANYNAME" & _
				" , FCCODE" & _
				" , FCNAME" & _
				" , FCAREA" & _
				" , FCOWNERSHIP" & _
				" , FCTYPE" & _
				" , FCCROP" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCTELEPHONE" & _
				" , FCFAX" & _
				" , FCEMAIL" & _
				" , ISACTIVE" & _
				" from " & txtTable & " ORDER BY FCCOMPANYCODE,ORDERDATA"
   elseif Request("txttable")="Users" Then
		q	=	" SELECT" & _
				" FCCODE, FCPASSWORD, FULLNAME,TITLE, DEPARTMENT, ACCESSRIGHTS, ISACTIVE,FIRSTADDRESS, VALID_FROM, VALID_UNTIL" & _
				" FROM USERS_TAB ORDER BY FCCODE"
     elseif Request("txttable")="UsersChangedPassword" Then
		q	=	" SELECT" & _
				" FCCODE, FCPASSWORD " & _
				" FROM USERS_TAB WHERE FCCODE ='" & Session("CodeName") & "' "				
   elseif Request("txttable")="UserProfile" Then
		q	=	" SELECT" & _
				" FCCODE, FCPASSWORD, FULLNAME,TITLE, DEPARTMENT, ACCESSRIGHTS, ISACTIVE, VALID_FROM, VALID_UNTIL, FCENTRY" & _
				" FROM USERS_TAB ORDER BY FCCODE"
   elseif Request("txttable")="PeriodControl" Then
		q	=	" select * from " & txtTable & " where fcba='" & session("fcBA") & "' order by POSTINGPOSITION"	
   elseif Request("txttable")="HistLog" Then
		q	=	" SELECT  *  FROM HISTORICAL_USER_TAB where username ='" & Session("CodeName") & "' and  bu='" & session("fcBA") & "'"
   elseif Request("txttable") = "Job" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY, UOM_UNIT, UNITOFMEASUREMENT, ACCOUNTCODE" & _
				" , JOB_FORBALANCESHEETCODE, JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSCODE, JOB_FORPROFITANDLOSSNAME" & _
				" , FCEDIT, FCIP, FCBA, JOB_FORACCOUNTTYPE" & _
				" FROM JOB WHERE FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE"
				' response.write q
    elseif Request("txttable") = "Jobmaster" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, JOBGROUPCODE, JOB_FORCROPTYPE, JOB_FORFIELDSTATUS, JOB_CATEGORY,  UOM_UNIT, UNITOFMEASUREMENT, ACCOUNTCODE" & _
				" , JOB_FORBALANCESHEETCODE, JOB_FORBALANCESHEETNAME, JOB_FORPROFITANDLOSSCODE, JOB_FORPROFITANDLOSSNAME" & _
				" , JOB_FORCASHFLOWCODE, JOB_FORCASHFLOWNAME, JOB_FORBUDGETCODE, JOB_FORBUDGETNAME,BUDGETQTY " & _				
				" , FCEDIT, FCIP, 'XXX' FCBA,JOB_FORACCOUNTTYPE" & _
				" FROM JOBMASTER ORDER BY FCCODE"
				' response.write q
   elseif Request("txttable") = "JobBahan" Then
		q	=	" SELECT" & _
				" a.FCCODE, a.FCNAME, ITEMCODE , b.FCNAME " & _
				" FROM JOBBAHAN a INNER JOIN STOCK b ON (a.ITEMCODE=b.FCCODE AND a.FCBA=b.FCBA) WHERE a.FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE"
				' response.write q
	elseif Request("txttable") = "JobGroup" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, PARENTCODE, PARENTDESC, PARENTDESC_ENG, FCENTRY, FCEDIT, FCIP, FCBA" & _
				" FROM JOBGROUP WHERE FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE"
	elseif Request("txttable") = "JobGroupMaster" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, PARENTCODE, PARENTDESC, PARENTDESC_ENG, FCENTRY, FCEDIT, FCIP,'XXX' FCBA" & _
				" FROM JOBGROUPMASTER  ORDER BY FCCODE"
    elseif Request("txttable") = "Jobbsmaster" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, FCNAME_ENG, FCENTRY, FCEDIT, FCIP,'XXX' FCBA" & _
				" FROM JOBBSMASTER  ORDER BY FCCODE"
	elseif Request("txttable") = "Jobismaster" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, FCNAME_ENG, FCENTRY, FCEDIT, FCIP,'XXX' FCBA" & _
				" FROM JOBISMASTER  ORDER BY FCCODE"	
	elseif Request("txttable") = "Jobcfmaster" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, FCNAME_ENG, FCENTRY, FCEDIT, FCIP,'XXX' FCBA" & _
				" FROM JOBCFMASTER  ORDER BY FCCODE"
	elseif Request("txttable") = "Jobbtmaster" Then
		q	=	" SELECT" & _
				" FCCODE, FCNAME, FCNAME_ENG, FCENTRY, FCEDIT, FCIP,'XXX' FCBA" & _
				" FROM JOBBTMASTER  ORDER BY FCCODE"				
   else
		q	=	" SELECT * FROM " & txtTable & " WHERE FCBA = '" & session("fcBA") & "'"
   end if
   'response.Write "xx:" & q
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr
   JumRec=1
   'intRecordsPerPage = 15
	
	Dim intRecordsPerPage,intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGad.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 45
	
	rsGad.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  rsGad.PageCount
	intRecords = rsGad.RecordCount
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

   rsGad.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount  
   nCountofField = 0   
   %>
   <table class="data" >
   <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-1
	  no=1
	%>
	<%If txtTable ="BusinessUnit" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">COMPANY CODE</Td>
		<td align="left" valign="top">COMPANY NAME</td>
		<Td align="left" valign="top">CODE</Td>
		<td align="left" valign="top">NAME</td>
		<Td align="left" valign="top">AREA</Td>
		<td align="left" valign="top">OWNERSHIP</td>
		<td align="left" valign="top">TYPE</td>
		<td align="left" valign="top">CROP</td>
		<Td align="left" valign="top">MANAGER</Td>
		<Td align="left" valign="top">KTU</Td>
		<td align="left" valign="top">TELEPON</td>
		<Td align="left" valign="top">FAX</Td>
		<td align="left" valign="top">E-MAIL</td>
		<td align="left" valign="top">ISACTIVE</td>
	</tr>
	<%ElseIf txtTable ="UserProfile" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td><td align="left" valign="top">PASSWORD</td>
		<Td align="left" valign="top">FULL NAME</Td><td align="left" valign="top">DEPARTEMENT</td>
		<Td align="left" valign="top">ACCESS RIGHT</Td>
	</tr>
	<%ElseIf txtTable ="Users" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">User Id</Td>
		<td align="left" valign="top">Password</td>
		<Td align="left" valign="top">Full Name</Td>
		<Td align="left" valign="top">Title</Td>
		<td align="left" valign="top">Department</td>
		<Td align="left" valign="top">Access Rights</Td>
		<Td align="left" valign="top">Active</Td>
		<Td align="left" valign="top">First Address</Td>
		<Td align="left" valign="top">Valid From</Td>
		<Td align="left" valign="top">Valid Until</Td>
	</tr>
	<%ElseIf txtTable ="UsersChangedPassword" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">User Id</Td>
		<td align="left" valign="top">Password</td>
	</tr>
	<%ElseIf txtTable ="Pattern" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">TABLE NAME</Td><td align="left" valign="top">DISPLAY PATTERN</td>
		<Td align="left" valign="top">EDIT PATTERN</Td>
	</tr>
	<%ElseIf txtTable ="Parameter" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">PARAMETER CODE</Td><td align="left" valign="top">PARAMETER NAME</td>
	</tr>

	<%ElseIf txtTable ="Trigger_Control" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">ACTIVITY</Td>
		<Td align="left" valign="top">NAME</Td>
		<td align="left" valign="top">SOURCE</td>
		<Td align="left" valign="top">TARGET</Td>
	</tr>
	<%ElseIf txtTable ="JobGroup" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td><td align="left" valign="top">NAME</td>
		<Td align="left" valign="top">PARENT</Td><Td align="left" valign="top">PARENTDESC</Td>
	</tr>
	<%ElseIf txtTable ="JobGroupMaster" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td><td align="left" valign="top">NAME</td>
		<Td align="left" valign="top">PARENT</Td><Td align="left" valign="top">PARENTDESC</Td>
		<Td align="left" valign="top">PARENTDESC(ENG)</Td>
	</tr>
	<%ElseIf txtTable ="Jobbsmaster" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td>
		<td align="left" valign="top">NAME</td>
		<Td align="left" valign="top">NAME(ENG)</Td>
	</tr>
	<%ElseIf txtTable ="Jobismaster" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td>
		<td align="left" valign="top">NAME</td>
		<Td align="left" valign="top">NAME(ENG)</Td>
	</tr>
	<%ElseIf txtTable ="Jobcfmaster" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td>
		<td align="left" valign="top">NAME</td>
		<Td align="left" valign="top">NAME(ENG)</Td>
	</tr>
	<%ElseIf txtTable ="Jobbtmaster" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td>
		<td align="left" valign="top">NAME</td>
		<Td align="left" valign="top">NAME(ENG)</Td>
	</tr>
	
	<%ElseIf txtTable ="GCMaster" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">ITEM</Td><td align="left" valign="top">CODE</td>
		<Td align="left" valign="top">NAME</Td>
	</tr>
	<%ElseIf txtTable ="Job" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td></Td>
		<td align="left" valign="top">NAME</td>
		<td align="left" valign="top">GROUP TYPE</td>
		<Td align="left" valign="top">CROP TYPE</Td>
		<Td align="left" valign="top">FIELD STATUS</Td>
		<Td align="left" valign="top">CATEGORY</Td>
		<Td align="left" valign="top">UOM</Td>
		<Td align="left" valign="top">ACCOUNT CODE</Td>
		<Td align="left" valign="top">JOB FOR BALANCESHEET CODE</Td>
		<Td align="left" valign="top">JOB FOR BALANCESHEET NAME</Td>
		<Td align="left" valign="top">JOB FOR PROFIT AND LOSS CODE</Td>
		<Td align="left" valign="top">JOB FOR PROFIT AND LOSS NAME</Td>
	</tr>
	<%ElseIf txtTable ="Jobmaster" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td></Td>
		<td align="left" valign="top">NAME</td>
		<td align="left" valign="top">GROUP TYPE</td>
		<Td align="left" valign="top">CROP TYPE</Td>
		<Td align="left" valign="top">FIELD STATUS</Td>
		<Td align="left" valign="top">CATEGORY</Td>
		<Td align="left" valign="top">UNIT-UOM</Td>
		<Td align="left" valign="top">OUTPUT-UOM</Td>
		<Td align="left" valign="top">ACCOUNT CODE</Td>
		<Td align="left" valign="top">JOB FOR BALANCESHEET CODE</Td>
		<Td align="left" valign="top">JOB FOR BALANCESHEET NAME</Td>
		<Td align="left" valign="top">JOB FOR PROFIT AND LOSS CODE</Td>
		<Td align="left" valign="top">JOB FOR PROFIT AND LOSS NAME</Td>
		<Td align="left" valign="top">JOB FOR CASHFLOW CODE</Td>
		<Td align="left" valign="top">JOB FOR CASHFLOW NAME</Td>
		<Td align="left" valign="top">JOB FOR BUDGET CODE</Td>
		<Td align="left" valign="top">JOB FOR BUDGET NAME</Td>
		<Td align="left" valign="top">BUDGET QTY</Td>
		
	</tr>
	<%ElseIf txtTable ="JobBahan" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD colspan="2">&nbsp;</TD>
		<Td align="left" valign="top">CODE</Td></Td>
		<td align="left" valign="top">NAME</td>
		<td align="left" valign="top">ITEMCODE</td>
		<Td align="left" valign="top">ITEMNAME</Td>
	</tr>
	<%ElseIf txtTable ="PeriodControl" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<Td align="left" valign="top">SYSTEM</Td></Td>
		<td align="left" valign="top">LAST PERIOD POSTED</td>
		<td align="left" valign="top">CURRENT DATE</td>
		<Td align="left" valign="top">POSTING POSITION</Td>
		<Td align="left" valign="top">BA</Td>
	</tr>
    
	<%ElseIf txtTable ="HistLog" then%>
	<tr>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<Td align="left" valign="top">USER NAME</Td>
		<td align="left" valign="top">IP ADDRESS</td>
		<td align="left" valign="top">BUSSINES UNIT</td>
		<Td align="left" valign="top">DATE ACTION</Td>
		<Td align="left" valign="top">TIME ACTION</Td>
	</tr>
	
	
	
	<%Else%>	
	<tr>
		<TD></TD><TD></TD>
		<td></td>
		<%for ni = 0 to nCountofField  %>
			<td align="left" valign="top"><%=rsGAD(ni).Name%></td>
		<%next%>
	</tr>
	<%End if%>
	  <%While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		    <td align="left" valign="top"><%=NO%></td>
			<%if txtTable <> "HistLog" then%>
			<%If txtTable <> "PeriodControl" then%>
				<%If txtTable = "GCMaster" then%>
					<td align="left" valign="top" width="10">
					<a href="TableStatistik.asp?fccode=<%=rsGAD(1)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a></td>
					<td align="left" valign="top" width="10">				
					<a href="TableStatistik.asp?fccode=<%=rsGAD(1)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></td>
				<%elseIf txtTable = "BusinessUnit" then%>
					<td align="left" valign="top" width="10">
					<a href="TableStatistik.asp?fccode=<%=rsGAD(2)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(2)%>" border="0"></a></td>
					<td align="left" valign="top" width="10">				
					<a href="TableStatistik.asp?fccode=<%=rsGAD(2)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(2)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></td>
				<%Else%>
					<td align="left" valign="top" width="10">
					<a href="TableStatistik.asp?fccode=<%=rsGAD(0)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a></td>
					<td align="left" valign="top" width="10">				
					<a href="TableStatistik.asp?fccode=<%=rsGAD(0)%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a></td>
				<%End if%>
			<%End IF%>
			<%End IF%>
		    <%If Request("txttable") ="Users" OR Request("txttable") ="Trigger_Control" or Request("txttable") ="PeriodControl" or Request("txttable") ="BusinessUnit" or Request("txttable") ="HistLog" then%>
				<%for ni = 0 to nCountofField %> 
					<!--Additional AST 20180503, for Handle Link Detail Business Unit -->
					<%If Request("txttable") ="BusinessUnit" And  rsGAD(ni).Name ="FCCODE" then %>
						<td align="left" width="597">
							<a target="_blank" href="BusinessUnitDetails.Asp?fccode=<%=rsGAD(2)%>&FCNAME=<%=rsGAD(3)%>"><%=rsGAD(ni)%>
							</a>						
	                    </td>
						
					<%Else%>
						<td align="left" title="<%=rsGAD(ni).Name%>" valign="top"><%=rsGAD(ni)%></td>
					<%End If%>
				<%next%>
			<%Else%>
				<%for ni = 0 to nCountofField -4%>
					<td align="left" title="<%=rsGAD(ni).Name%>" valign="top"><%=rsGAD(ni)%></td>
				<%next%>
			<%End If%>
			
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
   </TABLE>

<%'End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
