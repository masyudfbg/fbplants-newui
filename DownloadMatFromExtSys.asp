<%'Option Explicit%>
<%'On Error resume Next%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Download Material Master From Ext System</title>
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
</head>
<body>
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
Dim rsModule,qModule,qInsert, rsModuleCheck
		


v_FCCODE =request("v_FCCODE")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
	    q = " delete from user_fcba_tab " & _
			" where " & _
			" fcusercode = '" & request("v_FCCODE") & "'"	

		   connStr.execute(q)
		   cmdView	= "True"
	elseif Request("txtJenis")="New" or Request("txtJenis")="Edit" then
	
		   qCompanyAll = "Select MATERIALCODE, " & _
						 "		 MATERIALDESC, " & _
						 "		 MATERIALGROUP, " & _
						 "		 BASEUOM, " & _
						 "		 DIVISION, " & _
						 "		 EXTMATGROUP, " & _
						 "		 ISDELETE, " & _
						 "		 USERINPUT, " & _
						 "		 DATEINPUT, " & _
						 "		 USERMODIFY, " & _
						 "		 DATEMODIFY " & _
						 "  From MATERIALMASTER@IPLAS192L a " & _
						 "  where a.materialcode in (select b.materialcode " & _
						 "							 from Materialmastertrans@IPLAS192L b " & _
						 "							where b.bridgestate = 0 And b.errorflag=0" & _
						 "							  and b.appsys = 'IPLAS') " & _
						 " order by a.materialcode "
			  
	    Set rsModuleCheck = Server.CreateObject("ADODB.Recordset")
		rsModuleCheck.CursorType = 3
		rsModuleCheck.CursorLocation = 3
		rsModuleCheck.Open qCompanyAll, connstr
		
		doneContinue = "TRUE"
		tempCounterDefault = 0
		if Not rsModuleCheck.eof then 
		    While Not rsModuleCheck.EOF
			    v_UnixCode  = rsModuleCheck("MATERIALCODE") 
				if request(trim("data_" & v_UnixCode))="on" Then
					response.Write v_UnixCode
					strQueryStatus = "BEGIN BRIDGINSAP_FFB_API.UpdateGenMatMaster('" & rsModuleCheck("MATERIALCODE") & "'); END; "
	                connStr.execute(strQueryStatus)	
				end if
				rsModuleCheck.MoveNext
		    wend	
		end if	
		
		rsModuleCheck.Close
		set rsModuleCheck = Nothing
		response.Write "<font face=Verdana size=1 color=blue>Upload FFB To External System Completed..</font>"
		cmdView = "True"
		
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
 
 <form name="frmEntry" method="post" action="DownloadMatFromExtSys.Asp">
	<table class="data" width="500">
		
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

   txtCekUsers="True"


if cmdView = "True" and txtCekUsers="True" Then 
   Dim rsVH,intRecordsPerPage,cColor
   set rsVH = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
  
	  
             q = "Select MATERIALCODE, " & _
				 "		 MATERIALDESC, " & _
				 "		 MATERIALGROUP, " & _
				 "		 BASEUOM, " & _
				 "		 DIVISION, " & _
				 "		 EXTMATGROUP, " & _
				 "		 ISDELETE, " & _
				 "		 USERINPUT, " & _
				 "		 DATEINPUT, " & _
				 "		 USERMODIFY, " & _
				 "		 DATEMODIFY " & _
				 "  From MATERIALMASTER@IPLAS192L a " & _
				 "  where a.materialcode in (select b.materialcode " & _
				 "							 from Materialmastertrans@IPLAS192L b " & _
				 "							where b.bridgestate = 0 And b.errorflag=0" & _
				 "							  and b.appsys = 'IPLAS') " & _
				 " order by a.materialcode "
				
    
   Set rsVH = Server.CreateObject("ADODB.Recordset")
   rsVH.CursorType = 3
   rsVH.CursorLocation = 3
   rsVH.Open q, connstr%>
   <form action="DownloadMatFromExtSys.Asp" method="post" name="eForm">
   <table class="entry" cellpadding="1" cellspacing="1" width="620">
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<td align="center" valign="top" >Download</td>
		<td align="center" valign="top" >NO</td>
		<td align="center" valign="top" >MATERIALCODE</td>
		<td align="center" valign="top" >MATERIALDESC</td>
		<td align="center" valign="top" >MATERIALGROUP</td>
		<td align="center" valign="top" >BASEUOM</td>
		<td align="center" valign="top" >DIVISION</td>
		<td align="center" valign="top" >EXTMATGROUP</td>
		<td align="center" valign="top" >ISDELETE</td>
		<td align="center" valign="top" >USERINPUT</td>
		<td align="center" valign="top" >DATEINPUT </td>
		<td align="center" valign="top" >USERMODIFY </td>
		<td align="center" valign="top" >DATEMODIFY</td>
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
		  
          v_UnixCode  = rsVH("MATERIALCODE") 

		  %>
		  <tr bgcolor="<%=cColor%>">
			  
			  
			<td align="left" title="TEST" valign="top">
				  <input name="cb_<%=v_UnixCode%>" id="cb_<%=v_UnixCode%>" type="checkbox" value="off" onClick="SetCheckBox('<%=v_UnixCode%>', 'cb_<%=v_UnixCode%>')">
				  <div  id="div_<%=v_UnixCode%>" name="div_<%=v_UnixCode%>">
					 <input id="data_<%=v_UnixCode%>" name="data_<%=v_UnixCode%>" type="hidden" value="off" >
				  </div>			
			</td>	
            <td align="Right" valign="top"><%=NO%></td>
            <td align="Right" valign="top"><%=rsVH("MATERIALCODE")%></td>
			<td align="Right" valign="top"><%=rsVH("MATERIALDESC")%></td>
			<td align="Right" valign="top"><%=rsVH("MATERIALGROUP")%></td>
			<td align="Right" valign="top"><%=rsVH("BASEUOM")%></td>
			<td align="Right" valign="top"><%=rsVH("DIVISION")%></td>
			<td align="Right" valign="top"><%=rsVH("EXTMATGROUP")%></td>
			<td align="Right" valign="top"><%=rsVH("ISDELETE")%></td>
			<td align="Right" valign="top"><%=rsVH("USERINPUT")%></td>
			<td align="Right" valign="top"><%=rsVH("DATEINPUT")%></td>
			<td align="Right" valign="top"><%=rsVH("USERMODIFY")%></td>
			<td align="Right" valign="top"><%=rsVH("DATEMODIFY")%></td>
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
	  <input type="hidden" name="txtJenis" value="Edit">
	  <input Type="Submit"  name="Add" value="Submit">
	
	 
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
