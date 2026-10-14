<%Option Explicit%>
<%
'On Error resume Next 
%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Internal Destination</title>
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

/*
function CheckIsActive()
{
 var cIsActive = document.getElementById('ISACTIVE');
 if (cIsActive.checked != true )
	{
		alert("Is Active Belum di Click");
	}
}
*/

function calcParam()
{
	var cIsActivation = document.getElementById('ACTIVATION').checked;
	if(cIsActivation)
	{
		document.getElementById("hide_ACTIVATION").value ="1";
	}else
	{
		document.getElementById("hide_ACTIVATION").value ="0";
	}		
}

function validasi(form){
 
	<!-- Start additional Validation by AST 28 Juni 2018 -->
	if (form.INTDESTID.value == ""){
		alert("Anda belum mengisikan Internal Destination Id.");
		form.INTDESTID.focus();
		return (false);
	}
	
	if (form.INTDESTNAME.value == ""){
		alert("Anda belum mengisikan Internal Destination Name.");
		form.INTDESTNAME.focus();
		return (false);
	}
	
	/*
	var cActivation = document.getElementById('ACTIVATION').checked;
	if(cActivation)
		{
		}else
		{
			alert("Anda belum mengisikan Activation.");
			form.ACTIVATION.focus();
			return (false);
		}	
	*/
	
	if (form.ADDRESS.value == ""){
		alert("Anda belum mengisikan Address.");
		form.ADDRESS.focus();
		return (false);
	}	
	
	if (form.CITY.value == ""){
		alert("Anda belum mengisikan City.");
		form.CITY.focus();
		return (false);
	}		
	
	if (form.STATE.value == ""){
		alert("Anda belum mengisikan State.");
		form.STATE.focus();
		return (false);
	}
	
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
 
<%

 '-------------awal----------------------- 

Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim vINTDESTID, vINTDESTNAME, vACTIVATION, vADDRESS, vCITY, vSTATE, vREMARKS, txt_DestinationName, hide_ACTIVATION
Dim cmdView
txttable=request("txttable")


if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 
 	hide_ACTIVATION = 0
	txt_DestinationName=request("txt_DestinationName")
	hide_ACTIVATION=request("hide_ACTIVATION")
	
	vINTDESTID =request("INTDESTID")
	vINTDESTNAME =request("INTDESTNAME")
	vACTIVATION =request("ACTIVATION")
	vADDRESS =request("ADDRESS")
	vCITY =request("CITY")
	vSTATE =request("STATE")
	vREMARKS =request("REMARKS")

	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM INTDESTINATION WHERE  INTDESTID='" & vINTDESTID & "' AND fcBA='" & Session("fcBA") & "'"
		response.write q
		connStr.execute(q)
		cmdView	= "True"
		vINTDESTID =request("") ''For Clear Text
	   
	elseif Request("txtJenis")="New" then
		if Request("hide_ACTIVATION")="" then hide_ACTIVATION=0
	
		q	=	" INSERT INTO INTDESTINATION (" & _
				" INTDESTID, " & _
				" INTDESTNAME, " & _
				" ACTIVATION, " & _
				" ADDRESS, " & _
				" CITY, " & _
				" STATE, " & _
				" REMARKS, " & _
				" FCENTRY, " & _
				" FCEDIT, " & _
				" FCIP, " & _
				" FCBA, " & _
				" LASTUPDATE, " & _
				" LASTTIME " & _
				" ) values (" & _
				"   '" & vINTDESTID & "'" & _			
				" , '" & vINTDESTNAME & "'" & _
				" , '" & hide_ACTIVATION & "'" & _
				" , '" & vADDRESS & "'" & _
				" , '" & vCITY & "'" & _
				" , '" & vSTATE & "'" & _
				" , '" & vREMARKS & "'" & _
				" , '" & Session("CodeName") & "'" & _					
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _					
				" , '" & Session("fcBA") & "'" & _					
				" , sysdate " & _										
				" , To_Char(sysdate,'HH24:MI') " & _	
				" )"			 

		'response.write q
		'response.Flush()
		connStr.execute(q)
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	if Request("hide_ACTIVATION")="" then hide_ACTIVATION=0
	
       q=	"  UPDATE INTDESTINATION SET" & _ 
			"  INTDESTNAME ='" & vINTDESTNAME & "'" & _
			", ACTIVATION ='" & hide_ACTIVATION & "'" & _
			", ADDRESS ='" & vADDRESS & "'" & _
			", CITY ='" & vCITY &"'" & _
			", STATE ='" & vSTATE & "'" & _
			", REMARKS ='" & vREMARKS & "'" & _
			", LASTUPDATE = sysdate " & _
			", LASTTIME = To_Char(sysdate,'HH24:MI') " & _			
			"  WHERE  INTDESTID='" & vINTDESTID &"'"  & _
			"  AND fcBA='" & Session("fcBA") & "'" 
			
		''response.Write(q)
		connStr.execute(q)
		connStr.execute("commit")
		
		'response.Flush()
		''response.write q			
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

'' Event GetInitEdit / Mengisi Data
''============================================
if request("txtJenis")="Edit" then
 Dim rsVHx
 
q	=	" SELECT " & _
		" INTDESTID, INTDESTNAME, ACTIVATION, " & _
		" ADDRESS, CITY, STATE, REMARKS " & _			
		" FROM INTDESTINATION " & _
		" WHERE fcba = '" & Session("fcBA") & "'" & _
		" AND INTDESTID='" & vINTDESTID &"'"

''response.Write(q)				
Set rsVHx = Server.CreateObject("ADODB.Recordset")
rsVHx.Open q, connstr

	if not rsVHx.eof then 
		While not rsVHx.eof
			
			hide_ACTIVATION=rsVHx("ACTIVATION")
			vINTDESTID =rsVHx("INTDESTID")
			vINTDESTNAME =rsVHx("INTDESTNAME")
			vACTIVATION =rsVHx("ACTIVATION")
			vADDRESS =rsVHx("ADDRESS")
			vCITY =rsVHx("CITY")
			vSTATE =rsVHx("STATE")
			vREMARKS =rsVHx("REMARKS")
		
		rsVHx.movenext
		Wend
		end if
	end if		
	
		
 '-------------ISI Payment Form
  %>
  
<form name="frmEntry" method="post" action="IntDestination.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE INTERNAL DESTINATION</strong></font></div></td>
          </TR>
		<tr>
			<td align="left" width="15%">Destination Name</td><td align="left"><input type="text" size="15" name="txt_DestinationName" value="<%=request("txt_DestinationName")%>" ></td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;
		<a href="PrintIntDestination.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_DestinationName=<%=REQUEST("txt_DestinationName")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
		</tr>
    </table>  
 </form>
 <%
  
 if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
	
	'' Query Menampilkan Data di Table Bawah  
	q	=	" SELECT " & _
			" INTDESTID, INTDESTNAME, ACTIVATION, " & _
			" ADDRESS, CITY, STATE, REMARKS " & _			
			" FROM INTDESTINATION " & _
			" WHERE fcba = '" & Session("fcBA") & "'"
	
	if  len(trim(request("txt_DestinationName"))) > 0 then
		q = q & " AND INTDESTNAME='" & request("txt_DestinationName") & "'"
	end if
	q = q		
			
    Set rsGAD = Server.CreateObject("ADODB.Recordset")
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr 
   
  %>
   
   <!-- Form Entry And Edit --> 
   
   <form action="IntDestination.asp" method="post" name="eForm" onSubmit="javascript:return validasi(this)">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
					
		<tr>			
			<td width="189" align="left">Internal Destination Id</td>
			<%IF Request("txtJenis")="Edit" then%>			
				<td width="180" align="left"><input type="text" size="20" maxlength="20"  name="INTDESTID" id="INTDESTID" value="<%=vINTDESTID%>"readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
			<%else%>
				<td width="180" align="left"><input type="text" size="20" maxlength="20"  name="INTDESTID" id="INTDESTID" value="<%=vINTDESTID%>"><font color="#FF0000">*</font></td>
			<%end if%>
			
			<td width="149" align="left">Internal Destination Name</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="INTDESTNAME" id="INTDESTNAME" value="<%=vINTDESTNAME%>"></td>
			
			<td width="149" align="left"></td>	
			<td width="180" align="left"> 
				<input type="Hidden" name="hide_ACTIVATION" id="hide_ACTIVATION" value="<%=hide_ACTIVATION%>">
				<input type="checkbox" size="10" maxlength="30" onclick="calcParam()"  name="ACTIVATION" id="ACTIVATION" value="<%if Request("hide_ACTIVATION")="1" then%>" checked "<% end if %>" />  Activation
			</td>
		</tr>
		
		<tr>
			<td width="149" align="left">Address</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="ADDRESS" id="ADDRESS" value="<%=vADDRESS%>"></td>
			
			<td width="149" align="left">City</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="CITY" id="CITY" value="<%=vCITY%>"></td>
		</tr>
		<tr>
			<td width="149" align="left">State</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="STATE" id="STATE" value="<%=vSTATE%>"></td>

			<td width="149" align="left">Remarks</td>
			<td width="180" align="left"><input type="text" size="20" maxlength="30"  name="REMARKS" id="REMARKS" value="<%=vREMARKS%>"></td>
		</tr>

		<tr> 
			<td width="189" align="left"></td>	
		</tr>
		<tr> 
			<td width="189" align="left"></td>	
		</tr>
		
		<tr>
			<td colspan="2" align="left">
				  <input type="Hidden" name="txt_DestinationName" value="<%=REQUEST("INTDESTNAME")%>">				   
				<!--
				  <input type="text" name="hide_ISTRANSPORTER" value="<%=REQUEST("ISTRANSPORTER")%>">	
				  <input type="text" name="hide_ISBLOCK" 	   value="<%=REQUEST("ISBLOCK")%>">	
				  <input type="text" name="hide_ISACTIVE" 	   value="<%=REQUEST("ISACTIVE")%>">	
				-->
				
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
   </form>
  <!--#include file="validation_show.Asp"-->    
<%

	' Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	' howmanyfields=rsGAD.fields.count-1
	
	' JumRec=1
	' intRecordsPerPage = 20
	
	' rsGAD.PageSize = intRecordsPerPage
	' intNumberOfPages = 0
	' intNumberOfPages =  rsGAD.PageCount
	' intRecords = rsGAD.RecordCount
	''Set our AbsolutePage property.
	' if not rsGAD.EOF then
		' If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  ' rsGAD.AbsolutePage = 1
		' Else
		  ' rsGAD.AbsolutePage = Request("iPageNo")
		' End If  
	' end if
	
	' if Request("iPageNo") = "" then
		' nPage = 1
	' Else
		' nPage = Request("iPageNo")
	' End if
	
   ' rsGAD.PageSize = intRecordsPerPage
   ' intNumberOfPages = 0
   ' intNumberOfPages =  rsGAD.PageCount
   ' intRecords = rsGAD.RecordCount   
   
   %>
   <table class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD align="center"colspan="2">ACTION</TD>
		<TD align="left">Internal Destination Id</TD>	
		<TD align="left">Internal Destination Name</TD>	
		<TD align="left">Activation</TD>
		<TD align="left">Address</TD>	
		<TD align="left">City</TD>	
		<TD align="left">State</TD>	
		<TD align="left">Remarks</TD>
	</tr>
   
   <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-1
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="IntDestination.asp?INTDESTID=<%=rsGAD("INTDESTID")%>&txttable=<%=request("txttable")%>&hide_ACTIVATION=<%=rsGAD("ACTIVATION")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>		
			</td>
			<td valign="middle" width="30">
			<a href="IntDestination.asp?INTDESTID=<%=rsGAD("INTDESTID")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			
				<% if rsGAD(i).Name="ACTIVATION" Then %>  <!--Filter Data Only Checkbox -->
			
					<%if rsGAD(i).Name="ACTIVATION" AND rsGAD("ACTIVATION")="1" Then%>
						<td align="center" ><input type="checkbox" name="myTextEditBox" disabled readonly value="1" checked></td>
						<%Else%>
						<td align="center" ><input type="checkbox" name="myTextEditBox" disabled readonly value="0"></td>
					<%End If%>
					
				<%Else%>
					<td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%></td>
				<%End If%>	
			    
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
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->
</body>
</html>
