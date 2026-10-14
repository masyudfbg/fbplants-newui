<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Accrual THR dan Bonus</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script language="javascript" type="text/javascript" src="ajax/validate_ajax.js"></script>

<script language="javascript" type="text/javascript">

                function showHint(str)
					{
				var xmlhttp;
				var unit;
				var output;
				var oth;
				var jobname;
					if (str=="")
					  {
					  document.getElementById("mandays").innerHTML="";
					  document.getElementById("unit").innerHTML="";
					  document.getElementById("output").innerHTML="";
					  document.getElementById("oth").innerHTML="";
					  document.getElementById("jobname").innerHTML="";
				
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  xmlhttp=new XMLHttpRequest();
					  unit=new XMLHttpRequest();
					  output=new XMLHttpRequest();
					  oth=new XMLHttpRequest();
					  jobname=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
					  unit=new ActiveXObject("Microsoft.XMLHTTP");
					  output=new ActiveXObject("Microsoft.XMLHTTP");
					  oth=new ActiveXObject("Microsoft.XMLHTTP");
					  jobname=new ActiveXObject("Microsoft.XMLHTTP");
					  }
					  
					  
						xmlhttp.onreadystatechange=function()
						  {
						  if (xmlhttp.readyState==4 && xmlhttp.status==200)
							{
							document.getElementById("mandays").innerHTML=xmlhttp.responseText;
							}
						  }
						xmlhttp.open("GET","ajax/MANDAYS.asp?ajaxUoM=mandays&q="+str,true);
						xmlhttp.send();
						
						
						unit.onreadystatechange=function()
						  {
						  if (unit.readyState==4 && unit.status==200)
							{
							document.getElementById("unit").innerHTML=unit.responseText;
							}
						  }
						unit.open("GET","ajax/MANDAYS.asp?ajaxUoM=unit&q="+str,true);
						unit.send();
						
						output.onreadystatechange=function()
						  {
						  if (output.readyState==4 && output.status==200)
							{
							document.getElementById("output").innerHTML=output.responseText;
							}
						  }
						output.open("GET","ajax/MANDAYS.asp?ajaxUoM=output&q="+str,true);
						output.send();
						
						oth.onreadystatechange=function()
						  {
						  if (oth.readyState==4 && oth.status==200)
							{
							document.getElementById("oth").innerHTML=oth.responseText;
							}
						  }
						oth.open("GET","ajax/MANDAYS.asp?ajaxUoM=oth&q="+str,true);
						oth.send();
						
						jobname.onreadystatechange=function()
						  {
						  if (jobname.readyState==4 && jobname.status==200)
							{
							document.getElementById("jobname").innerHTML=jobname.responseText;
							}
						  }
						jobname.open("GET","ajax/MANDAYS.asp?ajaxUoM=jobname&q="+str,true);
						jobname.send();
					
					}

				

</script>
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
function periksa(form)
 { if(form.mandays.value > 1)
    {alert("mandays diisi paling besar 1 !");
     form.mandays.focus();
     form.mandays.select();
     return(false);
    }
 
  return(true);
 }
 

function harusangka(evt) {
    var e = evt || window.event;
    var key = e.keyCode || e.which;

    if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
    // numbers   
    key >= 48 && key <= 57 ||
	e.ctrlKey && key == 97 ||
    // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
    // left and right arrows
    key == 37 || key == 39 ||
    // Del and Ins
	key == 46 ) {
        // input is VALID
		}
	else {
			// input is INVALID
			e.returnValue = false;
			if (e.preventDefault) e.preventDefault();
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
Dim JOBCODE, LOCATIONTYPE, LOCATIONCODE, AMOUNT, REMARKS
Dim cmdView

Dim v_txtYear : v_txtYear = Request("txtYear")
Dim Header_Allowded : Header_Allowded = request("Header_Allowded")

Dim MinYr, MaxYr, mon, Yr
Yr = Year(Session("v_Employee"))
MinYr = Yr - 2
MaxYr = Yr  + 2

If len(v_txtyear)= 0 then v_txtyear = Yr

JOBCODE = request("JOBCODE")
LOCATIONTYPE = request("LOCATIONTYPE")
LOCATIONCODE = request("LOCATIONCODE")
AMOUNT = request("AMOUNT")
REMARKS = request("REMARKS")


if Request("txtJenis") = "New" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "Delete" then
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		cmdView	= "True"
	else
		cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis


txtTable = Request("txtTable")
txtModule = Request("txtModule")
txtSearch = Request("txtSearch")
txtField = Request("txtField")
txtJenis = request("txtJenis")

if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

'-------------ISI Payment Form
 
Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('12/31/" & v_txtYear & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('12/31/" & v_txtYear & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Employee')"
			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
'response.Write oSQLQuery
oRecordSet.Close
Set oRecordSet = Nothing
%>
<form name="frmEntry" method="post" action="accrual_thrbonus.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Year</td>
			<td align="left">
				<select name="txtyear">
				<%For i = MinYr to MaxYr%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtyear) then%> Selected <%End If%>><%=i%></option>
				<%Next%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td  align="left">Allowdedcode</td>
			<td align="left">
				<select name="Header_Allowded" >
					<%
					Dim rsallowded, strallowded
					strallowded = " select fccode, fcname" & _
								  " from allowded" & _
								  " where fcba = '" & Session("fcBA") & "'" & _
								  "       and fccode in ('A59')" & _
								  "       and man_auto = 'M'" & _
								  "       and includedpayrollprocessing = 'TRUE'" & _
								  " order by fccode asc"
					
					Set rsallowded = Server.CreateObject("ADODB.Recordset")
					rsallowded.Open strallowded, connstr, adOpenStatic 
					
					if not rsallowded.eof then
					Do While Not rsallowded.eof
					%>
						<option value="<%=rsallowded("fccode")%>" <%if rsallowded("fccode") = Header_Allowded then%> Selected <%end if%>><%=rsallowded("fccode")%> : <%=rsallowded("fcname")%></option>
					<%
						rsallowded.movenext
					loop
					end if
					rsallowded.Close
					set rsallowded = Nothing
					%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="Printaccrual_thrbonus.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
				<%
				dim dataModule,dataSubmodule,dataDetail,hiddenHadir
				dataModule 		= "Employee"
				dataSubmodule 	= "Accrual THR dan Bonus"
				if Header_Allowded = "" then
					dataDetail 	= ""
					hiddenHadir = "visibility: hidden;"
				else
					dataDetail 	= Header_Allowded
					hiddenHadir = "visibility: visible;"
				
				end if
				%>
				<a class="jqplot-image-button" style="<%=hiddenHadir%>" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td>
		</tr>
	</table>  
</form>
<%
if cmdView = "True" and len(trim(v_txtyear)) > 0 and len(trim(Header_Allowded)) > 0 Then
	Dim rsGAD,intRecordsPerPage,cColor
	set rsGAD = server.CreateObject("Adodb.recordset")
	q = " select * from dual"
	'response.Write q	
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.CursorType = 3
	rsGAD.CursorLocation = 3
	rsGAD.Open q, connstr
%>
   
	<form action="accrual_thrbonus.asp" method="post" name="eForm">
	<table cellpadding="0" cellspacing="0" width="50%">
			<tr>
				<td align="right">
					<font face="Verdana" size="1">Jobcode</font>
				</td>
				<td align="left">
					<input type="text" size="15" name="JOBCODE" value="<%=JOBCODE%>" onkeypress="return harusangka(event);">
				    <font color="#FF0000">&nbsp;*&nbsp;</font>
					<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Job&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font></a>
				</td>
			</tr>
			<tr>
				<td align="right">
					<font face="Verdana" size="1">Target Type</font>
				</td>
				<td align="left">
					<input type="text" size="15" name="LOCATIONTYPE" value="<%=LOCATIONTYPE%>">
					<font color="#FF0000">&nbsp;*&nbsp;</font>
					<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Location Type</font></a>
				</td>
			</tr>
			<tr>
				<td align="right">
					<font face="Verdana" size="1">Target Code</font>
				</td>
				<td align="left">
					<input type="text" size="15" name="LOCATIONCODE" value="<%=LOCATIONCODE%>">
					<font color="#FF0000">&nbsp;*&nbsp;</font>
				</td>
			</tr>

		<tr>
			<td align="right">
				<font face="Verdana" size="1">Amount</font>
			</td>
			<td align="left">
				<input type="text" size="20" name="AMOUNT" value="<%=AMOUNT%>" onkeypress="return harusangka(event);">
			</td>
		</tr>

		<tr>
			<td align="right">
				<font face="Verdana" size="1">Remarks</font>
			</td>
			<td align="left">
				<input type="text" size="50" name="REMARKS" value="<%=REMARKS%>">
			</td>
		</tr>
	    <TR>
			<td colspan="2">
				<input type="hidden" name="txtyear" value="<%=v_txtyear%>">
				<input type="hidden" name="Header_Allowded" value="<%=Header_Allowded%>">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%If val_IsTrue = "Y" Then%>
				<%IF Request("txtJenis")="Edit" then%>
					<input type="hidden" name="txtJenis" value="Edit">
					<input Type="Submit"  name="Add" value="Update"  OnClick ="periksa(this.form)" >
				<%Else%>
					<input type="hidden" name="txtJenis" value="New">
					<input Type="Submit"  name="Add" value="Add" OnClick ="periksa(this.form)">
				<%End if%>
				<%End if%>
				<input Type="Reset"  name="Reset" value="Reset">
				<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
			</td>
		</tr>
	</table>

	<table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="2">&nbsp;</td>
		<td align="left">Year</td>
		<td align="left">Allowded Code</td>
		<td align="left">Allowded Name</td>
		<td align="left">Jobcode</td>
		<td align="left">Target Type</td>
		<td align="left">Target Code</td>
		<td align="left">Amount</td>
		<td align="left">Remarks</td>
	</tr>
   <%
   
   if Not rsGAD.eof then
      nCountofField = rsGAD.fields.count - 1
	  no = 1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			
			<%If val_IsTrue = "Y" Then%>
				<td valign="middle" width="30">
					<a href="accrual_thrbonus.asp?txtyear=<%=v_txtYear%>&Header_Allowded=<%=Header_Allowded%>&JOBCODE=<%=rsGAD("JOBCODE")%>&locationtype=<%=rsGAD("locationtype")%>&locationcode=<%=rsGAD("locationcode")%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
				</td>
				<td>
					<a href="accrual_thrbonus.asp?txtyear=<%=v_txtYear%>&Header_Allowded=<%=Header_Allowded%>&JOBCODE=<%=rsGAD("JOBCODE")%>&locationtype=<%=rsGAD("locationtype")%>&locationcode=<%=rsGAD("locationcode")%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				</TD>
			<%else%>
				 <TD COLSPAN="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if%>	 
			</td>
		    <%for i = 0 to nCountofField - 4%>
			    <td align="left" title="<%=rsGAD(i).Name%>" valign="top"><%=rsGAD(i)%></td>
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

</form>
<script language="JavaScript" type="text/javascript">
//You should create the validator only after the definition of the HTML form
  var frmvalidator  = new Validator("eForm");
  frmvalidator.addValidation("EMPLOYEECODE","req","Please enter Emp.Code");
  frmvalidator.addValidation("JOBCODE","req","Please enter Job.Code");
  frmvalidator.addValidation("LOCATIONTYPE","req","Please enter Location Type");
  frmvalidator.addValidation("LOCATIONCODE","req","Please enter Location Code");
  frmvalidator.addValidation("MANDAYS","ltelmnt=1.001","Total Mandays perhari lebih dari 1");
  frmvalidator.addValidation("MANDAYS","gtelmnt=0","Total Mandays perhari lebih dari 1");

</script>
<!--  
  frmvalidator.addValidation("MANDAYS","numeric");
  frmvalidator.addValidation("OUTPUT","numeric");
  frmvalidator.addValidation("UNIT","numeric");
  frmvalidator.addValidation("RATE","numeric");
  frmvalidator.addValidation("OTHRS","numeric","5");

  frmvalidator.addValidation("LastName","req");
  frmvalidator.addValidation("LastName","maxlen=20");
  
  frmvalidator.addValidation("Email","maxlen=50");
  frmvalidator.addValidation("Email","req");
  frmvalidator.addValidation("Email","email");
  
  
  frmvalidator.addValidation("Address","maxlen=50");
  frmvalidator.addValidation("Country","dontselect=0");
-->
   
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
