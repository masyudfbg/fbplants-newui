<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<%
	' Note
	' 1. Tambahkan Record Cancelled table rowstate
	' 2. Buat Control Status dengan Luname RECURRING table controlstatus
	' 3. Tambahkan Package RECURRING_API
%>
<html>
<head>
<title>ePlaS - Piezo Type</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script language="javascript" type="text/javascript" src="ajax/validate_ajax.js"></script>
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-3.1.0.min.js"></script>
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
 

Dim rs_data, query_data
Dim totalamountupdate
Dim comm, param, fcba_, lookupcode_, date_from_, Date_until_, fcentry_, fcedit_, fcip_, ResultCommand_
Dim fccode, fcname, remarks

lookupcode = Request("lookupcode")
fccode = Request("fccode")
fcname = Request("fcname")
remarks = Request("remarks")



IF Request("txtJenis") = "New" OR Request("txtJenis") = "Header" OR Request("txtJenis") = "Delete" OR Request("txtJenis") = "Edit" THEN
	IF Request("txtJenis") = "Delete" THEN
		q =	" DELETE" & _
            " FROM piezotype_tab" & _
            " WHERE fccode = '" & fccode & "'"
		connStr.execute(q)
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Header" THEN
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "New" THEN
		q = " INSERT INTO piezotype_tab ( fccode," & _
			"                             fcname," & _
			"                             remarks," & _
			"                             fcentry," & _
			"                             fcedit," & _
			"                             fcip," & _
			"                             lastupdate," & _
			"                             lasttime)" & _
			" VALUES ( '" & fccode & "'," & _
			"          '" & fcname & "'," & _
			"          '" & remarks & "'," & _
			"          '" & Session("CodeName") & "'," & _
			"          '" & Session("CodeName") & "'," & _
			"          '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
			"          SYSDATE," & _
			"          TO_CHAR(SYSDATE, 'hh24:mi'))"
		connStr.execute(q)
			
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Edit" AND Request("Add") = "Update" THEN
		q =	" UPDATE piezotype_tab" & _
			" SET fcname = '" & fcname & "'" & _
			"     , remarks = '" & remarks & "'" & _
			" WHERE fccode = '" & fccode & "'"
		connStr.execute(q)
		cmdView	= "True"
	ELSE
		cmdView	= "True"
	END IF
	'Response.Write "sql: " & q
ELSE
    cmdView	= request("cmdView")
END IF

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis, cmdView

Dim lookupcode, documentname, startingdate, finisheddate,defaultexecuted, intervalinmonth, numberofruns, amount, isusesettlement, state, accrualjob
Dim jobcode, targettype, debetcredit, detailamount

txtTable = Request("txtTable")
txtModule = Request("txtModule")
txtSearch = Request("txtSearch")
txtField = Request("txtField")
txtJenis = request("txtJenis")


if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

 '-------------ISI Payment Form
%>
<form name="frmEntry" method="post" action="piezotype.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Search Code</td>
			<td align="left">
				<input type="text" maxlength = "20" name="lookupcode" size="30" value="<%=lookupcode%>">
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">
			</td>
		</tr>
	</table>  
</form>
<%
IF cmdView = "True" THEN
	Dim rs_datadtl,intRecordsPerPage,cColor
	q = " SELECT fccode," & _
	    "        fcname," & _
	    "        remarks," & _
	    "        fcentry," & _
	    "        fcedit," & _
	    "        fcip," & _
	    "        lastupdate," & _
	    "        lasttime" & _
	    " FROM piezotype_tab"
	if "" & lookupcode <> "" then
	    q = q & " WHERE fccode = '" & lookupcode & "'"
	end if
		q = q & " ORDER BY fccode"
	'response.Write q	
	Set rs_datadtl = Server.CreateObject("ADODB.Recordset")
	rs_datadtl.CursorType = 3
	rs_datadtl.CursorLocation = 3
	rs_datadtl.Open q, connstr
	
%>
   
<form action="piezotype.asp" method="post" name="eForm">
	<table class="data" width="100%">
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Code</font>
			</td>
			<td align="left">
				<input type="text" maxlength = "20" size="30" name="fccode" value="<%=fccode%>" <%if txtJenis = "Edit" then%>disabled<%end if%> required>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Description</font>
			</td>
			<td align="left">
				<input type="text" maxlength = "50" size="50" name="fcname" value="<%=fcname%>" required>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Remarks</font>
			</td>
			<td align="left">
				<textarea name="remarks" maxlength = "100" cols="50" rows="3"><%=remarks%></textarea>
			</td>
		</tr>
		
	    <TR>
			<td colspan="2" align="left">
				<input type="hidden" name="lookupcode" value="<%=lookupcode%>">
				<%if txtJenis = "Edit" then%>
				<input type="hidden" name="fccode" value="<%=fccode%>">
				<%end if%>
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
		<th>NO</th>
		<th colspan="2">&nbsp;</th>
		<th>CODE</th>
		<th>DESCRIPTION</th>
		<th>REMARKS</th>	
	</tr>
<%
	'response.write Session("v_Employee")
   
	if Not rs_datadtl.eof then
		nCountofField=rs_datadtl.fields.count-1
		no=1
		While Not rs_datadtl.EOF
			if no  mod 2 = 1 then
				cColor = "#EEEEEE"
			else
				cColor = "#E5E5E5"
			end if
%>
	<tr bgcolor="<%=cColor%>">
		<td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
				<a href="piezotype.asp?lookupcode=<%=lookupcode%>&fccode=<%=rs_datadtl("fccode")%>&fcname=<%=rs_datadtl("fcname")%>&remarks=<%=rs_datadtl("remarks")%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rs_datadtl(3)%>" border="0">
				</a>
			</td>
			<td>
				<a href="piezotype.asp?lookupcode=<%=lookupcode%>&fccode=<%=rs_datadtl("fccode")%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rs_datadtl(3)%>" border="0">
				</a>
			</td>
		<td align="left" title="<%=rs_datadtl("fccode").Name%>" valign="top"><%=rs_datadtl("fccode")%></td>
		<td align="left" title="<%=rs_datadtl("fcname").Name%>" valign="top"><%=rs_datadtl("fcname")%></td>
		<td align="left" title="<%=rs_datadtl("remarks").Name%>" valign="top"><%=rs_datadtl("remarks")%></td>
	</tr>
<%
			rs_datadtl.MoveNext
			intRecordsPerPage = intRecordsPerPage - 1
			no=no+1
		Wend
	else
		response.write ("<tr bgcolor='#E5E5E5'><td colspan='6' align='left'>Data is empty</td></tr>")
	end if 
	rs_datadtl.Close
	set rs_datadtl = Nothing
%>
	</table>
</form>

<%End if%>
		</div>
	</div>
	<!--#include file="bottom.jsp"-->
	
	<script src="/js/jquery.inputmask.bundle.min.js"></script>
	<script src="/js/inputmask.binding.js"></script>
</body>
</html>
