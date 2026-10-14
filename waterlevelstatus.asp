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
<title>ePlaS - Water Level Status</title>
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
Dim lowlevel, maxlevel, status, waterlevelid, is_in_data

lookupcode = Request("lookupcode")
lowlevel = Request("lowlevel")
maxlevel = Request("maxlevel")
status = Request("status")
waterlevelid = Request("waterlevelid")


IF Request("txtJenis") = "New" OR Request("txtJenis") = "Header" OR Request("txtJenis") = "Delete" OR Request("txtJenis") = "Edit" THEN
	IF Request("txtJenis") = "Delete" THEN
		q =	" DELETE" & _
            " FROM waterlevelstatus_tab" & _
            " WHERE waterlevelid = '" & waterlevelid & "'"
		connStr.execute(q)
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Header" THEN
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "New" THEN
	
		query_data = " SELECT COUNT(1) in_data" & _
		             " FROM waterlevelstatus_tab" & _
					 " WHERE (REPLACE('" & lowlevel & "', ',', '')" & " BETWEEN lowlevel AND maxlevel)" & _
					 "       OR (REPLACE('" & maxlevel & "', ',', '')" & " BETWEEN lowlevel AND maxlevel)"
		Set rs_data = Server.CreateObject("ADODB.Recordset")
		rs_data.Open query_data, connstr
		if not rs_data.EOF then
			is_in_data = Cdbl(rs_data("in_data"))
		else
			is_in_data = 0
		end if
		rs_data.Close
		Set rs_data = Nothing
		
		if is_in_data > 0 then
			err.number = 1001
			Err.Description = "Data sudah ada, tidak boleh ada range yang sama."
		else
			query_data = " SELECT NVL(MAX(TO_NUMBER(waterlevelid)), 0) maxid FROM waterlevelstatus_tab"
			Set rs_data = Server.CreateObject("ADODB.Recordset")
			rs_data.Open query_data, connstr
			if not rs_data.EOF then
				waterlevelid = Cdbl(rs_data("maxid"))
			else
				waterlevelid = 0
			end if
			rs_data.Close
			Set rs_data = Nothing
			
			waterlevelid = waterlevelid + 1
			
			q = " INSERT INTO waterlevelstatus_tab ( waterlevelid," & _
				"                                    lowlevel," & _
				"                                    maxlevel," & _
				"                                    status," & _
				"                                    fcentry," & _
				"                                    fcedit," & _
				"                                    fcip," & _
				"                                    lastupdate," & _
				"                                    lasttime)" & _
				" VALUES ( '" & waterlevelid & "'," & _
				"          REPLACE('" & lowlevel & "', ',', '')," & _
				"          REPLACE('" & maxlevel & "', ',', '')," & _
				"          '" & status & "'," & _
				"          '" & Session("CodeName") & "'," & _
				"          '" & Session("CodeName") & "'," & _
				"          '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
				"          SYSDATE," & _
				"          TO_CHAR(SYSDATE, 'hh24:mi'))"
			connStr.execute(q)
		end if
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Edit" AND Request("Add") = "Update" THEN
		query_data = " SELECT COUNT(1) in_data" & _
		             " FROM waterlevelstatus_tab" & _
					 " WHERE waterlevelid <> '" & waterlevelid & "' AND ((REPLACE('" & lowlevel & "', ',', '')" & " BETWEEN lowlevel AND maxlevel)" & _
					 "       OR (REPLACE('" & maxlevel & "', ',', '')" & " BETWEEN lowlevel AND maxlevel))"
		Set rs_data = Server.CreateObject("ADODB.Recordset")
		rs_data.Open query_data, connstr
		if not rs_data.EOF then
			is_in_data = Cdbl(rs_data("in_data"))
		else
			is_in_data = 0
		end if
		rs_data.Close
		Set rs_data = Nothing
		
		if is_in_data > 0 then
			err.number = 1001
			Err.Description = "Data sudah ada, tidak boleh ada range yang sama."
		else
			q =	" UPDATE waterlevelstatus_tab" & _
				" SET lowlevel = '" & lowlevel & "'" & _
				"     , maxlevel = '" & maxlevel & "'" & _
				"     , status = '" & status & "'" & _
				" WHERE waterlevelid = '" & waterlevelid & "'"
			connStr.execute(q)
		end if
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
<form name="frmEntry" method="post" action="waterlevelstatus.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Search Water Level</td>
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
	q = " SELECT waterlevelid," & _
	    "        lowlevel," & _
	    "        maxlevel," & _
	    "        status," & _
	    "        fcentry," & _
	    "        fcedit," & _
	    "        fcip," & _
	    "        lastupdate," & _
	    "        lasttime" & _
	    " FROM waterlevelstatus_tab"
	if "" & lookupcode <> "" then
	    q = q & " WHERE '" & lookupcode & "' BETWEEN lowlevel AND maxlevel"
	end if
		q = q & " ORDER BY waterlevelid"
	'response.Write q	
	Set rs_datadtl = Server.CreateObject("ADODB.Recordset")
	rs_datadtl.CursorType = 3
	rs_datadtl.CursorLocation = 3
	rs_datadtl.Open q, connstr
	
%>
   
<form action="waterlevelstatus.asp" method="post" name="eForm">
	<table class="data" width="100%">
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Low Level</font>
			</td>
			<td align="left">
				<input type="text" maxlength = "50" size="30" name="lowlevel" value="<%=lowlevel%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;" required>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Max Level</font>
			</td>
			<td align="left">
				<input type="text" maxlength = "50" size="30" name="maxlevel" value="<%=maxlevel%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;" required>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Status</font>
			</td>
			<td align="left">
				<input type="text" name="status" maxlength = "25" value="<%=status%>" size="50" required>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		
	    <TR>
			<td colspan="2" align="left">
				<input type="hidden" name="lookupcode" value="<%=lookupcode%>">
				<input type="hidden" name="waterlevelid" value="<%=waterlevelid%>">
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
		<th>LOW LEVEL</th>
		<th>MAX LEVEL</th>
		<th>STATUS</th>	
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
				<a href="waterlevelstatus.asp?lookupcode=<%=lookupcode%>&waterlevelid=<%=rs_datadtl("waterlevelid")%>&lowlevel=<%=rs_datadtl("lowlevel")%>&maxlevel=<%=rs_datadtl("maxlevel")%>&status=<%=rs_datadtl("status")%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rs_datadtl(3)%>" border="0">
				</a>
			</td>
			<td>
				<a href="waterlevelstatus.asp?lookupcode=<%=lookupcode%>&waterlevelid=<%=rs_datadtl("waterlevelid")%>&lowlevel=<%=rs_datadtl("lowlevel")%>&maxlevel=<%=rs_datadtl("maxlevel")%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rs_datadtl(3)%>" border="0">
				</a>
			</td>
		<td align="left" title="<%=rs_datadtl("lowlevel").Name%>" valign="top"><%=rs_datadtl("lowlevel")%></td>
		<td align="left" title="<%=rs_datadtl("maxlevel").Name%>" valign="top"><%=rs_datadtl("maxlevel")%></td>
		<td align="left" title="<%=rs_datadtl("status").Name%>" valign="top"><%=rs_datadtl("status")%></td>
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
