<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Authorize Template</title>
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

function CheckIsActive()
{
 var cIsActive = document.getElementById('ISACTIVE');
 if (cIsActive.checked != true )
	{
		alert("Is Active Belum di Click");
	}
}

function calcParam()
{
	var cIsActive = document.getElementById('ISACTIVE').checked;
	if(cIsActive)
	{
		document.getElementById("hide_ISACTIVE").value ="1";
	}else
	{
		document.getElementById("hide_ISACTIVE").value ="0";
	}
	
	var cIsBlock = document.getElementById('ISBLOCK').checked;
	if(cIsBlock)
	{
		document.getElementById("hide_ISBLOCK").value ="1";
	}else
	{
		document.getElementById("hide_ISBLOCK").value ="0";
	}
		
}

function loaddepreciationrange(){
	var method = document.getElementById("depreciationmethod").value;
	if (method === 'STRAIGHT LINE') {
		removeOptions(document.getElementById('depreciationrange'));
		populateSelect('depreciationrange', 1, 20);
	} else {
		removeOptions(document.getElementById('depreciationrange'));
		populateSelect('depreciationrange', 0, 0);
	}
}

function populateSelect(target, min, max){
	if (!target){
		return false;
	}
	else {
		var min = min || 0,
			max = max || min + 0;

		select = document.getElementById(target);

		for (var i = min; i<=max; i++){
			var opt = document.createElement('option');
			opt.innerHTML = "";
			opt.value = i;
			opt.innerHTML = i;
			select.appendChild(opt);
		}
	}
}

function removeOptions(selectbox)
{
    var i;
    for(i = selectbox.options.length - 1 ; i >= 0 ; i--)
    {
        selectbox.remove(i);
    }
}
//using the function:

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
searchdata = request("searchdata")
fccode = request("fccode")
fcname = request("fcname")
document_type = request("document_type")
document_classification = request("document_classification")
authorizebyamount = request("authorizebyamount")
amountrangefrom = request("amountrangefrom")
amountrangeuntil = request("amountrangeuntil")
description = request("description")
isactive = request("isactive")

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM authorizetemplate WHERE fccode = '" & fccode & "'"
		connStr.execute(q)
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		q = " INSERT INTO authorizetemplate ( fccode" & _
		    "                                 , fcname" & _
		    "                                 , document_type" & _
		    "                                 , document_classification" & _
		    "                                 , authorizebyamount" & _
		    "                                 , amountrangefrom" & _
		    "                                 , amountrangeuntil" & _
		    "                                 , description" & _
		    "                                 , isactive" & _
		    "                                 , fcentry" & _
		    "                                 , fcedit" & _
		    "                                 , fcip" & _
		    "                                 , lastupdate" & _
		    "                                 , lasttime)" & _
		    " VALUES ( '" & fccode & "'" & _
			"          , '" & fcname & "'" & _
			"          , '" & document_type & "'" & _
			"          , '" & document_classification & "'" & _
			"          , '" & authorizebyamount & "'" & _
			"          , REPLACE('" & amountrangefrom & "',',','')" & _
			"          , REPLACE('" & amountrangeuntil & "',',','')" & _
			"          , '" & description & "'" & _
			"          , '" & isactive & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

		'response.write q
		connStr.execute(q)
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		q = " UPDATE authorizetemplate" & _
            " SET fcname = '" & fcname & "'" & _
			"     , document_type = '" & document_type & "'" & _
			"     , document_classification = '" & document_classification & "'" & _
			"     , authorizebyamount = '" & authorizebyamount & "'" & _
			"     , amountrangefrom = REPLACE('" & amountrangefrom & "',',','')" & _
			"     , amountrangeuntil = REPLACE('" & amountrangeuntil & "',',','')" & _
			"     , description = '" & description & "'" & _
			"     , isactive = '" & isactive & "'" & _
			"     , fcedit = '" & Session("CodeName") & "'" & _
			"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"     , lastupdate = SYSDATE" & _
			"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
			" WHERE fccode = '" & fccode & "'" 
			
		connStr.execute(q)	
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

 '-------------ISI Payment Form
  %>
  
<form name="frmEntry" method="post" action="authorizetemplate.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
	<TR> 
		<td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE AUTHORIZE TEMPLATE</strong></font></div></td>
	</TR>
	<tr>
		<td align="left" width="20%">Search by Code or Name</td>
		<td align="left"><input type="text" size="30" name="searchdata" value="<%=searchdata%>"  ></td>
	</tr>	
	<tr>
		<td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;
		<a href="Printauthorizetemplate.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
	</tr>
	</table>  
</form>
<%
if cmdView = "True" Then
	q = " SELECT DISTINCT at.*, dt.fcname document_type_name, dc.fcname document_classification_name" & _
	    " FROM authorizetemplate at" & _
		" LEFT JOIN document_type dt ON at.document_type = dt.fccode" & _
		" LEFT JOIN document_classification dc ON at.document_classification = dc.fccode" & _
        " WHERE 1 = 1"

	if ""&searchdata <> "" then q = q & " AND (UPPER(at.fccode) LIKE UPPER('%"&searchdata&"%') OR UPPER(at.fcname) LIKE ('%"&searchdata&"%'))"
	if Request("txtJenis")="Edit" then q = q & " AND at.fccode = '" & fccode & "'"

	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.Open q, connstr
	
	if Request("txtJenis")="Edit" then
		if Not rsGAD.eof then
			fcname = rsGAD("fcname")
			document_type = rsGAD("document_type")
			document_classification = rsGAD("document_classification")
			authorizebyamount = rsGAD("authorizebyamount")
			amountrangefrom = rsGAD("amountrangefrom")
			amountrangeuntil = rsGAD("amountrangeuntil")
			description = rsGAD("description")
			isactive = rsGAD("isactive")
		end if
	end if

%>
   
   <!-- Form Entry And Edit --> 
   
   <form action="authorizetemplate.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
		
		<tr>
			<td width="20%" align="left">Authorize Template Code</td>
			<td align="left">
				<input type="text" size="60" maxlength="50" name="fccode" value="<%=fccode%>" <%IF Request("txtJenis")="Edit" then%> readonly="True" style="background:#ccc"<%end if%>><font color="#FF0000">*</font></td>
		</tr>	
			<td align="left">Authorize Template Name</td>
			<td align="left" width="100%" ><input type="text" size="60" maxlength="150" name="fcname" value="<%=fcname%>"><font color="#FF0000">*</font></td>		
		</tr>
		<tr>
			<td align="left">Module</td>
			<td align="left">
				<select name="document_type">
					<%
					qlov = "SELECT fccode, fcname FROM document_type ORDER BY orderdata"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if document_type = rslov("fccode") then%>selected<%end if%>><%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Classification</td>
			<td align="left">
				<select name="document_classification">
					<%
					qlov = "SELECT fccode, fcname, document_type FROM document_classification ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if cstr(""&document_classification) = cstr(""&rslov("fccode")) then%>selected<%end if%>><%=rslov("document_type")%> : <%=rslov("fccode")%>-<%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Authorize By Amount</td>
			<td align="left">
				<select name="authorizebyamount">
					<option value="TRUE" <%if isactive = "TRUE" then%>selected<%end if%>>TRUE</option>
					<option value="FALSE" <%if isactive = "FALSE" then%>selected<%end if%>>FALSE</option>
				</select>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Amount Range From</td>
			<td align="left">
				<input type="text" size="20" name="amountrangefrom" value="<%=amountrangefrom%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
			</td>
		</tr>
		
		<tr>
			<td align="left">Amount Range Until</td>
			<td align="left">
				<input type="text" size="20" name="amountrangeuntil" value="<%=amountrangeuntil%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
			</td>
		</tr>
		
		<tr>	
			<td align="left">Description</td>
			<td align="left" width="100%" ><input type="text" size="50" maxlength="100" name="description" value="<%=description%>"></td>		
		</tr>
		
		<tr>
			<td align="left">Is Active</td>
			<td align="left">
				<select name="isactive">
					<option value="TRUE" <%if isactive = "TRUE" then%>selected<%end if%>>TRUE</option>
					<option value="FALSE" <%if isactive = "FALSE" then%>selected<%end if%>>FALSE</option>
				</select>
				<font color="#FF0000">*</font>
			</td>
		</tr>
	
		<tr>
			<td colspan="2" align="left">
				  <input type="Hidden" name="searchdata" value="<%=searchdata%>">
				
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
   <table class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<TD align="left">No</TD>
		<TD align="center" colspan="2"></TD>
		<TD align="left">Code</TD>
		<TD align="left">Name</TD>
		<TD align="left">Module</TD>
		<TD align="left">Classification</TD>
		<TD align="left">Authorize By Amount</TD>
		<TD align="left">Amount Range From</TD>
		<TD align="left">Amount Range Until</TD>
		<TD align="left">Description</TD>
		<TD align="left">Is Active</TD>
	</tr>
<%
	if Not rsGAD.eof then
		NO = 0
		While Not rsGAD.EOF
			NO = NO +1
			if no  mod 2 = 1 then cColor = "#EEEEEE" else cColor = "#E5E5E5"
%>
			<tr bgcolor="<%=cColor%>">
				<td align="center"><%=NO%></td>
				<td valign="middle">
					<a href="authorizetemplate.asp?fccode=<%=rsGAD("fccode")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD("fccode")%>" border="0"></a>		
				</td>
				<td valign="middle">
					<a href="authorizetemplate.asp?fccode=<%=rsGAD("fccode")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD("fccode")%>" border="0">
					</a>
				</td>
		   
				<td align="left" title="<%=rsGAD("fccode").Name%>"><%=rsGAD("fccode")%></td>
				<td align="left" title="<%=rsGAD("fcname").Name%>"><%=rsGAD("fcname")%></td>
				<td align="left" title="<%=rsGAD("document_type_name").Name%>"><%=rsGAD("document_type_name")%></td>
				<td align="left" title="<%=rsGAD("document_classification_name").Name%>"><%=rsGAD("document_classification_name")%></td>
				<td align="left" title="<%=rsGAD("authorizebyamount").Name%>"><%=rsGAD("authorizebyamount")%></td>
				<td align="left" title="<%=rsGAD("amountrangefrom").Name%>"><%=rsGAD("amountrangefrom")%></td>
				<td align="left" title="<%=rsGAD("amountrangeuntil").Name%>"><%=rsGAD("amountrangeuntil")%></td>
				<td align="left" title="<%=rsGAD("description").Name%>"><%=rsGAD("description")%></td>
				<td align="left" title="<%=rsGAD("isactive").Name%>"><%=rsGAD("isactive")%></td>

			</tr>
<%
			rsGAD.MoveNext
		Wend
	end if 
	rsGAD.Close : set rsGAD = Nothing
%>
   </div>
   </TABLE>
<%End if%>  
</div></div>
  <!--#include file="bottom.jsp"-->
	<script src="/js/jquery.inputmask.bundle.min.js"></script>
	<script src="/js/inputmask.binding.js"></script>
</body>
</html>
