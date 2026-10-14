<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<title>Asset Group | ePlaS</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">
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
	// alert(method);
	if (method === 'STRAIGHT LINE') {
		$("#depreciationrange").show();
		$("#depreciationrange").val("");
		// $("#notdepreciable").hide();
		// removeOptions(document.getElementById('depreciationrange'));
		// populateSelect('depreciationrange', 1, 20);
	} else {
		$("#depreciationrange").val("NOT DEPRECIABLE");
		$("#depreciationrate").val(0);
		// $("#notdepreciable").show();
		// removeOptions(document.getElementById('depreciationrange'));
		// populateSelect('depreciationrange', 0, 0);
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

 		function hitungrumus(){
    	var depreciationrange = document.getElementById("depreciationrange");
    	var depreciationrate = document.getElementById("depreciationrate");
    	// var akp = document.getElementById("akp");
    	// var productivestand = document.getElementById("productivestand");
    	// var kg_taksasi = document.getElementById("kg_taksasi");
    	// var bjr_bbc = document.getElementById("bjr_bbc");
    	// var output_harvesting = document.getElementById("output_harvesting");
    	// var hk_harvesting = document.getElementById("hk_harvesting");

    	depreciationrate.value = Math.round( (100 / depreciationrange.value) );
    	// kg_taksasi.value = Math.round( (janjang_matang.value / field_sampling.value) * (100) * (productivestand.value * bjr_bbc.value) );
    	// hk_harvesting.value = Math.round( (kg_taksasi.value / output_harvesting.value) );
    	}



</script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-table.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-table.css?v=20261010c" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<script type="text/javascript" src="js/fbp-table.js?v=20261009c"></script>
<style>
/* assetgroup.asp: card & tabel selebar konten, judul kolom boleh turun baris */
body.page-legacy .box_both > form > table.data { width: 100% !important; }
body.page-legacy #assetGroupTable { width: 100% !important; }
body.page-legacy #assetGroupTable thead tr:not(.fbt-toolbar-row) > td { white-space: normal; vertical-align: bottom; line-height: 1.3; }
body.page-legacy #assetGroupTable tbody td { text-align: left; vertical-align: top; }
body.page-legacy #assetGroupTable tbody td:first-child { text-align: center; }
body.page-legacy #assetGroupTable tbody td:nth-child(n+9):nth-child(-n+11) { min-width: 150px; }
</style>
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
assettypecode = request("assettypecode")
depreciationmethod = request("depreciationmethod")
depreciationrange = request("depreciationrange")
depreciationrate = request("depreciationrate")
controljob = request("controljob")
description = request("description")
controljob_depreciation = request("controljob_depreciation")
controljob_accd_depreciation = request("controljob_accd_depreciation")

' if len(trim(depreciationmethod)) = 0 then depreciationmethod = "NOT DEPRECIABLE"

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM assetgroup WHERE fccode = '" & fccode & "'"
		connStr.execute(q)
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		if Request("depreciationmethod") = "NOT DEPRECIABLE"  then
			q = " INSERT INTO assetgroup ( fccode" & _
			    "                          , fcname" & _
			    "                          , assettypecode" & _
			    "                          , depreciationmethod" & _
			    "                          , depreciationrange" & _
			    "                          , controljob" & _
			    "                          , controljob_depreciation" & _
			    "                          , controljob_accd_depreciation" & _
			    "                          , description" & _
			    "                          , fcentry" & _
			    "                          , fcedit" & _
			    "                          , fcip" & _
			    "                          , lastupdate" & _
			    "                          , lasttime)" & _
			    " VALUES ( '" & fccode & "'" & _
				"          , '" & fcname & "'" & _
				"          , '" & assettypecode & "'" & _
				"          , '"& depreciationmethod &"'" & _
				"          , '0'" & _
				"          , '" & controljob & "'" & _
				"          , '" & controljob_depreciation & "'" & _
				"          , '" & controljob_accd_depreciation & "'" & _
				"          , '" & description & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Session("CodeName") & "'" & _
				"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"          , SYSDATE" & _
				"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

			' response.write q
			connStr.execute(q)
		else
			q = " INSERT INTO assetgroup ( fccode" & _
		    "                          , fcname" & _
		    "                          , assettypecode" & _
		    "                          , depreciationmethod" & _
		    "                          , depreciationrange" & _
		    "                          , controljob" & _
		    "                          , controljob_depreciation" & _
			"                          , controljob_accd_depreciation" & _
		    "                          , description" & _
		    "                          , fcentry" & _
		    "                          , fcedit" & _
		    "                          , fcip" & _
		    "                          , lastupdate" & _
		    "                          , lasttime)" & _
		    " VALUES ( '" & fccode & "'" & _
			"          , '" & fcname & "'" & _
			"          , '" & assettypecode & "'" & _
			"          , '" & depreciationmethod & "'" & _
			"          , '" & depreciationrange & "'" & _
			"          , '" & controljob & "'" & _
			"          , '" & controljob_depreciation & "'" & _
			"          , '" & controljob_accd_depreciation & "'" & _
			"          , '" & description & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Session("CodeName") & "'" & _
			"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
			"          , SYSDATE" & _
			"          , TO_CHAR(SYSDATE, 'hh24:mi'))"			 

		' response.write q
		connStr.execute(q)
		End if
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		if Request("depreciationmethod") = "NOT DEPRECIABLE"  then
			q = " UPDATE assetgroup" & _
	            " SET fcname = '" & fcname & "'" & _
				"     , assettypecode = '" & assettypecode & "'" & _
				"     , depreciationmethod = '" & depreciationmethod & "'" & _
				"     , depreciationrange = '0'" & _
				"     , controljob = '" & controljob & "'" & _
				"     , controljob_depreciation = '" & controljob_depreciation & "'" & _
				"     , controljob_accd_depreciation = '" & controljob_accd_depreciation & "'" & _
				"     , description = '" & description & "'" & _
				"     , fcedit = '" & Session("CodeName") & "'" & _
				"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"     , lastupdate = SYSDATE" & _
				"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
				" WHERE fccode = '" & fccode & "'" 
				
			connStr.execute(q)	
		Else
			q = " UPDATE assetgroup" & _
	            " SET fcname = '" & fcname & "'" & _
				"     , assettypecode = '" & assettypecode & "'" & _
				"     , depreciationmethod = '" & depreciationmethod & "'" & _
				"     , depreciationrange = '" & depreciationrange & "'" & _
				"     , controljob = '" & controljob & "'" & _
				"     , controljob_depreciation = '" & controljob_depreciation & "'" & _
				"     , controljob_accd_depreciation = '" & controljob_accd_depreciation & "'" & _
				"     , description = '" & description & "'" & _
				"     , fcedit = '" & Session("CodeName") & "'" & _
				"     , fcip = '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
				"     , lastupdate = SYSDATE" & _
				"     , lasttime = TO_CHAR(SYSDATE, 'hh24:mi')" & _
				" WHERE fccode = '" & fccode & "'" 
				
			connStr.execute(q)	
		End If
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
  
<form name="frmEntry" method="post" action="assetgroup.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
	<TR> 
		<td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE ASSET GROUP</strong></font></div></td>
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
		<a href="Printassetgroup.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
	</tr>
	</table>  
</form>
<%
if cmdView = "True" Then
	q = " SELECT ag.*, job.fcname jobname FROM assetgroup ag LEFT JOIN job ON (ag.controljob = job.fccode AND job.fcba = '" & Session("fcBA") & "')" & _
        " WHERE 1=1"

	if ""&searchdata <> "" then q = q & " AND (UPPER(ag.fccode) LIKE UPPER('%"&searchdata&"%') OR UPPER(ag.fcname) LIKE ('%"&searchdata&"%'))"
	if Request("txtJenis")="Edit" then q = q & " AND ag.fccode = '" & fccode & "'"

	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.Open q, connstr
	
	if Request("txtJenis")="Edit" then
		if Not rsGAD.eof then
			fcname = rsGAD("fcname")
			jobname = rsGAD("jobname")
			assettypecode = rsGAD("assettypecode")
			depreciationmethod = rsGAD("depreciationmethod")
			depreciationrange = rsGAD("depreciationrange")
			controljob = rsGAD("controljob")
			controljob_depreciation = rsGAD("controljob_depreciation")
			controljob_accd_depreciation = rsGAD("controljob_accd_depreciation")
			description = rsGAD("description")
		end if
	end if

%>
   
   <!-- Form Entry And Edit --> 
   
   <form action="assetgroup.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
		
		<tr>
			<td width="20%" align="left">Asset Group Code</td>
			<td align="left">
				<input type="text" size="15" maxlength="30" name="fccode" value="<%=fccode%>" <%IF Request("txtJenis")="Edit" then%> readonly="True" style="background:#999999"<%end if%>><font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="left">Asset Group Name</td>
			<td align="left" width="100%" ><input type="text" size="50" maxlength="100" name="fcname" value="<%=fcname%>"><font color="#FF0000">*</font></td>		
		</tr>
		<tr>
			<td align="left">Asset Type Code</td>
			<td align="left">
				<select name="assettypecode" class="chosen-select" style="width: 50%">
					<option></option>
					<option value="CAPEX" <%if assettypecode = "CAPEX" then%>selected<%end if%>>CAPEX</option>
					<option value="LVA" <%if assettypecode = "LVA" then%>selected<%end if%>>LVA</option>
				</select>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Commercial Depreciation Method</td>
			<td align="left">
				<select id="depreciationmethod" name="depreciationmethod" onchange="loaddepreciationrange();" class="chosen-select" style="width: 50%">
					<option></option>
					<option value="NOT DEPRECIABLE" <%if depreciationmethod = "NOT DEPRECIABLE" then%>selected<%end if%>>NOT DEPRECIABLE</option>
					<option value="STRAIGHT LINE" <%if depreciationmethod = "STRAIGHT LINE" then%>selected<%end if%>>STRAIGHT LINE</option>
				</select>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Commercial Depreciation Range (Month)</td>
			<td align="left">
				<input type="text" size="25" name="depreciationrange" id="depreciationrange" value="<%=depreciationrange%>" onkeyup="harusangka(this); hitungrumus(); ">
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Commercial Depreciation Rate (%)</td>
			<td align="left">
				<input type="text" size="15" id="depreciationrate" name="depreciationrate" value="<%=depreciationrate%>" readonly style = "background-color: #eee">
				<font color="#FF0000">*</font>
			</td>
		</tr>
	
		<tr>
			<td align="left">Control Job</td>
			<td align="left">
				<select name="controljob" class="chosen-select" style="width: 50%">
					<option></option>
					<%
					qlov = "SELECT DISTINCT fccode, fcname FROM job ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if controljob = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
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
			<td align="left">Control Job Depreciation</td>
			<td align="left">
				<select name="controljob_depreciation" class="chosen-select" style="width: 50%">
					<option></option>
					<%
					qlov = "SELECT DISTINCT fccode, fcname FROM job ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if controljob = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
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
			<td align="left">Control Job Accumulated Depreciation</td>
			<td align="left">
				<select name="controljob_accd_depreciation" class="chosen-select" style="width: 50%">
					<option></option>
					<%
					qlov = "SELECT DISTINCT fccode, fcname FROM job ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if controljob = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
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
			<td align="left">Description</td>
			<td align="left" width="100%" ><input type="text" size="50" maxlength="100" name="description" value="<%=description%>"></td>		
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
   <table class="data" id="assetGroupTable">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		<TD align="left">No</TD>
		<TD align="center"></TD>
		<TD align="center"></TD>
		<TD align="center">Asset Group Code</TD>
		<TD align="center">Asset Group Name</TD>
		<TD align="center">Asset Type</TD>
		<TD align="center">Depreciation Method</TD>
		<TD align="center">Depreciation Range</TD>
		<TD align="center">Control Job</TD>
		<TD align="center">Control Job Depreciation</TD>
		<TD align="center">Control Job Accumulated Depreciation</TD>
		<TD align="center">Description</TD>
	</tr>
	</thead>
	<tbody>
<%
	if rsGAD.eof then
%>
			<tr><td colspan="12" align="center">Data asset group tidak ditemukan.</td></tr>
<%
	end if
	if Not rsGAD.eof then
		NO = 0
		While Not rsGAD.EOF
			NO = NO +1
			if no  mod 2 = 1 then cColor = "#EEEEEE" else cColor = "#E5E5E5"
%>
			<tr bgcolor="<%=cColor%>">
				<td align="center"><%=NO%></td>
				<td valign="middle">
					<a href="assetgroup.asp?fccode=<%=rsGAD("fccode")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD("fccode")%>" border="0"></a>		
				</td>
				<td valign="middle">
					<a href="assetgroup.asp?fccode=<%=rsGAD("fccode")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD("fccode")%>" border="0">
					</a>
				</td>
		   
				<td align="left" title="<%=rsGAD("fccode").Name%>"><%=rsGAD("fccode")%></td>
				<td align="left" title="<%=rsGAD("fcname").Name%>"><%=rsGAD("fcname")%></td>
				<td align="left" title="<%=rsGAD("assettypecode").Name%>"><%=rsGAD("assettypecode")%></td>
				<td align="left" title="<%=rsGAD("depreciationmethod").Name%>"><%=rsGAD("depreciationmethod")%></td>
				<td align="left" title="<%=rsGAD("depreciationrange").Name%>"><%=rsGAD("depreciationrange")%></td>
				<td align="left" title="<%=rsGAD("controljob").Name%>"><%=rsGAD("controljob") & " - " & rsGAD("jobname")%></td>
				<td align="left" title="<%=rsGAD("controljob_depreciation").Name%>"><%=rsGAD("controljob_depreciation") & " - " & rsGAD("jobname") %></td>
				<td align="left" title="<%=rsGAD("controljob_accd_depreciation").Name%>"><%=rsGAD("controljob_accd_depreciation") & " - " & rsGAD("jobname")%></td>
				<td align="left" title="<%=rsGAD("description").Name%>"><%=rsGAD("description")%></td>

			</tr>
<%
			rsGAD.MoveNext
		Wend
	end if 
	rsGAD.Close : set rsGAD = Nothing
%>
   </tbody>
   </TABLE>
<% if Request("oldui") <> "1" then %>
   <script type="text/javascript">
	// Tabel daftar asset group: cari, jumlah baris, urut kolom, pagination (tampilan saja)
	if (window.FbpTable) FbpTable.init(document.getElementById('assetGroupTable'), { pageSize: 10 });
   </script>
<% end if %>
<%End if%>  
</div></div>
  <!--#include file="bottom.jsp"-->
  <script src="/js/jquery.inputmask.bundle.min.js"></script>
  <script src="/js/inputmask.binding.js"></script>
  <script src="plugins/chosen/chosen.jquery.min.js"></script>
  <script type="text/javascript">	
		// $(".chosen-select").empty();
		
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"45%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		// $(".chosen-select").trigger("chosen:updated");
  </script>
  <script type="text/javascript">
  	document.ready(function() {
	// getOvertime();
	// $("#foo").hide();
	//  if ($("#allowded").val() != "") {
	//         $("#foo").show();
	//     }else{
	//         $("#foo").hide();
	//     } 
	// document.getElementById("harvestingquality").style.display="none";
	loaddepreciationrange();
		});
  </script>
</body>
</html>
