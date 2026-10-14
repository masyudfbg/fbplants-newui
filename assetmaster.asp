<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<html>
<head>
<title>Asset Master | ePlaS</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
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
searchdata = request("searchdata")
fccode = request("fccode")
fcname = request("fcname")
unitofmeasurement = request("unitofmeasurement")
assetgroupcode = request("assetgroupcode")
description = request("description")
isactive = request("isactive")
assettaxcode = request("assettaxcode")

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM asset WHERE fccode = '" & fccode & "'"
		connStr.execute(q)
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		q = " INSERT INTO asset ( fccode" & _
		    "                          , fcname" & _
			"                          , assetgroupcode" & _
			"                          , assettaxcode" & _
		    "                          , unitofmeasurement" & _
		    "                          , description" & _
		    "                          , isactive" & _
		    "                          , fcentry" & _
		    "                          , fcedit" & _
		    "                          , fcip" & _
		    "                          , lastupdate" & _
		    "                          , lasttime)" & _
		    " VALUES ( '" & fccode & "'" & _
			"          , '" & fcname & "'" & _
			"          , '" & assetgroupcode & "'" & _
			"          , '" & assettaxcode & "'" & _
			"          , '" & unitofmeasurement & "'" & _
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
		q = " UPDATE asset" & _
            " SET fcname = '" & fcname & "'" & _
			"     , unitofmeasurement = '" & unitofmeasurement & "'" & _
			"     , assetgroupcode = '" & assetgroupcode & "'" & _
			"     , assettaxcode = '" & assettaxcode & "'" & _
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
  
<form name="frmEntry" method="post" action="assetmaster.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
	<TR> 
		<td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE ASSET MASTER</strong></font></div></td>
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
		<a href="Printassetmaster.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txt_CUSTOMERCODE=<%=REQUEST("txt_CUSTOMERCODE")%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
	</tr>
	</table>  
</form>
<%
if cmdView = "True" Then
	q = " SELECT am.*" & _
	    "        , DECODE(am.isactive,1,'TRUE', 'FALSE') isactivedesc" & _
	    "        , ag.fcname assetgroupname, at.fcname assettaxname" & _
	    " FROM asset am LEFT JOIN assetgroup ag ON ag.fccode = am.assetgroupcode " & _
	    "  LEFT JOIN fixedasset_taxgroup at ON at.fccode = am.assettaxcode" & _
        " WHERE 1=1"

	if ""&searchdata <> "" then q = q & " AND (UPPER(am.fccode) LIKE UPPER('%"&searchdata&"%') OR UPPER(am.fcname) LIKE ('%"&searchdata&"%'))"
	if Request("txtJenis")="Edit" then q = q & " AND am.fccode = '" & fccode & "'"
	' response.write q
	Set rsGAD = Server.CreateObject("ADODB.Recordset")
	rsGAD.Open q, connstr
	
	if Request("txtJenis")="Edit" then
		if Not rsGAD.eof then
			fcname = rsGAD("fcname")
			unitofmeasurement = rsGAD("unitofmeasurement")
			assetgroupcode = rsGAD("assetgroupcode")
			assettaxcode = rsGAD("assettaxcode")
			description = rsGAD("description")
			isactive = rsGAD("isactive")
		end if
	end if

%>
   
   <!-- Form Entry And Edit --> 
   
   <form action="assetmaster.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
		
		<tr>
			<td width="20%" align="left">Asset Master Code</td>
			<td align="left">
				<input type="text" size="15" maxlength="30" name="fccode" value="<%=fccode%>" <%IF Request("txtJenis")="Edit" then%> readonly="True" style="background:#999999"<%end if%>><font color="#FF0000">*</font></td>
		</tr>	
			<td align="left">Asset Master Name</td>
			<td align="left" width="100%" ><input type="text" size="50" maxlength="100" name="fcname" value="<%=fcname%>"><font color="#FF0000">*</font></td>		
		</tr>
		
		<tr>
			<td align="left">Asset Group</td>
			<td align="left">
				<select name="assetgroupcode" class="chosen-select" style="width: 50%">
					<option></option>
					<%
					qlov = "SELECT fccode, fcname FROM assetgroup ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if cstr(""&assetgroupcode) = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
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
			<td align="left">Asset Tax Group</td>
			<td align="left">
				<select name="assettaxcode" class="chosen-select" style="width: 50%">
					<option></option>
					<%
					qlov = "SELECT fccode, fcname FROM fixedasset_taxgroup ORDER BY fccode"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if cstr(""&assettaxcode) = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode") & ": " & rslov("fcname")%></option>
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
			<td align="left">Uom</td>
			<td align="left">
				<select name="unitofmeasurement" class="chosen-select" style="width: 30%">
					<option></option>
					<%
					qlov = "SELECT unitofmeasurecode fccode, unitofmeasuredesc fcname FROM unitofmeasure ORDER BY unitofmeasuredesc"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if cstr(""&unitofmeasurement) = rslov("fccode") then%>selected<%end if%>><%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		</tr>	
			<td align="left">Description</td>
			<td align="left" width="100%">
				<textarea name="description"><%=description%></textarea>
			</td>		
		</tr>
		<tr>
			<td align="left">Is Active</td>
			<td align="left">
				<select name="isactive" class="chosen-select" style="width: 10%">
					<option value="1" <%if isactive = "1" then%>selected<%end if%>>TRUE</option>
					<option value="0" <%if isactive = "0" then%>selected<%end if%>>FALSE</option>
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
		<TD align="left">Asset Master Code</TD>
		<TD align="left">Asset Master Name</TD>
		<TD align="left">Asset Group</TD>
		<TD align="left">Asset Tax Group</TD>
		<TD align="left">Uom</TD>
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
					<a href="assetmaster.asp?fccode=<%=rsGAD("fccode")%>&searchdata=<%=searchdata%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD("fccode")%>" border="0"></a>		
				</td>
				<td valign="middle">
					<a href="assetmaster.asp?fccode=<%=rsGAD("fccode")%>&searchdata=<%=searchdata%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');">
						<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD("fccode")%>" border="0">
					</a>
				</td>
		   
				<td align="left" title="<%=rsGAD("fccode").Name%>"><%=rsGAD("fccode")%></td>
				<td align="left" title="<%=rsGAD("fcname").Name%>"><%=rsGAD("fcname")%></td>
				<td align="left" title="<%=rsGAD("assetgroupcode").Name%>"><%=rsGAD("assetgroupcode") & " - " & rsGAD("assetgroupname")%></td>
				<td align="left" title="<%=rsGAD("assettaxcode").Name%>"><%=rsGAD("assettaxcode") & " - " & rsGAD("assettaxname")%></td>
				<td align="left" title="<%=rsGAD("unitofmeasurement").Name%>"><%=rsGAD("unitofmeasurement")%></td>
				<td align="left" title="<%=rsGAD("description").Name%>"><%=rsGAD("description")%></td>
				<td align="left" title="<%=rsGAD("isactivedesc").Name%>"><%=rsGAD("isactivedesc")%></td>

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
</body>
</html>
