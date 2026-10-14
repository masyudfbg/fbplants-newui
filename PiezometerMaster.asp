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
<title>ePlaS - Piezometer Master</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<style>
	.cut_input{margin-left:-10px;margin-top:20px;}
</style>
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

<script language="JavaScript"> 
function delete_element(id){
	Element.prototype.remove = function() {
		this.parentElement.removeChild(this);
	}
	NodeList.prototype.remove = HTMLCollection.prototype.remove = function() {
		for(var i = this.length - 1; i >= 0; i--) {
			if(this[i] && this[i].parentElement) {
				this[i].parentElement.removeChild(this[i]);
			}
		}
	}
	document.getElementById(id).remove();
}
function replace_comma(e){
	var value_str = e.value;
	var new_str = value_str.replace(/,/g, "");
	e.value = new_str;
}
function plusblock(){
	var jml = $(".plusblockoption").length;
	var counter = document.getElementById("countplusblockoption");
	var val_counter = counter.value;
	var x = 1;
	num = x + parseInt(counter.value)
	option = "'option"+num+"'";
	$("#plusblock").append('<div id="option'+num+'" style="margin-top:5px;"><input class="plusblockoption" name="fieldcode" value="" type="text" placeholder="Block Covered Area">&nbsp;<a class="cut_input" href="#" onclick="delete_element(' + option + ');"><img src="IMAGES/no_check.gif" alt="Delete" border="0"></a></div>');
	counter.value = num;
	//<input name="remarkdtl" value="" type="text" onblur="replace_comma(this)">
}
/*
function ConvertDMSToDD(degrees, minutes, seconds, direction) {
    var dd = degrees + minutes/60 + seconds/(60*60);
    if (direction == "X" || direction == "Y") {
        dd = dd * -1;
    } // Don't do anything for N or E
	alert(dd);
    return dd;
}

function ParseDMS(e,dir) {
	var input = e.value;
    var parts = input.split(/[^\d\w]+/);
    var cor = ConvertDMSToDD(parts[0], parts[1], parts[2], dir);
	var id_cor = document.getElementById("cor"+dir);
	//alert(parts[2]);
	id_cor.value = cor;
}
*/
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
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis, cmdView
txtTable = Request("txtTable")
txtModule = Request("txtModule")
txtSearch = Request("txtSearch")
txtField = Request("txtField")
txtJenis = request("txtJenis")


Dim rs_data, query_data
Dim qdelete,qdtl,rs_data_dtl
Dim search,piezoid,piezotype,longitude,latitude,remarks,fcba,fieldcode,isactive,fieldcode_split
Dim comm, param, fcba_, lookupcode_, date_from_, Date_until_, fcentry_, fcedit_, fcip_, ResultCommand_
'Dim remarkdtl_split,remarkdtl

search 		= trim(Request("search"))
piezoid 	= trim(Request("piezoid"))
piezotype 	= Request("piezotype")
longitude 	= trim(Request("longitude"))
latitude 	= trim(Request("latitude"))
remarks 	= Request("remarks")
fcba 		= Session("fcba") 
fieldcode	= trim(Request("fieldcode"))
isactive 	= Request("isactive")
'remarkdtl	= Request("remarkdtl")
fieldcode_split = Split(fieldcode,",")
'remarkdtl_split = Split(remarkdtl,",")

IF Request("txtJenis") = "New" OR Request("txtJenis") = "Header" OR Request("txtJenis") = "Delete" OR Request("txtJenis") = "Edit" THEN
	IF Request("txtJenis") = "Delete" THEN
		q =	" DELETE" & _
            " FROM PiezoDtl_tab" & _
            " WHERE piezoid = '" & piezoid & "'"
		connStr.execute(q)
		q =	" DELETE" & _
            " FROM piezo_tab" & _
            " WHERE piezoid = '" & piezoid & "'"
		connStr.execute(q)
		'response.write(q)
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Header" THEN
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "New" THEN
		q = " INSERT INTO piezo_tab(piezoid," & _
			"        piezotype," & _
			"        longitude," & _
			"        latitude," & _
			"        remarks," & _
			"        fcba," & _
			"        fieldcode," & _
			"        isactive," & _
			"        fcentry," & _
			"        fcedit," & _
			"        fcip," & _
			"        lastupdate," & _
			"        lasttime" & _
			") VALUES ( '" & piezoid & "'," & _
			"          '" & piezotype & "'," & _
			"          '" & longitude & "'," & _
			"          '" & latitude & "'," & _
			"          '" & remarks & "'," & _
			"          '" & fcba & "'," & _
			"          '" & fieldcode_split(0) & "'," & _
			"          '" & isactive & "'," & _
			"          '" & Session("CodeName") & "'," & _
			"          '" & Session("CodeName") & "'," & _
			"          '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
			"          SYSDATE," & _
			"          TO_CHAR(SYSDATE, 'hh24:mi'))"
		connStr.execute(q)
	if uBound(fieldcode_split)	> 0 then	
		qdelete =	" DELETE" & _
					" FROM PiezoDtl_tab" & _
					" WHERE piezoid = '" & piezoid & "'"
		connStr.execute(qdelete)
		for i=1 to uBound(fieldcode_split)	
			qdtl = " INSERT INTO PiezoDtl_tab(piezoid," & _
				"        fieldcode," & _
				"        fcba," & _
				"        isactive," & _
				"        fcentry," & _
				"        fcedit," & _
				"        fcip," & _
				"        lastupdate," & _
				"        lasttime" & _
				") VALUES ( '" & piezoid & "'," & _
				"          '" & fieldcode_split(i) & "'," & _
				"          '" & fcba & "'," & _
				"          '" & isactive & "'," & _
				"          '" & Session("CodeName") & "'," & _
				"          '" & Session("CodeName") & "'," & _
				"          '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
				"          SYSDATE," & _
				"          TO_CHAR(SYSDATE, 'hh24:mi'))"
				connStr.execute(qdtl)
		Next 
	END IF	
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Edit" AND Request("Add") = "Update" THEN
		q =	" UPDATE piezo_tab" & _
			" SET piezotype = '" & piezotype & "'" & _
			"     , longitude = '" & longitude & "'" & _
			"     , latitude = '" & latitude & "'" & _
			"     , remarks = '" & remarks & "'" & _
			"     , fieldcode = '" & trim(fieldcode_split(0)) & "'" & _
			"     , isactive = '" & isactive & "'" & _
			
			" WHERE piezoid = '" & piezoid & "'"
			'response.write q
		connStr.execute(q)
	if uBound(fieldcode_split)	> 0 then
		qdelete =	" DELETE" & _
					" FROM PiezoDtl_tab" & _
					" WHERE piezoid = '" & piezoid & "'"
		connStr.execute(qdelete)
		for i=1 to uBound(fieldcode_split)	
			qdtl = " INSERT INTO PiezoDtl_tab(piezoid," & _
				"        fieldcode," & _
				"        fcba," & _
				"        isactive," & _
				"        fcentry," & _
				"        fcedit," & _
				"        fcip," & _
				"        lastupdate," & _
				"        lasttime" & _
				") VALUES ( '" & piezoid & "'," & _
				"          '" & trim(fieldcode_split(i)) & "'," & _
				"          '" & fcba & "'," & _
				"          '" & isactive & "'," & _
				"          '" & Session("CodeName") & "'," & _
				"          '" & Session("CodeName") & "'," & _
				"          '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
				"          SYSDATE," & _
				"          TO_CHAR(SYSDATE, 'hh24:mi'))"
			'response.write qdtl & "###"
				connStr.execute(qdtl)
		Next 
	END IF	
		cmdView	= "True"
	ELSE
		cmdView	= "True"
	END IF
	'Response.Write "sql: " & q
ELSE
    cmdView	= request("cmdView")
END IF


if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

 '-------------ISI Payment Form
%>
<form name="frmEntry" method="post" action="PiezometerMaster.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Piezometer Search (id/type/blok)</td>
			<td align="left">
				<input type="text" maxlength = "20" name="search" size="30" value="<%=search%>">
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
	Dim rs_data_tab,intRecordsPerPage,cColor
	q = " select piezoid," & _
	    "        piezotype," & _
	    "        longitude," & _
		"        latitude," & _
		"        remarks," & _
		"        fcba," & _
		"        fieldcode," & _
		"        isactive," & _
	    "        fcentry," & _
	    "        fcedit," & _
	    "        fcip," & _
	    "        lastupdate," & _
	    "        lasttime" & _
	    " from piezo_tab WHERE fcba = '" & Session("fcba") & "'"
	if piezoid <> "" and txtJenis = "Edit" then 
	    q = q & " and piezoid = '" & piezoid & "'"
	elseif search <> "" then 
		q = q & " and (piezoid = '" & search & "' or piezotype = '" & search & "' or fieldcode = '" & search & "')"
	end if
		q = q & " order by piezoid"
	'response.Write q	
	Set rs_data_tab = Server.CreateObject("ADODB.Recordset")
	rs_data_tab.CursorType = 3
	rs_data_tab.CursorLocation = 3
	rs_data_tab.Open q, connstr
	
	if piezoid <> "" and txtJenis = "Edit" then 
		if Not rs_data_tab.eof then
			piezoid 	= rs_data_tab("piezoid")
			piezotype 	= rs_data_tab("piezotype")
			longitude 	= rs_data_tab("longitude")
			latitude 	= rs_data_tab("latitude")
			remarks 	= rs_data_tab("remarks")
			fcba 		= Session("fcba") 
			fieldcode	= rs_data_tab("fieldcode")
			isactive 	= rs_data_tab("isactive")
			qdtl = 	" select piezoid," & _
					"        fieldcode" & _
					" from PiezoDtl_tab WHERE piezoid = '" & piezoid & "' and fcba = '" & Session("fcba") & "'"
			'response.write qdtl
			Set rs_data_dtl = Server.CreateObject("ADODB.Recordset")
			rs_data_dtl.CursorType = 3
			rs_data_dtl.CursorLocation = 3
			rs_data_dtl.Open qdtl, connstr
		end if
	end if
%>
   
<form name="eForm" action="PiezometerMaster.asp" method="post">
	<table class="data" width="100%">
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Piezo ID</font>
			</td>
			<td align="left">
				<%if txtJenis = "Edit" then%><input name="piezoid" type="text" value="<%=piezoid%>" disabled><%end if%>
				<input  name="piezoid" value="<%=piezoid%>" type="<%if txtJenis = "Edit" then%>hidden<%else%>text<%end if%>" required>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Piezo Type</td>
			<td align="left"> 
				<select name="piezotype" required>
					<option value="" selected></option>
					<%
					Dim PiezoType_tab,strsql
					Set PiezoType_tab = Server.CreateObject("ADODB.Recordset")
					strsql =  " SELECT * from PiezoType_tab "
					PiezoType_tab.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
					if not PiezoType_tab.eof then
						DO while not PiezoType_tab.eof
							if piezotype = PiezoType_tab("fccode") then
							%>			
								<option value="<%=PiezoType_tab("fccode")%>" Selected><%=PiezoType_tab("fccode")%></option>
							<% Else %>	
								<option value="<%=PiezoType_tab("fccode")%>"><%=PiezoType_tab("fccode")%></option>
							<% End If
							
							PiezoType_tab.movenext
						loop
					End if
					PiezoType_tab.Close
					set PiezoType_tab = Nothing
					%>
				</select> <font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left" valign="top">
				<font face="Verdana" size="1" >Field Code</font>
				
			</td>
			<td align="left" id="plusblock">
				<input class="plusblockoption" name="fieldcode" value="<%=fieldcode%>" type="text" placeholder="Real Block"> <a class="button_menu_top" onclick="plusblock();"> + </a><br/>
			<% 
			no=0
			if piezoid <> "" and txtJenis = "Edit" then 
				if Not rs_data_dtl.eof then
						
						While Not rs_data_dtl.EOF 
						no = no+1
						%>
						<div id="option<%=no%>" style="margin-top:5px;"><input class="plusblockoption" name="fieldcode" value="<%=rs_data_dtl("fieldcode")%>" type="text" placeholder="Block Covered Area">&nbsp;<a class="cut_input" href="#" onclick="delete_element('option<%=no%>');"><img src="IMAGES/no_check.gif" alt="Delete" border="0"></a></div>
						
					<% rs_data_dtl.movenext
					Wend
						end if 
						rs_data_dtl.Close
						set rs_data_dtl = Nothing
			end if
				%>
				<input Type="hidden" value="<%=no%>" id="countplusblockoption">
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Latitude</font>
			</td>
			<td align="left">
				<input id="lat" type="number" name="latitude" value="<%=latitude%>" step="any" style="text-align: right;" required>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Longitude</font>
			</td>
			<td align="left">
				<input id="lng" type="number" name="longitude" step="any" value="<%=longitude%>" style="text-align: right;" required>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Isactive</td>
			<td align="left"> 
				<select name="isactive" required> 			
					<option value="TRUE" selected>True</option>
					<option value="FALSE" <%if isactive = "FALSE" then%>selected<%end if%>>False</option>
				</select> <font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Remark</font>
			</td>
			<td align="left">
				<input  name="remarks" value="<%=remarks%>" type="text">
			</td>
		</tr>
	    <tr>
			<td colspan="2" align="left">
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
		<th>PIEZO ID</th>
	    <th>PIEZO TYPE</th>
	    <th>LONGITUDE</th>
		<th>LATITUDE</th>
		<th>FIELD CODE</th>
		<th>REMARKS</th>
		<th>ISACTIVE</th>
	</tr>
<%
	'response.write Session("v_Employee")
   
	if Not rs_data_tab.eof then
		nCountofField=rs_data_tab.fields.count-1
		no=1
		While Not rs_data_tab.EOF
			if no  mod 2 = 1 then
				cColor = "#EEEEEE"
			else
				cColor = "#E5E5E5"
			end if
%>
	<tr bgcolor="<%=cColor%>">
		<td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
				<a href="PiezometerMaster.asp?piezoid=<%=rs_data_tab("piezoid")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rs_data_tab(3)%>" border="0">
				</a>
			</td>
			<td>
				<a href="PiezometerMaster.asp?piezoid=<%=rs_data_tab("piezoid")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rs_data_tab(3)%>" border="0">
				</a>
			</td>
	
		<td align="left" title="<%=rs_data_tab("piezoid").Name%>" valign="top"><%=rs_data_tab("piezoid")%></td>
		<td align="left" title="<%=rs_data_tab("piezotype").Name%>" valign="top"><%=rs_data_tab("piezotype")%></td>
		<td align="left" title="<%=rs_data_tab("longitude").Name%>" valign="top"><%=rs_data_tab("longitude")%></td>
		<td align="left" title="<%=rs_data_tab("latitude").Name%>" valign="top"><%=rs_data_tab("latitude")%></td>
		<td align="left" title="<%=rs_data_tab("fieldcode").Name%>" valign="top"><%=rs_data_tab("fieldcode")%></td>
		<td align="left" title="<%=rs_data_tab("remarks").Name%>" valign="top"><%=rs_data_tab("remarks")%></td>
		<td align="left" title="<%=rs_data_tab("isactive").Name%>" valign="top"><font color="red"><%=rs_data_tab("isactive")%></font></td>
	</tr>
<%
			rs_data_tab.MoveNext
			intRecordsPerPage = intRecordsPerPage - 1
			no=no+1
		Wend
	else
		response.write ("<tr bgcolor='#E5E5E5'><td colspan='9' align='left'>Data is empty</td></tr>")
	end if 
	rs_data_tab.Close
	set rs_data_tab = Nothing
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
