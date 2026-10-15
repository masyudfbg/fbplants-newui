<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/mm.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
	<title>ePlaS - Authorize Structure</title>
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
</script>

<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.]/g,'');}
</script>

<script language="javascript" type="text/javascript">
/** XHConn - Simple XMLHTTP Interface - bfults@gmail.com - 2005-04-08        **
 ** Code licensed under Creative Commons Attribution-ShareAlike License      **
 ** http://creativecommons.org/licenses/by-sa/2.0/                           **/
function XHConn()
{
  var xmlhttp, bComplete = false;
  try { xmlhttp = new ActiveXObject("Msxml2.XMLHTTP"); }
  catch (e) { try { xmlhttp = new ActiveXObject("Microsoft.XMLHTTP"); }
  catch (e) { try { xmlhttp = new XMLHttpRequest(); }
  catch (e) { xmlhttp = false; }}}
  if (!xmlhttp) return null;
  this.connect = function(sURL, sMethod, sVars, fnDone)
  {
    if (!xmlhttp) return false;
    bComplete = false;
    sMethod = sMethod.toUpperCase();
    try {
      if (sMethod == "GET")
      {
        xmlhttp.open(sMethod, sURL+"?"+sVars, true);
        sVars = "";
      }
      else
      {
        xmlhttp.open(sMethod, sURL, true);
        xmlhttp.setRequestHeader("Method", "POST "+sURL+" HTTP/1.1");
        xmlhttp.setRequestHeader("Content-Type",
          "application/x-www-form-urlencoded");
      }
      xmlhttp.onreadystatechange = function(){
        if (xmlhttp.readyState == 4 && !bComplete)
        {
          bComplete = true;
          fnDone(xmlhttp);
        }};
      xmlhttp.send(sVars);
    }
    catch(z) { return false; }
    return true;
  };
  return this;
}

  	// doAJAXCall : Generic AJAX Handler, used with XHConn
// Author : Bryce Christensen (www.esonica.com)
// PageURL : the server side page we are calling
// ReqType : either POST or GET, typically POST
// PostStr : parameter passed in a query string format 'param1=foo&param2=bar'
// FunctionName : the JS function that will handle the response

var doAJAXCall = function (PageURL, ReqType, PostStr, FunctionName) {

	// create the new object for doing the XMLHTTP Request
	var myConn = new XHConn();

	// check if the browser supports it
	if (myConn)	{
	    
	    // XMLHTTPRequest is supported by the browser, continue with the request
	    myConn.connect('' + PageURL + '', '' + ReqType + '', '' + PostStr + '', FunctionName);    
	} 
	else {
	    // Not support by this browser, alert the user
	    alert("XMLHTTP not available. Try a newer/better browser, this application will not work!");   
	}
}

// launched from button click 
var getAuthorizeTemplateInfo = function () {
	// build up the post string when passing variables to the server side page

	var authorizetemplatecode = document.getElementById("authorizetemplatecode").value;
	var PostStr = 'authorizetemplatecode='+authorizetemplatecode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getauthorizetemplateinfo.asp', 'POST', PostStr, showAuthorizeTemplateInfo);
}



// The function for handling the response from the server
var showAuthorizeTemplateInfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.document_type);
    // update the Div to show the result from the server
	document.getElementById("document_type").value = s.document_type;
	document.getElementById("document_classification_name").value = s.document_classification_name;
	document.getElementById("document_classification").value = s.document_classification;
}
</script>

<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">

<style>
.close-button {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 30px;
  height: 30px;
  position: relative;
  padding: 0 1px 0 0;
  border: none;
  -webkit-border-radius: 1em;
  border-radius: 1em;
  font: normal 8em/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #6d6d6d;
  cursor:pointer;
}

.close-button:hover {background: #f72c2c;}

.close-button::before {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 5px;
  height: 20px;
  position: absolute;
  content: "";
  top: 5px;
  left: 13px;
  border: none;
  font: normal 100%/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #ffffff;
  text-shadow: none;
  -webkit-transform: rotateZ(45deg)   ;
  transform: rotateZ(45deg)   ;
}

.close-button::after {
  display: inline-block;
  -webkit-box-sizing: content-box;
  -moz-box-sizing: content-box;
  box-sizing: content-box;
  width: 5px;
  height: 20px;
  position: absolute;
  content: "";
  top: 5px;
  left: 13px;
  border: none;
  font: normal 100%/normal Arial, Helvetica, sans-serif;
  color: rgba(0,0,0,1);
  -o-text-overflow: clip;
  text-overflow: clip;
  background: #ffffff;
  text-shadow: none;
  -webkit-transform: rotateZ(-45deg)   ;
  transform: rotateZ(-45deg)   ;
}
</style>

<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261015a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %>>
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
<%
'--Awal--
txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

authorizetemplatecode = Request("authorizetemplatecode")
document_type = Request("document_type")
document_classification = Request("document_classification")
document_classification_name = Request("document_classification_name")
fcba = Request("fcba")
runningnumber = Request("runningnumber")

arr_authorizetemplate_eof = true
oSql = " SELECT fccode, fcname FROM authorizetemplate WHERE isactive = 'TRUE' Order by fcname"

set oRs = server.CreateObject("Adodb.recordset")
oRs.Open oSql, connstr
if Not oRs.Eof then
	arr_authorizetemplate = oRs.getrows() : arr_authorizetemplate_eof = false
End if
oRs.Close : Set oRs = Nothing

arr_fcba_eof = true
oSql = " SELECT fccode, fcname FROM businessunit WHERE isactive = 'TRUE' ORDER BY orderdata"

set oRs = server.CreateObject("Adodb.recordset")
oRs.Open oSql, connstr
if Not oRs.Eof then
	arr_fcba = oRs.getrows() : arr_fcba_eof = false
End if
oRs.Close : Set oRs = Nothing


txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Header" then
		arr_insertdata_eof = true
		oSql = " SELECT fcba, accountingyear, accountingperiod, fieldcode, fccode" & _
               " FROM trans_bisa_agrotech_dtl" & _
               " WHERE fcba = '" & Session("fcBA") & "'" & _
               "       AND accountingyear = '" & txtyear & "'" & _
               "       AND accountingperiod = '" & txtmonth & "'" & _
               "       AND fieldcode = '" & txtblok & "'"

		set oRs = server.CreateObject("Adodb.recordset")
		oRs.Open oSql, connstr
		if Not oRs.Eof then
			arr_insertdata = oRs.getrows()
			arr_insertdata_eof = false
		End if
		oRs.Close
		Set oRs = Nothing
		
		if not arr_insertdata_eof then
			for nn=0 to ubound(arr_insertdata,2)
				q = " DELETE FROM trans_bisa_agrotech_dtl" & _
                    " WHERE fcba = '" & Session("fcBA") & "'" & _
                    "       AND accountingyear = '" & txtyear & "'" & _
                    "       AND accountingperiod = '" & txtmonth & "'" & _
                    "       AND fieldcode = '" & txtblok & "'" & _
					"       AND fccode = '" & arr_insertdata(4, nn) & "'"
				connStr.execute(q)
			next
		end if
		
		q = " DELETE FROM trans_bisa_agrotech_hdr" & _
			" WHERE fcba = '" & Session("fcBA") & "'" & _
			"       AND accountingyear = '" & txtyear & "'" & _
			"       AND accountingperiod = '" & txtmonth & "'" & _
			"       AND fieldcode = '" & txtblok & "'"
		connStr.execute(q)
		
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
		arr_insertdata_eof = true
		oSql = " SELECT authorizetemplatecode, authorizelevelcode, fcusercode, fcba" & _
               " FROM authorizestructure" & _
               " WHERE authorizetemplatecode = '" & authorizetemplatecode & "'" & _
               "       AND fcba = '" & fcba & "'"
		'response.write "<br/>"&oSql
		set oRs = server.CreateObject("Adodb.recordset")
		oRs.Open oSql, connstr
		if Not oRs.Eof then
			arr_insertdata = oRs.getrows()
			arr_insertdata_eof = false
		End if
		oRs.Close
		Set oRs = Nothing
		
		if not arr_insertdata_eof then
			for nn=0 to ubound(arr_insertdata,2)
				q = " DELETE FROM authorizestructure" & _
                    " WHERE authorizetemplatecode = '" & authorizetemplatecode & "'" & _
                    "       AND fcba = '" & fcBa & "'" & _
                    "       AND authorizelevelcode = '" & arr_insertdata(1, nn) & "'" & _
                    "       AND fcusercode = '" & arr_insertdata(2, nn) & "'"
				'response.write "<br/>"&q
				connStr.execute(q)
			next
		end if
		
		
			for nn=1 to cdbl(to_nol(request("runningnumber")))
				if cdbl(to_nol(request("cb_"&nn))) = 1 then
					q = " INSERT INTO authorizestructure ( authorizetemplatecode" & _
					    "                                  , authorizelevelcode" & _
					    "                                  , fcusercode" & _
					    "                                  , fcentry" & _
					    "                                  , fcedit" & _
					    "                                  , fcip" & _
					    "                                  , fcba" & _
					    "                                  , lastupdate" & _
					    "                                  , lasttime)" & _
						" VALUES ( '" & authorizetemplatecode & "'" & _
						"          , '" & request("al_"&nn) & "'" & _
						"          , '" & request("user_"&nn) & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Session("CodeName") & "'" & _
						"          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"          , '" & fcba & "'" & _
						"          , SYSDATE" & _
						"          , TO_CHAR(SYSDATE, 'hh24:mi') )"
					'response.write q
					connStr.execute(q)
				end if
			next
		
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

val_IsTrue = "N"
oSql = " SELECT rpt_bisa_agrotech_api.bolehinput('" & txtyear & "'||LPAD('" & txtmonth & "',2,0), '" & Session("fcBA") & "') istrue" & _
       " FROM DUAL"
			
set oRs = server.CreateObject("Adodb.recordset")
oRs.Open oSql, connstr
If Not oRs.Eof then
	val_IsTrue = oRs("istrue")
Else
	val_IsTrue = "N"
End if
oRs.Close
Set oRs = Nothing


If Len(Trim(CompleteStringError)) > 0 Then
	response.Write  "<div class = ""error"">" & CompleteStringError & "</div>"
End If
if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if
 %>
<form name="frmEntry" method="post" action="authorizestructure.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Authorize Template</td>
			<td align="left">
				<select id="authorizetemplatecode" name="authorizetemplatecode" style="width:75%;" class="chosen-select" onchange="javascript:getAuthorizeTemplateInfo();">
					<option value=""></option>
				<%
					if not arr_authorizetemplate_eof then
						for xi=0 to ubound(arr_authorizetemplate,2)
							if cstr(""&authorizetemplatecode) = cstr(""&arr_authorizetemplate(0, xi)) then
								response.write "<option value='"&arr_authorizetemplate(0, xi)&"' selected>"&arr_authorizetemplate(0, xi) & " : " & arr_authorizetemplate(1, xi)&"</option>"
							else
								response.write "<option value='"&arr_authorizetemplate(0, xi)&"'>"&arr_authorizetemplate(0, xi) & " : " & arr_authorizetemplate(1, xi)&"</option>"
							end if
						next
					end if
				%>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Module</td>
			<td align="left"><input type="text" size="50" id="document_type" name="document_type" value="<%=document_type%>" readonly="true"  style="background-color:#ccc;"></td>
		</tr>
		<tr>
			<td align="left">Classification</td>
			<td align="left"><input type="text" size="50" id="document_classification_name" name="document_classification_name" value="<%=document_classification_name%>" readonly="true"  style="background-color:#ccc;"></td>
			<input type="hidden" id="document_classification" name="document_classification" value="<%=document_classification%>">
		</tr>
		
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<select id="fcba" name="fcba" style="width:75%;" class="chosen-select">
					<option value=""></option>
				<%
					if not arr_fcba_eof then
						for xi=0 to ubound(arr_fcba,2)
							if cstr(""&fcba) = cstr(""&arr_fcba(0, xi)) then
								response.write "<option value='"&arr_fcba(0, xi)&"' selected>"&arr_fcba(0, xi) & " : " & arr_fcba(1, xi)&"</option>"
							else
								response.write "<option value='"&arr_fcba(0, xi)&"'>"&arr_fcba(0, xi) & " : " & arr_fcba(1, xi)&"</option>"
							end if
						next
					end if
				%>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;
				<%if val_IsTrue = "Y" then%>
				<a href="authorizestructure.asp?txtyear=<%=txtyear%>&txtmonth=<%=txtmonth%>&txtblok=<%=txtblok%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0">
				</a>
				<%end if%>
			</td>
		</tr>
	</table>  
</form>
<%
If cmdView = "True" and "" & authorizetemplatecode <> "" and "" & fcba <> "" Then	

	arr_data_eof = true
	q = " SELECT al.authorizelevelcode, al.authorizelevelname, al.fcusercode, uft.fcba, nvl(aus.isauthorize,0) isauthorize" & _
		" FROM (" & _
		" SELECT * FROM view_authorizelevel WHERE isactive = 'TRUE' AND startlevel = 'TRUE' AND isactive_user = 'TRUE'" & _
		" UNION ALL" & _
		" SELECT * FROM view_authorizelevel WHERE isactive = 'TRUE' AND startlevel = 'FALSE' AND finallevel = 'FALSE' AND isactive_user = 'TRUE'" & _
		" UNION ALL" & _
		" SELECT * FROM view_authorizelevel WHERE isactive = 'TRUE' AND finallevel = 'TRUE' AND isactive_user = 'TRUE'" & _
		" ) al" & _
		" INNER JOIN user_fcba_tab uft ON al.fcusercode = uft.fcusercode" & _
		" LEFT JOIN (SELECT authorizetemplatecode, authorizelevelcode, fcusercode, 1 isauthorize, fcba FROM authorizestructure ) aus ON (aus.authorizetemplatecode = '"&authorizetemplatecode&"' AND al.authorizelevelcode = aus.authorizelevelcode AND al.fcusercode = aus.fcusercode AND aus.fcba = uft.fcba)" & _
		" WHERE uft.fcba = '"&fcba&"' /*AND document_type = '"&document_type&"' AND document_classification = '"&document_classification&"'*/" & _
		" ORDER BY al.authorizelevelcode, al.fcusercode"

	'response.write q
	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open q, connstr
   
	if not oRs.eof then
		arr_data = oRs.getrows()
		arr_data_eof = false
	end if
	
	function f_rowspan (arr_, where_)
		result = 0
		arr1 = ""
		arr1 = arr_
		if not arr_data_eof then
			for xyz=0 to ubound(arr1,2)
				if cstr(""&arr1(0,xyz)) = cstr(""&where_) then result = result +1
			next
		end if
		
		f_rowspan = result
		
	end function
   
%>
	<form action="authorizestructure.asp" method="post" name="eForm">
	<table class="entry" cellpadding="1" cellspacing="1" width="100%" style="font-size:12px;">
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		   <td align="center">Code</td>
		   <td align="center">Description</td>
		   <td align="center">User Name</td>
		   <td align="center">Authorize</td>
		</tr>  

<%
	if not arr_data_eof then
		no = 0
		runningnumber = 0
		for xi=0 to ubound(arr_data,2)
			runningnumber = runningnumber +1
			curr_code = arr_data(0,xi)
			if curr_code <> prev_code then
				no = no +1
				if no  mod 2 = 1 then cColor = "background-color:#DDD;" else cColor = "background-color:#AAA;"
				response.write "<tr style='"&cColor&"'>"
				response.write "	<td align='center' valign='top' style='"&add_style&"' rowspan='"&f_rowspan(arr_data, arr_data(0,xi))&"'>"&arr_data(0,xi)&"</td>"
			else
				response.write "<tr style='"&cColor&"'>"
			end if
			response.write "	<td align='left' style='"&add_style&"'>"&arr_data(1,xi)&"</td>"
			response.write "	<td align='left' style='"&add_style&"'>"&arr_data(2,xi)&"</td>"
			
			if cdbl(to_nol(arr_data(4,xi))) = 1 then propchecked = "checked" else propchecked = ""

			response.write "	<td align='center' style='"&add_style&"'>"
			response.write "	    <input type='checkbox' name='cb_"&runningnumber&"' id='cb_"&runningnumber&"' class='item' value='"&cdbl(to_nol(arr_data(4,xi)))&"' "&propchecked&" onclick='handleClick(this);'>"
			response.write "	</td>"
			response.write "<input type='hidden' name=al_"&runningnumber&" value="&arr_data(0,xi)&">"
			response.write "<input type='hidden' name=user_"&runningnumber&" value="&arr_data(2,xi)&">"
			response.write "</tr>"
			
			prev_code = curr_code
			add_spasi = ""
			add_style = ""
			add_colspan = ""
		next
	else
		response.Write("<tr style='background-color:#DDD;'><td colspan='4' align='left'>Data is empty</td></tr>")	  
	end if 
%>
		<tr style='background-color:#fff;'>
			<td colspan="4" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="authorizetemplatecode" value="<%=authorizetemplatecode%>">
			<input type="hidden" name="document_type" value="<%=document_type%>">
			<input type="hidden" name="document_classification" value="<%=document_classification%>">
			<input type="hidden" name="document_classification_name" value="<%=document_classification_name%>">
			<input type="hidden" name="fcba" value="<%=fcba%>">
			<input type="hidden" name="runningnumber" value="<%=runningnumber%>">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<%IF Request("txtJenis")="Edit" then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit"  name="Add" value="Update">
			<%Else%>
				<input type="hidden" name="txtJenis" value="New">
				<input Type="Submit"  name="Add" value="Submit">
			<%End if%>
			<input Type="Reset"  name="Reset" value="Reset">
			<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
	</table>
</div>
</form>
<%End if%>
<!---------------akhir!--------------->
</div>
</div>
<!--#include file="bottom.jsp"-->	
	<script src="plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		
	</script>
	
	<script type="text/javascript">
		function handleClick(cb){
			if(cb.checked == 1){
				cb.value = 1;
			} else {
				cb.value = 0;
			}
        }
		
	
	</script>
	<script type="text/javascript">	
		$('#select-all').click( function () {
			$('.item').prop('checked', this.checked);
			$('.item').val(this.checked ? 1 : 0);
		})
	</script>
</body>
</html>
