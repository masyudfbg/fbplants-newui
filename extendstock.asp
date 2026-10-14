<%validatesession%>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/mm.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<!--#include virtual="/DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
	<title>Extend Stock - iPlaS</title>
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
var getMessage = function () {
	// build up the post string when passing variables to the server side page
	
	var fccode = document.getElementById("fccode").value;
	var PostStr = 'fccode='+fccode;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getstockinfo.asp', 'POST', PostStr, showMessageResponse);
}



// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.salescontractno);
    // update the Div to show the result from the server
	document.getElementById("fcname").value = s.fcname;
	document.getElementById("unitofmeasurement").value = s.unitofmeasurement;
}

var getTicket = function () {
	// build up the post string when passing variables to the server side page

	var chitno = document.getElementById("chitno").value;
	var PostStr = 'chitno='+chitno;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/GetTicket.asp', 'POST', PostStr, ShowTicketResponse);
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
<%
'--Awal--
Dim stringduedate : stringduedate = 0

Dim IsOverDuedate, varRemainDueDate, varDueDate, CompleteStringError, varErrorDueDate, stringremainqty, varRemainQty, IsOverRemainQty, varErrorRemainQty, varEditedQty

Dim qtyplanedit : qtyplanedit = Request("qtyplanedit")
Dim unitplanedit : unitplanedit = Request("unitplanedit")
Dim outputplanedit : outputplanedit = Request("outputplanedit")
Dim qtyfinaledit : qtyfinaledit = Request("qtyfinaledit")
Dim unitfinaledit : unitfinaledit = Request("unitfinaledit")
Dim outputfinaledit : outputfinaledit = Request("outputfinaledit")
Dim txtagreementlineno : txtagreementlineno = Request("txtagreementlineno")

Dim txtGRNCode : txtGRNCode = 0
Dim oRs : oRs = 0
Dim oSql : oSql = 0
Dim txtLocationCode : txtLocationCode = ""

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

fccode = Request("fccode")
fcname = Request("fcname")
unitofmeasurement = Request("unitofmeasurement")
max_no = Request("max_no")


txtJenis = request("txtJenis") '-- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
'response.write "txtjenis:" & txtJenis & " - " & Request("Add")
'--- Action ke Database di Oracle [ insert,edit,delete ]
if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Header" then		
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		cmdView	= "True"					
	elseif Request("txtJenis")="New" then
			for nn=0 to cdbl(to_nol(max_no))
				if cdbl(to_nol(request("cb_"&nn))) = 1 and cdbl(to_nol(request("has_"&nn))) = 0 then
					q = " INSERT INTO stock ( fccode" & _
					    "                     , fcname" & _
					    "                     , itemtype" & _
					    "                     , manufacturepartnumber" & _
					    "                     , manufacturecode" & _
					    "                     , supplierpartnumber" & _
					    "                     , suppliercode" & _
					    "                     , unitofmeasurement" & _
					    "                     , stockgroupcode" & _
					    "                     , activematerial" & _
					    "                     , validuntil" & _
					    "                     , isactive" & _
					    "                     , min_value" & _
					    "                     , max_value" & _
					    "                     , fcentry" & _
					    "                     , fcedit" & _
					    "                     , fcip" & _
					    "                     , fcba" & _
					    "                     , islesszero" & _
					    "                     , lastupdate" & _
					    "                     , lasttime)" & _
						" SELECT fccode" & _
					    "        , fcname" & _
					    "        , itemtype" & _
					    "        , manufacturepartnumber" & _
					    "        , manufacturecode" & _
					    "        , supplierpartnumber" & _
					    "        , suppliercode" & _
					    "        , unitofmeasurement" & _
					    "        , stockgroupcode" & _
					    "        , activematerial" & _
					    "        , validuntil" & _
					    "        , isactive" & _
					    "        , min_value" & _
					    "        , max_value" & _
					    "        , '" & Session("CodeName") & "' fcentry" & _
					    "        , '" & Session("CodeName") & "' fcedit" & _
					    "        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
					    "        , '" & request("bu_"&nn) & "' fcba" & _
					    "        , islesszero" & _
					    "        , SYSDATE lastupdate" & _
					    "        , TO_CHAR(SYSDATE, 'hh24:mi') lasttime" & _
					    " FROM bankstock" & _
					    " WHERE fccode = '" & fccode & "'"
					'response.write q
					connStr.execute(q)
					elseif cdbl(to_nol(request("cb_"&nn))) = 1 and cdbl(to_nol(request("has_"&nn))) = 1 then
					q = " MERGE INTO stock s1 " & _ 
					    " USING (SELECT fccode" & _
					    "          , fcname" & _
					    "          , unitofmeasurement,MANUFACTUREPARTNUMBER,MANUFACTURECODE,itemtype " & _
						"		FROM bankstock WHERE fccode = '" & fccode & "') s2 "  & _
						" ON (s1.fccode=s2.fccode) " & _ 
						" WHEN MATCHED THEN UPDATE SET s1.fcname=s2.fcname,s1.itemtype=s2.itemtype,s1.unitofmeasurement=s2.unitofmeasurement,s1.MANUFACTUREPARTNUMBER=s2.MANUFACTUREPARTNUMBER,s1.MANUFACTURECODE=s2.MANUFACTURECODE,s1.fcedit='" & Session("CodeName") & "',lastupdate=SYSDATE" & _
					    " WHERE s1.fcba = '" & request("bu_"&nn) & "'" 
			'		response.write q
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

If Len(Trim(CompleteStringError)) > 0 Then
	response.Write  "<div class = ""error"">" & CompleteStringError & "</div>"
End If
if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if
 %>
<form name="frmEntry" method="post" action="extendstock.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left" style="width:20%;">Stock Code</td>
			<td align="left">
				<input type="text" maxlength="15" id="fccode" name="fccode" size="30" value="<%=fccode%>" autocomplete="Off" onBlur="javascript:getMessage();">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Stock Name</td>
			<td align="left">
				<input type="text" maxlength="150" id="fcname" name="fcname" size="60" value="<%=fcname%>" readonly style="background-color:#eee;">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Uom</td>
			<td align="left">
				<input type="text" maxlength="15" id="unitofmeasurement" name="unitofmeasurement" size="30" value="<%=unitofmeasurement%>" readonly style="background-color:#eee;">
				<font color="#FF0000">&nbsp;*</font>
			</td>
			</td>
		</tr>
		
		<tr>
			<td colspan="2" align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">&nbsp;
				<!--
				<a href="extendstock.asp?txtyear=<%=txtyear%>&txtmonth=<%=txtmonth%>&txtblok=<%=txtblok%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0">
				</a>
				-->
			</td>
		</tr>
	</table>  
</form>
<%
If cmdView = "True" and "" & fccode <> ""and "" & fcname <> ""and "" & unitofmeasurement <> "" Then	

	arr_data_eof = true
	q = " SELECT bs.fcba, bu.fcname, CASE WHEN st.fccode IS NULL THEN 0 ELSE 1 END has_data FROM (" & _
		" SELECT bs.fccode, bs.fcname, bs.itemtype, bs.unitofmeasurement, bs.stockgroupcode" & _
		"        , bs.activematerial, bs.manufacturepartnumber, bs.manufacturecode, bs.supplierpartnumber, bs.suppliercode" & _
		"        , bs.islesszero, bs.validuntil, bs.isactive, bs.min_value, bs.max_value, uft.fcba" & _
		" FROM bankstock bs, user_fcba_tab uft" & _
		" WHERE bs.fccode = '" & fccode & "'" & _
		"       AND uft.fcusercode = '" & Session("CodeName") & "'" & _
		" ) bs LEFT JOIN stock st ON (bs.fccode = st.fccode AND bs.fcba = st.fcba)" & _
		" LEFT JOIN businessunit bu ON bu.fccode = bs.fcba" & _
		" ORDER BY bu.orderdata"

	set oRs = server.CreateObject("Adodb.recordset")
	oRs.Open q, connstr
   
	if not oRs.eof then
		arr_data = oRs.getrows()
		arr_data_eof = false
	end if

   
%>
	<form action="extendstock.asp" method="post" name="eForm">
	<table class="entry" cellpadding="1" cellspacing="1" width="100%" style="font-size:12px;">
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;">
		   <td align="center">Business Unit</td>
		   <td align="center">Description</td>
		   <td align="center">Extend</td>
		</tr>  

<%
	response.write "<tr style='background-color:#FF851B;'>"
	response.write "    <td colspan='2'>&nbsp;</td>"
	response.write "    <td align='center'><input type='checkbox' name='select-all' id='select-all'/> Check All/ Uncheck All</td>"
	response.write "</tr>"
	max_no = 0
	if not arr_data_eof then
		max_no = 0
		for xi=0 to ubound(arr_data,2)
			max_no = max_no +1
		next
	end if
	if not arr_data_eof then
		no = 1
		for xi=0 to ubound(arr_data,2)
		if xi mod 2 = 1 then cColor = "background-color:#e5e5e5;color:#111;" else cColor = "background-color:#eee;color:#111;"
			
			response.write "<tr style='"&cColor&"'>"
			response.write "	<td align='center' style='"&add_style&"'>"&arr_data(0,xi)&"</td>"
			response.write "	<td align='left' style='"&add_style&"' "&add_colspan&">"&add_spasi&arr_data(1,xi)&"</td>"
			propchecked = ""
			if cdbl(to_nol(arr_data(2,xi))) = 1 then
				propchecked = "checked"
			else
				propchecked = ""
			end if
			response.write "	<td align='center' style='"&add_style&"'>"
			response.write "	    <input type='checkbox' name='cb_"&no&"' id='cb_"&no&"' class='item' value='"&cdbl(to_nol(arr_data(2,xi)))&"' "&propchecked&" onclick='handleClick(this);'>"
			response.write "	</td>"
			response.write "	    <input type='hidden' name='bu_"&no&"' value='"&arr_data(0,xi)&"'>"
			response.write "	    <input type='hidden' name='has_"&no&"' value='"&cdbl(to_nol(arr_data(2,xi)))&"'>"
			response.write "</tr>"
			add_spasi = ""
			add_style = ""
			add_colspan = ""
			no = no +1
		next
	else
		response.Write("<tr><td colspan='3' align='left'>Data is empty</td></tr>")	  
	end if 
%>
		<tr style='background-color:#fff;'>
			<td colspan="3" align="left">
			<!-- diambil dari form yang diatas -->
			<input type="hidden" name="fccode" value="<%=fccode%>">
			<input type="hidden" name="fcname" value="<%=fcname%>">
			<input type="hidden" name="unitofmeasurement" value="<%=unitofmeasurement%>">
			<input type="hidden" name="max_no" value="<%=max_no%>">
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
