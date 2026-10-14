<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim textReportName : textReportName = "Update BKM"
Dim ScriptName : ScriptName = "Updatebkm.asp"

ScriptName =  ScriptName
%>
<html>
	<head>
<title>ePlaS - <%=PCase(textReportName)%></title>

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

 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> -->
<script language="JavaScript">
function showHint(str) {
if (str.length == 0) {
 str = "null";	
}
 fddate = document.getElementById("fddate").value;
 fcba = document.getElementById("fcba").value;
 gangcode = document.getElementById("gangcode").value;
 
    if (fcba.length == 0) {
    	alert(fcba);
        document.getElementById("status").innerHTML = "No Status";
        return;
    } else {
		document.getElementById("status").innerHTML = "<img src='css/images/load.gif'>";
        var xmlhttp = new XMLHttpRequest();
        xmlhttp.onreadystatechange = function() {
            if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
				output = xmlhttp.responseText;
				colosing = output.trim();
                
				if (colosing == "Approved"){
					document.getElementById("status").innerHTML = "<font color='red' >" + colosing + "</font>";
					document.getElementById("button_exe").innerHTML = '<input type="submit" value="UPDATE PLANNED" onclick=showHint("doupdate") >';
				}else if(colosing == "Not Yet Approved"){
					document.getElementById("status").innerHTML = "<font color='green' >" + colosing + "</font>";
					document.getElementById("button_exe").innerHTML = '<input type="button" style="color:#a0a0a0;" value="UPDATE PLANNED" disabled>';
				}else{
					document.getElementById("status").innerHTML = colosing;
					document.getElementById("button_exe").innerHTML = '<input type="button" style="color:#a0a0a0;" value="UPDATE PLANNED" disabled>';
				}
            }
        }
		 
		txt_FCBA = fcba.trim();
		// alert("ajax/update_bkmstatus.asp?fcba=" + txt_FCBA + "&gangcode=" + gangcode  + "&fddate=" + fddate  + "&exec=" + str );
        xmlhttp.open("GET", "ajax/update_bkmstatus.asp?fcba=" + txt_FCBA + "&gangcode=" + gangcode + "&fddate=" + fddate  + "&exec=" + str, true);
        xmlhttp.send();
    }
}
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
var getGangcode = function () {
	// build up the post string when passing variables to the server side page
	
	var sectionname = document.getElementById("sectionname").value;
	var selectedgangcode = document.getElementById("selectedgangcode").value;
	var PostStr = 'sectionname='+sectionname+'&gangcode='+selectedgangcode;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getlistgangcode_updbkm.asp', 'POST', PostStr, showGangcode);
}

// The function for handling the response from the server
var showGangcode = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("gangcodelist").innerHTML = "";
	document.getElementById("gangcodelist").innerHTML = response;

	var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"70%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
};

</script>
	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>

	<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %> onload="showHint('')">
		<!--#include file="top.jsp"-->
		<%
		Dim fddate : fddate = Request("fddate")
		Dim FCBA : FCBA = Request("FCBA")
		Dim GANGCODE : GANGCODE = Request("GANGCODE")
		dim status : status = "No Status"
		
		If Len(Trim(fddate)) = 0 Then
			fddate = textDate(Date)
		End if
		
		Dim txtTable, txtModule
		txtTable = Request("txtTable")
		txtModule = Request("txtModule")
		%>
		<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
		<div class="judul_head" valign="top" bgcolor="#fff" style=" width:85%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;">
				<div class="judul">
					Module : Generate Report/ <%=PCase(textReportName)%>
				</div>
				<div class="box_both" style="overflow:scroll;">
				<div style="border:1px #cccccc solid; padding:10px; width:400px;" >
					<table class = "data">
						<thead>
							
						</thead>
						<tbody>
							<tr>
								<td align="left" width="100">FCBA</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;<select id="fcba" name="txtFcba" onCHANGE="showHint('')" style="width:100px;">

										<%
										Dim rsFcba,strsql,v_txtFcba
										Set rsFcba = Server.CreateObject("ADODB.Recordset")
										rsFcba.CursorLocation = adUseClient
										strsql =    " select fcba from user_fcba_tab " & _
													" where " & _
													" fcusercode = '" & session("USERNAME") & "' "
										
										rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
										if not rsFcba.eof then
											DO while not rsFcba.eof
												if v_txtFcba = rsFcba("fcba") then
												%>			
													<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
												<% Else %>	
													<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
												<% End If
												
												rsFcba.movenext
											loop
										End if
										rsFcba.Close
										set rsFcba = Nothing
										%>
								</select>
								</td>
							</tr>
							<tr>
								<td align="left">SECTION NAME</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;
									<select name="sectionname" value="<%=sectionname%>" id="sectionname" onchange="javascript:getGangcode();" class="chosen-select" style="width: 100%;">
											<option value=""></option> 
												<% 	
													q = " SELECT DISTINCT emp.sectionname, sc.fcname" & _
														"    FROM employee emp" & _
														"         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
														"   WHERE fcba = '"& session("fcba") &"' ORDER BY emp.sectionname"
													' response.write q
													Set rsComboC = Server.CreateObject("ADODB.Recordset")
													rsComboC.CursorType = 3
													rsComboC.CursorLocation = 3
													rsComboC.Open q, connstr
													if not rsComboC.eof then%>
													<%
													Do While Not rsComboC.eof
													%>
														<option value="<%=rsComboC("sectionname")%>" 
															<%if request("sectionname") = cstr(""&rsComboC("sectionname")) then %> selected 
															<%End If%>><%=rsComboC("sectionname")%>
														</option> 
													<%rsComboC.movenext
													loop
													End if
													rsComboC.Close
													set rsComboC = Nothing
													q=""
												%>
										</select>  
								</td>		
							</tr>
							<tr>
								<td  align="left">GANG</td>
								<td align="left">
									<div id="gangcodelist"></div> 
								</td>
									<input type="Hidden" id="selectedgangcode" name="selectedgangcode" value="<%=request("selectedgangcode")%>">
							</tr>  
							<tr>
								<td align="left">DATE</td>
								<td align="left">
									:&nbsp;&nbsp;&nbsp;&nbsp;<input id="fddate" class="datepicker" type="text" maxlength="10" name="fddate" size="10" value="<%=fddate%>" onCHANGE="showHint('')"  onkeyup="showHint('')" style="width:100px;">
								</td>
							</tr>
							<tr>
								<td align="left">STATE</td>
								<td align="left" >
									:&nbsp;&nbsp;&nbsp;&nbsp;<span id="status" style="font-weight:bold; font-size:13px;"><%=status%></span>
									&nbsp;&nbsp;&nbsp;
								</td>
							</tr>
							<tr>
								<td colspan="2" align="left" style="padding-top:30px;"><span id="button_exe"></span></td>
							</tr>
							
						</tbody>
					</table>
				</div>
				</div>
		</div>
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

function getselectedgangcode() {
	var sgangcode = document.getElementById("selectedgangcode");
	sgangcode.value = document.getElementById("gangcode").value;
}

Document.prototype.ready = function(callback) {
  if(callback && typeof callback === 'function') {
    document.addEventListener("DOMContentLoaded", function() {
      if(document.readyState === "interactive" || document.readyState === "complete") {
        return callback();
      }
    });
  }
};
 

document.ready(function() {
	getGangcode(); 
	getselectedgangcode();
});
</script> 
	
	</body>
</html>