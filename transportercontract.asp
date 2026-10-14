<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - Transporter Contract</title>

<link rel="StyleSheet" href="dtree.css" type="text/css" />
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<script>
	function harusangka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
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

var getlistfcba = function () {
	// build up the post string when passing variables to the server side page
	
	var fccompanycode = document.getElementById("fccompanycode").value;
	var selectedfcba = document.getElementById("selectedfcba").value;
	var PostStr = 'fccompanycode='+fccompanycode+'&fcba='+selectedfcba;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getfcbabasedoncompany.asp', 'POST', PostStr, showlistfcba);
}

// The function for handling the response from the server
var showlistfcba = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;

    // update the Div to show the result from the server
	document.getElementById("fcba").innerHTML = "";
	document.getElementById("fcba").innerHTML = response;
}
</script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->
	<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
		</div>
   <div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">
 <!-------------awal!------------->
 <%

Dim qSupplyingDeclaration, rsSupplyingDeclarationInfo, tempQuantity, tempDecType
Dim v_FCENTRY, v_FCEDIT, v_FCIP, v_FCBA, v_FCCODE, v_FCNAME
Dim qSupplyingInfo, rsSupplyingInfo, qDeclarationType, rsDeclarationType
Dim txtmonth, txtyear 
Dim parcode, totalstand, ScriptName, transportercontractno
Dim QUANTITY, censuscode, ruleinfo
Dim REMARKS, transporter, losstolerance, shippingaddress, rate, unitofmeasure
Dim creationdate, startdate, enddate
Dim cmdView
Dim percent_bunch, remarks_bunch, percent_abw, remarks_abw, plantingdate, division 
Dim distribution1, distribution2, distribution3, distribution4
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

txtyear =request("txtyear")
transportercontractno =request("transportercontractno")
transporter=request("transporter")
remarks=request("remarks")
losstolerance=request("losstolerance")
shippingaddress = Request("shippingaddress")
rate = Request("rate")
startdate = Request("startdate")
enddate = Request("enddate")
unitofmeasure = Request("unitofmeasure")
creationdate = request("creationdate")

fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")
lineno = Request("lineno")

mon = Month(Session("v_Costbook"))
If len(txtmonth)=0 then txtmonth=mon

Yr = Year(Session("v_Costbook"))

If len(txtyear)= 0 then txtyear = Yr
transdate = txtmonth & "/1/" & txtyear

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   txtyear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   txtmonth	=	Month(Session("v_Costbook"))
end if


Function CheckFieldCode(initialFieldCode, initialFcba, initialDate)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" select count(1) total" & _
	                " from dual " & _
					" where field_api.GetValid_Activation('" & initialFieldCode & "', '" & initialFcba & "', to_date('"& initialDate &"', 'mm/dd/yyyy')) = 'Y'"
	' response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckFieldCode = initialAvailable
	'Select count(1) from dual where field_api.GetValid_Activation('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')) = 'Y'
	'select to_char (field_api.GetValid_Plantingdate('A01', 'KTP', to_date('03/19/2018', 'mm/dd/yyyy')),'yyyy') from dual
End Function



If Request("txtJenis") = "New" or Request("txtJenis") = "HeaderDelete" or Request("txtJenis") = "Delete" or Request("txtJenis") = "Edit" Then
	if Request("txtJenis")= "HeaderDelete" then
	    
        cmdView	= "True"
	elseif Request("txtJenis")= "Delete" then

			q = " delete from transportercontract " & _
				" where transportercontractno = '" & transportercontractno & "'" & _
				"       and fcBA = '" & fcba & "'" & _
				"		and lineno = '"& lineno &"'"
						
				' Response.Write q
				connStr.execute(q)
				cmdView	= "True"

	elseif Request("txtJenis")="New" then		
			If Len(Trim(transporter)) = 0 Then
				Err.Number = 1001
				Err.Description = "Transporter Date Harus diisi, tidak boleh kosong"
			Else
				'-- Get Data Max Linenokey
			querymaxlineno = " select nvl(max(lineno), 0) lineno" & _ 
							 " from transportercontract" & _ 
							 " where fcba = '" & fcba & "' and transportercontractno = '"& transportercontractno &"'"

			Set rsmaxlineno = Server.CreateObject("ADODB.Recordset")
			rsmaxlineno.CursorType = 3
			rsmaxlineno.CursorLocation = 3
			
			'response.write querymaxlineno
			
			rsmaxlineno.Open querymaxlineno, connstr
			if not rsmaxlineno.eof then
			   maxlineno = Cdbl(rsmaxlineno("lineno"))
			else
			   maxlineno = 0
			end if
			rsmaxlineno.close
			set rsmaxlineno = Nothing
			
			maxlineno = maxlineno + 1
				
				q = " INSERT INTO transportercontract" & _
					" (transportercontractno, creationdate, transporter, startdate, enddate, shippingaddress, rate, unitofmeasure, losstolerance, lineno, fcentry, fcedit, fcip, fcba, lastupdate, lasttime) VALUES (" & _
					" '" & transportercontractno & "'" & _
					" , to_date('" & creationdate & "','mm/dd/yyyy')" & _ 
					" , '" & transporter & "'" & _
					" , to_date('" & startdate & "','mm/dd/yyyy')" & _ 
					" , to_date('" & enddate & "','mm/dd/yyyy')" & _ 
					" , '" & shippingaddress & "'" & _
					" , REPLACE('" & rate & "',',','')" & _
					" , '" & unitofmeasure & "'" & _
					" , REPLACE('" & losstolerance & "',',','')" & _
					" , '" & maxlineno & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
					" , '" & fcba & "'" & _
					" , sysdate" & _
					" , to_char(sysdate, 'hh24:mi')" & _
					" )"
					' Response.Write q
					connStr.execute(q)
			End If
				
	cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			If Len(Trim(transporter)) = 0 Then
				Err.Number = 1001
				Err.Description = "Transporter Date Harus diisi, tidak boleh kosong"
			Else
					q =	" Update transportercontract" & _
						" set  transporter = '" & transporter & "'" & _
						"     , shippingaddress = '" & shippingaddress & "'" & _
						"     , rate = REPLACE('" & rate & "',',','')" & _
						"     , unitofmeasure = '" & unitofmeasure & "'" & _
						"     , losstolerance = REPLACE('" & losstolerance & "',',','')" & _
						"     , fcentry = '" & Session("CodeName") & "'" & _
						"     , fcedit = '" & Session("CodeName") & "'" & _
						" 	  , fcip ='" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
						"     , lastupdate = sysdate " & _
						"     , lasttime = to_char(sysdate,'hh24:mi') " & _
						" where transportercontractno = '" & transportercontractno & "'" & _
						"       and fcBA = '" & fcba & "'" & _
						"		and lineno = '"& lineno &"'"

						
		 			' response.write q
					connStr.execute(q)
					
				
			End If
		
		cmdView	= "True"
	else
		cmdView	= "True"
 	End if
	Else
	 	cmdview = request("cmdview")
 	End if
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")

if Len(Trim(creationdate)) = 0 Then	creationdate = textDate(date)
If Len(Trim(startdate)) = 0 Then startdate = textDate(date)
If Len(Trim(enddate)) = 0 Then	enddate = textDate(date)

If Err.Number <> 0 then
	Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if


oSQLQuery = " SELECT tc.*, bu.fccompanycode" & _
            " FROM transportercontract tc left join businessunit bu on bu.fccode = tc.fcba" & _
	        " WHERE tc.transportercontractno = '" & transportercontractno & "'"
'response.Write oSQLQuery

set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	creationdate = oRecordSet("creationdate")
	fccompanycode = oRecordSet("fccompanycode")
	fcba = oRecordSet("fcba")
	selectedfcba = oRecordSet("fcba")
End if
oRecordSet.Close : Set oRecordSet = Nothing

Dim val_IsTrue : val_IsTrue = "N"

oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & fcba & "'" & _
	        "       and to_date ('"& txtmonth &"'||'/01/'||'" & txtyear &"', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_char(to_date('"& txtmonth &"'||'/01/'||'" & txtyear &"','mm/dd/yyyy'),'yyyy')= to_char(currentdate,'yyyy')" & _
	        "       and upper(system) = upper('costbook')"
'response.Write oSQLQuery

set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
oRecordSet.Close : Set oRecordSet = Nothing

%>

 <!-------------ISI Supplying Declaration Form!------------->
 
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Transporter Contract</div><% end if %>
	<form name="frmEntry" method="post" action="<%=ScriptName%>">
		<table class="data" width="100%">
		<tr>
			<td align="left" width="20%">Transporter Contract No
				<td align="left">
				<input type="text" maxlength="50" name="transportercontractno" size="20" value="<%=transportercontractno%>" autocomplete="Off">
				<font color="#FF0000">*</font>
				</td>
			</td>
        </tr>
        <tr>
			<td align="left" width="20%">Creation Date 
			<td align="left">
				<input class="datepicker" type="text" size="15" name="creationdate" value="<%=creationdate%>">
				<font color="#FF0000">*</font>
			</td>
		</tr>
		
		<tr>
			<td align="left">Company</td>
			<td align="left">
				<select id="fccompanycode" name="fccompanycode" onchange="javascript:getlistfcba();">
					<option value=""></option>
					<%
					qlov = "SELECT DISTINCT bu.fccompanycode fccode FROM user_fcba_tab uft INNER JOIN businessunit bu ON bu.fccode = uft.fcba WHERE fcusercode = '" & Session("CodeName") & "'"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if fccompanycode = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%></option>
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
			<td align="left">Business Unit</td>
			<td align="left">
				<select id="fcba" name="fcba" onchange="javascript:getselectedfcba();">
					<option value=""></option>
				</select>
				<font color="#FF0000">*</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
		</tr>
				
		<tr>
			<td colspan="2"  align="left" class="fbp-doc-actions">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				
				<!-- <img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"> -->
				<!-- HADIR -->
				<!-- <%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Nursery"
				dataSubmodule 	= "Culling"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a> -->
				<!-- END:HADIR -->
			</td>
		</tr>
		</table>
	</form>
<%
if cmdView = "True" and len(trim(transportercontractno)) > 0 and len(trim(creationdate)) > 0 and len(trim(fcba)) > 0 Then
	Dim rsVH, intRecordsPerPage, cColor
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	=   " select * from transportercontract " & _
	    	" WHERE fcba = '"& fcba &"' " & _ 
		    " AND transportercontractno = '"& transportercontractno &"'"

		   ' " SELECT tc.transportercontractno, tc.creationdate, tc.transporter, tc.startdate, tc.enddate, tc.shippingaddress, tc.rate, tc.losstolerance, tc.unitofmeasure, " & _
		   ' " tr.FCNAME as TRANSPORTNAME, tr.FCCODE as TRANSPORTCODE " & _
		   ' " FROM transportercontract tc " & _ 
		   ' " INNER JOIN transporter tr ON (tc.fcba = tr.fcba AND tc.transportercontractno = tr.fccode) " & _
 		  '  " WHERE tc.fcba = '"& Session("fcBA") &"' " & _ 
		   ' " AND tc.transportercontractno = '"& transportercontractno &"'"

	' response.Write q

	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
	<form action="transportercontract.asp" method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
		
		<tr>
			<td align="left" width="20%">Transporter 
				<td align="left">
				<select name="transporter" title="Transporter">
					<option value="<%=transporter%>"></option> 
						<% 	Dim rsComboT
							q = "SELECT * FROM transporter WHERE fcba = '" & fcba & "' ORDER BY fccode"
							Set rsComboT = Server.CreateObject("ADODB.Recordset")
							rsComboT.Open q, connstr
							if not rsComboT.eof then%>
							<%
							Do While Not rsComboT.eof
							%>
								<option value="<%=rsComboT("fccode")%>" 
									<%if cstr(""&transporter)=cstr(""&rsComboT("fccode")) then %> selected 
									<%End If%>><%=rsComboT("fccode") & ": " & rsComboT("fcname")%>
								</option> 
							<%rsComboT.movenext
							loop
							End if
							rsComboT.Close
							set rsComboT = Nothing
							q=""
						%>
				</select> 
				<font color="#FF0000">*</font>
				</td>
			</td>

			<td align="left" width="20%">Shipping Address
				<td align="left">
				<textarea name="shippingaddress"><%=shippingaddress%></textarea>
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Start Date 
			<td align="left">
				<input class="datepicker" type="text" size="15" name="startdate" value="<%=startdate%>">
			</td>

			<td align="left" width="20%">End Date 
				<td align="left">
				<input class="datepicker" type="text" size="15" name="enddate" value="<%=enddate%>" >
				</td>
			</td>
		</tr>
		<tr>
			<td align="left" width="20%">Rate
				<td align="left">
				<input type="text" maxlength="15" name="rate" size="30" value="<%=rate%>" autocomplete="Off" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'">
				
				</td>
			</td>

			<td align="left" width="20%">Unit Of Measure
				<td align="left">
				<select name="unitofmeasure" title="Unit Of Measure Character a s/d z or 1 s/d 9">
					<option value="<%=UNITOFMEASURECODE%>"></option> 
						<% 	Dim rsComboP
							q = "SELECT * FROM UNITOFMEASURE WHERE fcba = '" & fcba & "'"
							Set rsComboP = Server.CreateObject("ADODB.Recordset")
							rsComboP.CursorType = 3
							rsComboP.CursorLocation = 3
							rsComboP.Open q, connstr
							if not rsComboP.eof then%>
							<%
							Do While Not rsComboP.eof
							%>
								<option value="<%=rsComboP("UNITOFMEASURECODE")%>" 
									<%if cstr(""&unitofmeasure)=cstr(""&rsComboP("UNITOFMEASURECODE")) then %> selected 
									<%End If%>><%=rsComboP("UNITOFMEASURECODE")%>
								</option> 
							<%rsComboP.movenext
							loop
							End if
							rsComboP.Close
							set rsComboP = Nothing
							q=""
						%>
				</select> 
			
				</td>
			</td>
		</tr>
		<tr>
		<td align="left" width="20%">Loss Tolerance
				<td align="left">
					<input type="text" name="losstolerance" size="30" value="<%=losstolerance%>" autocomplete="Off" onkeyup="return harusangka(this);"> %
				</td>
			</td>
		</tr>
			<tr>
				<td align="left" colspan="2">
					<input type="hidden" name="transportercontractno" value="<%=transportercontractno%>">
					<input type="hidden" name="creationdate" value="<%=creationdate%>">
					<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%If val_IsTrue = "Y" Then%>
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="txtJenis" value="Edit">
						<input type="hidden" name="plantingdate" value="<%=plantingdate%>">
						<input Type="Submit"  name="Add" value="Update">
					<%Else%>
						<input type="hidden" name="txtJenis" value="New">
						<input Type="Submit"  name="Add" value="Add">
					<%End if%>

					<input Type="Reset"  name="Reset" value="Reset">
					<input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
				<%End if%>
				</td>
			</tr>
		</table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<td align="left">NO</td>
		<td colspan="2" >&nbsp;</td>
		<td align="left">TRANSPORTER</td>
		<td align="left">START DATE</td>
		<td align="left">END DATE</td>
		<td align="left">SHIPPING ADDRESS</td>
		<td align="left">RATE</td>
		<td align="left">UNIT OF MEASURE</td>
		<td align="left">LOSS TOLERANCE</td>
	</tr>
	</thead>
	<tbody>
   
   <%
	if Not rsVH.eof then
		nCountofField = rsVH.fields.count - 1
		no = 1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=NO%></td>
						
			<td>
				
					<a href="transportercontract.asp?lineno=<%=RSvh("lineno")%>&transportercontractno=<%=RSvh("transportercontractno")%>&transporter=<%=rsVH("transporter")%>&startdate=<%=rsVH("startdate")%>&enddate=<%=rsVH("enddate")%>&shippingaddress=<%=rsVH("shippingaddress")%>&rate=<%=rsVH("rate")%>&unitofmeasure=<%=rsVH("unitofmeasure")%>&losstolerance=<%=rsVH("losstolerance")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=request("txtyear")%>" border="0"></a>
			</td>
			<td>
					<a href="transportercontract.asp?lineno=<%=RSvh("lineno")%>&transportercontractno=<%=RSvh("transportercontractno")%>&transporter=<%=rsVH("transporter")%>&startdate=<%=rsVH("startdate")%>&enddate=<%=rsVH("enddate")%>&shippingaddress=<%=rsVH("shippingaddress")%>&rate=<%=rsVH("rate")%>&unitofmeasure=<%=rsVH("unitofmeasure")%>&losstolerance=<%=rsVH("losstolerance")%>&fcba=<%=fcba%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=request("txtyear")%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
				
				
			</td>
			    <td align="center" title="<%=rsVH("transporter").Name%>" valign="top"><%=rsVH("transporter")%></td>
				<td align="center" title="<%=rsVH("startdate").Name%>" valign="top"><%=rsVH("startdate")%></td>
			    <td align="center" title="<%=rsVH("enddate").Name%>" valign="top"><%=rsVH("enddate")%></td>
			    <td align="center" title="<%=rsVH("shippingaddress").Name%>" valign="top"><%=rsVH("shippingaddress")%></td>
			    <td align="center" title="<%=rsVH("rate").Name%>" valign="top"><%=rsVH("rate")%></td>
			    <td align="center" title="<%=rsVH("unitofmeasure").Name%>" valign="top"><%=rsVH("unitofmeasure")%></td>
			    <td align="center" title="<%=rsVH("losstolerance").Name%>" valign="top"><%=rsVH("losstolerance")%></td>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  wend
   end if 
   rsVH.Close
   set rsVH = Nothing%>
   </tbody>
   </TABLE>
   </div>
<% if Request("oldui") <> "1" then %>
<script type="text/javascript">
// Status baris jadi badge (tampilan saja)
(function () {
	var cells = document.querySelectorAll('table.fbp-doc-lines tbody td[title]');
	for (var i = 0; i < cells.length; i++) {
		if (cells[i].getAttribute('title').toUpperCase() !== 'ROWSTATEID') continue;
		var s = (cells[i].textContent || '').replace(/\s+/g, ' ').trim();
		if (!s) continue;
		cells[i].parentNode.setAttribute('data-state', s);
		cells[i].innerHTML = '';
		var b = document.createElement('span');
		b.className = 'fbp-state';
		b.textContent = s;
		cells[i].appendChild(b);
	}
})();
</script>
<% end if %>
   </form>
<%End if%>
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->
		<script src="/js/jquery.inputmask.bundle.min.js"></script>
		<script src="/js/inputmask.binding.js"></script>
<script>

function getselectedfcba() {
	var sfcba = document.getElementById("selectedfcba");
	sfcba.value = document.getElementById("fcba").value;
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

// usage
document.ready(function() {
	getlistfcba();
});
</script>
</body>
</html>
