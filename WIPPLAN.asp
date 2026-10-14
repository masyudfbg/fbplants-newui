<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Contract Work in Progress Plan </title>
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
<!--<script language="JavaScript" src="/js/date-picker.js"></script>-->
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

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

var getcontractinfo = function () {
	// build up the post string when passing variables to the server side page
	var agreementcode = document.getElementById("agreementcode").value;
	var fcba = document.getElementById("selectedfcba").value;
	var plineno = document.getElementById("plineno").value;
	var PostStr = 'agreementcode='+agreementcode+'&fcba='+fcba+'&lineno='+plineno;
	//alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getcontractagreementinfo.asp', 'POST', PostStr, showcontractinfo);
	
}

// The function for handling the response from the server
var showcontractinfo = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // alert(s.itemcode);
    // update the Div to show the result from the server

	
    document.getElementById("jobcode").value = s.jobcode;
	document.getElementById("jobname").value = s.jobname;
	document.getElementById("rate").value = s.rate;
	document.getElementById("contract_ref").value = s.reference;
	
    
};
</script>

<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.]/g,'');}
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
Cek_IsDailyClosing = ""

private function get_targettype(jobcode, fcba)
	result_ = ""
	CheckQTY = " SELECT * FROM job WHERE fcba = '"&fcba&"' AND fccode = '"&jobcode&"'"
	' response.write CheckQTY
	Set rsCheckQTY = Server.CreateObject("ADODB.Recordset")
	rsCheckQTY.Open CheckQTY, connstr
	
	if Not rsCheckQTY.Eof then
		result_ = rsCheckQTY("job_category")
	Else
		result_ = ""
	End if
	get_targettype = result_
end function

fcCode =request("fcCode")
WIPDATE=REQUEST("WIPDATE")
QTY=request("QTY")
UNIT=request("UNIT")
OUTPUT=request("OUTPUT")
REFERENCE=request("REFERENCE")
REMARKS=request("REMARKS")

fccompanycode = request("fccompanycode")
fcba = request("fcba")
selectedfcba = Request("selectedfcba")

jobcode = Request("jobcode")

chargeunit = Request("chargeunit")
if len(trim(""&chargeunit))= 0 then 
	chargeunit = fcba
end if

locationtype = Request("locationtype")
locationcode = Request("locationcode")
lineno = Request("lineno")
maxlineno = Request("maxlineno")
		
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(fcba, WIPDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q = " DELETE FROM WORKINPROGRESSPLAN" & _
                " WHERE agreementcode = '" & FCCODE & "'" & _
                "       AND agreementlineno = '" & lineno & "'" & _
                "       AND fcba = '" & fcba & "'" & _
                "       AND wipdate = TO_DATE('" & WIPDATE & "','mm/dd/yyyy')" & _
				"       AND lineno = '" & maxlineno & "'"
					
			'Response.Write q
			ConnStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & WIPDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		locationtype = get_targettype(jobcode, fcba)
		Cek_IsDailyClosing = IsDailyClosing(fcba, WIPDATE)
		If Cek_IsDailyClosing = "FALSE" then
			'Validasi Qty
			CheckQTY = " SELECT fcba" & _
			           "        , agreementcode" & _
			           "        , NVL(SUM (totalagreement),0) totalagreement" & _
			           "        , NVL(SUM (totalwip),0) totalwip" & _
			           "        , NVL(SUM (totalagreement),0) - NVL(SUM (totalwip),0) selisih" & _
					   " FROM ( SELECT fcba, agreementcode, lineno agreementlineno, NVL(qty,0) totalagreement, 0 totalwip" & _
					   "        FROM contractagreementdetail WHERE agreementcode = '" & FCCODE & "' AND fcba = '" & fcba & "' AND lineno = '" & lineno & "'" & _
					   "        UNION ALL" &_
					   "        SELECT fcba, agreementcode, agreementlineno, 0 totalagreement, NVL(qty,0) totalwip" & _
					   "        FROM workinprogressplan WHERE agreementcode = '" & FCCODE & "' AND fcba = '" & fcba & "' AND agreementlineno = '" & lineno & "' ) dataall" & _
					   " WHERE fcba = '" & fcba & "' AND agreementcode = '" & FCCODE & "'" & _
					   " AND agreementlineno = '" & lineno & "' GROUP BY fcba, agreementcode"
			' response.write CheckQTY
			Set rsCheckQTY = Server.CreateObject("ADODB.Recordset")
			rsCheckQTY.Open CheckQTY, connstr
			
			if Not rsCheckQTY.Eof then
				varSelisih = cdbl(to_nol(rsCheckQTY("selisih")))
			Else
				varSelisih = 0		   
			End if

			if (varSelisih - cdbl(to_nol(QTY)) ) >=  0 then
				checkRate = " SELECT dtl.rate, hdr.finishdate" & _
				            " FROM contractagreementdetail dtl" & _
							" INNER JOIN contractagreement hdr ON (hdr.agreementcode = dtl.agreementcode AND hdr.fcba = dtl.fcba)" & _
							" WHERE dtl.fcba = '" & fcba & "'" & _
							"       AND dtl.agreementcode = '" & FCCODE & "'" & _
							"       AND dtl.lineno = '" & lineno &"'"
							
							
				Set rsCheckRate = Server.CreateObject("ADODB.Recordset")
				rsCheckRate.Open checkRate, connstr

				if Not rsCheckRate.Eof then
					varRate = cdbl(to_nol(rsCheckRate("rate")))
					varDueDate = rsCheckRate("finishdate")
				Else
					varRate = 0
					varDueDate = "1/1/0001"
				End if		

				rsCheckRate.Close
				set rsCheckRate = Nothing		

				
				strCheckDate = "select to_number(to_date('" &  WIPDATE & "','MM/DD/YYYY') - to_date('" &  varDueDate & "','MM/DD/YYYY')) as RemainDate from dual "
					
				Set rsCheckDate = Server.CreateObject("ADODB.Recordset")
				rsCheckDate.Open strCheckDate, connstr

				if Not rsCheckDate.Eof then
					varRemain 	= 	cdbl(to_nol(rsCheckDate("RemainDate")))
				Else
					varRemain 	= 	0
				End if	
				
				rsCheckDate.Close
				set rsCheckDate = Nothing
				
				maxlineno = 0
				strCheckDate = " SELECT NVL(MAX(lineno),0) +1 lineno FROM workinprogressplan WHERE agreementcode = '" & FCCODE & "'	AND fcba = '" & fcba & "' AND agreementlineno = '" & lineno & "'"
					
				Set rsCheckDate = Server.CreateObject("ADODB.Recordset")
				rsCheckDate.Open strCheckDate, connstr

				if Not rsCheckDate.Eof then
					maxlineno = cdbl(to_nol(rsCheckDate("lineno")))
				Else
					maxlineno = 0
				End if	
				
				rsCheckDate.Close
				set rsCheckDate = Nothing

				if varRemain <= 0 then		
					q = " INSERT INTO workinprogressplan ( agreementcode" & _
                        "                                  , wipdate" & _
                        "                                  , agreementlineno" & _
                        "                                  , lineno" & _
                        "                                  , qty" & _
                        "                                  , unit" & _
                        "                                  , output" & _
                        "                                  , postedflag" & _
                        "                                  , wipamt" & _
                        "                                  , reference" & _
                        "                                  , remarks" & _
						"                                  , chargeunit" & _
                        "                                  , jobcode" & _
                        "                                  , locationtype" & _
                        "                                  , locationcode" & _
						"                                  , fcentry" & _
                        "                                  , fcedit" & _
                        "                                  , fcip" & _
                        "                                  , fcba" & _
                        "                                  , lastupdate" & _
                        "                                  , lasttime )" & _
                        " VALUES ( '" & FCCODE & "'" & _
                        "          , to_date('" & WIPDATE & "','mm/dd/yyyy')" & _
                        "          , '" & lineno & "'" & _
                        "          , '" & maxlineno & "'" & _
                        "          , '" & QTY & "'" & _
                        "          , '" & UNIT & "'" & _
                        "          , '" & OUTPUT & "'" & _
                        "          , 0" & _
                        "          , '" & QTY * varRate & "'" & _
                        "          , '" & REFERENCE & "'" & _
                        "          , '" & REMARKS & "'" & _
						"          , '" & chargeunit & "'" & _
                        "          , '" & jobcode & "'" & _
                        "          , '" & locationtype & "'" & _
                        "          , '" & locationcode & "'" & _
                        "          , '" & Session("CodeName") & "'" & _
                        "          , '" & Session("CodeName") & "'" & _
                        "          , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _
                        "          , '" & fcba & "'" & _
                        "          , SYSDATE" & _
                        "          , TO_CHAR(SYSDATE,'HH24:MI') )"

						'Response.Write q	
						connStr.execute(q)
						cmdView	= "True"
						varError = ""
				else
					cmdView	= "True"
					varError = "Tanggal WIP harus lebih kecil dari Tanggal Jatuh Tempo pada Contract Aggreement "  
				end if
			else
				cmdView	= "True"
				varError = "Total Quantity yang anda masukan melebihi dari kontrak yang disepakati "  
			end if
			
			rsCheckQTY.Close
			Set rsCheckQTY = Nothing
		else
			err.number = 1001
			err.description = "Tanggal """ & WIPDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(fcba, WIPDATE)
		If Cek_IsDailyClosing = "FALSE" then
			'Validasi Qty
			CheckQTY = " SELECT fcba" & _
			           "        , agreementcode" & _
			           "        , NVL(SUM (totalagreement),0) totalagreement" & _
			           "        , NVL(SUM (totalwip),0) totalwip" & _
			           "        , NVL(SUM (totalagreement),0) - NVL(SUM (totalwip),0) selisih" & _
					   " FROM ( SELECT fcba, agreementcode, lineno, NVL(qty,0) totalagreement, 0 totalwip" & _
					   "        FROM contractagreementdetail WHERE agreementcode = '" & FCCODE & "' AND fcba = '" & fcba & "' AND lineno = '" & lineno & "'" & _
					   "        UNION ALL" &_
					   "        SELECT fcba, agreementcode, agreementlineno lineno, 0 totalagreement, NVL(qty,0) totalwip" & _
					   "        FROM workinprogressplan WHERE agreementcode = '" & FCCODE & "' AND fcba = '" & fcba & "' AND agreementlineno = '" & lineno & "' ) dataall" & _
					   " WHERE fcba = '" & fcba & "' AND agreementcode = '" & FCCODE & "'" & _
					   " AND lineno = '" & lineno & "' GROUP BY fcba, agreementcode"
			' response.write CheckQTY
			Set rsCheckQTY = Server.CreateObject("ADODB.Recordset")
			rsCheckQTY.Open CheckQTY, connstr
			
			if Not rsCheckQTY.Eof then
				varSelisih = cdbl(to_nol(rsCheckQTY("selisih")))
			Else
				varSelisih = 0		   
			End if
			
			checkRate = " SELECT dtl.rate, hdr.finishdate, dtl.lineno" & _
						" FROM contractagreementdetail dtl" & _
						" INNER JOIN contractagreement hdr ON (hdr.agreementcode = dtl.agreementcode AND hdr.fcba = dtl.fcba)" & _
						" WHERE dtl.fcba = '" & fcba & "'" & _
						"       AND dtl.agreementcode = '" & FCCODE & "'" & _
						"       AND dtl.lineno = '" & lineno &"'"
						
						
			Set rsCheckRate = Server.CreateObject("ADODB.Recordset")
			rsCheckRate.Open checkRate, connstr

			if Not rsCheckRate.Eof then
				varRate = cdbl(to_nol(rsCheckRate("rate")))
			Else
				varRate = 0
			End if	
		
			rsCheckRate.Close
			set rsCheckRate = Nothing
			
			CheckWIPQTY = " SELECT NVL(SUM(qty),0) qty" & _
                          " FROM workinprogressplan" & _
                          " WHERE agreementcode = '" & FCCODE & "'" & _
                          "       AND agreementlineno = '" & lineno & "'" & _
                          "       AND fcba = '" & fcba & "'" & _
                          "       AND wipdate = TO_DATE('" & WIPDATE & "','mm/dd/yyyy')" & _
                          "       AND lineno = '" & maxlineno & "'"
			
			Set rsCheckWIPQTY = Server.CreateObject("ADODB.Recordset")
			rsCheckWIPQTY.Open CheckWIPQTY, connstr
			
			if Not rsCheckWIPQTY.Eof then
				varWIPQTY = cdbl(to_nol(rsCheckWIPQTY("qty")))
			Else
				varWIPQTY = 0
			End if

			if ((varSelisih + varWIPQTY) - cdbl(to_nol(QTY))) >= 0 then			
				
				q = " UPDATE WORKINPROGRESSPLAN" & _
                    " SET jobcode = '" & jobcode & "'" & _
					"     ,locationtype = '" & locationtype & "'" & _
					"     ,locationcode = '" & locationcode & "'" & _
                    "     , qty = '" & qty & "'" & _
                    "     , unit = '" & unit & "'" & _
                    "     , output = '" & output & "'" & _
                    "     , wipamt = '" & cdbl(to_nol(qty)) *cdbl(to_nol(varRate)) & "'" & _
                    "     , reference = '" & reference & "'" & _
                    "     , remarks = '" & remarks & "'" & _
					"     , chargeunit = '" & chargeunit & "'" & _
                    " WHERE agreementcode = '" & FCCODE & "'" & _
                    "       AND agreementlineno = '" & lineno & "'" & _
                    "       AND fcba = '" & fcba & "'" & _
                    "       AND wipdate = TO_DATE('" & WIPDATE & "','mm/dd/yyyy')" & _
                    "       AND lineno = '" & maxlineno & "'"
				' response.write q
				ConnStr.execute(q)
					
				cmdView	= "True"
				varError = ""
						
			else
				cmdView	= "True"
				varError = "Total Quantity yang anda masukan melebihi dari kontrak yang disepakati "  
			end if
			
			rsCheckQTY.Close
			Set rsCheckQTY = Nothing
		else
			err.number = 1001
			err.description = "Tanggal """ & WIPDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
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

If Len(Trim(fcCode)) > 0 Then

	Dim rsCekAvailableCode
	set rsCekAvailableCode = server.CreateObject("Adodb.recordset")

	q = " SELECT hdr.*, bu.fccompanycode" & _
		" FROM contractagreement hdr" & _
		" LEFT JOIN businessunit bu on bu.fccode = hdr.fcba" & _
		" WHERE hdr.agreementcode = '" & fcCode & "'"
		
	rsCekAvailableCode.Open q, connstr
	
	If Not rsCekAvailableCode.EOF then
		FCCODE = rsCekAvailableCode("AGREEMENTCODE")
		fccompanycode = rsCekAvailableCode("fccompanycode")
		fcba = rsCekAvailableCode("fcba")
		selectedfcba = rsCekAvailableCode("fcba")
	Else
		Response.Write "<div class = ""error"">" & "Kontrak yang di maksud, tidak ada dalam Master Kontrak. Cek Kembali Kontrak anda" & "</div>"
		CmdView = "False"
	End if

	rsCekAvailableCode.Close
	set rsCekAvailableCode = Nothing

End If

txtTable	= Request("txtTable")
txtModule	= Request("txtModule")
txtSearch	= Request("txtSearch")
txtField	= Request("txtField")
txtJenis	= request("txtJenis")

If Len(Trim(varError)) > 0 Then
	response.Write  "<div class = ""error"">" & varError & "</div>"
End If

if err.number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if
 '-------------ISI Payment Form
 %>
 <form name="frmEntry" method="post" action="WIPPLAN.asp ">
	<table class="data" width="100%">
		<tr>
			<td  align="left" width="205">Agreement Code</td>
			<td align="left">
				<input type="text" maxlength="150" id="agreementcode" name="fcCODE" size="30" value="<%=fcCODE%>"><font color="#FF0000"> *</font>
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
				<font color="#FF0000"> *</font>
			</td>
		</tr>
		<tr>
			<td align="left">Business Unit</td>
			<td align="left">
				<select id="fcba" name="fcba" onchange="javascript:getselectedfcba();javascript:getlistcontractor();">
					<option value=""></option>
				</select>
				<font color="#FF0000"> *</font>
			</td>
			<input type="Hidden" id="selectedfcba" name="selectedfcba" value="<%=selectedfcba%>">
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<input class="datepicker" type="text" size="15" name="wipdate" value="<%=wipdate%>"><font color="#FF0000"> *</font>
			</td>
		</tr>
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
				<a href="PrintWIP.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
			</td>
		</tr>
    </table>  
 </form>
<%
	If cmdView = "True" And Len(Trim(""&fcCode)) > 0 And Len(Trim(""&fccompanycode)) > 0 And Len(Trim(""&fcba)) > 0 And Len(Trim(""&wipdate)) > 0 Then
		if request("txtJenis")="Edit" then
			q = " SELECT *" & _
			    " FROM workinprogressplan" & _
			    " WHERE agreementcode = '" & fcCode & "'" & _
			    "       AND agreementlineno = '" & agreementlineno & "'" & _
			    "       AND fcba = '" & fcba & "'" & _
			    "       AND wipdate = TO_DATE('" & wipdate & "', 'mm/dd/yyyy')" & _
				"       AND lineno = '" & lineno & "'"
			
			' response.Write q
			Set rsWC = Server.CreateObject("ADODB.Recordset")
			rsWC.Open q, connstr
			if not rsWC.eof then
			    chargeunit = rsWC("chargeunit")
				jobcode = rsWC("jobcode")
				locationtype = rsWC("locationtype")
				locationcode = rsWC("locationcode")
				qty = rsWC("qty")
				unit = rsWC("unit")
				output = rsWC("output")
				reference = rsWC("reference")
				remarks = rsWC("remarks")
			end if
		end if
	
		q = " SELECT agreementcode" & _
			"        , wipdate" & _
			"        , agreementlineno" & _
			"        , lineno" & _
			"        , chargeunit" & _
			"        , jobcode" & _
			"        , locationtype" & _
			"        , locationcode" & _
			"        , qty" & _
            "        , unit" & _
			"        , output" & _
			"        , reference" & _
			"        , remarks" & _
			" FROM workinprogressplan" & _
			" WHERE agreementcode = '" & fcCode & "'" & _
			"       AND fcba = '" & fcba & "'" & _
			"       AND wipdate = TO_DATE('" & wipdate & "', 'mm/dd/yyyy')" & _
			"       AND isgenerated IS NULL " & _ 
			" ORDER BY agreementcode, wipdate, agreementlineno, lineno"
			
		'response.Write q
		Set rsWC = Server.CreateObject("ADODB.Recordset")
		rsWC.Open q, connstr
%>
	<form action="WIPPLAN.asp " method="post" name="eForm">
		<table class="data" cellpadding="0" cellspacing="0" width="100%">
		
 		    <tr>
				<td align="left">Lineno</td>
				<td align="left">
					<select id="plineno" name="lineno" onchange="javascript:getcontractinfo();" onblur="javascript:getcontractinfo();" <%IF Request("txtJenis")="Edit" then%> disabled <%end if%>>
					<option value=""></option>
					<%
					qlov = "SELECT dtl.lineno fccode FROM contractagreementdetail dtl LEFT JOIN job ON (dtl.jobcode = job.fccode AND dtl.fcba = job.fcba) WHERE dtl.agreementcode = '" & fcCODE & "' AND dtl.fcba = '" & fcba & "' ORDER BY dtl.lineno"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if cstr(""&lineno) = cstr(""&rslov("fccode")) then%>selected<%end if%>><%=rslov("fccode")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
					</select>
					<font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Charge Unit </td>
				<td align="left">
					<select  name="chargeunit" id = "chargeunit" >
							<option value=""></option>
							<%
							qlov = 	 " select fccode,fcname " & _
									 "	from businessunit " & _
									 "	where fccompanycode = businessunit_api.get_CompanyCode('" & fcba & "') " & _
									 "	and businessunit_api.Isactive(fccode)='TRUE' " 
									 
							Set rslov = Server.CreateObject("ADODB.Recordset")
							rslov.Open qlov, connstr 
							if Not rslov.eof then
								While Not rslov.EOF
							%>
							<option value="<%=rslov("fccode")%>" <%if chargeunit = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%> </option>
							<%
									rslov.MoveNext
								Wend
							end if
							%>
					</select>
					<font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Job Code</td>
				<td align="left">
					<input type="text" size="30" id="jobcode" name="jobcode" value="<%=jobcode%>" style="background-color:#eee;" readonly><font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Job Desc</td>
				<td align="left">
					<input type="text" size="30" id="jobname" name="jobname" value="<%=jobname%>" style="background-color:#eee;" readonly><font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Rate</td>
				<td align="left">
					<input type="text" size="30" id="rate" name="rate" value="<%=rate%>" style="background-color:#eee;" readonly><font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Contract Ref</td>
				<td align="left">
					<input type="text" size="30" id="contract_ref" name="contract_ref" value="<%=contract_ref%>" style="background-color:#eee;" readonly><font color="#FF0000"> *</font>
				</td>
			</tr>
			
			<tr>
				<td align="left">Target Code</td>
				<td align="left">
					<input type="text" size="15" name="locationcode" value="<%=locationcode%>" ><font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Quantity</td>
				<td align="left">
					<input type="text" size="15" name="QTY" value="<%=QTY%>" onkeyup="checkInput(this);"><font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Unit</td>
				<td align="left">
					<input type="text" size="15" name="UNIT" value="<%=UNIT%>" onkeyup="checkInput(this);"><font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Output</td>
				<td align="left">
					<input type="text" size="15" name="OUTPUT" value="<%=OUTPUT%>" onkeyup="checkInput(this);"><font color="#FF0000"> *</font>
				</td>
			</tr>
			<tr>
				<td align="left">Reference</td>
				<td align="left">
					<input type="text" size="30" maxlength = "30" name="REFERENCE" value="<%=REFERENCE%>">
				</td>
			</tr>
			<tr>
				<td align="left">Remarks</td>
				<td align="left">
					<textarea  name="REMARKS" maxlength = "100" cols="50" rows="3"><%=REMARKS%></textarea>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="left">
					<input type="hidden" name="fcCode" value="<%=fcCode%>">
					<input type="hidden" name="wipdate" value="<%=wipdate%>">
					
					<input type="hidden" name="fccompanycode" value="<%=fccompanycode%>">
					<input type="hidden" name="fcba" value="<%=fcba%>">
					<input type="hidden" name="selectedfcba" value="<%=selectedfcba%>">
					
					<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
					<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
					<%IF Request("txtJenis")="Edit" then%>
						<input type="hidden" name="maxlineno" value="<%=maxlineno%>">
						<input type="hidden" name="lineno" value="<%=lineno%>">
					
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
		<br/>
		<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
			<TD align="left">No</TD>
			<TD colspan="2">&nbsp;</TD>
			<TD align="left">Agreement Code</TD>
			<TD align="left">WIP Date</TD>
			<TD align="left">Charge Unit</TD>
			<TD align="left">Jobcode</TD>
			<TD align="left">Target Type</TD>
			<TD align="left">Target Code</TD>
			<TD align="left">Quantity</TD>
			<TD align="left">Unit</TD>
			<TD align="left">Output</TD>
			<TD align="left">Reference</TD>
			<TD align="left">Remarks</TD>
		</tr>   
   <%
   if Not rsWC.eof then
      nCountofField=rsWC.fields.count-1
	  no=1
	  While Not rsWC.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
			<td align="left" valign="top"><%=NO%></td>
			<td valign="middle">
				<a href="WIPPLAN.asp?fcCode=<%=rsWC("AGREEMENTCODE")%>&lineno=<%=rsWC("AGREEMENTLINENO")%>&wipdate=<%=wipdate%>&maxlineno=<%=rsWC("lineno")%>&QTY=<%=rsWC("QTY")%>&UNIT=<%=rsWC("UNIT")%>&OUTPUT=<%=rsWC("OUTPUT")%>&REFERENCE=<%=rsWC("REFERENCE")%>&REMARKS=<%=rsWC("REMARKS")%>&locationcode=<%=rsWC("locationcode")%>&fcba=<%=fcba%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsWC(1)%>" border="0"></a>
			</td>
			<td>
				<a href="WIPPLAN.asp?fcCode=<%=rsWC("AGREEMENTCODE")%>&lineno=<%=rsWC("AGREEMENTLINENO")%>&wipdate=<%=wipdate%>&maxlineno=<%=rsWC("lineno")%>&fcba=<%=fcba%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsWC(1)%>" border="0"   onClick="return confirm('Are you sure to delete');"></a>
			</td>
			
			<td align="left" title="<%=rsWC("agreementcode").Name%>" valign="top"><%=rsWC("agreementcode")%></td>
			<td align="left" title="<%=rsWC("wipdate").Name%>" valign="top"><%=rsWC("wipdate")%></td>
			<td align="left" title="<%=rsWC("Chargeunit").Name%>" valign="top"><%=rsWC("Chargeunit")%></td>
			<td align="left" title="<%=rsWC("jobcode").Name%>" valign="top"><%=rsWC("jobcode")%></td>
			<td align="left" title="<%=rsWC("locationtype").Name%>" valign="top"><%=rsWC("locationtype")%></td>
			<td align="left" title="<%=rsWC("locationcode").Name%>" valign="top"><%=rsWC("locationcode")%></td>
			<td align="left" title="<%=rsWC("qty").Name%>" valign="top"><%=rsWC("qty")%></td>
			<td align="left" title="<%=rsWC("unit").Name%>" valign="top"><%=rsWC("unit")%></td>
			<td align="left" title="<%=rsWC("output").Name%>" valign="top"><%=rsWC("output")%></td>
			<td align="left" title="<%=rsWC("reference").Name%>" valign="top"><%=rsWC("reference")%></td>
			<td align="left" title="<%=rsWC("remarks").Name%>" valign="top"><%=rsWC("remarks")%></td>
		  </tr>
		  <%
		  rsWC.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsWC.Close
   set rsWC = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <!-------------akhir!------------->
</div></div>
  <!--#include file="bottom.jsp"-->
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
	getcontractinfo();
});
</script>
</body>
</html>
