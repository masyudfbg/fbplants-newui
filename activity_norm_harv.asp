<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Activity Norm Harvesting</title>
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

</script>
<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="plugins/responsivegridsystem/12cols.css" media="all">
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
var getuomjob = function () {
	// build up the post string when passing variables to the server side page
	var controljob = document.getElementById("controljob").value;
	var PostStr = 'controljob='+controljob;
	// alert(PostStr);
	// use the generic function to make the request
	doAJAXCall('ajax/getuomjob.asp', 'POST', PostStr, showMessageResponse);
	
}

// The function for handling the response from the server
var showMessageResponse = function (oXML) { 
    
    // get the response text, into a variable
    var response = oXML.responseText;
    var s = JSON.parse(response);
    // update the Div to show the result from the server
	document.getElementById("uomjob").value = s.uomjob;
}; 
</script>
<script>
	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
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
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
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
 
<%

 '-------------awal----------------------- 

Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim CUSTOMERCODE, txtDESCRIPTION, SUPPLIER, ICNO, BANK, BANKACCOUNTNO, COMPANYNO,OWN_ESTATE,CUSTOMER_FCBA,CUSTFFBTYPE,ISTRANSPORTER
Dim cmdView,txt_CUSTOMERCODE
Dim vSTREET, vCITY, vPOSTALCODE, vPROVINCE, vTELEPHONE, vFAX, vEMAIL, vISACTIVE, vISBLOCK, vISTRANSPORTER
Dim hide_ISACTIVE, hide_ISBLOCK, hide_ISTRANSPORTER
txttable=request("txttable")

plantingyear=request("plantingyear")
rotasi=request("rotasi")
basis_jjg=request("basis_jjg")
rate_sb=request("rate_sb")
lb_rate=request("lb_rate")
norma_hk=request("norma_hk")
rate1=request("rate1")

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM activity_norm_harv WHERE  plantingyear='" & plantingyear & "' AND fcBA='" & Session("fcBA") & "'"
		connStr.execute(q)
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		q	=	" INSERT INTO activity_norm_harv (" & _
				" plantingyear, " & _
				" rotasi, " & _
				" basis_jjg, " & _
				" rate_sb, " & _
				" lb_rate, " & _
				" norma_hk, " & _
				" rate1, " & _
				" FCENTRY, " & _
				" FCEDIT, " & _
				" FCIP, " & _
				" FCBA, " & _
				" LASTUPDATE, " & _
				" LASTTIME" & _
				" ) values (" & _
				"   '" & plantingyear & "'" & _			
				" , '" & rotasi & "'" & _
				" , '" & basis_jjg & "'" & _
				" , '" & rate_sb & "'" & _
				" , '" & lb_rate & "'" & _
				" , '" & norma_hk & "'" & _
				" , '" & rate1 & "'" & _
				" , '" & Session("CodeName") & "'" & _					
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") & "'" & _					
				" , '" & Session("fcBA") & "'" & _					
				" , sysdate " & _										
				" , To_Char(sysdate,'HH24:MI') " & _	
				" )"			 

		''response.write q
		'response.Flush()
		connStr.execute(q)
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q=	"  UPDATE activity_norm_harv SET" & _ 
			"  rotasi ='" & rotasi & "'" & _
			"  ,basis_jjg ='" & basis_jjg & "'" & _
			"  ,rate_sb ='" & rate_sb & "'" & _
			"  ,lb_rate ='" & lb_rate & "'" & _
			"  ,norma_hk ='" & norma_hk & "'" & _
			"  ,rate1 ='" & rate1 & "'" & _
			"  ,FCEDIT = '"& Session("CodeName") &"'" & _
			"  ,LASTUPDATE = sysdate" & _
			"  ,LASTTIME = To_Char(sysdate,'HH24:MI')" & _
			"  WHERE  plantingyear = '" & plantingyear &"'"  & _
			"  AND fcBA='" & Session("fcBA") & "'" 
			
		connStr.execute(q)
		connStr.execute("commit")
		''response.Write(q)
		'response.Flush()
		''response.write q			
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
  
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-sliders" aria-hidden="true"></i>Header Activity Norm Harvesting</div><% end if %>
<form name="frmEntry" method="post" action="activity_norm_harv.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
        <tr>
		    <td colspan="2">
		        <div align="left">
		            <font size="2" face="Verdana"><strong>MAINTENANCE ACTIVITY NORM HARVESTING</strong></font>
		        </div>
		    </td>
		</tr>
        <tr>
			<td align="left" width="15%">Planting Year</td><td align="left"><input type="text" size="15" name="plantingyear" value="<%=request("plantingyear")%>"  ></td>
		</tr>	
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;
		<a href="Printactivity_norm_harv.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&plantingyear=<%=REQUEST("plantingyear")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
		</tr>
    </table>  
 </form>
 <%
 if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   ' set rsGAD = server.CreateObject("Adodb.recordset")
    'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	
	''SUPPLIER,OWN_ESTATE, CUSTFFBTYPE, ISTRANSPORTER,
	'' Query Menampilkan Data di Table Bawah  
	q	=	" SELECT " & _
			" plantingyear, rotasi, basis_jjg, rate_sb, lb_rate, norma_hk, rate1 " & _
			" FROM activity_norm_harv " & _
			" WHERE fcba = '" & Session("fcBA") & "'" & _       
			" AND plantingyear ='"& request("plantingyear") &"' "	
		
    Set rsGAD = Server.CreateObject("ADODB.Recordset")
	' if  len(trim(request("CUSTOMERCODE"))) > 0 then
		' q = q & " and CUSTOMERCODE='" & request("CUSTOMERCODE") & "'"
	' end if
	' q = q
   
   ' response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr 
   
  %>
   
   <!-- Form Entry And Edit --> 
   
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="activity_norm_harv.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
		<%IF Request("txtJenis")="Edit" then%>
			<tr>
			    <td align="left"><font face="Verdana" size="1">Planting Year</font></td>
			    <td align="left"><input type="text" size="30" maxlength="100" value="<%=plantingyear%>" name="plantingyear" readonly style="background:#eee;"  title="DESCRIPTION Character a s/d z or 1 s/d 9"> <font color="#FF0000">*</font></td>
			</tr>
		<%ELse%>
			<tr>
			    <td align="left"><font face="Verdana" size="1">Planting Year</font></td>
			    <td align="left"><input type="text" size="30" maxlength="100" value="<%=plantingyear%>" name="plantingyear" onkeyup="hanyaAngka(this)" style="background:#fff;"  title="DESCRIPTION Character a s/d z or 1 s/d 9"> <font color="#FF0000">*</font></td>
			</tr>
		<%End If%>
		<tr>
		    <td align="left"><font face="Verdana" size="1">Rotasi</font></td>
		    <td align="left"><input type="text" size="30" maxlength="100" value="<%=rotasi%>" name="rotasi" onkeyup="hanyaAngka(this)" style="background:#fff;"  title="DESCRIPTION Character a s/d z or 1 s/d 9"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		    <td align="left"><font face="Verdana" size="1">Basis JJG</font></td>
		    <td align="left"><input type="text" size="30" maxlength="100" value="<%=basis_jjg%>" name="basis_jjg" onkeyup="hanyaAngka(this)" style="background:#fff;"  title="DESCRIPTION Character a s/d z or 1 s/d 9"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		    <td align="left"><font face="Verdana" size="1">Rate (SB)</font></td>
		    <td align="left"><input type="text" size="30" maxlength="100" value="<%=rate_sb%>" name="rate_sb" onkeyup="hanyaAngka(this)" style="background:#fff;"  title="DESCRIPTION Character a s/d z or 1 s/d 9"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		    <td align="left"><font face="Verdana" size="1">Rate (LB)</font></td>
		    <td align="left"><input type="text" size="30" maxlength="100" value="<%=lb_rate%>" name="lb_rate" onkeyup="hanyaAngka(this)" style="background:#fff;"  title="DESCRIPTION Character a s/d z or 1 s/d 9"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		    <td align="left"><font face="Verdana" size="1">Norma HK</font></td>
		    <td align="left"><input type="text" size="30" maxlength="100" value="<%=norma_hk%>" name="norma_hk" onkeyup="hanyaAngka(this)" style="background:#fff;"  title="DESCRIPTION Character a s/d z or 1 s/d 9"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		    <td align="left"><font face="Verdana" size="1">Rate A01</font></td>
		    <td align="left"><input type="text" size="30" maxlength="100" value="<%=rate1%>" name="rate1" onkeyup="hanyaAngka(this)" style="background:#fff;"  title="DESCRIPTION Character a s/d z or 1 s/d 9"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				  <input type="Hidden" name="txt_CUSTOMERCODE" value="<%=REQUEST("CUSTOMERCODE")%>">
				<!--
				  <input type="text" name="hide_ISTRANSPORTER" value="<%=REQUEST("ISTRANSPORTER")%>">	
				  <input type="text" name="hide_ISBLOCK" 	   value="<%=REQUEST("ISBLOCK")%>">	
				  <input type="text" name="hide_ISACTIVE" 	   value="<%=REQUEST("ISACTIVE")%>">	
				-->
				
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
  <!--#include file="validation_show.Asp"-->    
<%

	Dim intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
	howmanyfields=rsGAD.fields.count-1
	
	JumRec=1
	intRecordsPerPage = 20
	
	rsGAD.PageSize = intRecordsPerPage
	intNumberOfPages = 0
	intNumberOfPages =  rsGAD.PageCount
	intRecords = rsGAD.RecordCount
	'Set our AbsolutePage property.
	if not rsGAD.EOF then
		If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
		  rsGAD.AbsolutePage = 1
		Else
		  rsGAD.AbsolutePage = Request("iPageNo")
		End If  
	end if
	
	if Request("iPageNo") = "" then
		nPage = 1
	Else
		nPage = Request("iPageNo")
	End if
	
   rsGAD.PageSize = intRecordsPerPage
   intNumberOfPages = 0
   intNumberOfPages =  rsGAD.PageCount
   intRecords = rsGAD.RecordCount   
   %>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines" >
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
		<TD align="left">NO</TD>
		<TD align="center"colspan="2">ACTION</TD>
		<TD align="left">PLANTING YEAR</TD>	
		<TD align="left">ROTASI</TD>
		<TD align="left">BASIS JJG</TD>	 
		<TD align="left">RATE (SB)</TD>	 
		<TD align="left">RATE (LB)</TD>	 
		<TD align="left">NORMA HK</TD>	 
		<TD align="left">RATE (A01)</TD>	 
	</tr>
	</thead>
	<tbody>
   
   <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count - 5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="activity_norm_harv.asp?plantingyear=<%=rsGAD("plantingyear")%>&rotasi=<%=rsGAD("rotasi")%>&basis_jjg=<%=rsGAD("basis_jjg")%>&rate_sb=<%=rsGAD("rate_sb")%>&lb_rate=<%=rsGAD("lb_rate")%>&norma_hk=<%=rsGAD("norma_hk")%>&rate1=<%=rsGAD("Rate1")%>&txttable=<%=request("txttable")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>		
			</td>
			<td valign="middle" width="30">
			<a href="activity_norm_harv.asp?plantingyear=<%=rsGAD("plantingyear")%>&rotasi=<%=rsGAD("rotasi")%>&basis_jjg=<%=rsGAD("basis_jjg")%>&rate_sb=<%=rsGAD("rate_sb")%>&lb_rate=<%=rsGAD("lb_rate")%>&norma_hk=<%=rsGAD("norma_hk")%>&rate1=<%=rsGAD("Rate1")%>&txttable=<%=request("txttable")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
			<td align="left" title="<%=rsGAD("plantingyear").Name%>" valign="top"><%=rsGAD("plantingyear")%></td>
			<td align="left" title="<%=rsGAD("rotasi").Name%>" valign="top"><%=rsGAD("rotasi")%></td>
			<td align="left" title="<%=rsGAD("basis_jjg").Name%>" valign="top"><%=rsGAD("basis_jjg")%></td>
			<td align="left" title="<%=rsGAD("rate_sb").Name%>" valign="top"><%=rsGAD("rate_sb")%></td>
			<td align="left" title="<%=rsGAD("lb_rate").Name%>" valign="top"><%=rsGAD("lb_rate")%></td>
			<td align="left" title="<%=rsGAD("norma_hk").Name%>" valign="top"><%=rsGAD("norma_hk")%></td>
			<td align="left" title="<%=rsGAD("Rate1").Name%>" valign="top"><%=rsGAD("Rate1")%></td>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
   </tbody>
   </TABLE>
   </div>
<%End if%>  
 
  <%'-------------akhir----------------------%>
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
	getuomjob(); 
});
</script>
</body>
</html>
