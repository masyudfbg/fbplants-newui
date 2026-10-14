<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Vehicle Activity</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<script src="js/jquery-3.1.0.min.js"></script>
<script src="js/jquery.maskedinput.min.js"></script>
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
<script language="javascript" type="text/javascript">



                function showHint(str)
					{
				var hoursVehicle;
				var vehiclecode_;	
				
				vehiclecode_ = document.getElementById("combo_vehiclecode").value;
				
					if (str=="")
					  {
					  document.getElementById("hours").innerHTML="";
					  document.getElementById("unit").innerHTML="";
					  document.getElementById("output").innerHTML="";
					  
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  hoursVehicle=new XMLHttpRequest();
					  unitVehicle=new XMLHttpRequest();
					  outputVehicle=new XMLHttpRequest();

					  }
					else
					  {// code for IE6, IE5
					  hoursVehicle=new ActiveXObject("Microsoft.XMLHTTP");
					  unitVehicle=new ActiveXObject("Microsoft.XMLHTTP");
					  outputVehicle=new ActiveXObject("Microsoft.XMLHTTP");

					  }
					  
						hoursVehicle.onreadystatechange=function()
						  {
						  if (hoursVehicle.readyState==4 && hoursVehicle.status==200)
							{
							document.getElementById("hours").innerHTML=hoursVehicle.responseText;
							}
						  }
						hoursVehicle.open("GET","ajax/Vehicle_ann.asp?ajaxUoM=hours&FcCode="+ vehiclecode_ +"&q="+str,true);
						hoursVehicle.send();
						
						unitVehicle.onreadystatechange=function()
						  {
						  if (unitVehicle.readyState==4 && unitVehicle.status==200)
							{
							document.getElementById("unit").innerHTML=unitVehicle.responseText;
							}
						  }
						unitVehicle.open("GET","ajax/Vehicle_ann.asp?ajaxUoM=unit&FcCode="+ vehiclecode_ +"&q="+str,true);
						unitVehicle.send();
						
						outputVehicle.onreadystatechange=function()
						  {
						  if (outputVehicle.readyState==4 && outputVehicle.status==200)
							{
							document.getElementById("output").innerHTML=outputVehicle.responseText;
							}
						  }
						outputVehicle.open("GET","ajax/Vehicle_ann.asp?ajaxUoM=output&FcCode="+ vehiclecode_ +"&q="+str,true);
						outputVehicle.send();
	
					}
</script>
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

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
	var GetEmpName = function () {
		// build up the post string when passing variables to the server side page
		// var suppliercode = document.getElementById("suppliercode").value;
		var empcode = document.getElementById("empcode").value;
		var PostStr = 'validate=empcd&q='+empcode;
		// alert(PostStr);
		// use the generic function to make the request
		doAJAXCall('ajax/validate.asp', 'POST', PostStr, showMessageResponse);
		
	}

	// The function for handling the response from the server
	var showMessageResponse = function (oXML) { 
	    
	    // get the response text, into a variable
	    var response = oXML.responseText;
	    // alert(s.supplieraddress);
	    // update the Div to show the result from the server
	    // document.getElementById("suppliername").value = s.suppliername;
		// document.getElementById("supplieraddress").value = s.supplieraddress;
		document.getElementById("nonscript_empcd").innerHTML = response;
		// alert(s.suppliername);
	    
	}
	// launched from button click 
	var GetHelper = function () {
		// build up the post string when passing variables to the server side page
		// var suppliercode = document.getElementById("suppliercode").value;
		var helper = document.getElementById("helper").value;
		var PostStr = 'validate=empcd&q='+helper;
		// alert(PostStr);
		// use the generic function to make the request
		doAJAXCall('ajax/validate.asp', 'POST', PostStr, showHelper);
		
	}

	// The function for handling the response from the server
	var showHelper = function (oXML) { 
	    
	    // get the response text, into a variable
	    var response = oXML.responseText;
	    // alert(s.supplieraddress);
	    // update the Div to show the result from the server
	    // document.getElementById("suppliername").value = s.suppliername;
		// document.getElementById("supplieraddress").value = s.supplieraddress;
		document.getElementById("nonscript_empcd_helper").innerHTML = response;
		// alert(s.suppliername);
	    
	}
	// launched from button click 
	var GetHelpers = function () {
		// build up the post string when passing variables to the server side page
		// var suppliercode = document.getElementById("suppliercode").value;
		var helpers = document.getElementById("helpers").value;
		var PostStr = 'validate=empcd&q='+helpers;
		// alert(PostStr);
		// use the generic function to make the request
		doAJAXCall('ajax/validate.asp', 'POST', PostStr, showHelpers);
		
	}

	// The function for handling the response from the server
	var showHelpers = function (oXML) { 
	    
	    // get the response text, into a variable
	    var response = oXML.responseText;
	    // alert(s.supplieraddress);
	    // update the Div to show the result from the server
	    // document.getElementById("suppliername").value = s.suppliername;
		// document.getElementById("supplieraddress").value = s.supplieraddress;
		document.getElementById("nonscript_empcd_helpers").innerHTML = response;
		// alert(s.suppliername);
	    
	}
	var getTragettype = function () {
		// build up the post string when passing variables to the server side page
		var jobcode = document.getElementById("jobcode").value;
		var PostStr = 'jobcode='+jobcode;
		// alert(PostStr);
		// use the generic function to make the request
		doAJAXCall('ajax/GetPVInfo.asp', 'POST', PostStr, showTragettype);
		
	}

	// The function for handling the response from the server
	var showTragettype = function (oXML) { 
	     // alert("1");
	    // get the response text, into a variable
	    var response = oXML.responseText;
	    var s = JSON.parse(response);
	    // alert(response);
	    // update the Div to show the result from the server
	    document.getElementById("locationtype").value = s.target; 
		// document.getElementById("descpr").value = s.description; 
	}
</script>
<script language="JavaScript">
        function PopupTType() {
			var txtttype = document.getElementById("locationtypecode").value;
			var txtdate  = document.getElementById("vehdate").value;
			var newwin   = window.open("popup_loccode.Asp?txtSearch="+txtttype+"&txtdate="+txtdate+"&Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
        }

		function PopupTypeCC() {
			var chargeunit = $("#chargeunit_").val();
			txtttype = document.getElementById("chargetype_").value;
			var newwin   = window.open("popup_loccode_cc.Asp?txtSearch="+txtttype+"&chargeunit=" + chargeunit,"newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
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
 <!-------------awal------------->
 <%

Function CheckCharge(initType, initCode, initJob, initFcba, initDate, initSource)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	"SELECT F_CekTarget_byDate ('"& initType &"'," & _
					"                           '"& initCode &"'," & _
					"                           '"& initJob &"'," & _
					"                           '"& initFcba &"'," & _
					"                           TO_DATE ('"& initDate &"', 'MM/DD/YYYY')," & _
					"                           'VH') Total" & _
					"  FROM DUAL"

	 'response.write stravailable
	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = rsAvailable("Total")
	Else
		initialAvailable = 0
	End If
	strAvailable = ""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckCharge = initialAvailable
End Function

Function CheckAvailableEmployee(initialEmployee, initialFcba)
	Dim initialAvailable
	Dim rsAvailableEmployee, strAvailableEmployee
	
	strAvailableEmployee =	" Select Count(1) Total From Employee Where" & _
							" Fccode = '" & initialEmployee & "'" & _
							" And Fcba = '" & initialFcba & "'"
							
	Set rsAvailableEmployee = Server.CreateObject("ADODB.Recordset")
	rsAvailableEmployee.Open strAvailableEmployee, connstr, 3
	
	If Not rsAvailableEmployee.EOF Then
		initialAvailable = Cdbl(rsAvailableEmployee("Total"))
	Else
		initialAvailable = 0
	End If
	
	rsAvailableEmployee.Close
	set rsAvailableEmployee = Nothing

	CheckAvailableEmployee = initialAvailable
End Function

 Dim FCCODE,Hours,Time
 Dim FDDATE,PeriodDate, JOB
 Dim LOCATIONTYPE, LOCATIONcode
 Dim LOCATION,UNIT, OUTPUT, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA
 Dim cmdView
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 Dim Cek_IsAccClosing : Cek_IsAccClosing = ""
 Dim ACCOUNTINGYEAR, ACCOUNTINGPERIOD,ACCOUNTINGYEARPERIOD
 Dim CATEGORYTRANSPORT,TRIP
 
fddate =request("fdDate")
fcCode =request("fcCode")
time=request("time")
LOCATIONTYPE=request("LOCATIONTYPE")
LOCATIONcode=request("LOCATIONcode")
JOB=request("JOB")
hours=request("hours")
UNIT=request("UNIT")
OUTPUT=request("OUTPUT")
REMARKS=request("REMARKS")
reference=request("reference")
CATEGORYTRANSPORT = Request("CATEGORYTRANSPORT")
TRIP = Request("TRIP")
EMPLOYEECODE = request("EMPLOYEECODE")
HELPER1 = request("HELPER1")
HELPER2 = request("HELPER2")
PID = request("PID")
TYPE_KONTANAN = request("TYPE_KONTANAN")

' Added Variable ChargeJob, ChargeType, ChargeCode
chargeunit=Request("chargeunit")
chargejob=Request("chargejob")
chargetype=Request("chargetype")
chargecode=Request("chargecode")
businessunit=Request("chargeunit")
'businessunit=Request("businessunit")
' End Added

'response.write ("ata1 :" & businessunit)

 ACCOUNTINGPERIOD  = left(fddate,2)
 ACCOUNTINGYEAR    = right(fddate,4)

Sub DeleteDataVehicle
	q	=	" Delete from VehicleActivity where" & _
			" PID = '" & PID & "'" & _
			" and fcBA='" & Session("fcBA") & "'"
			
	' response.write q
	connStr.execute(q)
End Sub

Sub InsertDataVehicle

	qLineMax = " select count(1) fccode, fcba from intraco_tab where status='TRUE' and fccode = '"& JOB &"' group by fcba"			
	' response.write qLineMax
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		code = cdbl(to_nol(rsqLineMax("fccode")))
		'businessunit = rsqLineMax("fcba")
		businessunit = chargeunit
	end if 

	if code = 0 then
		if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 or len(trim(""&chargeunit)) <> 0 ) then
			err.number = 1001
			err.description = "Anda tidak harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
		else
			q	=	" Insert into VehicleACTIVITY (" & _
					" VEHICLECODE" & _
					" , PID" & _
					" , EMPLOYEECODE" & _
					" , HELPER1" & _
					" , HELPER2" & _
					" , VEHICLEDATE" & _
					" , TIME" & _
					" , JOBCODE" & _
					" , LOCATIONTYPE" & _
					" , LOCATIONCODE" & _
					" , chargetype" & _
		            " , chargecode" & _
		     		" , chargejob" & _
					" , chargeunit" & _
					" , CATEGORYTRANSPORT" & _
					" , TRIP" & _
					" , HOURS" & _
					" , UNIT" & _
					" , AMOUNT" & _
					" , OUTPUT" & _
					" , REF" & _
					" , REMARKS" & _
					" , TYPE_KONTANAN" & _
					" , FCENTRY" & _
					" , FCEDIT" & _
					" , FCIP" & _
					" , FCBA" & _
					" , LASTUPDATE" & _
					" , LASTTIME" & _
					" )" & _
					" values (" & _
					" '" & FCCODE & "'" & _
					" , LOWER(regexp_replace(rawtohex(sys_guid()), '([A-F0-9]{8})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{12})','\1-\2-\3-\4-\5'))" & _
					" , '" & EMPLOYEECODE & "'" & _
					" , '" & HELPER1 & "'" & _
					" , '" & HELPER2 & "'" & _
					" , to_date('" & fdDate & "','mm/dd/yyyy')" & _
					" , '" & TIME & "'" & _
					" , '" & JOB & "'" & _
					" , '" & LOCATIONTYPE & "'" & _
					" , '" & LOCATIONcode & "'" & _
					" , '" & chargetype & "'" & _ 
	         		" , '" & chargecode &  "'" & _
	         		" , '" & chargejob & "'" & _
					" , '" & chargeunit & "'" & _
					" , '" & CATEGORYTRANSPORT & "'" & _
					" , '" & TRIP & "'" & _
					" , '" & HOURS & "'" & _
					" , '" & UNIT & "'" & _
					" , '0'" & _
					" , '" & OUTPUT & "'" & _				
					" , '" & Trim(reference) & "'" & _
					" , '" & REMARKS & "'" & _
					" , '" & TYPE_KONTANAN & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					" , '" & Session("fcBA") & "'" & _
					" , sysdate" & _
					" , To_Char(sysdate,'HH24:MI')" & _
					" )"
					
			'response.Write(q)
			'response.Flush()
			connStr.execute(q)
			End If
		else
		if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 or len(trim(""&chargeunit)) = 0 ) then
			err.number = 1001
			err.description = "Anda harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
		else
			Dim val_Charge : val_Charge = 0
			val_Charge = CheckCharge(chargetype, chargecode, chargejob, businessunit, fdDate, sourcetype)
			If val_Charge <> "Success" Then
				Err.Number = 1001
				Err.Description = val_Charge
			Else
			q	=	" Insert into VehicleACTIVITY (" & _
					" VEHICLECODE" & _
					" , PID" & _
					" , EMPLOYEECODE" & _
					" , HELPER1" & _
					" , HELPER2" & _
					" , VEHICLEDATE" & _
					" , TIME" & _
					" , JOBCODE" & _
					" , LOCATIONTYPE" & _
					" , LOCATIONCODE" & _
					" , chargetype" & _
		            " , chargecode" & _
		     		" , chargejob" & _
					" , chargeunit" & _
					" , CATEGORYTRANSPORT" & _
					" , TRIP" & _
					" , HOURS" & _
					" , UNIT" & _
					" , AMOUNT" & _
					" , OUTPUT" & _
					" , REF" & _
					" , REMARKS" & _
					" , TYPE_KONTANAN" & _
					" , FCENTRY" & _
					" , FCEDIT" & _
					" , FCIP" & _
					" , FCBA" & _
					" , LASTUPDATE" & _
					" , LASTTIME" & _
					" )" & _
					" values (" & _
					" '" & FCCODE & "'" & _
					" , LOWER(regexp_replace(rawtohex(sys_guid()), '([A-F0-9]{8})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{12})','\1-\2-\3-\4-\5'))" & _
					" , '" & EMPLOYEECODE & "'" & _
					" , '" & HELPER1 & "'" & _
					" , '" & HELPER2 & "'" & _
					" , to_date('" & fdDate & "','mm/dd/yyyy')" & _
					" , '" & TIME & "'" & _
					" , '" & JOB & "'" & _
					" , '" & LOCATIONTYPE & "'" & _
					" , '" & LOCATIONcode & "'" & _
					" , '" & chargetype & "'" & _ 
	         		" , '" & chargecode &  "'" & _
	         		" , '" & chargejob & "'" & _
					" , '" & chargeunit & "'" & _
					" , '" & CATEGORYTRANSPORT & "'" & _
					" , '" & TRIP & "'" & _
					" , '" & HOURS & "'" & _
					" , '" & UNIT & "'" & _
					" , '0'" & _
					" , '" & OUTPUT & "'" & _				
					" , '" & Trim(reference) & "'" & _
					" , '" & REMARKS & "'" & _
					" , '" & TYPE_KONTANAN & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					" , '" & Session("fcBA") & "'" & _
					" , sysdate" & _
					" , To_Char(sysdate,'HH24:MI')" & _
					" )"
					
			'response.Write(q)
			'response.Flush()
			connStr.execute(q)
		End If
	End If
End If
End Sub

Sub UpdateDataVehicle
	qLineMax = " select count(1) fccode, fcba from intraco_tab where status='TRUE' and fccode = '"& JOB &"' group by fcba"			
	' response.write qLineMax
	set rsqLineMax = Server.CreateObject("ADODB.Recordset")
	rsqLineMax.Open qLineMax, connstr
	if not rsqLineMax.eof then
		code = cdbl(to_nol(rsqLineMax("fccode")))
		'businessunit = rsqLineMax("fcba")
		businessunit = chargeunit
	end if 
if code = 0 then
		if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 or len(trim(""&chargeunit)) <> 0 ) then
			err.number = 1001
			err.description = "Anda tidak harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
		else
		Call DeleteDataVehicle
		q	=	" Insert into VehicleACTIVITY (" & _
				" VEHICLECODE" & _
				" , PID" & _
				" , EMPLOYEECODE" & _
				" , HELPER1" & _
				" , HELPER2" & _
				" , VEHICLEDATE" & _
				" , TIME" & _
				" , JOBCODE" & _
				" , LOCATIONTYPE" & _
				" , LOCATIONCODE" & _
				" , chargetype" & _
	            " , chargecode" & _
	     		" , chargejob" & _
				" , chargeunit" & _
				" , CATEGORYTRANSPORT" & _
				" , TRIP" & _
				" , HOURS" & _
				" , UNIT" & _
				" , AMOUNT" & _
				" , OUTPUT" & _
				" , REF" & _
				" , REMARKS" & _
				" , TYPE_KONTANAN" & _
				" , FCENTRY" & _
				" , FCEDIT" & _
				" , FCIP" & _
				" , FCBA" & _
				" , LASTUPDATE" & _
				" , LASTTIME" & _
				" )" & _
				" values (" & _
				" '" & FCCODE & "'" & _
				" , LOWER(regexp_replace(rawtohex(sys_guid()), '([A-F0-9]{8})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{12})','\1-\2-\3-\4-\5'))" & _
				" , '" & EMPLOYEECODE & "'" & _
				" , '" & HELPER1 & "'" & _
				" , '" & HELPER2 & "'" & _
				" , to_date('" & fdDate & "','mm/dd/yyyy')" & _
				" , '" & TIME & "'" & _
				" , '" & JOB & "'" & _
				" , '" & LOCATIONTYPE & "'" & _
				" , '" & LOCATIONcode & "'" & _
				" , '" & chargetype & "'" & _ 
         		" , '" & chargecode &  "'" & _
         		" , '" & chargejob & "'" & _
				" , '" & chargeunit & "'" & _
				" , '" & CATEGORYTRANSPORT & "'" & _
				" , '" & TRIP & "'" & _
				" , '" & HOURS & "'" & _
				" , '" & UNIT & "'" & _
				" , '0'" & _
				" , '" & OUTPUT & "'" & _				
				" , '" & Trim(reference) & "'" & _
				" , '" & REMARKS & "'" & _
				" , '" & TYPE_KONTANAN & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Session("CodeName") & "'" & _
				" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
				" , '" & Session("fcBA") & "'" & _
				" , sysdate" & _
				" , To_Char(sysdate,'HH24:MI')" & _
				" )"
				
		'response.Write(q)
		'response.Flush()
		connStr.execute(q)
		End If
	Else
		Call DeleteDataVehicle
		Call InsertDataVehicle
	end if
End Sub
		

Cek_IsAccClosing = StatusAccClosingTgl(fddate,fcba) 
if 	Cek_IsAccClosing = "FALSE" then  
if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fdDate)
		If Cek_IsDailyClosing = "FALSE" then
			q	=	" Delete from VehicleActivity where" & _
					"  PID='" & PID & "'" & _
					" and fcBA='" & Session("fcBA") & "'" 
					
			' response.write q
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & fdDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fdDate)
		If Cek_IsDailyClosing = "FALSE" then
			If Len(Trim(EMPLOYEECODE)) > 0 Then
				If CheckAvailableEmployee(EMPLOYEECODE, Session("fcBA")) > 0 Then
					qLineMax = " select count(1) fccode, fcba from intraco_tab where status='TRUE' and  fccode = '"& JOB &"' group by fcba"			
					' response.write qLineMax
					set rsqLineMax = Server.CreateObject("ADODB.Recordset")
					rsqLineMax.Open qLineMax, connstr
					if not rsqLineMax.eof then
						code = cdbl(to_nol(rsqLineMax("fccode")))
						'businessunit = rsqLineMax("fcba")
						businessunit = chargeunit
					end if 

					if code = 0 then
						if ( len(trim(""&chargetype)) <> 0 or len(trim(""&chargecode)) <> 0 or len(trim(""&chargejob)) <> 0 or len(trim(""&chargeunit)) <> 0 ) then
							err.number = 1001
							err.description = "Anda tidak harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
						else
							If len(trim(""&unit)) = 0 Then
								err.number = 1001
								err.description = "Anda Harus Mengisi Kolom Unit! "
							Else
								q	=	" Insert into VehicleACTIVITY (" & _
										" VEHICLECODE" & _
										" , PID" & _
										" , EMPLOYEECODE" & _
										" , HELPER1" & _
										" , HELPER2" & _
										" , VEHICLEDATE" & _
										" , TIME" & _
										" , JOBCODE" & _
										" , LOCATIONTYPE" & _
										" , LOCATIONCODE" & _
										" , chargetype" & _
							            " , chargecode" & _
							     		" , chargejob" & _
										" , chargeunit" & _
										" , CATEGORYTRANSPORT" & _
										" , TRIP" & _
										" , HOURS" & _
										" , UNIT" & _
										" , AMOUNT" & _
										" , OUTPUT" & _
										" , REF" & _
										" , REMARKS" & _
										" , TYPE_KONTANAN" & _
										" , FCENTRY" & _
										" , FCEDIT" & _
										" , FCIP" & _
										" , FCBA" & _
										" , LASTUPDATE" & _
										" , LASTTIME" & _
										" )" & _
										" values (" & _
										" '" & FCCODE & "'" & _
										" , LOWER(regexp_replace(rawtohex(sys_guid()), '([A-F0-9]{8})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{12})','\1-\2-\3-\4-\5'))" & _
										" , '" & EMPLOYEECODE & "'" & _
										" , '" & HELPER1 & "'" & _
										" , '" & HELPER2 & "'" & _
										" , to_date('" & fdDate & "','mm/dd/yyyy')" & _
										" , '" & TIME & "'" & _
										" , '" & JOB & "'" & _
										" , '" & LOCATIONTYPE & "'" & _
										" , '" & LOCATIONcode & "'" & _
										" , '" & chargetype & "'" & _ 
						         		" , '" & chargecode &  "'" & _
						         		" , '" & chargejob & "'" & _
										" , '" & chargeunit & "'" & _
										" , '" & CATEGORYTRANSPORT & "'" & _
										" , '" & TRIP & "'" & _
										" , '" & HOURS & "'" & _
										" , '" & UNIT & "'" & _
										" , '0'" & _
										" , '" & OUTPUT & "'" & _				
										" , '" & Trim(reference) & "'" & _
										" , '" & REMARKS & "'" & _
										" , '" & TYPE_KONTANAN & "'" & _
										" , '" & Session("CodeName") & "'" & _
										" , '" & Session("CodeName") & "'" & _
										" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										" , '" & Session("fcBA") & "'" & _
										" , sysdate" & _
										" , To_Char(sysdate,'HH24:MI')" & _
										" )"
										
									'response.Write(q)
									'response.Flush()
									connStr.execute(q)
							End If
						End If
					Else
						if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 or len(trim(""&chargeunit)) = 0 ) then
							err.number = 1001
							err.description = "Anda harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
						else
							Dim val_Charge : val_Charge = 0
							val_Charge = CheckCharge(chargetype, chargecode, chargejob, businessunit, fdDate, sourcetype)
							If val_Charge <> "Success" Then
								Err.Number = 1001
								Err.Description = val_Charge
							Else
								If len(trim(""&unit)) = 0 Then
									err.number = 1001
									err.description = "Anda Harus Mengisi Kolom Unit! "
								Else
									q	=	" Insert into VehicleACTIVITY (" & _
											" VEHICLECODE" & _
											" , PID" & _
											" , EMPLOYEECODE" & _
											" , HELPER1" & _
											" , HELPER2" & _
											" , VEHICLEDATE" & _
											" , TIME" & _
											" , JOBCODE" & _
											" , LOCATIONTYPE" & _
											" , LOCATIONCODE" & _
											" , chargetype" & _
								            " , chargecode" & _
								     		" , chargejob" & _
											" , chargeunit" & _
											" , CATEGORYTRANSPORT" & _
											" , TRIP" & _
											" , HOURS" & _
											" , UNIT" & _
											" , AMOUNT" & _
											" , OUTPUT" & _
											" , REF" & _
											" , REMARKS" & _
											" , TYPE_KONTANAN" & _
											" , FCENTRY" & _
											" , FCEDIT" & _
											" , FCIP" & _
											" , FCBA" & _
											" , LASTUPDATE" & _
											" , LASTTIME" & _
											" )" & _
											" values (" & _
											" '" & FCCODE & "'" & _
											" , LOWER(regexp_replace(rawtohex(sys_guid()), '([A-F0-9]{8})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{4})([A-F0-9]{12})','\1-\2-\3-\4-\5'))" & _
											" , '" & EMPLOYEECODE & "'" & _
											" , '" & HELPER1 & "'" & _
											" , '" & HELPER2 & "'" & _
											" , to_date('" & fdDate & "','mm/dd/yyyy')" & _
											" , '" & TIME & "'" & _
											" , '" & JOB & "'" & _
											" , '" & LOCATIONTYPE & "'" & _
											" , '" & LOCATIONcode & "'" & _
											" , '" & chargetype & "'" & _ 
							         		" , '" & chargecode &  "'" & _
							         		" , '" & chargejob & "'" & _
											" , '" & chargeunit & "'" & _
											" , '" & CATEGORYTRANSPORT & "'" & _
											" , '" & TRIP & "'" & _
											" , '" & HOURS & "'" & _
											" , '" & UNIT & "'" & _
											" , '0'" & _
											" , '" & OUTPUT & "'" & _				
											" , '" & Trim(reference) & "'" & _
											" , '" & REMARKS & "'" & _
											" , '" & TYPE_KONTANAN & "'" & _
											" , '" & Session("CodeName") & "'" & _
											" , '" & Session("CodeName") & "'" & _
											" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
											" , '" & Session("fcBA") & "'" & _
											" , sysdate" & _
											" , To_Char(sysdate,'HH24:MI')" & _
											" )"
											
									'response.Write(q)
									'response.Flush()
									connStr.execute(q)
							End If
						End If
					End If
				End If
			Else
				err.description = "Kode Karyawan tidak termasuk dalam Master"
				err.number = 10001
			End If
		Else
			err.description = "Masukkan Kode Karyawan di kolom Mechanic"
			err.number = 10001
		End If
	else
		err.number = 1001
		err.description = "Tanggal """ & fdDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
	End If
	cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fdDate)
		If Cek_IsDailyClosing = "FALSE" then
			' Call UpdateDataVehicle
			qLineMax = " select count(1) fccode, fcba from intraco_tab where status='TRUE' and fccode = '"& JOB &"' group by fcba"			
			' response.write qLineMax
			set rsqLineMax = Server.CreateObject("ADODB.Recordset")
			rsqLineMax.Open qLineMax, connstr
			if not rsqLineMax.eof then
				code = cdbl(to_nol(rsqLineMax("fccode")))
				'businessunit = rsqLineMax("fcba")
				businessunit = chargeunit
			end if 
			if code = 0 then
				
				q	=	" Update VehicleACTIVITY set" & _
				        " CATEGORYTRANSPORT = '" & CATEGORYTRANSPORT & "'" & _
	                    " , TIME = '" & TIME & "'" & _
	                    " , EMPLOYEECODE = '" & EMPLOYEECODE & "'" & _
	                    " , HELPER1 = '" & HELPER1 & "'" & _
	                    " , HELPER2 = '" & HELPER2 & "'" & _
	                    " , JOBCODE = '" & JOB & "' " & _
	                    " , LOCATIONTYPE = '" & LOCATIONTYPE & "'" & _
	                    " , LOCATIONCODE = '" & LOCATIONCODE & "'" & _
						" , CHARGEUNIT = '" & CHARGEUNIT & "'" & _
	                    " , CHARGEJOB = '" & CHARGEJOB & "'" & _
	                    " , CHARGETYPE = '" & CHARGETYPE & "'" & _
	                    " , CHARGECODE = '" & CHARGECODE & "'" & _
	                    " , TRIP = '" & TRIP & "'" & _					
						" , HOURS = '" & HOURS & "'" & _
						" , UNIT = '" & UNIT & "'" & _
						" , OUTPUT = '" & OUTPUT & "'" & _
						" , ref = '" & Trim(reference) & "'" & _
						" , remarks = '" & remarks & "'" & _
						" , type_kontanan = '" & type_kontanan & "'" & _
						" , LASTUPDATE = sysdate " & _
						" , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						" , fcedit = '" & Session("CodeName") & "'" & _
						" where" & _
						"  PID='" & PID & "'" & _
						" and fcBA='" & Session("fcBA") & "'" 
						
				' response.write q
				connStr.execute(q)
				
			Else
			if ( len(trim(""&chargetype)) = 0 or len(trim(""&chargecode)) = 0 or len(trim(""&chargejob)) = 0 or len(trim(""&chargeunit)) = 0 ) then
				err.number = 1001
				err.description = "Anda harus input Charge Unit, Charge Job, Charge Type dan Charge Code! "
			else
				Dim val_Charges : val_Charges = 0
				val_Charges = CheckCharge(chargetype, chargecode, chargejob, businessunit, fdDate, sourcetype)
				If val_Charges <> "Success" Then
					Err.Number = 1001
					Err.Description = val_Charges
				Else
				q	=	" Update VehicleACTIVITY set" & _
				        " CATEGORYTRANSPORT = '" & CATEGORYTRANSPORT & "'" & _
	                    " , EMPLOYEECODE = '" & EMPLOYEECODE & "'" & _
	                    " , HELPER1 = '" & HELPER1 & "'" & _
	                    " , HELPER2 = '" & HELPER2 & "'" & _
	                    " , JOBCODE = '" & JOB & "' " & _
	                    " , LOCATIONTYPE = '" & LOCATIONTYPE & "'" & _
	                    " , LOCATIONCODE = '" & LOCATIONCODE & "'" & _
						" , CHARGEUNIT = '" & CHARGEUNIT & "'" & _
	                    " , CHARGEJOB = '" & CHARGEJOB & "'" & _
	                    " , CHARGETYPE = '" & CHARGETYPE & "'" & _
	                    " , CHARGECODE = '" & CHARGECODE & "'" & _
	                    " , TRIP = '" & TRIP & "'" & _					
						" , HOURS = '" & HOURS & "'" & _
						" , UNIT = '" & UNIT & "'" & _
						" , OUTPUT = '" & OUTPUT & "'" & _
						" , ref = '" & Trim(reference) & "'" & _
						" , remarks = '" & remarks & "'" & _
						" , type_kontanan = '" & type_kontanan & "'" & _
						" , LASTUPDATE = sysdate " & _
						" , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
						" , fcedit = '" & Session("CodeName") & "'" & _
						" where" & _
						"  PID='" & PID & "'" & _
						" and fcBA='" & Session("fcBA") & "'" 
						
				'response.write q
				connStr.execute(q)
			End If
		End If
		End If
			' connStr.execute("commit")
		else
			err.number = 1001
			err.description = "Tanggal """ & fdDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if
else
    err.number = 1001
	err.description = "Accounting Period  " & ACCOUNTINGPERIOD & "/" & ACCOUNTINGYEAR & "  LOCKED !!!"
	cmdView	= request("cmdView")
End if 


Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if Request("fddate")<>"" then
   fddate   =Request("fddate")
else
   fddate   =textDate(date)
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
txtJenis = request("txtJenis")
if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & fdDate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & fdDate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Vehicle')"
			
set oRecordSet = server.CreateObject("Adodb.recordset")
oRecordSet.Open oSQLQuery, connstr
If Not oRecordSet.Eof then
	val_IsTrue = oRecordSet("istrue")
Else
	val_IsTrue = "N"
End if
'response.Write oSQLQuery
oRecordSet.Close
Set oRecordSet = Nothing

%>
 <!-------------ISI Payment Form------------->

 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-tachometer" aria-hidden="true"></i>Header Vehicle Activity</div><% end if %>
 <form name="frmEntry" method="post" action="VehicleActivity.asp"> 
	<table class="data">
		<tr><td align="left">Date</td>
		<td align="left"><input class="datepicker" id="vehdate" type="text" maxlength="10" name="fddate" size="10" onchange="" value="<%=fdDate%>">
		<font color="#FF0000">&nbsp;*</font>
		</td></tr>	
		<tr><td  align="left">Vehicle Code</td>
		<td td  align="left">
        <select id="combo_vehiclecode" name="FCCODE" >
	    <%cmdCallCombo "Vehicle",FCCODE,"",""%>
        </select>
		<font color="#FF0000">&nbsp;*</font>
		</td>
		</tr>
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit" id="selectHeader"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintVehicle.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fdDate=<%=fddate%>&VehicleCode=<%=fcCode%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		<%
		dim dataModule,dataSubmodule,dataDetail,hiddenHadir
		dataModule 		= "Vehicle"
		dataSubmodule 	= "Vehicle Activity"
		if fcCode = "" then
			dataDetail 	= ""
			hiddenHadir = "visibility: hidden;"
		else
			dataDetail 	= fcCode
			hiddenHadir = "visibility: visible;"
		
		end if
		
		%>
		<a class="jqplot-image-button" style="<%=hiddenHadir%>" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" and len(trim(FCCODE)) > 0 Then
	Dim rsVH,intRecordsPerPage,cColor
	set rsVH = server.CreateObject("Adodb.recordset")
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	=	" select" & _
			" VEHICLECODE" & _
			" , EMPLOYEECODE" & _
			" , HELPER1" & _
			" , HELPER2" & _
			" , to_char(VEHICLEDATE,'mm/dd/yyyy') VEHICLEDATE" & _
			" , TIME" & _
			" , JOBCODE" & _
			" , LOCATIONTYPE" & _
			" , LOCATIONCODE" & _
			" , CHARGEUNIT" & _
			" , CHARGEJOB" & _
			" , CHARGETYPE" & _
			" , CHARGECODE" & _
			" , CATEGORYTRANSPORT" & _
			" , TYPE_KONTANAN" & _
			" , TRIP" & _
			" , HOURS" & _
			" , UNIT" & _
			" , AMOUNT" & _
			" , OUTPUT" & _
			" , TRIM(REF) REF" & _
			" , REMARKS" & _
			" , FCENTRY" & _
			" , FCEDIT" & _
			" , FCIP" & _
			" , FCBA" & _
			" , PID" & _ 
			" from VehicleActivity where" & _
			" fcBA = '" & session("fcBA") & "'"
			
If Len(Trim(fdDate)) > 0 Then
	q	=	q & " and VEHICLEDATE = to_date('" & fdDate & "', 'MM/DD/YYYY')"
End If

If Len(Trim(fcCode)) > 0 Then
	q	=	q & " and VehicleCode ='" & fcCode & "'"
End If

	'response.Write q
	response.Flush()	
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="VehicleActivity.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0">
		<!--Yang tidak bisa di edit adalah PK Table-->
			<tr>
				<td align="left">
					<font face="Verdana" size="1">Operator/Driver</font>
				</td>
				<td align="left">
					<input id="empcode" type="text" size="15" name="EMPLOYEECODE" value="<%=EMPLOYEECODE%>" onchange="GetEmpName(this.value);" >
					<span id="nonscript_empcd" style="color:black;">&nbsp;*&nbsp;
					<a id ="nonscriptempcd_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
					<font face="Verdana"   size="1">Employee</font></a></span>
				</td>	
			</tr>
			<tr>
				<td align="left">
					<font face="Verdana" size="1">Helper 1</font>
				</td>
				<td align="left">
					<input id="helper" type="text" size="15" name="HELPER1" value="<%=HELPER1%>" onchange="GetHelper(this.value);" >
					<span id="nonscript_empcd_helper" style="color:black;">&nbsp;*&nbsp;
					<a id ="nonscriptempcd_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
					<font face="Verdana"   size="1">Employee</font></a></span>
				</td>	
			</tr>
			<tr>
				<td align="left">
					<font face="Verdana" size="1">Helper 2</font>
				</td>
				<td align="left">
					<input id="helpers" type="text" size="15" name="HELPER2" value="<%=HELPER2%>" onchange="GetHelpers(this.value);" >
					<span id="nonscript_empcd_helpers" style="color:black;">&nbsp;*&nbsp;
					<a id ="nonscriptempcd_a" href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
					<font face="Verdana"   size="1">Employee</font></a></span>
				</td>	
			</tr>
 		     <tr><td align="Left">Time (HH:MM-HH:MM)</td><td align="Left"><input type="text" size="15" name="TIME" value="<%=TIME%>" placeholder="hh:mm-hh:mm" class="timepicker" autocomplete="off"></td></tr>
		     <tr>
		     	<td align="Left">Job Code</td>
		     	<td align="Left"><input onkeyup="showHint(this.value)" type="text" size="15" id="jobcode" name="Job" style="background:#dddddd" readonly   value="<%=JOB%>">&nbsp;<font color="#FF0000">*</font>&nbsp;
				<a href="javascript:launchwin('popup_activity.asp','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font><span class="satuan_item" id="jobname"></span></a>
				</td>
		     </tr>
		     <tr>
		     	<td align="Left">Target Type</td>
		     	<td align="Left"><input type="text" size="15" id="locationtypecode" name="locationtype" value="<%=locationtype%>" style="background:#dddddd" >&nbsp;<font color="#FF0000">*</font>&nbsp;
				<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Location Type</font></a>
		     	</td>
		     </tr>
		     <tr><td align="Left">Target Code</td><td align="Left"><input type="text" size="15" id="locationcode" name="LOCATIONcode" value="<%=LOCATIONcode%>" style="background:#dddddd" >&nbsp;<font color="#FF0000">*</font>&nbsp;
			 <a href="javascript:PopupTType('<%=fcba %>')"><font face="Verdana" size="1">Location Code</font> <span class="satuan_item" id="locationname"></span></a>
			 </td></tr>
		     <!-- Add Charge Job, Charge Type, Charge Code -->
		     <tr>
                <td align="left">
                    Charge Unit
                </td>
                <td align="left">
                   <select  name="chargeunit" id = "chargeunit_" >
							<option value=""></option>
							<%
							qlov = 	 " select fccode,fcname " & _
									 "	from businessunit " & _
									 "	where fccompanycode = businessunit_api.get_CompanyCode('" & session("fcba") & "') " & _
									 "	and businessunit_api.Isactive(fccode)='TRUE' " & _ 
									 "  and fccode <> '" & session("fcba") & "'"
									 
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
                </td>
            </tr>
			 <tr>
                <td align="left">
                    Charge Job
                </td>
                <td align="left">
                    <input type="text" size="15" id="chargejob_" name="chargejob" value="<%=chargejob%>" style="background:#dddddd" readonly>&nbsp;<font color="#FF0000">*</font>
                    <a href="javascript:launchwin('popup_activity_intraco.asp','newwindow','width='+(1000)+',height='+(705)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana"   size="1">Activity</font><span class="satuan_item" id="jobname"></span></a>
				</td>
            </tr>
            <tr>
                <td align="left">
                    Charge Type
                </td>
                <td align="left">
                    <input type="text" size="15" name="chargetype" id="chargetype_" value="<%=chargetype%>" style="background:#dddddd" readonly>&nbsp;<font color="#FF0000">*</font>
						<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=LocationType&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><font face="Verdana" size="1">Location Type</font></a>
                </td>
            </tr>
            <tr>
                <td align="left">
                    Charge Code
                </td>
                <td align="left">
                    <input type="text" size="15" name="chargecode" id="chargecode_" value="<%=chargecode%>" style="background:#dddddd" readonly>&nbsp;<font color="#FF0000">*</font>
					<a href="javascript:PopupTypeCC('')"><font face="Verdana" size="1">Location Code</font><span class="satuan_item" id="locationnamecc">&nbsp;</span></a></font>
                </td>
            </tr>
		<tr>
                <td align="left">
                    Category Transport
                </td>
				<% 	Dim rsCombo,Data_JobsAccrual
				    q = " SELECT fccode, fcname FROM vehicleactivitycategory WHERE fcba='" & session("fcba") & "' and Isactive='1' ORDER BY LPAD(fccode, 2, 0)"
					set rsCombo = Server.CreateObject("ADODB.Recordset")
					'response.write q
					rsCombo.CursorType = 3
					rsCombo.CursorLocation = 3
					rsCombo.Open q, connstr
					if not rsCombo.eof then
				%>
				<td align="left"> 
							<select name="CategoryTransport">    
								<% Do While Not rsCombo.eof %>
									
									  <option value="<%=rsCombo("fccode")%>" <% If Request("CategoryTransport") = rsCombo("fccode") Then %> Selected <%End if%>><%=rsCombo("fccode")%> : <%=rsCombo("fcname")%></option> 
								<%    rsCombo.movenext
								   loop

								
					End if
					rsCombo.Close
					set rsCombo = Nothing
					q=""
				%>
							</select> <font color="#FF0000">*</font>
				</td> 
        </tr>
        <tr>
        	<td align="left">
        		Type Kontanan
        	</td>
        	<td align="left">
        		<input type="radio" name="type_kontanan" value="Non Kontanan" checked> Non Kontanan<br>
        		<input type="radio" name="type_kontanan" value="Kontanan"> Kontanan<br> 
        	</td>
        </tr>
		<tr><td align="Left">Jml Trip</td><td align="Left"><input type="text" size="5" name="TRIP" value="<%=TRIP%>">&nbsp; Kali</td></tr>
		<tr><td align="Left">Hours (Jam Dunia)</td><td align="Left"><input type="text" size="15" name="HOURS" value="<%=HOURS%>"><span class="satuan_item" id="hours">&nbsp;...</span></td></tr>
		<tr><td align="Left">Km/Hm</td><td align="Left"><input type="text" size="15" name="UNIT" value="<%=UNIT%>"><span class="satuan_item" id="unit">&nbsp;...</span></td></tr>
		<tr><td align="Left">Output</td><td align="Left"><input type="text" size="15" name="OUTPUT" value="<%=OUTPUT%>"><span class="satuan_item" id="output">&nbsp;...</span></td></tr>
		<tr><td align="Left">Reference</td><td align="Left"><input type="text" size="50" name="Reference" value="<%=Reference%>"></td></tr>
		<tr><td align="Left">Remarks</td><td align="Left"><input type="text" size="50" name="REMARKS" value="<%=REMARKS%>"></td></tr>
	    <TR><td colspan="2"  align="Left">
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="fddate" value="<%=fddate%>">
		  <input type="hidden" name="PID" value="<%=PID%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <!-- Validasi Period Control In Here -->
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" id="btnUpdate" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit" id="btnNew" name="Add" value="Add">
		  <%End if%>
		  <!-- End Validasi Period Control In Here -->
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Item</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines">
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">CODE</TD>
	<TD align="left">EMPLOYEE CODE</TD>
	<TD align="left">HELPER 1</TD>	
	<TD align="left">HELPER 2</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">TIME</TD>	
	<TD align="left">JOB CODE</TD>	
	<TD align="left">LOC TYPE</TD>	
	<TD align="left">LOC CODE</TD>
	<TD align="left">CHARGE UNIT</TD>	
	<TD align="left">CHARGE JOB</TD>	
	<TD align="left">CHARGE TYPE</TD>	
	<TD align="left">CHARGE CODE</TD>
	<TD align="left">CAT TRANS</TD>
	<TD align="left">TYPE KONTANAN</TD>
	<TD align="left">JML TRIP</TD>
	<TD align="left">HOURS</TD>	
	<TD align="left">UNIT</TD>
	<TD align="left">AMOUNT</TD>
	<TD align="left">OUTPUT</TD>		
	<TD align="left">REFERENCES</TD>	
	<TD align="left">REMARKS</TD>	
	</tr>
	</thead>
	<tbody>
   
   <%
   if Not rsVH.eof then
      nCountofField=rsVH.fields.count-1
	  no=1
	  While Not rsVH.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
		  <!-- Validasi Period Control In Here -->
			<td valign="middle" width="30">
				<a href="VehicleActivity.Asp?PID=<%=rsVH("PID")%>&EMPLOYEECODE=<%=rsVH("EMPLOYEECODE")%>&helper1=<%=rsVH("helper1")%>&helper2=<%=rsVH("helper2")%>&fcCode=<%=rsVH("VEHICLECODE")%>&fdDate=<%=rsVH("VEHICLEDATE")%>&TIME=<%=rsVH("TIME")%>&JOB=<%=rsVH("JOBCODE")%>&LOCATIONTYPE=<%=rsVH("LOCATIONTYPE")%>&CATEGORYTRANSPORT=<%=rsVH("CATEGORYTRANSPORT")%>&TRIP=<%=rsVH("TRIP")%>&LOCATIONcode=<%=rsVH("LOCATIONCODE")%>&HOURS=<%=rsVH("HOURS")%>&UNIT=<%=rsVH("UNIT")%>&OUTPUT=<%=rsVH("OUTPUT")%>&reference=<%=rsVH("ref")%>&remarks=<%=rsVH("remarks")%>&chargeunit=<%=rsVH("chargeunit")%>&chargejob=<%=rsVH("chargejob")%>&chargetype=<%=rsVH("chargetype")%>&chargecode=<%=rsVH("chargecode")%>&type_kontanan=<%=rsVH("type_kontanan")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
			</td>
			<td>
				<a href="VehicleActivity.Asp?PID=<%=rsVH("PID")%>&fcCode=<%=rsVH("VEHICLECODE")%>&fdDate=<%=rsVH("VEHICLEDATE")%>&TIME=<%=rsVH("TIME")%>&JOB=<%=rsVH("JOBCODE")%>&LOCATIONTYPE=<%=rsVH("LOCATIONTYPE")%>&LOCATIONcode=<%=rsVH("LOCATIONCODE")%>&chargejob=<%=rsVH("chargejob")%>&chargetype=<%=rsVH("chargetype")%>&chargecode=<%=rsVH("chargecode")%>&type_kontanan=<%=rsVH("type_kontanan")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0" onClick="return confirm('Are you sure to delete');"></a>
			</td>
		  <!-- End Validasi Period Control In Here -->
		    <%for i = 0 to nCountofField-5%>
			    <td align="left" title="<%=rsVH(i).Name%>" valign="top"><%=rsVH(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsVH.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
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
 
  <!-------------akhir------------->
</div></div>
  <!--#include file="bottom.jsp"-->
<script type="text/javascript">

</script>
<script>
	$(function() {
	  $("input:text").keyup(function() {
			$(this).val($(this).val().replace(/|,|'|"|&|\r\n|\r|\n/g, ""));
	  });
	});

	$.mask.definitions['H'] = "[0-2]";
	$.mask.definitions['h'] = "[0-9]";
	$.mask.definitions['M'] = "[0-5]";
	$.mask.definitions['m'] = "[0-9]";
	$.mask.definitions['P'] = "[AaPp]";
	$.mask.definitions['p'] = "[Mm]";

	$(".timepicker").mask("Hh:Mm-Hh:Mm");


function get_Allowentry()
{
	var vehdate = $('#vehdate').val();
	if (vehdate != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + vehdate, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE"){
				$("input[name='Add']").show();
				
				
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$("input[name='Add']").hide();
				
			}
			
		});
		
	} 
}
	
</script>
</body>
</html>
