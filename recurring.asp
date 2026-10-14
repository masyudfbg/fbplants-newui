<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<%
	' Note
	' 1. Tambahkan Record Cancelled table rowstate
	' 2. Buat Control Status dengan Luname RECURRING table controlstatus
	' 3. Tambahkan Package RECURRING_API
%>
<html>
<head>
<title>ePlaS - Recurring</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script language="javascript" type="text/javascript" src="ajax/validate_ajax.js"></script>

<script language="javascript" type="text/javascript">

                function showHint(str)
					{
				var xmlhttp;
				var unit;
				var output;
				var oth;
				var jobname;
					if (str=="")
					  {
					  document.getElementById("mandays").innerHTML="";
					  document.getElementById("unit").innerHTML="";
					  document.getElementById("output").innerHTML="";
					  document.getElementById("oth").innerHTML="";
					  document.getElementById("jobname").innerHTML="";
				
					  return;
					  }
					if (window.XMLHttpRequest)
					  {// code for IE7+, Firefox, Chrome, Opera, Safari
					  xmlhttp=new XMLHttpRequest();
					  unit=new XMLHttpRequest();
					  output=new XMLHttpRequest();
					  oth=new XMLHttpRequest();
					  jobname=new XMLHttpRequest();
					  }
					else
					  {// code for IE6, IE5
					  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
					  unit=new ActiveXObject("Microsoft.XMLHTTP");
					  output=new ActiveXObject("Microsoft.XMLHTTP");
					  oth=new ActiveXObject("Microsoft.XMLHTTP");
					  jobname=new ActiveXObject("Microsoft.XMLHTTP");
					  }
					  
					  
						xmlhttp.onreadystatechange=function()
						  {
						  if (xmlhttp.readyState==4 && xmlhttp.status==200)
							{
							document.getElementById("mandays").innerHTML=xmlhttp.responseText;
							}
						  }
						xmlhttp.open("GET","ajax/MANDAYS.asp?ajaxUoM=mandays&q="+str,true);
						xmlhttp.send();
						
						
						unit.onreadystatechange=function()
						  {
						  if (unit.readyState==4 && unit.status==200)
							{
							document.getElementById("unit").innerHTML=unit.responseText;
							}
						  }
						unit.open("GET","ajax/MANDAYS.asp?ajaxUoM=unit&q="+str,true);
						unit.send();
						
						output.onreadystatechange=function()
						  {
						  if (output.readyState==4 && output.status==200)
							{
							document.getElementById("output").innerHTML=output.responseText;
							}
						  }
						output.open("GET","ajax/MANDAYS.asp?ajaxUoM=output&q="+str,true);
						output.send();
						
						oth.onreadystatechange=function()
						  {
						  if (oth.readyState==4 && oth.status==200)
							{
							document.getElementById("oth").innerHTML=oth.responseText;
							}
						  }
						oth.open("GET","ajax/MANDAYS.asp?ajaxUoM=oth&q="+str,true);
						oth.send();
						
						jobname.onreadystatechange=function()
						  {
						  if (jobname.readyState==4 && jobname.status==200)
							{
							document.getElementById("jobname").innerHTML=jobname.responseText;
							}
						  }
						jobname.open("GET","ajax/MANDAYS.asp?ajaxUoM=jobname&q="+str,true);
						jobname.send();
					
					}

				

</script>
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
function periksa(form)
 { if(form.mandays.value > 1)
    {alert("mandays diisi paling besar 1 !");
     form.mandays.focus();
     form.mandays.select();
     return(false);
    }
 
  return(true);
 }
 

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
 
Function GetYearofDate(date__)
	Dim result__
	Dim rsyearofdate, stryearofdate
	
	stryearofdate =	" SELECT TO_CHAR(TO_DATE('" & date__ & "', 'mm/dd/yyyy'), 'yyyy') year" & _
	                " FROM DUAL"
	'Response.Write stryearofdate
	Set rsyearofdate = Server.CreateObject("ADODB.Recordset")
	rsyearofdate.Open stryearofdate, connstr, 3
	
	If Not rsyearofdate.EOF Then
		result__ = Cdbl(rsyearofdate("year"))
	Else
		result__ = 2000
	End If
	
	stryearofdate =	""
	rsyearofdate.Close
	set rsyearofdate = Nothing

	GetYearofDate = result__
End Function

Function CheckTriggerControl(initialJobCode, initialSource, initialTarget)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable =	" SELECT COUNT(1) Total" & _
	                " FROM TRIGGER_CONTROL" & _
					" WHERE activity = '" & initialJobCode & "'" & _
					"       AND source = '" & initialSource & "'" & _
					"       AND target = '" & initialTarget & "'"
							
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

	CheckTriggerControl = initialAvailable
End Function

Function Checkdatadetail(docno__, jobcode__, targettype__)
	Dim result__
	Dim rsdatadetail, strdatadetail
	
	strdatadetail =	" SELECT COUNT(1) total" & _
	                " FROM recurringdtl" & _
					" WHERE fcba = '" & Session("fcBA") & "'" & _
					"       AND documentno = '" & docno__ & "'" & _
					"       AND jobcode = '" & jobcode__ & "'" & _
					"       AND targettype = '" & targettype__ & "'"
	'Response.Write strdatadetail
	Set rsdatadetail = Server.CreateObject("ADODB.Recordset")
	rsdatadetail.Open strdatadetail, connstr, 3
	
	If Not rsdatadetail.EOF Then
		result__ = Cdbl(rsdatadetail("total"))
	Else
		result__ = 0
	End If
	
	strdatadetail =	""
	rsdatadetail.Close
	set rsdatadetail = Nothing

	Checkdatadetail = result__
End Function

Dim rsrecurring, qrecurring
Dim totalamountupdate
Dim comm, param, fcba_, DocumentNo_, date_from_, Date_until_, fcentry_, fcedit_, fcip_, ResultCommand_

documentno = Request("documentno")
documentname = Request("documentname")
startingdate = Request("startingdate")
finisheddate = Request("finisheddate")
defaultexecuted = Request("defaultexecuted")
intervalinmonth = Request("intervalinmonth")
numberofruns = Request("numberofruns")
amount = Request("amount")
isusesettlement = Request("isusesettlement")
accrualjob = Request("accrualjob")
state = Request("state")
jobcode = Request("jobcode")
targettype = Request("targettype")
debetcredit = Request("debetcredit")
detailamount = Request("detailamount")

if "" & startingdate = "" then startingdate = textDate(date)
if "" & finisheddate = "" then finisheddate = textDate(date)
if "" & intervalinmonth = "" then intervalinmonth = 1
if "" & numberofruns = "" then numberofruns = 1
if "" & amount = "" then amount = 0
if "" & isusesettlement = "" then isusesettlement = "FALSE"
if "" & debetcredit = "" then debetcredit = "DEBET"
if "" & detailamount = "" then detailamount = 0
if "" & state = "" then state = "Planned"


IF Request("txtJenis") = "New" OR Request("txtJenis") = "Header" OR Request("txtJenis") = "Delete" OR Request("txtJenis") = "Edit" OR Request("txtJenis") = "Confirmed" OR Request("txtJenis") = "Cancelled" OR Request("txtJenis") = "Executed" OR Request("txtJenis") = "RePlanned" THEN
	IF Request("txtJenis") = "Delete" THEN
		q =	" DELETE" & _
            " FROM recurringdtl" & _
            " WHERE documentno = '" & documentno & "'" & _
            "       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
            "       AND fcba = '" & Session("fcBA") & "'" & _
			"       AND jobcode = '" & jobcode & "'"
		connStr.execute(q)
		
		qrecurring = " SELECT NVL(SUM(amount), 0) total" & _
					 " FROM recurringdtl" & _
					 " WHERE documentno = '" & documentno & "'" & _
					 "       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
					 "       AND fcba = '" & Session("fcBA") & "'"
		'Response.write qrecurring
		set rsrecurring = server.CreateObject("Adodb.recordset")
		rsrecurring.Open qrecurring, connstr
		
		IF NOT rsrecurring.EOF THEN
			totalamountupdate = Cdbl(rsrecurring("total"))
		ELSE
			totalamountupdate = 0
		END IF
		
		q =	" UPDATE recurring" & _
			" SET amount = '" & totalamountupdate & "'" & _
			" WHERE documentno = '" & documentno & "'" & _
			"       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
			"       AND fcba = '" & Session("fcBA") & "'"
		connStr.execute(q)		
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Header" THEN
		q =	" DELETE" & _
            " FROM recurringdtl" & _
            " WHERE documentno = '" & documentno & "'" & _
            "       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
            "       AND fcba = '" & Session("fcBA") & "'"
		connStr.execute(q)
		
		q =	" DELETE" & _
            " FROM recurring" & _
            " WHERE documentno = '" & documentno & "'" & _
            "       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
            "       AND fcba = '" & Session("fcBA") & "'"
		connStr.execute(q)
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "New" THEN
		IF GetYearofDate(startingdate) <> GetYearofDate(finisheddate) THEN
			err.number = 1001
			err.Description = " Starting Date tidak sama dengan Finished Date! Kolom tersebut harus dalam tahun yang sama."
		ELSE
			'--Insert data Header
			qrecurring = " SELECT documentno" & _
						 " FROM recurring" & _
						 " WHERE fcba = '" & Session("fcBA") & "'" & _
						 "       AND documentno = '" & documentno & "'"
			'Response.write qrecurring
			set rsrecurring = server.CreateObject("Adodb.recordset")
			rsrecurring.Open qrecurring, connstr
			IF rsrecurring.EOF THEN
				q = " INSERT INTO recurring ( documentno," & _
					"                         documentname," & _
					"                         accountingyear," & _
					"                         startingdate," & _
					"                         finishdate," & _
					"                         intervalinmonth," & _
					"                         rundate," & _
					"                         numberofrun," & _
					"                         currency," & _
					"                         amount," & _
					"                         fcba," & _
					"                         isusesettlement," & _
					"                         accrualjob," & _
					"                         status," & _
					"                         fcentry," & _
					"                         fcedit," & _
					"                         fcip," & _
					"                         lastupdate," & _
					"                         lasttime)" & _
					" VALUES ( '" & documentno & "'," & _
					"          '" & documentname & "'," & _
					"          '" & GetYearofDate(finisheddate) & "'," & _
					"          TO_DATE('" & startingdate & "', 'mm/dd/yyyy')," & _
					"          TO_DATE('" & finisheddate & "', 'mm/dd/yyyy')," & _
					"          '" & intervalinmonth & "'," & _
					"          TO_DATE('" & finisheddate & "', 'mm/dd/yyyy')," & _
					"          '" & numberofruns & "'," & _
					"          'IDR'," & _
					"          '" & amount & "'," & _
					"          '" & Session("fcBA") & "'," & _
					"          '" & isusesettlement & "'," & _
					"          '" & accrualjob & "'," & _
					"          'Planned'," & _
					"          '" & Session("CodeName") & "'," & _
					"          '" & Session("CodeName") & "'," & _
					"          '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
					"          SYSDATE," & _
					"          TO_CHAR(SYSDATE, 'hh24:mi'))"
				connStr.execute(q)
			END IF
			
			'--Insert data Detail
			IF CheckTriggerControl(jobcode, "BA", targettype) = 0 THEN
				err.number = 1001
				err.Description = " Jobcode: " & jobcode & " tidak bisa digunakan ke target: " & targettype & ", cek kembali jobcode dan target type nya!"
			ELSE
				IF Checkdatadetail(documentno, jobcode, targettype) <> 0 THEN
					err.number = 1001
					err.Description = " Jobcode: " & jobcode & " dan Target Type: " & targettype & " sudah ada didalam sistem dengan No Dokumen yang sama: " & documentno & ""
				ELSE
					
					q = " INSERT INTO recurringdtl ( documentno," & _
						"                            accountingyear," & _
						"                            debetcredit," & _
						"                            jobcode," & _
						"                            targettype," & _
						"                            amount," & _
						"                            fcba," & _
						"                            status," & _
						"                            fcentry," & _
						"                            fcedit," & _
						"                            fcip," & _
						"                            lastupdate," & _
						"                            lasttime)" & _
						" VALUES ( '" & documentno & "'," & _
						"          '" & GetYearofDate(finisheddate) & "'," & _
						"          '" & debetcredit & "'," & _
						"          '" & jobcode & "'," & _
						"          '" & targettype & "'," & _
						"          REPLACE('" & detailamount & "', ',', '')," & _
						"          '" & Session("fcBA") & "'," & _
						"          'Planned'," & _
						"          '" & Session("CodeName") & "'," & _
						"          '" & Session("CodeName") & "'," & _
						"          '" & Request.ServerVariables("REMOTE_ADDR") & "'," & _
						"          SYSDATE," & _
						"          TO_CHAR(SYSDATE, 'hh24:mi'))"
					connStr.execute(q)
					
					
					qrecurring = " SELECT NVL(SUM(amount), 0) total" & _
								 " FROM recurringdtl" & _
								 " WHERE documentno = '" & documentno & "'" & _
								 "       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
								 "       AND fcba = '" & Session("fcBA") & "'"
					'Response.write qrecurring
					set rsrecurring = server.CreateObject("Adodb.recordset")
					rsrecurring.Open qrecurring, connstr
					
					IF NOT rsrecurring.EOF THEN
						totalamountupdate = Cdbl(rsrecurring("total"))
					ELSE
						totalamountupdate = 0
					END IF
					
					q =	" UPDATE recurring" & _
						" SET amount = '" & totalamountupdate & "'" & _
						" WHERE documentno = '" & documentno & "'" & _
						"       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
						"       AND fcba = '" & Session("fcBA") & "'"
					connStr.execute(q)
				END IF
			END IF
		END IF
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Edit" AND Request("Add") = "Update" THEN
		IF CheckTriggerControl(jobcode, "BA", targettype) = 0 THEN
			err.number = 1001
			err.Description = " Jobcode: " & jobcode & " tidak bisa digunakan ke target: " & targettype & ", cek kembali jobcode dan target type nya!"
		ELSE
			q =	" UPDATE recurringdtl" & _
				" SET amount = REPLACE('" & detailamount & "', ',', '')" & _
				" WHERE documentno = '" & documentno & "'" & _
				"       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
				"       AND fcba = '" & Session("fcBA") & "'" & _
				"       AND jobcode = '" & jobcode & "'" & _
				"       AND targettype = '" & targettype & "'"
			'Response.Write "sql: " & q
			connStr.execute(q)
			
			
			qrecurring = " SELECT NVL(SUM(amount), 0) total" & _
						 " FROM recurringdtl" & _
						 " WHERE documentno = '" & documentno & "'" & _
						 "       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
						 "       AND fcba = '" & Session("fcBA") & "'"
			'Response.write qrecurring
			set rsrecurring = server.CreateObject("Adodb.recordset")
			rsrecurring.Open qrecurring, connstr
			
			IF NOT rsrecurring.EOF THEN
				totalamountupdate = Cdbl(rsrecurring("total"))
			ELSE
				totalamountupdate = 0
			END IF
			
			q =	" UPDATE recurring" & _
				" SET amount = '" & totalamountupdate & "'" & _
				" WHERE documentno = '" & documentno & "'" & _
				"       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
				"       AND fcba = '" & Session("fcBA") & "'"
			connStr.execute(q)
		END IF
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "RePlanned" THEN
		q = " BEGIN RECURRING_API.NextState('" & Session("fcBA") & "', '" & documentno & "', '" & GetYearofDate(finisheddate) & "', 'RePlan' ); END; "
		connStr.execute(q)
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Confirmed" THEN
		q = " BEGIN RECURRING_API.NextState('" & Session("fcBA") & "', '" & documentno & "', '" & GetYearofDate(finisheddate) & "', 'Confirm' ); END; "
		connStr.execute(q)
		cmdView	= "True"
		' Response.Redirect "/recurring.asp?documentno=" & documentno & "&finisheddate=" & finisheddate & ""
	ELSEIF Request("txtJenis") = "Cancelled" THEN
		q = " BEGIN RECURRING_API.NextState('" & Session("fcBA") & "', '" & documentno & "', '" & GetYearofDate(finisheddate) & "', 'Cancel' ); END; "
		connStr.execute(q)
		cmdView	= "True"
	ELSEIF Request("txtJenis") = "Executed" THEN
		Dim prmfcba, prmdocumentno, prmdatefrom, prmdateuntil, prmfcentry, prmfcedit, prmfcip, prmresult
		fcba_ = Session("fcBA")
		DocumentNo_ = documentno
		date_from_ = startingdate
		Date_until_ = finisheddate
		fcentry_ = Session("CodeName")
		fcedit_ = Session("CodeName")
		fcip_ = Request.ServerVariables("REMOTE_ADDR")
		
		Set comm = Server.CreateObject("ADODB.Command")
		With comm 
			.CommandText = "{ CALL RECURRING_API.CREATE_RCR_TRANS (?,?,?,?,?,?,?,?) }"
			.CommandType = adCmdText
			.ActiveConnection = connStr
			.Prepared = True
			Set prmfcba = .CreateParameter("",adVarchar,adParamInput,255)
			Set prmdocumentno = .CreateParameter("",adVarchar,adParamInput,255)
			Set prmdatefrom = .CreateParameter("",adDate,adParamInput)
			Set prmdateuntil = .CreateParameter("",adDate,adParamInput)
			Set prmfcentry = .CreateParameter("",adVarchar,adParamInput,255)
			Set prmfcedit = .CreateParameter("",adVarchar,adParamInput,255)
			Set prmfcip = .CreateParameter("",adVarchar,adParamInput,255)
			Set prmresult = .CreateParameter("",adLongVarChar,adParamOutput,1000)
			.Parameters.Append(prmfcba)
			.Parameters.Append(prmdocumentno)
			.Parameters.Append(prmdatefrom)
			.Parameters.Append(prmdateuntil)
			.Parameters.Append(prmfcentry)
			.Parameters.Append(prmfcedit)
			.Parameters.Append(prmfcip)
			.Parameters.Append(prmresult)
		End With

		With comm 
			.Parameters(0).Value = fcba_
			.Parameters(1).Value = DocumentNo_
			.Parameters(2).Value = date_from_
			.Parameters(3).Value = Date_until_
			.Parameters(4).Value = fcentry_
			.Parameters(5).Value = fcedit_
			.Parameters(6).Value = fcip_
			'.Parameters(7).Value = ResultCommand_
			ResultCommand_ = .Parameters(7).Value
			Set rsrecurring  = .Execute
			
			'Do Until rsrecurring.BOF Or rsrecurring.EOF
			  Response.Write(ResultCommand_)
			  'rsrecurring.MoveNext
			'Loop
		End With
		
		' q = "{ CALL RECURRING_API.CREATE_RCR_TRANS (?,?,?,?,?,?,?,?) }"
		' set comm = server.CreateObject ("ADODB.Command")
		
		' with comm
			' set .ActiveConnection = connStr
			' .CommandText = q
			' .CommandType = adCmdText
			' set fcba_ = .CreateParameter("fcba", adVarChar , adParamInput , , 0)
			' set DocumentNo_ = .CreateParameter("documentno", adVarChar , adParamInput , , 1)
			' set date_from_ = .CreateParameter("startingdate", adDate , adParamInput , , 2)
			' set Date_until_ = .CreateParameter("finisheddate", adDate , adParamInput , , 3)
			' set fcentry_ = .CreateParameter("fcentry", adVarChar , adParamInput , , 4)
			' set fcedit_ = .CreateParameter("fcedit", adVarChar , adParamInput , , 5)
			' set fcip_ = .CreateParameter("fcip", adVarChar , adParamInput , , 6)
			' set ResultCommand_ = .CreateParameter("resultcommand", adLongVarChar , adParamOutput , , 7)
			' .Parameters.Append fcba_
			' .Parameters.Append DocumentNo_
			' .Parameters.Append date_from_
			' .Parameters.Append Date_until_
			' .Parameters.Append fcentry_
			' .Parameters.Append fcedit_
			' .Parameters.Append fcip_
			' .Parameters.Append ResultCommand_
		' end with
		
		' set comm = server.CreateObject ("ADODB.Command")
		' set comm.ActiveConnection = connStr
		' comm.CommandText = "RECURRING_API.CREATE_RCR_TRANS"
		' comm.CommandType = adCmdStoredProc
		' comm.Parameters(0) = fcba_
		' comm.Parameters(1) = DocumentNo_
		' comm.Parameters(2) = date_from_
		' comm.Parameters(3) = Date_until_
		' comm.Parameters(4) = fcentry_
		' comm.Parameters(5) = fcedit_
		' comm.Parameters(6) = fcip_
		' comm.Execute
		' ResultCommand_ = comm.Parameters(7)
			
		' q = " VAR out_ VARCHAR2(1000); BEGIN RECURRING_API.CREATE_RCR_TRANS( '" & Session("fcBA") & "', '" & documentno & "', TO_DATE('" & startingdate & "', 'mm/dd/yyyy'), TO_DATE('" & finisheddate & "', 'mm/dd/yyyy'), '" & Session("CodeName") & "', '" & Session("CodeName") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', out_); END; "
		' Response.Write "sql: " & q
		' connStr.execute(q)
		
		
		q = " BEGIN RECURRING_API.NextState('" & Session("fcBA") & "', '" & documentno & "', '" & GetYearofDate(finisheddate) & "', 'Execute' ); END; "
		'connStr.execute(q)
		
		response.write ResultCommand_
		
		
		
		cmdView	= "True"
	ELSE
		cmdView	= "True"
	END IF
	'Response.Write "sql: " & q
ELSE
    cmdView	= request("cmdView")
END IF

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis, cmdView

Dim documentno, documentname, startingdate, finisheddate,defaultexecuted, intervalinmonth, numberofruns, amount, isusesettlement, state, accrualjob
Dim jobcode, targettype, debetcredit, detailamount

txtTable = Request("txtTable")
txtModule = Request("txtModule")
txtSearch = Request("txtSearch")
txtField = Request("txtField")
txtJenis = request("txtJenis")


qrecurring = " SELECT *" & _
             " FROM recurring" & _
             " WHERE documentno = '" & documentno & "'" & _
             "       AND accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
             "       AND fcba = '" & Session("fcBA") & "'"

set rsrecurring = server.CreateObject("Adodb.recordset")
rsrecurring.Open qrecurring, connstr
IF NOT rsrecurring.EOF THEN
	documentname = rsrecurring("documentname")
	startingdate = rsrecurring("startingdate")
	finisheddate = rsrecurring("finishdate")
	intervalinmonth = rsrecurring("intervalinmonth")
	numberofruns = rsrecurring("numberofrun")
	amount = rsrecurring("amount")
	isusesettlement = rsrecurring("isusesettlement")
	accrualjob = rsrecurring("accrualjob")
	state = rsrecurring("status")
END IF

if err.number <> 0 then
   Response.Write "<div class = ""error"">" & Err.Description & "</div>"
End if

if "" & ResultCommand_ <> "" then
	Response.Write "<div class = ""error"">" & ResultCommand_ & "</div>"
end if

 '-------------ISI Payment Form
%>
<form name="frmEntry" method="post" action="recurring.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Document No</td>
			<td align="left">
				<input type="text" maxlength = "35" name="documentno" size="30" value="<%=documentno%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Document Name</td>
			<td align="left">
				<input type="text" maxlength = "100" name="documentname" size="50" value="<%=documentname%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Starting Date</td>
			<td align="left">
				<input id="startingdate" class="datepicker" type="text" maxlength="10" name="startingdate" size="10" value="<%=startingdate%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Finished Date</td>
			<td align="left">
				<input class="datepicker" type="text" maxlength="10" name="finisheddate" size="10" value="<%=finisheddate%>" onchange="getnumberofruns(this.value);">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Default Executed Date per Month Type</td>
			<td align="left">
				<select name="defaultexecuted">
				<option value="month" <%if defaultexecuted = "month" then %> selected <%end if%>> Monthly</option>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Interval In Month</td>
			<td align="left">
				<input type="number" name="intervalinmonth" size="30" value="<%=intervalinmonth%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Number of Runs</td>
			<td align="left">
				<input id="numberofruns" type="number" name="numberofruns" size="30" value="<%=numberofruns%>" readonly="true">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
				<script>
					function getnumberofruns(date_){
						
						var start_date = new Date(document.getElementById("startingdate").value);
						var numberofruns = document.getElementById("numberofruns");
						var end_date = new Date(date_);
						total_months = (end_date.getFullYear() - start_date.getFullYear())*12 + (end_date.getMonth() - start_date.getMonth());
						if (total_months < 0){
							numberofruns.value = 0;
						} else {
							numberofruns.value = total_months +1;
						}
						
						//alert(numberofruns);
					}
				</script>
		<tr>
			<td align="left">Amount</td>
			<td align="left">
				<input disabled style="background:#ccc;" type="text" name="amount" size="30" value="<%=amount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Is Use Settlement</td>
			<td align="left">
				<select name="isusesettlement">
					<option value="FALSE" <%IF isusesettlement = "FALSE" THEN%> selected <%END IF%>>FALSE</option>
					<option value="TRUE" <%IF isusesettlement = "TRUE" THEN%> selected <%END IF%>>TRUE</option>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Accrual Job</td>
			<td align="left">
				<select name="accrualjob">
				<%
					Dim qaccrualjob, rsaccrualjob
					qaccrualjob = " SELECT fccode, fcname FROM job WHERE fcba = '" & Session("fcBA") & "' AND upper(fcname) like upper('accrual %') ORDER BY fccode"
					Set rsaccrualjob = Server.CreateObject("ADODB.Recordset")
					rsaccrualjob.Open qaccrualjob, connstr
					IF NOT rsaccrualjob.EOF THEN
						WHILE NOT rsaccrualjob.EOF
				%>
				<option value="<%=rsaccrualjob("fccode")%>" <%IF accrualjob = rsaccrualjob("fccode") THEN%> selected <%END IF%>>
					<%=rsaccrualjob("fccode")%> : <%=rsaccrualjob("fcname")%>
				</option>
				<%
							rsaccrualjob.MOVENEXT
						WEND
					END IF
				%>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">State</td>
			<td align="left">
				<input disabled type="text" maxlength="10" name="state" size="30" value="<%=state%>">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit" name="Add" value="Select">
				<%If state = "Planned" Then%>
				<a href="recurring.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&documentno=<%=documentno%>&startingdate=<%=startingdate%>&finisheddate=<%=finisheddate%>&defaultexecuted=<%=defaultexecuted%>&intervalinmonth=<%=intervalinmonth%>&numberofruns=<%=numberofruns%>&amount=<%=amount%>&isusesettlement=<%=isusesettlement%>&accrualjob=<%=accrualjob%>&txtJenis=Header">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=documentno%>" border="0" onClick="return confirm('Are you sure to delete');">
				</a>
				<%Else%>
					<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
				<%End if%>
				<% if state = "Planned" then %>
					&nbsp;&nbsp;
					<a href="recurring.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&documentno=<%=documentno%>&startingdate=<%=startingdate%>&finisheddate=<%=finisheddate%>&defaultexecuted=<%=defaultexecuted%>&intervalinmonth=<%=intervalinmonth%>&numberofruns=<%=numberofruns%>&amount=<%=amount%>&isusesettlement=<%=isusesettlement%>&accrualjob=<%=accrualjob%>&state=<%=state%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to confirm this data :  <%=documentno%> ?');">
						Confirm
					</a>
					&nbsp;&nbsp;
				<% elseif state = "Confirmed" then %>
					&nbsp;&nbsp;
					<a href="recurring.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&documentno=<%=documentno%>&startingdate=<%=startingdate%>&finisheddate=<%=finisheddate%>&defaultexecuted=<%=defaultexecuted%>&intervalinmonth=<%=intervalinmonth%>&numberofruns=<%=numberofruns%>&amount=<%=amount%>&isusesettlement=<%=isusesettlement%>&accrualjob=<%=accrualjob%>&state=<%=state%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=documentno%> ?');">
						Re-Plan
					</a>
					&nbsp;&nbsp;
					<a href="recurring.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&documentno=<%=documentno%>&startingdate=<%=startingdate%>&finisheddate=<%=finisheddate%>&defaultexecuted=<%=defaultexecuted%>&intervalinmonth=<%=intervalinmonth%>&numberofruns=<%=numberofruns%>&amount=<%=amount%>&isusesettlement=<%=isusesettlement%>&accrualjob=<%=accrualjob%>&state=<%=state%>&txtJenis=Executed" onClick="return confirm('Are you sure to Execute this data, Please check your detail data before once Executed!, Document No:  <%=documentno%> ?');">
						Execute
					</a>
					&nbsp;&nbsp;
					<a href="recurring.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&documentno=<%=documentno%>&startingdate=<%=startingdate%>&finisheddate=<%=finisheddate%>&defaultexecuted=<%=defaultexecuted%>&intervalinmonth=<%=intervalinmonth%>&numberofruns=<%=numberofruns%>&amount=<%=amount%>&isusesettlement=<%=isusesettlement%>&accrualjob=<%=accrualjob%>&state=<%=state%>&txtJenis=Cancelled" onClick="return confirm('Are you sure to Cancel this data : <%=documentno%> ?');">
						Cancel
					</a>
					&nbsp;&nbsp;
		<% end if%>
			</td>
		</tr>
	</table>  
</form>
<%
IF cmdView = "True" AND "" & documentno <> "" THEN
	Dim rsrecurringdtl,intRecordsPerPage,cColor
	q = " SELECT rd.documentno," & _
	    "        rd.accountingyear," & _
	    "        rd.debetcredit," & _
	    "        rd.jobcode," & _
	    "        jb.fcname jobname," & _
	    "        rd.targettype," & _
	    "        rd.amount," & _
	    "        rd.fcba," & _
	    "        rd.status," & _
	    "        rd.fcentry," & _
	    "        rd.fcedit," & _
	    "        rd.fcip," & _
	    "        rd.lastupdate," & _
	    "        rd.lasttime" & _
	    " FROM RECURRINGDTL rd" & _
	    " LEFT JOIN job jb ON (rd.fcba = jb.fcba AND rd.jobcode = jb.fccode)" & _
	    " WHERE rd.documentno = '" & documentno & "'" & _
		"       AND rd.accountingyear = '" & GetYearofDate(finisheddate) & "'" & _
		"       AND rd.fcba = '" & Session("fcBA") & "'" & _
		" ORDER BY rd.jobcode, rd.targettype"
	'response.Write q	
	Set rsrecurringdtl = Server.CreateObject("ADODB.Recordset")
	rsrecurringdtl.CursorType = 3
	rsrecurringdtl.CursorLocation = 3
	rsrecurringdtl.Open q, connstr
	
%>
   
<form action="recurring.asp" method="post" name="eForm">
	<table class="data" width="100%">
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Job Code</font>
			</td>
			<td align="left">
				<input type="text" size="15" name="jobcode" value="<%=jobcode%>" <%if txtJenis = "Edit" then%>disabled<%end if%>>
				<span class="satuan_item" id="jobname"></span>
				<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Target Type</font>
			</td>
			<td align="left">
				<input type="text" size="15" name="targettype" value="<%=targettype%>" <%if txtJenis = "Edit" then%>disabled<%end if%>>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Debet/Credit</font>
			</td>
			<td align="left">
				<select name="debetcredit">
					<option value="Db" <%IF debetcredit = "Db" THEN%> selected <%END IF%>>DEBET</option>
					<option value="Cr" <%IF debetcredit = "Cr" THEN%> selected <%END IF%>>CREDIT</option>
				<select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">
				<font face="Verdana" size="1">Amount</font>
			</td>
			<td align="left">
				<input type="text" size="15" name="detailamount" value="<%=detailamount%>" data-inputmask="'alias': 'numeric', 'groupSeparator': ',', 'autoGroup': true, 'digits': 0, 'digitsOptional': false, 'placeholder': '0'" style="text-align: right;">
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
	    <TR>
			<td colspan="2" align="left">
				<input type="hidden" name="documentno" value="<%=documentno%>">
				<input type="hidden" name="documentname" value="<%=documentname%>">
				<input type="hidden" name="startingdate" value="<%=startingdate%>">
				<input type="hidden" name="finisheddate" value="<%=finisheddate%>">
				<input type="hidden" name="defaultexecuted" value="<%=defaultexecuted%>">
				<input type="hidden" name="intervalinmonth" value="<%=intervalinmonth%>">
				<input type="hidden" name="numberofruns" value="<%=numberofruns%>">
				<input type="hidden" name="amount" value="<%=amount%>">
				<input type="hidden" name="isusesettlement" value="<%=isusesettlement%>">
				<input type="hidden" name="accrualjob" value="<%=accrualjob%>">
				<input type="hidden" name="state" value="<%=state%>">
				<%if txtJenis = "Edit" then%>
				<input type="hidden" name="jobcode" value="<%=jobcode%>">
				<input type="hidden" name="targettype" value="<%=targettype%>">
				<%end if%>
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<%If state = "Planned" Then%>
				<%IF Request("txtJenis")="Edit" then%>
					<input type="hidden" name="txtJenis" value="Edit">
					<input Type="Submit"  name="Add" value="Update"  OnClick ="periksa(this.form)" >
				<%Else%>
					<input type="hidden" name="txtJenis" value="New">
					<input Type="Submit"  name="Add" value="Add" OnClick ="periksa(this.form)">
				<%End if%>
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
		<th>JOB CODE</th>
		<th>JOB NAME</th>
		<th>TARGET TYPE</th>
		<th>DEBET/CREDIT</th>
		<th>AMOUNT</th>		
	</tr>
<%
	'response.write Session("v_Employee")
   
	if Not rsrecurringdtl.eof then
		nCountofField=rsrecurringdtl.fields.count-1
		no=1
		While Not rsrecurringdtl.EOF
			if no  mod 2 = 1 then
				cColor = "#EEEEEE"
			else
				cColor = "#E5E5E5"
			end if
%>
	<tr bgcolor="<%=cColor%>">
		<td align="left" valign="top"><%=NO%></td>
		<%If state = "Planned" Then%>
			<td valign="middle" width="30">
				<a href="recurring.asp?documentno=<%=documentno%>&startingdate=<%=startingdate%>&finisheddate=<%=finisheddate%>&defaultexecuted=<%=defaultexecuted%>&intervalinmonth=<%=intervalinmonth%>&numberofruns=<%=numberofruns%>&amount=<%=amount%>&isusesettlement=<%=isusesettlement%>&accrualjob=<%=accrualjob%>&state=<%=state%>&jobcode=<%=rsrecurringdtl("jobcode")%>&targettype=<%=rsrecurringdtl("targettype")%>&debetcredit=<%=rsrecurringdtl("debetcredit")%>&detailamount=<%=rsrecurringdtl("amount")%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsrecurringdtl(3)%>" border="0">
				</a>
			</td>
			<td>
				<a href="recurring.asp?documentno=<%=documentno%>&startingdate=<%=startingdate%>&finisheddate=<%=finisheddate%>&defaultexecuted=<%=defaultexecuted%>&intervalinmonth=<%=intervalinmonth%>&numberofruns=<%=numberofruns%>&amount=<%=amount%>&isusesettlement=<%=isusesettlement%>&accrualjob=<%=accrualjob%>&state=<%=state%>&jobcode=<%=rsrecurringdtl("jobcode")%>&targettype=<%=rsrecurringdtl("targettype")%>&debetcredit=<%=rsrecurringdtl("debetcredit")%>&txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsrecurringdtl(3)%>" border="0">
				</a>
			</td>
		<%Else%>
			<td colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
		<%End If%>
		<td align="left" title="<%=rsrecurringdtl("jobcode").Name%>" valign="top"><%=rsrecurringdtl("jobcode")%></td>
		<td align="left" title="<%=rsrecurringdtl("jobname").Name%>" valign="top"><%=rsrecurringdtl("jobname")%></td>
		<td align="left" title="<%=rsrecurringdtl("targettype").Name%>" valign="top"><%=rsrecurringdtl("targettype")%></td>
		<td align="left" title="<%=rsrecurringdtl("debetcredit").Name%>" valign="top"><%=rsrecurringdtl("debetcredit")%></td>
		<td align="right" title="<%=rsrecurringdtl("amount").Name%>" valign="top"><%=FormatNumber(rsrecurringdtl("amount"), 0)%></td>
	</tr>
<%
			rsrecurringdtl.MoveNext
			intRecordsPerPage = intRecordsPerPage - 1
			no=no+1
		Wend
	else
		response.write ("<tr bgcolor='#E5E5E5'><td colspan='9' align='left'>Data is empty</td></tr>")
	end if 
	rsrecurringdtl.Close
	set rsrecurringdtl = Nothing
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
