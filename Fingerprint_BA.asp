<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!--#include file="DailyValidation.inc"-->

<html>
<head>
<title>ePlaS - Berita Acara Fingerprint</title>
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
</script>
<script type='text/javascript'>
    //Replaces all non numeric characters
    function checkInput(element){ element.value = element.value.replace(/[^0-9\.\-]/g,'');}
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
'Fungsi cek employeecode sudah dateterminate atau belum
Function CheckAvailableEmployee(initialEmployee, initialFcba)
	Dim initialAvailable
	Dim rsAvailable, strAvailable
	
	strAvailable = " Select Count(1) Total" & _
	               " From employee" & _
				   " Where Fccode = '" & Trim(initialEmployee) & "'" & _
				   "       And Fcba = '" & Trim(initialFcba) & "'" & _
				   "       and dateterminate is null"

	Set rsAvailable = Server.CreateObject("ADODB.Recordset")
	rsAvailable.Open strAvailable, connstr, 3
	
	If Not rsAvailable.EOF Then
		initialAvailable = Cdbl(rsAvailable("Total"))
	Else
		initialAvailable = 0
	End If
	
	strAvailable =	""
	rsAvailable.Close
	set rsAvailable = Nothing

	CheckAvailableEmployee = initialAvailable
End Function
%>
 
 <%
 '-------------awal-----------------------
 Dim FDDATE,PeriodDate,PeriodMonth, Header_Allowded, empCode, AllowdedCode, fcCode
 Dim Amount, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, txtLineNo
 Dim cmdView, rsBA, qBA
 Dim qCheckStatus, txtClosingStatus, rsCheckStatus
 Dim valueAorD, objectRS, objectSQL
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
 
Dim v_txtYear : v_txtYear = Request("txtYear")
Dim v_txtMonth : v_txtMonth = Request("txtMonth")

Dim cek_ada_employee

Dim MinYr, MaxYr, mon, Yr
Yr = Year(Session("v_Employee"))
mon = Month(Session("v_Employee"))
MinYr = Yr - 2
MaxYr = Yr  + 2

If len(v_txtyear)= 0 then v_txtyear = Yr
If len(v_txtmonth) = 0 then v_txtmonth = mon

	
	PeriodMonth = request("PeriodMonth")
	PeriodDate = request("PeriodDate")
	Header_Allowded = request("Header_Allowded")
	
	empCode = request("empCode")
	fddate = request("fdDate")
	AllowdedCode = request("AllowdedCode")
	Amount = request("Amount")
	REMARKS = request("REMARKS")
	
	txtLineNo = request("txtLineNo")
 
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	qBA =	"select nvl(count(*),0) as JlhRecord from AllowDed where " & _
			" FCCODE='" & AllowdedCode & "' and MAN_AUTO='M' and fcBA='" & Session("fcBA") & "'" & _
			" and (CROP='" & Session("Crop") & "' or Crop='ALL')"
	Set rsBA = Server.CreateObject("ADODB.Recordset")
	rsBA.Open qBA, connstr,3,3
	if  cdbl(rsBA("JlhRecord")) > 0 then
		if Request("txtJenis")= "Delete" then
		
			'-- Check for Empallded PreProcess
			qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreAttendanceUpdate (to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'YYYY'))," & _ 
																				  " to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'MM')),'" & _ 
																					Session("fcBA") & "') closingStatus from dual "

			Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
			rsCheckStatus.CursorType = 3
			rsCheckStatus.CursorLocation = 3
			rsCheckStatus.Open qCheckStatus, connstr
			if not rsCheckStatus.eof then
			   txtClosingStatus = rsCheckStatus("closingStatus")
			else
			   Response.Write  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
			   txtClosingStatus="FALSE"
			end if
			rsCheckStatus.close
			set rsCheckStatus = Nothing



			IF txtClosingStatus = "TRUE" THEN
			
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fddate)
				If Cek_IsDailyClosing = "FALSE" then
		
					q	=	" Delete from EmpAllDed where" & _
							" ADHOCDATE = to_date('" & fddate & "','mm/dd/yyyy')" & _
							" and EMPCODE = '" & EMPCODE & "'" & _
							" And AllowdedCode = '" & AllowdedCode & "'" & _
							" And LINENO = '" & txtLineNo & "'" & _
							" and fcBA = '" & Session("fcBA") & "'"
					connStr.execute(q)
					
				else
					err.number = 1001
					err.description = "Tanggal """ & fddate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
				cmdView	= "True"
			else
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Unable to Update/Delete/Insert," & "</font>"
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Please OPEN your payroll Status </font>"
				cmdView	= "True"			
			end if				
				
		elseif Request("txtJenis")="New" then
			Dim rsGetMaxLineNo, strGetMaxLineNo, valueMaxLineNo
			
			'-- Check for Empallded PreProcess
			qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreAttendanceUpdate (to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'YYYY'))," & _ 
																				  " to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'MM')),'" & _ 
																					Session("fcBA") & "') closingStatus from dual "

			Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
			rsCheckStatus.CursorType = 3
			rsCheckStatus.CursorLocation = 3
			rsCheckStatus.Open qCheckStatus, connstr
			if not rsCheckStatus.eof then
			   txtClosingStatus = rsCheckStatus("closingStatus")
			else
			   Response.Write  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
			   txtClosingStatus="FALSE"
			end if
			rsCheckStatus.close
			set rsCheckStatus = Nothing
			
			IF txtClosingStatus = "TRUE" THEN
			
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fddate)
				If Cek_IsDailyClosing = "FALSE" then
			
					If Len(Trim(empCode)) > 0 Then
					
						objectSQL	=	" select allded" & _
										" from allowded" & _
										" where fccode = '" & AllowdedCode & "'" & _
										"       and fcba = '" & Session("FCBA") & "'"
						
						Set objectRS = Server.CreateObject("ADODB.Recordset")
						objectRS.Open objectSQL, connstr
						
						If Not objectRS.Eof Then
							valueAorD = objectRS("allded")
						Else
							valueAorD = ""
						End If
						objectRS.close
						set objectRS = Nothing
						
						If (valueAorD = "A" And Cdbl(Amount) < 0) then
							err.number = 10001
							err.description = "Allowded " & AllowdedCode & " adalah Pendapatan, maka Amount/ Rupiah harus Positif! tidak boleh Negatif. Silahkan Cek angka yang anda input!"
						Else
						
							cek_ada_employee = CheckAvailableEmployee(empCode, Session("FCBA"))
							
							if cek_ada_employee <> 0 then
				
								strGetMaxLineNo	=	" Select Nvl(Max(LINENO), 0) LineNo From Empallded Where Fcba = '" & Session("FCBA") & "'"
								
								Set rsGetMaxLineNo = Server.CreateObject("ADODB.Recordset")
								rsGetMaxLineNo.Open strGetMaxLineNo, connstr
								
								If Not rsGetMaxLineNo.Eof Then
									valueMaxLineNo = Cdbl(rsGetMaxLineNo("LineNo")) + 1
								Else
									valueMaxLineNo = 1
								End If
								
								q	=	" Insert into EmpAllDed (" & _
										" EMPCODE" & _
										" , ALLOWDEDCODE" & _
										" , ADHOCDATE" & _
										" , AMOUNT" & _
										" , FIXEDALLDED" & _
										" , FCENTRY" & _
										" , FCEDIT" & _
										" , FCIP" & _
										" , FCBA" & _
										" , REMARKS" & _
										" , LASTUPDATE" & _
										" , LASTTIME" & _
										" , LINENO" & _
										" ) " & _
										" values (" & _
										" '" & empCode & "'" & _
										" , '" & AllowdedCode & "'" & _
										" , to_date('" & fddate & "','mm/dd/yyyy') " & _
										" , " & Amount & "" & _
										" , 0" & _
										" , '" & Session("CodeName") & "'" & _
										" , '" & Session("CodeName") & "'" & _
										" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
										" , '" & Session("fcBA") & "'" & _
										" , '" & Remarks & "'" & _
										" , sysdate" & _
										" , To_Char(sysdate,'HH24:MI')" & _
										" , " & valueMaxLineNo & "" & _
										" )"
										
								'Response.Write q
								connStr.execute(q)
							else
								err.description = "Karyawan " & empCode & ", tidak ada di dalam master employee/ sudah dateterminate. Jadi kenapa di input premi manualnya?!!!"
								err.number = 10001
							end if
							
						End If
					Else
						err.description = "Kode Karyawan tidak boleh kosong"
						err.number = 10001
					End If
					
				else
					err.number = 1001
					err.description = "Tanggal """ & fddate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
				cmdView	= "True"
			else
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Unable to Update/Delete/Insert," & "</font>"
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Please OPEN your payroll Status </font>"
				cmdView	= "True"			
			end if
			
		elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
			'-- Check for Empallded PreProcess
			qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreAttendanceUpdate (to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'YYYY'))," & _ 
																				  " to_number(to_char(to_date('" & fdDate & "','mm/dd/yyyy'),'MM')),'" & _ 
																					Session("fcBA") & "') closingStatus from dual "

			Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
			rsCheckStatus.CursorType = 3
			rsCheckStatus.CursorLocation = 3
			rsCheckStatus.Open qCheckStatus, connstr
			if not rsCheckStatus.eof then
			   txtClosingStatus = rsCheckStatus("closingStatus")
			else
			   Response.Write  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
			   txtClosingStatus="FALSE"
			end if
			rsCheckStatus.close
			set rsCheckStatus = Nothing
			
			IF txtClosingStatus = "TRUE" THEN  
			
				Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), fddate)
				If Cek_IsDailyClosing = "FALSE" then
			
					objectSQL	=	" select allded" & _
									" from allowded" & _
									" where fccode = '" & AllowdedCode & "'" & _
									"       and fcba = '" & Session("FCBA") & "'"
					
					Set objectRS = Server.CreateObject("ADODB.Recordset")
					objectRS.Open objectSQL, connstr
					
					If Not objectRS.Eof Then
						valueAorD = objectRS("allded")
					Else
						valueAorD = ""
					End If
					objectRS.close
					set objectRS = Nothing
					
					If (valueAorD = "A" And Cdbl(Amount) < 0) then
						err.number = 10001
						err.description = "Allowded " & AllowdedCode & " adalah Pendapatan, maka Amount/ Rupiah harus Positif! tidak boleh Negatif. Silahkan Cek angka yang anda input!"
						cmdView	= "True"
					Else
					
						cek_ada_employee = CheckAvailableEmployee(empCode, Session("FCBA"))
							
							if cek_ada_employee <> 0 then
				
								q = " Update EmpAllDed" & _
									" set AMOUNT = " & Amount & "" & _
									"     , Remarks = '" & Remarks & "'" & _
									"     , LASTUPDATE = sysdate " & _
									"     , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
									" where ADHOCDATE = TO_DATE('" & fddate & "','mm/dd/yyyy')" & _
									"       and EMPCODE = '" & EMPCODE & "'" & _
									"       And AllowdedCode = '" & AllowdedCode & "'" & _
									"       And LINENO = '" & txtLineNo & "'" & _
									"       and fcBA = '" & Session("fcBA") & "'"
								
								'Response.Write q
								connStr.execute(q)
								connStr.execute("commit")
							else
								err.description = "Karyawan " & empCode & ", tidak ada di dalam master employee/ sudah dateterminate. Jadi kenapa di edit premi manualnya?!!!"
								err.number = 10001
							end if
						cmdView	= "True"
					End If
					
				else
					err.number = 1001
					err.description = "Tanggal """ & fddate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
				End If
				cmdView	= "True"
			else
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Unable to Update/Delete/Insert," & "</font>"
				response.Write "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Please OPEN your payroll Status </font>"
				cmdView	= "True"
			end if				
		else
			cmdView	= "True"
		end if	
	 else
		err.description= "Your are Allowance is Automatically or Allowance not your Crop Type... "
		err.number=10001
		cmdView	= "True"
	 end if
	 rsBA.close
	 set rsBA = Nothing
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if


Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
if len(trim(fddate)) <> 0 then
	fddate = Request("fddate")
else
	fddate = textdate(v_txtMonth & "/" & GetLastDay(v_txtMonth, v_txtYear) & "/" & v_txtYear)
end if

if Request("Amount")<>"" then
   Amount = Request("Amount")
else
   Amount = 0
end if

txtTable = Request("txtTable")
txtModule = Request("txtModule")
txtSearch = Request("txtSearch")
txtField = Request("txtField")
txtJenis = request("txtJenis")

if err.number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if
 '-------------ISI Payment Form
 
Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & v_txtmonth & "/01/" & v_txtyear & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & v_txtmonth & "/01/" & v_txtyear & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
	        "       and upper(system) = upper('Employee')"
			
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
 <form name="frmEntry" method="post" action="EmpAllDed.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Year/ Tahun</td>
			<td align="left">
				<select name="txtyear">
				<%For i = MinYr to MaxYr%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtyear) then%> Selected <%End If%>><%=i%></option>
				<%Next%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Month/ Bulan</td>
			<td align="left">
				<select name="txtmonth">
				<%For i = 1 to 12%>
					<Option Value="<%=i%>" <%If cint(i) = cint(v_txtmonth) then%> Selected <%End If%>><%=textMonth(i)%></option>
				<%Next%>
				</select>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;				
			</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" Then
	Dim rsVH, intRecordsPerPage, cColor
	set rsVH = server.CreateObject("Adodb.recordset")
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " select empcode" & _
		"		 ,  employee_api.get_EmployeeName (fcba,empcode) empname" & _
        "        , allowdedcode" & _
        "        , adhocdate" & _
        "        , amount" & _
        "        , remarks" & _
        "        , lineno" & _
        "        , fcentry" & _
        "        , fcedit" & _
        "        , fcip" & _
        "        , fcba" & _
        " from empallded" & _
		" where fcba = '" & Session("Fcba") & "'" & _
		"       and to_number(to_char(adhocdate,'yyyy')) = '" & v_txtYear & "'" & _
		"       and to_number(to_char(adhocdate,'mm')) = '" & v_txtMonth & "'" & _
		"       and allowded_api.get_man_auto(allowdedcode, fcba) = 'M'" & _
		"       and allowded_api.getincludedpayrollprocessing(allowdedcode, fcba) = 'TRUE'"
		
	if trim(Header_Allowded) <> "" or len(trim(Header_Allowded)) > 0 then
		q = q & " and allowdedcode = '" & Header_Allowded & "'"
	end if

	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
	
	' response.write q
%>
	<form action="EmpAllDed.Asp" method="post" name="eForm">
	<table class="data" cellpadding="0" cellspacing="0" width="100%">
		<!--Yang tidak bisa di edit adalah PK Table-->
		<tr>
			<td>EmpCode</td>
			<td align="left">
			<%if txtJenis = "Edit" then%>
				<input type="text" size="15" name="empCode" value="<%=empCode%>" readonly="True" style="background:#999999">
			<%else%>
				<input type="text" size="15" name="empCode" value="<%=empCode%>">
			<%end if%>
				<font color="#FF0000">&nbsp;*</font>
				<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Lookup Employee</a>
			</td>
		</tr>
		<tr>
			<td>Date</td>
			<td align="left">
			<%if txtJenis = "Edit" then%>
				<input class="datepicker" type="text" size="15" name="fdDate" value="<%=fdDate%>" readonly="True" style="background:#999999">
			<%else%>
				<input class="datepicker" type="text" size="15" name="fdDate" value="<%=fdDate%>">
			<%end if%>
				<font color="#FF0000">&nbsp;*</font>
			</td>
		</tr>
		<tr>
			<td>No.Berita Acara</td>
			<td align="left">
				<input type="text" size="50" name="REMARKS" value="<%=REMARKS%>">
			</td>
		</tr>
		<tr>
			<td>Remarks</td>
			<td align="left">
				<input type="text" size="50" name="REMARKS" value="<%=REMARKS%>">
			</td>
		</tr>
	    <tr>
			<td colspan="2" align="left">
			<input type="hidden" name="txtyear" value="<%=v_txtYear%>">
			<input type="hidden" name="txtMonth" value="<%=v_txtMonth%>">
			<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
			<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
			<!-- Period Control In Here -->
			<%IF Request("txtJenis")="Edit" then%>
				<input type="hidden" name="txtJenis" value="Edit">
				<input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
			<%Else%>
			   <input type="hidden" name="txtJenis" value="New">
			   <input Type="Submit"  name="Add" value="Add">
			<%End if%>
			<!-- Else Period Control -->
			<input Type="Reset"  name="Reset" value="Reset">
			<input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		</tr>
   </table>

   <table class="data" width="100%">
 <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">EMPLOYEE CODE</TD>	
	<TD align="left">EMPLOYEE NAME</TD>	
	<TD align="left">DATE</TD>	
	<TD align="left">No.BA</TD>	
	<TD align="left">REMARKS</TD>
	</tr>   
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
		  <td align="Right" valign="top" width="30"><%=NO%></td>
			<!-- Period Control In Here -->
				<td valign="middle" width="10">
					<a href="EmpAllDed.Asp?empCode=<%=rsVH("EMPCODE")%>&fdDate=<%=rsVH("ADHOCDATE")%>&AllowdedCode=<%=rsVH("ALLOWDEDCODE")%>&AMOUNT=<%=rsVH("AMOUNT")%>&Remarks=<%=rsVH("Remarks")%>&txtLineNo=<%=rsVH("LineNo")%>&PeriodMonth=<%=PeriodMonth%>&PeriodDate=<%=PeriodDate%>&Header_Allowded=<%=Header_Allowded%>&txtJenis=Edit">
					<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
				</td>
				<td width="10">
					<a href="EmpAllDed.Asp?empCode=<%=rsVH("EMPCODE")%>&fdDate=<%=rsVH("ADHOCDATE")%>&AllowdedCode=<%=rsVH("ALLOWDEDCODE")%>&txtLineNo=<%=rsVH("LineNo")%>&PeriodMonth=<%=PeriodMonth%>&PeriodDate=<%=PeriodDate%>&Header_Allowded=<%=Header_Allowded%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');">
					<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0"></a>
				</td>
			<!-- Else Period COntrol -->
				 <!-- <td width="10"colspan="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td> -->
			<!-- End Period Control -->	 

		    <%for i = 0 to nCountofField - 5%>
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
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>