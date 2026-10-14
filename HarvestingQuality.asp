<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Harvesting Quality</title>
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
   <div class="judul_head"  valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 <div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;"> 
 <!-------------awal------------->
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
DIM PeriodDate,	EMPCODE,FDDATE,FIELDCODE,UNDER_RIPE,OVER_RIPE,ABNORMAL,LONG_STALK,EATEN_BY_RAT,UNHARVEST_FFB,UNCOLLECT_LF_CIRCLE,UNCOLLECT_LF_PIECE
DIM UNARRANGE_FFB,UNPRUNE_FROND,QE_1,QE_2,QE_3,QE_4,QE_5,QE_6,QE_7,QE_8,QE_9,QE_10,QE_11
Dim cmdView

Dim Header_EmpCode
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

PeriodDate =request("PeriodDate")
Header_EmpCode=request("Header_EmpCode")
EMPCODE=request("EMPCODE")
FIELDCODE=request("FIELDCODE")


if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

UNDER_RIPE=To_Nol(request("UNDER_RIPE"))
OVER_RIPE=To_Nol(request("OVER_RIPE"))
ABNORMAL=To_Nol(request("ABNORMAL"))
LONG_STALK=To_Nol(request("LONG_STALK"))
EATEN_BY_RAT=To_Nol(request("EATEN_BY_RAT"))
UNHARVEST_FFB=To_Nol(request("UNHARVEST_FFB"))
UNCOLLECT_LF_CIRCLE=To_Nol(request("UNCOLLECT_LF_CIRCLE"))
UNCOLLECT_LF_PIECE=To_Nol(request("UNCOLLECT_LF_PIECE"))
UNARRANGE_FFB=To_Nol(request("UNARRANGE_FFB"))
UNPRUNE_FROND=To_Nol(request("UNPRUNE_FROND"))
QE_1=To_Nol(request("QE_1"))
QE_2=To_Nol(request("QE_2"))
QE_3=To_Nol(request("QE_3"))
QE_4=To_Nol(request("QE_4"))
QE_5=To_Nol(request("QE_5"))
QE_6=To_Nol(request("QE_6"))
QE_7=To_Nol(request("QE_7"))
QE_8=To_Nol(request("QE_8"))
QE_9=To_Nol(request("QE_9"))
QE_10=To_Nol(request("QE_10"))

	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PeriodDate)
		If Cek_IsDailyClosing = "FALSE" then
		
			q	=	" Delete from HarvestingQuality where" & _
					" EMPCODE = '" & EMPCODE & "'" & _
					" And FIELDCODE = '" & FIELDCODE & "'" & _
					" and fddate = to_Date('" & PeriodDate & "','MM/DD/YYYY')" & _
					" and fcBA = '" & Session("fcBA") & "'"
					
			'response.Write(q)
			connStr.execute(q)
			
		else
			err.number = 1001
			err.description = "Tanggal """ & PeriodDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="New" then	
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PeriodDate)
		If Cek_IsDailyClosing = "FALSE" then
		
			q	=	" Insert into HarvestingQuality (" & _
					" EMPCODE" & _
					" , FDDATE" & _
					" , FIELDCODE" & _
					" , UNDER_RIPE" & _
					" , OVER_RIPE" & _
					" , ABNORMAL" & _
					" , LONG_STALK" & _
					" , EATEN_BY_RAT" & _
					" , UNHARVEST_FFB" & _
					" , UNCOLLECT_LF_CIRCLE" & _
					" , UNCOLLECT_LF_PIECE" & _
					" , UNARRANGE_FFB" & _
					" , UNPRUNE_FROND" & _
					" , QE_1" & _
					" , QE_2" & _
					" , QE_3" & _
					" , QE_4" & _
					" , QE_5" & _
					" , QE_6" & _
					" , QE_7" & _
					" , QE_8" & _
					" , QE_9" & _
					" , QE_10" & _
					" , FCENTRY" & _
					" , FCEDIT" & _
					" , FCIP" & _
					" , FCBA" & _
					" , LASTUPDATE" & _
					" , LASTTIME" & _
					" ) " & _
					" values ("& _
					" '" & EMPCODE & "'," & _
					" to_date('" & PeriodDate & "','mm/dd/yyyy')" & _
					" , '" & FIELDCODE & "'" & _
					" , '" & UNDER_RIPE & "'" & _
					" , '" & OVER_RIPE & "'" & _
					" , '" & ABNORMAL & "'" & _
					" , '" & LONG_STALK & "'" & _
					" , '" & EATEN_BY_RAT & "'" & _
					" , '" & UNHARVEST_FFB & "'" & _
					" , '" & UNCOLLECT_LF_CIRCLE & "'" & _
					" , '" & UNCOLLECT_LF_PIECE & "'" & _
					" , '" & UNARRANGE_FFB & "'" & _
					" , '" & UNPRUNE_FROND & "'" & _
					" , '" & QE_1 & "'" & _
					" , '" & QE_2 & "'" & _
					" , '" & QE_3 & "'" & _
					" , '" & QE_4 & "'" & _
					" , '" & QE_5 & "'" & _
					" , '" & QE_6 & "'" & _
					" , '" & QE_7 & "'" & _
					" , '" & QE_8 & "'" & _
					" , '" & QE_9 & "'" & _
					" , '" & QE_10 & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Session("CodeName") & "'" & _
					" , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					" , '" & Session("fcBA") & "'" & _
					" , sysdate" & _
					" , To_Char(sysdate,'HH24:MI')" & _
					" )"
					
				'response.Write(q)
				response.Flush()
				connStr.execute(q)
				
		else
			err.number = 1001
			err.description = "Tanggal """ & PeriodDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
			cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PeriodDate)
		If Cek_IsDailyClosing = "FALSE" then
			
			q	=	" Update HarvestingQuality set" & _
					" EMPCODE = '" & EMPCODE & "'" & _
					" , FDDATE = to_date('" & PeriodDate & "','mm/dd/yyyy')" & _
					" , FIELDCODE = '" & FIELDCODE & "'" & _
					" , UNDER_RIPE = '" & UNDER_RIPE & "'" & _
					" , OVER_RIPE = '" & OVER_RIPE & "'" & _
					" , ABNORMAL = '" & ABNORMAL & "'" & _
					" , LONG_STALK = '" & LONG_STALK & "'" & _
					" , EATEN_BY_RAT = '" & EATEN_BY_RAT & "'" & _
					" , UNHARVEST_FFB = '" & UNHARVEST_FFB & "'" & _
					" , UNCOLLECT_LF_CIRCLE = '" & UNCOLLECT_LF_CIRCLE & "'" & _
					" , UNCOLLECT_LF_PIECE = '" & UNCOLLECT_LF_PIECE & "'" & _
					" , UNARRANGE_FFB = '" & UNARRANGE_FFB & "'" & _
					" , UNPRUNE_FROND = '" & UNPRUNE_FROND & "'" & _
					" , QE_1 = '" & QE_1 & "'" & _
					" , QE_2 = '" & QE_2 & "'" & _
					" , QE_3 = '" & QE_3 & "'" & _
					" , QE_4 = '" & QE_4 & "'" & _
					" , QE_5 = '" & QE_5 & "'" & _
					" , QE_6 = '" & QE_6 & "'" & _
					" , QE_7 = '" & QE_7 & "'" & _
					" , QE_8 = '" & QE_8 & "'" & _
					" , QE_9 = '" & QE_9 & "'" & _
					" , QE_10 = '" & QE_10 & "'" & _
					" , LASTUPDATE = sysdate " & _
					" , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					" where" & _
					" EMPCODE = '" & EMPCODE & "'" & _
					" And FIELDCODE = '" & FIELDCODE & "'" & _
					" and fddate = to_Date('" & PeriodDate & "','MM/DD/YYYY')" & _
					" and fcBA = '" & Session("fcBA") & "'"
					
					'response.write q
					connStr.execute(q)
					connStr.execute("commit")
					
		else
			err.number = 1001
			err.description = "Tanggal """ & PeriodDate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
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
if Request("PeriodDate")<>"" then
   PeriodDate = Request("PeriodDate")
else
   PeriodDate = textDate(date)
end if

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")
If Err.Number <> 0 then
   response.Write  "<div class = ""error"">" & err.description & "</div>"
End if



Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & PeriodDate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & PeriodDate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
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

<!-------------ISI Payment Form------------->
 <form name="frmEntry" method="post" action="HarvestingQuality.asp">
	<table class="data" width="100%">
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<input  class="datepicker" type="text" maxlength="10" name="PeriodDate" size="10" value="<%=PeriodDate%>">
			</td>
		</tr>
		<tr>
			<td align="left">Employee Code</td>
			<td align="left">
				<input type="text" maxlength="10" name="Header_EmpCode" size="10" value="<%=Header_EmpCode%>">&nbsp;&nbsp;<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Employee<font face="Verdana" size="1"></font></a>
			</td>
		</tr>	
		<tr>
			<td colspan="2"  align="left">
				<input Type="Hidden" Value="True" Name="cmdView">
				<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
				<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
				<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;<a href="PrintHarvestingQuality.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&PeriodDate=<%=PeriodDate%>" title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
				<!-- HADIR -->
				<%
				dim dataModule,dataSubmodule,dataDetail
				dataModule 		= "Harvesting"
				dataSubmodule 	= "Harvesting Quality"
				dataDetail 		= "NULL"
				%>
				<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
				<!-- END:HADIR -->
			</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" Then
	Dim rsVH,intRecordsPerPage,cColor
	set rsVH = server.CreateObject("Adodb.recordset")
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q	=	" select" & _
			" EMPCODE" & _
			" , FDDATE" & _
			" , FIELDCODE" & _
			" , UNDER_RIPE" & _
			" , OVER_RIPE" & _
			" , ABNORMAL" & _
			" , LONG_STALK" & _
			" , EATEN_BY_RAT" & _
			" , UNHARVEST_FFB" & _
			" , UNCOLLECT_LF_CIRCLE" & _
			" , UNCOLLECT_LF_PIECE" & _
			" , UNARRANGE_FFB" & _
			" , UNPRUNE_FROND" & _
			" , QE_1" & _
			" , QE_2" & _
			" , QE_3" & _
			" , QE_4" & _
			" , QE_5" & _
			" , QE_6" & _
			" , QE_7" & _
			" , QE_8" & _
			" , QE_9" & _
			" , QE_10" & _
			" , FCENTRY" & _
			" , FCEDIT" & _
			" , FCIP" & _
			" , FCBA" & _
			" , LASTUPDATE" & _
			" , LASTTIME" & _
			" from HarvestingQuality where" & _
			" fcBA = '" & Session("fcBA") & "'"

If Len(Trim(PeriodDate)) > 0 Then
	q	=	q & " and fddate = to_Date('" & PeriodDate & "','MM/DD/YYYY')"
End If

If Len(Trim(Header_EmpCode)) > 0 Then
	q	=	q & " and EMPCODE = '" & Header_EmpCode & "'"
End If

	'response.Write q
	Set rsVH = Server.CreateObject("ADODB.Recordset")
	rsVH.CursorType = 3
	rsVH.CursorLocation = 3
	rsVH.Open q, connstr
%>
   <form action="HarvestingQuality.Asp" method="post" name="eForm" style="display:table;">
   <table class="data" cellpadding="0" cellspacing="0" style=" width:850px;">
		<tr>
		<%IF txtJenis="Edit" Then%>
		<td align="left">Employee Code</td><td align="left"><input type="text" size="15" name="EMPCODE" value="<%=EMPCODE%>" readonly="True" style="background:#999999">&nbsp;<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Employee<font face="Verdana" size="1"></font></a></td>
		<%else%>
		<td align="left">Employee Code</td><td align="left"><input type="text" size="15" name="EMPCODE" value="<%=EMPCODE%>">&nbsp;<a href="javascript:launchwin('PrintData.Asp?txtModule=Data Master&txtTable=Employee&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">Employee<font face="Verdana" size="1"></font></a></td>
		<%end if%>
		<td align="left">Pelepah Sengkleh</td><td align="left"><input type="text" size="15" name="UNPRUNE_FROND" value="<%=UNPRUNE_FROND%>" onkeyup="checkInput(this);"></td>
		</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
		<td align="left">Field Code</td><td align="left"><input type="text" size="15" name="FIELDCODE" value="<%=FIELDCODE%>" readonly="True" style="background:#999999"></td>
		<%else%>
		<td align="left">Field Code</td><td align="left"><input type="text" size="15" name="FIELDCODE" value="<%=FIELDCODE%>"></td>
		<%end if%>
		<td align="left">Pelepah Tidak Disusun Rapi</td><td align="left"><input type="text" size="15" name="QE_1" value="<%=QE_1%>" onkeyup="checkInput(this);"></td>
		</tr>
		<tr>
		<td align="left">Buah Mentah</td><td align="left"><input type="text" size="15" name="UNDER_RIPE" value="<%=UNDER_RIPE%>" onkeyup="checkInput(this);"></td>
		<td align="left">Buah Matahari</td><td align="left"><input type="text" size="15" name="QE_2" value="<%=QE_2%>" onkeyup="checkInput(this);"></td>
		</tr>
		<tr>
		<td align="left">Buah Terlalu Matang</td><td align="left"><input type="text" size="15" name="OVER_RIPE" value="<%=OVER_RIPE%>" onkeyup="checkInput(this);"></td>
		<td align="left">Buah Busuk</td><td align="left"><input type="text" size="15" name="QE_3" value="<%=QE_3%>" onkeyup="checkInput(this);"></td>
		</tr>
		<tr>
		<td align="left">Abnormal</td><td align="left"><input type="text" size="15" name="ABNORMAL" value="<%=ABNORMAL%>" onkeyup="checkInput(this);"></td>
		<td align="left">Buah Mentah di Peram</td><td align="left"><input type="text" size="15" name="QE_4" value="<%=QE_4%>" onkeyup="checkInput(this);"></td>
		</tr>
		<tr>
		<td align="left">Tangkai Panjang</td><td align="left"><input type="text" size="15" name="LONG_STALK" value="<%=LONG_STALK%>" onkeyup="checkInput(this);"></td>
		<td align="left">Over Pruning</td><td align="left"><input type="text" size="15" name="QE_5" value="<%=QE_5%>" onkeyup="checkInput(this);"></td>
		</tr>
		<tr>
		<td align="left">Buah Dimakan Tikus</td><td align="left"><input type="text" size="15" name="EATEN_BY_RAT" value="<%=EATEN_BY_RAT%>" onkeyup="checkInput(this);"></td>
		<td align="left">Brondolan Tidak di Alas</td><td align="left"><input type="text" size="15" name="QE_6" value="<%=QE_6%>" onkeyup="checkInput(this);"></td>
		</tr>
		<td align="left">Buah Tidak Dipanen</td><td align="left"><input type="text" size="15" name="UNHARVEST_FFB" value="<%=UNHARVEST_FFB%>" onkeyup="checkInput(this);"></td>
		<td align="left">Brondolan Kotor/ Sampah</td><td align="left"><input type="text" size="15" name="QE_7" value="<%=QE_7%>" onkeyup="checkInput(this);"></td>
		<tr>
		<td align="left">Brondolan Dipiringan</td><td align="left"><input type="text" size="15" name="UNCOLLECT_LF_CIRCLE" value="<%=UNCOLLECT_LF_CIRCLE%>" onkeyup="checkInput(this);"></td>
		<td align="left">Buah di Belah</td><td align="left"><input type="text" size="15" name="QE_8" value="<%=QE_8%>" onkeyup="checkInput(this);"></td>
		</tr>
		<tr>
		<td align="left">Buah Tinggal Dipiringan</td><td align="left"><input type="text" size="15" name="UNCOLLECT_LF_PIECE" value="<%=UNCOLLECT_LF_PIECE%>" onkeyup="checkInput(this);"></td>
		<td align="left">Buah Tidak di Antrikan/ di Tulis</td><td align="left"><input type="text" size="15" name="QE_9" value="<%=QE_9%>" onkeyup="checkInput(this);"></td>
		</tr>
		<tr>
		<td align="left">Buah Tidak Disusun Rapi di TPH</td><td align="left"><input type="text" size="15" name="UNARRANGE_FFB" value="<%=UNARRANGE_FFB%>" onkeyup="checkInput(this);"></td>
		<td align="left">Buah Tidak Diletakan Di TPH</td><td align="left"><input type="text" size="15" name="QE_10" value="<%=QE_10%>" onkeyup="checkInput(this);"></td>
		</tr>
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="PeriodDate" value="<%=PeriodDate%>">
		  <input type="hidden" name="Header_EmpCode" value="<%=Header_EmpCode%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%If val_IsTrue = "Y" Then%>
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
		  <%Else%>
		       <input type="hidden" name="txtJenis" value="New">
		       <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">NIK</TD>	
	<TD align="left">Tanggal</TD>	
	<TD align="left">Blok</TD>	
	<TD align="left">Buah Mentah</TD>	
	<TD align="left">Buah Terlalu Matang</TD>	
	<TD align="left">Abnormal</TD>	
	<TD align="left">Tangkai Panjang</TD>	
	<TD align="left">Buah Dimakan Tikus</TD>	
	<TD align="left">Buah Tidak Dipanen</TD>	
	<TD align="left">Brondolan Dipiringan</TD>	
	<TD align="left">Buah Tinggal Dipiringan</TD>	
	<TD align="left">Buah Tidak Disusun Rapi</TD>	
	<TD align="left">Pelepah Sengkleh</TD>	
	<TD align="left">Pelepah Tidak Disusun Rapi</TD>	
	<TD align="left">Buah Matahari</TD>	
	<TD align="left">Buah Busuk</TD>	
	<TD align="left">Buah Mentah di Peram</TD>	
	<TD align="left">Over Pruning</TD>	
	<TD align="left">Brondolan Tidak di Alas</TD>	
	<TD align="left">Brondolan Kotor/ Sampah</TD>	
	<TD align="left">Buah di Belah</TD>	
	<TD align="left">Buah Tidak di Antrikan/ di Tulis</TD>	
	<TD align="left">Buah Tidak Diletakan Di TPH</TD>	
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
		  <td align="left" valign="top"><%=NO%></td>
		  <%If val_IsTrue = "Y" Then%>
			<td valign="middle" width="30">
				<a href="HarvestingQuality.Asp?EMPCODE=<%=rsVH("EMPCODE")%>&FIELDCODE=<%=rsVH("FIELDCODE")%>&PeriodDate=<%=rsVH("FDDATE")%>&UNDER_RIPE=<%=rsVH("UNDER_RIPE")%>&OVER_RIPE=<%=rsVH("OVER_RIPE")%>&ABNORMAL=<%=rsVH("ABNORMAL")%>&LONG_STALK=<%=rsVH("LONG_STALK")%>&EATEN_BY_RAT=<%=rsVH("EATEN_BY_RAT")%>&UNHARVEST_FFB=<%=rsVH("UNHARVEST_FFB")%>&UNCOLLECT_LF_CIRCLE=<%=rsVH("UNCOLLECT_LF_CIRCLE")%>&UNCOLLECT_LF_PIECE=<%=rsVH("UNCOLLECT_LF_PIECE")%>&UNARRANGE_FFB=<%=rsVH("UNARRANGE_FFB")%>&UNPRUNE_FROND=<%=rsVH("UNPRUNE_FROND")%>&QE_1=<%=rsVH("QE_1")%>&QE_2=<%=rsVH("QE_2")%>&QE_3=<%=rsVH("QE_3")%>&QE_4=<%=rsVH("QE_4")%>&QE_5=<%=rsVH("QE_5")%>&QE_6=<%=rsVH("QE_6")%>&QE_7=<%=rsVH("QE_7")%>&QE_8=<%=rsVH("QE_8")%>&QE_9=<%=rsVH("QE_9")%>&QE_10=<%=rsVH("QE_10")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsVH(1)%>" border="0"></a>
			</td>
			<td>
				<a href="HarvestingQuality.Asp?EMPCODE=<%=rsVH("EMPCODE")%>&FIELDCODE=<%=rsVH("FIELDCODE")%>&PeriodDate=<%=rsVH("FDDATE")%>&txtJenis=Delete" onClick="return confirm('Are you sure to delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsVH(1)%>" border="0"></a>
			</td>
			<%else%>
				 <TD COLSPAN="2"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if%>	 
		    <%for i = 0 to nCountofField - 6%>
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
 
  <!-------------akhir---------------------->
 </div> </div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
