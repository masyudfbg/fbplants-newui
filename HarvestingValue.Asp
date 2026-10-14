<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Nilai Panen</title>
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
 <!-------------awal------------->
 <%

Dim HARVESTINGVALUECODE, CUSTCODE, FDDATE, header_reference, header_remarks
Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim LINENO, DIVISIONINFO, FIELDCODEINFO, VEHICLECODEINFO, SAMPLE_QUANTITY, SAMPLE_WEIGHT
Dim UNDER_RIPE, FINE_RIPE,OVER_RIPE, LONG_STALK, EMPTY_BUNCH, ABNORMAL_RIPE, EATEN_BY_RAT, SPOILED_RIPE, DETACHED_RIPE, RUBBISH
Dim QE_1, QE_2, QE_3, QE_4, QE_5, QE_6, QE_7, QE_8, QE_9, QE_10
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
DIM strLineMax, rsstrLineMax, tempMaxLineNo
Dim cmdView
Dim rsHeaderHV
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
	'---- Data Header
	HARVESTINGVALUECODE	=	Request("HARVESTINGVALUECODE")
	FDDATE = Request("FDDATE")
	CUSTCODE = Request("CUSTCODE")
	header_reference = Request("header_reference")
	header_remarks = Request("header_remarks")


	if Len(Trim(FDDATE)) > 0 then
		FDDATE	=	Request("FDDATE")
	else
		FDDATE	=	textDate(date)
	end if

 '---- Data Details
LINENO			=	Request("LINENO")
DIVISIONINFO	=	Request("DIVISIONINFO")
FIELDCODEINFO	=	Request("FIELDCODEINFO")
VEHICLECODEINFO	=	Request("VEHICLECODEINFO")
SAMPLE_QUANTITY	= To_Nol	(Request("SAMPLE_QUANTITY"))
SAMPLE_WEIGHT	= To_Nol	(Request("SAMPLE_WEIGHT"))
UNDER_RIPE		= To_Nol	(Request("UNDER_RIPE"))
FINE_RIPE		= To_Nol	(Request("FINE_RIPE"))
OVER_RIPE		= To_Nol	(Request("OVER_RIPE"))
LONG_STALK		= To_Nol	(Request("LONG_STALK"))
EMPTY_BUNCH		= To_Nol	(Request("EMPTY_BUNCH"))
ABNORMAL_RIPE	= To_Nol	(Request("ABNORMAL_RIPE"))
EATEN_BY_RAT	= To_Nol	(Request("EATEN_BY_RAT"))
SPOILED_RIPE	= To_Nol	(Request("SPOILED_RIPE"))
DETACHED_RIPE	= To_Nol	(Request("DETACHED_RIPE"))
RUBBISH	  		= To_Nol	(Request("RUBBISH"))
ABNORMAL_RIPE	= To_Nol	(Request("ABNORMAL_RIPE"))
ABNORMAL_RIPE	= To_Nol	(Request("ABNORMAL_RIPE"))
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

if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" then
	if Request("txtJenis")= "Header" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), FDDATE)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			q	=	" Delete from harvestingvaluedetails" & _
					" where HARVESTINGVALUECODE = '" & HARVESTINGVALUECODE & "'" & _
					" and fcBA = '" & Session("fcBA") & "'"
			connStr.execute(q)
			
			if err.number = 0 then
				q	=	" Delete from harvestingvalue" & _
						" where HARVESTINGVALUECODE = '" & HARVESTINGVALUECODE & "'" & _
						" and fcBA = '" & Session("fcBA") & "'"
				connStr.execute(q)
				
				if err.number  = 0 then
					connStr.CommitTrans
				end if
			end if
			'response.Write(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & FDDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
		
    elseif Request("txtJenis")= "Details" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), FDDATE)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			q	=	" Delete from harvestingvaluedetails" & _
					" where  HARVESTINGVALUECODE = '" & HARVESTINGVALUECODE & "'" & _
					" and LINENO = '" & LINENO & "'" & _
					" and fcBA = '" & Session("fcBA") & "'"
			connStr.execute(q)
			
			if err.number  = 0 then
				connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & FDDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
		
	elseif Request("txtJenis")="New"  then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), FDDATE)
		If Cek_IsDailyClosing = "FALSE" then
			Dim rsPO
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			connstr.BeginTrans
			rsPO.Open " Select HARVESTINGVALUECODE from harvestingvalue where HARVESTINGVALUECODE='" & HARVESTINGVALUECODE & "' and fcba = '" & Session("fcBA") & "'", connstr
			IF rsPO.eof then
				q	=	" Insert into harvestingvalue" & _
						" ( HARVESTINGVALUECODE,CUSTCODE, FDDATE, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME )" & _
						" values ('" & HARVESTINGVALUECODE & "','" & CUSTCODE & "', to_date('" & FDDATE & "','mm/dd/yyyy'), '" & header_reference & "'" & _
						" , '" & header_remarks  & "', '" & Session("CodeName") & "', '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") &"', '" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
				'response.Write(q)
				connStr.execute(q)
			End if
			rsPO.Close
			set rsPO = Nothing	
			if err.number = 0 then
				strLineMax	=	" select max(to_number(nvl(LINENO,0))) maxLineIdKey from harvestingvaluedetails where " & _
								" HARVESTINGVALUECODE = '" & HARVESTINGVALUECODE & "' " & _ 
								" and fcBA='" & session("fcBA") & "'"

		  
				Set rsstrLineMax = Server.CreateObject("ADODB.Recordset")
				rsstrLineMax.Open strLineMax, connstr, adLockOptimistic, adOpenStatic
			
				if not rsstrLineMax.eof then
					if len(trim(rsstrLineMax("maxLineIdKey"))) <> 0 then 
						tempMaxLineNo = cdbl(rsstrLineMax("maxLineIdKey")) + 1
					else
						tempMaxLineNo = 1
					end if
				end if
						
				rsstrLineMax.Close
				set rsstrLineMax = Nothing 

				q	=	" Insert into harvestingvaluedetails (" & _
						" HARVESTINGVALUECODE" & _
						" ,LINENO" & _
						" ,DIVISIONINFO" & _
						" , FIELDCODEINFO" & _
						" , VEHICLECODEINFO" & _
						" , SAMPLE_QUANTITY" & _
						" , SAMPLE_WEIGHT" & _
						" , UNDER_RIPE" & _	
						" , FINE_RIPE" & _	
						" , OVER_RIPE" & _	
						" , LONG_STALK" & _	
						" , EMPTY_BUNCH" & _	
						" , ABNORMAL_RIPE" & _	
						" , EATEN_BY_RAT" & _	
						" , SPOILED_RIPE" & _	
						" , DETACHED_RIPE" & _	
						" , RUBBISH" & _	
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
						" '" & HARVESTINGVALUECODE & "'" & _
						" , '" & tempMaxLineNo & "'" & _
						" , '" & DIVISIONINFO & "' " & _
						" , '" & FIELDCODEINFO & "' " & _
						" , '" & VEHICLECODEINFO & "' " & _
						" , '" & SAMPLE_QUANTITY & "' " & _
						" , '" & SAMPLE_WEIGHT & "' " & _
						" , '" & UNDER_RIPE & "' " & _
						" , '" & FINE_RIPE & "' " & _
						" , '" & OVER_RIPE & "' " & _
						" , '" & LONG_STALK & "' " & _
						" , '" & EMPTY_BUNCH & "' " & _
						" , '" & ABNORMAL_RIPE & "' " & _
						" , '" & EATEN_BY_RAT & "'" & _
						" , '" & SPOILED_RIPE & "'" & _
						" , '" & DETACHED_RIPE & "'" & _
						" , '" & RUBBISH & "'" & _
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
				if err.number  = 0 then
					connStr.CommitTrans
				end if
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & FDDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
		
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), FDDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q	=	" Update harvestingvaluedetails set" & _
					" DIVISIONINFO = '" & DIVISIONINFO & "'" & _
					" , FIELDCODEINFO = '" & FIELDCODEINFO & "'" & _
					" , VEHICLECODEINFO = '" & VEHICLECODEINFO & "'" & _
					" , SAMPLE_QUANTITY = '" & SAMPLE_QUANTITY & "'" & _
					" , SAMPLE_WEIGHT = '" & SAMPLE_WEIGHT & "'" & _
					" , UNDER_RIPE = '" & UNDER_RIPE & "'" & _
					" , FINE_RIPE = '" & FINE_RIPE & "'" & _
					" , OVER_RIPE = '" & OVER_RIPE & "'" & _
					" , LONG_STALK = '" & LONG_STALK & "'" & _
					" , EMPTY_BUNCH = '" & EMPTY_BUNCH & "'" & _
					" , ABNORMAL_RIPE = '" & ABNORMAL_RIPE & "'" & _
					" , EATEN_BY_RAT = '" & EATEN_BY_RAT & "'" & _
					" , SPOILED_RIPE = '" & SPOILED_RIPE & "'" & _
					" , DETACHED_RIPE = '" & DETACHED_RIPE & "'" & _
					" , RUBBISH = '" & RUBBISH & "'" & _
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
					" HARVESTINGVALUECODE = '" & HARVESTINGVALUECODE & "'" & _
					" And LINENO = '" & LINENO & "'" & _
					" and fcBA = '" & Session("fcBA") & "'"
					
					'response.write q
					connStr.execute(q)
					connStr.execute("commit")
		else
			err.number = 1001
			err.description = "Tanggal """ & FDDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if 

Dim val_Tanggal, val_head_ref, val_head_rem
set rsHeaderHV = server.CreateObject("Adodb.recordset")
	q = " select to_char(fddate, 'mm/dd/yyyy') fddate" & _
	    "        , custcode" & _
		"        , reference" & _
	    "        , remarks" & _
	    " from harvestingvalue  " & _
        " where harvestingvaluecode = '" & HARVESTINGVALUECODE & "'" & _
        "       and fcBA = '" & session("fcBA") & "'"

	rsHeaderHV.Open q, connstr
	'response.Write q
	
if Not rsHeaderHV.Eof then
	FDDATE = rsHeaderHV("FDDATE")
	CUSTCODE = rsHeaderHV("custcode")
	header_reference = rsHeaderHV("REFERENCE")
	header_remarks = rsHeaderHV("REMARKS")
End if

rsHeaderHV.Close
set rsHeaderHV = Nothing


if err.number <> 0 then
    response.Write  "<div class = ""error"">" & err.description & "</div>"
End if
%>

<!-------------ISI Harvesting Write Off Form------------->
 <form name="frmEntry" method="post" action="HarvestingValue.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Harvesting Value Code</td>
		<td align="left"><input type="text" maxlength="50" name="HARVESTINGVALUECODE" size="30" value="<%=HARVESTINGVALUECODE %>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="FDDATE" size="10" value="<%=FDDATE%>">
		<!--<a href="javascript:show_calendar('frmEntry.writeoffdate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
		</td>
		</tr>
		
       <tr>
			
			<td align="left">Customer Code</td>
			<td align="left"> <select name="CUSTCODE">
			<%cmdCallComboSetup "CUSTOMER",CUSTCODE,"fcba",session("fcba")%>
			</select> <font color="#FF0000">*</font></td>
		
		</tr>
        
        
		<tr>
		<td align="left">Reference</td>
		<td align="left"><input type="text" maxlength="100" name="header_reference" size="30" value="<%=header_reference%>"></td>		
		</tr>

		
		<tr>
		<td align="left">Remarks</td>
		<td align="left"><textarea maxlength="1000" name="header_remarks"  cols="50" rows="3"><%=header_remarks%></textarea></td>		
		</tr>	

		</td>
		</tr>

	

		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
	    <a href="javascript:launchwin('PrintHarvestingValue.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&HARVESTINGVALUECODE=<%=HARVESTINGVALUECODE%>&FDDATE=<%=FDDATE%>&header_reference=<%=header_reference%>&header_remarks=<%=header_remarks%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<a href="HarvestingValue.Asp?HARVESTINGVALUECODE=<%=HARVESTINGVALUECODE%>&FDDATE=<%=FDDATE%>&txtJenis=Header" onClick="return confirm('Are you sure to delete');">
			<img src="IMAGES/delete.gif" alt="Delete Data : <%=HARVESTINGVALUECODE%>" border="0">
		</a>

		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(HARVESTINGVALUECODE)) > 4 Then
   Dim rsharvestingvalue,intRecordsPerPage,cColor
   set rsharvestingvalue = server.CreateObject("Adodb.recordset")
   q= " select" & _
			" HARVESTINGVALUECODE" & _
			" , LINENO" & _
			" , DIVISIONINFO" & _
			" , FIELDCODEINFO" & _
			" , VEHICLECODEINFO" & _
			" , SAMPLE_QUANTITY" & _
			" , SAMPLE_WEIGHT" & _
			" , UNDER_RIPE" & _
			" , FINE_RIPE" & _
			" , OVER_RIPE" & _
			" , LONG_STALK" & _
			" , EMPTY_BUNCH" & _
			" , ABNORMAL_RIPE" & _
			" , EATEN_BY_RAT" & _
			" , SPOILED_RIPE" & _
			" , DETACHED_RIPE" & _
			" , RUBBISH" & _
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
     " from harvestingvaluedetails  " & _
     " where HARVESTINGVALUECODE='" & HARVESTINGVALUECODE & "' and " & _
     " fcBA='" & session("fcBA") & "'"
	 
   'response.write q
   Set rsharvestingvalue = Server.CreateObject("ADODB.Recordset")
   rsharvestingvalue.CursorType = 3
   rsharvestingvalue.CursorLocation = 3
   rsharvestingvalue.Open q, connstr
   %>
   <form action="HarvestingValue.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
        
		
		<tr>
		 <td align="left">Division</td><td align="left"><input type="text" size="30" name="DIVISIONINFO" value="<%=DIVISIONINFO%>"></td>
         <td align="left">Buah Mentah Dibelah</td><td align="left"><input type="text" size="15" name="QE_1" value="<%=QE_1%>"></td>
		  </tr>
		
		<tr>
		
		<td align="left">Field Code</td><td align="left"><input type="text" size="30" name="FIELDCODEINFO" value="<%=FIELDCODEINFO%>"></td>
         <td align="left">Selisih Tandan</td><td align="left"><input type="text" size="15" name="QE_2" value="<%=QE_2%>"></td>
		
		</tr>
        
		<tr>
		<td align="left">Vehicle Code</td>
        <td align="left"><input type="text" maxlength="100" size="30" name="VEHICLECODEINFO" value="<%=VEHICLECODEINFO%>"></td>
        <td align="left">QE_3</td><td align="left"><input type="text" size="15" name="QE_3" value="<%=QE_3%>"></td>
		</tr>
        
		<tr>
			<td align="left">Bunch</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="SAMPLE_QUANTITY" value="<%=SAMPLE_QUANTITY%>"></td>
            <td align="left">QE_4</td><td align="left"><input type="text" size="15" name="QE_4" value="<%=QE_4%>"></td>
		</tr>
        
        <tr>
			<td align="left">Weight</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="SAMPLE_WEIGHT" value="<%=SAMPLE_WEIGHT%>"></td>
            <td align="left">QE_5</td><td align="left"><input type="text" size="15" name="QE_5" value="<%=QE_5%>"></td>
		</tr>
        
        <tr>
			<td align="left">Buah Mentah</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="UNDER_RIPE" value="<%=UNDER_RIPE%>"></td>
            <td align="left">QE_6</td><td align="left"><input type="text" size="15" name="QE_6" value="<%=QE_6%>"></td>
            
		</tr>
        
        <tr>
			<td align="left">Buah Masak</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="FINE_RIPE" value="<%=FINE_RIPE%>"></td>
            <td align="left">QE_7</td><td align="left"><input type="text" size="15" name="QE_7" value="<%=QE_7%>"></td>
           
		</tr>
        
        <tr>
			<td align="left">Buah Terlalu Masak</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="OVER_RIPE" value="<%=OVER_RIPE%>"></td>
            <td align="left">QE_8</td><td align="left"><input type="text" size="15" name="QE_8" value="<%=QE_8%>"></td>
            
		</tr>
        
        <tr>
			<td align="left">Tangkai Panjang</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="LONG_STALK" value="<%=LONG_STALK%>"></td>
            <td align="left">QE_9</td><td align="left"><input type="text" size="15" name="QE_9" value="<%=QE_9%>"></td>
		</tr>
        
        <tr>
			<td align="left">Buah Kosong</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="EMPTY_BUNCH" value="<%=EMPTY_BUNCH%>"></td>
            <td align="left">QE_10</td><td align="left"><input type="text" size="15" name="QE_10" value="<%=QE_10%>"></td>
		</tr>
        
        <tr>
			<td align="left">Buah Abnormal</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="ABNORMAL_RIPE" value="<%=ABNORMAL_RIPE%>"></td>
		</tr>
        
        <tr>
			<td align="left">Buah Dimakan Tikus</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="EATEN_BY_RAT" value="<%=EATEN_BY_RAT%>"></td>
		</tr>
        
        <tr>
			<td align="left">Buah Busuk</td>
             <td align="left"><input type="text" maxlength="100" size="15" name="SPOILED_RIPE" value="<%=SPOILED_RIPE%>"></td>
		</tr>
        
        <tr>
			<td align="left">Buah Bondolan</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="DETACHED_RIPE" value="<%=DETACHED_RIPE%>"></td>
		</tr>
        
        <tr>
			<td align="left">Sampah</td>
            <td align="left"><input type="text" maxlength="100" size="15" name="RUBBISH" value="<%=RUBBISH%>"></td>
		</tr>
        
		<tr>
			<td align="left" style="display:none;">LineNo</td>
			<td align="left"><input type="hidden" size="15" name="LINENO" value="<%=LINENO%>"></td>
		</tr>
		
		
	   <TR><td colspan="2" align="left">
		  <input type="hidden" name="HARVESTINGVALUECODE" value="<%=HARVESTINGVALUECODE%>">
          <input type="hidden" name="CUSTCODE" value="<%=CUSTCODE%>">
		  <input type="hidden" name="FDDATE" value="<%=FDDATE%>">
		  <input type="hidden" name="header_reference" value="<%=header_reference%>">
		  <input type="hidden" name="header_remarks" value="<%=header_remarks%>">
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		  <input type="hidden" name="txtJenis" value="Edit">
		  <input Type="Submit" style="background-color: #298CF7; color: #F7C610; font-weight: 
		       bold; font-size: 8pt; font-family:Verdana" name="Add" value="Update">
		  <%Else%>
		  <input type="hidden" name="txtJenis" value="New">
		  <input Type="Submit"  name="Add" value="Add">
		  <%End if%>
		  <input Type="Reset"  name="Reset" value="Reset">
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
		  </tr>
   </table>

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">Harvesting Code</TD>	
    <TD align="left">Line No</TD>
	<TD align="left">Division</TD>	
    <TD align="left">Field Code</TD>
    <TD align="left">Vehicle Code</TD>
    <TD align="left">Bunch</TD>
    <TD align="left">Weight</TD>	
	<TD align="left">Buah Mentah</TD>	
    <TD align="left">Buah Masak</TD>
	<TD align="left">Buah Terlalu Masak</TD>
	<TD align="left">Tangkai Panjang</TD>
    <TD align="left">Buah Kosong</TD>
    <TD align="left">Buah Abnormal</TD>	
	<TD align="left">Buah Dimakan Tikus</TD>	
	<TD align="left">Buah Busuk</TD>	
	<TD align="left">Buah Brondolan</TD>	
	<TD align="left">Sampah</TD>		
	<TD align="left">QE 1</TD>	
	<TD align="left">QE 2</TD>	
	<TD align="left">QE 3</TD>	
	<TD align="left">QE 4</TD>	
	<TD align="left">QE 5</TD>	
	<TD align="left">QE 6</TD>	
	<TD align="left">QE 7</TD>
    <TD align="left">QE 8</TD>
    <TD align="left">QE 9</TD>
    <TD align="left">QE 10</TD>
	</tr>
   <%
   if Not rsharvestingvalue.eof then
      nCountofField=rsharvestingvalue.fields.count-1
	  no=1
	  While Not rsharvestingvalue.EOF

		if no  mod 2 = 1 then
			cColor = "#EEEEEE"
		else
			cColor = "#E5E5E5"
		end if%>
		<tr bgcolor="<%=cColor%>">
		<td align="left" valign="top"><%=NO%></td>
		<td valign="middle" width="10">
			
	
			
			
			<%Dim rsCekClose,qCekClose
			Set rsCekClose = Server.CreateObject("ADODB.Recordset")
            rsCekClose.CursorType = 3
            rsCekClose.CursorLocation = 3
			qCekClose="Select HARVESTINGVALUECODE from harvestingvalue where HARVESTINGVALUECODE='" & rsharvestingvalue("HARVESTINGVALUECODE") & "'" & "  and fcba = '" & session("fcba") & "'"

			rsCekClose.Open qCekClose, connstr
			if Not rsCekClose.Eof then%>
				<a href="HarvestingValue.Asp?HARVESTINGVALUECODE=<%=rsharvestingvalue("HARVESTINGVALUECODE")%>&LINENO=<%=rsharvestingvalue("LINENO")%>&DIVISIONINFO=<%=rsharvestingvalue("DIVISIONINFO")%>&FIELDCODEINFO=<%=rsharvestingvalue("FIELDCODEINFO")%>&VEHICLECODEINFO=<%=rsharvestingvalue("VEHICLECODEINFO")%>&SAMPLE_QUANTITY=<%=rsharvestingvalue("SAMPLE_QUANTITY")%>&SAMPLE_WEIGHT=<%=rsharvestingvalue("SAMPLE_WEIGHT")%>&UNDER_RIPE=<%=rsharvestingvalue("UNDER_RIPE")%>&FINE_RIPE=<%=rsharvestingvalue("FINE_RIPE")%>&OVER_RIPE=<%=rsharvestingvalue("OVER_RIPE")%>&LONG_STALK=<%=rsharvestingvalue("LONG_STALK")%>&EMPTY_BUNCH=<%=rsharvestingvalue("EMPTY_BUNCH")%>&ABNORMAL_RIPE=<%=rsharvestingvalue("ABNORMAL_RIPE")%>&EATEN_BY_RAT=<%=rsharvestingvalue("EATEN_BY_RAT")%>&SPOILED_RIPE=<%=rsharvestingvalue("SPOILED_RIPE")%>&DETACHED_RIPE=<%=rsharvestingvalue("DETACHED_RIPE")%>&RUBBISH=<%=rsharvestingvalue("RUBBISH")%>&QE_1=<%=rsharvestingvalue("QE_1")%>&QE_2=<%=rsharvestingvalue("QE_2")%>&QE_3=<%=rsharvestingvalue("QE_3")%>&QE_4=<%=rsharvestingvalue("QE_4")%>&QE_5=<%=rsharvestingvalue("QE_5")%>&QE_6=<%=rsharvestingvalue("QE_6")%>&QE_7=<%=rsharvestingvalue("QE_7")%>&QE_8=<%=rsharvestingvalue("QE_8")%>&QE_9=<%=rsharvestingvalue("QE_9")%>&QE_10=<%=rsharvestingvalue("QE_10")%>&CUSTCODE=<%=CUSTCODE%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsharvestingvalue(1)%>" border="0"></a></td>
				<td valign="middle" width="10"><a href="HarvestingValue.Asp?HARVESTINGVALUECODE=<%=rsharvestingvalue("HARVESTINGVALUECODE")%>&FDDATE=<%=FDDATE%>&CUSTCODE=<%=CUSTCODE%>&LINENO=<%=rsharvestingvalue("LINENO")%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
				<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsharvestingvalue(1)%>" border="0"></a></td>
			<%Else%>	 
			    <td colspan="1"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if
			rsCekClose.close
			set rsCekClose = Nothing%>	 
		    <%for i = 0 to nCountofField%>
			    <td align="left" title="<%=rsharvestingvalue(i).Name%>" valign="top"><%=rsharvestingvalue(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsharvestingvalue.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsharvestingvalue.Close
   set rsharvestingvalue = Nothing%>
   
   </TABLE>
   </div>
   </form>
<%End if%>  
 
  <!-------------akhir---------------------->
</div> </div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
