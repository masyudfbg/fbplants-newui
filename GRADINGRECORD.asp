<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Grading Record</title>
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
 <%'-------------awal-----------------------
 %>
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim GRADINGNO, DATETIME, VEHICLENO, WEIGHT, BUNCHES, REMARKS, CHECKEDBY, WBT_CHITNO,F00,F5,TKOSONG,TPANJANG,DURA,BKECIL,AIR,SAMPAH,PASIR,LAIN
dim MENTAH,KMT,MT,TMT,BSK,AN
dim Perioddate,POSTINGDATE
Dim cmdView
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
txttable=request("txttable")

	Perioddate = request("Perioddate")

	GRADINGNO = request("GRADINGNO")
	DATETIME = request("DATETIME")
	VEHICLENO = request("VEHICLENO")
	WEIGHT = request("WEIGHT")
	BUNCHES = request("BUNCHES")
	REMARKS = request("REMARKS")
	CHECKEDBY = request("CHECKEDBY")
	WBT_CHITNO = request("WBT_CHITNO")
	F00 = request("F00")
	F5 = request("F5")
	TKOSONG = request("TKOSONG")
	TPANJANG = request("TPANJANG")
	DURA = request("DURA")
	BKECIL = request("BKECIL")
	AIR = request("AIR")
	SAMPAH = request("SAMPAH")
	PASIR = request("PASIR")
	LAIN = request("LAIN")
	POSTINGDATE = request("POSTINGDATE")
	MENTAH = request("mentah")
	KMT = request("kmt")
	MT = request("mt")
	TMT = request("tmt")
	BSK = request("bsk")
	AN = request("an")
	
	
	
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
    if F00 = "" then
	   F00=0
	end if
	if F5="" then
	   F5=0
	END IF
	IF TKOSONG="" THEN
	   TKOSONG=0
	END IF
	IF TPANJANG="" THEN
	   TPANJANG=0
	END IF
	IF DURA="" THEN
	   DURA=0
	END IF
	IF BKECIL = "" THEN
	   BKECIL=0
	END IF
	IF AIR="" THEN
	   AIR=0
	END IF
	IF SAMPAH="" THEN
	   SAMPAH=0
	END IF
	IF PASIR ="" THEN
	   PASIR =0
	END IF
	IF LAIN="" THEN
       LAIN=0
    END IF	   
	
	
	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), POSTINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q="Delete from GRADINGRECORD where GRADINGNO='" & GRADINGNO & "' And DATETIME =to_date('" & POSTINGDATE & "','mm/dd/yyyy') And WBT_CHITNO='" & WBT_CHITNO & "' And fcBA='" & Session("fcBA") & "'"
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & POSTINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), POSTINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q="Insert into GRADINGRECORD (GRADINGNO, DATETIME, VEHICLENO, WEIGHT, BUNCHES, REMARKS, CHECKEDBY, WBT_CHITNO, FCENTRY, FCEDIT, FCIP, FCBA,F00,F5,TKOSONG,TPANJANG,DURA,BKECIL,AIR,SAMPAH,PASIR,LAIN,POSTINGDATE, LASTUPDATE, LASTTIME,DATATYPE ) " & _
			"values ('" & GRADINGNO & "',to_date('" & Request("DATETIME") & "','mm/dd/yyyy'),'" &  VEHICLENO & "','" &  WEIGHT & "','" &  BUNCHES & "','" &  REMARKS & "','" &  CHECKEDBY & "','" &  WBT_CHITNO & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & _
			"'," & F00 & "," & F5 & "," & TKOSONG & "," & TPANJANG & "," & DURA & _ 
			" , " & BKECIL & ", " & AIR & ", " & SAMPAH & "," & PASIR & "," & LAIN &  _ 
			", to_Date('" & POSTINGDATE & "','MM/DD/YYYY'), sysdate, To_Char(sysdate,'HH24:MI'),'AM' )," & mentah & _ 
			"," & kmt & "," & mt & "," & tmt & "," & bsk & "," & an & "," & jk & _ 
			"," & tp & "," & bd & "," & sp & ""
			'response.Write(q)
			'response.Flush()
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & POSTINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), POSTINGDATE)
		If Cek_IsDailyClosing = "FALSE" then
			
			if MENTAH ="" THEN
				MENTAH =0
			END IF
			if KMT ="" THEN
				KMT =0
			END IF
			if MT ="" THEN
				MT =0
			END IF
			if TMT ="" THEN
				TMT =0
			END IF
			if BSK ="" THEN
				BSK =0
			END IF
			if AN ="" THEN
				AN =0
			END IF
			
			
		   q=" Update GRADINGRECORD set " &_
			" GRADINGNO ='" & GRADINGNO & "'" & _
			" ,DATETIME =to_date('" & Request("DATETIME") & "','mm/dd/yyyy') " & _
			" ,VEHICLENO='" & VEHICLENO & "'" & _
			" ,WEIGHT ='" & WEIGHT & "'" & _
			" ,BUNCHES ='" & BUNCHES & "'" & _
			" ,REMARKS ='" & REMARKS & "'" & _
			" ,CHECKEDBY ='" & CHECKEDBY & "'" & _
			" ,F00=" & F00 & "" & _
			" ,F5=" & F5 & "" & _
			" ,TKOSONG=" & TKOSONG & "" & _
			" ,TPANJANG=" & TPANJANG & "" & _
			" ,DURA=" & DURA & " " & _
			" ,BKECIL=" & BKECIL & "" & _
			" ,AIR=" & AIR & " "& _
			" ,SAMPAH=" & SAMPAH & "" & _
			" ,PASIR=" & PASIR & "" & _ 
			" ,LAIN=" & LAIN & "" & _
			" ,WBT_CHITNO ='" & WBT_CHITNO & "' " & _
			" , LASTUPDATE = sysdate " & _
			" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			" , mentah=" & MENTAH & " " & _ 
			" , kmt=" & KMT & " " &  _ 
			" , mt=" & MT & " " & _ 
			" , tmt=" & TMT & " " & _
			" , bsk=" & BSK & " " & _ 
			" , an=" & AN & " " & _
			" where  GRADINGNO='" & GRADINGNO & "'" & _
			" And DATETIME =to_date('" & Request("DATETIME") & "','mm/dd/yyyy') " & _
			" And POSTINGDATE=to_Date('" & POSTINGDATE & "','MM/DD/YYYY') " & _ 
			" And WBT_CHITNO='" & WBT_CHITNO & "' " & _
			" And fcBA='" & Session("fcBA") & "'"
		
			connStr.execute(q)
			connStr.execute("commit")
			'RESPONSE.Write "SQLnsddsdya : " & Q
			
			GRADINGNO =""
			DATETIME =""
			VEHICLENO=""
			WEIGHT =""
			BUNCHES =""
			REMARKS =""
			CHECKEDBY =""
			F00=""
			F5=""
			TKOSONG=""
			TPANJANG=""
			DURA=""
			BKECIL=""
			AIR=""
			SAMPAH=""
			PASIR=""
			LAIN=""
			WBT_CHITNO =""
			mentah=""
			kmt=""
			mt=""
			tmt=""
			bsk=""
			an=""
						
			 q=	" Update WEIGHBRIDGETICKET set " &_
				" JJG ='" & BUNCHES & "'" & _ 
				" , LASTUPDATE = sysdate " & _
				" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
				" where CHITNO='" & WBT_CHITNO & "' " & _
				" And fcBA='" & Session("fcBA") & "'"
				connStr.execute(q)
				connStr.execute("commit")
		else
			err.number = 1001
			err.description = "Tanggal """ & POSTINGDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
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
if len(trim(Perioddate)) <> 0 then
   Perioddate = Request("Perioddate")
else
   Perioddate = textDate(date)
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

 '-------------ISI Payment Form
 %>
 <form name="frmEntry" method="post" action="GRADINGRECORD.asp">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2" width="130"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE GRADING RECORD</strong></font></div></td>
          </TR>

		<tr><td align="left">Date</td>
		<td align="left" ><input class="datepicker" type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		
		</td></tr>
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintGRADINGRECORD.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		<!-- HADIR -->
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Weighbridge"
		dataSubmodule 	= "Grading Record"
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
	Dim rsGAD,intRecordsPerPage,cColor
	set rsGAD = server.CreateObject("Adodb.recordset")
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " select gradingno" & _
	    "        , datetime" & _
	    "        , vehicleno" & _
	    "        , weight" & _
	    "        , bunches" & _
	    "        , remarks" & _
	    "        , checkedby" & _
	    "        , wbt_chitno" & _
	    "        , f00" & _
	    "        , f5" & _
	    "        , dura" & _
	    "        , bkecil" & _
	    "        , air" & _
	    "        , sampah" & _
	    "        , pasir" & _
	    "        , lain" & _
		"        , mentah" & _
		"        , kmt" & _
		"        , mt" & _
		"        , tmt" & _
		"        , bsk" & _
		"        , an" & _
		"        , tkosong" & _
	    "        , tpanjang" & _
	    "        , postingdate" & _
	    "        , fcentry" & _
	    "        , fcedit" & _
	    "        , fcip" & _
	    "        , fcba" & _
		"        , DATATYPE" & _
	    " from gradingrecord" & _
	    " where fcBA = '" & session("fcBA") & "'"
	if  len(trim(Perioddate)) > 0 then
		q = q & " and to_char(postingdate,'yyyymmdd') = to_char(to_date('" & Perioddate & "', 'mm/dd/yyyy'),'yyyymmdd')"
	end if
	if Request("txtJenis")="Edit"  then
		q = q & " and to_char(postingdate,'yyyymmdd') = to_char(to_date('" & Perioddate & "', 'mm/dd/yyyy'),'yyyymmdd')"
	end if
	q = q 
  ' response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr
%>


   <form action="GRADINGRECORD.asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="90%" border="0">
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">Grading No</td><td align="left"><input type="text" size="15" maxlength="15" name="GRADINGNO" value="<%=GRADINGNO%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font> </td>
		<%else%>
			<td align="left">Grading No</td><td align="left"><input type="text" size="15" maxlength="15" name="GRADINGNO" value="<%=GRADINGNO%>"  ><font color="#FF0000">*</font> </td>
		<%end if%>
		</tr>
		<tr>
			<td align="left">Date</td>
			<td align="left">
				<%IF Request("txtJenis")="Edit" then%>
					<input type="text" maxlength="10" name="DATETIME" size="10" value="<%=DATETIME%>" readonly="True" style="background:#999999">
				<%Else%>
					<input class="datepicker" type="text" maxlength="10" name="DATETIME" size="10" value="<%=DATETIME%>">
				<%End if%>
				<font color="#FF0000">*</font> 
			</td>
		</tr>
		<tr>
			<td align="left">Posting Date</td>
			<td align="left">
				<%IF Request("txtJenis")="Edit" then%>
					<input type="text" maxlength="10" name="POSTINGDATE" size="10" value="<%=POSTINGDATE%>" readonly="True" style="background:#999999">
				<%Else%>
					<input class="datepicker" type="text" maxlength="10" name="POSTINGDATE" size="10" value="<%=POSTINGDATE%>">
				<%End if%>
				<font color="#FF0000">*</font> 
			</td>
		</tr>
		<TR>		
		</TR>
		<tr>
			<tr>
            <td align="left"><font face="Verdana" size="1">Vehicle</font></td><td align="left"><input type="text" size="15" maxlength="50" name="VEHICLENO" value="<%=VEHICLENO%>"><font color="#FF0000">*</font> </td>
			</tr>
		</TR>
		<TR>
			<td align="left">Weight (Kg)</td><td align="left"><input type="text" size="15" name="WEIGHT" value="<%=WEIGHT%>"><font color="#FF0000">*</font> </td>
		</tr>
		<tr>
		<td align="left">Bunches</td><td align="left"><input type="text" size="15" name="BUNCHES" readonly value="<%=BUNCHES%>"><font color="#FF0000">*</font> </td>
		</tr>
		<tr>
		<td align="left">Remarks</td><td align="left"><input type="text" size="50" maxlength="200" name="REMARKS" value="<%=REMARKS%>"></td>
		</tr>
		<tr>
		<td align="left">Checked By</td><td align="left"><input type="text" size="30" maxlength="50" name="CHECKEDBY" value="<%=CHECKEDBY%>"></td>
		</tr>
		<tr>
		<%IF Request("txtJenis")="Edit" then%>
			<td align="left">WBT Chitno</td><td align="left"><input type="text" size="15" maxlength="15" name="WBT_CHITNO" value="<%=WBT_CHITNO%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font> </td>
		<%Else%>			
			<td align="left">WBT Chitno</td><td align="left"><input type="text" size="15" maxlength="15" name="WBT_CHITNO" value="<%=WBT_CHITNO%>"><font color="#FF0000">*</font> </td>		
		<%End if%>
		</tr>
		<tr>
		<td align="left">F00</td><td align="left" style="width:20%;margin-left:0.3%"><input type="text" size="10" maxlength="10" name="F00" value="<%=F00%>"></td>
		<td align="left">MENTAH</td><td align="left"><input type="text" size="10" maxlength="10" name="mentah" value="<%=mentah%>"></td>
		</tr>
		
		<tr>
		<td align="left">F5</td><td align="left"><input type="text" size="10" maxlength="10" name="F5" value="<%=F5%>"></td>
		<td align="left">KURANG MATANG</td><td align="left"><input type="text" size="10" maxlength="10" name="kmt" value="<%=kmt%>"></td>
		</tr>
		<tr>
		<td align="left">TKOSONG</td><td align="left"><input type="text" size="10" maxlength="10" name="TKOSONG" value="<%=TKOSONG%>"></td>
		<td align="left">MATANG</td><td align="left"><input type="text" size="10" maxlength="10" name="mt" value="<%=mt%>"></td>
		</tr>
		<tr>
		<td align="left">TPANJANG</td><td align="left"><input type="text" size="10" maxlength="10" name="TPANJANG" value="<%=TPANJANG%>"></td>
		<td align="left">TERLALU MATANG</td><td align="left"><input type="text" size="10" maxlength="10" name="tmt" value="<%=tmt%>"></td>
		</tr>
		<tr>
		<td align="left">DURA</td><td align="left"><input type="text" size="10" maxlength="10" name="DURA" value="<%=DURA%>"></td>
		<td align="left">BUAH BUSUK</td><td align="left"><input type="text" size="10" maxlength="10" name="bsk" value="<%=bsk%>"></td>
		</tr>
		<tr>
		<td align="left">BKECIL</td><td align="left"><input type="text" size="10" maxlength="10" name="BKECIL" value="<%=BKECIL%>"></td>
		<td align="left">BUAH ABNORMAL</td><td align="left"><input type="text" size="10" maxlength="10" name="an" value="<%=an%>"></td>
		</tr>
		
		
	   <tr>
		<td align="left">AIR</td><td align="left"><input type="text" size="10" maxlength="10" name="AIR" value="<%=AIR%>"></td>
		
		</tr>
		<tr>
		<td align="left">SAMPAH</td><td align="left"><input type="text" size="10" maxlength="10" name="SAMPAH" value="<%=SAMPAH%>"></td>
		
		</tr>
		<tr>
		<td align="left">PASIR</td><td align="left"><input type="text" size="10" maxlength="10" name="PASIR" value="<%=PASIR%>"></td>
		
		</tr>
		<tr>
		<td align="left">LAIN</td><td align="left"><input type="text" size="10" maxlength="10" name="LAIN" value="<%=LAIN%>"></td>
		
		</tr>
		
		
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
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
		  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
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
   <table class="data" >
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;">
	<TD align="left">NO</TD>
	<% if rsGAD("DATATYPE") <> "AT"  then %>
		<TD colspan="2">&nbsp;</TD>
	<% ELSE %>	
		<TD colspan="2">&nbsp;</TD>
	<% end if %>
	<TD align="left">GRADING NO</TD>	
	<TD align="left">DATE</TD>	
	<TD align="left">VEHICLE</TD>	
	<TD align="left">WEIGHT</TD>	
	<TD align="left">BUNCHES</TD>	
	<TD align="left">REMARKS</TD>	
	<TD align="left">CHECKED BY</TD>	
	<TD align="left">CHITNO</TD>	
	
	<TD align="right">F00</TD>
	<TD align="right">F5</TD>
	<TD align="right">DURA</TD>
	<TD align="right">BKECIL</TD>
	<TD align="right">AIR</TD>
	<TD align="right">SAMPAH</TD>
	<TD align="right">PASIR</TD>
	<TD align="right">LAIN</TD>
	<TD align="right">MENTAH</TD>
	<TD align="right">KMT</TD>
	<TD align="right">MT</TD>
	<TD align="right">TMT</TD>
	<TD align="right">BSK</TD>
	<TD align="right">AN</TD>
	<TD align="right">JJKOSONG</TD>
	<TD align="right">TPANJANG</TD>
	<TD align="right">Posting Date</TD>
	</tr>
   
   <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-6
	  'RESPONSE.WRITE nCountofField
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="53" align="left" valign="top"><%=NO%></td>
		<% if rsGAD("DATATYPE") = "AT"  then %>
			<td colspan="2" align="center">Automate</td>  
		<% ELSE %>	
			<td valign="middle" width="23">
			<a href="GRADINGRECORD.asp?GRADINGNO=<%=rsGAD("GRADINGNO")%>&txttable=<%=request("txttable")%>&DATETIME=<%=rsGAD("DATETIME")%>&POSTINGDATE=<%=rsGAD("POSTINGDATE")%>&VEHICLENO=<%=rsGAD("VEHICLENO")%>&WEIGHT=<%=rsGAD("WEIGHT")%>&BUNCHES=<%=rsGAD("BUNCHES")%>&REMARKS=<%=rsGAD("REMARKS")%>&CHECKEDBY=<%=rsGAD("CHECKEDBY")%>&WBT_CHITNO=<%=rsGAD("WBT_CHITNO")%>&F00=<%=rsGAD("F00")%>&F5=<%=rsGAD("F5")%>&DURA=<%=rsGAD("DURA")%>&BKECIL=<%=rsGAD("BKECIL")%>&AIR=<%=rsGAD("AIR")%>&SAMPAH=<%=rsGAD("SAMPAH")%>&PASIR=<%=rsGAD("PASIR")%>&LAIN=<%=rsGAD("LAIN")%>&MENTAH=<%=rsGAD("MENTAH")%>&KMT=<%=rsGAD("KMT")%>&MT=<%=rsGAD("MT")%>&TMT=<%=rsGAD("TMT")%>&BSK=<%=rsGAD("BSK")%>&AN=<%=rsGAD("AN")%>&TKOSONG=<%=rsGAD("TKOSONG")%>&TPANJANG=<%=rsGAD("TPANJANG")%>&Perioddate=<%=rsGAD("postingdate")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="48">
			<a href="GRADINGRECORD.asp?GRADINGNO=<%=rsGAD("GRADINGNO")%>&txttable=<%=request("txttable")%>&POSTINGDATE=<%=rsGAD("postingdate")%>&Perioddate=<%=rsGAD("postingdate")%>&WBT_CHITNO=<%=rsGAD("WBT_CHITNO")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0"></a>
			</td>

		<% end if %>
		    <%for i = 0 to nCountofField%>
			    <td width="210" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsGAD.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   end if 
   rsGAD.Close
   set rsGAD = Nothing%>
   </div>
   </TABLE>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
