<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Shift Report</title>
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
 <%'-------------awal-----------------------
 %>
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
dim txt_SHIFTCODE,Perioddate
Dim PRODUCTIONDATE, SHIFT_SHIFTCODE, ACTUALSTART, ACTUALEND, CAGESPROCESSED, CAGEWEIGHT, RAMPBALANCE, CAGESBALANCE, PLATFORMBALANCE, DOWNTIME, CAUSE, TOTALPRODUCTIONHOUR
Dim cmdView
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""
txttable=request("txttable")

Perioddate=request("Perioddate")
txt_SHIFTCODE=request("txt_SHIFTCODE")
PRODUCTIONDATE=request("PRODUCTIONDATE")
SHIFT_SHIFTCODE=request("SHIFT_SHIFTCODE")
ACTUALSTART=request("ACTUALSTART")
ACTUALEND=request("ACTUALEND")
CAGESPROCESSED=request("CAGESPROCESSED")
CAGEWEIGHT=request("CAGEWEIGHT")
RAMPBALANCE=request("RAMPBALANCE")
CAGESBALANCE=request("CAGESBALANCE")
PLATFORMBALANCE=request("PLATFORMBALANCE")
DOWNTIME=request("DOWNTIME")
CAUSE=request("CAUSE")
TOTALPRODUCTIONHOUR = request("TOTALPRODUCTIONHOUR")

 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then

	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PRODUCTIONDATE)
		If Cek_IsDailyClosing = "FALSE" then
		   q="Delete from SHIFTREPORT where PRODUCTIONDATE=to_Date('" & PRODUCTIONDATE & "','MM/DD/YYYY') And ACTUALSTART='" & ACTUALSTART & "' And ACTUALEND='" & ACTUALEND & "' And fcBA='" & Session("fcBA") & "'"
		   'response.Write(q)
		   'response.Flush()
		   connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & PRODUCTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
       cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PRODUCTIONDATE)
		If Cek_IsDailyClosing = "FALSE" then
		  q="Insert into SHIFTREPORT (PRODUCTIONDATE, SHIFT_SHIFTCODE, ACTUALSTART, ACTUALEND, CAGESPROCESSED, CAGEWEIGHT, RAMPBALANCE, CAGESBALANCE, PLATFORMBALANCE, DOWNTIME, CAUSE,TOTALPRODUCTIONHOUR,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
			"values (to_Date('" & PRODUCTIONDATE & "','MM/DD/YYYY'),'" &  SHIFT_SHIFTCODE & "','" & ACTUALSTART & "','" & ACTUALEND & "'," &  CAGESPROCESSED & "," & _
			"" &  CAGEWEIGHT & "," &  RAMPBALANCE & "," &  CAGESBALANCE & "," &  PLATFORMBALANCE & ",'" & DOWNTIME & "','" &  CAUSE & "','" &  TOTALPRODUCTIONHOUR & "'," & _
			"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
			'response.Write(q)
			'Response.Flush()
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & PRODUCTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
        cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), PRODUCTIONDATE)
		If Cek_IsDailyClosing = "FALSE" then
		   q=" Update SHIFTREPORT set" & _
			" CAGESPROCESSED =" & CAGESPROCESSED & "" & _
			" ,RAMPBALANCE =" & RAMPBALANCE & "" & _
			" ,CAGESBALANCE =" & CAGESBALANCE & "" & _
			" ,PLATFORMBALANCE =" & PLATFORMBALANCE & " "& _
			" ,DOWNTIME ='" & DOWNTIME & "'" & _
			" ,CAUSE ='" & CAUSE & "'" & _
			" ,TOTALPRODUCTIONHOUR='" & TOTALPRODUCTIONHOUR & "'" & _
			" ,SHIFT_SHIFTCODE ='" & SHIFT_SHIFTCODE & "'" & _
			" ,CAGEWEIGHT =" & CAGEWEIGHT & "" & _
			" , LASTUPDATE = sysdate " & _
			" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			"where PRODUCTIONDATE=to_Date('" & PRODUCTIONDATE & "','MM/DD/YYYY') " & " And ACTUALSTART='" & ACTUALSTART & "' And ACTUALEND='" & ACTUALEND & "' And fcBA='" & Session("fcBA") & "'"
			connStr.execute(q)
			'response.Write(q)
			'response.Flush()
			connStr.execute("commit")
			'RESPONSE.Write "SQLnsddsdya : " & Q
		else
			err.number = 1001
			err.description = "Tanggal """ & PRODUCTIONDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
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

if len(trim(PRODUCTIONDATE)) <> 0 then
   PRODUCTIONDATE = Request("PRODUCTIONDATE")
else
   PRODUCTIONDATE = textDate(date)
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

 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Shift Report</div><% end if %>
 <form name="frmEntry" method="post" action="SHIFTREPORT.asp">
	<table class="data" width="100%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE SHIFT REPORT</strong></font></div></td>
          </TR>

		<tr><td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		
		</td></tr>
		<TR>
			<td align="left">Shift Code</td>
			<td align="left"> <select name="txt_SHIFTCODE">
			<%cmdCallComboSetup "SHIFT",REQUEST("txt_SHIFTCODE"),"fcba",session("fcba")%>
			</select></td>
		</tr>		
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintSHIFTREPORT.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_SHIFTCODE=<%=REQUEST("txt_SHIFTCODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		<!-- HADIR -->
		<%
		dim dataModule,dataSubmodule,dataDetail
		dataModule 		= "Mill Production"
		dataSubmodule 	= "Shift Report"
		dataDetail 		= "NULL"
		%>
		<a class="jqplot-image-button" style="" href="ajax/hadirAjax.asp?dataModule=<%=dataModule%>&dataSubmodule=<%=dataSubmodule%>&dataDetail=<%=dataDetail%>"><i class="fa fa-check-square-o"></i>&nbsp;Hadir</a>
		<!-- END:HADIR -->
		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   set rsGAD = server.CreateObject("Adodb.recordset")
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q="select PRODUCTIONDATE, SHIFT_SHIFTCODE, ACTUALSTART, ACTUALEND, CAGESPROCESSED, CAGEWEIGHT, " & _
      "RAMPBALANCE, CAGESBALANCE, PLATFORMBALANCE, DOWNTIME, CAUSE,TOTALPRODUCTIONHOUR, FCENTRY, FCEDIT, FCIP, FCBA from SHIFTREPORT where " & _
     " fcBA='" & session("fcBA") & "'"
	if  len(trim(Perioddate)) > 0 then
		q = q & " and to_char(PRODUCTIONDATE,'yyyymm') = to_char(to_date('" & Perioddate & "','MM/DD/YYYY'),'yyyymm')"
	end if
	if  len(trim(txt_SHIFTCODE)) > 0 then
		q = q & " and SHIFT_SHIFTCODE='" & txt_SHIFTCODE & "'"
	end if
	q = q  & " order by PRODUCTIONDATE, SHIFT_SHIFTCODE"
   'response.Write(q)
  'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr %>

   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="SHIFTREPORT.asp" method="post" name="eForm">
   <table class="data" width="100%" cellpadding="0" cellspacing="0">
		<td align="left">Date</td>
		<%IF txtJenis="Edit" Then%>
				<td align="left"><input type="text" maxlength="10" name="PRODUCTIONDATE" size="10" value="<%=PRODUCTIONDATE%>" readonly="True" style="background:#999999">
		<%Else%>
				<td align="left"><input class="datepicker" type="text" maxlength="10" name="PRODUCTIONDATE" size="10" value="<%=PRODUCTIONDATE%>">
		<%End if%>
				<font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<%IF txtJenis="Edit" Then%>
			<td align="left">Shift Code</td>
			<td align="left"> <select name="SHIFT_SHIFTCODE">
			<%cmdCallComboSetup "SHIFT",SHIFT_SHIFTCODE,"fcba",session("fcba")%>
			</select> <font color="#FF0000">*</font> </td>
		<%Else%>
			<td align="left">Shift Code</td>
			<td align="left"> <select name="SHIFT_SHIFTCODE">
			<%cmdCallComboSetup "SHIFT", SHIFT_SHIFTCODE,"fcba",session("fcba")%>
			</select> <font color="#FF0000">*</font> </td>
		<%End if%>
		</tr>
		<tr>
			<td align="left">Actual Start</td>
			<td align="left">
				<%IF txtJenis="Edit" Then%>
					<input type="text" maxlength="10" name="ACTUALSTART" size="10" value="<%=ACTUALSTART%>" readonly="True" style="background:#999999">
				<%Else%>
					<input type="text" maxlength="10" name="ACTUALSTART" size="10" value="<%=ACTUALSTART%>">
				<%End if%>
					<font color="#FF0000">*</font>
			</td>
		</tr>
		<tr>
			<td align="left">Actual End</td>
			<td align="left">
				<%IF txtJenis="Edit" Then%>
					<input type="text" maxlength="10" name="ACTUALEND" size="10" value="<%=ACTUALEND%>" readonly="True" style="background:#999999">
				<%Else%>
					<input type="text" maxlength="10" name="ACTUALEND" size="10" value="<%=ACTUALEND%>">
				<%End if%>
					<font color="#FF0000">*</font>
			</td>
		</TR>
		<TR>
		<td align="left">Cages Processed</td><td align="left"><input type="text" size="15" name="CAGESPROCESSED" value="<%=CAGESPROCESSED%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Cages Weight (Kg)</td><td align="left"><input type="text" size="15" name="CAGEWEIGHT" value="<%=CAGEWEIGHT%>"></td>
		</tr>

		<tr>
		<td align="left">Ramp Balance (Kg)</td><td align="left"><input type="text" size="15" name="RAMPBALANCE" value="<%=RAMPBALANCE%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Cages Balance (Kg)</td><td align="left"><input type="text" size="15" name="CAGESBALANCE" value="<%=CAGESBALANCE%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Plat Form Balance (Kg)</td><td align="left"><input type="text" size="15" name="PLATFORMBALANCE" value="<%=PLATFORMBALANCE%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Down Time</td><td align="left"><input type="text" size="15" name="DOWNTIME" value="<%=DOWNTIME%>"> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
		<td align="left">Cause</td><td align="left"><input type="text" size="50" maxlength="250" name="CAUSE" value="<%=CAUSE%>"></td>
		</tr>
		<tr>
		<td align="left">Processing Hours</td><td align="left"><input type="text" size="50" maxlength="250" name="TOTALPRODUCTIONHOUR" value="<%=TOTALPRODUCTIONHOUR%>"></td>
		</tr>
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_SHIFTCODE" value="<%=txt_SHIFTCODE%>">
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
   JumRec=1
   intRecordsPerPage = 20
	
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
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">SHIFT CODE</TD>	
	<TD align="left">ACTUAL START</TD>	
	<TD align="left">ACTUAL END</TD>	
	<TD align="left">CAGES PROCESSED</TD>	
	<TD align="left">CAGES WEIGHT</TD>	
	<TD align="left">RAMP BALANCES</TD>	
	<TD align="left">CAGE BALANCES</TD>	
	<TD align="left">PLAT FROM BALANCES</TD>	
	<TD align="left">DOWN TIME</TD>	
	<TD align="left">CAUSE</TD>	
	<TD align="left">PROCESSING HOURS</TD>	
	</tr>
	</thead>
	<tbody>
   
        <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count-5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td width="79" align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="21">
			
			<a href="SHIFTREPORT.asp?Perioddate=<%=rsGAD("PRODUCTIONDATE")%>&PRODUCTIONDATE=<%=rsGAD("PRODUCTIONDATE")%>&txttable=<%=request("txttable")%>&ACTUALSTART=<%=rsGAD("ACTUALSTART")%>&ACTUALEND=<%=rsGAD("ACTUALEND")%>&CAGESPROCESSED=<%=rsGAD("CAGESPROCESSED")%>&RAMPBALANCE=<%=rsGAD("RAMPBALANCE")%>&CAGESBALANCE=<%=rsGAD("CAGESBALANCE")%>&PLATFORMBALANCE=<%=rsGAD("PLATFORMBALANCE")%>&DOWNTIME=<%=rsGAD("DOWNTIME")%>&CAUSE=<%=rsGAD("CAUSE")%>&TOTALPRODUCTIONHOUR=<%=rsGAD("TOTALPRODUCTIONHOUR")%>&SHIFT_SHIFTCODE=<%=rsGAD("SHIFT_SHIFTCODE")%>&txt_SHIFTCODE=<%=rsGAD("SHIFT_SHIFTCODE")%>&CAGEWEIGHT=<%=rsGAD("CAGEWEIGHT")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="SHIFTREPORT.asp?Perioddate=<%=rsGAD("PRODUCTIONDATE")%>&PRODUCTIONDATE=<%=rsGAD("PRODUCTIONDATE")%>&SHIFT_SHIFTCODE=<%=rsGAD("SHIFT_SHIFTCODE")%>&txt_SHIFTCODE=<%=rsGAD("SHIFT_SHIFTCODE")%>&txttable=<%=request("txttable")%>&ACTUALSTART=<%=rsGAD("ACTUALSTART")%>&ACTUALEND=<%=rsGAD("ACTUALEND")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <%for i = 0 to nCountofField%>
			    <td width="597" align="left" valign="top" title="<%=rsGAD(i).Name%>"><%=rsGAD(i)%></td>
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
   </tbody>
   </TABLE>
   </div>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->

</body>
</html>
