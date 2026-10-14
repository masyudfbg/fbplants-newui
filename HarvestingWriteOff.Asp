<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Write Off Restan</title>
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

Dim writeoff_code, writeoffdate, header_reference, header_remarks
Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim FIELDCODE, QUANTITY, DETAIL_REFERENCE, DETAIL_REMARKS, LINENO
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
DIM strLineMax, rsstrLineMax, tempMaxLineNo
Dim cmdView
 Dim rsHeaderHWO
 Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 
	'---- Data Header
	writeoff_code	=	Request("writeoff_code")
	writeoffdate	=	Request("writeoffdate")
	header_reference=	Request("header_reference")
	header_remarks	=	Request("header_remarks")


	if Len(Trim(writeoffdate)) > 0 then
		writeoffdate	=	Request("writeoffdate")
	else
		writeoffdate	=	textDate(date)
	end if

 '---- Data Details
FIELDCODE		=	Request("FIELDCODE")
QUANTITY		=	To_Nol(Request("QUANTITY"))
DETAIL_REFERENCE=	Request("DETAIL_REFERENCE")
DETAIL_REMARKS	=	Request("DETAIL_REMARKS")
LINENO			=	Request("LINENO")


if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" then
	if Request("txtJenis")= "Header" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), writeoffdate)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			q	=	" Delete from harvestingwriteoffdetail" & _
					" where WRITEOFFCODE = '" & writeoff_code & "'" & _
					" and fcBA = '" & Session("fcBA") & "'"
			connStr.execute(q)
			
			if err.number = 0 then
				q	=	" Delete from harvestingwriteoff" & _
						" where WRITEOFFCODE = '" & writeoff_code & "'" & _
						" and fcBA = '" & Session("fcBA") & "'"
				connStr.execute(q)
				
				if err.number  = 0 then
					connStr.CommitTrans
				end if
			end if
			'response.Write(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & writeoffdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
		
    elseif Request("txtJenis")= "Details" then	
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), writeoffdate)
		If Cek_IsDailyClosing = "FALSE" then
			connstr.BeginTrans
			q	=	" Delete from harvestingwriteoffDetail" & _
					" where  WRITEOFFCODE = '" & writeoff_code & "'" & _
					" and Lineno = '" & LINENO & "'" & _
					" and fcBA = '" & Session("fcBA") & "'"
			connStr.execute(q)
			
			if err.number  = 0 then
				connStr.CommitTrans
			end if
		else
			err.number = 1001
			err.description = "Tanggal """ & writeoffdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		
		cmdView	= "True"
		
	elseif Request("txtJenis")="New"  then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), writeoffdate)
		If Cek_IsDailyClosing = "FALSE" then
			Dim rsPO
			Set rsPO = Server.CreateObject("ADODB.Recordset")
			connstr.BeginTrans
			rsPO.Open " Select WRITEOFFCODE from harvestingwriteoff where WRITEOFFCODE='" & writeoff_code & "' and fcba = '" & Session("fcBA") & "'", connstr
			IF rsPO.eof then
				q	=	" Insert into harvestingwriteoff" & _
						" ( WRITEOFFCODE, WRITEOFFDATE, REFERENCE, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME )" & _
						" values ('" & writeoff_code & "', to_date('" & writeoffdate & "','mm/dd/yyyy'), '" & header_reference & "'" & _
						" , '" & header_remarks  & "', '" & Session("CodeName") & "', '" & Session("CodeName") & "'" & _
						" , '" & Request.ServerVariables("REMOTE_ADDR") &"', '" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
				'response.Write(q)
				connStr.execute(q)
			End if
			rsPO.Close
			set rsPO = Nothing	
			if err.number = 0 then
				strLineMax	=	" select max(to_number(nvl(LINENO,0))) maxLineIdKey from harvestingwriteoffdetail where " & _
								" WRITEOFFCODE = '" & writeoff_code & "' " & _ 
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

				q	=	" Insert into harvestingwriteoffdetail (" & _
						" WRITEOFFCODE" & _
						" ,LINENO" & _
						" , WRITEOFFDATE" & _
						" , FIELDCODE" & _
						" , QUANTITY" & _
						" , REFERENCE" & _
						" , REMARKS" & _		
						" , FCENTRY" & _
						" , FCEDIT" & _
						" , FCIP" & _
						" , FCBA" & _
						" , LASTUPDATE" & _
						" , LASTTIME" & _
						" ) " & _
						" values ("& _
						" '" & writeoff_code & "'" & _
						" , '" & tempMaxLineNo & "'" & _
						" , to_date('" & writeoffdate & "','mm/dd/yyyy')" & _
						" , '" & FIELDCODE & "' , '" & QUANTITY * -1 & "'" & _
						" , '" & DETAIL_REFERENCE & "'" & _
						" , '" & DETAIL_REMARKS & "'" & _
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
			err.description = "Tanggal """ & writeoffdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
		
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), writeoffdate)
		If Cek_IsDailyClosing = "FALSE" then
			q	=	" Update harvestingwriteoffdetail set" & _
					" QUANTITY = '" & QUANTITY * -1 & "'" & _
					" , REFERENCE = '" & DETAIL_REFERENCE & "'" & _
					" , REMARKS = '" & DETAIL_REMARKS & "'" & _
					" , LASTUPDATE = sysdate " & _
					" , LASTTIME = To_Char(sysdate,'HH24:MI') "& _
					" where" & _
					" WRITEOFFCODE = '" & writeoff_code & "'" & _
					" And LINENO = '" & LINENO & "'" & _
					" And FIELDCODE = '" & FIELDCODE & "'" & _
					" and fcBA = '" & Session("fcBA") & "'"
					
					'response.write q
					connStr.execute(q)
					connStr.execute("commit")
		else
			err.number = 1001
			err.description = "Tanggal """ & writeoffdate & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
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
set rsHeaderHWO = server.CreateObject("Adodb.recordset")
	q	=	" select TO_CHAR(WRITEOFFDATE, 'MM/DD/YYYY') WRITEOFFDATE, REFERENCE, REMARKS from harvestingwriteoff  " & _
			" where writeoffcode = '" & writeoff_code & "'" & _
			" and fcBA = '" & session("fcBA") & "'"

	rsHeaderHWO.Open q, connstr
	'response.Write q
	
if Not rsHeaderHWO.Eof then
	writeoffdate		= rsHeaderHWO("WRITEOFFDATE")
	header_reference	= rsHeaderHWO("REFERENCE")
	header_remarks	= rsHeaderHWO("REMARKS")
End if

rsHeaderHWO.Close
set rsHeaderHWO = Nothing


if err.number <> 0 then
    response.Write  "<div class = ""error"">" & err.description & "</div>"
End if
%>

<!-------------ISI Harvesting Write Off Form------------->
 <form name="frmEntry" method="post" action="HarvestingWriteOff.Asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Write Off Code</td>
		<td align="left"><input type="text" maxlength="50" name="writeoff_code" size="30" value="<%=writeoff_code %>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr>
		<td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="writeoffdate" size="10" value="<%=writeoffdate%>">
		<!--<a href="javascript:show_calendar('frmEntry.writeoffdate');" onMouseOver="window.status='Date Picker';return true;" 
		onmouseout="window.status='';return true;">
		<img src="Images/show-calendar.gif" width=24 height=22 border=0 align=top></a>-->
		</td>
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
	    <a href="javascript:launchwin('PrintHWO.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&writeoffcode=<%=writeoff_code%>&writeoffdate=<%=writeoffdate%>&header_reference=<%=header_reference%>&header_remarks=<%=header_remarks%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<a href="HarvestingWriteOff.Asp?writeoff_code=<%=writeoff_code%>&writeoffdate=<%=writeoffdate%>&txtJenis=Header">
			<img src="IMAGES/delete.gif" alt="Delete Data : <%=writeoff_code%>" border="0" onClick="return confirm('Are you sure to delete');">
		</a>

		</td>
		</tr>
    </table>  
 </form>
<%if cmdView = "True" AND LEN(TRIM(writeoff_code)) > 4 Then
   Dim rsharvestingwriteoff,intRecordsPerPage,cColor
   set rsharvestingwriteoff = server.CreateObject("Adodb.recordset")
   q="select writeoffcode, LINENO, writeoffdate,fieldcode,quantity,reference,remarks from harvestingwriteoffdetail  " & _
     " where writeoffcode='" & writeoff_code & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   
   Set rsharvestingwriteoff = Server.CreateObject("ADODB.Recordset")
   rsharvestingwriteoff.CursorType = 3
   rsharvestingwriteoff.CursorLocation = 3
   rsharvestingwriteoff.Open q, connstr
   %>
   <form action="HarvestingWriteOff.Asp" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
        
		
		<tr>
		  <td align="left">Field Code</td>
		  
		  <td align="left">
			<%IF Request("txtJenis")="Edit" then%>
				<select name="FIELDCODE" disabled>
					<%cmdCallCombo "Field",FIELDCODE,"",""%>
				</select><font color="#FF0000">*</font>
				<input type="hidden" name="FIELDCODE" value="<%=FIELDCODE%>" />
			<%Else%>
				<select name="FIELDCODE">
					<%cmdCallCombo "Field",FIELDCODE,"",""%>
				</select><font color="#FF0000">*</font>
			<%End If%>
			
		  </td>
		  </tr>
		
		<tr>
		
		<td align="left">Janjang</td>
		<td align="left">
			<input type="text" size="15" name="QUANTITY" value="<%=QUANTITY%>">
			(+) Adjust Negative, (-) Adjust Positive
		</td>
		
		</tr>
		<tr>
		<td align="left">Reference</td><td align="left"><input type="text" maxlength="100" size="25" name="DETAIL_REFERENCE" value="<%=DETAIL_REFERENCE%>"></td>
		</tr>
		<tr>
			<td align="left">Remarks</td>
			<td align="left"><textarea maxlength="1000" name="DETAIL_REMARKS"  cols="50" rows="3"><%=DETAIL_REMARKS%></textarea></td>
		</tr>
		<tr>
			<td align="left" style="display:none;">LineNo</td>
			<td align="left"><input type="hidden" size="15" name="LINENO" value="<%=LINENO%>"></td>
		</tr>
		
		
	   <TR><td colspan="2" align="left">
		  <input type="hidden" name="writeoff_code" value="<%=writeoff_code%>">
		  <input type="hidden" name="writeoffdate" value="<%=writeoffdate%>">
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
	<TD align="left">Write Off Code</TD>
    <TD align="left">Line No</TD>		
	<TD align="left">Tanggal</TD>
	<TD align="left">Field</TD>
	<TD align="left">Janjang</TD>	
	<TD align="left">Reference</TD>	
	<TD align="left">Remarks</TD>	
	</tr>
   <%
   if Not rsharvestingwriteoff.eof then
      nCountofField=rsharvestingwriteoff.fields.count-1
	  no=1
	  While Not rsharvestingwriteoff.EOF

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
			qCekClose="Select writeoffcode from harvestingwriteoff where writeoffcode='" & rsharvestingwriteoff("writeoffcode") & "'" & "  and fcba = '" & session("fcba") & "'"

			rsCekClose.Open qCekClose, connstr
			if Not rsCekClose.Eof then%>
				<a href="HarvestingWriteOff.Asp?writeoff_code=<%=rsharvestingwriteoff("writeoffcode")%>&FIELDCODE=<%=rsharvestingwriteoff("FIELDCODE")%>&QUANTITY=<%=rsharvestingwriteoff("QUANTITY")%>&LINENO=<%=rsharvestingwriteoff("LINENO")%>&DETAIL_REFERENCE=<%=rsharvestingwriteoff("REFERENCE")%>&DETAIL_REMARKS=<%=rsharvestingwriteoff("REMARKS")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsharvestingwriteoff(1)%>" border="0"></a></td>
				<td valign="middle" width="10"><a href="HarvestingWriteOff.Asp?writeoff_code=<%=rsharvestingwriteoff("writeoffcode")%>&FIELDCODE=<%=rsharvestingwriteoff("FIELDCODE")%>&QUANTITY=<%=rsharvestingwriteoff("QUANTITY")%>&LINENO=<%=rsharvestingwriteoff("LINENO")%>&writeoffdate=<%=writeoffdate%>&txtJenis=Details" onClick="return confirm('Are you sure to delete');">
				<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsharvestingwriteoff(1)%>" border="0"></a></td>
			<%Else%>	 
			    <td colspan="1"><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
			<%End if
			rsCekClose.close
			set rsCekClose = Nothing%>	 
		    <%for i = 0 to nCountofField%>
			    <td align="left" title="<%=rsharvestingwriteoff(i).Name%>" valign="top"><%=rsharvestingwriteoff(i)%></td>
		    <%next%>
		  </tr>
		  <%
		  rsharvestingwriteoff.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsharvestingwriteoff.Close
   set rsharvestingwriteoff = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <!-------------akhir---------------------->
</div></div>
  <!--#include file="bottom.jsp"-->
</body>
</html>
