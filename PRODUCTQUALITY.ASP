<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<html>
<head>
<title>ePlaS - Product Quality</title>
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
%>
 <%


Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim QUALITYDATE, PSTORAGE_STORAGECODE, QUANTITY,FFA,MOISTURE,DIRT,OIL_LOSSES,KERNEL_LOSSES
dim txt_STORAGECODE,Perioddate
Dim Cek_IsDailyClosing : Cek_IsDailyClosing = ""

Dim cmdView,rsPR
txttable=request("txttable")

    txt_STORAGECODE=request("txt_STORAGECODE")
	Perioddate=request("Perioddate")
	QUALITYDATE=request("QUALITYDATE")
	PSTORAGE_STORAGECODE=request("PSTORAGE_STORAGECODE")
    FFA = request("FFA")
	MOISTURE = request("MOISTURE")
	DIRT = request("DIRT")
	OIL_LOSSES = request("OIL_LOSSES")
	KERNEL_LOSSES = request("KERNEL_LOSSES")
	
	
 if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
	if Request("txtJenis")= "Delete" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), QUALITYDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q="Delete from PRODUCTQUALITY where QUALITYDATE = to_Date('" & QUALITYDATE & "','MM/DD/YYYY') And PSTORAGE_STORAGECODE='" & PSTORAGE_STORAGECODE & "' And fcBA='" & Session("fcBA") & "'"
			'response.Write(q)
			connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & QUALITYDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), QUALITYDATE)
		If Cek_IsDailyClosing = "FALSE" then
			q="Insert into PRODUCTQUALITY (QUALITYDATE, PSTORAGE_STORAGECODE, FFA,MOISTURE,DIRT,OIL_LOSSES,KERNEL_LOSSES,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME) " & _
				"values (to_date('" & QUALITYDATE & "','MM/DD/YYYY') " & ",'" &  PSTORAGE_STORAGECODE & "'," &  FFA & "," &  MOISTURE & "," &  DIRT & "," &  OIL_LOSSES & "," &  KERNEL_LOSSES & "," & _
				"'" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
				Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI') )"
				'response.Write(q)
				'response.Flush()
				connStr.execute(q)
		else
			err.number = 1001
			err.description = "Tanggal """ & QUALITYDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
		End If
		cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
		Cek_IsDailyClosing = IsDailyClosing(Session("FCBA"), QUALITYDATE)
		If Cek_IsDailyClosing = "FALSE" then
		   q="Update PRODUCTQUALITY set " & _
			"  FFA = '" & FFA & "'" & _
			", MOISTURE = '" & MOISTURE & "'" & _
			", DIRT = '" & DIRT & "'" & _
			", OIL_LOSSES = '" & OIL_LOSSES & "'" & _
            ",  KERNEL_LOSSES = '" & KERNEL_LOSSES & "'" & _
			" , LASTUPDATE = sysdate " & _
			" , LASTTIME = To_Char(sysdate,'HH24:MI') " & _
			" where QUALITYDATE=to_Date('" & QUALITYDATE & "','MM/DD/YYYY') " & _ 
			" And PSTORAGE_STORAGECODE='" & PSTORAGE_STORAGECODE & "'" & _
			" And fcBA='" & Session("fcBA") & "'"
			
			'response.Write(q)
			connStr.execute(q)
			'response.Flush()
			connStr.execute("commit")
			'RESPONSE.Write "SQLnsddsdya : " & Q
		else
			err.number = 1001
			err.description = "Tanggal """ & QUALITYDATE & """, sudah di closing, sehingga tidak bisa melakukan insert/update/delete data, Silahkan konsultasikan dengan administrator anda"
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
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-flask" aria-hidden="true"></i>Header Product Quality</div><% end if %>
 <form name="frmEntry" method="post" action="PRODUCTQUALITY.ASP">
	<table class="data" width="70%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE PRODUCT QUALITY</strong></font></div></td>
          </TR>

		<tr><td align="left">Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="Perioddate" size="10" value="<%=Perioddate%>">
		
		</td></tr>
		<TR>
			<td align="left">Storage Code</td>
			<td align="left"> <select name="txt_STORAGECODE">
			<%cmdCallComboSetup "PRODUCTSTORAGE",txt_STORAGECODE,"fcba",session("fcba")%>
			</select></td>
		</tr>		
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
		<a href="PrintPRODUCTQUALITY.ASP?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&Perioddate=<%=Perioddate%>&txt_STORAGECODE=<%=REQUEST("txt_STORAGECODE")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>
		</td>
		</tr>
    </table>  
 </form>
<%
if cmdView = "True" Then
	Dim rsGAD,intRecordsPerPage,cColor
	set rsGAD = server.CreateObject("Adodb.recordset")
	
	'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	q = " select to_char(QUALITYDATE,'MM/DD/YYYY') QUALITYDATE, PSTORAGE_STORAGECODE, FFA,MOISTURE,DIRT,OIL_LOSSES,KERNEL_LOSSES,FCENTRY, FCEDIT, FCIP, FCBA, LASTUPDATE, LASTTIME" & _
	    " from PRODUCTQUALITY" & _
	    " where fcBA = '" & session("fcBA") & "'"
	if  len(trim(Perioddate)) > 0 then
	q = q & " and to_char(QUALITYDATE,'MM/DD/YYYY') = to_char(to_date('" & Perioddate & "', 'MM/DD/YYYY'),'MM/DD/YYYY')"
	end if
	if  len(trim(txt_STORAGECODE)) > 0 then
	q = q & " and PSTORAGE_STORAGECODE='" & txt_STORAGECODE & "'"
	end if
	q = q
	'response.Write(q)
	'response.Flush()
	rsGAD.CursorType = 3
	rsGAD.CursorLocation = 3
	rsGAD.Open q, connstr
%>


   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
   <form action="PRODUCTQUALITY.ASP" method="post" name="eForm">
   <table class="data" cellpadding="0" cellspacing="0" width="70%">
		<tr>
			<td align="left">Reading Date</td>
			<td align="left">
				<%IF Request("txtJenis")="Edit" then%>
				<input type="text" maxlength="10" name="QUALITYDATE" size="10" value="<%=QUALITYDATE%>" readonly="True" style="background:#999999">
				<%Else%>
				<input class="datepicker" type="text" maxlength="10" name="QUALITYDATE" size="10" value="<%=QUALITYDATE%>">
				<%End IF%>
				<font color="#FF0000">*</font>
			</td>
		</tr>
          <Tr> 
		<%IF Request("txtJenis")="Edit" then%>		
            <td align="left"><font face="Verdana" size="1">Storage Code</font></td>
            <td align="left"> <select name="PSTORAGE_STORAGECODE">
                <%cmdCallComboSetup "PRODUCTSTORAGE",PSTORAGE_STORAGECODE,"fcba",session("fcba")%>
              </select> <font color="#FF0000">*</font> </td>
		<%else%>
				<td align="left"><font face="Verdana" size="1">Storage Code</font></td>
				<td align="left"> <select name="PSTORAGE_STORAGECODE">
				<%cmdCallComboSetup "PRODUCTSTORAGE",PSTORAGE_STORAGECODE,"fcba",session("fcba")%>
				</select> <font color="#FF0000">*</font> </td>
		<%end if%>
          </Tr>
		
		<tr>
		<td align="left">FFA (%) </td><td align="left"><input type="text" size="15" name="FFA" value="<%=FFA%>"></td>
		</tr>
		<tr>
		<td align="left">MOISTURE (%)</td><td align="left"><input type="text" size="15" name="MOISTURE" value="<%=MOISTURE%>"></td>
		</tr>
		<tr>
		<td align="left">DIRT (%) </td><td align="left"><input type="text" size="15" name="DIRT" value="<%=DIRT%>"></td>
		</tr>
		<tr>
		<td align="left">OIL LOSSES (%)</td><td align="left"><input type="text" size="15" name="OIL_LOSSES" value="<%=OIL_LOSSES%>"></td>
		</tr>
		<tr>
		<td align="left">KERNEL LOSSES (%) </td><td align="left"><input type="text" size="15" name="KERNEL_LOSSES" value="<%=KERNEL_LOSSES%>"></td>
		</tr>
		
	    <TR><td colspan="2" align="left">
		  <input type="hidden" name="Perioddate" value="<%=Perioddate%>">
		  <input type="hidden" name="txt_STORAGECODE" value="<%=txt_STORAGECODE%>">
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
   <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Daftar Data</div><% end if %>
   <div class="fbp-doc-scroll">
   <table class="data fbp-doc-lines" >
	<thead>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD colspan="2">&nbsp;</TD>
	<TD align="left">DATE</TD>	
	<TD align="left">STORAGE CODE</TD>	
	<TD align="left">FFA (%)</TD>	
	<TD align="left">MOISTURE (%)</TD>	
	<TD align="left">DIRT (%)</TD>	
	<TD align="left">OIL LOSSES (%)</TD>
	<TD align="left">KERNEL LOSSES (%)</TD>
	</tr>
	</thead>
	<tbody>
        <%
   if Not rsGAD.eof then
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
			<a href="PRODUCTQUALITY.ASP?Perioddate=<%=Perioddate%>&txt_STORAGECODE=<%=txt_STORAGECODE%>&QUALITYDATE=<%=rsGAD("QUALITYDATE")%>&txttable=<%=request("txttable")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&FFA=<%=rsGAD("FFA")%>&MOISTURE=<%=rsGAD("MOISTURE")%>&DIRT=<%=rsGAD("DIRT")%>&OIL_LOSSES=<%=rsGAD("OIL_LOSSES")%>&KERNEL_LOSSES=<%=rsGAD("KERNEL_LOSSES")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
			<td valign="middle" width="20">
			<a href="PRODUCTQUALITY.ASP?Perioddate=<%=Perioddate%>&txt_STORAGECODE=<%=txt_STORAGECODE%>&QUALITYDATE=<%=rsGAD("QUALITYDATE")%>&PSTORAGE_STORAGECODE=<%=rsGAD("PSTORAGE_STORAGECODE")%>&txttable=<%=request("txttable")%>&txtJenis=Delete" onClick="return confirm('Are you sure to Delete');"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0"></a>
			</td>
		    <%for i = 0 to rsGAD.fields.count-7%>
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

