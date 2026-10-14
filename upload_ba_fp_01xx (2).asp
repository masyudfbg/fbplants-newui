<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<html>
<head>
<title>ePlaS - Upload Manual Fingerprint</title>
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
<%
'-------------awal-----------------------

Dim xCode,fcCode,cmdView,UPLOADDOCCODE,UPLOADDOCDATE,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
Dim Remarks
Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
Dim strLineMax, rsstrLineMax, tempMaxLineNo
Dim RowStateId
Dim PRCODE, PRLINEID
Dim tempStringError, rsTempPurchaseRequest
Dim totalPRData
Dim ApprovedQty, ORDEREDQTY
Dim strQueryStatus, rsQueryStatus 

Dim purchase_doc,item,vendor,matl_group,material,purchase_organization,puchase_group,plant,sloc,doc_date,order_quantity,order_unit,LineNo
Dim net_order_value,ccy,condition_type,amount_cond_type,transporter,pr_no,pr_item,po_release,error_flag,error_desc  
Dim fcfilename

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")

fcCode = request("fcCode")
if len(trim(fcCode)) = 0 then fcCode = "(NEW)"


txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]

UPLOADDOCDATE = Request("UPLOADDOCDATE")
if len(trim(UPLOADDOCDATE)) = 0 then UPLOADDOCDATE = textDate(date)


Dim kodekaryawan : kodekaryawan = request("kodekaryawan")
Dim timein : timein = Request("timein")
Dim timeout : timeout = request("timeout")
Dim keterangan : keterangan = request("keterangan")
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Validate" or Request("txtJenis")= "Transfer"  Then
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		q = " delete from userfinger.attendance_tab" & _
		    " where no_ba is not null" & _
		    "       and no_ba = '" & fcCode & "'" & _
		    "       and fcbatobe = '" & Session("fcBA") & "'"
		'response.write q
		connStr.execute(q)
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		connstr.BeginTrans
		q = " delete from userfinger.attendance_tab" & _
		    " where no_ba is not null" & _
		    "       and no_ba = '" & fcCode & "'" & _
		    "       and fcbatobe = '" & Session("fcBA") & "'" & _
			"       and empnotobe = '" & kodekaryawan & "'" & _
			"       and tdate = to_date('" & uploaddocdate & "', 'mm/dd/yyyy')"
		'response.write q
		connStr.execute(q)
		if err.number  = 0 then
			connStr.CommitTrans
		end if
		cmdView	= "True"
	elseif Request("txtJenis")="New" then
	  ' connstr.BeginTrans
	  ' if err.number  = 0 then
			' connStr.CommitTrans
	  ' end if
      ' cmdView	= "True"
	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
	   ' connstr.BeginTrans		
		' if err.number  = 0 then
		   ' connStr.CommitTrans
		' end if
        ' cmdView	= "True"
	elseif Request("txtJenis")= "Validate" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN userfinger.ManualFinger_API.ValidateDetail('" & fcCode & "','" & Session("fcBA") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Request.ServerVariables("REMOTE_ADDR") & "' ); END; "
	   'response.write strQueryStatus
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	   
	elseif Request("txtJenis")= "Transfer" then
	   connstr.BeginTrans
	   strQueryStatus = "BEGIN userfinger.ManualFinger_API.TransferDetail('" & fcCode & "','" & Session("fcBA") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & Request.ServerVariables("REMOTE_ADDR") & "' ); END; "
	   'response.write strQueryStatus
	   connStr.execute(strQueryStatus)	 
       cmdView	= "True"
	
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
	
 else
    cmdView	= request("cmdView")
	'Response.Write " else"
 End if

Dim rsgetdatahdr
set rsgetdatahdr = server.CreateObject("Adodb.recordset")
q = " select * from userfinger.uploadheader_mf_tab" & _
    " where docno = '" & fcCode & "'" & _
	"       and fcBA = '" & session("fcBA") & "'"
rsgetdatahdr.Open q, connstr
if Not rsgetdatahdr.Eof then
	fcfilename = rsgetdatahdr("filename")
	RowStateId = rsgetdatahdr("RowStateId")
End if
 
 
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if

Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
oSQLQuery = " select 'Y' istrue" & _
            " from periodcontrol" & _
	        " where fcba = '" & Session("fcBA") & "'" & _
	        "       and to_date ('" & uploaddocdate & "', 'mm/dd/yyyy') >= currentdate" & _
	        "       and to_date ('" & uploaddocdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
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

 '-------------ISI Payment Form 
 %>
 <form name="frmEntry" method="post" action="upload_ba_fp_01.asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">No. Document</td>
		<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">*</font></td>
		</tr>
		
		<tr>
			<td align="left">Nama File</td>
			 <td align="left">
			 <input type="text" maxlength="50" name="fcfilename" size="30" value="<%=fcfilename%>">
			 <input Type="Button"  name="Add" value="Upload" onClick="window.location.href='upload_ba_fp_02.asp'"></td>
		</tr>
		
		<tr>
			<td align="left">Status</td>
			<td align="left"><input disabled type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
			<a href="javascript:launchwin('PrintBAFP.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&UPLOADDOCDATE=<%=UPLOADDOCDATE%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" then %>
			<a href="upload_ba_fp_01.asp?fcCode=<%=fcCode%>&txtJenis=Validate" onClick="return confirm('Are you sure to Validate this data :  <%=fcCode%> ?');"> Validate </a>
		<% elseif RowStateId = "Validated" then %>	
			<a href="upload_ba_fp_01.asp?fcCode=<%=fcCode%>&txtJenis=Validate" onClick="return confirm('Are you sure to Validate this data :  <%=fcCode%> ?');"> Validate</a>
		<% elseif RowStateId = "Ready" then %>	
			<a href="upload_ba_fp_01.asp?fcCode=<%=fcCode%>&txtJenis=Transfer" onClick="return confirm('Are you sure to Transfer this data :  <%=fcCode%> ?');"> Transfer</a>
		<% end if%>
		<%If val_IsTrue = "Y" Then%>
			<a href="upload_ba_fp_01.asp?fcCode=<%=fcCode%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<%Else%>
			<img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18">
		<%End if%>
		
		</td>
		</tr>
    </table>  
 </form>
 
<%
	if cmdView = "True" AND LEN(TRIM(fcCode)) <> 0 Then
		Dim rsdata, intRecordsPerPage, cColor
		set rsdata = server.CreateObject("Adodb.recordset")
		
		q = " select fcba" & _
			"        , no_ba" & _
			"        , empno" & _
			"        , tdate" & _
			"        , timein" & _
			"        , timeout" & _
			"        , userfinger.desc_of_ba_reasontype(ba_reasontype) ba_reasontype" & _
			"        , error_flag" & _
			"        , error_desc" & _
			" from userfinger.uploaddetail_mf_tab" & _
			" where docno = '" & fcCode & "'" & _
			"       and fcba = '" & session("fcBA") & "'"
			
        'response.Write q
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open q, connstr
   
%>
   


   
   <form action="upload_ba_fp_01.asp" method="post" name="eForm">
   <!--
   <table class="data" cellpadding="0" cellspacing="0" width="100%">
    <%
	
	DIM DISABLED_
	IF Request("txtJenis")="Edit" then
	DISABLED_ = "DISABLED"
	Else
	DISABLED_ = ""
	END IF%>
		  
		<tr>
		   <td  align="left">Kode Karyawan</td>
		   <td align="left">			
				 <input type="text" size="15" name="kodekaryawan" value="<%=kodekaryawan%>" <%=DISABLED_%>>&nbsp;<font color="#FF0000">
			</td>	 
		</tr>


		<tr>
			<td align="left">Check In (hh:mm)</td>
			<td align="left"><input type="text" size="15" maxlength = "5" name="timein" value=<%=timein%>><font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Check Out (hh:mm)</td>
			<td align="left"><input type="text" size="15" maxlength = "5" name="timeout" value=<%=timeout%>><font color="#FF0000">
			</td>
		</tr>
		
		<tr>
			<td align="left">Keterangan</td>
			<td align="left"><textarea name="keterangan" maxlength = "100" cols="50" rows="3"><%=keterangan%></textarea>&nbsp;<font color="#FF0000">
		</tr>
		
	    <TR><td colspan="2" align="left">
	-->
		  <!-- diambil dari form yang diatas -->
	<!--
	      <input type="hidden" name="fcCode" value="<%=fcCode%>">
		  <input type="hidden" name="doc_date" value="<%=doc_date%>">
		  <input type="hidden" name="purchase_doc" value="<%=purchase_doc%>">
		  <input type="hidden" name="UPLOADDOCDATE" value="<%=UPLOADDOCDATE%>">
		  <input type="hidden" name="xCode" value="<%=xCode%>">
		  <input type="hidden" name="Remarks" value="<%=Remarks%>">
		  <input type="hidden" name="LineNo" value="<%=LineNo%>">
		  
		  <input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		  <input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		  <%IF Request("txtJenis")="Edit" then%>
		       <input type="hidden" name="txtJenis" value="Edit">
		       <input Type="Submit"  name="Add" value="Update">
		  <%Else%>
		       
		  <%End if%>
		  <input Type="Reset" name="Reset" value="Reset">
		  <input Type="Button" value="Back" name="B2" OnClick="javascript:history.back(1)"></td>
	  </tr>
   </table>
   -->

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<input type="hidden" name="fcCode" value="<%=fcCode%>">
	<input type="hidden" name="UPLOADDOCDATE" value="<%=UPLOADDOCDATE%>">
	<td align="left">NO</td>
	<td align="left">&nbsp;</td>
	<td align="left">No Berita Acara</td>	
	<td align="left">Tanggal</td>	
	<td align="left">Kode Karyawan</td>	
	<td align="left">Check In</td>	
	<td align="left">Check Out</td>	
	<td align="left">Keterangan</td>
	<td align="left">Error Flag</td>
	<td align="left">Error Desc</td>
	
	</tr>   
<%
   if Not rsdata.eof then
      nCountofField=rsdata.fields.count-1
	  no=1
	  While Not rsdata.EOF

		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if
%>
			<tr bgcolor="<%=cColor%>">
				<td align="left" valign="top"><%=NO%></td>
				<%If val_IsTrue = "Y" Then%>
				<!--
					<td valign="middle" width="10">
						<a href="upload_ba_fp_01.asp?fcCode=<%=rsdata("no_ba")%>&kodekaryawan=<%=rsdata("empnotobe")%>&uploaddocdate=<%=rsdata("tdate")%>&txtJenis=Edit">
							<img src="IMAGES/edit.gif" alt="Edit Data : <%=rsdata(2)%>" border="0">
						</a>
					</td>
				-->
					<td valign="middle" width="10">
						<a href="upload_ba_fp_01.asp?fcCode=<%=rsdata("no_ba")%>&kodekaryawan=<%=rsdata("empnotobe")%>&uploaddocdate=<%=rsdata("tdate")%>&txtJenis=Details">
							<img src="IMAGES/delete.gif" alt="Delete Data : <%=rsdata(1)%>" border="0" onClick="return confirm('Are you sure to delete');">
						</a>
					</td>
				<%Else%>
					<td><img src="icon/IconClose.JPG" alt="Data is Close" border="0" width="16" height="18"></td>
				<%End if%>
			
				<td align="left" title="<%=rsdata("no_ba").Name%>" valign="top"><%=rsdata("no_ba")%></td>
				<td align="left" title="<%=rsdata("tdate").Name%>" valign="top"><%=rsdata("tdate")%></td>
				<td align="left" title="<%=rsdata("empno").Name%>" valign="top"><%=rsdata("empno")%></td>
				<td align="left" title="<%=rsdata("timein").Name%>" valign="top"><%=rsdata("timein")%></td>
				<td align="left" title="<%=rsdata("timeout").Name%>" valign="top"><%=rsdata("timeout")%></td>
				<td align="left" title="<%=rsdata("ba_reasontype").Name%>" valign="top"><%=rsdata("ba_reasontype")%></td>
				<td align="left" title="<%=rsdata("error_flag").Name%>" valign="top"><%=rsdata("error_flag")%></td>
				<td align="left" title="<%=rsdata("error_desc").Name%>" valign="top"><%=rsdata("error_desc")%></td>
		  </tr>
		  <%
		  rsdata.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsdata.Close
   set rsdata = Nothing%>
   </div>
   </TABLE>
   </form>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->
 
</body>
</html>
