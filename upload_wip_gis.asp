<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<html>
<head>
<title>ePlaS - Upload Manual WIP GIS</title>
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
<script language="JavaScript">
        function PopupTType() {
		var txtyear_ = document.getElementById("txtYear_");
		var txtmonth_ = document.getElementById("txtMonth_");
		var yr_ = txtyear_.options[txtyear_.selectedIndex].value;
		var mn_ = txtmonth_.options[txtmonth_.selectedIndex].value;
		
		window.location.href = "upload_wip_gis.Asp?txtJenis=Preview&txtYear="+yr_+"&txtMonth="+mn_;
        }		
</script>  
<script language="JavaScript">
        function ShowUpload() {
		var txtyear_ = document.getElementById("txtYear_");
		var txtmonth_ = document.getElementById("txtMonth_");
		var yr_ = txtyear_.options[txtyear_.selectedIndex].value;
		var mn_ = txtmonth_.options[txtmonth_.selectedIndex].value;
		
		window.location.href = "upload_wip_gis_preview.Asp?txtYear="+yr_+"&txtMonth="+mn_;
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
<%
'-------------awal-----------------------
function copy_to_wipplan_gis
Dim q_copy
q_copy = " DELETE FROM workinprogressplan_gis WHERE reference='" & fcCode & "';" & _
		 " INSERT INTO workinprogressplan_gis(agreementcode, wipdate, agreementlineno, lineno, qty, output, postedflag, isgenerated, fcentry, fcedit, fcip, fcba, wipamt, reference, remarks, lastupdate, lasttime, agreementfinalno, unit, jobcode, locationtype, locationcode, chargeunit, isfinal) " & _ 
		 " WITH agreement_code AS " & _
		 " (" & _
		 "         SELECT DISTINCT a.agreementcode,a.jobcode,Trim(a.reference) actname,b.contractorcode, c.fccode fcba,c.fccompanycode,b.startdate, b.finishdate,a.lineno,a.rate  FROM contractagreementdetail a  " & _
		 " 				 INNER JOIN contractagreement b ON (a.agreementcode=b.agreementcode AND a.fcba=b.fcba ) " & _
		 "          INNER JOIN businessunit c ON (a.fcba=c.fccode)" & _
		 " 				 WHERE " & _
		 " 				 fctype='HO'" & _
		 " 				 AND isclose='N' " & _
		 " )" & _
		 " SELECT agreementcode,tanggal_wip,agreementlineno,Row_Number() OVER( ORDER BY agreementcode,tanggal_wip,job_code,remarks,locationcode) lineno," & _
		 "        qty,output,postedflg,isgenerate,fcentry,fcedit, fcip,fcba,wipamt," & _
		 "        reference,remarks,lastupdate,lasttime,agreementfinalno,unit,job_code,locationtype,locationcode,chargeunit,isfinal" & _
		 " FROM" & _
		 " (" & _
		 " SELECT y.agreementcode,x.tanggal_wip,y.lineno agreementlineno,0 lineno,Sum(luas) qty,0 output,0 postedflg,'FALSE' isgenerate,'" & Session("CodeName") & "' fcentry,'" & Session("CodeName") & "' fcedit,'::1' fcip,y.fcba,Sum(luas)*y.rate wipamt," & _
		 "        x.docno reference,x.actname remarks,SYSDATE lastupdate,'00:00' lasttime,'' agreementfinalno,Sum(luas) unit,x.job_code,'FF' locationtype,x.blok locationcode,x.kebun chargeunit,0 isfinal" & _
		 " FROM upload_dtl_wip_gis_tab  x " & _
		 " left JOIN agreement_code y ON (x.pt=y.fccompanycode AND x.kontraktor=y.contractorcode AND x.job_code=y.jobcode AND x.actname=y.actname AND tanggal_wip BETWEEN startdate AND finishdate)" & _
		 " WHERE x.docno='" & fcCode & "'" & _ 
		 " GROUP BY  y.agreementcode,x.tanggal_wip,y.fcba,y.lineno,y.rate,x.docno,x.actname,x.job_code,x.blok,x.kebun" & _
		 " ORDER BY  x.tanggal_wip,x.job_code,x.actname,x.blok " & _
		 " )a" & _
		 " ORDER BY lineno;"

		connStr.execute(q_copy)

end function


Dim xCode,fcCode,cmdView,UPLOADDOCCODE,UPLOADDOCDATE,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
Dim Remarks
Dim txtModule,txtTable,txtSearch,txtField,q,nCountofField,NO,txtJenis,txtYear,txtMonth
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
txtYear = request("txtYear")
txtMonth = request("txtMonth")

UPLOADDOCDATE = Request("UPLOADDOCDATE")
if len(trim(UPLOADDOCDATE)) = 0 then UPLOADDOCDATE = textDate(date)


Dim kodekaryawan : kodekaryawan = request("kodekaryawan")
Dim timein : timein = Request("timein")
Dim timeout : timeout = request("timeout")
Dim keterangan : keterangan = request("keterangan")
 
 '--- Action ke Database di Oracle [ insert,edit,delete ]
 if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Validate" or Request("txtJenis")= "Transfer" or Request("txtJenis")= "RePlanned"  Then
	REMARKS=request("REMARKS")
	if Request("txtJenis")= "Header" then
		connstr.BeginTrans
		q = " delete from upload_dtl_wip_gis_tab" & _
		    " where docno = '" & fcCode & "'" 
		'response.write q
		connStr.execute(q)
		
		q = " delete from upload_wip_gis_tab" & _
		    " where docno = '" & fcCode & "'" 
		'response.write q
		connStr.execute(q)
		
		if err.number  = 0 then
			connStr.CommitTrans
		else
			connStr.RollbackTrans
		end if
		cmdView	= "True"
	elseif Request("txtJenis")= "Details" then
		'connstr.BeginTrans
		'q = " delete from userfinger.attendance_tab" & _
		'    " where no_ba is not null" & _
		'    "       and no_ba = '" & fcCode & "'" & _
		'    "       and fcbatobe = '" & Session("fcBA") & "'" & _
		'	"       and empnotobe = '" & kodekaryawan & "'" & _
		'	"       and tdate = to_date('" & uploaddocdate & "', 'mm/dd/yyyy')"
		'response.write q
		'connStr.execute(q)
		'if err.number  = 0 then
		'	connStr.CommitTrans
		'end if
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
	   'connstr.BeginTrans
	   strQueryStatus = "UPDATE  upload_wip_gis_tab SET rowstateid='Approved',lastupdate=sysdate,fcentry='" & Session("CodeName") & "' WHERE docno='" & fcCode & "' "
	   'response.write strQueryStatus
	   connStr.execute(strQueryStatus)	 
	   call copy_to_wipplan_gis
       'cmdView	= "True"
	elseif Request("txtJenis")= "Transfer" then
	   'connstr.BeginTrans
	   'strQueryStatus = "BEGIN userfinger.ManualFinger_API.TransferDetail('" & fcCode & "','" & Session("fcBA") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "' ); END; "
	   'response.write strQueryStatus
	   'connStr.execute(strQueryStatus)	 
       'cmdView	= "True"
	elseif Request("txtJenis")= "RePlanned" then
	   'connstr.BeginTrans
	   'strQueryStatus = "BEGIN userfinger.ManualFinger_API.RePlannedDetail('" & fcCode & "','" & Session("fcBA") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "' ); END; "
	   'response.write strQueryStatus
	   'connStr.execute(strQueryStatus)	 
       'cmdView	= "True"
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
q = " select * from upload_wip_gis_tab" & _
    " where docno = '" & fcCode & "'" 
rsgetdatahdr.Open q, connstr
if Not rsgetdatahdr.Eof then
	fcfilename = rsgetdatahdr("filename")
	RowStateId = rsgetdatahdr("RowStateId")
End if
rsgetdatahdr.Close
Set rsgetdatahdr = Nothing
 
'if err.number <> 0 then
'	connStr.RollbackTrans
'    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
'End if

Dim val_IsTrue : val_IsTrue = "N"
Dim oRecordSet, oSQLQuery
'oSQLQuery = " select 'Y' istrue" & _
'            " from periodcontrol" & _
'	        " where fcba = '" & Session("fcBA") & "'" & _
'	        "       and to_date ('" & uploaddocdate & "', 'mm/dd/yyyy') >= currentdate" & _
'	        "       and to_date ('" & uploaddocdate & "', 'mm/dd/yyyy') <= last_day(add_months(currentdate, 1))" & _
'	        "       and upper(system) = upper('Employee')"
			
'set oRecordSet = server.CreateObject("Adodb.recordset")
'oRecordSet.Open oSQLQuery, connstr
'If Not oRecordSet.Eof then
'	val_IsTrue = oRecordSet("istrue")
'Else
	val_IsTrue = "N"
'End if
'response.Write oSQLQuery
'oRecordSet.Close
'Set oRecordSet = Nothing

 '-------------ISI Payment Form 
 %>
 <% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-upload" aria-hidden="true"></i>Header Upload WIP GIS</div><% end if %>
 <form name="frmEntry" method="post" action="upload_wip_gis.asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">No. Document</td>
		<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">*</font></td>
		</tr>
				<tr>
					<td align="left" width="150">Period</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM, v_txtMonth, v_txtYear
							mon = Month(Date)
							If len(v_txtMonth)= 0 then v_txtMonth = mon

							Yr = Year(Date)
							If len(v_txtYear) = 0 then v_txtYear = Yr
							MinYr = Yr - 2
							MaxYr = Yr + 2
							
						%>
					<td align="left">
							<select name="txtYear" id="txtYear_">
							<%	For i = MinYr to MaxYr %>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%	Next %>
							</select>
							<select name="txtMonth" id="txtMonth_"> 
							<%	For i = 1 to 12	%>
									<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%	Next %>
							</select>
					
				</tr>
		<tr>
			<td align="left">Nama File</td>
			 <td align="left">
			 <input type="text" maxlength="50" name="fcfilename" size="30" value="<%=fcfilename%>">
			 <input Type="Button"  name="Add" value="Upload" onClick="javascript:ShowUpload();"></td>
		</tr>
		
		<tr>
			<td align="left">Status</td>
			<td align="left"><input disabled type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr><td colspan="2"  align="left" class="fbp-doc-actions">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input type="Button" value="Preview" onClick="javascript:PopupTType();">&nbsp;
		<input Type="Submit"  name="Add" value="Select">&nbsp;&nbsp;
			<a href="javascript:launchwin('PrintBAFP.Asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=fcCode%>&UPLOADDOCDATE=<%=UPLOADDOCDATE%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" then %>
			<a href="upload_wip_gis.asp?fcCode=<%=fcCode%>&txtJenis=Validate" onClick="return confirm('Are you sure to Approved this data :  <%=fcCode%> ?');"> Approved </a>
		<% end if%>
		<!-- Period Control In Here -->
		<% if RowStateId <> "Approved" then %> 
			<a href="upload_wip_gis.asp?fcCode=<%=fcCode%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<% end if%>
		<!-- End Period Control In Here -->
		
		</td>
		</tr>
    </table>  
 </form>
 
<%
	if txtJenis = "Preview"  Then
		Dim rsdataPrev
		set rsdataPrev = server.CreateObject("Adodb.recordset")
		
		q = " select docno" & _
			"        , tdate" & _
			"        , filename,rowstateid" & _
			"        ,lastupdate,fcentry " & _
			" from upload_wip_gis_tab" & _
			" where to_char(tdate,'yyyy')='" & txtYear  & "' and to_number(to_char(tdate,'mm'))=" & txtMonth & _
			" order by lastupdate DESC"
			
        rsdataPrev.CursorType = 3
		rsdataPrev.CursorLocation = 3
		rsdataPrev.Open q, connstr
   
%>

<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-edit" aria-hidden="true"></i>Detail Item</div><% end if %>
<form action="upload_wip_gis.asp" method="post" name="eForm">
   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<input type="hidden" name="fcCode" value="<%=fcCode%>">
	<input type="hidden" name="txtYear" value="<%=txtYear%>">
	<td align="left">No Dokument</td>	
	<td align="left">Tanggal</td>	
	<td align="left">File</td>	
	<td align="left">Status</td>	
	<td align="left">LastUpdate</td>	
	<td align="left">FcEntry</td>	
	</tr>   
<%
   if Not rsdataPrev.eof then
      While Not rsdataPrev.EOF	  
%>
			<tr >
				<!-- Period Control In Here -->
					<td align="left"  valign="top">					
					<a href="upload_wip_gis.asp?fccode=<%=rsdataPrev("docno")%>&cmdView=TRUE">
					<font face="Verdana" size="1"><%=rsdataPrev("docno")%></font></a>
					</td>
					<td align="left"  valign="top"><%=rsdataPrev("tdate")%></td>
					<td align="left"  valign="top"><%=rsdataPrev("filename")%></td>
					<td align="left"  valign="top"><%=rsdataPrev("rowstateid")%></td>
					<td align="left"  valign="top"><%=rsdataPrev("lastupdate")%></td>
					<td align="left"  valign="top"><%=rsdataPrev("fcentry")%></td>
					
				<!-- End Period Control In Here -->
				
			    
			
		  </tr>
		  <%
		  rsdataPrev.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend
   else
      response.Write("Data is empty")	  
   end if 
   rsdataPrev.Close
   set rsdataPrev = Nothing%>
   </TABLE>
   </form>
<%End if%>  

<%
	if cmdView = "True" AND LEN(TRIM(fcCode)) <> 0 Then
		Dim rsdata, intRecordsPerPage, cColor
		set rsdata = server.CreateObject("Adodb.recordset")
		
		q = " select id " & _
			"        , docno" & _
			"        , pt" & _
			"        , kebun" & _
			"        , divisi" & _
			"        , blok" & _
			"        , job_code" & _
			"        , actname" & _
			"        , kontraktor" & _
			"        , tanggal_wip" & _
			"        , luas" & _
			" from upload_dtl_wip_gis_tab" & _
			" where docno = '" & fcCode & "' Order by id" 
			
        'response.Write q
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open q, connstr
   
%>
   


   
   <form action="upload_wip_gis.asp" method="post" name="eForm">
    <%
	
	DIM DISABLED_
	IF Request("txtJenis")="Edit" then
	DISABLED_ = "DISABLED"
	Else
	DISABLED_ = ""
	END IF%>
		  
	
	    <TR><td colspan="2" align="left">
		  <!-- diambil dari form yang diatas -->
	

   <table class="data">
	<BR>
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<input type="hidden" name="fcCode" value="<%=fcCode%>">
	<input type="hidden" name="UPLOADDOCDATE" value="<%=UPLOADDOCDATE%>">
	<td align="left">No</td>
	<td align="left">No.Doc</td>	
	<td align="left">PT</td>	
	<td align="left">Kebun</td>	
	<td align="left">Divisi</td>	
	<td align="left">Blok</td>	
	<td align="left">Job.Code</td>
	<td align="left">Aktivity</td>
	<td align="left">Kontraktor</td>
	<td align="left">Tgl.wip</td>
	<td align="left">Luas</td>
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
				<td align="left" valign="top"><%=rsdata("id")%></td>
				<!-- Period Control In Here -->
					<td align="left" valign="top"><%=rsdata("docno")%></td>
					<td align="left" valign="top"><%=rsdata("pt")%></td>
					<td align="left" valign="top"><%=rsdata("kebun")%></td>
					<td align="left" valign="top"><%=rsdata("divisi")%></td>
					<td align="left" valign="top"><%=rsdata("blok")%></td>
					<td align="left" valign="top"><%=rsdata("job_code")%></td>
					<td align="left" valign="top"><%=rsdata("actname")%></td>
					<td align="left" valign="top"><%=rsdata("kontraktor")%></td>
					<td align="left" valign="top"><%=rsdata("tanggal_wip")%></td>
					<td align="right" valign="top"><%=rsdata("luas")%></td>
				<!-- End Period Control In Here -->
				
			    
			
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
   </TABLE>
   </form>
<%End if%>  
 
  <%'-------------akhir----------------------%>
</div></div>
  <!--#include file="bottom.jsp"-->
 
</body>
</html>
