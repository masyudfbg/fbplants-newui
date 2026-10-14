<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->


<html>
<head>
<!--#include file="aspupload.asp"-->
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
function onSubmitForm() {
    var formDOMObj = document.frmSend;
    if (formDOMObj.txtFile.value == "" && formDOMObj.fcCode.value == "(NEW)")
        alert("Please press the Browse button and pick a file.")
	else
        return true;
    return false;
}
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

 Dim xCode,fcCode,cmdView,uploaddoccode,uploaddocdate,LineNoKey, ITEMCODE,QUANTITY,AMOUNT,OTHERITEMDESC
 Dim Remarks,fcfilename
 Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis
 Dim strLineMax, rsstrLineMax, tempMaxLineNo
 Dim RowStateId
 Dim PRCODE, PRLINEID
 Dim tempStringError, rsTempPurchaseRequest
 Dim totalPRData
 Dim ApprovedQty, ORDEREDQTY
 Dim strQueryStatus, rsQueryStatus 
 Dim txtYear,txtMonth

 Dim purchase_doc,item,vendor,matl_group,material,purchase_organization,puchase_group,plant,sloc,doc_date,order_quantity,order_unit
 Dim net_order_value,ccy,condition_type,amount_cond_type,transporter,pr_no,pr_item,po_release,error_flag,error_desc  
 Dim sql,namefile,temptxtJenis,fcCodeNew

 fcCode = "(NEW)"
 fcfilename = ""
 txtYear  = Request("txtYear")
 txtMonth = Request("txtMonth")
 
 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 temptxtJenis = Request("txtJenis")  
 
 Remarks = Request("Remarks")

 error_flag				= request("error_flag")
 error_desc 			= request("error_desc")
 
 dim rsAUTONUM,AutoNum,numq,txtdateNow
 Dim uploadsDirVar
 
'uploadsDirVar = Server.MapPath(".") & "\tempUploads\" 
 uploadsDirVar = "c:/temptemp/" 
  

function Environment()
    Dim fso, fileName, testFile, streamTest
    Environment = ""
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    if not fso.FolderExists(uploadsDirVar) then
        Environment = "<B>Folder " & uploadsDirVar & " does not exist.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    fileName = uploadsDirVar & "\test.txt"
    on error resume next
    Set testFile = fso.CreateTextFile(fileName, true)
    If Err.Number<>0 then
        Environment = "<B>Folder " & uploadsDirVar & " does not have write permissions.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
        exit function
    end if
    Err.Clear
    testFile.Close
    fso.DeleteFile(fileName)
    If Err.Number<>0 then
        Environment = "<B>Folder " & uploadsDirVar & " does not have delete permissions</B>, although it does have write permissions.<br>Change the permissions for IUSR_<I>computername</I> on this folder."
        exit function
    end if
    Err.Clear
    Set streamTest = Server.CreateObject("ADODB.Stream")
    If Err.Number<>0 then
        Environment = "<B>The ADODB object <I>Stream</I> is not available in your server.</B><br>Check the Requirements page for information about upgrading your ADODB libraries."
        exit function
    end if
    Set streamTest = Nothing
end function

function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey

    Set Upload = New ASPUpload
    Upload.Save(uploadsDirVar)

	'If something fails inside the script, but the exception is handled
	If Err.Number<>0 then Exit function
	
	 SaveFiles = ""
    ks = Upload.UploadedFiles.keys
    if (UBound(ks) <> -1) then
        'SaveFiles = "<B>Files uploaded:</B> "
        for each fileKey in Upload.UploadedFiles.keys
            SaveFiles = Upload.UploadedFiles(fileKey).FileName
			namefile = Upload.UploadedFiles(fileKey).FileName
        next
    else
        SaveFiles = "No file selected for upload or the file name specified in the upload form does not correspond to a valid file in the system."
    end if
	fcCode = Upload.form("fcCode")
	fcfilename = Upload.form("fcfilename")
	uploaddocdate = Upload.form("uploaddocdate")
	remarks = Upload.form("Remarks")
end function

Function runSQL(SQL,rs)
'on error resume next
	dim myrs

	set myRs = createobject("ADODB.recordset")
    myRs.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & uploadsDirVar & ";Extended Properties=""dBASE IV;""", 1, 3

	set rs = myRs
		if err then
			runSQL = false
			response.write err.description
		else
			runSQL = true
		end if
end function

Function validation_fcba(value_)
	Dim initFcba
	Dim rsFcba, strFcba
	
	strFcba =	" SELECT fccode FROM businessunit WHERE Upper(fccode) LIKE Upper('%" & value_ & "%')"
	Set rsFcba = Server.CreateObject("ADODB.Recordset")
	rsFcba.Open strFcba, connstr, 3
	
	
	If Not rsFcba.EOF Then
		initFcba = rsFcba("fccode")
	Else
		initFcba = "Error, kode unit " & value_ & " tidak ada"
	End If
	
	rsFcba.Close
	set rsFcba = Nothing

	validation_fcba = initFcba
end function

Function validation_Job_Blok(fcba_,period_,blok_, job_, act_)
	Dim initBlok
	Dim rsBlok, strBlok
	
	strBlok =	" SELECT  a.fccode blkcode,a.status blkstatus,b.fccode jobcode,b.fcname jobname  " & _
				" FROM field_hist_tab a " & _
				" INNER JOIN actmaster b ON (a.status=b.status) " & _
				" WHERE fcba='" & fcba_ & "' AND YEAR=Extract(YEAR from to_date('" & period_ & "','MM/DD/YYYY')) AND MONTH=Extract(MONTH from to_date('" & period_ & "','MM/DD/YYYY')) AND a.fccode LIKE '" & blok_ & "%'" & _
				" AND  Upper(b.jobcode) = '" & job_ & "' " & _ 
				" AND  Upper(b.fcname)  = '" & act_ & "' "
	Set rsBlok = Server.CreateObject("ADODB.Recordset")
	rsBlok.Open strBlok, connstr, 3
	
	
	If Not rsBlok.EOF Then
		rsBlok.MoveLast
		rsBlok.MoveFirst
		If rsBlok.RecordCount > 1 Then
			initBlok = "Error, ditemukan lebih dari satu record untuk kode blok :" & blok_ & " / kode job :" & job_
		Else
			initBlok = rsBlok("jobcode")
		End If
	Else
		initBlok = "Error, kode.job :" & job_  & " dengan kode.blok :" & blok_ & " beda status tanam / tidak ada"
	End If
	
	rsBlok.Close
	set rsBlok = Nothing

	validation_Job_Blok = initBlok
end function

Function validation_kontrak(comp_,period_,kontraktor_, job_)
	Dim initktk
	Dim rsktk, strktk
	
	strktk =	" SELECT DISTINCT a.agreementcode FROM contractagreementdetail a " & _ 
				" INNER JOIN contractagreement b ON (a.agreementcode=b.agreementcode AND a.fcba=b.fcba ) " & _
				" WHERE " & _
				" a.fcba IN (SELECT fccode FROM businessunit WHERE fctype='HO' AND fccompanycode='" & comp_ & "') " & _
				" AND isclose='N' " & _
				" AND jobcode='" & job_ & "' " & _
				" AND contractorcode='" & kontraktor_ & "' " & _
				" AND To_Date('"& period_ & "','MM/DD/YYYY') BETWEEN b.startdate AND b.finishdate" 

	Set rsktk = Server.CreateObject("ADODB.Recordset")
	rsktk.Open strktk, connstr, 3
	
	If Not rsktk.EOF Then
		initktk = rsktk("agreementcode")
	Else
		initktk = "Error, nomor kontrak tidak ada"
	End If
	
	rsktk.Close
	set rsktk = Nothing

	validation_kontrak = initktk
end function
			

			

'UPLOAD HTML ---------------------------->

Dim diagnostics

if Request.ServerVariables("REQUEST_METHOD") <> "POST" then
    diagnostics = Environment()
    if diagnostics<>"" then
        response.write "<div style=""margin-left:20; margin-top:30; margin-right:30; margin-bottom:30;"">"
        response.write diagnostics
        response.write "<p>After you correct this problem, reload the page."
        response.write "</div>"
    else
        response.write "<div>"
        OutputForm()
        response.write "</div>"
    end if
else
    SaveFiles()

'HTML HEADER UPLOAD ------->
	OutputForm()
	
'PROCESS ------------------->
	Dim rsPO
	if trim(fcCode) <> "" then
		
		   Set rsPO = Server.CreateObject("ADODB.Recordset")
			   'connstr.BeginTrans
			   'response.write "NAFIS FCcode : " & fcCode &  " FCBA :" & Session("FCBA")
			   tempString = " select *" & _
							" from upload_wip_gis_tab" & _
							" where docno = '" & fcCode & "'" & _
							"       and fcba = '" & Session("FCBA") & "'"  
			   'response.write tempString
			   rsPO.Open  tempString, connstr	   
			if Not rsPO.Eof then
				temptxtJenis = ""
				cmdView	= "True"
				Response.Redirect "upload_wip_gis.asp?fcCode=" & fcCode & "&cmdView=" & cmdView
			else
				temptxtJenis = "New"				
			End if
		rsPO.Close
		
	end if	
	
		'response.write ("JENIS : " & temptxtJenis & "<br/> ")
		
		if temptxtJenis="New" or temptxtJenis = "Header" or temptxtJenis = "Details" or temptxtJenis ="Edit" or temptxtJenis = "Confirmed" or temptxtJenis= "RePlanned" Then
		'response.write ("ata 1")
		
			if temptxtJenis="New" then



			uploaddocdate = txtMonth & "/1/" & txtyear

'response.write uploaddocdate

			Dim  Msql, qinsert, myRs
			Dim temp_id,temp_pt ,temp_kebun,temp_divisi ,temp_blok ,temp_jobcode ,temp_actname ,temp_kontraktor,temp_tglwip,temp_luas 
			Dim rscompare, sqlcompare, temp_result, cek_ba_reasontype, temp_instring, temp_data, Cek_Validasi, Display_Remarks,cek_bu,cek_jb,Msg_bu,Msg_jb,cek_ktk,Msg_ktk,numError
			no = 1
			
			
' Set up the recordset object
set myRs = createobject("ADODB.recordset")

' SQL query to select data from the .dbf file
SQL = "SELECT * FROM [line.dbf]"

' Connection string to read .dbf file
uploadsDirVar = "c:\\temptemp\\"
namefile = "line.dbf"
myRs.Open SQL, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & uploadsDirVar & ";Extended Properties=""dBASE IV;""", 1, 3
' Output the table
temp_instring = ""
cek_bu        = "TRUE"
response.write "<table border='1'>"
' Output the header row
response.write "<tr>"
For I = 0 To myRs.Fields.Count - 1
    response.write "<th>" & myRs.Fields.Item(I).Name & "</th>"
Next
response.write "<th>" & "Message" & "</th>"
response.write "</tr>"
' Process the data
numError = 0
do while not myRs.EOF
    response.write "<tr>"
	 Msg_bu = ""
	 Msg_jb = ""
	 Msg_ktk = ""
		
	For I = 0 To myRs.Fields.Count - 1
	  
		' Cek KOde Lokasi Kerja
		IF I=2 then
			temp_instring = validation_fcba(myRs.Fields.Item(I).Value)
			if Left(temp_instring,5) ="Error" then
				cek_bu = "FALSE"
				Msg_bu = temp_instring
				numError = numError + 1
			else
				cek_bu = "TRUE"
			end if
			
		end if
		
		' Cek Kode Blok, Kode Job 
		'fcba_,period_,blok_, job_
		IF I=8 then
			temp_instring = validation_Job_Blok(myRs.Fields.Item(2).Value,myRs.Fields.Item(8).Value,myRs.Fields.Item(4).Value,myRs.Fields.Item(5).Value,myRs.Fields.Item(6).Value)
			if Left(temp_instring,5) ="Error" then
			cek_jb = "FALSE"
			msg_jb = temp_instring
			numError = numError + 1
			else
			cek_jb = "TRUE"
			end if
			
		end if
		
		' cek kontrak
		'comp_,period_,kontraktor_, job_
		IF I=9 then
			temp_instring = validation_kontrak(myRs.Fields.Item(1).Value,myRs.Fields.Item(8).Value,myRs.Fields.Item(7).Value,myRs.Fields.Item(5).Value)
			if Left(temp_instring,5) ="Error" then
			cek_ktk = "FALSE"
			msg_ktk = temp_instring
			numError = numError + 1
			else
			cek_ktk = "TRUE"
			end if
			
		end if
		

		response.write "<td>" & myRs.Fields.Item(I).Value & "</td>"
		
	
		
    Next
		if Msg_bu="" and Msg_jb="" and Msg_ktk="" then 
			response.write "<td><font color='Blue'>" & "Ok" & "</font></td>"
			else
			response.write "<td><font color='Red'>" & Msg_bu & Msg_jb  & msg_ktk & "</font></td>"
		end if
		
	
	response.write "</tr>"
	myRs.MoveNext
loop
response.write "</table>"

' Close the recordset
myRs.Close
set myRs = nothing


		   If numError = 0 then
				' Response.write "<font color=""green"">"
				' Response.write "Data Boleh Upload!"
				' Response.write "</font>"
				
				 Set rsAUTONUM = Server.CreateObject("ADODB.Recordset")
				 numq = " SELECT NVL(controlnumber_api.get_nextnumber( UPPER('UPLOAD_WIPGIS')" & _
					    "                                              , " & txtyear& "" & _
					    "                                              , " & txtMonth & "" & _
					    "                                              , 'ARJ'), 'FALSE') AUTONUMBER" & _
					    " FROM DUAL "
					   
				' 'response.write numq
				 rsAUTONUM.Open  numq, connstr	   
				 if Not rsAUTONUM.Eof then
					AutoNum = rsAUTONUM("AUTONUMBER")
				 end if
				 fcCode = AutoNum
			
				 q = " MERGE INTO upload_wip_gis_tab x USING ( " & _ 
				     " SELECT   '" & fcCode & "' docno" & _
					 "          , to_date('" & uploaddocdate & "', 'mm/dd/yyyy') tdate" & _
					 "          , '" & namefile & "' filename " & _
					 "          , '" & Session("CodeName") & "' fcentry" & _
					 "          , '" & Request.ServerVariables("REMOTE_ADDR") &"' fcip" & _
					 "          , '" & Session("fcBA") & "' fcba" & _
					 "          , 'Planned' rowstateid" & _
					 "          , sysdate lastupdate" & _
					 "          , to_char(sysdate,'hh24:mi') lasttime " & _ 
					 " FROM DUAL ) y " & _ 
					 " ON (x.docno = y.docno ) " & _ 
					 " WHEN MATCHED THEN " & _
					 " UPDATE SET x.tdate=y.tdate , x.fcentry=y.fcentry , x.lastupdate=y.lastupdate " & _ 
					 " WHEN NOT MATCHED THEN " & _
					 " INSERT (docno , tdate, filename, fcentry, fcip, fcba, rowstateid , lastupdate , lasttime) " & _
					 " values (y.docno , y.tdate, y.filename, y.fcentry, y.fcip, y.fcba, y.rowstateid , y.lastupdate , y.lasttime) "
					
                 'response.write q 					
				 connStr.execute(q)

				' Loop untuk qry save
				 if runsql(sql,rs) then
					Do While Not rs.EOF
						
						For I = 0 To rs.Fields.Count - 1
										temp_id         = rs.Fields.Item(0).Value
										temp_pt         = rs.Fields.Item(1).Value
										temp_kebun      = rs.Fields.Item(2).Value
										temp_divisi     = rs.Fields.Item(3).Value
										temp_blok       = rs.Fields.Item(4).Value
										temp_jobcode    = rs.Fields.Item(5).Value
										temp_actname    = rs.Fields.Item(6).Value
										temp_kontraktor = rs.Fields.Item(7).Value
										temp_tglwip = rs.Fields.Item(8).Value 
										temp_luas   = round(rs.Fields.Item(9).Value,2)
							
					Next
							
					q = ""
					q = 	" MERGE INTO upload_dtl_wip_gis_tab e " & _
									     " USING (SELECT '" & fcCode & "' docno, " & _
										 "                " & temp_id & " id, " & _ 
										 "               '" & temp_pt & "' pt," & _
										 "               '" & temp_kebun & "'kebun ," & _
										 "				 '" & temp_divisi & "' divisi,      " & _ 
										 "               '" & temp_blok & "' blok,      " & _ 
										 "               '" & temp_jobcode & "' job_code,      " & _ 
										 "               '" & temp_actname & "' actname,      " & _ 
										 "               '" & temp_kontraktor & "' kontraktor,      " & _ 
										 "		 to_date('" & temp_tglwip & "','mm/dd/yyyy') tglwip, " & _   
										 "               '" & temp_luas & "' luas,      " & _ 
										 "				Sysdate lastupdate,                  " & _ 
										 "				to_char(sysdate, 'hh24:mi') lasttime," & _
										 "				'" & Session("CodeName") & "' fcentry " & _
										 "              FROM DUAL  " & _
										 "        )h" & _
										 " ON (e.id=h.id and e.docno = h.docno and e.kebun=h.kebun and e.blok=h.blok and e.job_code=h.job_code and e.actname=h.actname and e.tanggal_wip=h.tglwip) " & _
										 " WHEN MATCHED THEN " & _
										 " UPDATE SET e.luas=h.luas ,e.lastupdate = h.lastupdate,e.fcentry=h.fcentry " & _
										 " WHEN NOT MATCHED THEN " & _
										 " INSERT (docno , id, pt, kebun, divisi, blok, job_code , actname , kontraktor , tanggal_wip , luas , lastupdate , lasttime , fcentry) " & _
										 " values (h.docno , h.id , h.pt, h.kebun , h.divisi , h.blok , h.job_code , h.actname , h.kontraktor , h.tglwip , h.luas , h.lastupdate , h.lasttime, h.fcentry) "
										
									 'response.write q & "<br>"
									 connStr.execute(q)	
									 no = no + 1	
					
					rs.MoveNext
					Loop
					rs.Close
				 End if
				 Set rs = Nothing
				 cmdView	= "True"
				
				 Response.Write("<script>alert('Upload Data Sukses');</script>")

				 Response.write "<font color=""green"">"
				 Response.write "Upload Data Sukses"
				 Response.write "</font>"
				
				 'Response.Redirect "upload_wip_gis.asp?fcCode=" & fcCode & "&cmdView=" & cmdView
				
			 Else
			     Response.Write("<script>alert('Data tidak bisa di Upload! Periksa kembali Kolom file anda');</script>")
				 Response.write "<font color=""red"">"
				 Response.write "Data tidak bisa di Upload! Periksa kembali Kolom file anda"
				 Response.write "</font>"
			 End If
			
			 cmdView	= "True"
		 else
		
		 cmdView	= request("cmdView")
		 end if
	 end if

  ' END PROCESS ------------------->

end if


 ' txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 

 ' Uploader -------------------------->
 

function OutputForm()%>

<form name="frmSend" accept-charset="utf-8" ENCTYPE="multipart/form-data" onSubmit="return onSubmitForm();"  name="frmEntry" method="post" action="upload_wip_gis_preview.asp?txtYear=<%=txtYear%>&txtMonth=<%=txtMonth%>">
	<table class="data" width="100%">
		<tr>
		<td align="left">No. Document <%=temptxtJenis%></td>
		<td align="left">
			<input type="text" maxlength="15" name="fcCode" size="30" value="<%=fcCode%>" readonly="True" style="background:#999999">
			<font color="#FF0000"> * (Diisi otomatis oleh sistem)</font></td>
		</tr>
		
		<tr>
			<td align="left">Nama File</td>
            <td align="left">
			<!--<input type="text" maxlength="50" name="fcfilename" size="30" value="<%=fcfilename%>">-->
			<INPUT TYPE=FILE NAME="txtFile">
			<font color="#FF0000">*</font>
			</td>
			<td align="left">  
			
			</td>
		</tr>
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=txtTable%>" Name="txtTable">
		<input Type="Hidden" Value="<%=txtModule%>" Name="txtModule">
		<input Type="Hidden" Value="<%=txtMonth%>" Name="txtMonth">
		<input Type="Hidden" Value="<%=txtYear%>" Name="txtYear">
		<input type="hidden" name="txtJenis" value="New">
		<input Type="Submit" name="Add" value="Upload" name="cmdSubmit">&nbsp;&nbsp;
		<a href="upload_wip_gis.asp?"><input type=button value="Back">	</a>
		</td>
		</tr>
    </table>  
 </form>
<%
end function
 ' END Uploader -------------------------->
 


'End HTML --------------------->


 '--- Action ke Database di Oracle [ insert,edit,delete ]
 ' if Request("txtJenis")="New" or Request("txtJenis")= "Header" or Request("txtJenis")= "Details" or Request("txtJenis")="Edit" or Request("txtJenis")= "Confirmed" or Request("txtJenis")= "RePlanned" Then
	' REMARKS=request("REMARKS")
	
	' if Request("txtJenis")="New" then
	   ' Dim rsPO
	   ' Set rsPO = Server.CreateObject("ADODB.Recordset")
	   ' 'connstr.BeginTrans
	  
	   ' rsPO.Open "Select uploaddoccode from purchaseorderheader_sap where uploaddoccode='" & fcCode & "' AND FCBA = '" & Session("FCBA") & "'" , connstr
	   ' IF rsPO.eof then	   
	
	     ' q="Insert into purchaseorderheader_sap " & _
		 ' " ( UPLOADDOCCODE, uploaddocdate, filename, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, ROWSTATEID, LASTUPDATE, LASTTIME ) " & _
		 ' "values ('" & fcCode & "'," & " to_date('" & uploaddocdate & "','mm/dd/yyyy') " & ",'"  &  fcfilename & _
		 ' "','" & Remarks & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
		 ' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','Planned', sysdate, To_Char(sysdate,'HH24:MI'))"
		 ' response.write q
	     ' connStr.execute(q)
	   ' End if
	   ' rsPO.Close
	   ' response.write "Test123" 
	   ' IF err.number = 0 then	
	      ' response.write "Test456" 
	      ' Dim  Msql
		  ' set o = new ASPUpload
		  ' response.write "Test789" 
		  
		  ' ' sFileSplit = split(o.FileNameOf("txtFile"), "\")
		  
		  ' ' sFile = sFileSplit(Ubound(sFileSplit))
          ' ' o.FileInputName = "txtFile"
		  ' ' o.FileFullPath = Server.MapPath(".") & "\" & sFile
		 
		  ' sql = "SELECT * FROM [Sheet1$];"

		    ' IF runsql(sql,rs) then
				' Do While Not rs.EOF
					' Msql = ""
					
					' IF rs.Fields.Item(4).Value <> "" then
						' For I = 0 To rs.Fields.Count - 1
							' Msql = Msql & rs.Fields.Item(I).Value	& "','" 				   
						' Next
						' Msql = "'" & mid(Msql,1,len(Msql)-3) & "'" 				 
					' END IF	
					 
					 ' q= "Insert into  purchaseorderdetail_sap (UPLOADDOCCODE,PURCHASE_DOC,ITEM,VENDOR,MATL_GROUP,MATERIAL,PURCHASE_ORGANIZATION,PUCHASE_GROUP,PLANT,SLOC,DOC_DATE,ORDER_QUANTITY," & _
						' "ORDER_UNIT,NET_ORDER_VALUE,CCY,CONDITION_TYPE,AMOUNT_COND_TYPE,TRANSPORTER,PR_NO,PR_ITEM,PO_RELEASE,ERROR_FLAG,ERROR_DESC," & _
						' "FCENTRY,FCEDIT,FCIP,FCBA,LASTUPDATE,LASTTIME) " & _
						' "values ('" & fcCode & "'," & Msql & ",'FALSE','','','','" & _
						' Session("CodeName")  & "','" & Session("CodeName") & "','" & _
						' Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
					    ' response.write q
						' 'connStr.execute(q)	
						
					 ' rs.MoveNext
				' Loop
				' rs.Close
		    ' End if
		  ' Set rs = Nothing
		  ' If o.Error = "" then
		     ' response.write "Success" & o.FileFullPath & ""
		  ' Else
			 ' response.write "Failed due to the following error: " & o.Error
		  ' End if
		  ' set o = nothing 
          ' cmdView	= "True"
	   ' End if
    ' else
       ' cmdView	= request("cmdView")
    ' End if
' End if
'Dim rsPurchase
'set rsPurchase = server.CreateObject("Adodb.recordset")
'q = " select fcbatobe" & _
'	"        , empnotobe" & _
'	"        , tdate" & _
'	"        , timein" & _
'	"        , timeout" & _
'	"        , ba_reference" & _
'	" from userfinger.attendance_tab" & _
'	" where no_ba is not null" & _
'	"       and no_ba = '" & fcCode & "'" & _
'	"       and fcbatobe = '" & session("fcBA") & "'" 
'rsPurchase.Open q, connstr
'if Not rsPurchase.Eof then
'	uploaddocdate = rsPurchase("uploaddocdate")
'End if
'if err.number <> 0 then
'	connStr.RollbackTrans
'    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
'End if
 '-------------ISI Payment Form

'----akhir----------------------
%>
</div></div>
  <!--#include file="bottom.jsp"-->
 
</body>
</html>


 