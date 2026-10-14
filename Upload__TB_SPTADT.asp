<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->


<html>
<head>
<!--#include file="aspupload.asp"-->
<title>Upload Manual TB SPT&AUDIT</title>
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
 Dim txtYear,txtMonth,txtFcba

 Dim purchase_doc,item,vendor,matl_group,material,purchase_organization,puchase_group,plant,sloc,doc_date,order_quantity,order_unit
 Dim net_order_value,ccy,condition_type,amount_cond_type,transporter,pr_no,pr_item,po_release,error_flag,error_desc  
 Dim sql,namefile,temptxtJenis,fcCodeNew

 fcCode = "(NEW)"
 fcfilename = ""
 txtFcba  = Request("txtFCBA")
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
 
 uploadsDirVar = Server.MapPath(".") & "\tempUploads\" 
  

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
							" from userfinger.uploadheader_mf_tab" & _
							" where docno = '" & fcCode & "'" & _
							"       and fcba = '" & session("fcBA") & "'"  
			   'response.write tempString
			   rsPO.Open  tempString, connstr	   
			if Not rsPO.Eof then
				temptxtJenis = ""
				cmdView	= "True"
				'Response.Redirect "upload_ba_fp_01.asp?fcCode=" & fcCode & "&cmdView=" & cmdView
			else
				temptxtJenis = "New"				
			End if
		rsPO.Close
		
	end if	
	
		'response.write ("JENIS : " & temptxtJenis & "<br/> ")
		
		if temptxtJenis="New" or temptxtJenis = "Header" or temptxtJenis = "Details" or temptxtJenis ="Edit" or temptxtJenis = "Confirmed" or temptxtJenis= "RePlanned" Then
		'response.write ("ata 1")
		
			if temptxtJenis="New" then

Function runSQL(SQL,rs)
'on error resume next
	dim myrs

	set myRs = createobject("ADODB.recordset")
	myRs.Open SQL,"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & uploadsDirVar & namefile & ";Extended Properties=""Excel 12.0;IMEX=1;""", 1, 3
	set rs = myRs
		if err then
			runSQL = false
			response.write err.description
		else
			runSQL = true
		end if
end function
			
Function validation_ba_reasontype(initialvalue)
	Dim initialResult
	Dim rsValidation, strValidation
	
	strValidation =	" select count(1) total from dual where instr('"&initialvalue&"', 'FALSE') <> 0"

	Set rsValidation = Server.CreateObject("ADODB.Recordset")
	rsValidation.Open strValidation, connstr, 3
	
	If Not rsValidation.EOF Then
		initialResult = Cdbl(rsValidation("total"))
	Else
		initialResult = 1
	End If
	
	strValidation =	""
	rsValidation.Close
	set rsValidation = Nothing

	validation_ba_reasontype = initialResult
	
	' Return 1 Jika dari semua string terdapat kata-kata FALSE
	' Return 0 Jika lolos validasi
	
End Function

 uploaddocdate = txtMonth & "/1/" & txtyear

'response.write uploaddocdate

			Dim  Msql, qinsert, temp_tdate, temp_empnotobe, temp_timein, temp_timeout,temp_ba ,temp_tdateout
			Dim rscompare, sqlcompare, temp_result, cek_ba_reasontype, temp_instring, temp_data, Cek_Validasi, Display_Remarks
			no = 1
			sql = "SELECT * FROM [sheet1$];"
			
			if runsql(sql,rs) then
				temp_instring = ""
				cek_ba_reasontype = "TRUE"
				Do While Not rs.EOF
					Msql = ""
					temp_ba = ""
					For I = 0 To rs.Fields.Count - 1
						if I=0 then
						temp_ba = rs.Fields.Item(0).Value
						
							if trim(temp_ba) = "" or isNULL(temp_ba) then 
								cek_ba_reasontype = "FALSE"
							end if
						
						end if 
						
						
						
						if trim(temp_ba) <> "" then
							If I = 5 then
								temp_data = ""						
								temp_data = rs.Fields.Item(I).Value
									If NOT ( temp_data = "1" or temp_data = "2" or temp_data = "3" or temp_data = "4" or temp_data = "5" or temp_data = "6")then
										cek_ba_reasontype = "FALSE"
									End If						
							End If
						End If
					Next
					temp_instring = temp_instring & cek_ba_reasontype
					rs.MoveNext
					no=no+1
				Loop
				rs.Close
			End if
			Set rs = Nothing
			cmdView	= "True"
			
			'response.write temp_instring
			
			'Cek_Validasi = 0'
			Cek_Validasi = validation_ba_reasontype(temp_instring)
			
			If Cek_Validasi = 0 then
				Response.write "<font color=""green"">"
				Response.write "Data Boleh Upload!"
				Response.write "</font>"
				
				Set rsAUTONUM = Server.CreateObject("ADODB.Recordset")
				numq = " SELECT NVL(controlnumber_api.get_nextnumber( UPPER('UPLOAD_MF')" & _
					   "                                              , " & txtyear& "" & _
					   "                                              , " & txtMonth & "" & _
					   "                                              , '" & Session("FCBA") & "'), 'FALSE') AUTONUMBER" & _
					   " FROM DUAL "
					   
				'response.write numq
				rsAUTONUM.Open  numq, connstr	   
				if Not rsAUTONUM.Eof then
					AutoNum = rsAUTONUM("AUTONUMBER")
				end if
				fcCode = AutoNum
			
				q = " insert into userfinger.uploadheader_mf_tab( docno" & _
					"                                             , tdate" & _
					"                                             , filename" & _
					"                                             , fcentry" & _
					"                                             , fcip" & _
					"                                             , fcba" & _
					"                                             , rowstateid" & _
					"                                             , lastupdate" & _
					"                                             , lasttime)" & _
					" values ( '" & fcCode & "'" & _
					"          , to_date('" & uploaddocdate & "', 'mm/dd/yyyy')" & _
					"          , '" &  namefile & "'" & _
					"          , '" & Session("CodeName") & "'" & _
					"          , '" & Request.ServerVariables("REMOTE_ADDR") &"'" & _
					"          , '" & Session("fcBA") & "'" & _
					"          , 'Planned'" & _
					"          , sysdate" & _
					"          , to_char(sysdate,'hh24:mi'))"
				'response.write q
				connStr.execute(q)

				
				if runsql(sql,rs) then
					
					Do While Not rs.EOF
							Msql = ""
							'response.write ">" & rs.Fields.Item(0).Value & "<" & ">" & rs.Fields.Item(1).Value & "<" & ">" & rs.Fields.Item(2).Value & "<" & ">" & rs.Fields.Item(3).Value & "<" & ">" & rs.Fields.Item(4).Value & "<" & ">" & rs.Fields.Item(5).Value & "<" & ">" & rs.Fields.Item(6).Value & "<" & "</br>"
						
								IF rs.Fields.Item(0).Value <> "" and rs.Fields.Item(1).Value <> "" then
									For I = 0 To rs.Fields.Count - 1
										If I = 1 then
											'qinsert = ", '" & rs.Fields.Item(I).Value & "'"
											temp_empnotobe = rs.Fields.Item(I).Value
										elseif I = 2 then
											'qinsert = ", to_date('"& rs.Fields.Item(I).Value & "', 'mm/dd/yyyy')"
											temp_tdate = rs.Fields.Item(I).Value
											temp_tdateout = rs.Fields.Item(6).Value 
										elseif I = 3 then
											'qinsert = ", to_date('"& temp_tdate & " " & rs.Fields.Item(I).Value & ":00', 'mm/dd/yyyy hh24:mi:ss')"
											temp_timein = " to_date('"& temp_tdate & " " & rs.Fields.Item(I).Value & ":00', 'mm/dd/yyyy hh24:mi:ss') "
										elseif I = 4 then
											temp_timeout = " to_date('"& temp_tdateout & " " & rs.Fields.Item(I).Value & ":00', 'mm/dd/yyyy hh24:mi:ss') "
											
											set rscompare = server.CreateObject("Adodb.recordset")
											sqlcompare = " select 'TRUE' Result" & _
														 " from dual" & _
														 " where " & temp_timeout & " < " & temp_timein
											'response.write sqlcompare
											rscompare.open sqlcompare, connstr
											if not rscompare.EOF then
												temp_result = rscompare("Result")
											else
												temp_result = "FALSE"
											end if
											
											if temp_result = "TRUE" then
											    temp_timeout = " to_date('"& temp_tdate & " " & rs.Fields.Item(I).Value & ":00', 'mm/dd/yyyy hh24:mi:ss') + 1"
												'temp_timeout = ", to_date('"& temp_tdate & " " & rs.Fields.Item(I).Value & ":00', 'mm/dd/yyyy hh24:mi:ss') + 1"
											else
											    temp_timeout = " to_date('"& temp_tdateout & " " & rs.Fields.Item(I).Value & ":00', 'mm/dd/yyyy hh24:mi:ss') "
												'qinsert = ", to_date('"& temp_tdate & " " & rs.Fields.Item(I).Value & ":00', 'mm/dd/yyyy hh24:mi:ss')"
											end if
											' qinsert = ", to_date('"& temp_tdate & " " & rs.Fields.Item(I).Value & ":00', 'mm/dd/yyyy hh24:mi:ss')"
										'else
										'	qinsert = ", '" & rs.Fields.Item(I).Value & "'"
										end if
										'Msql = Msql & qinsert
									Next
									'Msql = "'" & mid(Msql,1,len(Msql)-3) & "'"
								
								 q = ""
								 q = 	" MERGE INTO userfinger.uploaddetail_mf_tab e " & _
										" USING (SELECT '" & fcCode & "' docno," & _
										"               '" & temp_ba & "'no_ba ," & _
										"				'" & temp_empnotobe & "' empno,      " & _  
										"				to_date('" & temp_tdate & "','mm/dd/yyyy') tdate, " & _   
										"				" & temp_timein  & " timein,         " & _
										"				" & temp_timeout & " timeout,        " & _
										"               '' ba_reasontype,                    " & _
										"				'" &Session("fcBA") & "' fcba,       " & _ 
										"				Sysdate lastupdate,                  " & _ 
										"				to_char(sysdate, 'hh24:mi') lasttime," & _
										"				NULL error_flag,                     " & _
										"				NULL error_desc                      " & _
										"               FROM DUAL  " & _
										"        )h" & _
										" ON (e.empno = h.empno and e.tdate=h.tdate) " & _
										" WHEN MATCHED THEN " & _
										" UPDATE SET e.docno=h.docno ,e.no_ba=h.no_ba,e.error_flag=h.error_flag, e.error_desc=h.error_desc,e.lastupdate = h.lastupdate " & _
										" WHEN NOT MATCHED THEN " & _
										" INSERT (docno , no_ba , empno , tdate , timein , timeout , ba_reasontype , fcba , lastupdate , lasttime , error_flag , error_desc) " & _
										" values (h.docno , h.no_ba , h.empno , h.tdate , h.timein , h.timeout , h.ba_reasontype , h.fcba , h.lastupdate , h.lasttime , h.error_flag , h.error_desc) "
										
										'response.write q & "<br>"
										connStr.execute(q)	
									no = no + 1
								END IF		
								rs.MoveNext
							Loop
							rs.Close
				End if
				Set rs = Nothing
				cmdView	= "True"
				
				Response.Redirect "upload_ba_fp_01.asp?fcCode=" & fcCode & "&cmdView=" & cmdView
				
			Else
				Response.write "<font color=""red"">"
				Response.write "Data tidak bisa di Upload! Periksa kembali Kolom file anda, Pastikan semua kolom terisi , untuk ba_resontype hanya berisikan 1/2/3/4/5/6"
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

<form name="frmSend" accept-charset="utf-8" ENCTYPE="multipart/form-data" onSubmit="return onSubmitForm();"  name="frmEntry" method="post" action="upload_ba_fp_02.asp?txtYear=<%=txtYear%>&txtMonth=<%=txtMonth%>">
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
		<input Type="Submit"  name="Add" value="Select" name="cmdSubmit">		
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


 