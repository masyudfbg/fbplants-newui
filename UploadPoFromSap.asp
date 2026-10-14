<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="adovbs.inc"-->


<html>
<head>
<!--#include file="aspupload.asp"-->
<title>ePlaS - Upload Purchase Order</title>
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

 Dim purchase_doc,item,vendor,matl_group,material,purchase_organization,puchase_group,plant,sloc,doc_date,order_quantity,order_unit
 Dim net_order_value,ccy,condition_type,amount_cond_type,transporter,pr_no,pr_item,po_release,error_flag,error_desc  
 Dim sql,namefile,temptxtJenis,fcCodeNew
 fcCode = "(NEW)"
 fcfilename = ""
 txtTable =Request("txtTable")
 txtModule=Request("txtModule")
 txtSearch=Request("txtSearch")
 txtField =Request("txtField")
 temptxtJenis = Request("txtJenis")  
  xCode = Request("xCode")
 Remarks = Request("Remarks")
 purchase_doc 			= request("purchase_doc")
 Item     				= request("Item")
 vendor       			= request("vendor")
 matl_group   			= request("matl_group")
 material     			= request("material")
 purchase_organization  = request("purchase_organization")
 puchase_group          = request("puchase_group")
 plant                  = request("plant")
 sloc                   = request("sloc")
 doc_date               = request("doc_date")
 order_quantity         = request("order_quantity")
 order_unit             = request("order_unit")
 net_order_value		= request("net_order_value")
 ccy					= request("ccy")
 condition_type			= request("condition_type")
 amount_cond_type		= request("amount_cond_type")
 transporter			= request("transporter")
 pr_no					= request("pr_no")
 pr_item				= request("pr_item")
 po_release				= request("po_release")
 error_flag				= request("error_flag")
 error_desc 			= request("error_desc")
 
 dim rsAUTONUM,AutoNum,numq,txtdateNow

 Set rsAUTONUM = Server.CreateObject("ADODB.Recordset")
 numq = "select nvl(controlnumber_api.get_nextnumber(upper('UPLOADPO'), " & year(date) & ", " & month(date) & ", '" & Session("FCBA") & "'), 'FALSE') prcode from dual "
 rsAUTONUM.Open  numq, connstr	   
			if Not rsAUTONUM.Eof then
				AutoNum = rsAUTONUM("prcode")
			end if
 fcCodeNew				= AutoNum
 '---- Data Details
 %>

 <%
 ' Add by NAF, 2012-12-1
 LINENOKEY=Request("LineNoKey")
 ' End Added by NAF

 ITEMCODE=Request("ITEMCODE")
 QUANTITY=Request("QUANTITY")
 AMOUNT =Request("AMOUNT")
 OTHERITEMDESC=request("OTHERITEMDESC")
 RowStateId = Request("RowStateId") 
 PRCODE = Request("PRCODE") 
 PRLINEID = Request("PRLINEID") 
 uploaddocdate = Request("uploaddocdate") 
  	'---- Data Header 
	if uploaddocdate  = "" then
	uploaddocdate   = textDate(date)  
	 else
	uploaddocdate   = uploaddocdate 
	end if


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
    response.write SaveFiles()

'HTML HEADER UPLOAD ------->
	OutputForm()
	
'PROCESS ------------------->
	Dim rsPO
	if trim(fcCode) <> "(NEW)" then
		
		   Set rsPO = Server.CreateObject("ADODB.Recordset")
			   'connstr.BeginTrans
			   'response.write "NAFIS FCcode : " & fcCode &  " FCBA :" & Session("FCBA")
			   tempString = "Select uploaddoccode from purchaseorderheader_sap where uploaddoccode='" & fcCode & "' AND FCBA = '" & Session("FCBA") & "'"   
			   'response.write tempString
			   rsPO.Open  tempString, connstr	   
			if Not rsPO.Eof then
				temptxtJenis = ""
				cmdView	= "True"
				Response.Redirect "POUpload.asp?fcCode=" & fcCode & "&cmdView=" & cmdView
				
			else
				temptxtJenis = "New"
				
			End if
		rsPO.Close
	else
	
		 Set rsPO = Server.CreateObject("ADODB.Recordset")
			   'connstr.BeginTrans
			  
			   tempString = "Select uploaddoccode from purchaseorderheader_sap where uploaddoccode='" & fcCodeNew & "' AND FCBA = '" & Session("FCBA") & "'"   
			   'response.write tempString
			   rsPO.Open  tempString, connstr	   
			if Not rsPO.Eof then
				temptxtJenis = ""
				cmdView	= "True"
				Response.Redirect "POUpload.asp?fcCode=" & fcCodeNew & "&cmdView=" & cmdView
				
			else
				temptxtJenis = "New"
				
			End if
		rsPO.Close

		fcCode = fcCodeNew
		
		
	end if	
	
		'response.write ("JENIS : " & temptxtJenis & "<br/> ")
		
		if temptxtJenis="New" or temptxtJenis = "Header" or temptxtJenis = "Details" or temptxtJenis ="Edit" or temptxtJenis = "Confirmed" or temptxtJenis= "RePlanned" Then
		'response.write ("ata 1")
		
			if temptxtJenis="New" then
		
			 q="Insert into purchaseorderheader_sap " & _
			 " ( UPLOADDOCCODE, uploaddocdate, filename, REMARKS, FCENTRY, FCEDIT, FCIP, FCBA, ROWSTATEID, LASTUPDATE, LASTTIME ) " & _
			 "values ('" & fcCode & "'," & " to_date('" & uploaddocdate & "','mm/dd/yyyy') " & ",'"  &  namefile & _
			 "','" & remarks & "','" & Session("CodeName") & "','" & Session("CodeName") & "','" & _
			 Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "','Planned', sysdate, To_Char(sysdate,'HH24:MI'))"
			 'response.write q
			 connStr.execute(q)

		   function runSQL(SQL,rs)

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

			Dim  Msql
			no = 1
			sql = "SELECT * FROM [sheet1$];"
				if runsql(sql,rs) then
				
				Do While Not rs.EOF
						Msql = ""
						
							IF rs.Fields.Item(0).Value <> "" and rs.Fields.Item(1).Value <> "" then
							'response.write "atwal 3"	
								For I = 0 To rs.Fields.Count - 1
									Msql = Msql & rs.Fields.Item(I).Value	& "','" 				   
								Next
								Msql = "'" & mid(Msql,1,len(Msql)-3) & "'" 				 
							
							 q=""
							 q= "Insert into  purchaseorderdetail_sap (UPLOADDOCCODE,Lineno,PURCHASE_DOC,ITEM,VENDOR,MATL_GROUP,MATERIAL,PURCHASE_ORGANIZATION,PUCHASE_GROUP,PLANT,SLOC,DOC_DATE,ORDER_QUANTITY," & _
								"ORDER_UNIT,NET_ORDER_VALUE,CCY,CONDITION_TYPE,AMOUNT_COND_TYPE,TRANSPORTER,PR_NO,PR_ITEM,PO_RELEASE,ERROR_FLAG,ERROR_DESC," & _
								"FCENTRY,FCEDIT,FCIP,FCBA,LASTUPDATE,LASTTIME) " & _
								"values ('" & fcCode & "'," & no & "," & Msql & ",'','','" & _
								Session("CodeName")  & "','" & Session("CodeName") & "','" & _
								Request.ServerVariables("REMOTE_ADDR") &"','" & Session("fcBA") & "', sysdate, To_Char(sysdate,'HH24:MI'))"
								response.write q
								
								connStr.execute(q)	
								no = no + 1
							END IF		
							 rs.MoveNext
						Loop
					rs.Close
				End if
				Set rs = Nothing
				cmdView	= "True"
				
				Response.Redirect "POUpload.asp?fcCode=" & fcCode & "&cmdView=" & cmdView
		else
		
		cmdView	= request("cmdView")
		end if
	end if

  ' END PROCESS ------------------->

end if


 ' txtJenis = request("txtJenis") '---- Untuk Membedakan Jenis Data Yang akan diolah [insert,edit,delete]
 

 ' Uploader -------------------------->
 

function OutputForm()%>

<form name="frmSend" accept-charset="utf-8" ENCTYPE="multipart/form-data" onSubmit="return onSubmitForm();"  name="frmEntry" method="post" action="UploadPoFromSap.asp">
	<table class="data" width="100%">
		<tr>
		<td align="left">Upload Doc Code</td>
		<td align="left"><input type="text" maxlength="50" name="fcCode" size="30" value="<%=fcCode%>"><font color="#FF0000">*</font></td>
		</tr>
		
		<tr>
		<td align="left">Upload Doc Date</td>
		<td align="left"><input class="datepicker" type="text" maxlength="10" name="uploaddocdate" size="10" value="<%=uploaddocdate%>">
		</td>
		</tr>
		
		<tr>
			<td align="left">File Name </td>
            <td align="left">
			<!--<input type="text" maxlength="50" name="fcfilename" size="30" value="<%=fcfilename%>">-->
			<INPUT TYPE=FILE NAME="txtFile">
			<font color="#FF0000">*</font>
			</td>
			<td align="left">  
			
			</td>
			<tr>
			<td align="left">Remarks</td>
			<td align="left"><textarea  name="Remarks"  cols="50" rows="3"><%=Remarks%></textarea></td>		
			</tr>
		</tr>
		
		
		<tr>
		<td align="left">Status</td>
		<td align="left"><input disabled type="text" maxlength="50" name="RowStateId" size="30" value="<%=RowStateId%>"><font color="#FF0000">*</font>
		</td>
		</tr>
		
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=txtTable%>" Name="txtTable">
		<input Type="Hidden" Value="<%=txtModule%>" Name="txtModule">
		<input type="hidden" name="txtJenis" value="New">
		<input Type="Submit"  name="Add" value="Select" name="cmdSubmit">&nbsp;&nbsp;
	    <a href="javascript:launchwin('PrintPO.Asp?txtTable=<%=txtTable%>&txtModule=<%=txtModule%>&fcCode=<%=fcCode%>&uploaddocdate=<%=uploaddocdate%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')"><img src="IMAGES/Print.gif" border="0" title="Print Preview"><font face="Verdana" size="1">Print</font></a>&nbsp;&nbsp;
		<a href="UploadPoFromSap.asp?fcCode=<%=fcCode%>&txtJenis=Header"><img src="IMAGES/delete.gif" alt="Delete Data : <%=fccode%>" border="0" onClick="return confirm('Are you sure to delete');"></a>&nbsp;&nbsp;
		<% if RowStateId = "Planned" then %>
		    <a href="UploadPoFromSap.asp?fcCode=<%=fcCode%>&txtJenis=Confirmed" onClick="return confirm('Are you sure to Confirm this data :  <%=fcCode%> ?');"> Confirmed </a>
		<% elseif RowStateId = "Confirmed" then %>	
		    <a href="UploadPoFromSap.asp?fcCode=<%=fcCode%>&txtJenis=RePlanned" onClick="return confirm('Are you sure to RePlan this data :  <%=fcCode%> ?');"> Re-Planned </a>
		<% end if%>
		
		
		
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
 Dim rsPurchase
 set rsPurchase = server.CreateObject("Adodb.recordset")
 q="select * from purchaseorderheader_sap  " & _
  " where uploaddoccode='" & fcCode & "' and " & _
  " fcBA='" & session("fcBA") & "'"
 rsPurchase.Open q, connstr
 if Not rsPurchase.Eof then
    uploaddocdate=rsPurchase("uploaddocdate")
	xCode=rsPurchase("filename")
	Remarks=rsPurchase("Remarks")
	RowStateId = rsPurchase("RowStateId")
 End if
if err.number <> 0 then
	connStr.RollbackTrans
    response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" &  err.description & "</font>"
End if
 '-------------ISI Payment Form

'----akhir----------------------
%>
</div></div>
  <!--#include file="bottom.jsp"-->
 
</body>
</html>


 