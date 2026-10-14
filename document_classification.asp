<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>ePlaS - Document Classification</title>
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

function CheckIsActive()
{
 var cIsActive = document.getElementById('ISACTIVE');
 if (cIsActive.checked != true )
	{
		alert("Is Active Belum di Click");
	}
}

function calcParam()
{
	var cIsActive = document.getElementById('ISACTIVE').checked;
	if(cIsActive)
	{
		document.getElementById("hide_ISACTIVE").value ="1";
	}else
	{
		document.getElementById("hide_ISACTIVE").value ="0";
	}
	
	var cIsBlock = document.getElementById('ISBLOCK').checked;
	if(cIsBlock)
	{
		document.getElementById("hide_ISBLOCK").value ="1";
	}else
	{
		document.getElementById("hide_ISBLOCK").value ="0";
	}
		
}

</script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy"<% end if %>>
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

Dim FCENTRY, FCEDIT, FCIP, FCBA
Dim CUSTOMERCODE, txtDESCRIPTION, SUPPLIER, ICNO, BANK, BANKACCOUNTNO, COMPANYNO,OWN_ESTATE,CUSTOMER_FCBA,CUSTFFBTYPE,ISTRANSPORTER
Dim cmdView,txt_CUSTOMERCODE
Dim vSTREET, vCITY, vPOSTALCODE, vPROVINCE, vTELEPHONE, vFAX, vEMAIL, vISACTIVE, vISBLOCK, vISTRANSPORTER
Dim hide_ISACTIVE, hide_ISBLOCK, hide_ISTRANSPORTER
txttable=request("txttable")
fccode =request("fccode")

if Request("txtJenis")="New" or Request("txtJenis")= "Delete" or Request("txtJenis")="Edit" Then
 
	txt_CUSTOMERCODE=request("txt_CUSTOMERCODE")
	CUSTOMERCODE=request("CUSTOMERCODE")
	txtDESCRIPTION=request("txtDESCRIPTION")
	''SUPPLIER=request("SUPPLIER")
	ICNO=request("ICNO")
	BANK=request("BANK")
	BANKACCOUNTNO=request("BANKACCOUNTNO")
	COMPANYNO=request("COMPANYNO")

	SUPPLIER = "N" 'Default value
	
	CUSTOMER_FCBA =request("CUSTOMER_FCBA")
	vSTREET =request("STREET")
	vCITY =request("CITY")
	vPOSTALCODE =request("POSTALCODE")
	vPROVINCE =request("PROVINCE")
	vTELEPHONE =request("TELEPHONE")
	vFAX =request("FAX")
	vEMAIL =request("EMAIL")
	
	hide_ISACTIVE = 0
	hide_ISBLOCK = 0
		
	hide_ISACTIVE = request("hide_ISACTIVE") 
	hide_ISBLOCK = request("hide_ISBLOCK") 
	
	vISACTIVE =request("ISACTIVE")
	vISBLOCK =request("ISBLOCK")

	fccode =request("fccode")
	fcname =request("fcname")
	v_type =request("v_type")

	if Request("txtJenis")= "Delete" then
		q	=	"DELETE FROM document_classification WHERE  fcCode='" & fccode & "'"
		connStr.execute(q)
		cmdView	= "True"
	   
	elseif Request("txtJenis")="New" then
		q	=	" INSERT INTO document_classification (" & _
				" fccode, " & _
				" fcname, " & _
				" document_type " & _
				" ) values (" & _
				"   '" & fccode & "'" & _			
				" , '" & fcname & "'" & _
				" , '" & v_type & "'" & _	
				" )"			 

		''response.write q
		'response.Flush()
		connStr.execute(q)
		cmdView	= "True"

	elseif Request("txtJenis")="Edit" and Request("Add")="Update" then
       q=	"  UPDATE document_classification SET" & _ 
			"  fcname ='" & fcname & "'" & _
			" , document_type ='" & v_type & "'" & _
			"  WHERE  fccode ='" & fcCode &"'"  
		connStr.execute(q)
		connStr.execute("commit")
		''response.Write(q)
		'response.Flush()
		''response.write q			
        cmdView	= "True"
    else
        cmdView	= "True"
	end if								 
	'Response.Write "sql " & q
else
    cmdView	= request("cmdView")
End if

Dim txtModule,txtTable,txtSearch,txtField,i,q,nCountofField,NO,txtJenis

txtTable =Request("txtTable")
txtModule=Request("txtModule")
txtSearch=Request("txtSearch")
txtField =Request("txtField")
'fcCode=request("fcCode")
txtJenis = request("txtJenis")

If Err.Number <> 0 then
   Response.Write "<font color=red>" & "Error Number : " & Err.Number &  " Desc. " & Err.Description & "</font>"
End if

 '-------------ISI Payment Form
  %>
  
<form name="frmEntry" method="post" action="document_classification.asp">
	<table class="data" width="94%" cellpadding="0" cellspacing="0">
          <TR> 
            <td colspan="2"><div align="left"><font size="2" face="Verdana"><strong>MAINTENANCE DOCUMENT CLASSIFICATION</strong></font></div></td>
          </TR>
		<tr>
			<td align="left" width="15%">Code</td><td align="left"><input type="text" size="15" name="fcCode" value="<%=request("fcCode")%>"  ></td>
		</tr>	
		<tr><td colspan="2"  align="left">
		<input Type="Hidden" Value="True" Name="cmdView">
		<input Type="Hidden" Value="<%=request("txtTable")%>" Name="txtTable">
		<input Type="Hidden" Value="<%=request("txtModule")%>" Name="txtModule">
		<input Type="Submit"  name="Add" value="Select">&nbsp;
		<a href="Printdocument_classification.asp?txtTable=<%=request("txtTable")%>&txtModule=<%=request("txtModule")%>&fcCode=<%=REQUEST("fcCode")%>" 
		title="Print Data" target="_blank"><img src="IMAGES/Print.gif" border="0"></a>		
		</td>
		</tr>
    </table>  
 </form>
 <%
 if cmdView = "True" Then
   Dim rsGAD,intRecordsPerPage,cColor
   ' set rsGAD = server.CreateObject("Adodb.recordset")
    'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
	
	''SUPPLIER,OWN_ESTATE, CUSTFFBTYPE, ISTRANSPORTER,
	'' Query Menampilkan Data di Table Bawah  
	q	=	" SELECT " & _
			" fccode, fcname, document_type " & _
			" FROM document_classification" & _
			" WHERE fccode like '" & fccode & "%'" 	
		
    Set rsGAD = Server.CreateObject("ADODB.Recordset")
	' if  len(trim(request("CUSTOMERCODE"))) > 0 then
		' q = q & " and CUSTOMERCODE='" & request("CUSTOMERCODE") & "'"
	' end if
	' q = q
   
   'response.Write(q)
   'response.Flush()
   rsGAD.CursorType = 3
   rsGAD.CursorLocation = 3
   rsGAD.Open q, connstr 
   
  %>
   
   <!-- Form Entry And Edit --> 
   
   <form action="document_classification.asp" method="post" name="eForm">
        <table class="data" cellpadding="0" cellspacing="0" width="100%" Border="0">
		
		<!--Garis Hijau /  Blok Greeen -->
		<tr bgcolor="#00a300"> 
			<td colspan="8">
				<div align="left">
					<font size="2" face="Verdana" color="#fff"><strong>Document Classification</strong></font>
				</div>
			</td>
		</tr>
		
		<tr>
			<%IF Request("txtJenis")="Edit" then%>
				<td width="13%" align="left">Code</td>
				<td width="64%" align="left"><input type="text" size="15" maxlength="15" name="fccode" value="<%=fccode%>" readonly="True" style="background:#999999"><font color="#FF0000">*</font></td>
			<%else%>
				<td width="2%" align="left">Code</td>
				<td width="11%" align="left"><input type="text" size="15" maxlength="15" name="fccode" value="<%=fccode%>"  ><font color="#FF0000">*</font></td>
			<%end if%>
		</tr>
		<tr>	
			<td align="left">Description</td>
			<td align="left" width="100%" ><input type="text" size="50" maxlength="255" name="fcname" value="<%=fcname%>"><font color="#FF0000">*</font></td>		
		</tr>
		<tr> 
            <td align="left"><font face="Verdana" size="1">Doc Type</font></td>
            <td align="left">
				<select name="v_type">
					<%
					qlov = "SELECT fccode, fcname FROM document_type ORDER BY orderdata"
					Set rslov = Server.CreateObject("ADODB.Recordset")
					rslov.Open qlov, connstr 
					if Not rslov.eof then
						While Not rslov.EOF
					%>
					<option value="<%=rslov("fccode")%>" <%if v_type = rslov("fccode") then%>selected<%end if%>><%=rslov("fccode")%>-<%=rslov("fcname")%></option>
					<%
							rslov.MoveNext
						Wend
					end if
					%>
				</select>
			</td>
        </tr>
		<tr> 
			<td width="189" align="left"></td>	
		</tr>
		<tr> 
			<td width="189" align="left"></td>	
		</tr>
		
		<tr>
			<td colspan="2" align="left">
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
				  <input Type="Button"  value="Back" name="B2" OnClick="javascript:history.back(1)">
			</td>
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
	<tr bgcolor="#0C5B24" style="color:#FFFFFF;" style="color:#FFFFFF;" >
	<TD align="left">NO</TD>
	<TD align="center"colspan="2">ACTION</TD>
	<TD align="left">CODE</TD>	
	<TD align="left">DESCRIPTION</TD>
	<TD align="left">DOC. TYPE</TD>
	</tr>
   
   <%
   if Not rsGAD.eof then
      nCountofField=rsGAD.fields.count - 5
	  no=1
	  While Not rsGAD.EOF
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" valign="top"><%=NO%></td>
			<td valign="middle" width="30">
			<a href="document_classification.asp?fcCode=<%=rsGAD("fcCode")%>&txttable=<%=request("txttable")%>&fcname=<%=rsGAD("fcname")%>&v_type=<%=rsGAD("document_type")%>&txtJenis=Edit"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsGAD(1)%>" border="0"></a>		
			</td>
			<td valign="middle" width="30">
			<a href="document_classification.asp?fcCode=<%=rsGAD("fcCode")%>&txttable=<%=request("txttable")%>&fcname=<%=rsGAD("fcname")%>&v_type=<%=rsGAD("document_type")%>&txtJenis=Delete"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsGAD(1)%>" border="0" onClick="return confirm('Are you sure to Delete');"></a>
			</td>
		    <td align="center" title="<%=rsGAD("fcCode").Name%>"><%=rsGAD("fccode")%></td>
		    <td align="center" title="<%=rsGAD("fcname").Name%>"><%=rsGAD("fcname")%></td>
		    <td align="center" title="<%=rsGAD("document_type").Name%>"><%=rsGAD("document_type")%></td>
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
