<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<title>ePlaS Reporting - Authorize Structure</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

searchdata = request("searchdata")
fccode = request("fccode")
fcname = request("fcname")
document_type = request("document_type")
document_classification = request("document_classification")
authorizebyamount = request("authorizebyamount")
amountrangefrom = request("amountrangefrom")
amountrangeuntil = request("amountrangeuntil")
description = request("description")
isactive = request("isactive")
businessunit = request("businessunit")
authorizetemplatecode = request("authorizetemplatecode")

 		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if

txtSQL = 	" SELECT al.authorizelevelcode, al.authorizelevelname, al.fcusercode, uft.fcba, nvl(aus.isauthorize,0) isauthorize, nvl(aus.authorizetemplatecode,'-') authorizetemplatecode ,aus.document_classification  ,aus.document_type " & _
			" FROM (" & _
			" SELECT * FROM view_authorizelevel WHERE isactive = 'TRUE' AND startlevel = 'TRUE' AND isactive_user = 'TRUE'" & _
			" UNION ALL" & _
			" SELECT * FROM view_authorizelevel WHERE isactive = 'TRUE' AND startlevel = 'FALSE' AND finallevel = 'FALSE' AND isactive_user = 'TRUE'" & _
			" UNION ALL" & _
			" SELECT * FROM view_authorizelevel WHERE isactive = 'TRUE' AND finallevel = 'TRUE' AND isactive_user = 'TRUE'" & _
			" ) al" & _
			" INNER JOIN user_fcba_tab uft ON al.fcusercode = uft.fcusercode" & _
			" LEFT JOIN ( " & _
			" SELECT authorizetemplatecode, authorizelevelcode, fcusercode, 1 isauthorize, x.fcba ," & _
			" y.document_type,z.fcname document_classification" & _
			" FROM authorizestructure x " & _
			" INNER JOIN  authorizetemplate y ON (x.authorizetemplatecode=y.fccode)" & _
			" INNER JOIN  document_classification z ON (y.document_classification=z.fccode AND y.document_type=z.document_type AND z.isactive='TRUE' ) " & _
			" ) aus ON (aus.authorizetemplatecode like '"&authorizetemplatecode&"%' AND al.authorizelevelcode = aus.authorizelevelcode AND al.fcusercode = aus.fcusercode AND aus.fcba = uft.fcba)" & _
			" WHERE uft.fcba = '"&businessunit&"'  AND isauthorize > 0 /*AND document_type = '"&document_type&"' AND document_classification = '"&document_classification&"'*/" & _
			" ORDER BY aus.document_type,aus.document_classification  ,al.authorizelevelcode, al.fcusercode"

' if  len(trim(fccode)) > 0 then
'     txtSQL =txtSQL &  "  and fccode = '"& fccode &"'"
' end if
' if  len(trim(document_type)) > 0 then
'     txtSQL =txtSQL &  "  and document_type = '"& document_type &"'"
' end if
' if  len(trim(document_classification)) > 0 then
'     txtSQL =txtSQL &  " and document_classification = '"& document_classification &"'"
' end if
' if  len(trim(isactive)) > 0 then
'     txtSQL =txtSQL &  " and isactive = '"& isactive &"'"
' end if

txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

'response.Write(txtsql)
response.Flush()
Set rsCat = Server.CreateObject("ADODB.Recordset")
	rsCat.CursorType = 3
	rsCat.CursorLocation = 3
	rsCat.Open txtSQL, connstr
%>
 <body>
		<div id="wrapper">
		<!--#include file="print_excel.asp"-->
 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>AUTHORIZE STRUCTURE</h2>
				</div>
			</div>
		 </th>
			</tr>
						<!--<div class="printhide">-->
				<tr>
					<td colspan="14" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
				    <th>Document.Code</th>
					<th>Document.Type</th>
					<th>Template.Code</th>
					<th>Level.Code</th>
					<th>Description</th>
					<th>User Name</th>
					<th>Authorize</th>
				</tr>
                </thead>
               <tbody>
	<%
	if Not rscat.eof then
		v_amountrangefrom = 0
		v_amountrangeuntil = 0
	While Not rscat.EOF


	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("document_classification")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("document_type")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("authorizetemplatecode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("authorizelevelcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("authorizelevelname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcusercode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("isauthorize")%></td>
		</tr>
	  <%

	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='14'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	</tbody>
</table>
   <table width="90%" class="data">
   	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=MgrName%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=KTUName%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
       </tbody>
   </table>
   </body>
   </html>