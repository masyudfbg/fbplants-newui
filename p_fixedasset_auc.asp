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
<title>ePlaS Reporting- Fiexed Asset AUC</title>
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

' v_txtCUSTOMERCODE= trim(request("periodyear"))
'-------------awal-----------------------
searchdata = request("searchdata")
fccode = request("fccode")
fcname = request("fcname")
assetcode = request("assetcode")
description = request("description")
isactive = request("isactive")
fccompanycode = request("fccompanycode")
businessunit = request("businessunit")
selectedfcba = request("selectedfcba")

acquisitionvalue = request("acquisitionvalue")
additionalvalue = request("additionalvalue")
totalvalue = request("totalvalue")

creationdate_from = request("creationdate_from")
creationdate_to = request("creationdate_to")

documentno = Request("documentno")
creationdate = Request("creationdate")
fccompanycode = Request("fccompanycode")
fcba = Request("fcba")
selectedfcba = Request("selectedfcba")
fixedassetcode = Request("fixedassetcode")
description = Request("description")
lineno = Request("lineno")

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

txtSQL = 	"	SELECT fah.documentno," & _ 
			"       fah.creationdate," & _ 
			"       fah.fixedassetcode," & _ 
			"       fah.description," & _ 
			"       fad.accountingyear," & _ 
			"       fad.accountingperiod," & _ 
			"       fad.sourcetype," & _ 
			"       fad.jobcode," & _ 
			"       fad.amount" & _ 
			"  FROM fixedasset_auc_hdr fah" & _ 
			"       INNER JOIN fixedasset_auc_dtl fad" & _ 
			"          ON (fah.documentno = fad.documentno AND fah.fcba = fad.fcba)" & _ 
			" WHERE fah.fcba = '"& businessunit &"'"  

if  len(trim(creationdate_from)) > 0 then
    txtSQL = txtSQL & " and fah.creationdate >=to_Date('" & creationdate_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(creationdate_from,1)
end if
if  len(trim(creationdate_to)) > 0 then
    txtSQL =txtSQL &  " and fah.creationdate <=to_Date('" & creationdate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(creationdate_to,1)
end if
if  len(trim(documentno)) > 0 then
    txtSQL =txtSQL &  " and fah.documentno = '"& documentno &"'"
end if
if  len(trim(fixedassetcode)) > 0 then
    txtSQL =txtSQL &  " and fah.fixedassetcode = '"& fixedassetcode &"'"
end if

' if  len(trim(documentno)) > 0 then
' 	txtSelect = txtSelect & " Document No" & documentno
' end if
' if  len(trim(fieldcode)) > 0 then
' 	txtSelect = txtSelect & " Blok " & fieldcode
' end if
' if  len(trim(fullstand)) > 0 then
' 	txtSelect = txtSelect & " Fullstand " & fullstand
' end if

' if len(trim(documentno)) > 0 then txtSQL = txtSQL & " and assetgroupcode = '"& assetgroupcode &"'": txtSelect = txtSelect & " Asset Group " & assetgroupcode
' if len(trim(fieldcode)) > 0 then txtSQL = txtSQL & " and depreciationmethod = '"& depreciationmethod &"'" : txtSelect = txtSelect & " Depreciation Method " & depreciationmethod
' if len(trim(fullstand)) > 0 then txtSQL = txtSQL & " and controljob = '"& controljob &"'" : txtSelect = txtSelect & " Control Job " & controljob


txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

' response.Write(txtsql)
'response.Flush()
%>
 <body>
		<div id="wrapper">
 <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>FIXED ASSET AUC</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th>Document No</th>
					<th>Creation Date</th>
					<th>Fixed Asset Code</th>
					<th>Description</th>
					<th>Source Type</th>
					<th>Job Code</th>
					<th>Amount</th>
				</tr>
				</thead>
               <tbody>
	<%

	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
		totalamount = 0
		nCountofField=rsCat.fields.count-1
	While Not rsCat.EOF

	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("documentno")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("creationdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fixedassetcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("description")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sourcetype")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("amount"))),2)%></td>
		</tr>
	  <%
	  totalamount = totalamount + cdbl(to_nol(rscat("amount"))) 
	  rsCat.MoveNext

	Wend
	Else
	response.write "<tr><td colspan='16'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="6"> <font face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(totalamount,2)%></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->