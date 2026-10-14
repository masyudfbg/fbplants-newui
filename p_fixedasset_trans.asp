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
<title>ePlaS Reporting- Fiexed Asset </title>
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
creationdate = request("creationdate")
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

txtSQL = 	" SELECT fa.fccode" & _
		    "        , fa.fcname" & _
		    "        , bu.fccompanycode" & _
		    "        , fa.fcba" & _
		    "        , fa.transdate" & _
			"        , fa.sivcode" & _
		    "        , fa.assetcode" & _
		    "        , ass.fcname assetname" & _
		    "        , fa.assetgroupcode" & _
		    "        , fa.assetgroupname" & _
		    "        , ass.unitofmeasurement" & _
		    "        , asg.depreciationrange" & _
		    "        , NVL( acquisitionvalue, 0) acquisitionvalue" & _
		    "        , NVL( additionalvalue, 0) additionalvalue" & _
		    "        , fa.description" & _
		    "        , fa.isactive" & _
		    "        , DECODE(fa.isactive,'1','TRUE', 'FALSE') isactivedesc" & _
		    " FROM fixedasset fa" & _
		    " LEFT JOIN asset ass ON ass.fccode = fa.assetcode" & _
			" LEFT JOIN assetgroup asg ON asg.fccode = ass.assetgroupcode" & _
			" LEFT JOIN businessunit bu ON fa.fcba = bu.fccode" & _
	        " WHERE 1=1"

if  len(trim(creationdate_from)) > 0 then
    txtSQL = txtSQL & " and fa.transdate >=to_Date('" & creationdate_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(creationdate_from,1)
end if
if  len(trim(creationdate_to)) > 0 then
    txtSQL =txtSQL &  " and fa.transdate <=to_Date('" & creationdate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(creationdate_to,1)
end if
if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  " and fa.fcba = '"& businessunit &"'"
end if
if  len(trim(assetcode)) > 0 then
    txtSQL =txtSQL &  " and fa.assetcode = '"& assetcode &"'"
end if
if  len(trim(isactive)) > 0 then
    txtSQL =txtSQL &  " and DECODE(fa.isactive,'1','TRUE', 'FALSE') = '"& isactive &"'"
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
txtSQL =txtSQL &  " ORDER BY fa.assetgroupcode, fa.fccode"

txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "" & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by WIDATE, WIDIVISION"

'response.Write(txtsql)
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
					<h2>FIXED ASSET</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th></th>
					<th>Fixed Asset Code</th>
					<th>Fixed Asset Name</th>
					<th>No.Referensi</th>
					<th>Creation Date</th>
					<th>Asset Code</th>
					<th>Asset Name</th>
					<th>Asset Group</th>
					<th>Asset Group Name</th>
					<th>UoM</th>
					<th>Depreciation Range</th>
					<th>Aquisition Value</th>
					<th>Additional Value</th>
					<th>Total Value</th>
					<th>Description</th>
				</tr>
				</thead>
               <tbody>
	<%

	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
		totalvalue = 0
		total = 0
		totalacquisitionvalue = 0
		totaladditionalvalue = 0 
	While Not rsCat.EOF
		 totalvalue = cdbl(to_nol(rscat("acquisitionvalue"))) + cdbl(to_nol(rscat("additionalvalue")))

	%>
		<tr>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1">
				<a href="fixedasset_trans_edit.Asp?txtModule=Finance%20and%20Accounting&fccode=<%=rscat("fccode")%>&fcba=<%=rscat("fcba")%>&txtJenis=Edit" target="_blank">
                	<img src="IMAGES/edit.gif" alt="Edit Data : " border="0">
				</a>
			</td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcName")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sivcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("assetcode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("assetname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("assetgroupcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("assetgroupname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("unitofmeasurement")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("depreciationrange")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("acquisitionvalue")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("additionalvalue")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=totalvalue%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("description")%></td>
		</tr>
	  <%
	 
		total = total + totalvalue
		totalacquisitionvalue = totalacquisitionvalue + cdbl(to_nol(rscat("acquisitionvalue")))
		totaladditionalvalue = totaladditionalvalue + cdbl(to_nol(rscat("additionalvalue")))
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
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="11"> <font face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(totalacquisitionvalue,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(totaladditionalvalue,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(total,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->