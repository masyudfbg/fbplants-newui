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
<title>ePlaS Reporting- Fixed Asset Write Off </title>
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
fctype = request("fctype")

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

txtSQL = 	" SELECT fwh.fccode," & _
			"        fwh.transdate," & _
			"        fwh.fctype," & _
			"        fwh.description," & _
			"        fwd.fixassetcode," & _
			"        fwd.fixassetname," & _
			"        fwd.fixassetamount," & _
			"        fwd.targettype_depreciation," & _
			"        fwd.targetcode_depreciation," & _
			"        fwd.jobcode_depreciation," & _
			"        fwd.accumulated_depreciation," & _
			"        fwd.targettype_others," & _
			"        fwd.targetcode_others," & _
			"        fwd.jobcode_others," & _
			"        fwd.ar_others," & _
			"        fwd.targettype_gain," & _
			"        fwd.targetcode_gain," & _
			"        fwd.jobcode_gain," & _
			"        fwd.gain_loss," & _
			"        fwd.remarks" & _
			"   FROM fixedasset_writeoff_det fwd" & _
			"        INNER JOIN fixedasset_writeoff fwh" & _
			"           ON (fwd.fcba = fwh.fcba AND fwd.fccode = fwh.fccode)" & _
			"  WHERE     fwd.fcba = '"& businessunit &"'" 

if  len(trim(creationdate_from)) > 0 then
    txtSQL = txtSQL & " and transdate >=to_Date('" & creationdate_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(creationdate_from,1)
end if
if  len(trim(creationdate_to)) > 0 then
    txtSQL =txtSQL &  " and transdate <=to_Date('" & creationdate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(creationdate_to,1)
end if
if  len(trim(fctype)) > 0 then
    txtSQL =txtSQL &  " and fctype = '"& fctype &"'"
end if
' if  len(trim(assetcode)) > 0 then
    ' txtSQL =txtSQL &  " and fa.assetcode = '"& assetcode &"'"
' end if
' if  len(trim(isactive)) > 0 then
    ' txtSQL =txtSQL &  " and fa.isactive = '"& isactive &"'"
' end if

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
txtSQL =txtSQL &  " ORDER BY fwh.fccode, fwd.fixassetcode, fwd.fixassetname"

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
         <th colspan="18" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>FIXED ASSET WRITE OFF</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="18" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th>Fixed Asset Write Off</th>
					<th>Fixed Asset Code</th>
					<th>Fixed Asset Name</th>
					<th>Creation Date</th>
					<th>Jobcode Depreciation</th>
					<th>Target Type Depreciation</th>
					<th>Target Code Depreciation</th>
					<th>Jobcode Others</th>
					<th>Target Type Others</th>
					<th>Target Code Others</th>
					<th>Jobcode Gain</th>
					<th>Target Type Gain</th>
					<th>Target Code Gain</th>
					<th>Fixed Asset Amount</th>
					<th>Accumulated Depreciation</th>
					<th>AR Others</th>
					<th>Gain Loss</th>
					<th>Remarks</th>
				</tr>
				</thead>
               <tbody>
	<%

	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
		v_fixassetamount = 0
		v_accumulated_depreciation = 0
		v_ar_others = 0
		v_gain_loss = 0 
	While Not rsCat.EOF

	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fixassetcode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fixassetname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode_depreciation")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targettype_depreciation")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targetcode_depreciation")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode_others")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targettype_others")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targetcode_others")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode_gain")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targettype_gain")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targetcode_gain")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("fixassetamount"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("accumulated_depreciation"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("ar_others"))),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(cdbl(to_nol(rscat("gain_loss"))),2)%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
	  <%
		v_fixassetamount = v_fixassetamount + cdbl(to_nol(rscat("fixassetamount")))
		v_accumulated_depreciation = v_accumulated_depreciation + cdbl(to_nol(rscat("accumulated_depreciation")))
		v_ar_others = v_ar_others + cdbl(to_nol(rscat("ar_others")))
		v_gain_loss = v_gain_loss + cdbl(to_nol(rscat("gain_loss")))
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
	<tr bgcolor="#F4F4A0" style="font-weight:bold;">
		<td align="Center" style="<%=cmdstyleLeft%>" colspan="13"> <font face="Verdana" size="1">T o t a l</td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_fixassetamount,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_accumulated_depreciation,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_ar_others,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_gain_loss,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"></td>
	</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->