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
<title>ePlaS Reporting- Fiexed Asset Adjustement </title>
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
fccompanycode = request("fccompanycode")
businessunit = request("businessunit")
selectedfcba = request("selectedfcba")

assetadjustcode = request("assetadjustcode")
 

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

txtSQL = 	" SELECT fccode, " & _
			"		 fctype, " & _
			"		 transdate, " & _
			"		 description, " & _
			"		 fixedassetcode, " & _
			"		 fixedassetname, " & _
			"		 jobcode, " & _
			"		 targettype, " & _
			"		 targetcode, " & _
			"		 adjustment_amount, " & _
			"		 adjustment_monthdepreciation, " & _
			"		 remarks, " & _
			"		 fcba " & _
		    " FROM Fixedasset_adjustment" & _
	        " WHERE 1=1" 

if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  " and fcba = '"& businessunit &"'"
end if
if  len(trim(assetadjustcode)) > 0 then
    txtSQL =txtSQL &  " and fccode = '"& assetadjustcode &"'"
end if
if  len(trim(assetcode)) > 0 then
    txtSQL =txtSQL &  " and fixedassetcode = '"& assetcode &"'"
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
txtSQL =txtSQL &  "  ORDER BY fccode "

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
					<h2>FIXED ASSET ADJUSTEMENT</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th>Fixed Asset Adjustement Code</th>
					<th>Fixed Asset Adjustement Type</th>
					<th>Creation Date</th>
					<th>Description	Fixed Asset Code</th>
					<th>Fixed Asset name</th>
					<th>Job code</th>
					<th>Target Type</th>
					<th>Target Code</th>
					<th>Adjustment Amount</th>
					<th>Adjustment Month Depreciation</th>
					<th>Remarks</th>
				</tr>
				</thead>
               <tbody>
	<%

	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	While Not rsCat.EOF

	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fctype")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transdate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("description")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fixedassetcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fixedassetname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targettype")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("targetcode")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("adjustment_amount")%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("adjustment_monthdepreciation")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
	  <%
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
	
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->