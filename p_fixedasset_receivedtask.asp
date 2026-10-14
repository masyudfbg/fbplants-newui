
 
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
<title>ePlaS Reporting- Fixed Asset Received Task</title>
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

fcname = request("fcname")

assetcode = request("assetcode")
fccompanycode = request("fccompanycode")
businessunit = request("businessunit")
selectedfcba = request("selectedfcba")
businessunit = request("businessunit")
TransportTaskNo = request("TransportTaskNo")
creationdate_from= trim(request("creationdate_from"))
creationdate_to=trim(request("creationdate_to"))

PSTORAGE_STORAGECODE = request("PSTORAGE_STORAGECODE")
 

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

txtSQL = 	"   Select h.transdate,h.fccode, h.fcba_source,h.fa_transportask,h.fa_transportaskdate,d.fixassetcode,d.fixassetname,d.assetcode,d.assetname, " & _
            "           d.assetgroupcode, d.assetgroupname, d.assettaxcode, d.assettaxname, d.uom, d.depreciationmonth, d.fixassetamount, " & _
            "          d.targettype_fixasset , d.targetcode_fixasset , d.sourcetype_depreciation, d.sourcecode_depreciation, d.jobcode_depreciation, " & _
            "          d.accumulated_depreciation, d.jobcode_intraco , d.sourcetype_intraco , d.sourcecode_intraco , d.intraco_amount, d.remarks, h.fcba " & _
			"	From FIXEDASSET_RECEIVEDTASK H " & _
			"	 Inner Join FIXEDASSET_RECEIVEDTASK_DET d " & _
			"		on h.fccode = d.fccode " & _
			"	   And h.fcba = d.fcba " & _
			"	WHERE 1 = 1 " 

if  len(trim(businessunit)) > 0 then
    txtSQL =txtSQL &  " and h.fcba = '"& businessunit &"'"
end if
if  len(trim(businessunitSource)) > 0 then
    txtSQL =txtSQL &  " and h.fcba_source = '"& businessunitSource &"'"
end if
if  len(trim(TransportTaskNo)) > 0 then
    txtSQL =txtSQL &  " and h.fccode = '"& TransportTaskNo &"'"
end if
if  len(trim(assetcode)) > 0 then
    txtSQL =txtSQL &  " and d.assetcode = '"& assetcode &"'"
end if
if  len(trim(creationdate_from)) > 0 then
    txtSQL = txtSQL & " and h.transdate >=to_Date('" & creationdate_from & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(creationdate_from,1)
end if
if  len(trim(creationdate_to)) > 0 then
    txtSQL =txtSQL &  " and h.transdate <=to_Date('" & creationdate_to & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(creationdate_to,1)
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
txtSQL =txtSQL &  "  ORDER BY h.transdate, h.fccode "

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
         <th colspan="28" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>FIXED ASSET RECEIVED TASK</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="28" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
				<tr>
					<th>DATE</th>
					<th>BUSINESS UNIT SOURCE</th>
					<th>FIXED ASSET TRANSPORT TASK SOURCE</th>
					<th>FIXED ASSET TRANSPORT TASK DATE</th>
					<th>FIXED ASSET TRANSPORT TASK NO</th>
					<th>FIXEDASSET CODE</th>
					<th>FIXEDASSET NAME</th>
					<th>ASSET CODE</th>
					<th>ASSET NAME</th>
					<th>ASSET GROUP CODE</th>
					<th>ASSET GROUP NAME</th>
					<th>ASSET TAX CODE</th>
					<th>ASSET TAX NAME</th>
					<th>UOM</th>
					<th>DEPRECIATION MONTH</th>
					<th>FIXEDASSET AMOUNT</th>
					<th>TARGET TYPE DEPRECIATION</th>
					<th>TARGET CODE DEPRECIATION</th>
					<th>SOURCE TYPE DEPRECIATION</th>
					<th>SOURCE CODE DEPRECIATION</th>
					<th>JOBCODE DEPRECIATION</th>
					<th>ACCUMULATED DEPRECIATION</th>
					<th>JOBCODE INTRCAO</th>
					<th>SOURCETYPE INTRACO</th>
					<th>SOURCECODE INTRACO</th>
					<th>INTRACO AMOUNT</th>
					<th>REMARKS</th>
				</tr>
				</thead>
               <tbody>
	<%

	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	While Not rsCat.EOF

	%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("h.transdate")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("h.fcba_source")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("h.fa_transportask")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("h.fa_transportaskdate")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("h.fccode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.fixassetcode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.fixassetname")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.assetcode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.assetname")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.assetgroupcode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.assetgroupname")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.assettaxcode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.assettaxname")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.uom")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.depreciationmonth")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.fixassetamount")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.targettype_fixasset")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.targetcode_fixasset")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.sourcetype_depreciation")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.sourcecode_depreciation")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.jobcode_depreciation")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.accumulated_depreciation")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.jobcode_intraco")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.sourcetype_intraco")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.sourcecode_intraco")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.intraco_amount")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("d.remarks")%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	Else
	response.write "<tr><td colspan='29'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->