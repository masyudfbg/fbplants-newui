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
<title>ePlaS Reporting- Fiexed Asset Tax Group </title>
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


assettaxgroupcode = request("assettaxgroupcode")
depreciationmethod = request("depreciationmethod")






txtSQL = 	" SELECT fccode, " & _
			"		 fcname, " & _
			"		 depreciationmethod, " & _
			"		 depreciationrange, " & _
			"		 depreciationrate, " & _
			"		 description " & _
		    " FROM Fixedasset_taxgroup" & _
	        " WHERE 1=1"


if  len(trim(assettaxgroupcode)) > 0 then
    txtSQL =txtSQL &  " and fccode = '"& assettaxgroupcode &"'"
end if
if  len(trim(depreciationmethod)) > 0 then
    txtSQL =txtSQL &  " and depreciationmethod = '"& depreciationmethod &"'"
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
				
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>FIXED ASSET TAX GROUP</h2>
				</div>
			</div>
</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="16" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
				</tr>
									

				<tr>
					<th>Asset Tax Group Code</th>
					<th>Asset Tax Group Name</th>
					<th>Depreciation Method</th>
					<th>Depreciation Range</th>
					<th>Depreciation Rate</th>
					<th>Description</th>
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
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("depreciationmethod")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("depreciationrange")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("depreciationrate")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("description")%></td>
			
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
