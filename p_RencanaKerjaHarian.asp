<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Rencana Kerja Harian</title>
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
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>

<%
Dim Columnx : Columnx = 0
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtFDDATE_From=trim(REQUEST("FDDATE_From"))
v_txtFDDATE_To=trim(REQUEST("FDDATE_To"))
v_txtGANGCODE=trim(REQUEST("GANGCODE"))
v_txtJOBCODE=trim(REQUEST("JOBCODE"))
v_txtLOCATIONTYPE=trim(REQUEST("LOCATIONTYPE"))
v_txtLOCATIONCODE=trim(REQUEST("LOCATIONCODE"))
v_txtsectionname = Request("txtsectionname")
v_cField=trim(REQUEST("cField"))
businessunit= trim(request("businessunit"))

		' q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
		' 	' response.write q
		' 	set rsCat = connStr.Execute(q)
		' 	if not rsCat.eof then
		' 	   fcbaname = rsCat("fcName")
		' 	   CompanyName=rsCat("FCCOMPANYNAME")
		' 	   MgrName = rsCat("FCMANAGER")
		' 	   KTUName = rsCat("FCKTU")
		' 	   CropName=rsCat("FCCROP")
		' 	   FcbaType=rsCat("fcType")
		' 	end if
' if  len(trim(v_cField)) = 0 then
' 	txtSQL = 	" select fcba" & _
' 				"        , nvl(sum(mandays),0) mandays" & _
' 				"        , nvl(sum(othrs),0) overtime" & _
' 				"        , nvl(sum(unit),0) unit" & _
' 				"        , nvl(sum(output),0) output" & _
' 				"        , nvl(sum(rate),0) rate" & _
' 				" from (" & _
' 				"     select gad.gangcode" & _
' 				"            , gad.fddate" & _
' 				"            , gad.supervision_1" & _
' 				"            , gad.supervision_2" & _
' 				"            , gad.supervision_3" & _
' 				"            , gad.supervision_4" & _
' 				"            , gad.employeecode" & _
' 				"            , EMPLOYEE_API.get_EmployeeName(gad.fcba, gad.employeecode) employeename" & _
' 				"            , EMPLOYEE_API.get_sectionname(gad.fcba, gad.employeecode) sectionname" & _
' 				"            , gad.attendance" & _
' 				"            , gad.jobcode" & _
' 				"            , JOB_API.get_jobcode_name(gad.jobcode, gad.fcba) jobname" & _
' 				"            , gad.locationtype" & _
' 				"            , gad.locationcode" & _
' 				"            , nvl(gad.mandays, 0) mandays" & _
' 				"            , nvl(gad.othrs, 0) othrs" & _
' 				"            , nvl(gad.rate, 0) rate" & _
' 				"            , nvl(gad.unit, 0) unit" & _
' 				"            , nvl(gad.output, 0) output" & _
' 				"            , gad.reference" & _
' 				"            , gad.remarks" & _
' 				"            , gad.fcba" & _
' 				"     from attendance_gad gad" & _
' 				"     where gad.fcba = '" & businessunit & "'" & _
' 				" ) where fcba = '" & businessunit & "'"
' else
	txtSQL =	" select id, transno, occode, divcode, gangcode, fieldcode, cakupan, transtype, jobcode, rkhdate, uom_code, pht_l, pht_p, phl_l, phl_p from EPCS.RKHHEADER" & _
	            " WHERE 1 = 1"
	response.write txtSQL


' end if

if  len(trim(v_txtFDDATE_From)) > 0 then
    txtSQL = txtSQL & " and fddate >= to_date('" & v_txtFDDATE_From & "','mm/dd/yyyy')"
	txtSelect = " From " & FormatDateTime(v_txtFDDATE_From,1)
end if
if  len(trim(v_txtFDDATE_To)) > 0 then
    txtSQL =txtSQL &  " and fddate <= to_date('" & v_txtFDDATE_To & "','mm/dd/yyyy')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtFDDATE_To,1)
end if
if  len(trim(v_txtsectionname)) > 0 then
    txtSQL = txtSQL & " and UPPER(sectionname) like UPPER ('%" & v_txtsectionname & "%')"
	txtSelect = txtSelect & "  Section Name " & (v_txtsectionname)
end if
if  len(trim(v_txtGANGCODE)) > 0 then
    txtSQL = txtSQL & " and UPPER(gangcode) like UPPER ('%" & v_txtGANGCODE & "%')"
	txtSelect = txtSelect & "  Gang Code " & (v_txtGANGCODE)
end if
if  len(trim(v_txtJOBCODE)) > 0 then
    txtSQL =txtSQL &  " and jobcode = '" & v_txtJOBCODE & "'"
	txtSelect = txtSelect & "  Job Code " & (v_txtJOBCODE)
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and locationtype = '" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & "  TT " & (v_txtLOCATIONTYPE)
end if 
if  len(trim(v_txtLOCATIONCODE)) > 0 then
    txtSQL =txtSQL &  " and locationcode = '" & v_txtLOCATIONCODE & "'"
	txtSelect = txtSelect & "  TC " & (v_txtLOCATIONCODE)
end if

txtSQL = txtSQL & 	" ORDER BY id," & _
						"      transno," & _
						"      occode," & _
						"      divcode"

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
         <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="20" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>RENCANA KERJA HARIAN</h2>
				</div>
			</div>
					
		</th>
        </tr>
        		<!--<div class="printhide">-->
						<tr>
							<th colspan="20"  style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
						</tr>
						<!--</div>-->
						<tr>
							<%	
							
							Set rsCat = Server.CreateObject("ADODB.Recordset")
							rsCat.CursorType = 3
							rsCat.CursorLocation = 3
							rsCat.Open txtSQL, connstr

							
							if Not rsCat.eof then
							   nCountofField=rsCat.fields.count-1
							end if
							NO=1%>
						</tr>
						<tr>
							<!--<th>No.</th>-->
							<th>ID</th>
							<th>Trans No</th>
							<th>Co Code</th>
							<th>Div Code</th>
							<th>Gang Code</th>
							<th>Field Code</th>
							<th>Cakupan</th>
							<th>Trans Type</th>
							<th>Job Code</th>
							<th>RKH Date</th>
							<th>UOM Code</th>
							<th>PHT L</th>
							<th>PHT P</th>
							<th>PHL L</th>
							<th>PHL P</th>
						</tr>
                        </thead>
                        <tbody>
							<%
							Dim v_Opening,v_Debit,v_Credit,v_Saldo,v_Quantity
							v_Quantity=0
							v_Amount = 0
							While Not rsCat.EOF%>
							<tr>
								<!--<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%> </td>-->
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("id")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transno")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("occode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("divcode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("gangcode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldcode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("cakupan")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("transtype")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("rkhdate")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("uom_code")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pht_l")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("pht_p")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("phl_l")%> </td>
								<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("phl_p")%> </td>
							</tr>
							  <%

							  rsCat.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							rsCat.close()
							rsCat=null%>
							<input  type="hidden" value="<%=txtTable%>" name="txtTable">
							<input  type="hidden" value="<%=txtModule%>" name="txtModule">
							</form>
                            </tbody>
						</table>
						<!--#include file="Bottom_Report.asp"-->
		</DIV>
	</BODY>
</HTML>