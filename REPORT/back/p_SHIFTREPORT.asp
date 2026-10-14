<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Shift Report</title>
</head>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtPRODUCTIONDATE_From= trim(request("PRODUCTIONDATE_From"))
v_txtPRODUCTIONDATE_To= trim(request("PRODUCTIONDATE_To"))
v_txtSHIFT_SHIFTCODE= trim(request("SHIFT_SHIFTCODE"))

txtSQL = "SELECT PRODUCTIONDATE, ACTUALSTART, ACTUALEND, CAGESPROCESSED, RAMPBALANCE, CAGESBALANCE, " & _
			"PLATFORMBALANCE, DOWNTIME, CAUSE, SHIFT_SHIFTCODE, CAGEWEIGHT FROM SHIFTREPORT" & _
					 " where fcba='" & session("fcBA") & "'"

if  len(trim(v_txtPRODUCTIONDATE_From)) > 0 then
    txtSQL = txtSQL & " and PRODUCTIONDATE >=to_Date('" & v_txtPRODUCTIONDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtPRODUCTIONDATE_From,1)
end if
if  len(trim(v_txtPRODUCTIONDATE_To)) > 0 then
    txtSQL =txtSQL &  " and PRODUCTIONDATE <=to_Date('" & v_txtPRODUCTIONDATE_To & "','MM/DD/YYYY')"
	xtSelect = txtSelect & "  to " & FormatDateTime(v_txtPRODUCTIONDATE_To,1)
end if
if  len(trim(v_txtSHIFT_SHIFTCODE)) > 0 then
    txtSQL = txtSQL & " and SHIFT_SHIFTCODE='" & v_txtSHIFT_SHIFTCODE & "'"
	txtSelect = txtSelect & " SHIFTCODE " & v_txtSHIFT_SHIFTCODE
end if

txtsql=(txtsql) 
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'& " group by CHITNO, FIELDCODE, CROPCODE, ProductCODE, COLLECTIONDATE,HARVESTINGDATE,RECEPTIONDATE, " & _
'                     "VEHICLE, LOADERCODE, RAMPATTENDANTCODE,WEIGHT"

'response.Write(txtsql)
'response.Flush()
%>
 <body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">




					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>SHIFT REPORT</h2>
				</div>
			</div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<div class="printhide">

				<tr>
					<td colspan="13"><%=txtSelect%></td>
				</tr>
				</div>
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Shift Code</th>
					<th>Production Date</th>
					<th>Actual Start</th>
					<th>Actual End</th>
					<th>Cages Processed</th>
					<th>Ramp Balance</th>
					<th>Cages Balance</th>
					<th>Platform Balance</th>
					<th>Down Time</th>
					<th>Cause Cage</th>
					<th>Weight</th>
				</tr>
	<%
	Dim v_Quantity
	v_Quantity=0
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SHIFT_SHIFTCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCTIONDATE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ACTUALSTART")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ACTUALEND")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CAGESPROCESSED")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("RAMPBALANCE")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CAGESBALANCE")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PLATFORMBALANCE")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DOWNTIME")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CAUSE")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CAGEWEIGHT")%>&nbsp;</td>
		</tr>
	  <%
	  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
</table>
<!--#include file="Bottom_Report.asp"-->