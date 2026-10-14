<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Grading Record</title>
</head>

<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtDATETIME_From= trim(request("DATETIME_From"))
v_txtDATETIME_To= trim(request("DATETIME_To"))
v_txtGRADINGNO= trim(request("GRADINGNO"))
v_txtWBT_CHITNO= trim(request("WBT_CHITNO"))
v_txtVEHICLENO=trim(request("VEHICLENO"))

txtSQL = "SELECT GRADINGNO, DATETIME, VEHICLENO, nvl(sum(WEIGHT),0) as WEIGHT, nvl(sum(BUNCHES),0) as BUNCHES, REMARKS, CHECKEDBY, WBT_CHITNO FROM GRADINGRECORD" & _
					 " where fcba='" & session("fcBA") & "'"

if  len(trim(v_txtDATETIME_From)) > 0 then
    txtSQL = txtSQL & " and DATETIME >=to_Date('" & v_txtDATETIME_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDATETIME_From,1)
end if
if  len(trim(v_txtDATETIME_To)) > 0 then
    txtSQL =txtSQL &  " and DATETIME <=to_Date('" & v_txtDATETIME_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtDATETIME_To,1)
end if
if  len(trim(v_txtGRADINGNO)) > 0 then
    txtSQL = txtSQL & " and GRADINGNO='" & v_txtGRADINGNO & "'"
	txtSelect = txtSelect & " Grading No " & v_txtGRADINGNO
end if
if  len(trim(v_txtWBT_CHITNO)) > 0 then
    txtSQL =txtSQL &  " and WBT_CHITNO='" & v_txtWBT_CHITNO & "'"
	txtSelect = txtSelect & " Chit No " & v_txtWBT_CHITNO
end if 
if  len(trim(v_txtVEHICLENO)) > 0 then
    txtSQL =txtSQL &  " and VEHICLENO='" & v_txtVEHICLENO & "'"
	txtSelect = txtSelect & " VEHICLENO " & v_txtVEHICLENO
end if 

txtsql=(txtsql) & " group by GRADINGNO, DATETIME, VEHICLENO, REMARKS, CHECKEDBY, WBT_CHITNO " 
'                     "VEHICLE, LOADERCODE, RAMPATTENDANTCODE,WEIGHT"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
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
					<h2>GRADING RECORD</h2>
				</div>
			</div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<div class="printhide">

				<tr>
					<td colspan="8"><%=txtSelect%></td>
				</tr>
				</div>
				<%	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO=1%>
				<tr>
					<th>Grading No</th>
					<th>Date</th>
					<th>Vehicle No</th>
					<th>Remarks</th>
					<th>Checked By</th>
					<th>Chit No</th>
					<th>Weight</th>
					<th>Bunches</th>
				</tr>
	<%
	Dim v_WEIGHT,v_BUNCHES
	v_WEIGHT=0
	v_BUNCHES=0
	While Not rsCat.EOF%>
		<tr>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("GRADINGNO")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DATETIME")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLENO")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARKS")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CHECKEDBY")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WBT_CHITNO")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("WEIGHT"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("BUNCHES"),2)%></td>
		</tr>
	  <%
	  v_WEIGHT=v_WEIGHT+trim(cdbl(rsCat("WEIGHT")))
	  v_BUNCHES=v_BUNCHES+trim(cdbl(rsCat("BUNCHES")))
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
	<tr bgcolor="#F4F4A0">
	<td align="Center" colspan="6"> <strong><i> <font face="Verdana" size="2">T o t a l</i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_WEIGHT,2)%></i></strong></td>
	<td align="Right"> <strong><i> <font face="Verdana" size="1"><%=FormatNumber(v_BUNCHES,2)%></i></strong></td>
	</tr>
	
</table>
<!--#include file="Bottom_Report.asp"-->