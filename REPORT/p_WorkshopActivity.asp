<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Workshop Activity</title>
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

v_WORKSHOPDATE_From=trim(request("DATEACTIVITY_From")) 
v_WORKSHOPDATE_To=trim(request("DATEACTIVITY_To"))
v_JOB=trim(request("Job"))
v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
v_LOCATIONCODE=trim(request("LOCATIONCODE"))

if request("txtKind") = "WorkshopRunningAccount" Then
	txtSQL = "select ws.WORKSHOPCODE, ws.WORKSHOPDATE, ws.TIME, ws.LOCATIONTYPE, ws.LOCATIONCODE,ws.JOB, j.fcname,ws.ref, ws.remarks," & _
     " nvl(sum(ws.HOURS),0) as Hours,nvl(sum(ws.UNIT),0) as Unit,nvl(sum(ws.OUTPUT),0) as Output,nvl(sum(AMOUNT),0) as Amount" & _
     " from WorkshopActivity ws,job j " & _
     " where (ws.job = j.fccode and ws.fcba = j.fcba) " & _
	 " AND WS.fcba='" & session("fcBA") & "'" & _
	 " AND j.fcba='" & session("fcBA") & "'" 
Else
	txtSQL = "select ws.WORKSHOPCODE, ws.WORKSHOPDATE, ws.TIME, ws.LOCATIONTYPE, ws.LOCATIONCODE,ws.JOB, j.fcname, ws.ref, ws.remarks," & _
     " nvl(sum(ws.HOURS),0) as Hours,nvl(sum(ws.UNIT),0) as Unit,nvl(sum(ws.OUTPUT),0) as Output,nvl(sum(AMOUNT),0) as Amount" & _
     " from WorkshopActivity ws,job j " & _
     " where (ws.job = j.fccode and ws.fcba = j.fcba) " & _
	 " AND WS.fcba='" & session("fcBA") & "'"	
End If


if  len(trim(v_WORKSHOPDATE_From)) > 0 then
    txtSQL = txtSQL & " and Ws.WORKSHOPDATE >=to_Date('" & v_WORKSHOPDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_WORKSHOPDATE_From,1)
end if
if  len(trim(v_WORKSHOPDATE_To)) > 0 then
    txtSQL =txtSQL &  " and Ws.WORKSHOPDATE <=to_Date('" & v_WORKSHOPDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_WORKSHOPDATE_To,1)
end if
if  len(trim(v_JOB)) > 0 then
    txtSQL = txtSQL & " and ws.JOB='" & v_JOB & "'"
	txtSelect = txtSelect & " Job Code " & v_JOB
end if
if  len(trim(v_LOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and ws.LOCATIONTYPE='" & v_LOCATIONTYPE & "'"
	txtSelect = txtSelect & " Target Type " & v_LOCATIONTYPE
end if 
if  len(trim(v_LOCATIONCODE)) > 0 then
    txtSQL =txtSQL &  " and ws.LOCATIONCODE='" & v_LOCATIONCODE & "'"
	txtSelect = txtSelect & " Target Code " & v_LOCATIONCODE
end if 
txtsql=(txtsql) & " group BY ws.WORKSHOPCODE, ws.WORKSHOPDATE, ws.TIME, ws.LOCATIONTYPE, ws.LOCATIONCODE,ws.JOB, j.fcname,ws.time,ws.ref, ws.remarks"
'response.Write(txtsql)
response.Flush()
%>
    <body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>WORKSHOP ACTIVITY</h2>
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
	NO=1
	Dim v_Quantity,v_qtyretrun,v_amountreceived,v_amountreturn
	v_HOURS=0
	v_UNIT=0
	v_OUTPUT=0
	v_AMOUNT=0
	%>
	</tr>
	<tr>
	    <th>Code</th>
		<th>Date</th>
		<th>Time</th>
		<th>Job Code</th>
		<th>Job Description</th>
		<th>Target Type</th>
		<th>Target Code</th>
		<th>Hours</th>
		<th>Unit</th>
		<th>Output</th>
		<th>Amount</th>
		<th>Reference</th>
		<th>Remarks</th>		
	</tr>
	<%While Not rsCat.EOF%>
	 	<tr>
		 
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WORKSHOPCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("WORKSHOPDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIME")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JOB")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>		
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONTYPE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONCODE")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Hours"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("UNIT"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Output"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Amount"),2)%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ref")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
	  <%
	  v_HOURS=v_HOURS+trim(cdbl(rsCat("Hours")))
	  v_UNIT  =v_UNIT+Trim(cdbl(rsCat("UNIT")))
	  v_OUTPUT =v_OUTPUT+Trim(cdbl(rsCat("Output")))
	  v_AMOUNT  =v_AMOUNT+trim(cdbl(rsCat("Amount")))
	  rsCat.MoveNext
	  response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="7"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_HOURS,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_UNIT,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_OUTPUT,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_AMOUNT,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
	
	</tr>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>