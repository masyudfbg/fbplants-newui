<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Vehicle Activity</title>
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

v_VEHICLEDATE_From=trim(request("VEHDATE_From"))
v_VEHICLEDATE_To=trim(request("VEHDATE_to"))
v_JOBCODE=trim(request("JOBCODE"))
v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
v_LocationCode=trim(request("LocationCode"))
v_VEHICLEGROUPCODE=trim(request("VEHICLEGROUPCODE"))
v_VEHICLECODE=trim(request("VEHICLECODE"))

if request("txtKind") = "VehicleRunningAccount" Then
txtSQL = "select vh.VEHICLECODE, vh.VEHICLEDATE, vh.TIME, vh.LOCATIONTYPE, vh.LOCATIONcode,vh.JOBCODE, j.fcname,vh.ref,vh.remarks," & _
     " NVL(sum(vh.HOURS),0) as Hours,NVL(sum(vh.UNIT),0) as Unit,NVL(sum(vh.OUTPUT),0) as Output,NVL(sum(AMOUNT),0) as Amount" & _
     " from VEHICLEACTIVITY vh,job j " & _
     " where (vh.jobcode = j.fccode and vh.fcba = j.fcba) " & _
	 " AND vh.fcba='" & session("fcBA") & "'"
Else
txtSQL = "select vh.VEHICLECODE, vh.VEHICLEDATE, vh.TIME, vh.LOCATIONTYPE, vh.LOCATIONcode,vh.JOBCODE, j.fcname,vh.ref,vh.remarks," & _
     " NVL(sum(vh.HOURS),0) as Hours,NVL(sum(vh.UNIT),0) as Unit,NVL(sum(vh.OUTPUT),0) as Output,NVL(sum(AMOUNT),0) as Amount" & _
     " from VEHICLEACTIVITY vh,job j " & _
     " where (vh.jobcode = j.fccode and vh.fcba = j.fcba) " & _
	 " AND vh.fcba='" & session("fcBA") & "'"
End If
	 
if  len(trim(v_VEHICLEDATE_From)) > 0 then
    txtSQL = txtSQL & " and vh.VEHICLEDATE >=to_Date('" & v_VEHICLEDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_VEHICLEDATE_From,1)
end if
if  len(trim(v_VEHICLEDATE_To)) > 0 then
    txtSQL =txtSQL &  " and vh.VEHICLEDATE <=to_Date('" & v_VEHICLEDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_VEHICLEDATE_To,1)
end if
if  len(trim(v_JOBCODE)) > 0 then
    txtSQL = txtSQL & " and vh.JOBCODE='" & v_JOBCODE & "'"
	txtSelect = txtSelect & " Job Code " & v_JOBCODE
end if
if  len(trim(v_LOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and vh.LOCATIONTYPE='" & v_LOCATIONTYPE & "'"
	txtSelect = txtSelect & " Target Type " & v_LOCATIONTYPE
end if 
if  len(trim(v_LocationCode)) > 0 then
    txtSQL =txtSQL &  " and vh.LOCATIONCODE='" & v_LocationCode & "'"
	txtSelect = txtSelect & " Target Code " & v_LocationCode
end if 
if  len(trim(v_VEHICLEGROUPCODE)) > 0 then
    txtSQL =txtSQL &  " and vh.VEHICLECODE in (select FCCODE from Vehicle where VEHICLEGROUPCODE='" & v_VEHICLEGROUPCODE & "')"
	txtSelect = txtSelect & " Vehicle Group " & v_VEHICLEGROUPCODE
end if 
if  len(trim(v_VEHICLECODE)) > 0 then
    txtSQL =txtSQL &  " and vh.VEHICLECODE='" & v_VEHICLECODE & "'"
	txtSelect = txtSelect & " Vehicle Code " & v_VEHICLECODE
end if 

txtsql=(txtsql) & " group BY vh.VEHICLECODE, vh.VEHICLEDATE, vh.TIME, vh.LOCATIONTYPE, vh.LOCATIONCODE,vh.JOBCODE, j.fcname,vh.ref,vh.remarks"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

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
					<h2>VEHICLE ACTIVITY</h2>
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
		<th>Job<br>Code</th>
		<th>Job<br>Description</th>
		<th>Target<br>Type</th>
		<th>Target<br>Code</th>
		<th>Hours</th>
		<th>Unit</th>
		<th>Output</th>
		<th>Amount</th>
		<th>Reference</th>
		<th>Remarks</th>	
	</tr>
	<%While Not rsCat.EOF%>
	 	<tr>
		 
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLECODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("VEHICLEDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TIME")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JOBCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>		
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONTYPE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONcode")%></td>
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