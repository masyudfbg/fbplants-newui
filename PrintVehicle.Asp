<%'Option Explicit
%>
<%' On Error GOTO CmdErrorlog 
%>
<%response.buffer=True%>
<%'validatesessionConst adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>RW Plant</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
</head>
<body>
 <%
    Dim txtNoWS,txtDateWS,txtRemarks,q,nCountofField,no,i
    Dim Remarks
    Dim rsWS,intRecordsPerPage,cColor,txtVehicle
	txtVehicle =cmdFindTitle("Vehicle","fcCode",request("VehicleCode"))
    %>
 <table width="700">
  <tr> 
    <td rowspan="6" width="100">
	<% if len(trim(PROFIL("REPORT_IMAGE"))) > 0 then %>
		<img src="IMAGES/<%=PROFIL("REPORT_IMAGE")%>" width="210" height="80">
	<% end if %>
	</td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">VEHICLE ACTIVITY</FONT></td></tr>
	<tr><td width="100">Date</td><td  width="300"><b><FONT color="#000099"><%=Request("fdDate")%></FONT></td></tr>
	<tr><td width="100">Vehicle</td><td  width="300"><b><FONT color="#000099"><%=Request("VehicleCode")%></FONT></td></tr>
 </table>
   <%
   Dim rsVehicle,nUnit,nHours,nOutput,nAmount
   set rsVehicle = server.CreateObject("Adodb.recordset")
   q="select ws.VEHICLEDATE, ws.LOCATIONTYPE, ws.LOCATIONCODE, j.fcname,ws.time," & _
     " NVL(sum(ws.HOURS),0) as Hours,NVL(sum(ws.UNIT),0) as Unit,NVL(sum(ws.OUTPUT),0) as Output,NVL(sum(AMOUNT),0) as Amount" & _
     " from VehicleActivity ws,job j " & _
     " where (ws.jobcode = j.fccode and ws.fcba =j.fcba)and ws.VehicleCode='" & request("VehicleCode") & _
	 "' and to_char(VEHICLEDATE,'mm/yyyy')=to_char(to_date('" &  Request("fdDate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy')  " & _
     " and ws.fcBA='" & session("fcBA") & "'" & _
	 "group by ws.VEHICLEDATE, ws.LOCATIONTYPE, ws.LOCATIONCODE, j.fcname,ws.time"
   'response.Write q	
   Set rsVehicle = Server.CreateObject("ADODB.Recordset")
   rsVehicle.CursorType = 3
   rsVehicle.CursorLocation = 3
   rsVehicle.Open q, connstr%>
   <br><br>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr >
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="right"><font face="Verdana" size="1">No</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Date</td></td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Target<br>Type</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Target<br>Code</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Job<br>Description</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Time</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Hours</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Unit</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Output</td>
   <td style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Amount</td>
   </tr>
   <%
   if Not rsVehicle.eof then
      nCountofField=rsVehicle.fields.count-1
	  no=1
	  nhours = nhours+cdbl(rsVehicle("hours"))
	  nunit = nunit+cdbl(rsVehicle("unit"))
	  noutput = noutput+cdbl(rsVehicle("output"))
	  nAmount = nAmount+cdbl(rsVehicle("Amount"))
	  While Not rsVehicle.EOF

		  'nQty = nQty+cdbl(rsVehicle("Quantity"))
		  'nAmount = nAmount+cdbl(rsVehicle("Amount"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr >
		  <td style="<%=cmdstyleLeft%>" align="left" valign="top"><font face="Verdana" size="1"><%=NO%></td>
          <td style="<%=cmdstyle%>" align="left" title="WS.ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsVehicle("VEHICLEDATE")%></td>
	      <td style="<%=cmdstyle%>" align="left" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsVehicle("locationtype")%></td>
	      <td style="<%=cmdstyle%>" align="left" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsVehicle("locationcode")%></td>
	      <td style="<%=cmdstyle%>" align="left" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsVehicle("fcname")%></td>
	      <td style="<%=cmdstyle%>" align="left" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsVehicle("time")%></td>
	      <td style="<%=cmdstyle%>" align="Right" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsVehicle("hours")%></td>
	      <td style="<%=cmdstyle%>" align="Right" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsVehicle("unit")%></td>
	      <td style="<%=cmdstyle%>" align="Right" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsVehicle("output")%></td>
	      <td style="<%=cmdstyle%>" align="Right" title="WS.Amount" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsVehicle("AMOUNT"),2)%></td>
		  </tr>
		  <%
		  rsVehicle.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  <tr >
	  <td style="<%=cmdstyleLeft%>" colspan="6" align="center"><font face="Verdana" size="1">T o t a l</td>
	  <td style="<%=cmdstyle%>" align="right"><font face="Verdana" size="1"><%=FormatNumber(nhours,2)%></td>
	  <td style="<%=cmdstyle%>" align="right"><font face="Verdana" size="1"><%=FormatNumber(nunit,2)%></td>
	  <td style="<%=cmdstyle%>" align="right"><font face="Verdana" size="1"><%=FormatNumber(noutput,2)%></td>
	  <td style="<%=cmdstyle%>" align="right"><font face="Verdana" size="1"><%=FormatNumber(nAmount,2)%></td>
	  </tr><%
   else
      response.Write("Data is empty")	  
   end if 
   rsVehicle.Close
   set rsVehicle = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->