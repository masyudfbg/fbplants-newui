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

   Dim rsWS,intRecordsPerPage,cColor,txtWorkshop
   set rsWS = server.CreateObject("Adodb.recordset")
   q="select * from WorkshopActivity  " & _
     " where WorkshopCode='" & request("WorkshopCode") & "' and to_char(WORKSHOPDATE,'mm/dd/yyyy')='" & Request("fdDate") & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsWS = Server.CreateObject("ADODB.Recordset")
   rsWS.Open q, connstr
   if not rsWS.eof then
      txtNoWS = rsWS("WorkshopCode")
	  txtDateWS = rsWS("WORKSHOPDATE")
	  txtWorkshop =cmdFindTitle("Workshop","fcCode",request("WorkshopCode"))
   End if
   rsWS.close
   set rsWS = Nothing%>
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
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">WORKSHOP ACTIVITY</FONT></td></tr>
	<tr><td width="100">Date</td><td  width="300"><b><FONT color="#000099"><%=Request("fdDate")%></FONT></td></tr>
	<tr><td width="100">Workshop</td><td  width="300"><b><FONT color="#000099"><%=Request("WorkshopCode")%></FONT></td></tr>
 </table>
   <%
   Dim rsWorkshop,nUnit,nHours,nOutput,nAmount
   set rsWorkshop = server.CreateObject("Adodb.recordset")
   q="select ws.WORKSHOPDATE, ws.LOCATIONTYPE, ws.LOCATIONcode, j.fcname,ws.time," & _
     "sum(ws.HOURS) as Hours,sum(ws.UNIT) as Unit,sum(ws.OUTPUT) as Output,sum(AMOUNT) as Amount" & _
     " from WorkshopActivity ws,job j " & _
     " where (ws.job = j.fccode and ws.fcba = j.fcba) and ws.WorkshopCode='" & request("WorkshopCode") & "' and to_char(ws.WORKSHOPDATE,'mm/dd/yyyy')='" & Request("fdDate") & "' and " & _
     " ws.fcBA='" & session("fcBA") & "'" & _
	 "group by ws.WORKSHOPDATE, ws.LOCATIONTYPE, ws.LOCATIONcode, j.fcname,ws.time"
   'response.Write q	
   Set rsWorkshop = Server.CreateObject("ADODB.Recordset")
   rsWorkshop.CursorType = 3
   rsWorkshop.CursorLocation = 3
   rsWorkshop.Open q, connstr%>
   <br><br>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr >
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="right"><font face="Verdana" size="1">No</td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Date</td></td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Location<br>Type</td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Location<br>Code</td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Job<br>Description</td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Time</td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Hours</td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Unit</td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Output</td>
   <td  style="<%=cmdstyleTOP%>" align="center"><font face="Verdana" size="1">Amount</td>
   </tr>
   <%
   if Not rsWorkshop.eof then
      nCountofField=rsWorkshop.fields.count-1
	  no=1
	  nhours=0
	  nunit = 0
	  noutput =0
	  nAmount=0
	  While Not rsWorkshop.EOF

		  nhours = nhours+cdbl(rsWorkshop("hours"))
		  nunit = nunit+cdbl(rsWorkshop("unit"))
		  noutput = noutput+cdbl(rsWorkshop("output"))
		  nAmount = nAmount+cdbl(rsWorkshop("Amount"))
		  
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr >
		  <td  style="<%=cmdstyleLeft%>" align="left" valign="top"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
          <td  style="<%=cmdstyle%>" align="left" title="Date" valign="top"><font face="Verdana" size="1"><%=rsWorkshop("WORKSHOPDATE")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>" align="left" title="locationtype" valign="top"><font face="Verdana" size="1"><%=rsWorkshop("locationtype")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>" align="left" title="locationcode" valign="top"><font face="Verdana" size="1"><%=rsWorkshop("locationcode")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>" align="left" title="fcname" valign="top"><font face="Verdana" size="1"><%=rsWorkshop("fcname")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>" align="left" title="time" valign="top"><font face="Verdana" size="1"><%=rsWorkshop("time")%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>" align="Right" title="hours" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsWorkshop("hours"),2)%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>" align="Right" title="unit" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsWorkshop("unit"),2)%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>" align="Right" title="output" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsWorkshop("output"),2)%>&nbsp;</td>
	      <td  style="<%=cmdstyle%>" align="Right" title="Amount" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsWorkshop("AMOUNT"),2)%>&nbsp;</td>
		  </tr>
		  <%
		  rsWorkshop.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="6"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nhours,2)%>&nbsp;</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nunit,2)%>&nbsp;</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(noutput,2)%>&nbsp;</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nAmount,2)%>&nbsp;</td>
	</tr>
	  
	  <%
   else
      response.Write("Data is empty")	  
   end if 
   rsWorkshop.Close
   set rsWorkshop = Nothing%>
   </TABLE>
 <!--#include file="Bottom_Report.asp"-->