<%'Option Explicit
%>
<%' On Error GOTO CmdErrorlog 
%>
<%response.buffer=True
Dim cmdstyle,cmdstyleTop,cmdstyleLeft
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<html>
<head>
<title>SAC Plant</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
</head>
<body>
 <%
 Dim txtNoGAD,txtDateGAD,txtRemarks,q,nCountofField,no,i
 Dim Remarks

   Dim rsGAD,intRecordsPerPage,cColor,txtSupplier
   set rsGAD = server.CreateObject("Adodb.recordset")
   q="select * from Attendance_GAD_temp " & _
     " where GangCode='" & request("GangCode") & "' and to_char(fdDate,'mm/dd/yyyy')='" & Request("fdDate") & "' and " & _
     " fcBA='" & session("fcBA") & "'"
   'response.Write q	
   Set rsGAD = Server.CreateObject("ADODB.Recordset")
   rsGAD.Open q, connstr
   if not rsGAD.eof then
      txtNoGAD = rsGAD("GangCode")
	  txtDateGAD = rsGAD("fdDATE")
   End if
   rsGAD.close
   set rsGAD = Nothing%>
 
<table width="804">
  <tr> 
    <td rowspan="6" width="100">
	<% if len(trim(PROFIL("REPORT_IMAGE"))) > 0 then %>
		<img src="IMAGES/<%=PROFIL("REPORT_IMAGE")%>" width="210" height="80">
	<% end if %>
	</td>
  </tr>
	<tr><td width="112"><FONT face="Verdana" size="1"color="#000099">Company</td><td  width="403"><b><FONT face="Verdana" size="1"color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="112"><FONT face="Verdana" size="1"color="#000099">Business Unit</td><td  width="403"><b><FONT face="Verdana" size="1"color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="112"><FONT face="Verdana" size="1"color="#000099">Report</td><td  width="403"><b><FONT face="Verdana" size="1" color="#000099">GANG ACTIVITY DETAILS & ATTENDANCE</FONT></td></tr>
	<tr><td width="112"><FONT face="Verdana" size="1"color="#000099">Date</td><td  width="403"><b><FONT face="Verdana" size="1" color="#000099"><%=FormatDateTime(Request("fdDate"),1)%></FONT></td></tr>
	<tr><td width="112"><FONT face="Verdana" size="1"color="#000099">Gang Code</td><td  width="403"><b><FONT face="Verdana" size="1" color="#000099"><%=Request("GangCode")%></FONT></td></tr>
 </table>
   <%
   Dim rsPurchaseOrder,nQty,nAmount
   set rsPurchaseOrder = server.CreateObject("Adodb.recordset")
   q="select att.GANGCODE,att.FDDATE,att.SUPERVISION_1,att.EMPLOYEECODE,emp.FCNAME as empName,att.Jobcode," & _
     "j.FCNAME as JobDesc,att.LOCATIONTYPE,att.ATTENDANCE,att.LOCATIONCODE,nvl(sum(mandays),0) as Mandays,nvl(sum(att.unit),0) as Unit,nvl(sum(att.OUTPUT),0) as Output," & _
	 "nvl(sum(othrs),0)as OTHRS,nvl(sum(rate),0) as Rate,att.REMARKS,att.REFERENCE from  attendance_gad_temp att,employee emp,job j " & _
     "where  (att.EMPLOYEECODE = emp.fcCode and att.fcba=emp.fcba) and (att.jobcode = j.fccode and att.fcba=j.fcba) and " & _
	 "att.GangCode='" & request("GangCode") & "' and  fddate = to_date('" & Request("fdDate") & "','mm/dd/yyyy') "  & _
     "and att.fcBA='" & session("fcBA") & "'" & _
     "group by att.GANGCODE,att.FDDATE,att.SUPERVISION_1,att.EMPLOYEECODE,emp.FCNAME ,att.jobcode,att.GANGCODE," & _
     "j.FCNAME ,att.LOCATIONTYPE,att.ATTENDANCE,att.LOCATIONCODE," & _
	 "att.REMARKS,att.REFERENCE " & _ 
	 "ORDER BY att.Jobcode,att.LOCATIONCODE,emp.FCNAME " 
	 

   'response.Write q	
   Set rsPurchaseOrder = Server.CreateObject("ADODB.Recordset")
   rsPurchaseOrder.CursorType = 3
   rsPurchaseOrder.CursorLocation = 3
   rsPurchaseOrder.Open q, connstr%>
   <br><br>
   <table Cellpadding="2" style="pointsize:8px" cellspacing="0">
   <tr>
   <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="center"><font face="Verdana" size="1">No</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">NIK</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Name</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Code</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Aktivitas</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Absensi</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Target<br>Type</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Target<br>Code</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Keterangan</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Reference</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">HK</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Unit</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Output</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Lembur</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Rate</td>
   </tr>
   <%
   if Not rsPurchaseOrder.eof then
      nCountofField=rsPurchaseOrder.fields.count-1
      Dim nMandays,nUnit,nOutput,nOTHRS,nRate
      nMandays=0
	  nUnit=0
	  nOutput=0
	  nOTHRS=0
	  nRate=0
	  No=1
	  While Not rsPurchaseOrder.EOF

	      nMandays=nMandays+CheckValue(rsPurchaseOrder("Mandays"))
	      nUnit=nUnit+CheckValue(rsPurchaseOrder("Unit"))
	      nOutput=nOutput+CheckValue(rsPurchaseOrder("Output"))
	      nOTHRS=nOTHRS+CheckValue(rsPurchaseOrder("OTHRS"))
	      nRate=nRate+CheckValue(rsPurchaseOrder("Rate"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr>
		  <td align="left"  style="<%=cmdstyleLeft%>" valign="top"><font face="Verdana" size="1"><%=NO%></td>
	      <td align="left" title="GAD.ITEMCODE" valign="top" style="<%=cmdstyle%>"><font pointsize="8px" face="Georgia, Times New Roman, Times, serif"><font face="Verdana" size="1"><%=rsPurchaseOrder("EMPLOYEECODE")%></font></td>
	      <td align="left" title="s.fcname" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPurchaseOrder("empName")%>&nbsp;</td>
	      <td align="left" title="GAD.OTHERITEMDESC" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPurchaseOrder("JobCode")%>&nbsp;</td>
	      <td align="left" title="GAD.OTHERITEMDESC" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPurchaseOrder("JobDesc")%>&nbsp;</td>
	      <td align="left" title="GAD.OTHERITEMDESC" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPurchaseOrder("ATTENDANCE")%>&nbsp;</td>
	      <td align="left" title="GAD.Quantity" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPurchaseOrder("LOCATIONTYPE")%>&nbsp;</td>
	      <td align="left" title="GAD.Quantity" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPurchaseOrder("LOCATIONCODE")%>&nbsp;</td>
	      <td align="left" title="GAD.Quantity" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPurchaseOrder("REMARKS")%>&nbsp;</td>
	      <td align="left" title="GAD.Quantity" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsPurchaseOrder("REFERENCE")%>&nbsp;</td>
	      <td align="Right" title="GAD.Amount" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(rsPurchaseOrder("Mandays"))%>&nbsp;</td>
	      <td align="Right" title="GAD.Amount" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(rsPurchaseOrder("Unit"))%>&nbsp;</td>
	      <td align="Right" title="GAD.Amount" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(rsPurchaseOrder("Output"))%>&nbsp;</td>
	      <td align="Right" title="GAD.Amount" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(rsPurchaseOrder("OTHRS"))%>&nbsp;</td>
	      <td align="Right" title="GAD.Amount" valign="top" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(rsPurchaseOrder("Rate"))%>&nbsp;</td>
		  </tr>
		  <%
		  rsPurchaseOrder.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  <tr>
	  <td colspan="10" style="<%=cmdstyleLeft%>" valign="top" align="center"><font face="Verdana" size="1">T o t a l&nbsp;</td>
	  <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(nMandays)%>&nbsp;</td>
	  <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(nUnit)%>&nbsp;</td>
	  <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(nOutput)%>&nbsp;</td>
	  <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(nOTHRS)%>&nbsp;</td>
	  <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=CheckValue(nRate)%>&nbsp;</td>
	  </tr><%
   else
      response.Write("Data is empty")	  
   end if 
   rsPurchaseOrder.Close
   set rsPurchaseOrder = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->