<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<head>
<title>RW Plant</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
<script language="JavaScript" src="date-picker.js"></script> 
</head>
<body>
 <%
    Dim txtNoWS,txtDateWS,txtRemarks,q,nCountofField,no,i
    Dim Remarks
    Dim rsWS,intRecordsPerPage,cColor,txtAgreementDetail
	'txtAgreementDetail =cmdFindTitle("AgreementDetail","fcCode",request("AgreementDetailCode"))
    %>
	 <table width="544">
	  <tr> 
		<td rowspan="6" width="100">
	<% if len(trim(PROFIL("REPORT_IMAGE"))) > 0 then %>
		<img src="IMAGES/<%=PROFIL("REPORT_IMAGE")%>" width="210" height="80">
	<% end if %>
	</td>
	  </tr>
		<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
		<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
		<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
		<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">AGGREMENT DETAILS</FONT></td></tr>
		<tr><td width="100">Code</td><td  width="300"><b><FONT color="#000099"><%=Request("fccode")%></FONT></td></tr>
</table>

   <%
   Dim rsAgreementDetail,QTY,nRetention,RATE,nAmount
   'set rsAgreementDetail = server.CreateObject("Adodb.recordset")
   if request("fcCode") <> "" and request("fdDate") <> "" then
   q="select AGREEMENTCODE, LINENO, JOB,b.fcname, LOCATIONTYPE, LOCATIONCODE, to_char(DUEDATE,'mm/dd/yyyy') as DUEDATE, " & _
     "RETENTION, QTY, RATE from Agreementdetail a ,job b  where a.job=b.fcCode and " & _
	 " AGREEMENTCODE ='" & request("fcCode") & "'" & _
	 " and a.fcBA='" & session("fcBA") & "'"
     '" and to_char(DUEDATE,'mm/yyyy')=to_char(to_date('" & request("fdDate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy') " & _
   elseif request("fcCode")="" and request("fdDate")<> "" then
   q="select AGREEMENTCODE, LINENO, JOB,b.fcname, LOCATIONTYPE, LOCATIONCODE, to_char(DUEDATE,'mm/dd/yyyy') as DUEDATE, " & _
     "RETENTION, QTY, RATE from Agreementdetail a ,job b  where a.job=b.fcCode and " & _
	 " and a.fcBA='" & session("fcBA") & "'"
	 '" to_char(DUEDATE,'mm/yyyy')=to_char(to_date('" & request("fdDate") & "','" & "MM/DD/YYYY" & "'),'mm/yyyy') " & _
   elseif request("fcCode")<>"" and  request("fdDate")="" then
   q="select AGREEMENTCODE, LINENO,  JOB,b.fcname, LOCATIONTYPE, LOCATIONCODE, to_char(DUEDATE,'mm/dd/yyyy') as DUEDATE, " & _
     "RETENTION, QTY, RATE from Agreementdetail a ,job b  where a.job=b.fcCode and " & _
	 " AGREEMENTCODE ='" & request("fcCode") & "'" & _
	 " and a.fcBA='" & session("fcBA") & "'"
   end if
   response.Write q	
   Set rsAgreementDetail = Server.CreateObject("ADODB.Recordset")
   rsAgreementDetail.CursorType = 3
   rsAgreementDetail.CursorLocation = 3
   rsAgreementDetail.Open q, connstr%>
   <br><br>
   
   <table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr>
   <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF">
   <font face="Verdana" size="1">No</td>	   
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Line</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Job.Code</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Job.Desc</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">TT</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">TC</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Due Date</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Retention</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Quantity</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Rate</td>
   </tr>
   <%
   if Not rsAgreementDetail.eof then
      nCountofField=rsAgreementDetail.fields.count-1
	  no=1
	  nRetention = 0
	  QTY = 0
	  RATE = 0
	  While Not rsAgreementDetail.EOF

	      nRetention = nRetention+cdbl(rsAgreementDetail("Retention"))
	      QTY = QTY+cdbl(rsAgreementDetail("QTY"))
	      RATE = RATE+cdbl(rsAgreementDetail("RATE"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
	      <td align="left" title="WS.LINENO" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LINENO")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("JOB")%></td>
	      <td align="left" title="WS.fcname" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("fcname")%></td>
	      <td align="left" title="WS.LOCATIONTYPE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LOCATIONTYPE")%></td>
	      <td align="left" title="WS.LOCATIONCODE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LOCATIONCODE")%></td>
	      <td align="left" title="WS.DUEDATE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("DUEDATE")%></td>
	      <td align="Right" title="WS.RETENTION" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsAgreementDetail("RETENTION"),2)%></td>
	      <td align="Right" title="WS.QTY" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsAgreementDetail("QTY"),2)%></td>
	      <td align="Right" title="WS.RATE" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsAgreementDetail("RATE"),2)%></td>
		  </tr>
		  <%
		  rsAgreementDetail.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	  <tr bgcolor="#CCCC99">
	  <td colspan="7" align="center" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1">T o t a l</td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nRetention,2)%></td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(QTY,2)%></td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(RATE,2)%></td>
	  </tr><%
   else
      response.Write("Data is empty")	  
   end if 
   rsAgreementDetail.Close
   set rsAgreementDetail = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->