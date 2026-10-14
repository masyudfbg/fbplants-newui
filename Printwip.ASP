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
		<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
		<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">WORK in PROGRESS</FONT></td></tr>
		<tr><td width="100">Code</td><td  width="300"><b><FONT color="#000099"><%=Request("fccode")%></FONT></td></tr>
	 </table>

   <%
   Dim rsAgreementDetail,QTY,nRetention,RATE,nAmount
   set rsAgreementDetail = server.CreateObject("Adodb.recordset")
   q="select wip.AGREEMENTCODE,to_char(wip.WIPDATE,'mm/dd/yyyy') as WIPDATE, wip.AGREEMENTLINENO, ag.LOCATIONTYPE,ag.LOCATIONCODE,ag.job,job.FCNAME,sum(wip.QTY) as QTY " & _
     " from WORKINPROGRESS wip ,AGREEMENTDETAIL AG,JOB  where (wip.AGREEMENTCODE=AG.AGREEMENTCODE and wip.AGREEMENTLINENO=ag.LINENO and wip.fcba=ag.fcba)" & _
	 " and (ag.job = job.FCCODE and ag.fcba = job.fcba) " & _
	 " and wip.AGREEMENTCODE ='" & request("fcCode") & "'" & _
	 " and wip.fcBA='" & session("fcBA") & "' group by wip.AGREEMENTCODE,to_char(wip.WIPDATE,'mm/dd/yyyy'), wip.AGREEMENTLINENO, ag.LOCATIONTYPE,ag.LOCATIONCODE,ag.job,job.FCNAME"
   'response.Write q	
   Set rsAgreementDetail = Server.CreateObject("ADODB.Recordset")
   rsAgreementDetail.CursorType = 3
   rsAgreementDetail.CursorLocation = 3
   rsAgreementDetail.Open q, connstr%>
   <br><br>
   <table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr>   
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">NO</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">WIP Date</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Line</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">TT</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">TC</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Job Code</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Job Desc.</td>
   <td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Qty</td>
   </tr>
   <%
   if Not rsAgreementDetail.eof then
      nCountofField=rsAgreementDetail.fields.count-1
	  no=1
	  QTY = 0
	  While Not rsAgreementDetail.EOF

		  Qty = Qty+cdbl(rsAgreementDetail("qty"))
		  'nAmount = nAmount+cdbl(rsAgreementDetail("Amount"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr bgcolor="<%=cColor%>">
		  <td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
	      <td align="left" title="WS.LINENO" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("WIPDATE")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("AGREEMENTLINENO")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LOCATIONTYPE")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("LOCATIONCODE")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("job")%></td>
	      <td align="left" title="WS.JOB" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsAgreementDetail("FCNAME")%></td>
	      <td align="Right" title="WS.QTY" valign="top"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsAgreementDetail("QTY"),2)%></td>
		  </tr>
		  <%
		  rsAgreementDetail.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  
		  no=no+1
	  Wend%>
	  <tr>
	  <td colspan="7" align="center" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1">T o t a l</td>
	  <td align="right"style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(QTY,2)%></td>
	  </tr><%
   else
      response.Write("Data is empty")	  
   end if 
   rsAgreementDetail.Close
   set rsAgreementDetail = Nothing%>
<!--#include file="Bottom_Report.asp"-->