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

v_PROJECTCODE=request("v_PROJECTCODE")

txtSQL = "select PROJECTCODE, ISSUEDATE, LOCATIONTYPE,LOCATIONCODE,JOBCODE,s.fcname,nvl(sum(QUANTITY),0) as Quantity,nvl(SUM(AMOUNT),0) AS AMOUNT  from PROJECT_ISSUE NO, JOB S where  " & _
         " (no.JOBCODE = S.FCCODE and no.fcba = s.fcba) and " & _
         " PROJECTCODE='" & v_PROJECTCODE & _
		  "' and no.fcBA='" & Session("fcBA") & "' group by PROJECTCODE, ISSUEDATE, LOCATIONTYPE,LOCATIONCODE,JOBCODE,s.fcname"		  
'response.Write(txtsql)


response.Flush()
%>
 <table width="700">
  <tr> 
    <td rowspan="5" width="100"><img src="IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PROJECT ISSUE</FONT></td></tr>
 </table><BR><BR>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
		<tr>
		<%	
		Set rsCat = connStr.Execute(txtSQL)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1%>
		</tr>
		<tr>
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">NO</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">PROJECT CODE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">ISSUE DATE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">TARGET TYPE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">TARGET CODE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">JOB</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">JOB DESCRIPTION</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QUANTITY</td>
	</tr>
    <% Dim nQty	  
	Set rsCat = Server.CreateObject("ADODB.Recordset")
   rsCat.CursorType = 3
   rsCat.CursorLocation = 3
   rsCat.Open txtSql, connstr
   if Not rsCat.eof then
	nCountofField=rsCat.fields.count-1
	no=1
	nQty=0
	nAmount=0
	While Not rsCat.EOF%>
			<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PROJECTCODE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ISSUEDATE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONTYPE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONCODE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JOBCODE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%>&nbsp;</td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("QUANTITY"),2)%>&nbsp;</td>
			</tr>
		  <%
				nQty=nQty+trim(cdbl(rsCat("QUANTITY")))
    		  rsCat.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="7"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nQty,2)%>&nbsp;</td>
	</tr>	  
   <%
   else
      response.Write("Data is empty")	  
   end if 
   rsCat.Close
   set rsCat = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->