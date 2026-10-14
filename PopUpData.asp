<%'Option Explicit
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>iPS - Popup</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
</head>
<body>
<table class="entry">
<tr>
 
 <table width="700">
  <tr> 
    <td rowspan="5" width="100"><img src="IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099"><%=request("txtModule")%>/<%=request("txtTable")%></FONT></td></tr>
 </table>
<tr>
 <td>
    <%' -----------------ATAS ISI%>
		    <!--#include file="browsetable.asp"-->
	</b><br>
	<form action="show.asp" method="post">
	<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	    <tr>
		<td align="center" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">No.</td>
	    <%	
		'response.Write(q)
		response.Flush()
		Set rsCat = connStr.Execute(q)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1
		if txtTable="LocationType" then
		   nCountofField=nCountofField+4
		end if%>
				<td align="center" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">CODE</td>
				<td align="center" style="<%=cmdstyleTop%>"><font face="Verdana" size="1">NAME</td>		
			<%for i = 2 to nCountofField-4%>
				<td align="center" style="<%=cmdstyleTop%>"><font face="Verdana" size="1"><%=rsCat(i).Name%></td>
			<%Next%>
		</tr>
		<%While Not rsCat.EOF%>
			<tr>
			<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
			<%for i = 0 to nCountofField-4
			    if len(trim(rscat(i).value))>0 then%>
				   <td style="<%=cmdstyle%>" align="left"><font face="Verdana" size="1">
				   <%if i =1 Then%>
				        <a href="#" onClick='self.opener.document.<%=request("p_item")%>.value="<%=rscat.Fields(0)%>";window.close()'><%=rsCat(i)%></a></td>
				   <%else%>
				        <%=rsCat(i)%></td>		
				   <%End if%>
				<%else%>
				   <td style="<%=cmdstyle%>" align="left"><font face="Verdana" size="1">-</td>
				<%end if%>
			<%next%>
			</tr>
		  <%
		  rsCat.MoveNext
		  no=no+1
		Wend
		rsCat.close()
		rsCat=null%>
		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		</form>
	</table>
  <%'------------------ BAWAH ISI%>
  </td>
</tr>
<!--#include file="Bottom_Report.asp"-->