<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - Project Issue</title>
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

v_txtPeriodDate=REQUEST("PeriodDate")

v_txtPODATE_From=trim(request("DATEACTIVITY_From"))
v_txtPODATE_To=trim(request("DATEACTIVITY_To"))
v_ProjectCode=trim(request("ProjectCode"))
v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
v_LOCATIONCODE=trim(request("LOCATIONCODE"))

txtSQL = "select ProjectCODE, ISSUEDATE, LOCATIONTYPE,LOCATIONCODE,JOBCODE,s.fcname,nvl(sum(QUANTITY),0) as Quantity,nvl(SUM(AMOUNT),0) AS AMOUNT  from PROJECT_ISSUE NO, JOB S where  " & _
         " (no.JOBCODE = S.FCCODE and no.fcba = s.fcba) and " & _
		  " no.fcBA='" & Session("fcBA") & "'"

if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and ISSUEDATE >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and ISSUEDATE <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
end if
if  len(trim(v_ProjectCode)) > 0 then
    txtSQL = txtSQL & " and ProjectCode='" & v_ProjectCode & "'"
end if
if  len(trim(v_LOCATIONTYPE)) > 0 then
    txtSQL = txtSQL & " and LOCATIONTYPE='" & v_LOCATIONTYPE & "'"
end if
if  len(trim(v_LOCATIONCODE)) > 0 then
    txtSQL = txtSQL & " and LOCATIONCODE='" & v_LOCATIONCODE & "'"
end if

txtSQL=(txtSQL) & " group by ProjectCODE, ISSUEDATE, LOCATIONTYPE,LOCATIONCODE,JOBCODE,s.fcname"
'response.Write(txtsql)
response.Flush()
%>
 <table width="543">
  <tr> 
    <td rowspan="6" width="100"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PROJECT ISSUE</FONT></td></tr>
 </table>
<BR><BR>
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
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">PRODUCTION CODE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">ISSUE DATE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">TARGET TYPE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">TARGET CODE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">JOB</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">JOB DESCRIPTION</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">QUANTITY</td>
	</tr>
    <%
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
	  		<tr>
			<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ProjectCODE")%>&nbsp;</td>
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