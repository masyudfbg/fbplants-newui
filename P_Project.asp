<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - Project</title>
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


v_Crop=trim(request("Crop"))
v_Source=trim(request("Source"))

if request("txtKind") = "ProjectList" Then
	txtSQL = "select * from Project where fcBA='" & Session("fcBA") & "'"
Else
	txtSQL = "select * from Project where fcBA='" & Session("fcBA") & "'"
End if

if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and DATEPLANTED >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtPODATE_From,1)
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and DATEPLANTED <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtPODATE_To,1)
end if
if  len(trim(v_ProjectCode)) > 0 then
    txtSQL = txtSQL & " and ProjectCode='" & v_ProjectCode & "'"
	txtSelect = txtSelect & " Project Code " & v_ProjectCode
end if
if  len(trim(v_Crop)) > 0 then
    txtSQL =txtSQL &  " and Crop='" & v_Crop & "'"
	txtSelect = txtSelect & " Crop " & v_Crop
end if 
if  len(trim(v_Source)) > 0 then
    txtSQL =txtSQL &  " and Source='" & v_Source & "'"
	txtSelect = txtSelect & " Source " & v_Source
end if 

'response.Write(txtsql)
response.Flush()
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
%>
 <table width="700">
  <tr> 
    <td rowspan="6" width="100"><img src="IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	
	<%if request("txtKind") ="ProjectList" Then%>
	
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PROJECT LIST</FONT></td></tr>
	
	<%Else%>
	
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PROJECT INFORMATION</FONT></td></tr>
	
	<%End if%>
 </table>
<BR><BR>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
    <td colspan="10" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" ><font face="Verdana" size="1"><%=txtSelect%></td>
	</tr>
	<tr>
    <td colspan="10">&nbsp;</td>
	</tr>
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
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Project<br>Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Description</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Date<br>Planted</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Crop</td>	
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Planting<br>Material</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Clone</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Source</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Type</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Status</td>
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
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ProjectCode")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DESCRIPTION")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DATEPLANTED")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CROP")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PLANTINGMATERIAL")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CLONE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SOURCE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TYPE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STATUS")%>&nbsp;</td>
			</tr>
		  <%
		  rsCat.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
   <%
   else
      response.Write("Data is empty")	  
   end if 
   rsCat.Close
   set rsCat = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->