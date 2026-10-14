<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - Project Adjusment</title>
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
v_PROJECTCode=trim(request("PROJECTCode"))


v_Crop=trim(request("Crop"))
v_Source=trim(request("Source"))


txtSQL = "select * from PROJECT_ADJUSMENT where fcBA='" & Session("fcBA") & "'"

if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and ADJUSMENTDATE >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtPODATE_From,1)
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and ADJUSMENTDATE <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
	txtSelect = " To " & FormatDateTime(v_txtPODATE_To,1)
end if
if  len(trim(v_PROJECTCode)) > 0 then
    txtSQL = txtSQL & " and PROJECTCode='" & v_PROJECTCode & "'"
	txtSelect = txtSelect & " PROJECTCode " & v_txtPSTORAGE_STORAGECODE
end if
txtsql=(txtsql)
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

'response.Write(txtsql)
response.Flush()
%>
 <table width="545">
  <tr> 
    <td rowspan="6" width="100"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PROJECT ADJUSMENT</FONT></td></tr>
	<tr><td width="100">Code</td><td  width="300"><b><FONT color="#000099"><%=v_PROJECTCode%></FONT></td></tr>
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
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">No</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Project Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Date Adjustment</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Quantity</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Amount</td>
	</tr>
    <%Dim nQty
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
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PROJECTCode")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ADJUSMENTDATE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FORMATNUMBER(rscat("QUANTITY"),2)%>&nbsp;</td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FORMATNUMBER(rscat("AMOUNT"),2)%>&nbsp;</td>
			</tr>
		  <%
		  nQty=nQty+trim(cdbl(rsCat("QUANTITY")))
		  nAmount=nAmount+trim(cdbl(rsCat("Amount")))
		  rsCat.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="3"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nQty,2)%>&nbsp;</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nAmount,2)%>&nbsp;</td>
	</tr>

   <%
   else
      response.Write("Data is empty")	  
   end if 
   rsCat.Close
   set rsCat = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->