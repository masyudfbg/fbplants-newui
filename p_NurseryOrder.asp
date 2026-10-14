<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - Nursery Order</title>
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
v_NurseryCode=trim(request("NurseryCode"))

txtSQL = "select NURSERYCODE, ORDERDATE, SUPPLIERCODE, s.fcname,nvl(sum(QUANTITY),0) as Quantity from NurseryOrder NO, Supplier S where  " & _
         " (no.SUPPLIERCODE = S.FCCODE and no.fcba = s.fcba) and " & _
		  " no.fcBA='" & Session("fcBA") & "'"

if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and ORDERDATE >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and ORDERDATE <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
end if
if  len(trim(v_NurseryCode)) > 0 then
    txtSQL = txtSQL & " and NurseryCode='" & v_NurseryCode & "'"
end if

txtSQL=(txtSQL) & " group by NURSERYCODE, ORDERDATE, SUPPLIERCODE, s.fcname"
'response.Write(txtsql)
response.Flush()
%>
 <table width="548">
  <tr> 
    <td rowspan="6" width="100"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">NURSERY ORDER</FONT></td></tr>
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
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">BATCH</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">ORDER DATE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">SUPPLIER CODE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">SUPPLIER NAME</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">ORDER QUANTITY</td>
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
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NURSERYCODE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ORDERDATE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SUPPLIERCODE")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FCNAME")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Quantity")%>&nbsp;</td>
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