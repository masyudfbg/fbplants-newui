<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - Project Ledger</title>
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

v_txtyear=trim(request("txtyear"))
v_txtmonth=trim(request("txtmonth"))
v_txtprojectcode=trim(request("txtprojectcode"))

txtSQL="SELECT PROJECTCODE, YEAR, MONTH, nvl(sum(OPENINGQTY),0) AS OPENINGQTY, nvl(sum(OPENINGAMOUNT),0) as OPENINGAMOUNT, " & _
        "nvl(sum(QTYRECEIVED),0) as QTYRECEIVED, nvl(sum(QTYCULLED),0) as QTYCULLED, " & _
		"nvl(sum(QTYISSUED),0) as QTYISSUED, nvl(sum(COSTPERSEEDLING),0) as COSTPERSEEDLING, nvl(sum(AMOUNTRECEIVED),0) as AMOUNTRECEIVED, " & _
		"nvl(sum(AMOUNTISSUED),0) as AMOUNTISSUED, " & _
		"nvl(sum(AMOUNTADJUSTMENT),0) as AMOUNTADJUSTMENT, " & _
		"nvl((sum(OPENINGQTY)+sum(QTYRECEIVED)-sum(QTYCULLED)-sum(QTYISSUED)),0) as QtySaldo," & _ 		 
         "nvl((sum(OPENINGAMOUNT)+sum(AMOUNTRECEIVED)-sum(AMOUNTISSUED)-sum(AMOUNTadjustment) ),0) as RpSaldo" & _ 		 
		 " from PROJECT_LEDGER  where fcBA='" & Session("fcBA") & "'"

if  len(trim(v_txtyear)) > 0 then
    txtSQL = txtSQL & " and year=" & v_txtyear
end if
if  len(trim(v_txtmonth)) > 0 then
    txtSQL =txtSQL &  " and Month=" & v_txtmonth
end if
if  len(trim(v_txtprojectcode)) > 0 then
    txtSQL =txtSQL &  " and PROJECTCODE like '" & v_txtprojectcode & "'"
end if 
txtsql=(txtsql) & " group by PROJECTCODE, YEAR, MONTH"
'response.Write(txtsql)
if v_txtprojectcode="" then
   v_txtprojectcode="ALL"
end if
response.Flush()
%>
 <table width="542">
  <tr> 
    <td rowspan="6" width="100"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PROJECT LEDGER</FONT></td></tr>
	<tr><td width="100">Period</td><td  width="300"><b><FONT color="#000099"><%=textMonth(v_txtmonth)%> , <%=v_txtyear%></FONT></td></tr>
	<tr><td width="100">Code</td><td  width="300"><b><FONT color="#000099"><%=v_txtprojectcode%></FONT></td></tr>
 </table>
<BR><BR>
<table width="200" cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
	<TR>
	<TD colspan="2" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF">&nbsp;</TD>
	<TD colspan="5" style="<%=cmdstyleTop%>" align="center"><font face="Verdana" size="1">Quantity</TD>
	<TD colspan="6" style="<%=cmdstyleTop%>" align="center"><font face="Verdana" size="1">Amount (Rp)</TD>
	</TR>
	<tr>
	    <td style="<%=cmdstyleLeft%>"><font face="Verdana" size="1">NO</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Project<br>Code</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Opening</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Order</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Adjustment</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Issue</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Saldo</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Opening</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Received</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Cost Per<br>Project</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Adjustment</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Issue</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">Saldo</td>
	</tr>
	<%While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsCat("PROJECTCODE")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("OPENINGQTY"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYRECEIVED"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYCULLED"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QTYISSUED"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("QtySaldo"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("OPENINGAMOUNT"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMOUNTRECEIVED"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("COSTPERSEEDLING"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMOUNTADJUSTMENT"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("AMOUNTISSUED"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rsCat("RpSaldo"),2)%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	  response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
</table>
<!--#include file="Bottom_Report.asp"-->