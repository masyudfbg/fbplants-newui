<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - Stock Product Ledger</title>
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
v_txtSTORAGE_CODE=trim(request("txtSTORAGE_CODE"))
v_txtPRODUCT_CODE=trim(request("txtPRODUCT_CODE"))


txtSQL = "Select TDATE,STORAGE_CODE,PRODUCT_CODE, nvl(SUM(QUANTITY_IN),0) AS QUANTITY_IN, nvl(SUM(QUANTITY_OUT),0) AS QUANTITY_OUT," & _
		"REFERENCES from MILL_PRODUCTION where FCBA='" & Session("fcBA") & "'"
if  len(trim(v_txtyear)) > 0 then
    txtSQL = txtSQL & " and to_number(to_Char(TDATE,'yyyy'))=" & v_txtyear
end if
if  len(trim(v_txtmonth)) > 0 then
    txtSQL =txtSQL &  " and to_number(to_Char(TDATE,'mm'))=" & v_txtmonth
end if
if  len(trim(v_txtSTORAGE_CODE)) > 0 then
    txtSQL = txtSQL & " and STORAGE_CODE= '"& v_txtSTORAGE_CODE & "'"
end if
if  len(trim(v_txtPRODUCT_CODE)) > 0 then
    txtSQL =txtSQL &  " and PRODUCT_CODE='" & v_txtPRODUCT_CODE & "'"
end if 
txtsql=(txtsql) & " group by TDATE,STORAGE_CODE,PRODUCT_CODE,REFERENCES"
'response.Write(txtsql)
'response.Flush()
%>
 <table width="544">
  <tr> 
    <td rowspan="5" width="100"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Manager</td><td  width="300"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">STOCK PRODUCT LEDGER</FONT></td></tr>
 </table>
<BR><BR>
<table width="805" cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
  <tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
  </tr>
	<TR>
	<TD colspan="5" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF">&nbsp;</TD>
	<TD colspan="2" style="<%=cmdstyleTop%>" align="center"><font face="Verdana" size="1">Quantity</TD>
	</TR>
	<tr>
	    <td style="<%=cmdstyleLeft%>">NO</td>
		<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1">DATE</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">STORAGE CODE</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">PRODUCT CODE</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">REFERENCES</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">QUANTITY IN</td>
		<td align="center" style="<%=cmdstyle%>"> <font face="Verdana" size="1">QUANTITY OUT</td>
	</tr>
	
<%
	Dim v_Quantity_In,v_Quantity_out
	v_Quantity_In=0
	v_Quantity_out=0
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("STORAGE_CODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("PRODUCT_CODE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REFERENCES")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QUANTITY_IN"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("QUANTITY_OUT"),2)%></td>
		</tr>
	  <%
	  v_Quantity_In=v_Quantity_In+trim(cdbl(rsCat("QUANTITY_IN")))
	  v_Quantity_out=v_Quantity_out+trim(cdbl(rsCat("QUANTITY_OUT")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>	
	
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="5"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity_In,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity_out,2)%></td>
	</tr>
	
</table>
<!--#include file="Bottom_Report.asp"-->