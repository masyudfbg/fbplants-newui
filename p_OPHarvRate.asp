<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - OP Harvesting Rate</title>
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


v_txtCROPCODE= trim(request("CROPCODE"))
v_txtTERRAIN_TYPE= trim(request("TERRAIN_TYPE"))

txtSQL = "SELECT CROPCODE, ABW_CENSUS, TERRAIN_TYPE, ABW_MIN, ABW_MAX, BASIS_QTY, BASIS_RATE, EXTRA_BASIS_RATE FROM OPHARVRATE" & _
		 " where fcba='" & session("fcBA") & "'"
if  len(trim(v_txtCROPCODE)) > 0 then
    txtSQL = txtSQL & " and CROPCODE='" & v_txtCROPCODE& "'"
end if
if  len(trim(v_txtTERRAIN_TYPE)) > 0 then
    txtSQL = txtSQL & " and TERRAIN_TYPE='" & v_txtTERRAIN_TYPE& "'"
end if
txtsql=(txtsql) 
'& " group by WIDATE, WIDIVISION"

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
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">OP HARVESTING RATE</FONT></td></tr>
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
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">CROP CODE</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">ABW AVERAGE</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">TERRAIN TYPE</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">ABW MIN</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">ABW MAX</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">BASIS QTY</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">BASIS RATE</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">EXTRA BASIS RATE</td>
	</tr>
	<%
	Dim v_TOTRNINCHES,v_TOTRNDAYS,v_AVGRAINDAY,v_AVGTEMP,v_AVGWIND
		'v_TOTRNINCHES=0
		'v_TOTRNDAYS=0
		'v_AVGRAINDAY=0
		'v_AVGTEMP=0
		'v_AVGWIND=0				
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("CROPCODE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ABW_CENSUS")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("TERRAIN_TYPE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ABW_MIN")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ABW_MAX")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BASIS_QTY")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("BASIS_RATE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("EXTRA_BASIS_RATE")%>&nbsp;</td>
		</tr>
	  <%
		'v_TOTRNINCHES=v_TOTRNINCHES+trim(cdbl(rsCat("TOTRNINCHES")))
		'v_TOTRNDAYS=v_TOTRNDAYS+trim(cdbl(rsCat("TOTRNDAYS")))
		'v_AVGRAINDAY=v_AVGRAINDAY+trim(cdbl(rsCat("AVGRAINDAY")))
		'v_AVGTEMP=v_AVGTEMP+trim(cdbl(rsCat("AVGTEMP")))
		'v_AVGWIND=v_AVGWIND+trim(cdbl(rsCat("AVGWIND")))		

	  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
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
	
</table>
<!--#include file="Bottom_Report.asp"-->