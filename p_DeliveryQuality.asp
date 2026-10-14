<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<title>RW Plant Reporting - Delivery Quality</title>
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

v_txtDELIVERYNOTE= trim(request("DELIVERYNOTE"))
v_txtQUALITYDATE_From= trim(request("QUALITYDATE_From"))
v_txtQUALITYDATE_To= trim(request("QUALITYDATE_To"))



txtSQL = "SELECT DELIVERYNOTE, QUALITYDATE, FFA, MOISTURE, DIRT, REMARKS, DIRT_AND_MOISTURE FROM DELIVERYQUALITY" & _
		 " where fcba='" & session("fcBA") & "'"

if  len(trim(v_txtQUALITYDATE_From)) > 0 then
    txtSQL = txtSQL & " and QUALITYDATE >=to_Date('" & v_txtQUALITYDATE_From & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtQUALITYDATE_To)) > 0 then
    txtSQL =txtSQL &  " and QUALITYDATE <=to_Date('" & v_txtQUALITYDATE_To & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtDELIVERYNOTE)) > 0 then
    txtSQL = txtSQL & " and DELIVERYNOTE='" & v_txtDELIVERYNOTE & "'"
end if


txtsql=(txtsql) 
'& " group by CHITNO, FIELDCODE, CROPCODE, ProductCODE, COLLECTIONDATE,HARVESTINGDATE,RECEPTIONDATE, " & _
'                     "VEHICLE, LOADERCODE, RAMPATTENDANTCODE,WEIGHT"

'response.Write(txtsql)
'response.Flush()
%>
 <table width="592">
  <tr> 
    <td rowspan="5" width="96"><img src="IMAGES/latar.jpg" width="80" height="80"></td>
  </tr>
	<tr><td width="134">Company</td>
	<td  width="346"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="134">Business Unit</td>
	<td  width="346"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="134">Manager</td>
	<td  width="346"><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td width="134">Report</td>
	<td  width="346"><b><FONT color="#000099">DELIVERY PRODUCT QUALITY</FONT></td>
	</tr>
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
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">RETURN DATE</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">DELIVERY NOTE CODE</td>	
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">FFA</td>	
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">MOISTURE</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">DIRT</td>	
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">REMARKS</td>	
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">DIRT AND MOISTURE</td>	
	</tr>
	<%
	Dim v_Quantity
	v_Quantity=0
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("QUALITYDATE")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DELIVERYNOTE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FFA")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MOISTURE")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DIRT")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("REMARKS")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DIRT_AND_MOISTURE")%>&nbsp;</td>
		</tr>
	  <%
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