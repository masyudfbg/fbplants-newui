<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Activity Norm Harvesting</title>
</head>
<body>
<%
plantingyear = request("plantingyear")
	if plantingyear <> "" then
	txtSQL = " SELECT PLANTINGYEAR, ROTASI, BASIS_JJG, RATE_SB, LB_RATE, NORMA_HK, RATE1" & _
             " FROM ACTIVITY_NORM_HARV" & _
             " WHERE fcba = '" & session("fcBA") & "'" & _
             "       AND PLANTINGYEAR = '" & plantingyear & "'"
	else
	txtSQL = " SELECT PLANTINGYEAR, ROTASI, BASIS_JJG, RATE_SB, LB_RATE, NORMA_HK, RATE1" & _
             " FROM ACTIVITY_NORM_HARV" & _
             " WHERE fcba = '" & session("fcBA") & "'"
	end if

'response.Write txtsql
Set rsCat = connStr.Execute(txtSQL)
%>
	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>Activity Norm Harvesting</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="7">&nbsp;</td>
						</tr>
						</div>
	</tr>
	<tr>
		<th>Planting Year</th>
		<th>Rotasi</th>
		<th>Basis Jjg</th>
		<th>Rate SB</th>
		<th>LB Rate</th>
		<th>Norm</th>
		<th>Rate</th>		
	</tr>
	<%While Not rsCat.EOF%>
	 	<tr>
		<td align="left"><font face="Verdana" size="1"><%=rscat("PLANTINGYEAR")%></td>		
		<td align="left"><font face="Verdana" size="1"><%=rscat("ROTASI")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("BASIS_JJG")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("RATE_SB")%></td>		 
		<td align="left"><font face="Verdana" size="1"><%=rscat("LB_RATE")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("NORMA_HK")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("RATE1")%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>