<%Option Explicit%>
<%response.buffer=True%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>iPS - Help</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
</head>
<body Onload=TeksJalan>
<!--#include file="top.jsp"-->

<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td> 
 
    <table class="data" >
     <%
		Dim q
		q="SELECT  *  FROM HISTORICAL_USER_TAB"
	    Set rs = connStr.Execute(q)
		%>			
				<tr>
					<tr bgcolor="#6666FF" >
					<Td align="left" valign="top">USER NAME</Td></Td>
					<td align="left" valign="top">IP ADDRESS</td>
					<td align="left" valign="top">BUSSINES UNIT</td>
					<Td align="left" valign="top">DATE ACTION</Td>
					<Td align="left" valign="top">TIME ACTION</Td>
				</tr>
			 <%do until rs.EOF%>		 
					<tr>	
						<td align="left" valign="top" ><%Response.Write(rs.fields("USERNAME"))%></td>
						<td align="left" valign="top" ><%Response.Write(rs.fields("IP_ADDRESS"))%></td>
						<td align="left" valign="center" ><%Response.Write(rs.fields("BU"))%></td>
						<td align="left" valign="top" ><%Response.Write(rs.fields("DATE_ACT"))%></td>
						<td align="left" valign="top" ><%Response.Write(rs.fields("TIME_ACT"))%></td
					</tr>	
					<%  rs.MoveNext %>
				 
				<%loop
				rs.close
				%>
		</table>	
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
</html>