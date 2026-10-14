<%Option Explicit%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>EPlas - Calculate Payroll 2</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
</head>
<body>
<table>
<%	
Dim q,nCountofField,no,i,txttable,txtmodule
q="Select * from accountcode where fccode like '" &  "1220%" & "'"
Set rsCat = connStr.Execute(q)
if Not rsCat.eof then
   nCountofField=rsCat.fields.count-1
end if
NO=1
for i = 0 to nCountofField%>
	<td align="center"><%=rsCat(i).Name%></td>
<%Next%>
</tr>
<%While Not rsCat.EOF%>
	<tr>
	<td align="left"><%=NO%></td>
	<%for i = 0 to nCountofField
		if len(trim(rscat(i).value))>0 then%>
		   <td align="left"><%=rsCat(i)%></td>
		<%else%>
		   <td align="left">-</td>
		<%end if%>
	<%next%>
	</tr>
  <%
  rsCat.MoveNext
  no=no+1
Wend
rsCat.close()
rsCat=null%>

<%
'Proses Closing
	Dim cn,  cmd, param
	Dim SQL
	SQL = "{call RunCalculate(?, {resultset 0, io_cursor})}"
	set cmd = server.CreateObject ("ADODB.Command")
	with cmd
		set .ActiveConnection	= connStr
	    .CommandText			= SQL
	    .CommandType			= 1		'adCmdText
	    'Pass in 0 to retrieve all of the records
	    set param = .CreateParameter("id", 131 , 1 , 1, "01-2008")  '<<131=adNumeric, 1=adParamInput
		
	    .Parameters.Append param
	end with
					
	set rs = server.CreateObject ( "ADODB.Recordset" ) 
	set rs = cmd.execute
%>	
</table>
<form action="Calculate.Asp" method="post">
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	<input type="button" value="Calculate Payroll">
</form>
</body>
</html>