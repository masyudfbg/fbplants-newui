<%Option Explicit%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<P>
<%
		
	Dim cn, cmd, param,i
	Dim SQL

	SQL = "{call curspkg_join.open_join_cursor1(?, {resultset 0, io_cursor})}"
	set cmd = server.CreateObject ("ADODB.Command")
	with cmd
		set .ActiveConnection	= connStr
	    .CommandText			= SQL
	    .CommandType			= 1		'adCmdText
	    'Pass in 0 to retrieve all of the records
		' Yang Paling Belakang adalah Parameter di Procedure 0 = semua data ; 1 empcode = 1
	    set param = .CreateParameter("id", 131 , 1 , , 0)  '<<131=adNumeric, 1=adParamInput 
	    .Parameters.Append param
	end with
					
	set rs = server.CreateObject ( "ADODB.Recordset" ) 
	set rs = cmd.execute
		    
	%>
	<TABLE WIDTH="80%" ALIGN=center BORDER=1 CELLSPACING=3 CELLPADDING=3>
	<%
	Response.Write ("<TR bgcolor=Gray>")
	For i = 0 To rs.Fields.Count - 1
		Response.Write ("<TD>" & rs.Fields(i).Name & "</TD>")
	Next
	Response.Write ("</TR>")
	Do until rs.EOF
		response.write ("<TR>")
		for i = 0 to rs.Fields.Count - 1
			If IsNull(rs(i).value) then 
				Response.Write ("<TD> NULL </TD>")
			else 
				Response.Write ("<TD>" & rs(i).value & "</TD>")
			end if 
		next 
		rs.MoveNext
		response.write ("</TR>")
	loop
	%>
	</TABLE>
	<%
	Set cmd = Nothing
	Set param = Nothing
	rs.close
	Set rs = Nothing
%>
</P>
</BODY>
</HTML>