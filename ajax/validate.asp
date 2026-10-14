<!--#include file="../connDb1.Inc"-->
<%
validate = Request("validate")
dim x
if validate = "svr1" then
x = "x"
elseif validate = "svr2" then
x = "x2"
elseif validate = "svr3" then
x = "x3"
elseif validate = "svr4" then
x = "x4"
elseif validate = "empcd" then
x = "x5"
end if

dim txtstr

response.expires=-1
	
	 str = " select * from employee where FCBA = '"& Session("fcBA") &"' and fccode = '" & Request("q") & "'"

					  Set rs = Server.CreateObject("ADODB.Recordset")

					  rs.Open str, connstr

						   

					  IF Not rs.eof then
						 txtstr = rs("FCNAME")
					  End if

					  rs.Close

					  set rs = Nothing  
					  
					  
	response.expires=-1
					  
	'get the q parameter from URL
	'm=ucase(request.querystring("q"))

	if len(trim(txtstr)) > 0 then
	 response.write(txtstr)
	 else
	 response.write(x)
	 end if

%>