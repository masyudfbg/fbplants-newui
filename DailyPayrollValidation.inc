<% 

Function IsDailyClosing(initfcba, initdate)
    Dim sqlisdailyclosing, rsisdailyclosing, ResultValue
	
	sqlisdailyclosing = " select nvl(dailyclosingpayroll_api.isdailyclosing ('" & initfcba & "'" & _
	                    "                                                     , to_date('" & initdate & "', 'mm/dd/yyyy')), 'FALSE') resultvalue" & _
						" from dual"

	set rsisdailyclosing = Server.CreateObject("ADODB.Recordset")
	rsisdailyclosing.Open sqlisdailyclosing, connstr
	if not rsisdailyclosing.eof then
		ResultValue = cstr(rsisdailyclosing("resultvalue"))
	else
		ResultValue = "FALSE"
	end if
	sqlisdailyclosing = ""
	rsisdailyclosing.Close
	set rsisdailyclosing = Nothing
	
	IsDailyClosing = ResultValue
	
End Function

 %>
