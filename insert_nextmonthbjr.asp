<%
Dim globalSQL, globalRS, Check_Validation

Response.Write "<br/> Processing Next Month ABW"

Check_Validation = "FALSE"

globalSQL = " select nvl(parameter_api.getparam1 ('HARVESTING_ABW', 'AUTOMATIC_PROCESS', '" & Session("FCBA") & "'), 'FALSE') details" & _
                   " from dual"

Set globalRS = Server.CreateObject("ADODB.Recordset")
globalRS.CursorType = 3
globalRS.CursorLocation = 3
globalRS.Open globalSQL, connstr

if not globalRS.eof then
    Check_Validation = globalRS("details")
else
    Check_Validation = "FALSE"
end if
globalRS.close
set globalRS = Nothing

Dim Param_ClosedMonth, Param_ClosedYear

Param_ClosedYear = Cdbl( Left(cmdAccountPeriodClose("Costbook"), 4) )
Param_ClosedMonth = Cdbl( Right(cmdAccountPeriodClose("Costbook"), 2) )


if Check_Validation = "TRUE" then
	globalSQL	= 	" BEGIN automatic_process_api.nextmonth_harvesting_abw('" & Session("FCBA") & "', " & Param_ClosedYear & ", " & Param_ClosedMonth & "); END; "	
	'Response.Write globalSQL & "<br/>"
	connStr.execute(globalSQL)
	Response.Write "<br/> Processing Next Month ABW, <font color=""green"">Success!!!</font>"
else
	Response.Write "<br/> &nbsp;&nbsp;&nbsp;&nbsp; <b><i>Your Business Unit Not Yet Configured For Automatic Process</i></b>"
	Response.Write "<br/> Processing Next Month ABW, <font color=""green"">Success!!!</font>"
end if
%>