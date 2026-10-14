<%

Function IsDailyClosing(initfcba, initdate)
    Dim sqlisdailyclosing, rsisdailyclosing, ResultValue
	
	sqlisdailyclosing = " select nvl(dailyclosingactivity_api.isdailyclosing ('" & initfcba & "'" & _
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

Function IsAccClosing(initfcba, initdate)
    Dim sqlisaccclosing, rsisaccclosing, ResultValue
	
	sqlisaccclosing = " select nvl(EMPALLDEDSTATUS_API.GetStatusClosing_acc  (to_number(to_char(to_date('" & initdate & "','mm/dd/yyyy'),'YYYY'))," & _ 
                                                                              " to_number(to_char(to_date('" & initdate & "','mm/dd/yyyy'),'MM'))," & _ 
	                                                                          "'" & initfcba & "'), 'FALSE') resultvalue" & _
						" from dual"

	set rsisaccclosing = Server.CreateObject("ADODB.Recordset")
	rsisaccclosing.Open sqlisaccclosing, connstr
	if not rsisaccclosing.eof then
		ResultValue = cstr(rsisaccclosing("resultvalue"))
	else
		ResultValue = "FALSE"
	end if
	sqlisaccclosing = ""
	rsisaccclosing.Close
	set rsisaccclosing = Nothing
	
	IsAccClosing = ResultValue
	
End Function

Function IsAccClosingPrd(initfcba, initmonth, inityear)
    Dim sqlisaccclosing, rsisaccclosing, ResultValue
	
	sqlisaccclosing = " select nvl(EMPALLDEDSTATUS_API.GetStatusClosing_acc  (to_number('" & initmonth & "')," & _ 
                                                                            " to_number('" & inityear & "')," & _ 
	                                                                       "'" & initfcba & "'), 'FALSE') resultvalue" & _
						" from dual"

	set rsisaccclosing = Server.CreateObject("ADODB.Recordset")
	rsisaccclosing.Open sqlisaccclosing, connstr
	if not rsisaccclosing.eof then
		ResultValue = cstr(rsisaccclosing("resultvalue"))
	else
		ResultValue = "FALSE"
	end if
	sqlisaccclosing = ""
	rsisaccclosing.Close
	set rsisaccclosing = Nothing
	
	IsAccClosingPrd = ResultValue
	
End Function

Function IsCanGenerate(initfcba, initdate)
    Dim sqlCanGenerate, rsCanGenerate, ResultValue
	
	sqlCanGenerate = " select count(*) jml from DAILYCOUNTINGGENERATE where fcba='" & initfcba & "'" & _
	                 " and to_char(TDATE,'YYYYMMDD') = to_char(to_date('" & initdate & "', 'mm/dd/yyyy'),'YYYYMMDD') " 


	ResultValue = "FALSE"
	set rsCanGenerate = Server.CreateObject("ADODB.Recordset")
	rsCanGenerate.Open sqlCanGenerate, connstr
	if not rsCanGenerate.eof then
	   if Cint(rsCanGenerate("jml")) <= 2 then
		ResultValue = "TRUE"
	   end if   
	end if
	sqlCanGenerate = ""
	rsCanGenerate.Close
	set rsCanGenerate = Nothing
	
	IsCanGenerate = ResultValue
	
End Function

%>