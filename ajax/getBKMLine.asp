<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
bkmcode = Request("bkmcode")
fcba = Request("fcba")
result = ""

html = html & "<select id='prlineid' name='prlineid' onchange='javascript:getStockCode();javascript:getselectedprline();' style='width: 40%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT distinct jobcode||locationtype||locationcode lineid, job_api.get_jobcode_name(jobcode,fcba)||'-'||locationtype||'-'||locationcode linedet" & _
	" FROM attendance_gad_temp " & _
	" WHERE bkmcode = '"& bkmcode &"' AND fcba = '"& fcba &"' ORDER BY jobcode,locationcode" 
	

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr

if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("lineid") & "'"
		if "" & prlineid <> "" then
			if cstr(""&prlineid) = cstr(""&rsCallCombo("lineid")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("linedet") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>