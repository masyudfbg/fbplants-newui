<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
fcba = Request("fcba")
activitycode = Request("activitycode")
result = ""

html = html & "<select id='activitycode' name='activitycode' onchange='javascript:getNoRKB(); javascript:getNormaHK(); javascript:getselectedactivity();' style='width: 30%' class='chosen-select'>"
html = html &"<option value=''></option>"

	q = "select fccode, fcname from jobmaster where jobgroupcode in (select fccode from jobgroupmaster where activity_heading = 'PNN')"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & activitycode <> "" then
			if activitycode = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fccode") & " : " & rsCallCombo("fcname")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>