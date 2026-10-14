<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
fcba = session("fcba")
result = ""


html = html & "<select id='jobcode' name='xjobCode' onchange='javascript:getselectedjobcode();javascript:getTragettype();javascript:getJobcodeIntraco();javascript:getViewcharge();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " select * from job where ACTIVATION = 'Y' AND fcba ='"& fcba &"' Order by fccode"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & fccode <> "" then
			if cstr(""&fccode) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
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