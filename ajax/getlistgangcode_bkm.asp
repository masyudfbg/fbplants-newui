<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
fcba = Request("fcba")
sectionname = Request("sectionname")
gangcode = Request("gangcode")
result = ""

html = html & "<select id='gangcode' name='gangcode' onchange='javascript:getselectedgangcode();' style='width: 30%' class='chosen-select'>"
html = html &"<option value=''></option>"

	q = " SELECT DISTINCT FCCODE as gangcode, fcname FROM GANG WHERE ACTIVATION = 'Y' AND fcba = '"& session("fcba") &"' AND remarks= '"& sectionname &"' ORDER BY gangcode ASC"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("gangcode") & "'"
		if "" & gangcode <> "" then
			if cstr(""&gangcode) = cstr(""&rsCallCombo("gangcode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("gangcode") & " : " & rsCallCombo("fcname") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>