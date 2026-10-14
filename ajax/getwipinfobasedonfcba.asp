<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccompanycode = Request("fccompanycode")
fcba = Request("fcba")
transdate = Request("transdate")
agreementcode = Request("agreementcode")
result = ""

html = html & "<select id='agreementcode' name='agreementcode' class='chosen-select' style='width:60%;' onchange='javascript:getretention(); javascript:getselectedagreement();javascript:getDP();'>"
html = html &"<option value=''></option>"

q = "  SELECT agreementcode" & _
	"    FROM workinprogressfinal" & _
	"   WHERE fcba = '"& fcba &"'" & _
	" GROUP BY agreementcode, fcba"


' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("agreementcode") & "'"
		if "" & agreementcode <> "" then
			if cstr(""&agreementcode) = cstr(""&rsCallCombo("agreementcode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("agreementcode")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>