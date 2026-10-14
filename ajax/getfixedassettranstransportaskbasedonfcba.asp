<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccompanycode = Request("fccompanycode")
fcba = Request("fcba")
fa_transportask = Request("fa_transportask")
result = ""

html = html & "<select id='fa_transportask' name='fa_transportask' onchange='javascript:getselectedtransportask(); javascript:getdatetransportask();' class='chosen-select' style='width: 50%;'>"
html = html &"<option value=''></option>"

q =  	" SELECT fatd.fccode" & _
		"  FROM fixedasset_transportask_det fatd" & _
		"       INNER JOIN intraco_tab it ON (it.fccode = fatd.jobcode_intraco)" & _
		" WHERE it.fcba = '"& session("fcba") &"' AND fatd.fcba = '"& fcba &"'" & _
		" MINUS" & _
		" SELECT fccode" & _
		"  FROM fixedasset_receivedtask" & _
		" WHERE fcba = '"& session("fcba") &"' AND fcba_source = '"& fcba &"'"

	' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & fa_transportask <> "" then
			if cstr(""&fa_transportask) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fccode")
		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>