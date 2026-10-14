<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
deliverytype = Request("deliverytype")
result = ""


html = html & "<select id='deliverytype' name='deliverytype' onchange='javascript:getselecteddeliverytype();' style='width: 75%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT fccode, fcname" & _
    " FROM parameter" & _
    " WHERE fcba = '" & fcba & "'" & _
	"      AND parcode = 'DELIVERY_TYPE'" & _
	" ORDER BY fccode" 
' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & deliverytype <> "" then
			if cstr(""&deliverytype) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
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