<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
result = ""


html = html & "<select id='suppliercode' name='suppliercode' onchange='javascript:getselectedsupplier();javascript:getPOgrn();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " select * from supplier where fcba ='"& session("fcba") &"' order by fccode"

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