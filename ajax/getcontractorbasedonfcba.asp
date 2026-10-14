<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
contractorcode = Request("contractorcode")
result = ""

q = " SELECT fccode, fcname" & _
	" FROM contractor" & _
	" WHERE fcba = '" & fcba & "'" & _
	" ORDER BY fccode"
' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & contractorcode <> "" then
			if cstr(""&contractorcode) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fccode") & ": " & rsCallCombo("fcname") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

else

html = "<option value=''></option>"
End if
rsCallCombo.Close
set rsCallCombo = Nothing  
					  
response.write(html)

%>