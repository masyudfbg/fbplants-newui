<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
customercode = Request("customercode")
result = ""

q = " SELECT customercode fccode, description fcname" & _
    " FROM customer" & _
    " WHERE fcba = '" & fcba & "'" & _
	"      AND supplier = 'N'" & _
	" ORDER BY customercode" 
' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & customercode <> "" then
			if cstr(""&customercode) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fccode") & " : " & rsCallCombo("fcname")

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