<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccompanycode = Request("fccompanycode")
fcba = Request("fcba")
customercode = Request("customercode")
result = ""

html = html & "<select id='customercode' name='customercode' onchange='javascript:getselectedcustomer();javascript:getlistsalescontract();' style='width: 55%;' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT distinct fccode, fcname" & _
	" FROM supplier" & _
	" WHERE fcba ='"& fcba &"' and controljob='2110103'" & _
	" GROUP BY fccode, fcname" & _
	" ORDER BY fccode"  

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

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>