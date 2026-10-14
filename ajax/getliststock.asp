<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
fcba = Request("fcba")
stock = Request("stock")
result = ""

html = html & "<select id='stock' name='stock' onchange='javascript:getselectedstock();' style='width: 80%' class='chosen-select'>"
html = html &"<option value=''></option>"

	q = "select  * from stock where fcba ='"& fcba &"' and stockgroupcode IN ('301','302','305')"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & stock <> "" then
			if cstr(""&stock) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fccode") & " : " & rsCallCombo("fcname")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select><font color='#FF0000'>&nbsp;*</font>"
response.write(html)

%>