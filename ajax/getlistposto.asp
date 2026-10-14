<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
fccode = Request("fccode")
result = ""

html = html & "<select id='fccode' name='fccode' onchange='javascript:getselectedpo();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " select * from purchaseorder where fcba ='"& fcba &"' and rowstateid not in ('Rejected','Planned') AND incoterm1 = 'LCO' order by pocode"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("pocode") & "'"
		if "" & fccode <> "" then
			if cstr(""&fccode) = cstr(""&rsCallCombo("pocode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("pocode") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>