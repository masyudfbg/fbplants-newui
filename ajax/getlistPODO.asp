<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
fcba = Request("fcba")
result = ""

html = html & "<select id='fccode' name='fccode' onchange='javascript:getselectedpo();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " select a.* from purchaseorder a left JOIN receivenote b ON (a.pocode=b.pocode AND a.fcba=b.fcba) where b.pocode IS NULL AND a.fcba ='"& fcba &"' AND a.rowstateid='Ordered' order by a.pocode"
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