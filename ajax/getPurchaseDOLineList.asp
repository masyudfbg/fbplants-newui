<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
prcode = Request("prcode")
fcba = Request("fcba")
doline = Request("doline")
deliveryorderno = Request("deliveryorderno")
fccompanycode = Request("fccompanycode")
result = ""

html = html & "<select id='doline' name='doline' onchange='javascript:getselecteddoline();javascript:getMessage();' style='width: 40%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT lineno" & _
	"  FROM purchasedeliveryorder" & _
	" WHERE deliveryorderno = '"& deliveryorderno &"' AND fcba = '"& fcba &"' ORDER BY lineno" 
	

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr

if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("lineno") & "'"
		if "" & doline <> "" then
			if cstr(""&doline) = cstr(""&rsCallCombo("lineno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("lineno") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>