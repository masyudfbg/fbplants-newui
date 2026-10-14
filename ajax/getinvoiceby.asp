<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccompanycode = Request("fccompanycode")
invoice_by = Request("invoice_by")
result = ""

html = html & "<select id='invoice_by' name='invoice_by' class='chosen-select' style='width:70%;' onchange='javascript:getselectedinvoice();'>"
html = html &"<option value=''></option>"

q = " SELECT DISTINCT fcba, fcname" & _
	"  FROM user_fcba_tab uft INNER JOIN businessunit bu ON (bu.fccode = uft.fcba)" & _
	"  INNER JOIN company co ON (co.companycode = bu.fccompanycode)" & _
	" WHERE fccompanycode = '" & fccompanycode & "' AND fcusercode = '"& session("codename") &"' AND bu.Isactive='TRUE'"  
 
' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fcba") & "'"
		if "" & invoice_by <> "" then
			if cstr(""&invoice_by) = cstr(""&rsCallCombo("fcba")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fcba") & " : " & rsCallCombo("fcname") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>