<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
invoicedpno = Request("invoicedpno")
fcba = Request("fcba")
linenoinvoicedp = Request("linenoinvoicedp")
result = ""

html = html & "<select id='linenoinvoicedp' name='linenoinvoicedp' onchange='javascript:getselectedlinenosales(); javascript:Getdetinvoicedp();' style='width: 20%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT lineno, dp_aftervat  FROM purchaseinvoicedp WHERE fcba = '"& fcba &"' AND invoicedpno = '"& invoicedpno &"'"   

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("lineno") & "'"
		if "" & linenoinvoicedp <> "" then
			if cstr(""&linenoinvoicedp) = cstr(""&rsCallCombo("lineno")) then html = html &  " selected"
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