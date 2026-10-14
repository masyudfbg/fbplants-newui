<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
invoicedpno = Request("invoicedpno")
fcba = Request("fcba")
result = ""

html = html & "<select id='invoicedpno' name='invoicedpno' onchange='javascript:getselectedinvoicedpno(); javascript:getlistlinenosales(); javascript:Getdetinvoicedp();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT invoicedpno FROM salesinvoicedp WHERE fcba = '"& fcba &"'"   

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("invoicedpno") & "'"
		if "" & invoicedpno <> "" then
			if cstr(""&invoicedpno) = cstr(""&rsCallCombo("invoicedpno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("invoicedpno")  

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>