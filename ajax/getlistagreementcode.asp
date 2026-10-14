<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%
txtcari = Request("keyword")
fcba = Request("fcBA")
txtsupplier = Request("contractor")

html = html & "<select id='agreementcode' name='agreementcode' class='chosen-select' style='width: 60%;' onchange='getselectedkontrak();'>"
html = html &"<option value=''></option>"

sqlview =  "select h.agreementcode as id,h.agreementcode as text " & _
" from contractagreement h " & _
" where  h.fcba='" & fcba & "' " & _
" and h.agreementcode like UPPER('%" & txtcari & "%') " & _ 
" and h.contractorcode = '" & txtsupplier & "'" & _
" group by h.agreementcode order by h.agreementcode asc"


Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open sqlview, connstr
if not rsCallCombo.eof then
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("id") & "'"
		if "" & fcba <> "" then
			if cstr(""&txtcari) = cstr(""&rsCallCombo("id")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("id") &" : "& rsCallCombo("text")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)
%>