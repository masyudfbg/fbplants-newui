<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
stono = Request("stono")
fcba = Request("fcba")
stolineno =Request("stolineno")
result = ""

html = html & "<select id='stolineno' name='stolineno' onchange='javascript:getPODO();' style='width: 30%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " select * from stocktransferorderdtl std " & _
	" inner join stocktransferorderhdr sth on (sth.pocode=std.pocode and sth.fcba=std.fcba and sth.stono = std.stono) " & _
	" where sth.fcba = '"& fcba &"' " & _
	" and sth.stono= '"& stono &"' "
' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("stolineno") & "'"
		if "" & stolineno <> "" then
			if cstr(""&stolineno) = cstr(""&rsCallCombo("stolineno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("stolineno")  

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>