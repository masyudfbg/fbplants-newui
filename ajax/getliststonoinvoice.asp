<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
pocode = Request("pocode")
stono = Request("stono")
result = ""


html = html & "<select id='stono' name='stono' onchange='javascript:getselectedstono();javascript:getstoinfo();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT sth.stono" & _
	"  FROM stocktransferorderdtl std" & _
	"       INNER JOIN stocktransferorderhdr sth" & _
	"          ON (sth.stono = std.stono AND sth.fcba = std.fcba)" & _
	" WHERE sth.fcba = '"& fcba &"' AND sth.pocode = '"& pocode &"'" 

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("stono") & "'"
		if "" & stono <> "" then
			if cstr(""&stono) = cstr(""&rsCallCombo("stono")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("stono")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing
			

html = html &"</select>"		  
response.write(html)

%>