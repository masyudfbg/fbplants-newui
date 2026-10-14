<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
xcode = Request("xcode")
stocktransno2 = Request("stocktransno2")
result = ""

html = html & "<select id='stocktransno2' name='stocktransno2' onchange='javascript:getselectedstogrn(); javascript:getlistdo();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = "  SELECT DISTINCT stono, pocode" & _
	"    FROM (SELECT a.stono, a.pocode" & _
	"            FROM stocktransferorderdtl a" & _
	"                 INNER JOIN stocktransferorderhdr b" & _
	"                    ON (a.stono = b.stono AND a.fcba = b.fcba)" & _
	"           WHERE     (a.quantityorder - NVL(do_quantity,0)) >= 0" & _
	"                 AND a.pocode = '"& xcode &"'" & _
	"                 AND a.fcba = '"& session("fcba") &"')" & _
	" ORDER BY stono" 

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("stono") & "'"
		if "" & stocktransno2 <> "" then
			if cstr(""&stocktransno2) = cstr(""&rsCallCombo("stono")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("stono") &" : "& rsCallCombo("pocode") 

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>