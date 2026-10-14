<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
lineno = Request("lineno")
fixassetcode = Request("fixassetcode")
result = ""

html = html & "<select id='lineno' name='lineno' onchange='javascript:getselectedline(); javascript:getAssetWriteOffInfo();' class='chosen-select' style='width:30%;'>"
html = html &"<option value=''></option>"

q = "  SELECT fal.lineno," & _ 
	"         fat.total," & _ 
	"         fal.yearlydepreciation," & _ 
	"         fat.fcname" & _ 
	"    FROM fixedasset fat" & _ 
	"         INNER JOIN Fixedasset_Depreciation_Ledger fal" & _ 
	"            ON (fat.fccode = fal.fccode AND fat.fcba = fal.fcba)" & _ 
	"   WHERE fat.fccode = '"& fixassetcode &"' AND fat.fcba = '"& session("fcba") &"'" & _ 
	"   ORDER BY lineno"  
  

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("lineno") & "'"
		if "" & lineno <> "" then
			if cstr(""&lineno) = cstr(""&rsCallCombo("lineno")) then html = html &  " selected"
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