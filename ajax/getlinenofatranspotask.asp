<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
lineno = Request("lineno")
fixassetcode = Request("fixassetcode")
fcba = Request("fcba")
result = ""

html = html & "<select id='lineno' name='lineno' onchange='javascript:getselectedline(); javascript:getAssetWriteOffInfo();' class='chosen-select' style='width:50%;'>"
html = html &"<option value=''></option>"

q = "  SELECT MAX(fatd.lineno) lineno" & _
	"    FROM fixedasset fat" & _ 
	"         INNER JOIN fixedasset_depreciation_ledger fatd" & _ 
	"            ON (fat.fccode = fatd.fccode AND fat.fcba = fatd.fcba)" & _ 
	"   WHERE fat.fccode = '"& fixassetcode &"' AND fatd.yearlydepreciation > 0 " & _ 
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