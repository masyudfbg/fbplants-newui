<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
fcba = Request("fcba")
unitkerja = Request("unitkerja")
result = ""

html = html & "<select id='unitkerja' name='unitkerja' onchange='javascript:getselectedsection(); javascript:getListfieldcode(); javascript:getNoRKB();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

	q = "SELECT '--ALL--' as sectionname, '' as fcname FROM DUAL UNION ALL SELECT DISTINCT sc.fccode, sc.fcname" & _
		"    FROM sectionname sc " & _
		"   order by fcname"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("sectionname") & "'"
		if "" & unitkerja <> "" then
			if unitkerja = cstr(""&rsCallCombo("sectionname")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("sectionname") & " : " & rsCallCombo("fcname")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>