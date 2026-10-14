<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
fcba = Request("fcba")
divisi = Request("divisi")
result = ""

html = html & "<select id='divisi' name='divisi' onchange='javascript:getselecteddivisi(); javascript:getlistgangcode();' style='width: 100%' class='chosen-select'>"
html = html &"<option value=''></option>"

	q = " SELECT '--ALL--' as sectionname, '' as fcname FROM DUAL UNION ALL SELECT DISTINCT emp.sectionname, sc.fcname" & _
		"    FROM employee emp" & _
		"         INNER JOIN sectionname sc ON (emp.sectionname = sc.fccode)" & _
		"   WHERE fcba = '"& fcba &"' order by sectionname" 

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("sectionname") & "'"
		if "" & divisi <> "" then
			if cstr(""&divisi) = cstr(""&rsCallCombo("sectionname")) then html = html &  " selected"
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