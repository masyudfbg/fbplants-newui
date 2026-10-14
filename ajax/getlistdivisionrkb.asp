<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
fcba = Request("fcba")
divisi = Request("afdeling")
result = ""


html = html & "<select id='divisi' name='divisi' onchange='javascript:getNoRKB(); javascript:getselectedafdeling(); javascript:gettargetcode();' style='width: 50%' class='chosen-select'>"
html = html &"<option value=''></option>"

	q = " SELECT '--ALL--' as sectionname, '' as fcname FROM DUAL UNION ALL SELECT DISTINCT sc.fccode, sc.fcname" & _
		"    FROM sectionname sc " & _
		"   order by fcname" 

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

html = html &"</select><font color='#FF0000'>&nbsp;*</font>"
response.write(html)

%>