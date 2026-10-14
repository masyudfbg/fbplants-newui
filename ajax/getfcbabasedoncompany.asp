<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccompanycode = Request("fccompanycode")
fcba = Request("fcba")
result = ""

q = " SELECT fcba FROM user_fcba_tab uft INNER JOIN businessunit bu ON bu.fccode = uft.fcba WHERE fccompanycode = '" & fccompanycode & "' AND uft.fcusercode = '" & Session("CodeName") & "'" 
' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fcba") & "'"
		if "" & fcba <> "" then
			if fcba = rsCallCombo("fcba") then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fcba")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

else

html = "<option value=''></option>"
End if
rsCallCombo.Close
set rsCallCombo = Nothing  
					  
response.write(html)

%>