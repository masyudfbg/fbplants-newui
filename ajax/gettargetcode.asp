<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fieldcode = Request("fieldcode")
fcba = Request("fcba")
activity = Request("activity")

'response.write activity

result = ""

	html = html & "<select id='fieldcode' name='fieldcode' onchange='javascript:getselectedfield(); javascript:getFieldInfo(); javascript:getNoRKB();' class='chosen-select' style='width:40%;'>"
	html = html &"<option value=''></option>"

		q = " SELECT field.fccode, field.fcname,field.HECTARAGEPLANTED FROM field "&_
			" LEFT join jobmaster on jobmaster.fccode ='"& activity &"' and field.DX_BLOCK_STATUS = jobmaster.job_forfieldstatus"&_
			" WHERE field.fcba = '"& fcba &"' AND field.activation = 'Y' " & _
			" UNION ALL " & _
			" SELECT nurserycode fccode,description fcnam,1 HECTARAGEPLANTED " & _
			" FROM nursery " & _
			" LEFT join jobmaster on jobmaster.fccode ='"& activity &"' and nursery.stage = jobmaster.job_category " & _
			" WHERE nursery.fcba='"& fcba &"' AND nursery.status='Y' "

	
	' response.Write q

	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	rsCallCombo.Open q, connstr
	if not rsCallCombo.eof then
		response.write "<option value=''></option>"
		Do While Not rsCallCombo.eof
			html = html & "<option value='" & rsCallCombo("fccode") & "'"
			if "" & fieldcode <> "" then
				if cstr(""&fieldcode) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
			end if
			html = html &  " >"
			html = html & rsCallCombo("fccode") & " - Blok:" & rsCallCombo("fcname") & " Ha:" & rsCallCombo("HECTARAGEPLANTED") 

			html = html &"</option>" 
			rsCallCombo.movenext
		loop

	End if
	rsCallCombo.Close
	set rsCallCombo = Nothing

	html = html &"</select><font color='#FF0000'>&nbsp;*</font>"
	response.write(html) 
%>