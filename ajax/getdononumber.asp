<!--#include file="../connDb1.Inc"-->
<!--#include file="JSON.asp"-->
<!--#include file="JSON_UTIL.asp"-->
<%
fcba = Session("fcba")
pocode = Request("pocode")
stono = Request("stono")
dono = Request("dono")
q = " SELECT DISTINCT dtl.do_no dono FROM deliveryorderpurchasedtl dtl" & _
    " INNER JOIN deliveryorderpurchasehdr hdr ON (hdr.fcba = dtl.fcba AND hdr.do_no = dtl.do_no)" & _
    " WHERE dtl.fcba = '" & fcba & "'" & _
    "       AND hdr.pocode = '" & pocode & "'" & _
    "       AND dtl.stono = '" & stono & "'"
if len(trim(dono)) > 0 then q = q & " AND dtl.do_no = '"&dono&"'"
'response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("dono") & "'"
		if "" & dono <> "" then
			if dono = rsCallCombo("dono") then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("dono")

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