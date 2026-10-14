<!--#include file="../connDb1.Inc"-->
<!--#include file="JSON.asp"-->
<!--#include file="JSON_UTIL.asp"-->
<%
fcba = Session("fcba")
pocode = Request("pocode")
stono = Request("stono")
result = ""

q = " SELECT DISTINCT dtl.stono" & _
    " FROM stocktransferorderhdr hdr" & _
    " INNER JOIN stocktransferorderdtl dtl ON (dtl.fcba = hdr.fcba AND dtl.stono = hdr.stono)" & _
    " WHERE dtl.fcba = '" & fcba & "' AND dtl.pocode = '" & pocode & "' AND NVL(dtl.do_quantity,0) < NVL(dtl.quantityorder,0)"
if len(trim(stono)) > 0 then q = q & " AND dtl.stono = '"&stono&"'"
'response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("stono") & "'"
		if "" & stono <> "" then
			if stono = rsCallCombo("stono") then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("stono")

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