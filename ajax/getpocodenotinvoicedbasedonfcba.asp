<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fcba = Request("fcba")
pocode = Request("pocode")
result = ""


html = html & "<select id='pocode' name='pocode' onchange='javascript:getselectedpocode();javascript:getpoinfo();' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q = " SELECT DISTINCT hdr.pocode fccode,sup.fcname " & _
	" FROM receivenotedetail dtl" & _
	" INNER JOIN receivenote hdr ON (hdr.receivenotecode = dtl.receivenotecode AND hdr.fcba = dtl.fcba)" & _
	" LEFT JOIN SUPPLIER sup ON hdr.suppliercode=sup.fccode and hdr.fcba=sup.fcba " & _ 
	" WHERE dtl.fcba = '" & fcba & "'" & _
	"       AND dtl.amountinvoice < dtl.amountpo_receive" & _
	" ORDER BY hdr.pocode"
' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("fccode") & "'"
		if "" & pocode <> "" then
			if cstr(""&pocode) = cstr(""&rsCallCombo("fccode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("fccode") & ":" & rsCallCombo("fcname")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing
			

html = html &"</select>"		  
response.write(html)

%>