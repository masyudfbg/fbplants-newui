<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
xcode = Request("xcode")
selectedsupplier = Request("selectedsupplier")
result = ""

html = html & "<select id='xcode' name='xcode' onchange='javascript:getselectedpo();javascript:getliststo()' style='width: 70%' class='chosen-select'>"
html = html &"<option value=''></option>"

q =	" Select distinct(a.pocode) pocode" & _
	" from PurchaseOrderDetails a" & _ 
	"      , purchaseorder b " & _
	" where a.pocode = b.pocode" & _
	"       and a.fcba = b.fcba" & _
	"       and nvl(b.rowstateid, 'Ordered') in ('Ordered','Confirmed')" & _
	"       and b.fcba = '" & session("fcba") & "'" & _
	"       and b.suppliercode = '" & selectedsupplier & "'" & _ 
	"       and a.received < a.quantity order by a.pocode Desc"
									
 'response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("pocode") & "'"
		if "" & xcode <> "" then
			if cstr(""&xcode) = cstr(""&rsCallCombo("pocode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("pocode")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>