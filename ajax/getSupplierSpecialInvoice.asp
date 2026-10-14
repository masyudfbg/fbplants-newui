<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
fcba = Request("fcba")
suppliercode = Request("suppliercode")
invoice_type = Request("invoice_type")
result = ""

html = html & "<select id='suppliercode' name='suppliercode' onchange='javascript:getselectedsupplier();' style='width: 60%' class='chosen-select'>"
html = html &"<option value=''></option>"

If invoice_type = "TBS" Then
	q = " select customercode as fccode, description as fcname from customer where fcba ='"& fcba &"' and supplier= 'Y'"
Else
	q = " select fccode, fcname from supplier where fcba ='"& fcba &"'"
End If

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
	  rsCallCombo.Open q, connstr
	  if not rsCallCombo.eof then
	  	response.write "<option value=''></option>"
			Do While Not rsCallCombo.eof
			 html = html & "<option value='" & rsCallCombo("fccode") & "'" 
			 if Request("fccode") = rsCallCombo("fccode") then html = html &  " selected"
			 html = html &  " >"
			 html = html & rsCallCombo("fccode") &" : "& rsCallCombo("fcname")

			html = html &"</option>" 
			rsCallCombo.movenext
		loop

	End if
	rsCallCombo.Close
	set rsCallCombo = Nothing 
					  
	response.write(html)
	html = html &"</select>"
' set rsrecurtrans = server.CreateObject("Adodb.recordset")
' rsrecurtrans.Open q, connstr

' if Not rsrecurtrans.eof then
' 	VAT = rsrecurtrans("VAT")
' 	salescontractno = rsrecurtrans("salescontractno")
' end if
' rsrecurtrans.Close
' set rsrecurtrans = Nothing

' response.write "{"
' response.write """salescontractno"" : """ & salescontractno & """"
' response.write "}"

%>