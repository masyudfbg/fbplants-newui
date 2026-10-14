<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
fcba = Request("fcba")
suppliercode = Request("suppliercode")
result = ""

html = html & "<select id='suppliercode' name='suppliercode' onchange='javascript:getselectedsupplier();' style='width: 60%' class='chosen-select'>"
html = html &"<option value=''></option>"

' q = " select sp.fccode as suppliercode, sp.fcname as suppliername, bu.fccode , bu.fcname, bu.fccompanycode " & _
' 	" from supplier sp Inner join businessunit bu ON (sp.fcba = bu.fccode)" & _
' 	" where bu.fccode = '"& fcba &"' and sp.fccode='"& fccode &"'"

q = " select * from supplier where fcba ='"& fcba &"'"

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