<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
businessunit = Request("businessunit")
result = ""

' q = " select sp.fccode as suppliercode, sp.fcname as suppliername, bu.fccode , bu.fcname, bu.fccompanycode " & _
' 	" from supplier sp Inner join businessunit bu ON (sp.fcba = bu.fccode)" & _
' 	" where bu.fccode = '"& fccode &"' and sp.fccode='"& suppliercode &"'"

q = " select * from purchaseorder po " & _
	" inner join purchaseorderdetails pod ON (po.pocode = pod.pocode and po.fcba = pod.fcba) where po.pocode ='"& fccode &"' and po.fcba='"& businessunit &"'"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
					  rsCallCombo.Open q, connstr
					  if not rsCallCombo.eof then
					  	response.write "<option value=''></option>"
							Do While Not rsCallCombo.eof
							 html = html & "<option value='" & rsCallCombo("linenokey") & "'" 
							 if cstr(""&Request("linenokey")) = cstr(""&rsCallCombo("linenokey") )then html = html &  " selected"
							 html = html &  " >"
							 html = html & rsCallCombo("linenokey")

							 html = html &"</option>" 
								rsCallCombo.movenext
							loop

					  else
					
					  	html = "<option value=''></option>"
					  End if
					  rsCallCombo.Close
					  set rsCallCombo = Nothing  
					  
	response.write(html)

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