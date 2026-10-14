<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
customercode = Request("customercode")
fcba = Request("fcba")
salescontractno = Request("salescontractno")
result = ""

q = " SELECT  sah.customercode, sah.salescontractno, sad.vat " & _
" FROM SALESAGREEMENTHDR sah" & _
" INNER JOIN SALESAGREEMENTDET sad ON (sah.salescontractno = sad.salescontractno AND sah.fcba = sad.fcba)" & _
" WHERE sah.customercode= '"& customercode &"' and sah.fcba = '" & fcba & "'" 


'response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
					  rsCallCombo.Open q, connstr
					  if not rsCallCombo.eof then
					  	response.write "<option value=''></option>"
							Do While Not rsCallCombo.eof
							 html = html & "<option value='" & rsCallCombo("salescontractno") & "'" 
							 if cstr(""&salescontractno) = cstr(""&rsCallCombo("salescontractno")) then html = html &  " selected"
							 html = html &  " >"
							 html = html & rsCallCombo("salescontractno")

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