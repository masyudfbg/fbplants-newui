<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
stono = Request("stono")
businessunit = Request("businessunit")
result = ""

' q = " select sp.fccode as suppliercode, sp.fcname as suppliername, bu.fccode , bu.fcname, bu.fccompanycode " & _
' 	" from supplier sp Inner join businessunit bu ON (sp.fcba = bu.fccode)" & _
' 	" where bu.fccode = '"& fccode &"' and sp.fccode='"& suppliercode &"'"

q = " select * from stocktransferorderdtl std " & _
	" inner join stocktransferorderhdr sth on (sth.pocode=std.pocode and sth.fcba=std.fcba and sth.stono = std.stono) " & _
	" where sth.fcba = '"& businessunit &"' " & _
	" and sth.stono= '"& stono &"' "

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
					  rsCallCombo.Open q, connstr
					  if not rsCallCombo.eof then
					  	response.write "<option value=''></option>"
							Do While Not rsCallCombo.eof
							 html = html & "<option value='" & rsCallCombo("stolineno") & "'" 
							 if Request("stolineno") = rsCallCombo("stolineno") then html = html &  " selected"
							 html = html &  " >"
							 html = html & rsCallCombo("stolineno")

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