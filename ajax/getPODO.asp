<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
fcba = Request("fcba")
result = ""

' q = " select sp.fccode as suppliercode, sp.fcname as suppliername, bu.fccode , bu.fcname, bu.fccompanycode " & _
' 	" from supplier sp Inner join businessunit bu ON (sp.fcba = bu.fccode)" & _
' 	" where bu.fccode = '"& fccode &"' and sp.fccode='"& suppliercode &"'"

q = " SELECT DISTINCT dtl.pocode" & _
	" FROM (" & _
	"     SELECT dtl.fcba, dtl.stono, dtl.stolineno" & _
	"     FROM STOCKTRANSFERORDERHDR hdr" & _
	"     INNER JOIN STOCKTRANSFERORDERDTL dtl ON (hdr.stono = dtl.stono AND hdr.fcba = dtl.fcba)" & _
	"     WHERE dtl.fcba = '"& fcba &"' AND dtl.do_quantity < dtl.quantityorder" & _
	"     MINUS" & _
	"     SELECT dtl.fcba, dtl.stono, dtl.stolineno" & _
	"     FROM DELIVERYORDERPURCHASEDTL dtl" & _
	"     WHERE dtl.fcba = '"& fcba &"'" & _
	" ) data" & _
	" INNER JOIN STOCKTRANSFERORDERDTL dtl ON (data.stono = dtl.stono AND data.fcba = dtl.fcba AND data.stolineno = dtl.stolineno) ORDER BY dtl.pocode"

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
					  rsCallCombo.Open q, connstr
					  if not rsCallCombo.eof then
					  	response.write "<option value=''></option>"
							Do While Not rsCallCombo.eof
							 html = html & "<option value='" & rsCallCombo("pocode") & "'" 
							 if Request("pocode") = rsCallCombo("pocode") then html = html &  " selected"
							 html = html &  " >"
							 html = html & rsCallCombo("pocode") 

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