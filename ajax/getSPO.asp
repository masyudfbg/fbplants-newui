<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccode = Request("fccode")
businessunit = Request("businessunit")
result = ""

q = " SELECT DISTINCT pocode" & _
	" FROM (" & _
	"     SELECT spo.fcba, spo.pocode, spo.linenokey, NVL(spo.spo_amount,0) spo_amount, NVL(inv.inv_amount,0) inv_amount" & _
	"     FROM (" & _
	"         SELECT fcba, pocode, linenokey, NVL(SUM(amountcharge),0) spo_amount" & _
	"         FROM servicepurchaseorderdetails spo" & _
	"         WHERE fcba = '"& businessunit &"'" & _
	"         GROUP BY fcba, pocode, linenokey) spo" & _
	"     LEFT JOIN ( SELECT fcba, pocode, polinenokey, NVL(SUM(totalamount),0) inv_amount" & _
	"                 FROM invoiceservicedtl spo" & _
	"                 WHERE fcba = '"& businessunit &"'" & _
	"                 GROUP BY fcba, pocode, polinenokey) inv ON (spo.fcba = inv.fcba AND spo.pocode = inv.pocode AND spo.linenokey = inv.polinenokey)" & _
	" )" & _
	" WHERE inv_amount < spo_amount" & _
	" ORDER BY pocode"


' q = " select * from servicepurchaseorder where fcba ='"& businessunit &"' order by pocode"

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