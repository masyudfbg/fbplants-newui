<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
documentno = Request("documentno")
orderdate = Request("orderdate")
result = ""

q = " SELECT rch.orderdate," & _
	"        rct.orderno," & _
	"        rct.stockcode," & _
	"        rct.pricereceived," & _
	"        rct.linenosource as lineby" & _
	"   FROM receivedtaskdtl_tab rct" & _
	"        INNER JOIN receivedtask_tab rch" & _
	"           ON (rct.orderno = rch.orderno AND rct.fcba = rch.fcba)" & _
	"  WHERE     rct.fcba = '"& session("fcba") &"'" & _
	"        AND rch.orderdate <= TO_DATE ('"& orderdate &"', 'mm/dd/yyyy') " & _
	"        AND rct.orderno = '"& documentno &"'"& _ 
	"		ORDER BY rct.orderno, rct.linenosource"  


' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
	 html = html & "<option value='" & rsCallCombo("lineby") & "'" 
	 if cdbl(to_nol(Request("lineby"))) = cdbl(to_nol(rsCallCombo("lineby"))) then html = html &  " selected"
	 html = html &  " >"
	 html = html & rsCallCombo("lineby")

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