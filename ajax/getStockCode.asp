<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
prcode = Request("prcode")
prlineid = Request("prlineid")
fcba = Request("fcba")
result = ""

' q = " select sp.fccode as suppliercode, sp.fcname as suppliername, bu.fccode , bu.fcname, bu.fccompanycode " & _
' 	" from supplier sp Inner join businessunit bu ON (sp.fcba = bu.fccode)" & _
' 	" where bu.fccode = '"& fccode &"' and sp.fccode='"& suppliercode &"'"

' itemcode = ""
' stockname = ""
' approvedqty = ""
' remainsqty = ""

q = " select prd.itemcode, prd.description, s.fcname, prd.approvedqty, " & _
	" (prd.approvedqty - prd.orderedqty) remainsqty from purchaserequestdetails prd " & _
	" inner join purchaserequest pr ON (pr.prcode = prd.prcode and pr.fcba = prd.fcba)" & _
	" left join stock s ON (s.fccode = prd.itemcode and s.fcba = prd.fcba)" & _
	" where pr.prcode ='"& prcode &"' and prd.lineid = '"& prlineid &"' and prd.fcba='" & fcba & "'"

' response.Write q

set rsrecurtrans = server.CreateObject("Adodb.recordset")
rsrecurtrans.Open q, connstr

if Not rsrecurtrans.eof then
	itemcode = rsrecurtrans("itemcode")
	stockname = rsrecurtrans("fcname") 
	approvedqty = rsrecurtrans("approvedqty")
	remainsqty = rsrecurtrans("remainsqty")
end if
rsrecurtrans.Close
set rsrecurtrans = Nothing

response.write "{"
response.write """itemcode"" : """ & itemcode & ""","
response.write """approvedqty"" : """ & approvedqty & ""","
response.write """remainsqty"" : """ & remainsqty & """," 
response.write """stockname"" : """ & stockname & """"
response.write "}"

%>