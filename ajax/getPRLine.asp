<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
prcode = Request("prcode")
fcba = Request("fcba")
prlineid = Request("prlineid")
pocode = Request("pocode")
fccompanycode = Request("fccompanycode")
result = ""

html = html & "<select id='prlineid' name='prlineid' onchange='javascript:getStockCode();javascript:getselectedprline();' style='width: 40%' class='chosen-select'>"
html = html &"<option value=''></option>"

' q = " SELECT * FROM purchaserequestdetails prd " & _
' 	" INNER JOIN purchaseorderdetails po ON (po.prcode = prd.prcode AND po.fcba = prd.fcba)" & _ 
' 	" INNER JOIN businessunit bu ON (bu.fccode = prd.fcba)" & _
' 	" INNER JOIN company co ON (co.companycode = bu.fccompanycode)" & _
' 	" WHERE prd.prcode='"& prcode &"' AND bu.fccompanycode = '"& fccompanycode &"'" 

' ' response.Write q

' Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
' rsCallCombo.Open q, connstr
' if not rsCallCombo.eof then
' 	Do While Not rsCallCombo.eof
' 		pocode = rsCallCombo("pocode")
' 		fcba = rsCallCombo("fcba")
' 	rsCallCombo.movenext
' 	loop

' End if
' rsCallCombo.Close
' set rsCallCombo = Nothing

' " MINUS" & _
	' " SELECT po.prlineid" & _
	' "  FROM purchaseorderdetails po" & _
	' "       INNER JOIN businessunit bu ON (bu.fccode = po.fcba)" & _
	' "       INNER JOIN company co ON (co.companycode = bu.fccompanycode)" & _
	' " WHERE po.prcode = '"& prcode &"' AND bu.fccompanycode ='"& fccompanycode &"') ORDER BY lineid"

q = " SELECT prd.lineid, prd.itemcode, stock_api.get_stockname(prd.itemcode, prd.fcba) itemname" & _
	"  FROM purchaserequest pr" & _
	"       INNER JOIN purchaserequestdetails prd" & _
	"          ON (pr.prcode = prd.prcode AND pr.fcba = prd.fcba)" & _
	"       INNER JOIN businessunit bu ON (bu.fccode = prd.fcba)" & _
	"       INNER JOIN company co ON (co.companycode = bu.fccompanycode)" & _
	" WHERE approvedqty - orderedqty > 0 AND pr.prcode = '"& prcode &"' AND pr.fcba = '"& fcba &"' ORDER BY prd.lineid" 
	

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr

if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("lineid") & "'"
		if "" & prlineid <> "" then
			if cstr(""&prlineid) = cstr(""&rsCallCombo("lineid")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("lineid") & " : " & rsCallCombo("itemcode") &" - "& rsCallCombo("itemname")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>