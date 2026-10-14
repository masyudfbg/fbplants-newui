<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		 
orderdate = Request("orderdate")
documentno = Request("documentno")
lineby = Request("lineby")
itemcode = Request("itemcode")
storecode = Request("storecode")
result = ""


' q = " SELECT 'GRN' type FROM DUAL" & _
' 	" UNION ALL" & _
' 	" SELECT 'RCT' type FROM DUAL" & _
' 	" UNION ALL" & _
' 	" SELECT 'WRO' type FROM DUAL"   

' ' response.write q
' set rsrecurtrans = server.CreateObject("Adodb.recordset")
' rsrecurtrans.Open q, connstr

' if Not rsrecurtrans.eof then
' 	typex=  rsrecurtrans("type")
' end if


' If typex = "GRN" then
	q = " SELECT rnd.purchaselinenokey as lineby" & _
		"   FROM receivenotedetail rnd" & _
		"        INNER JOIN receivenote rn" & _
		"           ON (rn.receivenotecode = rnd.receivenotecode AND rn.fcba = rnd.fcba)" & _
		"  WHERE     rn.fcba = '"& session("fcba") &"'" & _
		"        AND rn.receivenotecode = '"& documentno &"'" & _
		"        AND rnd.locationcode = '"& storecode &"'" & _
		"        AND rnd.purchaseitemcode = '"& itemcode &"'" & _ 
		"        AND rndate <= TO_DATE ('"& orderdate &"', 'mm/dd/yyyy')" & _
		"	ORDER BY rnd.purchaselinenokey"

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

' Else If typex = "RCT" then
' 	q = " SELECT rch.orderdate," & _
' 		"        rct.orderno," & _
' 		"        rct.stockcode," & _
' 		"        rct.pricereceived," & _
' 		"        rct.linenosource as lineby" & _
' 		"   FROM receivedtaskdtl_tab rct" & _
' 		"        INNER JOIN receivedtask_tab rch" & _
' 		"           ON (rct.orderno = rch.orderno AND rct.fcba = rch.fcba)" & _
' 		"  WHERE     rct.fcba = '"& session("fcba") &"'" & _
' 		"        AND rch.orderdate <= TO_DATE ('"& orderdate &"', 'mm/dd/yyyy') " & _
' 		"        AND rct.orderno = '"& documentno &"'"& _
' 		"        AND rct.locationcode = '"& storecode &"'" & _
' 		"        AND rct.stockcode = '"& itemcode &"'" & _  
' 		"		ORDER BY rct.orderno, rct.linenosource"  

' 	' response.Write q
' 	Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
' 	rsCallCombo.Open q, connstr

' 	if not rsCallCombo.eof then
' 		response.write "<option value=''></option>"
' 		Do While Not rsCallCombo.eof
' 		 html = html & "<option value='" & rsCallCombo("lineby") & "'" 
' 		 if cdbl(to_nol(Request("lineby"))) = cdbl(to_nol(rsCallCombo("lineby"))) then html = html &  " selected"
' 		 html = html &  " >"
' 		 html = html & rsCallCombo("lineby")

' 		 html = html &"</option>" 
' 			rsCallCombo.movenext
' 		loop
' 	else
' 		html = "<option value=''></option>"
' 	End if
' 	rsCallCombo.Close
' 	set rsCallCombo = Nothing  
						  
' 	response.write(html)	

' Else

' 	q = "   SELECT wo.wodate," & _
' 		"          wo.writeoffcode," & _
' 		"          wod.stockcode," & _
' 		"          (SUM (wod.amount) / SUM (wod.quantity)) price," & _
' 		"          wod.linenokey as lineby" & _
' 		"     FROM writeoffdetails wod" & _
' 		"          INNER JOIN writeoff wo" & _
' 		"             ON (wod.fcba = wo.fcba AND wod.writeoffcode = wo.writeoffcode)" & _
' 		"    WHERE     wod.fcba = '"& session("fcba") &"'" & _
' 		"          AND wod.writeoffcode = '"& documentno &"'" & _
' 		"          AND wod.stockcode = '"& itemcode &"'" & _
' 		"          AND wo.storecode = '"& storecode &"'" & _ 
' 		"          AND wo.wodate <=  TO_DATE ('"& orderdate &"', 'mm/dd/yyyy') " & _
' 		" GROUP BY wo.wodate," & _
' 		"          wo.writeoffcode," & _
' 		"          wod.stockcode," & _
' 		"          wod.linenokey" 

' 	' response.write q
' 	set rsCallCombo = server.CreateObject("Adodb.recordset")
' 	rsCallCombo.Open q, connstr

' 	if not rsCallCombo.eof then
' 		response.write "<option value=''></option>"
' 		Do While Not rsCallCombo.eof
' 		 html = html & "<option value='" & rsCallCombo("lineby") & "'" 
' 		 if cdbl(to_nol(Request("lineby"))) = cdbl(to_nol(rsCallCombo("lineby"))) then html = html &  " selected"
' 		 html = html &  " >"
' 		 html = html & rsCallCombo("lineby")

' 		 html = html &"</option>" 
' 			rsCallCombo.movenext
' 		loop
' 	else
' 		html = "<option value=''></option>"
' 	End if
' 	rsCallCombo.Close
' 	set rsCallCombo = Nothing  
						  
' 	response.write(html)	

' End If
' End If 
%>