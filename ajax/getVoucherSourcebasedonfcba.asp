<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
fccompanycode = Request("fccompanycode")
fcbasource = Request("fcbasource")
transdate = Request("transdate")
vouchersource = Request("vouchersource")
result = ""

html = html & "<select id='vouchersource' name='vouchersource' class='chosen-select' style='width:40%;' onchange='javascript:getselectedvoucher();'>"
html = html &"<option value=''></option>"

q = " SELECT DISTINCT pv.vouchercode" & _
	"  FROM paymentvoucher pv" & _
	"       INNER JOIN paymentvoucherdetail pvd" & _
	"          ON (pv.vouchercode = pvd.vouchercode AND pv.fcba = pvd.fcba)" & _
	"       INNER JOIN intraco_tab it ON (it.fccode = pvd.jobcode)" & _
	" WHERE     pv.fcba = '"& fcbasource &"'" & _
	"       AND pvd.chargejob IS NOT NULL" & _
	"       AND pvd.chargetype IS NOT NULL" & _
	"       AND pvd.chargecode IS NOT NULL" & _
	"       AND it.fcba = '"& session("fcba") &"'" & _
	" MINUS" & _
	" SELECT vouchersource" & _
	"  FROM approvalnotadebet" & _
	" WHERE fcba = '"& session("fcba") &"' AND fcbasource = '"& fcbasource &"'" 
  

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("vouchercode") & "'"
		if "" & vouchersource <> "" then
			if cstr(""&vouchersource) = cstr(""&rsCallCombo("vouchercode")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("vouchercode")

		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>