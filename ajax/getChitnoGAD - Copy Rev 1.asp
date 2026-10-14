<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
targetcode = Request("targetcode")
fddate = Request("fddate")
chitno = Request("chitno")
result = ""

html = html & "<select id='chitnoharvesting' name='chitno' onchange='javascript:getBjrSpbHarvesting();' class='chosen-select' style='width: 100%;'>"
html = html &"<option value=''></option>"

q = " SELECT distinct wb.spbno, spb.fieldcode, spb.receptiondate" & _
	"   FROM wb_harvesting wb" & _
	"        INNER JOIN harvestingspb spb" & _
	"           ON (wb.spbno = spb.spbno AND wb.fcba_estate = spb.fcba)" & _
	"        LEFT JOIN periodcontrol pc" & _
	"           ON (    TO_NUMBER (TO_CHAR (spb.receptiondate, 'mm')) = pc.currentaccountingperiod" & _
	"               AND TO_NUMBER (TO_CHAR (spb.receptiondate, 'yyyy')) = pc.currentaccountingyear AND spb.fcba = pc.fcba)" & _
	"  WHERE     fcba_mill = 'PSR'" & _
	"        AND fcba_estate = '"& session("fcba") &"'" & _
	"        AND fieldcode = '"& targetcode &"'" & _
	"        AND receptiondate = TO_DATE ('"& fddate &"', 'mm/dd/yyyy')" & _
	" 		 AND wb.own = 'OWN'" & _
	" ORDER BY wb.spbno, spb.receptiondate"
  

' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
	response.write "<option value=''></option>"
	Do While Not rsCallCombo.eof
		html = html & "<option value='" & rsCallCombo("spbno") & "'"
		if "" & chitno <> "" then
			if cstr(""&chitno) = cstr(""&rsCallCombo("spbno")) then html = html &  " selected"
		end if
		html = html &  " >"
		html = html & rsCallCombo("spbno") 
		html = html &"</option>" 
		rsCallCombo.movenext
	loop

End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>