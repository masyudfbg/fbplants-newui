<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/adovbs.Inc"-->
<!--#include virtual="/txtfunction.Inc"-->
<%		
transporttaskno = Request("transporttaskno")
xcode = Request("xcode")

html = html & "<select id='transporttaskno' name='transporttaskno' onchange='javascript:getselectedtransporttask();get_Allowentry();' class='chosen-select' style='width: 50%;'>"
html = html &"<option value=''></option>"

q = " select DISTINCT a.orderno, TO_CHAR(a.orderdate, 'MM/DD/YYYY') as orderdate from transporttask_tab a " & _
	" inner join transporttaskdtl_tab b on a.ORDERNO = b.ORDERNO " & _
	" WHERE b.fcba = '" & xcode & "' " & _
	" AND b.fcbadistination = '"& session("fcba") &"' " & _
	" AND a.rowstateid = 'Approved' " & _
	" AND a.postingflg = 1 " & _
	" AND(b.quantitytransfered - b.quantityreceived) > 0 " & _
	" ORDER BY a.orderno"
	' response.Write q

Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
rsCallCombo.Open q, connstr
if not rsCallCombo.eof then
    response.write "<option value=''></option>"
    Do While Not rsCallCombo.eof
        html = html & "<option value='" & rsCallCombo("orderno") & "'"
        if "" & transporttaskno <> "" then
            if transporttaskno = cstr(""&rsCallCombo("orderno")) then html = html & " selected"
        end if
        html = html & " >"
        html = html & rsCallCombo("orderno") & " | Tanggal : " & rsCallCombo("orderdate")
        html = html & "</option>" 
        rsCallCombo.movenext
    loop
End if
rsCallCombo.Close
set rsCallCombo = Nothing

html = html &"</select>"
response.write(html)

%>