<!--#include file="../connDb1.Inc"-->
<%
	txtfcbasource      = Request("fcbaSource")
    Dim html
	response.expires=-1
	
	stringCheck = " select distinct  a.orderno from TRANSPORTTASKDTL_TAB  a where fcba='" & txtfcbasource & "'  and  a.fcbadistination='" & txtfcbadistination & "'  and (quantitytransfered -quantityreceived ) > 0"

					  Set rsCallCombo = Server.CreateObject("ADODB.Recordset")
					  rsCallCombo.Open stringCheck, connstr
					  if not rsCallCombo.eof then
							html = "<select name='TransportTaskNo'>"
							Do While Not rsCallCombo.eof
							html = html & "<option value='" & rsCallCombo("orderno") & "'>" & rsCallCombo("orderno") & "</option>" & _
								rsCallCombo.movenext
							loop
							html = html & "</select>" 
							else
							
							html = "<select name='TransportTaskNo'></select>"
					  End if
					  rsCallCombo.Close
					  set rsCallCombo = Nothing  
					  
	response.expires=-1   
	response.write(html)
 
 
 %>