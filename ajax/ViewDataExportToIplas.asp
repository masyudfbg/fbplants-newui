<!--#include file="../connWBS.Inc"-->
<!--#include file="../txtFunction.inc"-->
<style>
*{font-size:10px;}
</style>


<% 
if LEN(TRIM(SESSION("username"))) > 0 then 

dim cek_to_srv,no
d=cdate(date)
set cek_to_srv = server.CreateObject("Adodb.recordset")
Q_tosrv   = "select * " & _
			"from wbs_datatbs_tab " & _
			"where wbsid||siteid in (select max(a.wbsid)||a.siteid " & _

			"				from wbs_datatbs_tab a " & _

			"			   where siteid is not null " & _

			"			   group by a.siteid) " & _

			"and siteid IN ( " & _
			"select b.fccode " & _
			"from IPLASPROD2.user_fcba_tab a, " & _
			"IPLASPROD2.businessunit b " & _
			"where a.fcba = b.fccode " & _
			"and a.fcusercode = '" & SESSION("username") & "' ) " 


Set cek_to_srv = Server.CreateObject("ADODB.Recordset")
'response.write Q_tosrv
cek_to_srv.Open Q_tosrv, connstr
if not cek_to_srv.eof then 
no = 1
	while not cek_to_srv.eof %>
	
<p>Export Data TBS: <%=cek_to_srv("siteid")%> Succsess <%=FormatDateTime(d,2)%>, Data terakhir : <%=cek_to_srv("datein")%>&nbsp;Time&nbsp;:&nbsp;<%=cek_to_srv("timein")%></p>

<% 
cek_to_srv.MoveNext
	'response.Flush()
	no=no+1
   Wend
end if
	cek_to_srv.close()
cek_to_srv=null
%>
	<p>Proccess...<span class="blink_me">|</span></p>
<%
	else
%>
	<p>Session Log Off...<span class="blink_me">|</span></p>
	
<%

end if

%>
