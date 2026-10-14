<!--#include file="../connWBS.Inc"-->
<!--#include file="../txtFunction.inc"-->

<%
if SESSION("username") = "" then %>

	<tr STYLE="BACKGROUND-COLOR:#FAFAFA;">
		<td colspan="4" align="CENTER">SESSION ANDA TELAH HABIS, SILAHKAN <a target="_blank" class="button_menu_top" style="color:#fff;" href="logout.asp">LOGIN</a> KEMBALI </td>
	</tr>
	
<%
else

FCBA_Q = "select b.fccode from IPLASPROD2.user_fcba_tab a, IPLASPROD2.businessunit b where a.fcba = b.fccode and a.fcusercode = '" & SESSION("username") & "' "
txtmonth = Lpad(month(date),"0","2")
txtDay = Lpad(Day(date),"0","2")
'response.Write SESSION("username")
Dim ERRORDATA,no,fcbaid,CHECK
   set ERRORDATA = server.CreateObject("Adodb.recordset")
   fcbaid = "and fcba IN (" & FCBA_Q & ") "
   
   q=	"select * from pesan_error_tab where nomer in (select max(a.nomer) " & _
		"from pesan_error_tab a " & _
		"where fcba is not null " & _
		"group by a.fcba) " & _
		fcbaid 
		
						
   'q="select a.* from pesan_error_tab a inner join (" & _
    ' "SELECT " & _
     '"fcba,max(tanggalerror) tanggalerror " & _
	 '"FROM pesan_error_tab where to_char(tanggalerror,'MM/DD/YYYY') = '" &x& "/" &day(date)& "/" & year(date) & "' " & fcbaid & _ 
	 '"group by fcba) b on (a.fcba = b.fcba and a.tanggalerror = b.tanggalerror) " & _
	 '"order by a.tanggalerror desc "
   'response.Write q	
   Set ERRORDATA = Server.CreateObject("ADODB.Recordset")
   ERRORDATA.Open q, connstr
%>
<%
							no = 1
							if not ERRORDATA.eof then %>
							
							<%
								while not ERRORDATA.eof
								SpiltName = Split(ERRORDATA("PESAN_EROR"), ";") 
								jml = UBound(SpiltName)
								jmlmssg = jml-1
								
								CHECK_SATUS = "SELECT CHITNO FROM IPLASPROD2.Weighbridgeticket WHERE CHITNO = '" & SpiltName(0) & "' AND FCBA = '" & ERRORDATA("FCBA") & "'"
								'RESPONSE.WRITE CHECK_SATUS
								Set CHECK = Server.CreateObject("ADODB.Recordset")
								CHECK.Open CHECK_SATUS, connstr
								if not CHECK.eof then %>
									
								<% 
								ELSE 
								%>
								<audio autoplay="autoplay">
									<source src="../ajax/sounds-948-just-like-magic.mp3" type="audio/mpeg" />
									<embed hidden="true" autostart="true" loop="false" src="../ajax/sounds-948-just-like-magic.mp3" />
								</audio>   
						<tr class="alert" style="background-color:#ff0000;color:#fff;">
								<td align="Left" ><font face="Verdana" size="1">
								<%=no%>
								</font>
								<td align="Left" ><font face="Verdana" size="1"><%=ERRORDATA("FCBA")%></td>
								<td align="Left" ><font face="Verdana" size="1"><%
								
								%>
								<div style="font-weight:bold;font-size:12px;"><%=SpiltName(0)%></div>
								
								<div style=' text-decoration: blink;'>
								<%
								For i = 1 To jmlmssg
								  response.write ("<font color='#fff'>Error " & i & ":</font>&nbsp;&nbsp;" & SpiltName(i) & "<br/>")
								Next
								%>
								</div>
								</td>
								<td align="Left" WIDTH="150" ><font face="Verdana" size="1"><%=ERRORDATA("TANGGALERROR")%></td>
						</tr>
						<%end if
							CHECK.close()
						CHECK=null%>
									
						  <%
						
							  ERRORDATA.MoveNext
							  'response.Flush()
							  no=no+1
							Wend
							else %>
							<tr STYLE="BACKGROUND-COLOR:#FAFAFA;">
								<td colspan="4" align="CENTER">
									<div  style="background-color:#FAFAFA;border:1px #ccc solid;padding-top:10px;padding-bottom:20px;">
									<div style="text-align:center;">PROCESSING UPLOAD DATA<BR/><img src="IMAGES/loader2.gif"></div>
								</div>	
								</td>
							</tr>
								
						<%	
							end if
							ERRORDATA.close()
						ERRORDATA=null
						
						
end if%>