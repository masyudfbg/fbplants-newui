<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

Dim TimeAwal,TimeAkhir,Hitung,Lama, qOpenEmpalldedStatus
TimeAwal=Time()

%>
	<div id="status_running"><img src="ICON/working.gif"></div><br/>
	<font face="Verdana" size="1">
	GL Last Period Close <%= Mid(Session("CurP_Costbook"),5,2) %><%=FormatDateTime(Session("cls_Costbook"),1)%> <br/> 
	GL Current Period is <%=FormatDateTime(Session("v_Costbook"),1)%> <br/><br/>
	Please Wait...
	<br/>
<%
Response.Write "Start ..." & Time() & "<br/>" 


Response.Write "to open period cost book Please wait, this action take several minute.." & "<br/>"																

'Start Codes for Update - Remove Columns
qNormalisasiDetail_UPDATE = " UPDATE iplasprod.periodcontrol" & _
							"    SET lastperiodposted =" & _
							"           (SELECT CASE" & _
							"                      WHEN currentaccountingperiod - 1 BETWEEN 1 AND 12" & _
							"                      THEN" & _
							"                         (SELECT LAST_DAY (ADD_MONTHS (lastperiodposted, -1))" & _
							"                            FROM periodcontrol" & _
							"                           WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook')" & _
							"                      ELSE" & _
							"                         lastperiodposted" & _
							"                   END" & _
							"                      AS lastperiodposted" & _
							"              FROM periodcontrol" & _
							"             WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook')," & _
							"        currentdate =" & _
							"           (SELECT CASE" & _
							"                      WHEN currentaccountingperiod - 1 < 12" & _
							"                      THEN" & _
							"                         (SELECT TRUNC (ADD_MONTHS (currentdate, -1), 'month')" & _
							"                            FROM periodcontrol" & _
							"                           WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook')" & _
							"                      ELSE" & _
							"                         currentdate" & _
							"                   END" & _
							"                      AS currentdate" & _
							"              FROM periodcontrol" & _
							"             WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook')," & _
							"        currentaccountingyear =" & _
							"           (SELECT CASE" & _
							"                      WHEN currentaccountingperiod - 1 = 0" & _
							"                      THEN" & _
							"                         currentaccountingyear - 1" & _
							"                      ELSE" & _
							"                         currentaccountingyear" & _
							"                   END" & _
							"                      AS currentaccountingyear" & _
							"              FROM periodcontrol" & _
							"             WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook')," & _
							"        currentaccountingperiod =" & _
							"           (SELECT CASE" & _
							"                      WHEN currentaccountingperiod - 1 = 0 THEN 15" & _
							"                      ELSE currentaccountingperiod - 1" & _
							"                   END" & _
							"                      AS currentaccountingperiod" & _
							"              FROM periodcontrol" & _
							"             WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook')," & _
							"        closeaccountingyear =" & _
							"           (SELECT CASE" & _
							"                      WHEN currentaccountingperiod - 1 = 1" & _
							"                      THEN" & _
							"                         closeaccountingyear - 1" & _
							"                      ELSE" & _
							"                         closeaccountingyear" & _
							"                   END" & _
							"                      AS closeaccountingyear" & _
							"              FROM periodcontrol" & _
							"             WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook')," & _
							"        closeaccountingperiod =" & _
							"           (SELECT CASE" & _
							"                      WHEN currentaccountingperiod - 1 = 1 THEN 15" & _
							"                      ELSE closeaccountingperiod - 1" & _
							"                   END" & _
							"                      AS closeaccountingperiod" & _
							"              FROM periodcontrol" & _
							"             WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook')" & _
							"  WHERE fcba = '" & session("fcba") & "' AND SYSTEM = 'Costbook'" 

connstr.execute(qNormalisasiDetail_UPDATE)
connStr.execute("Commit")
%>

<br/><br/><img src="images/ok.gif">Process Open Period Control are completed...

<%
Response.Write Session("USERNAME")
%>

<%
 if err.number <> 0 then
      
	connStr.RollbackTrans
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif> " & err.description & "</font>"
	cmdView	= "True"%>
    <script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<img src="images/error.gif"><%
 End if
%>
 <script language="JavaScript" type="text/javascript" >
 alert("Open Period Costbook Behasil.. \n Silahkan Login Kembali");
 window.location = "http://103.108.126.134/logout.asp";
 </script>
