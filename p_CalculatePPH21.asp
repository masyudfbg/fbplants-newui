<%
'Option Explicit
%>
<%
'On Error Resume Next
%>
<%response.buffer=True%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
Dim TimeAwal,TimeAkhir,Hitung,Lama, qMaxLineNo

TimeAwal=Time()
%>
<div id="status_running"><img src="ICON/working.gif"></div><br>
<font face="Verdana" size="1">
Process Last Period Close <%=MonthName(Month(Session("cls_Employee")))%>- <%=Year(Session("cls_Employee"))%> <br> 
Process Current Period is <%=MonthName(Month(Session("v_Employee")))%>-<%=Year(Session("v_Employee"))%><br><br>
Please Wait...<br>
<%
 Response.Write "Start ..." & Time() & "<br>" 
 if err.number <> 0 then
	connStr.RollbackTrans
	response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & "Error Number : " & err.number & " because : " & err.description & "</font>"
	cmdView	= "True"%>
    <script language="JavaScript" type="text/javascript" >
	document.getElementById('status_running').style.visibility = 'hidden';
	</script>
	<img src="images/error.gif"><%
 End if

dim rsEmpAllDed,rsPO,qReceive,qPO,qControlJob,nPurchaseControl,rsControlJob,txtControlJob,qAllowance,nQtyReceived,qNormalisasi,nPresentRate,nDays,nAmount
dim v_AllowanceCode, maxLineNo

Dim KAPTHLM  'KAPTHLM  = Kode Allowded Premi Tandan Hari Libur Mandor
Dim KAPTHL   'KAPTHL   = Kode Allowded Premi Tandan Hari Libur Anak Buah
Dim KAPPTHLM 'KAPPTHLM = Kode Allowded Potongan Premi Tandan Hari Libur Mandor

Dim nRateBagi 
Dim nFaktorKali 
Dim nPremiTandanHariLibur  


'-- Check for Empallded PreProcess
qCheckStatus = "select EMPALLDEDSTATUS_API.GetCheckPreProcess (" & Year(Session("v_Employee")) & "," & _
		                                                                Month(Session("v_Employee")) & ",'" & _
																		Session("fcBA") & "') closingStatus from dual "

Set rsCheckStatus = Server.CreateObject("ADODB.Recordset")
rsCheckStatus.CursorType = 3
rsCheckStatus.CursorLocation = 3
'response.write qCheckStatus
rsCheckStatus.Open qCheckStatus, connstr
if not rsCheckStatus.eof then
   txtClosingStatus = rsCheckStatus("closingStatus")
else
   Response.Write  "<font face=Verdana size=1 color=red>" & "EmpalldedStatus are not defined "
   txtClosingStatus="FALSE"
end if
rsCheckStatus.close
set rsCheckStatus = Nothing



Dim sqlNormalisasiEmployeeHist

IF txtClosingStatus = "TRUE" THEN  

	'---- Control Job
	qControlJob = " SELECT parameterdetail_api.get_param_1 ('" & Session("FCBA") & "', 'HARDCODE', 'CONTROLJOB_MODULE', 'EMPLOYEE') ControlJob FROM DUAL"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   txtControlJob=""
	end if

	if len(trim(""&txtControlJob)) <> 0 then

		qNormalisasi="Delete From TAXALLDED where  to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					 " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
					 " and ALLOWDEDCODE in (select FCCODE from ALLOWDED where MAN_AUTO='A' and fcBA='" & Session("fcBA") & "') " & _
					 " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)

		qNormalisasi="Delete From TAXALLDEDLEDGER where  MONTH=" & Month(Session("v_Employee")) &  _
					 " and YEAR=" & Year(Session("v_Employee")) & _
					 " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)
		
		%>
		
		
		<!--#include virtual="Payroll_Potongan_Pph21.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate  Tax (PPh21) is successfully...." & "<br><br></font>"%>
		
			
		
		<%	
		'qNormalisasi="Update EMPALLDED set Amount=amount*-1, lastupdate = sysdate, lasttime = to_char(sysdate,'HH24:MI') where  to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
		'			 " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
	'				 " and ALLOWDEDCODE in (select FCCODE from ALLOWDED where ALLDED<>'A' and amount > 0 and fcBA='" & Session("fcBA") & "') " & _
	'				 " and fcBA='" & Session("fcBA") & "'"
		'Response.Write(qNormalisasi)
		'connstr.execute(qNormalisasi)
		%>
		
		
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		<img src="images/ok.gif"><%
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."%>
	<%else%>
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
	<%
		TimeAkhir=Time()
		Lama = TimeAkhir-TimeAwal
		Response.Write "Finish ... " & time() & "<br><br>"
		Response.Write "<br/><img src='images/error.gif'/><font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For Employee is Empty"
		response.Flush()
	%>
	<%end if%>
<% else %>
	<%Response.Write  "<font face=Verdana size=1 color=red>" & "System are UNABLE to continue calculate the Payroll, Please Un-Close Financial Process "%>
<% end if %>