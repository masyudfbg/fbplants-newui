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
Employee Last Period Close <%=FormatDateTime(Session("cls_Employee"),1)%> <br> 
Employee Current Period is <%=FormatDateTime(Session("v_Employee"),1)%> <br><br>
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
	qControlJob="select ControlJob from Module where NAME='Employee'"
	Set rsControlJob = Server.CreateObject("ADODB.Recordset")
	rsControlJob.CursorType = 3
	rsControlJob.CursorLocation = 3
	rsControlJob.Open qControlJob, connstr
	if not rsControlJob.eof then
	   txtControlJob=rsControlJob("ControlJob")
	else
	   Response.Write  "<font face=Verdana size=1 color=red>" & "SAC Plant Error SAC20025 Control Job For Employee is Empty"
	   txtControlJob=""
	end if
	response.Flush()

	if txtControlJob<>"" then

		

		' To Pick max Line No in Cost_Activity
		qMaxLineNo = " select nvl(max(nvl(lineno, 0)), 0) maxLineNo " & _
						" from empallded " & _
						" where " & _
						" fcba = '" & Session("fcBA") & "' "
						
		Set rsqMaxLineNo = Server.CreateObject("ADODB.Recordset")
		rsqMaxLineNo.CursorType = 3
		rsqMaxLineNo.CursorLocation = 3
		rsqMaxLineNo.Open qMaxLineNo, connstr

		if not rsqMaxLineNo.eof then
			maxLineNo = cdbl(rsqMaxLineNo("maxLineNo"))
		else
			maxLineNo = 0
		end if
		rsqMaxLineNo.close
		set rsqMaxLineNo = Nothing		
		'---- Normalisasi

		qNormalisasiDetail ="Delete From EMPALLDED_HIST_TAB where " & _
					 " lineheaderno||fcba in (select lineno||fcba from empallded where	to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					 " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
					 " and ALLOWDEDCODE in (select FCCODE from ALLOWDED where MAN_AUTO='A' and fcBA='" & Session("fcBA") & "') " & _
					 " and fcBA='" & Session("fcBA") & "')"
		connstr.execute(qNormalisasiDetail)
		
		qNormalisasiDetail = " delete from empallded_hist_tab" & _
		                     " where fcba = '" & Session("fcBA") & "'" & _
							 "       and lineheaderno > " & maxLineNo & ""
		connstr.execute(qNormalisasiDetail)


		qNormalisasi="Delete From EMPALLDED where  to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					 " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
					 " and ALLOWDEDCODE in (select FCCODE from ALLOWDED where MAN_AUTO='A' and fcBA='" & Session("fcBA") & "') " & _
					 " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)
		

		qNormalisasi="Delete From TAXALLDED where  to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					 " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
					 " and ALLOWDEDCODE in (select FCCODE from ALLOWDED where MAN_AUTO='A' and fcBA='" & Session("fcBA") & "') " & _
					 " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)

		qNormalisasi="Delete From TAXALLDEDLEDGER where  MONTH=" & Month(Session("v_Employee")) &  _
					 " and YEAR=" & Year(Session("v_Employee")) & _
					 " and fcBA='" & Session("fcBA") & "'"
		connstr.execute(qNormalisasi)
		
		
		sqlNormalisasiEmployeeHist = " delete from employee_hist_tab" & _
                                     " where year = " & Year(Session("v_Employee")) & "" & _
                                     "       and month = " & Month(Session("v_Employee")) & "" & _
                                     "       and fcba = '" & Session("fcBA") & "'"

		connstr.execute(sqlNormalisasiEmployeeHist)
		
		sqlNormalisasiEmployeeHist = " insert into employee_hist_tab" & _
		                             " select " & Year(Session("v_Employee")) & " year" & _
		                             "        , " & Month(Session("v_Employee")) & " month" & _
		                             "        , fccode" & _
		                             "        , fcname" & _
		                             "        , sectionname" & _
		                             "        , gangcode" & _
		                             "        , dateofbirth" & _
		                             "        , religion" & _
		                             "        , race" & _
		                             "        , identitycard" & _
		                             "        , highesteducation" & _
		                             "        , yearofgraduation" & _
		                             "        , dateofjoin" & _
		                             "        , dateofdesignation" & _
		                             "        , presentgrade" & _
		                             "        , presentgrade_detail" & _
		                             "        , presentrate" & _
		                             "        , previousgrade" & _
		                             "        , previousgrade_detail" & _
		                             "        , previousrate" & _
		                             "        , maritalstatus" & _
		                             "        , spouse" & _
		                             "        , dependentchild" & _
		                             "        , jamsostekno" & _
		                             "        , jamsostekcategory" & _
		                             "        , position" & _
		                             "        , dateterminate" & _
		                             "        , nationality" & _
		                             "        , npwp" & _
		                             "        , naturetype" & _
		                             "        , includedtaxcalculation" & _
		                             "        , taxstatus" & _
		                             "        , paymenttype" & _
		                             "        , bankname" & _
		                             "        , accountno" & _
		                             "        , ownershipofaccount" & _
		                             "        , gender" & _
		                             "        , taxpaidbycompany" & _
		                             "        , address" & _
		                             "        , bpjskes_category" & _
		                             "        , bpjskes_no" & _
		                             "        , '" & Session("CodeName") & "' fcentry" & _
		                             "        , '" & Session("CodeName") & "' fcedit" & _
		                             "        , '" & Request.ServerVariables("REMOTE_ADDR") & "' fcip" & _
		                             "        , fcba" & _
		                             "        , sysdate lastupdate" & _
		                             "        , to_char(sysdate, 'hh24:mi') lasttime" & _
		                             " from employee" & _
                                     " where fcba = '" & Session("fcBA") & "'" & _
                                     "       and dateterminate is null"

		connstr.execute(sqlNormalisasiEmployeeHist)
		
		
		%>

		
        		
		<!--#include virtual="Payroll_Gaji.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Basic Salary is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Borongan.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Gaji Borongan is successfully...." & "<br><br></font>"%>
		
		<!--#include virtual="Payroll_TunjanganKeluarga.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Tunjangan Keluarga/Beras is successfully...." & "<br><br></font>"%>
		
		<!--#include virtual="Payroll_Jamsostek.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Jamsostek is successfully...." & "<br><br></font>"%>
		
		<!--#include virtual="payroll_bpjs_kesehatan.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate BPJS Kesehatan is successfully...." & "<br><br></font>"%>
		
		<!--#include virtual="Payroll_Lembur.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Over Time / Lembur is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Mangkir.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Mangkir is successfully...." & "<br><br></font>"%>

		
		<!--#include virtual="Payroll_IjinPribadi.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Ijin Pribadi is successfully...." & "<br><br></font>"%>
		
		<!--#include virtual="Payroll_Tandan_Fixed.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi / Potongan Panen is successfully...." & "<br><br></font>"%>
		
		
		<!--#include virtual="Payroll_TandanHariLibur_Fixed.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Tandan Hari Libur is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Mandore_2.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Mandor is successfully...." & "<br><br></font>"%>

		
		<!--#include virtual="Payroll_Kerani_Buah.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Kerani Buah is successfully...." & "<br><br></font>"%>
		
		<!--#include virtual="Payroll_Mandore1_NAF.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Mandor 1 by NAF is successfully...." & "<br><br></font>"%>	
		

		<!-----------NEW ADDITIONAL KONTANAN ----->
		<!--#include virtual="Payroll_TandanHariLiburMandor.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Kontanan Mandor (Supervisor 2) is successfully...." & "<br><br></font>"%>	

		<!--#include virtual="Payroll_TandanHariLiburKrani.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Kontanan Krani (Supervisor 3) is successfully...." & "<br><br></font>"%>	
		
		<!--#include virtual="Payroll_TandanHariLiburMandor1.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Kontanan Mandor 1 (Supervisor 1) is successfully...." & "<br><br></font>"%>	
		<!-----------END OF NEW ADDITIONAL KONTANAN ----->
		
		
		<!--#include virtual="Payroll_Semprot.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Semprot is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Semprot_Mandor.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Semprot Mandor is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Pupuk.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Pupuk is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Pupuk_Mandor.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Pupuk Mandor is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Until.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Until is successfully...." & "<br><br></font>"%>


		<!--#include virtual="Payroll_Additional.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi All Additioanl is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Tunas.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Tunas is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Muat.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Muat is successfully...." & "<br><br></font>"%>

		<!--#include virtual="Payroll_Transport.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Premi Transport is successfully...." & "<br><br></font>"%>
		

		<!--#include virtual="Payroll_Potongan_Pph21.Asp"-->
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Deduction of Tax (PPh21) is successfully...." & "<br><br></font>"%>
		
			
		<%Response.Write "<font face=Verdana size=1 color=blue>" & "Calculate Payroll is successfully...." & "<br><br></font>"%>

		<%	
		qNormalisasi="Update EMPALLDED set Amount=amount*-1, lastupdate = sysdate, lasttime = to_char(sysdate,'HH24:MI') where  to_number(to_Char(ADHOCDATE,'mm'))=" & Month(Session("v_Employee")) &  _
					 " and to_number(to_Char(ADHOCDATE,'yyyy'))=" & Year(Session("v_Employee")) & _
					 " and ALLOWDEDCODE in (select FCCODE from ALLOWDED where ALLDED<>'A' and amount > 0 and fcBA='" & Session("fcBA") & "') " & _
					 " and fcBA='" & Session("fcBA") & "'"
		'Response.Write(qNormalisasi)
		connstr.execute(qNormalisasi)%>
		
		
		<%
		qProcessPayroll = "BEGIN EMPALLDEDSTATUS_API.FinishProcessPayroll (" & Year(Session("v_Employee")) & "," & _
		                                                                Month(Session("v_Employee")) & ",'" & _
																		Session("fcBA") & "','" & _
																		Session("CodeName") & "','" & _
																		Session("CodeName") & "','" & _
																		Request.ServerVariables("REMOTE_ADDR") & "'); END;"
		
		connstr.execute(qProcessPayroll)
		
		%>
		
		
			
		<script language="JavaScript" type="text/javascript" >
		document.getElementById('status_running').style.visibility = 'hidden';
		</script>
		<img src="images/ok.gif"><%
		TimeAkhir=Time()
		Lama =TimeAkhir-TimeAwal
		Response.Write "Finish ... " & time() & "<br><br>"
		Response.Write " Time Elapsed " & Hour(Lama) & " Hours " & Minute(Lama) & " Minutes " &  Second(Lama) & " Seconds."%>
	<%end if%>
<% else %>

<%
	   Response.Write  "<font face=Verdana size=1 color=red>" & "System are UNABLE to continue calculate the Payroll, Please Un-Close Financial Process "
%>
  
<% end if %>

