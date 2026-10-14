<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtDATECREATED_From= trim(request("DATECREATED_From"))
v_txtDATECREATED_To=trim(request("DATECREATED_To"))
v_txtBANKCODE=trim(request("BANKCODE"))
txtControlJob=""
nOpening=0

q="SELECT CONTROLJOB FROM BANK WHERE FCCODE ='" & v_txtBANKCODE & _
  "' AND  FCBA= '" & Session("fcBA") & "'"
Set rsCARI = Server.CreateObject("ADODB.Recordset")
rsCARI.CursorType = 3
rsCARI.CursorLocation = 3
rsCARI.Open q, connstr	 
if Not rsCARI.eof Then
   txtControlJob=rsCARI("CONTROLJOB")
else
   txtControlJob=""   
End if
rsCari.Close
set rsCari=nothing

q="SELECT nvl(sum(OPENING_BALANCE),0) as OPENING_BALANCE  FROM TRIALBALANCE WHERE FCCODE ='" & txtControlJob & _
  "' AND  FCBA= '" & Session("fcBA") & "'" & _
  "  AND PERIOD_MONTH = to_char(to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY'),'MM')" & _
  "  AND PERIOD_year = to_char(to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY'),'YYYY')"
Set rsCARI = Server.CreateObject("ADODB.Recordset")
rsCARI.CursorType = 3
rsCARI.CursorLocation = 3
'response.Write q
response.Flush()
rsCARI.Open q, connstr	 
if Not rsCARI.eof Then
   nOpening=rsCARI("OPENING_BALANCE")
else
   nOpening=0
End if
rsCari.Close
set rsCari=nothing

' Amount diminuskan
' Karena kalau debet adalah kas bekurang
' Kalau Kredit maka kas bertambah
'-------------------------------- Control Job  -----------------------------------------------------------------------------------
txtSQL1 = "SELECT p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode, nvl(sum(pay.amount)*-1,0) as amount, pay.REFERENCE, pay.remarks" & _
		 " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
		 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
		 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" & _
		 " AND pay.fcba='" & session("fcBA") & "'"
if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL1 = txtSQL1 & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect1 = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL1 =txtSQL1 &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect1 = txtSelect1 & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtBANKCODE)) > 0 then
    txtSQL1 = txtSQL1 & " and (p.BANKCODE='" & v_txtBANKCODE & "')"
	txtSelect1 = txtSelect1 & " Bank Code " & v_txtBANKCODE
end if
txtsql1=(txtsql1) & " group by p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode,pay.REFERENCE, pay.remarks"

'-------------------------------- Jurnal ----------------------------------------------------------------------------------------

txtSQL2 = "SELECT p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode, nvl(sum(pay.amount),0) as amount, pay.REFERENCE, pay.remarks" & _
		 " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
		 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
		 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" & _
		 " AND pay.fcba='" & session("fcBA") & "'"
if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL2 = txtSQL2 & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect2 = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL2 =txtSQL2 &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect2 = txtSelect2 & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtBANKCODE)) > 0 then
    txtSQL2 = txtSQL2 & " and (pay.jobcode='" & txtControlJob & "')"
	txtSelect2 = txtSelect2 & " Bank Code " & v_txtBANKCODE
end if

txtsql2=(txtsql2) & " group by p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode,pay.REFERENCE, pay.remarks"

txtsql = txtsql1 & " union all " & txtsql2
txtSelect = txtSelect1
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected All Data"
end if
'response.Write(txtsql)
response.Flush()
%>
 <table width="700">
  <tr> 
    <td rowspan="5" width="200"><img src="../IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="150">COMPANY</td><td  width="300"><b><FONT color="#000099">: <%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="150">BUSINESS UNIT</td><td  width="300"><b><FONT color="#000099">: <%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="150">REPORT</td><td  width="300"><b><FONT color="#000099">: LAPORAN ARUS KAS</FONT></td></tr>
	
 </table>
<BR><BR>
<table Cellpadding="2" cellspacing="0">
	<tr>
    <td colspan="11" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" ><font face="Verdana" size="1"><%=txtSelect%></td>
	</tr>
	<tr>
    <td colspan="11">&nbsp;</td>
	</tr>
	<tr>
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">No</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Date</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Voucher Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Activity<br>Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Activity<br>Name</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Target<br>Type</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Target<br>Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Reference</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Remarks</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Amount</td>
	</tr>
	<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1
    v_Amount = 0%>
	<tr>
	<td align="left" style="<%=cmdstyleLeft%>" colspan="9"><font face="Verdana" size="1">Opening Balance.........</td>
	<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nOpening,2)%></td>
	
	<%
	v_Amount=cdbl(nOpening)
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Datecreated")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Vouchercode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jobcode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Fcname")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Locationtype")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Locationcode")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Reference")%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Remarks")%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Amount"),2)%></td>
		</tr>
	  <%
	  v_Amount=v_Amount+trim(cdbl(rsCat("Amount")))
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr>
	<td align="left" style="<%=cmdstyleLeft%>" colspan="9"> <font face="Verdana" size="1">Ending Balance.........</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Amount,2)%></td>
	</tr>
</table>
<!--#include file="../Bottom_Report.asp"-->