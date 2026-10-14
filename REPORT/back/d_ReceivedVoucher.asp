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
v_txtVOUCHERCODE=  trim(request("VOUCHERCODE"))
v_txtLOCATIONTYPE= trim(request("LOCATIONTYPE"))
v_txtITEMCODE=trim(request("ITEMCODE"))

txtSQL = "SELECT p.datecreated,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode, nvl(sum(pay.amount),0) as amount, pay.REFERENCE, pay.remarks" & _
		 " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
		 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
		 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" & _
		 " AND pay.fcba='" & session("fcBA") & "' and pay.vouchercode like 'PV%'"
if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL = txtSQL & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL =txtSQL &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtVOUCHERCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.VOUCHERCODE='" & v_txtVOUCHERCODE & "'"
	txtSelect = txtSelect & " Voucher No " & v_txtVOUCHERCODE
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and pay.locationtype='" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & " Location Type " & v_txtLOCATIONTYPE
end if 
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.locationcode='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Location Code " & v_txtITEMCODE
end if 

txtsql=(txtsql) & " group by p.datecreated,p.bankcode,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode,pay.REFERENCE, pay.remarks"
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
	<tr><td width="150">REPORT</td><td  width="300"><b><FONT color="#000099">: RECEIVED VOUCHER</FONT></td></tr>
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
    v_Amount = 0
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
		<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Datecreated")%>&nbsp;</td>
		
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Vouchercode")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Jobcode")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Fcname")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Locationtype")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Locationcode")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Reference")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Remarks")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Amount"),2)%>&nbsp;</td>
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
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="10"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Amount,2)%>&nbsp;</td>
	</tr>
</table>
<!--#include file="../Bottom_Report.asp"-->