<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Payment Voucher</title>
</head>
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
v_txtJOBCODE=Trim(REQUEST("JOBCODE"))
v_txtLOCATIONTYPE= trim(request("LOCATIONTYPE"))
v_txtITEMCODE=trim(request("ITEMCODE"))

txtSQL = "SELECT p.datecreated,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode, nvl(sum(pay.amount),0) as amount, pay.REFERENCE, pay.remarks" & _
		 " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
		 " WHERE ( ((pay.jobcode = job.fccode) AND (pay.FCBA = job.FCBA)) and " & _
		 "((pay.vouchercode = p.vouchercode) AND (pay.FCBA = p.FCBA)) )" & _
		 " AND pay.fcba='" & session("fcBA") & "'"
if  len(trim(v_txtDATECREATED_From)) > 0 then
    txtSQL = txtSQL & " and p.datecreated >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
    txtSQL =txtSQL &  " and p.datecreated <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtBANKCODE)) > 0 then
    txtSQL = txtSQL & " and p.BANKCODE='" & v_txtBANKCODE & "'"
	txtSelect = txtSelect & " Bank Code " & v_txtBANKCODE
end if
if  len(trim(v_txtVOUCHERCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.VOUCHERCODE='" & v_txtVOUCHERCODE & "'"
	txtSelect = txtSelect & " Voucher No " & v_txtVOUCHERCODE
end if 
if  len(trim(v_txtJOBCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.jobcode='" & v_txtJOBCODE & "'"
	txtSelect = txtSelect & "  Job Code " & v_txtJOBCODE
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and pay.locationtype='" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & " Location Type " & v_txtLOCATIONTYPE
end if 
if  len(trim(v_txtITEMCODE)) > 0 then
    txtSQL =txtSQL &  " and pay.locationcode='" & v_txtITEMCODE & "'"
	txtSelect = txtSelect & " Location Code " & v_txtITEMCODE
end if 

txtsql=(txtsql) & " group by p.datecreated,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		 " pay.locationcode,pay.REFERENCE, pay.remarks"
txtsql	=	txtsql & " Order By p.datecreated, pay.vouchercode asc, Amount Desc"
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'response.Write(txtsql)
response.Flush()
%>

	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>Bank Activity</h2>
				</div>
			</div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<div class="printhide">
				<tr>
					<td colspan="9"><%=txtSelect%></td>
				</tr>
				</div>
				<tr>
					 
					<th>Date</th>
					<th>Code</th>
					<th>Activity Code</th>
					<th>Activity Name</th>
					<th>Target Type</th>
					<th>Target Code</th>
					<th>Reference</th>
					<th>Remarks</th>
					<th>Amount</th>
				</tr>
				<%
				Dim v_Opening, v_Debit, v_Credit, v_Saldo	
				Set rsCat = connStr.Execute(txtSQL)
				if Not rsCat.eof then
					nCountofField=rsCat.fields.count-1
				end if
				NO	=	1
				v_Amount	=	0
				While Not rsCat.EOF
				%>
				<tr>
					
					<td align="Center"><%=rscat("Datecreated")%></td>
					<td align="left"><%=rscat("Vouchercode")%></td>
					<td align="left"><%=rscat("Jobcode")%></td>
					<td align="left"><%=rscat("Fcname")%></td>
					<td align="left"><%=rscat("Locationtype")%></td>
					<td align="left"><%=rscat("Locationcode")%></td>
					<td align="left"><%=rscat("Reference")%></td>
					<td align="left"><%=rscat("Remarks")%></td>
					<td align="Right"><%=FormatNumber(rscat("Amount"),2)%></td>
				</tr>
				<%
				v_Amount	=	v_Amount	+	trim(cdbl(rsCat("Amount")))
				'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
				rsCat.MoveNext
				'response.Flush()
				no	=	no	+	1
				Wend
				rsCat.close()
				rsCat	=	null
				%>
				<input  type="hidden" value="<%=txtTable%>" name="txtTable">
				<input  type="hidden" value="<%=txtModule%>" name="txtModule">
				</form>
				<tr bgcolor="#F4F4A0">
					<td align="Center" colspan="8"><strong><i>Total</i></strong></td>
					<td align="Right"><strong><i><%=FormatNumber(v_Amount,2)%></i></strong></td>
				</tr>
			</table>
			<!--#include file="Bottom_Report.asp"-->
		</div>
	</body>
</html>