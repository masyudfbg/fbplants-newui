<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>Recurring - iPlaS</title>
</head>
<%

accountingyear = request("txtyear")
documentno = request("documentno")
txtusesettlement = request("txtusesettlement")
txtstatus = request("txtstatus")

txtSQL = " SELECT re.fcba" & _
         "        , re.documentno" & _
         "        , re.documentname" & _
         "        , re.accountingyear" & _
         "        , re.startingdate" & _
         "        , re.finishdate" & _
         "        , re.numberofrun" & _
         "        , re.currency" & _
         "        , re.amount" & _
         "        , re.isusesettlement" & _
         "        , re.accrualjob" & _
         "        , re.status" & _
         "        , red.debetcredit" & _
         "        , red.jobcode" & _
         "        , red.targettype" & _
         "        , red.amount amountdtl" & _
         " FROM recurring re" & _
         " INNER JOIN recurringdtl red ON (re.documentno = red.documentno AND re.fcba = red.fcba)" & _
         " WHERE red.fcba = '" & session("fcBA") & "'" & _
		 "       AND re.accountingyear = '" & accountingyear & "'"
if "" & documentno <> "" then
    txtSQL = txtSQL & " AND red.documentno = '" & documentno & "'"
end if
if "" & txtusesettlement <> "" then
    txtSQL = txtSQL & " AND re.isusesettlement = '" & txtusesettlement & "'"
end if
if "" & txtstatus <> "" then
    txtSQL = txtSQL & " AND re.status = '" & txtstatus & "'"
end if
txtSQL = txtSQL & " ORDER BY red.documentno, red.jobcode, red.targettype"
'response.Write(txtsql)


Set rsCat = connStr.Execute(txtSQL)
arr_data_eof = true
if Not rsCat.eof then
	arr_data = rsCat.getrows
	arr_data_eof = false
end if
rsCat.close()
set rsCat = nothing

Function f_rowspan(data_array__, where__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if Cstr(arr_1(1, indx)) = Cstr(where__) then
			result_ = result_ +1
		end if
	next
	f_rowspan = result_
End Function
%>

	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>Recurring Master</h2>
				</div>
			</div>
			<div style="margin-bottom:10px;"></div>
			<table border="0" width="100%" id="product-table" class="headertable">
				<tr>
					<th>Document No</th>
					<th>Document Name</th>
					<th>Starting Date</th>
					<th>Finish Date</th>
					<th>Number Of Run</th>
					<th>Currency</th>
					<th>Amount</th>
					<th>Use Settlement</th>
					<th>Accrual Job</th>
					<th>Status</th>
					
					<th>Debet/Credit</th>
					<th>Jobcode</th>
					<th>Target Type</th>
					<th>Amount</th>
				</tr>
<%
				if not arr_data_eof then
					for xi=0 to ubound(arr_data,2)
						currdocno = arr_data(1, xi)
%>
						<tr>
						<%if currdocno <> prevdocno then%>
							<%p_span = f_rowspan(arr_data, arr_data(1, xi))%>
							<td align="left" valign="top" rowspan="<%=p_span%>"><%=currdocno%></td>
							<td align="left" valign="top" rowspan="<%=p_span%>"><%=arr_data(2, xi)%></td>
							<td align="left" valign="top" rowspan="<%=p_span%>"><%=arr_data(4, xi)%></td>
							<td align="left" valign="top" rowspan="<%=p_span%>"><%=arr_data(5, xi)%></td>
							<td align="center" valign="top" rowspan="<%=p_span%>"><%=arr_data(6, xi)%></td>
							<td align="center" valign="top" rowspan="<%=p_span%>"><%=arr_data(7, xi)%></td>
							<td align="right" valign="top" rowspan="<%=p_span%>"><%=formatnumber(cdbl(to_nol(arr_data(8, xi))), 2)%></td>
							<td align="center" valign="top" rowspan="<%=p_span%>"><%=arr_data(9, xi)%></td>
							<td align="left" valign="top" rowspan="<%=p_span%>"><%=arr_data(10, xi)%></td>
							<td align="left" valign="top" rowspan="<%=p_span%>"><%=arr_data(11, xi)%></td>
						<%end if%>
							<td align="center"><%=arr_data(12, xi)%></td>
							<td align="center"><%=arr_data(13, xi)%></td>
							<td align="center"><%=arr_data(14, xi)%></td>
							<td align="right"><%=formatnumber(cdbl(to_nol(arr_data(15, xi))), 2)%></td>
						</tr>
<%
						prevdocno = currdocno
					next
				end if
%>
			</table>
			<!--#include file="Bottom_Report.asp"-->
		</div>
	</body>
</html>