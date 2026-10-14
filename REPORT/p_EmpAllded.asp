<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Employee Allowance and Deduction</title>
</head>
<body>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

v_txtADHOCDATE_From=TRIM(REQUEST("FDDATE_From"))
v_txtADHOCDATE_To=TRIM(REQUEST("FDDATE_to"))
v_txtGANGCODE=TRIM(REQUEST("GANGCODE"))
v_txtALLOWDEDCODE=TRIM(REQUEST("Allowded"))
v_txtFCCODE=TRIM(REQUEST("EMPLOYEECODE"))
v_cField=trim(REQUEST("cField"))
v_cField1=trim(REQUEST("cField1"))

if  len(trim(v_cField)) = 0 and len(trim(v_cField1)) = 0 then
	txtSQL = "SELECT emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
		 " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
		 " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
		 " and emp.fcba='" & session("fcBA") & "'"
else
	if  len(trim(v_cField)) = 0 then
	txtSQL = "SELECT " & v_cField1 & _
		 " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
		 " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
		 " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
		 " and emp.fcba='" & session("fcBA") & "'"
	else
		if instr(v_cField1,"employee.fcname as empname") > 0 then	
			txtSQL = "SELECT " & v_cField & "," & v_cField1 & _
				 " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
				 " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
				 " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
				 " and emp.fcba='" & session("fcBA") & "'"
		else
			txtSQL = "SELECT " & v_cField & _
				 " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
				 " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
				 " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
				 " and emp.fcba='" & session("fcBA") & "'"
		End if		
	end if
end if
'response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & txtSQL & "</font>"
if  len(trim(v_txtADHOCDATE_From)) > 0 then
    txtSQL = txtSQL & " and emp.adhocdate>=to_Date('" & v_txtADHOCDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtADHOCDATE_From,1)
end if
if  len(trim(v_txtADHOCDATE_To)) > 0 then
    txtSQL =txtSQL &  " and emp.adhocdate <=to_Date('" & v_txtADHOCDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtADHOCDATE_To,1)
end if
if  len(trim(v_txtGANGCODE)) > 0 then
    txtSQL = txtSQL & " and employee.gangcode='" & v_txtGANGCODE & "'"
	txtSelect = txtSelect & "  Gang Code " & (v_txtGANGCODE)
end if
if  len(trim(v_txtFCCODE)) > 0 then
    txtSQL =txtSQL &  " and emp.empcode='" & v_txtFCCODE & "'"
	txtSelect = txtSelect & "  Employee Code " & (v_txtFCCODE)
end if 
if  len(trim(v_txtALLOWDEDCODE)) > 0 then
    txtSQL =txtSQL &  " and emp.ALLOWDEDCODE='" & v_txtALLOWDEDCODE & "'"
	txtSelect = txtSelect & "  Allowded Code " & (v_txtALLOWDEDCODE)
end if 
if  len(trim(v_cField)) = 0 and len(trim(v_cField1)) = 0 then
		txtSQL = (txtSQL) & " group by emp.allowdedcode"
	Else
		if instr(v_cField1,"employee.fcname as empname") > 0 then
			if  len(trim(v_cField)) = 0 then
			   txtsql=(txtsql) & " group by emp.allowdedcode,employee.fcname"
			else
			   txtsql=(txtsql) & " group by " & v_cField & ",emp.allowdedcode, employee.fcname"
			end if
		Else
		   txtsql=(txtsql) & " group by " & v_cField & ",emp.allowdedcode"
		End if	
	End if

%>
  <body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>PAYSLIP</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="7"><%=txtSelect%></td>
						</tr>
						</div>
  <% Dim Kolom
  Kolom =0
if instr(v_cField,"emp.adhocdate") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"emp.empcode") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField1,"employee.fcname as empname") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"employee.gangcode") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"emp.allowdedcode") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"allowded.fcname") > 0 then
	 Kolom=Kolom + 1
End if
	Kolom=KOLOM + 1%>
 
	<%	
	Set rsCat = connStr.Execute(txtSQL)
        'response.write txtsql
        'response.flush
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
	<tr>
			
		<%if instr(v_cField,"emp.adhocdate") > 0 then%>
			<th>Date</th>
		<%end if%>
        <%if instr(v_cField,"emp.empcode") > 0 then%>
			<th>Employee Code</th>
		<%end if%>
        <%if instr(v_cField1,"employee.fcname as empname") > 0 then%>
			<th>Employee Name</th>
		<%end if%>
        <%if instr(v_cField,"employee.gangcode") > 0 then%>
			<th>Gang Desc.</th>
		<%end if%>
        <%if instr(v_cField,"allowded.fcname") > 0 then%>
			<th>Allowance</th>
		<%end if%>
			<th>Allowance Code</th>
			<th>Amount</th>
	</tr>
	<%
	Dim v_amount
	v_amount=0
	While Not rsCat.EOF%>
		<tr>
		 
		<%if instr(v_cField,"emp.adhocdate") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("adhocdate")%></td>
		<%end if%>
        <%if instr(v_cField,"emp.empcode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("empcode")%> </td>
		<%end if%>
        <%if instr(v_cField1,"employee.fcname as empname") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("empname")%> </td>
		<%end if%>
        <%if instr(v_cField,"employee.gangcode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("gangcode")%> </td>
		<%end if%>
        <%if instr(v_cField,"allowded.fcname") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%> </td>
		<%end if%>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("allowdedcode")%> </td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),2)%></td>
		</tr>
	  <%
	  'if trim(mid(rsCat("allowdedcode"),1,1))="D" Then
	  '   nAmountAllowance = cdbl(rsCat("amount"))*-1
	  'else
	  '   nAmountAllowance = cdbl(rsCat("amount"))
	  'end if
	  v_amount=v_amount+cdbl(rsCat("amount"))
	  'response.write "v_amount : " & v_amount
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr bgcolor="#F4F4A0">
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="<%=KOLOM%>"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amount,2)%></td>
	</tr>

</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>