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

v_txtADHOCDATE_From=TRIM(REQUEST("FDDATE_From"))
v_txtADHOCDATE_To=TRIM(REQUEST("FDDATE_to"))
v_txtGANGCODE=TRIM(REQUEST("GANGCODE"))
v_txtALLOWDEDCODE=TRIM(REQUEST("Allowded"))
v_txtFCCODE=TRIM(REQUEST("EMPLOYEECODE"))
v_cField=trim(REQUEST("cField"))
v_cField1=trim(REQUEST("cField1"))

if  len(trim(v_cField)) = 0 and len(trim(v_cField1)) = 0 then
	txtSQL = "SELECT emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
		 " WHERE (emp.empcode = employee.fccode and emp.fcBA = employee.fcBA) " & _
		 " AND (emp.allowdedcode = allowded.FCcode and emp.fcBA = allowded.fcBA) " & _
		 " and emp.fcba='" & session("fcBA") & "'"
else
	if  len(trim(v_cField)) = 0 then
	txtSQL = "SELECT " & v_cField1 & _
		 " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
		 " WHERE (emp.empcode = employee.fccode and emp.fcBA = employee.fcBA) " & _
		 " AND (emp.allowdedcode = allowded.FCcode and emp.fcBA = allowded.fcBA) " & _
		 " and emp.fcba='" & session("fcBA") & "'"
	else
		if instr(v_cField1,"employee.fcname as empname") > 0 then	
			txtSQL = "SELECT " & v_cField & "," & v_cField1 & _
				 " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
				 " WHERE  (emp.empcode = employee.fccode and emp.fcBA = employee.fcBA) " & _
				 " AND (emp.allowdedcode = allowded.FCcode and emp.fcBA = allowded.fcBA) " & _
				 " and emp.fcba='" & session("fcBA") & "'"
		else
			txtSQL = "SELECT " & v_cField & _
				 " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
				 " WHERE (emp.empcode = employee.fccode and emp.fcBA = employee.fcBA) " & _
				 " AND (emp.allowdedcode = allowded.FCcode and emp.fcBA = allowded.fcBA) " & _
				 " and emp.fcba='" & session("fcBA") & "'"
		End if		
	end if
end if
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
	txtSelect = txtSelect & "  Gang : " & cmdFindTitle_BA("GANG","FCCODE",v_txtGANGCODE)
end if
if  len(trim(v_txtFCCODE)) > 0 then
    txtSQL =txtSQL &  " and emp.empcode='" & v_txtFCCODE & "'"
	txtSelect = txtSelect & "  Employee : " & cmdFindTitle_BA("Employee","fcCode",v_txtFCCODE)
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
 <table width="700">
  <tr> 
    <td rowspan="5" width="100"><img src="../IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">PAYSLIP</FONT></td></tr>
 </table>
<BR><BR>
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
	 Kolom=Kolom + 2
End if
if instr(v_cField,"emp.allowdedcode") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"allowded.fcname") > 0 then
	 Kolom=Kolom + 1
End if
	Kolom=KOLOM + 2%>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	<tr>
    <td colspan="12" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" ><font face="Verdana" size="1"><%=txtSelect%></td>
	</tr>
	<tr>
    <td>&nbsp;</td>
	</tr>
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	'response.Write(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
	<tr>
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">No</td>
		<%if instr(v_cField,"emp.adhocdate") > 0 then%>
			<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Date</td>
		<%end if%>
        <%if instr(v_cField,"emp.empcode") > 0 then%>
			<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Employee<br>Code</td>
		<%end if%>
        <%if instr(v_cField1,"employee.fcname as empname") > 0 then%>
			<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Employee<br>Name</td>	
		<%end if%>
        <%if instr(v_cField,"employee.gangcode") > 0 then%>
			<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Gang<br>Code.</td>	
			<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Gang<br>Desc.</td>	
		<%end if%>
        <%if instr(v_cField,"allowded.fcname") > 0 then%>
			<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Allowance</td>
		<%end if%>
			<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Allowance<br>Code</td>
			<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Amount</td>
	</tr>
	<%
	Dim v_amount
	v_amount=0
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<%if instr(v_cField,"emp.adhocdate") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("adhocdate")%></td>
		<%end if%>
        <%if instr(v_cField,"emp.empcode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("empcode")%>&nbsp;</td>
		<%end if%>
        <%if instr(v_cField1,"employee.fcname as empname") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("empname")%>&nbsp;</td>
		<%end if%>
        <%if instr(v_cField,"employee.gangcode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("gangcode")%>&nbsp;</td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=cmdFindTitle_BA("GANG","FCCODE",rscat("gangcode"))%>&nbsp;</td>
		<%end if%>
        <%if instr(v_cField,"allowded.fcname") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%>&nbsp;</td>
		<%end if%>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("allowdedcode")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),2)%></td>
		</tr>
	  <%
	  if trim(mid(rsCat("allowdedcode"),1,1))="D" Then
	     nAmountAllowance = cdbl(rsCat("amount"))*-1
	  else
	     nAmountAllowance = cdbl(rsCat("amount"))
	  end if
	  v_amount=v_amount+nAmountAllowance
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
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="<%=KOLOM%>"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amount,2)%></td>
	</tr>

</table>
<!--#include file="../Bottom_Report.asp"-->