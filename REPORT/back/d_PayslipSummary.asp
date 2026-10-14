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

Dim rsGroup

v_txtDATECREATED_From= trim(request("DATECREATED_From"))
v_txtDATECREATED_To=trim(request("DATECREATED_To"))
v_txtEmpCODE=trim(request("EMPCODE"))
v_txtGangCODE=  trim(request("GANGCODE"))

txtSQL = "select p.Empcode,e.fcName From Employee E,EmpAllded P Where " & _
         "(e.fccode= p.empcode and e.fcba = p.fcba)  and p.fcba='" & session("fcBA") & "'"
			 
if  len(trim(v_txtDATECREATED_From)) > 0 then
	txtSQL = txtSQL & " and p.adhocdate >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtDATECREATED_From,1)
end if
if  len(trim(v_txtDATECREATED_To)) > 0 then
	txtSQL =txtSQL &  " and p.adhocdate <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtDATECREATED_To,1)
end if
if  len(trim(v_txtEmpcode)) > 0 then
	txtSQL = txtSQL & " and p.EmpCode='" & v_txtEmpcode & "'"
	txtSelect = txtSelect & " Emp Code " & v_txtEmpcode
end if
if  len(trim(v_txtGangCode)) > 0 then
	txtSQL =txtSQL &  " and p.Empcode in (Select distinct(fccode) From Employee Where GangCode='" & v_txtGangCode & "' and fcba='" & session("fcBA") & "')"
	txtSelect = txtSelect & " Gang Code " & v_txtGangCode
end if 

txtsql=(txtsql) & " group by p.Empcode,e.fcName"

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
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
<tr><td width="150">REPORT</td><td  width="300"><b><FONT color="#000099">: PAYSLIP SUMMARY</FONT></td></tr>
</table>
<br/>
<table Cellpadding="2" cellspacing="0">
<tr>
<td colspan="14" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" ><font face="Verdana" size="1"><%=txtSelect%></td>
</tr><TR></tr>
<tr>
<td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">No</td>
<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Employee<br>ID</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Employee<br>Name</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Gaji</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Lembur</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Premi</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Tunjangan</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Borongan</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Jamsostek</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Potongan</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">PPH Pasal 21</td>
<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Total</td>
<td align="center" style="<%=cmdstyleTOP%>" colspan="2"> <font face="Verdana" size="1">Paraf</td>
</tr>
	<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo	
	Set rsEmployee = connStr.Execute(txtSQL)
	if Not rsEmployee.eof then
	   nCountofField=rsEmployee.fields.count-1
	end if
	NO=1
    v_Amount = 0
	nGrantTotal = 0
	While   Not rsEmployee.EOF
			txtSQL = 	" select" & _
						" fcba," & _
						" allded_group," & _
						" position" & _
						" from" & _
						" (" & _
						" SELECT" & _
						" ald.fcba fcba," & _
						" ald.allded_group allded_group," & _
						" par.param_1 position" & _
						" from allowded ald, parameter par" & _
						" where" & _
						" (ald.fcba = par.fcba and ald.allded_group = par.fccode)" & _
						" group by" & _
						" ald.fcba," & _
						" ald.allded_group," & _
						" par.param_1" & _
						" order by" & _
						" par.param_1)" & _
						" where fcba='" & session("fcBA") & "'"
			
			Set rsGroup = connStr.Execute(txtSQL)
			if Not rsGroup.eof then
				%>
				<tr>
				<td align="left" style="<%=cmdstyleLeft%>" height="30"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
				<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsEmployee("Empcode")%>&nbsp;</td>
				<td align="left" style="<%=cmdstyle%>" width="200"><font face="Verdana" size="1"><%=rsEmployee("fcName")%>&nbsp;</td>
				<%
				ntotal=0
				While   Not rsGroup.EOF
						txtPayslip = "select nvl(sum(Amount),0) as Amount From EmpAllded P Where p.fcba='" & session("fcBA") & "'" & _
								 " and Allowdedcode in (select fccode from Allowded where allded_group='" & rsGroup("allded_group") & "' and fcba='" & session("fcba") & "')" & _
								 " and p.empcode = '" & rsEmployee("Empcode") & "'" 
						nPayslip=0			 
						if  len(trim(v_txtDATECREATED_From)) > 0 then
							txtPayslip = txtPayslip & " and p.adhocdate >=to_Date('" & v_txtDATECREATED_From & "','MM/DD/YYYY')"
						end if
						if  len(trim(v_txtDATECREATED_To)) > 0 then
							txtPayslip =txtPayslip &  " and p.adhocdate <=to_Date('" & v_txtDATECREATED_To & "','MM/DD/YYYY')"
						end if
						
						Set rsPayslip = connStr.Execute(txtPayslip)
						response.Write allded_group
						if Not rsPayslip.eof then%>
						   <td style="<%=cmdstyle%>" align="right"><font face="Verdana" size="1"><%=FormatNumber(cdbl(rsPayslip("Amount")),0)%>&nbsp;</td>
						   <%
						Else
						   nPayslip=0
						   %>
						   <td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1">0</td>
						   <%   
						end if
						ntotal = ntotal+cdbl(rsPayslip("Amount"))
						rsGroup.MoveNext
				Wend
			    nGrantTotal = nGrantTotal+cdbl(ntotal)
				%>		
			   <td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(nTotal,0)%></td>
			   <%if no  mod 2 = 1 then%>
				   <td align="left" style="<%=cmdstyle%>" width="50" bgcolor="#CCCCCC"><font face="Verdana" size="1">&nbsp;</td>			
				   <td align="left" style="<%=cmdstyle%>" width="50"><font face="Verdana" size="1">&nbsp;</td>			
			   <%Else%>
				   <td align="left" style="<%=cmdstyle%>" width="50"><font face="Verdana" size="1">&nbsp;</td>			
				   <td align="left" style="<%=cmdstyle%>" width="50" bgcolor="#CCCCCC"><font face="Verdana" size="1">&nbsp;</font></td>			
			   <%End if%>
				</tr>		 
				<%
		    End if
			rsEmployee.MoveNext
			no=no+1
	Wend
	rsGroup.close()
	SET rsGroup = Nothing%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="11" height="30"> <font face="Verdana" size="1">T O T A L</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nGrantTotal,0)%>&nbsp;</td>
	<td align="left" style="<%=cmdstyle%>" width="50" bgcolor="#CCCCCC"><font face="Verdana" size="1">&nbsp;</td>
	<td align="left" style="<%=cmdstyle%>" width="50" bgcolor="#CCCCCC"><font face="Verdana" size="1">&nbsp;</td>
	</tr>
</table>
<!--#include file="../Bottom_Report.asp"-->