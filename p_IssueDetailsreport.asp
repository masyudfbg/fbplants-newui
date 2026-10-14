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

v_txtSIVDATE_From=Trim(REQUEST("SIVDATE_From"))
v_txtSIVDATE_To=Trim(REQUEST("SIVDATE_To"))
v_txtSIVCODE=Trim(REQUEST("SIVCODE"))
v_txtSTOCKCODE=Trim(REQUEST("STOCKCODE"))
v_txtJOBCODE=Trim(REQUEST("JOBCODE"))
v_txtLOCATIONTYPE=Trim(REQUEST("LOCATIONTYPE"))
v_txtLOCATIONCODE=Trim(REQUEST("LOCATIONCODE"))

txtSQL = "SELECT issue.sivdate,iss.sivcode, iss.stockcode, stock.fcname, iss.locationcode," & _
		 "iss.locationtype, nvl(sum(iss.quantity),0) as quantity, nvl(sum(iss.price),0) as price, nvl(sum(iss.amount),0) as amount, job.fcname as jobname, iss.reference as reference, iss.remarks as remarks" & _
		 " FROM issuedetails iss, issue issue, stock stock, job job" & _
		 " WHERE (((iss.sivcode = issue.sivcode) and (iss.fcBA = issue.fcBA))" & _
		 " AND ((iss.stockcode = stock.fccode) and (iss.fcBA= stock.fcBa)) " & _
		 " AND ((iss.jobcode = job.fccode) and (iss.fcBA = job.fcBA)) )" & _
		 " AND iss.fcba='" & session("fcBA") & "'"

if  len(trim(v_txtSIVDATE_From)) > 0 then
    txtSQL = txtSQL & " and issue.sivdate >=to_Date('" & v_txtSIVDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtSIVDATE_From,1)
end if
if  len(trim(v_txtSIVDATE_To)) > 0 then
    txtSQL =txtSQL &  " and issue.sivdate <=to_Date('" & v_txtSIVDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  To " & FormatDateTime(v_txtSIVDATE_To,1)
end if
if  len(trim(v_txtSIVCODE)) > 0 then
    txtSQL = txtSQL & " and iss.sivcode='" & v_txtSIVCODE & "'"
	txtSelect = txtSelect & "  SIV Code " & v_txtSIVCODE
end if
if  len(trim(v_txtSTOCKCODE)) > 0 then
    txtSQL =txtSQL &  " and iss.stockcode='" & v_txtSTOCKCODE & "'"
	txtSelect = txtSelect & "  Stock Code " & v_txtSTOCKCODE
end if 
if  len(trim(v_txtJOBCODE)) > 0 then
    txtSQL =txtSQL &  " and iss.jobcode='" & v_txtJOBCODE & "'"
	txtSelect = txtSelect & "  Job Code " & v_txtJOBCODE
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and iss.locationtype='" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & "  Target Type " & v_txtLOCATIONTYPE
end if 
if  len(trim(v_txtLOCATIONCODE)) > 0 then
    txtSQL =txtSQL &  " and iss.locationcode='" & v_txtLOCATIONCODE & "'"
	txtSelect = txtSelect & "  Target Code " & v_txtLOCATIONCODE
end if 

txtsql=(txtsql) & " group by issue.sivdate,iss.sivcode, iss.stockcode,stock.fcname,iss.locationcode,iss.locationtype,job.fcname,iss.reference, iss.remarks"
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
    <td rowspan="5" width="100"><img src="IMAGES/latar.jpg" width="210" height="80"></td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">STOCK ISSUE</FONT></td></tr>
 </table>
<BR><BR>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
	<tr>
    <td colspan="11" style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" ><font face="Verdana" size="1"><%=txtSelect%></td>
	</tr>
	<tr>
    <td colspan="10">&nbsp;</td>
	</tr>
	<tr>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	</tr>
	<tr>
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">No</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Date</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Issue<br>Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"><font face="Verdana" size="1">Stock<br>Code</td>	
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Stock<br>Desc.</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Job<br>Description</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Target<br>Type</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Target<br>Code</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Quantity</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Price</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Amount</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Reference</td>
		<td align="center" style="<%=cmdstyleTOP%>"> <font face="Verdana" size="1">Remarks</td>
	</tr>
	<%
	Dim v_Quantity,v_price,v_amount
	v_Quantity=0
	v_price=0
	v_amount=0
	While Not rsCat.EOF
	    
	    if cdbl(rscat("quantity")) <>0 and not isnull(rscat("quantity")) then
	       v_price=cdbl(rscat("amount")) / cdbl(rscat("quantity"))
		else
		   v_price=0   
		end if   
	    %>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sivdate")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sivcode")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("stockcode")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobname")%>&nbsp;</td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationtype")%>&nbsp;</td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationcode")%>&nbsp;</td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("quantity"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_price,2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),2)%></td>
		<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("reference")%></td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("remarks")%></td>
		</tr>
	  <%
	  'v_Quantity=v_Quantity+trim(cdbl(rsCat("quantity")))
	  v_Quantity  =v_Quantity+cdbl(rsCat("quantity"))
	  v_amount =v_amount+cdbl(rsCat("amount"))
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
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="8"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Quantity,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1">&nbsp;</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_amount,2)%></td>
	</tr>
</table>
<!--#include file="Bottom_Report.asp"-->