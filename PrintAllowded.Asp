<%'Option Explicit
%>
<%' On Error GOTO CmdErrorlog 
%>
<%response.buffer=True%>
<%
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"
%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
 <%

    Dim PeriodDate,PeriodMonth,Header_Employee,Header_Allowded
	
 	PeriodDate =request("PeriodDate")
	PeriodMonth =request("PeriodMonth")
	Header_Employee=request("Header_Employee")
	Header_Allowded=request("Header_Allowded")

    Dim txtNoWS,txtDateWS,txtRemarks,q,nCountofField,no,i
    Dim Remarks
    Dim rsWS,intRecordsPerPage,cColor,txtEmpAllDed
	txtEmpAllDed =cmdFindTitle("Allowded","fcCode",request("Header_Allowded"))
    %>

 <table width="541">
  <tr> 
    <td rowspan="6" width="100">
	<% if len(trim(PROFIL("REPORT_IMAGE"))) > 0 then %>
		<img src="IMAGES/<%=PROFIL("REPORT_IMAGE")%>" width="210" height="80">
	<% end if %>
	</td>
  </tr>
	<tr><td width="100">Company</td><td  width="300"><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td width="100">Business Unit</td><td  width="300"><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("FCBA")%></FONT></td></tr>
	<tr><td width="100">Report</td><td  width="300"><b><FONT color="#000099">ALLOWANCE AND DEDUCTION</FONT></td></tr>
 </table>
 	
   <%
   Dim q1,q2,q3,q4,q5,isiWhere
   q1=""
   q2=""
   q3=""
   q4=""
   q5=""
   isiWhere=""
   'Field Numeric Jangan di Sum karena ditampilkan mesti level detail
   q1="select EMPCODE, emp.fcname as empName,ALLOWDEDCODE, allo.fcname as ALLOWDEDDES, ADHOCDATE, AMOUNT as Amount " & _
      " from EmpAllDed EmpAll,Allowded ALLO,Employee Emp where (empall.ALLOWDEDCODE=allo.fccode and empall.fcBA=allo.fcBA) " & _
	  " and (empall.empcode=emp.fccode and empall.fcba=emp.fcba)"
   q = q1
   if trim(PeriodMonth) <> "" then
        q2 = " to_char(ADHOCDATE,'mm/yyyy')=to_char(to_date('" & PeriodMonth & "','" & "MM/DD/YYYY" & "'),'mm/yyyy') "
		q  = q1 + " and " + q2
		IsiWhere ="OK"
   end if
   if trim(PeriodDate) <> "" then
        q3 = " ADHOCDATE=to_date('" & PeriodDate & "','" & "MM/DD/YYYY" & "')"
	    q = q + " and " + q3
   end if
   if trim(Header_Employee) <> "" then
        q4 = " EMPCODE ='" & Header_Employee & "'"
	    q = q + " and " + q4
   end if
   if trim(Header_Allowded) <> "" then
        q5 = " ALLOWDEDCODE='" & Header_Allowded & "'"
	    q = q + " and " + q5
   end if
   
   'response.Write q	
   
   Dim rsEmpAllDed,nQty,nAmount
   Set rsEmpAllDed = Server.CreateObject("ADODB.Recordset")
   rsEmpAllDed.CursorType = 3
   rsEmpAllDed.CursorLocation = 3
   rsEmpAllDed.Open q, connstr%>
   <br><br>
<table Cellpadding="2" bgcolor="#F7F7F7" cellspacing="0">
   <tr >
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="right" colspan="7">
		<font face="Verdana" size="1">Selected by : Period = <%=PeriodDate%> ; Date = <%=PeriodMonth%> ; Employee Code = <%=Header_Employee%> ; Allowance = <%=Header_Allowded%>
		</td>
   <tr><td>&nbsp;</td></tr>
   <tr >
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF" align="right"><font face="Verdana" size="1">No</td>
   <td style="<%=cmdstyleTop%> align="center"><font face="Verdana" size="1">Date</td></td>
   <td style="<%=cmdstyleTop%> align="center"><font face="Verdana" size="1">Employee<br>Code</td>
   <td style="<%=cmdstyleTop%> align="center"><font face="Verdana" size="1">Employee<br>Name</td>
   <td style="<%=cmdstyleTop%> align="center"><font face="Verdana" size="1">Allowance<br>Code</td>
   <td style="<%=cmdstyleTop%> align="center"><font face="Verdana" size="1">Allowance<br>Description</td>
   <td style="<%=cmdstyleTop%> align="center"><font face="Verdana" size="1">Amount</td>
   </tr>
   <%
   if Not rsEmpAllDed.eof then
      nCountofField=rsEmpAllDed.fields.count-1
	  no=1
	  nQty=0
	  nAmount=0
	  While Not rsEmpAllDed.EOF

		  'nQty = nQty+cdbl(rsEmpAllDed("Quantity"))
		  nAmount = nAmount+cdbl(rsEmpAllDed("Amount"))
		  if no  mod 2 = 1 then
             cColor = "#EEEEEE"
		  else
		     cColor = "#E5E5E5"
		  end if%>
		  <tr >
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsEmpAllDed("ADHOCDATE")%>&nbsp;</td>
          <td align="left" style="<%=cmdstyle%>" title="WS.ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsEmpAllDed("EMPCODE")%>&nbsp;</td>
		  <td align="left" style="<%=cmdstyle%>" title="WS.ITEMCODE" valign="top"><font face="Verdana" size="1"><%=rsEmpAllDed("empName")%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsEmpAllDed("ALLOWDEDCODE")%>&nbsp;</td>
	      <td align="left" style="<%=cmdstyle%>" title="WS.OTHERITEMDESC" valign="top"><font face="Verdana" size="1"><%=rsEmpAllDed("ALLOWDEDDES")%>&nbsp;</td>
	      <td align="Right" style="<%=cmdstyle%>" title="WS.Amount" valign="top"><font face="Verdana" size="1"><%=FormatNumber(rsEmpAllDed("AMOUNT"),2)%>&nbsp;</td>
		  </tr>
		  <%
		  rsEmpAllDed.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	<tr>
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="6"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nAmount,2)%>&nbsp;</td>
	</tr>
<%
   else
      response.Write("Data is empty")	  
   end if 
   rsEmpAllDed.Close
   set rsEmpAllDed = Nothing%>
 </TABLE>
<!--#include file="Bottom_Report.asp"-->