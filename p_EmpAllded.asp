<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<title>ePlaS Reporting - Employee Allowance and Deduction</title>
</head>
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
v_txtGRADE=TRIM(REQUEST("Grade"))
v_cField=trim(REQUEST("cField"))
v_postingflg = trim(REQUEST("postingflg"))
' v_cField1=trim(REQUEST("cField1"))

' if  len(trim(v_cField)) = 0 and len(trim(v_cField1)) = 0 then
	' txtSQL = "SELECT emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
		 ' " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
		 ' " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
		 ' " and emp.fcba='" & session("fcBA") & "'"
' else
	' if  len(trim(v_cField)) = 0 then
	' txtSQL = "SELECT " & v_cField1 & _
		 ' " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
		 ' " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
		 ' " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
		 ' " and emp.fcba='" & session("fcBA") & "'"
	' else
		' if instr(v_cField1,"employee.fcname as empname") > 0 then	
			' txtSQL = "SELECT " & v_cField & "," & v_cField1 & _
				 ' " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
				 ' " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
				 ' " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
				 ' " and emp.fcba='" & session("fcBA") & "'"
		' else
			' txtSQL = "SELECT " & v_cField & _
				 ' " ,emp.allowdedcode,nvl(sum(emp.amount),0) as amount FROM empallded emp, employee employee, allowded allowded" & _
				 ' " WHERE ( ((emp.empcode = employee.fccode) and (emp.fcBA = employee.fcBA)) " & _
				 ' " AND ((emp.allowdedcode = allowded.FCcode) and (emp.fcBA = allowded.fcBA)) )" & _
				 ' " and emp.fcba='" & session("fcBA") & "'"
		' End if		
	' end if
' end if
'' response.Write  "<font color=#FF0000 pointsize=9 face=Arial, Helvetica, sans-serif>" & txtSQL & "</font>"
' if  len(trim(v_txtADHOCDATE_From)) > 0 then
    ' txtSQL = txtSQL & " and emp.adhocdate>=to_Date('" & v_txtADHOCDATE_From & "','MM/DD/YYYY')"
	' txtSelect = " From " & FormatDateTime(v_txtADHOCDATE_From,1)
' end if
' if  len(trim(v_txtADHOCDATE_To)) > 0 then
    ' txtSQL =txtSQL &  " and emp.adhocdate <=to_Date('" & v_txtADHOCDATE_To & "','MM/DD/YYYY')"
	' txtSelect = txtSelect & "  to " & FormatDateTime(v_txtADHOCDATE_To,1)
' end if
' if  len(trim(v_txtGANGCODE)) > 0 then
    ' txtSQL = txtSQL & " and employee.gangcode='" & v_txtGANGCODE & "'"
	' txtSelect = txtSelect & "  Gang Code " & (v_txtGANGCODE)
' end if
' if  len(trim(v_txtFCCODE)) > 0 then
    ' txtSQL =txtSQL &  " and emp.empcode='" & v_txtFCCODE & "'"
	' txtSelect = txtSelect & "  Employee Code " & (v_txtFCCODE)
' end if 
' if  len(trim(v_txtALLOWDEDCODE)) > 0 then
    ' txtSQL =txtSQL &  " and emp.ALLOWDEDCODE='" & v_txtALLOWDEDCODE & "'"
	' txtSelect = txtSelect & "  Allowded Code " & (v_txtALLOWDEDCODE)
' end if 
' if  len(trim(v_cField)) = 0 and len(trim(v_cField1)) = 0 then
		' txtSQL = (txtSQL) & " group by emp.allowdedcode"
	' Else
		' if instr(v_cField1,"employee.fcname as empname") > 0 then
			' if  len(trim(v_cField)) = 0 then
			   ' txtsql=(txtsql) & " group by emp.allowdedcode,employee.fcname"
			' else
			   ' txtsql=(txtsql) & " group by " & v_cField & ",emp.allowdedcode, employee.fcname"
			' end if
		' Else
		   ' txtsql=(txtsql) & " group by " & v_cField & ",emp.allowdedcode"
		' End if	
	' End if
	
If  Len(Trim(v_cField)) = 0 Then
	txtsql	=	txtsql &	" select fcba" & _
							"        , allowdedcode" & _
							"        , allowdedname" & _
							"        , nvl(sum(amount), 0) amount" & _
							" from (" & _
							"     select emp.fcba" & _
							"            , emp.adhocdate" & _
							"            , emp.empcode employeecode" & _
							"            , em.fcname employeename" & _
							"            , em.gangcode" & _
							"            , emp.allowdedcode" & _
							"            , ad.fcname allowdedname" & _
							"            , nvl(emp.amount, 0) amount" & _
							"     from empallded emp" & _
							"          , employee em" & _
							"          , allowded ad " & _
							"     where emp.empcode = em.fccode" & _
							"           and emp.fcba = em.fcba" & _
							"           and emp.allowdedcode = ad.fccode and em.fcba = ad.fcba " & _
							" 			and (em.dateterminate is NULL " & _ 
							"                or to_char(em.dateterminate,'MM/DD/YYYY') <= to_char(to_Date('" & v_txtADHOCDATE_To & "', 'MM/DD/YYYY'),'MM/DD/YYYY')) " & _
							"           and emp.empcode not in (SELECT fccode FROM view_payrolldetail where fcba = '" & session("fcBA") & "'  AND yearmonth= to_char(to_Date('" & v_txtADHOCDATE_From & "', 'MM/DD/YYYY'),'YYYYMM') GROUP BY fccode HAVING SUM(tot_salary) <= 0)" & _
							" 			and em.fcba = '" & session("fcBA") & "'" & _ 
							" 			and adhocdate >= to_Date('" & v_txtADHOCDATE_From & "', 'MM/DD/YYYY')" & _
							" 			and adhocdate <= to_Date('" & v_txtADHOCDATE_To & "', 'MM/DD/YYYY')" & _
							" union all "  & _
							" 		select d.FCBA,h.bkmdate adhocdate,p.EMPLOYEECODE,e.fcname employeename,e.GANGCODE, p.ALLOWDEDCODE,ad.fcname allowdedname,nvl(p.amount, 0) amount " & _
							"		from BKMPROSES p " & _
							"		inner join bkmhdr h on(p.bkmno=h.bkmno and p.fcba=h.fcba) " & _
							"		left join employee e on (e.fccode=p.EMPLOYEECODE and e.fcba=p.fcba) " & _
							"		left join BKMUPAHDTL d on (d.bkmno=p.bkmno and d.EMPLOYEECODE=p.EMPLOYEECODE and d.fcba=p.fcba and p.linenokey=d.linenokey) " & _
							"		left join allowded ad on (p.ALLOWDEDCODE = ad.fccode and p.fcba = ad.fcba ) " & _
							"		where h.bkmdate >= to_Date('" & v_txtADHOCDATE_From & "', 'MM/DD/YYYY') " & _
							"		and h.bkmdate <= to_Date('" & v_txtADHOCDATE_To & "', 'MM/DD/YYYY') and p.FCBA='" & session("fcBA") & "'   " & _
							"       and h.postingflg like '" & v_postingflg & "%'" & _ 
							" )" & _
							" where fcba = '" & session("fcBA") & "' AND allowdedcode not in('A17','A60')" 
Else
	txtsql	=	txtsql &	" select fcba" & _
							"        , allowdedcode" & _
							"        , " & v_cField & "" & _
							"        , allowdedname" & _
							"        , nvl(sum(amount), 0) amount" & _
							" from (" & _
							"     select emp.fcba" & _
							"            , emp.adhocdate" & _
							"            , emp.empcode employeecode" & _
							"            , em.fcname employeename" & _
							"            , em.presentgrade" & _
							"            , em.gangcode" & _
							"            , emp.allowdedcode" & _
							"            , ad.fcname allowdedname" & _
							"            , nvl(emp.amount, 0) amount" & _
							"     from empallded emp" & _
							"          , employee em" & _
							"          , allowded ad " & _
							"     where emp.empcode = em.fccode" & _
							"           and emp.fcba = em.fcba" & _
							"           and emp.allowdedcode = ad.fccode and em.fcba = ad.fcba " & _
							" 			and (em.dateterminate is NULL " & _ 
							"                or to_char(em.dateterminate,'MM/DD/YYYY') <= to_char(to_Date('" & v_txtADHOCDATE_From & "', 'MM/DD/YYYY'),'MM/DD/YYYY')) " & _
							"           and emp.empcode not in (SELECT fccode FROM view_payrolldetail where fcba = '" & session("fcBA") & "'  AND yearmonth= to_char(to_Date('" & v_txtADHOCDATE_From & "', 'MM/DD/YYYY'),'YYYYMM') GROUP BY fccode HAVING sum(tot_salary) <= 0)" & _
							" 			and em.fcba = '" & session("fcBA") & "'" & _ 
							" 			and adhocdate >= to_Date('" & v_txtADHOCDATE_From & "', 'MM/DD/YYYY')" & _
							" 			and adhocdate <= to_Date('" & v_txtADHOCDATE_To & "', 'MM/DD/YYYY')" & _
							"	 union all "  & _
							" 		select d.FCBA,h.bkmdate adhocdate,p.EMPLOYEECODE,e.fcname employeename,e.PRESENTGRADE,e.GANGCODE, p.ALLOWDEDCODE,ad.fcname allowdedname,nvl(p.amount, 0) amount " & _
							"		from BKMPROSES p " & _
							"		inner join bkmhdr h on(p.bkmno=h.bkmno and p.fcba=h.fcba) " & _
							"		left join employee e on (e.fccode=p.EMPLOYEECODE and e.fcba=p.fcba) " & _
							"		left join BKMUPAHDTL d on (d.bkmno=p.bkmno and d.EMPLOYEECODE=p.EMPLOYEECODE and d.fcba=p.fcba and p.linenokey=d.linenokey) " & _
							"		left join allowded ad on (p.ALLOWDEDCODE = ad.fccode and p.fcba = ad.fcba ) " & _
							"		where h.bkmdate >= to_Date('" & v_txtADHOCDATE_From & "', 'MM/DD/YYYY') " & _
							"		and h.bkmdate <= to_Date('" & v_txtADHOCDATE_To & "', 'MM/DD/YYYY') and p.FCBA='" & session("fcBA") & "'   " & _
							"       and h.postingflg like '" & v_postingflg & "%'" & _ 
							" )" & _
							" where fcba = '" & session("fcBA") & "' AND allowdedcode not in('A17','A60')" 
							
End If
if  len(trim(v_txtADHOCDATE_From)) > 0 then
    txtSQL = txtSQL & " and adhocdate >= to_Date('" & v_txtADHOCDATE_From & "', 'MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtADHOCDATE_From,1)
end if
if  len(trim(v_txtADHOCDATE_To)) > 0 then
    txtSQL = txtSQL &  " and adhocdate <= to_Date('" & v_txtADHOCDATE_To & "', 'MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtADHOCDATE_To,1)
end if
if  len(trim(v_txtGANGCODE)) > 0 then
    txtSQL = txtSQL & " and gangcode = '" & v_txtGANGCODE & "'"
	txtSelect = txtSelect & "  Gang Code " & (v_txtGANGCODE)
end if
if  len(trim(v_txtFCCODE)) > 0 then
    txtSQL = txtSQL &  " and employeecode = '" & v_txtFCCODE & "'"
	txtSelect = txtSelect & "  Employee Code " & (v_txtFCCODE)
end if
if  len(trim(v_txtGRADE)) > 0 then
    txtSQL = txtSQL &  " and presentgrade = '" & v_txtGRADE & "'"
	txtSelect = txtSelect & "  Employee Grade " & (v_txtGRADE)
end if
if  len(trim(v_txtALLOWDEDCODE)) > 0 then
    txtSQL = txtSQL &  " and allowdedcode = '" & v_txtALLOWDEDCODE & "'"
	txtSelect = txtSelect & "  Allowded Code " & (v_txtALLOWDEDCODE)
end if

' For Group By
If  Len(Trim(v_cField)) = 0 Then
	txtsql	=	txtsql &	" group by fcba" & _
							"          , allowdedcode" & _
							"          , allowdedname"
Else
	txtsql	=	txtsql &	" group by fcba" & _
							"          , allowdedcode" & _
							"          , allowdedname" & _
							"          , " & v_cField & ""
End If

' For Order By
txtsql =	txtsql & " order by fcba"
if instr(v_cField,"employeecode") > 0 then txtsql = txtsql & " , employeecode" End If
txtsql =	txtsql & " , allowdedcode"
if instr(v_cField,"adhocdate") > 0 then txtsql = txtsql & " , adhocdate" End If
	
' Response.Write txtsql

Dim Kolom : Kolom = 2
if instr(v_cField,"adhocdate") > 0 then
	 Kolom = Kolom + 1
End if
if instr(v_cField,"employeecode") > 0 then
	 Kolom = Kolom + 1
End if
if instr(v_cField,"employeename") > 0 then
	 Kolom = Kolom + 1
End if
if instr(v_cField,"presentgrade") > 0 then
	 Kolom = Kolom + 1
End if

if instr(v_cField,"gangcode") > 0 then
	 Kolom = Kolom + 1
End if

%>
  <body> 
		<div id="wrapper">
        <!--#include file="print_excel.asp"-->
        <table border="0" width="100%" id="product-table" class="xthf-site">
         <thead>
         <tr>
         <th colspan="<%=1+Kolom%>" style=" background-color:#FFFFFF; color:#000000; border:none">
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered" style="width:50%; margin:auto;">
					<h2>PAYSLIP DETAILS</h2>
				</div>
			</div>
			</th>
			</tr>
						<!--<div class="printhide">-->
						<tr>
							<td colspan="8" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></td>
						</tr>
						<!--</div>-->
 
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
			
		<%if instr(v_cField,"adhocdate") > 0 then%>
			<th>Date</th>
		<%end if%>
        <%if instr(v_cField,"employeecode") > 0 then%>
			<th>Employee Code</th>
		<%end if%>
        <%if instr(v_cField,"employeename") > 0 then%>
			<th>Employee Name</th>
		<%end if%>
		<%if instr(v_cField,"presentgrade") > 0 then%>
			<th>Employee Grade</th>
		<%end if%>
        <%if instr(v_cField,"gangcode") > 0 then%>
			<th>Gang Desc</th>
		<%end if%>
			<th>Allowance Code</th>
			<th>Allowance Desc</th>
			<th>Amount</th>
	</tr>
    </thead>
    <tbody>
	<%
	Dim v_amount
	v_amount=0
	While Not rsCat.EOF%>
		<tr>
		 
		<%if instr(v_cField,"adhocdate") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("adhocdate")%></td>
		<%end if%>
        <%if instr(v_cField,"employeecode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("employeecode")%> </td>
		<%end if%>
        <%if instr(v_cField,"employeename") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("employeename")%> </td>
		<%end if%>
		<%if instr(v_cField,"presentgrade") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("presentgrade")%> </td>
		<%end if%>
        <%if instr(v_cField,"gangcode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("gangcode")%> </td>
		<%end if%>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("allowdedcode")%> </td>
		<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("allowdedname")%> </td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("amount"),2)%></td>
		</tr>
	  <%
	  'if trim(mid(rsCat("allowdedcode"),1,1))="D" Then
	  '   nAmountAllowance = cdbl(rsCat("amount"))*-1
	  'else
	  '   nAmountAllowance = cdbl(rsCat("amount"))
	  'end if
	  v_amount = v_amount + cdbl(rsCat("amount"))
	  'response.write "v_amount : " & v_amount   vv
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
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="<%=KOLOM%>"> <font face="Verdana" size="1"><strong>T o t a l</strong></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(v_amount,2)%></strong></td>
	</tr>
</tbody>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>