<%Option Explicit%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>iPS - Data Print</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
</head>
<body>
<table class="entry">
<tr>
 
 <table class="entry">
  <tr> 
    <td rowspan="5"><img src="IMAGES/latar.jpg" width="90" height="100"></td>
  </tr>
	<tr><td width="40">Company</td><td><b><FONT color="#000099"><%=Session("COMPANY")%></FONT></td></tr>
	<tr><td>Business Unit</td><td><b><FONT color="#000099"><%=Session("BA")%> / <%=Session("KodeBA")%></FONT></td></tr>
	<tr><td>Manager</td><td><b><FONT color="#000099"><%=Session("FCMANAGER")%></FONT></td></tr>
	<tr><td>Module</td><td><b><FONT color="#000099"><%=request("txtModule")%>/<%=request("txtTable")%></FONT></td></tr>
 </table>
<tr>
 <td>
  <%' -----------------ATAS ISI%>
	
	<%
	Dim txtTable,txtModule,q,nCountofField,NO,txtSearch,txtField,txtWhere,txtColom,txtOrder
	Dim i
	txtTable =Request("txtTable")
	txtModule=Request("txtModule")
	txtSearch=Request("txtSearch")
	txtField =Request("txtField")

    If len(trim(txtSearch))>0 and len(trim(txtField)) > 0 then
		if instr(txtSearch,"/") <> 0  AND len(txtSearch)=10  then 'Date Type
		   txtWhere = txtField & " like TO_DATE('" & txtSearch & "','DD/MM/YYYY')"
		else
		   if instr(txtSearch,"=") > 0 then 'Data Fix
			  txtWhere = txtField & (txtSearch) 
		   elseif instr(txtSearch,">") > 0 then 'Data Fix
			  txtWhere = txtField & (txtSearch) 
		   elseif instr(txtSearch,"<") > 0 then 'Data Fix
			  txtWhere = txtField & (txtSearch) 
		   elseif instr(txtSearch,"<=") > 0 then 'Data Fix
			  txtWhere = txtField & (txtSearch) 
		   elseif instr(txtSearch,">=") > 0 then 'Data Fix
			  txtWhere = txtField & (txtSearch) 
		   elseif txtField <> "" and txtSearch <> "" then  'Caracter
			  txtWhere = " upper (" & (txtField) & ") like '%" & ucase(txtSearch) & "%'"
		   end if
		end if
		If txtModule="Data Setup" then  'Jika Date Setup Wherenya tidak ada yang lain sudah dibawah di sql
		   txtWhere=" Where " & txtWhere
		Else
		   txtWhere=" and " & txtWhere
		End if
    End if

	IF txtTable="Attendance_GAD" Then
	   if len(txtField)>0 and len(txtSearch) >0 Then
         q="SELECT att.gangcode, att.fddate, att.employeecode, employee.fcname, att.jobcode," & _
         " j.fcname, att.locationtype, att.locationcode, att.supervision_1," & _
         " att.supervision_2, att.supervision_3, att.supervision_4," & _
         " att.attendance, att.mandays, att.unit, att.output, att.othrs," & _
         " att.rate, att.reference, att.remarks, att.fcba" & _
         " FROM attendance_gad att, employee employee, job j " & _
         " WHERE " & _
		 "((employee.fccode = att.EMPLOYEECODE and employee.fcba=att.fcba) AND (att.jobcode = j.fccode and att.fcba=j.fcba))" & _
		 " and att.fcba='" & session("fcBa") & "'" & txtWhere
	   elseif len(Request("fdDate"))>0 and len(request("GangCode")) >0 Then	 
         q="SELECT att.gangcode, att.fddate, att.employeecode, employee.fcname, att.jobcode," & _
         " j.fcname, att.locationtype, att.locationcode, att.supervision_1," & _
         " att.supervision_2, att.supervision_3, att.supervision_4," & _
         " att.attendance, att.mandays, att.unit, att.output, att.othrs," & _
         " att.rate, att.reference, att.remarks, att.fcba" & _
         " FROM attendance_gad att, employee employee, job j " & _
         " WHERE " & _
		 "((employee.fccode = att.EMPLOYEECODE and employee.fcba=att.fcba) AND (att.jobcode = j.fccode and att.fcba=j.fcba))" & _
		 " and att.fcba='" & session("fcBa") & "' and to_char(att.fdDate,'mm/dd/yyyy')='" & request("fdDate") & _
		 "' and att.Gangcode='" & request("GangCode") & "'"
	   else
         q="SELECT att.gangcode, att.fddate, att.employeecode, employee.fcname, att.jobcode," & _
         " j.fcname, att.locationtype, att.locationcode, att.supervision_1," & _
         " att.supervision_2, att.supervision_3, att.supervision_4," & _
         " att.attendance, att.mandays, att.unit, att.output, att.othrs," & _
         " att.rate, att.reference, att.remarks, att.fcba" & _
         " FROM attendance_gad att, employee employee, job j " & _
         " WHERE " & _
		 "((employee.fccode = att.EMPLOYEECODE and employee.fcba=att.fcba) AND (att.jobcode = j.fccode and att.fcba=j.fcba))" & _
		 " and att.fcba='" & session("fcBa") & "'" 
	   end if 
    Elseif txtTable="Empallded" Then
	  if len(txtField)>0 and len(txtSearch) >0 Then
         q=" SELECT emp.adhocdate,emp.empcode, employee.fcname, allowded.description, emp.amount FROM empallded emp, employee employee,"& _
	       " allowded allowded WHERE ( (emp.empcode = employee.fccode) AND (emp.allowdedcode = allowded.FCcode))" & _
		   " and emp.fcba='" & session("fcBa") & "'"  & txtWhere
	  else
         q=" SELECT emp.adhocdate,emp.empcode, employee.fcname, allowded.description, emp.amount FROM empallded emp, employee employee,"& _
	       " allowded allowded WHERE ( (emp.empcode = employee.fccode) AND (emp.allowdedcode = allowded.FCcode)) and emp.fcba='" & session("fcBa") & "'" 
	  end if	
    Elseif txtTable="PurchaseOrderDetails" Then
	   if len(txtField)>0 and len(txtSearch) >0 Then
		 txtColom = "SELECT pur.pocode, pur.podate, sup.fcName as Supplier, ord.itemcode, s.FCNAME,sum(ord.quantity) as Quantity," & _
		 " sum(ord.amount) as Amount,pur.remarks   FROM purchaseorder pur, purchaseorderdetails ord,stock s,Supplier sup" & _
		 " WHERE ((pur.pocode = ord.pocode)) and ((ord.itemcode = s.fcCode)) and ((pur.SUPPLIERCODE=Sup.FcCode))" & _
		 " and pur.fcba='" & session("fcBA") & "'" & txtWhere
	   elseif len(Request("poDate"))>0 and len(request("fcCode")) >0  Then
		 txtColom = "SELECT pur.pocode, pur.podate, sup.fcName as Supplier, ord.itemcode, s.FCNAME,sum(ord.quantity) as Quantity," & _
		 " sum(ord.amount) as Amount,pur.remarks From PurchaseOrder pur, purchaseorderdetails ord,stock s,Supplier sup" & _
		 " WHERE ((pur.pocode = ord.pocode)) and ((ord.itemcode = s.fcCode)) and ((pur.SUPPLIERCODE=Sup.FcCode))" & _
		 " and pur.fcba='" & session("fcBA") & "' and pur.PoCode='" & request("fcCode") & "'"
	   else
		 txtColom = "SELECT pur.pocode, pur.podate, sup.fcName as Supplier, ord.itemcode, s.FCNAME,sum(ord.quantity) as Quantity," & _
		 " sum(ord.amount) as Amount,pur.remarks From purchaseorder pur, purchaseorderdetails ord,stock s,Supplier sup" & _
		 " WHERE ((pur.pocode = ord.pocode)) and ((ord.itemcode = s.fcCode)) and ((pur.SUPPLIERCODE=Sup.FcCode))" & _
		 " and pur.fcba='" & session("fcBA") & "'" 
	   end if	
	   q=txtColom & " group by pur.pocode, pur.podate, sup.fcName, ord.itemcode, s.FCNAME,pur.remarks ORDER BY pur.pocode"	  
	else
	   if len(trim(txtSearch))=0 OR len(trim(txtField))=0 then
		   q = "Select * from "&txtTable 
		   if txtModule <>  "Data Setup" then
			  q = q + " where fcba='" & session("fcBA") & "'"
		   End if
	   else
		   q = "Select * from "&txtTable& " where " & txtField & " like '%" & uCase(txtSearch) & "%'" 
		   if txtModule <>  "Data Setup" then
			  q = q + " AND fcBa='" & session("fcBa") & "'"
		   End if
	   end if   
	end if	
	response.Write "sql : " & q
	%>

	</b><br>
	<form action="show.asp" method="post">
	<table style="font:Arial, Helvetica, sans-serif" style="font-size:9px" 
	border=1 vspace="1" style="border:thin" cellspacing="0" bgcolor="#CCFFCC">
	    <tr>
		<td align="center">No.</td>
	    <%	
		Set rsCat = connStr.Execute(q)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1
		for i = 0 to nCountofField%>
		    <td align="center"><%=rsCat(i).Name%></td>
		<%Next%>
		</tr>
		<%While Not rsCat.EOF%>
			<tr>
			<td align="left"><%=NO%></td>
			<%for i = 0 to nCountofField
			    if len(trim(rscat(i).value))>0 then%>
				   <td align="left"><%=rsCat(i)%></td>
				<%else%>
				   <td align="left">-</td>
				<%end if%>
			<%next%>
			</tr>
		  <%
		  rsCat.MoveNext
		  no=no+1
		Wend
		rsCat.close()
		rsCat=null%>
		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		</form>
	</table>
  <%'------------------ BAWAH ISI%>
  </td>
</tr>
</body>
</html>	