<%Option Explicit%>
<%response.buffer=True%>
<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include virtual="txtFunction.inc"-->
<html>
<head>
<title>iPS</title>
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
</head>
<body>
<!--#include file="top.jsp"-->
<table width="100%">
<tr>
 <td width="22%" valign="top">
 <!--#include file="menu.asp"-->
 </td>
 <td width="78%" valign="top">
    <%' -----------------ATAS ISI%> <%
	Dim txtTable,txtModule,q,nCountofField,NO,txtSearch,txtField,txtColom,txtWhere,txtPeriod,txtOrder
	Dim i
	txtTable =Request("txtTable")
	txtModule=Request("txtModule")
	txtSearch=Request("txtSearch")
	txtField =Request("txtField")
	if (len(trim(txtSearch))=0 OR len(trim(txtField))=0) and (txtModule="Data Master" or txtModule="Data Setup") then
	   '----------- Special for Table Master & Setup -------------------------'
	   q = "Select * from "&txtTable 
	   if txtModule <>  "Data Setup" then
          q = q + " where fcba='" & session("fcBA") & "'"
	   End if
	elseif len(trim(txtSearch))<>0 and len(trim(txtField))<>0 and (txtModule="Data Master" or txtModule="Data Setup") then
	   txtWhere=""
	   txtColom=""
	   if instr(txtSearch,"/") = 0 then
	      if instr(txtSearch,"=") > 0 then
	         txtWhere = " where " & txtField & uCase(txtSearch) 
		  elseif txtField <> "" and txtSearch <> "" then
	         txtWhere = " where " & txtField & " like '%" & uCase(txtSearch) & "%'"
		  end if
	   else
	         txtWhere = " where " & txtField & " like TO_DATE('" & txtSearch & "','DD/MM/YYYY')"
	   end if

	   if txtWhere <> "" and txtModule <>  "Data Setup" then
	      q = "Select * from " & txtTable & txtWhere + " and fcba='" & session("fcBA") & "'"
       elseif txtWhere <> "" and txtModule =  "Data Setup" then
	      q = "Select * from " & txtTable & txtWhere
	   else
	      q = "Select * from " & txtTable	  
	   end if
	else
	   '----------- Special for Data Transaction -----------------------------'
	   txtWhere=""
	   txtColom=""
	   if instr(txtSearch,"/") = 0 then
	      if instr(txtSearch,"=") > 0 then
	         txtWhere = " and " & txtField & uCase(txtSearch) 
		  elseif txtField <> "" and txtSearch <> "" then
	         txtWhere = " and " & txtField & " like '%" & uCase(txtSearch) & "%'"
		  end if
	   else
	         txtWhere = " and " & txtField & " like TO_DATE('" & txtSearch & "','DD/MM/YYYY')"
	   end if
	   
       If txtTable="Attendance_GAD" then
	      txtColom = "SELECT att.fddate,att.employeecode, employee.fcname, job.fcname," & _
		             " att.locationtype, att.locationcode,att.mandays,att.unit,att.othrs" & _
					 " FROM attendance_gad att, job job, employee employee " & _
                     " WHERE ((att.jobcode = job.fccode) AND (att.employeecode = employee.fccode)) " & _
					 " AND ATT.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(att.fddate,'mm-yyyy')='" & "01-2006" & "'"
		  txtOrder = " ORDER BY EMPLOYEECODE,FDDATE"					 
	   elseif txtTable="PaymentVoucherDetail" then
	      txtColom = "SELECT p.datecreated,pay.vouchercode, pay.jobcode, job.fcname, pay.locationtype," & _
		             " pay.locationcode, pay.amount, pay.REFERENCE, pay.remarks" & _
                     " FROM paymentvoucherdetail pay, job job,paymentvoucher p" & _
                     " WHERE ((pay.jobcode = job.fccode)) and ((pay.vouchercode = p.vouchercode))" & _
					 " AND pay.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(p.datecreated,'mm-yyyy')='" & "01-2006" & "'"
		  txtOrder = " ORDER BY pay.vouchercode"					 		  
	   elseif txtTable="ReceiveNote" then
	      txtColom = "SELECT rec.receivenotecode, rec.pocode, rec.rndate, rd.locationtype," & _
                     "rd.locationcode, stock.fcname, rd.quantity, rd.qtyreturn," & _
                     "rd.amountreceived, rd.amountreturn  FROM receivenote rec, receivenotedetail rd, stock stock" & _
                     " WHERE (    (rec.receivenotecode = rd.receivenotecode)" & _
                     " AND (stock.fccode = rd.purchaseitemcode))" & _
					 " AND rec.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(rec.rndate,'mm-yyyy')='" & "03-2006" & "'"
		  txtOrder = " ORDER BY rec.receivenotecode"					 		  
	   elseif txtTable="IssueDetails" then
	      txtColom = "SELECT issue.sivdate,iss.sivcode, iss.stockcode, stock.fcname, iss.locationcode," & _
                     "iss.locationtype, iss.quantity, iss.price, iss.amount, job.fcname" & _
                     " FROM issuedetails iss, issue issue, stock stock, job job" & _
                     " WHERE (    (iss.sivcode = issue.sivcode)" & _
                     " AND (iss.stockcode = stock.fccode) AND (iss.jobcode = job.fccode))" & _
					 " AND iss.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(issue.sivdate,'mm-yyyy')='" & "03-2006" & "'"
		  txtOrder = " ORDER BY iss.sivcode"
	   elseif txtTable="WorkshopActivity" then
		  txtColom = "SELECT wor.dateactivity, wor.workshopcode, job.fcname, wor.locationtype," & _
					 " wor.LOCATION, wor.TIME, wor.hours, wor.amount" & _
					 " FROM workshopactivity wor, job job" & _
					 " WHERE ((wor.job = job.fccode))" & _
					 " AND wor.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(wor.dateactivity,'mm-yyyy')='" & "03-2006" & "'"
		  txtOrder = " ORDER BY wor.dateactivity"
	   elseif txtTable="VehicleActivity" then
		  txtColom = "SELECT veh.vehdate, veh.vehiclecode, veh.jobcode, job.fcname, veh.TIME," & _
					 " veh.locationtype, veh.locationcode, veh.quantity, veh.amount" & _
					 " FROM vehicleactivity veh, job job" & _
					 " WHERE ((veh.jobcode = job.fccode))" & _
					 " AND veh.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(veh.vehdate,'mm-yyyy')='" & "03-2006" & "'"
		  txtOrder = " ORDER BY veh.vehiclecode"
	   elseif txtTable="AGREEMENTDETAIL" then
		  txtColom = "SELECT * FROM WORKINPROGRESS" & _
		             " Where fcba='" & session("fcBA") & "'"
		  txtPeriod= " and to_char(WIPDATE,'mm-yyyy')='" & "03-2006" & "'"
		  txtOrder = " ORDER BY AGREEMENTCODE"
	   elseif txtTable="NurseryIssue" then
		  txtColom = "SELECT * FROM NURSERYISSUE" & _
		             " Where fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(ISSUEDATE,'mm-yyyy')='" & "03-2006" & "'"
		  txtOrder = " ORDER BY NURSERYCODE"
	   elseif txtTable="Empallded" then
		  txtColom = "SELECT emp.adhocdate,emp.empcode, employee.fcname, allowded.description," & _
					 " emp.amount FROM empallded emp, employee employee, allowded allowded" & _
					 " WHERE (    (emp.empcode = employee.fccode)" & _
					 " AND (emp.allowdedcode = allowded.allowdedcode))" & _
		             " and emp.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND to_char(emp.adhocdate,'mm-yyyy')='" & "03-2006" & "'"
		  txtOrder = " ORDER BY emp.empcode,emp.adhocdate"	  
	   elseif txtTable="HarvestingQuality" then
		  txtColom = "SELECT har.empcode, employee.fcname, har.fieldcode, har.under_ripe," & _
					 " har.over_ripe, har.abnormal, har.long_stalk, har.eaten_by_rat," & _
					 " har.unharvest_ffb, har.uncollect_lf_circle, har.uncollect_lf_piece," & _
					 " har.unarrange_ffb, har.unprune_frond, har.DAY, har.MONTH, har.YEAR" & _
					 " FROM harvestingquality har, employee employee" & _
					 " WHERE ((har.empcode = employee.fccode))" & _
		             " and har.fcba='" & session("fcBA") & "'"
		  txtPeriod= " AND har.Month='" & "3" & "' and har.Year='" & "2006" & "'"
		  txtOrder = " ORDER BY har.empcode"	  
	   elseif txtTable="Allowded" then
		  txtColom = "SELECT allowdedcode,description,jobcode from allowded har" & _
					 " WHERE har.fcba='" & session("fcBA") & "'"
		  txtOrder = " ORDER BY har.allowdedcode"	  
	   else
	      txtColom = "SELECT * FROM " & txtTable 
	   end if
	   
	   if txtWhere <> "" and txtPeriod<> "" and txtOrder <> ""  then
	      q = txtColom & txtPeriod & txtWhere & txtOrder
	   elseif txtWhere = "" and txtPeriod<> "" and txtOrder <> ""  then
	      q = txtColom & txtPeriod & txtOrder
	   else
	      q = txtColom
	   end if
	end if   
	Response.Write q
	%>
	<b><FONT face="Verdana, Arial, Helvetica, sans-serif" style="font-size:10px">
	   Module : <%=request("txtModule")%> / <%=txtTable%></FONT>
	</b><br>
	<form action="show.asp" method="post">
	<table style="font:Arial, Helvetica, sans-serif" style="font-size:9px" class="sample" cellspacing="0">
	    <%	
		Set rsCat = Server.CreateObject("ADODB.Recordset")
		rsCat.CursorType = 3
		rsCat.CursorLocation = 3
		rsCat.Open q, connstr
	   
		'Set rsCat = connStr.Execute(q)
		if Not rsCat.eof then
		    nCountofField=rsCat.fields.count-1
		end if
		NO=1%>
        
		<tr>
			<td colspan="<%=nCountofField+1%>" align="left" valign="middle"> 
              <%
			Dim txtLink
			if txtModule="Data Setup" or txtModule="Data Master" then
			   txtLink = "TableMaster"
			else
			   txtLink = txtTable
			end if
			%>
			<a href="<%=txtLink%>.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>"
			title="New : <%=txtTable%>">
			<img src="IMAGES/add2.gif" border="0">
			</a>
			Search Like Description : 
			<input type="text" title="Case is not Sensitive for date [dd/mm/yyyy]" name="txtSearch" 
			value="<%=txtSearch%>">&nbsp;from Column
			<input type="text" title="Case is not Sensitive" name="txtField" value="<%=txtField%>">&nbsp;&nbsp;
			<input type="submit" value="Select"
			style="background-color: #298CF7; color: #F7C610;font-weight: bold; font-size: 8pt; font-family:Verdana">
			<a href="PrintData.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>"
			 target="_blank">
			<img src="IMAGES/Print.gif" border="0" title="Print your Data">Print
			</a>
			</td>
		</tr>
        <!--#include file="judulshow.asp"-->
		<%
		Dim intRecordsPerPage,intNumberOfPages,intRecords,nPage,howmanyfields,JumRec
        howmanyfields=rsCat.fields.count-1
  
        JumRec=1
        intRecordsPerPage = 15
		
        rsCat.PageSize = intRecordsPerPage
        intNumberOfPages = 0
        intNumberOfPages =  rsCat.PageCount
        intRecords = rsCat.RecordCount
	    'Set our AbsolutePage property.
		if not rsCat.EOF then
			If (isEmpty(Request("iPageNo")) or trim(Request("iPageNo"))="") Then
			  rsCat.AbsolutePage = 1
			Else
			  rsCat.AbsolutePage = Request("iPageNo")
			End If  
        end if
	    if Request("iPageNo") = "" then
			nPage = 1
	    Else
			nPage = Request("iPageNo")
	    End if
		Dim cColor
		While Not rsCat.EOF and intRecordsPerPage > 0 
		    if intRecordsPerPage  mod 2 = 1 then
			   cColor = "#CCFF33"
			else
			   cColor = "#33FF00"
			end if   
		    %>
          
			<tr bgcolor="<%=cColor%>"> 
            <td> <a href="EditData.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>&txtCode=<%=rsCat(0)%>&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>"><img src="IMAGES/edit.gif" alt="Edit Data : <%=rsCat(1)%>" border="0"></a><a href="DeleteData.Asp?txtModule=<%=txtModule%>&txtTable=<%=txtTable%>&txtCode=<%=rsCat(0)%>&txtSearch=<%=txtSearch%>&txtField=<%=txtField%>&iPageno=<%=Request("iPageNo")%>"><img src="IMAGES/delete.gif" alt="Delete Data : <%=rsCat(1)%>" border="0"></a></td>
            <td align="left" valign="top"><%=NO%></td>
            <%for i = 0 to nCountofField%>
            <td align="left" title="<%=rsCat(i).Name%>" valign="top"><%=rsCat(i)%></td>
            <%next%>
            
			</tr>
		    
          <%
		    rsCat.MoveNext
		    intRecordsPerPage = intRecordsPerPage - 1
		    no=no+1
		Wend
		rsCat.close()
		rsCat=null%>
          
		<input  type="hidden" value="<%=txtTable%>" name="txtTable">
		<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		</form>
		<tr>
		<td colspan="10">
		   <%if intNumberOfPages<>1 then
			  For i = 1 to intNumberOfPages
				  Response.Write "<font size=""2"" face=""Arial Narrow"" color=""#FFFFFF""><a href=""show.asp?txtModule=" &_
				  txtModule & "&iPageNo=" &  i & "&txtTable=" & request("txtTable")& "&txtField=" & request("txtField")  & _
				  "&txtSearch=" & request("txtSearch") & _
				  """>" & i & "  "
			  Next
		   End if%>
		   </a><br>
		   <%if Request("iPageNo") = "" then%> <%nPage = 1%> <%Else%> <%nPage = Request("iPageNo")%> <%End if%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Page : <%=nPage%>
		   <font size="1" color="#000000" face="Arial, Helvetica, sans-serif">Total Records : <%=intRecords%></td>
		</tr>

		
	</table>
  
<%'------------------ BAWAH ISI%>
  </td>
</tr>
<tr>
  <td colspan="2">
  <!--#include file="bottom.jsp"-->
  </td>
</tr>
</table>
</body>
</html>