<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Gang Activity and Attendance</title>
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

v_txtFDDATE_From=trim(REQUEST("FDDATE_From"))
v_txtFDDATE_To=trim(REQUEST("FDDATE_To"))
v_txtGANGCODE=trim(REQUEST("GANGCODE"))
v_txtJOBCODE=trim(REQUEST("JOBCODE"))
v_txtEMPLOYEECODE=trim(REQUEST("EMPLOYEECODE"))
v_txtLOCATIONTYPE=trim(REQUEST("LOCATIONTYPE"))
v_txtLOCATIONCODE=trim(REQUEST("LOCATIONCODE"))
v_txtreport=trim(REQUEST("REPORT"))
v_cField=trim(REQUEST("cField"))
v_cField1=trim(REQUEST("cField1"))
if v_txtreport="ACTIVITY" Then
	if  len(trim(v_cField)) = 0 and len(trim(v_cField1)) = 0 then
	   txtSQL = "SELECT NVL(sum(gad.mandays),0) as mandays,NVL(sum(gad.unit),0) as unit,NVL(sum(gad.Rate),0) as Rate,NVL(sum(gad.Output),0) as Output,nvl(sum(gad.othrs),0) as othrs" & _
			 " FROM attendance_gad gad, job job, employee employee " & _
			 " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
			 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
			 " AND gad.fcba='" & session("fcBA") & "'"
	else
		if  len(trim(v_cField)) = 0 then
		   txtSQL = "SELECT " & v_cField1 & _
				 " ,NVL(sum(gad.mandays),0) as mandays,NVL(sum(gad.unit),0) as unit,NVL(sum(gad.Rate),0) as Rate,NVL(sum(gad.Output),0) as Output,nvl(sum(gad.othrs),0) as othrs" & _
				 " FROM attendance_gad gad, job job, employee employee " & _
				 " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
				 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
				 " AND gad.fcba='" & session("fcBA") & "'"
		else
			if instr(v_cField1,"job.fcname as jobname") > 0 then
			   txtSQL = "SELECT " & v_cField & "," & v_cField1 & _
					 " ,NVL(sum(gad.mandays),0) as mandays,NVL(sum(gad.unit),0) as unit,NVL(sum(gad.Rate),0) as Rate,NVL(sum(gad.Output),0) as Output,nvl(sum(gad.othrs),0) as othrs" & _
					 " FROM attendance_gad gad, job job, employee employee " & _
					 " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
					 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
					 " AND gad.fcba='" & session("fcBA") & "'"
			else
			   txtSQL = "SELECT " & v_cField  & _
					 " ,NVL(sum(gad.mandays),0) as mandays,NVL(sum(gad.unit),0) as unit,NVL(sum(gad.Rate),0) as Rate,NVL(sum(gad.Output),0) as Output,nvl(sum(gad.othrs),0) as othrs" & _
					 " FROM attendance_gad gad, job job, employee employee " & _
					 " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
					 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
					 " AND gad.fcba='" & session("fcBA") & "'"
			end if
		end if
	end if
else
	if  len(trim(v_cField)) = 0 and len(trim(v_cField1)) = 0 then
	   txtSQL = "SELECT gad.ATTENDANCE" & _
			 " FROM attendance_gad gad, job job, employee employee " & _
			 " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
			 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
			 " AND gad.fcba='" & session("fcBA") & "'"
	Else
		if  len(trim(v_cField)) = 0 then
		   txtSQL = "SELECT " & v_cField1 & _
				 " ,gad.ATTENDANCE" & _
				 " FROM attendance_gad gad, job job, employee employee " & _
				 " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
				 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
				 " AND gad.fcba='" & session("fcBA") & "'"
		else
			if instr(v_cField1,"job.fcname as jobname") > 0 then
			   txtSQL = "SELECT " & v_cField & "," & v_cField1 & _
					 " ,gad.ATTENDANCE" & _
					 " FROM attendance_gad gad, job job, employee employee " & _
					 " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
					 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
					 " AND gad.fcba='" & session("fcBA") & "'"
			Else
			   txtSQL = "SELECT " & v_cField & _
					 " ,gad.ATTENDANCE" & _
					 " FROM attendance_gad gad, job job, employee employee " & _
					 " WHERE ( ((gad.jobcode = job.fccode) AND (gad.fcBA = job.fcBA)) and " & _
					 " ((gad.employeecode = employee.fccode) and (gad.fcBA = employee.fcBA)) )" & _
					 " AND gad.fcba='" & session("fcBA") & "'"
			End if
		end if
	End if
end if


if  len(trim(v_txtFDDATE_From)) > 0 then
    txtSQL = txtSQL & " and gad.fddate >=to_Date('" & v_txtFDDATE_From & "','MM/DD/YYYY')"
	txtSelect = " From " & FormatDateTime(v_txtFDDATE_From,1)
end if
if  len(trim(v_txtFDDATE_To)) > 0 then
    txtSQL =txtSQL &  " and gad.fddate <=to_Date('" & v_txtFDDATE_To & "','MM/DD/YYYY')"
	txtSelect = txtSelect & "  to " & FormatDateTime(v_txtFDDATE_To,1)
end if
if  len(trim(v_txtGANGCODE)) > 0 then
    txtSQL = txtSQL & " and gad.gangcode='" & v_txtGANGCODE & "'"
	txtSelect = txtSelect & "  Gang Code " & (v_txtGANGCODE)
end if
if  len(trim(v_txtJOBCODE)) > 0 then
    txtSQL =txtSQL &  " and gad.jobcode='" & v_txtJOBCODE & "'"
	txtSelect = txtSelect & "  Job Code " & (v_txtJOBCODE)
end if 
if  len(trim(v_txtEMPLOYEECODE)) > 0 then
    txtSQL =txtSQL &  " and gad.employeecode='" & v_txtEMPLOYEECODE & "'"
	txtSelect = txtSelect & "  Emp. Code " & (v_txtEMPLOYEECODE)
end if 
if  len(trim(v_txtLOCATIONTYPE)) > 0 then
    txtSQL =txtSQL &  " and gad.locationtype='" & v_txtLOCATIONTYPE & "'"
	txtSelect = txtSelect & "  TT " & (v_txtLOCATIONTYPE)
end if 
if  len(trim(v_txtLOCATIONCODE)) > 0 then
    txtSQL =txtSQL &  " and gad.locationcode='" & v_txtLOCATIONCODE & "'"
	txtSelect = txtSelect & "  TC " & (v_txtLOCATIONCODE)
end if 

if v_txtreport="ATTENDANCE" Then
   txtSQL = txtSQL
else
	if  len(trim(v_cField)) = 0 and len(trim(v_cField1)) = 0 then
		txtSQL = txtSQL
	Else
		if instr(v_cField1,"job.fcname as jobname") > 0 then
			if  len(trim(v_cField)) = 0 then
			   txtsql=(txtsql) & " group by job.fcname"
			else
			   txtsql=(txtsql) & " group by " & v_cField & ",job.fcname"
			end if
		Else
		   txtsql=(txtsql) & " group by " & v_cField
		End if	
	End if
end if
if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if
'response.Write(txtsql)
'response.Flush()
%>
	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>ATTENDANCE GAD</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="12"><%=txtSelect%></td>
						</tr>
						</div>

  <% Dim Kolom
  Kolom =0
if instr(v_cField,"gad.fddate") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"gad.employeecode") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"employee.fcname") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"gad.jobcode") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField1,"job.fcname as jobname") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"gad.locationtype") > 0 then
	 Kolom=Kolom + 1
End if
if instr(v_cField,"gad.locationcode") > 0 then
	 Kolom=Kolom + 1
End if
	Kolom=KOLOM%>
	
<%if V_txtreport="ACTIVITY" then%>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	<tr>
	   
        <%if instr(v_cField,"gad.fddate") > 0 then%>
			<th>Date</th>
		<%end if%>
        <%if instr(v_cField,"gad.employeecode") > 0 then%>
			<th>Employee Code</th>
		<%end if%>
        <%if instr(v_cField,"employee.fcname") > 0 then%>
			<th>Employee Name</th>
		<%end if%>
		<%if instr(v_cField,"gad.jobcode") > 0 then%>
			<th>Jobcode</th>	
		<%end if%>
        <%if instr(v_cField1,"job.fcname as jobname") > 0 then%>
			<th>Job Description</th>	
		<%end if%>
        <%if instr(v_cField,"gad.locationtype") > 0 then%>
			<th>Target Type</th>
		<%end if%>
        <%if instr(v_cField,"gad.locationcode") > 0 then%>
			<th>Target Code</th>
		<%end if%>
				<th>Mandays</th>
				<th>Unit</th>
				<th>Rate</th>
				<th>Output</th>
				<th>Over Times</th>
	</tr>
	<%
	Dim v_mandays,v_unit,v_Rate,v_Output,v_othrs
	v_mandays=0
	v_unit=0
	v_Rate=0
	v_Output=0
	v_othrs=0
	While Not rsCat.EOF%>

		<tr>
		
		<%if instr(v_cField,"gad.fddate") > 0 then%>
			 <td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fddate")%></td>
		<%End if%>
        <%if instr(v_cField,"gad.employeecode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("employeecode")%> </td>
		<%End if%>
        <%if instr(v_cField,"employee.fcname") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%> </td>
		<%End if%>
		<%if instr(v_cField,"gad.jobcode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%> </td>
		<%End if%>
        <%if instr(v_cField1,"job.fcname as jobname") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobname")%> </td>
		<%End if%>
        <%if instr(v_cField,"gad.locationtype") > 0 then%>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationtype")%> </td>
		<%End if%>
        <%if instr(v_cField,"gad.locationcode") > 0 then%>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Locationcode")%> </td>
		<%End if%>
    	<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("mandays"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("unit"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Rate"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("Output"),2)%></td>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("othrs"),2)%></td>
		</tr>
	  <%
	  v_mandays=v_mandays+trim(cdbl(rsCat("mandays")))
	  v_unit  =v_unit+Trim(cdbl(rsCat("unit")))
	  v_Rate =v_Rate+Trim(cdbl(rsCat("Rate")))
	  v_Output =v_Output+Trim(cdbl(rsCat("Output")))
	  v_othrs=v_othrs+Trim(cdbl(rsCat("othrs")))
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
	<tr bgcolor="#F4F4A0">
	<td align="Center" style="<%=cmdstyleLeft%>" colspan= "<%=KOLOM%>"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_mandays,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_unit,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Rate,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_Output,2)%></td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(v_othrs,2)%></td>	
	</tr>

<%Elseif V_txtreport="ATTENDANCE" then%>
	<%	
	Set rsCat = connStr.Execute(txtSQL)
	if Not rsCat.eof then
	   nCountofField=rsCat.fields.count-1
	end if
	NO=1%>
	<tr>
	    <td style="<%=cmdstyleLeft%>;border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"><font face="Verdana" size="1">NO</td>
		<%if instr(v_cField,"gad.fddate") > 0 then%>
			<th>Date</th>
		<%end if%>
        <%if instr(v_cField,"gad.employeecode") > 0 then%>
			<th>Employee<br>Code</th>
		<%end if%>
        <%if instr(v_cField,"employee.fcname") > 0 then%>
			<th>Employee Name</th>	
		<%end if%>
		<%if instr(v_cField,"gad.jobcode") > 0 then%>
			<th>Jobcode</th>	
		<%end if%>
        <%if instr(v_cField1,"job.fcname as jobname") > 0 then%>
			<th>Job Description</th>	
		<%end if%>
        <%if instr(v_cField,"gad.locationtype") > 0 then%>
			<th>Target Type</th>
		<%end if%>
        <%if instr(v_cField,"gad.locationcode") > 0 then%>
			<th>Target Code</th>
		<%end if%>		
			<th>Attendance</th>
	</tr>
	<%
	While Not rsCat.EOF%>
		<tr>
		<td align="left" style="<%=cmdstyleLeft%>"><font face="Verdana" size="1"><%=NO%></td>
		<%if instr(v_cField,"gad.fddate") > 0 then%>
			 <td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fddate")%></td>
		<%End if%>
        <%if instr(v_cField,"gad.employeecode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("employeecode")%> </td>
		<%End if%>
        <%if instr(v_cField,"employee.fcname") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%> </td>
		<%End if%>
		<%if instr(v_cField,"gad.jobcode") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobcode")%> </td>
		<%End if%>
        <%if instr(v_cField1,"job.fcname as jobname") > 0 then%>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("jobname")%> </td>
		<%End if%>
        <%if instr(v_cField,"gad.locationtype") > 0 then%>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("locationtype")%> </td>
		<%End if%>
        <%if instr(v_cField,"gad.locationcode") > 0 then%>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("Locationcode")%> </td>
		<%End if%>
		<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("attendance")%></td>
		</tr>
	  <%
	  'response.write cdbl(trim(rsCat("MTHOPENBALANCE")))+1
	  if trim(rscat("attendance")) ="KJ" then
	     v_KJ = v_KJ + 1
	  elseif trim(rscat("attendance")) ="P1" then
	     v_P1 = v_P1 + 1
	  elseif trim(rscat("attendance")) ="P2" then
	     v_P2 = v_P2 + 1
	  elseif trim(rscat("attendance")) ="P3" then
	     v_P3 = v_P3 + 1
	  elseif trim(rscat("attendance")) ="S1" then
	     v_S1 = v_S1 + 1
	  elseif trim(rscat("attendance")) ="S2" then
	     v_S2 = v_S2 + 1
	  elseif trim(rscat("attendance")) ="H1" then
	     v_H1 = v_H1 + 1
	  elseif trim(rscat("attendance")) ="H2" then
	     v_H2 = v_H2 + 1
	  elseif trim(rscat("attendance")) ="CT" then
	     v_CT = v_CT + 1
	  elseif trim(rscat("attendance")) ="MK" then
	     v_MK = v_MK + 1
	  elseif trim(rscat("attendance")) ="ML" then
	     v_ML = v_ML + 1
	  elseif trim(rscat("attendance")) ="WH" then
	     v_WH = v_WH + 1
	  elseif trim(rscat("attendance")) ="WS" then
	     v_WS = v_WS + 1
	  elseif trim(rscat("attendance")) ="UN" then
	     v_UN = v_UN + 1
      End if		 
	  rsCat.MoveNext
	  'response.Flush()
	  no=no+1
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">KERJA (KJ)</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_KJ,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">KERJA (WS)</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_WS,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">KERJA (WH)</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_WH,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">UN</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_UN,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">ML</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_ML,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">MK</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_MK,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">P1</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_P1,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">P2</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_P2,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">P3</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_P3,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">S1</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_S1,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">S2</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_S2,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">H1</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_H1,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">H2</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_H2,2)%>&nbsp;</TD></tr>
	<tr><td align="RIGHT" style="<%=cmdstyleLeft%>" colspan=<%=KOLOM%>><font face="Verdana" size="1">CT</td>
	<TD style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(v_CT,2)%>&nbsp;</TD></tr>
<%end if%>
</table>
<!--#include file="Bottom_Report.asp"-->

		</DIV>
	</BODY>
</HTML>	