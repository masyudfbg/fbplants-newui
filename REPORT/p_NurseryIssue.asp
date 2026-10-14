<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Nursery Issue</title>
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

v_txtPeriodDate=REQUEST("PeriodDate")

v_txtPODATE_From=trim(request("DATEACTIVITY_From"))
v_txtPODATE_To=trim(request("DATEACTIVITY_To"))
v_NurseryCode=trim(request("NurseryCode"))
v_LOCATIONTYPE=trim(request("LOCATIONTYPE"))
v_LOCATIONCODE=trim(request("LOCATIONCODE"))

txtSQL = "select LINENOHIST, NURSERYCODE, ISSUEDATE, LOCATIONTYPE,LOCATIONCODE,JOBCODE,s.fcname,nvl(sum(QUANTITY),0) as Quantity,nvl(SUM(AMOUNT),0) AS AMOUNT  from NurseryIssue NO, JOB S where  " & _
         " (no.JOBCODE = S.FCCODE and no.fcba = s.fcba) and " & _
		  " no.fcBA='" & Session("fcBA") & "'"

if  len(trim(v_txtPODATE_From)) > 0 then
    txtSQL = txtSQL & " and ISSUEDATE >=to_Date('" & v_txtPODATE_From & "','MM/DD/YYYY')"
end if
if  len(trim(v_txtPODATE_To)) > 0 then
    txtSQL =txtSQL &  " and ISSUEDATE <=to_Date('" & v_txtPODATE_To & "','MM/DD/YYYY')"
end if
if  len(trim(v_NurseryCode)) > 0 then
    txtSQL = txtSQL & " and NurseryCode='" & v_NurseryCode & "'"
end if
if  len(trim(v_LOCATIONTYPE)) > 0 then
    txtSQL = txtSQL & " and LOCATIONTYPE='" & v_LOCATIONTYPE & "'"
end if
if  len(trim(v_LOCATIONCODE)) > 0 then
    txtSQL = txtSQL & " and LOCATIONCODE='" & v_LOCATIONCODE & "'"
end if

txtSQL=(txtSQL) & " group by LINENOHIST, NURSERYCODE, ISSUEDATE, LOCATIONTYPE,LOCATIONCODE,JOBCODE,s.fcname"
'response.Write(txtsql)
response.Flush()
%>
	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>NURSERY ISSUE</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="9"><%=txtSelect%></td>
						</tr>
						</div>
		<%	
		Set rsCat = connStr.Execute(txtSQL)
		if Not rsCat.eof then
		   nCountofField=rsCat.fields.count-1
		end if
		NO=1%>
		 
		<tr>
	     
		 <th>BATCH</th>
		 <th>ISSUE DATE</th>
		 <th>TARGET TYPE</th>
		 <th>TARGET CODE</th>
		 <th>JOB</th>
		 <th>JOB DESCRIPTION</th>
		 <th>QUANTITY</th>
		 <th>LINE HISTORY</th>
	</tr>
    <%
    Set rsCat = Server.CreateObject("ADODB.Recordset")
   rsCat.CursorType = 3
   rsCat.CursorLocation = 3
   rsCat.Open txtSql, connstr
   if Not rsCat.eof then
	nCountofField=rsCat.fields.count-1
	no=1
	nQty=0
	nAmount=0
	While Not rsCat.EOF%>
	  		<tr>
		 
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NURSERYCODE")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("ISSUEDATE")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONTYPE")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LOCATIONCODE")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JOBCODE")%></td>
			<td align="Left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcname")%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=formatnumber(rscat("QUANTITY"),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("LINENOHIST")%></td>
			</tr>
		  <%
				nQty=nQty+trim(cdbl(rsCat("QUANTITY")))
		  
		  rsCat.MoveNext
		  intRecordsPerPage = intRecordsPerPage - 1
		  no=no+1
	  Wend%>
	<tr bgcolor="#F4F4A0">
	<td align="Center" style="<%=cmdstyleLeft%>" colspan="6"> <font face="Verdana" size="1">T o t a l</td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(nQty,2)%> </td>
	<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
	</tr>	  
	  
   <%
   else
      response.Write("Data is empty")	  
   end if 
   rsCat.Close
   set rsCat = Nothing%>
   </TABLE>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>