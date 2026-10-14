<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Upkeep Rotation</title>
</head>
<body>
<%
txtYear = request("txtYear")
txtMonth = request("txtMonth")
jobcode = request("jobcode")
fieldcode = request("fieldcode")

	txtSQL = " SELECT ur.fcba" & _
	         "        , ur.lineno" & _
	         "        , ur.period_year" & _
	         "        , ur.period_month" & _
	         "        , ur.jobcode" & _
	         "        , j.fcname jobname" & _
	         "        , ur.fieldcode" & _
	         "        , ur.mandays" & _
	         "        , ur.orderdata" & _
             " FROM upkeeprotation ur" & _
			 " LEFT JOIN job j ON (ur.fcba = j.fcba AND ur.jobcode = j.fccode)" & _
             " WHERE ur.period_year = '" & txtYear & "'" & _
             "       AND ur.fcba = '" & session("fcBA") & "'"
	 
if  len(trim(txtYear)) > 0 then
	txtSelect = " Tahun " & txtYear
end if
if  len(trim(txtMonth)) > 0 then
	txtSQL = txtSQL &  " AND ur.period_month = '" & txtMonth & "'"
	txtSelect = txtSelect & " Bulan " & Bulan(txtMonth)
end if
if  len(trim(jobcode)) > 0 then
    txtSQL = txtSQL &  " AND ur.jobcode = '" & jobcode & "'"
	txtSelect = txtSelect & " Jobcode " & jobcode
end if 
if  len(trim(fieldcode)) > 0 then
    txtSQL =txtSQL &  " AND ur.fieldcode = '" & fieldcode & "'"
	txtSelect = txtSelect & " Fieldcode " & fieldcode
end if 

txtsql = txtsql & " ORDER BY ur.period_month, ur.jobcode, ur.orderdata "

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
end if

' response.Write txtsql
Set rsCat = connStr.Execute(txtSQL)
%>
	<body>
		<div id="wrapper">
			<div id="header-wrapper">
				<div id="header">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
				<div id="subheader" class="centered">
					<h2>Upkeep Rotation</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="5"><%=txtSelect%></td>
						</tr>
						</div>
	</tr>
	<tr>
	    <th>Month</th>
		<th>Jobcode</th>
		<th>Description</th>
		<th>Fieldcode</th>
		<th>Mandays</th>
	</tr>
	<%While Not rsCat.EOF%>
	 	<tr>
		 
		<td align="center"><font face="Verdana" size="1"><%=rscat("period_month")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("jobcode")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("jobname")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("fieldcode")%></td>
		<td align="right"><font face="Verdana" size="1"><%=rscat("mandays")%></td>
		</tr>
	  <%
	  rsCat.MoveNext
	Wend
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
	</form>
</table>
<!--#include file="Bottom_Report.asp"-->
</DIV>
</BODY>
</HTML>