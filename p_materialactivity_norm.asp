<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Material Activity Norm</title>
</head>
<body>
<%
jobcode = request("jobcode")
actmat = request("actmat")

	txtSQL = " SELECT job.fccode jobcode, job.fcname jobname, activematerial.fccode actmatcode, activematerial.fcname actmatname, " & _
			 " ur.UOMJOB, ur.ACTMAT, ur.UOMACTIVEMATERIAL, ur.ACTMATNORM" & _
             " FROM MATERIALACTIVITY_NORM ur" & _
			 " inner join job on ur.job = job.fccode and ur.fcba = job.fcba" & _
			 " inner join activematerial on ur.actmat = activematerial.fccode"  & _
             " WHERE ur.fcba = '" & session("fcBA") & "'"

if  len(trim(jobcode)) > 0 then
	txtSQL = txtSQL &  " AND ur.job = '" & jobcode & "'"
end if
if  len(trim(actmat)) > 0 then
	txtSQL = txtSQL &  " AND ur.ACTMAT = '" & actmat & "'"
end if

'response.Write txtsql
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
					<h2>Material Activity Norm</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="7">&nbsp;</td>
						</tr>
						</div>
	</tr>
	<tr>
		<th>Jobcode</th>
		<th>Jobcode UoM</th>
		<th>Active Material</th>
		<th>Active Material UoM</th>
		<th>Norm</th>
	</tr>
	<%While Not rsCat.EOF%>
	 	<tr>
		<td align="left"><font face="Verdana" size="1"><%=rscat("jobcode")%>-<%=rscat("jobname")%></td>		
		<td align="left"><font face="Verdana" size="1"><%=rscat("UOMJOB")%></td>
		<td align="left"><font face="Verdana" size="1"><%=rscat("actmatcode")%>-<%=rscat("actmatname")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("UOMACTIVEMATERIAL")%></td>		 
		<td align="left"><font face="Verdana" size="1"><%=rscat("ACTMATNORM")%></td>
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