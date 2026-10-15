<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Monthly Work Plan</title>
</head>
<body>
<%
txtYear = request("txtYear")
txtMonth = request("txtMonth")
jobcode = request("jobcode")
fieldcode = request("fieldcode")
txtTempJobcode = "1"
txtFieldcode = "1"
txtDivision = "1"

	txtSQL = " SELECT ur.fcba" & _
	         "        , ur.rkbno" & _
	         "        , ur.period_year" & _
	         "        , ur.period_month" & _
	         "        , jobgroup.fccode jobgroupcode " & _
	         "        , jobgroup.fcname jobgroupname " & _	 
			 "        , ur.activitycode" & _
	         "        , ur.fieldcode" & _
			 "        , field.division" & _
	         "        , j.fcname jobname" & _
			 "        , ur.mandays" & _
			 "        , ur.volume" & _
			 "        , ur.boronganpremi" & _
			 "        , ur.rate" & _
			 "        , ur.totalamount totalamounthead" & _
	         "        , stock.fcname material" & _
	         "        , det.qty" & _
			 "        , det.ratestock" & _
			 "        , det.totalamount" & _
			 "        , det.haplan" & _
			 "        , det.kg" & _
			 "        , det.jjg" & _
			 "        , det.rotasi" & _
			 "        , det.remarks" & _
             " FROM monthlyworkplandet det" & _
			 " inner join monthlyworkplan ur on det.rkbno = ur.rkbno and  det.fcba = ur.fcba" & _
			 " left join stock on det.fcba = stock.fcba and det.stockcode = stock.fccode"&_
			 " LEFT JOIN job j ON (ur.fcba = j.fcba AND ur.activitycode = j.fccode)" & _
			 " inner join jobgroup on j.JOBGROUPCODE = jobgroup.fccode and det.fcba = jobgroup.fcba" & _	
			 " inner join field on ur.fieldcode = field.fccode and ur.fcba = field.fcba	" & _
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
    txtSQL = txtSQL &  " AND ur.activitycode = '" & jobcode & "'"
	txtSelect = txtSelect & " activitycode " & jobcode
end if 
if  len(trim(fieldcode)) > 0 then
    txtSQL =txtSQL &  " AND ur.fieldcode = '" & fieldcode & "'"
	txtSelect = txtSelect & " Fieldcode " & fieldcode
end if 

txtsql = txtsql & " ORDER BY ur.period_month, ur.activitycode, det.lineno "

if len(trim(txtSelect)) > 0 then
   txtSelect = "Selected  by : " & txtSelect
else
   txtSelect = "Selected by All Data"
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
					<h2>Monthly Work Plan</h2>
				</div>
			</div>
					<table border="0" width="100%" id="product-table" class="headertable">
						<div class="printhide">
						<tr>
							<td colspan="22"><%=txtSelect%></td>
						</tr>
						</div>
	</tr>
	<tr>
		<th>RKB No</th>
	    <th>Month</th>
		<th>Group Code</th>
		<th>Group Name</th>
		<th>Jobcode</th>
		<th>Description</th>
		<th>Mandays</th>
		<th>Rate Mandays</th>
		<th>Amount Mandays</th>
		<th>Volume</th>
		<th>Borongan/Premi</th>
		<th>Division</th>
		<th>Fieldcode</th>
		<th>ha Plan</th>
		<th>Material</th>
		<th>Quantity</th>		
		<th>Rate</th>
		<th>Amount</th>		
		<th>Kilogram</th>
		<th>Janjang</th>
		<th>Rotasi</th>
		<th>Remarks</th>
	</tr>
	<%
	
	%>
	<%Do While Not rsCat.eof%>
	 	<tr>
		 <td align="center"><font face="Verdana" size="1"><%=rscat("rkbno")%></td>
		<td align="center"><font face="Verdana" size="1"><%=rscat("period_month")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("jobgroupcode")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("jobgroupname")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("activitycode")%></td>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("jobname")%></td>
		<%		
		if txtTempJobcode = Cdbl(rscat("activitycode")) then
		%>
			<td align="right"><font face="Verdana" size="1"></td>
			<td align="right"><font face="Verdana" size="1"></td>
			<td align="right"><font face="Verdana" size="1"></td>
			<td align="right"><font face="Verdana" size="1"></td>
			<td align="right"><font face="Verdana" size="1"></td>			
		<%
		else
		%>
			<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("mandays"))%></td>		
			<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("rate"))%></td>		
			<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("totalamounthead"))%></td>			
			<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("volume"))%></td>
			<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("boronganpremi"))%></td>
		<% end if %>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("division"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=rscat("fieldcode")%></td>
		<%
		if txtDivision = rscat("division") AND txtFieldcode = rscat("fieldcode") AND txtTempJobcode = Cdbl(rscat("activitycode")) then
		%>
			<td align="right"><font face="Verdana" size="1"></td>
		<%
		else
		%>
			<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("haplan"))%></td>
		<% end if %>
		<td align="Left"><font face="Verdana" size="1"><%=rscat("material")%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("qty"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("ratestock"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("totalamount"))%></td>		
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("kg"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("jjg"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("rotasi"))%></td>
		<td align="right"><font face="Verdana" size="1"><%=FbpNum2(rscat("remarks"))%></td>		
		</tr>
	  <%
	  
	  txtTempJobcode = Cdbl(rscat("activitycode"))
	  txtFieldcode = rscat("fieldcode")
	  txtDivision = rscat("division")
	  rsCat.MoveNext
	loop
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