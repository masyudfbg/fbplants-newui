<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim txtDate, txtFcba, txtJob_forfieldstatus
	Dim strMain, nSpanTotal
	Dim rsAttDaily, strAttDaily, v_txtFcba, v_txtJob_forfieldstatus
	Dim v_txtDate, NO, cClass
	
	txtDate					=	request("txtDate")
	txtFcba 				=   request("txtFcba")
	txtJob_forfieldstatus	=	request("txtJob_forfieldstatus")
	
	v_txtDate				=	txtDate
	v_txtFcba 				=   txtFcba
	v_txtJob_forfieldstatus	=	txtJob_forfieldstatus

	nSpanTotal = 1
	


	
	strMain = " select par_1.fcba, param_1, job_forfieldstatus, kbt_0, kht_0, khl_0, total_0," & _
" kbt_1, kht_1, khl_1, total_1," & _
" kbt_2, kht_2, khl_2, total_2" & _
" from" & _
" (SELECT fcba, job_forfieldstatus, sum(kbt_0) kbt_0, sum(kht_0) kht_0, sum(khl_0) khl_0, sum(total_0) total_0," & _
" sum(kbt_1) kbt_1, sum(kht_1) kht_1, sum(khl_1) khl_1, sum(total_1) total_1," & _
" sum(kbt_2) kbt_2, sum(kht_2) kht_2, sum(khl_2) khl_2, sum(total_2) total_2" & _
" FROM (SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate," & _
" 'VRA' job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0," & _
" 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1," & _
" 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2" & _
" FROM job" & _
" UNION" & _
" SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate," & _
" 'WRA' job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0," & _
" 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1," & _
" 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2" & _
" FROM job" & _
" UNION" & _
" SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate," & _
" job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0," & _
" 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1," & _
" 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2" & _
" FROM job" & _
" UNION" & _
" SELECT fcba, fddate, job_forfieldstatus, kbt kbt_0, kht kht_0," & _
" khl khl_0, (kbt + kht + khl) total_0, 0 kbt_1, 0 kht_1," & _
" 0 khl_1, 0 total_1, 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2" & _
" FROM view_attendance_daily_summmary" & _
" WHERE fddate = TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
" UNION" & _
" SELECT fcba, fddate, job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0," & _
" 0 total_0, kbt kbt_1, kht kht_1, khl khl_1," & _
" (kbt + kht + khl) total_1, 0 kbt_2, 0 kht_2, 0 khl_2," & _
" 0 total_2" & _
" FROM view_attendance_daily_summmary" & _
" WHERE fddate >=" & _
" TO_DATE ( TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'MM')" & _
" || '/01/'" & _
" || TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')," & _
" 'YYYY'" & _
" )," & _
" 'MM/DD/YYYY'" & _
" )" & _
" AND fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
" UNION" & _
" SELECT fcba, fddate, job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0," & _
" 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1, kbt kbt_2," & _
" kht kht_2, khl khl_2, (kbt + kht + khl) total_2" & _
" FROM view_attendance_daily_summmary" & _
" WHERE fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'))" & _
" GROUP BY fcba, job_forfieldstatus)" & _
" par_1," & _
" (select fcba, fccode, param_1 from parameter" & _
" where" & _
" param_1 is not null and" & _
" parcode = 'J0001'" & _
" order by param_1) par_2" & _
" where" & _
" par_1.job_forfieldstatus = par_2.fccode and" & _
" par_1.fcba = par_2.fcba" & _
" order by par_2.param_1" 


	strAttDaily	=	" SELECT FCBA, param_1 "
	if len(trim(v_txtjob_forfieldstatus)) > 0 then 
		if v_txtjob_forfieldstatus <> "--SUMMARY--" then
			strAttDaily = strAttDaily & " , job_forfieldstatus "
		end if
	else
		strAttDaily = strAttDaily & " , job_forfieldstatus "
	end if
	strAttDaily	=	strAttDaily	& " , sum(kbt_0) kbt_0, sum(kht_0) kht_0, sum(khl_0) khl_0, sum(total_0) total_0, " & _
					" sum(kbt_1) kbt_1, sum(kht_1) kht_1, sum(khl_1) khl_1, sum(total_1) total_1, " & _
					" sum(kbt_2) kbt_2, sum(kht_2) kht_2, sum(khl_2) khl_2, sum(total_2) total_2 "
					
	strAttDaily = strAttDaily &  " FROM ( " & strMain & " )  " 
	
	    strAttDaily = strAttDaily & " WHERE fcba = '" & Session("fcba") & "' "
	
	if len(trim(v_txtjob_forfieldstatus)) > 0 then 
		if v_txtjob_forfieldstatus <> "--SUMMARY--" then
			strAttDaily = strAttDaily & " AND job_forfieldstatus = '" & v_txtjob_forfieldstatus & "'"
		end if
	end if
	
	strAttDaily = strAttDaily & " GROUP BY FCBA, param_1 "
	if len(trim(v_txtjob_forfieldstatus)) > 0 then 
		if v_txtjob_forfieldstatus <> "--SUMMARY--" then
			strAttDaily = strAttDaily & " , job_forfieldstatus "
		end if
	else
		strAttDaily = strAttDaily & " , job_forfieldstatus "
	end if
	
	strAttDaily = strAttDaily & " ORDER BY FCBA, param_1 "
	if len(trim(v_txtjob_forfieldstatus)) > 0 then 
		if v_txtjob_forfieldstatus <> "--SUMMARY--" then
			strAttDaily = strAttDaily & " , job_forfieldstatus "
		end if
	else
		strAttDaily = strAttDaily & " , job_forfieldstatus "
	end if


	Set rsAttDaily = Server.CreateObject("ADODB.Recordset")
	rsAttDaily.Open strAttDaily, connstr, adOpenStatic 
	
%>


<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Daily Attendance Report</title>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>

<body>

	<div id="wrapper">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/DashboardManager.asp">Back to Report</a></td>
		</tr>
	</table>
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN PEMAKAIAN TENAGA KERJA</h3>
		<h5>TANGGAL: <%=v_txtDate%></h5>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>
		<th>STATUS</th>
		<th>&nbsp;</th>
	</tr>
	<tr>		
	<form name="frmEntry" method="post" action="p_AttendanceDaily_BA_NAF_M.asp">
		<td align="center">
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
			<input type="text" maxlength="10" name="txtDate" size="10" value="<%=v_txtDate%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input type="text" maxlength="10" name="txtDate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txtDate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>

		<td align="center">
		<select name="txtJob_forfieldstatus">	
		<option value="">--ALL--</option>
		<%
			Dim rsJob_forfieldstatus, strsql
			Set rsJob_forfieldstatus = Server.CreateObject("ADODB.Recordset")
			rsJob_forfieldstatus.CursorLocation = adUseClient
			
		    if len(trim(v_txtFcba)) > 0 then 
				strsql =  " SELECT '--SUMMARY--' as Job_forfieldstatus from DUAL UNION SELECT DISTINCT Job_forfieldstatus as Job_forfieldstatus FROM JOB WHERE FCBA = '" & v_txtFcba & "'"
			else
				strsql =  " SELECT '--SUMMARY--' as Job_forfieldstatus from DUAL UNION SELECT DISTINCT Job_forfieldstatus as Job_forfieldstatus FROM JOB WHERE FCBA = '" & Session("fcBA") & "'"
			end if
			
			
			rsJob_forfieldstatus.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsJob_forfieldstatus.eof then
				DO while not rsJob_forfieldstatus.eof
				if v_txtJob_forfieldstatus = rsJob_forfieldstatus("Job_forfieldstatus") then
		%>			<option value="<%=rsJob_forfieldstatus("Job_forfieldstatus")%>" Selected><%=rsJob_forfieldstatus("Job_forfieldstatus")%></option>
				<%else%>
					<option value="<%=rsJob_forfieldstatus("Job_forfieldstatus")%>"><%=rsJob_forfieldstatus("Job_forfieldstatus")%></option>
				<%end if%>
		<%			rsJob_forfieldstatus.movenext
				loop
			End if
			rsJob_forfieldstatus.Close
			set rsJob_forfieldstatus = nothing
		%>
		</select>		
		</td>
		
		
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="2">NO</th>
		
		
		<% if trim(request("txtjob_forfieldstatus")) <> "--SUMMARY--" then %>
		    <th rowspan="2">URAIAN</th>
		<% else
			nSpanTotal = nSpanTotal - 1 			
	    end if %>
		
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th colspan="4">HARI INI <%=v_txtDate%></th>
		<% else %>
		<th colspan="4">HARI INI</th>
		<% end if%>
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th colspan="4">SAMPAI HARI INI <%=v_txtDate%></th>
		<% else %>
		<th colspan="4">SAMPAI HARI INI</th>
		<% end if%>
		<th colspan="4">SAMPAI DENGAN BULAN INI</th>
	</tr>
	<tr>
		<th>KBT</th>
		<th>KHT</th>
		<th>KHL</th>
		<th>TOTAL</th>
		<th>KBT</th>
		<th>KHT</th>
		<th>KHL</th>
		<th>TOTAL</th>
		<th>KBT</th>
		<th>KHT</th>
		<th>KHL</th>
		<th>TOTAL</th>
	</tr>

<%
Dim nkbt_0, nkht_0, nkhl_0, nTotal_0
Dim nkbt_1, nkht_1, nkhl_1, nTotal_1
Dim nkbt_2, nkht_2, nkhl_2, nTotal_2
Dim sprevFcba, sprevjob_forfieldstatus, tempjob_forfieldstatus
	if not rsAttDaily.EOF then
		NO = 1
		nkbt_0 = 0
		nkht_0 = 0
		nkhl_0 = 0
		nTotal_0 = 0

		nkbt_1 = 0
		nkht_1 = 0
		nkhl_1 = 0
		nTotal_1 = 0

		nkbt_2 = 0
		nkht_2 = 0
		nkhl_2 = 0
		nTotal_2 = 0
		
		sprevFcba = ""
		sprevjob_forfieldstatus	=	""
		
		while not rsAttDaily.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>		
				
				<% if trim(request("txtjob_forfieldstatus")) <> "--SUMMARY--" then %>
				    <% tempjob_forfieldstatus = rsAttDaily("job_forfieldstatus") %>
							
				    <% if sprevjob_forfieldstatus <> rsAttDaily("job_forfieldstatus") then %>
				        <td align="left">
						<%=rsAttDaily("job_forfieldstatus")%>
					</td>
						<% sprevjob_forfieldstatus = rsAttDaily("job_forfieldstatus") %>
					<% else	%>
					    <td align="left"></td>					
					<% end if %>							 
				<% end if %>
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_0"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_0"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_0"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_0"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_1"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_1"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_1"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_1"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_2"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_2"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_2"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_2"),2)%></td>
			</tr>
<%
			NO = NO + 1			
			nkbt_0 = nkbt_0 + Cdbl(rsAttDaily("kbt_0"))
			nkht_0 = nkht_0 + Cdbl(rsAttDaily("kht_0"))
			nkhl_0 = nkhl_0 + Cdbl(rsAttDaily("khl_0"))
			nTotal_0 = nTotal_0 + Cdbl(rsAttDaily("total_0"))

			nkbt_1 = nkbt_1 + Cdbl(rsAttDaily("kbt_1"))
			nkht_1 = nkht_1 + Cdbl(rsAttDaily("kht_1"))
			nkhl_1 = nkhl_1 + Cdbl(rsAttDaily("khl_1"))
			nTotal_1 = nTotal_1 + Cdbl(rsAttDaily("total_1"))

			nkbt_2 = nkbt_2 + Cdbl(rsAttDaily("kbt_2"))
			nkht_2 = nkht_2 + Cdbl(rsAttDaily("kht_2"))
			nkhl_2 = nkhl_2 + Cdbl(rsAttDaily("khl_2"))
			nTotal_2 = nTotal_2 + Cdbl(rsAttDaily("total_2"))
			
			
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="<%= 1 + nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(nkbt_0,2)%></td>
				<td align="right"><%=FormatNumber(nkht_0,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_0,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_0,2)%></td>

				<td align="right"><%=FormatNumber(nkbt_1,2)%></td>
				<td align="right"><%=FormatNumber(nkht_1,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_1,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>

				<td align="right"><%=FormatNumber(nkbt_2,2)%></td>
				<td align="right"><%=FormatNumber(nkht_2,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_2,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>
				
				
			</tr>
</table>
</form>
<!--#include file="../Bottom_Report.asp"-->
