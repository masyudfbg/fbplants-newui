<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim rsAttDaily, strAttDaily, strGroup, strOrder, strQuery, v_txtFcba
	Dim v_txtDate, NO, cClass
	
	v_txtDate		=	trim(request.Form("txtDate"))
	v_txtFcba 		=   trim(request.Form("txtFcba"))	
	
	
	strAttDaily = " SELECT fcba,  job_forfieldstatus, sum(kbt_0) kbt_0, sum(kht_0) kht_0, sum(khl_0) khl_0, sum(total_0) total_0, " & _
													  " sum(kbt_1) kbt_1, sum(kht_1) kht_1, sum(khl_1) khl_1, sum(total_1) total_1, " & _
													  " sum(kbt_2) kbt_2, sum(kht_2) kht_2, sum(khl_2) khl_2, sum(total_2) total_2 " & _
					  " FROM (SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate, " & _
											" 'VRA' job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0, " & _
											" 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1, " & _
											" 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2 " & _
									   " FROM job " & _
							" UNION " & _
							" SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate, " & _
											" 'WRA' job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0, " & _
											" 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1, " & _
											" 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2 " & _
									   " FROM job " & _
							" UNION " & _
							" SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate, " & _
											" job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0, " & _
											" 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1, " & _
											" 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2 " & _
									   " FROM job " & _
							" UNION " & _
							" SELECT fcba, fddate, job_forfieldstatus, kbt kbt_0, kht kht_0, " & _
								   " khl khl_0, (kbt + kht + khl) total_0, 0 kbt_1, 0 kht_1, " & _
								   " 0 khl_1, 0 total_1, 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2 " & _
							  " FROM view_attendance_daily_summmary " & _
							 " WHERE fcba = '" & v_txtFcba & "' AND fddate = TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') " & _
							" UNION " & _
							" SELECT fcba, fddate, job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0, " & _
								   " 0 total_0, kbt kbt_1, kht kht_1, khl khl_1, " & _
								   " (kbt + kht + khl) total_1, 0 kbt_2, 0 kht_2, 0 khl_2, " & _
								   " 0 total_2 " & _
							  " FROM view_attendance_daily_summmary " & _
							 " WHERE fcba = '" & v_txtFcba & "' " &_ 
							     " AND fddate >= " & _
									  " TO_DATE (   TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'MM') " & _
											   " || '/01/' " & _
											   " || TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'), " & _
														   " 'YYYY' " & _
														  " ), " & _
											   " 'MM/DD/YYYY' " & _
											  " ) " & _
							     " AND fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') " & _
							" UNION " & _
							" SELECT fcba, fddate, job_forfieldstatus, 0 kbt_0, 0 kht_0, 0 khl_0, " & _
								   " 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1, kbt kbt_2, " & _
								   " kht kht_2, khl khl_2, (kbt + kht + khl) total_2 " & _
							  " FROM view_attendance_daily_summmary " & _
							 " WHERE fcba = '" & v_txtFcba & "' " & _
							    " AND fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')) " & _
							" WHERE fcba =  '" & v_txtFcba & "' " 
	
	
	
	' strAttDaily	=	" SELECT fcba, job_forfieldstatus, SUM (kbt) kbt, SUM (kht) kht, SUM (khl) khl" & _
					' " FROM (SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate," & _
					' " 'VRA' job_forfieldstatus, 'VH' job_category, 0 kbt," & _
					' " 0 kht, 0 khl" & _
					' " FROM job" & _
					' " UNION" & _
					' " SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate," & _
					' " 'WRA' job_forfieldstatus, 'WS' job_category, 0 kbt," & _
					' " 0 kht, 0 khl" & _
					' " FROM job" & _
					' " UNION" & _
					' " SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate," & _
					' " job_forfieldstatus, job_category, 0 kbt, 0 kht," & _
					' " 0 khl" & _
					' " FROM job" & _
					' " UNION" & _
					' " SELECT fcba, fddate, job_forfieldstatus, job_category," & _
					' " SUM (mandays) kbt, 0 kht, 0 khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KBT' AND job_forfieldstatus <> 'BTL'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, job_forfieldstatus, job_category, 0 kbt," & _
					' " SUM (mandays) kht, 0 khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KHT' AND job_forfieldstatus <> 'BTL'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, job_forfieldstatus, job_category, 0 kbt," & _
					' " 0 kht, SUM (mandays) khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KHL' AND job_forfieldstatus <> 'BTL'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, 'WRA' job_forfieldstatus, job_category," & _
					' " SUM (mandays) kbt, 0 kht, 0 khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KBT' AND jobcode = '6210100'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, 'WRA' job_forfieldstatus, job_category," & _
					' " 0 kbt, SUM (mandays) kht, 0 khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KHT' AND jobcode = '6210100'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, 'WRA' job_forfieldstatus, job_category," & _
					' " 0 kbt, 0 kht, SUM (mandays) khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KHL' AND jobcode = '6210100'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, 'VRA' job_forfieldstatus, job_category," & _
					' " SUM (mandays) kbt, 0 kht, 0 khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KBT' AND jobcode = '6310100'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, 'VRA' job_forfieldstatus, job_category," & _
					' " 0 kbt, SUM (mandays) kht, 0 khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KHT' AND jobcode = '6310100'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, 'VRA' job_forfieldstatus, job_category," & _
					' " 0 kbt, 0 kht, SUM (mandays) khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KHL' AND jobcode = '6310100'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, job_forfieldstatus, job_category," & _
					' " SUM (mandays) kbt, 0 kht, 0 khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KBT'" & _
					' " AND jobcode NOT IN ('6210100', '6310100')" & _
					' " AND job_forfieldstatus = 'BTL'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, job_forfieldstatus, job_category, 0 kbt," & _
					' " SUM (mandays) kht, 0 khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KHT'" & _
					' " AND jobcode NOT IN ('6210100', '6310100')" & _
					' " AND job_forfieldstatus = 'BTL'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category" & _
					' " UNION" & _
					' " SELECT fcba, fddate, job_forfieldstatus, job_category, 0 kbt," & _
					' " 0 kht, SUM (mandays) khl" & _
					' " FROM view_attendance_daily" & _
					' " WHERE presentgrade = 'KHL'" & _
					' " AND jobcode NOT IN ('6210100', '6310100')" & _
					' " AND job_forfieldstatus = 'BTL'" & _
					' " GROUP BY fcba, fddate, job_forfieldstatus, job_category)" & _
					' " WHERE fcba = '" & Session("fcba") & "' AND fddate = TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')"
						
	strGroup	=	" GROUP BY fcba, job_forfieldstatus"
	strOrder	=	" ORDER BY fcba, job_forfieldstatus"
	
	strQuery	=	strAttDaily + strGroup + strOrder

	Set rsAttDaily = Server.CreateObject("ADODB.Recordset")
	rsAttDaily.Open strQuery, connstr, adOpenStatic 
	
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
	<!--#include file="menutop.asp"-->
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN PEMAKAIAN TENAGA KERJA PER KEBUN</h3>
		<h5>TANGGAL: <%=v_txtDate%>, KEBUN : <%=v_txtFcba %></h5>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>
	    <th>KEBUN</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_AttendanceDaily_BA.asp">
		
		
		
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
			<select name="txtFcba">	
			<option value="">--ALL--</option>		
			<%
				Dim rsFcba, strsql1
				Set rsFcba = Server.CreateObject("ADODB.Recordset")
				rsFcba.CursorLocation = adUseClient
				strsql1 =  " SELECT fccode FROM businessunit "
				rsFcba.Open strsql1, connStr, adOpenForwardOnly, adLockReadOnly	
				if not rsFcba.eof then
					DO while not rsFcba.eof
					if v_txtFcba = rsFcba("fccode") then
			%>			<option value="<%=rsFcba("fccode")%>" Selected><%=rsFcba("fccode")%></option>
					<%else%>
						<option value="<%=rsFcba("fccode")%>"><%=rsFcba("fccode")%></option>
					<%end if%>
			<%			rsFcba.movenext
					loop
				End if
				rsFcba.Close
				set rsFcba = nothing
			%>
			</select>				
		</td>
		
		
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="2">NO</th>
		<th rowspan="2">URAIAN</th>
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
		
		
		while not rsAttDaily.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="left"><%=rsAttDaily(1)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(2),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(3),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(4),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(5),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily(6),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(7),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(8),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(9),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily(10),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(11),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(12),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily(13),2)%></td>
			</tr>
<%
			NO = NO + 1			
			nkbt_0 = nkbt_0 + Cdbl(rsAttDaily(2))
			nkht_0 = nkht_0 + Cdbl(rsAttDaily(3))
			nkhl_0 = nkhl_0 + Cdbl(rsAttDaily(4))
			nTotal_0 = nTotal_0 + Cdbl(rsAttDaily(5))

			nkbt_1 = nkbt_1 + Cdbl(rsAttDaily(6))
			nkht_1 = nkht_1 + Cdbl(rsAttDaily(7))
			nkhl_1 = nkhl_1 + Cdbl(rsAttDaily(8))
			nTotal_1 = nTotal_1 + Cdbl(rsAttDaily(9))

			nkbt_2 = nkbt_2 + Cdbl(rsAttDaily(10))
			nkht_2 = nkht_2 + Cdbl(rsAttDaily(11))
			nkhl_2 = nkhl_2 + Cdbl(rsAttDaily(12))
			nTotal_2 = nTotal_2 + Cdbl(rsAttDaily(13))
			
			
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="2">TOTAL</td>
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

<!--#include file="../Bottom_Report.asp"-->
